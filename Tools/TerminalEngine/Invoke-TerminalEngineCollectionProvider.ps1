[CmdletBinding()]
param(
    [string]$Phase = '',

    [string[]]$CandidateId = @(),

    [string]$Platform = '',

    [string]$Configuration = '',

    [string]$RepoRoot = '',

    [switch]$LoadOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RSProviderCandidateSetLeaf =
    'Specs\Terminal\TerminalEngineCandidateSet.v1.json'
$script:RSProviderCandidateReviewLeaf =
    'Specs\Terminal\TerminalEngineCandidateReview.v1.json'
$script:RSProviderCandidateAssessmentLeaf =
    'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json'
$script:RSProviderCandidateUniverseLeaf =
    'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json'
$script:RSProviderDescriptorSchemaLeaf =
    'Specs\Terminal\TerminalEngineCollectionDescriptor.schema.json'
$script:RSProviderDescriptorLeaf =
    'red-salamander\collection-descriptor.v1.json'
$script:RSProviderWrapperInputSchemaLeaf =
    'Tools\TerminalEngine\TerminalEngineCollectionDriverInput.schema.json'
$script:RSProviderWrapperResultSchemaLeaf =
    'Tools\TerminalEngine\TerminalEngineCollectionDriverResult.schema.json'
$script:RSProviderBootstrapLeaf =
    '.build\terminal-engine-spike\bootstrap'
$script:RSProviderLaneLeaf =
    '.build\terminal-engine-spike\lanes'
$script:RSProviderWrapperInputLeaf =
    'provider-input\candidate-wrapper-input.jcs.json'
$script:RSProviderWrapperResultLeaf =
    'candidate-wrapper-result.jcs.json'
$script:RSProviderBuildReceiptLeaf =
    'provider-input\build-identity.jcs.json'
$script:RSProviderSourceDependenciesLeaf =
    'provider-input\adapter-source-dependencies.json'
$script:RSProviderAdapterHeaderLeaf =
    'Tools\TerminalEngine\TerminalEngineGate0Adapter.h'
$script:RSProviderLayoutHeaderLeaf =
    'Tools\TerminalEngine\TerminalEngineGate0Layout.h'
$script:RSProviderHostHeadersInputId = 'host-wil-headers'
$script:RSProviderCapabilities = [uint64]127
$script:RSProviderDescriptorMaxBytes = [uint64]65536
$script:RSProviderWrapperInputMaxBytes = [uint64]32768
$script:RSProviderWrapperResultMaxBytes = [uint64]4096
$script:RSProviderLoadedRepoRoot = ''
$script:RSProviderEvaluatorModule = $null

function Get-RSProviderSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath
    )

    if (-not [IO.File]::Exists($LiteralPath)) {
        throw [IO.FileNotFoundException]::new(
            "Required file does not exist: $LiteralPath",
            $LiteralPath)
    }
    return (Get-FileHash -LiteralPath $LiteralPath -Algorithm SHA256).
        Hash.ToLowerInvariant()
}

function Get-RSProviderBytesSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToHexString(
            $sha256.ComputeHash($Bytes)).ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-RSProviderCanonicalRepoRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or
        -not [IO.Path]::IsPathFullyQualified($Path)) {
        throw [IO.InvalidDataException]::new(
            'RepoRoot must be an existing fully qualified directory.')
    }
    $resolved = @(Resolve-Path -LiteralPath $Path -ErrorAction Stop)
    if ($resolved.Count -ne 1 -or
        $resolved[0].Provider.Name -cne 'FileSystem' -or
        -not [IO.Directory]::Exists($resolved[0].ProviderPath)) {
        throw [IO.InvalidDataException]::new(
            'RepoRoot must resolve to exactly one filesystem directory.')
    }
    return [IO.Path]::GetFullPath($resolved[0].ProviderPath).
        TrimEnd([IO.Path]::DirectorySeparatorChar)
}

function Assert-RSProviderSafeSegment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if ($Value -cnotmatch '^[a-z0-9][a-z0-9._-]{0,63}$') {
        throw [IO.InvalidDataException]::new(
            "$Label is not a safe lowercase path segment: '$Value'.")
    }
}

function Assert-RSProviderSafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $baseName = $Value.Split('.')[0]
    if ($Value -cnotmatch '^[A-Za-z0-9][A-Za-z0-9._+-]{0,127}$' -or
        [IO.Path]::GetFileName($Value) -cne $Value -or
        $Value.EndsWith('.', [StringComparison]::Ordinal) -or
        $Value.EndsWith(' ', [StringComparison]::Ordinal) -or
        $baseName -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') {
        throw [IO.InvalidDataException]::new(
            "$Label is not one safe filesystem leaf: '$Value'.")
    }
}

function Get-RSProviderContainedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if ([IO.Path]::IsPathFullyQualified($RelativePath)) {
        throw [IO.InvalidDataException]::new(
            "$Label must be relative to its provider-owned root.")
    }
    $canonicalRoot = [IO.Path]::GetFullPath($Root).
        TrimEnd([IO.Path]::DirectorySeparatorChar)
    $candidate = [IO.Path]::GetFullPath((
            Join-Path $canonicalRoot $RelativePath))
    if (-not $candidate.StartsWith(
            $canonicalRoot + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw [IO.InvalidDataException]::new(
            "$Label escaped its provider-owned root.")
    }
    return $candidate
}

function Assert-RSProviderNoReparsePoint {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $root = [IO.Path]::GetFullPath($AllowedRoot).
        TrimEnd([IO.Path]::DirectorySeparatorChar)
    $current = [IO.Path]::GetFullPath($LiteralPath)
    if (-not $current.Equals(
            $root,
            [StringComparison]::OrdinalIgnoreCase) -and
        -not $current.StartsWith(
            $root + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw [IO.InvalidDataException]::new(
            "$Label escaped its allowed root.")
    }
    while (-not $current.Equals(
            $root,
            [StringComparison]::OrdinalIgnoreCase)) {
        if ([IO.File]::Exists($current) -or
            [IO.Directory]::Exists($current)) {
            if (((Get-Item -LiteralPath $current -Force).Attributes -band
                    [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw [IO.InvalidDataException]::new(
                    "$Label contains a reparse point.")
            }
        }
        $parent = [IO.Path]::GetDirectoryName($current)
        if ([string]::IsNullOrEmpty($parent) -or $parent -ceq $current) {
            throw [IO.InvalidDataException]::new(
                "$Label could not be proven inside its allowed root.")
        }
        $current = $parent
    }
    if ([IO.Directory]::Exists($root) -and
        ((Get-Item -LiteralPath $root -Force).Attributes -band
            [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw [IO.InvalidDataException]::new(
            "$Label has a reparse-point root.")
    }
}

function Get-RSProviderOrdinalStrings {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Value
    )

    $result = [string[]]@($Value)
    [Array]::Sort($result, [StringComparer]::Ordinal)
    return $result
}

function Assert-RSProviderExactProperties {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [string[]]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if ($Value -isnot [Collections.IDictionary]) {
        throw [IO.InvalidDataException]::new("$Label must be a JSON object.")
    }
    $actual = Get-RSProviderOrdinalStrings -Value (
        [string[]]@($Value.Keys | ForEach-Object { [string]$_ }))
    $orderedExpected = Get-RSProviderOrdinalStrings -Value $Expected
    if ($actual.Count -ne $orderedExpected.Count) {
        throw [IO.InvalidDataException]::new(
            "$Label does not have its exact closed property set.")
    }
    for ($index = 0; $index -lt $actual.Count; ++$index) {
        if ($actual[$index] -cne $orderedExpected[$index]) {
            throw [IO.InvalidDataException]::new(
                "$Label does not have its exact closed property set.")
        }
    }
}

function Read-RSProviderCanonicalJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [string]$ExpectedSha256 = '',

        [uint64]$MaximumBytes = [uint64]16777216,

        [switch]$PreliminaryUnauthenticated
    )

    $resolved = [IO.Path]::GetFullPath($LiteralPath)
    if (-not [IO.File]::Exists($resolved)) {
        throw [IO.FileNotFoundException]::new(
            "$Label does not exist.",
            $resolved)
    }
    Assert-RSProviderNoReparsePoint `
        -LiteralPath $resolved `
        -AllowedRoot $RepoRoot `
        -Label $Label
    $stream = [IO.FileStream]::new(
        $resolved,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::Read)
    try {
        $length = $stream.Length
        if ($length -le 0 -or [uint64]$length -gt $MaximumBytes -or
            $length -gt [int]::MaxValue) {
            throw [IO.InvalidDataException]::new(
                "$Label violates its fixed byte bound.")
        }
        $bytes = [byte[]]::new([int]$length)
        $offset = 0
        while ($offset -lt $bytes.Length) {
            $read = $stream.Read(
                $bytes,
                $offset,
                $bytes.Length - $offset)
            if ($read -le 0) {
                throw [IO.EndOfStreamException]::new(
                    "$Label ended before its locked byte length.")
            }
            $offset += $read
        }
    }
    finally {
        $stream.Dispose()
    }
    $actualSha256 = Get-RSProviderBytesSha256 -Bytes $bytes
    if (-not [string]::IsNullOrEmpty($ExpectedSha256)) {
        if ($ExpectedSha256 -cnotmatch '^[0-9a-f]{64}$' -or
            $actualSha256 -cne $ExpectedSha256) {
            throw [IO.InvalidDataException]::new(
                "$Label differs from its authenticated frozen identity.")
        }
    }
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $parsed = $json |
            ConvertFrom-Json `
                -AsHashtable `
                -Depth 100 `
                -NoEnumerate `
                -DateKind String
    }
    catch {
        throw [IO.InvalidDataException]::new(
            "$Label is not strict UTF-8 JSON.",
            $_.Exception)
    }
    # This mode is used only to reject an obviously non-executable cohort
    # before any repository module can execute. An executable cohort must be
    # replaced by the evaluator-authenticated closure before provider effects.
    if ($PreliminaryUnauthenticated) {
        return $parsed
    }
    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $canonical = ConvertTo-RSJcsUtf8Bytes -InputObject $parsed
    if (-not [Linq.Enumerable]::SequenceEqual(
            [byte[]]$bytes,
            [byte[]]$canonical)) {
        throw [IO.InvalidDataException]::new(
            "$Label is not exact RFC 8785 JCS.")
    }
    return $parsed
}

function Get-RSProviderEvaluatorModule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $evaluatorPath = [IO.Path]::GetFullPath(
        (Join-Path $RepoRoot 'Tools\TerminalEngineEvaluator.psm1'))
    if ($null -eq $script:RSProviderEvaluatorModule) {
        $imported = @(
            Import-Module `
                -Name $evaluatorPath `
                -Force `
                -PassThru)
        $exact = @(
            $imported | Where-Object {
                -not [string]::IsNullOrEmpty([string]$_.Path) -and
                [IO.Path]::GetFullPath([string]$_.Path).Equals(
                    $evaluatorPath,
                    [StringComparison]::OrdinalIgnoreCase)
            })
        if ($exact.Count -ne 1) {
            throw [IO.InvalidDataException]::new(
                'Provider could not bind the exact trusted terminal-engine evaluator module.')
        }
        $script:RSProviderEvaluatorModule = $exact[0]
    }
    $boundPath = [string]$script:RSProviderEvaluatorModule.Path
    if ([string]::IsNullOrEmpty($boundPath) -or
        -not [IO.Path]::GetFullPath($boundPath).Equals(
            $evaluatorPath,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw [IO.InvalidDataException]::new(
            'Provider evaluator binding changed or crossed repository roots.')
    }
    return ,$script:RSProviderEvaluatorModule
}

function Read-RSProviderCandidateGovernanceClosure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $evaluator = Get-RSProviderEvaluatorModule -RepoRoot $RepoRoot
    return & $evaluator {
        param($root)
        Read-RSCandidateGovernanceClosure -RepoRoot $root
    } $RepoRoot
}

function Write-RSProviderCanonicalJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [uint64]$MaximumBytes,

        [string]$SchemaFile = ''
    )

    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $bytes = ConvertTo-RSJcsUtf8Bytes -InputObject $Value
    if ($bytes.Length -le 0 -or [uint64]$bytes.Length -gt $MaximumBytes) {
        throw [IO.InvalidDataException]::new(
            'Provider JCS output violates its fixed byte bound.')
    }
    if (-not [string]::IsNullOrEmpty($SchemaFile)) {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        if (-not (Test-Json `
                -Json $json `
                -SchemaFile $SchemaFile `
                -ErrorAction Stop)) {
            throw [IO.InvalidDataException]::new(
                'Provider JCS output violates its closed schema.')
        }
    }
    $directory = [IO.Path]::GetDirectoryName(
        [IO.Path]::GetFullPath($LiteralPath))
    [void][IO.Directory]::CreateDirectory($directory)
    $stream = [IO.FileStream]::new(
        [IO.Path]::GetFullPath($LiteralPath),
        [IO.FileMode]::CreateNew,
        [IO.FileAccess]::Write,
        [IO.FileShare]::None)
    try {
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    }
    finally {
        $stream.Dispose()
    }
    return [pscustomobject]@{
        Path = [IO.Path]::GetFullPath($LiteralPath)
        Sha256 = Get-RSProviderSha256 -LiteralPath $LiteralPath
        SizeBytes = ([uint64]$bytes.Length).ToString(
            [Globalization.CultureInfo]::InvariantCulture)
    }
}

function Get-RSProviderDirectoryIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $root = [IO.Path]::GetFullPath($LiteralPath)
    if (-not [IO.Directory]::Exists($root)) {
        throw [IO.DirectoryNotFoundException]::new(
            "Required directory does not exist: $root")
    }
    Assert-RSProviderNoReparsePoint `
        -LiteralPath $root `
        -AllowedRoot $AllowedRoot `
        -Label 'provider directory identity root'
    Import-Module (Join-Path $RepoRoot (
            'Tools\TerminalEngine\TerminalEngineLifecycle.psm1'))
    return [string](Get-RSDirectoryIdentity -Path $root).sha256
}

function Get-RSProviderConfigurationLeaf {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Configuration
    )

    switch ($Configuration) {
        'Release' { return 'Release' }
        'ASan Debug' { return 'ASan-Debug' }
        default {
            throw [IO.InvalidDataException]::new(
                "Unsupported Gate-0 configuration '$Configuration'.")
        }
    }
}

function Get-RSProviderContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$CandidateId,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$Configuration
    )

    if ($Platform -cne 'x64' -and $Platform -cne 'ARM64') {
        throw [IO.InvalidDataException]::new(
            'Platform must be exactly x64 or ARM64.')
    }
    if ("$Platform|$Configuration" -cnotin @(
            'x64|Release',
            'x64|ASan Debug',
            'ARM64|Release')) {
        throw [IO.InvalidDataException]::new(
            "Unsupported Gate-0 collection lane '$Platform|$Configuration'.")
    }
    $candidateSetPath = Join-Path `
        $RepoRoot `
        $script:RSProviderCandidateSetLeaf
    $candidateSetSha256 =
        Get-RSProviderSha256 -LiteralPath $candidateSetPath
    $candidateSet = Read-RSProviderCanonicalJson `
        -LiteralPath $candidateSetPath `
        -RepoRoot $RepoRoot `
        -Label 'Frozen candidate set' `
        -PreliminaryUnauthenticated
    if ([string]$candidateSet['formatId'] -cne
            'red-salamander-terminal-engine-candidate-set' -or
        [int]$candidateSet['version'] -ne 1) {
        throw [IO.InvalidDataException]::new(
            'Frozen candidate-set identity is invalid.')
    }

    $candidateById =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::Ordinal)
    foreach ($candidate in [object[]]@($candidateSet['candidates'])) {
        $id = [string]$candidate['candidateId']
        Assert-RSProviderSafeSegment -Value $id -Label 'candidateId'
        if (-not $candidateById.TryAdd($id, $candidate)) {
            throw [IO.InvalidDataException]::new(
                "Frozen candidate set repeats '$id'.")
        }
    }
    $requested =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
    foreach ($id in $CandidateId) {
        Assert-RSProviderSafeSegment `
            -Value $id `
            -Label 'selected candidateId'
        if (-not $requested.Add($id)) {
            throw [IO.InvalidDataException]::new(
                "Selected candidate '$id' is repeated.")
        }
        if (-not $candidateById.ContainsKey($id)) {
            throw [IO.InvalidDataException]::new(
                "Selected candidate '$id' is not frozen.")
        }
    }
    if ($requested.Count -eq 0) {
        throw [IO.InvalidDataException]::new(
            'At least one frozen candidate must be selected.')
    }
    $selectedIds = Get-RSProviderOrdinalStrings -Value (
        [string[]]$requested)
    $allCandidateIds = Get-RSProviderOrdinalStrings -Value (
        [string[]]@($candidateById.Keys))
    if (($selectedIds -join "`0") -cne
        ($allCandidateIds -join "`0")) {
        throw [IO.InvalidDataException]::new(
            'Provider phases require the exact frozen candidate set; partial collection manifests are forbidden.')
    }
    $selectedCandidates = [object[]]@(
        $selectedIds | ForEach-Object { $candidateById[$_] })
    $executableCandidates = @(
        $selectedCandidates | Where-Object {
            [string]$_['scoredOrControl'] -ceq 'scored' -and
            [string]$_['disposition'] -ceq 'collect'
        })
    if ($executableCandidates.Count -eq 0) {
        throw [IO.InvalidDataException]::new(
            'Provider phases require an executable collect candidate; a static-only cohort cannot attest a lane.')
    }

    if ((Get-RSProviderSha256 -LiteralPath $candidateSetPath) -cne
        $candidateSetSha256) {
        throw [IO.InvalidDataException]::new(
            'Candidate set changed while the preliminary cohort was selected.')
    }
    $universePath = Join-Path `
        $RepoRoot `
        $script:RSProviderCandidateUniverseLeaf
    $candidateUniverseSha256 =
        Get-RSProviderSha256 -LiteralPath $universePath
    if ($candidateUniverseSha256 -cne
            [string]$candidateSet['candidateUniverseSha256']) {
        throw [IO.InvalidDataException]::new(
            'Candidate set is not bound to the exact round-4 universe.')
    }

    $governance = Read-RSProviderCandidateGovernanceClosure `
        -RepoRoot $RepoRoot
    if ($null -eq $governance -or
        [string]$governance.CandidateSetSha256 -cne $candidateSetSha256 -or
        [string]$governance.UniverseSha256 -cne $candidateUniverseSha256) {
        throw [IO.InvalidDataException]::new(
            'Provider governance closure does not match the preliminarily selected records.')
    }
    $candidateSet = $governance.CandidateSet
    $universe = $governance.Universe
    if ($null -eq $candidateSet -or $null -eq $universe) {
        throw [IO.InvalidDataException]::new(
            'Provider governance closure omitted a required manifest.')
    }
    $neutralityPolicySha256 = [string]$governance.NeutralityPolicySha256
    if ($neutralityPolicySha256 -cnotmatch '^[0-9a-f]{64}$') {
        throw [IO.InvalidDataException]::new(
            'Provider governance closure omitted the authenticated neutrality-policy identity.')
    }
    $candidateById =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::Ordinal)
    foreach ($candidate in [object[]]@($candidateSet['candidates'])) {
        $id = [string]$candidate['candidateId']
        Assert-RSProviderSafeSegment -Value $id -Label 'candidateId'
        if (-not $candidateById.TryAdd($id, $candidate)) {
            throw [IO.InvalidDataException]::new(
                "Governance closure repeats candidate '$id'.")
        }
    }
    $closedCandidateIds = Get-RSProviderOrdinalStrings -Value (
        [string[]]@($candidateById.Keys))
    if (($closedCandidateIds -join "`0") -cne
        ($selectedIds -join "`0")) {
        throw [IO.InvalidDataException]::new(
            'Governance closure changed the exact selected candidate cohort.')
    }
    $selectedCandidates = [object[]]@(
        $selectedIds | ForEach-Object { $candidateById[$_] })
    $closedExecutableCandidates = @(
        $selectedCandidates | Where-Object {
            [string]$_['scoredOrControl'] -ceq 'scored' -and
            [string]$_['disposition'] -ceq 'collect'
        })
    if ($closedExecutableCandidates.Count -eq 0) {
        throw [IO.InvalidDataException]::new(
            'Governance closure returned a static-only candidate cohort.')
    }

    $inputSets = @(
        [object[]]@($candidateSet['bootstrapInputSets']) |
            Where-Object { [string]$_['platform'] -ceq $Platform })
    if ($inputSets.Count -ne 1) {
        throw [IO.InvalidDataException]::new(
            "Frozen candidate set must contain exactly one $Platform input set.")
    }
    $inputSet = $inputSets[0]
    $inputs = [object[]]@($inputSet['inputs'])
    $bootstrapInputSetSha256 = [string]$inputSet['sha256']
    if ($bootstrapInputSetSha256 -cnotmatch '^[0-9a-f]{64}$') {
        throw [IO.InvalidDataException]::new(
            "Authenticated $Platform bootstrap input-set digest is invalid.")
    }
    $inputById =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::Ordinal)
    foreach ($input in $inputs) {
        $id = [string]$input['inputId']
        Assert-RSProviderSafeSegment -Value $id -Label 'bootstrap inputId'
        Assert-RSProviderSafeFileName `
            -Value ([string]$input['fileName']) `
            -Label "bootstrap input '$id' fileName"
        if (-not $inputById.TryAdd($id, $input)) {
            throw [IO.InvalidDataException]::new(
                "Frozen bootstrap input '$id' is repeated.")
        }
        if ([string]$input['sha256'] -cnotmatch '^[0-9a-f]{64}$') {
            throw [IO.InvalidDataException]::new(
                "Frozen bootstrap input '$id' has an invalid SHA-256.")
        }
        $uri = $null
        if (-not [Uri]::TryCreate(
                [string]$input['retrievalUrl'],
                [UriKind]::Absolute,
                [ref]$uri) -or
            $uri.Scheme -cne 'https' -or
            -not [string]::IsNullOrEmpty($uri.UserInfo)) {
            throw [IO.InvalidDataException]::new(
                "Frozen bootstrap input '$id' has an unsafe retrieval URL.")
        }
    }
    return [pscustomobject]@{
        RepoRoot = $RepoRoot
        CandidateSet = $candidateSet
        CandidateSetPath = $candidateSetPath
        CandidateSetSha256 = $candidateSetSha256
        CandidateUniverse = $universe
        CandidateUniversePath = $universePath
        CandidateUniverseSha256 = $candidateUniverseSha256
        CandidateReviewSha256 =
            [string]$governance.CandidateReviewSha256
        CandidateAssessmentSha256 =
            [string]$governance.CandidateAssessmentSha256
        GovernanceClosure = $governance
        NeutralityPolicySha256 = $neutralityPolicySha256
        CandidateById = $candidateById
        SelectedIds = $selectedIds
        SelectedCandidates = $selectedCandidates
        InputSet = $inputSet
        Inputs = $inputs
        InputById = $inputById
        BootstrapInputSetSha256 = $bootstrapInputSetSha256
        Platform = $Platform
        Configuration = $Configuration
    }
}

function Test-RSProviderInputSelected {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputRecord,

        [Parameter(Mandatory = $true)]
        [string[]]$SelectedIds
    )

    foreach ($consumer in [object[]]@($InputRecord['consumers'])) {
        if ($SelectedIds -ccontains [string]$consumer) {
            return $true
        }
    }
    return $false
}

function Get-RSProviderInputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BootstrapRoot,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputRecord
    )

    $inputId = [string]$InputRecord['inputId']
    $fileName = [string]$InputRecord['fileName']
    Assert-RSProviderSafeSegment -Value $inputId -Label 'bootstrap inputId'
    Assert-RSProviderSafeFileName `
        -Value $fileName `
        -Label "bootstrap input '$inputId' fileName"
    $platformRoot = Get-RSProviderContainedPath `
        -Root $BootstrapRoot `
        -RelativePath (Join-Path 'inputs' $Platform) `
        -Label 'platform input root'
    $inputRoot = Get-RSProviderContainedPath `
        -Root $platformRoot `
        -RelativePath $inputId `
        -Label "bootstrap input '$inputId' root"
    return Get-RSProviderContainedPath `
        -Root $inputRoot `
        -RelativePath $fileName `
        -Label "bootstrap input '$inputId'"
}

function Get-RSProviderVerifiedInput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BootstrapRoot,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputRecord
    )

    $target = Get-RSProviderInputPath `
        -BootstrapRoot $BootstrapRoot `
        -Platform $Platform `
        -InputRecord $InputRecord
    $directory = [IO.Path]::GetDirectoryName($target)
    [void][IO.Directory]::CreateDirectory($directory)
    Assert-RSProviderNoReparsePoint `
        -LiteralPath $target `
        -AllowedRoot $BootstrapRoot `
        -Label "bootstrap input '$($InputRecord['inputId'])'"
    $expected = [string]$InputRecord['sha256']
    if ([IO.File]::Exists($target)) {
        if ((Get-RSProviderSha256 -LiteralPath $target) -cne $expected) {
            throw [IO.InvalidDataException]::new(
                "Cached bootstrap input '$($InputRecord['inputId'])' digest mismatch.")
        }
        return $target
    }
    $temporary = Join-Path $directory (
        ".download-$([Guid]::NewGuid().ToString('N')).tmp")
    try {
        $null = Invoke-WebRequest `
            -Uri ([string]$InputRecord['retrievalUrl']) `
            -OutFile $temporary `
            -MaximumRedirection 5 `
            -UseBasicParsing
        if ((Get-RSProviderSha256 -LiteralPath $temporary) -cne $expected) {
            throw [IO.InvalidDataException]::new(
                "Downloaded bootstrap input '$($InputRecord['inputId'])' digest mismatch.")
        }
        [IO.File]::Move($temporary, $target, $false)
    }
    finally {
        if ([IO.File]::Exists($temporary)) {
            [IO.File]::Delete($temporary)
        }
    }
    return $target
}

function Get-RSProviderCandidateSourceInput {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate
    )

    $candidateId = [string]$Candidate['candidateId']
    $matches = @(
        $Context.Inputs | Where-Object {
            [string]$_['inputKind'] -ceq 'source-archive' -and
            [object[]]@($_['consumers']) -ccontains $candidateId -and
            [string]$_['sha256'] -ceq
                [string]$Candidate['sourceArchiveSha256']
        })
    if ($matches.Count -ne 1) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' must resolve to exactly one frozen source archive.")
    }
    return $matches[0]
}

function Get-RSProviderAssessmentCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate
    )

    $assessmentResult = $Context.GovernanceClosure.AssessmentResult
    if ($null -eq $assessmentResult -or
        $null -eq $assessmentResult.CandidateById) {
        throw [IO.InvalidDataException]::new(
            'Candidate governance closure omitted the authenticated assessment index.')
    }
    $candidateId = [string]$Candidate['candidateId']
    if (-not $assessmentResult.CandidateById.ContainsKey($candidateId)) {
        throw [IO.InvalidDataException]::new(
            "Collect candidate '$candidateId' is missing from the assessment.")
    }
    return $assessmentResult.CandidateById[$candidateId]
}

function Get-RSProviderAssessmentPatch {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate
    )

    return (Get-RSProviderAssessmentCandidate `
        -Context $Context `
        -Candidate $Candidate)['rawLedger']['patch']
}

function Get-RSProviderCollectionDescriptor {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$CandidateAuthoredPath
    )

    $candidateId = [string]$Candidate['candidateId']
    $descriptorPath = Get-RSProviderContainedPath `
        -Root $SourceRoot `
        -RelativePath $script:RSProviderDescriptorLeaf `
        -Label "candidate '$candidateId' collection descriptor"
    Import-Module (Join-Path $RepoRoot (
            'Tools\TerminalEngine\TerminalEngineCollectionDescriptor.psm1'))
    $descriptor = Read-RSTerminalEngineCollectionDescriptor `
        -LiteralPath $descriptorPath `
        -CandidateId $candidateId `
        -SchemaFile (Join-Path `
            $RepoRoot `
            $script:RSProviderDescriptorSchemaLeaf) `
        -RepoRoot $RepoRoot
    Assert-RSTerminalEngineCollectionDescriptorPathClosure `
        -Descriptor $descriptor `
        -SourceRoot $SourceRoot `
        -CandidateAuthoredPath $CandidateAuthoredPath `
        -CandidateId $candidateId
    return $descriptor
}

function Assert-RSProviderDescriptorAssessmentTruth {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$PatchRecord,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId
    )

    $derivedWrapper =
        [string]$Descriptor.BuildKind -ceq 'candidate-wrapper-v1'
    if ([string]$PatchRecord['collectionDescriptorSha256'] -cne
            [string]$Descriptor.Sha256 -or
        [string]$PatchRecord['buildKind'] -cne
            [string]$Descriptor.BuildKind -or
        [bool]$PatchRecord['maintainedBuildWrapper'] -ne
            $derivedWrapper -or
        [bool]$Descriptor.MaintainedBuildWrapper -ne
            $derivedWrapper) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' descriptor build truth differs from its frozen assessment.")
    }
}

function Get-RSProviderArtifactContract {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor
    )

    $value = if ($Descriptor.PSObject.Properties.Name -ccontains 'Value') {
        $Descriptor.Value
    }
    else {
        $Descriptor
    }
    $build = $value['build']
    $kind = [string]$build['kind']
    $adapter = $build['adapter']
    $private = [object[]]@($build['privateRuntimeLibraries'])
    $privateLeaves = [string[]]@(
        $private | ForEach-Object { [string]$_['fileName'] })
    $privateLeaves = Get-RSProviderOrdinalStrings -Value $privateLeaves
    return [pscustomobject]@{
        BuildKind = $kind
        MaintainedBuildWrapper = $kind -ceq 'candidate-wrapper-v1'
        AdapterDllLeaf = [string]$adapter['dllLeaf']
        AdapterProjectLeaf = if (
            $kind -ceq 'provider-generic-cmake-msvc-v1') {
            [string]$adapter['projectLeaf']
        }
        else {
            ''
        }
        StaticLibraries = if (
            $kind -ceq 'provider-generic-cmake-msvc-v1') {
            [object[]]@($build['staticLibraries'])
        }
        else {
            [object[]]@()
        }
        PrivateRuntimeLibraries = $private
        PrivateRuntimeLeaves = $privateLeaves
    }
}

function Get-RSProviderArchiveKind {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    if ($FileName.EndsWith(
            '.nupkg',
            [StringComparison]::OrdinalIgnoreCase)) {
        return 'nuget'
    }
    if ($FileName.EndsWith(
            '.zip',
            [StringComparison]::OrdinalIgnoreCase)) {
        return 'zip'
    }
    if ($FileName.EndsWith(
            '.tar',
            [StringComparison]::OrdinalIgnoreCase) -or
        $FileName.EndsWith(
            '.tar.gz',
            [StringComparison]::OrdinalIgnoreCase) -or
        $FileName.EndsWith(
            '.tgz',
            [StringComparison]::OrdinalIgnoreCase)) {
        return 'tar'
    }
    throw [IO.InvalidDataException]::new(
        "Frozen archive '$FileName' has no provider-supported kind.")
}

function Get-RSProviderExpandedInput {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [string]$BootstrapRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputRecord
    )

    $source = Get-RSProviderInputPath `
        -BootstrapRoot $BootstrapRoot `
        -Platform $Context.Platform `
        -InputRecord $InputRecord
    if ((Get-RSProviderSha256 -LiteralPath $source) -cne
            [string]$InputRecord['sha256']) {
        throw [IO.InvalidDataException]::new(
            "Frozen input '$($InputRecord['inputId'])' changed before expansion.")
    }
    $destination = Get-RSProviderContainedPath `
        -Root $BootstrapRoot `
        -RelativePath (Join-Path `
            'expanded' `
            (Join-Path `
                $Context.Platform `
                ([string]$InputRecord['inputId']))) `
        -Label "expanded input '$($InputRecord['inputId'])'"
    Remove-RSProviderExactDirectoryTree `
        -LiteralPath $destination `
        -AllowedRoot $BootstrapRoot `
        -Label "expanded input '$($InputRecord['inputId'])'"
    Import-Module (Join-Path $Context.RepoRoot (
            'Tools\TerminalEngine\TerminalEngineArchive.psm1'))
    $kind = Get-RSProviderArchiveKind `
        -FileName ([string]$InputRecord['fileName'])
    $plan = Get-RSTerminalEngineArchivePlan `
        -LiteralPath $source `
        -ArchiveKind $kind
    $null = Expand-RSTerminalEngineArchive `
        -LiteralPath $source `
        -DestinationRoot $destination `
        -ArchiveKind $kind `
        -PreflightPlan $plan
    return [pscustomobject]@{
        Root = $destination
        IdentitySha256 = Get-RSProviderDirectoryIdentity `
            -LiteralPath $destination `
            -AllowedRoot $BootstrapRoot `
            -RepoRoot $Context.RepoRoot
    }
}

function Get-RSProviderExecutableCandidates {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context
    )

    return [object[]]@(
        [object[]]@($Context.CandidateSet['candidates']) |
            Where-Object {
                [string]$_['scoredOrControl'] -ceq 'scored' -and
                [string]$_['disposition'] -ceq 'collect'
            })
}

function Resolve-RSProviderHostHeaders {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [string]$BootstrapRoot
    )

    $executableCandidates = Get-RSProviderExecutableCandidates `
        -Context $Context
    if ($executableCandidates.Count -eq 0) {
        throw [IO.InvalidDataException]::new(
            'A static-only candidate set has no Gate-0 host-header consumer.')
    }
    $matches = @(
        $Context.Inputs | Where-Object {
            [string]$_['inputId'] -ceq
                $script:RSProviderHostHeadersInputId
        })
    if ($matches.Count -ne 1 -or
        [string]$matches[0]['inputKind'] -cne 'dependency-input' -or
        [string]$matches[0]['fileName'] -cne 'host-wil-headers.zip') {
        throw [IO.InvalidDataException]::new(
            "Executable collection requires exactly one frozen candidate-neutral '$($script:RSProviderHostHeadersInputId)' dependency input named 'host-wil-headers.zip'.")
    }
    $input = $matches[0]
    $expectedConsumers = Get-RSProviderOrdinalStrings -Value (
        [string[]]@(
            $executableCandidates | ForEach-Object {
                [string]$_['candidateId']
            }))
    $actualConsumerValues = [string[]]@($input['consumers'])
    $consumerSet =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
    foreach ($consumer in $actualConsumerValues) {
        if (-not $consumerSet.Add($consumer)) {
            throw [IO.InvalidDataException]::new(
                'The frozen host-header input repeats an executable consumer.')
        }
    }
    $actualConsumers = Get-RSProviderOrdinalStrings -Value (
        $actualConsumerValues)
    if (($actualConsumers -join "`0") -cne
            ($expectedConsumers -join "`0")) {
        throw [IO.InvalidDataException]::new(
            'The frozen host-header input must be shared by the exact executable candidate cohort.')
    }

    $expanded = Get-RSProviderExpandedInput `
        -Context $Context `
        -BootstrapRoot $BootstrapRoot `
        -InputRecord $input
    $includeRoot = Get-RSProviderContainedPath `
        -Root $expanded.Root `
        -RelativePath 'include' `
        -Label 'candidate-neutral host-header include root'
    $treeSha256 = Get-RSProviderDirectoryIdentity `
        -LiteralPath $includeRoot `
        -AllowedRoot $BootstrapRoot `
        -RepoRoot $Context.RepoRoot
    Import-Module (Join-Path $Context.RepoRoot (
            'Tools\TerminalEngine\TerminalEngineLifecycle.psm1'))
    $verified = Resolve-RSTerminalGate0HostHeaderRoot `
        -HostHeaderIncludeRoot $includeRoot `
        -ExpectedTreeSha256 $treeSha256
    return [pscustomobject]@{
        InputId = $script:RSProviderHostHeadersInputId
        InputSha256 = [string]$input['sha256']
        IncludeRoot = [string]$verified.IncludeRoot
        TreeSha256 = [string]$verified.TreeSha256
    }
}

function Assert-RSProviderHostHeadersIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [string]$BootstrapRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$HostHeaders
    )

    $includeRoot = [IO.Path]::GetFullPath(
        [string]$HostHeaders.IncludeRoot)
    Assert-RSProviderNoReparsePoint `
        -LiteralPath $includeRoot `
        -AllowedRoot $BootstrapRoot `
        -Label 'candidate-neutral host-header include root'
    Import-Module (Join-Path $Context.RepoRoot (
            'Tools\TerminalEngine\TerminalEngineLifecycle.psm1'))
    return Resolve-RSTerminalGate0HostHeaderRoot `
        -HostHeaderIncludeRoot $includeRoot `
        -ExpectedTreeSha256 ([string]$HostHeaders.TreeSha256)
}

function Resolve-RSProviderDescriptorHeaderRoots {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor,

        [Parameter(Mandatory = $true)]
        [string]$BootstrapRoot,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId
    )

    $result =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::Ordinal)
    foreach ($package in [object[]]@(
            $Descriptor.Value['headerPackages'])) {
        $inputId = [string]$package['inputId']
        if (-not $Context.InputById.ContainsKey($inputId)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' header package '$inputId' is not frozen.")
        }
        $input = $Context.InputById[$inputId]
        if ([object[]]@($input['consumers']) -cnotcontains $CandidateId) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' header package '$inputId' is not assigned to the candidate.")
        }
        $expanded = Get-RSProviderExpandedInput `
            -Context $Context `
            -BootstrapRoot $BootstrapRoot `
            -InputRecord $input
        $includeRoot = Get-RSProviderContainedPath `
            -Root $expanded.Root `
            -RelativePath (
                [string]$package['includeRelativePath']).Replace('/', '\') `
            -Label "candidate '$CandidateId' header include root"
        if (-not [IO.Directory]::Exists($includeRoot)) {
            throw [IO.DirectoryNotFoundException]::new(
                "Candidate '$CandidateId' header include root is missing.")
        }
        $result.Add($inputId, [pscustomobject]@{
                Path = $includeRoot
                IdentitySha256 = Get-RSProviderDirectoryIdentity `
                    -LiteralPath $includeRoot `
                    -AllowedRoot $BootstrapRoot `
                    -RepoRoot $Context.RepoRoot
            })
    }
    return $result
}

function Resolve-RSProviderLegalSource {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Record,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$BootstrapRoot
    )

    $source = $Record['source']
    switch ([string]$source['sourceKind']) {
        'candidate-source' {
            return Get-RSProviderContainedPath `
                -Root $SourceRoot `
                -RelativePath (
                    [string]$source['sourceRelativePath']).Replace('/', '\') `
                -Label 'candidate legal source'
        }
        'frozen-input-whole' {
            $inputId = [string]$source['sourceInputId']
            if (-not $Context.InputById.ContainsKey($inputId)) {
                throw [IO.InvalidDataException]::new(
                    "Legal input '$inputId' is not frozen.")
            }
            return Get-RSProviderInputPath `
                -BootstrapRoot $BootstrapRoot `
                -Platform $Context.Platform `
                -InputRecord $Context.InputById[$inputId]
        }
        'frozen-input-member' {
            $inputId = [string]$source['sourceInputId']
            if (-not $Context.InputById.ContainsKey($inputId)) {
                throw [IO.InvalidDataException]::new(
                    "Legal archive input '$inputId' is not frozen.")
            }
            $expanded = Get-RSProviderExpandedInput `
                -Context $Context `
                -BootstrapRoot $BootstrapRoot `
                -InputRecord $Context.InputById[$inputId]
            return Get-RSProviderContainedPath `
                -Root $expanded.Root `
                -RelativePath (
                    [string]$source['archivePath']).Replace('/', '\') `
                -Label 'frozen legal archive member'
        }
        default {
            throw [IO.InvalidDataException]::new(
                'Descriptor legal source kind is unsupported.')
        }
    }
}

function Publish-RSProviderDescriptorNotices {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$BootstrapRoot,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId
    )

    $legal = $Descriptor.Value['legal']
    $outputPath = Get-RSProviderContainedPath `
        -Root $SourceRoot `
        -RelativePath (
            [string]$legal['outputRelativePath']).Replace('/', '\') `
        -Label "candidate '$CandidateId' generated notice output"
    if ([IO.File]::Exists($outputPath) -or
        [IO.Directory]::Exists($outputPath)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' generated notice output was not absent.")
    }
    $chunks = [Collections.Generic.List[byte[]]]::new()
    $utf8 = [Text.UTF8Encoding]::new($false, $true)
    $chunks.Add($utf8.GetBytes(
            "RedSalamander Terminal Engine Third-Party Notices`n"))
    $identityRows = [Collections.Generic.List[object]]::new()
    foreach ($dependency in [object[]]@($legal['dependencies'])) {
        $chunks.Add($utf8.GetBytes((
                "`nDependency: {0}`nSPDX: {1}`n" -f
                    [string]$dependency['dependencyId'],
                    [string]$dependency['spdx'])))
        foreach ($record in [object[]]@($dependency['inputs'])) {
            $path = Resolve-RSProviderLegalSource `
                -Context $Context `
                -Record $record `
                -SourceRoot $SourceRoot `
                -BootstrapRoot $BootstrapRoot
            if (-not [IO.File]::Exists($path)) {
                throw [IO.FileNotFoundException]::new(
                    "Candidate '$CandidateId' legal input is missing.",
                    $path)
            }
            $item = Get-Item -LiteralPath $path -Force
            $sha256 = Get-RSProviderSha256 -LiteralPath $path
            if ($sha256 -cne [string]$record['sha256'] -or
                $item.Length -ne [int64]$record['sizeBytes']) {
                throw [IO.InvalidDataException]::new(
                    "Candidate '$CandidateId' legal input identity is invalid.")
            }
            $bytes = [IO.File]::ReadAllBytes($path)
            try {
                $null = $utf8.GetString($bytes)
            }
            catch {
                throw [IO.InvalidDataException]::new(
                    "Candidate '$CandidateId' legal input is not strict UTF-8 text.",
                    $_.Exception)
            }
            $chunks.Add($utf8.GetBytes((
                    "Input: {0} ({1})`n" -f
                        [string]$record['destinationLeaf'],
                        [string]$record['kind'])))
            $chunks.Add($bytes)
            if ($bytes.Length -eq 0 -or
                $bytes[$bytes.Length - 1] -ne 10) {
                $chunks.Add([byte[]]@(10))
            }
            $identityRows.Add([ordered]@{
                    dependencyId = [string]$dependency['dependencyId']
                    destinationLeaf =
                        [string]$record['destinationLeaf']
                    sha256 = $sha256
                    sizeBytes = ([uint64]$item.Length).ToString(
                        [Globalization.CultureInfo]::InvariantCulture)
                })
        }
    }
    $total = [int64]0
    foreach ($chunk in $chunks) {
        $total += $chunk.Length
    }
    if ($total -le 0 -or $total -gt 67108864) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' generated notice output exceeds its bound.")
    }
    [void][IO.Directory]::CreateDirectory(
        [IO.Path]::GetDirectoryName($outputPath))
    $stream = [IO.FileStream]::new(
        $outputPath,
        [IO.FileMode]::CreateNew,
        [IO.FileAccess]::Write,
        [IO.FileShare]::None)
    try {
        foreach ($chunk in $chunks) {
            $stream.Write($chunk, 0, $chunk.Length)
        }
        $stream.Flush($true)
    }
    finally {
        $stream.Dispose()
    }
    Import-Module (
        Join-Path $Context.RepoRoot 'Tools\TerminalEvidence.psm1')
    return [pscustomobject]@{
        OutputPath = $outputPath
        OutputSha256 = Get-RSProviderSha256 -LiteralPath $outputPath
        OutputSizeBytes = ([uint64]$total).ToString(
            [Globalization.CultureInfo]::InvariantCulture)
        ClosureId = [string]$legal['closureId']
        ClosureSha256 = Get-RSJcsSha256 `
            -InputObject ([object[]]$identityRows.ToArray())
    }
}

function Resolve-RSProviderPathBindings {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor,

        [Parameter(Mandatory = $true)]
        [string]$BootstrapRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$HeaderRoots,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$NoticeState,

        [Parameter(Mandatory = $true)]
        [string]$BuildRoot
    )

    $providerPaths = [ordered]@{
        'adapter-header' = Join-Path `
            $Context.RepoRoot `
            $script:RSProviderAdapterHeaderLeaf
        'build-receipt' = Join-Path `
            $BuildRoot `
            $script:RSProviderBuildReceiptLeaf
        'layout-header' = Join-Path `
            $Context.RepoRoot `
            $script:RSProviderLayoutHeaderLeaf
        'notice-output' = [string]$NoticeState.OutputPath
        'source-dependencies' = Join-Path `
            $BuildRoot `
            $script:RSProviderSourceDependenciesLeaf
    }
    $result = [Collections.Generic.List[object]]::new()
    foreach ($binding in [object[]]@(
            $Descriptor.Value['build']['pathBindings'])) {
        $kind = [string]$binding['valueKind']
        $path = ''
        $identity = ''
        switch ($kind) {
            'frozen-input-file' {
                $inputId = [string]$binding['inputId']
                if (-not $Context.InputById.ContainsKey($inputId)) {
                    throw [IO.InvalidDataException]::new(
                        "Path binding references unknown input '$inputId'.")
                }
                $path = Get-RSProviderInputPath `
                    -BootstrapRoot $BootstrapRoot `
                    -Platform $Context.Platform `
                    -InputRecord $Context.InputById[$inputId]
                $identity = Get-RSProviderSha256 -LiteralPath $path
            }
            'header-package-root' {
                $inputId = [string]$binding['inputId']
                if (-not $HeaderRoots.ContainsKey($inputId)) {
                    throw [IO.InvalidDataException]::new(
                        "Path binding references unavailable header root '$inputId'.")
                }
                $path = [string]$HeaderRoots[$inputId].Path
                $identity = [string]$HeaderRoots[$inputId].IdentitySha256
            }
            'provider-path' {
                $providerPath = [string]$binding['providerPath']
                if (-not $providerPaths.Contains($providerPath)) {
                    throw [IO.InvalidDataException]::new(
                        "Path binding references unknown provider path '$providerPath'.")
                }
                $path = [string]$providerPaths[$providerPath]
                if ([IO.File]::Exists($path)) {
                    $identity = Get-RSProviderSha256 -LiteralPath $path
                }
                else {
                    Import-Module (
                        Join-Path $Context.RepoRoot (
                            'Tools\TerminalEvidence.psm1'))
                    $identity = Get-RSJcsSha256 -InputObject ([ordered]@{
                            reservedPath = [IO.Path]::GetFullPath($path)
                            state = 'absent-provider-output'
                        })
                }
            }
        }
        $result.Add([ordered]@{
                bindingKey = [string]$binding['bindingKey']
                valueKind = $kind
                path = [IO.Path]::GetFullPath($path)
                identitySha256 = $identity
            })
    }
    return [object[]]$result.ToArray()
}

function Read-RSProviderWrapperInput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$ResultPath,

        [string]$ExpectedRepoRoot = ''
    )

    $repoRoot = if (-not [string]::IsNullOrWhiteSpace($ExpectedRepoRoot)) {
        Get-RSProviderCanonicalRepoRoot -Path $ExpectedRepoRoot
    }
    elseif (-not [string]::IsNullOrWhiteSpace(
            $script:RSProviderLoadedRepoRoot)) {
        $script:RSProviderLoadedRepoRoot
    }
    else {
        throw [IO.InvalidDataException]::new(
            'Wrapper input validation requires a loaded RepoRoot.')
    }
    $input = [IO.Path]::GetFullPath($InputPath)
    $result = [IO.Path]::GetFullPath($ResultPath)
    $parsed = Read-RSProviderCanonicalJson `
        -LiteralPath $input `
        -RepoRoot $repoRoot `
        -Label 'Candidate-wrapper input' `
        -MaximumBytes $script:RSProviderWrapperInputMaxBytes
    $json = [IO.File]::ReadAllText(
        $input,
        [Text.UTF8Encoding]::new($false, $true))
    if (-not (Test-Json `
            -Json $json `
            -SchemaFile (Join-Path `
                $repoRoot `
                $script:RSProviderWrapperInputSchemaLeaf) `
            -ErrorAction Stop)) {
        throw [IO.InvalidDataException]::new(
            'Candidate-wrapper input violates its closed schema.')
    }
    $buildRoot = [IO.Path]::GetFullPath(
        [string]$parsed['paths']['buildRoot'])
    if (-not $input.StartsWith(
            $buildRoot + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase) -or
        -not $result.StartsWith(
            $buildRoot + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw [IO.InvalidDataException]::new(
            'Candidate-wrapper input or result escaped its build root.')
    }
    foreach ($name in [string[]]@(
            $parsed['paths'].Keys | ForEach-Object { [string]$_ })) {
        $path = [string]$parsed['paths'][$name]
        if (-not [IO.Path]::IsPathFullyQualified($path) -or
            [IO.Path]::GetFullPath($path) -cne $path) {
            throw [IO.InvalidDataException]::new(
                "Candidate-wrapper path '$name' is not canonical and absolute.")
        }
    }
    if ([IO.File]::Exists($result) -or [IO.Directory]::Exists($result)) {
        throw [IO.InvalidDataException]::new(
            'Candidate-wrapper result must be absent before launch.')
    }
    return $parsed
}

function New-RSProviderWrapperInput {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$PatchRecord,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$BuildRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$NoticeState,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$HostHeaders,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$PathBindings
    )

    if ([string]$Descriptor.BuildKind -cne 'candidate-wrapper-v1') {
        throw [IO.InvalidDataException]::new(
            'Provider-generic builds must not create a candidate-wrapper input.')
    }
    $contract = Get-RSProviderArtifactContract -Descriptor $Descriptor
    $adapterHeader = Join-Path `
        $Context.RepoRoot `
        $script:RSProviderAdapterHeaderLeaf
    $layoutHeader = Join-Path `
        $Context.RepoRoot `
        $script:RSProviderLayoutHeaderLeaf
    $record = [ordered]@{
        formatId = 'red-salamander-terminal-engine-wrapper-input'
        version = 1
        candidate = [ordered]@{
            candidateId = [string]$Candidate['candidateId']
            pin = [string]$Candidate['pin']
            pinSha256 = [string]$Candidate['pinSha256']
            sourceArchiveSha256 =
                [string]$Candidate['sourceArchiveSha256']
            patchSha256 = [string]$PatchRecord['patchSha256']
            resultTree = [string]$PatchRecord['resultTree']
            collectionDescriptorSha256 =
                [string]$Descriptor.Sha256
            buildKind = 'candidate-wrapper-v1'
        }
        lane = [ordered]@{
            platform = [string]$Context.Platform
            configuration = [string]$Context.Configuration
        }
        paths = [ordered]@{
            sourceRoot = [IO.Path]::GetFullPath($SourceRoot)
            buildRoot = [IO.Path]::GetFullPath($BuildRoot)
            noticeOutputPath =
                [IO.Path]::GetFullPath([string]$NoticeState.OutputPath)
            adapterHeaderPath =
                [IO.Path]::GetFullPath($adapterHeader)
            layoutHeaderPath =
                [IO.Path]::GetFullPath($layoutHeader)
            hostHeaderIncludeRoot =
                [IO.Path]::GetFullPath(
                    [string]$HostHeaders.IncludeRoot)
            msbuildPath = [IO.Path]::GetFullPath(
                (Get-RSProviderMSBuildPath))
            buildReceiptPath = [IO.Path]::GetFullPath((
                    Join-Path `
                        $BuildRoot `
                        $script:RSProviderBuildReceiptLeaf))
            sourceDependenciesPath = [IO.Path]::GetFullPath((
                    Join-Path `
                        $BuildRoot `
                        $script:RSProviderSourceDependenciesLeaf))
        }
        pathBindings = [object[]]$PathBindings
        artifacts = [ordered]@{
            adapterDllLeaf = [string]$contract.AdapterDllLeaf
            privateRuntimeLibraries =
                [string[]]@($contract.PrivateRuntimeLeaves)
        }
        gate0 = [ordered]@{
            adapterHeaderSha256 =
                Get-RSProviderSha256 -LiteralPath $adapterHeader
            layoutHeaderSha256 =
                Get-RSProviderSha256 -LiteralPath $layoutHeader
            hostHeaderTreeSha256 =
                [string]$HostHeaders.TreeSha256
            capabilities = $script:RSProviderCapabilities.ToString(
                [Globalization.CultureInfo]::InvariantCulture)
        }
    }
    $inputPath = Get-RSProviderContainedPath `
        -Root $BuildRoot `
        -RelativePath $script:RSProviderWrapperInputLeaf `
        -Label 'candidate-wrapper input'
    $resultPath = Get-RSProviderContainedPath `
        -Root $BuildRoot `
        -RelativePath $script:RSProviderWrapperResultLeaf `
        -Label 'candidate-wrapper result'
    $written = Write-RSProviderCanonicalJson `
        -LiteralPath $inputPath `
        -Value $record `
        -RepoRoot $Context.RepoRoot `
        -MaximumBytes $script:RSProviderWrapperInputMaxBytes `
        -SchemaFile (Join-Path `
            $Context.RepoRoot `
            $script:RSProviderWrapperInputSchemaLeaf)
    $parsed = Read-RSProviderWrapperInput `
        -InputPath $inputPath `
        -ResultPath $resultPath `
        -ExpectedRepoRoot $Context.RepoRoot
    return [pscustomobject]@{
        Path = $written.Path
        ResultPath = $resultPath
        Record = $parsed
        Sha256 = $written.Sha256
        SizeBytes = $written.SizeBytes
    }
}

function Read-RSProviderWrapperResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResultPath,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$BuildRoot
    )

    $result = [IO.Path]::GetFullPath($ResultPath)
    Assert-RSProviderNoReparsePoint `
        -LiteralPath $result `
        -AllowedRoot $BuildRoot `
        -Label 'candidate-wrapper result'
    $parsed = Read-RSProviderCanonicalJson `
        -LiteralPath $result `
        -RepoRoot $RepoRoot `
        -Label 'Candidate-wrapper result' `
        -MaximumBytes $script:RSProviderWrapperResultMaxBytes
    $json = [IO.File]::ReadAllText(
        $result,
        [Text.UTF8Encoding]::new($false, $true))
    if (-not (Test-Json `
            -Json $json `
            -SchemaFile (Join-Path `
                $RepoRoot `
                $script:RSProviderWrapperResultSchemaLeaf) `
            -ErrorAction Stop)) {
        throw [IO.InvalidDataException]::new(
            'Candidate-wrapper result violates its closed schema.')
    }
    $item = Get-Item -LiteralPath $result -Force
    return [pscustomobject]@{
        Record = $parsed
        Status = [string]$parsed['status']
        Failure = $parsed['failure']
        Sha256 = Get-RSProviderSha256 -LiteralPath $result
        SizeBytes = ([uint64]$item.Length).ToString(
            [Globalization.CultureInfo]::InvariantCulture)
    }
}

function Invoke-RSProviderIsolation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$DriverPath,

        [Parameter(Mandatory = $true)]
        [string[]]$DriverArguments,

        [Parameter(Mandatory = $true)]
        [string]$WorkingDirectory
    )

    . (Join-Path `
        $RepoRoot `
        'Tools\TerminalEngine\Invoke-RSIsolatedProcess.ps1') -LoadOnly
    $pwsh = [IO.Path]::GetFullPath((
            [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName))
    return Invoke-RSIsolatedProcess `
        -FilePath $pwsh `
        -ArgumentList (@(
                '-NoLogo',
                '-NoProfile',
                '-NonInteractive',
                '-ExecutionPolicy',
                'Bypass',
                '-File',
                $DriverPath) + $DriverArguments) `
        -CanaryExecutablePath $pwsh `
        -WorkingDirectory $WorkingDirectory `
        -CommandTimeoutSeconds 7200 `
        -ExclusiveDisposableRunner
}

function Assert-RSProviderIsolationAttestation {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Attestation,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId
    )

    if ($null -eq $Attestation -or
        [string]$Attestation.providerStatus -cne 'pass' -or
        -not [bool]$Attestation.isolationVerified -or
        -not [bool]$Attestation.rootCommandExecuted -or
        [string]$Attestation.enforcement -cne 'runner-enforced' -or
        -not [bool]$Attestation.ruleSetupVerified -or
        -not [bool]$Attestation.ruleRemovalVerified -or
        -not [bool]$Attestation.programIdentityStable -or
        -not [bool]$Attestation.scopeContract.
            exclusiveDisposableRunnerAcknowledged -or
        [int]$Attestation.rootCommand.result.exitCode -ne 0 -or
        [bool]$Attestation.rootCommand.result.timedOut -or
        -not [bool]$Attestation.rootCommand.result.outputComplete -or
        [object[]]@($Attestation.failures).Count -ne 0) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' did not run under a complete runner-enforced isolation attestation.")
    }
}

function Invoke-RSProviderWrapperDriver {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$WrapperInput,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$BuildRoot
    )

    if ([string]$Descriptor.BuildKind -cne 'candidate-wrapper-v1') {
        throw [IO.InvalidDataException]::new(
            'Provider-generic builds must not invoke a candidate wrapper.')
    }
    $driver = Get-RSProviderContainedPath `
        -Root $SourceRoot `
        -RelativePath (
            [string]$Descriptor.Value['build']['driverRelativePath']).
                Replace('/', '\') `
        -Label 'candidate-wrapper driver'
    if (-not $driver.EndsWith(
            '.ps1',
            [StringComparison]::Ordinal) -or
        -not [IO.File]::Exists($driver)) {
        throw [IO.InvalidDataException]::new(
            'Candidate-wrapper driver must be the exact patched PowerShell script.')
    }
    $isolation = Invoke-RSProviderIsolation `
        -RepoRoot $Context.RepoRoot `
        -DriverPath $driver `
        -DriverArguments @(
            '-InputPath',
            [string]$WrapperInput.Path,
            '-ResultPath',
            [string]$WrapperInput.ResultPath) `
        -WorkingDirectory $SourceRoot
    Assert-RSProviderIsolationAttestation `
        -Attestation $isolation `
        -CandidateId ([string]$Candidate['candidateId'])
    $result = Read-RSProviderWrapperResult `
        -ResultPath $WrapperInput.ResultPath `
        -RepoRoot $Context.RepoRoot `
        -BuildRoot $BuildRoot
    return [pscustomobject]@{
        Isolation = $isolation
        Result = $result
    }
}

function Get-RSProviderMSBuildPath {
    $paths = @(
        Get-Command msbuild.exe `
            -CommandType Application `
            -ErrorAction SilentlyContinue)
    if ($paths.Count -ne 1) {
        throw [IO.InvalidDataException]::new(
            'The collection runner must expose exactly one MSBuild executable.')
    }
    return [IO.Path]::GetFullPath($paths[0].Source)
}

function Get-RSProviderCandidateBuildRoot {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId
    )

    Assert-RSProviderSafeSegment `
        -Value $CandidateId `
        -Label 'candidate build root id'
    $lane = Join-Path `
        $Context.RepoRoot `
        $script:RSProviderLaneLeaf
    return Get-RSProviderContainedPath `
        -Root $lane `
        -RelativePath (Join-Path `
            $Context.Platform `
            (Join-Path `
                (Get-RSProviderConfigurationLeaf `
                    -Configuration $Context.Configuration) `
                $CandidateId)) `
        -Label "candidate '$CandidateId' build root"
}

function Remove-RSProviderExactDirectoryTree {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $target = [IO.Path]::GetFullPath($LiteralPath)
    $root = [IO.Path]::GetFullPath($AllowedRoot).
        TrimEnd([IO.Path]::DirectorySeparatorChar)
    if (-not $target.StartsWith(
            $root + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase) -or
        $target.Equals($root, [StringComparison]::OrdinalIgnoreCase)) {
        throw [IO.InvalidDataException]::new(
            "$Label is not one bounded child directory.")
    }
    if (-not [IO.Directory]::Exists($target)) {
        return
    }
    Assert-RSProviderNoReparsePoint `
        -LiteralPath $target `
        -AllowedRoot $root `
        -Label $Label
    foreach ($item in @(
            Get-ChildItem -LiteralPath $target -Force -Recurse)) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw [IO.InvalidDataException]::new(
                "$Label contains a reparse point.")
        }
    }
    Remove-Item -LiteralPath $target -Recurse -Force
}

function Invoke-RSProviderClean {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context
    )

    $laneRoot = Join-Path `
        $Context.RepoRoot `
        $script:RSProviderLaneLeaf
    foreach ($candidateId in $Context.SelectedIds) {
        Remove-RSProviderExactDirectoryTree `
            -LiteralPath (Get-RSProviderCandidateBuildRoot `
                -Context $Context `
                -CandidateId $candidateId) `
            -AllowedRoot $laneRoot `
            -Label "candidate '$candidateId' lane build root"
    }
    return [pscustomobject]@{
        clean = $true
        neutralityBaselineProof = $Context.NeutralityBaselineProof
        neutralityProof = $Context.NeutralityProof
    }
}

function Assert-RSProviderArtifactClosure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BuildRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor
    )

    $contract = Get-RSProviderArtifactContract -Descriptor $Descriptor
    $expected =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    [void]$expected.Add([string]$contract.AdapterDllLeaf)
    foreach ($leaf in $contract.PrivateRuntimeLeaves) {
        if (-not $expected.Add([string]$leaf)) {
            throw [IO.InvalidDataException]::new(
                'Descriptor artifact leaves collide under Windows case folding.')
        }
    }
    $actual =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @(
            Get-ChildItem `
                -LiteralPath $BuildRoot `
                -File `
                -Force `
                -Recurse `
                -Filter '*.dll')) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
            -not $actual.TryAdd($item.Name, $item)) {
            throw [IO.InvalidDataException]::new(
                'Build artifact closure has a reparse point or case-folding duplicate.')
        }
    }
    if ($actual.Count -ne $expected.Count) {
        throw [IO.InvalidDataException]::new(
            'Build artifact DLL closure differs from the descriptor.')
    }
    foreach ($leaf in $expected) {
        if (-not $actual.ContainsKey($leaf) -or
            [string]$actual[$leaf].Name -cne $leaf) {
            throw [IO.InvalidDataException]::new(
                "Required build artifact '$leaf' is missing or has different casing.")
        }
    }
    $adapter = $actual[[string]$contract.AdapterDllLeaf]
    return [pscustomobject]@{
        AdapterDll = [string]$adapter.FullName
        AdapterSha256 = Get-RSProviderSha256 `
            -LiteralPath $adapter.FullName
        PrivateRuntimePaths = [string[]]@(
            $contract.PrivateRuntimeLeaves | ForEach-Object {
                [string]$actual[[string]$_].FullName
            })
    }
}

function New-RSProviderStaticCandidateResult {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate
    )

    $candidateId = [string]$Candidate['candidateId']
    if ([string]$Candidate['scoredOrControl'] -ceq 'control') {
        Import-Module (
            Join-Path $Context.RepoRoot 'Tools\TerminalEvidence.psm1')
        $evidenceSha256 = Get-RSJcsSha256 -InputObject ([ordered]@{
                bootstrapInputSetSha256 =
                    [string]$Context.BootstrapInputSetSha256
                candidateId = $candidateId
                candidateSetSha256 = [string]$Context.CandidateSetSha256
                configuration = [string]$Context.Configuration
                constraintId = [string]$Candidate['constraintId']
                pinSha256 = [string]$Candidate['pinSha256']
                platform = [string]$Context.Platform
            })
        return [ordered]@{
            candidateId = $candidateId
            pin = [string]$Candidate['pin']
            pinSha256 = [string]$Candidate['pinSha256']
            outcome = 'constraint-control'
            controlFailure = [ordered]@{
                constraintId = [string]$Candidate['constraintId']
                evidenceSha256 = $evidenceSha256
            }
        }
    }
    if ([string]$Candidate['disposition'] -ceq 'source-rejected') {
        return [ordered]@{
            candidateId = $candidateId
            pin = [string]$Candidate['pin']
            pinSha256 = [string]$Candidate['pinSha256']
            outcome = 'source-rejected'
            sourceRejection = $Candidate['sourceRejection']
        }
    }
    throw [IO.InvalidDataException]::new(
        "Candidate '$candidateId' requires executable collection.")
}

function Invoke-RSProviderNeutralityProof {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [AllowEmptyCollection()]
        [object[]]$Descriptor = @()
    )

    if ([string]$Context.NeutralityPolicySha256 -cnotmatch
            '^[0-9a-f]{64}$') {
        throw [IO.InvalidDataException]::new(
            'Provider context omitted its authenticated neutrality policy.')
    }
    return Invoke-RSProviderAuthenticatedNeutralityProof `
        -RepoRoot $Context.RepoRoot `
        -ExpectedNeutralityPolicySha256 (
            [string]$Context.NeutralityPolicySha256) `
        -ExpectedUniverseSha256 (
            [string]$Context.CandidateUniverseSha256) `
        -ExpectedCandidateSetSha256 (
            [string]$Context.CandidateSetSha256) `
        -ExpectedCandidateReviewSha256 (
            [string]$Context.CandidateReviewSha256) `
        -ExpectedCandidateAssessmentSha256 (
            [string]$Context.CandidateAssessmentSha256) `
        -Descriptor $Descriptor `
        -FrozenInput $Context.Inputs
}

function Invoke-RSProviderAuthenticatedNeutralityProof {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedNeutralityPolicySha256,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedUniverseSha256,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedCandidateSetSha256,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedCandidateReviewSha256,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedCandidateAssessmentSha256,

        [AllowEmptyCollection()]
        [object[]]$Descriptor = @(),

        [AllowEmptyCollection()]
        [object[]]$FrozenInput = @()
    )

    $evaluator = Get-RSProviderEvaluatorModule -RepoRoot $RepoRoot
    return & $evaluator {
        param(
            $root,
            $policySha256,
            $universeSha256,
            $candidateSetSha256,
            $candidateReviewSha256,
            $candidateAssessmentSha256,
            $descriptors,
            $frozenInputs)
        Invoke-RSAuthenticatedCollectionNeutralityProof `
            -RepoRoot $root `
            -ExpectedNeutralityPolicySha256 $policySha256 `
            -ExpectedUniverseSha256 $universeSha256 `
            -ExpectedCandidateSetSha256 $candidateSetSha256 `
            -ExpectedCandidateReviewSha256 $candidateReviewSha256 `
            -ExpectedCandidateAssessmentSha256 (
                $candidateAssessmentSha256) `
            -Descriptor $descriptors `
            -FrozenInput $frozenInputs
    } `
        $RepoRoot `
        $ExpectedNeutralityPolicySha256 `
        $ExpectedUniverseSha256 `
        $ExpectedCandidateSetSha256 `
        $ExpectedCandidateReviewSha256 `
        $ExpectedCandidateAssessmentSha256 `
        $Descriptor `
        $FrozenInput
}

function Invoke-RSProviderBootstrap {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context
    )

    $bootstrapRoot = Join-Path `
        $Context.RepoRoot `
        $script:RSProviderBootstrapLeaf
    [void][IO.Directory]::CreateDirectory($bootstrapRoot)
    foreach ($input in $Context.Inputs) {
        if (Test-RSProviderInputSelected `
                -InputRecord $input `
                -SelectedIds $Context.SelectedIds) {
            $null = Get-RSProviderVerifiedInput `
                -BootstrapRoot $bootstrapRoot `
                -Platform $Context.Platform `
                -InputRecord $input
        }
    }
    $selectedExecutable = @(
        $Context.SelectedCandidates | Where-Object {
            [string]$_['scoredOrControl'] -ceq 'scored' -and
            [string]$_['disposition'] -ceq 'collect'
        })
    if ($selectedExecutable.Count -ne 0) {
        $null = Resolve-RSProviderHostHeaders `
            -Context $Context `
            -BootstrapRoot $bootstrapRoot
    }
    return [pscustomobject]@{
        bootstrap = $true
        bootstrapInputSetSha256 = $Context.BootstrapInputSetSha256
        neutralityBaselineProof = $Context.NeutralityBaselineProof
        neutralityProof = $Context.NeutralityProof
    }
}

function Invoke-RSProviderCollectCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate
    )

    $candidateId = [string]$Candidate['candidateId']
    $bootstrapRoot = Join-Path `
        $Context.RepoRoot `
        $script:RSProviderBootstrapLeaf
    $sourceRoot = Get-RSProviderContainedPath `
        -Root (Join-Path $bootstrapRoot 'sources') `
        -RelativePath $candidateId `
        -Label "candidate '$candidateId' patched source"
    if (-not [IO.Directory]::Exists($sourceRoot)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' patched source is missing; Bootstrap is incomplete.")
    }
    $patch = Get-RSProviderAssessmentPatch `
        -Context $Context `
        -Candidate $Candidate
    $changedPath = [string[]]@(
        & git -C $sourceRoot diff --cached --name-only --no-renames)
    if ($LASTEXITCODE -ne 0) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' changed-path discovery failed.")
    }
    $descriptor = Get-RSProviderCollectionDescriptor `
        -RepoRoot $Context.RepoRoot `
        -Candidate $Candidate `
        -SourceRoot $sourceRoot `
        -CandidateAuthoredPath $changedPath
    Assert-RSProviderDescriptorAssessmentTruth `
        -Descriptor $descriptor `
        -PatchRecord $patch `
        -CandidateId $candidateId
    $hostHeaders = Resolve-RSProviderHostHeaders `
        -Context $Context `
        -BootstrapRoot $bootstrapRoot
    $headerRoots = Resolve-RSProviderDescriptorHeaderRoots `
        -Context $Context `
        -Descriptor $descriptor `
        -BootstrapRoot $bootstrapRoot `
        -CandidateId $candidateId
    $noticeState = Publish-RSProviderDescriptorNotices `
        -Context $Context `
        -Descriptor $descriptor `
        -SourceRoot $sourceRoot `
        -BootstrapRoot $bootstrapRoot `
        -CandidateId $candidateId
    $buildRoot = Get-RSProviderCandidateBuildRoot `
        -Context $Context `
        -CandidateId $candidateId
    if (-not [IO.Directory]::Exists($buildRoot) -or
        @(Get-ChildItem -LiteralPath $buildRoot -Force).Count -ne 0) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' lane build root is not clean.")
    }
    $bindings = Resolve-RSProviderPathBindings `
        -Context $Context `
        -Descriptor $descriptor `
        -BootstrapRoot $bootstrapRoot `
        -HeaderRoots $headerRoots `
        -NoticeState $noticeState `
        -BuildRoot $buildRoot
    if ([string]$descriptor.BuildKind -ceq 'candidate-wrapper-v1') {
        $wrapperInput = New-RSProviderWrapperInput `
            -Context $Context `
            -Candidate $Candidate `
            -Descriptor $descriptor `
            -PatchRecord $patch `
            -SourceRoot $sourceRoot `
            -BuildRoot $buildRoot `
            -NoticeState $noticeState `
            -HostHeaders $hostHeaders `
            -PathBindings $bindings
        $null = Assert-RSProviderHostHeadersIdentity `
            -Context $Context `
            -BootstrapRoot $bootstrapRoot `
            -HostHeaders $hostHeaders
        $wrapper = Invoke-RSProviderWrapperDriver `
            -Context $Context `
            -Candidate $Candidate `
            -Descriptor $descriptor `
            -WrapperInput $wrapperInput `
            -SourceRoot $sourceRoot `
            -BuildRoot $buildRoot
        $null = Assert-RSProviderHostHeadersIdentity `
            -Context $Context `
            -BootstrapRoot $bootstrapRoot `
            -HostHeaders $hostHeaders
        if ([string]$wrapper.Result.Status -cne 'pass') {
            throw [IO.InvalidDataException]::new(
                "Candidate '$candidateId' wrapper reported a frozen build failure.")
        }
    }
    else {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' provider-generic build kind '$($descriptor.BuildKind)' has no completed isolated dispatcher.")
    }
    $artifacts = Assert-RSProviderArtifactClosure `
        -BuildRoot $buildRoot `
        -Descriptor $descriptor
    return [pscustomobject]@{
        Candidate = [ordered]@{
            candidateId = $candidateId
            pin = [string]$Candidate['pin']
            pinSha256 = [string]$Candidate['pinSha256']
            outcome = 'provider-contract-verified'
            descriptorSha256 = [string]$descriptor.Sha256
            buildKind = [string]$descriptor.BuildKind
            adapterSha256 = [string]$artifacts.AdapterSha256
        }
    }
}

function Invoke-RSProviderCollect {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Context
    )

    $executable = @(
        $Context.SelectedCandidates | Where-Object {
            [string]$_['scoredOrControl'] -ceq 'scored' -and
            [string]$_['disposition'] -ceq 'collect'
        })
    if ($executable.Count -eq 0) {
        throw [IO.InvalidDataException]::new(
            'Collect requires at least one selected executable candidate; static records cannot attest a lane.')
    }
    throw [IO.InvalidDataException]::new(
        'Round-3 freezes provider-neutral contracts only; executable build-kind collection requires a new approved universe round.')
}

$script:RSProviderLoadedRepoRoot = if (
    [string]::IsNullOrWhiteSpace($RepoRoot)) {
    ''
}
else {
    Get-RSProviderCanonicalRepoRoot -Path $RepoRoot
}

if ($LoadOnly) {
    return
}
if ($Phase -cnotin @('Bootstrap', 'Clean', 'Collect')) {
    throw [ArgumentException]::new(
        'Phase must be exactly Bootstrap, Clean, or Collect.')
}
$context = Get-RSProviderContext `
    -RepoRoot $script:RSProviderLoadedRepoRoot `
    -CandidateId $CandidateId `
    -Platform $Platform `
    -Configuration $Configuration
$neutrality = Invoke-RSProviderNeutralityProof -Context $context
$context | Add-Member `
    -NotePropertyName NeutralityBaselineProof `
    -NotePropertyValue $neutrality.Baseline
$context | Add-Member `
    -NotePropertyName NeutralityProof `
    -NotePropertyValue $neutrality.Runtime
$result = switch ($Phase) {
    'Bootstrap' { Invoke-RSProviderBootstrap -Context $context }
    'Clean' { Invoke-RSProviderClean -Context $context }
    'Collect' { Invoke-RSProviderCollect -Context $context }
}
Write-Output -NoEnumerate $result
