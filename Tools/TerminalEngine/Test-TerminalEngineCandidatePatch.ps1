[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[a-z0-9][a-z0-9._-]{0,63}$')]
    [string]$CandidateId,

    [Parameter(Mandatory = $true)]
    [ValidateSet('collect', 'adaptation-rejected')]
    [string]$Disposition,

    [Parameter(Mandatory = $true)]
    [string]$SourceRepository,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{40}$')]
    [string]$UpstreamCommit,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{40}$')]
    [string]$ExpectedUpstreamTree,

    [Parameter(Mandatory = $true)]
    [string]$PatchPath,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$ExpectedPatchSha256,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{40}$')]
    [string]$ExpectedResultTree,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 10000000)]
    [int]$ExpectedChangedMaintainedLines,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 100000)]
    [int]$ExpectedChangedMaintainedFiles,

    [Parameter(Mandatory = $true)]
    [ValidateRange(0, 100000)]
    [int]$ExpectedBinaryFiles,

    [Parameter(Mandatory = $true)]
    [string]$ConcernMappingPath,

    [AllowEmptyString()]
    [ValidateScript({
            [string]::IsNullOrEmpty($_) -or
            [string]$_ -cmatch '^[0-9a-f]{64}$'
        })]
    [string]$ExpectedCollectionDescriptorSha256 = '',

    [string]$OutputPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-RSCandidatePatchGit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Repository,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $output = @(& git -C $Repository @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) {
        $message = ($output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
        throw [InvalidOperationException]::new(
            "$Context failed with exit code $LASTEXITCODE.$([Environment]::NewLine)$message")
    }
    return ,([string[]]@($output | ForEach-Object { [string]$_ }))
}

function Get-RSCandidatePatchSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath
    )

    return (Get-FileHash -LiteralPath $LiteralPath -Algorithm SHA256).
        Hash.ToLowerInvariant()
}

function Get-RSCandidatePatchOrdinalStrings {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Value,

        [switch]$Unique
    )

    $sorted = [string[]]@($Value)
    [Array]::Sort($sorted, [StringComparer]::Ordinal)
    if (-not $Unique) {
        return ,$sorted
    }
    $result = [Collections.Generic.List[string]]::new()
    foreach ($entry in $sorted) {
        if ($result.Count -eq 0 -or
            -not [StringComparer]::Ordinal.Equals(
                $result[$result.Count - 1],
                $entry)) {
            $result.Add($entry)
        }
    }
    return ,$result.ToArray()
}

function Assert-RSCandidatePatchContainedFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $resolved = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $LiteralPath).Path)
    $root = [IO.Path]::GetFullPath($AllowedRoot)
    if (-not $resolved.StartsWith(
            $root + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase) -or
        ((Get-Item -LiteralPath $resolved -Force).Attributes -band
            [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw [IO.InvalidDataException]::new(
            "$Label must be a non-reparse file below '$root'.")
    }
    return $resolved
}

$repoRoot = [IO.Path]::GetFullPath((
        Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$source = [IO.Path]::GetFullPath((
        Resolve-Path -LiteralPath $SourceRepository).Path)
$insideWorkTree = Invoke-RSCandidatePatchGit `
    -Repository $source `
    -Arguments @('rev-parse', '--is-inside-work-tree') `
    -Context 'Git repository validation'
if ($insideWorkTree.Count -ne 1 -or $insideWorkTree[0] -cne 'true') {
    throw [IO.InvalidDataException]::new(
        "SourceRepository '$source' is not a Git work tree.")
}

$candidatePatchRoot = Join-Path $repoRoot 'External\TerminalEngine\CandidatePatches'
$resolvedPatch = Assert-RSCandidatePatchContainedFile `
    -LiteralPath $PatchPath `
    -AllowedRoot $candidatePatchRoot `
    -Label 'PatchPath'
$resolvedMapping = Assert-RSCandidatePatchContainedFile `
    -LiteralPath $ConcernMappingPath `
    -AllowedRoot $candidatePatchRoot `
    -Label 'ConcernMappingPath'
if ((Get-RSCandidatePatchSha256 -LiteralPath $resolvedPatch) -cne
    $ExpectedPatchSha256) {
    throw [IO.InvalidDataException]::new(
        "Candidate '$CandidateId' patch SHA-256 differs from the frozen value.")
}

$scratchRoot = [IO.Path]::GetFullPath((
        Join-Path ([IO.Path]::GetTempPath()) (
            'RedSalamander\TerminalEnginePatchProof')))
[void](New-Item -ItemType Directory -Path $scratchRoot -Force)
$scratch = [IO.Path]::GetFullPath((
        Join-Path $scratchRoot (
            "$CandidateId-$([Guid]::NewGuid().ToString('N').Substring(0, 12))")))
if (-not $scratch.StartsWith(
        $scratchRoot + [IO.Path]::DirectorySeparatorChar,
        [StringComparison]::OrdinalIgnoreCase)) {
    throw [IO.InvalidDataException]::new('Patch-proof scratch path escaped its root.')
}
[void](New-Item -ItemType Directory -Path $scratch)
$indexPath = Join-Path $scratch 'candidate.index'
$regeneratedPatch = Join-Path $scratch 'regenerated.patch'
$hadIndexEnvironment = Test-Path Env:GIT_INDEX_FILE
$previousIndexEnvironment = if ($hadIndexEnvironment) {
    [string]$env:GIT_INDEX_FILE
} else {
    ''
}

try {
    $env:GIT_INDEX_FILE = $indexPath
    [void](Invoke-RSCandidatePatchGit `
            -Repository $source `
            -Arguments @('read-tree', $UpstreamCommit) `
            -Context 'Reading the exact upstream tree into an isolated index')
    $upstreamTree = (Invoke-RSCandidatePatchGit `
            -Repository $source `
            -Arguments @('rev-parse', "$UpstreamCommit^{tree}") `
            -Context 'Resolving the exact upstream tree')[0].Trim()
    if ($upstreamTree -cne $ExpectedUpstreamTree) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' upstream tree differs from the frozen value.")
    }

    [void](Invoke-RSCandidatePatchGit `
            -Repository $source `
            -Arguments @(
                '-c',
                'core.autocrlf=false',
                'apply',
                '--cached',
                '--binary',
                '--whitespace=nowarn',
                $resolvedPatch) `
            -Context 'Applying the candidate patch to the isolated index')
    $resultTree = (Invoke-RSCandidatePatchGit `
            -Repository $source `
            -Arguments @('write-tree') `
            -Context 'Writing the candidate result tree')[0].Trim()
    if ($resultTree -cne $ExpectedResultTree) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' result tree differs from the frozen value.")
    }

    $numstat = Invoke-RSCandidatePatchGit `
        -Repository $source `
        -Arguments @(
            '-c',
            'core.autocrlf=false',
            'diff',
            '--cached',
            '--no-renames',
            '--numstat',
            $UpstreamCommit) `
        -Context 'Measuring candidate patch line counts'
    $nameStatus = Invoke-RSCandidatePatchGit `
        -Repository $source `
        -Arguments @(
            '-c',
            'core.autocrlf=false',
            'diff',
            '--cached',
            '--no-renames',
            '--name-status',
            $UpstreamCommit) `
        -Context 'Measuring candidate patch path counts'

    [int64]$changedLines = 0
    [int]$binaryFiles = 0
    $numstatPaths = [Collections.Generic.List[string]]::new()
    foreach ($line in $numstat) {
        if ($line -cnotmatch '^(?<added>-|[0-9]+)\t(?<deleted>-|[0-9]+)\t(?<path>.+)$') {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' produced an invalid numstat row.")
        }
        $path = [string]$Matches.path
        $numstatPaths.Add($path)
        if ($Matches.added -ceq '-' -or $Matches.deleted -ceq '-') {
            $binaryFiles++
        } else {
            $changedLines += [int64]$Matches.added + [int64]$Matches.deleted
        }
    }
    $changedPaths = [Collections.Generic.List[string]]::new()
    foreach ($line in $nameStatus) {
        if ($line -cnotmatch '^[AMDTU]\t(?<path>.+)$') {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' produced a rename or invalid name-status row.")
        }
        $changedPaths.Add([string]$Matches.path)
    }
    $orderedChangedPaths = Get-RSCandidatePatchOrdinalStrings `
        -Value $changedPaths.ToArray()
    $orderedNumstatPaths = Get-RSCandidatePatchOrdinalStrings `
        -Value $numstatPaths.ToArray()
    if (-not [Linq.Enumerable]::SequenceEqual(
            $orderedChangedPaths,
            $orderedNumstatPaths)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' numstat and name-status path sets differ.")
    }
    if ($changedLines -ne $ExpectedChangedMaintainedLines -or
        $changedPaths.Count -ne $ExpectedChangedMaintainedFiles -or
        $binaryFiles -ne $ExpectedBinaryFiles) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' patch metrics differ from the frozen values.")
    }

    [void](Invoke-RSCandidatePatchGit `
            -Repository $source `
            -Arguments @(
                '-c',
                'core.autocrlf=false',
                'diff',
                '--cached',
                '--binary',
                '--full-index',
                '--no-renames',
                "--output=$regeneratedPatch",
                $UpstreamCommit) `
            -Context 'Regenerating the canonical candidate patch')
    if ((Get-RSCandidatePatchSha256 -LiteralPath $regeneratedPatch) -cne
        $ExpectedPatchSha256) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' patch is not the exact canonical no-renames diff.")
    }

    Import-Module (Join-Path $repoRoot 'Tools\TerminalEvidence.psm1') -Force
    $mappingBytes = [IO.File]::ReadAllBytes($resolvedMapping)
    $mapping = [Text.Encoding]::UTF8.GetString($mappingBytes) |
        ConvertFrom-Json -AsHashtable -Depth 64
    $canonicalMappingBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $mapping
    if (-not [Linq.Enumerable]::SequenceEqual(
            [byte[]]$mappingBytes,
            [byte[]]$canonicalMappingBytes)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' concern mapping is not canonical JCS.")
    }
    $mappingJson = [Text.Encoding]::UTF8.GetString($canonicalMappingBytes)
    if (-not (Test-Json `
            -Json $mappingJson `
            -SchemaFile (Join-Path $repoRoot (
                'Specs\Terminal\TerminalEngineConcernMapping.schema.json')) `
            -ErrorAction SilentlyContinue)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' concern mapping fails its closed schema.")
    }
    foreach ($name in @('candidateId', 'upstreamTree', 'resultTree', 'patchSha256')) {
        $expected = switch ($name) {
            'candidateId' { $CandidateId }
            'upstreamTree' { $ExpectedUpstreamTree }
            'resultTree' { $ExpectedResultTree }
            'patchSha256' { $ExpectedPatchSha256 }
        }
        if ([string]$mapping[$name] -cne $expected) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' concern mapping $name differs.")
        }
    }
    $mappedPaths = [string[]]@(
        $mapping.pathConcerns | ForEach-Object { [string]$_.path })
    $orderedMappedPaths = Get-RSCandidatePatchOrdinalStrings -Value $mappedPaths
    if (-not [Linq.Enumerable]::SequenceEqual(
            $orderedMappedPaths,
            $mappedPaths) -or
        -not [Linq.Enumerable]::SequenceEqual(
            $orderedMappedPaths,
            $orderedChangedPaths)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' concern mapping is not an exact sorted patch-path closure.")
    }
    $logicalConcerns = Get-RSCandidatePatchOrdinalStrings `
        -Value ([string[]]@(
                $mapping.pathConcerns.concerns |
                    ForEach-Object { [string]$_ })) `
        -Unique
    foreach ($entry in $mapping.pathConcerns) {
        $entryConcerns = [string[]]@($entry.concerns)
        $sortedEntryConcerns = Get-RSCandidatePatchOrdinalStrings `
            -Value $entryConcerns `
            -Unique
        if (-not [Linq.Enumerable]::SequenceEqual(
                $entryConcerns,
                $sortedEntryConcerns)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' concern mapping contains unsorted concerns.")
        }
    }
    $mappingPayload = [ordered]@{
        candidateId = [string]$mapping.candidateId
        upstreamTree = [string]$mapping.upstreamTree
        resultTree = [string]$mapping.resultTree
        patchSha256 = [string]$mapping.patchSha256
        pathConcerns = [object[]]@($mapping.pathConcerns)
    }
    $mappingSha256 = Get-RSJcsSha256 -InputObject $mappingPayload
    if ([string]$mapping.mappingSha256 -cne $mappingSha256) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' concern mapping digest is invalid.")
    }

    $resultRoot = Join-Path $scratch 'result-tree'
    [void](New-Item -ItemType Directory -Path $resultRoot)
    $checkoutPrefix = $resultRoot.Replace('\', '/') + '/'
    [void](Invoke-RSCandidatePatchGit `
            -Repository $source `
            -Arguments @(
                'checkout-index',
                '--all',
                '--force',
                "--prefix=$checkoutPrefix") `
            -Context 'Materializing the exact candidate result tree')
    $descriptorModule = Join-Path $repoRoot (
        'Tools\TerminalEngine\TerminalEngineCollectionDescriptor.psm1')
    Import-Module $descriptorModule -Force
    $descriptorRelativePath =
        Get-RSTerminalEngineCollectionDescriptorLeaf
    $descriptorPath = Join-Path $resultRoot (
        $descriptorRelativePath.Replace('/', '\'))
    $descriptor = $null
    if ($Disposition -ceq 'collect') {
        if ([string]::IsNullOrEmpty(
                $ExpectedCollectionDescriptorSha256)) {
            throw [IO.InvalidDataException]::new(
                "Collect candidate '$CandidateId' requires a frozen collection descriptor SHA-256.")
        }
        $descriptor = Read-RSTerminalEngineCollectionDescriptor `
            -LiteralPath $descriptorPath `
            -CandidateId $CandidateId `
            -SchemaFile (Join-Path $repoRoot (
                    'Specs\Terminal\TerminalEngineCollectionDescriptor.schema.json')) `
            -RepoRoot $repoRoot
        if ([string]$descriptor.Sha256 -cne
            $ExpectedCollectionDescriptorSha256) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' collection descriptor SHA-256 differs from the frozen value.")
        }
        Assert-RSTerminalEngineCollectionDescriptorPathClosure `
            -Descriptor $descriptor `
            -SourceRoot $resultRoot `
            -CandidateAuthoredPath $orderedChangedPaths `
            -CandidateId $CandidateId
    }
    else {
        if (-not [string]::IsNullOrEmpty(
                $ExpectedCollectionDescriptorSha256)) {
            throw [IO.InvalidDataException]::new(
                "Adaptation-rejected candidate '$CandidateId' must not declare a collection descriptor SHA-256.")
        }
        if ([IO.File]::Exists($descriptorPath) -or
            [IO.Directory]::Exists($descriptorPath)) {
            throw [IO.InvalidDataException]::new(
                "Adaptation-rejected candidate '$CandidateId' must not contain a collection descriptor.")
        }
    }

    $proofPayload = [ordered]@{
        formatId = 'red-salamander-terminal-engine-patch-proof'
        version = 1
        candidateId = $CandidateId
        disposition = $Disposition
        upstreamCommit = $UpstreamCommit
        upstreamTree = $upstreamTree
        patchSha256 = $ExpectedPatchSha256
        patchSizeBytes = (Get-Item -LiteralPath $resolvedPatch -Force).Length
        resultTree = $resultTree
        changedMaintainedLines = $changedLines
        changedMaintainedFiles = $changedPaths.Count
        binaryFiles = $binaryFiles
        logicalConcerns = $logicalConcerns
        concernMappingSha256 = $mappingSha256
        status = 'verified'
    }
    if ($Disposition -ceq 'collect') {
        $proofPayload.Insert(
            $proofPayload.Count - 1,
            'collectionDescriptorRelativePath',
            $descriptorRelativePath)
        $proofPayload.Insert(
            $proofPayload.Count - 1,
            'collectionDescriptorSha256',
            [string]$descriptor.Sha256)
        $proofPayload.Insert(
            $proofPayload.Count - 1,
            'buildKind',
            [string]$descriptor.BuildKind)
        $proofPayload.Insert(
            $proofPayload.Count - 1,
            'maintainedBuildWrapper',
            [bool]$descriptor.MaintainedBuildWrapper)
    }
    $proof = [ordered]@{}
    foreach ($entry in $proofPayload.GetEnumerator()) {
        $proof[$entry.Key] = $entry.Value
    }
    $proof.proofSha256 = Get-RSJcsSha256 -InputObject $proofPayload
    if (-not [string]::IsNullOrEmpty($OutputPath)) {
        $resolvedOutput = [IO.Path]::GetFullPath($OutputPath)
        $parent = [IO.Path]::GetDirectoryName($resolvedOutput)
        if (-not [string]::IsNullOrEmpty($parent)) {
            [void](New-Item -ItemType Directory -Path $parent -Force)
        }
        [IO.File]::WriteAllBytes(
            $resolvedOutput,
            (ConvertTo-RSJcsUtf8Bytes -InputObject $proof))
    }
    return [pscustomobject]$proof
}
finally {
    if ($hadIndexEnvironment) {
        $env:GIT_INDEX_FILE = $previousIndexEnvironment
    } else {
        Remove-Item Env:GIT_INDEX_FILE -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $scratch) {
        $resolvedScratch = [IO.Path]::GetFullPath((
                Resolve-Path -LiteralPath $scratch).Path)
        if (-not $resolvedScratch.StartsWith(
                $scratchRoot + [IO.Path]::DirectorySeparatorChar,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw [IO.InvalidDataException]::new(
                'Refusing to remove a patch-proof scratch path outside its root.')
        }
        Remove-Item -LiteralPath $resolvedScratch -Recurse -Force
    }
}
