Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ValidationFingerprint.psm1') -Scope Local -ErrorAction Stop

$script:MaximumValidationEntryArtifactBytes = 64MB
$script:MaximumValidationRunArtifactBytes = 1GB
$script:ValidationFingerprintComponentReasons = [ordered]@{
    build_artifact = 'BUILD_ARTIFACT_CHANGED'
    logical_input = 'LOGICAL_INPUT_CHANGED'
    shared_contract = 'SHARED_CONTRACT_CHANGED'
    runner = 'RUNNER_CHANGED'
    environment = 'ENVIRONMENT_CHANGED'
    arguments = 'ARGUMENTS_CHANGED'
    coverage = 'COVERAGE_CHANGED'
    outcome_policy = 'OUTCOME_POLICY_CHANGED'
    capability = 'CAPABILITY_CHANGED'
    artifact_epoch = 'ARTIFACT_EPOCH_CHANGED'
    source_input = 'SOURCE_INPUT_CHANGED'
}

if (-not ('RedSalamander.ValidationEvidenceNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace RedSalamander
{
    [StructLayout(LayoutKind.Sequential)]
    public struct ValidationEvidenceFileInformation
    {
        public uint FileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWriteTime;
        public uint VolumeSerialNumber;
        public uint FileSizeHigh;
        public uint FileSizeLow;
        public uint NumberOfLinks;
        public uint FileIndexHigh;
        public uint FileIndexLow;
    }

    public static class ValidationEvidenceNative
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern SafeFileHandle CreateFileW(
            string fileName,
            uint desiredAccess,
            uint shareMode,
            IntPtr securityAttributes,
            uint creationDisposition,
            uint flagsAndAttributes,
            IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetFileInformationByHandle(
            SafeFileHandle file,
            out ValidationEvidenceFileInformation information);
    }
}
'@ -ErrorAction Stop
}

function Assert-RSValidationRunId {
    param([Parameter(Mandatory = $true)][string]$RunId)
    if ($RunId -cnotmatch '^[0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32}$') {
        throw "Invalid validation run ID '$RunId'."
    }
}

function Assert-RSValidationEvidenceRoot {
    param([Parameter(Mandatory = $true)][string]$EvidenceRoot)

    $root = [IO.Path]::GetFullPath($EvidenceRoot).TrimEnd('\')
    $anchor = [IO.Path]::GetPathRoot($root)
    $current = Get-Item -LiteralPath $anchor -Force
    foreach ($segment in $root.Substring($anchor.Length).Split('\', [StringSplitOptions]::RemoveEmptyEntries)) {
        $next = Join-Path $current.FullName $segment
        if (-not (Test-Path -LiteralPath $next)) { break }
        $current = Get-Item -LiteralPath $next -Force
        if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Validation evidence root contains a reparse point: $($current.FullName)"
        }
    }
    return $root
}

function Get-RSValidationDirectoryIdentity {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $item = Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
    if (-not $item.PSIsContainer -or
        ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Validation evidence identity requires a non-reparse directory: $LiteralPath"
    }

    $shareReadWriteDelete = [uint32]0x00000007
    $openExisting = [uint32]3
    $openReparsePointAndDirectory = [uint32]0x02200000
    $handle = [RedSalamander.ValidationEvidenceNative]::CreateFileW(
        $item.FullName,
        0,
        $shareReadWriteDelete,
        [IntPtr]::Zero,
        $openExisting,
        $openReparsePointAndDirectory,
        [IntPtr]::Zero)
    if ($handle.IsInvalid) {
        $errorCode = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        $handle.Dispose()
        throw [ComponentModel.Win32Exception]::new(
            $errorCode,
            "Unable to open validation evidence directory identity '$($item.FullName)'.")
    }

    try {
        $information = [RedSalamander.ValidationEvidenceFileInformation]::new()
        if (-not [RedSalamander.ValidationEvidenceNative]::GetFileInformationByHandle(
                $handle,
                [ref]$information)) {
            $errorCode = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            throw [ComponentModel.Win32Exception]::new(
                $errorCode,
                "Unable to read validation evidence directory identity '$($item.FullName)'.")
        }
        if (($information.FileAttributes -band [uint32][IO.FileAttributes]::ReparsePoint) -ne 0 -or
            ($information.FileAttributes -band [uint32][IO.FileAttributes]::Directory) -eq 0) {
            throw "Validation evidence identity changed to an unsafe filesystem object: $($item.FullName)"
        }
        return ('{0:x8}:{1:x8}{2:x8}' -f
            $information.VolumeSerialNumber,
            $information.FileIndexHigh,
            $information.FileIndexLow)
    }
    finally {
        $handle.Dispose()
    }
}

function Get-RSValidationRunsRoot {
    param(
        [Parameter(Mandatory = $true)][string]$EvidenceRoot,
        [switch]$AllowMissing
    )

    $root = Assert-RSValidationEvidenceRoot -EvidenceRoot $EvidenceRoot
    $runsRoot = [IO.Path]::GetFullPath((Join-Path $root 'runs')).TrimEnd('\')
    if (-not (Test-Path -LiteralPath $runsRoot)) {
        if ($AllowMissing) { return $null }
        throw "Validation evidence runs root does not exist: $runsRoot"
    }
    $runsItem = Get-Item -LiteralPath $runsRoot -Force -ErrorAction Stop
    if (-not $runsItem.PSIsContainer -or
        ($runsItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Validation evidence runs root must be a non-reparse directory: $runsRoot"
    }
    return $runsItem.FullName.TrimEnd('\')
}

function Resolve-RSValidationRunDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$EvidenceRoot,
        [Parameter(Mandatory = $true)][string]$RunReference,
        [Parameter(Mandatory = $true)][string]$SchemaRoot
    )

    $runsRoot = Get-RSValidationRunsRoot -EvidenceRoot $EvidenceRoot
    $candidate = $null
    if ($RunReference -cmatch '^[0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32}$') {
        Assert-RSValidationRunId -RunId $RunReference
        $candidate = Join-Path $runsRoot $RunReference
    } elseif ([IO.Path]::IsPathFullyQualified($RunReference)) {
        $requested = $RunReference.TrimEnd('\')
        $normalized = [IO.Path]::GetFullPath($RunReference).TrimEnd('\')
        if (-not [string]::Equals($requested, $normalized, [StringComparison]::Ordinal)) {
            throw "Validation run reference contains a path alias: $RunReference"
        }
        $candidate = $normalized
    } else {
        throw "Validation run reference must be one exact run ID or fully qualified run directory: $RunReference"
    }

    $parent = [IO.Path]::GetFullPath((Split-Path $candidate -Parent)).TrimEnd('\')
    if (-not [string]::Equals($parent, $runsRoot, [StringComparison]::Ordinal)) {
        throw "Validation run directory must be an exact direct child of '$runsRoot': $candidate"
    }
    $item = Get-Item -LiteralPath $candidate -Force -ErrorAction Stop
    if (-not $item.PSIsContainer -or
        ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Validation run directory must be a non-reparse directory: $candidate"
    }
    if (-not [string]::Equals($candidate, $item.FullName.TrimEnd('\'), [StringComparison]::Ordinal)) {
        throw "Validation run directory reference does not match the filesystem name exactly: $candidate"
    }

    $leaf = $item.Name
    Assert-RSValidationRunId -RunId $leaf
    $state = Read-RSValidationJsonFile -LiteralPath (Join-Path $item.FullName 'run-state.json') `
        -SchemaPath (Join-Path $SchemaRoot 'ValidationRunState.schema.json') -AllowedRoot $item.FullName
    if (-not [string]::Equals($leaf, [string]$state.run_id, [StringComparison]::Ordinal)) {
        throw "Validation run directory leaf '$leaf' does not match run state '$($state.run_id)'."
    }
    return [pscustomobject]@{
        RunDirectory = $item.FullName.TrimEnd('\')
        State = $state
        DirectoryIdentity = Get-RSValidationDirectoryIdentity -LiteralPath $item.FullName
    }
}

function Resolve-RSValidationPrunePlanEntry {
    param(
        [Parameter(Mandatory = $true)][string]$RunsRoot,
        [Parameter(Mandatory = $true)][object]$Entry
    )

    if ($null -eq $Entry.PSObject.Properties['LiteralPath'] -or
        $null -eq $Entry.PSObject.Properties['RunId'] -or
        $null -eq $Entry.PSObject.Properties['DirectoryIdentity']) {
        throw 'Validation evidence prune apply requires an exact plan entry.'
    }
    $runId = [string]$Entry.RunId
    $plannedIdentity = [string]$Entry.DirectoryIdentity
    Assert-RSValidationRunId -RunId $runId
    if ([string]::IsNullOrWhiteSpace($plannedIdentity)) {
        throw "Validation evidence prune plan has no directory identity for '$runId'."
    }

    $full = [IO.Path]::GetFullPath([string]$Entry.LiteralPath).TrimEnd('\')
    $parent = [IO.Path]::GetFullPath((Split-Path $full -Parent)).TrimEnd('\')
    if (-not [string]::Equals($parent, $RunsRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Validation evidence prune target must be a direct child of the runs root: $full"
    }
    $expectedPath = [IO.Path]::GetFullPath((Join-Path $RunsRoot $runId)).TrimEnd('\')
    if (-not [string]::Equals($full, $expectedPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Validation evidence prune target path does not match its run ID '$runId': $full"
    }

    $item = Get-Item -LiteralPath $full -Force -ErrorAction Stop
    if (-not $item.PSIsContainer -or
        ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Unsafe validation evidence prune target: $full"
    }
    $currentIdentity = Get-RSValidationDirectoryIdentity -LiteralPath $item.FullName
    if (-not [string]::Equals($currentIdentity, $plannedIdentity, [StringComparison]::Ordinal)) {
        throw "Validation evidence prune target identity changed after planning: $full"
    }
    return $item.FullName
}

function Get-RSValidationRecordDigest {
    param([Parameter(Mandatory = $true)][object]$Record, [Parameter(Mandatory = $true)][string[]]$ExcludedProperty)

    $projection = [ordered]@{}
    foreach ($property in $Record.PSObject.Properties) {
        if ($property.Name -notin $ExcludedProperty) { $projection[$property.Name] = $property.Value }
    }
    if ($Record -is [Collections.IDictionary]) {
        $projection = [ordered]@{}
        foreach ($key in $Record.Keys) {
            if ($key -notin $ExcludedProperty) { $projection[$key] = $Record[$key] }
        }
    }
    return Get-RSCanonicalJsonDigest -InputObject $projection
}

function Get-RSValidationArtifactIdentity {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $item = Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
    if ($item.PSIsContainer -or
        ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Validation result artifact must be a regular non-reparse file: $LiteralPath"
    }
    $stream = [IO.FileStream]::new($item.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $digest = [Convert]::ToHexString($sha256.ComputeHash($stream)).ToLowerInvariant()
        return [pscustomobject]@{ SizeBytes = [int64]$stream.Length; Sha256 = $digest }
    } finally {
        $sha256.Dispose()
        $stream.Dispose()
    }
}

function Publish-RSValidationResultArtifacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Run,
        [Parameter(Mandatory = $true)][string]$EntryId,
        [AllowNull()][AllowEmptyCollection()][string[]]$CandidatePath = @(),
        [AllowNull()][scriptblock]$AfterCopyTestHook = $null
    )

    Assert-RSValidationIdentifier -Id $EntryId
    if ($null -eq $Run.PSObject.Properties['ResultArtifactBytes']) {
        $Run | Add-Member -NotePropertyName ResultArtifactBytes -NotePropertyValue ([int64]0)
    }
    $sources = @()
    $names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $entryBytes = [int64]0
    foreach ($candidate in @($CandidatePath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $item = Get-Item -LiteralPath $candidate -Force -ErrorAction Stop
        if ($item.PSIsContainer -or
            ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Validation result artifact must be a regular non-reparse file: $candidate"
        }
        if (-not $names.Add($item.Name)) {
            throw "Validation result artifacts contain a duplicate Windows filename '$($item.Name)'."
        }
        $entryBytes += [int64]$item.Length
        if ($entryBytes -gt $script:MaximumValidationEntryArtifactBytes) {
            throw "Validation result artifacts for '$EntryId' exceed the 64 MiB entry quota."
        }
        $sources += $item
    }
    if ([int64]$Run.ResultArtifactBytes + $entryBytes -gt $script:MaximumValidationRunArtifactBytes) {
        throw "Validation result artifacts exceed the 1 GiB run quota."
    }
    if ($sources.Count -eq 0) { return @() }

    $resultRoot = Join-Path (Join-Path (Join-Path $Run.RunDirectory 'entries') $EntryId) 'results'
    [void](New-Item -ItemType Directory -Path $resultRoot -Force)
    $publishedPaths = [Collections.Generic.List[string]]::new()
    $records = [Collections.Generic.List[object]]::new()
    try {
        foreach ($source in $sources) {
            $destination = Join-Path $resultRoot $source.Name
            if (Test-Path -LiteralPath $destination) {
                throw "Validation result artifact destination already exists: $destination"
            }
            $temporary = Join-Path $resultRoot ('.{0}.{1}.tmp' -f $source.Name, [guid]::NewGuid().ToString('N'))
            $sourceStream = $null
            $destinationStream = $null
            $sha256 = $null
            try {
                $sourceStream = [IO.FileStream]::new($source.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
                $destinationStream = [IO.FileStream]::new($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
                $sha256 = [Security.Cryptography.SHA256]::Create()
                $buffer = [byte[]]::new(1MB)
                while (($read = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    [void]$sha256.TransformBlock($buffer, 0, $read, $null, 0)
                    $destinationStream.Write($buffer, 0, $read)
                }
                [void]$sha256.TransformFinalBlock([byte[]]::new(0), 0, 0)
                $copiedSize = [int64]$sourceStream.Length
                $copiedDigest = [Convert]::ToHexString($sha256.Hash).ToLowerInvariant()
                $destinationStream.Flush($true)
            } finally {
                if ($null -ne $sha256) { $sha256.Dispose() }
                if ($null -ne $destinationStream) { $destinationStream.Dispose() }
                if ($null -ne $sourceStream) { $sourceStream.Dispose() }
            }
            try {
                if ($null -ne $AfterCopyTestHook) {
                    & $AfterCopyTestHook $source.FullName $temporary
                }
                $sourceIdentity = Get-RSValidationArtifactIdentity -LiteralPath $source.FullName
                if ($sourceIdentity.SizeBytes -ne $copiedSize -or $sourceIdentity.Sha256 -cne $copiedDigest) {
                    throw "Validation result artifact source changed while being preserved: $($source.FullName)"
                }
                [IO.File]::Move($temporary, $destination)
                $publishedPaths.Add($destination)
                $destinationIdentity = Get-RSValidationArtifactIdentity -LiteralPath $destination
                if ($destinationIdentity.SizeBytes -ne $copiedSize -or $destinationIdentity.Sha256 -cne $copiedDigest) {
                    throw "Validation result artifact changed after atomic publication: $destination"
                }
                $records.Add([pscustomobject][ordered]@{
                        path = "entries/$EntryId/results/$($source.Name)"
                        size_bytes = $copiedSize
                        sha256 = $copiedDigest
                    })
            } finally {
                if (Test-Path -LiteralPath $temporary -PathType Leaf) {
                    Remove-Item -LiteralPath $temporary -Force
                }
            }
        }
    } catch {
        foreach ($publishedPath in $publishedPaths) {
            if (Test-Path -LiteralPath $publishedPath -PathType Leaf) {
                Remove-Item -LiteralPath $publishedPath -Force
            }
        }
        throw
    }
    $Run.ResultArtifactBytes = [int64]$Run.ResultArtifactBytes + $entryBytes
    return @($records | Sort-Object path)
}

function Publish-RSValidationDecisionRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RunDirectory,
        [Parameter(Mandatory = $true)][object]$Decision,
        [Parameter(Mandatory = $true)][string]$SchemaRoot
    )

    $runRoot = [IO.Path]::GetFullPath($RunDirectory).TrimEnd('\')
    $computedDigest = Get-RSValidationRecordDigest -Record $Decision -ExcludedProperty @('decision_digest')
    if ([string]$Decision.decision_digest -cne $computedDigest) {
        throw "Validation decision content does not match decision_digest '$($Decision.decision_digest)'."
    }
    $decisionDirectory = Join-Path $runRoot 'decisions'
    if (-not (Test-Path -LiteralPath $decisionDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $decisionDirectory)
    }
    $path = Join-Path $decisionDirectory ($computedDigest + '.json')
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        $stored = Read-RSValidationJsonFile -LiteralPath $path `
            -SchemaPath (Join-Path $SchemaRoot 'ValidationPlanDecision.schema.json') `
            -AllowedRoot $runRoot
        if ((Get-RSCanonicalJsonDigest -InputObject $stored) -cne
            (Get-RSCanonicalJsonDigest -InputObject $Decision)) {
            throw "Immutable validation decision already exists with different bytes: $path"
        }
        return $path
    }
    try {
        Write-RSAtomicCanonicalJsonFile -LiteralPath $path -InputObject $Decision `
            -SchemaPath (Join-Path $SchemaRoot 'ValidationPlanDecision.schema.json') `
            -AllowedRoot $runRoot -CreateNew
    } catch [IO.IOException] {
        # A concurrent publisher may have won the create-only race. Accept only
        # the identical canonical record; a different or corrupt record fails.
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw }
        $stored = Read-RSValidationJsonFile -LiteralPath $path `
            -SchemaPath (Join-Path $SchemaRoot 'ValidationPlanDecision.schema.json') `
            -AllowedRoot $runRoot
        if ((Get-RSCanonicalJsonDigest -InputObject $stored) -cne
            (Get-RSCanonicalJsonDigest -InputObject $Decision)) {
            throw "Immutable validation decision already exists with different bytes: $path"
        }
    }
    return $path
}

function Read-RSValidationDecisionRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RunDirectory,
        [Parameter(Mandatory = $true)][string]$DecisionDigest,
        [Parameter(Mandatory = $true)][string]$SchemaRoot
    )

    if ($DecisionDigest -cnotmatch '^[0-9a-f]{64}$') {
        throw "Validation decision digest is invalid: '$DecisionDigest'."
    }
    $runRoot = [IO.Path]::GetFullPath($RunDirectory).TrimEnd('\')
    $path = Join-Path $runRoot "decisions\$DecisionDigest.json"
    $decision = Read-RSValidationJsonFile -LiteralPath $path `
        -SchemaPath (Join-Path $SchemaRoot 'ValidationPlanDecision.schema.json') `
        -AllowedRoot $runRoot
    $computedDigest = Get-RSValidationRecordDigest -Record $decision -ExcludedProperty @('decision_digest')
    if ([string]$decision.decision_digest -cne $DecisionDigest -or $computedDigest -cne $DecisionDigest) {
        throw "Validation decision record does not match its content-addressed path: $path"
    }
    return $decision
}

function New-RSValidationEvidenceRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$EvidenceRoot,
        [Parameter(Mandatory = $true)][string]$RunId,
        [Parameter(Mandatory = $true)][object]$Plan,
        [Parameter(Mandatory = $true)][object]$Decision,
        [Parameter(Mandatory = $true)][object]$BuildReceipt,
        [Parameter(Mandatory = $true)][string]$SchemaRoot,
        [ValidateSet('Fresh', 'Resume', 'Affected')][string]$ValidationMode = 'Fresh',
        [ValidateRange(0, [int64]::MaxValue)][int64]$ArtifactEpoch = 0,
        [AllowNull()][string]$CampaignId = $null,
        [ValidateRange(1, [int64]::MaxValue)][int64]$CampaignSequence = 1,
        [AllowNull()][string]$ParentRunId = $null
    )

    Assert-RSValidationRunId -RunId $RunId
    if (-not [string]::IsNullOrEmpty($CampaignId)) { Assert-RSValidationRunId -RunId $CampaignId }
    if (-not [string]::IsNullOrEmpty($ParentRunId)) { Assert-RSValidationRunId -RunId $ParentRunId }
    if ($Plan.requested_suite -notin @('CI', 'Full')) { throw 'Durable validation evidence is limited to CI or Full plans.' }
    $root = Assert-RSValidationEvidenceRoot -EvidenceRoot $EvidenceRoot
    [void](New-Item -ItemType Directory -Path (Join-Path $root 'runs') -Force)
    $runDirectory = Join-Path (Join-Path $root 'runs') $RunId
    if (Test-Path -LiteralPath $runDirectory) { throw "Validation run already exists: $RunId" }
    [void](New-Item -ItemType Directory -Path $runDirectory)
    [void](New-Item -ItemType Directory -Path (Join-Path $runDirectory 'entries'))
    [void](New-Item -ItemType Directory -Path (Join-Path $runDirectory 'decisions'))
    Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $runDirectory 'validation-plan.json') `
        -InputObject $Plan -SchemaPath (Join-Path $SchemaRoot 'ValidationPlan.schema.json') -AllowedRoot $runDirectory
    $decisionPath = Publish-RSValidationDecisionRecord -RunDirectory $runDirectory `
        -Decision $Decision -SchemaRoot $SchemaRoot
    $entries = @($Decision.entries | ForEach-Object {
            $isReuse = ([string]$_.decision -eq 'REUSE')
            [pscustomobject][ordered]@{
                id = $_.id
                ordinal = [int64]$_.ordinal
                decision = if ($isReuse) { 'reuse' } else { 'execute' }
                status = if ($isReuse) { 'reused' } else { 'pending' }
                reason_codes = @($_.reason_codes)
                expected_fingerprint = $_.expected_fingerprint
                origin_run_id = if ($isReuse) { $_.origin_run_id } else { $null }
                origin_evidence_digest = if ($isReuse) { $_.origin_evidence_digest } else { $null }
                promotion_digest = if ($isReuse -and $null -ne $_.PSObject.Properties['origin_promotion_digest']) { $_.origin_promotion_digest } else { $null }
            }
        })
    $state = [pscustomobject][ordered]@{
        schema = 'red-salamander.validation-run-state.v1'
        canonicalization = 'rs-canonical-json-v1'
        run_id = $RunId
        campaign_id = if ([string]::IsNullOrEmpty($CampaignId)) { $RunId } else { $CampaignId }
        campaign_sequence = $CampaignSequence
        parent_run_id = if ([string]::IsNullOrEmpty($ParentRunId)) { $null } else { $ParentRunId }
        requested_suite = $Plan.requested_suite
        validation_mode = $ValidationMode
        snapshot_id = $Decision.snapshot_id
        build_receipt_id = $BuildReceipt.receipt_id
        artifact_epoch = $ArtifactEpoch
        mutation_state = 'stable'
        decision_digest = $Decision.decision_digest
        plan_digest = $Plan.plan_digest
        coverage_plan_digest = $Plan.coverage_plan_digest
        status = 'running'
        entries = $entries
    }
    Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $runDirectory 'run-state.json') `
        -InputObject $state -SchemaPath (Join-Path $SchemaRoot 'ValidationRunState.schema.json') -AllowedRoot $runDirectory
    return [pscustomobject]@{
        EvidenceRoot = $root; RunDirectory = $runDirectory; State = $state
        SchemaRoot = $SchemaRoot; DecisionPath = $decisionPath; ResultArtifactBytes = [int64]0
    }
}

function Write-RSValidationRunState {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Run)
    Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $Run.RunDirectory 'run-state.json') `
        -InputObject $Run.State -SchemaPath (Join-Path $Run.SchemaRoot 'ValidationRunState.schema.json') `
        -AllowedRoot $Run.RunDirectory
}

function Set-RSValidationEntryCheckpoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Run,
        [Parameter(Mandatory = $true)][string]$EntryId,
        [ValidateSet('pending', 'running', 'provisional-passed', 'promoted', 'failed', 'interrupted', 'invalidated', 'reused')]
        [string]$Status,
        [AllowNull()][object]$PromotionDigest = $null
    )

    $entry = @($Run.State.entries | Where-Object id -eq $EntryId)
    if ($entry.Count -ne 1) { throw "Validation run does not contain entry '$EntryId'." }
    $entry[0].status = $Status
    $entry[0].promotion_digest = if ($null -eq $PromotionDigest -or [string]::IsNullOrEmpty([string]$PromotionDigest)) {
        $null
    } else { [string]$PromotionDigest }
    Write-RSValidationRunState -Run $Run
}

function Get-RSValidationResultValue {
    param(
        [Parameter(Mandatory = $true)][object]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$Name,
        [AllowNull()][object]$DefaultValue = $null
    )

    foreach ($candidate in $Name) {
        if ($InputObject -is [Collections.IDictionary]) {
            if ($InputObject.Contains($candidate)) { return $InputObject[$candidate] }
        } elseif ($null -ne $InputObject.PSObject.Properties[$candidate]) {
            return $InputObject.PSObject.Properties[$candidate].Value
        }
    }
    return $DefaultValue
}

function New-RSValidationTerminalResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$EntryContract,
        [Parameter(Mandatory = $true)][object]$Result,
        [Parameter(Mandatory = $true)][bool]$ResultParsed,
        [Parameter(Mandatory = $true)][bool]$CoverageValid,
        [Parameter(Mandatory = $true)][bool]$RequiredArtifactsPresent
    )

    $entryId = [string](Get-RSValidationResultValue -InputObject $EntryContract -Name @('id', 'Id') -DefaultValue '')
    $kind = [string](Get-RSValidationResultValue -InputObject $EntryContract -Name @('kind', 'Kind') -DefaultValue '')
    Assert-RSValidationIdentifier -Id $entryId
    if ($kind -notin @('SelfTest', 'Executable', 'CppUnitTest', 'Pester', 'PowerShellScript')) {
        throw "Validation terminal result has unsupported kind '$kind'."
    }

    $reasons = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $exitCode = [int](Get-RSValidationResultValue -InputObject $Result -Name @('ExitCode', 'exit_code') -DefaultValue 1)
    $classification = [string](Get-RSValidationResultValue -InputObject $Result -Name @('Classification', 'classification') -DefaultValue '')
    if ($exitCode -ne 0) { [void]$reasons.Add('NONZERO_EXIT') }
    if ($classification -in @('FLAKY', 'REGRESSION', 'ISOLATION_SUSPECT', 'UNCLASSIFIED_FAILURE')) {
        [void]$reasons.Add('FAILURE_CLASSIFICATION')
    }
    if (-not $RequiredArtifactsPresent) { [void]$reasons.Add('REQUIRED_ARTIFACT_MISSING') }
    if (-not $CoverageValid) { [void]$reasons.Add('COVERAGE_INVALID') }

    $declaredPassed = [int64](Get-RSValidationResultValue -InputObject $Result -Name @('Passed', 'passed', 'PassedCount') -DefaultValue 0)
    $declaredFailed = [int64](Get-RSValidationResultValue -InputObject $Result -Name @('Failed', 'failed', 'FailedCount') -DefaultValue 0)
    $declaredSkipped = [int64](Get-RSValidationResultValue -InputObject $Result -Name @('Skipped', 'skipped', 'SkippedCount') -DefaultValue 0)
    $declaredTotal = [int64](Get-RSValidationResultValue -InputObject $Result -Name @('Total', 'total', 'TotalCount') `
        -DefaultValue ($declaredPassed + $declaredFailed + $declaredSkipped))
    if ($declaredPassed -lt 0 -or $declaredFailed -lt 0 -or $declaredSkipped -lt 0 -or $declaredTotal -lt 0) {
        [void]$reasons.Add('COUNT_INVALID')
    }

    if ($kind -in @('SelfTest', 'Pester')) {
        if (-not $ResultParsed) { [void]$reasons.Add('RESULT_NOT_PARSED') }
        if ($declaredTotal -eq 0) { [void]$reasons.Add('NO_TESTS_DISCOVERED') }
        if ($declaredPassed -eq 0 -and $declaredSkipped -gt 0 -and $declaredFailed -eq 0) {
            [void]$reasons.Add('ALL_TESTS_SKIPPED')
        }
        if ($declaredFailed -gt 0) { [void]$reasons.Add('DECLARED_FAILURES') }
        $pending = [int64](Get-RSValidationResultValue -InputObject $Result -Name @('Pending', 'PendingCount') -DefaultValue 0)
        $inconclusive = [int64](Get-RSValidationResultValue -InputObject $Result -Name @('Inconclusive', 'InconclusiveCount') -DefaultValue 0)
        if ($pending -gt 0 -or $inconclusive -gt 0) { [void]$reasons.Add('INCOMPLETE_TESTS') }
    }

    if ($kind -eq 'SelfTest') {
        $cases = @(Get-RSValidationResultValue -InputObject $Result -Name @('Cases', 'cases') -DefaultValue @())
        if ($cases.Count -eq 0) { [void]$reasons.Add('NO_CASE_RESULTS') }
        $actualPassed = [int64]0
        $actualFailed = [int64]0
        $actualSkipped = [int64]0
        $invalidStatus = $false
        $unallowedSkip = $false
        $forbiddenSkip = $false
        $outcome = Get-RSValidationResultValue -InputObject $EntryContract -Name @('outcome_contract', 'OutcomeContract') -DefaultValue ([pscustomobject]@{})
        $allowedSkipCases = @(
            @(Get-RSValidationResultValue -InputObject $outcome -Name @('static_skip_cases') -DefaultValue @()) +
            @(Get-RSValidationResultValue -InputObject $outcome -Name @('capability_bound_skip_cases') -DefaultValue @())
        )
        $forbiddenSkipClasses = @(Get-RSValidationResultValue -InputObject $outcome -Name @('forbidden_skip_classes') -DefaultValue @())
        foreach ($case in $cases) {
            $caseStatus = [string](Get-RSValidationResultValue -InputObject $case -Name @('status') -DefaultValue '')
            switch ($caseStatus) {
                'passed' { ++$actualPassed }
                'failed' { ++$actualFailed }
                'crashed' { ++$actualFailed }
                'skipped' {
                    ++$actualSkipped
                    $caseName = [string](Get-RSValidationResultValue -InputObject $case -Name @('name') -DefaultValue '')
                    $skipClass = [string](Get-RSValidationResultValue -InputObject $case -Name @('skip_class', 'skipClass') -DefaultValue '')
                    if ($skipClass -in $forbiddenSkipClasses) { $forbiddenSkip = $true }
                    if ($caseName -notin $allowedSkipCases -and $skipClass -notin $allowedSkipCases) { $unallowedSkip = $true }
                }
                default { $invalidStatus = $true }
            }
        }
        if ($invalidStatus) { [void]$reasons.Add('CASE_STATUS_INVALID') }
        if ($forbiddenSkip) { [void]$reasons.Add('FORBIDDEN_SKIP_CLASS') }
        if ($unallowedSkip) { [void]$reasons.Add('UNALLOWED_SKIP') }
        if ($actualPassed -ne $declaredPassed -or $actualFailed -ne $declaredFailed -or
            $actualSkipped -ne $declaredSkipped -or ($actualPassed + $actualFailed + $actualSkipped) -ne $declaredTotal) {
            [void]$reasons.Add('COUNT_MISMATCH')
        }
    } elseif ($kind -eq 'Pester' -and $declaredSkipped -gt 0) {
        # Pester entries currently declare no named skip allowlist. A skip is
        # incomplete coverage until an explicit result/outcome contract exists.
        [void]$reasons.Add('UNALLOWED_SKIP')
    }

    $terminal = [ordered]@{
        schema = 'red-salamander.validation-terminal-result.v1'
        canonicalization = 'rs-canonical-json-v1'
        terminal_result_digest = '0' * 64
        entry_id = $entryId
        kind = $kind
        status = if ($reasons.Count -eq 0) { 'PASSED' } else { 'FAILED' }
        reason_codes = @($reasons | Sort-Object)
        counts = [ordered]@{
            passed = [int64][Math]::Max(0, $declaredPassed)
            failed = [int64][Math]::Max(0, $declaredFailed)
            skipped = [int64][Math]::Max(0, $declaredSkipped)
            total = [int64][Math]::Max(0, $declaredTotal)
        }
    }
    $terminal.terminal_result_digest = Get-RSValidationRecordDigest -Record $terminal `
        -ExcludedProperty @('terminal_result_digest')
    $schemaRoot = Join-Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) 'Specs\Testing'
    $json = ConvertTo-RSCanonicalJson -InputObject $terminal
    if (-not (Test-Json -Json $json -SchemaFile (Join-Path $schemaRoot 'ValidationTerminalResult.schema.json') -ErrorAction SilentlyContinue)) {
        throw "Normalized validation terminal result does not satisfy its schema for '$entryId'."
    }
    return [pscustomobject]$terminal
}

function Publish-RSValidationEntryEvidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Run,
        [Parameter(Mandatory = $true)][object]$EntryContract,
        [Parameter(Mandatory = $true)][string]$EntryFingerprint,
        [string]$EntryFingerprintComponentVersion = 'validation-entry-components.v1',
        [AllowNull()][object]$EntryFingerprintComponents = $null,
        [Parameter(Mandatory = $true)][object]$TerminalResult,
        [Parameter(Mandatory = $true)][object]$PreSnapshot,
        [Parameter(Mandatory = $true)][object]$PostSnapshot,
        [Parameter(Mandatory = $true)][object]$BuildReceipt,
        [Parameter(Mandatory = $true)][datetime]$StartedUtc,
        [Parameter(Mandatory = $true)][datetime]$EndedUtc,
        [AllowNull()][AllowEmptyCollection()][object[]]$ResultArtifact = @(),
        [AllowNull()][AllowEmptyCollection()][string[]]$RequiredArtifact = @()
    )

    Assert-RSValidationIdentifier -Id ([string]$EntryContract.id)
    if ($EntryFingerprint -notmatch '^[0-9a-f]{64}$') { throw 'Entry evidence requires a canonical fingerprint.' }
    if ($EntryFingerprintComponentVersion -cne 'validation-entry-components.v1') {
        throw "Unsupported validation fingerprint component version '$EntryFingerprintComponentVersion'."
    }
    if ($null -eq $EntryFingerprintComponents) {
        $fallbackComponents = [ordered]@{}
        foreach ($name in $script:ValidationFingerprintComponentReasons.Keys) {
            $fallbackComponents[$name] = $EntryFingerprint
        }
        $EntryFingerprintComponents = [pscustomobject]$fallbackComponents
    }
    $componentNames = @($EntryFingerprintComponents.PSObject.Properties.Name | Sort-Object)
    $expectedComponentNames = @($script:ValidationFingerprintComponentReasons.Keys | Sort-Object)
    if ((Get-RSCanonicalJsonDigest -InputObject $componentNames) -cne
        (Get-RSCanonicalJsonDigest -InputObject $expectedComponentNames)) {
        throw 'Validation fingerprint component set is incomplete or unknown.'
    }
    foreach ($component in $EntryFingerprintComponents.PSObject.Properties) {
        if ([string]$component.Value -cnotmatch '^[0-9a-f]{64}$') {
            throw "Validation fingerprint component '$($component.Name)' is not a lowercase SHA-256 digest."
        }
    }
    # Explicitly passing an empty PowerShell array binds as $null across a
    # function boundary. Normalize both optional collections before StrictMode
    # property access; the runner always passes these parameters explicitly.
    $resultArtifacts = @($ResultArtifact | Where-Object {
            $null -ne $_ -and -not ($_ -is [string] -and [string]::IsNullOrWhiteSpace($_))
        })
    $requiredArtifacts = @($RequiredArtifact | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $resultPaths = @()
    foreach ($artifact in $resultArtifacts) {
        if ($null -eq $artifact.PSObject.Properties['path'] -or
            $null -eq $artifact.PSObject.Properties['size_bytes'] -or
            $null -eq $artifact.PSObject.Properties['sha256']) {
            throw 'Validation result artifact records require path, size_bytes, and sha256.'
        }
        $resultPaths += [string]$artifact.path
        $identity = Get-RSValidationArtifactIdentity -LiteralPath (Join-Path $Run.RunDirectory ([string]$artifact.path))
        if ($identity.SizeBytes -ne [int64]$artifact.size_bytes -or $identity.Sha256 -cne [string]$artifact.sha256) {
            throw "Validation result artifact does not match its content binding: $($artifact.path)"
        }
    }
    if ($resultPaths.Count -gt 0) { Assert-RSValidationPathSet -Path $resultPaths }
    if ($requiredArtifacts.Count -gt 0) { Assert-RSValidationPathSet -Path $requiredArtifacts }
    $missingRequired = @($requiredArtifacts | Where-Object {
            -not (Test-Path -LiteralPath (Join-Path $Run.RunDirectory $_) -PathType Leaf)
        })
    if ($TerminalResult.entry_id -ne $EntryContract.id -or
        $TerminalResult.terminal_result_digest -notmatch '^[0-9a-f]{64}$' -or
        (Get-RSValidationRecordDigest -Record $TerminalResult -ExcludedProperty @('terminal_result_digest')) -ne $TerminalResult.terminal_result_digest) {
        throw "Validation terminal result identity is invalid for '$($EntryContract.id)'."
    }
    $passed = $TerminalResult.status -eq 'PASSED'
    $barrierPassed = $PreSnapshot.snapshot_id -eq $PostSnapshot.snapshot_id -and $missingRequired.Count -eq 0
    $status = if ($passed -and $barrierPassed) { 'provisional-passed' } else { 'failed' }
    $artifactDigests = [ordered]@{}
    foreach ($output in @($BuildReceipt.outputs | Sort-Object id)) { $artifactDigests[[string]$output.id] = [string]$output.sha256 }
    $coverageDigest = Get-RSCanonicalJsonDigest -InputObject $EntryContract.coverage_contract
    $outcomeDigest = Get-RSCanonicalJsonDigest -InputObject $EntryContract.outcome_contract
    $duration = [int64][Math]::Max(0, [Math]::Round(($EndedUtc - $StartedUtc).TotalMilliseconds))
    $evidence = [ordered]@{
        schema = 'red-salamander.validation-entry-evidence.v1'
        canonicalization = 'rs-canonical-json-v1'
        evidence_digest = '0' * 64
        run_id = $Run.State.run_id
        entry_id = $EntryContract.id
        entry_fingerprint = $EntryFingerprint
        fingerprint_component_version = $EntryFingerprintComponentVersion
        fingerprint_components = $EntryFingerprintComponents
        plan_digest = $Run.State.plan_digest
        coverage_plan_digest = $Run.State.coverage_plan_digest
        terminal_result_digest = $TerminalResult.terminal_result_digest
        pre_snapshot_id = $PreSnapshot.snapshot_id
        post_snapshot_id = $PostSnapshot.snapshot_id
        build_receipt_id = $BuildReceipt.receipt_id
        artifact_epoch = [int64]$Run.State.artifact_epoch
        artifact_digests = $artifactDigests
        coverage_digest = $coverageDigest
        outcome_digest = $outcomeDigest
        status = $status
        result_artifacts = @($resultArtifacts | Sort-Object path)
        started_utc = $StartedUtc.ToUniversalTime().ToString('o')
        ended_utc = $EndedUtc.ToUniversalTime().ToString('o')
        duration_ms = $duration
    }
    $evidence.evidence_digest = Get-RSValidationRecordDigest -Record $evidence `
        -ExcludedProperty @('evidence_digest', 'started_utc', 'ended_utc', 'duration_ms')
    $entryRoot = Join-Path (Join-Path $Run.RunDirectory 'entries') $EntryContract.id
    $terminalRoot = Join-Path $entryRoot 'terminal-results'
    [void](New-Item -ItemType Directory -Path $terminalRoot -Force)
    $terminalPath = Join-Path $terminalRoot "$($TerminalResult.terminal_result_digest).json"
    Write-RSAtomicCanonicalJsonFile -LiteralPath $terminalPath -InputObject $TerminalResult `
        -SchemaPath (Join-Path $Run.SchemaRoot 'ValidationTerminalResult.schema.json') -AllowedRoot $Run.RunDirectory
    $evidenceRoot = Join-Path $entryRoot 'evidence'
    [void](New-Item -ItemType Directory -Path $evidenceRoot -Force)
    $evidencePath = Join-Path $evidenceRoot "$($evidence.evidence_digest).json"
    Write-RSAtomicCanonicalJsonFile -LiteralPath $evidencePath -InputObject $evidence `
        -SchemaPath (Join-Path $Run.SchemaRoot 'ValidationEntryEvidence.schema.json') -AllowedRoot $Run.RunDirectory
    Set-RSValidationEntryCheckpoint -Run $Run -EntryId $EntryContract.id -Status $status

    $promotion = $null
    $promotionPath = $null
    if ($status -eq 'provisional-passed') {
        $promotion = [ordered]@{
            schema = 'red-salamander.validation-promotion.v1'
            canonicalization = 'rs-canonical-json-v1'
            promotion_digest = '0' * 64
            origin_run_id = $Run.State.run_id
            entry_id = $EntryContract.id
            evidence_digest = $evidence.evidence_digest
            entry_fingerprint = $EntryFingerprint
            fingerprint_component_version = $EntryFingerprintComponentVersion
            fingerprint_components = $EntryFingerprintComponents
            plan_digest = $Run.State.plan_digest
            coverage_plan_digest = $Run.State.coverage_plan_digest
            terminal_result_digest = $TerminalResult.terminal_result_digest
            pre_snapshot_id = $PreSnapshot.snapshot_id
            post_snapshot_id = $PostSnapshot.snapshot_id
            build_receipt_id = $BuildReceipt.receipt_id
            artifact_epoch = [int64]$Run.State.artifact_epoch
            artifact_digests = $artifactDigests
            coverage_digest = $coverageDigest
            outcome_digest = $outcomeDigest
            result_artifacts = @($resultArtifacts | Sort-Object path)
            origin_terminalization_policy = 'entry-barrier-complete'
            promoted_utc = [datetime]::UtcNow.ToString('o')
        }
        $promotion.promotion_digest = Get-RSValidationRecordDigest -Record $promotion `
            -ExcludedProperty @('promotion_digest', 'promoted_utc')
        $promotionRoot = Join-Path $entryRoot 'promotions'
        [void](New-Item -ItemType Directory -Path $promotionRoot -Force)
        $promotionPath = Join-Path $promotionRoot "$($promotion.promotion_digest).json"
        Write-RSAtomicCanonicalJsonFile -LiteralPath $promotionPath -InputObject $promotion `
            -SchemaPath (Join-Path $Run.SchemaRoot 'ValidationPromotion.schema.json') -AllowedRoot $Run.RunDirectory
        Set-RSValidationEntryCheckpoint -Run $Run -EntryId $EntryContract.id -Status promoted `
            -PromotionDigest $promotion.promotion_digest
    }
    return [pscustomobject]@{
        Evidence = [pscustomobject]$evidence; EvidencePath = $evidencePath
        Promotion = if ($null -eq $promotion) { $null } else { [pscustomobject]$promotion }
        PromotionPath = $promotionPath; TerminalResult = $TerminalResult; TerminalResultPath = $terminalPath
        MissingRequiredArtifacts = $missingRequired
    }
}

function Complete-RSValidationEvidenceRun {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Run, [switch]$ForceFailed)
    $statuses = @($Run.State.entries | ForEach-Object { [string]$_.status })
    $Run.State.status = if ($ForceFailed -or @($statuses | Where-Object { $_ -eq 'failed' }).Count -gt 0) {
        'failed'
    } elseif ($Run.State.mutation_state -eq 'stable' -and
        @($statuses | Where-Object { $_ -notin @('promoted', 'reused') }).Count -eq 0) {
        'passed'
    } else {
        'interrupted'
    }
    Write-RSValidationRunState -Run $Run
    return $Run.State.status
}

function Read-RSVerifiedValidationPromotionEvidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OriginRunDirectory,
        [Parameter(Mandatory = $true)][string]$EntryId,
        [Parameter(Mandatory = $true)][string]$PromotionDigest,
        [Parameter(Mandatory = $true)][string]$SchemaRoot,
        [string]$ExpectedCoveragePlanDigest = '',
        [string]$ExpectedEntryFingerprint = '',
        [string]$ExpectedSnapshotId = '',
        [int64]$ExpectedArtifactEpoch = -1,
        [AllowNull()][object]$ExpectedBuildReceipt = $null,
        [switch]$RequireTerminalGreen
    )

    Assert-RSValidationIdentifier -Id $EntryId
    if ($PromotionDigest -notmatch '^[0-9a-f]{64}$') {
        throw "Validation promotion digest is invalid for '$EntryId'."
    }
    $suppliedOrigin = [IO.Path]::GetFullPath($OriginRunDirectory).TrimEnd('\')
    $evidenceRoot = Split-Path (Split-Path $suppliedOrigin -Parent) -Parent
    $resolvedOrigin = Resolve-RSValidationRunDirectory -EvidenceRoot $evidenceRoot `
        -RunReference $OriginRunDirectory -SchemaRoot $SchemaRoot
    $originRoot = $resolvedOrigin.RunDirectory
    $state = $resolvedOrigin.State
    if ($RequireTerminalGreen -and [string]$state.status -ne 'passed') {
        throw "Validation origin run is not terminal green: $($state.run_id)"
    }
    if ([string]$state.mutation_state -ne 'stable') {
        throw "Validation origin artifact epoch is not stable: $($state.mutation_state)"
    }

    $plan = Read-RSValidationJsonFile -LiteralPath (Join-Path $originRoot 'validation-plan.json') `
        -SchemaPath (Join-Path $SchemaRoot 'ValidationPlan.schema.json') -AllowedRoot $originRoot
    $decision = Read-RSValidationDecisionRecord -RunDirectory $originRoot `
        -DecisionDigest $state.decision_digest -SchemaRoot $SchemaRoot
    if ($plan.plan_digest -ne $state.plan_digest -or
        $plan.coverage_plan_digest -ne $state.coverage_plan_digest -or
        $plan.validation_mode -ne $state.validation_mode -or
        $plan.requested_suite -ne $state.requested_suite -or
        $decision.plan_digest -ne $state.plan_digest -or
        $decision.snapshot_id -ne $state.snapshot_id -or
        $decision.decision_digest -ne $state.decision_digest) {
        throw 'Validation origin plan, decision, and run-state identity do not agree.'
    }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedCoveragePlanDigest) -and
        $state.coverage_plan_digest -ne $ExpectedCoveragePlanDigest) {
        throw 'Validation origin coverage-plan identity does not match the current invocation.'
    }

    $stateEntry = @($state.entries | Where-Object id -eq $EntryId)
    $planEntry = @($plan.entries | Where-Object id -eq $EntryId)
    $decisionEntry = @($decision.entries | Where-Object id -eq $EntryId)
    if ($stateEntry.Count -ne 1 -or $planEntry.Count -ne 1 -or $decisionEntry.Count -ne 1) {
        throw "Validation origin does not contain one complete contract for '$EntryId'."
    }
    if ($stateEntry[0].status -ne 'promoted' -or
        $stateEntry[0].promotion_digest -ne $PromotionDigest -or
        $decisionEntry[0].expected_fingerprint -ne $stateEntry[0].expected_fingerprint) {
        throw "Validation origin entry '$EntryId' is not consistently promoted."
    }

    $promotionPath = Join-Path $originRoot "entries\$EntryId\promotions\$PromotionDigest.json"
    $promotion = Read-RSValidationJsonFile -LiteralPath $promotionPath `
        -SchemaPath (Join-Path $SchemaRoot 'ValidationPromotion.schema.json') -AllowedRoot $originRoot
    $computedPromotionDigest = Get-RSValidationRecordDigest -Record $promotion `
        -ExcludedProperty @('promotion_digest', 'promoted_utc')
    if ($promotion.promotion_digest -ne $PromotionDigest -or $computedPromotionDigest -ne $PromotionDigest) {
        throw "Validation promotion digest or filename mismatch for '$EntryId'."
    }

    $evidencePath = Join-Path $originRoot "entries\$EntryId\evidence\$($promotion.evidence_digest).json"
    $evidence = Read-RSValidationJsonFile -LiteralPath $evidencePath `
        -SchemaPath (Join-Path $SchemaRoot 'ValidationEntryEvidence.schema.json') -AllowedRoot $originRoot
    $computedEvidenceDigest = Get-RSValidationRecordDigest -Record $evidence `
        -ExcludedProperty @('evidence_digest', 'started_utc', 'ended_utc', 'duration_ms')
    if ($evidence.evidence_digest -ne $promotion.evidence_digest -or
        $computedEvidenceDigest -ne $promotion.evidence_digest) {
        throw "Validation evidence digest or filename mismatch for '$EntryId'."
    }

    if ($promotion.terminal_result_digest -ne $evidence.terminal_result_digest) {
        throw "Validation promotion and evidence terminal-result bindings disagree for '$EntryId'."
    }
    $terminalResultPath = Join-Path $originRoot "entries\$EntryId\terminal-results\$($evidence.terminal_result_digest).json"
    $terminalResult = Read-RSValidationJsonFile -LiteralPath $terminalResultPath `
        -SchemaPath (Join-Path $SchemaRoot 'ValidationTerminalResult.schema.json') -AllowedRoot $originRoot
    $computedTerminalResultDigest = Get-RSValidationRecordDigest -Record $terminalResult `
        -ExcludedProperty @('terminal_result_digest')
    if ($terminalResult.terminal_result_digest -ne $evidence.terminal_result_digest -or
        $computedTerminalResultDigest -ne $evidence.terminal_result_digest -or
        $terminalResult.entry_id -ne $EntryId -or
        $terminalResult.kind -ne $planEntry[0].kind -or
        $terminalResult.status -ne 'PASSED') {
        throw "Validation terminal result is missing, failed, or does not match '$EntryId'."
    }

    $expectedCoverageDigest = Get-RSCanonicalJsonDigest -InputObject $planEntry[0].coverage_contract
    $expectedOutcomeDigest = Get-RSCanonicalJsonDigest -InputObject $planEntry[0].outcome_contract
    $fingerprint = [string]$stateEntry[0].expected_fingerprint
    if (-not [string]::IsNullOrWhiteSpace($ExpectedEntryFingerprint) -and
        $fingerprint -ne $ExpectedEntryFingerprint) {
        throw "Validation origin fingerprint does not match the current entry '$EntryId'."
    }
    foreach ($record in @($promotion, $evidence)) {
        if ($record.entry_id -ne $EntryId -or
            $record.entry_fingerprint -ne $fingerprint -or
            $record.plan_digest -ne $state.plan_digest -or
            $record.coverage_plan_digest -ne $state.coverage_plan_digest -or
            $record.pre_snapshot_id -ne $state.snapshot_id -or
            $record.post_snapshot_id -ne $state.snapshot_id -or
            $record.build_receipt_id -ne $state.build_receipt_id -or
            [int64]$record.artifact_epoch -ne [int64]$state.artifact_epoch -or
            $record.coverage_digest -ne $expectedCoverageDigest -or
            $record.outcome_digest -ne $expectedOutcomeDigest) {
            throw "Validation promotion/evidence binding mismatch for '$EntryId'."
        }
    }
    if ($promotion.origin_run_id -ne $state.run_id -or
        $evidence.run_id -ne $state.run_id -or
        $evidence.status -ne 'provisional-passed' -or
        (Get-RSCanonicalJsonDigest -InputObject $promotion.artifact_digests) -ne
            (Get-RSCanonicalJsonDigest -InputObject $evidence.artifact_digests) -or
        $promotion.fingerprint_component_version -ne $evidence.fingerprint_component_version -or
        (Get-RSCanonicalJsonDigest -InputObject $promotion.fingerprint_components) -ne
            (Get-RSCanonicalJsonDigest -InputObject $evidence.fingerprint_components) -or
        (Get-RSCanonicalJsonDigest -InputObject @($promotion.result_artifacts | Sort-Object path)) -ne
            (Get-RSCanonicalJsonDigest -InputObject @($evidence.result_artifacts | Sort-Object path))) {
        throw "Validation promotion and evidence records disagree for '$EntryId'."
    }
    $verifiedArtifactBytes = [int64]0
    foreach ($artifact in @($evidence.result_artifacts)) {
        Assert-RSValidationPathSet -Path @([string]$artifact.path)
        $identity = Get-RSValidationArtifactIdentity -LiteralPath (Join-Path $originRoot ([string]$artifact.path))
        if ($identity.SizeBytes -ne [int64]$artifact.size_bytes -or $identity.Sha256 -cne [string]$artifact.sha256) {
            throw "Validation result artifact is missing or content-mismatched for '$EntryId': $($artifact.path)"
        }
        $verifiedArtifactBytes += $identity.SizeBytes
        if ($verifiedArtifactBytes -gt $script:MaximumValidationEntryArtifactBytes) {
            throw "Validation result artifacts for '$EntryId' exceed the 64 MiB entry quota."
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedSnapshotId) -and $state.snapshot_id -ne $ExpectedSnapshotId) {
        throw "Validation origin snapshot does not match the current entry '$EntryId'."
    }
    if ($ExpectedArtifactEpoch -ge 0 -and [int64]$state.artifact_epoch -ne $ExpectedArtifactEpoch) {
        throw "Validation origin artifact epoch does not match the current entry '$EntryId'."
    }
    if ($null -ne $ExpectedBuildReceipt) {
        if ($state.build_receipt_id -ne $ExpectedBuildReceipt.receipt_id) {
            throw "Validation origin build receipt does not match the current entry '$EntryId'."
        }
        $expectedArtifacts = [ordered]@{}
        foreach ($output in @($ExpectedBuildReceipt.outputs | Sort-Object id)) {
            $expectedArtifacts[[string]$output.id] = [string]$output.sha256
        }
        if ((Get-RSCanonicalJsonDigest -InputObject $expectedArtifacts) -ne
            (Get-RSCanonicalJsonDigest -InputObject $evidence.artifact_digests)) {
            throw "Validation origin artifact set is incomplete or mismatched for '$EntryId'."
        }
    }

    return [pscustomobject]@{
        State = $state; Plan = $plan; Decision = $decision
        StateEntry = $stateEntry[0]; PlanEntry = $planEntry[0]; DecisionEntry = $decisionEntry[0]
        Promotion = $promotion; Evidence = $evidence
        TerminalResult = $terminalResult
        ResultArtifactBytes = $verifiedArtifactBytes
        PromotionPath = $promotionPath; EvidencePath = $evidencePath; TerminalResultPath = $terminalResultPath
    }
}

function Get-RSValidationResumeCandidates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OriginRunDirectory,
        [Parameter(Mandatory = $true)][object]$CurrentPlan,
        [Parameter(Mandatory = $true)][object]$CurrentDecision,
        [Parameter(Mandatory = $true)][object]$CurrentSnapshot,
        [Parameter(Mandatory = $true)][object]$BuildReceipt,
        [hashtable]$CurrentFingerprintComponentsById = @{},
        [Parameter(Mandatory = $true)][string]$SchemaRoot
    )

    $suppliedOrigin = [IO.Path]::GetFullPath($OriginRunDirectory).TrimEnd('\')
    $evidenceRoot = Split-Path (Split-Path $suppliedOrigin -Parent) -Parent
    $resolvedOrigin = Resolve-RSValidationRunDirectory -EvidenceRoot $evidenceRoot `
        -RunReference $OriginRunDirectory -SchemaRoot $SchemaRoot
    $originRoot = $resolvedOrigin.RunDirectory
    $state = $resolvedOrigin.State
    $candidates = @()
    $verifiedRunArtifactBytes = [int64]0
    foreach ($entry in $CurrentPlan.entries) {
        $decisionEntry = @($CurrentDecision.entries | Where-Object id -eq $entry.id | Select-Object -First 1)
        $originEntry = @($state.entries | Where-Object id -eq $entry.id | Select-Object -First 1)
        $reason = 'NO_PRIOR_EVIDENCE'
        $promotion = $null
        $evidence = $null
        $reusable = $false
        if ($entry.cache_policy -ne 'ExactArtifact') {
            $reason = 'NON_CACHEABLE'
        } elseif ($originEntry.Count -eq 0 -or $originEntry[0].status -ne 'promoted' -or $null -eq $originEntry[0].promotion_digest) {
            $reason = 'PREVIOUS_NOT_GREEN'
        } elseif ($decisionEntry.Count -ne 1) {
            $reason = 'CORRUPT_EVIDENCE'
        } else {
            try {
                $verified = Read-RSVerifiedValidationPromotionEvidence `
                    -OriginRunDirectory $originRoot `
                    -EntryId $entry.id `
                    -PromotionDigest $originEntry[0].promotion_digest `
                    -SchemaRoot $SchemaRoot `
                    -RequireTerminalGreen
                if ($CurrentFingerprintComponentsById.ContainsKey([string]$entry.id)) {
                    $currentComponents = $CurrentFingerprintComponentsById[[string]$entry.id]
                    $originComponents = $verified.Evidence.fingerprint_components
                    $currentNames = @($currentComponents.digests.PSObject.Properties.Name | Sort-Object)
                    $originNames = @($originComponents.PSObject.Properties.Name | Sort-Object)
                    $knownNames = @($script:ValidationFingerprintComponentReasons.Keys | Sort-Object)
                    if ([string]$currentComponents.version -cne 'validation-entry-components.v1' -or
                        [string]$verified.Evidence.fingerprint_component_version -cne 'validation-entry-components.v1' -or
                        (Get-RSCanonicalJsonDigest $currentNames) -cne (Get-RSCanonicalJsonDigest $knownNames) -or
                        (Get-RSCanonicalJsonDigest $originNames) -cne (Get-RSCanonicalJsonDigest $knownNames)) {
                        $reason = 'SCHEMA_MISMATCH'
                        throw [InvalidOperationException]::new('COMPONENT_REASON_SELECTED')
                    }
                    foreach ($componentName in $script:ValidationFingerprintComponentReasons.Keys) {
                        if ([string]$currentComponents.digests.$componentName -cne [string]$originComponents.$componentName) {
                            $reason = $script:ValidationFingerprintComponentReasons[$componentName]
                            throw [InvalidOperationException]::new('COMPONENT_REASON_SELECTED')
                        }
                    }
                    if ([string]$originEntry[0].expected_fingerprint -cne [string]$decisionEntry[0].expected_fingerprint) {
                        $reason = 'CORRUPT_EVIDENCE'
                        throw [InvalidOperationException]::new('COMPONENT_REASON_SELECTED')
                    }
                } elseif ($originEntry[0].expected_fingerprint -ne $decisionEntry[0].expected_fingerprint) {
                    $reason = 'SOURCE_INPUT_CHANGED'
                    throw [InvalidOperationException]::new('COMPONENT_REASON_SELECTED')
                }
                $verified = Read-RSVerifiedValidationPromotionEvidence `
                    -OriginRunDirectory $originRoot `
                    -EntryId $entry.id `
                    -PromotionDigest $originEntry[0].promotion_digest `
                    -SchemaRoot $SchemaRoot `
                    -ExpectedCoveragePlanDigest $CurrentPlan.coverage_plan_digest `
                    -ExpectedEntryFingerprint $decisionEntry[0].expected_fingerprint `
                    -ExpectedSnapshotId $CurrentSnapshot.snapshot_id `
                    -ExpectedArtifactEpoch ([int64]$state.artifact_epoch) `
                    -ExpectedBuildReceipt $BuildReceipt `
                    -RequireTerminalGreen
                $promotion = $verified.Promotion
                $evidence = $verified.Evidence
                $verifiedRunArtifactBytes += [int64]$verified.ResultArtifactBytes
                if ($verifiedRunArtifactBytes -gt $script:MaximumValidationRunArtifactBytes) {
                    throw 'Verified validation result artifacts exceed the 1 GiB run quota.'
                }
                $reusable = $true
                $reason = 'REUSED_EXACT'
            } catch {
                if ($_.Exception.Message -eq 'COMPONENT_REASON_SELECTED') {
                    # The documented component-precedence branch selected it.
                } elseif ($_.Exception.Message -match 'snapshot|coverage-plan|fingerprint') {
                    $reason = 'SOURCE_INPUT_CHANGED'
                } elseif ($_.Exception.Message -match 'build receipt|artifact epoch|artifact set') {
                    $reason = 'BUILD_ARTIFACT_CHANGED'
                } else {
                    $reason = 'CORRUPT_EVIDENCE'
                }
                $reusable = $false
            }
        }
        $candidates += [pscustomobject][ordered]@{
            id = $entry.id; reusable = $reusable; reason_code = $reason
            origin_run_id = if ($reusable) { $state.run_id } else { $null }
            origin_evidence_digest = if ($reusable) { $evidence.evidence_digest } else { $null }
            origin_promotion_digest = if ($reusable) { $promotion.promotion_digest } else { $null }
        }
    }
    return $candidates
}

function New-RSValidationResumeDecision {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$FreshDecision, [Parameter(Mandatory = $true)][object[]]$Candidate)

    $byId = @{}
    foreach ($item in $Candidate) { $byId[[string]$item.id] = $item }
    $entries = foreach ($entry in $FreshDecision.entries) {
        $candidate = $byId[[string]$entry.id]
        [pscustomobject][ordered]@{
            id = $entry.id; ordinal = [int64]$entry.ordinal; affected = [bool]$entry.affected
            matched_impact_tags = @($entry.matched_impact_tags)
            decision = if ($candidate.reusable) { 'REUSE' } else { 'EXECUTE' }
            reason_codes = @([string]$candidate.reason_code)
            expected_fingerprint = $entry.expected_fingerprint
            origin_run_id = $candidate.origin_run_id
            origin_evidence_digest = $candidate.origin_evidence_digest
            origin_promotion_digest = $candidate.origin_promotion_digest
            estimated_duration_ms = if ($null -ne $entry.PSObject.Properties['estimated_duration_ms']) { [int64]$entry.estimated_duration_ms } else { [int64]0 }
        }
    }
    $executedDurationMeasure = $entries | Where-Object decision -eq EXECUTE | Measure-Object estimated_duration_ms -Sum
    $avoidedDurationMeasure = $entries | Where-Object decision -eq REUSE | Measure-Object estimated_duration_ms -Sum
    $executedDuration = if ($null -eq $executedDurationMeasure -or $null -eq $executedDurationMeasure.Sum) { 0L } else { [int64]$executedDurationMeasure.Sum }
    $avoidedDuration = if ($null -eq $avoidedDurationMeasure -or $null -eq $avoidedDurationMeasure.Sum) { 0L } else { [int64]$avoidedDurationMeasure.Sum }
    $record = [ordered]@{
        schema = $FreshDecision.schema; canonicalization = $FreshDecision.canonicalization
        decision_digest = '0' * 64; planner_mode = 'enforced'; snapshot_id = $FreshDecision.snapshot_id
        plan_digest = $FreshDecision.plan_digest; changed_paths = @($FreshDecision.changed_paths)
        matched_impact_rules = @($FreshDecision.matched_impact_rules); advisory_build_action = $FreshDecision.advisory_build_action
        entries = @($entries); estimated_fresh_duration_ms = [int64]$FreshDecision.estimated_fresh_duration_ms
        estimated_executed_duration_ms = $executedDuration
        estimated_avoided_duration_ms = $avoidedDuration
    }
    $record.decision_digest = Get-RSValidationRecordDigest -Record $record -ExcludedProperty @('decision_digest')
    return [pscustomobject]$record
}

function Repair-RSValidationInterruptedRunState {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Run)
    if ($Run.State.status -ne 'running') { return $false }
    foreach ($entry in $Run.State.entries) {
        if ($entry.status -in @('running', 'provisional-passed')) { $entry.status = 'interrupted' }
    }
    if ($Run.State.mutation_state -eq 'mutating') {
        $Run.State.mutation_state = 'interrupted'
    }
    $Run.State.status = 'interrupted'
    Write-RSValidationRunState -Run $Run
    return $true
}

function Start-RSValidationArtifactMutation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Run,
        [Parameter(Mandatory = $true)][object]$Plan,
        [Parameter(Mandatory = $true)][string[]]$WrittenArtifactId
    )

    if ($Run.State.status -ne 'running' -or $Run.State.mutation_state -ne 'stable') {
        throw "Validation artifact mutation requires one running stable epoch."
    }
    foreach ($artifactId in $WrittenArtifactId) { Assert-RSValidationIdentifier -Id $artifactId }
    $written = [Collections.Generic.HashSet[string]]::new($WrittenArtifactId, [StringComparer]::Ordinal)
    foreach ($planEntry in $Plan.entries) {
        $intersects = @($planEntry.artifact_read_ids | Where-Object { $written.Contains([string]$_) }).Count -gt 0
        if (-not $intersects) { continue }
        $stateEntry = @($Run.State.entries | Where-Object id -eq $planEntry.id)
        if ($stateEntry.Count -eq 1 -and $stateEntry[0].status -eq 'promoted') {
            $stateEntry[0].status = 'invalidated'
            $stateEntry[0].promotion_digest = $null
            $stateEntry[0].reason_codes = @('ARTIFACT_EPOCH_CHANGED')
        }
    }
    $Run.State.artifact_epoch = [int64]$Run.State.artifact_epoch + 1L
    $Run.State.mutation_state = 'mutating'
    Write-RSValidationRunState -Run $Run
    return [int64]$Run.State.artifact_epoch
}

function Complete-RSValidationArtifactMutation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Run,
        [Parameter(Mandatory = $true)][object]$Decision,
        [Parameter(Mandatory = $true)][object]$BuildReceipt,
        [Parameter(Mandatory = $true)][string]$MutatorEntryId,
        [Parameter(Mandatory = $true)][int64]$ExpectedEpoch
    )

    Assert-RSValidationIdentifier -Id $MutatorEntryId
    if ($Run.State.status -ne 'running' -or
        $Run.State.mutation_state -ne 'mutating' -or
        [int64]$Run.State.artifact_epoch -ne $ExpectedEpoch) {
        throw 'Validation artifact mutation completion does not match the active epoch.'
    }
    $decisionPath = Publish-RSValidationDecisionRecord -RunDirectory $Run.RunDirectory `
        -Decision $Decision -SchemaRoot $Run.SchemaRoot
    foreach ($decisionEntry in @($Decision.entries)) {
        $stateEntry = @($Run.State.entries | Where-Object id -eq $decisionEntry.id)
        if ($stateEntry.Count -ne 1) {
            throw "Artifact epoch decision lost entry '$($decisionEntry.id)'."
        }
        $stateEntry[0].decision = 'execute'
        $stateEntry[0].status = if ($stateEntry[0].id -eq $MutatorEntryId) { 'running' } else { 'pending' }
        $stateEntry[0].reason_codes = @($decisionEntry.reason_codes)
        $stateEntry[0].expected_fingerprint = $decisionEntry.expected_fingerprint
        $stateEntry[0].origin_run_id = $null
        $stateEntry[0].origin_evidence_digest = $null
        $stateEntry[0].promotion_digest = $null
    }
    $Run.State.snapshot_id = $Decision.snapshot_id
    $Run.State.build_receipt_id = $BuildReceipt.receipt_id
    $Run.State.decision_digest = $Decision.decision_digest
    $Run.State.mutation_state = 'stable'
    Write-RSValidationRunState -Run $Run
    $Run | Add-Member -NotePropertyName DecisionPath -NotePropertyValue $decisionPath -Force
    return [int64]$Run.State.artifact_epoch
}

function Test-RSValidationArtifactReadBarrier {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Run,
        [Parameter(Mandatory = $true)][object]$BuildReceipt,
        [Parameter(Mandatory = $true)][int64]$ExpectedEpoch
    )
    return $Run.State.mutation_state -eq 'stable' -and
        [int64]$Run.State.artifact_epoch -eq $ExpectedEpoch -and
        $Run.State.build_receipt_id -eq $BuildReceipt.receipt_id
}

function Get-RSValidationEvidencePrunePlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$EvidenceRoot,
        [ValidateRange(1, 10000)][int]$MaximumRuns = 20,
        [string[]]$KeepRunId = @()
    )

    $runsRoot = Get-RSValidationRunsRoot -EvidenceRoot $EvidenceRoot -AllowMissing
    if ($null -eq $runsRoot) { return @() }
    $keep = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($id in $KeepRunId) { Assert-RSValidationRunId $id; [void]$keep.Add($id) }
    $runs = @(Get-ChildItem -LiteralPath $runsRoot -Directory | Sort-Object Name -Descending)
    $schemaRoot = Join-Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) 'Specs\Testing'
    foreach ($run in $runs | Select-Object -First $MaximumRuns) { [void]$keep.Add($run.Name) }
    foreach ($run in $runs) {
        Assert-RSValidationRunId $run.Name
        if (($run.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Refusing reparse run: $($run.FullName)" }
        $statePath = Join-Path $run.FullName 'run-state.json'
        if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) { [void]$keep.Add($run.Name); continue }
        $state = Read-RSValidationJsonFile -LiteralPath $statePath `
            -SchemaPath (Join-Path (Split-Path $PSScriptRoot -Parent | Split-Path -Parent | Split-Path -Parent) 'Specs\Testing\ValidationRunState.schema.json') `
            -AllowedRoot $run.FullName
        if ($state.status -eq 'running') { [void]$keep.Add($run.Name) }
        foreach ($entry in $state.entries) {
            if ($null -ne $entry.origin_run_id) {
                $originRunId = [string]$entry.origin_run_id
                Assert-RSValidationRunId -RunId $originRunId
                if ($null -eq $entry.promotion_digest) {
                    throw "Validation prune provenance has no promotion digest for '$($entry.id)' in '$($state.run_id)'."
                }
                [void](Read-RSVerifiedValidationPromotionEvidence `
                        -OriginRunDirectory (Join-Path $runsRoot $originRunId) `
                        -EntryId ([string]$entry.id) `
                        -PromotionDigest ([string]$entry.promotion_digest) `
                        -SchemaRoot $schemaRoot `
                        -ExpectedCoveragePlanDigest ([string]$state.coverage_plan_digest) `
                        -RequireTerminalGreen)
                [void]$keep.Add($originRunId)
            }
        }
    }
    return @($runs | Where-Object { -not $keep.Contains($_.Name) } | ForEach-Object {
            [pscustomobject][ordered]@{
                LiteralPath = $_.FullName
                RunId = $_.Name
                DirectoryIdentity = Get-RSValidationDirectoryIdentity -LiteralPath $_.FullName
            }
        })
}

function Remove-RSValidationEvidenceRuns {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$EvidenceRoot,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$PrunePlan
    )

    $runsRoot = Get-RSValidationRunsRoot -EvidenceRoot $EvidenceRoot
    foreach ($entry in $PrunePlan) {
        $full = Resolve-RSValidationPrunePlanEntry -RunsRoot $runsRoot -Entry $entry
        if ($PSCmdlet.ShouldProcess($full, 'Remove validation evidence run')) {
            # Re-resolve the trusted chain and identity immediately before the
            # recursive operation so a plan/apply replacement cannot be deleted.
            $currentRunsRoot = Get-RSValidationRunsRoot -EvidenceRoot $EvidenceRoot
            if (-not [string]::Equals($currentRunsRoot, $runsRoot, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Validation evidence runs root changed after prune planning: $currentRunsRoot"
            }
            $full = Resolve-RSValidationPrunePlanEntry -RunsRoot $currentRunsRoot -Entry $entry
            Remove-Item -LiteralPath $full -Recurse -Force
        }
    }
}

function Get-RSValidationDurationEstimates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$EvidenceRoot,
        [Parameter(Mandatory = $true)][string]$SchemaRoot,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$EntryId,
        [ValidateRange(1, 200)][int]$MaximumRuns = 40,
        [ValidateRange(1, 20)][int]$MaximumSamplesPerEntry = 5
    )

    $result = @{}
    foreach ($id in $EntryId) { $result[$id] = [int64]0 }
    $root = Assert-RSValidationEvidenceRoot -EvidenceRoot $EvidenceRoot
    $runsRoot = Join-Path $root 'runs'
    if (-not (Test-Path -LiteralPath $runsRoot -PathType Container)) { return $result }

    $wanted = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($id in $EntryId) { [void]$wanted.Add($id) }
    $samples = @{}
    foreach ($run in @(Get-ChildItem -LiteralPath $runsRoot -Directory -Force | Sort-Object Name -Descending | Select-Object -First $MaximumRuns)) {
        if (($run.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
        try {
            $state = Read-RSValidationJsonFile -LiteralPath (Join-Path $run.FullName 'run-state.json') `
                -SchemaPath (Join-Path $SchemaRoot 'ValidationRunState.schema.json') -AllowedRoot $run.FullName
            foreach ($entry in @($state.entries | Where-Object { $_.status -eq 'promoted' -and $wanted.Contains([string]$_.id) })) {
                $id = [string]$entry.id
                if ($samples.ContainsKey($id) -and @($samples[$id]).Count -ge $MaximumSamplesPerEntry) { continue }
                $verified = Read-RSVerifiedValidationPromotionEvidence `
                    -OriginRunDirectory $run.FullName `
                    -EntryId $id `
                    -PromotionDigest $entry.promotion_digest `
                    -SchemaRoot $SchemaRoot `
                    -RequireTerminalGreen
                $evidence = $verified.Evidence
                if (-not $samples.ContainsKey($id)) { $samples[$id] = @() }
                $samples[$id] = @($samples[$id]) + [int64]$evidence.duration_ms
            }
        } catch {
            # History is advisory only. Invalid or incomplete records contribute no estimate
            # and can never alter fingerprints, affectedness, execution, or verdicts.
            continue
        }
    }
    foreach ($id in $EntryId) {
        $values = @($samples[$id] | Sort-Object)
        if ($values.Count -eq 0) { continue }
        $middle = [int][Math]::Floor($values.Count / 2)
        $result[$id] = if (($values.Count % 2) -eq 1) {
            [int64]$values[$middle]
        } else {
            [int64][Math]::Round(([double]$values[$middle - 1] + [double]$values[$middle]) / 2.0)
        }
    }
    return $result
}

Export-ModuleMember -Function @(
    'New-RSValidationEvidenceRun',
    'Resolve-RSValidationRunDirectory',
    'Publish-RSValidationResultArtifacts',
    'Publish-RSValidationDecisionRecord',
    'Read-RSValidationDecisionRecord',
    'Write-RSValidationRunState',
    'Set-RSValidationEntryCheckpoint',
    'New-RSValidationTerminalResult',
    'Publish-RSValidationEntryEvidence',
    'Complete-RSValidationEvidenceRun',
    'Read-RSVerifiedValidationPromotionEvidence',
    'Repair-RSValidationInterruptedRunState',
    'Start-RSValidationArtifactMutation',
    'Complete-RSValidationArtifactMutation',
    'Test-RSValidationArtifactReadBarrier',
    'Get-RSValidationResumeCandidates',
    'New-RSValidationResumeDecision',
    'Get-RSValidationDurationEstimates',
    'Get-RSValidationEvidencePrunePlan',
    'Remove-RSValidationEvidenceRuns'
)
