Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:HarnessVersion = '1.0.0'
$script:EvidenceJsonLeaf = 'terminal-engine-evidence.v1.json'
$script:EvidenceCompanionLeaf = 'terminal-engine-evidence.v1.sha256'
$script:CandidateSetLeaf = 'Specs\Terminal\TerminalEngineCandidateSet.v1.json'
$script:CandidateSetSchemaLeaf = 'Specs\Terminal\TerminalEngineCandidateSet.schema.json'
$script:SchemaLeaf = 'Specs\Terminal\TerminalEngineEvidence.schema.json'
$script:RequirementsLeaf = 'Specs\Terminal\TerminalEngineRequirements.v1.json'
$script:ScoreAnchorsLeaf = 'Specs\Terminal\TerminalEngineScoreAnchors.v1.json'
$script:FixtureCorpusLeaf = 'Specs\Terminal\TerminalEngineFixtureCorpus.v1.json'
$script:CandidateUniverseLeaf = 'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json'
$script:CandidateUniverseSchemaLeaf = (
    'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v4.schema.json')
$script:CandidateUniverseRound3Leaf = (
    'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json')
$script:CandidateUniverseRound3SchemaLeaf = (
    'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v3.schema.json')
$script:CandidateUniverseRound2Leaf = (
    'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json')
$script:CandidateUniverseRound2SchemaLeaf = (
    'Specs\Terminal\TerminalEngineCandidateUniverseTransition.schema.json')
$script:CandidateUniverseBaseLeaf = 'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json'
$script:CandidateUniverseBaseSchemaLeaf = (
    'Specs\Terminal\TerminalEngineCandidateUniverse.schema.json')
$script:CandidateReviewLeaf = 'Specs\Terminal\TerminalEngineCandidateReview.v1.json'
$script:CandidateReviewSchemaLeaf = 'Specs\Terminal\TerminalEngineCandidateReview.schema.json'
$script:CandidateAssessmentLeaf = 'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json'
$script:CandidateAssessmentSchemaLeaf = 'Specs\Terminal\TerminalEngineCandidateAssessment.schema.json'
$script:ConcernMappingSchemaLeaf = 'Specs\Terminal\TerminalEngineConcernMapping.schema.json'
$script:MeasurementProtocolLeaf = 'Specs\Terminal\TerminalEngineMeasurementProtocol.v1.json'
$script:CollectionProviderLeaf = 'Tools\TerminalEngine\Invoke-TerminalEngineCollectionProvider.ps1'
$script:HarnessFiles = @(
    'Directory.Build.props',
    'Directory.Build.targets',
    'RuntimeDependencies.props',
    'Tools\Evaluate-TerminalEngines.ps1',
    'Tools\TerminalEngineEvaluator.psm1',
    'Tools\TerminalEvidence.psm1',
    'Tools\TerminalEngine\Invoke-TerminalEngineCollectionProvider.ps1',
    'Tools\TerminalEngine\Invoke-RSIsolatedProcess.ps1',
    'Tools\TerminalEngine\TerminalEngineCollectionDescriptor.psm1',
    'Tools\TerminalEngine\TerminalEngineCollectionDescriptorContract.md',
    'Tools\TerminalEngine\TerminalEngineCollectionDriverContract.md',
    'Tools\TerminalEngine\TerminalEngineCollectionDriverInput.schema.json',
    'Tools\TerminalEngine\TerminalEngineCollectionDriverResult.schema.json',
    'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1',
    'Tools\TerminalEngine\Test-TerminalEngineCollectionNeutrality.ps1',
    'Tools\TerminalEngine\TerminalEnginePeIdentity.psm1',
    'Tools\TerminalEngine\TerminalEngineSourceReview.psm1',
    'Tools\TerminalEngine\TerminalEngineCandidateAssessment.psm1',
    'Tools\TerminalEngine\TerminalEngineC7Proof.psm1',
    'Tools\TerminalEngine\TerminalEngineC7ProofContract.md',
    'Tools\TerminalEngine\TerminalEngineGate0Adapter.h',
    'Tools\TerminalEngine\TerminalEngineGate0AdapterContract.md',
    'Tools\TerminalEngine\TerminalEngineGate0Harness.cpp',
    'Tools\TerminalEngine\TerminalEngineGate0Harness.vcxproj',
    'Tools\TerminalEngine\Invoke-TerminalEngineGate0Harness.ps1',
    'Tools\TerminalEngine\TerminalEngineGate0Layout.h',
    'Tools\TerminalEngine\TerminalEngineGate0Probe.cpp',
    'Tools\TerminalEngine\TerminalEngineGate0Probe.vcxproj',
    'Tools\TerminalEngine\Invoke-TerminalEngineGate0Probe.ps1',
    'Tools\TerminalEngine\TerminalEngineNotices.psm1',
    'Tools\TerminalEngine\Generate-TerminalEngineFixtures.ps1',
    'Tools\TerminalEngine\Generated\TerminalEngineFixtures.v1.h',
    'Specs\Terminal\TerminalEngineCandidateUniverse.schema.json',
    'Specs\Terminal\TerminalEngineCandidateUniverseTransition.schema.json',
    'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v3.schema.json',
    'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v4.schema.json',
    'Specs\Terminal\TerminalEngineCandidateSet.schema.json',
    'Specs\Terminal\TerminalEngineCandidateReview.schema.json',
    'Specs\Terminal\TerminalEngineCandidateAssessment.schema.json',
    'Specs\Terminal\TerminalEngineConcernMapping.schema.json',
    'Specs\Terminal\TerminalEngineCollectionDescriptor.schema.json',
    'Specs\Terminal\TerminalEngineCollectionNeutrality.schema.json',
    'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json',
    'Specs\Terminal\TerminalEngineC7Proof.schema.json',
    'Specs\Terminal\TerminalEngineC7ProofContract.v1.json',
    'Specs\Terminal\TerminalEngineEvidence.schema.json',
    'Tools\TerminalEngine\Test-TerminalEngineCandidatePatch.ps1',
    'vcpkg.json'
)
$script:Round3FrozenInputLeaves = [string[]]@(
    'Directory.Build.props',
    'Directory.Build.targets',
    'RuntimeDependencies.props',
    'Specs/Terminal/TerminalEngineC7Proof.schema.json',
    'Specs/Terminal/TerminalEngineC7ProofContract.v1.json',
    'Specs/Terminal/TerminalEngineCandidateReview.schema.json',
    'Specs/Terminal/TerminalEngineCandidateUniverse.v1.json',
    'Specs/Terminal/TerminalEngineCollectionNeutrality.v1.json',
    'Specs/Terminal/TerminalEngineEvidence.schema.json',
    'Specs/Terminal/TerminalEngineFixtureCorpus.v1.json',
    'Specs/Terminal/TerminalEngineMeasurementProtocol.v1.json',
    'Specs/Terminal/TerminalEngineRequirements.v1.json',
    'Specs/Terminal/TerminalEngineScoreAnchors.v1.json',
    'Tools/Evaluate-TerminalEngines.ps1',
    'Tools/New-TerminalEngineNotices.ps1',
    'Tools/TerminalEngine/Generate-TerminalEngineFixtures.ps1',
    'Tools/TerminalEngine/Generated/TerminalEngineFixtures.v1.h',
    'Tools/TerminalEngine/Invoke-RSIsolatedProcess.ps1',
    'Tools/TerminalEngine/Invoke-TerminalEngineCollectionProvider.ps1',
    'Tools/TerminalEngine/Invoke-TerminalEngineGate0Harness.ps1',
    'Tools/TerminalEngine/Invoke-TerminalEngineGate0Probe.ps1',
    'Tools/TerminalEngine/TerminalEngineC7Proof.psm1',
    'Tools/TerminalEngine/TerminalEngineC7ProofContract.md',
    'Tools/TerminalEngine/TerminalEngineCandidateAssessment.psm1',
    'Tools/TerminalEngine/TerminalEngineCollectionDescriptor.psm1',
    'Tools/TerminalEngine/TerminalEngineCollectionDescriptorContract.md',
    'Tools/TerminalEngine/TerminalEngineCollectionDriverContract.md',
    'Tools/TerminalEngine/TerminalEngineCollectionNeutrality.psm1',
    'Tools/TerminalEngine/TerminalEngineGate0Adapter.h',
    'Tools/TerminalEngine/TerminalEngineGate0AdapterContract.md',
    'Tools/TerminalEngine/TerminalEngineGate0Harness.cpp',
    'Tools/TerminalEngine/TerminalEngineGate0Harness.vcxproj',
    'Tools/TerminalEngine/TerminalEngineGate0Layout.h',
    'Tools/TerminalEngine/TerminalEngineGate0Probe.cpp',
    'Tools/TerminalEngine/TerminalEngineGate0Probe.vcxproj',
    'Tools/TerminalEngine/TerminalEngineNotices.psm1',
    'Tools/TerminalEngine/TerminalEnginePeIdentity.psm1',
    'Tools/TerminalEngine/TerminalEngineSourceReview.psm1',
    'Tools/TerminalEngine/Test-TerminalEngineCandidatePatch.ps1',
    'Tools/TerminalEngine/Test-TerminalEngineCollectionNeutrality.ps1',
    'Tools/TerminalEngineEvaluator.psm1',
    'Tools/TerminalEvidence.psm1',
    'Tools/Tests/TerminalEngineC7Proof.Tests.ps1',
    'Tools/Tests/TerminalEngineCandidateAssessment.Tests.ps1',
    'Tools/Tests/TerminalEngineCandidatePatch.Tests.ps1',
    'Tools/Tests/TerminalEngineCandidateUniverseTransition.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionDescriptor.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionDriver.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionNeutrality.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionProvider.Tests.ps1',
    'Tools/Tests/TerminalEngineEvaluator.Tests.ps1',
    'Tools/Tests/TerminalEngineFixtureGenerator.Tests.ps1',
    'Tools/Tests/TerminalEngineGate0Harness.Tests.ps1',
    'Tools/Tests/TerminalEngineGate0Probe.Tests.ps1',
    'Tools/Tests/TerminalEngineNetworkIsolation.Tests.ps1',
    'Tools/Tests/TerminalEngineNotices.Tests.ps1',
    'Tools/Tests/TerminalEnginePeIdentity.Tests.ps1',
    'Tools/Tests/TerminalEngineSourceReview.Tests.ps1',
    'Tools/Tests/TerminalEvidence.Tests.ps1',
    'vcpkg.json'
)
$script:Round4FrozenInputLeaves = [string[]]@(
    $script:Round3FrozenInputLeaves +
    @(
        'Specs/Terminal/TerminalEngineCandidateAssessment.schema.json',
        'Specs/Terminal/TerminalEngineCandidateSet.schema.json',
        'Specs/Terminal/TerminalEngineCandidateUniverse.v3.json',
        'Specs/Terminal/TerminalEngineCandidateUniverseTransition.v3.schema.json',
        'Specs/Terminal/TerminalEngineCandidateUniverseTransition.v4.schema.json'
    )
)
[Array]::Sort($script:Round4FrozenInputLeaves, [StringComparer]::Ordinal)
$script:CollectionNeutralitySurfaceLeaves = [string[]]@(
    '.github/workflows/terminal-engine-gate0.yml',
    'Directory.Build.props',
    'Directory.Build.targets',
    'RuntimeDependencies.props',
    'Specs/Terminal/TerminalEngineC7Proof.schema.json',
    'Specs/Terminal/TerminalEngineC7ProofContract.v1.json',
    'Specs/Terminal/TerminalEngineCollectionDescriptor.schema.json',
    'Tools/Evaluate-TerminalEngines.ps1',
    'Tools/TerminalEngine/Generate-TerminalEngineFixtures.ps1',
    'Tools/TerminalEngine/Generated/TerminalEngineFixtures.v1.h',
    'Tools/TerminalEngine/Invoke-RSIsolatedProcess.ps1',
    'Tools/TerminalEngine/Invoke-TerminalEngineCollectionProvider.ps1',
    'Tools/TerminalEngine/Invoke-TerminalEngineGate0Harness.ps1',
    'Tools/TerminalEngine/Invoke-TerminalEngineGate0Probe.ps1',
    'Tools/TerminalEngine/TerminalEngineArchive.psm1',
    'Tools/TerminalEngine/TerminalEngineC7Proof.psm1',
    'Tools/TerminalEngine/TerminalEngineC7ProofContract.md',
    'Tools/TerminalEngine/TerminalEngineCandidateAssessment.psm1',
    'Tools/TerminalEngine/TerminalEngineCollectionDescriptor.psm1',
    'Tools/TerminalEngine/TerminalEngineCollectionDescriptorContract.md',
    'Tools/TerminalEngine/TerminalEngineCollectionDriverContract.md',
    'Tools/TerminalEngine/TerminalEngineCollectionDriverInput.schema.json',
    'Tools/TerminalEngine/TerminalEngineCollectionDriverResult.schema.json',
    'Tools/TerminalEngine/TerminalEngineCollectionNeutrality.psm1',
    'Tools/TerminalEngine/TerminalEngineGate0Adapter.h',
    'Tools/TerminalEngine/TerminalEngineGate0AdapterContract.md',
    'Tools/TerminalEngine/TerminalEngineGate0ContractTestAdapter.cpp',
    'Tools/TerminalEngine/TerminalEngineGate0ContractTestAdapter.vcxproj',
    'Tools/TerminalEngine/TerminalEngineGate0Harness.cpp',
    'Tools/TerminalEngine/TerminalEngineGate0Harness.vcxproj',
    'Tools/TerminalEngine/TerminalEngineGate0Layout.h',
    'Tools/TerminalEngine/TerminalEngineGate0Probe.cpp',
    'Tools/TerminalEngine/TerminalEngineGate0Probe.vcxproj',
    'Tools/TerminalEngine/TerminalEngineLifecycle.psm1',
    'Tools/TerminalEngine/TerminalEngineNotices.psm1',
    'Tools/TerminalEngine/TerminalEnginePeIdentity.psm1',
    'Tools/TerminalEngine/TerminalEngineSourceReview.psm1',
    'Tools/TerminalEngine/Test-TerminalEngineCandidatePatch.ps1',
    'Tools/TerminalEngine/Test-TerminalEngineCollectionNeutrality.ps1',
    'Tools/TerminalEngineEvaluator.psm1',
    'Tools/TerminalEvidence.psm1',
    'Tools/Tests/TerminalEngineArchive.Tests.ps1',
    'Tools/Tests/TerminalEngineC7Proof.Tests.ps1',
    'Tools/Tests/TerminalEngineCandidateAssessment.Tests.ps1',
    'Tools/Tests/TerminalEngineCandidatePatch.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionDescriptor.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionDriver.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionNeutrality.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionProvider.Tests.ps1',
    'Tools/Tests/TerminalEngineNotices.Tests.ps1',
    'vcpkg.json'
)
[Array]::Sort(
    $script:CollectionNeutralitySurfaceLeaves,
    [StringComparer]::Ordinal)
$script:GateIds = [string[]](1..12 | ForEach-Object { "G$_" })
$script:CriterionIds = [string[]](1..11 | ForEach-Object { "C$_" })
$script:RequiredLaneKeys = [string[]]@(
    'ARM64|Release',
    'x64|ASan Debug',
    'x64|Release'
)
$script:TieMetricIds = [ordered]@{
    ownerDays = 'tie.owner-days'
    crossedAbiElements = 'tie.crossed-abi-elements'
    nonSystemRuntimeDependencies = 'tie.non-system-runtime-dependencies'
    addedReleasePackageBytes = 'tie.added-release-package-bytes'
}

function Get-RSPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [switch]$Required
    )

    $found = $false
    $value = $null
    if ($null -ne $InputObject -and $InputObject -is [Collections.IDictionary]) {
        foreach ($entry in $InputObject.GetEnumerator()) {
            if ($entry.Key -is [string] -and
                [string]::Equals([string]$entry.Key, $Name, [StringComparison]::Ordinal)) {
                $found = $true
                $value = $entry.Value
                break
            }
        }
    } elseif ($null -ne $InputObject) {
        foreach ($property in $InputObject.PSObject.Properties) {
            if ([string]::Equals($property.Name, $Name, [StringComparison]::Ordinal)) {
                $found = $true
                $value = $property.Value
                break
            }
        }
    }

    if ($Required -and -not $found) {
        throw [IO.InvalidDataException]::new("Required property '$Name' is missing.")
    }
    return [pscustomobject]@{
        Found = $found
        Value = $value
    }
}

function Get-RSRequiredString {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [string]$Pattern = '',

        [string]$Context = 'object'
    )

    $property = Get-RSPropertyValue -InputObject $InputObject -Name $Name -Required
    if ($property.Value -isnot [string] -or [string]::IsNullOrEmpty([string]$property.Value)) {
        throw [IO.InvalidDataException]::new("$Context property '$Name' must be a nonempty string.")
    }
    $value = [string]$property.Value
    if (-not [string]::IsNullOrEmpty($Pattern) -and $value -cnotmatch $Pattern) {
        throw [IO.InvalidDataException]::new("$Context property '$Name' has an invalid value.")
    }
    return $value
}

function Get-RSRequiredBoolean {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [string]$Context = 'object'
    )

    $property = Get-RSPropertyValue -InputObject $InputObject -Name $Name -Required
    if ($property.Value -isnot [bool]) {
        throw [IO.InvalidDataException]::new("$Context property '$Name' must be Boolean.")
    }
    return [bool]$property.Value
}

function Get-RSRequiredArray {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [string]$Context = 'object'
    )

    $property = Get-RSPropertyValue -InputObject $InputObject -Name $Name -Required
    if ($null -eq $property.Value -or $property.Value -is [string] -or
        $property.Value -isnot [Collections.IEnumerable]) {
        throw [IO.InvalidDataException]::new("$Context property '$Name' must be an array.")
    }
    return ,([object[]]@($property.Value))
}

function Get-RSRequiredUtcSecond {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [string]$Context = 'object'
    )

    $value = Get-RSRequiredString `
        -InputObject $InputObject `
        -Name $Name `
        -Context $Context
    if ($value.Length -ne 20 -or
        $value -cnotmatch (
            '^[0-9]{4}-(?:0[1-9]|1[0-2])-' +
            '(?:0[1-9]|[12][0-9]|3[01])T' +
            '(?:[01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]Z$')) {
        throw [IO.InvalidDataException]::new(
            "$Context property '$Name' must be an exact RFC 3339 UTC-second string.")
    }
    $parsed = [DateTimeOffset]::MinValue
    $styles = [Globalization.DateTimeStyles]::AssumeUniversal -bor
        [Globalization.DateTimeStyles]::AdjustToUniversal
    if (-not [DateTimeOffset]::TryParseExact(
            $value,
            "yyyy-MM-dd'T'HH:mm:ss'Z'",
            [Globalization.CultureInfo]::InvariantCulture,
            $styles,
            [ref]$parsed) -or
        $parsed.ToString(
            "yyyy-MM-dd'T'HH:mm:ss'Z'",
            [Globalization.CultureInfo]::InvariantCulture) -cne $value) {
        throw [IO.InvalidDataException]::new(
            "$Context property '$Name' must be an exact RFC 3339 UTC-second string.")
    }
    return $parsed
}

function Get-RSUniverseNominationWindowOpenUtc {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Universe,

        [string]$Context = 'candidate universe'
    )

    $universeFrozenUtc = Get-RSRequiredUtcSecond `
        -InputObject $Universe `
        -Name frozenUtc `
        -Context $Context
    $approvals = Get-RSRequiredArray `
        -InputObject $Universe `
        -Name approvals `
        -Context $Context
    if ($approvals.Count -eq 0) {
        throw [IO.InvalidDataException]::new(
            "$Context has no approval that can open the nomination window.")
    }
    $reviewerIds =
        [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $windowOpenUtc = [DateTimeOffset]::MinValue
    foreach ($approval in $approvals) {
        $reviewerId = Get-RSRequiredString `
            -InputObject $approval `
            -Name reviewerId `
            -Context "$Context approval"
        if (-not $reviewerIds.Add($reviewerId)) {
            throw [IO.InvalidDataException]::new(
                "$Context reviewer '$reviewerId' approved more than once.")
        }
        $reviewedUtc = Get-RSRequiredUtcSecond `
            -InputObject $approval `
            -Name reviewedUtc `
            -Context "$Context approval '$reviewerId'"
        if ($reviewedUtc -lt $universeFrozenUtc) {
            throw [IO.InvalidDataException]::new(
                "$Context approval '$reviewerId' predates the frozen proposal.")
        }
        if ($reviewedUtc -gt $windowOpenUtc) {
            $windowOpenUtc = $reviewedUtc
        }
    }
    return $windowOpenUtc
}

function Get-RSGovernanceFreezeUtc {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Record,

        [Parameter(Mandatory = $true)]
        [DateTimeOffset]$WindowOpenUtc,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $frozenUtc = Get-RSRequiredUtcSecond `
        -InputObject $Record `
        -Name frozenUtc `
        -Context $Context
    if ($frozenUtc -le $WindowOpenUtc) {
        throw [IO.InvalidDataException]::new(
            "$Context must freeze strictly after the nomination window opens.")
    }
    return $frozenUtc
}

function Assert-RSApprovalInsideGovernanceWindow {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Approval,

        [Parameter(Mandatory = $true)]
        [DateTimeOffset]$WindowOpenUtc,

        [Parameter(Mandatory = $true)]
        [DateTimeOffset]$GovernanceFreezeUtc,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    foreach ($item in $Approval) {
        $approvalId = Get-RSRequiredString `
            -InputObject $item `
            -Name approvalId `
            -Context $Context
        $reviewedUtc = Get-RSRequiredUtcSecond `
            -InputObject $item `
            -Name reviewedUtc `
            -Context "$Context '$approvalId'"
        if ($reviewedUtc -le $WindowOpenUtc) {
            throw [IO.InvalidDataException]::new(
                "$Context '$approvalId' is not strictly after the nomination window opened.")
        }
        if ($reviewedUtc -ge $GovernanceFreezeUtc) {
            throw [IO.InvalidDataException]::new(
                "$Context '$approvalId' is not strictly before the common governance freeze.")
        }
    }
}

function Get-RSFileSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not [IO.File]::Exists($Path)) {
        throw [IO.FileNotFoundException]::new("Required frozen input is missing: $Path", $Path)
    }
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $digest = $sha256.ComputeHash($stream)
    } finally {
        $sha256.Dispose()
        $stream.Dispose()
    }
    return ([BitConverter]::ToString($digest) -replace '-', '').ToLowerInvariant()
}

function Initialize-RSReadSnapshotNative {
    $typeName = 'RedSalamander.TerminalEngine.ReadSnapshotNative'
    if ($null -ne ($typeName -as [type])) {
        return
    }

    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace RedSalamander.TerminalEngine
{
    public static class ReadSnapshotNative
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct FileAttributeTagInfo
        {
            internal uint FileAttributes;
            internal uint ReparseTag;
        }

        [DllImport(
            "kernel32.dll",
            CharSet = CharSet.Unicode,
            ExactSpelling = true,
            SetLastError = true)]
        public static extern SafeFileHandle CreateFileW(
            string fileName,
            uint desiredAccess,
            uint shareMode,
            IntPtr securityAttributes,
            uint creationDisposition,
            uint flagsAndAttributes,
            IntPtr templateFile);

        [DllImport(
            "kernel32.dll",
            ExactSpelling = true,
            SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandleEx(
            SafeFileHandle file,
            int fileInformationClass,
            out FileAttributeTagInfo fileInformation,
            uint bufferSize);

        public static uint GetAttributesByHandle(SafeFileHandle handle)
        {
            FileAttributeTagInfo info;
            const int FileAttributeTagInfoClass = 9;
            if (!GetFileInformationByHandleEx(
                    handle,
                    FileAttributeTagInfoClass,
                    out info,
                    (uint)Marshal.SizeOf<FileAttributeTagInfo>()))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            return info.FileAttributes;
        }
    }
}
'@
}

function Open-RSReadSnapshotHandle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$Directory
    )

    Initialize-RSReadSnapshotNative
    $nativeType =
        'RedSalamander.TerminalEngine.ReadSnapshotNative' -as [type]
    $genericRead = [uint32]2147483648
    $fileShareRead = [uint32]0x00000001
    $openExisting = [uint32]3
    $openReparsePoint = [uint32]0x00200000
    $backupSemantics = [uint32]0x02000000
    $flags = if ($Directory) {
        $openReparsePoint -bor $backupSemantics
    } else {
        $openReparsePoint
    }
    $nativePath = if ($Path.StartsWith(
            '\\',
            [StringComparison]::Ordinal)) {
        '\\?\UNC\' + $Path.Substring(2)
    } else {
        '\\?\' + $Path
    }
    $handle = $nativeType::CreateFileW(
        $nativePath,
        $genericRead,
        $fileShareRead,
        [IntPtr]::Zero,
        $openExisting,
        $flags,
        [IntPtr]::Zero)
    if ($null -eq $handle -or $handle.IsInvalid) {
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        if ($null -ne $handle) {
            $handle.Dispose()
        }
        throw [ComponentModel.Win32Exception]::new(
            $lastError,
            "Cannot acquire immutable read snapshot for '$Path'.")
    }
    return ,$handle
}

function New-RSReadSnapshotLockSet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Path
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).
        TrimEnd([IO.Path]::DirectorySeparatorChar)
    $lockSet = [pscustomobject]@{
        RepoRoot = $root
        Handles = [Collections.Generic.List[IDisposable]]::new()
        DirectoryPath = [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
        FileStreamByPath =
            [Collections.Generic.Dictionary[string, IO.FileStream]]::new(
                [StringComparer]::OrdinalIgnoreCase)
    }
    try {
        Add-RSReadSnapshotLocks -LockSet $lockSet -Path $Path
        return $lockSet
    } catch {
        Close-RSReadSnapshotLockSet -LockSet $lockSet
        throw
    }
}

function Add-RSReadSnapshotLocks {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$LockSet,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Path
    )

    $root = [string]$LockSet.RepoRoot
    $rootPrefix = $root + [IO.Path]::DirectorySeparatorChar
    $files =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    $directories =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($inputPath in $Path) {
        if ([string]::IsNullOrWhiteSpace($inputPath)) {
            throw [IO.InvalidDataException]::new(
                'Read-snapshot paths must be nonempty.')
        }
        $absolute = if ([IO.Path]::IsPathRooted($inputPath)) {
            [IO.Path]::GetFullPath($inputPath)
        } else {
            [IO.Path]::GetFullPath((
                    Join-Path $root ($inputPath.Replace('/', '\'))))
        }
        if (-not $absolute.StartsWith(
                $rootPrefix,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw [IO.InvalidDataException]::new(
                "Read-snapshot path escapes the repository: '$inputPath'.")
        }
        [void]$files.Add($absolute)
        $directory = [IO.Path]::GetDirectoryName($absolute)
        while ($directory.StartsWith(
                $rootPrefix,
                [StringComparison]::OrdinalIgnoreCase)) {
            [void]$directories.Add($directory)
            $directory = [IO.Path]::GetDirectoryName($directory)
        }
        [void]$directories.Add($root)
    }

    $orderedDirectories = [string[]]@($directories)
    [Array]::Sort($orderedDirectories, [StringComparer]::OrdinalIgnoreCase)
    foreach ($directory in $orderedDirectories) {
        if ($LockSet.DirectoryPath.Contains($directory)) {
            continue
        }
        $attributes = [IO.File]::GetAttributes($directory)
        if (($attributes -band [IO.FileAttributes]::Directory) -eq 0 -or
            ($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw [IO.InvalidDataException]::new(
                "Read-snapshot directory is missing or is a reparse point: '$directory'.")
        }
        $handle = Open-RSReadSnapshotHandle -Path $directory -Directory
        $LockSet.Handles.Add($handle)
        $handleAttributes = (
            'RedSalamander.TerminalEngine.ReadSnapshotNative' -as [type]
        )::GetAttributesByHandle($handle)
        if (($handleAttributes -band [uint32][IO.FileAttributes]::Directory) -eq 0 -or
            ($handleAttributes -band [uint32][IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw [IO.InvalidDataException]::new(
                "Held read-snapshot directory is a reparse point: '$directory'.")
        }
        $attributes = [IO.File]::GetAttributes($directory)
        if (($attributes -band [IO.FileAttributes]::Directory) -eq 0 -or
            ($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw [IO.InvalidDataException]::new(
                "Read-snapshot directory changed or is a reparse point: '$directory'.")
        }
        [void]$LockSet.DirectoryPath.Add($directory)
    }

    $orderedFiles = [string[]]@($files)
    [Array]::Sort($orderedFiles, [StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $orderedFiles) {
        if ($LockSet.FileStreamByPath.ContainsKey($file)) {
            continue
        }
        $attributes = [IO.File]::GetAttributes($file)
        if (($attributes -band [IO.FileAttributes]::Directory) -ne 0 -or
            ($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw [IO.InvalidDataException]::new(
                "Read-snapshot file is missing or is a reparse point: '$file'.")
        }
        $handle = Open-RSReadSnapshotHandle -Path $file
        $handleAttributes = (
            'RedSalamander.TerminalEngine.ReadSnapshotNative' -as [type]
        )::GetAttributesByHandle($handle)
        if (($handleAttributes -band [uint32][IO.FileAttributes]::Directory) -ne 0 -or
            ($handleAttributes -band [uint32][IO.FileAttributes]::ReparsePoint) -ne 0) {
            $handle.Dispose()
            throw [IO.InvalidDataException]::new(
                "Held read-snapshot file is a reparse point: '$file'.")
        }
        $stream = [IO.FileStream]::new($handle, [IO.FileAccess]::Read)
        $LockSet.Handles.Add($stream)
        $attributes = [IO.File]::GetAttributes($file)
        if (($attributes -band [IO.FileAttributes]::Directory) -ne 0 -or
            ($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw [IO.InvalidDataException]::new(
                "Read-snapshot file changed or is a reparse point: '$file'.")
        }
        $LockSet.FileStreamByPath.Add($file, $stream)
    }
}

function Close-RSReadSnapshotLockSet {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$LockSet
    )

    for ($index = $LockSet.Handles.Count - 1; $index -ge 0; $index--) {
        $LockSet.Handles[$index].Dispose()
    }
    $LockSet.Handles.Clear()
    $LockSet.FileStreamByPath.Clear()
    $LockSet.DirectoryPath.Clear()
}

function Read-RSReadSnapshotBytes {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$LockSet,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int64]$MaximumLength = 16777216
    )

    $absolute = [IO.Path]::GetFullPath($Path)
    $stream = $null
    if (-not $LockSet.FileStreamByPath.TryGetValue(
            $absolute,
            [ref]$stream)) {
        throw [IO.InvalidDataException]::new(
            "Path is not part of the immutable read snapshot: '$Path'.")
    }
    $length = $stream.Length
    if ($length -lt 0 -or $length -gt $MaximumLength -or
        $length -gt [int]::MaxValue) {
        throw [IO.InvalidDataException]::new(
            "Read-snapshot file violates the fixed byte bound: '$Path'.")
    }
    $bytes = [byte[]]::new([int]$length)
    $stream.Position = 0
    $offset = 0
    while ($offset -lt $bytes.Length) {
        $read = $stream.Read($bytes, $offset, $bytes.Length - $offset)
        if ($read -eq 0) {
            throw [IO.EndOfStreamException]::new(
                "Read-snapshot file ended before its held length: '$Path'.")
        }
        $offset += $read
    }
    if ($stream.Length -ne $length) {
        throw [IO.InvalidDataException]::new(
            "Read-snapshot file length changed while reading: '$Path'.")
    }
    return ,$bytes
}

function Read-RSImmutableRepoFileBytes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int64]$MaximumLength = 16777216
    )

    $snapshot = New-RSReadSnapshotLockSet `
        -RepoRoot $RepoRoot `
        -Path @($Path)
    try {
        return ,(Read-RSReadSnapshotBytes `
                -LockSet $snapshot `
                -Path $Path `
                -MaximumLength $MaximumLength)
    } finally {
        Close-RSReadSnapshotLockSet -LockSet $snapshot
    }
}

function Import-RSExactModule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $expectedPath = [IO.Path]::GetFullPath($Path)
    $globalExecutionContext = (
        Get-Variable -Name ExecutionContext -Scope Global -ValueOnly)
    $hadGlobalExports = [bool](
        $globalExecutionContext.SessionState.InvokeCommand.GetCommands(
            '*',
            [Management.Automation.CommandTypes]::All,
            $true) |
            Where-Object {
                $null -ne $_.Module -and
                -not [string]::IsNullOrEmpty([string]$_.Module.Path) -and
                [string]::Equals(
                    [IO.Path]::GetFullPath([string]$_.Module.Path),
                    $expectedPath,
                    [StringComparison]::OrdinalIgnoreCase)
            } |
            Select-Object -First 1)
    $imported = [object[]]@(
        Import-Module `
            -Name $expectedPath `
            -Force `
            -PassThru)
    $matching = [object[]]@(
        $imported | Where-Object {
            -not [string]::IsNullOrEmpty([string]$_.Path) -and
            [string]::Equals(
                [IO.Path]::GetFullPath([string]$_.Path),
                $expectedPath,
                [StringComparison]::OrdinalIgnoreCase)
        })
    if ($matching.Count -ne 1) {
        throw [IO.InvalidDataException]::new(
            "The exact frozen module was not loaded: '$expectedPath'.")
    }
    $exactModule = $matching[0]
    if ($hadGlobalExports) {
        # -Force removes same-path exports from every visible scope. Re-expose
        # only the newly authenticated module when that path was already part
        # of the caller's global command surface.
        [void](Import-Module -ModuleInfo $exactModule -Global)
    }
    return $exactModule
}

function Get-RSRound4SnapshotPaths {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentUniversePath,

        [switch]$IncludeLiveGovernance
    )

    $result = [Collections.Generic.List[string]]::new()
    foreach ($path in $script:Round4FrozenInputLeaves) {
        $result.Add($path)
    }
    foreach ($path in $script:CollectionNeutralitySurfaceLeaves) {
        $result.Add($path)
    }
    foreach ($path in @(
            $CurrentUniversePath,
            $script:CandidateUniverseBaseLeaf,
            $script:CandidateUniverseRound2Leaf,
            $script:CandidateUniverseRound3Leaf,
            $script:CandidateUniverseBaseSchemaLeaf,
            $script:CandidateUniverseRound2SchemaLeaf,
            $script:CandidateUniverseRound3SchemaLeaf,
            $script:CandidateUniverseSchemaLeaf,
            $script:CandidateAssessmentSchemaLeaf,
            $script:CandidateReviewSchemaLeaf,
            $script:CandidateSetSchemaLeaf,
            $script:ConcernMappingSchemaLeaf,
            'Specs\Terminal\TerminalEngineCollectionDescriptor.schema.json',
            'Specs\Terminal\TerminalEngineCollectionNeutrality.schema.json',
            'Tools\TerminalEngine\TerminalEngineCollectionDriverInput.schema.json',
            'Tools\TerminalEngine\TerminalEngineCollectionDriverResult.schema.json')) {
        $result.Add($path)
    }
    if ($IncludeLiveGovernance) {
        foreach ($path in @(
                $script:CandidateSetLeaf,
                $script:CandidateReviewLeaf,
                $script:CandidateAssessmentLeaf)) {
            $result.Add($path)
        }
    }
    return [string[]]@($result)
}

function Import-RSAuthenticatedCollectionNeutralityModules {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    [void](Import-RSExactModule -Path (Join-Path $RepoRoot (
                'Tools\TerminalEvidence.psm1')))
    $neutralityModule = Import-RSExactModule -Path (Join-Path $RepoRoot (
            'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1'))
    $surface = [string[]]@(
        & $neutralityModule {
            Get-RSTerminalEngineCollectionNeutralitySurface
        })
    if (-not [Linq.Enumerable]::SequenceEqual(
            $surface,
            $script:CollectionNeutralitySurfaceLeaves)) {
        throw [IO.InvalidDataException]::new(
            'Authenticated collection-neutrality module surface differs from the trusted evaluator closure.')
    }
    return $neutralityModule
}

function ConvertFrom-RSReadSnapshotJson {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$LockSet,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $bytes = Read-RSReadSnapshotBytes -LockSet $LockSet -Path $Path
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        return $json |
            ConvertFrom-Json `
                -AsHashtable `
                -Depth 100 `
                -NoEnumerate `
                -DateKind String
    }
    catch [System.Text.DecoderFallbackException] {
        throw [IO.InvalidDataException]::new(
            "$Context is not strict UTF-8.",
            $_.Exception)
    }
    catch [System.Management.Automation.PSInvalidOperationException] {
        throw [IO.InvalidDataException]::new(
            "$Context is not duplicate-free JSON.",
            $_.Exception)
    }
}

function Get-RSGovernanceReferencedSnapshotPaths {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$LockSet,

        [Parameter(Mandatory = $true)]
        [string]$CandidateSetPath,

        [Parameter(Mandatory = $true)]
        [string]$CandidateAssessmentPath
    )

    $paths =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
    $candidateSet = ConvertFrom-RSReadSnapshotJson `
        -LockSet $LockSet `
        -Path $CandidateSetPath `
        -Context 'CandidateSet snapshot'
    $outcomes = Get-RSRequiredArray `
        -InputObject $candidateSet `
        -Name nominationOutcomes `
        -Context 'CandidateSet snapshot'
    if ($outcomes.Count -gt 64) {
        throw [IO.InvalidDataException]::new(
            'CandidateSet snapshot exceeds the pre-schema nomination bound.')
    }
    foreach ($outcome in $outcomes) {
        if ((Get-RSRequiredString `
                    -InputObject $outcome `
                    -Name outcome `
                    -Context 'CandidateSet snapshot nomination outcome') -cne
            'adaptation-rejected') {
            continue
        }
        $rejection = (Get-RSPropertyValue `
                -InputObject $outcome `
                -Name adaptationRejection `
                -Required).Value
        [void]$paths.Add((Get-RSRequiredString `
                    -InputObject $rejection `
                    -Name patchRelativePath `
                    -Pattern '^External/TerminalEngine/CandidatePatches/[a-z0-9][a-z0-9._-]{0,63}\.patch$' `
                    -Context 'CandidateSet adaptation rejection'))
        [void]$paths.Add((Get-RSRequiredString `
                    -InputObject $rejection `
                    -Name concernMappingRelativePath `
                    -Pattern '^External/TerminalEngine/CandidatePatches/[a-z0-9][a-z0-9._-]{0,63}\.concerns\.json$' `
                    -Context 'CandidateSet adaptation rejection'))
    }

    $assessment = ConvertFrom-RSReadSnapshotJson `
        -LockSet $LockSet `
        -Path $CandidateAssessmentPath `
        -Context 'CandidateAssessment snapshot'
    $candidates = Get-RSRequiredArray `
        -InputObject $assessment `
        -Name candidates `
        -Context 'CandidateAssessment snapshot'
    if ($candidates.Count -gt 16) {
        throw [IO.InvalidDataException]::new(
            'CandidateAssessment snapshot exceeds the pre-schema candidate bound.')
    }
    foreach ($candidate in $candidates) {
        $rawLedger = (Get-RSPropertyValue `
                -InputObject $candidate `
                -Name rawLedger `
                -Required).Value
        $patch = (Get-RSPropertyValue `
                -InputObject $rawLedger `
                -Name patch `
                -Required).Value
        [void]$paths.Add((Get-RSRequiredString `
                    -InputObject $patch `
                    -Name patchRelativePath `
                    -Pattern '^External/TerminalEngine/CandidatePatches/[a-z0-9][a-z0-9._-]{0,63}\.patch$' `
                    -Context 'CandidateAssessment patch'))
        [void]$paths.Add((Get-RSRequiredString `
                    -InputObject $patch `
                    -Name concernMappingRelativePath `
                    -Pattern '^External/TerminalEngine/CandidatePatches/[a-z0-9][a-z0-9._-]{0,63}\.concerns\.json$' `
                    -Context 'CandidateAssessment patch'))
    }
    $result = [string[]]@($paths)
    [Array]::Sort($result, [StringComparer]::Ordinal)
    return ,$result
}

function Assert-RSRound3CollectionNeutralityBaseline {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $policyRelativePath =
        'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json'
    $schemaRelativePath =
        'Specs\Terminal\TerminalEngineCollectionNeutrality.schema.json'
    $policyPath = Join-Path $RepoRoot $policyRelativePath
    $schemaPath = Join-Path $RepoRoot $schemaRelativePath
    if (-not [IO.File]::Exists($policyPath) -or
        -not [IO.File]::Exists($schemaPath)) {
        throw [IO.InvalidDataException]::new(
            'The frozen Round 3 collection-neutrality policy or schema is missing.')
    }

    $bytes = [IO.File]::ReadAllBytes($policyPath)
    if ($bytes.Length -eq 0 -or
        ($bytes.Length -ge 3 -and
            $bytes[0] -eq 0xef -and
            $bytes[1] -eq 0xbb -and
            $bytes[2] -eq 0xbf)) {
        throw [IO.InvalidDataException]::new(
            'The Round 3 collection-neutrality policy must be nonempty UTF-8 without BOM.')
    }
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $policy = $json |
            ConvertFrom-Json `
                -AsHashtable `
                -Depth 100 `
                -NoEnumerate `
                -DateKind String
    }
    catch [System.Text.DecoderFallbackException] {
        throw [IO.InvalidDataException]::new(
            'The Round 3 collection-neutrality policy is not strict UTF-8.',
            $_.Exception)
    }
    catch [System.Management.Automation.PSInvalidOperationException] {
        throw [IO.InvalidDataException]::new(
            'The Round 3 collection-neutrality policy is not duplicate-free JSON.',
            $_.Exception)
    }

    [void](Import-RSExactModule -Path (Join-Path $RepoRoot (
                'Tools\TerminalEvidence.psm1')))
    $canonicalBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $policy
    if (-not [Linq.Enumerable]::SequenceEqual(
            [byte[]]$bytes,
            [byte[]]$canonicalBytes)) {
        throw [IO.InvalidDataException]::new(
            'The Round 3 collection-neutrality policy must be exact RFC 8785 JCS bytes.')
    }
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath $schemaPath `
        -Context 'Frozen Round 3 collection-neutrality policy'

    $neutralityModule = Import-RSExactModule -Path (Join-Path $RepoRoot (
            'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1'))
    $proof = & $neutralityModule {
        param($root, $policyRecord)
        Invoke-RSTerminalEngineCollectionNeutralityScan `
            -RepoRoot $root `
            -PolicyRecord $policyRecord `
            -ValidateBaselineProof
    } $RepoRoot $policy
    if ([string]$proof.status -cne 'verified' -or
        [int]$proof.findingCount -ne 0) {
        throw [IO.InvalidDataException]::new(
            'The Round 3 collection-neutrality baseline is not verified.')
    }
}

function Read-RSCandidateConcernMapping {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [object]$PatchRecord,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId
    )

    $relativePath = Get-RSRequiredString `
        -InputObject $PatchRecord `
        -Name concernMappingRelativePath `
        -Context "candidate '$CandidateId' patch"
    $mappingPath = [IO.Path]::GetFullPath((Join-Path $RepoRoot $relativePath))
    $mappingRoot = [IO.Path]::GetFullPath((Join-Path $RepoRoot (
                'External\TerminalEngine\CandidatePatches')))
    if (-not $mappingPath.StartsWith(
            $mappingRoot + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' concern mapping is missing or unsafe.")
    }

    $rawBytes = Read-RSImmutableRepoFileBytes `
        -RepoRoot $RepoRoot `
        -Path $mappingPath
    $mapping = [Text.UTF8Encoding]::new($false, $true).GetString($rawBytes) |
        ConvertFrom-Json -AsHashtable -Depth 64
    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $canonicalBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $mapping
    if (-not [Linq.Enumerable]::SequenceEqual(
            [byte[]]$rawBytes,
            [byte[]]$canonicalBytes)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' concern mapping is not canonical JCS.")
    }
    Test-RSJsonSchema `
        -Json ([Text.Encoding]::UTF8.GetString($canonicalBytes)) `
        -SchemaPath (Join-Path $RepoRoot $script:ConcernMappingSchemaLeaf) `
        -Context "Candidate '$CandidateId' concern mapping"

    foreach ($name in @('candidateId', 'upstreamTree', 'resultTree', 'patchSha256')) {
        $expected = if ($name -ceq 'candidateId') {
            $CandidateId
        } else {
            Get-RSRequiredString -InputObject $PatchRecord -Name $name
        }
        if ((Get-RSRequiredString -InputObject $mapping -Name $name) -cne $expected) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' concern mapping $name differs from its patch record.")
        }
    }

    $pathConcerns = Get-RSRequiredArray `
        -InputObject $mapping `
        -Name pathConcerns `
        -Context "candidate '$CandidateId' concern mapping"
    $paths = [Collections.Generic.List[string]]::new()
    $union = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in $pathConcerns) {
        $path = Get-RSRequiredString -InputObject $entry -Name path
        if ([IO.Path]::IsPathRooted($path) -or
            $path.Contains('\') -or
            $path.IndexOfAny([char[]]@([char]0, [char]9, [char]10, [char]13)) -ge 0) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' concern mapping path '$path' is unsafe.")
        }
        $segments = [string[]]$path.Split('/')
        if ($segments.Count -eq 0 -or
            @($segments | Where-Object {
                    [string]::IsNullOrEmpty($_) -or $_ -ceq '.' -or $_ -ceq '..'
                }).Count -ne 0) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' concern mapping path '$path' is not normalized.")
        }
        $paths.Add($path)
        $entryConcerns = [string[]](Get-RSRequiredArray `
                -InputObject $entry `
                -Name concerns `
                -Context "candidate '$CandidateId' concern mapping path '$path'")
        Assert-RSOrdinalSortedUnique `
            -Value $entryConcerns `
            -Context "candidate '$CandidateId' concern mapping path '$path'"
        foreach ($concern in $entryConcerns) {
            [void]$union.Add($concern)
        }
    }
    Assert-RSOrdinalSortedUnique `
        -Value $paths.ToArray() `
        -Context "candidate '$CandidateId' concern mapping paths"
    $expectedConcerns = [string[]](Get-RSRequiredArray `
            -InputObject $PatchRecord `
            -Name logicalConcerns `
            -Context "candidate '$CandidateId' patch")
    Assert-RSExactOrdinalSet `
        -Actual ([string[]]@($union)) `
        -Expected $expectedConcerns `
        -Context "candidate '$CandidateId' concern mapping taxonomy closure"

    $mappingPayload = [ordered]@{
        candidateId = Get-RSRequiredString -InputObject $mapping -Name candidateId
        upstreamTree = Get-RSRequiredString -InputObject $mapping -Name upstreamTree
        resultTree = Get-RSRequiredString -InputObject $mapping -Name resultTree
        patchSha256 = Get-RSRequiredString -InputObject $mapping -Name patchSha256
        pathConcerns = $pathConcerns
    }
    $mappingSha256 = Get-RSJcsSha256 -InputObject $mappingPayload
    if ($mappingSha256 -cne
        (Get-RSRequiredString -InputObject $mapping -Name mappingSha256) -or
        $mappingSha256 -cne
        (Get-RSRequiredString -InputObject $PatchRecord -Name concernMappingSha256)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' concern mapping digest is invalid.")
    }
    return $mapping
}

function Get-RSOrdinalSortedStrings {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Value
    )

    $result = [string[]]@($Value)
    [Array]::Sort($result, [StringComparer]::Ordinal)
    return ,$result
}

function Assert-RSOrdinalSortedUnique {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    for ($index = 1; $index -lt $Value.Count; ++$index) {
        $comparison = [StringComparer]::Ordinal.Compare($Value[$index - 1], $Value[$index])
        if ($comparison -eq 0) {
            throw [IO.InvalidDataException]::new("$Context contains duplicate '$($Value[$index])'.")
        }
        if ($comparison -gt 0) {
            throw [IO.InvalidDataException]::new("$Context is not in ascending ordinal order.")
        }
    }
}

function Sort-RSObjectsOrdinal {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$InputObject,

        [Parameter(Mandatory = $true)]
        [string[]]$KeyProperty
    )

    $result = [object[]]@($InputObject)
    $comparison = [Comparison[object]] {
        param($left, $right)
        foreach ($name in $KeyProperty) {
            $leftProperty = Get-RSPropertyValue -InputObject $left -Name $name -Required
            $rightProperty = Get-RSPropertyValue -InputObject $right -Name $name -Required
            if ($leftProperty.Value -isnot [string] -or $rightProperty.Value -isnot [string]) {
                throw [IO.InvalidDataException]::new("Ordinal sort key '$name' must be a string.")
            }
            $order = [StringComparer]::Ordinal.Compare(
                [string]$leftProperty.Value,
                [string]$rightProperty.Value)
            if ($order -ne 0) {
                return $order
            }
        }
        return 0
    }
    [Array]::Sort($result, $comparison)
    return ,$result
}

function Assert-RSExactOrdinalSet {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Actual,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $actualSorted = Get-RSOrdinalSortedStrings -Value $Actual
    $expectedSorted = Get-RSOrdinalSortedStrings -Value $Expected
    if (($actualSorted -join "`n") -cne ($expectedSorted -join "`n")) {
        throw [IO.InvalidDataException]::new("$Context does not match the frozen candidate set.")
    }
    Assert-RSOrdinalSortedUnique -Value $actualSorted -Context $Context
}

function Get-RSRepositoryState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $commitOutput = @(& git -C $RepoRoot rev-parse HEAD 2>&1)
    if ($LASTEXITCODE -ne 0 -or $commitOutput.Count -ne 1 -or
        [string]$commitOutput[0] -cnotmatch '^[0-9a-f]{40}$') {
        throw [IO.InvalidDataException]::new('Unable to resolve one exact repository commit.')
    }
    $statusOutput = @(& git -C $RepoRoot status --porcelain=v1 --untracked-files=all 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw [IO.InvalidDataException]::new('Unable to verify repository cleanliness.')
    }
    if ($statusOutput.Count -ne 0) {
        throw [IO.InvalidDataException]::new(
            'TerminalEngine evidence requires a clean repository, including no untracked files.')
    }

    return [pscustomobject]@{
        Commit = [string]$commitOutput[0]
        Clean = $true
    }
}

function Get-RSTerminalHarnessSourceSha256 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $evidenceModulePath = Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1'
    Import-Module $evidenceModulePath

    $identities = [object[]]@(
        foreach ($leaf in $script:HarnessFiles) {
            [ordered]@{
                leaf = $leaf.Replace('\', '/')
                sha256 = Get-RSFileSha256 -Path (Join-Path $RepoRoot $leaf)
            }
        }
    )
    return Get-RSJcsIdentitySetSha256 -Records $identities -KeyProperty leaf
}

function Get-RSCurrentCandidateUniverseFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $transitionPath = Join-Path $RepoRoot $script:CandidateUniverseLeaf
    if ([IO.File]::Exists($transitionPath)) {
        return $transitionPath
    }
    throw [IO.InvalidDataException]::new(
        'The required round-4 candidate universe is missing; current authority cannot fall back.')
}

function Get-RSFrozenState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet
    )

    return [ordered]@{
        requirementsSha256 = Get-RSFileSha256 -Path (Join-Path $RepoRoot $script:RequirementsLeaf)
        scoreAnchorsSha256 = Get-RSFileSha256 -Path (Join-Path $RepoRoot $script:ScoreAnchorsLeaf)
        fixtureCorpusSha256 = Get-RSFileSha256 -Path (Join-Path $RepoRoot $script:FixtureCorpusLeaf)
        measurementProtocolSha256 = Get-RSFileSha256 -Path (Join-Path $RepoRoot $script:MeasurementProtocolLeaf)
        candidateUniverseSha256 = Get-RSFileSha256 -Path (
            Get-RSCurrentCandidateUniverseFilePath -RepoRoot $RepoRoot)
        candidateSetSha256 = Get-RSCandidateSetSha256 -CandidateSet $CandidateSet
        candidateReviewSha256 = Get-RSFileSha256 -Path (Join-Path $RepoRoot $script:CandidateReviewLeaf)
        candidateAssessmentSha256 = Get-RSFileSha256 -Path (
            Join-Path $RepoRoot $script:CandidateAssessmentLeaf)
    }
}

function Read-RSCandidateUniverse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $path = Join-Path $RepoRoot $script:CandidateUniverseBaseLeaf
    $schemaPath = Join-Path $RepoRoot $script:CandidateUniverseBaseSchemaLeaf
    if (-not [IO.File]::Exists($path) -or -not [IO.File]::Exists($schemaPath)) {
        throw [IO.InvalidDataException]::new(
            'The frozen candidate universe or its closed schema is missing.')
    }
    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -eq 0 -or
        ($bytes.Length -ge 3 -and
            $bytes[0] -eq 0xef -and
            $bytes[1] -eq 0xbb -and
            $bytes[2] -eq 0xbf)) {
        throw [IO.InvalidDataException]::new(
            'The candidate universe must be nonempty UTF-8 without BOM.')
    }
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $universe = $json |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    } catch {
        throw [IO.InvalidDataException]::new(
            'The candidate universe is not duplicate-free UTF-8 JSON.',
            $_.Exception)
    }
    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $canonicalBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $universe
    if ([Convert]::ToBase64String($bytes) -cne
        [Convert]::ToBase64String($canonicalBytes)) {
        throw [IO.InvalidDataException]::new(
            'The candidate universe must be exact RFC 8785 JCS bytes.')
    }
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath $schemaPath `
        -Context 'Frozen candidate universe'

    $proposal = [ordered]@{}
    foreach ($entry in $universe.GetEnumerator()) {
        if ([string]$entry.Key -cnotin @('proposalSha256', 'approvals')) {
            $proposal[[string]$entry.Key] = $entry.Value
        }
    }
    $proposalSha256 = Get-RSJcsSha256 -InputObject $proposal
    if ($proposalSha256 -cne
        (Get-RSRequiredString `
            -InputObject $universe `
            -Name proposalSha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'candidate universe')) {
        throw [IO.InvalidDataException]::new(
            'The candidate-universe proposal digest is not derived from its exact proposal.')
    }

    $authorId = Get-RSRequiredString `
        -InputObject $universe `
        -Name authorId `
        -Pattern '^[a-z0-9][a-z0-9._-]{0,63}$' `
        -Context 'candidate universe'
    $seedCandidates = Get-RSRequiredArray `
        -InputObject $universe `
        -Name seedCandidates `
        -Context 'candidate universe'
    $seedIds = [string[]]@(
        $seedCandidates | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        })
    Assert-RSOrdinalSortedUnique -Value $seedIds -Context 'candidate-universe seed candidates'
    $nominationIds = [string[]]@(
        foreach ($seed in $seedCandidates) {
            $permission = Get-RSRequiredString `
                -InputObject $seed `
                -Name adaptationPermission `
                -Context 'candidate-universe seed'
            $nomination = Get-RSPropertyValue `
                -InputObject $seed `
                -Name allowedNominationId `
                -Required
            if ($permission -ceq 'allowed-existing-subsystem') {
                if ($nomination.Value -isnot [string] -or
                    [string]::IsNullOrEmpty([string]$nomination.Value)) {
                    throw [IO.InvalidDataException]::new(
                        'Every allowed seed must name exactly one nomination opportunity.')
                }
                [string]$nomination.Value
            } elseif ($null -ne $nomination.Value) {
                throw [IO.InvalidDataException]::new(
                    'A forbidden absent-subsystem seed cannot name a nomination.')
            }
        })
    if ($nominationIds.Count -ne 0) {
        $sortedNominationIds = Get-RSOrdinalSortedStrings -Value $nominationIds
        Assert-RSOrdinalSortedUnique `
            -Value $sortedNominationIds `
            -Context 'candidate-universe nomination IDs'
    }
    $controls = Get-RSRequiredArray -InputObject $universe -Name controls
    $controlIds = [string[]]@(
        $controls | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        })
    Assert-RSOrdinalSortedUnique -Value $controlIds -Context 'candidate-universe controls'

    $selection = (Get-RSPropertyValue `
            -InputObject (Get-RSRequirements -RepoRoot $RepoRoot) `
            -Name selection `
            -Required).Value
    $requiredApprovals = [int](Get-RSPropertyValue `
            -InputObject $selection `
            -Name reviewerApprovalCount `
            -Required).Value
    $approvals = Get-RSRequiredArray -InputObject $universe -Name approvals
    if ($approvals.Count -ne $requiredApprovals) {
        throw [IO.InvalidDataException]::new(
            "Candidate universe requires exactly $requiredApprovals independent approvals.")
    }
    $approvalIds = [string[]]@(
        $approvals | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name approvalId
        })
    Assert-RSOrdinalSortedUnique -Value $approvalIds -Context 'candidate-universe approvals'
    $reviewerIds = [string[]]@()
    foreach ($approval in $approvals) {
        if ((Get-RSRequiredString -InputObject $approval -Name proposalSha256) -cne
            $proposalSha256) {
            throw [IO.InvalidDataException]::new(
                'Candidate-universe approval is bound to a different proposal.')
        }
        $reviewerId = Get-RSRequiredString -InputObject $approval -Name reviewerId
        if ($reviewerId -ceq $authorId) {
            throw [IO.InvalidDataException]::new(
                'Candidate-universe author cannot approve the universe.')
        }
        $reviewerIds += $reviewerId
    }
    $sortedReviewerIds = Get-RSOrdinalSortedStrings -Value $reviewerIds
    Assert-RSOrdinalSortedUnique `
        -Value $sortedReviewerIds `
        -Context 'candidate-universe reviewers'
    $null = Get-RSUniverseNominationWindowOpenUtc `
        -Universe $universe `
        -Context 'candidate universe'

    return $universe
}

function Read-RSRound2CandidateUniverseCore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [switch]$HistoricalBase
    )

    $path = Join-Path $RepoRoot $script:CandidateUniverseRound2Leaf
    if (-not [IO.File]::Exists($path)) {
        throw [IO.InvalidDataException]::new(
            'The exact round-2 candidate-universe transition is missing; the named reader cannot fall back.')
    }
    $schemaPath = Join-Path $RepoRoot $script:CandidateUniverseRound2SchemaLeaf
    if (-not [IO.File]::Exists($path) -or -not [IO.File]::Exists($schemaPath)) {
        throw [IO.InvalidDataException]::new(
            'The frozen candidate-universe transition or its closed schema is missing.')
    }

    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -eq 0 -or
        ($bytes.Length -ge 3 -and
            $bytes[0] -eq 0xef -and
            $bytes[1] -eq 0xbb -and
            $bytes[2] -eq 0xbf)) {
        throw [IO.InvalidDataException]::new(
            'The candidate-universe transition must be nonempty UTF-8 without BOM.')
    }
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $transition = $json |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    } catch {
        throw [IO.InvalidDataException]::new(
            'The candidate-universe transition is not duplicate-free UTF-8 JSON.',
            $_.Exception)
    }

    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $canonicalBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $transition
    if ([Convert]::ToBase64String($bytes) -cne
        [Convert]::ToBase64String($canonicalBytes)) {
        throw [IO.InvalidDataException]::new(
            'The candidate-universe transition must be exact RFC 8785 JCS bytes.')
    }
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath $schemaPath `
        -Context 'Frozen candidate-universe transition'

    $proposal = [ordered]@{}
    foreach ($entry in $transition.GetEnumerator()) {
        if ([string]$entry.Key -cnotin @('proposalSha256', 'approvals')) {
            $proposal[[string]$entry.Key] = $entry.Value
        }
    }
    $proposalSha256 = Get-RSJcsSha256 -InputObject $proposal
    if ($proposalSha256 -cne
        (Get-RSRequiredString `
            -InputObject $transition `
            -Name proposalSha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'candidate-universe transition')) {
        throw [IO.InvalidDataException]::new(
            'The candidate-universe transition proposal digest is not derived from its exact proposal.')
    }

    $authorId = Get-RSRequiredString `
        -InputObject $transition `
        -Name authorId `
        -Pattern '^[a-z0-9][a-z0-9._-]{0,63}$' `
        -Context 'candidate-universe transition'
    $selection = (Get-RSPropertyValue `
            -InputObject (Get-RSRequirements -RepoRoot $RepoRoot) `
            -Name selection `
            -Required).Value
    $requiredApprovals = [int](Get-RSPropertyValue `
            -InputObject $selection `
            -Name reviewerApprovalCount `
            -Required).Value
    $approvals = Get-RSRequiredArray -InputObject $transition -Name approvals
    if ($approvals.Count -ne $requiredApprovals) {
        throw [IO.InvalidDataException]::new(
            "Candidate-universe transition requires exactly $requiredApprovals independent approvals.")
    }
    $approvalIds = [string[]]@(
        $approvals | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name approvalId
        })
    Assert-RSOrdinalSortedUnique `
        -Value $approvalIds `
        -Context 'candidate-universe transition approvals'
    $reviewerIds = [string[]]@()
    foreach ($approval in $approvals) {
        if ((Get-RSRequiredString -InputObject $approval -Name proposalSha256) -cne
            $proposalSha256) {
            throw [IO.InvalidDataException]::new(
                'Candidate-universe transition approval is bound to a different proposal.')
        }
        $reviewerId = Get-RSRequiredString -InputObject $approval -Name reviewerId
        if ($reviewerId -ceq $authorId) {
            throw [IO.InvalidDataException]::new(
                'Candidate-universe transition author cannot approve the transition.')
        }
        $reviewerIds += $reviewerId
    }
    Assert-RSOrdinalSortedUnique `
        -Value (Get-RSOrdinalSortedStrings -Value $reviewerIds) `
        -Context 'candidate-universe transition reviewers'
    $null = Get-RSUniverseNominationWindowOpenUtc `
        -Universe $transition `
        -Context 'candidate-universe transition'

    $base = (Get-RSPropertyValue `
            -InputObject $transition `
            -Name baseUniverse `
            -Required).Value
    $baseRelativePath = Get-RSRequiredString `
        -InputObject $base `
        -Name relativePath `
        -Context 'candidate-universe transition base'
    if ($baseRelativePath -cne $script:CandidateUniverseBaseLeaf.Replace('\', '/')) {
        throw [IO.InvalidDataException]::new(
            'The candidate-universe transition references an unexpected base universe.')
    }
    $baseSha256 = Get-RSRequiredString `
        -InputObject $base `
        -Name sha256 `
        -Pattern '^[0-9a-f]{64}$' `
        -Context 'candidate-universe transition base'
    if ((Get-RSFileSha256 `
                -Path (Join-Path $RepoRoot $script:CandidateUniverseBaseLeaf)) -cne
        $baseSha256) {
        throw [IO.InvalidDataException]::new(
            'The candidate-universe transition base digest does not match the frozen file.')
    }

    $governanceSchemas = (Get-RSPropertyValue `
            -InputObject $transition `
            -Name governanceSchemas `
            -Required).Value
    foreach ($schemaName in @('baseUniverse', 'transition')) {
        $schemaRecord = (Get-RSPropertyValue `
                -InputObject $governanceSchemas `
                -Name $schemaName `
                -Required).Value
        $schemaRelativePath = Get-RSRequiredString `
            -InputObject $schemaRecord `
            -Name relativePath `
            -Context "candidate-universe $schemaName schema"
        $expectedSchemaSha256 = Get-RSRequiredString `
            -InputObject $schemaRecord `
            -Name sha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context "candidate-universe $schemaName schema"
        if ((Get-RSFileSha256 `
                    -Path (Join-Path $RepoRoot $schemaRelativePath.Replace('/', '\'))) -cne
                $expectedSchemaSha256) {
            throw [IO.InvalidDataException]::new(
                "Candidate-universe governance schema '$schemaName' has changed.")
        }
    }

    $frozenInputs = Get-RSRequiredArray `
        -InputObject $transition `
        -Name frozenInputs `
        -Context 'candidate-universe transition'
    $seenFrozenInputs =
        [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($inputRecord in $frozenInputs) {
        $relativePath = Get-RSRequiredString `
            -InputObject $inputRecord `
            -Name relativePath `
            -Context 'candidate-universe transition frozen input'
        if (-not $seenFrozenInputs.Add($relativePath)) {
            throw [IO.InvalidDataException]::new(
                "Candidate-universe transition repeats frozen input '$relativePath'.")
        }
        $expectedSha256 = Get-RSRequiredString `
            -InputObject $inputRecord `
            -Name sha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'candidate-universe transition frozen input'
        if (-not $HistoricalBase) {
            $inputPath = Join-Path $RepoRoot $relativePath.Replace('/', '\')
            if ((Get-RSFileSha256 -Path $inputPath) -cne $expectedSha256) {
                throw [IO.InvalidDataException]::new(
                    "Candidate-universe transition frozen input '$relativePath' has changed.")
            }
        }
    }

    $corrections = Get-RSRequiredArray `
        -InputObject $transition `
        -Name conformanceCorrections `
        -Context 'candidate-universe transition'
    $correctionIds = Get-RSOrdinalSortedStrings -Value (
        [string[]]@(
            $corrections | ForEach-Object {
                Get-RSRequiredString -InputObject $_ -Name correctionId
            }))
    $expectedCorrectionIds = [string[]]@(
        'c7-proof-origin-closure',
        'portable-package-measurement',
        'positive-kitty-png-execution')
    if (($correctionIds -join "`0") -cne ($expectedCorrectionIds -join "`0")) {
        throw [IO.InvalidDataException]::new(
            'Candidate-universe transition conformance corrections are incomplete.')
    }

    $universe = Read-RSCandidateUniverse -RepoRoot $RepoRoot
    $policy = (Get-RSPropertyValue `
            -InputObject $transition `
            -Name policy `
            -Required).Value
    $transitionCaps = (Get-RSPropertyValue `
            -InputObject $policy `
            -Name hardCaps `
            -Required).Value
    $basePolicy = (Get-RSPropertyValue `
            -InputObject $universe `
            -Name adaptationPolicy `
            -Required).Value
    $baseCaps = (Get-RSPropertyValue `
            -InputObject $basePolicy `
            -Name hardCaps `
            -Required).Value
    if ((Get-RSJcsSha256 -InputObject $transitionCaps) -cne
        (Get-RSJcsSha256 -InputObject $baseCaps)) {
        throw [IO.InvalidDataException]::new(
            'Round 2 claims unchanged adaptation caps, but its cap object differs from round 1.')
    }

    $seeds = Get-RSRequiredArray `
        -InputObject $universe `
        -Name seedCandidates `
        -Context 'candidate universe'
    $seedById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($seed in $seeds) {
        $seedById.Add(
            (Get-RSRequiredString -InputObject $seed -Name candidateId),
            $seed)
    }
    $overrides = Get-RSRequiredArray `
        -InputObject $policy `
        -Name nominationOverrides `
        -Context 'candidate-universe transition policy'
    $newNominationIds =
        [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($override in $overrides) {
        $seedId = Get-RSRequiredString -InputObject $override -Name seedCandidateId
        $priorNominationId = Get-RSRequiredString `
            -InputObject $override `
            -Name priorNominationId
        $nominationId = Get-RSRequiredString -InputObject $override -Name nominationId
        if (-not $seedById.ContainsKey($seedId) -or
            [string]$seedById[$seedId]['allowedNominationId'] -cne $priorNominationId -or
            -not $newNominationIds.Add($nominationId)) {
            throw [IO.InvalidDataException]::new(
                "Candidate-universe transition nomination override for '$seedId' is not exact.")
        }
        $seedById[$seedId]['allowedNominationId'] = $nominationId
    }

    $universe['version'] = 2
    $universe['frozenUtc'] = $transition['frozenUtc']
    $universe['authorId'] = $authorId
    $universe['proposalSha256'] = $proposalSha256
    $universe['approvals'] = $approvals
    return $universe
}

function Read-RSRound2CandidateUniverse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Read-RSRound2CandidateUniverseCore -RepoRoot $RepoRoot
}

function Read-RSHistoricalRound2CandidateUniverse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Read-RSRound2CandidateUniverseCore `
        -RepoRoot $RepoRoot `
        -HistoricalBase
}

function Read-RSRound3CandidateUniverseCore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [switch]$HistoricalBase
    )

    $path = Join-Path $RepoRoot $script:CandidateUniverseRound3Leaf
    if (-not [IO.File]::Exists($path)) {
        throw [IO.InvalidDataException]::new(
            'The exact round-3 candidate-universe transition is missing; the named reader cannot fall back.')
    }
    $schemaPath = Join-Path $RepoRoot $script:CandidateUniverseRound3SchemaLeaf
    if (-not [IO.File]::Exists($schemaPath)) {
        throw [IO.InvalidDataException]::new(
            'The frozen round-3 candidate-universe transition schema is missing.')
    }

    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -eq 0 -or
        ($bytes.Length -ge 3 -and
            $bytes[0] -eq 0xef -and
            $bytes[1] -eq 0xbb -and
            $bytes[2] -eq 0xbf)) {
        throw [IO.InvalidDataException]::new(
            'The round-3 candidate-universe transition must be nonempty UTF-8 without BOM.')
    }
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $transition = $json |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    } catch {
        throw [IO.InvalidDataException]::new(
            'The round-3 candidate-universe transition is not duplicate-free UTF-8 JSON.',
            $_.Exception)
    }

    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $canonicalBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $transition
    if ([Convert]::ToBase64String($bytes) -cne
        [Convert]::ToBase64String($canonicalBytes)) {
        throw [IO.InvalidDataException]::new(
            'The round-3 candidate-universe transition must be exact RFC 8785 JCS bytes.')
    }
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath $schemaPath `
        -Context 'Frozen round-3 candidate-universe transition'

    $proposal = [ordered]@{}
    foreach ($entry in $transition.GetEnumerator()) {
        if ([string]$entry.Key -cnotin @('proposalSha256', 'approvals')) {
            $proposal[[string]$entry.Key] = $entry.Value
        }
    }
    $proposalSha256 = Get-RSJcsSha256 -InputObject $proposal
    if ($proposalSha256 -cne
        (Get-RSRequiredString `
            -InputObject $transition `
            -Name proposalSha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'round-3 candidate-universe transition')) {
        throw [IO.InvalidDataException]::new(
            'The round-3 candidate-universe proposal digest is not derived from its exact proposal.')
    }

    $authorId = Get-RSRequiredString `
        -InputObject $transition `
        -Name authorId `
        -Pattern '^[a-z0-9][a-z0-9._-]{0,63}$' `
        -Context 'round-3 candidate-universe transition'
    $selection = (Get-RSPropertyValue `
            -InputObject (Get-RSRequirements -RepoRoot $RepoRoot) `
            -Name selection `
            -Required).Value
    $requiredApprovals = [int](Get-RSPropertyValue `
            -InputObject $selection `
            -Name reviewerApprovalCount `
            -Required).Value
    $approvals = Get-RSRequiredArray -InputObject $transition -Name approvals
    if ($approvals.Count -ne $requiredApprovals) {
        throw [IO.InvalidDataException]::new(
            "Round 3 requires exactly $requiredApprovals independent approvals.")
    }
    $approvalIds = [string[]]@(
        $approvals | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name approvalId
        })
    Assert-RSOrdinalSortedUnique `
        -Value $approvalIds `
        -Context 'round-3 candidate-universe approvals'
    $reviewerIds = [string[]]@()
    foreach ($approval in $approvals) {
        if ((Get-RSRequiredString -InputObject $approval -Name proposalSha256) -cne
            $proposalSha256) {
            throw [IO.InvalidDataException]::new(
                'A round-3 approval is bound to a different proposal.')
        }
        $reviewerId = Get-RSRequiredString -InputObject $approval -Name reviewerId
        if ($reviewerId -ceq $authorId) {
            throw [IO.InvalidDataException]::new(
                'The round-3 author cannot approve the transition.')
        }
        $reviewerIds += $reviewerId
    }
    Assert-RSOrdinalSortedUnique `
        -Value (Get-RSOrdinalSortedStrings -Value $reviewerIds) `
        -Context 'round-3 candidate-universe reviewers'
    $null = Get-RSUniverseNominationWindowOpenUtc `
        -Universe $transition `
        -Context 'round-3 candidate-universe transition'

    $base = (Get-RSPropertyValue `
            -InputObject $transition `
            -Name baseUniverse `
            -Required).Value
    $baseRelativePath = Get-RSRequiredString `
        -InputObject $base `
        -Name relativePath `
        -Context 'round-3 candidate-universe base'
    if ($baseRelativePath -cne
        $script:CandidateUniverseRound2Leaf.Replace('\', '/')) {
        throw [IO.InvalidDataException]::new(
            'Round 3 references an unexpected base universe.')
    }
    $baseSha256 = Get-RSRequiredString `
        -InputObject $base `
        -Name sha256 `
        -Pattern '^[0-9a-f]{64}$' `
        -Context 'round-3 candidate-universe base'
    if ((Get-RSFileSha256 `
                -Path (Join-Path $RepoRoot $script:CandidateUniverseRound2Leaf)) -cne
        $baseSha256) {
        throw [IO.InvalidDataException]::new(
            'The round-3 base digest does not match the frozen round-2 file.')
    }

    $governanceSchemas = (Get-RSPropertyValue `
            -InputObject $transition `
            -Name governanceSchemas `
            -Required).Value
    $schemaNames = [string[]]@(
        $governanceSchemas.GetEnumerator() |
            ForEach-Object { [string]$_.Key })
    Assert-RSOrdinalSortedUnique `
        -Value $schemaNames `
        -Context 'round-3 governance-schema names'
    foreach ($schemaName in $schemaNames) {
        $schemaRecord = (Get-RSPropertyValue `
                -InputObject $governanceSchemas `
                -Name $schemaName `
                -Required).Value
        $schemaRelativePath = Get-RSRequiredString `
            -InputObject $schemaRecord `
            -Name relativePath `
            -Context "round-3 $schemaName schema"
        $expectedSchemaSha256 = Get-RSRequiredString `
            -InputObject $schemaRecord `
            -Name sha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context "round-3 $schemaName schema"
        $verifyHistoricalSchema =
            -not $HistoricalBase -or
            $schemaName -cin @('round2Transition', 'round3Transition')
        if ($verifyHistoricalSchema -and
            (Get-RSFileSha256 `
                    -Path (Join-Path $RepoRoot $schemaRelativePath.Replace('/', '\'))) -cne
                $expectedSchemaSha256) {
            throw [IO.InvalidDataException]::new(
                "Round-3 governance schema '$schemaName' has changed.")
        }
    }

    $frozenInputs = Get-RSRequiredArray `
        -InputObject $transition `
        -Name frozenInputs `
        -Context 'round-3 candidate-universe transition'
    $seenFrozenInputs =
        [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $frozenInputPaths = [Collections.Generic.List[string]]::new()
    foreach ($inputRecord in $frozenInputs) {
        $relativePath = Get-RSRequiredString `
            -InputObject $inputRecord `
            -Name relativePath `
            -Context 'round-3 frozen input'
        if (-not $seenFrozenInputs.Add($relativePath)) {
            throw [IO.InvalidDataException]::new(
                "Round 3 repeats frozen input '$relativePath'.")
        }
        $frozenInputPaths.Add($relativePath)
        $expectedSha256 = Get-RSRequiredString `
            -InputObject $inputRecord `
            -Name sha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'round-3 frozen input'
        if (-not $HistoricalBase -and
            (Get-RSFileSha256 `
                    -Path (Join-Path $RepoRoot $relativePath.Replace('/', '\'))) -cne
                $expectedSha256) {
            throw [IO.InvalidDataException]::new(
                "Round-3 frozen input '$relativePath' has changed.")
        }
    }
    Assert-RSOrdinalSortedUnique `
        -Value $frozenInputPaths.ToArray() `
        -Context 'round-3 frozen-input paths'
    if (($frozenInputPaths.ToArray() -join "`0") -cne
        ($script:Round3FrozenInputLeaves -join "`0")) {
        $missingInputs = [string[]]@(
            $script:Round3FrozenInputLeaves |
                Where-Object { -not $seenFrozenInputs.Contains($_) })
        $unexpectedInputs = [string[]]@(
            $frozenInputPaths |
                Where-Object {
                    $_ -cnotin $script:Round3FrozenInputLeaves
                })
        throw [IO.InvalidDataException]::new(
            "Round 3 must freeze the exact mandatory input closure. Missing=[$($missingInputs -join ', ')]; unexpected=[$($unexpectedInputs -join ', ')].")
    }

    if (-not $HistoricalBase) {
        Assert-RSRound3CollectionNeutralityBaseline -RepoRoot $RepoRoot
    }

    $corrections = Get-RSRequiredArray `
        -InputObject $transition `
        -Name conformanceCorrections `
        -Context 'round-3 candidate-universe transition'
    $correctionIds = [string[]]@(
        $corrections | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name correctionId
        })
    if ($correctionIds.Count -ne 1 -or
        $correctionIds[0] -cne 'cross-repository-adaptation-ledger-closure') {
        throw [IO.InvalidDataException]::new(
            'The round-3 cross-repository conformance correction is incomplete.')
    }

    $universe = Read-RSHistoricalRound2CandidateUniverse -RepoRoot $RepoRoot
    $policy = (Get-RSPropertyValue `
            -InputObject $transition `
            -Name policy `
            -Required).Value
    $transitionCaps = (Get-RSPropertyValue `
            -InputObject $policy `
            -Name hardCaps `
            -Required).Value
    $basePolicy = (Get-RSPropertyValue `
            -InputObject $universe `
            -Name adaptationPolicy `
            -Required).Value
    $baseCaps = (Get-RSPropertyValue `
            -InputObject $basePolicy `
            -Name hardCaps `
            -Required).Value
    if ((Get-RSJcsSha256 -InputObject $transitionCaps) -cne
        (Get-RSJcsSha256 -InputObject $baseCaps)) {
        throw [IO.InvalidDataException]::new(
            'Round 3 claims unchanged caps, but its cap object differs from round 2.')
    }

    $seeds = Get-RSRequiredArray `
        -InputObject $universe `
        -Name seedCandidates `
        -Context 'round-2 effective candidate universe'
    $seedById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($seed in $seeds) {
        $seedById.Add(
            (Get-RSRequiredString -InputObject $seed -Name candidateId),
            $seed)
    }
    $overrides = Get-RSRequiredArray `
        -InputObject $policy `
        -Name nominationOverrides `
        -Context 'round-3 candidate-universe policy'
    $newNominationIds =
        [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($override in $overrides) {
        $seedId = Get-RSRequiredString `
            -InputObject $override `
            -Name seedCandidateId
        $priorNominationId = Get-RSRequiredString `
            -InputObject $override `
            -Name priorNominationId
        $nominationId = Get-RSRequiredString `
            -InputObject $override `
            -Name nominationId
        if (-not $seedById.ContainsKey($seedId) -or
            [string]$seedById[$seedId]['allowedNominationId'] -cne
                $priorNominationId -or
            -not $newNominationIds.Add($nominationId)) {
            throw [IO.InvalidDataException]::new(
                "Round-3 nomination override for '$seedId' is not exact.")
        }
        $seedById[$seedId]['allowedNominationId'] = $nominationId
    }

    $universe['version'] = 3
    $universe['frozenUtc'] = $transition['frozenUtc']
    $universe['authorId'] = $authorId
    $universe['proposalSha256'] = $proposalSha256
    $universe['approvals'] = $approvals
    return $universe
}

function Read-RSRound3CandidateUniverse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Read-RSRound3CandidateUniverseCore -RepoRoot $RepoRoot
}

function Read-RSHistoricalRound3CandidateUniverse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Read-RSRound3CandidateUniverseCore `
        -RepoRoot $RepoRoot `
        -HistoricalBase
}

function Assert-RSRound4AuthenticationPreflight {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Transition
    )

    $expectedGovernanceSchemas = [ordered]@{
        candidateAssessment = 'Specs/Terminal/TerminalEngineCandidateAssessment.schema.json'
        candidateReview = 'Specs/Terminal/TerminalEngineCandidateReview.schema.json'
        candidateSet = 'Specs/Terminal/TerminalEngineCandidateSet.schema.json'
        concernMapping = 'Specs/Terminal/TerminalEngineConcernMapping.schema.json'
        driverInput = 'Tools/TerminalEngine/TerminalEngineCollectionDriverInput.schema.json'
        driverResult = 'Tools/TerminalEngine/TerminalEngineCollectionDriverResult.schema.json'
        hostNeutrality = 'Specs/Terminal/TerminalEngineCollectionNeutrality.schema.json'
        providerDescriptor = 'Specs/Terminal/TerminalEngineCollectionDescriptor.schema.json'
        round2Transition = 'Specs/Terminal/TerminalEngineCandidateUniverseTransition.schema.json'
        round3Transition = 'Specs/Terminal/TerminalEngineCandidateUniverseTransition.v3.schema.json'
        round4Transition = 'Specs/Terminal/TerminalEngineCandidateUniverseTransition.v4.schema.json'
    }
    $governanceSchemas = (Get-RSPropertyValue `
            -InputObject $Transition `
            -Name governanceSchemas `
            -Required).Value
    if ($governanceSchemas -isnot [Collections.IDictionary]) {
        throw [IO.InvalidDataException]::new(
            'Round-4 governanceSchemas must be one object.')
    }
    $schemaNames = [string[]]@(
        $governanceSchemas.GetEnumerator() |
            ForEach-Object { [string]$_.Key })
    Assert-RSOrdinalSortedUnique `
        -Value $schemaNames `
        -Context 'round-4 governance-schema names'
    Assert-RSExactOrdinalSet `
        -Actual $schemaNames `
        -Expected ([string[]]@($expectedGovernanceSchemas.Keys)) `
        -Context 'round-4 governance-schema closure'
    foreach ($schemaName in $schemaNames) {
        $schemaRecord = (Get-RSPropertyValue `
                -InputObject $governanceSchemas `
                -Name $schemaName `
                -Required).Value
        $relativePath = Get-RSRequiredString `
            -InputObject $schemaRecord `
            -Name relativePath `
            -Context "round-4 $schemaName schema"
        if ($relativePath -cne [string]$expectedGovernanceSchemas[$schemaName]) {
            throw [IO.InvalidDataException]::new(
                "Round-4 governance schema '$schemaName' has an unexpected path.")
        }
        $expectedSha256 = Get-RSRequiredString `
            -InputObject $schemaRecord `
            -Name sha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context "round-4 $schemaName schema"
        if ((Get-RSFileSha256 -Path (Join-Path $RepoRoot (
                    $relativePath.Replace('/', '\')))) -cne
            $expectedSha256) {
            throw [IO.InvalidDataException]::new(
                "Round-4 governance schema '$schemaName' has changed.")
        }
    }

    $frozenInputs = Get-RSRequiredArray `
        -InputObject $Transition `
        -Name frozenInputs `
        -Context 'round-4 candidate-universe transition'
    $seen =
        [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $shaByPath =
        [Collections.Generic.Dictionary[string, string]]::new(
            [StringComparer]::Ordinal)
    $paths = [Collections.Generic.List[string]]::new()
    foreach ($record in $frozenInputs) {
        $relativePath = Get-RSRequiredString `
            -InputObject $record `
            -Name relativePath `
            -Context 'round-4 frozen input'
        if (-not $seen.Add($relativePath)) {
            throw [IO.InvalidDataException]::new(
                "Round 4 repeats frozen input '$relativePath'.")
        }
        $paths.Add($relativePath)
        $shaByPath.Add(
            $relativePath,
            (Get-RSRequiredString `
                -InputObject $record `
                -Name sha256 `
                -Pattern '^[0-9a-f]{64}$' `
                -Context 'round-4 frozen input'))
    }
    Assert-RSOrdinalSortedUnique `
        -Value $paths.ToArray() `
        -Context 'round-4 frozen-input paths'
    if (($paths.ToArray() -join "`0") -cne
        ($script:Round4FrozenInputLeaves -join "`0")) {
        throw [IO.InvalidDataException]::new(
            'Round 4 must freeze the exact mandatory input closure before any frozen path is read.')
    }
    foreach ($relativePath in $script:Round4FrozenInputLeaves) {
        if ((Get-RSFileSha256 -Path (Join-Path $RepoRoot (
                    $relativePath.Replace('/', '\')))) -cne
            $shaByPath[$relativePath]) {
            throw [IO.InvalidDataException]::new(
                "Round-4 frozen input '$relativePath' has changed.")
        }
    }
    return $shaByPath
}

function Write-RSTrustedGovernanceJsonValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [Text.Json.Utf8JsonWriter]$Writer
    )

    if ($null -eq $Value) {
        $Writer.WriteNullValue()
        return
    }
    if ($Value -is [Collections.IDictionary]) {
        $names = [Collections.Generic.List[string]]::new()
        foreach ($entry in $Value.GetEnumerator()) {
            if ($entry.Key -isnot [string]) {
                throw [IO.InvalidDataException]::new(
                    'Trusted governance JSON object keys must be strings.')
            }
            $names.Add([string]$entry.Key)
        }
        $orderedNames = [string[]]@($names)
        [Array]::Sort($orderedNames, [StringComparer]::Ordinal)
        $Writer.WriteStartObject()
        foreach ($name in $orderedNames) {
            $Writer.WritePropertyName($name)
            Write-RSTrustedGovernanceJsonValue `
                -Value $Value[$name] `
                -Writer $Writer
        }
        $Writer.WriteEndObject()
        return
    }
    if ($Value -is [Collections.IEnumerable] -and
        $Value -isnot [string]) {
        $Writer.WriteStartArray()
        foreach ($entry in $Value) {
            Write-RSTrustedGovernanceJsonValue `
                -Value $entry `
                -Writer $Writer
        }
        $Writer.WriteEndArray()
        return
    }
    if ($Value -is [string]) {
        $Writer.WriteStringValue([string]$Value)
        return
    }
    if ($Value -is [bool]) {
        $Writer.WriteBooleanValue([bool]$Value)
        return
    }
    if ($Value -is [byte] -or
        $Value -is [sbyte] -or
        $Value -is [int16] -or
        $Value -is [uint16] -or
        $Value -is [int32] -or
        $Value -is [uint32] -or
        $Value -is [int64]) {
        $Writer.WriteNumberValue([int64]$Value)
        return
    }
    throw [IO.InvalidDataException]::new(
        "Trusted governance JSON contains unsupported type '$($Value.GetType().FullName)'.")
}

function ConvertTo-RSTrustedGovernanceJcsUtf8Bytes {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject
    )

    $stream = [IO.MemoryStream]::new()
    $options = [Text.Json.JsonWriterOptions]::new()
    $options.Indented = $false
    $options.SkipValidation = $false
    $options.Encoder =
        [Text.Encodings.Web.JavaScriptEncoder]::UnsafeRelaxedJsonEscaping
    $writer = [Text.Json.Utf8JsonWriter]::new($stream, $options)
    try {
        Write-RSTrustedGovernanceJsonValue `
            -Value $InputObject `
            -Writer $writer
        $writer.Flush()
        return ,$stream.ToArray()
    } finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Get-RSTrustedGovernanceJcsSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject
    )

    return Get-RSBytesSha256 -Bytes (
        ConvertTo-RSTrustedGovernanceJcsUtf8Bytes `
            -InputObject $InputObject)
}

function Assert-RSRound4AuthorityPreflight {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Transition,

        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    $canonicalBytes = ConvertTo-RSTrustedGovernanceJcsUtf8Bytes `
        -InputObject $Transition
    if (-not [Linq.Enumerable]::SequenceEqual(
            [byte[]]$Bytes,
            [byte[]]$canonicalBytes)) {
        throw [IO.InvalidDataException]::new(
            'The round-4 candidate-universe transition must be exact trusted RFC 8785 JCS bytes.')
    }
    $proposal = [ordered]@{}
    foreach ($entry in $Transition.GetEnumerator()) {
        if ([string]$entry.Key -cnotin @('proposalSha256', 'approvals')) {
            $proposal[[string]$entry.Key] = $entry.Value
        }
    }
    $proposalSha256 = Get-RSTrustedGovernanceJcsSha256 `
        -InputObject $proposal
    if ($proposalSha256 -cne
        (Get-RSRequiredString `
            -InputObject $Transition `
            -Name proposalSha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'round-4 candidate-universe transition')) {
        throw [IO.InvalidDataException]::new(
            'The round-4 proposal digest is not derived from its exact trusted proposal.')
    }

    $authorId = Get-RSRequiredString `
        -InputObject $Transition `
        -Name authorId `
        -Pattern '^[a-z0-9][a-z0-9._-]{0,63}$' `
        -Context 'round-4 candidate-universe transition'
    $approvals = Get-RSRequiredArray `
        -InputObject $Transition `
        -Name approvals `
        -Context 'round-4 candidate-universe transition'
    $trustedRequiredApprovals = 2
    if ($approvals.Count -ne $trustedRequiredApprovals) {
        throw [IO.InvalidDataException]::new(
            "Round 4 requires exactly $trustedRequiredApprovals independent approvals at the evaluator trust root.")
    }
    $approvalIds = [string[]]@(
        $approvals | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name approvalId
        })
    Assert-RSOrdinalSortedUnique `
        -Value $approvalIds `
        -Context 'round-4 candidate-universe approvals'
    $reviewerIds = [string[]]@()
    foreach ($approval in $approvals) {
        if ((Get-RSRequiredString `
                    -InputObject $approval `
                    -Name proposalSha256) -cne
            $proposalSha256) {
            throw [IO.InvalidDataException]::new(
                'A round-4 approval is bound to a different proposal.')
        }
        $reviewerId = Get-RSRequiredString `
            -InputObject $approval `
            -Name reviewerId
        if ($reviewerId -ceq $authorId) {
            throw [IO.InvalidDataException]::new(
                'The round-4 author cannot approve the transition.')
        }
        $reviewerIds += $reviewerId
    }
    Assert-RSOrdinalSortedUnique `
        -Value (Get-RSOrdinalSortedStrings -Value $reviewerIds) `
        -Context 'round-4 candidate-universe reviewers'
    $null = Get-RSUniverseNominationWindowOpenUtc `
        -Universe $Transition `
        -Context 'round-4 candidate-universe transition'

    $base = (Get-RSPropertyValue `
            -InputObject $Transition `
            -Name baseUniverse `
            -Required).Value
    if ((Get-RSRequiredString `
                -InputObject $base `
                -Name relativePath `
                -Context 'round-4 candidate-universe base') -cne
        $script:CandidateUniverseRound3Leaf.Replace('\', '/')) {
        throw [IO.InvalidDataException]::new(
            'Round 4 references an unexpected base universe.')
    }
    if ((Get-RSFileSha256 -Path (Join-Path $RepoRoot (
                $script:CandidateUniverseRound3Leaf))) -cne
        (Get-RSRequiredString `
            -InputObject $base `
            -Name sha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'round-4 candidate-universe base')) {
        throw [IO.InvalidDataException]::new(
            'The round-4 base digest does not match the frozen round-3 file.')
    }
}

function Read-RSCurrentCandidateUniverseCore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [switch]$IncludeRecordIdentity
    )

    $path = Join-Path $RepoRoot $script:CandidateUniverseLeaf
    if (-not [IO.File]::Exists($path)) {
        throw [IO.InvalidDataException]::new(
            'The required round-4 candidate universe is missing; current authority cannot fall back.')
    }
    $schemaPath = Join-Path $RepoRoot $script:CandidateUniverseSchemaLeaf
    if (-not [IO.File]::Exists($schemaPath)) {
        throw [IO.InvalidDataException]::new(
            'The frozen round-4 candidate-universe transition schema is missing.')
    }

    $readSnapshot = New-RSReadSnapshotLockSet `
        -RepoRoot $RepoRoot `
        -Path (Get-RSRound4SnapshotPaths `
            -CurrentUniversePath $script:CandidateUniverseLeaf)
    try {
    $bytes = Read-RSReadSnapshotBytes `
        -LockSet $readSnapshot `
        -Path $path
    if ($bytes.Length -eq 0 -or
        ($bytes.Length -ge 3 -and
            $bytes[0] -eq 0xef -and
            $bytes[1] -eq 0xbb -and
            $bytes[2] -eq 0xbf)) {
        throw [IO.InvalidDataException]::new(
            'The round-4 candidate-universe transition must be nonempty UTF-8 without BOM.')
    }
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $transition = $json |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    } catch {
        throw [IO.InvalidDataException]::new(
            'The round-4 candidate-universe transition is not duplicate-free UTF-8 JSON.',
            $_.Exception)
    }

    Assert-RSRound4AuthorityPreflight `
        -RepoRoot $RepoRoot `
        -Transition $transition `
        -Bytes $bytes
    $null = Assert-RSRound4AuthenticationPreflight `
        -RepoRoot $RepoRoot `
        -Transition $transition
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath $schemaPath `
        -Context 'Frozen round-4 candidate-universe transition'
    $null = Import-RSAuthenticatedCollectionNeutralityModules `
        -RepoRoot $RepoRoot
    $canonicalBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $transition
    if ([Convert]::ToBase64String($bytes) -cne
        [Convert]::ToBase64String($canonicalBytes)) {
        throw [IO.InvalidDataException]::new(
            'The round-4 candidate-universe transition must be exact RFC 8785 JCS bytes.')
    }
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath $schemaPath `
        -Context 'Frozen round-4 candidate-universe transition'

    $proposal = [ordered]@{}
    foreach ($entry in $transition.GetEnumerator()) {
        if ([string]$entry.Key -cnotin @('proposalSha256', 'approvals')) {
            $proposal[[string]$entry.Key] = $entry.Value
        }
    }
    $proposalSha256 = Get-RSJcsSha256 -InputObject $proposal
    if ($proposalSha256 -cne
        (Get-RSRequiredString `
            -InputObject $transition `
            -Name proposalSha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'round-4 candidate-universe transition')) {
        throw [IO.InvalidDataException]::new(
            'The round-4 candidate-universe proposal digest is not derived from its exact proposal.')
    }

    $authorId = Get-RSRequiredString `
        -InputObject $transition `
        -Name authorId `
        -Pattern '^[a-z0-9][a-z0-9._-]{0,63}$' `
        -Context 'round-4 candidate-universe transition'
    $selection = (Get-RSPropertyValue `
            -InputObject (Get-RSRequirements -RepoRoot $RepoRoot) `
            -Name selection `
            -Required).Value
    $requiredApprovals = [int](Get-RSPropertyValue `
            -InputObject $selection `
            -Name reviewerApprovalCount `
            -Required).Value
    $approvals = Get-RSRequiredArray -InputObject $transition -Name approvals
    if ($approvals.Count -ne $requiredApprovals) {
        throw [IO.InvalidDataException]::new(
            "Round 4 requires exactly $requiredApprovals independent approvals.")
    }
    $approvalIds = [string[]]@(
        $approvals | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name approvalId
        })
    Assert-RSOrdinalSortedUnique `
        -Value $approvalIds `
        -Context 'round-4 candidate-universe approvals'
    $reviewerIds = [string[]]@()
    foreach ($approval in $approvals) {
        if ((Get-RSRequiredString -InputObject $approval -Name proposalSha256) -cne
            $proposalSha256) {
            throw [IO.InvalidDataException]::new(
                'A round-4 approval is bound to a different proposal.')
        }
        $reviewerId = Get-RSRequiredString -InputObject $approval -Name reviewerId
        if ($reviewerId -ceq $authorId) {
            throw [IO.InvalidDataException]::new(
                'The round-4 author cannot approve the transition.')
        }
        $reviewerIds += $reviewerId
    }
    Assert-RSOrdinalSortedUnique `
        -Value (Get-RSOrdinalSortedStrings -Value $reviewerIds) `
        -Context 'round-4 candidate-universe reviewers'
    $null = Get-RSUniverseNominationWindowOpenUtc `
        -Universe $transition `
        -Context 'round-4 candidate-universe transition'

    $base = (Get-RSPropertyValue `
            -InputObject $transition `
            -Name baseUniverse `
            -Required).Value
    $baseRelativePath = Get-RSRequiredString `
        -InputObject $base `
        -Name relativePath `
        -Context 'round-4 candidate-universe base'
    if ($baseRelativePath -cne
        $script:CandidateUniverseRound3Leaf.Replace('\', '/')) {
        throw [IO.InvalidDataException]::new(
            'Round 4 references an unexpected base universe.')
    }
    $baseSha256 = Get-RSRequiredString `
        -InputObject $base `
        -Name sha256 `
        -Pattern '^[0-9a-f]{64}$' `
        -Context 'round-4 candidate-universe base'
    if ((Get-RSFileSha256 `
                -Path (Join-Path $RepoRoot $script:CandidateUniverseRound3Leaf)) -cne
        $baseSha256) {
        throw [IO.InvalidDataException]::new(
            'The round-4 base digest does not match the frozen round-3 file.')
    }

    $expectedGovernanceSchemas = [ordered]@{
        candidateAssessment = 'Specs/Terminal/TerminalEngineCandidateAssessment.schema.json'
        candidateReview = 'Specs/Terminal/TerminalEngineCandidateReview.schema.json'
        candidateSet = 'Specs/Terminal/TerminalEngineCandidateSet.schema.json'
        concernMapping = 'Specs/Terminal/TerminalEngineConcernMapping.schema.json'
        driverInput = 'Tools/TerminalEngine/TerminalEngineCollectionDriverInput.schema.json'
        driverResult = 'Tools/TerminalEngine/TerminalEngineCollectionDriverResult.schema.json'
        hostNeutrality = 'Specs/Terminal/TerminalEngineCollectionNeutrality.schema.json'
        providerDescriptor = 'Specs/Terminal/TerminalEngineCollectionDescriptor.schema.json'
        round2Transition = 'Specs/Terminal/TerminalEngineCandidateUniverseTransition.schema.json'
        round3Transition = 'Specs/Terminal/TerminalEngineCandidateUniverseTransition.v3.schema.json'
        round4Transition = 'Specs/Terminal/TerminalEngineCandidateUniverseTransition.v4.schema.json'
    }
    $governanceSchemas = (Get-RSPropertyValue `
            -InputObject $transition `
            -Name governanceSchemas `
            -Required).Value
    $schemaNames = [string[]]@(
        $governanceSchemas.GetEnumerator() |
            ForEach-Object { [string]$_.Key })
    Assert-RSOrdinalSortedUnique `
        -Value $schemaNames `
        -Context 'round-4 governance-schema names'
    Assert-RSExactOrdinalSet `
        -Actual $schemaNames `
        -Expected ([string[]]@($expectedGovernanceSchemas.Keys)) `
        -Context 'round-4 governance-schema closure'
    foreach ($schemaName in $schemaNames) {
        $schemaRecord = (Get-RSPropertyValue `
                -InputObject $governanceSchemas `
                -Name $schemaName `
                -Required).Value
        $schemaRelativePath = Get-RSRequiredString `
            -InputObject $schemaRecord `
            -Name relativePath `
            -Context "round-4 $schemaName schema"
        if ($schemaRelativePath -cne [string]$expectedGovernanceSchemas[$schemaName]) {
            throw [IO.InvalidDataException]::new(
                "Round-4 governance schema '$schemaName' has an unexpected path.")
        }
        $expectedSchemaSha256 = Get-RSRequiredString `
            -InputObject $schemaRecord `
            -Name sha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context "round-4 $schemaName schema"
        if ((Get-RSFileSha256 `
                -Path (Join-Path $RepoRoot $schemaRelativePath.Replace('/', '\'))) -cne
            $expectedSchemaSha256) {
            throw [IO.InvalidDataException]::new(
                "Round-4 governance schema '$schemaName' has changed.")
        }
    }

    $frozenInputs = Get-RSRequiredArray `
        -InputObject $transition `
        -Name frozenInputs `
        -Context 'round-4 candidate-universe transition'
    $seenFrozenInputs =
        [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $frozenInputSha256ByRelativePath =
        [Collections.Generic.Dictionary[string, string]]::new(
            [StringComparer]::Ordinal)
    $frozenInputPaths = [Collections.Generic.List[string]]::new()
    foreach ($inputRecord in $frozenInputs) {
        $relativePath = Get-RSRequiredString `
            -InputObject $inputRecord `
            -Name relativePath `
            -Context 'round-4 frozen input'
        if (-not $seenFrozenInputs.Add($relativePath)) {
            throw [IO.InvalidDataException]::new(
                "Round 4 repeats frozen input '$relativePath'.")
        }
        $frozenInputPaths.Add($relativePath)
        $expectedSha256 = Get-RSRequiredString `
            -InputObject $inputRecord `
            -Name sha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'round-4 frozen input'
        $frozenInputSha256ByRelativePath.Add(
            $relativePath,
            $expectedSha256)
    }
    Assert-RSOrdinalSortedUnique `
        -Value $frozenInputPaths.ToArray() `
        -Context 'round-4 frozen-input paths'
    if (($frozenInputPaths.ToArray() -join "`0") -cne
        ($script:Round4FrozenInputLeaves -join "`0")) {
        $missingInputs = [string[]]@(
            $script:Round4FrozenInputLeaves |
                Where-Object { -not $seenFrozenInputs.Contains($_) })
        $unexpectedInputs = [string[]]@(
            $frozenInputPaths |
                Where-Object { $_ -cnotin $script:Round4FrozenInputLeaves })
        throw [IO.InvalidDataException]::new(
            "Round 4 must freeze the exact mandatory input closure. Missing=[$($missingInputs -join ', ')]; unexpected=[$($unexpectedInputs -join ', ')].")
    }
    foreach ($relativePath in $script:Round4FrozenInputLeaves) {
        if ((Get-RSFileSha256 `
                -Path (Join-Path $RepoRoot $relativePath.Replace('/', '\'))) -cne
            $frozenInputSha256ByRelativePath[$relativePath]) {
            throw [IO.InvalidDataException]::new(
                "Round-4 frozen input '$relativePath' has changed.")
        }
    }

    Assert-RSRound3CollectionNeutralityBaseline -RepoRoot $RepoRoot

    $corrections = Get-RSRequiredArray `
        -InputObject $transition `
        -Name conformanceCorrections `
        -Context 'round-4 candidate-universe transition'
    $correctionIds = [string[]]@(
        $corrections | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name correctionId
        })
    $expectedCorrectionIds = [string[]]@(
        'cross-repository-adaptation-ledger-closure',
        'candidate-governance-chronology-closure'
    )
    if (($correctionIds -join "`0") -cne ($expectedCorrectionIds -join "`0")) {
        throw [IO.InvalidDataException]::new(
            'The round-4 ordered conformance-correction closure is incomplete.')
    }

    $universe = Read-RSHistoricalRound3CandidateUniverse -RepoRoot $RepoRoot
    $policy = (Get-RSPropertyValue `
            -InputObject $transition `
            -Name policy `
            -Required).Value
    $transitionCaps = (Get-RSPropertyValue `
            -InputObject $policy `
            -Name hardCaps `
            -Required).Value
    $basePolicy = (Get-RSPropertyValue `
            -InputObject $universe `
            -Name adaptationPolicy `
            -Required).Value
    $baseCaps = (Get-RSPropertyValue `
            -InputObject $basePolicy `
            -Name hardCaps `
            -Required).Value
    if ((Get-RSJcsSha256 -InputObject $transitionCaps) -cne
        (Get-RSJcsSha256 -InputObject $baseCaps)) {
        throw [IO.InvalidDataException]::new(
            'Round 4 claims unchanged caps, but its cap object differs from round 3.')
    }

    $carryForward = Get-RSRequiredArray `
        -InputObject $policy `
        -Name nominationCarryForward `
        -Context 'round-4 candidate-universe policy'
    $allowedSeeds = @(
        (Get-RSRequiredArray `
            -InputObject $universe `
            -Name seedCandidates `
            -Context 'round-3 effective candidate universe') |
            Where-Object {
                (Get-RSRequiredString `
                    -InputObject $_ `
                    -Name adaptationPermission) -ceq 'allowed-existing-subsystem'
            }
    )
    $expectedSeedIds = [string[]]@(
        $allowedSeeds | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        })
    $carrySeedIds = [string[]]@(
        $carryForward | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name seedCandidateId
        })
    Assert-RSOrdinalSortedUnique `
        -Value $carrySeedIds `
        -Context 'round-4 nomination carry-forward'
    Assert-RSExactOrdinalSet `
        -Actual $carrySeedIds `
        -Expected $expectedSeedIds `
        -Context 'round-4 nomination carry-forward seeds'
    $allowedById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($seed in $allowedSeeds) {
        $allowedById.Add(
            (Get-RSRequiredString -InputObject $seed -Name candidateId),
            $seed)
    }
    foreach ($record in $carryForward) {
        $seedCandidateId = Get-RSRequiredString `
            -InputObject $record `
            -Name seedCandidateId
        if ((Get-RSRequiredString -InputObject $record -Name nominationId) -cne
            (Get-RSRequiredString `
                -InputObject $allowedById[$seedCandidateId] `
                -Name allowedNominationId)) {
            throw [IO.InvalidDataException]::new(
                "Round-4 nomination carry-forward for '$seedCandidateId' changed its identity.")
        }
    }

    $universe['version'] = 4
    $universe['frozenUtc'] = $transition['frozenUtc']
    $universe['authorId'] = $authorId
    $universe['proposalSha256'] = $proposalSha256
    $universe['approvals'] = $approvals
    if ($IncludeRecordIdentity) {
        return [pscustomobject]@{
            Manifest = $universe
            Sha256 = Get-RSBytesSha256 -Bytes $bytes
            FrozenInputSha256ByRelativePath =
                $frozenInputSha256ByRelativePath
        }
    }
    return $universe
    } finally {
        Close-RSReadSnapshotLockSet -LockSet $readSnapshot
    }
}

function Read-RSCurrentCandidateUniverse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Read-RSCurrentCandidateUniverseCore -RepoRoot $RepoRoot
}

function Assert-RSCandidateSetUniverseClosure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Universe,

        [Parameter(Mandatory = $true)]
        [DateTimeOffset]$WindowOpenUtc,

        [Parameter(Mandatory = $true)]
        [DateTimeOffset]$GovernanceFreezeUtc
    )

    $setAuthorId = Get-RSRequiredString `
        -InputObject $CandidateSet `
        -Name authorId `
        -Context 'candidate set'
    $candidates = Get-RSRequiredArray `
        -InputObject $CandidateSet `
        -Name candidates `
        -Context 'candidate set'
    $candidateById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($candidate in $candidates) {
        $candidateById.Add(
            (Get-RSRequiredString -InputObject $candidate -Name candidateId),
            $candidate)
    }

    $outcomes = Get-RSRequiredArray `
        -InputObject $CandidateSet `
        -Name nominationOutcomes `
        -Context 'candidate set'
    $outcomeIds = [string[]]@(
        $outcomes | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name seedCandidateId
        })
    Assert-RSOrdinalSortedUnique `
        -Value $outcomeIds `
        -Context 'candidate-set nomination outcomes'
    $seedCandidates = Get-RSRequiredArray `
        -InputObject $Universe `
        -Name seedCandidates `
        -Context 'candidate universe'
    $seedIds = [string[]]@(
        $seedCandidates | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        })
    Assert-RSExactOrdinalSet `
        -Actual $outcomeIds `
        -Expected $seedIds `
        -Context 'candidate-set nomination closure'
    $outcomeById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($outcome in $outcomes) {
        $outcomeById.Add(
            (Get-RSRequiredString -InputObject $outcome -Name seedCandidateId),
            $outcome)
    }

    $policy = (Get-RSPropertyValue `
            -InputObject $Universe `
            -Name adaptationPolicy `
            -Required).Value
    $caps = (Get-RSPropertyValue `
            -InputObject $policy `
            -Name hardCaps `
            -Required).Value
    $taxonomy = [string[]](Get-RSRequiredArray `
            -InputObject $policy `
            -Name logicalConcernTaxonomy `
            -Context 'candidate-universe adaptation policy')
    $patchMetric = (Get-RSPropertyValue `
            -InputObject $policy `
            -Name patchMetricAlgorithm `
            -Required).Value
    $requiredApprovals = [int](Get-RSPropertyValue `
            -InputObject $patchMetric `
            -Name requiredIndependentApprovals `
            -Required).Value
    $expectedCandidateIds = [Collections.Generic.List[string]]::new()

    foreach ($seed in $seedCandidates) {
        $seedCandidateId = Get-RSRequiredString `
            -InputObject $seed `
            -Name candidateId
        $outcome = $outcomeById[$seedCandidateId]
        foreach ($name in @('pin', 'pinSha256', 'sourceArchiveSha256')) {
            if ((Get-RSRequiredString -InputObject $outcome -Name $name) -cne
                (Get-RSRequiredString -InputObject $seed -Name $name)) {
                throw [IO.InvalidDataException]::new(
                    "Nomination outcome '$seedCandidateId' $name differs from its seed.")
            }
        }
        $permission = Get-RSRequiredString `
            -InputObject $seed `
            -Name adaptationPermission
        $outcomeKind = Get-RSRequiredString `
            -InputObject $outcome `
            -Name outcome
        if ($permission -ceq 'forbidden-absent-subsystem') {
            if ($outcomeKind -cne 'source-rejected') {
                throw [IO.InvalidDataException]::new(
                    "Forbidden seed '$seedCandidateId' must be source-rejected.")
            }
            if (-not $candidateById.ContainsKey($seedCandidateId)) {
                throw [IO.InvalidDataException]::new(
                    "Source-rejected seed '$seedCandidateId' is absent from the candidate set.")
            }
            $candidate = $candidateById[$seedCandidateId]
            if ((Get-RSRequiredString -InputObject $candidate -Name disposition) -cne
                'source-rejected') {
                throw [IO.InvalidDataException]::new(
                    "Forbidden seed '$seedCandidateId' has a non-rejected candidate record.")
            }
            foreach ($name in @('pin', 'pinSha256', 'sourceArchiveSha256')) {
                if ((Get-RSRequiredString -InputObject $candidate -Name $name) -cne
                    (Get-RSRequiredString -InputObject $seed -Name $name)) {
                    throw [IO.InvalidDataException]::new(
                        "Source-rejected candidate '$seedCandidateId' $name differs from its seed.")
                }
            }
            $outcomeRejection = (Get-RSPropertyValue `
                    -InputObject $outcome `
                    -Name sourceRejection `
                    -Required).Value
            $candidateRejection = (Get-RSPropertyValue `
                    -InputObject $candidate `
                    -Name sourceRejection `
                    -Required).Value
            if ((Get-RSJcsSha256 -InputObject $outcomeRejection) -cne
                (Get-RSJcsSha256 -InputObject $candidateRejection)) {
                throw [IO.InvalidDataException]::new(
                    "Source rejection '$seedCandidateId' differs between nomination and candidate records.")
            }
            $expectedCandidateIds.Add($seedCandidateId)
            continue
        }

        if ($permission -cne 'allowed-existing-subsystem') {
            throw [IO.InvalidDataException]::new(
                "Seed '$seedCandidateId' has an unsupported adaptation permission.")
        }
        $nominationId = Get-RSRequiredString `
            -InputObject $outcome `
            -Name nominationId
        $allowedNominationId = Get-RSRequiredString `
            -InputObject $seed `
            -Name allowedNominationId
        if ($nominationId -cne $allowedNominationId) {
            throw [IO.InvalidDataException]::new(
                "Seed '$seedCandidateId' used an unfrozen nomination id.")
        }
        if ($outcomeKind -ceq 'admitted') {
            if (-not $candidateById.ContainsKey($nominationId)) {
                throw [IO.InvalidDataException]::new(
                    "Admitted nomination '$nominationId' is absent from candidates.")
            }
            $candidate = $candidateById[$nominationId]
            if ((Get-RSRequiredString -InputObject $candidate -Name disposition) -cne
                'collect') {
                throw [IO.InvalidDataException]::new(
                    "Admitted nomination '$nominationId' is not collectable.")
            }
            foreach ($name in @('pin', 'pinSha256', 'sourceArchiveSha256')) {
                if ((Get-RSRequiredString -InputObject $candidate -Name $name) -cne
                    (Get-RSRequiredString -InputObject $seed -Name $name)) {
                    throw [IO.InvalidDataException]::new(
                        "Admitted nomination '$nominationId' $name differs from its seed.")
                }
            }
            $expectedCandidateIds.Add($nominationId)
            continue
        }
        if ($outcomeKind -cne 'adaptation-rejected') {
            throw [IO.InvalidDataException]::new(
                "Allowed seed '$seedCandidateId' lacks an admitted or measured-rejected outcome.")
        }
        if ($candidateById.ContainsKey($nominationId)) {
            throw [IO.InvalidDataException]::new(
                "Adaptation-rejected nomination '$nominationId' cannot enter collection.")
        }

        $rejection = (Get-RSPropertyValue `
                -InputObject $outcome `
                -Name adaptationRejection `
                -Required).Value
        $patchRelativePath = Get-RSRequiredString `
            -InputObject $rejection `
            -Name patchRelativePath
        $repositoryRoot = [IO.Path]::GetFullPath($RepoRoot)
        $patchPath = [IO.Path]::GetFullPath((
                Join-Path $repositoryRoot $patchRelativePath))
        $patchRoot = [IO.Path]::GetFullPath((Join-Path $repositoryRoot (
                    'External\TerminalEngine\CandidatePatches')))
        if (-not $patchPath.StartsWith(
                $patchRoot + [IO.Path]::DirectorySeparatorChar,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' patch bytes are missing or differ.")
        }
        $patchBytes = Read-RSImmutableRepoFileBytes `
            -RepoRoot $repositoryRoot `
            -Path $patchPath
        if ((Get-RSBytesSha256 -Bytes $patchBytes) -cne
            (Get-RSRequiredString -InputObject $rejection -Name patchSha256) -or
            $patchBytes.Length -ne
            [int64](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name patchSizeBytes `
                -Required).Value) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' patch bytes are missing or differ.")
        }
        $changedLines = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name changedMaintainedLines `
                -Required).Value
        $changedFiles = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name changedMaintainedFiles `
                -Required).Value
        $binaryFiles = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name binaryFiles `
                -Required).Value
        $concerns = [string[]](Get-RSRequiredArray `
                -InputObject $rejection `
                -Name logicalConcerns `
                -Context "adaptation rejection '$nominationId'")
        Assert-RSOrdinalSortedUnique `
            -Value $concerns `
            -Context "adaptation rejection '$nominationId' concerns"
        foreach ($concern in $concerns) {
            if ($taxonomy -cnotcontains $concern) {
                throw [IO.InvalidDataException]::new(
                    "Adaptation rejection '$nominationId' uses unknown concern '$concern'.")
            }
        }
        [void](Read-RSCandidateConcernMapping `
                -RepoRoot $repositoryRoot `
                -PatchRecord $rejection `
                -CandidateId $nominationId)
        $newToolchains = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name newToolchainFamilies `
                -Required).Value
        $runtimeLeaves = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name newNonSystemRuntimeLeaves `
                -Required).Value
        $adapterClass = Get-RSRequiredString `
            -InputObject $rejection `
            -Name adapterClass
        $adapterFloor = if ($adapterClass -ceq
            'same-language-or-existing-c-abi') { 3 } else { 5 }
        $initialAdapterDays = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name initialAdapterDays `
                -Required).Value
        if ($initialAdapterDays -lt $adapterFloor) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' adapter estimate is below its floor.")
        }
        $declaredPatchDays = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name declaredInitialPatchDays `
                -Required).Value
        $patchFloor = [Math]::Max(
            $declaredPatchDays,
            [Math]::Max(
                2 * $concerns.Count,
                [int][Math]::Ceiling([decimal]$changedLines / [decimal]256)))
        $initialPatchDays = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name initialPatchDays `
                -Required).Value
        if ($initialPatchDays -ne $patchFloor) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' patch estimate is not the mechanical floor.")
        }
        $perUpgradeFloor = 1 +
            [int][Math]::Ceiling([decimal]$changedFiles / [decimal]8) +
            $newToolchains +
            $runtimeLeaves
        $perUpgradeDays = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name perUpgradeDays `
                -Required).Value
        if ($perUpgradeDays -lt $perUpgradeFloor) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' upgrade estimate is below its floor.")
        }
        $touchesSecuritySurface = Get-RSRequiredBoolean `
            -InputObject $rejection `
            -Name patchTouchesParserDecoderAccounting
        $securityDays = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name annualSecurityResponseDays `
                -Required).Value
        if ($securityDays -lt $(if ($touchesSecuritySurface) { 2 } else { 1 })) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' security estimate is below its floor.")
        }
        $noticeDays = [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name releaseNoticeDays `
                -Required).Value
        if ($noticeDays -lt 1) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' notice estimate is below its floor.")
        }
        $ownerDays = [int][Math]::Ceiling(
            ([decimal]$initialAdapterDays +
                [decimal]$initialPatchDays +
                [decimal]9 * [decimal]$perUpgradeDays +
                [decimal]3 * [decimal]$securityDays +
                [decimal]$noticeDays) *
            [decimal]1.20)
        if ($ownerDays -ne [int](Get-RSPropertyValue `
                -InputObject $rejection `
                -Name ownerDays `
                -Required).Value) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' owner-days value is not mechanical.")
        }

        $derivedFailures = [Collections.Generic.List[string]]::new()
        if ($ownerDays -gt [int]$caps.ownerDays) {
            $derivedFailures.Add('owner-days-cap')
        }
        if ($changedFiles -gt [int]$caps.changedMaintainedFiles) {
            $derivedFailures.Add('changed-files-cap')
        }
        if ($changedLines -gt [int]$caps.changedMaintainedLines) {
            $derivedFailures.Add('changed-lines-cap')
        }
        if ($concerns.Count -gt [int]$caps.logicalConcerns) {
            $derivedFailures.Add('logical-concerns-cap')
        }
        if ($newToolchains -gt [int]$caps.newToolchainFamilies) {
            $derivedFailures.Add('new-toolchain-cap')
        }
        if ($runtimeLeaves -gt [int]$caps.newNonSystemRuntimeLeaves) {
            $derivedFailures.Add('runtime-leaf-cap')
        }
        if ($binaryFiles -ne 0) {
            $derivedFailures.Add('binary-patch-forbidden')
        }
        $failureCodes = [string[]](Get-RSRequiredArray `
                -InputObject $rejection `
                -Name failureCodes `
                -Context "adaptation rejection '$nominationId'")
        Assert-RSOrdinalSortedUnique `
            -Value $failureCodes `
            -Context "adaptation rejection '$nominationId' failure codes"
        Assert-RSExactOrdinalSet `
            -Actual $failureCodes `
            -Expected $derivedFailures.ToArray() `
            -Context "adaptation rejection '$nominationId' mechanical failures"
        if ($derivedFailures.Count -eq 0) {
            throw [IO.InvalidDataException]::new(
                "Adaptation-rejected nomination '$nominationId' does not exceed a frozen cap.")
        }

        $glueAuthorIds = [string[]](Get-RSRequiredArray `
                -InputObject $rejection `
                -Name glueAuthorIds `
                -Context "adaptation rejection '$nominationId'")
        Assert-RSOrdinalSortedUnique `
            -Value $glueAuthorIds `
            -Context "adaptation rejection '$nominationId' glue authors"
        $evidencePayload = [ordered]@{}
        foreach ($property in $rejection.GetEnumerator()) {
            if ([string]$property.Key -cnotin @('approvals', 'evidenceSha256')) {
                $evidencePayload[[string]$property.Key] = $property.Value
            }
        }
        $evidenceSha256 = Get-RSJcsSha256 -InputObject ([ordered]@{
                seedCandidateId = $seedCandidateId
                nominationId = $nominationId
                pin = Get-RSRequiredString -InputObject $outcome -Name pin
                pinSha256 = Get-RSRequiredString -InputObject $outcome -Name pinSha256
                sourceArchiveSha256 = Get-RSRequiredString `
                    -InputObject $outcome `
                    -Name sourceArchiveSha256
                adaptationRejection = $evidencePayload
            })
        if ($evidenceSha256 -cne
            (Get-RSRequiredString -InputObject $rejection -Name evidenceSha256)) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' evidence digest is invalid.")
        }
        $approvals = Get-RSRequiredArray `
            -InputObject $rejection `
            -Name approvals `
            -Context "adaptation rejection '$nominationId'"
        if ($approvals.Count -ne $requiredApprovals) {
            throw [IO.InvalidDataException]::new(
                "Adaptation rejection '$nominationId' lacks exactly $requiredApprovals approvals.")
        }
        Assert-RSApprovalInsideGovernanceWindow `
            -Approval $approvals `
            -WindowOpenUtc $WindowOpenUtc `
            -GovernanceFreezeUtc $GovernanceFreezeUtc `
            -Context "adaptation rejection '$nominationId' approval"
        $approvalIds = [string[]]@(
            $approvals | ForEach-Object {
                Get-RSRequiredString -InputObject $_ -Name approvalId
            })
        Assert-RSOrdinalSortedUnique `
            -Value $approvalIds `
            -Context "adaptation rejection '$nominationId' approvals"
        $reviewerIds = [string[]]@()
        foreach ($approval in $approvals) {
            if ((Get-RSRequiredString `
                    -InputObject $approval `
                    -Name evidenceSha256) -cne $evidenceSha256) {
                throw [IO.InvalidDataException]::new(
                    "Adaptation rejection '$nominationId' approval is bound to another record.")
            }
            $reviewerId = Get-RSRequiredString `
                -InputObject $approval `
                -Name reviewerId
            if ($reviewerId -ceq $setAuthorId -or
                $glueAuthorIds -ccontains $reviewerId) {
                throw [IO.InvalidDataException]::new(
                    "Adaptation rejection '$nominationId' was approved by an author.")
            }
            $reviewerIds += $reviewerId
        }
        $reviewerIds = Get-RSOrdinalSortedStrings -Value $reviewerIds
        Assert-RSOrdinalSortedUnique `
            -Value $reviewerIds `
            -Context "adaptation rejection '$nominationId' reviewers"
    }

    $controls = Get-RSRequiredArray `
        -InputObject $Universe `
        -Name controls `
        -Context 'candidate universe'
    foreach ($control in $controls) {
        $controlId = Get-RSRequiredString -InputObject $control -Name candidateId
        if (-not $candidateById.ContainsKey($controlId)) {
            throw [IO.InvalidDataException]::new(
                "Frozen control '$controlId' is absent from the candidate set.")
        }
        $candidate = $candidateById[$controlId]
        foreach ($name in @('pin', 'pinSha256', 'constraintId')) {
            if ((Get-RSRequiredString -InputObject $candidate -Name $name) -cne
                (Get-RSRequiredString -InputObject $control -Name $name)) {
                throw [IO.InvalidDataException]::new(
                    "Control '$controlId' $name differs from the universe.")
            }
        }
        $expectedCandidateIds.Add($controlId)
    }
    Assert-RSExactOrdinalSet `
        -Actual ([string[]]@($candidateById.Keys)) `
        -Expected $expectedCandidateIds.ToArray() `
        -Context 'candidate-set admitted/rejected/control closure'
}

function Read-RSCandidateSetPreflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [switch]$IncludeRecordIdentity
    )

    $path = Join-Path $RepoRoot $script:CandidateSetLeaf
    $schemaPath = Join-Path $RepoRoot $script:CandidateSetSchemaLeaf
    if (-not [IO.File]::Exists($path) -or -not [IO.File]::Exists($schemaPath)) {
        throw [IO.InvalidDataException]::new(
            'The frozen candidate set or its closed schema has not been checked in.')
    }
    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -eq 0 -or
        ($bytes.Length -ge 3 -and $bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf)) {
        throw [IO.InvalidDataException]::new('The candidate set must be nonempty UTF-8 without BOM.')
    }
    try {
        $text = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $candidateSet = $text |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    } catch {
        throw [IO.InvalidDataException]::new('The candidate set is not valid duplicate-free UTF-8 JSON.', $_.Exception)
    }
    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $canonicalCandidateSetBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $candidateSet
    if ([Convert]::ToBase64String($bytes) -cne
        [Convert]::ToBase64String($canonicalCandidateSetBytes)) {
        throw [IO.InvalidDataException]::new(
            'The candidate set must be exact RFC 8785 JCS bytes; ambiguous or noncanonical JSON is forbidden.')
    }
    Test-RSJsonSchema `
        -Json $text `
        -SchemaPath $schemaPath `
        -Context 'Frozen candidate set'

    $formatId = Get-RSRequiredString -InputObject $candidateSet -Name formatId -Context 'candidate set'
    if ($formatId -cne 'red-salamander-terminal-engine-candidate-set') {
        throw [IO.InvalidDataException]::new('The candidate set formatId is not supported.')
    }
    $versionProperty = Get-RSPropertyValue -InputObject $candidateSet -Name version -Required
    if ($versionProperty.Value -ne 1) {
        throw [IO.InvalidDataException]::new('The candidate set version must be 1.')
    }
    $universe = Read-RSCurrentCandidateUniverse -RepoRoot $RepoRoot
    $windowOpenUtc = Get-RSUniverseNominationWindowOpenUtc `
        -Universe $universe `
        -Context 'current candidate universe'
    $governanceFreezeUtc = Get-RSGovernanceFreezeUtc `
        -Record $candidateSet `
        -WindowOpenUtc $windowOpenUtc `
        -Context 'candidate set'
    $candidateUniverseSha256 = Get-RSRequiredString `
        -InputObject $candidateSet `
        -Name candidateUniverseSha256 `
        -Pattern '^[0-9a-f]{64}$' `
        -Context 'candidate set'
    if ($candidateUniverseSha256 -cne
        (Get-RSFileSha256 -Path (
                Get-RSCurrentCandidateUniverseFilePath -RepoRoot $RepoRoot))) {
        throw [IO.InvalidDataException]::new(
            'The candidate set is not bound to the frozen candidate universe.')
    }
    $candidates = Get-RSRequiredArray -InputObject $candidateSet -Name candidates -Context 'candidate set'
    if ($candidates.Count -eq 0) {
        throw [IO.InvalidDataException]::new('The candidate set must declare at least one candidate.')
    }
    $candidateReviewSha256 = Get-RSRequiredString `
        -InputObject $candidateSet `
        -Name candidateReviewSha256 `
        -Pattern '^[0-9a-f]{64}$' `
        -Context 'candidate set'
    $candidateReviewPath = Join-Path $RepoRoot $script:CandidateReviewLeaf
    if (-not [IO.File]::Exists($candidateReviewPath) -or
        $candidateReviewSha256 -cne (Get-RSFileSha256 -Path $candidateReviewPath)) {
        throw [IO.InvalidDataException]::new(
            'The candidate set candidate-review digest does not match the frozen review.')
    }
    $candidateAssessmentSha256 = Get-RSRequiredString `
        -InputObject $candidateSet `
        -Name candidateAssessmentSha256 `
        -Pattern '^[0-9a-f]{64}$' `
        -Context 'candidate set'
    $candidateAssessmentPath = Join-Path $RepoRoot $script:CandidateAssessmentLeaf
    if (-not [IO.File]::Exists($candidateAssessmentPath) -or
        $candidateAssessmentSha256 -cne
        (Get-RSFileSha256 -Path $candidateAssessmentPath)) {
        throw [IO.InvalidDataException]::new(
            'The candidate set candidate-assessment digest does not match the frozen assessment.')
    }

    $ids = [string[]]@()
    foreach ($candidate in $candidates) {
        $candidateId = Get-RSRequiredString `
            -InputObject $candidate `
            -Name candidateId `
            -Pattern '^[a-z0-9][a-z0-9._-]{0,63}$' `
            -Context 'candidate-set candidate'
        $null = Get-RSRequiredString -InputObject $candidate -Name pin -Context "candidate '$candidateId'"
        $null = Get-RSRequiredString `
            -InputObject $candidate `
            -Name pinSha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context "candidate '$candidateId'"
        $kind = Get-RSRequiredString `
            -InputObject $candidate `
            -Name scoredOrControl `
            -Context "candidate '$candidateId'"
        if ($kind -cne 'scored' -and $kind -cne 'control') {
            throw [IO.InvalidDataException]::new(
                "Candidate '$candidateId' scoredOrControl must be 'scored' or 'control'.")
        }
        if ($kind -ceq 'control') {
            $constraintId = Get-RSRequiredString `
                -InputObject $candidate `
                -Name constraintId `
                -Context "control '$candidateId'"
            if ($constraintId -cnotin @(
                    'candidate-owned-control',
                    'web-runtime-second-stack',
                    'new-parser-maintenance')) {
                throw [IO.InvalidDataException]::new(
                    "Control '$candidateId' has an unsupported constraintId.")
            }
        } else {
            $null = Get-RSRequiredString `
                -InputObject $candidate `
                -Name sourceArchiveSha256 `
                -Pattern '^[0-9a-f]{64}$' `
                -Context "candidate '$candidateId'"
            $disposition = Get-RSRequiredString `
                -InputObject $candidate `
                -Name disposition `
                -Context "candidate '$candidateId'"
            if ($disposition -cne 'collect' -and $disposition -cne 'source-rejected') {
                throw [IO.InvalidDataException]::new(
                    "Candidate '$candidateId' has an unsupported disposition.")
            }
            if ($disposition -ceq 'source-rejected') {
                $sourceRejection = (Get-RSPropertyValue `
                        -InputObject $candidate `
                        -Name sourceRejection `
                        -Required).Value
                foreach ($name in @(
                        'gateId',
                        'failureCode',
                        'candidateReviewSha256',
                        'blockerId',
                        'blockerSha256',
                        'sourceArchiveSha256')) {
                    $null = Get-RSRequiredString `
                        -InputObject $sourceRejection `
                        -Name $name `
                        -Context "candidate '$candidateId' source rejection"
                }
                if ((Get-RSRequiredString `
                            -InputObject $sourceRejection `
                            -Name candidateReviewSha256) -cne $candidateReviewSha256) {
                    throw [IO.InvalidDataException]::new(
                        "Candidate '$candidateId' source rejection is not bound to the frozen review.")
                }
            }
        }
        $ids += $candidateId
    }
    Assert-RSOrdinalSortedUnique -Value $ids -Context 'candidate-set candidates'
    Assert-RSCandidateSetUniverseClosure `
        -RepoRoot $RepoRoot `
        -CandidateSet $candidateSet `
        -Universe $universe `
        -WindowOpenUtc $windowOpenUtc `
        -GovernanceFreezeUtc $governanceFreezeUtc

    $bootstrapInputs = Get-RSRequiredArray `
        -InputObject $candidateSet `
        -Name bootstrapInputSets `
        -Context 'candidate set'
    if ($bootstrapInputs.Count -ne 2) {
        throw [IO.InvalidDataException]::new(
            'The candidate set must declare exactly x64 and ARM64 bootstrap input-set locks.')
    }
    $bootstrapPlatforms = [string[]]@()
    foreach ($inputSet in $bootstrapInputs) {
        $platform = Get-RSRequiredString -InputObject $inputSet -Name platform -Context 'bootstrap input set'
        if ($platform -cne 'x64' -and $platform -cne 'ARM64') {
            throw [IO.InvalidDataException]::new('Bootstrap input-set platform must be x64 or ARM64.')
        }
        $declaredInputSetSha256 = Get-RSRequiredString `
            -InputObject $inputSet `
            -Name sha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context "bootstrap input set '$platform'"
        $inputs = Get-RSRequiredArray `
            -InputObject $inputSet `
            -Name inputs `
            -Context "bootstrap input set '$platform'"
        $inputIds = [string[]]@(
            $inputs | ForEach-Object {
                Get-RSRequiredString `
                    -InputObject $_ `
                    -Name inputId `
                    -Context "bootstrap input set '$platform'"
            })
        Assert-RSOrdinalSortedUnique `
            -Value $inputIds `
            -Context "bootstrap input set '$platform' inputs"
        foreach ($input in $inputs) {
            $inputId = Get-RSRequiredString -InputObject $input -Name inputId
            $consumers = [string[]](Get-RSRequiredArray `
                    -InputObject $input `
                    -Name consumers `
                    -Context "bootstrap input '$inputId'")
            Assert-RSOrdinalSortedUnique `
                -Value $consumers `
                -Context "bootstrap input '$inputId' consumers"
            foreach ($consumer in $consumers) {
                if ($ids -cnotcontains $consumer) {
                    throw [IO.InvalidDataException]::new(
                        "Bootstrap input '$inputId' names unknown consumer '$consumer'.")
                }
            }
        }
        if ((Get-RSJcsSha256 -InputObject $inputs) -cne
            $declaredInputSetSha256) {
            throw [IO.InvalidDataException]::new(
                "Bootstrap input set '$platform' digest is not derived from its exact inputs.")
        }

        foreach ($candidate in $candidates) {
            if ((Get-RSRequiredString `
                    -InputObject $candidate `
                    -Name scoredOrControl) -cne 'scored') {
                continue
            }
            $candidateId = Get-RSRequiredString `
                -InputObject $candidate `
                -Name candidateId
            $sourceInputs = @(
                $inputs | Where-Object {
                    (Get-RSRequiredString `
                        -InputObject $_ `
                        -Name inputKind) -ceq 'source-archive' -and
                    [string[]](Get-RSRequiredArray `
                        -InputObject $_ `
                        -Name consumers `
                        -Context "bootstrap input set '$platform'") -ccontains
                        $candidateId
                })
            if ($sourceInputs.Count -ne 1) {
                throw [IO.InvalidDataException]::new(
                    "Scored candidate '$candidateId' must have exactly one source archive in '$platform'.")
            }
            $sourceInput = $sourceInputs[0]
            if ((Get-RSRequiredString -InputObject $sourceInput -Name version) -cne
                    (Get-RSRequiredString -InputObject $candidate -Name pin) -or
                (Get-RSRequiredString -InputObject $sourceInput -Name sha256) -cne
                    (Get-RSRequiredString `
                        -InputObject $candidate `
                        -Name sourceArchiveSha256)) {
                throw [IO.InvalidDataException]::new(
                    "Scored candidate '$candidateId' source bootstrap identity differs in '$platform'.")
            }
        }
        $bootstrapPlatforms += $platform
    }
    Assert-RSExactOrdinalSet `
        -Actual $bootstrapPlatforms `
        -Expected ([string[]]@('ARM64', 'x64')) `
        -Context 'bootstrap input-set platforms'

    if ($IncludeRecordIdentity) {
        return [pscustomobject]@{
            Manifest = $candidateSet
            Sha256 = Get-RSBytesSha256 -Bytes $bytes
        }
    }
    return $candidateSet
}

function Read-RSCandidateSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return (Read-RSCandidateGovernanceClosure -RepoRoot $RepoRoot).CandidateSet
}

function Get-RSCandidateSetSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet
    )

    return Get-RSJcsSha256 -InputObject $CandidateSet
}

function Read-RSCandidateReview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet,

        [switch]$VerifySourceArchives
    )

    $path = Join-Path $RepoRoot $script:CandidateReviewLeaf
    $schemaPath = Join-Path $RepoRoot $script:CandidateReviewSchemaLeaf
    if (-not [IO.File]::Exists($path) -or -not [IO.File]::Exists($schemaPath)) {
        throw [IO.InvalidDataException]::new(
            'The frozen candidate review or its schema is missing.')
    }
    $universe = Read-RSCurrentCandidateUniverse -RepoRoot $RepoRoot
    if ($VerifySourceArchives -and [int]$universe['version'] -ge 4) {
        throw [IO.InvalidDataException]::new(
            'Round 4 does not authorize released-handle source-archive verification; a future universe must freeze one complete archive-and-verifier snapshot before executable collection.')
    }
    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -eq 0 -or
        ($bytes.Length -ge 3 -and
            $bytes[0] -eq 0xef -and
            $bytes[1] -eq 0xbb -and
            $bytes[2] -eq 0xbf)) {
        throw [IO.InvalidDataException]::new(
            'The candidate review must be nonempty UTF-8 without BOM.')
    }
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $review = $json |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    } catch {
        throw [IO.InvalidDataException]::new(
            'The candidate review is not duplicate-free UTF-8 JSON.',
            $_.Exception)
    }
    [void](Import-RSExactModule -Path (Join-Path $RepoRoot (
                'Tools\TerminalEvidence.psm1')))
    $canonicalReviewBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $review
    if ([Convert]::ToBase64String($bytes) -cne
        [Convert]::ToBase64String($canonicalReviewBytes)) {
        throw [IO.InvalidDataException]::new(
            'The candidate review must be exact RFC 8785 JCS bytes; ambiguous or noncanonical JSON is forbidden.')
    }
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath $schemaPath `
        -Context 'Frozen candidate review'

    $actualSha256 = Get-RSBytesSha256 -Bytes $bytes
    $expectedSha256 = Get-RSRequiredString `
        -InputObject $CandidateSet `
        -Name candidateReviewSha256 `
        -Pattern '^[0-9a-f]{64}$' `
        -Context 'candidate set'
    if ($actualSha256 -cne $expectedSha256) {
        throw [IO.InvalidDataException]::new(
            'The frozen candidate review digest differs from the candidate set.')
    }

    $windowOpenUtc = Get-RSUniverseNominationWindowOpenUtc `
        -Universe $universe `
        -Context 'current candidate universe'
    $reviewFreezeUtc = Get-RSGovernanceFreezeUtc `
        -Record $review `
        -WindowOpenUtc $windowOpenUtc `
        -Context 'candidate review'
    $setFreezeUtc = Get-RSGovernanceFreezeUtc `
        -Record $CandidateSet `
        -WindowOpenUtc $windowOpenUtc `
        -Context 'candidate set'
    if ((Get-RSRequiredString `
                -InputObject $review `
                -Name frozenUtc `
                -Context 'candidate review') -cne
        (Get-RSRequiredString `
                -InputObject $CandidateSet `
                -Name frozenUtc `
                -Context 'candidate set')) {
        throw [IO.InvalidDataException]::new(
            'Candidate review and candidate set do not share one exact governance freeze.')
    }
    $approvals = Get-RSRequiredArray `
        -InputObject $review `
        -Name approvals `
        -Context 'candidate review'
    Assert-RSApprovalInsideGovernanceWindow `
        -Approval $approvals `
        -WindowOpenUtc $windowOpenUtc `
        -GovernanceFreezeUtc $reviewFreezeUtc `
        -Context 'candidate-review approval'

    $authorId = Get-RSRequiredString `
        -InputObject $review `
        -Name authorId `
        -Pattern '^[a-z0-9][a-z0-9._-]{0,63}$' `
        -Context 'candidate review'
    $reviewCandidates = Get-RSRequiredArray `
        -InputObject $review `
        -Name candidates `
        -Context 'candidate review'
    $reviewIds = [string[]]@(
        $reviewCandidates | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        })
    Assert-RSOrdinalSortedUnique `
        -Value $reviewIds `
        -Context 'candidate-review candidates'

    $scoredCandidates = @(
        (Get-RSCandidateSetEntries -CandidateSet $CandidateSet) |
            Where-Object {
                (Get-RSRequiredString -InputObject $_ -Name scoredOrControl) -ceq 'scored'
            })
    $scoredIds = [string[]]@(
        $scoredCandidates | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        })
    Assert-RSExactOrdinalSet `
        -Actual $reviewIds `
        -Expected $scoredIds `
        -Context 'candidate-review scored candidates'

    $frozenById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($candidate in $scoredCandidates) {
        $frozenById.Add(
            (Get-RSRequiredString -InputObject $candidate -Name candidateId),
            $candidate)
    }
    $blockerBySha = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    $blockerShaByCandidateAndId = [Collections.Generic.Dictionary[string, string]]::new(
        [StringComparer]::Ordinal)
    $glueAuthorsByBlockerSha = [Collections.Generic.Dictionary[string, string[]]]::new(
        [StringComparer]::Ordinal)
    $sourceProofs = [Collections.Generic.List[object]]::new()
    foreach ($reviewCandidate in $reviewCandidates) {
        $candidateId = Get-RSRequiredString -InputObject $reviewCandidate -Name candidateId
        $frozenCandidate = $frozenById[$candidateId]
        foreach ($name in @(
                'pin',
                'pinSha256',
                'sourceArchiveSha256',
                'disposition')) {
            if ((Get-RSRequiredString -InputObject $reviewCandidate -Name $name) -cne
                (Get-RSRequiredString -InputObject $frozenCandidate -Name $name)) {
                throw [IO.InvalidDataException]::new(
                    "Candidate-review '$candidateId' $name differs from the candidate set.")
            }
        }
        $blockers = Get-RSRequiredArray `
            -InputObject $reviewCandidate `
            -Name blockers `
            -Context "candidate-review '$candidateId'"
        $blockerIds = [string[]]@(
            $blockers | ForEach-Object {
                Get-RSRequiredString -InputObject $_ -Name blockerId
            })
        if ($blockerIds.Count -ne 0) {
            Assert-RSOrdinalSortedUnique `
                -Value $blockerIds `
                -Context "candidate-review '$candidateId' blockers"
        }
        $glueAuthorIds = [string[]](Get-RSRequiredArray `
                -InputObject $reviewCandidate `
                -Name glueAuthorIds `
                -Context "candidate-review '$candidateId'")
        if ($glueAuthorIds.Count -ne 0) {
            Assert-RSOrdinalSortedUnique `
                -Value $glueAuthorIds `
                -Context "candidate-review '$candidateId' glue authors"
        }
        $archivePath = $null
        if ($VerifySourceArchives -and $blockers.Count -ne 0) {
            $archiveRelativePath = Get-RSRequiredString `
                -InputObject $reviewCandidate `
                -Name archiveRelativePath `
                -Context "candidate-review '$candidateId'"
            $archiveRoot = [IO.Path]::GetFullPath((
                    Join-Path $RepoRoot (
                        '.build\terminal-engine-spike\bootstrap\codeload')))
            $archivePath = [IO.Path]::GetFullPath((
                    Join-Path $RepoRoot $archiveRelativePath))
            if (-not $archivePath.StartsWith(
                    $archiveRoot + [IO.Path]::DirectorySeparatorChar,
                    [StringComparison]::OrdinalIgnoreCase) -or
                -not [IO.File]::Exists($archivePath) -or
                ((Get-Item -LiteralPath $archivePath -Force).Attributes -band
                    [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw [IO.InvalidDataException]::new(
                    "Candidate-review '$candidateId' source archive is missing or unsafe.")
            }
        }
        foreach ($blocker in $blockers) {
            $source = (Get-RSPropertyValue `
                    -InputObject $blocker `
                    -Name source `
                    -Required).Value
            $lineStart = [int](Get-RSPropertyValue `
                    -InputObject $source `
                    -Name lineStart `
                    -Required).Value
            $lineEnd = [int](Get-RSPropertyValue `
                    -InputObject $source `
                    -Name lineEnd `
                    -Required).Value
            if ($lineEnd -lt $lineStart) {
                throw [IO.InvalidDataException]::new(
                    "Candidate-review '$candidateId' has a reversed source line range.")
            }
            $blockerSha256 = Get-RSJcsSha256 -InputObject ([ordered]@{
                    candidateId = $candidateId
                    pin = Get-RSRequiredString -InputObject $reviewCandidate -Name pin
                    pinSha256 = Get-RSRequiredString -InputObject $reviewCandidate -Name pinSha256
                    sourceArchiveSha256 = Get-RSRequiredString `
                        -InputObject $reviewCandidate `
                        -Name sourceArchiveSha256
                    blocker = $blocker
                })
            if ($blockerBySha.ContainsKey($blockerSha256)) {
                throw [IO.InvalidDataException]::new(
                    "Candidate review contains duplicate blocker identity '$blockerSha256'.")
            }
            $blockerBySha.Add($blockerSha256, $blocker)
            $glueAuthorsByBlockerSha.Add($blockerSha256, $glueAuthorIds)
            $blockerShaByCandidateAndId.Add(
                "$candidateId`0$(Get-RSRequiredString -InputObject $blocker -Name blockerId)",
                $blockerSha256)
            if ($VerifySourceArchives) {
                [void](Import-RSExactModule -Path (
                        Join-Path $RepoRoot (
                            'Tools\TerminalEngine\TerminalEngineSourceReview.psm1')))
                $proof = Test-RSSourceReviewBlocker `
                    -ArchivePath $archivePath `
                    -ExpectedArchiveSha256 (Get-RSRequiredString `
                        -InputObject $reviewCandidate `
                        -Name sourceArchiveSha256) `
                    -Blocker $blocker
                $sourceProofs.Add([ordered]@{
                        candidateId = $candidateId
                        blockerId = Get-RSRequiredString `
                            -InputObject $blocker `
                            -Name blockerId
                        blockerSha256 = $blockerSha256
                        proofSha256 = Get-RSJcsSha256 -InputObject $proof
                    })
            }
        }
    }

    $approvalIds = [string[]]@(
        $approvals | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name approvalId
        })
    Assert-RSOrdinalSortedUnique `
        -Value $approvalIds `
        -Context 'candidate-review approvals'
    $reviewersByBlocker = @{}
    foreach ($approval in $approvals) {
        $blockerSha256 = Get-RSRequiredString -InputObject $approval -Name blockerSha256
        if (-not $blockerBySha.ContainsKey($blockerSha256)) {
            throw [IO.InvalidDataException]::new(
                "Candidate-review approval refers to unknown blocker '$blockerSha256'.")
        }
        $reviewerId = Get-RSRequiredString -InputObject $approval -Name reviewerId
        if ($reviewerId -ceq $authorId) {
            throw [IO.InvalidDataException]::new(
                "Candidate-review author '$authorId' cannot approve a source blocker.")
        }
        if ($glueAuthorsByBlockerSha[$blockerSha256] -ccontains $reviewerId) {
            throw [IO.InvalidDataException]::new(
                "Candidate-glue author '$reviewerId' cannot approve blocker '$blockerSha256'.")
        }
        if (-not $reviewersByBlocker.ContainsKey($blockerSha256)) {
            $reviewersByBlocker[$blockerSha256] =
                [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        }
        if (-not $reviewersByBlocker[$blockerSha256].Add($reviewerId)) {
            throw [IO.InvalidDataException]::new(
                "Reviewer '$reviewerId' approved blocker '$blockerSha256' more than once.")
        }
    }

    foreach ($frozenCandidate in $scoredCandidates) {
        $candidateId = Get-RSRequiredString -InputObject $frozenCandidate -Name candidateId
        if ((Get-RSRequiredString -InputObject $frozenCandidate -Name disposition) -cne
            'source-rejected') {
            continue
        }
        $sourceRejection = (Get-RSPropertyValue `
                -InputObject $frozenCandidate `
                -Name sourceRejection `
                -Required).Value
        $blockerId = Get-RSRequiredString -InputObject $sourceRejection -Name blockerId
        $key = "$candidateId`0$blockerId"
        if (-not $blockerShaByCandidateAndId.ContainsKey($key)) {
            throw [IO.InvalidDataException]::new(
                "Frozen source rejection '$candidateId/$blockerId' is absent from the review.")
        }
        $computedSha256 = $blockerShaByCandidateAndId[$key]
        if ($computedSha256 -cne
            (Get-RSRequiredString -InputObject $sourceRejection -Name blockerSha256)) {
            throw [IO.InvalidDataException]::new(
                "Frozen source rejection '$candidateId/$blockerId' has a wrong blocker digest.")
        }
        $blocker = $blockerBySha[$computedSha256]
        foreach ($pair in @(
                @('gateId', 'gateId'),
                @('failureCode', 'failureCode'))) {
            if ((Get-RSRequiredString -InputObject $sourceRejection -Name $pair[0]) -cne
                (Get-RSRequiredString -InputObject $blocker -Name $pair[1])) {
                throw [IO.InvalidDataException]::new(
                    "Frozen source rejection '$candidateId/$blockerId' differs from its blocker.")
            }
        }
        $source = (Get-RSPropertyValue -InputObject $blocker -Name source -Required).Value
        if ((Get-RSRequiredString -InputObject $sourceRejection -Name sourceArchiveSha256) -cne
            (Get-RSRequiredString -InputObject $frozenCandidate -Name sourceArchiveSha256)) {
            throw [IO.InvalidDataException]::new(
                "Frozen source rejection '$candidateId/$blockerId' has a wrong archive digest.")
        }
        $requiredApprovals = [int](Get-RSPropertyValue `
                -InputObject (Get-RSPropertyValue `
                    -InputObject (Get-RSRequirements -RepoRoot $RepoRoot) `
                    -Name sourceReviewPolicy `
                    -Required).Value `
                -Name requiredIndependentApprovals `
                -Required).Value
        if (-not $reviewersByBlocker.ContainsKey($computedSha256) -or
            $reviewersByBlocker[$computedSha256].Count -ne $requiredApprovals) {
            throw [IO.InvalidDataException]::new(
                "Frozen source rejection '$candidateId/$blockerId' lacks exactly $requiredApprovals independent approvals.")
        }
    }

    return [pscustomobject]@{
        Manifest = $review
        Sha256 = $actualSha256
        BlockerBySha256 = $blockerBySha
        SourceProofs = $sourceProofs.ToArray()
    }
}

function Read-RSCandidateAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet
    )

    $path = Join-Path $RepoRoot $script:CandidateAssessmentLeaf
    $schemaPath = Join-Path $RepoRoot $script:CandidateAssessmentSchemaLeaf
    if (-not [IO.File]::Exists($path) -or -not [IO.File]::Exists($schemaPath)) {
        throw [IO.InvalidDataException]::new(
            'The frozen candidate assessment or its closed schema is missing.')
    }
    $universe = Read-RSCurrentCandidateUniverse -RepoRoot $RepoRoot
    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -eq 0 -or
        ($bytes.Length -ge 3 -and
            $bytes[0] -eq 0xef -and
            $bytes[1] -eq 0xbb -and
            $bytes[2] -eq 0xbf)) {
        throw [IO.InvalidDataException]::new(
            'The candidate assessment must be nonempty UTF-8 without BOM.')
    }
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $assessment = $json |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    }
    catch {
        throw [IO.InvalidDataException]::new(
            'The candidate assessment is not duplicate-free UTF-8 JSON.',
            $_.Exception)
    }

    [void](Import-RSExactModule -Path (Join-Path $RepoRoot (
                'Tools\TerminalEvidence.psm1')))
    if ([Convert]::ToBase64String($bytes) -cne
        [Convert]::ToBase64String(
            (ConvertTo-RSJcsUtf8Bytes -InputObject $assessment))) {
        throw [IO.InvalidDataException]::new(
            'The candidate assessment must be exact RFC 8785 JCS bytes.')
    }
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath $schemaPath `
        -Context 'Frozen candidate assessment'

    $actualSha256 = Get-RSBytesSha256 -Bytes $bytes
    if ($actualSha256 -cne
        (Get-RSRequiredString `
            -InputObject $CandidateSet `
            -Name candidateAssessmentSha256 `
            -Pattern '^[0-9a-f]{64}$' `
            -Context 'candidate set')) {
        throw [IO.InvalidDataException]::new(
            'The frozen candidate assessment digest differs from the candidate set.')
    }

    $windowOpenUtc = Get-RSUniverseNominationWindowOpenUtc `
        -Universe $universe `
        -Context 'current candidate universe'
    $assessmentFreezeUtc = Get-RSGovernanceFreezeUtc `
        -Record $assessment `
        -WindowOpenUtc $windowOpenUtc `
        -Context 'candidate assessment'
    $setFreezeUtc = Get-RSGovernanceFreezeUtc `
        -Record $CandidateSet `
        -WindowOpenUtc $windowOpenUtc `
        -Context 'candidate set'
    if ((Get-RSRequiredString `
                -InputObject $assessment `
                -Name frozenUtc `
                -Context 'candidate assessment') -cne
        (Get-RSRequiredString `
                -InputObject $CandidateSet `
                -Name frozenUtc `
                -Context 'candidate set')) {
        throw [IO.InvalidDataException]::new(
            'Candidate assessment and candidate set do not share one exact governance freeze.')
    }
    if ((Get-RSRequiredString `
            -InputObject $assessment `
            -Name candidateUniverseSha256) -cne
        (Get-RSFileSha256 -Path (
            Get-RSCurrentCandidateUniverseFilePath -RepoRoot $RepoRoot))) {
        throw [IO.InvalidDataException]::new(
            'The candidate assessment is not bound to the frozen universe.')
    }
    if ((Get-RSRequiredString `
            -InputObject $assessment `
            -Name scoreAnchorsSha256) -cne
        (Get-RSFileSha256 -Path (
            Join-Path $RepoRoot $script:ScoreAnchorsLeaf))) {
        throw [IO.InvalidDataException]::new(
            'The candidate assessment is not bound to the frozen score anchors.')
    }

    $assessmentAuthorId = Get-RSRequiredString `
        -InputObject $assessment `
        -Name authorId `
        -Context 'candidate assessment'
    $candidates = Get-RSRequiredArray `
        -InputObject $assessment `
        -Name candidates `
        -Context 'candidate assessment'
    foreach ($candidate in $candidates) {
        $candidateId = Get-RSRequiredString `
            -InputObject $candidate `
            -Name candidateId `
            -Context 'candidate assessment'
        Assert-RSApprovalInsideGovernanceWindow `
            -Approval (Get-RSRequiredArray `
                -InputObject $candidate `
                -Name approvals `
                -Context "candidate assessment '$candidateId'") `
            -WindowOpenUtc $windowOpenUtc `
            -GovernanceFreezeUtc $assessmentFreezeUtc `
            -Context "candidate assessment '$candidateId' approval"
    }
    $candidateIds = [string[]]@(
        $candidates | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        })
    Assert-RSOrdinalSortedUnique `
        -Value $candidateIds `
        -Context 'candidate-assessment candidates'

    $collectCandidates = @(
        (Get-RSCandidateSetEntries -CandidateSet $CandidateSet) |
            Where-Object {
                (Get-RSRequiredString -InputObject $_ -Name scoredOrControl) -ceq
                    'scored' -and
                (Get-RSRequiredString -InputObject $_ -Name disposition) -ceq
                    'collect'
            })
    $collectIds = [string[]]@(
        $collectCandidates | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        })
    Assert-RSExactOrdinalSet `
        -Actual $candidateIds `
        -Expected $collectIds `
        -Context 'candidate-assessment collect candidates'

    $frozenById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($candidate in $collectCandidates) {
        $frozenById.Add(
            (Get-RSRequiredString -InputObject $candidate -Name candidateId),
            $candidate)
    }

    [void](Import-RSExactModule -Path (
            Join-Path $RepoRoot (
                'Tools\TerminalEngine\TerminalEngineCandidateAssessment.psm1')))
    $requiredApprovals = [int](Get-RSPropertyValue `
            -InputObject (Get-RSPropertyValue `
                -InputObject $universe `
                -Name adaptationPolicy `
                -Required).Value `
            -Name patchMetricAlgorithm `
            -Required).Value.requiredIndependentApprovals
    $candidateById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    $admissionById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    $staticRatingsById = [Collections.Generic.Dictionary[string, object[]]]::new(
        [StringComparer]::Ordinal)
    foreach ($candidate in $candidates) {
        $candidateId = Get-RSRequiredString -InputObject $candidate -Name candidateId
        $frozenCandidate = $frozenById[$candidateId]
        foreach ($name in @('pin', 'pinSha256', 'sourceArchiveSha256')) {
            if ((Get-RSRequiredString -InputObject $candidate -Name $name) -cne
                (Get-RSRequiredString -InputObject $frozenCandidate -Name $name)) {
                throw [IO.InvalidDataException]::new(
                    "Candidate assessment '$candidateId' $name differs from the candidate set.")
            }
        }

        $ledgerSha256 = Get-RSCandidateLedgerSha256 -Candidate $candidate
        if ($ledgerSha256 -cne
            (Get-RSRequiredString -InputObject $candidate -Name ledgerSha256)) {
            throw [IO.InvalidDataException]::new(
                "Candidate assessment '$candidateId' ledger digest is invalid.")
        }
        $rawLedger = (Get-RSPropertyValue `
                -InputObject $candidate `
                -Name rawLedger `
                -Required).Value
        [void](Assert-RSAssessmentDescriptorBuildTruth `
                -Candidate $candidate)
        $patch = (Get-RSPropertyValue `
                -InputObject $rawLedger `
                -Name patch `
                -Required).Value
        $patchRelativePath = Get-RSRequiredString `
            -InputObject $patch `
            -Name patchRelativePath
        $patchPath = [IO.Path]::GetFullPath((Join-Path $RepoRoot $patchRelativePath))
        $patchRoot = [IO.Path]::GetFullPath((Join-Path $RepoRoot (
                    'External\TerminalEngine\CandidatePatches')))
        if (-not $patchPath.StartsWith(
                $patchRoot + [IO.Path]::DirectorySeparatorChar,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw [IO.InvalidDataException]::new(
                "Candidate assessment '$candidateId' patch file is missing or unsafe.")
        }
        $patchBytes = Read-RSImmutableRepoFileBytes `
            -RepoRoot $RepoRoot `
            -Path $patchPath
        if ((Get-RSBytesSha256 -Bytes $patchBytes) -cne
            (Get-RSRequiredString -InputObject $patch -Name patchSha256) -or
            $patchBytes.Length -ne
            [int64](Get-RSPropertyValue `
                -InputObject $patch `
                -Name patchSizeBytes `
                -Required).Value) {
            throw [IO.InvalidDataException]::new(
                "Candidate assessment '$candidateId' patch bytes differ from the raw ledger.")
        }
        [void](Read-RSCandidateConcernMapping `
                -RepoRoot $RepoRoot `
                -PatchRecord $patch `
                -CandidateId $candidateId)
        $nominationMatches = @(
            (Get-RSRequiredArray `
                -InputObject $CandidateSet `
                -Name nominationOutcomes `
                -Context 'candidate set') |
                Where-Object {
                    (Get-RSRequiredString -InputObject $_ -Name outcome) -ceq
                        'admitted' -and
                    (Get-RSRequiredString -InputObject $_ -Name nominationId) -ceq
                        $candidateId
                })
        if ($nominationMatches.Count -ne 1 -or
            (Get-RSRequiredString `
                -InputObject $nominationMatches[0] `
                -Name ledgerSha256) -cne $ledgerSha256) {
            throw [IO.InvalidDataException]::new(
                "Candidate assessment '$candidateId' is not bound to its admitted nomination.")
        }
        $glueAuthorIds = [string[]](Get-RSRequiredArray `
                -InputObject $candidate `
                -Name glueAuthorIds `
                -Context "candidate assessment '$candidateId'")
        Assert-RSOrdinalSortedUnique `
            -Value $glueAuthorIds `
            -Context "candidate assessment '$candidateId' glue authors"
        $approvals = Get-RSRequiredArray `
            -InputObject $candidate `
            -Name approvals `
            -Context "candidate assessment '$candidateId'"
        if ($approvals.Count -ne $requiredApprovals) {
            throw [IO.InvalidDataException]::new(
                "Candidate assessment '$candidateId' requires exactly $requiredApprovals approvals.")
        }
        $approvalIds = [string[]]@(
            $approvals | ForEach-Object {
                Get-RSRequiredString -InputObject $_ -Name approvalId
            })
        Assert-RSOrdinalSortedUnique `
            -Value $approvalIds `
            -Context "candidate assessment '$candidateId' approvals"
        $reviewerIds = [string[]]@()
        foreach ($approval in $approvals) {
            if ((Get-RSRequiredString -InputObject $approval -Name ledgerSha256) -cne
                $ledgerSha256) {
                throw [IO.InvalidDataException]::new(
                    "Candidate assessment '$candidateId' approval is bound to another ledger.")
            }
            $reviewerId = Get-RSRequiredString `
                -InputObject $approval `
                -Name reviewerId
            if ($reviewerId -ceq $assessmentAuthorId -or
                $glueAuthorIds -ccontains $reviewerId) {
                throw [IO.InvalidDataException]::new(
                    "Candidate assessment '$candidateId' was approved by an author.")
            }
            $reviewerIds += $reviewerId
        }
        $reviewerIds = Get-RSOrdinalSortedStrings -Value $reviewerIds
        Assert-RSOrdinalSortedUnique `
            -Value $reviewerIds `
            -Context "candidate assessment '$candidateId' reviewers"

        $admission = Get-RSCandidateAdmission `
            -Candidate $candidate `
            -CandidateUniverse $universe
        if (-not $admission.admitted) {
            throw [IO.InvalidDataException]::new(
                "Collect candidate '$candidateId' exceeds the frozen adaptation caps.")
        }
        $staticRatings = [object[]]@(
            Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $universe)
        $candidateById.Add($candidateId, $candidate)
        $admissionById.Add($candidateId, $admission)
        $staticRatingsById.Add($candidateId, $staticRatings)
    }

    return [pscustomobject]@{
        Manifest = $assessment
        Sha256 = $actualSha256
        CandidateById = $candidateById
        AdmissionById = $admissionById
        StaticRatingsById = $staticRatingsById
    }
}

function Read-RSCandidateGovernanceClosure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $universePath = [IO.Path]::GetFullPath(
        (Get-RSCurrentCandidateUniverseFilePath -RepoRoot $RepoRoot))
    $candidateSetPath = [IO.Path]::GetFullPath(
        (Join-Path $RepoRoot $script:CandidateSetLeaf))
    $candidateReviewPath = [IO.Path]::GetFullPath(
        (Join-Path $RepoRoot $script:CandidateReviewLeaf))
    $candidateAssessmentPath = [IO.Path]::GetFullPath(
        (Join-Path $RepoRoot $script:CandidateAssessmentLeaf))
    $recordPaths = [ordered]@{
        Universe = $universePath
        CandidateSet = $candidateSetPath
        CandidateReview = $candidateReviewPath
        CandidateAssessment = $candidateAssessmentPath
    }
    $readSnapshot = New-RSReadSnapshotLockSet `
        -RepoRoot $RepoRoot `
        -Path (Get-RSRound4SnapshotPaths `
            -CurrentUniversePath $universePath `
            -IncludeLiveGovernance)
    try {
        $referencedPaths = Get-RSGovernanceReferencedSnapshotPaths `
            -LockSet $readSnapshot `
            -CandidateSetPath $candidateSetPath `
            -CandidateAssessmentPath $candidateAssessmentPath
        Add-RSReadSnapshotLocks `
            -LockSet $readSnapshot `
            -Path $referencedPaths

        $startingDigests = [ordered]@{}
        foreach ($entry in $recordPaths.GetEnumerator()) {
            $startingDigests[$entry.Key] = Get-RSFileSha256 -Path $entry.Value
        }

        $universeResult = Read-RSCurrentCandidateUniverseCore `
            -RepoRoot $RepoRoot `
            -IncludeRecordIdentity
        $null = Import-RSAuthenticatedCollectionNeutralityModules `
            -RepoRoot $RepoRoot
        [void](Import-RSExactModule -Path (Join-Path $RepoRoot (
                    'Tools\TerminalEngine\TerminalEngineCandidateAssessment.psm1')))
        $universe = $universeResult.Manifest
        $candidateSetResult = Read-RSCandidateSetPreflight `
            -RepoRoot $RepoRoot `
            -IncludeRecordIdentity
        $candidateSet = $candidateSetResult.Manifest
        $review = Read-RSCandidateReview `
            -RepoRoot $RepoRoot `
            -CandidateSet $candidateSet
        $assessment = Read-RSCandidateAssessment `
            -RepoRoot $RepoRoot `
            -CandidateSet $candidateSet

        $reviewFrozenUtc = Get-RSRequiredString `
            -InputObject $review.Manifest `
            -Name frozenUtc `
            -Context 'candidate review'
        $assessmentFrozenUtc = Get-RSRequiredString `
            -InputObject $assessment.Manifest `
            -Name frozenUtc `
            -Context 'candidate assessment'
        $setFrozenUtc = Get-RSRequiredString `
            -InputObject $candidateSet `
            -Name frozenUtc `
            -Context 'candidate set'
        if ($reviewFrozenUtc -cne $assessmentFrozenUtc -or
            $reviewFrozenUtc -cne $setFrozenUtc) {
            throw [IO.InvalidDataException]::new(
                'CandidateReview, CandidateAssessment, and CandidateSet must share one exact governance freeze.')
        }
        $windowOpenUtc = Get-RSUniverseNominationWindowOpenUtc `
            -Universe $universe `
            -Context 'current candidate universe'
        $null = Get-RSGovernanceFreezeUtc `
            -Record $candidateSet `
            -WindowOpenUtc $windowOpenUtc `
            -Context 'candidate governance closure'

        $endingUniversePath = [IO.Path]::GetFullPath(
            (Get-RSCurrentCandidateUniverseFilePath -RepoRoot $RepoRoot))
        if ($endingUniversePath -cne $universePath) {
            throw [IO.InvalidDataException]::new(
                'The current candidate-universe path changed while its governance closure was read.')
        }
        $endingDigests = [ordered]@{}
        foreach ($entry in $recordPaths.GetEnumerator()) {
            $endingDigests[$entry.Key] = Get-RSFileSha256 -Path $entry.Value
        }
        $parsedDigests = [ordered]@{
            Universe = $universeResult.Sha256
            CandidateSet = $candidateSetResult.Sha256
            CandidateReview = $review.Sha256
            CandidateAssessment = $assessment.Sha256
        }
        foreach ($name in $startingDigests.Keys) {
            if ([string]$startingDigests[$name] -cne
                    [string]$endingDigests[$name]) {
                throw [IO.InvalidDataException]::new(
                    "Candidate governance record '$name' changed while its closure was read.")
            }
            if ([string]$parsedDigests[$name] -cne
                    [string]$startingDigests[$name]) {
                throw [IO.InvalidDataException]::new(
                    "Candidate governance record '$name' parsed bytes differ from its locked identity.")
            }
        }
        $neutralityPolicyRelativePath =
            'Specs/Terminal/TerminalEngineCollectionNeutrality.v1.json'
        $frozenInputSha256ByRelativePath =
            $universeResult.FrozenInputSha256ByRelativePath
        if ($null -eq $frozenInputSha256ByRelativePath -or
            -not $frozenInputSha256ByRelativePath.ContainsKey(
                $neutralityPolicyRelativePath)) {
            throw [IO.InvalidDataException]::new(
                'The authenticated universe omitted the collection-neutrality policy identity.')
        }
        $neutralityPolicySha256 = [string](
            $frozenInputSha256ByRelativePath[$neutralityPolicyRelativePath])
        if ($neutralityPolicySha256 -cnotmatch '^[0-9a-f]{64}$') {
            throw [IO.InvalidDataException]::new(
                'The authenticated collection-neutrality policy identity is invalid.')
        }

        return [pscustomobject]@{
            Universe = $universe
            CandidateSet = $candidateSet
            Review = $review.Manifest
            Assessment = $assessment.Manifest
            ReviewResult = $review
            AssessmentResult = $assessment
            UniverseSha256 = [string]$startingDigests.Universe
            CandidateSetSha256 = [string]$startingDigests.CandidateSet
            CandidateReviewSha256 = [string]$startingDigests.CandidateReview
            CandidateAssessmentSha256 = [string]$startingDigests.CandidateAssessment
            NeutralityPolicySha256 = $neutralityPolicySha256
            UniverseOpenUtc = $windowOpenUtc.ToString(
                "yyyy-MM-dd'T'HH:mm:ss'Z'",
                [Globalization.CultureInfo]::InvariantCulture)
            GovernanceFreezeUtc = $setFrozenUtc
        }
    } finally {
        Close-RSReadSnapshotLockSet -LockSet $readSnapshot
    }
}

function Invoke-RSAuthenticatedCollectionNeutralityProof {
    [CmdletBinding()]
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

    $expectedIdentities = [ordered]@{
        NeutralityPolicy = $ExpectedNeutralityPolicySha256
        Universe = $ExpectedUniverseSha256
        CandidateSet = $ExpectedCandidateSetSha256
        CandidateReview = $ExpectedCandidateReviewSha256
        CandidateAssessment = $ExpectedCandidateAssessmentSha256
    }
    foreach ($entry in $expectedIdentities.GetEnumerator()) {
        if ([string]$entry.Value -cnotmatch '^[0-9a-f]{64}$') {
            throw [IO.InvalidDataException]::new(
                "Expected $($entry.Key) identity is invalid.")
        }
    }
    $universePath = [IO.Path]::GetFullPath(
        (Get-RSCurrentCandidateUniverseFilePath -RepoRoot $RepoRoot))
    $candidateSetPath = [IO.Path]::GetFullPath(
        (Join-Path $RepoRoot $script:CandidateSetLeaf))
    $candidateAssessmentPath = [IO.Path]::GetFullPath(
        (Join-Path $RepoRoot $script:CandidateAssessmentLeaf))
    $snapshot = New-RSReadSnapshotLockSet `
        -RepoRoot $RepoRoot `
        -Path (Get-RSRound4SnapshotPaths `
            -CurrentUniversePath $universePath `
            -IncludeLiveGovernance)
    try {
        $referencedPaths = Get-RSGovernanceReferencedSnapshotPaths `
            -LockSet $snapshot `
            -CandidateSetPath $candidateSetPath `
            -CandidateAssessmentPath $candidateAssessmentPath
        Add-RSReadSnapshotLocks `
            -LockSet $snapshot `
            -Path $referencedPaths

        $closure = Read-RSCandidateGovernanceClosure -RepoRoot $RepoRoot
        if ([string]$closure.NeutralityPolicySha256 -cne
            $ExpectedNeutralityPolicySha256) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality policy differs from the authenticated provider context.')
        }
        $actualGovernanceIdentities = [ordered]@{
            Universe = [string]$closure.UniverseSha256
            CandidateSet = [string]$closure.CandidateSetSha256
            CandidateReview = [string]$closure.CandidateReviewSha256
            CandidateAssessment =
                [string]$closure.CandidateAssessmentSha256
        }
        foreach ($entry in $actualGovernanceIdentities.GetEnumerator()) {
            if ([string]$entry.Value -cne
                [string]$expectedIdentities[$entry.Key]) {
                throw [IO.InvalidDataException]::new(
                    "$($entry.Key) identity differs from the authenticated provider context.")
            }
        }
        $neutralityModule =
            Import-RSAuthenticatedCollectionNeutralityModules `
                -RepoRoot $RepoRoot
        $policyPath = Join-Path $RepoRoot (
            'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json')
        $policyBytes = Read-RSReadSnapshotBytes `
            -LockSet $snapshot `
            -Path $policyPath
        if ((Get-RSBytesSha256 -Bytes $policyBytes) -cne
            $ExpectedNeutralityPolicySha256) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality policy bytes differ from their authenticated identity.')
        }
        $policy = ConvertFrom-RSReadSnapshotJson `
            -LockSet $snapshot `
            -Path $policyPath `
            -Context 'Collection-neutrality policy'
        $canonicalPolicyBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $policy
        if (-not [Linq.Enumerable]::SequenceEqual(
                [byte[]]$policyBytes,
                [byte[]]$canonicalPolicyBytes)) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality policy must be exact RFC 8785 JCS bytes.')
        }
        Test-RSJsonSchema `
            -Json ([Text.UTF8Encoding]::new(
                    $false,
                    $true).GetString($policyBytes)) `
            -SchemaPath (Join-Path $RepoRoot (
                    'Specs\Terminal\TerminalEngineCollectionNeutrality.schema.json')) `
            -Context 'Collection-neutrality policy'

        return & $neutralityModule {
            param(
                $root,
                $policyRecord,
                $governance,
                $descriptors,
                $frozenInputs)
            $baseline = Invoke-RSTerminalEngineCollectionNeutralityScan `
                -RepoRoot $root `
                -PolicyRecord $policyRecord `
                -ValidateBaselineProof
            $runtime = Invoke-RSTerminalEngineCollectionNeutralityScan `
                -RepoRoot $root `
                -GovernanceObject $governance `
                -Descriptor $descriptors `
                -FrozenInput $frozenInputs `
                -PolicyRecord $policyRecord
            return [pscustomobject]@{
                Baseline = $baseline
                Runtime = $runtime
            }
        } `
            $RepoRoot `
            $policy `
            @($closure.Universe, $closure.CandidateSet) `
            $Descriptor `
            $FrozenInput
    } finally {
        Close-RSReadSnapshotLockSet -LockSet $snapshot
    }
}

function Get-RSBootstrapInputSetSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet,

        [Parameter(Mandatory = $true)]
        [string]$Platform
    )

    $inputSets = Get-RSRequiredArray `
        -InputObject $CandidateSet `
        -Name bootstrapInputSets `
        -Context 'candidate set'
    $matching = @(
        $inputSets | Where-Object {
            (Get-RSRequiredString -InputObject $_ -Name platform -Context 'bootstrap input set') -ceq $Platform
        }
    )
    if ($matching.Count -ne 1) {
        throw [IO.InvalidDataException]::new(
            "The candidate set must contain one bootstrap lock for $Platform.")
    }
    return Get-RSRequiredString `
        -InputObject $matching[0] `
        -Name sha256 `
        -Pattern '^[0-9a-f]{64}$' `
        -Context "bootstrap input set '$Platform'"
}

function Get-RSCandidateSetEntries {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet
    )

    return Get-RSRequiredArray -InputObject $CandidateSet -Name candidates -Context 'candidate set'
}

function Assert-RSExecutableCandidateCohort {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet,

        [string[]]$CandidateId = @()
    )

    $collectCandidates = @(
        (Get-RSCandidateSetEntries -CandidateSet $CandidateSet) |
            Where-Object {
                ($CandidateId.Count -eq 0 -or
                    $CandidateId -ccontains
                        (Get-RSRequiredString `
                            -InputObject $_ `
                            -Name candidateId `
                            -Context 'candidate-set candidate')) -and
                (Get-RSRequiredString `
                    -InputObject $_ `
                    -Name scoredOrControl `
                    -Context 'candidate-set candidate') -ceq 'scored' -and
                (Get-RSRequiredString `
                    -InputObject $_ `
                    -Name disposition `
                    -Context 'scored candidate') -ceq 'collect'
            }
    )
    if ($collectCandidates.Count -eq 0) {
        $scope = if ($CandidateId.Count -eq 0) {
            'frozen candidate set'
        } else {
            'selected candidate cohort'
        }
        throw [IO.InvalidDataException]::new(
            "The $scope must contain at least one scored candidate with disposition 'collect'; static-only evidence cannot enter collection or finalization.")
    }
}

function Resolve-RSCandidateSelection {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Candidate,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet
    )

    $allCandidates = Get-RSCandidateSetEntries -CandidateSet $CandidateSet
    $allIds = [string[]]@(
        $allCandidates | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        }
    )
    if ($Candidate.Count -eq 1 -and $Candidate[0] -ceq 'All') {
        return ,$allIds
    }
    if ($Candidate -ccontains 'All') {
        throw [ArgumentException]::new("Candidate 'All' cannot be combined with explicit IDs.")
    }
    foreach ($candidateId in $Candidate) {
        if ($candidateId -cnotmatch '^[a-z0-9][a-z0-9._-]{0,63}$') {
            throw [ArgumentException]::new("Candidate '$candidateId' is not a valid exact candidate ID.")
        }
    }
    $sorted = Get-RSOrdinalSortedStrings -Value $Candidate
    Assert-RSOrdinalSortedUnique -Value $sorted -Context 'Candidate parameter'
    foreach ($candidateId in $sorted) {
        if ($allIds -cnotcontains $candidateId) {
            throw [ArgumentException]::new("Candidate '$candidateId' is not in the frozen candidate set.")
        }
    }
    return ,$sorted
}

function Test-RSMapContainsKey {
    param(
        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$Map,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -ne $Map.PSObject.Methods['ContainsKey']) {
        return [bool]$Map.ContainsKey($Name)
    }
    return [bool]$Map.Contains($Name)
}

function Assert-RSTerminalEngineInvocation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$BoundParameters
    )

    foreach ($requiredName in @('Mode', 'Candidate', 'EvidenceRoot')) {
        if (-not (Test-RSMapContainsKey -Map $BoundParameters -Name $requiredName)) {
            throw [ArgumentException]::new("-$requiredName is required.")
        }
    }
    $mode = [string]$BoundParameters['Mode']
    if ($mode -cne 'Collect' -and $mode -cne 'Finalize') {
        throw [ArgumentException]::new("-Mode must be exactly Collect or Finalize.")
    }
    $candidate = [string[]]@($BoundParameters['Candidate'])
    if ($candidate.Count -eq 0 -or @($candidate | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count -ne 0) {
        throw [ArgumentException]::new('-Candidate must contain at least one nonempty value.')
    }
    if ([string]::IsNullOrWhiteSpace([string]$BoundParameters['EvidenceRoot'])) {
        throw [ArgumentException]::new('-EvidenceRoot must be a nonempty path.')
    }

    if ($mode -ceq 'Collect') {
        if ($candidate.Count -ne 1 -or $candidate[0] -cne 'All') {
            throw [ArgumentException]::new(
                'Collect requires -Candidate All; partial collection manifests are not finalizable.')
        }
        foreach ($requiredName in @('Platform', 'Configuration', 'Bootstrap', 'Clean', 'Offline')) {
            if (-not (Test-RSMapContainsKey -Map $BoundParameters -Name $requiredName)) {
                throw [ArgumentException]::new("Collect requires explicit -$requiredName.")
            }
        }
        foreach ($switchName in @('Bootstrap', 'Clean', 'Offline')) {
            if (-not [bool]$BoundParameters[$switchName]) {
                throw [ArgumentException]::new("Collect requires -$switchName to be true.")
            }
        }
        foreach ($forbiddenName in @('EvidenceManifest', 'RequireAllEvidence')) {
            if (Test-RSMapContainsKey -Map $BoundParameters -Name $forbiddenName) {
                throw [ArgumentException]::new("-$forbiddenName is Finalize-only.")
            }
        }
        $platform = [string]$BoundParameters['Platform']
        $configuration = [string]$BoundParameters['Configuration']
        if ($platform -cne 'x64' -and $platform -cne 'ARM64') {
            throw [ArgumentException]::new('-Platform must be exactly x64 or ARM64.')
        }
        if ($configuration -cne 'Release' -and $configuration -cne 'ASan Debug') {
            throw [ArgumentException]::new('-Configuration must be exactly Release or ASan Debug.')
        }
        if ($platform -ceq 'ARM64' -and $configuration -ceq 'ASan Debug') {
            throw [ArgumentException]::new('Gate 0 has no ARM64 ASan Debug collection lane.')
        }
    } else {
        foreach ($forbiddenName in @('Platform', 'Configuration', 'Bootstrap', 'Clean', 'Offline')) {
            if (Test-RSMapContainsKey -Map $BoundParameters -Name $forbiddenName) {
                throw [ArgumentException]::new("-$forbiddenName is Collect-only.")
            }
        }
        if (-not (Test-RSMapContainsKey -Map $BoundParameters -Name 'EvidenceManifest')) {
            throw [ArgumentException]::new('Finalize requires exactly three explicit -EvidenceManifest paths.')
        }
        $manifests = [string[]]@($BoundParameters['EvidenceManifest'])
        if ($manifests.Count -ne 3 -or
            @($manifests | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count -ne 0) {
            throw [ArgumentException]::new('Finalize requires exactly three nonempty -EvidenceManifest paths.')
        }
        if ($candidate.Count -ne 1 -or $candidate[0] -cne 'All') {
            throw [ArgumentException]::new(
                "Finalize requires -Candidate All so the frozen scored/control set cannot be narrowed.")
        }
    }

    return [pscustomobject]@{
        Mode = $mode
        Candidate = $candidate
        EvidenceRoot = [string]$BoundParameters['EvidenceRoot']
    }
}

function Test-RSJsonSchema {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Json,

        [Parameter(Mandatory = $true)]
        [string]$SchemaPath,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $valid = Test-Json `
        -Json $Json `
        -SchemaFile $SchemaPath `
        -ErrorAction SilentlyContinue `
        -WarningAction SilentlyContinue 2>$null
    if (-not $valid) {
        throw [IO.InvalidDataException]::new("$Context does not satisfy TerminalEngine evidence schema v1.")
    }
}

function Assert-RSGateSemantics {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Requirements
    )

    $candidateId = Get-RSRequiredString -InputObject $Candidate -Name candidateId
    $outcome = Get-RSRequiredString -InputObject $Candidate -Name outcome -Context "candidate '$candidateId'"
    if ($outcome -ceq 'constraint-control' -or $outcome -ceq 'source-rejected') {
        return
    }

    $gates = Get-RSRequiredArray -InputObject $Candidate -Name gates -Context "candidate '$candidateId'"
    if ($gates.Count -ne 12) {
        throw [IO.InvalidDataException]::new("Candidate '$candidateId' must contain G1 through G12.")
    }
    for ($index = 0; $index -lt 12; ++$index) {
        $actualGateId = Get-RSRequiredString -InputObject $gates[$index] -Name gateId
        if ($actualGateId -cne $script:GateIds[$index]) {
            throw [IO.InvalidDataException]::new("Candidate '$candidateId' gate order is invalid.")
        }
    }

    if ($outcome -ceq 'complete') {
        foreach ($gate in $gates) {
            if ((Get-RSRequiredString -InputObject $gate -Name status) -cne 'pass' -or
                (Get-RSRequiredString -InputObject $gate -Name resultCategory) -cne 'passed') {
                throw [IO.InvalidDataException]::new(
                    "Complete candidate '$candidateId' contains a non-pass gate.")
            }
            $null = Get-RSRequiredString `
                -InputObject $gate `
                -Name evidenceSha256 `
                -Pattern '^[0-9a-f]{64}$'
        }
        return
    }
    if ($outcome -cne 'early-failure') {
        throw [IO.InvalidDataException]::new("Candidate '$candidateId' has an unsupported outcome.")
    }

    $earlyFailureProperty = Get-RSPropertyValue -InputObject $Candidate -Name earlyFailure -Required
    $firstFailedGate = Get-RSRequiredString `
        -InputObject $earlyFailureProperty.Value `
        -Name firstFailedGate `
        -Context "candidate '$candidateId' early failure"
    $failedIndex = [Array]::IndexOf($script:GateIds, $firstFailedGate)
    if ($failedIndex -lt 0) {
        throw [IO.InvalidDataException]::new("Candidate '$candidateId' has an invalid first failed gate.")
    }
    for ($index = 0; $index -lt 12; ++$index) {
        $status = Get-RSRequiredString -InputObject $gates[$index] -Name status
        $category = Get-RSRequiredString -InputObject $gates[$index] -Name resultCategory
        $evidenceProperty = Get-RSPropertyValue -InputObject $gates[$index] -Name evidenceSha256 -Required
        if ($index -lt $failedIndex) {
            if ($status -cne 'pass' -or
                $category -cne 'passed' -or
                $evidenceProperty.Value -isnot [string] -or
                [string]$evidenceProperty.Value -cnotmatch '^[0-9a-f]{64}$') {
                throw [IO.InvalidDataException]::new(
                    "Candidate '$candidateId' gate transition before $firstFailedGate is invalid.")
            }
        } elseif ($index -eq $failedIndex) {
            if ($status -cne 'fail' -or $category -cne 'mandatory-failure' -or
                $evidenceProperty.Value -isnot [string] -or
                [string]$evidenceProperty.Value -cnotmatch '^[0-9a-f]{64}$') {
                throw [IO.InvalidDataException]::new(
                    "Candidate '$candidateId' failed-gate record is invalid.")
            }
        } else {
            if ($status -cne 'not-run' -or
                $category -cne "failed-$firstFailedGate" -or
                $null -ne $evidenceProperty.Value) {
                throw [IO.InvalidDataException]::new(
                    "Candidate '$candidateId' post-failure gate transition is invalid.")
            }
        }
    }

    $failureCode = Get-RSRequiredString `
        -InputObject $earlyFailureProperty.Value `
        -Name failureCode `
        -Context "candidate '$candidateId' early failure"
    $requirementGates = Get-RSRequiredArray -InputObject $Requirements -Name gates -Context 'requirements'
    $gateRequirements = @(
        $requirementGates | Where-Object {
            (Get-RSRequiredString -InputObject $_ -Name gateId) -ceq $firstFailedGate
        }
    )
    if ($gateRequirements.Count -ne 1) {
        throw [IO.InvalidDataException]::new("Requirements do not define exactly one $firstFailedGate.")
    }
    $allowedCodes = [string[]](Get-RSRequiredArray `
            -InputObject $gateRequirements[0] `
            -Name failureCodes `
            -Context $firstFailedGate)
    if ($allowedCodes -cnotcontains $failureCode) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' failure code '$failureCode' is not allowed for $firstFailedGate.")
    }
}

function Assert-RSSourceRejectionSemantics {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Candidate,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$FrozenCandidate,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Requirements
    )

    $candidateId = Get-RSRequiredString -InputObject $Candidate -Name candidateId
    $disposition = Get-RSRequiredString `
        -InputObject $FrozenCandidate `
        -Name disposition `
        -Context "candidate '$candidateId'"
    if ($disposition -cne 'source-rejected') {
        throw [IO.InvalidDataException]::new(
            "Collected source rejection for '$candidateId' was not frozen before collection.")
    }

    $actual = (Get-RSPropertyValue `
            -InputObject $Candidate `
            -Name sourceRejection `
            -Required).Value
    $expected = (Get-RSPropertyValue `
            -InputObject $FrozenCandidate `
            -Name sourceRejection `
            -Required).Value
    foreach ($name in @(
            'gateId',
            'failureCode',
            'candidateReviewSha256',
            'blockerId',
            'blockerSha256',
            'sourceArchiveSha256')) {
        if ((Get-RSRequiredString -InputObject $actual -Name $name) -cne
            (Get-RSRequiredString -InputObject $expected -Name $name)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$candidateId' source rejection '$name' differs from the frozen review.")
        }
    }

    $gateId = Get-RSRequiredString -InputObject $actual -Name gateId
    $sourcePolicy = (Get-RSPropertyValue `
            -InputObject $Requirements `
            -Name sourceReviewPolicy `
            -Required).Value
    $eligibleGates = [string[]](Get-RSRequiredArray `
            -InputObject $sourcePolicy `
            -Name eligibleGates `
            -Context 'source review policy')
    if ($eligibleGates -cnotcontains $gateId) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' uses source rejection for ineligible gate '$gateId'.")
    }
    $gateRequirements = @(
        (Get-RSRequiredArray -InputObject $Requirements -Name gates -Context 'requirements') |
            Where-Object {
                (Get-RSRequiredString -InputObject $_ -Name gateId) -ceq $gateId
            })
    if ($gateRequirements.Count -ne 1) {
        throw [IO.InvalidDataException]::new(
            "Requirements do not define exactly one source-review gate '$gateId'.")
    }
    $allowedCodes = [string[]](Get-RSRequiredArray `
            -InputObject $gateRequirements[0] `
            -Name failureCodes `
            -Context $gateId)
    $failureCode = Get-RSRequiredString -InputObject $actual -Name failureCode
    if ($allowedCodes -cnotcontains $failureCode) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' source failure '$failureCode' is not allowed for $gateId.")
    }
}

function Assert-RSCollectionSemantics {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Manifest,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Requirements,

        [string[]]$ExpectedCandidateId = @()
    )

    if ((Get-RSRequiredString -InputObject $Manifest -Name recordKind) -cne 'collection') {
        throw [IO.InvalidDataException]::new('Input record is not a collection manifest.')
    }
    Assert-RSExecutableCandidateCohort -CandidateSet $CandidateSet
    $platform = Get-RSRequiredString -InputObject $Manifest -Name platform
    $configuration = Get-RSRequiredString -InputObject $Manifest -Name configuration
    if ("$platform|$configuration" -cnotin $script:RequiredLaneKeys) {
        throw [IO.InvalidDataException]::new(
            "Collection lane '$platform|$configuration' is not a required Gate-0 lane.")
    }

    $nativeProperty = Get-RSPropertyValue -InputObject $Manifest -Name nativeAttestation -Required
    $processMachine = Get-RSRequiredString -InputObject $nativeProperty.Value -Name processMachine
    $nativeMachine = Get-RSRequiredString -InputObject $nativeProperty.Value -Name nativeMachine
    $isNative = Get-RSRequiredBoolean -InputObject $nativeProperty.Value -Name isNative
    $expectedNativeMachine = if ($platform -ceq 'x64') { 'AMD64' } else { 'ARM64' }
    if ($processMachine -cne 'UNKNOWN' -or $nativeMachine -cne $expectedNativeMachine -or -not $isNative) {
        throw [IO.InvalidDataException]::new(
            "Collection lane '$platform|$configuration' is not native $expectedNativeMachine evidence.")
    }
    $windowsBuild = Get-RSRequiredString `
        -InputObject $nativeProperty.Value `
        -Name windowsBuild `
        -Pattern '^10\.0\.[0-9]{5}\.[0-9]+$'
    $versionParts = $windowsBuild.Split('.')
    $buildNumber = [int]::Parse(
        $versionParts[2],
        [Globalization.CultureInfo]::InvariantCulture)
    $revisionNumber = [int]::Parse(
        $versionParts[3],
        [Globalization.CultureInfo]::InvariantCulture)
    if ($buildNumber -lt 22000) {
        throw [IO.InvalidDataException]::new(
            "Collection lane '$platform|$configuration' ran below the supported Windows minimum.")
    }
    if ((Get-RSRequiredString `
                -InputObject $nativeProperty.Value `
                -Name minimumSupportedBuild) -cne '22000.2600') {
        throw [IO.InvalidDataException]::new(
            'Native attestation does not bind the frozen Windows minimum 22000.2600.')
    }
    $minimumOsProof = Get-RSRequiredString `
        -InputObject $nativeProperty.Value `
        -Name minimumOsProof
    if ($minimumOsProof -cne 'native-smoke' -and
        $minimumOsProof -cne 'pe-import-audit') {
        throw [IO.InvalidDataException]::new(
            'Native attestation has an unsupported minimum-OS proof method.')
    }
    if ($minimumOsProof -ceq 'native-smoke' -and
        ($buildNumber -ne 22000 -or $revisionNumber -lt 2600)) {
        throw [IO.InvalidDataException]::new(
            'A native-smoke minimum-OS proof must run on build 22000 revision 2600 or newer.')
    }
    $null = Get-RSRequiredString `
        -InputObject $nativeProperty.Value `
        -Name minimumOsEvidenceSha256 `
        -Pattern '^[0-9a-f]{64}$'

    $policyProperty = Get-RSPropertyValue -InputObject $Manifest -Name collectionPolicy -Required
    foreach ($flag in @('bootstrap', 'clean', 'offline')) {
        if (-not (Get-RSRequiredBoolean -InputObject $policyProperty.Value -Name $flag)) {
            throw [IO.InvalidDataException]::new("Collection policy '$flag' must be true.")
        }
    }
    if ((Get-RSRequiredString -InputObject $policyProperty.Value -Name networkIsolation) -cne
        'runner-enforced') {
        throw [IO.InvalidDataException]::new('Collection network isolation is not runner-enforced.')
    }
    $actualInputSet = Get-RSRequiredString `
        -InputObject $policyProperty.Value `
        -Name bootstrapInputSetSha256 `
        -Pattern '^[0-9a-f]{64}$'
    $expectedInputSet = Get-RSBootstrapInputSetSha256 `
        -CandidateSet $CandidateSet `
        -Platform $platform
    if ($actualInputSet -cne $expectedInputSet) {
        throw [IO.InvalidDataException]::new(
            "Collection bootstrap input-set digest does not match the frozen $platform lock.")
    }

    $frozenProperty = Get-RSPropertyValue -InputObject $Manifest -Name frozen -Required
    $candidateSetSha256 = Get-RSRequiredString `
        -InputObject $frozenProperty.Value `
        -Name candidateSetSha256 `
        -Pattern '^[0-9a-f]{64}$'
    if ($candidateSetSha256 -cne (Get-RSCandidateSetSha256 -CandidateSet $CandidateSet)) {
        throw [IO.InvalidDataException]::new('Collection candidate-set digest is not current.')
    }

    $frozenEntries = Get-RSCandidateSetEntries -CandidateSet $CandidateSet
    $frozenById = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    foreach ($entry in $frozenEntries) {
        $frozenById.Add(
            (Get-RSRequiredString -InputObject $entry -Name candidateId),
            $entry)
    }

    $candidates = Get-RSRequiredArray -InputObject $Manifest -Name candidates -Context 'collection'
    $candidateIds = [string[]]@(
        foreach ($candidate in $candidates) {
            Get-RSRequiredString `
                -InputObject $candidate `
                -Name candidateId `
                -Pattern '^[a-z0-9][a-z0-9._-]{0,63}$'
        }
    )
    Assert-RSOrdinalSortedUnique -Value $candidateIds -Context 'collection candidates'
    if ($ExpectedCandidateId.Count -ne 0) {
        Assert-RSExactOrdinalSet `
            -Actual $candidateIds `
            -Expected $ExpectedCandidateId `
            -Context 'collection candidates'
    }

    foreach ($candidate in $candidates) {
        $candidateId = Get-RSRequiredString -InputObject $candidate -Name candidateId
        if (-not $frozenById.ContainsKey($candidateId)) {
            throw [IO.InvalidDataException]::new(
                "Collection candidate '$candidateId' is not in the frozen set.")
        }
        $frozen = $frozenById[$candidateId]
        foreach ($name in @('pin', 'pinSha256')) {
            if ((Get-RSRequiredString -InputObject $candidate -Name $name) -cne
                (Get-RSRequiredString -InputObject $frozen -Name $name)) {
                throw [IO.InvalidDataException]::new(
                    "Collection candidate '$candidateId' $name differs from the frozen set.")
            }
        }
        $kind = Get-RSRequiredString -InputObject $frozen -Name scoredOrControl
        $outcome = Get-RSRequiredString -InputObject $candidate -Name outcome
        if ($kind -ceq 'control') {
            if ($outcome -cne 'constraint-control') {
                throw [IO.InvalidDataException]::new(
                    "Frozen control '$candidateId' is not a constraint-control record.")
            }
            $controlFailure = (Get-RSPropertyValue `
                    -InputObject $candidate `
                    -Name controlFailure `
                    -Required).Value
            $constraintId = Get-RSRequiredString -InputObject $controlFailure -Name constraintId
            if ($constraintId -cne (Get-RSRequiredString -InputObject $frozen -Name constraintId)) {
                throw [IO.InvalidDataException]::new(
                    "Control '$candidateId' constraint differs from the frozen set.")
            }
        } elseif ($outcome -ceq 'constraint-control') {
            throw [IO.InvalidDataException]::new(
                "Scored candidate '$candidateId' cannot become a constraint control.")
        } elseif ((Get-RSRequiredString `
                    -InputObject $frozen `
                    -Name disposition `
                    -Context "candidate '$candidateId'") -ceq 'source-rejected' -and
            $outcome -cne 'source-rejected') {
            throw [IO.InvalidDataException]::new(
                "Frozen source-rejected candidate '$candidateId' cannot be collected as '$outcome'.")
        }

        Assert-RSGateSemantics -Candidate $candidate -Requirements $Requirements
        if ($outcome -ceq 'source-rejected') {
            Assert-RSSourceRejectionSemantics `
                -Candidate $candidate `
                -FrozenCandidate $frozen `
                -Requirements $Requirements
        }
        if ($outcome -ceq 'complete') {
            $evidenceObjects = Get-RSRequiredArray `
                -InputObject $candidate `
                -Name evidenceObjects `
                -Context "candidate '$candidateId'"
            $evidenceIds = [string[]]@(
                $evidenceObjects | ForEach-Object {
                    Get-RSRequiredString -InputObject $_ -Name evidenceId
                }
            )
            Assert-RSOrdinalSortedUnique `
                -Value $evidenceIds `
                -Context "candidate '$candidateId' evidence objects"
            foreach ($requiredEvidenceId in @(
                    'abi.build-identity',
                    'abi.fuzz-probe',
                    'abi.layout-pair',
                    'abi.source-dependencies')) {
                if ($evidenceIds -cnotcontains
                    $requiredEvidenceId) {
                    throw [IO.InvalidDataException]::new(
                        "Complete candidate '$candidateId' lacks provider-owned '$requiredEvidenceId' proof.")
                }
            }

            $measurements = Get-RSRequiredArray `
                -InputObject $candidate `
                -Name measurements `
                -Context "candidate '$candidateId'"
            $metricIds = [string[]]@(
                $measurements | ForEach-Object {
                    Get-RSRequiredString -InputObject $_ -Name metricId
                }
            )
            Assert-RSOrdinalSortedUnique `
                -Value $metricIds `
                -Context "candidate '$candidateId' measurements"
        }
    }
}

function Read-RSTerminalEvidenceManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $resolvedPath = [IO.Path]::GetFullPath($Path)
    if (-not [IO.File]::Exists($resolvedPath)) {
        throw [IO.FileNotFoundException]::new("Evidence manifest does not exist: $resolvedPath", $resolvedPath)
    }
    if (-not [string]::Equals(
            [IO.Path]::GetFileName($resolvedPath),
            $script:EvidenceJsonLeaf,
            [StringComparison]::Ordinal)) {
        throw [IO.InvalidDataException]::new(
            "Evidence manifest leaf must be exactly $script:EvidenceJsonLeaf.")
    }
    $bytes = [IO.File]::ReadAllBytes($resolvedPath)
    if ($bytes.Length -eq 0 -or
        ($bytes.Length -ge 3 -and $bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf)) {
        throw [IO.InvalidDataException]::new('Evidence manifest must be nonempty UTF-8 without BOM.')
    }
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $manifest = $json |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    } catch {
        throw [IO.InvalidDataException]::new(
            'Evidence manifest is not valid duplicate-free UTF-8 JSON.',
            $_.Exception)
    }

    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $canonicalBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $manifest
    if ($canonicalBytes.Length -ne $bytes.Length -or
        -not [Linq.Enumerable]::SequenceEqual([byte[]]$bytes, [byte[]]$canonicalBytes)) {
        throw [IO.InvalidDataException]::new('Evidence manifest bytes are not exact RFC-8785 JCS.')
    }
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath (Join-Path $RepoRoot $script:SchemaLeaf) `
        -Context 'Evidence manifest'

    $digest = Get-RSFileSha256 -Path $resolvedPath
    $companionPath = Join-Path ([IO.Path]::GetDirectoryName($resolvedPath)) $script:EvidenceCompanionLeaf
    if (-not [IO.File]::Exists($companionPath)) {
        throw [IO.InvalidDataException]::new('Evidence digest companion is missing.')
    }
    $companionBytes = [IO.File]::ReadAllBytes($companionPath)
    $expectedCompanion = [Text.Encoding]::ASCII.GetBytes($digest + "`n")
    if ($companionBytes.Length -ne 65 -or
        -not [Linq.Enumerable]::SequenceEqual([byte[]]$companionBytes, [byte[]]$expectedCompanion)) {
        throw [IO.InvalidDataException]::new(
            'Evidence digest companion does not match the exact manifest bytes.')
    }

    $runId = Get-RSRequiredString -InputObject $manifest -Name runId
    $machineHash = Get-RSRequiredString -InputObject $manifest -Name machineHash
    $runDirectory = [IO.Path]::GetDirectoryName($resolvedPath)
    $areaDirectory = [IO.Path]::GetDirectoryName($runDirectory)
    $machineDirectory = [IO.Path]::GetDirectoryName($areaDirectory)
    if (-not [string]::Equals(
            [IO.Path]::GetFileName($runDirectory),
            $runId,
            [StringComparison]::Ordinal) -or
        -not [string]::Equals(
            [IO.Path]::GetFileName($areaDirectory),
            'TerminalEngine',
            [StringComparison]::Ordinal) -or
        -not [string]::Equals(
            [IO.Path]::GetFileName($machineDirectory),
            $machineHash,
            [StringComparison]::Ordinal)) {
        throw [IO.InvalidDataException]::new(
            'Evidence manifest path does not match machineHash/TerminalEngine/runId.')
    }

    return [pscustomobject]@{
        Path = $resolvedPath
        Digest = $digest
        Manifest = $manifest
    }
}

function Get-RSBytesSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [byte[]]$Bytes
    )

    return [Convert]::ToHexString(
        [Security.Cryptography.SHA256]::HashData($Bytes)).
        ToLowerInvariant()
}

function Get-RSRequirements {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Get-Content `
        -LiteralPath (Join-Path $RepoRoot $script:RequirementsLeaf) `
        -Raw `
        -Encoding UTF8 |
        ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
}

function Get-RSScoreAnchors {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Get-Content `
        -LiteralPath (Join-Path $RepoRoot $script:ScoreAnchorsLeaf) `
        -Raw `
        -Encoding UTF8 |
        ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
}

function Assert-RSCurrentIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Manifest,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$RepositoryState,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Frozen,

        [Parameter(Mandatory = $true)]
        [string]$SchemaSha256,

        [Parameter(Mandatory = $true)]
        [string]$HarnessSourceSha256
    )

    if ((Get-RSRequiredString -InputObject $Manifest -Name repositoryCommit) -cne
        [string]$RepositoryState.Commit -or
        -not (Get-RSRequiredBoolean -InputObject $Manifest -Name repositoryClean)) {
        throw [IO.InvalidDataException]::new(
            'Evidence repository identity does not match the current clean checkout.')
    }
    if ((Get-RSRequiredString -InputObject $Manifest -Name schemaSha256) -cne $SchemaSha256) {
        throw [IO.InvalidDataException]::new('Evidence schema digest is not current.')
    }
    $harness = (Get-RSPropertyValue -InputObject $Manifest -Name harness -Required).Value
    if ((Get-RSRequiredString -InputObject $harness -Name version) -cne $script:HarnessVersion -or
        (Get-RSRequiredString -InputObject $harness -Name sourceSha256) -cne $HarnessSourceSha256) {
        throw [IO.InvalidDataException]::new('Evidence harness identity is not current.')
    }
    $manifestFrozen = (Get-RSPropertyValue -InputObject $Manifest -Name frozen -Required).Value
    foreach ($name in @(
            'requirementsSha256',
            'scoreAnchorsSha256',
            'fixtureCorpusSha256',
            'measurementProtocolSha256',
            'candidateUniverseSha256',
            'candidateSetSha256',
            'candidateReviewSha256',
            'candidateAssessmentSha256')) {
        if ((Get-RSRequiredString -InputObject $manifestFrozen -Name $name) -cne
            (Get-RSRequiredString -InputObject $Frozen -Name $name)) {
            throw [IO.InvalidDataException]::new("Evidence frozen identity '$name' is not current.")
        }
    }
}

function Get-RSCollectionCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Manifest,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId
    )

    $candidates = Get-RSRequiredArray -InputObject $Manifest -Name candidates
    $matching = @(
        $candidates | Where-Object {
            (Get-RSRequiredString -InputObject $_ -Name candidateId) -ceq $CandidateId
        }
    )
    if ($matching.Count -ne 1) {
        throw [IO.InvalidDataException]::new(
            "Collection must contain exactly one '$CandidateId' record.")
    }
    return $matching[0]
}

function Get-RSScoringEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [object[]]$Collections,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$ScoreAnchors,

        [Parameter(Mandatory = $true)]
        [object[]]$StaticRatings
    )

    Import-Module (
        Join-Path $PSScriptRoot (
            'TerminalEngine\TerminalEngineCandidateAssessment.psm1'))
    Import-Module (
        Join-Path $PSScriptRoot (
            'TerminalEngine\TerminalEngineC7Proof.psm1'))
    $criteriaDefinitions = Get-RSRequiredArray `
        -InputObject $ScoreAnchors `
        -Name criteria `
        -Context 'score anchors'
    $staticById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($rating in $StaticRatings) {
        $criterionId = Get-RSRequiredString -InputObject $rating -Name criterionId
        if ($staticById.ContainsKey($criterionId)) {
            throw [IO.InvalidDataException]::new(
                "Assessment repeats static rating '$criterionId'.")
        }
        $staticById.Add($criterionId, $rating)
    }
    Assert-RSExactOrdinalSet `
        -Actual ([string[]]@($staticById.Keys)) `
        -Expected ([string[]]@(
            'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C10', 'C11')) `
        -Context "candidate '$CandidateId' static ratings"

    $x64Release = @(
        $Collections | Where-Object {
            (Get-RSRequiredString -InputObject $_.Manifest -Name platform) -ceq
                'x64' -and
            (Get-RSRequiredString -InputObject $_.Manifest -Name configuration) -ceq
                'Release'
        })
    if ($x64Release.Count -ne 1) {
        throw [IO.InvalidDataException]::new(
            'Scoring requires exactly one x64 Release collection.')
    }
    $criteria = [object[]]@()
    $total = [decimal]0
    foreach ($criterionDefinition in $criteriaDefinitions) {
        $criterionId = Get-RSRequiredString -InputObject $criterionDefinition -Name criterionId
        $weightProperty = Get-RSPropertyValue -InputObject $criterionDefinition -Name weight -Required
        $weight = [int]$weightProperty.Value
        $anchors = Get-RSRequiredArray `
            -InputObject $criterionDefinition `
            -Name anchors `
            -Context "criterion '$criterionId'"
        $anchorById = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
        foreach ($anchor in $anchors) {
            $anchorById.Add((Get-RSRequiredString -InputObject $anchor -Name anchorId), $anchor)
        }

        $derived = $null
        $evidenceRefs = [object[]]@()
        if ($criterionId -ceq 'C9') {
            $candidate = Get-RSCollectionCandidate `
                -Manifest $x64Release[0].Manifest `
                -CandidateId $CandidateId
            $measurements = Get-RSRequiredArray `
                -InputObject $candidate `
                -Name measurements
            $derived = Get-RSCandidatePerformanceRating -Measurements $measurements
            $performanceMetricIds = [string[]]@(
                'added-release-package-bytes',
                'create-destroy-us',
                'dirty-snapshot-us',
                'eight-session-retained-bytes',
                'parse-8mib-us')
            $evidenceRefs = [object[]]@(
                foreach ($metricId in $performanceMetricIds) {
                    [ordered]@{
                        inputManifestSha256 = [string]$x64Release[0].Digest
                        evidenceKind = 'metric'
                        evidenceId = $metricId
                    }
                })
        }
        else {
            if (-not $staticById.ContainsKey($criterionId)) {
                throw [IO.InvalidDataException]::new(
                    "Assessment has no static rating for '$criterionId'.")
            }
            $derived = $staticById[$criterionId]
            $scoreEvidenceId = "score.$($criterionId.ToLowerInvariant())"
            $derivedEvidenceSha256 = Get-RSRequiredString `
                -InputObject $derived `
                -Name evidenceSha256
            $evidenceRefs = [object[]]@(
                foreach ($collection in $Collections) {
                    $candidate = Get-RSCollectionCandidate `
                        -Manifest $collection.Manifest `
                        -CandidateId $CandidateId
                    $objects = Get-RSRequiredArray `
                        -InputObject $candidate `
                        -Name evidenceObjects
                    $matches = @(
                        $objects | Where-Object {
                            (Get-RSRequiredString `
                                -InputObject $_ `
                                -Name evidenceId) -ceq $scoreEvidenceId
                        })
                    if ($matches.Count -ne 1 -or
                        (Get-RSRequiredString `
                            -InputObject $matches[0] `
                            -Name sha256) -cne $derivedEvidenceSha256) {
                        throw [IO.InvalidDataException]::new(
                            "Candidate '$CandidateId' '$scoreEvidenceId' is not bound to the reviewed raw ledger.")
                    }
                    [ordered]@{
                        inputManifestSha256 = [string]$collection.Digest
                        evidenceKind = 'object'
                        evidenceId = $scoreEvidenceId
                    }
                })
        }

        $selectedRating = [int](Get-RSPropertyValue `
                -InputObject $derived `
                -Name rating `
                -Required).Value
        $selectedAnchorId = Get-RSRequiredString `
            -InputObject $derived `
            -Name anchorId
        if ($selectedRating -lt 0 -or $selectedRating -gt 5 -or
            -not $anchorById.ContainsKey($selectedAnchorId) -or
            [int](Get-RSPropertyValue `
                -InputObject $anchorById[$selectedAnchorId] `
                -Name rating `
                -Required).Value -ne $selectedRating) {
            throw [IO.InvalidDataException]::new(
                "Criterion '$criterionId' did not derive one matching frozen integer anchor.")
        }

        if ($criterionId -ceq 'C7' -and $selectedRating -eq 5) {
            $proofRefs = [Collections.Generic.List[object]]::new()
            $proofHashes = [Collections.Generic.HashSet[string]]::new(
                [StringComparer]::Ordinal)
            foreach ($collection in $Collections) {
                $candidate = Get-RSCollectionCandidate `
                    -Manifest $collection.Manifest `
                    -CandidateId $CandidateId
                $objects = Get-RSRequiredArray `
                    -InputObject $candidate `
                    -Name evidenceObjects
                $matches = @(
                    $objects | Where-Object {
                        (Get-RSRequiredString `
                            -InputObject $_ `
                            -Name evidenceId) -ceq 'score.c7-proof'
                    })
                if ($matches.Count -ne 1) {
                    throw [IO.InvalidDataException]::new(
                        "Candidate '$CandidateId' C7 rating 5 requires exactly one provider-owned score.c7-proof in every lane.")
                }
                Assert-RSTerminalEngineC7EvidenceObject `
                    -EvidenceObject $matches[0] `
                    -RepoRoot $RepoRoot
                $proofSha256 = Get-RSRequiredString `
                    -InputObject $matches[0] `
                    -Name sha256
                if (-not $proofHashes.Add($proofSha256)) {
                    throw [IO.InvalidDataException]::new(
                        "Candidate '$CandidateId' replays one C7 proof identity across multiple lanes.")
                }
                $proofRefs.Add([ordered]@{
                        inputManifestSha256 = [string]$collection.Digest
                        evidenceKind = 'object'
                        evidenceId = 'score.c7-proof'
                    })
            }
            $evidenceRefs = [object[]]@(
                @($evidenceRefs) + @($proofRefs.ToArray()))
        }

        $evidenceRefs = Sort-RSObjectsOrdinal `
            -InputObject $evidenceRefs `
            -KeyProperty inputManifestSha256, evidenceKind, evidenceId
        $contribution = ([decimal]$weight * [decimal]$selectedRating) / [decimal]5
        $total += $contribution
        $criteria += [ordered]@{
            criterionId = $criterionId
            weight = $weight
            rating = $selectedRating
            matchedAnchorId = $selectedAnchorId
            contribution = $contribution.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture)
            evidenceRefs = $evidenceRefs
        }
    }

    if (($criteria | ForEach-Object { $_.criterionId }) -join ',' -cne
        ($script:CriterionIds -join ',')) {
        throw [IO.InvalidDataException]::new('Frozen score criteria are not exactly C1 through C11.')
    }

    return [pscustomobject]@{
        Criteria = $criteria
        TotalDecimal = $total
        Total = $total.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture)
    }
}

function Get-RSTieTuple {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$X64ReleaseManifest,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$AssessmentCandidate,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Admission
    )

    $candidate = Get-RSCollectionCandidate `
        -Manifest $X64ReleaseManifest `
        -CandidateId $CandidateId
    $measurements = Get-RSRequiredArray -InputObject $candidate -Name measurements
    $ledger = (Get-RSPropertyValue `
            -InputObject $AssessmentCandidate `
            -Name rawLedger `
            -Required).Value
    $boundary = (Get-RSPropertyValue `
            -InputObject $ledger `
            -Name boundary `
            -Required).Value
    $integration = (Get-RSPropertyValue `
            -InputObject $ledger `
            -Name integration `
            -Required).Value
    $derivedCounts = [ordered]@{
        'crossed-abi-elements' = [UInt64](Get-RSPropertyValue `
                -InputObject $boundary `
                -Name crossedAbiElements `
                -Required).Value
        'non-system-runtime-dependencies' = [UInt64](Get-RSPropertyValue `
                -InputObject $integration `
                -Name nonSystemRuntimeDependencies `
                -Required).Value
    }
    foreach ($entry in $derivedCounts.GetEnumerator()) {
        $matches = @(
            $measurements | Where-Object {
                (Get-RSRequiredString -InputObject $_ -Name metricId) -ceq
                    [string]$entry.Key
            }
        )
        if ($matches.Count -ne 1) {
            throw [IO.InvalidDataException]::new(
                "Complete candidate '$CandidateId' must contain one '$($entry.Key)' x64 Release metric.")
        }
        $unit = Get-RSRequiredString -InputObject $matches[0] -Name unit
        $samples = Get-RSRequiredArray -InputObject $matches[0] -Name samples
        $warmup = (Get-RSPropertyValue -InputObject $matches[0] -Name warmupCount -Required).Value
        if ($unit -cne 'count' -or $samples.Count -ne 1 -or [int]$warmup -ne 0 -or
            $samples[0] -isnot [string] -or [string]$samples[0] -cnotmatch
            '^(?:0|[1-9][0-9]{0,19})$') {
            throw [IO.InvalidDataException]::new(
                "Metric '$($entry.Key)' must be one canonical uint64 count with no warmup.")
        }
        $parsed = [UInt64]0
        if (-not [UInt64]::TryParse(
                [string]$samples[0],
                [Globalization.NumberStyles]::None,
                [Globalization.CultureInfo]::InvariantCulture,
                [ref]$parsed)) {
            throw [IO.InvalidDataException]::new(
                "Metric '$($entry.Key)' exceeds uint64.")
        }
        if ($parsed -ne [UInt64]$entry.Value) {
            throw [IO.InvalidDataException]::new(
                "Metric '$($entry.Key)' differs from the reviewed raw ledger.")
        }
    }

    $packageMatches = @(
        $measurements | Where-Object {
            (Get-RSRequiredString -InputObject $_ -Name metricId) -ceq
                'added-release-package-bytes'
        })
    if ($packageMatches.Count -ne 1 -or
        (Get-RSRequiredString -InputObject $packageMatches[0] -Name unit) -cne
            'bytes' -or
        [int](Get-RSPropertyValue `
            -InputObject $packageMatches[0] `
            -Name warmupCount `
            -Required).Value -ne 0) {
        throw [IO.InvalidDataException]::new(
            'added-release-package-bytes must be one un-warmed x64 Release byte measurement.')
    }
    $packageSamples = Get-RSRequiredArray `
        -InputObject $packageMatches[0] `
        -Name samples
    if ($packageSamples.Count -ne 1 -or
        $packageSamples[0] -isnot [string] -or
        [string]$packageSamples[0] -cnotmatch '^(?:0|[1-9][0-9]{0,19})$') {
        throw [IO.InvalidDataException]::new(
            'added-release-package-bytes must have exactly one canonical uint64 sample.')
    }
    $packageBytes = [UInt64]0
    if (-not [UInt64]::TryParse(
            [string]$packageSamples[0],
            [Globalization.NumberStyles]::None,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$packageBytes)) {
        throw [IO.InvalidDataException]::new(
            'added-release-package-bytes exceeds uint64.')
    }

    return [ordered]@{
        ownerDays = ([UInt64](Get-RSPropertyValue `
                    -InputObject $Admission `
                    -Name ownerDays `
                    -Required).Value).ToString(
                [Globalization.CultureInfo]::InvariantCulture)
        crossedAbiElements = $derivedCounts['crossed-abi-elements'].ToString(
            [Globalization.CultureInfo]::InvariantCulture)
        nonSystemRuntimeDependencies =
            $derivedCounts['non-system-runtime-dependencies'].ToString(
                [Globalization.CultureInfo]::InvariantCulture)
        addedReleasePackageBytes = $packageBytes.ToString(
            [Globalization.CultureInfo]::InvariantCulture)
    }
}

function Compare-RSTieTuple {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Left,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Right
    )

    foreach ($name in $script:TieMetricIds.Keys) {
        $leftValue = [UInt64]::Parse(
            (Get-RSRequiredString -InputObject $Left -Name $name),
            [Globalization.CultureInfo]::InvariantCulture)
        $rightValue = [UInt64]::Parse(
            (Get-RSRequiredString -InputObject $Right -Name $name),
            [Globalization.CultureInfo]::InvariantCulture)
        if ($leftValue -lt $rightValue) {
            return -1
        }
        if ($leftValue -gt $rightValue) {
            return 1
        }
    }
    return 0
}

function New-RSSelection {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ScoredResults,

        [Parameter(Mandatory = $true)]
        [decimal]$Threshold,

        [Parameter(Mandatory = $true)]
        [decimal]$CohortDelta
    )

    if ($ScoredResults.Count -eq 0) {
        return [ordered]@{
            status = 'no-scored-candidate'
            rawLeaderScore = $null
            cohort = [object[]]@()
            winnerCandidateId = $null
        }
    }
    $rawLeader = [decimal]0
    foreach ($result in $ScoredResults) {
        $value = [decimal]::Parse(
            (Get-RSRequiredString -InputObject $result -Name total),
            [Globalization.CultureInfo]::InvariantCulture)
        if ($value -gt $rawLeader) {
            $rawLeader = $value
        }
    }
    $rawLeaderText = $rawLeader.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture)
    if ($rawLeader -lt $Threshold) {
        return [ordered]@{
            status = 'below-threshold'
            rawLeaderScore = $rawLeaderText
            cohort = [object[]]@()
            winnerCandidateId = $null
        }
    }

    $cohortResults = [object[]]@(
        $ScoredResults | Where-Object {
            $score = [decimal]::Parse(
                (Get-RSRequiredString -InputObject $_ -Name total),
                [Globalization.CultureInfo]::InvariantCulture)
            $score -ge $Threshold -and $score -ge ($rawLeader - $CohortDelta)
        }
    )
    $cohortIds = Get-RSOrdinalSortedStrings -Value ([string[]]@(
            $cohortResults | ForEach-Object {
                Get-RSRequiredString -InputObject $_ -Name candidateId
            }
        ))

    $ranked = [object[]]@($cohortResults)
    $tieComparison = [Comparison[object]] {
        param($left, $right)
        $comparison = Compare-RSTieTuple `
            -Left (Get-RSPropertyValue -InputObject $left -Name tieTuple -Required).Value `
            -Right (Get-RSPropertyValue -InputObject $right -Name tieTuple -Required).Value
        if ($comparison -ne 0) {
            return $comparison
        }
        return 0
    }
    [Array]::Sort($ranked, $tieComparison)
    $isTied = $ranked.Count -gt 1 -and
        (Compare-RSTieTuple `
            -Left (Get-RSPropertyValue -InputObject $ranked[0] -Name tieTuple -Required).Value `
            -Right (Get-RSPropertyValue -InputObject $ranked[1] -Name tieTuple -Required).Value) -eq 0
    if ($isTied) {
        return [ordered]@{
            status = 'tied-first-no-winner'
            rawLeaderScore = $rawLeaderText
            cohort = $cohortIds
            winnerCandidateId = $null
        }
    }
    return [ordered]@{
        status = 'winner'
        rawLeaderScore = $rawLeaderText
        cohort = $cohortIds
        winnerCandidateId = Get-RSRequiredString -InputObject $ranked[0] -Name candidateId
    }
}

function New-RSRunDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EvidenceRoot,

        [Parameter(Mandatory = $true)]
        [string]$MachineHash
    )

    $resolvedRoot = [IO.Path]::GetFullPath($EvidenceRoot)
    for ($attempt = 0; $attempt -lt 3; ++$attempt) {
        $now = [DateTime]::UtcNow
        $runId = $now.ToString('yyyy-MM-dd_HHmmss', [Globalization.CultureInfo]::InvariantCulture)
        $runDirectory = Join-Path $resolvedRoot "$MachineHash\TerminalEngine\$runId"
        $areaDirectory = [IO.Path]::GetDirectoryName($runDirectory)
        $reservationPath = Join-Path $areaDirectory ".$runId.reserve"
        $reservation = $null
        try {
            $null = [IO.Directory]::CreateDirectory($areaDirectory)
            $reservation = [IO.FileStream]::new(
                $reservationPath,
                [IO.FileMode]::CreateNew,
                [IO.FileAccess]::Write,
                [IO.FileShare]::None,
                1,
                [IO.FileOptions]::DeleteOnClose)
            if ([IO.Directory]::Exists($runDirectory)) {
                throw [IO.IOException]::new("TerminalEngine runId '$runId' already exists.")
            }
            $null = [IO.Directory]::CreateDirectory($runDirectory)
            if (@([IO.Directory]::EnumerateFileSystemEntries($runDirectory)).Count -ne 0) {
                throw [IO.IOException]::new("New TerminalEngine run directory '$runId' is not empty.")
            }
            return [pscustomobject]@{
                RunId = $runId
                CreatedUtc = $now.ToString(
                    'yyyy-MM-ddTHH:mm:ssZ',
                    [Globalization.CultureInfo]::InvariantCulture)
                Directory = $runDirectory
            }
        } catch [IO.IOException] {
            # A second-level collision is retried at the next UTC second.
        } finally {
            if ($null -ne $reservation) {
                $reservation.Dispose()
            }
        }
        $milliseconds = 1100 - [DateTime]::UtcNow.Millisecond
        [Threading.Thread]::Sleep([Math]::Max(1, $milliseconds))
    }
    throw [IO.IOException]::new('Unable to allocate a create-new UTC TerminalEngine run directory.')
}

function Write-RSTerminalEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Manifest,

        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $json = [Text.Encoding]::UTF8.GetString((ConvertTo-RSJcsUtf8Bytes -InputObject $Manifest))
    Test-RSJsonSchema `
        -Json $json `
        -SchemaPath (Join-Path $RepoRoot $script:SchemaLeaf) `
        -Context 'Generated evidence manifest'
    return Write-RSEvidenceRecord `
        -Manifest $Manifest `
        -Directory $Directory `
        -JsonLeaf $script:EvidenceJsonLeaf `
        -CompanionLeaf $script:EvidenceCompanionLeaf
}

function Invoke-RSTerminalEngineFinalization {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$EvidenceManifest,

        [Parameter(Mandatory = $true)]
        [string]$EvidenceRoot,

        [switch]$RequireAllEvidence
    )

    $candidateSet = Read-RSCandidateSetPreflight -RepoRoot $RepoRoot
    Assert-RSExecutableCandidateCohort -CandidateSet $candidateSet
    $governance = Read-RSCandidateGovernanceClosure -RepoRoot $RepoRoot
    $candidateSet = $governance.CandidateSet
    $candidateAssessment = $governance.AssessmentResult
    $requirements = Get-RSRequirements -RepoRoot $RepoRoot
    $scoreAnchors = Get-RSScoreAnchors -RepoRoot $RepoRoot
    $repositoryState = Get-RSRepositoryState -RepoRoot $RepoRoot
    $schemaSha256 = Get-RSFileSha256 -Path (Join-Path $RepoRoot $script:SchemaLeaf)
    $harnessSourceSha256 = Get-RSTerminalHarnessSourceSha256 -RepoRoot $RepoRoot
    $frozen = Get-RSFrozenState -RepoRoot $RepoRoot -CandidateSet $candidateSet

    $collections = [object[]]@(
        foreach ($path in $EvidenceManifest) {
            Read-RSTerminalEvidenceManifest -Path $path -RepoRoot $RepoRoot
        }
    )
    $manifestPaths = [string[]]@($collections | ForEach-Object Path)
    Assert-RSOrdinalSortedUnique `
        -Value (Get-RSOrdinalSortedStrings -Value $manifestPaths) `
        -Context 'EvidenceManifest paths'

    $laneKeys = [string[]]@()
    foreach ($collection in $collections) {
        $manifest = $collection.Manifest
        Assert-RSCurrentIdentity `
            -Manifest $manifest `
            -RepositoryState $repositoryState `
            -Frozen $frozen `
            -SchemaSha256 $schemaSha256 `
            -HarnessSourceSha256 $harnessSourceSha256
        Assert-RSCollectionSemantics `
            -Manifest $manifest `
            -CandidateSet $candidateSet `
            -Requirements $requirements
        $laneKeys += "{0}|{1}" -f
            (Get-RSRequiredString -InputObject $manifest -Name platform),
            (Get-RSRequiredString -InputObject $manifest -Name configuration)
    }
    Assert-RSExactOrdinalSet `
        -Actual $laneKeys `
        -Expected $script:RequiredLaneKeys `
        -Context 'finalization lanes'

    $collections = Sort-RSObjectsOrdinal `
        -InputObject ([object[]]@(
                foreach ($collection in $collections) {
                    [pscustomobject]@{
                        SortPlatform = Get-RSRequiredString -InputObject $collection.Manifest -Name platform
                        SortConfiguration = Get-RSRequiredString -InputObject $collection.Manifest -Name configuration
                        Path = $collection.Path
                        Digest = $collection.Digest
                        Manifest = $collection.Manifest
                    }
                }
            )) `
        -KeyProperty SortPlatform, SortConfiguration

    $x64Release = @(
        $collections | Where-Object {
            $_.SortPlatform -ceq 'x64' -and $_.SortConfiguration -ceq 'Release'
        }
    )[0]
    $x64Asan = @(
        $collections | Where-Object {
            $_.SortPlatform -ceq 'x64' -and $_.SortConfiguration -ceq 'ASan Debug'
        }
    )[0]
    $x64ReleaseInputSet = Get-RSRequiredString `
        -InputObject (Get-RSPropertyValue `
                -InputObject $x64Release.Manifest `
                -Name collectionPolicy `
                -Required).Value `
        -Name bootstrapInputSetSha256
    $x64AsanInputSet = Get-RSRequiredString `
        -InputObject (Get-RSPropertyValue `
                -InputObject $x64Asan.Manifest `
                -Name collectionPolicy `
                -Required).Value `
        -Name bootstrapInputSetSha256
    if ($x64ReleaseInputSet -cne $x64AsanInputSet) {
        throw [IO.InvalidDataException]::new(
            'x64 Release and ASan Debug bootstrap input-set digests differ.')
    }

    $allFrozenCandidates = Get-RSCandidateSetEntries -CandidateSet $candidateSet
    $allFrozenIds = [string[]]@(
        $allFrozenCandidates | ForEach-Object {
            Get-RSRequiredString -InputObject $_ -Name candidateId
        }
    )
    foreach ($collection in $collections) {
        $actualIds = [string[]]@(
            (Get-RSRequiredArray -InputObject $collection.Manifest -Name candidates) |
                ForEach-Object {
                    Get-RSRequiredString -InputObject $_ -Name candidateId
                }
        )
        Assert-RSExactOrdinalSet `
            -Actual $actualIds `
            -Expected $allFrozenIds `
            -Context "collection '$($collection.Digest)' candidates"
    }

    $candidateResults = [object[]]@()
    $controlResults = [object[]]@()
    foreach ($frozenCandidate in $allFrozenCandidates) {
        $candidateId = Get-RSRequiredString -InputObject $frozenCandidate -Name candidateId
        $kind = Get-RSRequiredString -InputObject $frozenCandidate -Name scoredOrControl
        $laneCandidates = [object[]]@(
            foreach ($collection in $collections) {
                [pscustomobject]@{
                    Collection = $collection
                    Candidate = Get-RSCollectionCandidate `
                        -Manifest $collection.Manifest `
                        -CandidateId $candidateId
                }
            }
        )
        if ($kind -ceq 'control') {
            $constraintIds = [string[]]@(
                foreach ($laneCandidate in $laneCandidates) {
                    $controlFailure = (Get-RSPropertyValue `
                            -InputObject $laneCandidate.Candidate `
                            -Name controlFailure `
                            -Required).Value
                    Get-RSRequiredString -InputObject $controlFailure -Name constraintId
                }
            )
            if (@($constraintIds | Select-Object -Unique -CaseInsensitive:$false).Count -ne 1) {
                throw [IO.InvalidDataException]::new(
                    "Control '$candidateId' constraintId differs across lanes.")
            }
            $failureReferences = [object[]]@(
                foreach ($laneCandidate in $laneCandidates) {
                    $manifest = $laneCandidate.Collection.Manifest
                    $controlFailure = (Get-RSPropertyValue `
                            -InputObject $laneCandidate.Candidate `
                            -Name controlFailure `
                            -Required).Value
                    [ordered]@{
                        platform = Get-RSRequiredString -InputObject $manifest -Name platform
                        configuration = Get-RSRequiredString -InputObject $manifest -Name configuration
                        inputManifestSha256 = [string]$laneCandidate.Collection.Digest
                        evidenceSha256 = Get-RSRequiredString `
                            -InputObject $controlFailure `
                            -Name evidenceSha256
                    }
                }
            )
            $failureReferences = Sort-RSObjectsOrdinal `
                -InputObject $failureReferences `
                -KeyProperty platform, configuration
            $controlResults += [ordered]@{
                candidateId = $candidateId
                outcome = 'constraint-control'
                constraintId = $constraintIds[0]
                failureReferences = $failureReferences
            }
            continue
        }

        $sourceRejections = @(
            $laneCandidates | Where-Object {
                (Get-RSRequiredString -InputObject $_.Candidate -Name outcome) -ceq 'source-rejected'
            }
        )
        if ($sourceRejections.Count -ne 0) {
            if ($sourceRejections.Count -ne 3) {
                throw [IO.InvalidDataException]::new(
                    "Source-rejected candidate '$candidateId' must have the same rejection in all three lanes.")
            }
            $tupleIdentities = [string[]]@(
                foreach ($laneCandidate in $sourceRejections) {
                    $sourceRejection = (Get-RSPropertyValue `
                            -InputObject $laneCandidate.Candidate `
                            -Name sourceRejection `
                            -Required).Value
                    @(
                        'gateId',
                        'failureCode',
                        'candidateReviewSha256',
                        'blockerId',
                        'blockerSha256',
                        'sourceArchiveSha256'
                    ) | ForEach-Object {
                        Get-RSRequiredString -InputObject $sourceRejection -Name $_
                    } | Join-String -Separator '|'
                })
            if (@($tupleIdentities | Select-Object -Unique -CaseInsensitive:$false).Count -ne 1) {
                throw [IO.InvalidDataException]::new(
                    "Source-rejected candidate '$candidateId' differs across lanes.")
            }
            $failureReferences = [object[]]@(
                foreach ($laneCandidate in $sourceRejections) {
                    $sourceRejection = (Get-RSPropertyValue `
                            -InputObject $laneCandidate.Candidate `
                            -Name sourceRejection `
                            -Required).Value
                    [ordered]@{
                        inputManifestSha256 = [string]$laneCandidate.Collection.Digest
                        gateId = Get-RSRequiredString -InputObject $sourceRejection -Name gateId
                        failureCode = Get-RSRequiredString -InputObject $sourceRejection -Name failureCode
                        candidateReviewSha256 = Get-RSRequiredString `
                            -InputObject $sourceRejection `
                            -Name candidateReviewSha256
                        blockerId = Get-RSRequiredString -InputObject $sourceRejection -Name blockerId
                        blockerSha256 = Get-RSRequiredString `
                            -InputObject $sourceRejection `
                            -Name blockerSha256
                        sourceArchiveSha256 = Get-RSRequiredString `
                            -InputObject $sourceRejection `
                            -Name sourceArchiveSha256
                    }
                })
            $failureReferences = Sort-RSObjectsOrdinal `
                -InputObject $failureReferences `
                -KeyProperty inputManifestSha256, gateId, failureCode, blockerId
            $candidateResults += [ordered]@{
                candidateId = $candidateId
                outcome = 'mandatory-failed'
                eligibility = 'mandatory-failed'
                failureReferences = $failureReferences
            }
            continue
        }

        $earlyFailures = @(
            $laneCandidates | Where-Object {
                (Get-RSRequiredString -InputObject $_.Candidate -Name outcome) -ceq 'early-failure'
            }
        )
        if ($earlyFailures.Count -ne 0) {
            $failureReferences = [object[]]@(
                foreach ($laneCandidate in $earlyFailures) {
                    $earlyFailure = (Get-RSPropertyValue `
                            -InputObject $laneCandidate.Candidate `
                            -Name earlyFailure `
                            -Required).Value
                    [ordered]@{
                        inputManifestSha256 = [string]$laneCandidate.Collection.Digest
                        firstFailedGate = Get-RSRequiredString `
                            -InputObject $earlyFailure `
                            -Name firstFailedGate
                        failureCode = Get-RSRequiredString `
                            -InputObject $earlyFailure `
                            -Name failureCode
                    }
                }
            )
            $failureReferences = Sort-RSObjectsOrdinal `
                -InputObject $failureReferences `
                -KeyProperty inputManifestSha256, firstFailedGate, failureCode
            $candidateResults += [ordered]@{
                candidateId = $candidateId
                outcome = 'mandatory-failed'
                eligibility = 'mandatory-failed'
                failureReferences = $failureReferences
            }
            continue
        }
        foreach ($laneCandidate in $laneCandidates) {
            if ((Get-RSRequiredString -InputObject $laneCandidate.Candidate -Name outcome) -cne 'complete') {
                throw [IO.InvalidDataException]::new(
                    "Scored candidate '$candidateId' has neither complete, early-failure, nor source-rejected evidence.")
            }
        }

        $scoring = Get-RSScoringEvidence `
            -RepoRoot $RepoRoot `
            -Collections $collections `
            -CandidateId $candidateId `
            -ScoreAnchors $scoreAnchors `
            -StaticRatings $candidateAssessment.StaticRatingsById[$candidateId]
        $tieTuple = Get-RSTieTuple `
            -X64ReleaseManifest $x64Release.Manifest `
            -CandidateId $candidateId `
            -AssessmentCandidate $candidateAssessment.CandidateById[$candidateId] `
            -Admission $candidateAssessment.AdmissionById[$candidateId]
        $eligibility = if ($scoring.TotalDecimal -ge [decimal]80) {
            'eligible'
        } else {
            'below-threshold'
        }
        $candidateResults += [ordered]@{
            candidateId = $candidateId
            outcome = 'scored-complete'
            eligibility = $eligibility
            criteria = $scoring.Criteria
            total = $scoring.Total
            tieTuple = $tieTuple
        }
    }

    $candidateResults = Sort-RSObjectsOrdinal -InputObject $candidateResults -KeyProperty candidateId
    $controlResults = Sort-RSObjectsOrdinal -InputObject $controlResults -KeyProperty candidateId
    $scoredResults = [object[]]@(
        $candidateResults | Where-Object {
            (Get-RSRequiredString -InputObject $_ -Name outcome) -ceq 'scored-complete'
        }
    )
    $selectionProperty = Get-RSPropertyValue `
        -InputObject $requirements `
        -Name selection `
        -Required
    $threshold = [decimal]::Parse(
        (Get-RSRequiredString -InputObject $selectionProperty.Value -Name eligibleThreshold),
        [Globalization.CultureInfo]::InvariantCulture)
    $cohortDelta = [decimal]::Parse(
        (Get-RSRequiredString -InputObject $selectionProperty.Value -Name inclusiveCohortDelta),
        [Globalization.CultureInfo]::InvariantCulture)
    $selection = New-RSSelection `
        -ScoredResults $scoredResults `
        -Threshold $threshold `
        -CohortDelta $cohortDelta

    $machineHash = Get-RSTestMachineHash
    $run = New-RSRunDirectory -EvidenceRoot $EvidenceRoot -MachineHash $machineHash
    $inputManifests = [object[]]@(
        foreach ($collection in $collections) {
            [ordered]@{
                platform = [string]$collection.SortPlatform
                configuration = [string]$collection.SortConfiguration
                manifestSha256 = [string]$collection.Digest
            }
        }
    )
    $manifest = [ordered]@{
        formatId = 'red-salamander-terminal-engine-evidence'
        schemaVersion = 1
        recordKind = 'finalization'
        runId = $run.RunId
        createdUtc = $run.CreatedUtc
        repositoryCommit = [string]$repositoryState.Commit
        repositoryClean = $true
        machineHash = $machineHash
        schemaSha256 = $schemaSha256
        harness = [ordered]@{
            version = $script:HarnessVersion
            sourceSha256 = $harnessSourceSha256
        }
        frozen = $frozen
        inputManifests = $inputManifests
        candidateResults = $candidateResults
        controlResults = $controlResults
        selection = $selection
    }
    $manifestPath = Write-RSTerminalEvidence `
        -Manifest $manifest `
        -Directory $run.Directory `
        -RepoRoot $RepoRoot
    $exitCode = if ((Get-RSRequiredString -InputObject $selection -Name status) -ceq 'winner') {
        0
    } else {
        20
    }
    return [pscustomobject]@{
        ManifestPath = $manifestPath
        ExitCode = $exitCode
        SelectionStatus = Get-RSRequiredString -InputObject $selection -Name status
    }
}

function Invoke-RSCollectionProviderPhase {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProviderPath,

        [Parameter(Mandatory = $true)]
        [string]$Phase,

        [Parameter(Mandatory = $true)]
        [string[]]$CandidateId,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $output = @(
        & $ProviderPath `
            -Phase $Phase `
            -CandidateId $CandidateId `
            -Platform $Platform `
            -Configuration $Configuration `
            -RepoRoot $RepoRoot
    )
    if ($output.Count -ne 1 -or $null -eq $output[0]) {
        throw [IO.InvalidDataException]::new(
            "Collection provider phase '$Phase' must return exactly one structured object.")
    }
    return $output[0]
}

function Invoke-RSTerminalEngineCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$Candidate,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$EvidenceRoot,

        [string]$CollectionProviderPath = ''
    )

    $candidateSet = Read-RSCandidateSetPreflight -RepoRoot $RepoRoot
    Assert-RSExecutableCandidateCohort -CandidateSet $candidateSet
    $selectedIds = Resolve-RSCandidateSelection -Candidate $Candidate -CandidateSet $candidateSet
    Assert-RSExecutableCandidateCohort `
        -CandidateSet $candidateSet `
        -CandidateId $selectedIds
    if ($Candidate.Count -ne 1 -or $Candidate[0] -cne 'All') {
        throw [ArgumentException]::new(
            'Collect requires Candidate All; partial collection manifests are not finalizable.')
    }
    $governance = Read-RSCandidateGovernanceClosure -RepoRoot $RepoRoot
    $candidateSet = $governance.CandidateSet
    $requirements = Get-RSRequirements -RepoRoot $RepoRoot
    $repositoryState = Get-RSRepositoryState -RepoRoot $RepoRoot
    $schemaSha256 = Get-RSFileSha256 -Path (Join-Path $RepoRoot $script:SchemaLeaf)
    $harnessSourceSha256 = Get-RSTerminalHarnessSourceSha256 -RepoRoot $RepoRoot
    $frozen = Get-RSFrozenState -RepoRoot $RepoRoot -CandidateSet $candidateSet
    if ([string]::IsNullOrEmpty($CollectionProviderPath)) {
        $CollectionProviderPath = Join-Path $RepoRoot $script:CollectionProviderLeaf
    }
    if (-not [IO.File]::Exists($CollectionProviderPath)) {
        throw [IO.InvalidDataException]::new(
            "Candidate collection provider is missing: $script:CollectionProviderLeaf")
    }

    $bootstrap = Invoke-RSCollectionProviderPhase `
        -ProviderPath $CollectionProviderPath `
        -Phase Bootstrap `
        -CandidateId $selectedIds `
        -Platform $Platform `
        -Configuration $Configuration `
        -RepoRoot $RepoRoot
    if (-not (Get-RSRequiredBoolean -InputObject $bootstrap -Name bootstrap)) {
        throw [IO.InvalidDataException]::new('Collection provider did not attest bootstrap=true.')
    }
    $bootstrapInputSetSha256 = Get-RSRequiredString `
        -InputObject $bootstrap `
        -Name bootstrapInputSetSha256 `
        -Pattern '^[0-9a-f]{64}$'
    if ($bootstrapInputSetSha256 -cne
        (Get-RSBootstrapInputSetSha256 -CandidateSet $candidateSet -Platform $Platform)) {
        throw [IO.InvalidDataException]::new(
            'Collection provider bootstrap identity differs from the frozen lock.')
    }
    $null = Read-RSCandidateReview `
        -RepoRoot $RepoRoot `
        -CandidateSet $candidateSet `
        -VerifySourceArchives

    $clean = Invoke-RSCollectionProviderPhase `
        -ProviderPath $CollectionProviderPath `
        -Phase Clean `
        -CandidateId $selectedIds `
        -Platform $Platform `
        -Configuration $Configuration `
        -RepoRoot $RepoRoot
    if (-not (Get-RSRequiredBoolean -InputObject $clean -Name clean)) {
        throw [IO.InvalidDataException]::new('Collection provider did not attest clean=true.')
    }

    $collected = Invoke-RSCollectionProviderPhase `
        -ProviderPath $CollectionProviderPath `
        -Phase Collect `
        -CandidateId $selectedIds `
        -Platform $Platform `
        -Configuration $Configuration `
        -RepoRoot $RepoRoot
    if (-not (Get-RSRequiredBoolean -InputObject $collected -Name offline) -or
        (Get-RSRequiredString -InputObject $collected -Name networkIsolation) -cne 'runner-enforced') {
        throw [IO.InvalidDataException]::new(
            'Collection provider did not prove offline runner-enforced isolation.')
    }
    $nativeAttestation = (Get-RSPropertyValue `
            -InputObject $collected `
            -Name nativeAttestation `
            -Required).Value
    $toolchain = (Get-RSPropertyValue `
            -InputObject $collected `
            -Name toolchain `
            -Required).Value
    $candidates = Get-RSRequiredArray -InputObject $collected -Name candidates
    $postCollectionRepositoryState = Get-RSRepositoryState -RepoRoot $RepoRoot
    if ([string]$postCollectionRepositoryState.Commit -cne [string]$repositoryState.Commit) {
        throw [IO.InvalidDataException]::new(
            'Repository commit changed during TerminalEngine collection.')
    }

    $machineHash = Get-RSTestMachineHash
    $manifest = [ordered]@{
        formatId = 'red-salamander-terminal-engine-evidence'
        schemaVersion = 1
        recordKind = 'collection'
        runId = '2000-01-01_000000'
        createdUtc = '2000-01-01T00:00:00Z'
        repositoryCommit = [string]$repositoryState.Commit
        repositoryClean = $true
        machineHash = $machineHash
        schemaSha256 = $schemaSha256
        harness = [ordered]@{
            version = $script:HarnessVersion
            sourceSha256 = $harnessSourceSha256
        }
        frozen = $frozen
        platform = $Platform
        configuration = $Configuration
        nativeAttestation = $nativeAttestation
        toolchain = $toolchain
        collectionPolicy = [ordered]@{
            bootstrap = $true
            clean = $true
            offline = $true
            bootstrapInputSetSha256 = $bootstrapInputSetSha256
            networkIsolation = 'runner-enforced'
        }
        candidates = $candidates
    }
    Assert-RSCollectionSemantics `
        -Manifest $manifest `
        -CandidateSet $candidateSet `
        -Requirements $requirements `
        -ExpectedCandidateId $selectedIds
    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1')
    $preflightJson = [Text.Encoding]::UTF8.GetString(
        (ConvertTo-RSJcsUtf8Bytes -InputObject $manifest))
    Test-RSJsonSchema `
        -Json $preflightJson `
        -SchemaPath (Join-Path $RepoRoot $script:SchemaLeaf) `
        -Context 'Collected evidence manifest'

    $run = New-RSRunDirectory -EvidenceRoot $EvidenceRoot -MachineHash $machineHash
    $manifest['runId'] = $run.RunId
    $manifest['createdUtc'] = $run.CreatedUtc
    $manifestPath = Write-RSTerminalEvidence `
        -Manifest $manifest `
        -Directory $run.Directory `
        -RepoRoot $RepoRoot
    return [pscustomobject]@{
        ManifestPath = $manifestPath
        ExitCode = 0
        SelectionStatus = 'collection'
    }
}

Export-ModuleMember -Function `
    Assert-RSTerminalEngineInvocation, `
    Get-RSTerminalHarnessSourceSha256, `
    Read-RSCandidateUniverse, `
    Read-RSCurrentCandidateUniverse, `
    Read-RSHistoricalRound3CandidateUniverse, `
    Read-RSCandidateSet, `
    Read-RSCandidateReview, `
    Read-RSCandidateAssessment, `
    Read-RSCandidateGovernanceClosure, `
    Invoke-RSAuthenticatedCollectionNeutralityProof, `
    Read-RSTerminalEvidenceManifest, `
    Assert-RSCollectionSemantics, `
    New-RSSelection, `
    Invoke-RSTerminalEngineCollection, `
    Invoke-RSTerminalEngineFinalization
