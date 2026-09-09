[CmdletBinding()]
param(
    [string]$RepoRoot = '',

    [string[]]$GovernanceFile = @(),

    [string[]]$DescriptorFile = @(),

    [string[]]$FrozenInputFile = @(),

    [string[]]$AdditionalLiteral = @(),

    [string]$PolicyFile = (
        'Specs/Terminal/TerminalEngineCollectionNeutrality.v1.json'),

    [switch]$ValidateBaselineProof
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
}
$RepoRoot = [IO.Path]::GetFullPath($RepoRoot)
Import-Module (Join-Path $RepoRoot (
        'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1')) `
    -Force
Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1') -Force

function Read-RSNeutralityJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath
    )

    $absolute = if ([IO.Path]::IsPathFullyQualified($LiteralPath)) {
        [IO.Path]::GetFullPath($LiteralPath)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $RepoRoot $LiteralPath))
    }
    if (-not [IO.File]::Exists($absolute)) {
        throw [IO.FileNotFoundException]::new(
            "Neutrality governance input '$LiteralPath' is absent.",
            $absolute)
    }
    return [IO.File]::ReadAllText(
        $absolute,
        [Text.UTF8Encoding]::new($false, $true)) |
        ConvertFrom-Json `
            -AsHashtable `
            -Depth 100 `
            -NoEnumerate `
            -DateKind String
}

$governance = [object[]]@(
    $GovernanceFile | ForEach-Object {
        Read-RSNeutralityJson -LiteralPath $_
    })
$descriptors = [object[]]@(
    $DescriptorFile | ForEach-Object {
        Read-RSNeutralityJson -LiteralPath $_
    })
$frozenInputs = [object[]]@(
    $FrozenInputFile | ForEach-Object {
        Read-RSNeutralityJson -LiteralPath $_
    })
$policy = Read-RSNeutralityJson -LiteralPath $PolicyFile
$policyPath = if ([IO.Path]::IsPathFullyQualified($PolicyFile)) {
    [IO.Path]::GetFullPath($PolicyFile)
}
else {
    [IO.Path]::GetFullPath((Join-Path $RepoRoot $PolicyFile))
}
$policyBytes = [IO.File]::ReadAllBytes($policyPath)
$canonicalPolicyBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $policy
if (-not [Linq.Enumerable]::SequenceEqual(
        [byte[]]$policyBytes,
        [byte[]]$canonicalPolicyBytes) -or
    -not (Test-Json `
        -Json ([Text.Encoding]::UTF8.GetString($canonicalPolicyBytes)) `
        -SchemaFile (Join-Path $RepoRoot (
                'Specs\Terminal\TerminalEngineCollectionNeutrality.schema.json')) `
        -ErrorAction Stop)) {
    throw [IO.InvalidDataException]::new(
        'Collection-neutrality policy is not canonical JCS or violates its closed schema.')
}

Invoke-RSTerminalEngineCollectionNeutralityScan `
    -RepoRoot $RepoRoot `
    -GovernanceObject $governance `
    -Descriptor $descriptors `
    -FrozenInput $frozenInputs `
    -AdditionalLiteral $AdditionalLiteral `
    -PolicyRecord $policy `
    -ValidateBaselineProof:$ValidateBaselineProof
