Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) (
        'TerminalEvidence.psm1'))

$script:RSCollectionNeutralitySurface = [string[]]@(
    '.github/workflows/terminal-engine-gate0.yml',
    'Directory.Build.props',
    'Directory.Build.targets',
    'RuntimeDependencies.props',
    'Specs/Terminal/TerminalEngineCollectionDescriptor.schema.json',
    'Specs/Terminal/TerminalEngineC7Proof.schema.json',
    'Specs/Terminal/TerminalEngineC7ProofContract.v1.json',
    'Tools/Evaluate-TerminalEngines.ps1',
    'Tools/TerminalEvidence.psm1',
    'Tools/TerminalEngineEvaluator.psm1',
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
    'Tools/Tests/TerminalEngineArchive.Tests.ps1',
    'Tools/Tests/TerminalEngineC7Proof.Tests.ps1',
    'Tools/Tests/TerminalEngineCandidateAssessment.Tests.ps1',
    'Tools/Tests/TerminalEngineCandidatePatch.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionDriver.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionDescriptor.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionNeutrality.Tests.ps1',
    'Tools/Tests/TerminalEngineCollectionProvider.Tests.ps1',
    'Tools/Tests/TerminalEngineNotices.Tests.ps1',
    'vcpkg.json')
[Array]::Sort($script:RSCollectionNeutralitySurface, [StringComparer]::Ordinal)

function Get-RSTerminalEngineCollectionNeutralitySurface {
    return [string[]]@($script:RSCollectionNeutralitySurface)
}

function Add-RSCollectionNeutralityValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [Collections.Generic.HashSet[string]]$Result
    )

    $normalized = $Value.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($normalized) -or
        $normalized.Length -gt 256) {
        return
    }
    [void]$Result.Add($normalized)
    $variants = [string[]]@(
        [Text.RegularExpressions.Regex]::Replace(
            $normalized,
            '-rs-v[0-9]+$',
            ''),
        [Text.RegularExpressions.Regex]::Replace(
            $normalized,
            '-rs$',
            ''),
        [Text.RegularExpressions.Regex]::Replace(
            $normalized,
            '-upstream$',
            ''),
        [Text.RegularExpressions.Regex]::Replace(
            $normalized,
            '-core$',
            ''),
        [Text.RegularExpressions.Regex]::Replace(
            $normalized,
            '^(?:source|notice|archive)-',
            ''),
        [Text.RegularExpressions.Regex]::Replace(
            $normalized,
            '(?:\.tar\.gz|\.nupkg|\.vcxproj|\.dll|\.lib|\.zip|\.txt|\.json)$',
            ''))
    foreach ($variant in $variants) {
        if ($variant -cne $normalized -and $variant.Length -ge 3) {
            [void]$Result.Add($variant)
        }
    }
}

function Add-RSCollectionNeutralityExecutableCandidateIds {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [Collections.Generic.HashSet[string]]$Result
    )

    if ($null -eq $Value) {
        return
    }
    if ($Value -is [Collections.IDictionary]) {
        if ($Value.Contains('candidateId') -and
            $Value.Contains('scoredOrControl') -and
            $Value.Contains('disposition') -and
            [string]$Value['scoredOrControl'] -ceq 'scored' -and
            [string]$Value['disposition'] -ceq 'collect') {
            [void]$Result.Add([string]$Value['candidateId'])
        }
        foreach ($entry in $Value.GetEnumerator()) {
            Add-RSCollectionNeutralityExecutableCandidateIds `
                -Value $entry.Value `
                -Result $Result
        }
        return
    }
    if ($Value -is [Collections.IEnumerable] -and
        $Value -isnot [string]) {
        foreach ($entry in $Value) {
            Add-RSCollectionNeutralityExecutableCandidateIds `
                -Value $entry `
                -Result $Result
        }
    }
}

function Add-RSCollectionNeutralityInputRecords {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [Collections.Generic.List[object]]$Result
    )

    if ($null -eq $Value) {
        return
    }
    if ($Value -is [Collections.IDictionary]) {
        if ($Value.Contains('inputId') -and
            $Value.Contains('inputKind') -and
            $Value.Contains('fileName') -and
            $Value.Contains('consumers')) {
            [void]$Result.Add($Value)
        }
        foreach ($entry in $Value.GetEnumerator()) {
            Add-RSCollectionNeutralityInputRecords `
                -Value $entry.Value `
                -Result $Result
        }
        return
    }
    if ($Value -is [Collections.IEnumerable] -and
        $Value -isnot [string]) {
        foreach ($entry in $Value) {
            Add-RSCollectionNeutralityInputRecords `
                -Value $entry `
                -Result $Result
        }
    }
}

function Get-RSCollectionNeutralityNeutralHostInputLiterals {
    param(
        [AllowEmptyCollection()]
        [object[]]$GovernanceObject = @(),

        [AllowEmptyCollection()]
        [object[]]$Record = @()
    )

    $executableCandidates =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
    foreach ($value in $GovernanceObject) {
        Add-RSCollectionNeutralityExecutableCandidateIds `
            -Value $value `
            -Result $executableCandidates
    }
    $result =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    if ($executableCandidates.Count -eq 0) {
        return ,$result
    }

    $inputs = [Collections.Generic.List[object]]::new()
    foreach ($value in @($GovernanceObject) + @($Record)) {
        Add-RSCollectionNeutralityInputRecords `
            -Value $value `
            -Result $inputs
    }
    $hostRecordFound = $false
    $hostRecordInvalid = $false
    foreach ($input in $inputs) {
        if ([string]$input['inputId'] -cne 'host-wil-headers') {
            continue
        }
        $hostRecordFound = $true
        if ([string]$input['inputKind'] -cne 'dependency-input' -or
            [string]$input['fileName'] -cne 'host-wil-headers.zip') {
            $hostRecordInvalid = $true
            continue
        }
        $consumers =
            [Collections.Generic.HashSet[string]]::new(
                [StringComparer]::Ordinal)
        $valid = $true
        foreach ($consumer in [object[]]@($input['consumers'])) {
            if ($consumer -isnot [string] -or
                -not $consumers.Add([string]$consumer)) {
                $valid = $false
                break
            }
        }
        if (-not $valid -or
            $consumers.Count -ne $executableCandidates.Count) {
            $hostRecordInvalid = $true
            continue
        }
        foreach ($candidateId in $executableCandidates) {
            if (-not $consumers.Contains($candidateId)) {
                $valid = $false
                break
            }
        }
        if ($valid) {
            continue
        }
        $hostRecordInvalid = $true
    }
    if ($hostRecordFound -and -not $hostRecordInvalid) {
        [void]$result.Add('host-wil-headers')
        [void]$result.Add('host-wil-headers.zip')
    }
    return ,$result
}

function Add-RSCollectionNeutralityObjectValues {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [Collections.Generic.HashSet[string]]$Result,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [Collections.Generic.HashSet[string]]$NeutralHostInputLiteral
    )

    if ($null -eq $Value) {
        return
    }
    if ($Value -is [Collections.IDictionary]) {
        $isToolchainInput =
            $Value.Contains('inputKind') -and
            [string]$Value['inputKind'] -ceq 'toolchain-input'
        foreach ($entry in $Value.GetEnumerator()) {
            $name = [string]$entry.Key
            if ($name -cin @(
                    'candidateId',
                    'seedCandidateId',
                    'nominationId',
                    'priorNominationId',
                    'dependencyId',
                    'inputId',
                    'sourceInputId',
                    'closureId',
                    'libraryId',
                    'destinationLeaf',
                    'fileName',
                    'dllLeaf',
                    'projectLeaf',
                    'libraryLeaf')) {
                if ($entry.Value -is [string] -and
                    -not $NeutralHostInputLiteral.Contains(
                        [string]$entry.Value)) {
                    if ($isToolchainInput -and $name -ceq 'fileName') {
                        $exact = ([string]$entry.Value).
                            Trim().
                            ToLowerInvariant()
                        if (-not [string]::IsNullOrWhiteSpace($exact) -and
                            $exact.Length -le 256) {
                            [void]$Result.Add($exact)
                        }
                    }
                    else {
                        Add-RSCollectionNeutralityValue `
                            -Value ([string]$entry.Value) `
                            -Result $Result
                    }
                }
            }
            Add-RSCollectionNeutralityObjectValues `
                -Value $entry.Value `
                -Result $Result `
                -NeutralHostInputLiteral $NeutralHostInputLiteral
        }
        return
    }
    if ($Value -is [Collections.IEnumerable] -and
        $Value -isnot [string]) {
        foreach ($entry in $Value) {
            Add-RSCollectionNeutralityObjectValues `
                -Value $entry `
                -Result $Result `
                -NeutralHostInputLiteral $NeutralHostInputLiteral
        }
    }
}

function Add-RSCollectionNeutralityCandidatePathValues {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [Collections.Generic.HashSet[string]]$Result
    )

    if ($null -eq $Value) {
        return
    }
    if ($Value -is [Collections.IDictionary]) {
        foreach ($entry in $Value.GetEnumerator()) {
            $name = [string]$entry.Key
            if ($entry.Value -is [string] -and
                ($name.EndsWith(
                        'RelativePath',
                        [StringComparison]::Ordinal) -or
                    $name -ceq 'relativePath')) {
                foreach ($segment in (
                        ([string]$entry.Value) -split '[/\\]')) {
                    $compactSegment =
                        ($segment.ToLowerInvariant() -replace '[^a-z0-9]', '')
                    if ($compactSegment.Length -lt 3) {
                        continue
                    }
                    $candidateBearing = $false
                    foreach ($known in [string[]]@($Result)) {
                        $compactKnown = $known -replace '[^a-z0-9]', ''
                        if ($compactKnown.Length -ge 3 -and
                            $compactSegment.Contains(
                                $compactKnown,
                                [StringComparison]::Ordinal)) {
                            $candidateBearing = $true
                            break
                        }
                    }
                    if ($candidateBearing) {
                        Add-RSCollectionNeutralityValue `
                            -Value $segment `
                            -Result $Result
                    }
                }
            }
            Add-RSCollectionNeutralityCandidatePathValues `
                -Value $entry.Value `
                -Result $Result
        }
        return
    }
    if ($Value -is [Collections.IEnumerable] -and
        $Value -isnot [string]) {
        foreach ($entry in $Value) {
            Add-RSCollectionNeutralityCandidatePathValues `
                -Value $entry `
                -Result $Result
        }
    }
}

function Get-RSTerminalEngineCollectionForbiddenLiteral {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()]
        [object[]]$GovernanceObject = @(),

        [AllowEmptyCollection()]
        [object[]]$Descriptor = @(),

        [AllowEmptyCollection()]
        [object[]]$FrozenInput = @(),

        [AllowEmptyCollection()]
        [string[]]$AdditionalLiteral = @()
    )

    $result =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    $neutralHostInputLiteral =
        Get-RSCollectionNeutralityNeutralHostInputLiterals `
            -GovernanceObject $GovernanceObject `
            -Record (@($Descriptor) + @($FrozenInput))
    foreach ($record in @($GovernanceObject) + @($Descriptor) + @($FrozenInput)) {
        Add-RSCollectionNeutralityObjectValues `
            -Value $record `
            -Result $result `
            -NeutralHostInputLiteral $neutralHostInputLiteral
    }
    foreach ($record in @($GovernanceObject) + @($Descriptor) + @($FrozenInput)) {
        Add-RSCollectionNeutralityCandidatePathValues `
            -Value $record `
            -Result $result
    }
    foreach ($literal in $AdditionalLiteral) {
        Add-RSCollectionNeutralityValue -Value $literal -Result $result
    }
    $ordered = [string[]]@($result)
    [Array]::Sort($ordered, [StringComparer]::Ordinal)
    return $ordered
}

function Get-RSCollectionNeutralitySha256Text {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($Value)
    return [Convert]::ToHexString(
        [Security.Cryptography.SHA256]::HashData($bytes)).
        ToLowerInvariant()
}

function Assert-RSCollectionNeutralityExactProperties {
    param(
        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$Value,

        [Parameter(Mandatory = $true)]
        [string[]]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $actual = [string[]]@($Value.Keys | ForEach-Object { [string]$_ })
    $expectedSorted = [string[]]@($Expected)
    [Array]::Sort($actual, [StringComparer]::Ordinal)
    [Array]::Sort($expectedSorted, [StringComparer]::Ordinal)
    if (-not [Linq.Enumerable]::SequenceEqual($actual, $expectedSorted)) {
        throw [IO.InvalidDataException]::new(
            "$Label does not have its exact closed property set.")
    }
}

function Assert-RSCollectionNeutralityOrdinalStrings {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $sorted = [string[]]@($Value)
    [Array]::Sort($sorted, [StringComparer]::Ordinal)
    if (-not [Linq.Enumerable]::SequenceEqual($Value, $sorted)) {
        throw [IO.InvalidDataException]::new(
            "$Label must be sorted ordinally.")
    }
    for ($index = 1; $index -lt $Value.Count; ++$index) {
        if ($Value[$index - 1] -ceq $Value[$index]) {
            throw [IO.InvalidDataException]::new(
                "$Label must not contain duplicates.")
        }
    }
}

function Get-RSCollectionNeutralityIdentifierFacts {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text
    )

    $identifiers =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    $segments =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($match in [Text.RegularExpressions.Regex]::Matches(
            $Text,
            '[A-Za-z][A-Za-z0-9_]*',
            [Text.RegularExpressions.RegexOptions]::CultureInvariant)) {
        [void]$identifiers.Add((
                ([string]$match.Value).ToLowerInvariant() -replace
                    '[^a-z0-9]', ''))
        foreach ($underscorePart in ([string]$match.Value).Split('_')) {
            $camel = [Text.RegularExpressions.Regex]::Replace(
                $underscorePart,
                '(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])',
                ' ')
            foreach ($part in $camel.Split(
                    ' ',
                    [StringSplitOptions]::RemoveEmptyEntries)) {
                [void]$segments.Add($part.ToLowerInvariant())
            }
        }
    }
    return [pscustomobject]@{
        Identifiers = $identifiers
        Segments = $segments
    }
}

function Get-RSCollectionNeutralityAstLiteral {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [Management.Automation.Language.Ast]$Ast
    )

    if ($Ast -is
        [Management.Automation.Language.StringConstantExpressionAst]) {
        return [string]$Ast.Value
    }
    if ($Ast -is
            [Management.Automation.Language.ExpandableStringExpressionAst] -and
        @($Ast.NestedExpressions).Count -eq 0) {
        return [string]$Ast.Value
    }
    if ($Ast -is [Management.Automation.Language.BinaryExpressionAst] -and
        $Ast.Operator -eq
            [Management.Automation.Language.TokenKind]::Plus) {
        $left = Get-RSCollectionNeutralityAstLiteral -Ast $Ast.Left
        $right = Get-RSCollectionNeutralityAstLiteral -Ast $Ast.Right
        if ($null -ne $left -and $null -ne $right) {
            return "$left$right"
        }
    }
    return $null
}

function Get-RSCollectionNeutralityStringCompositions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseInput(
        $Text,
        [ref]$tokens,
        [ref]$errors)
    $result =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($binary in $ast.FindAll(
            {
                param($node)
                $node -is
                    [Management.Automation.Language.BinaryExpressionAst] -and
                $node.Operator -eq
                    [Management.Automation.Language.TokenKind]::Plus
            },
            $true)) {
        $literal = Get-RSCollectionNeutralityAstLiteral -Ast $binary
        if ($null -ne $literal) {
            $compact =
                $literal.ToLowerInvariant() -replace '[^a-z0-9]', ''
            $key = "$compact`0$($binary.Extent.StartLineNumber)"
            [void]$result.TryAdd($key, [pscustomobject]@{
                    Compact = $compact
                    Line = [int]$binary.Extent.StartLineNumber
                })
        }
    }
    foreach ($invocation in $ast.FindAll(
            {
                param($node)
                $node -is
                    [Management.Automation.Language.InvokeMemberExpressionAst]
            },
            $true)) {
        $member = Get-RSCollectionNeutralityAstLiteral `
            -Ast $invocation.Member
        if ([string]$member -cne 'Concat') {
            continue
        }
        $builder = [Text.StringBuilder]::new()
        $complete = $invocation.Arguments.Count -ge 2
        foreach ($argument in $invocation.Arguments) {
            $literal = Get-RSCollectionNeutralityAstLiteral -Ast $argument
            if ($null -eq $literal) {
                $complete = $false
                break
            }
            [void]$builder.Append($literal)
        }
        if ($complete) {
            $compact =
                $builder.ToString().ToLowerInvariant() -replace
                    '[^a-z0-9]', ''
            $key = "$compact`0$($invocation.Extent.StartLineNumber)"
            [void]$result.TryAdd($key, [pscustomobject]@{
                    Compact = $compact
                    Line = [int]$invocation.Extent.StartLineNumber
                })
        }
    }

    $extension = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($extension -cin @(
            '.c', '.cc', '.cpp', '.cxx', '.h', '.hh', '.hpp',
            '.zig', '.rs')) {
        $quoted = [Text.RegularExpressions.Regex]::Matches(
            $Text,
            '"(?:\\.|[^"\\])*"',
            [Text.RegularExpressions.RegexOptions]::CultureInvariant)
        for ($index = 1; $index -lt $quoted.Count; ++$index) {
            $left = $quoted[$index - 1]
            $right = $quoted[$index]
            $betweenOffset = $left.Index + $left.Length
            $between = $Text.Substring(
                $betweenOffset,
                $right.Index - $betweenOffset)
            if (-not [Text.RegularExpressions.Regex]::IsMatch(
                    $between,
                    '^(?:(?:\s+)|(?:/\*[\s\S]*?\*/)|(?://[^\r\n]*(?:\r?\n|$)))*$',
                    [Text.RegularExpressions.RegexOptions]::CultureInvariant)) {
                continue
            }
            $leftValue = $left.Value.Substring(1, $left.Length - 2)
            $rightValue = $right.Value.Substring(1, $right.Length - 2)
            $compact = (
                "$leftValue$rightValue".ToLowerInvariant() -replace
                    '[^a-z0-9]', '')
            $prefix = $Text.Substring(0, $left.Index)
            $line = 1 + (
                [Text.RegularExpressions.Regex]::Matches(
                    $prefix,
                    "\n",
                    [Text.RegularExpressions.RegexOptions]::CultureInvariant)
            ).Count
            $key = "$compact`0$line"
            [void]$result.TryAdd($key, [pscustomobject]@{
                    Compact = $compact
                    Line = [int]$line
                })
        }
    }
    return [object[]]@($result.Values)
}

function Test-RSCollectionNeutralityLiteralHit {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Line,

        [Parameter(Mandatory = $true)]
        [string]$Literal,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$IdentifierFacts
    )

    $escaped = [Text.RegularExpressions.Regex]::Escape($Literal)
    if ([Text.RegularExpressions.Regex]::IsMatch(
            $Line,
            "(?i)$escaped",
            [Text.RegularExpressions.RegexOptions]::CultureInvariant)) {
        return $true
    }
    $compact = $Literal -replace '[^a-z0-9]', ''
    if ($compact.Length -le 3) {
        return $IdentifierFacts.Segments.Contains($compact)
    }
    foreach ($identifier in $IdentifierFacts.Identifiers) {
        if ($identifier.Contains(
                $compact,
                [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Initialize-RSCollectionNeutralityNative {
    $typeName =
        'RedSalamander.TerminalEngine.CollectionNeutralitySnapshotNative'
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
    public static class CollectionNeutralitySnapshotNative
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

function Read-RSCollectionNeutralitySurfaceSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    Initialize-RSCollectionNeutralityNative
    $nativeType =
        'RedSalamander.TerminalEngine.CollectionNeutralitySnapshotNative' -as
            [type]
    $nativePath = if ($Path.StartsWith(
            '\\',
            [StringComparison]::Ordinal)) {
        '\\?\UNC\' + $Path.Substring(2)
    } else {
        '\\?\' + $Path
    }
    $handle = $nativeType::CreateFileW(
        $nativePath,
        [uint32]2147483648,
        [uint32]1,
        [IntPtr]::Zero,
        [uint32]3,
        [uint32]0x00200000,
        [IntPtr]::Zero)
    if ($null -eq $handle -or $handle.IsInvalid) {
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        if ($null -ne $handle) {
            $handle.Dispose()
        }
        throw [ComponentModel.Win32Exception]::new(
            $lastError,
            "Cannot acquire collection-neutrality surface '$RelativePath'.")
    }
    $handleAttributes = $nativeType::GetAttributesByHandle($handle)
    if (($handleAttributes -band [uint32][IO.FileAttributes]::Directory) -ne 0 -or
        ($handleAttributes -band [uint32][IO.FileAttributes]::ReparsePoint) -ne 0) {
        $handle.Dispose()
        throw [IO.InvalidDataException]::new(
            "Collection-neutrality host surface '$RelativePath' is a reparse point.")
    }
    $stream = [IO.FileStream]::new($handle, [IO.FileAccess]::Read)
    try {
        $attributes = [IO.File]::GetAttributes($Path)
        $length = $stream.Length
        if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
            $length -lt 0 -or
            $length -gt 4194304 -or
            $length -gt [int]::MaxValue) {
            throw [IO.InvalidDataException]::new(
                "Collection-neutrality host surface '$RelativePath' violates the fixed file bound.")
        }
        $bytes = [byte[]]::new([int]$length)
        $offset = 0
        while ($offset -lt $bytes.Length) {
            $read = $stream.Read(
                $bytes,
                $offset,
                $bytes.Length - $offset)
            if ($read -eq 0) {
                throw [IO.EndOfStreamException]::new(
                    "Collection-neutrality host surface '$RelativePath' ended before its held length.")
            }
            $offset += $read
        }
        if ($stream.Length -ne $length) {
            throw [IO.InvalidDataException]::new(
                "Collection-neutrality host surface '$RelativePath' changed while its snapshot was read.")
        }
        $text = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        return [pscustomobject]@{
            Bytes = $bytes
            Text = $text
            Sha256 = [Convert]::ToHexString(
                [Security.Cryptography.SHA256]::HashData($bytes)).
                ToLowerInvariant()
            SizeBytes = $length
        }
    } finally {
        $stream.Dispose()
    }
}

function Invoke-RSTerminalEngineCollectionNeutralityScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [AllowEmptyCollection()]
        [object[]]$GovernanceObject = @(),

        [AllowEmptyCollection()]
        [object[]]$Descriptor = @(),

        [AllowEmptyCollection()]
        [object[]]$FrozenInput = @(),

        [AllowEmptyCollection()]
        [string[]]$AdditionalLiteral = @(),

        [AllowEmptyCollection()]
        [object[]]$Exception = @(),

        [AllowNull()]
        [object]$PolicyRecord = $null,

        [switch]$ValidateBaselineProof
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).
        TrimEnd([IO.Path]::DirectorySeparatorChar)
    if ($null -ne $PolicyRecord) {
        if ($PolicyRecord -isnot [Collections.IDictionary]) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality policy must be one object.')
        }
        Assert-RSCollectionNeutralityExactProperties `
            -Value $PolicyRecord `
            -Expected @(
                'formatId',
                'version',
                'surfacePaths',
                'baselineAdditionalLiterals',
                'exceptions',
                'baselineProof') `
            -Label 'Collection-neutrality policy'
        if ([string]$PolicyRecord['formatId'] -cne
                'red-salamander-terminal-engine-collection-neutrality' -or
            [int]$PolicyRecord['version'] -ne 1) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality policy identity is invalid.')
        }
        $policySurface = [string[]]@($PolicyRecord['surfacePaths'])
        Assert-RSCollectionNeutralityOrdinalStrings `
            -Value $policySurface `
            -Label 'Collection-neutrality policy surfacePaths'
        if (-not [Linq.Enumerable]::SequenceEqual(
                $policySurface,
                $script:RSCollectionNeutralitySurface)) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality policy surfacePaths differ from the executable scanner surface.')
        }
        $baselineAdditional = [string[]]@(
            $PolicyRecord['baselineAdditionalLiterals'])
        Assert-RSCollectionNeutralityOrdinalStrings `
            -Value $baselineAdditional `
            -Label 'Collection-neutrality policy baselineAdditionalLiterals'
        $AdditionalLiteral = [string[]]@(
            $baselineAdditional + $AdditionalLiteral)
        if ($Exception.Count -ne 0) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality exceptions must come only from the frozen policy record.')
        }
        $Exception = [object[]]@($PolicyRecord['exceptions'])
    }
    $forbidden = [string[]]@(
        Get-RSTerminalEngineCollectionForbiddenLiteral `
            -GovernanceObject $GovernanceObject `
            -Descriptor $Descriptor `
            -FrozenInput $FrozenInput `
            -AdditionalLiteral $AdditionalLiteral)
    if ($forbidden.Count -eq 0) {
        throw [IO.InvalidDataException]::new(
            'Collection-neutrality scan has no frozen forbidden literals.')
    }

    $exceptionByKey =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::Ordinal)
    $previousExceptionKey = $null
    foreach ($record in $Exception) {
        if ($record -isnot [Collections.IDictionary]) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality exception must be one object.')
        }
        Assert-RSCollectionNeutralityExactProperties `
            -Value $record `
            -Expected @('path', 'line', 'literal', 'lineSha256') `
            -Label 'Collection-neutrality exception'
        $path = [string]$record['path']
        $line = [int]$record['line']
        $literal = ([string]$record['literal']).ToLowerInvariant()
        $lineSha256 = [string]$record['lineSha256']
        if ($path -cnotin $script:RSCollectionNeutralitySurface -or
            $line -lt 1 -or
            $literal -cnotin $forbidden -or
            $lineSha256 -cnotmatch '^[0-9a-f]{64}$') {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality exception is not an exact frozen path, line, hash, and forbidden literal.')
        }
        $key = "$path`0$line`0$literal"
        $sortKey = "$path`0$($line.ToString(
                'D10',
                [Globalization.CultureInfo]::InvariantCulture))`0$literal"
        if ($null -ne $previousExceptionKey -and
            [StringComparer]::Ordinal.Compare(
                $previousExceptionKey,
                $sortKey) -ge 0) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality exceptions must be sorted and unique.')
        }
        $previousExceptionKey = $sortKey
        if (-not $exceptionByKey.TryAdd($key, $record)) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality exception is duplicated.')
        }
    }
    $usedExceptions =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
    $surfaceRecords = [Collections.Generic.List[object]]::new()
    $violations = [Collections.Generic.List[string]]::new()

    foreach ($relative in $script:RSCollectionNeutralitySurface) {
        $absolute = [IO.Path]::GetFullPath((
                Join-Path $root ($relative.Replace('/', '\'))))
        if (-not $absolute.StartsWith(
                $root + [IO.Path]::DirectorySeparatorChar,
                [StringComparison]::OrdinalIgnoreCase) -or
            -not [IO.File]::Exists($absolute)) {
            throw [IO.FileNotFoundException]::new(
                "Required collection-neutrality host surface '$relative' is absent.",
                $absolute)
        }
        $surfaceSnapshot = Read-RSCollectionNeutralitySurfaceSnapshot `
            -Path $absolute `
            -RelativePath $relative
        $text = [string]$surfaceSnapshot.Text
        $compositions = [object[]]@(
            Get-RSCollectionNeutralityStringCompositions `
                -Text $text `
                -RelativePath $relative)
        $lines = [Text.RegularExpressions.Regex]::Split(
            $text,
            '\r?\n',
            [Text.RegularExpressions.RegexOptions]::CultureInvariant)
        foreach ($literal in $forbidden) {
            $compact = $literal -replace '[^a-z0-9]', ''
            foreach ($composition in $compositions) {
                if ([string]$composition.Compact -cne $compact) {
                    continue
                }
                $compositionLine = [int]$composition.Line
                $compositionKey =
                    "$relative`0$compositionLine`0$literal"
                $exceptionRecord = $null
                if ($exceptionByKey.TryGetValue(
                        $compositionKey,
                        [ref]$exceptionRecord) -and
                    [string]$exceptionRecord['lineSha256'] -ceq
                        (Get-RSCollectionNeutralitySha256Text `
                            -Value $lines[$compositionLine - 1])) {
                    [void]$usedExceptions.Add($compositionKey)
                    continue
                }
                $violations.Add(
                    "$relative`:$compositionLine`: constructed frozen collection literal")
            }
            for ($lineIndex = 0; $lineIndex -lt $lines.Count; ++$lineIndex) {
                $identifierFacts =
                    Get-RSCollectionNeutralityIdentifierFacts `
                        -Text $lines[$lineIndex]
                if (-not (Test-RSCollectionNeutralityLiteralHit `
                        -Line $lines[$lineIndex] `
                        -Literal $literal `
                        -IdentifierFacts $identifierFacts)) {
                    continue
                }
                $lineNumber = $lineIndex + 1
                $key = "$relative`0$lineNumber`0$literal"
                $exceptionRecord = $null
                if ($exceptionByKey.TryGetValue($key, [ref]$exceptionRecord) -and
                    [string]$exceptionRecord['lineSha256'] -ceq
                        (Get-RSCollectionNeutralitySha256Text `
                            -Value $lines[$lineIndex])) {
                    [void]$usedExceptions.Add($key)
                    continue
                }
                $violations.Add(
                    "$relative`:$lineNumber`: forbidden frozen collection literal")
            }
        }
        $surfaceRecords.Add([ordered]@{
                path = $relative
                sha256 = [string]$surfaceSnapshot.Sha256
                sizeBytes = $surfaceSnapshot.SizeBytes.ToString(
                    [Globalization.CultureInfo]::InvariantCulture)
            })
    }

    if ($usedExceptions.Count -ne $exceptionByKey.Count) {
        throw [IO.InvalidDataException]::new(
            'Collection-neutrality exception set contains a stale or nonmatching record.')
    }
    if ($violations.Count -ne 0) {
        throw [IO.InvalidDataException]::new(
            "Collection-neutrality scan failed:`n$($violations -join "`n")")
    }

    $exceptionPayload = [object[]]@(
        $Exception | Sort-Object `
            @{ Expression = { [string]$_['path'] } },
            @{ Expression = { [int]$_['line'] } },
            @{ Expression = { [string]$_['literal'] } })
    $payload = [ordered]@{
        algorithmId = 'terminal-engine-collection-neutrality-v1'
        surfaceSetSha256 = Get-RSJcsSha256 `
            -InputObject ([object[]]$surfaceRecords.ToArray())
        forbiddenLiteralSetSha256 = Get-RSJcsSha256 `
            -InputObject ([string[]]$forbidden)
        exceptionSetSha256 = Get-RSJcsSha256 `
            -InputObject $exceptionPayload
        findingCount = 0
        status = 'verified'
    }
    $proof = [ordered]@{}
    foreach ($entry in $payload.GetEnumerator()) {
        $proof[$entry.Key] = $entry.Value
    }
    $proof['proofSha256'] = Get-RSJcsSha256 -InputObject $payload
    if ($null -ne $PolicyRecord) {
        $baselineProof = $PolicyRecord['baselineProof']
        if ($baselineProof -isnot [Collections.IDictionary]) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality baselineProof must be one object.')
        }
        Assert-RSCollectionNeutralityExactProperties `
            -Value $baselineProof `
            -Expected @(
                'algorithmId',
                'surfaceSetSha256',
                'forbiddenLiteralSetSha256',
                'exceptionSetSha256',
                'findingCount',
                'status',
                'proofSha256') `
            -Label 'Collection-neutrality baselineProof'
        if ([string]$baselineProof['surfaceSetSha256'] -cne
            [string]$payload.surfaceSetSha256) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality surface identity differs from the frozen policy baseline.')
        }
    }
    if ($ValidateBaselineProof) {
        if ($null -eq $PolicyRecord) {
            throw [IO.InvalidDataException]::new(
                'Baseline-proof validation requires a policy record.')
        }
        if ((Get-RSJcsSha256 -InputObject $baselineProof) -cne
                (Get-RSJcsSha256 -InputObject $proof)) {
            throw [IO.InvalidDataException]::new(
                'Collection-neutrality baseline proof differs from the current exact scan.')
        }
    }
    return [pscustomobject]$proof
}

Export-ModuleMember -Function @(
    'Get-RSTerminalEngineCollectionNeutralitySurface',
    'Get-RSTerminalEngineCollectionForbiddenLiteral',
    'Invoke-RSTerminalEngineCollectionNeutralityScan')
