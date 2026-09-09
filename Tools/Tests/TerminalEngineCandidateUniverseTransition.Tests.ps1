Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngineEvaluator.psm1'
$evidenceModulePath = Join-Path $repoRoot 'Tools\TerminalEvidence.psm1'
$neutralityModulePath = Join-Path $repoRoot (
    'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1')
Import-Module $evidenceModulePath -Force -Global
$neutralityModule = Import-Module `
    $neutralityModulePath `
    -Force `
    -Global `
    -PassThru
$script:RSTestNeutralitySurface = [string[]]@(
    & $neutralityModule {
        Get-RSTerminalEngineCollectionNeutralitySurface
    })
if ($script:RSTestNeutralitySurface.Count -ne 51) {
    throw 'The transition-test neutrality surface must contain exactly 51 paths.'
}
Import-Module $modulePath -Force -Global
$script:RSTestRound4Superseded =
    Test-Path -LiteralPath (Join-Path $repoRoot 'Plugins\Terminal')

function Assert-RSTestTransitionThrows {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Action,

        [Parameter(Mandatory)]
        [string]$MessageFragment
    )

    $caught = $null
    try {
        & $Action
    }
    catch [System.Exception] {
        $caught = $_.Exception
    }
    if ($null -eq $caught) {
        throw "Expected transition validation to throw '$MessageFragment'."
    }
    if (-not $caught.Message.Contains(
            $MessageFragment,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw (
            "Expected transition validation message containing '$MessageFragment', " +
            "got '$($caught.Message)'.")
    }
}

function Read-RSTestJson {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return [Text.UTF8Encoding]::new($false, $true).GetString(
        [IO.File]::ReadAllBytes($Path)) |
        ConvertFrom-Json `
            -AsHashtable `
            -Depth 100 `
            -NoEnumerate `
            -DateKind String
}

function Invoke-RSTestEvaluatorPrivate {
    param(
        [Parameter(Mandatory)]
        [string]$Command,

        [Parameter(Mandatory)]
        [string]$Root
    )

    $module = Get-Module -Name TerminalEngineEvaluator
    if ($null -eq $module) {
        throw 'TerminalEngineEvaluator must be imported before private test invocation.'
    }
    return & $module {
        param($CommandName, $RepoRoot)
        & $CommandName -RepoRoot $RepoRoot
    } $Command $Root
}

function New-RSTestTransitionRepository {
    param(
        [Parameter(Mandatory)]
        [string]$Root
    )

    $transitionRelative =
        'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json'
    $transitionSource = Join-Path $repoRoot $transitionRelative
    if (-not [IO.File]::Exists($transitionSource)) {
        throw 'The canonical Round 3 transition must exist before its tests run.'
    }
    $transition = Read-RSTestJson -Path $transitionSource
    $relativePaths =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($relativePath in @(
            'Tools\TerminalEngineEvaluator.psm1',
            'Tools\TerminalEvidence.psm1',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.schema.json',
            'Specs\Terminal\TerminalEngineCandidateUniverseTransition.schema.json',
            'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v3.schema.json',
            'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v4.schema.json',
            'Specs\Terminal\TerminalEngineCollectionNeutrality.schema.json')) {
        [void]$relativePaths.Add($relativePath)
    }
    foreach ($inputRecord in @($transition['frozenInputs'])) {
        [void]$relativePaths.Add(
            ([string]$inputRecord['relativePath']).Replace('/', '\'))
    }
    foreach ($schemaRecord in $transition['governanceSchemas'].Values) {
        [void]$relativePaths.Add(
            ([string]$schemaRecord['relativePath']).Replace('/', '\'))
    }
    foreach ($surfacePath in $script:RSTestNeutralitySurface) {
        [void]$relativePaths.Add($surfacePath.Replace('/', '\'))
    }

    foreach ($relativePath in $relativePaths) {
        $source = Join-Path $repoRoot $relativePath
        if (-not [IO.File]::Exists($source)) {
            throw "Required Round 3 test input is absent: $relativePath"
        }
        $destination = Join-Path $Root $relativePath
        $null = [IO.Directory]::CreateDirectory(
            [IO.Path]::GetDirectoryName($destination))
        [IO.File]::Copy($source, $destination)
    }
    return $Root
}

function Edit-RSTestTransition {
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [scriptblock]$Edit,

        [switch]$RebindProposal,

        [switch]$Pretty
    )

    Import-Module $evidenceModulePath -Force -Global
    $path = Join-Path $Root (
        'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json')
    $transition = Read-RSTestJson -Path $path
    & $Edit $transition
    if ($RebindProposal) {
        $proposal = [ordered]@{}
        foreach ($entry in $transition.GetEnumerator()) {
            if ([string]$entry.Key -cnotin @('proposalSha256', 'approvals')) {
                $proposal[[string]$entry.Key] = $entry.Value
            }
        }
        $transition['proposalSha256'] =
            Get-RSJcsSha256 -InputObject $proposal
        foreach ($approval in @($transition['approvals'])) {
            $approval['proposalSha256'] = $transition['proposalSha256']
        }
    }
    $bytes = if ($Pretty) {
        [Text.UTF8Encoding]::new($false).GetBytes(
            ($transition | ConvertTo-Json -Depth 100))
    }
    else {
        ConvertTo-RSJcsUtf8Bytes -InputObject $transition
    }
    [IO.File]::WriteAllBytes($path, $bytes)
}

function New-RSTestRound4TransitionRepository {
    param(
        [Parameter(Mandatory)]
        [string]$Root
    )

    $null = New-RSTestTransitionRepository -Root $Root
    $source = Join-Path $repoRoot (
        'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json')
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
        throw "Missing canonical Round 4 candidate universe: $source"
    }
    $target = Join-Path $Root (
        'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json')
    Copy-Item -LiteralPath $source -Destination $target
    return $Root
}

function Edit-RSTestRound4Transition {
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [scriptblock]$Edit,

        [switch]$RebindProposal,

        [switch]$Pretty
    )

    Import-Module $evidenceModulePath -Force -Global
    $path = Join-Path $Root (
        'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json')
    $transition = Read-RSTestJson -Path $path
    & $Edit $transition
    if ($RebindProposal) {
        $proposal = [ordered]@{}
        foreach ($entry in $transition.GetEnumerator()) {
            if ([string]$entry.Key -cnotin @('proposalSha256', 'approvals')) {
                $proposal[[string]$entry.Key] = $entry.Value
            }
        }
        $transition['proposalSha256'] =
            Get-RSJcsSha256 -InputObject $proposal
        foreach ($approval in @($transition['approvals'])) {
            $approval['proposalSha256'] = $transition['proposalSha256']
        }
    }
    $bytes = if ($Pretty) {
        [Text.UTF8Encoding]::new($false).GetBytes(
            ($transition | ConvertTo-Json -Depth 100))
    }
    else {
        ConvertTo-RSJcsUtf8Bytes -InputObject $transition
    }
    [IO.File]::WriteAllBytes($path, $bytes)
}

Describe 'TerminalEngine candidate-universe Round 3 transition' {
    It 'force-reloads only exact authenticated modules after round-4 preflight' {
        $tokens = $null
        $parseErrors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile(
            $modulePath,
            [ref]$tokens,
            [ref]$parseErrors)
        @($parseErrors).Count | Should Be 0
        $forcedImports = [object[]]@($ast.FindAll({
                    param($node)
                    if ($node -isnot
                        [Management.Automation.Language.CommandAst] -or
                        $node.GetCommandName() -cne 'Import-Module') {
                        return $false
                    }
                    return @($node.CommandElements | Where-Object {
                                $_ -is
                                    [Management.Automation.Language.CommandParameterAst] -and
                                $_.ParameterName -ieq 'Force'
                            }).Count -gt 0
                }, $true))
        $forcedImports.Count | Should Be 1
        $owner = $forcedImports[0].Parent
        while ($null -ne $owner -and
            $owner -isnot
                [Management.Automation.Language.FunctionDefinitionAst]) {
            $owner = $owner.Parent
        }
        $null -eq $owner | Should Be $false
        $owner.Name | Should Be 'Import-RSExactModule'
        @($forcedImports[0].CommandElements | Where-Object {
                $_ -is
                    [Management.Automation.Language.CommandParameterAst] -and
                $_.ParameterName -ieq 'PassThru'
            }).Count | Should Be 1
        $ownerText = [string]$owner.Extent.Text
        $ownerText.Contains(
            '[IO.Path]::GetFullPath([string]$_.Path)',
            [StringComparison]::Ordinal) | Should Be $true
        $ownerText.Contains(
            '$matching.Count -ne 1',
            [StringComparison]::Ordinal) | Should Be $true

        $currentReader = [object[]]@($ast.FindAll({
                    param($node)
                    $node -is
                        [Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -ceq 'Read-RSCurrentCandidateUniverseCore'
                }, $true))
        $currentReader.Count | Should Be 1
        $readerText = [string]$currentReader[0].Extent.Text
        $authorityIndex = $readerText.IndexOf(
            'Assert-RSRound4AuthorityPreflight',
            [StringComparison]::Ordinal)
        $authenticationIndex = $readerText.IndexOf(
            'Assert-RSRound4AuthenticationPreflight',
            [StringComparison]::Ordinal)
        $importIndex = $readerText.IndexOf(
            'Import-RSAuthenticatedCollectionNeutralityModules',
            [StringComparison]::Ordinal)
        ($authorityIndex -ge 0) | Should Be $true
        ($authenticationIndex -gt $authorityIndex) | Should Be $true
        ($importIndex -gt $authenticationIndex) | Should Be $true
    }

    It 'preserves caller-global dependency exports after a historical round-3 read' {
        $expectedExports = [string[]]@(
            'ConvertTo-RSJcsUtf8Bytes',
            'Get-RSTerminalEngineCollectionNeutralitySurface',
            'Invoke-RSTerminalEngineCollectionNeutralityScan')
        foreach ($commandName in $expectedExports) {
            (Get-Command $commandName -ErrorAction Stop).Name |
                Should Be $commandName
        }

        $historical = Read-RSHistoricalRound3CandidateUniverse -RepoRoot $repoRoot
        $historical['version'] | Should Be 3

        foreach ($commandName in $expectedExports) {
            (Get-Command $commandName -ErrorAction Stop).Name |
                Should Be $commandName
        }
    }

    It 'keeps direct v2 validation strict while v3 consumes v2 as authenticated history' {
        Assert-RSTestTransitionThrows `
            -Action {
                Invoke-RSTestEvaluatorPrivate `
                    -Command 'Read-RSRound2CandidateUniverse' `
                    -Root $repoRoot
            } `
            -MessageFragment 'frozen input'

        $historical = Invoke-RSTestEvaluatorPrivate `
            -Command 'Read-RSHistoricalRound2CandidateUniverse' `
            -Root $repoRoot
        $historical['version'] | Should Be 2

        $historical = Read-RSHistoricalRound3CandidateUniverse -RepoRoot $repoRoot
        $historical['version'] | Should Be 3
    }

    It 'resolves the exact approved v3 transition over immutable v2 and v1' {
        $universe = Read-RSHistoricalRound3CandidateUniverse -RepoRoot $repoRoot
        $universe['version'] | Should Be 3
        [string[]]@(
            @($universe['seedCandidates']) |
                Where-Object { $null -ne $_['allowedNominationId'] } |
                ForEach-Object { [string]$_['allowedNominationId'] }) |
            Should Be @(
                'contour-rs-v3',
                'ghostty-rs-v4',
                'wezterm-rs-v3')
        $universe['adaptationPolicy']['hardCaps']['ownerDays'] | Should Be 60
    }

    It 'never downgrades a missing named universe version' {
        $missingRound4 = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'missing-round4')
        Remove-Item -LiteralPath (Join-Path $missingRound4 (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json'))
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSCurrentCandidateUniverse -RepoRoot $missingRound4
            } `
            -MessageFragment 'round-4 candidate universe is missing'

        $missingRound3 = New-RSTestTransitionRepository -Root (
            Join-Path $TestDrive 'missing-round3')
        Remove-Item -LiteralPath (Join-Path $missingRound3 (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json'))
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSHistoricalRound3CandidateUniverse -RepoRoot $missingRound3
            } `
            -MessageFragment 'round-3 candidate-universe transition is missing'

        $missingRound2 = New-RSTestTransitionRepository -Root (
            Join-Path $TestDrive 'missing-round2')
        Remove-Item -LiteralPath (Join-Path $missingRound2 (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json'))
        Assert-RSTestTransitionThrows `
            -Action {
                Invoke-RSTestEvaluatorPrivate `
                    -Command 'Read-RSHistoricalRound2CandidateUniverse' `
                    -Root $missingRound2
            } `
            -MessageFragment 'round-2 candidate-universe transition is missing'
    }

    It 'preserves the exact historical 60-input closure and invalid closeout' {
        $transitionPath = Join-Path $repoRoot (
            'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json')
        $transition = Read-RSTestJson -Path $transitionPath
        @($transition['frozenInputs']).Count | Should Be 60
        [string]$transition['roundCloseout']['status'] |
            Should Be 'invalid-no-selection'

        [string]$transition['frozenInputs'][7]['sha256'] |
            Should Be (
                '174a064cd879cbc0d82e1953437bca82e7dfacdd777132f738a761f24d9f3285')
    }

    It 'rejects duplicate reviewer identities even when the schema passes' {
        $fixture = New-RSTestTransitionRepository -Root (
            Join-Path $TestDrive 'duplicate-reviewer')
        Edit-RSTestTransition -Root $fixture -Edit {
            param($transition)
            $transition['approvals'][1]['reviewerId'] =
                $transition['approvals'][0]['reviewerId']
        }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSHistoricalRound3CandidateUniverse -RepoRoot $fixture
            } `
            -MessageFragment 'reviewers'
    }

    It 'rejects a stale proposal digest' {
        $fixture = New-RSTestTransitionRepository -Root (
            Join-Path $TestDrive 'stale-proposal')
        Edit-RSTestTransition -Root $fixture -Edit {
            param($transition)
            $transition['frozenUtc'] = '2026-07-30T19:04:49Z'
        }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSHistoricalRound3CandidateUniverse -RepoRoot $fixture
            } `
            -MessageFragment 'proposal digest'
    }

    It 'rejects noncanonical transition bytes' {
        $fixture = New-RSTestTransitionRepository -Root (
            Join-Path $TestDrive 'noncanonical')
        Edit-RSTestTransition -Root $fixture -Pretty -Edit {
            param($transition)
        }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSHistoricalRound3CandidateUniverse -RepoRoot $fixture
            } `
            -MessageFragment 'exact RFC 8785 JCS'
    }

    It 'rejects a cross-paired nomination tuple at the closed schema' {
        $fixture = New-RSTestTransitionRepository -Root (
            Join-Path $TestDrive 'cross-pair')
        Edit-RSTestTransition -Root $fixture -RebindProposal -Edit {
            param($transition)
            $transition['policy']['nominationOverrides'][0][
                'priorNominationId'] = 'ghostty-rs-v3'
        }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSHistoricalRound3CandidateUniverse -RepoRoot $fixture
            } `
            -MessageFragment 'does not satisfy TerminalEngine evidence schema'
    }

    It 'rejects a weakened premeasurement rule even with a rebound proposal' {
        $fixture = New-RSTestTransitionRepository -Root (
            Join-Path $TestDrive 'weakened-rule')
        Edit-RSTestTransition -Root $fixture -RebindProposal -Edit {
            param($transition)
            $transition['policy']['preMeasurementRules'][0] =
                'Collection may start before candidate-set approval.'
        }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSHistoricalRound3CandidateUniverse -RepoRoot $fixture
            } `
            -MessageFragment 'does not satisfy TerminalEngine evidence schema'
    }
}

Describe 'TerminalEngine candidate-universe Round 4 transition' {
    It 'preserves the exact approved round-3 artifact as immutable history' {
        (Get-FileHash `
            -LiteralPath (Join-Path $repoRoot (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json')) `
            -Algorithm SHA256).Hash.ToLowerInvariant() |
            Should Be (
                '5c38444417ad246148928206c95acef9d84238d5f3f8116bb88fdc5fa115f63e')
    }

    It 'carries policy through one closed lexical mapping and adds only chronology' {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'v4-policy-equivalence')
        $v3 = Read-RSTestJson -Path (Join-Path $fixture (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json'))
        $v4 = Read-RSTestJson -Path (Join-Path $fixture (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json'))
        $expected = [string[]]@(
            [string[]]$v3.policy.preMeasurementRules +
            @(
                'Current-universe approval opens the round: every source, adaptation-rejection, and admitted-candidate approval is an exact canonical UTC-second instant strictly after the latest universe approval and strictly before one byte-identical CandidateReview/CandidateAssessment/CandidateSet freeze instant. Equal boundaries, anticipatory approvals, mismatched freezes, and provider paths that skip the complete governance closure are invalid before any collection effect.'
            ))
        $expected[0] = $expected[0] -creplace (
            'round-3 native'), 'round-4 native'
        $expected[2] = $expected[2] -creplace (
            'Round 3 has a closed'), (
            'Round 4 retains the round-3 closed')
        $expected[6] = $expected[6] -creplace (
            'round 3 has no'), 'round 4 has no'
        ([string[]]$v4.policy.preMeasurementRules -join "`0") |
            Should Be ($expected -join "`0")
        ([string[]]@(
                $v4.conformanceCorrections |
                    ForEach-Object { [string]$_.correctionId }) -join "`0") |
            Should Be (
                'cross-repository-adaptation-ledger-closure' + "`0" +
                'candidate-governance-chronology-closure')
        @($v4.frozenInputs).Count | Should Be 65
        @($v4.policy.nominationCarryForward).Count | Should Be 3
    }

    It 'resolves an authenticated v4 over exact historical v3' -Skip:$script:RSTestRound4Superseded {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'valid-v4')
        $universe = Read-RSCurrentCandidateUniverse -RepoRoot $fixture
        $universe.version | Should Be 4
        [string[]]@(
            @($universe.seedCandidates) |
                Where-Object { $null -ne $_.allowedNominationId } |
                ForEach-Object { [string]$_.allowedNominationId }) |
            Should Be @(
                'contour-rs-v3',
                'ghostty-rs-v4',
                'wezterm-rs-v3')
    }

    It 'locks every real-v4 static closure class before authority is interpreted' {
        $relativePaths = [string[]]@(
            'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v4.schema.json',
            'Specs\Terminal\TerminalEngineRequirements.v1.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.schema.json',
            'Tools\TerminalEngine\TerminalEngineArchive.psm1')
        foreach ($relativePath in $relativePaths) {
            $caseName = ($relativePath -replace '[^A-Za-z0-9]+', '-').
                Trim('-')
            $fixture = New-RSTestRound4TransitionRepository -Root (
                Join-Path $TestDrive "v4-writer-$caseName")
            $path = Join-Path $fixture $relativePath
            $writer = [IO.FileStream]::new(
                $path,
                [IO.FileMode]::Open,
                [IO.FileAccess]::Write,
                [IO.FileShare]::ReadWrite)
            try {
                Assert-RSTestTransitionThrows `
                    -Action {
                        Read-RSCurrentCandidateUniverse -RepoRoot $fixture
                    } `
                    -MessageFragment 'Cannot acquire immutable read snapshot'
            }
            finally {
                $writer.Dispose()
            }
        }
    }

    It 'rejects a stable modified frozen module before its top-level code runs' {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'v4-modified-module-marker')
        $marker = Join-Path $fixture 'modified-module-executed.marker'
        $escapedMarker = $marker.Replace("'", "''")
        [IO.File]::AppendAllText(
            (Join-Path $fixture (
                'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1')),
            "`n[IO.File]::WriteAllText('$escapedMarker','executed')`n",
            [Text.UTF8Encoding]::new($false))

        Assert-RSTestTransitionThrows `
            -Action { Read-RSCurrentCandidateUniverse -RepoRoot $fixture } `
            -MessageFragment 'frozen input'
        [IO.File]::Exists($marker) | Should Be $false
    }

    It 'uses a literal-two trust root before a self-declared schema or module can run' {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'v4-self-declared-one-approval')
        $marker = Join-Path $fixture 'self-declared-module-executed.marker'
        $escapedMarker = $marker.Replace("'", "''")
        $moduleRelative =
            'Tools/TerminalEngine/TerminalEngineCollectionNeutrality.psm1'
        $modulePath = Join-Path $fixture $moduleRelative.Replace('/', '\')
        [IO.File]::AppendAllText(
            $modulePath,
            "`n[IO.File]::WriteAllText('$escapedMarker','executed')`n",
            [Text.UTF8Encoding]::new($false))

        $requirementsRelative =
            'Specs/Terminal/TerminalEngineRequirements.v1.json'
        $requirementsPath =
            Join-Path $fixture $requirementsRelative.Replace('/', '\')
        $requirements = Read-RSTestJson -Path $requirementsPath
        $requirements.selection.reviewerApprovalCount = 1
        [IO.File]::WriteAllBytes(
            $requirementsPath,
            (ConvertTo-RSJcsUtf8Bytes -InputObject $requirements))

        $schemaRelative =
            'Specs/Terminal/TerminalEngineCandidateUniverseTransition.v4.schema.json'
        $schemaPath = Join-Path $fixture $schemaRelative.Replace('/', '\')
        $permissiveSchema = [ordered]@{
            '$schema' = 'https://json-schema.org/draft/2020-12/schema'
            type = 'object'
        }
        [IO.File]::WriteAllBytes(
            $schemaPath,
            (ConvertTo-RSJcsUtf8Bytes -InputObject $permissiveSchema))

        $shaByRelativePath = @{
            $moduleRelative = (
                Get-FileHash -LiteralPath $modulePath -Algorithm SHA256).
                Hash.ToLowerInvariant()
            $requirementsRelative = (
                Get-FileHash -LiteralPath $requirementsPath -Algorithm SHA256).
                Hash.ToLowerInvariant()
            $schemaRelative = (
                Get-FileHash -LiteralPath $schemaPath -Algorithm SHA256).
                Hash.ToLowerInvariant()
        }
        Edit-RSTestRound4Transition `
            -Root $fixture `
            -RebindProposal `
            -Edit {
                param($transition)
                $transition.approvals =
                    [object[]]@($transition.approvals[0])
                $transition.governanceSchemas.round4Transition.sha256 =
                    $shaByRelativePath[$schemaRelative]
                foreach ($relativePath in $shaByRelativePath.Keys) {
                    $matches = [object[]]@(
                        $transition.frozenInputs |
                            Where-Object {
                                [string]$_.relativePath -ceq $relativePath
                            })
                    if ($matches.Count -ne 1) {
                        throw "Round-4 test input '$relativePath' is not frozen exactly once."
                    }
                    $matches[0].sha256 =
                        $shaByRelativePath[$relativePath]
                }
            }

        Assert-RSTestTransitionThrows `
            -Action { Read-RSCurrentCandidateUniverse -RepoRoot $fixture } `
            -MessageFragment (
                'requires exactly 2 independent approvals at the evaluator trust root')
        [IO.File]::Exists($marker) | Should Be $false
    }

    It 'authenticates the historical v2 file and v1 reader schema under real v4' -Skip:$script:RSTestRound4Superseded {
        $v2Fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'v4-modified-v2')
        [IO.File]::AppendAllText(
            (Join-Path $v2Fixture (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json')),
            ' ',
            [Text.UTF8Encoding]::new($false))
        Assert-RSTestTransitionThrows `
            -Action { Read-RSCurrentCandidateUniverse -RepoRoot $v2Fixture } `
            -MessageFragment 'round-3 base digest'

        $v1SchemaFixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'v4-modified-v1-schema')
        [IO.File]::AppendAllText(
            (Join-Path $v1SchemaFixture (
                'Specs\Terminal\TerminalEngineCandidateUniverse.schema.json')),
            ' ',
            [Text.UTF8Encoding]::new($false))
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSCurrentCandidateUniverse -RepoRoot $v1SchemaFixture
            } `
            -MessageFragment "governance schema 'baseUniverse'"
    }

    It 'rejects deleting v4 instead of downgrading to historical v3' {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'missing-v4')
        Remove-Item -LiteralPath (Join-Path $fixture (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json'))
        Assert-RSTestTransitionThrows `
            -Action { Read-RSCurrentCandidateUniverse -RepoRoot $fixture } `
            -MessageFragment 'round-4 candidate universe is missing'
    }

    It 'rejects every lexically noncanonical universe timestamp after full proposal rebinding' {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'bad-v4-time')
        $path = Join-Path $fixture (
            'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json')
        $original = [IO.File]::ReadAllBytes($path)
        foreach ($badValue in @(
                '2026-07-31T10:00:00+00:00',
                '2026-07-31T10:00:00.000Z',
                '2026-07-31T10:00:00z',
                ' 2026-07-31T10:00:00Z',
                '2026-07-31T10:00:00Z ')) {
            [IO.File]::WriteAllBytes($path, $original)
            Edit-RSTestRound4Transition `
                -Root $fixture `
                -RebindProposal `
                -Edit {
                    param($transition)
                    $transition.frozenUtc = $badValue
                }
            Assert-RSTestTransitionThrows `
                -Action {
                    Read-RSCurrentCandidateUniverse -RepoRoot $fixture
                } `
                -MessageFragment (
                    'must be an exact RFC 3339 UTC-second string')
        }
    }

    It 'rejects an impossible universe calendar instant after proposal rebinding' {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'impossible-v4-time')
        Edit-RSTestRound4Transition `
            -Root $fixture `
            -RebindProposal `
            -Edit {
                param($transition)
                $transition.frozenUtc = '2026-02-30T10:00:00Z'
            }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSCurrentCandidateUniverse -RepoRoot $fixture
            } `
            -MessageFragment 'must be an exact RFC 3339 UTC-second string'
    }

    It 'rejects lexically noncanonical universe approval timestamps' {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'bad-v4-approval-time')
        $path = Join-Path $fixture (
            'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json')
        $original = [IO.File]::ReadAllBytes($path)
        foreach ($badValue in @(
                '2026-07-31T10:01:00+00:00',
                '2026-07-31T10:01:00.000Z',
                '2026-07-31T10:01:00z')) {
            [IO.File]::WriteAllBytes($path, $original)
            Edit-RSTestRound4Transition `
                -Root $fixture `
                -Edit {
                    param($transition)
                    $transition.approvals[0].reviewedUtc = $badValue
                }
            Assert-RSTestTransitionThrows `
                -Action {
                    Read-RSCurrentCandidateUniverse -RepoRoot $fixture
                } `
                -MessageFragment (
                    'must be an exact RFC 3339 UTC-second string')
        }
    }

    It 'rejects an impossible universe approval calendar instant' {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'impossible-v4-approval-time')
        Edit-RSTestRound4Transition `
            -Root $fixture `
            -Edit {
                param($transition)
                $transition.approvals[0].reviewedUtc =
                    '2026-02-30T10:01:00Z'
            }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSCurrentCandidateUniverse -RepoRoot $fixture
            } `
            -MessageFragment 'must be an exact RFC 3339 UTC-second string'
    }

    It 'rejects a fully rebound one-byte v4 policy weakening' -Skip:$script:RSTestRound4Superseded {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'weakened-v4-rule')
        Edit-RSTestRound4Transition `
            -Root $fixture `
            -RebindProposal `
            -Edit {
                param($transition)
                $transition.policy.preMeasurementRules[10] =
                    $transition.policy.preMeasurementRules[10].Replace(
                        'strictly after',
                        'at or after')
            }
        Assert-RSTestTransitionThrows `
            -Action { Read-RSCurrentCandidateUniverse -RepoRoot $fixture } `
            -MessageFragment 'does not satisfy TerminalEngine evidence schema'
    }

    It 'rejects rebound v4 base, governance-schema, or frozen-input digests' {
        $baseFixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'bad-v4-base')
        Edit-RSTestRound4Transition `
            -Root $baseFixture `
            -RebindProposal `
            -Edit {
                param($transition)
                $transition.baseUniverse.sha256 = '1' * 64
            }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSCurrentCandidateUniverse -RepoRoot $baseFixture
            } `
            -MessageFragment 'base digest'

        $schemaFixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'bad-v4-schema')
        Edit-RSTestRound4Transition `
            -Root $schemaFixture `
            -RebindProposal `
            -Edit {
                param($transition)
                $transition.governanceSchemas.hostNeutrality.sha256 =
                    '1' * 64
            }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSCurrentCandidateUniverse -RepoRoot $schemaFixture
            } `
            -MessageFragment 'governance schema'

        $inputFixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'bad-v4-input')
        Edit-RSTestRound4Transition `
            -Root $inputFixture `
            -RebindProposal `
            -Edit {
                param($transition)
                $transition.frozenInputs[0].sha256 = '1' * 64
            }
        Assert-RSTestTransitionThrows `
            -Action {
                Read-RSCurrentCandidateUniverse -RepoRoot $inputFixture
            } `
            -MessageFragment 'frozen input'
    }

    It 'rejects v4 host-surface drift through the current neutrality proof' -Skip:$script:RSTestRound4Superseded {
        $fixture = New-RSTestRound4TransitionRepository -Root (
            Join-Path $TestDrive 'v4-neutrality-surface-drift')
        $surfacePath = Join-Path $fixture (
            'Tools\TerminalEngine\TerminalEngineGate0ContractTestAdapter.cpp')
        [IO.File]::AppendAllText(
            $surfacePath,
            "`n",
            [Text.UTF8Encoding]::new($false))
        Assert-RSTestTransitionThrows `
            -Action { Read-RSCurrentCandidateUniverse -RepoRoot $fixture } `
            -MessageFragment (
                'Collection-neutrality surface identity differs from the frozen policy baseline')
    }
}
