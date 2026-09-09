Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot (
    'Tools\TerminalEngine\TerminalEngineC7Proof.psm1')
$evidenceModulePath = Join-Path $repoRoot 'Tools\TerminalEvidence.psm1'
Import-Module $evidenceModulePath -Force
Import-Module $modulePath -Force

function ConvertTo-RSC7TestJcsBytes {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject
    )

    return ,([RedSalamander.Tools.TerminalEvidenceJcs]::Serialize($InputObject))
}

function Test-RSC7ActionThrows {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    try {
        & $Action
        return $false
    }
    catch {
        return $true
    }
}

function New-RSC7TestInputs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $stateJcs = Get-RSTerminalEngineC7ExpectedStateJcs `
        -RepoRoot $RepositoryRoot
    return [pscustomobject]@{
        StateJcs = $stateJcs
        States = [ordered]@{
            'all-at-once' = $stateJcs
            'bytewise' = $stateJcs
            'terminator-split' = $stateJcs
        }
        OrderingJcs = (
            '{"admittedOrder":"protocol-reply,input,resize,side-channel",' +
            '"reordered":false}')
        OrderingFlags = [string[]]@(
            'action-observed',
            'checked-arithmetic',
            'order-preserved')
    }
}

function Invoke-RSC7TestProof {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Inputs
    )

    return New-RSTerminalEngineC7ProviderProof `
        -RepoRoot $RepositoryRoot `
        -CandidateId candidate-alpha `
        -PinSha256 ('1' * 64) `
        -SourceArchiveSha256 ('2' * 64) `
        -Platform x64 `
        -Configuration Release `
        -AdapterIdentitySha256 ('3' * 64) `
        -StateJcsBySchedule $Inputs.States `
        -OrderingStateJcs $Inputs.OrderingJcs `
        -OrderingActionFlags $Inputs.OrderingFlags
}

Describe 'Candidate-neutral TerminalEngine C7 score proof' {
    BeforeAll {
        $script:C7Inputs = New-RSC7TestInputs -RepositoryRoot $repoRoot
    }

    It 'freezes 31 concrete composition cases and the three exact boundary cases' {
        $context = Read-RSTerminalEngineC7ProofContract -RepoRoot $repoRoot
        $matrix = Invoke-RSTerminalEngineC7AuthenticationMatrix `
            -Contract $context.Contract
        $boundaries = Invoke-RSTerminalEngineC7BoundaryMatrix `
            -Contract $context.Contract

        $matrix.Count | Should Be 31
        @($matrix | Where-Object classification -CEQ 'positive').Count |
            Should Be 4
        @($matrix | Where-Object classification -CNE 'positive').Count |
            Should Be 27
        @($matrix | Where-Object caseId -CEQ 'valid-idle-composition')[0].
            idleAuthorized | Should Be $true
        @($matrix | Where-Object caseId -CEQ 'replay-invalidates')[0].
            epochStatus | Should Be 'invalidated'
        @($matrix | Where-Object caseId -CEQ 'out-of-order-invalidates')[0].
            decisions[-1] | Should Be 'epoch-invalidated'

        $boundaries.Count | Should Be 3
        $boundaries[0].inputValue | Should Be '2097152'
        $boundaries[0].decision | Should Be 'idle-authorized'
        $boundaries[1].inputValue | Should Be '2097153'
        $boundaries[1].decision | Should Be 'rejected-authentication'
        $boundaries[2].inputValue | Should Be '18446744073709551615'
        $boundaries[2].epochStatus | Should Be 'invalidated'
    }

    It 'reconstructs the exact bounded synthetic engine stream' {
        $bytes = Get-RSTerminalEngineC7FixedStreamBytes -RepoRoot $repoRoot
        $bytes.Length | Should Be 192
        $sha256 = [Convert]::ToHexString(
            [Security.Cryptography.SHA256]::HashData($bytes)).
            ToLowerInvariant()
        $sha256 | Should Be (
            'c52a8932860a92d8b51f96bd807a15b01ae49e870bdc7541cec0e1bd8e141688')
    }

    It 'creates one bounded content-free proof bound to engine, order, auth, and lane' {
        $result = Invoke-RSC7TestProof `
            -RepositoryRoot $repoRoot `
            -Inputs $script:C7Inputs
        $result.EvidenceObject.evidenceId | Should Be 'score.c7-proof'
        $result.EvidenceObject.resultCategory | Should Be 'verified'
        [UInt64]$result.EvidenceObject.sizeBytes | Should BeLessThan 32769
        $result.Proof.authentication.caseCount | Should Be '31'
        $result.Proof.authentication.boundaryCaseCount | Should Be '3'
        $result.Proof.engine.scheduleStates.Count | Should Be 3
        @($result.Proof.engine.scheduleStates.sha256 |
                Select-Object -Unique).Count | Should Be 1

        $proofJcs = [Text.Encoding]::UTF8.GetString(
            (ConvertTo-RSC7TestJcsBytes -InputObject $result.Proof))
        foreach ($forbidden in @(
                '00112233445566778899aabbccddeeff',
                'file://localhost/C:/rs-c7',
                'rs-c7-title',
                'https://example.invalid/rs-c7',
                'echo rs-c7',
                'terminalText',
                'payload')) {
            $proofJcs | Should Not Match ([regex]::Escape($forbidden))
        }
    }

    It 'rejects engine-state, schedule, order, and action-flag mismatches' {
        $state = $script:C7Inputs.StateJcs |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
        $state.hostParserCount = '1'
        $badState = [Text.Encoding]::UTF8.GetString(
            (ConvertTo-RSC7TestJcsBytes -InputObject $state))
        $badStates = [ordered]@{
            'all-at-once' = $badState
            'bytewise' = $script:C7Inputs.StateJcs
            'terminator-split' = $script:C7Inputs.StateJcs
        }
        (Test-RSC7ActionThrows -Action {
            New-RSTerminalEngineC7ProviderProof `
                -RepoRoot $repoRoot `
                -CandidateId candidate-alpha `
                -PinSha256 ('1' * 64) `
                -SourceArchiveSha256 ('2' * 64) `
                -Platform x64 `
                -Configuration Release `
                -AdapterIdentitySha256 ('3' * 64) `
                -StateJcsBySchedule $badStates `
                -OrderingStateJcs $script:C7Inputs.OrderingJcs `
                -OrderingActionFlags $script:C7Inputs.OrderingFlags
        }) | Should Be $true

        $missingSchedule = [ordered]@{
            'all-at-once' = $script:C7Inputs.StateJcs
            'bytewise' = $script:C7Inputs.StateJcs
        }
        (Test-RSC7ActionThrows -Action {
            New-RSTerminalEngineC7ProviderProof `
                -RepoRoot $repoRoot `
                -CandidateId candidate-alpha `
                -PinSha256 ('1' * 64) `
                -SourceArchiveSha256 ('2' * 64) `
                -Platform x64 `
                -Configuration Release `
                -AdapterIdentitySha256 ('3' * 64) `
                -StateJcsBySchedule $missingSchedule `
                -OrderingStateJcs $script:C7Inputs.OrderingJcs `
                -OrderingActionFlags $script:C7Inputs.OrderingFlags
        }) | Should Be $true

        (Test-RSC7ActionThrows -Action {
            New-RSTerminalEngineC7ProviderProof `
                -RepoRoot $repoRoot `
                -CandidateId candidate-alpha `
                -PinSha256 ('1' * 64) `
                -SourceArchiveSha256 ('2' * 64) `
                -Platform x64 `
                -Configuration Release `
                -AdapterIdentitySha256 ('3' * 64) `
                -StateJcsBySchedule $script:C7Inputs.States `
                -OrderingStateJcs (
                    '{"admittedOrder":"input,protocol-reply,resize,side-channel",' +
                    '"reordered":true}') `
                -OrderingActionFlags $script:C7Inputs.OrderingFlags
        }) | Should Be $true

        (Test-RSC7ActionThrows -Action {
            New-RSTerminalEngineC7ProviderProof `
                -RepoRoot $repoRoot `
                -CandidateId candidate-alpha `
                -PinSha256 ('1' * 64) `
                -SourceArchiveSha256 ('2' * 64) `
                -Platform x64 `
                -Configuration Release `
                -AdapterIdentitySha256 ('3' * 64) `
                -StateJcsBySchedule $script:C7Inputs.States `
                -OrderingStateJcs $script:C7Inputs.OrderingJcs `
                -OrderingActionFlags @(
                    'action-observed',
                    'checked-arithmetic')
        }) | Should Be $true
    }

    It 'rejects driver injection and adds only one provider-owned sorted descriptor' {
        $candidate = [ordered]@{
            candidateId = 'candidate-alpha'
            pinSha256 = ('1' * 64)
            outcome = 'complete'
            evidenceObjects = [object[]]@(
                [ordered]@{
                    evidenceId = 'score.c7'
                    sha256 = ('4' * 64)
                    sizeBytes = '1'
                    resultCategory = 'c7-r5'
                },
                [ordered]@{
                    evidenceId = 'score.c8'
                    sha256 = ('5' * 64)
                    sizeBytes = '1'
                    resultCategory = 'c8-r3'
                })
        }
        $null = Add-RSTerminalEngineC7ProviderEvidence `
            -CandidateResult $candidate `
            -RepoRoot $repoRoot `
            -SourceArchiveSha256 ('2' * 64) `
            -Platform x64 `
            -Configuration Release `
            -AdapterIdentitySha256 ('3' * 64) `
            -StateJcsBySchedule $script:C7Inputs.States `
            -OrderingStateJcs $script:C7Inputs.OrderingJcs `
            -OrderingActionFlags $script:C7Inputs.OrderingFlags
        ($candidate.evidenceObjects.evidenceId -join ',') |
            Should Be 'score.c7,score.c7-proof,score.c8'

        (Test-RSC7ActionThrows -Action {
            Add-RSTerminalEngineC7ProviderEvidence `
                -CandidateResult $candidate `
                -RepoRoot $repoRoot `
                -SourceArchiveSha256 ('2' * 64) `
                -Platform x64 `
                -Configuration Release `
                -AdapterIdentitySha256 ('3' * 64) `
                -StateJcsBySchedule $script:C7Inputs.States `
                -OrderingStateJcs $script:C7Inputs.OrderingJcs `
                -OrderingActionFlags $script:C7Inputs.OrderingFlags
        }) | Should Be $true
    }

    It 'rejects replayable or unbounded evaluator descriptors' {
        (Test-RSC7ActionThrows -Action {
            Assert-RSTerminalEngineC7EvidenceObject `
                -EvidenceObject ([ordered]@{
                    evidenceId = 'score.c7-proof'
                    sha256 = ('1' * 64)
                    sizeBytes = '32769'
                    resultCategory = 'verified'
                }) `
                -RepoRoot $repoRoot
        }) | Should Be $true
        (Test-RSC7ActionThrows -Action {
            Assert-RSTerminalEngineC7EvidenceObject `
                -EvidenceObject ([ordered]@{
                    evidenceId = 'score.c7-proof'
                    sha256 = ('1' * 64)
                    sizeBytes = '100'
                    resultCategory = 'claimed'
                }) `
                -RepoRoot $repoRoot
        }) | Should Be $true
    }
}

InModuleScope TerminalEngineC7Proof {
    function New-RSC7ConcreteTestContext {
        return New-RSC7StateMachineContext
    }

    function New-RSC7ConcreteTestObservation {
        return [ordered]@{
            typedEventKind = 'input-start'
            engineEventOrdinal = '6'
            livePromptSpanAvailable = $true
            primaryScreen = $true
            commandRunning = $false
            alternateScreen = $false
            nestedOwner = $false
        }
    }

    function New-RSC7ConcreteTestFrame {
        param(
            [string]$ProtocolVersion = '1',
            [string]$Nonce = '00112233445566778899aabbccddeeff',
            [string]$Epoch = '7',
            [string]$Sequence = '2',
            [string]$EventKind = 'activity',
            [string]$OperationId = 'idle-2',
            [string]$Activity = 'idle-at-primary-prompt',
            [string]$EngineEventOrdinal = '6'
        )

        return [ordered]@{
            protocolVersion = $ProtocolVersion
            nonce = $Nonce
            epoch = $Epoch
            sequence = $Sequence
            eventKind = $EventKind
            operationId = $OperationId
            payload = [ordered]@{
                activity = $Activity
                engineEventOrdinal = $EngineEventOrdinal
            }
        }
    }

    function Start-RSC7ConcreteTestEpoch {
        $context = New-RSC7ConcreteTestContext
        $handshake = New-RSC7ConcreteStep `
            -Context $context `
            -Token handshake-valid `
            -MaxFrameBytes ([UInt64]2097152)
        $decision = Invoke-RSC7FrameStateMachine `
            -Context $context `
            -Step $handshake `
            -MaxFrameBytes ([UInt64]2097152)
        if ($decision -cne 'handshake-accepted') {
            throw 'C7 concrete test handshake failed.'
        }
        return $context
    }

    function Invoke-RSC7ConcreteTestFrame {
        param(
            [Parameter(Mandatory = $true)]
            [Collections.IDictionary]$Context,

            [Parameter(Mandatory = $true)]
            [byte[]]$FrameBytes,

            [AllowNull()]
            [object]$Observation = $null,

            [string]$Origin = 'authenticated-side-channel'
        )

        if ($null -eq $Observation) {
            $Observation = New-RSC7ConcreteTestObservation
        }
        $step = [pscustomobject]@{
            HasFrame = $true
            Origin = $Origin
            FrameBytes = $FrameBytes
            Observation = $Observation
        }
        return Invoke-RSC7FrameStateMachine `
            -Context $Context `
            -Step $step `
            -MaxFrameBytes ([UInt64]2097152)
    }

    Describe 'Concrete C7 authenticated-frame parser and composition state machine' {
        It 'rejects each independently mutated authentication field' {
            $mutations = [object[]]@(
                @{ Name = 'protocolVersion'; Value = '2' },
                @{ Name = 'nonce'; Value = 'ffeeddccbbaa99887766554433221100' },
                @{ Name = 'epoch'; Value = '8' },
                @{ Name = 'sequence'; Value = '3' },
                @{ Name = 'eventKind'; Value = 'unknown' },
                @{ Name = 'operationId'; Value = 'bad operation' })
            foreach ($mutation in $mutations) {
                $context = Start-RSC7ConcreteTestEpoch
                $frame = New-RSC7ConcreteTestFrame
                $frame[[string]$mutation.Name] = [string]$mutation.Value
                $decision = Invoke-RSC7ConcreteTestFrame `
                    -Context $context `
                    -FrameBytes (ConvertTo-RSJcsUtf8Bytes -InputObject $frame)
                $decision | Should Be 'rejected-authentication'
                [bool]$context.invalidated | Should Be $true
            }

            foreach ($payloadMutation in @(
                    @{ Name = 'activity'; Value = 'running' },
                    @{ Name = 'engineEventOrdinal'; Value = '0' })) {
                $context = Start-RSC7ConcreteTestEpoch
                $frame = New-RSC7ConcreteTestFrame
                $frame.payload[[string]$payloadMutation.Name] =
                    [string]$payloadMutation.Value
                $decision = Invoke-RSC7ConcreteTestFrame `
                    -Context $context `
                    -FrameBytes (ConvertTo-RSJcsUtf8Bytes -InputObject $frame)
                $decision | Should Be 'rejected-authentication'
                [bool]$context.invalidated | Should Be $true
            }
        }

        It 'rejects every missing field, duplicate key, unknown property, and noncanonical encoding' {
            foreach ($name in @(
                    'protocolVersion',
                    'nonce',
                    'epoch',
                    'sequence',
                    'eventKind',
                    'operationId',
                    'payload')) {
                $context = Start-RSC7ConcreteTestEpoch
                $frame = New-RSC7ConcreteTestFrame
                [void]$frame.Remove($name)
                $decision = Invoke-RSC7ConcreteTestFrame `
                    -Context $context `
                    -FrameBytes (ConvertTo-RSJcsUtf8Bytes -InputObject $frame)
                $decision | Should Be 'rejected-authentication'
            }

            $context = Start-RSC7ConcreteTestEpoch
            $frame = New-RSC7ConcreteTestFrame
            $frame['unknown'] = 'x'
            (Invoke-RSC7ConcreteTestFrame `
                -Context $context `
                -FrameBytes (ConvertTo-RSJcsUtf8Bytes -InputObject $frame)) |
                Should Be 'rejected-authentication'

            $context = Start-RSC7ConcreteTestEpoch
            $frame = New-RSC7ConcreteTestFrame
            $canonical = [Text.Encoding]::UTF8.GetString(
                (ConvertTo-RSJcsUtf8Bytes -InputObject $frame))
            $duplicate = '{"epoch":"7",' + $canonical.Substring(1)
            (Invoke-RSC7ConcreteTestFrame `
                -Context $context `
                -FrameBytes ([Text.UTF8Encoding]::new(
                    $false,
                    $true).GetBytes($duplicate))) |
                Should Be 'rejected-authentication'

            $context = Start-RSC7ConcreteTestEpoch
            $noncanonical = New-RSC7ConcreteTestFrame |
                ConvertTo-Json -Compress -Depth 10
            (Invoke-RSC7ConcreteTestFrame `
                -Context $context `
                -FrameBytes ([Text.UTF8Encoding]::new(
                    $false,
                    $true).GetBytes($noncanonical))) |
                Should Be 'rejected-authentication'

            $context = Start-RSC7ConcreteTestEpoch
            (Invoke-RSC7ConcreteTestFrame `
                -Context $context `
                -FrameBytes ([byte[]]@(0xff, 0xfe, 0xfd))) |
                Should Be 'rejected-authentication'
        }

        It 'denies each independently mismatched engine observation without invalidating auth' {
            $mutations = [object[]]@(
                @{ Name = 'typedEventKind'; Value = 'none' },
                @{ Name = 'engineEventOrdinal'; Value = '5' },
                @{ Name = 'livePromptSpanAvailable'; Value = $false },
                @{ Name = 'primaryScreen'; Value = $false },
                @{ Name = 'commandRunning'; Value = $true },
                @{ Name = 'alternateScreen'; Value = $true },
                @{ Name = 'nestedOwner'; Value = $true })
            foreach ($mutation in $mutations) {
                $context = Start-RSC7ConcreteTestEpoch
                $observation = New-RSC7ConcreteTestObservation
                $observation[[string]$mutation.Name] = $mutation.Value
                $decision = Invoke-RSC7ConcreteTestFrame `
                    -Context $context `
                    -FrameBytes (ConvertTo-RSJcsUtf8Bytes `
                        -InputObject (New-RSC7ConcreteTestFrame)) `
                    -Observation $observation
                $decision | Should Be 'denied-engine-state'
                [bool]$context.invalidated | Should Be $false
                [UInt64]$context.nextSequence | Should Be ([UInt64]3)
            }
        }

        It 'accepts the exact 2 MiB frame, rejects plus one, and fails uint64 wrap' {
            $max = [UInt64]2097152

            $context = Start-RSC7ConcreteTestEpoch
            $exact = New-RSC7SizedActivityStep `
                -Context $context `
                -TargetSizeBytes $max
            ([byte[]]$exact.FrameBytes).Length | Should Be 2097152
            (Invoke-RSC7FrameStateMachine `
                -Context $context `
                -Step $exact `
                -MaxFrameBytes $max) | Should Be 'idle-authorized'

            $context = Start-RSC7ConcreteTestEpoch
            $plusOne = New-RSC7SizedActivityStep `
                -Context $context `
                -TargetSizeBytes ($max + [UInt64]1)
            ([byte[]]$plusOne.FrameBytes).Length | Should Be 2097153
            (Invoke-RSC7FrameStateMachine `
                -Context $context `
                -Step $plusOne `
                -MaxFrameBytes $max) | Should Be 'rejected-authentication'
            [bool]$context.invalidated | Should Be $true

            $context = New-RSC7ConcreteTestContext
            $context.active = $true
            $context.nextSequence = [UInt64]::MaxValue
            [void]$context.seenOperationIds.Add('handshake-1')
            $maxStep = New-RSC7ConcreteStep `
                -Context $context `
                -Token idle-valid `
                -MaxFrameBytes $max
            (Invoke-RSC7FrameStateMachine `
                -Context $context `
                -Step $maxStep `
                -MaxFrameBytes $max) | Should Be 'rejected-authentication'
            [bool]$context.invalidated | Should Be $true
        }
    }
}
