Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\TerminalEvidence.psm1') -Force

$script:ContractLeaf = 'Specs\Terminal\TerminalEngineC7ProofContract.v1.json'
$script:SchemaLeaf = 'Specs\Terminal\TerminalEngineC7Proof.schema.json'
$script:EvidenceId = 'score.c7-proof'
$script:ResultCategory = 'verified'
$script:ProofNonce = '00112233445566778899aabbccddeeff'
$script:ProofEpoch = '7'
$script:ExpectedCaseIds = [string[]]@(
    'valid-idle-composition',
    'typed-event-without-auth',
    'auth-without-typed-event',
    'ordinary-osc133-cannot-authorize',
    'ordinary-osc7-cannot-authorize',
    'remote-output-cannot-authorize',
    'nested-output-does-not-invalidate',
    'nested-owner-pauses-then-resumes',
    'running-command-denied',
    'alternate-screen-denied',
    'non-primary-screen-denied',
    'missing-live-prompt-denied',
    'wrong-nonce-invalidates',
    'missing-nonce-invalidates',
    'wrong-epoch-invalidates',
    'malformed-frame-invalidates',
    'oversized-frame-invalidates',
    'protocol-version-invalidates',
    'replay-invalidates',
    'duplicate-operation-invalidates',
    'sequence-gap-invalidates',
    'out-of-order-invalidates',
    'later-marker-cannot-recover-invalid-epoch',
    'duplicate-handshake-invalidates',
    'valid-engine-mismatch-consumes-then-recovers',
    'post-handshake-wrong-nonce-invalidates',
    'post-handshake-missing-nonce-invalidates',
    'post-handshake-wrong-epoch-invalidates',
    'post-handshake-malformed-invalidates',
    'post-handshake-oversized-invalidates',
    'post-handshake-protocol-version-invalidates'
)

function Get-RSC7Property {
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
                [string]::Equals(
                    [string]$entry.Key,
                    $Name,
                    [StringComparison]::Ordinal)) {
                $found = $true
                $value = $entry.Value
                break
            }
        }
    }
    elseif ($null -ne $InputObject) {
        foreach ($property in $InputObject.PSObject.Properties) {
            if ([string]::Equals(
                    $property.Name,
                    $Name,
                    [StringComparison]::Ordinal)) {
                $found = $true
                $value = $property.Value
                break
            }
        }
    }

    if ($Required -and -not $found) {
        throw [IO.InvalidDataException]::new(
            "C7 proof object is missing required property '$Name'.")
    }
    return [pscustomobject]@{
        Found = $found
        Value = $value
    }
}

function Get-RSC7RequiredString {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $property = Get-RSC7Property -InputObject $InputObject -Name $Name -Required
    if ($property.Value -isnot [string] -or
        [string]::IsNullOrEmpty([string]$property.Value)) {
        throw [IO.InvalidDataException]::new(
            "C7 proof property '$Name' must be a nonempty string.")
    }
    return [string]$property.Value
}

function Get-RSC7RequiredArray {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $property = Get-RSC7Property -InputObject $InputObject -Name $Name -Required
    if ($property.Value -is [string] -or
        $property.Value -isnot [Collections.IEnumerable]) {
        throw [IO.InvalidDataException]::new(
            "C7 proof property '$Name' must be an array.")
    }
    return [object[]]@($property.Value)
}

function Assert-RSC7ExactProperties {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string[]]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $actual = [string[]]@(
        if ($InputObject -is [Collections.IDictionary]) {
            foreach ($entry in $InputObject.GetEnumerator()) {
                if ($entry.Key -isnot [string]) {
                    throw [IO.InvalidDataException]::new(
                        "$Context contains a non-string property name.")
                }
                [string]$entry.Key
            }
        }
        elseif ($null -ne $InputObject) {
            $InputObject.PSObject.Properties.Name
        }
    )
    [Array]::Sort($actual, [StringComparer]::Ordinal)
    $expectedSorted = [string[]]$Expected.Clone()
    [Array]::Sort($expectedSorted, [StringComparer]::Ordinal)
    if (($actual -join "`0") -cne ($expectedSorted -join "`0")) {
        throw [IO.InvalidDataException]::new(
            "$Context properties differ from the frozen closed contract.")
    }
}

function Assert-RSC7ExactStringSet {
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

    $actualSorted = [string[]]$Actual.Clone()
    $expectedSorted = [string[]]$Expected.Clone()
    [Array]::Sort($actualSorted, [StringComparer]::Ordinal)
    [Array]::Sort($expectedSorted, [StringComparer]::Ordinal)
    if (($actualSorted -join "`0") -cne ($expectedSorted -join "`0")) {
        throw [IO.InvalidDataException]::new(
            "$Context differs from the frozen exact set.")
    }
    for ($index = 1; $index -lt $actualSorted.Count; ++$index) {
        if ($actualSorted[$index - 1] -ceq $actualSorted[$index]) {
            throw [IO.InvalidDataException]::new(
                "$Context contains duplicate '$($actualSorted[$index])'.")
        }
    }
}

function Assert-RSC7Sha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    if ($Value -cnotmatch '^[0-9a-f]{64}$') {
        throw [IO.InvalidDataException]::new(
            "$Context must be 64 lowercase hexadecimal SHA-256 characters.")
    }
}

function Get-RSC7UInt64 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    [UInt64]$parsed = 0
    if ($Value -cnotmatch '^(?:0|[1-9][0-9]{0,19})$' -or
        -not [UInt64]::TryParse(
            $Value,
            [Globalization.NumberStyles]::None,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$parsed)) {
        throw [IO.InvalidDataException]::new(
            "$Context must be a canonical unsigned 64-bit decimal string.")
    }
    return $parsed
}

function Import-RSC7EvidenceModule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    Import-Module (Join-Path $RepoRoot 'Tools\TerminalEvidence.psm1') -Force
}

function Get-RSTerminalEngineC7FixedStreamBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $contractContext = Read-RSTerminalEngineC7ProofContract -RepoRoot $RepoRoot
    $escape = [char]0x1b
    $stream = [string]::Concat([string[]]@(
            "${escape}]0;rs-c7-title${escape}\",
            "${escape}]7;file://localhost/C:/rs-c7${escape}\",
            "${escape}]133;A${escape}\",
            '$ ',
            "${escape}]133;B${escape}\",
            "echo rs-c7`r`n",
            "${escape}]133;C;cmdline_url=echo%20rs-c7${escape}\",
            "${escape}]8;;https://example.invalid/rs-c7${escape}\",
            'x',
            "${escape}]8;;${escape}\",
            "`r`n",
            "${escape}]133;D;23${escape}\",
            "${escape}]133;A${escape}\",
            '$ ',
            "${escape}]133;B${escape}\"
        ))
    $bytes = [Text.UTF8Encoding]::new($false, $true).GetBytes($stream)
    $engineProbe = (Get-RSC7Property `
            -InputObject $contractContext.Contract `
            -Name engineProbe `
            -Required).Value
    $expectedSize = Get-RSC7UInt64 `
        -Value (Get-RSC7RequiredString `
            -InputObject $engineProbe `
            -Name fixedStreamSizeBytes) `
        -Context 'C7 fixed stream size'
    if ([UInt64]$bytes.Length -ne $expectedSize) {
        throw [IO.InvalidDataException]::new(
            "C7 fixed stream size '$($bytes.Length)' differs from '$expectedSize'.")
    }
    $actualSha256 = [Convert]::ToHexString(
        [Security.Cryptography.SHA256]::HashData($bytes)).ToLowerInvariant()
    $expectedSha256 = Get-RSC7RequiredString `
        -InputObject $engineProbe `
        -Name fixedStreamSha256
    if ($actualSha256 -cne $expectedSha256) {
        throw [IO.InvalidDataException]::new(
            'C7 fixed stream bytes differ from the frozen identity.')
    }
    return ,$bytes
}

function New-RSC7ConcreteStep {
    param(
        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$Context,

        [Parameter(Mandatory = $true)]
        [string]$Token,

        [Parameter(Mandatory = $true)]
        [UInt64]$MaxFrameBytes
    )

    $validObservation = [ordered]@{
        typedEventKind = 'input-start'
        engineEventOrdinal = '6'
        livePromptSpanAvailable = $true
        primaryScreen = $true
        commandRunning = $false
        alternateScreen = $false
        nestedOwner = $false
    }
    if ($Token -ceq 'idle-typed-no-auth') {
        return [pscustomobject]@{
            HasFrame = $false
            Origin = 'none'
            FrameBytes = [byte[]]@()
            Observation = $validObservation
        }
    }

    $escape = [char]0x1b
    if ($Token -ceq 'idle-ordinary-osc133') {
        return [pscustomobject]@{
            HasFrame = $true
            Origin = 'terminal-output'
            FrameBytes = [Text.UTF8Encoding]::new($false, $true).GetBytes(
                "${escape}]133;A${escape}\")
            Observation = $validObservation
        }
    }
    if ($Token -ceq 'idle-ordinary-osc7') {
        return [pscustomobject]@{
            HasFrame = $true
            Origin = 'terminal-output'
            FrameBytes = [Text.UTF8Encoding]::new($false, $true).GetBytes(
                "${escape}]7;file://remote.invalid/tmp${escape}\")
            Observation = $validObservation
        }
    }

    $isHandshake = $Token.StartsWith(
        'handshake-',
        [StringComparison]::Ordinal)
    $sequence = if ($isHandshake) {
        [UInt64]1
    }
    else {
        [UInt64]$Context['nextSequence']
    }
    if ($Token -ceq 'idle-replay') {
        $sequence = if ([UInt64]$Context['nextSequence'] -eq 0) {
            [UInt64]0
        }
        else {
            [UInt64]$Context['nextSequence'] - [UInt64]1
        }
    }
    elseif ($Token -in @('idle-gap', 'idle-out-of-order')) {
        $sequence = [UInt64]$Context['nextSequence'] + [UInt64]1
    }

    $operationId = if ($isHandshake) {
        'handshake-1'
    }
    else {
        "idle-$sequence"
    }
    if ($Token -ceq 'idle-duplicate-operation') {
        $operationId = [string]$Context['lastOperationId']
    }

    $payload = if ($isHandshake) {
        [ordered]@{
            capability = 'prompt-activity-v1'
        }
    }
    else {
        [ordered]@{
            activity = 'idle-at-primary-prompt'
            engineEventOrdinal = '6'
        }
    }
    $frame = [ordered]@{
        protocolVersion = '1'
        nonce = $script:ProofNonce
        epoch = $script:ProofEpoch
        sequence = $sequence.ToString(
            [Globalization.CultureInfo]::InvariantCulture)
        eventKind = if ($isHandshake) { 'handshake' } else { 'activity' }
        operationId = $operationId
        payload = $payload
    }

    switch ($Token) {
        { $_ -in @('handshake-wrong-nonce', 'idle-wrong-nonce') } {
            $frame['nonce'] = 'ffeeddccbbaa99887766554433221100'
            break
        }
        { $_ -in @('handshake-missing-nonce', 'idle-missing-nonce') } {
            [void]$frame.Remove('nonce')
            break
        }
        { $_ -in @('handshake-wrong-epoch', 'idle-wrong-epoch') } {
            $frame['epoch'] = '8'
            break
        }
        { $_ -in @('handshake-protocol-version', 'idle-protocol-version') } {
            $frame['protocolVersion'] = '2'
            break
        }
        'idle-auth-no-typed' {
            $validObservation['typedEventKind'] = 'none'
            break
        }
        'idle-running' {
            $validObservation['commandRunning'] = $true
            break
        }
        'idle-alt-screen' {
            $validObservation['alternateScreen'] = $true
            break
        }
        'idle-not-primary' {
            $validObservation['primaryScreen'] = $false
            break
        }
        'idle-no-live-span' {
            $validObservation['livePromptSpanAvailable'] = $false
            break
        }
        'idle-nested-auth' {
            $validObservation['nestedOwner'] = $true
            break
        }
    }

    if ($Token -in @('handshake-malformed', 'idle-malformed')) {
        $frameBytes = [Text.UTF8Encoding]::new($false, $true).GetBytes(
            '{"protocolVersion":')
    }
    elseif ($Token -in @('handshake-oversized', 'idle-oversized')) {
        $payload['padding'] = 'x' * ([int]$MaxFrameBytes)
        $frameBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $frame
    }
    else {
        $frameBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $frame
    }

    $origin = if ($Token -ceq 'idle-remote-output') {
        'remote-output'
    }
    elseif ($Token -ceq 'idle-nested-output') {
        'nested-output'
    }
    else {
        'authenticated-side-channel'
    }
    return [pscustomobject]@{
        HasFrame = $true
        Origin = $origin
        FrameBytes = [byte[]]$frameBytes
        Observation = $validObservation
    }
}

function Set-RSC7EpochInvalid {
    param(
        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$Context
    )

    $Context['invalidated'] = $true
    $Context['idleAuthorized'] = $false
}

function Test-RSC7FrameShape {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Frame
    )

    try {
        Assert-RSC7ExactProperties `
            -InputObject $Frame `
            -Expected @(
                'protocolVersion',
                'nonce',
                'epoch',
                'sequence',
                'eventKind',
                'operationId',
                'payload') `
            -Context 'C7 authenticated frame'
    }
    catch {
        return $false
    }
    foreach ($name in @(
            'protocolVersion',
            'nonce',
            'epoch',
            'sequence',
            'eventKind',
            'operationId')) {
        $value = (Get-RSC7Property `
                -InputObject $Frame `
                -Name $name `
                -Required).Value
        if ($value -isnot [string] -or
            [string]::IsNullOrEmpty([string]$value)) {
            return $false
        }
    }
    return $true
}

function Test-RSC7EngineObservation {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Observation,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedOrdinal
    )

    try {
        Assert-RSC7ExactProperties `
            -InputObject $Observation `
            -Expected @(
                'typedEventKind',
                'engineEventOrdinal',
                'livePromptSpanAvailable',
                'primaryScreen',
                'commandRunning',
                'alternateScreen',
                'nestedOwner') `
            -Context 'C7 engine observation'
    }
    catch {
        return $false
    }
    foreach ($name in @(
            'livePromptSpanAvailable',
            'primaryScreen',
            'commandRunning',
            'alternateScreen',
            'nestedOwner')) {
        if ((Get-RSC7Property `
                -InputObject $Observation `
                -Name $name `
                -Required).Value -isnot [bool]) {
            return $false
        }
    }
    return (
        (Get-RSC7RequiredString `
            -InputObject $Observation `
            -Name typedEventKind) -ceq 'input-start' -and
        (Get-RSC7RequiredString `
            -InputObject $Observation `
            -Name engineEventOrdinal) -ceq $ExpectedOrdinal -and
        [bool](Get-RSC7Property `
            -InputObject $Observation `
            -Name livePromptSpanAvailable `
            -Required).Value -and
        [bool](Get-RSC7Property `
            -InputObject $Observation `
            -Name primaryScreen `
            -Required).Value -and
        -not [bool](Get-RSC7Property `
            -InputObject $Observation `
            -Name commandRunning `
            -Required).Value -and
        -not [bool](Get-RSC7Property `
            -InputObject $Observation `
            -Name alternateScreen `
            -Required).Value -and
        -not [bool](Get-RSC7Property `
            -InputObject $Observation `
            -Name nestedOwner `
            -Required).Value)
}

function Invoke-RSC7FrameStateMachine {
    param(
        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$Context,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Step,

        [Parameter(Mandatory = $true)]
        [UInt64]$MaxFrameBytes
    )

    $Context['idleAuthorized'] = $false
    if (-not [bool]$Step.HasFrame) {
        return 'rejected-no-authenticated-frame'
    }
    if ([string]$Step.Origin -cne 'authenticated-side-channel') {
        return 'ignored-untrusted-origin'
    }
    if ([bool]$Context['invalidated']) {
        return 'epoch-invalidated'
    }
    $frameBytes = [byte[]]$Step.FrameBytes
    if ([UInt64]$frameBytes.Length -gt $MaxFrameBytes) {
        Set-RSC7EpochInvalid -Context $Context
        return 'rejected-authentication'
    }

    $frame = $null
    try {
        $frameText = [Text.UTF8Encoding]::new($false, $true).GetString(
            $frameBytes)
        $frame = $frameText |
            ConvertFrom-Json `
                -AsHashtable `
                -Depth 20 `
                -NoEnumerate `
                -DateKind String
        if (-not (Test-RSC7FrameShape -Frame $frame)) {
            throw [IO.InvalidDataException]::new(
                'Authenticated frame shape is invalid.')
        }
        $canonicalFrame = [Text.Encoding]::UTF8.GetString(
            (ConvertTo-RSJcsUtf8Bytes -InputObject $frame))
        if ($canonicalFrame -cne $frameText) {
            throw [IO.InvalidDataException]::new(
                'Authenticated frame is not canonical JSON.')
        }
    }
    catch {
        Set-RSC7EpochInvalid -Context $Context
        return 'rejected-authentication'
    }

    $protocolVersion = Get-RSC7RequiredString `
        -InputObject $frame `
        -Name protocolVersion
    $nonce = Get-RSC7RequiredString -InputObject $frame -Name nonce
    $epoch = Get-RSC7RequiredString -InputObject $frame -Name epoch
    $sequenceText = Get-RSC7RequiredString `
        -InputObject $frame `
        -Name sequence
    $eventKind = Get-RSC7RequiredString `
        -InputObject $frame `
        -Name eventKind
    $operationId = Get-RSC7RequiredString `
        -InputObject $frame `
        -Name operationId
    [UInt64]$sequence = 0
    $identityValid = (
        $protocolVersion -ceq '1' -and
        $nonce -cmatch '^[0-9a-f]{32}$' -and
        $nonce -ceq $script:ProofNonce -and
        $epoch -ceq $script:ProofEpoch -and
        $sequenceText -cmatch '^(?:0|[1-9][0-9]{0,19})$' -and
        [UInt64]::TryParse(
            $sequenceText,
            [Globalization.NumberStyles]::None,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$sequence) -and
        $operationId -cmatch '^[a-z0-9][a-z0-9._-]{0,63}$')
    if (-not $identityValid) {
        Set-RSC7EpochInvalid -Context $Context
        return 'rejected-authentication'
    }

    $payload = (Get-RSC7Property `
            -InputObject $frame `
            -Name payload `
            -Required).Value
    if ($eventKind -ceq 'handshake') {
        $payloadValid = $true
        try {
            Assert-RSC7ExactProperties `
                -InputObject $payload `
                -Expected @('capability') `
                -Context 'C7 handshake payload'
            $payloadValid = (
                (Get-RSC7RequiredString `
                    -InputObject $payload `
                    -Name capability) -ceq 'prompt-activity-v1')
        }
        catch {
            $payloadValid = $false
        }
        if ([bool]$Context['active'] -or
            $sequence -ne [UInt64]1 -or
            $operationId -cne 'handshake-1' -or
            -not $payloadValid) {
            Set-RSC7EpochInvalid -Context $Context
            return 'rejected-authentication'
        }
        [void]$Context['seenOperationIds'].Add($operationId)
        $Context['active'] = $true
        $Context['nextSequence'] = [UInt64]2
        return 'handshake-accepted'
    }

    if ($eventKind -cne 'activity' -or
        -not [bool]$Context['active'] -or
        $sequence -ne [UInt64]$Context['nextSequence'] -or
        $Context['seenOperationIds'].Contains($operationId)) {
        Set-RSC7EpochInvalid -Context $Context
        return 'rejected-authentication'
    }

    $payloadValid = $true
    try {
        $payloadProperties = [string[]]@(
            if ($payload -is [Collections.IDictionary]) {
                foreach ($entry in $payload.GetEnumerator()) {
                    [string]$entry.Key
                }
            }
            else {
                $payload.PSObject.Properties.Name
            })
        $expectedPayloadProperties = if ($payloadProperties -ccontains 'padding') {
            [string[]]@('activity', 'engineEventOrdinal', 'padding')
        }
        else {
            [string[]]@('activity', 'engineEventOrdinal')
        }
        Assert-RSC7ExactProperties `
            -InputObject $payload `
            -Expected $expectedPayloadProperties `
            -Context 'C7 activity payload'
        $payloadValid = (
            (Get-RSC7RequiredString `
                -InputObject $payload `
                -Name activity) -ceq 'idle-at-primary-prompt' -and
            (Get-RSC7RequiredString `
                -InputObject $payload `
                -Name engineEventOrdinal) -cmatch '^[1-9][0-9]{0,19}$' -and
            (-not ($payloadProperties -ccontains 'padding') -or
                (Get-RSC7Property `
                    -InputObject $payload `
                    -Name padding `
                    -Required).Value -is [string]))
    }
    catch {
        $payloadValid = $false
    }
    if (-not $payloadValid -or $sequence -eq [UInt64]::MaxValue) {
        Set-RSC7EpochInvalid -Context $Context
        return 'rejected-authentication'
    }

    [void]$Context['seenOperationIds'].Add($operationId)
    $Context['lastOperationId'] = $operationId
    $Context['nextSequence'] = $sequence + [UInt64]1
    $engineEventOrdinal = Get-RSC7RequiredString `
        -InputObject $payload `
        -Name engineEventOrdinal
    if (-not (Test-RSC7EngineObservation `
            -Observation $Step.Observation `
            -ExpectedOrdinal $engineEventOrdinal)) {
        return 'denied-engine-state'
    }
    $Context['idleAuthorized'] = $true
    return 'idle-authorized'
}

function Invoke-RSTerminalEngineC7AuthenticationMatrix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Contract
    )

    $authentication = (Get-RSC7Property `
            -InputObject $Contract `
            -Name authentication `
            -Required).Value
    $maxFrameBytes = Get-RSC7UInt64 `
        -Value (Get-RSC7RequiredString `
            -InputObject $authentication `
            -Name maxFrameBytes) `
        -Context 'C7 authenticated frame byte limit'
    $cases = Get-RSC7RequiredArray -InputObject $authentication -Name cases
    $results = [Collections.Generic.List[object]]::new()
    foreach ($case in $cases) {
        $context = [ordered]@{
            active = $false
            invalidated = $false
            nextSequence = [UInt64]1
            idleAuthorized = $false
            seenOperationIds = [Collections.Generic.HashSet[string]]::new(
                [StringComparer]::Ordinal)
            lastOperationId = ''
        }
        $decisions = [Collections.Generic.List[string]]::new()
        foreach ($tokenValue in (Get-RSC7RequiredArray `
                -InputObject $case `
                -Name scenario)) {
            if ($tokenValue -isnot [string] -or
                [string]::IsNullOrEmpty([string]$tokenValue)) {
                throw [IO.InvalidDataException]::new(
                    'C7 authentication scenario tokens must be nonempty strings.')
            }
            $step = New-RSC7ConcreteStep `
                -Context $context `
                -Token ([string]$tokenValue) `
                -MaxFrameBytes $maxFrameBytes
            $decisions.Add((Invoke-RSC7FrameStateMachine `
                    -Context $context `
                    -Step $step `
                    -MaxFrameBytes $maxFrameBytes))
        }
        $epochStatus = if ([bool]$context['invalidated']) {
            'invalidated'
        }
        elseif ([bool]$context['active']) {
            'active'
        }
        else {
            'awaiting-handshake'
        }
        $results.Add([ordered]@{
                caseId = Get-RSC7RequiredString -InputObject $case -Name caseId
                classification = Get-RSC7RequiredString `
                    -InputObject $case `
                    -Name classification
                decisions = [string[]]$decisions.ToArray()
                epochStatus = $epochStatus
                idleAuthorized = [bool]$context['idleAuthorized']
            })
    }
    return [object[]]$results.ToArray()
}

function New-RSC7StateMachineContext {
    param()

    return [ordered]@{
        active = $false
        invalidated = $false
        nextSequence = [UInt64]1
        idleAuthorized = $false
        seenOperationIds = [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
        lastOperationId = ''
    }
}

function New-RSC7SizedActivityStep {
    param(
        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$Context,

        [Parameter(Mandatory = $true)]
        [UInt64]$TargetSizeBytes
    )

    $sequence = [UInt64]$Context['nextSequence']
    $payload = [ordered]@{
        activity = 'idle-at-primary-prompt'
        engineEventOrdinal = '6'
        padding = ''
    }
    $frame = [ordered]@{
        protocolVersion = '1'
        nonce = $script:ProofNonce
        epoch = $script:ProofEpoch
        sequence = $sequence.ToString(
            [Globalization.CultureInfo]::InvariantCulture)
        eventKind = 'activity'
        operationId = "idle-$sequence"
        payload = $payload
    }
    $emptyBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $frame
    if ($TargetSizeBytes -lt [UInt64]$emptyBytes.Length -or
        $TargetSizeBytes -gt [UInt64][Int32]::MaxValue) {
        throw [IO.InvalidDataException]::new(
            'C7 sized frame target cannot be represented by the bounded synthetic payload.')
    }
    $paddingLength = [int]($TargetSizeBytes - [UInt64]$emptyBytes.Length)
    $payload['padding'] = 'x' * $paddingLength
    $frameBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $frame
    if ([UInt64]$frameBytes.Length -ne $TargetSizeBytes) {
        throw [IO.InvalidDataException]::new(
            'C7 sized frame generator did not reach its exact byte target.')
    }
    return [pscustomobject]@{
        HasFrame = $true
        Origin = 'authenticated-side-channel'
        FrameBytes = [byte[]]$frameBytes
        Observation = [ordered]@{
            typedEventKind = 'input-start'
            engineEventOrdinal = '6'
            livePromptSpanAvailable = $true
            primaryScreen = $true
            commandRunning = $false
            alternateScreen = $false
            nestedOwner = $false
        }
    }
}

function Invoke-RSTerminalEngineC7BoundaryMatrix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Contract
    )

    $authentication = (Get-RSC7Property `
            -InputObject $Contract `
            -Name authentication `
            -Required).Value
    $maxFrameBytes = Get-RSC7UInt64 `
        -Value (Get-RSC7RequiredString `
            -InputObject $authentication `
            -Name maxFrameBytes) `
        -Context 'C7 authenticated frame byte limit'
    $cases = Get-RSC7RequiredArray `
        -InputObject $authentication `
        -Name boundaryCases
    $results = [Collections.Generic.List[object]]::new()
    foreach ($case in $cases) {
        $context = New-RSC7StateMachineContext
        $precondition = Get-RSC7RequiredString `
            -InputObject $case `
            -Name precondition
        if ($precondition -ceq 'active-after-handshake') {
            $handshake = New-RSC7ConcreteStep `
                -Context $context `
                -Token handshake-valid `
                -MaxFrameBytes $maxFrameBytes
            $handshakeDecision = Invoke-RSC7FrameStateMachine `
                -Context $context `
                -Step $handshake `
                -MaxFrameBytes $maxFrameBytes
            if ($handshakeDecision -cne 'handshake-accepted') {
                throw [IO.InvalidDataException]::new(
                    'C7 boundary precondition handshake failed.')
            }
        }
        elseif ($precondition -ceq
            'active-with-next-sequence-at-uint64-max') {
            $context['active'] = $true
            $context['nextSequence'] = [UInt64]::MaxValue
            [void]$context['seenOperationIds'].Add('handshake-1')
        }
        else {
            throw [IO.InvalidDataException]::new(
                "Unknown C7 boundary precondition '$precondition'.")
        }

        $inputValue = Get-RSC7UInt64 `
            -Value (Get-RSC7RequiredString `
                -InputObject $case `
                -Name inputValue) `
            -Context 'C7 boundary input value'
        $caseId = Get-RSC7RequiredString -InputObject $case -Name caseId
        $step = if ($caseId -in @(
                'frame-size-exact-boundary',
                'frame-size-plus-one')) {
            New-RSC7SizedActivityStep `
                -Context $context `
                -TargetSizeBytes $inputValue
        }
        elseif ($caseId -ceq 'sequence-max-wrap-fatal') {
            New-RSC7ConcreteStep `
                -Context $context `
                -Token idle-valid `
                -MaxFrameBytes $maxFrameBytes
        }
        else {
            throw [IO.InvalidDataException]::new(
                "Unknown C7 boundary case '$caseId'.")
        }
        $decision = Invoke-RSC7FrameStateMachine `
            -Context $context `
            -Step $step `
            -MaxFrameBytes $maxFrameBytes
        $epochStatus = if ([bool]$context['invalidated']) {
            'invalidated'
        }
        elseif ([bool]$context['active']) {
            'active'
        }
        else {
            'awaiting-handshake'
        }
        $results.Add([ordered]@{
                caseId = $caseId
                inputValue = $inputValue.ToString(
                    [Globalization.CultureInfo]::InvariantCulture)
                decision = $decision
                epochStatus = $epochStatus
                idleAuthorized = [bool]$context['idleAuthorized']
            })
    }
    return [object[]]$results.ToArray()
}

function Assert-RSC7BoundaryMatrix {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Contract
    )

    $authentication = (Get-RSC7Property `
            -InputObject $Contract `
            -Name authentication `
            -Required).Value
    $cases = Get-RSC7RequiredArray `
        -InputObject $authentication `
        -Name boundaryCases
    $expectedIds = [string[]]@(
        'frame-size-exact-boundary',
        'frame-size-plus-one',
        'sequence-max-wrap-fatal')
    $actualIds = [string[]]@(
        $cases | ForEach-Object {
            Get-RSC7RequiredString -InputObject $_ -Name caseId
        })
    if (($actualIds -join "`0") -cne ($expectedIds -join "`0")) {
        throw [IO.InvalidDataException]::new(
            'C7 boundary cases differ from the frozen ordered set.')
    }

    $results = Invoke-RSTerminalEngineC7BoundaryMatrix -Contract $Contract
    for ($index = 0; $index -lt $cases.Count; ++$index) {
        if ([string]$results[$index].decision -cne
                (Get-RSC7RequiredString `
                    -InputObject $cases[$index] `
                    -Name expectedDecision) -or
            [string]$results[$index].epochStatus -cne
                (Get-RSC7RequiredString `
                    -InputObject $cases[$index] `
                    -Name expectedEpochStatus) -or
            [bool]$results[$index].idleAuthorized -ne
                [bool](Get-RSC7Property `
                    -InputObject $cases[$index] `
                    -Name expectedIdleAuthorized `
                    -Required).Value) {
            throw [IO.InvalidDataException]::new(
                "C7 boundary case '$($actualIds[$index])' differs from its frozen outcome.")
        }
    }
}

function Assert-RSC7AuthenticationMatrix {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Contract
    )

    $authentication = (Get-RSC7Property `
            -InputObject $Contract `
            -Name authentication `
            -Required).Value
    $cases = Get-RSC7RequiredArray -InputObject $authentication -Name cases
    $actualCaseIds = [string[]]@(
        $cases | ForEach-Object {
            Get-RSC7RequiredString -InputObject $_ -Name caseId
        }
    )
    if (($actualCaseIds -join "`0") -cne ($script:ExpectedCaseIds -join "`0")) {
        throw [IO.InvalidDataException]::new(
            'C7 authentication cases differ from the frozen ordered 31-case matrix.')
    }
    Assert-RSC7ExactStringSet `
        -Actual $actualCaseIds `
        -Expected $script:ExpectedCaseIds `
        -Context 'C7 authentication case IDs'

    $results = Invoke-RSTerminalEngineC7AuthenticationMatrix -Contract $Contract
    for ($index = 0; $index -lt $cases.Count; ++$index) {
        $expected = (Get-RSC7Property `
                -InputObject $cases[$index] `
                -Name expected `
                -Required).Value
        $expectedDecisions = [string[]]@(
            Get-RSC7RequiredArray -InputObject $expected -Name decisions)
        $actualDecisions = [string[]]@($results[$index].decisions)
        if (($actualDecisions -join "`0") -cne
            ($expectedDecisions -join "`0") -or
            [string]$results[$index].epochStatus -cne
                (Get-RSC7RequiredString -InputObject $expected -Name epochStatus) -or
            [bool]$results[$index].idleAuthorized -ne
                [bool](Get-RSC7Property `
                    -InputObject $expected `
                    -Name idleAuthorized `
                    -Required).Value) {
            throw [IO.InvalidDataException]::new(
                "C7 authentication case '$($actualCaseIds[$index])' does not match its frozen outcome.")
        }
    }
    Assert-RSC7BoundaryMatrix -Contract $Contract
}

function Assert-RSC7Contract {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Contract,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    if ((Get-RSC7RequiredString -InputObject $Contract -Name formatId) -cne
        'red-salamander-terminal-engine-c7-proof-contract') {
        throw [IO.InvalidDataException]::new('C7 proof contract formatId is not v1.')
    }
    $version = (Get-RSC7Property `
            -InputObject $Contract `
            -Name version `
            -Required).Value
    if ([int]$version -ne 1) {
        throw [IO.InvalidDataException]::new('C7 proof contract version is not 1.')
    }

    $correction = (Get-RSC7Property `
            -InputObject $Contract `
            -Name correction `
            -Required).Value
    foreach ($name in @(
            'frozenBeforePerformance',
            'preservesMandatoryCorpus')) {
        if (-not [bool](Get-RSC7Property `
                -InputObject $correction `
                -Name $name `
                -Required).Value) {
            throw [IO.InvalidDataException]::new(
                "C7 proof correction '$name' must remain true.")
        }
    }
    foreach ($name in @(
            'priorCollectionReusable',
            'priorPerformanceReusable')) {
        if ([bool](Get-RSC7Property `
                -InputObject $correction `
                -Name $name `
                -Required).Value) {
            throw [IO.InvalidDataException]::new(
                "C7 proof correction '$name' must remain false.")
        }
    }

    $engineProbe = (Get-RSC7Property `
            -InputObject $Contract `
            -Name engineProbe `
            -Required).Value
    $scheduleIds = [string[]]@(
        Get-RSC7RequiredArray -InputObject $engineProbe -Name schedules)
    $expectedSchedules = [string[]]@(
        'all-at-once',
        'bytewise',
        'terminator-split')
    if (($scheduleIds -join "`0") -cne ($expectedSchedules -join "`0")) {
        throw [IO.InvalidDataException]::new(
            'C7 engine schedules differ from the frozen ordered set.')
    }
    if ((Get-RSC7RequiredString `
            -InputObject $engineProbe `
            -Name maxStateBytes) -cne '8192') {
        throw [IO.InvalidDataException]::new(
            'C7 max engine state must remain exactly 8192 bytes.')
    }

    Import-RSC7EvidenceModule -RepoRoot $RepoRoot
    $ordering = (Get-RSC7Property `
            -InputObject $engineProbe `
            -Name requiredOrderingFixture `
            -Required).Value
    $orderingJcs = Get-RSC7RequiredString -InputObject $ordering -Name stateJcs
    $orderingObject = $orderingJcs |
        ConvertFrom-Json -AsHashtable -Depth 20 -NoEnumerate -DateKind String
    $canonicalOrdering = [Text.Encoding]::UTF8.GetString(
        (ConvertTo-RSJcsUtf8Bytes -InputObject $orderingObject))
    if ($canonicalOrdering -cne $orderingJcs -or
        (Get-RSJcsSha256 -InputObject $orderingObject) -cne
            (Get-RSC7RequiredString -InputObject $ordering -Name stateSha256)) {
        throw [IO.InvalidDataException]::new(
            'C7 required ordering fixture state is not canonical or hash-bound.')
    }

    $providerEvidence = (Get-RSC7Property `
            -InputObject $Contract `
            -Name providerEvidence `
            -Required).Value
    if ((Get-RSC7RequiredString `
            -InputObject $providerEvidence `
            -Name maxProofBytes) -cne '32768') {
        throw [IO.InvalidDataException]::new(
            'C7 provider proof limit must remain exactly 32768 bytes.')
    }
    Assert-RSC7AuthenticationMatrix -Contract $Contract
}

function Read-RSTerminalEngineC7ProofContract {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $resolvedRepoRoot = [IO.Path]::GetFullPath($RepoRoot)
    $contractPath = Join-Path $resolvedRepoRoot $script:ContractLeaf
    $schemaPath = Join-Path $resolvedRepoRoot $script:SchemaLeaf
    if (-not [IO.File]::Exists($contractPath) -or
        -not [IO.File]::Exists($schemaPath)) {
        throw [IO.FileNotFoundException]::new(
            'The frozen C7 proof contract or its closed schema is missing.')
    }
    $raw = [IO.File]::ReadAllText(
        $contractPath,
        [Text.UTF8Encoding]::new($false, $true))
    $valid = Test-Json `
        -Json $raw `
        -SchemaFile $schemaPath `
        -ErrorAction Stop
    if (-not $valid) {
        throw [IO.InvalidDataException]::new(
            'C7 proof contract does not satisfy its closed v1 schema.')
    }
    $contract = $raw |
        ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    Assert-RSC7Contract -Contract $contract -RepoRoot $resolvedRepoRoot
    Import-RSC7EvidenceModule -RepoRoot $resolvedRepoRoot
    return [pscustomobject]@{
        Contract = $contract
        ContractSha256 = Get-RSJcsSha256 -InputObject $contract
        ContractPath = $contractPath
        SchemaPath = $schemaPath
    }
}

function Get-RSTerminalEngineC7ExpectedStateJcs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $context = Read-RSTerminalEngineC7ProofContract -RepoRoot $RepoRoot
    $engineProbe = (Get-RSC7Property `
            -InputObject $context.Contract `
            -Name engineProbe `
            -Required).Value
    $expectedState = (Get-RSC7Property `
            -InputObject $engineProbe `
            -Name expectedState `
            -Required).Value
    Import-RSC7EvidenceModule -RepoRoot $RepoRoot
    return [Text.Encoding]::UTF8.GetString(
        (ConvertTo-RSJcsUtf8Bytes -InputObject $expectedState))
}

function Assert-RSC7CanonicalState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StateJcs,

        [Parameter(Mandatory = $true)]
        [UInt64]$MaxBytes,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedJcs,

        [Parameter(Mandatory = $true)]
        [string]$Context,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $bytes = [Text.UTF8Encoding]::new($false, $true).GetBytes($StateJcs)
    if ([UInt64]$bytes.Length -gt $MaxBytes) {
        throw [IO.InvalidDataException]::new(
            "$Context exceeds the frozen $MaxBytes-byte limit.")
    }
    $parsed = $StateJcs |
        ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    Import-RSC7EvidenceModule -RepoRoot $RepoRoot
    $canonical = [Text.Encoding]::UTF8.GetString(
        (ConvertTo-RSJcsUtf8Bytes -InputObject $parsed))
    if ($canonical -cne $StateJcs) {
        throw [IO.InvalidDataException]::new(
            "$Context is not exact RFC 8785 canonical JSON.")
    }
    if ($StateJcs -cne $ExpectedJcs) {
        throw [IO.InvalidDataException]::new(
            "$Context differs from the frozen C7 engine observation.")
    }
    return [pscustomobject]@{
        Sha256 = Get-RSJcsSha256 -InputObject $parsed
        SizeBytes = [string]$bytes.Length
    }
}

function New-RSTerminalEngineC7ProviderProof {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId,

        [Parameter(Mandatory = $true)]
        [string]$PinSha256,

        [Parameter(Mandatory = $true)]
        [string]$SourceArchiveSha256,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Release', 'ASan Debug')]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$AdapterIdentitySha256,

        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$StateJcsBySchedule,

        [Parameter(Mandatory = $true)]
        [string]$OrderingStateJcs,

        [Parameter(Mandatory = $true)]
        [string[]]$OrderingActionFlags
    )

    if ($CandidateId -cnotmatch '^[a-z0-9][a-z0-9._-]{0,63}$') {
        throw [IO.InvalidDataException]::new(
            'C7 candidateId is not a bounded identifier.')
    }
    Assert-RSC7Sha256 -Value $PinSha256 -Context 'C7 pin identity'
    Assert-RSC7Sha256 `
        -Value $SourceArchiveSha256 `
        -Context 'C7 source archive identity'
    Assert-RSC7Sha256 `
        -Value $AdapterIdentitySha256 `
        -Context 'C7 adapter identity'

    if (($Platform -ceq 'ARM64' -and $Configuration -cne 'Release')) {
        throw [IO.InvalidDataException]::new(
            'C7 proof supports only the frozen ARM64 Release lane.')
    }

    $contractContext = Read-RSTerminalEngineC7ProofContract -RepoRoot $RepoRoot
    $contract = $contractContext.Contract
    $engineProbe = (Get-RSC7Property `
            -InputObject $contract `
            -Name engineProbe `
            -Required).Value
    $scheduleIds = [string[]]@(
        Get-RSC7RequiredArray -InputObject $engineProbe -Name schedules)
    Assert-RSC7ExactStringSet `
        -Actual ([string[]]@($StateJcsBySchedule.Keys)) `
        -Expected $scheduleIds `
        -Context 'C7 engine-state schedules'
    $maxStateBytes = Get-RSC7UInt64 `
        -Value (Get-RSC7RequiredString `
            -InputObject $engineProbe `
            -Name maxStateBytes) `
        -Context 'C7 engine-state byte limit'
    $expectedStateJcs = Get-RSTerminalEngineC7ExpectedStateJcs `
        -RepoRoot $RepoRoot
    $scheduleStates = [Collections.Generic.List[object]]::new()
    foreach ($scheduleId in $scheduleIds) {
        $value = $StateJcsBySchedule[$scheduleId]
        if ($value -isnot [string]) {
            throw [IO.InvalidDataException]::new(
                "C7 engine state '$scheduleId' must be a canonical JSON string.")
        }
        $state = Assert-RSC7CanonicalState `
            -StateJcs ([string]$value) `
            -MaxBytes $maxStateBytes `
            -ExpectedJcs $expectedStateJcs `
            -Context "C7 '$scheduleId' engine state" `
            -RepoRoot $RepoRoot
        $scheduleStates.Add([ordered]@{
                scheduleId = $scheduleId
                sha256 = [string]$state.Sha256
                sizeBytes = [string]$state.SizeBytes
            })
    }

    $ordering = (Get-RSC7Property `
            -InputObject $engineProbe `
            -Name requiredOrderingFixture `
            -Required).Value
    $expectedOrderingJcs = Get-RSC7RequiredString `
        -InputObject $ordering `
        -Name stateJcs
    $orderingState = Assert-RSC7CanonicalState `
        -StateJcs $OrderingStateJcs `
        -MaxBytes ([UInt64]1024) `
        -ExpectedJcs $expectedOrderingJcs `
        -Context 'C7 write-response-order state' `
        -RepoRoot $RepoRoot
    $requiredActionFlags = [string[]]@(
        Get-RSC7RequiredArray `
            -InputObject $ordering `
            -Name requiredActionFlags)
    Assert-RSC7ExactStringSet `
        -Actual $OrderingActionFlags `
        -Expected $requiredActionFlags `
        -Context 'C7 write-response-order action flags'

    Import-RSC7EvidenceModule -RepoRoot $RepoRoot
    $matrix = Invoke-RSTerminalEngineC7AuthenticationMatrix -Contract $contract
    $boundaryMatrix = Invoke-RSTerminalEngineC7BoundaryMatrix `
        -Contract $contract
    $positiveCount = @(
        $matrix | Where-Object classification -CEQ 'positive').Count
    $negativeCount = $matrix.Count - $positiveCount
    $arguments = (Get-RSC7Property `
            -InputObject $engineProbe `
            -Name arguments `
            -Required).Value
    $proof = [ordered]@{
        formatId = 'red-salamander-terminal-engine-c7-proof'
        version = 1
        contractSha256 = [string]$contractContext.ContractSha256
        candidateId = $CandidateId
        pinSha256 = $PinSha256
        sourceArchiveSha256 = $SourceArchiveSha256
        lane = [ordered]@{
            platform = $Platform
            configuration = $Configuration
        }
        adapterIdentitySha256 = $AdapterIdentitySha256
        engine = [ordered]@{
            actionId = Get-RSC7RequiredString `
                -InputObject $engineProbe `
                -Name actionId
            argumentsSha256 = Get-RSJcsSha256 -InputObject $arguments
            fixedStreamSha256 = Get-RSC7RequiredString `
                -InputObject $engineProbe `
                -Name fixedStreamSha256
            scheduleStates = [object[]]$scheduleStates.ToArray()
        }
        ordering = [ordered]@{
            actionId = Get-RSC7RequiredString `
                -InputObject $ordering `
                -Name actionId
            fixtureId = Get-RSC7RequiredString `
                -InputObject $ordering `
                -Name fixtureId
            scheduleId = Get-RSC7RequiredString `
                -InputObject $ordering `
                -Name scheduleId
            stateSha256 = [string]$orderingState.Sha256
            actionFlagsSha256 = Get-RSJcsSha256 `
                -InputObject ([string[]]$requiredActionFlags)
        }
        authentication = [ordered]@{
            caseCount = [string]$matrix.Count
            positiveCount = [string]$positiveCount
            negativeCount = [string]$negativeCount
            matrixSha256 = Get-RSJcsSha256 -InputObject $matrix
            boundaryCaseCount = [string]$boundaryMatrix.Count
            boundaryMatrixSha256 = Get-RSJcsSha256 `
                -InputObject $boundaryMatrix
        }
        verdict = 'verified'
    }
    $proofBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $proof
    $providerEvidence = (Get-RSC7Property `
            -InputObject $contract `
            -Name providerEvidence `
            -Required).Value
    $maxProofBytes = Get-RSC7UInt64 `
        -Value (Get-RSC7RequiredString `
            -InputObject $providerEvidence `
            -Name maxProofBytes) `
        -Context 'C7 provider proof byte limit'
    if ([UInt64]$proofBytes.Length -gt $maxProofBytes) {
        throw [IO.InvalidDataException]::new(
            "C7 provider proof exceeds the frozen $maxProofBytes-byte limit.")
    }
    $proofSha256 = Get-RSJcsSha256 -InputObject $proof
    return [pscustomobject]@{
        Proof = $proof
        EvidenceObject = [ordered]@{
            evidenceId = $script:EvidenceId
            sha256 = $proofSha256
            sizeBytes = [string]$proofBytes.Length
            resultCategory = $script:ResultCategory
        }
    }
}

function Assert-RSTerminalEngineC7EvidenceObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$EvidenceObject,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    Assert-RSC7ExactProperties `
        -InputObject $EvidenceObject `
        -Expected @('evidenceId', 'sha256', 'sizeBytes', 'resultCategory') `
        -Context 'C7 provider evidence descriptor'
    if ((Get-RSC7RequiredString `
            -InputObject $EvidenceObject `
            -Name evidenceId) -cne $script:EvidenceId -or
        (Get-RSC7RequiredString `
            -InputObject $EvidenceObject `
            -Name resultCategory) -cne $script:ResultCategory) {
        throw [IO.InvalidDataException]::new(
            'C7 provider evidence descriptor has the wrong identity or status.')
    }
    Assert-RSC7Sha256 `
        -Value (Get-RSC7RequiredString `
            -InputObject $EvidenceObject `
            -Name sha256) `
        -Context 'C7 provider evidence hash'
    $sizeBytes = Get-RSC7UInt64 `
        -Value (Get-RSC7RequiredString `
            -InputObject $EvidenceObject `
            -Name sizeBytes) `
        -Context 'C7 provider evidence size'
    $contractContext = Read-RSTerminalEngineC7ProofContract -RepoRoot $RepoRoot
    $providerEvidence = (Get-RSC7Property `
            -InputObject $contractContext.Contract `
            -Name providerEvidence `
            -Required).Value
    $maxProofBytes = Get-RSC7UInt64 `
        -Value (Get-RSC7RequiredString `
            -InputObject $providerEvidence `
            -Name maxProofBytes) `
        -Context 'C7 provider proof byte limit'
    if ($sizeBytes -eq 0 -or $sizeBytes -gt $maxProofBytes) {
        throw [IO.InvalidDataException]::new(
            'C7 provider evidence size is zero or exceeds its frozen bound.')
    }
}

function Add-RSTerminalEngineC7ProviderEvidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$CandidateResult,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$SourceArchiveSha256,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Release', 'ASan Debug')]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$AdapterIdentitySha256,

        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$StateJcsBySchedule,

        [Parameter(Mandatory = $true)]
        [string]$OrderingStateJcs,

        [Parameter(Mandatory = $true)]
        [string[]]$OrderingActionFlags
    )

    if ([string]$CandidateResult['outcome'] -cne 'complete') {
        throw [IO.InvalidDataException]::new(
            'Provider cannot attach C7 score proof to an incomplete candidate.')
    }
    $candidateId = [string]$CandidateResult['candidateId']
    $pinSha256 = [string]$CandidateResult['pinSha256']
    $byId = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($record in @($CandidateResult['evidenceObjects'])) {
        $evidenceId = Get-RSC7RequiredString `
            -InputObject $record `
            -Name evidenceId
        if (-not $byId.TryAdd($evidenceId, $record)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$candidateId' repeats evidence object '$evidenceId'.")
        }
    }
    if ($byId.ContainsKey($script:EvidenceId)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' driver attempted to inject provider-owned '$($script:EvidenceId)'.")
    }

    $providerProof = New-RSTerminalEngineC7ProviderProof `
        -RepoRoot $RepoRoot `
        -CandidateId $candidateId `
        -PinSha256 $pinSha256 `
        -SourceArchiveSha256 $SourceArchiveSha256 `
        -Platform $Platform `
        -Configuration $Configuration `
        -AdapterIdentitySha256 $AdapterIdentitySha256 `
        -StateJcsBySchedule $StateJcsBySchedule `
        -OrderingStateJcs $OrderingStateJcs `
        -OrderingActionFlags $OrderingActionFlags
    [void]$byId.Add(
        $script:EvidenceId,
        $providerProof.EvidenceObject)
    $ids = [string[]]@($byId.Keys)
    [Array]::Sort($ids, [StringComparer]::Ordinal)
    $ordered = [Collections.Generic.List[object]]::new()
    foreach ($evidenceId in $ids) {
        $ordered.Add($byId[$evidenceId])
    }
    $CandidateResult['evidenceObjects'] = [object[]]$ordered.ToArray()
    return $providerProof
}

Export-ModuleMember -Function `
    Add-RSTerminalEngineC7ProviderEvidence, `
    Assert-RSTerminalEngineC7EvidenceObject, `
    Get-RSTerminalEngineC7ExpectedStateJcs, `
    Get-RSTerminalEngineC7FixedStreamBytes, `
    Invoke-RSTerminalEngineC7AuthenticationMatrix, `
    Invoke-RSTerminalEngineC7BoundaryMatrix, `
    New-RSTerminalEngineC7ProviderProof, `
    Read-RSTerminalEngineC7ProofContract
