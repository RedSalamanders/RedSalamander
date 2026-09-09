[CmdletBinding(DefaultParameterSetName = 'Verify')]
param(
    [string]$Corpus = (Join-Path $PSScriptRoot '..\..\Specs\Terminal\TerminalEngineFixtureCorpus.v1.json'),

    [string]$Header = (Join-Path $PSScriptRoot 'Generated\TerminalEngineFixtures.v1.h'),

    [Parameter(ParameterSetName = 'Verify')]
    [switch]$Verify,

    [Parameter(Mandatory, ParameterSetName = 'Update')]
    [switch]$Update
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\TerminalEvidence.psm1') -Force
$script:ZlibBombBase64 = $null

function Get-RSSha256Hex {
    param([Parameter(Mandatory)][byte[]]$Bytes)

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToHexString($sha.ComputeHash($Bytes)).ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function ConvertTo-RSHex {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    return [Convert]::ToHexString($Bytes).ToLowerInvariant()
}

function New-RSLiteralTextStep {
    param([Parameter(Mandatory)][string]$Text)
    return [ordered]@{
        op = 'LiteralHex'
        hex = ConvertTo-RSHex -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($Text))
        byteValue = 0
        repeatCount = [uint64]0
    }
}

function New-RSLiteralHexStep {
    param([Parameter(Mandatory)][string]$Hex)
    if ($Hex -cnotmatch '^(?:[0-9a-f]{2})+$') {
        throw "Literal hex is invalid: $Hex"
    }
    return [ordered]@{
        op = 'LiteralHex'
        hex = $Hex
        byteValue = 0
        repeatCount = [uint64]0
    }
}

function New-RSRepeatByteStep {
    param(
        [Parameter(Mandatory)][ValidateRange(0, 255)][int]$ByteValue,
        [Parameter(Mandatory)][uint64]$RepeatCount
    )
    return [ordered]@{
        op = 'RepeatByte'
        hex = ''
        byteValue = [byte]$ByteValue
        repeatCount = $RepeatCount
    }
}

function Get-RSZlibBombBase64 {
    if ($null -ne $script:ZlibBombBase64) {
        return $script:ZlibBombBase64
    }

    $compressed = [IO.MemoryStream]::new()
    try {
        $zlib = [IO.Compression.ZLibStream]::new(
            $compressed,
            [IO.Compression.CompressionLevel]::SmallestSize,
            $true)
        try {
            $zeroes = [byte[]]::new(65536)
            [uint64]$remaining = 1048577
            while ($remaining -ne 0) {
                $count = [int][Math]::Min([uint64]$zeroes.Length, $remaining)
                $zlib.Write($zeroes, 0, $count)
                $remaining -= [uint64]$count
            }
        }
        finally {
            $zlib.Dispose()
        }
        $script:ZlibBombBase64 = [Convert]::ToBase64String($compressed.ToArray())
        return $script:ZlibBombBase64
    }
    finally {
        $compressed.Dispose()
    }
}

function Get-RSFixtureInputPlan {
    param([Parameter(Mandatory)][string]$FixtureId)

    $esc = [char]0x1b
    switch ($FixtureId) {
        'alternate-screen-roundtrip' {
            return @(
                (New-RSLiteralTextStep "primary${esc}[?1049halternate${esc}[?1049l"))
        }
        'bracketed-paste-modes' {
            return @((New-RSLiteralTextStep "${esc}[?2004h"))
        }
        'color-style-matrix' {
            return @((New-RSLiteralTextStep (
                        "${esc}[1;2;3;4:3;5;7;9;31;48;5;200mA" +
                        "${esc}[38;2;1;2;3;48;2;4;5;6mB${esc}[0mC")))
        }
        'cursor-and-palette' {
            return @((New-RSLiteralTextStep (
                        "${esc}[5;7H${esc}[?25l${esc}]4;1;rgb:11/22/33${esc}\" +
                        "${esc}]10;rgb:44/55/66${esc}\")))
        }
        'focus-and-mouse-matrix' {
            return @((New-RSLiteralTextStep (
                        "${esc}[?1000h${esc}[?1002h${esc}[?1003h${esc}[?1005h" +
                        "${esc}[?1006h${esc}[?1015h${esc}[?1016h${esc}[?1004h")))
        }
        'hyperlink-boundaries' {
            return @(
                (New-RSLiteralTextStep "${esc}]8;"),
                (New-RSRepeatByteStep -ByteValue 112 -RepeatCount 1024),
                (New-RSLiteralTextStep ';u'),
                (New-RSLiteralTextStep "${esc}\p${esc}]8;"),
                (New-RSRepeatByteStep -ByteValue 113 -RepeatCount 1025),
                (New-RSLiteralTextStep ';u'),
                (New-RSLiteralTextStep "${esc}\q${esc}]8;;"),
                (New-RSRepeatByteStep -ByteValue 117 -RepeatCount 8192),
                (New-RSLiteralTextStep "${esc}\u${esc}]8;;"),
                (New-RSRepeatByteStep -ByteValue 118 -RepeatCount 8193),
                (New-RSLiteralTextStep "${esc}\v"))
        }
        'kitty-delete-generation' {
            return @((New-RSLiteralTextStep (
                        "${esc}_Gf=32,s=1,v=1,i=7,a=T;/wAA/w==${esc}\" +
                        "${esc}_Ga=p,i=7,p=11,q=2;${esc}\" +
                        "${esc}_Ga=d,d=i,i=7;${esc}\")))
        }
        'kitty-dimension-overflow' {
            return @((New-RSLiteralTextStep (
                        "${esc}_Gf=32,s=4294967295,v=4294967295,i=8,a=T;AA==${esc}\")))
        }
        'kitty-direct-png' {
            return @((New-RSLiteralTextStep (
                        "${esc}_Gf=100,i=13,a=T;" +
                        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ' +
                        'AAAADUlEQVR42mP4z8DwHwAFAAH/VscvDQAAAABJRU5ErkJggg==' +
                        "${esc}\")))
        }
        'kitty-direct-rgb' {
            return @((New-RSLiteralTextStep (
                        "${esc}_Gf=24,s=1,v=1,i=1,a=T;/wAA${esc}\")))
        }
        'kitty-direct-rgba-chunked' {
            return @((New-RSLiteralTextStep (
                        "${esc}_Gf=32,s=1,v=1,i=2,m=1;/wAA${esc}\" +
                        "${esc}_Gi=2,m=0;/w==${esc}\")))
        }
        'kitty-pixel-boundary' {
            return @((New-RSLiteralTextStep 'RSFIXTURE:kitty-pixel-boundary'))
        }
        'kitty-png-bomb' {
            return @((New-RSLiteralTextStep (
                        "${esc}_Gf=100,i=9,a=T;" +
                        'iVBORw0KGgoAAAANSUhEUgAA////AAD///8IAQAAAAAA' +
                        "${esc}\")))
        }
        'kitty-replacement-transient' {
            return @((New-RSLiteralTextStep 'RSFIXTURE:kitty-replacement-transient'))
        }
        'kitty-three-layers-crop' {
            return @((New-RSLiteralTextStep (
                        "${esc}_Gf=32,s=1,v=1,i=3,a=T;/wAA/w==${esc}\" +
                        "${esc}_Ga=p,i=3,p=1,z=-2,x=0,y=0,w=1,h=1;${esc}\" +
                        "${esc}_Ga=p,i=3,p=2,z=-1,x=0,y=0,w=1,h=1;${esc}\" +
                        "${esc}_Ga=p,i=3,p=3,z=2,x=0,y=0,w=1,h=1;${esc}\")))
        }
        'kitty-zlib-bomb' {
            $bomb = Get-RSZlibBombBase64
            return @((New-RSLiteralTextStep (
                        "${esc}_Gf=32,o=z,s=512,v=512,i=10,a=T;$bomb${esc}\")))
        }
        'malformed-utf8-vt' {
            return @((New-RSLiteralHexStep (
                        'f0808080c0afeda0801b5b39393939393939393939393939393939393939393939393939393939393939')))
        }
        'nonkitty-generic-boundary' {
            return @(
                (New-RSLiteralTextStep "${esc}P"),
                (New-RSRepeatByteStep -ByteValue 97 -RepeatCount 65536),
                (New-RSLiteralTextStep "${esc}\"),
                (New-RSLiteralTextStep "${esc}P"),
                (New-RSRepeatByteStep -ByteValue 98 -RepeatCount 65537),
                (New-RSLiteralTextStep "${esc}\Z"))
        }
        'osc52-boundaries' {
            return @(
                (New-RSLiteralTextStep "${esc}]52;c;"),
                (New-RSRepeatByteStep -ByteValue 65 -RepeatCount 1048576),
                (New-RSLiteralTextStep ([string][char]7)),
                (New-RSLiteralTextStep "${esc}]52;c;"),
                (New-RSRepeatByteStep -ByteValue 65 -RepeatCount 1048577),
                (New-RSLiteralTextStep ([string][char]7)))
        }
        'pending-effects-boundary' {
            return @((New-RSLiteralTextStep 'RSFIXTURE:pending-effects-boundary'))
        }
        'resize-reflow-selection' {
            return @((New-RSLiteralTextStep (
                        "logical-line-1 e$([char]0x301) $([char]::ConvertFromUtf32(0x1f642))`r`nlogical-line-2 wide:$([char]0x754c)")))
        }
        'scrollback-byte-boundary' {
            return @(
                (New-RSRepeatByteStep -ByteValue 10 -RepeatCount 67108865))
        }
        'snapshot-dirty-retry' {
            return @((New-RSLiteralTextStep "snapshot-one`r`nsnapshot-two"))
        }
        'snapshot-mutation-barrier' {
            return @((New-RSLiteralTextStep (
                        "borrowed${esc}_Gf=32,s=1,v=1,i=12,a=T;/wAA/w==${esc}\")))
        }
        'synchronized-output' {
            return @((New-RSLiteralTextStep (
                        "${esc}[?2026hfirst`r`nsecond${esc}[?2026l")))
        }
        'teardown-callback-quiet' {
            return @((New-RSLiteralTextStep 'RSFIXTURE:teardown-callback-quiet'))
        }
        'title-cwd-boundaries' {
            return @(
                (New-RSLiteralTextStep "${esc}]2;"),
                (New-RSRepeatByteStep -ByteValue 116 -RepeatCount 4096),
                (New-RSLiteralTextStep ([string][char]7)),
                (New-RSLiteralTextStep "${esc}]2;"),
                (New-RSRepeatByteStep -ByteValue 117 -RepeatCount 4097),
                (New-RSLiteralTextStep ([string][char]7)),
                (New-RSLiteralTextStep "${esc}]7;file://localhost/"),
                (New-RSRepeatByteStep -ByteValue 99 -RepeatCount 131072),
                (New-RSLiteralTextStep ([string][char]7)),
                (New-RSLiteralTextStep "${esc}]7;file://localhost/"),
                (New-RSRepeatByteStep -ByteValue 100 -RepeatCount 131073),
                (New-RSLiteralTextStep ([string][char]7)))
        }
        'unicode-grapheme-wide' {
            return @((New-RSLiteralTextStep (
                        "e$([char]0x301) $([char]0x754c) " +
                        "$([char]::ConvertFromUtf32(0x1f469))$([char]0x200d)$([char]::ConvertFromUtf32(0x1f4bb)) " +
                        "$([char]0x2764)$([char]0xfe0f)")))
        }
        'write-response-order' {
            return @((New-RSLiteralTextStep (
                        "${esc}[5n${esc}[6n${esc}]2;ordered${esc}\")))
        }
        default {
            throw "No exact input plan is defined for fixture '$FixtureId'."
        }
    }
}

function Get-RSExpectedCheck {
    param(
        [Parameter(Mandatory)][string]$FixtureId,
        [Parameter(Mandatory)][string]$Expectation
    )

    $values = switch ($FixtureId) {
        'alternate-screen-roundtrip' {
            [ordered]@{ alternateDiscarded = $true; primarySurvived = $true; primaryText = 'primary' }
        }
        'bracketed-paste-modes' {
            [ordered]@{ disabled = 'payload'; enabled = '\u001b[200~payload\u001b[201~'; ordered = $true }
        }
        'color-style-matrix' {
            [ordered]@{ indexed = $true; styles = 'bold,faint,italic,underline,blink,inverse,strike'; truecolor = '1,2,3/4,5,6' }
        }
        'cursor-and-palette' {
            [ordered]@{ cursorVisible = $false; palette1 = '11,22,33'; defaultForeground = '44,55,66' }
        }
        'focus-and-mouse-matrix' {
            [ordered]@{ focus = $true; modes = '1000,1002,1003,1005,1006,1015,1016'; ordered = $true }
        }
        'hyperlink-boundaries' {
            [ordered]@{ parameterBoundary = '1024'; recordBoundary = '16384'; uriBoundary = '8192'; plusOneDiscarded = $true }
        }
        'kitty-delete-generation' {
            [ordered]@{ imageGenerationChanged = $true; placementGenerationChanged = $true; placementOnlyMutationDistinct = $true }
        }
        'kitty-dimension-overflow' {
            [ordered]@{ allocationAttempted = $false; checkedArithmetic = $true; rejected = $true }
        }
        'kitty-direct-png' {
            [ordered]@{
                format = 'png'
                height = '1'
                imageId = '13'
                immutableAfterDelete = $true
                pixel = 'ff0000ff'
                snapshotOwned = $true
                width = '1'
            }
        }
        'kitty-direct-rgb' {
            [ordered]@{ format = 'rgb'; height = '1'; immutable = $true; pixel = 'ff0000'; width = '1' }
        }
        'kitty-direct-rgba-chunked' {
            [ordered]@{ atomicGenerationDelta = '1'; format = 'rgba'; pixel = 'ff0000ff' }
        }
        'kitty-pixel-boundary' {
            [ordered]@{ boundary = '16000000'; boundaryAcceptedWhenLedgersFit = $true; plusOnePredecodeRejected = $true }
        }
        'kitty-png-bomb' {
            [ordered]@{ decodedAllocationAttempted = $false; dimensionsPreflighted = $true; rejected = $true }
        }
        'kitty-replacement-transient' {
            [ordered]@{ atomicFailure = $true; oldRetainedUntilPublication = $true; simultaneousReservations = 'old,new,compressed,decoded,conversion,publication' }
        }
        'kitty-three-layers-crop' {
            [ordered]@{ crop = '0,0,1,1'; layers = '-2,-1,2'; placementCount = '3' }
        }
        'kitty-zlib-bomb' {
            [ordered]@{ boundedInflater = $true; outputBeyondReservation = $false; processAlive = $true }
        }
        'malformed-utf8-vt' {
            [ordered]@{ deterministic = $true; processAlive = $true; replacement = 'fffd' }
        }
        'nonkitty-generic-boundary' {
            [ordered]@{ boundary = '65536'; boundaryAccepted = $true; plusOneDiscarded = $true }
        }
        'osc52-boundaries' {
            [ordered]@{ decodedBoundary = '786432'; encodedBoundary = '1048576'; partialEffect = $false; plusOneDiscarded = $true }
        }
        'pending-effects-boundary' {
            [ordered]@{ byteBoundary = '2097152'; descriptorBoundary = '256'; plusOneAtomic = $true }
        }
        'resize-reflow-selection' {
            [ordered]@{ graphemeStable = $true; logicalLines = '2'; selectionStable = $true; wideCellStable = $true }
        }
        'scrollback-byte-boundary' {
            [ordered]@{ boundary = '67108864'; plusOnePreallocationHandled = $true; unit = 'bytes' }
        }
        'snapshot-dirty-retry' {
            [ordered]@{ failedPublicationRetried = $true; fullSnapshotOnRetry = $true; notificationLost = $false }
        }
        'snapshot-mutation-barrier' {
            [ordered]@{ borrowedAfterEnd = 'invalid'; copiedSnapshotStable = $true; mutationBarrier = $true }
        }
        'synchronized-output' {
            [ordered]@{ publishedDuringSync = $false; releaseAtomic = $true; retainedDirty = $true }
        }
        'teardown-callback-quiet' {
            [ordered]@{ callbackAfterDetach = '0'; producersStoppedFirst = $true; quietBeforeUnload = $true }
        }
        'title-cwd-boundaries' {
            [ordered]@{ cwdBoundary = '131072'; plusOneDiscarded = $true; titleBoundary = '4096' }
        }
        'unicode-grapheme-wide' {
            [ordered]@{ combiningStable = $true; variationStable = $true; wideContinuationStable = $true; zwjStable = $true }
        }
        'write-response-order' {
            [ordered]@{ admittedOrder = 'protocol-reply,input,resize,side-channel'; reordered = $false }
        }
        default {
            throw "No expected check is defined for fixture '$FixtureId'."
        }
    }
    $isLimitFixture = $FixtureId -in @(
        'hyperlink-boundaries',
        'kitty-dimension-overflow',
        'kitty-pixel-boundary',
        'kitty-png-bomb',
        'kitty-replacement-transient',
        'kitty-zlib-bomb',
        'nonkitty-generic-boundary',
        'osc52-boundaries',
        'pending-effects-boundary',
        'scrollback-byte-boundary',
        'title-cwd-boundaries')
    $resourceValues = [ordered]@{
        allReservationsReleased = $true
        checkedArithmetic = $true
        limitFixture = $isLimitFixture
        peakWithinFrozenLimits = $true
        rejectionBeforeAllocation = $isLimitFixture
    }
    $stateValue = [Text.Encoding]::UTF8.GetString(
        (ConvertTo-RSJcsUtf8Bytes -InputObject $values))
    $resourceValue = [Text.Encoding]::UTF8.GetString(
        (ConvertTo-RSJcsUtf8Bytes -InputObject $resourceValues))
    return @(
        [ordered]@{
            checkId = 'resource-trace'
            expectedValue = $resourceValue
            expectedSha256 = Get-RSSha256Hex -Bytes (
                [Text.UTF8Encoding]::new($false).GetBytes($resourceValue))
            requirement = $Expectation
        },
        [ordered]@{
            checkId = 'state'
            expectedValue = $stateValue
            expectedSha256 = Get-RSSha256Hex -Bytes (
                [Text.UTF8Encoding]::new($false).GetBytes($stateValue))
            requirement = $Expectation
        })
}

function New-RSFixtureAction {
    param(
        [Parameter(Mandatory)][string]$ActionId,
        [Parameter(Mandatory)][AllowEmptyCollection()][object]$Arguments
    )

    $argumentsJcs = [Text.Encoding]::UTF8.GetString(
        (ConvertTo-RSJcsUtf8Bytes -InputObject $Arguments))
    return [pscustomobject]@{
        ActionId = $ActionId
        ArgumentsJcs = $argumentsJcs
        Sha256 = Get-RSJcsSha256 -InputObject ([ordered]@{
                actionId = $ActionId
                arguments = $Arguments
            })
    }
}

function Get-RSFixtureActions {
    param([Parameter(Mandatory)][string]$FixtureId)

    $actions = [Collections.Generic.List[object]]::new()
    $actions.Add((New-RSFixtureAction -ActionId 'feed-current-schedule' -Arguments (
                [ordered]@{ input = 'inputPlan'; schedule = 'current' })))
    switch ($FixtureId) {
        'bracketed-paste-modes' {
            $actions.Add((New-RSFixtureAction -ActionId 'encode-paste-disabled' -Arguments (
                        [ordered]@{ payloadHex = '7061796c6f6164' })))
            $actions.Add((New-RSFixtureAction -ActionId 'encode-paste-enabled' -Arguments (
                        [ordered]@{ payloadHex = '7061796c6f6164' })))
        }
        'focus-and-mouse-matrix' {
            $actions.Add((New-RSFixtureAction -ActionId 'encode-focus-gained-lost' -Arguments (
                        [ordered]@{ order = 'gained,lost' })))
            $actions.Add((New-RSFixtureAction -ActionId 'encode-mouse-matrix' -Arguments (
                        [ordered]@{ modes = '1000,1002,1003,1005,1006,1015,1016' })))
        }
        'hyperlink-boundaries' {
            $actions.Add((New-RSFixtureAction -ActionId 'probe-osc8-limits' -Arguments (
                        [ordered]@{ parameter = '1024'; record = '16384'; uri = '8192' })))
        }
        'kitty-direct-png' {
            $actions.Add((New-RSFixtureAction -ActionId 'snapshot-static-kitty' -Arguments (
                        [ordered]@{
                            expectedDecodedRgbaHex = 'ff0000ff'
                            imageId = '13'
                            mutation = 'delete-image-and-placement'
                            requireOwnedCopyStable = $true
                        })))
        }
        'kitty-pixel-boundary' {
            $actions.Add((New-RSFixtureAction -ActionId 'probe-kitty-pixel-limit' -Arguments (
                        [ordered]@{ boundary = '16000000'; plusOne = '16000001' })))
        }
        'kitty-replacement-transient' {
            $actions.Add((New-RSFixtureAction -ActionId 'probe-kitty-replacement-ledger' -Arguments (
                        [ordered]@{ forceFailureAfterNewReservation = $true })))
        }
        'nonkitty-generic-boundary' {
            $actions.Add((New-RSFixtureAction -ActionId 'probe-generic-escape-limit' -Arguments (
                        [ordered]@{ boundary = '65536'; plusOne = '65537' })))
        }
        'osc52-boundaries' {
            $actions.Add((New-RSFixtureAction -ActionId 'probe-osc52-limits' -Arguments (
                        [ordered]@{ decoded = '786432'; encoded = '1048576' })))
        }
        'pending-effects-boundary' {
            $actions.Add((New-RSFixtureAction -ActionId 'probe-pending-effect-ledger' -Arguments (
                        [ordered]@{ bytes = '2097152'; descriptors = '256' })))
        }
        'resize-reflow-selection' {
            $actions.Add((New-RSFixtureAction -ActionId 'select-logical-range' -Arguments (
                        [ordered]@{ end = '1:12'; start = '0:0' })))
            $actions.Add((New-RSFixtureAction -ActionId 'resize-sequence' -Arguments (
                        [ordered]@{ dimensions = '80x24,20x24,100x24' })))
        }
        'scrollback-byte-boundary' {
            $actions.Add((New-RSFixtureAction -ActionId 'probe-terminal-text-ledger' -Arguments (
                        [ordered]@{ boundary = '67108864'; plusOne = '67108865' })))
        }
        'snapshot-dirty-retry' {
            $actions.Add((New-RSFixtureAction -ActionId 'snapshot-copy-fail-retry' -Arguments (
                        [ordered]@{ failAfterCells = '1'; requireFullRetry = $true })))
        }
        'snapshot-mutation-barrier' {
            $actions.Add((New-RSFixtureAction -ActionId 'snapshot-end-mutate-probe' -Arguments (
                        [ordered]@{ mutate = 'resize-and-write'; retainBorrow = $false })))
        }
        'teardown-callback-quiet' {
            $actions.Add((New-RSFixtureAction -ActionId 'teardown-sequence' -Arguments (
                        [ordered]@{ order = 'stop-producers,cancel,stop-posting,detach-callback,release,unload' })))
        }
        'title-cwd-boundaries' {
            $actions.Add((New-RSFixtureAction -ActionId 'probe-title-cwd-limits' -Arguments (
                        [ordered]@{ cwd = '131072'; title = '4096' })))
        }
        'write-response-order' {
            $actions.Add((New-RSFixtureAction -ActionId 'admit-external-events' -Arguments (
                        [ordered]@{ order = 'input,resize,side-channel' })))
        }
    }
    return $actions.ToArray()
}

function Get-RSInputIdentity {
    param([Parameter(Mandatory)][object[]]$Plan)

    $hash = [Security.Cryptography.IncrementalHash]::CreateHash(
        [Security.Cryptography.HashAlgorithmName]::SHA256)
    [uint64]$size = 0
    $repeatBuffer = [byte[]]::new(65536)
    try {
        foreach ($step in $Plan) {
            if ([string]$step.op -ceq 'LiteralHex') {
                $bytes = [Convert]::FromHexString([string]$step.hex)
                $hash.AppendData($bytes)
                $size += [uint64]$bytes.LongLength
                continue
            }
            if ([string]$step.op -cne 'RepeatByte') {
                throw "Unsupported input op '$($step.op)'."
            }
            [Array]::Fill($repeatBuffer, [byte]$step.byteValue)
            [uint64]$remaining = [uint64]$step.repeatCount
            while ($remaining -ne 0) {
                $count = [int][Math]::Min([uint64]$repeatBuffer.Length, $remaining)
                $hash.AppendData($repeatBuffer, 0, $count)
                $size += [uint64]$count
                $remaining -= [uint64]$count
            }
        }
        return [pscustomobject]@{
            Size = $size
            Sha256 = [Convert]::ToHexString($hash.GetHashAndReset()).ToLowerInvariant()
        }
    }
    finally {
        $hash.Dispose()
    }
}

function Expand-RSInputPlan {
    param(
        [Parameter(Mandatory)][object[]]$Plan,
        [Parameter(Mandatory)][ValidateRange(0, 4194304)][uint64]$MaximumBytes
    )

    $identity = Get-RSInputIdentity -Plan $Plan
    if ($identity.Size -gt $MaximumBytes) {
        throw "Input of $($identity.Size) bytes exceeds terminator-cut expansion bound."
    }
    $stream = [IO.MemoryStream]::new([int]$identity.Size)
    try {
        foreach ($step in $Plan) {
            if ([string]$step.op -ceq 'LiteralHex') {
                $bytes = [Convert]::FromHexString([string]$step.hex)
                $stream.Write($bytes)
            }
            else {
                $buffer = [byte[]]::new(65536)
                [Array]::Fill($buffer, [byte]$step.byteValue)
                [uint64]$remaining = [uint64]$step.repeatCount
                while ($remaining -ne 0) {
                    $count = [int][Math]::Min([uint64]$buffer.Length, $remaining)
                    $stream.Write($buffer, 0, $count)
                    $remaining -= [uint64]$count
                }
            }
        }
        return $stream.ToArray()
    }
    finally {
        $stream.Dispose()
    }
}

function Get-RSTerminatorCuts {
    param([Parameter(Mandatory)][byte[]]$Bytes)

    $cuts = [Collections.Generic.SortedSet[uint64]]::new()
    $inCsi = $false
    for ($index = 1; $index -lt $Bytes.Length; ++$index) {
        $current = $Bytes[$index]
        $previous = $Bytes[$index - 1]
        if ($previous -eq 0x1b -and $current -eq 0x5b) {
            $inCsi = $true
        }
        elseif ($inCsi -and $current -ge 0x40 -and $current -le 0x7e) {
            [void]$cuts.Add([uint64]$index)
            if (($index + 1) -lt $Bytes.Length) {
                [void]$cuts.Add([uint64]($index + 1))
            }
            $inCsi = $false
        }
        if ($current -eq 0x07 -or
            $current -eq 0x1b -or
            $previous -eq 0x1b -or
            ($current -band 0xc0) -eq 0x80) {
            [void]$cuts.Add([uint64]$index)
        }
    }
    return [uint64[]]@($cuts)
}

function ConvertTo-RSCppString {
    param([Parameter(Mandatory)][string]$Value)
    return $Value.Replace('\', '\\').Replace('"', '\"')
}

if (-not [IO.File]::Exists($Corpus)) {
    throw "Fixture corpus is missing: $Corpus"
}
$corpusBytes = [IO.File]::ReadAllBytes($Corpus)
$corpusText = [Text.UTF8Encoding]::new($false, $true).GetString($corpusBytes)
$corpusObject = $corpusText |
    ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
if ([string]$corpusObject.formatId -cne
    'red-salamander-terminal-engine-fixture-corpus' -or
    [int]$corpusObject.version -ne 1) {
    throw 'The fixture corpus format/version is unsupported.'
}

$scheduleAlgorithms = @{
    'all-at-once' = 'AllAtOnce'
    'bytewise' = 'Bytewise'
    'fibonacci' = 'Fibonacci'
    'terminator-split' = 'TerminatorCuts'
}
$fixtureRecords = [Collections.Generic.List[object]]::new()
foreach ($fixture in @($corpusObject.fixtures)) {
    $fixtureId = [string]$fixture.fixtureId
    $plan = [object[]](Get-RSFixtureInputPlan -FixtureId $fixtureId)
    $identity = Get-RSInputIdentity -Plan $plan
    $checks = [object[]]@(
        Get-RSExpectedCheck `
            -FixtureId $fixtureId `
            -Expectation ([string]$fixture.expectation))
    $actions = [object[]](Get-RSFixtureActions -FixtureId $fixtureId)
    $schedules = [Collections.Generic.List[object]]::new()
    foreach ($scheduleId in [string[]]$fixture.schedules) {
        if (-not $scheduleAlgorithms.ContainsKey($scheduleId)) {
            throw "Fixture '$fixtureId' names unknown schedule '$scheduleId'."
        }
        $cuts = [uint64[]]@()
        if ($scheduleId -ceq 'terminator-split') {
            $cuts = Get-RSTerminatorCuts -Bytes (
                Expand-RSInputPlan -Plan $plan -MaximumBytes 4194304)
        }
        $scheduleIdentity = [ordered]@{
            algorithm = [string]$scheduleAlgorithms[$scheduleId]
            cutOffsets = [object[]]@($cuts | ForEach-Object {
                    $_.ToString([Globalization.CultureInfo]::InvariantCulture)
                })
            inputSizeBytes = $identity.Size.ToString(
                [Globalization.CultureInfo]::InvariantCulture)
            scheduleId = $scheduleId
        }
        $schedules.Add([pscustomobject]@{
                ScheduleId = $scheduleId
                Algorithm = [string]$scheduleAlgorithms[$scheduleId]
                Cuts = $cuts
                Sha256 = Get-RSJcsSha256 -InputObject $scheduleIdentity
            })
    }
    $sortedSchedules = [object[]]@(
        $schedules | Sort-Object -Property ScheduleId -CaseSensitive)
    $sortedChecks = [object[]]@(
        $checks | Sort-Object -Property checkId -CaseSensitive)
    $contract = [ordered]@{
        actions = [object[]]@($actions | ForEach-Object {
                [ordered]@{
                    actionId = [string]$_.ActionId
                    actionSha256 = [string]$_.Sha256
                }
            })
        checks = [object[]]@($sortedChecks | ForEach-Object {
                [ordered]@{
                    checkId = [string]$_.checkId
                    expectedSha256 = [string]$_.expectedSha256
                }
            })
        fixtureId = $fixtureId
        inputSha256 = [string]$identity.Sha256
        inputSizeBytes = $identity.Size.ToString(
            [Globalization.CultureInfo]::InvariantCulture)
        schedules = [object[]]@($sortedSchedules | ForEach-Object {
                [ordered]@{
                    scheduleId = [string]$_.ScheduleId
                    scheduleSha256 = [string]$_.Sha256
                }
            })
    }
    $fixtureRecords.Add([pscustomobject]@{
            FixtureId = $fixtureId
            Symbol = ($fixtureId -replace '[^A-Za-z0-9]', '_')
            Plan = $plan
            InputSize = [uint64]$identity.Size
            InputSha256 = [string]$identity.Sha256
            Schedules = $sortedSchedules
            Checks = $sortedChecks
            Actions = $actions
            ContractSha256 = Get-RSJcsSha256 -InputObject $contract
        })
}
$fixtureRecordsArray = [object[]]@(
    $fixtureRecords | Sort-Object -Property FixtureId -CaseSensitive)
$performanceFixtureIds = [string[]]@(
    'alternate-screen-roundtrip',
    'color-style-matrix',
    'synchronized-output',
    'unicode-grapheme-wide',
    'write-response-order')
$recordById = @{}
foreach ($record in $fixtureRecordsArray) {
    $recordById[[string]$record.FixtureId] = $record
}
$performanceHash = [Security.Cryptography.IncrementalHash]::CreateHash(
    [Security.Cryptography.HashAlgorithmName]::SHA256)
try {
    [uint64]$performanceRemaining = 8388608
    [uint64]$performanceCounter = 0
    while ($performanceRemaining -ne 0) {
        $fixtureId = $performanceFixtureIds[
            [int]($performanceCounter % [uint64]$performanceFixtureIds.Count)]
        $payload = Expand-RSInputPlan `
            -Plan ([object[]]$recordById[$fixtureId].Plan) `
            -MaximumBytes 4194304
        $suffix = [Text.Encoding]::ASCII.GetBytes(
            ('#{0:x16}' -f $performanceCounter) + "`n")
        [uint64]$recordSize = [uint64]$payload.Length + [uint64]$suffix.Length
        if ($recordSize -gt $performanceRemaining) {
            $filler = [byte[]]::new([int]$performanceRemaining)
            [Array]::Fill($filler, [byte]0x78)
            $performanceHash.AppendData($filler)
            $performanceRemaining = 0
            break
        }
        $performanceHash.AppendData($payload)
        $performanceHash.AppendData($suffix)
        $performanceRemaining -= $recordSize
        $performanceCounter++
    }
    $performancePayloadSha256 = [Convert]::ToHexString(
        $performanceHash.GetHashAndReset()).ToLowerInvariant()
}
finally {
    $performanceHash.Dispose()
}
$manifestSha256 = Get-RSJcsIdentitySetSha256 `
    -Records ([object[]]@($fixtureRecordsArray | ForEach-Object {
                [ordered]@{
                    fixtureId = [string]$_.FixtureId
                    contractSha256 = [string]$_.ContractSha256
                }
            })) `
    -KeyProperty fixtureId

$builder = [Text.StringBuilder]::new()
[void]$builder.AppendLine('#pragma once')
[void]$builder.AppendLine()
[void]$builder.AppendLine('#include <array>')
[void]$builder.AppendLine('#include <cstddef>')
[void]$builder.AppendLine('#include <cstdint>')
[void]$builder.AppendLine('#include <span>')
[void]$builder.AppendLine('#include <string_view>')
[void]$builder.AppendLine()
[void]$builder.AppendLine('namespace RedSalamander::TerminalEngine::Fixtures::V1 {')
[void]$builder.AppendLine('enum class InputOp : std::uint8_t { LiteralHex, RepeatByte };')
[void]$builder.AppendLine('struct InputStep final { InputOp op; std::string_view hex; std::uint8_t byteValue; std::uint64_t repeatCount; };')
[void]$builder.AppendLine('enum class ScheduleAlgorithm : std::uint8_t { AllAtOnce, Bytewise, Fibonacci, TerminatorCuts };')
[void]$builder.AppendLine('struct ScheduleSpec final { std::string_view scheduleId; ScheduleAlgorithm algorithm; std::span<const std::uint64_t> cutOffsets; std::string_view scheduleSha256; };')
[void]$builder.AppendLine('struct Check final { std::string_view checkId; std::string_view expectedValue; std::string_view expectedSha256; };')
[void]$builder.AppendLine('struct Action final { std::string_view actionId; std::string_view argumentsJcs; std::string_view actionSha256; };')
[void]$builder.AppendLine('struct Fixture final { std::string_view fixtureId; std::span<const InputStep> inputPlan; std::uint64_t inputSizeBytes; std::string_view inputSha256; std::span<const ScheduleSpec> schedules; std::span<const Action> actions; std::span<const Check> checks; std::string_view contractSha256; };')
[void]$builder.AppendLine('struct PerformancePayloadSpec final { std::span<const std::string_view> fixtureIds; std::uint64_t sizeBytes; std::string_view sha256; };')
[void]$builder.AppendLine()
foreach ($record in $fixtureRecordsArray) {
    $symbol = [string]$record.Symbol
    [void]$builder.AppendLine("inline constexpr std::array<InputStep, $($record.Plan.Count)> k_${symbol}_input{")
    foreach ($step in $record.Plan) {
        [void]$builder.AppendLine((
                '    InputStep{{InputOp::{0}, "{1}", {2}, {3}ULL}},' -f
                $step.op,
                $step.hex,
                $step.byteValue,
                $step.repeatCount))
    }
    [void]$builder.AppendLine('};')
    foreach ($schedule in $record.Schedules) {
        if ($schedule.Cuts.Count -eq 0) {
            continue
        }
        $cutValues = ($schedule.Cuts | ForEach-Object { "$_`ULL" }) -join ', '
        $cutSymbol = ([string]$schedule.ScheduleId -replace '[^A-Za-z0-9]', '_')
        [void]$builder.AppendLine(
            "inline constexpr std::array<std::uint64_t, $($schedule.Cuts.Count)> k_${symbol}_${cutSymbol}_cuts{$cutValues};")
    }
    [void]$builder.AppendLine(
        "inline constexpr std::array<ScheduleSpec, $($record.Schedules.Count)> k_${symbol}_schedules{")
    foreach ($schedule in $record.Schedules) {
        $cutsExpression = 'std::span<const std::uint64_t>{}'
        if ($schedule.Cuts.Count -ne 0) {
            $cutSymbol = ([string]$schedule.ScheduleId -replace '[^A-Za-z0-9]', '_')
            $cutsExpression = "k_${symbol}_${cutSymbol}_cuts"
        }
        [void]$builder.AppendLine((
                '    ScheduleSpec{{"{0}", ScheduleAlgorithm::{1}, {2}, "{3}"}},' -f
                $schedule.ScheduleId,
                $schedule.Algorithm,
                $cutsExpression,
                $schedule.Sha256))
    }
    [void]$builder.AppendLine('};')
    [void]$builder.AppendLine(
        "inline constexpr std::array<Check, $($record.Checks.Count)> k_${symbol}_checks{")
    foreach ($check in $record.Checks) {
        [void]$builder.AppendLine((
                '    Check{{"{0}", "{1}", "{2}"}},' -f
                (ConvertTo-RSCppString $check.checkId),
                (ConvertTo-RSCppString $check.expectedValue),
                $check.expectedSha256))
    }
    [void]$builder.AppendLine('};')
    [void]$builder.AppendLine(
        "inline constexpr std::array<Action, $($record.Actions.Count)> k_${symbol}_actions{")
    foreach ($action in $record.Actions) {
        [void]$builder.AppendLine((
                '    Action{{"{0}", "{1}", "{2}"}},' -f
                (ConvertTo-RSCppString $action.ActionId),
                (ConvertTo-RSCppString $action.ArgumentsJcs),
                $action.Sha256))
    }
    [void]$builder.AppendLine('};')
}
[void]$builder.AppendLine()
[void]$builder.AppendLine(
    "inline constexpr std::string_view kFixtureManifestSha256 = `"$manifestSha256`";")
[void]$builder.AppendLine(
    "inline constexpr std::array<std::string_view, $($performanceFixtureIds.Count)> kPerformanceFixtureIds{")
foreach ($fixtureId in $performanceFixtureIds) {
    [void]$builder.AppendLine("    `"$fixtureId`",")
}
[void]$builder.AppendLine('};')
[void]$builder.AppendLine(
    "inline constexpr PerformancePayloadSpec kPerformancePayload{kPerformanceFixtureIds, 8388608ULL, `"$performancePayloadSha256`"};")
[void]$builder.AppendLine(
    "inline constexpr std::array<Fixture, $($fixtureRecordsArray.Count)> kFixtures{")
foreach ($record in $fixtureRecordsArray) {
    $symbol = [string]$record.Symbol
    [void]$builder.AppendLine((
            '    Fixture{{"{0}", k_{1}_input, {2}ULL, "{3}", k_{1}_schedules, k_{1}_actions, k_{1}_checks, "{4}"}},' -f
            $record.FixtureId,
            $symbol,
            $record.InputSize,
            $record.InputSha256,
            $record.ContractSha256))
}
[void]$builder.AppendLine('};')
[void]$builder.AppendLine(
    '[[nodiscard]] inline constexpr std::span<const Fixture> GetFixtures() noexcept { return kFixtures; }')
[void]$builder.AppendLine('} // namespace RedSalamander::TerminalEngine::Fixtures::V1')

$generated = $builder.ToString().Replace("`r`n", "`n")
$generatedBytes = [Text.UTF8Encoding]::new($false).GetBytes($generated)
$generatedHeaderSha256 = Get-RSSha256Hex -Bytes $generatedBytes
$generatorSourceSha256 = (Get-FileHash `
        -LiteralPath $PSCommandPath `
        -Algorithm SHA256).Hash.ToLowerInvariant()
if (-not $Update -and (
        [string]$corpusObject.fixtureManifestSha256 -cne $manifestSha256 -or
        [string]$corpusObject.generatedHeaderSha256 -cne $generatedHeaderSha256 -or
        [string]$corpusObject.generatorSourceSha256 -cne $generatorSourceSha256)) {
    throw 'The corpus generator/header/fixture-manifest identities are not frozen to the current bytes.'
}
if ($Update) {
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Header)) | Out-Null
    [IO.File]::WriteAllBytes($Header, $generatedBytes)
}
else {
    $actualHeaderBytes = if ([IO.File]::Exists($Header)) {
        [IO.File]::ReadAllBytes($Header)
    }
    else {
        [byte[]]@()
    }
    if ([Convert]::ToBase64String($actualHeaderBytes) -cne
        [Convert]::ToBase64String($generatedBytes)) {
        throw 'The checked fixture header differs from the canonical generator output.'
    }
}

[pscustomobject]@{
    verified = $true
    fixtureCount = $fixtureRecordsArray.Count
    fixtureManifestSha256 = $manifestSha256
    corpusSha256 = Get-RSSha256Hex -Bytes $corpusBytes
    headerSha256 = $generatedHeaderSha256
}
