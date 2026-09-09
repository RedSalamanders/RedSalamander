Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\TerminalEvidence.psm1'
$schemaPath = Join-Path $repoRoot 'Specs\Terminal\TerminalEngineEvidence.schema.json'
$corpusRoot = Join-Path $repoRoot 'Specs\Terminal\TerminalEngineEvidenceCorpus'
. (Join-Path $repoRoot 'Tools\TestRunPlan.ps1')
Import-Module $modulePath -Force

function ConvertFrom-RSCorpusJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function ConvertTo-RSJcsString {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value
    )

    return [Text.Encoding]::UTF8.GetString((ConvertTo-RSJcsUtf8Bytes -InputObject $Value))
}

function Assert-RSOrdinalSequence {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Values,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $expected = [string[]]@($Values)
    [Array]::Sort($expected, [StringComparer]::Ordinal)
    if (($Values -join "`n") -cne ($expected -join "`n")) {
        throw "$Description is not in ascending ordinal order."
    }
    if (@($Values | Select-Object -Unique).Count -ne $Values.Count) {
        throw "$Description contains a duplicate ID."
    }
}

function Assert-RSThrows {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    $threw = $false
    try {
        & $Action | Out-Null
    } catch {
        $threw = $true
    }
    if (-not $threw) {
        throw 'Expected the action to throw a terminating exception.'
    }
}

function Get-RSSha256Hex {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $digest = $sha256.ComputeHash($Bytes)
    } finally {
        $sha256.Dispose()
    }
    return ([BitConverter]::ToString($digest) -replace '-', '').ToLowerInvariant()
}

Describe 'Terminal evidence public module contract' {
    It 'exports exactly the five frozen functions' {
        $actual = [string[]]@(
            Get-Command -Module TerminalEvidence -CommandType Function |
                Select-Object -ExpandProperty Name
        )
        [Array]::Sort($actual, [StringComparer]::Ordinal)
        $expected = [string[]]@(
            'ConvertTo-RSJcsUtf8Bytes',
            'Get-RSJcsIdentitySetSha256',
            'Get-RSJcsSha256',
            'Get-RSTestMachineHash',
            'Write-RSEvidenceRecord'
        )
        ($actual -join ',') | Should Be ($expected -join ',')
    }
}

Describe 'RFC 8785 JCS and SHA-256' {
    It 'matches the RFC number, literal, escaping, and property-order example' {
        $sampleString = [string]::Concat(
            [char]0x20ac,
            '$',
            [char]0x000f,
            [char]0x000a,
            "A'B",
            '"',
            '\',
            '"',
            '/')
        $value = [ordered]@{
            numbers = [object[]]@(333333333.33333329, 1E+30, 4.50, 2e-3, 1e-27)
            string = $sampleString
            literals = [object[]]@($null, $true, $false)
        }
        $expected = [string]::Concat(
            '{"literals":[null,true,false],"numbers":[333333333.3333333,1e+30,4.5,0.002,1e-27],"string":"',
            [char]0x20ac,
            '$\u000f\nA',
            "'",
            'B\"\\\"/"}')

        ConvertTo-RSJcsString -Value $value | Should Be $expected
    }

    It 'uses ECMAScript fixed and exponential number boundaries' {
        ConvertTo-RSJcsString -Value ([double]1e-7) | Should Be '1e-7'
        ConvertTo-RSJcsString -Value ([double]1e-6) | Should Be '0.000001'
        ConvertTo-RSJcsString -Value ([double]1e20) | Should Be '100000000000000000000'
        ConvertTo-RSJcsString -Value ([double]1e21) | Should Be '1e+21'
        ConvertTo-RSJcsString -Value ([double]-0.0) | Should Be '0'
    }

    It 'returns UTF-8 without a BOM or trailing newline' {
        $bytes = ConvertTo-RSJcsUtf8Bytes -InputObject ([ordered]@{ text = [string][char]0x00e9 })
        ($bytes -is [byte[]]) | Should Be $true
        $bytes[0] | Should Be 0x7b
        $bytes[$bytes.Length - 1] | Should Be 0x7d
        (($bytes.Length -ge 3) -and $bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf) |
            Should Be $false
    }

    It 'sorts object names with StringComparer Ordinal semantics' {
        $value = [Collections.Specialized.OrderedDictionary]::new([StringComparer]::Ordinal)
        $value.Add('a', 1)
        $value.Add('A', 2)
        $value.Add('aa', 3)
        ConvertTo-RSJcsString -Value $value | Should Be '{"A":2,"a":1,"aa":3}'
    }

    It 'hashes the exact canonical bytes' {
        $value = [ordered]@{ b = 1; a = [string][char]0x00e9 }
        $bytes = ConvertTo-RSJcsUtf8Bytes -InputObject $value
        $expected = Get-RSSha256Hex -Bytes $bytes
        Get-RSJcsSha256 -InputObject $value | Should Be $expected
    }

    It 'rejects non-finite, decimal, unsafe-integer, malformed-Unicode, and cyclic values' {
        Assert-RSThrows { ConvertTo-RSJcsUtf8Bytes -InputObject ([double]::NaN) -ErrorAction Stop }
        Assert-RSThrows { ConvertTo-RSJcsUtf8Bytes -InputObject ([double]::PositiveInfinity) -ErrorAction Stop }
        Assert-RSThrows { ConvertTo-RSJcsUtf8Bytes -InputObject ([decimal]1.25) -ErrorAction Stop }
        Assert-RSThrows { ConvertTo-RSJcsUtf8Bytes -InputObject ([long]9007199254740992) -ErrorAction Stop }
        Assert-RSThrows { ConvertTo-RSJcsUtf8Bytes -InputObject ([string][char]0xd800) -ErrorAction Stop }

        $cycle = [Collections.ArrayList]::new()
        [void]$cycle.Add($cycle)
        Assert-RSThrows { ConvertTo-RSJcsUtf8Bytes -InputObject $cycle -ErrorAction Stop }
    }
}

Describe 'JCS identity-set digest' {
    It 'sorts tuples field by field using ordinal comparison and retains full records' {
        $records = [object[]]@(
            [ordered]@{ candidateId = 'a'; pin = 'z'; scoredOrControl = 'scored'; retained = 'second' },
            [ordered]@{ candidateId = 'A'; pin = 'z'; scoredOrControl = 'scored'; retained = 'first' },
            [ordered]@{ candidateId = 'a'; pin = 'a'; scoredOrControl = 'control'; retained = 'middle' }
        )
        $expectedOrder = [object[]]@($records[1], $records[2], $records[0])
        $expected = Get-RSJcsSha256 -InputObject $expectedOrder

        Get-RSJcsIdentitySetSha256 `
            -Records $records `
            -KeyProperty candidateId, pin, scoredOrControl |
            Should Be $expected
    }

    It 'rejects missing, null, non-string, and duplicate tuple keys' {
        Assert-RSThrows {
            Get-RSJcsIdentitySetSha256 `
                -Records @([ordered]@{ candidateId = 'a' }) `
                -KeyProperty candidateId, pin `
                -ErrorAction Stop
        }
        Assert-RSThrows {
            Get-RSJcsIdentitySetSha256 `
                -Records @([ordered]@{ candidateId = 'a'; pin = $null }) `
                -KeyProperty candidateId, pin `
                -ErrorAction Stop
        }
        Assert-RSThrows {
            Get-RSJcsIdentitySetSha256 `
                -Records @([ordered]@{ candidateId = 'a'; pin = 1 }) `
                -KeyProperty candidateId, pin `
                -ErrorAction Stop
        }
        Assert-RSThrows {
            Get-RSJcsIdentitySetSha256 `
                -Records @(
                    [ordered]@{ candidateId = 'a'; pin = 'same'; retained = 1 },
                    [ordered]@{ candidateId = 'a'; pin = 'same'; retained = 2 }
                ) `
                -KeyProperty candidateId, pin `
                -ErrorAction Stop
        }
    }
}

Describe 'TerminalEngine evidence schema corpus' {
    It 'accepts every positive corpus record and rejects every negative corpus record' {
        $positiveFiles = @(Get-ChildItem -LiteralPath $corpusRoot -Filter 'positive.*.json' -File)
        $negativeFiles = @(Get-ChildItem -LiteralPath $corpusRoot -Filter 'negative.*.json' -File)
        $positiveFiles.Count | Should BeGreaterThan 0
        $negativeFiles.Count | Should BeGreaterThan 0

        foreach ($file in $positiveFiles) {
            if (-not (Test-Json -LiteralPath $file.FullName -SchemaFile $schemaPath -ErrorAction SilentlyContinue)) {
                throw "Positive schema corpus record was rejected: $($file.Name)"
            }
        }
        foreach ($file in $negativeFiles) {
            if (Test-Json -LiteralPath $file.FullName -SchemaFile $schemaPath -ErrorAction SilentlyContinue) {
                throw "Negative schema corpus record was accepted: $($file.Name)"
            }
        }
    }

    It 'keeps the checked-in ID arrays sorted and unique and the fixed sequences exact' {
        $collection = ConvertFrom-RSCorpusJson -Path (Join-Path $corpusRoot 'positive.collection.json')
        Assert-RSOrdinalSequence `
            -Values ([string[]]@($collection.candidates.candidateId)) `
            -Description 'collection candidates'

        $complete = @($collection.candidates | Where-Object outcome -eq 'complete')[0]
        Assert-RSOrdinalSequence `
            -Values ([string[]]@($complete.evidenceObjects.evidenceId)) `
            -Description 'complete evidence objects'
        Assert-RSOrdinalSequence `
            -Values ([string[]]@($complete.measurements.metricId)) `
            -Description 'complete measurements'
        ($complete.gates.gateId -join ',') | Should Be ((1..12 | ForEach-Object { "G$_" }) -join ',')

        $earlyFailureCandidate = @($collection.candidates | Where-Object outcome -eq 'early-failure')[0]
        $failedGate = [string]$earlyFailureCandidate.earlyFailure.firstFailedGate
        $failedIndex = [int]$failedGate.Substring(1) - 1
        for ($index = 0; $index -lt $earlyFailureCandidate.gates.Count; ++$index) {
            $gate = $earlyFailureCandidate.gates[$index]
            if ($index -lt $failedIndex) {
                $gate.status | Should Be 'pass'
                $gate.resultCategory | Should Be 'passed'
            } elseif ($index -eq $failedIndex) {
                $gate.status | Should Be 'fail'
                $gate.resultCategory | Should Be 'mandatory-failure'
            } else {
                $gate.status | Should Be 'not-run'
                $gate.resultCategory | Should Be ("failed-$failedGate")
                $null -eq $gate.evidenceSha256 | Should Be $true
            }
        }

        $finalization = ConvertFrom-RSCorpusJson -Path (Join-Path $corpusRoot 'positive.finalization.json')
        Assert-RSOrdinalSequence `
            -Values ([string[]]@($finalization.candidateResults.candidateId)) `
            -Description 'final candidate results'
        Assert-RSOrdinalSequence `
            -Values ([string[]]@($finalization.controlResults.candidateId)) `
            -Description 'final control results'
        $criterion = @($finalization.candidateResults | Where-Object outcome -eq 'scored-complete')[0].criteria
        ($criterion.criterionId -join ',') | Should Be ((1..11 | ForEach-Object { "C$_" }) -join ',')
        ($criterion.weight -join ',') | Should Be '18,14,12,10,10,8,8,7,6,4,3'
        ($finalization.inputManifests | ForEach-Object { "$($_.platform)|$($_.configuration)" }) -join ',' |
            Should Be 'ARM64|Release,x64|ASan Debug,x64|Release'
        Assert-RSOrdinalSequence `
            -Values ([string[]]@($finalization.selection.cohort)) `
            -Description 'selection cohort'
    }

    It 'rejects a mutation of the fixed gate order and uint64 upper bound' {
        $collection = ConvertFrom-RSCorpusJson -Path (Join-Path $corpusRoot 'positive.collection.json')
        $firstGate = $collection.candidates[0].gates[0]
        $collection.candidates[0].gates[0] = $collection.candidates[0].gates[1]
        $collection.candidates[0].gates[1] = $firstGate
        $json = ConvertTo-RSJcsString -Value $collection
        (Test-Json -Json $json -SchemaFile $schemaPath -ErrorAction SilentlyContinue) | Should Be $false

        $collection = ConvertFrom-RSCorpusJson -Path (Join-Path $corpusRoot 'positive.collection.json')
        $collection.candidates[0].measurements[0].samples[2] = '18446744073709551616'
        $json = ConvertTo-RSJcsString -Value $collection
        (Test-Json -Json $json -SchemaFile $schemaPath -ErrorAction SilentlyContinue) | Should Be $false
    }
}

Describe 'Transactional evidence archive writer' {
    It 'writes one canonical manifest and one lowercase digest companion using create-new semantics' {
        $machineHash = Get-RSTestMachineHash
        $runId = '2026-07-30_160000'
        $scratch = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'terminal-evidence-pester' `
            -Case ([guid]::NewGuid().ToString('N'))
        $directory = Join-Path $scratch "$machineHash\TerminalEngine\$runId"
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
        $manifest = [ordered]@{
            runId = $runId
            machineHash = $machineHash
            z = 2
            a = [string][char]0x00e9
        }

        $result = Write-RSEvidenceRecord `
            -Manifest $manifest `
            -Directory $directory `
            -JsonLeaf 'terminal-engine-evidence.v1.json' `
            -CompanionLeaf 'terminal-engine-evidence.v1.sha256'
        $expectedPath = [IO.Path]::GetFullPath((Join-Path $directory 'terminal-engine-evidence.v1.json'))
        $result | Should Be $expectedPath

        $files = @(Get-ChildItem -LiteralPath $directory -File)
        $files.Count | Should Be 2
        $jsonBytes = [IO.File]::ReadAllBytes($expectedPath)
        [Text.Encoding]::UTF8.GetString($jsonBytes) | Should Be (ConvertTo-RSJcsString -Value $manifest)
        $jsonBytes[$jsonBytes.Length - 1] | Should Not Be 0x0a

        $expectedDigest = Get-RSSha256Hex -Bytes $jsonBytes
        $companionBytes = [IO.File]::ReadAllBytes((Join-Path $directory 'terminal-engine-evidence.v1.sha256'))
        [Text.Encoding]::ASCII.GetString($companionBytes) | Should Be ($expectedDigest + "`n")
        $companionBytes.Length | Should Be 65

        Assert-RSThrows {
            Write-RSEvidenceRecord `
                -Manifest $manifest `
                -Directory $directory `
                -JsonLeaf 'terminal-engine-evidence.v1.json' `
                -CompanionLeaf 'terminal-engine-evidence.v1.sha256' `
                -ErrorAction Stop
        }
        @(Get-ChildItem -LiteralPath $directory -File).Count | Should Be 2
    }

    It 'rejects run, profile, actual-machine, and leaf mismatches before publishing' {
        $machineHash = Get-RSTestMachineHash
        $runId = '2026-07-30_160001'
        $scratch = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'terminal-evidence-pester' `
            -Case ([guid]::NewGuid().ToString('N'))
        $directory = Join-Path $scratch "$machineHash\TerminalEngine\$runId"
        New-Item -ItemType Directory -Path $directory -Force | Out-Null

        Assert-RSThrows {
            Write-RSEvidenceRecord `
                -Manifest ([ordered]@{ runId = '2026-07-30_160002'; machineHash = $machineHash }) `
                -Directory $directory `
                -JsonLeaf 'record.json' `
                -CompanionLeaf 'record.sha256' `
                -ErrorAction Stop
        }
        Assert-RSThrows {
            Write-RSEvidenceRecord `
                -Manifest ([ordered]@{ runId = $runId; machineHash = '000000000000' }) `
                -Directory $directory `
                -JsonLeaf 'record.json' `
                -CompanionLeaf 'record.sha256' `
                -ErrorAction Stop
        }
        Assert-RSThrows {
            Write-RSEvidenceRecord `
                -Manifest ([ordered]@{ runId = $runId; machineHash = $machineHash }) `
                -Directory $directory `
                -JsonLeaf '..\record.json' `
                -CompanionLeaf 'record.sha256' `
                -ErrorAction Stop
        }
        @(Get-ChildItem -LiteralPath $directory -File).Count | Should Be 0
    }
}

Describe 'Test machine hash compatibility' {
    It 'matches the SelfTest MachineGuid/computer-name UTF-8 SHA-256 algorithm' {
        $seed = $null
        try {
            $seed = [string](Get-ItemPropertyValue `
                    -LiteralPath 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography' `
                    -Name MachineGuid `
                    -ErrorAction Stop)
        } catch {
            $seed = [Environment]::MachineName
        }
        if ([string]::IsNullOrEmpty($seed)) {
            throw 'Test precondition failed: neither MachineGuid nor computer name is available.'
        }
        $seedBytes = [Text.UTF8Encoding]::new($false, $true).GetBytes($seed)
        $expected = (Get-RSSha256Hex -Bytes $seedBytes).Substring(0, 12)

        $actual = Get-RSTestMachineHash
        $actual | Should Match '^[0-9a-f]{12}$'
        $actual | Should Be $expected
    }
}
