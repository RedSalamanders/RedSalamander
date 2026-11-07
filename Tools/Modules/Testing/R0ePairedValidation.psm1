Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'TestSandbox.psm1') -Scope Local

# These exact case bodies and shared measured/predicate/remainder projections were
# independently reviewed. A different body requires a new review, not a new hash here.
function Get-RSR0eCaseIdentity {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SourceText)
    $text = $SourceText.Replace("`r`n", "`n")
    $match = [regex]::Match($text, '(?ms)^case SelfTestState::Step::FileOps_ProviderCapabilityMatrix:\s*\n.*?(?=^case SelfTestState::Step::)')
    if (-not $match.Success) { throw 'R0e case body not found.' }
    $case = $match.Value.TrimEnd() + "`n"
    $hash = { param([string]$value) [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($value))) }
    $full = & $hash $case
    if ($full -cnotin @('AB2DE533491C1816F373655E164C89A38ED38D3EC3D6DE6B8F90502BBB3F9517','A3541BB88E72E6BC41EAD4C77D37F9136A39EF123C8AAB780ECEF63039F020CE')) { throw 'Unreviewed R0e case body.' }
    $start = 'providerOperationCallCount == 0u && r0eElapsedUs < kR0eSelfTestCeilingUs,'
    $end = "`n        {`n            return true;`n        }`n`n        std::string missingCreateDirectoryJson(localCapabilitiesJsonSnapshot);"
    $i = $case.IndexOf($start, [StringComparison]::Ordinal)
    if ($i -lt 0 -or $case.IndexOf($start, $i + 1, [StringComparison]::Ordinal) -ge 0) { throw 'R0e predicate anchor missing or repeated.' }
    $j = $case.IndexOf($end, $i + $start.Length, [StringComparison]::Ordinal)
    if ($j -lt 0 -or $case.IndexOf($end, $j + 1, [StringComparison]::Ordinal) -ge 0) { throw 'R0e suffix anchor missing or repeated.' }
    $prefix = & $hash ($case.Substring(0, $i + $start.Length))
    $suffix = & $hash ($case.Substring($j))
    if ($prefix -cne 'B061625FED5FFB3ADF57BCE2B0FA02E57C055B4E0EEE034628F0719852611833' -or
        $suffix -cne 'BAE1AC8644A3D3994F4DFCB39E763F67859562A095C13903EC3AD71C73582C6B') { throw 'R0e measured case projection changed.' }
    return [pscustomobject]@{ full=$full; prefix=$prefix; suffix=$suffix }
}

function Test-RSR0eInteger {
    param($Value)
    return ($null -ne $Value -and $Value -isnot [bool] -and
        ($Value -is [int] -or $Value -is [long] -or $Value -is [uint32] -or $Value -is [uint64]) -and
        $Value -ge 0 -and $Value -le [long]::MaxValue)
}

function Test-RSR0eUnitScale {
    param($Value)
    return ($null -ne $Value -and $Value -isnot [bool] -and
        ($Value -is [int] -or $Value -is [long] -or $Value -is [uint32] -or $Value -is [uint64] -or
         $Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) -and $Value -eq 1)
}

function Test-RSR0eSample {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Summary,
          [Parameter(Mandatory)][System.Collections.IDictionary]$Native,
          [Parameter(Mandatory)][object[]]$Metrics,
          [Parameter(Mandatory)][ValidateSet('baseline','candidate')][string]$Side,
          [Parameter(Mandatory)][string]$RunId,
          [Parameter(Mandatory)][ValidateSet('x64','ARM64')][string]$Platform,
          [Parameter(Mandatory)][int]$RunnerExit)
    if ($RunId -cnotmatch '^\d{8}T\d{6}Z-\d+-[0-9a-f]{32}$' -or $Summary.schema -cne 'red-salamander.run-all-tests.v1' -or
        $Summary.run_id -cne $RunId -or $Summary.suite -cne 'FileOps' -or $Summary.case_filter -cne 'FileOps_ProviderCapabilityMatrix' -or
        $Summary.platform -cne $Platform -or $Summary.configuration -cne 'Release' -or $Summary.timeout_scale -ne 1 -or $Summary.fail_fast -ne $false -or
        $Native.suite -cne 'FileOperations' -or $Native.timeout_scale -ne 1 -or $Native.repeat_count -ne 1 -or
        $Native.fail_fast -isnot [bool] -or $Native.fail_fast) { throw 'R0e invocation/profile mismatch.' }
    if (-not (Test-RSR0eUnitScale $Summary.timeout_scale) -or -not (Test-RSR0eUnitScale $Native.timeout_scale)) {
        throw 'Invalid R0e timeout scale.'
    }
    foreach ($field in @('exit_code','passed','failed','skipped','total')) {
        if (-not (Test-RSR0eInteger $Summary[$field])) { throw "Invalid runner count: $field" }
    }
    foreach ($field in @('repeat_count','passed','failed','skipped')) {
        if (-not (Test-RSR0eInteger $Native[$field])) { throw "Invalid native count: $field" }
    }
    if ($Summary.fail_fast -isnot [bool]) { throw 'Invalid runner fail_fast.' }
    $q = $Summary.quarantine
    foreach ($field in @('active_count','invalid_count','repair_attempt_count')) {
        if (-not (Test-RSR0eInteger $q[$field])) { throw "Invalid quarantine count: $field" }
    }
    if ($q -isnot [System.Collections.IDictionary] -or $q.is_valid -isnot [bool] -or -not $q.is_valid -or
        $q.has_blocking_entries -isnot [bool] -or $q.has_blocking_entries -or $q.active_count -ne 0 -or $q.invalid_count -ne 0 -or
        @($q.active_entries).Count -ne 0 -or @($q.invalid_entries).Count -ne 0 -or @($q.errors).Count -ne 0 -or $q.repair_attempt_count -ne 0) {
        throw 'R0e quarantine is not clean.'
    }
    $cases = @($Native.cases)
    foreach ($case in $cases) { if (-not (Test-RSR0eInteger $case.repeat_index) -or $case.repeat_index -ne 1) { throw 'Invalid R0e repeat index.' } }
    $names = @('Setup','FileOps_ProviderCapabilityMatrix','Cleanup_RestorePluginConfig')
    if ($cases.Count -ne 3 -or @($cases | Where-Object { $_.name -cnotin $names }).Count -ne 0) { throw 'Unexpected R0e cases.' }
    foreach ($name in $names) { if (@($cases | Where-Object { $_.name -ceq $name }).Count -ne 1) { throw 'Missing or duplicate R0e case.' } }
    foreach ($name in @('Setup','Cleanup_RestorePluginConfig')) {
        if (@($cases | Where-Object { $_.name -ceq $name -and $_.status -ceq 'passed' }).Count -ne 1) { throw 'R0e setup/cleanup failed.' }
    }
    $sample = @($cases | Where-Object { $_.name -ceq 'FileOps_ProviderCapabilityMatrix' })[0]
    $metricNames = @('RouteContainmentUs','CapabilityParseWorkUs','AdmissionWorkUs','RejectionChecksUs','CapabilityParseCount','AdmissionDecisionCount','RejectedUncontainedCount','ProviderCallCount')
    if ($Metrics.Count -ne 8) { throw 'Unexpected R0e metric inventory.' }
    $values = @{}; $identity = $null
    foreach ($name in $metricNames) {
        $rows = @($Metrics | Where-Object { $_.metric -ceq "FileOps.SelfTest.R0e.$name" })
        if ($rows.Count -ne 1) { throw "Missing or duplicate R0e metric: $name" }
        $row = $rows[0]
        foreach ($field in @('hr','durationUs','value0','value1','value','process','thread')) {
            if (-not (Test-RSR0eInteger $row.$field)) { throw "Invalid $field in R0e metric: $name" }
        }
        if ($row.hr -ne 0 -or $row.process -eq 0 -or $row.thread -eq 0) { throw "Invalid R0e metric identity or result: $name" }
        foreach ($field in @('scenario','build','branch','commit','machineHash','runId')) {
            if ($row.$field -isnot [string]) { throw "Missing R0e metadata: $field" }
        }
        if ($row.scenario -cne 'SelfTest' -or $row.build -cne 'Release' -or
            [string]::IsNullOrWhiteSpace($row.machineHash) -or [string]::IsNullOrWhiteSpace($row.runId)) { throw 'Invalid R0e metric metadata.' }
        if ($null -eq $identity) { $identity = $row } else {
            foreach ($field in @('process','thread','scenario','build','branch','commit','machineHash','runId')) {
                if ($row.$field -cne $identity.$field) { throw "R0e metric identity mismatch: $field" }
            }
        }
        if ($name -ceq 'RouteContainmentUs') {
            if ($row.unit -cne 'us' -or $row.detail -cne '4096-capability-admission-decisions' -or
                $row.value0 -ne 4096 -or $row.value1 -ne 4096 -or $row.value -ne $row.durationUs) { throw 'Invalid R0e route timing row.' }
            $values[$name]=[int64]$row.durationUs
        } else {
            if ($row.unit -cne 'count' -or $row.detail -cne 'value' -or $row.durationUs -ne 0 -or
                $row.value1 -ne 0 -or $row.value -ne $row.value0) { throw "Invalid R0e value row: $name" }
            $values[$name]=[int64]$row.value0
        }
    }
    if ($values.RouteContainmentUs -le 0 -or $values.CapabilityParseWorkUs -gt $values.RouteContainmentUs -or
        $values.AdmissionWorkUs -gt $values.RouteContainmentUs -or $values.CapabilityParseCount -ne 4096 -or
        $values.AdmissionDecisionCount -ne 4096 -or $values.RejectedUncontainedCount -ne 1 -or $values.ProviderCallCount -ne 0) { throw 'R0e metric invariants failed.' }
    $red = $sample.status -ceq 'failed' -and $values.RouteContainmentUs -ge 2000000
    if ($red) {
        if ($Side -cne 'baseline' -or $RunnerExit -ne 1 -or $Summary.exit_code -ne 1 -or $Summary.passed -ne 2 -or $Summary.failed -ne 1 -or
            $Summary.skipped -ne 0 -or $Native.passed -ne 2 -or $Native.failed -ne 1 -or $Native.skipped -ne 0 -or
            $sample.reason -cnotmatch '^R0e capability/admission containment must remain complete, provider-call-free, and within the build-specific ceiling(?:\.| \(parses=4096/4096, decisions=4096, rejected=1, providerCalls=0, elapsed=\d+ us, ceiling=2000000 us, parseWork=\d+ us, admissionWork=\d+ us, rejectionChecks=\d+ us\)\.)$') {
            throw 'Baseline is not the sole reviewed timing-only failure.'
        }
        if ($sample.reason -cmatch 'elapsed=(\d+) us, ceiling=2000000 us, parseWork=(\d+) us, admissionWork=(\d+) us, rejectionChecks=(\d+) us') {
            if ([int64]$Matches[1] -ne $values.RouteContainmentUs -or [int64]$Matches[2] -ne $values.CapabilityParseWorkUs -or
                [int64]$Matches[3] -ne $values.AdmissionWorkUs -or [int64]$Matches[4] -ne $values.RejectionChecksUs) { throw 'RED diagnostic disagrees with raw metrics.' }
        }
    } elseif ($sample.status -cne 'passed' -or $RunnerExit -ne 2 -or $Summary.exit_code -ne 2 -or $Summary.passed -ne 3 -or $Summary.failed -ne 0 -or
        $Summary.skipped -ne 0 -or $Native.passed -ne 3 -or $Native.failed -ne 0 -or $Native.skipped -ne 0 -or $values.RouteContainmentUs -ge 2000000) {
        throw 'R0e selected case did not pass the unchanged ceiling.'
    }
    if ($Summary.total -ne 3) { throw 'R0e runner case count mismatch.' }
    return [pscustomobject]@{ side=$Side; run_id=$RunId; machine_hash=$identity.machineHash; metric_run_id=$identity.runId;
        process=$identity.process; thread=$identity.thread; branch=$identity.branch; commit=$identity.commit; red=$red; metrics=$values }
}

# This reducer consumes samples already checked against closed archives by the caller;
# it does not independently establish archive provenance.
function Get-RSR0eAssessment {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Baseline, [Parameter(Mandatory)][object[]]$Candidate)
    if ($Baseline.Count -ne 5 -or $Candidate.Count -ne 5) { throw 'R0e requires exactly five processes per side.' }
    $all = @($Baseline) + @($Candidate)
    foreach ($item in $all) {
        if ($item.red -isnot [bool] -or $item.metrics -isnot [System.Collections.IDictionary] -or
            $item.run_id -cnotmatch '^\d{8}T\d{6}Z-\d+-[0-9a-f]{32}$' -or
            -not (Test-RSR0eInteger $item.process) -or $item.process -eq 0 -or
            -not (Test-RSR0eInteger $item.thread) -or $item.thread -eq 0 -or
            [string]::IsNullOrWhiteSpace($item.metric_run_id)) { throw 'Invalid reduced sample identity.' }
        foreach ($field in @('RouteContainmentUs','CapabilityParseWorkUs','AdmissionWorkUs','RejectionChecksUs','CapabilityParseCount',
                             'AdmissionDecisionCount','RejectedUncontainedCount','ProviderCallCount')) {
            if (-not (Test-RSR0eInteger $item.metrics[$field])) { throw "Invalid reduced sample metric: $field" }
        }
        if ($item.metrics.RouteContainmentUs -le 0 -or $item.red -ne ($item.metrics.RouteContainmentUs -ge 2000000) -or
            $item.metrics.CapabilityParseCount -ne 4096 -or $item.metrics.AdmissionDecisionCount -ne 4096 -or
            $item.metrics.RejectedUncontainedCount -ne 1 -or $item.metrics.ProviderCallCount -ne 0) { throw 'Invalid reduced sample timing or counts.' }
    }
    if (@($all.run_id | Sort-Object -Unique).Count -ne 10 -or @($all.machine_hash | Sort-Object -Unique).Count -ne 1 -or
        [string]::IsNullOrWhiteSpace([string]$all[0].machine_hash) -or @($Baseline | Where-Object side -cne 'baseline').Count -ne 0 -or
        @($Candidate | Where-Object { $_.side -cne 'candidate' -or $_.red }).Count -ne 0) { throw 'R0e paired identities are inconsistent.' }
    $baseTimes=@($Baseline | ForEach-Object { [int64]$_.metrics.RouteContainmentUs } | Sort-Object)
    $candidateTimes=@($Candidate | ForEach-Object { [int64]$_.metrics.RouteContainmentUs } | Sort-Object)
    $threshold=[Math]::Max(1.5 * $baseTimes[4], $baseTimes[4] + 50000)
    if ($candidateTimes[4] -gt $threshold) { throw 'R0e matched p95 exceeds the unchanged ceiling.' }
    $red=@($Baseline | Where-Object red).Count -gt 0
    return [pscustomobject]@{ baseline_us=$baseTimes; candidate_us=$candidateTimes; baseline_p95_us=$baseTimes[4]; candidate_p95_us=$candidateTimes[4];
        threshold_us=$threshold; baseline_red=$red; comparison_passed=$true; outcome='NOT_EVALUATED' }
}

function Read-RSR0eArchive {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$TestRoot, [Parameter(Mandatory)][string]$RunId)
    if ($RunId -cnotmatch '^\d{8}T\d{6}Z-\d+-[0-9a-f]{32}$') { throw 'Invalid R0e run ID.' }
    $run = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $TestRoot "evidence/completed-runs/$RunId") -RequireDescendant
    $records = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $TestRoot "evidence/run-archive-records/$RunId") -RequireDescendant
    $intentPath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $records 'intent.json') -RequireDescendant
    $completePath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $records 'complete.json') -RequireDescendant
    $intent = Get-Content -LiteralPath $intentPath -Raw | ConvertFrom-Json -AsHashtable
    $complete = Get-Content -LiteralPath $completePath -Raw | ConvertFrom-Json -AsHashtable
    $source = [IO.Path]::GetFullPath((Join-Path $TestRoot "runs/$RunId"))
    if ($intent.schema -cne 'red-salamander.run-archive.v1' -or $intent.run_id -cne $RunId -or $intent.disposition -cne 'completed' -or
        $intent.source -ine $source -or $intent.destination -ine $run -or $complete.schema -cne 'red-salamander.run-archive-complete.v1' -or
        $complete.source -ine $source -or $complete.destination -ine $run -or $complete.intent_sha256 -ine (Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash) {
        throw 'R0e archive closure is invalid.'
    }
    $files = @{}
    foreach ($relative in @('artifacts\selftest\last_run\fileops\results.json',
                            'artifacts\selftest\last_run\perf\perf_metrics.jsonl','artifacts\selftest\last_run\fileops\trace.txt')) {
        $path = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $run $relative) -RequireDescendant
        $rows = @($intent.inventory | Where-Object { $_.path -ieq $relative -and -not $_.directory })
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "R0e archived file missing: $relative" }
        $sha = (Get-FileHash -LiteralPath $path -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($rows.Count -ne 1 -or [int64]$rows[0].length -ne (Get-Item -LiteralPath $path).Length -or $rows[0].sha256 -ine $sha) {
            throw "R0e archived bytes do not match inventory: $relative"
        }
        $files[$relative] = @{path=$path;sha256=$sha}
    }
    # The runner writes its summary after archive closure, so it is observed separately;
    # the command binds its parsed identity/profile/times to the invocation and log.
    $summaryRelative = 'artifacts\selftest\last_run\run-all-tests-results.json'
    $summaryPath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $run $summaryRelative) -RequireDescendant
    $summaryBytes = [IO.File]::ReadAllBytes($summaryPath)
    $summaryText = [Text.Encoding]::UTF8.GetString($summaryBytes).TrimStart([char]0xFEFF)
    $summaryHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($summaryBytes))
    $files[$summaryRelative] = @{path=$summaryPath;sha256=$summaryHash;length=$summaryBytes.Length;provenance='post-archive-observed'}
    return @{summary_text=$summaryText; summary_bytes=$summaryBytes;
        native=(Get-Content -LiteralPath $files['artifacts\selftest\last_run\fileops\results.json'].path -Raw | ConvertFrom-Json -AsHashtable);
        metrics=@(Get-Content -LiteralPath $files['artifacts\selftest\last_run\perf\perf_metrics.jsonl'].path | ForEach-Object { $_ | ConvertFrom-Json } |
            Where-Object { $_.metric -clike 'FileOps.SelfTest.R0e.*' });files=$files;intent_sha256=$complete.intent_sha256}
}

Export-ModuleMember -Function Get-RSR0eCaseIdentity, Test-RSR0eSample, Get-RSR0eAssessment, Read-RSR0eArchive
