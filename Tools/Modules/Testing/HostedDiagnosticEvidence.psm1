Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'TestSuitePlan.psm1') -Scope Local -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'TestRunPlan.Common.psm1') -Scope Local -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'ValidationEvidence.psm1') -Scope Local
Import-Module (Join-Path $PSScriptRoot 'TestSandbox.psm1') -Scope Local
Import-Module (Join-Path $PSScriptRoot '../Build/BuildEvidence.psm1') -Scope Local -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'ValidationFingerprint.psm1') -Scope Local -ErrorAction Stop

function Get-RSHostedPairedDispatch {
    param([Parameter(Mandatory)][string]$EventName, [Parameter(Mandatory)][string]$EventJson,
          [Parameter(Mandatory)][string]$CandidateHead)
    if ($EventName -cne 'workflow_dispatch' -or $CandidateHead -cnotmatch '^[a-f0-9]{40}$' -or $CandidateHead -cmatch '^0+$') {
        throw 'R0e paired validation requires a manual dispatch and a nonzero immutable candidate SHA.'
    }
    try { $event = $EventJson | ConvertFrom-Json -AsHashtable -ErrorAction Stop }
    catch { throw 'R0e paired dispatch event is invalid JSON.' }
    if ($event -isnot [System.Collections.IDictionary] -or -not $event.Contains('inputs') -or
        $event.inputs -isnot [System.Collections.IDictionary] -or -not $event.inputs.Contains('scope') -or
        -not $event.inputs.Contains('baseline_ref') -or -not $event.inputs.Contains('configuration') -or
        $event.inputs.scope -cne 'R0ePaired' -or $event.inputs.configuration -cne 'Release' -or $event.inputs.baseline_ref -isnot [string] -or
        $event.inputs.baseline_ref -cnotmatch '^[a-f0-9]{40}$' -or $event.inputs.baseline_ref -cmatch '^0+$') {
        throw 'R0e paired dispatch requires its exact scope and a nonzero immutable baseline SHA.'
    }
    return @{ baseline=$event.inputs.baseline_ref; candidate=$CandidateHead }
}

# This adapter-only check deliberately accepts filtered runner exit 2, unlike the
# general terminal reducer, but only after that reducer validates every native case.
# It never converts the repository's NOT_EVALUATED coverage into a Full result.
function Get-RSHostedDiagnosticSelection {
    param([Parameter(Mandatory)][ValidateSet('CommandsHandoff','FileOpsVolumes','ViewerSpaceCapture')][string]$Scope)
    switch ($Scope) {
        'CommandsHandoff' { return @{ suite='Commands'; case='cmd_pane_selection_select_next_keeps_navigation_shell_stable'; native='commands'; nativeSuite='Commands'; entry='selftest.commands' } }
        'FileOpsVolumes' { return @{ suite='FileOps'; case='R4A19_DiscoveryIndependentVolumes'; native='fileops'; nativeSuite='FileOperations'; entry='selftest.file-operations' } }
        'ViewerSpaceCapture' { return @{ suite='Commands'; case='cmd_viewer_space_explorer_analyst_exact_nested_detail'; native='commands'; nativeSuite='Commands'; entry='selftest.commands' } }
    }
}

function Get-RSHostedDiagnosticRunId {
    param([Parameter(Mandatory)][string]$Output)
    $matches = [regex]::Matches($Output, '(?m)^\s*Run id:\s*([0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32})\s*$')
    if ($matches.Count -ne 1) { throw 'Expected one exact runner-owned diagnostic run ID.' }
    return $matches[0].Groups[1].Value
}

function Assert-RSHostedDiagnosticSelection {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Selection)
    $known = @((Get-RSHostedDiagnosticSelection CommandsHandoff), (Get-RSHostedDiagnosticSelection FileOpsVolumes), (Get-RSHostedDiagnosticSelection ViewerSpaceCapture))
    $matches = @($known | Where-Object {
        $_.suite -ceq $Selection.suite -and $_.case -ceq $Selection.case -and $_.native -ceq $Selection.native -and
        $_.nativeSuite -ceq $Selection.nativeSuite -and $_.entry -ceq $Selection.entry
    })
    if ($matches.Count -ne 1) { throw 'Unknown hosted diagnostic selection.' }
}

function Test-RSHostedTimeWindow {
    param(
        [Parameter(Mandatory)][object]$Summary,
        [Parameter(Mandatory)][DateTimeOffset]$InvocationStartedUtc,
        [Parameter(Mandatory)][DateTimeOffset]$InvocationEndedUtc
    )

    if ($InvocationStartedUtc -gt $InvocationEndedUtc) {
        throw 'Invocation window bounds are inverted.'
    }
    $times = @()
    foreach ($value in @($Summary.run_started_utc, $Summary.run_ended_utc)) {
        $parsed = [DateTimeOffset]::MinValue
        if ($value -is [datetime]) {
            if ($value.Kind -ne [DateTimeKind]::Utc) { throw 'Diagnostic summary time is invalid.' }
            $parsed = [DateTimeOffset]$value
        } elseif ($value -is [datetimeoffset]) {
            $parsed = $value
        } elseif ($value -isnot [string] -or $value -cnotmatch '(Z|[+-]\d{2}:\d{2})$' -or
            -not [DateTimeOffset]::TryParse($value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            throw 'Diagnostic summary time is invalid.'
        }
        $times += $parsed.ToUniversalTime()
    }
    if ($times.Count -ne 2 -or $times[0] -lt $InvocationStartedUtc -or $times[1] -gt $InvocationEndedUtc -or $times[0] -gt $times[1]) {
        throw 'Diagnostic summary falls outside the invocation window.'
    }
}

function Assert-RSHostedCleanQuarantine {
    param([Parameter(Mandatory)][object]$Summary)

    $q = $Summary.quarantine
    if ($q -isnot [System.Collections.IDictionary] -or -not $q.Contains('is_valid') -or -not $q.Contains('has_blocking_entries') -or
        -not $q.Contains('active_count') -or -not $q.Contains('active_entries') -or -not $q.Contains('invalid_count') -or
        -not $q.Contains('invalid_entries') -or -not $q.Contains('errors') -or $q.is_valid -isnot [bool] -or
        $q.has_blocking_entries -isnot [bool] -or $q.is_valid -cne $true -or
        $q.has_blocking_entries -cne $false -or [int]$q.active_count -ne 0 -or [int]$q.invalid_count -ne 0 -or
        @($q.active_entries).Count -ne 0 -or @($q.invalid_entries).Count -ne 0 -or @($q.errors).Count -ne 0) {
        throw 'Missing, invalid, or blocking quarantine state.'
    }
}

function Test-RSHostedDiagnosticOutcome {
    param([Parameter(Mandatory)][int]$ExitCode, [Parameter(Mandatory)][string]$RunId,
          [Parameter(Mandatory)][System.Collections.IDictionary]$Selection,
          [Parameter(Mandatory)][System.Collections.IDictionary]$Summary,
          [Parameter(Mandatory)][System.Collections.IDictionary]$Native,
          [Parameter(Mandatory)][string]$Platform, [Parameter(Mandatory)][string]$Configuration,
          [Parameter(Mandatory)][bool]$ReceiptVerified,
          [Parameter(Mandatory)][string]$ExpectedExePath,
          [Parameter(Mandatory)][DateTimeOffset]$InvocationStartedUtc, [Parameter(Mandatory)][DateTimeOffset]$InvocationEndedUtc)
    Assert-RSHostedDiagnosticSelection $Selection
    if ($RunId -cnotmatch '^[0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32}$') { throw 'Invalid diagnostic run ID.' }
    if (-not $ReceiptVerified -or $ExitCode -ne 2 -or $Summary.schema -cne 'red-salamander.run-all-tests.v1' -or
        $Summary.run_id -cne $RunId -or $Summary.suite -cne $Selection.suite -or $Summary.case_filter -cne $Selection.case -or
        $Summary.platform -cne $Platform -or $Summary.configuration -cne $Configuration -or [int]$Summary.exit_code -ne 2) {
        throw 'Filtered runner, profile, or source receipt does not match the diagnostic request.'
    }
    # V1 has no receipt ID; bind the summary to this child invocation and executable.
    if ($Summary.executable -isnot [string] -or $Summary.executable -ine $ExpectedExePath) {
        throw 'Diagnostic executable or invocation window does not match.'
    }
    Test-RSHostedTimeWindow -Summary $Summary -InvocationStartedUtc $InvocationStartedUtc -InvocationEndedUtc $InvocationEndedUtc
    Assert-RSHostedCleanQuarantine -Summary $Summary
    if (-not $Native.Contains('suite') -or $Native.suite -cne $Selection.nativeSuite -or -not $Native.Contains('cases') -or
        -not $Native.Contains('passed') -or -not $Native.Contains('failed') -or -not $Native.Contains('skipped')) {
        throw 'Native diagnostic result is incomplete.'
    }
    $selected = @($Native.cases | Where-Object { $_.name -ceq $Selection.case })
    if ($selected.Count -ne 1 -or $selected[0].status -cne 'passed') { throw 'The exact selected native case did not pass once.' }
    $allowedCases = @($Selection.case)
    if ($Selection.suite -ceq 'FileOps') {
        foreach ($required in @('Setup', 'Cleanup_RestorePluginConfig')) {
            $case = @($Native.cases | Where-Object { $_.name -ceq $required })
            if ($case.Count -ne 1 -or $case[0].status -cne 'passed') { throw 'Required FileOps setup or cleanup did not pass.' }
            $allowedCases += $required
        }
    }
    if (@($Native.cases | Where-Object { $_.name -cnotin $allowedCases }).Count -ne 0) { throw 'Unexpected native cases in the filtered diagnostic.' }
    $result = @{ exit_code=0; passed=$Native.passed; failed=$Native.failed; skipped=$Native.skipped; cases=@($Native.cases) }
    $entry = @{ id=$Selection.entry; kind='SelfTest'; outcome_contract=@{ static_skip_cases=@(); capability_bound_skip_cases=@(); forbidden_skip_classes=@() } }
    $terminal = New-RSValidationTerminalResult -EntryContract $entry -Result $result -ResultParsed $true -CoverageValid $true -RequiredArtifactsPresent $true
    if ($terminal.status -cne 'PASSED' -or [int]$Native.skipped -ne 0 -or [int]$Summary.failed -ne 0 -or [int]$Summary.skipped -ne 0 -or
        [int]$Summary.passed -ne [int]$Native.passed -or [int]$Summary.total -ne @($Native.cases).Count) {
        throw 'Native diagnostic cases are incomplete, skipped, or failed.'
    }
    return $true
}

function Read-RSHostedVerifiedArchiveRecord {
    param(
        [Parameter(Mandatory)][string]$TestRoot,
        [Parameter(Mandatory)][string]$RunId
    )

    if ($RunId -cnotmatch '^[0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32}$') { throw 'Invalid diagnostic run ID.' }
    $run = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $TestRoot "evidence/completed-runs/$RunId") -RequireDescendant
    $records = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $TestRoot "evidence/run-archive-records/$RunId") -RequireDescendant
    $intentPath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $records 'intent.json') -RequireDescendant
    $completePath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $records 'complete.json') -RequireDescendant
    $expectedSource = [IO.Path]::GetFullPath((Join-Path $TestRoot "runs/$RunId"))
    $intent = Get-Content -LiteralPath $intentPath -Raw | ConvertFrom-Json -AsHashtable
    $complete = Get-Content -LiteralPath $completePath -Raw | ConvertFrom-Json -AsHashtable
    if ($intent.schema -cne 'red-salamander.run-archive.v1' -or $intent.run_id -cne $RunId -or $intent.disposition -cne 'completed' -or
        $intent.destination -ine $run -or $intent.source -ine $expectedSource -or $complete.source -ine $expectedSource -or
        $complete.schema -cne 'red-salamander.run-archive-complete.v1' -or $complete.destination -ine $run -or
        $complete.intent_sha256 -ine (Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash) {
        throw 'Diagnostic archive proof is invalid.'
    }
    return @{
        RunDirectory = $run
        Intent = $intent
        ArchiveIntentSha256 = $complete.intent_sha256
    }
}

function Read-RSHostedDiagnosticEvidence {
    param([Parameter(Mandatory)][string]$TestRoot, [Parameter(Mandatory)][string]$RunId,
          [Parameter(Mandatory)][System.Collections.IDictionary]$Selection)
    Assert-RSHostedDiagnosticSelection $Selection
    $archive = Read-RSHostedVerifiedArchiveRecord -TestRoot $TestRoot -RunId $RunId
    $run = $archive.RunDirectory
    $intent = $archive.Intent
    $relative = 'artifacts\selftest\last_run\' + $Selection.native + '\results.json'
    $nativePath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $run $relative) -RequireDescendant
    $rows = @($intent.inventory | Where-Object { $_.path -ieq $relative -and -not $_.directory })
    if ($rows.Count -ne 1 -or [int64]$rows[0].length -ne (Get-Item -LiteralPath $nativePath).Length -or
        $rows[0].sha256 -ine (Get-FileHash -LiteralPath $nativePath -Algorithm SHA256).Hash) { throw 'Native diagnostic result does not match the archive.' }
    $summaryPath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $run 'artifacts/selftest/last_run/run-all-tests-results.json') -RequireDescendant
    $summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json -AsHashtable
    if ($summary.Contains('schema') -and $summary.schema -ceq 'red-salamander.run-all-tests.v2') {
        # Durable PR/CI/Full runs retain v1 beside their stable v2 summary.
        $summaryPath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $run 'artifacts/selftest/last_run/run-all-tests-results-v1.json') -RequireDescendant
    }
    return @{ summary=(Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json -AsHashtable);
        native=(Get-Content -LiteralPath $nativePath -Raw | ConvertFrom-Json -AsHashtable);
        summarySha256=(Get-FileHash -LiteralPath $summaryPath -Algorithm SHA256).Hash;
        nativeSha256=(Get-FileHash -LiteralPath $nativePath -Algorithm SHA256).Hash;
        archiveIntentSha256=$archive.ArchiveIntentSha256 }
}

function Get-RSHostedArchiveMemberPath {
    param($Archive, [string]$TestRoot, [string]$RelativePath)
    $rows = @($Archive.Intent.inventory | Where-Object { $_.path -ieq $RelativePath -and -not $_.directory })
    $path = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $Archive.RunDirectory $RelativePath) -RequireDescendant
    if ($rows.Count -ne 1 -or [int64]$rows[0].length -ne (Get-Item -LiteralPath $path).Length -or
        $rows[0].sha256 -ine (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash) {
        throw "Archived diagnostic member does not match its inventory: $RelativePath"
    }
    return $path
}

function Assert-RSHostedV1CommandsProof {
    param($Archive, [string]$TestRoot, [string]$RepoRoot, [string]$CaseFilter,
        [string]$Platform, [string]$Configuration, $Summary, $Receipt, $Snapshot,
        [DateTimeOffset]$InvocationStartedUtc, [DateTimeOffset]$InvocationEndedUtc)

    $candidates = @($Archive.Intent.inventory | Where-Object {
        -not $_.directory -and $_.path -imatch '^artifacts\\selftest\\(?:previous_run|last_run)\\entry-fingerprints\.json$'
    })
    if ($candidates.Count -ne 1) { throw 'Missing or ambiguous archived prelaunch fingerprint proof.' }
    $prelaunch = Split-Path -Parent $candidates[0].path
    $records = @{}
    foreach ($pair in @(
            @('fingerprints','entry-fingerprints.json','ValidationEntryFingerprints.schema.json'),
            @('plan','validation-plan.json','ValidationPlan.schema.json'),
            @('decision','validation-plan-decision.json','ValidationPlanDecision.schema.json'))) {
        $path = Get-RSHostedArchiveMemberPath $Archive $TestRoot (Join-Path $prelaunch $pair[1])
        $records[$pair[0]] = Read-RSValidationJsonFile -LiteralPath $path -AllowedRoot $TestRoot `
            -SchemaPath (Join-Path $RepoRoot "Specs/Testing/$($pair[2])")
    }
    $fp = $records.fingerprints; $plan = $records.plan; $decision = $records.decision
    $exe = [IO.Path]::GetFullPath((Join-Path $RepoRoot ".build/$Platform/$Configuration/RedSalamander.exe"))
    $entries = @(Get-RSTestRunPlan -Suite Commands -RepoRoot $RepoRoot -Platform $Platform -Configuration $Configuration `
        -RedSalamanderExePath $exe -CaseFilter $CaseFilter)
    $expectedPlan = New-RSValidationPlan -RepoRoot $RepoRoot -RequestedSuite Commands -Entries $entries -ValidationMode Fresh -CoverageScope filtered
    if ((Get-RSCanonicalJsonDigest -InputObject $plan) -cne (Get-RSCanonicalJsonDigest -InputObject $expectedPlan) -or
        $plan.plan_digest -cne $expectedPlan.plan_digest -or $plan.coverage_plan_digest -cne $expectedPlan.coverage_plan_digest -or
        $fp.plan_digest -cne $plan.plan_digest -or $fp.coverage_plan_digest -cne $plan.coverage_plan_digest -or
        $fp.snapshot_id -cne $Snapshot.SnapshotId -or $fp.build_receipt_id -cne $Receipt.receipt_id -or $fp.artifact_epoch -ne 0 -or
        @($fp.entries).Count -ne 1 -or $fp.entries[0].id -cne 'selftest.commands') {
        throw 'Archived Commands plan, profile, filter, source or receipt proof does not match.'
    }
    $fpProjection = [ordered]@{}
    foreach ($key in @('schema','canonicalization','plan_digest','coverage_plan_digest','snapshot_id','build_receipt_id','artifact_epoch','entries')) {
        $fpProjection[$key] = $fp.$key
    }
    $row = $fp.entries[0]
    $rowDigest = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{
        schema='red-salamander.entry-fingerprint.v2'; canonicalization='rs-canonical-json-v1'
        component_digest_version=$row.component_digest_version; component_digests=$row.component_digests })
    if ($fp.fingerprints_digest -cne (Get-RSCanonicalJsonDigest -InputObject $fpProjection) -or $row.fingerprint -cne $rowDigest) {
        throw 'Archived prelaunch fingerprint digest is invalid.'
    }
    # Reuse the fingerprint owner for deterministic source/receipt/plan components;
    # environment and runner discovery remain the archived invocation's observations.
    $expected = Get-RSValidationEntryFingerprint -EntryContract $expectedPlan.entries[0] `
        -CoveragePlanDigest $expectedPlan.coverage_plan_digest -WorkspaceSnapshot @{snapshot_id=$Snapshot.SnapshotId} -BuildReceipt $Receipt
    foreach ($component in @('arguments','coverage','outcome_policy','shared_contract','source_input','artifact_epoch','build_artifact')) {
        if ($row.component_digests.$component -cne $expected.ComponentDigests.$component) { throw "Archived fingerprint component mismatch: $component" }
    }
    $projection = [ordered]@{}
    foreach ($property in $decision.PSObject.Properties) { if ($property.Name -cne 'decision_digest') { $projection[$property.Name] = $property.Value } }
    if ($decision.decision_digest -cne (Get-RSCanonicalJsonDigest -InputObject $projection) -or
        $decision.planner_mode -cne 'shadow' -or
        $decision.plan_digest -cne $plan.plan_digest -or $decision.snapshot_id -cne $fp.snapshot_id -or
        @($decision.entries).Count -ne 1 -or $decision.entries[0].id -cne 'selftest.commands' -or
        $decision.entries[0].ordinal -ne 0 -or
        $decision.entries[0].decision -cne 'EXECUTE' -or $decision.entries[0].expected_fingerprint -cne $row.fingerprint -or
        $null -ne $decision.entries[0].origin_run_id -or $null -ne $decision.entries[0].origin_evidence_digest) {
        throw 'Archived Commands execution decision does not match its prelaunch proof.'
    }
    $snapshots = @()
    foreach ($relative in @((Join-Path $prelaunch 'workspace-initial.json'),
            'artifacts/selftest/last_run/workspace-after-selftest.commands.json', 'artifacts/selftest/last_run/workspace-final.json')) {
        $path = Get-RSHostedArchiveMemberPath $Archive $TestRoot $relative.Replace('/','\')
        $record = Read-RSValidationJsonFile -LiteralPath $path -AllowedRoot $TestRoot -SchemaPath (Join-Path $RepoRoot 'Specs/Testing/WorkspaceSnapshot.schema.json')
        if ($record.snapshot_id -cne $fp.snapshot_id -or $record.acquisition_before -cne $fp.snapshot_id -or
            $record.acquisition_after -cne $fp.snapshot_id -or $record.git_head -cne $Snapshot.GitHead -or
            $record.repo_identity -cne $Receipt.repo_identity) { throw 'Archived Commands workspace source identity changed.' }
        Test-RSHostedTimeWindow -Summary @{run_started_utc=$record.created_utc;run_ended_utc=$record.created_utc} `
            -InvocationStartedUtc $InvocationStartedUtc -InvocationEndedUtc $InvocationEndedUtc
        $snapshots += [DateTimeOffset]$record.created_utc
    }
    $suites = @($Summary.suites)
    if ($suites.Count -ne 1 -or $suites[0].suite -cne 'Commands' -or $suites[0].kind -cne 'SelfTest' -or
        $suites[0].classification -cne 'PASSED' -or [int]$suites[0].exit_code -ne 0 -or
        [int64]$suites[0].passed -ne [int64]$Summary.passed -or [int64]$suites[0].failed -ne 0 -or
        [int64]$suites[0].skipped -ne 0 -or @($suites[0].retry_attempts).Count -ne 0) {
        throw 'Commands v1 native execution did not pass once without retries.'
    }
    $exits = @($Archive.Intent.inventory | Where-Object {
        -not $_.directory -and $_.path -imatch '^artifacts\\native-process-exits\\[0-9a-f]{32}\.json$'
    })
    if ($exits.Count -ne 1) { throw 'Missing or ambiguous owned native exit proof.' }
    $path = Get-RSHostedArchiveMemberPath $Archive $TestRoot $exits[0].path
    $exit = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    if ($exit.schema -cne 'red-salamander.native-process-exit.v1' -or ($exit.process_id -isnot [long] -and $exit.process_id -isnot [int]) -or
        $exit.process_id -le 0 -or $exit.executable -ine $exe -or [int]$exit.exit_code -ne 0) { throw 'Owned native PID, executable or exit code does not match.' }
    Test-RSHostedTimeWindow -Summary @{run_started_utc=$exit.observed_utc;run_ended_utc=$exit.observed_utc} `
        -InvocationStartedUtc ([DateTimeOffset]$Summary.run_started_utc) -InvocationEndedUtc ([DateTimeOffset]$Summary.run_ended_utc)
    $observed = [DateTimeOffset]$exit.observed_utc
    if ($snapshots[0] -gt [DateTimeOffset]$Summary.run_started_utc -or $snapshots[1] -lt $observed -or
        $snapshots[2] -lt $snapshots[1]) { throw 'Archived prelaunch/postlaunch snapshot ordering is invalid.' }
    $metrics = Get-RSHostedArchiveMemberPath $Archive $TestRoot 'artifacts\selftest\last_run\perf\perf_metrics.jsonl'
    $count = 0
    $reader = [IO.File]::OpenText($metrics)
    try {
        while ($null -ne ($line = $reader.ReadLine())) {
            $metric = $line | ConvertFrom-Json -AsHashtable -ErrorAction Stop
            if (($metric.process -isnot [long] -and $metric.process -isnot [int]) -or $metric.process -ne $exit.process_id) {
                throw 'Native metric PID does not match the owned exit.'
            }
            Test-RSHostedTimeWindow -Summary @{run_started_utc=$metric.timestamp;run_ended_utc=$metric.timestamp} `
                -InvocationStartedUtc ([DateTimeOffset]$Summary.run_started_utc) -InvocationEndedUtc $observed
            ++$count
        }
    } finally { $reader.Dispose() }
    if ($count -eq 0) { throw 'Owned native metric evidence is empty.' }
}

function Assert-RSHostedFilteredCaseRun {
    <#
    .SYNOPSIS
        Verifies v1 single-suite or v2 durable filtered Commands evidence, including
        archived prelaunch source/receipt/plan and owned native PID/time proof for v1.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$ExitCode,
        [Parameter(Mandatory)][string]$Output,
        [Parameter(Mandatory)][string]$TestRoot,
        [Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][ValidateSet('Commands')][string]$Suite,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9_]+$')][string]$CaseFilter,
        [Parameter(Mandatory)][string]$Platform,
        [Parameter(Mandatory)][string]$Configuration,
        [Parameter(Mandatory)][DateTimeOffset]$InvocationStartedUtc,
        [Parameter(Mandatory)][DateTimeOffset]$InvocationEndedUtc
    )

    if ($ExitCode -ne 2) {
        throw "Filtered hosted runner for '$Suite' filter '$CaseFilter' failed with exit code $ExitCode."
    }
    $runId = Get-RSHostedDiagnosticRunId -Output $Output
    $archive = Read-RSHostedVerifiedArchiveRecord -TestRoot $TestRoot -RunId $runId
    $run = $archive.RunDirectory
    $lastRun = Join-Path $run 'artifacts/selftest/last_run'
    $v2Path = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $lastRun 'run-all-tests-results.json') -RequireDescendant
    $v1Path = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $lastRun 'run-all-tests-results-v1.json') -RequireDescendant
    $stable = Get-Content -LiteralPath $v2Path -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    $v2 = $null
    if ($stable.schema -ceq 'red-salamander.run-all-tests.v2') {
        $v2 = Read-RSValidationJsonFile -LiteralPath $v2Path -AllowedRoot $TestRoot `
            -SchemaPath (Join-Path $RepoRoot 'Specs/Testing/RunAllTestsSummary.schema.json')
        $v1 = Get-Content -LiteralPath $v1Path -Raw | ConvertFrom-Json -AsHashtable
    } elseif ($stable.schema -ceq 'red-salamander.run-all-tests.v1') {
        $v1 = $stable
    } else { throw 'Unsupported filtered runner summary schema.' }

    if ($null -ne $v2) {
    if ($v2.schema -cne 'red-salamander.run-all-tests.v2' -or $v2.run_id -cne $runId -or
        $v2.requested_gate -cne $Suite -or $v2.change_verdict -cne 'PASSED' -or
        $v2.validation_mode -cne 'Fresh' -or $v2.coverage_scope -cne 'filtered' -or
        @($v2.entries).Count -ne 1 -or $v2.entries[0].id -cne 'selftest.commands' -or
        $v2.entries[0].execution -cne 'executed' -or $v2.entries[0].promotion_state -cne 'promoted') {
        throw "Filtered hosted v2 summary for '$CaseFilter' did not reach change_verdict PASSED."
    }
    if ($v2.repository_verdict -cne 'NOT_EVALUATED') {
        throw "Filtered hosted exit 2 for '$CaseFilter' expected repository_verdict NOT_EVALUATED but got '$($v2.repository_verdict)'."
    }
    if ([int]$v2.counts.executed -ne 1 -or [int]$v2.counts.promoted -ne 1 -or
        [int]$v2.counts.reused -ne 0 -or [int]$v2.counts.provisional -ne 0 -or
        [int]$v2.counts.invalidated -ne 0 -or [int]$v2.counts.incomplete -ne 0 -or
        [int64]$v2.case_counts.passed -lt 1 -or [int64]$v2.case_counts.failed -ne 0 -or [int64]$v2.case_counts.skipped -ne 0) {
        throw "Filtered hosted case counts for '$CaseFilter' are incomplete, skipped, or failed."
    }
    }

    if ($v1.schema -cne 'red-salamander.run-all-tests.v1' -or $v1.run_id -cne $runId -or
        $v1.suite -cne $Suite -or $v1.case_filter -cne $CaseFilter -or
        $v1.platform -cne $Platform -or $v1.configuration -cne $Configuration -or
        [int]$v1.exit_code -ne $ExitCode -or
        [int64]$v1.passed -lt 1 -or [int64]$v1.failed -ne 0 -or [int64]$v1.skipped -ne 0) {
        throw "Filtered hosted v1 summary for '$CaseFilter' does not match the expected run identity or counts."
    }
    Test-RSHostedTimeWindow -Summary $v1 -InvocationStartedUtc $InvocationStartedUtc -InvocationEndedUtc $InvocationEndedUtc
    Assert-RSHostedCleanQuarantine -Summary $v1

    # Aggregates are written after archival, so native bytes (not aggregates) are
    # authenticated against its inventory. Require actual selected passing cases.
    $relative = 'artifacts\selftest\last_run\commands\results.json'
    $nativePath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $run $relative) -RequireDescendant
    $rows = @($archive.Intent.inventory | Where-Object { $_.path -ieq $relative -and -not $_.directory })
    if ($rows.Count -ne 1 -or [int64]$rows[0].length -ne (Get-Item -LiteralPath $nativePath).Length -or
        $rows[0].sha256 -ine (Get-FileHash -LiteralPath $nativePath -Algorithm SHA256).Hash) {
        throw 'Native filtered result does not match the archive.'
    }
    $native = Get-Content -LiteralPath $nativePath -Raw | ConvertFrom-Json -AsHashtable
    if ($native.suite -cne 'Commands') { throw 'Native filtered suite mismatch.' }
    if ($null -eq $v2) {
        Test-RSHostedTimeWindow -Summary @{run_started_utc=$native.run_started_utc;run_ended_utc=$native.run_started_utc} `
            -InvocationStartedUtc ([DateTimeOffset]$v1.run_started_utc) -InvocationEndedUtc ([DateTimeOffset]$v1.run_ended_utc)
    }
    $names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($case in @($native.cases)) {
        $matches = [string]::Equals($case.name, $CaseFilter, [StringComparison]::OrdinalIgnoreCase) -or
            ($CaseFilter.EndsWith('_') -and ([string]$case.name).StartsWith($CaseFilter, [StringComparison]::OrdinalIgnoreCase))
        if (-not $matches -or $case.status -cne 'passed' -or -not $names.Add([string]$case.name)) {
            throw 'Native filtered cases are unexpected, duplicated, or not passing.'
        }
    }
    if ($names.Count -lt 1 -or [int64]$native.passed -ne $names.Count -or [int64]$native.failed -ne 0 -or
        [int64]$native.skipped -ne 0 -or [int64]$v1.total -ne $names.Count -or
        [int64]$v1.passed -ne $names.Count -or ($null -ne $v2 -and ([int64]$v2.case_counts.total -ne $names.Count -or
        [int64]$v2.case_counts.passed -ne $names.Count))) { throw 'Native filtered case counts do not match runner summaries.' }

    $buildOutput = Join-Path $RepoRoot ".build/$Platform/$Configuration"
    $receipt = Read-RSBuildReceipt -BuildOutputDir $buildOutput -SchemaPath (Join-Path $RepoRoot 'Specs/Build/BuildReceipt.schema.json')
    $snapshot = Get-RSBuildSourceSnapshot -RepoRoot $RepoRoot
    $requiredExe = ".build/$Platform/$Configuration/RedSalamander.exe"
    if ($v1.executable -ine [IO.Path]::GetFullPath((Join-Path $RepoRoot $requiredExe)) -or
        $null -eq $receipt -or ($null -ne $v2 -and ($v2.snapshot_id -cne $snapshot.SnapshotId -or
        $v2.build_receipt_id -cne $receipt.receipt_id)) -or
        -not (Test-RSBuildReceiptForArtifactUse -Receipt $receipt -SourceSnapshot $snapshot -Platform $Platform `
            -Configuration $Configuration -RepoRoot $RepoRoot -RequiredRelativePath @($requiredExe) `
            -RequiredProjectSelection @('RedSalamander/RedSalamander.vcxproj'))) {
        throw 'Native filtered source, executable, or artifact receipt mismatch.'
    }
    if ($null -eq $v2) {
        Assert-RSHostedV1CommandsProof -Archive $archive -TestRoot $TestRoot -RepoRoot $RepoRoot -CaseFilter $CaseFilter `
            -Platform $Platform -Configuration $Configuration -Summary $v1 -Receipt $receipt -Snapshot $snapshot `
            -InvocationStartedUtc $InvocationStartedUtc -InvocationEndedUtc $InvocationEndedUtc
    }

    return [pscustomobject]@{
        RunId = $runId
        RepositoryVerdict = 'NOT_EVALUATED'
        ChangeVerdict = 'PASSED'
        Passed = [int64]$v1.passed
        ArchiveIntentSha256 = $archive.ArchiveIntentSha256
    }
}

Export-ModuleMember -Function Get-RSHostedDiagnosticSelection, Get-RSHostedDiagnosticRunId, Test-RSHostedDiagnosticOutcome, Read-RSHostedDiagnosticEvidence, Assert-RSHostedFilteredCaseRun, Get-RSHostedPairedDispatch
