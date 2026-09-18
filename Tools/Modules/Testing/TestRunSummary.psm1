Set-StrictMode -Version Latest

Import-Module (Join-Path $PSScriptRoot 'ValidationFingerprint.psm1') -Scope Local -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'ValidationEvidence.psm1') -Scope Local -ErrorAction Stop

Import-Module (Join-Path $PSScriptRoot 'TestInvocation.psm1') -Force -Scope Local

function Get-RSValidationExitCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('PASSED', 'FAILED', 'NOT_EVALUATED', 'INVALID_CLI_OR_PLAN', 'BLOCKED_ATTESTATION', 'CORRUPT_EVIDENCE')]
        [string]$TerminalState
    )

    switch ($TerminalState) {
        'PASSED' { return 0 }
        'FAILED' { return 1 }
        'NOT_EVALUATED' { return 2 }
        'INVALID_CLI_OR_PLAN' { return 3 }
        'BLOCKED_ATTESTATION' { return 4 }
        'CORRUPT_EVIDENCE' { return 5 }
    }
    throw "Unhandled validation terminal state '$TerminalState'."
}

function Convert-RSTestCaseForRunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Case
    )

    [pscustomobject]@{
        name = [string](Get-RSObjectValue -Object $Case -Names @('name', 'Name', 'CaseName') -DefaultValue '(unnamed case)')
        status = [string](Get-RSObjectValue -Object $Case -Names @('status', 'Status') -DefaultValue '')
        reason = [string](Get-RSObjectValue -Object $Case -Names @('reason', 'Reason') -DefaultValue '')
        duration_ms = [uint64](Get-RSObjectValue -Object $Case -Names @('durationMs', 'duration_ms', 'Duration') -DefaultValue 0)
        attempt = [int](Get-RSObjectValue -Object $Case -Names @('attempt', 'repeat_index', 'repeatIndex') -DefaultValue 1)
    }
}

function Convert-RSTestRetryAttemptForRunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$RetryAttempt
    )

    [pscustomobject]@{
        name = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('name', 'Name') -DefaultValue '(unnamed retry)')
        mode = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('mode', 'Mode') -DefaultValue '')
        exit_code = [int](Get-RSObjectValue -Object $RetryAttempt -Names @('exit_code', 'ExitCode') -DefaultValue 0)
        passed = [bool](Get-RSObjectValue -Object $RetryAttempt -Names @('passed', 'Passed') -DefaultValue $false)
        duration_ms = [uint64](Get-RSObjectValue -Object $RetryAttempt -Names @('duration_ms', 'DurationMs', 'WallMs') -DefaultValue 0)
        output_log_path = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('output_log_path', 'OutputLogPath') -DefaultValue '')
        reason = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('reason', 'Reason') -DefaultValue '')
        shuffle_seed = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('shuffle_seed', 'ShuffleSeed') -DefaultValue '')
    }
}

function Convert-RSTestQuarantineRepairAttemptForRunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$RepairAttempt
    )

    [pscustomobject]@{
        harness = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('harness', 'Harness') -DefaultValue '')
        name = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('name', 'Name') -DefaultValue '')
        mode = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('mode', 'Mode') -DefaultValue '')
        owner = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('owner', 'Owner') -DefaultValue '')
        expires = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('expires', 'Expires') -DefaultValue '')
        issue = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('issue', 'Issue') -DefaultValue '')
        exit_code = [int](Get-RSObjectValue -Object $RepairAttempt -Names @('exit_code', 'ExitCode') -DefaultValue -1)
        passed = [bool](Get-RSObjectValue -Object $RepairAttempt -Names @('passed', 'Passed') -DefaultValue $false)
        duration_ms = [uint64](Get-RSObjectValue -Object $RepairAttempt -Names @('duration_ms', 'DurationMs', 'WallMs') -DefaultValue 0)
        output_log_path = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('output_log_path', 'OutputLogPath') -DefaultValue '')
        reason = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('reason', 'Reason') -DefaultValue '')
    }
}

function New-RSTestRunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'PR', 'CI', 'Full')]
        [string]$Suite,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$ExePath,

        [Parameter(Mandatory = $true)]
        [string]$ArtifactRoot,

        [Parameter(Mandatory = $true)]
        [datetime]$RunStartedUtc,

        [Parameter(Mandatory = $true)]
        [datetime]$RunEndedUtc,

        [uint64]$DurationMs = 0,

        [int]$ExitCode = 0,

        [double]$TimeoutMultiplier = 1.0,

        [bool]$FailFast = $false,

        [string]$CaseFilter = '',

        [string]$TestRoot = '',

        [string]$RunId = '',

        [object]$QuarantineStatus = $null,

        [object[]]$QuarantineRepairAttempts = @(),

        [object]$TestSandboxAudit = $null,

        [object[]]$Results = @()
    )

    $totalPassed = 0
    $totalFailed = 0
    $totalSkipped = 0
    $classificationCounts = [ordered]@{
        passed = 0
        flaky = 0
        regression = 0
        isolation_suspect = 0
        unclassified_failure = 0
    }
    $summarySuites = @()

    foreach ($result in @($Results)) {
        $passed = [int](Get-RSObjectValue -Object $result -Names @('Passed', 'passed') -DefaultValue 0)
        $failed = [int](Get-RSObjectValue -Object $result -Names @('Failed', 'failed') -DefaultValue 0)
        $skipped = [int](Get-RSObjectValue -Object $result -Names @('Skipped', 'skipped') -DefaultValue 0)

        $totalPassed += $passed
        $totalFailed += $failed
        $totalSkipped += $skipped

        $cases = @(Get-RSObjectValue -Object $result -Names @('Cases', 'cases') -DefaultValue @() | ForEach-Object {
                Convert-RSTestCaseForRunSummary -Case $_
            })
        $failures = @(Get-RSObjectValue -Object $result -Names @('Failures', 'failures') -DefaultValue @() | ForEach-Object {
                Convert-RSTestCaseForRunSummary -Case $_
            })
        $retryAttempts = @(Get-RSObjectValue -Object $result -Names @('RetryAttempts', 'retry_attempts') -DefaultValue @() | ForEach-Object {
                Convert-RSTestRetryAttemptForRunSummary -RetryAttempt $_
            })
        $classification = [string](Get-RSObjectValue -Object $result -Names @('Classification', 'classification') -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($classification)) {
            $classification = if (Test-RSTestResultPassed -Result $result) { 'PASSED' } else { 'UNCLASSIFIED_FAILURE' }
        }
        $classificationReason = [string](Get-RSObjectValue -Object $result -Names @('ClassificationReason', 'classification_reason') -DefaultValue '')

        switch ($classification) {
            'PASSED' { ++$classificationCounts.passed }
            'FLAKY' { ++$classificationCounts.flaky }
            'REGRESSION' { ++$classificationCounts.regression }
            'ISOLATION_SUSPECT' { ++$classificationCounts.isolation_suspect }
            default { ++$classificationCounts.unclassified_failure }
        }

        $summarySuites += [pscustomobject]@{
            suite = [string](Get-RSObjectValue -Object $result -Names @('Name', 'suite', 'name') -DefaultValue '(unnamed suite)')
            kind = [string](Get-RSObjectValue -Object $result -Names @('Kind', 'kind') -DefaultValue '')
            exit_code = [int](Get-RSObjectValue -Object $result -Names @('ExitCode', 'exit_code') -DefaultValue 0)
            duration_ms = [uint64](Get-RSObjectValue -Object $result -Names @('WallMs', 'DurationMs', 'duration_ms') -DefaultValue 0)
            classification = $classification
            classification_reason = $classificationReason
            shuffle_seed = [string](Get-RSObjectValue -Object $result -Names @('ShuffleSeed', 'shuffle_seed') -DefaultValue '')
            repeat_count = [int](Get-RSObjectValue -Object $result -Names @('RepeatCount', 'repeat_count') -DefaultValue 1)
            passed = $passed
            failed = $failed
            skipped = $skipped
            output_log_path = [string](Get-RSObjectValue -Object $result -Names @('OutputLogPath', 'output_log_path') -DefaultValue '')
            retry_attempts = $retryAttempts
            cases = $cases
            failures = $failures
        }
    }

    $quarantineActiveEntries = @()
    $quarantineInvalidEntries = @()
    $quarantineErrors = @()
    $quarantineIsValid = $true
    $quarantineHasBlockingEntries = $false
    if ($null -ne $QuarantineStatus) {
        $quarantineActiveEntries = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('ActiveEntries', 'active_entries') -DefaultValue @())
        $quarantineInvalidEntries = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('InvalidEntries', 'invalid_entries') -DefaultValue @())
        $quarantineErrors = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('Errors', 'errors') -DefaultValue @())
        $quarantineIsValid = [bool](Get-RSObjectValue -Object $QuarantineStatus -Names @('IsValid', 'is_valid') -DefaultValue $true)
        $quarantineHasBlockingEntries = [bool](Get-RSObjectValue -Object $QuarantineStatus -Names @('HasBlockingEntries', 'has_blocking_entries') -DefaultValue $false)
    }
    $quarantineRepairAttempts = @($QuarantineRepairAttempts | ForEach-Object {
            Convert-RSTestQuarantineRepairAttemptForRunSummary -RepairAttempt $_
        })
    $quarantineRepairReproducedCount = @($quarantineRepairAttempts | Where-Object { -not [bool]$_.passed }).Count

    [pscustomobject]@{
        schema = 'red-salamander.run-all-tests.v1'
        run_started_utc = $RunStartedUtc.ToString('o')
        run_ended_utc = $RunEndedUtc.ToString('o')
        duration_ms = $DurationMs
        suite = $Suite
        platform = $Platform
        configuration = $Configuration
        executable = $ExePath
        artifact_root = $ArtifactRoot
        test_root = $TestRoot
        run_id = $RunId
        fail_fast = $FailFast
        timeout_scale = $TimeoutMultiplier
        case_filter = $CaseFilter
        exit_code = $ExitCode
        passed = $totalPassed
        failed = $totalFailed
        skipped = $totalSkipped
        total = ($totalPassed + $totalFailed + $totalSkipped)
        classifications = [pscustomobject]$classificationCounts
        quarantine = [pscustomobject]@{
            is_valid = $quarantineIsValid
            has_blocking_entries = $quarantineHasBlockingEntries
            active_count = @($quarantineActiveEntries).Count
            invalid_count = @($quarantineInvalidEntries).Count
            active_entries = $quarantineActiveEntries
            invalid_entries = $quarantineInvalidEntries
            errors = $quarantineErrors
            repair_attempt_count = @($quarantineRepairAttempts).Count
            repair_reproduced_count = $quarantineRepairReproducedCount
            repair_attempts = $quarantineRepairAttempts
        }
        test_sandbox_audit = $TestSandboxAudit
        suites = $summarySuites
    }
}

function New-RSValidationRunSummaryV2 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Run,
        [Parameter(Mandatory = $true)][object]$Plan,
        [Parameter(Mandatory = $true)][object]$Decision,
        [Parameter(Mandatory = $true)][object]$LegacySummary,
        [Parameter(Mandatory = $true)][string]$EvidenceRoot,
        [Parameter(Mandatory = $true)][string]$SchemaRoot,
        [Parameter(Mandatory = $true)][ValidateSet('complete', 'filtered', 'family', 'capability-reduced')][string]$CoverageScope
    )

    $entryRows = @()
    $verifiedRunArtifactBytes = [int64]0
    $corrupt = $Plan.plan_digest -ne $Run.State.plan_digest -or
        $Plan.coverage_plan_digest -ne $Run.State.coverage_plan_digest -or
        $Plan.validation_mode -ne $Run.State.validation_mode -or
        $Plan.requested_suite -ne $Run.State.requested_suite -or
        $Decision.plan_digest -ne $Run.State.plan_digest
    foreach ($stateEntry in @($Run.State.entries | Sort-Object ordinal)) {
        $planEntry = @($Plan.entries | Where-Object id -eq $stateEntry.id | Select-Object -First 1)
        if ($planEntry.Count -ne 1) { throw "Aggregate plan entry is missing: $($stateEntry.id)" }
        $originRunId = if ($stateEntry.status -eq 'reused') { [string]$stateEntry.origin_run_id } else { [string]$Run.State.run_id }
        $originDirectory = Join-Path (Join-Path $EvidenceRoot 'runs') $originRunId
        $promotion = $null
        $evidence = $null
        if ($null -ne $stateEntry.promotion_digest) {
            try {
                $verified = Read-RSVerifiedValidationPromotionEvidence `
                    -OriginRunDirectory $originDirectory `
                    -EntryId $stateEntry.id `
                    -PromotionDigest $stateEntry.promotion_digest `
                    -SchemaRoot $SchemaRoot `
                    -ExpectedCoveragePlanDigest $Plan.coverage_plan_digest `
                    -ExpectedEntryFingerprint $stateEntry.expected_fingerprint `
                    -RequireTerminalGreen:($stateEntry.status -eq 'reused')
                $promotion = $verified.Promotion
                $evidence = $verified.Evidence
                $verifiedRunArtifactBytes += [int64]$verified.ResultArtifactBytes
                if ($verifiedRunArtifactBytes -gt 1GB) {
                    throw 'Verified validation result artifacts exceed the 1 GiB run quota.'
                }
                if ($stateEntry.status -eq 'reused' -and
                    ($stateEntry.origin_run_id -ne $verified.State.run_id -or
                     $stateEntry.origin_evidence_digest -ne $evidence.evidence_digest -or
                     $Run.State.build_receipt_id -ne $evidence.build_receipt_id)) {
                    throw "Reused validation state does not match verified origin '$($stateEntry.id)'."
                }
            } catch {
                $corrupt = $true
                $promotion = $null
                $evidence = $null
            }
        } elseif ($stateEntry.status -in @('promoted', 'reused')) {
            $corrupt = $true
        }
        $promotionState = switch ([string]$stateEntry.status) {
            'promoted' { if ($null -ne $promotion) { 'promoted' } else { 'incomplete' } }
            'reused' { if ($null -ne $promotion) { 'promoted' } else { 'incomplete' } }
            'provisional-passed' { 'provisional' }
            'invalidated' { 'invalidated' }
            default { 'incomplete' }
        }
        $evidenceDigest = if ($null -ne $evidence) { [string]$evidence.evidence_digest } else {
            Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ state = 'missing'; run_id = $Run.State.run_id; entry_id = $stateEntry.id })
        }
        $entryRows += [pscustomobject][ordered]@{
            id = $stateEntry.id
            execution = if ($stateEntry.status -eq 'reused') { 'reused' } else { 'executed' }
            promotion_state = $promotionState
            evidence_digest = $evidenceDigest
            promotion_digest = if ($null -eq $promotion) { $null } else { [string]$promotion.promotion_digest }
            origin_run_id = if ($null -eq $promotion) { $null } else { $originRunId }
            artifact_epoch = if ($null -eq $evidence) { [int64]$Run.State.artifact_epoch } else { [int64]$evidence.artifact_epoch }
            artifact_read_ids = @($planEntry[0].artifact_read_ids)
            artifact_write_ids = @($planEntry[0].artifact_write_ids)
            produces_artifact_ids = @($planEntry[0].produces_artifact_ids)
            cache_policy = $planEntry[0].cache_policy
            reason_codes = @($stateEntry.reason_codes)
            coverage_digest = if ($null -eq $evidence) { Get-RSCanonicalJsonDigest $planEntry[0].coverage_contract } else { [string]$evidence.coverage_digest }
            outcome_digest = if ($null -eq $evidence) { Get-RSCanonicalJsonDigest $planEntry[0].outcome_contract } else { [string]$evidence.outcome_digest }
            capability_digest = $null
            duration_ms = if ($null -eq $evidence) { 0L } else { [int64]$evidence.duration_ms }
        }
    }
    $counts = [ordered]@{
        executed = [int64]@($entryRows | Where-Object execution -eq executed).Count
        reused = [int64]@($entryRows | Where-Object execution -eq reused).Count
        provisional = [int64]@($entryRows | Where-Object promotion_state -eq provisional).Count
        promoted = [int64]@($entryRows | Where-Object promotion_state -eq promoted).Count
        invalidated = [int64]@($entryRows | Where-Object promotion_state -eq invalidated).Count
        incomplete = [int64]@($entryRows | Where-Object promotion_state -eq incomplete).Count
    }
    $hasFailure = [int]$LegacySummary.exit_code -ne 0 -or $Run.State.status -eq 'failed'
    $complete = $CoverageScope -eq 'complete' -and $entryRows.Count -eq @($Plan.entries).Count -and
        $counts.provisional -eq 0 -and $counts.invalidated -eq 0 -and $counts.incomplete -eq 0
    $repositoryVerdict = if ($corrupt) { 'BLOCKED' } elseif ($hasFailure) { 'FAILED' } elseif ($complete) { 'PASSED' } else { 'NOT_EVALUATED' }
    $changeVerdict = if ($corrupt) { 'BLOCKED' } elseif ($hasFailure) { 'FAILED' } elseif ($counts.incomplete -eq 0 -and $counts.invalidated -eq 0) { 'PASSED' } else { 'NOT_EVALUATED' }
    $commands = @($entryRows | Where-Object id -eq 'selftest.commands' | Select-Object -First 1)
    $commandsState = if ($commands.Count -eq 0) { 'absent' } elseif ($commands[0].promotion_state -ne 'promoted') { 'invalid' } elseif ($commands[0].execution -eq 'reused') { 'exact-reused' } else { 'fresh' }
    $avoidedMeasure = @($entryRows | Where-Object execution -eq reused | Measure-Object duration_ms -Sum)
    $avoidedDuration = if ($avoidedMeasure.Count -eq 0 -or $null -eq $avoidedMeasure[0].Sum) { 0L } else { [int64]$avoidedMeasure[0].Sum }
    return [pscustomobject][ordered]@{
        schema = 'red-salamander.run-all-tests.v2'
        run_id = $Run.State.run_id
        validation_mode = $Run.State.validation_mode
        requested_gate = $Plan.requested_suite
        coverage_scope = $CoverageScope
        snapshot_id = $Run.State.snapshot_id
        build_receipt_id = $Run.State.build_receipt_id
        plan_digest = $Run.State.plan_digest
        coverage_plan_digest = $Run.State.coverage_plan_digest
        decision_digest = $Decision.decision_digest
        repository_verdict = $repositoryVerdict
        change_verdict = $changeVerdict
        counts = [pscustomobject]$counts
        case_counts = [pscustomobject][ordered]@{
            passed = [int64]$LegacySummary.passed; failed = [int64]$LegacySummary.failed
            skipped = [int64]$LegacySummary.skipped; total = [int64]$LegacySummary.total
        }
        duration_ms = [int64]$LegacySummary.duration_ms
        estimated_avoided_duration_ms = $avoidedDuration
        commands_broad_seal_state = $commandsState
        entries = $entryRows
    }
}

function Convert-RSTestRunSummaryToCaseHistoryRows {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Summary
    )

    $runId = [string](Get-RSObjectValue -Object $Summary -Names @('run_id', 'RunId') -DefaultValue '')
    $runStartedUtc = [string](Get-RSObjectValue -Object $Summary -Names @('run_started_utc', 'RunStartedUtc') -DefaultValue '')
    $suiteSelection = [string](Get-RSObjectValue -Object $Summary -Names @('suite', 'Suite') -DefaultValue '')
    $rows = @()

    foreach ($suite in @(Get-RSObjectValue -Object $Summary -Names @('suites', 'Suites') -DefaultValue @())) {
        $harness = [string](Get-RSObjectValue -Object $suite -Names @('suite', 'name', 'Name') -DefaultValue '')
        $classification = [string](Get-RSObjectValue -Object $suite -Names @('classification', 'Classification') -DefaultValue '')
        $suiteSeed = [string](Get-RSObjectValue -Object $suite -Names @('shuffle_seed', 'ShuffleSeed', 'seed') -DefaultValue '')

        foreach ($case in @(Get-RSObjectValue -Object $suite -Names @('cases', 'Cases') -DefaultValue @())) {
            $rows += [pscustomobject][ordered]@{
                schema = 'red-salamander.case-history.v1'
                run_id = $runId
                run_started_utc = $runStartedUtc
                suite = $suiteSelection
                harness = $harness
                case = [string](Get-RSObjectValue -Object $case -Names @('name', 'Name') -DefaultValue '')
                duration_ms = [uint64](Get-RSObjectValue -Object $case -Names @('duration_ms', 'durationMs', 'Duration') -DefaultValue 0)
                status = [string](Get-RSObjectValue -Object $case -Names @('status', 'Status') -DefaultValue '')
                reason = [string](Get-RSObjectValue -Object $case -Names @('reason', 'Reason') -DefaultValue '')
                classification = $classification
                seed = $suiteSeed
                attempt = [int](Get-RSObjectValue -Object $case -Names @('attempt', 'repeat_index', 'repeatIndex') -DefaultValue 1)
                source = 'case'
                mode = ''
            }
        }

        $retryAttemptIndex = 0
        foreach ($retryAttempt in @(Get-RSObjectValue -Object $suite -Names @('retry_attempts', 'RetryAttempts') -DefaultValue @())) {
            ++$retryAttemptIndex
            $retryPassed = [bool](Get-RSObjectValue -Object $retryAttempt -Names @('passed', 'Passed') -DefaultValue $false)
            $retrySeed = [string](Get-RSObjectValue -Object $retryAttempt -Names @('shuffle_seed', 'ShuffleSeed') -DefaultValue '')
            if ([string]::IsNullOrWhiteSpace($retrySeed)) {
                $retrySeed = $suiteSeed
            }

            $rows += [pscustomobject][ordered]@{
                schema = 'red-salamander.case-history.v1'
                run_id = $runId
                run_started_utc = $runStartedUtc
                suite = $suiteSelection
                harness = $harness
                case = [string](Get-RSObjectValue -Object $retryAttempt -Names @('name', 'Name') -DefaultValue '')
                duration_ms = [uint64](Get-RSObjectValue -Object $retryAttempt -Names @('duration_ms', 'DurationMs', 'WallMs') -DefaultValue 0)
                status = if ($retryPassed) { 'passed' } else { 'failed' }
                reason = [string](Get-RSObjectValue -Object $retryAttempt -Names @('reason', 'Reason') -DefaultValue '')
                classification = $classification
                seed = $retrySeed
                attempt = $retryAttemptIndex
                source = 'retry'
                mode = [string](Get-RSObjectValue -Object $retryAttempt -Names @('mode', 'Mode') -DefaultValue '')
            }
        }
    }

    return $rows
}

function Convert-RSTestCaseHistoryRowsToJsonl {
    param(
        [object[]]$Rows = @()
    )

    return (@($Rows) | ForEach-Object {
            $_ | ConvertTo-Json -Compress -Depth 8
        }) -join "`n"
}

function Convert-RSTestCaseHistoryRowsToDashboardMarkdown {
    param(
        [object[]]$Rows = @(),

        [Parameter(Mandatory = $true)]
        [object]$Summary,

        [int]$TopCount = 20
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $runId = [string](Get-RSObjectValue -Object $Summary -Names @('run_id', 'RunId') -DefaultValue '')
    $suite = [string](Get-RSObjectValue -Object $Summary -Names @('suite', 'Suite') -DefaultValue '')
    $lines.Add('# Run-AllTests Case Dashboard')
    $lines.Add('')
    $lines.Add(('- Run: `{0}`' -f $runId))
    $lines.Add(('- Suite: `{0}`' -f $suite))
    $lines.Add('')

    $lines.Add('## Top Slowest Cases')
    $lines.Add('')
    $lines.Add('| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |')
    $lines.Add('|---|---|---|---|---|---:|---|---:|')
    $topRows = @($Rows | Sort-Object -Property duration_ms -Descending | Select-Object -First $TopCount)
    foreach ($row in $topRows) {
        $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} |' -f `
                    (ConvertTo-RSMarkdownTableCell $row.harness), `
                    (ConvertTo-RSMarkdownTableCell $row.case), `
                    (ConvertTo-RSMarkdownTableCell $row.source), `
                    (ConvertTo-RSMarkdownTableCell $row.status), `
                    (ConvertTo-RSMarkdownTableCell $row.classification), `
                    (ConvertTo-RSMarkdownTableCell $row.duration_ms), `
                    (ConvertTo-RSMarkdownTableCell $row.seed), `
                    (ConvertTo-RSMarkdownTableCell $row.attempt)))
    }
    if (@($topRows).Count -eq 0) {
        $lines.Add('| (none) |  |  |  |  | 0 |  | 0 |')
    }

    $lines.Add('')
    $lines.Add('## Failure And Retry Evidence')
    $lines.Add('')
    $lines.Add('| Harness | Case | Source | Status | Classification | Seed | Reason |')
    $lines.Add('|---|---|---|---|---|---|---|')
    $evidenceRows = @($Rows | Where-Object { $_.status -ne 'passed' -or $_.source -eq 'retry' })
    foreach ($row in $evidenceRows) {
        $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} |' -f `
                    (ConvertTo-RSMarkdownTableCell $row.harness), `
                    (ConvertTo-RSMarkdownTableCell $row.case), `
                    (ConvertTo-RSMarkdownTableCell $row.source), `
                    (ConvertTo-RSMarkdownTableCell $row.status), `
                    (ConvertTo-RSMarkdownTableCell $row.classification), `
                    (ConvertTo-RSMarkdownTableCell $row.seed), `
                    (ConvertTo-RSMarkdownTableCell $row.reason)))
    }
    if (@($evidenceRows).Count -eq 0) {
        $lines.Add('| (none) |  |  |  |  |  |  |')
    }

    return ($lines -join "`n")
}

function ConvertTo-RSMarkdownTableCell {
    param(
        [AllowNull()]
        [object]$Value = $null
    )

    if ($null -eq $Value) {
        return ''
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    return $text.Replace('|', '\|').Replace("`r`n", '<br>').Replace("`n", '<br>').Replace("`r", '<br>')
}

function Convert-RSTestRunSummaryToGitHubStepSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Summary
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $schema = [string](Get-RSObjectValue -Object $Summary -Names @('schema') -DefaultValue '')
    if ($schema -eq 'red-salamander.run-all-tests.v2') {
        $lines.Add('## RedSalamander Validation Summary')
        $lines.Add('')
        $lines.Add("- Repository verdict: **$($Summary.repository_verdict)**")
        $lines.Add("- Change verdict: **$($Summary.change_verdict)**")
        $lines.Add("- Gate/mode: ``$($Summary.requested_gate)`` / ``$($Summary.validation_mode)``")
        $lines.Add("- Coverage: ``$($Summary.coverage_scope)``")
        $lines.Add(("- Entries: executed={0} reused={1} promoted={2} provisional={3} invalidated={4} incomplete={5}" -f `
                    $Summary.counts.executed, $Summary.counts.reused, $Summary.counts.promoted, `
                    $Summary.counts.provisional, $Summary.counts.invalidated, $Summary.counts.incomplete))
        $lines.Add(("- Cases: passed={0} failed={1} skipped={2} total={3}" -f `
                    $Summary.case_counts.passed, $Summary.case_counts.failed, $Summary.case_counts.skipped, $Summary.case_counts.total))
        $lines.Add("- Estimated avoided duration: $($Summary.estimated_avoided_duration_ms) ms")
        $lines.Add('')
        $lines.Add('| Entry | Execution | Promotion | Origin | Reasons |')
        $lines.Add('|---|---:|---:|---|---|')
        foreach ($entry in $Summary.entries) {
            $lines.Add(('| {0} | {1} | {2} | {3} | {4} |' -f `
                        (ConvertTo-RSMarkdownTableCell $entry.id), $entry.execution, $entry.promotion_state, `
                        (ConvertTo-RSMarkdownTableCell $entry.origin_run_id), `
                        (ConvertTo-RSMarkdownTableCell (@($entry.reason_codes) -join ', '))))
        }
        return (($lines.ToArray()) -join [Environment]::NewLine) + [Environment]::NewLine
    }
    $exitCode = [int](Get-RSObjectValue -Object $Summary -Names @('exit_code', 'ExitCode') -DefaultValue 1)
    $status = if ($exitCode -eq 0) { 'PASSED' } else { 'FAILED' }
    $suite = [string](Get-RSObjectValue -Object $Summary -Names @('suite', 'Suite') -DefaultValue '')
    $runId = [string](Get-RSObjectValue -Object $Summary -Names @('run_id', 'RunId') -DefaultValue '')
    $markdownTick = [char]96

    $lines.Add('## RedSalamander Test Summary')
    $lines.Add('')
    $lines.Add("- Status: **$status**")
    $lines.Add("- Suite: $markdownTick$suite$markdownTick")
    if (-not [string]::IsNullOrWhiteSpace($runId)) {
        $lines.Add("- Run id: $markdownTick$runId$markdownTick")
    }
    $lines.Add(("- Cases: passed={0} failed={1} skipped={2} total={3}" -f `
                (Get-RSObjectValue -Object $Summary -Names @('passed') -DefaultValue 0), `
                (Get-RSObjectValue -Object $Summary -Names @('failed') -DefaultValue 0), `
                (Get-RSObjectValue -Object $Summary -Names @('skipped') -DefaultValue 0), `
                (Get-RSObjectValue -Object $Summary -Names @('total') -DefaultValue 0)))

    $classifications = Get-RSObjectValue -Object $Summary -Names @('classifications') -DefaultValue $null
    if ($null -ne $classifications) {
        $lines.Add(("- Classifications: passed={0} flaky={1} regression={2} isolation_suspect={3} unclassified_failure={4}" -f `
                    (Get-RSObjectValue -Object $classifications -Names @('passed') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $classifications -Names @('flaky') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $classifications -Names @('regression') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $classifications -Names @('isolation_suspect') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $classifications -Names @('unclassified_failure') -DefaultValue 0)))
    }

    $lines.Add('')
    $lines.Add('### Suites')
    $lines.Add('')
    $lines.Add('| Suite | Classification | Passed | Failed | Skipped | Retry attempts |')
    $lines.Add('|---|---:|---:|---:|---:|---:|')
    foreach ($summarySuite in @(Get-RSObjectValue -Object $Summary -Names @('suites') -DefaultValue @())) {
        $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} |' -f `
                    (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $summarySuite -Names @('suite') -DefaultValue '')), `
                    (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $summarySuite -Names @('classification') -DefaultValue '')), `
                    (Get-RSObjectValue -Object $summarySuite -Names @('passed') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $summarySuite -Names @('failed') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $summarySuite -Names @('skipped') -DefaultValue 0), `
                    @((Get-RSObjectValue -Object $summarySuite -Names @('retry_attempts') -DefaultValue @())).Count))
    }

    $quarantine = Get-RSObjectValue -Object $Summary -Names @('quarantine') -DefaultValue $null
    if ($null -ne $quarantine) {
        $lines.Add('')
        $lines.Add('### Quarantine')
        $lines.Add('')
        $lines.Add(("- Active: {0}; invalid: {1}; repair_attempts: {2}; repair_reproduced: {3}" -f `
                    (Get-RSObjectValue -Object $quarantine -Names @('active_count') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $quarantine -Names @('invalid_count') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $quarantine -Names @('repair_attempt_count') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $quarantine -Names @('repair_reproduced_count') -DefaultValue 0)))

        $activeEntries = @(Get-RSObjectValue -Object $quarantine -Names @('active_entries') -DefaultValue @())
        if (@($activeEntries).Count -gt 0) {
            $lines.Add('')
            $lines.Add('| Active harness | Case | Owner | Expires | Issue |')
            $lines.Add('|---|---|---|---|---|')
            foreach ($entry in $activeEntries) {
                $lines.Add(('| {0} | {1} | {2} | {3} | {4} |' -f `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('harness') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('name') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('owner') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('expires') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('issue') -DefaultValue ''))))
            }
        }

        $repairAttempts = @(Get-RSObjectValue -Object $quarantine -Names @('repair_attempts') -DefaultValue @())
        if (@($repairAttempts).Count -gt 0) {
            $lines.Add('')
            $lines.Add('| Repair status | Harness | Case | Owner | Expires | Reason |')
            $lines.Add('|---|---|---|---|---|---|')
            foreach ($attempt in $repairAttempts) {
                $repairStatus = if ([bool](Get-RSObjectValue -Object $attempt -Names @('passed') -DefaultValue $false)) { 'REPAIR PASS' } else { 'REPAIR FAIL' }
                $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} |' -f `
                            $repairStatus, `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('harness') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('name') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('owner') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('expires') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('reason') -DefaultValue ''))))
            }
        }

        $errors = @(Get-RSObjectValue -Object $quarantine -Names @('errors') -DefaultValue @())
        if (@($errors).Count -gt 0) {
            $lines.Add('')
            $lines.Add('Invalid quarantine entries:')
            foreach ($errorMessage in $errors) {
                $lines.Add("- $(ConvertTo-RSMarkdownTableCell $errorMessage)")
            }
        }
    }

    return (($lines.ToArray()) -join [Environment]::NewLine) + [Environment]::NewLine
}

Export-ModuleMember -Function @(
    'Get-RSValidationExitCode',
    'Convert-RSTestCaseForRunSummary',
    'Convert-RSTestRetryAttemptForRunSummary',
    'Convert-RSTestQuarantineRepairAttemptForRunSummary',
    'New-RSTestRunSummary',
    'New-RSValidationRunSummaryV2',
    'Convert-RSTestRunSummaryToCaseHistoryRows',
    'Convert-RSTestCaseHistoryRowsToJsonl',
    'Convert-RSTestCaseHistoryRowsToDashboardMarkdown',
    'Convert-RSTestRunSummaryToGitHubStepSummary'
)
