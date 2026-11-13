Set-StrictMode -Version Latest
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
        # Current runners publish v2 at the stable path and retain v1 beside it.
        $summaryPath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $run 'artifacts/selftest/last_run/run-all-tests-results-v1.json') -RequireDescendant
    }
    return @{ summary=(Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json -AsHashtable);
        native=(Get-Content -LiteralPath $nativePath -Raw | ConvertFrom-Json -AsHashtable);
        summarySha256=(Get-FileHash -LiteralPath $summaryPath -Algorithm SHA256).Hash;
        nativeSha256=(Get-FileHash -LiteralPath $nativePath -Algorithm SHA256).Hash;
        archiveIntentSha256=$archive.ArchiveIntentSha256 }
}

function Assert-RSHostedFilteredCaseRun {
    <#
    .SYNOPSIS
        Verifies a Run-AllTests.ps1 -CaseFilter invocation from its archived evidence.
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
    $v2 = Read-RSValidationJsonFile -LiteralPath $v2Path -AllowedRoot $TestRoot `
        -SchemaPath (Join-Path $RepoRoot 'Specs/Testing/RunAllTestsSummary.schema.json')
    $v1 = Get-Content -LiteralPath $v1Path -Raw | ConvertFrom-Json -AsHashtable

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
        [int64]$v1.passed -ne $names.Count -or [int64]$v2.case_counts.total -ne $names.Count -or
        [int64]$v2.case_counts.passed -ne $names.Count) { throw 'Native filtered case counts do not match runner summaries.' }

    $buildOutput = Join-Path $RepoRoot ".build/$Platform/$Configuration"
    $receipt = Read-RSBuildReceipt -BuildOutputDir $buildOutput -SchemaPath (Join-Path $RepoRoot 'Specs/Build/BuildReceipt.schema.json')
    $snapshot = Get-RSBuildSourceSnapshot -RepoRoot $RepoRoot
    $requiredExe = ".build/$Platform/$Configuration/RedSalamander.exe"
    if ($v1.executable -ine [IO.Path]::GetFullPath((Join-Path $RepoRoot $requiredExe)) -or
        $null -eq $receipt -or $v2.snapshot_id -cne $snapshot.SnapshotId -or
        $v2.build_receipt_id -cne $receipt.receipt_id -or
        -not (Test-RSBuildReceiptForArtifactUse -Receipt $receipt -SourceSnapshot $snapshot -Platform $Platform `
            -Configuration $Configuration -RepoRoot $RepoRoot -RequiredRelativePath @($requiredExe) `
            -RequiredProjectSelection @('RedSalamander/RedSalamander.vcxproj'))) {
        throw 'Native filtered source, executable, or artifact receipt mismatch.'
    }

    return [pscustomobject]@{
        RunId = $runId
        RepositoryVerdict = [string]$v2.repository_verdict
        ChangeVerdict = [string]$v2.change_verdict
        Passed = [int64]$v2.case_counts.passed
        ArchiveIntentSha256 = $archive.ArchiveIntentSha256
    }
}

Export-ModuleMember -Function Get-RSHostedDiagnosticSelection, Get-RSHostedDiagnosticRunId, Test-RSHostedDiagnosticOutcome, Read-RSHostedDiagnosticEvidence, Assert-RSHostedFilteredCaseRun, Get-RSHostedPairedDispatch
