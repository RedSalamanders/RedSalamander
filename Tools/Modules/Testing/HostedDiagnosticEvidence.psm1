Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'ValidationEvidence.psm1') -Scope Local
Import-Module (Join-Path $PSScriptRoot 'TestSandbox.psm1') -Scope Local

# This adapter-only check deliberately accepts filtered runner exit 2, unlike the
# general terminal reducer, but only after that reducer validates every native case.
# It never converts the repository's NOT_EVALUATED coverage into a Full result.
function Get-RSHostedDiagnosticSelection {
    param([Parameter(Mandatory)][ValidateSet('CommandsHandoff','FileOpsVolumes')][string]$Scope)
    switch ($Scope) {
        'CommandsHandoff' { return @{ suite='Commands'; case='cmd_pane_selection_select_next_keeps_navigation_shell_stable'; native='commands'; nativeSuite='Commands'; entry='selftest.commands' } }
        'FileOpsVolumes' { return @{ suite='FileOps'; case='R4A19_DiscoveryIndependentVolumes'; native='fileops'; nativeSuite='FileOperations'; entry='selftest.file-operations' } }
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
    $known = @((Get-RSHostedDiagnosticSelection CommandsHandoff), (Get-RSHostedDiagnosticSelection FileOpsVolumes))
    $matches = @($known | Where-Object {
        $_.suite -ceq $Selection.suite -and $_.case -ceq $Selection.case -and $_.native -ceq $Selection.native -and
        $_.nativeSuite -ceq $Selection.nativeSuite -and $_.entry -ceq $Selection.entry
    })
    if ($matches.Count -ne 1) { throw 'Unknown hosted diagnostic selection.' }
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
    if ($Summary.executable -isnot [string] -or $Summary.executable -ine $ExpectedExePath -or
        $InvocationStartedUtc -gt $InvocationEndedUtc) { throw 'Diagnostic executable or invocation window does not match.' }
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
    $q = $Summary.quarantine
    if ($q -isnot [System.Collections.IDictionary] -or -not $q.Contains('is_valid') -or -not $q.Contains('has_blocking_entries') -or
        -not $q.Contains('active_count') -or -not $q.Contains('active_entries') -or -not $q.Contains('invalid_count') -or
        -not $q.Contains('invalid_entries') -or -not $q.Contains('errors') -or $q.is_valid -isnot [bool] -or
        $q.has_blocking_entries -isnot [bool] -or $q.is_valid -cne $true -or
        $q.has_blocking_entries -cne $false -or [int]$q.active_count -ne 0 -or [int]$q.invalid_count -ne 0 -or
        @($q.active_entries).Count -ne 0 -or @($q.invalid_entries).Count -ne 0 -or @($q.errors).Count -ne 0) {
        throw 'Missing, invalid, or blocking quarantine state.'
    }
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

function Read-RSHostedDiagnosticEvidence {
    param([Parameter(Mandatory)][string]$TestRoot, [Parameter(Mandatory)][string]$RunId,
          [Parameter(Mandatory)][System.Collections.IDictionary]$Selection)
    Assert-RSHostedDiagnosticSelection $Selection
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
        $complete.intent_sha256 -ine (Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash) { throw 'Diagnostic archive proof is invalid.' }
    $relative = 'artifacts\selftest\last_run\' + $Selection.native + '\results.json'
    $nativePath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $run $relative) -RequireDescendant
    $rows = @($intent.inventory | Where-Object { $_.path -ieq $relative -and -not $_.directory })
    if ($rows.Count -ne 1 -or [int64]$rows[0].length -ne (Get-Item -LiteralPath $nativePath).Length -or
        $rows[0].sha256 -ine (Get-FileHash -LiteralPath $nativePath -Algorithm SHA256).Hash) { throw 'Native diagnostic result does not match the archive.' }
    $summaryPath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $run 'artifacts/selftest/last_run/run-all-tests-results.json') -RequireDescendant
    return @{ summary=(Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json -AsHashtable);
        native=(Get-Content -LiteralPath $nativePath -Raw | ConvertFrom-Json -AsHashtable);
        summarySha256=(Get-FileHash -LiteralPath $summaryPath -Algorithm SHA256).Hash;
        nativeSha256=(Get-FileHash -LiteralPath $nativePath -Algorithm SHA256).Hash;
        archiveIntentSha256=$complete.intent_sha256 }
}

Export-ModuleMember -Function Get-RSHostedDiagnosticSelection, Get-RSHostedDiagnosticRunId, Test-RSHostedDiagnosticOutcome, Read-RSHostedDiagnosticEvidence
