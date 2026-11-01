param([string]$EvidenceRoot = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
$taskUtf8 = [Text.UTF8Encoding]::new($false)

$profiles = @{}
foreach ($variant in @('baseline','candidate')) {
    $profile = Get-Content -LiteralPath (Join-Path $EvidenceRoot "$variant-build-receipt.json") -Raw | ConvertFrom-Json
    if ($profile.configuration -ne 'Release' -or $profile.platform -ne 'x64' -or -not $profile.tests_enabled -or
        'compiler-debug-information=embedded' -notin $profile.build_arguments -or $profile.exit_code -ne 0) { throw "Invalid $variant build profile." }
    $profiles[$variant] = $profile
}
foreach ($field in @('toolchain_identity_digest','vcpkg_identity_digest')) {
    if ($profiles.baseline.$field -cne $profiles.candidate.$field) { throw "Build identity mismatch: $field" }
}
if (($profiles.baseline.build_arguments -join '|') -cne ($profiles.candidate.build_arguments -join '|')) { throw 'Build arguments differ.' }
$taskBaselineBinding = Get-Content -LiteralPath (Join-Path $EvidenceRoot 'baseline-source-binding.json') -Raw | ConvertFrom-Json
$taskCodeBindingPath = Join-Path $EvidenceRoot 'isolation-source-binding-final.json'
$taskCodeBinding = Get-Content -LiteralPath $taskCodeBindingPath -Raw | ConvertFrom-Json
$taskRuntimeBinding = Get-Content -LiteralPath (Join-Path $EvidenceRoot 'candidate-perf-build-binding.json') -Raw | ConvertFrom-Json
if ($taskRuntimeBinding.schema -ne 'red-salamander.i27-perf-build-binding.v2' -or
    $taskRuntimeBinding.receipt_id -ne $profiles.candidate.receipt_id -or
    $taskRuntimeBinding.source_snapshot_id -ne $profiles.candidate.source_snapshot_id -or
    $taskRuntimeBinding.code_binding_sha256 -ne (Get-FileHash -LiteralPath $taskCodeBindingPath -Algorithm SHA256).Hash.ToLowerInvariant()) {
    throw 'Final candidate source/runtime binding is stale.'
}
$taskExpectedRoots = @{baseline=$taskBaselineBinding.frozen_checkout; candidate=$taskCodeBinding.isolated}
$taskCaptureBindings = @(Get-Content -LiteralPath (Join-Path $EvidenceRoot 'candidate-capture-build-bindings.json') -Raw | ConvertFrom-Json)
if ($taskCaptureBindings.Count -ne 5) { throw 'Candidate capture build bindings are incomplete.' }

function Get-Median([double[]]$Values) {
    if ($Values.Count -eq 0) { throw 'A metric has no samples.' }
    $sorted = @($Values | Sort-Object)
    $middle = [int][Math]::Floor($sorted.Count / 2)
    if (($sorted.Count % 2) -ne 0) { return $sorted[$middle] }
    return ($sorted[$middle - 1] + $sorted[$middle]) / 2.0
}

$captures = @()
foreach ($variant in @('baseline', 'candidate')) {
    foreach ($index in 1..5) {
        $root = Join-Path $EvidenceRoot "$variant-$index/artifacts/selftest/last_run"
        $report = Get-Content -LiteralPath (Join-Path $root 'run-all-tests-results.json') -Raw | ConvertFrom-Json
        $native = Get-Content -LiteralPath (Join-Path $root 'fileops/results.json') -Raw | ConvertFrom-Json
        if ($variant -eq 'candidate') {
            $taskCaptureBinding = @($taskCaptureBindings | Where-Object capture -eq $index)
            if ($taskCaptureBinding.Count -ne 1 -or $taskCaptureBinding[0].run_id -ne $report.run_id -or
                $taskCaptureBinding[0].receipt_id -ne $profiles.candidate.receipt_id -or
                $taskCaptureBinding[0].source_snapshot_id -ne $profiles.candidate.source_snapshot_id -or
                $taskCaptureBinding[0].runtime_binding_sha256 -ne (Get-FileHash -LiteralPath (Join-Path $EvidenceRoot 'candidate-perf-build-binding.json') -Algorithm SHA256).Hash.ToLowerInvariant() -or
                $taskCaptureBinding[0].modules_checked_before_and_after -ne $taskRuntimeBinding.modules.Count) {
                throw "Candidate capture runtime identity is not bound: $index"
            }
        }
        $taskExpectedExe = [IO.Path]::GetFullPath((Join-Path $taskExpectedRoots[$variant] '.build/x64/Release/RedSalamander.exe'))
        if ([IO.Path]::GetFullPath($report.executable) -ne $taskExpectedExe) { throw "Wrong executable checkout: $variant/$index" }
        if ($report.suite -ne 'FileOps' -or $report.platform -ne 'x64' -or $report.configuration -ne 'Release' -or
            $report.case_filter -ne 'FileOps_IntegratedPresentationPerf' -or $report.exit_code -ne 0 -or
            $native.passed -ne 1 -or $native.cases.Count -ne 1 -or $native.failed -ne 0 -or $native.skipped -ne 0 -or
            @($native.cases | Where-Object { $_.name -eq 'FileOps_IntegratedPresentationPerf' -and $_.status -eq 'passed' }).Count -ne 1 -or
            -not $report.test_sandbox_audit.is_clean) { throw "Invalid $variant capture $index." }
        $metricPath = Join-Path $root 'perf/perf_metrics.jsonl'
        $allLines = @(Get-Content -LiteralPath $metricPath)
        $selectedLines = @()
        $rows = @()
        foreach ($line in $allLines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $row = $line | ConvertFrom-Json
            if ($row.metric -like 'FileOps.IntegratedPresentation.*') { $selectedLines += $line; $rows += $row }
        }
        $machines = @($rows.machineHash | Sort-Object -Unique)
        if ($machines.Count -ne 1 -or $machines[0] -ne '4cb089111a23') { throw 'Machine control changed.' }
        if ($rows.Count -ne 275 -or @($rows | Where-Object { $_.build -cne 'Release' -or $_.commit -cne $profiles[$variant].git_head }).Count -ne 0 -or
            @($rows.runId | Sort-Object -Unique).Count -ne 1 -or @($rows.process | Sort-Object -Unique).Count -ne 1) { throw 'Metric provenance changed.' }
        $runStart = [DateTimeOffset]::Parse($report.run_started_utc)
        $runEnd = [DateTimeOffset]::Parse($report.run_ended_utc)
        foreach ($row in $rows) {
            $timestamp = [DateTimeOffset]::Parse($row.timestamp)
            if ($timestamp -lt $runStart -or $timestamp -gt $runEnd) { throw 'Metric timestamp falls outside its runner capture.' }
        }
        $controls = @($rows | Where-Object metric -in @(
            'FileOps.IntegratedPresentation.WorkloadFNV64','FileOps.IntegratedPresentation.ScanStatusWorkloadFNV64',
            'FileOps.IntegratedPresentation.ClientPixels','FileOps.IntegratedPresentation.ScanStatusClientPixels') |
            Select-Object metric,detail,durationUs,value0,value1)
        if ($controls.Count -ne 8 -or @($controls | Group-Object metric,detail | Where-Object Count -ne 1).Count -ne 0 -or
            @($rows | Where-Object { $_.metric -ne 'FileOps.IntegratedPresentation.SurfaceCounters' -and $_.hr -ne 0 }).Count -ne 0) { throw 'Invalid or duplicate control/measurement record.' }
        $captures += [pscustomobject]@{
            variant=$variant; capture=$index; run_id=$report.run_id; native=$native; rows=$rows; controls=$controls
            opaque_digest_tokens=@($selectedLines | Where-Object { ($_ | ConvertFrom-Json).metric -like '*WorkloadFNV64' } | ForEach-Object {
                $taskDigestToken=[regex]::Match($_, '"durationUs"\s*:\s*([0-9]+)')
                if (-not $taskDigestToken.Success) { throw 'Opaque workload digest token is missing.' }
                $taskDigestToken.Groups[1].Value
            })
            raw_sha256=(Get-FileHash -LiteralPath $metricPath -Algorithm SHA256).Hash.ToLowerInvariant()
            raw_bytes=(Get-Item -LiteralPath $metricPath).Length; raw_rows=$allLines.Count; claim_rows=$selectedLines.Count
            claim_lines=$selectedLines
        }
    }
}
if (@($captures.run_id | Sort-Object -Unique).Count -ne 10) { throw 'A capture run was reused.' }
$reference = $captures[0].controls | ConvertTo-Json -Depth 5 -Compress
foreach ($capture in $captures) {
    if (($capture.controls | ConvertTo-Json -Depth 5 -Compress) -cne $reference) { throw 'Fixture, DPI or geometry control changed.' }
    if ($capture.opaque_digest_tokens.Count -ne 2 -or ($capture.opaque_digest_tokens -join '|') -cne ($captures[0].opaque_digest_tokens -join '|')) { throw 'Exact workload digest token changed.' }
    $referencePhases = @($captures[0].rows | Where-Object metric -eq 'FileOps.IntegratedPresentation.Memory' | ForEach-Object detail)
    $capturePhases = @($capture.rows | Where-Object metric -eq 'FileOps.IntegratedPresentation.Memory' | ForEach-Object detail)
    if (($referencePhases -join '|') -cne ($capturePhases -join '|')) { throw 'Resource phases changed.' }
}
$comparisons = @()
$metrics = @('GraphOpenRedraw8Us','GraphCollapsedRedraw8Us','ChurnUpdate8Us','IssuesDetailRedraw8Us',
    'ScanStatusChangingRedraw8Us','ScanStatusCleanRedraw8Us')
foreach ($metric in $metrics) {
    $details = if ($metric -like 'ScanStatus*') { @('compare','change-case','change-attributes','file-list') } else { @('') }
    foreach ($detail in $details) {
    $before = @(); $after = @(); $samples = @()
    foreach ($capture in $captures) {
        $values = @($capture.rows | Where-Object { $_.metric -eq "FileOps.IntegratedPresentation.$metric" -and ($detail -eq '' -or $_.detail -eq $detail) } | ForEach-Object { [double]$_.durationUs })
        $expectedSamples = if ($metric -eq 'ChurnUpdate8Us') { 32 } else { 5 }
        if ($values.Count -ne $expectedSamples) { throw "Timing sample count differs: $metric/$detail $($values.Count)" }
        $median = Get-Median $values
        if ($capture.variant -eq 'baseline') { $before += $median } else { $after += $median }
        $samples += [pscustomobject]@{variant=$capture.variant;capture=$capture.capture;samples=$values.Count;median_us=$median}
    }
    $baseline = Get-Median $before; $candidate = Get-Median $after
    $comparisons += [pscustomobject]@{metric=$metric;detail=$detail;baseline_median_us=$baseline;candidate_median_us=$candidate;
        delta_percent=100.0*($candidate-$baseline)/$baseline;baseline_capture_medians=$before;candidate_capture_medians=$after;captures=$samples}
    }
}
$resources = @()
foreach ($capture in $captures) {
    $memory = @($capture.rows | Where-Object metric -eq 'FileOps.IntegratedPresentation.Memory')
    $cpu = @($capture.rows | Where-Object metric -eq 'FileOps.IntegratedPresentation.Cpu')
    $handles = @($capture.rows | Where-Object metric -eq 'FileOps.IntegratedPresentation.Handles')
    $gui = @($capture.rows | Where-Object metric -eq 'FileOps.IntegratedPresentation.GuiAndAttachments')
    foreach ($samples in @($memory,$cpu,$handles,$gui)) {
        if ($samples.Count -ne 29 -or ($samples.detail -join '|') -cne ($memory.detail -join '|')) { throw 'Resource observation count/order differs.' }
    }
    $cold = @($memory | Where-Object detail -eq 'cold')[0]
    $closed = @($memory | Where-Object detail -eq 'scan-file-list-closed')[0]
    $cpuCold = @($cpu | Where-Object detail -eq 'cold')[0]
    $cpuClosed = @($cpu | Where-Object detail -eq 'scan-file-list-closed')[0]
    $resources += [pscustomobject]@{variant=$capture.variant;capture=$capture.capture;
        peak_observed_private_bytes=($memory.value0 | Measure-Object -Maximum).Maximum;
        peak_observed_working_set_bytes=($memory.value1 | Measure-Object -Maximum).Maximum;
        cold_private_bytes=$cold.value0;closed_private_bytes=$closed.value0;closed_minus_cold_private_bytes=$closed.value0-$cold.value0;
        scenario_ui_cpu_us=$cpuClosed.value0-$cpuCold.value0;scenario_process_cpu_us=$cpuClosed.value1-$cpuCold.value1;
        peak_observed_handles=($handles.value0 | Measure-Object -Maximum).Maximum;
        peak_observed_gdi=($gui.value0 | Measure-Object -Maximum).Maximum;peak_observed_user=($gui.value1 | Measure-Object -Maximum).Maximum;
        closed_minus_cold_handles=$handles[-1].value0-$handles[0].value0;
        closed_minus_cold_hosts=$handles[-1].value1-$handles[0].value1;
        closed_minus_cold_gdi=$gui[-1].value0-$gui[0].value0;closed_minus_cold_user=$gui[-1].value1-$gui[0].value1;
        closed_minus_cold_attachments=$gui[-1].durationUs-$gui[0].durationUs;
        cold_handles=$handles[0].value0;closed_handles=$handles[-1].value0;
        cold_gui_and_attachments=$gui[0];closed_gui_and_attachments=$gui[-1]}
}
$resourceComparison = @()
foreach ($field in @('peak_observed_private_bytes','peak_observed_working_set_bytes','closed_minus_cold_private_bytes',
    'scenario_ui_cpu_us','scenario_process_cpu_us','peak_observed_handles','peak_observed_gdi','peak_observed_user',
    'closed_minus_cold_handles','closed_minus_cold_hosts','closed_minus_cold_gdi','closed_minus_cold_user','closed_minus_cold_attachments')) {
    $before = @($resources | Where-Object variant -eq baseline | ForEach-Object { [double]$_.$field })
    $after = @($resources | Where-Object variant -eq candidate | ForEach-Object { [double]$_.$field })
    $baseline = Get-Median $before; $candidate = Get-Median $after
    $resourceComparison += [pscustomobject]@{metric=$field;baseline_median=$baseline;candidate_median=$candidate;delta=$candidate-$baseline;
        delta_percent=if($baseline -gt 0 -and $field -notlike 'closed_minus_cold_*'){100.0*($candidate-$baseline)/$baseline}else{$null};baseline_captures=$before;candidate_captures=$after}
}
$result = [pscustomobject]@{schema='red-salamander.fileops-i27-paired-analysis.v2';method='Five fresh processes per variant; median of per-capture medians per timing/scenario; medians of matched observations for resources.';
    build_bindings=@(@($profiles.baseline,$profiles.candidate) | Select-Object receipt_id,source_snapshot_id,git_head,configuration,platform,build_arguments,toolchain_identity_digest,vcpkg_identity_digest);
    controls=$captures[0].controls;timings=$comparisons;resource_comparisons=$resourceComparison;resources=$resources;
    captures=@($captures | Select-Object variant,capture,run_id,raw_sha256,raw_bytes,raw_rows,claim_rows);
    limitations=@('Eight synchronous RedrawWindow calls per redraw batch; not monitor-presented FPS.',
        'Memory peaks are maxima of labelled observations, not continuous allocator peaks.',
        'CPU counters have OS sampling granularity and include the fixture message pump.',
        'Integrated hidden interval remains 350ms; longer hidden/minimized timer behavior has a separate native regression.',
        'WorkloadFNV64 durationUs contains an opaque digest; ClientPixels durationUs contains DPI, not elapsed time. These controls are compared for identity only.',
        'GuiAndAttachments durationUs contains a DxUi attachment count, not elapsed time. Closed-minus-cold resource deltas are signed counts/bytes and have no percentage.',
        'Driver allocation, queue depth and GPU peak bytes remain unavailable as recorded by the fixture.')}
$result | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'performance-comparison.json') -Encoding utf8
$taskTypedControls = @()
foreach ($taskLine in $captures[0].claim_lines) {
    $taskRow = $taskLine | ConvertFrom-Json
    if ($taskRow.metric -like '*WorkloadFNV64') {
        $taskDigestMatch = [regex]::Match($taskLine, '"durationUs"\s*:\s*([0-9]+)')
        if (-not $taskDigestMatch.Success) { throw 'Opaque workload digest token is missing.' }
        $taskTypedControls += [pscustomobject]@{metric=$taskRow.metric;workload=$taskRow.detail;
            digest_u64=$taskDigestMatch.Groups[1].Value;object_count=$taskRow.value0;row_or_operation_count=$taskRow.value1;
            raw_field_semantics='durationUs is an opaque digest in this legacy fixture'}
    } elseif ($taskRow.metric -like '*ClientPixels') {
        $taskTypedControls += [pscustomobject]@{metric=$taskRow.metric;surface=$taskRow.detail;
            dpi=$taskRow.durationUs;width_pixels=$taskRow.value0;height_pixels=$taskRow.value1;
            raw_field_semantics='durationUs is DPI in this legacy fixture'}
    }
}
if ($taskTypedControls.Count -ne 8) { throw 'Typed fixture controls are incomplete.' }
$taskTypedReport = [ordered]@{schema='red-salamander.fileops-i27-typed-controls.v2';
    comparison_sha256=(Get-FileHash -LiteralPath (Join-Path $EvidenceRoot 'performance-comparison.json') -Algorithm SHA256).Hash.ToLowerInvariant();
    controls=$taskTypedControls;captures=@($captures | Select-Object variant,capture,run_id,raw_sha256);
    attachment_semantics='GuiAndAttachments.durationUs is an attachment count, not elapsed time; all 29 labelled rows per capture are retained in claim-rows.jsonl.'}
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'performance-typed-controls.json'), ($taskTypedReport | ConvertTo-Json -Depth 8), $taskUtf8)
$comparisons | Select-Object metric,detail,baseline_median_us,candidate_median_us,delta_percent | Format-Table -AutoSize
$resourceComparison | Select-Object metric,baseline_median,candidate_median,delta_percent | Format-Table -AutoSize
