# Root evidence-only curation of the two closed, reviewed F7 inputs. No native/build/checkout operation.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$taskBase = 'C:/RedSalamander.Perf/evidence/i27-followup-20261002-0e139fb/resume-20261003T121349Z'
$taskOutput = Join-Path $taskBase ('f7-curation-external-' + [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ'))
if (Test-Path -LiteralPath $taskOutput) { throw 'Existing curation preserved.' }
[void](New-Item -ItemType Directory -Path $taskOutput)
$taskInputs = @(
    @{label='baseline';run='20261003T182209Z-51656-126e7ed59f724fefae9d270d755c8153';packet='f7-complete-original-buffers-idle-agreed-20261003T182209Z-51656-126e7ed59f724fefae9d270d755c8153'},
    @{label='candidate';run='20261003T191317Z-82816-199e9f25af4d49739d1410c6881e3e0d';packet='f7-complete-stack-buffers-idle-comparison-20261003T191317Z-82816-199e9f25af4d49739d1410c6881e3e0d'}
)
$taskUtf8 = [Text.UTF8Encoding]::new($false, $true)
$taskSummary = [Collections.Generic.List[object]]::new()
foreach ($taskInput in $taskInputs) {
    $taskTarget = Join-Path $taskOutput $taskInput.label
    [void](New-Item -ItemType Directory -Path $taskTarget)
    $taskManifestPath = Join-Path $taskBase ($taskInput.packet + '/C-archive/archive-manifest.json')
    $taskManifest = Get-Content -LiteralPath $taskManifestPath -Raw | ConvertFrom-Json -DateKind String
    $taskWhole = Join-Path $taskBase ($taskInput.packet + '/C-archive/whole-run')
    $taskSource = Join-Path $taskWhole 'artifacts/selftest/last_run/perf/perf_metrics.jsonl'
    $taskReportPath = Join-Path $taskWhole 'artifacts/selftest/last_run/run-all-tests-results.json'
    $taskReport = Get-Content -LiteralPath $taskReportPath -Raw | ConvertFrom-Json -DateKind String
    if ($taskManifest.schema -cne 'i27.v2-pair-whole-run-archive.v1' -or
        [IO.Path]::GetFullPath($taskManifest.source) -cne [IO.Path]::GetFullPath(('C:/RedSalamander.Perf/runs/' + $taskInput.run)) -or
        [IO.Path]::GetFullPath($taskManifest.whole_run) -cne [IO.Path]::GetFullPath($taskWhole) -or
        $taskManifest.complete -isnot [bool] -or -not $taskManifest.complete -or $null -ne $taskManifest.error -or
        $taskReport.run_id -cne $taskInput.run -or $taskReport.case_filter -cne 'FileOps_AdmissionBaseline' -or
        $taskReport.passed -ne 3 -or $taskReport.failed -ne 0 -or $taskReport.skipped -ne 0 -or $taskReport.exit_code -ne 0) {
        throw 'Expected exact closed successful F7 run.'
    }
    foreach ($taskPath in @($taskManifestPath,$taskSource,$taskReportPath)) {
        $taskCursor = Get-Item -LiteralPath $taskPath -Force
        if ($taskCursor -isnot [IO.FileInfo]) { throw 'Expected ordinary retained file.' }
        while ($null -ne $taskCursor) {
            if ($taskCursor.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Retained input has a reparse component.' }
            $taskCursor = if ($taskCursor -is [IO.FileInfo]) { $taskCursor.Directory } else { $taskCursor.Parent }
        }
    }
    $taskSourceRelative = [IO.Path]::GetRelativePath($taskWhole,$taskSource)
    $taskBound = @($taskManifest.copied | Where-Object path -ceq $taskSourceRelative)
    $taskSourceSha = (Get-FileHash -LiteralPath $taskSource -Algorithm SHA256).Hash.ToLowerInvariant()
    $taskSourceBytes = (Get-Item -LiteralPath $taskSource -Force).Length
    if ($taskBound.Count -ne 1 -or $taskBound[0].sha256 -cne $taskSourceSha -or $taskBound[0].bytes -ne $taskSourceBytes) {
        throw 'Original metric bytes differ from their closed manifest.'
    }
    $taskReportBound = @($taskManifest.copied | Where-Object path -ceq ([IO.Path]::GetRelativePath($taskWhole,$taskReportPath)))
    if ($taskReportBound.Count -ne 1 -or $taskReportBound[0].sha256 -cne (Get-FileHash -LiteralPath $taskReportPath -Algorithm SHA256).Hash.ToLowerInvariant() -or
        $taskReportBound[0].bytes -ne (Get-Item -LiteralPath $taskReportPath -Force).Length) { throw 'Original runner report differs from its closed manifest.' }
    # The native sink writes UTF-8, without BOM, and LF. Check that before reproducing selected raw lines.
    $taskProbe = [IO.File]::OpenRead($taskSource)
    try {
        $taskFirst = [byte[]]::new(4096)
        $taskFirstCount = $taskProbe.Read($taskFirst,0,$taskFirst.Length)
        if ($taskFirstCount -ge 3 -and $taskFirst[0] -eq 239 -and $taskFirst[1] -eq 187 -and $taskFirst[2] -eq 191) { throw 'Unexpected original BOM.' }
        $taskFirstLf = [Array]::IndexOf($taskFirst,[byte]10,0,$taskFirstCount)
        if ($taskFirstLf -lt 0 -or ($taskFirstLf -gt 0 -and $taskFirst[$taskFirstLf-1] -eq 13)) { throw 'Expected original native LF delimiter.' }
    } finally { $taskProbe.Dispose() }
    $taskReader = [IO.StreamReader]::new($taskSource,$taskUtf8,$false,65536)
    $taskWriter = $null
    $taskSegments = [Collections.Generic.List[object]]::new()
    $taskLineNumbers = [Collections.Generic.List[long]]::new()
    $taskRowCount = 0L; $taskOriginalLines = 0L; $taskSelectedBytes = 0L
    $taskMetricRun = $null; $taskMetricProcess = $null
    $taskSegment = $null; $taskSegmentOrdinal = 0; $taskSegmentBytes = 0L
    try {
        while ($null -ne ($taskLine = $taskReader.ReadLine())) {
            ++$taskOriginalLines
            if ($taskLine.IndexOf('FileOps.AdmissionBaseline.',[StringComparison]::Ordinal) -lt 0) { continue }
            $taskRow = $taskLine | ConvertFrom-Json -DateKind String
            if (-not $taskRow.metric.StartsWith('FileOps.AdmissionBaseline.',[StringComparison]::Ordinal)) { continue }
            if ($taskRow.machineHash -cne '4cb089111a23' -or $taskRow.build -cne 'Release' -or
                $taskRow.branch -cne 'fileops-safety-followup' -or $taskRow.commit -cne '' -or
                $taskRow.hr -ne $(if ($taskRow.metric -ceq 'FileOps.AdmissionBaseline.RouteSkipped') {1} else {0})) {
                throw 'Selected original metric profile or HRESULT differs.'
            }
            if ($null -eq $taskMetricRun) { $taskMetricRun = $taskRow.runId; $taskMetricProcess = $taskRow.process }
            elseif ($taskRow.runId -cne $taskMetricRun -or $taskRow.process -ne $taskMetricProcess) { throw 'Selected native metric process/run identity changed.' }
            $taskBytes = $taskUtf8.GetByteCount($taskLine) + 1
            if ($taskBytes -gt 2MB) { throw 'Original claim row exceeds archive file cap.' }
            if ($null -eq $taskWriter -or $taskSegmentBytes + $taskBytes -gt 2MB) {
                if ($null -ne $taskWriter) { $taskWriter.Dispose(); $taskWriter = $null }
                ++$taskSegmentOrdinal
                $taskSegmentFolder = Join-Path $taskTarget ('raw-{0:D2}' -f $taskSegmentOrdinal)
                [void](New-Item -ItemType Directory -Path $taskSegmentFolder)
                $taskSegmentPath = Join-Path $taskSegmentFolder 'perf_metrics.jsonl'
                $taskWriter = [IO.StreamWriter]::new($taskSegmentPath,$false,$taskUtf8,65536)
                $taskWriter.NewLine = "`n"
                $taskSegmentBytes = 0L
                $taskSegment = [pscustomobject]@{path=$taskSegmentPath;rows=0L;bytes=0L;first_original_line=$taskOriginalLines;last_original_line=$taskOriginalLines;sha256=$null}
                $taskSegments.Add($taskSegment)
            }
            $taskWriter.WriteLine($taskLine)
            $taskSegment.bytes += $taskBytes; ++$taskSegment.rows; $taskSegment.last_original_line = $taskOriginalLines
            $taskSegmentBytes += $taskBytes; $taskSelectedBytes += $taskBytes; ++$taskRowCount
            $taskLineNumbers.Add($taskOriginalLines)
        }
    } finally {
        if ($null -ne $taskWriter) { $taskWriter.Dispose() }
        $taskReader.Dispose()
    }
    foreach ($taskSegment in $taskSegments) {
        if ((Get-Item -LiteralPath $taskSegment.path).Length -ne $taskSegment.bytes) { throw 'Selected raw segment size differs.' }
        $taskSegment.sha256 = (Get-FileHash -LiteralPath $taskSegment.path -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    if ($taskRowCount -ne 5694 -or $taskSourceSha -cne (Get-FileHash -LiteralPath $taskSource -Algorithm SHA256).Hash.ToLowerInvariant()) {
        throw 'Complete original row selection or unchanged source binding failed.'
    }
    $taskRecord = [pscustomobject]@{
        schema='i27.f7-original-row-selection.v1';variant=$taskInput.label;run_id=$taskInput.run
        original_path=$taskSource;original_sha256=$taskSourceSha;original_bytes=$taskSourceBytes;original_line_count=$taskOriginalLines
        archive_manifest=$taskManifestPath;archive_manifest_sha256=(Get-FileHash -LiteralPath $taskManifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
        selected_rows=$taskRowCount;selected_bytes=$taskSelectedBytes;native_metric_run_id=$taskMetricRun;native_metric_process=$taskMetricProcess;original_line_numbers=@($taskLineNumbers.ToArray());segments=@($taskSegments.ToArray())
        selector='All original rows whose metric begins FileOps.AdmissionBaseline.; includes all qualified, prime, control and S_FALSE skip rows in original order.'
        byte_rule='Unchanged original UTF-8 row text plus original native LF delimiter; no JSON rewriting, numeric conversion, sampling, adverse-row removal or separate-run folder splitting.'
        native_commit_exception='Original empty commit field retained; reviewed canonical receipt/source binding remains separately required.'
        curator=$PSCommandPath;curator_sha256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant()
        qualification='Evidence curation only; original failures/skips, scoped assessment and all other I27 gates remain unchanged.'
    }
    [IO.File]::WriteAllText((Join-Path $taskTarget 'row-selection.json'),($taskRecord|ConvertTo-Json -Depth 10),$taskUtf8)
    $taskSummary.Add([pscustomobject]@{variant=$taskInput.label;rows=$taskRowCount;bytes=$taskSelectedBytes;segments=$taskSegments.Count;original_bytes=$taskSourceBytes;original_lines=$taskOriginalLines;folder=$taskTarget})
}
[IO.File]::WriteAllText((Join-Path $taskOutput 'summary.json'),($taskSummary.ToArray()|ConvertTo-Json -Depth 6),$taskUtf8)
$taskSummary.ToArray() | ConvertTo-Json -Depth 6
