<#
.SYNOPSIS
    UNEXECUTED proposal: curate all fixed F6/v2 rows from sixteen exact closed captures.
.DESCRIPTION
    Derived from root-reviewed Curate-ClosedF7Rows.reviewed.ps1. Preflights all literal
    successful controls before streaming. Root must invoke only after source-frozen
    comparisons have fully closed. Copies original row text/native LF, never JSON values.
.PARAMETER ClosedComparisonsConfirmed
    Root's explicit assertion that all six F6 and ten v2 contained captures are closed.
.OUTPUTS
    External row-selection records, raw segments and summary/failure record; no publication.
.NOTES
    PowerShell 7.5+ is required for ConvertFrom-Json -DateKind String. No native/build,
    source/output reads, process/focus operation, cleanup, source rewrite or qualification.
#>
[CmdletBinding()]
param([Parameter(Mandatory)][switch]$ClosedComparisonsConfirmed)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if (-not $ClosedComparisonsConfirmed) {throw 'Root must confirm all comparisons are closed before curation.'}
if ($PSVersionTable.PSVersion -lt [version]'7.5') {throw 'PowerShell 7.5+ required for exact string date parsing.'}
$taskBase='C:/RedSalamander.Perf/evidence/i27-followup-20261002-0e139fb/resume-20261003T121349Z'
$taskModelPath=Join-Path $taskBase 'final-run-review/Curate-ClosedF6V2Rows.proposed.inputs.json'
$taskModelSha='1355b36d3120c925cd2284cdd66fa3397432922b9b7ad2f3c69bb8a86f289638'
$taskUtf8=[Text.UTF8Encoding]::new($false,$true)
function Get-TaskSha([string]$Path) {(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Assert-TaskFile([string]$Path) {
    $full=[IO.Path]::GetFullPath($Path);$root=[IO.Path]::GetFullPath($taskBase).TrimEnd('\')
    if (-not $full.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase) -or $full.Substring(2).Contains(':')) {throw 'Exact governed evidence path required.'}
    $cursor=Get-Item -LiteralPath $full -Force
    if ($cursor -isnot [IO.FileInfo]) {throw 'Expected exact ordinary evidence file.'}
    while ($null -ne $cursor) {
        if ($cursor.Attributes -band [IO.FileAttributes]::ReparsePoint) {throw 'Evidence has reparse component.'}
        $cursor=if ($cursor -is [IO.FileInfo]) {$cursor.Directory}else{$cursor.Parent}
    }
    return $full
}
function Read-TaskJson([string]$Path) {Get-Content -LiteralPath (Assert-TaskFile $Path) -Raw | ConvertFrom-Json -DateKind String}
function Get-TaskBoundFile($Manifest,[string]$Whole,[string]$Relative) {
    $bound=@($Manifest.copied | Where-Object {([string]$_.path -replace '\\','/') -ceq ($Relative -replace '\\','/')})
    $path=Assert-TaskFile (Join-Path $Whole $Relative)
    if ($bound.Count -ne 1 -or $bound[0].sha256 -cnotmatch '^[0-9a-f]{64}$' -or (Get-Item -LiteralPath $path).Length -ne $bound[0].bytes) {throw 'Exact archive copied-byte identity is missing/mismatched.'}
    return [pscustomobject]@{path=$path;bytes=[long]$bound[0].bytes;sha256=$bound[0].sha256}
}
function Assert-TaskCounts($Record,[int]$Expected) {
    if ($Record.passed -ne $Expected -or $Record.failed -ne 0 -or $Record.skipped -ne 0) {throw 'Only complete successful declared native captures may be curated; original failures remain failures.'}
}
$taskModelPath=Assert-TaskFile $taskModelPath
if ((Get-TaskSha $taskModelPath) -cne $taskModelSha) {throw 'Literal input proposal changed; re-review before curation.'}
$taskInputs=@(Read-TaskJson $taskModelPath)
if ($taskInputs.Count -ne 16 -or @($taskInputs | Where-Object family -CEQ F6).Count -ne 6 -or @($taskInputs | Where-Object family -CEQ V2).Count -ne 10) {throw 'Expected exactly six F6 plus ten v2 literal inputs.'}
$taskOutput=Join-Path $taskBase ('f6-v2-curation-external-'+[DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ')+'-'+[Guid]::NewGuid().ToString('N'))
if (Test-Path -LiteralPath $taskOutput) {throw 'Existing output preserved.'}
[void](New-Item -ItemType Directory -Path $taskOutput)
$taskSummary=[Collections.Generic.List[object]]::new();$taskPrepared=[Collections.Generic.List[object]]::new()
$taskFailure=$null
try {
    # Small exact metadata preflight for ALL sixteen captures; no raw scan before it passes.
    foreach ($taskInput in $taskInputs) {
        $taskPacket=Join-Path $taskBase $taskInput.packet_relative
        $taskCapture=Join-Path $taskBase $taskInput.capture_relative
        $taskManifestPath=Assert-TaskFile (Join-Path $taskCapture $taskInput.manifest_relative)
        $taskWhole=[IO.Path]::GetFullPath((Join-Path $taskCapture $taskInput.whole_relative))
        $taskManifest=Read-TaskJson $taskManifestPath
        $taskFamily=$taskInput.family;$taskExpectedRows=if ($taskFamily -ceq 'F6') {208}else{275}
        $taskCase=if ($taskFamily -ceq 'F6') {'FileOps_SafetyRemediationContracts'}else{'FileOps_IntegratedPresentationPerf'}
        $taskNativePassed=if ($taskFamily -ceq 'F6') {3}else{1}
        $taskExpectedSchema=if ($taskFamily -ceq 'F6') {'i27.f6-whole-run-archive.v1'}else{'i27.v2-pair-whole-run-archive.v1'}
        if ($taskManifest.schema -cne $taskExpectedSchema -or $taskManifest.complete -isnot [bool] -or -not $taskManifest.complete -or
            $null -ne $taskManifest.error -or [IO.Path]::GetFullPath($taskManifest.source) -cne [IO.Path]::GetFullPath(('C:/RedSalamander.Perf/runs/'+$taskInput.run_id)) -or
            [IO.Path]::GetFullPath($taskManifest.whole_run) -cne $taskWhole) {throw 'Exact closed C whole-archive manifest required.'}
        $taskBindingPath=Assert-TaskFile (Join-Path $taskPacket 'interleaved-run-bindings.json')
        $taskBindings=@(Read-TaskJson $taskBindingPath)
        $taskBinding=$taskBindings[[int]$taskInput.ordinal-1]
        $taskBindingRole=if ($taskFamily -ceq 'F6') {$taskBinding.mode}else{$taskBinding.variant}
        if ($taskBinding.run_id -cne $taskInput.run_id -or $taskBindingRole -cne $taskInput.role -or $taskBinding.pair -ne $taskInput.pair -or
            $taskBinding.receipt_id -cne $taskInput.receipt_id -or $taskBinding.immutable_receipt_sha256 -cne $taskInput.receipt_sha -or
            $taskBinding.exit_code -ne 0 -or $null -ne $taskBinding.runner_error) {throw 'Literal interleaved ordinal/run/role/receipt binding failed.'}
        if ($taskFamily -ceq 'F6') {
            $taskRoots=@($taskBinding.roots | Where-Object root -CEQ 'C:\RedSalamander.Perf')
            if ($taskRoots.Count -ne 1 -or -not $taskRoots[0].present -or $taskRoots[0].archived_run -cne $taskWhole -or $null -ne $taskRoots[0].archive_error) {throw 'F6 exact C root binding failed.'}
            $taskFreshPath=Assert-TaskFile (Join-Path $taskPacket 'fresh-build-binding.json');$taskFresh=Read-TaskJson $taskFreshPath
        } else {
            if ($taskBinding.archived_run -cne $taskWhole -or $null -ne $taskBinding.archive_error) {throw 'V2 exact C archive binding failed.'}
            $taskFreshPath=Assert-TaskFile (Join-Path $taskPacket 'fresh-build-bindings.json');$taskFresh=(Read-TaskJson $taskFreshPath).($taskInput.role)
        }
        $taskExpectedCheckout=if ($taskFamily -ceq 'V2' -and $taskInput.role -ceq 'baseline') {'C:\RsI26V2'}else{'C:\RsI27Followup'}
        if ($taskFresh.repo -cne $taskExpectedCheckout -or $taskFresh.profile -cne ($taskExpectedCheckout+'\.build\x64\Release') -or
            $taskFresh.receipt.receipt_id -cne $taskInput.receipt_id -or $taskFresh.record_sha256 -cne $taskInput.receipt_sha -or
            $taskFresh.source_snapshot_id -cne $taskInput.source_snapshot -or $taskFresh.receipt.source_snapshot_id -cne $taskInput.source_snapshot -or
            $taskFresh.receipt.git_head -cne $taskInput.canonical_head -or $taskFresh.receipt.platform -cne 'x64' -or
            $taskFresh.receipt.configuration -cne 'Release' -or $taskFresh.receipt.target -cne 'solution' -or -not $taskFresh.receipt.tests_enabled) {throw 'Archived fresh build metadata does not bind the frozen source/profile/receipt.'}
        $taskValidationPath=Assert-TaskFile (Join-Path $taskCapture 'capture-validation.json');$taskValidation=Read-TaskJson $taskValidationPath
        if ($taskValidation.valid -isnot [bool] -or -not $taskValidation.valid -or $null -ne $taskValidation.error) {throw 'Capture-validation is not a closed valid record.'}
        $taskControlsPath=Assert-TaskFile (Join-Path $taskCapture 'capture-controls.json');$taskControls=Read-TaskJson $taskControlsPath
        if ($taskControls.machine -cne '4cb089111a23' -or $taskControls.process -ne $taskInput.native_process -or
            $taskControls.metric_run -cne $taskInput.native_metric_run -or $taskControls.fixture_sha256 -cne $taskInput.fixture_sha -or
            $taskControls.metric_metadata.metric_commit -cne $taskInput.native_commit -or $taskControls.metric_metadata.metric_branch -cne $taskInput.native_branch -or
            $taskControls.metric_metadata.canonical_git_head -cne $taskInput.canonical_head -or $taskControls.metric_metadata.source_snapshot_id -cne $taskInput.source_snapshot -or
            $taskControls.metric_metadata.exception -cne $taskInput.metadata_exception) {throw 'Literal capture/native machine/run/process/commit controls differ.'}
        if ($taskFamily -ceq 'F6') {
            if ($taskControls.mode -cne $taskInput.role -or $taskControls.f6_rows -ne 208 -or @($taskControls.samples).Count -ne 18) {throw 'F6 mode/208 rows/18 intervals mismatch.'}
        } elseif ($taskControls.integrated_rows -ne 275 -or $taskControls.controls -cne $taskInput.expected_controls -or $taskControls.metric_metadata.variant -cne $taskInput.role) {throw 'V2 275 rows or exact geometry/workload controls mismatch.'}
        $taskSource=Get-TaskBoundFile $taskManifest $taskWhole 'artifacts/selftest/last_run/perf/perf_metrics.jsonl'
        if ($taskSource.sha256 -cne $taskControls.raw_metric_sha256) {throw 'Control raw digest does not bind whole-archive copied digest.'}
        $taskReport=Get-TaskBoundFile $taskManifest $taskWhole 'artifacts/selftest/last_run/run-all-tests-results.json'
        $taskNative=Get-TaskBoundFile $taskManifest $taskWhole 'artifacts/selftest/last_run/fileops/results.json'
        $taskTrace=Get-TaskBoundFile $taskManifest $taskWhole 'artifacts/selftest/last_run/fileops/trace.txt'
        $taskRunnerTrace=Get-TaskBoundFile $taskManifest $taskWhole 'artifacts/selftest/last_run/trace.txt'
        foreach ($taskSmall in @($taskReport,$taskNative,$taskTrace,$taskRunnerTrace)) {if ((Get-TaskSha $taskSmall.path) -cne $taskSmall.sha256) {throw 'Native/report/trace original differs from closed archive digest.'}}
        $taskReportRecord=Read-TaskJson $taskReport.path;$taskNativeRecord=Read-TaskJson $taskNative.path
        Assert-TaskCounts $taskReportRecord $taskNativePassed;Assert-TaskCounts $taskNativeRecord $taskNativePassed
        if ($taskReportRecord.schema -cne 'red-salamander.run-all-tests.v1' -or $taskReportRecord.run_id -cne $taskInput.run_id -or $taskReportRecord.case_filter -cne $taskCase -or
            $taskReportRecord.configuration -cne 'Release' -or $taskReportRecord.platform -cne 'x64' -or $taskReportRecord.exit_code -ne 0 -or
            $taskNativeRecord.suite -cne 'FileOperations' -or $taskNativeRecord.repeat_count -ne 1 -or
            @($taskNativeRecord.cases).Count -ne $taskNativePassed -or @($taskNativeRecord.cases | Where-Object status -CNE passed).Count -ne 0 -or
            @($taskNativeRecord.cases | Where-Object name -CEQ $taskCase).Count -ne 1) {throw 'Exact native/report case/profile identity failed.'}
        if ($taskFamily -ceq 'F6' -and $taskControls.trace_sha256 -cne $taskRunnerTrace.sha256) {throw 'F6 full self-test trace control digest differs.'}
        $taskMetadataFiles=@($taskManifestPath,$taskBindingPath,$taskFreshPath,$taskValidationPath,$taskControlsPath)
        $taskMetadata=@($taskMetadataFiles | ForEach-Object {[pscustomobject]@{path=$_;bytes=(Get-Item -LiteralPath $_).Length;sha256=(Get-TaskSha $_)}})
        $taskPrepared.Add([pscustomobject]@{input=$taskInput;source=$taskSource;report=$taskReport;native=$taskNative;trace=$taskTrace;runner_trace=$taskRunnerTrace;metadata=$taskMetadata;expected_rows=$taskExpectedRows})
    }
    foreach ($taskReady in $taskPrepared) {
        $taskInput=$taskReady.input;$taskSource=$taskReady.source.path
        $taskTarget=Join-Path $taskOutput ($taskInput.family.ToLowerInvariant()+'/'+$taskInput.run_id)
        [void](New-Item -ItemType Directory -Path $taskTarget -Force)
        $taskPrefix=if ($taskInput.family -ceq 'F6') {'FileOps.PopupTelemetry.'}else{'FileOps.IntegratedPresentation.'}
        $taskSourceBefore=Get-TaskSha $taskSource
        $taskRows=0L;$taskLines=0L;$taskBytes=0L;$taskDecodedLfBytes=0L
        $taskLineNumbers=[Collections.Generic.List[long]]::new();$taskSegments=[Collections.Generic.List[object]]::new()
        $taskReader=$null;$taskWriter=$null;$taskSegment=$null;$taskSegmentBytes=0L;$taskOrdinal=0;$taskRunError=$null;$taskSourceAfter=$null
        try {
            if ($taskSourceBefore -cne $taskReady.source.sha256) {throw 'Original raw SHA changed before curation.'}
            $taskProbe=[IO.File]::OpenRead($taskSource)
            try {
                $taskFirst=[byte[]]::new(4096);$taskCount=$taskProbe.Read($taskFirst,0,$taskFirst.Length)
                if ($taskCount -ge 3 -and $taskFirst[0] -eq 239 -and $taskFirst[1] -eq 187 -and $taskFirst[2] -eq 191) {throw 'Native stream BOM rejected.'}
                $taskLf=[Array]::IndexOf($taskFirst,[byte]10,0,$taskCount)
                if ($taskLf -lt 0 -or ($taskLf -gt 0 -and $taskFirst[$taskLf-1] -eq 13)) {throw 'Native LF required.'}
            } finally {$taskProbe.Dispose()}
            $taskReader=[IO.StreamReader]::new($taskSource,$taskUtf8,$false,65536)
            while ($null -ne ($taskLine=$taskReader.ReadLine())) {
                ++$taskLines;$taskLineBytes=$taskUtf8.GetByteCount($taskLine)+1;$taskDecodedLfBytes+=$taskLineBytes
                if ($taskLine.IndexOf($taskPrefix,[StringComparison]::Ordinal) -lt 0) {continue}
                $taskRow=$taskLine | ConvertFrom-Json -DateKind String
                if (-not $taskRow.metric.StartsWith($taskPrefix,[StringComparison]::Ordinal)) {continue}
                $taskExpectedHr=if ($taskInput.family -ceq 'V2' -and $taskRow.metric -ceq 'FileOps.IntegratedPresentation.SurfaceCounters') {[long]2147500033}else{0L}
                if ($taskRow.machineHash -cne '4cb089111a23' -or $taskRow.build -cne 'Release' -or $taskRow.runId -cne $taskInput.native_metric_run -or
                    $taskRow.process -ne $taskInput.native_process -or $taskRow.commit -cne $taskInput.native_commit -or
                    $taskRow.branch -cne $taskInput.native_branch -or $taskRow.hr -ne $taskExpectedHr) {throw 'Exact selected native row profile/run/process/HRESULT differs.'}
                if ($taskLineBytes -gt 2MB) {throw 'One original row exceeds the file cap.'}
                if ($null -eq $taskWriter -or $taskSegmentBytes+$taskLineBytes -gt 2MB) {
                    if ($null -ne $taskWriter) {$taskWriter.Dispose();$taskWriter=$null}
                    ++$taskOrdinal;$taskSegmentFolder=Join-Path $taskTarget ('raw-{0:D2}' -f $taskOrdinal)
                    [void](New-Item -ItemType Directory -Path $taskSegmentFolder)
                    $taskSegmentPath=Join-Path $taskSegmentFolder 'perf_metrics.jsonl'
                    $taskWriter=[IO.StreamWriter]::new($taskSegmentPath,$false,$taskUtf8,65536);$taskWriter.NewLine="`n"
                    $taskSegmentBytes=0L;$taskSegment=[pscustomobject]@{path=$taskSegmentPath;rows=0L;bytes=0L;first_original_line=$taskLines;last_original_line=$taskLines;sha256=$null}
                    $taskSegments.Add($taskSegment)
                }
                # Write original row TEXT, never converted values. UInt64 digits remain verbatim.
                $taskWriter.WriteLine($taskLine);$taskSegment.bytes+=$taskLineBytes;++$taskSegment.rows;$taskSegment.last_original_line=$taskLines
                $taskSegmentBytes+=$taskLineBytes;$taskBytes+=$taskLineBytes;++$taskRows;$taskLineNumbers.Add($taskLines)
            }
            if ($taskRows -ne $taskReady.expected_rows -or $taskDecodedLfBytes -ne $taskReady.source.bytes) {throw 'Exact all-row count or whole native-LF byte accounting failed.'}
            foreach ($taskMeta in $taskReady.metadata) {if ((Get-TaskSha $taskMeta.path) -cne $taskMeta.sha256) {throw 'Closure/control metadata changed during curation.'}}
            foreach ($taskSmall in @($taskReady.report,$taskReady.native,$taskReady.trace,$taskReady.runner_trace)) {if ((Get-TaskSha $taskSmall.path) -cne $taskSmall.sha256) {throw 'Original native/report/trace changed during curation.'}}
        } catch {$taskRunError=$_.Exception.Message}
        finally {
            if ($null -ne $taskWriter) {$taskWriter.Dispose()};if ($null -ne $taskReader) {$taskReader.Dispose()}
            try {
                $taskSourceAfter=Get-TaskSha $taskSource
                if ($taskSourceBefore -cne $taskSourceAfter -or (Get-Item -LiteralPath $taskSource).Length -ne $taskReady.source.bytes) {throw 'Original raw changed after curation.'}
                foreach ($taskSegment in $taskSegments) {
                    if ((Get-Item -LiteralPath $taskSegment.path).Length -ne $taskSegment.bytes -or $taskSegment.bytes -gt 2MB) {throw 'Segment bytes differ/overflow.'}
                    $taskSegment.sha256=Get-TaskSha $taskSegment.path
                }
            } catch {if ($null -eq $taskRunError) {$taskRunError=$_.Exception.Message}else{$taskRunError+='; '+$_.Exception.Message}}
            $taskRecord=[ordered]@{schema='i27.f6-v2-original-row-selection.proposed.v1';status=$(if ($null -eq $taskRunError) {'curated-original-rows'}else{'failed'});error=$taskRunError;family=$taskInput.family;role=$taskInput.role;ordinal=$taskInput.ordinal;pair=$taskInput.pair;run_id=$taskInput.run_id;native_metric_run_id=$taskInput.native_metric_run;native_metric_process=$taskInput.native_process;original_path=$taskSource;original_sha256_before=$taskSourceBefore;original_sha256_after=$taskSourceAfter;archive_copied_sha256=$taskReady.source.sha256;original_bytes=$taskReady.source.bytes;original_line_count=$taskLines;decoded_native_lf_bytes=$taskDecodedLfBytes;selected_rows=$taskRows;expected_rows=$taskReady.expected_rows;selected_bytes=$taskBytes;original_line_numbers=@($taskLineNumbers.ToArray());segments=@($taskSegments.ToArray());closed_metadata=$taskReady.metadata;native_result=$taskReady.native;runner_report=$taskReady.report;trace=$taskReady.trace;runner_trace=$taskReady.runner_trace;literal_controls=$taskInput;input_model_sha256=$taskModelSha;curator=$PSCommandPath;curator_sha256=(Get-TaskSha $PSCommandPath);selector=('ALL original rows with metric prefix '+$taskPrefix+'; no sampling/adverse removal');byte_rule='Strict UTF-8 original text plus native LF, whole-stream decoded-LF byte count must equal source size; no numeric rewrite or UInt64 conversion in output';qualification='Curation only; native original failures stay failures; no nonregression, GPU peak, FPS, global-clean, Full or I27 qualification.'}
            $taskRecordPath=Join-Path $taskTarget 'row-selection.json'
            [IO.File]::WriteAllText($taskRecordPath,($taskRecord | ConvertTo-Json -Depth 12)+"`n",$taskUtf8)
            $taskRecordBytes=(Get-Item -LiteralPath $taskRecordPath).Length
            if ($taskRecordBytes -gt 2MB -or $taskBytes+$taskRecordBytes -gt 5MB) {
                $taskRecord.status='failed';$taskRecord.error='Per-file or same-original-run curation cap exceeded; no resegmentation into invented runs.'
                [IO.File]::WriteAllText($taskRecordPath,($taskRecord | ConvertTo-Json -Depth 12)+"`n",$taskUtf8)
                $taskRunError=$taskRecord.error
            }
            $taskSummary.Add([pscustomobject]@{family=$taskInput.family;role=$taskInput.role;run_id=$taskInput.run_id;rows=$taskRows;bytes=$taskBytes;segments=$taskSegments.Count;folder=$taskTarget;status=$taskRecord.status;error=$taskRecord.error})
        }
        if ($null -ne $taskRunError) {throw 'Curation failed; preserve partial output and every original; do not publish.'}
    }
} catch {$taskFailure=$_.Exception.Message}
finally {
    $taskFinal=[ordered]@{schema='i27.f6-v2-curation-summary.proposed.v1';status=$(if ($null -eq $taskFailure -and $taskSummary.Count -eq 16) {'curated-original-rows'}else{'failed'});error=$taskFailure;input_model=$taskModelPath;input_model_sha256=$taskModelSha;expected_original_runs=16;completed_original_runs=$taskSummary.Count;preflight_passed_original_runs=$taskPrepared.Count;runs=@($taskSummary.ToArray());output=$taskOutput;qualification='Evidence-only curation; no publication or nonregression/Full/I27 qualification.'}
    [IO.File]::WriteAllText((Join-Path $taskOutput 'summary.json'),($taskFinal | ConvertTo-Json -Depth 8)+"`n",$taskUtf8)
    Write-Output ('CURATION_PACKET='+$taskOutput)
}
if ($null -ne $taskFailure -or $taskSummary.Count -ne 16) {exit 1}
exit 0
