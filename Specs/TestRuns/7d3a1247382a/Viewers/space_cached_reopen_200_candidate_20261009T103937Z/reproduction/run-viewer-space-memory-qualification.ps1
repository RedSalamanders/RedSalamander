[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateSet('FourRootAbba', 'CandidateHeavy', 'CachedReopen200')][string]$Mode,
    [Parameter(Mandatory = $true)][string]$ExpectedCandidateReceiptId,
    [Parameter(Mandatory = $true)][string]$ExpectedCandidateBuildSourceSnapshotId,
    [Parameter(Mandatory = $true)][string]$ExpectedCandidateBuildLogPath,
    [string]$ExpectedBaselineReceiptId = '',
    [string]$ExpectedBaselineBuildLogPath = '',
    [string]$ExpectedBaselineBuildSourceSnapshotId = '',
    [string]$BaselineRoot = 'C:\Users\PVB\.codex\worktrees\viewer-space-baseline\RedSalamander',
    [string]$CandidateRoot = 'C:\Users\PVB\.codex\worktrees\viewer-space-explorer-analyst\RedSalamander',
    [string]$TestRoot = 'D:\RedSalamander.Perf'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$testRootFull = [IO.Path]::GetFullPath($TestRoot).TrimEnd('\')
if ($testRootFull -cne 'D:\RedSalamander.Perf') { throw "Reviewed run is bound to D:\RedSalamander.Perf; got '$testRootFull'." }
$marker = Join-Path $testRootFull '.red-salamander-test-root'
if (-not (Test-Path -LiteralPath $marker -PathType Leaf) -or
    (Get-Content -LiteralPath $marker -Raw).TrimEnd("`r", "`n") -cne 'red-salamander.test-sandbox-root.v1') {
    throw "Test root is not an already initialized marked sandbox: $testRootFull"
}
$BaselineRoot = [IO.Path]::GetFullPath($BaselineRoot)
$CandidateRoot = [IO.Path]::GetFullPath($CandidateRoot)
Import-Module (Join-Path $CandidateRoot 'Tools\Modules\Testing\TestSandbox.psm1') -Force -Scope Local -ErrorAction Stop
Import-Module (Join-Path $CandidateRoot 'Tools\Modules\Build\ProcessStreaming.psm1') -Force -Scope Local -ErrorAction Stop
$python = (Get-Command python -ErrorAction Stop).Source
$pwsh = (Get-Command pwsh -ErrorAction Stop).Source
$curator = Join-Path $CandidateRoot '.build\curate-viewer-space-memory-run.py'
if (-not (Test-Path -LiteralPath $curator -PathType Leaf)) { throw "Missing ignored curator: $curator" }

$isFourRootAbba = $Mode -eq 'FourRootAbba'
$isCandidateHeavy = $Mode -eq 'CandidateHeavy'
$isHeavy = $isFourRootAbba -or $isCandidateHeavy
if ($isFourRootAbba -and ([string]::IsNullOrWhiteSpace($ExpectedBaselineReceiptId) -or [string]::IsNullOrWhiteSpace($ExpectedBaselineBuildSourceSnapshotId) -or [string]::IsNullOrWhiteSpace($ExpectedBaselineBuildLogPath))) {
    throw 'FourRootAbba requires exact baseline Release receipt ID, build source_snapshot_id, and archived MSBuild log path.'
}
if ($isCandidateHeavy -and (-not [string]::IsNullOrWhiteSpace($ExpectedBaselineReceiptId) -or -not [string]::IsNullOrWhiteSpace($ExpectedBaselineBuildSourceSnapshotId) -or -not [string]::IsNullOrWhiteSpace($ExpectedBaselineBuildLogPath))) {
    throw 'CandidateHeavy is candidate-only and rejects baseline receipt, build source_snapshot_id, or log inputs.'
}
$caseName = if ($isHeavy) { 'cmd_viewer_space_memory_heavy_four_root_optin' } else { 'cmd_viewer_space_cache_reopen_200_first_map_optin' }
$caseFilter = $caseName
$campaignId = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
$campaignRoot = Assert-RSTestSandboxContainedPath -TestRoot $testRootFull -Path (Join-Path $testRootFull "evidence\viewer-space-memory-$($Mode.ToLowerInvariant())-$campaignId") -RequireDescendant
New-Item -ItemType Directory -Path $campaignRoot | Out-Null
$logsRoot = Join-Path $campaignRoot 'runner-logs'
New-Item -ItemType Directory -Path $logsRoot | Out-Null
$rawSourcesRoot = Join-Path $campaignRoot 'raw-sources'
New-Item -ItemType Directory -Path $rawSourcesRoot | Out-Null
$manifestPath = Join-Path $campaignRoot 'campaign-manifest.json'
$manifest = [ordered]@{
    schema = 'viewer-space.memory-qualification.v1'; campaign_id = $campaignId; mode = $Mode
    created_utc = [DateTime]::UtcNow.ToString('o'); test_root = $testRootFull
    platform = 'x64'; configuration = 'Release'; case_filter = $caseFilter
    ordering = if ($isFourRootAbba) { 'baseline warmup, candidate warmup, then three serial ABBA blocks; each block has two mirrored adjacent baseline/candidate contrasts (six pairs total).' } elseif ($isCandidateHeavy) { 'single candidate-only four-root 100,000-file-per-root lifecycle run; no baseline warmup, no candidate/baseline comparison, no memory-reduction ratio.' } else { 'single candidate-only 200-sample cached-reopen run; not a baseline comparison.' }
    timeout_ms = if ($isHeavy) { 1200000 } else { 420000 }
    archive_limits = @{ per_file_bytes = 2097152; per_run_bytes = 5242880 }
    expected_receipts = @{ baseline = if ($isFourRootAbba) { $ExpectedBaselineReceiptId } else { $null }; candidate = $ExpectedCandidateReceiptId }
    expected_build_source_snapshot_ids = @{ baseline = if ($isFourRootAbba) { $ExpectedBaselineBuildSourceSnapshotId } else { $null }; candidate = $ExpectedCandidateBuildSourceSnapshotId }
    expected_build_log_paths = @{ baseline = if ($isFourRootAbba) { $ExpectedBaselineBuildLogPath } else { $null }; candidate = $ExpectedCandidateBuildLogPath }
    workspace_snapshots = 'Capture each runner workspace-final.json snapshot_id independently; it is distinct from the build receipt source_snapshot_id.'
    raw_sources_root = 'raw-sources'; reproduction = @{}; runs = @(); measurements = @()
}
if ($isCandidateHeavy) {
    $manifest.qualification_scope = 'candidate-only-four-root-lifecycle'
    $manifest.memory_reduction_target = @{ targetPercent = 70; status = 'UNMEASURED'; reason = 'A candidate-only run has no baseline denominator.' }
}

function Save-Manifest {
    $temp = "$manifestPath.tmp"
    [IO.File]::WriteAllText($temp, ($manifest | ConvertTo-Json -Depth 32) + "`n", [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temp -Destination $manifestPath -Force | Out-Null
}
function Get-Sha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Assert-ExpectedBuildLogPath([string]$Repo, [string]$ExpectedPath) {
    if ([string]::IsNullOrWhiteSpace($ExpectedPath) -or -not [IO.Path]::IsPathRooted($ExpectedPath)) { throw "An absolute exact archived MSBuild log path is required for '$Repo'." }
    $logRoot = [IO.Path]::GetFullPath((Join-Path $Repo '.build\logs')).TrimEnd('\') + '\'
    $exactPath = [IO.Path]::GetFullPath($ExpectedPath)
    if (-not $exactPath.StartsWith($logRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Expected MSBuild log must be beneath this repository's .build\logs directory: $exactPath"
    }
    if (-not (Test-Path -LiteralPath $exactPath -PathType Leaf)) { throw "Expected exact MSBuild log does not exist: $exactPath" }
    $item = Get-Item -LiteralPath $exactPath -Force
    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Expected MSBuild log cannot be a reparse point: $exactPath" }
    return $item
}

# Seal the ignored launcher/analysis inputs with each campaign before running.
$reproductionRoot = Join-Path $campaignRoot 'reproduction'
New-Item -ItemType Directory -Path $reproductionRoot | Out-Null
foreach ($path in @($PSCommandPath, $curator, (Join-Path $CandidateRoot '.build\analyze-viewer-space-memory-campaign.py'))) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing campaign reproduction input: $path" }
    $copy = Join-Path $reproductionRoot ([IO.Path]::GetFileName($path))
    [IO.File]::Copy($path, $copy, $false)
    $sourceHash = Get-Sha256 $path
    if ((Get-Sha256 $copy) -cne $sourceHash) { throw "Campaign helper preservation mismatch: $path" }
    $manifest.reproduction[[IO.Path]::GetFileName($path)] = @{ sha256 = $sourceHash; bytes = (Get-Item -LiteralPath $copy).Length; path = "reproduction\$([IO.Path]::GetFileName($path))" }
}
Save-Manifest

function Invoke-One([string]$Side, [string]$Role, [int]$Block, [int]$Position) {
    $repo = if ($Side -eq 'baseline') { $BaselineRoot } else { $CandidateRoot }
    $expectedReceipt = if ($Side -eq 'baseline') { $ExpectedBaselineReceiptId } else { $ExpectedCandidateReceiptId }
    $expectedSourceSnapshot = if ($Side -eq 'baseline') { $ExpectedBaselineBuildSourceSnapshotId } else { $ExpectedCandidateBuildSourceSnapshotId }
    $receiptPath = Join-Path $repo '.build\x64\Release\build-receipt.json'
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { throw "Missing Release receipt: $receiptPath" }
    $receiptBytes = [IO.File]::ReadAllBytes($receiptPath)
    $receipt = [Text.Encoding]::UTF8.GetString($receiptBytes) | ConvertFrom-Json
    if ($receipt.schema -cne 'red-salamander.build-receipt.v1' -or $receipt.configuration -cne 'Release' -or $receipt.platform -cne 'x64' -or
        [int]$receipt.exit_code -ne 0 -or $receipt.tests_enabled -ne $true -or $receipt.receipt_id -cne $expectedReceipt -or
        $receipt.source_snapshot_id -cne $expectedSourceSnapshot) {
        throw "Exact verified Release receipt mismatch for $Side ($receiptPath)."
    }
    $receiptHash = Get-Sha256 $receiptPath
    $expectedBuildLogPath = if ($Side -eq 'baseline') { $ExpectedBaselineBuildLogPath } else { $ExpectedCandidateBuildLogPath }
    $buildLog = Assert-ExpectedBuildLogPath $repo $expectedBuildLogPath
    $buildInputs = Join-Path $campaignRoot "build-inputs\$Side"
    if (-not (Test-Path -LiteralPath $buildInputs -PathType Container)) { New-Item -ItemType Directory -Path $buildInputs | Out-Null }
    $sealedReceipt = Join-Path $buildInputs 'build-receipt.json'
    $sealedBuildLog = Join-Path $buildInputs 'msbuild.log'
    if (-not (Test-Path -LiteralPath $sealedReceipt -PathType Leaf)) {
        [IO.File]::Copy($receiptPath, $sealedReceipt, $false)
        $buildLogBytes = [IO.File]::ReadAllBytes($buildLog.FullName)
        if ($buildLogBytes.Length -gt (5MB - (Get-Item -LiteralPath $sealedReceipt).Length)) { throw 'Build input archive would exceed the 5 MiB per-run evidence bound.' }
        $buildParts = @()
        for ($offset = 0; $offset -lt $buildLogBytes.Length; $offset += 2MB) {
            $length = [Math]::Min(2MB, $buildLogBytes.Length - $offset)
            $partName = if ($buildParts.Count -eq 0) { 'msbuild.log' } else { "msbuild-part$($buildParts.Count + 1).log" }
            $partPath = Join-Path $buildInputs $partName
            $part = [byte[]]::new($length)
            [Array]::Copy($buildLogBytes, $offset, $part, 0, $length)
            [IO.File]::WriteAllBytes($partPath, $part)
            $buildParts += [pscustomobject]@{ path=$partPath; bytes=$length; sha256=(Get-Sha256 $partPath) }
        }
        $sealedBuildLog = $buildParts[0].path
        $buildLogIndex = [ordered]@{ source_path=$buildLog.FullName; source_bytes=$buildLogBytes.Length; source_sha256=(Get-Sha256 $buildLog.FullName); chunks=$buildParts }
        [IO.File]::WriteAllText((Join-Path $buildInputs 'msbuild-log-index.json'), ($buildLogIndex | ConvertTo-Json -Depth 8) + "`n", [Text.UTF8Encoding]::new($false))
    }
    $buildInputFiles = @(Get-ChildItem -LiteralPath $buildInputs -File)
    if (@($buildInputFiles | Where-Object { $_.Length -gt 2MB }).Count -gt 0 -or
        [long](($buildInputFiles | Measure-Object -Property Length -Sum).Sum) -gt 5MB) {
        throw 'Sealed build receipt/log evidence violates the 2 MiB per-file or 5 MiB per-build evidence bound.'
    }
    if ((Get-Sha256 $sealedReceipt) -cne $receiptHash) { throw "Sealed receipt changed for $Side." }
    $buildLogIndex = Get-Content -LiteralPath (Join-Path $buildInputs 'msbuild-log-index.json') -Raw | ConvertFrom-Json
    $buildLogHash = [string]$buildLogIndex.source_sha256
    if ([IO.Path]::GetFullPath([string]$buildLogIndex.source_path) -cne [IO.Path]::GetFullPath($buildLog.FullName) -or
        [long]$buildLogIndex.source_bytes -ne [long]$buildLog.Length -or $buildLogHash -cne (Get-Sha256 $buildLog.FullName)) {
        throw "Exact MSBuild log path/length/SHA changed since it was sealed for $Side."
    }
    $label = if ($Role -eq 'measurement') { "$campaignId-b$Block-p$Position-$Side" } elseif ($Role -eq 'candidate-heavy') { "$campaignId-candidate-heavy" } else { "$campaignId-warmup-$Side" }
    $stdout = Join-Path $logsRoot "$label.stdout.txt"; $stderr = Join-Path $logsRoot "$label.stderr.txt"
    $temp = Assert-RSTestSandboxContainedPath -TestRoot $testRootFull -Path (Join-Path $campaignRoot "temp\$label") -RequireDescendant
    New-Item -ItemType Directory -Path $temp -Force | Out-Null
    $runEnv = @{
        REDSALAMANDER_TEST_ROOT = $testRootFull; TEMP = $temp; TMP = $temp
        REDSALAMANDER_VIEWERSPACE_MEMORY_HEAVY = if ($isHeavy) { '1' } else { $null }
        REDSALAMANDER_VIEWERSPACE_CACHED_REOPEN_200 = if (-not $isHeavy) { '1' } else { $null }
        REDSALAMANDER_VIEWERSPACE_SCAN_PHASES = $null
        REDSALAMANDER_VIEWERSPACE_PUBLICATION_PHASES = $null
        REDSALAMANDER_VIEWERSPACE_TEST_CACHE_RESTORE_PHASES_PER_TURN = $null
        REDSALAMANDER_VIEWERSPACE_TEST_CACHE_CAP_BYTES = $null
        REDSALAMANDER_VIEWERSPACE_TEST_CACHE_AGGREGATE_BYTES = $null
        REDSALAMANDER_VIEWERSPACE_TEST_CORRUPT_CACHE_NAME_OFFSET = $null
        REDSALAMANDER_VIEWERSPACE_TEST_CORRUPT_CACHE_ROOT_STATE = $null
        REDSALAMANDER_VIEWERSPACE_TEST_CORRUPT_CACHE_CHILD_OWNER = $null
        REDSALAMANDER_VIEWERSPACE_TEST_CORRUPT_CACHE_MISSING_CHILD = $null
        REDSALAMANDER_VIEWERSPACE_TEST_SCAN_THREADS = $null
    }
    $runnerArgs = @('-NoLogo','-NoProfile','-File',(Join-Path $repo 'Tools\Run-AllTests.ps1'),
        '-Suite','Commands','-SkipBuild','-CaseFilter',$caseFilter,'-SelfTestRepeat','1',
        '-TestRoot',$testRootFull,'-Platform','x64','-Configuration','Release')
    $timeout = if ($isHeavy) { 1200000 } else { 420000 }
    $buildLogIndexPath = Join-Path $buildInputs 'msbuild-log-index.json'
    $attempt = [pscustomobject][ordered]@{
        label=$label; side=$Side; role=$Role; block=$Block; position=$Position; run_id=''
        status='STARTING'; failure_phase=$null; error=$null; started_utc=[DateTime]::UtcNow.ToString('o'); finished_utc=$null
        exit_code=$null; summary_exit_code=$null; case_status='pending'; receipt_id=$receipt.receipt_id
        receipt_source_snapshot_id=$receipt.source_snapshot_id; receipt_sha256=$receiptHash; workspace_snapshot_id=$null
        receipt_ended_utc=$receipt.ended_utc; expected_build_log_path=$expectedBuildLogPath
        build_log_source_path=$buildLogIndex.source_path; build_log_index_path=$buildLogIndexPath
        build_log_sha256=$buildLogHash; build_log_bytes=$buildLogIndex.source_bytes
        stdout_sha256=$null; stderr_sha256=$null; stdout_bytes=$null; stderr_bytes=$null
        raw_metrics_original_path=$null; raw_metrics_preserved_path=$null; raw_metrics_bytes=$null; raw_metrics_sha256=$null
        archive_path=(Join-Path $campaignRoot "runs\$label"); curation_status=$null; curation_errors=@()
    }
    $manifest.runs = @($manifest.runs) + @($attempt)
    $measurementAttempt = $null
    if ($Role -eq 'measurement') {
        $measurementAttempt = [pscustomobject]@{label=$label;side=$Side;block=$Block;position=$Position;run_id=''}
        $manifest.measurements = @($manifest.measurements) + @($measurementAttempt)
    }
    Save-Manifest
    try {
        $exit = Invoke-RSStreamingProcess -FilePath $pwsh -Arguments $runnerArgs -WorkingDirectory $repo -LogPath $stdout -ErrorLogPath $stderr `
            -TimeoutMilliseconds $timeout -AdditionalEnvironment $runEnv
    } catch {
        $attempt.status = 'FAILED'; $attempt.failure_phase = 'runner'; $attempt.error = $_.Exception.Message
        $attempt.finished_utc = [DateTime]::UtcNow.ToString('o')
        if (Test-Path -LiteralPath $stdout -PathType Leaf) { $attempt.stdout_bytes = (Get-Item -LiteralPath $stdout).Length; $attempt.stdout_sha256 = Get-Sha256 $stdout }
        if (Test-Path -LiteralPath $stderr -PathType Leaf) { $attempt.stderr_bytes = (Get-Item -LiteralPath $stderr).Length; $attempt.stderr_sha256 = Get-Sha256 $stderr }
        Save-Manifest
        throw
    }
    $attempt.exit_code = $exit; $attempt.status = 'RUN_RETURNED'
    Save-Manifest
    try {
        if (Test-Path -LiteralPath $stdout -PathType Leaf) { $attempt.stdout_bytes = (Get-Item -LiteralPath $stdout).Length; $attempt.stdout_sha256 = Get-Sha256 $stdout }
        if (Test-Path -LiteralPath $stderr -PathType Leaf) { $attempt.stderr_bytes = (Get-Item -LiteralPath $stderr).Length; $attempt.stderr_sha256 = Get-Sha256 $stderr }
        $output = if (Test-Path -LiteralPath $stdout -PathType Leaf) { Get-Content -LiteralPath $stdout -Raw } else { '' }
        $matches = [regex]::Matches($output, 'Run id:\s*([0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32})')
        if ($matches.Count -ne 1) { throw "Expected one canonical Run id; exit=$exit. Logs: $stdout / $stderr" }
        $runId = $matches[0].Groups[1].Value
    } catch {
        $attempt.status = 'FAILED'; $attempt.failure_phase = 'run-id'; $attempt.error = $_.Exception.Message
        $attempt.finished_utc = [DateTime]::UtcNow.ToString('o'); Save-Manifest
        throw
    }
    $attempt.run_id = $runId; $attempt.status = 'RUNNING'
    if ($null -ne $measurementAttempt) { $measurementAttempt.run_id = $runId }
    Save-Manifest
    $runRoot = Join-Path (Join-Path $testRootFull 'runs') $runId
    $raw = Join-Path $runRoot 'artifacts\selftest\last_run'
    $archive = Join-Path $campaignRoot "runs\$label"
    $rawMetricsOriginal = Join-Path $raw 'perf\perf_metrics.jsonl'
    $rawMetricsPreserved = Join-Path (Join-Path $rawSourcesRoot $label) 'perf_metrics.jsonl'
    $rawMetricsSourceIndex = Join-Path (Join-Path $rawSourcesRoot $label) 'source-index.json'
    $rawSourceRecord = $null
    try {
        $rawSourceDirectory = Assert-RSTestSandboxContainedPath -TestRoot $testRootFull -Path (Join-Path $rawSourcesRoot $label) -RequireDescendant
        if (Test-Path -LiteralPath $rawSourceDirectory) { throw "Refusing to overwrite raw-source evidence: $rawSourceDirectory" }
        New-Item -ItemType Directory -Path $rawSourceDirectory | Out-Null
        if (Test-Path -LiteralPath $rawMetricsOriginal -PathType Leaf) {
            $rawMetricsOriginalFull = Assert-RSTestSandboxContainedPath -TestRoot $testRootFull -Path $rawMetricsOriginal -RequireDescendant
            $rawMetricsPreservedFull = Assert-RSTestSandboxContainedPath -TestRoot $testRootFull -Path $rawMetricsPreserved -RequireDescendant
            $rawMetricsSourceIndexFull = Assert-RSTestSandboxContainedPath -TestRoot $testRootFull -Path $rawMetricsSourceIndex -RequireDescendant
            [IO.File]::Copy($rawMetricsOriginalFull, $rawMetricsPreservedFull, $false)
            $rawLength = [long](Get-Item -LiteralPath $rawMetricsOriginalFull).Length
            $rawHash = Get-Sha256 $rawMetricsOriginalFull
            if ([long](Get-Item -LiteralPath $rawMetricsPreservedFull).Length -ne $rawLength -or (Get-Sha256 $rawMetricsPreservedFull) -cne $rawHash) {
                throw "Campaign raw-source preservation was not byte-exact: $label"
            }
            $rawSourceRecord = [ordered]@{ original_path=$rawMetricsOriginalFull; preserved_path=$rawMetricsPreservedFull; bytes=$rawLength; sha256=$rawHash }
            [IO.File]::WriteAllText($rawMetricsSourceIndexFull, ($rawSourceRecord | ConvertTo-Json -Depth 4) + "`n", [Text.UTF8Encoding]::new($false))
            $attempt.raw_metrics_original_path = $rawSourceRecord.original_path
            $attempt.raw_metrics_preserved_path = $rawSourceRecord.preserved_path
            $attempt.raw_metrics_bytes = $rawSourceRecord.bytes
            $attempt.raw_metrics_sha256 = $rawSourceRecord.sha256
        } else {
            $attempt.raw_metrics_original_path = [IO.Path]::GetFullPath($rawMetricsOriginal)
            $attempt.error = 'Raw perf_metrics.jsonl was missing; the run will be preserved as unqualified.'
        }
        $attempt.status = 'CURATION_PENDING'
        Save-Manifest
    } catch {
        $attempt.status = 'FAILED'; $attempt.failure_phase = 'raw_source_preservation'; $attempt.error = $_.Exception.Message
        $attempt.finished_utc = [DateTime]::UtcNow.ToString('o'); Save-Manifest
        throw
    }
    $curateArgs = @('-X','utf8',$curator,$raw,$archive,$runId,$Side,$sealedReceipt,$expectedReceipt,$expectedSourceSnapshot,$caseName,$Mode,$stdout,$stderr,$buildLogIndexPath,$rawMetricsPreserved,$rawMetricsSourceIndex)
    try {
        $curateOutput = @(& $python @curateArgs)
        $curationExit = [int]$LASTEXITCODE
        if ($curationExit -ne 0 -or $curateOutput.Count -eq 0) { throw "Raw run curation failed closed (exit $curationExit)." }
        $curation = [string]$curateOutput[-1] | ConvertFrom-Json
    } catch {
        $attempt.status = 'FAILED'; $attempt.failure_phase = 'curation'; $attempt.error = "Raw run curation failed for $label; raw-source copy and available run artifacts were retained. $($_.Exception.Message)"
        $attempt.exit_code = $exit; $attempt.curation_status = 'PRESERVED_UNQUALIFIED'
        $attempt.curation_errors = @($attempt.error); $attempt.finished_utc = [DateTime]::UtcNow.ToString('o')
        Save-Manifest
        throw $attempt.error
    }
    $attempt.curation_status = [string]$curation.qualificationStatus
    $attempt.curation_errors = @($curation.errors)
    Save-Manifest
    try {
        $summaryPath = Join-Path $raw 'run-all-tests-results.json'
    $summary = $null
    if (Test-Path -LiteralPath $summaryPath -PathType Leaf) { $summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json }
    $allCaseRows = if ($null -ne $summary) { @($summary.suites | ForEach-Object { $_.cases } | ForEach-Object { $_ }) } else { @() }
    $caseRows = @($allCaseRows | Where-Object { $_.name -ceq $caseName })
    $passed = $exit -eq 0 -and $curation.qualificationStatus -ceq 'qualified' -and $null -ne $summary -and
        $summary.schema -ceq 'red-salamander.run-all-tests.v1' -and $summary.exit_code -eq 0 -and $summary.run_id -ceq $runId -and
        $summary.platform -ceq 'x64' -and $summary.configuration -ceq 'Release' -and $summary.case_filter -ceq $caseFilter -and
        $allCaseRows.Count -eq 1 -and $caseRows.Count -eq 1 -and $caseRows[0].status -ceq 'passed'
    if ((Get-Sha256 $receiptPath) -cne $receiptHash) { throw "Release receipt changed during $label." }
    $workspaceSnapshotId = $null
    $workspacePath = Join-Path $raw 'workspace-final.json'
    if (Test-Path -LiteralPath $workspacePath -PathType Leaf) { $workspaceSnapshotId = (Get-Content -LiteralPath $workspacePath -Raw | ConvertFrom-Json).snapshot_id }
    $attempt.exit_code = $exit
    $attempt.summary_exit_code = if ($null -ne $summary) { $summary.exit_code } else { $null }
    $attempt.case_status = if ($caseRows.Count -eq 1) { $caseRows[0].status } else { 'missing' }
    $attempt.workspace_snapshot_id = $workspaceSnapshotId
    $attempt.stdout_sha256 = Get-Sha256 $stdout; $attempt.stderr_sha256 = Get-Sha256 $stderr
    $attempt.stdout_bytes = (Get-Item -LiteralPath $stdout).Length; $attempt.stderr_bytes = (Get-Item -LiteralPath $stderr).Length
    $attempt.status = if ($passed) { 'PASSED' } else { 'FAILED' }
    $attempt.finished_utc = [DateTime]::UtcNow.ToString('o')
    if (-not $passed) {
        $attempt.failure_phase = 'qualification'
        $attempt.error = if ($curation.errors.Count -gt 0) { @($curation.errors) -join '; ' } else { 'Run summary, case result, or exact execution status failed qualification.' }
    }
    } catch {
        $attempt.status = 'FAILED'; $attempt.failure_phase = 'post_curation_evidence'; $attempt.error = $_.Exception.Message
        $attempt.finished_utc = [DateTime]::UtcNow.ToString('o')
        Save-Manifest
        throw
    }
    Save-Manifest
    Write-Host "$label $($attempt.status) run=$runId"
    if (-not $passed) { throw "Stopping after failed/non-exact case $label. See manifest and preserved artifacts." }
}

if ($isFourRootAbba) {
    Invoke-One 'baseline' 'warmup' 0 0
    Invoke-One 'candidate' 'warmup' 0 0
    for ($block = 1; $block -le 3; ++$block) {
        Invoke-One 'baseline' 'measurement' $block 1
        Invoke-One 'candidate' 'measurement' $block 2
        Invoke-One 'candidate' 'measurement' $block 3
        Invoke-One 'baseline' 'measurement' $block 4
    }
} elseif ($isCandidateHeavy) {
    Invoke-One 'candidate' 'candidate-heavy' 0 1
} else {
    Invoke-One 'candidate' 'measurement' 0 1
}
$manifest.finished_utc = [DateTime]::UtcNow.ToString('o')
Save-Manifest
Write-Host "Completed $Mode. Manifest: $manifestPath"
