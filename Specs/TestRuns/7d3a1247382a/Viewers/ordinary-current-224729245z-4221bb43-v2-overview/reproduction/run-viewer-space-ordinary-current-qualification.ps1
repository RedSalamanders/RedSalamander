[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidatePattern('^[0-9a-f]{64}$')][string]$ExpectedBaselineReceiptId,
    [Parameter(Mandatory = $true)][string]$ExpectedBaselineBuildSourceSnapshotId,
    [Parameter(Mandatory = $true)][string]$ExpectedBaselineBuildLogPath,
    [Parameter(Mandatory = $true)][ValidatePattern('^[0-9a-f]{64}$')][string]$ExpectedCandidateReceiptId,
    [Parameter(Mandatory = $true)][string]$ExpectedCandidateBuildSourceSnapshotId,
    [Parameter(Mandatory = $true)][string]$ExpectedCandidateBuildLogPath,
    [string]$BaselineRoot = 'C:\Users\PVB\.codex\worktrees\viewer-space-baseline\RedSalamander',
    [string]$CandidateRoot = 'C:\Users\PVB\.codex\worktrees\viewer-space-explorer-analyst\RedSalamander',
    [string]$TestRoot = 'D:\RedSalamander.Perf'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$requiredMarker = Join-Path $TestRoot '.red-salamander-test-root'
if (-not (Test-Path -LiteralPath $requiredMarker -PathType Leaf) -or
    (Get-Content -LiteralPath $requiredMarker -Raw).TrimEnd("`r", "`n") -cne 'red-salamander.test-sandbox-root.v1') {
    throw "TestRoot must already be an initialized, marked sandbox: $TestRoot"
}
$testRootFull = [IO.Path]::GetFullPath($TestRoot).TrimEnd('\')
if ($testRootFull -cne 'D:\RedSalamander.Perf') {
    throw "This reviewed campaign is bound to D:\RedSalamander.Perf; got '$testRootFull'."
}

$BaselineRoot = [IO.Path]::GetFullPath($BaselineRoot)
$CandidateRoot = [IO.Path]::GetFullPath($CandidateRoot)
Import-Module (Join-Path $CandidateRoot 'Tools\Modules\Testing\TestSandbox.psm1') -Force -Scope Local -ErrorAction Stop
foreach ($repo in @($BaselineRoot, $CandidateRoot)) {
    if (-not (Test-Path -LiteralPath (Join-Path $repo 'Tools\Run-AllTests.ps1') -PathType Leaf)) {
        throw "Canonical runner is missing from '$repo'."
    }
    if (-not (Test-Path -LiteralPath (Join-Path $repo '.build\x64\Release\build-receipt.json') -PathType Leaf)) {
        throw "Verified x64 Release receipt pointer is missing from '$repo'. Wait for its Release rebuild to finish."
    }
}

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function ConvertTo-DxUiCanonicalJson($Value) {
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [System.Collections.IDictionary]) {
        $parts = foreach ($key in @($Value.Keys | Sort-Object -CaseSensitive)) {
            $keyJson = ConvertTo-Json -InputObject ([string]$key) -Compress
            "$keyJson`:$((ConvertTo-DxUiCanonicalJson $Value[$key]))"
        }
        return '{' + ($parts -join ',') + '}'
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $values = [ordered]@{}
        foreach ($property in @($Value.PSObject.Properties | Sort-Object Name -CaseSensitive)) { $values[$property.Name] = $property.Value }
        return ConvertTo-DxUiCanonicalJson $values
    }
    if ($Value -is [System.Array]) { return '[' + ((@($Value | ForEach-Object { ConvertTo-DxUiCanonicalJson $_ })) -join ',') + ']' }
    return ConvertTo-Json -InputObject $Value -Compress -Depth 100
}

function Assert-DxUiUniqueJsonProperties([System.Text.Json.JsonElement]$Element) {
    if ($Element.ValueKind -eq [System.Text.Json.JsonValueKind]::Object) {
        $names = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($property in $Element.EnumerateObject()) {
            if (-not $names.Add($property.Name)) { throw "Duplicate DxUi sidecar JSON property '$($property.Name)'." }
            Assert-DxUiUniqueJsonProperties $property.Value
        }
    } elseif ($Element.ValueKind -eq [System.Text.Json.JsonValueKind]::Array) {
        foreach ($value in $Element.EnumerateArray()) { Assert-DxUiUniqueJsonProperties $value }
    }
}

function Get-DxUiSidecar([string]$Side, [string]$Repo) {
    $path = Join-Path $Repo '.build\x64\Release\RedSalamander.exe.DxUi.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing $Side DxUi runtime sidecar: $path" }
    $item = Get-Item -LiteralPath $path -Force
    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "$Side DxUi runtime sidecar cannot be a reparse point: $path" }
    $bytes = [IO.File]::ReadAllBytes($path)
    $document = $null
    try {
        $text = [Text.Encoding]::UTF8.GetString($bytes)
        $document = [System.Text.Json.JsonDocument]::Parse($text)
        Assert-DxUiUniqueJsonProperties $document.RootElement
        $value = $text | ConvertFrom-Json -ErrorAction Stop
    } catch { throw "Invalid $Side DxUi runtime sidecar JSON: $path" }
    finally { if ($null -ne $document) { $document.Dispose() } }
    $expectedKeys = @('repository','commit','apiRevision','platform','configuration','publicHeadersGitTree','buildIdentity','archiveSha256','module','moduleSha256','linkCommandSha256') | Sort-Object -CaseSensitive
    $actualKeys = @($value.PSObject.Properties.Name | Sort-Object -CaseSensitive)
    if (($actualKeys -join [char]0) -cne ($expectedKeys -join [char]0) -or
        (@($value.buildIdentity.PSObject.Properties.Name | Sort-Object -CaseSensitive) -join [char]0) -cne (@('Fingerprint','Identity') -join [char]0)) {
        throw "Incomplete or unsupported $Side DxUi runtime sidecar schema: $path"
    }
    if ($value.repository -isnot [string] -or [string]::IsNullOrWhiteSpace($value.repository) -or
        $value.commit -cnotmatch '^[0-9a-f]{40}$' -or (($value.apiRevision -isnot [int]) -and ($value.apiRevision -isnot [long])) -or $value.apiRevision -le 0 -or
        $value.platform -cne 'x64' -or $value.configuration -cne 'Release' -or
        $value.publicHeadersGitTree -cnotmatch '^[0-9a-f]{40}$' -or $value.archiveSha256 -cnotmatch '^[0-9a-fA-F]{64}$' -or
        $value.module -cne 'RedSalamander.exe' -or $value.moduleSha256 -cnotmatch '^[0-9a-fA-F]{64}$' -or
        $value.linkCommandSha256 -cnotmatch '^[0-9a-fA-F]{64}$' -or
        $null -eq $value.buildIdentity -or $value.buildIdentity.Fingerprint -cnotmatch '^[0-9a-fA-F]{64}$' -or
        $null -eq $value.buildIdentity.Identity) {
        throw "Incomplete or unsupported $Side DxUi runtime sidecar schema/profile: $path"
    }
    $modulePath = Join-Path $Repo '.build\x64\Release\RedSalamander.exe'
    if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf) -or
        (Get-Sha256 $modulePath) -cne ([string]$value.moduleSha256).ToLowerInvariant()) {
        throw "$Side DxUi runtime sidecar does not identify the current RedSalamander.exe bytes: $modulePath"
    }
    return [pscustomobject]@{ side=$Side; path=$path; repo=$Repo; bytes=$bytes; sha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)).ToLowerInvariant(); value=$value }
}

function Assert-DxUiRuntimePair($Pair) {
    $baseline = $Pair.baseline.value; $candidate = $Pair.candidate.value
    foreach ($field in @('repository','commit','apiRevision','platform','configuration','publicHeadersGitTree','archiveSha256','module')) {
        if ([string]$baseline.$field -cne [string]$candidate.$field) { throw "Baseline/candidate DxUi runtime mismatch in '$field'." }
    }
    if ((ConvertTo-DxUiCanonicalJson $baseline.buildIdentity) -cne (ConvertTo-DxUiCanonicalJson $candidate.buildIdentity)) {
        throw 'Baseline/candidate DxUi buildIdentity differs.'
    }
}

function Get-DxUiRuntimePair {
    $pair = [ordered]@{ baseline = Get-DxUiSidecar 'baseline' $BaselineRoot; candidate = Get-DxUiSidecar 'candidate' $CandidateRoot }
    Assert-DxUiRuntimePair $pair
    return $pair
}

function Assert-DxUiRuntimePairUnchanged {
    foreach ($side in @('baseline','candidate')) {
        $pin = $dxuiRuntimePair[$side]
        $current = Get-DxUiSidecar $side $pin.repo
        if ($current.sha256 -cne $pin.sha256 -or $current.bytes.Length -ne $pin.bytes.Length) { throw "$side DxUi runtime sidecar changed during the campaign." }
        $archived = Join-Path $campaignRoot "raw-sources\dxui\$side.DxUi.json"
        if (-not (Test-Path -LiteralPath $archived -PathType Leaf) -or (Get-Sha256 $archived) -cne $pin.sha256 -or
            (Get-Item -LiteralPath $archived).Length -ne $pin.bytes.Length) { throw "Preserved $side DxUi runtime sidecar failed its pinned SHA/length check." }
    }
    Assert-DxUiRuntimePair $dxuiRuntimePair
}

function Get-DxUiManifestIdentity {
    $matched = $dxuiRuntimePair.baseline.value
    $entries = [ordered]@{}
    New-Item -ItemType Directory -Path (Join-Path $campaignRoot 'raw-sources\dxui') -Force | Out-Null
    foreach ($side in @('baseline','candidate')) {
        $pin = $dxuiRuntimePair[$side]
        $relative = "raw-sources/dxui/$side.DxUi.json"
        $destination = Join-Path $campaignRoot ($relative.Replace('/','\'))
        [IO.File]::WriteAllBytes($destination, $pin.bytes)
        if ((Get-Sha256 $destination) -cne $pin.sha256 -or (Get-Item -LiteralPath $destination).Length -ne $pin.bytes.Length) { throw "Could not preserve exact $side DxUi sidecar bytes." }
        $entries[$side] = [ordered]@{ source_path=$pin.path; preserved_path=$relative; bytes=$pin.bytes.Length; sha256=$pin.sha256 }
    }
    return [ordered]@{
        status='matched'; platform=$matched.platform; configuration=$matched.configuration
        dependency=@{ repository=$matched.repository; commit=$matched.commit; apiRevision=$matched.apiRevision; archiveSha256=$matched.archiveSha256; publicHeadersGitTree=$matched.publicHeadersGitTree; buildIdentity=$matched.buildIdentity }
        sidecars=$entries
    }
}

function Get-PinnedBuild([string]$Side, [string]$Repo, [string]$ReceiptId, [string]$SourceSnapshotId, [string]$BuildLogPath) {
    if ([string]::IsNullOrWhiteSpace($ReceiptId) -or [string]::IsNullOrWhiteSpace($SourceSnapshotId) -or
        [string]::IsNullOrWhiteSpace($BuildLogPath) -or -not [IO.Path]::IsPathRooted($BuildLogPath)) {
        throw "Explicit receipt, source_snapshot_id, and exact absolute MSBuild log path are required for $Side."
    }
    $receiptPath = Join-Path $Repo '.build\x64\Release\build-receipt.json'
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { throw "Missing exact Release receipt for ${Side}: $receiptPath" }
    $receiptItem = Get-Item -LiteralPath $receiptPath -Force
    if (($receiptItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Release receipt cannot be a reparse point: $receiptPath" }
    $receipt = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
    if ($receipt.schema -cne 'red-salamander.build-receipt.v1' -or $receipt.configuration -cne 'Release' -or
        $receipt.platform -cne 'x64' -or [int]$receipt.exit_code -ne 0 -or $receipt.tests_enabled -ne $true -or
        $receipt.receipt_id -cne $ReceiptId -or $receipt.source_snapshot_id -cne $SourceSnapshotId) {
        throw "Exact x64 Release receipt/source snapshot mismatch for $Side."
    }
    $logRoot = [IO.Path]::GetFullPath((Join-Path $Repo '.build\logs')).TrimEnd('\') + '\'
    $exactLog = [IO.Path]::GetFullPath($BuildLogPath)
    $logDirectory = Get-Item -LiteralPath $logRoot.TrimEnd('\') -Force -ErrorAction Stop
    if (-not $exactLog.StartsWith($logRoot, [StringComparison]::OrdinalIgnoreCase) -or
        -not (Test-Path -LiteralPath $exactLog -PathType Leaf)) {
        throw "Exact MSBuild log must exist beneath ${logRoot}: $exactLog"
    }
    $logItem = Get-Item -LiteralPath $exactLog -Force
    if (($logDirectory.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
        ($logItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
        $logItem.DirectoryName -cne $logDirectory.FullName.TrimEnd('\') -or $logItem.Length -le 0 -or $logItem.Length -gt 2MB) {
        throw "Exact MSBuild log is linked, empty, or exceeds the 2 MiB evidence bound: $exactLog"
    }
    return [ordered]@{
        side=$Side; repository=$Repo; receipt_id=$ReceiptId; source_snapshot_id=$SourceSnapshotId
        receipt_path=$receiptPath; receipt_sha256=(Get-Sha256 $receiptPath); receipt_bytes=$receiptItem.Length
        build_log_source_path=$exactLog; build_log_sha256=(Get-Sha256 $exactLog); build_log_bytes=$logItem.Length
    }
}

$pinnedBuilds = [ordered]@{
    baseline = (Get-PinnedBuild 'baseline' $BaselineRoot $ExpectedBaselineReceiptId $ExpectedBaselineBuildSourceSnapshotId $ExpectedBaselineBuildLogPath)
    candidate = (Get-PinnedBuild 'candidate' $CandidateRoot $ExpectedCandidateReceiptId $ExpectedCandidateBuildSourceSnapshotId $ExpectedCandidateBuildLogPath)
}
$dxuiRuntimePair = Get-DxUiRuntimePair
$campaignId = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
$campaignRoot = Join-Path $testRootFull "evidence\viewer-space-ordinary-current-$campaignId"
$campaignRoot = Assert-RSTestSandboxContainedPath -TestRoot $testRootFull -Path $campaignRoot -RequireDescendant
if (Test-Path -LiteralPath $campaignRoot) {
    throw "Campaign evidence directory already exists: $campaignRoot"
}
$logsRoot = Join-Path $campaignRoot 'runner-logs'
New-Item -ItemType Directory -Path $logsRoot -Force | Out-Null
$rawSourcesRoot = Join-Path $campaignRoot 'raw-sources'
New-Item -ItemType Directory -Path $rawSourcesRoot -Force | Out-Null

$buildInputsRoot = Join-Path $campaignRoot 'build-inputs'
foreach ($side in @('baseline', 'candidate')) {
    $sideRoot = Join-Path $buildInputsRoot $side
    New-Item -ItemType Directory -Path $sideRoot -Force | Out-Null
    $build = $pinnedBuilds[$side]
    $receiptCopy = Join-Path $sideRoot 'build-receipt.json'
    $logCopy = Join-Path $sideRoot 'msbuild.log'
    [IO.File]::Copy($build.receipt_path, $receiptCopy, $false)
    [IO.File]::Copy($build.build_log_source_path, $logCopy, $false)
    if ((Get-Sha256 $receiptCopy) -cne $build.receipt_sha256 -or (Get-Sha256 $logCopy) -cne $build.build_log_sha256) {
        throw "Pinned build inputs changed while being sealed for $side."
    }
    $build.receipt_archive_path = "build-inputs/$side/build-receipt.json"
    $build.build_log_archive_path = "build-inputs/$side/msbuild.log"
}

$manifestPath = Join-Path $campaignRoot 'campaign-manifest.json'
$proposalRoot = Join-Path $CandidateRoot '.build\codex-proposals\ordinary-current-qualification'
$curatorPath = Join-Path $proposalRoot 'curate-viewer-space-ordinary-current-evidence.py'
if (-not (Test-Path -LiteralPath $curatorPath -PathType Leaf)) {
    throw "Existing ViewerSpace evidence curator is missing: $curatorPath"
}
$python = (Get-Command python -ErrorAction Stop).Source
$pwsh = (Get-Command pwsh -ErrorAction Stop).Source
Import-Module (Join-Path $CandidateRoot 'Tools\Modules\Build\ProcessStreaming.psm1') -Force -Scope Local -ErrorAction Stop

# Preserve the exact ignored orchestration helpers used by this campaign, before
# any measurements or stale-run cleanup. These are evidence, not source inputs.
$reproductionRoot = Join-Path $campaignRoot 'reproduction'
New-Item -ItemType Directory -Path $reproductionRoot | Out-Null
$reproductionInputs = [ordered]@{}
foreach ($helper in @($PSCommandPath, $curatorPath, (Join-Path $proposalRoot 'analyze-viewer-space-ordinary-current-qualification.py'))) {
    $copyPath = Join-Path $reproductionRoot ([IO.Path]::GetFileName($helper))
    [IO.File]::Copy($helper, $copyPath, $false)
    $sourceHash = (Get-FileHash -LiteralPath $helper -Algorithm SHA256).Hash.ToLowerInvariant()
    $copyHash = (Get-FileHash -LiteralPath $copyPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($sourceHash -cne $copyHash) { throw "Campaign helper preservation mismatch: $helper" }
    $reproductionInputs[[IO.Path]::GetFileName($helper)] = [ordered]@{
        path = "reproduction/$([IO.Path]::GetFileName($helper))"
        sha256 = $sourceHash; bytes = (Get-Item -LiteralPath $copyPath).Length
    }
}

$manifest = [ordered]@{
    schema = 'viewer-space.apprelease-comparison-campaign.v1'
    campaign_id = $campaignId
    created_utc = [DateTime]::UtcNow.ToString('o')
    test_root = $testRootFull
    configuration = 'Release'
    platform = 'x64'
    runner_arguments = @('-Suite', 'Commands', '-SkipBuild', '-CaseFilter', 'viewer_space_perf_',
        '-SelfTestRepeat', '1', '-TestRoot', $testRootFull,
        '-Platform', 'x64', '-Configuration', 'Release')
    opt_in_environment = @{ REDSALAMANDER_VIEWERSPACE_LARGE_PERF = '1' }
    baseline_root = $BaselineRoot
    candidate_root = $CandidateRoot
    qualification_scope = 'ordinary-current-x64-release-large-and-small-scan-comparison'
    expected_builds = $pinnedBuilds
    case_filter = 'viewer_space_perf_'
    expected_cases = @('viewer_space_perf_small_progressive', 'viewer_space_perf_large_optin')
    expected_fixture_counts = @{ small = @{ files=10000; directories=121 }; large = @{ files=100000; directories=501 } }
    ordering = 'baseline warmup, candidate warmup, then ABBA x 3 (A=baseline, B=candidate); one fresh runner process per observation; serial only.'
    process_helper = 'Tools/Modules/Build/ProcessStreaming.psm1::Invoke-RSStreamingProcess with contained child process and a 30 minute outer watchdog.'
    launcher_temp_root = 'campaign-specific TEMP/TMP beneath the marked TestSandbox evidence directory; Run-AllTests then switches to its run-owned process-temp directory.'
    observer_policy = 'Both branches aggregate queue diagnostics in producers and emit cumulative UI snapshots at 250 ms plus terminal/teardown. Baseline source includes matched benchmark harness and diagnostic-only aggregation; original scanner/model/layout production algorithms remain unchanged.'
    archive_label_policy = 'campaign-id plus warmup side, or campaign-id plus ABBA block and position; campaign evidence targets fail closed.'
    reproduction_inputs = $reproductionInputs
    dxui_runtime = Get-DxUiManifestIdentity
    measurements = @()
    runs = @()
}

function Write-CampaignManifest {
    $json = $manifest | ConvertTo-Json -Depth 32
    $temporary = "$manifestPath.tmp"
    [IO.File]::WriteAllText($temporary, $json + "`n", [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporary -Destination $manifestPath -Force | Out-Null
}

Write-CampaignManifest

function Assert-PinnedBuildUnchanged([string]$Side) {
    $build = $pinnedBuilds[$Side]
    $receipt = Join-Path $build.repository '.build\x64\Release\build-receipt.json'
    if (-not (Test-Path -LiteralPath $receipt -PathType Leaf) -or (Get-Sha256 $receipt) -cne $build.receipt_sha256) {
        throw "Pinned $Side Release receipt changed during the campaign."
    }
    $log = Get-Item -LiteralPath $build.build_log_source_path -ErrorAction Stop
    if ($log.Length -ne $build.build_log_bytes -or (Get-Sha256 $log.FullName) -cne $build.build_log_sha256) {
        throw "Exact pinned $Side MSBuild log changed during the campaign."
    }
}

function Invoke-Measurement([string]$Side, [string]$Stage, [bool]$IsWarmup, [int]$Block, [int]$Position) {
    Assert-DxUiRuntimePairUnchanged
    $repo = if ($Side -eq 'baseline') { $BaselineRoot } else { $CandidateRoot }
    $runner = Join-Path $repo 'Tools\Run-AllTests.ps1'
    $archiveLabel = if ($IsWarmup) { "$campaignId-$Stage" } else { "$campaignId-b$Block-p$Position" }
    $receiptPath = Join-Path $repo '.build\x64\Release\build-receipt.json'
    $receiptBeforeBytes = [IO.File]::ReadAllBytes($receiptPath)
    $receiptBefore = [Text.Encoding]::UTF8.GetString($receiptBeforeBytes) | ConvertFrom-Json
    $expectedBuild = $pinnedBuilds[$Side]
    if ($receiptBefore.configuration -cne 'Release' -or $receiptBefore.platform -cne 'x64' -or
        [int]$receiptBefore.exit_code -ne 0 -or $receiptBefore.tests_enabled -ne $true -or
        $receiptBefore.receipt_id -cne $expectedBuild.receipt_id -or
        $receiptBefore.source_snapshot_id -cne $expectedBuild.source_snapshot_id) {
        throw "Receipt does not match the explicitly pinned x64 Release identity: $receiptPath"
    }
    Assert-PinnedBuildUnchanged $Side
    $receiptBeforeHash = Get-Sha256 $receiptPath

    $stdoutPath = Join-Path $logsRoot "$archiveLabel.stdout.txt"
    $stderrPath = Join-Path $logsRoot "$archiveLabel.stderr.txt"
    $runTempRoot = Join-Path $campaignRoot "launcher-temp\$archiveLabel"
    $runTempRoot = Assert-RSTestSandboxContainedPath -TestRoot $testRootFull -Path $runTempRoot -RequireDescendant
    New-Item -ItemType Directory -Path $runTempRoot -Force | Out-Null
    $runnerArguments = @('-NoLogo', '-NoProfile', '-File', $runner,
        '-Suite', 'Commands', '-SkipBuild', '-CaseFilter', 'viewer_space_perf_',
        '-SelfTestRepeat', '1', '-TestRoot', $testRootFull,
        '-Platform', 'x64', '-Configuration', 'Release')
    $exitCode = Invoke-RSStreamingProcess `
        -FilePath $pwsh `
        -Arguments $runnerArguments `
        -WorkingDirectory $repo `
        -LogPath $stdoutPath `
        -ErrorLogPath $stderrPath `
        -TimeoutMilliseconds 1800000 `
        -AdditionalEnvironment @{
            REDSALAMANDER_VIEWERSPACE_LARGE_PERF = '1'
            REDSALAMANDER_TEST_ROOT = $testRootFull
            TEMP = $runTempRoot
            TMP = $runTempRoot
        }

    Assert-DxUiRuntimePairUnchanged

    $runnerOutput = (Get-Content -LiteralPath $stdoutPath -Raw) -replace '\x1B\[[0-?]*[ -/]*[@-~]', ''
    $stdoutItem = Get-Item -LiteralPath $stdoutPath -Force
    $stderrItem = Get-Item -LiteralPath $stderrPath -Force
    if ($stdoutItem.Length -gt 2MB -or $stderrItem.Length -gt 2MB -or
        ($stdoutItem.Length + $stderrItem.Length) -gt 5MB) {
        throw "Captured runner logs exceed the 2 MiB per-file or 5 MiB per-run evidence bound for '$archiveLabel'."
    }
    $runIdMatches = [regex]::Matches($runnerOutput, 'Run id:\s*([0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32})')
    if ($runIdMatches.Count -ne 1) {
        throw "Runner '$archiveLabel' did not emit exactly one canonical Run id. Exit=$exitCode; see '$stdoutPath' and '$stderrPath'."
    }
    $runId = $runIdMatches[0].Groups[1].Value
    $runRoot = Join-Path (Join-Path $testRootFull 'runs') $runId
    $raw = Join-Path $runRoot 'artifacts\selftest\last_run'
    $curationExitCode = -1
    $curatedPath = $null
    $curationResult = $null
    if (Test-Path -LiteralPath $raw -PathType Container) {
        $perfMetricsPath = Join-Path $raw 'perf\perf_metrics.jsonl'
        if (-not (Test-Path -LiteralPath $perfMetricsPath -PathType Leaf)) {
            throw "Cannot curate '$archiveLabel' before cleanup: perf metrics are missing from the canonical run artifacts."
        }
        $expectedArchivePath = Join-Path $campaignRoot "curated\$Side\$archiveLabel"
        if (Test-Path -LiteralPath $expectedArchivePath) {
            throw "Refusing to overwrite existing curated archive: $expectedArchivePath"
        }
        $curatorOutput = @(& $python -X utf8 $curatorPath $campaignRoot $runId $archiveLabel $Side `
            ([string]$receiptBefore.receipt_id) ([string]$receiptBefore.source_snapshot_id))
        $curationExitCode = [int]$LASTEXITCODE
        if ($curationExitCode -eq 0 -and $curatorOutput.Count -gt 0) {
            $curationResult = [string]$curatorOutput[-1] | ConvertFrom-Json
            $curatedPath = [string]$curationResult.path
            $curatedPrefix = (Join-Path $campaignRoot 'curated').TrimEnd('\') + '\'
            if (-not $curatedPath.StartsWith($curatedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Curator returned an archive outside campaign evidence: '$curatedPath'"
            }
            if ([IO.Path]::GetFullPath($curatedPath) -cne [IO.Path]::GetFullPath($expectedArchivePath)) {
                throw "Curator path differs from the verified no-overwrite target: '$curatedPath'"
            }
        }
    }
    $summaryPath = Join-Path $raw 'run-all-tests-results.json'
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) {
        throw "Canonical runner summary is missing for '$archiveLabel': $summaryPath"
    }
    $summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
    $summaryHash = Get-Sha256 $summaryPath
    $caseSummary = @($summary.suites | ForEach-Object { $_.cases } | ForEach-Object { $_ })
    $expectedCases = @('viewer_space_perf_small_progressive', 'viewer_space_perf_large_optin')
    $actualNames = @($caseSummary | ForEach-Object { [string]$_.name } | Sort-Object -Unique)
    $caseStatusesValid = ($caseSummary.Count -eq 2 -and
        (@($caseSummary | Where-Object { [string]$_.status -ine 'passed' }).Count -eq 0) -and
        (@($expectedCases | Where-Object { $_ -notin $actualNames }).Count -eq 0))
    $runPassed = $exitCode -eq 0 -and $summary.exit_code -eq 0 -and
        $summary.run_id -ceq $runId -and $summary.platform -ceq 'x64' -and
        $summary.configuration -ceq 'Release' -and $summary.case_filter -ceq 'viewer_space_perf_' -and
        $caseStatusesValid

    $receiptAfterHash = Get-Sha256 $receiptPath
    $receiptAfter = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
    if ($receiptAfterHash -cne $receiptBeforeHash -or $receiptAfter.receipt_id -cne $receiptBefore.receipt_id -or
        $receiptAfter.source_snapshot_id -cne $expectedBuild.source_snapshot_id) {
        throw "Release receipt changed during '$archiveLabel'; measurements are invalid."
    }
    Assert-PinnedBuildUnchanged $Side
    $snapshotPath = Join-Path $raw 'workspace-final.json'
    $workspaceSnapshotId = $null
    $workspaceSnapshotSha256 = $null
    if (Test-Path -LiteralPath $snapshotPath -PathType Leaf) {
        $snapshot = Get-Content -LiteralPath $snapshotPath -Raw | ConvertFrom-Json
        $workspaceSnapshotId = [string]$snapshot.snapshot_id
        $workspaceSnapshotSha256 = Get-Sha256 $snapshotPath
    }

    $runRecord = [ordered]@{
        label = $archiveLabel
        stage = $Stage
        side = $Side
        role = if ($IsWarmup) { 'warmup' } else { 'measurement' }
        block = $Block
        position = $Position
        workspace_root = $repo
        run_id = $runId
        process_exit_code = $exitCode
        summary_exit_code = $summary.exit_code
        summary_schema = $summary.schema
        summary_path = $summaryPath
        summary_sha256 = $summaryHash
        case_filter = $summary.case_filter
        cases = $caseSummary
        case_statuses_valid = $caseStatusesValid
        build_receipt_id = [string]$receiptBefore.receipt_id
        build_receipt_path = $receiptPath
        build_receipt_sha256 = $receiptBeforeHash
        build_source_snapshot_id = [string]$receiptBefore.source_snapshot_id
        build_log_source_path = [string]$expectedBuild.build_log_source_path
        build_log_sha256 = [string]$expectedBuild.build_log_sha256
        build_log_bytes = [long]$expectedBuild.build_log_bytes
        curated_evidence_sha256 = [string]$curationResult.evidenceSha256
        curated_evidence_bytes = [long]$curationResult.evidenceBytes
        curated_provenance_sha256 = [string]$curationResult.provenanceSha256
        curated_provenance_bytes = [long]$curationResult.provenanceBytes
        curated_total_bytes = [long]$curationResult.bytes
        raw_metrics_sha256 = [string]$curationResult.rawMetricsSha256
        workspace_snapshot_id = $workspaceSnapshotId
        workspace_snapshot_sha256 = $workspaceSnapshotSha256
        runner_stdout_path = $stdoutPath
        runner_stdout_sha256 = Get-Sha256 $stdoutPath
        runner_stdout_bytes = [long]$stdoutItem.Length
        runner_stderr_path = $stderrPath
        runner_stderr_sha256 = Get-Sha256 $stderrPath
        runner_stderr_bytes = [long]$stderrItem.Length
        curated_path = $curatedPath
        curation_exit_code = $curationExitCode
        status = if ($runPassed -and $curationExitCode -eq 0 -and $null -ne $curationResult -and
            $curationResult.qualificationStatus -ceq 'qualified' -and $null -ne $curatedPath -and
            (Test-Path -LiteralPath $curatedPath -PathType Container)) { 'PASSED' } else { 'FAILED' }
    }
    $manifest.runs = @($manifest.runs) + @([pscustomobject]$runRecord)
    if (-not $IsWarmup) {
        $manifest.measurements = @($manifest.measurements) + @([pscustomobject]@{
            label = $archiveLabel
            side = $Side
            block = $Block
            position = $Position
            run_id = $runId
        })
    }
    Write-CampaignManifest
    Write-Host "$archiveLabel $($runRecord.status) run=$runId receipt=$($receiptBefore.receipt_id)"
    if ($runRecord.status -ne 'PASSED') {
        throw "Stopping campaign after unsuccessful run '$archiveLabel'. See the manifest and captured artifacts."
    }
}

Invoke-Measurement 'baseline' 'warmup_baseline' $true 0 0
Invoke-Measurement 'candidate' 'warmup_candidate' $true 0 0
for ($block = 1; $block -le 3; ++$block) {
    Invoke-Measurement 'baseline' 'measurement' $false $block 1
    Invoke-Measurement 'candidate' 'measurement' $false $block 2
    Invoke-Measurement 'candidate' 'measurement' $false $block 3
    Invoke-Measurement 'baseline' 'measurement' $false $block 4
}

$manifest['finished_utc'] = [DateTime]::UtcNow.ToString('o')
Write-CampaignManifest
Write-Host "Completed six Release measurements per side. Manifest: $manifestPath"
Write-Host "Analysis command: python -X utf8 `"$(Join-Path $proposalRoot 'analyze-viewer-space-ordinary-current-qualification.py')`" `"$manifestPath`""
