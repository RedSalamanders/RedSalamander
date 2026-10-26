param(
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'

# This private helper refreshes only the three named documentation PNGs and
# their provenance sidecar. It is preview-only unless -Execute is explicit.
# All source evidence is read from the immutable archive and root review files;
# no source is reconstructed from cleaned run directories.
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$testRoot = Join-Path $repoRoot 'Specs\TestRuns\7d3a1247382a'
$commandsArchive = Join-Path $testRoot 'Validation\space_release_prepush_commands_captures_20261009T213050Z'
$nativeArchive = Join-Path $testRoot 'Validation\space_release_prepush_native_captures_20261009T213050Z'
$buildArchive = Join-Path $testRoot 'Build\space_candidate_release_prepush_reattest_20261009T214710Z'
$evidenceRoot = 'D:\RedSalamander.Perf\evidence\viewer-space-finish-all-20261009'
$commandsReviewPath = Join-Path $evidenceRoot 'root-release-commands-capture-review.json'
$nativeReviewPath = Join-Path $evidenceRoot 'root-release-native-reports-build-and-debug-focus-replay.json'
$durableReviewArchive = Join-Path $testRoot 'Validation\space_finish_all_root_review_20261010'
$durableCommandsReviewPath = Join-Path $durableReviewArchive 'root-release-commands-capture-review.json'
$durableNativeReviewPath = Join-Path $durableReviewArchive 'root-release-native-reports-build-and-debug-focus-replay.json'
$historicalProvenanceArchiveRelative = 'Specs/TestRuns/7d3a1247382a/Validation/space_finish_all_root_review_20261010/root-doc-provenance-before-final.json'
$historicalProvenanceArchive = Join-Path $repoRoot $historicalProvenanceArchiveRelative
$historicalProvenanceSha256 = 'f43758ab71c3695860165c945fea73eafe73cf7aa0095f1b598c3b8f32ef085b'
$commandsRunId = '20261009T213050Z-43264-4bde93ea855e44c885a3277a68e89038'
$receiptId = '37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd'
$sourceSnapshotId = '9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554'
$nativePngRelative = 'viewer-pe/viewer-space-lifecycle-96129062/viewer-space-02-partial-scan-observed-totals-inflight.png'
$nativeProvenanceRelative = 'viewer-pe/viewer-space-lifecycle-96129062/viewer-space-02-partial-scan-observed-totals-inflight.provenance.json'

function Assert-NoReparsePathComponents([string] $Path) {
    $current = [System.IO.Path]::GetFullPath($Path)
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        $attributes = $null
        try {
            $attributes = [System.IO.File]::GetAttributes($current)
        } catch [System.IO.FileNotFoundException] {
            # A not-yet-created destination is allowed; its ancestors are still checked.
        } catch [System.IO.DirectoryNotFoundException] {
            # A missing descendant does not hide any existing ancestor.
        }
        if ($null -ne $attributes -and ($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Refusing reparse-point path component '$current' in '$Path'."
        }
        $parent = [System.IO.Path]::GetDirectoryName($current)
        if ([string]::IsNullOrWhiteSpace($parent) -or
            [string]::Equals($parent, $current, [System.StringComparison]::OrdinalIgnoreCase)) { break }
        $current = $parent
    }
}

function Assert-Equal([object] $Actual, [object] $Expected, [string] $Label) {
    if ($Actual -cne $Expected) {
        throw "$Label mismatch. Expected '$Expected', got '$Actual'."
    }
}

function Assert-ContainedPath([string] $Root, [string] $Path, [string] $Label) {
    $fullRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if (-not $fullPath.StartsWith($fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label escapes its pinned root: $fullPath"
    }
    return $fullPath
}

function Get-ManifestEntry([string] $ManifestPath, [string] $RelativePath, [string] $ExpectedBytes, [string] $ExpectedHash) {
    $matches = @(
        Get-Content -LiteralPath $ManifestPath | ForEach-Object {
            $columns = $_ -split "`t", 4
            if ($columns.Count -eq 4 -and $columns[0] -ceq $RelativePath) {
                [pscustomobject]@{ Path = $columns[0]; Bytes = $columns[1]; Sha256 = $columns[2]; AbsoluteDiagnostic = $columns[3] }
            }
        }
    )
    if ($matches.Count -ne 1) { throw "Expected exactly one manifest record for $RelativePath in $ManifestPath." }
    Assert-Equal $matches[0].Bytes $ExpectedBytes "$RelativePath byte count in manifest"
    Assert-Equal $matches[0].Sha256 $ExpectedHash "$RelativePath SHA-256 in manifest"
    return $matches[0]
}

function Get-VerifiedArchiveFile([string] $ArchiveRoot, [string] $RelativePath, [string] $ExpectedBytes, [string] $ExpectedHash, [string] $ManifestPath) {
    if ($RelativePath.StartsWith('/') -or $RelativePath.Contains('..')) { throw "Unsafe archive-relative path: $RelativePath" }
    $path = Assert-ContainedPath $ArchiveRoot (Join-Path $ArchiveRoot ($RelativePath -replace '/', '\')) 'Archive path'
    Assert-NoReparsePathComponents $path
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing archive payload: $path" }
    $null = Get-ManifestEntry $ManifestPath $RelativePath $ExpectedBytes $ExpectedHash
    $item = Get-Item -LiteralPath $path -Force
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Refusing reparse-point archive payload: $path" }
    if ([string]$item.Length -cne $ExpectedBytes) { throw "Archive payload length mismatch: $path" }
    $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-Equal $hash $ExpectedHash "$RelativePath payload SHA-256"
    return $path
}

function Get-PngDimensions([string] $Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 24 -or $bytes[0] -ne 137 -or $bytes[1] -ne 80 -or $bytes[2] -ne 78 -or $bytes[3] -ne 71) {
        throw "Not a valid PNG header: $Path"
    }
    $width = [System.Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($bytes, 16))
    $height = [System.Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($bytes, 20))
    return @{ width = $width; height = $height }
}

function Find-ReviewPayload($Review, [string] $Path, [string] $Hash, [string] $Bytes, [string] $ContainerProperty, [string] $PathProperty, [string] $HashProperty, [string] $BytesProperty) {
    $matches = @($Review.$ContainerProperty | Where-Object { $_.$PathProperty -ceq $Path })
    if ($matches.Count -ne 1) { throw "Root visual review does not contain exactly one payload record for $Path." }
    Assert-Equal $matches[0].$HashProperty $Hash "Root-review hash for $Path"
    Assert-Equal ([string]$matches[0].$BytesProperty) $Bytes "Root-review byte count for $Path"
}

# Validate repository paths before touching archive payloads or destinations.
Assert-NoReparsePathComponents $repoRoot
foreach ($path in @($testRoot, $commandsArchive, $nativeArchive, $buildArchive, (Join-Path $repoRoot 'docs\res'))) {
    Assert-NoReparsePathComponents $path
}

# Root review inputs live outside the checkout. Use the repository's canonical
# marked-root authorization before reading any record, then constrain each file
# beneath that root; the shared helper performs fresh component checks.
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\TestSandbox.psm1') -Force
$authorizedPerfRoot = Get-RSTestSandboxRoot -RepoRoot $repoRoot -TestRootOverride 'D:\RedSalamander.Perf' -SelfTestRootOverride '' -LocalAppDataRoot ''
if (-not [string]::Equals($authorizedPerfRoot.TrimEnd('\'), 'D:\RedSalamander.Perf', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Unexpected authorized review evidence root: $authorizedPerfRoot"
}
$commandsReviewPath = Assert-RSTestSandboxContainedPath -TestRoot $authorizedPerfRoot -Path $commandsReviewPath -RequireDescendant
$nativeReviewPath = Assert-RSTestSandboxContainedPath -TestRoot $authorizedPerfRoot -Path $nativeReviewPath -RequireDescendant

# Pin current documentary destinations; nothing else under docs/res is eligible.
$destinations = @{
    explorer = Assert-ContainedPath (Join-Path $repoRoot 'docs\res') (Join-Path $repoRoot 'docs\res\viewer-space-explorer.png') 'Explorer destination'
    rainbow = Assert-ContainedPath (Join-Path $repoRoot 'docs\res') (Join-Path $repoRoot 'docs\res\viewer-space-analyst-rainbow.png') 'Rainbow destination'
    inflight = Assert-ContainedPath (Join-Path $repoRoot 'docs\res') (Join-Path $repoRoot 'docs\res\viewer-space-inflight.png') 'In-flight destination'
    provenance = Assert-ContainedPath (Join-Path $repoRoot 'docs\res') (Join-Path $repoRoot 'docs\res\viewer-space-captures.provenance.json') 'Provenance destination'
}
foreach ($destination in $destinations.Values) {
    Assert-NoReparsePathComponents $destination
    if (Test-Path -LiteralPath $destination) {
        $item = Get-Item -LiteralPath $destination -Force
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Refusing reparse-point destination: $destination" }
    }
}

$commandsManifest = Join-Path $commandsArchive 'manifest.tsv'
$commandsCaptureRoot = Join-Path $commandsArchive 'selftest\last_run\perf\viewer_space_captures'
$commandsCaptureRelative = @{
    explorer = 'selftest/last_run/perf/viewer_space_captures/explorer-light.png'
    rainbow = 'selftest/last_run/perf/viewer_space_captures/analyst-rainbow.png'
    provenance = 'selftest/last_run/perf/viewer_space_captures/provenance.json'
}
$files = @{
    explorer = Get-VerifiedArchiveFile $commandsArchive $commandsCaptureRelative.explorer '503887' 'd34452778aa6413f65b8df79446db6c37fa7a2e707aca9c6c3b0618f87b416d7' $commandsManifest
    rainbow = Get-VerifiedArchiveFile $commandsArchive $commandsCaptureRelative.rainbow '383793' '3960a081a380ee9708c3e9cc7564542e9a811bc0ffaf96066abda089c3bdf41b' $commandsManifest
    commandsProvenance = Get-VerifiedArchiveFile $commandsArchive $commandsCaptureRelative.provenance '4093' '8b5f5d7b6561baa9cf0a46f9566d6c9188789e0af6a11bccff704eccaaaf56da' $commandsManifest
}
$nativeManifest = Join-Path $nativeArchive 'manifest.tsv'
$files.native = Get-VerifiedArchiveFile $nativeArchive $nativePngRelative '107897' '828e81d84b269d22079bb8a179b130dabac72d1f8fe284d30aa587b5466c66c1' $nativeManifest
$files.nativeProvenance = Get-VerifiedArchiveFile $nativeArchive $nativeProvenanceRelative '2033' '80968baca7f16da588182082fe0ccbb84d2375e1cc1bcf392b08d22275e5e547' $nativeManifest

$commandsPairPath = Join-Path $commandsArchive 'capture-pairs.json'
$commandsPairs = Get-Content -LiteralPath $commandsPairPath -Raw | ConvertFrom-Json
Assert-Equal $commandsPairs.scope 'commands-captures' 'Commands capture-pairs scope'
Assert-Equal $commandsPairs.captureCount 11 'Commands capture-pairs count'
Assert-Equal $commandsPairs.runId $commandsRunId 'Commands capture-pairs run id'

$commandsProvenance = Get-Content -LiteralPath $files.commandsProvenance -Raw | ConvertFrom-Json
Assert-Equal $commandsProvenance.scenario 'cmd_viewer_space_explorer_analyst_exact_nested_detail' 'Commands scenario'
Assert-Equal $commandsProvenance.captureKind 'actual-application-harness' 'Commands capture kind'
Assert-Equal $commandsProvenance.dpi 144 'Commands capture DPI'
Assert-Equal $commandsProvenance.noActivation $true 'Commands no-activation flag'
Assert-Equal $commandsProvenance.logicalBytes 67108864 'Commands fixture logical bytes'
Assert-Equal $commandsProvenance.folders 10 'Commands fixture folder count'
Assert-Equal $commandsProvenance.files 64 'Commands fixture file count'
$explorerFrame = @($commandsProvenance.frames | Where-Object file -CEQ 'explorer-light.png')
$rainbowFrame = @($commandsProvenance.frames | Where-Object file -CEQ 'analyst-rainbow.png')
if ($explorerFrame.Count -ne 1 -or $rainbowFrame.Count -ne 1) { throw 'Commands provenance is missing a unique selected frame.' }
foreach ($selected in @(
    @{ path = $commandsCaptureRelative.explorer; bytes = 503887; hash = 'd34452778aa6413f65b8df79446db6c37fa7a2e707aca9c6c3b0618f87b416d7' },
    @{ path = $commandsCaptureRelative.rainbow; bytes = 383793; hash = '3960a081a380ee9708c3e9cc7564542e9a811bc0ffaf96066abda089c3bdf41b' }
)) {
    $pair = @($commandsPairs.captures | Where-Object { $_.png.path -ceq $selected.path })
    if ($pair.Count -ne 1) { throw "Commands capture-pairs is missing $($selected.path)." }
    Assert-Equal $pair[0].png.sha256 $selected.hash "Commands capture-pairs hash for $($selected.path)"
    Assert-Equal ([string]$pair[0].png.bytes) ([string]$selected.bytes) "Commands capture-pairs bytes for $($selected.path)"
    Assert-Equal $pair[0].provenance[0].path $commandsCaptureRelative.provenance "Shared Commands provenance path for $($selected.path)"
    Assert-Equal $pair[0].provenance[0].sha256 '8b5f5d7b6561baa9cf0a46f9566d6c9188789e0af6a11bccff704eccaaaf56da' "Shared Commands provenance hash for $($selected.path)"
}

$nativePairs = Get-Content -LiteralPath (Join-Path $nativeArchive 'capture-pairs.json') -Raw | ConvertFrom-Json
Assert-Equal $nativePairs.scope 'native-captures' 'Native capture-pairs scope'
Assert-Equal $nativePairs.runId $commandsRunId 'Native capture-pairs run id'
$nativePair = @($nativePairs.captures | Where-Object { $_.png.path -ceq $nativePngRelative })
if ($nativePair.Count -ne 1) { throw 'Native capture-pairs must bind exactly the selected 02 frame.' }
Assert-Equal $nativePair[0].png.sha256 '828e81d84b269d22079bb8a179b130dabac72d1f8fe284d30aa587b5466c66c1' 'Native capture-pairs PNG hash'
Assert-Equal $nativePair[0].provenance[0].path $nativeProvenanceRelative 'Native adjacent provenance path'
Assert-Equal $nativePair[0].provenance[0].sha256 '80968baca7f16da588182082fe0ccbb84d2375e1cc1bcf392b08d22275e5e547' 'Native capture-pairs provenance hash'
$nativeProvenance = Get-Content -LiteralPath $files.nativeProvenance -Raw | ConvertFrom-Json
Assert-Equal $nativeProvenance.valid $true 'Native provenance validity'
Assert-Equal $nativeProvenance.capture_hresult 0 'Native capture HRESULT'
Assert-Equal $nativeProvenance.foreground_preserved $true 'Native foreground preservation'
Assert-Equal $nativeProvenance.dpi 144 'Native capture DPI'
Assert-Equal $nativeProvenance.capture_route 'Windows.Graphics.Capture + DirectComposition, no activation' 'Native capture route'
Assert-Equal $nativeProvenance.root_total_bytes_after 384 'Native observed logical bytes'
Assert-Equal $nativeProvenance.scanned_folders_after 1 'Native observed folders'
Assert-Equal $nativeProvenance.scanned_files_after 2 'Native observed files'

$receiptPath = Join-Path $buildArchive 'build-receipt.json'
Assert-NoReparsePathComponents $receiptPath
$receipt = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
Assert-Equal $receipt.receipt_id $receiptId 'Build receipt id'
Assert-Equal $receipt.source_snapshot_id $sourceSnapshotId 'Build source snapshot id'
Assert-Equal $receipt.platform 'x64' 'Build platform'
Assert-Equal $receipt.configuration 'Release' 'Build configuration'
Assert-Equal $receipt.tests_enabled $true 'Build tests-enabled flag'
$receiptSha = (Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash.ToLowerInvariant()
$receiptExe = @($receipt.outputs | Where-Object relative_path -CEQ '.build/x64/Release/RedSalamander.exe')
$receiptDll = @($receipt.outputs | Where-Object relative_path -CEQ '.build/x64/Release/Plugins/ViewerSpace.dll')
$receiptViewerTests = @($receipt.outputs | Where-Object relative_path -CEQ '.build/x64/Release/ViewerPETests.exe')
if ($receiptExe.Count -ne 1 -or $receiptDll.Count -ne 1 -or $receiptViewerTests.Count -ne 1) { throw 'Immutable build receipt lacks unique app/ViewerSpace/test outputs.' }
Assert-Equal $nativeProvenance.executable_sha256 $receiptViewerTests[0].sha256 'Native executable hash against immutable receipt'
Assert-Equal $nativeProvenance.viewer_space_dll_sha256 $receiptDll[0].sha256 'Native ViewerSpace hash against immutable receipt'

$commandsReview = Get-Content -LiteralPath $commandsReviewPath -Raw | ConvertFrom-Json
Assert-Equal $commandsReview.reviewer 'root' 'Commands visual reviewer'
Assert-Equal $commandsReview.run_id $commandsRunId 'Commands reviewed run id'
Assert-Equal $commandsReview.receipt_id $receiptId 'Commands reviewed receipt id'
Assert-Equal $commandsReview.build_source_snapshot_id $sourceSnapshotId 'Commands reviewed build source'
Assert-Equal $commandsReview.capture_provenance_sha256 '8b5f5d7b6561baa9cf0a46f9566d6c9188789e0af6a11bccff704eccaaaf56da' 'Commands reviewed provenance hash'
if ($commandsReview.scope -cne 'whole-frame visual inspection of eleven original-resolution harness frames') { throw 'Unexpected Commands review scope.' }
foreach ($pair in @(
    @{ path = 'selftest/last_run/perf/viewer_space_captures/explorer-light.png'; hash = 'd34452778aa6413f65b8df79446db6c37fa7a2e707aca9c6c3b0618f87b416d7'; bytes = '503887' },
    @{ path = 'selftest/last_run/perf/viewer_space_captures/analyst-rainbow.png'; hash = '3960a081a380ee9708c3e9cc7564542e9a811bc0ffaf96066abda089c3bdf41b'; bytes = '383793' }
)) {
    Find-ReviewPayload $commandsReview $pair.path $pair.hash $pair.bytes 'payload_replay' 'path' 'sha256' 'bytes'
}
foreach ($binary in @(
    @{ path = '.build/x64/Release/RedSalamander.exe'; hash = $receiptExe[0].sha256; bytes = [string]$receiptExe[0].size_bytes },
    @{ path = '.build/x64/Release/Plugins/ViewerSpace.dll'; hash = $receiptDll[0].sha256; bytes = [string]$receiptDll[0].size_bytes }
)) {
    Find-ReviewPayload $commandsReview $binary.path $binary.hash $binary.bytes 'binaries' 'path' 'sha256' 'bytes'
}

$nativeReview = Get-Content -LiteralPath $nativeReviewPath -Raw | ConvertFrom-Json
Assert-Equal $nativeReview.reviewer 'root' 'Native visual reviewer'
$nativeReviewFrame = @($nativeReview.lifecycle_frames_reviewed_at_original_resolution | Where-Object {
    ($_.png -replace '\\', '/') -ceq $nativePngRelative
})
if ($nativeReviewFrame.Count -ne 1) { throw 'Root native review does not contain the selected exact 02 frame.' }
Assert-Equal $nativeReviewFrame[0].png_sha256 '828e81d84b269d22079bb8a179b130dabac72d1f8fe284d30aa587b5466c66c1' 'Native root-review PNG hash'
Assert-Equal $nativeReviewFrame[0].provenance_sha256 '80968baca7f16da588182082fe0ccbb84d2375e1cc1bcf392b08d22275e5e547' 'Native root-review provenance hash'
$nativeArchiveRecord = @($nativeReview.archives | Where-Object archive -CEQ 'Validation/space_release_prepush_native_captures_20261009T213050Z')
if ($nativeArchiveRecord.Count -ne 1) { throw 'Root review does not identify the selected immutable native capture archive.' }
if ($Execute) {
    foreach ($pair in @(
        @{ original = $commandsReviewPath; archived = $durableCommandsReviewPath },
        @{ original = $nativeReviewPath; archived = $durableNativeReviewPath }
    )) {
        Assert-NoReparsePathComponents $pair.archived
        if (-not (Test-Path -LiteralPath $pair.archived -PathType Leaf)) { throw "Root's durable review-record archive is required before -Execute: $($pair.archived)" }
        $originalHash = (Get-FileHash -LiteralPath $pair.original -Algorithm SHA256).Hash.ToLowerInvariant()
        $archivedHash = (Get-FileHash -LiteralPath $pair.archived -Algorithm SHA256).Hash.ToLowerInvariant()
        Assert-Equal $archivedHash $originalHash "Durable root-review archive binding for $($pair.archived)"
    }
}

# Validate the report archive context too, without claiming its overall run passed.
$reportsArchive = Join-Path $testRoot 'Validation\space_release_prepush_reports_20261009T213050Z'
$runResultsPath = Join-Path $reportsArchive 'selftest\last_run\run-all-tests-results.json'
Assert-NoReparsePathComponents $reportsArchive
Assert-NoReparsePathComponents $runResultsPath
$runResults = Get-Content -LiteralPath $runResultsPath -Raw | ConvertFrom-Json
Assert-Equal $runResults.build_receipt_id $receiptId 'Release PrePush test-results receipt binding'
Assert-Equal $runResults.run_id $commandsRunId 'Release PrePush test-results run binding'
Assert-Equal $runResults.schema 'red-salamander.run-all-tests.v2' 'Release PrePush results schema'
Assert-Equal $runResults.repository_verdict 'FAILED' 'Release PrePush repository verdict'
Assert-Equal $runResults.change_verdict 'FAILED' 'Release PrePush change verdict'
Assert-Equal $runResults.case_counts.passed 2125 'Release PrePush passed count'
Assert-Equal $runResults.case_counts.failed 16 'Release PrePush failed count'
Assert-Equal $runResults.case_counts.skipped 264 'Release PrePush skipped count'
Assert-Equal $runResults.case_counts.total 2405 'Release PrePush total count'
if ($runResults.entries.Count -ne 20 -or $runResults.counts.promoted -ne 15 -or $runResults.counts.incomplete -ne 5) {
    throw 'Expected the broader Release PrePush report to record its failed overall outcome.'
}

$destMap = @{
    explorer = $destinations.explorer
    rainbow = $destinations.rainbow
    inflight = $destinations.inflight
}
$sourceMap = @{
    explorer = $files.explorer
    rainbow = $files.rainbow
    inflight = $files.native
}
$assetMetadata = @{}
foreach ($key in @('explorer', 'rainbow', 'inflight')) {
    $src = $sourceMap[$key]
    $hash = (Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash.ToLowerInvariant()
    $dim = Get-PngDimensions $src
    $assetMetadata[$key] = @{
        destination = $destMap[$key].Substring($repoRoot.Length + 1).Replace('\', '/')
        archiveSource = if ($key -eq 'inflight') { "Specs/TestRuns/7d3a1247382a/Validation/space_release_prepush_native_captures_20261009T213050Z/$nativePngRelative" } else { "Specs/TestRuns/7d3a1247382a/Validation/space_release_prepush_commands_captures_20261009T213050Z/$($commandsCaptureRelative[$key])" }
        sourceRunId = $commandsRunId
        sha256 = $hash
        bytes = (Get-Item -LiteralPath $src).Length
        pngPixelDimensions = $dim
    }
}

$provenancePath = $destinations.provenance
Assert-NoReparsePathComponents $provenancePath
$currentProvenance = Get-Content -LiteralPath $provenancePath -Raw | ConvertFrom-Json -AsHashtable
if ($Execute) {
    $preRefreshHash = (Get-FileHash -LiteralPath $provenancePath -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-Equal $preRefreshHash $historicalProvenanceSha256 'Pre-refresh documentation provenance SHA-256'
    Assert-NoReparsePathComponents $historicalProvenanceArchive
    if (-not (Test-Path -LiteralPath $historicalProvenanceArchive -PathType Leaf)) {
        throw "Root's byte-preserved pre-refresh provenance archive is required before -Execute: $historicalProvenanceArchive"
    }
    $archivedOldHash = (Get-FileHash -LiteralPath $historicalProvenanceArchive -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-Equal $archivedOldHash $historicalProvenanceSha256 'Archived pre-refresh provenance SHA-256'
}
if (-not $currentProvenance.ContainsKey('historicalCaptures')) { $currentProvenance['historicalCaptures'] = @() }
$oldExplorerRun = [string]$currentProvenance.releaseExplorerAnalystCapture.sourceRunId
$oldNativeRun = [string]$currentProvenance.nativeReleaseInflightCapture.sourceRunId
$oldHistoryKey = "$oldExplorerRun|$oldNativeRun"
$alreadyArchived = @($currentProvenance.historicalCaptures | Where-Object historyKey -CEQ $oldHistoryKey).Count -gt 0
if (($oldExplorerRun -or $oldNativeRun) -and -not $alreadyArchived) {
    $currentProvenance.historicalCaptures += @{
        historyKey = $oldHistoryKey
        preservedAtRefresh = '2026-10-10'
        releaseExplorerAnalystCapture = $currentProvenance.releaseExplorerAnalystCapture
        nativeReleaseInflightCapture = $currentProvenance.nativeReleaseInflightCapture
        qualificationLimits = $currentProvenance.qualificationLimits
        originalSidecarArchive = $historicalProvenanceArchiveRelative
        originalSidecarSha256 = $historicalProvenanceSha256
        note = 'Previous provenance records retained unchanged as historical archive references; cleaned source runs were not reconstructed.'
    }
}
$currentProvenance['priorProvenanceArchive'] = @{
    path = $historicalProvenanceArchiveRelative
    sha256 = $historicalProvenanceSha256
    note = 'Root-created byte-preserved snapshot of the complete pre-refresh sidecar; included to retain the original document provenance exactly.'
}

$commandsCaptureRecord = @{
    scenario = 'cmd_viewer_space_explorer_analyst_exact_nested_detail'
    caseResult = 'passed'
    sourceRunId = $commandsRunId
    sourceWorkspaceSnapshotId = $sourceSnapshotId
    parentCommandsRunContext = 'The Release PrePush run failed overall (2125 passed, 16 failed, 264 skipped of 2405). Its Commands suite was 938 passed, 8 failed, 4 skipped. The scoped ViewerSpace Commands cases were 34 passed, 0 failed, 4 declared opt-in skipped, including successful 20K and 50K cases. These capture results do not imply a green repository run.'
    captureAcceptance = 'Root accepted all 11 complete original-resolution frames in this exact 144-DPI Release capture set. This exact set includes current presented generations for the 503-DIP full-height map and 455-DIP body boundary. It does not qualify other DPI, locale, provider, platform, performance or manual accessibility profiles.'
    captureReview = 'Specs/TestRuns/7d3a1247382a/Validation/space_finish_all_root_review_20261010/root-release-commands-capture-review.json'
    build = @{
        configuration = 'Release'; platform = 'x64'; testsEnabled = $true
        receiptId = $receiptId; immutableReceiptSha256 = $receiptSha
        buildSnapshotId = $sourceSnapshotId
        immutableReceipt = 'Specs/TestRuns/7d3a1247382a/Build/space_candidate_release_prepush_reattest_20261009T214710Z/build-receipt.json'
        applicationSha256 = $receiptExe[0].sha256
        viewerSpaceDllSha256 = $receiptDll[0].sha256
    }
    captureProvenance = 'Specs/TestRuns/7d3a1247382a/Validation/space_release_prepush_commands_captures_20261009T213050Z/selftest/last_run/perf/viewer_space_captures/provenance.json'
    captureProvenanceSha256 = '8b5f5d7b6561baa9cf0a46f9566d6c9188789e0af6a11bccff704eccaaaf56da'
    dpi = 144
    fixture = @{
        scenario = $commandsProvenance.scenario
        logicalBytes = $commandsProvenance.logicalBytes
        folders = $commandsProvenance.folders
        files = $commandsProvenance.files
        noActivation = $commandsProvenance.noActivation
        source = 'Archive shared sibling provenance.json; immutable archive manifest and root-reviewed capture record.'
    }
    captures = @(
        @{
            viewMode = 'Explorer'; file = 'viewer-space-explorer.png'; theme = 'Light'
            source = $assetMetadata.explorer.archiveSource
            sha256 = $assetMetadata.explorer.sha256; bytes = $assetMetadata.explorer.bytes
            pngPixelDimensions = $assetMetadata.explorer.pngPixelDimensions
            harnessClientViewport = @{ widthPx = $explorerFrame[0].widthPx; heightPx = $explorerFrame[0].heightPx; dpi = $commandsProvenance.dpi }
            captureManifestRecord = "Specs/TestRuns/7d3a1247382a/Validation/space_release_prepush_commands_captures_20261009T213050Z/$($commandsCaptureRelative.provenance)#/frames[file=explorer-light.png]"
            wholeFrameReview = 'accepted-for-exact-144dpi-complete-frame'
        },
        @{
            viewMode = 'Analyst'; file = 'viewer-space-analyst-rainbow.png'; theme = 'Rainbow'
            source = $assetMetadata.rainbow.archiveSource
            sha256 = $assetMetadata.rainbow.sha256; bytes = $assetMetadata.rainbow.bytes
            pngPixelDimensions = $assetMetadata.rainbow.pngPixelDimensions
            harnessClientViewport = @{ widthPx = $rainbowFrame[0].widthPx; heightPx = $rainbowFrame[0].heightPx; dpi = $commandsProvenance.dpi }
            captureManifestRecord = "Specs/TestRuns/7d3a1247382a/Validation/space_release_prepush_commands_captures_20261009T213050Z/$($commandsCaptureRelative.provenance)#/frames[file=analyst-rainbow.png]"
            wholeFrameReview = 'accepted-for-exact-144dpi-complete-frame'
        }
    )
}

$nativeCaptureRecord = @{
    scenario = '02-partial-scan-observed-totals-inflight'
    captureResult = 'valid; root-reviewed at original resolution'
    sourceRunId = $commandsRunId
    source = "Specs/TestRuns/7d3a1247382a/Validation/space_release_prepush_native_captures_20261009T213050Z/$nativePngRelative"
    sha256 = $assetMetadata.inflight.sha256
    bytes = $assetMetadata.inflight.bytes
    pngPixelDimensions = $assetMetadata.inflight.pngPixelDimensions
    dpi = $nativeProvenance.dpi
    build = @{
        configuration = 'Release'; platform = 'x64'; testsEnabled = $true
        receiptId = $receiptId; immutableReceiptSha256 = $receiptSha
        buildSnapshotId = $sourceSnapshotId
        immutableReceipt = 'Specs/TestRuns/7d3a1247382a/Build/space_candidate_release_prepush_reattest_20261009T214710Z/build-receipt.json'
        executableSha256 = $nativeProvenance.executable_sha256
        viewerSpaceDllSha256 = $nativeProvenance.viewer_space_dll_sha256
    }
    captureProvenance = "Specs/TestRuns/7d3a1247382a/Validation/space_release_prepush_native_captures_20261009T213050Z/$nativeProvenanceRelative"
    captureProvenanceSha256 = '80968baca7f16da588182082fe0ccbb84d2375e1cc1bcf392b08d22275e5e547'
    captureAcceptance = 'Root reviewed this provenance-bearing 02 partial-scan frame at original resolution as part of the 144-DPI native lifecycle review. The separate additional in-flight image has only limited text provenance and is excluded.'
    wholeFrameReview = @{
        reviewer = 'root'
        source = 'Specs/TestRuns/7d3a1247382a/Validation/space_finish_all_root_review_20261010/root-release-native-reports-build-and-debug-focus-replay.json'
        scenario = '02-partial-scan-observed-totals-inflight'
        status = 'accepted'
    }
    captureRoute = $nativeProvenance.capture_route
    observedDuringCapture = @{
        analysisRows = $nativeProvenance.analysis_rows_after
        files = $nativeProvenance.scanned_files_after
        logicalBytes = $nativeProvenance.root_total_bytes_after
        folders = $nativeProvenance.scanned_folders_after
        lastPresentHResult = $nativeProvenance.last_render_present_hr_after
        activeWorkerCallbacks = $nativeProvenance.gate_callbacks_in_flight_after
    }
    adjacentProvenancePolicy = 'Copied with the original source archive and SHA-256 recorded above; not synthesized.'
}

$currentProvenance['releaseExplorerAnalystCapture'] = $commandsCaptureRecord
$currentProvenance['nativeReleaseInflightCapture'] = $nativeCaptureRecord
$currentProvenance['qualificationLimits'] = @(
    'The Release PrePush run that supplied these captures failed overall (2125 passed, 16 failed, 264 skipped of 2405); its Commands suite also had failures. The selected ViewerSpace Commands scope passed 34 cases and skipped 4 declared opt-ins, including passing 20K and 50K cases. These scoped results do not replace the failed repository verdict.',
    'Root accepted the exact 11 complete Commands frames and the proper 144-DPI lifecycle frames only for the documented deterministic English scenarios. The selected in-flight image has valid adjacent provenance. An additional in-flight image has limited text provenance and is excluded.',
    'This evidence does not establish an overall performance gain, a green pull request, app ARM64 runtime qualification, other DPI or locale/provider profiles, or manual accessibility qualification.'
)
$newProvenanceText = ($currentProvenance | ConvertTo-Json -Depth 40) + "`n"

$receiptRelative = 'Specs/TestRuns/7d3a1247382a/Build/space_candidate_release_prepush_reattest_20261009T214710Z/build-receipt.json'
$manifestRows = @(
    foreach ($key in @('explorer', 'rainbow', 'inflight')) {
        [pscustomobject]@{
            destination = $destMap[$key].Substring($repoRoot.Length + 1).Replace('\', '/')
            source = $assetMetadata[$key].archiveSource
            sha256 = $assetMetadata[$key].sha256
            bytes = $assetMetadata[$key].bytes
            pngPixelDimensions = '{0}x{1}' -f $assetMetadata[$key].pngPixelDimensions.width, $assetMetadata[$key].pngPixelDimensions.height
        }
    }
)

Write-Output "Mode: $(if ($Execute) { 'EXECUTE (explicit)' } else { 'PREVIEW ONLY' })"
Write-Output "Release run: $commandsRunId"
Write-Output "Build receipt: $receiptId (SHA-256 $receiptSha), source $sourceSnapshotId"
Write-Output 'Validated root reviews: Commands 11 original-resolution frames; native exact 02 frame with adjacent provenance.'
foreach ($row in $manifestRows) {
    Write-Output ('{0} <- {1} | SHA-256 {2} | {3} bytes | PNG {4}' -f $row.destination, $row.source, $row.sha256, $row.bytes, $row.pngPixelDimensions)
}
Write-Output "The provenance update retains prior records under historicalCaptures and references root's byte-preserved sidecar archive $historicalProvenanceArchiveRelative (SHA-256 $historicalProvenanceSha256)."
Write-Output 'It records the broader Release PrePush failure and adds no overall-green or performance claim.'
if (-not $Execute) { return }

# All inputs, bindings, hashes, image metadata and destinations have been
# validated above. Only the three pinned PNG destinations and provenance JSON
# are written. PNGs are copied byte-for-byte from immutable archives.
foreach ($key in @('explorer', 'rainbow', 'inflight')) {
    [System.IO.File]::Copy($sourceMap[$key], $destMap[$key], $true)
}
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($provenancePath, $newProvenanceText, $utf8NoBom)
foreach ($key in @('explorer', 'rainbow', 'inflight')) {
    $actualHash = (Get-FileHash -LiteralPath $destMap[$key] -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-Equal $actualHash $assetMetadata[$key].sha256 "Written destination $key SHA-256"
}
Write-Output 'Executed byte copies and provenance update; all three destination PNG hashes match their immutable archive sources.'
