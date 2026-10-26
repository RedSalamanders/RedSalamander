[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$spaceRepo = [IO.Path]::GetFullPath((Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$spaceAuthority = 'D:\RedSalamander.Perf'
if ((Get-Content -Raw -LiteralPath (Join-Path $spaceAuthority '.red-salamander-test-root')).Trim() -cne 'red-salamander.test-sandbox-root.v1') { throw 'Ownership marker mismatch' }
$spaceCampaign = Get-Content -Raw -LiteralPath (Join-Path $spaceAuthority 'evidence\viewer-space-memory-retainedreference-20261009T231508122Z-399289c9\campaign-manifest.json') | ConvertFrom-Json
if (-not $spaceCampaign.finished_utc -or @($spaceCampaign.runs).Count -ne 14 -or @($spaceCampaign.runs | Where-Object status -NE 'PASSED').Count) { throw 'Controlled campaign is not terminal with fourteen passes; no overlapping run' }
$spaceEvidence = Join-Path $spaceAuthority 'evidence\viewer-space-finish-all-20261009\hover-navigation-current'
if (Test-Path -LiteralPath $spaceEvidence) { throw 'No overwrite of previous focused attempt' }
$spaceFrozen = Get-Content -Raw -LiteralPath (Join-Path (Split-Path -Parent $spaceEvidence) 'queued-path-frozen-source.json') | ConvertFrom-Json
foreach ($spacePin in $spaceFrozen.WorkspaceSnapshot.files) {
    $spaceFile = Join-Path $spaceRepo $spacePin.path
    if ((Get-Item -LiteralPath $spaceFile).Length -ne $spacePin.worktree_size_bytes -or (Get-FileHash -LiteralPath $spaceFile -Algorithm SHA256).Hash.ToLowerInvariant() -cne $spacePin.worktree_sha256) { throw "Frozen source changed: $($spacePin.path)" }
}
$spaceReceiptPath = Join-Path $spaceRepo '.build\x64\Release\build-receipt.json'
$spaceReceipt = Get-Content -Raw -LiteralPath $spaceReceiptPath | ConvertFrom-Json
if ($spaceReceipt.receipt_id -cne '37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd' -or $spaceReceipt.source_snapshot_id -cne $spaceFrozen.SnapshotId -or -not $spaceReceipt.tests_enabled) { throw 'Exact test-enabled Release receipt unavailable' }
Import-Module (Join-Path $spaceRepo 'Tools\Modules\Testing\TestSandbox.psm1') -Force -Scope Local
Import-Module (Join-Path $spaceRepo 'Tools\Modules\Build\ProcessStreaming.psm1') -Force -Scope Local
$spaceEvidence = Assert-RSTestSandboxContainedPath -TestRoot $spaceAuthority -Path $spaceEvidence -RequireDescendant
New-Item -ItemType Directory -Path $spaceEvidence | Out-Null
$spaceTemp = Join-Path $spaceEvidence 'temp'
New-Item -ItemType Directory -Path $spaceTemp | Out-Null
$spaceRunId = '{0}-{1}-{2}' -f [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ'), $PID, [Guid]::NewGuid().ToString('N')
$spaceRequest = [ordered]@{runId=$spaceRunId; case='cmd_viewer_space_hover_does_not_rebuild_static_cache'; receiptId=$spaceReceipt.receipt_id; sourceSnapshotId=$spaceReceipt.source_snapshot_id; startedUtc=[DateTime]::UtcNow.ToString('o'); finishedUtc=$null; exitCode=$null; purpose='Recover missing current hover/navigation metric artifacts; no broader verdict replacement'; desktopAgreement='User authorized governed broad Commands focus-taking during this implementation session'}
$spaceRequestPath = Join-Path $spaceEvidence 'request.json'
$spaceRequest | ConvertTo-Json | Set-Content -LiteralPath $spaceRequestPath -Encoding utf8
Copy-Item -LiteralPath $spaceReceiptPath -Destination (Join-Path $spaceEvidence 'build-receipt.json')
$spaceEnv = @{REDSALAMANDER_TEST_ROOT=$spaceAuthority; TEMP=$spaceTemp; TMP=$spaceTemp}
foreach ($spaceVariable in @(Get-ChildItem Env: | Where-Object Name -Like 'REDSALAMANDER_VIEWERSPACE_*')) { $spaceEnv[$spaceVariable.Name] = $null }
$spaceArgs = @('-NoLogo','-NoProfile','-File',(Join-Path $spaceRepo 'Tools\Run-AllTests.ps1'),'-Suite','Commands','-SkipBuild','-RunId',$spaceRunId,'-CaseFilter',$spaceRequest.case,'-SelfTestRepeat','1','-TestRoot',$spaceAuthority,'-Platform','x64','-Configuration','Release')
$spaceExit = Invoke-RSStreamingProcess -FilePath (Get-Command pwsh).Source -Arguments $spaceArgs -WorkingDirectory $spaceRepo -LogPath (Join-Path $spaceEvidence 'runner-stdout.log') -ErrorLogPath (Join-Path $spaceEvidence 'runner-stderr.log') -TimeoutMilliseconds 420000 -AdditionalEnvironment $spaceEnv
$spaceRequest.finishedUtc = [DateTime]::UtcNow.ToString('o')
$spaceRequest.exitCode = $spaceExit
$spaceRequest | ConvertTo-Json | Set-Content -LiteralPath $spaceRequestPath -Encoding utf8
Write-Output ('Focused run {0}; exit {1}; evidence {2}' -f $spaceRunId,$spaceExit,$spaceEvidence)
exit $spaceExit
