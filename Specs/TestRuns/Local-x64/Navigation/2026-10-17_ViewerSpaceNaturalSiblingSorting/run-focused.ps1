param([Parameter(Mandatory=$true)][string]$Phase,[string]$Configuration='Debug',[string]$Case='TestViewerSpaceAnalysisHostVirtualGridAndLifetime')
$ErrorActionPreference='Stop'
$repositoryRoot='Z:\src\RedSalamander'
Import-Module (Join-Path $repositoryRoot 'Tools\Modules\Testing\TestRunPlan.psm1')
Import-Module (Join-Path $repositoryRoot 'Tools\Modules\Build\ArtifactOperationLock.psm1')
Import-Module (Join-Path $repositoryRoot 'Tools\Modules\Build\ProcessStreaming.psm1')
$context=New-RSTestRunContext -RepoRoot $repositoryRoot -TestRootOverride 'Z:\RedSalamander.Perf' -SelfTestRootOverride ''
$phaseRoot=Join-Path $PSScriptRoot $Phase
New-Item -ItemType Directory -Path $phaseRoot -Force | Out-Null
$profileRoot=Join-Path $context.RunRoot 'scratch\profile'
New-Item -ItemType Directory -Path $profileRoot -Force | Out-Null
New-Item -ItemType Directory -Path $context.LegacySelfTestRoot -Force | Out-Null
$outputRoot=Join-Path $repositoryRoot ".build\x64\$Configuration"
$executable=Join-Path $outputRoot 'ViewerPETests.exe'
$plugin=Join-Path $outputRoot 'Plugins\ViewerSpace.dll'
$lock=Enter-RSArtifactOperationLock -RepoRoot $repositoryRoot -Operation 'ViewerSpace natural sibling focused native diagnostic' -Scope @{configuration=$Configuration;platform='x64';suite=$Case}
try {
    $receiptSource=Join-Path $outputRoot 'build-receipt.json'
    $pointer=Get-Content $receiptSource -Raw | ConvertFrom-Json
    $receiptPath=Join-Path $outputRoot "receipts\$($pointer.receipt_id).json"
    Copy-Item -LiteralPath $receiptPath -Destination (Join-Path $phaseRoot 'build-receipt.json')
    $metadata=[ordered]@{scope='Focused native diagnostic; not a canonical aggregate or Full';phase=$Phase;configuration=$Configuration;case=$Case;arguments=@('--no-activate',$Case);run_id=$context.RunId;receipt_id=$pointer.receipt_id;started_utc=[DateTime]::UtcNow.ToString('o');inputs=@(@('Plugins/ViewerSpace/ViewerSpace.cpp','Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp') | ForEach-Object {[pscustomobject]@{path=$_;sha256=(Get-FileHash -LiteralPath (Join-Path $repositoryRoot $_) -Algorithm SHA256).Hash}});binaries=@($executable,$plugin | ForEach-Object {[pscustomobject]@{path=$_;sha256=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash}})}
    $exitCode=Invoke-RSStreamingProcess -FilePath $executable -Arguments @('--no-activate',$Case) -WorkingDirectory $outputRoot -LogPath (Join-Path $phaseRoot 'trace.txt') -ErrorLogPath (Join-Path $phaseRoot 'stderr.txt') -TimeoutMilliseconds 45000 -AdditionalEnvironment @{REDSALAMANDER_TEST_ROOT=$context.TestRoot;REDSALAMANDER_SELFTEST_ROOT=$context.LegacySelfTestRoot;LOCALAPPDATA=$profileRoot;TEMP=$profileRoot;TMP=$profileRoot}
    $metadata.exit_code=$exitCode
    $metadata.ended_utc=[DateTime]::UtcNow.ToString('o')
    $metadata | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $phaseRoot 'result.json') -Encoding utf8
    Write-Host "Focused native exit: $exitCode; phase: $Phase"
} finally {Exit-RSArtifactOperationLock -Lock $lock}
