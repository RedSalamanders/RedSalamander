[CmdletBinding()]
param([switch]$Execute)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$spaceRepo = [IO.Path]::GetFullPath((Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$spaceAuthority = 'D:\RedSalamander.Perf'
if ((Get-Content -Raw -LiteralPath (Join-Path $spaceAuthority '.red-salamander-test-root')).Trim() -cne 'red-salamander.test-sandbox-root.v1') { throw 'Ownership marker mismatch' }
$spaceEvidence = Join-Path $spaceAuthority 'evidence\viewer-space-finish-all-20261009'
$spaceParent = [IO.Path]::GetFullPath((Join-Path $spaceRepo 'Specs\TestRuns\7d3a1247382a\Viewers'))
$spacePrefix = 'ordinary-current-224729245z-4221bb43'
$spaceDestination = [IO.Path]::GetFullPath((Join-Path $spaceEvidence 'rejected-ordinary-v1-curation'))
if (-not $spaceDestination.StartsWith($spaceAuthority + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Destination containment failure' }
if (Test-Path -LiteralPath $spaceDestination) { throw 'Preservation destination already exists; no overwrite' }

function Assert-SpaceOrdinaryAncestors([string]$Path) {
    $spaceCursor = [IO.Path]::GetFullPath($Path)
    while (-not [string]::IsNullOrEmpty($spaceCursor)) {
        if (Test-Path -LiteralPath $spaceCursor) {
            if (((Get-Item -Force -LiteralPath $spaceCursor).Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Reparse ancestor: $spaceCursor" }
        }
        $spaceCursor = Split-Path -Parent $spaceCursor
    }
}
Assert-SpaceOrdinaryAncestors $spaceParent
Assert-SpaceOrdinaryAncestors $spaceDestination
$spaceOverview = Join-Path $spaceParent ($spacePrefix + '-overview')
$spaceIndex = Get-Content -Raw -LiteralPath (Join-Path $spaceOverview 'archive-index.json') | ConvertFrom-Json
$spaceNames = @($spacePrefix + '-overview') + @($spaceIndex.sibling_archive_folders)
if ($spaceNames.Count -ne 15 -or @($spaceNames | Sort-Object -Unique).Count -ne 15) { throw 'Expected exactly 15 distinct v1 archives' }
$spaceFolders = @()
foreach ($spaceName in $spaceNames) {
    if ($spaceName -notmatch '^ordinary-current-224729245z-4221bb43-(overview|warmup-(baseline|candidate)|b0[1-3]-p0[1-4]-(baseline|candidate))$') { throw 'Unexpected archive name' }
    $spaceSource = [IO.Path]::GetFullPath((Join-Path $spaceParent $spaceName))
    if ((Split-Path -Parent $spaceSource) -cne $spaceParent) { throw 'Source containment failure' }
    Assert-SpaceOrdinaryAncestors $spaceSource
    $spaceRelative = 'Specs/TestRuns/7d3a1247382a/Viewers/' + $spaceName + '/'
    $spaceTracked = @(& git -C $spaceRepo ls-files -- $spaceRelative)
    if ($LASTEXITCODE -ne 0 -or $spaceTracked.Count -ne 0) { throw 'Rejected archive must be untracked' }
    $spaceLeafIndex = Get-Content -Raw -LiteralPath (Join-Path $spaceSource 'archive-index.json') | ConvertFrom-Json
    if ($spaceLeafIndex.campaign_id -cne '20261009T224729245Z-4221bb43' -or $spaceLeafIndex.archive_folder -cne $spaceName -or $spaceLeafIndex.campaign_manifest_sha256 -cne '5e16e6a0e8ca43d226e516cfa4df4c8a636e2b52b3465173e6d0cbc68adb5eba') { throw 'Archive ownership mismatch' }
    $spaceFiles = @()
    $spacePending = [Collections.Generic.Stack[string]]::new()
    $spacePending.Push($spaceSource)
    while ($spacePending.Count -gt 0) {
        foreach ($spaceItem in Get-ChildItem -Force -LiteralPath $spacePending.Pop()) {
            if (($spaceItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Reparse item in rejected archive' }
            if ($spaceItem.PSIsContainer) { $spacePending.Push($spaceItem.FullName); continue }
            $spaceFiles += [pscustomobject]@{ path = $spaceItem.FullName.Substring($spaceSource.Length + 1); bytes = $spaceItem.Length; sha256 = (Get-FileHash -LiteralPath $spaceItem.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }
        }
    }
    $spaceFolders += [pscustomobject]@{source = $spaceSource; destination = Join-Path $spaceDestination $spaceName; files = $spaceFiles}
}
if (-not $Execute) { [pscustomobject]@{mode='PREVIEW_ONLY'; folders=$spaceFolders.Count; destination=$spaceDestination} | ConvertTo-Json; return }
$spaceTerminal = Get-Content -Raw -LiteralPath (Join-Path $spaceAuthority 'evidence\viewer-space-memory-retainedreference-20261009T231508122Z-399289c9\campaign-manifest.json') | ConvertFrom-Json
if (-not $spaceTerminal.finished_utc -and -not @($spaceTerminal.runs | Where-Object status -EQ 'FAILED').Count) { throw 'Defer filesystem preservation until the native campaign is terminal' }
$null = New-Item -ItemType Directory -Path $spaceDestination
foreach ($spaceFolder in $spaceFolders) {
    if (Test-Path -LiteralPath $spaceFolder.destination) { throw 'Destination appeared during preservation' }
    Assert-SpaceOrdinaryAncestors $spaceFolder.source
    Assert-SpaceOrdinaryAncestors $spaceFolder.destination
    # Exact absolute source/destination and every no-follow child were verified above.
    Move-Item -LiteralPath $spaceFolder.source -Destination $spaceFolder.destination
    foreach ($spaceFile in $spaceFolder.files) {
        $spaceFilePath = Join-Path $spaceFolder.destination $spaceFile.path
        if ((Get-Item -LiteralPath $spaceFilePath).Length -ne $spaceFile.bytes -or (Get-FileHash -LiteralPath $spaceFilePath -Algorithm SHA256).Hash.ToLowerInvariant() -cne $spaceFile.sha256) { throw 'Moved file differs from pre-move bytes' }
    }
}
$spaceRecord = [pscustomobject]@{reviewer='root'; status='PRESERVED_REJECTED_CURATION'; campaignId='20261009T224729245Z-4221bb43'; reason='The v1 raw byte shards were named complete JSONL files; v2 corrected naming and replay maps without altering benchmark sources.'; folders=$spaceFolders; mismatches=0}
$spaceRecordPath = Join-Path $spaceEvidence 'root-rejected-ordinary-v1-preservation.json'
if (Test-Path -LiteralPath $spaceRecordPath) { throw 'Preservation record already exists' }
[IO.File]::WriteAllText($spaceRecordPath, ($spaceRecord | ConvertTo-Json -Depth 8) + "`n", [Text.UTF8Encoding]::new($false))
[pscustomobject]@{status='PRESERVED_REJECTED_CURATION'; folders=$spaceFolders.Count; mismatches=0; record=$spaceRecordPath} | ConvertTo-Json
