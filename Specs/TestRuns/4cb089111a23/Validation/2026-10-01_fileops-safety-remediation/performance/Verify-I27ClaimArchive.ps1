param([Parameter(Mandatory=$true)][string]$ArchiveRoot)
$ErrorActionPreference='Stop'
$taskRoot=[IO.Path]::GetFullPath($ArchiveRoot).TrimEnd('\','/')
$taskPrefix=$taskRoot+[IO.Path]::DirectorySeparatorChar
$taskManifestPath=Join-Path $taskRoot 'performance/selection-manifest.json'
$taskSelections=@(Get-Content -LiteralPath $taskManifestPath -Raw|ConvertFrom-Json)
if($taskSelections.Count -ne 23){throw 'All original, interleaved and unchanged-control captures must be retained.'}
$taskRecords=[Collections.Generic.List[object]]::new()
$taskReferenceDigests=$null
foreach($taskSelection in $taskSelections){
    $taskPath=[IO.Path]::GetFullPath((Join-Path $taskRoot $taskSelection.claim_path))
    if(-not $taskPath.StartsWith($taskPrefix,[StringComparison]::OrdinalIgnoreCase) -or
       (Get-Item -LiteralPath $taskPath).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Claim storage escapes archive root.'}
    if((Get-FileHash -LiteralPath $taskPath -Algorithm SHA256).Hash.ToLowerInvariant() -ne $taskSelection.stored_sha256 -or
       (Get-Item -LiteralPath $taskPath).Length -ne $taskSelection.stored_bytes){throw 'Compressed claim identity differs.'}
    if($taskSelection.claim_bytes -gt 2MB -or $taskSelection.encoding -notlike 'gzip *'){throw 'Unexpected claim encoding or decoded size.'}
    $taskInput=[IO.File]::OpenRead($taskPath)
    $taskOutput=[IO.MemoryStream]::new()
    try{
        $taskZip=[IO.Compression.GZipStream]::new($taskInput,[IO.Compression.CompressionMode]::Decompress,$true)
        try{
            $taskBuffer=[byte[]]::new(8192)
            while(($taskRead=$taskZip.Read($taskBuffer,0,$taskBuffer.Length)) -gt 0){
                if($taskOutput.Length+$taskRead -gt $taskSelection.claim_bytes){throw 'Decoded claims exceed declared bounded length.'}
                $taskOutput.Write($taskBuffer,0,$taskRead)
            }
        }finally{$taskZip.Dispose()}
        $taskBytes=$taskOutput.ToArray()
    }finally{$taskOutput.Dispose();$taskInput.Dispose()}
    if($taskBytes.Length -ne $taskSelection.claim_bytes -or
       [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($taskBytes)).ToLowerInvariant() -ne $taskSelection.claim_sha256){throw 'Decoded claim identity differs.'}
    $taskText=[Text.UTF8Encoding]::new($false,$true).GetString($taskBytes)
    if($taskText.Contains("`r") -or -not $taskText.EndsWith("`n")){throw 'Claim normalization differs from declared LF contract.'}
    $taskLines=@($taskText.Split([char]10)|Where-Object {-not [string]::IsNullOrWhiteSpace($_)})
    if($taskLines.Count -ne 275 -or $taskSelection.claim_rows -ne 275 -or @($taskSelection.zero_based_raw_line_indices).Count -ne 275){throw 'Claim row count differs.'}
    $taskPrevious=-1
    foreach($taskIndex in $taskSelection.zero_based_raw_line_indices){
        if($taskIndex -le $taskPrevious -or $taskIndex -ge $taskSelection.raw_line_count){throw 'Source claim ordering is not strictly increasing and bounded.'}
        $taskPrevious=$taskIndex
    }
    $taskRows=@($taskLines|ForEach-Object {$_|ConvertFrom-Json})
    if(@($taskRows|Where-Object { $_.metric -notlike 'FileOps.IntegratedPresentation.*' -or $_.machineHash -cne '4cb089111a23' -or $_.build -cne 'Release'}).Count -ne 0){throw 'Claim row scenario, machine or build differs.'}
    $taskDigests=@($taskLines|Where-Object {($_|ConvertFrom-Json).metric -like '*WorkloadFNV64'}|ForEach-Object {
        $taskMatch=[regex]::Match($_,'"durationUs"\s*:\s*([0-9]+)')
        if(-not $taskMatch.Success){throw 'Opaque workload digest token missing.'}
        $taskMatch.Groups[1].Value
    })
    if($taskDigests.Count -ne 2){throw 'Opaque digest controls are incomplete.'}
    if($null -eq $taskReferenceDigests){$taskReferenceDigests=$taskDigests -join '|'}
    elseif(($taskDigests -join '|') -cne $taskReferenceDigests){throw 'Exact workload digest changed across archived captures.'}
    $taskRecords.Add([pscustomobject]@{dataset=$taskSelection.dataset;variant=$taskSelection.variant;capture=$taskSelection.capture;
        run_id=$taskSelection.run_id;claim_path=$taskSelection.claim_path;rows_parsed=$taskRows.Count;
        decoded_sha256=$taskSelection.claim_sha256;stored_sha256=$taskSelection.stored_sha256})
}
$taskHistoryCount=0
$taskHistoryPath=Join-Path $taskRoot 'lossless-history-manifest.json'
if(Test-Path -LiteralPath $taskHistoryPath){
    foreach($taskHistory in @(Get-Content -LiteralPath $taskHistoryPath -Raw|ConvertFrom-Json)){
        $taskStored=[IO.Path]::GetFullPath((Join-Path $taskRoot $taskHistory.stored_path))
        if(-not $taskStored.StartsWith($taskPrefix,[StringComparison]::OrdinalIgnoreCase) -or $taskHistory.original_bytes -gt 2MB){throw 'Historical storage path or decoded size is invalid.'}
        if((Get-FileHash -LiteralPath $taskStored -Algorithm SHA256).Hash.ToLowerInvariant() -ne $taskHistory.stored_sha256){throw 'Historical compressed identity differs.'}
        $taskInput=[IO.File]::OpenRead($taskStored)
        $taskOutput=[IO.MemoryStream]::new()
        try{
            $taskZip=[IO.Compression.GZipStream]::new($taskInput,[IO.Compression.CompressionMode]::Decompress,$true)
            try{
                $taskBuffer=[byte[]]::new(8192)
                while(($taskRead=$taskZip.Read($taskBuffer,0,$taskBuffer.Length)) -gt 0){
                    if($taskOutput.Length+$taskRead -gt $taskHistory.original_bytes){throw 'Decoded historical bytes exceed declared bound.'}
                    $taskOutput.Write($taskBuffer,0,$taskRead)
                }
            }finally{$taskZip.Dispose()}
            if($taskOutput.Length -ne $taskHistory.original_bytes -or [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($taskOutput.ToArray())).ToLowerInvariant() -ne $taskHistory.original_sha256){throw 'Historical decoded identity differs.'}
        }finally{$taskOutput.Dispose();$taskInput.Dispose()}
        ++$taskHistoryCount
    }
}
$taskResult=[ordered]@{schema='red-salamander.i27-lossless-claim-verification.v1';timestamp_utc=[DateTime]::UtcNow.ToString('o');historical_traces_verified=$taskHistoryCount;
    captures_verified=$taskRecords.Count;rows_parsed=($taskRecords|Measure-Object rows_parsed -Sum).Sum;
    selection_manifest_sha256=(Get-FileHash -LiteralPath $taskManifestPath -Algorithm SHA256).Hash.ToLowerInvariant();records=$taskRecords.ToArray()}
[IO.File]::WriteAllText((Join-Path $taskRoot 'performance/claim-verification.json'),($taskResult|ConvertTo-Json -Depth 8 -Compress),[Text.UTF8Encoding]::new($false))
Write-Output "Verified $($taskRecords.Count) lossless capture files, $($taskResult.rows_parsed) ordered parsed claim rows, and $taskHistoryCount complete historical traces."
