$ErrorActionPreference = 'Stop'
$repo = 'Z:\src\RedSalamander'
$root = $PSScriptRoot
$temp = Join-Path $root 'temp'
$checkoutParent = Join-Path $root 'c'
$checkout = Join-Path $checkoutParent 'c'
$candidateIndex = Join-Path $root 'candidate.index'
$objectDir = Join-Path $root 'objects'
New-Item -ItemType Directory -Path $temp,$checkoutParent,$objectDir -Force | Out-Null
$env:TEMP = $temp
$env:TMP = $temp
$env:GIT_OBJECT_DIRECTORY = $objectDir
$env:GIT_ALTERNATE_OBJECT_DIRECTORIES = Join-Path $repo '.git\objects'
$env:GIT_OPTIONAL_LOCKS = '0'
$env:GIT_NO_LAZY_FETCH = '1'
Remove-Item Env:GIT_INDEX_FILE -ErrorAction SilentlyContinue
Remove-Item Env:GIT_WORK_TREE -ErrorAction SilentlyContinue

function Invoke-Git([string[]]$Arguments, [string]$IndexPath = $null, [string]$WorkTree = $null) {
    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = (Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $psi.CreateNoWindow = $true
    $psi.WorkingDirectory = $repo
    $psi.Environment['GIT_OBJECT_DIRECTORY'] = Join-Path $root 'objects'
    $psi.Environment['GIT_ALTERNATE_OBJECT_DIRECTORIES'] = Join-Path $repo '.git\objects'
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [System.Text.UTF8Encoding]::new($false)
    $psi.StandardErrorEncoding = [System.Text.UTF8Encoding]::new($false)
    foreach ($arg in @('--no-optional-locks','--no-lazy-fetch','-c','gc.auto=0','-c','maintenance.auto=false') + $Arguments) { [void]$psi.ArgumentList.Add($arg) }
    if ($IndexPath) { $psi.Environment['GIT_INDEX_FILE'] = $IndexPath } else { $psi.Environment.Remove('GIT_INDEX_FILE') }
    if ($WorkTree) { $psi.Environment['GIT_WORK_TREE'] = $WorkTree } else { $psi.Environment.Remove('GIT_WORK_TREE') }
    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $psi
    try {
    if (-not $process.Start()) { throw 'Could not start Git.' }
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    $process.WaitForExit()
    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()
    [pscustomobject]@{ ExitCode = $process.ExitCode; Stdout = $stdout; Stderr = $stderr }
    } finally { $process.Dispose() }
}
function Require-Git([string[]]$Arguments, [string]$IndexPath = $null, [string]$WorkTree = $null) {
    $result = Invoke-Git $Arguments $IndexPath $WorkTree
    if ($result.ExitCode -ne 0) { throw "Git failed ($($result.ExitCode)): git $($Arguments -join ' ')`n$($result.Stderr)$($result.Stdout)" }
    return $result.Stdout
}
function Hash-Bytes([byte[]]$Bytes) { [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($Bytes)) }
function Hash-Text([string]$Text) { Hash-Bytes ([System.Text.UTF8Encoding]::new($false).GetBytes($Text)) }
function Get-ObjectStoreFingerprint() {
    $base=Join-Path $repo '.git\objects'
    $items=[Collections.Generic.List[string]]::new()
    $stack=[Collections.Generic.Stack[string]]::new()
    $stack.Push($base)
    [uint64]$bytes=0
    while($stack.Count -gt 0) {
        foreach($path in [IO.Directory]::EnumerateFileSystemEntries($stack.Pop())) {
            $attr=[IO.File]::GetAttributes($path)
            if(($attr -band [IO.FileAttributes]::ReparsePoint) -ne 0){throw "Object store reparse point: $path"}
            if(($attr -band [IO.FileAttributes]::Directory) -ne 0){$stack.Push($path);continue}
            $info=[IO.FileInfo]::new($path)
            $relative=[IO.Path]::GetRelativePath($base,$path).Replace('\','/')
            $digest=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            $items.Add("$relative|$($info.Length)|$digest")
            $bytes += [uint64]$info.Length
        }
    }
    $ordered=@($items)
    [Array]::Sort($ordered,[StringComparer]::Ordinal)
    [pscustomobject]@{files=$ordered.Length;bytes=$bytes;sha256=(Hash-Text ($ordered -join "`n"))}
}

function Compare-RSManifestRelocationTextAttributes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Manifest,
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$CandidateIndex,
        [ValidateRange(1, 32)][int]$BatchPathCount = 32
    )

    $RepoRoot = [IO.Path]::GetFullPath($RepoRoot)
    $CandidateIndex = [IO.Path]::GetFullPath($CandidateIndex)
    $sourceCommit = [string]$Manifest.source_commit
    if ($sourceCommit -cnotmatch '^[0-9a-f]{40}$') {
        throw 'Relocation manifest source_commit must be a full lowercase Git commit SHA.'
    }
    $entries = @($Manifest.entries)
    if ($entries.Count -ne 3002) { throw "Expected 3002 exact relocation pairs; found $($entries.Count)." }

    $oldPaths = [Collections.Generic.List[string]]::new()
    $newPaths = [Collections.Generic.List[string]]::new()
    $oldSet = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $newSet = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in $entries) {
        $oldPath = [string]$entry.old_path
        $newPath = [string]$entry.new_path
        foreach ($path in @($oldPath, $newPath)) {
            if ([string]::IsNullOrWhiteSpace($path) -or $path.StartsWith('/') -or $path.Contains('\') -or
                $path.Contains(':') -or $path -match '[\x00-\x1f\x7f]' -or $path -match '(^|/)\.\.?(/|$)') {
                throw "Relocation manifest contains a non-repository-relative path: '$path'."
            }
        }
        if (-not $oldSet.Add($oldPath) -or -not $newSet.Add($newPath)) {
            throw "Relocation manifest contains a duplicate old or new path: '$oldPath' -> '$newPath'."
        }
        $oldPaths.Add($oldPath)
        $newPaths.Add($newPath)
    }

    # Require-Git is supplied by the stock-checkout proof script. Its UTF-8 text
    # result preserves embedded NULs from `check-attr -z`.
    $readAttributes = {
        param(
            [Parameter(Mandatory = $true)][string[]]$Paths,
            [AllowEmptyString()][string]$SourceCommit = '',
            [AllowNull()][string]$IndexPath = $null
        )
        $attributes = @('text', 'eol')
        $valuesByPath = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
        for ($offset = 0; $offset -lt $Paths.Count; $offset += $BatchPathCount) {
            $last = [Math]::Min($offset + $BatchPathCount - 1, $Paths.Count - 1)
            $batch = @($Paths[$offset..$last])
            $argumentPathCharacters = ($batch | ForEach-Object { $_.Length } | Measure-Object -Sum).Sum
            if ($argumentPathCharacters -gt 32768) { throw 'A check-attr path batch exceeds the 32 KiB argument budget.' }

            $arguments = [Collections.Generic.List[string]]::new()
            foreach ($argument in @('-C', $RepoRoot, '-c', 'core.quotePath=false', 'check-attr', '-z')) { $arguments.Add($argument) }
            if (-not [string]::IsNullOrWhiteSpace($SourceCommit)) {
                $arguments.Add("--source=$SourceCommit")
            } else {
                $arguments.Add('--cached')
            }
            foreach ($attribute in $attributes) { $arguments.Add($attribute) }
            $arguments.Add('--')
            foreach ($path in $batch) { $arguments.Add($path) }
            $raw = Require-Git $arguments.ToArray() $IndexPath $RepoRoot
            $tokens = @($raw.Split([char]0, [StringSplitOptions]::RemoveEmptyEntries))
            $expectedTokenCount = $batch.Count * $attributes.Count * 3
            if ($tokens.Count -ne $expectedTokenCount) {
                throw "git check-attr returned $($tokens.Count) NUL fields; expected $expectedTokenCount for this path batch."
            }
            $requested = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
            foreach ($path in $batch) { [void]$requested.Add($path) }
            for ($index = 0; $index -lt $tokens.Count; $index += 3) {
                $path = $tokens[$index]
                $attribute = $tokens[$index + 1]
                $value = $tokens[$index + 2]
                if (-not $requested.Contains($path) -or $attribute -cnotin $attributes) {
                    throw "git check-attr returned an unrequested record: path='$path' attribute='$attribute'."
                }
                if (-not $valuesByPath.ContainsKey($path)) {
                    $valuesByPath.Add($path, [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal))
                }
                $perPath = [Collections.Generic.Dictionary[string, string]]$valuesByPath[$path]
                if (-not $perPath.TryAdd($attribute, $value)) { throw "git check-attr duplicated '$attribute' for '$path'." }
            }
            foreach ($path in $batch) {
                if (-not $valuesByPath.ContainsKey($path)) { throw "git check-attr omitted path '$path'." }
                $perPath = [Collections.Generic.Dictionary[string, string]]$valuesByPath[$path]
                foreach ($attribute in $attributes) {
                    if (-not $perPath.ContainsKey($attribute)) { throw "git check-attr omitted '$attribute' for '$path'." }
                }
            }
        }
        return ,$valuesByPath
    }

    $oldAttributes = & $readAttributes -Paths $oldPaths.ToArray() -SourceCommit $sourceCommit
    $newAttributes = & $readAttributes -Paths $newPaths.ToArray() -IndexPath $CandidateIndex
    $mismatches = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $entries.Count; $index++) {
        $entry = $entries[$index]
        $oldPath = [string]$entry.old_path
        $newPath = [string]$entry.new_path
        foreach ($attribute in @('text', 'eol')) {
            $oldValue = ([Collections.Generic.Dictionary[string, string]]$oldAttributes[$oldPath])[$attribute]
            $newValue = ([Collections.Generic.Dictionary[string, string]]$newAttributes[$newPath])[$attribute]
            if ($oldValue -cne $newValue) {
                $mismatches.Add([pscustomobject]@{
                    OldPath = $oldPath; NewPath = $newPath; Attribute = $attribute
                    SourceValue = $oldValue; CandidateValue = $newValue
                })
            }
        }
    }
    return [pscustomobject]@{
        ComparedPaths = $entries.Count
        ComparedAttributes = $entries.Count * 2
        MismatchCount = $mismatches.Count
        Mismatches = @($mismatches)
    }
}

$runId = Split-Path -Leaf (Split-Path -Parent $root)
$indexSource = Join-Path $repo '.git\index'
$headBefore = (Require-Git @('-C',$repo,'rev-parse','HEAD')).Trim()
$statusBefore = Require-Git @('-C',$repo,'-c','core.quotePath=false','status','--porcelain=v1','-z','--untracked-files=all')
$statusBeforePath = Join-Path $root 'status-before.porcelain-z'
[System.IO.File]::WriteAllText($statusBeforePath, $statusBefore, [System.Text.UTF8Encoding]::new($false))
$indexBeforeHash = (Get-FileHash -LiteralPath $indexSource -Algorithm SHA256).Hash
$attrsBeforeHash = (Get-FileHash -LiteralPath (Join-Path $repo '.gitattributes') -Algorithm SHA256).Hash
$gitVersion = (Require-Git @('--version')).Trim()
$config = [ordered]@{}
foreach ($name in @('core.longpaths','core.symlinks','core.autocrlf','core.splitIndex','core.filemode')) {
    $value = Invoke-Git @('-C',$repo,'config','--show-origin','--get',$name)
    $config[$name] = if ($value.ExitCode -eq 0) { $value.Stdout.Trim() } else { '<unset>' }
}
$sharedBefore = @(Get-ChildItem -LiteralPath (Join-Path $repo '.git') -Filter 'sharedindex.*' -File -ErrorAction SilentlyContinue | ForEach-Object { '{0}:{1}' -f $_.Name,(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash })
$objectStoreBefore = Get-ObjectStoreFingerprint

$proof = [ordered]@{
    runId=$runId; repo=$repo; checkoutRoot=$checkout; temporaryIndex=$candidateIndex
    gitVersion=$gitVersion; config=$config; invocationOverrides=@('core.longpaths=false','core.symlinks=false','core.splitIndex=false')
    headBefore=$headBefore; objectStoreDirectory=(Join-Path $root 'objects'); objectStoreUsesRepoReadOnlyAlternate=(Join-Path $repo '.git\objects'); repoObjectStoreSha256Before=$objectStoreBefore.sha256; repoObjectStoreFileCountBefore=$objectStoreBefore.files; repoObjectStoreBytesBefore=$objectStoreBefore.bytes; repoObjectStoreSha256After=$null; repoObjectStoreFileCountAfter=$null; repoObjectStoreBytesAfter=$null; repoObjectStoreUnchanged=$null; isolatedObjectFileCount=$null; sourceIndexSha256Before=$indexBeforeHash; sourceStatusSha256Before=(Hash-Text $statusBefore)
    attributesSha256Before=$attrsBeforeHash; sourceIndexSha256After=$null; sourceStatusSha256After=$null; sourceStatusUnchanged=$null
    sharedIndexFilesBefore=$sharedBefore; sharedIndexFilesAfter=$null; candidateIndexSha256BeforeAdd=$null; candidateIndexSha256AfterAdd=$null
    entries=0; modes=@{}; unsupportedModes=@(); nonzeroStages=0; maxRelativePathChars=0; maxRelativePath=$null
    maxCheckoutAbsolutePathChars=0; maxCheckoutAbsolutePath=$null; checkoutFiles=0; checkoutDirectories=0
    reparseEntries=@(); hardlinkCountFailures=@(); checkoutRoundTripEntries=0; checkoutRoundTripMismatches=@(); alternateCheckoutStatusClean=$false
    immutableRelocationEntries=0; immutableRelocationTotalBytes=0; immutableRelocationFailures=@(); sourceHeadAfter=$null; durationSeconds=$null
    result='incomplete'; failure=$null
}
$started = [System.Diagnostics.Stopwatch]::StartNew()
try {
    if (Test-Path -LiteralPath $checkout) { throw "Checkout destination must be fresh: $checkout" }
    Copy-Item -LiteralPath $indexSource -Destination $candidateIndex
    $proof.candidateIndexSha256BeforeAdd = (Get-FileHash -LiteralPath $candidateIndex -Algorithm SHA256).Hash
    Write-Output "Captured index/status. Candidate temp index: $candidateIndex"
    [void](Require-Git @('-C',$repo,'-c','core.splitIndex=false','update-index','--no-split-index') $candidateIndex)
    [void](Require-Git @('-C',$repo,'-c','core.longpaths=false','-c','core.symlinks=false','-c','core.splitIndex=false','add','-A') $candidateIndex)
    $proof.candidateIndexSha256AfterAdd = (Get-FileHash -LiteralPath $candidateIndex -Algorithm SHA256).Hash
    Write-Output 'Temporary-index git add -A completed.'

    $stageText = Require-Git @('-C',$repo,'-c','core.quotePath=false','-c','core.splitIndex=false','ls-files','--stage','-z') $candidateIndex
    $records = @($stageText.Split([char]0,[System.StringSplitOptions]::RemoveEmptyEntries))
    $entries = [System.Collections.Generic.List[object]]::new()
    $pathSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $modeCounts = @{}
    $badModes = [System.Collections.Generic.List[object]]::new()
    $nonzeroStages = 0
    $maxRel = 0; $maxRelPath = ''
    $maxAbs = 0; $maxAbsPath = ''
    foreach ($record in $records) {
        $tab = $record.IndexOf("`t")
        if ($tab -lt 0) { throw 'Malformed git ls-files stage record.' }
        $meta = $record.Substring(0,$tab).Split(' ')
        if ($meta.Length -ne 3) { throw 'Malformed stage metadata.' }
        $mode = $meta[0]; $oid = $meta[1]; $stage = [int]$meta[2]; $path = $record.Substring($tab+1)
        if ($stage -ne 0) { $nonzeroStages++ }
        if ($modeCounts.ContainsKey($mode)) { $modeCounts[$mode]++ } else { $modeCounts[$mode]=1 }
        if ($mode -notin @('100644','100755')) { $badModes.Add([ordered]@{mode=$mode;path=$path;oid=$oid;stage=$stage}) }
        if (-not $pathSet.Add($path)) { throw "Case-insensitive duplicate path in index: $path" }
        $nativeRel = $path.Replace('/','\')
        if ($nativeRel.Length -gt $maxRel) { $maxRel=$nativeRel.Length; $maxRelPath=$path }
        $absolute = [System.IO.Path]::GetFullPath((Join-Path $checkout $nativeRel))
        if ($absolute.Length -gt $maxAbs) { $maxAbs=$absolute.Length; $maxAbsPath=$absolute }
        $entries.Add([ordered]@{mode=$mode;oid=$oid;stage=$stage;path=$path})
    }
    $proof.entries=$entries.Count; $proof.modes=$modeCounts; $proof.unsupportedModes=@($badModes); $proof.nonzeroStages=$nonzeroStages
    $proof.maxRelativePathChars=$maxRel; $proof.maxRelativePath=$maxRelPath; $proof.maxCheckoutAbsolutePathChars=$maxAbs; $proof.maxCheckoutAbsolutePath=$maxAbsPath
    $inventoryJson = ConvertTo-Json -InputObject @($entries) -Depth 4 -Compress
    $inventoryBytes = [System.Text.UTF8Encoding]::new($false).GetBytes($inventoryJson)
    $inventoryPath = Join-Path $root 'staged-modes-paths.json.gz'
    $file = [System.IO.File]::Create($inventoryPath)
    try { $gzip=[System.IO.Compression.GZipStream]::new($file,[System.IO.Compression.CompressionLevel]::Optimal); try{$gzip.Write($inventoryBytes,0,$inventoryBytes.Length)}finally{$gzip.Dispose()} } finally{$file.Dispose()}
    if ($maxAbs -ge 260) { throw "Stock checkout path budget exceeded: $maxAbs characters." }
    if ($nonzeroStages -ne 0) { throw "Candidate index contains $nonzeroStages non-stage-zero entries." }
    if ($badModes.Count -ne 0) { throw "Candidate index contains unsupported/link modes: $($badModes.Count). See proof." }

    New-Item -ItemType Directory -Path $checkout | Out-Null
    [void](Require-Git @('-C',$repo,'-c','core.longpaths=false','-c','core.symlinks=false','-c','core.splitIndex=false','checkout-index','--all','--force',"--prefix=$($checkout.TrimEnd('\'))\") $candidateIndex)
    Write-Output "checkout-index completed ($($entries.Count) tracked paths)."

    Add-Type -Path (Join-Path $root 'GitLinkProbe.cs')
    $actualPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $reparse = [System.Collections.Generic.List[string]]::new()
    $hardlinks = [System.Collections.Generic.List[string]]::new()
    $fileCount=0; $dirCount=0
    $stack=[System.Collections.Generic.Stack[string]]::new(); $stack.Push($checkout)
    while($stack.Count -gt 0) {
        $directory=$stack.Pop()
        foreach($entryPath in [System.IO.Directory]::EnumerateFileSystemEntries($directory)) {
            $attrs=[System.IO.File]::GetAttributes($entryPath)
            if (($attrs -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { $reparse.Add($entryPath); continue }
            if (($attrs -band [System.IO.FileAttributes]::Directory) -ne 0) { $dirCount++; $stack.Push($entryPath); continue }
            $fileCount++
            $relative=[System.IO.Path]::GetRelativePath($checkout,$entryPath).Replace('\','/')
            if (-not $actualPaths.Add($relative)) { throw "Duplicate checkout path: $relative" }
            $handle=[System.IO.File]::OpenHandle($entryPath,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
            try { $info=[RsGitLinkProbe+BY_HANDLE_FILE_INFORMATION]::new(); if (-not [RsGitLinkProbe]::GetFileInformationByHandle($handle,[ref]$info)) { throw "GetFileInformationByHandle failed: $entryPath" }; if ($info.NumberOfLinks -ne 1) { $hardlinks.Add("$($info.NumberOfLinks):$entryPath") } }
            finally { $handle.Dispose() }
        }
    }
    $missing=[System.Collections.Generic.List[string]]::new()
    foreach($item in $entries) { if (-not $actualPaths.Contains([string]$item.path)) { $missing.Add([string]$item.path) } }
    $extra=[System.Collections.Generic.List[string]]::new()
    foreach($actual in $actualPaths) { if (-not $pathSet.Contains($actual)) { $extra.Add($actual) } }
    $proof.checkoutFiles=$fileCount; $proof.checkoutDirectories=$dirCount; $proof.reparseEntries=@($reparse); $proof.hardlinkCountFailures=@($hardlinks)
    if ($missing.Count -or $extra.Count) { throw "Checkout path-set mismatch; missing=$($missing.Count), extra=$($extra.Count)." }
    if ($reparse.Count -or $hardlinks.Count) { throw "Checkout contains reparse points or hardlinked files." }

    # Re-stage the checkout through Git clean filters into only the private index, then
    # compare every candidate-index mode/OID/stage/path with the original staged inventory.
    $restage = Invoke-Git @('-C',$repo,'-c','core.longpaths=false','-c','core.symlinks=false','-c','core.splitIndex=false','add','--renormalize','-A') $candidateIndex $checkout
    if ($restage.ExitCode -ne 0) { throw "Alternate-worktree git add failed: $($restage.Stderr)$($restage.Stdout)" }
    $afterStageText = Require-Git @('-C',$repo,'-c','core.quotePath=false','-c','core.splitIndex=false','ls-files','--stage','-z') $candidateIndex
    $afterRecords = @($afterStageText.Split([char]0,[System.StringSplitOptions]::RemoveEmptyEntries))
    $changedRecords = [System.Collections.Generic.List[object]]::new()
    $beforeMap = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::Ordinal)
    foreach ($entry in $entries) { $beforeMap[[string]$entry.path] = "$($entry.mode)|$($entry.oid)|$($entry.stage)" }
    $afterMap = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::Ordinal)
    foreach ($record in $afterRecords) {
        $tab = $record.IndexOf("`t"); if ($tab -lt 0) { throw 'Malformed re-stage record.' }
        $meta = $record.Substring(0,$tab).Split(' '); if ($meta.Length -ne 3) { throw 'Malformed re-stage metadata.' }
        $path = $record.Substring($tab+1); $value = "$($meta[0])|$($meta[1])|$($meta[2])"
        $afterMap[$path] = $value
        if (-not $beforeMap.ContainsKey($path) -or $beforeMap[$path] -cne $value) { $changedRecords.Add([ordered]@{path=$path;after=$value}) }
    }
    foreach ($entry in $entries) { if (-not $afterMap.ContainsKey([string]$entry.path)) { $changedRecords.Add([ordered]@{path=$entry.path;before="$($entry.mode)|$($entry.oid)|$($entry.stage)";after=$null}) } }
    $proof.checkoutRoundTripEntries = $afterRecords.Count
    $proof.checkoutRoundTripMismatches = @($changedRecords)
    $roundTripStatus = Invoke-Git @('-C',$repo,'-c','core.quotePath=false','status','--porcelain=v1','-z','--untracked-files=all') $candidateIndex $checkout
    $proof.alternateCheckoutStatusClean = $roundTripStatus.ExitCode -eq 0 -and $roundTripStatus.Stdout.Length -eq 0
    if ($afterRecords.Count -ne $entries.Count -or $changedRecords.Count -ne 0) { throw 'Checkout clean-filter re-stage mode/OID/stage/path parity failed.' }

    $manifest=Get-Content -LiteralPath (Join-Path $repo 'Specs\Testing\ProtectedPathRelocations.json') -Raw | ConvertFrom-Json
    if(@($manifest.entries).Count -ne 3002){throw 'The exact 3002-path migration manifest is required.'}
    Import-Module (Join-Path $repo 'Tools\Modules\Testing\ValidationFingerprint.psm1') -Force
    $blobDigests=Get-RSValidationGitBlobContentSha256Map -RepoRoot $repo -BlobId @($manifest.entries | ForEach-Object {[string]$_.blob})
    $physicalBaseline=Get-Content -LiteralPath (Join-Path $repo 'Specs\TestRuns\Local-x64\MergeReview\2026-10-31_Remediation\stock-checkout-source-byte-comparison.json') -Raw | ConvertFrom-Json
    $expectedPhysical=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
    foreach($item in $physicalBaseline.firstMismatches){$expectedPhysical.Add([string]$item.path,$item)}
    if($expectedPhysical.Count -ne 4){throw 'The previously witnessed four physical autocrlf differences are required.'}
    $actualPhysical=[Collections.Generic.List[object]]::new()
    [uint64]$sourceRelocationBytes=0
    [uint64]$checkoutRelocationBytes=0
    foreach($relocation in $manifest.entries) {
        $path=[string]$relocation.new_path
        $expectedIndex="$($relocation.mode)|$($relocation.blob)|0"
        if(-not $beforeMap.ContainsKey($path) -or $beforeMap[$path] -cne $expectedIndex){$proof.immutableRelocationFailures += "index:$path";continue}
        if($blobDigests[[string]$relocation.blob] -cne [string]$relocation.blob_content_sha256){$proof.immutableRelocationFailures += "canonical-blob:$path";continue}
        $sourcePath=Join-Path $repo ($path.Replace('/','\'))
        $checkoutPath=Join-Path $checkout ($path.Replace('/','\'))
        $sourceHash=(Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $checkoutHash=(Get-FileHash -LiteralPath $checkoutPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $sourceBytes=(Get-Item -LiteralPath $sourcePath).Length
        $checkoutBytes=(Get-Item -LiteralPath $checkoutPath).Length
        $sourceRelocationBytes += [uint64]$sourceBytes
        $checkoutRelocationBytes += [uint64]$checkoutBytes
        if($sourceHash -cne $checkoutHash) {
            if(-not $expectedPhysical.ContainsKey($path)){throw "Unreviewed physical checkout difference: $path"}
            $expected=$expectedPhysical[$path]
            if($sourceHash -cne [string]$expected.sourceSha256 -or $checkoutHash -cne [string]$expected.checkoutSha256 -or
                $sourceBytes -ne $expected.sourceBytes -or $checkoutBytes -ne $expected.checkoutBytes){throw "Physical autocrlf witness changed: $path"}
            $actualPhysical.Add([pscustomobject]@{path=$path;sourceSha256=$sourceHash;checkoutSha256=$checkoutHash;sourceBytes=$sourceBytes;checkoutBytes=$checkoutBytes})
        }
        $proof.immutableRelocationEntries++
    }
    if($proof.immutableRelocationFailures.Count -ne 0 -or $proof.immutableRelocationEntries -ne 3002){throw 'Immutable relocation canonical blob/mode checks failed.'}
    if($actualPhysical.Count -ne 4 -or ($checkoutRelocationBytes-$sourceRelocationBytes) -ne 30){throw 'Physical checkout changes differ from the four pre-existing autocrlf conversions.'}
    $proof.immutableRelocationTotalBytes=$sourceRelocationBytes
    $proof['checkoutRelocationBytes']=$checkoutRelocationBytes
    $proof['physicalAutocrlfDifferences']=@($actualPhysical)
    $proof['rawIdenticalRelocations']=3002-$actualPhysical.Count
    $attrParity=Compare-RSManifestRelocationTextAttributes -Manifest $manifest -RepoRoot $repo -CandidateIndex $candidateIndex
    $proof['relocationAttributeParity']=$attrParity
    if($attrParity.MismatchCount -ne 0){throw 'Old/new relocation text/eol policies differ.'}
    $proof['objectStoreFingerprintMethod']='Sorted relative path, byte length and SHA-256 content; no timestamps.'
    $proof['alternateStatusAcceptance']='Diagnostic only. Acceptance uses exact clean-filter mode/OID/stage/path parity.'

    $proof.result='pass'
} catch {
    $proof.result='fail'; $proof.failure=$_.Exception.Message
} finally {
    $started.Stop(); $proof.durationSeconds=[Math]::Round($started.Elapsed.TotalSeconds,3)
    try { $proof.sourceIndexSha256After=(Get-FileHash -LiteralPath $indexSource -Algorithm SHA256).Hash } catch { $proof.sourceIndexSha256After='unreadable' }
    try {
        $statusAfter=Require-Git @('-C',$repo,'-c','core.quotePath=false','status','--porcelain=v1','-z','--untracked-files=all')
        [System.IO.File]::WriteAllText((Join-Path $root 'status-after.porcelain-z'),$statusAfter,[System.Text.UTF8Encoding]::new($false))
        $proof.sourceStatusSha256After=Hash-Text $statusAfter
        $proof.sourceStatusUnchanged=($proof.sourceStatusSha256Before -ceq $proof.sourceStatusSha256After)
        $proof.sourceHeadAfter=(Require-Git @('-C',$repo,'rev-parse','HEAD')).Trim()
        $proof.attributesSha256After=(Get-FileHash -LiteralPath (Join-Path $repo '.gitattributes') -Algorithm SHA256).Hash
        $proof.sharedIndexFilesAfter=@(Get-ChildItem -LiteralPath (Join-Path $repo '.git') -Filter 'sharedindex.*' -File -ErrorAction SilentlyContinue | ForEach-Object { '{0}:{1}' -f $_.Name,(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash })
        $objectStoreAfter=Get-ObjectStoreFingerprint; $proof.repoObjectStoreSha256After=$objectStoreAfter.sha256; $proof.repoObjectStoreFileCountAfter=$objectStoreAfter.files; $proof.repoObjectStoreBytesAfter=$objectStoreAfter.bytes; $proof.repoObjectStoreUnchanged=($objectStoreBefore.sha256 -ceq $objectStoreAfter.sha256); $proof.isolatedObjectFileCount=@(Get-ChildItem -LiteralPath (Join-Path $root 'objects') -Recurse -File -Force).Count
    } catch { $proof.finalSnapshotFailure=$_.Exception.Message }
    $proof.sourceIndexUnchanged=($proof.sourceIndexSha256Before -ceq $proof.sourceIndexSha256After)
    $proof.sourceHeadUnchanged=($proof.headBefore -ceq $proof.sourceHeadAfter)
    $proof.attributesUnchanged=($attrsBeforeHash -ceq $proof.attributesSha256After)
    $proof.sharedIndexFilesUnchanged=((@($proof.sharedIndexFilesBefore) -join "`n") -ceq (@($proof.sharedIndexFilesAfter) -join "`n"))
    if ($proof.result -eq 'pass' -and (-not $proof.sourceIndexUnchanged -or -not $proof.sourceStatusUnchanged -or -not $proof.sourceHeadUnchanged -or -not $proof.attributesUnchanged -or -not $proof.sharedIndexFilesUnchanged -or -not $proof.repoObjectStoreUnchanged)) { $proof.result='fail'; $proof.failure='Source index/status/HEAD/attributes, git shared-index files, or repository object store changed during isolated proof.' }
    $json=ConvertTo-Json -InputObject $proof -Depth 12
    [System.IO.File]::WriteAllText((Join-Path $root 'stock-checkout-result.json'),$json,[System.Text.UTF8Encoding]::new($false))
    Write-Output $json
}
if ($proof.result -ne 'pass') { throw "Stock checkout proof failed: $($proof.failure)" }

