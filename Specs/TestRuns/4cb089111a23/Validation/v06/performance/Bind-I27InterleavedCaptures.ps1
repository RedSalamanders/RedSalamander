param()
$ErrorActionPreference = 'Stop'
$taskRoot = Join-Path $PSScriptRoot 'paired-current-r21'
$taskBindingPath = Join-Path $taskRoot 'candidate-perf-build-binding.json'
$taskBinding = Get-Content -LiteralPath $taskBindingPath -Raw | ConvertFrom-Json
$taskCaptures = @(Get-Content -LiteralPath (Join-Path $taskRoot 'interleaved-capture-bindings.json') -Raw | ConvertFrom-Json)
if ($taskCaptures.Count -ne 10) { throw 'All ten interleaved process captures must be present.' }
$taskCandidates = [Collections.Generic.List[object]]::new()
$taskReferences = @{}
foreach ($taskCapture in $taskCaptures) {
    $taskVariant = $taskCapture.variant
    if ($taskVariant -notin @('baseline', 'candidate')) { throw 'Unknown interleaved variant.' }
    $taskRepo = if ($taskVariant -eq 'baseline') { 'C:/RsI26Adopt' } else { 'C:/RsI27Safety' }
    $taskReceipt = Get-Content -LiteralPath (Join-Path $taskRoot "$taskVariant-build-receipt.json") -Raw | ConvertFrom-Json
    $taskRun = Get-Content -LiteralPath (Join-Path $taskRoot "$taskVariant-$($taskCapture.capture)/artifacts/selftest/last_run/run-all-tests-results.json") -Raw | ConvertFrom-Json
    if ($taskCapture.run_id -ne $taskRun.run_id -or $taskCapture.receipt_id -ne $taskReceipt.receipt_id -or
        $taskCapture.source_snapshot_id -ne $taskReceipt.source_snapshot_id -or
        $taskCapture.modules_checked_before_and_after -ne @($taskCapture.runtime).Count) {
        throw 'Recorded interleaved source/runtime identity differs from its actual capture.'
    }
    $taskModuleSet = @($taskCapture.runtime | ForEach-Object {
        [pscustomobject]@{ relative_path = ([IO.Path]::GetRelativePath($taskRepo, $_.path) -replace '\\', '/'); sha256 = $_.sha256 }
    } | Sort-Object relative_path)
    $taskSerialized = $taskModuleSet | ConvertTo-Json -Depth 4 -Compress
    if ($taskReferences.ContainsKey($taskVariant)) {
        if ($taskReferences[$taskVariant] -cne $taskSerialized) { throw 'A variant runtime changed between captures.' }
    } else { $taskReferences[$taskVariant] = $taskSerialized }
    foreach ($taskModule in $taskCapture.runtime) {
        if ((Get-FileHash -LiteralPath $taskModule.path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $taskModule.sha256) {
            throw 'A captured runtime module no longer matches the retained identity.'
        }
    }
    if ($taskVariant -eq 'candidate') {
        foreach ($taskExpectedModule in $taskBinding.modules) {
            $taskMatches = @($taskModuleSet | Where-Object {
                $_.relative_path -ceq $taskExpectedModule.relative_path -and $_.sha256 -ceq $taskExpectedModule.sha256
            })
            if ($taskMatches.Count -ne 1) { throw 'Recorded candidate runtime lacks a final source-bound module or its hash differs.' }
        }
        $taskCandidates.Add([pscustomobject]@{
            capture = $taskCapture.capture; run_id = $taskCapture.run_id
            receipt_id = $taskCapture.receipt_id; source_snapshot_id = $taskCapture.source_snapshot_id
            runtime_binding_sha256 = (Get-FileHash -LiteralPath $taskBindingPath -Algorithm SHA256).Hash.ToLowerInvariant()
            modules_checked_before_and_after = @($taskBinding.modules).Count
            total_build_modules_checked_before_and_after = $taskCapture.modules_checked_before_and_after
            derivation = 'All source-bound runtime modules are members of the complete pre/post-checked build inventory; additional standalone test executables remain in the original interleaved records.'
        })
    }
}
foreach ($taskVariant in @('baseline', 'candidate')) {
    $taskIndices = @($taskCaptures | Where-Object variant -eq $taskVariant | ForEach-Object capture | Sort-Object)
    if (($taskIndices -join ',') -ne '1,2,3,4,5') { throw 'Capture index is missing or duplicated.' }
}
$taskExpectedOrder = 'baseline1,candidate1,candidate2,baseline2,baseline3,candidate3,candidate4,baseline4,baseline5,candidate5'
if ((@($taskCaptures | ForEach-Object { "$($_.variant)$($_.capture)" }) -join ',') -ne $taskExpectedOrder) {
    throw 'Recorded process order is not the stated alternating interleave.'
}
[IO.File]::WriteAllText((Join-Path $taskRoot 'candidate-capture-build-bindings.json'), ($taskCandidates.ToArray() | Sort-Object capture | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
Write-Output 'Derived five comparer bindings from the ten existing SHA-verified interleaved capture records; no run was repeated.'
