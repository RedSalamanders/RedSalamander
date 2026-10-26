$ErrorActionPreference = 'Stop'
$evidence = $PSScriptRoot
$plan = Get-Content -LiteralPath "$evidence/plan.json" -Raw | ConvertFrom-Json
& powercfg /getactivescheme | Set-Content -LiteralPath "$evidence/power-scheme.txt"
$receipts = @()
foreach ($entry in $plan) {
    $d = $entry.directory
    $before = @(Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^(MSBuild|cl|link|DxUiTests|DxUiBenchmark|RedSalamander|RedXeTests|RedXeBenchmark)\.exe$' } | Select-Object ProcessId,ParentProcessId,Name,CommandLine)
    $before | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath "$d/processes-before.json"
    $busy = @($before | Where-Object { $_.Name -ne 'MSBuild.exe' -or $_.CommandLine -notmatch '/nodemode:1' })
    if ($busy.Count) { throw "Another build/test/application is active; preserve processes-before.json at $d and resume later without discarding this attempt." }
    $exe = Join-Path $entry.repo '.build/x64/Release/RedSalamander.exe'
    if ((Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.exe_sha256) { throw 'Binary drift.' }
    $started = [DateTimeOffset]::UtcNow.ToString('o')
    & pwsh -NoProfile -File "$d/run.ps1" *> "$d/driver.log"
    $code = $LASTEXITCODE
    $after = @(Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^(MSBuild|cl|link|DxUiTests|DxUiBenchmark|RedSalamander|RedXeTests|RedXeBenchmark)\.exe$' } | Select-Object ProcessId,ParentProcessId,Name,CommandLine)
    $after | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath "$d/processes-after.json"
    $receipts += [ordered]@{ ordinal=$entry.ordinal; policy=$entry.policy; startedUtc=$started; completedUtc=[DateTimeOffset]::UtcNow.ToString('o'); exit=$code }
    $receipts | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath "$evidence/invocations.json"
    Write-Output "Invocation $($entry.ordinal) $($entry.policy): exit $code"
    if ($code -ne 0) { throw "Invocation failed; preserve $d." }
    if (@($after | Where-Object { $_.Name -ne 'MSBuild.exe' -or $_.CommandLine -notmatch '/nodemode:1' }).Count) { throw 'Build/test activity appeared during the invocation; keep evidence and review before resuming.' }
}
