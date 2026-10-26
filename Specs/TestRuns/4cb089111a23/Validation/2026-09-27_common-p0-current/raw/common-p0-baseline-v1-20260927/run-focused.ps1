$ErrorActionPreference = 'Stop'
$repo = 'C:/RsI26P0'
$evidence = 'C:/RedSalamander.Perf/evidence/i26-ui/common-p0-baseline-v1-20260927'
Set-Location -LiteralPath $repo
$expectedHead = 'e61e349c3e8d133e9d52a4c63ba4c49266bc4db5'
if ((& git rev-parse HEAD).Trim() -ne $expectedHead) { throw 'Qualification source changed.' }
# All explicit qualification source overlays are bound by candidate-inputs.json.

# Use the repository's CI integration, only in this child process. No user-wide import.
$env:VCPKG_ROOT = 'C:/RsI26Vcpkg'
$env:VcpkgRoot = $env:VCPKG_ROOT
$env:VCPKG_DEFAULT_TRIPLET = 'x64-windows'
$env:VCPKG_INSTALLED_DIR = "$repo/.build/vcpkg_installed/x64"
$installed = "$env:VCPKG_INSTALLED_DIR/x64-windows"
$env:ForceImportBeforeCppTargets = "$env:VCPKG_ROOT/scripts/buildsystems/msbuild/vcpkg.props"
$env:ForceImportAfterCppTargets = "$env:VCPKG_ROOT/scripts/buildsystems/msbuild/vcpkg.targets"
$env:VcpkgConfiguration = 'Release'
$env:VcpkgManifestInstall = 'false'
$env:VcpkgXUseBuiltInApplocalDeps = 'false'
$env:VCPkgLocalAppDataDisabled = 'true'
$env:CL = "/I`"$installed/include`" $env:CL"
$env:LINK = "/LIBPATH:`"$installed/lib`" /LIBPATH:`"$installed/lib/manual-link`" $env:LINK"
$env:Path = "$installed/bin;$env:Path"
$env:RSBuildEnableTests = 'true'


Import-Module ./Tools/Modules/Testing/TestRunPlan.psm1 -Force
$inputs = Get-Content -LiteralPath "$evidence/candidate-inputs.json" -Raw | ConvertFrom-Json
foreach ($inputFile in $inputs) {
    if ((Get-FileHash -LiteralPath (Join-Path $repo $inputFile.path) -Algorithm SHA256).Hash.ToLowerInvariant() -ne $inputFile.sha256) { throw "Frozen input changed: $($inputFile.path)" }
}
$cases = [ordered]@{ common_smoke = 'FileOps_CommonPresentationPerf' }
$receipts = @()
foreach ($entry in $cases.GetEnumerator()) {
    $plan = @(Get-RSTestRunPlan -Suite FileOps -CaseFilter $entry.Value -RepoRoot $repo -Platform x64 -Configuration Release -RedSalamanderExePath "$repo/.build/x64/Release/RedSalamander.exe")
    $plan | Select-Object Id,Kind,RequiresInteractiveDesktop | ConvertTo-Json -Depth 6 | Set-Content "$evidence/$($entry.Key)-plan.json"
    if (@($plan | Where-Object RequiresInteractiveDesktop).Count -ne 0) { throw 'Foreground validation is pending; refusing this native case.' }
    $started = [DateTimeOffset]::UtcNow.ToString('o')
    & ./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter $entry.Value -Configuration Release -Platform x64 -SkipBuild -TestRoot E:/RedSalamander.Perf *> "$evidence/$($entry.Key).log"
    $code = $LASTEXITCODE
    Set-Content -LiteralPath "$evidence/$($entry.Key)-exit.txt" -Value $code
    $receipts += [ordered]@{ case = $entry.Value; startedUtc = $started; completedUtc = [DateTimeOffset]::UtcNow.ToString('o'); exit = $code }
    $runMatch = [regex]::Match((Get-Content -LiteralPath "$evidence/$($entry.Key).log" -Raw), 'Run id:\s+([0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]+)')
    if ($runMatch.Success) {
        $sourceRun = Join-Path 'E:/RedSalamander.Perf/runs' $runMatch.Groups[1].Value
        $archiveRun = "$evidence/$($entry.Key)-runner-artifacts"
        if (Test-Path -LiteralPath $archiveRun) { throw 'Runner archive already exists.' }
        Copy-Item -LiteralPath $sourceRun -Destination $archiveRun -Recurse
        $copiedHashes = @(Get-ChildItem -LiteralPath $sourceRun -File -Recurse | ForEach-Object {
            $relative = [IO.Path]::GetRelativePath($sourceRun, $_.FullName)
            $sourceHash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            if ((Get-FileHash -LiteralPath (Join-Path $archiveRun $relative) -Algorithm SHA256).Hash.ToLowerInvariant() -ne $sourceHash) { throw "Runner archive differs: $relative" }
            [ordered]@{ path = $relative.Replace('\','/'); sha256 = $sourceHash }
        })
        $copiedHashes | ConvertTo-Json -Depth 6 | Set-Content "$evidence/$($entry.Key)-runner-copy-hashes.json"
        Write-Output "Archived $($entry.Key) runner: $($copiedHashes.Count) verified files."
    } else { Write-Output 'No application run id: preserve this as a runner/setup failure.' }
    Get-Content -LiteralPath "$evidence/$($entry.Key).log" -Tail 28
}
$receipts | ConvertTo-Json -Depth 6 | Set-Content "$evidence/invocations.json"
# Candidate must pass both cases; retain every failure without weakening its assertions.
if (@($receipts | Where-Object {$_.exit -ne 0}).Count -ne 0) { exit 1 }
exit 0





