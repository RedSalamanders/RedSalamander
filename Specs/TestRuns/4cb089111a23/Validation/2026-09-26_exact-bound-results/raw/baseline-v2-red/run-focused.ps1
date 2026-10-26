$ErrorActionPreference = 'Stop'
$repo = 'C:/RsI26Final'
$evidence = 'C:/RedSalamander.Perf/evidence/i26-ui/exact-bound-results-baseline-v2-20260923'
Set-Location -LiteralPath $repo
$expectedHead = 'ccebd5af68eb45f2e9deed64a7063780e71eb1ce'
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
$cases = [ordered]@{ batch = 'FileOps_BatchRenameExactResultCarrier'; recycle = 'Phase10_DeferredConsentAndRecycleEscalation' }
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
    Get-Content -LiteralPath "$evidence/$($entry.Key).log" -Tail 28
}
$receipts | ConvertTo-Json -Depth 6 | Set-Content "$evidence/invocations.json"
# Preserve both independent baseline results. The missing-carrier assertions are expected to fail;
# a build or runner failure must be distinguished from that exact expected assertion by review.
if (@($receipts | Where-Object {$_.exit -ne 0}).Count -ne 0) { exit 1 }
exit 0


