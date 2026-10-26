$ErrorActionPreference = 'Stop'
$repo = 'C:/RsI26Final'
$evidence = 'C:/RedSalamander.Perf/evidence/i26-ui/terminal-group-debug-20260923'
Set-Location -LiteralPath $repo
$expectedHead = 'ccebd5af68eb45f2e9deed64a7063780e71eb1ce'
if ((& git rev-parse HEAD).Trim() -ne $expectedHead) { throw 'Qualification source changed.' }
$inputs = Get-Content -LiteralPath "$evidence/candidate-inputs.json" -Raw | ConvertFrom-Json
foreach ($inputFile in $inputs) {
 foreach ($sourceRepo in @($repo, 'Z:/src/RedSalamander')) {
  if ((Get-FileHash -LiteralPath (Join-Path $sourceRepo $inputFile.path) -Algorithm SHA256).Hash.ToLowerInvariant() -ne $inputFile.sha256) { throw "Candidate changed: $sourceRepo/$($inputFile.path)" }
 }
}

# Use the repository's CI integration, only in this child process. No user-wide import.
$env:VCPKG_ROOT = 'C:/RsI26Vcpkg'
$env:VcpkgRoot = $env:VCPKG_ROOT
$env:VCPKG_DEFAULT_TRIPLET = 'x64-windows'
$env:VCPKG_INSTALLED_DIR = "$repo/.build/vcpkg_installed/x64"
$installed = "$env:VCPKG_INSTALLED_DIR/x64-windows"
$env:ForceImportBeforeCppTargets = "$env:VCPKG_ROOT/scripts/buildsystems/msbuild/vcpkg.props"
$env:ForceImportAfterCppTargets = "$env:VCPKG_ROOT/scripts/buildsystems/msbuild/vcpkg.targets"
$env:VcpkgConfiguration = 'Debug'
$env:VcpkgManifestInstall = 'false'
$env:VcpkgXUseBuiltInApplocalDeps = 'false'
$env:VCPkgLocalAppDataDisabled = 'true'
$env:CL = "/I`"$installed/include`" $env:CL"
$env:LINK = "/LIBPATH:`"$installed/debug/lib`" /LIBPATH:`"$installed/debug/lib/manual-link`" $env:LINK"
$env:Path = "$installed/debug/bin;$env:Path"
$env:RSBuildEnableTests = 'true'


Import-Module ./Tools/Modules/Testing/TestRunPlan.psm1 -Force
$inputs = Get-Content -LiteralPath "$evidence/candidate-inputs.json" -Raw | ConvertFrom-Json
foreach ($inputFile in $inputs) { if ((Get-FileHash -LiteralPath (Join-Path $repo $inputFile.path) -Algorithm SHA256).Hash.ToLowerInvariant() -ne $inputFile.sha256) { throw "Candidate changed: $($inputFile.path)" } }
$plan = @(Get-RSTestRunPlan -Suite FileOps -CaseFilter FileOps_VisualGallery -RepoRoot $repo -Platform x64 -Configuration Debug -RedSalamanderExePath "$repo/.build/x64/Debug/RedSalamander.exe")
$plan | Select-Object Id,Kind,RequiresInteractiveDesktop | ConvertTo-Json -Depth 6 | Set-Content "$evidence/gallery-plan.json"
if (@($plan | Where-Object RequiresInteractiveDesktop).Count -ne 0) { throw 'Refusing foreground validation while desktop availability is pending.' }
& ./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_VisualGallery -Configuration Debug -Platform x64 -SkipBuild -TestRoot E:/RedSalamander.Perf *> "$evidence/gallery.log"
$testExit = $LASTEXITCODE
Set-Content "$evidence/gallery-exit.txt" $testExit
Get-Content "$evidence/gallery.log" -Tail 35
exit $testExit










