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


& ./build.ps1 -Configuration Debug -Platform x64 -MaxCpuCount 4 *> "$evidence/build.log"
$buildExit = $LASTEXITCODE
Set-Content "$evidence/build-exit.txt" $buildExit
Get-Content "$evidence/build.log" -Tail 35
if ($buildExit -ne 0) { exit $buildExit }; Copy-Item -LiteralPath "$repo/.build/x64/Debug/build-receipt.json" -Destination "$evidence/build-receipt.json"




& 'C:/RedSalamander.Perf/evidence/i26-ui/terminal-group-debug-20260923/capture-candidate.ps1'
exit $LASTEXITCODE






