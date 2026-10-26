$ErrorActionPreference = 'Stop'
$repo = 'C:/RsI26Final'
$evidence = 'C:/RedSalamander.Perf/evidence/i26-ui/terminal-facts-qualified-v2-20260923'
Set-Location -LiteralPath $repo
$expectedHead = 'ccebd5af68eb45f2e9deed64a7063780e71eb1ce'
if ((& git rev-parse HEAD).Trim() -ne $expectedHead) { throw 'Qualification source changed.' }
# 25 explicit candidate source files are bound by candidate-inputs.json; baseline captures are retained.

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


& ./build.ps1 -Configuration Release -Platform x64 -MaxCpuCount 4 *> "$evidence/build.log"
$buildExit = $LASTEXITCODE
Set-Content "$evidence/build-exit.txt" $buildExit
Get-Content "$evidence/build.log" -Tail 35
if ($buildExit -ne 0) { exit $buildExit }; Copy-Item -LiteralPath "$repo/.build/x64/Release/build-receipt.json" -Destination "$evidence/build-receipt.json"




& 'C:/RedSalamander.Perf/evidence/i26-ui/terminal-facts-qualified-v2-20260923/capture-candidate.ps1'
exit $LASTEXITCODE






