$ErrorActionPreference = 'Stop'
$repo = 'C:/RsI26P0'
$evidence = 'C:/RedSalamander.Perf/evidence/i26-ui/common-p0-baseline-v1-20260927'
Set-Location -LiteralPath $repo
$expectedHead = 'e61e349c3e8d133e9d52a4c63ba4c49266bc4db5'
if ((& git rev-parse HEAD).Trim() -ne $expectedHead) { throw 'Qualification source changed.' }
# Frozen manifest includes the exact-bound candidate and identical deterministic fixture hooks.
$inputs = Get-Content -LiteralPath "$evidence/candidate-inputs.json" -Raw | ConvertFrom-Json
foreach ($inputFile in $inputs) {
    if ((Get-FileHash -LiteralPath (Join-Path $repo $inputFile.path) -Algorithm SHA256).Hash.ToLowerInvariant() -ne $inputFile.sha256) { throw "Frozen input changed before build: $($inputFile.path)" }
}

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


& ./vcpkg-install.ps1 -Platform x64 -VcpkgExe C:/RsI26Vcpkg/vcpkg.exe *> "$evidence/dependencies.log"
$dependencyExit = $LASTEXITCODE
Set-Content "$evidence/dependencies-exit.txt" $dependencyExit
Get-Content "$evidence/dependencies.log" -Tail 10
if ($dependencyExit -ne 0) { exit $dependencyExit }
& ./build.ps1 -Configuration Release -Platform x64 -MaxCpuCount 4 *> "$evidence/build.log"
$buildExit = $LASTEXITCODE
Set-Content "$evidence/build-exit.txt" $buildExit
Get-Content "$evidence/build.log" -Tail 35
if ($buildExit -ne 0) { exit $buildExit }; Copy-Item -LiteralPath "$repo/.build/x64/Release/build-receipt.json" -Destination "$evidence/build-receipt.json"




& 'C:/RedSalamander.Perf/evidence/i26-ui/common-p0-baseline-v1-20260927/run-focused.ps1'
exit $LASTEXITCODE






