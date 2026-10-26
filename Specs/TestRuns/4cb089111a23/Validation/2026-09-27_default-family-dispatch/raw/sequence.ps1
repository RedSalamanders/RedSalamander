$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
& "$root/family-dispatch-v6-release-20260927/build-and-test.ps1"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& "$root/family-dispatch-v6-debug-20260927/build-and-test.ps1"
exit $LASTEXITCODE
