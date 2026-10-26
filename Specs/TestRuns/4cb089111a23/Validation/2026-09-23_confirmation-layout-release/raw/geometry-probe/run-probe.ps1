param([string]$SourceRoot='C:/RsI26Final',[string]$Label='old')
$ErrorActionPreference='Stop'
$probe = 'C:/RedSalamander.Perf/evidence/i26-ui/alert-geometry-probe-20260923'
Set-Location -LiteralPath $probe
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$install = & $vswhere -latest -prerelease -products '*' -property installationPath
Import-Module (Join-Path $install 'Common7/Tools/Microsoft.VisualStudio.DevShell.dll')
Enter-VsDevShell -VsInstallPath $install -SkipAutomaticLocation -DevCmdArguments '-arch=x64 -host_arch=x64' | Out-Null
$library = 'C:/RsI26Final/.build/dependencies/DxUi/acb51c8c8deec6239e5ff077f09d457e9e4ce9b8bd7b1d6164839a36d7d30a91/x64/Release/DxUi.lib'
$probeSource = if ($Label -eq 'old') { "$probe/probe-before-host-opt-in.cpp" } else { "$probe/probe.cpp" }
$compilerArgs = @('/nologo','/std:c++latest','/EHsc','/W4','/permissive-','/utf-8','/MD','/DUNICODE','/D_UNICODE','/DENABLE_TESTS','/I',"$SourceRoot/RedSalamander/Ui",'/I',"$SourceRoot/RedSalamander",'/I',"$SourceRoot/Common",'/I','C:/RsI26Final/RedSalamander','/I','C:/RsI26Final/.build/vcpkg_installed/x64/x64-windows/include','/I','C:/RsI26Final/.build/dependencies/DxUi/source/78b3de389a189c7f86f611787e0489fb6d474218/include',"/Fe:$probe/$Label.exe","/Fo:$probe/$Label.obj",$probeSource,$library,'dwrite.lib','d2d1.lib','user32.lib','gdi32.lib','ole32.lib','advapi32.lib')
& cl.exe @compilerArgs *> "$probe/$Label-build.log"
if ($LASTEXITCODE -ne 0) { Get-Content "$probe/$Label-build.log" -Tail 15; exit $LASTEXITCODE }
& "$probe/$Label.exe" *> "$probe/$Label-run.log"
$code=$LASTEXITCODE
Get-Content "$probe/$Label-run.log"
Set-Content -LiteralPath "$probe/$Label-exit.txt" $code
Get-FileHash -LiteralPath "$SourceRoot/RedSalamander/Ui/AlertOverlay.h" | ConvertTo-Json | Set-Content "$probe/$Label-source.json"
exit $code
