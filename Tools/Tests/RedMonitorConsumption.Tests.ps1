$repoRoot=Split-Path (Split-Path $PSScriptRoot)
Describe 'Independent RedMonitor consumption' {
 It 'verifies the single vendored header and license entirely offline' {
  & (Join-Path $PSScriptRoot 'Helpers/RedMonitorClient.ps1') -RepositoryRoot $repoRoot
 }
 It 'owns no viewer or viewer-test build target' {
  $solution=Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander.sln') -Raw
  $solution|Should Not Match 'RedSalamanderMonitor.*vcxproj|MonitorTest.*vcxproj'
  foreach($path in @('RedSalamanderMonitor/RedSalamanderMonitor.vcxproj','Tests/MonitorTest/MonitorTest.vcxproj')){
   Test-Path -LiteralPath (Join-Path $repoRoot $path)|Should Be $false
  }
  $shim=Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Restore-RedMonitor.ps1') -Raw
  $shim|Should Match 'Verify-RedMonitorClient'
  $shim|Should Not Match 'git clone|build.ps1|Copy-Item'
 }
 It 'normalizes its existing per-module providers and discovers only an installed helper' {
  $helpers=Get-Content -LiteralPath (Join-Path $repoRoot 'Common/Helpers.h') -Raw
  $helpers|Should Match 'RedMonitor::WriteNormalizedEvent'
  $helpers|Should Match 'g_RedSalamanderProvider'
  $startup=Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander/RedSalamander.cpp') -Raw
  $startup|Should Match 'RedMonitor::LaunchInstalledViewer'
  $startup|Should Match 'RedMonitor::IsOrdinaryStartup'
  $startup|Should Not Match 'QueueRedSalamanderMonitorLaunch'
 }
}

Describe 'Direct native RedMonitor dependency verification' -Tag RequiresBuildToolchain {
 BeforeAll {
  $script:clientGuardRoot=Split-Path (Split-Path $PSScriptRoot)
  Import-Module (Join-Path $clientGuardRoot 'Tools/Modules/Testing/TestSandbox.psm1')
  $vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
  $installation=(& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath).Trim()
  $relative=if([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64'){'MSBuild/Current/Bin/arm64/MSBuild.exe'}else{'MSBuild/Current/Bin/amd64/MSBuild.exe'}
  $script:clientGuardMSBuild=Join-Path $installation $relative
  [xml]$owner=Get-Content -LiteralPath (Join-Path $clientGuardRoot 'Directory.Build.targets') -Raw
  $script:clientGuardTarget=($owner.Project.Target|Where-Object Name -eq 'RSVerifyRedMonitorClient').OuterXml
 }
 It 'verifies a direct producer before build while excluding resources and external dependencies' -TestCases @(
  @{name='direct-plugin';native=$true;firstParty='true';tamper=$false;expected=0;verified=$true},
  @{name='tampered-direct-plugin';native=$true;firstParty='true';tamper=$true;expected=1;verified=$false},
  @{name='resource-only';native=$false;firstParty='true';tamper=$true;expected=0;verified=$false},
  @{name='external-dependency';native=$true;firstParty='false';tamper=$true;expected=0;verified=$false}
 ) {
  param($name,$native,$firstParty,$tamper,$expected,$verified)
  $scratch=New-RSTestSandboxScratchDirectory -RepoRoot $clientGuardRoot -Harness 'RedMonitorDirectBuild' -Case ($name+'-'+[guid]::NewGuid().ToString('N'))
  foreach($directory in 'Tools','Dependencies','ThirdParty/RedMonitor/include/RedMonitor'){[void][IO.Directory]::CreateDirectory((Join-Path $scratch $directory))}
  foreach($relativePath in 'Tools/Verify-RedMonitorClient.ps1','Dependencies/RedMonitor.lock.json','ThirdParty/RedMonitor/include/RedMonitor/Client.h','ThirdParty/RedMonitor/LICENSE.txt'){
   Copy-Item -LiteralPath (Join-Path $clientGuardRoot $relativePath) -Destination (Join-Path $scratch $relativePath)
  }
  if($tamper){[IO.File]::AppendAllText((Join-Path $scratch 'ThirdParty/RedMonitor/include/RedMonitor/Client.h'),'tampered')}
  [IO.File]::WriteAllText((Join-Path $scratch 'guard.targets'),('<Project>'+$clientGuardTarget+'</Project>'))
  $items=if($native){'<ItemGroup><ClCompile Include="producer.cpp" /></ItemGroup>'}else{'<ItemGroup><ResourceCompile Include="resources.rc" /></ItemGroup>'}
  $project='<Project><PropertyGroup><RSIsFirstPartyProject>'+ $firstParty +'</RSIsFirstPartyProject></PropertyGroup>'+$items+
   '<Import Project="guard.targets" /><Target Name="PrepareForBuild" /><Target Name="Build" DependsOnTargets="PrepareForBuild"><WriteLinesToFile File="admitted.txt" Lines="build-admitted" /></Target></Project>'
  $projectPath=Join-Path $scratch 'FileSystemDummy.vcxproj'
  [IO.File]::WriteAllText($projectPath,$project)
  $log=Join-Path $scratch 'msbuild.log'
  & $clientGuardMSBuild $projectPath /nologo /noautoresponse /nodeReuse:false /t:Build /v:minimal *> $log
  $code=$LASTEXITCODE
  ($code -eq 0)|Should Be ($expected -eq 0)
  (Test-Path -LiteralPath (Join-Path $scratch 'admitted.txt'))|Should Be ($expected -eq 0)
  if($verified){[IO.File]::ReadAllText($log)|Should Match 'Verified local RedMonitor client'}
  if($expected -ne 0){[IO.File]::ReadAllText($log)|Should Match 'Missing or modified RedMonitor client'}
 }
}
