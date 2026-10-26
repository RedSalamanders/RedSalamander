<#
.SYNOPSIS
Exercises offline RedMonitor client verification with isolated immutable-file fixtures.
#>
[CmdletBinding()]
param([string]$RepositoryRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(-not $RepositoryRoot){
 $RepositoryRoot=$PSScriptRoot
 while(-not(Test-Path -LiteralPath (Join-Path $RepositoryRoot 'Dependencies/RedMonitor.lock.json'))){
  $parent=Split-Path $RepositoryRoot
  if(-not $parent -or $parent -eq $RepositoryRoot){throw 'Cannot locate consumer root.'}
  $RepositoryRoot=$parent
 }
}
$verify=Join-Path $RepositoryRoot 'Verify-RedMonitorClient.ps1'
if(-not(Test-Path -LiteralPath $verify)){$verify=Join-Path $RepositoryRoot 'Tools/Verify-RedMonitorClient.ps1'}
Import-Module (Join-Path $RepositoryRoot 'Tools/Modules/Testing/TestSandbox.psm1')
$fixture=New-RSTestSandboxScratchDirectory -RepoRoot $RepositoryRoot -Harness 'RedMonitorClient' -Case ('Offline-'+[guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory((Join-Path $fixture 'Dependencies'))|Out-Null
[IO.Directory]::CreateDirectory((Join-Path $fixture 'ThirdParty'))|Out-Null
Copy-Item -LiteralPath (Join-Path $RepositoryRoot 'ThirdParty/RedMonitor') -Destination (Join-Path $fixture 'ThirdParty') -Recurse
$lockPath=Join-Path $fixture 'Dependencies/RedMonitor.lock.json'
$baseline=[IO.File]::ReadAllText((Join-Path $RepositoryRoot 'Dependencies/RedMonitor.lock.json'))
[IO.File]::WriteAllText($lockPath,$baseline)
$results=[Collections.Generic.List[object]]::new()
function Invoke-Control([bool]$Pass,[string]$Name){
 $start=[Diagnostics.ProcessStartInfo]::new((Get-Process -Id $PID).Path)
 foreach($arg in @('-NoProfile','-File',$verify,'-RepositoryRoot',$fixture)){$start.ArgumentList.Add($arg)}
 $start.Environment['PATH']=Join-Path $env:SystemRoot 'System32'
 $start.UseShellExecute=$false;$start.CreateNoWindow=$true
 $start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
 $process=[Diagnostics.Process]::new();$process.StartInfo=$start
 try{
  if(-not $process.Start()){throw 'Cannot start verification control.'}
  $stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
  if(-not $process.WaitForExit(30000)){$process.Kill($true);$process.WaitForExit();throw 'Owned verification control timed out.'}
  [IO.File]::WriteAllText((Join-Path $fixture ($Name+'.log')),$stdout.GetAwaiter().GetResult()+$stderr.GetAwaiter().GetResult())
  $results.Add([pscustomobject]@{name=$Name;expectedSuccess=$Pass;exitCode=$process.ExitCode})
  if(($process.ExitCode -eq 0) -ne $Pass){throw "Client control failed: $Name"}
  if(Test-Path -LiteralPath (Join-Path $fixture '.build')){throw 'Verification created build/cache output.'}
  Write-Host "PASS RedMonitor client $Name"
 }finally{$process.Dispose()}
}
Invoke-Control $true 'offline-exact-bytes'
foreach($relative in @('include/RedMonitor/Client.h','LICENSE.txt')){
 $path=Join-Path $fixture ('ThirdParty/RedMonitor/'+$relative)
 $bytes=[IO.File]::ReadAllBytes($path)
 [IO.File]::AppendAllText($path,'corrupt')
 Invoke-Control $false ('tampered-'+[IO.Path]::GetFileName($path))
 [IO.File]::WriteAllBytes($path,$bytes)
 Move-Item -LiteralPath $path -Destination ($path+'.saved')
 Invoke-Control $false ('missing-'+[IO.Path]::GetFileName($path))
 Move-Item -LiteralPath ($path+'.saved') -Destination $path
}
foreach($field in @('repository','commit','apiRevision')){
 $pin=$baseline|ConvertFrom-Json
 $pin.$field=if($field -eq 'apiRevision'){99}else{'foreign'}
 $pin|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $lockPath
 Invoke-Control $false ('invalid-'+$field)
}
foreach($kind in @('duplicate','traversal','hash')){
 $pin=$baseline|ConvertFrom-Json
 switch($kind){
  duplicate {$pin.files[1]=$pin.files[0]}
  traversal {$pin.files[0].path='../Client.h'}
  hash {$pin.files[0].sha256='invalid'}
 }
 $pin|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $lockPath
 Invoke-Control $false ('invalid-'+$kind)
}
[IO.File]::WriteAllText($lockPath,$baseline)
$vendor=Join-Path $fixture 'ThirdParty/RedMonitor'
$savedVendor=$vendor+'.saved'
Move-Item -LiteralPath $vendor -Destination $savedVendor
try {
 New-RSTestSandboxLink -ItemType Junction -Path $vendor -Target $savedVendor|Out-Null
 Get-Item -LiteralPath $vendor -Force|Select-Object FullName,LinkType,Target|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $fixture 'reparse-control.json')
 Invoke-Control $false 'reparse-vendor'
} finally {
 if(Test-Path -LiteralPath $vendor){
  $link=Get-Item -LiteralPath $vendor -Force
  if(($link.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 -or $link.LinkType -ne 'Junction' -or
     [IO.Path]::GetFullPath([string]$link.Target) -ine [IO.Path]::GetFullPath($savedVendor)){
   throw 'Refusing to unlink a changed RedMonitor fixture path.'
  }
  # Remove only this known contained junction, without traversing its preserved target.
  [IO.Directory]::Delete($vendor,$false)
 }
 if(Test-Path -LiteralPath $savedVendor){Move-Item -LiteralPath $savedVendor -Destination $vendor}
}
# Keep immutable files, logs and link metadata as ordinary archivable evidence.
$results|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $fixture 'receipt.json')
Write-Host "PASS $($results.Count) offline client controls; retained $fixture"
