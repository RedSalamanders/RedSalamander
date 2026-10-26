param([string]$Phase)
$ErrorActionPreference='Stop'
$quiet=[Diagnostics.Stopwatch]::StartNew()
$previous=''
while($true){
 $nodes=@()
 $blocked=$false
 $tools=@(Get-Process -Name MSBuild,cl,link,lib,rc,mt,cmake,ninja,clang,clang-cl,lld-link,csc,vbc,dotnet -ErrorAction SilentlyContinue)
 if(@($tools|Where-Object ProcessName -ne 'MSBuild').Count){$blocked=$true}
 foreach($detail in @(Get-CimInstance Win32_Process -Filter "Name='MSBuild.exe'")){
  if($detail.CommandLine-notmatch'/nodemode:1 /nodeReuse:true'-or(Get-Process -Id $detail.ParentProcessId -ErrorAction SilentlyContinue)){$blocked=$true;continue}
  $p=Get-Process -Id $detail.ProcessId -ErrorAction SilentlyContinue
  if(!$p){$blocked=$true;continue}
  $nodes += [pscustomobject]@{pid=$p.Id;start_ticks=$p.StartTime.ToUniversalTime().Ticks;cpu_ticks=$p.TotalProcessorTime.Ticks;parent_pid=$detail.ParentProcessId;command_line=$detail.CommandLine;observed_utc=[DateTime]::UtcNow.ToString('o')}
 }
 $identity=(@($nodes|Sort-Object pid|ForEach-Object {'{0}:{1}:{2}'-f$_.pid,$_.start_ticks,$_.cpu_ticks})-join';')
 if($blocked-or$identity-ne$previous){$quiet.Restart()}
 $previous=$identity
 if(!$blocked-and$quiet.Elapsed.TotalSeconds-ge30){break}
 Start-Sleep -Seconds 2
}
ConvertTo-Json -InputObject @($nodes) -Depth 5|Set-Content (Join-Path $PSScriptRoot 'idle-node-exclusions.json')
ConvertTo-Json -InputObject @($nodes) -Depth 5|Set-Content (Join-Path $PSScriptRoot "$Phase-idle-node-exclusions.json")
[ordered]@{checked_utc=[DateTime]::UtcNow.ToString('o');minimum_stable_cpu_seconds=30;observed_idle_seconds=$quiet.Elapsed.TotalSeconds;node_count=$nodes.Count;passed=$true;note='Prepares exact dead-parent MSBuild node identities for the unchanged benchmark observer. The benchmark still requires its separate 30-second gate and samples the entire runner window.'}|ConvertTo-Json|Set-Content (Join-Path $PSScriptRoot "$Phase-idle-node-verification.json")
