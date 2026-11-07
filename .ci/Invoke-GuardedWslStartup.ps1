# Private, separately approved WSL startup-console exception. Called only while
# the parent owns the canonical interactive mutex; no other window is changed.
function Invoke-PrivateGuardedWslStartupClose {
 param([Parameter(Mandatory)][string]$EvidenceDirectory,[Parameter(Mandatory)]$StockFailure,[Parameter(Mandatory)][bool]$Approved)
 if(-not $Approved -or $env:GITHUB_ACTIONS -cne 'true' -or $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or
    $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander' -or $env:GITHUB_EVENT_NAME -cne 'workflow_dispatch' -or
    [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -cne 'Arm64' -or [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString() -cne 'Arm64'){
  throw 'Private WSL exception requires explicit approval and the project-hosted native ARM64 manual role.'
 }
 $intent=[ordered]@{schema='private.wsl-startup-close-intent.v1';approved=$Approved;scope='exact System32 WSL update startup console only';created_utc=[DateTimeOffset]::UtcNow.ToString('o');coverage='NOT_EVALUATED'}
 $intentPath=Join-Path $EvidenceDirectory 'wsl-startup-intent.json';Write-PrivateCreateOnlyJson $intentPath $intent
 $result=[ordered]@{schema='private.wsl-startup-close-result.v1';intent_sha256=(Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash;closed=$false;process=$null;thread=$null;window=$null;session=$null;creation_utc=$null;argument_count=$null;arguments_match=$false;error=$null;finished_utc=$null}
 $caught=$null;$window=$null
 try {
  $context=(Get-RSHostedDesktopDiagnostics).DesktopContext
  if($context.WindowStation -cne 'WinSta0' -or $context.ThreadDesktop -cne 'Default' -or $context.InputDesktop -cne 'Default' -or
     $context.WindowStationError -ne 0 -or $context.ThreadDesktopError -ne 0 -or $context.InputDesktopError -ne 0){throw 'Caller and input must be WinSta0/Default.'}
  if(-not ('PrivateWslStartup.CapturedWindow' -as [type])){Add-Type -Path (Join-Path $PSScriptRoot 'GuardedWslStartup.cs')}
  $window=[PrivateWslStartup.CapturedWindow]::Capture()
  $result.process=$window.ProcessId;$result.thread=$window.ThreadId;$result.window=$window.Window;$result.session=$window.SessionId;$result.creation_utc=$window.CreationUtc.ToString('o')
  if($StockFailure.last_observation.kind -cne 'Unknown' -or [long]$StockFailure.last_observation.window -ne $window.Window -or
     [uint32]$StockFailure.last_observation.process -ne $window.ProcessId -or [uint32]$StockFailure.last_observation.thread -ne $window.ThreadId -or
     [uint32]$StockFailure.last_observation.session -ne $window.SessionId){throw 'WSL identity does not match the saved stock-preparation refusal.'}
  $current=Get-CimInstance -ClassName Win32_Process -Filter ("ProcessId = {0}" -f $window.ProcessId)
  if($null -eq $current -or [uint32]$current.ProcessId -ne $window.ProcessId -or $null -eq $current.CreationDate -or
     [Math]::Abs(($current.CreationDate.ToUniversalTime() - $window.CreationUtc).Ticks) -ge 10 -or
     -not [string]::Equals([string]$current.ExecutablePath, (Join-Path $env:SystemRoot 'System32/wsl.exe'),[StringComparison]::OrdinalIgnoreCase)){throw 'Retained WSL process and CIM creation/image identity do not match.'}
  $count=0;$matches=[PrivateWslStartup.CapturedWindow]::ExactArguments([string]$current.CommandLine,[ref]$count)
  $result.argument_count=$count;$result.arguments_match=$matches
  if(-not $matches){throw 'WSL startup command arguments do not match the approved update sequence.'}
  $context=(Get-RSHostedDesktopDiagnostics).DesktopContext
  if($context.WindowStation -cne 'WinSta0' -or $context.ThreadDesktop -cne 'Default' -or $context.InputDesktop -cne 'Default' -or
     $context.WindowStationError -ne 0 -or $context.ThreadDesktopError -ne 0 -or $context.InputDesktopError -ne 0){throw 'Caller or input desktop changed before WSL close.'}
  $window.CloseOnce();$result.closed=$true
 }catch{$caught=$_;$result.error=$_.Exception.Message}
 finally {
  if($window){$window.Dispose()}
  $result.finished_utc=[DateTimeOffset]::UtcNow.ToString('o')
  try{Write-PrivateCreateOnlyJson (Join-Path $EvidenceDirectory 'wsl-startup-result.json') $result}catch{
   if($caught){throw [AggregateException]::new('WSL guard and audit result both failed.',[Exception[]]@($caught.Exception,$_.Exception))};throw
  }
 }
 if($caught){throw $caught}
}
