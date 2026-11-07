# Private read-only session/desktop observations; no input or state transition.
function Initialize-PrivateHostedInputCapability {
 if (-not ('PrivateHosted.InputCapability' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'HostedInputCapability.cs') }
 if (-not ('PrivateHosted.InputSecurity' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'HostedInputSecurity.cs') }
}
function Get-PrivateHostedInputCapability {
 Initialize-PrivateHostedInputCapability
 $names=$null;$namesError=$null
 try{$names=(Get-RSHostedDesktopDiagnostics).DesktopContext}catch{$namesError=$_.Exception.ToString()}
 [ordered]@{schema='private.hosted-input-capability.v1';observed_utc=[DateTimeOffset]::UtcNow.ToString('o');native=[PrivateHosted.InputCapability]::Read();security=[PrivateHosted.InputSecurity]::Read();desktop_context=$names;desktop_context_error=$namesError;interpretation='Read-only; desktop name alone does not establish a connected session.'}
}
