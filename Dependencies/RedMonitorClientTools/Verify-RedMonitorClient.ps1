<#
.SYNOPSIS
Verifies the locally vendored, immutable RedMonitor client without network, checkout or build work.
.DESCRIPTION
Runs on Windows PowerShell 5.1 (the MSBuild RequirePinnedRedMonitorClient gate, so IDE builds need no
PowerShell 7) and on PowerShell 7. Links inside the checkout are refused; folders above it are the user's.
.PARAMETER RepositoryRoot
Consumer repository containing Dependencies/RedMonitor.lock.json and ThirdParty/RedMonitor. A relative path
resolves against the caller's PowerShell location.
#>
[CmdletBinding()]
param([string]$RepositoryRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
# Windows PowerShell leaves $PSScriptRoot empty in a parameter default; MSBuild runs this from the project folder.
if(-not $RepositoryRoot){$RepositoryRoot=$PSScriptRoot}
$root=[IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RepositoryRoot)).TrimEnd('\','/')
function Assert-CheckoutPath([string]$Path){
    # From the input up to and including the root, before anything is read. A junction or synced folder above the
    # root (where the user chose to clone) is not inspected.
    $current=[IO.Path]::GetFullPath($Path)
    while($true){
        if((Test-Path -LiteralPath $current) -and ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){
            throw "Reparse RedMonitor client path: $current"
        }
        if([string]::Equals($current.TrimEnd('\','/'),$root,[StringComparison]::OrdinalIgnoreCase)){return}
        $current=Split-Path -Parent $current
        if(-not $current){throw "RedMonitor client path outside the checkout: $Path"}
    }
}
function Get-Sha256([string]$Path){
    # .NET directly: Windows PowerShell's Get-FileHash is a module function that a PowerShell 7 PSModulePath
    # inherited through build.ps1 and MSBuild hides.
    $sha=[Security.Cryptography.SHA256]::Create()
    try{
        $stream=[IO.File]::OpenRead($Path)
        try{return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')}finally{$stream.Dispose()}
    }finally{$sha.Dispose()}
}
$lock=Join-Path $root 'Dependencies/RedMonitor.lock.json'
Assert-CheckoutPath $lock
$pin=Get-Content -LiteralPath $lock -Raw|ConvertFrom-Json
if($pin.repository -cne 'https://github.com/RedSalamanders/RedMonitor' -or $pin.commit -cnotmatch '^[0-9a-f]{40}$' -or $pin.apiRevision -ne 1){
    throw 'Invalid immutable RedMonitor client lock.'
}
$expected=@('include/RedMonitor/Client.h','LICENSE.txt')
if(@($pin.files).Count -ne $expected.Count){throw 'Unexpected RedMonitor client file inventory.'}
foreach($relative in $expected){
    $entries=@($pin.files|Where-Object path -CEQ $relative)
    if($entries.Count -ne 1 -or $entries[0].sha256 -cnotmatch '^[A-Fa-f0-9]{64}$'){throw "Invalid client manifest entry: $relative"}
    $file=Join-Path $root ('ThirdParty/RedMonitor/'+$relative)
    Assert-CheckoutPath $file
    if(-not(Test-Path -LiteralPath $file -PathType Leaf) -or (Get-Sha256 $file) -ine $entries[0].sha256){
        throw "Missing or modified RedMonitor client: $relative. Use the explicit client updater."
    }
}
Write-Host "Verified local RedMonitor client $($pin.commit); no viewer dependency."
