<#
.SYNOPSIS
Verifies the locally vendored, immutable RedMonitor client without network, checkout or build work.
.PARAMETER RepositoryRoot
Consumer repository containing Dependencies/RedMonitor.lock.json and ThirdParty/RedMonitor.
.DESCRIPTION
Checks or explicitly updates the minimal client dependency. Ordinary verification is local and read-only.
.OUTPUTS
Console result; the explicit updater also writes the two vendored inputs and immutable lock.
.NOTES
Never builds, stages or launches the independent RedMonitor helper. Does not terminate any process.
.EXAMPLE
./Tools/Verify-RedMonitorClient.ps1
#>
[CmdletBinding()]
param([string]$RepositoryRoot=(Split-Path $PSScriptRoot))
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath($RepositoryRoot)
$pin=Get-Content -LiteralPath (Join-Path $root 'Dependencies/RedMonitor.lock.json') -Raw|ConvertFrom-Json
if($pin.repository -cne 'https://github.com/RedSalamanders/RedMonitor' -or $pin.commit -cnotmatch '^[0-9a-f]{40}$' -or $pin.apiRevision -ne 1){
    throw 'Invalid immutable RedMonitor client lock.'
}
$expected=@('include/RedMonitor/Client.h','LICENSE.txt')
if($pin.files.Count -ne $expected.Count){throw 'Unexpected RedMonitor client file inventory.'}
foreach($relative in $expected){
    $entries=@($pin.files|Where-Object path -CEQ $relative)
    if($entries.Count -ne 1 -or $entries[0].sha256 -cnotmatch '^[A-Fa-f0-9]{64}$'){throw "Invalid client manifest entry: $relative"}
    $file=Join-Path $root ('ThirdParty/RedMonitor/'+$relative)
    if((Test-Path -LiteralPath $file) -and ((Get-Item -LiteralPath $file -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){
        throw "Reparse client input: $file"
    }
    for($directory=[IO.DirectoryInfo]::new((Split-Path $file));$directory;$directory=$directory.Parent){
        if($directory.Exists -and ($directory.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw "Reparse client directory: $file"}
    }
    if(-not(Test-Path -LiteralPath $file -PathType Leaf) -or (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -ine $entries[0].sha256){
        throw "Missing or modified RedMonitor client: $relative. Use the explicit client updater."
    }
}
Write-Host "Verified local RedMonitor client $($pin.commit); no viewer dependency."
