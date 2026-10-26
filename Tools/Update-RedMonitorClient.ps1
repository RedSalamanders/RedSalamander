<#
.SYNOPSIS
Explicitly updates only the RedMonitor client header, license and immutable consumer lock.
.DESCRIPTION
Uses authenticated gh contents APIs at one full commit. Ordinary builds never run this updater.
.PARAMETER Commit
Full immutable RedSalamanders/RedMonitor commit reviewed for adoption.
.PARAMETER RepositoryRoot
Consumer repository to update. Review and commit the resulting header/license/lock together.
.OUTPUTS
Console result; the explicit updater also writes the two vendored inputs and immutable lock.
.NOTES
Never builds, stages or launches the independent RedMonitor helper. Does not terminate any process.
.EXAMPLE
./Tools/Update-RedMonitorClient.ps1 -Commit 0000000000000000000000000000000000000000
#>
[CmdletBinding(SupportsShouldProcess)]
param([Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{40}$')][string]$Commit,[string]$RepositoryRoot=(Split-Path $PSScriptRoot))
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath($RepositoryRoot)
$files=@('include/RedMonitor/Client.h','LICENSE.txt')
$payload=@(foreach($path in $files){
    $response=& gh api "repos/RedSalamanders/RedMonitor/contents/$($path)?ref=$Commit"
    if($LASTEXITCODE -ne 0){throw "Cannot fetch the pinned client input: $path"}
    $item=($response -join [Environment]::NewLine)|ConvertFrom-Json
    if($item.type -cne 'file' -or $item.encoding -cne 'base64' -or $item.path -cne $path){throw 'Unexpected client contents response.'}
    $bytes=[Convert]::FromBase64String($item.content)
    [pscustomobject]@{path=$path;bytes=$bytes;sha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)).ToLowerInvariant()}
})
if([Text.Encoding]::UTF8.GetString($payload[0].bytes) -notmatch 'kApiRevision\s*=\s*1\s*;'){throw 'Unsupported RedMonitor client API.'}
$vendor=Join-Path $root 'ThirdParty/RedMonitor'
for($current=[IO.DirectoryInfo]::new($vendor);$current;$current=$current.Parent){
    if($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Reparse client update destination.'}
}
if(-not $PSCmdlet.ShouldProcess($vendor,"Update client at $Commit")){return}
foreach($file in $payload){
    $destination=Join-Path $vendor $file.path
    New-Item -ItemType Directory -Force -Path (Split-Path $destination)|Out-Null
    if(Test-Path -LiteralPath $destination){
        if((Get-Item -LiteralPath $destination -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Reparse client file destination.'}
    }
    [IO.File]::WriteAllBytes($destination,$file.bytes)
}
[IO.File]::WriteAllText((Join-Path $vendor '.gitattributes'),"* -text"+[Environment]::NewLine)
New-Item -ItemType Directory -Force -Path (Join-Path $root 'Dependencies')|Out-Null
$lock=[ordered]@{repository='https://github.com/RedSalamanders/RedMonitor';commit=$Commit;apiRevision=1;files=@($payload|ForEach-Object{[ordered]@{path=$_.path;sha256=$_.sha256}})}
$lock|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $root 'Dependencies/RedMonitor.lock.json') -Encoding utf8NoBOM
Write-Host "Updated header-only RedMonitor client to $Commit. No viewer source or binaries were fetched."
