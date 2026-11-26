<#
.SYNOPSIS
Explicitly updates only the RedMonitor client header, license and immutable consumer lock.
.DESCRIPTION
Uses authenticated gh contents APIs at one full commit. Preserves an existing tooling pin.
.PARAMETER Commit
Full immutable RedSalamanders/RedMonitor commit reviewed for adoption.
.PARAMETER RepositoryRoot
Consumer repository to update. Review and commit the resulting header/license/lock together.
#>
[CmdletBinding(SupportsShouldProcess)]
param([Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{40}$')][string]$Commit,[string]$RepositoryRoot=$PSScriptRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
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
if(-not $PSCmdlet.ShouldProcess($vendor,"Update client at $Commit")){return}
$lockPath=Join-Path $root 'Dependencies/RedMonitor.lock.json'
$attributesPath=Join-Path $vendor '.gitattributes'
Assert-CheckoutPath $lockPath
Assert-CheckoutPath $attributesPath
foreach($path in $files){Assert-CheckoutPath (Join-Path $vendor $path)}
$lock=[ordered]@{repository='https://github.com/RedSalamanders/RedMonitor';commit=$Commit;apiRevision=1;files=@($payload|ForEach-Object{[ordered]@{path=$_.path;sha256=$_.sha256}})}
if(Test-Path -LiteralPath $lockPath){
    $existingFile=Get-Item -LiteralPath $lockPath -Force
    if($existingFile.PSIsContainer -or ($existingFile.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Invalid client lock file.'}
    $existing=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($lockPath,[Text.UTF8Encoding]::new($false,$true))) -AsHashtable
    if($existing -isnot [Collections.IDictionary]){throw 'Invalid client lock document.'}
    if($existing.Contains('tooling')){$lock['tooling']=$existing['tooling']}
}
foreach($file in $payload){
    $destination=Join-Path $vendor $file.path
    New-Item -ItemType Directory -Force -Path (Split-Path $destination)|Out-Null
    [IO.File]::WriteAllBytes($destination,$file.bytes)
}
[IO.File]::WriteAllText($attributesPath,"* -text"+[Environment]::NewLine)
New-Item -ItemType Directory -Force -Path (Join-Path $root 'Dependencies')|Out-Null
$lock|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $lockPath -Encoding utf8NoBOM
Write-Host "Updated header-only RedMonitor client to $Commit. No viewer source or binaries were fetched."
