[CmdletBinding()]param([string]$Root='', [string]$ScratchRoot='')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
if(-not $Root){$Root=(& git -C $PSScriptRoot rev-parse --show-toplevel).Trim()}
if(-not $ScratchRoot){
    if((Split-Path $Root -Leaf)-eq 'bp28' -and $Root -like '*RedSalamander.Perf*' -or
       (Test-Path -LiteralPath (Join-Path $Root 'RedSalamander/RedSalamander.vcxproj'))){
        Import-Module (Join-Path $Root 'Tools/Modules/Testing/TestSandbox.psm1') -Scope Local
        $ScratchRoot=New-RSTestSandboxScratchDirectory -RepoRoot $Root -Harness BuildConsole -Case Policy -RunId ('console-'+[guid]::NewGuid().ToString('N'))
    }else{$ScratchRoot=Join-Path $Root ('.build/BuildConsoleTests/'+[guid]::NewGuid().ToString('N'))}
}
[void][IO.Directory]::CreateDirectory($ScratchRoot)
Import-Module (Join-Path $Root 'Build/BuildConsole.psm1') -Scope Local -ErrorAction Stop
function Require([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message}}
function Refuses([scriptblock]$Body,[string]$Pattern){
    $failed=$false
    try{& $Body}catch{if($_.Exception.Message-notmatch $Pattern){throw};$failed=$true}
    Require $failed "Expected refusal: $Pattern"
}
$clean=Join-Path $ScratchRoot 'clean.log'
[IO.File]::WriteAllText($clean,'Build succeeded. 0 Warning(s) 0 Error(s)')
Assert-NativeBuildDiagnostics -LogPath $clean
foreach($diagnostic in @(
    'source.cpp(4): warning C4100: unused parameter',
    '2>C:\work\source.cpp(4): warning C4365: sign conversion',
    'LINK : warning LNK4099: missing PDB',
    'cl : Command line warning D9025 : overriding option',
    'project.targets(8,3): warning MSB1234: build diagnostic',
    'warning : custom build diagnostic'
)){
    $warn=Join-Path $ScratchRoot 'warning.log'
    [IO.File]::WriteAllText($warn,$diagnostic)
    $facts=Get-NativeBuildDiagnostics -LogPath $warn
    Require ($facts.WarningCount-eq 1 -and $facts.ErrorCount-eq 0) "Missed diagnostic: $diagnostic"
    Refuses {Assert-NativeBuildDiagnostics -LogPath $warn} 'Zero-warning'
}
$errorLog=Join-Path $ScratchRoot 'error.log'
[IO.File]::WriteAllText($errorLog,'file.cpp(1): fatal error C1083: missing header')
Refuses {Assert-NativeBuildDiagnostics -LogPath $errorLog} 'Zero-warning'
Refuses {Assert-NativeBuildDiagnostics -LogPath (Join-Path $ScratchRoot 'missing.log')} 'missing'
$binary=Join-Path $ScratchRoot "synthetic # app's É.exe"
# Deliberately not executable: writing a result must never launch it.
[IO.File]::WriteAllText($binary,'A synthetic artifact, never run.')
$plain=ConvertTo-NativeBuildLink -LiteralPath $binary -HyperlinkMode Never
Require ($plain -ceq [IO.Path]::GetFullPath($binary)) 'Plain link mode changed the path.'
$link=ConvertTo-NativeBuildLink -LiteralPath $binary -HyperlinkMode Always
Require ($link.Contains('file://localhost/') -and $link.Contains('%20') -and $link.Contains('%23')) 'File hyperlink was not URI-escaped and locally hosted.'
Require ($link.Contains([string][char]27+']8;;')) 'OSC 8 opening sequence missing.'
Refuses {ConvertTo-NativeBuildLink -LiteralPath ('unsafe'+[char]27+'path') -HyperlinkMode Always} 'control characters'
$records=@(Write-NativeBuildResult -Project Fixture -Platform x64 -Configuration Release -LogPath $clean -Artifacts @($binary) -Executable $binary -ColorMode Always -HyperlinkMode Always 6>&1)
$messages=@($records|ForEach-Object{[string]$_.MessageData})
Require (@($messages|Where-Object{$_ -like 'BUILD SUCCEEDED:*'}).Count-eq 1) 'Common success header missing.'
Require (@($messages|Where-Object{$_ -like 'Run:*'}).Count-eq 1) 'Executable link missing.'
Require (@($messages|Where-Object{$_ -like "& *''*"}).Count-eq 1) 'Quoted executable fallback must escape apostrophes.'
foreach($record in $records){Require ([string]$record.MessageData.ForegroundColor -eq 'Green') 'Successful result line is not green.'}
Refuses {Write-NativeBuildResult -Project Fixture -Platform x64 -Configuration Release -LogPath $errorLog -Artifacts @($binary)} 'Zero-warning'
Refuses {Write-NativeBuildResult -Project Fixture -Platform x64 -Configuration Release -LogPath $clean -Artifacts @((Join-Path $ScratchRoot 'absent.exe'))} 'missing'

$paths=if(Test-Path -LiteralPath (Join-Path $Root 'Build/DxUiPaths.psm1')){'Build/DxUiPaths.psm1'}
       elseif(Test-Path -LiteralPath (Join-Path $Root 'Tools/Modules/Build/DxUiPaths.psm1')){'Tools/Modules/Build/DxUiPaths.psm1'}else{''}
if($paths){
    Import-Module (Join-Path $Root $paths) -Scope Local -ErrorAction Stop
    $dependency=Join-Path $ScratchRoot 'dependency'
    $commit='0123456789abcdef'+('1'*24)
    $source=Get-DxUiSourcePath -DependencyRoot $dependency -Commit $commit
    Require ((Split-Path $source -Leaf)-ceq '0123456789abcdef') 'Source folder is not the agreed 16-character prefix.'
    $fingerprint=('a'*16)+('b'*48)
    $identity=[ordered]@{Fingerprint=$fingerprint;Identity=[ordered]@{commit=$commit;platform='x64'}}
    $first=Initialize-DxUiOutputRoot -DependencyRoot $dependency -BuildIdentity $identity
    Require ((Split-Path $first.TrimEnd('\') -Leaf)-ceq ('a'*16)) 'Output folder is not shortened.'
    $again=Initialize-DxUiOutputRoot -DependencyRoot $dependency -BuildIdentity $identity
    Require ($first-ceq $again) 'Identical identity should reuse its verified root.'
    $marker=Join-Path $first 'DxUi.build-identity.json'
    $before=(Get-FileHash -LiteralPath $marker -Algorithm SHA256).Hash
    $collision=[ordered]@{Fingerprint=('a'*16)+('c'*48);Identity=[ordered]@{commit=$commit;platform='x64'}}
    Refuses {Initialize-DxUiOutputRoot -DependencyRoot $dependency -BuildIdentity $collision} 'collision'
    Require ((Get-FileHash -LiteralPath $marker -Algorithm SHA256).Hash-ceq $before) 'Collision changed existing identity data.'
    $unverified=Join-Path $dependency ('d'*16)
    [void][IO.Directory]::CreateDirectory($unverified)
    [IO.File]::WriteAllText((Join-Path $unverified 'foreign.txt'),'Retain this data.')
    $foreign=[ordered]@{Fingerprint=('d'*64);Identity=[ordered]@{commit=$commit;platform='x64'}}
    Refuses {Initialize-DxUiOutputRoot -DependencyRoot $dependency -BuildIdentity $foreign} 'unverified'
    Require (Test-Path -LiteralPath (Join-Path $unverified 'foreign.txt')) 'Unverified output was removed.'
}
Write-Host 'PASS zero-warning results, green output, safe terminal links, plain launch commands and short-path collision refusal.'
