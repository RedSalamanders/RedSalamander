#Requires -Version 7.2
# Real imported compiler/linker profiles, pinned archives, and deliberately broken controls; no compilation or desktop use.
[CmdletBinding()]param([string]$Root='', [switch]$RequirePinnedArchives)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(-not $Root){$Root=(& git -C $PSScriptRoot rev-parse --show-toplevel).Trim()}
$Root=[IO.Path]::GetFullPath($Root)
if(-not $IsWindows){Write-Host 'SKIP Release build-policy evaluation: Windows MSVC required; native qualification remains open.';return}
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$installation=(& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath).Trim()
$msbuild=Join-Path $installation 'MSBuild/Current/Bin/amd64/MSBuild.exe'
if([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()-eq 'Arm64'){$msbuild=Join-Path $installation 'MSBuild/Current/Bin/arm64/MSBuild.exe'}
if(-not(Test-Path -LiteralPath $msbuild)){throw 'Install Visual Studio 2026 with MSVC 14.52 C++ tools.'}
$scratch=Join-Path $Root ('.build/ReleaseBuildPolicy/'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($scratch)
$dependencySource=''
$pinPath=Join-Path $Root 'Dependencies/DxUi.lock.json'
if(Test-Path -LiteralPath $pinPath){
    $pin=Get-Content -LiteralPath $pinPath -Raw|ConvertFrom-Json
    $candidateSource=Join-Path $Root ".build/dependencies/DxUi/source/$($pin.commit)"
    if(Test-Path -LiteralPath (Join-Path $candidateSource 'src/DxUi.vcxproj')){$dependencySource=$candidateSource}
    elseif($RequirePinnedArchives){throw 'Restore the exact DxUi source before required archive-policy evaluation.'}
    else{Write-Host 'Pinned archive evaluation pending: restore the exact DxUi source and rerun with -RequirePinnedArchives.'}
}
$probe=Join-Path $scratch 'Read.targets'
[IO.File]::WriteAllText($probe,'<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003"><ItemGroup><Link Include="ReleasePolicyLinkProbe"/></ItemGroup></Project>')
function Read-Profile([string]$Project,[string]$Configuration,[string]$Platform,[string]$ProbePath=$probe,[string[]]$Extra=@()){
    $arguments=@($Project,'/nologo',"/p:Configuration=$Configuration","/p:Platform=$Platform","/p:ForceImportAfterCppTargets=$ProbePath",'-getItem:ClCompile,Link,ProjectReference','-getProperty:ConfigurationType,VCToolsVersion,WholeProgramOptimization,LinkIncremental')+$Extra
    if($dependencySource){$arguments+="/p:DxUiRoot=$dependencySource"}
    $output=@(& $msbuild @arguments 2>&1)
    if($LASTEXITCODE-ne 0){throw "Evaluation failed: $Project $Configuration $Platform`n$($output -join "`n")"}
    return ($output -join "`n")|ConvertFrom-Json -AsHashtable
}
function Assert-Profile($Value,[string]$Configuration,[string]$Label,[switch]$PinnedArchive){
    if($Value.Properties.VCToolsVersion-notmatch '^14[.]52[.]\d+$'){throw "$Label selected unsupported MSVC $($Value.Properties.VCToolsVersion)."}
    foreach($item in $Value.Items.ClCompile){
        if($item['ExcludedFromBuild']-eq 'true'){continue}
        if($Configuration-eq 'Release'){
            $packaged=$item['FunctionLevelLinking']-eq 'true' -or ($PinnedArchive -and $item['FunctionLevelLinking']-ne 'false' -and $item['Optimization']-in @('MaxSpeed','MinSpace'))
            if(-not $packaged){throw "$Label $($item.Identity) disables function-level linking."}
            if($item['WholeProgramOptimization']-ne 'true'){throw "$Label $($item.Identity) disables /GL."}
            if($item['AdditionalOptions']-match '(?i)(?:^|\s)/(?:GL-|Gy-)(?:\s|$)'){throw "$Label overrides the Release compiler policy."}
        }elseif($item['WholeProgramOptimization']-eq 'true'){throw "$Label unexpectedly enables /GL in $Configuration."}
    }
    $link=@($Value.Items.Link|Where-Object Identity -eq 'ReleasePolicyLinkProbe')[0]
    if($Configuration-eq 'Release' -and $Value.Properties.ConfigurationType-in @('Application','DynamicLibrary')){
        if($link['OptimizeReferences']-ne 'true'){throw "$Label disables /OPT:REF."}
        if($link['LinkTimeCodeGeneration']-ne 'UseLinkTimeCodeGeneration'){throw "$Label does not use full /LTCG."}
        if($Value.Properties.LinkIncremental-eq 'true'){throw "$Label enables incremental linking with full LTCG."}
        if($link['AdditionalOptions']-match '(?i)/(?:OPT:NOREF|LTCG:OFF|INCREMENTAL(?:\s|$))'){throw "$Label overrides the Release linker policy."}
    }
    if($Configuration-ne 'Release' -and $link['LinkTimeCodeGeneration']-notin @($null,'','Default')){throw "$Label unexpectedly enables Release LTCG in $Configuration."}
}
$projects=@(& git -C $Root ls-files -- '*.vcxproj')
if($LASTEXITCODE-ne 0){throw 'Cannot enumerate tracked native projects.'}
$checked=0;$archiveChecked=@{};$first=$null
foreach($relative in $projects){
    $path=Join-Path $Root $relative
    [xml]$xml=Get-Content -LiteralPath $path -Raw
    if(-not $xml.SelectNodes('//*[local-name()="ClCompile"][@Include]').Count){continue}
    foreach($configuration in 'Debug','Release','ASan Debug'){foreach($platform in 'x64','ARM64'){
        $value=Read-Profile $path $configuration $platform
        Assert-Profile $value $configuration "$relative $platform $configuration"
        if($configuration-eq 'Release' -and $platform-eq 'x64' -and -not $first){$first=$path}
        foreach($reference in $value.Items.ProjectReference){
            if($reference.Identity-notmatch '[\\/]src[\\/]DxUi[.]vcxproj$'){continue}
            $archive=[IO.Path]::GetFullPath([IO.Path]::Combine([IO.Path]::GetDirectoryName($path),$reference.Identity))
            if($archive-eq [IO.Path]::GetFullPath((Join-Path $Root 'src/DxUi.vcxproj'))){continue}
            $properties=@{}
            foreach($pair in ([string]$reference['AdditionalProperties']).Split(';',[StringSplitOptions]::RemoveEmptyEntries)){
                $parts=$pair.Split('=',2);if($parts.Count-eq 2){$properties[$parts[0]]=$parts[1]}
            }
            $expected=if($configuration-eq 'Release'){'true'}else{'false'}
            if($properties['WholeProgramOptimization']-ne $expected -or $properties['VCToolsVersion']-ne $value.Properties.VCToolsVersion){throw "$relative does not forward exact toolset and WPO to pinned DxUi."}
            $key="$archive|$configuration|$platform"
            if(-not $archiveChecked.ContainsKey($key)){
                $extra=@("/p:WholeProgramOptimization=$expected","/p:VCToolsVersion=$($properties['VCToolsVersion'])")
                $child=Read-Profile $archive $configuration $platform -Extra $extra
                Assert-Profile $child $configuration "Pinned DxUi $platform $configuration" -PinnedArchive
                $archiveChecked[$key]=$true
            }
        }
        ++$checked
    }}
}
if(-not $first){throw 'No native Release profile was evaluated.'}
$negative=Join-Path $scratch 'Broken.targets'
[IO.File]::WriteAllText($negative,'<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003"><ItemGroup><ClCompile Update="@(ClCompile)"><WholeProgramOptimization>false</WholeProgramOptimization></ClCompile><Link Include="ReleasePolicyLinkProbe"><OptimizeReferences>false</OptimizeReferences></Link></ItemGroup></Project>')
$rejected=$false
try{Assert-Profile (Read-Profile $first 'Release' 'x64' $negative) 'Release' 'deliberately broken profile'}catch{if($_.Exception.Message-notmatch 'disables /GL'){throw};$rejected=$true}
if(-not $rejected){throw 'The evaluated compiler-negative control was accepted.'}
$broken=Read-Profile $first 'Release' 'x64'
foreach($item in $broken.Items.ClCompile){$item['WholeProgramOptimization']='true'}
$broken.Properties.ConfigurationType='Application'
foreach($item in $broken.Items.Link){$item['OptimizeReferences']='false'}
$rejected=$false
try{Assert-Profile $broken 'Release' 'broken final link'}catch{if($_.Exception.Message-notmatch 'disables /OPT:REF'){throw};$rejected=$true}
if(-not $rejected){throw 'The linker-negative control was accepted.'}
if($RequirePinnedArchives -and (Test-Path -LiteralPath $pinPath) -and $archiveChecked.Count-ne 6){throw 'Required pinned archive coverage is incomplete.'}
Write-Host "PASS Release policy: $checked native profiles, $($archiveChecked.Count) pinned-archive profiles, compiler/linker negative controls."

# Provisioning must remain read-only on local hosts and refuse a missing required family
# before downloading or starting any installer. The ordinary selected installation also
# has to satisfy the same SDK/target/ASan/LLVM check that CI uses.
$ciSetup=Join-Path $Root '.github/actions/setup-msvc/Setup.ps1'
& pwsh -NoProfile -File $ciSetup -VerifyOnly -Platform x64
if($LASTEXITCODE-ne 0){throw 'Read-only CI toolchain verification failed.'}
$fixture=Join-Path $scratch 'ci-provisioning-refusal'
$fixtureScript=Join-Path $fixture '.github/actions/setup-msvc/Setup.ps1'
[void][IO.Directory]::CreateDirectory((Split-Path $fixtureScript))
Copy-Item -LiteralPath $ciSetup -Destination $fixtureScript
[IO.File]::WriteAllText((Join-Path $fixture 'Directory.Build.props'),'<Project><PropertyGroup><RequiredMsvcToolsVersion>14.99</RequiredMsvcToolsVersion></PropertyGroup></Project>')
foreach($readOnly in @($true,$false)){
    $start=[Diagnostics.ProcessStartInfo]::new('pwsh')
    $start.UseShellExecute=$false;$start.CreateNoWindow=$true
    $start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
    foreach($argument in @('-NoProfile','-File',$fixtureScript,'-Platform','x64')){$start.ArgumentList.Add($argument)}
    if($readOnly){$start.ArgumentList.Add('-VerifyOnly')}
    $start.Environment['GITHUB_ACTIONS']='false'
    $process=[Diagnostics.Process]::new();$process.StartInfo=$start
    try{
        [void]$process.Start();$stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
        if(-not $process.WaitForExit(30000)){$process.Kill($true);throw 'Owned CI refusal control timed out.'}
        $message=$stdout.GetAwaiter().GetResult()+$stderr.GetAwaiter().GetResult()
        $expected=if($readOnly){'discovered installation lacks default MSVC 14.99'}else{'Automatic provisioning is restricted'}
        if($process.ExitCode-eq 0-or $message-notmatch $expected){throw "CI provisioning refusal control failed: $message"}
    }finally{$process.Dispose()}
}
Write-Host 'PASS CI toolchain verification and missing-family/local-install refusal controls.'