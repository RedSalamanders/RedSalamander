Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$script:localSdkCandidateRepo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

Describe 'Explicit local SDK application build' {
    It 'passes the canonical solution target and unchanged source, descriptor and profile to the broker' {
        & {
            $root=$localSdkCandidateRepo
            Import-Module (Join-Path $root 'Tools/Modules/Build/BuildProjectSelection.psm1')
            $selection=Get-RSBuildProjectUnion -SolutionPath (Join-Path $root 'RedSalamander.sln') -ProjectNames RedSalamander
            $file=Join-Path $root 'Tools/Build-LocalSdkCandidate.ps1'
            $tokens=$null;$errors=$null
            $ast=[Management.Automation.Language.Parser]::ParseFile($file,[ref]$tokens,[ref]$errors)
            $errors.Count | Should Be 0
            function Invoke-CiLocalSdkApplicationBuild {
                param($ProjectRoot,$Descriptor,$ExpectedDescriptorSha256,$Platform,$MSBuildPath,$BuildContract,$BuildNumber)
                $script:localSdkReceived=$PSBoundParameters
                return 'original broker result'
            }
            $Descriptor='C:/synthetic-not-read/local snapshot.json';$DescriptorSha256='a'*64
            $Platform='ARM64';$BuildNumber=123;$MSBuildPath='C:/synthetic-not-launched/MSBuild.exe'
            foreach($name in @('$buildContract','$result')){
                $assignments=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -ceq $name},$true))
                $assignments.Count | Should Be 1
                . ([scriptblock]::Create($assignments[0].Extent.Text))
            }
            $localSdkReceived.ProjectRoot | Should Be $root
            $localSdkReceived.Descriptor | Should Be $Descriptor
            $localSdkReceived.ExpectedDescriptorSha256 | Should Be $DescriptorSha256
            $localSdkReceived.Platform | Should Be ARM64
            $localSdkReceived.BuildNumber | Should Be 123
            $localSdkReceived.MSBuildPath | Should Be $MSBuildPath
            $result | Should Be 'original broker result'
            $buildContract.Count | Should Be 9
            $buildContract.target | Should Be $selection.MSBuildTarget
            $buildContract.metadataProject | Should Be 'RedSalamander/RedSalamander.vcxproj'
            $buildContract.sdkProperty | Should Be RSDxUiSdkRoot
            $buildContract.outputProperty | Should Be RSOutputRoot
            $buildContract.testsProperty | Should Be RSBuildEnableTests
            $buildContract.versionProperty | Should Be RSVersionBuildNumber
        }
    }

    It 'refuses compiler-free execution before missing descriptor, tool or output discovery' {
        $prior=$env:CI_NATIVE_BUILDS_DISABLED
        try{
            $env:CI_NATIVE_BUILDS_DISABLED='true'
            { & (Join-Path $localSdkCandidateRepo 'Tools/Build-LocalSdkCandidate.ps1') -Descriptor 'missing' -DescriptorSha256 ('a'*64) -MSBuildPath 'missing' -BuildNumber 123 } | Should Throw 'require local Windows native execution'
        }finally{[Environment]::SetEnvironmentVariable('CI_NATIVE_BUILDS_DISABLED',$prior,'Process')}
    }

    It 'persists abandoned dependency evidence and releases its lock on refusal and success' {
        & {
            $tokens=$null;$errors=$null
            $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $localSdkCandidateRepo 'Tools/Build-LocalSdkCandidate.ps1'),[ref]$tokens,[ref]$errors)
            $blocks=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.TryStatementAst]},$true))
            $blocks.Count | Should Be 1
            $action=[scriptblock]::Create($blocks[0].Extent.Text)
            $root='C:/synthetic-not-read';$scope=@{configuration='Release';platform='x64';coordination='shared-dependency'}
            $Descriptor='missing';$DescriptorSha256='a'*64;$Platform='x64';$MSBuildPath='missing';$buildContract=@{};$BuildNumber=123
            function Set-RSArtifactOperationContaminated {param($RepoRoot,$Reason,$AbandonedOwner,$Scope)$script:localSdkLockProbe.marked++;$script:localSdkLockProbe.contaminated=$true}
            function Read-RSArtifactOperationContamination {param($RepoRoot,$Scope)return $script:localSdkLockProbe.contaminated}
            function Assert-RSNoResidualArtifactToolProcesses {param($RepoRoot,$Scope)$script:localSdkLockProbe.checked++}
            function Invoke-CiLocalSdkApplicationBuild {param($ProjectRoot,$Descriptor,$ExpectedDescriptorSha256,$Platform,$MSBuildPath,$BuildContract,$BuildNumber)$script:localSdkLockProbe.built++}
            function Exit-RSArtifactOperationLock {param($Lock)$script:localSdkLockProbe.closed++}
            foreach($state in @('abandoned','contaminated','clean')){
                $lease=[pscustomobject]@{WasAbandoned=($state -ceq 'abandoned');AbandonedOwner=@{source='synthetic'}}
                $script:localSdkLockProbe=@{marked=0;checked=0;built=0;closed=0;contaminated=($state -ceq 'contaminated')}
                if($state -ceq 'clean'){& $action}else{{& $action} | Should Throw 'refuses abandoned or contaminated dependency inputs'}
                $localSdkLockProbe.closed | Should Be 1
                $localSdkLockProbe.marked | Should Be $(if($state -ceq 'abandoned'){1}else{0})
                $localSdkLockProbe.built | Should Be $(if($state -ceq 'clean'){1}else{0})
                $localSdkLockProbe.checked | Should Be $(if($state -ceq 'clean'){1}else{0})
            }
        }
    }
}

Describe 'Actual local SDK output ownership' -Tag RequiresBuildToolchain {
    BeforeAll {
        Import-Module (Join-Path $localSdkCandidateRepo 'Tools/Modules/Testing/TestSandbox.psm1')
        Import-Module (Join-Path $localSdkCandidateRepo 'Tools/Modules/Build/BuildProjectSelection.psm1')
        # Metadata fixtures deliberately select the installed native VS instance;
        # they never provision tools, execute a receipt-supplied path or compile.
        $vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installation=(& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath).Trim()
        $hostFolder=if([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64'){'arm64'}else{'amd64'}
        $script:localSdkCandidateMSBuild=Join-Path $installation "MSBuild/Current/Bin/$hostFolder/MSBuild.exe"
    }

    It 'requires prepared terminal inputs without invoking the ordinary dependency builder in the local recipe' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $localSdkCandidateRepo -Harness LocalSdkTerminalInputs -Case ([guid]::NewGuid().ToString('N'))
        $source=[xml](Get-Content (Join-Path $localSdkCandidateRepo 'Directory.Build.targets') -Raw)
        $target=@($source.Project.Target|Where-Object Name -ceq RSPrepareTerminalRuntime)
        $target.Count | Should Be 1
        $project=Join-Path $fixture 'Terminal.vcxproj'
        [IO.File]::WriteAllText($project,('<Project>'+$target[0].OuterXml+'</Project>'))
        $tools=Join-Path $fixture 'Tools';[void][IO.Directory]::CreateDirectory($tools)
        $marker=Join-Path $tools 'unexpected-build.txt'
        [IO.File]::WriteAllText((Join-Path $tools 'Build-GhosttyTerminalRuntime.ps1'),'param($Platform) [IO.File]::WriteAllText((Join-Path $PSScriptRoot "unexpected-build.txt"),"builder invoked")')
        $arguments=@($project,'/nologo','/noautoresponse','/nr:false','/v:quiet','/t:RSPrepareTerminalRuntime','/p:RSIsFirstPartyProject=true','/p:Platform=x64')
        $failure=(& $localSdkCandidateMSBuild @arguments '/p:CiSdkContract=LocalUnpinnedAbiV1' 2>&1|Out-String)
        $LASTEXITCODE | Should Not Be 0
        $failure | Should Match 'Prepare terminal runtime inputs through the ordinary build'
        Test-Path -LiteralPath $marker | Should Be $false
        $runtime=Join-Path $fixture '.build/TerminalEngine/Product/x64/bin/ghostty-vt.dll'
        [void][IO.Directory]::CreateDirectory((Split-Path $runtime))
        [IO.File]::WriteAllText($runtime,'synthetic existence control; not runtime validation')
        & $localSdkCandidateMSBuild @arguments '/p:CiSdkContract=LocalUnpinnedAbiV1' *> $null
        $LASTEXITCODE | Should Be 0
        Test-Path -LiteralPath $marker | Should Be $false
        & $localSdkCandidateMSBuild @arguments *> $null
        $LASTEXITCODE | Should Be 0
        Test-Path -LiteralPath $marker | Should Be $true
    }

    It 'preserves canonical outputs, rejects downgraded alternate roots and reaches the live local output guard' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $localSdkCandidateRepo -Harness LocalSdkOutputGuard -Case ([guid]::NewGuid().ToString('N'))
        $source=[xml](Get-Content (Join-Path $localSdkCandidateRepo 'Directory.Build.targets') -Raw)
        $guard=@($source.Project.Target|Where-Object Name -ceq RSValidateOutputRoot)
        $guard.Count | Should Be 1
        $project=Join-Path $fixture 'OutputGuard.vcxproj'
        [IO.File]::WriteAllText($project,('<Project>'+$guard[0].OuterXml+'<Target Name="RSRequireLocalSdkOutputLease"><Error Text="ACTUAL_LOCAL_OUTPUT_GUARD_REACHED" /></Target></Project>'))
        function Invoke-RootOutputGuard([string]$Output,[string]$Contract,[bool]$Sdk){
            $arguments=@($project,'/nologo','/noautoresponse','/nr:false','/v:quiet','/t:RSValidateOutputRoot','/p:RSIsFirstPartyProject=true',('/p:RSOutputRoot='+$Output),('/p:CiSdkContract='+$Contract))
            if($Sdk){$arguments+=@('/p:RSDxUiSdkRoot=synthetic-sdk','/p:RSCiEngineRoot=synthetic-engine')}
            $text=(& $localSdkCandidateMSBuild @arguments 2>&1|Out-String)
            return @{exit=$LASTEXITCODE;text=$text}
        }
        (Invoke-RootOutputGuard (Join-Path $fixture '.build/') '' $false).exit | Should Be 0
        foreach($contract in @('','QualifiedAbiV2')){
            $refused=Invoke-RootOutputGuard (Join-Path $fixture 'alternate/') $contract $true
            $refused.exit | Should Not Be 0
            $refused.text | Should Match 'Separate RedSalamander output roots require'
        }
        $local=Invoke-RootOutputGuard (Join-Path $fixture 'alternate/') LocalUnpinnedAbiV1 $true
        $local.exit | Should Not Be 0
        $local.text | Should Match ACTUAL_LOCAL_OUTPUT_GUARD_REACHED
        Test-Path -LiteralPath (Join-Path $fixture 'alternate') | Should Be $false
    }

    It 'keeps every native solution project in isolated outputs while retaining ordinary dependency inputs' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $localSdkCandidateRepo -Harness LocalSdkOutputMetadata -Case ([guid]::NewGuid().ToString('N'))
        $sdk=Join-Path $fixture 'sdk';[void][IO.Directory]::CreateDirectory($sdk)
        [IO.File]::WriteAllText((Join-Path $sdk 'DxUi.Sdk.props'),'<Project><PropertyGroup><DxUiSdkRoot>$(MSBuildThisFileDirectory)</DxUiSdkRoot></PropertyGroup></Project>')
        $isolated=Join-Path $fixture 'outputs'
        $projects=@(Get-RSSolutionProjects -SolutionPath (Join-Path $localSdkCandidateRepo 'RedSalamander.sln') | Where-Object {$_.Path -and $_.RelativePath -like '*.vcxproj'})
        $projects.Count | Should BeGreaterThan 0
        foreach($project in $projects){
            $text=& $localSdkCandidateMSBuild $project.Path '/nologo' '/noautoresponse' '/nr:false' '/v:quiet' '/p:Platform=x64' '/p:Configuration=Release' ('/p:SolutionDir='+$localSdkCandidateRepo+'\') ('/p:RSOutputRoot='+$isolated+'\') ('/p:RSDxUiSdkRoot='+$sdk) ('/p:RSCiEngineRoot='+(Join-Path $localSdkCandidateRepo '.ci-engine')) '/p:CiSdkContract=LocalUnpinnedAbiV1' '-getProperty:OutDir,IntDir,VcpkgInstalledDir,VcpkgEnableManifest,VcpkgManifestInstall'
            $LASTEXITCODE | Should Be 0
            $metadata=($text -join "`n"|ConvertFrom-Json).Properties
            foreach($name in @('OutDir','IntDir')){[IO.Path]::GetFullPath($metadata.$name).StartsWith(($isolated+[IO.Path]::DirectorySeparatorChar),[StringComparison]::OrdinalIgnoreCase) | Should Be $true}
            [IO.Path]::GetFullPath($metadata.VcpkgInstalledDir).StartsWith((Join-Path $localSdkCandidateRepo '.build/vcpkg_installed/'),[StringComparison]::OrdinalIgnoreCase) | Should Be $true
            $metadata.VcpkgManifestInstall | Should Be false
            if($project.Name -in @('RedSalamander','Common','FileSystem7z','FileSystemDummy')){$metadata.VcpkgEnableManifest | Should Be true}
        }
        Test-Path -LiteralPath $isolated | Should Be $false
        Write-Host ('PASS isolated output/dependency metadata for '+$projects.Count+' native solution projects; no compilation or ABI acceptance.')
    }
}
