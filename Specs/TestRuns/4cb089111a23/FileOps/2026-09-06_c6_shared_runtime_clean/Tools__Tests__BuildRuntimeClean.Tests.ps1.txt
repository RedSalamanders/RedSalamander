Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/ProcessStreaming.psm1') -ErrorAction Stop

# This fixture models shared-output ownership, not a compiled product. It imports
# the installed MSBuild cleanup implementation and the actual repository targets;
# no production output, vcpkg install, or toolchain file is modified.
function New-RSSharedRuntimeCleanFixture {
    $root = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness build-runtime-clean -Case ([Guid]::NewGuid().ToString('N'))
    foreach ($directory in @('project', 'output', 'output-sibling', 'intermediate', 'installed')) {
        [void](New-Item -ItemType Directory -Path (Join-Path $root $directory))
    }
    $files = [ordered]@{
        Own = 'output/Consumer.dll'
        Shared = 'output/z.dll'
        SharedUpper = 'output/SQLITE3.DLL'
        Private = 'intermediate/private.dll'
        Source = 'installed/z.dll'
        Sibling = 'output-sibling/z.dll'
        Orphan = 'output/Consumer.old.pdb'
    }
    foreach ($key in @($files.Keys)) {
        $files[$key] = Join-Path $root $files[$key]
        [IO.File]::WriteAllText($files[$key], "owned-fixture-$key", [Text.UTF8Encoding]::new($false))
    }
    $history = Join-Path $root 'intermediate/Consumer.FileListAbsolute.txt'
    [IO.File]::WriteAllLines($history, [string[]]@($files.Values), [Text.UTF8Encoding]::new($false))
    $project = Join-Path $root 'project/Consumer.vcxproj'
    $targetsPath = [Security.SecurityElement]::Escape((Join-Path $repoRoot 'Directory.Build.targets'))
    $testTargetsPath = [Security.SecurityElement]::Escape((Join-Path $repoRoot 'Tests/Directory.Build.targets'))
    $source = @'
<Project>
  <PropertyGroup>
    <RSFixtureRoot>$(MSBuildProjectDirectory)\..\</RSFixtureRoot>
    <OutDir>$([MSBuild]::NormalizeDirectory('$(RSFixtureRoot)output'))</OutDir>
    <OutputPath>$(OutDir)</OutputPath>
    <IntermediateOutputPath>$([MSBuild]::NormalizeDirectory('$(RSFixtureRoot)intermediate'))</IntermediateOutputPath>
    <IntDir>$(IntermediateOutputPath)</IntDir>
    <TargetName>Consumer</TargetName>
    <AssemblyName>Consumer</AssemblyName>
    <OutputType>Library</OutputType>
    <TargetPath>$(OutDir)Consumer.dll</TargetPath>
    <CleanFile>Consumer.FileListAbsolute.txt</CleanFile>
    <RSIsFirstPartyProject Condition="'$(RSIsFirstPartyProject)'==''">true</RSIsFirstPartyProject>
    <RSVersioningEnabled>false</RSVersioningEnabled>
    <ConfigurationType>DynamicLibrary</ConfigurationType>
    <RSIsTestProject Condition="'$(RSFixtureDisabledTests)'=='true'">true</RSIsTestProject>
    <RSBuildEnableTests Condition="'$(RSFixtureDisabledTests)'=='true'">false</RSBuildEnableTests>
  </PropertyGroup>
  <Import Project="$(MSBuildBinPath)\Microsoft.Common.CurrentVersion.targets" />
  <Import Project="__ROOT_TARGETS__" Condition="'$(RSFixtureDisabledTests)'!='true'" />
  <Import Project="__TEST_TARGETS__" Condition="'$(RSFixtureDisabledTests)'=='true'" />
  <!-- Deliberately registered AFTER the repo hook: vcpkg's link-skipped path
       also supplies FileWrites immediately before this common MSBuild target. -->
  <Target Name="FixtureRegisterWrites" BeforeTargets="_CleanGetCurrentAndPriorFileWrites">
    <ItemGroup>
      <FileWrites Include="$(TargetPath);$(IntDir)private.dll" />
      <FileWrites Include="$(OutDir)z.dll;$(OutDir)SQLITE3.DLL" Condition="'$(RSFixtureProducer)'=='link' or '$(RSFixtureProducer)'=='preserve'" />
    </ItemGroup>
  </Target>
</Project>
'@
    [IO.File]::WriteAllText($project, $source.Replace('__ROOT_TARGETS__', $targetsPath).Replace('__TEST_TARGETS__', $testTargetsPath), [Text.UTF8Encoding]::new($false))
    return [pscustomobject]@{ Root = $root; Project = $project; History = $history; Files = $files }
}

function Invoke-RSSharedRuntimeCleanFixture {
    param($Fixture, [string]$Target, [string[]]$Properties = @())
    $log = Join-Path $Fixture.Root (([Guid]::NewGuid().ToString('N')) + '.log')
    $arguments = @('-NoProfile', '-File', (Join-Path $repoRoot 'Tools/Invoke-SanitizedMsbuild.ps1'),
        $Fixture.Project, "/t:$Target", "/p:Configuration=$env:RS_VALIDATION_CONFIGURATION", "/p:Platform=$env:RS_VALIDATION_PLATFORM",
        '/nologo', '/v:minimal', '/nr:false') + $Properties
    $exitCode = Invoke-RSStreamingProcess -FilePath 'pwsh.exe' -Arguments $arguments -WorkingDirectory $repoRoot `
        -LogPath $log -TimeoutMilliseconds 60000 -DelegateArtifactOperation
    if ($exitCode -ne 0) { throw "MSBuild fixture failed ($exitCode). See $log" }
    (Get-Content -LiteralPath $log -Raw) | Should Not Match '(?im)\bwarning\s*(?:[A-Z]+\d+\s*)?:'
}

Describe 'Shared runtime DLL Clean ownership' -Tag RequiresBuildToolchain {
    BeforeAll {
        if ($env:RS_VALIDATION_PLATFORM -notin @('x64', 'ARM64') -or
            $env:RS_VALIDATION_CONFIGURATION -notin @('Debug', 'Release', 'ASan Debug')) {
            throw 'Select RS_VALIDATION_PLATFORM and RS_VALIDATION_CONFIGURATION before invoking build-toolchain coverage.'
        }
    }

    It 'migrates old explicit Clean records without deleting another project dependency' {
        $fixture = New-RSSharedRuntimeCleanFixture
        Invoke-RSSharedRuntimeCleanFixture $fixture CoreClean
        foreach ($key in @('Shared', 'SharedUpper', 'Source', 'Sibling')) {
            (Get-Content -LiteralPath $fixture.Files[$key] -Raw) | Should Be "owned-fixture-$key"
        }
        foreach ($key in @('Own', 'Private', 'Orphan')) {
            Test-Path -LiteralPath $fixture.Files[$key] | Should Be $false
        }
    }

    It 'preserves DLLs used by an earlier rebuilt project even when deletion is denied' {
        $fixture = New-RSSharedRuntimeCleanFixture
        $lease = [IO.File]::Open($fixture.Files.Shared, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        try { Invoke-RSSharedRuntimeCleanFixture $fixture CoreClean }
        finally { $lease.Dispose() }
        Test-Path -LiteralPath $fixture.Files.Own | Should Be $false
        (Get-Content -LiteralPath $fixture.Files.Shared -Raw) | Should Be 'owned-fixture-Shared'
    }

    It 'migrates orphaned shared entries before incremental deletion' {
        $fixture = New-RSSharedRuntimeCleanFixture
        Invoke-RSSharedRuntimeCleanFixture $fixture IncrementalClean
        foreach ($key in @('Own', 'Shared', 'SharedUpper', 'Private', 'Source', 'Sibling')) {
            Test-Path -LiteralPath $fixture.Files[$key] | Should Be $true
        }
        Test-Path -LiteralPath $fixture.Files.Orphan | Should Be $false
        $history = @(Get-Content -LiteralPath $fixture.History)
        ($history -contains $fixture.Files.Shared) | Should Be $false
        ($history -contains $fixture.Files.SharedUpper) | Should Be $false
    }

    It 'excludes link and link-skipped dependency writes from new cleanup ownership' {
        foreach ($producer in @('link', 'preserve')) {
            foreach ($target in @('IncrementalClean', '_CleanRecordFileWrites')) {
                $fixture = New-RSSharedRuntimeCleanFixture
                Invoke-RSSharedRuntimeCleanFixture $fixture $target @("/p:RSFixtureProducer=$producer")
                $history = @(Get-Content -LiteralPath $fixture.History)
                ($history -contains $fixture.Files.Shared) | Should Be $false
                ($history -contains $fixture.Files.SharedUpper) | Should Be $false
                ($history -contains $fixture.Files.Own) | Should Be $true
                ($history -contains $fixture.Files.Private) | Should Be $true
                Test-Path -LiteralPath $fixture.Files.Shared | Should Be $true
            }
        }
    }

    It 'allows late consumer Clean after another project has staged a newer dependency' {
        $fixture = New-RSSharedRuntimeCleanFixture
        Invoke-RSSharedRuntimeCleanFixture $fixture _CleanRecordFileWrites @('/p:RSFixtureProducer=preserve')
        [IO.File]::WriteAllText($fixture.Files.Shared, 'newer-producer-bytes', [Text.UTF8Encoding]::new($false))
        Invoke-RSSharedRuntimeCleanFixture $fixture CoreClean
        (Get-Content -LiteralPath $fixture.Files.Shared -Raw) | Should Be 'newer-producer-bytes'
        Test-Path -LiteralPath $fixture.Files.Own | Should Be $false
    }

    It 'records only owned outputs when there is no prior cleanup history' {
        $fixture = New-RSSharedRuntimeCleanFixture
        Remove-Item -LiteralPath $fixture.History
        Invoke-RSSharedRuntimeCleanFixture $fixture _CleanRecordFileWrites @('/p:RSFixtureProducer=link')
        $history = @(Get-Content -LiteralPath $fixture.History)
        ($history -contains $fixture.Files.Own) | Should Be $true
        ($history -contains $fixture.Files.Private) | Should Be $true
        ($history -contains $fixture.Files.Shared) | Should Be $false
        ($history -contains $fixture.Files.SharedUpper) | Should Be $false
    }

    It 'normalizes path spelling while retaining exact target and private DLL ownership' {
        $fixture = New-RSSharedRuntimeCleanFixture
        $alias = Join-Path $fixture.Root 'output/../output/Z.DLL'
        [IO.File]::AppendAllText($fixture.History, "$alias`n", [Text.UTF8Encoding]::new($false))
        Invoke-RSSharedRuntimeCleanFixture $fixture CoreClean
        Test-Path -LiteralPath $fixture.Files.Shared | Should Be $true
        Test-Path -LiteralPath $fixture.Files.Own | Should Be $false
        Test-Path -LiteralPath $fixture.Files.Private | Should Be $false
    }

    It 'protects shared runtimes when a prior test project becomes a production utility' {
        $fixture = New-RSSharedRuntimeCleanFixture
        Invoke-RSSharedRuntimeCleanFixture $fixture 'CoreClean;RSCleanDisabledTestOutputs' @('/p:RSFixtureDisabledTests=true')
        Test-Path -LiteralPath $fixture.Files.Shared | Should Be $true
        Test-Path -LiteralPath $fixture.Files.SharedUpper | Should Be $true
        Test-Path -LiteralPath $fixture.Files.Own | Should Be $false
    }

    It 'does not change cleanup behavior for a non-first-party project' {
        $fixture = New-RSSharedRuntimeCleanFixture
        Invoke-RSSharedRuntimeCleanFixture $fixture CoreClean @('/p:RSIsFirstPartyProject=false')
        Test-Path -LiteralPath $fixture.Files.Shared | Should Be $false
        Test-Path -LiteralPath $fixture.Files.Source | Should Be $true
        Test-Path -LiteralPath $fixture.Files.Sibling | Should Be $true
    }
}
