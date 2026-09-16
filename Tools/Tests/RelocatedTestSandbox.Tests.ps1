Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/ProcessStreaming.psm1') -ErrorAction Stop

Describe 'Native test sandbox after package relocation' -Tag RequiresBuildToolchain {
    It 'uses an authorized explicit root without a checkout and still rejects invalid roots' {
        if ($env:RS_VALIDATION_PLATFORM -notin @('x64', 'ARM64') -or
            $env:RS_VALIDATION_CONFIGURATION -notin @('Debug', 'Release', 'ASan Debug')) {
            throw 'Select RS_VALIDATION_PLATFORM and RS_VALIDATION_CONFIGURATION for native fixture coverage.'
        }
        $fixture = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness relocated-test-sandbox -Case ([Guid]::NewGuid().ToString('N'))
        $source = @'
#include "TestSupport.h"
#include <fstream>

#if defined(RS_FIXTURE_ASAN) && !defined(__SANITIZE_ADDRESS__)
#error The ASan Debug fixture must be instrumented.
#endif

int wmain()
{
    namespace Support = RedSalamander::TestSupport;
    std::error_code ec;
    if (!Common::Testing::FindTestRepositoryRoot(ec).empty() || !ec) return 10;
    const auto expected = Support::GetEnvironmentString(Support::kTestRootEnvironmentVariable);
    if (expected.empty()) return 11;
    const auto actual = Support::ResolveTestSandboxBase(ec);
    if (ec || !Common::Testing::TestSandboxPathEquals(actual, expected)) return 12;
    const auto scratch = Support::AcquireTestDirectory(
        {.harnessSegment = L"relocated-native-proof", .leafSegment = L"writer",
         .fallbackRunIdPrefix = L"relocated-native-proof"}, ec);
    if (ec || scratch.empty() || !Common::Testing::IsSameOrDescendantTestSandboxPath(scratch, actual)) return 13;
    std::ofstream proof(scratch / L"proof.txt");
    proof << "relocated writer";
    proof.close();
    if (!proof) return 14;
    for (const auto invalid : {L"C:\\Windows:untrusted", L"C:\\Windows", L"Z:\\RedSalamander.Perf\\..\\outside"})
    {
        Support::ScopedEnvironmentVariable overrideRoot(Support::kTestRootEnvironmentVariable, invalid);
        if (!Support::ResolveTestSandboxBase(ec).empty() || !ec) return 15;
    }
    Support::ScopedEnvironmentVariable unsetRoot(Support::kTestRootEnvironmentVariable, std::nullopt);
    if (!Support::ResolveTestSandboxBase(ec).empty() || !ec) return 16;
    return 0;
}
'@
        [IO.File]::WriteAllText((Join-Path $fixture 'Relocated.cpp'), $source, [Text.UTF8Encoding]::new($false))
        $includes = [Security.SecurityElement]::Escape("$repoRoot\Tests\TestSupport;$repoRoot\Common")
        $runtimeTargets = [Security.SecurityElement]::Escape((Join-Path $repoRoot 'Directory.Build.targets'))
        $projectSource = @'
<Project DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <ItemGroup Label="ProjectConfigurations">
    <ProjectConfiguration Include="Debug|x64"><Configuration>Debug</Configuration><Platform>x64</Platform></ProjectConfiguration>
    <ProjectConfiguration Include="Release|x64"><Configuration>Release</Configuration><Platform>x64</Platform></ProjectConfiguration>
    <ProjectConfiguration Include="ASan Debug|x64"><Configuration>ASan Debug</Configuration><Platform>x64</Platform></ProjectConfiguration>
    <ProjectConfiguration Include="Debug|ARM64"><Configuration>Debug</Configuration><Platform>ARM64</Platform></ProjectConfiguration>
    <ProjectConfiguration Include="Release|ARM64"><Configuration>Release</Configuration><Platform>ARM64</Platform></ProjectConfiguration>
    <ProjectConfiguration Include="ASan Debug|ARM64"><Configuration>ASan Debug</Configuration><Platform>ARM64</Platform></ProjectConfiguration>
  </ItemGroup>
  <PropertyGroup Label="Globals"><WindowsTargetPlatformVersion>10.0</WindowsTargetPlatformVersion></PropertyGroup>
  <Import Project="$(VCTargetsPath)\Microsoft.Cpp.Default.props" />
  <PropertyGroup Label="Configuration">
    <ConfigurationType>Application</ConfigurationType><PlatformToolset>v145</PlatformToolset>
    <CharacterSet>Unicode</CharacterSet><UseDebugLibraries Condition="'$(Configuration)'!='Release'">true</UseDebugLibraries>
    <EnableASAN Condition="'$(Configuration)'=='ASan Debug'">true</EnableASAN>
  </PropertyGroup>
  <Import Project="$(VCTargetsPath)\Microsoft.Cpp.props" />
  <PropertyGroup><OutDir>$(MSBuildProjectDirectory)\bin\</OutDir><IntDir>$(MSBuildProjectDirectory)\obj\</IntDir><LinkIncremental>false</LinkIncremental></PropertyGroup>
  <ItemDefinitionGroup>
    <ClCompile>
      <AdditionalIncludeDirectories>__INCLUDES__</AdditionalIncludeDirectories><LanguageStandard>stdcpplatest</LanguageStandard>
      <PreprocessorDefinitions>UNICODE;_UNICODE;NOMINMAX</PreprocessorDefinitions><ExceptionHandling>Sync</ExceptionHandling>
      <WarningLevel>Level4</WarningLevel><TreatWarningAsError>true</TreatWarningAsError><SDLCheck>true</SDLCheck><ConformanceMode>true</ConformanceMode>
      <DebugInformationFormat>ProgramDatabase</DebugInformationFormat><BasicRuntimeChecks>Default</BasicRuntimeChecks>
    </ClCompile>
    <Link><SubSystem>Console</SubSystem></Link>
  </ItemDefinitionGroup>
  <ItemDefinitionGroup Condition="'$(Configuration)'=='ASan Debug'"><ClCompile><PreprocessorDefinitions>RS_FIXTURE_ASAN;%(PreprocessorDefinitions)</PreprocessorDefinitions></ClCompile></ItemDefinitionGroup>
  <ItemGroup><ClCompile Include="Relocated.cpp" /></ItemGroup>
  <Import Project="$(VCTargetsPath)\Microsoft.Cpp.targets" />
  <!-- Reuse the product's selected-toolset runtime staging, including native ARM64 ASAN. -->
  <Import Project="__RUNTIME_TARGETS__" />
</Project>
'@
        $project = Join-Path $fixture 'Relocated.vcxproj'
        [IO.File]::WriteAllText($project, $projectSource.Replace('__INCLUDES__', $includes).Replace('__RUNTIME_TARGETS__', $runtimeTargets), [Text.UTF8Encoding]::new($false))
        $arguments = @('-NoProfile', '-File', (Join-Path $repoRoot 'Tools/Invoke-SanitizedMsbuild.ps1'),
            $project, '/t:Build', "/p:Configuration=$env:RS_VALIDATION_CONFIGURATION", "/p:Platform=$env:RS_VALIDATION_PLATFORM",
            '/nologo', '/v:minimal', '/nr:false')
        $buildExit = Invoke-RSStreamingProcess -FilePath 'pwsh.exe' -Arguments $arguments -WorkingDirectory $fixture `
            -LogPath (Join-Path $fixture 'build.log') -TimeoutMilliseconds 120000 -DelegateArtifactOperation
        $buildExit | Should Be 0
        $runExit = Invoke-RSStreamingProcess -FilePath (Join-Path $fixture 'bin/Relocated.exe') -WorkingDirectory $fixture `
            -LogPath (Join-Path $fixture 'run.log') -TimeoutMilliseconds 30000
        $runExit | Should Be 0
    }
}
