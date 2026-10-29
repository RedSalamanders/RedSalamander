$ciWarningRepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $ciWarningRepoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Scope Local
Import-Module (Join-Path $ciWarningRepoRoot 'Build/BuildConsole.psm1') -Scope Local
Import-Module (Join-Path $ciWarningRepoRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -Scope Local

Describe 'CI dependency header warning classification' {
    It 'classifies the injected include path without suppressing first-party warnings' {
        if (-not $IsWindows) { throw 'This native warning-policy proof requires Windows and MSVC.' }
        $vs = Get-RSVcpkgVisualStudioInstallation
        $msbuild = Join-Path $vs 'MSBuild/Current/Bin/amd64/MSBuild.exe'
        $platform = 'x64'
        if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64') {
            $msbuild = Join-Path $vs 'MSBuild/Current/Bin/arm64/MSBuild.exe'
            $platform = 'ARM64'
        }
        $version = (Get-RSVcpkgDefaultToolset -Installation $vs).Version
        $version | Should Match '^14[.]52[.]\d+$'
        $fixture = New-RSTestSandboxScratchDirectory -RepoRoot $ciWarningRepoRoot `
            -Harness 'ci-external-warning-policy' -Case ([guid]::NewGuid().ToString('N'))
        $installed = Join-Path $fixture 'installed'
        $include = Join-Path $installed 'fixture-triplet/include'
        [void][IO.Directory]::CreateDirectory($include)
        $header = Join-Path $include 'CiExternalWarningProbe.h'
        [IO.File]::WriteAllText($header, "#pragma once`n#if CI_WARNING_PROBE_UNDEFINED_EXTERNAL`n#endif`ninline int ExternalProbe() { return 0; }`n")
        $source = Join-Path $fixture 'Probe.cpp'
        $plainSource = "#include <CiExternalWarningProbe.h>`nint Probe() { return ExternalProbe(); }`n"
        $emptyImports = @{}
        foreach ($role in 'DirectoryProps', 'DirectoryTargets', 'BeforeCppTargets', 'AfterCppTargets') {
            $emptyImports[$role] = Join-Path $fixture ($role + '.props')
            [IO.File]::WriteAllText($emptyImports[$role], '<Project/>')
        }
        $project = Join-Path $fixture 'CiExternalWarningProbe.vcxproj'
        $xml = @'
<Project DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <ItemGroup Label="ProjectConfigurations"><ProjectConfiguration Include="Debug|__PLATFORM__"><Configuration>Debug</Configuration><Platform>__PLATFORM__</Platform></ProjectConfiguration></ItemGroup>
  <PropertyGroup Label="Globals"><WindowsTargetPlatformVersion>10.0.26100.0</WindowsTargetPlatformVersion></PropertyGroup>
  <Import Project="$(VCTargetsPath)\Microsoft.Cpp.Default.props"/>
  <PropertyGroup Label="Configuration"><ConfigurationType>StaticLibrary</ConfigurationType><PlatformToolset>v145</PlatformToolset><UseDebugLibraries>true</UseDebugLibraries></PropertyGroup>
  <Import Project="$(VCTargetsPath)\Microsoft.Cpp.props"/>
  <PropertyGroup><OutDir>__OUTPUT__\</OutDir><IntDir>__OUTPUT__\obj\</IntDir></PropertyGroup>
  <ItemDefinitionGroup><ClCompile><PrecompiledHeader>NotUsing</PrecompiledHeader><LanguageStandard>stdcpplatest</LanguageStandard></ClCompile></ItemDefinitionGroup>
  <ItemGroup><ClCompile Include="Probe.cpp"/></ItemGroup>
  <Import Project="$(VCTargetsPath)\Microsoft.Cpp.targets"/>
  <Import Project="__POLICY__"/>
</Project>
'@
        $xml = $xml.Replace('__PLATFORM__', $platform).Replace('__OUTPUT__', [Security.SecurityElement]::Escape((Join-Path $fixture 'output'))).Replace('__POLICY__', [Security.SecurityElement]::Escape((Join-Path $ciWarningRepoRoot 'Build/PinnedDxUiPolicy.targets')))
        [IO.File]::WriteAllText($project, $xml)
        $names = @('CL', 'ExternalIncludePath', 'VCPKG_INSTALLED_DIR', 'VCPKG_DEFAULT_TRIPLET', '_MSPDBSRV_ENDPOINT_')
        $saved = @{}
        foreach ($name in $names) { $saved[$name] = [Environment]::GetEnvironmentVariable($name, 'Process') }
        try {
            $env:CL = '/I"' + $include + '"'
            [Environment]::SetEnvironmentVariable('ExternalIncludePath', $null, 'Process')
            $env:VCPKG_DEFAULT_TRIPLET = 'fixture-triplet'
            $env:_MSPDBSRV_ENDPOINT_ = 'RSCiWarningProbe' + [guid]::NewGuid().ToString('N')
            $defaultExternalPaths = @()
            foreach ($phase in 'before', 'external', 'first-party') {
                [IO.File]::WriteAllText($source, $(if ($phase -eq 'first-party') {
                    $plainSource + "#if CI_WARNING_PROBE_UNDEFINED_FIRST_PARTY`n#endif`n"
                } else { $plainSource }))
                [Environment]::SetEnvironmentVariable('VCPKG_INSTALLED_DIR', $(if ($phase -eq 'before') { $null } else { $installed }), 'Process')
                $log = Join-Path $fixture ($phase + '.log')
                $profileArguments = @($project, '/nologo', '/p:Configuration=Debug', "/p:Platform=$platform",
                    "/p:VCToolsVersion=$version", '/p:VcpkgEnabled=false', "/p:DirectoryBuildPropsPath=$($emptyImports.DirectoryProps)", "/p:DirectoryBuildTargetsPath=$($emptyImports.DirectoryTargets)",
                    "/p:ForceImportBeforeCppTargets=$($emptyImports.BeforeCppTargets)", "/p:ForceImportAfterCppTargets=$($emptyImports.AfterCppTargets)")
                $arguments = $profileArguments + @('/v:minimal', '/t:Rebuild')
                & $msbuild @arguments *> $log
                $LASTEXITCODE | Should Be 0
                $diagnostics = Get-NativeBuildDiagnostics -LogPath $log
                $diagnostics.ErrorCount | Should Be 0
                $text = Get-Content -LiteralPath $log -Raw
                if ($phase -in 'before', 'external') {
                    $profileText = @(& $msbuild @profileArguments -getProperty:ExternalIncludePath)
                    $LASTEXITCODE | Should Be 0
                    $profile = $profileText -join "`n"
                    $paths = @($profile.Split(';') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                    if ($phase -eq 'before') { $defaultExternalPaths = $paths; $defaultExternalPaths.Count | Should BeGreaterThan 0 }
                }
                if ($phase -eq 'external') {
                    $diagnostics.WarningCount | Should Be 0
                    Assert-NativeBuildDiagnostics -LogPath $log
                    $profile | Should Match ([regex]::Escape($include))
                    foreach ($originalPath in $defaultExternalPaths) { ($paths -contains $originalPath) | Should Be $true }
                } else {
                    $diagnostics.WarningCount | Should Be 1
                    $text | Should Match 'warning C4668'
                    if ($phase -eq 'before') { $text | Should Match 'CiExternalWarningProbe[.]h' }
                    else { $text | Should Match 'Probe[.]cpp' }
                    $rejected = $false
                    try { Assert-NativeBuildDiagnostics -LogPath $log }
                    catch { $_.Exception.Message | Should Match 'Zero-warning build validation failed'; $rejected = $true }
                    $rejected | Should Be $true
                }
                Write-Host "CI include probe $phase : $($diagnostics.WarningCount) warning(s), zero-warning gate checked; log=$log"
            }
        } finally {
            foreach ($name in $names) { [Environment]::SetEnvironmentVariable($name, $saved[$name], 'Process') }
        }
    }
}
