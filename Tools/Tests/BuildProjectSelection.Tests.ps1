Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperModule = Join-Path $repoRoot 'Tools\Modules\Build\BuildProjectSelection.psm1'
$solutionPath = Join-Path $repoRoot 'RedSalamander.sln'
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -ErrorAction Stop

Describe 'Build project selection helper' {
    BeforeAll {
        Import-Module $helperModule -Force -ErrorAction Stop
    }

    It 'builds plugin projects directly from their vcxproj files' {
        $selection = Get-RSBuildSelection `
            -SolutionPath $solutionPath `
            -SolutionDir $repoRoot `
            -ProjectName 'ViewerText'

        $selection.BuildProjectDirectly | Should Be $true
        $selection.BuildInput | Should Match ([Regex]::Escape('Plugins\ViewerText\ViewerText.vcxproj') + '$')
        $selection.MSBuildTarget | Should Be 'Build'
        $selection.CleanTarget | Should Be 'Clean'
    }

    It 'builds test projects directly from their vcxproj files' {
        $selection = Get-RSBuildSelection `
            -SolutionPath $solutionPath `
            -SolutionDir $repoRoot `
            -ProjectName 'ProductUiTests'

        $selection.BuildProjectDirectly | Should Be $true
        $selection.BuildInput | Should Match ([Regex]::Escape('Tests\ProductUiTests\ProductUiTests.vcxproj') + '$')
        $selection.MSBuildTarget | Should Be 'Build'
        $selection.CleanTarget | Should Be 'Clean'
    }

    It 'keeps RedSalamander on the solution build graph' {
        $selection = Get-RSBuildSelection `
            -SolutionPath $solutionPath `
            -SolutionDir $repoRoot `
            -ProjectName 'RedSalamander'

        $selection.BuildProjectDirectly | Should Be $false
        $selection.BuildInput | Should Be $solutionPath
        $selection.MSBuildTarget | Should Be 'RedSalamander'
        $selection.CleanTarget | Should Be 'RedSalamander:Clean'
    }

    It 'uses the rebuild target directly for vcxproj rebuilds' {
        $selection = Get-RSBuildSelection `
            -SolutionPath $solutionPath `
            -SolutionDir $repoRoot `
            -ProjectName 'ViewerText' `
            -Rebuild

        $selection.BuildProjectDirectly | Should Be $true
        $selection.MSBuildTarget | Should Be 'Rebuild'
        $selection.CleanTarget | Should Be 'Clean'
    }

    It 'retains the full solution folder path for production localizations and plugins' {
        $projects = @(Get-RSSolutionProjects -SolutionPath $solutionPath)
        ($projects | Where-Object Name -eq 'RedSalamander-fr-FR').SolutionTarget | Should Be 'Solution Items\RedSalamander-fr-FR'
        ($projects | Where-Object Name -eq 'ViewerText').SolutionTarget | Should Be 'Plugins\ViewerText'
        ($projects | Where-Object Name -eq 'RedSalamander').SolutionTarget | Should Be 'RedSalamander'
    }

    It 'builds one canonical union through the solution graph and keeps partial receipt scope' {
        $selection = Get-RSBuildProjectUnion -SolutionPath $solutionPath -ProjectNames @('ViewerText', 'ProductUiTests', 'viewertext')
        $selection.BuildInput | Should Be $solutionPath
        $selection.ReceiptTarget | Should Be 'selected-projects'
        ($selection.ProjectSelection -join '|') | Should Be 'Plugins/ViewerText/ViewerText.vcxproj|Tests/ProductUiTests/ProductUiTests.vcxproj'
        $selection.MSBuildTarget | Should Be 'Plugins\ViewerText;Tests\ProductUiTests'
        $selection.CleanTarget | Should Be 'Plugins\ViewerText:Clean;Tests\ProductUiTests:Clean'
        $reordered = Get-RSBuildProjectUnion -SolutionPath $solutionPath -ProjectNames @('ProductUiTests', 'ViewerText') -Rebuild
        ($reordered.ProjectSelection -join '|') | Should Be ($selection.ProjectSelection -join '|')
        $reordered.MSBuildTarget | Should Be 'Plugins\ViewerText:Rebuild;Tests\ProductUiTests:Rebuild'
    }

    It 'declares runtime-loaded viewer and contract inputs in the actual native build graph' {
        $projects = @(Get-RSSolutionProjects -SolutionPath $solutionPath)
        $owners = @{
            ViewerPETests = @('Common', 'RedSalamander', 'ViewerPE', 'ViewerWeb', 'ViewerImgRaw', 'ViewerText', 'ViewerSpace', 'ViewerVLC')
            PluginContractTests = @('FileSystem', 'FileSystem7z', 'FileSystemCurl', 'FileSystemS3', 'FileSystemGoogleDrive', 'FileSystemMicrosoftDrive', 'FileSystemMtp', 'FileSystemDummy', 'Terminal', 'Terminal-ja-JP', 'ViewerPE', 'ViewerWeb', 'ViewerImgRaw', 'ViewerText', 'ViewerSpace', 'ViewerVLC', 'ViewerSqlite')
        }
        foreach ($owner in $owners.Keys) {
            $ownerProject = @($projects | Where-Object Name -ceq $owner)[0]
            [xml]$xml = Get-Content -LiteralPath $ownerProject.Path -Raw
            $references = @($xml.SelectNodes('//*[local-name()="ProjectReference"]'))
            foreach ($name in $owners[$owner]) {
                $producer = @($projects | Where-Object Name -ceq $name)[0]
                $matching = @($references | Where-Object {
                    [IO.Path]::GetFullPath((Join-Path (Split-Path $ownerProject.Path) $_.Include)) -ieq $producer.Path
                })
                $matching.Count | Should Be 1
                $matching[0].Project.Trim('{}') | Should Be $producer.Guid
                # These are runtime fixtures, never additional link inputs.
                $matching[0].LinkLibraryDependencies | Should Be 'false'
            }
        }
    }

    It 'refuses unknown, folder, empty and injected target names' {
        foreach ($invalid in @('DoesNotExist', 'Plugins', '', 'ViewerText;Build')) {
            $attempt = Invoke-RSTestBlock { Get-RSBuildProjectUnion -SolutionPath $solutionPath -ProjectNames @('ProductUiTests', $invalid) }
            $attempt.Succeeded | Should Be $false
        }
    }

    It 'rejects incompatible public union roles before tool discovery or module import' {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'build.ps1'), [ref]$tokens, [ref]$errors)
        @($errors).Count | Should Be 0
        $guard = @($ast.FindAll({param($node)
            $node -is [Management.Automation.Language.IfStatementAst] -and
            $node.Extent.Text -match 'ProjectNames cannot be combined'
        }, $true))
        $guard.Count | Should Be 1
        $firstImport = @($ast.FindAll({param($node)
            $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Import-Module'
        }, $true) | Sort-Object {$_.Extent.StartOffset} | Select-Object -First 1)[0]
        ($guard[0].Extent.EndOffset -lt $firstImport.Extent.StartOffset) | Should Be $true
        $ProjectNames = @('ViewerText', 'ProductUiTests')
        $ProjectName = ''; $ProductionOnly = $false; $Zip = $false; $GenerateWingetManifest = $false; $OfficialRelease = $false
        foreach ($conflictingParameter in @('ProjectName', 'ProductionOnly', 'Zip', 'GenerateWingetManifest', 'OfficialRelease')) {
            Set-Variable -Name $conflictingParameter -Value $(if ($conflictingParameter -eq 'ProjectName') {'Common'} else {$true})
            { . ([scriptblock]::Create($guard[0].Extent.Text)) } | Should Throw 'ProjectNames cannot be combined'
            Set-Variable -Name $conflictingParameter -Value $(if ($conflictingParameter -eq 'ProjectName') {''} else {$false})
        }
        . ([scriptblock]::Create($guard[0].Extent.Text))
    }

}

Describe 'Actual nested solution target selection' -Tag RequiresBuildToolchain {
    BeforeAll {
        Import-Module $helperModule -Force -ErrorAction Stop
    }
    It 'invokes nested and reserved-name solution targets through actual MSBuild without compiling C++' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -ErrorAction Stop
        $installation = Get-RSVcpkgVisualStudioInstallation
        $hostFolder = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -ceq 'Arm64') {'arm64'} else {'amd64'}
        $msbuild = Join-Path $installation "MSBuild/Current/Bin/$hostFolder/MSBuild.exe"
        $fixture = Join-Path $TestDrive 'solution-targets'
        [void][IO.Directory]::CreateDirectory($fixture)
        $projectText = @'
<Project DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <Target Name="Build"><WriteLinesToFile File="$(MSBuildProjectDirectory)/selected.txt" Lines="$(MSBuildProjectName)" Overwrite="false" /></Target>
  <Target Name="Rebuild" DependsOnTargets="Build" />
</Project>
'@
        [IO.File]::WriteAllText((Join-Path $fixture 'nested.vcxproj'), $projectText)
        [IO.File]::WriteAllText((Join-Path $fixture 'reserved.vcxproj'), $projectText)
        $text = @'
Microsoft Visual Studio Solution File, Format Version 12.00
Project("{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}") = "Plug.in", "nested.vcxproj", "{11111111-1111-1111-1111-111111111111}"
EndProject
Project("{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}") = "Build", "reserved.vcxproj", "{22222222-2222-2222-2222-222222222222}"
EndProject
Project("{2150E333-8FDC-42A3-9474-1A3956D46DE8}") = "Outer.Folder", "Outer.Folder", "{33333333-3333-3333-3333-333333333333}"
EndProject
Project("{2150E333-8FDC-42A3-9474-1A3956D46DE8}") = "Child", "Child", "{44444444-4444-4444-4444-444444444444}"
EndProject
Global
  GlobalSection(SolutionConfigurationPlatforms) = preSolution
    Release|x64 = Release|x64
  EndGlobalSection
  GlobalSection(ProjectConfigurationPlatforms) = postSolution
    {11111111-1111-1111-1111-111111111111}.Release|x64.ActiveCfg = Release|x64
    {11111111-1111-1111-1111-111111111111}.Release|x64.Build.0 = Release|x64
    {22222222-2222-2222-2222-222222222222}.Release|x64.ActiveCfg = Release|x64
    {22222222-2222-2222-2222-222222222222}.Release|x64.Build.0 = Release|x64
  EndGlobalSection
  GlobalSection(NestedProjects) = preSolution
    {11111111-1111-1111-1111-111111111111} = {44444444-4444-4444-4444-444444444444}
    {44444444-4444-4444-4444-444444444444} = {33333333-3333-3333-3333-333333333333}
  EndGlobalSection
EndGlobal
'@
        $sln = Join-Path $fixture 'fixture.sln'
        [IO.File]::WriteAllText($sln, $text)
        $projects = @(Get-RSSolutionProjects -SolutionPath $sln)
        $nested = ($projects | Where-Object Name -eq 'Plug.in').SolutionTarget
        $reserved = ($projects | Where-Object Name -eq 'Build').SolutionTarget
        $nested | Should Be 'Outer_Folder\Child\Plug_in'
        $reserved | Should Be 'Solution:Build'
        $union = Get-RSBuildProjectUnion -SolutionPath $sln -ProjectNames @('Build', 'Plug.in', 'build') -Rebuild
        $union.ProjectNames.Count | Should Be 2
        $output = @(& $msbuild $union.BuildInput ("/t:" + $union.MSBuildTarget) /p:Configuration=Release /p:Platform=x64 /nologo /v:quiet /nr:false 2>&1)
        $output | Set-Content -LiteralPath (Join-Path $fixture 'msbuild.log')
        $LASTEXITCODE | Should Be 0
        $selected = @(Get-Content -LiteralPath (Join-Path $fixture 'selected.txt'))
        ($selected -ccontains 'nested') | Should Be $true
        ($selected -ccontains 'reserved') | Should Be $true
        $selected.Count | Should Be 2
        [IO.File]::WriteAllText($sln, $text.Replace('{44444444-4444-4444-4444-444444444444} = {33333333-3333-3333-3333-333333333333}', '{44444444-4444-4444-4444-444444444444} = {44444444-4444-4444-4444-444444444444}'))
        { Get-RSSolutionProjects -SolutionPath $sln } | Should Throw 'Cyclic solution nesting'
        [IO.File]::WriteAllText($sln, $text.Replace('{44444444-4444-4444-4444-444444444444} = {33333333-3333-3333-3333-333333333333}', '{44444444-4444-4444-4444-444444444444} = {22222222-2222-2222-2222-222222222222}'))
        { Get-RSSolutionProjects -SolutionPath $sln } | Should Throw 'Unknown, non-folder or duplicate solution nesting'
    }
}
