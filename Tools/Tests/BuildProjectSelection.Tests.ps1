Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperModule = Join-Path $repoRoot 'Tools\Modules\Build\BuildProjectSelection.psm1'
$solutionPath = Join-Path $repoRoot 'RedSalamander.sln'

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
            -ProjectName 'DxUiTests'

        $selection.BuildProjectDirectly | Should Be $true
        $selection.BuildInput | Should Match ([Regex]::Escape('Tests\DxUiTests\DxUiTests.vcxproj') + '$')
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
}
