Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Describe 'Viewer chrome keyboard contracts' {
    BeforeAll {
        $viewerSources = @(
            'Plugins\ViewerPE\ViewerPE.cpp',
            'Plugins\ViewerWeb\ViewerWeb.cpp',
            'Plugins\ViewerImgRaw\ViewerImgRaw.cpp',
            'Plugins\ViewerText\ViewerText.cpp'
        )
    }

    It 'routes standalone file-combo keyboard handling through the shared helper' {
        foreach ($relativePath in $viewerSources) {
            $sourcePath = Join-Path $repoRoot $relativePath
            Test-Path $sourcePath | Should Be $true
            $source = Get-Content -Path $sourcePath -Raw

            $source | Should Match 'ViewerFileComboHost\.h'
            $source | Should Match 'DispatchFileComboHostWndProc'
            $source | Should Match 'ConfigureFileComboKeyboard'
            $source | Should Not Match 'WM_KEYUP\s*&&\s*\(wp\s*==\s*VK_ESCAPE\s*\|\|\s*wp\s*==\s*VK_TAB\)'
            $source | Should Not Match 'PostMessageW\(root,\s*WM_CLOSE'
        }
    }

    It 'documents the shared Escape focus-cancel-close contract' {
        $pluginSpecPath = Join-Path $repoRoot 'Specs\Plugins\Plugins_ViewerPlugins.md'
        $spaceSpecPath = Join-Path $repoRoot 'Specs\Plugins\Plugins_ViewerSpace.md'
        $rawSpecPath = Join-Path $repoRoot 'Specs\Plugins\Plugins_ViewerImgRaw.md'

        $pluginSpec = Get-Content -Path $pluginSpecPath -Raw
        $spaceSpec = Get-Content -Path $spaceSpecPath -Raw
        $rawSpec = Get-Content -Path $rawSpecPath -Raw

        $pluginSpec | Should Match 'focus is inside viewer chrome'
        $pluginSpec | Should Match 'returns focus to the main viewer surface'
        $pluginSpec | Should Match 'cancellable work is active'
        $pluginSpec | Should Match 'posts `WM_CLOSE`'
        $spaceSpec | Should Match '`Esc`: cancel scan if scanning; otherwise close viewer'
        $rawSpec | Should Match '`Esc`: cancel active RAW decoding if loading; otherwise close viewer'
    }
}

Describe 'Launcher subsystem contracts' {
    It 'uses a Windows-subsystem WinGet alias launcher and a console companion for foreground waits' {
        $guiProjectPath = Join-Path $repoRoot 'RedLauncher\RedLauncher.vcxproj'
        $consoleProjectPath = Join-Path $repoRoot 'RedLauncher\RedLauncherConsole.vcxproj'
        $solutionPath = Join-Path $repoRoot 'RedSalamander.sln'
        $zipScriptPath = Join-Path $repoRoot 'Installer\zip\build-zip.ps1'

        Test-Path $guiProjectPath | Should Be $true
        Test-Path $consoleProjectPath | Should Be $true

        $guiProject = Get-Content -Path $guiProjectPath -Raw
        $consoleProject = Get-Content -Path $consoleProjectPath -Raw
        $solution = Get-Content -Path $solutionPath -Raw
        $zipScript = Get-Content -Path $zipScriptPath -Raw

        $guiProject | Should Match '<SubSystem>Windows</SubSystem>'
        $guiProject | Should Match '<PreprocessorDefinitions>REDLAUNCHER_GUI=1;'
        $consoleProject | Should Match '<SubSystem>Console</SubSystem>'
        $consoleProject | Should Match '<TargetName>RedLauncherConsole</TargetName>'
        $consoleProject | Should Match '<PreprocessorDefinitions>REDLAUNCHER_CONSOLE=1;'
        $solution | Should Match 'RedLauncherConsole'
        $zipScript | Should Match 'RedLauncherConsole\.exe'
    }
}
