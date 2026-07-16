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
            $source | Should Match 'ViewerFileComboHost::MessageMayOpenWindowComboPopup'
            $source | Should Match 'ViewerFileComboHost::ComputeStandaloneComboPopupHeightPx'
            $source | Should Not Match '(?m)^\s*\[\[nodiscard\]\]\s+bool\s+MessageMayOpenWindowComboPopup\s*\('
            $source | Should Not Match 'WM_KEYUP\s*&&\s*\(wp\s*==\s*VK_ESCAPE\s*\|\|\s*wp\s*==\s*VK_TAB\)'
            $source | Should Not Match 'PostMessageW\(root,\s*WM_CLOSE'
        }
    }

    It 'routes viewer Unicode text copies through the shared ownership helper' {
        $clipboardSources = @(
            'Plugins\ViewerWeb\ViewerWeb.cpp',
            'Plugins\ViewerText\ViewerText.Text.cpp',
            'Plugins\ViewerText\ViewerText.Hex.cpp'
        )

        foreach ($relativePath in $clipboardSources) {
            $sourcePath = Join-Path $repoRoot $relativePath
            $source = Get-Content -Path $sourcePath -Raw

            $source | Should Match 'UnicodeClipboard\.h'
            $source | Should Match 'Common::Clipboard::TrySetUnicodeText'
            $source | Should Not Match '(?m)^\s*bool\s+CopyUnicodeTextToClipboard\s*\('
            $source | Should Not Match 'SetClipboardData\(CF_UNICODETEXT'
        }
    }

    It 'routes first-party viewer title bars through the shared DWM attribute policy' {
        $titleBarSources = @(
            'Plugins\ViewerPE\ViewerPE.cpp',
            'Plugins\ViewerWeb\ViewerWeb.cpp',
            'Plugins\ViewerImgRaw\ViewerImgRaw.cpp',
            'Plugins\ViewerText\ViewerText.cpp',
            'Plugins\ViewerVLC\ViewerVLC.cpp',
            'Plugins\ViewerSqlite\ViewerSqlite.cpp'
        )

        foreach ($relativePath in $titleBarSources) {
            $sourcePath = Join-Path $repoRoot $relativePath
            $source = Get-Content -Path $sourcePath -Raw

            $source | Should Match 'ViewerTitleBarTheme\.h'
            $source | Should Match 'ViewerChrome::ApplyTitleBarTheme'
            $source | Should Not Match 'kDwmwa(Caption|Border|Text)Color'
        }
    }

    It 'does not persist OnCreate hwnd captures in viewer chrome callbacks' {
        $chromeCallbackSources = @(
            'Plugins\ViewerPE\ViewerPE.cpp',
            'Plugins\ViewerWeb\ViewerWeb.cpp',
            'Plugins\ViewerImgRaw\ViewerImgRaw.cpp',
            'Plugins\ViewerText\ViewerText.cpp',
            'Plugins\ViewerSpace\ViewerSpace.cpp'
        )

        foreach ($relativePath in $chromeCallbackSources) {
            $sourcePath = Join-Path $repoRoot $relativePath
            Test-Path $sourcePath | Should Be $true
            $source = Get-Content -Path $sourcePath -Raw

            $source | Should Not Match '(SetOnSelectionChanged|ConfigureFileComboKeyboard|SetRefreshMenuStateCallback|SetOnTabBoundary|SetOnEscape)\s*\([^`r`n]*\[[^\]]*\bhwnd\b'
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
    It 'uses one detached console-policy WinGet alias launcher for normal and foreground waits' {
        $launcherProjectPath = Join-Path $repoRoot 'RedLauncher\RedLauncher.vcxproj'
        $launcherManifestPath = Join-Path $repoRoot 'RedLauncher\res\exe.manifest'
        $solutionPath = Join-Path $repoRoot 'RedSalamander.sln'
        $zipScriptPath = Join-Path $repoRoot 'Installer\zip\build-zip.ps1'

        Test-Path $launcherProjectPath | Should Be $true
        Test-Path (Join-Path $repoRoot 'RedLauncher\RedLauncherConsole.vcxproj') | Should Be $false
        Test-Path $launcherManifestPath | Should Be $true

        $launcherProject = Get-Content -Path $launcherProjectPath -Raw
        $launcherManifest = Get-Content -Path $launcherManifestPath -Raw
        $solution = Get-Content -Path $solutionPath -Raw
        $zipScript = Get-Content -Path $zipScriptPath -Raw

        $launcherProject | Should Match '<SubSystem>Console</SubSystem>'
        $launcherProject | Should Match '<AdditionalManifestFiles>\$\(ProjectDir\)res\\exe\.manifest</AdditionalManifestFiles>'
        $launcherManifest | Should Match 'consoleAllocationPolicy'
        $launcherManifest | Should Match '>detached<'
        $solution | Should Not Match 'RedLauncherConsole'
        $zipScript | Should Not Match 'RedLauncherConsole\.exe'
    }
}
