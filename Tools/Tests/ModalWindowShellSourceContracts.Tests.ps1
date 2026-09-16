Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Describe 'Modal window shell contracts' {
    BeforeAll {
        $appSource = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\RedSalamander.cpp') -Raw
        $shellSource = Get-Content -LiteralPath (Join-Path $repoRoot 'Common\ModalWindowShell.h') -Raw

        $aboutStart = $appSource.IndexOf('class AboutDialogWindow final', [StringComparison]::Ordinal)
        $aboutEnd = $appSource.IndexOf('class FatalErrorDialogWindow final', $aboutStart, [StringComparison]::Ordinal)
        $fatalStart = $aboutEnd
        $fatalEnd = $appSource.IndexOf('void ShowFatalErrorDialog(', $fatalStart, [StringComparison]::Ordinal)

        $aboutSource = $appSource.Substring($aboutStart, $aboutEnd - $aboutStart)
        $fatalSource = $appSource.Substring($fatalStart, $fatalEnd - $fatalStart)
    }

    It 'routes About and Fatal Error through the shared modal owner and message-loop shell' {
        foreach ($dialogSource in @($aboutSource, $fatalSource)) {
            $dialogSource | Should Match 'Common::ModalWindowShell\s+modalShell'
            $dialogSource | Should Match 'modalShell\.CreateCentered\(createOptions, hwnd\)'
            $dialogSource | Should Match 'modalShell\.ShowAndRun\(_hWnd\.get\(\), _done, _result,'
            $dialogSource | Should Not Match 'EnableWindow\(_ownerWindow'
            $dialogSource | Should Not Match 'GetMessageW\('
            $dialogSource | Should Not Match 'CreateWindowExW\('
            $dialogSource | Should Not Match 'DestroyWindow\(_hWnd\.get\(\)\)'
        }
    }

    It 'keeps quit propagation and owner restoration inside the narrow shell' {
        $shellSource | Should Match 'DxUi::RunDxUiModalLoop\(hwnd, loopOptions\)'
        $shellSource | Should Not Match 'GetMessageW\(|PostQuitMessage\('
        $shellSource | Should Match 'EnableWindow\(_ownerWindow, FALSE\)'
        $shellSource | Should Match 'EnableWindow\(_ownerWindow, TRUE\)'
        $shellSource | Should Match 'SetActiveWindow\(_ownerWindow\)'
        $shellSource | Should Not Match 'WM_QUERYENDSESSION|WM_ENDSESSION'
    }

    It 'does not extend the modal helper to unrelated application paths' {
        ([regex]::Matches($appSource, 'Common::ModalWindowShell\s+modalShell')).Count | Should Be 2
    }
}
