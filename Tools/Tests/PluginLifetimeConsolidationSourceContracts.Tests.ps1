Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Describe 'Plugin lifetime consolidation source contracts' {
    BeforeAll {
        $helpersSource = Get-Content -LiteralPath (Join-Path $repoRoot 'Common\Helpers.h') -Raw
        $googleHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'Plugins\FileSystemGoogleDrive\FileSystemGoogleDrive.h') -Raw
        $googleSource = Get-Content -LiteralPath (Join-Path $repoRoot 'Plugins\FileSystemGoogleDrive\FileSystemGoogleDrive.cpp') -Raw
        $microsoftHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.h') -Raw
        $microsoftSource = Get-Content -LiteralPath (Join-Path $repoRoot 'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.cpp') -Raw
        $managerHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\FileSystemPluginManager.h') -Raw
        $lookupConsumers = @(
            'RedSalamander\CompareDirectoriesWindow.cpp'
            'RedSalamander\FindFilesWindow.cpp'
            'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'
            'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'
        ) | ForEach-Object { Get-Content -LiteralPath (Join-Path $repoRoot $_) -Raw }
        $threadpoolSources = Get-ChildItem -LiteralPath (Join-Path $repoRoot 'Plugins'), (Join-Path $repoRoot 'RedSalamander') -Recurse -Include *.cpp,*.h |
            ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }
    }

    It 'uses the shared generation and drain state for the two equivalent navigation callback families' {
        foreach ($header in @($googleHeader, $microsoftHeader)) {
            $header | Should Match 'RegistrationCallbackState<INavigationMenuCallback>'
            $header | Should Match 'NavigationMenuCallbackState\s+_navigationMenuCallbackState'
            $header | Should Not Match '_navigationMenuCallbacksInFlight'
            $header | Should Not Match '_navigationMenuDrainCv'
        }

        foreach ($source in @($googleSource, $microsoftSource)) {
            $source | Should Match '_navigationMenuCallbackState\.Set\(callback, cookie\)'
            $source | Should Match '_navigationMenuCallbackState\.TryCapture\(snapshot\)'
            $source | Should Match '_navigationMenuCallbackState\.TryEnter\(snapshot, callback, cookie\)'
            $source | Should Match '_navigationMenuCallbackState\.FinishInvoke\(\)'
        }
    }

    It 'centralizes callback-return module-pin transfer' {
        $helpersSource | Should Match 'void\s+TransferModulePinToCallbackReturn\(PTP_CALLBACK_INSTANCE instance, wil::unique_hmodule& modulePin\)\s+noexcept'
        ([regex]::Matches($helpersSource, 'FreeLibraryWhenCallbackReturns\(')).Count | Should Be 1
        foreach ($source in $threadpoolSources) {
            $source | Should Not Match 'FreeLibraryWhenCallbackReturns\('
        }
    }

    It 'exposes one case-insensitive manager lookup with an explicit pointer lease' {
        $managerHeader | Should Match 'returned non-owning pointer remains valid only until'
        $managerHeader | Should Match '\[\[nodiscard\]\]\s+const PluginEntry\* FindPluginById\(std::wstring_view pluginId\) const noexcept;'
        foreach ($source in $lookupConsumers) {
            $source | Should Match 'FileSystemPluginManager::GetInstance\(\)\.FindPluginById\(pluginId\)'
            $source | Should Not Match 'CompareStringOrdinal\(entry\.id\.c_str\(\)'
        }
    }
}
