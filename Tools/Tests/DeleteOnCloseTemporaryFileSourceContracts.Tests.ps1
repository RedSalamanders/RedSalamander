Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force

Describe 'Delete-on-close temporary-file source contracts' {
    BeforeAll {
        $helper = Get-RSText -Path 'Common\DeleteOnCloseTemporaryFile.h'
        $curlPolicy = Get-RSText -Path 'Plugins\FileSystemCurl\FileSystemCurl.Internal.h'
        $s3Policy = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.Internal.h'
        $microsoftDrive = Get-RSText -Path 'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.cpp'
        $fileActionHeader = Get-RSText -Path 'RedSalamander\FileActionLauncher.h'
        $fileActionSource = Get-RSText -Path 'RedSalamander\FileActionLauncher.cpp'
        $fileActionSelfTest = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'
        $searchService = Get-RSText -Path 'RedSalamanderSearchService\Main.cpp'

        $providerSources = @(
            Get-ChildItem -LiteralPath (Join-Path $repoRoot 'Plugins\FileSystemCurl') -Recurse -File -Include '*.h', '*.cpp'
            Get-ChildItem -LiteralPath (Join-Path $repoRoot 'Plugins\FileSystemS3') -Recurse -File -Include '*.h', '*.cpp'
            Get-ChildItem -LiteralPath (Join-Path $repoRoot 'Plugins\FileSystemMicrosoftDrive') -Recurse -File -Include '*.h', '*.cpp'
        ) | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }
        $providerText = $providerSources -join "`n"
    }

    It 'owns the reservation until a valid WIL delete-on-close handle takes over' {
        $helper | Should Match 'class TemporaryPathReservation final'
        $helper | Should Match 'GetTempFileNameW\('
        $helper | Should Match '~TemporaryPathReservation\(\) noexcept[\s\S]*DeleteFileW\(_path\.c_str\(\)\)'
        $helper | Should Match 'wil::unique_hfile file\(CreateFileW\('
        $helper | Should Match 'options\.flagsAndAttributes \| FILE_FLAG_DELETE_ON_CLOSE'
        $helper | Should Match 'reservation\.TransferToHandle\(\)[\s\S]*outFile = std::move\(file\)'
        $helper | Should Match 'FailNextDeleteOnCloseTemporaryFileOpen'
    }

    It 'centralizes provider allocation while retaining provider-specific prefixes and flags' {
        $providerText | Should Not Match 'GetTempFileNameW\('
        $providerText | Should Match 'CreateDeleteOnCloseTemporaryFile\('
        $curlPolicy | Should Match 'kCurlTemporaryFileOptions[\s\S]*\.prefix\s*=\s*L"rsc"[\s\S]*FILE_ATTRIBUTE_TEMPORARY \| FILE_FLAG_SEQUENTIAL_SCAN'
        $s3Policy | Should Match 'kS3TemporaryFileOptions[\s\S]*\.prefix\s*=\s*L"rs3"'
        $microsoftDrive | Should Match 'DeleteOnCloseTemporaryFileOptions tempOptions[\s\S]*\.prefix\s*=\s*L"rsm"[\s\S]*CreateDeleteOnCloseTemporaryFile\(tempOptions'
    }

    It 'keeps selected-path files under a distinct identity-bearing process-lifetime lease with bounded launch cleanup' {
        $fileActionHeader | Should Match 'class SelectedPathsFileLease'
        $fileActionHeader | Should Match 'FILE_ID_INFO _fileIdentity'
        $fileActionHeader | Should Match 'FILE_ID_INFO _rootIdentity'
        $fileActionHeader | Should Match 'LaunchExternalPlan\(LaunchPlan plan'
        $fileActionHeader | Should Match 'selectedPathsMaximumRetentionMs\s*=\s*10u \* 60u \* 1000u'
        $fileActionSource | Should Match 'FILE_ATTRIBUTE_DIRECTORY \| FILE_ATTRIBUTE_REPARSE_POINT'
        $fileActionSource | Should Match 'constexpr DWORD kMaximumRetentionMs = 10u \* 60u \* 1000u'
        $fileActionSource | Should Not Match 'cleanupFilesAfterExit'
    }

    It 'cleans only the recorded file identity through one no-follow disposition handle' {
        $fileActionSource | Should Match 'QueryFileIdentity\(file\.get\(\), fileIdentity\)'
        $fileActionSource | Should Match 'file\.reset\(\);[\s\S]*rootHandle\.reset\(\);[\s\S]*outLease = std::move\(lease\)'
        $fileActionSource | Should Match 'DELETE \| FILE_READ_ATTRIBUTES[\s\S]*FILE_SHARE_READ \| FILE_SHARE_WRITE[\s\S]*FILE_FLAG_OPEN_REPARSE_POINT'
        $fileActionSource | Should Match 'EqualFileIdentity\(currentFileIdentity, expectedFileIdentity\)'
        $fileActionSource | Should Match 'SetFileInformationByHandle\(file, FileDispositionInfo'
        $fileActionSource | Should Not Match 'std::filesystem::remove\(file'
    }

    It 'recovers only bounded private-root manifests off the launch hot path' {
        $fileActionSource | Should Match 'TrySubmitThreadpoolCallback\(SelectedPathsRecoveryCallback'
        $fileActionSource | Should Match 'constexpr size_t kMaximumPasses = 32u'
        $fileActionSource | Should Match 'SelectedPathsRecoveryLimits kLimits\{24ull \* 60ull \* 60ull \* 1000ull, 128u, 16u, 20u\}'
        $fileActionSource | Should Match 'stoppedByEntryBound[\s\S]*stoppedByDeleteBound[\s\S]*stoppedByTimeBound'
        $fileActionSource | Should Match 'fileaction\.selected_paths_recovery\.inspected'
        $fileActionSource | Should Match 'fileaction\.selected_paths_recovery\.deleted'
        $fileActionSource | Should Match 'fileaction\.selected_paths_recovery\.bound_hits'
        $fileActionSource | Should Not Match 'IsOwnedSelectedPathsFileName'
        $fileActionSource | Should Not Match 'ScavengeSelectedPathsFiles'
        $fileActionSelfTest | Should Match 'sharedTempShape[\s\S]*rsa0001\.tmp[\s\S]*shared-temp shape'
    }

    It 'creates selected-path manifests exclusively in the private app-data root' {
        $fileActionSource | Should Match '#include "AppDataPaths\.h"'
        $fileActionSource | Should Match 'AppDataPaths::GetLocalAppDataPath\(\)'
        $fileActionSource | Should Match 'Common::Paths::CreateUniqueFileInDirectory\('
        $fileActionSource | Should Match 'CREATE_NEW'
        $fileActionSource | Should Not Match 'GetTempFileNameW\('
        $fileActionSource | Should Not Match 'GetTempPathW\('
    }

    It 'prevalidates exact records and streams them without an aggregate payload' {
        $fileActionSource | Should Match 'std::span<const std::filesystem::path>'
        $fileActionSource | Should Match 'ValidateSelectedPathRecord'
        $fileActionSource | Should Match 'Common::HandleIo::WriteAll'
        $fileActionSource | Should Match 'kMaximumSelectedPathsWriteBufferBytes\s*=\s*64u \* 1024u'
        $fileActionSource | Should Match 'std::vector<std::byte> writeBuffer'
        $fileActionSource | Should Not Match 'std::vector<std::filesystem::path> selectedPaths = effectiveContext\.selectedPaths'
        $fileActionSource | Should Not Match 'std::wstring text;'
    }

    It 'reuses canonical ordinary Windows argument quoting' {
        $fileActionSource | Should Match '#include "ProcessCommandLine\.h"'
        $fileActionSource | Should Match 'Common::Process::QuoteWindowsCommandLineArgument\('
        $fileActionSource | Should Not Match 'void AppendWindowsQuotedArgument\('
        $fileActionSource | Should Match 'AppendWindowsQuotedArgumentContent[\s\S]*different semantics'
    }

    It 'uses the real child parser to prove different-identity replacement survival' {
        $searchService | Should Match 'replaceManifest[\s\S]*DeleteFileW\(argv\.get\(\)\[2\]\)[\s\S]*CREATE_NEW'
        $fileActionSelfTest | Should Match 'file_action_selected_paths_identity_cleanup'
        $fileActionSelfTest | Should Match 'file_action_selected_paths_recovery_bounds'
        $fileActionSelfTest | Should Match 'child replacement sentinel'
    }
}
