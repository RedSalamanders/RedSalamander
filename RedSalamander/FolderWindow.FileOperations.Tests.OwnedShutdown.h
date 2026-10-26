#pragma once
#ifdef ENABLE_TESTS
#include "Blake3Digest.h"
#include "PlugInterfaces/FileSystem.h"
#include <cstdint>
#include <filesystem>
#include <string>
#include <string_view>
#include <windows.h>

// EXE-private test authority and an owned event gate. This is never plugin ABI.
class FolderWindow;
namespace OwnedShutdownSelfTest
{
inline constexpr std::wstring_view kCase   = L"FileOps_OwnedProductionShutdown";
inline constexpr const wchar_t* kRole      = L"REDSALAMANDER_TEST_D10_CHILD";
inline constexpr const wchar_t* kCaseRoot  = L"REDSALAMANDER_TEST_D10_CASE_ROOT";
inline constexpr const wchar_t* kNonce     = L"REDSALAMANDER_TEST_D10_NONCE";
inline constexpr const wchar_t* kParentPid = L"REDSALAMANDER_TEST_D10_PARENT_PID";
struct StageProof final
{
    bool entered   = false;
    bool released  = false;
    bool returned  = false;
    uint64_t bytes = 0u;
    HRESULT status = E_PENDING;
    std::wstring source, destination, stage;
    Common::Crypto::Blake3Digest digest{};
};
HRESULT ArmAuthority(HWND mainWindow,
                     std::wstring_view currentCase,
                     const std::filesystem::path& caseRoot,
                     std::wstring_view runId,
                     std::wstring_view nonce,
                     DWORD parentPid) noexcept;
bool UsesProductionRoute(HWND mainWindow) noexcept;
bool OwnerHasActiveOperations(FolderWindow& owner) noexcept;
HRESULT ArmStagePause(const std::filesystem::path& source, const std::filesystem::path& destination) noexcept;
HRESULT MaybePauseStageAfterWrite(std::wstring_view source,
                                  std::wstring_view destination,
                                  std::wstring_view stage,
                                  IFileSystemBoundObject* ownedStage,
                                  const FileSystemOptions* options,
                                  uint64_t bytes) noexcept;
StageProof GetStageProof();
void ReleaseStagePause() noexcept;
void RecordParkedProof(bool originalsPreserved, uint32_t heartbeatTicks, uint64_t heartbeatMs, bool stopping, bool foregroundPreserved) noexcept;
void RecordCompletion(HRESULT status, bool exactUnpublishedAndStageRemoved) noexcept;
void RecordFailure(std::wstring_view reason) noexcept;
void RecordCloseAttempt(bool active) noexcept;
void RecordDeferredClose(bool active) noexcept;
void RecordAcceptedQuietClose(bool quiet) noexcept;
void RecordNormalSettingsSave(bool ownedSave, HRESULT status) noexcept;
bool RecordNormalDestroy(HWND mainWindow, uint64_t normalTeardownCount) noexcept;
} // namespace OwnedShutdownSelfTest
#endif
