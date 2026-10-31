#include "FolderWindow.FileOperations.Tests.Runner.h"

// FileOperations self-test â€” tick-driven async state machine.
//
// Architecture
// ------------
// The self-test runs as a cooperative state machine driven by the UI thread:
//   1. The host creates a timer and calls Tick(hwnd) on each tick.
//   2. Tick() advances the current step, starts async file-ops tasks, and
//      polls for completion via NotifyTaskCompleted() callbacks.
//   3. When Tick() returns true the run is complete (IsDone() == true).
//
// Active phase order
// ------------------
// kFileOpsPhaseOrder controls which Step enum values are exercised and in which
// order.  Adding a new step to the enum alone does not run it â€” it must also be
// appended to kFileOpsPhaseOrder.

#ifdef ENABLE_TESTS

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL move-only wrappers trigger deleted special member warnings in this TU

#include <algorithm>
#include <array>
#include <atomic>
#include <charconv>
#include <chrono>
#include <cstring>
#include <cwctype>
#include <filesystem>
#include <format>
#include <fstream>
#include <iterator>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>

#include <AccCtrl.h>
#include <AclAPI.h>
#include <TlHelp32.h>
#include <UIAutomation.h>
#include <cfapi.h>
#include <winhttp.h>
#include <winioctl.h>

#pragma comment(lib, "CldApi.lib") // Test-only, fully hydrated placeholder fixture; no cloud account or callback provider.

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unreferenced inline helpers
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

#include "BatchRenameExecutionEngine.h"
#include "BatchRenameWindow.h"
#include "Blake3Digest.h"
#include "Common.Tests.SandboxPath.h"
#include "Common.Tests.WindowActivationGuard.h"
#include "CompareDirectoriesWindow.h"
#include "ConnectionProfileUtils.h"
#include "ContentDigest.h"
#include "DirectoryInfoCache.h"
#include "FileOperationDurableStore.h"
#include "FileOperationItemResultReducer.h"
#include "FileOperationMoveBreadcrumb.h"
#include "FileOperationsConflictBroker.h"
#include "FileSystemPluginManager.h"
#include "FileSystemRouteContract.h"
#include "FileSystemRouteProviderBase.h"
#include "FluentIcons.h"
#include "FolderView.h"
#include "FolderWindow.FileOperations.IssuesPane.h"
#include "FolderWindow.FileOperations.IssuesProjection.h"
#include "FolderWindow.FileOperations.Popup.h"
#include "FolderWindow.FileOperations.PromptText.h"
#include "FolderWindow.FileOperations.State.Private.h"
#include "FolderWindow.FileOperations.Tests.OwnedShutdown.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "FolderWindow.h"
#include "HostServices.h"
#include "LocalFileTransaction.h"
#include "LocalizationManager.h"
#include "NavigationLocation.h"
#include "RedSalamander.h"
#include "SplashScreen.h"
#include "StringConversion.h"
#include "TestSupport/TestSupport.Tests.ChildProcess.h"
#include "TestSupport/TestSupport.Tests.DesktopObservation.h"
#include "TestSupport/TestSupport.Tests.DirectedSelfTestInputWarning.h"
#include "TestSupport/TestSupport.Tests.OwnedVirtualDisk.h"
#include "TestSupport/TestSupport.Tests.Runner.h"
#include "TestSupport/TestSupport.Tests.WindowScreenshot.h"
#include "Ui/AlertOverlayWindow.h"
#include "WindowMessages.h"
#include "YyjsonHelpers.h"
#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514) // Common/Helpers.h uses WIL types and triggers /Wall noise in this TU
#include "Helpers.h"
#include "PathUtils.h"
#pragma warning(pop)

static Common::Settings::Settings& g_settings = GetApplicationSettingsForSelfTest();

namespace
{
constexpr std::wstring_view kFolderWindowClassName    = L"RedSalamander.FolderWindow";
constexpr std::wstring_view kFolderViewClassName      = L"RedSalamanderFolderView";
constexpr std::wstring_view kPopupClassName           = L"RedSalamander.FileOperationsPopup";
constexpr std::wstring_view kPluginIdLocal            = L"builtin/file-system";
constexpr std::wstring_view kPluginIdDummy            = L"builtin/file-system-dummy";
constexpr std::wstring_view kPluginId7z               = L"builtin/file-system-7z";
constexpr std::wstring_view kPluginIdFtp              = L"builtin/file-system-ftp";
constexpr std::wstring_view kPluginIdSftp             = L"builtin/file-system-sftp";
constexpr std::wstring_view kPluginIdScp              = L"builtin/file-system-scp";
constexpr std::wstring_view kPluginIdImap             = L"builtin/file-system-imap";
constexpr std::wstring_view kPluginIdGoogleDrive      = L"builtin/file-system-gdrive";
constexpr std::wstring_view kPluginIdS3               = L"builtin/file-system-s3";
constexpr std::wstring_view kPluginIdS3Table          = L"builtin/file-system-s3table";
constexpr std::wstring_view kPluginIdOneDrivePersonal = L"builtin/file-system-onedrive-personal";
constexpr std::wstring_view kPluginIdOneDriveBusiness = L"builtin/file-system-onedrive-business";
constexpr std::wstring_view kPluginIdSharePoint       = L"builtin/file-system-sharepoint";
constexpr std::wstring_view kAlternateTestRootEnvVar  = L"REDSALAMANDER_TEST_ALTERNATE_ROOT";

// Synthetic provider for route-containment witnesses. By default it declares an uncontained route
// and never returns from a provider operation (the R0e rejection witness). With
// `boundedBlockedRead` it declares a bounded route and every mutation call blocks inside a
// synchronous Win32 read on an anonymous pipe with no writer, the same pending-IRP state a worker
// is left in by a dead SMB share (the R0f-SMB containment witness).
class UncontainedNeverReturningFileSystem final : public IFileSystem, public FileSystemRouteCapabilitiesBase
{
public:
    explicit UncontainedNeverReturningFileSystem(std::string capabilitiesJson,
                                                 bool exposeTypedRoute   = true,
                                                 bool typedCopyOperation = true,
                                                 bool boundedBlockedRead = false) noexcept
        : _capabilitiesJson(std::move(capabilitiesJson)),
          _exposeTypedRoute(exposeTypedRoute),
          _typedCopyOperation(typedCopyOperation),
          _boundedBlockedRead(boundedBlockedRead)
    {
        if (_boundedBlockedRead)
        {
            HANDLE readEnd  = nullptr;
            HANDLE writeEnd = nullptr;
            if (CreatePipe(&readEnd, &writeEnd, nullptr, 0u) != FALSE)
            {
                _pipeRead.reset(readEnd);
                _pipeWrite.reset(writeEnd);
            }
        }
    }

    [[nodiscard]] bool BlockedReadReady() const noexcept
    {
        return ! _boundedBlockedRead || (_pipeRead && _pipeWrite);
    }

    [[nodiscard]] bool WaitForBlockedReadReturn(DWORD timeoutMs) const noexcept
    {
        const ULONGLONG deadline = GetTickCount64() + timeoutMs;
        while (! _blockedReadReturned.load(std::memory_order_acquire))
        {
            if (GetTickCount64() >= deadline)
            {
                return false;
            }
            Sleep(10u);
        }
        return true;
    }

    [[nodiscard]] HRESULT BlockedReadStatus() const noexcept
    {
        return _blockedReadStatus.load(std::memory_order_acquire);
    }

    // Lets a still-wedged read complete so the process can shut down after a failed witness.
    void ReleaseBlockedRead() noexcept
    {
        if (_pipeWrite)
        {
            DWORD wrote = 0u;
            static_cast<void>(WriteFile(_pipeWrite.get(), "x", 1u, &wrote, nullptr));
        }
    }

    UncontainedNeverReturningFileSystem(const UncontainedNeverReturningFileSystem&)            = delete;
    UncontainedNeverReturningFileSystem& operator=(const UncontainedNeverReturningFileSystem&) = delete;
    UncontainedNeverReturningFileSystem(UncontainedNeverReturningFileSystem&&)                 = delete;
    UncontainedNeverReturningFileSystem& operator=(UncontainedNeverReturningFileSystem&&)      = delete;

    [[nodiscard]] uint64_t ProviderOperationCallCount() const noexcept
    {
        return _providerOperationCallCount.load(std::memory_order_acquire);
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** object) noexcept override
    {
        if (! object)
        {
            return E_POINTER;
        }
        *object = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *object = static_cast<IFileSystem*>(this);
        }
        else if (riid == __uuidof(IFileSystemPathCapabilities2))
        {
            *object = static_cast<IFileSystemPathCapabilities2*>(this);
        }
        else if (riid == __uuidof(IFileSystemRouteCapabilities) && _exposeTypedRoute)
        {
            *object = static_cast<IFileSystemRouteCapabilities*>(static_cast<FileSystemRouteCapabilitiesBase*>(this));
        }
        else
        {
            return E_NOINTERFACE;
        }
        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG remaining = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE GetPathCapabilities(const wchar_t* path, FileSystemOperation operation, const char** jsonUtf8) noexcept override
    {
        if (! jsonUtf8)
        {
            return E_POINTER;
        }
        *jsonUtf8 = nullptr;
        if (! path || path[0] == L'\0' || operation < FILESYSTEM_COPY || operation > FILESYSTEM_CREATE_DIRECTORY)
        {
            return E_INVALIDARG;
        }
        *jsonUtf8 = _capabilitiesJson.c_str();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t*, IFilesInformation**) noexcept override
    {
        return _boundedBlockedRead ? E_NOTIMPL : WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t*, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t*, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE DeleteItem(const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE
    RenameItem(const wchar_t*, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE
    CopyItems(const wchar_t* const*, unsigned long, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE
    MoveItems(const wchar_t* const*, unsigned long, const wchar_t*, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE
    DeleteItems(const wchar_t* const*, unsigned long, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE
    RenameItems(const FileSystemRenamePair*, unsigned long, FileSystemFlags, const FileSystemOptions*, IFileSystemCallback*, void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t*, FileSystemOperation, FileSystemTransferEndpoint, FileSystemTransferHints*) noexcept override
    {
        return _boundedBlockedRead ? E_NOTIMPL : WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t*, FileSystemStorageCharacteristics*) noexcept override
    {
        return _boundedBlockedRead ? E_NOTIMPL : WaitForeverAfterProviderOperationCall();
    }

protected:
    HRESULT BuildFileSystemRouteDescriptor(const wchar_t*, FileSystemOperation, FileSystemRouteDescriptor& descriptor) noexcept override
    {
        descriptor.providerId             = _boundedBlockedRead ? L"selftest/bounded-blocked-read" : L"selftest/uncontained-never-return";
        descriptor.pathProfileId          = _boundedBlockedRead ? L"selftest-bounded-blocked-read" : L"selftest-uncontained";
        descriptor.rootId                 = _boundedBlockedRead ? L"selftest-bounded-blocked-read/root" : L"selftest-uncontained/root";
        descriptor.availability           = FILESYSTEM_ROUTE_AVAILABLE;
        descriptor.cancellationRoute      = _boundedBlockedRead ? FILESYSTEM_CANCELLATION_BOUNDED : FILESYSTEM_CANCELLATION_UNCONTAINED;
        descriptor.namespaceKind          = FILESYSTEM_NAMESPACE_PROVIDER_VIRTUAL_FOLDER;
        descriptor.componentComparison    = FILESYSTEM_ROUTE_COMPONENT_ORDINAL_CASE_SENSITIVE;
        descriptor.caseOnlyRename         = FILESYSTEM_ROUTE_CASE_ONLY_UNSUPPORTED;
        descriptor.copyOperation          = _typedCopyOperation;
        descriptor.readOperation          = true;
        descriptor.writeOperation         = true;
        descriptor.pathTextStableIdentity = true;
        descriptor.exportCopyAll          = true;
        descriptor.importCopyAll          = true;
        return S_OK;
    }

private:
    HRESULT WaitForeverAfterProviderOperationCall() noexcept
    {
        _providerOperationCallCount.fetch_add(1u, std::memory_order_release);
        if (_boundedBlockedRead)
        {
            std::byte buffer[16]{};
            DWORD read           = 0u;
            const BOOL ok        = _pipeRead ? ReadFile(_pipeRead.get(), buffer, static_cast<DWORD>(sizeof(buffer)), &read, nullptr) : FALSE;
            const HRESULT status = ok != FALSE ? S_OK : HRESULT_FROM_WIN32(GetLastError());
            _blockedReadStatus.store(status, std::memory_order_release);
            _blockedReadReturned.store(true, std::memory_order_release);
            return FAILED(status) ? status : E_UNEXPECTED;
        }
        _neverRelease.wait(false, std::memory_order_acquire);
        return E_UNEXPECTED;
    }

    std::atomic_ulong _refCount{1u};
    std::atomic<uint64_t> _providerOperationCallCount{0u};
    std::atomic<bool> _neverRelease{false};
    std::atomic<bool> _blockedReadReturned{false};
    std::atomic<HRESULT> _blockedReadStatus{E_PENDING};
    wil::unique_handle _pipeRead;
    wil::unique_handle _pipeWrite;
    std::string _capabilitiesJson;
    bool _exposeTypedRoute   = true;
    bool _typedCopyOperation = true;
    bool _boundedBlockedRead = false;
};

[[nodiscard]] Common::Settings::FileOperationsSettings& EnsureFileOperationsSettingsForSelfTest() noexcept
{
    // An otherwise-default File Operations block may be pruned between cases. Tests that mutate
    // runtime-only File Operations settings must materialize their block immediately before use.
    if (! g_settings.fileOperations.has_value())
    {
        g_settings.fileOperations.emplace();
    }
    return g_settings.fileOperations.value();
}

struct RecycleBinBatchTestSnapshot final
{
    uint64_t batchCalls     = 0;
    uint64_t requestedItems = 0;
    uint64_t observedItems  = 0;
    uint64_t failedItems    = 0;
    uint64_t fallbackCount  = 0;
    uint64_t maxBatchSize   = 0;
};

[[nodiscard]] HRESULT ReadRecycleBinBatchTestSnapshot(IFileSystem* fileSystem, bool reset, RecycleBinBatchTestSnapshot& snapshot) noexcept
{
    using PfnGetSnapshot = HRESULT(__stdcall*)(BOOL, uint64_t*, uint64_t*, uint64_t*, uint64_t*, uint64_t*, uint64_t*);

    snapshot = {};
    if (fileSystem == nullptr)
    {
        return E_POINTER;
    }

    // The plugin manager can hold a live module independently of its staged pathname. Resolve the
    // module from this exact implementation's vtable so the snapshot always matches the object that
    // performed the operation.
    const auto objectWords    = reinterpret_cast<const void* const*>(fileSystem);
    const void* vtableAddress = objectWords[0];
    if (vtableAddress == nullptr)
    {
        return E_UNEXPECTED;
    }
    MEMORY_BASIC_INFORMATION moduleMemory{};
    if (VirtualQuery(vtableAddress, &moduleMemory, sizeof(moduleMemory)) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    if (moduleMemory.AllocationBase == nullptr)
    {
        return E_UNEXPECTED;
    }
    const HMODULE module = static_cast<HMODULE>(moduleMemory.AllocationBase);

    const FARPROC getSnapshotProc = GetProcAddress(module, "RedSalamanderFileSystemTestGetRecycleBinBatchSnapshot");
#pragma warning(push)
#pragma warning(disable : 4191) // Win32 exports are resolved as FARPROC; validate null before calling the typed test-only export.
    const auto getSnapshot = reinterpret_cast<PfnGetSnapshot>(getSnapshotProc);
#pragma warning(pop)
    if (getSnapshot == nullptr)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    return getSnapshot(reset ? TRUE : FALSE,
                       &snapshot.batchCalls,
                       &snapshot.requestedItems,
                       &snapshot.observedItems,
                       &snapshot.failedItems,
                       &snapshot.fallbackCount,
                       &snapshot.maxBatchSize);
}

[[nodiscard]] HWND FindCurrentProcessPopupWindow() noexcept
{
    // File Operations selftests can run while another RedSalamander process is
    // open. A desktop-global FindWindowW lookup may bind the fixture to that
    // process's popup and make task/popup ordering nondeterministic.
    HWND candidate = nullptr;
    while ((candidate = FindWindowExW(nullptr, candidate, kPopupClassName.data(), nullptr)) != nullptr)
    {
        DWORD processId = 0;
        static_cast<void>(GetWindowThreadProcessId(candidate, &processId));
        if (processId == GetCurrentProcessId())
        {
            return candidate;
        }
    }
    return nullptr;
}

constexpr std::wstring_view kSelfTestEnvConnFtp                          = L"REDSALAMANDER_SELFTEST_CONN_FTP";
constexpr std::wstring_view kSelfTestEnvConnSftp                         = L"REDSALAMANDER_SELFTEST_CONN_SFTP";
constexpr std::wstring_view kSelfTestEnvConnScp                          = L"REDSALAMANDER_SELFTEST_CONN_SCP";
constexpr std::wstring_view kSelfTestEnvConnImap                         = L"REDSALAMANDER_SELFTEST_CONN_IMAP";
constexpr std::wstring_view kSelfTestEnvConnS3                           = L"REDSALAMANDER_SELFTEST_CONN_S3";
constexpr std::wstring_view kSelfTestEnvConnS3Alt                        = L"REDSALAMANDER_SELFTEST_CONN_S3_ALT";
constexpr std::wstring_view kSelfTestEnvConnOneDrivePersonal             = L"REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL";
constexpr std::wstring_view kSelfTestEnvConnOneDriveBusiness             = L"REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS";
constexpr std::wstring_view kSelfTestEnvConnSharePoint                   = L"REDSALAMANDER_SELFTEST_CONN_SHAREPOINT";
constexpr std::wstring_view kSelfTestEnvBandwidthThrottleWorkerMode      = L"REDSALAMANDER_FILEOPS_BW_WORKER_MODE";
constexpr std::wstring_view kSelfTestEnvDeleteToctouSwapPath             = L"REDSALAMANDER_FILEOPS_DELETE_TOCTOU_SWAP_PATH";
constexpr std::wstring_view kSelfTestEnvDeleteToctouSwapTarget           = L"REDSALAMANDER_FILEOPS_DELETE_TOCTOU_SWAP_TARGET";
constexpr std::wstring_view kSelfTestEnvDeleteToctouSwapFired            = L"REDSALAMANDER_FILEOPS_DELETE_TOCTOU_SWAP_FIRED";
constexpr std::wstring_view kSelfTestEnvMoveDestinationRacePath          = L"REDSALAMANDER_FILEOPS_MOVE_DESTINATION_RACE_PATH";
constexpr std::wstring_view kSelfTestEnvMoveDestinationRaceFired         = L"REDSALAMANDER_FILEOPS_MOVE_DESTINATION_RACE_FIRED";
constexpr std::wstring_view kSelfTestEnvStagedCopyPromoteFailPath        = L"REDSALAMANDER_FILEOPS_STAGED_COPY_PROMOTE_FAIL_PATH";
constexpr std::wstring_view kSelfTestEnvStagedCopyPromoteFailFired       = L"REDSALAMANDER_FILEOPS_STAGED_COPY_PROMOTE_FAIL_FIRED";
constexpr std::wstring_view kSelfTestEnvFinalAttributesFailPath          = L"REDSALAMANDER_FILEOPS_FINAL_ATTRIBUTES_FAIL_PATH";
constexpr std::wstring_view kSelfTestEnvFinalAttributesFailFired         = L"REDSALAMANDER_FILEOPS_FINAL_ATTRIBUTES_FAIL_FIRED";
constexpr std::wstring_view kSelfTestEnvAbortOwnedStageUnknownPath       = L"REDSALAMANDER_FILEOPS_ABORT_OWNED_STAGE_UNKNOWN_PATH";
constexpr std::wstring_view kSelfTestEnvAbortOwnedStageUnknownFired      = L"REDSALAMANDER_FILEOPS_ABORT_OWNED_STAGE_UNKNOWN_FIRED";
constexpr std::wstring_view kSelfTestEnvDirectFinalRollbackSwapPath      = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ROLLBACK_SWAP_PATH";
constexpr std::wstring_view kSelfTestEnvDirectFinalRollbackSwapMovedPath = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ROLLBACK_SWAP_MOVED_PATH";
constexpr std::wstring_view kSelfTestEnvDirectFinalRollbackSwapFired     = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ROLLBACK_SWAP_FIRED";
constexpr std::wstring_view kSelfTestEnvDirectFinalAbortFailPath         = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ABORT_FAIL_PATH";
constexpr std::wstring_view kSelfTestEnvDirectFinalAbortFailFired        = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ABORT_FAIL_FIRED";
// I35: persistent (not one-shot) Local-plugin injection that makes BindObject report an identity-less
// location (FileIdInfo unsupported) for the configured root and everything below it.
constexpr std::wstring_view kSelfTestEnvBindIdentityUnsupportedPath = L"REDSALAMANDER_FILEOPS_BIND_IDENTITY_UNSUPPORTED_PATH";

constexpr std::wstring_view kSelfTestDefaultConnFtp              = L"FileOpsSelfTest FTP";
constexpr std::wstring_view kSelfTestDefaultConnSftp             = L"FileOpsSelfTest SFTP";
constexpr std::wstring_view kSelfTestDefaultConnScp              = L"FileOpsSelfTest SCP";
constexpr std::wstring_view kSelfTestDefaultConnImap             = L"FileOpsSelfTest IMAP";
constexpr std::wstring_view kSelfTestDefaultConnS3               = L"FileOpsSelfTest S3";
constexpr std::wstring_view kSelfTestDefaultConnS3Alt            = L"FileOpsSelfTest S3 Alt";
constexpr std::wstring_view kSelfTestDefaultConnOneDrivePersonal = L"FileOpsSelfTest OneDrive Personal";
constexpr std::wstring_view kSelfTestDefaultConnOneDriveBusiness = L"FileOpsSelfTest OneDrive Business";
constexpr std::wstring_view kSelfTestDefaultConnSharePoint       = L"FileOpsSelfTest SharePoint";

// HARD REQUIREMENT (Remote selftests):
// Any selftest phase that performs *remote* file operations (copy/move/delete) MUST be sandboxed to a dedicated,
// test-only folder/prefix on the remote side. Never run destructive tests against '/', a home directory, or any
// user-managed data. Remote file-op phases must:
//   - refuse/skip when the ConnectionProfile.initialPath is not a dedicated selftest root,
//   - create a unique per-run subfolder under that root,
//   - only touch/delete paths under that per-run subfolder.
// S3 gets one narrow exception: bucket root is allowed when the bucket name itself clearly contains 'selftest',
// because that bucket is already dedicated and the test phases still create an isolated per-run subfolder under it.
//
// Phase 16 currently validates secret retrieval for all remote plugins and also runs a sandboxed
// OneDrive Personal file-ops smoke path (remote read, local->remote copy, remote->local copy,
// remote move/rename, and remote delete). The sandbox requirement stays enforced for every remote
// provider so destructive coverage never escapes the dedicated selftest root.

constexpr ULONGLONG kDefaultTimeoutMs = 60'000ull;

// Keep the owned callback complete before SelfTestState instantiates unique_ptr cleanup.
struct WatchCallback final : public IFileSystemDirectoryWatchCallback
{
    WatchCallback()                                = default;
    WatchCallback(const WatchCallback&)            = delete;
    WatchCallback(WatchCallback&&)                 = delete;
    WatchCallback& operator=(const WatchCallback&) = delete;
    WatchCallback& operator=(WatchCallback&&)      = delete;

    std::atomic<uint64_t> callbackCount{0};
    std::atomic<uint64_t> overflowCount{0};
    std::atomic<uint64_t> badNotificationSizeCount{0};

    HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification, void* /*cookie*/) noexcept override
    {
        callbackCount.fetch_add(1, std::memory_order_relaxed);
        if (notification == nullptr || notification->sizeBytes != sizeof(FileSystemDirectoryChangeNotification))
        {
            badNotificationSizeCount.fetch_add(1, std::memory_order_relaxed);
        }

        if (notification && notification->sizeBytes == sizeof(FileSystemDirectoryChangeNotification) && notification->overflow)
        {
            overflowCount.fetch_add(1, std::memory_order_relaxed);
        }
        return S_OK;
    }
};

struct CompletedTaskInfo
{
    HRESULT hr                                    = S_OK;
    ULONGLONG completionTick                      = 0;
    bool discoveryClosed                          = false;
    bool firstMutationBeforeDiscoveryClosed       = false;
    bool discoverySkipped                         = false;
    uint64_t discoveredTotalBytes                 = 0;
    unsigned long discoveredFiles                 = 0;
    unsigned long discoveredDirectories           = 0;
    uint64_t copiedFileCount                      = 0u;
    bool copiedFileCountKnown                     = false;
    bool discoveredBytesKnown                     = false;
    uint64_t discoveryDurationUs                  = 0;
    uint64_t discoveryMaxQueueDepth               = 0;
    uint64_t discoveryStarvationCount             = 0;
    uint64_t discoveryCallbackCount               = 0;
    uint64_t discoveryCallbackUs                  = 0;
    uint64_t discoveryLockWaitUs                  = 0;
    uint64_t discoveryFirstMutationUs             = (std::numeric_limits<uint64_t>::max)();
    uint64_t discoveryBytesCompletedWhileOpen     = (std::numeric_limits<uint64_t>::max)();
    uint64_t discoveryMutationsCompletedWhileOpen = (std::numeric_limits<uint64_t>::max)();
    // Must stay 0: discovery closure is one-way and means the totals are final.
    uint64_t discoveryGrowthAfterCloseCount              = (std::numeric_limits<uint64_t>::max)();
    uint64_t progressCallbackCount                       = 0;
    bool started                                         = false;
    unsigned long progressTotalItems                     = 0;
    unsigned long progressCompletedItems                 = 0;
    uint64_t progressCompletedBytes                      = 0;
    unsigned long completedFiles                         = 0;
    unsigned long completedFolders                       = 0;
    uint64_t conflictWaitUs                              = 0;
    uint64_t conflictConvergenceWaitUs                   = 0;
    uint64_t conflictPromptCount                         = 0;
    uint64_t interlockWaitUs                             = 0;
    uint64_t interlockWaitCount                          = 0;
    uint64_t bridgeDirectoryEnsureCount                  = 0;
    uint64_t bridgeSourceDirectoryEnumerationCount       = 0;
    uint64_t bridgeFileAdmissionCount                    = 0;
    uint64_t bridgeMaxActiveFileTransfers                = 0;
    uint64_t bridgeEarlyFileStartCount                   = 0;
    uint64_t bridgeAdmissionMaxQueueDepth                = 0;
    uint64_t bridgeTraversalMaxDepth                     = 0;
    uint64_t bridgeTraversalMaxRetainedEntries           = 0;
    uint64_t bridgeTraversalMaxQueuedPathBytes           = 0;
    uint64_t bridgeTraversalMaxMetadataBytes             = 0;
    uint64_t bridgeTraversalLimitHitCount                = 0;
    uint64_t conflictExpectedDestinationBoundCount       = 0;
    uint64_t conflictExpectedDestinationUnavailableCount = 0;
    uint64_t conflictExpectedDestinationReturnedCount    = 0;
    uint64_t bridgeImmediateDestinationSizeProbeUs       = 0;
    uint64_t bridgeImmediateDestinationSizeProbeCount    = 0;
    uint64_t bridgeCommitSizeProofCount                  = 0;
    uint64_t bridgeCommitSizeProofFallbackCount          = 0;
    uint64_t verificationUs                              = 0;
    uint64_t verificationReadBytes                       = 0;
    uint64_t verificationReadCalls                       = 0;
    uint64_t verificationProviderProofCount              = 0;
    uint64_t verificationHostReadbackCount               = 0;
    uint64_t verificationTotalBytes                      = 0;
    uint64_t verificationCompletedBytes                  = 0;
    unsigned int configuredMaxConcurrency                = 0;
    std::vector<std::optional<HRESULT>> sourceItemStatuses;
    std::vector<std::optional<FileOperations::FileOperationItemResult>> sourceItemResults;
    // I35 R10.5: the progress-report cadence the host measured for the whole task.
    FolderWindow::FileOperationState::Task::ProgressCadenceStats progressCadence{};
    // I35 R10.5: the destination storage kind the concurrency budget consulted (UNKNOWN when it did not).
    uint32_t autoConcurrencyDestinationStorageKind = FILESYSTEM_STORAGE_UNKNOWN;
    // I35 R10.2: the side of the transfer a stall of the task would be named after.
    FolderWindow::FileOperationState::ProgressStallSide progressStallSide = FolderWindow::FileOperationState::ProgressStallSide::Storage;
    // The host bridge's stage writers: one per file the bridge copied (zero on a provider's own route).
    uint64_t bridgeStageCreateCount = 0;
};

// This fixture owns a local Cloud Files registration across asynchronous host-copy ticks.
// Its path storage outlives the WIL registration owner; cleanup records its HRESULT for assertions.
struct PlaceholderSyncRoot final
{
    using SetCompatibilityModeFn = CHAR(NTAPI*)(CHAR);
    std::filesystem::path path;
    HRESULT cleanupHr                     = S_OK;
    bool registered                       = false;
    SetCompatibilityModeFn setProcessMode = nullptr;
    SetCompatibilityModeFn setThreadMode  = nullptr;
    CHAR previousProcessMode              = -1;
    CHAR previousThreadMode               = -1;
};

void UnregisterPlaceholderSyncRoot(PlaceholderSyncRoot* root) noexcept
{
    root->cleanupHr  = root->registered ? CfUnregisterSyncRoot(root->path.c_str()) : S_OK;
    root->registered = false;
    if (root->previousThreadMode >= 0 && root->setThreadMode(root->previousThreadMode) < 0)
    {
        root->cleanupHr = E_FAIL;
    }
    if (root->previousProcessMode >= 0 && root->setProcessMode(root->previousProcessMode) < 0)
    {
        root->cleanupHr = E_FAIL;
    }
    root->previousThreadMode  = -1;
    root->previousProcessMode = -1;
    if (FAILED(root->cleanupHr))
    {
        Debug::Error(L"Placeholder fixture registration/mode cleanup failed (hr=0x{:08X}).", static_cast<unsigned long>(root->cleanupHr));
    }
}

using UniquePlaceholderSyncRoot = wil::unique_any<PlaceholderSyncRoot*, decltype(&UnregisterPlaceholderSyncRoot), UnregisterPlaceholderSyncRoot>;

struct SelfTestState
{
    // Explicitly delete copy/move operations (self-test state is not copyable/movable).
    SelfTestState()                                = default;
    SelfTestState(const SelfTestState&)            = delete;
    SelfTestState(SelfTestState&&)                 = delete;
    SelfTestState& operator=(const SelfTestState&) = delete;
    SelfTestState& operator=(SelfTestState&&)      = delete;

    enum class Step
    {
        Idle,
        Setup,
        FileOps_CommonPresentationPerf,
        FileOps_IntegratedPresentationPerf,
        FileOps_VisualGallery,
        FileOps_VisualGalleryInteraction,
        FileOps_QueuedAdmissionContracts,
        FileOps_SafetyRemediationContracts,
        FileOps_ConflictMessageText,
        FileOps_ReviewRemediationContracts,
        FileOps_RepeatDecisionContracts,
        FileOps_SkipCountingPresentation,
        FileOps_PermanentDeleteCompletionProgress,
        FileOps_AdmissionBaseline,
        FileOps_OwnedProductionShutdown,
        FileOps_CrossVolumeMoveAfterPartialPublicationStorage,
        FileOps_RepeatedNonClipboardCopyStorage,
        FileOps_RepeatedNonClipboardMoveStorage,
        FileOps_TerminalFactsProjection,
        FileOps_ClearCompletedScope,
        FileOps_PerItemCancelUndispatchedRoots,
        FileOps_AdmissionFailureReachesPane,
        FileOps_DeferredExitRefusesNewWork,
        FileOps_CompareWindowCloseConsent,
        FileOps_PopupFailureSurfaceRecovery,
        FileOps_PopupStableHostedTree,
        FileOps_PopupIdleCadence,
        FileOps_AlertOverlayLifetime,
        FileOps_AlertOverlayUxAccessibility,
        FileOps_AlertOverlayModelessDecision,
        FileOps_DecisionMirror,
        FileOps_DecisionMirrorSameHostOverlap,
        FileOps_DecisionMirrorIngressAnchor,
        FileOps_DecisionMirrorCompareWindow,
        FileOps_DecisionMirrorConsecutivePrompts,
        FileOps_DecisionMirrorQueueOrder,
        FileOps_BatchRenameExactResultCarrier,
        FileOps_BatchRenameNeverRunRowsNotAttempted,
        FileOps_NotesSeverity,
        FileOps_NotesMetadata,
        FileOps_CrossProviderMtpNewFile,
        FileOps_VerifyAvailability,
        FileOps_SizelessSourceNote,
        FileOps_CrossRootLocalCopy,
        FileOps_ArchiveNestedCopyLocal,
        FileOps_ArchiveNestedCopyUnc,
        FileOps_CancelAllNeverStarted,
        FileOps_SmbCopyThroughputVariance,
        FileOps_LocalCopyThroughput,
        FileOps_SmbNativeRouteConcurrencyBudget,
        FileOps_SameProviderBridgeDestinationStorage,
        FileOps_SmbManagedMoveLeavesNoStage,
        FileOps_RateWindowDisplay,
        FileOps_ProgressPhaseAccounting,
        FileOps_StallLabelLayout,
        FileOps_SmallFileCopyCost,
        Delete_IdentityLessLocationByName,
        Delete_IdentityLessSharingAndInterlock,
        Delete_IdentityLessQueuedRootOverlap,
        Delete_IdentityLessPinnedTreeOtherSpelling,
        Delete_IdentityLessPinnedTreeArchiveWriter,
        Delete_IdentityLessPinnedTreeCardPending,
        Delete_IdentityLessPinnedTreeConcurrentWriter,
        Delete_IdentityLessPinnedTreeBudget,
        Delete_IdentityLessPinnedTreeChangedEntry,
        Delete_IdentityLessPinnedTreeEscalationPartial,
        Delete_IdentityLessPinnedTreeLinkSwap,
        Delete_IdentityLessPinnedTreeSubfolderOpen,
        Delete_IdentityLessPinnedTreeRecycleOverBudget,
        Delete_IdentityLessPinnedTreeHolders,
        Delete_IdentityLessPinnedTreeScale,
        Delete_IdentityLessFolderChangeAndHolders,
        Delete_IdentityLessOpenFailureAndArchiveOrigins,
        FileOps_IdentityLessGuardPartition,
        FileOps_LeftoverCleanupWalkBounds,
        FileOps_OwnedStageCleanupBudget,
        Delete_NasFileIdInfoUnsupportedByName,
        FileOps_CopyMergeIntoExistingFolder,
        FileOps_MoveMergeIntoExistingFolderSameVolume,
        Beeline_RenameMergeSkipKeepsSourceFolder,
        FileOps_RenameMergeEmptyFolderCompletes,
        FileOps_ReparseDirectoryMergeIntoExistingFolder,
        FileOps_ProviderCapabilityMatrix,
        FileOps_ParallelGraphFairColorWeight,
        FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof,
        FileOps_ReparseLiteralCopyKeepsTargets,
        FileOps_DeleteToctouSwapGuard,
        FileOps_NativeMoveRacedDestinationNotReplaced,
        FileOps_NativeMoveMissingPathKnownFailure,
        FileOps_ResolvedItemsExactDestinations,
        Riptide_MoveDistinctSameSizeFilePreservesSource,
        Riptide_ReparseNativeMoveRelocatesLinkObject,
        Fairstream_ManagedMovePreservesUncopiedNewFiles,
        Fairstream_MoveMergeReadonlyDestinationFolder,
        Fairstream_OverwriteGrantIsOneShot,
        Fairstream_ConflictPromptsSerializedUnderParallelism,
        Fairstream_ReparseReplaceNonEmptyDirRequiresConsent,
        Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate,
        Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent,
        Riptide_ReparseMoveSourceNeverUsesDirectoryMerge,
        Fairstream_MovedTreeKeepsLiteralLinks,
        Fairstream_CrossFsConcurrentMoveUsesBridge,
        Fairstream_RerunCopyIdenticalCollisionPrompts,
        Fairstream_MoveSameSizeCollisionPrompts,
        Fairstream_GraphBandsFairUnderParallelCopy,
        Fairstream_ManagedMergeMovePreservesChildren,
        Fairstream_JunctionMergeTargetRequiresConsent,
        Fairstream_CopyIntoSelfAliasRejected,
        Fairstream_StorageKindProbed,
        Fairstream_ParallelDeleteContinuesPastLockedChild,
        Fairstream_BridgePerFileConflictSkips,
        Cinderstar_BridgeMovePerFileConflictSkipsParallel,
        Riptide_BridgeNestedDirVsFileSkipContinuesSiblings,
        Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch,
        Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial,
        Fairstream_SaturationConcurrentCopiesMakeProgress,
        Fairstream_DiscoveryAheadOverlapsTransfer,
        Riptide_LiveFinishedSnapshotCarriesDiagnostics,
        Riptide_BridgeSequentialContinueOnErrorCopiesSiblings,
        Causeway_BridgeRejectsHostileChildNames,
        Causeway_BridgeProviderOutputContracts,
        Causeway_BridgeSchedulingAndResourceContracts,
        Causeway_BridgeFileReparsePolicy,
        Beeline_CopySkipLinksKeepsPlaceholders,
        Causeway_BridgeFailureStatusAndPausedReader,
        Floodgate_CrossFsCopyGetSizeFailureRefusesCommit,
        Floodgate_CrossFsMoveGetSizeFailurePreservesSource,
        Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource,
        Cinderstar_LegacyWriterEventuallyConsistentMove,
        Cinderstar_LegacyWriterPermanentMissMove,
        Cinderstar_LegacyWriterWrongSizeMove,
        Cinderstar_LegacyWriterConcurrentReplacementCopy,
        Cinderstar_LegacyWriterCancelDuringBackoff,
        Floodgate_CrossFsCopyOnlyUsesCopyVerification,
        Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree,
        Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup,
        Floodgate_CrossFsDirectoryCopyOnlyRetainsSource,
        Floodgate_LocalWriterOverwriteIsStaged,
        Floodgate_LocalCopyOverwriteIsStaged,
        Floodgate_LocalCopyNewNameConcurrentReplacementSurvives,
        Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete,
        Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker,
        Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker,
        Floodgate_InlineF2WorkerQueueBypassAndInterlock,
        C1_InlineRenameNameRefusedOnCard,
        R4A19_DiscoveryMeasurementFacts,
        R4A19_DiscoveryProviderControls,
        R4A19_DiscoveryIndependentVolumes,
        R0fSmb_BlockedSynchronousCallCancelReturns,
        R0fSmb_LoopbackReadWriteCreateDelete,
        R0fSmb_LoopbackRecycleNeverDestroys,
        R0fSmb_NestedTreeCopyToShare,
        R0fSmb_BridgeTreeCopyToShare,
        R0fSmb_ManagedMoveTreeToShare,
        R0fCurl_FakeFtpReadWriteCreateDelete,
        BR5_CurlHostDirectoryMove,
        BR5_CurlHostDirectoryMoveRefused,
        R0fS3_FakeS3ReadWriteCreateDelete,
        R0fGraph_FakeGraphReadWriteCreateRenameRecycle,
        R0fGDrive_FakeDriveReadWriteCreateMoveDelete,
        C0_GDriveCommittedCopyResponseLost,
        C0_GDriveCommittedDeleteResponseLost,
        C0_NativeCopyFailurePreservesKnownAxes,
        C0_MutationReceiptPrefixBoundary,
        C0_LocalKnownNoCommitReceipts,
        C0_CurlNativeDeleteGuards,
        C0_CurlCommittedMutationResponseLost,
        C0_CurlHostCommittedDeleteResponseLost,
        C0_CurlPartialTreeFailure,
        C0_CurlHostPartialTreeFailure,
        C0_CurlNativeDeleteLateListing,
        C0_CurlDeleteTraversalBounds,
        C0_CurlDirectorySizeTruth,
        C0_CurlMovePreflightTruth,
        C0_CurlCopyTraversalTruth,
        C0_CurlEntryLookupTruth,
        C0_CurlImapListingTruth,
        C0_CurlImapTransportTruth,
        R3_1_IdentityLessReplaceDummy,
        R3_2_WriterProofDummy,
        R3_3_PublicationFaultMatrix,
        R3_4_ReplaceWithoutOccupantTokenRefused,
        RC3_9_TransferDestinationNameRefused,
        R4T_DeepTreeCopyCompletesIteratively,
        R4T2_DeepTreeLocalDeleteCompletes,
        R4T3_WideDirectoryCopyCompletes,
        R4T4_LocalManyDirectoriesCopy,
        R4T5_LocalCopyHoldsNoDirectoryHandle,
        R4T6_LocalDirectoryRestoreChecksIdentity,
        R4T7_BridgeCopyHoldsNoDirectoryHandle,
        R4T8_BridgeDirectoryRestoreChecksIdentity,
        R4T9_BridgeManyDirectoriesCopy,
        R4A02_DeleteOverCopySourceWarnsAndQueues,
        R4A02_DontStartCancelsBeforeMutation,
        R4A02_ReadReadDoesNotWarn,
        R4A02_SameDestinationWarns,
        R4A02_RunConcurrentSameDestinationStaysSafe,
        R4A02_LiveOutputGuardChoices,
        R4A02_DisjointWritesDoNotWarn,
        R4A02_RenamePublishKeepsIndexPrecise,
        R4A02_LatePublisherStillWarns,
        Phase5_DiscoverySingleTraversal,
        Phase5_DiscoveryCancelReleasesSlot,
        Phase5_CleanupDrainsOutstandingTasks,
        Phase5_DiscoveryCancelLatencyLocal,
        Phase5_DiscoverySkipContinues,
        Phase5_CancelQueuedTask,
        BR3_CancelQueuePausedTransfer,
        Phase5_SwitchParallelToWaitDuringDiscovery,
        Phase5_SwitchWaitToParallelResume,
        Phase6_PopupRateSmoothing,
        Phase6_PopupSmokeResizeAndPause,
        Phase6_DeleteBytesMeaningful,
        Phase6_LocalBandwidthThrottle,
        Phase6_ParallelBandwidthThrottleFairness,
        Phase7_WatcherChurn,
        Phase7_CacheBorrowNoWatchInvalidation,
        Phase7_CrossPaneVisibleRefreshLocal,
        Phase7_CrossPaneVisibleRefreshDummy,
        Phase7_CrossPaneRelocateLocal,
        Phase7_LargeDirectoryEnumeration,
        Phase7_ParallelCopyMoveKnobs,
        Phase7_CopyMoveConcurrency16Perf,
        Phase7_AutoConcurrencyHints,
        Phase7_PerItemDirectoryCopyInFlightLines,
        Phase7_CopyItemsSingleFolderRecursiveParallelism,
        Phase7_CopyItemsMultiRootUnevenRecursiveParallelism,
        Phase7_CopyRecursiveParallelismMatrix,
        Phase7_CopyItemsSaturatedPoolProducersHelp,
        Phase7_CopyItemWaitingProducerSeesCancel,
        Phase7_CopyItemWalkAtLimitParks,
        Phase7_DeleteItemsNestedRootsWake,
        Phase7_SharedPerItemScheduler,
        Phase7_ParallelDeleteKnobs,
        Phase7_RecycleBinBatchDelete,
        Phase7_RecycleBinBatchDeleteMultiBatch,
        Phase7_RecycleBinTestItemsPurged,
        Phase8_DefaultBandwidthLimitFromSettings,
        Phase8_TightDefaults_NoOverwrite,
        Phase8_InvalidDestinationRejected,
        Phase8_InvalidSizeBytesRejected,
        Phase8_PerItemOrchestration,
        Phase9_ConflictPrompt_OverwriteReplaceReadonly,
        Phase9_ConflictPrompt_ApplyToAllUiCache,
        Phase9_ConflictPrompt_KeepBothNestedCacheEligibility,
        Phase9_ConflictPrompt_TypeMismatchNoOverwrite,
        Phase9_ConflictPrompt_LocalFileOntoDirectory,
        Phase9_ConflictPrompt_SkipApplyToAll,
        Phase9_ConflictPrompt_RetryCap,
        Phase9_ConflictPrompt_SkipContinuesDirectoryCopy,
        Phase9_PerItemConcurrency,
        Phase10_PermanentDelete,
        Phase10_DeferredConsentAndRecycleEscalation,
        Phase10_RecycleDestroyedIsNeverRecycled,
        Phase10_RecycleEscalationDirectoryTree,
        Phase10_RecycleEscalationTaskScopedGrant,
        Phase10_RecycleEscalationBulkItemsContinueOnError,
        Phase10_RecycleEscalationBulkItemsResumeBounded,
        Phase10_MetadataPreservationAndSourceRetention,
        Phase10_TypedResultsAndConsumers,
        Phase10_ClipboardAdmissionAndRetainedActions,
        Phase10_ContentVerification,
        Phase10_ArtifactTouchGuard,
        R1d_PreparingLifecycle,
        C1_ExitCloseDeferredUntilTasksQuiet,
        Beeline_SameVolumeTreeMoveIsRename,
        Beeline_LoopbackShareMoveIsRename,
        C1_PermanentDeleteConfirmsOnCard,
        Phase11_CrossFileSystemBridge,
        Phase11_BridgeSingleFolderParallelCopyInFlightLines,
        Phase11_BridgeMultiFolderParallelCopyInFlightLines,
        Phase11_BridgePipelineDummyToDummyPerf,
        Phase11_ConnectionOverridePrecedence,
        Phase11_ConnectionOverrideGlobalGate,
        Phase11_ConnectionOverrideClamp,
        Phase12_ReparsePointPolicy,
        Phase13_PostMortemDiagnostics,
        Phase14_PopupHostLifetimeGuard,
        Phase15_FileSystem7zReadSeekSmoke,
        Phase15_FileSystem7zMountPathImpact,
        Phase16_RemoteWatchContractExposure,
        Phase16_RemoteFtpSecret,
        Phase16_RemoteFtpSandbox,
        Phase16_RemoteSftpSecret,
        Phase16_RemoteSftpSandbox,
        Phase16_RemoteScpSecret,
        Phase16_RemoteScpSandbox,
        Phase16_RemoteImapSecret,
        Phase16_RemoteImapSandbox,
        Phase16_RemoteS3Secret,
        Phase16_RemoteS3Sandbox,
        Phase16_RemoteS3FileOps,
        Phase16_RemoteOneDrivePersonalSecret,
        Phase16_RemoteOneDrivePersonalSandbox,
        Phase16_RemoteOneDrivePersonalFileOps,
        Phase16_RemoteOneDriveBusinessSecret,
        Phase16_RemoteOneDriveBusinessSandbox,
        Phase16_RemoteSharePointSecret,
        Phase16_RemoteSharePointSandbox,
        DiscoveryScope_LeafCopy,
        DiscoveryScope_LeafMove,
        DiscoveryScope_ManagedDirectory,
        DiscoveryScope_NativeRenameMerge,
        DiscoveryScope_MixedRoots,
        DiscoveryScope_ZeroByteAndVerifiedLeaf,
        DiscoveryScope_LocalDirectApi,
        DiscoveryScope_DummyDirectApi,
        DiscoveryScope_CurlDirectApi,
        DiscoveryScope_S3DirectApi,
        DiscoveryScope_GDriveDirectApi,
        DiscoveryScope_GraphDirectApi,
        DiscoveryScope_MtpDirectApi,
        DiscoveryScope_FolderClosesBeforeLastFile,
        DiscoveryScope_WideFolderClosesWithinQueueLead,
        DiscoveryScope_NestedTreeMetadataRestoredAfterChildren,
        DiscoveryScope_ManagedMoveListingCompleteBeforeRemoval,
        DiscoveryScope_WorkerFailureIsNotCanceled,
        DiscoveryScope_FlatFolderUsesFullConcurrency,
        DiscoveryScope_ListingSilenceNoted,
        DiscoveryScope_TenThousandSmallFilesCloseAfterListing,
        DiscoveryScope_RetryDoesNotGrowAfterClose,
        DiscoveryScope_HeldFolderFailureStillCloses,
        DiscoveryScope_FolderLiveProgress,
        DiscoveryScope_MultipleFoldersLiveProgress,
        FileOps_SelectionShapeLocal,
        FileOps_SelectionShapeSmb,
        Cleanup_RestorePluginConfig,
        Done,
        Failed,
    };

    std::atomic<bool> running{false};
    std::atomic<bool> done{false};
    std::atomic<bool> failed{false};
    Step step = Step::Idle;
    SelfTest::SelfTestOptions options;
    std::wstring runFilter;
    std::vector<Step> activePhaseOrder;
    std::vector<Step> reportedPhaseOrder;
    uint32_t stepState                        = 0;
    uint64_t batchRenameExactPromptGeneration = 0u;
    uint64_t batchRenameExactCurrentTaskId    = 0u;
    std::optional<bool> batchRenameExactAutoAcceptBefore;
    wil::com_ptr<IFileSystem> batchRenameExactPaneFileSystem;
    uint64_t runStartTick = 0;

    std::vector<SelfTest::SelfTestCaseResult> phaseResults;
    bool phaseInProgress     = false;
    ULONGLONG phaseStartTick = 0;
    bool phaseFailed         = false;
    std::wstring phaseName;
    std::wstring phaseFailureMessage;

    HWND mainWindow = nullptr;

    std::filesystem::path tempRoot;
    // Test-only storage proof failure preserves governed scratch after normal worker drain.
    bool retainTempRootOnFailure = false;
    std::filesystem::path fileOpsAlternateVolumeRoot;
    PlaceholderSyncRoot placeholderSyncRootStorage;
    UniquePlaceholderSyncRoot placeholderSyncRoot;

    wil::com_ptr<IFileSystem> fsLocal;
    wil::com_ptr<IInformations> infoLocal;
    std::string localConfigOriginal;
    bool localConfigDirty = false;

    wil::com_ptr<IFileSystem> fsDummy;
    wil::com_ptr<IInformations> infoDummy;
    std::string dummyConfigOriginal;
    bool dummyConfigDirty = false;

    // The module pin is declared before its objects so the COM owners are released first.
    std::unique_ptr<RedSalamander::TestSupport::ScopedEnvironmentVariable> r4A19MtpProfile;
    std::optional<std::wstring> r4A19MtpOriginalLocalAppData;
    wil::unique_hmodule r4A19MtpModule;
    wil::com_ptr<IFileSystem> r4A19Mtp;
    wil::com_ptr<IFileSystemIO> r4A19MtpIo;

    // Keep the DLL and result storage alive until the native fixture worker has joined.
    wil::unique_hmodule c0CurlModule;
    unsigned int c0CurlPassed = 0u;
    unsigned int c0CurlFailed = 0u;
    std::atomic<HRESULT> c0CurlResult{E_PENDING};
    std::jthread c0CurlWorker;

    // C0_LocalKnownNoCommitReceipts holds the moved file open (no FILE_SHARE_DELETE) until the task ends.
    wil::unique_hfile c0LocalLockedHandle;

    bool fileOperationsBackedUp = false;
    std::optional<Common::Settings::FileOperationsSettings> fileOperationsOriginal;
    bool appThemeBackedUp = false;
    std::optional<AppTheme> appThemeOriginal;

    bool connectionsBackedUp = false;
    std::optional<Common::Settings::ConnectionsSettings> connectionsOriginal;
    std::wstring connOverrideProfileName;

    wil::com_ptr<IFileSystem> fs7z;
    wil::com_ptr<IInformations> info7z;
    std::string config7zOriginal;
    bool config7zDirty = false;

    std::vector<std::wstring> dummyPaths;
    std::string dummySeedConfig;

    wil::com_ptr<IFileSystem> fsRemoteS3;
    std::wstring remoteS3ProfileName;
    std::wstring remoteS3CaseRootConn;
    std::wstring remoteS3AltCaseRootConn;
    std::wstring remoteS3UploadDirConn;
    std::wstring remoteS3MoveDirConn;
    std::wstring remoteS3SeedFileConn;
    std::wstring remoteS3UploadedFileConn;
    std::wstring remoteS3MovedFileConn;
    std::wstring remoteS3RenamedFileConn;
    std::wstring remoteS3UploadedDirConn;
    std::filesystem::path remoteS3DisplayPath;
    std::filesystem::path remoteS3Workspace;
    std::filesystem::path remoteS3UploadSource;
    std::filesystem::path remoteS3DownloadDir;
    std::filesystem::path remoteS3DownloadedFile;
    std::filesystem::path remoteS3DirUploadSource;
    std::filesystem::path remoteS3DirMoveDownloadDir;
    std::filesystem::path remoteS3DirMovedFile;
    std::string remoteS3Payload;

    wil::com_ptr<IFileSystem> fsRemoteOneDrivePersonal;
    std::wstring remoteOneDrivePersonalProfileName;
    std::wstring remoteOneDrivePersonalCaseRootConn;
    std::wstring remoteOneDrivePersonalUploadDirConn;
    std::wstring remoteOneDrivePersonalMoveDirConn;
    std::wstring remoteOneDrivePersonalSeedFileConn;
    std::wstring remoteOneDrivePersonalUploadedFileConn;
    std::wstring remoteOneDrivePersonalMovedFileConn;
    std::filesystem::path remoteOneDrivePersonalDisplayPath;
    std::filesystem::path remoteOneDrivePersonalWorkspace;
    std::filesystem::path remoteOneDrivePersonalUploadSource;
    std::filesystem::path remoteOneDrivePersonalDownloadDir;
    std::filesystem::path remoteOneDrivePersonalDownloadedFile;
    std::string remoteOneDrivePersonalPayload;

    FolderWindow* folderWindow                = nullptr;
    FolderWindow::FileOperationState* fileOps = nullptr;

    std::optional<std::uint64_t> taskA;
    wil::com_ptr_nothrow<IInvokeProvider> queuedStartAction;
    wil::com_ptr_nothrow<IInvokeProvider> queuedReorderAction;
    wil::com_ptr_nothrow<IInvokeProvider> queuedCancelAction;
    wil::com_ptr_nothrow<IInvokeProvider> discoverySkipAction;
    std::optional<std::uint64_t> taskB;
    std::optional<std::uint64_t> taskC;
    bool r4a02ConcurrentAdmissionObserved = false;
    uint32_t r4a02LiveOutputScenario      = 0u;
    bool r4a02LiveOutputQueueObserved     = false;
    bool inlineF2QueueModeOriginal        = true;
    bool r4a02QueueModeOriginal           = true;
    std::optional<bool> r4A19IndependentQueueModeOriginal;
    HWND c1DeferredCloseTarget           = nullptr;
    ULONGLONG c1HeartbeatStartTick       = 0;
    unsigned c1HeartbeatTicks            = 0u;
    uint64_t inlineF2PresentedBaseline   = 0;
    uint64_t inlineF2SilentCleanBaseline = 0;
    std::optional<std::uint64_t> queuePausedTask;
    RECT popupOriginalRect{};
    bool popupOriginalRectValid = false;

    wil::com_ptr<IFileSystemDirectoryWatch> directoryWatch;
    std::unique_ptr<WatchCallback> directoryWatchCallback;
    std::filesystem::path watchDir;
    uint64_t watchCounter = 0;
    wil::unique_handle lockedFileHandle;
    wil::unique_handle lockedDestinationHandle;

    size_t copyKnobIndex                         = 0;
    size_t copyKnobRetryCount                    = 0;
    size_t deleteKnobIndex                       = 0;
    bool copySpeedLimitCleared                   = false;
    bool copyPromptValidated                     = false;
    bool copyKnobObservedPerCallShare            = false;
    unsigned int copyKnobObservedActiveCalls     = 0;
    uint64_t copyKnobObservedDesiredSpeedLimit   = 0;
    uint64_t copyKnobObservedAppliedSpeedLimit   = 0;
    uint64_t copyKnobObservedEffectiveSpeedLimit = 0;
    ULONGLONG copyTaskStartTick                  = 0;
    size_t connGateMaxActiveCopyStreams          = 0;
    bool connGateObservedSaturation              = false;
    size_t keepBothNestedScenarioIndex           = 0;
    ULONGLONG keepBothNestedScenarioStartTick    = 0;
    std::atomic<bool> phase10ClipboardPublishedNonRunning{false};
    std::atomic<bool> phase10ClipboardReadinessBarrierCalled{false};
    std::atomic<bool> phase10ClipboardSlowGateRelease{false};
    bool phase10ClipboardReturnedBeforeReadiness = false;
    std::atomic<bool> r1dPreparingGateEntered{false};
    std::atomic<bool> r1dPreparingGateRelease{false};
    bool r1dAdmissionReturnedBeforeReadiness   = false;
    bool r1dMoveBreadcrumbPublishedBeforeReady = false;
    uint64_t r1dPresentationBaseline           = 0u;
    bool phase10MetadataSparseSeeded           = false;
    int64_t phase10MetadataExpectedWriteTime   = 0;

    std::wstring failureMessage;
    bool autoDismissSuccessOriginal            = false;
    ULONGLONG stepStartTick                    = 0;
    ULONGLONG markerTick                       = 0;
    size_t recycleMalformedRound               = 0;
    DWORD beelineVolumeSerial                  = 0;
    ULONGLONG beelineFileIndex                 = 0;
    uint64_t beelineCopyMs                     = 0;
    ULONGLONG r4A19ScenarioStartTick           = 0;
    uint64_t r4A19VolumeCIsolatedUs            = 0;
    uint64_t r4A19VolumeDIsolatedUs            = 0;
    bool r4A19IndependentProgressOverlap       = false;
    ULONGLONG lastProgressLogTick              = 0;
    unsigned long pauseResumeBaselineItems     = 0;
    uint64_t pauseResumeBaselineBytes          = 0;
    size_t baselineThreadCount                 = 0;
    ULONGLONG localBandwidthRunStartTick       = 0;
    ULONGLONG localBandwidthCancelHandlerTick  = 0;
    ULONGLONG localBandwidthCancelStartTick    = 0;
    ULONGLONG localBandwidthCancelReturnTick   = 0;
    uint64_t localBandwidthDurationUs          = 0;
    uint64_t localBandwidthDurationLeadUs      = 0;
    uint64_t localBandwidthCancelLatencyUs     = 0;
    uint64_t localBandwidthMaxWindowBytes      = 0;
    uint64_t localBandwidthMaxSampleDeltaBytes = 0;
    std::vector<std::pair<ULONGLONG, uint64_t>> localBandwidthSamples;
    bool bandwidthThrottleWorkerModeEnvBackedUp    = false;
    bool bandwidthThrottleWorkerModeEnvHadOriginal = false;
    std::wstring bandwidthThrottleWorkerModeEnvOriginal;

    // Holds an exclusive handle across ticks for Fairstream_ParallelDeleteContinuesPastLockedChild.
    wil::unique_handle holdOpenHandle;

    // Best graph-fairness layout values observed while the task ran (the popup can auto-close
    // at completion, so the final poll may no longer see it).
    uint32_t fairstreamGraphMaxMultiBuckets             = 0;
    uint32_t fairstreamGraphSingleBuckets               = 0;
    uint32_t fairstreamGraphDistinctHues                = 0;
    double fairstreamGraphMinShare                      = 0.0;
    double fairstreamGraphMaxShare                      = 0.0;
    size_t fairstreamGraphMaxInFlight                   = 0;
    uint32_t fairstreamGraphAccumCalls                  = 0;
    uint32_t fairstreamGraphLastPending                 = 0;
    uint32_t fairstreamGraphAccumMaxStreams             = 0;
    bool fairstreamGraphFairnessObserved                = false;
    size_t fairstreamSaturationDispatchMetricBaseline   = 0;
    size_t fairstreamSaturationMaxDistinctInFlightTrees = 0;
    size_t storageClampMaxInFlight                      = 0;

    ULONGLONG parallelBandwidthRunStartTick                  = 0;
    uint64_t parallelBandwidthBaselineUs                     = 0;
    uint64_t parallelBandwidthCandidateUs                    = 0;
    uint64_t parallelBandwidthBaselineMaxSkewBytes           = 0;
    uint64_t parallelBandwidthCandidateMaxSkewBytes          = 0;
    uint64_t parallelBandwidthBaselineMaxCallbackDeltaBytes  = 0;
    uint64_t parallelBandwidthCandidateMaxCallbackDeltaBytes = 0;
    size_t parallelBandwidthBaselineMaxActive                = 0;
    size_t parallelBandwidthCandidateMaxActive               = 0;
    uint64_t parallelBandwidthBaselineSamples                = 0;
    uint64_t parallelBandwidthCandidateSamples               = 0;
    ULONGLONG defaultSpeedLimitRunStartTick                  = 0;
    uint64_t defaultSpeedLimitBaselineUs                     = 0;
    uint64_t defaultSpeedLimitCandidateUs                    = 0;
    std::string defaultSpeedLimitDummyConfigSnapshot;
    ULONGLONG recycleBinBatchRunStartTick                   = 0;
    uint64_t recycleBinBatchBaselineUs                      = 0;
    uint64_t recycleBinBatchCandidateUs                     = 0;
    ULONGLONG recycleResumeScaleStartTick                   = 0;
    size_t recycleEscalationStopPrompted                    = 0;
    size_t recycleEscalationStopReplaced                    = 0;
    ULONGLONG copyMoveConcurrencyPerfRunStartTick           = 0;
    uint64_t copyMoveConcurrencyPerfBaselineUs              = 0;
    uint64_t copyMoveConcurrencyPerfCandidateUs             = 0;
    unsigned int copyMoveConcurrencyPerfBaselineConfigured  = 0;
    unsigned int copyMoveConcurrencyPerfCandidateConfigured = 0;
    size_t copyMoveConcurrencyPerfBaselineMaxActive         = 0;
    size_t copyMoveConcurrencyPerfCandidateMaxActive        = 0;
    ULONGLONG autoConcurrencyRunStartTick                   = 0;
    uint64_t autoConcurrencyManualUs                        = 0;
    uint64_t autoConcurrencyAutoUs                          = 0;
    unsigned int autoConcurrencyManualConfigured            = 0;
    unsigned int autoConcurrencyAutoConfigured              = 0;
    unsigned int autoDeleteConfigured                       = 0;
    bool autoConcurrencyAutoPopupObserved                   = false;
    unsigned int autoConcurrencyAutoPopupResolved           = 0;
    unsigned int autoConcurrencyAutoPopupApplied            = 0;
    bool autoConcurrencyAutoCompletedPopupObserved          = false;
    unsigned int autoConcurrencyAutoCompletedPopupResolved  = 0;
    unsigned int autoConcurrencyAutoCompletedPopupApplied   = 0;
    bool autoDeletePopupObserved                            = false;
    unsigned int autoDeletePopupResolved                    = 0;
    unsigned int autoDeletePopupApplied                     = 0;
    bool autoDeleteCompletedPopupObserved                   = false;
    unsigned int autoDeleteCompletedPopupResolved           = 0;
    unsigned int autoDeleteCompletedPopupApplied            = 0;
    ULONGLONG bridgePipelineRunStartTick                    = 0;
    uint64_t bridgePipelineBaselineUs                       = 0;
    uint64_t bridgePipelineCandidateUs                      = 0;
    ULONGLONG connOverridePerfRunStartTick                  = 0;
    uint64_t connOverridePerfBaselineUs                     = 0;
    uint64_t connOverridePerfCandidateUs                    = 0;
    unsigned int connOverridePerfBaselineConfigured         = 0;
    unsigned int connOverridePerfCandidateConfigured        = 0;
    bool connOverridePerfCandidatePopupObserved             = false;
    unsigned int connOverridePerfCandidatePopupResolved     = 0;
    unsigned int connOverridePerfCandidatePopupApplied      = 0;
    std::string dummyConfigSnapshot;
    std::string connOverrideDummyConfigSnapshot;
    std::unordered_map<std::uint64_t, CompletedTaskInfo> completedTasks;

    // Phase 14 â€” UI lifetime guard regression.
    std::optional<std::uint64_t> phase14InfoTask;
    std::atomic<bool> phase14ShutdownDone{false};
};

SelfTestState& GetState() noexcept
{
    // Intentionally leak to avoid static destruction order issues on process exit:
    // the plugin manager may unload modules before this state releases COM pointers.
    static SelfTestState* state = []() noexcept { return new SelfTestState{}; }();
    return *state;
}

std::wstring_view StepToString(SelfTestState::Step step) noexcept
{
    switch (step)
    {
        case SelfTestState::Step::Idle: return L"Idle";
        case SelfTestState::Step::Setup: return L"Setup";
        case SelfTestState::Step::FileOps_CommonPresentationPerf: return L"FileOps_CommonPresentationPerf";
        case SelfTestState::Step::FileOps_IntegratedPresentationPerf: return L"FileOps_IntegratedPresentationPerf";
        case SelfTestState::Step::FileOps_VisualGallery: return L"FileOps_VisualGallery";
        case SelfTestState::Step::FileOps_VisualGalleryInteraction: return L"FileOps_VisualGalleryInteraction";
        case SelfTestState::Step::FileOps_QueuedAdmissionContracts: return L"FileOps_QueuedAdmissionContracts";
        case SelfTestState::Step::FileOps_SafetyRemediationContracts: return L"FileOps_SafetyRemediationContracts";
        case SelfTestState::Step::FileOps_ConflictMessageText: return L"FileOps_ConflictMessageText";
        case SelfTestState::Step::FileOps_ReviewRemediationContracts: return L"FileOps_ReviewRemediationContracts";
        case SelfTestState::Step::FileOps_RepeatDecisionContracts: return L"FileOps_RepeatDecisionContracts";
        case SelfTestState::Step::FileOps_SkipCountingPresentation: return L"FileOps_SkipCountingPresentation";
        case SelfTestState::Step::FileOps_PermanentDeleteCompletionProgress: return L"FileOps_PermanentDeleteCompletionProgress";
        case SelfTestState::Step::FileOps_AdmissionBaseline: return L"FileOps_AdmissionBaseline";
        case SelfTestState::Step::FileOps_OwnedProductionShutdown: return L"FileOps_OwnedProductionShutdown";
        case SelfTestState::Step::FileOps_CrossVolumeMoveAfterPartialPublicationStorage: return L"FileOps_CrossVolumeMoveAfterPartialPublicationStorage";
        case SelfTestState::Step::FileOps_RepeatedNonClipboardCopyStorage: return L"FileOps_RepeatedNonClipboardCopyStorage";
        case SelfTestState::Step::FileOps_RepeatedNonClipboardMoveStorage: return L"FileOps_RepeatedNonClipboardMoveStorage";
        case SelfTestState::Step::FileOps_TerminalFactsProjection: return L"FileOps_TerminalFactsProjection";
        case SelfTestState::Step::FileOps_ClearCompletedScope: return L"FileOps_ClearCompletedScope";
        case SelfTestState::Step::FileOps_PerItemCancelUndispatchedRoots: return L"FileOps_PerItemCancelUndispatchedRoots";
        case SelfTestState::Step::FileOps_AdmissionFailureReachesPane: return L"FileOps_AdmissionFailureReachesPane";
        case SelfTestState::Step::FileOps_DeferredExitRefusesNewWork: return L"FileOps_DeferredExitRefusesNewWork";
        case SelfTestState::Step::FileOps_CompareWindowCloseConsent: return L"FileOps_CompareWindowCloseConsent";
        case SelfTestState::Step::FileOps_PopupFailureSurfaceRecovery: return L"FileOps_PopupFailureSurfaceRecovery";
        case SelfTestState::Step::FileOps_PopupStableHostedTree: return L"FileOps_PopupStableHostedTree";
        case SelfTestState::Step::FileOps_PopupIdleCadence: return L"FileOps_PopupIdleCadence";
        case SelfTestState::Step::FileOps_AlertOverlayLifetime: return L"FileOps_AlertOverlayLifetime";
        case SelfTestState::Step::FileOps_AlertOverlayUxAccessibility: return L"FileOps_AlertOverlayUxAccessibility";
        case SelfTestState::Step::FileOps_AlertOverlayModelessDecision: return L"FileOps_AlertOverlayModelessDecision";
        case SelfTestState::Step::FileOps_DecisionMirror: return L"FileOps_DecisionMirror";
        case SelfTestState::Step::FileOps_DecisionMirrorSameHostOverlap: return L"FileOps_DecisionMirrorSameHostOverlap";
        case SelfTestState::Step::FileOps_DecisionMirrorIngressAnchor: return L"FileOps_DecisionMirrorIngressAnchor";
        case SelfTestState::Step::FileOps_DecisionMirrorCompareWindow: return L"FileOps_DecisionMirrorCompareWindow";
        case SelfTestState::Step::FileOps_DecisionMirrorConsecutivePrompts: return L"FileOps_DecisionMirrorConsecutivePrompts";
        case SelfTestState::Step::FileOps_DecisionMirrorQueueOrder: return L"FileOps_DecisionMirrorQueueOrder";
        case SelfTestState::Step::FileOps_BatchRenameExactResultCarrier: return L"FileOps_BatchRenameExactResultCarrier";
        case SelfTestState::Step::FileOps_BatchRenameNeverRunRowsNotAttempted: return L"FileOps_BatchRenameNeverRunRowsNotAttempted";
        case SelfTestState::Step::FileOps_NotesSeverity: return L"FileOps_NotesSeverity";
        case SelfTestState::Step::FileOps_NotesMetadata: return L"FileOps_NotesMetadata";
        case SelfTestState::Step::FileOps_CrossProviderMtpNewFile: return L"FileOps_CrossProviderMtpNewFile";
        case SelfTestState::Step::FileOps_VerifyAvailability: return L"FileOps_VerifyAvailability";
        case SelfTestState::Step::FileOps_SizelessSourceNote: return L"FileOps_SizelessSourceNote";
        case SelfTestState::Step::FileOps_CrossRootLocalCopy: return L"FileOps_CrossRootLocalCopy";
        case SelfTestState::Step::FileOps_ArchiveNestedCopyLocal: return L"FileOps_ArchiveNestedCopyLocal";
        case SelfTestState::Step::FileOps_ArchiveNestedCopyUnc: return L"FileOps_ArchiveNestedCopyUnc";
        case SelfTestState::Step::FileOps_CancelAllNeverStarted: return L"FileOps_CancelAllNeverStarted";
        case SelfTestState::Step::FileOps_SmbCopyThroughputVariance: return L"FileOps_SmbCopyThroughputVariance";
        case SelfTestState::Step::FileOps_LocalCopyThroughput: return L"FileOps_LocalCopyThroughput";
        case SelfTestState::Step::FileOps_SmbNativeRouteConcurrencyBudget: return L"FileOps_SmbNativeRouteConcurrencyBudget";
        case SelfTestState::Step::FileOps_SameProviderBridgeDestinationStorage: return L"FileOps_SameProviderBridgeDestinationStorage";
        case SelfTestState::Step::FileOps_SmbManagedMoveLeavesNoStage: return L"FileOps_SmbManagedMoveLeavesNoStage";
        case SelfTestState::Step::FileOps_RateWindowDisplay: return L"FileOps_RateWindowDisplay";
        case SelfTestState::Step::FileOps_ProgressPhaseAccounting: return L"FileOps_ProgressPhaseAccounting";
        case SelfTestState::Step::FileOps_StallLabelLayout: return L"FileOps_StallLabelLayout";
        case SelfTestState::Step::FileOps_SmallFileCopyCost: return L"FileOps_SmallFileCopyCost";
        case SelfTestState::Step::Delete_IdentityLessLocationByName: return L"Delete_IdentityLessLocationByName";
        case SelfTestState::Step::Delete_IdentityLessSharingAndInterlock: return L"Delete_IdentityLessSharingAndInterlock";
        case SelfTestState::Step::Delete_IdentityLessQueuedRootOverlap: return L"Delete_IdentityLessQueuedRootOverlap";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeOtherSpelling: return L"Delete_IdentityLessPinnedTreeOtherSpelling";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeArchiveWriter: return L"Delete_IdentityLessPinnedTreeArchiveWriter";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeCardPending: return L"Delete_IdentityLessPinnedTreeCardPending";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeConcurrentWriter: return L"Delete_IdentityLessPinnedTreeConcurrentWriter";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeBudget: return L"Delete_IdentityLessPinnedTreeBudget";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeChangedEntry: return L"Delete_IdentityLessPinnedTreeChangedEntry";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeEscalationPartial: return L"Delete_IdentityLessPinnedTreeEscalationPartial";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeLinkSwap: return L"Delete_IdentityLessPinnedTreeLinkSwap";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeSubfolderOpen: return L"Delete_IdentityLessPinnedTreeSubfolderOpen";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeRecycleOverBudget: return L"Delete_IdentityLessPinnedTreeRecycleOverBudget";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeHolders: return L"Delete_IdentityLessPinnedTreeHolders";
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeScale: return L"Delete_IdentityLessPinnedTreeScale";
        case SelfTestState::Step::Delete_IdentityLessFolderChangeAndHolders: return L"Delete_IdentityLessFolderChangeAndHolders";
        case SelfTestState::Step::Delete_IdentityLessOpenFailureAndArchiveOrigins: return L"Delete_IdentityLessOpenFailureAndArchiveOrigins";
        case SelfTestState::Step::FileOps_IdentityLessGuardPartition: return L"FileOps_IdentityLessGuardPartition";
        case SelfTestState::Step::FileOps_LeftoverCleanupWalkBounds: return L"FileOps_LeftoverCleanupWalkBounds";
        case SelfTestState::Step::FileOps_OwnedStageCleanupBudget: return L"FileOps_OwnedStageCleanupBudget";
        case SelfTestState::Step::Delete_NasFileIdInfoUnsupportedByName: return L"Delete_NasFileIdInfoUnsupportedByName";
        case SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder: return L"FileOps_CopyMergeIntoExistingFolder";
        case SelfTestState::Step::FileOps_MoveMergeIntoExistingFolderSameVolume: return L"FileOps_MoveMergeIntoExistingFolderSameVolume";
        case SelfTestState::Step::Beeline_RenameMergeSkipKeepsSourceFolder: return L"Beeline_RenameMergeSkipKeepsSourceFolder";
        case SelfTestState::Step::FileOps_RenameMergeEmptyFolderCompletes: return L"FileOps_RenameMergeEmptyFolderCompletes";
        case SelfTestState::Step::Riptide_MoveDistinctSameSizeFilePreservesSource: return L"Riptide_MoveDistinctSameSizeFilePreservesSource";
        case SelfTestState::Step::Riptide_ReparseNativeMoveRelocatesLinkObject: return L"Riptide_ReparseNativeMoveRelocatesLinkObject";
        case SelfTestState::Step::Fairstream_ManagedMovePreservesUncopiedNewFiles: return L"Fairstream_ManagedMovePreservesUncopiedNewFiles";
        case SelfTestState::Step::Fairstream_MoveMergeReadonlyDestinationFolder: return L"Fairstream_MoveMergeReadonlyDestinationFolder";
        case SelfTestState::Step::Fairstream_OverwriteGrantIsOneShot: return L"Fairstream_OverwriteGrantIsOneShot";
        case SelfTestState::Step::Fairstream_ConflictPromptsSerializedUnderParallelism: return L"Fairstream_ConflictPromptsSerializedUnderParallelism";
        case SelfTestState::Step::Fairstream_ReparseReplaceNonEmptyDirRequiresConsent: return L"Fairstream_ReparseReplaceNonEmptyDirRequiresConsent";
        case SelfTestState::Step::Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate: return L"Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate";
        case SelfTestState::Step::Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent: return L"Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent";
        case SelfTestState::Step::Riptide_ReparseMoveSourceNeverUsesDirectoryMerge: return L"Riptide_ReparseMoveSourceNeverUsesDirectoryMerge";
        case SelfTestState::Step::Fairstream_MovedTreeKeepsLiteralLinks: return L"Fairstream_MovedTreeKeepsLiteralLinks";
        case SelfTestState::Step::Fairstream_CrossFsConcurrentMoveUsesBridge: return L"Fairstream_CrossFsConcurrentMoveUsesBridge";
        case SelfTestState::Step::Fairstream_RerunCopyIdenticalCollisionPrompts: return L"Fairstream_RerunCopyIdenticalCollisionPrompts";
        case SelfTestState::Step::Fairstream_MoveSameSizeCollisionPrompts: return L"Fairstream_MoveSameSizeCollisionPrompts";
        case SelfTestState::Step::Fairstream_GraphBandsFairUnderParallelCopy: return L"Fairstream_GraphBandsFairUnderParallelCopy";
        case SelfTestState::Step::Fairstream_ManagedMergeMovePreservesChildren: return L"Fairstream_ManagedMergeMovePreservesChildren";
        case SelfTestState::Step::Fairstream_JunctionMergeTargetRequiresConsent: return L"Fairstream_JunctionMergeTargetRequiresConsent";
        case SelfTestState::Step::Fairstream_CopyIntoSelfAliasRejected: return L"Fairstream_CopyIntoSelfAliasRejected";
        case SelfTestState::Step::Fairstream_StorageKindProbed: return L"Fairstream_StorageKindProbed";
        case SelfTestState::Step::Fairstream_ParallelDeleteContinuesPastLockedChild: return L"Fairstream_ParallelDeleteContinuesPastLockedChild";
        case SelfTestState::Step::Fairstream_BridgePerFileConflictSkips: return L"Fairstream_BridgePerFileConflictSkips";
        case SelfTestState::Step::Cinderstar_BridgeMovePerFileConflictSkipsParallel: return L"Cinderstar_BridgeMovePerFileConflictSkipsParallel";
        case SelfTestState::Step::Riptide_BridgeNestedDirVsFileSkipContinuesSiblings: return L"Riptide_BridgeNestedDirVsFileSkipContinuesSiblings";
        case SelfTestState::Step::Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch: return L"Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch";
        case SelfTestState::Step::Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial:
            return L"Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial";
        case SelfTestState::Step::Fairstream_SaturationConcurrentCopiesMakeProgress: return L"Fairstream_SaturationConcurrentCopiesMakeProgress";
        case SelfTestState::Step::Fairstream_DiscoveryAheadOverlapsTransfer: return L"Fairstream_DiscoveryAheadOverlapsTransfer";
        case SelfTestState::Step::Riptide_LiveFinishedSnapshotCarriesDiagnostics: return L"Riptide_LiveFinishedSnapshotCarriesDiagnostics";
        case SelfTestState::Step::Riptide_BridgeSequentialContinueOnErrorCopiesSiblings: return L"Riptide_BridgeSequentialContinueOnErrorCopiesSiblings";
        case SelfTestState::Step::Causeway_BridgeRejectsHostileChildNames: return L"Causeway_BridgeRejectsHostileChildNames";
        case SelfTestState::Step::Causeway_BridgeProviderOutputContracts: return L"Causeway_BridgeProviderOutputContracts";
        case SelfTestState::Step::Causeway_BridgeSchedulingAndResourceContracts: return L"Causeway_BridgeSchedulingAndResourceContracts";
        case SelfTestState::Step::Causeway_BridgeFileReparsePolicy: return L"Causeway_BridgeFileReparsePolicy";
        case SelfTestState::Step::Beeline_CopySkipLinksKeepsPlaceholders: return L"Beeline_CopySkipLinksKeepsPlaceholders";
        case SelfTestState::Step::Causeway_BridgeFailureStatusAndPausedReader: return L"Causeway_BridgeFailureStatusAndPausedReader";
        case SelfTestState::Step::Floodgate_CrossFsCopyGetSizeFailureRefusesCommit: return L"Floodgate_CrossFsCopyGetSizeFailureRefusesCommit";
        case SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailurePreservesSource: return L"Floodgate_CrossFsMoveGetSizeFailurePreservesSource";
        case SelfTestState::Step::Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource:
            return L"Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource";
        case SelfTestState::Step::Cinderstar_LegacyWriterEventuallyConsistentMove: return L"Cinderstar_LegacyWriterEventuallyConsistentMove";
        case SelfTestState::Step::Cinderstar_LegacyWriterPermanentMissMove: return L"Cinderstar_LegacyWriterPermanentMissMove";
        case SelfTestState::Step::Cinderstar_LegacyWriterWrongSizeMove: return L"Cinderstar_LegacyWriterWrongSizeMove";
        case SelfTestState::Step::Cinderstar_LegacyWriterConcurrentReplacementCopy: return L"Cinderstar_LegacyWriterConcurrentReplacementCopy";
        case SelfTestState::Step::Cinderstar_LegacyWriterCancelDuringBackoff: return L"Cinderstar_LegacyWriterCancelDuringBackoff";
        case SelfTestState::Step::Floodgate_CrossFsCopyOnlyUsesCopyVerification: return L"Floodgate_CrossFsCopyOnlyUsesCopyVerification";
        case SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree:
            return L"Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree";
        case SelfTestState::Step::Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup: return L"Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup";
        case SelfTestState::Step::Floodgate_CrossFsDirectoryCopyOnlyRetainsSource: return L"Floodgate_CrossFsDirectoryCopyOnlyRetainsSource";
        case SelfTestState::Step::Floodgate_LocalWriterOverwriteIsStaged: return L"Floodgate_LocalWriterOverwriteIsStaged";
        case SelfTestState::Step::Floodgate_LocalCopyOverwriteIsStaged: return L"Floodgate_LocalCopyOverwriteIsStaged";
        case SelfTestState::Step::Floodgate_LocalCopyNewNameConcurrentReplacementSurvives: return L"Floodgate_LocalCopyNewNameConcurrentReplacementSurvives";
        case SelfTestState::Step::Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete:
            return L"Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete";
        case SelfTestState::Step::Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker:
            return L"Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker";
        case SelfTestState::Step::Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker:
            return L"Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker";
        case SelfTestState::Step::Floodgate_InlineF2WorkerQueueBypassAndInterlock: return L"Floodgate_InlineF2WorkerQueueBypassAndInterlock";
        case SelfTestState::Step::C1_InlineRenameNameRefusedOnCard: return L"C1_InlineRenameNameRefusedOnCard";
        case SelfTestState::Step::R4A19_DiscoveryMeasurementFacts: return L"R4A19_DiscoveryMeasurementFacts";
        case SelfTestState::Step::R4A19_DiscoveryProviderControls: return L"R4A19_DiscoveryProviderControls";
        case SelfTestState::Step::R4A19_DiscoveryIndependentVolumes: return L"R4A19_DiscoveryIndependentVolumes";
        case SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns: return L"R0fSmb_BlockedSynchronousCallCancelReturns";
        case SelfTestState::Step::R0fSmb_LoopbackReadWriteCreateDelete: return L"R0fSmb_LoopbackReadWriteCreateDelete";
        case SelfTestState::Step::R0fSmb_LoopbackRecycleNeverDestroys: return L"R0fSmb_LoopbackRecycleNeverDestroys";
        case SelfTestState::Step::R0fSmb_NestedTreeCopyToShare: return L"R0fSmb_NestedTreeCopyToShare";
        case SelfTestState::Step::R0fSmb_BridgeTreeCopyToShare: return L"R0fSmb_BridgeTreeCopyToShare";
        case SelfTestState::Step::R0fSmb_ManagedMoveTreeToShare: return L"R0fSmb_ManagedMoveTreeToShare";
        case SelfTestState::Step::R0fCurl_FakeFtpReadWriteCreateDelete: return L"R0fCurl_FakeFtpReadWriteCreateDelete";
        case SelfTestState::Step::BR5_CurlHostDirectoryMove: return L"BR5_CurlHostDirectoryMove";
        case SelfTestState::Step::BR5_CurlHostDirectoryMoveRefused: return L"BR5_CurlHostDirectoryMoveRefused";
        case SelfTestState::Step::R0fS3_FakeS3ReadWriteCreateDelete: return L"R0fS3_FakeS3ReadWriteCreateDelete";
        case SelfTestState::Step::R0fGraph_FakeGraphReadWriteCreateRenameRecycle: return L"R0fGraph_FakeGraphReadWriteCreateRenameRecycle";
        case SelfTestState::Step::R0fGDrive_FakeDriveReadWriteCreateMoveDelete: return L"R0fGDrive_FakeDriveReadWriteCreateMoveDelete";
        case SelfTestState::Step::C0_GDriveCommittedCopyResponseLost: return L"C0_GDriveCommittedCopyResponseLost";
        case SelfTestState::Step::C0_GDriveCommittedDeleteResponseLost: return L"C0_GDriveCommittedDeleteResponseLost";
        case SelfTestState::Step::C0_NativeCopyFailurePreservesKnownAxes: return L"C0_NativeCopyFailurePreservesKnownAxes";
        case SelfTestState::Step::C0_MutationReceiptPrefixBoundary: return L"C0_MutationReceiptPrefixBoundary";
        case SelfTestState::Step::C0_LocalKnownNoCommitReceipts: return L"C0_LocalKnownNoCommitReceipts";
        case SelfTestState::Step::C0_CurlNativeDeleteGuards: return L"C0_CurlNativeDeleteGuards";
        case SelfTestState::Step::C0_CurlCommittedMutationResponseLost: return L"C0_CurlCommittedMutationResponseLost";
        case SelfTestState::Step::C0_CurlHostCommittedDeleteResponseLost: return L"C0_CurlHostCommittedDeleteResponseLost";
        case SelfTestState::Step::C0_CurlPartialTreeFailure: return L"C0_CurlPartialTreeFailure";
        case SelfTestState::Step::C0_CurlHostPartialTreeFailure: return L"C0_CurlHostPartialTreeFailure";
        case SelfTestState::Step::C0_CurlNativeDeleteLateListing: return L"C0_CurlNativeDeleteLateListing";
        case SelfTestState::Step::C0_CurlDeleteTraversalBounds: return L"C0_CurlDeleteTraversalBounds";
        case SelfTestState::Step::C0_CurlDirectorySizeTruth: return L"C0_CurlDirectorySizeTruth";
        case SelfTestState::Step::C0_CurlMovePreflightTruth: return L"C0_CurlMovePreflightTruth";
        case SelfTestState::Step::C0_CurlCopyTraversalTruth: return L"C0_CurlCopyTraversalTruth";
        case SelfTestState::Step::C0_CurlEntryLookupTruth: return L"C0_CurlEntryLookupTruth";
        case SelfTestState::Step::C0_CurlImapListingTruth: return L"C0_CurlImapListingTruth";
        case SelfTestState::Step::C0_CurlImapTransportTruth: return L"C0_CurlImapTransportTruth";
        case SelfTestState::Step::R3_1_IdentityLessReplaceDummy: return L"R3_1_IdentityLessReplaceDummy";
        case SelfTestState::Step::R3_2_WriterProofDummy: return L"R3_2_WriterProofDummy";
        case SelfTestState::Step::R3_3_PublicationFaultMatrix: return L"R3_3_PublicationFaultMatrix";
        case SelfTestState::Step::R3_4_ReplaceWithoutOccupantTokenRefused: return L"R3_4_ReplaceWithoutOccupantTokenRefused";
        case SelfTestState::Step::RC3_9_TransferDestinationNameRefused: return L"RC3_9_TransferDestinationNameRefused";
        case SelfTestState::Step::R4T_DeepTreeCopyCompletesIteratively: return L"R4T_DeepTreeCopyCompletesIteratively";
        case SelfTestState::Step::R4T2_DeepTreeLocalDeleteCompletes: return L"R4T2_DeepTreeLocalDeleteCompletes";
        case SelfTestState::Step::R4T3_WideDirectoryCopyCompletes: return L"R4T3_WideDirectoryCopyCompletes";
        case SelfTestState::Step::R4T4_LocalManyDirectoriesCopy: return L"R4T4_LocalManyDirectoriesCopy";
        case SelfTestState::Step::R4T5_LocalCopyHoldsNoDirectoryHandle: return L"R4T5_LocalCopyHoldsNoDirectoryHandle";
        case SelfTestState::Step::R4T6_LocalDirectoryRestoreChecksIdentity: return L"R4T6_LocalDirectoryRestoreChecksIdentity";
        case SelfTestState::Step::R4T7_BridgeCopyHoldsNoDirectoryHandle: return L"R4T7_BridgeCopyHoldsNoDirectoryHandle";
        case SelfTestState::Step::R4T8_BridgeDirectoryRestoreChecksIdentity: return L"R4T8_BridgeDirectoryRestoreChecksIdentity";
        case SelfTestState::Step::R4T9_BridgeManyDirectoriesCopy: return L"R4T9_BridgeManyDirectoriesCopy";
        case SelfTestState::Step::R4A02_DeleteOverCopySourceWarnsAndQueues: return L"R4A02_DeleteOverCopySourceWarnsAndQueues";
        case SelfTestState::Step::R4A02_DontStartCancelsBeforeMutation: return L"R4A02_DontStartCancelsBeforeMutation";
        case SelfTestState::Step::R4A02_ReadReadDoesNotWarn: return L"R4A02_ReadReadDoesNotWarn";
        case SelfTestState::Step::R4A02_SameDestinationWarns: return L"R4A02_SameDestinationWarns";
        case SelfTestState::Step::R4A02_RunConcurrentSameDestinationStaysSafe: return L"R4A02_RunConcurrentSameDestinationStaysSafe";
        case SelfTestState::Step::R4A02_LiveOutputGuardChoices: return L"R4A02_LiveOutputGuardChoices";
        case SelfTestState::Step::R4A02_DisjointWritesDoNotWarn: return L"R4A02_DisjointWritesDoNotWarn";
        case SelfTestState::Step::R4A02_RenamePublishKeepsIndexPrecise: return L"R4A02_RenamePublishKeepsIndexPrecise";
        case SelfTestState::Step::R4A02_LatePublisherStillWarns: return L"R4A02_LatePublisherStillWarns";
        case SelfTestState::Step::FileOps_ReparseDirectoryMergeIntoExistingFolder: return L"FileOps_ReparseDirectoryMergeIntoExistingFolder";
        case SelfTestState::Step::FileOps_ProviderCapabilityMatrix: return L"FileOps_ProviderCapabilityMatrix";
        case SelfTestState::Step::FileOps_ParallelGraphFairColorWeight: return L"FileOps_ParallelGraphFairColorWeight";
        case SelfTestState::Step::FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof: return L"FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof";
        case SelfTestState::Step::FileOps_ReparseLiteralCopyKeepsTargets: return L"FileOps_ReparseLiteralCopyKeepsTargets";
        case SelfTestState::Step::FileOps_DeleteToctouSwapGuard: return L"FileOps_DeleteToctouSwapGuard";
        case SelfTestState::Step::FileOps_NativeMoveRacedDestinationNotReplaced: return L"FileOps_NativeMoveRacedDestinationNotReplaced";
        case SelfTestState::Step::FileOps_NativeMoveMissingPathKnownFailure: return L"FileOps_NativeMoveMissingPathKnownFailure";
        case SelfTestState::Step::FileOps_ResolvedItemsExactDestinations: return L"FileOps_ResolvedItemsExactDestinations";
        case SelfTestState::Step::Phase5_DiscoverySingleTraversal: return L"Phase5_DiscoverySingleTraversal";
        case SelfTestState::Step::Phase5_DiscoveryCancelReleasesSlot: return L"Phase5_DiscoveryCancelReleasesSlot";
        case SelfTestState::Step::Phase5_CleanupDrainsOutstandingTasks: return L"Phase5_CleanupDrainsOutstandingTasks";
        case SelfTestState::Step::Phase5_DiscoveryCancelLatencyLocal: return L"Phase5_DiscoveryCancelLatencyLocal";
        case SelfTestState::Step::Phase5_DiscoverySkipContinues: return L"Phase5_DiscoverySkipContinues";
        case SelfTestState::Step::Phase5_CancelQueuedTask: return L"Phase5_CancelQueuedTask";
        case SelfTestState::Step::BR3_CancelQueuePausedTransfer: return L"BR3_CancelQueuePausedTransfer";
        case SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery: return L"Phase5_SwitchParallelToWaitDuringDiscovery";
        case SelfTestState::Step::Phase5_SwitchWaitToParallelResume: return L"Phase5_SwitchWaitToParallelResume";
        case SelfTestState::Step::Phase6_PopupRateSmoothing: return L"Phase6_PopupRateSmoothing";
        case SelfTestState::Step::Phase6_PopupSmokeResizeAndPause: return L"Phase6_PopupSmokeResizeAndPause";
        case SelfTestState::Step::Phase6_DeleteBytesMeaningful: return L"Phase6_DeleteBytesMeaningful";
        case SelfTestState::Step::Phase6_LocalBandwidthThrottle: return L"Phase6_LocalBandwidthThrottle";
        case SelfTestState::Step::Phase6_ParallelBandwidthThrottleFairness: return L"Phase6_ParallelBandwidthThrottleFairness";
        case SelfTestState::Step::Phase7_WatcherChurn: return L"Phase7_WatcherChurn";
        case SelfTestState::Step::Phase7_CacheBorrowNoWatchInvalidation: return L"Phase7_CacheBorrowNoWatchInvalidation";
        case SelfTestState::Step::Phase7_CrossPaneVisibleRefreshLocal: return L"Phase7_CrossPaneVisibleRefreshLocal";
        case SelfTestState::Step::Phase7_CrossPaneVisibleRefreshDummy: return L"Phase7_CrossPaneVisibleRefreshDummy";
        case SelfTestState::Step::Phase7_CrossPaneRelocateLocal: return L"Phase7_CrossPaneRelocateLocal";
        case SelfTestState::Step::Phase7_LargeDirectoryEnumeration: return L"Phase7_LargeDirectoryEnumeration";
        case SelfTestState::Step::Phase7_ParallelCopyMoveKnobs: return L"Phase7_ParallelCopyMoveKnobs";
        case SelfTestState::Step::Phase7_CopyMoveConcurrency16Perf: return L"Phase7_CopyMoveConcurrency16Perf";
        case SelfTestState::Step::Phase7_AutoConcurrencyHints: return L"Phase7_AutoConcurrencyHints";
        case SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines: return L"Phase7_PerItemDirectoryCopyInFlightLines";
        case SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism: return L"Phase7_CopyItemsSingleFolderRecursiveParallelism";
        case SelfTestState::Step::Phase7_CopyItemsMultiRootUnevenRecursiveParallelism: return L"Phase7_CopyItemsMultiRootUnevenRecursiveParallelism";
        case SelfTestState::Step::Phase7_CopyRecursiveParallelismMatrix: return L"Phase7_CopyRecursiveParallelismMatrix";
        case SelfTestState::Step::Phase7_CopyItemsSaturatedPoolProducersHelp: return L"Phase7_CopyItemsSaturatedPoolProducersHelp";
        case SelfTestState::Step::Phase7_CopyItemWaitingProducerSeesCancel: return L"Phase7_CopyItemWaitingProducerSeesCancel";
        case SelfTestState::Step::Phase7_CopyItemWalkAtLimitParks: return L"Phase7_CopyItemWalkAtLimitParks";
        case SelfTestState::Step::Phase7_DeleteItemsNestedRootsWake: return L"Phase7_DeleteItemsNestedRootsWake";
        case SelfTestState::Step::Phase7_SharedPerItemScheduler: return L"Phase7_SharedPerItemScheduler";
        case SelfTestState::Step::Phase7_ParallelDeleteKnobs: return L"Phase7_ParallelDeleteKnobs";
        case SelfTestState::Step::Phase7_RecycleBinBatchDelete: return L"Phase7_RecycleBinBatchDelete";
        case SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch: return L"Phase7_RecycleBinBatchDeleteMultiBatch";
        case SelfTestState::Step::Phase7_RecycleBinTestItemsPurged: return L"Phase7_RecycleBinTestItemsPurged";
        case SelfTestState::Step::Phase8_DefaultBandwidthLimitFromSettings: return L"Phase8_DefaultBandwidthLimitFromSettings";
        case SelfTestState::Step::Phase8_TightDefaults_NoOverwrite: return L"Phase8_TightDefaults_NoOverwrite";
        case SelfTestState::Step::Phase8_InvalidDestinationRejected: return L"Phase8_InvalidDestinationRejected";
        case SelfTestState::Step::Phase8_InvalidSizeBytesRejected: return L"Phase8_InvalidSizeBytesRejected";
        case SelfTestState::Step::Phase8_PerItemOrchestration: return L"Phase8_PerItemOrchestration";
        case SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly: return L"Phase9_ConflictPrompt_OverwriteReplaceReadonly";
        case SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache: return L"Phase9_ConflictPrompt_ApplyToAllUiCache";
        case SelfTestState::Step::Phase9_ConflictPrompt_KeepBothNestedCacheEligibility: return L"Phase9_ConflictPrompt_KeepBothNestedCacheEligibility";
        case SelfTestState::Step::Phase9_ConflictPrompt_TypeMismatchNoOverwrite: return L"Phase9_ConflictPrompt_TypeMismatchNoOverwrite";
        case SelfTestState::Step::Phase9_ConflictPrompt_LocalFileOntoDirectory: return L"Phase9_ConflictPrompt_LocalFileOntoDirectory";
        case SelfTestState::Step::Phase9_ConflictPrompt_SkipApplyToAll: return L"Phase9_ConflictPrompt_SkipApplyToAll";
        case SelfTestState::Step::Phase9_ConflictPrompt_RetryCap: return L"Phase9_ConflictPrompt_RetryCap";
        case SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy: return L"Phase9_ConflictPrompt_SkipContinuesDirectoryCopy";
        case SelfTestState::Step::Phase9_PerItemConcurrency: return L"Phase9_PerItemConcurrency";
        case SelfTestState::Step::Phase10_PermanentDelete: return L"Phase10_PermanentDelete";
        case SelfTestState::Step::Phase10_DeferredConsentAndRecycleEscalation: return L"Phase10_DeferredConsentAndRecycleEscalation";
        case SelfTestState::Step::Phase10_RecycleDestroyedIsNeverRecycled: return L"Phase10_RecycleDestroyedIsNeverRecycled";
        case SelfTestState::Step::Phase10_RecycleEscalationDirectoryTree: return L"Phase10_RecycleEscalationDirectoryTree";
        case SelfTestState::Step::Phase10_RecycleEscalationTaskScopedGrant: return L"Phase10_RecycleEscalationTaskScopedGrant";
        case SelfTestState::Step::Phase10_RecycleEscalationBulkItemsContinueOnError: return L"Phase10_RecycleEscalationBulkItemsContinueOnError";
        case SelfTestState::Step::Phase10_RecycleEscalationBulkItemsResumeBounded: return L"Phase10_RecycleEscalationBulkItemsResumeBounded";
        case SelfTestState::Step::Phase10_MetadataPreservationAndSourceRetention: return L"Phase10_MetadataPreservationAndSourceRetention";
        case SelfTestState::Step::Phase10_TypedResultsAndConsumers: return L"Phase10_TypedResultsAndConsumers";
        case SelfTestState::Step::Phase10_ClipboardAdmissionAndRetainedActions: return L"Phase10_ClipboardAdmissionAndRetainedActions";
        case SelfTestState::Step::Phase10_ContentVerification: return L"Phase10_ContentVerification";
        case SelfTestState::Step::Phase10_ArtifactTouchGuard: return L"Phase10_ArtifactTouchGuard";
        case SelfTestState::Step::R1d_PreparingLifecycle: return L"R1d_PreparingLifecycle";
        case SelfTestState::Step::C1_ExitCloseDeferredUntilTasksQuiet: return L"C1_ExitCloseDeferredUntilTasksQuiet";
        case SelfTestState::Step::Beeline_SameVolumeTreeMoveIsRename: return L"Beeline_SameVolumeTreeMoveIsRename";
        case SelfTestState::Step::Beeline_LoopbackShareMoveIsRename: return L"Beeline_LoopbackShareMoveIsRename";
        case SelfTestState::Step::C1_PermanentDeleteConfirmsOnCard: return L"C1_PermanentDeleteConfirmsOnCard";
        case SelfTestState::Step::Phase11_CrossFileSystemBridge: return L"Phase11_CrossFileSystemBridge";
        case SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines: return L"Phase11_BridgeSingleFolderParallelCopyInFlightLines";
        case SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines: return L"Phase11_BridgeMultiFolderParallelCopyInFlightLines";
        case SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf: return L"Phase11_BridgePipelineDummyToDummyPerf";
        case SelfTestState::Step::Phase11_ConnectionOverridePrecedence: return L"Phase11_ConnectionOverridePrecedence";
        case SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate: return L"Phase11_ConnectionOverrideGlobalGate";
        case SelfTestState::Step::Phase11_ConnectionOverrideClamp: return L"Phase11_ConnectionOverrideClamp";
        case SelfTestState::Step::Phase12_ReparsePointPolicy: return L"Phase12_ReparsePointPolicy";
        case SelfTestState::Step::Phase13_PostMortemDiagnostics: return L"Phase13_PostMortemDiagnostics";
        case SelfTestState::Step::Phase14_PopupHostLifetimeGuard: return L"Phase14_PopupHostLifetimeGuard";
        case SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke: return L"Phase15_FileSystem7zReadSeekSmoke";
        case SelfTestState::Step::Phase15_FileSystem7zMountPathImpact: return L"Phase15_FileSystem7zMountPathImpact";
        case SelfTestState::Step::Phase16_RemoteWatchContractExposure: return L"Phase16_RemoteWatchContractExposure";
        case SelfTestState::Step::Phase16_RemoteFtpSecret: return L"Phase16_RemoteFtpSecret";
        case SelfTestState::Step::Phase16_RemoteFtpSandbox: return L"Phase16_RemoteFtpSandbox";
        case SelfTestState::Step::Phase16_RemoteSftpSecret: return L"Phase16_RemoteSftpSecret";
        case SelfTestState::Step::Phase16_RemoteSftpSandbox: return L"Phase16_RemoteSftpSandbox";
        case SelfTestState::Step::Phase16_RemoteScpSecret: return L"Phase16_RemoteScpSecret";
        case SelfTestState::Step::Phase16_RemoteScpSandbox: return L"Phase16_RemoteScpSandbox";
        case SelfTestState::Step::Phase16_RemoteImapSecret: return L"Phase16_RemoteImapSecret";
        case SelfTestState::Step::Phase16_RemoteImapSandbox: return L"Phase16_RemoteImapSandbox";
        case SelfTestState::Step::Phase16_RemoteS3Secret: return L"Phase16_RemoteS3Secret";
        case SelfTestState::Step::Phase16_RemoteS3Sandbox: return L"Phase16_RemoteS3Sandbox";
        case SelfTestState::Step::Phase16_RemoteS3FileOps: return L"Phase16_RemoteS3FileOps";
        case SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret: return L"Phase16_RemoteOneDrivePersonalSecret";
        case SelfTestState::Step::Phase16_RemoteOneDrivePersonalSandbox: return L"Phase16_RemoteOneDrivePersonalSandbox";
        case SelfTestState::Step::Phase16_RemoteOneDrivePersonalFileOps: return L"Phase16_RemoteOneDrivePersonalFileOps";
        case SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret: return L"Phase16_RemoteOneDriveBusinessSecret";
        case SelfTestState::Step::Phase16_RemoteOneDriveBusinessSandbox: return L"Phase16_RemoteOneDriveBusinessSandbox";
        case SelfTestState::Step::Phase16_RemoteSharePointSecret: return L"Phase16_RemoteSharePointSecret";
        case SelfTestState::Step::Phase16_RemoteSharePointSandbox: return L"Phase16_RemoteSharePointSandbox";
        case SelfTestState::Step::DiscoveryScope_LeafCopy: return L"DiscoveryScope_LeafCopy";
        case SelfTestState::Step::DiscoveryScope_LeafMove: return L"DiscoveryScope_LeafMove";
        case SelfTestState::Step::DiscoveryScope_ManagedDirectory: return L"DiscoveryScope_ManagedDirectory";
        case SelfTestState::Step::DiscoveryScope_NativeRenameMerge: return L"DiscoveryScope_NativeRenameMerge";
        case SelfTestState::Step::DiscoveryScope_MixedRoots: return L"DiscoveryScope_MixedRoots";
        case SelfTestState::Step::DiscoveryScope_ZeroByteAndVerifiedLeaf: return L"DiscoveryScope_ZeroByteAndVerifiedLeaf";
        case SelfTestState::Step::DiscoveryScope_LocalDirectApi: return L"DiscoveryScope_LocalDirectApi";
        case SelfTestState::Step::DiscoveryScope_DummyDirectApi: return L"DiscoveryScope_DummyDirectApi";
        case SelfTestState::Step::DiscoveryScope_CurlDirectApi: return L"DiscoveryScope_CurlDirectApi";
        case SelfTestState::Step::DiscoveryScope_S3DirectApi: return L"DiscoveryScope_S3DirectApi";
        case SelfTestState::Step::DiscoveryScope_GDriveDirectApi: return L"DiscoveryScope_GDriveDirectApi";
        case SelfTestState::Step::DiscoveryScope_GraphDirectApi: return L"DiscoveryScope_GraphDirectApi";
        case SelfTestState::Step::DiscoveryScope_MtpDirectApi: return L"DiscoveryScope_MtpDirectApi";
        case SelfTestState::Step::DiscoveryScope_FolderClosesBeforeLastFile: return L"DiscoveryScope_FolderClosesBeforeLastFile";
        case SelfTestState::Step::DiscoveryScope_WideFolderClosesWithinQueueLead: return L"DiscoveryScope_WideFolderClosesWithinQueueLead";
        case SelfTestState::Step::DiscoveryScope_NestedTreeMetadataRestoredAfterChildren: return L"DiscoveryScope_NestedTreeMetadataRestoredAfterChildren";
        case SelfTestState::Step::DiscoveryScope_ManagedMoveListingCompleteBeforeRemoval: return L"DiscoveryScope_ManagedMoveListingCompleteBeforeRemoval";
        case SelfTestState::Step::DiscoveryScope_WorkerFailureIsNotCanceled: return L"DiscoveryScope_WorkerFailureIsNotCanceled";
        case SelfTestState::Step::DiscoveryScope_FlatFolderUsesFullConcurrency: return L"DiscoveryScope_FlatFolderUsesFullConcurrency";
        case SelfTestState::Step::DiscoveryScope_ListingSilenceNoted: return L"DiscoveryScope_ListingSilenceNoted";
        case SelfTestState::Step::DiscoveryScope_TenThousandSmallFilesCloseAfterListing: return L"DiscoveryScope_TenThousandSmallFilesCloseAfterListing";
        case SelfTestState::Step::DiscoveryScope_RetryDoesNotGrowAfterClose: return L"DiscoveryScope_RetryDoesNotGrowAfterClose";
        case SelfTestState::Step::DiscoveryScope_HeldFolderFailureStillCloses: return L"DiscoveryScope_HeldFolderFailureStillCloses";
        case SelfTestState::Step::DiscoveryScope_FolderLiveProgress: return L"DiscoveryScope_FolderLiveProgress";
        case SelfTestState::Step::DiscoveryScope_MultipleFoldersLiveProgress: return L"DiscoveryScope_MultipleFoldersLiveProgress";
        case SelfTestState::Step::FileOps_SelectionShapeLocal: return L"FileOps_SelectionShapeLocal";
        case SelfTestState::Step::FileOps_SelectionShapeSmb: return L"FileOps_SelectionShapeSmb";
        case SelfTestState::Step::Cleanup_RestorePluginConfig: return L"Cleanup_RestorePluginConfig";
        case SelfTestState::Step::Done: return L"Done";
        case SelfTestState::Step::Failed: return L"Failed";
        default: break;
    }
    return L"(unknown)";
}

constexpr auto kFileOpsPhaseOrder = std::to_array<SelfTestState::Step>({
    SelfTestState::Step::Setup,                            // Environment setup and plugin loading
    SelfTestState::Step::FileOps_QueuedAdmissionContracts, // Private canonical admission and clipboard release

    SelfTestState::Step::FileOps_SafetyRemediationContracts,
    SelfTestState::Step::FileOps_ConflictMessageText,
    SelfTestState::Step::FileOps_ReviewRemediationContracts,
    SelfTestState::Step::FileOps_RepeatDecisionContracts,
    SelfTestState::Step::FileOps_SkipCountingPresentation,
    SelfTestState::Step::FileOps_PermanentDeleteCompletionProgress,
    SelfTestState::Step::FileOps_AdmissionBaseline,
    SelfTestState::Step::FileOps_OwnedProductionShutdown,
    SelfTestState::Step::FileOps_CrossVolumeMoveAfterPartialPublicationStorage,
    SelfTestState::Step::FileOps_RepeatedNonClipboardCopyStorage,
    SelfTestState::Step::FileOps_RepeatedNonClipboardMoveStorage,
    SelfTestState::Step::FileOps_TerminalFactsProjection,               // Completed summary reaches the production popup with independent terminal counts
    SelfTestState::Step::FileOps_ClearCompletedScope,                   // Clear completed clears only painted results and keeps interrupted-Move breadcrumbs
    SelfTestState::Step::FileOps_PerItemCancelUndispatchedRoots,        // Cancel at a prompt leaves never-dispatched per-item roots untouched, not unknown
    SelfTestState::Step::FileOps_AdmissionFailureReachesPane,           // I30 host#2 - a failed pane Copy/Move/Delete admission reaches the pane alert
    SelfTestState::Step::FileOps_DeferredExitRefusesNewWork,            // I30 x-lifetime#4 (Q14) - a pending exit refuses new work and never re-prompts
    SelfTestState::Step::FileOps_CompareWindowCloseConsent,             // I30 a2-4 - Compare window close and the application exit cover Compare engines
    SelfTestState::Step::FileOps_PopupFailureSurfaceRecovery,           // I30 popup-a#1 (Q6) - popup frame/layout failures retry; failure surface recovers
    SelfTestState::Step::FileOps_PopupStableHostedTree,                 // I30 popup-b#1 (Q7) - hosted tree reconciled by identity; held press survives churn
    SelfTestState::Step::FileOps_PopupIdleCadence,                      // I30 x-perf#3/#5, x-lifetime#3/#7 - waiting tasks and idle UI stop per-tick work
    SelfTestState::Step::FileOps_AlertOverlayLifetime,                  // I30 alert#2/#9/#10, x-ux#12 - shared-anchor tracking; prompt exits when owner dies
    SelfTestState::Step::FileOps_AlertOverlayUxAccessibility,           // I30 alert#1/#3-#8 - no option jump, UIA events/marshal, mnemonics, names
    SelfTestState::Step::FileOps_AlertOverlayModelessDecision,          // I35 S5 R8.13 - modeless overlay: focus only when owned, close is not Escape
    SelfTestState::Step::FileOps_DecisionMirror,                        // I35 S5 R8.1-R8.9 - pane mirror of a card decision, first answer wins
    SelfTestState::Step::FileOps_DecisionMirrorSameHostOverlap,         // I35 S5 R8.2/R8.6 - overlap consent: Queue after default, Escape Don't start
    SelfTestState::Step::FileOps_DecisionMirrorIngressAnchor,           // I35 S5 R8.10 - pre-task prompts anchor on the originating pane
    SelfTestState::Step::FileOps_DecisionMirrorCompareWindow,           // I35 S5 R8.4/R8.8 - per-window mirrors; Compare close tears its mirror down
    SelfTestState::Step::FileOps_DecisionMirrorConsecutivePrompts,      // I35 S5 D-8.1 - one flash per run of questions; one Apply-to-all flag
    SelfTestState::Step::FileOps_DecisionMirrorQueueOrder,              // I35 S5 R8.6 - shown decision kept; the oldest pending decision next
    SelfTestState::Step::FileOps_BatchRenameExactResultCarrier,         // Post-preflight rename conflict reports exact accepted prompt and mutation result
    SelfTestState::Step::FileOps_BatchRenameNeverRunRowsNotAttempted,   // I30 state-c#1 - a Batch Rename row a stopped run never reached is NotAttempted
    SelfTestState::Step::FileOps_NotesSeverity,                         // I35 S1 - questions and user skips are notes, real failures still count
    SelfTestState::Step::FileOps_NotesMetadata,                         // I35 S1/S1b - metadata outcomes classify alike on native and bridge Copy
    SelfTestState::Step::FileOps_CrossProviderMtpNewFile,               // I35 XP-01 - a new file copied to an MTP device publishes through its writer
    SelfTestState::Step::FileOps_VerifyAvailability,                    // I35 XP-04 - Verify never removes a destination's only publication route
    SelfTestState::Step::FileOps_SizelessSourceNote,                    // I35 XP-10 - a sizeless source fails per file and is named once
    SelfTestState::Step::FileOps_CrossRootLocalCopy,                    // I35 XP-09 - a cross-root Local Copy runs the provider route
    SelfTestState::Step::FileOps_ArchiveNestedCopyLocal,                // Archive nested Windows joins retain only the leaf on local Copy.
    SelfTestState::Step::FileOps_ArchiveNestedCopyUnc,                  // The same real archive bridge through sandbox-only SMB.
    SelfTestState::Step::FileOps_CancelAllNeverStarted,                 // I35 S3 - Cancel all: never-started line, no counted problem, one-time group collapse
    SelfTestState::Step::FileOps_SmbCopyThroughputVariance,             // I35 S6 R10.5 - corpus to the SMB self-test share: report cadence, stalls, variation
    SelfTestState::Step::FileOps_LocalCopyThroughput,                   // I35 S6 R10.6 - the same corpus within the local root: the regression reference
    SelfTestState::Step::FileOps_SmbNativeRouteConcurrencyBudget,       // I35 S6 R10.5 - native route budget is min(source, destination) with candidate c
    SelfTestState::Step::FileOps_SameProviderBridgeDestinationStorage,  // A same-provider bridge route consults its destination's storage
    SelfTestState::Step::FileOps_SmbManagedMoveLeavesNoStage,           // I35 S6 R10.5 - Move to the share cancelled in each phase leaves no stage
    SelfTestState::Step::FileOps_RateWindowDisplay,                     // I35 S6 R10.2 - speed display: 5 s window, named stalls, stall band, restart
    SelfTestState::Step::FileOps_ProgressPhaseAccounting,               // I35 S6 - unsized reports name no phase; completions counted once across expiry
    SelfTestState::Step::FileOps_StallLabelLayout,                      // I35 S6 R10.2 - a stall name never moves the plot; the band never hatches labels
    SelfTestState::Step::FileOps_SmallFileCopyCost,                     // I35 S6 - small-file Copy wall time on the provider route and the bridge
    SelfTestState::Step::Delete_IdentityLessLocationByName,             // I35 R7.3 - identity-less Delete by name: consent, corroboration, leftovers, Recycle
    SelfTestState::Step::Delete_IdentityLessSharingAndInterlock,        // I35 R7.1/R7.3 - sharing failure gets Retry; Indeterminate and held paths refused
    SelfTestState::Step::Delete_IdentityLessQueuedRootOverlap,          // I35 R7.3 - a task admitted after a queued delete by name respects its root
    SelfTestState::Step::Delete_IdentityLessPinnedTreeOtherSpelling,    // I35 R7.3 - data moved in through another spelling survives the pinned tree
    SelfTestState::Step::Delete_IdentityLessPinnedTreeArchiveWriter,    // I35 R7.3 - an archive Pack wrote into the tree survives the delete
    SelfTestState::Step::Delete_IdentityLessPinnedTreeCardPending,      // I35 R7.3 - published before the card; a deep change refuses
    SelfTestState::Step::Delete_IdentityLessPinnedTreeConcurrentWriter, // I35 R7.3 - a file written after the listing survives; partial receipt
    SelfTestState::Step::Delete_IdentityLessPinnedTreeBudget,           // I35 R7.3 - a tree over the by-name budget is refused with its reason
    SelfTestState::Step::Delete_IdentityLessPinnedTreeChangedEntry,     // I35 R7.3 - a listed entry changed or gone after the listing stops the delete, nothing
                                                                        // deleted
    SelfTestState::Step::Delete_IdentityLessPinnedTreeEscalationPartial, // I35 R7.3 - an escalated delete by name reports its partial stop and receipt
    SelfTestState::Step::Delete_IdentityLessPinnedTreeLinkSwap,          // I35 R7.3 - a folder swapped for a junction never redirects the delete
    SelfTestState::Step::Delete_IdentityLessPinnedTreeSubfolderOpen,     // I35 R7.3 - a sub-folder the location refuses to open is an open failure
    SelfTestState::Step::Delete_IdentityLessPinnedTreeRecycleOverBudget, // I35 R7.3 - a Recycle over the by-name bounds keeps its reason
    SelfTestState::Step::Delete_IdentityLessPinnedTreeHolders,           // I35 R7.3 - a decided holder refuses the delete at both re-checks
    SelfTestState::Step::Delete_IdentityLessPinnedTreeScale,        // I35 R7.3 - a few thousand entries deleted by name, whole and in chunks (timings logged)
    SelfTestState::Step::Delete_IdentityLessFolderChangeAndHolders, // I35 R7.3 - folder child change and late holders refuse a delete by name
    SelfTestState::Step::Delete_IdentityLessOpenFailureAndArchiveOrigins, // I35 R7.1/R7.3 - a failed open and archive cleanups never delete by name
    SelfTestState::Step::FileOps_IdentityLessGuardPartition,              // I35 R7.3/R7.6 - unidentifiable Delete items leave the guard; other touches refused
    SelfTestState::Step::FileOps_LeftoverCleanupWalkBounds,               // I35 R7.5 - the leftover listing's walk keeps and reports its bounds
    SelfTestState::Step::FileOps_OwnedStageCleanupBudget,                 // I35 R7.4 - owned-stage cleanup keeps its budget under the cancel watch
    SelfTestState::Step::Delete_NasFileIdInfoUnsupportedByName,           // I35 R7.3 - the NAS without FileIdInfo deletes by name (path.smb.primary)
    SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder,             // Clearflow Phase 1 - copy folder into existing destination folder merges
    SelfTestState::Step::FileOps_MoveMergeIntoExistingFolderSameVolume,   // Clearflow Phase 1 - same-volume move folder into existing destination folder merges
    SelfTestState::Step::Beeline_RenameMergeSkipKeepsSourceFolder,        // Beeline - rename merge answers Overwrite/Skip per child and keeps the source folder
    SelfTestState::Step::FileOps_RenameMergeEmptyFolderCompletes,         // I30 x-datasafety#3 - an empty-folder rename merge completes as Moved
    SelfTestState::Step::FileOps_ReparseDirectoryMergeIntoExistingFolder, // Clearflow Phase 1 - directory reparse copy requires overwrite for existing dir
    SelfTestState::Step::FileOps_ProviderCapabilityMatrix,                // Clearflow Phase 7 - provider capability matrix and offline conformance
    SelfTestState::Step::FileOps_ParallelGraphFairColorWeight,            // Clearflow Phase 3 - parallel graph hue weights are fair
    SelfTestState::Step::FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof, // Clearflow Phase 4 - Copy-only Move retains source without destination
                                                                                   // proof
    SelfTestState::Step::FileOps_ReparseLiteralCopyKeepsTargets,          // Clearflow Phase 4 - copied reparse links keep their stored targets (C3 literal)
    SelfTestState::Step::FileOps_DeleteToctouSwapGuard,                   // Clearflow Phase 4 - delete path rejects stale check/act swaps
    SelfTestState::Step::FileOps_NativeMoveRacedDestinationNotReplaced,   // I30 fs-fileops#2 - Native Move never replaces a destination absent at its probe
    SelfTestState::Step::FileOps_NativeMoveMissingPathKnownFailure,       // I30 fs-fileops#3 - a vanished source or missing parent is a known failure
    SelfTestState::Step::FileOps_ResolvedItemsExactDestinations,          // Clearflow Phase 4 - resolved compare-sync items use exact destinations
    SelfTestState::Step::Riptide_MoveDistinctSameSizeFilePreservesSource, // Riptide R0-1 - same size/mtime move conflict must preserve source
    SelfTestState::Step::Riptide_ReparseNativeMoveRelocatesLinkObject,    // Riptide R0-2 - native reparse Move relocates only the link object
    SelfTestState::Step::Fairstream_ManagedMovePreservesUncopiedNewFiles, // Fairstream 1A - Managed cleanup removes only exact copied entries
    SelfTestState::Step::Fairstream_MoveMergeReadonlyDestinationFolder,   // Fairstream 2C - readonly destination folder merges without prompt
    SelfTestState::Step::Fairstream_OverwriteGrantIsOneShot,              // Fairstream 1D - single Overwrite answer authorizes one file only
    SelfTestState::Step::Fairstream_ConflictPromptsSerializedUnderParallelism, // Fairstream 1C - concurrent conflicts never clobber the active prompt
    SelfTestState::Step::Fairstream_ReparseReplaceNonEmptyDirRequiresConsent,  // Fairstream 1E - non-empty dir replace needs its own answered conflict
    SelfTestState::Step::Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate,   // Floodgate FS-Track-C - apply-to-all decisions stay within one risk bucket
    SelfTestState::Step::Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent,   // Riptide R2-1 - empty real dir cannot be silently converted to a junction
    SelfTestState::Step::Riptide_ReparseMoveSourceNeverUsesDirectoryMerge,     // Riptide R2-1 - source junction never uses directory-merge rename path
    SelfTestState::Step::Fairstream_MovedTreeKeepsLiteralLinks,                // Fairstream 1E - moved trees keep literal link text
    SelfTestState::Step::Fairstream_CrossFsConcurrentMoveUsesBridge,           // Fairstream 1B - concurrent cross-FS move uses the bridge
    SelfTestState::Step::Fairstream_RerunCopyIdenticalCollisionPrompts, // Fairstream 2D - identical destinations require an explicit typed conflict decision
    SelfTestState::Step::Fairstream_MoveSameSizeCollisionPrompts,       // Fairstream 2D - same-size/same-mtime collisions still prompt when bytes differ
    SelfTestState::Step::Fairstream_GraphBandsFairUnderParallelCopy,    // Fairstream 4 - live graph history shows 4 equal bands for 4 equal streams
    SelfTestState::Step::Fairstream_ManagedMergeMovePreservesChildren,  // managed same-volume folder merge preserves all source/destination children
    SelfTestState::Step::Fairstream_JunctionMergeTargetRequiresConsent, // Fairstream 6B - junction destination never merges silently
    SelfTestState::Step::Fairstream_CopyIntoSelfAliasRejected,          // Fairstream 6C - aliased copy-into-self rejected via canonical paths
    SelfTestState::Step::Fairstream_StorageKindProbed,                  // Fairstream 5D - storage profile clamps recursive copy fan-out
    SelfTestState::Step::Fairstream_ParallelDeleteContinuesPastLockedChild,  // Fairstream 7D - locked child does not abort a continue-on-error parallel delete
    SelfTestState::Step::Fairstream_BridgePerFileConflictSkips,              // Cinderstar 19 - serial cross-FS MOVE Skip retains only the skipped source child
    SelfTestState::Step::Cinderstar_BridgeMovePerFileConflictSkipsParallel,  // Cinderstar 19 - configured-parallel MOVE traverses the distinct cleanup path
    SelfTestState::Step::Riptide_BridgeNestedDirVsFileSkipContinuesSiblings, // Riptide R1-1 - bridge child directory-vs-file Skip continues siblings
    SelfTestState::Step::Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch,       // Riptide R1-1 - directory-over-readonly-file is a type mismatch;
                                                                                      // Keep Both preserves the blocker
    SelfTestState::Step::Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial, // Riptide R1-1 - bridge CreateDirectory race re-probes existing file
    SelfTestState::Step::Fairstream_SaturationConcurrentCopiesMakeProgress, // Fairstream 5E - concurrent directory copies share the dynamic-job pool without
                                                                            // stalling
    SelfTestState::Step::Fairstream_DiscoveryAheadOverlapsTransfer,         // single traversal discovers ahead while mutation progresses; totals close
    SelfTestState::Step::Riptide_LiveFinishedSnapshotCarriesDiagnostics,    // Riptide R2-6 - live just-finished popup row keeps diagnostics
    SelfTestState::Step::Riptide_BridgeSequentialContinueOnErrorCopiesSiblings, // Riptide R2-7 - sequential bridge honors continue-on-error
    SelfTestState::Step::Causeway_BridgeRejectsHostileChildNames,               // Causeway CW-1 - hostile provider names never reach path joins
    SelfTestState::Step::Causeway_BridgeProviderOutputContracts,           // Causeway CW-2/CW-2A/CW-T1 - reject invalid reader/writer counts and short streams
    SelfTestState::Step::Causeway_BridgeSchedulingAndResourceContracts,    // Causeway CW-3/CW-15/CW-25/CW-26 - nested scheduling, cloud listings, error
                                                                           // vocabulary, memory budget
    SelfTestState::Step::Causeway_BridgeFileReparsePolicy,                 // Causeway CW-8 - file reparse entries obey Skip/Preserve
    SelfTestState::Step::Beeline_CopySkipLinksKeepsPlaceholders,           // Beeline - Copy Skip skips name-surrogate links only; WOF placeholders copy
    SelfTestState::Step::Causeway_BridgeFailureStatusAndPausedReader,      // Causeway CW-18/CW-24 - paused-reader shutdown and real worker HRESULT propagation
    SelfTestState::Step::Floodgate_CrossFsCopyGetSizeFailureRefusesCommit, // Floodgate FG-P0-2 - COPY refuses unverifiable unknown-size sources
    SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailurePreservesSource, // Floodgate FG-P0-2 - MOVE preserves source when source size is unverifiable
    SelfTestState::Step::Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource, // Floodgate FG-A1 - MOVE preserves source after destination re-stat
                                                                                        // failure
    SelfTestState::Step::Cinderstar_LegacyWriterEventuallyConsistentMove,  // Cinderstar 20 - Copy-only publication converges after one final-path miss
    SelfTestState::Step::Cinderstar_LegacyWriterPermanentMissMove,         // Cinderstar 20 - exhausted final-path misses preserve the Copy-only source
    SelfTestState::Step::Cinderstar_LegacyWriterWrongSizeMove,             // Cinderstar 20 - confirmed wrong size is not retried and preserves source
    SelfTestState::Step::Cinderstar_LegacyWriterConcurrentReplacementCopy, // Cinderstar 20 - post-publication cleanup preserves a replacement
    SelfTestState::Step::Cinderstar_LegacyWriterCancelDuringBackoff,       // Cinderstar 20 - cancellation interrupts Copy-only publication retry backoff
    SelfTestState::Step::Floodgate_CrossFsCopyOnlyUsesCopyVerification,    // P1.3 - CopyOnly uses Copy publication verification and retains source
    SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree, // Floodgate FG-P0-2 - source/destination guards under parallel walk
    SelfTestState::Step::Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup,              // P1.3 - CopyOnly never arms or enters source cleanup
                                                                                         // corruption
    SelfTestState::Step::Floodgate_CrossFsDirectoryCopyOnlyRetainsSource,                // P1.3 - directory CopyOnly retains source and no cleanup manifest
    SelfTestState::Step::Floodgate_LocalWriterOverwriteIsStaged, // Floodgate FG-P0-3 - overwrite writer abort preserves existing destination
    SelfTestState::Step::Floodgate_LocalCopyOverwriteIsStaged,   // Floodgate FS-Track-A - same-FS copy overwrite preserves existing destination until promote
    SelfTestState::Step::Floodgate_LocalCopyNewNameConcurrentReplacementSurvives,     // R0a FOS-01 - exact abort preserves a foreign pathname replacement
    SelfTestState::Step::Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete,  // R0a FOS-01 - known retained partial final leaf is explicit
    SelfTestState::Step::Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker, // Riptide R2-9 - shared scheduler shutdown waits for active callbacks
    SelfTestState::Step::Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker,   // Riptide R2-9 - host per-item scheduler shutdown waits for active
                                                                                      // callbacks
    SelfTestState::Step::Floodgate_InlineF2WorkerQueueBypassAndInterlock, // P1.5 - worker-only F2 bypasses Queue but waits on exact parent interlock
    SelfTestState::Step::C1_InlineRenameNameRefusedOnCard,                // C1 (b0) - an inline rename the provider refuses fails on its card, not in admission
    SelfTestState::Step::R4A19_DiscoveryMeasurementFacts,                 // R4-A19 - observation-only discovery baseline facts
    SelfTestState::Step::R4A19_DiscoveryProviderControls,                 // R4-A19 - delayed Dummy and serialized fake-MTP controls
    SelfTestState::Step::R4A19_DiscoveryIndependentVolumes,               // R4-A19 - marked C/D independent task overlap
    SelfTestState::Step::DiscoveryScope_LeafCopy,                         // I21 - a known leaf closes its exact total before payload work
    SelfTestState::Step::DiscoveryScope_LeafMove,                         // I21 - the same leaf rule holds through exact Move source cleanup
    SelfTestState::Step::DiscoveryScope_ManagedDirectory,                 // I21 - nested bound cleanup cannot close the walker's root
    SelfTestState::Step::DiscoveryScope_NativeRenameMerge,                // I21 - a refused Native rename hands discovery to the merge walk
    SelfTestState::Step::DiscoveryScope_MixedRoots,                       // I21 - the aggregate closes once on exact summed totals
    SelfTestState::Step::DiscoveryScope_ZeroByteAndVerifiedLeaf, // I21 - a known empty file is a known zero total; verification admits the same-endpoint bridge
    SelfTestState::Step::DiscoveryScope_LocalDirectApi,          // I23 - one governed provider call is one discovery scope with one closure
    SelfTestState::Step::DiscoveryScope_DummyDirectApi,          // I22 - Dummy reports every direct route before it mutates
    SelfTestState::Step::DiscoveryScope_CurlDirectApi,           // I22 - Curl reports on its enumeration, singular and bulk
    SelfTestState::Step::DiscoveryScope_S3DirectApi,             // I22 - S3 reports its exact plan total before object work
    SelfTestState::Step::DiscoveryScope_GDriveDirectApi,         // I22 - Google Drive reports the leaf exactly and the tree per child
    SelfTestState::Step::DiscoveryScope_GraphDirectApi,          // I22 - Microsoft Drive Native Move is one closed item-based record
    SelfTestState::Step::DiscoveryScope_MtpDirectApi,            // I22 - MTP reports the leaf from the listing it already holds
    SelfTestState::Step::DiscoveryScope_FolderClosesBeforeLastFile,              // I35 S2 - every walker closes a folder at its listing's end
    SelfTestState::Step::DiscoveryScope_WideFolderClosesWithinQueueLead,         // I35 S2 - a wide folder closes within the queue lead
    SelfTestState::Step::DiscoveryScope_NestedTreeMetadataRestoredAfterChildren, // I35 S2 - directory metadata lands after the last child
    SelfTestState::Step::DiscoveryScope_ManagedMoveListingCompleteBeforeRemoval, // I35 S2 - a Managed Move lists before it removes
    SelfTestState::Step::DiscoveryScope_WorkerFailureIsNotCanceled,              // I35 S2 - a worker's failure is never a cancel
    SelfTestState::Step::DiscoveryScope_FlatFolderUsesFullConcurrency,           // I35 S2 - the reservation ends with the listing
    SelfTestState::Step::DiscoveryScope_ListingSilenceNoted,                     // I35 S2 - one listing-paused note, no closure
    SelfTestState::Step::DiscoveryScope_TenThousandSmallFilesCloseAfterListing,  // I35 D-6.2 - a 10,000-file folder closes once listed
    SelfTestState::Step::DiscoveryScope_RetryDoesNotGrowAfterClose,              // I35 S2 R6.3 - a retried listing counts every entry once
    SelfTestState::Step::DiscoveryScope_HeldFolderFailureStillCloses,            // I35 S2 R6.1 - a skipped held folder no longer holds the scope open
    SelfTestState::Step::DiscoveryScope_FolderLiveProgress,                      // I37 - real nested copy counters and rendered per-file progress
    SelfTestState::Step::DiscoveryScope_MultipleFoldersLiveProgress,             // I37 - distinct roots publish monotonic aggregate progress
    SelfTestState::Step::FileOps_SelectionShapeLocal,
    SelfTestState::Step::FileOps_SelectionShapeSmb,
    SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns, // R0f-SMB - a wedged synchronous provider call returns on cancel
    SelfTestState::Step::R0fSmb_LoopbackReadWriteCreateDelete,       // R0f-SMB - read/write/create/rename/delete on a loopback share
    SelfTestState::Step::R0fSmb_LoopbackRecycleNeverDestroys,        // R0f-SMB - Recycle on a share keeps every item and asks before permanent delete
    SelfTestState::Step::R0fSmb_NestedTreeCopyToShare,               // I35 - a nested tree copies to the registered SMB share, folder times restored
    SelfTestState::Step::R0fSmb_BridgeTreeCopyToShare,               // I35 - a nested tree and a flat folder copy to the SMB share through the host bridge
    SelfTestState::Step::R0fSmb_ManagedMoveTreeToShare,        // I35 - a nested tree and a flat folder move to the SMB share (Managed Move), no stage left
    SelfTestState::Step::R0fCurl_FakeFtpReadWriteCreateDelete, // R0f-Curl - cross-provider copy both ways, create/rename/delete on fake FTP
    SelfTestState::Step::BR5_CurlHostDirectoryMove,
    SelfTestState::Step::BR5_CurlHostDirectoryMoveRefused,
    SelfTestState::Step::R0fS3_FakeS3ReadWriteCreateDelete,              // R0f-S3 - cross-provider copy both ways, create/rename/delete on fake S3
    SelfTestState::Step::R0fGraph_FakeGraphReadWriteCreateRenameRecycle, // R0f-Graph - cross-provider copy both ways, create/rename/recycle on fake Graph
    SelfTestState::Step::R0fGDrive_FakeDriveReadWriteCreateMoveDelete,   // R0f-GDrive - cross-provider copy both ways, create/rename/move/delete on fake Drive
    SelfTestState::Step::C0_GDriveCommittedCopyResponseLost,             // C0/C2 - committed remote copy must never be replayed after response loss
    SelfTestState::Step::C0_GDriveCommittedDeleteResponseLost,
    SelfTestState::Step::C0_NativeCopyFailurePreservesKnownAxes,
    SelfTestState::Step::C0_MutationReceiptPrefixBoundary,
    SelfTestState::Step::C0_LocalKnownNoCommitReceipts, // C0 - Local Cancel/refused-rename/merge-cancel receipts stay known, never uncertain
    SelfTestState::Step::C0_CurlNativeDeleteGuards,
    SelfTestState::Step::C0_CurlCommittedMutationResponseLost,
    SelfTestState::Step::C0_CurlHostCommittedDeleteResponseLost,
    SelfTestState::Step::C0_CurlPartialTreeFailure,
    SelfTestState::Step::C0_CurlHostPartialTreeFailure,
    SelfTestState::Step::C0_CurlNativeDeleteLateListing,
    SelfTestState::Step::C0_CurlDeleteTraversalBounds,
    SelfTestState::Step::C0_CurlDirectorySizeTruth,
    SelfTestState::Step::C0_CurlMovePreflightTruth,
    SelfTestState::Step::C0_CurlCopyTraversalTruth,
    SelfTestState::Step::C0_CurlEntryLookupTruth,
    SelfTestState::Step::C0_CurlImapListingTruth,
    SelfTestState::Step::C0_CurlImapTransportTruth,     // C0 - production IMAP owners through real libcurl against a loopback server
    SelfTestState::Step::R3_1_IdentityLessReplaceDummy, // R3-1 - conditional replace through an atomic-final writer without bound authority
    SelfTestState::Step::R3_2_WriterProofDummy,         // R3-2 - writer content proof: Verified Copy without readback, Managed Move cleanup after proof
    SelfTestState::Step::R3_3_PublicationFaultMatrix,   // R3-3 - every bridge fault hook x route x Copy/Move against the publication invariants
    SelfTestState::Step::R3_4_ReplaceWithoutOccupantTokenRefused, // R0-RC3 - a granted replacement without an occupant token is refused, nothing written
    SelfTestState::Step::RC3_9_TransferDestinationNameRefused,    // R0-RC3 (9) - a destination leaf the provider rejects is refused at admission
    SelfTestState::Step::R4T_DeepTreeCopyCompletesIteratively,    // R4-T1 - a 300-level tree copies and moves through the iterative bridge walkers
    SelfTestState::Step::R4T2_DeepTreeLocalDeleteCompletes,       // R4-T2 - a 300-level Local chain beyond MAX_PATH is deleted through the frame-based walk
    SelfTestState::Step::R4T3_WideDirectoryCopyCompletes, // R4-T3 - a 20,000-child directory copies through the bridge with its listing reported, not capped
    SelfTestState::Step::R4T4_LocalManyDirectoriesCopy,   // I35 - a 1,885-folder Local tree copies three times, every folder's time restored
    SelfTestState::Step::R4T5_LocalCopyHoldsNoDirectoryHandle,        // I35 - the Local walker keeps no handle on a folder whose children copy
    SelfTestState::Step::R4T6_LocalDirectoryRestoreChecksIdentity,    // I35 - a folder's times are restored only on the folder the walker created
    SelfTestState::Step::R4T7_BridgeCopyHoldsNoDirectoryHandle,       // I35 - the host bridge keeps no handle on a folder whose children copy
    SelfTestState::Step::R4T8_BridgeDirectoryRestoreChecksIdentity,   // I35 - the bridge restores a folder's times only on the folder it created
    SelfTestState::Step::R4T9_BridgeManyDirectoriesCopy,              // I35 - a 1,885-folder tree copies twice through the bridge, times restored
    SelfTestState::Step::R4A02_DeleteOverCopySourceWarnsAndQueues,    // R4-A02-1 - a Delete over a running Copy's source warns and queues behind it
    SelfTestState::Step::R4A02_DontStartCancelsBeforeMutation,        // R4-A02-1 - Don't start ends the overlapping task before any mutation
    SelfTestState::Step::R4A02_ReadReadDoesNotWarn,                   // R4-A02-1 - two Copies reading one source warn nothing
    SelfTestState::Step::R4A02_SameDestinationWarns,                  // R4-A02-1 - simultaneous Copies publishing the same leaf warn before mutation
    SelfTestState::Step::R4A02_RunConcurrentSameDestinationStaysSafe, // R4-A02-2 - Run admits only the disclosed overlap and both result sets stay truthful
    SelfTestState::Step::R4A02_LiveOutputGuardChoices,                // R4-A02-2 - uncovered live output parks; Skip/Queue/destructive and overflow stay safe
    SelfTestState::Step::R4A02_DisjointWritesDoNotWarn,               // R4-A02-1 - disjoint source and destination roots warn nothing
    SelfTestState::Step::R4A02_RenamePublishKeepsIndexPrecise, // Step 3 (2026-09-06) - a rename publication is indexed under its parent scope, no host-wide
                                                               // overflow
    SelfTestState::Step::R4A02_LatePublisherStillWarns, // Step 4 (2026-09-06) - a task that publishes after a newer peer prepared still warns and queues after
                                                        // it
    SelfTestState::Step::Phase5_DiscoverySingleTraversal,                  // Phase 5 — one traversal feeds execution and closes totals
    SelfTestState::Step::Phase5_DiscoveryCancelReleasesSlot,               // Phase 5 — cancel releases discovery/execution capacity
    SelfTestState::Step::Phase5_CleanupDrainsOutstandingTasks,             // Failed-family cleanup drains paused and queued work
    SelfTestState::Step::Phase5_DiscoveryCancelLatencyLocal,               // Phase 5 — local discovery/transfer observes bounded cancellation
    SelfTestState::Step::Phase5_DiscoverySkipContinues,                    // Phase 5 — Skip switches to just-in-time and completes the same selection
    SelfTestState::Step::BR3_CancelQueuePausedTransfer,                    // Cancellation drains independently of the paused queue predecessor
    SelfTestState::Step::Phase5_CancelQueuedTask,                          // Phase 5 â€” canceling a queued (not-yet-running) task
    SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery,       // Phase 5 — mode switch parallel→wait while discovery remains open
    SelfTestState::Step::Phase5_SwitchWaitToParallelResume,                // Phase 5 â€” mode switch waitâ†’parallel and resume
    SelfTestState::Step::Phase6_PopupRateSmoothing,                        // Phase 6 - popup rate/ETA smoothing contract
    SelfTestState::Step::Phase6_PopupSmokeResizeAndPause,                  // Phase 6 â€” popup resize and pause-button interaction
    SelfTestState::Step::Phase6_DeleteBytesMeaningful,                     // Phase 6 â€” delete reports meaningful byte counts in progress
    SelfTestState::Step::Phase6_LocalBandwidthThrottle,                    // Phase 6 â€” local CopyFileEx bandwidth throttle duration and cancel latency
    SelfTestState::Step::Phase6_ParallelBandwidthThrottleFairness,         // Phase 6 â€” parallel throttle fairness: shared-only vs per-worker budget
    SelfTestState::Step::Phase7_WatcherChurn,                              // Phase 7 â€” directory watcher fires correctly under heavy churn
    SelfTestState::Step::Phase7_CacheBorrowNoWatchInvalidation,            // Phase 7 â€” cached-but-unwatched folders still invalidate on routed mutations
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshLocal,              // Phase 7 â€” visible local panes auto-refresh on watcher-driven changes
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshDummy,              // Phase 7 â€” visible dummy panes stay in sync after in-app mutation routing
    SelfTestState::Step::Phase7_CrossPaneRelocateLocal,                    // Phase 7 â€” deleting a visible subtree relocates impacted panes
    SelfTestState::Step::Phase7_LargeDirectoryEnumeration,                 // Phase 7 â€” enumerate a directory with many entries
    SelfTestState::Step::Phase7_ParallelCopyMoveKnobs,                     // Phase 7 â€” speed limits and parallelism knobs for copy/move
    SelfTestState::Step::Phase7_CopyMoveConcurrency16Perf,                 // Phase 7 / 8.2 â€” compare copy/move cap 8 vs 16 on the same machine
    SelfTestState::Step::Phase7_AutoConcurrencyHints,                      // Phase 7 â€” auto mode resolves copy/delete concurrency from storage hints
    SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines,         // Phase 7 â€” per-item directory copy uses internal parallelism (popup lines)
    SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism, // Phase 7 - plugin CopyItem/CopyItems(count==1) use recursive internal parallelism
    SelfTestState::Step::Phase7_CopyItemsMultiRootUnevenRecursiveParallelism,  // Phase 7 - multi-root copies rebalance uneven recursive subtrees
    SelfTestState::Step::Phase7_CopyRecursiveParallelismMatrix,                // Phase 7 - recursive copy/move matrix shapes and edge paths
    SelfTestState::Step::Phase7_CopyItemsSaturatedPoolProducersHelp,           // Phase 7 - bulk CopyItems roots on every scheduler worker still finish
    SelfTestState::Step::Phase7_CopyItemWaitingProducerSeesCancel,             // Phase 7 - a producer waiting for queue room sees the caller's cancel
    SelfTestState::Step::Phase7_CopyItemWalkAtLimitParks,                      // Phase 7 - a walk at its active limit is parked, not dispatched in a loop
    SelfTestState::Step::Phase7_DeleteItemsNestedRootsWake,                    // Phase 7 - a parked parallel delete is woken by its own step
    SelfTestState::Step::Phase7_SharedPerItemScheduler,                        // Phase 7 â€” shared per-item scheduler across parallel tasks
    SelfTestState::Step::Phase7_ParallelDeleteKnobs,                           // Phase 7 â€” speed limits and parallelism knobs for delete
    SelfTestState::Step::Phase7_RecycleBinBatchDelete,                         // Phase 7 â€” multi-item recycle-bin delete batches sibling items
    SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch,               // Phase 7 â€” recycle-bin delete spans multiple sibling batches when >500 inputs
    SelfTestState::Step::Phase7_RecycleBinTestItemsPurged,                     // Phase 7 - a recycled test item leaves the Recycle Bin when its case ends
    SelfTestState::Step::Phase8_DefaultBandwidthLimitFromSettings,             // Phase 8 â€” new copy/move tasks snapshot the global default speed limit
    SelfTestState::Step::Phase8_TightDefaults_NoOverwrite,                     // Phase 8 â€” no-overwrite default returns correct HRESULT
    SelfTestState::Step::Phase8_InvalidDestinationRejected,                    // Phase 8 â€” invalid destination is rejected before op starts
    SelfTestState::Step::Phase8_InvalidSizeBytesRejected,                      // Phase 8 â€” invalid sizeBytes is rejected at the ABI boundary
    SelfTestState::Step::Phase8_PerItemOrchestration,                          // Phase 8 â€” per-item mode orchestrates items one by one
    SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly,       // Phase 9 â€” overwrite read-only via conflict prompt
    SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache,              // Phase 9 â€” apply-to-all caching in conflict prompt UI
    SelfTestState::Step::Phase9_ConflictPrompt_KeepBothNestedCacheEligibility, // Phase 9 — cached Keep Both stays top-level; nested explicit choice remains
    SelfTestState::Step::Phase9_ConflictPrompt_TypeMismatchNoOverwrite,        // Phase 9 — typed file-on-folder conflict never exposes Overwrite
    SelfTestState::Step::Phase9_ConflictPrompt_LocalFileOntoDirectory,         // Phase 9 - local file-on-directory prompt excludes Overwrite
    SelfTestState::Step::Phase9_ConflictPrompt_SkipApplyToAll,                 // Phase 9 â€” Skip + All similar conflict decision cache
    SelfTestState::Step::Phase9_ConflictPrompt_RetryCap,                       // Phase 9 â€” retry cap in conflict prompt
    SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy,     // Phase 9 â€” skip continues directory copy
    SelfTestState::Step::Phase9_PerItemConcurrency,                            // Phase 9 â€” per-item mode with concurrent operations
    SelfTestState::Step::Phase10_PermanentDelete,                              // Phase 10 â€” confirmed permanent delete
    SelfTestState::Step::Phase10_DeferredConsentAndRecycleEscalation,          // Phase 10 / P4.2 — typed deferred consent and exact Recycle escalation
    SelfTestState::Step::Phase10_RecycleDestroyedIsNeverRecycled,              // Phase 10 - a Recycle the provider reports as destroyed fails, unescalated
    SelfTestState::Step::Phase10_RecycleEscalationDirectoryTree,               // Phase 10 - a failed folder Recycle escalates to a counted recursive delete
    SelfTestState::Step::Phase10_RecycleEscalationTaskScopedGrant,             // Phase 10 - one task-scoped Delete all remaining permanently grant (Q2)
    SelfTestState::Step::Phase10_RecycleEscalationBulkItemsContinueOnError,    // Phase 10 - pane Delete honours continue-on-error; folder escalation
    SelfTestState::Step::Phase10_RecycleEscalationBulkItemsResumeBounded,      // Phase 10 - pane Delete resume: bounded provider calls, unknown never resumed
    SelfTestState::Step::Phase10_MetadataPreservationAndSourceRetention,       // Phase 10 / P4.3 — exact metadata and security-loss retention
    SelfTestState::Step::Phase10_TypedResultsAndConsumers,                     // Phase 10 / P4.4 — typed item axes and exact consumer removal
    SelfTestState::Step::Phase10_ClipboardAdmissionAndRetainedActions,         // Phase 10 / P4.5 — accepted cut-list barrier and retained actions
    SelfTestState::Step::Phase10_ContentVerification,                          // Phase 10 / P4.6 — optional BLAKE3 verification and truthful outcomes
    SelfTestState::Step::Phase10_ArtifactTouchGuard,                           // Phase 10 / P5.4 — external launch artifact warning and exact revalidation
    SelfTestState::Step::R1d_PreparingLifecycle,                               // R1d — common pre-consumption lifecycle and immutable preparation facts
    SelfTestState::Step::C1_ExitCloseDeferredUntilTasksQuiet, // C1 - application exit drains cancellation off the UI thread; the close resumes after the last
                                                              // reap
    SelfTestState::Step::Beeline_SameVolumeTreeMoveIsRename,  // Beeline - a same-volume file and tree Move is one Native rename per item, identity preserved
    SelfTestState::Step::Beeline_LoopbackShareMoveIsRename,   // Beeline - the SMB profile takes the same route: a Move inside the loopback share alias is one
                                                              // Native rename per item
    SelfTestState::Step::C1_PermanentDeleteConfirmsOnCard,    // C1 - the permanent-delete confirmation is a card consent after Preparing pins the roots
    SelfTestState::Step::Phase11_CrossFileSystemBridge,       // Phase 11 â€” copy/move across different file-system plugins
    SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines, // Phase 11 â€” bridge: single-folder copy uses within-folder parallelism
    SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines,  // Phase 11 â€” bridge: multi-folder copy keeps bounded early file admission
    SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf,              // Phase 11 â€” bridge: pipeline overlaps dummy read/write chunk latency
    SelfTestState::Step::Phase11_ConnectionOverridePrecedence,                // Phase 11 / 8 â€” non-zero @conn override takes precedence over plugin default
    SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate,                // Phase 11 â€” connection overrides apply globally across tasks
    SelfTestState::Step::Phase11_ConnectionOverrideClamp,                     // Phase 11 â€” connection manager overrides clamp per-task concurrency
    SelfTestState::Step::Phase12_ReparsePointPolicy,                          // Phase 12 â€” reparse-point (symlink/junction) handling policy
    SelfTestState::Step::Phase13_PostMortemDiagnostics,                       // Phase 13 â€” post-mortem diagnostics on task failure
    SelfTestState::Step::Phase14_PopupHostLifetimeGuard,                      // Phase 14 â€” popup host lifetime guard (no UAF on late input)
    SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke,                   // Phase 15 â€” 7z IFileReader read/seek smoke on a large (>32MB) entry
    SelfTestState::Step::Phase15_FileSystem7zMountPathImpact,                 // Phase 15 â€” 7z mounted panes retarget/exit on backing archive changes
    SelfTestState::Step::Phase16_RemoteWatchContractExposure,                 // Phase 16 â€” mutable remote plugins expose the watch contract offline
    SelfTestState::Step::Phase16_RemoteFtpSecret,                             // Phase 16 â€” secure secret retrieval (FTP)
    SelfTestState::Step::Phase16_RemoteFtpSandbox,                            // Phase 16 â€” remote sandbox root configuration (FTP)
    SelfTestState::Step::Phase16_RemoteSftpSecret,                            // Phase 16 â€” secure secret retrieval (SFTP)
    SelfTestState::Step::Phase16_RemoteSftpSandbox,                           // Phase 16 â€” remote sandbox root configuration (SFTP)
    SelfTestState::Step::Phase16_RemoteScpSecret,                             // Phase 16 â€” secure secret retrieval (SCP)
    SelfTestState::Step::Phase16_RemoteScpSandbox,                            // Phase 16 â€” remote sandbox root configuration (SCP)
    SelfTestState::Step::Phase16_RemoteImapSecret,                            // Phase 16 â€” secure secret retrieval (IMAP)
    SelfTestState::Step::Phase16_RemoteImapSandbox,                           // Phase 16 â€” remote sandbox root configuration (IMAP)
    SelfTestState::Step::Phase16_RemoteS3Secret,                              // Phase 16 â€” secure secret retrieval (S3)
    SelfTestState::Step::Phase16_RemoteS3Sandbox,                             // Phase 16 â€” remote sandbox root configuration (S3)
    SelfTestState::Step::Phase16_RemoteS3FileOps,                             // Phase 16 â€” sandboxed CRUD/copy-move coverage (S3)
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret,                // Phase 16 â€” secure secret retrieval (OneDrive Personal)
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSandbox,               // Phase 16 â€” remote sandbox root configuration (OneDrive Personal)
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalFileOps,               // Phase 16 â€” sandboxed CRUD/copy-move coverage (OneDrive Personal)
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret,                // Phase 16 â€” secure secret retrieval (OneDrive Business)
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSandbox,               // Phase 16 â€” remote sandbox root configuration (OneDrive Business)
    SelfTestState::Step::Phase16_RemoteSharePointSecret,                      // Phase 16 â€” secure secret retrieval (SharePoint)
    SelfTestState::Step::Phase16_RemoteSharePointSandbox,                     // Phase 16 â€” remote sandbox root configuration (SharePoint)
    SelfTestState::Step::Cleanup_RestorePluginConfig                          // Restore plugin config and delete temp files
});

constexpr std::array<SelfTestState::Step, 5> kFileOpsFamilyClearflowPhase01{{
    SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder,
    SelfTestState::Step::FileOps_MoveMergeIntoExistingFolderSameVolume,
    SelfTestState::Step::Beeline_RenameMergeSkipKeepsSourceFolder,
    SelfTestState::Step::FileOps_RenameMergeEmptyFolderCompletes,
    SelfTestState::Step::FileOps_ReparseDirectoryMergeIntoExistingFolder,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyClearflowPhase07{{
    SelfTestState::Step::FileOps_ProviderCapabilityMatrix,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyClearflowPhase03{{
    SelfTestState::Step::FileOps_ParallelGraphFairColorWeight,
}};

constexpr std::array<SelfTestState::Step, 6> kFileOpsFamilyClearflowPhase04{{
    SelfTestState::Step::FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof,
    SelfTestState::Step::FileOps_ReparseLiteralCopyKeepsTargets,
    SelfTestState::Step::FileOps_DeleteToctouSwapGuard,
    SelfTestState::Step::FileOps_NativeMoveRacedDestinationNotReplaced,
    SelfTestState::Step::FileOps_NativeMoveMissingPathKnownFailure,
    SelfTestState::Step::FileOps_ResolvedItemsExactDestinations,
}};

constexpr std::array<SelfTestState::Step, 55> kFileOpsFamilyFairstream{{
    SelfTestState::Step::Riptide_MoveDistinctSameSizeFilePreservesSource,
    SelfTestState::Step::Riptide_ReparseNativeMoveRelocatesLinkObject,
    SelfTestState::Step::Fairstream_ManagedMovePreservesUncopiedNewFiles,
    SelfTestState::Step::Fairstream_MoveMergeReadonlyDestinationFolder,
    SelfTestState::Step::Fairstream_OverwriteGrantIsOneShot,
    SelfTestState::Step::Fairstream_ConflictPromptsSerializedUnderParallelism,
    SelfTestState::Step::Fairstream_ReparseReplaceNonEmptyDirRequiresConsent,
    SelfTestState::Step::Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate,
    SelfTestState::Step::Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent,
    SelfTestState::Step::Riptide_ReparseMoveSourceNeverUsesDirectoryMerge,
    SelfTestState::Step::Fairstream_MovedTreeKeepsLiteralLinks,
    SelfTestState::Step::Fairstream_CrossFsConcurrentMoveUsesBridge,
    SelfTestState::Step::Fairstream_RerunCopyIdenticalCollisionPrompts,
    SelfTestState::Step::Fairstream_MoveSameSizeCollisionPrompts,
    SelfTestState::Step::Fairstream_GraphBandsFairUnderParallelCopy,
    SelfTestState::Step::Fairstream_ManagedMergeMovePreservesChildren,
    SelfTestState::Step::Fairstream_JunctionMergeTargetRequiresConsent,
    SelfTestState::Step::Fairstream_CopyIntoSelfAliasRejected,
    SelfTestState::Step::Fairstream_StorageKindProbed,
    SelfTestState::Step::Fairstream_ParallelDeleteContinuesPastLockedChild,
    SelfTestState::Step::Fairstream_BridgePerFileConflictSkips,
    SelfTestState::Step::Cinderstar_BridgeMovePerFileConflictSkipsParallel,
    SelfTestState::Step::Riptide_BridgeNestedDirVsFileSkipContinuesSiblings,
    SelfTestState::Step::Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch,
    SelfTestState::Step::Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial,
    SelfTestState::Step::Fairstream_SaturationConcurrentCopiesMakeProgress,
    SelfTestState::Step::Fairstream_DiscoveryAheadOverlapsTransfer,
    SelfTestState::Step::Riptide_LiveFinishedSnapshotCarriesDiagnostics,
    SelfTestState::Step::Riptide_BridgeSequentialContinueOnErrorCopiesSiblings,
    SelfTestState::Step::Causeway_BridgeRejectsHostileChildNames,
    SelfTestState::Step::Causeway_BridgeProviderOutputContracts,
    SelfTestState::Step::Causeway_BridgeSchedulingAndResourceContracts,
    SelfTestState::Step::Causeway_BridgeFileReparsePolicy,
    SelfTestState::Step::Beeline_CopySkipLinksKeepsPlaceholders,
    SelfTestState::Step::Causeway_BridgeFailureStatusAndPausedReader,
    SelfTestState::Step::Floodgate_CrossFsCopyGetSizeFailureRefusesCommit,
    SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailurePreservesSource,
    SelfTestState::Step::Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource,
    SelfTestState::Step::Cinderstar_LegacyWriterEventuallyConsistentMove,
    SelfTestState::Step::Cinderstar_LegacyWriterPermanentMissMove,
    SelfTestState::Step::Cinderstar_LegacyWriterWrongSizeMove,
    SelfTestState::Step::Cinderstar_LegacyWriterConcurrentReplacementCopy,
    SelfTestState::Step::Cinderstar_LegacyWriterCancelDuringBackoff,
    SelfTestState::Step::Floodgate_CrossFsCopyOnlyUsesCopyVerification,
    SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree,
    SelfTestState::Step::Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup,
    SelfTestState::Step::Floodgate_CrossFsDirectoryCopyOnlyRetainsSource,
    SelfTestState::Step::Floodgate_LocalWriterOverwriteIsStaged,
    SelfTestState::Step::Floodgate_LocalCopyOverwriteIsStaged,
    SelfTestState::Step::Floodgate_LocalCopyNewNameConcurrentReplacementSurvives,
    SelfTestState::Step::Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete,
    SelfTestState::Step::Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker,
    SelfTestState::Step::Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker,
    SelfTestState::Step::Floodgate_InlineF2WorkerQueueBypassAndInterlock,
    SelfTestState::Step::C1_InlineRenameNameRefusedOnCard,
}};

constexpr std::array<SelfTestState::Step, 4> kFileOpsFamilyClearflowPhase05{{
    SelfTestState::Step::Phase5_DiscoverySingleTraversal,
    SelfTestState::Step::Phase5_DiscoveryCancelLatencyLocal,
    SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines,
    SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism,
}};

constexpr std::array<SelfTestState::Step, 9> kFileOpsFamilyPhase05{{
    SelfTestState::Step::Phase5_DiscoverySingleTraversal,
    SelfTestState::Step::Phase5_DiscoveryCancelReleasesSlot,
    SelfTestState::Step::Phase5_CleanupDrainsOutstandingTasks,
    SelfTestState::Step::Phase5_DiscoveryCancelLatencyLocal,
    SelfTestState::Step::Phase5_DiscoverySkipContinues,
    SelfTestState::Step::BR3_CancelQueuePausedTransfer,
    SelfTestState::Step::Phase5_CancelQueuedTask,
    SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery,
    SelfTestState::Step::Phase5_SwitchWaitToParallelResume,
}};

constexpr std::array<SelfTestState::Step, 5> kFileOpsFamilyPhase06{{
    SelfTestState::Step::Phase6_PopupRateSmoothing,
    SelfTestState::Step::Phase6_PopupSmokeResizeAndPause,
    SelfTestState::Step::Phase6_DeleteBytesMeaningful,
    SelfTestState::Step::Phase6_LocalBandwidthThrottle,
    SelfTestState::Step::Phase6_ParallelBandwidthThrottleFairness,
}};

constexpr std::array<SelfTestState::Step, 22> kFileOpsFamilyPhase07{{
    SelfTestState::Step::Phase7_WatcherChurn,
    SelfTestState::Step::Phase7_CacheBorrowNoWatchInvalidation,
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshLocal,
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshDummy,
    SelfTestState::Step::Phase7_CrossPaneRelocateLocal,
    SelfTestState::Step::Phase7_LargeDirectoryEnumeration,
    SelfTestState::Step::Phase7_ParallelCopyMoveKnobs,
    SelfTestState::Step::Phase7_CopyMoveConcurrency16Perf,
    SelfTestState::Step::Phase7_AutoConcurrencyHints,
    SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines,
    SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism,
    SelfTestState::Step::Phase7_CopyItemsMultiRootUnevenRecursiveParallelism,
    SelfTestState::Step::Phase7_CopyRecursiveParallelismMatrix,
    SelfTestState::Step::Phase7_CopyItemsSaturatedPoolProducersHelp,
    SelfTestState::Step::Phase7_CopyItemWaitingProducerSeesCancel,
    SelfTestState::Step::Phase7_CopyItemWalkAtLimitParks,
    SelfTestState::Step::Phase7_DeleteItemsNestedRootsWake,
    SelfTestState::Step::Phase7_SharedPerItemScheduler,
    SelfTestState::Step::Phase7_ParallelDeleteKnobs,
    SelfTestState::Step::Phase7_RecycleBinBatchDelete,
    SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch,
    SelfTestState::Step::Phase7_RecycleBinTestItemsPurged,
}};

constexpr std::array<SelfTestState::Step, 5> kFileOpsFamilyPhase08{{
    SelfTestState::Step::Phase8_DefaultBandwidthLimitFromSettings,
    SelfTestState::Step::Phase8_TightDefaults_NoOverwrite,
    SelfTestState::Step::Phase8_InvalidDestinationRejected,
    SelfTestState::Step::Phase8_InvalidSizeBytesRejected,
    SelfTestState::Step::Phase8_PerItemOrchestration,
}};

constexpr std::array<SelfTestState::Step, 9> kFileOpsFamilyPhase09{{
    SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly,
    SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache,
    SelfTestState::Step::Phase9_ConflictPrompt_KeepBothNestedCacheEligibility,
    SelfTestState::Step::Phase9_ConflictPrompt_TypeMismatchNoOverwrite,
    SelfTestState::Step::Phase9_ConflictPrompt_LocalFileOntoDirectory,
    SelfTestState::Step::Phase9_ConflictPrompt_SkipApplyToAll,
    SelfTestState::Step::Phase9_ConflictPrompt_RetryCap,
    SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy,
    SelfTestState::Step::Phase9_PerItemConcurrency,
}};

constexpr std::array<SelfTestState::Step, 12> kFileOpsFamilyPhase10{{
    SelfTestState::Step::Phase10_PermanentDelete,
    SelfTestState::Step::Phase10_DeferredConsentAndRecycleEscalation,
    SelfTestState::Step::Phase10_RecycleDestroyedIsNeverRecycled,
    SelfTestState::Step::Phase10_RecycleEscalationDirectoryTree,
    SelfTestState::Step::Phase10_RecycleEscalationTaskScopedGrant,
    SelfTestState::Step::Phase10_RecycleEscalationBulkItemsContinueOnError,
    SelfTestState::Step::Phase10_RecycleEscalationBulkItemsResumeBounded,
    SelfTestState::Step::Phase10_MetadataPreservationAndSourceRetention,
    SelfTestState::Step::Phase10_TypedResultsAndConsumers,
    SelfTestState::Step::Phase10_ClipboardAdmissionAndRetainedActions,
    SelfTestState::Step::Phase10_ContentVerification,
    SelfTestState::Step::Phase10_ArtifactTouchGuard,
}};

constexpr std::array<SelfTestState::Step, 3> kFileOpsFamilyR4A19DiscoveryBaseline{{
    SelfTestState::Step::R4A19_DiscoveryMeasurementFacts,
    SelfTestState::Step::R4A19_DiscoveryProviderControls,
    SelfTestState::Step::R4A19_DiscoveryIndependentVolumes,
}};

constexpr std::array<SelfTestState::Step, 25> kFileOpsFamilyDiscoveryScope{{
    SelfTestState::Step::DiscoveryScope_LeafCopy,
    SelfTestState::Step::DiscoveryScope_LeafMove,
    SelfTestState::Step::DiscoveryScope_ManagedDirectory,
    SelfTestState::Step::DiscoveryScope_NativeRenameMerge,
    SelfTestState::Step::DiscoveryScope_MixedRoots,
    SelfTestState::Step::DiscoveryScope_ZeroByteAndVerifiedLeaf,
    SelfTestState::Step::DiscoveryScope_LocalDirectApi,
    SelfTestState::Step::DiscoveryScope_DummyDirectApi,
    SelfTestState::Step::DiscoveryScope_CurlDirectApi,
    SelfTestState::Step::DiscoveryScope_S3DirectApi,
    SelfTestState::Step::DiscoveryScope_GDriveDirectApi,
    SelfTestState::Step::DiscoveryScope_GraphDirectApi,
    SelfTestState::Step::DiscoveryScope_MtpDirectApi,
    SelfTestState::Step::DiscoveryScope_FolderClosesBeforeLastFile,
    SelfTestState::Step::DiscoveryScope_WideFolderClosesWithinQueueLead,
    SelfTestState::Step::DiscoveryScope_NestedTreeMetadataRestoredAfterChildren,
    SelfTestState::Step::DiscoveryScope_ManagedMoveListingCompleteBeforeRemoval,
    SelfTestState::Step::DiscoveryScope_WorkerFailureIsNotCanceled,
    SelfTestState::Step::DiscoveryScope_FlatFolderUsesFullConcurrency,
    SelfTestState::Step::DiscoveryScope_ListingSilenceNoted,
    SelfTestState::Step::DiscoveryScope_TenThousandSmallFilesCloseAfterListing,
    SelfTestState::Step::DiscoveryScope_RetryDoesNotGrowAfterClose,
    SelfTestState::Step::DiscoveryScope_HeldFolderFailureStillCloses,
    SelfTestState::Step::DiscoveryScope_FolderLiveProgress,
    SelfTestState::Step::DiscoveryScope_MultipleFoldersLiveProgress,
}};

constexpr std::array<SelfTestState::Step, 5> kFileOpsFamilyR1dPreparing{{
    SelfTestState::Step::R1d_PreparingLifecycle,
    SelfTestState::Step::C1_ExitCloseDeferredUntilTasksQuiet,
    SelfTestState::Step::Beeline_SameVolumeTreeMoveIsRename,
    SelfTestState::Step::Beeline_LoopbackShareMoveIsRename,
    SelfTestState::Step::C1_PermanentDeleteConfirmsOnCard,
}};

constexpr std::array<SelfTestState::Step, 6> kFileOpsFamilyR0fSmbContainment{{
    SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns,
    SelfTestState::Step::R0fSmb_LoopbackReadWriteCreateDelete,
    SelfTestState::Step::R0fSmb_LoopbackRecycleNeverDestroys,
    SelfTestState::Step::R0fSmb_NestedTreeCopyToShare,
    SelfTestState::Step::R0fSmb_BridgeTreeCopyToShare,
    SelfTestState::Step::R0fSmb_ManagedMoveTreeToShare,
}};

constexpr std::array<SelfTestState::Step, 3> kFileOpsFamilyR0fCurlContainment{{
    SelfTestState::Step::R0fCurl_FakeFtpReadWriteCreateDelete,
    SelfTestState::Step::BR5_CurlHostDirectoryMove,
    SelfTestState::Step::BR5_CurlHostDirectoryMoveRefused,
}};

// C0 receipts and Curl truth: these steps were outside every family, so no Fresh Full executed them
// ("not executed" in the gate records through gate #20). Families are what the harness runs.
constexpr std::array<SelfTestState::Step, 3> kFileOpsFamilyC0MutationReceipts{{
    SelfTestState::Step::C0_NativeCopyFailurePreservesKnownAxes,
    SelfTestState::Step::C0_MutationReceiptPrefixBoundary,
    SelfTestState::Step::C0_LocalKnownNoCommitReceipts,
}};

constexpr std::array<SelfTestState::Step, 13> kFileOpsFamilyC0CurlTruth{{
    SelfTestState::Step::C0_CurlNativeDeleteGuards,
    SelfTestState::Step::C0_CurlCommittedMutationResponseLost,
    SelfTestState::Step::C0_CurlHostCommittedDeleteResponseLost,
    SelfTestState::Step::C0_CurlPartialTreeFailure,
    SelfTestState::Step::C0_CurlHostPartialTreeFailure,
    SelfTestState::Step::C0_CurlNativeDeleteLateListing,
    SelfTestState::Step::C0_CurlDeleteTraversalBounds,
    SelfTestState::Step::C0_CurlDirectorySizeTruth,
    SelfTestState::Step::C0_CurlMovePreflightTruth,
    SelfTestState::Step::C0_CurlCopyTraversalTruth,
    SelfTestState::Step::C0_CurlEntryLookupTruth,
    SelfTestState::Step::C0_CurlImapListingTruth,
    SelfTestState::Step::C0_CurlImapTransportTruth,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyR0fS3Containment{{
    SelfTestState::Step::R0fS3_FakeS3ReadWriteCreateDelete,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyR0fGraphContainment{{
    SelfTestState::Step::R0fGraph_FakeGraphReadWriteCreateRenameRecycle,
}};

constexpr std::array<SelfTestState::Step, 3> kFileOpsFamilyR0fGDriveContainment{{
    SelfTestState::Step::R0fGDrive_FakeDriveReadWriteCreateMoveDelete,
    SelfTestState::Step::C0_GDriveCommittedCopyResponseLost,
    SelfTestState::Step::C0_GDriveCommittedDeleteResponseLost,
}};

constexpr std::array<SelfTestState::Step, 5> kFileOpsFamilyR3OwnedPublication{{
    SelfTestState::Step::R3_1_IdentityLessReplaceDummy,
    SelfTestState::Step::R3_2_WriterProofDummy,
    SelfTestState::Step::R3_3_PublicationFaultMatrix,
    SelfTestState::Step::R3_4_ReplaceWithoutOccupantTokenRefused,
    SelfTestState::Step::RC3_9_TransferDestinationNameRefused,
}};

constexpr std::array<SelfTestState::Step, 9> kFileOpsFamilyR4Traversal{{
    SelfTestState::Step::R4T_DeepTreeCopyCompletesIteratively,
    SelfTestState::Step::R4T2_DeepTreeLocalDeleteCompletes,
    SelfTestState::Step::R4T3_WideDirectoryCopyCompletes,
    SelfTestState::Step::R4T4_LocalManyDirectoriesCopy,
    SelfTestState::Step::R4T5_LocalCopyHoldsNoDirectoryHandle,
    SelfTestState::Step::R4T6_LocalDirectoryRestoreChecksIdentity,
    SelfTestState::Step::R4T7_BridgeCopyHoldsNoDirectoryHandle,
    SelfTestState::Step::R4T8_BridgeDirectoryRestoreChecksIdentity,
    SelfTestState::Step::R4T9_BridgeManyDirectoriesCopy,
}};

constexpr std::array<SelfTestState::Step, 9> kFileOpsFamilyR4Overlap{{
    SelfTestState::Step::R4A02_DeleteOverCopySourceWarnsAndQueues,
    SelfTestState::Step::R4A02_DontStartCancelsBeforeMutation,
    SelfTestState::Step::R4A02_ReadReadDoesNotWarn,
    SelfTestState::Step::R4A02_SameDestinationWarns,
    SelfTestState::Step::R4A02_RunConcurrentSameDestinationStaysSafe,
    SelfTestState::Step::R4A02_LiveOutputGuardChoices,
    SelfTestState::Step::R4A02_DisjointWritesDoNotWarn,
    SelfTestState::Step::R4A02_RenamePublishKeepsIndexPrecise,
    SelfTestState::Step::R4A02_LatePublisherStillWarns,
}};

constexpr std::array<SelfTestState::Step, 7> kFileOpsFamilyPhase11{{
    SelfTestState::Step::Phase11_CrossFileSystemBridge,
    SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines,
    SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines,
    SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf,
    SelfTestState::Step::Phase11_ConnectionOverridePrecedence,
    SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate,
    SelfTestState::Step::Phase11_ConnectionOverrideClamp,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPhase12{{
    SelfTestState::Step::Phase12_ReparsePointPolicy,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPhase13{{
    SelfTestState::Step::Phase13_PostMortemDiagnostics,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPhase14{{
    SelfTestState::Step::Phase14_PopupHostLifetimeGuard,
}};

constexpr std::array<SelfTestState::Step, 2> kFileOpsFamilyPhase15{{
    SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke,
    SelfTestState::Step::Phase15_FileSystem7zMountPathImpact,
}};

constexpr std::array<SelfTestState::Step, 19> kFileOpsFamilyPhase16{{
    SelfTestState::Step::Phase16_RemoteWatchContractExposure,
    SelfTestState::Step::Phase16_RemoteFtpSecret,
    SelfTestState::Step::Phase16_RemoteFtpSandbox,
    SelfTestState::Step::Phase16_RemoteSftpSecret,
    SelfTestState::Step::Phase16_RemoteSftpSandbox,
    SelfTestState::Step::Phase16_RemoteScpSecret,
    SelfTestState::Step::Phase16_RemoteScpSandbox,
    SelfTestState::Step::Phase16_RemoteImapSecret,
    SelfTestState::Step::Phase16_RemoteImapSandbox,
    SelfTestState::Step::Phase16_RemoteS3Secret,
    SelfTestState::Step::Phase16_RemoteS3Sandbox,
    SelfTestState::Step::Phase16_RemoteS3FileOps,
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret,
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSandbox,
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalFileOps,
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret,
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSandbox,
    SelfTestState::Step::Phase16_RemoteSharePointSecret,
    SelfTestState::Step::Phase16_RemoteSharePointSandbox,
}};

struct FileOpsFamilyDefinition
{
    std::wstring_view name;
    std::span<const SelfTestState::Step> phases;
};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyQueuedAdmissionContracts{{
    SelfTestState::Step::FileOps_QueuedAdmissionContracts,
}};
constexpr std::array<SelfTestState::Step, 6> kFileOpsFamilySafetyRemediationContracts{{
    SelfTestState::Step::FileOps_SafetyRemediationContracts,
    SelfTestState::Step::FileOps_ConflictMessageText,
    SelfTestState::Step::FileOps_ReviewRemediationContracts,
    SelfTestState::Step::FileOps_RepeatDecisionContracts,
    SelfTestState::Step::FileOps_SkipCountingPresentation,
    SelfTestState::Step::FileOps_PermanentDeleteCompletionProgress,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyOwnedProductionShutdown{{
    SelfTestState::Step::FileOps_OwnedProductionShutdown,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyAdmissionBaseline{{
    SelfTestState::Step::FileOps_AdmissionBaseline,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyCrossVolumeMovePartialStorage{{
    SelfTestState::Step::FileOps_CrossVolumeMoveAfterPartialPublicationStorage,
}};
constexpr std::array<SelfTestState::Step, 2> kFileOpsFamilyRepeatedNonClipboardStorage{{
    SelfTestState::Step::FileOps_RepeatedNonClipboardCopyStorage,
    SelfTestState::Step::FileOps_RepeatedNonClipboardMoveStorage,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyTerminalFactsProjection{{
    SelfTestState::Step::FileOps_TerminalFactsProjection,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyClearCompletedScope{{
    SelfTestState::Step::FileOps_ClearCompletedScope,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPerItemCancelUndispatchedRoots{{
    SelfTestState::Step::FileOps_PerItemCancelUndispatchedRoots,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyAdmissionFailureReachesPane{{
    SelfTestState::Step::FileOps_AdmissionFailureReachesPane,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyDeferredExitRefusesNewWork{{
    SelfTestState::Step::FileOps_DeferredExitRefusesNewWork,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyCompareWindowCloseConsent{{
    SelfTestState::Step::FileOps_CompareWindowCloseConsent,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPopupFailureSurfaceRecovery{{
    SelfTestState::Step::FileOps_PopupFailureSurfaceRecovery,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPopupStableHostedTree{{
    SelfTestState::Step::FileOps_PopupStableHostedTree,
}};
constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPopupIdleCadence{{
    SelfTestState::Step::FileOps_PopupIdleCadence,
}};
constexpr std::array<SelfTestState::Step, 2> kFileOpsFamilyAlertOverlayLifetime{{
    SelfTestState::Step::FileOps_AlertOverlayLifetime,
    SelfTestState::Step::FileOps_AlertOverlayUxAccessibility,
}};
constexpr std::array<SelfTestState::Step, 7> kFileOpsFamilyDecisionMirror{{
    SelfTestState::Step::FileOps_AlertOverlayModelessDecision,
    SelfTestState::Step::FileOps_DecisionMirror,
    SelfTestState::Step::FileOps_DecisionMirrorSameHostOverlap,
    SelfTestState::Step::FileOps_DecisionMirrorIngressAnchor,
    SelfTestState::Step::FileOps_DecisionMirrorCompareWindow,
    SelfTestState::Step::FileOps_DecisionMirrorConsecutivePrompts,
    SelfTestState::Step::FileOps_DecisionMirrorQueueOrder,
}};

constexpr std::array<SelfTestState::Step, 2> kFileOpsFamilyBatchRenameExactResult{{
    SelfTestState::Step::FileOps_BatchRenameExactResultCarrier,
    SelfTestState::Step::FileOps_BatchRenameNeverRunRowsNotAttempted,
}};

constexpr std::array<SelfTestState::Step, 2> kFileOpsFamilyNotes{{
    SelfTestState::Step::FileOps_NotesSeverity,
    SelfTestState::Step::FileOps_NotesMetadata,
}};

constexpr std::array<SelfTestState::Step, 2> kFileOpsFamilySelectionShape{{
    SelfTestState::Step::FileOps_SelectionShapeLocal,
    SelfTestState::Step::FileOps_SelectionShapeSmb,
}};

constexpr std::array<SelfTestState::Step, 6> kFileOpsFamilyCrossProviderRoutes{{
    SelfTestState::Step::FileOps_CrossProviderMtpNewFile,
    SelfTestState::Step::FileOps_VerifyAvailability,
    SelfTestState::Step::FileOps_SizelessSourceNote,
    SelfTestState::Step::FileOps_CrossRootLocalCopy,
    SelfTestState::Step::FileOps_ArchiveNestedCopyLocal,
    SelfTestState::Step::FileOps_ArchiveNestedCopyUnc,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyCancelAll{{
    SelfTestState::Step::FileOps_CancelAllNeverStarted,
}};

constexpr std::array<SelfTestState::Step, 8> kFileOpsFamilyThroughput{{
    SelfTestState::Step::FileOps_SmbCopyThroughputVariance,
    SelfTestState::Step::FileOps_LocalCopyThroughput,
    SelfTestState::Step::FileOps_SmbNativeRouteConcurrencyBudget,
    SelfTestState::Step::FileOps_SameProviderBridgeDestinationStorage,
    SelfTestState::Step::FileOps_SmbManagedMoveLeavesNoStage,
    SelfTestState::Step::FileOps_RateWindowDisplay,
    SelfTestState::Step::FileOps_ProgressPhaseAccounting,
    SelfTestState::Step::FileOps_StallLabelLayout,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilySmallFileCost{{
    SelfTestState::Step::FileOps_SmallFileCopyCost,
}};

constexpr std::array<SelfTestState::Step, 21> kFileOpsFamilyIdentityLess{{
    SelfTestState::Step::Delete_IdentityLessLocationByName,
    SelfTestState::Step::Delete_IdentityLessSharingAndInterlock,
    SelfTestState::Step::Delete_IdentityLessQueuedRootOverlap,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeOtherSpelling,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeArchiveWriter,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeCardPending,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeConcurrentWriter,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeBudget,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeChangedEntry,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeEscalationPartial,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeLinkSwap,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeSubfolderOpen,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeRecycleOverBudget,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeHolders,
    SelfTestState::Step::Delete_IdentityLessPinnedTreeScale,
    SelfTestState::Step::Delete_IdentityLessFolderChangeAndHolders,
    SelfTestState::Step::Delete_IdentityLessOpenFailureAndArchiveOrigins,
    SelfTestState::Step::FileOps_IdentityLessGuardPartition,
    SelfTestState::Step::FileOps_LeftoverCleanupWalkBounds,
    SelfTestState::Step::FileOps_OwnedStageCleanupBudget,
    SelfTestState::Step::Delete_NasFileIdInfoUnsupportedByName,
}};

constexpr auto kFileOpsFamilyDefinitions = std::to_array<FileOpsFamilyDefinition>({
    FileOpsFamilyDefinition{L"FileOpsFamily_QueuedAdmissionContracts", {kFileOpsFamilyQueuedAdmissionContracts}},
    FileOpsFamilyDefinition{L"FileOpsFamily_SafetyRemediationContracts", {kFileOpsFamilySafetyRemediationContracts}},
    FileOpsFamilyDefinition{L"FileOpsFamily_AdmissionBaseline", {kFileOpsFamilyAdmissionBaseline}},
    FileOpsFamilyDefinition{L"FileOpsFamily_OwnedProductionShutdown", {kFileOpsFamilyOwnedProductionShutdown}},
    FileOpsFamilyDefinition{L"FileOpsFamily_CrossVolumeMovePartialStorage", {kFileOpsFamilyCrossVolumeMovePartialStorage}},
    FileOpsFamilyDefinition{L"FileOpsFamily_RepeatedNonClipboardStorage", {kFileOpsFamilyRepeatedNonClipboardStorage}},
    FileOpsFamilyDefinition{L"FileOpsFamily_TerminalFactsProjection", {kFileOpsFamilyTerminalFactsProjection}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearCompletedScope", {kFileOpsFamilyClearCompletedScope}},
    FileOpsFamilyDefinition{L"FileOpsFamily_PerItemCancelUndispatchedRoots", {kFileOpsFamilyPerItemCancelUndispatchedRoots}},
    FileOpsFamilyDefinition{L"FileOpsFamily_AdmissionFailureReachesPane", {kFileOpsFamilyAdmissionFailureReachesPane}},
    FileOpsFamilyDefinition{L"FileOpsFamily_DeferredExitRefusesNewWork", {kFileOpsFamilyDeferredExitRefusesNewWork}},
    FileOpsFamilyDefinition{L"FileOpsFamily_CompareWindowCloseConsent", {kFileOpsFamilyCompareWindowCloseConsent}},
    FileOpsFamilyDefinition{L"FileOpsFamily_PopupFailureSurfaceRecovery", {kFileOpsFamilyPopupFailureSurfaceRecovery}},
    FileOpsFamilyDefinition{L"FileOpsFamily_PopupStableHostedTree", {kFileOpsFamilyPopupStableHostedTree}},
    FileOpsFamilyDefinition{L"FileOpsFamily_PopupIdleCadence", {kFileOpsFamilyPopupIdleCadence}},
    FileOpsFamilyDefinition{L"FileOpsFamily_AlertOverlayLifetime", {kFileOpsFamilyAlertOverlayLifetime}},
    FileOpsFamilyDefinition{L"FileOpsFamily_DecisionMirror", {kFileOpsFamilyDecisionMirror}},
    FileOpsFamilyDefinition{L"FileOpsFamily_BatchRenameExactResult", {kFileOpsFamilyBatchRenameExactResult}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Notes", {kFileOpsFamilyNotes}},
    FileOpsFamilyDefinition{L"FileOpsFamily_SelectionShape", {kFileOpsFamilySelectionShape}},
    FileOpsFamilyDefinition{L"FileOpsFamily_CrossProviderRoutes", {kFileOpsFamilyCrossProviderRoutes}},
    FileOpsFamilyDefinition{L"FileOpsFamily_CancelAll", {kFileOpsFamilyCancelAll}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Throughput", {kFileOpsFamilyThroughput}},
    FileOpsFamilyDefinition{L"FileOpsFamily_SmallFileCost", {kFileOpsFamilySmallFileCost}},
    FileOpsFamilyDefinition{L"FileOpsFamily_IdentityLess", {kFileOpsFamilyIdentityLess}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase01_DirectoryMerge", {kFileOpsFamilyClearflowPhase01}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase03_GraphFairness", {kFileOpsFamilyClearflowPhase03}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase04_Security", {kFileOpsFamilyClearflowPhase04}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Fairstream", {kFileOpsFamilyFairstream}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase05_PerfEvaluation", {kFileOpsFamilyClearflowPhase05}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase07_ProviderMatrix", {kFileOpsFamilyClearflowPhase07}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase05_Discovery", {kFileOpsFamilyPhase05}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R4A19DiscoveryBaseline", {kFileOpsFamilyR4A19DiscoveryBaseline}},
    FileOpsFamilyDefinition{L"FileOpsFamily_DiscoveryScope", {kFileOpsFamilyDiscoveryScope}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R0fSmbContainment", {kFileOpsFamilyR0fSmbContainment}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R0fCurlContainment", {kFileOpsFamilyR0fCurlContainment}},
    FileOpsFamilyDefinition{L"FileOpsFamily_C0MutationReceipts", {kFileOpsFamilyC0MutationReceipts}},
    FileOpsFamilyDefinition{L"FileOpsFamily_C0CurlTruth", {kFileOpsFamilyC0CurlTruth}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R0fS3Containment", {kFileOpsFamilyR0fS3Containment}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R0fGraphContainment", {kFileOpsFamilyR0fGraphContainment}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R0fGDriveContainment", {kFileOpsFamilyR0fGDriveContainment}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R3OwnedPublication", {kFileOpsFamilyR3OwnedPublication}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R4Traversal", {kFileOpsFamilyR4Traversal}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R4Overlap", {kFileOpsFamilyR4Overlap}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase06_PopupAndDelete", {kFileOpsFamilyPhase06}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase07_WatchAndParallelism", {kFileOpsFamilyPhase07}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase08_Validation", {kFileOpsFamilyPhase08}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase09_ConflictPrompt", {kFileOpsFamilyPhase09}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase10_DeleteValidation", {kFileOpsFamilyPhase10}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R1dPreparingLifecycle", {kFileOpsFamilyR1dPreparing}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase11_BridgeAndConnections", {kFileOpsFamilyPhase11}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase12_Reparse", {kFileOpsFamilyPhase12}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase13_PostMortem", {kFileOpsFamilyPhase13}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase14_PopupLifetime", {kFileOpsFamilyPhase14}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase15_FileSystem7z", {kFileOpsFamilyPhase15}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase16_Remote", {kFileOpsFamilyPhase16}},
});

static_assert(
    []
{
    // Broad runs dispatch child families, while reporting this default phase order.
    // A phase registered in only one list would silently go unexecuted until the
    // parent runner rejects its unexpected skip. Overlapping existing families are valid.
    std::array<bool, static_cast<size_t>(SelfTestState::Step::Failed) + 1u> covered{};
    for (const auto& family : kFileOpsFamilyDefinitions)
    {
        for (const SelfTestState::Step step : family.phases)
            covered.at(static_cast<size_t>(step)) = true;
    }
    for (const SelfTestState::Step step : kFileOpsPhaseOrder)
    {
        if (step != SelfTestState::Step::Setup && step != SelfTestState::Step::Cleanup_RestorePluginConfig && ! covered.at(static_cast<size_t>(step)))
            return false;
    }
    return true;
}(),
    "Every default FileOps phase must be dispatched by a child family.");

struct RunSelection
{
    bool recognized = true;
    std::vector<SelfTestState::Step> activePhases;
    std::vector<SelfTestState::Step> reportedPhases;
};

[[nodiscard]] bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept;
[[nodiscard]] bool StartsWithIgnoreCase(std::wstring_view text, std::wstring_view prefix) noexcept;

[[nodiscard]] bool IsPhaseSelected(const SelfTestState& state, SelfTestState::Step step) noexcept
{
    return std::find(state.activePhaseOrder.begin(), state.activePhaseOrder.end(), step) != state.activePhaseOrder.end();
}

[[nodiscard]] std::optional<SelfTestState::Step> NextSelectedPhaseAfterCurrent(const SelfTestState& state) noexcept
{
    const auto current = std::find(state.activePhaseOrder.begin(), state.activePhaseOrder.end(), state.step);
    if (current == state.activePhaseOrder.end())
    {
        return std::nullopt;
    }

    const auto next = std::next(current);
    if (next == state.activePhaseOrder.end())
    {
        return std::nullopt;
    }

    return *next;
}

[[nodiscard]] const FileOpsFamilyDefinition* FindFamilyByName(std::wstring_view name) noexcept
{
    for (const FileOpsFamilyDefinition& family : kFileOpsFamilyDefinitions)
    {
        if (EqualsIgnoreCase(name, family.name))
        {
            return &family;
        }
    }

    return nullptr;
}

[[nodiscard]] std::optional<SelfTestState::Step> FindPhaseByName(std::wstring_view name) noexcept
{
    for (const SelfTestState::Step step : kFileOpsPhaseOrder)
    {
        if (step == SelfTestState::Step::Setup || step == SelfTestState::Step::Cleanup_RestorePluginConfig)
        {
            continue;
        }

        if (EqualsIgnoreCase(name, StepToString(step)))
        {
            return step;
        }
    }

    return std::nullopt;
}

[[nodiscard]] std::vector<SelfTestState::Step> FindPhasesByPrefix(std::wstring_view prefix)
{
    std::vector<SelfTestState::Step> matches;
    if (prefix.empty())
    {
        return matches;
    }

    for (const SelfTestState::Step step : kFileOpsPhaseOrder)
    {
        if (step == SelfTestState::Step::Setup || step == SelfTestState::Step::Cleanup_RestorePluginConfig)
        {
            continue;
        }

        if (StartsWithIgnoreCase(StepToString(step), prefix))
        {
            matches.push_back(step);
        }
    }

    return matches;
}

[[nodiscard]] RunSelection ResolveRunSelection(std::wstring_view filter)
{
    RunSelection selection{};
    if (EqualsIgnoreCase(filter, L"FileOps_CommonPresentationPerf"))
    {
        selection.reportedPhases = {SelfTestState::Step::FileOps_CommonPresentationPerf};
        selection.activePhases   = selection.reportedPhases;
        return selection;
    }
    if (EqualsIgnoreCase(filter, L"FileOps_IntegratedPresentationPerf"))
    {
        selection.reportedPhases = {SelfTestState::Step::FileOps_IntegratedPresentationPerf};
        selection.activePhases   = selection.reportedPhases;
        return selection;
    }
    // Opt-in documentation fixture; not added to broad runtime/performance qualification.
    if (EqualsIgnoreCase(filter, L"FileOps_VisualGallery") || EqualsIgnoreCase(filter, L"FileOps_VisualGalleryInteraction"))
    {
        selection.reportedPhases = {EqualsIgnoreCase(filter, L"FileOps_VisualGalleryInteraction") ? SelfTestState::Step::FileOps_VisualGalleryInteraction
                                                                                                  : SelfTestState::Step::FileOps_VisualGallery};
        selection.activePhases   = selection.reportedPhases;
        return selection;
    }

    if (filter.empty())
    {
        selection.reportedPhases.reserve(kFileOpsPhaseOrder.size());
        for (const SelfTestState::Step step : kFileOpsPhaseOrder)
        {
            selection.reportedPhases.push_back(step);
            if (step != SelfTestState::Step::Setup && step != SelfTestState::Step::Cleanup_RestorePluginConfig)
            {
                selection.activePhases.push_back(step);
            }
        }
        return selection;
    }

    if (const FileOpsFamilyDefinition* family = FindFamilyByName(filter))
    {
        selection.reportedPhases.push_back(SelfTestState::Step::Setup);
        for (const SelfTestState::Step step : family->phases)
        {
            selection.reportedPhases.push_back(step);
            selection.activePhases.push_back(step);
        }
        selection.reportedPhases.push_back(SelfTestState::Step::Cleanup_RestorePluginConfig);
        return selection;
    }

    if (const std::optional<SelfTestState::Step> phase = FindPhaseByName(filter); phase.has_value())
    {
        if (phase.value() == SelfTestState::Step::Phase5_SwitchWaitToParallelResume)
        {
            // The resume assertion consumes the queue-paused task deliberately established by the
            // immediately preceding transition phase. Keep exact-case retries deterministic by
            // scheduling that setup phase explicitly instead of relying on a previous process run.
            selection.reportedPhases = {SelfTestState::Step::Setup,
                                        SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery,
                                        phase.value(),
                                        SelfTestState::Step::Cleanup_RestorePluginConfig};
            selection.activePhases   = {SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery, phase.value()};
            return selection;
        }

        selection.reportedPhases = {SelfTestState::Step::Setup, phase.value(), SelfTestState::Step::Cleanup_RestorePluginConfig};
        selection.activePhases   = {phase.value()};
        return selection;
    }

    std::vector<SelfTestState::Step> prefixMatches = FindPhasesByPrefix(filter);
    if (! prefixMatches.empty())
    {
        selection.reportedPhases.push_back(SelfTestState::Step::Setup);
        for (const SelfTestState::Step step : prefixMatches)
        {
            selection.reportedPhases.push_back(step);
            selection.activePhases.push_back(step);
        }
        selection.reportedPhases.push_back(SelfTestState::Step::Cleanup_RestorePluginConfig);
        return selection;
    }

    selection.recognized = false;
    return selection;
}

[[nodiscard]] std::vector<std::wstring> BuildRunFiltersImpl(std::wstring_view filter)
{
    std::vector<std::wstring> filters;
    if (filter.empty())
    {
        filters.reserve(kFileOpsFamilyDefinitions.size());
        for (const FileOpsFamilyDefinition& family : kFileOpsFamilyDefinitions)
        {
            filters.emplace_back(family.name);
        }
        return filters;
    }

    if (ResolveRunSelection(filter).recognized)
    {
        filters.emplace_back(filter);
    }

    return filters;
}

void AppendLog(std::wstring_view message) noexcept
{
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, std::format(L"[{}] {}", GetTickCount64(), message));
    SelfTest::AppendSelfTestTrace(std::format(L"[{}] {}", GetTickCount64(), message));
}

void RecordCurrentPhase(SelfTestState& state, SelfTest::SelfTestCaseResult::Status status, std::wstring_view reason = {}) noexcept
{
    if (! state.phaseInProgress || state.phaseName.empty())
    {
        return;
    }

    const auto now            = GetTickCount64();
    const uint64_t durationMs = (now >= state.phaseStartTick) ? (now - state.phaseStartTick) : 0;

    SelfTest::SelfTestCaseResult item{};
    item.name       = state.phaseName;
    item.status     = status;
    item.durationMs = durationMs;
    if (! reason.empty())
    {
        item.reason = reason;
    }

    // Every phase boundary, failed phases included, purges the Recycle Bin items the phase created;
    // any it leaves fails the phase and the run.
    uint64_t purgedRecycledItems = 0u;
    if (std::wstring recycleLeak = SelfTest::PurgeRecycledTestItems(&purgedRecycledItems); ! recycleLeak.empty())
    {
        AppendLog(std::format(L"{}: {}", state.phaseName, recycleLeak));
        item.status = SelfTest::SelfTestCaseResult::Status::failed;
        item.reason = item.reason.empty() ? recycleLeak : std::format(L"{} {}", item.reason, recycleLeak);
        if (state.failureMessage.empty())
        {
            state.failureMessage = std::move(recycleLeak);
        }
        state.failed.store(true, std::memory_order_release);
    }
    else if (purgedRecycledItems != 0u)
    {
        AppendLog(std::format(L"{}: purged {} recycled test items from the Recycle Bin", state.phaseName, purgedRecycledItems));
    }

    state.phaseResults.push_back(std::move(item));
    state.phaseInProgress = false;
    state.phaseName.clear();
    state.phaseStartTick = 0;
    state.phaseFailed    = false;
    state.phaseFailureMessage.clear();
}

// Purges the Recycle Bin items a finished recycle task created and requires exactly expectedCount of
// them, so a progress sink that stopped journaling fails the phase instead of leaking its items
// past the boundary purge. Returns an empty string on success.
[[nodiscard]] std::wstring PurgeExpectedRecycledTestItems(uint64_t expectedCount) noexcept
{
    uint64_t purged        = 0u;
    std::wstring purgeLeak = SelfTest::PurgeRecycledTestItems(&purged);
    if (! purgeLeak.empty())
    {
        return purgeLeak;
    }
    if (purged != expectedCount)
    {
        return std::format(L"the Local provider journaled {} recycled test items for purging; expected {}.", purged, expectedCount);
    }
    return {};
}

// This asynchronous fixture spans presentation ticks. Restore only admission of
// future tasks, including on failure/cancellation; never reschedule live tasks or
// persist a preference merely to qualify the independent-volume control.
void RestoreIndependentVolumeQueueMode(SelfTestState& state) noexcept
{
    if (! state.r4A19IndependentQueueModeOriginal.has_value())
        return;
    const bool original = state.r4A19IndependentQueueModeOriginal.value();
    state.r4A19IndependentQueueModeOriginal.reset();
    if (state.fileOps)
        state.fileOps->SetQueueNewTasks(original);
}

void BeginPhase(SelfTestState& state, SelfTestState::Step step) noexcept
{
    RestoreIndependentVolumeQueueMode(state);
    if (state.phaseInProgress)
    {
        RecordCurrentPhase(state, state.phaseFailed ? SelfTest::SelfTestCaseResult::Status::failed : SelfTest::SelfTestCaseResult::Status::passed);
    }

    if (step == SelfTestState::Step::Done || step == SelfTestState::Step::Failed || step == SelfTestState::Step::Idle)
    {
        return;
    }

    state.phaseInProgress = true;
    state.phaseStartTick  = GetTickCount64();
    state.phaseFailed     = false;
    state.phaseFailureMessage.clear();
    state.phaseName = StepToString(step);
}

void NextStep(SelfTestState& state, SelfTestState::Step next) noexcept
{
    SelfTestState::Step resolved = next;
    if (next != SelfTestState::Step::Setup && next != SelfTestState::Step::Cleanup_RestorePluginConfig && next != SelfTestState::Step::Done &&
        next != SelfTestState::Step::Failed && next != SelfTestState::Step::Idle)
    {
        if (state.activePhaseOrder.empty())
        {
            resolved = SelfTestState::Step::Cleanup_RestorePluginConfig;
        }
        else if (! IsPhaseSelected(state, next))
        {
            if (state.step == SelfTestState::Step::Setup)
            {
                resolved = state.activePhaseOrder.front();
            }
            else if (const std::optional<SelfTestState::Step> selectedNext = NextSelectedPhaseAfterCurrent(state); selectedNext.has_value())
            {
                resolved = selectedNext.value();
            }
            else
            {
                resolved = SelfTestState::Step::Cleanup_RestorePluginConfig;
            }
        }
    }

    AppendLog(std::format(L"NextStep: {}", StepToString(resolved)));
    SplashScreen::IfExistSetText(std::format(L"Self-test: {}", StepToString(resolved)));
    BeginPhase(state, resolved);
    state.step          = resolved;
    state.stepStartTick = GetTickCount64();
    state.stepState     = 0;
    state.markerTick    = 0;
}

bool HasTimedOut(const SelfTestState& state, ULONGLONG nowTick, ULONGLONG timeoutMs = kDefaultTimeoutMs) noexcept
{
    return nowTick >= state.stepStartTick && (nowTick - state.stepStartTick) > SelfTest::ScaleTimeout(timeoutMs);
}

[[nodiscard]] std::wstring NormalizePluginPathForSelfTest(std::wstring_view rawPath) noexcept;
void CleanupRemoteOneDrivePersonalCase(SelfTestState& state) noexcept;
void CleanupRemoteS3Case(SelfTestState& state) noexcept;
void PruneEmptyAlternateVolumeSandboxParents(const std::filesystem::path& sandboxRoot) noexcept;
[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindLoadedPluginEntry(std::wstring_view pluginId) noexcept;

[[nodiscard]] bool PerformCleanup(SelfTestState& state) noexcept;

void Fail(std::wstring_view message) noexcept
{
    SelfTestState& state = GetState();
    if (state.done.load(std::memory_order_acquire))
    {
        return;
    }

    state.failureMessage = std::wstring(message);
    state.phaseFailed    = true;
    if (state.phaseFailureMessage.empty())
    {
        state.phaseFailureMessage = message;
    }
    state.failed.store(true, std::memory_order_release);
    AppendLog(std::format(L"FAIL: {}", state.failureMessage));
    Debug::Error(L"FileOpsSelfTest FAILED: {}", state.failureMessage);

    // Record the current phase as failed, then run cleanup immediately. Many self-test call sites do
    // `Fail(...); return true;` which would otherwise short-circuit the FSM and skip cleanup.
    RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::failed, state.failureMessage);
    BeginPhase(state, SelfTestState::Step::Cleanup_RestorePluginConfig);
    const bool cleaned = PerformCleanup(state);
    RecordCurrentPhase(state,
                       cleaned ? SelfTest::SelfTestCaseResult::Status::passed : SelfTest::SelfTestCaseResult::Status::failed,
                       cleaned ? L"" : L"Outstanding operations did not drain; provider settings and scratch were preserved.");

    state.step = SelfTestState::Step::Done;
    state.running.store(false, std::memory_order_release);
    state.done.store(true, std::memory_order_release);
}

FolderWindow* TryGetFolderWindow(HWND mainWindow) noexcept
{
    if (! mainWindow)
    {
        return nullptr;
    }

    const HWND folderWindowHwnd = FindWindowExW(mainWindow, nullptr, kFolderWindowClassName.data(), nullptr);
    if (! folderWindowHwnd)
    {
        return nullptr;
    }

    return reinterpret_cast<FolderWindow*>(GetWindowLongPtrW(folderWindowHwnd, GWLP_USERDATA));
}

FolderWindow::FileOperationState* TryGetFileOps(FolderWindow* folderWindow) noexcept
{
    if (! folderWindow)
    {
        return nullptr;
    }

    return folderWindow->DebugGetFileOperationState();
}

FolderView* TryGetFolderView(FolderWindow* folderWindow, FolderWindow::Pane pane) noexcept
{
    if (! folderWindow)
    {
        return nullptr;
    }

    const HWND hwnd = folderWindow->GetFolderViewHwnd(pane);
    if (! hwnd)
    {
        return nullptr;
    }

    return reinterpret_cast<FolderView*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
}

bool BackupPluginConfiguration(IInformations* info, std::string& outConfigUtf8) noexcept
{
    if (! info)
    {
        return false;
    }

    const char* config = nullptr;
    const HRESULT hr   = info->GetConfiguration(&config);
    if (FAILED(hr) || ! config)
    {
        return false;
    }

    outConfigUtf8 = config;
    return true;
}

bool SetPluginConfiguration(IInformations* info, std::string_view configUtf8) noexcept
{
    if (! info)
    {
        return false;
    }

    Common::Settings::JsonValue parsedConfiguration;
    if (FAILED(Common::Settings::ParseJsonValue(configUtf8, parsedConfiguration)))
    {
        return false;
    }

    std::string owned(configUtf8);
    owned.push_back('\0');
    const HRESULT hr = info->SetConfiguration(owned.c_str());
    if (SUCCEEDED(hr))
    {
        SelfTestState& state = GetState();
        if (state.running.load(std::memory_order_acquire))
        {
            if (info == state.infoLocal.get())
            {
                state.localConfigDirty                                                   = true;
                g_settings.plugins.configurationByPluginId[std::wstring(kPluginIdLocal)] = std::move(parsedConfiguration);
            }
            else if (info == state.infoDummy.get())
            {
                state.dummyConfigDirty                                                   = true;
                g_settings.plugins.configurationByPluginId[std::wstring(kPluginIdDummy)] = std::move(parsedConfiguration);
            }
            else if (info == state.info7z.get())
            {
                state.config7zDirty                                                   = true;
                g_settings.plugins.configurationByPluginId[std::wstring(kPluginId7z)] = std::move(parsedConfiguration);
            }
        }
    }
    return SUCCEEDED(hr);
}

std::optional<std::string> TryGetPropertyFieldValue(IFileSystemIO* io, const wchar_t* path, std::string_view key) noexcept
{
    if (! io || ! path || key.empty())
    {
        return std::nullopt;
    }

    const char* jsonUtf8 = nullptr;
    const HRESULT hr     = io->GetItemProperties(path, &jsonUtf8);
    if (FAILED(hr) || ! jsonUtf8)
    {
        return std::nullopt;
    }

    std::string_view json(jsonUtf8);
    const std::string needle = std::format("\"key\":\"{}\"", key);
    const size_t keyPos      = json.find(needle);
    if (keyPos == std::string_view::npos)
    {
        return std::nullopt;
    }

    const size_t valuePos = json.find("\"value\":\"", keyPos + needle.size());
    if (valuePos == std::string_view::npos)
    {
        return std::nullopt;
    }

    const size_t begin = valuePos + std::string_view("\"value\":\"").size();
    const size_t end   = json.find('"', begin);
    if (end == std::string_view::npos || end < begin)
    {
        return std::nullopt;
    }

    return std::string(json.substr(begin, end - begin));
}

[[maybe_unused]] std::optional<uint32_t> TryGetPropertyFieldUInt32(IFileSystemIO* io, const wchar_t* path, std::string_view key) noexcept
{
    const auto valueOpt = TryGetPropertyFieldValue(io, path, key);
    if (! valueOpt.has_value())
    {
        return std::nullopt;
    }

    uint32_t parsed                   = 0;
    const char* begin                 = valueOpt->c_str();
    const char* end                   = begin + valueOpt->size();
    const std::from_chars_result conv = std::from_chars(begin, end, parsed);
    if (conv.ec != std::errc{} || conv.ptr != end)
    {
        return std::nullopt;
    }

    return parsed;
}

// This test-only state owns every operation in its app process. Quiesce those workers before
// changing provider configuration or scratch; a failed family must not poison the next family.
[[nodiscard]] bool CancelOutstandingSelfTestOperations(FolderWindow::FileOperationState* fileOps) noexcept
{
    if (! fileOps || ! fileOps->HasActiveOperations())
    {
        return true;
    }

    fileOps->CancelAll();
    const auto drained = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return ! fileOps->HasActiveOperations();
    }, {.timeout = SelfTest::Scale(std::chrono::milliseconds(30'000)), .operationName = L"FileOps family cancellation and completion drain"});
    if (! drained.conditionMet)
    {
        AppendLog(drained.timeoutDiagnostic);
        return false;
    }
    AppendLog(std::format(L"FileOps family cancellation drained in {} ms.", drained.elapsed.count()));
    return true;
}

[[nodiscard]] bool PerformCleanup(SelfTestState& state) noexcept
{
    RestoreIndependentVolumeQueueMode(state);
    AppendLog(L"PerformCleanup: begin");
    SetFileOpsBridgeProducerDelayForSelfTest(0);
    SetFileOpsBridgeTraversalDepthLimitForSelfTest(0);
    SetFileOpsBridgePipelineModeForSelfTest(FileOpsBridgePipelineMode::Default);
    SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest());
    SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
    SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeFailNextDestinationGetSizeAttemptsForSelfTest());
    SetFileOpsBridgeFailNextStageEntropyForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeFailNextStageEntropyAttemptsForSelfTest());
    SetFileOpsBridgeFailNextDestinationOpenForSelfTest(0, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
    static_cast<void>(TakeFileOpsBridgeFailNextDestinationOpenAttemptsForSelfTest());
    SetFileOpsBridgeNullNextSourceReaderForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeNullNextSourceReaderAttemptsForSelfTest());
    SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeReportWrongDestinationSizeAttemptsForSelfTest());
    SetFileOpsBridgeOverReportNextReadForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeOverReportNextReadAttemptsForSelfTest());
    SetFileOpsBridgePrematureEofNextReadForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgePrematureEofNextReadAttemptsForSelfTest());
    SetFileOpsBridgeUnderConsumeNextWriteForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeUnderConsumeNextWriteAttemptsForSelfTest());
    SetFileOpsBridgeOverReportNextWriteForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeOverReportNextWriteAttemptsForSelfTest());
    SetFileOpsBridgeInjectHostileChildNamesForSelfTest(false);
    static_cast<void>(TakeFileOpsBridgeInjectHostileChildNameAttemptsForSelfTest());
    SetFileOpsBridgeInjectFileReparseForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeInjectFileReparseAttemptsForSelfTest());
    SetFileOpsBridgeReparsePolicyOverrideForSelfTest(FileOpsBridgeReparsePolicyOverride::None);
    SetFileOpsAutoConcurrencyOverrideForSelfTest(false, 1u, FILESYSTEM_STORAGE_UNKNOWN);
    ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
    SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
    ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
    SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
    SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(0u);
    static_cast<void>(TakeFileOpsNativeMoveCreateDirectoryRaceAttemptsForSelfTest());
    SetFileOpsManagedCleanupKnownNonCommitForSelfTest(HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), 0u);
    static_cast<void>(TakeFileOpsManagedCleanupKnownNonCommitAttemptsForSelfTest());
    SetFileOpsPermanentDeleteKnownNonCommitForSelfTest(HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), 0u);
    static_cast<void>(TakeFileOpsPermanentDeleteKnownNonCommitAttemptsForSelfTest());
    SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(0u);
    static_cast<void>(TakeFileOpsManagedCleanupUnknownOutcomeAttemptsForSelfTest());
    ReleaseFileOpsBridgePublishedDestinationRetryPauseForSelfTest();
    SetFileOpsBridgePublishedDestinationRetryPauseForSelfTest(false);
    static_cast<void>(TakeFileOpsBridgeReplacePublishedDestinationAttemptsForSelfTest());
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_BRIDGE_REPLACE_PUBLISHED_DESTINATION_PATH", nullptr));
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_BRIDGE_REPLACE_PUBLISHED_DESTINATION_PAYLOAD", nullptr));
    ReleaseFileOpsKeepBothNestedConflictPauseForSelfTest();
    SetFileOpsKeepBothNestedConflictPauseForSelfTest(false);
    ReleaseFileOpsBatchRenameBeforeMutationPauseForSelfTest();
    SetFileOpsBatchRenameBeforeMutationPauseForSelfTest(false);
    SetFileOpsBatchRenameReceiptFaultForSelfTest(FileOpsBatchRenameReceiptFaultForSelfTest::None);
    SetFileOpsRecycleDeleteReceiptFaultForSelfTest(FileOpsRecycleDeleteReceiptFaultForSelfTest::None);
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_RECYCLE_FAIL_PATH", nullptr));
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_RECYCLE_UNKNOWN_PATH", nullptr));
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_RECYCLE_UNKNOWN_HOLD", nullptr));
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_RECYCLE_DESTROY_PATH", nullptr));
    ReleaseFileOpsPermanentDeleteBeforeBindPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeBindPauseForSelfTest(false);
    ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
    ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
    SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_CASE_RENAME_ENTROPY_FAIL_PATH", nullptr));
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_PATH", nullptr));
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_FIRED", nullptr));

    // The host timer guards against reentrant Tick calls while the shared message pump dispatches
    // completion notifications. All injected worker barriers above are released before cancellation.
    if (! CancelOutstandingSelfTestOperations(state.fileOps))
    {
        state.failed.store(true, std::memory_order_release);
        state.phaseFailed = true;
        if (state.failureMessage.empty())
        {
            state.failureMessage = L"FileOps cleanup could not drain outstanding operations.";
        }
        return false;
    }
    if (state.batchRenameExactAutoAcceptBefore.has_value())
    {
        HostResetPromptShutdown();
        HostSetAutoAcceptPrompts(state.batchRenameExactAutoAcceptBefore.value());
        state.batchRenameExactAutoAcceptBefore.reset();
    }
    state.batchRenameExactPaneFileSystem.reset();
    state.batchRenameExactCurrentTaskId = 0u;

    if (state.fileOps)
    {
        AppendLog(L"PerformCleanup: restore auto-dismiss");
        state.fileOps->SetAutoDismissSuccess(state.autoDismissSuccessOriginal);
        AppendLog(L"PerformCleanup: restore auto-dismiss done");
    }

    if (state.fileOperationsBackedUp)
    {
        AppendLog(L"PerformCleanup: restoring file-operations settings snapshot");
        g_settings.fileOperations = state.fileOperationsOriginal;
        AppendLog(L"PerformCleanup: restored file-operations settings snapshot");
    }

    if (state.appThemeBackedUp && state.folderWindow && state.appThemeOriginal.has_value())
    {
        AppendLog(L"PerformCleanup: restoring app theme snapshot");
        state.folderWindow->ApplyTheme(state.appThemeOriginal.value());
        state.appThemeOriginal.reset();
        state.appThemeBackedUp = false;
        AppendLog(L"PerformCleanup: restored app theme snapshot");
    }

    AppendLog(L"PerformCleanup: restore plugin/config state");
    const auto restorePluginConfigIfNeeded = [&](const wchar_t* name, IInformations* info, const std::string& originalConfig, bool& dirtyFlag) noexcept
    {
        if (! info || originalConfig.empty())
        {
            return;
        }

        if (! dirtyFlag)
        {
            AppendLog(std::format(L"PerformCleanup: {} config was not mutated", name ? name : L"(plugin)"));
            return;
        }

        std::string currentConfig;
        if (! BackupPluginConfiguration(info, currentConfig))
        {
            AppendLog(std::format(L"PerformCleanup: {} current-config read failed", name ? name : L"(plugin)"));
            return;
        }

        if (currentConfig == originalConfig)
        {
            AppendLog(std::format(L"PerformCleanup: {} config already restored", name ? name : L"(plugin)"));
            dirtyFlag = false;
            return;
        }

        AppendLog(std::format(L"PerformCleanup: restoring {} config", name ? name : L"(plugin)"));
        static_cast<void>(SetPluginConfiguration(info, originalConfig));
        AppendLog(std::format(L"PerformCleanup: restored {} config", name ? name : L"(plugin)"));
        dirtyFlag = false;
    };

    restorePluginConfigIfNeeded(L"local", state.infoLocal.get(), state.localConfigOriginal, state.localConfigDirty);
    // Skip restoring FileSystemDummy between families. Each setup pass reseeds it explicitly and
    // repeated cleanup-time config reads/restores can hang during late aggregate-family transitions.
    state.dummyConfigDirty = false;
    AppendLog(L"PerformCleanup: dummy config restore skipped; next setup reseeds it deterministically");
    restorePluginConfigIfNeeded(L"7z", state.info7z.get(), state.config7zOriginal, state.config7zDirty);
    if (state.connectionsOriginal.has_value())
    {
        AppendLog(L"PerformCleanup: restoring connections snapshot");
        g_settings.connections = state.connectionsOriginal.value();
        AppendLog(L"PerformCleanup: restored connections snapshot");
    }
    if (state.bandwidthThrottleWorkerModeEnvBackedUp)
    {
        AppendLog(L"PerformCleanup: restoring bandwidth-throttle worker-mode env override");
        if (state.bandwidthThrottleWorkerModeEnvHadOriginal)
        {
            static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvBandwidthThrottleWorkerMode.data(), state.bandwidthThrottleWorkerModeEnvOriginal.c_str()));
        }
        else
        {
            static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvBandwidthThrottleWorkerMode.data(), nullptr));
        }
        state.bandwidthThrottleWorkerModeEnvBackedUp    = false;
        state.bandwidthThrottleWorkerModeEnvHadOriginal = false;
        state.bandwidthThrottleWorkerModeEnvOriginal.clear();
    }
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapTarget.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvMoveDestinationRacePath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvMoveDestinationRaceFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapMovedPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalAbortFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalAbortFailFired.data(), nullptr));
    // Persistent (not one-shot) injection: a case that failed before its own scope guard ran must not
    // leave every later bind under that root identity-less.
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvBindIdentityUnsupportedPath.data(), nullptr));
    AppendLog(L"PerformCleanup: restore plugin/config state done");

    AppendLog(L"PerformCleanup: remote cleanup");
    CleanupRemoteS3Case(state);
    CleanupRemoteOneDrivePersonalCase(state);

    AppendLog(L"PerformCleanup: release watches");
    state.directoryWatchCallback.reset();
    state.directoryWatch.reset();
    state.holdOpenHandle.reset();
    state.lockedFileHandle.reset();
    state.lockedDestinationHandle.reset();

    state.placeholderSyncRoot.reset();
    if (state.retainTempRootOnFailure && ! state.fileOpsAlternateVolumeRoot.empty())
    {
        AppendLog(L"PerformCleanup: alternate-volume storage/mount witness retained for review.");
        AppendLog(state.fileOpsAlternateVolumeRoot.native());
    }
    else if (! state.fileOpsAlternateVolumeRoot.empty())
    {
        AppendLog(std::format(L"PerformCleanup: delete alternate-volume root {}", state.fileOpsAlternateVolumeRoot.wstring()));
        std::error_code cleanupEc;
        const std::filesystem::path alternateRoot = std::move(state.fileOpsAlternateVolumeRoot);
        std::filesystem::remove_all(alternateRoot, cleanupEc);
        PruneEmptyAlternateVolumeSandboxParents(alternateRoot);
    }

    if (state.retainTempRootOnFailure && ! state.tempRoot.empty())
    {
        AppendLog(L"PerformCleanup: worker drain/settings restore completed; storage-proof scratch retained for review.");
        AppendLog(state.tempRoot.native());
    }
    else if (! state.tempRoot.empty())
    {
        AppendLog(std::format(L"PerformCleanup: delete temp root {}", state.tempRoot.wstring()));
        bool removed = false;
        for (int i = 0; i < 3; ++i)
        {
            removed = SelfTest::RemoveAll(state.tempRoot);
            if (removed)
            {
                break;
            }
            ::Sleep(100);
        }
        if (! removed)
        {
            Debug::Warning(L"FileOpsSelfTest: cleanup could not delete temp root: {}", state.tempRoot.wstring());
        }
    }

    state.r4A19MtpProfile.reset();
    AppendLog(L"PerformCleanup: defer COM/plugin release to process shutdown");
    AppendLog(L"PerformCleanup: done");
    return true;
}

bool LoadPlugins(SelfTestState& state) noexcept
{
    if (state.fsLocal && state.infoLocal && state.fsDummy && state.infoDummy && state.fs7z && state.info7z)
    {
        return true;
    }

    FileSystemPluginManager& mgr = FileSystemPluginManager::GetInstance();

    const auto tryAssignFromLoadedEntry = [&](std::wstring_view pluginId) noexcept
    {
        if (const FileSystemPluginManager::PluginEntry* entry = FindLoadedPluginEntry(pluginId))
        {
            if (EqualsIgnoreCase(pluginId, kPluginIdLocal))
            {
                state.fsLocal   = entry->fileSystem;
                state.infoLocal = entry->informations;
            }
            else if (EqualsIgnoreCase(pluginId, kPluginIdDummy))
            {
                state.fsDummy   = entry->fileSystem;
                state.infoDummy = entry->informations;
            }
            else if (EqualsIgnoreCase(pluginId, kPluginId7z))
            {
                state.fs7z   = entry->fileSystem;
                state.info7z = entry->informations;
            }
        }
    };

    tryAssignFromLoadedEntry(kPluginIdLocal);
    tryAssignFromLoadedEntry(kPluginIdDummy);
    tryAssignFromLoadedEntry(kPluginId7z);

    if (! state.fsLocal || ! state.infoLocal)
    {
        static_cast<void>(mgr.EnablePlugin(kPluginIdLocal, g_settings));
        tryAssignFromLoadedEntry(kPluginIdLocal);
    }
    if (! state.fsDummy || ! state.infoDummy)
    {
        static_cast<void>(mgr.EnablePlugin(kPluginIdDummy, g_settings));
        tryAssignFromLoadedEntry(kPluginIdDummy);
    }
    if (! state.fs7z || ! state.info7z)
    {
        static_cast<void>(mgr.EnablePlugin(kPluginId7z, g_settings));
        tryAssignFromLoadedEntry(kPluginId7z);
    }

    for (const auto& p : mgr.GetPlugins())
    {
        if (! state.fsLocal && p.id == kPluginIdLocal)
        {
            state.fsLocal   = p.fileSystem;
            state.infoLocal = p.informations;
        }
        else if (! state.fsDummy && p.id == kPluginIdDummy)
        {
            state.fsDummy   = p.fileSystem;
            state.infoDummy = p.informations;
        }
        else if (! state.fs7z && p.id == kPluginId7z)
        {
            state.fs7z   = p.fileSystem;
            state.info7z = p.informations;
        }
    }

    return state.fsLocal && state.infoLocal && state.fsDummy && state.infoDummy;
}

std::vector<std::wstring> ListDirectories(IFileSystem* fs, std::wstring_view path, size_t maxCount) noexcept
{
    std::vector<std::wstring> out;
    if (! fs)
    {
        return out;
    }

    wil::com_ptr<IFilesInformation> files;
    const std::wstring pathW(path);
    const HRESULT hr = fs->ReadDirectoryInfo(pathW.c_str(), files.put());
    if (FAILED(hr) || ! files)
    {
        return out;
    }

    FileInfo* head = nullptr;
    if (FAILED(files->GetBuffer(&head)) || ! head)
    {
        return out;
    }

    for (FileInfo* entry = head; entry;)
    {
        const bool isDirectory = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (isDirectory && entry->FileNameSize >= sizeof(wchar_t))
        {
            const size_t charCount = entry->FileNameSize / sizeof(wchar_t);
            std::wstring name(entry->FileName, entry->FileName + charCount);
            if (name != L"." && name != L"..")
            {
                out.push_back(std::move(name));
                if (out.size() >= maxCount)
                {
                    break;
                }
            }
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }

    return out;
}

size_t GetDirectoryEntryCount(IFileSystem* fs, std::wstring_view path) noexcept
{
    if (! fs)
    {
        return 0;
    }

    wil::com_ptr<IFilesInformation> files;
    const std::wstring pathW(path);
    const HRESULT hr = fs->ReadDirectoryInfo(pathW.c_str(), files.put());
    if (FAILED(hr) || ! files)
    {
        return 0;
    }

    unsigned long count = 0;
    if (FAILED(files->GetCount(&count)))
    {
        return 0;
    }

    return static_cast<size_t>(count);
}

uint64_t GetDirectoryImmediateFileBytes(IFileSystem* fs, std::wstring_view path) noexcept
{
    if (! fs)
    {
        return 0;
    }

    wil::com_ptr<IFilesInformation> files;
    const std::wstring pathW(path);
    const HRESULT hr = fs->ReadDirectoryInfo(pathW.c_str(), files.put());
    if (FAILED(hr) || ! files)
    {
        return 0;
    }

    FileInfo* head = nullptr;
    if (FAILED(files->GetBuffer(&head)) || ! head)
    {
        return 0;
    }

    uint64_t totalBytes = 0;
    for (FileInfo* entry = head; entry;)
    {
        const bool isDirectory = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (! isDirectory && entry->EndOfFile > 0)
        {
            totalBytes += static_cast<uint64_t>(entry->EndOfFile);
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }

    return totalBytes;
}

bool EnsureDummyFolderExists(IFileSystem* fs, std::wstring_view destinationFolder) noexcept
{
    if (! fs || destinationFolder.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemDirectoryOperations), dirOps.put_void());
    if (FAILED(hr) || ! dirOps)
    {
        AppendLog(std::format(
            L"EnsureDummyFolderExists missing IFileSystemDirectoryOperations folder={} hr=0x{:08X}", destinationFolder, static_cast<unsigned long>(hr)));
        return false;
    }

    const std::wstring destinationText(destinationFolder);
    hr            = dirOps->CreateDirectory(destinationText.c_str());
    const bool ok = SUCCEEDED(hr) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    if (! ok)
    {
        AppendLog(std::format(L"EnsureDummyFolderExists failed folder={} hr=0x{:08X}", destinationFolder, static_cast<unsigned long>(hr)));
    }
    return ok;
}

struct ProviderCapabilitySnapshot
{
    bool copyOperation            = false;
    bool moveOperation            = false;
    bool nativeMoveOperation      = false;
    bool deleteOperation          = false;
    bool renameOperation          = false;
    bool recycleOperation         = false;
    bool createDirectoryOperation = false;
    bool properties               = false;
    bool read                     = false;
    bool write                    = false;
    uint64_t copyMoveMax          = 0;
    uint64_t deleteMax            = 0;
    uint64_t deleteRecycleMax     = 0;
    bool exportCopyWildcard       = false;
    bool exportMoveWildcard       = false;
    bool importCopyWildcard       = false;
    bool importMoveWildcard       = false;
    bool preserveFileLink         = false;
    bool preserveDirectoryLink    = false;
    bool retargetInTree           = false;
    bool exactLinkRemoval         = false;
    bool pathIdentityPresent      = false;
    bool pathTextStableIdentity   = false;
    std::wstring componentComparison;
    wchar_t preferredSeparator = L'\0';
    std::wstring acceptedSeparators;
    bool casePreserving = false;
    std::wstring caseOnlyRename;
};

class FlagSensitiveAtomicWriter final : public IFileSystemAtomicWriter
{
public:
    explicit FlagSensitiveAtomicWriter(FileSystemFlags expectedFlags) noexcept : _expectedFlags(expectedFlags)
    {
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystemAtomicWriter))
        {
            *ppvObject = static_cast<IFileSystemAtomicWriter*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        return 1u;
    }

    HRESULT STDMETHODCALLTYPE SupportsAtomicWriterCommit(const wchar_t* path, FileSystemFlags flags, BOOL* supported) noexcept override
    {
        if (! path || path[0] == L'\0')
        {
            return E_INVALIDARG;
        }
        if (! supported)
        {
            return E_POINTER;
        }

        _observedFlags = flags;
        ++_probeCount;
        *supported = static_cast<uint32_t>(flags) == static_cast<uint32_t>(_expectedFlags) ? TRUE : FALSE;
        return S_OK;
    }

    [[nodiscard]] FileSystemFlags ObservedFlags() const noexcept
    {
        return _observedFlags;
    }

    [[nodiscard]] unsigned int ProbeCount() const noexcept
    {
        return _probeCount;
    }

private:
    FileSystemFlags _expectedFlags = FILESYSTEM_FLAG_NONE;
    FileSystemFlags _observedFlags = FILESYSTEM_FLAG_NONE;
    unsigned int _probeCount       = 0u;
};

[[nodiscard]] bool TryReadProviderPathIdentity(yyjson_val* root, ProviderCapabilitySnapshot& snapshot, std::wstring& reasonOut) noexcept;

[[nodiscard]] bool CapabilityListContainsWildcard(yyjson_val* value) noexcept
{
    if (! value || ! yyjson_is_arr(value))
    {
        return false;
    }

    const size_t count = yyjson_arr_size(value);
    for (size_t i = 0; i < count; ++i)
    {
        yyjson_val* item = yyjson_arr_get(value, i);
        if (! item || ! yyjson_is_str(item))
        {
            continue;
        }

        const char* text = yyjson_get_str(item);
        if (text && std::string_view(text) == "*")
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool TryGetJsonBool(yyjson_val* object, const char* key, bool& valueOut) noexcept
{
    valueOut          = false;
    yyjson_val* value = object ? yyjson_obj_get(object, key) : nullptr;
    if (! value || ! yyjson_is_bool(value))
    {
        return false;
    }

    valueOut = yyjson_get_bool(value) != 0;
    return true;
}

[[nodiscard]] bool TryGetJsonUInt(yyjson_val* object, const char* key, uint64_t& valueOut) noexcept
{
    valueOut          = 0;
    yyjson_val* value = object ? yyjson_obj_get(object, key) : nullptr;
    if (! value)
    {
        return false;
    }

    if (yyjson_is_uint(value))
    {
        valueOut = yyjson_get_uint(value);
        return true;
    }

    if (yyjson_is_int(value))
    {
        const int64_t signedValue = yyjson_get_int(value);
        if (signedValue >= 0)
        {
            valueOut = static_cast<uint64_t>(signedValue);
            return true;
        }
    }

    return false;
}

// Capability JSON remains a diagnostics contract. Keep its shape validation in
// selftest code so no executable host path can accidentally regain JSON authority.
[[nodiscard]] bool AreFileSystemCapabilitiesV2ValidForSelfTest(std::string_view jsonUtf8) noexcept
{
    if (jsonUtf8.empty())
    {
        return false;
    }

    std::string jsonCopy(jsonUtf8);
    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc(
        yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, nullptr), &yyjson_doc_free);
    if (! doc)
    {
        return false;
    }
    yyjson_val* root    = yyjson_doc_get_root(doc.get());
    yyjson_val* version = root ? yyjson_obj_get(root, "version") : nullptr;
    if (! root || ! yyjson_is_obj(root) || ! version || ! yyjson_is_int(version) || yyjson_get_int(version) != 2)
    {
        return false;
    }

    constexpr std::array<const char*, 11u> requiredObjects{{
        "operations",
        "concurrency",
        "transfer",
        "names",
        "identity",
        "publication",
        "links",
        "metadata",
        "verification",
        "cancellation",
        "directories",
    }};
    for (const char* name : requiredObjects)
    {
        yyjson_val* value = yyjson_obj_get(root, name);
        if (! value || ! yyjson_is_obj(value))
        {
            return false;
        }
    }
    for (const char* name : {"pathProfile", "rootId"})
    {
        yyjson_val* value = yyjson_obj_get(root, name);
        if (! value || ! yyjson_is_str(value) || yyjson_get_len(value) == 0u)
        {
            return false;
        }
    }
    if (! ParseDiagnosticFileSystemPathIdentityContractFromRoot(root, {}).has_value())
    {
        return false;
    }

    yyjson_val* operations = yyjson_obj_get(root, "operations");
    bool ignored           = false;
    for (const char* name : {"copy", "move", "nativeMove", "delete", "rename", "createDirectory", "properties", "read", "write", "recycle"})
    {
        if (! TryGetJsonBool(operations, name, ignored))
        {
            return false;
        }
    }
    const auto requireBooleans = [&](const char* objectName, std::initializer_list<const char*> memberNames) noexcept
    {
        yyjson_val* object = yyjson_obj_get(root, objectName);
        for (const char* member : memberNames)
        {
            if (! TryGetJsonBool(object, member, ignored))
            {
                return false;
            }
        }
        return true;
    };
    if (! requireBooleans("identity", {"boundDelete", "conditionalDelete"}) ||
        ! requireBooleans("publication", {"exclusiveStage", "conditionalPublish", "committedSize"}) ||
        ! requireBooleans("links", {"preserveFileLink", "preserveDirectoryLink", "retargetInTree", "exactLinkRemoval"}))
    {
        return false;
    }
    yyjson_val* verification = yyjson_obj_get(root, "verification");
    yyjson_val* cancellation = yyjson_obj_get(root, "cancellation");
    if (! TryGetJsonBool(verification, "hostReadback", ignored) || ! TryGetJsonBool(cancellation, "abort", ignored) ||
        ! TryGetJsonBool(cancellation, "deadline", ignored))
    {
        return false;
    }

    yyjson_val* routeClass = yyjson_obj_get(cancellation, "routeClass");
    yyjson_val* watchdog   = yyjson_obj_get(cancellation, "providerWatchdogTimeoutMs");
    uint64_t watchdogMs    = 0u;
    if (! routeClass || ! yyjson_is_str(routeClass) || ! TryGetJsonUInt(cancellation, "providerWatchdogTimeoutMs", watchdogMs) ||
        watchdog != yyjson_obj_get(cancellation, "providerWatchdogTimeoutMs"))
    {
        return false;
    }
    const std::string_view routeText(yyjson_get_str(routeClass), yyjson_get_len(routeClass));
    if ((routeText == "providerWatchdog") != (watchdogMs != 0u) || (routeText != "bounded" && routeText != "uncontained" && routeText != "providerWatchdog"))
    {
        return false;
    }

    yyjson_val* names     = yyjson_obj_get(root, "names");
    uint64_t maxComponent = 0u;
    if (! TryGetJsonUInt(names, "maxComponentUtf16", maxComponent) || maxComponent == 0u)
    {
        return false;
    }
    yyjson_val* proof = yyjson_obj_get(verification, "providerProof");
    if (! proof || ! yyjson_is_str(proof))
    {
        return false;
    }
    const std::string_view proofText(yyjson_get_str(proof), yyjson_get_len(proof));
    return proofText == "none" || proofText == "blake3-bound-object" || proofText == "writer-digest";
}

[[nodiscard]] std::wstring Utf16FromUtf8ForFileOpsSelfTest(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

struct PerfMetricSample
{
    uint64_t durationUs = 0;
    uint64_t value0     = 0;
    uint64_t value1     = 0;
    HRESULT hr          = S_OK;
};

[[nodiscard]] bool TryReadPerfMetricSamples(std::string_view metricName, std::vector<PerfMetricSample>& samplesOut) noexcept
{
    samplesOut.clear();
    if (metricName.empty())
    {
        return false;
    }

    try
    {
        const std::filesystem::path path = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
        std::ifstream input(path, std::ios::binary);
        if (! input)
        {
            return false;
        }

        std::string line;
        while (std::getline(input, line))
        {
            if (line.empty())
            {
                continue;
            }

            std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc(yyjson_read_opts(line.data(), line.size(), YYJSON_READ_ALLOW_BOM, nullptr, nullptr),
                                                                        &yyjson_doc_free);
            if (! doc)
            {
                continue;
            }

            yyjson_val* root = yyjson_doc_get_root(doc.get());
            if (! root || ! yyjson_is_obj(root))
            {
                continue;
            }

            yyjson_val* metric = yyjson_obj_get(root, "metric");
            if (! metric || ! yyjson_is_str(metric))
            {
                continue;
            }

            const char* metricText = yyjson_get_str(metric);
            if (! metricText || std::string_view(metricText, yyjson_get_len(metric)) != metricName)
            {
                continue;
            }

            PerfMetricSample sample{};
            uint64_t hrValue = 0;
            if (! TryGetJsonUInt(root, "durationUs", sample.durationUs) || ! TryGetJsonUInt(root, "value0", sample.value0) ||
                ! TryGetJsonUInt(root, "value1", sample.value1) || ! TryGetJsonUInt(root, "hr", hrValue))
            {
                continue;
            }
            sample.hr = static_cast<HRESULT>(static_cast<uint32_t>(hrValue));
            samplesOut.push_back(sample);
        }

        return true;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Selftest metric evidence is best-effort; callers fail explicitly when it is missing.
        return false;
    }
}

[[nodiscard]] bool TryReadProviderCapabilities(IFileSystem* fs, ProviderCapabilitySnapshot& snapshot, std::wstring& reasonOut) noexcept
{
    snapshot = {};
    reasonOut.clear();
    if (! fs)
    {
        reasonOut = L"file system is null";
        return false;
    }

    const char* jsonUtf8 = nullptr;
    const HRESULT hr     = fs->GetPathCapabilities(L"/", FILESYSTEM_COPY, &jsonUtf8);
    if (FAILED(hr) || ! jsonUtf8 || jsonUtf8[0] == '\0')
    {
        reasonOut = std::format(L"GetPathCapabilities failed hr=0x{:08X}", static_cast<unsigned long>(hr));
        return false;
    }

    std::string jsonCopy(jsonUtf8);
    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc(
        yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, nullptr), &yyjson_doc_free);
    if (! doc)
    {
        reasonOut = L"capabilities JSON did not parse";
        return false;
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root))
    {
        reasonOut = L"capabilities root is not an object";
        return false;
    }

    yyjson_val* version = yyjson_obj_get(root, "version");
    if (! version || ! yyjson_is_int(version) || yyjson_get_int(version) != 2)
    {
        reasonOut = L"capabilities version is not 2";
        return false;
    }

    yyjson_val* operations = yyjson_obj_get(root, "operations");
    if (! operations || ! yyjson_is_obj(operations))
    {
        reasonOut = L"capabilities operations object is missing";
        return false;
    }

    yyjson_val* concurrency = yyjson_obj_get(root, "concurrency");
    if (! concurrency || ! yyjson_is_obj(concurrency))
    {
        reasonOut = L"capabilities concurrency object is missing";
        return false;
    }

    if (! TryGetJsonBool(operations, "copy", snapshot.copyOperation) || ! TryGetJsonBool(operations, "move", snapshot.moveOperation) ||
        ! TryGetJsonBool(operations, "nativeMove", snapshot.nativeMoveOperation) || ! TryGetJsonBool(operations, "delete", snapshot.deleteOperation) ||
        ! TryGetJsonBool(operations, "rename", snapshot.renameOperation) ||
        ! TryGetJsonBool(operations, "createDirectory", snapshot.createDirectoryOperation) || ! TryGetJsonBool(operations, "properties", snapshot.properties) ||
        ! TryGetJsonBool(operations, "read", snapshot.read) || ! TryGetJsonBool(operations, "write", snapshot.write) ||
        ! TryGetJsonBool(operations, "recycle", snapshot.recycleOperation))
    {
        reasonOut = L"capabilities operations object is missing required boolean fields";
        return false;
    }

    if (! TryGetJsonUInt(concurrency, "copyMoveMax", snapshot.copyMoveMax) || ! TryGetJsonUInt(concurrency, "deleteMax", snapshot.deleteMax) ||
        ! TryGetJsonUInt(concurrency, "deleteRecycleBinMax", snapshot.deleteRecycleMax))
    {
        reasonOut = L"capabilities concurrency object is missing required numeric fields";
        return false;
    }

    yyjson_val* cross = yyjson_obj_get(root, "transfer");
    if (! cross || ! yyjson_is_obj(cross))
    {
        reasonOut = L"capabilities transfer object is missing";
        return false;
    }

    yyjson_val* exportNode = yyjson_obj_get(cross, "export");
    yyjson_val* importNode = yyjson_obj_get(cross, "import");
    if (! exportNode || ! yyjson_is_obj(exportNode) || ! importNode || ! yyjson_is_obj(importNode))
    {
        reasonOut = L"capabilities transfer import/export objects are missing";
        return false;
    }

    snapshot.exportCopyWildcard = CapabilityListContainsWildcard(yyjson_obj_get(exportNode, "copy"));
    snapshot.exportMoveWildcard = CapabilityListContainsWildcard(yyjson_obj_get(exportNode, "move"));
    snapshot.importCopyWildcard = CapabilityListContainsWildcard(yyjson_obj_get(importNode, "copy"));
    snapshot.importMoveWildcard = CapabilityListContainsWildcard(yyjson_obj_get(importNode, "move"));

    yyjson_val* links = yyjson_obj_get(root, "links");
    if (! links || ! yyjson_is_obj(links) || ! TryGetJsonBool(links, "preserveFileLink", snapshot.preserveFileLink) ||
        ! TryGetJsonBool(links, "preserveDirectoryLink", snapshot.preserveDirectoryLink) ||
        ! TryGetJsonBool(links, "retargetInTree", snapshot.retargetInTree) || ! TryGetJsonBool(links, "exactLinkRemoval", snapshot.exactLinkRemoval))
    {
        reasonOut = L"capabilities links object is missing required boolean fields";
        return false;
    }

    if (! TryReadProviderPathIdentity(root, snapshot, reasonOut))
    {
        return false;
    }

    return true;
}

[[nodiscard]] bool TryReadTypedProviderCapabilities(
    IFileSystem* fs, std::wstring_view providerId, ProviderCapabilitySnapshot& snapshot, std::wstring& reasonOut, bool* usedArenaFallback = nullptr) noexcept
{
    snapshot = {};
    reasonOut.clear();
    if (usedArenaFallback != nullptr)
    {
        *usedArenaFallback = false;
    }
    const FileSystemRouteContract::QueryResult route = FileSystemRouteContract::Query(fs, L"/", FILESYSTEM_COPY, providerId);
    if (route.state != FileSystemRouteContract::QueryState::Available || ! route.snapshot.pathIdentity.has_value())
    {
        reasonOut =
            std::format(L"typed route query failed state={} hr=0x{:08X}", static_cast<unsigned int>(route.state), static_cast<unsigned long>(route.status));
        return false;
    }
    if (usedArenaFallback != nullptr)
    {
        *usedArenaFallback = route.usedArenaFallback;
    }

    const FileSystemRouteContract::Snapshot& facts = route.snapshot;
    snapshot.copyOperation                         = facts.copyOperation;
    snapshot.moveOperation                         = facts.moveOperation;
    snapshot.nativeMoveOperation                   = facts.nativeMoveOperation;
    snapshot.deleteOperation                       = facts.deleteOperation;
    snapshot.renameOperation                       = facts.renameOperation;
    snapshot.recycleOperation                      = facts.recycleOperation;
    snapshot.createDirectoryOperation              = facts.createDirectoryOperation;
    snapshot.properties                            = facts.properties;
    snapshot.read                                  = facts.read;
    snapshot.write                                 = facts.write;
    snapshot.copyMoveMax                           = facts.copyMoveMaxConcurrency;
    snapshot.deleteMax                             = facts.deleteMaxConcurrency;
    snapshot.deleteRecycleMax                      = facts.deleteRecycleBinMaxConcurrency;
    snapshot.preserveFileLink                      = facts.preserveFileLink;
    snapshot.preserveDirectoryLink                 = facts.preserveDirectoryLink;
    snapshot.retargetInTree                        = facts.retargetInTree;
    snapshot.exactLinkRemoval                      = facts.exactLinkRemoval;
    snapshot.pathIdentityPresent                   = true;
    snapshot.pathTextStableIdentity                = facts.pathIdentity->pathTextStableIdentity;
    snapshot.componentComparison =
        facts.pathIdentity->componentComparison == FileSystemPathComponentComparison::OrdinalIgnoreCase ? L"ordinalIgnoreCase" : L"ordinalCaseSensitive";
    snapshot.preferredSeparator = facts.pathIdentity->preferredSeparator;
    snapshot.acceptedSeparators = facts.pathIdentity->acceptedSeparators;
    snapshot.casePreserving     = facts.pathIdentity->casePreserving;
    switch (facts.pathIdentity->caseOnlyRename)
    {
        case FileSystemPathCaseOnlyRename::Supported: snapshot.caseOnlyRename = L"supported"; break;
        case FileSystemPathCaseOnlyRename::NoOp: snapshot.caseOnlyRename = L"noOp"; break;
        case FileSystemPathCaseOnlyRename::Unsupported: snapshot.caseOnlyRename = L"unsupported"; break;
        case FileSystemPathCaseOnlyRename::NotApplicable: snapshot.caseOnlyRename = L"notApplicable"; break;
    }

    const auto peerAllowed = [&](FileSystemOperation operation, FileSystemTransferPeerRole role) noexcept
    {
        const FileSystemRouteContract::BooleanResult peer = FileSystemRouteContract::QueryTransferPeerAllowed(fs, L"/", operation, role, L"selftest/peer");
        return peer.state == FileSystemRouteContract::QueryState::Available && peer.value;
    };
    snapshot.exportCopyWildcard = peerAllowed(FILESYSTEM_COPY, FILESYSTEM_TRANSFER_PEER_EXPORT);
    snapshot.exportMoveWildcard = peerAllowed(FILESYSTEM_MOVE, FILESYSTEM_TRANSFER_PEER_EXPORT);
    snapshot.importCopyWildcard = peerAllowed(FILESYSTEM_COPY, FILESYSTEM_TRANSFER_PEER_IMPORT);
    snapshot.importMoveWildcard = peerAllowed(FILESYSTEM_MOVE, FILESYSTEM_TRANSFER_PEER_IMPORT);
    return true;
}

[[nodiscard]] bool TryReadProviderPathIdentity(yyjson_val* root, ProviderCapabilitySnapshot& snapshot, std::wstring& reasonOut) noexcept
{
    yyjson_val* identity = yyjson_obj_get(root, "names");
    if (! identity || ! yyjson_is_obj(identity))
    {
        reasonOut = L"capabilities names object is missing";
        return false;
    }

    if (! TryGetJsonBool(identity, "pathTextStableIdentity", snapshot.pathTextStableIdentity))
    {
        reasonOut = L"capabilities names pathTextStableIdentity is missing or not boolean";
        return false;
    }

    const auto readJsonString = [&](const char* key, std::wstring& valueOut) noexcept -> bool
    {
        yyjson_val* value = yyjson_obj_get(identity, key);
        if (! value || ! yyjson_is_str(value))
        {
            reasonOut = std::format(L"capabilities names {} is missing or not string", Utf16FromUtf8ForFileOpsSelfTest(key));
            return false;
        }

        const char* text = yyjson_get_str(value);
        const size_t len = yyjson_get_len(value);
        valueOut         = text ? Utf16FromUtf8ForFileOpsSelfTest(std::string_view{text, len}) : std::wstring{};
        if (valueOut.empty())
        {
            reasonOut = std::format(L"capabilities names {} is empty or invalid UTF-8", Utf16FromUtf8ForFileOpsSelfTest(key));
            return false;
        }
        return true;
    };

    if (! readJsonString("comparison", snapshot.componentComparison))
    {
        return false;
    }
    if (snapshot.componentComparison != L"ordinalIgnoreCase" && snapshot.componentComparison != L"ordinalCaseSensitive")
    {
        reasonOut = L"capabilities names comparison is not concrete";
        return false;
    }

    std::wstring preferredSeparator;
    if (! readJsonString("preferredSeparator", preferredSeparator) || preferredSeparator.size() != 1u)
    {
        reasonOut = L"capabilities names preferredSeparator must be exactly one character";
        return false;
    }
    snapshot.preferredSeparator = preferredSeparator.front();

    yyjson_val* accepted = yyjson_obj_get(identity, "acceptedSeparators");
    if (! accepted || ! yyjson_is_arr(accepted) || yyjson_arr_size(accepted) == 0u)
    {
        reasonOut = L"capabilities names acceptedSeparators is missing or empty";
        return false;
    }

    snapshot.acceptedSeparators.clear();
    const size_t acceptedCount = yyjson_arr_size(accepted);
    for (size_t index = 0; index < acceptedCount; ++index)
    {
        yyjson_val* item = yyjson_arr_get(accepted, index);
        if (! item || ! yyjson_is_str(item))
        {
            reasonOut = L"capabilities names acceptedSeparators contains a non-string value";
            return false;
        }

        const char* text       = yyjson_get_str(item);
        std::wstring separator = text ? Utf16FromUtf8ForFileOpsSelfTest(std::string_view{text, yyjson_get_len(item)}) : std::wstring{};
        if (separator.size() != 1u)
        {
            reasonOut = L"capabilities names acceptedSeparators entries must be exactly one character";
            return false;
        }
        if (snapshot.acceptedSeparators.find(separator.front()) == std::wstring::npos)
        {
            snapshot.acceptedSeparators.push_back(separator.front());
        }
    }

    if (snapshot.acceptedSeparators.find(snapshot.preferredSeparator) == std::wstring::npos)
    {
        reasonOut = L"capabilities names acceptedSeparators must include preferredSeparator";
        return false;
    }

    if (! TryGetJsonBool(identity, "casePreserving", snapshot.casePreserving))
    {
        reasonOut = L"capabilities names casePreserving is missing or not boolean";
        return false;
    }

    if (! readJsonString("caseOnlyRename", snapshot.caseOnlyRename))
    {
        return false;
    }
    if (snapshot.caseOnlyRename != L"supported" && snapshot.caseOnlyRename != L"noOp" && snapshot.caseOnlyRename != L"unsupported" &&
        snapshot.caseOnlyRename != L"notApplicable")
    {
        reasonOut = L"capabilities names caseOnlyRename is not a supported value";
        return false;
    }

    snapshot.pathIdentityPresent = true;
    return true;
}

[[nodiscard]] bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (std::towlower(a[i]) != std::towlower(b[i]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool StartsWithIgnoreCase(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (prefix.size() > text.size())
    {
        return false;
    }

    for (size_t i = 0; i < prefix.size(); ++i)
    {
        if (std::towlower(text[i]) != std::towlower(prefix[i]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] std::filesystem::path MakePluginDisplayPath(std::wstring_view pluginShortId, std::wstring_view pluginPath) noexcept
{
    std::wstring pathText(pluginPath);
    if (pathText.empty())
    {
        pathText = L"/";
    }
    else if (pathText.front() != L'/')
    {
        pathText.insert(pathText.begin(), L'/');
    }

    std::wstring displayPath;
    displayPath.reserve(pluginShortId.size() + 1u + pathText.size());
    displayPath.append(pluginShortId);
    displayPath.push_back(L':');
    displayPath.append(pathText);
    return std::filesystem::path(displayPath);
}

[[maybe_unused]] [[nodiscard]] std::filesystem::path MakeMountedDisplayPath(std::wstring_view pluginShortId,
                                                                            const std::filesystem::path& instanceContext,
                                                                            std::wstring_view pluginPath) noexcept
{
    std::wstring pathText(pluginPath);
    if (pathText.empty())
    {
        pathText = L"/";
    }
    else if (pathText.front() != L'/')
    {
        pathText.insert(pathText.begin(), L'/');
    }

    std::wstring displayPath;
    const std::wstring contextText = instanceContext.wstring();
    displayPath.reserve(pluginShortId.size() + 1u + contextText.size() + 1u + pathText.size());
    displayPath.append(pluginShortId);
    displayPath.push_back(L':');
    displayPath.append(contextText);
    displayPath.push_back(L'|');
    displayPath.append(pathText);
    return std::filesystem::path(displayPath);
}

[[nodiscard]] bool CurrentPanePathEquals(const FolderWindow* folderWindow, FolderWindow::Pane pane, const std::filesystem::path& expected) noexcept
{
    if (! folderWindow)
    {
        return false;
    }

    const std::optional<std::filesystem::path> current = folderWindow->GetCurrentPath(pane);
    return current.has_value() && EqualsIgnoreCase(current->wstring(), expected.wstring());
}

[[nodiscard]] bool DummyWriteTextFile(IFileSystem* fs, std::wstring_view path, std::string_view contents, bool allowOverwrite = false) noexcept
{
    if (! fs || path.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    const HRESULT hrIo = fs->QueryInterface(IID_PPV_ARGS(io.addressof()));
    if (FAILED(hrIo) || ! io)
    {
        return false;
    }

    const std::wstring pathText(path);
    wil::com_ptr<IFileWriter> writer;
    const FileSystemFlags flags = allowOverwrite ? FILESYSTEM_FLAG_ALLOW_OVERWRITE : FILESYSTEM_FLAG_NONE;
    const HRESULT hrWriter      = io->CreateFileWriter(pathText.c_str(), flags, writer.put());
    if (FAILED(hrWriter) || ! writer)
    {
        return false;
    }

    if (! contents.empty())
    {
        unsigned long written = 0;
        const HRESULT hrWrite = writer->Write(contents.data(), static_cast<unsigned long>(contents.size()), &written);
        if (FAILED(hrWrite) || written != static_cast<unsigned long>(contents.size()))
        {
            return false;
        }
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] std::wstring ToPluginPathText(const std::filesystem::path& path) noexcept
{
    return path.generic_wstring();
}

[[nodiscard]] bool WriteFileBytesFsIo(
    const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, const void* data, size_t sizeBytes, bool allowOverwrite = false) noexcept
{
    if (! io)
    {
        return false;
    }

    if (sizeBytes > 0 && ! data)
    {
        return false;
    }

    if (sizeBytes > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return false;
    }

    wil::com_ptr<IFileWriter> writer;
    const std::wstring pathText = ToPluginPathText(path);
    const FileSystemFlags flags = allowOverwrite ? FILESYSTEM_FLAG_ALLOW_OVERWRITE : FILESYSTEM_FLAG_NONE;
    const HRESULT createHr      = io->CreateFileWriter(pathText.c_str(), flags, writer.put());
    if (FAILED(createHr) || ! writer)
    {
        return false;
    }

    if (sizeBytes > 0)
    {
        unsigned long written = 0;
        const HRESULT writeHr = writer->Write(data, static_cast<unsigned long>(sizeBytes), &written);
        if (FAILED(writeHr) || written != static_cast<unsigned long>(sizeBytes))
        {
            return false;
        }
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] bool WriteFileTextFsIo(const wil::com_ptr<IFileSystemIO>& io,
                                     const std::filesystem::path& path,
                                     std::string_view text,
                                     bool allowOverwrite = false) noexcept
{
    return WriteFileBytesFsIo(io, path, text.data(), text.size(), allowOverwrite);
}

[[nodiscard]] bool WritePatternFileFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, uint64_t sizeBytes) noexcept
{
    if (! io)
    {
        return false;
    }

    wil::com_ptr<IFileWriter> writer;
    const std::wstring pathText = ToPluginPathText(path);
    const HRESULT createHr      = io->CreateFileWriter(pathText.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (FAILED(createHr) || ! writer)
    {
        return false;
    }

    constexpr size_t kChunkBytes = 32 * 1024;
    std::array<unsigned char, kChunkBytes> buffer{};
    for (size_t i = 0; i < kChunkBytes; ++i)
    {
        buffer[i] = static_cast<unsigned char>((i * 131u) ^ 0x5Au);
    }

    uint64_t remaining = sizeBytes;
    while (remaining > 0)
    {
        const unsigned long chunk = static_cast<unsigned long>((std::min<uint64_t>)(remaining, static_cast<uint64_t>(kChunkBytes)));
        unsigned long written     = 0;
        const HRESULT writeHr     = writer->Write(buffer.data(), chunk, &written);
        if (FAILED(writeHr) || written != chunk)
        {
            return false;
        }

        remaining -= chunk;
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] bool ReadFileTextFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, std::string& textOut) noexcept
{
    textOut.clear();
    if (! io)
    {
        return false;
    }

    wil::com_ptr<IFileReader> reader;
    const std::wstring pathText = ToPluginPathText(path);
    const HRESULT readerHr      = io->CreateFileReader(pathText.c_str(), reader.put());
    if (FAILED(readerHr) || ! reader)
    {
        return false;
    }

    uint64_t sizeBytes   = 0;
    const HRESULT sizeHr = reader->GetSize(&sizeBytes);
    if (FAILED(sizeHr) || sizeBytes > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
    {
        return false;
    }

    textOut.resize(static_cast<size_t>(sizeBytes));
    size_t totalRead = 0;
    while (totalRead < textOut.size())
    {
        unsigned long chunkRead = 0;
        const unsigned long requestBytes =
            static_cast<unsigned long>((std::min)(textOut.size() - totalRead, static_cast<size_t>((std::numeric_limits<unsigned long>::max)())));
        const HRESULT readHr = reader->Read(textOut.data() + totalRead, requestBytes, &chunkRead);
        if (FAILED(readHr))
        {
            textOut.clear();
            return false;
        }

        if (chunkRead == 0)
        {
            textOut.clear();
            return false;
        }

        totalRead += chunkRead;
    }

    return true;
}

[[nodiscard]] bool GetFileSizeFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, uint64_t& sizeBytesOut) noexcept
{
    sizeBytesOut = 0;
    if (! io)
    {
        return false;
    }

    wil::com_ptr<IFileReader> reader;
    const std::wstring pathText = ToPluginPathText(path);
    const HRESULT readerHr      = io->CreateFileReader(pathText.c_str(), reader.put());
    if (FAILED(readerHr) || ! reader)
    {
        return false;
    }

    return SUCCEEDED(reader->GetSize(&sizeBytesOut));
}

[[nodiscard]] bool EnsureDirectoryExistsFsOps(const wil::com_ptr<IFileSystemDirectoryOperations>& ops, const std::filesystem::path& path) noexcept
{
    if (! ops)
    {
        return false;
    }

    const std::filesystem::path normalized = path.lexically_normal();
    std::filesystem::path current          = normalized.root_path();
    for (const auto& part : normalized.relative_path())
    {
        current /= part;
        const HRESULT hr = ops->CreateDirectory(current.c_str());
        if (SUCCEEDED(hr) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            continue;
        }

        return false;
    }

    return true;
}

[[nodiscard]] std::wstring JoinPluginPathForSelfTest(std::wstring_view base, std::wstring_view leaf) noexcept
{
    std::wstring result = NormalizePluginPathForSelfTest(base);
    if (result.empty())
    {
        return {};
    }

    if (! leaf.empty())
    {
        if (result.back() != L'/')
        {
            result.push_back(L'/');
        }

        result.append(leaf);
    }

    return NormalizePluginPathForSelfTest(result);
}

[[nodiscard]] std::wstring MakeConnectionPathForSelfTest(std::wstring_view profileName, std::wstring_view pluginPath) noexcept
{
    if (profileName.empty() || pluginPath.empty() || pluginPath.front() != L'/')
    {
        return {};
    }

    return std::format(L"/@conn:{}{}", profileName, pluginPath);
}

[[nodiscard]] bool PathExistsFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, unsigned long* attributes = nullptr) noexcept
{
    if (! io)
    {
        return false;
    }

    unsigned long attrs           = 0;
    const std::wstring pluginPath = ToPluginPathText(path);
    const HRESULT hr              = io->GetAttributes(pluginPath.c_str(), &attrs);
    if (FAILED(hr))
    {
        return false;
    }

    if (attributes)
    {
        *attributes = attrs;
    }

    return true;
}

void ResetRemoteS3State(SelfTestState& state) noexcept
{
    state.fsRemoteS3.reset();
    state.remoteS3ProfileName.clear();
    state.remoteS3CaseRootConn.clear();
    state.remoteS3AltCaseRootConn.clear();
    state.remoteS3UploadDirConn.clear();
    state.remoteS3MoveDirConn.clear();
    state.remoteS3SeedFileConn.clear();
    state.remoteS3UploadedFileConn.clear();
    state.remoteS3MovedFileConn.clear();
    state.remoteS3RenamedFileConn.clear();
    state.remoteS3UploadedDirConn.clear();
    state.remoteS3DisplayPath.clear();
    state.remoteS3Workspace.clear();
    state.remoteS3UploadSource.clear();
    state.remoteS3DownloadDir.clear();
    state.remoteS3DownloadedFile.clear();
    state.remoteS3DirUploadSource.clear();
    state.remoteS3DirMoveDownloadDir.clear();
    state.remoteS3DirMovedFile.clear();
    state.remoteS3Payload.clear();
}

void CleanupRemoteS3Case(SelfTestState& state) noexcept
{
    if (state.fsRemoteS3 && ! state.remoteS3CaseRootConn.empty())
    {
        static_cast<void>(state.fsRemoteS3->DeleteItem(state.remoteS3CaseRootConn.c_str(),
                                                       static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                       nullptr,
                                                       nullptr,
                                                       nullptr));
    }

    if (state.fsRemoteS3 && ! state.remoteS3AltCaseRootConn.empty())
    {
        static_cast<void>(state.fsRemoteS3->DeleteItem(state.remoteS3AltCaseRootConn.c_str(),
                                                       static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                       nullptr,
                                                       nullptr,
                                                       nullptr));
    }

    ResetRemoteS3State(state);
}

void ResetRemoteOneDrivePersonalState(SelfTestState& state) noexcept
{
    state.fsRemoteOneDrivePersonal.reset();
    state.remoteOneDrivePersonalProfileName.clear();
    state.remoteOneDrivePersonalCaseRootConn.clear();
    state.remoteOneDrivePersonalUploadDirConn.clear();
    state.remoteOneDrivePersonalMoveDirConn.clear();
    state.remoteOneDrivePersonalSeedFileConn.clear();
    state.remoteOneDrivePersonalUploadedFileConn.clear();
    state.remoteOneDrivePersonalMovedFileConn.clear();
    state.remoteOneDrivePersonalDisplayPath.clear();
    state.remoteOneDrivePersonalWorkspace.clear();
    state.remoteOneDrivePersonalUploadSource.clear();
    state.remoteOneDrivePersonalDownloadDir.clear();
    state.remoteOneDrivePersonalDownloadedFile.clear();
    state.remoteOneDrivePersonalPayload.clear();
}

void CleanupRemoteOneDrivePersonalCase(SelfTestState& state) noexcept
{
    if (state.fsRemoteOneDrivePersonal && ! state.remoteOneDrivePersonalCaseRootConn.empty())
    {
        static_cast<void>(
            state.fsRemoteOneDrivePersonal->DeleteItem(state.remoteOneDrivePersonalCaseRootConn.c_str(),
                                                       static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                       nullptr,
                                                       nullptr,
                                                       nullptr));
    }

    ResetRemoteS3State(state);
    ResetRemoteOneDrivePersonalState(state);
}

[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindLoadedPluginEntry(std::wstring_view pluginId) noexcept
{
    if (pluginId.empty())
    {
        return nullptr;
    }

    FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
    for (const auto& entry : pluginManager.GetPlugins())
    {
        if (! entry.id.empty() && EqualsIgnoreCase(entry.id, pluginId))
        {
            return &entry;
        }
    }

    return nullptr;
}

[[nodiscard]] std::wstring MakeUniqueConnectionProfileName(std::wstring_view baseName) noexcept
{
    std::wstring base(baseName);
    if (base.empty())
    {
        base = L"FileOpsSelfTestDummy";
    }

    const auto isTaken = [&](std::wstring_view candidate) noexcept -> bool
    {
        if (! g_settings.connections)
        {
            return false;
        }

        for (const Common::Settings::ConnectionProfile& profile : g_settings.connections->items)
        {
            if (! profile.name.empty() && EqualsIgnoreCase(profile.name, candidate))
            {
                return true;
            }
        }

        return false;
    };

    if (! isTaken(base))
    {
        return base;
    }

    for (unsigned int i = 2; i < 1000; ++i)
    {
        const std::wstring candidate = std::format(L"{} ({})", base, i);
        if (! isTaken(candidate))
        {
            return candidate;
        }
    }

    // Best-effort fallback (unlikely).
    return std::format(L"{} ({})", base, GetTickCount64());
}

[[nodiscard]] std::wstring NewGuidString() noexcept
{
    std::wstring id;
    return SUCCEEDED(Common::Settings::CreateConnectionProfileId(id)) ? id : std::wstring{};
}

[[nodiscard]] Common::Settings::JsonValue MakeJsonObjectWithUIntMembers(std::initializer_list<std::pair<std::string_view, uint64_t>> members) noexcept
{
    auto obj = std::make_shared<Common::Settings::JsonObject>();
    obj->members.reserve(members.size());

    for (const auto& [key, value] : members)
    {
        Common::Settings::JsonValue member{};
        member.value = static_cast<uint64_t>(value);
        obj->members.emplace_back(std::string(key), std::move(member));
    }

    Common::Settings::JsonValue out{};
    out.value = obj;
    return out;
}

void RemoveConnectionProfileByName(std::wstring_view name) noexcept
{
    if (! g_settings.connections || name.empty())
    {
        return;
    }

    auto& items = g_settings.connections->items;
    items.erase(std::remove_if(items.begin(),
                               items.end(),
                               [&](const Common::Settings::ConnectionProfile& profile) noexcept
    { return ! profile.name.empty() && EqualsIgnoreCase(profile.name, name); }),
                items.end());
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0;
    while (start < text.size())
    {
        const wchar_t ch = text[start];
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        ++start;
    }

    size_t end = text.size();
    while (end > start)
    {
        const wchar_t ch = text[end - 1u];
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] std::wstring GetEnvVarTrimmed(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return {};
    }

    std::wstring key(name);
    DWORD required = GetEnvironmentVariableW(key.c_str(), nullptr, 0);
    if (required == 0)
    {
        return {};
    }

    std::wstring value;
    value.resize(required);
    const DWORD written = GetEnvironmentVariableW(key.c_str(), value.data(), required);
    if (written == 0 || written >= required)
    {
        return {};
    }

    value.resize(written);
    return TrimWhitespace(value);
}

void SecureClearAndFreeSecret(wil::unique_cotaskmem_string& secret) noexcept
{
    if (wchar_t* text = secret.get())
    {
        const size_t len = wcslen(text);
        if (len > 0)
        {
            SecureZeroMemory(text, len * sizeof(wchar_t));
        }
    }
    secret.reset();
}

[[nodiscard]] bool IsHostConnectionUiUnavailableHr(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
}

struct PhaseCheckResult
{
    SelfTest::SelfTestCaseResult::Status status = SelfTest::SelfTestCaseResult::Status::skipped;
    std::wstring reason;
};

struct ResolvedRemoteProfile
{
    const Common::Settings::ConnectionProfile* profile = nullptr;
    std::wstring profileName;
    bool usedFallback = false;
};

[[nodiscard]] ResolvedRemoteProfile ResolveRemoteConnectionProfile(std::wstring_view envVarName,
                                                                   std::wstring_view defaultProfileName,
                                                                   std::wstring_view expectedPluginId) noexcept
{
    ResolvedRemoteProfile resolved{};
    const std::wstring overrideName = GetEnvVarTrimmed(envVarName);
    resolved.profileName            = ! overrideName.empty() ? overrideName : std::wstring(defaultProfileName);
    resolved.profile                = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, resolved.profileName);
    if (resolved.profile || ! overrideName.empty() || ! g_settings.connections)
    {
        return resolved;
    }

    for (const Common::Settings::ConnectionProfile& profile : g_settings.connections->items)
    {
        if (profile.name.empty() || profile.pluginId.empty() || ! EqualsIgnoreCase(profile.pluginId, expectedPluginId))
        {
            continue;
        }

        bool selfTestNamed = false;
        for (size_t i = 0; i + 8u <= profile.name.size(); ++i)
        {
            if (std::towlower(profile.name[i + 0u]) == L's' && std::towlower(profile.name[i + 1u]) == L'e' && std::towlower(profile.name[i + 2u]) == L'l' &&
                std::towlower(profile.name[i + 3u]) == L'f' && std::towlower(profile.name[i + 4u]) == L't' && std::towlower(profile.name[i + 5u]) == L'e' &&
                std::towlower(profile.name[i + 6u]) == L's' && std::towlower(profile.name[i + 7u]) == L't')
            {
                selfTestNamed = true;
                break;
            }
        }

        if (! selfTestNamed)
        {
            continue;
        }

        resolved.profile      = &profile;
        resolved.profileName  = profile.name;
        resolved.usedFallback = true;
        break;
    }

    return resolved;
}

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSecret(std::wstring_view protocolLabel,
                                                           std::wstring_view envVarName,
                                                           std::wstring_view defaultProfileName,
                                                           std::wstring_view expectedPluginId) noexcept
{
    const ResolvedRemoteProfile resolved               = ResolveRemoteConnectionProfile(envVarName, defaultProfileName, expectedPluginId);
    const std::wstring& profileName                    = resolved.profileName;
    const Common::Settings::ConnectionProfile* profile = resolved.profile;
    if (! profile)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: connection profile not found (set {} or create '{}').", protocolLabel, envVarName, defaultProfileName)};
    }

    if (profile->pluginId.empty() || ! EqualsIgnoreCase(profile->pluginId, expectedPluginId))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile targets a different plugin.", protocolLabel)};
    }

    if (profile->authMode == Common::Settings::ConnectionAuthMode::Anonymous)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: authMode=anonymous (no secret needed).", protocolLabel)};
    }

    if (profile->id.empty())
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile is missing a stable id.", protocolLabel)};
    }

    const bool bypassHello = g_settings.connections ? g_settings.connections->bypassWindowsHello : false;
    if (profile->requireWindowsHello && ! bypassHello)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: requireWindowsHello=true (enable bypassWindowsHello or disable the profile flag for automation).", protocolLabel)};
    }

    if (! profile->savePassword)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: savePassword=false (secret is not persisted).", protocolLabel)};
    }

    HostConnectionSecretKind kind = HOST_CONNECTION_SECRET_PASSWORD;
    if (profile->authMode == Common::Settings::ConnectionAuthMode::SshKey)
    {
        kind = HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE;
    }
    else if (profile->authMode == Common::Settings::ConnectionAuthMode::OAuth2Pkce)
    {
        kind = HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN;
    }

    wil::com_ptr<IHostConnections> hostConnections;
    const HRESULT hrQI = GetHostServices()->QueryInterface(IID_PPV_ARGS(hostConnections.addressof()));
    if (FAILED(hrQI) || ! hostConnections)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: missing IHostConnections. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(hrQI))};
    }

    wil::unique_cotaskmem_string secret;
    const HRESULT hrSecret = hostConnections->GetConnectionSecret(profileName.c_str(), kind, nullptr, secret.put());
    if (hrSecret == HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
    {
        SecureClearAndFreeSecret(secret);
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: secret not found.", protocolLabel)};
    }

    if (FAILED(hrSecret))
    {
        SecureClearAndFreeSecret(secret);
        if (IsHostConnectionUiUnavailableHr(hrSecret))
        {
            return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                    .reason = std::format(
                        L"{}: host connection UI unavailable for GetConnectionSecret. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(hrSecret))};
        }

        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: GetConnectionSecret failed. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(hrSecret))};
    }

    if (! secret)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: GetConnectionSecret returned success but no secret.", protocolLabel)};
    }

    SecureClearAndFreeSecret(secret);
    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

[[nodiscard]] bool ContainsIgnoreCase(std::wstring_view text, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }

    if (text.size() < needle.size())
    {
        return false;
    }

    for (size_t i = 0; i + needle.size() <= text.size(); ++i)
    {
        bool match = true;
        for (size_t j = 0; j < needle.size(); ++j)
        {
            if (std::towlower(text[i + j]) != std::towlower(needle[j]))
            {
                match = false;
                break;
            }
        }

        if (match)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring NormalizePluginPathForSelfTest(std::wstring_view rawPath) noexcept
{
    std::wstring path = TrimWhitespace(rawPath);
    for (wchar_t& ch : path)
    {
        if (ch == L'\\')
        {
            ch = L'/';
        }
    }

    while (path.size() > 1u && path.back() == L'/')
    {
        path.pop_back();
    }

    return path;
}

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSandbox(std::wstring_view protocolLabel,
                                                            std::wstring_view envVarName,
                                                            std::wstring_view defaultProfileName,
                                                            std::wstring_view expectedPluginId) noexcept
{
    const ResolvedRemoteProfile resolved               = ResolveRemoteConnectionProfile(envVarName, defaultProfileName, expectedPluginId);
    const Common::Settings::ConnectionProfile* profile = resolved.profile;
    if (! profile)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: connection profile not found (set {} or create '{}').", protocolLabel, envVarName, defaultProfileName)};
    }

    if (profile->pluginId.empty() || ! EqualsIgnoreCase(profile->pluginId, expectedPluginId))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile targets a different plugin.", protocolLabel)};
    }

    const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
    if (initialPath.empty() || initialPath[0] != L'/')
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must be an absolute plugin path (starting with '/').", protocolLabel)};
    }

    if (initialPath == L"/")
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must point to a dedicated selftest folder/prefix (not '/').", protocolLabel)};
    }

    if (ContainsIgnoreCase(initialPath, L"/@conn:"))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must not include the host-reserved '/@conn:' prefix.", protocolLabel)};
    }

    size_t segmentCount = 0;
    std::wstring firstSegment;
    for (size_t i = 0; i < initialPath.size();)
    {
        while (i < initialPath.size() && initialPath[i] == L'/')
        {
            ++i;
        }

        const size_t start = i;
        while (i < initialPath.size() && initialPath[i] != L'/')
        {
            ++i;
        }

        if (i == start)
        {
            continue;
        }

        const std::wstring_view segment = std::wstring_view(initialPath).substr(start, i - start);
        if (segment == L"." || segment == L"..")
        {
            return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                    .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must not contain '.' or '..' segments.", protocolLabel)};
        }

        if (segmentCount == 0)
        {
            firstSegment.assign(segment);
        }
        ++segmentCount;
    }

    const bool isS3                  = expectedPluginId == kPluginIdS3;
    const bool dedicatedS3BucketRoot = isS3 && segmentCount == 1 && ContainsIgnoreCase(firstSegment, L"selftest");
    if (isS3 && segmentCount < 2 && ! dedicatedS3BucketRoot)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must include a bucket and a dedicated selftest prefix (e.g. "
                                      L"'/bucket/red-salamander-selftest'), or use a dedicated selftest bucket root (e.g. '/redsalamander-selftest').",
                                      protocolLabel)};
    }

    if (! ContainsIgnoreCase(initialPath, L"selftest"))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason =
                    std::format(L"{}: HARD REQUIREMENT: initialPath must include 'selftest' (case-insensitive) to prove it is test-only.", protocolLabel)};
    }

    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

std::filesystem::path GetTempRootPath() noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::FileOperations);
    if (suiteRoot.empty())
    {
        return {};
    }
    return suiteRoot / L"work";
}

[[nodiscard]] std::wstring MakeExtendedPathForSelfTest(const std::filesystem::path& inputPath);

// The administrative-share alias of a sandbox path (\\localhost\<drive>$\...) names the same volume
// through the local-win32-smb profile: the one second qualified endpoint a single-volume machine
// has, so cross-endpoint routes can be exercised without a second disk. Empty off a drive letter.
[[nodiscard]] std::filesystem::path LoopbackShareAlias(const std::filesystem::path& localPath) noexcept
{
    // The share exposes the whole volume; a test may name only a path inside its marked sandbox through it.
    std::error_code sandboxEc;
    const std::wstring driveRootName = localPath.root_name().native();
    if (driveRootName.size() != 2u || driveRootName[1] != L':' || ! Common::Testing::RequireSandboxedExposureTarget(localPath, sandboxEc))
    {
        return {};
    }
    return std::filesystem::path(std::format(LR"(\\localhost\{}$\)", driveRootName[0])) / localPath.lexically_relative(localPath.root_path());
}

// A SUBST drive (no broadcast) naming one directory inside the marked sandbox; any other directory is refused.
// The removal deletes only that exact definition.
[[nodiscard]] bool DefineSandboxedSubstDrive(const std::wstring& deviceName, const std::filesystem::path& directory) noexcept
{
    std::error_code sandboxEc;
    if (! Common::Testing::RequireSandboxedExposureTarget(directory, sandboxEc))
    {
        ::SetLastError(ERROR_BAD_PATHNAME);
        return false;
    }
    const std::wstring rawTarget = std::wstring(L"\\??\\") + directory.native();
    return DefineDosDeviceW(DDD_RAW_TARGET_PATH | DDD_NO_BROADCAST_SYSTEM, deviceName.c_str(), rawTarget.c_str()) != FALSE;
}

void RemoveSandboxedSubstDrive(const std::wstring& deviceName, const std::filesystem::path& directory) noexcept
{
    const std::wstring rawTarget = std::wstring(L"\\??\\") + directory.native();
    constexpr DWORD removeFlags  = DDD_REMOVE_DEFINITION | DDD_EXACT_MATCH_ON_REMOVE | DDD_RAW_TARGET_PATH | DDD_NO_BROADCAST_SYSTEM;
    static_cast<void>(DefineDosDeviceW(removeFlags, deviceName.c_str(), rawTarget.c_str()));
}

// Refused means the sandbox guard rejected the alias (the root is not inside a marked sandbox): callers fail the
// case. Only Unreachable (the share is not served on this machine) is an environment skip.
enum class LoopbackShareStatus : uint8_t
{
    Reachable,
    Unreachable,
    Refused,
};

[[nodiscard]] LoopbackShareStatus ProbeLoopbackShare(const std::filesystem::path& sandboxRoot) noexcept
{
    const std::filesystem::path alias = LoopbackShareAlias(sandboxRoot);
    if (alias.empty())
    {
        return LoopbackShareStatus::Refused;
    }
    return GetFileAttributesW(alias.c_str()) != INVALID_FILE_ATTRIBUTES ? LoopbackShareStatus::Reachable : LoopbackShareStatus::Unreachable;
}

bool RecreateEmptyDirectory(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    constexpr DWORD kRetryDelayMs = 50u;
    const ULONGLONG deadline      = ::GetTickCount64() + SelfTest::ScaleTimeout(6'000u);
    do
    {
        static_cast<void>(SelfTest::RemoveAll(path));

        const std::filesystem::path extendedPath(MakeExtendedPathForSelfTest(path));
        std::error_code ec;
        ec.clear();
        static_cast<void>(std::filesystem::create_directories(extendedPath, ec));
        if (ec)
        {
            ::Sleep(kRetryDelayMs);
            continue;
        }

        const bool empty = std::filesystem::is_empty(extendedPath, ec);
        if (! ec && empty)
        {
            return true;
        }

        ::Sleep(kRetryDelayMs);
    } while (::GetTickCount64() < deadline);

    return false;
}

std::vector<std::filesystem::path> CollectFiles(const std::filesystem::path& dir, size_t maxCount) noexcept
{
    std::vector<std::filesystem::path> out;
    std::error_code ec;
    for (std::filesystem::directory_iterator it(dir, ec); ! ec && it != std::filesystem::directory_iterator{}; it.increment(ec))
    {
        const auto& entry = *it;
        const bool isFile = entry.is_regular_file(ec);
        if (ec)
        {
            break;
        }
        if (! isFile)
        {
            continue;
        }

        out.push_back(entry.path());
        if (out.size() >= maxCount)
        {
            break;
        }
    }
    return out;
}

size_t CountFiles(const std::filesystem::path& dir) noexcept
{
    size_t count = 0;
    std::error_code ec;
    const std::filesystem::path extendedDir(MakeExtendedPathForSelfTest(dir));
    for (std::filesystem::directory_iterator it(extendedDir, ec); ! ec && it != std::filesystem::directory_iterator{}; it.increment(ec))
    {
        if (it->is_regular_file(ec))
        {
            ++count;
        }
    }
    return count;
}

size_t CountFilesRecursive(const std::filesystem::path& dir) noexcept
{
    size_t count = 0;
    std::error_code ec;
    const std::filesystem::path extendedDir(MakeExtendedPathForSelfTest(dir));
    for (std::filesystem::recursive_directory_iterator it(extendedDir, std::filesystem::directory_options::skip_permission_denied, ec);
         ! ec && it != std::filesystem::recursive_directory_iterator{};
         it.increment(ec))
    {
        if (it->is_regular_file(ec))
        {
            ++count;
        }
    }
    return count;
}

struct FileOpsRecursiveProgressRecorder final : IFileSystemCallback
{
    std::wstring matchingNeedle;
    std::array<uint64_t, 32> streamIds{};
    std::array<uint64_t, 32> matchingStreamIds{};
    size_t streamCount          = 0;
    size_t matchingStreamCount  = 0;
    uint64_t progressCount      = 0;
    uint64_t matchingProgress   = 0;
    uint64_t completedCount     = 0;
    HRESULT lastCompletedStatus = E_PENDING;
    std::optional<FileSystemItemMutationResult> lastMutationResult;
    std::array<std::optional<FileSystemItemMutationResult>, 32> itemMutationResults{};
    std::array<HRESULT, 32> itemCompletedStatuses{};
    std::array<uint32_t, 32> itemCompletionCounts{};
    DWORD callerThreadId               = 0;
    bool matchingProgressOnOtherThread = false;
    uint64_t issueCount                = 0;
    HRESULT lastIssueStatus            = E_PENDING;
    bool cancelAfterProgress           = false;
    // Cancels at the first report that carries an item's written bytes; a file's announcement at zero
    // bytes (I35 R10.1) comes before its destination exists.
    bool cancelAfterTransferredProgress = false;
    bool cancelRequested                = false;
    FileSystemIssueAction issueAction   = FileSystemIssueAction::Skip;
    mutable std::mutex mutex;

    static void RecordStream(std::array<uint64_t, 32>& ids, size_t& count, uint64_t streamId) noexcept
    {
        for (size_t i = 0; i < count; ++i)
        {
            if (ids[i] == streamId)
            {
                return;
            }
        }

        if (count < ids.size())
        {
            ids[count++] = streamId;
        }
    }

    HRESULT STDMETHODCALLTYPE FileSystemProgress([[maybe_unused]] FileSystemOperation operationType,
                                                 [[maybe_unused]] unsigned long totalItems,
                                                 [[maybe_unused]] unsigned long completedItems,
                                                 [[maybe_unused]] uint64_t totalBytes,
                                                 [[maybe_unused]] uint64_t completedBytes,
                                                 const wchar_t* currentSourcePath,
                                                 [[maybe_unused]] const wchar_t* currentDestinationPath,
                                                 uint64_t currentItemTotalBytes,
                                                 uint64_t currentItemCompletedBytes,
                                                 [[maybe_unused]] FileSystemOptions* options,
                                                 uint64_t progressStreamId,
                                                 [[maybe_unused]] void* cookie) noexcept override
    {
        std::scoped_lock lock(mutex);
        ++progressCount;
        if (currentSourcePath && currentItemTotalBytes > 0)
        {
            RecordStream(streamIds, streamCount, progressStreamId);
            if (! matchingNeedle.empty())
            {
                const std::wstring_view sourceView{currentSourcePath};
                if (sourceView.find(std::wstring_view{matchingNeedle}) != std::wstring_view::npos)
                {
                    ++matchingProgress;
                    RecordStream(matchingStreamIds, matchingStreamCount, progressStreamId);
                    matchingProgressOnOtherThread = matchingProgressOnOtherThread || (callerThreadId != 0 && ::GetCurrentThreadId() != callerThreadId);
                }
            }
            if (cancelAfterProgress || (cancelAfterTransferredProgress && currentItemCompletedBytes > 0u))
            {
                cancelRequested = true;
            }
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemItemCompleted([[maybe_unused]] FileSystemOperation operationType,
                                                      unsigned long itemIndex,
                                                      [[maybe_unused]] const wchar_t* sourcePath,
                                                      [[maybe_unused]] const wchar_t* destinationPath,
                                                      HRESULT status,
                                                      const FileSystemItemMutationResult* mutationResult,
                                                      [[maybe_unused]] FileSystemOptions* options,
                                                      [[maybe_unused]] void* cookie) noexcept override
    {
        std::scoped_lock lock(mutex);
        ++completedCount;
        lastCompletedStatus = status;
        lastMutationResult  = FileSystemRouteContract::SnapshotItemMutationResult(mutationResult);
        if (itemIndex < itemMutationResults.size())
        {
            ++itemCompletionCounts[itemIndex];
            itemCompletedStatuses[itemIndex] = status;
            itemMutationResults[itemIndex]   = lastMutationResult;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* cancel, [[maybe_unused]] void* cookie) noexcept override
    {
        if (! cancel)
        {
            return E_POINTER;
        }

        std::scoped_lock lock(mutex);
        *cancel = cancelRequested ? TRUE : FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemIssue([[maybe_unused]] FileSystemOperation operationType,
                                              [[maybe_unused]] const wchar_t* sourcePath,
                                              [[maybe_unused]] const wchar_t* destinationPath,
                                              [[maybe_unused]] HRESULT status,
                                              FileSystemIssueAction* action,
                                              IFileSystemBoundObject** expectedDestination,
                                              [[maybe_unused]] FileSystemOptions* options,
                                              [[maybe_unused]] void* cookie) noexcept override
    {
        if (! action || ! expectedDestination)
        {
            return E_POINTER;
        }
        *expectedDestination = nullptr;

        std::scoped_lock lock(mutex);
        ++issueCount;
        lastIssueStatus = status;
        *action         = issueAction;
        return S_OK;
    }
};

struct FileOpsRecorderOperationControl final : IFileSystemOperationControl
{
    explicit FileOpsRecorderOperationControl(FileOpsRecursiveProgressRecorder& recorder) noexcept : _recorder(recorder)
    {
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, [[maybe_unused]] void* cookie) noexcept override
    {
        return _recorder.FileSystemShouldCancel(abort, nullptr);
    }

    HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, [[maybe_unused]] void* cookie) noexcept override
    {
        if (mode == nullptr)
        {
            return E_POINTER;
        }
        *mode = FILESYSTEM_DISCOVERY_JUST_IN_TIME;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress* progress, [[maybe_unused]] void* cookie) noexcept override
    {
        return progress != nullptr && progress->sizeBytes == sizeof(FileSystemDiscoveryProgress) ? S_OK : E_INVALIDARG;
    }

private:
    FileOpsRecursiveProgressRecorder& _recorder;
};

bool WaitForFileCount(const std::filesystem::path& dir, size_t expectedCount, ULONGLONG timeoutMs) noexcept
{
    const ULONGLONG startTick = GetTickCount64();
    for (;;)
    {
        if (CountFiles(dir) == expectedCount)
        {
            return true;
        }

        const ULONGLONG nowTick = GetTickCount64();
        if ((nowTick - startTick) >= timeoutMs)
        {
            return CountFiles(dir) == expectedCount;
        }

        ::Sleep(50);
    }
}

bool WriteTestFile(const std::filesystem::path& path, size_t bytes) noexcept
{
    std::error_code ec;
    const std::filesystem::path parent = path.parent_path();
    if (! parent.empty())
    {
        std::filesystem::create_directories(std::filesystem::path(MakeExtendedPathForSelfTest(parent)), ec);
    }

    const std::wstring extendedPath = MakeExtendedPathForSelfTest(path);

    DWORD lastError = 0;
    wil::unique_handle h;
    for (int attempt = 0; attempt < 20; ++attempt)
    {
        h.reset(CreateFileW(extendedPath.c_str(),
                            GENERIC_WRITE,
                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                            nullptr,
                            CREATE_ALWAYS,
                            FILE_ATTRIBUTE_NORMAL,
                            nullptr));
        if (h)
        {
            break;
        }

        lastError = GetLastError();
        if (lastError == ERROR_ACCESS_DENIED)
        {
            static_cast<void>(SetFileAttributesW(extendedPath.c_str(), FILE_ATTRIBUTE_NORMAL));
        }

        if (lastError != ERROR_SHARING_VIOLATION && lastError != ERROR_LOCK_VIOLATION && lastError != ERROR_ACCESS_DENIED)
        {
            break;
        }

        ::Sleep(50);
    }

    if (! h)
    {
        SetLastError(lastError);
        return false;
    }

    std::vector<unsigned char> buffer;
    buffer.resize(std::min<size_t>(bytes, 64 * 1024));
    for (size_t i = 0; i < buffer.size(); ++i)
    {
        buffer[i] = static_cast<unsigned char>((i * 131u) ^ 0x5Au);
    }

    size_t remaining = bytes;
    while (remaining > 0)
    {
        const DWORD chunk = static_cast<DWORD>(std::min<size_t>(remaining, buffer.size()));
        DWORD written     = 0;
        if (! WriteFile(h.get(), buffer.data(), chunk, &written, nullptr) || written != chunk)
        {
            return false;
        }
        remaining -= chunk;
    }

    return true;
}

bool WriteFilledTestFile(const std::filesystem::path& path, size_t bytes, unsigned char value) noexcept
{
    std::error_code ec;
    const std::filesystem::path parent = path.parent_path();
    if (! parent.empty())
    {
        std::filesystem::create_directories(parent, ec);
    }

    wil::unique_handle h(CreateFileW(
        path.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! h)
    {
        return false;
    }

    std::vector<unsigned char> buffer(std::min<size_t>(bytes, 64 * 1024), value);
    size_t remaining = bytes;
    while (remaining > 0)
    {
        const DWORD chunk = static_cast<DWORD>(std::min<size_t>(remaining, buffer.size()));
        DWORD written     = 0;
        if (! WriteFile(h.get(), buffer.data(), chunk, &written, nullptr) || written != chunk)
        {
            return false;
        }
        remaining -= chunk;
    }

    return true;
}

bool SeedCopyKnobSourceFiles(const std::filesystem::path& src) noexcept
{
    if (! RecreateEmptyDirectory(src))
    {
        return false;
    }

    for (int i = 0; i < 40; ++i)
    {
        const std::filesystem::path file = src / std::format(L"small_{:03}.bin", i);
        if (! WriteTestFile(file, 4096))
        {
            return false;
        }
    }

    for (int i = 0; i < 3; ++i)
    {
        const std::filesystem::path file = src / std::format(L"medium_{:03}.bin", i);
        if (! WriteTestFile(file, 2 * 1024 * 1024))
        {
            return false;
        }
    }

    return true;
}

bool WriteAllToHandle(HANDLE handle, const void* data, size_t bytes) noexcept
{
    if (! handle || handle == INVALID_HANDLE_VALUE || ! data)
    {
        return false;
    }

    const auto* src  = static_cast<const unsigned char*>(data);
    size_t remaining = bytes;
    while (remaining > 0)
    {
        const DWORD chunk = static_cast<DWORD>(std::min<size_t>(remaining, static_cast<size_t>(std::numeric_limits<DWORD>::max())));
        DWORD written     = 0;
        if (! WriteFile(handle, src, chunk, &written, nullptr) || written != chunk)
        {
            return false;
        }
        src += chunk;
        remaining -= chunk;
    }

    return true;
}

[[maybe_unused]] bool VerifyPatternBytes(const unsigned char* data, size_t size, uint64_t baseOffset) noexcept
{
    if (! data)
    {
        return false;
    }

    for (size_t i = 0; i < size; ++i)
    {
        const unsigned char expected = static_cast<unsigned char>((baseOffset + i) & 0xFFu);
        if (data[i] != expected)
        {
            return false;
        }
    }

    return true;
}

#pragma pack(push, 1)
struct ZipLocalFileHeader
{
    uint32_t signature;
    uint16_t versionNeeded;
    uint16_t flags;
    uint16_t compressionMethod;
    uint16_t lastModTime;
    uint16_t lastModDate;
    uint32_t crc32;
    uint32_t compressedSize;
    uint32_t uncompressedSize;
    uint16_t fileNameLength;
    uint16_t extraFieldLength;
};

struct ZipDataDescriptor
{
    uint32_t signature;
    uint32_t crc32;
    uint32_t compressedSize;
    uint32_t uncompressedSize;
};

struct ZipCentralDirectoryHeader
{
    uint32_t signature;
    uint16_t versionMadeBy;
    uint16_t versionNeeded;
    uint16_t flags;
    uint16_t compressionMethod;
    uint16_t lastModTime;
    uint16_t lastModDate;
    uint32_t crc32;
    uint32_t compressedSize;
    uint32_t uncompressedSize;
    uint16_t fileNameLength;
    uint16_t extraFieldLength;
    uint16_t fileCommentLength;
    uint16_t diskNumberStart;
    uint16_t internalFileAttributes;
    uint32_t externalFileAttributes;
    uint32_t localHeaderOffset;
};

struct ZipEndOfCentralDirectory
{
    uint32_t signature;
    uint16_t diskNumber;
    uint16_t centralDirDiskNumber;
    uint16_t entriesThisDisk;
    uint16_t entriesTotal;
    uint32_t centralDirSize;
    uint32_t centralDirOffset;
    uint16_t commentLength;
};
#pragma pack(pop)

static_assert(sizeof(ZipLocalFileHeader) == 30);
static_assert(sizeof(ZipDataDescriptor) == 16);
static_assert(sizeof(ZipCentralDirectoryHeader) == 46);
static_assert(sizeof(ZipEndOfCentralDirectory) == 22);

uint32_t UpdateCrc32(uint32_t crc, const unsigned char* data, size_t size) noexcept
{
    static const std::array<uint32_t, 256> table = []() noexcept
    {
        std::array<uint32_t, 256> t{};
        for (uint32_t i = 0; i < t.size(); ++i)
        {
            uint32_t c = i;
            for (int bit = 0; bit < 8; ++bit)
            {
                c = (c & 1u) ? (0xEDB88320u ^ (c >> 1u)) : (c >> 1u);
            }
            t[i] = c;
        }
        return t;
    }();

    uint32_t c = crc;
    for (size_t i = 0; i < size; ++i)
    {
        c = table[(c ^ data[i]) & 0xFFu] ^ (c >> 8u);
    }
    return c;
}

[[maybe_unused]] bool CreateZipArchiveWithStoredPatternFile(const std::filesystem::path& zipPath, std::string_view entryName, uint64_t fileSizeBytes) noexcept
{
    if (entryName.empty())
    {
        return false;
    }

    if (entryName.size() > static_cast<size_t>(std::numeric_limits<uint16_t>::max()))
    {
        return false;
    }

    if (fileSizeBytes > static_cast<uint64_t>(std::numeric_limits<uint32_t>::max()))
    {
        return false;
    }

    std::error_code ec;
    const std::filesystem::path parent = zipPath.parent_path();
    if (! parent.empty())
    {
        std::filesystem::create_directories(parent, ec);
    }

    wil::unique_handle h(CreateFileW(
        zipPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! h)
    {
        return false;
    }

    constexpr uint32_t kLocalSig = 0x04034B50u;
    constexpr uint32_t kCenSig   = 0x02014B50u;
    constexpr uint32_t kEocdSig  = 0x06054B50u;
    constexpr uint32_t kDescSig  = 0x08074B50u;
    constexpr uint16_t kVer20    = 20u;

    const uint16_t nameLen = static_cast<uint16_t>(entryName.size());

    ZipLocalFileHeader local{};
    local.signature         = kLocalSig;
    local.versionNeeded     = kVer20;
    local.flags             = 0x0008u; // data descriptor present
    local.compressionMethod = 0u;      // store
    local.lastModTime       = 0u;
    local.lastModDate       = 0u;
    local.crc32             = 0u;
    local.compressedSize    = 0u;
    local.uncompressedSize  = 0u;
    local.fileNameLength    = nameLen;
    local.extraFieldLength  = 0u;

    if (! WriteAllToHandle(h.get(), &local, sizeof(local)) || ! WriteAllToHandle(h.get(), entryName.data(), entryName.size()))
    {
        return false;
    }

    std::vector<unsigned char> buffer;
    buffer.resize(256u * 1024u);

    uint32_t crc    = 0xFFFFFFFFu;
    uint64_t offset = 0;
    while (offset < fileSizeBytes)
    {
        const size_t toWrite = static_cast<size_t>(std::min<uint64_t>(static_cast<uint64_t>(buffer.size()), fileSizeBytes - offset));
        for (size_t i = 0; i < toWrite; ++i)
        {
            buffer[i] = static_cast<unsigned char>((offset + i) & 0xFFu);
        }

        crc = UpdateCrc32(crc, buffer.data(), toWrite);
        if (! WriteAllToHandle(h.get(), buffer.data(), toWrite))
        {
            return false;
        }

        offset += static_cast<uint64_t>(toWrite);
    }
    crc ^= 0xFFFFFFFFu;

    const uint32_t size32 = static_cast<uint32_t>(fileSizeBytes);
    ZipDataDescriptor desc{};
    desc.signature        = kDescSig;
    desc.crc32            = crc;
    desc.compressedSize   = size32;
    desc.uncompressedSize = size32;
    if (! WriteAllToHandle(h.get(), &desc, sizeof(desc)))
    {
        return false;
    }

    const uint32_t centralDirOffset = static_cast<uint32_t>(sizeof(local) + static_cast<size_t>(nameLen) + static_cast<size_t>(fileSizeBytes) + sizeof(desc));

    ZipCentralDirectoryHeader cen{};
    cen.signature              = kCenSig;
    cen.versionMadeBy          = kVer20;
    cen.versionNeeded          = kVer20;
    cen.flags                  = local.flags;
    cen.compressionMethod      = local.compressionMethod;
    cen.lastModTime            = local.lastModTime;
    cen.lastModDate            = local.lastModDate;
    cen.crc32                  = crc;
    cen.compressedSize         = size32;
    cen.uncompressedSize       = size32;
    cen.fileNameLength         = nameLen;
    cen.extraFieldLength       = 0u;
    cen.fileCommentLength      = 0u;
    cen.diskNumberStart        = 0u;
    cen.internalFileAttributes = 0u;
    cen.externalFileAttributes = 0u;
    cen.localHeaderOffset      = 0u;

    if (! WriteAllToHandle(h.get(), &cen, sizeof(cen)) || ! WriteAllToHandle(h.get(), entryName.data(), entryName.size()))
    {
        return false;
    }

    const uint32_t centralDirSize = static_cast<uint32_t>(sizeof(cen) + static_cast<size_t>(nameLen));

    ZipEndOfCentralDirectory eocd{};
    eocd.signature            = kEocdSig;
    eocd.diskNumber           = 0u;
    eocd.centralDirDiskNumber = 0u;
    eocd.entriesThisDisk      = 1u;
    eocd.entriesTotal         = 1u;
    eocd.centralDirSize       = centralDirSize;
    eocd.centralDirOffset     = centralDirOffset;
    eocd.commentLength        = 0u;

    if (! WriteAllToHandle(h.get(), &eocd, sizeof(eocd)))
    {
        return false;
    }

    return true;
}

std::optional<FolderWindow::FileOperationState::Task::ConflictPromptState> TryGetConflictPromptCopy(FolderWindow::FileOperationState::Task* task) noexcept
{
    if (! task)
    {
        return std::nullopt;
    }

    std::scoped_lock lock(task->_conflictArbiter.mutex);
    // Conflict prompts publish an immediate non-actionable "Reading details" state before
    // no-follow metadata resolves. Most mutation selftests consume decisions, so expose only the
    // stable actionable prompt here. Presentation tests inspect the arbiter directly when they
    // need to prove that the provisional state is visible.
    if (! task->_conflictArbiter.prompt.active || task->_conflictArbiter.prompt.metadataLoading)
    {
        return std::nullopt;
    }

    return task->_conflictArbiter.prompt;
}

bool InvokePopupSelfTest(HWND popup, const FileOperationsPopupInternal::PopupSelfTestInvoke& invoke) noexcept
{
    if (! popup)
    {
        return false;
    }

    static_cast<void>(SendMessageW(popup, WndMsg::kFileOpsPopupSelfTestInvoke, 0, reinterpret_cast<LPARAM>(&invoke)));
    return true;
}

// FileOps policy: only the current UI thread's visible menu owned by this exact task popup.

// Resolve actual retained task controls by their stable task/action identity, never their row ordinal.
// UIA Button::Invoke requests focus internally; the mandatory harness blocker below
// prevents native activation while retaining the real hosted callback path.
struct FileOpsTaskAccessibilityProbe final
{
    bool traversalComplete = false;
    std::wstring statusName;
};

[[nodiscard]] bool FindFileOpsTaskAction(HWND popup,
                                         FileOperationsPopupInternal::PopupHitTest::Kind kind,
                                         uint64_t taskId,
                                         wil::com_ptr_nothrow<IInvokeProvider>& action,
                                         FileOpsTaskAccessibilityProbe* probe = nullptr)
{
    action.reset();
    if (probe)
        *probe = {};
    const std::wstring expectedId = std::format(L"FileOperations.Action.{}.{}.0", static_cast<uint32_t>(kind), taskId);
    const std::wstring statusId   = probe ? std::format(L"FileOperations.Task.{}.Status", taskId) : std::wstring{};
    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> root;
    root.attach(DxUi::CreateWindowHostAccessibilityProvider(popup));
    wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
    if (! root || FAILED(root.query_to(fragment.put())))
        return false;
    std::vector<wil::com_ptr_nothrow<IRawElementProviderFragment>> pending;
    pending.push_back(std::move(fragment));
    size_t visited = 0u;
    while (! pending.empty() && ++visited <= 512u)
    {
        auto current = std::move(pending.back());
        pending.pop_back();
        wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
        wil::unique_variant id;
        if (FAILED(current.query_to(simple.put())) || ! simple || FAILED(simple->GetPropertyValue(UIA_AutomationIdPropertyId, &id)))
            return false;
        if (probe && id.vt == VT_BSTR && id.bstrVal && statusId == std::wstring_view(id.bstrVal, SysStringLen(id.bstrVal)))
        {
            wil::unique_variant name;
            if (FAILED(simple->GetPropertyValue(UIA_NamePropertyId, &name)) || name.vt != VT_BSTR || ! name.bstrVal)
                return false;
            probe->statusName.assign(name.bstrVal, SysStringLen(name.bstrVal));
        }
        if (id.vt == VT_BSTR && id.bstrVal && expectedId == std::wstring_view(id.bstrVal, SysStringLen(id.bstrVal)))
        {
            wil::unique_variant type;
            wil::unique_variant enabled;
            wil::com_ptr_nothrow<IUnknown> pattern;
            return SUCCEEDED(simple->GetPropertyValue(UIA_ControlTypePropertyId, &type)) && type.vt == VT_I4 && type.lVal == UIA_ButtonControlTypeId &&
                   SUCCEEDED(simple->GetPropertyValue(UIA_IsEnabledPropertyId, &enabled)) && enabled.vt == VT_BOOL && enabled.boolVal == VARIANT_TRUE &&
                   SUCCEEDED(simple->GetPatternProvider(UIA_InvokePatternId, pattern.put())) && pattern && SUCCEEDED(pattern.query_to(action.put())) && action;
        }
        wil::com_ptr_nothrow<IRawElementProviderFragment> child;
        if (FAILED(current->Navigate(NavigateDirection_FirstChild, child.put())))
            return false;
        size_t siblings = 0u;
        while (child)
        {
            if (++siblings > 128u || visited + pending.size() >= 512u)
                return false;
            wil::com_ptr_nothrow<IRawElementProviderFragment> next;
            if (FAILED(child->Navigate(NavigateDirection_NextSibling, next.put())))
                return false;
            pending.push_back(std::move(child));
            child = std::move(next);
        }
    }
    // Absence is evidence only after a successful complete bounded traversal.
    // The caller can additionally require the task's positive status witness.
    if (probe)
        probe->traversalComplete = pending.empty();
    return false;
}

[[nodiscard]] bool InvokeFileOpsTaskActionWithoutActivation(IInvokeProvider& action, bool permitRetired = false) noexcept
{
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread())
    {
        AppendLog(L"QueueControlInvoke: refused because the mandatory no-activation blocker is absent.");
        return false; // Never dispatch a focus-requesting Button outside the nonactivating lane.
    }
    // A retained provider can outlive a row rebuild. Record the identity it resolves
    // immediately before invocation, without substituting a newly found provider.
    wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
    wil::unique_variant automationId;
    HRESULT identityHr = action.QueryInterface(IID_PPV_ARGS(simple.put()));
    if (SUCCEEDED(identityHr) && simple)
        identityHr = simple->GetPropertyValue(UIA_AutomationIdPropertyId, &automationId);
    const std::wstring resolvedId =
        automationId.vt == VT_BSTR && automationId.bstrVal ? std::wstring(automationId.bstrVal, SysStringLen(automationId.bstrVal)) : std::wstring{};
    const HWND focusBefore      = GetFocus();
    const HWND foregroundBefore = GetForegroundWindow();
    const HRESULT result        = action.Invoke();
    const HWND focusAfter       = GetFocus();
    const HWND foregroundAfter  = GetForegroundWindow();
    DWORD foregroundProcess     = 0u;
    if (foregroundAfter)
        GetWindowThreadProcessId(foregroundAfter, &foregroundProcess);
    const bool acquiredForeground = foregroundBefore != foregroundAfter && foregroundProcess == GetCurrentProcessId();
    const bool acceptedResult     = SUCCEEDED(result) || (permitRetired && (result == UIA_E_ELEMENTNOTAVAILABLE || result == UIA_E_NOTSUPPORTED));
    AppendLog(std::format(L"QueueControlInvoke: permitRetired={} identityHr=0x{:08X} resolvedId='{}' invokeHr=0x{:08X} acceptedResult={} "
                          L"focusBefore=0x{:X} focusAfter=0x{:X} foregroundBefore=0x{:X} foregroundAfter=0x{:X} foregroundPid={} "
                          L"testPid={} acquiredForeground={}",
                          permitRetired,
                          static_cast<uint32_t>(identityHr),
                          resolvedId,
                          static_cast<uint32_t>(result),
                          acceptedResult,
                          reinterpret_cast<uintptr_t>(focusBefore),
                          reinterpret_cast<uintptr_t>(focusAfter),
                          reinterpret_cast<uintptr_t>(foregroundBefore),
                          reinterpret_cast<uintptr_t>(foregroundAfter),
                          foregroundProcess,
                          GetCurrentProcessId(),
                          acquiredForeground));
    return focusAfter == focusBefore && ! acquiredForeground && acceptedResult;
}

// Invoke queues the real menu action; the phased test observes callback completion on a later tick.

// This fixture has unique history labels. Its caller verifies the committed
// destination and admission key after the real keyboard callback completes.

bool TryGetPopupTaskSnapshot(FolderWindow::FileOperationState* fileOps, uint64_t taskId, FileOperationsPopupInternal::TaskSnapshot& out) noexcept
{
    if (! fileOps || taskId == 0)
    {
        return false;
    }

    const HWND popup = fileOps->GetPopupHwndForSelfTest();
    return popup && DebugGetFileOperationsPopupTaskSnapshot(popup, taskId, out);
}

bool PromptHasAction(const FolderWindow::FileOperationState::Task::ConflictPromptState& prompt,
                     FolderWindow::FileOperationState::Task::ConflictAction action) noexcept
{
    for (size_t i = 0; i < prompt.actionCount; ++i)
    {
        if (prompt.actions[i] == action)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring MakeExtendedPathForSelfTest(const std::filesystem::path& inputPath)
{
    std::wstring normalized = inputPath.native();
    std::ranges::replace(normalized, L'/', L'\\');
    if (normalized.empty())
    {
        normalized = L".";
    }

    if (normalized.rfind(L"\\\\?\\", 0) != 0)
    {
        const DWORD required = GetFullPathNameW(normalized.c_str(), 0, nullptr, nullptr);
        if (required != 0)
        {
            std::wstring absolute(static_cast<size_t>(required) + 1u, L'\0');
            const DWORD written = GetFullPathNameW(normalized.c_str(), static_cast<DWORD>(absolute.size()), absolute.data(), nullptr);
            if (written != 0)
            {
                absolute.resize(static_cast<size_t>(written));
                normalized = std::move(absolute);
            }
        }
    }

    if (normalized.rfind(L"\\\\?\\", 0) == 0)
    {
        return normalized;
    }
    if (normalized.rfind(L"\\\\", 0) == 0)
    {
        return std::wstring(L"\\\\?\\UNC\\") + normalized.substr(2);
    }
    return std::wstring(L"\\\\?\\") + normalized;
}

[[nodiscard]] bool FileSizeEquals(const std::filesystem::path& path, uint64_t expectedBytes) noexcept
{
    const std::wstring extendedPath = MakeExtendedPathForSelfTest(path);
    wil::unique_handle file(CreateFileW(extendedPath.c_str(),
                                        FILE_READ_ATTRIBUTES,
                                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                        nullptr,
                                        OPEN_EXISTING,
                                        FILE_ATTRIBUTE_NORMAL,
                                        nullptr));
    if (! file)
    {
        return false;
    }

    LARGE_INTEGER size{};
    if (! GetFileSizeEx(file.get(), &size) || size.QuadPart < 0)
    {
        return false;
    }

    return static_cast<uint64_t>(size.QuadPart) == expectedBytes;
}

[[nodiscard]] bool FilesEqualBytes(const std::filesystem::path& lhs, const std::filesystem::path& rhs) noexcept
{
    const auto openForRead = [&](const std::filesystem::path& inputPath) noexcept -> wil::unique_handle
    {
        const std::wstring extendedPath = MakeExtendedPathForSelfTest(inputPath);
        return wil::unique_handle(CreateFileW(extendedPath.c_str(),
                                              GENERIC_READ,
                                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                              nullptr,
                                              OPEN_EXISTING,
                                              FILE_ATTRIBUTE_NORMAL,
                                              nullptr));
    };

    wil::unique_handle lhsFile = openForRead(lhs);
    wil::unique_handle rhsFile = openForRead(rhs);
    if (! lhsFile || ! rhsFile)
    {
        return false;
    }

    LARGE_INTEGER lhsSize{};
    LARGE_INTEGER rhsSize{};
    if (! GetFileSizeEx(lhsFile.get(), &lhsSize) || ! GetFileSizeEx(rhsFile.get(), &rhsSize) || lhsSize.QuadPart < 0 || rhsSize.QuadPart < 0 ||
        lhsSize.QuadPart != rhsSize.QuadPart || static_cast<uint64_t>(lhsSize.QuadPart) > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
    {
        return false;
    }

    const auto readAll = [](HANDLE file, std::vector<char>& bytes) noexcept -> bool
    {
        size_t offset = 0;
        while (offset < bytes.size())
        {
            const DWORD toRead = static_cast<DWORD>(std::min<size_t>(bytes.size() - offset, static_cast<size_t>((std::numeric_limits<DWORD>::max)())));
            DWORD read         = 0;
            if (! ReadFile(file, bytes.data() + offset, toRead, &read, nullptr) || read == 0)
            {
                return false;
            }
            offset += read;
        }
        return true;
    };

    std::vector<char> lhsBytes(static_cast<size_t>(lhsSize.QuadPart));
    std::vector<char> rhsBytes(static_cast<size_t>(rhsSize.QuadPart));
    if (! readAll(lhsFile.get(), lhsBytes) || ! readAll(rhsFile.get(), rhsBytes))
    {
        return false;
    }

    return lhsBytes == rhsBytes;
}

bool CreateDeleteTree(const std::filesystem::path& root, int directories, int filesPerDirectory, size_t bytesPerFile) noexcept
{
    if (! RecreateEmptyDirectory(root))
    {
        return false;
    }

    for (int d = 0; d < directories; ++d)
    {
        const std::filesystem::path sub = root / std::format(L"dir_{:02}", d);
        std::error_code ec;
        std::filesystem::create_directories(sub, ec);
        if (ec)
        {
            return false;
        }

        for (int f = 0; f < filesPerDirectory; ++f)
        {
            const std::filesystem::path file = sub / std::format(L"file_{:03}.txt", f);
            if (! WriteTestFile(file, bytesPerFile))
            {
                return false;
            }
        }
    }

    return true;
}

[[nodiscard]] bool ReadUtf16TextFile(const std::filesystem::path& path, std::wstring& textOut) noexcept
{
    textOut.clear();

    std::ifstream stream(path, std::ios::binary);
    if (! stream)
    {
        return false;
    }

    stream.seekg(0, std::ios::end);
    const std::streamoff size = stream.tellg();
    if (size < 0 || (size % static_cast<std::streamoff>(sizeof(wchar_t))) != 0)
    {
        return false;
    }

    stream.seekg(0, std::ios::beg);
    std::wstring text(static_cast<size_t>(size / static_cast<std::streamoff>(sizeof(wchar_t))), L'\0');
    if (! text.empty() && ! stream.read(reinterpret_cast<char*>(text.data()), size))
    {
        textOut.clear();
        return false;
    }

    if (! text.empty() && text.front() == 0xFEFF)
    {
        text.erase(text.begin());
    }

    textOut = std::move(text);
    return true;
}

[[nodiscard]] bool TryFindDiagnosticLine(
    std::wstring_view diagnosticsText, uint64_t taskId, std::wstring_view category, std::wstring& lineOut, std::wstring_view requiredNeedle = {}) noexcept
{
    lineOut.clear();
    const std::wstring taskNeedle     = std::format(L"\"task\":{}", taskId);
    const std::wstring categoryNeedle = std::format(L"\"category\":\"{}\"", category);

    // The diagnostics log accumulates across app runs and task ids restart per run, so the same
    // "task":N can appear for unrelated operations from earlier runs. Keep the LAST match: that
    // is the current run's line.
    bool found       = false;
    size_t searchPos = 0;
    while ((searchPos = diagnosticsText.find(taskNeedle, searchPos)) != std::wstring_view::npos)
    {
        const size_t lineStart            = diagnosticsText.rfind(L'\n', searchPos);
        const size_t contentStart         = lineStart == std::wstring_view::npos ? 0u : (lineStart + 1u);
        const size_t lineEnd              = diagnosticsText.find(L'\n', searchPos);
        const size_t contentEnd           = lineEnd == std::wstring_view::npos ? diagnosticsText.size() : lineEnd;
        const std::wstring_view candidate = diagnosticsText.substr(contentStart, contentEnd - contentStart);
        if (candidate.find(categoryNeedle) != std::wstring_view::npos && (requiredNeedle.empty() || candidate.find(requiredNeedle) != std::wstring_view::npos))
        {
            lineOut.assign(candidate);
            if (! lineOut.empty() && lineOut.back() == L'\r')
            {
                lineOut.pop_back();
            }
            found = true;
        }

        searchPos += taskNeedle.size();
    }

    return found;
}

bool CreateDeleteTreeWithRetry(
    const std::filesystem::path& root, int directories, int filesPerDirectory, size_t bytesPerFile, std::wstring_view label, int maxAttempts = 3) noexcept
{
    constexpr DWORD kRetryDelayMs = 200;
    for (int attempt = 1; attempt <= maxAttempts; ++attempt)
    {
        if (CreateDeleteTree(root, directories, filesPerDirectory, bytesPerFile))
        {
            return true;
        }

        if (attempt < maxAttempts)
        {
            AppendLog(std::format(L"Setup: retrying {} seed (attempt {}/{})", label, attempt + 1, maxAttempts));
            ::Sleep(kRetryDelayMs);
        }
    }

    return false;
}

bool CreateSiblingFiles(const std::filesystem::path& root,
                        int fileCount,
                        size_t bytesPerFile,
                        std::vector<std::filesystem::path>& outFiles,
                        std::wstring_view fileStem = L"batch") noexcept
{
    outFiles.clear();
    if (! RecreateEmptyDirectory(root))
    {
        return false;
    }

    if (fileCount <= 0)
    {
        return true;
    }

    outFiles.reserve(static_cast<size_t>(fileCount));
    for (int index = 0; index < fileCount; ++index)
    {
        const std::filesystem::path file = root / std::format(L"{}_{:04}.bin", fileStem, index);
        if (! WriteTestFile(file, bytesPerFile))
        {
            return false;
        }

        outFiles.push_back(file);
    }

    return true;
}

[[nodiscard]] size_t GetProcessThreadCount() noexcept
{
    const DWORD pid = GetCurrentProcessId();

    wil::unique_handle snapshot(CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0));
    if (! snapshot)
    {
        return 0;
    }

    THREADENTRY32 entry{};
    entry.dwSize = sizeof(entry);

    size_t count = 0;
    if (Thread32First(snapshot.get(), &entry))
    {
        do
        {
            if (entry.th32OwnerProcessID == pid)
            {
                ++count;
            }
            entry.dwSize = sizeof(entry);
        } while (Thread32Next(snapshot.get(), &entry));
    }

    return count;
}

struct Phase14ShutdownWork final
{
    FolderWindow::FileOperationState* fileOps = nullptr;
    std::atomic<bool>* done                   = nullptr;
};

void CALLBACK Phase14ShutdownCallback(PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
{
    std::unique_ptr<Phase14ShutdownWork> work(static_cast<Phase14ShutdownWork*>(context));
    if (! work)
    {
        return;
    }

    if (work->fileOps)
    {
        work->fileOps->Shutdown();
    }

    if (work->done)
    {
        work->done->store(true, std::memory_order_release);
    }
}

struct ReparsePointHeader
{
    DWORD tag        = 0;
    USHORT dataBytes = 0;
    USHORT reserved  = 0;
};
static_assert(sizeof(ReparsePointHeader) == 8);

struct MountPointReparseHeader
{
    USHORT substituteOffset = 0;
    USHORT substituteLength = 0;
    USHORT printOffset      = 0;
    USHORT printLength      = 0;
};
static_assert(sizeof(MountPointReparseHeader) == 8);

struct SymbolicLinkReparseHeader
{
    USHORT substituteOffset = 0;
    USHORT substituteLength = 0;
    USHORT printOffset      = 0;
    USHORT printLength      = 0;
    ULONG flags             = 0;
};
static_assert(sizeof(SymbolicLinkReparseHeader) == 12);

constexpr ULONG kSymlinkRelativeFlag = 0x00000001u;

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

[[nodiscard]] std::wstring NormalizePathForCompare(std::wstring path)
{
    std::ranges::replace(path, L'/', L'\\');

    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        path = std::wstring(L"\\\\") + path.substr(8);
    }
    else if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        path = path.substr(4);
    }

    size_t rootLength = 0;
    if (path.size() >= 2 && path[1] == L':')
    {
        rootLength = (path.size() >= 3 && IsPathSeparator(path[2])) ? 3u : 2u;
    }
    else if (path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\')
    {
        size_t firstSep  = path.find_first_of(L"\\/", 2);
        size_t secondSep = (firstSep == std::wstring::npos) ? std::wstring::npos : path.find_first_of(L"\\/", firstSep + 1);
        rootLength       = (secondSep == std::wstring::npos) ? path.size() : (secondSep + 1);
    }
    else if (! path.empty() && IsPathSeparator(path.front()))
    {
        rootLength = 1u;
    }

    while (path.size() > rootLength && ! path.empty() && IsPathSeparator(path.back()))
    {
        path.pop_back();
    }

    if (! path.empty())
    {
        if (path.size() > static_cast<size_t>(std::numeric_limits<int>::max() - 1))
        {
            return path;
        }

        std::wstring lower(path.size() + 1, L'\0');
        const int written =
            LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, path.c_str(), -1, lower.data(), static_cast<int>(lower.size()), nullptr, nullptr, 0);
        if (written > 0)
        {
            lower.resize(static_cast<size_t>(written) - 1);
            path = std::move(lower);
        }
    }
    return path;
}

[[nodiscard]] std::wstring NtPathToWin32Path(std::wstring_view path)
{
    if (path.rfind(L"\\??\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\??\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    return std::wstring(path);
}

[[nodiscard]] std::optional<std::wstring> TryGetDirectoryReparseTargetAbsolute(const std::filesystem::path& linkPath) noexcept
{
    wil::unique_handle handle(CreateFileW(linkPath.c_str(),
                                          FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
    if (! handle)
    {
        return std::nullopt;
    }

    alignas(8) std::array<std::byte, MAXIMUM_REPARSE_DATA_BUFFER_SIZE> buffer{};
    DWORD bytesReturned = 0;
    if (! DeviceIoControl(handle.get(), FSCTL_GET_REPARSE_POINT, nullptr, 0, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesReturned, nullptr))
    {
        return std::nullopt;
    }

    if (bytesReturned < sizeof(ReparsePointHeader))
    {
        return std::nullopt;
    }

    const auto* header        = reinterpret_cast<const ReparsePointHeader*>(buffer.data());
    const std::byte* payload  = buffer.data() + sizeof(ReparsePointHeader);
    const size_t payloadBytes = bytesReturned - sizeof(ReparsePointHeader);

    auto readPath = [&](USHORT offsetBytes, USHORT lengthBytes, size_t fixedHeaderBytes) -> std::optional<std::wstring>
    {
        if ((offsetBytes % sizeof(wchar_t)) != 0u || (lengthBytes % sizeof(wchar_t)) != 0u)
        {
            return std::nullopt;
        }
        if (payloadBytes < fixedHeaderBytes)
        {
            return std::nullopt;
        }
        const size_t pathBytes = payloadBytes - fixedHeaderBytes;
        if (offsetBytes > pathBytes || lengthBytes > pathBytes || (static_cast<size_t>(offsetBytes) + static_cast<size_t>(lengthBytes)) > pathBytes)
        {
            return std::nullopt;
        }
        const wchar_t* text = reinterpret_cast<const wchar_t*>(payload + fixedHeaderBytes + offsetBytes);
        return std::wstring(text, text + (lengthBytes / sizeof(wchar_t)));
    };

    if (header->tag == IO_REPARSE_TAG_MOUNT_POINT)
    {
        if (payloadBytes < sizeof(MountPointReparseHeader))
        {
            return std::nullopt;
        }

        const auto* mount = reinterpret_cast<const MountPointReparseHeader*>(payload);
        auto substitute   = readPath(mount->substituteOffset, mount->substituteLength, sizeof(MountPointReparseHeader));
        if (! substitute.has_value())
        {
            return std::nullopt;
        }

        std::wstring absolute = NtPathToWin32Path(substitute.value());
        absolute              = NormalizePathForCompare(absolute);
        return absolute;
    }

    if (header->tag == IO_REPARSE_TAG_SYMLINK)
    {
        if (payloadBytes < sizeof(SymbolicLinkReparseHeader))
        {
            return std::nullopt;
        }

        const auto* symlink = reinterpret_cast<const SymbolicLinkReparseHeader*>(payload);
        auto substitute     = readPath(symlink->substituteOffset, symlink->substituteLength, sizeof(SymbolicLinkReparseHeader));
        if (! substitute.has_value())
        {
            return std::nullopt;
        }

        std::wstring target = substitute.value();
        if ((symlink->flags & kSymlinkRelativeFlag) != 0u)
        {
            std::filesystem::path absolutePath = std::filesystem::path(linkPath).parent_path() / std::filesystem::path(target);
            target                             = absolutePath.lexically_normal().wstring();
        }
        else
        {
            target = NtPathToWin32Path(target);
        }

        target = NormalizePathForCompare(target);
        return target;
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<DWORD> TryGetReparseTag(const std::filesystem::path& path) noexcept
{
    wil::unique_handle handle(CreateFileW(path.c_str(),
                                          FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
    if (! handle)
    {
        return std::nullopt;
    }

    alignas(8) std::array<std::byte, MAXIMUM_REPARSE_DATA_BUFFER_SIZE> buffer{};
    DWORD bytesReturned = 0;
    if (! DeviceIoControl(handle.get(), FSCTL_GET_REPARSE_POINT, nullptr, 0, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesReturned, nullptr))
    {
        return std::nullopt;
    }

    if (bytesReturned < sizeof(ReparsePointHeader))
    {
        return std::nullopt;
    }

    const auto* header = reinterpret_cast<const ReparsePointHeader*>(buffer.data());
    return header->tag;
}

[[nodiscard]] bool TryCreateJunction(const std::filesystem::path& junctionPath, const std::filesystem::path& targetDirectoryPath) noexcept
{
    std::error_code ec;
    if (! Common::Testing::RequireSandboxedLinkTarget(junctionPath, targetDirectoryPath, ec))
    {
        return false;
    }

    // Junction must be an empty directory when applying the mount-point reparse buffer.
    std::filesystem::remove_all(junctionPath, ec);
    ec.clear();
    std::filesystem::create_directories(junctionPath, ec);
    if (ec)
    {
        return false;
    }

    const std::filesystem::path targetAbs = std::filesystem::absolute(targetDirectoryPath, ec);
    if (ec)
    {
        return false;
    }

    std::wstring target = targetAbs.wstring();
    if (target.empty())
    {
        return false;
    }
    if (target.back() != L'\\' && target.back() != L'/')
    {
        target.push_back(L'\\');
    }

    std::wstring substitute = L"\\??\\";
    substitute.append(target);

    const size_t substituteBytes = substitute.size() * sizeof(wchar_t);
    const size_t printBytes      = target.size() * sizeof(wchar_t);
    const size_t pathBufferBytes = substituteBytes + sizeof(wchar_t) + printBytes + sizeof(wchar_t);

    constexpr size_t kMountPointHeaderBytes = sizeof(USHORT) * 4; // offsets/lengths
    const size_t mountPointBytes            = kMountPointHeaderBytes + pathBufferBytes;
    if (mountPointBytes > static_cast<size_t>(std::numeric_limits<USHORT>::max()))
    {
        return false;
    }

    const size_t totalBytes = sizeof(ReparsePointHeader) + mountPointBytes;
    if (totalBytes > MAXIMUM_REPARSE_DATA_BUFFER_SIZE)
    {
        return false;
    }

    std::vector<std::byte> buffer(totalBytes);
    auto* header      = reinterpret_cast<ReparsePointHeader*>(buffer.data());
    header->tag       = IO_REPARSE_TAG_MOUNT_POINT;
    header->dataBytes = static_cast<USHORT>(mountPointBytes);
    header->reserved  = 0;

    auto* mountHeader             = reinterpret_cast<MountPointReparseHeader*>(buffer.data() + sizeof(ReparsePointHeader));
    mountHeader->substituteOffset = 0;
    mountHeader->substituteLength = static_cast<USHORT>(substituteBytes);
    mountHeader->printOffset      = static_cast<USHORT>(substituteBytes + sizeof(wchar_t));
    mountHeader->printLength      = static_cast<USHORT>(printBytes);

    std::byte* pathBuffer = buffer.data() + sizeof(ReparsePointHeader) + sizeof(MountPointReparseHeader);
    std::memcpy(pathBuffer, substitute.data(), substituteBytes);
    std::memset(pathBuffer + substituteBytes, 0, sizeof(wchar_t));
    std::memcpy(pathBuffer + substituteBytes + sizeof(wchar_t), target.data(), printBytes);
    std::memset(pathBuffer + substituteBytes + sizeof(wchar_t) + printBytes, 0, sizeof(wchar_t));

    wil::unique_handle handle(
        CreateFileW(junctionPath.c_str(), GENERIC_WRITE, 0, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS, nullptr));
    if (! handle)
    {
        return false;
    }

    DWORD ignored = 0;
    if (! DeviceIoControl(handle.get(), FSCTL_SET_REPARSE_POINT, buffer.data(), static_cast<DWORD>(buffer.size()), nullptr, 0, &ignored, nullptr))
    {
        return false;
    }

    return true;
}

[[nodiscard]] std::optional<std::wstring> TryGetVolumeGuidPathForPath(const std::filesystem::path& path) noexcept
{
    const std::wstring text = path.wstring();
    if (text.empty())
    {
        return std::nullopt;
    }

    std::array<wchar_t, MAX_PATH> volumePath{};
    if (GetVolumePathNameW(text.c_str(), volumePath.data(), static_cast<DWORD>(volumePath.size())) == FALSE)
    {
        return std::nullopt;
    }

    std::array<wchar_t, MAX_PATH> volumeGuidPath{};
    if (GetVolumeNameForVolumeMountPointW(volumePath.data(), volumeGuidPath.data(), static_cast<DWORD>(volumeGuidPath.size())) == FALSE)
    {
        return std::nullopt;
    }

    return std::wstring(volumeGuidPath.data());
}

void PruneEmptyAlternateVolumeSandboxParents(const std::filesystem::path& sandboxRoot) noexcept
{
    try
    {
        const std::filesystem::path suiteRoot   = sandboxRoot.parent_path();
        const std::filesystem::path scratchRoot = suiteRoot.parent_path();
        const std::filesystem::path runRoot     = scratchRoot.parent_path();
        const std::filesystem::path runsRoot    = runRoot.parent_path();
        const std::filesystem::path testRoot    = runsRoot.parent_path();

        const std::optional<std::filesystem::path> authorizedRoot =
            Common::Testing::GetDedicatedExternalTestSandboxBase(testRoot, Common::Testing::kTestSandboxDirectoryName);
        if (suiteRoot.empty() || scratchRoot.filename() != L"scratch" || runsRoot.filename() != L"runs" || ! authorizedRoot.has_value() ||
            ! Common::Testing::TestSandboxPathEquals(testRoot, authorizedRoot.value()) || ! Common::Testing::HasValidTestSandboxMarker(testRoot))
        {
            return;
        }

        const auto pruneIfEmpty = [](const std::filesystem::path& path) noexcept
        {
            std::error_code ec;
            if (! std::filesystem::is_directory(path, ec) || ec)
            {
                return false;
            }

            ec.clear();
            if (! std::filesystem::is_empty(path, ec) || ec)
            {
                return false;
            }

            ec.clear();
            return std::filesystem::remove(path, ec) && ! ec;
        };

        if (! pruneIfEmpty(suiteRoot))
        {
            return;
        }
        if (! pruneIfEmpty(scratchRoot))
        {
            return;
        }
        if (! pruneIfEmpty(runRoot) || ! pruneIfEmpty(runsRoot))
        {
            return;
        }

        std::error_code ec;
        const std::filesystem::path markerPath = testRoot / std::wstring(Common::Testing::kTestSandboxMarkerFileName);
        if (! std::filesystem::remove(markerPath, ec) || ec)
        {
            return;
        }
        static_cast<void>(pruneIfEmpty(testRoot));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Cleanup is best-effort at the noexcept self-test boundary; the next runner cleanup pass owns any leftover empty parents.
        Debug::Warning(L"FileOpsSelfTest: alternate-volume TestSandbox parent pruning failed.");
    }
}

[[nodiscard]] std::optional<std::filesystem::path> TryCreateAlternateWritableVolumeSelfTestRoot(const std::filesystem::path& referencePath,
                                                                                                std::wstring& detail,
                                                                                                bool allowOwnedVirtualDisk = false) noexcept
{
    detail.clear();

    std::array<wchar_t, 2> alternateVolumeAuthorization{};
    if (GetEnvironmentVariableW(
            L"REDSALAMANDER_TEST_ALLOW_ALTERNATE_VOLUME", alternateVolumeAuthorization.data(), static_cast<DWORD>(alternateVolumeAuthorization.size())) != 1u ||
        alternateVolumeAuthorization[0] != L'1')
    {
        detail = L"alternate-volume-test-root-not-authorized";
        return std::nullopt;
    }

    const std::optional<std::wstring> referenceVolume = TryGetVolumeGuidPathForPath(referencePath);
    if (! referenceVolume.has_value())
    {
        detail = L"reference-volume-unavailable";
        return std::nullopt;
    }

    const auto tryCandidateRoot = [&](const std::filesystem::path& root) -> std::optional<std::filesystem::path>
    {
        const std::wstring rootText = root.wstring();
        if (GetDriveTypeW(rootText.c_str()) != DRIVE_FIXED)
        {
            detail = std::format(L"alternate-volume-not-fixed-{}", rootText);
            return std::nullopt;
        }

        const std::optional<std::wstring> candidateVolume = TryGetVolumeGuidPathForPath(root);
        if (! candidateVolume.has_value())
        {
            detail = std::format(L"alternate-volume-unavailable-{}", rootText);
            return std::nullopt;
        }
        if (candidateVolume.value() == referenceVolume.value())
        {
            detail = L"alternate-volume-matches-primary";
            return std::nullopt;
        }

        const SelfTest::TestSandbox alternateVolumeSandbox =
            SelfTest::AcquireTestSandboxOnVolume(SelfTest::SelfTestSuite::FileOperations, L"real_cross_volume_move", root);
        if (! alternateVolumeSandbox.IsValid())
        {
            detail = std::format(L"test-sandbox-unavailable-{}", rootText);
            return std::nullopt;
        }

        const std::filesystem::path& candidate = alternateVolumeSandbox.root;
        const std::filesystem::path probe      = candidate / L"probe.tmp";
        {
            std::ofstream out(probe, std::ios::binary | std::ios::trunc);
            if (! out)
            {
                std::error_code ec;
                std::filesystem::remove_all(candidate, ec);
                PruneEmptyAlternateVolumeSandboxParents(candidate);
                detail = std::format(L"probe-open-failed-{}", rootText);
                return std::nullopt;
            }
            out << "probe";
        }

        std::error_code ec;
        std::filesystem::remove(probe, ec);
        return candidate;
    };

    const std::wstring configuredRootText = GetEnvVarTrimmed(kAlternateTestRootEnvVar);
    if (! configuredRootText.empty())
    {
        const std::filesystem::path configuredRoot = std::filesystem::path(configuredRootText).lexically_normal();
        const std::optional<std::filesystem::path> authorizedRoot =
            Common::Testing::GetDedicatedExternalTestSandboxBase(configuredRoot, Common::Testing::kTestSandboxDirectoryName);
        if (! authorizedRoot.has_value() || ! Common::Testing::TestSandboxPathEquals(configuredRoot, authorizedRoot.value()))
        {
            detail = L"configured-alternate-test-root-unauthorized";
            return std::nullopt;
        }
        return tryCandidateRoot(authorizedRoot.value().root_path());
    }

    const DWORD driveMask = GetLogicalDrives();
    if (driveMask == 0u)
    {
        detail = L"logical-drives-unavailable";
        return std::nullopt;
    }

    for (wchar_t drive = L'A'; drive <= L'Z'; ++drive)
    {
        const DWORD bit = 1u << static_cast<DWORD>(drive - L'A');
        if ((driveMask & bit) == 0u)
        {
            continue;
        }

        std::wstring rootText;
        rootText.push_back(drive);
        rootText.append(L":\\");
        if (GetDriveTypeW(rootText.c_str()) != DRIVE_FIXED)
        {
            continue;
        }

        const std::filesystem::path root(rootText);
        const std::optional<std::wstring> candidateVolume = TryGetVolumeGuidPathForPath(root);
        if (! candidateVolume.has_value() || candidateVolume.value() == referenceVolume.value())
        {
            continue;
        }
        const std::optional<std::filesystem::path> candidate = tryCandidateRoot(root);
        if (candidate.has_value())
        {
            return candidate;
        }
    }

    if (allowOwnedVirtualDisk)
    {
        // A one-volume hosted runner already has a harness-owned VHDX for FileOps.
        // Reuse only that exact validated mount; never adopt or mount another volume.
        using namespace RedSalamander::TestSupport;
        const std::wstring handedMount  = GetEnvironmentString(kOwnedVirtualDiskMountEnvironmentVariable);
        const std::wstring handedVolume = GetEnvironmentString(kOwnedVirtualDiskVolumeEnvironmentVariable);
        const std::wstring testRoot     = GetEnvironmentString(L"REDSALAMANDER_TEST_ROOT");
        const std::wstring runId        = GetEnvironmentString(L"REDSALAMANDER_TEST_RUN_ID");
        const auto expectedMount =
            testRoot.empty() || ! IsSafeTestRunId(runId) ? std::filesystem::path{} : GetOwnedVirtualDiskMountPath(testRoot, runId, L"f7-admission-baseline");
        const std::filesystem::path mountPath(handedMount);
        detail = EvaluateOwnedVirtualDiskMount(ObserveOwnedVirtualDiskMount(mountPath, handedVolume), handedVolume, mountPath, expectedMount);
        if (! detail.empty())
            return std::nullopt;
        const auto volume = TryGetVolumeGuidPathForPath(mountPath);
        if (! volume.has_value() || _wcsicmp(volume.value().c_str(), referenceVolume.value().c_str()) == 0)
        {
            detail = L"test-owned-volume-not-distinct-from-primary";
            return std::nullopt;
        }
        const auto candidate = mountPath / std::format(L"independent-volume-{}-{}", GetCurrentProcessId(), GetTickCount64());
        if (! SelfTest::EnsureDirectory(candidate))
        {
            detail = L"test-owned-alternate-case-root-unavailable";
            return std::nullopt;
        }
        AppendLog(std::format(L"Independent-volume fixture uses test-owned VHDX: mount={} volume={}", mountPath.native(), volume.value()));
        return candidate;
    }

    if (detail.empty())
    {
        detail = L"no-alternate-writable-fixed-volume";
    }
    return std::nullopt;
}

[[nodiscard]] bool TryDenyListDirectoryToEveryone(const std::filesystem::path& path) noexcept
{
    std::array<std::byte, SECURITY_MAX_SID_SIZE> sidBuffer{};
    DWORD sidSize = static_cast<DWORD>(sidBuffer.size());
    if (! CreateWellKnownSid(WinWorldSid, nullptr, sidBuffer.data(), &sidSize))
    {
        return false;
    }

    PACL existingDacl                       = nullptr;
    PSECURITY_DESCRIPTOR securityDescriptor = nullptr;
    const DWORD getSecurityError            = GetNamedSecurityInfoW(
        const_cast<wchar_t*>(path.c_str()), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, &existingDacl, nullptr, &securityDescriptor);
    if (getSecurityError != ERROR_SUCCESS || ! securityDescriptor)
    {
        return false;
    }

    wil::unique_hlocal_ptr<void> ownedSecurityDescriptor(securityDescriptor);

    EXPLICIT_ACCESSW denyEntry{};
    denyEntry.grfAccessPermissions = FILE_LIST_DIRECTORY;
    denyEntry.grfAccessMode        = DENY_ACCESS;
    denyEntry.grfInheritance       = NO_INHERITANCE;
    denyEntry.Trustee.TrusteeForm  = TRUSTEE_IS_SID;
    denyEntry.Trustee.TrusteeType  = TRUSTEE_IS_WELL_KNOWN_GROUP;
    denyEntry.Trustee.ptstrName    = reinterpret_cast<wchar_t*>(sidBuffer.data());

    PACL newDacl                = nullptr;
    const DWORD setEntriesError = SetEntriesInAclW(1, &denyEntry, existingDacl, &newDacl);
    if (setEntriesError != ERROR_SUCCESS || ! newDacl)
    {
        return false;
    }

    wil::unique_hlocal_ptr<ACL> ownedNewDacl(newDacl);

    const DWORD setSecurityError =
        SetNamedSecurityInfoW(const_cast<wchar_t*>(path.c_str()), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, ownedNewDacl.get(), nullptr);
    return setSecurityError == ERROR_SUCCESS;
}

std::optional<std::uint64_t> StartFileOperationAndGetId(
    FolderWindow::FileOperationState* fileOps,
    FileSystemOperation operation,
    FolderWindow::Pane sourcePane,
    std::optional<FolderWindow::Pane> destinationPane,
    const wil::com_ptr<IFileSystem>& fileSystem,
    std::vector<std::filesystem::path> sourcePaths,
    std::filesystem::path destinationFolder,
    FileSystemFlags flags,
    bool waitForOthers,
    uint64_t initialSpeedLimitBytesPerSecond                           = 0,
    FolderWindow::FileOperationState::ExecutionMode executionMode      = FolderWindow::FileOperationState::ExecutionMode::PerItem,
    bool requireConfirmation                                           = false,
    wil::com_ptr<IFileSystem> destinationFileSystem                    = nullptr,
    std::vector<FolderWindow::ResolvedFileOperationItem> resolvedItems = {},
    std::optional<uint32_t> moveClipboardSequence                      = std::nullopt,
    std::function<HRESULT()> preWorkerReleaseBarrier                   = {},
    std::function<HRESULT()> preConsumptionDecisionGate                = {},
    HRESULT* admissionStatus                                           = nullptr) noexcept
{
    if (admissionStatus)
        *admissionStatus = E_POINTER;
    if (! fileOps)
    {
        return std::nullopt;
    }

    uint64_t startedTaskId = 0;
    const HRESULT hrStart  = fileOps->AdmitOperation(operation,
                                                     sourcePane,
                                                     destinationPane,
                                                     fileSystem,
                                                     std::move(sourcePaths),
                                                     std::move(destinationFolder),
                                                     flags,
                                                     waitForOthers,
                                                     initialSpeedLimitBytesPerSecond,
                                                     executionMode,
                                                     requireConfirmation,
                                                     std::move(destinationFileSystem),
                                                     &startedTaskId,
                                                     std::move(resolvedItems),
                                                     {},
                                                     {},
                                                     {},
                                                     std::nullopt,
                                                     FileOperations::DeleteOrigin::PaneCommand,
                                                     std::nullopt,
                                                     moveClipboardSequence,
                                                     std::move(preWorkerReleaseBarrier),
                                                     std::move(preConsumptionDecisionGate));
    if (admissionStatus)
        *admissionStatus = hrStart;
    if (FAILED(hrStart) || startedTaskId == 0)
    {
        return std::nullopt;
    }
    return startedTaskId;
}

struct DummyReentrantWatchCallback final : public IFileSystemDirectoryWatchCallback
{
    IFileSystemDirectoryWatch* watch = nullptr;
    std::wstring watchedPath;
    std::atomic<uint64_t> callbackCount{0};
    std::atomic<uint64_t> changeCount{0};
    std::atomic<bool> unwatchAttempted{false};
    std::atomic<HRESULT> unwatchHr{S_OK};
    std::atomic<uint64_t> renamedOldCount{0};
    std::atomic<uint64_t> renamedNewCount{0};

    HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification, void* /*cookie*/) noexcept override
    {
        callbackCount.fetch_add(1u, std::memory_order_relaxed);
        if (notification != nullptr && notification->changes != nullptr)
        {
            changeCount.fetch_add(notification->changeCount, std::memory_order_relaxed);
            for (unsigned long index = 0; index < notification->changeCount; ++index)
            {
                if (notification->changes[index].action == FILESYSTEM_DIR_CHANGE_RENAMED_OLD_NAME)
                {
                    renamedOldCount.fetch_add(1u, std::memory_order_relaxed);
                }
                else if (notification->changes[index].action == FILESYSTEM_DIR_CHANGE_RENAMED_NEW_NAME)
                {
                    renamedNewCount.fetch_add(1u, std::memory_order_relaxed);
                }
            }
        }

        bool expected = false;
        if (watch != nullptr && unwatchAttempted.compare_exchange_strong(expected, true, std::memory_order_acq_rel))
        {
            unwatchHr.store(watch->UnwatchDirectory(watchedPath.c_str()), std::memory_order_release);
        }

        return S_OK;
    }
};

// Records every discovery report one governed provider call makes, so a fixture can judge the
// stream itself rather than the host's view of it: the host clamps a regressing cumulative value
// to a zero delta and counts late discovery without rejecting either, so a provider that reports
// out of order, or never closes, passes any test that only checks return codes.
struct DirectApiDiscoveryRecorder final : IFileSystemOperationControl
{
    struct Event final
    {
        uint64_t bytes       = 0;
        uint64_t files       = 0;
        uint64_t directories = 0;
        bool closed          = false;
    };

    std::mutex mutex;
    std::vector<Event> events;

    HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void* /*cookie*/) noexcept override
    {
        if (abort != nullptr)
        {
            *abort = FALSE;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void* /*cookie*/) noexcept override
    {
        if (mode != nullptr)
        {
            *mode = FILESYSTEM_DISCOVERY_AHEAD;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress* progress, void* /*cookie*/) noexcept override
    {
        if (progress == nullptr || progress->sizeBytes != sizeof(FileSystemDiscoveryProgress))
        {
            return E_INVALIDARG;
        }
        std::scoped_lock lock(mutex);
        events.push_back(Event{progress->discoveredBytes, progress->discoveredFiles, progress->discoveredDirectories, progress->traversalClosed != FALSE});
        return S_OK;
    }
};

struct ExpectedDiscoveryTotals final
{
    uint64_t bytes = 0;
    uint64_t files = 0;
    // Left unset when the provider's directory accounting is not part of the claim under test.
    std::optional<uint64_t> directories;
};

// Returns a failure description, or an empty string when the call's scope is well formed: the
// cumulative stream never decreases, exactly one report closes it, that closure is the last
// event, and it closed on the exact expected totals.
[[nodiscard]] std::wstring DescribeDiscoveryScopeDefect(std::wstring_view callName,
                                                        const DirectApiDiscoveryRecorder& recorder,
                                                        const ExpectedDiscoveryTotals& expected) noexcept
{
    if (recorder.events.empty())
    {
        return std::format(L"{} reported no discovery at all.", callName);
    }

    size_t closedCount      = 0;
    size_t firstClosedIndex = recorder.events.size();
    uint64_t previousBytes  = 0;
    uint64_t previousFiles  = 0;
    uint64_t previousDirs   = 0;
    for (size_t index = 0; index < recorder.events.size(); ++index)
    {
        const auto& event = recorder.events[index];
        if (event.bytes < previousBytes || event.files < previousFiles || event.directories < previousDirs)
        {
            return std::format(L"{} cumulative totals went backwards at report {} ({}/{}/{} after {}/{}/{}).",
                               callName,
                               index,
                               event.bytes,
                               event.files,
                               event.directories,
                               previousBytes,
                               previousFiles,
                               previousDirs);
        }
        previousBytes = event.bytes;
        previousFiles = event.files;
        previousDirs  = event.directories;
        if (event.closed)
        {
            ++closedCount;
            firstClosedIndex = (std::min)(firstClosedIndex, index);
        }
    }

    if (closedCount != 1u)
    {
        return std::format(L"{} closed its governed scope {} times; a call owns exactly one closure.", callName, closedCount);
    }
    if (firstClosedIndex + 1u != recorder.events.size())
    {
        return std::format(L"{} kept discovering after it closed: closure at report {} of {}.", callName, firstClosedIndex, recorder.events.size());
    }

    const auto& final = recorder.events.back();
    if (final.bytes != expected.bytes || final.files != expected.files ||
        (expected.directories.has_value() && final.directories != expected.directories.value()))
    {
        return std::format(L"{} closed on wrong totals (bytes={}/{} files={}/{} directories={}/{}).",
                           callName,
                           final.bytes,
                           expected.bytes,
                           final.files,
                           expected.files,
                           final.directories,
                           expected.directories.has_value() ? std::to_wstring(expected.directories.value()) : std::wstring(L"any"));
    }
    return {};
}

// Drives one provider call under a fresh recorder and returns its defect description, or an empty
// string. A failed call is a defect in its own right: the fixture never learns what a call that
// did not run would have reported.
template <typename Invoke>
[[nodiscard]] std::wstring RunDirectApiDiscoveryCall(std::wstring_view callName, const ExpectedDiscoveryTotals& expected, Invoke&& invoke) noexcept
{
    DirectApiDiscoveryRecorder recorder{};
    FileSystemOptions options{};
    options.sizeBytes              = sizeof(FileSystemOptions);
    options.operationControl       = &recorder;
    options.operationControlCookie = nullptr;
    const HRESULT hr               = invoke(options);
    if (FAILED(hr))
    {
        return std::format(L"{} failed (hr=0x{:08X}).", callName, static_cast<unsigned long>(hr));
    }
    return DescribeDiscoveryScopeDefect(callName, recorder, expected);
}

// A provider instance over a fake backend the plugin starts through its own self-test export. The
// fake stops and the module unloads with this object, so a step with several exits cannot leak
// either.
struct FakeBackendProvider final
{
    FakeBackendProvider() noexcept                             = default;
    FakeBackendProvider(const FakeBackendProvider&)            = delete;
    FakeBackendProvider& operator=(const FakeBackendProvider&) = delete;
    FakeBackendProvider(FakeBackendProvider&&)                 = delete;
    FakeBackendProvider& operator=(FakeBackendProvider&&)      = delete;
    ~FakeBackendProvider()
    {
        Stop();
    }

    wil::unique_hmodule module;
    void* endpoint    = nullptr;
    unsigned int port = 0u;
    std::string stopExport;
    std::string logExport;
    wil::com_ptr<IFileSystem> fileSystem;
    wil::com_ptr<IFileSystemIO> io;

    // The fake's own request log, for a failure message that says what the provider actually
    // asked of the backend. Empty when the fake keeps none.
    [[nodiscard]] std::wstring RequestLog() const noexcept
    {
        using RequestLogFn = HRESULT(__stdcall*)(void*, wchar_t*, unsigned int) noexcept;
        if (endpoint == nullptr || ! module || logExport.empty())
        {
            return {};
        }
        const FARPROC logAddress = GetProcAddress(module.get(), logExport.c_str());
        if (logAddress == nullptr)
        {
            return {};
        }
        wchar_t buffer[4096]{};
#pragma warning(push)
#pragma warning(disable : 4191) // The named self-test export fixes the typed ABI.
        const HRESULT hr = reinterpret_cast<RequestLogFn>(logAddress)(endpoint, buffer, static_cast<unsigned int>(std::size(buffer)));
#pragma warning(pop)
        return SUCCEEDED(hr) ? std::wstring(buffer) : std::wstring{};
    }

    void Stop() noexcept
    {
        io.reset();
        fileSystem.reset();
        if (endpoint != nullptr && module && ! stopExport.empty())
        {
            using StopFakeFn = void(__stdcall*)(void*) noexcept;
            if (const FARPROC stopAddress = GetProcAddress(module.get(), stopExport.c_str()))
            {
#pragma warning(push)
#pragma warning(disable : 4191) // The named self-test export fixes the typed ABI.
                reinterpret_cast<StopFakeFn>(stopAddress)(endpoint);
#pragma warning(pop)
            }
        }
        endpoint = nullptr;
        port     = 0u;
    }
};

// Starts the plugin's fake backend, creates a provider instance over it, and applies the
// configuration the caller builds from the fake's port. On failure `detail` says which step.
template <typename BuildConfiguration>
[[nodiscard]] HRESULT StartFakeBackendProvider(std::wstring_view pluginId,
                                               std::string_view startExport,
                                               std::string_view stopExport,
                                               BuildConfiguration&& buildConfiguration,
                                               FakeBackendProvider& provider,
                                               std::wstring& detail,
                                               std::string_view logExport = {}) noexcept
{
    using StartFakeFn = HRESULT(__stdcall*)(unsigned int*, void**) noexcept;
    using CreateFn    = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

    provider.Stop();
    provider.stopExport.assign(stopExport);
    provider.logExport.assign(logExport);
    FARPROC startAddress = nullptr;
    HRESULT hr           = SelfTest::LoadPluginSelfTestExport(pluginId, startExport, provider.module, startAddress);
    if (FAILED(hr) || startAddress == nullptr)
    {
        detail = std::format(L"the fake backend export is unavailable (hr=0x{:08X})", static_cast<unsigned long>(FAILED(hr) ? hr : E_FAIL));
        return FAILED(hr) ? hr : E_FAIL;
    }
    const FARPROC createAddress = GetProcAddress(provider.module.get(), "RedSalamanderCreate");
    if (createAddress == nullptr)
    {
        detail = L"the plugin factory export is unavailable";
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }
#pragma warning(push)
#pragma warning(disable : 4191) // The named self-test exports fix the typed ABI.
    hr = reinterpret_cast<StartFakeFn>(startAddress)(&provider.port, &provider.endpoint);
#pragma warning(pop)
    if (FAILED(hr) || provider.endpoint == nullptr || provider.port == 0u)
    {
        detail = std::format(L"the fake backend did not start (hr=0x{:08X})", static_cast<unsigned long>(hr));
        return FAILED(hr) ? hr : E_FAIL;
    }
    FactoryOptions factoryOptions{};
    factoryOptions.debugLevel = DEBUG_LEVEL_NONE;
#pragma warning(push)
#pragma warning(disable : 4191)
    hr = reinterpret_cast<CreateFn>(createAddress)(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), pluginId.data(), provider.fileSystem.put_void());
#pragma warning(pop)
    if (FAILED(hr) || ! provider.fileSystem)
    {
        detail = std::format(L"the provider instance was not created (hr=0x{:08X})", static_cast<unsigned long>(hr));
        return FAILED(hr) ? hr : E_FAIL;
    }
    wil::com_ptr<IInformations> information;
    const std::string configuration = buildConfiguration(provider.port);
    if (FAILED(provider.fileSystem->QueryInterface(__uuidof(IInformations), information.put_void())) || ! information ||
        FAILED(information->SetConfiguration(configuration.c_str())))
    {
        detail = L"the provider did not accept the fixture configuration";
        return E_FAIL;
    }
    if (FAILED(provider.fileSystem->QueryInterface(__uuidof(IFileSystemIO), provider.io.put_void())) || ! provider.io)
    {
        detail = L"the provider lacks IFileSystemIO";
        return E_NOINTERFACE;
    }
    return S_OK;
}

// The selection every provider discovery fixture drives: three plain leaves of distinct known
// sizes and one flat directory of four equal children, seeded through the provider's own writer
// so the fixture only depends on the routes under test. Sizes are odd on purpose so a total that
// merely looks plausible cannot match by accident.
struct DiscoveryFixtureSelection final
{
    static constexpr uint64_t kLeafBytesA     = 1024u + 7u;
    static constexpr uint64_t kLeafBytesB     = 2048u + 11u;
    static constexpr uint64_t kLeafBytesC     = 3072u + 13u;
    static constexpr uint64_t kTreeChildBytes = 512u + 3u;
    static constexpr uint64_t kTreeChildCount = 4u;
    static constexpr uint64_t kLeafBytes      = kLeafBytesA + kLeafBytesB + kLeafBytesC;
    static constexpr uint64_t kTreeBytes      = kTreeChildCount * kTreeChildBytes;

    std::wstring root;
    std::wstring source;
    std::wstring copyDestination;
    std::wstring bulkCopyDestination;
    std::wstring moveDestination;
    std::wstring leafA;
    std::wstring leafB;
    std::wstring leafC;
    std::wstring tree;

    [[nodiscard]] std::wstring Under(std::wstring_view folder, std::wstring_view leaf) const
    {
        std::wstring joined(folder);
        if (joined.empty() || joined.back() != L'/')
        {
            joined.push_back(L'/');
        }
        joined.append(leaf);
        return joined;
    }
};

// Seeds (or re-seeds) the selection under `root`, which must already be a valid location on the
// provider. Existing fixture content is removed first so a step can reuse the same root between
// calls that consume it. `createTreeMarker` is false for a provider whose directories are only
// prefixes and whose transfer of an explicit marker object is a separate, known defect.
[[nodiscard]] bool SeedDiscoveryFixtureSelection(IFileSystem* fileSystem,
                                                 std::wstring_view root,
                                                 DiscoveryFixtureSelection& selection,
                                                 bool createTreeMarker = true) noexcept
{
    if (fileSystem == nullptr || root.empty())
    {
        return false;
    }
    selection.root                = std::wstring(root);
    selection.source              = selection.Under(root, L"src");
    selection.copyDestination     = selection.Under(root, L"copy-dst");
    selection.bulkCopyDestination = selection.Under(root, L"copy-bulk-dst");
    selection.moveDestination     = selection.Under(root, L"move-dst");
    selection.leafA               = selection.Under(selection.source, L"leaf-a.bin");
    selection.leafB               = selection.Under(selection.source, L"leaf-b.bin");
    selection.leafC               = selection.Under(selection.source, L"leaf-c.bin");
    selection.tree                = selection.Under(selection.source, L"tree");

    if (! EnsureDummyFolderExists(fileSystem, selection.root))
    {
        return false;
    }
    constexpr FileSystemFlags kRemoveFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    for (const std::wstring* folder : {&selection.source, &selection.copyDestination, &selection.bulkCopyDestination, &selection.moveDestination})
    {
        // A missing folder is the normal first-run state; anything else is reported by the seed below.
        static_cast<void>(fileSystem->DeleteItem(folder->c_str(), kRemoveFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(fileSystem, *folder))
        {
            return false;
        }
    }
    if (createTreeMarker && ! EnsureDummyFolderExists(fileSystem, selection.tree))
    {
        return false;
    }

    const auto write = [&](const std::wstring& path, uint64_t bytes) noexcept
    {
        const std::string contents(static_cast<size_t>(bytes), 'd');
        return DummyWriteTextFile(fileSystem, path, contents);
    };
    if (! write(selection.leafA, DiscoveryFixtureSelection::kLeafBytesA) || ! write(selection.leafB, DiscoveryFixtureSelection::kLeafBytesB) ||
        ! write(selection.leafC, DiscoveryFixtureSelection::kLeafBytesC))
    {
        return false;
    }
    for (uint64_t index = 0u; index < DiscoveryFixtureSelection::kTreeChildCount; ++index)
    {
        if (! write(selection.Under(selection.tree, std::format(L"child-{:02}.bin", index)), DiscoveryFixtureSelection::kTreeChildBytes))
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] std::wstring JoinDiscoveryDefects(const std::vector<std::wstring>& defects, std::wstring_view fakeRequestLog = {})
{
    std::wstring joined;
    for (const std::wstring& defect : defects)
    {
        if (! joined.empty())
        {
            joined.append(L" | ");
        }
        joined.append(defect);
    }
    if (! fakeRequestLog.empty())
    {
        joined.append(L" fixture:");
        joined.append(fakeRequestLog);
    }
    return joined;
}

// What the host must see after a same-endpoint leaf Copy on a provider's direct route: the
// provider closed the scope on the leaf's exact size before the transfer, so no mutation completed
// while discovery was still open and nothing grew after closure. Empty when the task is right.
[[nodiscard]] std::wstring DescribeHostLeafDiscoveryDefect(std::wstring_view label, const CompletedTaskInfo& info, uint64_t expectedBytes) noexcept
{
    if (FAILED(info.hr))
    {
        return std::format(L"{}: the host leaf task failed (hr=0x{:08X}).", label, static_cast<unsigned long>(info.hr));
    }
    if (info.bridgeFileAdmissionCount != 0ull)
    {
        // The bridge owns its own exact leaf record, so a task that took it proves nothing about the provider.
        return std::format(
            L"{}: the host leaf task took the bridge instead of the provider's direct route (bridge admissions={}).", label, info.bridgeFileAdmissionCount);
    }
    if (! info.discoveryClosed || info.discoveredTotalBytes != expectedBytes || info.discoveredFiles != 1ul ||
        info.discoveryMutationsCompletedWhileOpen != 0ull || info.discoveryGrowthAfterCloseCount != 0ull)
    {
        return std::format(L"{}: the host leaf task did not see a closed exact record before the transfer (closed={} bytes={}/{} files={} "
                           L"mutationsWhileOpen={} growthAfterClose={}).",
                           label,
                           info.discoveryClosed ? 1 : 0,
                           info.discoveredTotalBytes,
                           expectedBytes,
                           info.discoveredFiles,
                           info.discoveryMutationsCompletedWhileOpen,
                           info.discoveryGrowthAfterCloseCount);
    }
    return {};
}

} // namespace

std::vector<std::wstring> FileOperationsSelfTest::BuildRunFilters(const SelfTest::SelfTestOptions& options)
{
    return BuildRunFiltersImpl(options.caseFilter);
}

std::vector<std::wstring> FileOperationsSelfTest::BuildExpectedCaseNames(const SelfTest::SelfTestOptions& options)
{
    const RunSelection selection = ResolveRunSelection(options.caseFilter);
    if (! selection.recognized)
    {
        return {};
    }

    std::vector<std::wstring> names;
    names.reserve(selection.reportedPhases.size());
    for (const SelfTestState::Step step : selection.reportedPhases)
    {
        names.emplace_back(StepToString(step));
    }
    return names;
}

void FileOperationsSelfTest::Start(HWND mainWindow, const SelfTest::SelfTestOptions& options) noexcept
{
    SelfTestState& state = GetState();
    // Do not let a later aggregate family recreate the scratch retained by D6.
    // Cleanup/Stop still cancel and drain normally; review requires a fresh process/run root.
    if (state.retainTempRootOnFailure)
    {
        AppendLog(L"Start: refused reuse after retained storage-proof failure; scratch requires review.");
        return;
    }
    if (state.running.exchange(true, std::memory_order_acq_rel))
    {
        AppendLog(L"Start: skipped because a run is already marked running");
        return;
    }

    AppendLog(L"Start: reset begin");
    state.options   = options;
    state.runFilter = options.caseFilter;
    state.done.store(false, std::memory_order_release);
    state.failed.store(false, std::memory_order_release);
    state.failureMessage.clear();
    state.mainWindow = mainWindow;
    state.tempRoot.clear();
    // Keep shared plugin COM instances and their original configuration snapshots alive for the
    // whole aggregate run. Cleanup already defers plugin release to process shutdown, and releasing
    // them between families can hang the UI during family-to-family transitions.
    state.connOverrideProfileName.clear();
    // Keep the seeded FileSystemDummy selection across aggregate families. Setup reapplies the same
    // seed config and destination folders without rediscovering the source path on every transition.
    state.localConfigDirty = false;
    state.dummyConfigDirty = false;
    state.config7zDirty    = false;
    AppendLog(L"Start: clearing remote case state");
    ResetRemoteOneDrivePersonalState(state);
    state.folderWindow = nullptr;
    state.fileOps      = nullptr;
    state.taskA.reset();
    state.taskB.reset();
    state.batchRenameExactPromptGeneration = 0u;
    state.batchRenameExactCurrentTaskId    = 0u;
    state.batchRenameExactAutoAcceptBefore.reset();
    state.batchRenameExactPaneFileSystem.reset();
    state.taskC.reset();
    state.queuePausedTask.reset();
    state.popupOriginalRect      = {};
    state.popupOriginalRectValid = false;
    state.directoryWatch.reset();
    state.directoryWatchCallback.reset();
    state.watchDir.clear();
    state.watchCounter = 0;
    state.lockedFileHandle.reset();
    state.lockedDestinationHandle.reset();
    state.copyKnobIndex                       = 0;
    state.copyKnobRetryCount                  = 0;
    state.deleteKnobIndex                     = 0;
    state.copySpeedLimitCleared               = false;
    state.copyPromptValidated                 = false;
    state.copyKnobObservedPerCallShare        = false;
    state.copyKnobObservedActiveCalls         = 0;
    state.copyKnobObservedDesiredSpeedLimit   = 0;
    state.copyKnobObservedAppliedSpeedLimit   = 0;
    state.copyKnobObservedEffectiveSpeedLimit = 0;
    state.phase10ClipboardPublishedNonRunning.store(false, std::memory_order_release);
    state.phase10ClipboardReadinessBarrierCalled.store(false, std::memory_order_release);
    state.phase10ClipboardSlowGateRelease.store(false, std::memory_order_release);
    state.phase10ClipboardReturnedBeforeReadiness = false;
    state.autoDismissSuccessOriginal              = false;
    state.fileOperationsBackedUp                  = false;
    state.fileOperationsOriginal.reset();
    state.appThemeBackedUp = false;
    state.appThemeOriginal.reset();
    state.copyTaskStartTick                 = 0;
    state.localBandwidthRunStartTick        = 0;
    state.localBandwidthCancelHandlerTick   = 0;
    state.localBandwidthCancelStartTick     = 0;
    state.localBandwidthCancelReturnTick    = 0;
    state.localBandwidthDurationUs          = 0;
    state.localBandwidthDurationLeadUs      = 0;
    state.localBandwidthCancelLatencyUs     = 0;
    state.localBandwidthMaxWindowBytes      = 0;
    state.localBandwidthMaxSampleDeltaBytes = 0;
    state.localBandwidthSamples.clear();
    state.bandwidthThrottleWorkerModeEnvBackedUp    = false;
    state.bandwidthThrottleWorkerModeEnvHadOriginal = false;
    state.bandwidthThrottleWorkerModeEnvOriginal.clear();
    state.parallelBandwidthRunStartTick                   = 0;
    state.parallelBandwidthBaselineUs                     = 0;
    state.parallelBandwidthCandidateUs                    = 0;
    state.parallelBandwidthBaselineMaxSkewBytes           = 0;
    state.parallelBandwidthCandidateMaxSkewBytes          = 0;
    state.parallelBandwidthBaselineMaxCallbackDeltaBytes  = 0;
    state.parallelBandwidthCandidateMaxCallbackDeltaBytes = 0;
    state.parallelBandwidthBaselineMaxActive              = 0;
    state.parallelBandwidthCandidateMaxActive             = 0;
    state.parallelBandwidthBaselineSamples                = 0;
    state.parallelBandwidthCandidateSamples               = 0;
    state.defaultSpeedLimitRunStartTick                   = 0;
    state.defaultSpeedLimitBaselineUs                     = 0;
    state.defaultSpeedLimitCandidateUs                    = 0;
    state.defaultSpeedLimitDummyConfigSnapshot.clear();
    state.completedTasks.clear();
    state.phase14InfoTask.reset();
    state.phase14ShutdownDone.store(false, std::memory_order_release);
    ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
    SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
    ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
    SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
    ReleaseFileOpsPermanentDeleteBeforeBindPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeBindPauseForSelfTest(false);
    ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
    ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
    SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
    SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(0u);
    static_cast<void>(TakeFileOpsNativeMoveCreateDirectoryRaceAttemptsForSelfTest());
    SetFileOpsManagedCleanupKnownNonCommitForSelfTest(HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), 0u);
    static_cast<void>(TakeFileOpsManagedCleanupKnownNonCommitAttemptsForSelfTest());
    SetFileOpsPermanentDeleteKnownNonCommitForSelfTest(HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), 0u);
    static_cast<void>(TakeFileOpsPermanentDeleteKnownNonCommitAttemptsForSelfTest());
    SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(0u);
    static_cast<void>(TakeFileOpsManagedCleanupUnknownOutcomeAttemptsForSelfTest());
    SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
    static_cast<void>(TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest());

    state.activePhaseOrder.clear();
    state.reportedPhaseOrder.clear();
    state.phaseResults.clear();
    state.phaseInProgress = false;
    state.phaseStartTick  = 0;
    state.phaseFailed     = false;
    state.phaseName.clear();
    state.phaseFailureMessage.clear();
    AppendLog(L"Start: resolve selection");

    const RunSelection selection = ResolveRunSelection(options.caseFilter);
    if (! selection.recognized)
    {
        state.running.store(false, std::memory_order_release);
        state.failed.store(true, std::memory_order_release);
        state.done.store(true, std::memory_order_release);
        state.failureMessage = std::format(L"FileOpsSelfTest unknown case/family filter '{}'.", options.caseFilter);
        AppendLog(std::format(L"FAIL: {}", state.failureMessage));
        Debug::Error(L"FileOpsSelfTest FAILED: {}", state.failureMessage);
        return;
    }

    AppendLog(L"Start: initialize phase order");
    state.activePhaseOrder   = selection.activePhases;
    state.reportedPhaseOrder = selection.reportedPhases;
    state.step = EqualsIgnoreCase(options.caseFilter, L"FileOps_VisualGallery") ? SelfTestState::Step::FileOps_VisualGallery : SelfTestState::Step::Setup;
    if (EqualsIgnoreCase(options.caseFilter, L"FileOps_VisualGalleryInteraction"))
        state.step = SelfTestState::Step::FileOps_VisualGalleryInteraction;
    if (EqualsIgnoreCase(options.caseFilter, L"FileOps_CommonPresentationPerf"))
        state.step = SelfTestState::Step::FileOps_CommonPresentationPerf;
    if (EqualsIgnoreCase(options.caseFilter, L"FileOps_IntegratedPresentationPerf"))
        state.step = SelfTestState::Step::FileOps_IntegratedPresentationPerf;
    state.runStartTick        = GetTickCount64();
    state.stepStartTick       = static_cast<ULONGLONG>(state.runStartTick);
    state.markerTick          = 0;
    state.baselineThreadCount = 0;
    BeginPhase(state, state.step);
    AppendLog(L"Start: setup ready");
    AppendLog(L"Start");
    Debug::Info(L"FileOpsSelfTest: started");
}

namespace
{
using TickResult = std::optional<bool>;

[[nodiscard]] HRESULT PopulateBatchRenameExactResultProviderFacts(IFileSystem* fileSystem, BatchRenameExecutionOp& operation) noexcept;

[[nodiscard]] bool RunFileOpsCommittedRetainedDeleteChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    auto* folder      = TryGetFolderWindow(state.mainWindow);
    auto* operations  = TryGetFileOps(folder);
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps committed-retained Delete: {}", reason));
        return false;
    };
    if (! folder || ! operations || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner or initial quiet point unavailable");

    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground            = GetForegroundWindow();
    const auto sourceFileSystem      = folder->GetFileSystem(FolderWindow::Pane::Left);
    const auto destinationFileSystem = folder->GetFileSystem(FolderWindow::Pane::Right);
    wil::com_ptr<IFileSystemIO> sourceIo;
    if (! sourceFileSystem || ! destinationFileSystem || ! sourceFileSystem.try_query_to(sourceIo.put()) ||
        folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal || folder->GetFileSystemPluginId(FolderWindow::Pane::Right) != kPluginIdLocal)
        return reject(L"actual Local pane filesystems unavailable");

    constexpr std::wstring_view pathVariable  = L"REDSALAMANDER_FILEOPS_DELETE_COMMITTED_RETAINED_PATH";
    constexpr std::wstring_view firedVariable = L"REDSALAMANDER_FILEOPS_DELETE_COMMITTED_RETAINED_FIRED";
    if (! GetEnvVarTrimmed(pathVariable).empty() || ! GetEnvVarTrimmed(firedVariable).empty())
        return reject(L"one-shot injection already configured");
    const auto root                      = state.tempRoot / L"committed-retained-delete";
    const auto source                    = root / L"retained_r\u00E9vision_\u6771\u4EAC.bin";
    constexpr size_t payloadBytes        = 4099u;
    constexpr unsigned char payloadValue = 0xA7u;
    if (! RecreateEmptyDirectory(root) || ! WriteFilledTestFile(source, payloadBytes, payloadValue))
        return reject(L"owned distinct payload could not be staged");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    std::string originalBytes;
    if (! ReadFileTextFsIo(sourceIo, source, originalBytes) || originalBytes.size() != payloadBytes ||
        ! std::ranges::all_of(originalBytes, [](char value) noexcept { return static_cast<unsigned char>(value) == payloadValue; }))
        return reject(L"original exact bytes unavailable");
    Common::Crypto::Blake3Hasher originalHasher;
    originalHasher.Update(std::as_bytes(std::span<const char>(originalBytes.data(), originalBytes.size())));
    const auto originalDigest = originalHasher.Finalize();

    const bool autoDismiss    = operations->GetAutoDismissSuccess();
    const bool autoAccept     = HostGetAutoAcceptPrompts();
    const auto promptOverride = HostGetTestPromptResultOverride();
    const auto restore        = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(pathVariable.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(firedVariable.data(), nullptr));
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        operations->SetAutoDismissSuccess(autoDismiss);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
    });
    operations->SetAutoDismissSuccess(false);
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();

    uint64_t taskId = 0u;
    std::atomic<unsigned int> callbackCount{0u};
    std::atomic<HRESULT> callbackResult{E_PENDING};
    const uint64_t subscription = folder->AddFileOperationCompletedCallback([&](const FolderWindow::FileOperationCompletedEvent& event) noexcept
    {
        if (event.taskId == taskId)
        {
            callbackResult.store(event.hr, std::memory_order_release);
            callbackCount.fetch_add(1u, std::memory_order_acq_rel);
        }
    });
    const auto unsubscribe      = wil::scope_exit([&]() noexcept { folder->RemoveFileOperationCompletedCallback(subscription); });
    // Drain before unsubscribing or destroying any state borrowed by the callback.
    const auto drain = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(pathVariable.data(), nullptr));
        static_cast<void>(CancelOutstandingSelfTestOperations(operations));
        if (taskId != 0u)
        {
            operations->DismissCompletedTask(taskId);
            operations->DebugRemoveDiagnosticsForTask(taskId);
        }
        operations->DebugResetPopupForSelfTest();
    });
    if (! SetEnvironmentVariableW(pathVariable.data(), source.c_str()))
        return reject(L"one-shot injection could not be configured");
    const auto admitted = StartFileOperationAndGetId(
        operations, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, sourceFileSystem, {source}, {}, FILESYSTEM_FLAG_RECURSIVE, false);
    if (! admitted.has_value())
        return reject(L"normal PermanentDelete admission failed");
    taskId = admitted.value();
    std::optional<Task::ConflictPromptState> prompt;
    const auto consent = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        prompt = TryGetConflictPromptCopy(operations->FindTask(taskId));
        return prompt.has_value();
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"committed-retained Delete normal consent"});
    auto* task = operations->FindTask(taskId);
    if (! consent.conditionMet || ! task || prompt.value().bucket != Task::ConflictBucket::PermanentDeleteConfirmation ||
        prompt.value().sourcePath != source.native() || ! PromptHasAction(prompt.value(), Task::ConflictAction::PermanentDelete))
        return reject(L"normal identity-bound PermanentDelete consent unavailable");
    if (! task->SubmitConflictDecision(prompt.value().generation, Task::ConflictAction::PermanentDelete, false))
        return reject(L"normal PermanentDelete consent was rejected");
    const auto quiet = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return callbackCount.load(std::memory_order_acquire) == 1u && ! operations->FindTask(taskId) && ! operations->HasActiveOperations();
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"committed-retained Delete worker quiet point"});
    if (! quiet.conditionMet || callbackResult.load(std::memory_order_acquire) != E_UNEXPECTED || GetEnvVarTrimmed(firedVariable) != L"1")
        return reject(L"actual injected task did not complete once with E_UNEXPECTED");

    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto summary = std::ranges::find_if(summaries, [taskId](const auto& value) noexcept { return value.taskId == taskId; });
    if (summary == summaries.end() || summary->resultHr != E_UNEXPECTED || summary->errorCount == 0u || summary->retainedSourceCount != 1u ||
        summary->removedSourceCount != 0u || summary->unknownSourceCount != 0u || summary->indeterminateItemCount != 0u ||
        ! std::ranges::any_of(summary->issueDiagnostics, [taskId](const auto& issue) noexcept {
        return issue.taskId == taskId && issue.category == L"task.result" && issue.status == E_UNEXPECTED;
    }))
        return reject(L"real completed summary did not retain Failed/Retained truth and task.result diagnostic");
    FileOperationsPopupInternal::TaskSnapshot terminal{};
    if (! DebugGetFileOperationsPopupTaskSnapshot(operations->GetPopupHwndForSelfTest(), taskId, terminal) || ! terminal.finished ||
        terminal.statusKind != FileOperationsPopupInternal::TaskSnapshot::StatusKind::Failed)
        return reject(L"real completed task card did not project Failed");
    std::string retainedBytes;
    if (! ReadFileTextFsIo(sourceIo, source, retainedBytes) || retainedBytes != originalBytes)
        return reject(L"retained source changed exact original bytes");
    Common::Crypto::Blake3Hasher retainedHasher;
    retainedHasher.Update(std::as_bytes(std::span<const char>(retainedBytes.data(), retainedBytes.size())));
    if (retainedHasher.Finalize() != originalDigest)
        return reject(L"retained source changed canonical BLAKE3 digest");

    auto lifetime = std::make_shared<int>(1);
    wil::unique_hwnd issues(FileOperationsIssuesPane::Create(operations, folder, state.mainWindow, lifetime));
    // This private pane is not registered as the application pane. Expire its
    // owner token before teardown so it cannot persist test placement/view state.
    const auto closeIssues = wil::scope_exit([&]() noexcept
    {
        lifetime.reset();
        issues.reset();
    });
    FileOperationsIssuesPane::SelfTestSnapshot selected{};
    if (! issues || ! FileOperationsIssuesPane::SelfTestSelectTask(issues.get(), taskId) ||
        ! FileOperationsIssuesPane::SelfTestSetTechnicalExpanded(issues.get(), true))
        return reject(L"actual Issues projection could not select the task and expand technical columns");
    // Consent is an earlier S_OK diagnostic for this same task. Select the actual
    // result row explicitly rather than assuming generic task selection orders errors first.
    if (! FileOperationsIssuesPane::SelfTestSelectDiagnostic(issues.get(), taskId, L"task.result", E_UNEXPECTED))
        return reject(L"actual Issues projection omitted task.result/E_UNEXPECTED");
    ShowWindow(issues.get(), SW_SHOWNOACTIVATE);
    RedrawWindow(issues.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(issues.get(), selected, true) || selected.selectionCount != 1u ||
        selected.primarySelectedTaskId != taskId || selected.primarySelectedRowId == 0u)
        return reject(L"actual Issues projection could not select the completed task");
    const auto statusColumn   = std::ranges::find(selected.columnIds, L"hresult");
    const auto categoryColumn = std::ranges::find(selected.columnIds, L"category");
    if (statusColumn == selected.columnIds.end() || categoryColumn == selected.columnIds.end())
        return reject(L"Issues status/category columns unavailable");
    const size_t statusIndex   = static_cast<size_t>(statusColumn - selected.columnIds.begin());
    const size_t categoryIndex = static_cast<size_t>(categoryColumn - selected.columnIds.begin());
    if (statusIndex >= selected.selectedCells.size() || categoryIndex >= selected.selectedCells.size())
        return reject(L"selected actual Issues row omitted numeric status/category cells");
    if (selected.selectedCells[categoryIndex] != L"task.result")
        return reject(std::format(L"selected actual Issues category='{}', expected='task.result'", selected.selectedCells[categoryIndex]));
    if (selected.selectedCells[statusIndex] != L"0x8000FFFF")
        return reject(std::format(L"selected actual Issues HRESULT='{}', expected='0x8000FFFF'", selected.selectedCells[statusIndex]));
    const HWND actualForeground = GetForegroundWindow();
    if (actualForeground != foreground)
        return reject(std::format(L"Issues changed foreground: expected=0x{:X}, actual=0x{:X}",
                                  reinterpret_cast<uintptr_t>(foreground),
                                  reinterpret_cast<uintptr_t>(actualForeground)));
    if (FileOperationsIssuesPane::SelfTestSelectDiagnostic(issues.get(), 0u, L"task.result", E_UNEXPECTED) ||
        FileOperationsIssuesPane::SelfTestSelectDiagnostic(issues.get(), taskId, L"missing-test-category", E_UNEXPECTED) ||
        FileOperationsIssuesPane::SelfTestSelectDiagnostic(issues.get(), taskId, L"task.result", S_OK))
        return reject(L"Issues diagnostic selector accepted a wrong task/category/HRESULT");
    FileOperationsIssuesPane::SelfTestSnapshot afterRefusal{};
    if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(issues.get(), afterRefusal, true) || afterRefusal.selectionCount != 1u ||
        afterRefusal.primarySelectedRowId != selected.primarySelectedRowId || afterRefusal.selectedCells != selected.selectedCells ||
        GetForegroundWindow() != foreground)
        return reject(L"refused diagnostic selection changed the actual row or foreground");
    AppendLog(std::format(L"Committed-retained Delete task={}: result=0x8000FFFF errors={} retained=1 removed=0 unknown=0 indeterminate=0; "
                          L"task.result retained; Issues selected category='{}' status='{}'; bytes={} preserved BLAKE3={}",
                          taskId,
                          summary->errorCount,
                          selected.selectedCells[categoryIndex],
                          selected.selectedCells[statusIndex],
                          retainedBytes.size(),
                          Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(originalDigest))));
    return true;
}

[[nodiscard]] bool RunFileOpsRecursiveDeleteAfterRemovalCancelChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Recursive Delete after-removal cancel: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner or initial quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground       = GetForegroundWindow();
    const auto sourceFileSystem = folder->GetFileSystem(FolderWindow::Pane::Left);
    wil::com_ptr<IFileSystemIO> sourceIo;
    wil::com_ptr<IInformations> information;
    if (! sourceFileSystem || folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal || ! sourceFileSystem.try_query_to(sourceIo.put()) ||
        ! sourceFileSystem.try_query_to(information.put()))
        return reject(L"actual Local source pane interfaces unavailable");

    const auto root   = state.tempRoot / L"recursive-delete-after-removal-cancel";
    const auto source = root / L"source_r\u00E9vision_\u6771\u4EAC";
    bool cleanupQuiet = true;
    if (! RecreateEmptyDirectory(root))
        return reject(L"owned fixture root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept
    {
        if (cleanupQuiet)
            static_cast<void>(SelfTest::RemoveAll(root));
        else
            AppendLog(L"Recursive Delete cancel cleanup: quiet point unavailable; owned fixture retained.");
    });
    struct Seed final
    {
        std::filesystem::path relative;
        size_t bytes;
        unsigned char value;
    };
    const std::array<Seed, 5> seeds{{
        {L"Alpha.bin", 4099u, 0xA7u},
        {L"Bravo_r\u00E9vision.bin", 2053u, 0x3Cu},
        {L"Charlie_\u6771\u4EAC.bin", 1031u, 0xD2u},
        {std::filesystem::path(L"subtree") / L"leaf.bin", 521u, 0x6Eu},
        {std::filesystem::path(L"subtree") / L"deeper" / L"leaf.bin", 263u, 0x91u},
    }};
    for (const auto& seed : seeds)
    {
        if (! WriteFilledTestFile(source / seed.relative, seed.bytes, seed.value))
            return reject(L"distinct fixture payloads unavailable");
    }

    // EXE-local FileOps witness, deliberately limited to a small, owned ordinary-file
    // tree. Unlike a general filesystem walker it rejects every reparse point and any
    // entry outside this closed seed inventory; it never traverses a link or an unknown
    // tree. Provider IO and canonical BLAKE3 retain exact bytes as well as hashes.
    struct TreeEntry final
    {
        std::wstring relative;
        DWORD attributes = 0u;
        std::string bytes;
        Common::Crypto::Blake3Digest digest{};
        bool operator==(const TreeEntry&) const noexcept = default;
    };
    std::vector<std::wstring> expectedNames{L".", L"subtree", L"subtree\\deeper"};
    for (const auto& seed : seeds)
        expectedNames.push_back(seed.relative.native());
    std::ranges::sort(expectedNames);
    const auto capture = [&](std::vector<TreeEntry>& tree) noexcept
    {
        tree.clear();
        std::vector<std::filesystem::path> directories{source};
        size_t capturedBytes = 0u;
        for (size_t index = 0u; index < directories.size(); ++index)
        {
            const auto directory            = directories[index];
            const DWORD directoryAttributes = GetFileAttributesW(directory.c_str());
            if (directoryAttributes == INVALID_FILE_ATTRIBUTES || (directoryAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u ||
                (directoryAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
                return false;
            if (index == 0u)
                tree.push_back(TreeEntry{L".", directoryAttributes, {}, {}});
            std::error_code ec;
            std::filesystem::directory_iterator entry(directory, std::filesystem::directory_options::none, ec);
            if (ec)
                return false;
            for (; entry != std::filesystem::directory_iterator{}; entry.increment(ec))
            {
                if (ec || tree.size() >= expectedNames.size())
                    return false;
                const auto path = entry->path();
                TreeEntry observed;
                observed.relative = path.lexically_relative(source).native();
                if (! std::ranges::binary_search(expectedNames, observed.relative))
                    return false;
                observed.attributes = GetFileAttributesW(path.c_str());
                if (observed.attributes == INVALID_FILE_ATTRIBUTES || (observed.attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
                    return false;
                if ((observed.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
                {
                    if (directories.size() >= 3u)
                        return false;
                    directories.push_back(path);
                }
                else
                {
                    const auto seed = std::ranges::find_if(seeds, [&](const Seed& value) noexcept { return value.relative.native() == observed.relative; });
                    if (seed == seeds.end() || ! ReadFileTextFsIo(sourceIo, path, observed.bytes) || observed.bytes.size() != seed->bytes ||
                        capturedBytes > 64u * 1024u - observed.bytes.size())
                        return false;
                    capturedBytes += observed.bytes.size();
                    Common::Crypto::Blake3Hasher hasher;
                    hasher.Update(std::as_bytes(std::span<const char>(observed.bytes.data(), observed.bytes.size())));
                    observed.digest = hasher.Finalize();
                }
                tree.push_back(std::move(observed));
            }
            if (ec)
                return false;
        }
        std::ranges::sort(tree, {}, &TreeEntry::relative);
        return true;
    };
    std::vector<TreeEntry> original;
    if (! capture(original) || original.size() != expectedNames.size())
        return reject(L"original closed tree unavailable");
    for (size_t index = 0u; index < original.size(); ++index)
    {
        if (original[index].relative != expectedNames[index])
            return reject(L"original closed tree differs from the exact seed inventory");
    }
    for (const auto& seed : seeds)
    {
        const auto entry = std::ranges::find(original, seed.relative.native(), &TreeEntry::relative);
        if (entry == original.end() || entry->bytes.size() != seed.bytes ||
            ! std::ranges::all_of(entry->bytes, [value = seed.value](char byte) noexcept { return static_cast<unsigned char>(byte) == value; }))
            return reject(L"original bytes differ from the distinct sentinel payloads");
    }

    std::string originalConfiguration;
    if (! BackupPluginConfiguration(information.get(), originalConfiguration))
        return reject(L"actual pane provider configuration unavailable");
    bool configurationRestored      = false;
    const auto restoreConfiguration = wil::scope_exit([&]() noexcept
    {
        if (! configurationRestored && FAILED(information->SetConfiguration(originalConfiguration.c_str())))
            AppendLog(L"Recursive Delete cancel cleanup: provider configuration restoration FAILED.");
    });
    if (FAILED(information->SetConfiguration(R"json({"concurrencyMode":"manual","deleteMaxConcurrency":1})json")))
        return reject(L"serial Local configuration refused");
    const auto route = FileSystemRouteContract::Query(sourceFileSystem.get(), source.native(), FILESYSTEM_DELETE, kPluginIdLocal);
    if (route.state != FileSystemRouteContract::QueryState::Available || route.snapshot.deleteMaxConcurrency != 1u)
        return reject(L"actual pane provider did not establish serial Delete");

    const bool autoDismiss         = operations->GetAutoDismissSuccess();
    const bool autoAccept          = HostGetAutoAcceptPrompts();
    const auto promptOverride      = HostGetTestPromptResultOverride();
    const auto restorePresentation = wil::scope_exit([&]() noexcept
    {
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        operations->SetAutoDismissSuccess(autoDismiss);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
    });
    operations->SetAutoDismissSuccess(false);
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();

    try
    {
        constexpr std::wstring_view parentVariable  = L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_PARENT";
        constexpr std::wstring_view enteredVariable = L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_ENTERED_EVENT";
        constexpr std::wstring_view releaseVariable = L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_RELEASE_EVENT";
        constexpr std::wstring_view removedVariable = L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_REMOVED_PATH";
        if (! GetEnvVarTrimmed(parentVariable).empty() || ! GetEnvVarTrimmed(enteredVariable).empty() || ! GetEnvVarTrimmed(releaseVariable).empty() ||
            ! GetEnvVarTrimmed(removedVariable).empty())
            return reject(L"another removal gate is already configured");
        const auto eventStem   = std::format(L"Local\\RedSalamander.D1.{}.{}", GetCurrentProcessId(), GetTickCount64());
        const auto enteredName = eventStem + L".entered";
        const auto releaseName = eventStem + L".release";
        wil::unique_handle entered(CreateEventW(nullptr, TRUE, FALSE, enteredName.c_str()));
        const DWORD enteredError = GetLastError();
        if (! entered || enteredError == ERROR_ALREADY_EXISTS)
            return reject(L"case-owned entered event unavailable or collided");
        wil::unique_handle release(CreateEventW(nullptr, TRUE, FALSE, releaseName.c_str()));
        const DWORD releaseError = GetLastError();
        if (! release || releaseError == ERROR_ALREADY_EXISTS)
            return reject(L"case-owned release event unavailable or collided");
        const auto extendedSource = MakeExtendedPathForSelfTest(source);
        RedSalamander::TestSupport::ScopedEnvironmentVariable parentScope(parentVariable, extendedSource);
        RedSalamander::TestSupport::ScopedEnvironmentVariable enteredScope(enteredVariable, enteredName);
        RedSalamander::TestSupport::ScopedEnvironmentVariable releaseScope(releaseVariable, releaseName);
        RedSalamander::TestSupport::ScopedEnvironmentVariable removedScope(removedVariable, std::nullopt);
        struct Observation final
        {
            uint64_t taskId            = 0u;
            unsigned int callbackCount = 0u;
            HRESULT result             = E_PENDING;
        };
        // Completion callbacks run on the owner thread. Shared capture, never a borrowed
        // stack, remains safe even if failure cleanup cannot establish the quiet point.
        auto observation            = std::make_shared<Observation>();
        const uint64_t subscription = folder->AddFileOperationCompletedCallback([observation](const FolderWindow::FileOperationCompletedEvent& event) noexcept
        {
            if (event.taskId == observation->taskId)
            {
                observation->result = event.hr;
                ++observation->callbackCount;
            }
        });
        const auto drain            = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(SetEvent(release.get()));
            cleanupQuiet = CancelOutstandingSelfTestOperations(operations);
            if (cleanupQuiet && observation->taskId != 0u)
            {
                const auto retired = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                    return ! operations->FindTask(observation->taskId);
                }, {.timeout = SelfTest::Scale(std::chrono::seconds(30)), .operationName = L"D1 failure cleanup task retirement"});
                cleanupQuiet       = retired.conditionMet;
            }
            folder->RemoveFileOperationCompletedCallback(subscription);
            if (cleanupQuiet && observation->taskId != 0u)
            {
                operations->DismissCompletedTask(observation->taskId);
                operations->DebugRemoveDiagnosticsForTask(observation->taskId);
                state.completedTasks.erase(observation->taskId);
                operations->DebugResetPopupForSelfTest();
            }
        });
        const auto admitted         = StartFileOperationAndGetId(
            operations, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, sourceFileSystem, {source}, {}, FILESYSTEM_FLAG_RECURSIVE, false);
        if (! admitted.has_value())
            return reject(L"normal recursive PermanentDelete admission failed");
        observation->taskId = admitted.value();
        cleanupQuiet        = false;
        std::optional<Task::ConflictPromptState> prompt;
        const auto consent = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            prompt = TryGetConflictPromptCopy(operations->FindTask(observation->taskId));
            return prompt.has_value();
        },
            {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"D1 normal identity-bound PermanentDelete consent"});
        auto* task = operations->FindTask(observation->taskId);
        if (! consent.conditionMet || ! task || prompt.value().bucket != Task::ConflictBucket::PermanentDeleteConfirmation ||
            prompt.value().sourcePath != source.native() || ! PromptHasAction(prompt.value(), Task::ConflictAction::PermanentDelete) ||
            ! task->SubmitConflictDecision(prompt.value().generation, Task::ConflictAction::PermanentDelete, false))
            return reject(L"normal identity-bound PermanentDelete consent was unavailable or rejected");
        task            = nullptr; // Never retain a Task borrow across the next message pump.
        const auto gate = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return WaitForSingleObject(entered.get(), 0u) == WAIT_OBJECT_0;
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"D1 first real immediate child removed"});
        if (! gate.conditionMet || observation->callbackCount != 0u || GetForegroundWindow() != foreground || ! GetEnvVarTrimmed(parentVariable).empty())
            return reject(L"actual post-removal one-shot gate unavailable or prematurely completed/activated");
        const auto removedPath = GetEnvVarTrimmed(removedVariable);
        const auto removedSeed = std::ranges::find_if(seeds, [&](const Seed& seed) noexcept {
            return seed.relative.parent_path().empty() && MakeExtendedPathForSelfTest(source / seed.relative) == removedPath;
        });
        if (removedSeed == seeds.end())
            return reject(L"gate receipt is not an exact immediate ordinary child path");
        std::vector<TreeEntry> cut;
        if (! capture(cut) || std::ranges::find(cut, removedSeed->relative.native(), &TreeEntry::relative) != cut.end())
            return reject(L"recorded removed child still present or surviving tree is not closed");
        for (const auto& seed : seeds)
        {
            if (! seed.relative.parent_path().empty() || seed.relative == removedSeed->relative)
                continue;
            const auto retained = std::ranges::find(cut, seed.relative.native(), &TreeEntry::relative);
            const auto before   = std::ranges::find(original, seed.relative.native(), &TreeEntry::relative);
            if (retained == cut.end() || before == original.end() || ! (*retained == *before))
                return reject(L"first immediate removal did not leave both other ordinary files unchanged");
        }
        uint64_t removedBytes = 0u;
        size_t removedEntries = 0u;
        for (const auto& before : original)
        {
            const auto retained = std::ranges::find(cut, before.relative, &TreeEntry::relative);
            if (retained == cut.end())
            {
                ++removedEntries;
                removedBytes += before.bytes.size();
                AppendLog(std::format(L"D1 task={} removed-at-cut path={} bytes={} BLAKE3={}",
                                      observation->taskId,
                                      before.relative,
                                      before.bytes.size(),
                                      (before.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u
                                          ? L"<directory; no content digest>"
                                          : Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(before.digest))));
            }
            else if (! (*retained == before))
                return reject(L"a surviving entry changed before the cancellation cut");
        }
        task = operations->FindTask(observation->taskId);
        if (! task)
            return reject(L"admitted worker disappeared while parked at the removal gate");
        task->RequestCancel();
        task = nullptr;
        if (! SetEvent(release.get()))
            return reject(L"case-owned removal gate could not be released");
        // No deliberate delay to age a provider cancellation cache. Resumption follows
        // the ordinary always-built traversal checks immediately after RequestCancel.
        const auto quiet = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return observation->callbackCount == 1u && ! operations->FindTask(observation->taskId) && ! operations->HasActiveOperations();
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"D1 canceled worker post-join quiet point"});
        if (! quiet.conditionMet || observation->result != HRESULT_FROM_WIN32(ERROR_CANCELLED))
            return reject(L"real operation did not retire once with cancellation");
        cleanupQuiet = true;
        std::vector<TreeEntry> after;
        if (! capture(after) || after != cut || GetEnvVarTrimmed(removedVariable) != removedPath || GetForegroundWindow() != foreground)
            return reject(L"post-cancel tree/path/bytes/BLAKE3 changed or foreground was acquired");
        const auto completion = state.completedTasks.find(observation->taskId);
        if (completion == state.completedTasks.end() || completion->second.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) ||
            completion->second.configuredMaxConcurrency != 1u || completion->second.progressCompletedBytes > removedBytes ||
            completion->second.sourceItemResults.size() != 1u || ! completion->second.sourceItemResults.front().has_value())
            return reject(L"actual completed-task concurrency/accounting/dense root result contradicts storage");
        const auto& result = completion->second.sourceItemResults.front().value();
        if (result.completion != FileOperations::ItemCompletion::Canceled || result.status != HRESULT_FROM_WIN32(ERROR_CANCELLED) ||
            result.sourceDisposition != FileOperations::SourceDisposition::Retained || result.publication != FileOperations::PublicationState::NotAttempted ||
            result.ownedStageDisposition != FileOperations::OwnedStageDisposition::NotApplicable)
            return reject(L"typed root axes lost Canceled/Retained/nonpublication/no-stage truth");
        FileOperationsPopupInternal::TaskSnapshot terminal{};
        if (! DebugGetFileOperationsPopupTaskSnapshot(operations->GetPopupHwndForSelfTest(), observation->taskId, terminal) || ! terminal.finished ||
            terminal.statusKind != FileOperationsPopupInternal::TaskSnapshot::StatusKind::Canceled)
            return reject(L"real completed card did not project Canceled");
        for (const auto& entry : after)
        {
            AppendLog(std::format(L"D1 task={} retained-at-cut-and-quiet path={} attributes=0x{:08X} bytes={} BLAKE3={}; exact bytes unchanged",
                                  observation->taskId,
                                  entry.relative,
                                  entry.attributes,
                                  entry.bytes.size(),
                                  (entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u
                                      ? L"<directory; no content digest>"
                                      : Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(entry.digest))));
        }
        if (FAILED(information->SetConfiguration(originalConfiguration.c_str())))
            return reject(L"provider configuration restoration failed after the quiet point");
        std::string restoredConfiguration;
        if (! BackupPluginConfiguration(information.get(), restoredConfiguration) || restoredConfiguration != originalConfiguration)
            return reject(L"provider configuration did not restore exactly");
        configurationRestored = true;
        AppendLog(std::format(
            L"D1 task={}: real removed child={}; removed entries={} bytes={}; closed surviving entries={} unchanged after cancel; "
            L"one callback, Canceled/Retained, serial provider/host budgets=1, observed progress bytes={}, configuration restored, foreground preserved",
            observation->taskId,
            removedPath,
            removedEntries,
            removedBytes,
            after.size(),
            completion->second.progressCompletedBytes));
        return true;
    }
    catch (const std::system_error& error)
    {
        // The canonical environment scope can throw on Win32 failure at this noexcept
        // selftest boundary; all gate/callback guards unwind before reporting failure.
        Debug::Error(L"D1 environment setup failed (error={}).", error.code().value());
        return reject(L"canonical environment scope setup failed");
    }
}

[[nodiscard]] bool RunFileOpsEmptySelectionAdmissionChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    auto* folder      = TryGetFolderWindow(state.mainWindow);
    auto* operations  = TryGetFileOps(folder);
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Empty-selection admission: {}", reason));
        return false;
    };
    if (! folder || ! operations || operations->HasActiveOperations())
        return reject(L"idle owner unavailable");
    const auto sourceFileSystem      = folder->GetFileSystem(FolderWindow::Pane::Left);
    const auto destinationFileSystem = folder->GetFileSystem(FolderWindow::Pane::Right);
    wil::com_ptr<IFileSystemIO> sourceIo;
    wil::com_ptr<IFileSystemIO> destinationIo;
    if (! sourceFileSystem || ! destinationFileSystem || ! sourceFileSystem.try_query_to(sourceIo.put()) ||
        ! destinationFileSystem.try_query_to(destinationIo.put()) || folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal ||
        folder->GetFileSystemPluginId(FolderWindow::Pane::Right) != kPluginIdLocal)
        return reject(L"actual Local pane filesystems unavailable");

    const HWND foreground  = GetForegroundWindow();
    const HWND popup       = operations->GetPopupHwndForSelfTest();
    const auto prompts     = HostGetTestPromptRequestCount();
    const auto payloads    = HostGetTestSynchronousPayloadRegistrationCount();
    const auto alerts      = HostGetTestWindowAlertOverlayCount();
    const auto hosts       = DxUi::DebugGetAttachedWindowHostCount();
    const auto attachments = DxUi::DebugGetSharedWindowHostAttachmentCountForThread(GetCurrentThreadId());
    std::vector<Task*> tasks;
    operations->CollectTasks(tasks);
    std::vector<uint64_t> originalTaskIds;
    for (const auto* task : tasks)
        originalTaskIds.push_back(task->GetId());
    std::ranges::sort(originalTaskIds);
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
    operations->CollectCompletedTasks(completed);
    std::vector<uint64_t> originalCompletedIds;
    for (const auto& task : completed)
        originalCompletedIds.push_back(task.taskId);
    std::ranges::sort(originalCompletedIds);

    const auto root        = state.tempRoot / L"empty-selection-admission";
    const auto source      = root / L"src" / L"source-sentinel.bin";
    const auto destination = root / L"dst" / L"destination-sentinel.bin";
    if (! RecreateEmptyDirectory(root))
        return reject(L"owned root unavailable");
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        // No callback borrows this stack. If admission regresses, join any unexpected
        // worker before removing its private filesystem fixture or retained result.
        if (! CancelOutstandingSelfTestOperations(operations))
        {
            AppendLog(L"Empty-selection admission cleanup: worker quiet point failed; owned root retained.");
            return;
        }
        operations->CollectCompletedTasks(completed);
        for (const auto& task : completed)
        {
            if (! std::ranges::binary_search(originalCompletedIds, task.taskId))
            {
                operations->DismissCompletedTask(task.taskId);
                operations->DebugRemoveDiagnosticsForTask(task.taskId);
            }
        }
        if (! popup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        static_cast<void>(SelfTest::RemoveAll(root));
    });
    if (! WriteFilledTestFile(source, 4099u, 0xA7u) || ! WriteFilledTestFile(destination, 2053u, 0x3Cu))
        return reject(L"distinct sentinel payloads unavailable");

    // This fixture is closed: two immediate directories, each with one regular
    // sentinel. Observe only these three directories; never descend into links.
    using InventoryEntry = std::pair<std::filesystem::path, std::filesystem::file_type>;
    const auto inventory = [&](std::vector<InventoryEntry>& result) noexcept
    {
        result.clear();
        std::error_code ec;
        for (const auto& directory : std::array{root, source.parent_path(), destination.parent_path()})
        {
            const auto status = std::filesystem::symlink_status(directory, ec);
            if (ec || std::filesystem::is_symlink(status) || ! std::filesystem::is_directory(status))
                return false;
            std::filesystem::directory_iterator entry(directory, std::filesystem::directory_options::none, ec);
            if (ec)
                return false;
            for (; entry != std::filesystem::directory_iterator{}; entry.increment(ec))
            {
                if (ec || result.size() >= 4u)
                    return false;
                const auto entryStatus = entry->symlink_status(ec);
                if (ec || std::filesystem::is_symlink(entryStatus) ||
                    (! std::filesystem::is_directory(entryStatus) && ! std::filesystem::is_regular_file(entryStatus)))
                    return false;
                result.emplace_back(entry->path().lexically_relative(root), entryStatus.type());
            }
            if (ec)
                return false;
        }
        std::ranges::sort(result);
        return true;
    };
    std::vector<InventoryEntry> expectedInventory{
        {source.parent_path().lexically_relative(root), std::filesystem::file_type::directory},
        {source.lexically_relative(root), std::filesystem::file_type::regular},
        {destination.parent_path().lexically_relative(root), std::filesystem::file_type::directory},
        {destination.lexically_relative(root), std::filesystem::file_type::regular},
    };
    std::ranges::sort(expectedInventory);
    std::vector<InventoryEntry> beforeInventory;
    if (! inventory(beforeInventory) || beforeInventory != expectedInventory)
        return reject(L"bounded original inventory differs from the two-directory/two-file fixture");
    const std::array paths{source, destination};
    const std::array io{sourceIo, destinationIo};
    const std::array<size_t, 2> sizes{4099u, 2053u};
    const std::array<unsigned char, 2> values{0xA7u, 0x3Cu};
    std::array<std::string, 2> originalBytes;
    std::array<Common::Crypto::Blake3Digest, 2> originalDigests;
    for (size_t index = 0u; index < paths.size(); ++index)
    {
        if (! ReadFileTextFsIo(io[index], paths[index], originalBytes[index]) || originalBytes[index].size() != sizes[index] ||
            ! std::ranges::all_of(originalBytes[index], [value = values[index]](char byte) noexcept { return static_cast<unsigned char>(byte) == value; }))
            return reject(L"original exact sentinel bytes unavailable");
        Common::Crypto::Blake3Hasher hasher;
        hasher.Update(std::as_bytes(std::span<const char>(originalBytes[index].data(), originalBytes[index].size())));
        originalDigests[index] = hasher.Finalize();
    }
    for (const auto operation : std::array<FileSystemOperation, 3>{{FILESYSTEM_COPY, FILESYSTEM_MOVE, FILESYSTEM_DELETE}})
    {
        uint64_t taskId  = 0xD9u;
        const HRESULT hr = operations->AdmitOperation(operation,
                                                      FolderWindow::Pane::Left,
                                                      operation == FILESYSTEM_DELETE ? std::nullopt : std::optional{FolderWindow::Pane::Right},
                                                      sourceFileSystem,
                                                      {},
                                                      operation == FILESYSTEM_DELETE ? std::filesystem::path{} : destination.parent_path(),
                                                      FILESYSTEM_FLAG_NONE,
                                                      false,
                                                      0u,
                                                      FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                      false,
                                                      destinationFileSystem,
                                                      &taskId);
        operations->CollectTasks(tasks);
        std::vector<uint64_t> taskIds;
        for (const auto* task : tasks)
            taskIds.push_back(task->GetId());
        std::ranges::sort(taskIds);
        operations->CollectCompletedTasks(completed);
        std::vector<uint64_t> completedIds;
        for (const auto& task : completed)
            completedIds.push_back(task.taskId);
        std::ranges::sort(completedIds);
        std::vector<InventoryEntry> afterInventory;
        if (hr != S_FALSE || taskId != 0u || taskIds != originalTaskIds || completedIds != originalCompletedIds || operations->HasActiveOperations() ||
            ! inventory(afterInventory) || afterInventory != beforeInventory)
            return reject(std::format(L"operation={} hr=0x{:08X} taskId={}: refusal/task/inventory contract failed",
                                      static_cast<unsigned int>(operation),
                                      static_cast<unsigned long>(hr),
                                      taskId));
        for (size_t index = 0u; index < paths.size(); ++index)
        {
            std::string retainedBytes;
            if (! ReadFileTextFsIo(io[index], paths[index], retainedBytes) || retainedBytes != originalBytes[index])
                return reject(L"empty selection changed exact sentinel bytes");
            Common::Crypto::Blake3Hasher hasher;
            hasher.Update(std::as_bytes(std::span<const char>(retainedBytes.data(), retainedBytes.size())));
            const auto retainedDigest = hasher.Finalize();
            if (retainedDigest != originalDigests[index])
                return reject(L"empty selection changed canonical sentinel BLAKE3");
            AppendLog(std::format(L"Empty-selection operation={} path={} bytes={} before BLAKE3={} after BLAKE3={}; exact bytes/inventory retained",
                                  static_cast<unsigned int>(operation),
                                  paths[index].native(),
                                  retainedBytes.size(),
                                  Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(originalDigests[index])),
                                  Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(retainedDigest))));
        }
        if (GetForegroundWindow() != foreground || operations->GetPopupHwndForSelfTest() != popup || HostGetTestPromptRequestCount() != prompts ||
            HostGetTestSynchronousPayloadRegistrationCount() != payloads || HostGetTestWindowAlertOverlayCount() != alerts ||
            DxUi::DebugGetAttachedWindowHostCount() != hosts || DxUi::DebugGetSharedWindowHostAttachmentCountForThread(GetCurrentThreadId()) != attachments)
            return reject(L"empty selection changed foreground, popup or host observations");
        AppendLog(std::format(L"Empty-selection operation={}: S_FALSE, taskId=0, no new active/completed task; foreground/popup/host counts retained",
                              static_cast<unsigned int>(operation)));
    }
    return true;
}

[[nodiscard]] bool RunFileOpsCompletionRecoveryChecks(SelfTestState& state) noexcept
{
    using namespace std::chrono_literals;
    auto* folder      = TryGetFolderWindow(state.mainWindow);
    auto* operations  = TryGetFileOps(folder);
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps completion recovery: {}", reason));
        return false;
    };
    if (! folder || ! operations || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner or initial quiet point unavailable");

    const HWND foreground  = GetForegroundWindow();
    const DWORD ownerTid   = GetCurrentThreadId();
    const bool autoDismiss = operations->GetAutoDismissSuccess();
    operations->SetAutoDismissSuccess(false);
    operations->DebugSetTaskPresentationNowTickForSelfTest(10'000u);
    const auto restore = wil::scope_exit([&]() noexcept
    {
        ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
        ReleaseFileOpsFallbackWorkerReturnPauseForSelfTest();
        operations->DebugForceAllFileOperationCompletedPostsFailure(false);
        operations->DebugSuppressTaskPresentationPostsForSelfTest(false);
        operations->DebugForceAllOrphanedCompletionDrainSubmissionFailures(false);
        static_cast<void>(CancelOutstandingSelfTestOperations(operations));
        operations->DebugSetTaskPresentationNowTickForSelfTest(std::nullopt);
        operations->SetAutoDismissSuccess(autoDismiss);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
    });

    const auto root = state.tempRoot / L"completion-recovery";
    if (! RecreateEmptyDirectory(root))
        return reject(L"owned root could not be created");
    for (unsigned int scenario = 0u; scenario < 5u; ++scenario)
    {
        operations->DebugResetPopupForSelfTest();
        SendMessageW(folder->GetHwnd(), WndMsg::kFileOperationPresentationChanged, 0, 0);
        operations->DebugSuppressTaskPresentationPostsForSelfTest(true);
        const auto caseRoot    = root / std::format(L"case-{}", scenario);
        const auto source      = caseRoot / L"source" / L"payload.bin";
        const auto reference   = caseRoot / L"reference.bin";
        const auto destination = scenario == 3u ? source.parent_path() / L"renamed.bin" : caseRoot / L"destination" / L"payload.bin";
        if (! WriteTestFile(source, 4096u + scenario) || ! WriteTestFile(reference, 4096u + scenario) || ! SelfTest::EnsureDirectory(destination.parent_path()))
            return reject(L"owned byte references could not be created");

        uint64_t taskId = 0u;
        std::atomic<unsigned int> callbackCount{0u};
        std::atomic<DWORD> callbackTid{0u};
        std::atomic<HRESULT> callbackResult{E_PENDING};
        const uint64_t subscription = folder->AddFileOperationCompletedCallback([&](const FolderWindow::FileOperationCompletedEvent& event) noexcept
        {
            if (event.taskId == taskId)
            {
                callbackTid.store(GetCurrentThreadId(), std::memory_order_release);
                callbackResult.store(event.hr, std::memory_order_release);
                callbackCount.fetch_add(1u, std::memory_order_acq_rel);
            }
        });
        // The drain scope precedes unsubscribing, so a failed witness cannot leave a
        // callback borrowing the case's stack after return.
        const auto unsubscribe = wil::scope_exit([&]() noexcept { folder->RemoveFileOperationCompletedCallback(subscription); });
        const auto drain       = wil::scope_exit([&]() noexcept
        {
            ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
            ReleaseFileOpsFallbackWorkerReturnPauseForSelfTest();
            operations->DebugForceAllFileOperationCompletedPostsFailure(false);
            operations->DebugSuppressTaskPresentationPostsForSelfTest(false);
            operations->DebugForceAllOrphanedCompletionDrainSubmissionFailures(false);
            const bool quiet = CancelOutstandingSelfTestOperations(operations);
            if (quiet && taskId != 0u)
            {
                // Each scenario proves exactly one retained result before retiring only its own completion.
                operations->DismissCompletedTask(taskId);
                operations->DebugRemoveDiagnosticsForTask(taskId);
                state.completedTasks.erase(taskId);
            }
            else if (! quiet)
            {
                state.retainTempRootOnFailure = true;
                AppendLog(std::format(L"Completion recovery scenario {} could not drain its owned task {} during cleanup.", scenario, taskId));
            }
        });
        SetFileOpsPostFinishedCompletionPauseForSelfTest(true);
        SetFileOpsFallbackWorkerReturnPauseForSelfTest(true);
        operations->DebugForceAllFileOperationCompletedPostsFailure(scenario != 2u);
        operations->DebugForceAllOrphanedCompletionDrainSubmissionFailures(scenario == 4u);
        operations->DebugForceNextFileOperationCompletedPostFailure();
        if (scenario == 1u)
            operations->DebugForceNextOrphanedCompletionDrainSubmissionFailure();
        std::optional<uint64_t> admitted;
        HRESULT admissionHr = E_PENDING;
        if (scenario == 3u)
        {
            const auto paneFileSystem = folder->GetFileSystem(FolderWindow::Pane::Left);
            if (! paneFileSystem)
                return reject(L"scheduled rename left-pane filesystem unavailable");
            BatchRenameExecutionOp rename{};
            rename.originalSource = source;
            rename.finalLeaf      = destination.filename().native();
            const HRESULT factsHr = PopulateBatchRenameExactResultProviderFacts(paneFileSystem.get(), rename);
            if (factsHr != S_OK)
                return reject(
                    std::format(L"scenario {} scheduled rename provider facts unavailable: hr=0x{:08X}", scenario, static_cast<unsigned long>(factsHr)));
            uint64_t renamedTaskId = 0u;
            std::vector<BatchRenameExecutionOp> renames;
            renames.push_back(std::move(rename));
            // Scheduled rename admission requires the owner's exact pane instance.
            admissionHr = operations->AdmitBatchRename(FolderWindow::Pane::Left, paneFileSystem, std::move(renames), {}, {}, &renamedTaskId);
            if (admissionHr == S_OK && renamedTaskId != 0u)
                admitted = renamedTaskId;
        }
        else
        {
            uint64_t copiedTaskId = 0u;
            admissionHr           = operations->AdmitOperation(FILESYSTEM_COPY,
                                                               FolderWindow::Pane::Left,
                                                               FolderWindow::Pane::Right,
                                                               state.fsLocal,
                                                               {source},
                                                               destination.parent_path(),
                                                               FILESYSTEM_FLAG_NONE,
                                                               false,
                                                               0u,
                                                               FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                               false,
                                                               state.fsLocal,
                                                               &copiedTaskId);
            if (SUCCEEDED(admissionHr) && copiedTaskId != 0u)
                admitted = copiedTaskId;
        }
        if (! admitted.has_value())
            return reject(std::format(L"scenario {} real Local {} admission failed: hr=0x{:08X}",
                                      scenario,
                                      scenario == 3u ? L"BatchRename" : L"Copy",
                                      static_cast<unsigned long>(admissionHr)));
        taskId              = admitted.value();
        const auto terminal = RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept {
            return HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest();
        }, {.timeout = SelfTest::Scale(8s), .operationName = L"completion flag published before worker return"});
        if (! terminal.conditionMet || callbackCount.load(std::memory_order_acquire) != 0u || ! operations->HasActiveOperations())
            return reject(L"terminal flag incorrectly established a worker quiet point");
        // Retain a summary in this fixture even for a sub-500ms Copy. This changes
        // only presentation retention; the actual Local mutation/result remains untouched.
        auto* task = operations->FindTask(taskId);
        if (! task)
            return reject(L"parked worker lost Task storage");
        task->_presentationState.store(FolderWindow::FileOperationState::Task::TaskPresentationState::Presented, std::memory_order_release);
        ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
        const auto parked = RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept {
            return HasFileOpsFallbackWorkerReturnPauseEnteredForSelfTest();
        }, {.timeout = SelfTest::Scale(8s), .operationName = L"fallback worker parked after transfer or admission refusal"});
        if (! parked.conditionMet || operations->DebugFallbackCompletionCountForSelfTest() != 1u || callbackCount.load(std::memory_order_acquire) != 0u)
            return reject(L"pending fallback was consumed before its worker joined");
        if (scenario == 1u && operations->DebugLastRestoredCompletionWorkerTaskId() != taskId)
            return reject(L"threadpool refusal did not restore worker ownership");

        // Exercise the real early payload-less wake and a stale nonzero registry
        // token. Neither may borrow the pending completion or destroy a live worker.
        SendMessageW(folder->GetHwnd(), WndMsg::kFileOperationCompleted, static_cast<WPARAM>(taskId), 0);
        SendMessageW(folder->GetHwnd(), WndMsg::kFileOperationCompleted, static_cast<WPARAM>(taskId), 1);
        const uint64_t ticksBefore = operations->DebugCompletionRecoveryTickCountForSelfTest();
        const auto responsive      = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return operations->DebugCompletionRecoveryTickCountForSelfTest() >= ticksBefore + 2u;
        }, {.timeout = SelfTest::Scale(1500ms), .operationName = L"owner heartbeat while orphan worker remains parked"});
        if (! responsive.conditionMet || responsive.elapsed > SelfTest::Scale(1500ms) || callbackCount.load(std::memory_order_acquire) != 0u ||
            ! operations->FindTask(taskId) || operations->DebugFallbackCompletionCountForSelfTest() != 1u)
            return reject(L"owner joined a parked worker, stopped its heartbeat, or consumed an unready fallback");

        ReleaseFileOpsFallbackWorkerReturnPauseForSelfTest();
        const auto completed = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return callbackCount.load(std::memory_order_acquire) == 1u && ! operations->FindTask(taskId);
        }, {.timeout = SelfTest::Scale(8s), .operationName = L"owner recovery without any successful completion post"});
        if (! completed.conditionMet || callbackTid.load(std::memory_order_acquire) != ownerTid ||
            operations->DebugLastFileOperationCompletionApplyTid() != ownerTid || operations->DebugLastFileOperationCompletionWorkerTid() == ownerTid ||
            operations->DebugFallbackCompletionCountForSelfTest() != 0u)
            return reject(L"completion did not apply exactly once on the owner with empty fallback/task ownership");
        SendMessageW(folder->GetHwnd(), WndMsg::kFileOperationCompleted, static_cast<WPARAM>(taskId), 0);
        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        operations->CollectCompletedTasks(summaries);
        const DWORD sourceAttributes    = GetFileAttributesW(source.c_str());
        const bool renamedSourceAbsent  = sourceAttributes == INVALID_FILE_ATTRIBUTES && GetLastError() == ERROR_FILE_NOT_FOUND;
        const auto summaryCount         = std::ranges::count_if(summaries, [taskId](const auto& summary) noexcept { return summary.taskId == taskId; });
        const bool sourcePreserved      = scenario == 3u ? renamedSourceAbsent : FilesEqualBytes(source, reference);
        const bool destinationPreserved = FilesEqualBytes(destination, reference);
        const bool foregroundPreserved  = GetForegroundWindow() == foreground;
        if (callbackCount.load(std::memory_order_acquire) != 1u || callbackResult.load(std::memory_order_acquire) != S_OK || summaryCount != 1 ||
            ! sourcePreserved || ! destinationPreserved || ! foregroundPreserved)
            return reject(
                std::format(L"scenario {} callbackCount={} result=0x{:08X} summaryCount={} sourcePreserved={} destinationPreserved={} foregroundPreserved={}",
                            scenario,
                            callbackCount.load(std::memory_order_acquire),
                            static_cast<unsigned long>(callbackResult.load(std::memory_order_acquire)),
                            summaryCount,
                            sourcePreserved,
                            destinationPreserved,
                            foregroundPreserved));
        AppendLog(std::format(L"Completion recovery scenario {}: ownerTid={} callbackCount={} joined=true bytes=preserved",
                              scenario,
                              ownerTid,
                              callbackCount.load(std::memory_order_acquire)));
    }
    return true;
}

[[nodiscard]] bool RunFileOpsStoppingSnapshotChecks(SelfTestState& state) noexcept
{
    using namespace FileOperationsPopupInternal;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps stopping snapshot: {}", reason));
        return false;
    };
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation window guard unavailable");
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! operations || ! state.fsLocal || operations->HasActiveOperations())
        return reject(L"fixture requires an idle Local operation state");

    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui && ! g_settings.ui->language.empty() && g_settings.ui->language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui->language};
    auto restoreLanguage = wil::scope_exit([&]() noexcept { static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage)); });
    const std::array<std::pair<std::wstring_view, std::wstring_view>, 5> cultures{
        {{L"en-US", L"Stopping"}, {L"fr-FR", L"Arrêt en cours"}, {L"cs-CZ", L"Zastavování"}, {L"ja-JP", L"停止中"}, {L"sk-SK", L"Zastavovanie"}}};
    for (const auto& [culture, expected] : cultures)
    {
        if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, std::wstring(culture)})) ||
            LoadStringResource(nullptr, IDS_FILEOPS_STATUS_STOPPING) != expected || ! DebugValidateFileOperationsStoppingPresentation())
            return reject(std::format(L"{} exact localized label or pure status/summary/taskbar precedence failed", culture));
    }
    static_cast<void>(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, L"en-US"}));

    const bool previousFooterOnly  = operations->GetPopupFooterOnly();
    const bool previousCompact     = operations->GetPopupCompactDensity();
    const bool previousAutoDismiss = operations->GetAutoDismissSuccess();
    operations->SetPopupFooterOnly(false);
    operations->SetPopupCompactDensity(false);
    operations->SetAutoDismissSuccess(false);
    auto restoreSettings                  = wil::scope_exit([&]() noexcept
    {
        operations->SetPopupFooterOnly(previousFooterOnly);
        operations->SetPopupCompactDensity(previousCompact);
        operations->SetAutoDismissSuccess(previousAutoDismiss);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(10000u));
    });
    const auto root                       = state.tempRoot / L"stopping-snapshot";
    const auto source                     = root / L"src" / L"collision.bin";
    const auto destination                = root / L"dst" / L"collision.bin";
    const auto sourceReference            = root / L"source-reference.bin";
    const auto destinationReference       = root / L"destination-reference.bin";
    const auto hiddenSource               = root / L"hidden-src" / L"hidden.bin";
    const auto hiddenDestination          = root / L"hidden-dst" / L"hidden.bin";
    const auto hiddenSourceReference      = root / L"hidden-source-reference.bin";
    const auto hiddenDestinationReference = root / L"hidden-destination-reference.bin";
    if (! RecreateEmptyDirectory(root) || ! WriteTestFile(source, 4096u) || ! WriteTestFile(sourceReference, 4096u) || ! WriteTestFile(destination, 8192u) ||
        ! WriteTestFile(destinationReference, 8192u) || ! WriteTestFile(hiddenSource, 1024u) || ! WriteTestFile(hiddenSourceReference, 1024u) ||
        ! WriteTestFile(hiddenDestination, 2048u) || ! WriteTestFile(hiddenDestinationReference, 2048u))
        return reject(L"owned collision scratch could not be created");

    std::optional<uint64_t> taskId;
    std::optional<uint64_t> hiddenTaskId;
    // The readiness callback owns its gate even if a failed completion drain
    // outlives this stack frame. Its wait also has a finite failure fallback.
    struct HiddenReadinessGate final
    {
        std::atomic<bool> entered{false};
        wil::unique_handle release{CreateEventW(nullptr, TRUE, FALSE, nullptr)};
    };
    const auto hiddenGate = std::make_shared<HiddenReadinessGate>();
    if (! hiddenGate->release)
        return reject(L"hidden readiness release event could not be created");
    // This existing decoration barrier observes only explicit release/bailout,
    // never Task cancellation, so the unfinished observation is deterministic.
    SetFileOpsConflictMetadataPauseForSelfTest(true, static_cast<ULONGLONG>(SelfTest::Scale(std::chrono::milliseconds(30'000)).count()));
    auto releaseAndDrain  = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEvent(hiddenGate->release.get()));
        ReleaseFileOpsConflictMetadataPauseForSelfTest();
        SetFileOpsConflictMetadataPauseForSelfTest(false, 0u);
        operations->DebugSetTaskPresentationNowTickForSelfTest(std::nullopt);
        if (hiddenTaskId)
            if (auto* hiddenTask = operations->FindTask(hiddenTaskId.value()))
                hiddenTask->RequestCancel();
        if (taskId)
        {
            if (auto* task = operations->FindTask(taskId.value()))
                task->RequestCancel();
            static_cast<void>(CancelOutstandingSelfTestOperations(operations));
            operations->DismissCompletedTask(taskId.value());
            if (hiddenTaskId)
                operations->DismissCompletedTask(hiddenTaskId.value());
            if (const HWND popup = operations->GetPopupHwndForSelfTest(); popup && ! operations->HasActiveOperations())
                PostMessageW(popup, WM_CLOSE, 0, 0);
        }
    });
    const HWND foreground = GetForegroundWindow();
    taskId                = StartFileOperationAndGetId(operations,
                                                       FILESYSTEM_COPY,
                                                       FolderWindow::Pane::Left,
                                                       FolderWindow::Pane::Right,
                                                       state.fsLocal,
                                                       {source},
                                                       destination.parent_path(),
                                                       FILESYSTEM_FLAG_NONE,
                                                       false,
                                                       0u,
                                                       FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                       false,
                                                       state.fsLocal);
    if (! taskId)
        return reject(L"Local collision was not admitted");
    const auto blocked = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return HasFileOpsConflictMetadataPauseEnteredForSelfTest() && operations->GetPopupHwndForSelfTest() != nullptr;
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"Stopping snapshot metadata barrier"});
    if (! blocked.conditionMet)
        return reject(blocked.timeoutDiagnostic);
    const HWND popup = operations->GetPopupHwndForSelfTest();
    TaskSnapshot before{};
    auto* task = operations->FindTask(taskId.value());
    if (! task || ! DebugGetFileOperationsPopupTaskSnapshot(popup, taskId.value(), before) || ! before.conflict.active || before.finished)
        return reject(L"barrier did not retain an unfinished real conflict");
    const uint64_t initialLimit = before.desiredSpeedLimitBytesPerSecond;
    task->RequestCancel();
    task->SetPaused(true); // Stale row and bulk-pause callers must not re-arm a pause.
    task->SetDesiredSpeedLimit(1234u);
    operations->SetAllRunningTasksPaused(true);
    if (operations->RunQueuedTaskNow(taskId.value()) || operations->MoveQueuedTask(taskId.value(), true))
        return reject(L"stopping task still accepts queue decisions");
    TaskSnapshot stopping{};
    if (! HasFileOpsConflictMetadataPauseEnteredForSelfTest() || ! DebugGetFileOperationsPopupTaskSnapshot(popup, taskId.value(), stopping) ||
        ! stopping.cancellationRequested || stopping.finished || stopping.statusKind != TaskSnapshot::StatusKind::Stopping || stopping.paused ||
        stopping.conflict.active || stopping.conflict.buttonsPublishable || stopping.conflict.actionCount != 0u || stopping.conflict.inspection ||
        stopping.canMoveQueueUp || stopping.canMoveQueueDown || stopping.desiredSpeedLimitBytesPerSecond != initialLimit ||
        stopping.completedBytes != before.completedBytes || stopping.totalBytes != before.totalBytes ||
        stopping.currentSourcePath != before.currentSourcePath || stopping.currentDestinationPath != before.currentDestinationPath)
        return reject(L"RequestCancel did not synchronously project authority, withdraw decisions and retain facts");

    constexpr std::array<PopupHitTest::Kind, 10> decisions{{PopupHitTest::Kind::TaskPause,
                                                            PopupHitTest::Kind::TaskCancel,
                                                            PopupHitTest::Kind::TaskStartNow,
                                                            PopupHitTest::Kind::TaskQueueMoveUp,
                                                            PopupHitTest::Kind::TaskQueueMoveDown,
                                                            PopupHitTest::Kind::TaskSpeedLimit,
                                                            PopupHitTest::Kind::TaskSkip,
                                                            PopupHitTest::Kind::TaskConflictAction,
                                                            PopupHitTest::Kind::TaskConflictToggleApplyToAll,
                                                            PopupHitTest::Kind::TaskConflictMore}};
    for (const auto kind : decisions)
    {
        PopupSelfTestInvoke invoke{};
        invoke.kind             = kind;
        invoke.taskId           = taskId.value();
        invoke.promptGeneration = before.conflict.generation;
        if (DebugInvokeFileOperationsPopup(popup, invoke))
            return reject(L"stale task decision was dispatched after cancellation");
    }
    const auto previousPromptOverride = HostGetTestPromptResultOverride();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    auto restorePromptOverride = wil::scope_exit([&]() noexcept
    {
        if (previousPromptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(previousPromptOverride);
    });
    const auto promptCount     = HostGetTestPromptRequestCount();
    PopupSelfTestInvoke cancelAll{};
    cancelAll.kind = PopupHitTest::Kind::FooterCancelAll;
    static_cast<void>(DebugInvokeFileOperationsPopup(popup, cancelAll));
    if (HostGetTestPromptRequestCount() != promptCount)
        return reject(L"all-stopping work opened another Cancel all consent prompt");
    for (size_t pass = 0u; pass < 3u; ++pass)
        RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    PopupLayoutDebugSnapshot layout{};
    layout.taskId = taskId.value();
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, layout) || ! layout.usesDxUiHost || layout.taskStatusKind != TaskSnapshot::StatusKind::Stopping ||
        layout.taskStatusVisualTone != static_cast<uint32_t>(PopupStatusVisualTone::Neutral) || ! layout.taskStatusChipActuallyPainted ||
        layout.taskStatusChipPaintedText != L"Stopping" || layout.taskStatusHeaderPaintedText.find(L"Stopping") != std::wstring::npos ||
        layout.taskStatusActiveStateCount != 1u || layout.globalRunningCount != 1u || layout.globalNeedAttentionCount != 0u ||
        layout.footerPauseResumeAllVisible || layout.footerAggregateEtaVisible || layout.taskPauseVisible || layout.taskCancelVisible ||
        layout.taskSpeedLimitVisible || layout.taskStartNowVisible || layout.taskQueueMoveUpVisible || layout.taskQueueMoveDownVisible ||
        layout.conflictButtonsPublishable)
        return reject(L"retained/native painted model retained actionable or duplicate stopping presentation");
    for (const auto kind : decisions)
    {
        wil::com_ptr_nothrow<IInvokeProvider> action;
        FileOpsTaskAccessibilityProbe probe{};
        if (FindFileOpsTaskAction(popup, kind, taskId.value(), action, &probe))
            return reject(L"UIA exposes an actionable stopping task decision");
        if (! probe.traversalComplete || probe.statusName != L"Stopping")
            return reject(L"UIA stopping status was absent or its bounded provider traversal failed");
    }
    if (! FilesEqualBytes(source, sourceReference) || ! FilesEqualBytes(destination, destinationReference) || GetForegroundWindow() != foreground)
        return reject(L"cancel-pending observation changed source/destination bytes or foreground");
    if (! DebugEnterFileOperationsPopupFailureSurface(popup))
        return reject(L"native failure surface could not be exercised without activation");
    PopupLayoutDebugSnapshot native{};
    native.taskId           = taskId.value();
    const auto cancelLabel  = LoadStringResource(nullptr, IDS_FILEOPS_BTN_CANCEL_ALL);
    const HWND nativeCancel = FindWindowExW(popup, nullptr, L"BUTTON", cancelLabel.c_str());
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, native) || ! native.failureSurfaceActive || native.usesDxUiHost ||
        native.failureSurfaceText.find(L"Stopping") == std::wstring::npos || ! nativeCancel || IsWindowEnabled(nativeCancel) != FALSE)
        return reject(L"native failure surface did not expose stopping text with Cancel all unavailable");
    SendMessageW(nativeCancel, BM_CLICK, 0, 0);
    if (HostGetTestPromptRequestCount() != promptCount || GetForegroundWindow() != foreground)
        return reject(L"native stopping cancel binding opened consent or changed foreground");

    operations->DebugSetTaskPresentationNowTickForSelfTest(10'000u);
    hiddenTaskId = StartFileOperationAndGetId(operations,
                                              FILESYSTEM_COPY,
                                              FolderWindow::Pane::Left,
                                              FolderWindow::Pane::Right,
                                              state.fsLocal,
                                              {hiddenSource},
                                              hiddenDestination.parent_path(),
                                              FILESYSTEM_FLAG_NONE,
                                              false,
                                              0u,
                                              FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                              false,
                                              state.fsLocal,
                                              {},
                                              std::nullopt,
                                              {},
                                              [hiddenGate]() noexcept
    {
        hiddenGate->entered.store(true, std::memory_order_release);
        const DWORD wait = WaitForSingleObject(hiddenGate->release.get(), 30'000u);
        return wait == WAIT_OBJECT_0 ? S_OK : wait == WAIT_TIMEOUT ? HRESULT_FROM_WIN32(ERROR_TIMEOUT) : HRESULT_FROM_WIN32(GetLastError());
    });
    if (! hiddenTaskId)
        return reject(L"hidden Preparing Copy was not admitted");
    const auto hiddenBlocked = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return hiddenGate->entered.load(std::memory_order_acquire);
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"Hidden Preparing Copy readiness barrier"});
    if (! hiddenBlocked.conditionMet)
        return reject(hiddenBlocked.timeoutDiagnostic);
    const auto* hiddenTask = operations->FindTask(hiddenTaskId.value());
    TaskSnapshot hiddenSnapshot{};
    if (! hiddenTask || ! hiddenTask->IsPresentationHidden() || hiddenTask->CancelRequestedDurable() ||
        hiddenTask->GetLifecyclePhase() != FileOperations::TaskLifecyclePhase::AwaitingAcceptance ||
        hiddenTask->_selectedRootReadinessComplete.load(std::memory_order_acquire) ||
        DebugGetFileOperationsPopupTaskSnapshot(popup, hiddenTaskId.value(), hiddenSnapshot) || HostGetTestPromptRequestCount() != promptCount ||
        ! HasFileOpsConflictMetadataPauseEnteredForSelfTest())
        return reject(L"readiness-held Copy was not truly hidden outside the visible stopping snapshot");
    // The native fallback refreshes engine availability on its governed timer.
    // Pump that normal path while the reveal clock and readiness gate stay held.
    const auto cancelAvailable = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return IsWindowEnabled(nativeCancel) != FALSE;
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"Cancel all includes hidden Preparing Copy"});
    if (! cancelAvailable.conditionMet)
        return reject(L"engine Cancel all availability excluded the hidden Preparing Copy");
    hiddenTask = operations->FindTask(hiddenTaskId.value());
    if (! hiddenTask || ! hiddenTask->IsPresentationHidden() || DebugGetFileOperationsPopupTaskSnapshot(popup, hiddenTaskId.value(), hiddenSnapshot))
        return reject(L"hidden Copy became visible before global cancellation consent");
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    SendMessageW(nativeCancel, BM_CLICK, 0, 0);
    hiddenTask = operations->FindTask(hiddenTaskId.value());
    TaskSnapshot stillStopping{};
    if (HostGetTestPromptRequestCount() != promptCount + 1u || ! hiddenTask || ! hiddenTask->CancelRequestedDurable() ||
        ! HasFileOpsConflictMetadataPauseEnteredForSelfTest() || ! DebugGetFileOperationsPopupTaskSnapshot(popup, taskId.value(), stillStopping) ||
        stillStopping.finished || ! stillStopping.cancellationRequested || stillStopping.statusKind != TaskSnapshot::StatusKind::Stopping ||
        ! FilesEqualBytes(source, sourceReference) || ! FilesEqualBytes(destination, destinationReference) ||
        ! FilesEqualBytes(hiddenSource, hiddenSourceReference) || ! FilesEqualBytes(hiddenDestination, hiddenDestinationReference) ||
        GetForegroundWindow() != foreground)
        return reject(L"one global consent did not cancel hidden work while preserving visible stopping work, exact bytes and foreground");
    static_cast<void>(SetEvent(hiddenGate->release.get()));
    ReleaseFileOpsConflictMetadataPauseForSelfTest();
    operations->DebugSetTaskPresentationNowTickForSelfTest(std::nullopt);
    const auto finished = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return ! operations->HasActiveOperations();
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"Stopping snapshot cancellation completion"});
    if (! finished.conditionMet)
        return reject(finished.timeoutDiagnostic);
    TaskSnapshot terminal{};
    if (! DebugGetFileOperationsPopupTaskSnapshot(popup, taskId.value(), terminal) || ! terminal.finished ||
        terminal.statusKind != TaskSnapshot::StatusKind::Canceled || ! FilesEqualBytes(source, sourceReference) ||
        ! FilesEqualBytes(destination, destinationReference) || ! FilesEqualBytes(hiddenSource, hiddenSourceReference) ||
        ! FilesEqualBytes(hiddenDestination, hiddenDestinationReference))
        return reject(L"actual terminal cancellation or preserved collision bytes were lost");
    TaskSnapshot hiddenTerminal{};
    if (! DebugGetFileOperationsPopupTaskSnapshot(popup, hiddenTaskId.value(), hiddenTerminal) || ! hiddenTerminal.finished ||
        hiddenTerminal.statusKind != TaskSnapshot::StatusKind::Canceled || HostGetTestPromptRequestCount() != promptCount + 1u)
        return reject(L"hidden Copy lost its actual terminal cancellation or repeated global consent");
    return true;
}

[[nodiscard]] bool RunFileOpsCompletionOwnerDestroyChecks(SelfTestState& state) noexcept
{
    using namespace std::chrono_literals;
    auto* primary           = TryGetFolderWindow(state.mainWindow);
    auto* primaryOperations = TryGetFileOps(primary);
    if (! primary || ! primaryOperations || primaryOperations->HasActiveOperations())
        return false;
    const HWND foreground  = GetForegroundWindow();
    const HWND primaryHwnd = primary->GetHwnd();
    const auto root        = state.tempRoot / L"completion-owner-destroy";
    const auto source      = root / L"src" / L"original.bin";
    const auto reference   = root / L"reference.bin";
    const auto destination = root / L"dst" / L"original.bin";
    if (! RecreateEmptyDirectory(root) || ! WriteTestFile(source, 4097u) || ! WriteTestFile(reference, 4097u) ||
        ! SelfTest::EnsureDirectory(destination.parent_path()))
        return false;

    Common::Settings::Settings settings;
    settings.fileOperations.emplace();
    wil::unique_hwnd container(
        CreateWindowExW(0, L"STATIC", L"Owned completion teardown fixture", WS_POPUP, 0, 0, 640, 480, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr));
    if (! container)
        return false;
    FolderWindow secondary;
    secondary.SetSettings(&settings);
    const HWND ownerHwnd = secondary.Create(container.get(), 0, 0, 640, 480);
    auto* operations     = secondary.DebugGetFileOperationState();
    if (! ownerHwnd || ! operations)
        return false;
    const HRESULT sourcePaneHr       = secondary.SetFileSystemPluginForPane(FolderWindow::Pane::Left, kPluginIdLocal);
    const HRESULT destinationPaneHr  = secondary.SetFileSystemPluginForPane(FolderWindow::Pane::Right, kPluginIdLocal);
    const auto sourceFileSystem      = secondary.GetFileSystem(FolderWindow::Pane::Left);
    const auto destinationFileSystem = secondary.GetFileSystem(FolderWindow::Pane::Right);
    const std::wstring sourcePlugin(secondary.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring destinationPlugin(secondary.GetFileSystemPluginId(FolderWindow::Pane::Right));
    if (FAILED(sourcePaneHr) || FAILED(destinationPaneHr) || ! sourceFileSystem || ! destinationFileSystem)
    {
        AppendLog(std::format(L"Completion secondary owner Local pane setup failed: sourceHr=0x{:08X} destinationHr=0x{:08X} sourceFs={} destinationFs={}",
                              static_cast<unsigned long>(sourcePaneHr),
                              static_cast<unsigned long>(destinationPaneHr),
                              static_cast<bool>(sourceFileSystem),
                              static_cast<bool>(destinationFileSystem)));
        return false;
    }
    uint64_t taskId = 0u;
    std::optional<CompletedTaskInfo> savedCompletion;
    std::atomic<unsigned int> callbacks{0u};
    std::atomic<HRESULT> result{E_PENDING};
    const uint64_t subscription = secondary.AddFileOperationCompletedCallback([&](const FolderWindow::FileOperationCompletedEvent& event) noexcept
    {
        result.store(event.hr, std::memory_order_release);
        callbacks.fetch_add(1u, std::memory_order_acq_rel);
        secondary.Destroy();
    });
    const auto destroy          = wil::scope_exit([&]() noexcept
    {
        ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
        ReleaseFileOpsFallbackWorkerReturnPauseForSelfTest();
        secondary.RemoveFileOperationCompletedCallback(subscription);
        secondary.Destroy();
        if (taskId != 0u)
        {
            if (savedCompletion.has_value())
                state.completedTasks.insert_or_assign(taskId, savedCompletion.value());
            else
                state.completedTasks.erase(taskId);
        }
    });
    operations->DebugForceAllFileOperationCompletedPostsFailure(true);
    operations->DebugSuppressTaskPresentationPostsForSelfTest(true);
    SetFileOpsPostFinishedCompletionPauseForSelfTest(true);
    // Both endpoints belong to this actual secondary owner, rather than borrowing
    // a primary fixture instance or inferring an omitted destination filesystem.
    const HRESULT admissionHr = operations->AdmitOperation(FILESYSTEM_COPY,
                                                           FolderWindow::Pane::Left,
                                                           FolderWindow::Pane::Right,
                                                           sourceFileSystem,
                                                           {source},
                                                           destination.parent_path(),
                                                           FILESYSTEM_FLAG_NONE,
                                                           false,
                                                           0u,
                                                           FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                           false,
                                                           destinationFileSystem,
                                                           &taskId);
    if (FAILED(admissionHr) || taskId == 0u)
    {
        AppendLog(std::format(L"Completion secondary owner Copy admission failed: hr=0x{:08X} taskId={} sourcePlugin={} destinationPlugin={}",
                              static_cast<unsigned long>(admissionHr),
                              taskId,
                              sourcePlugin,
                              destinationPlugin));
        return false;
    }
    const auto priorCompletion = state.completedTasks.find(taskId);
    if (priorCompletion != state.completedTasks.end())
        savedCompletion = priorCompletion->second;
    const auto parked = RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept {
        return HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest();
    }, {.timeout = SelfTest::Scale(8s), .operationName = L"secondary owner completion before callbacks"});
    if (! parked.conditionMet)
        return false;
    ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
    const auto destroyed = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return callbacks.load(std::memory_order_acquire) != 0u && ! IsWindow(ownerHwnd);
    }, {.timeout = SelfTest::Scale(8s), .operationName = L"completion callback destroys actual secondary owner"});
    const bool ownerGone = IsWindow(ownerHwnd) == FALSE;
    // The ordinary debug getter initializes on demand. A teardown witness must
    // observe the old state without recreating it after the HWND has gone.
    const bool stateGone             = secondary.DebugGetFileOperationState(false) == nullptr;
    const bool primaryPreserved      = primary->GetHwnd() == primaryHwnd;
    const bool primaryStatePreserved = TryGetFileOps(primary) == primaryOperations;
    const bool foregroundPreserved   = GetForegroundWindow() == foreground;
    const bool sourcePreserved       = FilesEqualBytes(source, reference);
    const bool destinationPreserved  = FilesEqualBytes(destination, reference);
    const HRESULT completionHr       = result.load(std::memory_order_acquire);
    const bool passed = destroyed.conditionMet && callbacks.load(std::memory_order_acquire) == 1u && completionHr == S_OK && ownerGone && stateGone &&
                        primaryPreserved && primaryStatePreserved && foregroundPreserved && sourcePreserved && destinationPreserved;
    AppendLog(std::format(L"Completion owner teardown: passed={} completedWait={} callbacks={} result=0x{:08X} ownerGone={} stateGone={} "
                          L"primaryPreserved={} primaryStatePreserved={} foregroundPreserved={} sourcePreserved={} destinationPreserved={} "
                          L"admission=0x{:08X} sourcePlugin={} destinationPlugin={}",
                          passed,
                          destroyed.conditionMet,
                          callbacks.load(std::memory_order_acquire),
                          static_cast<unsigned long>(completionHr),
                          ownerGone,
                          stateGone,
                          primaryPreserved,
                          primaryStatePreserved,
                          foregroundPreserved,
                          sourcePreserved,
                          destinationPreserved,
                          static_cast<unsigned long>(admissionHr),
                          sourcePlugin,
                          destinationPlugin));
    return passed;
}

[[nodiscard]] bool RecordFileOpsPopupIdleTelemetry(SelfTestState& state) noexcept
{
    using namespace std::chrono_literals;
    using Snapshot    = FileOperationsPopupInternal::PopupTelemetryDebugSnapshot;
    auto* folder      = TryGetFolderWindow(state.mainWindow);
    auto* operations  = TryGetFileOps(folder);
    const auto reject = [&state](std::wstring_view reason) noexcept
    {
        state.retainTempRootOnFailure = true;
        AppendLog(std::format(L"Popup telemetry fixture: {}", reason));
        return false;
    };
    // Only an explicit exact-case request can select the current-source
    // counterfactual. Full/family/prefix selections retain production scheduling.
    constexpr std::wstring_view modeVariable = L"REDSALAMANDER_SELFTEST_F6_IDLE_MODE";
    const auto requestedMode                 = RedSalamander::TestSupport::ReadEnvironmentValue(modeVariable);
    if (requestedMode.error != ERROR_SUCCESS)
        return reject(std::format(L"F6 mode environment read failed error={}", requestedMode.error));
    bool idleSuspensionEnabled = true;
    if (requestedMode.value.has_value())
    {
        if (! EqualsIgnoreCase(state.runFilter, L"FileOps_SafetyRemediationContracts"))
            return reject(L"F6 mode requires explicit FileOps_SafetyRemediationContracts case selection");
        if (requestedMode.value.value() == L"suspension-off")
            idleSuspensionEnabled = false;
        else if (requestedMode.value.value() != L"suspension-on")
            return reject(L"F6 mode must be exactly suspension-on or suspension-off");
    }
    if (! folder || ! operations || operations->HasActiveOperations())
        return reject(L"idle owner unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground  = GetForegroundWindow();
    uint64_t taskId        = 0u;
    uint64_t infoId        = 0u;
    const bool autoDismiss = operations->GetAutoDismissSuccess();
    const auto root        = state.tempRoot / L"popup-telemetry";
    const auto source      = root / L"source" / L"payload.bin";
    const auto destination = root / L"destination" / L"payload.bin";
    const auto reference   = root / L"reference.bin";
    if (! RecreateEmptyDirectory(root) || ! WriteTestFile(source, 8193u) || ! WriteTestFile(reference, 8193u) ||
        ! SelfTest::EnsureDirectory(destination.parent_path()))
        return reject(L"owned byte references unavailable");
    bool workerQuietForCleanup = true;
    const auto removeRoot      = wil::scope_exit([&]() noexcept
    {
        if (workerQuietForCleanup && ! state.retainTempRootOnFailure)
            static_cast<void>(SelfTest::RemoveAll(root));
        else
            AppendLog(std::format(L"Popup telemetry retains owned scratch: {} workerQuiet={}", root.native(), workerQuietForCleanup));
    });
    bool cleanupAttempted      = false;
    bool cleanupSucceeded      = false;
    const auto cleanup         = [&]() noexcept
    {
        if (cleanupAttempted)
            return cleanupSucceeded;
        cleanupAttempted = true;
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
        ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
        SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
        const bool quiet      = CancelOutstandingSelfTestOperations(operations);
        workerQuietForCleanup = quiet;
        if (! quiet)
        {
            state.retainTempRootOnFailure = true;
            // Initial admission proved this owner had no other active work. The
            // terminal failure path joins this fixture's workers before settings
            // restoration; sticky retention stops all subsequent family reuse.
            AppendLog(L"Popup telemetry bounded drain failed; hard-retiring the owner and retaining scratch before settings restoration.");
            operations->Shutdown();
            workerQuietForCleanup = ! operations->HasActiveOperations();
        }
        if (quiet)
        {
            if (taskId != 0u)
                operations->DismissCompletedTask(taskId);
            if (infoId != 0u)
                operations->DismissInformationalTask(infoId);
            operations->DebugResetPopupForSelfTest();
        }
        operations->SetAutoDismissSuccess(autoDismiss);
        const bool saved = operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u);
        cleanupSucceeded = quiet && saved;
        return cleanupSucceeded;
    };
    auto restore = wil::scope_exit([&]() noexcept { static_cast<void>(cleanup()); });
    operations->SetAutoDismissSuccess(false);
    operations->DebugEnsurePopupVisibleForSelfTest();
    const HWND popup = operations->GetPopupHwndForSelfTest();
    if (! popup)
        return reject(L"retained popup unavailable");

    Snapshot originalPolicy{};
    if (! DebugGetFileOperationsPopupTelemetrySnapshot(popup, originalPolicy) || ! originalPolicy.idlePolicyEnabled)
        return reject(L"owned popup did not begin with production idle suspension enabled");
    if (! DebugSetFileOperationsPopupIdleSuspensionForF6(popup, operations, idleSuspensionEnabled))
        return reject(L"owned popup policy scope unavailable");
    auto restoreIdleSuspension = wil::scope_exit([&]() noexcept
    {
        // Reacquire through the owner; never retain a popup-state borrow across a pump.
        if (operations->GetPopupHwndForSelfTest() == popup)
            static_cast<void>(DebugSetFileOperationsPopupIdleSuspensionForF6(popup, operations, true));
    });
    Debug::Perf::Emit(L"FileOps.PopupTelemetry.PolicyMode",
                      idleSuspensionEnabled ? L"current-source-suspension-on" : L"current-source-suspension-off-counterfactual",
                      0u,
                      idleSuspensionEnabled ? 1u : 0u,
                      650u,
                      S_OK);
    AppendLog(std::format(L"Popup telemetry mode: current-source suspensionEnabled={} explicitCase={}",
                          idleSuspensionEnabled,
                          EqualsIgnoreCase(state.runFilter, L"FileOps_SafetyRemediationContracts")));

    // This fixture samples interval deltas alongside telemetry counters; unlike
    // presentation fixtures it observes naturally dispatched timers, never redraws.
    struct ResourceSample final
    {
        uint64_t privateBytes = 0u, workingSet = 0u, uiCpuUs = 0u, processCpuUs = 0u;
        uint64_t handles = 0u, gdi = 0u, user = 0u, hosts = 0u, attachments = 0u;
    };
    const auto captureResources = [&](ResourceSample& value, std::wstring_view name, std::wstring_view boundary) noexcept
    {
        const auto memory = FolderWindowFileOperationsStateInternal::CaptureProcessMemorySnapshot();
        DWORD handles     = 0u;
        FILETIME created{}, exited{}, kernel{}, user{};
        const auto micros = [](FILETIME time) noexcept { return ((static_cast<uint64_t>(time.dwHighDateTime) << 32u) | time.dwLowDateTime) / 10u; };
        if (memory.privateBytes == 0u || memory.workingSetBytes == 0u)
            return reject(
                std::format(L"{} {} CaptureProcessMemorySnapshot returned private={} working={}", name, boundary, memory.privateBytes, memory.workingSetBytes));
        if (! GetProcessHandleCount(GetCurrentProcess(), &handles))
        {
            const DWORD error = GetLastError();
            return reject(std::format(L"{} {} GetProcessHandleCount failed error={}", name, boundary, error));
        }
        if (! GetProcessTimes(GetCurrentProcess(), &created, &exited, &kernel, &user))
        {
            const DWORD error = GetLastError();
            return reject(std::format(L"{} {} GetProcessTimes failed error={}", name, boundary, error));
        }
        value.privateBytes = memory.privateBytes;
        value.workingSet   = memory.workingSetBytes;
        value.processCpuUs = micros(kernel) + micros(user);
        if (! GetThreadTimes(GetCurrentThread(), &created, &exited, &kernel, &user))
        {
            const DWORD error = GetLastError();
            return reject(std::format(L"{} {} GetThreadTimes failed error={}", name, boundary, error));
        }
        value.uiCpuUs     = micros(kernel) + micros(user);
        value.handles     = handles;
        value.gdi         = GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS);
        value.user        = GetGuiResources(GetCurrentProcess(), GR_USEROBJECTS);
        value.hosts       = DxUi::DebugGetAttachedWindowHostCount();
        value.attachments = DxUi::DebugGetSharedWindowHostAttachmentCountForThread(GetCurrentThreadId());
        return true;
    };
    const auto interval = [](std::chrono::milliseconds duration) noexcept
    {
        const auto end = std::chrono::steady_clock::now() + duration;
        return RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept { return std::chrono::steady_clock::now() >= end; },
                                                             {.timeout = SelfTest::Scale(2s), .operationName = L"popup natural telemetry interval"});
    };
    const std::array<std::pair<std::wstring_view, int>, 3> displayStates{
        {{L"visible", SW_SHOWNOACTIVATE}, {L"hidden", SW_HIDE}, {L"minimized", SW_SHOWMINNOACTIVE}}};
    const auto measure = [&](std::wstring_view phase, bool work) noexcept
    {
        for (const auto& [display, show] : displayStates)
        {
            const auto name = std::format(L"{}-{}", phase, display);
            ShowWindow(popup, show);
            const auto warmed = interval(work ? 150ms : 700ms);
            if (! warmed.conditionMet)
                return reject(std::format(L"{} warm-up pump did not meet its condition elapsedMs={}", name, warmed.elapsed.count()));
            if (! work)
            {
                Snapshot settling{};
                if (! DebugGetFileOperationsPopupTelemetrySnapshot(popup, settling))
                    return reject(std::format(L"{} idle scheduling observation unavailable", name));
                // Observe actual motion independently of cadence in both modes.
                // The query reads animation state only; it builds no task/rate snapshot.
                bool visibleMotion = true;
                const auto settled = RedSalamander::TestSupport::PumpMessagesUntil(
                    [&]() noexcept
                {
                    return DebugGetFileOperationsPopupUiMotionForF6(popup, operations, visibleMotion) && ! visibleMotion &&
                           DebugGetFileOperationsPopupTelemetrySnapshot(popup, settling) && settling.timerIntervalMs == (idleSuspensionEnabled ? 0u : 100u);
                },
                    {.timeout = SelfTest::Scale(2s), .operationName = L"bounded completed popup motion quiet transition"});
                if (! DebugGetFileOperationsPopupTelemetrySnapshot(popup, settling))
                    return reject(std::format(L"{} settled scalar observation unavailable", name));
                Debug::Perf::Emit(L"FileOps.PopupTelemetry.IdleMotionSettlementUs",
                                  name,
                                  static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(settled.elapsed).count()),
                                  settling.timerIntervalMs,
                                  settled.conditionMet ? 1u : 0u,
                                  settled.conditionMet ? S_OK : HRESULT_FROM_WIN32(ERROR_TIMEOUT));
                if (! settled.conditionMet)
                    return reject(std::format(L"{} motion did not settle elapsedMs={} visibleMotion={} intervalMs={} suspensionEnabled={}",
                                              name,
                                              settled.elapsed.count(),
                                              visibleMotion,
                                              settling.timerIntervalMs,
                                              settling.idlePolicyEnabled));
            }
            if ((display == L"visible" && (! IsWindowVisible(popup) || IsIconic(popup))) || (display == L"hidden" && IsWindowVisible(popup)) ||
                (display == L"minimized" && ! IsIconic(popup)))
                return reject(std::format(L"{} did not establish its actual display state", name));
            ResourceSample resourcesBefore{}, resourcesAfter{};
            Snapshot before{}, after{};
            if (! captureResources(resourcesBefore, name, L"before"))
                return false; // The exact resource API already logged its failure.
            if (! DebugGetFileOperationsPopupTelemetrySnapshot(popup, before))
                return reject(std::format(L"{} before scalar observation unavailable popup={} currentPopup={} isWindow={} thread={}",
                                          name,
                                          reinterpret_cast<uintptr_t>(popup),
                                          reinterpret_cast<uintptr_t>(operations->GetPopupHwndForSelfTest()),
                                          IsWindow(popup),
                                          GetWindowThreadProcessId(popup, nullptr)));
            const auto elapsed = interval(650ms);
            if (! elapsed.conditionMet)
                return reject(std::format(L"{} measured pump did not meet its condition elapsedMs={}", name, elapsed.elapsed.count()));
            if (! DebugGetFileOperationsPopupTelemetrySnapshot(popup, after))
                return reject(std::format(L"{} after scalar observation unavailable popup={} currentPopup={} isWindow={} thread={}",
                                          name,
                                          reinterpret_cast<uintptr_t>(popup),
                                          reinterpret_cast<uintptr_t>(operations->GetPopupHwndForSelfTest()),
                                          IsWindow(popup),
                                          GetWindowThreadProcessId(popup, nullptr)));
            if (! captureResources(resourcesAfter, name, L"after"))
                return false; // The exact resource API already logged its failure.
            const HWND actualForeground = GetForegroundWindow();
            if (actualForeground != foreground)
            {
                DWORD expectedProcess = 0u, actualProcess = 0u;
                const DWORD expectedThread = GetWindowThreadProcessId(foreground, &expectedProcess);
                const DWORD actualThread   = GetWindowThreadProcessId(actualForeground, &actualProcess);
                return reject(std::format(
                    L"{} foreground changed expectedHwnd={} expectedPid={} expectedTid={} actualHwnd={} actualPid={} actualTid={} active={} focus={}",
                    name,
                    reinterpret_cast<uintptr_t>(foreground),
                    expectedProcess,
                    expectedThread,
                    reinterpret_cast<uintptr_t>(actualForeground),
                    actualProcess,
                    actualThread,
                    reinterpret_cast<uintptr_t>(GetActiveWindow()),
                    reinterpret_cast<uintptr_t>(GetFocus())));
            }
            const auto ticks     = after.timerTicks - before.timerTicks;
            const auto snapshots = after.taskSnapshots - before.taskSnapshots;
            const auto rates     = after.rateSnapshots - before.rateSnapshots;
            const auto taskbar   = after.taskbarAttempts - before.taskbarAttempts;
            AppendLog(std::format(L"Popup telemetry {}: elapsedMs={} timerTicks={} taskSnapshots={} rateSnapshots={} rateUpdates={} "
                                  L"taskbarAttempts={} taskbarSuccesses={} intervalMs={} policy={} lastTaskbarState={}",
                                  name,
                                  elapsed.elapsed.count(),
                                  ticks,
                                  snapshots,
                                  rates,
                                  after.rateUpdates - before.rateUpdates,
                                  taskbar,
                                  after.taskbarSuccesses - before.taskbarSuccesses,
                                  after.timerIntervalMs,
                                  after.idlePolicyEnabled,
                                  after.lastTaskbarState));
            const uint64_t elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed.elapsed).count());
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.Counters", name, elapsedUs, ticks, snapshots, S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.Rates", name, 0u, rates, after.rateUpdates - before.rateUpdates, S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.Taskbar", name, 0u, taskbar, after.taskbarSuccesses - before.taskbarSuccesses, S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.CpuUs",
                              name,
                              elapsedUs,
                              resourcesAfter.uiCpuUs - resourcesBefore.uiCpuUs,
                              resourcesAfter.processCpuUs - resourcesBefore.processCpuUs,
                              S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.MemoryBefore", name, 0u, resourcesBefore.privateBytes, resourcesBefore.workingSet, S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.MemoryAfter", name, 0u, resourcesAfter.privateBytes, resourcesAfter.workingSet, S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.Handles", name, 0u, resourcesBefore.handles, resourcesAfter.handles, S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.Hosts", name, 0u, resourcesBefore.hosts, resourcesAfter.hosts, S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.Attachments", name, 0u, resourcesBefore.attachments, resourcesAfter.attachments, S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.GuiBefore", name, 0u, resourcesBefore.gdi, resourcesBefore.user, S_OK);
            Debug::Perf::Emit(L"FileOps.PopupTelemetry.GuiAfter", name, 0u, resourcesAfter.gdi, resourcesAfter.user, S_OK);
            if (work && (ticks == 0u || taskbar == 0u || after.timerIntervalMs == 0u))
                return reject(std::format(L"{} stopped active telemetry", name));
            if (before.idlePolicyEnabled != idleSuspensionEnabled || after.idlePolicyEnabled != idleSuspensionEnabled)
                return reject(std::format(L"{} lost its scoped policy mode", name));
            if (! work && idleSuspensionEnabled &&
                (ticks != 0u || snapshots != 0u || rates != 0u || after.rateUpdates != before.rateUpdates || taskbar != 0u || after.timerIntervalMs != 0u))
                return reject(std::format(L"{} retained idle telemetry", name));
            if (! work && ! idleSuspensionEnabled && (ticks == 0u || snapshots == 0u || taskbar == 0u || after.timerIntervalMs != 100u))
                return reject(std::format(L"{} counterfactual did not retain real idle ticks/snapshots/taskbar submissions at 100ms", name));
            if (display != L"visible" && rates != 0u)
                return reject(std::format(L"{} built hidden/minimized graph rates", name));
        }
        return true;
    };
    if (! measure(L"initial-idle", false))
        return false;
    // Hold a real publication and then its real terminal completion. The bounded
    // worker barrier contributes process CPU; only paired identical runs compare it.
    SetFileOpsLiveOutputPublishedPauseForSelfTest(true);
    SetFileOpsPostFinishedCompletionPauseForSelfTest(true);
    workerQuietForCleanup = false;
    const auto admitted   = StartFileOperationAndGetId(operations,
                                                       FILESYSTEM_COPY,
                                                       FolderWindow::Pane::Left,
                                                       FolderWindow::Pane::Right,
                                                       folder->GetFileSystem(FolderWindow::Pane::Left),
                                                       {source},
                                                       destination.parent_path(),
                                                       FILESYSTEM_FLAG_NONE,
                                                       false,
                                                       0u,
                                                       FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                       false,
                                                       folder->GetFileSystem(FolderWindow::Pane::Right));
    if (! admitted.has_value())
        return reject(L"real Copy admission failed");
    taskId             = admitted.value();
    const auto running = RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept { return HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest(); },
                                                                       {.timeout = SelfTest::Scale(8s), .operationName = L"telemetry actual publication gate"});
    auto* task         = operations->FindTask(taskId);
    if (! running.conditionMet || ! task || ! task->HasEnteredOperation() || task->_taskFinished.load(std::memory_order_acquire))
        return reject(L"actual running Copy did not reach its publication gate");
    task->_presentationState.store(FolderWindow::FileOperationState::Task::TaskPresentationState::Presented, std::memory_order_release);
    if (! measure(L"running", true) || ! HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest())
        return reject(L"running measurement outlived its real publication gate");
    ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
    const auto completing = RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept {
        return HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest();
    }, {.timeout = SelfTest::Scale(8s), .operationName = L"telemetry actual completion gate"});
    task                  = operations->FindTask(taskId);
    if (! completing.conditionMet || ! task || ! task->_taskFinished.load(std::memory_order_acquire) || ! measure(L"completing", true) ||
        ! HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest())
        return reject(L"completing measurement lost its owned terminal worker");
    ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
    const auto quiet = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return ! operations->FindTask(taskId) && ! operations->HasActiveOperations();
    }, {.timeout = SelfTest::Scale(8s), .operationName = L"telemetry actual worker quiet point"});
    if (! quiet.conditionMet || operations->GetPopupHwndForSelfTest() != popup || ! FilesEqualBytes(source, reference) ||
        ! FilesEqualBytes(destination, reference) || ! measure(L"completed-idle", false))
        return reject(L"completion lost retained popup, exact bytes, or idle measurement");
    Snapshot final{};
    if (! DebugGetFileOperationsPopupTelemetrySnapshot(popup, final) || final.lastTaskbarState != static_cast<uint32_t>(TBPF_NOPROGRESS))
        return reject(L"last completion did not publish the final taskbar model");
    FolderWindow::InformationalTaskUpdate information{};
    information.title      = L"Vérification informative de la suspension de la télémétrie";
    information.scanActive = true;
    infoId                 = operations->CreateOrUpdateInformationalTask(information);
    if (infoId == 0u || ! measure(L"informational-running", true))
        return reject(L"informational admission did not rearm telemetry");
    information.taskId     = infoId;
    information.scanActive = false;
    information.finished   = true;
    information.resultHr   = S_OK;
    static_cast<void>(operations->CreateOrUpdateInformationalTask(information));
    if (! measure(L"informational-completed-idle", false))
        return false;
    // Outside the unchanged measurement intervals, replay Explorer's taskbar
    // event on the stopped, minimized popup. It must use its cached final model.
    const UINT taskbarCreated = RegisterWindowMessageW(L"TaskbarButtonCreated");
    Snapshot replay{};
    if (taskbarCreated == 0u || ! DebugGetFileOperationsPopupTelemetrySnapshot(popup, final))
        return reject(L"Explorer taskbar-button event or final scalar observation unavailable");
    SendMessageW(popup, taskbarCreated, 0u, 0);
    if (! DebugGetFileOperationsPopupTelemetrySnapshot(popup, replay) || replay.taskSnapshots != final.taskSnapshots ||
        replay.lastTaskbarState != static_cast<uint32_t>(TBPF_NOPROGRESS) || replay.taskbarAttempts != final.taskbarAttempts + 1u ||
        (replay.idlePolicyEnabled && replay.timerIntervalMs != 0u) || GetForegroundWindow() != foreground)
        return reject(L"Explorer did not replay the final taskbar model from its cache without idle telemetry");
    AppendLog(std::format(L"Popup taskbar replay: snapshots={} attempts={} successes={} intervalMs={}",
                          replay.taskSnapshots - final.taskSnapshots,
                          replay.taskbarAttempts - final.taskbarAttempts,
                          replay.taskbarSuccesses - final.taskbarSuccesses,
                          replay.timerIntervalMs));
    if (! DebugSetFileOperationsPopupIdleSuspensionForF6(popup, operations, true))
        return reject(L"owned popup policy restoration failed");
    restoreIdleSuspension.release();
    if (! cleanup())
        return reject(L"worker cleanup or restored settings save did not reach its quiet point");
    restore.release();
    const HWND finalForeground = GetForegroundWindow();
    if (finalForeground != foreground)
        return reject(std::format(L"final cleanup foreground changed expectedHwnd={} actualHwnd={}",
                                  reinterpret_cast<uintptr_t>(foreground),
                                  reinterpret_cast<uintptr_t>(finalForeground)));
    return true;
}

// I30 Wave 3 item 3 (x-perf#3, x-perf#5, x-lifetime#3, x-lifetime#7, model#3): the popup and the
// owner heartbeat stop doing per-tick work that changes nothing on screen.
[[nodiscard]] bool RunFileOpsCompletedCardAutoCollapseChecks(SelfTestState& state) noexcept
{
    using namespace std::chrono_literals;
    using namespace FileOperationsPopupInternal;
    const auto reject = [&state](std::wstring_view reason) noexcept
    {
        state.retainTempRootOnFailure = true;
        AppendLog(std::format(L"Completed-card auto-collapse: FAILED: {}", reason));
        return false;
    };
    auto* folder = TryGetFolderWindow(state.mainWindow);
    if (! folder)
        return reject(L"owner unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground   = GetForegroundWindow();
    DWORD foregroundProcess = 0u;
    static_cast<void>(GetWindowThreadProcessId(foreground, &foregroundProcess));
    // Retaining height is an explicit same-binary counterfactual of the old completion path.
    // Fixed geometry prevents window auto-sizing or deferred accessible text from masking the race.
    struct Scenario
    {
        bool snap;
        bool unknown;
        bool reducedMotion;
    };
    constexpr std::array scenarios{Scenario{false, false, false}, Scenario{true, false, false}, Scenario{true, true, false}, Scenario{true, false, true}};
    for (const auto& scenario : scenarios)
    {
        Common::Settings::Settings settings{};
        settings.fileOperations.emplace();
        auto owner = std::make_unique<FolderWindow>();
        owner->SetSettings(&settings);
        AppTheme theme              = folder->GetTheme();
        theme.reducedMotionOverride = scenario.reducedMotion;
        owner->ApplyTheme(theme);
        FolderWindow::FileOperationState privateOperations(*owner);
        auto lifetime = std::make_shared<int>(0);
        wil::unique_hwnd popup(FileOperationsPopup::Create(&privateOperations, owner.get(), state.mainWindow, lifetime));
        auto closePopup  = wil::scope_exit([&]() noexcept
        {
            lifetime.reset();
            popup.reset();
        });
        auto* popupState = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA)) : nullptr;
        if (! popupState)
            return reject(L"private popup unavailable");
        popupState->debugFixedGeometry           = true;
        popupState->debugAutoCollapseSnapEnabled = scenario.snap;
        constexpr uint64_t taskId                = 0xF260202610180001ull;
        TaskSnapshot task{};
        task.taskId                 = taskId;
        task.operation              = scenario.unknown ? FILESYSTEM_MOVE : FILESYSTEM_COPY;
        task.lifecyclePhase         = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
        task.started                = true;
        task.paused                 = true;
        task.discoveryClosed        = true;
        task.discoveredBytesKnown   = true;
        task.hasProgressCallbacks   = true;
        task.operationStartTick     = GetTickCount64();
        task.currentSourcePath      = L"C:\\CollapseFixture\\album";
        task.currentDestinationPath = L"D:\\CollapseFixture\\album";
        task.totalBytes             = 32ull << 20;
        task.completedBytes         = task.totalBytes;
        task.totalItems             = 9u;
        task.completedItems         = task.totalItems;
        // Preserve the actual unset collapse choice and height history across the live transition.
        popupState->DebugUpdateLiveVisualScenario({task});
        SetWindowPos(popup.get(), nullptr, 40, 40, 860, 700, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
        const auto publicationReady = std::chrono::steady_clock::now() + 350ms;
        if (! RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept { return std::chrono::steady_clock::now() >= publicationReady; },
                                                            {.timeout = SelfTest::Scale(2s), .operationName = L"settle initial card text"})
                  .conditionMet)
            return reject(L"initial card did not settle");
        task.paused                 = false;
        task.finished               = true;
        task.lifecyclePhase         = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Terminal);
        task.completedBytes         = task.totalBytes;
        task.completedItems         = task.totalItems;
        task.resultHr               = scenario.unknown ? HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) : S_OK;
        task.hasIndeterminateResult = scenario.unknown;
        popupState->DebugUpdateLiveVisualScenario({task});
        const auto started = std::chrono::steady_clock::now();
        popupState->DebugRunIdleTimerTickForSelfTest(popup.get());
        UpdateWindow(popup.get());
        const UINT firstTimer = popupState->DebugTelemetrySnapshot().timerIntervalMs;
        PopupLayoutDebugSnapshot first{};
        first.taskId = taskId;
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), first) || ! first.taskCompactRow)
            return reject(L"completion did not logically collapse the card");
        const float firstHeight = first.taskPaintedCardRect.bottom - first.taskPaintedCardRect.top;
        // Pump only: a forced paint here would hide the missing-animation-wake defect.
        const auto observeUntil = std::chrono::steady_clock::now() + 900ms;
        static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return std::chrono::steady_clock::now() >= observeUntil;
        }, {.timeout = SelfTest::Scale(2s), .operationName = L"timer-driven completed-card settlement"}));
        PopupLayoutDebugSnapshot settled{};
        settled.taskId = taskId;
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), settled))
            return reject(L"settled card unavailable");
        const float height             = settled.taskPaintedCardRect.bottom - settled.taskPaintedCardRect.top;
        const float compactHeight      = 44.0f * static_cast<float>(GetDpiForWindow(popup.get())) / 96.0f;
        const bool firstRetainedHeight = first.taskHeightAnimationUnsettled && firstHeight > compactHeight + 2.0f;
        const bool quiet               = popupState->DebugTelemetrySnapshot().timerIntervalMs == 0u;
        const auto expectedStatus      = scenario.unknown ? TaskSnapshot::StatusKind::Indeterminate : TaskSnapshot::StatusKind::Done;
        const bool correct             = settled.taskStatusKind == expectedStatus && quiet &&
                                         (! scenario.snap ? firstRetainedHeight
                                                          : ! first.taskHeightAnimationUnsettled && std::abs(firstHeight - compactHeight) <= 2.0f &&
                                                                ! settled.taskHeightAnimationUnsettled && std::abs(height - compactHeight) <= 2.0f);
        const uint64_t elapsedUs       = Debug::Perf::ElapsedUs(started);
        const std::wstring facts =
            std::format(L"snap={};unknown={};reduced={};firstTimer={};firstHeight={:.1f};height={:.1f};compact={:.1f};firstRetained={};quiet={}",
                        scenario.snap,
                        scenario.unknown,
                        scenario.reducedMotion,
                        firstTimer,
                        firstHeight,
                        height,
                        compactHeight,
                        firstRetainedHeight,
                        quiet);
        AppendLog(std::format(L"Completed-card auto-collapse: {} {}", correct ? L"passed" : L"FAILED", facts));
        Debug::Perf::Emit(L"FileOps.Popup.AutoCollapse.ObservationUs", facts, elapsedUs, firstTimer, static_cast<uint64_t>(height), correct ? S_OK : E_FAIL);
        if (! correct)
            return reject(facts);
    }
    const HWND finalForeground = GetForegroundWindow();
    if (finalForeground == foreground)
        return true;
    DWORD finalForegroundProcess = 0u;
    static_cast<void>(GetWindowThreadProcessId(finalForeground, &finalForegroundProcess));
    return reject(std::format(L"foreground changed: before=0x{:X} beforePid={} after=0x{:X} afterPid={} selfPid={}",
                              reinterpret_cast<uintptr_t>(foreground),
                              foregroundProcess,
                              reinterpret_cast<uintptr_t>(finalForeground),
                              finalForegroundProcess,
                              GetCurrentProcessId()));
}

[[nodiscard]] bool RunFileOpsPopupIdleCadenceChecks(SelfTestState& state) noexcept
{
    using namespace std::chrono_literals;
    using namespace FileOperationsPopupInternal;
    using State       = FolderWindow::FileOperationState;
    const auto reject = [&state](std::wstring_view reason) noexcept
    {
        state.retainTempRootOnFailure = true;
        AppendLog(std::format(L"Popup idle cadence: FAILED: {}", reason));
        return false;
    };
    const auto note = [](std::wstring_view text) noexcept { AppendLog(std::format(L"Popup idle cadence: {}", text)); };
    // Cadence assertions record and continue, so one run measures every state (also against a
    // counterfactual build); fixture failures still stop at once through reject.
    size_t failures = 0u;
    const auto fail = [&](std::wstring_view reason) noexcept
    {
        ++failures;
        state.retainTempRootOnFailure = true;
        AppendLog(std::format(L"Popup idle cadence: FAILED: {}", reason));
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || operations->HasActiveOperations())
        return reject(L"idle owner unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground = GetForegroundWindow();
    const auto pumpFor    = [](std::chrono::milliseconds duration) noexcept
    {
        const auto end = std::chrono::steady_clock::now() + duration;
        return RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept { return std::chrono::steady_clock::now() >= end; },
                                                             {.timeout = duration + SelfTest::Scale(2s), .operationName = L"popup cadence interval"});
    };
    const auto uiThreadCpuUs = []() noexcept -> uint64_t
    {
        FILETIME created{}, exited{}, kernel{}, user{};
        if (! GetThreadTimes(GetCurrentThread(), &created, &exited, &kernel, &user))
            return 0u;
        const auto micros = [](FILETIME time) noexcept { return ((static_cast<uint64_t>(time.dwHighDateTime) << 32u) | time.dwLowDateTime) / 10u; };
        return micros(kernel) + micros(user);
    };

    // 1. Synthetic popup on a private engine: the cadence each declared task state selects, and the
    //    path-detail motion check (x-lifetime#3). Nothing here admits work or touches the user's popup.
    {
        Common::Settings::Settings settings{};
        settings.fileOperations.emplace();
        auto owner = std::make_unique<FolderWindow>();
        owner->SetSettings(&settings);
        State privateOperations(*owner);
        auto lifetime = std::make_shared<int>(0);
        wil::unique_hwnd popup(FileOperationsPopup::Create(&privateOperations, folder, state.mainWindow, lifetime));
        auto closePopup  = wil::scope_exit([&]() noexcept
        {
            lifetime.reset();
            popup.reset();
        });
        auto* popupState = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA)) : nullptr;
        if (! popupState)
            return reject(L"could not create the private production popup");
        const auto paint = [&]() noexcept { RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN); };

        constexpr uint64_t taskId = 0xF26A202610090331ull;
        TaskSnapshot running{};
        running.taskId                     = taskId;
        running.operation                  = FILESYSTEM_COPY;
        running.lifecyclePhase             = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
        running.started                    = true;
        running.discoveryClosed            = true;
        running.hasProgressCallbacks       = true;
        running.lastProgressCallbackTick   = GetTickCount64();
        running.operationStartTick         = running.lastProgressCallbackTick > 250u ? running.lastProgressCallbackTick - 250u : 0u;
        running.currentSourcePath          = L"C:\\Cadence\\Source\\Æble-שלום\\payload.bin";
        running.currentDestinationPath     = L"D:\\Cadence\\Destination\\Æble-שלום\\payload.bin";
        running.totalBytes                 = 64ull << 20;
        running.completedBytes             = 16ull << 20;
        running.discoveredBytesKnown       = true;
        running.exactDiscoveredBytes       = running.totalBytes;
        running.totalItems                 = 2u;
        running.completedItems             = 1u;
        TaskSnapshot paused                = running;
        paused.paused                      = true;
        TaskSnapshot pausedUnknown         = paused;
        pausedUnknown.discoveredBytesKnown = false;
        pausedUnknown.exactDiscoveredBytes = 0u;
        TaskSnapshot prompting             = running;
        prompting.conflict.active          = true;
        prompting.conflict.generation      = 7u;
        prompting.conflict.sourcePath      = running.currentSourcePath;
        prompting.conflict.destinationPath = running.currentDestinationPath;
        TaskSnapshot queued                = running;
        queued.started                     = false;
        queued.operationStartTick          = 0u;
        queued.hasProgressCallbacks        = false;
        queued.waitingInQueue              = true;
        TaskSnapshot finished              = running;
        finished.finished                  = true;

        ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
        struct Scenario final
        {
            std::wstring_view name;
            const TaskSnapshot* task = nullptr;
            UINT intervalMs          = 0u;
        };
        const std::array<Scenario, 6> scenarios{{
            {L"synthetic-running", &running, 100u},
            {L"synthetic-paused", &paused, 1000u},
            {L"synthetic-paused-unknown", &pausedUnknown, 1000u},
            {L"synthetic-prompting", &prompting, 1000u},
            {L"synthetic-queued", &queued, 1000u},
            {L"synthetic-completed-idle", &finished, 0u},
        }};
        for (const Scenario& scenario : scenarios)
        {
            popupState->DebugSetVisualScenario({*scenario.task}, false, true);
            paint();
            PopupTelemetryDebugSnapshot settled{};
            const auto reached = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                settled = popupState->DebugTelemetrySnapshot();
                return settled.timerIntervalMs == scenario.intervalMs && ! popupState->DebugHasVisibleUiMotion(popup.get());
            },
                {.timeout = SelfTest::Scale(3s), .operationName = L"synthetic popup cadence"});
            if (! reached.conditionMet)
                fail(std::format(L"{} kept intervalMs={} (expected {}) visibleMotion={}",
                                 scenario.name,
                                 settled.timerIntervalMs,
                                 scenario.intervalMs,
                                 popupState->DebugHasVisibleUiMotion(popup.get())));
            // A scenario change lets DxUi finish its own bounded control transitions; measure the steady state.
            static_cast<void>(pumpFor(600ms));
            if (scenario.task == &pausedUnknown)
            {
                PopupLayoutDebugSnapshot unknownBar{};
                unknownBar.taskId = taskId;
                if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), unknownBar) || ! unknownBar.taskHostedProgressPresent ||
                    ! unknownBar.taskHostedProgressIndeterminate || ! unknownBar.taskHostedProgressAnimationPaused || ! unknownBar.taskHostedProgressEnabled)
                    fail(L"paused unknown progress must keep its enabled indeterminate control with motion suspended");
            }
            const PopupTelemetryDebugSnapshot before = popupState->DebugTelemetrySnapshot();
            const auto elapsed                       = pumpFor(1200ms);
            const PopupTelemetryDebugSnapshot after  = popupState->DebugTelemetrySnapshot();
            const uint64_t ticks                     = after.timerTicks - before.timerTicks;
            const uint64_t renders                   = after.taskSnapshots - before.taskSnapshots;
            const uint64_t elapsedUs                 = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed.elapsed).count());
            Debug::Perf::Emit(L"FileOps.PopupCadence.Counters", scenario.name, elapsedUs, ticks, renders, S_OK);
            note(std::format(
                L"{}: elapsedMs={} timerTicks={} renders={} intervalMs={}", scenario.name, elapsed.elapsed.count(), ticks, renders, after.timerIntervalMs));
            if (! elapsed.conditionMet || after.timerIntervalMs != scenario.intervalMs)
                fail(std::format(L"{} changed its cadence while measured intervalMs={}", scenario.name, after.timerIntervalMs));
            if (scenario.intervalMs == 100u && ticks < 6u)
                fail(std::format(L"{} stopped active telemetry: timerTicks={}", scenario.name, ticks));
            if (scenario.intervalMs == 1000u && (ticks > 2u || renders > 3u))
                fail(std::format(L"{} kept repainting a waiting task: timerTicks={} renders={}", scenario.name, ticks, renders));
            if (scenario.intervalMs == 0u && ticks != 0u)
                fail(std::format(L"{} kept an idle timer: timerTicks={}", scenario.name, ticks));
        }

        // A finished card's bar after the idle transition. The last tick decides idle before any frame shows the
        // task finished; the frame it requests starts the bar's easing toward 100% and must not be the last frame
        // (WM_PAINT refreshes the timer after each frame, which keeps it running while the bar eases).
        {
            TaskSnapshot halfway   = running;
            halfway.completedBytes = halfway.totalBytes / 2u;
            popupState->DebugSetVisualScenario({halfway}, true, true);
            paint();
            PopupTelemetryDebugSnapshot halfwaySettled{};
            static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                halfwaySettled = popupState->DebugTelemetrySnapshot();
                return halfwaySettled.timerIntervalMs == 100u && ! popupState->DebugHasVisibleUiMotion(popup.get());
            },
                {.timeout = SelfTest::Scale(3s), .operationName = L"half-way bar settles"}));
            paint(); // a frame after the easing ended shows its final value
            PopupLayoutDebugSnapshot bar{};
            bar.taskId = taskId;
            if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), bar) || ! bar.taskHostedProgressPresent ||
                std::abs(bar.taskHostedProgressValue - 50.0) > 0.5)
                return reject(std::format(
                    L"the half-way card's bar is not at 50% (present={} value={:.1f})", bar.taskHostedProgressPresent, bar.taskHostedProgressValue));
            TaskSnapshot done = halfway;
            done.finished     = true;
            done.resultHr     = S_OK;
            popupState->DebugSetVisualScenario({done}, true, true);
            // Production order: the tick turns the popup idle first, then the frame it requested is painted.
            popupState->DebugRunIdleTimerTickForSelfTest(popup.get());
            paint();
            // The frame that started the bar's easing must leave the timer running: nothing else is guaranteed to
            // paint again (DxUi's own control transitions may, which hides the defect in this fixture).
            const PopupTelemetryDebugSnapshot afterFinishedFrame = popupState->DebugTelemetrySnapshot();
            if (afterFinishedFrame.timerIntervalMs == 0u)
                fail(L"the frame that started a finished card's bar easing left no timer running, so its bar can stop part-way beside its 100%");
            const auto finishedBar = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                bar        = {};
                bar.taskId = taskId;
                return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), bar) && bar.taskHostedProgressPresent && bar.taskHostedProgressValue >= 99.99;
            },
                {.timeout = SelfTest::Scale(2s), .operationName = L"finished card's bar reaches 100%"});
            note(std::format(L"finished-after-idle: barPresent={} barValue={:.1f}", bar.taskHostedProgressPresent, bar.taskHostedProgressValue));
            if (! finishedBar.conditionMet)
                fail(std::format(L"a finished card's bar stopped at {:.1f}% after the idle transition (its percentage reads 100%)",
                                 bar.taskHostedProgressValue));
        }

        // I35 (owner report 2026-10-13): a finished Verify On card says whether every byte was verified. The
        // verification half of its bar carries visible marks at any share, the text beside its compact meter reads
        // 100% only when every byte was verified (otherwise "Not verified" or the share verified), and the bar's
        // value agrees with what is drawn.
        {
            struct VerifyCase final
            {
                TaskSnapshot task;
                std::wstring label;
                bool marks          = false;
                double marksPercent = 0.0;
                double value        = 0.0;
            };
            TaskSnapshot done               = running;
            done.lifecyclePhase             = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Terminal);
            done.finished                   = true;
            done.resultHr                   = S_OK;
            done.completedBytes             = done.totalBytes;
            done.completedItems             = done.totalItems;
            done.verificationRequested      = true;
            done.verificationState          = static_cast<uint8_t>(FileOperations::VerificationState::Verified);
            done.verificationTotalBytes     = done.totalBytes;
            done.verificationCompletedBytes = done.totalBytes;
            std::array<VerifyCase, 4> cases{};
            cases[0]                                 = {done, FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_PERCENT, 100u), true, 100.0, 100.0};
            cases[0].task.taskId                     = taskId + 10u;
            cases[1]                                 = cases[0];
            cases[1].task.taskId                     = taskId + 11u;
            cases[1].task.verificationState          = static_cast<uint8_t>(FileOperations::VerificationState::Unavailable);
            cases[1].task.verificationCompletedBytes = 0u;
            cases[1].label                           = LoadStringResource(nullptr, IDS_FILEOPS_COMPACT_NOT_VERIFIED);
            cases[1].marksPercent                    = 0.0;
            cases[1].value                           = 50.0;
            cases[2]                                 = cases[1];
            cases[2].task.taskId                     = taskId + 12u;
            // Exactly 40%: the label rounds the share down, so it never claims more than was verified.
            cases[2].task.verificationTotalBytes     = 1000u;
            cases[2].task.verificationCompletedBytes = 400u;
            cases[2].label                           = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_VERIFIED_PERCENT, 40u);
            cases[2].marksPercent                    = 40.0;
            cases[2].value                           = 70.0;
            cases[3]                                 = cases[0];
            cases[3].task.taskId                     = taskId + 13u;
            cases[3].task.verificationRequested      = false;
            cases[3].task.verificationState          = static_cast<uint8_t>(FileOperations::VerificationState::NotRequested);
            cases[3].task.verificationTotalBytes     = 0u;
            cases[3].task.verificationCompletedBytes = 0u;
            cases[3].marks                           = false;
            std::vector<TaskSnapshot> verifyTasks;
            for (const VerifyCase& entry : cases)
                verifyTasks.push_back(entry.task);
            popupState->DebugSetVisualScenario(std::move(verifyTasks), true, true);
            paint();
            std::wstring mismatch;
            const auto agrees = [&]() noexcept
            {
                mismatch.clear();
                for (const VerifyCase& expected : cases)
                {
                    PopupLayoutDebugSnapshot card{};
                    card.taskId = expected.task.taskId;
                    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), card) || ! card.taskHostedProgressPresent ||
                        card.taskCompactMeterLabel != expected.label || card.taskVerificationMarksVisible != expected.marks ||
                        (expected.marks && ! card.taskVerificationMarksOutlined) ||
                        (expected.marks && std::abs(card.taskVerificationMarksPercent - expected.marksPercent) > 0.5) ||
                        std::abs(card.taskHostedProgressValue - expected.value) > 0.5)
                    {
                        mismatch = std::format(L"task +{}: label='{}' (expected '{}'), marks={} {:.1f}% (expected {} {:.1f}%), value={:.1f} (expected {:.1f})",
                                               expected.task.taskId - taskId,
                                               card.taskCompactMeterLabel,
                                               expected.label,
                                               card.taskVerificationMarksVisible,
                                               card.taskVerificationMarksPercent,
                                               expected.marks,
                                               expected.marksPercent,
                                               card.taskHostedProgressValue,
                                               expected.value);
                        return false;
                    }
                }
                return true;
            };
            const auto verifyCards = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                paint();
                return agrees();
            },
                {.timeout = SelfTest::Scale(2s), .operationName = L"finished Verify On cards"});
            note(std::format(L"finished Verify On cards: {}", verifyCards.conditionMet ? std::wstring(L"agree") : mismatch));
            if (! verifyCards.conditionMet)
                fail(std::format(L"a finished Verify On card does not say whether every byte was verified: {}", mismatch));
        }

        // I35 (owner report 2026-10-13): a card whose height motion got no frame after its deadline (every tick of a busy
        // UI thread came late) still settles. The tick that stops the timer asks for the frame that resolves the motion;
        // nothing else is guaranteed to paint again.
        {
            popupState->DebugSetIdleSuspensionEnabled(popup.get(), true);
            TaskSnapshot settledCard   = running;
            settledCard.taskId         = taskId + 20u;
            settledCard.lifecyclePhase = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Terminal);
            settledCard.finished       = true;
            settledCard.resultHr       = S_OK;
            settledCard.completedBytes = settledCard.totalBytes;
            settledCard.completedItems = settledCard.totalItems;
            popupState->DebugSetVisualScenario({settledCard}, false, true);
            paint();
            static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                paint();
                return ! popupState->DebugHasVisibleUiMotion(popup.get());
            },
                {.timeout = SelfTest::Scale(3s), .operationName = L"card settles before the late tick"}));
            popupState->DebugAgeTaskHeightMotionForSelfTest(settledCard.taskId, 120.0, 1000u);
            PopupLayoutDebugSnapshot stale{};
            stale.taskId         = settledCard.taskId;
            const bool staleRead = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), stale);
            popupState->DebugRunIdleTimerTickForSelfTest(popup.get());
            RECT pendingUpdate{};
            const bool frameRequested = GetUpdateRect(popup.get(), &pendingUpdate, FALSE) != FALSE;
            UpdateWindow(popup.get());
            PopupLayoutDebugSnapshot settled{};
            settled.taskId         = settledCard.taskId;
            const bool settledRead = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), settled);
            note(std::format(L"late tick: stale={} frameRequested={} unsettledAfter={}",
                             stale.taskHeightAnimationUnsettled,
                             frameRequested,
                             settled.taskHeightAnimationUnsettled));
            if (! staleRead || ! stale.taskHeightAnimationUnsettled || ! settledRead)
                fail(L"the late-tick fixture could not leave a card height away from its target");
            else if (! frameRequested || settled.taskHeightAnimationUnsettled)
                fail(L"the timer tick after a card's height motion ended without the frame that settles it, so the card keeps a stale height");
        }

        // x-lifetime#3: a resize left pending when path detail opens must not hold the 16 ms
        // motion cadence for as long as the detail stays open.
        popupState->DebugSetVisualScenario({paused}, false, true);
        paint();
        popupState->DebugOpenPathDetail(std::format(L"FileOperations.Task.{}.Path.Source.0", taskId));
        paint();
        if (! popupState->DebugPathDetailField())
            return reject(L"the paused card's source path detail did not open");
        popupState->DebugArmAutoResizeMotionForSelfTest();
        paint();
        PopupTelemetryDebugSnapshot detailSettled{};
        const auto detailQuiet = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            detailSettled = popupState->DebugTelemetrySnapshot();
            return detailSettled.timerIntervalMs == 1000u && ! popupState->DebugHasVisibleUiMotion(popup.get());
        },
            {.timeout = SelfTest::Scale(3s), .operationName = L"path detail motion settles"});
        static_cast<void>(pumpFor(600ms));
        const PopupTelemetryDebugSnapshot detailBefore = popupState->DebugTelemetrySnapshot();
        const auto detailElapsed                       = pumpFor(1200ms);
        const PopupTelemetryDebugSnapshot detailAfter  = popupState->DebugTelemetrySnapshot();
        const uint64_t detailTicks                     = detailAfter.timerTicks - detailBefore.timerTicks;
        const uint64_t detailRenders                   = detailAfter.taskSnapshots - detailBefore.taskSnapshots;
        Debug::Perf::Emit(L"FileOps.PopupCadence.Counters",
                          L"synthetic-paused-path-detail",
                          static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(detailElapsed.elapsed).count()),
                          detailTicks,
                          detailRenders,
                          S_OK);
        note(std::format(L"synthetic-paused-path-detail: settled={} timerTicks={} renders={} intervalMs={} visibleMotion={}",
                         detailQuiet.conditionMet,
                         detailTicks,
                         detailRenders,
                         detailAfter.timerIntervalMs,
                         popupState->DebugHasVisibleUiMotion(popup.get())));
        if (! detailQuiet.conditionMet || ! popupState->DebugPathDetailField() || detailTicks > 2u || detailRenders > 3u)
            fail(std::format(
                L"open path detail kept the motion cadence: intervalMs={} timerTicks={} renders={}", detailAfter.timerIntervalMs, detailTicks, detailRenders));
    }

    // 2. The real engine: a Copy with a heavy retained history that first waits on a real conflict
    //    prompt (with a second Copy queued behind it), then runs held, paused, resumed and finally
    //    completes through the fallback queue. Asserts the popup cadence, the owner heartbeat and
    //    the per-tick summary copies.
    const bool autoDismiss       = operations->GetAutoDismissSuccess();
    const auto root              = state.tempRoot / L"popup-idle-cadence";
    const auto source            = root / L"source" / L"payload.bin";
    const auto destination       = root / L"destination";
    const auto queuedSource      = root / L"queued-source" / L"payload.bin";
    const auto queuedDestination = root / L"queued-destination";
    if (! RecreateEmptyDirectory(root) || ! WriteTestFile(source, 8193u) || ! WriteTestFile(destination / L"payload.bin", 4097u) ||
        ! WriteTestFile(queuedSource, 4099u) || ! SelfTest::EnsureDirectory(queuedDestination))
        return reject(L"owned scratch unavailable");
    static constexpr std::array<uint64_t, 4> historyIds{{0xF26A2026100903A1ull, 0xF26A2026100903A2ull, 0xF26A2026100903A3ull, 0xF26A2026100903A4ull}};
    constexpr size_t kHistoryPaths  = 2000u;
    constexpr size_t kHistoryIssues = 128u;
    uint64_t taskId                 = 0u;
    uint64_t queuedTaskId           = 0u;
    bool workerQuiet                = true;
    bool cleanupAttempted           = false;
    bool cleanupSucceeded           = false;
    const auto cleanup              = [&]() noexcept
    {
        if (cleanupAttempted)
            return cleanupSucceeded;
        cleanupAttempted = true;
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
        ReleaseFileOpsFallbackWorkerReturnPauseForSelfTest();
        operations->DebugSuppressTaskPresentationPostsForSelfTest(false);
        workerQuiet = CancelOutstandingSelfTestOperations(operations);
        if (! workerQuiet)
        {
            state.retainTempRootOnFailure = true;
            AppendLog(L"Popup idle cadence: bounded drain failed; hard-retiring the owner before settings restoration.");
            operations->Shutdown();
            workerQuiet = ! operations->HasActiveOperations();
        }
        for (const uint64_t id : historyIds)
            operations->DebugRemoveDiagnosticsForTask(id);
        if (workerQuiet)
        {
            if (taskId != 0u)
                operations->DismissCompletedTask(taskId);
            if (queuedTaskId != 0u)
                operations->DismissCompletedTask(queuedTaskId);
            operations->DebugResetPopupForSelfTest();
        }
        operations->SetAutoDismissSuccess(autoDismiss);
        const bool saved = operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u);
        cleanupSucceeded = workerQuiet && saved;
        return cleanupSucceeded;
    };
    auto restore    = wil::scope_exit([&]() noexcept { static_cast<void>(cleanup()); });
    auto removeRoot = wil::scope_exit([&]() noexcept
    {
        if (workerQuiet && ! state.retainTempRootOnFailure)
            static_cast<void>(SelfTest::RemoveAll(root));
    });

    for (const uint64_t id : historyIds)
    {
        State::CompletedTaskSummary summary{};
        summary.taskId              = id;
        summary.operation           = FILESYSTEM_MOVE;
        summary.resultHr            = S_OK;
        summary.completedTick       = GetTickCount64();
        summary.totalItems          = static_cast<unsigned long>(kHistoryPaths);
        summary.completedItems      = static_cast<unsigned long>(kHistoryPaths);
        summary.retainedSourceCount = static_cast<unsigned long>(kHistoryPaths);
        summary.warningCount        = static_cast<unsigned long>(kHistoryIssues);
        summary.sourcePath          = std::format(L"C:\\Cadence\\History\\{:X}", id);
        summary.retainedSourcePaths.reserve(kHistoryPaths);
        for (size_t index = 0u; index < kHistoryPaths; ++index)
            summary.retainedSourcePaths.emplace_back(std::format(L"C:\\Cadence\\History\\{:X}\\kept-{:05}.bin", id, index));
        summary.issueDiagnostics.reserve(kHistoryIssues);
        for (size_t index = 0u; index < kHistoryIssues; ++index)
        {
            State::TaskDiagnosticEntry issue{};
            issue.taskId     = id;
            issue.operation  = FILESYSTEM_MOVE;
            issue.severity   = State::DiagnosticSeverity::Warning;
            issue.status     = HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION);
            issue.category   = L"cadence";
            issue.message    = std::format(L"Retained cadence issue {}", index);
            issue.sourcePath = std::format(L"C:\\Cadence\\History\\{:X}\\kept-{:05}.bin", id, index);
            summary.issueDiagnostics.push_back(std::move(issue));
        }
        operations->DebugAppendCompletedTaskForSelfTest(std::move(summary));
    }

    // model#3: the Issues snapshot carries the retained issue entries, never the per-root paths,
    // and copies nothing while the revision is unchanged.
    {
        std::vector<State::TaskDiagnosticEntry> diagnostics;
        std::vector<State::TaskDiagnosticEntry> completedIssues;
        uint64_t revision = 0u;
        if (! operations->CollectIssuesSnapshotIfChanged(0u, diagnostics, completedIssues, revision))
            return reject(L"the Issues snapshot reported no change for a new revision");
        const auto historyIssues = std::ranges::count_if(
            completedIssues, [&](const State::TaskDiagnosticEntry& issue) noexcept { return std::ranges::find(historyIds, issue.taskId) != historyIds.end(); });
        std::vector<State::TaskDiagnosticEntry> unchangedDiagnostics;
        std::vector<State::TaskDiagnosticEntry> unchangedIssues;
        uint64_t unchangedRevision = revision;
        const bool copiedAgain     = operations->CollectIssuesSnapshotIfChanged(revision, unchangedDiagnostics, unchangedIssues, unchangedRevision);
        note(std::format(L"issues snapshot: completedIssues={} history={} copiedAgain={}", completedIssues.size(), historyIssues, copiedAgain));
        if (static_cast<size_t>(historyIssues) != historyIds.size() * kHistoryIssues || copiedAgain || ! unchangedIssues.empty())
            fail(std::format(L"the Issues snapshot carried {} retained issue entries (expected {}) or copied again on an unchanged revision",
                             historyIssues,
                             historyIds.size() * kHistoryIssues));
    }

    operations->SetAutoDismissSuccess(false);
    operations->DebugEnsurePopupVisibleForSelfTest();
    const HWND popup = operations->GetPopupHwndForSelfTest();
    if (! popup)
        return reject(L"the File Operations popup is unavailable");
    ShowWindow(popup, SW_SHOWNOACTIVATE);

    struct Sample final
    {
        PopupTelemetryDebugSnapshot popup{};
        uint64_t recoveryTicks      = 0u;
        uint64_t presentationBuilds = 0u;
        uint64_t projectableScans   = 0u;
        uint64_t uiCpuUs            = 0u;
    };
    const auto sample = [&]() noexcept
    {
        Sample value{};
        static_cast<void>(DebugGetFileOperationsPopupTelemetrySnapshot(popup, value.popup));
        value.recoveryTicks      = operations->DebugCompletionRecoveryTickCountForSelfTest();
        value.presentationBuilds = operations->DebugCompletedPresentationBuildCountForSelfTest();
        value.projectableScans   = operations->DebugProjectableIssueScanCountForSelfTest();
        value.uiCpuUs            = uiThreadCpuUs();
        return value;
    };
    const auto waitForInterval = [&](UINT intervalMs, std::chrono::milliseconds timeout, std::wstring_view operation) noexcept
    {
        PopupTelemetryDebugSnapshot observed{};
        return RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return DebugGetFileOperationsPopupTelemetrySnapshot(popup, observed) && observed.timerIntervalMs == intervalMs;
        }, {.timeout = timeout, .operationName = operation});
    };
    const auto measure = [&](std::wstring_view name, std::chrono::milliseconds duration, Sample& before, Sample& after) noexcept
    {
        before                   = sample();
        const auto elapsed       = pumpFor(duration);
        after                    = sample();
        const uint64_t elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed.elapsed).count());
        const uint64_t ticks     = after.popup.timerTicks - before.popup.timerTicks;
        const uint64_t renders   = after.popup.taskSnapshots - before.popup.taskSnapshots;
        const uint64_t heartbeat = after.recoveryTicks - before.recoveryTicks;
        Debug::Perf::Emit(L"FileOps.PopupCadence.Counters", name, elapsedUs, ticks, renders, S_OK);
        Debug::Perf::Emit(L"FileOps.PopupCadence.OwnerHeartbeat", name, elapsedUs, heartbeat, operations->DebugLastPresentationHeartbeatMsForSelfTest(), S_OK);
        Debug::Perf::Emit(L"FileOps.PopupCadence.CompletedCopies",
                          name,
                          ticks,
                          after.presentationBuilds - before.presentationBuilds,
                          after.projectableScans - before.projectableScans,
                          S_OK);
        Debug::Perf::Emit(L"FileOps.PopupCadence.UiCpuUs", name, elapsedUs, after.uiCpuUs - before.uiCpuUs, 0u, S_OK);
        note(std::format(L"{}: elapsedMs={} timerTicks={} renders={} intervalMs={} ownerTicks={} heartbeatMs={} presentationBuilds={} "
                         L"projectableScans={} uiCpuUs={}",
                         name,
                         elapsed.elapsed.count(),
                         ticks,
                         renders,
                         after.popup.timerIntervalMs,
                         heartbeat,
                         operations->DebugLastPresentationHeartbeatMsForSelfTest(),
                         after.presentationBuilds - before.presentationBuilds,
                         after.projectableScans - before.projectableScans,
                         after.uiCpuUs - before.uiCpuUs));
        return elapsed.conditionMet;
    };
    const auto work = [](State::PopupTelemetryWork value) noexcept { return static_cast<unsigned>(value); };

    // Prompt-blocked: the destination already holds the payload, so the Copy first waits on a real
    // published conflict prompt; the popup follows it at the waiting cadence (x-perf#5).
    SetFileOpsLiveOutputPublishedPauseForSelfTest(true);
    workerQuiet         = false;
    const auto admitted = StartFileOperationAndGetId(operations,
                                                     FILESYSTEM_COPY,
                                                     FolderWindow::Pane::Left,
                                                     FolderWindow::Pane::Right,
                                                     folder->GetFileSystem(FolderWindow::Pane::Left),
                                                     {source},
                                                     destination,
                                                     FILESYSTEM_FLAG_NONE,
                                                     false,
                                                     0u,
                                                     State::ExecutionMode::PerItem,
                                                     false,
                                                     folder->GetFileSystem(FolderWindow::Pane::Right));
    if (! admitted.has_value())
        return reject(L"the real Copy admission failed");
    taskId              = admitted.value();
    State::Task* task   = nullptr;
    const auto prompted = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        task = operations->FindTask(taskId);
        return TryGetConflictPromptCopy(task).has_value();
    },
        {.timeout = SelfTest::Scale(8s), .operationName = L"cadence copy conflict prompt"});
    if (! prompted.conditionMet || ! task)
        return reject(L"the real Copy did not publish its conflict prompt");
    if (const State::PopupTelemetryWork promptWork = operations->GetPopupTelemetryWork(); promptWork != State::PopupTelemetryWork::Waiting)
        fail(std::format(L"a published prompt was classified {} (expected Waiting)", work(promptWork)));
    if (! waitForInterval(1000u, SelfTest::Scale(2s), L"prompting copy waiting cadence").conditionMet)
        fail(L"the prompting Copy kept the 100 ms telemetry cadence");
    static_cast<void>(pumpFor(600ms)); // the prompt card's bounded control transitions finish first
    Sample before{}, after{};
    if (! measure(L"engine-prompting", 1200ms, before, after))
        return reject(L"prompting measurement pump failed");
    if (after.popup.timerTicks - before.popup.timerTicks > 2u || after.popup.taskSnapshots - before.popup.taskSnapshots > 3u)
        fail(std::format(L"the prompting Copy kept waking the UI: timerTicks={} renders={}",
                         after.popup.timerTicks - before.popup.timerTicks,
                         after.popup.taskSnapshots - before.popup.taskSnapshots));

    // Queued: a Copy that waits for others is held in the queue behind the prompting Copy and is
    // waiting work; canceling it makes it active telemetry at once (it may roll back or clean up).
    {
        const auto queuedAdmitted = StartFileOperationAndGetId(operations,
                                                               FILESYSTEM_COPY,
                                                               FolderWindow::Pane::Left,
                                                               FolderWindow::Pane::Right,
                                                               folder->GetFileSystem(FolderWindow::Pane::Left),
                                                               {queuedSource},
                                                               queuedDestination,
                                                               FILESYSTEM_FLAG_NONE,
                                                               true,
                                                               0u,
                                                               State::ExecutionMode::PerItem,
                                                               false,
                                                               folder->GetFileSystem(FolderWindow::Pane::Right));
        if (! queuedAdmitted.has_value())
            return reject(L"the queued Copy admission failed");
        queuedTaskId            = queuedAdmitted.value();
        State::Task* queuedTask = nullptr;
        const auto queued       = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            queuedTask = operations->FindTask(queuedTaskId);
            return queuedTask && queuedTask->IsWaitingInQueue() && ! queuedTask->HasEnteredOperation();
        },
            {.timeout = SelfTest::Scale(4s), .operationName = L"cadence copy waits in the queue"});
        if (! queued.conditionMet || ! queuedTask)
            return reject(L"the second Copy did not wait in the queue");
        const State::PopupTelemetryWork queuedWork = operations->GetPopupTelemetryWork();
        const auto queuedCadence                   = waitForInterval(1000u, SelfTest::Scale(2s), L"queued copy waiting cadence");
        queuedTask->RequestCancel();
        const State::PopupTelemetryWork stoppingWork = operations->GetPopupTelemetryWork();
        note(std::format(L"queued: work={} intervalReached={} canceledWork={}", work(queuedWork), queuedCadence.conditionMet, work(stoppingWork)));
        if (queuedWork != State::PopupTelemetryWork::Waiting || ! queuedCadence.conditionMet || stoppingWork != State::PopupTelemetryWork::Active)
            fail(std::format(L"queued classification work={} cadenceReached={} canceledWork={} (expected Waiting, true, Active)",
                             work(queuedWork),
                             queuedCadence.conditionMet,
                             work(stoppingWork)));
        const auto queuedGone = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return ! operations->FindTask(queuedTaskId);
        }, {.timeout = SelfTest::Scale(8s), .operationName = L"canceled queued copy completes"});
        if (! queuedGone.conditionMet)
            return reject(L"the canceled queued Copy did not complete");
    }

    // Answering the prompt posts a presentation change, so the 100 ms cadence returns at once. The
    // answer follows a slow popup tick whose own motion (if any) has settled, and re-arms the owner
    // heartbeat, so for 400 ms (scaled for slowed runs) neither a slow tick nor a 16 ms motion tick
    // can stand in for the post.
    {
        // The removed queued card may still resize the popup; let that motion finish first.
        static_cast<void>(pumpFor(1000ms));
        auto* popupState = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup, GWLP_USERDATA));
        if (! popupState)
            return reject(L"the File Operations popup state is unavailable");
        bool popupTicked = false;
        for (int attempt = 0; attempt < 4 && ! popupTicked; ++attempt)
        {
            if (! waitForInterval(1000u, SelfTest::Scale(2s), L"prompting copy waiting cadence before the answer").conditionMet)
                continue;
            PopupTelemetryDebugSnapshot start{};
            static_cast<void>(DebugGetFileOperationsPopupTelemetrySnapshot(popup, start));
            const auto ticked = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                PopupTelemetryDebugSnapshot observed{};
                return DebugGetFileOperationsPopupTelemetrySnapshot(popup, observed) && observed.timerTicks != start.timerTicks;
            },
                {.timeout = SelfTest::Scale(2s), .operationName = L"popup waiting tick before the answer"});
            if (! ticked.conditionMet)
                continue;
            const ULONGLONG tickedAt = GetTickCount64();
            static_cast<void>(pumpFor(150ms));
            popupTicked = waitForInterval(1000u, 300ms, L"popup motion after the waiting tick").conditionMet && ! popupState->DebugHasVisibleUiMotion(popup) &&
                          GetTickCount64() - tickedAt <= 450u;
        }
        SendMessageW(folder->GetHwnd(), WndMsg::kFileOperationPresentationChanged, 0, 0); // re-arms the owner heartbeat now
        PopupTelemetryDebugSnapshot beforeAnswer{};
        static_cast<void>(DebugGetFileOperationsPopupTelemetrySnapshot(popup, beforeAnswer));
        const uint64_t ownerTicksBeforeAnswer = operations->DebugCompletionRecoveryTickCountForSelfTest();
        task->DebugSubmitCurrentConflictDecision(State::Task::ConflictAction::Overwrite, false);
        PopupTelemetryDebugSnapshot observed{};
        const auto answered = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return DebugGetFileOperationsPopupTelemetrySnapshot(popup, observed) && observed.timerIntervalMs == 100u;
        }, {.timeout = SelfTest::Scale(400ms), .operationName = L"answered prompt telemetry cadence"});
        note(std::format(L"answer restored the 100 ms cadence: {} in {} ms (intervalMs={}, aligned={}, popupTicks={}, renders={}, ownerRefreshes={})",
                         answered.conditionMet,
                         answered.elapsed.count(),
                         observed.timerIntervalMs,
                         popupTicked,
                         observed.timerTicks - beforeAnswer.timerTicks,
                         observed.taskSnapshots - beforeAnswer.taskSnapshots,
                         operations->DebugCompletionRecoveryTickCountForSelfTest() - ownerTicksBeforeAnswer));
        Debug::Perf::Emit(L"FileOps.PopupCadence.PromptAnswerResumeUs",
                          L"engine-prompt-answer",
                          static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(answered.elapsed).count()),
                          answered.conditionMet ? 1u : 0u,
                          observed.timerIntervalMs,
                          answered.conditionMet ? S_OK : E_FAIL);
        if (! popupTicked || ! answered.conditionMet)
            fail(std::format(L"answering the prompt kept the waiting cadence for {} ms (intervalMs={})", answered.elapsed.count(), observed.timerIntervalMs));
    }
    const auto running = RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept { return HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest(); },
                                                                       {.timeout = SelfTest::Scale(8s), .operationName = L"cadence copy publication gate"});
    if (! running.conditionMet || operations->FindTask(taskId) != task || ! task->HasEnteredOperation() || task->_taskFinished.load(std::memory_order_acquire))
        return reject(L"the answered Copy did not reach its publication gate");
    task->_presentationState.store(State::Task::TaskPresentationState::Presented, std::memory_order_release);

    // Running: 100 ms frames, but the retained history is projected once, not copied per tick
    // (x-perf#3), and the owner heartbeat has backed off with no fallback pending (x-lifetime#7).
    if (! waitForInterval(100u, SelfTest::Scale(2s), L"running copy telemetry cadence").conditionMet)
        fail(L"the running Copy did not select the 100 ms telemetry cadence");
    static_cast<void>(pumpFor(1200ms)); // let the admission's fast heartbeat back off
    if (! measure(L"engine-running", 1200ms, before, after))
        return reject(L"running measurement pump failed");
    if (after.popup.timerIntervalMs != 100u || after.popup.timerTicks - before.popup.timerTicks < 6u)
        fail(L"the running Copy lost its 100 ms telemetry");
    if (after.presentationBuilds - before.presentationBuilds > 1u || after.projectableScans - before.projectableScans > 1u)
        fail(std::format(L"the popup rebuilt the retained history {} times and rescanned issues {} times in {} ticks",
                         after.presentationBuilds - before.presentationBuilds,
                         after.projectableScans - before.projectableScans,
                         after.popup.timerTicks - before.popup.timerTicks));
    if (operations->HasPendingFallbackCompletions() || operations->DebugLastPresentationHeartbeatMsForSelfTest() != State::kCompletionRecoveryIdleHeartbeatMs ||
        after.recoveryTicks - before.recoveryTicks > 3u)
        fail(std::format(L"the owner heartbeat did not back off: heartbeatMs={} ownerTicks={}",
                         operations->DebugLastPresentationHeartbeatMsForSelfTest(),
                         after.recoveryTicks - before.recoveryTicks));

    // Paused: the waiting cadence for the popup, the backed-off heartbeat for the owner (x-perf#5).
    task->SetPaused(true);
    if (! waitForInterval(1000u, SelfTest::Scale(2s), L"paused copy waiting cadence").conditionMet)
        fail(L"the paused Copy kept the 100 ms telemetry cadence");
    static_cast<void>(pumpFor(600ms)); // the Paused card's bounded control transitions finish first
    if (! measure(L"engine-paused", 2200ms, before, after))
        return reject(L"paused measurement pump failed");
    if (after.popup.timerTicks - before.popup.timerTicks > 3u || after.popup.taskSnapshots - before.popup.taskSnapshots > 4u ||
        after.recoveryTicks - before.recoveryTicks > 4u)
        fail(std::format(L"the paused Copy kept waking the UI: timerTicks={} renders={} ownerTicks={}",
                         after.popup.timerTicks - before.popup.timerTicks,
                         after.popup.taskSnapshots - before.popup.taskSnapshots,
                         after.recoveryTicks - before.recoveryTicks));

    // Resume re-evaluates at once through the presentation post, not on the next slow tick.
    task->SetPaused(false);
    const auto resumed = waitForInterval(100u, SelfTest::Scale(500ms), L"resumed copy telemetry cadence");
    note(std::format(L"resume re-armed the 100 ms cadence in {} ms", resumed.elapsed.count()));
    if (! resumed.conditionMet)
        fail(L"resuming the Copy did not restore the 100 ms cadence promptly");

    // Prompt-blocked: the engine classifies a published decision as waiting, and its metadata load
    // (which still changes the card) as active. The flags are restored before any pump.
    {
        std::scoped_lock promptLock(task->_conflictArbiter.mutex);
        task->_conflictArbiter.prompt.active          = true;
        task->_conflictArbiter.prompt.metadataLoading = false;
    }
    const State::PopupTelemetryWork promptWork = operations->GetPopupTelemetryWork();
    {
        std::scoped_lock promptLock(task->_conflictArbiter.mutex);
        task->_conflictArbiter.prompt.metadataLoading = true;
    }
    const State::PopupTelemetryWork loadingWork = operations->GetPopupTelemetryWork();
    {
        std::scoped_lock promptLock(task->_conflictArbiter.mutex);
        task->_conflictArbiter.prompt.active          = false;
        task->_conflictArbiter.prompt.metadataLoading = false;
    }
    const State::PopupTelemetryWork resumedWork = operations->GetPopupTelemetryWork();
    if (promptWork != State::PopupTelemetryWork::Waiting || loadingWork != State::PopupTelemetryWork::Active ||
        resumedWork != State::PopupTelemetryWork::Active)
        fail(std::format(L"prompt classification prompt={} loading={} after={}",
                         static_cast<unsigned>(promptWork),
                         static_cast<unsigned>(loadingWork),
                         static_cast<unsigned>(resumedWork)));

    // x-lifetime#7: the Copy completes with every completion post failing, and its worker parks
    // before the reaper can join it. An early payload-less wake for that pending fallback returns
    // the backed-off owner heartbeat to 100 ms at once. Presentation posts are suppressed and the
    // owner heartbeat is re-armed just before, so no other refresh can pick 100 ms in its place.
    {
        operations->DebugSuppressTaskPresentationPostsForSelfTest(true);
        SetFileOpsFallbackWorkerReturnPauseForSelfTest(true);
        SendMessageW(folder->GetHwnd(), WndMsg::kFileOperationPresentationChanged, 0, 0);
        const UINT heartbeatBefore = operations->DebugLastPresentationHeartbeatMsForSelfTest();
        const ULONGLONG armedTick  = GetTickCount64();
        operations->DebugForceNextFileOperationCompletedPostFailure();
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        // Wait without dispatching posted messages or timers (only cross-thread sends), so the
        // owner runs no heartbeat between the fallback and the wake.
        bool parked                 = false;
        const uint64_t parkWindowMs = SelfTest::ScaleTimeout(600u);
        while (GetTickCount64() - armedTick < parkWindowMs)
        {
            if (HasFileOpsFallbackWorkerReturnPauseEnteredForSelfTest() && operations->DebugFallbackCompletionCountForSelfTest() == 1u)
            {
                parked = true;
                break;
            }
            MSG sent{};
            static_cast<void>(PeekMessageW(&sent, nullptr, 0, 0, PM_NOREMOVE | PM_QS_SENDMESSAGE));
            Sleep(1);
        }
        const ULONGLONG wakeTick   = GetTickCount64();
        const uint64_t ticksBefore = operations->DebugCompletionRecoveryTickCountForSelfTest();
        if (parked)
            SendMessageW(folder->GetHwnd(), WndMsg::kFileOperationCompleted, static_cast<WPARAM>(taskId), 0);
        const auto wakeWindow     = SelfTest::Scale(400ms);
        const auto woken          = pumpFor(wakeWindow);
        const uint64_t ownerTicks = operations->DebugCompletionRecoveryTickCountForSelfTest() - ticksBefore;
        const UINT heartbeatAfter = operations->DebugLastPresentationHeartbeatMsForSelfTest();
        Debug::Perf::Emit(L"FileOps.PopupCadence.OwnerHeartbeat",
                          L"engine-fallback-wake",
                          static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(woken.elapsed).count()),
                          ownerTicks,
                          heartbeatAfter,
                          S_OK);
        note(std::format(L"fallback wake: parked={} parkedAfterMs={} heartbeatBefore={} ownerTicks={} heartbeatAfter={}",
                         parked,
                         wakeTick - armedTick,
                         heartbeatBefore,
                         ownerTicks,
                         heartbeatAfter));
        if (! parked)
            return reject(L"the Copy did not park its worker behind a pending fallback completion");
        if (heartbeatBefore != State::kCompletionRecoveryIdleHeartbeatMs || ownerTicks < 2u || heartbeatAfter != State::kCompletionRecoveryHeartbeatMs)
            fail(std::format(L"the payload-less wake left the owner heartbeat backed off: before={} ownerTicks={} after={} (in {} ms)",
                             heartbeatBefore,
                             ownerTicks,
                             heartbeatAfter,
                             wakeWindow.count()));
        ReleaseFileOpsFallbackWorkerReturnPauseForSelfTest();
        const auto adopted = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return ! operations->FindTask(taskId) && operations->DebugFallbackCompletionCountForSelfTest() == 0u;
        }, {.timeout = SelfTest::Scale(8s), .operationName = L"cadence copy fallback completion adopted"});
        operations->DebugSuppressTaskPresentationPostsForSelfTest(false);
        if (! adopted.conditionMet)
            return reject(L"the pending fallback completion was not adopted after its worker joined");
    }

    if (GetForegroundWindow() != foreground)
        note(L"the foreground window changed during the measurement (environmental; no fixture activates a window)");

    if (! cleanup())
        return reject(L"worker cleanup or the restored settings save did not reach its quiet point");
    restore.release();
    if (failures != 0u)
        return reject(std::format(L"{} cadence assertion(s) failed", failures));
    return true;
}

// I30 Wave 4 item 1 (alert#2, alert#9, alert#10, x-ux#12): the host prompt scenario. A window-scope
// alert and a host prompt share one tool window; the alert is cleared while the prompt's nested loop
// pumps, the window is resized and then destroyed. UI thread only.
struct AlertOverlayLifetimePromptScenario final
{
    HWND target           = nullptr;
    HWND windowAlert      = nullptr;
    size_t baselineAlerts = 0u;
    UINT_PTR stepTimer    = 0u;
    UINT_PTR watchdog     = 0u;
    bool stepRan          = false;
    bool sawBothOverlays  = false;
    bool alertCleared     = false;
    bool promptStillShown = false;
    bool promptFollows    = false;
    bool watchdogFired    = false;
};
AlertOverlayLifetimePromptScenario* g_alertOverlayLifetimePrompt = nullptr;

[[nodiscard]] std::vector<HWND> FindVisibleAlertOverlaysOwnedBy(HWND owner) noexcept
{
    std::vector<HWND> found;
    HWND candidate = nullptr;
    while ((candidate = FindWindowExW(nullptr, candidate, L"RedSalamander.AlertOverlayWindow", nullptr)) != nullptr)
    {
        if (GetWindow(candidate, GW_OWNER) == owner && IsWindowVisible(candidate) != FALSE)
            found.push_back(candidate);
    }
    return found;
}

[[nodiscard]] bool WindowCoversClientOf(HWND window, HWND target) noexcept
{
    RECT client{};
    RECT windowRect{};
    if (! window || GetClientRect(target, &client) == 0 || GetWindowRect(window, &windowRect) == 0)
        return false;
    POINT corners[2]{{client.left, client.top}, {client.right, client.bottom}};
    MapWindowPoints(target, nullptr, corners, 2);
    return windowRect.left == corners[0].x && windowRect.top == corners[0].y && windowRect.right == corners[1].x && windowRect.bottom == corners[1].y;
}

void CALLBACK AlertOverlayLifetimePromptStep(HWND, UINT, UINT_PTR timerId, DWORD) noexcept
{
    AlertOverlayLifetimePromptScenario* const scenario = g_alertOverlayLifetimePrompt;
    if (! scenario)
        return;
    if (timerId == scenario->watchdog)
    {
        // The prompt outlived its owner: end it through application prompt shutdown and fail.
        KillTimer(nullptr, timerId);
        scenario->watchdog      = 0u;
        scenario->watchdogFired = true;
        HostBeginPromptShutdown();
        return;
    }
    if (timerId != scenario->stepTimer)
        return;
    KillTimer(nullptr, timerId);
    scenario->stepTimer = 0u;
    if (scenario->stepRan)
        return;
    scenario->stepRan = true;

    const std::vector<HWND> overlays = FindVisibleAlertOverlaysOwnedBy(scenario->target);
    HWND prompt                      = nullptr;
    for (const HWND overlay : overlays)
    {
        if (overlay != scenario->windowAlert)
            prompt = overlay;
    }
    scenario->sawBothOverlays  = overlays.size() == 2u && prompt != nullptr && std::ranges::find(overlays, scenario->windowAlert) != overlays.end();
    scenario->alertCleared     = SUCCEEDED(HostClearAlert(HOST_ALERT_SCOPE_WINDOW, scenario->target)) && IsWindow(scenario->windowAlert) == FALSE &&
                                 HostGetTestWindowAlertOverlayCount() == scenario->baselineAlerts;
    scenario->promptStillShown = prompt != nullptr && IsWindowVisible(prompt) != FALSE;
    RECT targetRect{};
    static_cast<void>(GetWindowRect(scenario->target, &targetRect));
    SetWindowPos(scenario->target, nullptr, targetRect.left + 24, targetRect.top + 12, 440, 300, SWP_NOZORDER | SWP_NOACTIVATE);
    scenario->promptFollows = WindowCoversClientOf(prompt, scenario->target);
    // The user closes the tool window while the prompt is up.
    DestroyWindow(scenario->target);
}

[[nodiscard]] bool RunFileOpsAlertOverlayLifetimeChecks(SelfTestState& state) noexcept
{
    uint32_t failures = 0u;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Alert overlay lifetime: FAILED: {}", reason));
        return false;
    };
    const auto require = [&failures](bool condition, std::wstring_view what) noexcept
    {
        if (! condition)
        {
            ++failures;
            AppendLog(std::format(L"Alert overlay lifetime: FAILED: {}", what));
        }
    };
    if (! state.mainWindow || IsWindow(state.mainWindow) == FALSE || HostIsPromptShutdown())
        return reject(L"live main window without application prompt shutdown unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const uint64_t foregroundBefore = RedSalamander::Ui::DebugGetAlertOverlayForegroundRequestCount();

    // Overlays sharing one anchor chain independently; an owner destroyed without the hook clears visibility.
    RedSalamander::Ui::AlertOverlayTrackingSelfTestResult tracking{};
    if (! RedSalamander::Ui::DebugRunAlertOverlayTrackingSelfTest(state.mainWindow, tracking))
        return reject(L"the overlay tracking fixture could not run");
    require(tracking.siblingDetachKeepsTracking, L"removing one overlay unhooked the other overlay on the same anchor (it stopped following resizes)");
    require(tracking.laterSubclassKept, L"the last overlay to detach dropped a subclass installed after it");
    require(tracking.originalProcedureRestored, L"the anchor's original procedure or hook state was not restored after every overlay detached");
    require(tracking.ownerDestroyedClearsVisibility, L"an overlay destroyed with its owner still reports visible when no hook saw the owner go");
    require(tracking.closeHidesWithoutEscapeBinding, L"Escape on a closable alert without an escape binding did not hide it");
    require(tracking.targetDestroyedCallbackReported, L"the owner was not told that the tracked window died, or could not delete the overlay there");
    require(tracking.foregroundRequests == 0u, std::format(L"hiding modeless alerts requested the foreground {} time(s)", tracking.foregroundRequests));

    RECT mainRect{};
    static_cast<void>(GetWindowRect(state.mainWindow, &mainRect));
    const auto createTarget = [&mainRect]() noexcept
    {
        return CreateWindowExW(
            0, L"STATIC", L"", WS_POPUP | WS_VISIBLE, mainRect.left + 64, mainRect.top + 64, 380, 260, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr);
    };

    AlertOverlayLifetimePromptScenario scenario{};
    scenario.baselineAlerts = HostGetTestWindowAlertOverlayCount();
    scenario.target         = createTarget();
    const auto cleanup      = wil::scope_exit([&scenario]() noexcept
    {
        if (scenario.stepTimer != 0u)
            KillTimer(nullptr, scenario.stepTimer);
        if (scenario.watchdog != 0u)
            KillTimer(nullptr, scenario.watchdog);
        g_alertOverlayLifetimePrompt = nullptr;
        if (scenario.target && IsWindow(scenario.target) != FALSE)
            DestroyWindow(scenario.target);
        if (scenario.watchdogFired)
            HostResetPromptShutdown();
    });
    if (! scenario.target)
        return reject(L"the prompt's tool window could not be created");

    HostAlertRequest alert{};
    alert.sizeBytes    = sizeof(alert);
    alert.scope        = HOST_ALERT_SCOPE_WINDOW;
    alert.modality     = HOST_ALERT_MODELESS;
    alert.severity     = HOST_ALERT_INFO;
    alert.targetWindow = scenario.target;
    alert.title        = L"Alert overlay lifetime";
    alert.message      = L"Window-scope alert shown next to a prompt.";
    alert.closable     = TRUE;
    if (FAILED(HostShowAlert(alert)) || HostGetTestWindowAlertOverlayCount() != scenario.baselineAlerts + 1u)
        return reject(L"the window-scope alert was not shown");
    const std::vector<HWND> shown = FindVisibleAlertOverlaysOwnedBy(scenario.target);
    scenario.windowAlert          = shown.size() == 1u ? shown.front() : nullptr;
    if (! scenario.windowAlert)
        return reject(std::format(L"expected one visible window-scope alert overlay, found {}", shown.size()));

    wil::com_ptr<IHostPrompts> prompts;
    if (FAILED(GetHostServices()->QueryInterface(__uuidof(IHostPrompts), prompts.put_void())) || ! prompts)
        return reject(L"IHostPrompts unavailable");
    HostPromptRequest prompt{};
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_WARNING;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = scenario.target;
    prompt.title         = L"Alert overlay lifetime";
    prompt.message       = L"Prompt on the same window as an alert.";
    prompt.defaultResult = HOST_PROMPT_RESULT_NONE;
    prompt.presentation  = HOST_PROMPT_PRESENTATION_DEFAULT;

    g_alertOverlayLifetimePrompt = &scenario;
    scenario.stepTimer           = SetTimer(nullptr, 0u, 10u, &AlertOverlayLifetimePromptStep);
    scenario.watchdog =
        SetTimer(nullptr, 0u, static_cast<UINT>(std::min<uint64_t>(SelfTest::ScaleTimeout(10'000u), USER_TIMER_MAXIMUM)), &AlertOverlayLifetimePromptStep);
    if (scenario.stepTimer == 0u || scenario.watchdog == 0u)
        return reject(L"the prompt step or watchdog timer could not be started");
    // The real host prompt (the interface bypasses the test prompt overrides).
    HostPromptResult result = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt  = prompts->ShowPrompt(&prompt, nullptr, &result);
    require(scenario.stepRan, L"the prompt returned before its nested loop dispatched the scenario step");
    require(scenario.sawBothOverlays, L"the prompt and the window-scope alert were not both shown on the tool window");
    require(scenario.alertCleared, L"clearing the window-scope alert during the prompt did not destroy it and drop its entry");
    require(scenario.promptStillShown, L"clearing the window-scope alert hid the prompt");
    require(scenario.promptFollows, L"after the alert was cleared the prompt no longer followed its tool window (shared hook lost)");
    require(! scenario.watchdogFired, L"the prompt's nested loop kept running after its tool window was destroyed (ended by the watchdog)");
    require(SUCCEEDED(hrPrompt) && result == HOST_PROMPT_RESULT_CANCEL,
            std::format(L"the prompt whose owner died returned hr=0x{:08X} result={} instead of the escape result",
                        static_cast<unsigned long>(hrPrompt),
                        static_cast<int>(result)));
    require(IsWindow(scenario.target) == FALSE, L"the tool window was not destroyed by the scenario step");

    // A window-scope alert whose target dies leaves no map entry behind.
    HWND shortLived = createTarget();
    if (! shortLived)
        return reject(L"the short-lived target window could not be created");
    alert.targetWindow   = shortLived;
    const bool shownLive = SUCCEEDED(HostShowAlert(alert)) && HostGetTestWindowAlertOverlayCount() == scenario.baselineAlerts + 1u;
    DestroyWindow(shortLived);
    require(shownLive && HostGetTestWindowAlertOverlayCount() == scenario.baselineAlerts,
            std::format(
                L"a window-scope alert outlived its target window: {} entries, baseline {}", HostGetTestWindowAlertOverlayCount(), scenario.baselineAlerts));

    const uint64_t foregroundRequests = RedSalamander::Ui::DebugGetAlertOverlayForegroundRequestCount() - foregroundBefore;
    require(foregroundRequests == 0u, std::format(L"hiding the alerts and the orphaned prompt requested the foreground {} time(s)", foregroundRequests));
    if (failures != 0u)
        return reject(std::format(L"{} alert overlay lifetime assertion(s) failed", failures));
    AppendLog(L"Alert overlay lifetime: shared-anchor tracking, orphaned prompt exit, target-death cleanup and no-foreground hides verified.");
    return true;
}

[[nodiscard]] bool RunFileOpsAlertOverlayUxAccessibilityChecks(SelfTestState& state) noexcept
{
    uint32_t failures = 0u;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Alert overlay UX/accessibility: FAILED: {}", reason));
        return false;
    };
    const auto require = [&failures](bool condition, std::wstring_view what) noexcept
    {
        if (! condition)
        {
            ++failures;
            AppendLog(std::format(L"Alert overlay UX/accessibility: FAILED: {}", what));
        }
    };
    if (! state.mainWindow || IsWindow(state.mainWindow) == FALSE)
        return reject(L"live main window unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    // Run in French so the UI Automation names must come from the localized string resources.
    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui->language.empty() && g_settings.ui->language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui->language};
    const auto restoreLanguage = wil::scope_exit([&]() noexcept { static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage)); });
    if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, L"fr-FR"})))
        return reject(L"the fr-FR language could not be applied");

    RedSalamander::Ui::AlertOverlayUxAccessibilitySelfTestResult result{};
    if (! RedSalamander::Ui::DebugRunAlertOverlayUxAccessibilitySelfTest(state.mainWindow, result))
        return reject(L"the alert overlay UX/accessibility fixture could not run");

    require(result.optionFixtureScrolls, L"fixture: the option alert does not scroll or the target option is not fully visible");
    require(result.pointerActivationKeepsScroll,
            std::format(L"clicking a visible option moved the scrolled content (offset {:.1f} -> {:.1f} DIP)",
                        result.optionScrollBeforeDip,
                        result.optionScrollAfterDip));
    require(result.arrowActivationKeepsScroll, L"cycling a visible option with an arrow key moved the scrolled content");
    require(result.spaceActivationKeepsScroll, L"cycling a visible option with Space moved the scrolled content");
    require(result.iconFixtureScrolls, L"fixture: the tall alert does not scroll or shows no icon");
    require(result.iconInsideVisibleBody,
            std::format(L"the prompt icon [{:.1f}, {:.1f}] is not centred in the visible body [{:.1f}, {:.1f}]",
                        result.iconRect.top,
                        result.iconRect.bottom,
                        result.bodyViewport.top,
                        result.bodyViewport.bottom));
    require(result.tabRaisesFocusEvent, L"Tab moved the focus inside the focused alert without a UIA focus-changed event on the new control");
    require(result.uiaSetFocusRaisesFocusEvent, L"UIA SetFocus on a button raised no focus-changed event on it");
    require(result.arrowRaisesFocusEvent, L"an arrow key moved the button focus without a UIA focus-changed event");
    require(result.optionCycleRaisesNameChange, L"Space changed the focused option's value without a UIA name-changed event (or raised a focus event)");
    require(result.unfocusedMoveRaisesNoFocusEvent, L"a focus move inside an alert that does not have the keyboard focus raised a UIA focus event");
    require(result.replacementRaisesStructureChange, L"replacing the model of a visible alert raised no UIA structure-changed event on the new root");
    require(result.replacementAnnouncesText, L"replacing the model of a visible alert with new text raised no UIA notification");
    require(result.replacementRaisesFocusEvent, L"replacing the model of a focused alert raised no UIA focus event on the new focused control");
    require(result.reshowKeepsRestoreFocus, L"a blocking alert re-shown while focused replaced the focus it restores with the overlay itself");
    require(result.overlayNeverRecordedAsRestoreFocus, L"a blocking alert recorded its own overlay window as the focus to restore");
    require(result.unmarkedLabelsHaveNoMnemonic, L"Alt+C invoked a button whose label has no '&' marker (first-letter guess)");
    require(result.markedMnemonicInvokesItsButton, L"Alt+key did not invoke exactly the button whose '&' marker names that key");
    require(result.ambiguousMnemonicIgnored, L"Alt+key invoked a button although two labels mark the same key");
    require(result.markersHiddenFromName, L"a button's UIA name still shows its '&' marker or drops a literal '&&'");
    const std::wstring expectedRoot  = LoadStringResource(nullptr, IDS_ALERT_OVERLAY_NAME);
    const std::wstring expectedClose = LoadStringResource(nullptr, IDS_ALERT_OVERLAY_CLOSE);
    require(expectedRoot == L"Alerte" && result.untitledRootName == expectedRoot,
            std::format(L"the untitled alert is named \"{}\" in fr-FR, expected \"{}\"", result.untitledRootName, expectedRoot));
    require(expectedClose == L"Fermer" && result.closeName == expectedClose,
            std::format(L"the close button is named \"{}\" in fr-FR, expected \"{}\"", result.closeName, expectedClose));
    require(result.forgedUiaCallIgnored, L"a UIA marshal message carrying a raw call pointer from an arbitrary sender ran the call");
    require(result.crossThreadUiaCallMarshals, L"a UIA provider called from another thread did not run on the overlay's thread");
    require(result.untakenSyncPayloadReclaimed, L"a synchronous payload no receiver took was not reclaimed when SendMessagePayload returned");
    require(result.foregroundRequests == 0u, std::format(L"the fixture's alerts requested the foreground {} time(s)", result.foregroundRequests));
    if (failures != 0u)
        return reject(std::format(L"{} alert overlay UX/accessibility assertion(s) failed", failures));
    AppendLog(L"Alert overlay UX/accessibility: steady option activation, visible-body icon, UIA focus/structure/notification events, restore "
              L"focus, '&' mnemonics, localized names and the token-only UIA marshal verified.");
    return true;
}

#include "FileOperations.Scope.Tests.DeleteProgress.cpp"
#include "FolderWindow.FileOperations.Tests.AclDeniedChild.cpp"
#include "FolderWindow.FileOperations.Tests.AdmissionBaseline.cpp"
#include "FolderWindow.FileOperations.Tests.BatchRenameExactResult.cpp"
#include "FolderWindow.FileOperations.Tests.CancelAll.cpp"
#include "FolderWindow.FileOperations.Tests.CaseOnlyStorage.cpp"
#include "FolderWindow.FileOperations.Tests.CommonPresentationPerf.cpp"
#include "FolderWindow.FileOperations.Tests.ConflictBroker.cpp"
#include "FolderWindow.FileOperations.Tests.CrossProviderRoutes.cpp"

// Review remediation calls the cross-provider checks above; preserve this implementation order.
#include "FileOperations.Scope.Tests.ReviewRemediation.cpp"

#include "FolderWindow.FileOperations.Tests.DecisionMirror.cpp"
#include "FolderWindow.FileOperations.Tests.DiscoveryClosure.cpp"
#include "FolderWindow.FileOperations.Tests.IdentityLess.cpp"
#include "FolderWindow.FileOperations.Tests.IntegratedPresentationPerf.cpp"
#include "FolderWindow.FileOperations.Tests.IssuesProjection.cpp"
#include "FolderWindow.FileOperations.Tests.ItemResultReducer.cpp"
#include "FolderWindow.FileOperations.Tests.ManagedMovePartialWrite.cpp"
#include "FolderWindow.FileOperations.Tests.Notes.cpp"
#include "FolderWindow.FileOperations.Tests.OverwritePreservation.cpp"
#include "FolderWindow.FileOperations.Tests.OwnedShutdown.cpp"
#include "FolderWindow.FileOperations.Tests.PathDetailFocus.cpp"
#include "FolderWindow.FileOperations.Tests.QueuedAdmission.cpp"
#include "FolderWindow.FileOperations.Tests.RepeatedStorage.cpp"
#include "FolderWindow.FileOperations.Tests.SameObject.cpp"
#include "FolderWindow.FileOperations.Tests.SmallFileCost.cpp"
#include "FolderWindow.FileOperations.Tests.Throughput.cpp"
#include "FolderWindow.FileOperations.Tests.VisualGallery.cpp"

[[nodiscard]] std::wstring_view ValidateCompletedTerminalFactsProjection(SelfTestState& state) noexcept
{
    using Summary  = FolderWindow::FileOperationState::CompletedTaskSummary;
    using Snapshot = FileOperationsPopupInternal::TaskSnapshot;

    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || operations->HasActiveOperations() || operations->GetPopupHwndForSelfTest())
    {
        return L"Terminal-facts projection requires an idle File Operations state with no application popup.";
    }

    constexpr uint64_t taskId = 0xF26A000000000001ull;
    std::vector<Summary> existing;
    operations->CollectCompletedTasks(existing);
    if (std::ranges::any_of(existing, [](const Summary& value) noexcept { return value.taskId == taskId; }))
    {
        return L"Terminal-facts projection task ID is already in use.";
    }

    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        return L"Terminal-facts projection could not acquire the no-activation window guard.";
    }

    Summary completed{};
    completed.taskId                           = taskId;
    completed.operation                        = FILESYSTEM_MOVE;
    completed.resultHr                         = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    completed.warningCount                     = 1u;
    completed.completedTick                    = GetTickCount64();
    completed.discoveryClosed                  = true;
    completed.totalItems                       = 10u;
    completed.completedItems                   = 10u;
    completed.publishedItemCount               = 9u;
    completed.unknownPublicationItemCount      = 1u;
    completed.removedSourceCount               = 7u;
    completed.retainedSourceCount              = 2u;
    completed.retainedSourceNativeCount        = 1u;
    completed.unknownSourceCount               = 1u;
    completed.verifiedItemCount                = 6u;
    completed.verificationFailedItemCount      = 1u;
    completed.verificationUnavailableItemCount = 1u;
    completed.verificationCanceledItemCount    = 1u;
    completed.verificationProblemItemCount     = 3u;
    completed.indeterminateItemCount           = 1u;
    completed.sourcePath                       = L"C:\\TerminalFacts\\source.txt";
    completed.destinationPath                  = L"D:\\TerminalFacts\\source.txt";
    completed.resultSummary                    = L"Nine published; one publication unknown; two sources kept; one source unknown.";

    FolderWindow::FileOperationState::TaskDiagnosticEntry issue{};
    issue.taskId    = taskId;
    issue.operation = FILESYSTEM_MOVE;
    issue.severity  = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
    issue.status    = completed.resultHr;
    issue.category  = L"task.summary";
    issue.message   = completed.resultSummary;
    completed.issueDiagnostics.push_back(std::move(issue));

    operations->DebugAppendCompletedTaskForSelfTest(std::move(completed));
    const auto dismissTask = wil::scope_exit([operations]() noexcept { operations->DismissCompletedTask(taskId); });

    auto lifetime = std::make_shared<int>(0);
    wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
    auto closePopup = wil::scope_exit([&]() noexcept
    {
        // This private hidden popup is not registered as the application's popup. Expire its
        // host token before WM_NCDESTROY so OnPopupDestroyed does not save application placement.
        lifetime.reset();
        popup.reset();
    });
    if (! popup)
    {
        return L"Terminal-facts projection could not create its private production popup.";
    }

    Snapshot projected{};
    if (! DebugGetFileOperationsPopupTaskSnapshot(popup.get(), taskId, projected) || ! projected.finished)
    {
        return L"The private production popup did not project its completed task summary.";
    }
    if (projected.publishedItemCount != 9u || projected.unknownPublicationItemCount != 1u || projected.removedSourceCount != 7u ||
        projected.retainedSourceCount != 2u || projected.retainedSourceNativeCount != 1u || projected.unknownSourceCount != 1u ||
        projected.verifiedItemCount != 6u || projected.verificationFailedItemCount != 1u || projected.verificationUnavailableItemCount != 1u ||
        projected.verificationCanceledItemCount != 1u || projected.verificationProblemItemCount != 3u)
    {
        return L"The popup lost independent completed publication, source, or verification counts.";
    }
    if (projected.resultSummary != L"Nine published; one publication unknown; two sources kept; one source unknown." || ! projected.verificationRequested ||
        projected.verificationState != static_cast<uint8_t>(FileOperations::VerificationState::Failed))
    {
        return L"The popup changed the exact terminal summary or legacy selected verification state.";
    }
    return {};
}

// Q8 (popup-c#2): Clear completed removes only the results its last painted header covered and
// never acknowledges an interrupted-Move breadcrumb, with the Completed group collapsed or expanded.
[[nodiscard]] std::wstring_view ValidateClearCompletedScope(SelfTestState& state) noexcept
{
    using Summary = FolderWindow::FileOperationState::CompletedTaskSummary;
    using Hit     = FileOperationsPopupInternal::PopupHitTest::Kind;
    using Layout  = FileOperationsPopupInternal::PopupLayoutDebugSnapshot;

    auto* folder = TryGetFolderWindow(state.mainWindow);
    if (! folder || state.tempRoot.empty())
        return L"Clear-completed scope requires the live folder window and the suite temp root.";

    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return L"Clear-completed scope could not acquire the no-activation window guard.";

    const std::filesystem::path root           = state.tempRoot / L"clear-completed-scope";
    const std::filesystem::path breadcrumbRoot = root / L"state";
    if (! SelfTest::EnsureDirectory(breadcrumbRoot))
        return L"Clear-completed scope could not create its owned durable store.";
    FileOperationMoveBreadcrumb::SetRootForSelfTest(breadcrumbRoot);
    const auto restoreRoot = wil::scope_exit([]() noexcept { FileOperationMoveBreadcrumb::ClearRootForSelfTest(); });

    FileOperationMoveBreadcrumb::Record record{};
    record.taskId                         = 0x5A01u;
    record.admittedStrategies.nativeItems = 1u;
    record.sourceRootCount                = 1u;
    record.sourceRootSamples              = {FileOperationMoveBreadcrumb::QualifiedLocation{
        .pluginId           = L"builtin/file-system",
        .pluginShortId      = L"file",
        .instanceId         = L"host/default",
        .profileId          = L"local",
        .rootId             = L"clear-source-root",
        .representativePath = root / L"source.txt",
    }};
    record.destination                    = FileOperationMoveBreadcrumb::QualifiedLocation{
        .pluginId           = L"builtin/file-system",
        .pluginShortId      = L"file",
        .instanceId         = L"host/default",
        .profileId          = L"local",
        .rootId             = L"clear-destination-root",
        .representativePath = root / L"destination",
    };
    FileOperationMoveBreadcrumb::Breadcrumb breadcrumb;
    if (FAILED(FileOperationMoveBreadcrumb::Breadcrumb::Create(record, breadcrumb)) ||
        FAILED(breadcrumb.Advance(FileOperationMoveBreadcrumb::DurablePhase::Executing)))
        return L"Clear-completed scope could not persist its interrupted-Move breadcrumb.";
    const std::filesystem::path breadcrumbPath = breadcrumb.GetPath();

    // A private engine loads the breadcrumb as a restart notice exactly as startup does, and
    // keeps the live host's completed history and application popup out of this witness.
    Common::Settings::Settings settings{};
    settings.fileOperations.emplace();
    auto owner = std::make_unique<FolderWindow>();
    owner->SetSettings(&settings);
    FolderWindow::FileOperationState operations(*owner);
    if (operations.GetPopupFooterOnly())
        return L"Clear-completed scope requires the default card list, not footer-only mode.";

    const auto retained = [&operations](uint64_t taskId) noexcept
    {
        std::vector<Summary> completed;
        operations.CollectCompletedTasks(completed);
        return std::ranges::any_of(completed, [taskId](const Summary& item) noexcept { return item.taskId == taskId; });
    };
    uint64_t noticeId = 0u;
    {
        std::vector<Summary> loaded;
        operations.CollectCompletedTasks(loaded);
        const auto notice = std::ranges::find_if(loaded, [&](const Summary& item) noexcept { return item.interruptedMoveBreadcrumbPath == breadcrumbPath; });
        if (notice != loaded.end() && notice->interruptedMoveNotice)
            noticeId = notice->taskId;
    }
    if (noticeId == 0u)
        return L"The restored interrupted-Move breadcrumb did not project as a notice card.";

    constexpr uint64_t kOrdinaryBase = 0xC1EA000000000000ull;
    const auto appendCompleted       = [&operations](uint64_t taskId) noexcept
    {
        Summary summary{};
        summary.taskId        = taskId;
        summary.operation     = FILESYSTEM_COPY;
        summary.resultHr      = S_OK;
        summary.completedTick = GetTickCount64();
        operations.DebugAppendCompletedTaskForSelfTest(std::move(summary));
    };
    appendCompleted(kOrdinaryBase + 1u);
    appendCompleted(kOrdinaryBase + 2u);

    auto lifetime = std::make_shared<int>(0);
    wil::unique_hwnd popup(FileOperationsPopup::Create(&operations, folder, state.mainWindow, lifetime));
    auto closePopup = wil::scope_exit([&]() noexcept
    {
        // Private popup: expire its host token before WM_NCDESTROY, as the terminal-facts witness does.
        lifetime.reset();
        popup.reset();
    });
    if (! popup)
        return L"Clear-completed scope could not create its private production popup.";
    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
    const auto paint  = [&popup]() noexcept { RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN); };
    const auto invoke = [&popup](Hit kind, uint64_t taskId) noexcept
    {
        FileOperationsPopupInternal::PopupSelfTestInvoke request{};
        request.kind   = kind;
        request.taskId = taskId;
        return DebugInvokeFileOperationsPopup(popup.get(), request);
    };
    const auto layoutFor = [&popup](uint64_t taskId, Layout& layout) noexcept
    {
        layout        = {};
        layout.taskId = taskId;
        return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), layout);
    };

    for (const bool collapsed : {true, false})
    {
        const std::wstring_view mode = collapsed ? L"collapsed" : L"expanded";
        const uint64_t seenA         = kOrdinaryBase + (collapsed ? 1u : 3u);
        const uint64_t seenB         = kOrdinaryBase + (collapsed ? 2u : 4u);
        const uint64_t late          = kOrdinaryBase + (collapsed ? 5u : 6u);
        paint();
        Layout layout{};
        if (layoutFor(seenA, layout) && layout.completedGroupExpanded == collapsed)
        {
            static_cast<void>(invoke(Hit::CompletedGroupToggle, 0u));
            paint();
        }
        Layout notice{};
        const bool groupReady  = layoutFor(seenA, layout) && layout.completedGroupVisible && layout.completedGroupExpanded != collapsed &&
                                 layout.completedGroupClearVisible && layout.taskHiddenByCompletedGroup == collapsed;
        const bool noticeShown = layoutFor(noticeId, notice) && ! notice.taskHiddenByCompletedGroup;
        AppendLog(std::format(L"Clear-completed scope {}: groupVisible={} expanded={} clearVisible={} memberHidden={} noticeHidden={}",
                              mode,
                              layout.completedGroupVisible,
                              layout.completedGroupExpanded,
                              layout.completedGroupClearVisible,
                              layout.taskHiddenByCompletedGroup,
                              notice.taskHiddenByCompletedGroup));
        if (! groupReady)
            return L"The Completed group did not publish Clear completed for its two painted ordinary results.";
        if (! noticeShown)
            return L"The Completed group hid the interrupted-Move notice behind a badge.";

        // A result that finishes after the last paint was never shown to the user.
        appendCompleted(late);
        if (! invoke(Hit::CompletedGroupClear, 0u))
            return L"The private popup did not dispatch Clear completed.";
        const bool seenCleared = ! retained(seenA) && ! retained(seenB);
        const bool lateKept    = retained(late);
        const bool noticeKept  = retained(noticeId);
        std::error_code existsError;
        const bool breadcrumbKept = std::filesystem::exists(breadcrumbPath, existsError) && ! existsError;
        AppendLog(std::format(L"Clear-completed scope {} after Clear: seenCleared={} lateKept={} noticeKept={} breadcrumbKept={}",
                              mode,
                              seenCleared,
                              lateKept,
                              noticeKept,
                              breadcrumbKept));
        if (! seenCleared)
            return L"Clear completed did not remove the results its painted header covered.";
        if (! lateKept)
            return L"Clear completed removed a result that finished after the last paint.";
        if (! noticeKept || ! breadcrumbKept)
            return L"Clear completed acknowledged the interrupted-Move notice or deleted its breadcrumb file.";

        // Start the expanded pass from two fresh ordinary results.
        if (collapsed)
        {
            operations.DismissCompletedTask(late);
            appendCompleted(kOrdinaryBase + 3u);
            appendCompleted(kOrdinaryBase + 4u);
        }
    }

    // OnActivatedHit reports 0 for a handled Dismiss; assert on its effect instead.
    static_cast<void>(invoke(Hit::TaskDismiss, noticeId));
    std::error_code existsError;
    const bool breadcrumbGone = ! std::filesystem::exists(breadcrumbPath, existsError) && ! existsError;
    if (retained(noticeId) || ! breadcrumbGone)
        return L"The notice's own Dismiss did not acknowledge its exact breadcrumb.";
    return {};
}

// I30 popup-a#1 / x-arch#2 (Q6): a frame the attached host cannot prepare is skipped and retried,
// a text measurement failure degrades the layout, and only a persistent failure selects the native
// failure surface, which reopening the same HWND or the backoff timer leaves once the host attaches.
// Every step must keep the conflict decision published on the hosted popup.
[[nodiscard]] bool RunFileOpsPopupFailureSurfaceRecoveryChecks(SelfTestState& state) noexcept
{
    using namespace std::chrono_literals;
    using Snapshot    = FileOperationsPopupInternal::TaskSnapshot;
    using Hit         = FileOperationsPopupInternal::PopupHitTest::Kind;
    using Layout      = FileOperationsPopupInternal::PopupLayoutDebugSnapshot;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Popup failure-surface recovery: {}", reason));
        return false;
    };

    auto* folder = TryGetFolderWindow(state.mainWindow);
    if (! folder)
        return reject(L"the live folder window is unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"could not acquire the no-activation window guard");
    const HWND foreground     = GetForegroundWindow();
    const auto resetInjection = wil::scope_exit([]() noexcept
    {
        DebugFailNextFileOperationsD2DTargetAttempts(0u);
        DebugFailNextFileOperationsPopupLayoutMeasurements(0u);
        DebugFailNextFileOperationsDxUiHostAttachAttempts(0u);
    });

    // A private engine keeps the application's popup and tasks out of this witness.
    Common::Settings::Settings settings{};
    settings.fileOperations.emplace();
    auto owner = std::make_unique<FolderWindow>();
    owner->SetSettings(&settings);
    FolderWindow::FileOperationState operations(*owner);
    if (operations.GetPopupFooterOnly())
        return reject(L"the private engine must default to the card list, not footer-only mode");

    auto lifetime = std::make_shared<int>(0);
    wil::unique_hwnd popup(FileOperationsPopup::Create(&operations, folder, state.mainWindow, lifetime));
    auto closePopup = wil::scope_exit([&]() noexcept
    {
        // Private popup: expire its host token before WM_NCDESTROY, as the terminal-facts witness does.
        lifetime.reset();
        popup.reset();
    });
    auto* popupState =
        popup ? reinterpret_cast<FileOperationsPopupInternal::FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA)) : nullptr;
    if (! popupState)
        return reject(L"could not create the private production popup");

    Snapshot task{};
    task.taskId                                 = 1u;
    task.operation                              = FILESYSTEM_COPY;
    task.lifecyclePhase                         = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
    task.started                                = true;
    task.discoveryClosed                        = true;
    task.totalBytes                             = 4096u;
    task.totalItems                             = 1u;
    task.plannedItems                           = 1u;
    task.currentSourcePath                      = L"C:\\Recovery\\report.txt";
    task.currentDestinationPath                 = L"D:\\Recovery\\report.txt";
    task.conflict.active                        = true;
    task.conflict.generation                    = 1u;
    task.conflict.bucket                        = static_cast<uint8_t>(FileOperations::ConflictClass::RegularFileExists);
    task.conflict.status                        = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    task.conflict.sourcePath                    = task.currentSourcePath;
    task.conflict.destinationPath               = task.currentDestinationPath;
    task.conflict.sourceName                    = L"report.txt";
    task.conflict.sourceParentPath              = L"C:\\Recovery";
    task.conflict.destinationName               = L"report.txt";
    task.conflict.destinationParentPath         = L"D:\\Recovery";
    task.conflict.sourceMetadata.available      = true;
    task.conflict.sourceMetadata.sizeKnown      = true;
    task.conflict.sourceMetadata.sizeBytes      = 4096u;
    task.conflict.sourceMetadata.lastWriteTime  = 134'338'680'000'000'000ll;
    task.conflict.destinationMetadata           = task.conflict.sourceMetadata;
    task.conflict.destinationMetadata.sizeBytes = 2048u;
    popupState->DebugSetVisualScenario({task}, false, true);
    popupState->debugFixedGeometry = true;

    const UINT dpi      = GetDpiForWindow(popup.get());
    const auto resizeTo = [&](int widthDip) noexcept
    {
        SetWindowPos(popup.get(),
                     nullptr,
                     0,
                     0,
                     MulDiv(widthDip, static_cast<int>(dpi), 96),
                     MulDiv(900, static_cast<int>(dpi), 96),
                     SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    };
    const auto paint      = [&]() noexcept { RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN); };
    const auto childCount = [&]() noexcept
    {
        size_t count = 0u;
        for (HWND child = GetWindow(popup.get(), GW_CHILD); child != nullptr; child = GetWindow(child, GW_HWNDNEXT))
            ++count;
        return count;
    };
    const auto describe = [&](std::wstring_view step) noexcept
    {
        Layout layout{};
        layout.taskId   = task.taskId;
        const bool read = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), layout);
        AppendLog(std::format(L"Popup failure-surface recovery {}: read={} hosted={} surface={} ready={} surfaceChildren={} children={} prepared={} "
                              L"publishable={} primary={} hostedActions={}",
                              step,
                              read,
                              layout.usesDxUiHost,
                              layout.failureSurfaceActive,
                              layout.failureSurfaceReady,
                              layout.failureSurfaceChildCount,
                              childCount(),
                              popupState->DebugHasCurrentPreparedFrame(),
                              layout.conflictButtonsPublishable,
                              layout.conflictPrimaryActionCount,
                              popupState->DebugHostedHitCountForSelfTest(Hit::TaskConflictAction, task.taskId)));
        return layout;
    };
    // The hosted popup is back: no native child, a prepared frame and every primary decision published.
    const auto hostedWithDecisions = [&]() noexcept
    {
        Layout layout{};
        layout.taskId = task.taskId;
        return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), layout) && layout.usesDxUiHost && ! layout.failureSurfaceActive &&
               layout.failureSurfaceChildCount == 0u && childCount() == 0u && popupState->DebugHasCurrentPreparedFrame() && layout.conflictButtonsPublishable &&
               layout.conflictPrimaryActionCount > 0u &&
               popupState->DebugHostedHitCountForSelfTest(Hit::TaskConflictAction, task.taskId) == layout.conflictPrimaryActionCount;
    };
    const auto onFailureSurface = [&]() noexcept
    {
        Layout layout{};
        layout.taskId = task.taskId;
        return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), layout) && layout.failureSurfaceActive && layout.failureSurfaceReady &&
               ! layout.usesDxUiHost && layout.failureSurfaceChildCount == 3u && childCount() == 3u;
    };
    const auto waitFor = [](auto&& condition, std::wstring_view operation, std::chrono::milliseconds timeout) noexcept
    { return RedSalamander::TestSupport::PumpMessagesUntil(condition, {.timeout = SelfTest::Scale(timeout), .operationName = operation}).conditionMet; };

    resizeTo(760);
    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
    paint();
    if (! hostedWithDecisions())
    {
        static_cast<void>(describe(L"baseline"));
        return reject(L"the hosted popup did not publish the conflict decision before any injected failure");
    }

    // popup-a#1: one frame the host cannot prepare (ResizeBuffers after a GPU reset) is skipped,
    // not a reason for the permanent native surface; the popup timer prepares the next one.
    DebugFailNextFileOperationsD2DTargetAttempts(1u);
    paint();
    const Layout skipped = describe(L"after one failed frame");
    if (skipped.failureSurfaceActive || childCount() != 0u || popupState->DebugHasCurrentPreparedFrame())
        return reject(L"one failed frame selected the failure surface (or was not injected) instead of being retried");
    if (! waitFor(hostedWithDecisions, L"retry of a skipped popup frame", 5000ms))
    {
        static_cast<void>(describe(L"after the retry"));
        return reject(L"the popup timer did not prepare the skipped frame and restore the conflict decision");
    }

    // x-arch#2: a text measurement failure degrades the conflict actions to full-width rows.
    DebugFailNextFileOperationsPopupLayoutMeasurements(1000u);
    resizeTo(700);
    paint();
    {
        const Layout degraded = describe(L"with failing measurements");
        size_t fullWidthRows  = 0u;
        float firstLeft       = 0.0f;
        float firstRight      = 0.0f;
        for (const auto& button : popupState->DebugVisualButtons())
        {
            if (button.hit.kind != Hit::TaskConflictAction || button.hit.taskId != task.taskId)
                continue;
            if (fullWidthRows == 0u)
            {
                firstLeft  = button.bounds.left;
                firstRight = button.bounds.right;
            }
            fullWidthRows += button.bounds.left == firstLeft && button.bounds.right == firstRight ? 1u : 0u;
        }
        if (degraded.failureSurfaceActive || childCount() != 0u || ! hostedWithDecisions() || fullWidthRows != degraded.conflictPrimaryActionCount)
            return reject(L"a measurement failure selected the failure surface or withdrew a conflict decision instead of a fallback layout");
    }
    DebugFailNextFileOperationsPopupLayoutMeasurements(0u);
    resizeTo(760);
    paint();
    if (! hostedWithDecisions())
        return reject(L"the measured conflict layout did not return once measurement succeeded");

    // A persistent failure still reaches the failure surface, only after repeated failures over time.
    DebugFailNextFileOperationsD2DTargetAttempts(1000u);
    ULONGLONG failingSince = GetTickCount64();
    paint();
    if (! waitFor(onFailureSurface, L"persistent frame failure selects the failure surface", 10'000ms))
    {
        static_cast<void>(describe(L"persistent failure"));
        return reject(L"a persistent frame failure never reached the native failure surface");
    }
    // The popup requires five failures over at least two seconds (Specs/UI/UI_FileOperationsPopup.md).
    if (GetTickCount64() - failingSince < 1900u)
        return reject(L"the failure surface was selected before the retry window elapsed");

    // Close and reopen the same HWND once the device recovers: the hosted popup returns.
    // The re-attach backoff is at least one second away, so only WM_SHOWWINDOW(TRUE) can leave
    // the surface before any message is pumped; the timer path is proven separately below.
    DebugFailNextFileOperationsD2DTargetAttempts(0u);
    if (! onFailureSurface())
        return reject(L"the failure surface was left before the popup was reopened");
    ShowWindow(popup.get(), SW_HIDE);
    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
    {
        Layout reopened{};
        reopened.taskId = task.taskId;
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), reopened) || ! reopened.usesDxUiHost || reopened.failureSurfaceActive ||
            reopened.failureSurfaceChildCount != 0u || childCount() != 0u)
        {
            static_cast<void>(describe(L"right after reopening"));
            return reject(L"WM_SHOWWINDOW(TRUE) did not re-attach the host on the failure surface");
        }
    }
    // a3-2: no frame has published the decision since the host returned, so Escape binds no decision
    // (the failure surface never showed it) and closes the popup instead of submitting it.
    if (popupState->DebugConflictEscapeTaskIdForSelfTest() != 0u)
        return reject(std::format(L"the re-attached host kept the escape decision of task {} from before the failure surface",
                                  popupState->DebugConflictEscapeTaskIdForSelfTest()));
    SendMessageW(popup.get(), WM_KEYDOWN, VK_ESCAPE, 1);
    SendMessageW(popup.get(), WM_KEYUP, VK_ESCAPE, 0);
    if (IsWindowVisible(popup.get()))
        return reject(L"Escape right after the host re-attached was consumed as a decision instead of closing the popup");
    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
    paint();
    if (! waitFor(hostedWithDecisions, L"reopened popup re-attaches its host", 5000ms))
    {
        static_cast<void>(describe(L"after reopening"));
        return reject(L"reopening the same HWND did not re-attach the host and restore the conflict decision");
    }

    // Without reopening, the backoff timer re-attempts the host as well.
    DebugFailNextFileOperationsD2DTargetAttempts(1000u);
    failingSince = GetTickCount64();
    paint();
    if (! waitFor(onFailureSurface, L"second persistent failure selects the failure surface", 10'000ms))
        return reject(L"the second persistent frame failure never reached the native failure surface");
    DebugFailNextFileOperationsD2DTargetAttempts(0u);
    if (! waitFor(hostedWithDecisions, L"backoff timer re-attaches the host", 10'000ms))
    {
        static_cast<void>(describe(L"after the re-attach timer"));
        return reject(L"the backoff timer did not re-attach the host from the failure surface");
    }
    static_cast<void>(describe(L"recovered"));
    AppendLog(std::format(L"Popup failure-surface recovery: second surface selected after {} ms.", GetTickCount64() - failingSince));
    if (GetForegroundWindow() != foreground)
        return reject(L"recovery changed the foreground window");
    return true;
}

// I30 host#2: a pane Copy/Move/Delete whose admission fails must not vanish. A recovery-heartbeat
// failure is refused before any prompt; one after an accepted prompt still reaches the pane alert.
[[nodiscard]] bool RunFileOpsAdmissionFailureReachesPaneChecks(SelfTestState& state) noexcept
{
    constexpr FolderWindow::Pane kSource      = FolderWindow::Pane::Left;
    constexpr FolderWindow::Pane kDestination = FolderWindow::Pane::Right;
    const HRESULT armFailure                  = HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA);
    const auto reject                         = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Admission failure reaches pane: {}", reason));
        return false;
    };

    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || state.tempRoot.empty() || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"live owner, temp root or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const auto root                                          = state.tempRoot / L"admission-failure-pane";
    const auto sourceRoot                                    = root / L"source";
    const auto destination                                   = root / L"destination";
    const auto payload                                       = sourceRoot / L"payload.bin";
    const auto sentinel                                      = destination / L"sentinel.bin";
    const std::optional<std::filesystem::path> originalLeft  = folder->GetCurrentPath(kSource);
    const std::optional<std::filesystem::path> originalRight = folder->GetCurrentPath(kDestination);
    const HWND originalPopup                                 = operations->GetPopupHwndForSelfTest();
    const bool autoAccept                                    = HostGetAutoAcceptPrompts();
    const auto promptOverride                                = HostGetTestPromptResultOverride();
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
    operations->CollectCompletedTasks(completed);
    const size_t originalCompletedCount = completed.size();
    bool passed                         = false;
    const auto cleanup                  = wil::scope_exit([&]() noexcept
    {
        operations->DebugFailRecoveryTimerArmForSelfTest(-1);
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        folder->DismissPaneAlertOverlay(kSource);
        const bool quiet = CancelOutstandingSelfTestOperations(operations);
        if (! originalPopup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        if (originalLeft.has_value())
            folder->SetFolderPath(kSource, originalLeft.value());
        if (originalRight.has_value())
            folder->SetFolderPath(kDestination, originalRight.value());
        if (! passed || ! quiet)
            state.retainTempRootOnFailure = true;
    });

    if ((folder->GetFileSystemPluginId(kSource) != kPluginIdLocal && FAILED(folder->SetFileSystemPluginForPane(kSource, kPluginIdLocal))) ||
        (folder->GetFileSystemPluginId(kDestination) != kPluginIdLocal && FAILED(folder->SetFileSystemPluginForPane(kDestination, kPluginIdLocal))))
        return reject(L"Local panes unavailable");
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(sourceRoot) || ! RecreateEmptyDirectory(destination) ||
        ! WriteFilledTestFile(payload, 4096u, 0x5Au) || ! WriteFilledTestFile(sentinel, 512u, 0xA5u))
        return reject(L"owned source and destination folders unavailable");
    folder->SetFolderPath(kSource, sourceRoot);
    folder->SetFolderPath(kDestination, destination);
    const auto listed = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        return CurrentPanePathEquals(folder, kSource, sourceRoot) && CurrentPanePathEquals(folder, kDestination, destination) &&
               folder->DebugHasItemDisplayName(kSource, L"payload.bin") && folder->DebugHasItemDisplayName(kDestination, L"sentinel.bin");
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"admission-failure pane listing"});
    if (! listed.conditionMet)
        return reject(L"the panes did not list the owned source and destination folders");

    // Accept every prompt, so a prompt shown before the fault is observable rather than a cancel.
    HostSetAutoAcceptPrompts(false);
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    const auto exists        = [](const std::filesystem::path& path) noexcept { return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES; };
    const std::wstring hrHex = std::format(L"{:08X}", static_cast<unsigned long>(armFailure));
    const auto attempt       = [&](std::wstring_view name, UINT titleId, int32_t passingArms, uint64_t expectedPrompts, const auto& command) noexcept
    {
        FolderView::AlertOverlayDebugSnapshot alert{};
        folder->DismissPaneAlertOverlay(kSource);
        const auto cleared = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return folder->DebugGetPaneAlertSnapshot(kSource, alert) && ! alert.visible;
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(2)), .operationName = L"admission-failure alert dismissal"});
        if (! cleared.conditionMet || ! folder->DebugFocusItemByDisplayName(kSource, L"payload.bin"))
            return reject(std::format(L"{}: the source pane alert did not clear or the payload could not be focused", name));

        const uint64_t promptsBefore = HostGetTestPromptRequestCount();
        operations->DebugFailRecoveryTimerArmForSelfTest(passingArms);
        command();
        operations->DebugFailRecoveryTimerArmForSelfTest(-1);
        const uint64_t prompts = HostGetTestPromptRequestCount() - promptsBefore;
        const auto shown       = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return folder->DebugGetPaneAlertSnapshot(kSource, alert) && alert.visible;
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(2)), .operationName = L"admission-failure pane alert"});
        operations->CollectCompletedTasks(completed);
        const std::wstring expectedTitle = LoadStringResource(nullptr, titleId);
        const bool untouched             = exists(payload) && exists(sentinel) && ! exists(destination / payload.filename());
        AppendLog(std::format(L"Admission failure reaches pane {}: prompts={} (expected {}) alertVisible={} alertHr=0x{:08X} title='{}' "
                              L"message='{}' active={} completed={} (was {}) untouched={}",
                              name,
                              prompts,
                              expectedPrompts,
                              alert.visible,
                              static_cast<unsigned long>(alert.hr),
                              alert.title,
                              alert.message,
                              operations->HasActiveOperations(),
                              completed.size(),
                              originalCompletedCount,
                              untouched));
        if (prompts != expectedPrompts)
            return reject(std::format(L"{}: {} prompt(s) were shown, expected {}", name, prompts, expectedPrompts));
        if (! shown.conditionMet || alert.kind != FolderView::ErrorOverlayKind::Operation || alert.hr != armFailure || alert.title != expectedTitle ||
            alert.message.find(hrHex) == std::wstring::npos)
            return reject(std::format(L"{}: the failed admission did not reach the source pane alert", name));
        if (operations->HasActiveOperations() || completed.size() != originalCompletedCount || ! untouched)
            return reject(std::format(L"{}: a refused admission published a task or changed the owned files", name));
        return true;
    };

    // Copy and Move with options always ask first: a heartbeat failure must be refused before the
    // prompt, and a failure after the accepted prompt must still be reported. Delete and Permanent
    // delete (whose consent is on the task card) report it as well. Delete runs last, after the
    // fault has proved it refuses admission, so a Recycle can never reach the user's bin.
    passed = attempt(L"Copy before consent", IDS_FILEOP_OPERATION_COPY, 0, 0u, [&]() { folder->CommandCopyToOtherPane(kSource, true); }) &&
             attempt(L"Copy after consent", IDS_FILEOP_OPERATION_COPY, 1, 1u, [&]() { folder->CommandCopyToOtherPane(kSource, true); }) &&
             attempt(L"Move before consent", IDS_FILEOP_OPERATION_MOVE, 0, 0u, [&]() { folder->CommandMoveToOtherPane(kSource, true); }) &&
             attempt(L"Move after consent", IDS_FILEOP_OPERATION_MOVE, 1, 1u, [&]() { folder->CommandMoveToOtherPane(kSource, true); }) &&
             attempt(L"Permanent delete", IDS_FILEOP_OPERATION_DELETE, 0, 0u, [&]() { folder->CommandPermanentDelete(kSource); }) &&
             attempt(L"Delete", IDS_FILEOP_OPERATION_DELETE, 0, 0u, [&]() { folder->CommandDelete(kSource); });
    if (! passed)
        return false;

    // A Copy the operation-artifact touch guard blocks is explained by the guard's blocked prompt
    // (Close is its default and Escape answer), which says the operation was not started; the pane must not add a second, generic alert. The
    // Dummy provider offers no object binding, so an artifact-shaped destination name cannot be
    // verified and the guard blocks the Copy.
    passed                                                 = false;
    constexpr std::wstring_view kArtifactLeaf              = L"staged.rs_tmp_0123456789abcdef0123456789abcdef_1";
    const std::wstring artifactRoot                        = std::format(L"/admission-failure-artifact-{}", GetTickCount64());
    const std::wstring artifactSource                      = artifactRoot + L"/source";
    const std::wstring artifactDestination                 = artifactRoot + L"/destination";
    const std::wstring artifactFile                        = artifactSource + L"/" + std::wstring(kArtifactLeaf);
    const auto restoreLocalPanes                           = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(folder->SetFileSystemPluginForPane(kSource, kPluginIdLocal));
        static_cast<void>(folder->SetFileSystemPluginForPane(kDestination, kPluginIdLocal));
        if (state.fsDummy)
            static_cast<void>(state.fsDummy->DeleteItem(
                artifactRoot.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
    });
    const FileSystemPluginManager::PluginEntry* dummyEntry = FindLoadedPluginEntry(kPluginIdDummy);
    const std::wstring_view dummyShortId = dummyEntry && ! dummyEntry->shortId.empty() ? std::wstring_view(dummyEntry->shortId) : std::wstring_view(L"dummy");
    if (! state.fsDummy || ! EnsureDummyFolderExists(state.fsDummy.get(), artifactRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), artifactSource) ||
        ! EnsureDummyFolderExists(state.fsDummy.get(), artifactDestination) || ! DummyWriteTextFile(state.fsDummy.get(), artifactFile, "artifact") ||
        FAILED(folder->SetFileSystemPluginForPane(kSource, kPluginIdDummy)) || FAILED(folder->SetFileSystemPluginForPane(kDestination, kPluginIdDummy)))
        return reject(L"blocked artifact Copy: the Dummy panes and folders are unavailable");
    const std::filesystem::path artifactSourceDisplay      = MakePluginDisplayPath(dummyShortId, artifactSource);
    const std::filesystem::path artifactDestinationDisplay = MakePluginDisplayPath(dummyShortId, artifactDestination);
    folder->SetFolderPath(kSource, artifactSourceDisplay);
    folder->SetFolderPath(kDestination, artifactDestinationDisplay);
    FolderView::AlertOverlayDebugSnapshot artifactAlert{};
    const auto artifactListed = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        return CurrentPanePathEquals(folder, kSource, artifactSourceDisplay) && CurrentPanePathEquals(folder, kDestination, artifactDestinationDisplay) &&
               folder->DebugHasItemDisplayName(kSource, kArtifactLeaf);
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"blocked-artifact pane listing"});
    folder->DismissPaneAlertOverlay(kSource);
    const auto artifactCleared = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return folder->DebugGetPaneAlertSnapshot(kSource, artifactAlert) && ! artifactAlert.visible;
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(2)), .operationName = L"blocked-artifact alert dismissal"});
    if (! artifactListed.conditionMet || ! artifactCleared.conditionMet || ! folder->DebugFocusItemByDisplayName(kSource, kArtifactLeaf))
        return reject(L"blocked artifact Copy: the Dummy panes did not list the artifact, the alert did not clear or the artifact could not be focused");
    const uint64_t artifactPromptsBefore = HostGetTestPromptRequestCount();
    folder->CommandCopyToOtherPane(kSource, false);
    const uint64_t artifactPrompts = HostGetTestPromptRequestCount() - artifactPromptsBefore;
    // Give a second alert time to appear; none may.
    const auto artifactShown = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return folder->DebugGetPaneAlertSnapshot(kSource, artifactAlert) && artifactAlert.visible;
    }, {.timeout = SelfTest::Scale(std::chrono::milliseconds(750)), .operationName = L"blocked-artifact second alert"});
    operations->CollectCompletedTasks(completed);
    const std::wstring copiedArtifact = artifactDestination + L"/" + std::wstring(kArtifactLeaf);
    wil::com_ptr<IFileSystemIO> artifactIo;
    unsigned long artifactAttributes = 0u;
    const bool artifactUntouched     = SUCCEEDED(state.fsDummy->QueryInterface(IID_PPV_ARGS(artifactIo.addressof()))) && artifactIo &&
                                       SUCCEEDED(artifactIo->GetAttributes(artifactFile.c_str(), &artifactAttributes)) &&
                                       FAILED(artifactIo->GetAttributes(copiedArtifact.c_str(), &artifactAttributes));
    AppendLog(std::format(L"Admission failure reaches pane blocked artifact Copy: prompts={} (expected 1) secondAlert={} alertHr=0x{:08X} "
                          L"message='{}' active={} completed={} (was {}) untouched={}",
                          artifactPrompts,
                          artifactShown.conditionMet,
                          static_cast<unsigned long>(artifactAlert.hr),
                          artifactAlert.message,
                          operations->HasActiveOperations(),
                          completed.size(),
                          originalCompletedCount,
                          artifactUntouched));
    if (artifactPrompts != 1u)
        return reject(std::format(L"blocked artifact Copy: {} prompt(s) were shown, expected the guard's one blocked warning", artifactPrompts));
    if (artifactShown.conditionMet)
        return reject(L"blocked artifact Copy: a second, generic pane alert followed the guard's blocked warning");
    if (operations->HasActiveOperations() || completed.size() != originalCompletedCount || ! artifactUntouched)
        return reject(L"blocked artifact Copy: the blocked admission published a task or changed the owned files");
    passed = true;
    return passed;
}

// A refused admission must never show the FolderView rename prompt or the Change Case prompt. While
// the refusal routes run, a thread timer counts and cancels either one, so a regression fails the case
// instead of blocking it in a modal loop.
std::atomic<uint32_t> g_deferredExitModalPromptCount{0u};

void CALLBACK CancelDeferredExitModalPrompt(HWND, UINT, UINT_PTR, DWORD) noexcept
{
    if (GetFolderViewRenamePromptHandle() != nullptr && DebugCancelFolderViewRenamePrompt())
        g_deferredExitModalPromptCount.fetch_add(1u, std::memory_order_acq_rel);
    if (GetFolderViewChangeCasePromptHandle() != nullptr && DebugCancelFolderViewChangeCasePrompt())
        g_deferredExitModalPromptCount.fetch_add(1u, std::memory_order_acq_rel);
}

// The close targets record whether the exit was still pending when the resumed WM_CLOSE reached
// them: anything dispatched ahead of that WM_CLOSE (a posted Change Case result or Delete command)
// must have met the pending exit and been refused, not started new work. UI thread only.
const FolderWindow* g_deferredExitCloseFolder = nullptr;
WNDPROC g_deferredExitCloseTargetBaseProc     = nullptr;
int g_deferredExitPendingAtClose              = -1; // -1: no WM_CLOSE yet; 0: not pending; 1: pending.

LRESULT CALLBACK RecordDeferredExitCloseTarget(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_CLOSE && g_deferredExitCloseFolder != nullptr)
        g_deferredExitPendingAtClose = g_deferredExitCloseFolder->IsFileOperationsExitPending() ? 1 : 0;
    return RedSalamander::Win32Callback::CallWindowProcNoThrow(g_deferredExitCloseTargetBaseProc, hwnd, msg, wp, lp);
}

// I30 x-lifetime#4 (owner decision Q14): an accepted "Cancel all and exit" refuses new work. The exit
// consent is asked only while a task can still be canceled, a second close attempt keeps the pending
// exit without asking again, every admission route (main panes, FolderView prompts, Batch Rename,
// Change Case, Find Files and a real Compare Directories window) is refused before any prompt,
// clipboard use or breadcrumb with a pane alert, and a task queued behind the cancelled work never starts.
[[nodiscard]] bool RunFileOpsDeferredExitRefusesNewWorkChecks(SelfTestState& state) noexcept
{
    constexpr FolderWindow::Pane kSource      = FolderWindow::Pane::Left;
    constexpr FolderWindow::Pane kDestination = FolderWindow::Pane::Right;
    const auto reject                         = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Deferred exit refuses new work: {}", reason));
        return false;
    };

    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || state.tempRoot.empty() || ! CancelOutstandingSelfTestOperations(operations) ||
        folder->IsFileOperationsExitPending())
        return reject(L"live owner, Local file system, temp root or an idle operation state without a pending exit unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const auto root            = state.tempRoot / L"deferred-exit-refusal";
    const auto breadcrumbRoot  = root / L"state";
    const auto holdSource      = root / L"hold-src" / L"collision.bin";
    const auto holdDestination = root / L"hold-dst" / L"collision.bin";
    const auto queuedSource    = root / L"queued-src" / L"queued.bin";
    const auto queuedTarget    = root / L"queued-dst";
    const auto sourceRoot      = root / L"source";
    const auto destination     = root / L"destination";
    const auto payload         = sourceRoot / L"payload.bin";
    const auto sentinel        = destination / L"sentinel.bin";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(breadcrumbRoot) || ! RecreateEmptyDirectory(sourceRoot) ||
        ! RecreateEmptyDirectory(destination) || ! RecreateEmptyDirectory(queuedTarget) || ! WriteFilledTestFile(holdSource, 4096u, 0x11u) ||
        ! WriteFilledTestFile(holdDestination, 8192u, 0x22u) || ! WriteFilledTestFile(queuedSource, 2048u, 0x33u) ||
        ! WriteFilledTestFile(payload, 4096u, 0x5Au) || ! WriteFilledTestFile(sentinel, 512u, 0xA5u))
        return reject(L"owned scratch folders could not be created");
    FileOperationMoveBreadcrumb::SetRootForSelfTest(breadcrumbRoot);

    const std::optional<std::filesystem::path> originalLeft  = folder->GetCurrentPath(kSource);
    const std::optional<std::filesystem::path> originalRight = folder->GetCurrentPath(kDestination);
    const HWND originalPopup                                 = operations->GetPopupHwndForSelfTest();
    const bool autoAccept                                    = HostGetAutoAcceptPrompts();
    const auto promptOverride                                = HostGetTestPromptResultOverride();
    // Message-only stand-ins for the main window: the default WM_CLOSE handling destroys them, which
    // is how the resumed close is observed.
    std::array<HWND, 2> closeTargets{};
    for (HWND& target : closeTargets)
    {
        target = CreateWindowExW(0, L"STATIC", L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, nullptr, nullptr);
        if (target != nullptr)
            g_deferredExitCloseTargetBaseProc = reinterpret_cast<WNDPROC>(
                RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(target, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(&RecordDeferredExitCloseTarget)));
    }
    g_deferredExitCloseFolder     = folder;
    HWND compareWindow            = nullptr;
    FolderWindow* compareFolder   = nullptr;
    UINT_PTR modalPromptTimer     = 0u;
    const auto closeOwnedTopLevel = [](HWND window) noexcept
    {
        if (window == nullptr || IsWindow(window) == FALSE)
            return;
        PostMessageW(window, WM_CLOSE, 0, 0);
        static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return IsWindow(window) == FALSE;
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"deferred-exit window close"}));
    };
    bool passed        = false;
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        if (modalPromptTimer != 0u)
            KillTimer(nullptr, modalPromptTimer);
        FolderView::DebugSetNextMoveSelectedItemsDestinationForSelfTest(std::nullopt);
        closeOwnedTopLevel(GetBatchRenameWindowHandle());
        closeOwnedTopLevel(compareWindow);
        ReleaseFileOpsConflictMetadataPauseForSelfTest();
        SetFileOpsConflictMetadataPauseForSelfTest(false, 0u);
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        folder->DismissPaneAlertOverlay(kSource);
        folder->DismissPaneAlertOverlay(kDestination);
        const bool quiet = CancelOutstandingSelfTestOperations(operations);
        // With nothing active this disarms any exit the checks left pending, so later cases can admit.
        if (quiet && folder->IsFileOperationsExitPending())
            static_cast<void>(folder->CancelAllFileOperationsThenClose(nullptr));
        for (HWND target : closeTargets)
            if (target != nullptr && IsWindow(target) != FALSE)
                DestroyWindow(target);
        g_deferredExitCloseFolder = nullptr;
        FileOperationMoveBreadcrumb::ClearRootForSelfTest();
        if (! originalPopup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        if (originalLeft.has_value())
            folder->SetFolderPath(kSource, originalLeft.value());
        if (originalRight.has_value())
            folder->SetFolderPath(kDestination, originalRight.value());
        if (! passed || ! quiet)
            state.retainTempRootOnFailure = true;
    });
    if (closeTargets[0] == nullptr || closeTargets[1] == nullptr)
        return reject(L"message-only close targets could not be created");

    if ((folder->GetFileSystemPluginId(kSource) != kPluginIdLocal && FAILED(folder->SetFileSystemPluginForPane(kSource, kPluginIdLocal))) ||
        (folder->GetFileSystemPluginId(kDestination) != kPluginIdLocal && FAILED(folder->SetFileSystemPluginForPane(kDestination, kPluginIdLocal))))
        return reject(L"Local panes unavailable");
    folder->SetFolderPath(kSource, sourceRoot);
    folder->SetFolderPath(kDestination, destination);
    const auto listed = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        return CurrentPanePathEquals(folder, kSource, sourceRoot) && CurrentPanePathEquals(folder, kDestination, destination) &&
               folder->DebugHasItemDisplayName(kSource, L"payload.bin") && folder->DebugHasItemDisplayName(kDestination, L"sentinel.bin");
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"deferred-exit pane listing"});
    if (! listed.conditionMet)
        return reject(L"the panes did not list the owned source and destination folders");
    const wil::com_ptr<IFileSystem> paneFileSystem = folder->GetFileSystem(kSource);
    if (! paneFileSystem)
        return reject(L"the source pane has no file system");

    // Batch Rename and a real Compare Directories window are opened before the exit is accepted, so
    // their later Rename and sync reach the host while it is pending. Compare owns its own FolderWindow
    // and engine, which must still honor the application window's pending exit.
    if (GetBatchRenameWindowHandle() != nullptr || GetCompareDirectoriesWindowHandle() != nullptr)
        return reject(L"a Batch Rename or Compare Directories window was already open");
    if (! folder->DebugFocusItemByDisplayName(kSource, L"payload.bin"))
        return reject(L"the payload could not be focused for Batch Rename");
    folder->CommandBatchRename(kSource);
    const auto batchOpened = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return GetBatchRenameWindowHandle() != nullptr;
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"deferred-exit Batch Rename open"});
    BatchRename::Rules batchRules{};
    batchRules.nameTemplate = L"renamed{ext}";
    if (! batchOpened.conditionMet || ! DebugSetBatchRenameWindowRules(batchRules))
        return reject(L"Batch Rename did not open for the payload or did not take its rules");
    const auto batchReady = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        static_cast<void>(DebugFlushBatchRenameWindowPendingPreview());
        BatchRenameDebugSnapshot snapshot{};
        return DebugGetBatchRenameWindowSnapshot(snapshot) && snapshot.renameButtonEnabled;
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"deferred-exit Batch Rename preview"});
    if (! batchReady.conditionMet)
        return reject(L"the Batch Rename preview never enabled Rename");

    SendMessageW(state.mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const auto compareOpened = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        compareWindow = GetCompareDirectoriesWindowHandle();
        return compareWindow != nullptr;
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"deferred-exit Compare open"});
    compareFolder = compareOpened.conditionMet ? DebugGetCompareDirectoriesFolderWindowForWindow(compareWindow) : nullptr;
    if (! compareFolder)
        return reject(L"Compare Directories did not open for the two panes");
    SendMessageW(compareWindow, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_RESCAN, 0), 0);
    const auto compared = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        CompareDirectoriesRunDebugSnapshot snapshot{};
        return DebugGetCompareDirectoriesRunSnapshotForWindow(compareWindow, snapshot) && snapshot.compareStarted && snapshot.compareActive &&
               ! snapshot.compareRunPending && compareFolder->DebugHasItemDisplayName(kSource, L"payload.bin") &&
               compareFolder->DebugHasItemDisplayName(kDestination, L"sentinel.bin");
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(15)), .operationName = L"deferred-exit Compare run"});
    if (! compared.conditionMet)
        return reject(L"the Compare Directories run did not settle on the owned folders");

    const auto exists = [](const std::filesystem::path& path) noexcept { return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES; };
    // A task held in its conflict metadata barrier owns the running slot and ignores its cancel
    // until released, so the exit stays pending for as long as the checks need it.
    const auto admitHeldCopy = [&]() noexcept -> std::optional<uint64_t>
    {
        SetFileOpsConflictMetadataPauseForSelfTest(true, static_cast<ULONGLONG>(SelfTest::Scale(std::chrono::milliseconds(60'000)).count()));
        const std::optional<uint64_t> taskId = StartFileOperationAndGetId(operations,
                                                                          FILESYSTEM_COPY,
                                                                          kSource,
                                                                          kDestination,
                                                                          state.fsLocal,
                                                                          {holdSource},
                                                                          holdDestination.parent_path(),
                                                                          FILESYSTEM_FLAG_NONE,
                                                                          false,
                                                                          0u,
                                                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                          false,
                                                                          state.fsLocal);
        if (! taskId)
            return std::nullopt;
        const auto held  = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return HasFileOpsConflictMetadataPauseEnteredForSelfTest();
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"deferred-exit held copy"});
        const auto* task = operations->FindTask(taskId.value());
        if (! held.conditionMet || ! task || ! task->HasEnteredOperation())
            return std::nullopt;
        return taskId;
    };
    // The resumed close posts WM_CLOSE to the target once the last task is reaped. The exit must stay
    // pending until that WM_CLOSE is handled, and end right after it so nothing later is blocked.
    const auto drainToClose = [&](HWND target, std::wstring_view name) noexcept
    {
        g_deferredExitPendingAtClose = -1;
        ReleaseFileOpsConflictMetadataPauseForSelfTest();
        SetFileOpsConflictMetadataPauseForSelfTest(false, 0u);
        const auto closed = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return ! operations->HasActiveOperations() && IsWindow(target) == FALSE && ! folder->IsFileOperationsExitPending();
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(30)), .operationName = L"deferred-exit close resumption"});
        AppendLog(std::format(L"Deferred exit {}: closed={} pendingAtClose={}", name, closed.conditionMet, g_deferredExitPendingAtClose));
        if (! closed.conditionMet)
            return reject(std::format(L"{}: the deferred close did not resume once the cancelled work was reaped", name));
        if (g_deferredExitPendingAtClose != 1)
            return reject(std::format(L"{}: the exit stopped refusing new work before its resumed WM_CLOSE was handled", name));
        return true;
    };

    // 1. With a cancelable task the exit consent is asked once; accepting arms the deferred close and
    //    cancels both the running task and the one waiting in the queue behind it.
    const std::optional<uint64_t> heldTaskId = admitHeldCopy();
    if (! heldTaskId)
        return reject(L"the held Copy did not reach its conflict barrier while running");
    const std::optional<uint64_t> queuedTaskId =
        StartFileOperationAndGetId(operations, FILESYSTEM_COPY, kSource, kDestination, state.fsLocal, {queuedSource}, queuedTarget, FILESYSTEM_FLAG_NONE, true);
    if (! queuedTaskId)
        return reject(L"the queued Copy was not admitted");
    const auto queued = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        const auto* task = operations->FindTask(queuedTaskId.value());
        return task && task->IsWaitingInQueue() && ! task->HasEnteredOperation();
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"deferred-exit queued copy"});
    if (! queued.conditionMet)
        return reject(L"the second Copy did not wait in the queue behind the running one");

    HostSetAutoAcceptPrompts(false);
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    uint64_t prompts = HostGetTestPromptRequestCount();
    if (folder->ConfirmCancelAllFileOperations(closeTargets[0]) || HostGetTestPromptRequestCount() != prompts + 1u || ! folder->IsFileOperationsExitPending() ||
        IsWindow(closeTargets[0]) == FALSE)
        return reject(L"accepting the exit consent did not ask exactly once and defer the close");
    const auto* heldTask   = operations->FindTask(heldTaskId.value());
    const auto* queuedTask = operations->FindTask(queuedTaskId.value());
    if (! heldTask || ! heldTask->CancelRequestedDurable() || (queuedTask && ! queuedTask->CancelRequestedDurable()))
        return reject(L"the accepted exit did not cancel the running and the queued Copy");

    // 2. The cancelled queued task ends without ever entering its operation.
    const auto queuedDone = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return state.completedTasks.contains(queuedTaskId.value()) && ! operations->FindTask(queuedTaskId.value());
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"deferred-exit queued copy cancellation"});
    if (! queuedDone.conditionMet || state.completedTasks[queuedTaskId.value()].hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) ||
        exists(queuedTarget / queuedSource.filename()))
        return reject(L"the cancelled queued Copy did not end canceled without writing its destination");
    if (! folder->IsFileOperationsExitPending() || IsWindow(closeTargets[0]) == FALSE || ! operations->HasActiveOperations())
        return reject(L"the close resumed while the held task was still owned");

    // 3. A second close attempt keeps the pending exit, does not ask again and shows the popup.
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    prompts = HostGetTestPromptRequestCount();
    if (folder->ConfirmCancelAllFileOperations(closeTargets[0]) || HostGetTestPromptRequestCount() != prompts)
        return reject(L"a close attempt while the exit was pending asked for consent again");
    const HWND popup = operations->GetPopupHwndForSelfTest();
    if (! folder->IsFileOperationsExitPending() || IsWindow(closeTargets[0]) == FALSE || ! popup || IsWindowVisible(popup) == FALSE)
        return reject(L"a second close attempt disarmed the pending exit or did not bring the popup forward");

    // 4. Every admission route is refused before any prompt. Prompts are accepted here, so a route
    //    that still reached its confirmation would be counted.
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    const auto taskIds = [&]() noexcept
    {
        std::vector<FolderWindow::FileOperationState::Task*> tasks;
        operations->CollectTasks(tasks);
        std::vector<uint64_t> ids;
        for (const auto* task : tasks)
            ids.push_back(task->GetId());
        std::ranges::sort(ids);
        return ids;
    };
    const auto breadcrumbFiles = [&]() noexcept
    {
        std::error_code ec;
        size_t count = 0u;
        for (auto it = std::filesystem::recursive_directory_iterator(breadcrumbRoot, ec); ! ec && it != std::filesystem::recursive_directory_iterator();
             it.increment(ec))
            count += it->is_regular_file(ec) ? 1u : 0u;
        return count;
    };
    const std::vector<uint64_t> liveIds = taskIds();
    const size_t breadcrumbsBefore      = breadcrumbFiles();
    const auto compareTaskCount         = [&]() noexcept
    {
        std::vector<FolderWindow::FileOperationState::Task*> tasks;
        if (auto* compareOperations = TryGetFileOps(compareFolder))
            compareOperations->CollectTasks(tasks);
        return tasks.size();
    };
    const auto dismissAlerts = [&](FolderWindow* owner) noexcept
    {
        owner->DismissPaneAlertOverlay(kSource);
        owner->DismissPaneAlertOverlay(kDestination);
        FolderView::AlertOverlayDebugSnapshot left{};
        FolderView::AlertOverlayDebugSnapshot right{};
        return owner->DebugGetPaneAlertSnapshot(kSource, left) && ! left.visible && owner->DebugGetPaneAlertSnapshot(kDestination, right) && ! right.visible;
    };
    modalPromptTimer = SetTimer(nullptr, 0u, 50u, &CancelDeferredExitModalPrompt);
    if (modalPromptTimer == 0u)
        return reject(L"the modal prompt watchdog timer could not be started");
    // alertFolder is the FolderWindow whose pane must show the refusal: the main one, or Compare's own.
    const auto attemptIn =
        [&](FolderWindow* alertFolder, std::wstring_view name, std::optional<FolderWindow::Pane> alertPane, UINT operationId, const auto& route) noexcept
    {
        FolderView::AlertOverlayDebugSnapshot alert{};
        const auto cleared = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return dismissAlerts(folder) && dismissAlerts(compareFolder);
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(2)), .operationName = L"deferred-exit alert dismissal"});
        if (! cleared.conditionMet || ! folder->DebugFocusItemByDisplayName(kSource, L"payload.bin"))
            return reject(std::format(L"{}: the pane alerts did not clear or the payload could not be focused", name));

        const uint64_t promptsBefore      = HostGetTestPromptRequestCount();
        const uint32_t modalPromptsBefore = g_deferredExitModalPromptCount.load(std::memory_order_acquire);
        const size_t batchWindowsBefore   = DebugGetBatchRenameWindowCount();
        // A modal prompt runs inside route() until the watchdog cancels it, so it is counted on return.
        const HRESULT hr = route();
        const uint64_t routePrompts =
            HostGetTestPromptRequestCount() - promptsBefore + (g_deferredExitModalPromptCount.load(std::memory_order_acquire) - modalPromptsBefore);
        bool alertShown = true;
        if (alertPane.has_value())
        {
            const auto shown                 = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return alertFolder->DebugGetPaneAlertSnapshot(alertPane.value(), alert) && alert.visible;
            }, {.timeout = SelfTest::Scale(std::chrono::seconds(2)), .operationName = L"deferred-exit refusal alert"});
            const std::wstring expectedTitle = LoadStringResource(nullptr, IDS_CAPTION_FILEOPS_EXIT);
            const std::wstring expectedMessage =
                FormatStringResource(nullptr, IDS_FMT_PANE_FILEOP_REFUSED_EXIT_PENDING, LoadStringResource(nullptr, operationId));
            alertShown = shown.conditionMet && alert.kind == FolderView::ErrorOverlayKind::Operation &&
                         alert.hr == FolderWindow::kFileOperationRefusedForPendingExit && alert.title == expectedTitle && alert.message == expectedMessage;
        }
        const bool untouched = exists(payload) && exists(sentinel) && ! exists(destination / payload.filename()) && ! exists(sourceRoot / L"renamed.bin");
        const size_t batchWindowsAfter = DebugGetBatchRenameWindowCount();
        AppendLog(std::format(L"Deferred exit refuses {}: hr=0x{:08X} prompts={} alertShown={} title='{}' message='{}' tasks={} (was {}) "
                              L"compareTasks={} batchWindows={} (was {}) breadcrumbs={} untouched={}",
                              name,
                              static_cast<unsigned long>(hr),
                              routePrompts,
                              alertShown,
                              alert.title,
                              alert.message,
                              taskIds().size(),
                              liveIds.size(),
                              compareTaskCount(),
                              batchWindowsAfter,
                              batchWindowsBefore,
                              breadcrumbFiles(),
                              untouched));
        if (routePrompts != 0u)
            return reject(std::format(L"{}: a refused admission showed {} prompt(s)", name, routePrompts));
        if (FAILED(hr))
            return reject(std::format(L"{}: the route did not report the exit refusal it was expected to return", name));
        if (! alertShown)
            return reject(std::format(L"{}: the refusal did not reach the pane as the localized exit alert", name));
        if (batchWindowsAfter != batchWindowsBefore)
            return reject(std::format(L"{}: a refused admission opened a Batch Rename window", name));
        if (taskIds() != liveIds || compareTaskCount() != 0u || breadcrumbFiles() != breadcrumbsBefore || ! untouched)
            return reject(std::format(L"{}: a refused admission published a task, wrote a breadcrumb or changed the owned files", name));
        return true;
    };
    const auto attempt = [&](std::wstring_view name, std::optional<FolderWindow::Pane> alertPane, UINT operationId, const auto& route) noexcept
    { return attemptIn(folder, name, alertPane, operationId, route); };
    const auto refusedHr = [&](std::wstring_view name, HRESULT hr, HRESULT expected) noexcept
    {
        if (hr == expected)
            return true;
        AppendLog(std::format(
            L"Deferred exit refuses {}: returned 0x{:08X}, expected 0x{:08X}", name, static_cast<unsigned long>(hr), static_cast<unsigned long>(expected)));
        return false;
    };
    const HRESULT kRefused = FolderWindow::kFileOperationRefusedForPendingExit;
    FolderView* destinationView =
        reinterpret_cast<FolderView*>(GetWindowLongPtrW(folder->GetFolderViewHwnd(kDestination), GWLP_USERDATA)); // FolderView stores itself there.
    if (! destinationView)
        return reject(L"the destination FolderView is unavailable");
    std::atomic<uint32_t> clipboardBarrierCalls{0u};
    const bool refused = attempt(L"pane Copy with options",
                                 kSource,
                                 IDS_FILEOP_OPERATION_COPY,
                                 [&]() noexcept
    {
        folder->CommandCopyToOtherPane(kSource, true);
        return S_OK;
    }) &&
                         attempt(L"pane Move",
                                 kSource,
                                 IDS_FILEOP_OPERATION_MOVE,
                                 [&]() noexcept
    {
        folder->CommandMoveToOtherPane(kSource, false);
        return S_OK;
    }) &&
                         attempt(L"pane Permanent delete",
                                 kSource,
                                 IDS_FILEOP_OPERATION_DELETE,
                                 [&]() noexcept
    {
        folder->CommandPermanentDelete(kSource);
        return S_OK;
    }) &&
                         attempt(L"drag and drop",
                                 kDestination,
                                 IDS_FILEOP_OPERATION_COPY,
                                 [&]() noexcept
    {
        DWORD performed  = DROPEFFECT_NONE;
        const HRESULT hr = destinationView->DebugPerformFileDropForSelfTest({payload}, DROPEFFECT_COPY, &performed);
        // The pane alert explains the refusal; FolderView must not replace it with a generic failure.
        return refusedHr(L"drag and drop", hr, S_OK) ? S_OK : E_FAIL;
    }) &&
                         attempt(L"clipboard Move admission",
                                 std::nullopt,
                                 IDS_FILEOP_OPERATION_MOVE,
                                 [&]() noexcept
    {
        uint64_t taskId  = 0u;
        const HRESULT hr = operations->AdmitOperation(FILESYSTEM_MOVE,
                                                      kSource,
                                                      kDestination,
                                                      paneFileSystem,
                                                      {payload},
                                                      destination,
                                                      FILESYSTEM_FLAG_RECURSIVE,
                                                      false,
                                                      0u,
                                                      FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                      false,
                                                      nullptr,
                                                      &taskId,
                                                      {},
                                                      {},
                                                      {},
                                                      {},
                                                      std::nullopt,
                                                      FileOperations::DeleteOrigin::PaneCommand,
                                                      std::nullopt,
                                                      0x51A7u,
                                                      [&]() noexcept
        {
            clipboardBarrierCalls.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        });
        return refusedHr(L"clipboard Move admission", hr, kRefused) && taskId == 0u && clipboardBarrierCalls.load(std::memory_order_acquire) == 0u ? S_OK
                                                                                                                                                   : E_FAIL;
    }) &&
                         attempt(L"inline rename admission",
                                 std::nullopt,
                                 IDS_FILEOP_OPERATION_RENAME,
                                 [&]() noexcept
    {
        uint64_t taskId  = 0u;
        const HRESULT hr = operations->AdmitInlineRename(kSource, paneFileSystem, payload, L"renamed.bin", &taskId);
        return refusedHr(L"inline rename admission", hr, kRefused) && taskId == 0u ? S_OK : E_FAIL;
    }) &&
                         attempt(L"Batch Rename admission",
                                 std::nullopt,
                                 IDS_FILEOP_OPERATION_RENAME,
                                 [&]() noexcept
    {
        BatchRenameExecutionOp rename{};
        rename.originalSource = payload;
        rename.finalLeaf      = L"renamed.bin";
        uint64_t taskId       = 0u;
        const HRESULT hr      = operations->AdmitBatchRename(kSource, paneFileSystem, {std::move(rename)}, {}, {}, &taskId);
        return refusedHr(L"Batch Rename admission", hr, kRefused) && taskId == 0u ? S_OK : E_FAIL;
    }) &&
                         attempt(L"Find Files permanent delete",
                                 kSource,
                                 IDS_FILEOP_OPERATION_DELETE,
                                 [&]() noexcept
    {
        uint64_t taskId  = 0u;
        const HRESULT hr = folder->StartFileOperationForResolvedPaths(
            kPluginIdLocal, std::wstring_view{}, FILESYSTEM_DELETE, {payload}, FILESYSTEM_FLAG_RECURSIVE, true, &taskId);
        return refusedHr(L"Find Files permanent delete", hr, kRefused) && taskId == 0u ? S_OK : E_FAIL;
    }) &&
                         attempt(L"Find Files copy to a folder",
                                 kSource,
                                 IDS_FILEOP_OPERATION_COPY,
                                 [&]() noexcept
    {
        uint64_t taskId  = 0u;
        const HRESULT hr = folder->StartFileOperationForResolvedPathsToDestination(
            kPluginIdLocal, std::wstring_view{}, FILESYSTEM_COPY, {payload}, destination, FILESYSTEM_FLAG_RECURSIVE, &taskId);
        return refusedHr(L"Find Files copy to a folder", hr, kRefused) && taskId == 0u ? S_OK : E_FAIL;
    }) &&
                         attempt(L"resolved items to the other pane",
                                 kSource,
                                 IDS_FILEOP_OPERATION_MOVE,
                                 [&]() noexcept
    {
        FolderWindow::ResolvedFileOperationItem item{};
        item.sourcePath      = payload;
        item.destinationPath = destination / payload.filename();
        std::optional<std::filesystem::path> destinationFolder;
        uint64_t taskId  = 0u;
        const HRESULT hr = folder->StartFileOperationForResolvedItemsToOtherPane(
            kPluginIdLocal, std::wstring_view{}, FILESYSTEM_MOVE, {std::move(item)}, &destinationFolder, &taskId, std::wstring{});
        return refusedHr(L"resolved items to the other pane", hr, kRefused) && taskId == 0u ? S_OK : E_FAIL;
    });
    if (! refused)
        return false;

    // Host commands and FolderView-owned entry points that would otherwise ask first: the rename
    // prompt, the Move folder picker, the Change Case prompt and the Batch Rename window.
    const HWND sourceViewHwnd = folder->GetFolderViewHwnd(kSource);
    if (! sourceViewHwnd)
        return reject(L"the source FolderView window is unavailable");
    const bool refusedBeforePrompts = attempt(L"pane Rename command",
                                              kSource,
                                              IDS_FILEOP_OPERATION_RENAME,
                                              [&]() noexcept
    {
        folder->CommandRename(kSource);
        return S_OK;
    }) &&
                                      attempt(L"context menu Rename",
                                              kSource,
                                              IDS_FILEOP_OPERATION_RENAME,
                                              [&]() noexcept
    {
        SendMessageW(sourceViewHwnd, WM_COMMAND, MAKEWPARAM(IDM_FOLDERVIEW_CONTEXT_RENAME, 0), 0);
        return S_OK;
    }) &&
                                      attempt(L"FolderView F2",
                                              kSource,
                                              IDS_FILEOP_OPERATION_RENAME,
                                              [&]() noexcept
    {
        SendMessageW(sourceViewHwnd, WM_KEYDOWN, VK_F2, 0);
        return S_OK;
    }) &&
                                      attempt(L"context menu Move",
                                              kSource,
                                              IDS_FILEOP_OPERATION_MOVE,
                                              [&]() noexcept
    {
        // The debug destination stands in for the folder picker: a refused Move must not take it.
        FolderView::DebugSetNextMoveSelectedItemsDestinationForSelfTest(destination);
        SendMessageW(sourceViewHwnd, WM_COMMAND, MAKEWPARAM(IDM_FOLDERVIEW_CONTEXT_MOVE, 0), 0);
        const bool pickerSkipped = FolderView::DebugIsNextMoveSelectedItemsDestinationPendingForSelfTest();
        FolderView::DebugSetNextMoveSelectedItemsDestinationForSelfTest(std::nullopt);
        if (! pickerSkipped)
            AppendLog(L"Deferred exit refuses context menu Move: the folder picker stand-in was consumed");
        return pickerSkipped ? S_OK : E_FAIL;
    }) &&
                                      attempt(L"Change Case command",
                                              kSource,
                                              IDS_FILEOP_OPERATION_RENAME,
                                              [&]() noexcept
    {
        folder->CommandChangeCase(kSource);
        return S_OK;
    }) &&
                                      attempt(L"Batch Rename execution",
                                              kSource,
                                              IDS_FILEOP_OPERATION_RENAME,
                                              [&]() noexcept
    {
        const HRESULT hr = DebugExecuteBatchRenameWindow();
        BatchRenameDebugSnapshot snapshot{};
        const bool reported = DebugGetBatchRenameWindowSnapshot(snapshot);
        AppendLog(std::format(L"Deferred exit refuses Batch Rename execution: failedRows={} canceled={} firstFailure=0x{:08X} text='{}'",
                              snapshot.lastExecutionFailedRows,
                              snapshot.lastExecutionCanceled,
                              static_cast<unsigned long>(snapshot.lastExecutionFirstFailure),
                              snapshot.lastExecutionFirstFailureText));
        // Nothing was attempted, so no row may be reported as failed.
        return refusedHr(L"Batch Rename execution", hr, kRefused) && reported && snapshot.lastExecutionFailedRows == 0u && snapshot.lastExecutionCanceled &&
                       snapshot.lastExecutionFirstFailure == kRefused
                   ? S_OK
                   : E_FAIL;
    });
    if (! refusedBeforePrompts)
        return false;
    closeOwnedTopLevel(GetBatchRenameWindowHandle());
    if (GetBatchRenameWindowHandle() != nullptr)
        return reject(L"the Batch Rename window did not close");
    const bool refusedSecondaryWindows = attempt(L"Batch Rename command",
                                                 kSource,
                                                 IDS_FILEOP_OPERATION_RENAME,
                                                 [&]() noexcept
    {
        folder->CommandBatchRename(kSource);
        return GetBatchRenameWindowHandle() == nullptr ? S_OK : E_FAIL;
    }) &&
                                         attemptIn(compareFolder,
                                                   L"Compare Directories sync",
                                                   kSource,
                                                   IDS_FILEOP_OPERATION_COPY,
                                                   [&]() noexcept
    {
        compareFolder->SetActivePane(kSource);
        if (! compareFolder->DebugFocusItemByDisplayName(kSource, L"payload.bin"))
            return E_FAIL;
        SendMessageW(compareWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_COPY_TO_OTHER, 0), 0);
        return S_OK;
    }) &&
                                         attemptIn(compareFolder,
                                                   L"Compare pane Permanent delete",
                                                   kSource,
                                                   IDS_FILEOP_OPERATION_DELETE,
                                                   [&]() noexcept
    {
        compareFolder->SetActivePane(kSource);
        if (! compareFolder->DebugFocusItemByDisplayName(kSource, L"payload.bin"))
            return E_FAIL;
        SendMessageW(compareWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_PERMANENT_DELETE, 0), 0);
        return S_OK;
    });
    if (! refusedSecondaryWindows)
        return false;

    // 5. Releasing the held task lets the pending exit resume, and the cancelled queued task never ran.
    if (! drainToClose(closeTargets[0], L"accepted exit"))
        return false;
    const auto heldDone = state.completedTasks.find(heldTaskId.value());
    if (heldDone == state.completedTasks.end() || SUCCEEDED(heldDone->second.hr) || exists(queuedTarget / queuedSource.filename()) || ! exists(holdDestination))
        return reject(L"the held Copy did not end as a cancellation, or the queued Copy wrote its destination");

    // 6. When every task is already Stopping there is nothing to consent to: the close is deferred
    //    without a prompt and resumes once the task is reaped.
    const std::optional<uint64_t> stoppingTaskId = admitHeldCopy();
    if (! stoppingTaskId)
        return reject(L"the second held Copy did not reach its conflict barrier while running");
    if (auto* stoppingTask = operations->FindTask(stoppingTaskId.value()))
        stoppingTask->RequestCancel();
    if (operations->HasCancelableOperations() || ! operations->HasActiveOperations())
        return reject(L"the cancelled held Copy was not the only, already Stopping, task");
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    prompts = HostGetTestPromptRequestCount();
    if (folder->ConfirmCancelAllFileOperations(closeTargets[1]) || HostGetTestPromptRequestCount() != prompts || ! folder->IsFileOperationsExitPending())
        return reject(L"an exit with only Stopping work asked for consent or did not defer the close");
    if (! drainToClose(closeTargets[1], L"all-Stopping exit"))
        return false;

    passed = true;
    return true;
}

// I30 a2-4: a Compare Directories window owns its own FolderWindow and File Operations engine.
// Closing it with work running asks "cancel and close, or keep running": keeping it leaves the task
// running; accepting cancels it through the normal cancel path, refuses new Compare work with the
// window-close alert and closes the window once its engine is quiet. The application exit consent
// counts the Compare task, "Cancel all and exit" cancels it, a second close brings the Compare
// engine's popup forward and the deferred close waits until that engine is quiet too.
[[nodiscard]] bool RunFileOpsCompareWindowCloseConsentChecks(SelfTestState& state) noexcept
{
    constexpr FolderWindow::Pane kSource      = FolderWindow::Pane::Left;
    constexpr FolderWindow::Pane kDestination = FolderWindow::Pane::Right;
    const auto reject                         = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Compare window close consent: {}", reason));
        return false;
    };

    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || state.tempRoot.empty() || ! CancelOutstandingSelfTestOperations(operations) ||
        folder->IsFileOperationsExitPending())
        return reject(L"live owner, Local file system, temp root or an idle operation state without a pending exit unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const auto root                   = state.tempRoot / L"compare-window-close-consent";
    const auto holdSource             = root / L"hold-src" / L"collision.bin";
    const auto holdDestination        = root / L"hold-dst" / L"collision.bin";
    const auto sourceRoot             = root / L"source";
    const auto destination            = root / L"destination";
    const auto payload                = sourceRoot / L"payload.bin";
    const auto sentinel               = destination / L"sentinel.bin";
    constexpr uintmax_t kHoldDestSize = 8192u;
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(sourceRoot) || ! RecreateEmptyDirectory(destination) ||
        ! WriteFilledTestFile(holdSource, 4096u, 0x11u) || ! WriteFilledTestFile(holdDestination, static_cast<size_t>(kHoldDestSize), 0x22u) ||
        ! WriteFilledTestFile(payload, 4096u, 0x5Au) || ! WriteFilledTestFile(sentinel, 512u, 0xA5u))
        return reject(L"owned scratch folders could not be created");

    const std::optional<std::filesystem::path> originalLeft  = folder->GetCurrentPath(kSource);
    const std::optional<std::filesystem::path> originalRight = folder->GetCurrentPath(kDestination);
    const HWND originalPopup                                 = operations->GetPopupHwndForSelfTest();
    const bool autoAccept                                    = HostGetAutoAcceptPrompts();
    const auto promptOverride                                = HostGetTestPromptResultOverride();
    // A message-only stand-in for the main window: the default WM_CLOSE handling destroys it, which
    // is how the resumed application close is observed.
    HWND closeTarget = CreateWindowExW(0, L"STATIC", L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, nullptr, nullptr);
    if (closeTarget != nullptr)
        g_deferredExitCloseTargetBaseProc = reinterpret_cast<WNDPROC>(
            RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(closeTarget, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(&RecordDeferredExitCloseTarget)));
    g_deferredExitCloseFolder   = folder;
    HWND compareWindow          = nullptr;
    FolderWindow* compareFolder = nullptr;
    // compareFolder belongs to the Compare window: never use it once that window is gone.
    const auto compareAlive      = [&]() noexcept { return compareWindow != nullptr && compareFolder != nullptr && IsWindow(compareWindow) != FALSE; };
    const auto compareOperations = [&]() noexcept { return compareAlive() ? TryGetFileOps(compareFolder) : nullptr; };
    bool passed                  = false;
    const auto cleanup           = wil::scope_exit([&]() noexcept
    {
        ReleaseFileOpsConflictMetadataPauseForSelfTest();
        SetFileOpsConflictMetadataPauseForSelfTest(false, 0u);
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        bool quiet = CancelOutstandingSelfTestOperations(operations);
        if (compareAlive())
        {
            quiet = CancelOutstandingSelfTestOperations(compareOperations()) && quiet;
            compareFolder->DismissPaneAlertOverlay(kSource);
            const HWND window = compareWindow;
            PostMessageW(window, WM_CLOSE, 0, 0);
            static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return IsWindow(window) == FALSE;
            }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"compare-close-consent window close"}));
        }
        compareFolder = nullptr;
        // With nothing active this disarms any exit the checks left pending, so later cases can admit.
        if (quiet && folder->IsFileOperationsExitPending())
            static_cast<void>(folder->CancelAllFileOperationsThenClose(nullptr));
        if (closeTarget != nullptr && IsWindow(closeTarget) != FALSE)
            DestroyWindow(closeTarget);
        g_deferredExitCloseFolder = nullptr;
        if (! originalPopup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        if (originalLeft.has_value())
            folder->SetFolderPath(kSource, originalLeft.value());
        if (originalRight.has_value())
            folder->SetFolderPath(kDestination, originalRight.value());
        if (! passed || ! quiet)
            state.retainTempRootOnFailure = true;
    });
    if (closeTarget == nullptr)
        return reject(L"the message-only close target could not be created");

    if ((folder->GetFileSystemPluginId(kSource) != kPluginIdLocal && FAILED(folder->SetFileSystemPluginForPane(kSource, kPluginIdLocal))) ||
        (folder->GetFileSystemPluginId(kDestination) != kPluginIdLocal && FAILED(folder->SetFileSystemPluginForPane(kDestination, kPluginIdLocal))))
        return reject(L"Local panes unavailable");
    folder->SetFolderPath(kSource, sourceRoot);
    folder->SetFolderPath(kDestination, destination);
    const auto listed = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        return CurrentPanePathEquals(folder, kSource, sourceRoot) && CurrentPanePathEquals(folder, kDestination, destination) &&
               folder->DebugHasItemDisplayName(kSource, L"payload.bin") && folder->DebugHasItemDisplayName(kDestination, L"sentinel.bin");
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"compare-close-consent pane listing"});
    if (! listed.conditionMet)
        return reject(L"the panes did not list the owned source and destination folders");

    if (GetCompareDirectoriesWindowHandle() != nullptr)
        return reject(L"a Compare Directories window was already open");
    SendMessageW(state.mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const auto compareOpened = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        compareWindow = GetCompareDirectoriesWindowHandle();
        return compareWindow != nullptr;
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"compare-close-consent Compare open"});
    compareFolder = compareOpened.conditionMet ? DebugGetCompareDirectoriesFolderWindowForWindow(compareWindow) : nullptr;
    if (! compareFolder)
        return reject(L"Compare Directories did not open for the two panes");
    SendMessageW(compareWindow, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_RESCAN, 0), 0);
    const auto compared = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        CompareDirectoriesRunDebugSnapshot snapshot{};
        return DebugGetCompareDirectoriesRunSnapshotForWindow(compareWindow, snapshot) && snapshot.compareStarted && snapshot.compareActive &&
               ! snapshot.compareRunPending && compareFolder->DebugHasItemDisplayName(kSource, L"payload.bin") &&
               compareFolder->DebugHasItemDisplayName(kDestination, L"sentinel.bin");
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(15)), .operationName = L"compare-close-consent Compare run"});
    if (! compared.conditionMet || ! compareOperations())
        return reject(L"the Compare Directories run did not settle on the owned folders");

    const auto exists              = [](const std::filesystem::path& path) noexcept { return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES; };
    const auto holdDestinationKept = [&]() noexcept
    {
        std::error_code ec;
        return std::filesystem::file_size(holdDestination, ec) == kHoldDestSize && ! ec;
    };
    // A Compare engine task held in its conflict metadata barrier owns the running slot and ignores
    // its cancel until released, so a close stays pending for as long as the checks need it.
    const auto admitHeldCompareCopy = [&]() noexcept -> std::optional<uint64_t>
    {
        auto* compareOps = compareOperations();
        if (! compareOps)
            return std::nullopt;
        SetFileOpsConflictMetadataPauseForSelfTest(true, static_cast<ULONGLONG>(SelfTest::Scale(std::chrono::milliseconds(60'000)).count()));
        const std::optional<uint64_t> taskId = StartFileOperationAndGetId(compareOps,
                                                                          FILESYSTEM_COPY,
                                                                          kSource,
                                                                          kDestination,
                                                                          state.fsLocal,
                                                                          {holdSource},
                                                                          holdDestination.parent_path(),
                                                                          FILESYSTEM_FLAG_NONE,
                                                                          false,
                                                                          0u,
                                                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                          false,
                                                                          state.fsLocal);
        if (! taskId)
            return std::nullopt;
        const auto held  = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return HasFileOpsConflictMetadataPauseEnteredForSelfTest();
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"compare-close-consent held copy"});
        const auto* task = compareOps->FindTask(taskId.value());
        if (! held.conditionMet || ! task || ! task->HasEnteredOperation())
            return std::nullopt;
        return taskId;
    };
    const auto lastPromptIs = [](UINT titleId, UINT messageId) noexcept
    {
        HostPromptDebugSnapshot prompt{};
        const bool matches = HostGetTestLastPromptDebugSnapshot(prompt) && prompt.title == LoadStringResource(nullptr, titleId) &&
                             prompt.message == LoadStringResource(nullptr, messageId);
        AppendLog(std::format(L"Compare window close consent: last prompt title='{}' message='{}' matches={}", prompt.title, prompt.message, matches));
        return matches;
    };

    // 1. Closing the Compare window with its own task running asks the window-close consent, and
    //    "keep running" leaves the window open and the task running.
    const std::optional<uint64_t> firstTaskId = admitHeldCompareCopy();
    if (! firstTaskId)
        return reject(L"the Compare engine's Copy did not reach its conflict barrier while running");
    if (operations->HasActiveOperations())
        return reject(L"the main engine was not idle while only the Compare engine held a task");
    HostSetAutoAcceptPrompts(false);
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    uint64_t prompts = HostGetTestPromptRequestCount();
    SendMessageW(compareWindow, WM_CLOSE, 0, 0);
    uint64_t asked = HostGetTestPromptRequestCount() - prompts;
    AppendLog(std::format(L"Compare window close consent keep running: prompts={} windowAlive={}", asked, IsWindow(compareWindow) != FALSE));
    if (! compareAlive())
        return reject(L"closing the Compare window with a running task destroyed it without asking");
    if (asked != 1u || ! lastPromptIs(IDS_CAPTION_FILEOPS_CLOSE_WINDOW, IDS_MSG_FILEOPS_CANCEL_ALL_CLOSE_WINDOW))
        return reject(L"closing the Compare window with a running task did not ask its cancel-and-close consent exactly once");
    const auto* firstTask = compareOperations()->FindTask(firstTaskId.value());
    if (! firstTask || firstTask->CancelRequestedDurable() || compareFolder->IsFileOperationsExitPending() || folder->IsFileOperationsExitPending())
        return reject(L"\"keep running\" canceled the Compare task or left a close pending");

    // 2. The application exit consent counts the Compare task although the main engine is idle, and
    //    declining it changes nothing.
    prompts = HostGetTestPromptRequestCount();
    if (folder->ConfirmCancelAllFileOperations(closeTarget) || HostGetTestPromptRequestCount() != prompts + 1u ||
        ! lastPromptIs(IDS_CAPTION_FILEOPS_EXIT, IDS_MSG_FILEOPS_CANCEL_ALL_EXIT))
        return reject(L"the application exit consent did not count the Compare engine's running task");
    firstTask = compareOperations()->FindTask(firstTaskId.value());
    if (! firstTask || firstTask->CancelRequestedDurable() || folder->IsFileOperationsExitPending() || IsWindow(closeTarget) == FALSE)
        return reject(L"a declined application exit canceled the Compare task or closed the application");

    // 3. "Cancel all and exit" cancels the Compare task and defers the application close.
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    prompts = HostGetTestPromptRequestCount();
    if (folder->ConfirmCancelAllFileOperations(closeTarget) || HostGetTestPromptRequestCount() != prompts + 1u || ! folder->IsFileOperationsExitPending() ||
        ! compareFolder->IsFileOperationsExitPending() || IsWindow(closeTarget) == FALSE)
        return reject(L"accepting the application exit did not ask once and defer the close");
    firstTask = compareOperations()->FindTask(firstTaskId.value());
    if (! firstTask || ! firstTask->CancelRequestedDurable())
        return reject(L"\"Cancel all and exit\" did not cancel the Compare engine's task");

    // 4. A second application close does not ask again and brings the Compare engine's popup forward,
    //    since only that engine still has work. A Compare window close while the exit is pending keeps
    //    the window open without asking.
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    if (const HWND shownPopup = compareOperations()->GetPopupHwndForSelfTest())
        ShowWindow(shownPopup, SW_HIDE);
    const HWND mainPopupBefore     = operations->GetPopupHwndForSelfTest();
    const bool mainPopupWasVisible = mainPopupBefore != nullptr && IsWindowVisible(mainPopupBefore) != FALSE;
    prompts                        = HostGetTestPromptRequestCount();
    if (folder->ConfirmCancelAllFileOperations(closeTarget) || HostGetTestPromptRequestCount() != prompts || ! folder->IsFileOperationsExitPending())
        return reject(L"a second application close while the exit was pending asked again or disarmed it");
    const HWND comparePopup = compareOperations()->GetPopupHwndForSelfTest();
    const HWND mainPopup    = operations->GetPopupHwndForSelfTest();
    AppendLog(std::format(L"Compare window close consent second close: comparePopupVisible={} mainPopupVisible={} (was {})",
                          comparePopup != nullptr && IsWindowVisible(comparePopup) != FALSE,
                          mainPopup != nullptr && IsWindowVisible(mainPopup) != FALSE,
                          mainPopupWasVisible));
    if (! comparePopup || IsWindowVisible(comparePopup) == FALSE)
        return reject(L"a second application close did not bring the Compare engine's popup forward");
    SendMessageW(compareWindow, WM_CLOSE, 0, 0);
    if (! compareAlive() || HostGetTestPromptRequestCount() != prompts)
        return reject(L"closing the Compare window during the pending application exit asked or closed it while its task was still stopping");

    // 5. The deferred application close waits for the Compare engine, then resumes; the exit stays
    //    pending until its WM_CLOSE is handled.
    const bool waited = RedSalamander::TestSupport::PumpMessagesUntil(
                            [&]() noexcept { return IsWindow(closeTarget) == FALSE; },
                            {.timeout = SelfTest::Scale(std::chrono::milliseconds(500)), .operationName = L"compare-close-consent premature close"})
                            .conditionMet == false;
    if (! waited || ! folder->IsFileOperationsExitPending())
        return reject(L"the deferred application close resumed while the Compare engine still owned its task");
    g_deferredExitPendingAtClose = -1;
    ReleaseFileOpsConflictMetadataPauseForSelfTest();
    SetFileOpsConflictMetadataPauseForSelfTest(false, 0u);
    const auto exitClosed = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        auto* compareOps = compareOperations();
        return compareOps && ! compareOps->HasActiveOperations() && IsWindow(closeTarget) == FALSE && ! folder->IsFileOperationsExitPending();
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(30)), .operationName = L"compare-close-consent application close resumption"});
    AppendLog(
        std::format(L"Compare window close consent application exit: closed={} pendingAtClose={}", exitClosed.conditionMet, g_deferredExitPendingAtClose));
    if (! exitClosed.conditionMet)
        return reject(L"the deferred application close did not resume once the Compare engine was quiet");
    if (g_deferredExitPendingAtClose != 1)
        return reject(L"the application exit stopped refusing new work before its resumed WM_CLOSE was handled");
    if (! compareAlive() || ! holdDestinationKept())
        return reject(L"the application exit closed the Compare window itself or the canceled Copy changed its destination");

    // 6. Accepting the Compare window's own consent cancels its task, refuses new Compare work with
    //    the window-close alert while the application stays usable, and closes the window once quiet.
    const std::optional<uint64_t> secondTaskId = admitHeldCompareCopy();
    if (! secondTaskId)
        return reject(L"the second Compare engine Copy did not reach its conflict barrier while running");
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    prompts = HostGetTestPromptRequestCount();
    SendMessageW(compareWindow, WM_CLOSE, 0, 0);
    if (! compareAlive() || HostGetTestPromptRequestCount() != prompts + 1u ||
        ! lastPromptIs(IDS_CAPTION_FILEOPS_CLOSE_WINDOW, IDS_MSG_FILEOPS_CANCEL_ALL_CLOSE_WINDOW))
        return reject(L"accepting the Compare window close did not ask once and keep the window until its task stopped");
    const auto* secondTask = compareOperations()->FindTask(secondTaskId.value());
    if (! secondTask || ! secondTask->CancelRequestedDurable() || ! compareFolder->IsFileOperationsExitPending() || folder->IsFileOperationsExitPending())
        return reject(L"the accepted Compare window close did not cancel its task, refuse Compare work, or left the application refusing work");

    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    compareFolder->DismissPaneAlertOverlay(kSource);
    compareFolder->SetActivePane(kSource);
    if (! compareFolder->DebugFocusItemByDisplayName(kSource, L"payload.bin"))
        return reject(L"the payload could not be focused in the Compare pane");
    prompts = HostGetTestPromptRequestCount();
    SendMessageW(compareWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_PERMANENT_DELETE, 0), 0);
    FolderView::AlertOverlayDebugSnapshot alert{};
    const auto alertShown            = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return compareAlive() && compareFolder->DebugGetPaneAlertSnapshot(kSource, alert) && alert.visible;
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(2)), .operationName = L"compare-close-consent refusal alert"});
    const std::wstring expectedTitle = LoadStringResource(nullptr, IDS_CAPTION_FILEOPS_CLOSE_WINDOW);
    const std::wstring expectedMessage =
        FormatStringResource(nullptr, IDS_FMT_PANE_FILEOP_REFUSED_CLOSE_PENDING, LoadStringResource(nullptr, IDS_FILEOP_OPERATION_DELETE));
    std::vector<FolderWindow::FileOperationState::Task*> compareTasks;
    if (auto* compareOps = compareOperations())
        compareOps->CollectTasks(compareTasks);
    AppendLog(std::format(L"Compare window close consent refusal: prompts={} alertShown={} hr=0x{:08X} title='{}' message='{}' compareTasks={} payload={}",
                          HostGetTestPromptRequestCount() - prompts,
                          alertShown.conditionMet,
                          static_cast<unsigned long>(alert.hr),
                          alert.title,
                          alert.message,
                          compareTasks.size(),
                          exists(payload)));
    if (HostGetTestPromptRequestCount() != prompts || ! alertShown.conditionMet || alert.kind != FolderView::ErrorOverlayKind::Operation ||
        alert.hr != FolderWindow::kFileOperationRefusedForPendingExit || alert.title != expectedTitle || alert.message != expectedMessage)
        return reject(L"a Compare pane Delete during the accepted window close was not refused with the window-close alert before any prompt");
    if (compareTasks.size() != 1u || ! exists(payload))
        return reject(L"a refused Compare pane Delete published a task or deleted the payload");

    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    prompts = HostGetTestPromptRequestCount();
    SendMessageW(compareWindow, WM_CLOSE, 0, 0);
    if (! compareAlive() || HostGetTestPromptRequestCount() != prompts)
        return reject(L"a second Compare window close while its close was pending asked again or closed it early");

    const HWND closingCompareWindow = compareWindow;
    ReleaseFileOpsConflictMetadataPauseForSelfTest();
    SetFileOpsConflictMetadataPauseForSelfTest(false, 0u);
    const auto compareClosed = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return IsWindow(closingCompareWindow) == FALSE;
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(30)), .operationName = L"compare-close-consent Compare window close resumption"});
    if (! compareClosed.conditionMet)
        return reject(L"the accepted Compare window close did not resume once its engine was quiet");
    compareWindow = nullptr;
    compareFolder = nullptr;
    if (! holdDestinationKept() || folder->IsFileOperationsExitPending() || operations->HasActiveOperations())
        return reject(L"the canceled Compare Copy changed its destination or the application was left refusing work");

    passed = true;
    return true;
}

constexpr std::string_view kLocalManualCopyConcurrency4ForSelfTest =
    R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve"})json";
constexpr int64_t kSelfTestDayTicks = 864'000'000'000;
// 2001-01-01 00:00:00 UTC in FILETIME ticks: directory times far from any copy time.
constexpr int64_t kSelfTestYear2001Ticks = 126'227'808'000'000'000;

// A directory's last-write time in FILETIME ticks (-1 when it cannot be read).
[[nodiscard]] int64_t DirectoryLastWriteTimeForSelfTest(const std::filesystem::path& directory) noexcept
{
    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (GetFileAttributesExW(MakeExtendedPathForSelfTest(directory).c_str(), GetFileExInfoStandard, &data) == FALSE)
    {
        return -1;
    }
    return static_cast<int64_t>((static_cast<uint64_t>(data.ftLastWriteTime.dwHighDateTime) << 32u) | data.ftLastWriteTime.dwLowDateTime);
}

[[nodiscard]] bool SetDirectoryLastWriteTimeForSelfTest(const std::filesystem::path& directory, int64_t lastWriteTime) noexcept
{
    wil::unique_hfile handle(CreateFileW(MakeExtendedPathForSelfTest(directory).c_str(),
                                         FILE_WRITE_ATTRIBUTES,
                                         FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                         nullptr,
                                         OPEN_EXISTING,
                                         FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                         nullptr));
    const uint64_t ticks = static_cast<uint64_t>(lastWriteTime);
    const FILETIME time{static_cast<DWORD>(ticks & 0xFFFFFFFFu), static_cast<DWORD>(ticks >> 32u)};
    return handle && SetFileTime(handle.get(), nullptr, nullptr, &time) != FALSE;
}

// Every entry below `directory` as a relative path (directories end with a separator), sorted ignoring case;
// false when the tree cannot be listed.
[[nodiscard]] bool CollectRelativeTreeForSelfTest(const std::filesystem::path& directory, std::vector<std::wstring>& entries) noexcept
{
    entries.clear();
    const std::filesystem::path extended(MakeExtendedPathForSelfTest(directory));
    const size_t prefixLength = extended.native().size();
    std::error_code ec;
    for (std::filesystem::recursive_directory_iterator it(extended, std::filesystem::directory_options::none, ec);
         ! ec && it != std::filesystem::recursive_directory_iterator{};
         it.increment(ec))
    {
        std::wstring relative = it->path().native().substr((std::min)(prefixLength, it->path().native().size()));
        while (! relative.empty() && relative.front() == L'\\')
        {
            relative.erase(relative.begin());
        }
        std::error_code kindError;
        if (it->is_directory(kindError))
        {
            relative.push_back(L'\\');
        }
        if (kindError)
        {
            return false;
        }
        entries.push_back(std::move(relative));
    }
    if (ec)
    {
        return false;
    }
    std::ranges::sort(entries, [](const std::wstring& left, const std::wstring& right) noexcept { return _wcsicmp(left.c_str(), right.c_str()) < 0; });
    return true;
}

// A folder of the many-folders fixture: `relative` below the tree root, the root itself when empty.
[[nodiscard]] std::filesystem::path ManyDirectoriesFolderForSelfTest(const std::filesystem::path& treeRoot, const std::wstring& relative)
{
    return relative.empty() ? treeRoot : treeRoot / relative;
}

// The many-folders fixture gives folder i the last-write time 2001 + i seconds.
[[nodiscard]] constexpr int64_t ManyDirectoriesFolderTimeForSelfTest(size_t index) noexcept
{
    return kSelfTestYear2001Ticks + static_cast<int64_t>(index) * 10'000'000;
}

// The registered SMB self-test root (path.smb.primary: REDSALAMANDER_SELFTEST_SMB_ROOT) when it is a UNC directory below
// the share with a 'selftest' segment and no dot segment and it can be reached; otherwise nullopt, with the reason in
// `unavailableReason`. A case creates, and removes, only one unique child of it.
[[nodiscard]] std::optional<std::wstring> ResolveSmbSelfTestRootForSelfTest(std::wstring& unavailableReason) noexcept
{
    unavailableReason.clear();
    std::wstring smbRootText = GetEnvVarTrimmed(L"REDSALAMANDER_SELFTEST_SMB_ROOT");
    std::ranges::replace(smbRootText, L'/', L'\\');
    while (smbRootText.size() > 2u && smbRootText.back() == L'\\')
    {
        smbRootText.pop_back();
    }
    if (smbRootText.empty())
    {
        unavailableReason = L"path.smb.primary is not configured";
        return std::nullopt;
    }
    size_t segmentCount     = 0u;
    bool hasSelfTestSegment = false;
    bool hasDotSegment      = false;
    for (size_t index = 2u; index < smbRootText.size();)
    {
        while (index < smbRootText.size() && smbRootText[index] == L'\\')
        {
            ++index;
        }
        const size_t start = index;
        while (index < smbRootText.size() && smbRootText[index] != L'\\')
        {
            ++index;
        }
        if (start == index)
        {
            continue;
        }
        const std::wstring_view segment = std::wstring_view(smbRootText).substr(start, index - start);
        ++segmentCount;
        hasDotSegment      = hasDotSegment || segment == L"." || segment == L"..";
        hasSelfTestSegment = hasSelfTestSegment || ContainsIgnoreCase(segment, L"selftest");
    }
    if (! smbRootText.starts_with(L"\\\\") || segmentCount < 3u || hasDotSegment || ! hasSelfTestSegment)
    {
        unavailableReason = L"path.smb.primary must be a UNC directory below the share with a 'selftest' segment and no dot segment";
        return std::nullopt;
    }
    const DWORD rootAttributes = GetFileAttributesW(smbRootText.c_str());
    const DWORD rootError      = rootAttributes == INVALID_FILE_ATTRIBUTES ? GetLastError() : ERROR_SUCCESS;
    if (rootAttributes == INVALID_FILE_ATTRIBUTES || (rootAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
    {
        unavailableReason = std::format(L"path.smb.primary cannot be reached or is not a directory (error {})", rootError);
        return std::nullopt;
    }
    return smbRootText;
}

// The nested fixture of the SMB cases: four folder levels with files at every level, an empty folder and a folder of
// 300 small files. Each folder gets its own last-write time in 2001 once its contents exist, so a folder whose basic
// information was not restored after its children were copied shows the copy time instead.
struct SmbTreeFixtureDirectory final
{
    const wchar_t* relative;
    unsigned int files;
};
constexpr std::array<SmbTreeFixtureDirectory, 7> kSmbTreeFixtureDirectories{{
    {L"", 2u},
    {L"empty", 0u},
    {L"level2", 3u},
    {L"level2\\level3", 2u},
    {L"level2\\level3\\level4", 2u},
    {L"level2\\sibling", 1u},
    {L"many", 300u},
}};
// The flat folder of the SMB bridge cases: 40 files of 1 to 4 KiB and its own 2001 time.
constexpr unsigned int kSmbFlatFixtureFiles = 40u;
constexpr int64_t kSmbFlatFixtureTime       = kSelfTestYear2001Ticks + 40 * kSelfTestDayTicks;
// A share may keep times more coarsely than NTFS; two seconds covers FAT-style resolution.
constexpr int64_t kSmbLastWriteToleranceTicks = 20'000'000;

[[nodiscard]] std::filesystem::path SmbTreeFixturePath(const std::filesystem::path& root, size_t index)
{
    const wchar_t* const relative = kSmbTreeFixtureDirectories[index].relative;
    return *relative == L'\0' ? root : root / relative;
}

[[nodiscard]] constexpr int64_t SmbTreeFixtureTime(size_t index) noexcept
{
    return kSelfTestYear2001Ticks + static_cast<int64_t>(index) * kSelfTestDayTicks;
}

// Writes the nested fixture at `root`; returns its file count, or nullopt.
[[nodiscard]] std::optional<size_t> WriteSmbTreeFixtureForSelfTest(const std::filesystem::path& root) noexcept
{
    size_t files = 0u;
    for (size_t index = 0u; index < kSmbTreeFixtureDirectories.size(); ++index)
    {
        const std::filesystem::path path = SmbTreeFixturePath(root, index);
        std::error_code createError;
        static_cast<void>(std::filesystem::create_directories(path, createError));
        if (createError)
        {
            return std::nullopt;
        }
        for (unsigned int file = 0u; file < kSmbTreeFixtureDirectories[index].files; ++file)
        {
            const size_t bytes = 512u + (files % 7u) * 129u;
            if (! WriteFilledTestFile(path / std::format(L"file-{:03}.bin", file), bytes, static_cast<unsigned char>(0x21u + files % 90u)))
            {
                return std::nullopt;
            }
            ++files;
        }
    }
    for (size_t index = 0u; index < kSmbTreeFixtureDirectories.size(); ++index)
    {
        if (! SetDirectoryLastWriteTimeForSelfTest(SmbTreeFixturePath(root, index), SmbTreeFixtureTime(index)))
        {
            return std::nullopt;
        }
    }
    return files;
}

[[nodiscard]] bool WriteSmbFlatFixtureForSelfTest(const std::filesystem::path& root) noexcept
{
    for (unsigned int file = 0u; file < kSmbFlatFixtureFiles; ++file)
    {
        if (! WriteFilledTestFile(root / std::format(L"flat-{:02}.bin", file), 1024u + file * 77u, static_cast<unsigned char>(0x30u + file)))
        {
            return false;
        }
    }
    return SetDirectoryLastWriteTimeForSelfTest(root, kSmbFlatFixtureTime);
}

// True when the folder's last-write time is within the share tolerance of `expected`.
[[nodiscard]] bool LastWriteTimeNearForSelfTest(const std::filesystem::path& directory, int64_t expected) noexcept
{
    const int64_t actual = DirectoryLastWriteTimeForSelfTest(directory);
    return actual >= 0 && (actual >= expected ? actual - expected : expected - actual) <= kSmbLastWriteToleranceTicks;
}

// Empty when `copied` holds exactly the tree at `expected` (the same entries with the same bytes, and no stage or other
// operation artifact); otherwise what differs.
[[nodiscard]] std::wstring DescribeTreeDifferenceForSelfTest(const std::filesystem::path& expected, const std::filesystem::path& copied) noexcept
{
    std::vector<std::wstring> expectedEntries;
    std::vector<std::wstring> copiedEntries;
    if (! CollectRelativeTreeForSelfTest(expected, expectedEntries) || ! CollectRelativeTreeForSelfTest(copied, copiedEntries))
    {
        return std::format(L"'{}' or '{}' could not be listed", expected.native(), copied.native());
    }
    const auto leftover = std::ranges::find_if(copiedEntries, [](const std::wstring& entry) noexcept { return ContainsIgnoreCase(entry, L".rs_"); });
    if (leftover != copiedEntries.end())
    {
        return std::format(L"'{}' kept a stage or another operation artifact: '{}'", copied.filename().native(), *leftover);
    }
    if (! std::ranges::equal(
            copiedEntries, expectedEntries, [](const std::wstring& left, const std::wstring& right) noexcept { return EqualsIgnoreCase(left, right); }))
    {
        return std::format(L"'{}' differs from its source: {} entries, {} expected", copied.filename().native(), copiedEntries.size(), expectedEntries.size());
    }
    for (const std::wstring& entry : expectedEntries)
    {
        if (! entry.ends_with(L'\\') && ! FilesEqualBytes(expected / entry, copied / entry))
        {
            return std::format(L"'{}\\{}' differs from its source", copied.filename().native(), entry);
        }
    }
    return {};
}

// One task's diagnostics.
[[nodiscard]] std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> CollectTaskDiagnosticsForSelfTest(FolderWindow::FileOperationState* operations,
                                                                                                                   uint64_t taskId) noexcept
{
    std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> entries;
    operations->CollectDiagnostics(entries);
    std::erase_if(entries, [taskId](const FolderWindow::FileOperationState::TaskDiagnosticEntry& entry) noexcept { return entry.taskId != taskId; });
    return entries;
}

[[nodiscard]] bool HasDiagnosticCategoryForSelfTest(const std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry>& entries,
                                                    std::wstring_view category) noexcept
{
    return std::ranges::any_of(entries,
                               [category](const FolderWindow::FileOperationState::TaskDiagnosticEntry& entry) noexcept { return entry.category == category; });
}

[[nodiscard]] std::wstring DescribeTaskDiagnosticsForSelfTest(const std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry>& entries) noexcept
{
    std::wstring text;
    for (size_t index = 0u; index < entries.size() && index < 8u; ++index)
    {
        const FolderWindow::FileOperationState::TaskDiagnosticEntry& entry = entries[index];
        text += std::format(L" [{} 0x{:08X} '{}' {}]", entry.category, static_cast<unsigned long>(entry.status), entry.destinationPath, entry.message);
    }
    return text;
}

// I35 NAS nested Copy (owner report 2026-10-12: copying a folder with subfolders to the NAS fails): a
// host-admitted Local Copy of a nested tree to a unique child of the registered SMB self-test root
// (path.smb.primary), through the Local provider's own walker. The walker used to keep every folder it
// published open, with DELETE access, until the folder's children finished; a server that refuses every
// rename inside a folder held that way (the Samba NAS probed on 2026-10-12) failed each nested publication
// and the task stopped with 0x800703E4. Environment-gated: skipped, with the reason, when the root is not
// configured, is not a dedicated self-test UNC directory, or cannot be reached.
[[nodiscard]] bool RunFileOpsSmbNestedTreeCopyChecks(SelfTestState& state, std::wstring& unavailableReason) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"SMB nested tree Copy: {}", reason));
        return false;
    };
    const std::optional<std::wstring> smbRoot = ResolveSmbSelfTestRootForSelfTest(unavailableReason);
    if (! smbRoot.has_value())
    {
        return false;
    }

    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
    {
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        return reject(L"no-activation guard unavailable");
    }

    const std::filesystem::path localBase = state.tempRoot / L"smb-nested-copy";
    const std::filesystem::path source    = localBase / L"tree";
    const std::filesystem::path shareBase = std::filesystem::path(smbRoot.value()) / std::format(L"nested-copy-{}-{}", GetCurrentProcessId(), GetTickCount64());
    const std::filesystem::path destination = shareBase / L"dst";
    const std::filesystem::path copied      = destination / source.filename();
    const bool verify                       = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
    std::optional<uint64_t> taskId;
    const auto restore = wil::scope_exit([&]() noexcept
    {
        const bool quiet = CancelOutstandingSelfTestOperations(operations);
        if (quiet && taskId.has_value())
        {
            operations->DismissCompletedTask(taskId.value());
            operations->DebugRemoveDiagnosticsForTask(taskId.value());
            state.completedTasks.erase(taskId.value());
        }
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verify;
        // Only this run's unique child of the share root is removed.
        static_cast<void>(SelfTest::RemoveAll(shareBase));
        static_cast<void>(SelfTest::RemoveAll(localBase));
    });

    if (! RecreateEmptyDirectory(source))
    {
        return reject(L"local source root unavailable");
    }
    const std::optional<size_t> sourceFiles = WriteSmbTreeFixtureForSelfTest(source);
    if (! sourceFiles.has_value())
    {
        return reject(L"local source fixture unavailable");
    }
    if (! RecreateEmptyDirectory(destination))
    {
        return reject(std::format(L"share destination '{}' unavailable", destination.native()));
    }

    // Verify Off keeps the provider's own route: the Local walker publishes every folder itself.
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    const auto startedAt                                      = std::chrono::steady_clock::now();
    taskId                                                    = StartFileOperationAndGetId(operations,
                                                                                           FILESYSTEM_COPY,
                                                                                           FolderWindow::Pane::Left,
                                                                                           FolderWindow::Pane::Right,
                                                                                           state.fsLocal,
                                                                                           {source},
                                                                                           destination,
                                                                                           FILESYSTEM_FLAG_RECURSIVE,
                                                                                           false);
    if (! taskId.has_value())
    {
        return reject(L"the Copy to the share was not admitted");
    }
    // A failure the walker cannot go past asks for a decision; nothing in this Copy may get there, so a
    // decision ends the case at once with what the task reported.
    std::optional<FolderWindow::FileOperationState::Task::ConflictPromptState> prompt;
    const auto finished = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        if (FolderWindow::FileOperationState::Task* task = operations->FindTask(taskId.value()); task != nullptr)
        {
            prompt = TryGetConflictPromptCopy(task);
        }
        return prompt.has_value() || (state.completedTasks.contains(taskId.value()) && operations->FindTask(taskId.value()) == nullptr);
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(240)), .operationName = L"SMB nested tree Copy"});
    const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
    const auto diagnostics   = CollectTaskDiagnosticsForSelfTest(operations, taskId.value());
    if (prompt.has_value())
    {
        return reject(std::format(L"the Copy to the share stopped on a decision: status=0x{:08X} bucket={} source='{}' destination='{}'{}",
                                  static_cast<unsigned long>(prompt->status),
                                  static_cast<unsigned int>(prompt->bucket),
                                  prompt->sourcePath,
                                  prompt->destinationPath,
                                  DescribeTaskDiagnosticsForSelfTest(diagnostics)));
    }
    const auto completed = state.completedTasks.find(taskId.value());
    if (! finished.conditionMet || completed == state.completedTasks.end())
    {
        return reject(std::format(L"the Copy to the share did not finish{}", DescribeTaskDiagnosticsForSelfTest(diagnostics)));
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.SmbNestedTreeCopy",
                      L"levels=4;folders=7;files=310",
                      elapsedUs,
                      kSmbTreeFixtureDirectories.size(),
                      sourceFiles.value(),
                      completed->second.hr);
    if (completed->second.hr != S_OK)
    {
        return reject(std::format(L"the Copy of the nested tree to the share failed: hr=0x{:08X}{}",
                                  static_cast<unsigned long>(completed->second.hr),
                                  DescribeTaskDiagnosticsForSelfTest(diagnostics)));
    }

    // The share holds exactly the source tree: every folder and file with the same bytes, and nothing else (no
    // stage such as .rs_copy_tmp_ left behind).
    if (const std::wstring difference = DescribeTreeDifferenceForSelfTest(source, copied); ! difference.empty())
    {
        return reject(difference);
    }
    for (size_t index = 0u; index < kSmbTreeFixtureDirectories.size(); ++index)
    {
        if (! LastWriteTimeNearForSelfTest(SmbTreeFixturePath(copied, index), SmbTreeFixtureTime(index)))
        {
            return reject(std::format(L"the copied folder '{}' has last-write time {} instead of its source's {}: its basic information was not restored",
                                      kSmbTreeFixtureDirectories[index].relative,
                                      DirectoryLastWriteTimeForSelfTest(SmbTreeFixturePath(copied, index)),
                                      SmbTreeFixtureTime(index)));
        }
    }
    AppendLog(std::format(L"SMB nested tree Copy: {} folders and {} files copied to the share in {} ms; every folder's last-write time restored",
                          kSmbTreeFixtureDirectories.size(),
                          sourceFiles.value(),
                          elapsedUs / 1000u));
    return true;
}

// I35 NAS nested Copy, host bridge routes (S6 saw a Managed Move to the NAS keep a stage even for a flat folder): a bridge
// Copy (the destination provider named explicitly, the same Local instance, takes the host bridge) and a Managed Move (a
// Move to another volume) of the nested tree and of a flat folder of 40 files, to a unique child of the registered SMB
// self-test root. The bridge kept each folder it created open, with DELETE access, until the folder's children finished;
// a server that refuses every rename inside a folder held that way (the Samba NAS probed on 2026-10-12) failed each
// child's stage publication (bridge.publication.unknown, the visible stage retained). Every item must end S_OK through
// the bridge, the share must hold exactly the source trees with no stage, no publication outcome may be unknown, and a
// Move must leave no source. Folder times come back, or the task carries the one note that the destination keeps no
// timestamps (a destination without FileIdInfo cannot bind a folder again). Environment-gated as
// R0fSmb_NestedTreeCopyToShare.
[[nodiscard]] bool RunFileOpsSmbBridgeTreeTransferChecks(SelfTestState& state, std::wstring& unavailableReason, bool move) noexcept
{
    const std::wstring_view name = move ? std::wstring_view(L"SMB Managed Move") : std::wstring_view(L"SMB bridge Copy");
    const auto reject            = [name](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"{}: {}", name, reason));
        return false;
    };
    const std::optional<std::wstring> smbRoot = ResolveSmbSelfTestRootForSelfTest(unavailableReason);
    if (! smbRoot.has_value())
    {
        return false;
    }
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
    {
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        return reject(L"no-activation guard unavailable");
    }

    const std::filesystem::path localBase = state.tempRoot / (move ? L"smb-bridge-move" : L"smb-bridge-copy");
    const std::filesystem::path sources   = localBase / L"src";
    const std::filesystem::path expected  = localBase / L"expected";
    const std::filesystem::path shareBase =
        std::filesystem::path(smbRoot.value()) / std::format(L"{}-{}-{}", move ? L"bridge-move" : L"bridge-copy", GetCurrentProcessId(), GetTickCount64());
    const std::filesystem::path destination = shareBase / L"dst";
    const bool verify                       = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
    std::optional<uint64_t> taskId;
    const auto restore = wil::scope_exit([&]() noexcept
    {
        const bool quiet = CancelOutstandingSelfTestOperations(operations);
        if (quiet && taskId.has_value())
        {
            operations->DismissCompletedTask(taskId.value());
            operations->DebugRemoveDiagnosticsForTask(taskId.value());
            state.completedTasks.erase(taskId.value());
        }
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verify;
        // Only this run's unique child of the share root is removed.
        static_cast<void>(SelfTest::RemoveAll(shareBase));
        static_cast<void>(SelfTest::RemoveAll(localBase));
    });

    // The sources, and the same trees again from the same generator: what the share must hold once a Move removed them.
    if (! RecreateEmptyDirectory(localBase))
    {
        return reject(L"local case root unavailable");
    }
    size_t treeFiles = 0u;
    for (const std::filesystem::path& root : {sources, expected})
    {
        const std::optional<size_t> files = WriteSmbTreeFixtureForSelfTest(root / L"tree");
        if (! files.has_value() || ! WriteSmbFlatFixtureForSelfTest(root / L"flat"))
        {
            return reject(L"local fixture unavailable");
        }
        treeFiles = files.value();
    }
    if (! RecreateEmptyDirectory(destination))
    {
        return reject(std::format(L"share destination '{}' unavailable", destination.native()));
    }

    // Verify Off: bridge Copy and managed Move both use the parallel walker; Move retains exact source cleanup authority.
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    std::vector<std::filesystem::path> items{sources / L"tree", sources / L"flat"};
    const auto startedAt = std::chrono::steady_clock::now();
    if (move)
    {
        taskId = StartFileOperationAndGetId(operations,
                                            FILESYSTEM_MOVE,
                                            FolderWindow::Pane::Left,
                                            FolderWindow::Pane::Right,
                                            state.fsLocal,
                                            std::move(items),
                                            destination,
                                            FILESYSTEM_FLAG_RECURSIVE,
                                            false);
    }
    else
    {
        taskId = StartFileOperationAndGetId(operations,
                                            FILESYSTEM_COPY,
                                            FolderWindow::Pane::Left,
                                            FolderWindow::Pane::Right,
                                            state.fsLocal,
                                            std::move(items),
                                            destination,
                                            FILESYSTEM_FLAG_RECURSIVE,
                                            false,
                                            0u,
                                            FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                            false,
                                            state.fsLocal);
    }
    if (! taskId.has_value())
    {
        return reject(L"the transfer to the share was not admitted");
    }
    std::optional<FolderWindow::FileOperationState::Task::ConflictPromptState> prompt;
    const auto finished = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        if (FolderWindow::FileOperationState::Task* task = operations->FindTask(taskId.value()); task != nullptr)
        {
            prompt = TryGetConflictPromptCopy(task);
        }
        return prompt.has_value() || (state.completedTasks.contains(taskId.value()) && operations->FindTask(taskId.value()) == nullptr);
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(300)), .operationName = name});
    const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
    const auto diagnostics   = CollectTaskDiagnosticsForSelfTest(operations, taskId.value());
    if (prompt.has_value())
    {
        return reject(std::format(L"the transfer stopped on a decision: status=0x{:08X} bucket={} source='{}' destination='{}'{}",
                                  static_cast<unsigned long>(prompt->status),
                                  static_cast<unsigned int>(prompt->bucket),
                                  prompt->sourcePath,
                                  prompt->destinationPath,
                                  DescribeTaskDiagnosticsForSelfTest(diagnostics)));
    }
    const auto completed = state.completedTasks.find(taskId.value());
    if (! finished.conditionMet || completed == state.completedTasks.end())
    {
        return reject(std::format(L"the transfer did not finish{}", DescribeTaskDiagnosticsForSelfTest(diagnostics)));
    }
    const CompletedTaskInfo& info = completed->second;
    Debug::Perf::Emit(L"FileOps.SelfTest.SmbBridgeTreeTransfer",
                      move ? L"operation=move;folders=8;files=350" : L"operation=copy;folders=8;files=350",
                      elapsedUs,
                      info.bridgeSourceDirectoryEnumerationCount,
                      treeFiles + kSmbFlatFixtureFiles,
                      info.hr);
    if (info.hr != S_OK)
    {
        return reject(std::format(L"the transfer failed: hr=0x{:08X}{}", static_cast<unsigned long>(info.hr), DescribeTaskDiagnosticsForSelfTest(diagnostics)));
    }
    // The bridge walker lists every source folder itself, including the parallel managed Move route.
    if (info.bridgeSourceDirectoryEnumerationCount == 0u)
    {
        return reject(L"the transfer did not go through the host bridge (it listed no source folder)");
    }
    if (HasDiagnosticCategoryForSelfTest(diagnostics, L"bridge.publication.unknown"))
    {
        return reject(std::format(L"a publication outcome was unknown{}", DescribeTaskDiagnosticsForSelfTest(diagnostics)));
    }
    for (const wchar_t* leaf : {L"tree", L"flat"})
    {
        if (const std::wstring difference = DescribeTreeDifferenceForSelfTest(expected / leaf, destination / leaf); ! difference.empty())
        {
            return reject(difference);
        }
        std::error_code existsError;
        if (move && (std::filesystem::exists(sources / leaf, existsError) || existsError))
        {
            return reject(std::format(L"the Move left its source '{}' behind{}", leaf, DescribeTaskDiagnosticsForSelfTest(diagnostics)));
        }
    }
    // Folder times come back, or the one note that the destination keeps no timestamps names their loss.
    const size_t folders = kSmbTreeFixtureDirectories.size() + 1u;
    size_t untimed       = LastWriteTimeNearForSelfTest(destination / L"flat", kSmbFlatFixtureTime) ? 0u : 1u;
    for (size_t index = 0u; index < kSmbTreeFixtureDirectories.size(); ++index)
    {
        untimed += LastWriteTimeNearForSelfTest(SmbTreeFixturePath(destination / L"tree", index), SmbTreeFixtureTime(index)) ? 0u : 1u;
    }
    const bool noted = HasDiagnosticCategoryForSelfTest(diagnostics, L"metadata.basic.unsupported");
    if (untimed != 0u && ! noted)
    {
        return reject(std::format(L"{} of {} folders lost their source's last-write time without the note that the destination keeps no timestamps{}",
                                  untimed,
                                  folders,
                                  DescribeTaskDiagnosticsForSelfTest(diagnostics)));
    }
    AppendLog(std::format(L"{}: {} folders and {} files through the bridge in {} ms ({} source folders listed by the bridge); {} of {} folder times "
                          L"restored{}",
                          name,
                          folders,
                          treeFiles + kSmbFlatFixtureFiles,
                          elapsedUs / 1000u,
                          info.bridgeSourceDirectoryEnumerationCount,
                          folders - untimed,
                          folders,
                          noted ? L", the destination-keeps-no-timestamps note recorded" : L""));
    return true;
}

[[nodiscard]] bool RunFileOpsLargeSelectionCountingChecks(SelfTestState& state, ClosureCaseScope& scope, const std::filesystem::path& root) noexcept
{
    constexpr unsigned long kEmptyFiles = 4096u;
    constexpr size_t kPayloadBytes      = 8u * 1024u * 1024u;
    const auto source                   = root / L"wide-selection";
    const auto destination              = root / L"wide-selection-out";
    if (! RecreateEmptyDirectory(source) || ! RecreateEmptyDirectory(destination) || ! WriteFilledTestFile(source / L"000-payload.bin", kPayloadBytes, 0x65u))
        return false;
    std::vector<std::filesystem::path> paths{source / L"000-payload.bin"};
    paths.reserve(kEmptyFiles + 1u);
    for (unsigned long index = 0u; index < kEmptyFiles; ++index)
    {
        const auto file = source / std::format(L"empty-{:04}.bin", index);
        if (! WriteFilledTestFile(file, 0u, 0u))
            return false;
        paths.push_back(file);
    }
    bool counted                          = false;
    bool correctLiveTotals                = false;
    bool duplicateDidNotGrow              = false;
    bool changedMetadataWithdrewCertainty = false;
    const auto observe                    = [&](FolderWindow::FileOperationState::Task& task) noexcept
    {
        if (counted || ! task._discoveryClosed.load(std::memory_order_acquire))
            return;
        counted           = true;
        correctLiveTotals = task._discoveredFileCount.load(std::memory_order_acquire) == kEmptyFiles + 1u &&
                            task._discoveredTotalBytes.load(std::memory_order_acquire) == kPayloadBytes &&
                            task._discoveredBytesKnown.load(std::memory_order_acquire);
        // Protocol controls on the real counted task: a retry repeats the same leaf snapshot;
        // a later authoritative source size differs. Neither may double count or reopen discovery,
        // and the latter must withdraw certainty instead of publishing a false fixed byte total.
        FolderWindow::FileOperationState::Task::PerItemCallbackCookie cookie{0u};
        FileSystemDiscoveryProgress progress{};
        progress.sizeBytes       = sizeof(progress);
        progress.discoveredFiles = 1u;
        progress.discoveredBytes = kPayloadBytes;
        progress.traversalClosed = TRUE;
        const HRESULT duplicate  = task.FileSystemReportDiscoveryProgress(&progress, &cookie);
        duplicateDidNotGrow      = SUCCEEDED(duplicate) && task._discoveredTotalBytes.load(std::memory_order_acquire) == kPayloadBytes &&
                                   task._discoveredFileCount.load(std::memory_order_acquire) == kEmptyFiles + 1u &&
                                   task._discoveredBytesKnown.load(std::memory_order_acquire);
        ++progress.discoveredBytes;
        const HRESULT changed            = task.FileSystemReportDiscoveryProgress(&progress, &cookie);
        changedMetadataWithdrewCertainty = SUCCEEDED(changed) && ! task._discoveredBytesKnown.load(std::memory_order_acquire) &&
                                           task._discoveryClosed.load(std::memory_order_acquire) &&
                                           task._discoveredTotalBytes.load(std::memory_order_acquire) == kPayloadBytes;
    };
    const auto info = scope.Run(FILESYSTEM_COPY, state.fsLocal, std::move(paths), destination, 512u * 1024u, observe, [&]() noexcept { return counted; });
    if (! info.has_value())
        return false;
    std::vector<std::wstring> sourceEntries;
    std::vector<std::wstring> destinationEntries;
    const bool sourceIntact = CollectRelativeTreeForSelfTest(source, sourceEntries) && sourceEntries.size() == kEmptyFiles + 1u &&
                              FileSizeEquals(source / L"000-payload.bin", kPayloadBytes);
    const bool destinationClean =
        CollectRelativeTreeForSelfTest(destination, destinationEntries) &&
        std::ranges::none_of(destinationEntries, [](const std::wstring& entry) noexcept { return ContainsIgnoreCase(entry, L".rs_"); });
    const bool passed        = counted && correctLiveTotals && duplicateDidNotGrow && changedMetadataWithdrewCertainty &&
                               info->hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && info->discoveryDurationUs <= 2'000'000u &&
                               info->discoveryBytesCompletedWhileOpen == 0u && info->discoveryGrowthAfterCloseCount == 0u && sourceIntact && destinationClean;
    const std::wstring facts = std::format(
        L"selected={};discoveryUs={};bytesWhileOpen={};liveTotals={};duplicateStable={};changedUncertain={};sourceIntact={};destinationClean={};hr=0x{:08X}",
        kEmptyFiles + 1u,
        info->discoveryDurationUs,
        info->discoveryBytesCompletedWhileOpen,
        correctLiveTotals,
        duplicateDidNotGrow,
        changedMetadataWithdrewCertainty,
        sourceIntact,
        destinationClean,
        static_cast<unsigned long>(info->hr));
    AppendLog(std::format(L"Large selection counting: {} {}", passed ? L"passed" : L"FAILED", facts));
    Debug::Perf::Emit(L"FileOps.SelfTest.SelectedFilesConvergence", facts, info->discoveryDurationUs, kEmptyFiles + 1u, 2'000'000u, passed ? S_OK : E_FAIL);
    return passed;
}

[[nodiscard]] bool RunFileOpsParallelMoveSafetyChecks(SelfTestState& state,
                                                      FolderWindow::FileOperationState* operations,
                                                      ClosureCaseScope& scope,
                                                      const std::filesystem::path& localRoot,
                                                      const std::filesystem::path& destinationRoot) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;
    for (const std::wstring_view mode : {L"new-child", L"cancel", L"retry", L"skip", L"unknown"})
    {
        const bool newChild    = mode == L"new-child";
        const bool cancel      = mode == L"cancel";
        const bool retry       = mode == L"retry";
        const bool skip        = mode == L"skip";
        const bool unknown     = mode == L"unknown";
        const auto source      = localRoot / mode / L"tree";
        const auto expected    = localRoot / mode / L"expected";
        const auto destination = destinationRoot / mode;
        if (! RecreateEmptyDirectory(source) || ! RecreateEmptyDirectory(expected) || ! RecreateEmptyDirectory(destination))
            return false;
        for (unsigned int index = 0u; index < 4u; ++index)
        {
            const auto name  = std::format(L"file-{}.bin", index);
            const auto value = static_cast<unsigned char>(0x61u + index);
            if (! WriteFilledTestFile(source / name, 256u * 1024u + index, value) || ! WriteFilledTestFile(expected / name, 256u * 1024u + index, value))
                return false;
        }
        const auto resetHooks = wil::scope_exit([]() noexcept
        {
            ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            SetFileOpsManagedCleanupKnownNonCommitForSelfTest(HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), 0u);
            SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(0u);
        });
        static_cast<void>(TakeFileOpsManagedCleanupKnownNonCommitAttemptsForSelfTest());
        static_cast<void>(TakeFileOpsManagedCleanupUnknownOutcomeAttemptsForSelfTest());
        SetFileOpsManagedCleanupKnownNonCommitForSelfTest(HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), retry || skip ? 1u : 0u);
        SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(unknown ? 1u : 0u);
        // All four publications must finish before injecting cleanup outcomes. A single entered
        // pause plus four I/O permits also admits files still publishing, whose canceled SMB
        // mutation/abort can legitimately need recovery. Hold the exact cleanup checkpoint instead.
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(true, 60'000ull);
        bool acted                  = false;
        bool injectionSucceeded     = false;
        unsigned int answered       = 0u;
        uint64_t answeredGeneration = 0u;
        std::wstring keptFile;
        const auto observe = [&](Task& task) noexcept
        {
            if (! acted && GetFileOpsBridgeMoveSourceCleanupPausedCountForSelfTest() == 4u && task._discoveryClosed.load(std::memory_order_acquire) &&
                task._bridgeActiveFileTransfers.load(std::memory_order_acquire) == 4u)
            {
                acted = true;
                if (cancel)
                    task.RequestCancel();
                else if (newChild)
                    injectionSucceeded = WriteFilledTestFile(source / L"unselected.bin", 512u, 0x99u) &&
                                         WriteFilledTestFile(localRoot / mode / L"unselected-expected.bin", 512u, 0x99u);
                ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
                SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            }
        };
        const auto answer = [&](Task& task) noexcept -> bool
        {
            const auto prompt = TryGetConflictPromptCopy(&task);
            if (! prompt.has_value() || prompt->generation == answeredGeneration)
                return true;
            const auto action = retry ? Task::ConflictAction::Retry : Task::ConflictAction::Skip;
            if ((! retry && ! skip) || prompt->bucket != Task::ConflictBucket::SharingViolation || ! PromptHasAction(prompt.value(), action))
                return false;
            answeredGeneration = prompt->generation;
            keptFile           = prompt->sourcePath;
            ++answered;
            task.DebugSubmitCurrentConflictDecision(action, false);
            return true;
        };
        const auto info = scope.Run(FILESYSTEM_MOVE, state.fsLocal, {source}, destination, 0u, observe, {}, {}, answer);
        if (! info.has_value())
        {
            AppendLog(std::format(L"Parallel Move safety {}: unexpected decision or timeout{}",
                                  mode,
                                  DescribeTaskDiagnosticsForSelfTest(CollectTaskDiagnosticsForSelfTest(operations, scope.LastTaskId()))));
            return false;
        }
        const auto copied = destination / L"tree";
        std::vector<std::wstring> destinationEntries;
        const bool destinationClean =
            CollectRelativeTreeForSelfTest(copied, destinationEntries) &&
            std::ranges::none_of(destinationEntries, [](const std::wstring& entry) noexcept { return ContainsIgnoreCase(entry, L".rs_"); });
        const bool sourceIntact        = ! cancel || DescribeTreeDifferenceForSelfTest(expected, source).empty();
        const bool destinationComplete = DescribeTreeDifferenceForSelfTest(expected, copied).empty();
        bool passed =
            acted && info->bridgeMaxActiveFileTransfers == 4u && info->discoveryGrowthAfterCloseCount == 0u && destinationClean && destinationComplete;
        std::error_code sourceError;
        if (cancel)
        {
            passed = passed && info->hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && sourceIntact;
            for (const auto& entry : destinationEntries)
                passed = passed && FilesEqualBytes(expected / entry, copied / entry);
        }
        else if (unknown)
        {
            passed = passed && info->hr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && answered == 0u && std::filesystem::exists(source, sourceError) &&
                     ! sourceError && ! info->sourceItemResults.empty() && info->sourceItemResults.front().has_value() &&
                     info->sourceItemResults.front().value().sourceDisposition == FileOperations::SourceDisposition::Unknown;
        }
        else
        {
            passed = passed && DescribeTreeDifferenceForSelfTest(expected, copied).empty();
            if (retry)
                passed = passed && info->hr == S_OK && answered == 1u && ! std::filesystem::exists(source, sourceError) && ! sourceError;
            else
            {
                passed = passed && info->hr == S_FALSE && ! info->sourceItemResults.empty() && info->sourceItemResults.front().has_value() &&
                         info->sourceItemResults.front().value().sourceDisposition == FileOperations::SourceDisposition::Retained;
                if (newChild)
                    passed = passed && acted && injectionSucceeded &&
                             FilesEqualBytes(source / L"unselected.bin", localRoot / mode / L"unselected-expected.bin") &&
                             ! std::filesystem::exists(copied / L"unselected.bin", sourceError) && ! sourceError;
                else
                    passed = passed && answered == 1u && ! keptFile.empty() &&
                             FilesEqualBytes(std::filesystem::path(keptFile), expected / std::filesystem::path(keptFile).filename());
            }
        }
        const std::wstring facts =
            std::format(L"mode={};activeMax={};acted={};answers={};clean={};destinationComplete={};sourceIntact={};growth={};hr=0x{:08X}",
                        mode,
                        info->bridgeMaxActiveFileTransfers,
                        acted,
                        answered,
                        destinationClean,
                        destinationComplete,
                        sourceIntact,
                        info->discoveryGrowthAfterCloseCount,
                        static_cast<unsigned long>(info->hr));
        AppendLog(std::format(L"Parallel Move safety: {} {}", passed ? L"passed" : L"FAILED", facts));
        Debug::Perf::Emit(L"FileOps.SelfTest.ParallelMoveSafety", facts, 0u, info->bridgeMaxActiveFileTransfers, 4u, passed ? S_OK : E_FAIL);
        if (! passed)
            return false;
    }
    return true;
}

// Bounded recursive SMB witness: empty directories, three levels, zero-byte content and exact
// post-order source cleanup. The separate legacy 350-file/two-root cost scenario keeps its own gate.
[[nodiscard]] bool RunFileOpsNestedParallelMoveChecks(SelfTestState& state,
                                                      FolderWindow::FileOperationState* operations,
                                                      ClosureCaseScope& scope,
                                                      const std::filesystem::path& localRoot,
                                                      const std::filesystem::path& destination) noexcept
{
    const auto source   = localRoot / L"nested";
    const auto expected = localRoot / L"expected";
    const std::array<std::wstring_view, 6u> directories{{L"", L"empty", L"alpha", L"alpha/empty", L"beta", L"beta/nested"}};
    uint64_t totalBytes = 0u;
    for (const auto& root : {source, expected})
    {
        for (const auto directory : directories)
            if (! RecreateEmptyDirectory(root / directory))
                return false;
        for (unsigned int index = 0u; index < 12u; ++index)
        {
            const auto directory = index < 4u ? L"" : index < 8u ? L"alpha" : L"beta/nested";
            const size_t bytes   = (index + 1u) * 1024u;
            if (! WriteFilledTestFile(root / directory / std::format(L"file-{:02}.bin", index), bytes, static_cast<unsigned char>(0x42u + index)))
                return false;
            if (root == source)
                totalBytes += bytes;
        }
        if (! WriteFilledTestFile(root / L"beta/zero.bin", 0u, 0u))
            return false;
        for (size_t index = 0u; index < directories.size(); ++index)
            if (! SetDirectoryLastWriteTimeForSelfTest(root / directories[index], ManyDirectoriesFolderTimeForSelfTest(index)))
                return false;
    }
    if (! RecreateEmptyDirectory(destination))
        return false;
    const auto start = std::chrono::steady_clock::now();
    const auto info  = scope.Run(FILESYSTEM_MOVE, state.fsLocal, {source}, destination, 0u);
    if (! info.has_value())
        return false;
    const auto copied = destination / L"nested";
    std::error_code sourceError;
    bool timestamps = true;
    for (size_t index = 0u; index < directories.size(); ++index)
        timestamps = LastWriteTimeNearForSelfTest(copied / directories[index], ManyDirectoriesFolderTimeForSelfTest(index)) && timestamps;
    const auto contentDefect = DescribeTreeDifferenceForSelfTest(expected, copied);
    // Match the existing SMB metadata contract: an identity-less NAS may refuse an exact
    // directory rebind. Its explicit unsupported-metadata note must disclose unpreserved times.
    const bool unsupportedMetadataNoted =
        HasDiagnosticCategoryForSelfTest(CollectTaskDiagnosticsForSelfTest(operations, scope.LastTaskId()), L"metadata.basic.unsupported");
    const bool passed  = info->hr == S_OK && info->bridgeMaxActiveFileTransfers == 4u && info->discoveryClosed && info->discoveredFiles == 13u &&
                         info->discoveredDirectories == directories.size() && info->discoveredTotalBytes == totalBytes &&
                         info->bridgeSourceDirectoryEnumerationCount == directories.size() && info->discoveryGrowthAfterCloseCount == 0u &&
                         contentDefect.empty() && (timestamps || unsupportedMetadataNoted) && ! std::filesystem::exists(source, sourceError) && ! sourceError;
    const auto elapsed = Debug::Perf::ElapsedUs(start);
    const auto facts =
        std::format(L"files={};directories={};enumerations={};activeMax={};bytes={};timestamps={};unsupportedMetadataNoted={};content={};hr=0x{:08X}",
                    info->discoveredFiles,
                    info->discoveredDirectories,
                    info->bridgeSourceDirectoryEnumerationCount,
                    info->bridgeMaxActiveFileTransfers,
                    info->discoveredTotalBytes,
                    timestamps,
                    unsupportedMetadataNoted,
                    contentDefect,
                    static_cast<unsigned long>(info->hr));
    AppendLog(std::format(L"Nested parallel Move: {} {}", passed ? L"passed" : L"FAILED", facts));
    Debug::Perf::Emit(L"FileOps.SelfTest.NestedParallelMove", facts, elapsed, info->bridgeMaxActiveFileTransfers, 4u, passed ? S_OK : E_FAIL);
    return passed;
}

// I42: run identical fresh fixtures as a directory and as the equivalent Ctrl+A leaf list. Observe
// actual admitted native transfer phases, not selected-root worker counts, and compare every destination byte.
[[nodiscard]] bool RunFileOpsSelectionShapeChecks(SelfTestState& state, bool smb, std::wstring& unavailableReason) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Selection shape: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner/provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const auto localRoot                  = state.tempRoot / (smb ? L"selection-shape-smb" : L"selection-shape-local");
    std::filesystem::path destinationRoot = localRoot / L"out";
    if (smb)
    {
        const auto share = ResolveSmbSelfTestRootForSelfTest(unavailableReason);
        if (! share.has_value())
            return false;
        destinationRoot = std::filesystem::path(share.value()) / std::format(L"selection-shape-{}-{}", GetCurrentProcessId(), GetTickCount64());
    }
    std::string savedConfiguration;
    if (! BackupPluginConfiguration(state.infoLocal.get(), savedConfiguration) || ! SetPluginConfiguration(state.infoLocal.get(), savedConfiguration) ||
        ! BackupPluginConfiguration(state.infoLocal.get(), savedConfiguration))
        return reject(L"configuration backup unavailable");
    const auto restore = wil::scope_exit([&]() noexcept
    {
        if (CancelOutstandingSelfTestOperations(operations))
        {
            SetFileOpsAutoConcurrencyOverrideForSelfTest(false, 0u, FILESYSTEM_STORAGE_UNKNOWN);
            static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), savedConfiguration));
            static_cast<void>(SelfTest::RemoveAll(destinationRoot));
            static_cast<void>(SelfTest::RemoveAll(localRoot));
        }
    });
    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    constexpr unsigned long kFiles                            = 8u;
    // Multiple native/bridge pump chunks are required: a one-chunk file has no partial-byte
    // observation, so counting only partial transfers would falsely report zero active workers.
    constexpr size_t kBytesPerFile = 4u * 1024u * 1024u;
    constexpr uint64_t kTotalBytes = uint64_t{kFiles} * kBytesPerFile;
    std::vector<std::wstring> defects;
    constexpr unsigned int routeCount = 3u;
    for (unsigned int route = 0u; route < routeCount; ++route)
    {
        const bool move                        = smb && route == 2u;
        const bool bridge                      = smb && route == 1u;
        const bool autoWithInactiveOne         = ! smb && route == 1u;
        const unsigned int expectedConcurrency = ! smb && route == 2u ? 1u : 4u;
        const std::string config               = std::format(R"json({{"concurrencyMode":"{}","copyMoveMaxConcurrency":{}}})json",
                                                             autoWithInactiveOne ? "auto" : "manual",
                                                             autoWithInactiveOne ? 1u : expectedConcurrency);
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
            return reject(L"test concurrency configuration unavailable");
        SetFileOpsAutoConcurrencyOverrideForSelfTest(autoWithInactiveOne, 4u, FILESYSTEM_STORAGE_SSD);
        for (const bool selectedLeaves : {false, true})
        {
            const std::wstring label =
                std::format(L"{}-{}-route{}-{}", smb ? L"smb" : L"local", move ? L"move" : L"copy", route, selectedLeaves ? L"leaves" : L"folder");
            const auto source      = localRoot / label / L"album";
            const auto expected    = localRoot / label / L"expected";
            const auto destination = destinationRoot / label;
            if (! RecreateEmptyDirectory(source) || ! RecreateEmptyDirectory(expected) || ! RecreateEmptyDirectory(destination))
                return reject(std::format(L"{}: fixture directory unavailable", label));
            std::vector<std::filesystem::path> leaves;
            for (unsigned long index = 0u; index <= kFiles; ++index)
            {
                const auto leaf    = std::format(L"file-{:02}.bin", index);
                const size_t bytes = index == kFiles ? 0u : kBytesPerFile;
                const auto value   = static_cast<unsigned char>(0x31u + index);
                if (! WriteFilledTestFile(source / leaf, bytes, value) || ! WriteFilledTestFile(expected / leaf, bytes, value))
                    return reject(std::format(L"{}: source bytes unavailable", label));
                leaves.push_back(source / leaf);
            }
            size_t maximumActiveFiles      = 0u;
            size_t maximumPartialByteFiles = 0u;
            bool liveStatusConverged       = false;
            const auto observe             = [&](FolderWindow::FileOperationState::Task& task) noexcept
            {
                if (selectedLeaves && ! liveStatusConverged && task._discoveryClosed.load(std::memory_order_acquire))
                {
                    operations->DebugEnsurePopupVisibleForSelfTest();
                    FileOperationsPopupInternal::TaskSnapshot snapshot{};
                    FileOperationsPopupInternal::TaskProgressPresentationDebugSnapshot presentation{};
                    if (DebugGetFileOperationsPopupTaskSnapshot(operations->GetPopupHwndForSelfTest(), task.GetId(), snapshot))
                    {
                        DebugResolveFileOperationsWholeTaskProgress(snapshot, presentation);
                        liveStatusConverged = snapshot.discoveryClosed && snapshot.discoveredFileCount == kFiles + 1u && snapshot.totalBytes == kTotalBytes &&
                                              snapshot.discoveredBytesKnown && presentation.determinate && ! presentation.showDiscoveryActivity &&
                                              ! presentation.provisional;
                    }
                }
                if (move || bridge)
                {
                    maximumActiveFiles = (std::max)(maximumActiveFiles, static_cast<size_t>(task._bridgeActiveFileTransfers.load(std::memory_order_acquire)));
                    return;
                }
                std::scoped_lock lock(task._inFlightFilesMutex);
                size_t active               = 0u;
                size_t partialByteTransfers = 0u;
                for (size_t index = 0u; index < task._inFlightFileCount; ++index)
                {
                    const auto& file = task._inFlightFiles[index];
                    // Commit/publication still owns the admitted file's I/O permits. Buffered
                    // small files may report their only byte chunk already complete, so retain
                    // these real work phases; an Opening row waiting for a permit does not count.
                    const bool partialBytePhase = (file.completedBytes > 0u && file.completedBytes < file.totalBytes) ||
                                                  file.phase == FileSystemFileWorkPhase::Finalizing || file.phase == FileSystemFileWorkPhase::Publishing;
                    // This native phase follows exact handles, metadata preparation,
                    // buffer allocation and admission, before the first read. It is
                    // real transfer work at zero/full byte callback boundaries; a
                    // Begin-only row waiting for admission still does not count.
                    const bool activePhase = partialBytePhase || (file.explicitLifecycle && file.phase == FileSystemFileWorkPhase::Transferring);
                    if (! file.terminal && file.totalBytes > 0u && activePhase)
                        ++active;
                    if (! file.terminal && file.totalBytes > 0u && partialBytePhase)
                        ++partialByteTransfers;
                }
                maximumActiveFiles      = (std::max)(maximumActiveFiles, active);
                maximumPartialByteFiles = (std::max)(maximumPartialByteFiles, partialByteTransfers);
            };
            const auto started = std::chrono::steady_clock::now();
            const auto info    = scope.Run(move ? FILESYSTEM_MOVE : FILESYSTEM_COPY,
                                           state.fsLocal,
                                           selectedLeaves ? std::move(leaves) : std::vector<std::filesystem::path>{source},
                                           destination,
                                           8u * 1024u * 1024u,
                                           observe,
                                           {},
                                           {},
                                           {},
                                           FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                           FILESYSTEM_FLAG_NONE,
                                           1024u,
                                           bridge);
            if (! info.has_value())
                return reject(std::format(L"{}: transfer did not complete or requested an unexpected decision{}",
                                          label,
                                          DescribeTaskDiagnosticsForSelfTest(CollectTaskDiagnosticsForSelfTest(operations, scope.LastTaskId()))));
            maximumActiveFiles               = (std::max)(maximumActiveFiles, static_cast<size_t>(info->bridgeMaxActiveFileTransfers));
            const uint64_t elapsedUs         = Debug::Perf::ElapsedUs(started);
            const auto copied                = selectedLeaves ? destination : destination / L"album";
            const std::wstring contentDefect = DescribeTreeDifferenceForSelfTest(expected, copied);
            std::error_code sourceError;
            const bool sourceExists = std::filesystem::exists(source, sourceError);
            bool sourcesCorrect     = ! sourceError;
            if (move)
            {
                sourcesCorrect = sourcesCorrect && (selectedLeaves ? std::filesystem::is_empty(source, sourceError) && ! sourceError : ! sourceExists);
            }
            else
            {
                sourcesCorrect = sourcesCorrect && DescribeTreeDifferenceForSelfTest(expected, source).empty();
            }
            const bool correct       = info->hr == S_OK && info->discoveryClosed && info->discoveredFiles == kFiles + 1u &&
                                       info->discoveredTotalBytes == kTotalBytes && info->discoveryGrowthAfterCloseCount == 0u &&
                                       info->discoveryBytesCompletedWhileOpen < kTotalBytes / 8u && maximumActiveFiles == expectedConcurrency &&
                                       contentDefect.empty() && sourcesCorrect && (! selectedLeaves || liveStatusConverged);
            const std::wstring facts = std::format(L"scenario={};configured={};activeMax={};discoveryUs={};bytesWhileOpen={};files={};bytes={};growth={};"
                                                   L"elapsedUs={};contents={};sources={};liveStatus={};hr=0x{:08X}",
                                                   label,
                                                   info->configuredMaxConcurrency,
                                                   maximumActiveFiles,
                                                   info->discoveryDurationUs,
                                                   info->discoveryBytesCompletedWhileOpen,
                                                   info->discoveredFiles,
                                                   info->discoveredTotalBytes,
                                                   info->discoveryGrowthAfterCloseCount,
                                                   elapsedUs,
                                                   contentDefect.empty(),
                                                   sourcesCorrect,
                                                   liveStatusConverged,
                                                   static_cast<unsigned long>(info->hr));
            AppendLog(std::format(L"Selection shape: {} {}", correct ? L"passed" : L"FAILED", facts));
            Debug::Perf::Emit(L"FileOps.SelfTest.SelectionShape", facts, elapsedUs, maximumActiveFiles, expectedConcurrency, correct ? S_OK : E_FAIL);
            AppendLog(std::format(L"Selection shape phase witness: scenario={};actualActiveMax={};partialByteActiveMax={};expected={}",
                                  label,
                                  maximumActiveFiles,
                                  maximumPartialByteFiles,
                                  expectedConcurrency));
            Debug::Perf::Emit(L"FileOps.SelfTest.SelectionShape.PhaseConcurrency", label, elapsedUs, maximumActiveFiles, maximumPartialByteFiles, S_OK);
            if (! correct)
                defects.push_back(std::format(L"{} {}", facts, contentDefect));
        }
    }
    if (defects.empty())
    {
        if (smb)
        {
            if (! RunFileOpsParallelMoveSafetyChecks(state, operations, scope, localRoot / L"safety", destinationRoot / L"safety"))
                defects.push_back(L"parallel managed Move safety failed");
            if (! RunFileOpsNestedParallelMoveChecks(state, operations, scope, localRoot / L"nested", destinationRoot / L"nested"))
                defects.push_back(L"nested managed Move failed");
        }
        else if (! RunFileOpsLargeSelectionCountingChecks(state, scope, localRoot))
            defects.push_back(L"large fresh selected-file counting failed");
    }
    if (! scope.Close())
        return reject(L"worker/settings cleanup did not settle");
    SetFileOpsAutoConcurrencyOverrideForSelfTest(false, 0u, FILESYSTEM_STORAGE_UNKNOWN);
    std::string restored;
    if (! SetPluginConfiguration(state.infoLocal.get(), savedConfiguration) || ! BackupPluginConfiguration(state.infoLocal.get(), restored) ||
        restored != savedConfiguration)
        return reject(L"configuration was not restored");
    for (const auto& defect : defects)
        AppendLog(std::format(L"Selection shape defect: {}", defect));
    return defects.empty();
}

// The many-folders fixture: 12 branches of 12 of 12 leaves (1,885 folders with the root), one 64-byte file in each of the
// 144 middle folders, and folder i's last-write time 2001 + i seconds. `directories` lists the folders relative to the
// root, the root itself first as the empty entry.
[[nodiscard]] bool WriteManyDirectoriesFixtureForSelfTest(const std::filesystem::path& source,
                                                          std::vector<std::wstring>& directories,
                                                          size_t& sourceFiles) noexcept
{
    constexpr unsigned int kFanOut = 12u;
    constexpr size_t kFileBytes    = 64u;
    directories.clear();
    sourceFiles = 0u;
    directories.reserve(1u + kFanOut + kFanOut * kFanOut + kFanOut * kFanOut * kFanOut);
    directories.emplace_back();
    for (unsigned int branch = 0u; branch < kFanOut; ++branch)
    {
        directories.push_back(std::format(L"b{:02}", branch));
        for (unsigned int middle = 0u; middle < kFanOut; ++middle)
        {
            directories.push_back(std::format(L"b{:02}\\m{:02}", branch, middle));
            for (unsigned int leaf = 0u; leaf < kFanOut; ++leaf)
            {
                directories.push_back(std::format(L"b{:02}\\m{:02}\\l{:02}", branch, middle, leaf));
            }
        }
    }
    if (! RecreateEmptyDirectory(source))
    {
        return false;
    }
    for (const std::wstring& relative : directories)
    {
        std::error_code createError;
        static_cast<void>(std::filesystem::create_directory(ManyDirectoriesFolderForSelfTest(source, relative), createError));
        if (createError)
        {
            return false;
        }
        if (std::ranges::count(relative, L'\\') == 1)
        {
            if (! WriteFilledTestFile(
                    ManyDirectoriesFolderForSelfTest(source, relative) / L"f.bin", kFileBytes, static_cast<unsigned char>(0x41u + sourceFiles % 26u)))
            {
                return false;
            }
            ++sourceFiles;
        }
    }
    for (size_t index = 0u; index < directories.size(); ++index)
    {
        if (! SetDirectoryLastWriteTimeForSelfTest(ManyDirectoriesFolderForSelfTest(source, directories[index]), ManyDirectoriesFolderTimeForSelfTest(index)))
        {
            return false;
        }
    }
    return true;
}

// Empty when every folder of the many-folders fixture under `copied` has its source's last-write time; otherwise the first
// that does not.
[[nodiscard]] std::wstring FirstManyDirectoriesFolderWithoutTimeForSelfTest(const std::filesystem::path& copied,
                                                                            const std::vector<std::wstring>& directories) noexcept
{
    for (size_t index = 0u; index < directories.size(); ++index)
    {
        if (DirectoryLastWriteTimeForSelfTest(ManyDirectoriesFolderForSelfTest(copied, directories[index])) != ManyDirectoriesFolderTimeForSelfTest(index))
        {
            return directories[index].empty() ? std::wstring(L"(root)") : directories[index];
        }
    }
    return {};
}

// I35 NAS nested Copy, perf witness: a 1,885-folder Local tree (12 branches of 12 of 12 leaves, one 64-byte file
// in each of the 144 middle folders) is copied three times through the provider's own walker at configured
// concurrency 4. Every copy must hold exactly the source tree, and every folder must get its source's
// last-write time back. FileOps.SelfTest.LocalManyDirectoriesCopy records each copy's wall time.
[[nodiscard]] bool RunFileOpsLocalManyDirectoriesCopyChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Local many-folders Copy: {}", reason));
        return false;
    };
    if (! state.fsLocal || ! state.infoLocal || ! SetPluginConfiguration(state.infoLocal.get(), kLocalManualCopyConcurrency4ForSelfTest))
    {
        return reject(L"Local provider or its manual copy/move concurrency 4 unavailable");
    }
    const std::filesystem::path base        = state.tempRoot / L"many-folders-copy";
    const std::filesystem::path source      = base / L"src" / L"tree";
    const std::filesystem::path destination = base / L"dst";
    const auto removeBase                   = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(base)); });

    constexpr unsigned int kRepetitions = 3u;
    std::vector<std::wstring> directories; // relative to the tree root; the root itself is empty
    size_t sourceFiles = 0u;
    if (! WriteManyDirectoriesFixtureForSelfTest(source, directories, sourceFiles))
    {
        return reject(L"source fixture unavailable");
    }
    std::vector<std::wstring> sourceEntries;
    if (! CollectRelativeTreeForSelfTest(source, sourceEntries) || sourceEntries.size() != directories.size() - 1u + sourceFiles)
    {
        return reject(L"the source tree could not be listed");
    }

    FileSystemOptions options{};
    options.sizeBytes = sizeof(FileSystemOptions);
    uint64_t totalUs  = 0u;
    uint64_t bestUs   = (std::numeric_limits<uint64_t>::max)();
    for (unsigned int repetition = 0u; repetition < kRepetitions; ++repetition)
    {
        if (! RecreateEmptyDirectory(destination))
        {
            return reject(L"destination unavailable");
        }
        const std::filesystem::path copied = destination / source.filename();
        const auto startedAt               = std::chrono::steady_clock::now();
        const HRESULT copyHr               = state.fsLocal->CopyItem(source.c_str(), copied.c_str(), FILESYSTEM_FLAG_RECURSIVE, &options, nullptr, nullptr);
        const uint64_t elapsedUs           = Debug::Perf::ElapsedUs(startedAt);
        Debug::Perf::Emit(
            L"FileOps.SelfTest.LocalManyDirectoriesCopy", L"folders=1885;files=144;concurrency=4", elapsedUs, directories.size(), repetition, copyHr);
        if (copyHr != S_OK)
        {
            return reject(std::format(L"copy {} failed: hr=0x{:08X}", repetition, static_cast<unsigned long>(copyHr)));
        }
        std::vector<std::wstring> copiedEntries;
        if (! CollectRelativeTreeForSelfTest(copied, copiedEntries) ||
            ! std::ranges::equal(
                copiedEntries, sourceEntries, [](const std::wstring& left, const std::wstring& right) noexcept { return EqualsIgnoreCase(left, right); }))
        {
            return reject(
                std::format(L"copy {} differs from its source: {} entries copied, {} expected", repetition, copiedEntries.size(), sourceEntries.size()));
        }
        if (const std::wstring untimed = FirstManyDirectoriesFolderWithoutTimeForSelfTest(copied, directories); ! untimed.empty())
        {
            return reject(std::format(L"copy {}: the folder '{}' did not get its source's last-write time back", repetition, untimed));
        }
        totalUs += elapsedUs;
        bestUs = (std::min)(bestUs, elapsedUs);
    }
    AppendLog(std::format(L"Local many-folders Copy: {} folders and {} files copied {} times, mean {} ms, best {} ms",
                          directories.size(),
                          sourceFiles,
                          kRepetitions,
                          totalUs / kRepetitions / 1000u,
                          bestUs / 1000u));
    return true;
}

// True when `path` opens with share mode 0 and read access: nothing else has the file open. A scanner may hold a
// new file for a moment, so a sharing violation is retried briefly.
[[nodiscard]] bool FileOpensExclusivelyForSelfTest(const wchar_t* path) noexcept
{
    const std::wstring extended = MakeExtendedPathForSelfTest(path);
    for (unsigned int attempt = 0u; attempt < 25u; ++attempt)
    {
        const wil::unique_hfile file(CreateFileW(extended.c_str(), GENERIC_READ, 0u, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
        if (file)
        {
            return true;
        }
        if (GetLastError() != ERROR_SHARING_VIOLATION)
        {
            return false;
        }
        Sleep(20u);
    }
    return false;
}

// A Local CopyItem on its own thread whose callback holds one file progress report until released, which
// pauses the walker inside that file's item with every folder above the file published and still waiting
// for its children.
struct HeldLocalCopyCallback final : IFileSystemCallback
{
    wil::unique_event_nothrow held;
    wil::unique_event_nothrow release;
    // Hold the report CopyFileInternal makes after the file's copy route returned (the file is complete and the
    // walker has no handle on it left), not the first report.
    bool holdAfterPublication = false;
    std::atomic<bool> holding{false};
    // The top-level item's completion, read once the copy thread has ended.
    HRESULT itemStatus = E_PENDING;
    std::optional<FileSystemItemMutationResult> receipt;

    HRESULT STDMETHODCALLTYPE FileSystemProgress(FileSystemOperation /*operationType*/,
                                                 unsigned long /*totalItems*/,
                                                 unsigned long /*completedItems*/,
                                                 uint64_t /*totalBytes*/,
                                                 uint64_t /*completedBytes*/,
                                                 const wchar_t* /*currentSourcePath*/,
                                                 const wchar_t* currentDestinationPath,
                                                 uint64_t currentItemTotalBytes,
                                                 uint64_t currentItemCompletedBytes,
                                                 FileSystemOptions* /*options*/,
                                                 uint64_t /*progressStreamId*/,
                                                 void* /*cookie*/) noexcept override
    {
        if (currentItemTotalBytes == 0u || currentDestinationPath == nullptr || holding.load(std::memory_order_acquire))
        {
            return S_OK;
        }
        // A new file is written at its final name, so it exists while it is copied; only after the copy route
        // returned does the walker hold no handle on it, and only then does it open exclusively.
        if (holdAfterPublication && (currentItemCompletedBytes != currentItemTotalBytes || ! FileOpensExclusivelyForSelfTest(currentDestinationPath)))
        {
            return S_OK;
        }
        if (holding.exchange(true, std::memory_order_acq_rel))
        {
            return S_OK;
        }
        held.SetEvent();
        static_cast<void>(WaitForSingleObject(release.get(), INFINITE));
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation /*operationType*/,
                                                      unsigned long /*itemIndex*/,
                                                      const wchar_t* /*sourcePath*/,
                                                      const wchar_t* /*destinationPath*/,
                                                      HRESULT status,
                                                      const FileSystemItemMutationResult* mutationResult,
                                                      FileSystemOptions* /*options*/,
                                                      void* /*cookie*/) noexcept override
    {
        itemStatus = status;
        receipt    = FileSystemRouteContract::SnapshotItemMutationResult(mutationResult);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* cancel, void* /*cookie*/) noexcept override
    {
        if (! cancel)
        {
            return E_POINTER;
        }
        *cancel = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation /*operationType*/,
                                              const wchar_t* /*sourcePath*/,
                                              const wchar_t* /*destinationPath*/,
                                              HRESULT /*status*/,
                                              FileSystemIssueAction* action,
                                              IFileSystemBoundObject** expectedDestination,
                                              FileSystemOptions* /*options*/,
                                              void* /*cookie*/) noexcept override
    {
        if (! action || ! expectedDestination)
        {
            return E_POINTER;
        }
        *expectedDestination = nullptr;
        *action              = FileSystemIssueAction::Cancel;
        return S_OK;
    }
};

// The copy, its callback and its paths live on the heap: a copy that does not end in time keeps using them, so
// they are then left allocated rather than freed under it.
struct HeldLocalCopyJob final
{
    HeldLocalCopyCallback callback{};
    FileSystemOptions options{};
    std::wstring source;
    std::wstring destination;
    std::wstring destinationFolder;
    IFileSystem* fileSystem = nullptr;
    // CopyItems at concurrency 1 into destinationFolder runs the serial walker (CopyDirectoryInternal); CopyItem runs
    // the producer/worker walker, which a recursive folder always takes otherwise.
    bool serialWalker = false;
    HRESULT hr        = E_PENDING;
};

// Owns one held copy: starts it on its own thread and, however the case ends, releases the held report and waits
// for the copy before the job can go away.
class HeldLocalCopyRun final
{
public:
    HeldLocalCopyRun() noexcept                          = default;
    HeldLocalCopyRun(const HeldLocalCopyRun&)            = delete;
    HeldLocalCopyRun& operator=(const HeldLocalCopyRun&) = delete;
    HeldLocalCopyRun(HeldLocalCopyRun&&)                 = delete;
    HeldLocalCopyRun& operator=(HeldLocalCopyRun&&)      = delete;
    ~HeldLocalCopyRun()
    {
        static_cast<void>(Finish());
    }

    [[nodiscard]] HeldLocalCopyJob* Start(IFileSystem* fileSystem,
                                          const std::filesystem::path& source,
                                          const std::filesystem::path& destination,
                                          bool holdAfterPublication,
                                          bool serialWalker) noexcept
    {
        _job.reset(new (std::nothrow) HeldLocalCopyJob{});
        if (! _job || FAILED(_job->callback.held.create(wil::EventOptions::ManualReset)) ||
            FAILED(_job->callback.release.create(wil::EventOptions::ManualReset)))
        {
            return nullptr;
        }
        _job->callback.holdAfterPublication = holdAfterPublication;
        _job->options.sizeBytes             = sizeof(FileSystemOptions);
        // Concurrency 1 keeps a bulk call's folder on the serial walker; 0 leaves the provider's own (4).
        _job->options.copyMoveMaxConcurrency = serialWalker ? 1u : 0u;
        _job->source                         = source.native();
        _job->destination                    = destination.native();
        // CopyItems names the folder the copy goes into and keeps the source's leaf, which destination ends with.
        _job->destinationFolder = destination.parent_path().native();
        _job->fileSystem        = fileSystem;
        _job->serialWalker      = serialWalker;
        _thread.reset(CreateThread(nullptr,
                                   0u,
                                   [](void* parameter) noexcept -> DWORD
        {
            auto* job = static_cast<HeldLocalCopyJob*>(parameter);
            if (job->serialWalker)
            {
                const wchar_t* const sources[] = {job->source.c_str()};
                job->hr =
                    job->fileSystem->CopyItems(sources, 1u, job->destinationFolder.c_str(), FILESYSTEM_FLAG_RECURSIVE, &job->options, &job->callback, nullptr);
                return 0u;
            }
            job->hr =
                job->fileSystem->CopyItem(job->source.c_str(), job->destination.c_str(), FILESYSTEM_FLAG_RECURSIVE, &job->options, &job->callback, nullptr);
            return 0u;
        },
                                   _job.get(),
                                   0u,
                                   nullptr));
        return _thread ? _job.get() : nullptr;
    }

    [[nodiscard]] bool WaitHeld() const noexcept
    {
        return _job && _thread && WaitForSingleObject(_job->callback.held.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(30'000u))) == WAIT_OBJECT_0;
    }

    // Releases the held report and waits for the copy; false when it did not end (the job then stays allocated).
    [[nodiscard]] bool Finish() noexcept
    {
        if (! _thread)
        {
            return true;
        }
        _job->callback.release.SetEvent();
        const bool ended = WaitForSingleObject(_thread.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(60'000u))) == WAIT_OBJECT_0;
        if (! ended)
        {
            static_cast<void>(_job.release()); // still in use by the copy
        }
        _thread.reset();
        return ended;
    }

private:
    std::unique_ptr<HeldLocalCopyJob> _job;
    wil::unique_handle _thread;
};

// I35 NAS nested Copy: the Local walkers keep no handle on a folder they created while the folder's children copy.
// A chain tree\a\b\c with one file is copied with that file's first progress report held; meanwhile every
// folder already published must open with share mode 0 (read access, so the open takes part in sharing). The
// handle with DELETE access the walkers used to keep until a folder's children finished made that open fail
// with ERROR_SHARING_VIOLATION, and an SMB server refuses every rename inside a folder held that way. Once the
// copy has ended, every folder must have its source's last-write time back. Both walkers run (CopyItem: the
// producer/worker walker; CopyItems at concurrency 1: the serial walker), and every failing walker is reported.
[[nodiscard]] bool RunFileOpsLocalCopyHoldsNoDirectoryHandleChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Local Copy holds no folder handle: {}", reason));
        return false;
    };
    if (! state.fsLocal || ! state.infoLocal || ! SetPluginConfiguration(state.infoLocal.get(), kLocalManualCopyConcurrency4ForSelfTest))
    {
        return reject(L"Local provider or its manual copy/move concurrency 4 unavailable");
    }
    const std::filesystem::path base   = state.tempRoot / L"walker-no-folder-handle";
    const std::filesystem::path source = base / L"src" / L"tree";
    const auto removeBase              = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(base)); });
    const std::filesystem::path held   = std::filesystem::path(L"a") / L"b" / L"c" / L"held.bin";
    // tree, tree\a, tree\a\b and tree\a\b\c below a root, and each one's source time.
    constexpr size_t kFolders = 4u;
    const auto folderAt       = [](const std::filesystem::path& root, size_t depth) noexcept
    {
        std::filesystem::path path = root;
        for (size_t level = 0u; level < depth; ++level)
        {
            path /= std::wstring(1u, static_cast<wchar_t>(L'a' + level));
        }
        return path;
    };
    const auto folderTime = [](size_t depth) noexcept { return kSelfTestYear2001Ticks + static_cast<int64_t>(depth) * kSelfTestDayTicks; };
    if (! RecreateEmptyDirectory(base) || ! WriteFilledTestFile(source / held, 256u * 1024u, 0x5Du))
    {
        return reject(L"fixture unavailable");
    }
    for (size_t depth = 0u; depth < kFolders; ++depth)
    {
        if (! SetDirectoryLastWriteTimeForSelfTest(folderAt(source, depth), folderTime(depth)))
        {
            return reject(L"fixture times unavailable");
        }
    }

    std::wstring failures;
    for (const bool serialWalker : {false, true})
    {
        const std::wstring_view walker          = serialWalker ? L"serial walker" : L"producer/worker walker";
        const std::filesystem::path destination = base / (serialWalker ? L"dst-serial" : L"dst-parallel");
        const std::filesystem::path copied      = destination / L"tree";
        if (! RecreateEmptyDirectory(destination))
        {
            return reject(std::format(L"{}: destination unavailable", walker));
        }
        HeldLocalCopyRun run;
        HeldLocalCopyJob* const job = run.Start(state.fsLocal.get(), source, copied, false, serialWalker);
        if (job == nullptr || ! run.WaitHeld())
        {
            return reject(std::format(L"{}: the copy did not report its file's progress", walker));
        }
        std::wstring refused;
        for (size_t depth = 0u; depth < kFolders; ++depth)
        {
            const std::filesystem::path directory = folderAt(copied, depth);
            DWORD error                           = ERROR_SUCCESS;
            for (unsigned int attempt = 0u; attempt < 20u; ++attempt)
            {
                // Closed again at once: the copy goes on only after every folder was checked.
                wil::unique_hfile exclusive(CreateFileW(MakeExtendedPathForSelfTest(directory).c_str(),
                                                        GENERIC_READ,
                                                        0u,
                                                        nullptr,
                                                        OPEN_EXISTING,
                                                        FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                                        nullptr));
                error = exclusive ? ERROR_SUCCESS : GetLastError();
                if (error != ERROR_SHARING_VIOLATION)
                {
                    break;
                }
                // Another program (a scanner) may hold the folder for a moment; the walker's own handle stays.
                Sleep(50u);
            }
            if (error != ERROR_SUCCESS)
            {
                refused += std::format(L" '{}' (error {})", directory.lexically_relative(destination).native(), error);
            }
        }
        if (! run.Finish())
        {
            return reject(std::format(L"{}: the copy did not end after its progress report was released", walker));
        }
        if (! refused.empty())
        {
            failures += std::format(L" [{}: folders the walker had published could not be opened exclusively while their children copied:{}]", walker, refused);
            continue;
        }
        if (job->hr != S_OK || ! FilesEqualBytes(source / held, copied / held))
        {
            failures += std::format(L" [{}: the copy did not complete: hr=0x{:08X}]", walker, static_cast<unsigned long>(job->hr));
            continue;
        }
        std::wstring untimed;
        for (size_t depth = 0u; depth < kFolders; ++depth)
        {
            const std::filesystem::path directory = folderAt(copied, depth);
            if (DirectoryLastWriteTimeForSelfTest(directory) != folderTime(depth))
            {
                untimed += std::format(L" '{}'", directory.lexically_relative(destination).native());
            }
        }
        if (! untimed.empty())
        {
            failures += std::format(L" [{}: folders without their source's last-write time once the copy ended:{}]", walker, untimed);
        }
    }
    return failures.empty() ? true : reject(failures);
}

// I35 NAS nested Copy: holding no handle, the walker finds a folder it created again by its path when the
// folder's children have finished, and restores its basic information only when the path still names that
// folder. tree\folder\last.bin is copied with the report the walker makes after publishing last.bin held, so
// the folder's restore is pending. Runs: unchanged (restored); the folder renamed away and another folder
// created at its path (the replacement keeps its time and stays empty, the renamed folder keeps last.bin, and
// the receipt reports the folder's times lost with ERROR_FILE_INVALID, which the host warns about, as for the
// host bridge); the folder replaced by a junction (refused with ERROR_DIRECTORY, its target untouched). With
// FileIdInfo unsupported (injection) the volume serial number
// and file index prove the unchanged run and refuse the replaced one. Every run goes through both walkers
// (CopyItem: the producer/worker walker; CopyItems at concurrency 1: the serial walker), and every failing run
// is reported.
[[nodiscard]] bool RunFileOpsLocalDirectoryRestoreChecksIdentityChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Local folder restore checks identity: {}", reason));
        return false;
    };
    if (! state.fsLocal || ! state.infoLocal || ! SetPluginConfiguration(state.infoLocal.get(), kLocalManualCopyConcurrency4ForSelfTest))
    {
        return reject(L"Local provider or its manual copy/move concurrency 4 unavailable");
    }
    enum class Replacement : uint8_t
    {
        None,
        Directory,
        Junction,
    };
    struct Run final
    {
        const wchar_t* name;
        bool withoutFileId;
        Replacement replacement;
    };
    constexpr std::array<Run, 5> kRuns{{
        {L"file-id-unchanged", false, Replacement::None},
        {L"file-id-replaced", false, Replacement::Directory},
        {L"file-id-junction", false, Replacement::Junction},
        {L"file-index-unchanged", true, Replacement::None},
        {L"file-index-replaced", true, Replacement::Directory},
    }};
    constexpr int64_t kTreeTime        = kSelfTestYear2001Ticks;
    constexpr int64_t kFolderTime      = kSelfTestYear2001Ticks + kSelfTestDayTicks;
    constexpr int64_t kReplacementTime = kSelfTestYear2001Ticks + 791 * kSelfTestDayTicks;
    constexpr int64_t kTargetTime      = kSelfTestYear2001Ticks + 1000 * kSelfTestDayTicks;
    const std::filesystem::path root   = state.tempRoot / L"walker-restore-identity";
    const auto removeRoot              = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    std::wstring failures;
    for (const bool serialWalker : {false, true})
    {
        const std::wstring_view walker = serialWalker ? L"serial" : L"parallel";
        for (const Run& run : kRuns)
        {
            const std::wstring name                 = std::format(L"{}/{}", walker, run.name);
            const std::filesystem::path base        = root / walker / run.name;
            const std::filesystem::path source      = base / L"src" / L"tree";
            const std::filesystem::path destination = base / L"dst";
            const std::filesystem::path copied      = destination / L"tree";
            const std::filesystem::path folder      = copied / L"folder";
            const std::filesystem::path moved       = copied / L"folder-moved";
            const std::filesystem::path target      = base / L"junction-target";
            if (! RecreateEmptyDirectory(base) || ! WriteFilledTestFile(source / L"folder" / L"last.bin", 64u * 1024u, 0x6Bu) ||
                ! RecreateEmptyDirectory(destination) || ! RecreateEmptyDirectory(target) ||
                ! SetDirectoryLastWriteTimeForSelfTest(source / L"folder", kFolderTime) || ! SetDirectoryLastWriteTimeForSelfTest(source, kTreeTime) ||
                ! SetDirectoryLastWriteTimeForSelfTest(target, kTargetTime))
            {
                return reject(std::format(L"{}: fixture unavailable", name));
            }
            if (run.withoutFileId && SetEnvironmentVariableW(kSelfTestEnvBindIdentityUnsupportedPath.data(), destination.c_str()) == FALSE)
            {
                return reject(std::format(L"{}: the FileIdInfo-unsupported injection could not be set", name));
            }
            // Declared before the copy, so the injection stays until the copy has ended.
            const auto clearInjection =
                wil::scope_exit([]() noexcept { static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvBindIdentityUnsupportedPath.data(), nullptr)); });

            HeldLocalCopyRun copy;
            HeldLocalCopyJob* const job = copy.Start(state.fsLocal.get(), source, copied, true, serialWalker);
            if (job == nullptr || ! copy.WaitHeld())
            {
                return reject(std::format(L"{}: the walker never reported last.bin published", name));
            }
            if (run.replacement != Replacement::None)
            {
                DWORD moveError = ERROR_SUCCESS;
                for (unsigned int attempt = 0u; attempt < 20u; ++attempt)
                {
                    moveError = MoveFileExW(MakeExtendedPathForSelfTest(folder).c_str(), MakeExtendedPathForSelfTest(moved).c_str(), 0u) != FALSE
                                    ? ERROR_SUCCESS
                                    : GetLastError();
                    if (moveError != ERROR_SHARING_VIOLATION && moveError != ERROR_ACCESS_DENIED)
                    {
                        break;
                    }
                    Sleep(50u); // a scanner may hold last.bin for a moment
                }
                const bool replaced = moveError == ERROR_SUCCESS && (run.replacement == Replacement::Directory
                                                                         ? CreateDirectoryW(MakeExtendedPathForSelfTest(folder).c_str(), nullptr) != FALSE &&
                                                                               SetDirectoryLastWriteTimeForSelfTest(folder, kReplacementTime)
                                                                         : TryCreateJunction(folder, target));
                if (! replaced)
                {
                    return reject(std::format(L"{}: the published folder could not be renamed away and replaced (error {})", name, moveError));
                }
            }
            if (! copy.Finish())
            {
                return reject(std::format(L"{}: the copy did not end after its progress report was released", name));
            }
            const HeldLocalCopyCallback& callback = job->callback;
            if (job->hr != S_OK || callback.itemStatus != S_OK || ! callback.receipt.has_value() ||
                ! FileSystemItemMutationResultHasMetadataOutcome(callback.receipt.value()))
            {
                failures += std::format(L" [{}: the copy did not complete with a receipt: hr=0x{:08X} item=0x{:08X}]",
                                        name,
                                        static_cast<unsigned long>(job->hr),
                                        static_cast<unsigned long>(callback.itemStatus));
                continue;
            }
            const FileSystemItemMutationResult& receipt = callback.receipt.value();
            const bool basicLost                        = (receipt.metadataLostFeatures & FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES) != 0u;
            // A refused restore is a lost write the host warns about (metadata.basic.lost): the receipt carries the refusal
            // (another folder at the path, or a link), never the unsupported status of a destination that keeps no timestamps.
            const HRESULT expectedRefusal   = HRESULT_FROM_WIN32(run.replacement == Replacement::Junction ? ERROR_DIRECTORY : ERROR_FILE_INVALID);
            const bool basicLostWithRefusal = basicLost && receipt.metadataFirstFailure == expectedRefusal;
            if (DirectoryLastWriteTimeForSelfTest(copied) != kTreeTime)
            {
                failures += std::format(L" [{}: the copied tree did not get its source's last-write time back]", name);
                continue;
            }
            if (run.replacement == Replacement::None)
            {
                if (DirectoryLastWriteTimeForSelfTest(folder) != kFolderTime || basicLost)
                {
                    failures += std::format(L" [{}: the unchanged folder was not restored (time {}, lost=0x{:X})]",
                                            name,
                                            DirectoryLastWriteTimeForSelfTest(folder),
                                            receipt.metadataLostFeatures);
                }
                continue;
            }
            if (! FilesEqualBytes(source / L"folder" / L"last.bin", moved / L"last.bin"))
            {
                failures += std::format(L" [{}: the folder renamed away did not keep its child]", name);
                continue;
            }
            if (run.replacement == Replacement::Directory)
            {
                std::vector<std::wstring> replacementEntries;
                if (DirectoryLastWriteTimeForSelfTest(folder) != kReplacementTime || ! CollectRelativeTreeForSelfTest(folder, replacementEntries) ||
                    ! replacementEntries.empty())
                {
                    failures += std::format(L" [{}: the restore touched the folder that replaced the one the walker created (time {})]",
                                            name,
                                            DirectoryLastWriteTimeForSelfTest(folder));
                    continue;
                }
            }
            else
            {
                const DWORD attributes = GetFileAttributesW(MakeExtendedPathForSelfTest(folder).c_str());
                if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0u ||
                    DirectoryLastWriteTimeForSelfTest(target) != kTargetTime)
                {
                    failures += std::format(L" [{}: the restore touched the junction that replaced the folder, or its target]", name);
                    continue;
                }
            }
            if (! basicLostWithRefusal)
            {
                failures += std::format(L" [{}: a refused restore must report the folder's times lost with its refusal 0x{:08X}, which the host "
                                        L"warns about (lost=0x{:X}, first failure=0x{:08X})]",
                                        name,
                                        static_cast<unsigned long>(expectedRefusal),
                                        receipt.metadataLostFeatures,
                                        static_cast<unsigned long>(receipt.metadataFirstFailure));
            }
        }
    }
    return failures.empty() ? true : reject(failures);
}

// tree, tree\a, tree\a\b and so on below `root`: depth 0 is `root` itself.
[[nodiscard]] std::filesystem::path ChainFolderForSelfTest(const std::filesystem::path& root, size_t depth)
{
    std::filesystem::path path = root;
    for (size_t level = 0u; level < depth; ++level)
    {
        path /= std::wstring(1u, static_cast<wchar_t>(L'a' + level));
    }
    return path;
}

// Copies `source` into `destinationFolder` through the host bridge (the destination provider named explicitly: the same
// Local instance), Verify On selecting the sequential walker and Verify Off the parallel one, with the bridge held just
// before it restores the created folder `heldFolder`; runs `onHeld` while it is held, then lets the copy finish. Empty
// when the copy reached the hold and finished; otherwise what went wrong. `info` and `diagnostics` describe the task.
[[nodiscard]] std::wstring RunHeldBridgeCopyForSelfTest(SelfTestState& state,
                                                        FolderWindow::FileOperationState* operations,
                                                        const std::filesystem::path& source,
                                                        const std::filesystem::path& destinationFolder,
                                                        bool sequentialWalker,
                                                        const std::filesystem::path& heldFolder,
                                                        const std::function<void()>& onHeld,
                                                        CompletedTaskInfo& info,
                                                        std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry>& diagnostics) noexcept
{
    info = {};
    diagnostics.clear();
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = sequentialWalker;
    SetFileOpsBridgeDirectoryRestorePauseForSelfTest(heldFolder.native());
    const auto disarm                    = wil::scope_exit([]() noexcept { SetFileOpsBridgeDirectoryRestorePauseForSelfTest({}); });
    const std::optional<uint64_t> taskId = StartFileOperationAndGetId(operations,
                                                                      FILESYSTEM_COPY,
                                                                      FolderWindow::Pane::Left,
                                                                      FolderWindow::Pane::Right,
                                                                      state.fsLocal,
                                                                      {source},
                                                                      destinationFolder,
                                                                      FILESYSTEM_FLAG_RECURSIVE,
                                                                      false,
                                                                      0u,
                                                                      FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                      false,
                                                                      state.fsLocal);
    if (! taskId.has_value())
    {
        return L"the bridge Copy was not admitted";
    }
    const auto forget = wil::scope_exit([&]() noexcept
    {
        if (CancelOutstandingSelfTestOperations(operations))
        {
            operations->DismissCompletedTask(taskId.value());
            operations->DebugRemoveDiagnosticsForTask(taskId.value());
            state.completedTasks.erase(taskId.value());
        }
    });
    std::optional<FolderWindow::FileOperationState::Task::ConflictPromptState> prompt;
    const auto finishedOrAsked = [&]() noexcept
    {
        if (FolderWindow::FileOperationState::Task* task = operations->FindTask(taskId.value()); task != nullptr)
        {
            prompt = TryGetConflictPromptCopy(task);
        }
        return prompt.has_value() || (state.completedTasks.contains(taskId.value()) && operations->FindTask(taskId.value()) == nullptr);
    };
    bool held = false;
    static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        held = HasFileOpsBridgeDirectoryRestorePauseEnteredForSelfTest();
        return held || finishedOrAsked();
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(60)), .operationName = L"bridge Copy reaching the held folder restore"}));
    if (held && onHeld)
    {
        onHeld();
    }
    ReleaseFileOpsBridgeDirectoryRestorePauseForSelfTest();
    const auto finished = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept { return finishedOrAsked(); }, {.timeout = SelfTest::Scale(std::chrono::seconds(60)), .operationName = L"held bridge Copy"});
    diagnostics = CollectTaskDiagnosticsForSelfTest(operations, taskId.value());
    if (prompt.has_value())
    {
        return std::format(L"the bridge Copy stopped on a decision: status=0x{:08X} destination='{}'{}",
                           static_cast<unsigned long>(prompt->status),
                           prompt->destinationPath,
                           DescribeTaskDiagnosticsForSelfTest(diagnostics));
    }
    const auto completed = state.completedTasks.find(taskId.value());
    if (! finished.conditionMet || completed == state.completedTasks.end())
    {
        return std::format(L"the bridge Copy did not finish{}", DescribeTaskDiagnosticsForSelfTest(diagnostics));
    }
    info = completed->second;
    if (! held)
    {
        return std::format(L"the bridge never reached the restore of '{}' (hr=0x{:08X}, folders ensured {}){}",
                           heldFolder.filename().native(),
                           static_cast<unsigned long>(info.hr),
                           info.bridgeDirectoryEnsureCount,
                           DescribeTaskDiagnosticsForSelfTest(diagnostics));
    }
    return {};
}

// I35 NAS nested Copy, host bridge: the bridge keeps no handle on a folder it created while the folder's children copy.
// tree\a\b\c\held.bin is copied through the bridge with the restore of tree\a\b\c held (its child done; tree, a and b
// waiting for it); meanwhile every created folder must open with share mode 0 and read access (an open that takes part
// in sharing). The created folder's authority the bridge kept, with DELETE access, until the folder's children finished
// made that open fail with ERROR_SHARING_VIOLATION, and an SMB server refuses every rename inside a folder held that way.
// Once the copy has ended every folder must have its source's last-write time back. Both bridge walkers run (Verify Off:
// parallel; Verify On: sequential), and every failing walker is reported.
[[nodiscard]] bool RunFileOpsBridgeCopyHoldsNoDirectoryHandleChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Bridge Copy holds no folder handle: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
    {
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        return reject(L"no-activation guard unavailable");
    }
    const std::filesystem::path base   = state.tempRoot / L"bridge-no-folder-handle";
    const std::filesystem::path source = base / L"src" / L"tree";
    const std::filesystem::path held   = std::filesystem::path(L"a") / L"b" / L"c" / L"held.bin";
    const bool verify                  = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
    const auto restore                 = wil::scope_exit([&]() noexcept
    {
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verify;
        static_cast<void>(SelfTest::RemoveAll(base));
    });
    constexpr size_t kFolders          = 4u;
    const auto folderTime              = [](size_t depth) noexcept { return kSelfTestYear2001Ticks + static_cast<int64_t>(depth) * kSelfTestDayTicks; };
    if (! RecreateEmptyDirectory(base) || ! WriteFilledTestFile(source / held, 256u * 1024u, 0x5Du))
    {
        return reject(L"fixture unavailable");
    }
    for (size_t depth = 0u; depth < kFolders; ++depth)
    {
        if (! SetDirectoryLastWriteTimeForSelfTest(ChainFolderForSelfTest(source, depth), folderTime(depth)))
        {
            return reject(L"fixture times unavailable");
        }
    }

    std::wstring failures;
    for (const bool sequentialWalker : {false, true})
    {
        const std::wstring_view walker          = sequentialWalker ? L"sequential walker (Verify On)" : L"parallel walker";
        const std::filesystem::path destination = base / (sequentialWalker ? L"dst-sequential" : L"dst-parallel");
        const std::filesystem::path copied      = destination / L"tree";
        if (! RecreateEmptyDirectory(destination))
        {
            return reject(std::format(L"{}: destination unavailable", walker));
        }
        std::wstring refused;
        const auto probe = [&]()
        {
            for (size_t depth = 0u; depth < kFolders; ++depth)
            {
                const std::filesystem::path directory = ChainFolderForSelfTest(copied, depth);
                DWORD error                           = ERROR_SUCCESS;
                for (unsigned int attempt = 0u; attempt < 20u; ++attempt)
                {
                    // Closed again at once: the copy goes on only after every folder was checked.
                    const wil::unique_hfile exclusive(CreateFileW(MakeExtendedPathForSelfTest(directory).c_str(),
                                                                  GENERIC_READ,
                                                                  0u,
                                                                  nullptr,
                                                                  OPEN_EXISTING,
                                                                  FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                                                  nullptr));
                    error = exclusive ? ERROR_SUCCESS : GetLastError();
                    if (error != ERROR_SHARING_VIOLATION)
                    {
                        break;
                    }
                    // Another program (a scanner) may hold the folder for a moment; the bridge's own handle stays.
                    Sleep(50u);
                }
                if (error != ERROR_SUCCESS)
                {
                    refused += std::format(L" '{}' (error {})", directory.lexically_relative(destination).native(), error);
                }
            }
        };
        CompletedTaskInfo info{};
        std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> diagnostics;
        const std::wstring failure = RunHeldBridgeCopyForSelfTest(
            state, operations, source, destination, sequentialWalker, ChainFolderForSelfTest(copied, kFolders - 1u), probe, info, diagnostics);
        if (! failure.empty())
        {
            failures += std::format(L" [{}: {}]", walker, failure);
            continue;
        }
        if (! refused.empty())
        {
            failures += std::format(L" [{}: folders the bridge had published could not be opened exclusively while their children copied:{}]", walker, refused);
            continue;
        }
        if (info.hr != S_OK || ! FilesEqualBytes(source / held, copied / held))
        {
            failures += std::format(L" [{}: the copy did not complete: hr=0x{:08X}{}]",
                                    walker,
                                    static_cast<unsigned long>(info.hr),
                                    DescribeTaskDiagnosticsForSelfTest(diagnostics));
            continue;
        }
        std::wstring untimed;
        for (size_t depth = 0u; depth < kFolders; ++depth)
        {
            const std::filesystem::path directory = ChainFolderForSelfTest(copied, depth);
            if (DirectoryLastWriteTimeForSelfTest(directory) != folderTime(depth))
            {
                untimed += std::format(L" '{}'", directory.lexically_relative(destination).native());
            }
        }
        if (! untimed.empty())
        {
            failures += std::format(L" [{}: folders without their source's last-write time once the copy ended:{}{}]",
                                    walker,
                                    untimed,
                                    DescribeTaskDiagnosticsForSelfTest(diagnostics));
        }
    }
    return failures.empty() ? true : reject(failures);
}

// I35 NAS nested Copy, host bridge: holding no handle, the bridge binds a created folder again by its path when the
// folder's children have finished, and restores its basic information only through a bound object whose identity is the
// one recorded at publication. tree\folder\last.bin is copied through the bridge with the restore of tree\folder held.
// Runs: unchanged (restored); the folder renamed away and another folder created at its path (the replacement keeps its
// time and stays empty, the renamed folder keeps last.bin, and the task warns that the folder's metadata was lost); the
// folder replaced by a junction (refused, its target untouched); and, with FileIdInfo unsupported on the destination
// (injection, as on the NAS), unchanged: the folder cannot be bound again, so nothing is written to it and the task
// carries the one note that the destination keeps no timestamps. Every run goes through both walkers (Verify Off:
// parallel; Verify On: sequential), and every failing run is reported.
[[nodiscard]] bool RunFileOpsBridgeDirectoryRestoreChecksIdentityChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Bridge folder restore checks identity: {}", reason));
        return false;
    };
    auto* folderWindow = TryGetFolderWindow(state.mainWindow);
    auto* operations   = TryGetFileOps(folderWindow);
    if (! folderWindow || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
    {
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        return reject(L"no-activation guard unavailable");
    }
    enum class Replacement : uint8_t
    {
        None,
        Directory,
        Junction,
    };
    struct Run final
    {
        const wchar_t* name;
        bool withoutFileId;
        Replacement replacement;
    };
    constexpr std::array<Run, 4> kRuns{{
        {L"unchanged", false, Replacement::None},
        {L"replaced", false, Replacement::Directory},
        {L"junction", false, Replacement::Junction},
        {L"identity-less", true, Replacement::None},
    }};
    constexpr int64_t kTreeTime        = kSelfTestYear2001Ticks;
    constexpr int64_t kFolderTime      = kSelfTestYear2001Ticks + kSelfTestDayTicks;
    constexpr int64_t kReplacementTime = kSelfTestYear2001Ticks + 791 * kSelfTestDayTicks;
    constexpr int64_t kTargetTime      = kSelfTestYear2001Ticks + 1000 * kSelfTestDayTicks;
    const std::filesystem::path root   = state.tempRoot / L"bridge-restore-identity";
    const bool verify                  = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
    const auto restore                 = wil::scope_exit([&]() noexcept
    {
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verify;
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvBindIdentityUnsupportedPath.data(), nullptr));
        static_cast<void>(SelfTest::RemoveAll(root));
    });
    std::wstring failures;
    for (const bool sequentialWalker : {false, true})
    {
        const std::wstring_view walker = sequentialWalker ? L"sequential" : L"parallel";
        for (const Run& run : kRuns)
        {
            const std::wstring name                 = std::format(L"{}/{}", walker, run.name);
            const std::filesystem::path base        = root / walker / run.name;
            const std::filesystem::path source      = base / L"src" / L"tree";
            const std::filesystem::path destination = base / L"dst";
            const std::filesystem::path copied      = destination / L"tree";
            const std::filesystem::path folder      = copied / L"folder";
            const std::filesystem::path moved       = copied / L"folder-moved";
            const std::filesystem::path target      = base / L"junction-target";
            if (! RecreateEmptyDirectory(base) || ! WriteFilledTestFile(source / L"folder" / L"last.bin", 64u * 1024u, 0x6Bu) ||
                ! RecreateEmptyDirectory(destination) || ! RecreateEmptyDirectory(target) ||
                ! SetDirectoryLastWriteTimeForSelfTest(source / L"folder", kFolderTime) || ! SetDirectoryLastWriteTimeForSelfTest(source, kTreeTime) ||
                ! SetDirectoryLastWriteTimeForSelfTest(target, kTargetTime))
            {
                return reject(std::format(L"{}: fixture unavailable", name));
            }
            if (run.withoutFileId && SetEnvironmentVariableW(kSelfTestEnvBindIdentityUnsupportedPath.data(), destination.c_str()) == FALSE)
            {
                return reject(std::format(L"{}: the FileIdInfo-unsupported injection could not be set", name));
            }
            // Declared before the copy, so the injection stays until the copy has ended.
            const auto clearInjection =
                wil::scope_exit([]() noexcept { static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvBindIdentityUnsupportedPath.data(), nullptr)); });

            std::wstring replaceFailure;
            const auto replace = [&]()
            {
                if (run.replacement == Replacement::None)
                {
                    return;
                }
                DWORD moveError = ERROR_SUCCESS;
                for (unsigned int attempt = 0u; attempt < 20u; ++attempt)
                {
                    moveError = MoveFileExW(MakeExtendedPathForSelfTest(folder).c_str(), MakeExtendedPathForSelfTest(moved).c_str(), 0u) != FALSE
                                    ? ERROR_SUCCESS
                                    : GetLastError();
                    if (moveError != ERROR_SHARING_VIOLATION && moveError != ERROR_ACCESS_DENIED)
                    {
                        break;
                    }
                    Sleep(50u); // a scanner may hold last.bin for a moment
                }
                const bool replaced = moveError == ERROR_SUCCESS && (run.replacement == Replacement::Directory
                                                                         ? CreateDirectoryW(MakeExtendedPathForSelfTest(folder).c_str(), nullptr) != FALSE &&
                                                                               SetDirectoryLastWriteTimeForSelfTest(folder, kReplacementTime)
                                                                         : TryCreateJunction(folder, target));
                if (! replaced)
                {
                    replaceFailure = std::format(L"the published folder could not be renamed away and replaced (error {})", moveError);
                }
            };
            CompletedTaskInfo info{};
            std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> diagnostics;
            const std::wstring failure =
                RunHeldBridgeCopyForSelfTest(state, operations, source, destination, sequentialWalker, folder, replace, info, diagnostics);
            if (! failure.empty())
            {
                failures += std::format(L" [{}: {}]", name, failure);
                continue;
            }
            if (! replaceFailure.empty())
            {
                failures += std::format(L" [{}: {}]", name, replaceFailure);
                continue;
            }
            if (info.hr != S_OK)
            {
                failures += std::format(L" [{}: the copy did not complete: hr=0x{:08X}{}]",
                                        name,
                                        static_cast<unsigned long>(info.hr),
                                        DescribeTaskDiagnosticsForSelfTest(diagnostics));
                continue;
            }
            const bool lost             = HasDiagnosticCategoryForSelfTest(diagnostics, L"metadata.basic.lost");
            const bool notKept          = HasDiagnosticCategoryForSelfTest(diagnostics, L"metadata.basic.unsupported");
            const int64_t folderTimeNow = DirectoryLastWriteTimeForSelfTest(folder);
            const int64_t treeTimeNow   = DirectoryLastWriteTimeForSelfTest(copied);
            if (run.withoutFileId)
            {
                // Not bound again: nothing written, and the one note names it.
                if (! notKept || lost || folderTimeNow == kFolderTime || treeTimeNow == kTreeTime)
                {
                    failures += std::format(L" [{}: a folder the bridge cannot bind again must keep the copy's time and carry the "
                                            L"destination-keeps-no-timestamps note (note={}, warning={}, folder time {}, tree time {}){}]",
                                            name,
                                            notKept,
                                            lost,
                                            folderTimeNow,
                                            treeTimeNow,
                                            DescribeTaskDiagnosticsForSelfTest(diagnostics));
                }
                continue;
            }
            if (treeTimeNow != kTreeTime)
            {
                failures += std::format(
                    L" [{}: the copied tree did not get its source's last-write time back{}]", name, DescribeTaskDiagnosticsForSelfTest(diagnostics));
                continue;
            }
            if (run.replacement == Replacement::None)
            {
                if (folderTimeNow != kFolderTime || lost || notKept)
                {
                    failures += std::format(L" [{}: the unchanged folder was not restored (time {}, warning={}, note={}){}]",
                                            name,
                                            folderTimeNow,
                                            lost,
                                            notKept,
                                            DescribeTaskDiagnosticsForSelfTest(diagnostics));
                }
                continue;
            }
            if (! FilesEqualBytes(source / L"folder" / L"last.bin", moved / L"last.bin"))
            {
                failures += std::format(L" [{}: the folder renamed away did not keep its child]", name);
                continue;
            }
            if (run.replacement == Replacement::Directory)
            {
                std::vector<std::wstring> replacementEntries;
                if (folderTimeNow != kReplacementTime || ! CollectRelativeTreeForSelfTest(folder, replacementEntries) || ! replacementEntries.empty())
                {
                    failures += std::format(L" [{}: the restore touched the folder that replaced the one the bridge created (time {})]", name, folderTimeNow);
                    continue;
                }
            }
            else
            {
                const DWORD attributes = GetFileAttributesW(MakeExtendedPathForSelfTest(folder).c_str());
                if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0u ||
                    DirectoryLastWriteTimeForSelfTest(target) != kTargetTime)
                {
                    failures += std::format(L" [{}: the restore touched the junction that replaced the folder, or its target]", name);
                    continue;
                }
            }
            if (! lost)
            {
                failures += std::format(L" [{}: a refused restore must warn that the folder's metadata was lost (note={}){}]",
                                        name,
                                        notKept,
                                        DescribeTaskDiagnosticsForSelfTest(diagnostics));
            }
        }
    }
    return failures.empty() ? true : reject(failures);
}

// I35 NAS nested Copy, host bridge perf witness: the 1,885-folder tree of R4T4 copied twice through the host bridge
// (Verify Off: the parallel walker). Every copy must hold exactly the source tree and give every folder its source's
// last-write time back. FileOps.SelfTest.BridgeManyDirectoriesCopy records each copy's wall time.
[[nodiscard]] bool RunFileOpsBridgeManyDirectoriesCopyChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Bridge many-folders Copy: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
    {
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        return reject(L"no-activation guard unavailable");
    }
    const std::filesystem::path base        = state.tempRoot / L"bridge-many-folders";
    const std::filesystem::path source      = base / L"src" / L"tree";
    const std::filesystem::path destination = base / L"dst";
    const bool verify                       = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
    std::optional<uint64_t> taskId;
    const auto forgetTask = [&]() noexcept
    {
        if (taskId.has_value() && CancelOutstandingSelfTestOperations(operations))
        {
            operations->DismissCompletedTask(taskId.value());
            operations->DebugRemoveDiagnosticsForTask(taskId.value());
            state.completedTasks.erase(taskId.value());
        }
        taskId.reset();
    };
    const auto restore = wil::scope_exit([&]() noexcept
    {
        forgetTask();
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verify;
        static_cast<void>(SelfTest::RemoveAll(base));
    });
    std::vector<std::wstring> directories;
    size_t sourceFiles = 0u;
    if (! WriteManyDirectoriesFixtureForSelfTest(source, directories, sourceFiles))
    {
        return reject(L"source fixture unavailable");
    }
    std::vector<std::wstring> sourceEntries;
    if (! CollectRelativeTreeForSelfTest(source, sourceEntries) || sourceEntries.size() != directories.size() - 1u + sourceFiles)
    {
        return reject(L"the source tree could not be listed");
    }

    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    constexpr unsigned int kRepetitions                       = 2u;
    uint64_t totalUs                                          = 0u;
    for (unsigned int repetition = 0u; repetition < kRepetitions; ++repetition)
    {
        if (! RecreateEmptyDirectory(destination))
        {
            return reject(L"destination unavailable");
        }
        const std::filesystem::path copied = destination / source.filename();
        const auto startedAt               = std::chrono::steady_clock::now();
        taskId                             = StartFileOperationAndGetId(operations,
                                                                        FILESYSTEM_COPY,
                                                                        FolderWindow::Pane::Left,
                                                                        FolderWindow::Pane::Right,
                                                                        state.fsLocal,
                                                                        {source},
                                                                        destination,
                                                                        FILESYSTEM_FLAG_RECURSIVE,
                                                                        false,
                                                                        0u,
                                                                        FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                        false,
                                                                        state.fsLocal);
        if (! taskId.has_value())
        {
            return reject(std::format(L"copy {} was not admitted", repetition));
        }
        std::optional<FolderWindow::FileOperationState::Task::ConflictPromptState> prompt;
        const auto finished = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            if (FolderWindow::FileOperationState::Task* task = operations->FindTask(taskId.value()); task != nullptr)
            {
                prompt = TryGetConflictPromptCopy(task);
            }
            return prompt.has_value() || (state.completedTasks.contains(taskId.value()) && operations->FindTask(taskId.value()) == nullptr);
        },
            {.timeout = SelfTest::Scale(std::chrono::seconds(300)), .operationName = L"bridge many-folders Copy"});
        const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
        const auto diagnostics   = CollectTaskDiagnosticsForSelfTest(operations, taskId.value());
        const auto completed     = state.completedTasks.find(taskId.value());
        if (prompt.has_value() || ! finished.conditionMet || completed == state.completedTasks.end())
        {
            return reject(std::format(L"copy {} asked a question or did not finish{}", repetition, DescribeTaskDiagnosticsForSelfTest(diagnostics)));
        }
        const CompletedTaskInfo info = completed->second;
        forgetTask();
        Debug::Perf::Emit(L"FileOps.SelfTest.BridgeManyDirectoriesCopy",
                          L"folders=1885;files=144;walker=parallel",
                          elapsedUs,
                          info.bridgeDirectoryEnsureCount,
                          repetition,
                          info.hr);
        if (info.hr != S_OK || info.bridgeDirectoryEnsureCount < directories.size())
        {
            return reject(std::format(L"copy {} failed or did not go through the bridge: hr=0x{:08X}, folders ensured {}{}",
                                      repetition,
                                      static_cast<unsigned long>(info.hr),
                                      info.bridgeDirectoryEnsureCount,
                                      DescribeTaskDiagnosticsForSelfTest(diagnostics)));
        }
        std::vector<std::wstring> copiedEntries;
        if (! CollectRelativeTreeForSelfTest(copied, copiedEntries) ||
            ! std::ranges::equal(
                copiedEntries, sourceEntries, [](const std::wstring& left, const std::wstring& right) noexcept { return EqualsIgnoreCase(left, right); }))
        {
            return reject(
                std::format(L"copy {} differs from its source: {} entries copied, {} expected", repetition, copiedEntries.size(), sourceEntries.size()));
        }
        if (const std::wstring untimed = FirstManyDirectoriesFolderWithoutTimeForSelfTest(copied, directories); ! untimed.empty())
        {
            return reject(std::format(L"copy {}: the folder '{}' did not get its source's last-write time back", repetition, untimed));
        }
        totalUs += elapsedUs;
    }
    AppendLog(std::format(L"Bridge many-folders Copy: {} folders and {} files copied {} times through the bridge, mean {} ms",
                          directories.size(),
                          sourceFiles,
                          kRepetitions,
                          totalUs / kRepetitions / 1000u));
    return true;
}

#pragma warning(push)
#pragma warning(disable : 4061) // Each helper intentionally owns only its phase subset; unmatched steps return nullopt.
#pragma warning(disable : 4883) // This test-only dispatcher includes the intentionally large Fairstream phase corpus; runtime optimization is irrelevant.

[[nodiscard]] TickResult TickFairstreamA(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_A
#include "FolderWindow.FileOperations.Tests.Fairstream.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_A
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickFairstreamB(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_B
#include "FolderWindow.FileOperations.Tests.Fairstream.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_B
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickFairstreamC(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_C
#include "FolderWindow.FileOperations.Tests.Fairstream.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_C
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickCopyMerge(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_COPY_MERGE
#include "FolderWindow.FileOperations.Tests.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_COPY_MERGE
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickMoveMerge(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_MOVE_MERGE
#include "FolderWindow.FileOperations.Tests.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_MOVE_MERGE
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickReparseMerge(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_REPARSE_MERGE
#include "FolderWindow.FileOperations.Tests.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_REPARSE_MERGE
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickProviderMatrix(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PROVIDER_MATRIX
#include "FolderWindow.FileOperations.Tests.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PROVIDER_MATRIX
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase5(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE5
#include "FolderWindow.FileOperations.Tests.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE5
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase6(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE6
#include "FolderWindow.FileOperations.Tests.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE6
        default: return std::nullopt;
    }
}

// The provider discovery-parity fixtures start fake backends and hold them across ticks, so they
// keep their frame out of the Phase 5 dispatcher like the other large corpora.
[[nodiscard]] TickResult TickDiscoveryProviders(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_DISCOVERY_PROVIDERS
#include "FolderWindow.FileOperations.Tests.DiscoveryProviders.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_DISCOVERY_PROVIDERS
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase7(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE7
#include "FolderWindow.FileOperations.Tests.Phases07_09.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE7
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase8(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE8
#include "FolderWindow.FileOperations.Tests.Phases07_09.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE8
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase9(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE9
#include "FolderWindow.FileOperations.Tests.Phases07_09.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE9
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase10(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE10
#include "FolderWindow.FileOperations.Tests.Phases10_13.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE10
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase11(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE11
#include "FolderWindow.FileOperations.Tests.Phases10_13.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE11
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase12(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE12
#include "FolderWindow.FileOperations.Tests.Phases10_13.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE12
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase13(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE13
#include "FolderWindow.FileOperations.Tests.Phases10_13.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE13
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhases14To16(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#include "FolderWindow.FileOperations.Tests.Phases14_16.cpp"
        return std::nullopt;
    }

#pragma warning(pop)
} // namespace

#pragma warning(push)
#pragma warning(disable : 4061) // Tick dispatches all non-setup steps through the phase helpers above.
bool FileOperationsSelfTest::Tick(HWND /*mainWindow*/) noexcept
{
    SelfTestState& state = GetState();
    if (! state.running.load(std::memory_order_acquire))
    {
        return false;
    }

    if (state.done.load(std::memory_order_acquire))
    {
        return true;
    }

    // Keep the very large phase-local test fixtures off this dispatcher frame. In Debug builds,
    // placing every included phase in this one switch made Tick reserve roughly 780 KiB and left
    // too little of the default 1 MiB UI-thread stack for nested popup rendering.
    const auto dispatchPhase = [&](const auto& tickPhase) noexcept -> TickResult { return tickPhase(state); };
    if (const TickResult result = dispatchPhase(TickFairstreamA); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickFairstreamB); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickFairstreamC); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickBatchRenameExactResultCarrier); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickCopyMerge); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickMoveMerge); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickReparseMerge); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickProviderMatrix); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhase5); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhase6); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickDiscoveryProviders); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhase7); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhase8); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhase9); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhase10); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhase11); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhase12); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhase13); result.has_value())
    {
        return result.value();
    }
    if (const TickResult result = dispatchPhase(TickPhases14To16); result.has_value())
    {
        return result.value();
    }

    switch (state.step)
    {
        case SelfTestState::Step::FileOps_QueuedAdmissionContracts:
        {
            SelfTest::CaseState contracts;
            if (! ValidateQueuedAdmissionContracts(state, contracts))
            {
                Fail(contracts.failure);
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_SafetyRemediationContracts);
            return false;
        }
        case SelfTestState::Step::FileOps_SafetyRemediationContracts:
        {
            auto* operations = TryGetFileOps(TryGetFolderWindow(state.mainWindow));
            if (! operations)
                return false;
            bool contractsPassed = true;
            const auto check     = [&](std::wstring_view name, const auto& runCheck) noexcept
            {
                if (state.retainTempRootOnFailure)
                {
                    AppendLog(std::format(L"Safety contract {}: NOT EXECUTED after retained failure", name));
                    contractsPassed = false;
                    return;
                }
                const bool passed = runCheck();
                AppendLog(std::format(L"Safety contract {}: {}", name, passed ? L"passed" : L"FAILED"));
                contractsPassed = contractsPassed && passed;
            };
            check(L"mutation interlock indexed lookup differential",
                  [&]
            {
                std::wstring failure;
                const bool passed = FolderWindowFileOperationsStateInternal::RunMutationInterlockLookupDifferentialForSelfTest(failure);
                if (! passed)
                    AppendLog(std::format(L"Interlock lookup differential FAILED: {}", failure));
                return passed;
            });
            check(L"conflict broker", [&] { return RunFileOpsConflictBrokerChecks(); });
            check(L"ordinary Skip inherits Apply to all for AccessDenied across files/folders, while unchecked and different errors/profiles/roots prompt",
                  [&] { return RunFileOpsAccessDeniedSkipScopeChecks(state); });
            check(L"receipt matrix", [&] { return RunFileOperationItemResultReducerSelfTest(); });
            check(L"missing-receipt dense task results", [&] { return RunFileOperationCanceledMissingReceiptTaskSelfTest(*operations); });
            check(L"per-item mutation gate dense task results", [&] { return RunFileOperationUndispatchedRootGateTaskSelfTest(*operations); });
            check(L"descendant Skip with unknown axes completes with skipped items",
                  [&] { return RunFileOperationSkippedUnknownAxesTaskSelfTest(*operations); });
            check(L"partial Move task.result text follows typed source dispositions",
                  [&] { return RunFileOperationPartialMoveResultMessageSelfTest(*operations); });
            check(L"receipt status and aggregate", [&] { return RunFileOperationReceiptStatusTaskSelfTest(*operations); });
            check(L"a committed per-item receipt stays committed for its item", [&] { return RunFileOperationStickyCommittedReceiptSelfTest(*operations); });
            check(L"empty-selection production Copy/Move/Delete admission, exact bytes/hash and bounded inventory",
                  [&] { return RunFileOpsEmptySelectionAdmissionChecks(state); });
            check(L"actual same-object Copy KeepBoth and Move/hard-link refusal, provider identity and exact bytes/hash/inventory",
                  [&] { return RunFileOpsSameObjectBytePreservationChecks(state); });
            check(L"actual host case-only file rename and separate provider temp second-leg rollback, exact casing/file ID/bytes/BLAKE3/no leftovers",
                  [&] { return RunFileOpsCaseOnlyFileRenameStorageChecks(state); });
            check(L"actual locked overwrite refusal and post-write stage failure, original bytes/hash/file IDs and closed inventory",
                  [&] { return RunFileOpsOverwritePreservationChecks(state); });
            check(L"actual PermanentDelete committed-retained receipt, exact bytes/hash and Issues",
                  [&] { return RunFileOpsCommittedRetainedDeleteChecks(state); });
            check(L"actual serial Local recursive PermanentDelete after-child-removal cancel, exact bytes/BLAKE3 and closed surviving tree",
                  [&] { return RunFileOpsRecursiveDeleteAfterRemovalCancelChecks(state); });
            check(L"Local missing-source producer and finalizer",
                  [&] { return RunFileOperationNativeMoveMissingSourceTaskSelfTest(*operations, state.fsLocal, state.tempRoot); });
            check(L"Local empty-directory Copy no-op and publication",
                  [&] { return RunFileOperationLocalEmptyDirectoryCopyNoOpSelfTest(*operations, state.fsLocal, state.infoLocal.get(), state.tempRoot); });
            check(L"actual ordinary-child DACL denied serial recursive Copy, restored security and exact source/destination trees",
                  [&] { return RunFileOpsAclDeniedChildCopyChecks(state); });
            check(L"Issues projection and lifecycle", [&] { return RunFileOpsIssuesProjectionChecks(state); });
            check(L"cached apply-to-all keeps one Issues row and one JSONL record per originating prompt with its count, and never evicts a real error",
                  [&] { return RunFileOpsIssueReservoirAggregationChecks(state); });
            check(L"bounded Issues reservoirs evict unpaired records before real errors and whole prompt rows oldest first",
                  [&] { return RunFileOpsIssueReservoirEvictionChecks(state); });
            check(L"Issues pane shows chronological numeric-sortable rows with only applicable facts, an instruction on an empty first open, one "
                  L"selected row, CRLF copy, a named detail region, no rebuild on unprojectable records or minimize",
                  [&] { return RunFileOpsIssuesPaneUxChecks(state); });
            check(L"path-detail no-activation reconstruction", [&] { return RunFileOpsPathDetailNoActivationRebuildChecks(state); });
            check(L"task-owned path region focus never binds Escape to another task's decision",
                  [&] { return RunFileOpsTaskOwnedRegionFocusEscapeChecks(state); });
            check(L"popup announces Canceled, Outcome unknown and fast failures, counts paused tasks apart, exposes path rows as Buttons and lets "
                  L"Space activate a focused Cancel or Skip",
                  [&] { return RunFileOpsPopupStatusAccessibilityChecks(state); });
            check(L"conflict inspection follows a refresh, ends Canceled on a user cancel, breaks a same-name tie by kind and ends Stale with its "
                  L"popup, prompt, cleared pane path or disconnect",
                  [&] { return RunFileOpsConflictInspectionSettlementChecks(state); });
            const FileOperations::QualifiedEndpoint dummyEndpoint{
                .pluginId               = L"builtin/file-system-dummy",
                .instanceId             = L"host/default",
                .profileId              = L"dummy-local",
                .rootId                 = L"dummy-root",
                .pathIdentity           = FileSystemPathIdentity{.pathTextStableIdentity = true,
                                                                 .componentComparison    = FileSystemPathComponentComparison::OrdinalIgnoreCase,
                                                                 .preferredSeparator     = L'\\',
                                                                 .acceptedSeparators     = L"\\/",
                                                                 .casePreserving         = true,
                                                                 .caseOnlyRename         = FileSystemPathCaseOnlyRename::Supported},
                .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
            };
            check(L"provider cancellation after mutation",
                  [&]
            {
                return RunFileOperationCanceledProviderNullReceiptSelfTest(
                    *operations, state.fsDummy, dummyEndpoint, std::format(L"{}-{}", GetCurrentProcessId(), GetTickCount64()));
            });
            check(L"Stopping status precedence, localized summary and real RequestCancel snapshot", [&] { return RunFileOpsStoppingSnapshotChecks(state); });
            check(L"all-posts-fail owner recovery and parked-worker quiet point", [&] { return RunFileOpsCompletionRecoveryChecks(state); });
            check(L"completion callback destroys real secondary owner", [&] { return RunFileOpsCompletionOwnerDestroyChecks(state); });
            check(L"retained popup idle telemetry measurement", [&] { return RecordFileOpsPopupIdleTelemetry(state); });
            if (! contractsPassed)
            {
                Fail(L"File Operations safety contracts failed; see individual witnesses in the trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_ConflictMessageText);
            return false;
        }
        case SelfTestState::Step::FileOps_ConflictMessageText:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress       = true;
            const auto reset = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsConflictMessageTextChecks() || ! RunFileOpsIssuesTypedProblemTextChecks(state) ||
                ! RunFileOpsIssuesProducerResourceTextChecks(state))
            {
                Fail(L"File Operations conflict message text failed; see localized question and Issues witnesses in the trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_ReviewRemediationContracts);
            return false;
        }
        case SelfTestState::Step::FileOps_ReviewRemediationContracts:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress       = true;
            const auto reset = wil::scope_exit([]() noexcept { inProgress = false; });
            RedSalamander::TestSupport::ScopedWindowActivationBlocker activationBlocker;
            if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! activationBlocker.Start())
            {
                Fail(L"Review remediation requires the no-activation guard.");
                return true;
            }
            // These ownership/alias controls need no mounted disk or desktop
            // input; run them in the ordinary remediation case as well as F7.
            std::wstring ownedVolumeFailure;
            if (! AdmissionBaselineSelfTest::RunOwnedVolumeRefusalControls(state.tempRoot / L"review-owned-volume-refusal-controls", ownedVolumeFailure))
            {
                Fail(std::format(L"Owned-volume operation-path controls failed: {}", ownedVolumeFailure));
                return true;
            }
            AppendLog(L"Owned-volume operation-path controls passed without mounting a disk or taking focus.");
            // Private predicates use their own engine; the real clipboard-ready journey
            // retains its strict idle-owner precondition under the no-activation guard.
            SelfTest::CaseState admissionContracts;
            if (! ValidateQueuedAdmissionContracts(state, admissionContracts))
            {
                Fail(admissionContracts.failure);
                return true;
            }
            if (! RunFileOpsDeleteParentPinChecks(state) || ! RunFileOpsCrossRootSkipContractChecks(state) ||
                ! RunFileOperationQueuedMissingSourceTaskSelfTest(*state.fileOps, state.fsLocal, state.tempRoot) ||
                ! RunFileOperationUnverifiableDestinationCollisionSelfTest(*state.fileOps, state.fsLocal, state.tempRoot) ||
                ! RunFileOpsCompletedCardAutoCollapseChecks(state))
            {
                Fail(L"File Operations review regression failed; see the trace.");
                return true;
            }
            std::wstring presentationFailure;
            if (! DebugValidateFileOperationsMissingSourcePresentation(presentationFailure))
            {
                Fail(presentationFailure);
                return true;
            }
            if (! CaptureFileOpsWaitingCards(state))
            {
                Fail(L"Waiting-card presentation failed; see WaitingCard witnesses in the trace.");
                return true;
            }
            if (! CaptureFileOpsDocumentation(state))
            {
                Fail(L"Documentation screenshots failed; see DocumentationPrivacy witnesses in the trace.");
                return true;
            }
            if (! CaptureFileOpsCardHierarchy(state))
            {
                Fail(L"Card hierarchy failed; see CardHierarchy subject, geometry and disclosure witnesses in the trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_RepeatDecisionContracts);
            return false;
        }
        case SelfTestState::Step::FileOps_RepeatDecisionContracts:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress       = true;
            const auto reset = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! DebugFileOpsConflictActionPolicyCoverageForSelfTest() || ! RunFileOpsAccessDeniedSkipScopeChecks(state) ||
                ! RunFileOpsMetadataConsentRootScopeChecks(state) || ! RunFileOpsKeepBothNestedRepeatChecks(state))
            {
                Fail(L"File Operations repeated-choice contracts failed; see policy witnesses in the trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_SkipCountingPresentation);
            return false;
        }
        case SelfTestState::Step::FileOps_SkipCountingPresentation:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress       = true;
            const auto reset = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsSkipCountingPresentationChecks(state))
            {
                Fail(L"File Operations Skip counting presentation failed; see counter, layout and capture witnesses in the trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_PermanentDeleteCompletionProgress);
            return false;
        }
        case SelfTestState::Step::FileOps_PermanentDeleteCompletionProgress:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress       = true;
            const auto reset = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsPermanentDeleteProgressChecks(state) || ! RunFileOpsReadonlyDeleteUserConsentChecks(state) ||
                ! RunFileOpsSelectedReadonlyDeleteUserConsentChecks(state) || ! RunFileOpsBridgeCompletedFolderChecks(state))
            {
                Fail(L"File Operations permanent delete progress failed; see committed-work and readonly witnesses in the trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_AdmissionBaseline);
            return false;
        }
        case SelfTestState::Step::FileOps_AdmissionBaseline:
        {
            static bool baselineInProgress = false;
            if (baselineInProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            baselineInProgress = true;
            const auto reset   = wil::scope_exit([]() noexcept { baselineInProgress = false; });
            if (! RunFileOpsAdmissionBaseline(state))
            {
                Fail(L"Actual UI admission baseline failed; see F7 per-route/sample trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_OwnedProductionShutdown);
            return false;
        }
        case SelfTestState::Step::FileOps_OwnedProductionShutdown:
        {
            if (RedSalamander::TestSupport::GetEnvironmentString(OwnedShutdownSelfTest::kRole) == L"owned-child-v1")
                return OwnedShutdownFixture::ChildTick(state);
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress       = true;
            const auto reset = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! OwnedShutdownFixture::Parent(state))
            {
                Fail(L"Actual owned production shutdown failed; retain D10 child tree and close receipt.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_CrossVolumeMoveAfterPartialPublicationStorage);
            return false;
        }
        case SelfTestState::Step::FileOps_CrossVolumeMoveAfterPartialPublicationStorage:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            std::wstring unavailableReason;
            const bool passed = RunFileOpsManagedMovePartialWriteChecks(state, unavailableReason);
            if (! passed && ! unavailableReason.empty())
            {
                Debug::Perf::Emit(
                    L"FileOps.SelfTest.CrossVolumeMoveAfterPartialPublicationStorage", std::format(L"skip={}", unavailableReason), 0u, 0u, 0u, S_FALSE);
                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, unavailableReason);
                NextStep(state, SelfTestState::Step::FileOps_RepeatedNonClipboardCopyStorage);
                return false;
            }
            if (! passed)
            {
                Fail(L"Real cross-volume managed Move storage witness failed; see retained D2 trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_RepeatedNonClipboardCopyStorage);
            return false;
        }
        case SelfTestState::Step::FileOps_RepeatedNonClipboardCopyStorage:
        case SelfTestState::Step::FileOps_RepeatedNonClipboardMoveStorage:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            const bool move            = state.step == SelfTestState::Step::FileOps_RepeatedNonClipboardMoveStorage;
            std::wstring unavailableReason;
            const bool passed = RunFileOpsRepeatedNonClipboardStorageChecks(state, move ? FILESYSTEM_MOVE : FILESYSTEM_COPY, unavailableReason);
            const auto next   = move ? SelfTestState::Step::FileOps_TerminalFactsProjection : SelfTestState::Step::FileOps_RepeatedNonClipboardMoveStorage;
            if (! passed && ! unavailableReason.empty())
            {
                Debug::Perf::Emit(L"FileOps.SelfTest.RepeatedNonClipboardStorage", std::format(L"skip={}", unavailableReason), 0u, 0u, 0u, S_FALSE);
                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, unavailableReason);
                NextStep(state, next);
                return false;
            }
            if (! passed)
            {
                Fail(L"Repeated non-clipboard real storage witness failed; see retained D11 trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, next);
            return false;
        }
        case SelfTestState::Step::FileOps_TerminalFactsProjection:
        {
            if (! TryGetFolderWindow(state.mainWindow))
                return false;
            const std::wstring_view failure = ValidateCompletedTerminalFactsProjection(state);
            if (! failure.empty())
            {
                // Fail performs suite cleanup; the validator's popup and injected task are now gone.
                Fail(failure);
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_ClearCompletedScope);
            return false;
        }
        case SelfTestState::Step::FileOps_ClearCompletedScope:
        {
            if (! TryGetFolderWindow(state.mainWindow))
                return false;
            const std::wstring_view failure = ValidateClearCompletedScope(state);
            if (! failure.empty())
            {
                // The validator's private engine, popup and breadcrumb store are already gone.
                Fail(failure);
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_PerItemCancelUndispatchedRoots);
            return false;
        }
        case SelfTestState::Step::FileOps_PerItemCancelUndispatchedRoots:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            const bool passed          = RunFileOpsPerItemCancelUndispatchedRootsChecks(state, FILESYSTEM_COPY, 1u) &&
                                         RunFileOpsPerItemCancelUndispatchedRootsChecks(state, FILESYSTEM_MOVE, 1u) &&
                                         RunFileOpsPerItemCancelUndispatchedRootsChecks(state, FILESYSTEM_COPY, 2u) &&
                                         RunFileOpsPerItemCancelUndispatchedRootsChecks(state, FILESYSTEM_MOVE, 2u);
            if (! passed)
            {
                Fail(L"Per-item cancel violated the proved completion/dispatch states or storage; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_AdmissionFailureReachesPane);
            return false;
        }
        case SelfTestState::Step::FileOps_AdmissionFailureReachesPane:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsAdmissionFailureReachesPaneChecks(state))
            {
                Fail(L"A failed pane Copy/Move/Delete admission was dropped silently or prompted before failing; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_DeferredExitRefusesNewWork);
            return false;
        }
        case SelfTestState::Step::FileOps_DeferredExitRefusesNewWork:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsDeferredExitRefusesNewWorkChecks(state))
            {
                Fail(L"A pending application exit admitted new work, asked for consent again, or let queued work start; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_CompareWindowCloseConsent);
            return false;
        }
        case SelfTestState::Step::FileOps_CompareWindowCloseConsent:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsCompareWindowCloseConsentChecks(state))
            {
                Fail(L"A Compare Directories window close or the application exit bypassed the Compare engine's running work; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_PopupFailureSurfaceRecovery);
            return false;
        }
        case SelfTestState::Step::FileOps_PopupFailureSurfaceRecovery:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsPopupFailureSurfaceRecoveryChecks(state))
            {
                Fail(L"A transient popup frame or layout failure left the conflict decision unreachable, or the failure surface did not recover; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_PopupStableHostedTree);
            return false;
        }
        case SelfTestState::Step::FileOps_PopupStableHostedTree:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsPopupStableHostedTreeChecks(state))
            {
                Fail(L"The popup rebuilt its hosted control tree, lost a held press, moved focus or republished live text on every frame; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_PopupIdleCadence);
            return false;
        }
        case SelfTestState::Step::FileOps_PopupIdleCadence:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsPopupIdleCadenceChecks(state))
            {
                Fail(L"A paused, queued or prompt-blocked task, open path detail or retained history kept the popup or owner timer doing per-tick work; see "
                     L"trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_AlertOverlayLifetime);
            return false;
        }
        case SelfTestState::Step::FileOps_AlertOverlayLifetime:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsAlertOverlayLifetimeChecks(state))
            {
                Fail(L"An alert overlay lost its anchor tracking, kept a prompt loop alive after its owner died, leaked a window-scope entry or "
                     L"requested the foreground; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_AlertOverlayUxAccessibility);
            return false;
        }
        case SelfTestState::Step::FileOps_AlertOverlayUxAccessibility:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsAlertOverlayUxAccessibilityChecks(state))
            {
                Fail(L"An alert overlay moved its content on option activation, misplaced its icon, raised no UI Automation event, restored focus to "
                     L"itself, guessed a mnemonic, named itself in English or ran a forged UI Automation marshal call; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_AlertOverlayModelessDecision);
            return false;
        }
        case SelfTestState::Step::FileOps_AlertOverlayModelessDecision:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsAlertOverlayModelessChecks(state))
            {
                Fail(L"A modeless alert overlay took the focus it did not own, acted on keys without the focus, ran Escape from its close glyph, "
                     L"did not announce, lost its anchor tracking or requested the foreground; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_DecisionMirror);
            return false;
        }
        case SelfTestState::Step::FileOps_DecisionMirror:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsDecisionMirrorChecks(state))
            {
                Fail(state.phaseFailureMessage.empty() ? L"The decision mirror failed without a specific diagnostic; see trace."
                                                       : std::wstring_view(state.phaseFailureMessage));
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_DecisionMirrorSameHostOverlap);
            return false;
        }
        case SelfTestState::Step::FileOps_DecisionMirrorSameHostOverlap:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsDecisionMirrorOverlapChecks(state))
            {
                Fail(L"The decision mirror did not present the same-host overlap consent with the engine's default and Escape, or its Escape "
                     L"reached another task; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_DecisionMirrorIngressAnchor);
            return false;
        }
        case SelfTestState::Step::FileOps_DecisionMirrorIngressAnchor:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsDecisionMirrorIngressChecks(state))
            {
                Fail(L"A pre-task Copy or Delete prompt was not anchored on the pane the request came from; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_DecisionMirrorCompareWindow);
            return false;
        }
        case SelfTestState::Step::FileOps_DecisionMirrorCompareWindow:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsDecisionMirrorCompareWindowChecks(state))
            {
                Fail(L"A Compare Directories window's decision mirror was not independent of the main window's, or outlived its window; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_DecisionMirrorConsecutivePrompts);
            return false;
        }
        case SelfTestState::Step::FileOps_DecisionMirrorConsecutivePrompts:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsDecisionMirrorConsecutiveChecks(state))
            {
                Fail(L"A run of decisions flashed the popup's taskbar button more than once or blanked the status segment, or Apply to all on the "
                     L"mirror was not the card's engine flag; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_DecisionMirrorQueueOrder);
            return false;
        }
        case SelfTestState::Step::FileOps_DecisionMirrorQueueOrder:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsDecisionMirrorQueueOrderChecks(state))
            {
                Fail(L"The decision mirror jumped away from a shown decision or did not mirror the oldest pending decision next; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_BatchRenameExactResultCarrier);
            return false;
        }
        case SelfTestState::Step::FileOps_BatchRenameNeverRunRowsNotAttempted:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsBatchRenameNeverRunRowChecks(state))
            {
                Fail(L"Batch Rename reported a row the stopped run never reached with another row's outcome; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_NotesSeverity);
            return false;
        }
        case SelfTestState::Step::FileOps_NotesSeverity:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsNotesSeverityChecks(state))
            {
                Fail(L"A question answered, a skip the user chose or a hidden Info log level counted as a warning, or a real failure did not; see "
                     L"trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_NotesMetadata);
            return false;
        }
        case SelfTestState::Step::FileOps_NotesMetadata:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsNotesMetadataChecks(state))
            {
                Fail(L"Metadata outcomes were not classified alike on the native and bridge Copy routes, or the native route lost metadata "
                     L"silently; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_CrossProviderMtpNewFile);
            return false;
        }
        case SelfTestState::Step::R0fSmb_NestedTreeCopyToShare:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            std::wstring unavailableReason;
            const bool passed = RunFileOpsSmbNestedTreeCopyChecks(state, unavailableReason);
            if (! passed && ! unavailableReason.empty())
            {
                AppendLog(std::format(L"SMB nested tree Copy: environment-gated case skipped: {}", unavailableReason));
                Debug::Perf::Emit(L"FileOps.SelfTest.SmbNestedTreeCopy", std::format(L"skip={}", unavailableReason), 0u, 0u, 0u, S_FALSE);
                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, unavailableReason);
                NextStep(state, SelfTestState::Step::R0fSmb_BridgeTreeCopyToShare);
                return false;
            }
            if (! passed)
            {
                Fail(L"A Local Copy of a nested tree to the registered SMB share failed, was incomplete, kept a stage, or did not restore a "
                     L"folder's last-write time; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::R0fSmb_BridgeTreeCopyToShare);
            return false;
        }
        case SelfTestState::Step::R0fSmb_BridgeTreeCopyToShare:
        case SelfTestState::Step::R0fSmb_ManagedMoveTreeToShare:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            const bool move            = state.step == SelfTestState::Step::R0fSmb_ManagedMoveTreeToShare;
            const SelfTestState::Step next =
                move ? SelfTestState::Step::R0fCurl_FakeFtpReadWriteCreateDelete : SelfTestState::Step::R0fSmb_ManagedMoveTreeToShare;
            std::wstring unavailableReason;
            const bool passed = RunFileOpsSmbBridgeTreeTransferChecks(state, unavailableReason, move);
            if (! passed && ! unavailableReason.empty())
            {
                AppendLog(std::format(L"SMB bridge {}: environment-gated case skipped: {}", move ? L"Managed Move" : L"Copy", unavailableReason));
                Debug::Perf::Emit(L"FileOps.SelfTest.SmbBridgeTreeTransfer", std::format(L"skip={}", unavailableReason), 0u, 0u, 0u, S_FALSE);
                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, unavailableReason);
                NextStep(state, next);
                return false;
            }
            if (! passed)
            {
                Fail(move ? L"A Managed Move of a nested tree and a flat folder to the registered SMB share failed, was incomplete, kept a stage or "
                            L"a source, or left a publication outcome unknown; see trace."
                          : L"A bridge Copy of a nested tree and a flat folder to the registered SMB share failed, was incomplete, kept a stage, "
                            L"or left a publication outcome unknown; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, next);
            return false;
        }
        case SelfTestState::Step::R4T4_LocalManyDirectoriesCopy:
        {
            if (! RunFileOpsLocalManyDirectoriesCopyChecks(state))
            {
                Fail(L"A Local Copy of a 1,885-folder tree failed, was incomplete, or did not restore every folder's last-write time; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::R4T5_LocalCopyHoldsNoDirectoryHandle);
            return false;
        }
        case SelfTestState::Step::R4T5_LocalCopyHoldsNoDirectoryHandle:
        {
            if (! RunFileOpsLocalCopyHoldsNoDirectoryHandleChecks(state))
            {
                Fail(L"The Local Copy walker kept a folder it created open while the folder's children copied; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::R4T6_LocalDirectoryRestoreChecksIdentity);
            return false;
        }
        case SelfTestState::Step::R4T6_LocalDirectoryRestoreChecksIdentity:
        {
            if (! RunFileOpsLocalDirectoryRestoreChecksIdentityChecks(state))
            {
                Fail(L"The Local Copy walker restored a folder's times on an object it did not create, or did not report a refused restore; see "
                     L"trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::R4T7_BridgeCopyHoldsNoDirectoryHandle);
            return false;
        }
        case SelfTestState::Step::R4T7_BridgeCopyHoldsNoDirectoryHandle:
        case SelfTestState::Step::R4T8_BridgeDirectoryRestoreChecksIdentity:
        case SelfTestState::Step::R4T9_BridgeManyDirectoriesCopy:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                     = true;
            const auto resetInProgress     = wil::scope_exit([]() noexcept { inProgress = false; });
            const SelfTestState::Step step = state.step;
            bool passed                    = false;
            std::wstring_view failure;
            SelfTestState::Step next = SelfTestState::Step::R4A02_DeleteOverCopySourceWarnsAndQueues;
            if (step == SelfTestState::Step::R4T7_BridgeCopyHoldsNoDirectoryHandle)
            {
                passed  = RunFileOpsBridgeCopyHoldsNoDirectoryHandleChecks(state);
                failure = L"The host bridge kept a folder it created open while the folder's children copied; see trace.";
                next    = SelfTestState::Step::R4T8_BridgeDirectoryRestoreChecksIdentity;
            }
            else if (step == SelfTestState::Step::R4T8_BridgeDirectoryRestoreChecksIdentity)
            {
                passed  = RunFileOpsBridgeDirectoryRestoreChecksIdentityChecks(state);
                failure = L"The host bridge restored a folder's times on an object it did not create, wrote to a folder it could not bind "
                          L"again, or did not report a refused restore; see trace.";
                next    = SelfTestState::Step::R4T9_BridgeManyDirectoriesCopy;
            }
            else
            {
                passed  = RunFileOpsBridgeManyDirectoriesCopyChecks(state);
                failure = L"A bridge Copy of a 1,885-folder tree failed, was incomplete, or did not restore every folder's last-write time; see "
                          L"trace.";
            }
            if (! passed)
            {
                Fail(failure);
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, next);
            return false;
        }
        case SelfTestState::Step::FileOps_CrossProviderMtpNewFile:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsCrossProviderMtpNewFileChecks(state))
            {
                Fail(L"A new file copied to an MTP device did not publish through the device writer, or left a temporary object; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_VerifyAvailability);
            return false;
        }
        case SelfTestState::Step::FileOps_VerifyAvailability:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsVerifyAvailabilityChecks(state))
            {
                Fail(L"Requested verification removed a destination's only publication route, or a writer-digest destination lost Verify; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_SizelessSourceNote);
            return false;
        }
        case SelfTestState::Step::FileOps_SizelessSourceNote:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsSizelessSourceNoteChecks(state))
            {
                Fail(L"A source without file sizes did not fail per file with one Information note naming the cause; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_CrossRootLocalCopy);
            return false;
        }
        case SelfTestState::Step::FileOps_CrossRootLocalCopy:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            const auto result          = RunFileOpsCrossRootLocalCopyChecks(state);
            if (result.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(result.reason);
                return true;
            }
            RecordCurrentPhase(state, result.status, result.reason);
            NextStep(state, SelfTestState::Step::FileOps_ArchiveNestedCopyLocal);
            return false;
        }
        case SelfTestState::Step::FileOps_ArchiveNestedCopyLocal:
        case SelfTestState::Step::FileOps_ArchiveNestedCopyUnc:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            const bool network         = state.step == SelfTestState::Step::FileOps_ArchiveNestedCopyUnc;
            const auto result          = RunFileOpsArchiveNestedCopyChecks(state, network);
            if (result.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(result.reason);
                return true;
            }
            RecordCurrentPhase(state, result.status, result.reason);
            NextStep(state, network ? SelfTestState::Step::FileOps_CancelAllNeverStarted : SelfTestState::Step::FileOps_ArchiveNestedCopyUnc);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_FolderClosesBeforeLastFile:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeFolderClosesBeforeLastFileChecks(state))
            {
                Fail(L"A folder's discovery closed only after its last file on at least one walker; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_WideFolderClosesWithinQueueLead);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_WideFolderClosesWithinQueueLead:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeWideFolderClosesWithinQueueLeadChecks(state))
            {
                Fail(L"A folder wider than the queue lead closed only near its last file; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_NestedTreeMetadataRestoredAfterChildren);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_NestedTreeMetadataRestoredAfterChildren:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeNestedTreeMetadataChecks(state))
            {
                Fail(L"A created directory lost its source timestamps (metadata applied before a child finished); see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_ManagedMoveListingCompleteBeforeRemoval);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_ManagedMoveListingCompleteBeforeRemoval:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            std::wstring unavailableReason;
            const bool passed = RunDiscoveryScopeManagedMoveListingChecks(state, unavailableReason);
            if (! passed && ! unavailableReason.empty())
            {
                Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryClosure.ManagedMoveWitness", std::format(L"skip={}", unavailableReason), 0u, 0u, 0u, S_FALSE);
                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, unavailableReason);
                NextStep(state, SelfTestState::Step::DiscoveryScope_WorkerFailureIsNotCanceled);
                return false;
            }
            if (! passed)
            {
                Fail(L"A Managed Move skipped a child or kept its source folder; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_WorkerFailureIsNotCanceled);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_WorkerFailureIsNotCanceled:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeWorkerFailureIsNotCanceledChecks(state))
            {
                Fail(L"A worker's failure ended the recursive copy as canceled or was masked; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_FlatFolderUsesFullConcurrency);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_FlatFolderUsesFullConcurrency:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeFlatFolderUsesFullConcurrencyChecks(state))
            {
                Fail(L"A flat folder's copy never reached the configured concurrency after its listing; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_ListingSilenceNoted);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_ListingSilenceNoted:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeListingSilenceNotedChecks(state))
            {
                Fail(L"A listing silence during a running copy was not noted exactly once while discovery was open; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_TenThousandSmallFilesCloseAfterListing);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_TenThousandSmallFilesCloseAfterListing:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeTenThousandSmallFilesChecks(state))
            {
                Fail(L"A 10,000-file folder did not close once it was listed; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_RetryDoesNotGrowAfterClose);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_RetryDoesNotGrowAfterClose:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeRetryDoesNotGrowAfterCloseChecks(state))
            {
                Fail(L"A retried directory listing counted an entry twice or missed a subtree; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_HeldFolderFailureStillCloses);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_HeldFolderFailureStillCloses:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeHeldFolderFailureStillClosesChecks(state))
            {
                Fail(L"A skipped held folder kept the discovery scope open; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_FolderLiveProgress);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_FolderLiveProgress:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeFolderLiveProgressChecks(state))
            {
                Fail(L"A real folder copy lost its child totals or individual-file progress; see recorded facts and captures.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::DiscoveryScope_MultipleFoldersLiveProgress);
            return false;
        }
        case SelfTestState::Step::DiscoveryScope_MultipleFoldersLiveProgress:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDiscoveryScopeMultipleFoldersLiveProgressChecks(state))
            {
                Fail(L"A real multiple-folder copy published inconsistent live progress or contents; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_SelectionShapeLocal);
            return false;
        }
        case SelfTestState::Step::FileOps_SelectionShapeLocal:
        case SelfTestState::Step::FileOps_SelectionShapeSmb:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress       = true;
            const auto reset = wil::scope_exit([]() noexcept { inProgress = false; });
            const bool smb   = state.step == SelfTestState::Step::FileOps_SelectionShapeSmb;
            std::wstring unavailableReason;
            const bool passed = RunFileOpsSelectionShapeChecks(state, smb, unavailableReason);
            if (! unavailableReason.empty())
            {
                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, unavailableReason);
            }
            else if (! passed)
            {
                Fail(L"Copy/Move selection shapes did not use the same file concurrency or converge counting early; see SelectionShape metrics.");
                return true;
            }
            else
            {
                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            }
            NextStep(state, smb ? SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns : SelfTestState::Step::FileOps_SelectionShapeSmb);
            return false;
        }
        case SelfTestState::Step::FileOps_CancelAllNeverStarted:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunFileOpsCancelAllNeverStartedChecks(state))
            {
                Fail(L"Cancel all did not end a queued task as a clean never-started cancel, or the completed group did not collapse exactly "
                     L"once; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_SmbCopyThroughputVariance);
            return false;
        }
        case SelfTestState::Step::FileOps_SmbCopyThroughputVariance:
        case SelfTestState::Step::FileOps_LocalCopyThroughput:
        case SelfTestState::Step::FileOps_SmbNativeRouteConcurrencyBudget:
        case SelfTestState::Step::FileOps_SameProviderBridgeDestinationStorage:
        case SelfTestState::Step::FileOps_SmbManagedMoveLeavesNoStage:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            ThroughputCaseResult result{};
            std::wstring_view failure;
            SelfTestState::Step next = SelfTestState::Step::FileOps_RateWindowDisplay;
            switch (state.step)
            {
                case SelfTestState::Step::FileOps_SmbCopyThroughputVariance:
                    result  = RunFileOpsSmbCopyThroughputVarianceChecks(state);
                    failure = L"SMB throughput could not measure the corpus copy";
                    next    = SelfTestState::Step::FileOps_LocalCopyThroughput;
                    break;
                case SelfTestState::Step::FileOps_LocalCopyThroughput:
                    result  = RunFileOpsLocalCopyThroughputChecks(state);
                    failure = L"Local throughput could not measure the corpus copy";
                    next    = SelfTestState::Step::FileOps_SmbNativeRouteConcurrencyBudget;
                    break;
                case SelfTestState::Step::FileOps_SmbNativeRouteConcurrencyBudget:
                    result  = RunFileOpsSmbNativeRouteConcurrencyBudgetChecks(state);
                    failure = L"The native route did not budget concurrency as min(source, destination)";
                    next    = SelfTestState::Step::FileOps_SameProviderBridgeDestinationStorage;
                    break;
                case SelfTestState::Step::FileOps_SameProviderBridgeDestinationStorage:
                    result  = RunFileOpsSameProviderBridgeDestinationStorageChecks(state);
                    failure = L"A same-provider bridge route did not consult its destination's storage";
                    next    = SelfTestState::Step::FileOps_SmbManagedMoveLeavesNoStage;
                    break;
                default:
                    result  = RunFileOpsSmbManagedMoveLeavesNoStageChecks(state);
                    failure = L"A Move to the share cancelled in one of its phases left a stage or a partial file";
                    break;
            }
            if (! result.passed && ! result.skipped)
            {
                Fail(std::format(L"{}: {}.", failure, result.reason));
                return true;
            }
            RecordCurrentPhase(
                state, result.skipped ? SelfTest::SelfTestCaseResult::Status::skipped : SelfTest::SelfTestCaseResult::Status::passed, result.reason);
            NextStep(state, next);
            return false;
        }
        case SelfTestState::Step::FileOps_RateWindowDisplay:
        {
            const ThroughputCaseResult result = RunFileOpsRateWindowDisplayChecks();
            if (! result.passed)
            {
                Fail(std::format(L"The speed display broke an R10.2 rule: {}.", result.reason));
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_ProgressPhaseAccounting);
            return false;
        }
        case SelfTestState::Step::FileOps_ProgressPhaseAccounting:
        {
            const ThroughputCaseResult result = RunFileOpsProgressPhaseAccountingChecks(state);
            if (! result.passed)
            {
                Fail(std::format(L"The host's progress accounting broke an I35 S6 rule: {}.", result.reason));
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_StallLabelLayout);
            return false;
        }
        case SelfTestState::Step::FileOps_StallLabelLayout:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                        = true;
            const auto resetInProgress        = wil::scope_exit([]() noexcept { inProgress = false; });
            const ThroughputCaseResult result = RunFileOpsStallLabelLayoutChecks(state);
            if (! result.passed)
            {
                Fail(std::format(L"A stall name moved the throughput graph: {}.", result.reason));
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_SmallFileCopyCost);
            return false;
        }
        case SelfTestState::Step::FileOps_SmallFileCopyCost:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            std::wstring failure;
            if (! RunFileOpsSmallFileCopyCostChecks(state, failure))
            {
                Fail(std::format(L"The small-file Copy cost could not be measured: {}.", failure));
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessLocationByName);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessLocationByName:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessLocationByNameChecks(state))
            {
                Fail(L"A location without file identities did not delete by name behind one consent, deleted a changed item, asked the touch guard "
                     L"for a leftover it hands to that consent, or did not escalate a refused Recycle by name; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessSharingAndInterlock);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessSharingAndInterlock:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessSharingAndInterlockChecks(state))
            {
                Fail(L"A delete by name fell back on a sharing failure, was offered for an item that could not be opened, or ran on a path "
                     L"another task holds; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessQueuedRootOverlap);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessQueuedRootOverlap:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            std::wstring skipped;
            if (! RunDeleteIdentityLessQueuedRootOverlapChecks(state, skipped))
            {
                Fail(L"A task admitted after a queued delete by name was not advised about its root, ran inside it before it, or lost data to it; "
                     L"see trace.");
                return true;
            }
            RecordCurrentPhase(state, skipped.empty() ? SelfTest::SelfTestCaseResult::Status::passed : SelfTest::SelfTestCaseResult::Status::skipped, skipped);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeOtherSpelling);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeOtherSpelling:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            std::wstring skipped;
            if (! RunDeleteIdentityLessPinnedTreeOtherSpellingChecks(state, skipped))
            {
                Fail(L"A delete by name removed data moved into its tree through another spelling of the share; see trace.");
                return true;
            }
            RecordCurrentPhase(state, skipped.empty() ? SelfTest::SelfTestCaseResult::Status::passed : SelfTest::SelfTestCaseResult::Status::skipped, skipped);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeArchiveWriter);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeArchiveWriter:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            std::wstring skipped;
            if (! RunDeleteIdentityLessPinnedTreeArchiveWriterChecks(state, skipped))
            {
                Fail(L"A delete by name removed an archive Pack wrote into its tree while it waited; see trace.");
                return true;
            }
            RecordCurrentPhase(state, skipped.empty() ? SelfTest::SelfTestCaseResult::Status::passed : SelfTest::SelfTestCaseResult::Status::skipped, skipped);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeCardPending);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeCardPending:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeCardPendingChecks(state))
            {
                Fail(L"A delete by name was invisible to later tasks while its card waited, or deleted a tree that changed deep inside; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeConcurrentWriter);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeConcurrentWriter:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeConcurrentWriterChecks(state))
            {
                Fail(L"A delete by name removed a file written into its tree after its listing, or did not report the partial delete; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeBudget);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeBudget:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeBudgetChecks(state))
            {
                Fail(L"A tree over the by-name budget was offered a delete by name or refused without its reason; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeChangedEntry);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeChangedEntry:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeChangedEntryChecks(state))
            {
                Fail(L"A listed entry that changed after the listing did not stop the delete by name with nothing deleted; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeEscalationPartial);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeEscalationPartial:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeEscalationPartialChecks(state))
            {
                Fail(L"A Recycle escalated to a delete by name did not report its partial delete; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeLinkSwap);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeLinkSwap:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeLinkSwapChecks(state))
            {
                Fail(L"A junction swapped into the tree after the listing redirected a delete by name; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeSubfolderOpen);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeSubfolderOpen:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeSubfolderOpenChecks(state))
            {
                Fail(L"A sub-folder refused with ERROR_NOT_SUPPORTED was not reported as an open failure; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeRecycleOverBudget);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeRecycleOverBudget:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeRecycleOverBudgetChecks(state))
            {
                Fail(L"A refused Recycle over the by-name bounds did not keep the reason of its pin; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeHolders);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeHolders:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeHolderChecks(state))
            {
                Fail(L"A task that holds the folder and will not wait did not refuse a delete by name; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessPinnedTreeScale);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessPinnedTreeScale:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessPinnedTreeScaleChecks(state))
            {
                Fail(L"A delete by name of a tree of a few thousand entries did not delete it; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessFolderChangeAndHolders);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessFolderChangeAndHolders:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessFolderChangeAndHolderChecks(state))
            {
                Fail(L"A folder deleted by name was deleted although a child changed after its pin, or although another task took it after the "
                     L"answer; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Delete_IdentityLessOpenFailureAndArchiveOrigins);
            return false;
        }
        case SelfTestState::Step::Delete_IdentityLessOpenFailureAndArchiveOrigins:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            if (! RunDeleteIdentityLessOpenFailureAndArchiveOriginChecks(state))
            {
                Fail(L"An open that failed with ERROR_NOT_SUPPORTED was treated as an identity-less object, or a Pack or Unpack cleanup deleted by "
                     L"name; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_IdentityLessGuardPartition);
            return false;
        }
        case SelfTestState::Step::FileOps_IdentityLessGuardPartition:
        {
            if (! RunFileOpsIdentityLessGuardPartitionChecks(state))
            {
                Fail(L"The touch guard kept an unidentifiable Delete item as a receipt item, accepted another touch of it, or misclassified a "
                     L"provider staging name; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_LeftoverCleanupWalkBounds);
            return false;
        }
        case SelfTestState::Step::FileOps_LeftoverCleanupWalkBounds:
        {
            if (! RunFileOpsLeftoverCleanupWalkBoundsChecks(state))
            {
                Fail(L"The leftover listing's walk listed the wrong rows, followed a link, or did not report a bound or a stop; see trace.");
                return true;
            }
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::FileOps_OwnedStageCleanupBudget);
            return false;
        }
        case SelfTestState::Step::FileOps_OwnedStageCleanupBudget:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            std::wstring skipped;
            if (! RunFileOpsOwnedStageCleanupBudgetChecks(state, skipped))
            {
                Fail(L"The cancel watch cut short an owned-stage cleanup within its budget: the bridge stage was retained or reported; see trace.");
                return true;
            }
            RecordCurrentPhase(state, skipped.empty() ? SelfTest::SelfTestCaseResult::Status::passed : SelfTest::SelfTestCaseResult::Status::skipped, skipped);
            NextStep(state, SelfTestState::Step::Delete_NasFileIdInfoUnsupportedByName);
            return false;
        }
        case SelfTestState::Step::Delete_NasFileIdInfoUnsupportedByName:
        {
            static bool inProgress = false;
            if (inProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            inProgress                 = true;
            const auto resetInProgress = wil::scope_exit([]() noexcept { inProgress = false; });
            std::wstring skipped;
            if (! RunDeleteNasFileIdInfoUnsupportedByNameChecks(state, skipped))
            {
                Fail(L"The registered share without FileIdInfo did not bind as Unsupported, corroborate with its file index, or delete by name; "
                     L"see trace.");
                return true;
            }
            RecordCurrentPhase(state, skipped.empty() ? SelfTest::SelfTestCaseResult::Status::passed : SelfTest::SelfTestCaseResult::Status::skipped, skipped);
            NextStep(state, SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder);
            return false;
        }
        case SelfTestState::Step::FileOps_CommonPresentationPerf:
        {
            static bool measurementInProgress = false;
            if (measurementInProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            measurementInProgress = true;
            auto resetMeasurement = wil::scope_exit([] { measurementInProgress = false; });
            const bool measured   = RunFileOpsCommonPresentationPerf(state);
            RecordCurrentPhase(state,
                               measured ? SelfTest::SelfTestCaseResult::Status::passed : SelfTest::SelfTestCaseResult::Status::failed,
                               measured ? L"" : L"Common presentation measurement failed; see trace.");
            state.failed.store(! measured, std::memory_order_release);
            state.step = SelfTestState::Step::Done;
            state.running.store(false, std::memory_order_release);
            state.done.store(true, std::memory_order_release);
            return true;
        }
        case SelfTestState::Step::FileOps_IntegratedPresentationPerf:
        {
            static bool measurementInProgress = false;
            if (measurementInProgress || ! TryGetFolderWindow(state.mainWindow))
                return false;
            measurementInProgress = true;
            auto resetMeasurement = wil::scope_exit([] { measurementInProgress = false; });
            const bool measured   = RunFileOpsIntegratedPresentationPerf(state);
            RecordCurrentPhase(state,
                               measured ? SelfTest::SelfTestCaseResult::Status::passed : SelfTest::SelfTestCaseResult::Status::failed,
                               measured ? L"" : L"Integrated presentation measurement failed; see trace.");
            state.failed.store(! measured, std::memory_order_release);
            state.step = SelfTestState::Step::Done;
            state.running.store(false, std::memory_order_release);
            state.done.store(true, std::memory_order_release);
            return true;
        }
        case SelfTestState::Step::FileOps_VisualGallery:
        case SelfTestState::Step::FileOps_VisualGalleryInteraction:
        {
            static bool captureInProgress = false;
            if (captureInProgress)
                return false;
            if (! TryGetFolderWindow(state.mainWindow))
                return false;
            captureInProgress   = true;
            auto resetCapture   = wil::scope_exit([] { captureInProgress = false; });
            const bool captured = CaptureFileOpsVisualGallery(state);
            RecordCurrentPhase(state,
                               captured ? SelfTest::SelfTestCaseResult::Status::passed : SelfTest::SelfTestCaseResult::Status::failed,
                               captured ? L"" : L"Scenario capture failed; see trace.");
            state.failed.store(! captured, std::memory_order_release);
            state.step = SelfTestState::Step::Done;
            state.running.store(false, std::memory_order_release);
            state.done.store(true, std::memory_order_release);
            return true;
        }
        case SelfTestState::Step::Setup:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 30'000ull))
            {
                const HWND folderWindowHwnd = state.mainWindow ? FindWindowExW(state.mainWindow, nullptr, kFolderWindowClassName.data(), nullptr) : nullptr;
                const HWND folderViewA      = folderWindowHwnd ? FindWindowExW(folderWindowHwnd, nullptr, kFolderViewClassName.data(), nullptr) : nullptr;
                const HWND folderViewB      = folderViewA ? FindWindowExW(folderWindowHwnd, folderViewA, kFolderViewClassName.data(), nullptr) : nullptr;

                const FolderView* viewA = folderViewA ? reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderViewA, GWLP_USERDATA)) : nullptr;
                const FolderView* viewB = folderViewB ? reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderViewB, GWLP_USERDATA)) : nullptr;

                const bool cbA = viewA ? viewA->DebugHasFileOperationRequestCallback() : false;
                const bool cbB = viewB ? viewB->DebugHasFileOperationRequestCallback() : false;

                Fail(std::format(L"Setup timed out (folderWindow={} folderViewA={} folderViewB={} callbackA={} callbackB={}).",
                                 folderWindowHwnd != nullptr,
                                 folderViewA != nullptr,
                                 folderViewB != nullptr,
                                 cbA,
                                 cbB));
                return true;
            }

            state.folderWindow = TryGetFolderWindow(state.mainWindow);
            if (! state.folderWindow)
            {
                return false;
            }

            state.fileOps = TryGetFileOps(state.folderWindow);
            if (! state.fileOps)
            {
                return false;
            }
            AppendLog(L"Setup: resolved file-operations state");
            state.autoDismissSuccessOriginal               = state.fileOps->GetAutoDismissSuccess();
            state.fileOperationsOriginal                   = g_settings.fileOperations;
            state.fileOperationsBackedUp                   = true;
            state.appThemeOriginal                         = state.folderWindow->GetTheme();
            state.appThemeBackedUp                         = true;
            AppTheme deterministicMotionTheme              = state.appThemeOriginal.value();
            deterministicMotionTheme.reducedMotionOverride = false;
            state.folderWindow->ApplyTheme(deterministicMotionTheme);
            Common::Settings::FileOperationsSettings& fileOperations = EnsureFileOperationsSettingsForSelfTest();
            fileOperations.crossFsBridgeBufferSizeKB                 = 4096u;
            fileOperations.defaultBandwidthLimitBytesPerSecond       = 0;

            if (! LoadPlugins(state))
            {
                return false;
            }
            AppendLog(L"Setup: plugins loaded");

            {
                const HWND folderWindowHwnd = FindWindowExW(state.mainWindow, nullptr, kFolderWindowClassName.data(), nullptr);
                if (! folderWindowHwnd)
                {
                    return false;
                }

                const HWND folderViewA = FindWindowExW(folderWindowHwnd, nullptr, kFolderViewClassName.data(), nullptr);
                if (! folderViewA)
                {
                    return false;
                }

                const HWND folderViewB = FindWindowExW(folderWindowHwnd, folderViewA, kFolderViewClassName.data(), nullptr);
                if (! folderViewB)
                {
                    return false;
                }

                const FolderView* viewA = reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderViewA, GWLP_USERDATA));
                const FolderView* viewB = reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderViewB, GWLP_USERDATA));
                if (! viewA || ! viewB)
                {
                    return false;
                }

                if (! viewA->DebugHasFileOperationRequestCallback() || ! viewB->DebugHasFileOperationRequestCallback())
                {
                    return false;
                }
            }
            AppendLog(L"Setup: folder views expose file-operation callbacks");

            {
                const HRESULT leftHr  = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Left, kPluginIdLocal);
                const HRESULT rightHr = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Right, kPluginIdLocal);
                if (FAILED(leftHr) || FAILED(rightHr))
                {
                    Fail(std::format(L"Setup failed to set panes to local filesystem plugin (left=0x{:08X} right=0x{:08X}).",
                                     static_cast<unsigned long>(leftHr),
                                     static_cast<unsigned long>(rightHr)));
                    return true;
                }
            }
            AppendLog(L"Setup: panes switched to local filesystem plugin");

            if (state.localConfigOriginal.empty())
            {
                static_cast<void>(BackupPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal));
            }
            if (state.dummyConfigOriginal.empty())
            {
                static_cast<void>(BackupPluginConfiguration(state.infoDummy.get(), state.dummyConfigOriginal));
            }
            if (state.config7zOriginal.empty() && state.info7z)
            {
                static_cast<void>(BackupPluginConfiguration(state.info7z.get(), state.config7zOriginal));
            }
            if (! state.connectionsBackedUp)
            {
                state.connectionsOriginal = g_settings.connections;
                state.connectionsBackedUp = true;
            }
            AppendLog(L"Setup: original plugin/config state available");

            const auto buildDummySeedConfig = [](unsigned int seed) noexcept
            { return std::format(R"json({{"maxChildrenPerDirectory":128,"maxDepth":10,"seed":{},"latencyMs":5,"virtualSpeedLimit":"0"}})json", seed); };

            const std::array<std::wstring_view, 9> dummyDestinationFolders{L"/dest-a",
                                                                           L"/dest-b",
                                                                           L"/dest-skip-a",
                                                                           L"/dest-skip-b",
                                                                           L"/dest-queued-a",
                                                                           L"/dest-queued-b",
                                                                           L"/dest-queued-c",
                                                                           L"/dest-wait-a",
                                                                           L"/dest-wait-b"};
            const auto ensureDummyDestinationFolders = [&]() noexcept
            {
                for (const auto& folder : dummyDestinationFolders)
                {
                    if (! EnsureDummyFolderExists(state.fsDummy.get(), folder))
                    {
                        Fail(std::format(L"Failed to create dummy destination folder: {}", folder));
                        return false;
                    }
                }
                return true;
            };

            if (state.dummyPaths.empty() || state.dummySeedConfig.empty())
            {
                AppendLog(L"Setup: selecting dummy source path");
                const auto toDummyPath = [](std::wstring_view leaf) -> std::wstring
                {
                    if (leaf.empty())
                    {
                        return L"/";
                    }

                    const wchar_t first = leaf.front();
                    if (first == L'/' || first == L'\\')
                    {
                        return std::wstring(leaf);
                    }

                    return std::format(L"/{}", leaf);
                };

                const auto trySeed = [&](unsigned int seed) noexcept -> bool
                {
                    AppendLog(std::format(L"Setup: applying dummy seed {}", seed));
                    const std::string config = buildDummySeedConfig(seed);
                    if (! SetPluginConfiguration(state.infoDummy.get(), config))
                    {
                        AppendLog(std::format(L"Setup: failed to apply dummy seed {}", seed));
                        return false;
                    }

                    const std::vector<std::wstring> dirs = ListDirectories(state.fsDummy.get(), L"/", 64);
                    AppendLog(std::format(L"Setup: dummy seed {} listed {} candidate directories", seed, dirs.size()));
                    std::wstring bestCandidate;
                    size_t bestChildren = 0;
                    uint64_t bestBytes  = 0;

                    std::wstring firstNonEmpty;
                    size_t firstNonEmptyChildren = 0;

                    for (const auto& dir : dirs)
                    {
                        const std::wstring candidate = toDummyPath(dir);
                        if (candidate == L"/")
                        {
                            continue;
                        }

                        const size_t childCount = GetDirectoryEntryCount(state.fsDummy.get(), candidate);
                        if (childCount == 0u)
                        {
                            continue;
                        }

                        if (firstNonEmpty.empty())
                        {
                            firstNonEmpty         = candidate;
                            firstNonEmptyChildren = childCount;
                        }

                        const uint64_t bytes = GetDirectoryImmediateFileBytes(state.fsDummy.get(), candidate);
                        if (bytes > bestBytes)
                        {
                            bestCandidate = candidate;
                            bestChildren  = childCount;
                            bestBytes     = bytes;
                        }
                    }

                    if (bestCandidate.empty() && ! firstNonEmpty.empty())
                    {
                        bestCandidate = firstNonEmpty;
                        bestChildren  = firstNonEmptyChildren;
                    }

                    if (! bestCandidate.empty())
                    {
                        state.dummyPaths.push_back(bestCandidate);
                        state.dummyPaths.push_back(bestCandidate);
                        state.dummySeedConfig = config;
                        AppendLog(std::format(L"Dummy selection seed={} path={} children={} bytes={}", seed, bestCandidate, bestChildren, bestBytes));
                        return true;
                    }

                    return false;
                };

                const std::array<unsigned int, 4> seeds{42u, 1337u, 2026u, 7u};
                for (const unsigned int seed : seeds)
                {
                    if (trySeed(seed))
                    {
                        break;
                    }
                }

                if (state.dummyPaths.empty())
                {
                    Fail(L"FileSystemDummy did not provide a non-empty directory for discovery tests.");
                    return true;
                }
                AppendLog(L"Setup: dummy source path ready");
            }
            else
            {
                AppendLog(L"Setup: reapplying cached dummy seed config");
                if (! SetPluginConfiguration(state.infoDummy.get(), state.dummySeedConfig))
                {
                    Fail(L"Failed to reapply cached dummy seed config.");
                    return true;
                }
                AppendLog(L"Setup: cached dummy seed config reapplied");
            }

            // FileSystemDummy's batch operations require these folders to already exist. Provider
            // config changes reset the dummy tree, so recreate them after every seed reapply.
            if (! ensureDummyDestinationFolders())
            {
                return true;
            }

            if (state.tempRoot.empty())
            {
                AppendLog(L"Setup: recreating temp root");
                state.tempRoot = GetTempRootPath();
                if (! RecreateEmptyDirectory(state.tempRoot))
                {
                    Fail(L"Failed to create temp root directory for self-test.");
                    return true;
                }
                AppendLog(std::format(L"Setup: temp root created at {}", state.tempRoot.wstring()));

                const std::filesystem::path src      = state.tempRoot / L"copy-src";
                const std::filesystem::path dst      = state.tempRoot / L"copy-dst";
                const std::filesystem::path del      = state.tempRoot / L"delete-tree";
                const std::filesystem::path en       = state.tempRoot / L"enum";
                const std::filesystem::path watch    = state.tempRoot / L"watch";
                const std::filesystem::path preA     = state.tempRoot / L"discovery-a";
                const std::filesystem::path preB     = state.tempRoot / L"discovery-b";
                const bool needsDeleteTree           = IsPhaseSelected(state, SelfTestState::Step::Phase6_DeleteBytesMeaningful);
                const bool needsDiscoverySwitchTrees = IsPhaseSelected(state, SelfTestState::Step::Phase5_DiscoverySkipContinues) ||
                                                       IsPhaseSelected(state, SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery) ||
                                                       IsPhaseSelected(state, SelfTestState::Step::Phase5_SwitchWaitToParallelResume);

                std::error_code ec;
                if (! SeedCopyKnobSourceFiles(src))
                {
                    Fail(L"Failed to seed copy-src directory.");
                    return true;
                }
                AppendLog(L"Setup: copy-src seeded");

                ec.clear();
                std::filesystem::create_directories(dst, ec);
                if (ec)
                {
                    Fail(L"Failed to create copy-dst directory.");
                    return true;
                }
                AppendLog(L"Setup: copy-dst created");

                if (needsDeleteTree)
                {
                    std::filesystem::create_directories(del, ec);
                    if (ec)
                    {
                        Fail(L"Failed to create delete-tree directory.");
                        return true;
                    }
                    AppendLog(L"Setup: delete-tree root created");
                }
                else
                {
                    AppendLog(L"Setup: delete-tree skipped");
                }

                std::filesystem::create_directories(en, ec);
                if (ec)
                {
                    Fail(L"Failed to create enum directory.");
                    return true;
                }
                AppendLog(L"Setup: enum directory created");

                std::filesystem::create_directories(watch, ec);
                if (ec)
                {
                    Fail(L"Failed to create watch directory.");
                    return true;
                }
                AppendLog(L"Setup: watch directory created");

                // Keep this tree large enough that delete progress callbacks occur beyond the initial throttle window,
                // so delete completedBytes > 0 is observable while the task is running.
                if (needsDeleteTree)
                {
                    if (! CreateDeleteTree(del, 10, 300, 1))
                    {
                        Fail(L"Failed to create delete-tree.");
                        return true;
                    }
                    AppendLog(L"Setup: delete-tree seeded");
                }

                if (needsDiscoverySwitchTrees)
                {
                    if (! CreateDeleteTreeWithRetry(preA, 10, 200, 1, L"discovery tree A") ||
                        ! CreateDeleteTreeWithRetry(preB, 10, 200, 1, L"discovery tree B"))
                    {
                        Fail(L"Failed to create discovery trees.");
                        return true;
                    }
                    AppendLog(L"Setup: discovery trees seeded");
                }
                else
                {
                    AppendLog(L"Setup: discovery trees skipped");
                }

                AppendLog(L"Setup: temp root ready");
            }

            AppendLog(L"Setup: complete");
            NextStep(state, SelfTestState::Step::FileOps_QueuedAdmissionContracts);
            return false;
        }

        case SelfTestState::Step::FileOps_ParallelGraphFairColorWeight:
        {
            FileOperationsPopupInternal::GraphHueWeightDebugSnapshot graph{};
            if (! DebugBuildFileOperationsGraphFairColorWeightSnapshot(graph))
            {
                Fail(L"FileOps_ParallelGraphFairColorWeight could not build a graph hue-weight snapshot.");
                return true;
            }

            Debug::Perf::Emit(L"FileOps.Popup.Graph.HueWeightCount",
                              L"scenario=four-equal-streams",
                              static_cast<uint64_t>(graph.hueCount),
                              static_cast<uint64_t>(graph.totalWeight),
                              4u,
                              S_OK);

            if (graph.hueCount != 4u)
            {
                Fail(std::format(L"FileOps_ParallelGraphFairColorWeight expected 4 hue buckets for 4 equal streams, got {}.", graph.hueCount));
                return true;
            }

            if (graph.totalWeight <= 0.0)
            {
                Fail(L"FileOps_ParallelGraphFairColorWeight produced no graph weight.");
                return true;
            }

            double minShare = 1.0;
            double maxShare = 0.0;
            for (size_t i = 0; i < graph.hueCount; ++i)
            {
                const double share = graph.weights[i] / graph.totalWeight;
                minShare           = std::min(minShare, share);
                maxShare           = std::max(maxShare, share);
            }

            Debug::Perf::Emit(L"FileOps.Popup.Graph.HueWeightSpreadPermille",
                              L"scenario=four-equal-streams",
                              static_cast<uint64_t>((maxShare - minShare) * 1000.0),
                              static_cast<uint64_t>(minShare * 1000.0),
                              static_cast<uint64_t>(maxShare * 1000.0),
                              S_OK);

            if (minShare < 0.20 || maxShare > 0.30)
            {
                Fail(std::format(L"FileOps_ParallelGraphFairColorWeight expected each stream share near 25%, got min={:.3f}, max={:.3f}.", minShare, maxShare));
                return true;
            }

            NextStep(state, SelfTestState::Step::Phase5_DiscoverySingleTraversal);
            return false;
        }

        default: break;
    }

    return false;
}
#pragma warning(pop)

void FileOperationsSelfTest::NotifyTaskCompleted(std::uint64_t taskId, HRESULT hr) noexcept
{
    SelfTestState& state = GetState();
    if (! state.running.load(std::memory_order_acquire))
    {
        return;
    }

    CompletedTaskInfo info{};
    info.hr             = hr;
    info.completionTick = GetTickCount64();
    if (state.fileOps)
    {
        if (auto* task = state.fileOps->FindTask(taskId))
        {
            info.discoveryClosed                             = task->_discoveryClosed.load(std::memory_order_acquire);
            info.firstMutationBeforeDiscoveryClosed          = task->_firstMutationBeforeDiscoveryClosed.load(std::memory_order_acquire);
            info.discoverySkipped                            = task->_discoverySkipped.load(std::memory_order_acquire);
            info.discoveredTotalBytes                        = task->_discoveredTotalBytes.load(std::memory_order_acquire);
            info.discoveredFiles                             = task->_discoveredFileCount.load(std::memory_order_acquire);
            info.copiedFileCount                             = task->_copiedFileCount.load(std::memory_order_acquire);
            info.copiedFileCountKnown                        = task->_copiedFileCountKnown.load(std::memory_order_acquire);
            info.discoveredBytesKnown                        = task->_discoveredBytesKnown.load(std::memory_order_acquire);
            info.discoveredDirectories                       = task->_discoveredDirectoryCount.load(std::memory_order_acquire);
            info.discoveryDurationUs                         = task->_perf.discoveryOpenUs.load(std::memory_order_acquire);
            info.discoveryMaxQueueDepth                      = task->_discoveryMaxQueueDepth.load(std::memory_order_acquire);
            info.discoveryStarvationCount                    = task->_discoveryStarvationCount.load(std::memory_order_acquire);
            info.discoveryCallbackCount                      = task->_perf.discoveryCallbackCount.load(std::memory_order_acquire);
            info.discoveryCallbackUs                         = task->_perf.discoveryCallbackUs.load(std::memory_order_acquire);
            info.discoveryLockWaitUs                         = task->_perf.discoveryLockWaitUs.load(std::memory_order_acquire);
            info.discoveryFirstMutationUs                    = task->_discoveryFirstMutationUs.load(std::memory_order_acquire);
            info.discoveryBytesCompletedWhileOpen            = task->_discoveryCompletedBytesWhileOpen.load(std::memory_order_acquire);
            info.discoveryMutationsCompletedWhileOpen        = task->_discoveryCompletedMutationsWhileOpen.load(std::memory_order_acquire);
            info.discoveryGrowthAfterCloseCount              = task->_discoveryGrowthAfterCloseCount.load(std::memory_order_acquire);
            info.progressCallbackCount                       = task->_progressCallbackCount.load(std::memory_order_acquire);
            info.started                                     = task->HasStarted();
            info.conflictWaitUs                              = task->_perf.conflictWaitUs;
            info.conflictConvergenceWaitUs                   = task->_perf.conflictConvergenceWaitUs;
            info.conflictPromptCount                         = task->_perf.conflictPromptCount;
            info.interlockWaitUs                             = task->_perf.interlockWaitUs;
            info.interlockWaitCount                          = task->_perf.interlockWaitCount;
            info.bridgeDirectoryEnsureCount                  = task->_bridgeDirectoryEnsureCount.load(std::memory_order_acquire);
            info.bridgeSourceDirectoryEnumerationCount       = task->_bridgeSourceDirectoryEnumerationCount.load(std::memory_order_acquire);
            info.bridgeFileAdmissionCount                    = task->_bridgeFileAdmissionCount.load(std::memory_order_acquire);
            info.bridgeMaxActiveFileTransfers                = task->_bridgeMaxActiveFileTransfers.load(std::memory_order_acquire);
            info.bridgeEarlyFileStartCount                   = task->_bridgeFileStartedBeforeProducerDone.load(std::memory_order_acquire);
            info.bridgeAdmissionMaxQueueDepth                = task->_bridgeAdmissionMaxQueueDepth.load(std::memory_order_acquire);
            info.bridgeTraversalMaxDepth                     = task->_bridgeTraversalMaxDepth.load(std::memory_order_acquire);
            info.bridgeTraversalMaxRetainedEntries           = task->_bridgeTraversalMaxRetainedEntries.load(std::memory_order_acquire);
            info.bridgeTraversalMaxQueuedPathBytes           = task->_bridgeTraversalMaxQueuedPathBytes.load(std::memory_order_acquire);
            info.bridgeTraversalMaxMetadataBytes             = task->_bridgeTraversalMaxMetadataBytes.load(std::memory_order_acquire);
            info.bridgeTraversalLimitHitCount                = task->_bridgeTraversalLimitHitCount.load(std::memory_order_acquire);
            info.conflictExpectedDestinationBoundCount       = task->_conflictExpectedDestinationBoundCount.load(std::memory_order_acquire);
            info.conflictExpectedDestinationUnavailableCount = task->_conflictExpectedDestinationUnavailableCount.load(std::memory_order_acquire);
            info.conflictExpectedDestinationReturnedCount    = task->_conflictExpectedDestinationReturnedCount.load(std::memory_order_acquire);
            info.bridgeImmediateDestinationSizeProbeUs       = task->_perf.bridgeImmediateDestinationSizeProbeUs.load(std::memory_order_acquire);
            info.bridgeImmediateDestinationSizeProbeCount    = task->_perf.bridgeImmediateDestinationSizeProbeCount.load(std::memory_order_acquire);
            info.bridgeCommitSizeProofCount                  = task->_perf.bridgeCommitSizeProofCount.load(std::memory_order_acquire);
            info.bridgeCommitSizeProofFallbackCount          = task->_perf.bridgeCommitSizeProofFallbackCount.load(std::memory_order_acquire);
            info.verificationUs                              = task->_perf.verificationUs.load(std::memory_order_acquire);
            info.verificationReadBytes                       = task->_perf.verificationReadBytes.load(std::memory_order_acquire);
            info.verificationReadCalls                       = task->_perf.verificationReadCalls.load(std::memory_order_acquire);
            info.verificationProviderProofCount              = task->_perf.verificationProviderProofCount.load(std::memory_order_acquire);
            info.verificationHostReadbackCount               = task->_perf.verificationHostReadbackCount.load(std::memory_order_acquire);
            info.verificationTotalBytes                      = task->_verificationTotalBytes.load(std::memory_order_acquire);
            info.verificationCompletedBytes                  = task->_verificationCompletedBytes.load(std::memory_order_acquire);
            info.configuredMaxConcurrency                    = task->_effectiveConcurrencyBudget.load(std::memory_order_acquire);
            if (info.configuredMaxConcurrency == 0u)
            {
                info.configuredMaxConcurrency = task->_perItemMaxConcurrencyBudget;
            }
            info.progressCadence                       = task->SnapshotProgressCadence();
            info.autoConcurrencyDestinationStorageKind = task->_autoConcurrencyDestinationStorageKind.load(std::memory_order_acquire);
            info.progressStallSide                     = task->_progressStallSide.load(std::memory_order_acquire);
            info.bridgeStageCreateCount                = task->_perf.bridgeStageCreateCount.load(std::memory_order_acquire);
            {
                std::scoped_lock lock(task->_progressMutex);
                info.progressTotalItems     = task->_progressTotalItems;
                info.progressCompletedItems = task->_progressCompletedItems;
                info.progressCompletedBytes = task->_progressCompletedBytes;
                info.completedFiles         = task->_completedTopLevelFiles;
                info.completedFolders       = task->_completedTopLevelFolders;
            }
            {
                std::scoped_lock lock(task->_sourceItemStatusMutex);
                info.sourceItemStatuses.reserve(task->_sourceItemResultBuilders.size());
                info.sourceItemResults.reserve(task->_sourceItemResultBuilders.size());
                for (const FolderWindow::FileOperationState::Task::SourceItemResultBuilder& builder : task->_sourceItemResultBuilders)
                {
                    info.sourceItemStatuses.push_back(builder.status);
                    info.sourceItemResults.push_back(builder.terminal);
                }
            }
        }
    }

    state.completedTasks[taskId] = info;
}

bool FileOperationsSelfTest::IsRunning() noexcept
{
    return GetState().running.load(std::memory_order_acquire);
}

bool FileOperationsSelfTest::IsDone() noexcept
{
    return GetState().done.load(std::memory_order_acquire);
}

SelfTest::SelfTestSuiteResult FileOperationsSelfTest::GetSuiteResult() noexcept
{
    SelfTestState& state = GetState();

    SelfTest::SelfTestSuiteResult result{};
    result.suite = SelfTest::SelfTestSuite::FileOperations;

    const ULONGLONG nowTick = GetTickCount64();
    if (state.runStartTick != 0 && nowTick >= static_cast<ULONGLONG>(state.runStartTick))
    {
        result.durationMs = static_cast<uint64_t>(nowTick - static_cast<ULONGLONG>(state.runStartTick));
    }

    result.failureMessage = state.failureMessage;

    std::span<const SelfTestState::Step> reportedPhases = kFileOpsPhaseOrder;
    if (! state.reportedPhaseOrder.empty())
    {
        reportedPhases = state.reportedPhaseOrder;
    }

    result.cases.reserve(reportedPhases.size());
    for (const SelfTestState::Step step : reportedPhases)
    {
        const std::wstring_view expected = StepToString(step);
        const auto it                    = std::find_if(
            state.phaseResults.begin(), state.phaseResults.end(), [&](const SelfTest::SelfTestCaseResult& item) noexcept { return item.name == expected; });
        if (it != state.phaseResults.end())
        {
            SelfTest::AppendCaseResult(result, *it);
            continue;
        }

        const std::wstring_view reason = state.failed.load(std::memory_order_acquire) ? L"not reached (aborted due to failure)" : L"not reached";
        SelfTest::AppendCaseResult(result, expected, SelfTest::SelfTestCaseResult::Status::skipped, reason, 0);
    }

    return result;
}

bool FileOperationsSelfTest::DidFail() noexcept
{
    return GetState().failed.load(std::memory_order_acquire);
}

bool FileOperationsSelfTest::HasRetainedFailureEvidence() noexcept
{
    return GetState().retainTempRootOnFailure;
}

std::wstring_view FileOperationsSelfTest::FailureMessage() noexcept
{
    return GetState().failureMessage;
}

#pragma warning(pop)

#else

bool FileOperationsSelfTest::Tick(HWND /*mainWindow*/) noexcept
{
    return false;
}
void FileOperationsSelfTest::NotifyTaskCompleted(std::uint64_t /*taskId*/, HRESULT /*hr*/) noexcept
{
}
bool FileOperationsSelfTest::IsRunning() noexcept
{
    return false;
}
bool FileOperationsSelfTest::IsDone() noexcept
{
    return false;
}
bool FileOperationsSelfTest::DidFail() noexcept
{
    return false;
}
std::wstring_view FileOperationsSelfTest::FailureMessage() noexcept
{
    return {};
}

#endif // ENABLE_TESTS
