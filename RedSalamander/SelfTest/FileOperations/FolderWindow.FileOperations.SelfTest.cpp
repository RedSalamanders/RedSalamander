#include "FolderWindow.FileOperations.SelfTest.h"

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
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include <AccCtrl.h>
#include <AclAPI.h>
#include <TlHelp32.h>
#include <winioctl.h>
#include <cfapi.h>
#include <winhttp.h>

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

#include "Blake3Digest.h"
#include "ConnectionProfileUtils.h"
#include "ContentDigest.h"
#include "DirectoryInfoCache.h"
#include "FileSystemRouteContract.h"
#include "FileSystemRouteProviderBase.h"
#include "FileSystemPluginManager.h"
#include "FluentIcons.h"
#include "FolderView.h"
#include "FolderWindow.FileOperations.Popup.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "FolderWindow.h"
#include "HostServices.h"
#include "RedSalamander.h"
#include "SplashScreen.h"
#include "TestSandboxPath.h"
#include "WindowMessages.h"
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
                                                 bool exposeTypedRoute = true,
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

    HRESULT STDMETHODCALLTYPE GetPathCapabilities(const wchar_t* path,
                                                  FileSystemOperation operation,
                                                  const char** jsonUtf8) noexcept override
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
    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t*,
                                       const wchar_t*,
                                       FileSystemFlags,
                                       const FileSystemOptions*,
                                       IFileSystemCallback*,
                                       void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t*,
                                       const wchar_t*,
                                       FileSystemFlags,
                                       const FileSystemOptions*,
                                       IFileSystemCallback*,
                                       void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE DeleteItem(const wchar_t*,
                                         FileSystemFlags,
                                         const FileSystemOptions*,
                                         IFileSystemCallback*,
                                         void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t*,
                                         const wchar_t*,
                                         FileSystemFlags,
                                         const FileSystemOptions*,
                                         IFileSystemCallback*,
                                         void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const*,
                                        unsigned long,
                                        const wchar_t*,
                                        FileSystemFlags,
                                        const FileSystemOptions*,
                                        IFileSystemCallback*,
                                        void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const*,
                                        unsigned long,
                                        const wchar_t*,
                                        FileSystemFlags,
                                        const FileSystemOptions*,
                                        IFileSystemCallback*,
                                        void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const*,
                                          unsigned long,
                                          FileSystemFlags,
                                          const FileSystemOptions*,
                                          IFileSystemCallback*,
                                          void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair*,
                                          unsigned long,
                                          FileSystemFlags,
                                          const FileSystemOptions*,
                                          IFileSystemCallback*,
                                          void*) noexcept override
    {
        return WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t*,
                                               FileSystemOperation,
                                               FileSystemTransferEndpoint,
                                               FileSystemTransferHints*) noexcept override
    {
        return _boundedBlockedRead ? E_NOTIMPL : WaitForeverAfterProviderOperationCall();
    }
    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t*, FileSystemStorageCharacteristics*) noexcept override
    {
        return _boundedBlockedRead ? E_NOTIMPL : WaitForeverAfterProviderOperationCall();
    }

protected:
    HRESULT BuildFileSystemRouteDescriptor(const wchar_t*,
                                           FileSystemOperation,
                                           FileSystemRouteDescriptor& descriptor) noexcept override
    {
        descriptor.providerId = _boundedBlockedRead ? L"selftest/bounded-blocked-read" : L"selftest/uncontained-never-return";
        descriptor.pathProfileId = _boundedBlockedRead ? L"selftest-bounded-blocked-read" : L"selftest-uncontained";
        descriptor.rootId = _boundedBlockedRead ? L"selftest-bounded-blocked-read/root" : L"selftest-uncontained/root";
        descriptor.availability = FILESYSTEM_ROUTE_AVAILABLE;
        descriptor.cancellationRoute = _boundedBlockedRead ? FILESYSTEM_CANCELLATION_BOUNDED : FILESYSTEM_CANCELLATION_UNCONTAINED;
        descriptor.namespaceKind = FILESYSTEM_NAMESPACE_PROVIDER_VIRTUAL_FOLDER;
        descriptor.componentComparison = FILESYSTEM_ROUTE_COMPONENT_ORDINAL_CASE_SENSITIVE;
        descriptor.caseOnlyRename = FILESYSTEM_ROUTE_CASE_ONLY_UNSUPPORTED;
        descriptor.copyOperation = _typedCopyOperation;
        descriptor.readOperation = true;
        descriptor.writeOperation = true;
        descriptor.pathTextStableIdentity = true;
        descriptor.exportCopyAll = true;
        descriptor.importCopyAll = true;
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
    bool _exposeTypedRoute = true;
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
    const auto objectWords = reinterpret_cast<const void* const*>(fileSystem);
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

struct WatchCallback;

struct CompletedTaskInfo
{
    HRESULT hr                            = S_OK;
    ULONGLONG completionTick              = 0;
    bool discoveryClosed                 = false;
    bool firstMutationBeforeDiscoveryClosed   = false;
    bool discoverySkipped                   = false;
    uint64_t discoveredTotalBytes            = 0;
    unsigned long discoveredFiles            = 0;
    unsigned long discoveredDirectories      = 0;
    uint64_t discoveryDurationUs            = 0;
    uint64_t discoveryMaxQueueDepth       = 0;
    uint64_t discoveryStarvationCount     = 0;
    uint64_t discoveryCallbackCount       = 0;
    uint64_t discoveryCallbackUs          = 0;
    uint64_t discoveryLockWaitUs          = 0;
    uint64_t discoveryFirstMutationUs = (std::numeric_limits<uint64_t>::max)();
    uint64_t discoveryBytesCompletedWhileOpen = (std::numeric_limits<uint64_t>::max)();
    uint64_t discoveryMutationsCompletedWhileOpen = (std::numeric_limits<uint64_t>::max)();
    uint64_t progressCallbackCount             = 0;
    bool started                          = false;
    unsigned long progressTotalItems      = 0;
    unsigned long progressCompletedItems  = 0;
    uint64_t progressCompletedBytes       = 0;
    unsigned long completedFiles          = 0;
    unsigned long completedFolders        = 0;
    uint64_t conflictWaitUs               = 0;
    uint64_t conflictConvergenceWaitUs    = 0;
    uint64_t conflictPromptCount          = 0;
    uint64_t interlockWaitUs              = 0;
    uint64_t interlockWaitCount           = 0;
    uint64_t bridgeDirectoryEnsureCount   = 0;
    uint64_t bridgeSourceDirectoryEnumerationCount = 0;
    uint64_t bridgeFileAdmissionCount     = 0;
    uint64_t bridgeEarlyFileStartCount    = 0;
    uint64_t bridgeAdmissionMaxQueueDepth = 0;
    uint64_t bridgeTraversalMaxDepth = 0;
    uint64_t bridgeTraversalMaxRetainedEntries = 0;
    uint64_t bridgeTraversalMaxQueuedPathBytes = 0;
    uint64_t bridgeTraversalMaxMetadataBytes = 0;
    uint64_t bridgeTraversalLimitHitCount = 0;
    uint64_t conflictExpectedDestinationBoundCount = 0;
    uint64_t conflictExpectedDestinationUnavailableCount = 0;
    uint64_t conflictExpectedDestinationReturnedCount = 0;
    uint64_t bridgeImmediateDestinationSizeProbeUs = 0;
    uint64_t bridgeImmediateDestinationSizeProbeCount = 0;
    uint64_t bridgeCommitSizeProofCount = 0;
    uint64_t bridgeCommitSizeProofFallbackCount = 0;
    uint64_t verificationUs = 0;
    uint64_t verificationReadBytes = 0;
    uint64_t verificationReadCalls = 0;
    uint64_t verificationProviderProofCount = 0;
    uint64_t verificationHostReadbackCount = 0;
    uint64_t verificationTotalBytes = 0;
    uint64_t verificationCompletedBytes = 0;
    unsigned int configuredMaxConcurrency = 0;
    std::vector<std::optional<HRESULT>> sourceItemStatuses;
    std::vector<std::optional<FileOperations::FileOperationItemResult>> sourceItemResults;
};

// This fixture owns a local Cloud Files registration across asynchronous host-copy ticks.
// Its path storage outlives the WIL registration owner; cleanup records its HRESULT for assertions.
struct PlaceholderSyncRoot final
{
    using SetCompatibilityModeFn = CHAR(NTAPI*)(CHAR);
    std::filesystem::path path;
    HRESULT cleanupHr = S_OK;
    bool registered = false;
    SetCompatibilityModeFn setProcessMode = nullptr;
    SetCompatibilityModeFn setThreadMode = nullptr;
    CHAR previousProcessMode = -1;
    CHAR previousThreadMode = -1;
};

void UnregisterPlaceholderSyncRoot(PlaceholderSyncRoot* root) noexcept
{
    root->cleanupHr = root->registered ? CfUnregisterSyncRoot(root->path.c_str()) : S_OK;
    root->registered = false;
    if (root->previousThreadMode >= 0 && root->setThreadMode(root->previousThreadMode) < 0)
    {
        root->cleanupHr = E_FAIL;
    }
    if (root->previousProcessMode >= 0 && root->setProcessMode(root->previousProcessMode) < 0)
    {
        root->cleanupHr = E_FAIL;
    }
    root->previousThreadMode = -1;
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
        FileOps_CopyMergeIntoExistingFolder,
        FileOps_MoveMergeIntoExistingFolderSameVolume,
        Beeline_RenameMergeSkipKeepsSourceFolder,
        FileOps_ReparseDirectoryMergeIntoExistingFolder,
        FileOps_ProviderCapabilityMatrix,
        FileOps_ParallelGraphFairColorWeight,
        FileOps_CrossVolumeMovePartialFailureStatus,
        FileOps_ReparseLiteralCopyKeepsTargets,
        FileOps_DeleteToctouSwapGuard,
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
        Phase7_SharedPerItemScheduler,
        Phase7_ParallelDeleteKnobs,
        Phase7_RecycleBinBatchDelete,
        Phase7_RecycleBinBatchDeleteMultiBatch,
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
    uint32_t stepState    = 0;
    uint64_t runStartTick = 0;

    std::vector<SelfTest::SelfTestCaseResult> phaseResults;
    bool phaseInProgress     = false;
    ULONGLONG phaseStartTick = 0;
    bool phaseFailed         = false;
    std::wstring phaseName;
    std::wstring phaseFailureMessage;

    HWND mainWindow = nullptr;

    std::filesystem::path tempRoot;
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
    std::optional<std::uint64_t> taskB;
    std::optional<std::uint64_t> taskC;
    bool r4a02ConcurrentAdmissionObserved = false;
    uint32_t r4a02LiveOutputScenario      = 0u;
    bool r4a02LiveOutputQueueObserved     = false;
    bool inlineF2QueueModeOriginal = true;
    bool r4a02QueueModeOriginal    = true;
    HWND c1DeferredCloseTarget     = nullptr;
    ULONGLONG c1HeartbeatStartTick = 0;
    unsigned c1HeartbeatTicks      = 0u;
    uint64_t inlineF2PresentedBaseline = 0;
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
    bool r1dAdmissionReturnedBeforeReadiness = false;
    bool r1dMoveBreadcrumbPublishedBeforeReady = false;
    uint64_t r1dPresentationBaseline = 0u;
    bool phase10MetadataSparseSeeded              = false;
    int64_t phase10MetadataExpectedWriteTime       = 0;

    std::wstring failureMessage;
    bool autoDismissSuccessOriginal            = false;
    ULONGLONG stepStartTick                    = 0;
    ULONGLONG markerTick                       = 0;
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
    ULONGLONG localBandwidthCancelStartTick    = 0;
    uint64_t localBandwidthDurationUs          = 0;
    uint64_t localBandwidthDurationLeadUs      = 0;
    uint64_t localBandwidthCancelLatencyUs     = 0;
    uint64_t localBandwidthMaxWindowBytes      = 0;
    uint64_t localBandwidthMaxSampleDeltaBytes = 0;
    std::vector<std::pair<ULONGLONG, uint64_t>> localBandwidthSamples;
    bool bandwidthThrottleWorkerModeEnvBackedUp    = false;
    bool bandwidthThrottleWorkerModeEnvHadOriginal = false;
    std::wstring bandwidthThrottleWorkerModeEnvOriginal;

    // Holds a no-delete-share handle open across ticks to make a move's source-delete phase
    // fail deterministically (used by FileOps_CrossVolumeMovePartialFailureStatus).
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

    ULONGLONG parallelBandwidthRunStartTick         = 0;
    uint64_t parallelBandwidthBaselineUs            = 0;
    uint64_t parallelBandwidthCandidateUs           = 0;
    uint64_t parallelBandwidthBaselineMaxSkewBytes  = 0;
    uint64_t parallelBandwidthCandidateMaxSkewBytes = 0;
    uint64_t parallelBandwidthBaselineMaxCallbackDeltaBytes  = 0;
    uint64_t parallelBandwidthCandidateMaxCallbackDeltaBytes = 0;
    size_t parallelBandwidthBaselineMaxActive       = 0;
    size_t parallelBandwidthCandidateMaxActive      = 0;
    uint64_t parallelBandwidthBaselineSamples       = 0;
    uint64_t parallelBandwidthCandidateSamples      = 0;
    ULONGLONG defaultSpeedLimitRunStartTick         = 0;
    uint64_t defaultSpeedLimitBaselineUs            = 0;
    uint64_t defaultSpeedLimitCandidateUs           = 0;
    std::string defaultSpeedLimitDummyConfigSnapshot;
    ULONGLONG recycleBinBatchRunStartTick                   = 0;
    uint64_t recycleBinBatchBaselineUs                      = 0;
    uint64_t recycleBinBatchCandidateUs                     = 0;
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
        case SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder: return L"FileOps_CopyMergeIntoExistingFolder";
        case SelfTestState::Step::FileOps_MoveMergeIntoExistingFolderSameVolume: return L"FileOps_MoveMergeIntoExistingFolderSameVolume";
        case SelfTestState::Step::Beeline_RenameMergeSkipKeepsSourceFolder: return L"Beeline_RenameMergeSkipKeepsSourceFolder";
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
        case SelfTestState::Step::Cinderstar_BridgeMovePerFileConflictSkipsParallel:
            return L"Cinderstar_BridgeMovePerFileConflictSkipsParallel";
        case SelfTestState::Step::Riptide_BridgeNestedDirVsFileSkipContinuesSiblings: return L"Riptide_BridgeNestedDirVsFileSkipContinuesSiblings";
        case SelfTestState::Step::Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch:
            return L"Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch";
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
        case SelfTestState::Step::Cinderstar_LegacyWriterEventuallyConsistentMove:
            return L"Cinderstar_LegacyWriterEventuallyConsistentMove";
        case SelfTestState::Step::Cinderstar_LegacyWriterPermanentMissMove: return L"Cinderstar_LegacyWriterPermanentMissMove";
        case SelfTestState::Step::Cinderstar_LegacyWriterWrongSizeMove: return L"Cinderstar_LegacyWriterWrongSizeMove";
        case SelfTestState::Step::Cinderstar_LegacyWriterConcurrentReplacementCopy:
            return L"Cinderstar_LegacyWriterConcurrentReplacementCopy";
        case SelfTestState::Step::Cinderstar_LegacyWriterCancelDuringBackoff: return L"Cinderstar_LegacyWriterCancelDuringBackoff";
        case SelfTestState::Step::Floodgate_CrossFsCopyOnlyUsesCopyVerification:
            return L"Floodgate_CrossFsCopyOnlyUsesCopyVerification";
        case SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree:
            return L"Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree";
        case SelfTestState::Step::Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup: return L"Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup";
        case SelfTestState::Step::Floodgate_CrossFsDirectoryCopyOnlyRetainsSource:
            return L"Floodgate_CrossFsDirectoryCopyOnlyRetainsSource";
        case SelfTestState::Step::Floodgate_LocalWriterOverwriteIsStaged: return L"Floodgate_LocalWriterOverwriteIsStaged";
        case SelfTestState::Step::Floodgate_LocalCopyOverwriteIsStaged: return L"Floodgate_LocalCopyOverwriteIsStaged";
        case SelfTestState::Step::Floodgate_LocalCopyNewNameConcurrentReplacementSurvives:
            return L"Floodgate_LocalCopyNewNameConcurrentReplacementSurvives";
        case SelfTestState::Step::Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete:
            return L"Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete";
        case SelfTestState::Step::Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker:
            return L"Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker";
        case SelfTestState::Step::Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker:
            return L"Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker";
        case SelfTestState::Step::Floodgate_InlineF2WorkerQueueBypassAndInterlock:
            return L"Floodgate_InlineF2WorkerQueueBypassAndInterlock";
        case SelfTestState::Step::C1_InlineRenameNameRefusedOnCard: return L"C1_InlineRenameNameRefusedOnCard";
        case SelfTestState::Step::R4A19_DiscoveryMeasurementFacts: return L"R4A19_DiscoveryMeasurementFacts";
        case SelfTestState::Step::R4A19_DiscoveryProviderControls: return L"R4A19_DiscoveryProviderControls";
        case SelfTestState::Step::R4A19_DiscoveryIndependentVolumes: return L"R4A19_DiscoveryIndependentVolumes";
        case SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns: return L"R0fSmb_BlockedSynchronousCallCancelReturns";
        case SelfTestState::Step::R0fSmb_LoopbackReadWriteCreateDelete: return L"R0fSmb_LoopbackReadWriteCreateDelete";
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
        case SelfTestState::Step::FileOps_CrossVolumeMovePartialFailureStatus: return L"FileOps_CrossVolumeMovePartialFailureStatus";
        case SelfTestState::Step::FileOps_ReparseLiteralCopyKeepsTargets: return L"FileOps_ReparseLiteralCopyKeepsTargets";
        case SelfTestState::Step::FileOps_DeleteToctouSwapGuard: return L"FileOps_DeleteToctouSwapGuard";
        case SelfTestState::Step::FileOps_ResolvedItemsExactDestinations: return L"FileOps_ResolvedItemsExactDestinations";
        case SelfTestState::Step::Phase5_DiscoverySingleTraversal: return L"Phase5_DiscoverySingleTraversal";
        case SelfTestState::Step::Phase5_DiscoveryCancelReleasesSlot: return L"Phase5_DiscoveryCancelReleasesSlot";
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
        case SelfTestState::Step::Phase7_SharedPerItemScheduler: return L"Phase7_SharedPerItemScheduler";
        case SelfTestState::Step::Phase7_ParallelDeleteKnobs: return L"Phase7_ParallelDeleteKnobs";
        case SelfTestState::Step::Phase7_RecycleBinBatchDelete: return L"Phase7_RecycleBinBatchDelete";
        case SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch: return L"Phase7_RecycleBinBatchDeleteMultiBatch";
        case SelfTestState::Step::Phase8_DefaultBandwidthLimitFromSettings: return L"Phase8_DefaultBandwidthLimitFromSettings";
        case SelfTestState::Step::Phase8_TightDefaults_NoOverwrite: return L"Phase8_TightDefaults_NoOverwrite";
        case SelfTestState::Step::Phase8_InvalidDestinationRejected: return L"Phase8_InvalidDestinationRejected";
        case SelfTestState::Step::Phase8_InvalidSizeBytesRejected: return L"Phase8_InvalidSizeBytesRejected";
        case SelfTestState::Step::Phase8_PerItemOrchestration: return L"Phase8_PerItemOrchestration";
        case SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly: return L"Phase9_ConflictPrompt_OverwriteReplaceReadonly";
        case SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache: return L"Phase9_ConflictPrompt_ApplyToAllUiCache";
        case SelfTestState::Step::Phase9_ConflictPrompt_KeepBothNestedCacheEligibility:
            return L"Phase9_ConflictPrompt_KeepBothNestedCacheEligibility";
        case SelfTestState::Step::Phase9_ConflictPrompt_TypeMismatchNoOverwrite: return L"Phase9_ConflictPrompt_TypeMismatchNoOverwrite";
        case SelfTestState::Step::Phase9_ConflictPrompt_LocalFileOntoDirectory: return L"Phase9_ConflictPrompt_LocalFileOntoDirectory";
        case SelfTestState::Step::Phase9_ConflictPrompt_SkipApplyToAll: return L"Phase9_ConflictPrompt_SkipApplyToAll";
        case SelfTestState::Step::Phase9_ConflictPrompt_RetryCap: return L"Phase9_ConflictPrompt_RetryCap";
        case SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy: return L"Phase9_ConflictPrompt_SkipContinuesDirectoryCopy";
        case SelfTestState::Step::Phase9_PerItemConcurrency: return L"Phase9_PerItemConcurrency";
        case SelfTestState::Step::Phase10_PermanentDelete: return L"Phase10_PermanentDelete";
        case SelfTestState::Step::Phase10_DeferredConsentAndRecycleEscalation:
            return L"Phase10_DeferredConsentAndRecycleEscalation";
        case SelfTestState::Step::Phase10_MetadataPreservationAndSourceRetention:
            return L"Phase10_MetadataPreservationAndSourceRetention";
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
        case SelfTestState::Step::Cleanup_RestorePluginConfig: return L"Cleanup_RestorePluginConfig";
        case SelfTestState::Step::Done: return L"Done";
        case SelfTestState::Step::Failed: return L"Failed";
        default: break;
    }
    return L"(unknown)";
}

constexpr auto kFileOpsPhaseOrder = std::to_array<SelfTestState::Step>({
    SelfTestState::Step::Setup,                                           // Environment setup and plugin loading
    SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder,             // Clearflow Phase 1 - copy folder into existing destination folder merges
    SelfTestState::Step::FileOps_MoveMergeIntoExistingFolderSameVolume,   // Clearflow Phase 1 - same-volume move folder into existing destination folder merges
    SelfTestState::Step::Beeline_RenameMergeSkipKeepsSourceFolder,        // Beeline - rename merge answers Overwrite/Skip per child and keeps the source folder
    SelfTestState::Step::FileOps_ReparseDirectoryMergeIntoExistingFolder, // Clearflow Phase 1 - directory reparse copy requires overwrite for existing dir
    SelfTestState::Step::FileOps_ProviderCapabilityMatrix,                // Clearflow Phase 7 - provider capability matrix and offline conformance
    SelfTestState::Step::FileOps_ParallelGraphFairColorWeight,            // Clearflow Phase 3 - parallel graph hue weights are fair
    SelfTestState::Step::FileOps_CrossVolumeMovePartialFailureStatus,     // Clearflow Phase 4 - partial move status is explicit
    SelfTestState::Step::FileOps_ReparseLiteralCopyKeepsTargets,   // Clearflow Phase 4 - copied reparse links keep their stored targets (C3 literal)
    SelfTestState::Step::FileOps_DeleteToctouSwapGuard,                   // Clearflow Phase 4 - delete path rejects stale check/act swaps
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
    SelfTestState::Step::Fairstream_MovedTreeKeepsLiteralLinks,           // Fairstream 1E - moved trees keep literal link text
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
    SelfTestState::Step::C1_InlineRenameNameRefusedOnCard,               // C1 (b0) - an inline rename the provider refuses fails on its card, not in admission
    SelfTestState::Step::R4A19_DiscoveryMeasurementFacts,                 // R4-A19 - observation-only discovery baseline facts
    SelfTestState::Step::R4A19_DiscoveryProviderControls,                 // R4-A19 - delayed Dummy and serialized fake-MTP controls
    SelfTestState::Step::R4A19_DiscoveryIndependentVolumes,               // R4-A19 - marked C/D independent task overlap
    SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns,      // R0f-SMB - a wedged synchronous provider call returns on cancel
    SelfTestState::Step::R0fSmb_LoopbackReadWriteCreateDelete,            // R0f-SMB - read/write/create/rename/delete on a loopback share
    SelfTestState::Step::R0fCurl_FakeFtpReadWriteCreateDelete,            // R0f-Curl - cross-provider copy both ways, create/rename/delete on fake FTP
    SelfTestState::Step::BR5_CurlHostDirectoryMove,
    SelfTestState::Step::BR5_CurlHostDirectoryMoveRefused,
    SelfTestState::Step::R0fS3_FakeS3ReadWriteCreateDelete,               // R0f-S3 - cross-provider copy both ways, create/rename/delete on fake S3
    SelfTestState::Step::R0fGraph_FakeGraphReadWriteCreateRenameRecycle,  // R0f-Graph - cross-provider copy both ways, create/rename/recycle on fake Graph
    SelfTestState::Step::R0fGDrive_FakeDriveReadWriteCreateMoveDelete,    // R0f-GDrive - cross-provider copy both ways, create/rename/move/delete on fake Drive
    SelfTestState::Step::C0_GDriveCommittedCopyResponseLost,              // C0/C2 - committed remote copy must never be replayed after response loss
    SelfTestState::Step::C0_GDriveCommittedDeleteResponseLost,
    SelfTestState::Step::C0_NativeCopyFailurePreservesKnownAxes,
    SelfTestState::Step::C0_MutationReceiptPrefixBoundary,
    SelfTestState::Step::C0_LocalKnownNoCommitReceipts,                   // C0 - Local Cancel/refused-rename/merge-cancel receipts stay known, never uncertain
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
    SelfTestState::Step::C0_CurlImapTransportTruth,                       // C0 - production IMAP owners through real libcurl against a loopback server
    SelfTestState::Step::R3_1_IdentityLessReplaceDummy,                   // R3-1 - conditional replace through an atomic-final writer without bound authority
    SelfTestState::Step::R3_2_WriterProofDummy,       // R3-2 - writer content proof: Verified Copy without readback, Managed Move cleanup after proof
    SelfTestState::Step::R3_3_PublicationFaultMatrix, // R3-3 - every bridge fault hook x route x Copy/Move against the publication invariants
    SelfTestState::Step::R3_4_ReplaceWithoutOccupantTokenRefused, // R0-RC3 - a granted replacement without an occupant token is refused, nothing written
    SelfTestState::Step::RC3_9_TransferDestinationNameRefused,    // R0-RC3 (9) - a destination leaf the provider rejects is refused at admission
    SelfTestState::Step::R4T_DeepTreeCopyCompletesIteratively,    // R4-T1 - a 300-level tree copies and moves through the iterative bridge walkers
    SelfTestState::Step::R4T2_DeepTreeLocalDeleteCompletes,       // R4-T2 - a 300-level Local chain beyond MAX_PATH is deleted through the frame-based walk
    SelfTestState::Step::R4T3_WideDirectoryCopyCompletes, // R4-T3 - a 20,000-child directory copies through the bridge with its listing reported, not capped
    SelfTestState::Step::R4A02_DeleteOverCopySourceWarnsAndQueues,    // R4-A02-1 - a Delete over a running Copy's source warns and queues behind it
    SelfTestState::Step::R4A02_DontStartCancelsBeforeMutation,        // R4-A02-1 - Don't start ends the overlapping task before any mutation
    SelfTestState::Step::R4A02_ReadReadDoesNotWarn,                   // R4-A02-1 - two Copies reading one source warn nothing
    SelfTestState::Step::R4A02_SameDestinationWarns,                  // R4-A02-1 - simultaneous Copies publishing the same leaf warn before mutation
    SelfTestState::Step::R4A02_RunConcurrentSameDestinationStaysSafe, // R4-A02-2 - Run admits only the disclosed overlap and both result sets stay truthful
    SelfTestState::Step::R4A02_LiveOutputGuardChoices,                // R4-A02-2 - uncovered live output parks; Skip/Queue/destructive and overflow stay safe
    SelfTestState::Step::R4A02_DisjointWritesDoNotWarn,               // R4-A02-1 - disjoint source and destination roots warn nothing
    SelfTestState::Step::R4A02_RenamePublishKeepsIndexPrecise,         // Step 3 (2026-09-06) - a rename publication is indexed under its parent scope, no host-wide overflow
    SelfTestState::Step::R4A02_LatePublisherStillWarns,               // Step 4 (2026-09-06) - a task that publishes after a newer peer prepared still warns and queues after it
    SelfTestState::Step::Phase5_DiscoverySingleTraversal,             // Phase 5 — one traversal feeds execution and closes totals
    SelfTestState::Step::Phase5_DiscoveryCancelReleasesSlot,          // Phase 5 — cancel releases discovery/execution capacity
    SelfTestState::Step::Phase5_DiscoveryCancelLatencyLocal,          // Phase 5 — local discovery/transfer observes bounded cancellation
    SelfTestState::Step::Phase5_DiscoverySkipContinues,               // Phase 5 — Skip switches to just-in-time and completes the same selection
    SelfTestState::Step::BR3_CancelQueuePausedTransfer,               // Cancellation drains independently of the paused queue predecessor
    SelfTestState::Step::Phase5_CancelQueuedTask,                     // Phase 5 â€” canceling a queued (not-yet-running) task
    SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery,  // Phase 5 — mode switch parallel→wait while discovery remains open
    SelfTestState::Step::Phase5_SwitchWaitToParallelResume,           // Phase 5 â€” mode switch waitâ†’parallel and resume
    SelfTestState::Step::Phase6_PopupRateSmoothing,                   // Phase 6 - popup rate/ETA smoothing contract
    SelfTestState::Step::Phase6_PopupSmokeResizeAndPause,             // Phase 6 â€” popup resize and pause-button interaction
    SelfTestState::Step::Phase6_DeleteBytesMeaningful,                // Phase 6 â€” delete reports meaningful byte counts in progress
    SelfTestState::Step::Phase6_LocalBandwidthThrottle,               // Phase 6 â€” local CopyFileEx bandwidth throttle duration and cancel latency
    SelfTestState::Step::Phase6_ParallelBandwidthThrottleFairness,    // Phase 6 â€” parallel throttle fairness: shared-only vs per-worker budget
    SelfTestState::Step::Phase7_WatcherChurn,                         // Phase 7 â€” directory watcher fires correctly under heavy churn
    SelfTestState::Step::Phase7_CacheBorrowNoWatchInvalidation,       // Phase 7 â€” cached-but-unwatched folders still invalidate on routed mutations
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshLocal,         // Phase 7 â€” visible local panes auto-refresh on watcher-driven changes
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshDummy,         // Phase 7 â€” visible dummy panes stay in sync after in-app mutation routing
    SelfTestState::Step::Phase7_CrossPaneRelocateLocal,               // Phase 7 â€” deleting a visible subtree relocates impacted panes
    SelfTestState::Step::Phase7_LargeDirectoryEnumeration,            // Phase 7 â€” enumerate a directory with many entries
    SelfTestState::Step::Phase7_ParallelCopyMoveKnobs,                // Phase 7 â€” speed limits and parallelism knobs for copy/move
    SelfTestState::Step::Phase7_CopyMoveConcurrency16Perf,            // Phase 7 / 8.2 â€” compare copy/move cap 8 vs 16 on the same machine
    SelfTestState::Step::Phase7_AutoConcurrencyHints,                 // Phase 7 â€” auto mode resolves copy/delete concurrency from storage hints
    SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines,    // Phase 7 â€” per-item directory copy uses internal parallelism (popup lines)
    SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism, // Phase 7 - plugin CopyItem/CopyItems(count==1) use recursive internal parallelism
    SelfTestState::Step::Phase7_CopyItemsMultiRootUnevenRecursiveParallelism,  // Phase 7 - multi-root copies rebalance uneven recursive subtrees
    SelfTestState::Step::Phase7_CopyRecursiveParallelismMatrix,                // Phase 7 - recursive copy/move matrix shapes and edge paths
    SelfTestState::Step::Phase7_SharedPerItemScheduler,                        // Phase 7 â€” shared per-item scheduler across parallel tasks
    SelfTestState::Step::Phase7_ParallelDeleteKnobs,                           // Phase 7 â€” speed limits and parallelism knobs for delete
    SelfTestState::Step::Phase7_RecycleBinBatchDelete,                         // Phase 7 â€” multi-item recycle-bin delete batches sibling items
    SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch,               // Phase 7 â€” recycle-bin delete spans multiple sibling batches when >500 inputs
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
    SelfTestState::Step::Phase10_MetadataPreservationAndSourceRetention,       // Phase 10 / P4.3 — exact metadata and security-loss retention
    SelfTestState::Step::Phase10_TypedResultsAndConsumers,                     // Phase 10 / P4.4 — typed item axes and exact consumer removal
    SelfTestState::Step::Phase10_ClipboardAdmissionAndRetainedActions,         // Phase 10 / P4.5 — accepted cut-list barrier and retained actions
    SelfTestState::Step::Phase10_ContentVerification,                          // Phase 10 / P4.6 — optional BLAKE3 verification and truthful outcomes
    SelfTestState::Step::Phase10_ArtifactTouchGuard,                           // Phase 10 / P5.4 — external launch artifact warning and exact revalidation
    SelfTestState::Step::R1d_PreparingLifecycle,                               // R1d — common pre-consumption lifecycle and immutable preparation facts
    SelfTestState::Step::C1_ExitCloseDeferredUntilTasksQuiet,                  // C1 - application exit drains cancellation off the UI thread; the close resumes after the last reap
    SelfTestState::Step::Beeline_SameVolumeTreeMoveIsRename,                   // Beeline - a same-volume file and tree Move is one Native rename per item, identity preserved
    SelfTestState::Step::Beeline_LoopbackShareMoveIsRename,                    // Beeline - the SMB profile takes the same route: a Move inside the loopback share alias is one Native rename per item
    SelfTestState::Step::C1_PermanentDeleteConfirmsOnCard,                     // C1 - the permanent-delete confirmation is a card consent after Preparing pins the roots
    SelfTestState::Step::Phase11_CrossFileSystemBridge,                        // Phase 11 â€” copy/move across different file-system plugins
    SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines,  // Phase 11 â€” bridge: single-folder copy uses within-folder parallelism
    SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines,   // Phase 11 â€” bridge: multi-folder copy keeps bounded early file admission
    SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf,               // Phase 11 â€” bridge: pipeline overlaps dummy read/write chunk latency
    SelfTestState::Step::Phase11_ConnectionOverridePrecedence,                 // Phase 11 / 8 â€” non-zero @conn override takes precedence over plugin default
    SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate,                 // Phase 11 â€” connection overrides apply globally across tasks
    SelfTestState::Step::Phase11_ConnectionOverrideClamp,                      // Phase 11 â€” connection manager overrides clamp per-task concurrency
    SelfTestState::Step::Phase12_ReparsePointPolicy,                           // Phase 12 â€” reparse-point (symlink/junction) handling policy
    SelfTestState::Step::Phase13_PostMortemDiagnostics,                        // Phase 13 â€” post-mortem diagnostics on task failure
    SelfTestState::Step::Phase14_PopupHostLifetimeGuard,                       // Phase 14 â€” popup host lifetime guard (no UAF on late input)
    SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke,                    // Phase 15 â€” 7z IFileReader read/seek smoke on a large (>32MB) entry
    SelfTestState::Step::Phase15_FileSystem7zMountPathImpact,                  // Phase 15 â€” 7z mounted panes retarget/exit on backing archive changes
    SelfTestState::Step::Phase16_RemoteWatchContractExposure,                  // Phase 16 â€” mutable remote plugins expose the watch contract offline
    SelfTestState::Step::Phase16_RemoteFtpSecret,                              // Phase 16 â€” secure secret retrieval (FTP)
    SelfTestState::Step::Phase16_RemoteFtpSandbox,                             // Phase 16 â€” remote sandbox root configuration (FTP)
    SelfTestState::Step::Phase16_RemoteSftpSecret,                             // Phase 16 â€” secure secret retrieval (SFTP)
    SelfTestState::Step::Phase16_RemoteSftpSandbox,                            // Phase 16 â€” remote sandbox root configuration (SFTP)
    SelfTestState::Step::Phase16_RemoteScpSecret,                              // Phase 16 â€” secure secret retrieval (SCP)
    SelfTestState::Step::Phase16_RemoteScpSandbox,                             // Phase 16 â€” remote sandbox root configuration (SCP)
    SelfTestState::Step::Phase16_RemoteImapSecret,                             // Phase 16 â€” secure secret retrieval (IMAP)
    SelfTestState::Step::Phase16_RemoteImapSandbox,                            // Phase 16 â€” remote sandbox root configuration (IMAP)
    SelfTestState::Step::Phase16_RemoteS3Secret,                               // Phase 16 â€” secure secret retrieval (S3)
    SelfTestState::Step::Phase16_RemoteS3Sandbox,                              // Phase 16 â€” remote sandbox root configuration (S3)
    SelfTestState::Step::Phase16_RemoteS3FileOps,                              // Phase 16 â€” sandboxed CRUD/copy-move coverage (S3)
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret,                 // Phase 16 â€” secure secret retrieval (OneDrive Personal)
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSandbox,                // Phase 16 â€” remote sandbox root configuration (OneDrive Personal)
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalFileOps,                // Phase 16 â€” sandboxed CRUD/copy-move coverage (OneDrive Personal)
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret,                 // Phase 16 â€” secure secret retrieval (OneDrive Business)
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSandbox,                // Phase 16 â€” remote sandbox root configuration (OneDrive Business)
    SelfTestState::Step::Phase16_RemoteSharePointSecret,                       // Phase 16 â€” secure secret retrieval (SharePoint)
    SelfTestState::Step::Phase16_RemoteSharePointSandbox,                      // Phase 16 â€” remote sandbox root configuration (SharePoint)
    SelfTestState::Step::Cleanup_RestorePluginConfig                           // Restore plugin config and delete temp files
});

constexpr std::array<SelfTestState::Step, 4> kFileOpsFamilyClearflowPhase01{{
    SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder,
    SelfTestState::Step::FileOps_MoveMergeIntoExistingFolderSameVolume,
    SelfTestState::Step::Beeline_RenameMergeSkipKeepsSourceFolder,
    SelfTestState::Step::FileOps_ReparseDirectoryMergeIntoExistingFolder,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyClearflowPhase07{{
    SelfTestState::Step::FileOps_ProviderCapabilityMatrix,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyClearflowPhase03{{
    SelfTestState::Step::FileOps_ParallelGraphFairColorWeight,
}};

constexpr std::array<SelfTestState::Step, 4> kFileOpsFamilyClearflowPhase04{{
    SelfTestState::Step::FileOps_CrossVolumeMovePartialFailureStatus,
    SelfTestState::Step::FileOps_ReparseLiteralCopyKeepsTargets,
    SelfTestState::Step::FileOps_DeleteToctouSwapGuard,
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

constexpr std::array<SelfTestState::Step, 8> kFileOpsFamilyPhase05{{
    SelfTestState::Step::Phase5_DiscoverySingleTraversal,
    SelfTestState::Step::Phase5_DiscoveryCancelReleasesSlot,
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

constexpr std::array<SelfTestState::Step, 17> kFileOpsFamilyPhase07{{
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
    SelfTestState::Step::Phase7_SharedPerItemScheduler,
    SelfTestState::Step::Phase7_ParallelDeleteKnobs,
    SelfTestState::Step::Phase7_RecycleBinBatchDelete,
    SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch,
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

constexpr std::array<SelfTestState::Step, 7> kFileOpsFamilyPhase10{{
    SelfTestState::Step::Phase10_PermanentDelete,
    SelfTestState::Step::Phase10_DeferredConsentAndRecycleEscalation,
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

constexpr std::array<SelfTestState::Step, 5> kFileOpsFamilyR1dPreparing{{
    SelfTestState::Step::R1d_PreparingLifecycle,
    SelfTestState::Step::C1_ExitCloseDeferredUntilTasksQuiet,
    SelfTestState::Step::Beeline_SameVolumeTreeMoveIsRename,
    SelfTestState::Step::Beeline_LoopbackShareMoveIsRename,
    SelfTestState::Step::C1_PermanentDeleteConfirmsOnCard,
}};

constexpr std::array<SelfTestState::Step, 2> kFileOpsFamilyR0fSmbContainment{{
    SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns,
    SelfTestState::Step::R0fSmb_LoopbackReadWriteCreateDelete,
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

constexpr std::array<SelfTestState::Step, 3> kFileOpsFamilyR4Traversal{{
    SelfTestState::Step::R4T_DeepTreeCopyCompletesIteratively,
    SelfTestState::Step::R4T2_DeepTreeLocalDeleteCompletes,
    SelfTestState::Step::R4T3_WideDirectoryCopyCompletes,
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

constexpr auto kFileOpsFamilyDefinitions = std::to_array<FileOpsFamilyDefinition>({
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase01_DirectoryMerge", {kFileOpsFamilyClearflowPhase01}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase03_GraphFairness", {kFileOpsFamilyClearflowPhase03}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase04_Security", {kFileOpsFamilyClearflowPhase04}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Fairstream", {kFileOpsFamilyFairstream}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase05_PerfEvaluation", {kFileOpsFamilyClearflowPhase05}},
    FileOpsFamilyDefinition{L"FileOpsFamily_ClearflowPhase07_ProviderMatrix", {kFileOpsFamilyClearflowPhase07}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase05_Discovery", {kFileOpsFamilyPhase05}},
    FileOpsFamilyDefinition{L"FileOpsFamily_R4A19DiscoveryBaseline", {kFileOpsFamilyR4A19DiscoveryBaseline}},
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

    state.phaseResults.push_back(std::move(item));
    state.phaseInProgress = false;
    state.phaseName.clear();
    state.phaseStartTick = 0;
    state.phaseFailed    = false;
    state.phaseFailureMessage.clear();
}

void BeginPhase(SelfTestState& state, SelfTestState::Step step) noexcept
{
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

void PerformCleanup(SelfTestState& state) noexcept;

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
    PerformCleanup(state);
    RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);

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
                state.localConfigDirty = true;
                g_settings.plugins.configurationByPluginId[std::wstring(kPluginIdLocal)] = std::move(parsedConfiguration);
            }
            else if (info == state.infoDummy.get())
            {
                state.dummyConfigDirty = true;
                g_settings.plugins.configurationByPluginId[std::wstring(kPluginIdDummy)] = std::move(parsedConfiguration);
            }
            else if (info == state.info7z.get())
            {
                state.config7zDirty = true;
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

void PerformCleanup(SelfTestState& state) noexcept
{
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
    ReleaseFileOpsPermanentDeleteBeforeBindPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeBindPauseForSelfTest(false);
    ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
    ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
    SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_CASE_RENAME_ENTROPY_FAIL_PATH", nullptr));
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_PATH", nullptr));
    static_cast<void>(SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_FIRED", nullptr));

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
    if (! state.fileOpsAlternateVolumeRoot.empty())
    {
        AppendLog(std::format(L"PerformCleanup: delete alternate-volume root {}", state.fileOpsAlternateVolumeRoot.wstring()));
        std::error_code cleanupEc;
        const std::filesystem::path alternateRoot = std::move(state.fileOpsAlternateVolumeRoot);
        std::filesystem::remove_all(alternateRoot, cleanupEc);
        PruneEmptyAlternateVolumeSandboxParents(alternateRoot);
    }

    if (! state.tempRoot.empty())
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

    AppendLog(L"PerformCleanup: defer COM/plugin release to process shutdown");
    AppendLog(L"PerformCleanup: done");
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
    bool copyOperation       = false;
    bool moveOperation       = false;
    bool nativeMoveOperation = false;
    bool deleteOperation     = false;
    bool renameOperation     = false;
    bool recycleOperation    = false;
    bool createDirectoryOperation = false;
    bool properties          = false;
    bool read                = false;
    bool write               = false;
    uint64_t copyMoveMax        = 0;
    uint64_t deleteMax          = 0;
    uint64_t deleteRecycleMax   = 0;
    bool exportCopyWildcard     = false;
    bool exportMoveWildcard     = false;
    bool importCopyWildcard     = false;
    bool importMoveWildcard     = false;
    bool preserveFileLink       = false;
    bool preserveDirectoryLink  = false;
    bool retargetInTree         = false;
    bool exactLinkRemoval       = false;
    bool pathIdentityPresent    = false;
    bool pathTextStableIdentity = false;
    std::wstring componentComparison;
    wchar_t preferredSeparator = L'\0';
    std::wstring acceptedSeparators;
    bool casePreserving = false;
    std::wstring caseOnlyRename;
};

class FlagSensitiveAtomicWriter final : public IFileSystemAtomicWriter
{
public:
    explicit FlagSensitiveAtomicWriter(FileSystemFlags expectedFlags) noexcept : _expectedFlags(expectedFlags) {}

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
        yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, nullptr),
        &yyjson_doc_free);
    if (! doc)
    {
        return false;
    }
    yyjson_val* root = yyjson_doc_get_root(doc.get());
    yyjson_val* version = root ? yyjson_obj_get(root, "version") : nullptr;
    if (! root || ! yyjson_is_obj(root) || ! version || ! yyjson_is_int(version) || yyjson_get_int(version) != 2)
    {
        return false;
    }

    constexpr std::array<const char*, 11u> requiredObjects{{
        "operations", "concurrency", "transfer", "names", "identity", "publication", "links",
        "metadata", "verification", "cancellation", "directories",
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
    bool ignored = false;
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
    yyjson_val* watchdog = yyjson_obj_get(cancellation, "providerWatchdogTimeoutMs");
    uint64_t watchdogMs = 0u;
    if (! routeClass || ! yyjson_is_str(routeClass) || ! TryGetJsonUInt(cancellation, "providerWatchdogTimeoutMs", watchdogMs) ||
        watchdog != yyjson_obj_get(cancellation, "providerWatchdogTimeoutMs"))
    {
        return false;
    }
    const std::string_view routeText(yyjson_get_str(routeClass), yyjson_get_len(routeClass));
    if ((routeText == "providerWatchdog") != (watchdogMs != 0u) ||
        (routeText != "bounded" && routeText != "uncontained" && routeText != "providerWatchdog"))
    {
        return false;
    }

    yyjson_val* names = yyjson_obj_get(root, "names");
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
    const HRESULT hr = fs->GetPathCapabilities(L"/", FILESYSTEM_COPY, &jsonUtf8);
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
        ! TryGetJsonBool(operations, "nativeMove", snapshot.nativeMoveOperation) ||
        ! TryGetJsonBool(operations, "delete", snapshot.deleteOperation) || ! TryGetJsonBool(operations, "rename", snapshot.renameOperation) ||
        ! TryGetJsonBool(operations, "createDirectory", snapshot.createDirectoryOperation) ||
        ! TryGetJsonBool(operations, "properties", snapshot.properties) || ! TryGetJsonBool(operations, "read", snapshot.read) ||
        ! TryGetJsonBool(operations, "write", snapshot.write) || ! TryGetJsonBool(operations, "recycle", snapshot.recycleOperation))
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
    if (! links || ! yyjson_is_obj(links) ||
        ! TryGetJsonBool(links, "preserveFileLink", snapshot.preserveFileLink) ||
        ! TryGetJsonBool(links, "preserveDirectoryLink", snapshot.preserveDirectoryLink) ||
        ! TryGetJsonBool(links, "retargetInTree", snapshot.retargetInTree) ||
        ! TryGetJsonBool(links, "exactLinkRemoval", snapshot.exactLinkRemoval))
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

[[nodiscard]] bool TryReadTypedProviderCapabilities(IFileSystem* fs,
                                                    std::wstring_view providerId,
                                                    ProviderCapabilitySnapshot& snapshot,
                                                    std::wstring& reasonOut,
                                                    bool* usedArenaFallback = nullptr) noexcept
{
    snapshot = {};
    reasonOut.clear();
    if (usedArenaFallback != nullptr)
    {
        *usedArenaFallback = false;
    }
    const FileSystemRouteContract::QueryResult route =
        FileSystemRouteContract::Query(fs, L"/", FILESYSTEM_COPY, providerId);
    if (route.state != FileSystemRouteContract::QueryState::Available || ! route.snapshot.pathIdentity.has_value())
    {
        reasonOut = std::format(L"typed route query failed state={} hr=0x{:08X}",
                                static_cast<unsigned int>(route.state),
                                static_cast<unsigned long>(route.status));
        return false;
    }
    if (usedArenaFallback != nullptr)
    {
        *usedArenaFallback = route.usedArenaFallback;
    }

    const FileSystemRouteContract::Snapshot& facts = route.snapshot;
    snapshot.copyOperation = facts.copyOperation;
    snapshot.moveOperation = facts.moveOperation;
    snapshot.nativeMoveOperation = facts.nativeMoveOperation;
    snapshot.deleteOperation = facts.deleteOperation;
    snapshot.renameOperation = facts.renameOperation;
    snapshot.recycleOperation = facts.recycleOperation;
    snapshot.createDirectoryOperation = facts.createDirectoryOperation;
    snapshot.properties = facts.properties;
    snapshot.read = facts.read;
    snapshot.write = facts.write;
    snapshot.copyMoveMax = facts.copyMoveMaxConcurrency;
    snapshot.deleteMax = facts.deleteMaxConcurrency;
    snapshot.deleteRecycleMax = facts.deleteRecycleBinMaxConcurrency;
    snapshot.preserveFileLink = facts.preserveFileLink;
    snapshot.preserveDirectoryLink = facts.preserveDirectoryLink;
    snapshot.retargetInTree = facts.retargetInTree;
    snapshot.exactLinkRemoval = facts.exactLinkRemoval;
    snapshot.pathIdentityPresent = true;
    snapshot.pathTextStableIdentity = facts.pathIdentity->pathTextStableIdentity;
    snapshot.componentComparison = facts.pathIdentity->componentComparison == FileSystemPathComponentComparison::OrdinalIgnoreCase
        ? L"ordinalIgnoreCase"
        : L"ordinalCaseSensitive";
    snapshot.preferredSeparator = facts.pathIdentity->preferredSeparator;
    snapshot.acceptedSeparators = facts.pathIdentity->acceptedSeparators;
    snapshot.casePreserving = facts.pathIdentity->casePreserving;
    switch (facts.pathIdentity->caseOnlyRename)
    {
        case FileSystemPathCaseOnlyRename::Supported: snapshot.caseOnlyRename = L"supported"; break;
        case FileSystemPathCaseOnlyRename::NoOp: snapshot.caseOnlyRename = L"noOp"; break;
        case FileSystemPathCaseOnlyRename::Unsupported: snapshot.caseOnlyRename = L"unsupported"; break;
        case FileSystemPathCaseOnlyRename::NotApplicable: snapshot.caseOnlyRename = L"notApplicable"; break;
    }

    const auto peerAllowed = [&](FileSystemOperation operation, FileSystemTransferPeerRole role) noexcept
    {
        const FileSystemRouteContract::BooleanResult peer =
            FileSystemRouteContract::QueryTransferPeerAllowed(fs, L"/", operation, role, L"selftest/peer");
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

[[nodiscard]] bool WriteFileBytesFsIo(const wil::com_ptr<IFileSystemIO>& io,
                                      const std::filesystem::path& path,
                                      const void* data,
                                      size_t sizeBytes,
                                      bool allowOverwrite = false) noexcept
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
    const std::wstring driveRootName = localPath.root_name().native();
    if (driveRootName.size() != 2u || driveRootName[1] != L':')
    {
        return {};
    }
    return std::filesystem::path(std::format(LR"(\\localhost\{}$\)", driveRootName[0])) / localPath.lexically_relative(localPath.root_path());
}

[[nodiscard]] bool LoopbackShareReachable(const std::filesystem::path& sandboxRoot) noexcept
{
    const std::filesystem::path alias = LoopbackShareAlias(sandboxRoot);
    return ! alias.empty() && GetFileAttributesW(alias.c_str()) != INVALID_FILE_ATTRIBUTES;
}

bool RecreateEmptyDirectory(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    constexpr DWORD kRetryDelayMs = 50u;
    const ULONGLONG deadline       = ::GetTickCount64() + SelfTest::ScaleTimeout(6'000u);
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
    size_t streamCount                = 0;
    size_t matchingStreamCount        = 0;
    uint64_t progressCount            = 0;
    uint64_t matchingProgress         = 0;
    uint64_t completedCount           = 0;
    HRESULT lastCompletedStatus       = E_PENDING;
    std::optional<FileSystemItemMutationResult> lastMutationResult;
    uint64_t issueCount               = 0;
    HRESULT lastIssueStatus           = E_PENDING;
    bool cancelAfterProgress          = false;
    bool cancelRequested              = false;
    FileSystemIssueAction issueAction = FileSystemIssueAction::Skip;
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
                                                 [[maybe_unused]] uint64_t currentItemCompletedBytes,
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
                }
            }
            if (cancelAfterProgress)
            {
                cancelRequested = true;
            }
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemItemCompleted([[maybe_unused]] FileSystemOperation operationType,
                                                      [[maybe_unused]] unsigned long itemIndex,
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
        lastMutationResult = FileSystemRouteContract::SnapshotItemMutationResult(mutationResult);
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
        *action = issueAction;
        return S_OK;
    }
};

struct FileOpsRecorderOperationControl final : IFileSystemOperationControl
{
    explicit FileOpsRecorderOperationControl(FileOpsRecursiveProgressRecorder& recorder) noexcept : _recorder(recorder) {}

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

    HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress* progress,
                                                                 [[maybe_unused]] void* cookie) noexcept override
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
        h.reset(CreateFileW(
            extendedPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
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

        const std::optional<std::filesystem::path> authorizedRoot = Common::Testing::GetDedicatedExternalTestSandboxBase(
            testRoot,
            Common::Testing::kTestSandboxDirectoryName);
        if (suiteRoot.empty() || scratchRoot.filename() != L"scratch" || runsRoot.filename() != L"runs" ||
            ! authorizedRoot.has_value() || ! Common::Testing::TestSandboxPathEquals(testRoot, authorizedRoot.value()) ||
            ! Common::Testing::HasValidTestSandboxMarker(testRoot))
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
                                                                                                std::wstring& detail) noexcept
{
    detail.clear();

    std::array<wchar_t, 2> alternateVolumeAuthorization{};
    if (GetEnvironmentVariableW(L"REDSALAMANDER_TEST_ALLOW_ALTERNATE_VOLUME",
                                alternateVolumeAuthorization.data(),
                                static_cast<DWORD>(alternateVolumeAuthorization.size())) != 1u ||
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
        const std::filesystem::path probe = candidate / L"probe.tmp";
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
        const std::optional<std::filesystem::path> authorizedRoot = Common::Testing::GetDedicatedExternalTestSandboxBase(
            configuredRoot,
            Common::Testing::kTestSandboxDirectoryName);
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
    std::function<HRESULT()> preConsumptionDecisionGate                = {}) noexcept
{
    if (! fileOps)
    {
        return std::nullopt;
    }

    uint64_t startedTaskId = 0;
    const HRESULT hrStart = fileOps->AdmitOperation(operation,
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
    if (FAILED(hrStart) || startedTaskId == 0)
    {
        return std::nullopt;
    }
    return startedTaskId;
}

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
    state.autoDismissSuccessOriginal          = false;
    state.fileOperationsBackedUp              = false;
    state.fileOperationsOriginal.reset();
    state.appThemeBackedUp = false;
    state.appThemeOriginal.reset();
    state.copyTaskStartTick                 = 0;
    state.localBandwidthRunStartTick        = 0;
    state.localBandwidthCancelStartTick     = 0;
    state.localBandwidthDurationUs          = 0;
    state.localBandwidthDurationLeadUs      = 0;
    state.localBandwidthCancelLatencyUs     = 0;
    state.localBandwidthMaxWindowBytes      = 0;
    state.localBandwidthMaxSampleDeltaBytes = 0;
    state.localBandwidthSamples.clear();
    state.bandwidthThrottleWorkerModeEnvBackedUp    = false;
    state.bandwidthThrottleWorkerModeEnvHadOriginal = false;
    state.bandwidthThrottleWorkerModeEnvOriginal.clear();
    state.parallelBandwidthRunStartTick          = 0;
    state.parallelBandwidthBaselineUs            = 0;
    state.parallelBandwidthCandidateUs           = 0;
    state.parallelBandwidthBaselineMaxSkewBytes  = 0;
    state.parallelBandwidthCandidateMaxSkewBytes = 0;
    state.parallelBandwidthBaselineMaxCallbackDeltaBytes  = 0;
    state.parallelBandwidthCandidateMaxCallbackDeltaBytes = 0;
    state.parallelBandwidthBaselineMaxActive     = 0;
    state.parallelBandwidthCandidateMaxActive    = 0;
    state.parallelBandwidthBaselineSamples       = 0;
    state.parallelBandwidthCandidateSamples      = 0;
    state.defaultSpeedLimitRunStartTick          = 0;
    state.defaultSpeedLimitBaselineUs            = 0;
    state.defaultSpeedLimitCandidateUs           = 0;
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
    state.activePhaseOrder    = selection.activePhases;
    state.reportedPhaseOrder  = selection.reportedPhases;
    state.step                = SelfTestState::Step::Setup;
    state.runStartTick        = GetTickCount64();
    state.stepStartTick       = static_cast<ULONGLONG>(state.runStartTick);
    state.markerTick          = 0;
    state.baselineThreadCount = 0;
    BeginPhase(state, SelfTestState::Step::Setup);
    AppendLog(L"Start: setup ready");
    AppendLog(L"Start");
    Debug::Info(L"FileOpsSelfTest: started");
}

namespace
{
using TickResult = std::optional<bool>;

#pragma warning(push)
#pragma warning(disable : 4061) // Each helper intentionally owns only its phase subset; unmatched steps return nullopt.
#pragma warning(disable : 4883) // This test-only dispatcher includes the intentionally large Fairstream phase corpus; runtime optimization is irrelevant.

[[nodiscard]] TickResult TickFairstreamA(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_A
#include "FolderWindow.FileOperations.SelfTest.Fairstream.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_A
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickFairstreamB(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_B
#include "FolderWindow.FileOperations.SelfTest.Fairstream.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_B
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickFairstreamC(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_C
#include "FolderWindow.FileOperations.SelfTest.Fairstream.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_C
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickCopyMerge(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_COPY_MERGE
#include "FolderWindow.FileOperations.SelfTest.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_COPY_MERGE
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickMoveMerge(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_MOVE_MERGE
#include "FolderWindow.FileOperations.SelfTest.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_MOVE_MERGE
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickReparseMerge(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_REPARSE_MERGE
#include "FolderWindow.FileOperations.SelfTest.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_REPARSE_MERGE
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickProviderMatrix(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PROVIDER_MATRIX
#include "FolderWindow.FileOperations.SelfTest.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PROVIDER_MATRIX
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase5(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE5
#include "FolderWindow.FileOperations.SelfTest.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE5
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase6(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE6
#include "FolderWindow.FileOperations.SelfTest.Phases05_06.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE6
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase7(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE7
#include "FolderWindow.FileOperations.SelfTest.Phases07_09.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE7
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase8(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE8
#include "FolderWindow.FileOperations.SelfTest.Phases07_09.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE8
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase9(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE9
#include "FolderWindow.FileOperations.SelfTest.Phases07_09.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE9
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase10(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE10
#include "FolderWindow.FileOperations.SelfTest.Phases10_13.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE10
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase11(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE11
#include "FolderWindow.FileOperations.SelfTest.Phases10_13.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE11
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase12(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE12
#include "FolderWindow.FileOperations.SelfTest.Phases10_13.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE12
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhase13(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#define FILEOPS_SELFTEST_INCLUDE_PHASE13
#include "FolderWindow.FileOperations.SelfTest.Phases10_13.cpp"
#undef FILEOPS_SELFTEST_INCLUDE_PHASE13
        default: return std::nullopt;
    }
}

[[nodiscard]] TickResult TickPhases14To16(SelfTestState& state) noexcept
{
    switch (state.step)
    {
#include "FolderWindow.FileOperations.SelfTest.Phases14_16.cpp"
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
    const auto dispatchPhase = [&](const auto& tickPhase) noexcept -> TickResult
    {
        return tickPhase(state);
    };
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
            state.autoDismissSuccessOriginal = state.fileOps->GetAutoDismissSuccess();
            state.fileOperationsOriginal     = g_settings.fileOperations;
            state.fileOperationsBackedUp     = true;
            state.appThemeOriginal           = state.folderWindow->GetTheme();
            state.appThemeBackedUp           = true;
            AppTheme deterministicMotionTheme = state.appThemeOriginal.value();
            deterministicMotionTheme.reducedMotionOverride = false;
            state.folderWindow->ApplyTheme(deterministicMotionTheme);
            Common::Settings::FileOperationsSettings& fileOperations = EnsureFileOperationsSettingsForSelfTest();
            fileOperations.crossFsBridgeBufferSizeKB                    = 4096u;
            fileOperations.defaultBandwidthLimitBytesPerSecond          = 0;

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

                const std::filesystem::path src                     = state.tempRoot / L"copy-src";
                const std::filesystem::path dst                     = state.tempRoot / L"copy-dst";
                const std::filesystem::path del                     = state.tempRoot / L"delete-tree";
                const std::filesystem::path en                      = state.tempRoot / L"enum";
                const std::filesystem::path watch                   = state.tempRoot / L"watch";
                const std::filesystem::path preA                    = state.tempRoot / L"discovery-a";
                const std::filesystem::path preB                    = state.tempRoot / L"discovery-b";
                const bool needsDeleteTree                          = IsPhaseSelected(state, SelfTestState::Step::Phase6_DeleteBytesMeaningful);
                const bool needsDiscoverySwitchTrees                  = IsPhaseSelected(state, SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery) ||
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
            NextStep(state, SelfTestState::Step::Phase5_DiscoverySingleTraversal);
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
                info.discoveryClosed               = task->_discoveryClosed.load(std::memory_order_acquire);
                info.firstMutationBeforeDiscoveryClosed = task->_firstMutationBeforeDiscoveryClosed.load(std::memory_order_acquire);
                info.discoverySkipped                 = task->_discoverySkipped.load(std::memory_order_acquire);
                info.discoveredTotalBytes              = task->_discoveredTotalBytes.load(std::memory_order_acquire);
                info.discoveredFiles                   = task->_discoveredFileCount.load(std::memory_order_acquire);
                info.discoveredDirectories             = task->_discoveredDirectoryCount.load(std::memory_order_acquire);
                info.discoveryDurationUs              = task->_perf.discoveryOpenUs.load(std::memory_order_acquire);
                info.discoveryMaxQueueDepth         = task->_discoveryMaxQueueDepth.load(std::memory_order_acquire);
                info.discoveryStarvationCount       = task->_discoveryStarvationCount.load(std::memory_order_acquire);
                info.discoveryCallbackCount         = task->_perf.discoveryCallbackCount.load(std::memory_order_acquire);
                info.discoveryCallbackUs            = task->_perf.discoveryCallbackUs.load(std::memory_order_acquire);
                info.discoveryLockWaitUs            = task->_perf.discoveryLockWaitUs.load(std::memory_order_acquire);
                info.discoveryFirstMutationUs = task->_discoveryFirstMutationUs.load(std::memory_order_acquire);
                info.discoveryBytesCompletedWhileOpen = task->_discoveryCompletedBytesWhileOpen.load(std::memory_order_acquire);
                info.discoveryMutationsCompletedWhileOpen = task->_discoveryCompletedMutationsWhileOpen.load(std::memory_order_acquire);
                info.progressCallbackCount          = task->_progressCallbackCount.load(std::memory_order_acquire);
                info.started                        = task->HasStarted();
                info.conflictWaitUs                 = task->_perf.conflictWaitUs;
                info.conflictConvergenceWaitUs      = task->_perf.conflictConvergenceWaitUs;
                info.conflictPromptCount            = task->_perf.conflictPromptCount;
                info.interlockWaitUs                 = task->_perf.interlockWaitUs;
                info.interlockWaitCount              = task->_perf.interlockWaitCount;
                info.bridgeDirectoryEnsureCount     = task->_bridgeDirectoryEnsureCount.load(std::memory_order_acquire);
                info.bridgeSourceDirectoryEnumerationCount = task->_bridgeSourceDirectoryEnumerationCount.load(std::memory_order_acquire);
                info.bridgeFileAdmissionCount       = task->_bridgeFileAdmissionCount.load(std::memory_order_acquire);
                info.bridgeEarlyFileStartCount      = task->_bridgeFileStartedBeforeProducerDone.load(std::memory_order_acquire);
                info.bridgeAdmissionMaxQueueDepth   = task->_bridgeAdmissionMaxQueueDepth.load(std::memory_order_acquire);
                info.bridgeTraversalMaxDepth = task->_bridgeTraversalMaxDepth.load(std::memory_order_acquire);
                info.bridgeTraversalMaxRetainedEntries = task->_bridgeTraversalMaxRetainedEntries.load(std::memory_order_acquire);
                info.bridgeTraversalMaxQueuedPathBytes = task->_bridgeTraversalMaxQueuedPathBytes.load(std::memory_order_acquire);
                info.bridgeTraversalMaxMetadataBytes = task->_bridgeTraversalMaxMetadataBytes.load(std::memory_order_acquire);
                info.bridgeTraversalLimitHitCount = task->_bridgeTraversalLimitHitCount.load(std::memory_order_acquire);
                info.conflictExpectedDestinationBoundCount =
                    task->_conflictExpectedDestinationBoundCount.load(std::memory_order_acquire);
                info.conflictExpectedDestinationUnavailableCount =
                    task->_conflictExpectedDestinationUnavailableCount.load(std::memory_order_acquire);
                info.conflictExpectedDestinationReturnedCount =
                    task->_conflictExpectedDestinationReturnedCount.load(std::memory_order_acquire);
                info.bridgeImmediateDestinationSizeProbeUs =
                    task->_perf.bridgeImmediateDestinationSizeProbeUs.load(std::memory_order_acquire);
                info.bridgeImmediateDestinationSizeProbeCount =
                    task->_perf.bridgeImmediateDestinationSizeProbeCount.load(std::memory_order_acquire);
                info.bridgeCommitSizeProofCount = task->_perf.bridgeCommitSizeProofCount.load(std::memory_order_acquire);
                info.bridgeCommitSizeProofFallbackCount = task->_perf.bridgeCommitSizeProofFallbackCount.load(std::memory_order_acquire);
                info.verificationUs = task->_perf.verificationUs.load(std::memory_order_acquire);
                info.verificationReadBytes = task->_perf.verificationReadBytes.load(std::memory_order_acquire);
                info.verificationReadCalls = task->_perf.verificationReadCalls.load(std::memory_order_acquire);
                info.verificationProviderProofCount = task->_perf.verificationProviderProofCount.load(std::memory_order_acquire);
                info.verificationHostReadbackCount = task->_perf.verificationHostReadbackCount.load(std::memory_order_acquire);
                info.verificationTotalBytes = task->_verificationTotalBytes.load(std::memory_order_acquire);
                info.verificationCompletedBytes = task->_verificationCompletedBytes.load(std::memory_order_acquire);
                info.configuredMaxConcurrency       = task->_effectiveConcurrencyBudget.load(std::memory_order_acquire);
                if (info.configuredMaxConcurrency == 0u)
                {
                    info.configuredMaxConcurrency = task->_perItemMaxConcurrencyBudget;
                }
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

            const std::wstring_view reason =
                state.failed.load(std::memory_order_acquire) ? L"not reached (aborted due to failure)" : L"not reached";
            SelfTest::AppendCaseResult(result, expected, SelfTest::SelfTestCaseResult::Status::skipped, reason, 0);
        }

        return result;
    }

    bool FileOperationsSelfTest::DidFail() noexcept
    {
        return GetState().failed.load(std::memory_order_acquire);
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
