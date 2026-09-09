#include "FolderWindow.FileOperations.State.Private.h"
#include "FolderWindow.FileOperationsInternal.h"

#include "Blake3Digest.h"
#include "ConnectionProfileUtils.h"
#include "ContentDigest.h"
#include "FileOperationTraversalPolicy.h"
#include "FileSystemPathIdentity.h"
#include "FileSystemRouteContract.h"
#include "FolderWindow.FileOperations.IssuesPane.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "PathUtils.h"
#include "Resource.h"
#include "SessionState.h"
#include "SettingsHotReload.h"
#include "SettingsSave.h"
#include "SettingsStore.h"
#include "SynchronousIoCancelWatch.h"

#include <algorithm>
#include <array>
#include <bcrypt.h>
#include <cassert>
#include <chrono>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <deque>
#include <functional>
#include <iterator>
#include <psapi.h>
#include <set>

#pragma comment(lib, "bcrypt.lib")
#include <shellapi.h>
#include <system_error>
#include <thread>
#include <tuple>
#include <unordered_map>
#include <utility>

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

namespace FolderWindowFileOperationsStateInternal
{
using Task = FolderWindow::FileOperationState::Task;

using FileSystemRouteContract::IsValidItemMutationResultPrefix;

[[nodiscard]] FileOperations::OwnedStageDisposition GetOwnedStageDisposition(const FileSystemItemMutationResult& result) noexcept
{
    if (! FileSystemItemMutationResultHasOwnedStageDisposition(result))
    {
        return FileOperations::OwnedStageDisposition::NotApplicable;
    }
    switch (result.ownedStageDisposition)
    {
        case FileSystemOwnedStageDisposition::NotCreated: return FileOperations::OwnedStageDisposition::NotCreated;
        case FileSystemOwnedStageDisposition::Removed: return FileOperations::OwnedStageDisposition::Removed;
        case FileSystemOwnedStageDisposition::Published: return FileOperations::OwnedStageDisposition::Published;
        case FileSystemOwnedStageDisposition::Retained: return FileOperations::OwnedStageDisposition::Retained;
        case FileSystemOwnedStageDisposition::Unknown: return FileOperations::OwnedStageDisposition::Unknown;
        case FileSystemOwnedStageDisposition::RetainedIncomplete: return FileOperations::OwnedStageDisposition::RetainedIncomplete;
        case FileSystemOwnedStageDisposition::NotApplicable:
        default: return FileOperations::OwnedStageDisposition::NotApplicable;
    }
}

[[nodiscard]] bool OwnedStageDispositionIsIndeterminate(FileOperations::OwnedStageDisposition disposition) noexcept
{
    return disposition == FileOperations::OwnedStageDisposition::Unknown || disposition == FileOperations::OwnedStageDisposition::RetainedIncomplete;
}

[[nodiscard]] HRESULT AdvanceValidatedFileInfoEntry(FileInfo* entry, const std::byte* bufferBase, const std::byte* bufferEnd, FileInfo*& nextOut) noexcept;
[[nodiscard]] HRESULT TryGetValidatedFileInfoName(FileInfo* entry,
                                                  const std::byte* bufferBase,
                                                  const std::byte* bufferEnd,
                                                  std::wstring_view& nameOut) noexcept;

#ifdef ENABLE_TESTS
std::atomic<unsigned int> g_fileOpsBridgePipelineMode{static_cast<unsigned int>(FileOpsBridgePipelineMode::Default)};
std::atomic<unsigned int> g_fileOpsBridgeProducerDelayMs{0};
std::atomic<uint64_t> g_fileOpsBridgeTraversalDepthLimitOverride{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextFileCopyCount{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextFileCopyAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextSourceGetSizeCount{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextSourceGetSizeAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationGetSizeCount{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationGetSizeAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationOpenCount{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationOpenAttempts{0};
std::atomic<HRESULT> g_fileOpsBridgeFailNextDestinationOpenStatus{HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)};
std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationBasicInfoCount{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationBasicInfoAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeNullNextSourceReaderCount{0};
std::atomic<unsigned long> g_fileOpsBridgeNullNextSourceReaderAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeReportWrongDestinationSizeCount{0};
std::atomic<unsigned long> g_fileOpsBridgeReportWrongDestinationSizeAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextStageEntropyCount{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextStageEntropyAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeOverReportNextReadCount{0};
std::atomic<unsigned long> g_fileOpsBridgeOverReportNextReadAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgePrematureEofNextReadCount{0};
std::atomic<unsigned long> g_fileOpsBridgePrematureEofNextReadAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeUnderConsumeNextWriteCount{0};
std::atomic<unsigned long> g_fileOpsBridgeUnderConsumeNextWriteAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeOverReportNextWriteCount{0};
std::atomic<unsigned long> g_fileOpsBridgeOverReportNextWriteAttempts{0};
std::atomic<bool> g_fileOpsBridgeInjectHostileChildNames{false};
std::atomic<unsigned long> g_fileOpsBridgeInjectHostileChildNameAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeInjectFileReparseCount{0};
std::atomic<unsigned long> g_fileOpsBridgeInjectFileReparseAttempts{0};
std::atomic<int> g_fileOpsBridgeReparsePolicyOverride{-1};
std::atomic<unsigned long> g_fileOpsBridgeReplacePublishedDestinationAttempts{0};
std::atomic<unsigned long> g_fileOpsManagedCleanupKnownNonCommitCount{0};
std::atomic<unsigned long> g_fileOpsManagedCleanupKnownNonCommitAttempts{0};
std::atomic<HRESULT> g_fileOpsManagedCleanupKnownNonCommitStatus{HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION)};
std::atomic<unsigned long> g_fileOpsPermanentDeleteKnownNonCommitCount{0};
std::atomic<unsigned long> g_fileOpsPermanentDeleteKnownNonCommitAttempts{0};
std::atomic<HRESULT> g_fileOpsPermanentDeleteKnownNonCommitStatus{HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION)};
std::atomic<unsigned long> g_fileOpsManagedCleanupUnknownOutcomeCount{0};
std::atomic<unsigned long> g_fileOpsManagedCleanupUnknownOutcomeAttempts{0};
std::atomic<unsigned long> g_fileOpsVerificationForceHostReadbackCount{0};
std::atomic<unsigned long> g_fileOpsVerificationForceHostReadbackAttempts{0};
std::atomic<unsigned long> g_fileOpsVerificationForceUnavailableCount{0};
std::atomic<unsigned long> g_fileOpsVerificationForceUnavailableAttempts{0};
std::atomic<unsigned long> g_fileOpsVerificationForceMismatchCount{0};
std::atomic<unsigned long> g_fileOpsVerificationForceMismatchAttempts{0};
std::atomic<bool> g_fileOpsAutoConcurrencyOverrideEnabled{false};
std::atomic<unsigned int> g_fileOpsAutoConcurrencyOverridePreferred{1};
std::atomic<uint32_t> g_fileOpsAutoConcurrencyOverrideStorageKind{FILESYSTEM_STORAGE_UNKNOWN};

SelfTestPausePoint g_fileOpsPostFinishedCompletionPausePoint;
SelfTestPausePoint g_fileOpsBridgeMoveSourceCleanupPausePoint;
SelfTestPausePoint g_fileOpsBridgePublishedDestinationRetryPausePoint;
SelfTestPausePoint g_fileOpsConflictMetadataPausePoint;
SelfTestPausePoint g_fileOpsKeepBothNestedConflictPausePoint;
SelfTestPausePoint g_fileOpsInlineRenameBeforeMutationPausePoint;
SelfTestPausePoint g_fileOpsBatchRenameBeforeExecutionPausePoint;
SelfTestPausePoint g_fileOpsRecycleEscalationBeforeBindPausePoint;
SelfTestPausePoint g_fileOpsPermanentDeleteBeforeBindPausePoint;
SelfTestPausePoint g_fileOpsPermanentDeleteBeforeRecheckPausePoint;
SelfTestPausePoint g_fileOpsPermanentDeleteBeforeLiveOutputGuardPausePoint;
SelfTestPausePoint g_fileOpsLiveOutputPublishedPausePoint;
SelfTestPausePoint g_fileOpsVerificationReadbackPausePoint;
std::atomic<unsigned long> g_fileOpsNativeMoveCreateDirectoryRaceCount{0u};
std::atomic<unsigned long> g_fileOpsNativeMoveCreateDirectoryRaceAttempts{0u};
std::atomic<DWORD> g_fileOpsInlineRenameAdmissionThreadId{0u};
std::atomic<DWORD> g_fileOpsInlineRenameExecutionThreadId{0u};
std::atomic<unsigned long> g_fileOpsInlineRenameExecutionAttempts{0u};
std::atomic<ULONGLONG> g_fileOpsConflictMetadataPauseBailoutMs{5'000ull};

[[nodiscard]] unsigned long GetInFlightFileCountSnapshot(Task& task) noexcept
{
    std::scoped_lock lock(task._inFlightFilesMutex);
    return static_cast<unsigned long>(task._inFlightFileCount);
}
#endif

[[nodiscard]] uint64_t PerfNowUs() noexcept
{
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now().time_since_epoch()).count());
}

[[nodiscard]] uint64_t PerfElapsedUs(uint64_t startUs) noexcept
{
    const uint64_t nowUs = PerfNowUs();
    return (nowUs >= startUs) ? (nowUs - startUs) : 0;
}

void AtomicMax(std::atomic<uint64_t>& target, uint64_t value) noexcept
{
    uint64_t current = target.load(std::memory_order_acquire);
    while (current < value && ! target.compare_exchange_weak(current, value, std::memory_order_acq_rel, std::memory_order_acquire))
    {
    }
}

[[nodiscard]] std::optional<std::wstring> TryGetUncShareRootBoundary(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return std::nullopt;
    }

    std::wstring normalized(text);
    std::ranges::replace(normalized, L'/', L'\\');

    size_t serverStart = 0u;
    if (normalized.rfind(L"\\\\?\\UNC\\", 0u) == 0u)
    {
        serverStart = 8u;
    }
    else if (normalized.rfind(L"\\\\", 0u) == 0u && normalized.rfind(L"\\\\?\\", 0u) != 0u)
    {
        serverStart = 2u;
    }
    else
    {
        return std::nullopt;
    }

    const size_t serverEnd = normalized.find(L'\\', serverStart);
    if (serverEnd == std::wstring::npos || serverEnd == serverStart)
    {
        return std::nullopt;
    }

    const size_t shareStart = serverEnd + 1u;
    if (shareStart >= normalized.size() || normalized[shareStart] == L'\\')
    {
        return std::nullopt;
    }

    size_t shareEnd = normalized.find(L'\\', shareStart);
    if (shareEnd == std::wstring::npos)
    {
        shareEnd = normalized.size();
    }
    if (shareEnd <= shareStart)
    {
        return std::nullopt;
    }

    return normalized.substr(0u, shareEnd);
}

[[nodiscard]] bool IsUncShareRootBoundary(std::wstring_view path, const std::optional<std::wstring>& shareRoot) noexcept
{
    if (! shareRoot.has_value())
    {
        return false;
    }

    std::wstring normalized(path);
    for (wchar_t& ch : normalized)
    {
        if (ch == L'/')
        {
            ch = L'\\';
        }
    }
    while (normalized.size() > shareRoot.value().size() && normalized.back() == L'\\')
    {
        normalized.pop_back();
    }

    return OrdinalString::EqualsNoCase(normalized, shareRoot.value());
}

#ifdef ENABLE_TESTS
[[nodiscard]] FileOpsBridgePipelineMode GetBridgePipelineModeOverride() noexcept
{
    const unsigned int raw = g_fileOpsBridgePipelineMode.load(std::memory_order_acquire);
    switch (static_cast<FileOpsBridgePipelineMode>(raw))
    {
        case FileOpsBridgePipelineMode::Default: return FileOpsBridgePipelineMode::Default;
        case FileOpsBridgePipelineMode::Disabled: return FileOpsBridgePipelineMode::Disabled;
        case FileOpsBridgePipelineMode::Enabled: return FileOpsBridgePipelineMode::Enabled;
        default: return FileOpsBridgePipelineMode::Default;
    }
}

[[nodiscard]] unsigned int GetBridgeProducerDelayMsForSelfTest() noexcept
{
    return g_fileOpsBridgeProducerDelayMs.load(std::memory_order_acquire);
}

[[nodiscard]] std::optional<std::wstring> TryReadEnvironmentVariableForSelfTest(const wchar_t* name) noexcept
{
    if (name == nullptr || name[0] == L'\0')
    {
        return std::nullopt;
    }

    const DWORD required = GetEnvironmentVariableW(name, nullptr, 0u);
    if (required == 0u)
    {
        return std::nullopt;
    }

    std::wstring value(required, L'\0');
    const DWORD written = GetEnvironmentVariableW(name, value.data(), required);
    if (written == 0u || written >= required)
    {
        return std::nullopt;
    }

    value.resize(written);
    return value;
}

[[nodiscard]] std::string NarrowEnvironmentPayloadForSelfTest(std::wstring_view payload) noexcept
{
    std::string bytes;
    bytes.reserve(payload.size());
    for (const wchar_t ch : payload)
    {
        if (ch > 0xFF)
        {
            return {};
        }
        bytes.push_back(static_cast<char>(ch));
    }
    return bytes;
}

[[nodiscard]] bool ConsumeBridgeFailNextFileCopyForSelfTest() noexcept
{
    unsigned long remaining = g_fileOpsBridgeFailNextFileCopyCount.load(std::memory_order_acquire);
    while (remaining > 0u)
    {
        if (g_fileOpsBridgeFailNextFileCopyCount.compare_exchange_weak(remaining, remaining - 1u, std::memory_order_acq_rel, std::memory_order_acquire))
        {
            g_fileOpsBridgeFailNextFileCopyAttempts.fetch_add(1u, std::memory_order_acq_rel);
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool ConsumeBridgeFailNextSourceGetSizeForSelfTest() noexcept
{
    unsigned long remaining = g_fileOpsBridgeFailNextSourceGetSizeCount.load(std::memory_order_acquire);
    while (remaining > 0u)
    {
        if (g_fileOpsBridgeFailNextSourceGetSizeCount.compare_exchange_weak(remaining, remaining - 1u, std::memory_order_acq_rel, std::memory_order_acquire))
        {
            g_fileOpsBridgeFailNextSourceGetSizeAttempts.fetch_add(1u, std::memory_order_acq_rel);
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool ConsumeBridgeCounterForSelfTest(std::atomic<unsigned long>& remainingCount, std::atomic<unsigned long>& attemptCount) noexcept
{
    unsigned long remaining = remainingCount.load(std::memory_order_acquire);
    while (remaining > 0u)
    {
        if (remainingCount.compare_exchange_weak(remaining, remaining - 1u, std::memory_order_acq_rel, std::memory_order_acquire))
        {
            attemptCount.fetch_add(1u, std::memory_order_acq_rel);
            return true;
        }
    }

    return false;
}

[[nodiscard]] HRESULT MaybeInjectHostileBridgeChildNamesForSelfTest(FileInfo* head, std::byte* bufferBase, std::byte* bufferEnd) noexcept
{
    if (! g_fileOpsBridgeInjectHostileChildNames.exchange(false, std::memory_order_acq_rel))
    {
        return S_FALSE;
    }
    g_fileOpsBridgeInjectHostileChildNameAttempts.fetch_add(1u, std::memory_order_acq_rel);

    static constexpr wchar_t kEmbeddedNulName[] = {L'n', L'u', L'l', L'\0', L't', L'a', L'i', L'l', L'.', L't', L'x', L't'};
    static constexpr wchar_t kControlName[]     = {L'c', L't', L'l', static_cast<wchar_t>(1), L'.', L't', L'x', L't'};
    constexpr std::array<std::wstring_view, 8> kHostileNames{{
        L"..\\escape.txt",
        L"a\\b.txt",
        std::wstring_view(kEmbeddedNulName, std::size(kEmbeddedNulName)),
        std::wstring_view(kControlName, std::size(kControlName)),
        L"x:stream",
        L"CON",
        L"Case.txt",
        L"case.txt",
    }};

    FileInfo* entry = head;
    for (const std::wstring_view hostileName : kHostileNames)
    {
        if (entry == nullptr)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const auto* entryBytes       = reinterpret_cast<const std::byte*>(entry);
        const size_t available       = static_cast<size_t>(bufferEnd - entryBytes);
        const size_t recordBytes     = entry->NextEntryOffset != 0u ? static_cast<size_t>(entry->NextEntryOffset) : available;
        constexpr size_t kNameOffset = offsetof(FileInfo, FileName);
        const size_t hostileBytes    = hostileName.size() * sizeof(wchar_t);
        if (recordBytes < kNameOffset || hostileBytes > recordBytes - kNameOffset)
        {
            return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
        }

        std::memcpy(entry->FileName, hostileName.data(), hostileBytes);
        entry->FileNameSize = static_cast<unsigned long>(hostileBytes);

        FileInfo* next          = nullptr;
        const HRESULT advanceHr = AdvanceValidatedFileInfoEntry(entry, bufferBase, bufferEnd, next);
        if (advanceHr == S_FALSE)
        {
            entry = nullptr;
        }
        else if (FAILED(advanceHr))
        {
            return advanceHr;
        }
        else
        {
            entry = next;
        }
    }
    return S_OK;
}

[[nodiscard]] HRESULT MaybeInjectBridgeFileReparseForSelfTest(FileInfo* head, std::byte* bufferBase, std::byte* bufferEnd) noexcept
{
    if (! ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeInjectFileReparseCount, g_fileOpsBridgeInjectFileReparseAttempts))
    {
        return S_FALSE;
    }

    constexpr std::wstring_view kInjectedFileName = L"reparse-file.bin";
    FileInfo* entry                               = head;
    while (entry != nullptr)
    {
        std::wstring_view name;
        HRESULT hr = TryGetValidatedFileInfoName(entry, bufferBase, bufferEnd, name);
        if (FAILED(hr))
        {
            return hr;
        }
        if (name == kInjectedFileName && (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
        {
            entry->FileAttributes |= FILE_ATTRIBUTE_REPARSE_POINT;
            return S_OK;
        }

        FileInfo* next = nullptr;
        hr             = AdvanceValidatedFileInfoEntry(entry, bufferBase, bufferEnd, next);
        if (hr == S_FALSE)
        {
            break;
        }
        if (FAILED(hr))
        {
            return hr;
        }
        entry = next;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

enum class SelfTestBridgeIoRole : unsigned char
{
    Source,
    Destination,
};

struct SelfTestBridgeFileReader final : IFileReader, IFileReaderOperationControl
{
    SelfTestBridgeFileReader(wil::com_ptr<IFileReader> inner, SelfTestBridgeIoRole role) noexcept : _inner(std::move(inner)), _role(role)
    {
        if (_inner)
        {
            static_cast<void>(_inner->QueryInterface(IID_PPV_ARGS(_operationControl.addressof())));
        }
    }
    SelfTestBridgeFileReader(const SelfTestBridgeFileReader&)            = delete;
    SelfTestBridgeFileReader& operator=(const SelfTestBridgeFileReader&) = delete;
    SelfTestBridgeFileReader(SelfTestBridgeFileReader&&)                 = delete;
    SelfTestBridgeFileReader& operator=(SelfTestBridgeFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileReaderOperationControl) && _operationControl)
        {
            *ppvObject = static_cast<IFileReaderOperationControl*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (current == 0u)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE SetOperationControl(const FileSystemOptions* options) noexcept override
    {
        return _operationControl ? _operationControl->SetOperationControl(options) : E_NOINTERFACE;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (sizeBytes == nullptr)
        {
            return E_POINTER;
        }
        if (! _inner)
        {
            return E_POINTER;
        }
        if (_role == SelfTestBridgeIoRole::Source && ConsumeBridgeFailNextSourceGetSizeForSelfTest())
        {
            return HRESULT_FROM_WIN32(ERROR_READ_FAULT);
        }
        if (_role == SelfTestBridgeIoRole::Destination &&
            ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeFailNextDestinationGetSizeCount, g_fileOpsBridgeFailNextDestinationGetSizeAttempts))
        {
            return HRESULT_FROM_WIN32(ERROR_READ_FAULT);
        }
        const HRESULT hr = _inner->GetSize(sizeBytes);
        if (SUCCEEDED(hr) && _role == SelfTestBridgeIoRole::Destination &&
            ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeReportWrongDestinationSizeCount, g_fileOpsBridgeReportWrongDestinationSizeAttempts))
        {
            *sizeBytes = *sizeBytes == std::numeric_limits<uint64_t>::max() ? *sizeBytes - 1u : *sizeBytes + 1u;
        }
        return hr;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        return _inner ? _inner->Seek(offset, origin, newPosition) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (bytesRead == nullptr)
        {
            return E_POINTER;
        }
        if (! _inner)
        {
            return E_POINTER;
        }

        if (_role == SelfTestBridgeIoRole::Source &&
            ConsumeBridgeCounterForSelfTest(g_fileOpsBridgePrematureEofNextReadCount, g_fileOpsBridgePrematureEofNextReadAttempts))
        {
            *bytesRead = 0;
            return S_OK;
        }

        const HRESULT hr = _inner->Read(buffer, bytesToRead, bytesRead);
        if (SUCCEEDED(hr) && _role == SelfTestBridgeIoRole::Source && bytesToRead < std::numeric_limits<unsigned long>::max() &&
            ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeOverReportNextReadCount, g_fileOpsBridgeOverReportNextReadAttempts))
        {
            *bytesRead = bytesToRead + 1u;
        }
        return hr;
    }

private:
    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileReader> _inner;
    wil::com_ptr<IFileReaderOperationControl> _operationControl; // R0f-Curl-OR1
    SelfTestBridgeIoRole _role = SelfTestBridgeIoRole::Source;
};

struct SelfTestBridgeFileWriter final : IFileWriter,
                                        IFileWriterExpectedSize,
                                        IFileWriterCommitSizeProof,
                                        IFileWriterExpectedReplacement,
                                        IFileWriterContentProof
{
    explicit SelfTestBridgeFileWriter(wil::com_ptr<IFileWriter> inner) noexcept : _inner(std::move(inner))
    {
        if (_inner)
        {
            static_cast<void>(_inner->QueryInterface(IID_PPV_ARGS(_commitSizeProof.addressof())));
            static_cast<void>(_inner->QueryInterface(IID_PPV_ARGS(_expectedSizeWriter.addressof())));
            static_cast<void>(_inner->QueryInterface(IID_PPV_ARGS(_expectedReplacement.addressof())));
            static_cast<void>(_inner->QueryInterface(IID_PPV_ARGS(_contentProof.addressof())));
        }
    }
    SelfTestBridgeFileWriter(const SelfTestBridgeFileWriter&)            = delete;
    SelfTestBridgeFileWriter& operator=(const SelfTestBridgeFileWriter&) = delete;
    SelfTestBridgeFileWriter(SelfTestBridgeFileWriter&&)                 = delete;
    SelfTestBridgeFileWriter& operator=(SelfTestBridgeFileWriter&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileWriter))
        {
            *ppvObject = static_cast<IFileWriter*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileWriterCommitSizeProof) && _commitSizeProof)
        {
            *ppvObject = static_cast<IFileWriterCommitSizeProof*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileWriterExpectedSize) && _expectedSizeWriter)
        {
            *ppvObject = static_cast<IFileWriterExpectedSize*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileWriterExpectedReplacement) && _expectedReplacement)
        {
            *ppvObject = static_cast<IFileWriterExpectedReplacement*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileWriterContentProof) && _contentProof)
        {
            *ppvObject = static_cast<IFileWriterContentProof*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (current == 0u)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetPosition(uint64_t* positionBytes) noexcept override
    {
        return _inner ? _inner->GetPosition(positionBytes) : E_POINTER;
    }

    // R3-1: the decorator is transparent for the replace expectation.
    HRESULT STDMETHODCALLTYPE SetExpectedReplacement(const FileSystemBasicInformation* expected) noexcept override
    {
        return _expectedReplacement ? _expectedReplacement->SetExpectedReplacement(expected) : E_NOINTERFACE;
    }

    HRESULT STDMETHODCALLTYPE GetContentProofAlgorithms(uint32_t* algorithmMask) noexcept override
    {
        return _contentProof ? _contentProof->GetContentProofAlgorithms(algorithmMask) : E_NOINTERFACE;
    }

    HRESULT STDMETHODCALLTYPE GetCommittedContentProof(FileSystemContentProof* proof) noexcept override
    {
        return _contentProof ? _contentProof->GetCommittedContentProof(proof) : E_NOINTERFACE;
    }

    HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept override
    {
        if (bytesWritten == nullptr)
        {
            return E_POINTER;
        }
        if (! _inner)
        {
            return E_POINTER;
        }

        if (bytesToWrite > 0u && ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeUnderConsumeNextWriteCount, g_fileOpsBridgeUnderConsumeNextWriteAttempts))
        {
            unsigned long persistedBytes = 0;
            const HRESULT hr             = _inner->Write(buffer, bytesToWrite - 1u, &persistedBytes);
            if (SUCCEEDED(hr))
            {
                *bytesWritten = bytesToWrite;
            }
            return hr;
        }

        const HRESULT hr = _inner->Write(buffer, bytesToWrite, bytesWritten);
        if (SUCCEEDED(hr) && bytesToWrite < std::numeric_limits<unsigned long>::max() &&
            ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeOverReportNextWriteCount, g_fileOpsBridgeOverReportNextWriteAttempts))
        {
            *bytesWritten = bytesToWrite + 1u;
        }
        return hr;
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override
    {
        return _inner ? _inner->Commit() : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetCommittedSize(uint64_t* sizeBytes) noexcept override
    {
        return _commitSizeProof ? _commitSizeProof->GetCommittedSize(sizeBytes) : E_NOINTERFACE;
    }

    HRESULT STDMETHODCALLTYPE SetExpectedSize(uint64_t sizeBytes) noexcept override
    {
        return _expectedSizeWriter ? _expectedSizeWriter->SetExpectedSize(sizeBytes) : E_NOINTERFACE;
    }

private:
    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileWriter> _inner;
    wil::com_ptr<IFileWriterExpectedSize> _expectedSizeWriter;
    wil::com_ptr<IFileWriterExpectedReplacement> _expectedReplacement;
    wil::com_ptr<IFileWriterCommitSizeProof> _commitSizeProof;
    wil::com_ptr<IFileWriterContentProof> _contentProof; // R3-2
};

struct SelfTestBridgeIoDecorator final : IFileSystemIO
{
    SelfTestBridgeIoDecorator(wil::com_ptr<IFileSystemIO> inner, SelfTestBridgeIoRole role) noexcept : _inner(std::move(inner)), _role(role)
    {
    }
    SelfTestBridgeIoDecorator(const SelfTestBridgeIoDecorator&)            = delete;
    SelfTestBridgeIoDecorator& operator=(const SelfTestBridgeIoDecorator&) = delete;
    SelfTestBridgeIoDecorator(SelfTestBridgeIoDecorator&&)                 = delete;
    SelfTestBridgeIoDecorator& operator=(SelfTestBridgeIoDecorator&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystemIO))
        {
            *ppvObject = static_cast<IFileSystemIO*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (current == 0u)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override
    {
        return _inner ? _inner->GetAttributes(path, fileAttributes) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override
    {
        if (reader == nullptr)
        {
            return E_POINTER;
        }
        *reader = nullptr;
        if (! _inner)
        {
            return E_POINTER;
        }

        if (_role == SelfTestBridgeIoRole::Destination &&
            ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeFailNextDestinationOpenCount, g_fileOpsBridgeFailNextDestinationOpenAttempts))
        {
            return g_fileOpsBridgeFailNextDestinationOpenStatus.load(std::memory_order_acquire);
        }
        if (_role == SelfTestBridgeIoRole::Source &&
            ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeNullNextSourceReaderCount, g_fileOpsBridgeNullNextSourceReaderAttempts))
        {
            return S_OK;
        }

        wil::com_ptr<IFileReader> innerReader;
        const HRESULT hr = _inner->CreateFileReader(path, innerReader.addressof());
        if (FAILED(hr))
        {
            return hr;
        }
        if (! innerReader)
        {
            return E_POINTER;
        }

        auto* decoratedReader = new (std::nothrow) SelfTestBridgeFileReader(std::move(innerReader), _role);
        if (decoratedReader == nullptr)
        {
            return E_OUTOFMEMORY;
        }

        wil::com_ptr<IFileReader> decorated;
        decorated.attach(decoratedReader);
        *reader = decorated.detach();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override
    {
        if (writer == nullptr)
        {
            return E_POINTER;
        }
        *writer = nullptr;
        if (! _inner)
        {
            return E_POINTER;
        }

        wil::com_ptr<IFileWriter> innerWriter;
        const HRESULT hr = _inner->CreateFileWriter(path, flags, innerWriter.addressof());
        if (FAILED(hr))
        {
            return hr;
        }
        if (! innerWriter)
        {
            return E_POINTER;
        }
        if (_role != SelfTestBridgeIoRole::Destination)
        {
            *writer = innerWriter.detach();
            return S_OK;
        }

        auto* decoratedWriter = new (std::nothrow) SelfTestBridgeFileWriter(std::move(innerWriter));
        if (decoratedWriter == nullptr)
        {
            return E_OUTOFMEMORY;
        }

        wil::com_ptr<IFileWriter> decorated;
        decorated.attach(decoratedWriter);
        *writer = decorated.detach();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override
    {
        return _inner ? _inner->GetFileBasicInformation(path, info) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override
    {
        return _inner ? _inner->SetFileBasicInformation(path, info) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override
    {
        return _inner ? _inner->GetItemProperties(path, jsonUtf8) : E_POINTER;
    }

private:
    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystemIO> _inner;
    SelfTestBridgeIoRole _role = SelfTestBridgeIoRole::Source;
};

[[nodiscard]] HRESULT DecorateBridgeIoForSelfTest(wil::com_ptr<IFileSystemIO>& io, SelfTestBridgeIoRole role) noexcept
{
    if (! io)
    {
        return E_POINTER;
    }

    auto* decoratedIo = new (std::nothrow) SelfTestBridgeIoDecorator(io, role);
    if (decoratedIo == nullptr)
    {
        return E_OUTOFMEMORY;
    }

    wil::com_ptr<IFileSystemIO> decorated;
    decorated.attach(decoratedIo);
    io = std::move(decorated);
    return S_OK;
}

void MaybePauseAfterTaskFinishedBeforeSummaryForSelfTest() noexcept
{
    g_fileOpsPostFinishedCompletionPausePoint.Pause(5'000ull);
}

std::mutex g_fileOpsPreparedTransferStrategyMutex;
std::array<PreparedTransferRecordForSelfTest, 64> g_fileOpsPreparedTransferStrategies{};
size_t g_fileOpsPreparedTransferStrategyNext = 0u;

void RecordPreparedTransferStrategyForSelfTest(const FolderWindow::FileOperationState::Task& task) noexcept
{
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = task.LoadPlans();
    if (! plans)
    {
        return;
    }
    PreparedTransferRecordForSelfTest record{.taskId = task.GetId()};
    for (const FileOperations::FileOperationPlan& plan : *plans)
    {
        if (const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan); transfer != nullptr)
        {
            if (record.planCount == 0u)
            {
                record.firstStrategy = static_cast<uint8_t>(transfer->strategy);
            }
            record.strategyMask |= 1u << static_cast<uint32_t>(transfer->strategy);
            ++record.planCount;
        }
    }
    if (record.planCount == 0u)
    {
        return;
    }
    std::scoped_lock lock(g_fileOpsPreparedTransferStrategyMutex);
    g_fileOpsPreparedTransferStrategies[g_fileOpsPreparedTransferStrategyNext] = record;
    g_fileOpsPreparedTransferStrategyNext = (g_fileOpsPreparedTransferStrategyNext + 1u) % g_fileOpsPreparedTransferStrategies.size();
}

void MaybePauseBeforeBridgeMoveSourceCleanupForSelfTest() noexcept
{
    g_fileOpsBridgeMoveSourceCleanupPausePoint.Pause(5'000ull);
}

void MaybeCreateNativeMoveDestinationDirectoryRaceForSelfTest(const std::wstring& destinationPath) noexcept
{
    if (! ConsumeBridgeCounterForSelfTest(g_fileOpsNativeMoveCreateDirectoryRaceCount, g_fileOpsNativeMoveCreateDirectoryRaceAttempts))
    {
        return;
    }

    const BOOL created = CreateDirectoryW(destinationPath.c_str(), nullptr);
    if (created == 0 && GetLastError() != ERROR_ALREADY_EXISTS)
    {
        Debug::ErrorWithLastError(L"FileOps self-test could not create the Native Move destination-directory race at '{}'.", destinationPath);
    }
}

void MaybeInjectBridgeCreateDirectoryRaceForSelfTest(IFileSystemIO& destinationIo, const std::wstring& destinationPath) noexcept
{
    constexpr const wchar_t* kRacePathEnv = L"REDSALAMANDER_FILEOPS_BRIDGE_CREATE_DIRECTORY_RACE_PATH";

    const std::optional<std::wstring> configured = TryReadEnvironmentVariableForSelfTest(kRacePathEnv);
    if (! configured.has_value())
    {
        return;
    }

    if (CompareStringOrdinal(configured->c_str(), -1, destinationPath.c_str(), -1, TRUE) != CSTR_EQUAL)
    {
        return;
    }

    static_cast<void>(SetEnvironmentVariableW(kRacePathEnv, nullptr));

    wil::com_ptr<IFileWriter> writer;
    const HRESULT hrWriter = destinationIo.CreateFileWriter(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, writer.addressof());
    if (SUCCEEDED(hrWriter) && writer)
    {
        static_cast<void>(writer->Commit());
    }
}

void MaybeReplaceBridgeDestinationFromEnvironmentForSelfTest(IFileSystemIO& destinationIo,
                                                             const std::wstring& destinationPath,
                                                             const wchar_t* pathEnvironmentName,
                                                             const wchar_t* payloadEnvironmentName,
                                                             std::atomic<unsigned long>& attemptCount) noexcept
{
    const std::optional<std::wstring> configured = TryReadEnvironmentVariableForSelfTest(pathEnvironmentName);
    auto normalizeForCompare                     = [](std::wstring value) noexcept
    {
        std::replace(value.begin(), value.end(), L'/', L'\\');
        return value;
    };

    if (! configured.has_value())
    {
        return;
    }

    const std::wstring configuredPath = normalizeForCompare(*configured);
    const std::wstring candidatePath  = normalizeForCompare(destinationPath);
    if (CompareStringOrdinal(configuredPath.c_str(), -1, candidatePath.c_str(), -1, TRUE) != CSTR_EQUAL)
    {
        return;
    }

    static_cast<void>(SetEnvironmentVariableW(pathEnvironmentName, nullptr));
    const std::optional<std::wstring> payload = TryReadEnvironmentVariableForSelfTest(payloadEnvironmentName);
    if (! payload.has_value())
    {
        return;
    }

    const std::string bytes = NarrowEnvironmentPayloadForSelfTest(*payload);
    if (bytes.size() > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return;
    }

    attemptCount.fetch_add(1u, std::memory_order_acq_rel);

    wil::com_ptr<IFileWriter> writer;
    const HRESULT hrWriter = destinationIo.CreateFileWriter(destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.addressof());
    if (FAILED(hrWriter) || ! writer)
    {
        return;
    }

    if (! bytes.empty())
    {
        unsigned long written = 0;
        const HRESULT hrWrite = writer->Write(bytes.data(), static_cast<unsigned long>(bytes.size()), &written);
        if (FAILED(hrWrite) || written != static_cast<unsigned long>(bytes.size()))
        {
            return;
        }
    }

    static_cast<void>(writer->Commit());
}

void MaybeReplacePublishedBridgeDestinationBeforeVerificationForSelfTest(IFileSystemIO& destinationIo, const std::wstring& destinationPath) noexcept
{
    MaybeReplaceBridgeDestinationFromEnvironmentForSelfTest(destinationIo,
                                                            destinationPath,
                                                            L"REDSALAMANDER_FILEOPS_BRIDGE_REPLACE_PUBLISHED_DESTINATION_PATH",
                                                            L"REDSALAMANDER_FILEOPS_BRIDGE_REPLACE_PUBLISHED_DESTINATION_PAYLOAD",
                                                            g_fileOpsBridgeReplacePublishedDestinationAttempts);
}

void MaybePauseBeforePublishedDestinationRetryBackoffForSelfTest() noexcept
{
    g_fileOpsBridgePublishedDestinationRetryPausePoint.Pause(5'000ull);
}

[[nodiscard]] HRESULT DecorateBridgeWriterForSelfTest(wil::com_ptr<IFileWriter>& writer) noexcept
{
    if (! writer)
    {
        return E_POINTER;
    }

    auto* decoratedWriter = new (std::nothrow) SelfTestBridgeFileWriter(writer);
    if (decoratedWriter == nullptr)
    {
        return E_OUTOFMEMORY;
    }

    wil::com_ptr<IFileWriter> decorated;
    decorated.attach(decoratedWriter);
    writer = std::move(decorated);
    return S_OK;
}
#endif

struct AutoConcurrencyResolution final
{
    unsigned int concurrency = 0u;
    uint32_t storageKind     = FILESYSTEM_STORAGE_UNKNOWN;

    [[nodiscard]] bool HasValue() const noexcept
    {
        return concurrency > 0u;
    }
};

[[nodiscard]] ReparsePointPolicy ParseReparsePointPolicy(std::string_view text) noexcept
{
    if (text == "skip" || text == "followTargets")
    {
        return ReparsePointPolicy::Skip;
    }

    // `copyReparse` is the legacy persisted spelling for Preserve. No stored value can
    // restore target-following behavior.
    return ReparsePointPolicy::Preserve;
}

[[nodiscard]] FileSystemConcurrencyMode ParseConcurrencyMode(std::string_view text) noexcept
{
    if (text == "manual")
    {
        return FileSystemConcurrencyMode::Manual;
    }

    return FileSystemConcurrencyMode::Auto;
}

[[nodiscard]] const wchar_t* ConcurrencyModeToString(FileSystemConcurrencyMode mode) noexcept
{
    return mode == FileSystemConcurrencyMode::Manual ? L"manual" : L"auto";
}

[[nodiscard]] const wchar_t* StorageKindToString(uint32_t storageKind) noexcept
{
    switch (storageKind)
    {
        case FILESYSTEM_STORAGE_HDD: return L"hdd";
        case FILESYSTEM_STORAGE_SSD: return L"ssd";
        case FILESYSTEM_STORAGE_NVME: return L"nvme";
        case FILESYSTEM_STORAGE_NETWORK_SHARE: return L"networkShare";
        case FILESYSTEM_STORAGE_CLOUD: return L"cloud";
        case FILESYSTEM_STORAGE_VIRTUAL: return L"virtual";
        case FILESYSTEM_STORAGE_MEMORY: return L"memory";
        default: return L"unknown";
    }
}

struct ParsedFileSystemConfiguration final
{
    ParsedFileSystemConfiguration()                                                = default;
    ParsedFileSystemConfiguration(const ParsedFileSystemConfiguration&)            = delete;
    ParsedFileSystemConfiguration& operator=(const ParsedFileSystemConfiguration&) = delete;
    ParsedFileSystemConfiguration(ParsedFileSystemConfiguration&&)                 = default;
    ParsedFileSystemConfiguration& operator=(ParsedFileSystemConfiguration&&)      = default;

    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc{nullptr, &yyjson_doc_free};
    yyjson_val* root = nullptr;

    [[nodiscard]] explicit operator bool() const noexcept
    {
        return doc != nullptr && root != nullptr;
    }
};

[[nodiscard]] ParsedFileSystemConfiguration TryParseFileSystemConfiguration(const wil::com_ptr<IFileSystem>& fileSystem) noexcept
{
    ParsedFileSystemConfiguration parsed;
    if (! fileSystem)
    {
        return parsed;
    }

    wil::com_ptr<IInformations> informations;
    if (FAILED(fileSystem->QueryInterface(__uuidof(IInformations), informations.put_void())) || ! informations)
    {
        return parsed;
    }

    const char* configurationJsonUtf8 = nullptr;
    if (FAILED(informations->GetConfiguration(&configurationJsonUtf8)) || ! configurationJsonUtf8)
    {
        return parsed;
    }

    const size_t configurationBytes = std::strlen(configurationJsonUtf8);
    if (configurationBytes == 0)
    {
        return parsed;
    }

    parsed.doc.reset(yyjson_read(configurationJsonUtf8, configurationBytes, YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM));
    if (! parsed.doc)
    {
        return parsed;
    }

    parsed.root = yyjson_doc_get_root(parsed.doc.get());
    if (! parsed.root || ! yyjson_is_obj(parsed.root))
    {
        parsed.root = nullptr;
    }

    return parsed;
}

[[nodiscard]] std::optional<FileSystemConcurrencyMode> TryGetConcurrencyModeFromFileSystem(const wil::com_ptr<IFileSystem>& fileSystem) noexcept
{
    ParsedFileSystemConfiguration parsed = TryParseFileSystemConfiguration(fileSystem);
    if (! parsed)
    {
        return std::nullopt;
    }

    yyjson_val* modeVal = yyjson_obj_get(parsed.root, "concurrencyMode");
    if (! modeVal || ! yyjson_is_str(modeVal))
    {
        return FileSystemConcurrencyMode::Auto;
    }

    const char* modeText = yyjson_get_str(modeVal);
    if (! modeText || modeText[0] == '\0')
    {
        return FileSystemConcurrencyMode::Auto;
    }

    return ParseConcurrencyMode(modeText);
}

[[nodiscard]] AutoConcurrencyResolution ResolveAutoPerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                         const std::vector<std::filesystem::path>& paths,
                                                                         FileSystemOperation operation,
                                                                         unsigned int uiMax) noexcept;
[[nodiscard]] AutoConcurrencyResolution ResolveAutoPerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                         const std::filesystem::path& path,
                                                                         FileSystemOperation operation,
                                                                         unsigned int uiMax) noexcept;

[[nodiscard]] ReparsePointPolicy GetReparsePointPolicyFromSettings(const Common::Settings::Settings& settings, const std::wstring& pluginId) noexcept
{
    const auto it = settings.plugins.configurationByPluginId.find(pluginId);
    if (it == settings.plugins.configurationByPluginId.end())
    {
        return ReparsePointPolicy::Preserve;
    }

    const Common::Settings::JsonValue& config = it->second;
    if (! std::holds_alternative<Common::Settings::JsonValue::ObjectPtr>(config.value))
    {
        return ReparsePointPolicy::Preserve;
    }

    const auto obj = std::get<Common::Settings::JsonValue::ObjectPtr>(config.value);
    if (! obj)
    {
        return ReparsePointPolicy::Preserve;
    }

    for (const auto& member : obj->members)
    {
        if (member.first != "reparsePointPolicy")
        {
            continue;
        }

        const Common::Settings::JsonValue& v = member.second;
        if (! std::holds_alternative<std::string>(v.value))
        {
            return ReparsePointPolicy::Preserve;
        }

        const std::string& text = std::get<std::string>(v.value);
        return ParseReparsePointPolicy(text);
    }

    return ReparsePointPolicy::Preserve;
}

[[nodiscard]] uint64_t MeasurePathBytes(std::wstring_view path) noexcept
{
    return static_cast<uint64_t>(path.size()) * sizeof(wchar_t);
}

constexpr ULONGLONG kVisibleProgressPathRefreshIntervalMs = 100ull;

void UpdateTrackedPath(std::wstring& target, const wchar_t* source, uint64_t& bytesCounter, uint64_t& appliedCounter, uint64_t& skippedCounter) noexcept
{
    const std::wstring_view sourceView = (source && source[0] != L'\0') ? std::wstring_view(source) : std::wstring_view{};
    if (target == sourceView)
    {
        ++skippedCounter;
        return;
    }

    target.assign(sourceView);
    bytesCounter += MeasurePathBytes(sourceView);
    ++appliedCounter;
}

void UpdateTrackedPathIfPresent(
    std::wstring& target, const wchar_t* source, uint64_t& bytesCounter, uint64_t& appliedCounter, uint64_t& skippedCounter) noexcept
{
    if (! source || source[0] == L'\0')
    {
        return;
    }

    UpdateTrackedPath(target, source, bytesCounter, appliedCounter, skippedCounter);
}

[[nodiscard]] bool IsSameOrChildPath(std::wstring_view root, std::wstring_view candidate) noexcept;

void PublishDiagnosticPathSnapshotLocked(FolderWindow::FileOperationState::Task& task)
{
    using Task = FolderWindow::FileOperationState::Task;

    auto snapshot                                 = std::make_shared<Task::DiagnosticPathSnapshot>();
    snapshot->progressSourcePath                  = task._progressSourcePath;
    snapshot->progressDestinationPath             = task._progressDestinationPath;
    snapshot->lastProgressCallbackSourcePath      = task._lastProgressCallbackSourcePath;
    snapshot->lastProgressCallbackDestinationPath = task._lastProgressCallbackDestinationPath;

    std::shared_ptr<const Task::DiagnosticPathSnapshot> publishedSnapshot = std::move(snapshot);
    task._publishedDiagnosticPathSnapshot.store(std::move(publishedSnapshot), std::memory_order_release);
}

void PublishProgressCountersLocked(FolderWindow::FileOperationState::Task& task) noexcept
{
    task._publishedProgressTotalItems.store(task._progressTotalItems, std::memory_order_release);
    task._publishedProgressCompletedItems.store(task._progressCompletedItems, std::memory_order_release);
    task._publishedProgressTotalBytes.store(task._progressTotalBytes, std::memory_order_release);
    task._publishedProgressCompletedBytes.store(task._progressCompletedBytes, std::memory_order_release);
    task._publishedProgressItemTotalBytes.store(task._progressItemTotalBytes, std::memory_order_release);
    task._publishedProgressItemCompletedBytes.store(task._progressItemCompletedBytes, std::memory_order_release);
}

struct TopLevelCompletionSnapshot
{
    unsigned long completedFiles   = 0;
    unsigned long completedFolders = 0;
};

void StorePublishedTopLevelCompletionSnapshot(FolderWindow::FileOperationState::Task& task, const TopLevelCompletionSnapshot& snapshot) noexcept
{
    task._publishedCompletedTopLevelFiles.store(snapshot.completedFiles, std::memory_order_release);
    task._publishedCompletedTopLevelFolders.store(snapshot.completedFolders, std::memory_order_release);
}

[[maybe_unused]] TopLevelCompletionSnapshot LoadTopLevelCompletionSnapshot(const FolderWindow::FileOperationState::Task& task) noexcept
{
    TopLevelCompletionSnapshot snapshot{};
    snapshot.completedFiles   = task._publishedCompletedTopLevelFiles.load(std::memory_order_acquire);
    snapshot.completedFolders = task._publishedCompletedTopLevelFolders.load(std::memory_order_acquire);
    return snapshot;
}

TopLevelCompletionSnapshot MarkTopLevelItemCompleted(FolderWindow::FileOperationState::Task& task, size_t index) noexcept
{
    TopLevelCompletionSnapshot snapshot{};
    std::scoped_lock lock(task._topLevelCompletionMutex);
    if (index < task._topLevelItemCompleted.size() && task._topLevelItemCompleted[index] == 0)
    {
        task._topLevelItemCompleted[index] = 1;
        if (index < task._topLevelItemKinds.size())
        {
            const auto kind = task._topLevelItemKinds[index];
            if (kind == Task::TopLevelItemKind::File)
            {
                if (task._completedTopLevelFiles < std::numeric_limits<unsigned long>::max())
                {
                    ++task._completedTopLevelFiles;
                }
            }
            else if (kind == Task::TopLevelItemKind::Folder)
            {
                if (task._completedTopLevelFolders < std::numeric_limits<unsigned long>::max())
                {
                    ++task._completedTopLevelFolders;
                }
            }
        }
    }

    snapshot.completedFiles   = task._completedTopLevelFiles;
    snapshot.completedFolders = task._completedTopLevelFolders;
    return snapshot;
}

PublishedProgressSnapshot CapturePublishedProgressSnapshotLocked(const FolderWindow::FileOperationState::Task& task) noexcept
{
    PublishedProgressSnapshot snapshot{};
    snapshot.totalItems                 = task._progressTotalItems;
    snapshot.completedItems             = task._progressCompletedItems;
    snapshot.totalBytes                 = task._progressTotalBytes;
    snapshot.completedBytes             = task._progressCompletedBytes;
    snapshot.itemTotalBytes             = task._progressItemTotalBytes;
    snapshot.itemCompletedBytes         = task._progressItemCompletedBytes;
    snapshot.completedFiles             = task._publishedCompletedTopLevelFiles.load(std::memory_order_relaxed);
    snapshot.completedFolders           = task._publishedCompletedTopLevelFolders.load(std::memory_order_relaxed);
    snapshot.progressCallbackCount      = task._progressCallbackCount.load(std::memory_order_relaxed);
    snapshot.itemCompletedCallbackCount = task._itemCompletedCallbackCount.load(std::memory_order_relaxed);
    return snapshot;
}

void StorePublishedProgressSnapshot(FolderWindow::FileOperationState::Task& task, const PublishedProgressSnapshot& snapshot) noexcept
{
    task._publishedProgressTotalItems.store(snapshot.totalItems, std::memory_order_release);
    task._publishedProgressCompletedItems.store(snapshot.completedItems, std::memory_order_release);
    task._publishedProgressTotalBytes.store(snapshot.totalBytes, std::memory_order_release);
    task._publishedProgressCompletedBytes.store(snapshot.completedBytes, std::memory_order_release);
    task._publishedProgressItemTotalBytes.store(snapshot.itemTotalBytes, std::memory_order_release);
    task._publishedProgressItemCompletedBytes.store(snapshot.itemCompletedBytes, std::memory_order_release);
}

PublishedProgressSnapshot LoadPublishedProgressSnapshot(const FolderWindow::FileOperationState::Task& task) noexcept
{
    PublishedProgressSnapshot snapshot{};
    snapshot.totalItems                 = task._publishedProgressTotalItems.load(std::memory_order_acquire);
    snapshot.completedItems             = task._publishedProgressCompletedItems.load(std::memory_order_acquire);
    snapshot.totalBytes                 = task._publishedProgressTotalBytes.load(std::memory_order_acquire);
    snapshot.completedBytes             = task._publishedProgressCompletedBytes.load(std::memory_order_acquire);
    snapshot.itemTotalBytes             = task._publishedProgressItemTotalBytes.load(std::memory_order_acquire);
    snapshot.itemCompletedBytes         = task._publishedProgressItemCompletedBytes.load(std::memory_order_acquire);
    snapshot.completedFiles             = task._publishedCompletedTopLevelFiles.load(std::memory_order_acquire);
    snapshot.completedFolders           = task._publishedCompletedTopLevelFolders.load(std::memory_order_acquire);
    snapshot.progressCallbackCount      = task._progressCallbackCount.load(std::memory_order_acquire);
    snapshot.itemCompletedCallbackCount = task._itemCompletedCallbackCount.load(std::memory_order_acquire);
    return snapshot;
}

void CopyEffectiveProgressPathsLocked(const FolderWindow::FileOperationState::Task& task,
                                      std::wstring& sourcePath,
                                      std::wstring& destinationPath,
                                      ULONGLONG* lastProgressCallbackTick = nullptr) noexcept
{
    sourcePath      = ! task._lastProgressCallbackSourcePath.empty() ? task._lastProgressCallbackSourcePath : task._progressSourcePath;
    destinationPath = ! task._lastProgressCallbackDestinationPath.empty() ? task._lastProgressCallbackDestinationPath : task._progressDestinationPath;
    if (lastProgressCallbackTick != nullptr)
    {
        *lastProgressCallbackTick = task._lastProgressCallbackTick;
    }
}

Task::ProgressStreamPerf& FindOrAddProgressStreamPerfLocked(Task& task, const void* cookieKey, uint64_t progressStreamId) noexcept
{
    for (size_t i = 0; i < task._progressStreamPerfCount; ++i)
    {
        auto& entry = task._progressStreamPerf[i];
        if (entry.cookieKey == cookieKey && entry.progressStreamId == progressStreamId)
        {
            return entry;
        }
    }

    size_t index = task._progressStreamPerfCount;
    if (index < task._progressStreamPerf.size())
    {
        ++task._progressStreamPerfCount;
    }
    else
    {
        index                = 0;
        ULONGLONG oldestTick = task._progressStreamPerf[0].lastUpdateTick;
        for (size_t i = 1; i < task._progressStreamPerfCount; ++i)
        {
            const ULONGLONG tick = task._progressStreamPerf[i].lastUpdateTick;
            if (tick == 0 || (oldestTick != 0 && tick < oldestTick))
            {
                index      = i;
                oldestTick = tick;
            }
        }
    }

    auto& entry                  = task._progressStreamPerf[index];
    entry.cookieKey              = cookieKey;
    entry.progressStreamId       = progressStreamId;
    entry.callbackCount          = 0;
    entry.callbackUs             = 0;
    entry.lockWaitUs             = 0;
    entry.callbackGapCount       = 0;
    entry.callbackGapMs          = 0;
    entry.callbackGapBytes       = 0;
    entry.maxCallbackGapMs       = 0;
    entry.maxCallbackGapBytes    = 0;
    entry.maxCallbackDeltaBytes  = 0;
    entry.lastItemCompletedBytes = 0;
    entry.firstUpdateTick        = 0;
    entry.lastUpdateTick         = 0;
    return entry;
}

void NoteProgressStreamPerf(Task& task,
                            const void* cookieKey,
                            uint64_t progressStreamId,
                            ULONGLONG progressCallbackTick,
                            uint64_t currentItemCompletedBytes,
                            uint64_t lockWaitUs,
                            uint64_t callbackUs) noexcept
{
    std::scoped_lock lock(task._progressStreamPerfMutex);
    auto& entry = FindOrAddProgressStreamPerfLocked(task, cookieKey, progressStreamId);
    // A reused stream can move to a new file, resetting current-item progress to a lower value.
    const uint64_t itemDeltaBytes =
        currentItemCompletedBytes >= entry.lastItemCompletedBytes ? (currentItemCompletedBytes - entry.lastItemCompletedBytes) : currentItemCompletedBytes;
    entry.maxCallbackDeltaBytes = (std::max)(entry.maxCallbackDeltaBytes, itemDeltaBytes);
    if (entry.callbackCount == 0)
    {
        entry.firstUpdateTick = progressCallbackTick;
    }
    else if (entry.lastUpdateTick != 0 && progressCallbackTick >= entry.lastUpdateTick)
    {
        const uint64_t gapMs = static_cast<uint64_t>(progressCallbackTick - entry.lastUpdateTick);
        entry.callbackGapMs += gapMs;
        entry.callbackGapBytes += itemDeltaBytes;
        ++entry.callbackGapCount;
        if (gapMs > entry.maxCallbackGapMs)
        {
            entry.maxCallbackGapMs    = gapMs;
            entry.maxCallbackGapBytes = itemDeltaBytes;
        }
    }

    ++entry.callbackCount;
    entry.lockWaitUs += lockWaitUs;
    entry.callbackUs += callbackUs;
    entry.lastItemCompletedBytes = currentItemCompletedBytes;
    entry.lastUpdateTick         = progressCallbackTick;
}

struct PerItemInFlightAggregate
{
    uint64_t completedBytes = 0;
    uint64_t completedItems = 0;
    uint64_t totalItems     = 0;
    size_t activeCount      = 0;
};

struct PerItemInFlightUpdateResult
{
    PerItemInFlightAggregate aggregate{};
    bool evicted              = false;
    const void* evictedCookie = nullptr;
};

struct PerItemInFlightFinishResult
{
    PerItemInFlightAggregate aggregate{};
    uint64_t completedBytes = 0;
    uint64_t completedItems = 0;
    uint64_t totalItems     = 0;
};

void AddPerItemAggregateValue(uint64_t& target, uint64_t value) noexcept
{
    if (std::numeric_limits<uint64_t>::max() - target < value)
    {
        target = std::numeric_limits<uint64_t>::max();
    }
    else
    {
        target += value;
    }
}

void SubtractPerItemAggregateValue(uint64_t& target, uint64_t value) noexcept
{
    target = (target >= value) ? (target - value) : 0;
}

void RemovePerItemInFlightEntryFromAggregate(Task& task, const Task::PerItemInFlightCall& entry) noexcept
{
    SubtractPerItemAggregateValue(task._perItemInFlightCompletedBytes, entry.completedBytes);
    SubtractPerItemAggregateValue(task._perItemInFlightCompletedItems, static_cast<uint64_t>(entry.completedItems));
    SubtractPerItemAggregateValue(task._perItemInFlightTotalItems, static_cast<uint64_t>(entry.totalItems));
}

void AddPerItemInFlightEntryToAggregate(Task& task, const Task::PerItemInFlightCall& entry) noexcept
{
    AddPerItemAggregateValue(task._perItemInFlightCompletedBytes, entry.completedBytes);
    AddPerItemAggregateValue(task._perItemInFlightCompletedItems, static_cast<uint64_t>(entry.completedItems));
    AddPerItemAggregateValue(task._perItemInFlightTotalItems, static_cast<uint64_t>(entry.totalItems));
}

PerItemInFlightAggregate SummarizePerItemInFlightCallsLocked(Task& task) noexcept
{
    PerItemInFlightAggregate aggregate{};
    aggregate.activeCount    = task._perItemInFlightCallCount;
    aggregate.completedBytes = task._perItemInFlightCompletedBytes;
    aggregate.completedItems = task._perItemInFlightCompletedItems;
    aggregate.totalItems     = task._perItemInFlightTotalItems;
    return aggregate;
}

void InitializePerItemInFlightEntry(Task::PerItemInFlightCall& entry, const void* cookieKey, ULONGLONG nowTick) noexcept
{
    entry                = {};
    entry.cookie         = cookieKey;
    entry.lastUpdateTick = nowTick;
}

PerItemInFlightAggregate ResetPerItemInFlightCalls(Task& task, const void* cookieKey, ULONGLONG nowTick) noexcept
{
    std::scoped_lock lock(task._perItemInFlightCallsMutex);
    task._perItemInFlightCallCount      = 0;
    task._perItemInFlightCompletedBytes = 0;
    task._perItemInFlightCompletedItems = 0;
    task._perItemInFlightTotalItems     = 0;
    if (! task._perItemInFlightCalls.empty())
    {
        InitializePerItemInFlightEntry(task._perItemInFlightCalls[0], cookieKey, nowTick);
        task._perItemInFlightCallCount = 1;
    }
    return SummarizePerItemInFlightCallsLocked(task);
}

PerItemInFlightAggregate BeginPerItemInFlightCall(Task& task, const void* cookieKey, ULONGLONG nowTick) noexcept
{
    std::scoped_lock lock(task._perItemInFlightCallsMutex);

    size_t found = task._perItemInFlightCallCount;
    for (size_t i = 0; i < task._perItemInFlightCallCount; ++i)
    {
        if (task._perItemInFlightCalls[i].cookie == cookieKey)
        {
            found = i;
            break;
        }
    }

    if (found < task._perItemInFlightCallCount)
    {
        RemovePerItemInFlightEntryFromAggregate(task, task._perItemInFlightCalls[found]);
        InitializePerItemInFlightEntry(task._perItemInFlightCalls[found], cookieKey, nowTick);
    }
    else if (task._perItemInFlightCallCount < task._perItemInFlightCalls.size())
    {
        InitializePerItemInFlightEntry(task._perItemInFlightCalls[task._perItemInFlightCallCount], cookieKey, nowTick);
        ++task._perItemInFlightCallCount;
    }
    else if (! task._perItemInFlightCalls.empty())
    {
        size_t replaceIndex  = 0;
        ULONGLONG oldestTick = task._perItemInFlightCalls[0].lastUpdateTick;
        for (size_t i = 1; i < task._perItemInFlightCallCount; ++i)
        {
            const ULONGLONG tick = task._perItemInFlightCalls[i].lastUpdateTick;
            if (tick == 0 || (oldestTick != 0 && tick < oldestTick))
            {
                replaceIndex = i;
                oldestTick   = tick;
            }
        }

        RemovePerItemInFlightEntryFromAggregate(task, task._perItemInFlightCalls[replaceIndex]);
        InitializePerItemInFlightEntry(task._perItemInFlightCalls[replaceIndex], cookieKey, nowTick);
    }

    return SummarizePerItemInFlightCallsLocked(task);
}

PerItemInFlightUpdateResult UpdatePerItemInFlightCall(
    Task& task, const void* cookieKey, unsigned long completedItems, uint64_t completedBytes, unsigned long totalItems, ULONGLONG nowTick) noexcept
{
    PerItemInFlightUpdateResult result{};
    std::scoped_lock lock(task._perItemInFlightCallsMutex);

    size_t found = task._perItemInFlightCallCount;
    for (size_t i = 0; i < task._perItemInFlightCallCount; ++i)
    {
        if (task._perItemInFlightCalls[i].cookie == cookieKey)
        {
            found = i;
            break;
        }
    }

    if (found < task._perItemInFlightCallCount)
    {
        RemovePerItemInFlightEntryFromAggregate(task, task._perItemInFlightCalls[found]);
        task._perItemInFlightCalls[found].completedItems = completedItems;
        task._perItemInFlightCalls[found].completedBytes = completedBytes;
        task._perItemInFlightCalls[found].lastUpdateTick = nowTick;
        if (totalItems > 0)
        {
            task._perItemInFlightCalls[found].totalItems = (std::max)(task._perItemInFlightCalls[found].totalItems, totalItems);
        }
        AddPerItemInFlightEntryToAggregate(task, task._perItemInFlightCalls[found]);
    }
    else
    {
        const auto populateEntry = [&](Task::PerItemInFlightCall& entry) noexcept
        {
            entry.cookie         = cookieKey;
            entry.completedItems = completedItems;
            entry.completedBytes = completedBytes;
            entry.totalItems     = totalItems;
            entry.lastUpdateTick = nowTick;
        };

        if (task._perItemInFlightCallCount < task._perItemInFlightCalls.size())
        {
            populateEntry(task._perItemInFlightCalls[task._perItemInFlightCallCount]);
            AddPerItemInFlightEntryToAggregate(task, task._perItemInFlightCalls[task._perItemInFlightCallCount]);
            ++task._perItemInFlightCallCount;
        }
        else if (! task._perItemInFlightCalls.empty())
        {
            size_t replaceIndex  = 0;
            ULONGLONG oldestTick = task._perItemInFlightCalls[0].lastUpdateTick;
            for (size_t i = 1; i < task._perItemInFlightCallCount; ++i)
            {
                const ULONGLONG tick = task._perItemInFlightCalls[i].lastUpdateTick;
                if (tick == 0 || (oldestTick != 0 && tick < oldestTick))
                {
                    replaceIndex = i;
                    oldestTick   = tick;
                }
            }

            result.evicted       = true;
            result.evictedCookie = task._perItemInFlightCalls[replaceIndex].cookie;
            RemovePerItemInFlightEntryFromAggregate(task, task._perItemInFlightCalls[replaceIndex]);
            populateEntry(task._perItemInFlightCalls[replaceIndex]);
            AddPerItemInFlightEntryToAggregate(task, task._perItemInFlightCalls[replaceIndex]);
        }
    }

    result.aggregate = SummarizePerItemInFlightCallsLocked(task);
    return result;
}

PerItemInFlightFinishResult FinishPerItemInFlightCall(Task& task, const void* cookieKey) noexcept
{
    PerItemInFlightFinishResult result{};
    std::scoped_lock lock(task._perItemInFlightCallsMutex);

    for (size_t i = 0; i < task._perItemInFlightCallCount; ++i)
    {
        if (task._perItemInFlightCalls[i].cookie == cookieKey)
        {
            result.completedItems = task._perItemInFlightCalls[i].completedItems;
            result.completedBytes = task._perItemInFlightCalls[i].completedBytes;
            result.totalItems     = static_cast<uint64_t>(task._perItemInFlightCalls[i].totalItems);
            RemovePerItemInFlightEntryFromAggregate(task, task._perItemInFlightCalls[i]);
            task._perItemInFlightCalls[i] = task._perItemInFlightCalls[task._perItemInFlightCallCount - 1u];
            --task._perItemInFlightCallCount;
            break;
        }
    }

    result.aggregate = SummarizePerItemInFlightCallsLocked(task);
    return result;
}

size_t GetPerItemInFlightCallCountSnapshot(Task& task) noexcept
{
    std::scoped_lock lock(task._perItemInFlightCallsMutex);
    return task._perItemInFlightCallCount;
}

PerItemInFlightAggregate GetPerItemInFlightAggregate(Task& task) noexcept
{
    std::scoped_lock lock(task._perItemInFlightCallsMutex);
    return SummarizePerItemInFlightCallsLocked(task);
}

void ApplyCallbackBandwidthLimit(Task& task, FileSystemOptions* options, unsigned int perItemActiveCallsSnapshot) noexcept
{
    if (options == nullptr || (task._operation != FILESYSTEM_COPY && task._operation != FILESYSTEM_MOVE))
    {
        return;
    }

    const uint64_t pluginEffective = options->bandwidthLimitBytesPerSecond;
    const uint64_t desiredTotal    = task._desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

    if (task._executionMode == FolderWindow::FileOperationState::ExecutionMode::PerItem && task._perItemMaxConcurrency > 1u)
    {
        uint64_t desiredPerCall = desiredTotal;
        if (desiredTotal > 0)
        {
            const unsigned int activeCalls = std::max(1u, perItemActiveCallsSnapshot);
            desiredPerCall                 = std::max<uint64_t>(uint64_t{1}, desiredTotal / static_cast<uint64_t>(activeCalls));
        }

        // Keep the UI limit line in task units (total), while applying the per-call share to the plugin.
        task._effectiveSpeedLimitBytesPerSecond.store(desiredTotal, std::memory_order_release);
        options->bandwidthLimitBytesPerSecond = desiredPerCall;
        task._appliedSpeedLimitBytesPerSecond.store(desiredPerCall, std::memory_order_release);
        return;
    }

    const uint64_t applied = task._appliedSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    task._effectiveSpeedLimitBytesPerSecond.store(pluginEffective, std::memory_order_release);
    if (desiredTotal != applied)
    {
        options->bandwidthLimitBytesPerSecond = desiredTotal;
        task._appliedSpeedLimitBytesPerSecond.store(desiredTotal, std::memory_order_release);
    }
}

void UpdateProgressPathState(Task& task,
                             FolderWindow::FileOperationState::Task::PerItemCallbackCookie* perItemCookie,
                             const wchar_t* currentSourcePath,
                             const wchar_t* currentDestinationPath,
                             ULONGLONG nowTick) noexcept
{
    std::scoped_lock lock(task._progressPathMutex);

    bool publishDiagnosticPathSnapshot = false;
    const std::wstring_view currentSourceView =
        (currentSourcePath && currentSourcePath[0] != L'\0') ? std::wstring_view(currentSourcePath) : std::wstring_view{};
    const std::wstring_view currentDestinationView =
        (currentDestinationPath && currentDestinationPath[0] != L'\0') ? std::wstring_view(currentDestinationPath) : std::wstring_view{};
    const bool sourceChanged                = task._progressSourcePath != currentSourceView;
    const bool destinationChanged           = task._progressDestinationPath != currentDestinationView;
    const bool shouldApplyVisiblePathUpdate = task._lastVisibleProgressPathUpdateTick == 0 || nowTick < task._lastVisibleProgressPathUpdateTick ||
                                              (nowTick - task._lastVisibleProgressPathUpdateTick) >= kVisibleProgressPathRefreshIntervalMs;

    if (sourceChanged)
    {
        if (shouldApplyVisiblePathUpdate)
        {
            task._progressSourcePath.assign(currentSourceView);
            task._perf.progressPathUpdateBytes += MeasurePathBytes(currentSourceView);
            ++task._perf.progressPathUpdateAppliedCount;
            publishDiagnosticPathSnapshot = true;
        }
        else
        {
            ++task._perf.progressPathUpdateThrottledCount;
        }
    }
    else
    {
        ++task._perf.progressPathUpdateSkippedCount;
    }

    if (destinationChanged)
    {
        if (shouldApplyVisiblePathUpdate)
        {
            task._progressDestinationPath.assign(currentDestinationView);
            task._perf.progressPathUpdateBytes += MeasurePathBytes(currentDestinationView);
            ++task._perf.progressPathUpdateAppliedCount;
            publishDiagnosticPathSnapshot = true;
        }
        else
        {
            ++task._perf.progressPathUpdateThrottledCount;
        }
    }
    else
    {
        ++task._perf.progressPathUpdateSkippedCount;
    }

    if ((sourceChanged || destinationChanged) && shouldApplyVisiblePathUpdate)
    {
        task._lastVisibleProgressPathUpdateTick = nowTick;
    }
    if (task._lastProgressCallbackSourcePath != currentSourceView)
    {
        task._lastProgressCallbackSourcePath.assign(currentSourceView);
        publishDiagnosticPathSnapshot = true;
    }
    if (task._lastProgressCallbackDestinationPath != currentDestinationView)
    {
        task._lastProgressCallbackDestinationPath.assign(currentDestinationView);
        publishDiagnosticPathSnapshot = true;
    }
    task._lastProgressCallbackTick = nowTick;

    if (perItemCookie != nullptr)
    {
        if (currentSourcePath && currentSourcePath[0] != L'\0')
        {
            if (perItemCookie->lastProgressSourcePath == currentSourceView)
            {
                ++task._perf.progressPathUpdateSkippedCount;
            }
            else
            {
                perItemCookie->lastProgressSourcePath.assign(currentSourceView);
                task._perf.progressPathUpdateBytes += MeasurePathBytes(currentSourceView);
                ++task._perf.progressPathUpdateAppliedCount;
            }
        }
        if (currentDestinationPath && currentDestinationPath[0] != L'\0')
        {
            if (perItemCookie->lastProgressDestinationPath == currentDestinationView)
            {
                ++task._perf.progressPathUpdateSkippedCount;
            }
            else
            {
                perItemCookie->lastProgressDestinationPath.assign(currentDestinationView);
                task._perf.progressPathUpdateBytes += MeasurePathBytes(currentDestinationView);
                ++task._perf.progressPathUpdateAppliedCount;
            }
        }
    }

    if (publishDiagnosticPathSnapshot)
    {
        PublishDiagnosticPathSnapshotLocked(task);
    }
}

void UpdateItemCompletedPathState(Task& task,
                                  FolderWindow::FileOperationState::Task::PerItemCallbackCookie* perItemCookie,
                                  const wchar_t* sourcePath,
                                  const wchar_t* destinationPath) noexcept
{
    std::scoped_lock lock(task._progressPathMutex);

    bool publishDiagnosticPathSnapshot = false;
    if (! task._lastProgressCallbackSourcePath.empty() && sourcePath && sourcePath[0] != L'\0')
    {
        ++task._perf.itemCompletedPathUpdateSkippedCount;
    }
    else
    {
        const std::wstring_view sourceView = (sourcePath && sourcePath[0] != L'\0') ? std::wstring_view(sourcePath) : std::wstring_view{};
        publishDiagnosticPathSnapshot |= (task._progressSourcePath != sourceView);
        UpdateTrackedPath(task._progressSourcePath,
                          sourcePath,
                          task._perf.itemCompletedPathUpdateBytes,
                          task._perf.itemCompletedPathUpdateAppliedCount,
                          task._perf.itemCompletedPathUpdateSkippedCount);
    }
    if (! task._lastProgressCallbackDestinationPath.empty() && destinationPath && destinationPath[0] != L'\0')
    {
        ++task._perf.itemCompletedPathUpdateSkippedCount;
    }
    else
    {
        const std::wstring_view destinationView = (destinationPath && destinationPath[0] != L'\0') ? std::wstring_view(destinationPath) : std::wstring_view{};
        publishDiagnosticPathSnapshot |= (task._progressDestinationPath != destinationView);
        UpdateTrackedPath(task._progressDestinationPath,
                          destinationPath,
                          task._perf.itemCompletedPathUpdateBytes,
                          task._perf.itemCompletedPathUpdateAppliedCount,
                          task._perf.itemCompletedPathUpdateSkippedCount);
    }

    if (perItemCookie != nullptr)
    {
        if (perItemCookie->lastProgressSourcePath.empty() && sourcePath && sourcePath[0] != L'\0')
        {
            const std::wstring_view sourceView(sourcePath);
            perItemCookie->lastProgressSourcePath.assign(sourceView);
            task._perf.itemCompletedPathUpdateBytes += MeasurePathBytes(sourceView);
            ++task._perf.itemCompletedPathUpdateAppliedCount;
        }
        else if (sourcePath && sourcePath[0] != L'\0')
        {
            ++task._perf.itemCompletedPathUpdateSkippedCount;
        }
        if (perItemCookie->lastProgressDestinationPath.empty() && destinationPath && destinationPath[0] != L'\0')
        {
            const std::wstring_view destinationView(destinationPath);
            perItemCookie->lastProgressDestinationPath.assign(destinationView);
            task._perf.itemCompletedPathUpdateBytes += MeasurePathBytes(destinationView);
            ++task._perf.itemCompletedPathUpdateAppliedCount;
        }
        else if (destinationPath && destinationPath[0] != L'\0')
        {
            ++task._perf.itemCompletedPathUpdateSkippedCount;
        }
    }

    if (publishDiagnosticPathSnapshot)
    {
        PublishDiagnosticPathSnapshotLocked(task);
    }
}

void SaturatingAtomicAdd(std::atomic<uint64_t>& target, uint64_t value) noexcept
{
    uint64_t current = target.load(std::memory_order_acquire);
    for (;;)
    {
        const uint64_t desired = current > std::numeric_limits<uint64_t>::max() - value ? std::numeric_limits<uint64_t>::max() : current + value;
        if (target.compare_exchange_weak(current, desired, std::memory_order_acq_rel, std::memory_order_acquire))
        {
            return;
        }
    }
}

void UpdateInFlightFileProgress(Task& task,
                                const void* cookieKey,
                                uint64_t progressStreamId,
                                const wchar_t* currentSourcePath,
                                uint64_t currentItemTotalBytes,
                                uint64_t currentItemCompletedBytes,
                                ULONGLONG nowTick,
                                bool discoveryOpenAtCallback) noexcept
{
    if (currentSourcePath == nullptr || currentSourcePath[0] == L'\0')
    {
        return;
    }

    std::scoped_lock lock(task._inFlightFilesMutex);

    constexpr ULONGLONG kExpiryMsActive    = 10'000ull;
    constexpr ULONGLONG kExpiryMsCompleted = 300ull;

    size_t write = 0;
    for (size_t read = 0; read < task._inFlightFileCount; ++read)
    {
        const Task::InFlightFileProgress& entry = task._inFlightFiles[read];
        const bool completed                    = entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes;
        const ULONGLONG expiryMs                = completed ? kExpiryMsCompleted : kExpiryMsActive;
        const bool expired                      = entry.lastUpdateTick != 0 && nowTick >= entry.lastUpdateTick && (nowTick - entry.lastUpdateTick) > expiryMs;
        if (expired)
        {
            continue;
        }

        if (write != read)
        {
            task._inFlightFiles[write] = std::move(task._inFlightFiles[read]);
        }
        ++write;
    }
    task._inFlightFileCount = write;

    const uint64_t streamKey = progressStreamId;
    size_t found             = task._inFlightFileCount;
    for (size_t i = 0; i < task._inFlightFileCount; ++i)
    {
        if (task._inFlightFiles[i].cookieKey == cookieKey && task._inFlightFiles[i].progressStreamId == streamKey)
        {
            found = i;
            break;
        }
    }

    if (found < task._inFlightFileCount)
    {
        if (task._inFlightFiles[found].sourcePath != currentSourcePath)
        {
            task._inFlightFiles[found].sourcePath.assign(currentSourcePath);
            task._inFlightFiles[found].completionCountedWhileDiscoveryOpen = false;
        }
        task._inFlightFiles[found].totalBytes     = currentItemTotalBytes;
        task._inFlightFiles[found].completedBytes = currentItemCompletedBytes;
        task._inFlightFiles[found].lastUpdateTick = nowTick;
        if (discoveryOpenAtCallback && currentItemTotalBytes > 0u && currentItemCompletedBytes >= currentItemTotalBytes &&
            ! task._inFlightFiles[found].completionCountedWhileDiscoveryOpen)
        {
            task._inFlightFiles[found].completionCountedWhileDiscoveryOpen = true;
            SaturatingAtomicAdd(task._discoveryCompletedMutationsWhileOpen, 1u);
        }
        return;
    }

    Task::InFlightFileProgress added{};
    added.cookieKey        = cookieKey;
    added.progressStreamId = streamKey;
    added.sourcePath       = currentSourcePath;
    added.totalBytes       = currentItemTotalBytes;
    added.completedBytes   = currentItemCompletedBytes;
    added.lastUpdateTick   = nowTick;
    if (discoveryOpenAtCallback && currentItemTotalBytes > 0u && currentItemCompletedBytes >= currentItemTotalBytes)
    {
        added.completionCountedWhileDiscoveryOpen = true;
        SaturatingAtomicAdd(task._discoveryCompletedMutationsWhileOpen, 1u);
    }

    if (task._inFlightFileCount < task._inFlightFiles.size())
    {
        task._inFlightFiles[task._inFlightFileCount] = std::move(added);
        ++task._inFlightFileCount;
        return;
    }

    if (! task._inFlightFiles.empty())
    {
        size_t replaceIndex  = 0;
        ULONGLONG oldestTick = task._inFlightFiles[0].lastUpdateTick;
        for (size_t i = 1; i < task._inFlightFileCount; ++i)
        {
            const ULONGLONG tick = task._inFlightFiles[i].lastUpdateTick;
            if (tick == 0 || (oldestTick != 0 && tick < oldestTick))
            {
                replaceIndex = i;
                oldestTick   = tick;
            }
        }

        ++task._perf.progressInFlightEvictions;
        task._inFlightFiles[replaceIndex] = std::move(added);
    }
}

void RemoveInFlightFileBySourcePath(Task& task, const wchar_t* sourcePath) noexcept
{
    if (sourcePath == nullptr || sourcePath[0] == L'\0')
    {
        return;
    }

    std::scoped_lock lock(task._inFlightFilesMutex);
    for (size_t i = 0; i < task._inFlightFileCount; ++i)
    {
        if (task._inFlightFiles[i].sourcePath == sourcePath)
        {
            for (size_t j = i + 1u; j < task._inFlightFileCount; ++j)
            {
                task._inFlightFiles[j - 1u] = std::move(task._inFlightFiles[j]);
            }
            --task._inFlightFileCount;
            break;
        }
    }
}

Task::ConflictWorkerPerf& FindOrAddConflictWorkerPerfLocked(Task& task, const void* cookieKey, ULONGLONG nowTick) noexcept
{
    for (size_t i = 0; i < task._conflictWorkerPerfCount; ++i)
    {
        auto& entry = task._conflictWorkerPerf[i];
        if (entry.cookieKey == cookieKey)
        {
            entry.lastUpdateTick = nowTick;
            return entry;
        }
    }

    size_t index = task._conflictWorkerPerfCount;
    if (index < task._conflictWorkerPerf.size())
    {
        ++task._conflictWorkerPerfCount;
    }
    else
    {
        index                = 0;
        ULONGLONG oldestTick = task._conflictWorkerPerf[0].lastUpdateTick;
        for (size_t i = 1; i < task._conflictWorkerPerfCount; ++i)
        {
            const ULONGLONG tick = task._conflictWorkerPerf[i].lastUpdateTick;
            if (tick == 0 || (oldestTick != 0 && tick < oldestTick))
            {
                index      = i;
                oldestTick = tick;
            }
        }
    }

    auto& entry          = task._conflictWorkerPerf[index];
    entry.cookieKey      = cookieKey;
    entry.promptCount    = 0;
    entry.waitUs         = 0;
    entry.lastUpdateTick = nowTick;
    return entry;
}

void NoteConflictWorkerWait(Task& task, const void* cookieKey, uint64_t waitUs) noexcept
{
    const ULONGLONG nowTick = GetTickCount64();
    std::scoped_lock lock(task._conflictArbiter.mutex);
    auto& entry = FindOrAddConflictWorkerPerfLocked(task, cookieKey, nowTick);
    ++entry.promptCount;
    entry.waitUs += waitUs;
}

[[nodiscard]] bool TryApplyDiagnosticPathSnapshot(std::wstring& resolvedPath, std::wstring_view fallbackPath, std::wstring_view candidatePath) noexcept
{
    if (candidatePath.empty())
    {
        return false;
    }

    if (! fallbackPath.empty() && ! IsSameOrChildPath(fallbackPath, candidatePath))
    {
        return false;
    }

    resolvedPath.assign(candidatePath);
    return true;
}

[[nodiscard]] std::pair<std::wstring, std::wstring> GetMostSpecificPathsForDiagnostics(
    const FolderWindow::FileOperationState::Task& task,
    const FolderWindow::FileOperationState::Task::PerItemCallbackCookie* perItemCookie,
    std::wstring_view sourceFallback,
    std::wstring_view destinationFallback) noexcept
{
    std::wstring source(sourceFallback);
    std::wstring destination(destinationFallback);

    bool sourceResolved      = false;
    bool destinationResolved = false;
    if (perItemCookie != nullptr)
    {
        sourceResolved      = TryApplyDiagnosticPathSnapshot(source, sourceFallback, perItemCookie->lastProgressSourcePath);
        destinationResolved = TryApplyDiagnosticPathSnapshot(destination, destinationFallback, perItemCookie->lastProgressDestinationPath);
    }

    // Conflict resolution now converges workers at checkpoints, but diagnostics should still
    // read the last published snapshot instead of taking _progressMutex again here.
    const auto publishedSnapshot = task._publishedDiagnosticPathSnapshot.load(std::memory_order_acquire);
    if (publishedSnapshot)
    {
        if (! sourceResolved)
        {
            sourceResolved = TryApplyDiagnosticPathSnapshot(source, sourceFallback, publishedSnapshot->lastProgressCallbackSourcePath);
        }
        if (! sourceResolved)
        {
            sourceResolved = TryApplyDiagnosticPathSnapshot(source, sourceFallback, publishedSnapshot->progressSourcePath);
        }

        if (! destinationResolved)
        {
            destinationResolved = TryApplyDiagnosticPathSnapshot(destination, destinationFallback, publishedSnapshot->lastProgressCallbackDestinationPath);
        }
        if (! destinationResolved)
        {
            destinationResolved = TryApplyDiagnosticPathSnapshot(destination, destinationFallback, publishedSnapshot->progressDestinationPath);
        }
    }

    return {std::move(source), std::move(destination)};
}

[[nodiscard]] size_t GetPositiveSizeOrDefault(const std::optional<uint32_t>& value, size_t defaultValue) noexcept
{
    if (! value.has_value() || value.value() == 0)
    {
        return defaultValue;
    }

    return static_cast<size_t>(value.value());
}

[[nodiscard]] ULONGLONG GetPositiveIntervalOrDefault(const std::optional<uint32_t>& value, ULONGLONG defaultValue) noexcept
{
    if (! value.has_value() || value.value() == 0)
    {
        return defaultValue;
    }

    return static_cast<ULONGLONG>(value.value());
}

void CleanupDiagnosticsFilesInDirectory(const std::filesystem::path& directory,
                                        std::wstring_view filePrefix,
                                        std::wstring_view fileExtension,
                                        size_t maxFilesToKeep) noexcept
{
    if (directory.empty() || maxFilesToKeep == 0)
    {
        return;
    }

    struct DiagnosticsFileForCleanup final
    {
        std::filesystem::path path;
        std::filesystem::file_time_type lastWriteTime{};
    };

    std::error_code ec;
    std::vector<DiagnosticsFileForCleanup> files;
    for (std::filesystem::directory_iterator it(directory, ec), end; ! ec && it != end; it.increment(ec))
    {
        const std::filesystem::directory_entry& de = *it;
        if (! de.is_regular_file(ec))
        {
            continue;
        }

        const std::wstring fileNameText = de.path().filename().wstring();
        if (fileNameText.size() < (filePrefix.size() + fileExtension.size()))
        {
            continue;
        }
        if (fileNameText.rfind(filePrefix.data(), 0) != 0)
        {
            continue;
        }
        if (de.path().extension().wstring() != fileExtension)
        {
            continue;
        }

        std::error_code timeEc;
        const std::filesystem::file_time_type lastWriteTime = de.last_write_time(timeEc);
        files.push_back(DiagnosticsFileForCleanup{
            .path          = de.path(),
            .lastWriteTime = timeEc ? std::filesystem::file_time_type::min() : lastWriteTime,
        });
    }

    if (files.size() <= maxFilesToKeep)
    {
        return;
    }

    std::sort(files.begin(),
              files.end(),
              [](const DiagnosticsFileForCleanup& left, const DiagnosticsFileForCleanup& right)
    {
        if (left.lastWriteTime != right.lastWriteTime)
        {
            return left.lastWriteTime > right.lastWriteTime;
        }
        return left.path > right.path;
    });
    for (size_t i = maxFilesToKeep; i < files.size(); ++i)
    {
        std::filesystem::remove(files[i].path, ec);
    }
}

[[nodiscard]] bool GetAutoDismissSuccessFromSettings(const Common::Settings::Settings& settings) noexcept
{
    if (! settings.fileOperations.has_value())
    {
        return false;
    }

    return settings.fileOperations->autoDismissSuccess;
}

[[nodiscard]] bool GetPopupFooterOnlyFromSettings(const Common::Settings::Settings& settings) noexcept
{
    if (! settings.fileOperations.has_value())
    {
        return false;
    }

    return settings.fileOperations->popupFooterOnly;
}

[[nodiscard]] bool GetPopupCompactDensityFromSettings(const Common::Settings::Settings& settings) noexcept
{
    if (! settings.fileOperations.has_value())
    {
        return false;
    }

    return settings.fileOperations->popupCompactDensity;
}

void PruneFileOperationsSettingsIfDefault(Common::Settings::Settings& settings) noexcept
{
    if (settings.fileOperations.has_value() && ! Common::Settings::HasNonDefaultFileOperationsSettings(settings.fileOperations.value()))
    {
        settings.fileOperations.reset();
    }
}

constexpr unsigned int kDefaultCrossFsBridgeBufferSizeKB = 4096u;
constexpr unsigned int kMinCrossFsBridgeBufferSizeKB     = 512u;
constexpr unsigned int kMaxCrossFsBridgeBufferSizeKB     = 16384u;
constexpr uint64_t kCrossFsBridgeBufferBudgetBytes       = 256ull * 1024ull * 1024ull;

class CrossFsBridgeBufferBudget final
{
public:
    CrossFsBridgeBufferBudget()                                            = default;
    ~CrossFsBridgeBufferBudget()                                           = default;
    CrossFsBridgeBufferBudget(const CrossFsBridgeBufferBudget&)            = delete;
    CrossFsBridgeBufferBudget& operator=(const CrossFsBridgeBufferBudget&) = delete;
    CrossFsBridgeBufferBudget(CrossFsBridgeBufferBudget&&)                 = delete;
    CrossFsBridgeBufferBudget& operator=(CrossFsBridgeBufferBudget&&)      = delete;

    [[nodiscard]] bool Acquire(uint64_t bytes, const std::atomic<bool>& cancelled, std::stop_token stopToken) noexcept
    {
        if (bytes == 0u || bytes > kCrossFsBridgeBufferBudgetBytes)
        {
            return false;
        }

        std::unique_lock lock(_mutex);
        while (bytes > kCrossFsBridgeBufferBudgetBytes - _inUseBytes)
        {
            if (cancelled.load(std::memory_order_acquire) || stopToken.stop_requested())
            {
                return false;
            }
            _cv.wait_for(lock, std::chrono::milliseconds(50));
        }

        _inUseBytes += bytes;
        _peakBytes = (std::max)(_peakBytes, _inUseBytes);
        return true;
    }

    void Release(uint64_t bytes) noexcept
    {
        {
            std::scoped_lock lock(_mutex);
            _inUseBytes = bytes <= _inUseBytes ? _inUseBytes - bytes : 0u;
        }
        _cv.notify_all();
    }

    [[nodiscard]] uint64_t PeakBytes() const noexcept
    {
        std::scoped_lock lock(_mutex);
        return _peakBytes;
    }

    [[nodiscard]] uint64_t InUseBytes() const noexcept
    {
        std::scoped_lock lock(_mutex);
        return _inUseBytes;
    }

    void ResetPeak() noexcept
    {
        std::scoped_lock lock(_mutex);
        _peakBytes = _inUseBytes;
    }

private:
    mutable std::mutex _mutex;
    std::condition_variable _cv;
    uint64_t _inUseBytes = 0u;
    uint64_t _peakBytes  = 0u;
};

CrossFsBridgeBufferBudget& GetCrossFsBridgeBufferBudget() noexcept
{
    static CrossFsBridgeBufferBudget budget;
    return budget;
}

class CrossFsBridgeBufferLease final
{
public:
    CrossFsBridgeBufferLease() = default;
    ~CrossFsBridgeBufferLease() noexcept
    {
        Reset();
    }

    CrossFsBridgeBufferLease(const CrossFsBridgeBufferLease&)            = delete;
    CrossFsBridgeBufferLease& operator=(const CrossFsBridgeBufferLease&) = delete;
    CrossFsBridgeBufferLease(CrossFsBridgeBufferLease&&)                 = delete;
    CrossFsBridgeBufferLease& operator=(CrossFsBridgeBufferLease&&)      = delete;

    [[nodiscard]] bool Acquire(uint64_t bytes, const std::atomic<bool>& cancelled, std::stop_token stopToken) noexcept
    {
        if (_bytes != 0u || ! GetCrossFsBridgeBufferBudget().Acquire(bytes, cancelled, stopToken))
        {
            return false;
        }
        _bytes = bytes;
        return true;
    }

    void Reset() noexcept
    {
        if (_bytes != 0u)
        {
            GetCrossFsBridgeBufferBudget().Release(_bytes);
            _bytes = 0u;
        }
    }

private:
    uint64_t _bytes = 0u;
};
constexpr uint64_t kDefaultBandwidthLimitBytesPerSecond = 0;
[[nodiscard]] uint64_t GetFileOpsTraversalMaxDepth() noexcept
{
#ifdef ENABLE_TESTS
    const uint64_t overrideLimit = g_fileOpsBridgeTraversalDepthLimitOverride.load(std::memory_order_acquire);
    if (overrideLimit != 0u)
    {
        return overrideLimit;
    }
#endif
    // R4-T1: the bridge walkers keep their frames on an explicit stack, so depth is reported through
    // bridge.traversal.* counters and is not a terminal limit; the retained work-entry, path-text and
    // metadata ceilings stay task-terminal.
    return (std::numeric_limits<uint64_t>::max)();
}

[[nodiscard]] size_t GetBridgeAdmissionQueueLimit() noexcept
{
    return Common::FileOperations::kDiscoveryTarget;
}

[[nodiscard]] HRESULT TryGetValidatedFileInfoName(FileInfo* entry, const std::byte* bufferBase, const std::byte* bufferEnd, std::wstring_view& nameOut) noexcept
{
    nameOut = {};
    if (entry == nullptr || bufferBase == nullptr || bufferEnd == nullptr || bufferEnd < bufferBase)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const auto* entryBytes = reinterpret_cast<const std::byte*>(entry);
    if (entryBytes < bufferBase || entryBytes > bufferEnd)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const size_t availableBytes  = static_cast<size_t>(bufferEnd - entryBytes);
    constexpr size_t kNameOffset = offsetof(FileInfo, FileName);
    if (availableBytes < kNameOffset)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if ((entry->FileNameSize % sizeof(wchar_t)) != 0u || entry->FileNameSize > availableBytes - kNameOffset)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    nameOut = std::wstring_view(entry->FileName, static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t));
    return S_OK;
}

[[nodiscard]] HRESULT ValidateBridgeStructuralChildName(const std::wstring_view name) noexcept
{
    if (name.empty() || name == L"." || name == L"..")
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    for (const wchar_t ch : name)
    {
        if (ch == L'\\' || ch == L'/' || ch == L'\0' || ch < L' ' || ch == static_cast<wchar_t>(0x7F))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
    }

    return S_OK;
}

[[nodiscard]] bool IsReservedWindowsBridgeChildName(std::wstring_view name) noexcept
{
    while (! name.empty() && (name.back() == L' ' || name.back() == L'.'))
    {
        name.remove_suffix(1u);
    }

    if (const size_t firstDot = name.find(L'.'); firstDot != std::wstring_view::npos)
    {
        name = name.substr(0u, firstDot);
    }
    while (! name.empty() && name.back() == L' ')
    {
        name.remove_suffix(1u);
    }

    if (OrdinalString::EqualsNoCase(name, L"CON") || OrdinalString::EqualsNoCase(name, L"PRN") || OrdinalString::EqualsNoCase(name, L"AUX") ||
        OrdinalString::EqualsNoCase(name, L"NUL") || OrdinalString::EqualsNoCase(name, L"CONIN$") || OrdinalString::EqualsNoCase(name, L"CONOUT$"))
    {
        return true;
    }

    if (name.size() != 4u || (name[3] < L'0' || name[3] > L'9'))
    {
        return false;
    }

    return OrdinalString::StartsWithNoCase(name, L"COM") || OrdinalString::StartsWithNoCase(name, L"LPT");
}

[[nodiscard]] HRESULT ValidateWindowsBridgeChildName(const std::wstring_view name) noexcept
{
    constexpr size_t kMaximumWindowsComponentLength = 255u;
    if (name.empty() || name.size() > kMaximumWindowsComponentLength || name.back() == L' ' || name.back() == L'.' ||
        name.find_first_of(L":*?\"<>|") != std::wstring_view::npos || IsReservedWindowsBridgeChildName(name))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    return S_OK;
}

[[nodiscard]] bool UsesOrdinalIgnoreCaseComponents(IFileSystem& fileSystem, std::wstring_view providerPath, std::wstring_view pluginId) noexcept
{
    const FileSystemRouteContract::QueryResult route = FileSystemRouteContract::Query(&fileSystem, providerPath, FILESYSTEM_COPY, pluginId);
    return route.state == FileSystemRouteContract::QueryState::Available && route.snapshot.pathIdentity.has_value() &&
           route.snapshot.pathIdentity->pathTextStableIdentity &&
           route.snapshot.pathIdentity->componentComparison == FileSystemPathComponentComparison::OrdinalIgnoreCase;
}

// Orders one directory's enumerated child names the way the destination compares components:
// ordinal, or ordinal-ignore-case when the destination folds case. The names are views into the
// frame's own IFilesInformation buffer, so registering a name retains no copy (R4-T3).
struct BridgeChildNameLess final
{
    bool ignoreCase = false;

    [[nodiscard]] bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
    {
        return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), ignoreCase ? TRUE : FALSE) ==
               CSTR_LESS_THAN;
    }
};
using BridgeChildNameSet = std::set<std::wstring_view, BridgeChildNameLess>;

[[nodiscard]] HRESULT AdvanceValidatedFileInfoEntry(FileInfo* entry, const std::byte* bufferBase, const std::byte* bufferEnd, FileInfo*& nextOut) noexcept
{
    nextOut = nullptr;
    if (entry == nullptr || bufferBase == nullptr || bufferEnd == nullptr || bufferEnd < bufferBase)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (entry->NextEntryOffset == 0)
    {
        return S_FALSE;
    }

    const auto* entryBytes = reinterpret_cast<const std::byte*>(entry);
    if (entryBytes < bufferBase || entryBytes > bufferEnd)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const size_t availableBytes = static_cast<size_t>(bufferEnd - entryBytes);
    if (entry->NextEntryOffset < sizeof(FileInfo) || entry->NextEntryOffset > availableBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const auto* nextBytes = entryBytes + entry->NextEntryOffset;
    if (nextBytes < bufferBase || nextBytes > bufferEnd)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    const size_t remainingBytes = static_cast<size_t>(bufferEnd - nextBytes);
    if (remainingBytes < sizeof(FileInfo))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    nextOut = reinterpret_cast<FileInfo*>(const_cast<std::byte*>(nextBytes));
    return S_OK;
}

[[nodiscard]] unsigned long GetCrossFsBridgeBufferBytesFromSettings(const Common::Settings::Settings* settings) noexcept
{
    uint32_t bufferSizeKB = kDefaultCrossFsBridgeBufferSizeKB;
    if (settings && settings->fileOperations.has_value())
    {
        bufferSizeKB = settings->fileOperations->crossFsBridgeBufferSizeKB;
    }

    bufferSizeKB                    = std::clamp(bufferSizeKB, kMinCrossFsBridgeBufferSizeKB, kMaxCrossFsBridgeBufferSizeKB);
    constexpr uint64_t kBytesPerKiB = 1024ull;
    const uint64_t bytes64          = static_cast<uint64_t>(bufferSizeKB) * kBytesPerKiB;
    return bytes64 > static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()) ? std::numeric_limits<unsigned long>::max()
                                                                                      : static_cast<unsigned long>(bytes64);
}

[[nodiscard]] unsigned long ClampCrossFsBridgeBufferBytes(uint32_t preferredBytes) noexcept
{
    constexpr uint64_t kMinBytes = static_cast<uint64_t>(kMinCrossFsBridgeBufferSizeKB) * 1024ull;
    constexpr uint64_t kMaxBytes = static_cast<uint64_t>(kMaxCrossFsBridgeBufferSizeKB) * 1024ull;
    const uint64_t clampedBytes  = (std::clamp)(static_cast<uint64_t>(preferredBytes), kMinBytes, kMaxBytes);
    return static_cast<unsigned long>((std::min)(clampedBytes, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max())));
}

struct AdaptiveBridgeTuning final
{
    unsigned long bufferBytes = 0u;
    DWORD progressPeriodMs    = 200u;
    uint32_t latencyClass     = FILESYSTEM_TRANSFER_LATENCY_UNKNOWN;
    uint32_t flags            = FILESYSTEM_TRANSFER_HINT_NONE;
};

[[nodiscard]] AdaptiveBridgeTuning ResolveAdaptiveCrossFsBridgeTuning(unsigned long configuredBytes,
                                                                      IFileSystem& sourceFileSystem,
                                                                      const wchar_t* sourcePath,
                                                                      IFileSystem& destinationFileSystem,
                                                                      const wchar_t* destinationPath,
                                                                      FileSystemOperation operationType) noexcept
{
    const unsigned long defaultBytes = kDefaultCrossFsBridgeBufferSizeKB * 1024u;
    if (configuredBytes == 0)
    {
        configuredBytes = defaultBytes;
    }
    AdaptiveBridgeTuning tuning{.bufferBytes = configuredBytes};
    if (sourcePath == nullptr || destinationPath == nullptr)
    {
        return tuning;
    }

    unsigned long resolvedBytes = 0;
    bool sawBufferHint          = false;
    const auto applyHints       = [&](IFileSystem& fileSystem, const wchar_t* path, FileSystemTransferEndpoint endpoint) noexcept
    {
        FileSystemTransferHints hints{};
        hints.sizeBytes  = sizeof(hints);
        const HRESULT hr = fileSystem.GetTransferHints(path, operationType, endpoint, &hints);
        if (FAILED(hr))
        {
            return;
        }

        tuning.latencyClass = (std::max)(tuning.latencyClass, hints.latencyClass);
        tuning.flags |= hints.flags;
        if (hints.preferredProgressPeriodMs != 0u)
        {
            tuning.progressPeriodMs = (std::max)(tuning.progressPeriodMs, static_cast<DWORD>(hints.preferredProgressPeriodMs));
        }
        if (hints.preferredBufferBytes == 0u)
        {
            return;
        }

        const unsigned long hintedBytes = ClampCrossFsBridgeBufferBytes(hints.preferredBufferBytes);
        sawBufferHint                   = true;
        resolvedBytes                   = (std::max)(resolvedBytes, hintedBytes);
    };

    applyHints(sourceFileSystem, sourcePath, FILESYSTEM_TRANSFER_SOURCE_READ);
    applyHints(destinationFileSystem, destinationPath, FILESYSTEM_TRANSFER_DESTINATION_WRITE);
    if (sawBufferHint)
    {
        // The setting is a fallback. Provider hints remain active even when the user changed it;
        // this avoids silently disabling WAN/cloud tuning for every non-default preference value.
        tuning.bufferBytes = resolvedBytes;
    }
    else if (tuning.latencyClass >= FILESYSTEM_TRANSFER_LATENCY_WAN || (tuning.flags & FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS) != 0u)
    {
        tuning.bufferBytes = (std::max)(configuredBytes, ClampCrossFsBridgeBufferBytes(8u * 1024u * 1024u));
    }
    return tuning;
}

[[nodiscard]] uint64_t GetDefaultBandwidthLimitBytesPerSecondFromSettings(const Common::Settings::Settings* settings) noexcept
{
    if (! settings || ! settings->fileOperations.has_value())
    {
        return kDefaultBandwidthLimitBytesPerSecond;
    }

    return settings->fileOperations->defaultBandwidthLimitBytesPerSecond;
}

void SetAutoDismissSuccessInSettings(Common::Settings::Settings& settings, bool enabled) noexcept
{
    if (settings.fileOperations.has_value())
    {
        settings.fileOperations->autoDismissSuccess = enabled;
    }
    else if (enabled)
    {
        settings.fileOperations.emplace();
        settings.fileOperations->autoDismissSuccess = true;
    }

    PruneFileOperationsSettingsIfDefault(settings);
}

void SetPopupFooterOnlyInSettings(Common::Settings::Settings& settings, bool footerOnly) noexcept
{
    if (settings.fileOperations.has_value())
    {
        settings.fileOperations->popupFooterOnly = footerOnly;
    }
    else if (footerOnly)
    {
        settings.fileOperations.emplace();
        settings.fileOperations->popupFooterOnly = true;
    }

    PruneFileOperationsSettingsIfDefault(settings);
}

void SetPopupCompactDensityInSettings(Common::Settings::Settings& settings, bool compactDensity) noexcept
{
    if (settings.fileOperations.has_value())
    {
        settings.fileOperations->popupCompactDensity = compactDensity;
    }
    else if (compactDensity)
    {
        settings.fileOperations.emplace();
        settings.fileOperations->popupCompactDensity = true;
    }

    PruneFileOperationsSettingsIfDefault(settings);
}

[[nodiscard]] DiagnosticsSettings GetDiagnosticsSettingsFromSettings(const Common::Settings::Settings* settings) noexcept
{
    DiagnosticsSettings diagnostics{};
    if (! settings || ! settings->fileOperations.has_value())
    {
        return diagnostics;
    }

    const auto& fileOperations                 = settings->fileOperations.value();
    diagnostics.maxDiagnosticsInMemory         = GetPositiveSizeOrDefault(fileOperations.maxDiagnosticsInMemory, diagnostics.maxDiagnosticsInMemory);
    diagnostics.maxDiagnosticsPerFlush         = GetPositiveSizeOrDefault(fileOperations.maxDiagnosticsPerFlush, diagnostics.maxDiagnosticsPerFlush);
    diagnostics.maxDiagnosticsLogFiles         = std::max<size_t>(1u, static_cast<size_t>(fileOperations.maxDiagnosticsLogFiles));
    diagnostics.maxDiagnosticsIssueReportFiles = GetPositiveSizeOrDefault(fileOperations.maxIssueReportFiles, diagnostics.maxDiagnosticsIssueReportFiles);
    diagnostics.diagnosticsFlushIntervalMs = GetPositiveIntervalOrDefault(fileOperations.diagnosticsFlushIntervalMs, diagnostics.diagnosticsFlushIntervalMs);
    diagnostics.diagnosticsCleanupIntervalMs =
        GetPositiveIntervalOrDefault(fileOperations.diagnosticsCleanupIntervalMs, diagnostics.diagnosticsCleanupIntervalMs);
    diagnostics.infoEnabled  = fileOperations.diagnosticsInfoEnabled;
    diagnostics.debugEnabled = fileOperations.diagnosticsDebugEnabled;
    return diagnostics;
}

[[nodiscard]] const wchar_t* OperationToString(FileSystemOperation operation) noexcept
{
    switch (operation)
    {
        case FILESYSTEM_COPY: return L"copy";
        case FILESYSTEM_MOVE: return L"move";
        case FILESYSTEM_DELETE: return L"delete";
        case FILESYSTEM_RENAME: return L"rename";
        case FILESYSTEM_CREATE_DIRECTORY: return L"create-directory";
        default: return L"unknown";
    }
}

[[nodiscard]] bool IsCancellationStatus(HRESULT hr) noexcept
{
    return hr == E_ABORT || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

[[nodiscard]] const wchar_t* DiagnosticSeverityToString(FolderWindow::FileOperationState::DiagnosticSeverity severity) noexcept
{
    switch (severity)
    {
        case FolderWindow::FileOperationState::DiagnosticSeverity::Debug: return L"debug";
        case FolderWindow::FileOperationState::DiagnosticSeverity::Info: return L"info";
        case FolderWindow::FileOperationState::DiagnosticSeverity::Warning: return L"warning";
        case FolderWindow::FileOperationState::DiagnosticSeverity::Error: return L"error";
        default: return L"unknown";
    }
}

[[nodiscard]] ProcessMemorySnapshot CaptureProcessMemorySnapshot() noexcept
{
    ProcessMemorySnapshot snapshot{};

    PROCESS_MEMORY_COUNTERS_EX counters{};
    counters.cb = sizeof(counters);
    if (K32GetProcessMemoryInfo(GetCurrentProcess(), reinterpret_cast<PROCESS_MEMORY_COUNTERS*>(&counters), static_cast<DWORD>(sizeof(counters))) == 0)
    {
        return snapshot;
    }

    snapshot.workingSetBytes = static_cast<uint64_t>(counters.WorkingSetSize);
    snapshot.privateBytes    = static_cast<uint64_t>(counters.PrivateUsage);
    return snapshot;
}

[[nodiscard]] const wchar_t* Win32ErrorToSymbolicName(DWORD error) noexcept
{
    switch (error)
    {
        case ERROR_SUCCESS: return L"ERROR_SUCCESS";
        case ERROR_ACCESS_DENIED: return L"ERROR_ACCESS_DENIED";
        case ERROR_ALREADY_EXISTS: return L"ERROR_ALREADY_EXISTS";
        case ERROR_FILE_EXISTS: return L"ERROR_FILE_EXISTS";
        case ERROR_FILE_NOT_FOUND: return L"ERROR_FILE_NOT_FOUND";
        case ERROR_PATH_NOT_FOUND: return L"ERROR_PATH_NOT_FOUND";
        case ERROR_SHARING_VIOLATION: return L"ERROR_SHARING_VIOLATION";
        case ERROR_LOCK_VIOLATION: return L"ERROR_LOCK_VIOLATION";
        case ERROR_DISK_FULL: return L"ERROR_DISK_FULL";
        case ERROR_HANDLE_DISK_FULL: return L"ERROR_HANDLE_DISK_FULL";
        case ERROR_CANCELLED: return L"ERROR_CANCELLED";
        case ERROR_NOT_SUPPORTED: return L"ERROR_NOT_SUPPORTED";
        case ERROR_INVALID_NAME: return L"ERROR_INVALID_NAME";
        case ERROR_INVALID_PARAMETER: return L"ERROR_INVALID_PARAMETER";
        case ERROR_DIRECTORY: return L"ERROR_DIRECTORY";
        case ERROR_PARTIAL_COPY: return L"ERROR_PARTIAL_COPY";
        case ERROR_BAD_LENGTH: return L"ERROR_BAD_LENGTH";
        case ERROR_ARITHMETIC_OVERFLOW: return L"ERROR_ARITHMETIC_OVERFLOW";
        default: return nullptr;
    }
}

[[nodiscard]] std::wstring FormatDiagnosticHresultName(HRESULT hr) noexcept
{
    const wchar_t* known = nullptr;
    switch (hr)
    {
        case S_OK: known = L"S_OK"; break;
        case S_FALSE: known = L"S_FALSE"; break;
        case E_ABORT: known = L"E_ABORT"; break;
        case E_ACCESSDENIED: known = L"E_ACCESSDENIED"; break;
        case E_FAIL: known = L"E_FAIL"; break;
        case E_INVALIDARG: known = L"E_INVALIDARG"; break;
        case E_NOINTERFACE: known = L"E_NOINTERFACE"; break;
        case E_NOTIMPL: known = L"E_NOTIMPL"; break;
        case E_OUTOFMEMORY: known = L"E_OUTOFMEMORY"; break;
        case E_POINTER: known = L"E_POINTER"; break;
        case E_UNEXPECTED: known = L"E_UNEXPECTED"; break;
        default: break;
    }
    if (known)
    {
        return known;
    }

    if (HRESULT_FACILITY(hr) == FACILITY_WIN32)
    {
        const DWORD code = HRESULT_CODE(static_cast<DWORD>(hr));
        if (const wchar_t* win32Name = Win32ErrorToSymbolicName(code))
        {
            return win32Name;
        }

        return std::format(L"WIN32_ERROR_{}", static_cast<unsigned long>(code));
    }

    return std::format(L"HRESULT_0x{:08X}", static_cast<unsigned long>(hr));
}

[[nodiscard]] std::wstring FormatDiagnosticStatusText(HRESULT hr) noexcept
{
    return FormatHResultMessage(hr);
}

[[nodiscard]] std::wstring EscapeDiagnosticField(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    std::wstring escaped;
    escaped.reserve(text.size());
    for (wchar_t ch : text)
    {
        if (ch == L'\r' || ch == L'\n' || ch == L'\t')
        {
            escaped.push_back(L' ');
        }
        else
        {
            escaped.push_back(ch);
        }
    }

    return escaped;
}

[[nodiscard]] std::wstring EscapeDiagnosticJsonString(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    std::wstring escaped;
    escaped.reserve(text.size());
    for (wchar_t ch : text)
    {
        switch (ch)
        {
            case L'\\': escaped.append(L"\\\\"); break;
            case L'"': escaped.append(L"\\\""); break;
            case L'\b': escaped.append(L"\\b"); break;
            case L'\f': escaped.append(L"\\f"); break;
            case L'\n': escaped.append(L"\\n"); break;
            case L'\r': escaped.append(L"\\r"); break;
            case L'\t': escaped.append(L"\\t"); break;
            default:
                if (ch < 0x20)
                {
                    std::format_to(std::back_inserter(escaped), L"\\u{:04X}", static_cast<unsigned>(ch));
                }
                else
                {
                    escaped.push_back(ch);
                }
                break;
        }
    }

    return escaped;
}

[[nodiscard]] std::wstring_view TrimTrailingSeparators(std::wstring_view path) noexcept
{
    while (! path.empty())
    {
        const wchar_t last = path.back();
        if (last != L'\\' && last != L'/')
        {
            break;
        }
        path.remove_suffix(1);
    }
    return path;
}

[[nodiscard]] bool IsSameOrChildPath(std::wstring_view root, std::wstring_view candidate) noexcept
{
    root      = TrimTrailingSeparators(root);
    candidate = TrimTrailingSeparators(candidate);

    if (root.empty() || candidate.size() < root.size())
    {
        return false;
    }

    if (root.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }

    if (! OrdinalString::StartsWithNoCase(candidate, root))
    {
        return false;
    }

    if (candidate.size() == root.size())
    {
        return true;
    }

    const wchar_t next = candidate[root.size()];
    return next == L'\\' || next == L'/';
}

[[nodiscard]] std::wstring_view GetPathLeaf(std::wstring_view path) noexcept
{
    const std::wstring_view trimmed = TrimTrailingSeparators(path);
    if (trimmed.empty())
    {
        return trimmed;
    }

    const size_t pos = trimmed.find_last_of(L"\\/");
    if (pos == std::wstring_view::npos)
    {
        return trimmed;
    }

    return trimmed.substr(pos + 1);
}

[[nodiscard]] wchar_t GuessPreferredSeparator(std::wstring_view folder) noexcept
{
    const bool hasForward = folder.find(L'/') != std::wstring_view::npos;
    const bool hasBack    = folder.find(L'\\') != std::wstring_view::npos;
    if (hasForward && ! hasBack)
    {
        return L'/';
    }
    return L'\\';
}

[[nodiscard]] std::wstring JoinFolderAndLeaf(std::wstring_view folder, std::wstring_view leaf) noexcept
{
    if (folder.empty())
    {
        return std::wstring(leaf);
    }

    std::wstring result(folder);
    const wchar_t sep = GuessPreferredSeparator(folder);
    if (! result.empty())
    {
        const wchar_t last = result.back();
        if (last != L'\\' && last != L'/')
        {
            result.push_back(sep);
        }
    }
    result.append(leaf);
    return result;
}

[[nodiscard]] unsigned int DetermineConfiguredPerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                    std::wstring_view providerPath,
                                                                    std::wstring_view pluginId,
                                                                    FileSystemOperation operation,
                                                                    FileSystemFlags flags,
                                                                    unsigned int uiMax) noexcept
{
    if (! fileSystem || providerPath.empty() || uiMax == 0u)
    {
        return 1u;
    }

    const bool isCopyMove = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
    const bool isDelete   = operation == FILESYSTEM_DELETE;
    if (! isCopyMove && ! isDelete)
    {
        return 1u;
    }

    const FileSystemRouteContract::QueryResult route = FileSystemRouteContract::Query(fileSystem.get(), providerPath, operation, pluginId);
    if (route.state != FileSystemRouteContract::QueryState::Available)
    {
        return 1u;
    }

    uint32_t concurrency = 1u;
    if (isCopyMove)
    {
        concurrency = route.snapshot.copyMoveMaxConcurrency;
    }
    else if (isDelete)
    {
        concurrency = (flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0 ? route.snapshot.deleteRecycleBinMaxConcurrency : route.snapshot.deleteMaxConcurrency;
    }

    return std::clamp(concurrency, 1u, uiMax);
}

[[nodiscard]] std::optional<AutoConcurrencyResolution> TryGetStoragePreferredMaxConcurrencyForPath(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                                                   const std::filesystem::path& path,
                                                                                                   FileSystemOperation operation) noexcept
{
    if (! fileSystem || path.empty())
    {
        return std::nullopt;
    }

#ifdef ENABLE_TESTS
    if (g_fileOpsAutoConcurrencyOverrideEnabled.load(std::memory_order_acquire))
    {
        AutoConcurrencyResolution resolution{};
        resolution.concurrency = std::max(1u, g_fileOpsAutoConcurrencyOverridePreferred.load(std::memory_order_acquire));
        resolution.storageKind = g_fileOpsAutoConcurrencyOverrideStorageKind.load(std::memory_order_acquire);
        return resolution;
    }
#endif

    FileSystemStorageCharacteristics characteristics{};
    characteristics.sizeBytes = sizeof(FileSystemStorageCharacteristics);
    if (FAILED(fileSystem->GetStorageCharacteristics(path.c_str(), &characteristics)))
    {
        return std::nullopt;
    }

    unsigned int preferred = 0u;
    if (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE)
    {
        preferred = characteristics.preferredCopyMoveConcurrency;
    }
    else if (operation == FILESYSTEM_DELETE)
    {
        preferred = characteristics.preferredDeleteConcurrency;
    }

    if (preferred == 0u)
    {
        return std::nullopt;
    }

    AutoConcurrencyResolution resolution{};
    resolution.concurrency = preferred;
    resolution.storageKind = characteristics.storageKind;
    return resolution;
}

[[nodiscard]] unsigned int DetermineAutoPerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                              const std::vector<std::filesystem::path>& paths,
                                                              FileSystemOperation operation,
                                                              unsigned int uiMax) noexcept
{
    AutoConcurrencyResolution resolution{};
    if (uiMax == 0u)
    {
        return 0u;
    }

    for (const auto& path : paths)
    {
        const auto preferred = TryGetStoragePreferredMaxConcurrencyForPath(fileSystem, path, operation);
        if (! preferred.has_value())
        {
            continue;
        }

        const unsigned int clamped = std::clamp(preferred->concurrency, 1u, uiMax);
        if (! resolution.HasValue() || clamped < resolution.concurrency)
        {
            resolution.concurrency = clamped;
            resolution.storageKind = preferred->storageKind;
            continue;
        }

        if (clamped == resolution.concurrency && resolution.storageKind != preferred->storageKind)
        {
            resolution.storageKind = FILESYSTEM_STORAGE_UNKNOWN;
        }
    }

    return resolution.concurrency;
}

[[nodiscard]] AutoConcurrencyResolution ResolveAutoPerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                         const std::vector<std::filesystem::path>& paths,
                                                                         FileSystemOperation operation,
                                                                         unsigned int uiMax) noexcept
{
    if (! fileSystem || uiMax == 0u)
    {
        return {};
    }

    AutoConcurrencyResolution resolution{};
    for (const auto& path : paths)
    {
        if (const auto preferred = TryGetStoragePreferredMaxConcurrencyForPath(fileSystem, path, operation); preferred.has_value())
        {
            const unsigned int clamped = std::clamp(preferred->concurrency, 1u, uiMax);
            if (! resolution.HasValue() || clamped < resolution.concurrency)
            {
                resolution.concurrency = clamped;
                resolution.storageKind = preferred->storageKind;
            }
            else if (clamped == resolution.concurrency && resolution.storageKind != preferred->storageKind)
            {
                resolution.storageKind = FILESYSTEM_STORAGE_UNKNOWN;
            }
        }
    }

    return resolution;
}

[[nodiscard]] unsigned int DetermineAutoPerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                              const std::filesystem::path& path,
                                                              FileSystemOperation operation,
                                                              unsigned int uiMax) noexcept
{
    return ResolveAutoPerItemMaxConcurrency(fileSystem, path, operation, uiMax).concurrency;
}

[[nodiscard]] AutoConcurrencyResolution ResolveAutoPerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                         const std::filesystem::path& path,
                                                                         FileSystemOperation operation,
                                                                         unsigned int uiMax) noexcept
{
    if (uiMax == 0u)
    {
        return {};
    }

    if (const auto preferred = TryGetStoragePreferredMaxConcurrencyForPath(fileSystem, path, operation); preferred.has_value())
    {
        AutoConcurrencyResolution resolution{};
        resolution.concurrency = std::clamp(preferred->concurrency, 1u, uiMax);
        resolution.storageKind = preferred->storageKind;
        return resolution;
    }

    return {};
}

[[nodiscard]] bool ShouldUseAutoPerItemConcurrency(const wil::com_ptr<IFileSystem>& fileSystem, FileSystemOperation operation, FileSystemFlags flags) noexcept
{
    const bool isCopyMove = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
    const bool isDelete   = operation == FILESYSTEM_DELETE;
    if (! isCopyMove && ! isDelete)
    {
        return false;
    }

    if (isDelete && (flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0)
    {
        // Recycle Bin deletes still use the explicit shell-oriented cap; storage hints do not model shell batching cost.
        return false;
    }

    const auto mode = TryGetConcurrencyModeFromFileSystem(fileSystem);
    return mode.has_value() && mode.value() == FileSystemConcurrencyMode::Auto;
}

[[nodiscard]] unsigned int DeterminePerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                          const std::vector<std::filesystem::path>& paths,
                                                          std::wstring_view pluginId,
                                                          FileSystemOperation operation,
                                                          FileSystemFlags flags,
                                                          unsigned int uiMax) noexcept
{
    if (ShouldUseAutoPerItemConcurrency(fileSystem, operation, flags))
    {
        if (const unsigned int autoConcurrency = DetermineAutoPerItemMaxConcurrency(fileSystem, paths, operation, uiMax); autoConcurrency > 0u)
        {
            return autoConcurrency;
        }
    }

    return paths.empty() ? 1u : DetermineConfiguredPerItemMaxConcurrency(fileSystem, paths.front().native(), pluginId, operation, flags, uiMax);
}

[[nodiscard]] unsigned int DeterminePerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                          const std::filesystem::path& path,
                                                          std::wstring_view pluginId,
                                                          FileSystemOperation operation,
                                                          FileSystemFlags flags,
                                                          unsigned int uiMax) noexcept
{
    if (ShouldUseAutoPerItemConcurrency(fileSystem, operation, flags))
    {
        if (const unsigned int autoConcurrency = DetermineAutoPerItemMaxConcurrency(fileSystem, path, operation, uiMax); autoConcurrency > 0u)
        {
            return autoConcurrency;
        }
    }

    return DetermineConfiguredPerItemMaxConcurrency(fileSystem, path.native(), pluginId, operation, flags, uiMax);
}

[[nodiscard]] std::wstring ResolveCircuitBreakerConnectionId(const Common::Settings::Settings* settings, std::wstring_view pluginPath) noexcept
{
    if (const auto connName = ConnectionProfileUtils::TryParseConnNameFromPluginPath(pluginPath); connName.has_value())
    {
        if (const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(settings, *connName);
            profile && ! profile->id.empty())
        {
            return profile->id;
        }
    }

    return {};
}

class ConnectionConcurrencyLimiter final
{
public:
    ConnectionConcurrencyLimiter()  = default;
    ~ConnectionConcurrencyLimiter() = default;

    ConnectionConcurrencyLimiter(const ConnectionConcurrencyLimiter&)            = delete;
    ConnectionConcurrencyLimiter& operator=(const ConnectionConcurrencyLimiter&) = delete;
    ConnectionConcurrencyLimiter(ConnectionConcurrencyLimiter&&)                 = delete;
    ConnectionConcurrencyLimiter& operator=(ConnectionConcurrencyLimiter&&)      = delete;

    enum class Kind : uint8_t
    {
        CopyMove,
        Delete,
    };

    class Permit final
    {
    public:
        Permit() = default;

        Permit(ConnectionConcurrencyLimiter* limiter, std::wstring connectionId, Kind kind) noexcept
            : _limiter(limiter),
              _connectionId(std::move(connectionId)),
              _kind(kind)
        {
        }

        Permit(const Permit&)            = delete;
        Permit& operator=(const Permit&) = delete;

        Permit(Permit&& other) noexcept : _limiter(std::exchange(other._limiter, nullptr)), _connectionId(std::move(other._connectionId)), _kind(other._kind)
        {
        }

        Permit& operator=(Permit&& other) noexcept
        {
            if (this == &other)
            {
                return *this;
            }

            Release();
            _limiter      = std::exchange(other._limiter, nullptr);
            _connectionId = std::move(other._connectionId);
            _kind         = other._kind;
            return *this;
        }

        ~Permit()
        {
            Release();
        }

        explicit operator bool() const noexcept
        {
            return _limiter != nullptr;
        }

    private:
        void Release() noexcept
        {
            if (! _limiter)
            {
                return;
            }

            _limiter->Release(_connectionId, _kind);
            _limiter = nullptr;
        }

        ConnectionConcurrencyLimiter* _limiter = nullptr;
        std::wstring _connectionId;
        Kind _kind = Kind::CopyMove;
    };

    template <typename CancelPredicate>
    [[nodiscard]] Permit AcquireCopyMove(std::wstring_view connectionId, uint32_t max, CancelPredicate&& shouldCancel) noexcept
    {
        return Acquire(connectionId, Kind::CopyMove, max, std::forward<CancelPredicate>(shouldCancel));
    }

    template <typename CancelPredicate>
    [[nodiscard]] Permit AcquireDelete(std::wstring_view connectionId, uint32_t max, CancelPredicate&& shouldCancel) noexcept
    {
        return Acquire(connectionId, Kind::Delete, max, std::forward<CancelPredicate>(shouldCancel));
    }

private:
    struct Entry final
    {
        uint32_t maxCopyMove      = 1;
        uint32_t inFlightCopyMove = 0;
        uint32_t maxDelete        = 1;
        uint32_t inFlightDelete   = 0;
    };

    void Release(const std::wstring& connectionId, Kind kind) noexcept
    {
        std::lock_guard lock(_mutex);

        const auto it = _entries.find(connectionId);
        if (it == _entries.end())
        {
            return;
        }

        Entry& entry = it->second;
        if (kind == Kind::CopyMove)
        {
            if (entry.inFlightCopyMove > 0)
            {
                --entry.inFlightCopyMove;
            }
        }
        else
        {
            if (entry.inFlightDelete > 0)
            {
                --entry.inFlightDelete;
            }
        }

        _cv.notify_all();
    }

    template <typename CancelPredicate>
    [[nodiscard]] Permit Acquire(std::wstring_view connectionId, Kind kind, uint32_t max, CancelPredicate&& shouldCancel) noexcept
    {
        if (connectionId.empty())
        {
            return {};
        }

        std::wstring key(connectionId);
        const uint32_t maxEffective = (std::max)(1u, max);

        std::unique_lock lock(_mutex);
        for (;;)
        {
            if (shouldCancel())
            {
                return {};
            }

            Entry& entry = _entries[key];
            if (kind == Kind::CopyMove)
            {
                entry.maxCopyMove = maxEffective;
                if (entry.inFlightCopyMove < entry.maxCopyMove)
                {
                    ++entry.inFlightCopyMove;
                    return Permit(this, std::move(key), kind);
                }
            }
            else
            {
                entry.maxDelete = maxEffective;
                if (entry.inFlightDelete < entry.maxDelete)
                {
                    ++entry.inFlightDelete;
                    return Permit(this, std::move(key), kind);
                }
            }

            _cv.wait_for(lock, std::chrono::milliseconds(100));
        }
    }

    std::mutex _mutex;
    std::condition_variable _cv;
    std::unordered_map<std::wstring, Entry> _entries;
};

ConnectionConcurrencyLimiter& GetConnectionConcurrencyLimiter() noexcept
{
    static ConnectionConcurrencyLimiter limiter;
    return limiter;
}

[[nodiscard]] std::optional<DWORD> Win32ErrorFromHRESULT(HRESULT hr) noexcept
{
    if (hr == E_ACCESSDENIED)
    {
        return ERROR_ACCESS_DENIED;
    }
    if (hr == E_ABORT)
    {
        return ERROR_CANCELLED;
    }

    if (HRESULT_FACILITY(hr) == FACILITY_WIN32)
    {
        return HRESULT_CODE(hr);
    }

    return std::nullopt;
}

[[nodiscard]] bool IsNetworkOfflineError(DWORD error) noexcept
{
    switch (error)
    {
        case ERROR_BAD_NETPATH:
        case ERROR_BAD_NET_NAME:
        case ERROR_NETNAME_DELETED:
        case ERROR_NETWORK_UNREACHABLE:
        case ERROR_HOST_UNREACHABLE:
        case ERROR_PORT_UNREACHABLE:
        case ERROR_CONNECTION_UNAVAIL:
        case ERROR_NOT_CONNECTED:
        case ERROR_CONNECTION_REFUSED:
        case ERROR_NO_NETWORK:
        case ERROR_NETWORK_ACCESS_DENIED: return true;
        default: return false;
    }
}

[[nodiscard]] bool IsCircuitBreakerAuthError(DWORD error) noexcept
{
    switch (error)
    {
        case ERROR_INVALID_PASSWORD:
        case ERROR_LOGON_FAILURE: return true;
        default: return false;
    }
}

[[nodiscard]] bool IsCircuitBreakerTransientError(DWORD error) noexcept
{
    if (IsNetworkOfflineError(error))
    {
        return true;
    }

    switch (error)
    {
        case ERROR_BAD_NET_RESP:
        case ERROR_CONNECTION_ABORTED:
        case ERROR_SEM_TIMEOUT:
        case ERROR_TIMEOUT:
        case ERROR_UNEXP_NET_ERR: return true;
        default: return false;
    }
}

[[nodiscard]] bool ShouldCountCircuitBreakerFailure(HRESULT hr) noexcept
{
    if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
    {
        return false;
    }

    const std::optional<DWORD> errorOpt = Win32ErrorFromHRESULT(hr);
    const DWORD error                   = errorOpt.value_or(0);
    if (errorOpt.has_value() && IsCircuitBreakerAuthError(error))
    {
        return false;
    }

    return errorOpt.has_value() && IsCircuitBreakerTransientError(error);
}

class ConnectionCircuitBreaker final
{
public:
    ConnectionCircuitBreaker()  = default;
    ~ConnectionCircuitBreaker() = default;

    ConnectionCircuitBreaker(const ConnectionCircuitBreaker&)            = delete;
    ConnectionCircuitBreaker& operator=(const ConnectionCircuitBreaker&) = delete;
    ConnectionCircuitBreaker(ConnectionCircuitBreaker&&)                 = delete;
    ConnectionCircuitBreaker& operator=(ConnectionCircuitBreaker&&)      = delete;

    // Returns true if the request should proceed, false if it should fail fast.
    [[nodiscard]] bool ShouldAllow(std::initializer_list<std::wstring_view> connectionIds) noexcept
    {
        std::wstring_view id1;
        std::wstring_view id2;
        for (const std::wstring_view id : connectionIds)
        {
            if (id.empty())
            {
                continue;
            }

            if (id1.empty())
            {
                id1 = id;
                continue;
            }

            if (id2.empty() && ! OrdinalString::EqualsNoCase(id1, id))
            {
                id2 = id;
                continue;
            }
        }

        if (id1.empty() && id2.empty())
        {
            return true;
        }

        const ULONGLONG nowTick = GetTickCount64();

        std::lock_guard lock(_mutex);

        const bool deny = wouldDenyLocked(id1, nowTick) || wouldDenyLocked(id2, nowTick);
        if (deny)
        {
            return false;
        }

        // Allow: mark any open connections as having an in-flight probe.
        markProbeLocked(id1, nowTick);
        markProbeLocked(id2, nowTick);
        return true;
    }

    void RecordSuccess(std::initializer_list<std::wstring_view> connectionIds) noexcept
    {
        recordResult(connectionIds, S_OK);
    }

    void RecordFailure(std::initializer_list<std::wstring_view> connectionIds, HRESULT hr) noexcept
    {
        recordResult(connectionIds, hr);
    }

private:
    static constexpr ULONGLONG kWindowMs       = 30'000ull;
    static constexpr size_t kFailureThreshold  = 5u;
    static constexpr ULONGLONG kCooldownMs     = 30'000ull;
    static constexpr ULONGLONG kProbeBackoffMs = 5'000ull;

    [[nodiscard]] static std::wstring MakeEntryKey(std::wstring_view connectionId)
    {
        // Connection IDs are GUID strings; treat them case-insensitively by normalizing to lowercase.
        std::wstring key(connectionId);
        for (wchar_t& ch : key)
        {
            if (ch >= L'A' && ch <= L'Z')
            {
                ch = static_cast<wchar_t>(ch - L'A' + L'a');
            }
        }
        return key;
    }

    struct Entry final
    {
        std::deque<ULONGLONG> transientFailureTicks;
        ULONGLONG openUntilTick        = 0;
        ULONGLONG nextProbeAllowedTick = 0;
        bool probeInFlight             = false;
    };

    void pruneLocked(Entry& entry, ULONGLONG nowTick) noexcept
    {
        while (! entry.transientFailureTicks.empty())
        {
            const ULONGLONG oldest = entry.transientFailureTicks.front();
            if (nowTick >= oldest && (nowTick - oldest) > kWindowMs)
            {
                entry.transientFailureTicks.pop_front();
                continue;
            }
            break;
        }
    }

    [[nodiscard]] bool wouldDenyLocked(std::wstring_view connectionId, ULONGLONG nowTick) noexcept
    {
        if (connectionId.empty())
        {
            return false;
        }

        const std::wstring key = MakeEntryKey(connectionId);
        auto it                = _entries.find(key);
        if (it == _entries.end())
        {
            return false;
        }

        Entry& entry = it->second;
        pruneLocked(entry, nowTick);

        if (entry.openUntilTick <= nowTick)
        {
            entry.openUntilTick        = 0;
            entry.probeInFlight        = false;
            entry.nextProbeAllowedTick = 0;
            if (entry.transientFailureTicks.empty())
            {
                _entries.erase(it);
            }
            return false;
        }

        if (entry.probeInFlight)
        {
            return true;
        }

        return nowTick < entry.nextProbeAllowedTick;
    }

    void markProbeLocked(std::wstring_view connectionId, ULONGLONG nowTick) noexcept
    {
        if (connectionId.empty())
        {
            return;
        }

        const std::wstring key = MakeEntryKey(connectionId);
        auto it                = _entries.find(key);
        if (it == _entries.end())
        {
            return;
        }

        Entry& entry = it->second;
        if (entry.openUntilTick > nowTick)
        {
            entry.probeInFlight        = true;
            entry.nextProbeAllowedTick = nowTick + kProbeBackoffMs;
        }
        else if (entry.transientFailureTicks.empty())
        {
            _entries.erase(it);
        }
    }

    void recordResult(std::initializer_list<std::wstring_view> connectionIds, HRESULT hr) noexcept
    {
        std::wstring_view id1;
        std::wstring_view id2;
        for (const std::wstring_view id : connectionIds)
        {
            if (id.empty())
            {
                continue;
            }

            if (id1.empty())
            {
                id1 = id;
                continue;
            }

            if (id2.empty() && ! OrdinalString::EqualsNoCase(id1, id))
            {
                id2 = id;
                continue;
            }
        }

        if (id1.empty() && id2.empty())
        {
            return;
        }

        const ULONGLONG nowTick     = GetTickCount64();
        const bool countableFailure = FAILED(hr) && ShouldCountCircuitBreakerFailure(hr);
        const bool isSuccess        = SUCCEEDED(hr);

        std::lock_guard lock(_mutex);

        const auto apply = [&](std::wstring_view id) noexcept
        {
            if (id.empty())
            {
                return;
            }

            std::wstring key = MakeEntryKey(id);
            auto it          = _entries.find(key);

            if (isSuccess)
            {
                if (it != _entries.end())
                {
                    _entries.erase(it);
                }
                return;
            }

            if (! countableFailure)
            {
                if (it != _entries.end())
                {
                    it->second.probeInFlight = false;
                    if (it->second.openUntilTick <= nowTick && it->second.transientFailureTicks.empty())
                    {
                        _entries.erase(it);
                    }
                }
                return;
            }

            if (it == _entries.end())
            {
                auto [insertedIt, inserted] = _entries.emplace(std::move(key), Entry{});
                it                          = insertedIt;
            }

            Entry& entry        = it->second;
            entry.probeInFlight = false;
            pruneLocked(entry, nowTick);
            entry.transientFailureTicks.push_back(nowTick);
            pruneLocked(entry, nowTick);

            if (entry.transientFailureTicks.size() >= kFailureThreshold)
            {
                entry.transientFailureTicks.clear();
                entry.openUntilTick        = nowTick + kCooldownMs;
                entry.nextProbeAllowedTick = nowTick;
            }
        };

        apply(id1);
        apply(id2);
    }

    std::mutex _mutex;
    std::unordered_map<std::wstring, Entry> _entries;
};

ConnectionCircuitBreaker& GetConnectionCircuitBreaker() noexcept
{
    static ConnectionCircuitBreaker breaker;
    return breaker;
}

template <typename Fn>
[[nodiscard]] HRESULT RunWithCircuitBreaker(ConnectionCircuitBreaker& breaker,
                                            std::wstring_view sourceConnectionId,
                                            std::wstring_view destinationConnectionId,
                                            Fn&& fn) noexcept
{
    const bool hasCircuitBreakerConnection = ! sourceConnectionId.empty() || ! destinationConnectionId.empty();
    if (hasCircuitBreakerConnection && ! breaker.ShouldAllow({sourceConnectionId, destinationConnectionId}))
    {
        return HRESULT_FROM_WIN32(ERROR_NO_NETWORK);
    }

    const HRESULT hr = fn();

    if (hasCircuitBreakerConnection)
    {
        if (SUCCEEDED(hr))
        {
            breaker.RecordSuccess({sourceConnectionId, destinationConnectionId});
        }
        else
        {
            breaker.RecordFailure({sourceConnectionId, destinationConnectionId}, hr);
        }
    }

    return hr;
}

[[nodiscard]] bool IsPathTooLongError(DWORD error) noexcept
{
    switch (error)
    {
        case ERROR_FILENAME_EXCED_RANGE:
        case ERROR_BUFFER_OVERFLOW: return true;
        default: return false;
    }
}

[[nodiscard]] bool IsCopyMoveOperation(FileSystemOperation operation) noexcept
{
    return operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
}

[[nodiscard]] bool IsDestinationCollisionOperation(FileSystemOperation operation) noexcept
{
    return IsCopyMoveOperation(operation) || operation == FILESYSTEM_RENAME;
}

[[nodiscard]] bool IsTraversalResourceLimitStatus(HRESULT status) noexcept
{
    return status == HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW) || status == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
}

[[nodiscard]] bool IsDirectoryReparsePoint(const wil::com_ptr<IFileSystemIO>& fileSystemIo, std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    unsigned long attributes = 0;
    if (fileSystemIo)
    {
        const HRESULT hr = fileSystemIo->GetAttributes(std::wstring(path).c_str(), &attributes);
        if (FAILED(hr))
        {
            return false;
        }
    }
    else
    {
        const DWORD win32 = GetFileAttributesW(std::wstring(path).c_str());
        if (win32 == INVALID_FILE_ATTRIBUTES)
        {
            return false;
        }
        attributes = win32;
    }

    return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
}

[[nodiscard]] Task::ConflictBucket ClassifyConflictBucket(FileSystemOperation operation,
                                                          FileSystemFlags flags,
                                                          const wil::com_ptr<IFileSystemIO>& fileSystemIo,
                                                          HRESULT status,
                                                          std::wstring_view sourcePath,
                                                          std::wstring_view destinationPath,
                                                          bool unsupportedReparseHint) noexcept
{
    if (status == HRESULT_FROM_WIN32(ERROR_CANCELLED) || status == E_ABORT)
    {
        return Task::ConflictBucket::Unknown;
    }

    if (unsupportedReparseHint)
    {
        return Task::ConflictBucket::UnsupportedReparse;
    }

    if (operation == FILESYSTEM_DELETE && (flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0)
    {
        // Deleting via the recycle bin is handled by the shell and can fail for a variety of reasons
        // (including cases that would succeed as a direct delete). Offer a permanent-delete fallback.
        return Task::ConflictBucket::RecycleFailed;
    }

    const std::optional<DWORD> errorOpt = Win32ErrorFromHRESULT(status);
    const DWORD error                   = errorOpt.value_or(0);

    switch (error)
    {
        case ERROR_ALREADY_EXISTS:
        case ERROR_FILE_EXISTS: return Task::ConflictBucket::RegularFileExists;
        // Metadata is read before the prompt becomes visible. This provisional class is refined to
        // TypeMismatch or DestinationLink and never gains an Overwrite action as a folder bucket.
        case ERROR_DIR_NOT_EMPTY:
        case ERROR_DATATYPE_MISMATCH:
            if (IsCopyMoveOperation(operation))
            {
                return Task::ConflictBucket::TypeMismatch;
            }
            break;
        // A junction/symlink occupying the destination is never a silent merge target.
        case ERROR_REPARSE_POINT_ENCOUNTERED:
            if (IsCopyMoveOperation(operation))
            {
                return Task::ConflictBucket::DestinationLink;
            }
            break;
        case ERROR_INVALID_NAME:
        case ERROR_BAD_PATHNAME:
            if (IsCopyMoveOperation(operation) || operation == FILESYSTEM_RENAME)
            {
                return Task::ConflictBucket::NameNotRepresentable;
            }
            break;
        case ERROR_CANT_RESOLVE_FILENAME:
            if (IsCopyMoveOperation(operation))
            {
                return Task::ConflictBucket::TargetConflict;
            }
            break;
        case ERROR_SHARING_VIOLATION:
        case ERROR_LOCK_VIOLATION: return Task::ConflictBucket::SharingViolation;
        case ERROR_DISK_FULL:
        case ERROR_HANDLE_DISK_FULL: return Task::ConflictBucket::DiskFull;
        default: break;
    }

    if (IsPathTooLongError(error))
    {
        return Task::ConflictBucket::PathTooLong;
    }

    if (IsNetworkOfflineError(error))
    {
        return Task::ConflictBucket::NetworkOffline;
    }

    if (error == ERROR_NOT_SUPPORTED && IsCopyMoveOperation(operation) && IsDirectoryReparsePoint(fileSystemIo, sourcePath))
    {
        return Task::ConflictBucket::UnsupportedReparse;
    }

    if (error == ERROR_ACCESS_DENIED)
    {
        const bool isDelete           = operation == FILESYSTEM_DELETE;
        const std::wstring_view probe = isDelete ? sourcePath : destinationPath;

        if (! probe.empty())
        {
            unsigned long attributes = 0;
            bool gotAttributes       = false;
            if (fileSystemIo)
            {
                gotAttributes = SUCCEEDED(fileSystemIo->GetAttributes(std::wstring(probe).c_str(), &attributes));
            }
            else
            {
                const DWORD win32 = GetFileAttributesW(std::wstring(probe).c_str());
                if (win32 != INVALID_FILE_ATTRIBUTES)
                {
                    attributes    = win32;
                    gotAttributes = true;
                }
            }

            if (gotAttributes && (attributes & FILE_ATTRIBUTE_READONLY) != 0)
            {
                return Task::ConflictBucket::ReadOnlyRegularFile;
            }
        }

        return Task::ConflictBucket::AccessDenied;
    }

    return Task::ConflictBucket::Unknown;
}

// Retry is offered only where retrying can plausibly succeed. Existing-destination and
// read-only collisions are deterministic — they have dedicated resolution actions instead, and
// a Retry button there is a dead end that pads the prompt.
[[nodiscard]] bool IsRetryableConflictBucket(Task::ConflictBucket bucket) noexcept
{
    return bucket != Task::ConflictBucket::UnsupportedReparse && bucket != Task::ConflictBucket::RegularFileExists &&
           bucket != Task::ConflictBucket::ReadOnlyRegularFileExists && bucket != Task::ConflictBucket::TypeMismatch &&
           bucket != Task::ConflictBucket::DestinationLink && bucket != Task::ConflictBucket::ReadOnlyRegularFile &&
           bucket != Task::ConflictBucket::NameNotRepresentable && bucket != Task::ConflictBucket::TargetConflict;
}

using ConflictBucket = Task::ConflictBucket;
using ConflictAction = Task::ConflictAction;

[[nodiscard]] wil::com_ptr<IFileSystemIO> QueryFileSystemIo(IFileSystem* fileSystem) noexcept
{
    wil::com_ptr<IFileSystemIO> result;
    if (fileSystem)
    {
        static_cast<void>(fileSystem->QueryInterface(IID_PPV_ARGS(result.addressof())));
    }
    return result;
}

[[nodiscard]] Task::ConflictPromptState::ItemMetadata ReadConflictItemMetadata(IFileSystem* fileSystem,
                                                                               IFileSystemIO* io,
                                                                               std::wstring_view path,
                                                                               bool useWin32Metadata) noexcept
{
    Task::ConflictPromptState::ItemMetadata result{};
    if (path.empty())
    {
        return result;
    }

    const std::wstring pathText(path);

    const auto applyBoundObjectKind = [&](Task::ConflictPromptState::ItemMetadata& metadata) noexcept -> bool
    {
        if (fileSystem == nullptr)
        {
            return false;
        }
        wil::com_ptr<IFileSystemObjectBinding> binding;
        if (FAILED(fileSystem->QueryInterface(__uuidof(IFileSystemObjectBinding), binding.put_void())) || ! binding)
        {
            return false;
        }
        constexpr FileSystemBindFlags flags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
        wil::com_ptr<IFileSystemBoundObject> bound;
        if (FAILED(binding->BindObject(pathText.c_str(), flags, bound.put())) || ! bound)
        {
            return false;
        }
        FileSystemBoundObjectSnapshot snapshot{};
        snapshot.sizeBytes = sizeof(snapshot);
        if (FAILED(bound->GetSnapshot(&snapshot)))
        {
            return false;
        }
        switch (snapshot.kind)
        {
            case FILESYSTEM_BOUND_REGULAR_FILE:
                metadata.objectKindKnown = true;
                metadata.isDirectory     = false;
                metadata.isLink          = false;
                return true;
            case FILESYSTEM_BOUND_DIRECTORY:
                metadata.objectKindKnown = true;
                metadata.isDirectory     = true;
                metadata.isLink          = false;
                return true;
            case FILESYSTEM_BOUND_LINK:
                metadata.objectKindKnown = true;
                metadata.isLink          = true;
                return true;
            case FILESYSTEM_BOUND_OTHER:
            default: return false;
        }
    };

#ifdef ENABLE_TESTS
    g_fileOpsConflictMetadataPausePoint.Pause(g_fileOpsConflictMetadataPauseBailoutMs.load(std::memory_order_acquire));
#endif

    if (useWin32Metadata && NavigationLocation::LooksLikeWindowsAbsolutePath(path))
    {
        WIN32_FILE_ATTRIBUTE_DATA data{};
        if (GetFileAttributesExW(pathText.c_str(), GetFileExInfoStandard, &data) == FALSE)
        {
            return result;
        }

        result.available       = true;
        result.objectKindKnown = (data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0u;
        result.attributes      = data.dwFileAttributes;
        result.isDirectory     = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if ((data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
        {
            const bool boundKindKnown = applyBoundObjectKind(result);
            if (! boundKindKnown || result.isLink)
            {
                const DWORD openFlags = FILE_FLAG_OPEN_REPARSE_POINT | (result.isDirectory ? FILE_FLAG_BACKUP_SEMANTICS : 0u);
                wil::unique_hfile handle(CreateFileW(pathText.c_str(),
                                                     FILE_READ_ATTRIBUTES,
                                                     FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                     nullptr,
                                                     OPEN_EXISTING,
                                                     openFlags,
                                                     nullptr));
                if (handle)
                {
                    FILE_ATTRIBUTE_TAG_INFO tagInfo{};
                    if (GetFileInformationByHandleEx(handle.get(), FileAttributeTagInfo, &tagInfo, sizeof(tagInfo)) != FALSE)
                    {
                        result.objectKindKnown = true;
                        result.isLink          = IsReparseTagNameSurrogate(tagInfo.ReparseTag) != FALSE;
                        if (result.isLink)
                        {
                            if (tagInfo.ReparseTag == IO_REPARSE_TAG_MOUNT_POINT)
                            {
                                result.linkKind = Task::ConflictLinkKind::Junction;
                            }
                            else if (tagInfo.ReparseTag == IO_REPARSE_TAG_SYMLINK)
                            {
                                result.linkKind = result.isDirectory ? Task::ConflictLinkKind::SymbolicDirectory : Task::ConflictLinkKind::SymbolicFile;
                            }
                        }
                    }
                    else
                    {
                        result.objectKindKnown = false;
                    }
                }
                else
                {
                    result.objectKindKnown = false;
                }
            }
        }
        const uint64_t lastWriteTime =
            (static_cast<uint64_t>(data.ftLastWriteTime.dwHighDateTime) << 32) | static_cast<uint64_t>(data.ftLastWriteTime.dwLowDateTime);
        result.lastWriteTime = static_cast<__int64>(lastWriteTime);
        if (! result.isDirectory)
        {
            result.sizeKnown = true;
            result.sizeBytes = (static_cast<uint64_t>(data.nFileSizeHigh) << 32) | static_cast<uint64_t>(data.nFileSizeLow);
        }
        return result;
    }

    if (! io)
    {
        return result;
    }

    FileSystemBasicInformation basic{};
    basic.sizeBytes       = sizeof(basic);
    const HRESULT basicHr = io->GetFileBasicInformation(pathText.c_str(), &basic);
    if (SUCCEEDED(basicHr))
    {
        result.available       = true;
        result.attributes      = basic.attributes;
        result.isDirectory     = (basic.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        result.objectKindKnown = (basic.attributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0u;
        if (! result.objectKindKnown)
        {
            static_cast<void>(applyBoundObjectKind(result));
        }
        result.lastWriteTime = basic.lastWriteTime;
        return result;
    }

    unsigned long attributes   = 0;
    const HRESULT attributesHr = io->GetAttributes(pathText.c_str(), &attributes);
    if (FAILED(attributesHr))
    {
        return result;
    }

    result.available       = true;
    result.attributes      = attributes;
    result.isDirectory     = (result.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    result.objectKindKnown = (result.attributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0u;
    if (! result.objectKindKnown)
    {
        static_cast<void>(applyBoundObjectKind(result));
    }
    result.lastWriteTime = 0;
    return result;
}

[[nodiscard]] Task::ConflictItemKind ConflictItemKindFromMetadata(const Task::ConflictPromptState::ItemMetadata& metadata) noexcept
{
    if (! metadata.available || ! metadata.objectKindKnown)
    {
        return Task::ConflictItemKind::Unknown;
    }
    if (metadata.isLink)
    {
        return Task::ConflictItemKind::Link;
    }
    return metadata.isDirectory ? Task::ConflictItemKind::Directory : Task::ConflictItemKind::RegularFile;
}

[[nodiscard]] ConflictBucket RefineTypedConflictBucket(FileSystemOperation operation,
                                                       ConflictBucket provisionalBucket,
                                                       const Task::ConflictPromptState::ItemMetadata& sourceMetadata,
                                                       const Task::ConflictPromptState::ItemMetadata& destinationMetadata) noexcept
{
    if (! IsDestinationCollisionOperation(operation))
    {
        return provisionalBucket;
    }

    if (destinationMetadata.available && destinationMetadata.objectKindKnown && destinationMetadata.isLink)
    {
        return ConflictBucket::DestinationLink;
    }

    if (provisionalBucket == ConflictBucket::DestinationLink && destinationMetadata.available && destinationMetadata.objectKindKnown &&
        ! destinationMetadata.isLink)
    {
        provisionalBucket = ConflictBucket::RegularFileExists;
    }

    const bool isCollisionClass = provisionalBucket == ConflictBucket::RegularFileExists || provisionalBucket == ConflictBucket::ReadOnlyRegularFile ||
                                  provisionalBucket == ConflictBucket::ReadOnlyRegularFileExists || provisionalBucket == ConflictBucket::TypeMismatch;
    if (! isCollisionClass || ! sourceMetadata.available || ! destinationMetadata.available)
    {
        return provisionalBucket;
    }

    const Task::ConflictItemKind sourceKind      = ConflictItemKindFromMetadata(sourceMetadata);
    const Task::ConflictItemKind destinationKind = ConflictItemKindFromMetadata(destinationMetadata);
    if (sourceKind != destinationKind)
    {
        return ConflictBucket::TypeMismatch;
    }

    if (sourceKind == Task::ConflictItemKind::RegularFile)
    {
        return (destinationMetadata.attributes & FILE_ATTRIBUTE_READONLY) != 0 ? ConflictBucket::ReadOnlyRegularFileExists : ConflictBucket::RegularFileExists;
    }

    // Folder-on-folder belongs to merge and should not reach a conflict prompt. If a provider does
    // report this shape, fail closed with non-destructive Keep Both/Skip/Cancel rather than inventing
    // a folder overwrite grant.
    return ConflictBucket::TypeMismatch;
}

[[nodiscard]] Task::ConflictDecisionScope BuildConflictDecisionScope(const Task& task,
                                                                     const Task::PerItemCallbackCookie* perItemCookie,
                                                                     ConflictBucket bucket,
                                                                     const Task::ConflictPromptState::ItemMetadata& sourceMetadata,
                                                                     const Task::ConflictPromptState::ItemMetadata& destinationMetadata) noexcept
{
    Task::ConflictDecisionScope scope{};
    scope.bucket               = bucket;
    scope.sourceKind           = ConflictItemKindFromMetadata(sourceMetadata);
    scope.destinationKind      = ConflictItemKindFromMetadata(destinationMetadata);
    scope.destinationLinkKind  = destinationMetadata.linkKind;
    scope.sourceProfileId      = task._sourcePluginId;
    scope.destinationProfileId = task._destinationPluginId.empty() ? task._sourcePluginId : task._destinationPluginId;

    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = task.LoadPlans();
    if (plans && perItemCookie != nullptr)
    {
        size_t remainingIndex = perItemCookie->itemIndex;
        for (const FileOperations::FileOperationPlan& plan : *plans)
        {
            const auto* transferPlan = std::get_if<FileOperations::TransferPlan>(&plan);
            if (! transferPlan)
            {
                continue;
            }
            if (remainingIndex < transferPlan->selectedItems.size())
            {
                scope.sourceProfileId      = transferPlan->sourceEndpoint.profileId;
                scope.destinationProfileId = transferPlan->destinationEndpoint.profileId;
                scope.destinationRootId    = transferPlan->destinationEndpoint.rootId;
                break;
            }
            remainingIndex -= transferPlan->selectedItems.size();
        }
    }

    switch (bucket)
    {
        case ConflictBucket::RegularFileExists:
        case ConflictBucket::ReadOnlyRegularFileExists:
            scope.applyToAllEligible = scope.sourceKind == Task::ConflictItemKind::RegularFile && scope.destinationKind == Task::ConflictItemKind::RegularFile;
            break;
        case ConflictBucket::TypeMismatch:
            scope.applyToAllEligible = scope.sourceKind != Task::ConflictItemKind::Unknown && scope.destinationKind != Task::ConflictItemKind::Unknown &&
                                       scope.sourceKind != scope.destinationKind;
            break;
        case ConflictBucket::ReadOnlyRegularFile:
        case ConflictBucket::NameNotRepresentable:
        case ConflictBucket::SharingViolation:
        case ConflictBucket::AccessDenied:
        case ConflictBucket::DiskFull:
        case ConflictBucket::PathTooLong:
        case ConflictBucket::NetworkOffline: scope.applyToAllEligible = true; break;
        case ConflictBucket::InsufficientSpace:
        case ConflictBucket::SpaceUnknown:
        case ConflictBucket::EfsPlaintext:
        case ConflictBucket::SparseInflation:
        case ConflictBucket::PlaceholderHydration:
        case ConflictBucket::MetadataLoss: scope.applyToAllEligible = true; break;
        case ConflictBucket::DestinationLink: scope.applyToAllEligible = scope.destinationLinkKind != Task::ConflictLinkKind::Unknown; break;
        case ConflictBucket::TargetConflict:
        case ConflictBucket::RecycleFailed:
        case ConflictBucket::SameHostOverlap:
        case ConflictBucket::SameHostLiveOutput:
        case ConflictBucket::PermanentDeleteConfirmation:
        case ConflictBucket::UnsupportedReparse:
        case ConflictBucket::Unknown:
        case ConflictBucket::Count:
        default: scope.applyToAllEligible = false; break;
    }

    if (perItemCookie == nullptr && IsCopyMoveOperation(task._operation))
    {
        scope.applyToAllEligible = false;
    }
    return scope;
}

struct ConflictActionPolicy
{
    std::array<ConflictAction, Task::ConflictPromptState::kMaxActions> actions{};
    size_t actionCount = 0;
    std::array<ConflictAction, Task::ConflictPromptState::kMaxActions> primaryActions{};
    size_t primaryActionCount = 0;
    std::array<ConflictAction, Task::ConflictPromptState::kMaxActions> overflowActions{};
    size_t overflowActionCount   = 0;
    ConflictAction defaultAction = ConflictAction::None;
    ConflictAction escapeAction  = ConflictAction::None;
    bool applyToAllEligible      = false;
    bool skipAllEligible         = false;
    bool buttonsPublishable      = false;

    void Add(ConflictAction action) noexcept
    {
        if (actionCount < actions.size())
        {
            actions[actionCount] = action;
            ++actionCount;
        }
    }

    [[nodiscard]] bool Contains(ConflictAction action) const noexcept
    {
        return std::ranges::find(actions.begin(), actions.begin() + static_cast<std::ptrdiff_t>(actionCount), action) !=
               actions.begin() + static_cast<std::ptrdiff_t>(actionCount);
    }

    void AddPrimaryIfAvailable(ConflictAction action) noexcept
    {
        if (primaryActionCount >= 3u || ! Contains(action))
        {
            return;
        }
        if (std::ranges::find(primaryActions.begin(), primaryActions.begin() + static_cast<std::ptrdiff_t>(primaryActionCount), action) !=
            primaryActions.begin() + static_cast<std::ptrdiff_t>(primaryActionCount))
        {
            return;
        }
        primaryActions[primaryActionCount] = action;
        ++primaryActionCount;
    }

    void FinalizePresentation(ConflictBucket bucket, bool scopedApplyToAllEligible) noexcept
    {
        applyToAllEligible = scopedApplyToAllEligible;
        skipAllEligible    = Contains(ConflictAction::SkipAll);
        buttonsPublishable = actionCount > 0u;
        if (! buttonsPublishable)
        {
            return;
        }

        if (Contains(ConflictAction::Cancel))
        {
            defaultAction = ConflictAction::Cancel;
            escapeAction  = ConflictAction::Cancel;
        }

        switch (bucket)
        {
            case ConflictBucket::RegularFileExists:
                AddPrimaryIfAvailable(ConflictAction::Overwrite);
                AddPrimaryIfAvailable(ConflictAction::KeepBoth);
                break;
            case ConflictBucket::ReadOnlyRegularFileExists:
                AddPrimaryIfAvailable(ConflictAction::ReplaceReadOnly);
                AddPrimaryIfAvailable(ConflictAction::KeepBoth);
                break;
            case ConflictBucket::TypeMismatch: AddPrimaryIfAvailable(ConflictAction::KeepBoth); break;
            case ConflictBucket::DestinationLink:
                AddPrimaryIfAvailable(ConflictAction::ReplaceLink);
                AddPrimaryIfAvailable(ConflictAction::KeepBoth);
                break;
            case ConflictBucket::RecycleFailed:
            case ConflictBucket::PermanentDeleteConfirmation:
                AddPrimaryIfAvailable(ConflictAction::Cancel);
                AddPrimaryIfAvailable(ConflictAction::PermanentDelete);
                break;
            case ConflictBucket::InsufficientSpace:
            case ConflictBucket::SpaceUnknown:
            case ConflictBucket::SparseInflation:
            case ConflictBucket::PlaceholderHydration: AddPrimaryIfAvailable(ConflictAction::Proceed); break;
            case ConflictBucket::SameHostOverlap:
                AddPrimaryIfAvailable(ConflictAction::Proceed);
                AddPrimaryIfAvailable(ConflictAction::RunConcurrently);
                // Queue after is the safe scheduling default. Escape remains Don't start.
                defaultAction = ConflictAction::Proceed;
                break;
            case ConflictBucket::SameHostLiveOutput:
                AddPrimaryIfAvailable(ConflictAction::QueueUntilOtherTaskFinishes);
                AddPrimaryIfAvailable(ConflictAction::InvalidateLiveOutput);
                // Waiting for the named live publisher is the non-destructive default.
                defaultAction = ConflictAction::QueueUntilOtherTaskFinishes;
                break;
            case ConflictBucket::EfsPlaintext:
            case ConflictBucket::MetadataLoss:
                AddPrimaryIfAvailable(ConflictAction::Proceed);
                AddPrimaryIfAvailable(ConflictAction::RetainSource);
                break;
            case ConflictBucket::ReadOnlyRegularFile: AddPrimaryIfAvailable(ConflictAction::ReplaceReadOnly); break;
            case ConflictBucket::NameNotRepresentable: AddPrimaryIfAvailable(ConflictAction::KeepBoth); break;
            case ConflictBucket::TargetConflict:
            case ConflictBucket::AccessDenied:
            case ConflictBucket::SharingViolation:
            case ConflictBucket::DiskFull:
            case ConflictBucket::PathTooLong:
            case ConflictBucket::NetworkOffline:
            case ConflictBucket::UnsupportedReparse:
            case ConflictBucket::Unknown:
            case ConflictBucket::Count:
            default: AddPrimaryIfAvailable(ConflictAction::Retry); break;
        }

        // Cancel is the safe action and must remain directly visible rather than
        // being displaced into More by a three-button layout.
        AddPrimaryIfAvailable(ConflictAction::Cancel);
        AddPrimaryIfAvailable(ConflictAction::Skip);

        for (size_t index = 0u; index < actionCount; ++index)
        {
            const ConflictAction action = actions[index];
            if (std::ranges::find(primaryActions.begin(), primaryActions.begin() + static_cast<std::ptrdiff_t>(primaryActionCount), action) !=
                primaryActions.begin() + static_cast<std::ptrdiff_t>(primaryActionCount))
            {
                continue;
            }
            if (overflowActionCount < overflowActions.size())
            {
                overflowActions[overflowActionCount] = action;
                ++overflowActionCount;
            }
        }
    }
};

[[nodiscard]] constexpr bool IsDeferredConsentBucket(ConflictBucket bucket) noexcept
{
    switch (bucket)
    {
        case ConflictBucket::RecycleFailed:
        case ConflictBucket::InsufficientSpace:
        case ConflictBucket::SpaceUnknown:
        case ConflictBucket::EfsPlaintext:
        case ConflictBucket::SparseInflation:
        case ConflictBucket::PlaceholderHydration:
        case ConflictBucket::MetadataLoss:
        case ConflictBucket::SameHostOverlap:
        case ConflictBucket::SameHostLiveOutput:
        case ConflictBucket::PermanentDeleteConfirmation: return true;
        case ConflictBucket::RegularFileExists:
        case ConflictBucket::ReadOnlyRegularFile:
        case ConflictBucket::ReadOnlyRegularFileExists:
        case ConflictBucket::TypeMismatch:
        case ConflictBucket::DestinationLink:
        case ConflictBucket::NameNotRepresentable:
        case ConflictBucket::TargetConflict:
        case ConflictBucket::SharingViolation:
        case ConflictBucket::AccessDenied:
        case ConflictBucket::DiskFull:
        case ConflictBucket::PathTooLong:
        case ConflictBucket::NetworkOffline:
        case ConflictBucket::UnsupportedReparse:
        case ConflictBucket::Unknown:
        case ConflictBucket::Count: return false;
    }
    return false;
}

[[nodiscard]] ConflictActionPolicy BuildConflictActionPolicy(ConflictBucket bucket,
                                                             bool allowRetry,
                                                             bool allowKeepBoth,
                                                             bool applyToAllEligible,
                                                             bool allowDestructiveReplacement,
                                                             bool allowConcurrentRun    = false,
                                                             bool allowLiveInvalidation = true) noexcept
{
    ConflictActionPolicy policy{};

    if (bucket == ConflictBucket::PermanentDeleteConfirmation)
    {
        // C1: the initial permanent-delete confirmation on the card. Cancel is first and the
        // presentation default; there is nothing to skip and nothing to apply to all.
        policy.Add(ConflictAction::Cancel);
        policy.Add(ConflictAction::PermanentDelete);
        policy.FinalizePresentation(bucket, false);
        return policy;
    }

    if (bucket == ConflictBucket::RecycleFailed)
    {
        // Recycle escalation is a one-item destructive consent gate. Cancel is deliberately first
        // and is the presentation default; Apply-to-all and Skip All are never eligible here.
        policy.Add(ConflictAction::Cancel);
        policy.Add(ConflictAction::PermanentDelete);
        policy.Add(ConflictAction::Skip);
        policy.FinalizePresentation(bucket, false);
        return policy;
    }

    switch (bucket)
    {
        case ConflictBucket::RegularFileExists:
            if (allowDestructiveReplacement)
            {
                policy.Add(ConflictAction::Overwrite);
            }
            if (allowKeepBoth)
            {
                policy.Add(ConflictAction::KeepBoth);
            }
            break;
        case ConflictBucket::ReadOnlyRegularFileExists:
            if (allowDestructiveReplacement)
            {
                policy.Add(ConflictAction::ReplaceReadOnly);
            }
            if (allowKeepBoth)
            {
                policy.Add(ConflictAction::KeepBoth);
            }
            break;
        case ConflictBucket::TypeMismatch:
            if (allowKeepBoth)
            {
                policy.Add(ConflictAction::KeepBoth);
            }
            break;
        case ConflictBucket::DestinationLink:
            if (allowDestructiveReplacement)
            {
                policy.Add(ConflictAction::ReplaceLink);
            }
            if (allowKeepBoth)
            {
                policy.Add(ConflictAction::KeepBoth);
            }
            break;
        case ConflictBucket::ReadOnlyRegularFile:
            if (allowDestructiveReplacement)
            {
                policy.Add(ConflictAction::ReplaceReadOnly);
            }
            break;
        case ConflictBucket::RecycleFailed: break;
        case ConflictBucket::InsufficientSpace:
        case ConflictBucket::SpaceUnknown:
        case ConflictBucket::SparseInflation:
        case ConflictBucket::PlaceholderHydration: policy.Add(ConflictAction::Proceed); break;
        case ConflictBucket::EfsPlaintext:
        case ConflictBucket::MetadataLoss:
            policy.Add(ConflictAction::Proceed);
            policy.Add(ConflictAction::RetainSource);
            break;
        case ConflictBucket::SameHostOverlap:
            policy.Add(ConflictAction::Proceed);
            if (allowConcurrentRun)
            {
                policy.Add(ConflictAction::RunConcurrently);
            }
            break;
        case ConflictBucket::SameHostLiveOutput:
            policy.Add(ConflictAction::QueueUntilOtherTaskFinishes);
            policy.Add(ConflictAction::Skip);
            if (allowLiveInvalidation)
            {
                policy.Add(ConflictAction::InvalidateLiveOutput);
            }
            break;
        case ConflictBucket::NameNotRepresentable:
            if (allowKeepBoth)
            {
                policy.Add(ConflictAction::KeepBoth);
            }
            break;
        case ConflictBucket::TargetConflict:
        case ConflictBucket::AccessDenied:
        case ConflictBucket::SharingViolation:
        case ConflictBucket::DiskFull:
        case ConflictBucket::PathTooLong:
        case ConflictBucket::NetworkOffline:
        case ConflictBucket::UnsupportedReparse:
        case ConflictBucket::PermanentDeleteConfirmation:
        case ConflictBucket::Unknown:
        case ConflictBucket::Count:
        default: break;
    }

    if (allowRetry && ! IsDeferredConsentBucket(bucket))
    {
        policy.Add(ConflictAction::Retry);
    }

    if (! IsDeferredConsentBucket(bucket) || bucket == ConflictBucket::RecycleFailed)
    {
        policy.Add(ConflictAction::Skip);
        if (applyToAllEligible)
        {
            policy.Add(ConflictAction::SkipAll);
        }
    }
    policy.Add(ConflictAction::Cancel);
    policy.FinalizePresentation(bucket, applyToAllEligible);
    return policy;
}

[[nodiscard]] bool IsCacheableConflictDecision(ConflictAction action) noexcept
{
    return action != ConflictAction::Retry && action != ConflictAction::Cancel && action != ConflictAction::None && action != ConflictAction::SkipAll &&
           action != ConflictAction::RunConcurrently && action != ConflictAction::QueueUntilOtherTaskFinishes && action != ConflictAction::InvalidateLiveOutput;
}

[[nodiscard]] bool IsCachedConflictDecisionEligible(const Task& task,
                                                    const Task::PerItemCallbackCookie* perItemCookie,
                                                    const Task::ConflictDecisionScope& scope,
                                                    std::wstring_view conflictDestinationPath,
                                                    ConflictAction action) noexcept
{
    if (! scope.applyToAllEligible)
    {
        return false;
    }

    if (action != ConflictAction::KeepBoth)
    {
        return true;
    }

    return task._executionMode == FolderWindow::FileOperationState::ExecutionMode::PerItem &&
           (task._operation == FILESYSTEM_COPY || task._operation == FILESYSTEM_MOVE) &&
           (scope.bucket == ConflictBucket::RegularFileExists || scope.bucket == ConflictBucket::ReadOnlyRegularFileExists ||
            scope.bucket == ConflictBucket::TypeMismatch) &&
           perItemCookie != nullptr && ! perItemCookie->operationDestinationPath.empty() &&
           NavigationLocation::EqualsNoCase(conflictDestinationPath, perItemCookie->operationDestinationPath);
}

[[nodiscard]] bool IsKeepBothActionEligible(const Task& task,
                                            const Task::PerItemCallbackCookie* perItemCookie,
                                            ConflictBucket bucket,
                                            std::wstring_view conflictDestinationPath) noexcept
{
    if (task._executionMode != FolderWindow::FileOperationState::ExecutionMode::PerItem ||
        (task._operation != FILESYSTEM_COPY && task._operation != FILESYSTEM_MOVE) ||
        (bucket != ConflictBucket::RegularFileExists && bucket != ConflictBucket::ReadOnlyRegularFileExists && bucket != ConflictBucket::TypeMismatch &&
         bucket != ConflictBucket::DestinationLink && bucket != ConflictBucket::NameNotRepresentable) ||
        perItemCookie == nullptr)
    {
        return false;
    }

    const bool providerHandlesNested = NavigationLocation::EqualsNoCase(task._sourcePluginId, L"builtin/file-system") &&
                                       NavigationLocation::EqualsNoCase(task._destinationPluginId, L"builtin/file-system");
    return FileOperations::IsProviderKeepBothDestinationEligible(providerHandlesNested, conflictDestinationPath, perItemCookie->operationDestinationPath);
}

[[nodiscard]] bool IsMissingPathStatus(HRESULT status) noexcept
{
    return status == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || status == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ||
           status == HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

[[nodiscard]] HRESULT FindAvailableUniqueSiblingPath(
    IFileSystemIO* destinationIo, std::wstring_view destinationPath, bool isDirectory, size_t& nextOrdinal, std::wstring& pathOut) noexcept
{
    pathOut.clear();
    if (! destinationIo || destinationPath.empty())
    {
        return E_INVALIDARG;
    }

    constexpr size_t kMaximumKeepBothAttempts = 10'000u;
    nextOrdinal                               = std::max<size_t>(2u, nextOrdinal);
    for (size_t attempt = 0u; attempt < kMaximumKeepBothAttempts; ++attempt, ++nextOrdinal)
    {
        std::wstring candidate;
        const HRESULT candidateHr = Common::Paths::BuildUniqueSiblingPathCandidate(destinationPath, isDirectory, nextOrdinal, candidate);
        if (FAILED(candidateHr))
        {
            return candidateHr;
        }

        unsigned long attributes = 0;
        const HRESULT probeHr    = destinationIo->GetAttributes(candidate.c_str(), &attributes);
        if (IsMissingPathStatus(probeHr))
        {
            pathOut = std::move(candidate);
            ++nextOrdinal;
            return S_OK;
        }
        if (FAILED(probeHr))
        {
            return probeHr;
        }
    }
    return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
}

void StoreConflictDecisionInCacheLocked(Task::ConflictArbiter& arbiter, const Task::ConflictDecisionScope& scope, ConflictAction action) noexcept
{
    if (! scope.applyToAllEligible || ! IsCacheableConflictDecision(action))
    {
        return;
    }

    for (std::optional<Task::CachedConflictDecision>& entry : arbiter.decisionCache)
    {
        if (entry.has_value() && entry->scope == scope)
        {
            entry->action = action;
            return;
        }
    }

    const auto empty =
        std::ranges::find_if(arbiter.decisionCache, [](const std::optional<Task::CachedConflictDecision>& entry) noexcept { return ! entry.has_value(); });
    if (empty != arbiter.decisionCache.end())
    {
        *empty = Task::CachedConflictDecision{scope, action};
    }
}

[[nodiscard]] std::optional<ConflictAction> LoadEligibleConflictDecisionFromCacheLocked(const Task& task,
                                                                                        const Task::PerItemCallbackCookie* perItemCookie,
                                                                                        const Task::ConflictDecisionScope& scope,
                                                                                        std::wstring_view conflictDestinationPath) noexcept
{
    for (const std::optional<Task::CachedConflictDecision>& entry : task._conflictArbiter.decisionCache)
    {
        if (! entry.has_value() || entry->scope != scope)
        {
            continue;
        }
        if (IsCachedConflictDecisionEligible(task, perItemCookie, scope, conflictDestinationPath, entry->action))
        {
            return entry->action;
        }
    }
    return std::nullopt;
}

void ClearConflictPrompt(Task& task, std::optional<Task::CachedConflictDecision> cachedDecision = std::nullopt) noexcept
{
    {
        std::scoped_lock lock(task._conflictArbiter.mutex);
        if (cachedDecision.has_value())
        {
            StoreConflictDecisionInCacheLocked(task._conflictArbiter, cachedDecision->scope, cachedDecision->action);
        }
        task._conflictArbiter.prompt        = {};
        task._conflictArbiter.ownerThreadId = 0;
        task._conflictArbiter.decisionAction.reset();
        task._conflictArbiter.decisionApplyToAll = false;
    }

    if (task._conflictArbiter.decisionEvent)
    {
        static_cast<void>(ResetEvent(task._conflictArbiter.decisionEvent.get()));
    }

    task._conflictArbiter.cv.notify_all();
}

struct DeferredConsentPromptFacts
{
    FileOperations::DeferredConsentRisk risk = FileOperations::DeferredConsentRisk::InsufficientSpace;
    uint8_t overlapProblem                   = 0u; // R4-A02-1
    uint64_t overlapTaskId                   = 0u;
    bool allowConcurrentRun                  = false;
    bool allowLiveInvalidation               = true;
    std::optional<size_t> itemIndex;
    std::wstring destinationRootId;
    bool applyToAllEligible = false;
    bool itemCountKnown     = false;
    uint64_t itemCount      = 0;
    bool bytesKnown         = false;
    uint64_t bytes          = 0;
    std::optional<FileOperations::ProviderIdentitySnapshot> sourceIdentity;
    std::optional<FileOperations::ProviderIdentitySnapshot> destinationIdentity;
    std::wstring consentDetail; // C1: what a permanent-delete confirmation deletes
};

void SetConflictPromptLocked(Task& task,
                             ConflictBucket bucket,
                             HRESULT status,
                             std::wstring sourcePath,
                             std::wstring destinationPath,
                             Task::ConflictPromptState::ItemMetadata sourceMetadata,
                             Task::ConflictPromptState::ItemMetadata destinationMetadata,
                             Task::ConflictDecisionScope decisionScope,
                             const ConflictActionPolicy& actionPolicy,
                             unsigned int attemptCount,
                             const DeferredConsentPromptFacts* consentFacts = nullptr) noexcept
{
    if (task._conflictArbiter.decisionEvent)
    {
        static_cast<void>(ResetEvent(task._conflictArbiter.decisionEvent.get()));
    }

    Task::ConflictPromptState prompt{};
    prompt.active              = true;
    prompt.metadataLoading     = false;
    prompt.deferredConsent     = consentFacts != nullptr;
    prompt.bucket              = bucket;
    prompt.status              = status;
    prompt.sourcePath          = std::move(sourcePath);
    prompt.destinationPath     = std::move(destinationPath);
    prompt.sourceMetadata      = std::move(sourceMetadata);
    prompt.destinationMetadata = std::move(destinationMetadata);
    prompt.decisionScope       = std::move(decisionScope);
    prompt.applyToAllEligible  = actionPolicy.applyToAllEligible;
    prompt.applyToAllChecked   = false;
    prompt.skipAllEligible     = actionPolicy.skipAllEligible;
    prompt.buttonsPublishable  = actionPolicy.buttonsPublishable;
    prompt.retryFailed         = attemptCount != 0u;
    prompt.attemptCount        = attemptCount;
    if (consentFacts != nullptr)
    {
        prompt.consentDetail       = consentFacts->consentDetail;
        prompt.factItemCountKnown  = consentFacts->itemCountKnown;
        prompt.factItemCount       = consentFacts->itemCount;
        prompt.factBytesKnown      = consentFacts->bytesKnown;
        prompt.factBytes           = consentFacts->bytes;
        prompt.overlapProblem      = consentFacts->overlapProblem;
        prompt.overlapTaskId       = consentFacts->overlapTaskId;
        prompt.sourceIdentity      = consentFacts->sourceIdentity;
        prompt.destinationIdentity = consentFacts->destinationIdentity;
    }

    prompt.actions             = actionPolicy.actions;
    prompt.actionCount         = actionPolicy.actionCount;
    prompt.primaryActions      = actionPolicy.primaryActions;
    prompt.primaryActionCount  = actionPolicy.primaryActionCount;
    prompt.overflowActions     = actionPolicy.overflowActions;
    prompt.overflowActionCount = actionPolicy.overflowActionCount;
    prompt.defaultAction       = actionPolicy.defaultAction;
    prompt.escapeAction        = actionPolicy.escapeAction;

    task._conflictArbiter.prompt        = std::move(prompt);
    task._conflictArbiter.ownerThreadId = GetCurrentThreadId();

    task._conflictArbiter.decisionAction.reset();
    task._conflictArbiter.decisionApplyToAll = false;
}

void SetConflictPromptMetadataLoadingLocked(Task& task,
                                            ConflictBucket bucket,
                                            HRESULT status,
                                            std::wstring sourcePath,
                                            std::wstring destinationPath,
                                            unsigned int attemptCount,
                                            const DeferredConsentPromptFacts* consentFacts) noexcept
{
    if (task._conflictArbiter.decisionEvent)
    {
        static_cast<void>(ResetEvent(task._conflictArbiter.decisionEvent.get()));
    }

    Task::ConflictPromptState prompt{};
    prompt.active          = true;
    prompt.metadataLoading = true;
    prompt.deferredConsent = consentFacts != nullptr;
    prompt.bucket          = bucket;
    prompt.status          = status;
    prompt.sourcePath      = std::move(sourcePath);
    prompt.destinationPath = std::move(destinationPath);
    prompt.retryFailed     = attemptCount != 0u;
    prompt.attemptCount    = attemptCount;
    if (consentFacts != nullptr)
    {
        prompt.consentDetail       = consentFacts->consentDetail;
        prompt.factItemCountKnown  = consentFacts->itemCountKnown;
        prompt.factItemCount       = consentFacts->itemCount;
        prompt.factBytesKnown      = consentFacts->bytesKnown;
        prompt.factBytes           = consentFacts->bytes;
        prompt.overlapProblem      = consentFacts->overlapProblem;
        prompt.overlapTaskId       = consentFacts->overlapTaskId;
        prompt.sourceIdentity      = consentFacts->sourceIdentity;
        prompt.destinationIdentity = consentFacts->destinationIdentity;
    }

    task._conflictArbiter.prompt        = std::move(prompt);
    task._conflictArbiter.ownerThreadId = GetCurrentThreadId();
    task._conflictArbiter.decisionAction.reset();
    task._conflictArbiter.decisionApplyToAll = false;
}

struct ConflictPromptBeginResult
{
    ConflictAction action = ConflictAction::None;
    bool ownsPrompt       = false;
    bool fromCache        = false;
    ConflictBucket bucket = ConflictBucket::Unknown;
    Task::ConflictDecisionScope decisionScope{};
};

[[nodiscard]] ConflictPromptBeginResult BeginConflictPrompt(Task& task,
                                                            const Task::PerItemCallbackCookie* perItemCookie,
                                                            ConflictBucket bucket,
                                                            HRESULT status,
                                                            std::wstring_view sourcePath,
                                                            std::wstring_view destinationPath,
                                                            bool allowRetry,
                                                            unsigned int attemptCount,
                                                            bool ignoreCachedDecision,
                                                            bool allowDestructiveReplacement               = false,
                                                            const DeferredConsentPromptFacts* consentFacts = nullptr) noexcept
{
    auto [promptSourcePath, promptDestinationPath] = GetMostSpecificPathsForDiagnostics(task, perItemCookie, sourcePath, destinationPath);
    const bool sourceUsesWin32Metadata             = NavigationLocation::EqualsNoCase(task._sourcePluginId, L"builtin/file-system");
    const bool destinationUsesWin32Metadata        = NavigationLocation::EqualsNoCase(task._destinationPluginId, L"builtin/file-system") ||
                                                     (task._destinationPluginId.empty() && ! task._destinationFileSystem && sourceUsesWin32Metadata);

    std::unique_lock lock(task._conflictArbiter.mutex);
    task._conflictArbiter.cv.wait(lock, [&]() noexcept {
        return ! task._conflictArbiter.prompt.active || task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested();
    });

    if (task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested())
    {
        return {ConflictAction::Cancel, false, false, bucket, {}};
    }

    // Publish a stable, non-actionable attention state before any provider or Win32 metadata
    // read can block. No decision is accepted until the complete typed action set replaces it.
    SetConflictPromptMetadataLoadingLocked(task, bucket, status, promptSourcePath, promptDestinationPath, attemptCount, consentFacts);
    lock.unlock();

    task.RequestPresentationReveal();

    const uint64_t metadataStartUs             = PerfNowUs();
    const wil::com_ptr<IFileSystemIO> sourceIo = QueryFileSystemIo(task._fileSystem.get());
    wil::com_ptr<IFileSystemIO> destinationIo  = task._destinationFileSystem ? QueryFileSystemIo(task._destinationFileSystem.get()) : sourceIo;
    const Task::ConflictPromptState::ItemMetadata sourceMetadata =
        ReadConflictItemMetadata(task._fileSystem.get(), sourceIo.get(), promptSourcePath, sourceUsesWin32Metadata);
    const Task::ConflictPromptState::ItemMetadata destinationMetadata =
        ReadConflictItemMetadata(task._destinationFileSystem ? task._destinationFileSystem.get() : task._fileSystem.get(),
                                 destinationIo.get(),
                                 promptDestinationPath,
                                 destinationUsesWin32Metadata);
    const uint64_t metadataUs = PerfElapsedUs(metadataStartUs);
    task._perf.conflictMetadataUs.fetch_add(metadataUs, std::memory_order_relaxed);
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(
            L"FileOps.Conflict.MetadataPromptUs", L"", metadataUs, sourceMetadata.available ? 1u : 0u, destinationMetadata.available ? 1u : 0u, S_OK);
    }

    bucket                                    = RefineTypedConflictBucket(task._operation, bucket, sourceMetadata, destinationMetadata);
    const bool allowKeepBoth                  = IsKeepBothActionEligible(task, perItemCookie, bucket, promptDestinationPath);
    Task::ConflictDecisionScope decisionScope = BuildConflictDecisionScope(task, perItemCookie, bucket, sourceMetadata, destinationMetadata);
    if (consentFacts != nullptr)
    {
        decisionScope.destinationRootId  = consentFacts->destinationRootId;
        decisionScope.applyToAllEligible = consentFacts->applyToAllEligible;
    }
    const ConflictActionPolicy actionPolicy = BuildConflictActionPolicy(bucket,
                                                                        allowRetry,
                                                                        allowKeepBoth,
                                                                        decisionScope.applyToAllEligible,
                                                                        allowDestructiveReplacement,
                                                                        consentFacts != nullptr && consentFacts->allowConcurrentRun,
                                                                        consentFacts == nullptr || consentFacts->allowLiveInvalidation);

    lock.lock();

    if (task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested())
    {
        lock.unlock();
        ClearConflictPrompt(task);
        return {ConflictAction::Cancel, false, false, bucket, decisionScope};
    }

    if (! ignoreCachedDecision)
    {
        if (const std::optional<ConflictAction> cachedDecision =
                LoadEligibleConflictDecisionFromCacheLocked(task, perItemCookie, decisionScope, promptDestinationPath);
            cachedDecision.has_value() && actionPolicy.Contains(cachedDecision.value()))
        {
            lock.unlock();
            ClearConflictPrompt(task);
            return {cachedDecision.value(), false, true, bucket, decisionScope};
        }
    }

    SetConflictPromptLocked(task,
                            bucket,
                            status,
                            promptSourcePath,
                            promptDestinationPath,
                            sourceMetadata,
                            destinationMetadata,
                            decisionScope,
                            actionPolicy,
                            attemptCount,
                            consentFacts);
    lock.unlock();

    task.RequestActionablePromptPresentation();

    task.LogDiagnostic(FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                       status,
                       L"item.conflict.prompt",
                       attemptCount != 0u ? L"Conflict prompt shown after a failed attempt." : L"Conflict prompt shown for item.",
                       promptSourcePath,
                       promptDestinationPath);

    return {ConflictAction::None, true, false, bucket, decisionScope};
}

[[nodiscard]] ConflictPromptBeginResult BeginProviderReturnedConflictPrompt(Task& task,
                                                                            const Task::PerItemCallbackCookie* perItemCookie,
                                                                            ConflictBucket bucket,
                                                                            HRESULT status,
                                                                            std::wstring_view sourcePath,
                                                                            std::wstring_view destinationPath,
                                                                            bool allowRetry,
                                                                            unsigned int attemptCount,
                                                                            bool ignoreCachedDecision) noexcept
{
    // A conflict that escaped the provider/bridge callback has no exact destination authority
    // attached to it. The outer executor can offer non-destructive actions only.
    return BeginConflictPrompt(task, perItemCookie, bucket, status, sourcePath, destinationPath, allowRetry, attemptCount, ignoreCachedDecision, false);
}

[[nodiscard]] std::pair<ConflictAction, bool> WaitForConflictDecision(Task& task,
                                                                      const void* cookieKey,
                                                                      const Task::ConflictDecisionScope& decisionScope) noexcept
{
    bool consentPrompt = false;
    {
        std::scoped_lock lock(task._conflictArbiter.mutex);
        consentPrompt = task._conflictArbiter.prompt.deferredConsent;
    }
    const uint64_t perfStartUs   = PerfNowUs();
    const auto perfCallbackScope = wil::scope_exit([&] noexcept
    {
        const uint64_t waitUs = PerfElapsedUs(perfStartUs);
        if (consentPrompt)
        {
            task._perf.consentWaitUs += waitUs;
            ++task._perf.consentPromptCount;
        }
        else
        {
            task._perf.conflictWaitUs += waitUs;
            ++task._perf.conflictPromptCount;
        }
        NoteConflictWorkerWait(task, cookieKey, waitUs);
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(consentPrompt ? L"FileOps.Consent.WaitUs" : L"FileOps.Conflict.WaitUs", L"", waitUs, 0u, 0u, S_OK);
        }
    });

    if (! task._conflictArbiter.decisionEvent)
    {
        ClearConflictPrompt(task);
        return {ConflictAction::Cancel, false};
    }

    for (;;)
    {
        if (task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested())
        {
            ClearConflictPrompt(task);
            return {ConflictAction::Cancel, false};
        }

        const DWORD wait = WaitForSingleObject(task._conflictArbiter.decisionEvent.get(), 50);
        if (wait == WAIT_OBJECT_0)
        {
            break;
        }
    }

    ConflictAction action = ConflictAction::Cancel;
    bool applyToAll       = false;
    {
        std::scoped_lock lock(task._conflictArbiter.mutex);
        action     = task._conflictArbiter.decisionAction.value_or(ConflictAction::Cancel);
        applyToAll = task._conflictArbiter.decisionApplyToAll;
    }

    if (applyToAll && IsCacheableConflictDecision(action))
    {
        ClearConflictPrompt(task, Task::CachedConflictDecision{decisionScope, action});
        return {action, true};
    }

    ClearConflictPrompt(task);
    return {action, applyToAll};
}

[[nodiscard]] constexpr ConflictBucket ConflictBucketFromDeferredConsentRisk(FileOperations::DeferredConsentRisk risk) noexcept
{
    switch (risk)
    {
        case FileOperations::DeferredConsentRisk::InsufficientSpace: return ConflictBucket::InsufficientSpace;
        case FileOperations::DeferredConsentRisk::SpaceUnknown: return ConflictBucket::SpaceUnknown;
        case FileOperations::DeferredConsentRisk::EfsPlaintext: return ConflictBucket::EfsPlaintext;
        case FileOperations::DeferredConsentRisk::SparseInflation: return ConflictBucket::SparseInflation;
        case FileOperations::DeferredConsentRisk::PlaceholderHydration: return ConflictBucket::PlaceholderHydration;
        case FileOperations::DeferredConsentRisk::MetadataLoss: return ConflictBucket::MetadataLoss;
        case FileOperations::DeferredConsentRisk::RecycleEscalation: return ConflictBucket::RecycleFailed;
        case FileOperations::DeferredConsentRisk::SameHostOverlap: return ConflictBucket::SameHostOverlap;
        case FileOperations::DeferredConsentRisk::SameHostLiveOutput: return ConflictBucket::SameHostLiveOutput;
        case FileOperations::DeferredConsentRisk::PermanentDelete: return ConflictBucket::PermanentDeleteConfirmation;
        default: return ConflictBucket::Unknown;
    }
}

struct DeferredConsentRequest
{
    FileOperations::DeferredConsentRisk risk = FileOperations::DeferredConsentRisk::InsufficientSpace;
    uint8_t overlapProblem                   = 0u; // R4-A02-1
    uint64_t overlapTaskId                   = 0u;
    std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations> overlapTaskIds{};
    size_t overlapTaskIdCount                        = 0u;
    bool allowConcurrentRun                          = false;
    bool allowLiveInvalidation                       = true;
    HRESULT status                                   = E_FAIL;
    const Task::PerItemCallbackCookie* perItemCookie = nullptr;
    std::wstring_view sourcePath;
    std::wstring_view destinationPath;
    std::optional<size_t> itemIndex;
    std::wstring destinationRootId;
    bool applyToAllEligible = false;
    bool itemCountKnown     = false;
    uint64_t itemCount      = 0;
    bool bytesKnown         = false;
    uint64_t bytes          = 0;
    std::optional<FileOperations::ProviderIdentitySnapshot> sourceIdentity;
    std::optional<FileOperations::ProviderIdentitySnapshot> destinationIdentity;
    std::wstring consentDetail; // C1: what a permanent-delete confirmation deletes
};

struct DeferredConsentResult
{
    HRESULT status        = E_FAIL;
    ConflictAction action = ConflictAction::Cancel;
    bool applyToAll       = false;
};

[[nodiscard]] bool StoreDeferredConsentReceipt(Task& task, const DeferredConsentRequest& request, ConflictAction action, bool applyToAll) noexcept
{
    FileOperations::DeferredConsentReceipt receipt{};
    receipt.risk                = request.risk;
    receipt.taskId              = task._taskId;
    receipt.itemIndex           = applyToAll ? std::nullopt : request.itemIndex;
    receipt.destinationRootId   = request.destinationRootId;
    receipt.sourceIdentity      = request.sourceIdentity;
    receipt.destinationIdentity = request.destinationIdentity;
    if (request.overlapTaskIdCount > 0u)
    {
        auto related         = std::make_shared<FileOperations::SameHostOverlapRelationSet>();
        related->taskIds     = request.overlapTaskIds;
        related->taskIdCount = request.overlapTaskIdCount;
        receipt.relatedTasks = std::move(related);
    }
    receipt.decision = action;

    std::scoped_lock lock(task._conflictArbiter.mutex);
    for (std::optional<FileOperations::DeferredConsentReceipt>& existing : task._conflictArbiter.consentReceipts)
    {
        if (existing.has_value() && existing->risk == receipt.risk && existing->taskId == receipt.taskId && existing->itemIndex == receipt.itemIndex &&
            existing->destinationRootId == receipt.destinationRootId)
        {
            existing = std::move(receipt);
            return true;
        }
    }
    const auto empty = std::ranges::find_if(task._conflictArbiter.consentReceipts,
                                            [](const std::optional<FileOperations::DeferredConsentReceipt>& entry) noexcept { return ! entry.has_value(); });
    if (empty == task._conflictArbiter.consentReceipts.end())
    {
        return false;
    }
    *empty = std::move(receipt);
    return true;
}

[[nodiscard]] DeferredConsentResult RequestDeferredConsent(Task& task, const DeferredConsentRequest& request) noexcept
{
    DeferredConsentPromptFacts facts{};
    facts.risk                  = request.risk;
    facts.overlapProblem        = request.overlapProblem;
    facts.overlapTaskId         = request.overlapTaskId;
    facts.allowConcurrentRun    = request.allowConcurrentRun;
    facts.allowLiveInvalidation = request.allowLiveInvalidation;
    facts.itemIndex             = request.itemIndex;
    facts.destinationRootId     = request.destinationRootId;
    facts.applyToAllEligible    = request.applyToAllEligible && request.risk != FileOperations::DeferredConsentRisk::RecycleEscalation;
    facts.itemCountKnown        = request.itemCountKnown;
    facts.itemCount             = request.itemCount;
    facts.bytesKnown            = request.bytesKnown;
    facts.bytes                 = request.bytes;
    facts.sourceIdentity        = request.sourceIdentity;
    facts.destinationIdentity   = request.destinationIdentity;
    facts.consentDetail         = request.consentDetail;

    const ConflictPromptBeginResult promptBegin = BeginConflictPrompt(task,
                                                                      request.perItemCookie,
                                                                      ConflictBucketFromDeferredConsentRisk(request.risk),
                                                                      request.status,
                                                                      request.sourcePath,
                                                                      request.destinationPath,
                                                                      false,
                                                                      0u,
                                                                      false,
                                                                      false,
                                                                      &facts);
    ConflictAction action                       = promptBegin.action;
    bool applyToAll                             = false;
    if (promptBegin.ownsPrompt)
    {
        const auto decision = WaitForConflictDecision(task, request.perItemCookie, promptBegin.decisionScope);
        action              = decision.first;
        applyToAll          = decision.second;
    }

    if (action == ConflictAction::Cancel)
    {
        return {HRESULT_FROM_WIN32(ERROR_CANCELLED), action, false};
    }
    if (promptBegin.fromCache)
    {
        return {S_OK, action, true};
    }
    if (! StoreDeferredConsentReceipt(task, request, action, applyToAll))
    {
        task.LogDiagnostic(FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                           HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY),
                           L"consent.receipt.limit",
                           L"The bounded deferred-consent receipt table is full; the affected mutation was not authorized.",
                           request.sourcePath,
                           request.destinationPath);
        return {HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY), ConflictAction::Cancel, false};
    }
    return {S_OK, action, applyToAll};
}

[[nodiscard]] bool IsModifierConflictAction(ConflictAction action) noexcept
{
    switch (action)
    {
        case ConflictAction::Overwrite:
        case ConflictAction::ReplaceReadOnly:
        case ConflictAction::ReplaceLink:
        case ConflictAction::PermanentDelete: return true;
        case ConflictAction::KeepBoth:
        case ConflictAction::Proceed:
        case ConflictAction::RetainSource:
        case ConflictAction::None:
        case ConflictAction::Retry:
        case ConflictAction::Skip:
        case ConflictAction::SkipAll:
        case ConflictAction::RunConcurrently:
        case ConflictAction::QueueUntilOtherTaskFinishes:
        case ConflictAction::InvalidateLiveOutput:
        case ConflictAction::Cancel:
        default: return false;
    }
}

class PerItemExecutionPolicy
{
public:
    virtual ~PerItemExecutionPolicy()                            = default;
    [[nodiscard]] virtual HRESULT Process(size_t index) noexcept = 0;
};

template <typename Callable> class ReferencedPerItemExecutionPolicy final : public PerItemExecutionPolicy
{
public:
    explicit ReferencedPerItemExecutionPolicy(Callable& callable) noexcept : _callable(callable)
    {
    }

    ReferencedPerItemExecutionPolicy(const ReferencedPerItemExecutionPolicy&)            = delete;
    ReferencedPerItemExecutionPolicy& operator=(const ReferencedPerItemExecutionPolicy&) = delete;

    [[nodiscard]] HRESULT Process(size_t index) noexcept override
    {
        return _callable(index);
    }

private:
    Callable& _callable;
};

class PerItemTaskScheduler final
{
public:
    enum class FailurePolicy : uint8_t
    {
        RecordOnly,
        CancelTask,
    };

    PerItemTaskScheduler() = default;
    ~PerItemTaskScheduler() noexcept
    {
        Shutdown();
    }

    PerItemTaskScheduler(const PerItemTaskScheduler&)            = delete;
    PerItemTaskScheduler(PerItemTaskScheduler&&)                 = delete;
    PerItemTaskScheduler& operator=(const PerItemTaskScheduler&) = delete;
    PerItemTaskScheduler& operator=(PerItemTaskScheduler&&)      = delete;

    struct PerfSnapshot final
    {
        uint64_t dequeueAttempts = 0;
        uint64_t dequeueSuccess  = 0;
        uint64_t waitForWorkUs   = 0;
        uint64_t processIndexUs  = 0;
    };

    struct Job final
    {
        Job() noexcept = default;

        Job(const Job&)            = delete;
        Job(Job&&)                 = delete;
        Job& operator=(const Job&) = delete;
        Job& operator=(Job&&)      = delete;

        Task* task = nullptr;
        // Non-owning stack references. WaitJob and Shutdown join every callback before the
        // operation context can unwind.
        PerItemExecutionPolicy* itemPolicy = nullptr;
        std::atomic<HRESULT>* firstFailure = nullptr;
        FailurePolicy failurePolicy        = FailurePolicy::RecordOnly;
        size_t totalItems                  = 0;
        unsigned int maxConcurrency        = 1;

        // Protected by the scheduler mutex.
        size_t nextIndex      = 0;
        unsigned int inFlight = 0;
        bool done             = false;

        std::atomic<uint64_t> perfDequeueAttempts{0};
        std::atomic<uint64_t> perfDequeueSuccess{0};
        std::atomic<uint64_t> perfWaitForWorkUs{0};
        std::atomic<uint64_t> perfProcessIndexUs{0};

        std::mutex doneMutex;
        std::condition_variable doneCv;
    };

    using JobPtr = std::shared_ptr<Job>;

    [[nodiscard]] PerfSnapshot CapturePerfSnapshot() const noexcept
    {
        PerfSnapshot snapshot{};
        snapshot.dequeueAttempts = _perfDequeueAttempts.load(std::memory_order_acquire);
        snapshot.dequeueSuccess  = _perfDequeueSuccess.load(std::memory_order_acquire);
        snapshot.waitForWorkUs   = _perfWaitForWorkUs.load(std::memory_order_acquire);
        snapshot.processIndexUs  = _perfProcessIndexUs.load(std::memory_order_acquire);
        return snapshot;
    }

    [[nodiscard]] PerfSnapshot SnapshotPerf(const JobPtr& job) const noexcept
    {
        PerfSnapshot snapshot{};
        if (! job)
        {
            return snapshot;
        }

        snapshot.dequeueAttempts = job->perfDequeueAttempts.load(std::memory_order_acquire);
        snapshot.dequeueSuccess  = job->perfDequeueSuccess.load(std::memory_order_acquire);
        snapshot.waitForWorkUs   = job->perfWaitForWorkUs.load(std::memory_order_acquire);
        snapshot.processIndexUs  = job->perfProcessIndexUs.load(std::memory_order_acquire);
        return snapshot;
    }

    JobPtr StartJob(Task* task,
                    unsigned int maxConcurrency,
                    size_t totalItems,
                    PerItemExecutionPolicy& itemPolicy,
                    std::atomic<HRESULT>& firstFailure,
                    FailurePolicy failurePolicy)
    {
        auto job            = std::make_shared<Job>();
        job->task           = task;
        job->totalItems     = totalItems;
        job->itemPolicy     = &itemPolicy;
        job->firstFailure   = &firstFailure;
        job->failurePolicy  = failurePolicy;
        job->maxConcurrency = std::max(1u, maxConcurrency);

        ensureWorkers();

        if (_workers.empty())
        {
            for (size_t i = 0; i < job->totalItems; ++i)
            {
                if (isTaskCancelled(*job))
                {
                    break;
                }
                const HRESULT itemHr = job->itemPolicy->Process(i);
                recordItemResult(*job, itemHr);
            }

            {
                std::scoped_lock lock(job->doneMutex);
                job->done = true;
            }
            job->doneCv.notify_all();
            return job;
        }

        {
            std::scoped_lock lock(_mutex);
            _jobs.push_back(job);
            _rrCursor = _jobs.size() - 1u; // Bias next dequeue to the newly-added job to reduce start latency/starvation.
        }

        _cv.notify_all();
        return job;
    }

    void WaitJob(const JobPtr& job) noexcept
    {
        if (! job)
        {
            return;
        }

        if (s_currentScheduler != this)
        {
            std::unique_lock lock(job->doneMutex);
            job->doneCv.wait(lock, [&]() noexcept { return job->done; });
            return;
        }

        // A scheduler worker may create a nested job (the bridge directory producer does this for
        // within-folder file copies). Blocking that worker here can park every worker behind nested
        // jobs that nobody is left to execute. Participate in the target job until it completes.
        for (;;)
        {
            {
                std::scoped_lock doneLock(job->doneMutex);
                if (job->done)
                {
                    return;
                }
            }

            JobPtr workJob;
            size_t index = 0;
            {
                std::scoped_lock schedulerLock(_mutex);
                cleanupJobsLocked();
                static_cast<void>(tryDequeueFromJobLocked(job, workJob, index));
            }

            if (workJob)
            {
                processDequeuedWork(workJob, index);
                continue;
            }

            std::unique_lock doneLock(job->doneMutex);
            job->doneCv.wait_for(doneLock, std::chrono::milliseconds(10), [&]() noexcept { return job->done; });
        }
    }

    void NotifyWorkAvailable() noexcept
    {
        _cv.notify_all();
    }

#ifdef ENABLE_TESTS
    [[nodiscard]] bool EnsureWorkersAvailableForSelfTest() noexcept
    {
        ensureWorkers();
        return ! _workers.empty();
    }

    [[nodiscard]] unsigned int WorkerCountForSelfTest() const noexcept
    {
        return _workerCount.load(std::memory_order_acquire);
    }
#endif

    void Shutdown() noexcept
    {
        std::vector<std::jthread> workers;
        {
            std::scoped_lock lock(_initMutex);
            if (! _initialized)
            {
                return;
            }

            for (std::jthread& worker : _workers)
            {
                worker.request_stop();
            }

            workers      = std::move(_workers);
            _initialized = false;
            _workerCount.store(0u, std::memory_order_release);
        }

        _cv.notify_all();

        // FileOperationState stops producers before shutting down this scheduler.
        // Join workers first so stack-backed operation contexts cannot unwind while callbacks are still active.
        workers.clear();

        {
            std::scoped_lock lock(_mutex);
            for (const JobPtr& job : _jobs)
            {
                finishJobLocked(job);
            }
            _jobs.clear();
            _rrCursor = 0;
        }

        _cv.notify_all();
    }

private:
    [[nodiscard]] size_t countActiveJobsLocked() const noexcept
    {
        size_t active = 0;
        for (const JobPtr& job : _jobs)
        {
            if (! job)
            {
                continue;
            }

            if (isTaskCancelled(*job) || isTaskPaused(*job))
            {
                continue;
            }

            if (job->nextIndex >= job->totalItems)
            {
                continue;
            }

            ++active;
        }

        return active;
    }

    [[nodiscard]] unsigned int effectiveMaxConcurrencyLocked(const Job& job) const noexcept
    {
        unsigned int maxConc = std::max(1u, job.maxConcurrency);

        // Sharing policy: when multiple jobs are active, avoid letting a single job occupy every worker thread.
        // Keep at least one worker available for each other active job so new tasks can start promptly.
        const unsigned int workerCount = _workerCount.load(std::memory_order_acquire);
        if (workerCount <= 1u)
        {
            if (Debug::Perf::IsCaptureEnabled())
            {
                Debug::Perf::Emit(L"FileOps.Scheduler.EffectiveConcurrency",
                                  std::format(L"maxConcurrency={} activeJobs={} workerCount={} cap={}", job.maxConcurrency, 0u, workerCount, 1u),
                                  0,
                                  1u,
                                  workerCount,
                                  S_OK);
            }
            return 1u;
        }

        const size_t activeJobs = countActiveJobsLocked();
        if (activeJobs <= 1u)
        {
            // Starvation guard: reserve one worker so a second job can begin without waiting for an in-flight
            // long-running file operation to complete.
            if (job.totalItems > 1u)
            {
                maxConc = std::min<unsigned int>(maxConc, workerCount - 1u);
            }

            const unsigned int cap = std::max(1u, maxConc);
            if (Debug::Perf::IsCaptureEnabled())
            {
                Debug::Perf::Emit(L"FileOps.Scheduler.EffectiveConcurrency",
                                  std::format(L"maxConcurrency={} activeJobs={} workerCount={} cap={}", job.maxConcurrency, activeJobs, workerCount, cap),
                                  0,
                                  cap,
                                  workerCount,
                                  S_OK);
            }
            return cap;
        }

        const unsigned int cap    = (activeJobs >= static_cast<size_t>(workerCount)) ? 1u : (workerCount - static_cast<unsigned int>(activeJobs) + 1u);
        maxConc                   = std::min<unsigned int>(maxConc, std::max(1u, cap));
        const unsigned int result = std::max(1u, maxConc);
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Scheduler.EffectiveConcurrency",
                              std::format(L"maxConcurrency={} activeJobs={} workerCount={} cap={}", job.maxConcurrency, activeJobs, workerCount, result),
                              0,
                              result,
                              workerCount,
                              S_OK);
        }
        return result;
    }

    void ensureWorkers()
    {
        std::scoped_lock lock(_initMutex);
        if (_initialized)
        {
            return;
        }

        unsigned int workerCount = std::thread::hardware_concurrency();
        if (workerCount == 0)
        {
            workerCount = 4;
        }

        constexpr unsigned int kMaxWorkers = static_cast<unsigned int>(Task::kMaxInFlightFiles);
        workerCount                        = std::max(1u, std::min(workerCount, kMaxWorkers));
        _workerCount.store(workerCount, std::memory_order_release);

        _workers.reserve(workerCount);
        for (unsigned int i = 0; i < workerCount; ++i)
        {
            try
            {
                _workers.emplace_back([this](std::stop_token stopToken) noexcept { workerMain(stopToken); });
            }
            catch (const std::system_error&)
            {
                break;
            }
        }

        _workerCount.store(static_cast<unsigned int>(_workers.size()), std::memory_order_release);
        _initialized = true;
    }

    [[nodiscard]] bool isTaskCancelled(const Job& job) const noexcept
    {
        if (! job.task)
        {
            return true;
        }
        return job.task->_cancelled.load(std::memory_order_acquire) || job.task->_stopToken.stop_requested();
    }

    [[nodiscard]] bool isTaskPaused(const Job& job) const noexcept
    {
        if (! job.task)
        {
            return false;
        }
        return job.task->IsPaused() || job.task->IsQueuePaused();
    }

    void finishJobLocked(const JobPtr& job) noexcept
    {
        if (! job)
        {
            return;
        }

        {
            std::scoped_lock lock(job->doneMutex);
            job->done = true;
        }
        job->doneCv.notify_all();
    }

    void cleanupJobsLocked() noexcept
    {
        size_t write = 0;
        for (size_t read = 0; read < _jobs.size(); ++read)
        {
            const JobPtr& job = _jobs[read];
            if (! job)
            {
                continue;
            }

            const bool cancelled = isTaskCancelled(*job);
            const bool finished  = job->nextIndex >= job->totalItems;
            if ((cancelled || finished) && job->inFlight == 0)
            {
                finishJobLocked(job);
                continue;
            }

            if (write != read)
            {
                _jobs[write] = job;
            }
            ++write;
        }

        if (write < _jobs.size())
        {
            _jobs.resize(write);
        }

        if (_rrCursor >= _jobs.size())
        {
            _rrCursor = 0;
        }
    }

    [[nodiscard]] bool hasSchedulableWorkLocked() noexcept
    {
        cleanupJobsLocked();

        for (const JobPtr& job : _jobs)
        {
            if (! job)
            {
                continue;
            }

            if (isTaskCancelled(*job) || isTaskPaused(*job))
            {
                continue;
            }

            if (job->inFlight >= effectiveMaxConcurrencyLocked(*job))
            {
                continue;
            }

            if (job->nextIndex >= job->totalItems)
            {
                continue;
            }

            return true;
        }

        return false;
    }

    [[nodiscard]] bool tryDequeueWorkLocked(JobPtr& outJob, size_t& outIndex) noexcept
    {
        _perfDequeueAttempts.fetch_add(1u, std::memory_order_relaxed);
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Scheduler.DequeueAttempts", L"", 0, 1u, 0u, S_OK);
        }

        const size_t jobCount = _jobs.size();
        if (jobCount == 0)
        {
            return false;
        }

        const size_t start = (jobCount > 0) ? (_rrCursor % jobCount) : 0;
        for (size_t attempt = 0; attempt < jobCount; ++attempt)
        {
            const size_t idx = (start + attempt) % jobCount;
            JobPtr& job      = _jobs[idx];
            if (! job)
            {
                continue;
            }

            if (isTaskCancelled(*job) || isTaskPaused(*job))
            {
                continue;
            }

            if (job->inFlight >= effectiveMaxConcurrencyLocked(*job))
            {
                continue;
            }

            if (job->nextIndex >= job->totalItems)
            {
                continue;
            }

            outJob   = job;
            outIndex = job->nextIndex;
            job->nextIndex += 1;
            job->inFlight += 1;

            _perfDequeueSuccess.fetch_add(1u, std::memory_order_relaxed);
            if (Debug::Perf::IsCaptureEnabled())
            {
                Debug::Perf::Emit(L"FileOps.Scheduler.DequeueSuccess", L"", 0, 1u, 0u, S_OK);
                Debug::Perf::Emit(L"FileOps.Scheduler.ScanJobsPerAttempt", L"", 0, static_cast<uint64_t>(attempt + 1u), 0u, S_OK);
            }

            _rrCursor = (idx + 1u) % jobCount;
            return true;
        }

        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Scheduler.ScanJobsPerAttempt", L"", 0, static_cast<uint64_t>(jobCount), 0u, S_OK);
        }
        return false;
    }

    [[nodiscard]] bool tryDequeueFromJobLocked(const JobPtr& requestedJob, JobPtr& outJob, size_t& outIndex) noexcept
    {
        if (! requestedJob || isTaskCancelled(*requestedJob) || isTaskPaused(*requestedJob) ||
            requestedJob->inFlight >= effectiveMaxConcurrencyLocked(*requestedJob) || requestedJob->nextIndex >= requestedJob->totalItems)
        {
            return false;
        }

        const auto it = std::find(_jobs.begin(), _jobs.end(), requestedJob);
        if (it == _jobs.end())
        {
            return false;
        }

        outJob   = requestedJob;
        outIndex = requestedJob->nextIndex;
        requestedJob->nextIndex += 1u;
        requestedJob->inFlight += 1u;
        _perfDequeueAttempts.fetch_add(1u, std::memory_order_relaxed);
        _perfDequeueSuccess.fetch_add(1u, std::memory_order_relaxed);
        return true;
    }

    void processDequeuedWork(const JobPtr& job, size_t index) noexcept
    {
        if (job && job->itemPolicy)
        {
            const uint64_t processStartUs = PerfNowUs();
            // R0f: a provider call wedged on a dead share returns once this task is canceled.
            const Common::SynchronousIoCancelWatch::Scope cancelWatch(job->task != nullptr ? &Task::ShouldCancelSynchronousIo : nullptr, job->task);
            const HRESULT itemHr = job->itemPolicy->Process(index);
            recordItemResult(*job, itemHr);
            const uint64_t processUs = PerfElapsedUs(processStartUs);
            _perfProcessIndexUs.fetch_add(processUs, std::memory_order_relaxed);
            if (Debug::Perf::IsCaptureEnabled())
            {
                Debug::Perf::Emit(L"FileOps.Scheduler.ProcessIndexUs", L"", processUs, static_cast<uint64_t>(index), 0u, S_OK);
            }
        }

        {
            std::scoped_lock lock(_mutex);
            if (job && job->inFlight > 0)
            {
                job->inFlight -= 1u;
            }
            cleanupJobsLocked();
        }
        _cv.notify_all();
    }

    static void recordItemResult(const Job& job, HRESULT itemHr) noexcept
    {
        if (SUCCEEDED(itemHr) || job.firstFailure == nullptr)
        {
            return;
        }
        HRESULT expected = S_OK;
        job.firstFailure->compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel);
        if (job.failurePolicy == FailurePolicy::CancelTask && job.task != nullptr)
        {
            job.task->RequestCancel();
        }
    }

    void workerMain(std::stop_token stopToken) noexcept
    {
        [[maybe_unused]] auto coInit                  = wil::CoInitializeEx_failfast();
        PerItemTaskScheduler* const previousScheduler = s_currentScheduler;
        s_currentScheduler                            = this;
        const auto restoreScheduler                   = wil::scope_exit([&]() noexcept { s_currentScheduler = previousScheduler; });

        for (;;)
        {
            JobPtr job;
            size_t index = 0;
            {
                std::unique_lock lock(_mutex);
                const uint64_t waitStartUs = PerfNowUs();
                _cv.wait(lock, [&]() noexcept { return stopToken.stop_requested() || hasSchedulableWorkLocked(); });
                const uint64_t waitUs = PerfElapsedUs(waitStartUs);
                _perfWaitForWorkUs.fetch_add(waitUs, std::memory_order_relaxed);
                if (Debug::Perf::IsCaptureEnabled())
                {
                    Debug::Perf::Emit(L"FileOps.Scheduler.WaitForWorkUs", L"", waitUs, 0u, 0u, S_OK);
                }
                if (stopToken.stop_requested())
                {
                    return;
                }

                cleanupJobsLocked();
                if (! tryDequeueWorkLocked(job, index))
                {
                    continue;
                }
            }

            processDequeuedWork(job, index);
        }
    }

private:
    std::mutex _mutex;
    std::condition_variable _cv;
    std::vector<JobPtr> _jobs;
    size_t _rrCursor = 0;

    std::mutex _initMutex;
    bool _initialized = false;
    std::vector<std::jthread> _workers;
    std::atomic<unsigned int> _workerCount{0};
    inline static thread_local PerItemTaskScheduler* s_currentScheduler = nullptr;
    std::atomic<uint64_t> _perfDequeueAttempts{0};
    std::atomic<uint64_t> _perfDequeueSuccess{0};
    std::atomic<uint64_t> _perfWaitForWorkUs{0};
    std::atomic<uint64_t> _perfProcessIndexUs{0};
};

PerItemTaskScheduler& GetPerItemTaskScheduler() noexcept
{
    static PerItemTaskScheduler scheduler;
    return scheduler;
}

#ifdef ENABLE_TESTS
bool RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTestInternal(FolderWindow::FileOperationState& state) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;

    Task task(state);
    PerItemTaskScheduler scheduler;
    if (! scheduler.EnsureWorkersAvailableForSelfTest())
    {
        Debug::Error(L"FileOps host scheduler selftest failed: Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker should create workers.");
        return false;
    }

    struct ProbeState
    {
        ProbeState()                             = default;
        ProbeState(const ProbeState&)            = delete;
        ProbeState(ProbeState&&)                 = delete;
        ProbeState& operator=(const ProbeState&) = delete;
        ProbeState& operator=(ProbeState&&)      = delete;

        std::mutex mutex;
        std::condition_variable cv;
        bool workerEntered    = false;
        bool releaseWorker    = false;
        bool callbackActive   = false;
        bool workerExited     = false;
        bool waiterReturned   = false;
        bool shutdownReturned = false;
    } probe;

    const auto blockedWorker = [&](size_t) noexcept -> HRESULT
    {
        std::unique_lock lock(probe.mutex);
        probe.workerEntered  = true;
        probe.callbackActive = true;
        probe.cv.notify_all();
        probe.cv.wait(lock, [&]() noexcept { return probe.releaseWorker; });
        probe.callbackActive = false;
        probe.workerExited   = true;
        probe.cv.notify_all();
        return S_OK;
    };
    ReferencedPerItemExecutionPolicy blockedWorkerPolicy(blockedWorker);
    std::atomic<HRESULT> firstFailure{S_OK};
    const auto job = scheduler.StartJob(&task, 1u, 1u, blockedWorkerPolicy, firstFailure, PerItemTaskScheduler::FailurePolicy::RecordOnly);

    const auto waitFor = [&](auto predicate, std::chrono::milliseconds timeout) noexcept -> bool
    {
        std::unique_lock lock(probe.mutex);
        return probe.cv.wait_for(lock, timeout, predicate);
    };

    if (! waitFor([&]() noexcept { return probe.workerEntered; }, std::chrono::milliseconds(5000)))
    {
        {
            std::scoped_lock lock(probe.mutex);
            probe.releaseWorker = true;
        }
        probe.cv.notify_all();
        scheduler.Shutdown();
        Debug::Error(
            L"FileOps host scheduler selftest failed: Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker should enter the blocked worker callback.");
        return false;
    }

    std::jthread waiter([&]() noexcept
    {
        scheduler.WaitJob(job);
        {
            std::scoped_lock lock(probe.mutex);
            probe.waiterReturned = true;
        }
        probe.cv.notify_all();
    });

    std::jthread shutdownThread([&]() noexcept
    {
        scheduler.Shutdown();
        {
            std::scoped_lock lock(probe.mutex);
            probe.shutdownReturned = true;
        }
        probe.cv.notify_all();
    });

    static_cast<void>(waitFor([&]() noexcept { return probe.waiterReturned || probe.shutdownReturned; }, std::chrono::milliseconds(250)));

    bool waiterReturnedEarly   = false;
    bool shutdownReturnedEarly = false;
    {
        std::scoped_lock lock(probe.mutex);
        waiterReturnedEarly   = probe.waiterReturned;
        shutdownReturnedEarly = probe.shutdownReturned;
        probe.releaseWorker   = true;
    }
    probe.cv.notify_all();

    const bool drained =
        waitFor([&]() noexcept { return probe.workerExited && probe.waiterReturned && probe.shutdownReturned; }, std::chrono::milliseconds(5000));
    if (waiterReturnedEarly)
    {
        Debug::Error(
            L"FileOps host scheduler selftest failed: Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker released WaitJob before the worker exited.");
    }
    if (shutdownReturnedEarly)
    {
        Debug::Error(L"FileOps host scheduler selftest failed: Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker returned from shutdown before the "
                     L"worker exited.");
    }
    if (! drained)
    {
        Debug::Error(L"FileOps host scheduler selftest failed: Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker did not drain after release.");
    }

    return ! waiterReturnedEarly && ! shutdownReturnedEarly && drained;
}

bool RunFileOpsPerItemSchedulerNestedSaturationSelfTestForSelfTestInternal(FolderWindow::FileOperationState& state) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;

    Task task(state);
    PerItemTaskScheduler scheduler;
    if (! scheduler.EnsureWorkersAvailableForSelfTest())
    {
        Debug::Error(L"FileOps nested scheduler saturation selftest could not create workers.");
        return false;
    }

    const unsigned int workerCount = scheduler.WorkerCountForSelfTest();
    if (workerCount == 0u)
    {
        Debug::Error(L"FileOps nested scheduler saturation selftest observed zero workers.");
        return false;
    }

    struct SaturationState final
    {
        SaturationState()                                  = default;
        SaturationState(const SaturationState&)            = delete;
        SaturationState(SaturationState&&)                 = delete;
        SaturationState& operator=(const SaturationState&) = delete;
        SaturationState& operator=(SaturationState&&)      = delete;

        std::mutex mutex;
        std::condition_variable cv;
        unsigned int entered = 0u;
        bool release         = false;
        bool timedOut        = false;
        std::atomic<unsigned int> nestedCompleted{0u};
    } saturation;

    std::atomic<HRESULT> outerFirstFailure{S_OK};
    const auto outerWorker = [&](size_t) noexcept -> HRESULT
    {
        {
            std::unique_lock lock(saturation.mutex);
            ++saturation.entered;
            if (saturation.entered == workerCount)
            {
                saturation.release = true;
                saturation.cv.notify_all();
            }
            else if (! saturation.cv.wait_for(lock, std::chrono::seconds(5), [&]() noexcept { return saturation.release; }))
            {
                saturation.timedOut = true;
                saturation.release  = true;
                saturation.cv.notify_all();
            }
        }

        const auto nestedWorker = [&](size_t) noexcept -> HRESULT
        {
            saturation.nestedCompleted.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        };
        ReferencedPerItemExecutionPolicy nestedPolicy(nestedWorker);
        std::atomic<HRESULT> nestedFirstFailure{S_OK};
        const auto nestedJob = scheduler.StartJob(&task, 1u, 1u, nestedPolicy, nestedFirstFailure, PerItemTaskScheduler::FailurePolicy::RecordOnly);
        scheduler.WaitJob(nestedJob);
        return S_OK;
    };
    ReferencedPerItemExecutionPolicy outerPolicy(outerWorker);

    std::vector<PerItemTaskScheduler::JobPtr> outerJobs;
    outerJobs.reserve(workerCount);
    for (unsigned int outerIndex = 0u; outerIndex < workerCount; ++outerIndex)
    {
        outerJobs.push_back(scheduler.StartJob(&task, 1u, 1u, outerPolicy, outerFirstFailure, PerItemTaskScheduler::FailurePolicy::RecordOnly));
    }

    for (const auto& outerJob : outerJobs)
    {
        scheduler.WaitJob(outerJob);
    }
    scheduler.Shutdown();

    const unsigned int nestedCompleted = saturation.nestedCompleted.load(std::memory_order_acquire);
    if (saturation.timedOut || saturation.entered != workerCount || nestedCompleted != workerCount)
    {
        Debug::Error(L"FileOps nested scheduler saturation selftest failed: workers={} entered={} nestedCompleted={} timedOut={}.",
                     workerCount,
                     saturation.entered,
                     nestedCompleted,
                     saturation.timedOut ? 1 : 0);
        return false;
    }
    return true;
}

bool RunFileOpsPerItemSchedulerFailurePolicySelfTestForSelfTestInternal(FolderWindow::FileOperationState& state) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;

    PerItemTaskScheduler scheduler;
    if (! scheduler.EnsureWorkersAvailableForSelfTest())
    {
        Debug::Error(L"FileOps scheduler failure-policy selftest could not create workers.");
        return false;
    }

    Task recordOnlyTask(state);
    std::atomic_uint32_t recordOnlyProcessed{0u};
    const auto recordOnlyWorker = [&](const size_t index) noexcept -> HRESULT
    {
        recordOnlyProcessed.fetch_add(1u, std::memory_order_acq_rel);
        return index == 0u ? E_OUTOFMEMORY : S_OK;
    };
    ReferencedPerItemExecutionPolicy recordOnlyPolicy(recordOnlyWorker);
    std::atomic<HRESULT> recordOnlyFailure{S_OK};
    const auto recordOnlyJob =
        scheduler.StartJob(&recordOnlyTask, 2u, 2u, recordOnlyPolicy, recordOnlyFailure, PerItemTaskScheduler::FailurePolicy::RecordOnly);
    scheduler.WaitJob(recordOnlyJob);

    Task cancelTask(state);
    const auto cancelWorker = [](size_t) noexcept -> HRESULT { return E_OUTOFMEMORY; };
    ReferencedPerItemExecutionPolicy cancelPolicy(cancelWorker);
    std::atomic<HRESULT> cancelFailure{S_OK};
    const auto cancelJob = scheduler.StartJob(&cancelTask, 1u, 1u, cancelPolicy, cancelFailure, PerItemTaskScheduler::FailurePolicy::CancelTask);
    scheduler.WaitJob(cancelJob);
    scheduler.Shutdown();

    const bool passed = recordOnlyFailure.load(std::memory_order_acquire) == E_OUTOFMEMORY && recordOnlyProcessed.load(std::memory_order_acquire) == 2u &&
                        ! recordOnlyTask._cancelled.load(std::memory_order_acquire) && cancelFailure.load(std::memory_order_acquire) == E_OUTOFMEMORY &&
                        cancelTask._cancelled.load(std::memory_order_acquire);
    if (! passed)
    {
        Debug::Error(L"FileOps scheduler failure-policy selftest failed: recordHr=0x{:08X}, recordProcessed={}, recordCancelled={}, cancelHr=0x{:08X}, "
                     L"cancelCancelled={}.",
                     static_cast<unsigned long>(recordOnlyFailure.load(std::memory_order_acquire)),
                     recordOnlyProcessed.load(std::memory_order_acquire),
                     recordOnlyTask._cancelled.load(std::memory_order_acquire) ? 1u : 0u,
                     static_cast<unsigned long>(cancelFailure.load(std::memory_order_acquire)),
                     cancelTask._cancelled.load(std::memory_order_acquire) ? 1u : 0u);
    }
    return passed;
}

bool RunFileOpsWorkerStartGateCancellationSelfTestForSelfTestInternal(FolderWindow::FileOperationState& state) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;
    using namespace std::chrono_literals;

    const auto runScenario = [&](const bool requestStop) noexcept -> bool
    {
        Task task(state);
        // This focused test owns no published task. A null state makes ThreadMain return
        // immediately after the start gate while still exercising the real wait/wake path.
        task._state = nullptr;
        std::atomic_bool returned{false};
        std::jthread worker([&](const std::stop_token stopToken) noexcept
        {
            task.ThreadMain(stopToken);
            returned.store(true, std::memory_order_release);
        });

        const auto waitUntil = [](const auto& predicate, const std::chrono::milliseconds timeout) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + timeout;
            while (! predicate() && std::chrono::steady_clock::now() < deadline)
            {
                std::this_thread::sleep_for(1ms);
            }
            return predicate();
        };
        const bool entered = waitUntil([&]() noexcept { return task._debugWorkerStartGateWaiting.load(std::memory_order_acquire); }, 2s);
        if (requestStop)
        {
            worker.request_stop();
        }
        else
        {
            task.RequestCancel();
        }
        const bool woke = waitUntil([&]() noexcept { return returned.load(std::memory_order_acquire); }, 2s);
        if (! woke)
        {
            task._workerReleased.store(true, std::memory_order_release);
            task._workerReleased.notify_all();
            worker.request_stop();
        }
        return entered && woke;
    };

    const bool cancelPassed = runScenario(false);
    const bool stopPassed   = runScenario(true);
    if (! cancelPassed || ! stopPassed)
    {
        Debug::Error(L"FileOps worker start-gate selftest failed: cancelWake={}, stopWake={}.", cancelPassed ? 1u : 0u, stopPassed ? 1u : 0u);
    }
    return cancelPassed && stopPassed;
}

bool RunFileOpsBridgePausedReaderStopSelfTestForSelfTestInternal(FolderWindow::FileOperationState& state) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;

    Task task(state);
    task.SetPaused(true);
    std::atomic<bool> externalStop{false};
    std::mutex mutex;
    std::condition_variable cv;
    bool entered  = false;
    bool returned = false;

    std::jthread reader([&]() noexcept
    {
        {
            std::scoped_lock lock(mutex);
            entered = true;
        }
        cv.notify_all();
        task.WaitWhilePaused(&externalStop);
        {
            std::scoped_lock lock(mutex);
            returned = true;
        }
        cv.notify_all();
    });

    {
        std::unique_lock lock(mutex);
        if (! cv.wait_for(lock, std::chrono::seconds(2), [&]() noexcept { return entered; }))
        {
            externalStop.store(true, std::memory_order_release);
            task.WakePauseWaiters();
            task.SetPaused(false);
            Debug::Error(L"FileOps paused-reader stop selftest did not enter the pause wait.");
            return false;
        }
    }

    const ULONGLONG stopTick = GetTickCount64();
    externalStop.store(true, std::memory_order_release);
    task.WakePauseWaiters();
    bool stoppedPromptly = false;
    {
        std::unique_lock lock(mutex);
        stoppedPromptly = cv.wait_for(lock, std::chrono::seconds(2), [&]() noexcept { return returned; });
    }
    task.SetPaused(false);
    if (! stoppedPromptly || GetTickCount64() - stopTick > 2000ull)
    {
        Debug::Error(L"FileOps paused-reader stop selftest did not unwind promptly after writer-side stop.");
        return false;
    }
    reader.join();

    // Hold a peer exactly between its false conflict predicate and condition-variable sleep.
    // Cancellation must synchronize with that mutex before notifying, or its wake is lost.
    Task conflictTask(state);
    conflictTask._conflictArbiter.prompt.active = true;
    conflictTask._conflictArbiter.ownerThreadId = GetCurrentThreadId();
    conflictTask._dbgConflictWaitBeforeSleepGate.store(true, std::memory_order_release);
    returned = false;
    std::jthread peer([&]() noexcept
    {
        conflictTask.WaitWhilePaused();
        {
            std::scoped_lock lock(mutex);
            returned = true;
        }
        cv.notify_all();
    });

    const ULONGLONG gateDeadline = GetTickCount64() + 2000ull;
    while (! conflictTask._dbgConflictWaitBeforeSleepReached.load(std::memory_order_acquire) && GetTickCount64() < gateDeadline)
    {
        Sleep(1);
    }
    const bool reachedGate = conflictTask._dbgConflictWaitBeforeSleepReached.load(std::memory_order_acquire);
    std::atomic<bool> cancelReturned{false};
    std::jthread canceller([&]() noexcept
    {
        conflictTask.RequestCancel();
        cancelReturned.store(true, std::memory_order_release);
        cv.notify_all();
    });

    // The fixed notifier blocks on the peer's mutex. The old notifier completes here,
    // before that peer can sleep, giving the RED control an explicit lost-wake ordering.
    {
        std::unique_lock lock(mutex);
        static_cast<void>(cv.wait_for(lock, std::chrono::milliseconds(100), [&]() noexcept { return cancelReturned.load(std::memory_order_acquire); }));
    }
    const bool notifiedBeforeSleep = cancelReturned.load(std::memory_order_acquire);
    conflictTask._dbgConflictWaitBeforeSleepGate.store(false, std::memory_order_release);
    conflictTask._dbgConflictWaitBeforeSleepGate.notify_all();
    bool peerReturned = false;
    {
        std::unique_lock lock(mutex);
        peerReturned = cv.wait_for(lock, std::chrono::seconds(2), [&]() noexcept { return returned; });
    }
    if (! peerReturned)
    {
        // Failure-only recovery keeps the diagnostic RED run and subsequent shutdown finite.
        {
            std::scoped_lock lock(conflictTask._conflictArbiter.mutex);
            conflictTask._conflictArbiter.prompt.active = false;
        }
        conflictTask._conflictArbiter.cv.notify_all();
    }
    peer.join();
    canceller.join();
    const bool conflictPassed = reachedGate && ! notifiedBeforeSleep && peerReturned;
    Debug::Perf::Emit(L"FileOps.SelfTest.ConflictCancelWake",
                      L"peer-held-between-predicate-and-sleep",
                      0u,
                      conflictPassed ? 1u : 0u,
                      notifiedBeforeSleep ? 1u : 0u,
                      conflictPassed ? S_OK : E_FAIL);
    if (! conflictPassed)
    {
        Debug::Error(L"FileOps conflict cancellation lost-wake selftest failed: reached={}, earlyNotify={}, returned={}.",
                     reachedGate ? 1u : 0u,
                     notifiedBeforeSleep ? 1u : 0u,
                     peerReturned ? 1u : 0u);
    }
    return conflictPassed;
}

bool RunFileOpsBridgeDirectoryBufferValidationSelfTestForSelfTestInternal() noexcept
{
    struct alignas(FileInfo) AlignedFileInfoBuffer
    {
        std::array<std::byte, sizeof(FileInfo)> bytes{};
    } buffer;

    auto* entry            = reinterpret_cast<FileInfo*>(buffer.bytes.data());
    entry->NextEntryOffset = 0;
    entry->FileNameSize    = static_cast<unsigned long>(sizeof(wchar_t) * 8u);

    std::wstring_view name;
    HRESULT hr = TryGetValidatedFileInfoName(entry, buffer.bytes.data(), buffer.bytes.data() + buffer.bytes.size(), name);
    if (hr != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
    {
        Debug::Error(L"FileOps bridge validation selftest failed: overrun FileInfo name should be rejected.");
        return false;
    }

    FileInfo* next = nullptr;
    hr             = AdvanceValidatedFileInfoEntry(entry, buffer.bytes.data(), buffer.bytes.data() + buffer.bytes.size(), next);
    if (hr != S_FALSE || next != nullptr)
    {
        Debug::Error(L"FileOps bridge validation selftest failed: terminal FileInfo entry should return S_FALSE.");
        return false;
    }

    entry->FileNameSize    = 0;
    entry->NextEntryOffset = static_cast<unsigned long>(sizeof(FileInfo) - 1u);
    hr                     = AdvanceValidatedFileInfoEntry(entry, buffer.bytes.data(), buffer.bytes.data() + buffer.bytes.size(), next);
    if (hr != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
    {
        Debug::Error(L"FileOps bridge validation selftest failed: short NextEntryOffset should be rejected.");
        return false;
    }

    return true;
}
#endif
} // namespace FolderWindowFileOperationsStateInternal

using namespace FolderWindowFileOperationsStateInternal;

#ifdef ENABLE_TESTS
HRESULT FolderWindow::FileOperationState::Task::DebugRequestDeferredConsentForSelfTest(FileOperations::DeferredConsentRisk risk,
                                                                                       bool applyToAllEligible,
                                                                                       bool itemCountKnown,
                                                                                       uint64_t itemCount,
                                                                                       bool bytesKnown,
                                                                                       uint64_t bytes,
                                                                                       ConflictAction* selectedAction) noexcept
{
    if (selectedAction == nullptr)
    {
        return E_POINTER;
    }
    *selectedAction = ConflictAction::Cancel;

    DeferredConsentRequest request{};
    request.risk                       = risk;
    request.status                     = HRESULT_FROM_WIN32(ERROR_REQUEST_ABORTED);
    request.sourcePath                 = L"selftest-source";
    request.destinationPath            = L"selftest-destination";
    request.itemIndex                  = 0u;
    request.destinationRootId          = L"selftest-root";
    request.applyToAllEligible         = applyToAllEligible;
    request.itemCountKnown             = itemCountKnown;
    request.itemCount                  = itemCount;
    request.bytesKnown                 = bytesKnown;
    request.bytes                      = bytes;
    const DeferredConsentResult result = RequestDeferredConsent(*this, request);
    *selectedAction                    = result.action;
    return result.status;
}
#endif

namespace
{
struct AtomicWriterRoute final
{
    FileSystemFlags effectiveFinalFlags  = FILESYSTEM_FLAG_NONE;
    FileSystemFlags fallbackSiblingFlags = FILESYSTEM_FLAG_NONE;
    bool useAtomicFinalWriter            = false;
};

[[nodiscard]] AtomicWriterRoute ResolveAtomicWriterRoute(IFileSystemAtomicWriter* atomicWriter,
                                                         const wchar_t* destinationPath,
                                                         FileSystemFlags taskFlags,
                                                         bool overwriteGranted,
                                                         bool replaceReadOnlyGranted) noexcept
{
    AtomicWriterRoute route{};
    uint32_t effectiveBits = static_cast<uint32_t>(taskFlags);
    if (overwriteGranted || replaceReadOnlyGranted)
    {
        effectiveBits |= static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE);
    }
    if (replaceReadOnlyGranted)
    {
        effectiveBits |= static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    }

    route.effectiveFinalFlags  = static_cast<FileSystemFlags>(effectiveBits);
    route.fallbackSiblingFlags = static_cast<FileSystemFlags>(
        effectiveBits & ~(static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY)));

    if (atomicWriter && destinationPath && destinationPath[0] != L'\0')
    {
        BOOL supported             = FALSE;
        const HRESULT capabilityHr = atomicWriter->SupportsAtomicWriterCommit(destinationPath, route.effectiveFinalFlags, &supported);
        route.useAtomicFinalWriter = SUCCEEDED(capabilityHr) && supported == TRUE;
    }

    return route;
}

// R3-1: a destination without object binding may still replace one occupant conditionally when its
// atomic-final writer accepts the overwrite flag (the writer then carries the occupant the user saw
// through IFileWriterExpectedReplacement).
[[nodiscard]] bool DestinationSupportsAtomicReplace(IFileSystem* destinationFs, std::wstring_view destinationPath) noexcept
{
    if (destinationFs == nullptr || destinationPath.empty())
    {
        return false;
    }
    wil::com_ptr<IFileSystemAtomicWriter> atomicWriter;
    if (FAILED(destinationFs->QueryInterface(IID_PPV_ARGS(atomicWriter.addressof()))) || ! atomicWriter)
    {
        return false;
    }
    try
    {
        const std::wstring path(destinationPath);
        BOOL supported = FALSE;
        return SUCCEEDED(atomicWriter->SupportsAtomicWriterCommit(path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, &supported)) && supported == TRUE;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
}
// R3-2: the digests the host is willing to compute for a writer proof, strongest first.
constexpr std::array<Common::Crypto::ContentDigestAlgorithm, 5> kWriterProofAlgorithmPreference{{
    Common::Crypto::ContentDigestAlgorithm::Sha256,
    Common::Crypto::ContentDigestAlgorithm::Sha1,
    Common::Crypto::ContentDigestAlgorithm::QuickXor,
    Common::Crypto::ContentDigestAlgorithm::Crc64Nvme,
    Common::Crypto::ContentDigestAlgorithm::Md5,
}};
} // namespace

#ifdef ENABLE_TESTS
bool ResolveFileOpsAtomicWriterRouteForSelfTest(IFileSystemAtomicWriter* atomicWriter,
                                                const wchar_t* destinationPath,
                                                FileSystemFlags taskFlags,
                                                bool overwriteGranted,
                                                bool replaceReadOnlyGranted,
                                                FileSystemFlags& effectiveFinalFlags,
                                                FileSystemFlags& fallbackSiblingFlags) noexcept
{
    const AtomicWriterRoute route = ResolveAtomicWriterRoute(atomicWriter, destinationPath, taskFlags, overwriteGranted, replaceReadOnlyGranted);
    effectiveFinalFlags           = route.effectiveFinalFlags;
    fallbackSiblingFlags          = route.fallbackSiblingFlags;
    return route.useAtomicFinalWriter;
}

bool RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept
{
    return RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTestInternal(state);
}

bool RunFileOpsPerItemSchedulerNestedSaturationSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept
{
    return RunFileOpsPerItemSchedulerNestedSaturationSelfTestForSelfTestInternal(state);
}

bool RunFileOpsPerItemSchedulerFailurePolicySelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept
{
    return RunFileOpsPerItemSchedulerFailurePolicySelfTestForSelfTestInternal(state);
}

bool RunFileOpsWorkerStartGateCancellationSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept
{
    return RunFileOpsWorkerStartGateCancellationSelfTestForSelfTestInternal(state);
}

bool RunFileOpsBridgePausedReaderStopSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept
{
    return RunFileOpsBridgePausedReaderStopSelfTestForSelfTestInternal(state);
}

bool RunFileOpsBridgeDirectoryBufferValidationSelfTestForSelfTest() noexcept
{
    return RunFileOpsBridgeDirectoryBufferValidationSelfTestForSelfTestInternal();
}

void ResetFileOpsBridgeBufferBudgetPeakForSelfTest() noexcept
{
    GetCrossFsBridgeBufferBudget().ResetPeak();
}

uint64_t GetFileOpsBridgeBufferBudgetPeakForSelfTest() noexcept
{
    return GetCrossFsBridgeBufferBudget().PeakBytes();
}

uint64_t GetFileOpsBridgeBufferBudgetInUseForSelfTest() noexcept
{
    return GetCrossFsBridgeBufferBudget().InUseBytes();
}

bool IsFileOpsCircuitBreakerTransientErrorForSelfTest(DWORD error) noexcept
{
    return IsCircuitBreakerTransientError(error);
}

bool DebugReadFileOpsConflictMetadataForSelfTest(IFileSystemIO* io, std::wstring_view path, FileOpsConflictMetadataDebugResult& out) noexcept
{
    const Task::ConflictPromptState::ItemMetadata metadata = ReadConflictItemMetadata(nullptr, io, path, false);
    out.available                                          = metadata.available;
    out.isDirectory                                        = metadata.isDirectory;
    out.attributes                                         = metadata.attributes;
    out.lastWriteTime                                      = metadata.lastWriteTime;
    return metadata.available;
}

bool DebugFileOpsUnboundCollisionWithholdsDestructiveActionsForSelfTest() noexcept
{
    const ConflictActionPolicy regular  = BuildConflictActionPolicy(ConflictBucket::RegularFileExists, false, true, true, false);
    const ConflictActionPolicy readOnly = BuildConflictActionPolicy(ConflictBucket::ReadOnlyRegularFileExists, false, true, true, false);
    const ConflictActionPolicy link     = BuildConflictActionPolicy(ConflictBucket::DestinationLink, false, true, true, false);
    return ! regular.Contains(ConflictAction::Overwrite) && regular.Contains(ConflictAction::KeepBoth) &&
           ! readOnly.Contains(ConflictAction::ReplaceReadOnly) && readOnly.Contains(ConflictAction::KeepBoth) &&
           ! link.Contains(ConflictAction::ReplaceLink) && link.Contains(ConflictAction::KeepBoth);
}

bool DebugFileOpsConflictActionPolicyCoverageForSelfTest() noexcept
{
    const auto matchesActions = [](const auto& actual, size_t actualCount, std::initializer_list<ConflictAction> expected) noexcept
    {
        return actualCount == expected.size() &&
               std::equal(actual.begin(), actual.begin() + static_cast<std::ptrdiff_t>(actualCount), expected.begin(), expected.end());
    };
    const auto matchesPolicy = [&](const ConflictActionPolicy& policy,
                                   std::initializer_list<ConflictAction> actions,
                                   std::initializer_list<ConflictAction> primary,
                                   std::initializer_list<ConflictAction> overflow,
                                   bool applyToAllEligible,
                                   bool skipAllEligible,
                                   ConflictAction expectedDefault = ConflictAction::Cancel) noexcept
    {
        return matchesActions(policy.actions, policy.actionCount, actions) && matchesActions(policy.primaryActions, policy.primaryActionCount, primary) &&
               matchesActions(policy.overflowActions, policy.overflowActionCount, overflow) && policy.defaultAction == expectedDefault &&
               policy.escapeAction == ConflictAction::Cancel && policy.applyToAllEligible == applyToAllEligible && policy.skipAllEligible == skipAllEligible &&
               policy.buttonsPublishable;
    };

    const ConflictActionPolicy regularAllowed   = BuildConflictActionPolicy(ConflictBucket::RegularFileExists, false, true, true, true);
    const ConflictActionPolicy regularWithheld  = BuildConflictActionPolicy(ConflictBucket::RegularFileExists, false, true, true, false);
    const ConflictActionPolicy readOnlyAllowed  = BuildConflictActionPolicy(ConflictBucket::ReadOnlyRegularFileExists, false, true, true, true);
    const ConflictActionPolicy readOnlyWithheld = BuildConflictActionPolicy(ConflictBucket::ReadOnlyRegularFileExists, false, true, true, false);
    const ConflictActionPolicy typeMismatch     = BuildConflictActionPolicy(ConflictBucket::TypeMismatch, false, true, true, true);
    const ConflictActionPolicy destinationLink  = BuildConflictActionPolicy(ConflictBucket::DestinationLink, false, true, true, true);
    const ConflictActionPolicy targetConflict   = BuildConflictActionPolicy(ConflictBucket::TargetConflict, false, false, false, false);
    const ConflictActionPolicy accessRetry      = BuildConflictActionPolicy(ConflictBucket::AccessDenied, true, false, true, false);
    const ConflictActionPolicy recycle          = BuildConflictActionPolicy(ConflictBucket::RecycleFailed, true, true, true, true);
    const ConflictActionPolicy efsConsent       = BuildConflictActionPolicy(ConflictBucket::EfsPlaintext, false, false, true, false);
    const ConflictActionPolicy overlapBounded   = BuildConflictActionPolicy(ConflictBucket::SameHostOverlap, false, false, false, false, true);
    const ConflictActionPolicy overlapOverflow  = BuildConflictActionPolicy(ConflictBucket::SameHostOverlap, false, false, false, false, false);
    const ConflictActionPolicy liveBounded      = BuildConflictActionPolicy(ConflictBucket::SameHostLiveOutput, false, false, false, false, false, true);
    const ConflictActionPolicy liveOverflow     = BuildConflictActionPolicy(ConflictBucket::SameHostLiveOutput, false, false, false, false, false, false);

    Task::ConflictPromptState loading{};
    loading.active          = true;
    loading.metadataLoading = true;

    return matchesPolicy(regularAllowed,
                         {ConflictAction::Overwrite, ConflictAction::KeepBoth, ConflictAction::Skip, ConflictAction::SkipAll, ConflictAction::Cancel},
                         {ConflictAction::Overwrite, ConflictAction::KeepBoth, ConflictAction::Cancel},
                         {ConflictAction::Skip, ConflictAction::SkipAll},
                         true,
                         true) &&
           matchesPolicy(regularWithheld,
                         {ConflictAction::KeepBoth, ConflictAction::Skip, ConflictAction::SkipAll, ConflictAction::Cancel},
                         {ConflictAction::KeepBoth, ConflictAction::Cancel, ConflictAction::Skip},
                         {ConflictAction::SkipAll},
                         true,
                         true) &&
           matchesPolicy(readOnlyAllowed,
                         {ConflictAction::ReplaceReadOnly, ConflictAction::KeepBoth, ConflictAction::Skip, ConflictAction::SkipAll, ConflictAction::Cancel},
                         {ConflictAction::ReplaceReadOnly, ConflictAction::KeepBoth, ConflictAction::Cancel},
                         {ConflictAction::Skip, ConflictAction::SkipAll},
                         true,
                         true) &&
           matchesPolicy(readOnlyWithheld,
                         {ConflictAction::KeepBoth, ConflictAction::Skip, ConflictAction::SkipAll, ConflictAction::Cancel},
                         {ConflictAction::KeepBoth, ConflictAction::Cancel, ConflictAction::Skip},
                         {ConflictAction::SkipAll},
                         true,
                         true) &&
           matchesPolicy(typeMismatch,
                         {ConflictAction::KeepBoth, ConflictAction::Skip, ConflictAction::SkipAll, ConflictAction::Cancel},
                         {ConflictAction::KeepBoth, ConflictAction::Cancel, ConflictAction::Skip},
                         {ConflictAction::SkipAll},
                         true,
                         true) &&
           matchesPolicy(destinationLink,
                         {ConflictAction::ReplaceLink, ConflictAction::KeepBoth, ConflictAction::Skip, ConflictAction::SkipAll, ConflictAction::Cancel},
                         {ConflictAction::ReplaceLink, ConflictAction::KeepBoth, ConflictAction::Cancel},
                         {ConflictAction::Skip, ConflictAction::SkipAll},
                         true,
                         true) &&
           matchesPolicy(targetConflict, {ConflictAction::Skip, ConflictAction::Cancel}, {ConflictAction::Cancel, ConflictAction::Skip}, {}, false, false) &&
           matchesPolicy(accessRetry,
                         {ConflictAction::Retry, ConflictAction::Skip, ConflictAction::SkipAll, ConflictAction::Cancel},
                         {ConflictAction::Retry, ConflictAction::Cancel, ConflictAction::Skip},
                         {ConflictAction::SkipAll},
                         true,
                         true) &&
           matchesPolicy(recycle,
                         {ConflictAction::Cancel, ConflictAction::PermanentDelete, ConflictAction::Skip},
                         {ConflictAction::Cancel, ConflictAction::PermanentDelete, ConflictAction::Skip},
                         {},
                         false,
                         false) &&
           matchesPolicy(efsConsent,
                         {ConflictAction::Proceed, ConflictAction::RetainSource, ConflictAction::Cancel},
                         {ConflictAction::Proceed, ConflictAction::RetainSource, ConflictAction::Cancel},
                         {},
                         true,
                         false) &&
           matchesPolicy(overlapBounded,
                         {ConflictAction::Proceed, ConflictAction::RunConcurrently, ConflictAction::Cancel},
                         {ConflictAction::Proceed, ConflictAction::RunConcurrently, ConflictAction::Cancel},
                         {},
                         false,
                         false,
                         ConflictAction::Proceed) &&
           matchesPolicy(overlapOverflow,
                         {ConflictAction::Proceed, ConflictAction::Cancel},
                         {ConflictAction::Proceed, ConflictAction::Cancel},
                         {},
                         false,
                         false,
                         ConflictAction::Proceed) &&
           matchesPolicy(liveBounded,
                         {ConflictAction::QueueUntilOtherTaskFinishes, ConflictAction::Skip, ConflictAction::InvalidateLiveOutput, ConflictAction::Cancel},
                         {ConflictAction::QueueUntilOtherTaskFinishes, ConflictAction::InvalidateLiveOutput, ConflictAction::Cancel},
                         {ConflictAction::Skip},
                         false,
                         false,
                         ConflictAction::QueueUntilOtherTaskFinishes) &&
           matchesPolicy(liveOverflow,
                         {ConflictAction::QueueUntilOtherTaskFinishes, ConflictAction::Skip, ConflictAction::Cancel},
                         {ConflictAction::QueueUntilOtherTaskFinishes, ConflictAction::Cancel, ConflictAction::Skip},
                         {},
                         false,
                         false,
                         ConflictAction::QueueUntilOtherTaskFinishes) &&
           loading.actionCount == 0u && loading.primaryActionCount == 0u && loading.overflowActionCount == 0u &&
           loading.defaultAction == ConflictAction::None && loading.escapeAction == ConflictAction::None && ! loading.applyToAllEligible &&
           ! loading.skipAllEligible && ! loading.buttonsPublishable;
}
#endif

FolderWindow::FileOperationState::Task::Task(FileOperationState& state) noexcept : _state(&state), _folderWindow(&state._owner)
{
    _conflictArbiter.decisionEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
}

std::shared_ptr<const FileOperations::FileOperationPlanGroup> FolderWindow::FileOperationState::Task::LoadPlans() const noexcept
{
    return _plans.load(std::memory_order_acquire);
}

void FolderWindow::FileOperationState::Task::StorePlans(std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans) noexcept
{
    _plans.store(std::move(plans), std::memory_order_release);
}

FolderWindow::FileOperationState::Task::TaskLifecyclePhase FolderWindow::FileOperationState::Task::GetLifecyclePhase() const noexcept
{
    return _lifecyclePhase.load(std::memory_order_acquire);
}

std::shared_ptr<const FolderWindow::FileOperationState::Task::PreparationSnapshot> FolderWindow::FileOperationState::Task::LoadPreparationSnapshot()
    const noexcept
{
    return _preparationSnapshot.load(std::memory_order_acquire);
}

void FolderWindow::FileOperationState::Task::PublishLifecyclePhase(const TaskLifecyclePhase phase) noexcept
{
    TaskLifecyclePhase current = _lifecyclePhase.load(std::memory_order_acquire);
    while (current < phase && ! _lifecyclePhase.compare_exchange_weak(current, phase, std::memory_order_acq_rel, std::memory_order_acquire))
    {
    }
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemProgress(FileSystemOperation operationType,
                                                                                     unsigned long totalItems,
                                                                                     unsigned long completedItems,
                                                                                     uint64_t totalBytes,
                                                                                     uint64_t completedBytes,
                                                                                     const wchar_t* currentSourcePath,
                                                                                     const wchar_t* currentDestinationPath,
                                                                                     uint64_t currentItemTotalBytes,
                                                                                     uint64_t currentItemCompletedBytes,
                                                                                     FileSystemOptions* options,
                                                                                     uint64_t progressStreamId,
                                                                                     void* cookie) noexcept
{
#ifdef ENABLE_TESTS
    assert(_dbgCallbackActiveScopeCount.load(std::memory_order_acquire) > 0u);
#endif

    if (operationType != _operation)
    {
        return S_OK;
    }

    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }

    const ULONGLONG nowTick                   = GetTickCount64();
    const uint64_t perfStartUs                = PerfNowUs();
    const bool discoveryOpenAtCallback        = ! _discoveryClosed.load(std::memory_order_acquire);
    bool trackProgressStreamPerf              = false;
    uint64_t progressLockWaitUs               = 0;
    uint64_t progressStreamItemCompletedBytes = 0;
    const auto perfCallbackScope              = wil::scope_exit([&] noexcept
    {
        const uint64_t callbackUs = PerfElapsedUs(perfStartUs);
        _perf.progressCallbackUs.fetch_add(callbackUs, std::memory_order_relaxed);
        if (trackProgressStreamPerf)
        {
            NoteProgressStreamPerf(*this, cookie, progressStreamId, nowTick, progressStreamItemCompletedBytes, progressLockWaitUs, callbackUs);
        }
    });

#ifdef ENABLE_TESTS
    bool warnSingleInFlightProgress         = false;
    unsigned int dbgConfiguredConcurrency   = 1u;
    size_t dbgInFlightFileCount             = 0;
    unsigned long dbgPlannedTopLevelFiles   = 0;
    unsigned long dbgPlannedTopLevelFolders = 0;
    bool warnPerItemInFlightEviction        = false;
    const void* dbgPerItemEvictedCookie     = nullptr;
    size_t dbgPerItemCapacity               = 0;
    size_t dbgPerItemInFlightCount          = 0;
#endif

    PerItemInFlightUpdateResult perItemInFlightUpdate{};
    PerItemInFlightAggregate perItemInFlightAggregate{};
    unsigned int perItemBandwidthActiveCalls = 1u;
    if (_executionMode == ExecutionMode::PerItem)
    {
        if (cookie != nullptr)
        {
            perItemInFlightUpdate    = UpdatePerItemInFlightCall(*this, cookie, completedItems, completedBytes, totalItems, nowTick);
            perItemInFlightAggregate = perItemInFlightUpdate.aggregate;
        }
        else
        {
            perItemInFlightAggregate = GetPerItemInFlightAggregate(*this);
        }

        perItemBandwidthActiveCalls = std::max(1u, static_cast<unsigned int>(perItemInFlightAggregate.activeCount));
    }

    const uint64_t previousProgressCallbackCount = _progressCallbackCount.fetch_add(1, std::memory_order_relaxed);
    if (previousProgressCallbackCount == 0)
    {
        const ULONGLONG opStartTick = _operationStartTick.load(std::memory_order_acquire);
        if (opStartTick != 0 && nowTick >= opStartTick)
        {
            _perf.progressFirstCallbackDelayMs = static_cast<uint64_t>(nowTick - opStartTick);
        }
    }
    PublishedProgressSnapshot publishedProgressSnapshot{};

    {
        const uint64_t progressLockWaitStartUs = PerfNowUs();
        std::scoped_lock lock(_progressMutex);
        progressLockWaitUs = PerfElapsedUs(progressLockWaitStartUs);
        _perf.progressLockWaitUs += progressLockWaitUs;
        if (progressLockWaitUs > 0)
        {
            ++_perf.progressLockContentionCount;
        }
        const uint64_t progressLockHoldStartUs   = PerfNowUs();
        const auto progressLockHoldScope         = wil::scope_exit([&] noexcept { _perf.progressLockHoldUs += PerfElapsedUs(progressLockHoldStartUs); });
        trackProgressStreamPerf                  = true;
        const uint64_t completedBytesBefore      = _progressCompletedBytes;
        const unsigned long completedItemsBefore = _progressCompletedItems;
        // A transfer callback while traversal is open proves the one-pass scheduler admitted a
        // safe mutation without waiting for a whole-tree total.
        if (! _firstMutationBeforeDiscoveryClosed.load(std::memory_order_acquire) && ! _discoveryClosed.load(std::memory_order_acquire))
        {
            _firstMutationBeforeDiscoveryClosed.store(true, std::memory_order_release);
        }
        if (_executionMode == ExecutionMode::PerItem)
        {
            if (_perItemTotalItems > 0 && _operation != FILESYSTEM_DELETE)
            {
                _progressTotalItems = (std::max)(_progressTotalItems, _perItemTotalItems);
            }

            if (perItemInFlightUpdate.evicted)
            {
                ++_perf.perItemInFlightEvictions;
#ifdef ENABLE_TESTS
                if (_dbgLastPerItemInFlightEvictWarnTick == 0 ||
                    (nowTick >= _dbgLastPerItemInFlightEvictWarnTick && (nowTick - _dbgLastPerItemInFlightEvictWarnTick) > 5'000ull))
                {
                    _dbgLastPerItemInFlightEvictWarnTick = nowTick;
                    warnPerItemInFlightEviction          = true;
                    dbgPerItemEvictedCookie              = perItemInFlightUpdate.evictedCookie;
                    dbgPerItemCapacity                   = _perItemInFlightCalls.size();
                    dbgPerItemInFlightCount              = perItemInFlightAggregate.activeCount;
                }
#endif
            }

            const uint64_t mappedCompletedBytes = _perItemCompletedBytes + perItemInFlightAggregate.completedBytes;
            _progressCompletedBytes             = (std::max)(_progressCompletedBytes, mappedCompletedBytes);

            if (_operation == FILESYSTEM_DELETE)
            {
                const uint64_t mappedTotalItems = _perItemTotalEntryCount + perItemInFlightAggregate.totalItems;
                if (mappedTotalItems > 0)
                {
                    const uint64_t clamped = std::min<uint64_t>(mappedTotalItems, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()));
                    _progressTotalItems    = (std::max)(_progressTotalItems, static_cast<unsigned long>(clamped));
                }

                const uint64_t mappedCompletedItems = _perItemCompletedEntryCount + perItemInFlightAggregate.completedItems;
                const uint64_t clamped  = std::min<uint64_t>(mappedCompletedItems, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()));
                _progressCompletedItems = (std::max)(_progressCompletedItems, static_cast<unsigned long>(clamped));
            }
            else
            {
                _progressCompletedItems = (std::max)(_progressCompletedItems, _perItemCompletedItems);
            }
        }
        else
        {
            if (totalItems > 0)
            {
                _progressTotalItems = (std::max)(_progressTotalItems, totalItems);
            }
            _progressCompletedItems = (std::max)(_progressCompletedItems, completedItems);
            if (totalBytes > 0)
            {
                _progressTotalBytes = (std::max)(_progressTotalBytes, totalBytes);
            }
            _progressCompletedBytes = (std::max)(_progressCompletedBytes, completedBytes);
        }

        if (discoveryOpenAtCallback)
        {
            const uint64_t completedBytesDelta = _progressCompletedBytes >= completedBytesBefore ? _progressCompletedBytes - completedBytesBefore : 0u;
            const uint64_t completedItemsDelta =
                _progressCompletedItems >= completedItemsBefore ? static_cast<uint64_t>(_progressCompletedItems - completedItemsBefore) : 0u;
            NoteDiscoveryCompletionWhileOpen(perfStartUs, completedBytesDelta, currentItemTotalBytes == 0u ? completedItemsDelta : 0u);
        }

        _progressItemTotalBytes          = currentItemTotalBytes;
        _progressItemCompletedBytes      = currentItemCompletedBytes;
        publishedProgressSnapshot        = CapturePublishedProgressSnapshotLocked(*this);
        progressStreamItemCompletedBytes = currentItemCompletedBytes;

#ifdef ENABLE_TESTS
        dbgConfiguredConcurrency  = _dbgConfiguredMaxConcurrency;
        dbgPlannedTopLevelFiles   = _plannedTopLevelFiles;
        dbgPlannedTopLevelFolders = _plannedTopLevelFolders;
#endif
    }

    StorePublishedProgressSnapshot(*this, publishedProgressSnapshot);

    UpdateProgressPathState(*this,
                            (_executionMode == ExecutionMode::PerItem && cookie != nullptr) ? static_cast<PerItemCallbackCookie*>(cookie) : nullptr,
                            currentSourcePath,
                            currentDestinationPath,
                            nowTick);

    ApplyCallbackBandwidthLimit(*this, options, perItemBandwidthActiveCalls);

    if (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE)
    {
        UpdateInFlightFileProgress(
            *this, cookie, progressStreamId, currentSourcePath, currentItemTotalBytes, currentItemCompletedBytes, nowTick, discoveryOpenAtCallback);
    }

#ifdef ENABLE_TESTS
    dbgInFlightFileCount = GetInFlightFileCountSnapshot(*this);
    if ((_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE) && dbgConfiguredConcurrency > 1u)
    {
        const unsigned long plannedTopLevelItems = dbgPlannedTopLevelFiles + dbgPlannedTopLevelFolders;
        const bool likelyParallelWork            = dbgPlannedTopLevelFolders > 0 || plannedTopLevelItems > 1u;

        if (likelyParallelWork)
        {
            if (dbgInFlightFileCount > 1u)
            {
                _dbgObservedMultipleInFlightFiles = true;
                _dbgSingleInFlightStartTick       = 0;
            }
            else if (! _dbgObservedMultipleInFlightFiles)
            {
                if (_dbgSingleInFlightStartTick == 0)
                {
                    _dbgSingleInFlightStartTick = nowTick;
                }
                else if (_dbgLastSingleInFlightWarnTick == 0 && nowTick >= _dbgSingleInFlightStartTick && (nowTick - _dbgSingleInFlightStartTick) > 15'000ull)
                {
                    _dbgLastSingleInFlightWarnTick = nowTick;
                    warnSingleInFlightProgress     = true;
                }
            }
        }
        else
        {
            _dbgSingleInFlightStartTick = 0;
        }
    }
    else
    {
        _dbgSingleInFlightStartTick = 0;
    }

    if (warnSingleInFlightProgress)
    {
        Debug::Warning(
            L"FileOps: expected multiple in-flight file progress lines but observed <= 1 for >15s (taskId={} op={} execMode={} configuredConcurrency={} "
            L"plannedFiles={} plannedFolders={} inFlightFiles={} cookie={:p} streamId={}).",
            _taskId,
            static_cast<unsigned int>(_operation),
            static_cast<unsigned int>(_executionMode),
            dbgConfiguredConcurrency,
            dbgPlannedTopLevelFiles,
            dbgPlannedTopLevelFolders,
            dbgInFlightFileCount,
            cookie,
            progressStreamId);
    }

    if (warnPerItemInFlightEviction)
    {
        Debug::Warning(L"FileOps: per-item in-flight call table overflow; evicted oldest entry (taskId={} op={} execMode={} perItemTableSize={} "
                       L"perItemInFlightCount={} newCookie={:p} evictedCookie={:p}).",
                       _taskId,
                       static_cast<unsigned int>(_operation),
                       static_cast<unsigned int>(_executionMode),
                       dbgPerItemCapacity,
                       dbgPerItemInFlightCount,
                       cookie,
                       dbgPerItemEvictedCookie);
    }
#endif

    WaitWhilePaused();

    if (_cancelled.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return S_OK;
}

HRESULT FolderWindow::FileOperationState::Task::ReportVerificationProgress(const wchar_t* sourcePath,
                                                                           const wchar_t* destinationPath,
                                                                           uint64_t itemTotalBytes,
                                                                           uint64_t itemCompletedBytes,
                                                                           uint64_t completedBytes,
                                                                           bool active) noexcept
{
    _verificationActive.store(active, std::memory_order_release);
    _verificationItemTotalBytes.store(itemTotalBytes, std::memory_order_release);
    _verificationItemCompletedBytes.store(itemCompletedBytes, std::memory_order_release);
    _verificationCompletedBytes.store(completedBytes, std::memory_order_release);
    if (active)
    {
        UpdateProgressPathState(*this, nullptr, sourcePath, destinationPath, GetTickCount64());
    }
    WaitWhilePaused();
    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemItemCompleted(FileSystemOperation operationType,
                                                                                          unsigned long itemIndex,
                                                                                          const wchar_t* sourcePath,
                                                                                          const wchar_t* destinationPath,
                                                                                          HRESULT status,
                                                                                          const FileSystemItemMutationResult* mutationResult,
                                                                                          FileSystemOptions* options,
                                                                                          void* cookie) noexcept
{
#ifdef ENABLE_TESTS
    assert(_dbgCallbackActiveScopeCount.load(std::memory_order_acquire) > 0u);
#endif

    if (operationType != _operation)
    {
        return S_OK;
    }

    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    auto mutationSnapshot        = FileSystemRouteContract::SnapshotItemMutationResult(mutationResult);
    const bool malformedMutation = mutationResult != nullptr && ! mutationSnapshot.has_value();
    if (malformedMutation)
    {
        // An invalid completion replaces any preceding child receipt with unknown
        // truth, even when the provider ignores our E_INVALIDARG return. Keeping a
        // present unknown receipt also prevents the intentional-partial exception
        // from treating this malformed completion as an ordinary receipt-less Skip.
        mutationSnapshot = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, FALSE};
    }
    mutationResult = mutationSnapshot.has_value() ? &mutationSnapshot.value() : nullptr;

    const uint64_t perfStartUs                = PerfNowUs();
    const auto perfCallbackScope              = wil::scope_exit([&] noexcept { _perf.itemCompletedCallbackUs += PerfElapsedUs(perfStartUs); });
    unsigned int perItemBandwidthActiveCalls  = 1u;
    const uint64_t itemCompletedCallbackCount = _itemCompletedCallbackCount.fetch_add(1, std::memory_order_relaxed) + 1u;
    PublishedProgressSnapshot publishedProgressSnapshot{};

    PerItemCallbackCookie* const perItemCookie =
        (_executionMode == ExecutionMode::PerItem && cookie != nullptr) ? static_cast<PerItemCallbackCookie*>(cookie) : nullptr;
    size_t sourceIndex = static_cast<size_t>(itemIndex);
    if (perItemCookie != nullptr)
    {
        sourceIndex = perItemCookie->itemIndex;
    }
    bool logRetainedCleanupDebt = false;
    if (sourceIndex < _sourcePaths.size())
    {
        std::scoped_lock lock(_sourceItemStatusMutex);
        if (_sourceItemResultBuilders.size() < _sourcePaths.size())
        {
            _sourceItemResultBuilders.resize(_sourcePaths.size());
        }
        SourceItemResultBuilder& builder = _sourceItemResultBuilders[sourceIndex];
        builder.status                   = status;
        if (mutationResult != nullptr)
        {
            logRetainedCleanupDebt = ! builder.mutation.has_value() && mutationResult->outcomeKnown == TRUE && mutationResult->mutationCommitted == TRUE &&
                                     FileSystemItemMutationResultHasOwnedStageDisposition(*mutationResult) &&
                                     mutationResult->ownedStageDisposition == FileSystemOwnedStageDisposition::Retained;
            builder.mutation       = mutationSnapshot;
        }
        else
        {
            // A later unknown completion must not inherit an earlier child/attempt's receipt.
            builder.mutation.reset();
        }
    }

    if (malformedMutation)
    {
        return E_INVALIDARG;
    }

    if (_executionMode != ExecutionMode::PerItem)
    {
        StorePublishedTopLevelCompletionSnapshot(*this, MarkTopLevelItemCompleted(*this, static_cast<size_t>(itemIndex)));
    }

    {
        const uint64_t progressLockWaitStartUs = PerfNowUs();
        std::scoped_lock lock(_progressMutex);
        const uint64_t itemCompletedLockWaitUs = PerfElapsedUs(progressLockWaitStartUs);
        _perf.itemCompletedLockWaitUs += itemCompletedLockWaitUs;
        if (itemCompletedLockWaitUs > 0)
        {
            ++_perf.itemCompletedLockContentionCount;
        }
        const uint64_t progressLockHoldStartUs = PerfNowUs();
        const auto progressLockHoldScope       = wil::scope_exit([&] noexcept { _perf.itemCompletedLockHoldUs += PerfElapsedUs(progressLockHoldStartUs); });
        if (_executionMode != ExecutionMode::PerItem)
        {
            const unsigned long completedItemsClamped = static_cast<unsigned long>(std::min(itemCompletedCallbackCount, static_cast<uint64_t>(ULONG_MAX)));
            _progressCompletedItems                   = (std::max)(_progressCompletedItems, completedItemsClamped);
        }
        _lastItemIndex            = itemIndex;
        _lastItemHr               = status;
        publishedProgressSnapshot = CapturePublishedProgressSnapshotLocked(*this);
    }

    StorePublishedProgressSnapshot(*this, publishedProgressSnapshot);

    if (_executionMode == ExecutionMode::PerItem)
    {
        perItemBandwidthActiveCalls = std::max(1u, static_cast<unsigned int>(GetPerItemInFlightCallCountSnapshot(*this)));
    }

    UpdateItemCompletedPathState(*this, perItemCookie, sourcePath, destinationPath);

    if (destinationPath != nullptr && destinationPath[0] != L'\0' && status != S_FALSE && SUCCEEDED(status) &&
        (mutationResult == nullptr || mutationResult->mutationCommitted != FALSE))
    {
        NoteLiveOutputPublished(destinationPath);
    }

    if (logRetainedCleanupDebt)
    {
        LogDiagnostic(DiagnosticSeverity::Warning,
                      S_FALSE,
                      L"item.cleanup.retained",
                      LoadStringResource(nullptr, IDS_FILEOPS_CLEANUP_ITEM_RETAINED),
                      sourcePath != nullptr ? std::wstring_view(sourcePath) : std::wstring_view{},
                      destinationPath != nullptr ? std::wstring_view(destinationPath) : std::wstring_view{});
    }

    ApplyCallbackBandwidthLimit(*this, options, perItemBandwidthActiveCalls);

    RemoveInFlightFileBySourcePath(*this, sourcePath);

    if (status == S_FALSE)
    {
        _observedSkipAction.store(true, std::memory_order_release);
        if (perItemCookie != nullptr)
        {
            perItemCookie->explicitSkipObserved.store(true, std::memory_order_release);
        }
        LogDiagnostic(DiagnosticSeverity::Info,
                      S_FALSE,
                      L"item.completed.skipped",
                      L"Provider completed the item without rewriting the destination.",
                      sourcePath != nullptr ? std::wstring_view(sourcePath) : std::wstring_view{},
                      destinationPath != nullptr ? std::wstring_view(destinationPath) : std::wstring_view{});
    }

    if (_cancelled.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemShouldCancel(BOOL* pCancel, void* /*cookie*/) noexcept
{
#ifdef ENABLE_TESTS
    assert(_dbgCallbackActiveScopeCount.load(std::memory_order_acquire) > 0u);
#endif

    if (! pCancel)
    {
        return E_POINTER;
    }

    const bool cancel = _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested();
    *pCancel          = cancel ? TRUE : FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemIssue(FileSystemOperation operationType,
                                                                                  const wchar_t* sourcePath,
                                                                                  const wchar_t* destinationPath,
                                                                                  HRESULT status,
                                                                                  FileSystemIssueAction* action,
                                                                                  IFileSystemBoundObject** expectedDestination,
                                                                                  [[maybe_unused]] FileSystemOptions* options,
                                                                                  void* cookie) noexcept
{
#ifdef ENABLE_TESTS
    assert(_dbgCallbackActiveScopeCount.load(std::memory_order_acquire) > 0u);
#endif

    if (! action || ! expectedDestination)
    {
        return E_POINTER;
    }

    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }

    *action              = FileSystemIssueAction::Cancel;
    *expectedDestination = nullptr;

    if (IsTraversalResourceLimitStatus(status))
    {
        // Traversal ceilings are quantitative task stop conditions, never user-resolvable
        // conflicts. Preserve the exact status and do not strand the worker behind a prompt.
        return status;
    }

    WaitWhilePaused();

    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    const std::wstring_view sourceText      = sourcePath ? sourcePath : L"";
    const std::wstring_view destinationText = destinationPath ? destinationPath : L"";

    wil::com_ptr<IFileSystemBoundObject> exactDestination;
    HRESULT exactDestinationHr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    bool atomicReplaceEligible = false; // R3-1: no bound authority, but a conditional atomic-final replace exists
    if (! destinationText.empty())
    {
        const wil::com_ptr<IFileSystem>& destinationFileSystem = _destinationFileSystem ? _destinationFileSystem : _fileSystem;
        wil::com_ptr<IFileSystemObjectBinding> destinationBinding;
        exactDestinationHr =
            destinationFileSystem ? destinationFileSystem->QueryInterface(__uuidof(IFileSystemObjectBinding), destinationBinding.put_void()) : E_NOINTERFACE;
        if (SUCCEEDED(exactDestinationHr) && destinationBinding)
        {
            constexpr FileSystemBindFlags conflictBindFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA |
                                                                                               FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION);
            exactDestinationHr                              = destinationBinding->BindObject(destinationText.data(), conflictBindFlags, exactDestination.put());
            if (SUCCEEDED(exactDestinationHr) && ! exactDestination)
            {
                exactDestinationHr = E_UNEXPECTED;
            }
        }
        if (SUCCEEDED(exactDestinationHr) && exactDestination)
        {
            _conflictExpectedDestinationBoundCount.fetch_add(1u, std::memory_order_relaxed);
        }
        else if (! destinationBinding && destinationFileSystem)
        {
            atomicReplaceEligible = DestinationSupportsAtomicReplace(destinationFileSystem.get(), destinationText);
        }
    }

    PerItemCallbackCookie* perItemCookie = nullptr;
    if (_executionMode == ExecutionMode::PerItem && cookie != nullptr)
    {
        perItemCookie = static_cast<PerItemCallbackCookie*>(cookie);
    }

    ConflictBucket bucket = ClassifyConflictBucket(operationType, _flags, wil::com_ptr<IFileSystemIO>{}, status, sourceText, destinationText, false);
#ifdef ENABLE_TESTS
    if (bucket == ConflictBucket::RegularFileExists && perItemCookie != nullptr && ! perItemCookie->operationDestinationPath.empty() &&
        ! NavigationLocation::EqualsNoCase(destinationText, perItemCookie->operationDestinationPath))
    {
        g_fileOpsKeepBothNestedConflictPausePoint.Pause(10'000ull);
    }
#endif
    if (bucket == ConflictBucket::RecycleFailed)
    {
        auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, perItemCookie, sourceText, destinationText);
        LogDiagnostic(
            DiagnosticSeverity::Error, status, L"delete.recycleBin.item", L"Recycle Bin delete failed for item.", diagnosticSource, diagnosticDestination);
        // A provider issue callback cannot carry the initial per-item mutation truth or the retained
        // source authority required by the Recycle escalation contract. The executor owns that gate
        // after the initial Recycle call completes; never hand a pathname-based PermanentDelete action
        // back into the provider.
        *action = FileSystemIssueAction::Cancel;
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const size_t bucketIndex = static_cast<size_t>(bucket);

    // C2: a retryable bucket stays retryable; each explicit Retry is one attempt, and the prompt
    // says how many have failed so far.
    const bool allowRetry     = IsRetryableConflictBucket(bucket);
    unsigned int attemptCount = 0u;
    if (perItemCookie != nullptr && bucketIndex < perItemCookie->issueRetryCounts.size())
    {
        attemptCount = perItemCookie->issueRetryCounts[bucketIndex];
    }

    // Replace link stays bound-only: a link needs exact identity. Files on an identity-less route
    // are replaced conditionally through the writer contract instead.
    const bool allowDestructiveReplacement = static_cast<bool>(exactDestination) || (atomicReplaceEligible && bucket != ConflictBucket::DestinationLink);
    const ConflictPromptBeginResult promptBegin =
        BeginConflictPrompt(*this, perItemCookie, bucket, status, sourceText, destinationText, allowRetry, attemptCount, false, allowDestructiveReplacement);
    bucket                           = promptBegin.bucket;
    const size_t resolvedBucketIndex = static_cast<size_t>(bucket);
    ConflictAction decision          = promptBegin.action;
    if (promptBegin.ownsPrompt)
    {
        const auto result = WaitForConflictDecision(*this, cookie, promptBegin.decisionScope);
        decision          = result.first;
    }

    const auto returnDestructiveAction = [&](FileSystemIssueAction selected) noexcept -> HRESULT
    {
        switch (GuardLiveOutputBeforeInvalidation(destinationText, FileOperations::MutationInterlockAccess::PublishDestination))
        {
            case LiveOutputGuardDisposition::Skip: *action = FileSystemIssueAction::Skip; return S_OK;
            case LiveOutputGuardDisposition::RetryCurrentMutation: *action = FileSystemIssueAction::Retry; return S_OK;
            case LiveOutputGuardDisposition::Cancel: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            case LiveOutputGuardDisposition::Proceed: break;
        }
        if (FAILED(exactDestinationHr) || ! exactDestination)
        {
            if (atomicReplaceEligible && selected != FileSystemIssueAction::ReplaceLink)
            {
                // R3-1: the decision is carried to the atomic-final writer as the occupant the user saw.
                *action              = selected;
                *expectedDestination = nullptr;
                _conflictAtomicReplaceGrantCount.fetch_add(1u, std::memory_order_relaxed);
                return S_OK;
            }
            _conflictExpectedDestinationUnavailableCount.fetch_add(1u, std::memory_order_relaxed);
            return FAILED(exactDestinationHr) ? exactDestinationHr : HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        *action              = selected;
        *expectedDestination = exactDestination.detach();
        _conflictExpectedDestinationReturnedCount.fetch_add(1u, std::memory_order_relaxed);
        return S_OK;
    };

    switch (decision)
    {
        case ConflictAction::Overwrite: return returnDestructiveAction(FileSystemIssueAction::Overwrite);
        case ConflictAction::ReplaceReadOnly: return returnDestructiveAction(FileSystemIssueAction::ReplaceReadOnly);
        case ConflictAction::ReplaceLink: return returnDestructiveAction(FileSystemIssueAction::ReplaceLink);
        case ConflictAction::PermanentDelete: *action = FileSystemIssueAction::Cancel; return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        case ConflictAction::KeepBoth:
            if (NavigationLocation::EqualsNoCase(_sourcePluginId, L"builtin/file-system") &&
                NavigationLocation::EqualsNoCase(_destinationPluginId, L"builtin/file-system"))
            {
                *action = FileSystemIssueAction::KeepBoth;
                return S_OK;
            }
            if (_executionMode == ExecutionMode::PerItem && perItemCookie != nullptr && ! perItemCookie->operationDestinationPath.empty() &&
                NavigationLocation::EqualsNoCase(destinationText, perItemCookie->operationDestinationPath))
            {
                perItemCookie->keepBothRequested = true;
                *action                          = FileSystemIssueAction::Skip;
                return S_OK;
            }
            *action = FileSystemIssueAction::Cancel;
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        case ConflictAction::Retry:
            if (perItemCookie != nullptr && resolvedBucketIndex < perItemCookie->issueRetryCounts.size())
            {
                ++perItemCookie->issueRetryCounts[resolvedBucketIndex];
            }
            *action = FileSystemIssueAction::Retry;
            return S_OK;
        case ConflictAction::Skip:
        case ConflictAction::SkipAll:
        {
            auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, perItemCookie, sourceText, destinationText);
            LogDiagnostic(
                DiagnosticSeverity::Warning, status, L"item.conflict.skip", L"Conflict action Skip item selected.", diagnosticSource, diagnosticDestination);
            _observedSkipAction.store(true, std::memory_order_release);
            if (perItemCookie != nullptr)
            {
                perItemCookie->explicitSkipObserved.store(true, std::memory_order_release);
            }
            *action = FileSystemIssueAction::Skip;
            return S_OK;
        }
        case ConflictAction::Cancel:
        case ConflictAction::Proceed:
        case ConflictAction::RetainSource:
        case ConflictAction::RunConcurrently:
        case ConflictAction::QueueUntilOtherTaskFinishes:
        case ConflictAction::InvalidateLiveOutput:
        case ConflictAction::None:
        default: *action = FileSystemIssueAction::Cancel; return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
}

void FolderWindow::FileOperationState::Task::BeginDiscovery() noexcept
{
    if (_operation != FILESYSTEM_COPY && _operation != FILESYSTEM_MOVE && _operation != FILESYSTEM_DELETE)
    {
        _discoveryClosed.store(true, std::memory_order_release);
        return;
    }

    {
        std::scoped_lock lock(_discoveryMutex);
        _discoveryItemClosed.assign(_sourcePaths.size(), 0u);
        _closedDiscoveryItemCount = 0u;
    }
    const ULONGLONG nowTick = GetTickCount64();
    _discoverySkipped.store(false, std::memory_order_release);
    _discoveryClosed.store(_sourcePaths.empty(), std::memory_order_release);
    _discoveryAheadActive.store(! _sourcePaths.empty(), std::memory_order_release);
    _discoveryStartTick.store(nowTick, std::memory_order_release);
    _discoveryStartPerfUs.store(PerfNowUs(), std::memory_order_release);
    _discoveryClosedTick.store(_sourcePaths.empty() ? nowTick : 0u, std::memory_order_release);
    _discoverySkipRequestedTick.store(0u, std::memory_order_release);
    _discoveryReservationReleasedTick.store(0u, std::memory_order_release);
    _discoveryFirstMutationUs.store(0u, std::memory_order_release);
    _discoveryCompletedBytesWhileOpen.store(0u, std::memory_order_release);
    _discoveryCompletedMutationsWhileOpen.store(0u, std::memory_order_release);
    _discoveredTotalBytes.store(0u, std::memory_order_release);
    _discoveredFileCount.store(0u, std::memory_order_release);
    _discoveredDirectoryCount.store(0u, std::memory_order_release);
    _discoveryMaxQueueDepth.store(0u, std::memory_order_release);
    _discoveryStarvationCount.store(0u, std::memory_order_release);
    _firstMutationBeforeDiscoveryClosed.store(false, std::memory_order_release);
}

void FolderWindow::FileOperationState::Task::NoteDiscoveryCompletionWhileOpen(const uint64_t observedAtPerfUs,
                                                                              const uint64_t completedBytes,
                                                                              const uint64_t completedMutations) noexcept
{
    if ((completedBytes == 0u && completedMutations == 0u) || _discoveryClosed.load(std::memory_order_acquire))
    {
        return;
    }

    uint64_t noFirstMutation            = 0u;
    const uint64_t discoveryStartPerfUs = _discoveryStartPerfUs.load(std::memory_order_acquire);
    const uint64_t firstMutationUs =
        (std::max)(uint64_t{1u}, observedAtPerfUs >= discoveryStartPerfUs ? observedAtPerfUs - discoveryStartPerfUs : uint64_t{0u});
    static_cast<void>(
        _discoveryFirstMutationUs.compare_exchange_strong(noFirstMutation, firstMutationUs, std::memory_order_acq_rel, std::memory_order_acquire));
    _firstMutationBeforeDiscoveryClosed.store(true, std::memory_order_release);
    SaturatingAtomicAdd(_discoveryCompletedBytesWhileOpen, completedBytes);
    SaturatingAtomicAdd(_discoveryCompletedMutationsWhileOpen, completedMutations);
}

void FolderWindow::FileOperationState::Task::SkipDiscovery() noexcept
{
    if (_discoveryClosed.load(std::memory_order_acquire))
    {
        return;
    }

    const ULONGLONG requestedTick = GetTickCount64();
    ULONGLONG noRequest           = 0u;
    if (! _discoverySkipRequestedTick.compare_exchange_strong(noRequest, requestedTick, std::memory_order_acq_rel) ||
        _discoverySkipped.exchange(true, std::memory_order_acq_rel))
    {
        return;
    }

    // The host owns the discovery-ahead reservation. Release it as part of the same
    // one-way transition; a provider that is inside an active call observes JIT mode
    // at its next bounded FileSystemGetDiscoveryMode checkpoint.
    _discoveryReservationReleasedTick.store(requestedTick, std::memory_order_release);
    _discoveryAheadActive.store(false, std::memory_order_release);
    LogDiagnostic(DiagnosticSeverity::Info,
                  S_FALSE,
                  L"discovery.skip",
                  L"User stopped discovery-ahead; traversal continues just in time without skipping safety checks.");
    _pauseCv.notify_all();
}

void FolderWindow::FileOperationState::Task::CloseDiscovery() noexcept
{
    if (_discoveryClosed.exchange(true, std::memory_order_acq_rel))
    {
        return;
    }

    _discoveryAheadActive.store(false, std::memory_order_release);
    if (_verificationRequested.load(std::memory_order_acquire))
    {
        AtomicMax(_verificationTotalBytes, _discoveredTotalBytes.load(std::memory_order_acquire));
    }
    const ULONGLONG closedTick = GetTickCount64();
    _discoveryClosedTick.store(closedTick, std::memory_order_release);
    if (_discoveryReservationReleasedTick.load(std::memory_order_acquire) == 0u)
    {
        _discoveryReservationReleasedTick.store(closedTick, std::memory_order_release);
    }

    const ULONGLONG startTick = _discoveryStartTick.load(std::memory_order_acquire);
    const uint64_t durationUs = startTick != 0u && closedTick >= startTick ? (closedTick - startTick) * 1000ull : 0u;
    _perf.discoveryOpenUs.store(durationUs, std::memory_order_release);
    LogDiagnostic(DiagnosticSeverity::Debug,
                  S_OK,
                  L"discovery.closed",
                  std::format(L"Single-pass traversal closed (elapsedMs={}, bytes={:L}, files={:L}, dirs={:L}, skippedAhead={}).",
                              durationUs / 1000ull,
                              _discoveredTotalBytes.load(std::memory_order_acquire),
                              _discoveredFileCount.load(std::memory_order_acquire),
                              _discoveredDirectoryCount.load(std::memory_order_acquire),
                              _discoverySkipped.load(std::memory_order_acquire) ? L"true" : L"false"));
    _pauseCv.notify_all();
}

void FolderWindow::FileOperationState::Task::MarkDiscoveryItemClosed(size_t sourceIndex) noexcept
{
    bool allClosed = false;
    {
        std::scoped_lock lock(_discoveryMutex);
        if (sourceIndex >= _discoveryItemClosed.size() || _discoveryItemClosed[sourceIndex] != 0u)
        {
            return;
        }
        _discoveryItemClosed[sourceIndex] = 1u;
        ++_closedDiscoveryItemCount;
        allClosed = _closedDiscoveryItemCount == _discoveryItemClosed.size();
    }
    if (allClosed)
    {
        CloseDiscovery();
    }
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void* /*cookie*/) noexcept
{
    if (mode == nullptr)
    {
        return E_POINTER;
    }
    const bool skipped = _discoverySkipped.load(std::memory_order_acquire);
    *mode              = skipped || _discoveryClosed.load(std::memory_order_acquire) ? FILESYSTEM_DISCOVERY_JUST_IN_TIME : FILESYSTEM_DISCOVERY_AHEAD;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress* progress,
                                                                                                    void* cookie) noexcept
{
    if (progress == nullptr)
    {
        return E_POINTER;
    }
    if (progress->sizeBytes != sizeof(FileSystemDiscoveryProgress))
    {
        return E_INVALIDARG;
    }

    const uint64_t startedUs = PerfNowUs();
    _perf.discoveryCallbackCount.fetch_add(1u, std::memory_order_relaxed);
    const auto recordDuration = wil::scope_exit([&]() noexcept { _perf.discoveryCallbackUs.fetch_add(PerfElapsedUs(startedUs), std::memory_order_relaxed); });

    auto* itemCookie          = static_cast<PerItemCallbackCookie*>(cookie);
    uint64_t bytesDelta       = progress->discoveredBytes;
    uint64_t filesDelta       = progress->discoveredFiles;
    uint64_t directoriesDelta = progress->discoveredDirectories;
    size_t sourceIndex        = 0u;
    if (itemCookie != nullptr)
    {
        sourceIndex = itemCookie->itemIndex;
        // A provider may report from concurrent workers. Serialize the cumulative-to-delta state,
        // and never reinterpret an out-of-order or retry-reset cumulative value as new discovery.
        std::scoped_lock lock(_discoveryMutex);
        bytesDelta       = progress->discoveredBytes >= itemCookie->lastDiscoveredBytes ? progress->discoveredBytes - itemCookie->lastDiscoveredBytes : 0u;
        filesDelta       = progress->discoveredFiles >= itemCookie->lastDiscoveredFiles ? progress->discoveredFiles - itemCookie->lastDiscoveredFiles : 0u;
        directoriesDelta = progress->discoveredDirectories >= itemCookie->lastDiscoveredDirectories
                               ? progress->discoveredDirectories - itemCookie->lastDiscoveredDirectories
                               : 0u;
        itemCookie->lastDiscoveredBytes       = (std::max)(itemCookie->lastDiscoveredBytes, progress->discoveredBytes);
        itemCookie->lastDiscoveredFiles       = (std::max)(itemCookie->lastDiscoveredFiles, progress->discoveredFiles);
        itemCookie->lastDiscoveredDirectories = (std::max)(itemCookie->lastDiscoveredDirectories, progress->discoveredDirectories);
    }

    SaturatingAtomicAdd(_discoveredTotalBytes, bytesDelta);

    const auto addClampedUlong = [](std::atomic<unsigned long>& target, uint64_t value) noexcept
    {
        unsigned long current = target.load(std::memory_order_acquire);
        for (;;)
        {
            const uint64_t desired64    = (std::min)(static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()), static_cast<uint64_t>(current) + value);
            const unsigned long desired = static_cast<unsigned long>(desired64);
            if (target.compare_exchange_weak(current, desired, std::memory_order_acq_rel))
            {
                return;
            }
        }
    };
    addClampedUlong(_discoveredFileCount, filesDelta);
    addClampedUlong(_discoveredDirectoryCount, directoriesDelta);
    AtomicMax(_discoveryMaxQueueDepth, progress->queuedItems);

    {
        const uint64_t lockStartedUs = PerfNowUs();
        std::scoped_lock lock(_progressMutex);
        _perf.discoveryLockWaitUs.fetch_add(PerfElapsedUs(lockStartedUs), std::memory_order_relaxed);
        _progressTotalBytes = (std::max)(_progressTotalBytes, _discoveredTotalBytes.load(std::memory_order_acquire));
        if (_operation == FILESYSTEM_DELETE)
        {
            const uint64_t discoveredItems = static_cast<uint64_t>(_discoveredFileCount.load(std::memory_order_acquire)) +
                                             static_cast<uint64_t>(_discoveredDirectoryCount.load(std::memory_order_acquire));
            _progressTotalItems = static_cast<unsigned long>((std::min)(discoveredItems, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max())));
        }
        PublishProgressCountersLocked(*this);
    }

    if (progress->traversalClosed != FALSE)
    {
        MarkDiscoveryItemClosed(sourceIndex);
    }
    return S_OK;
}
// Managed Move has already published the destination. Any remaining failure belongs exclusively
// to exact source cleanup, so destination attributes must not reinterpret ACCESS_DENIED as a
// replacement/read-only conflict. Retry/Skip/Cancel are the only applicable actions here.
[[nodiscard]] Task::ConflictBucket ClassifyManagedCleanupConflictBucket(HRESULT status) noexcept
{
    const std::optional<DWORD> error = Win32ErrorFromHRESULT(status);
    if (error.has_value() && error.value() == ERROR_ACCESS_DENIED)
    {
        return Task::ConflictBucket::AccessDenied;
    }

    return ClassifyConflictBucket(FILESYSTEM_DELETE, FILESYSTEM_FLAG_NONE, {}, status, std::wstring_view{}, std::wstring_view{}, false);
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemShouldAbort(BOOL* abort, void* cookie) noexcept
{
    return FileSystemShouldCancel(abort, cookie);
}

HRESULT FolderWindow::FileOperationState::Task::PrepareMutationInterlockScopes() noexcept
{
    _mutationInterlockScopes.clear();
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (! plans || plans->empty() || ! _fileSystem)
    {
        return E_UNEXPECTED;
    }

    constexpr FileSystemBindFlags bindFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    const uint64_t startedUs                = PerfNowUs();
    std::vector<std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode>> authorityNodes;
    uint64_t bindCount = 0u;
    const auto finish  = [&](HRESULT hr) noexcept
    {
        const auto countAccess = [&](FileOperations::MutationInterlockAccess access) noexcept
        {
            return static_cast<uint64_t>(std::ranges::count_if(
                _mutationInterlockScopes, [&](const FileOperations::MutationInterlockScope& scope) noexcept { return scope.access == access; }));
        };
        Debug::Perf::Emit(L"FileOps.Interlock.PrepareUs",
                          OperationToString(_operation),
                          PerfElapsedUs(startedUs),
                          static_cast<uint64_t>(_mutationInterlockScopes.size()),
                          static_cast<uint64_t>(authorityNodes.size()),
                          hr);
        Debug::Perf::Emit(L"FileOps.Interlock.ScopeCount",
                          L"read-source",
                          0u,
                          countAccess(FileOperations::MutationInterlockAccess::ReadSource),
                          static_cast<uint64_t>(_mutationInterlockScopes.size()),
                          hr);
        Debug::Perf::Emit(L"FileOps.Interlock.ScopeCount",
                          L"write-source",
                          0u,
                          countAccess(FileOperations::MutationInterlockAccess::WriteSource),
                          static_cast<uint64_t>(_mutationInterlockScopes.size()),
                          hr);
        Debug::Perf::Emit(L"FileOps.Interlock.ScopeCount",
                          L"publish-destination",
                          0u,
                          countAccess(FileOperations::MutationInterlockAccess::PublishDestination),
                          static_cast<uint64_t>(_mutationInterlockScopes.size()),
                          hr);
        Debug::Perf::Emit(L"FileOps.Interlock.BindCount", OperationToString(_operation), 0u, bindCount, static_cast<uint64_t>(authorityNodes.size()), hr);
        return hr;
    };

    const auto addScope = [&](IFileSystem* fileSystem,
                              const FileOperations::QualifiedEndpoint& endpoint,
                              std::wstring_view providerPath,
                              FileOperations::MutationInterlockAccess access,
                              bool bindingRequired,
                              bool retainDeleteAuthority) noexcept -> HRESULT
    {
        if (fileSystem == nullptr || providerPath.empty() || ! endpoint.pathIdentity.has_value())
        {
            return E_INVALIDARG;
        }
        const auto existingScope = std::ranges::find_if(_mutationInterlockScopes,
                                                        [&](const FileOperations::MutationInterlockScope& existing) noexcept
        {
            return FileOperations::QualifiedEndpointsReferToSameRoot(existing.endpoint, endpoint) &&
                   EquivalentPath(endpoint.pathIdentity.value(), existing.providerPath, providerPath);
        });
        if (existingScope != _mutationInterlockScopes.end())
        {
            if (retainDeleteAuthority && ! existingScope->exactDeleteAuthority)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
            if (existingScope->access == FileOperations::MutationInterlockAccess::ReadSource)
            {
                existingScope->access = access;
            }
            else if (access == FileOperations::MutationInterlockAccess::WriteSource)
            {
                existingScope->access = access;
            }
            return S_OK;
        }

        const auto findAuthorityNode = [&](std::wstring_view path) noexcept -> std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode>
        {
            const auto existing = std::ranges::find_if(authorityNodes,
                                                       [&](const std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode>& node) noexcept
            {
                return node && FileOperations::QualifiedEndpointsReferToSameRoot(node->endpoint, endpoint) &&
                       EquivalentPath(endpoint.pathIdentity.value(), node->retained.providerPath, path);
            });
            return existing != authorityNodes.end() ? *existing : nullptr;
        };

        FileOperations::MutationInterlockScope scope{};
        scope.endpoint             = endpoint;
        scope.providerPath         = std::wstring(providerPath);
        scope.access               = access;
        scope.exactDeleteAuthority = retainDeleteAuthority;

        const auto retainBoundChain = [&](std::wstring startingPath, FileOperations::BoundObjectAuthority startingAuthority) noexcept -> HRESULT
        {
            std::vector<FileOperations::RetainedPathAuthority> retainedChain;
            retainedChain.emplace_back(FileOperations::RetainedPathAuthority{
                .providerPath = startingPath,
                .authority    = std::move(startingAuthority),
            });
            std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode> retainedParent;

            std::wstring current = std::move(startingPath);
            for (size_t depth = 0u; depth < 1024u; ++depth)
            {
                std::wstring parent;
                if (! TryGetFileSystemParentPath(endpoint.pathIdentity.value(), current, parent) ||
                    EquivalentPath(endpoint.pathIdentity.value(), current, parent))
                {
                    current.clear();
                    break;
                }

                if (std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode> existingParent = findAuthorityNode(parent))
                {
                    retainedParent = std::move(existingParent);
                    current.clear();
                    break;
                }

                ++bindCount;
                FileOperations::ObjectBindingResult ancestor = FileOperations::BindObjectAuthority(fileSystem, parent, endpoint.profileId, bindFlags);
                if (ancestor.state == FileOperations::ObjectBindingState::Bound)
                {
                    retainedChain.emplace_back(FileOperations::RetainedPathAuthority{
                        .providerPath = parent,
                        .authority    = std::move(ancestor.authority),
                    });
                    current = std::move(parent);
                    continue;
                }
                if (ancestor.state == FileOperations::ObjectBindingState::Indeterminate ||
                    ancestor.state == FileOperations::ObjectBindingState::ProviderContractViolation)
                {
                    return FAILED(ancestor.status) ? ancestor.status : E_UNEXPECTED;
                }
                if (ancestor.state == FileOperations::ObjectBindingState::Unsupported)
                {
                    if (bindingRequired)
                    {
                        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                    }
                    scope.conservativeIdentityDomain = true;
                }
                current.clear();
                break;
            }
            if (! current.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
            }

            for (auto retained = retainedChain.rbegin(); retained != retainedChain.rend(); ++retained)
            {
                auto node      = std::make_shared<FileOperations::MutationInterlockAuthorityNode>(FileOperations::MutationInterlockAuthorityNode{
                    .endpoint = endpoint,
                    .retained = std::move(*retained),
                    .parent   = std::move(retainedParent),
                });
                retainedParent = node;
                authorityNodes.emplace_back(std::move(node));
            }
            scope.root = std::move(retainedParent);
            return S_OK;
        };

        if (std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode> existingRoot = findAuthorityNode(providerPath))
        {
            scope.root = std::move(existingRoot);
        }
        else
        {
            const FileSystemBindFlags rootBindFlags = retainDeleteAuthority ? static_cast<FileSystemBindFlags>(bindFlags | FILESYSTEM_BIND_DELETE) : bindFlags;
            ++bindCount;
            FileOperations::ObjectBindingResult root = FileOperations::BindObjectAuthority(fileSystem, providerPath, endpoint.profileId, rootBindFlags);
            if (root.state == FileOperations::ObjectBindingState::Bound)
            {
                const HRESULT retainHr = retainBoundChain(std::wstring(providerPath), std::move(root.authority));
                if (FAILED(retainHr))
                {
                    return retainHr;
                }
            }
            else if (root.state == FileOperations::ObjectBindingState::Missing && ! retainDeleteAuthority)
            {
                std::wstring current(providerPath);
                for (size_t depth = 0u; depth < 1024u; ++depth)
                {
                    std::wstring parent;
                    if (! TryGetFileSystemParentPath(endpoint.pathIdentity.value(), current, parent) ||
                        EquivalentPath(endpoint.pathIdentity.value(), current, parent))
                    {
                        scope.conservativeIdentityDomain = true;
                        current.clear();
                        break;
                    }
                    if (std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode> existingParent = findAuthorityNode(parent))
                    {
                        scope.root = std::move(existingParent);
                        current.clear();
                        break;
                    }

                    ++bindCount;
                    FileOperations::ObjectBindingResult ancestor = FileOperations::BindObjectAuthority(fileSystem, parent, endpoint.profileId, bindFlags);
                    if (ancestor.state == FileOperations::ObjectBindingState::Bound)
                    {
                        const HRESULT retainHr = retainBoundChain(parent, std::move(ancestor.authority));
                        if (FAILED(retainHr))
                        {
                            return retainHr;
                        }
                        current.clear();
                        break;
                    }
                    if (ancestor.state == FileOperations::ObjectBindingState::Indeterminate ||
                        ancestor.state == FileOperations::ObjectBindingState::ProviderContractViolation)
                    {
                        return FAILED(ancestor.status) ? ancestor.status : E_UNEXPECTED;
                    }
                    if (ancestor.state == FileOperations::ObjectBindingState::Unsupported)
                    {
                        scope.conservativeIdentityDomain = true;
                        current.clear();
                        break;
                    }
                    current = std::move(parent);
                }
                if (! current.empty())
                {
                    return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
                }
            }
            else if (root.state == FileOperations::ObjectBindingState::Unsupported)
            {
                if (bindingRequired)
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
                scope.conservativeIdentityDomain = true;
            }
            else
            {
                return FAILED(root.status) ? root.status : E_UNEXPECTED;
            }
        }

        for (std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode> current = scope.root; current; current = current->parent)
        {
            std::wstring relativePath;
            if (! TryGetFileSystemRelativePath(endpoint.pathIdentity.value(), current->retained.providerPath, providerPath, relativePath))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            std::optional<std::wstring> relativePathKey = TryMakePathKey(endpoint.pathIdentity.value(), relativePath);
            scope.anchors.emplace_back(FileOperations::MutationInterlockAnchor{
                .authority       = current,
                .relativePath    = std::move(relativePath),
                .relativePathKey = std::move(relativePathKey),
            });
        }

        _mutationInterlockScopes.emplace_back(std::move(scope));
        return S_OK;
    };

    for (const FileOperations::FileOperationPlan& plan : *plans)
    {
        const HRESULT planHr = std::visit(
            [&](const auto& typedPlan) noexcept -> HRESULT
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
            {
                IFileSystem* const destinationFileSystem = _destinationFileSystem ? _destinationFileSystem.get() : _fileSystem.get();
                HRESULT hr                               = S_OK;
                for (size_t sourceIndex = 0u; sourceIndex < typedPlan.selectedItems.size(); ++sourceIndex)
                {
                    std::wstring destinationLeaf;
                    if (! FileOperations::TryResolveTransferDestinationProviderPath(typedPlan, sourceIndex, destinationLeaf))
                    {
                        return E_INVALIDARG;
                    }
                    hr = addScope(destinationFileSystem,
                                  typedPlan.destinationEndpoint,
                                  destinationLeaf,
                                  FileOperations::MutationInterlockAccess::PublishDestination,
                                  false,
                                  false);
                    if (FAILED(hr))
                    {
                        return hr;
                    }
                }
                const FileOperations::MutationInterlockAccess sourceAccess =
                    typedPlan.intent == FileOperations::TransferIntent::Copy || typedPlan.strategy == FileOperations::OperationStrategy::CopyOnly
                        ? FileOperations::MutationInterlockAccess::ReadSource
                        : FileOperations::MutationInterlockAccess::WriteSource;
                for (const FileOperations::QualifiedSourceItem& item : typedPlan.selectedItems)
                {
                    hr = addScope(_fileSystem.get(), typedPlan.sourceEndpoint, item.providerPath, sourceAccess, false, false);
                    if (FAILED(hr))
                    {
                        return hr;
                    }
                }
                return S_OK;
            }
            else if constexpr (std::is_same_v<Plan, FileOperations::DeletePlan>)
            {
                const bool permanentDelete = typedPlan.mode == FileOperations::DeleteMode::Permanent;
                // Native-authority deletes (no bound objects) keep the interlock scope for overlap
                // detection but cannot retain exact delete authority; the provider owns the receipt.
                const bool exactDeleteAuthority = permanentDelete && ! typedPlan.nativeAuthority;
                for (const FileOperations::QualifiedSourceItem& item : typedPlan.selectedItems)
                {
#ifdef ENABLE_TESTS
                    if (permanentDelete)
                    {
                        g_fileOpsPermanentDeleteBeforeBindPausePoint.Pause(10'000ull);
                    }
#endif
                    const HRESULT hr = addScope(_fileSystem.get(),
                                                typedPlan.endpoint,
                                                item.providerPath,
                                                FileOperations::MutationInterlockAccess::WriteSource,
                                                exactDeleteAuthority,
                                                exactDeleteAuthority);
                    if (FAILED(hr))
                    {
                        if (exactDeleteAuthority)
                        {
                            LogDiagnostic(DiagnosticSeverity::Error,
                                          hr,
                                          L"preparation.deleteRoot",
                                          L"The selected item could not be pinned for deletion; nothing was deleted.",
                                          item.providerPath);
                        }
                        return hr;
                    }
                    if (permanentDelete && typedPlan.nativeAuthority)
                    {
                        // C10: a provider that can name the object pins its identity here; the card's
                        // re-check and the delete compare against it. A provider that cannot keeps
                        // the path delete, and the card says so.
                        const HRESULT pinHr = PinNativeDeleteIdentity(typedPlan.endpoint, item.providerPath);
                        if (FAILED(pinHr))
                        {
                            LogDiagnostic(DiagnosticSeverity::Error,
                                          pinHr,
                                          L"preparation.deleteRoot",
                                          L"The selected item could not be identified for deletion; nothing was deleted.",
                                          item.providerPath);
                            return pinHr;
                        }
                    }
                    // C1: the exact authority retained above pins the object the card confirms;
                    // RunPermanentDeleteConfirmation re-checks every pinned pathname after the answer.
                }
                return S_OK;
            }
            else
            {
                // Rename publications land under these parent WriteSource scopes; the live-output
                // index and its overflow fallback correlate them there.
                _publishesUnderWriteSourceScopes = true;
                for (const FileOperations::RenameStep& step : typedPlan.finalMappings)
                {
                    std::wstring parent;
                    if (! TryGetFileSystemParentPath(typedPlan.endpoint.pathIdentity.value(), step.source.providerPath, parent))
                    {
                        return E_INVALIDARG;
                    }
                    const HRESULT hr =
                        addScope(_fileSystem.get(), typedPlan.endpoint, parent, FileOperations::MutationInterlockAccess::WriteSource, true, false);
                    if (FAILED(hr))
                    {
                        return hr;
                    }
                }
                return S_OK;
            }
        },
            plan);
        if (FAILED(planHr))
        {
            return finish(planHr);
        }
    }

    std::ranges::sort(_mutationInterlockScopes,
                      [](const FileOperations::MutationInterlockScope& left, const FileOperations::MutationInterlockScope& right) noexcept
    {
        return std::tie(left.endpoint.pluginId, left.endpoint.instanceId, left.endpoint.profileId, left.endpoint.rootId, left.providerPath) <
               std::tie(right.endpoint.pluginId, right.endpoint.instanceId, right.endpoint.profileId, right.endpoint.rootId, right.providerPath);
    });
    return finish(S_OK);
}

HRESULT FolderWindow::FileOperationState::Task::BuildPreparationSnapshot(const uint64_t selectedRootReadinessUs) noexcept
{
    const uint64_t startedUs                                                  = PerfNowUs();
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (! plans || plans->empty())
    {
        return E_UNEXPECTED;
    }

    PreparationSnapshot snapshot{};
    snapshot.taskId    = _taskId;
    snapshot.operation = _operation;
    snapshot.status    = S_OK;
    snapshot.scopes.reserve(_mutationInterlockScopes.size());

    std::array<uint64_t, 4u> strategyCounts{};
    const auto addCount = [](uint64_t& target, const size_t count) noexcept
    {
        const uint64_t converted =
            count > static_cast<size_t>((std::numeric_limits<uint64_t>::max)()) ? (std::numeric_limits<uint64_t>::max)() : static_cast<uint64_t>(count);
        target = converted > (std::numeric_limits<uint64_t>::max)() - target ? (std::numeric_limits<uint64_t>::max)() : target + converted;
    };

    for (const FileOperations::FileOperationPlan& plan : *plans)
    {
        std::visit(
            [&](const auto& typedPlan) noexcept
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
            {
                addCount(snapshot.selectedRootCount, typedPlan.selectedItems.size());
                const size_t strategyIndex = static_cast<size_t>(typedPlan.strategy);
                if (strategyIndex < strategyCounts.size())
                {
                    addCount(strategyCounts[strategyIndex], typedPlan.selectedItems.size());
                }
                if (typedPlan.strategy == FileOperations::OperationStrategy::CopyOnly)
                {
                    addCount(snapshot.copyOnlyCount, typedPlan.selectedItems.size());
                }
            }
            else if constexpr (std::is_same_v<Plan, FileOperations::DeletePlan>)
            {
                addCount(snapshot.selectedRootCount, typedPlan.selectedItems.size());
            }
            else
            {
                addCount(snapshot.selectedRootCount, typedPlan.finalMappings.size());
            }
        },
            plan);
    }

    for (size_t strategyIndex = 0u; strategyIndex < strategyCounts.size(); ++strategyIndex)
    {
        if (strategyCounts[strategyIndex] == 0u)
        {
            continue;
        }
        snapshot.strategies.emplace_back(FileOperations::PreparationStrategyFact{
            .strategy          = static_cast<FileOperations::OperationStrategy>(strategyIndex),
            .selectedRootCount = strategyCounts[strategyIndex],
        });
    }

    for (const FileOperations::MutationInterlockScope& scope : _mutationInterlockScopes)
    {
        const auto endpointIt = std::ranges::find_if(snapshot.endpoints,
                                                     [&](const FileOperations::PreparationEndpointFact& endpoint) noexcept
        {
            return endpoint.pluginId == scope.endpoint.pluginId && endpoint.instanceId == scope.endpoint.instanceId &&
                   endpoint.profileId == scope.endpoint.profileId && endpoint.rootId == scope.endpoint.rootId;
        });
        size_t endpointIndex  = static_cast<size_t>(endpointIt - snapshot.endpoints.begin());
        if (endpointIt == snapshot.endpoints.end())
        {
            endpointIndex = snapshot.endpoints.size();
            if (endpointIndex > static_cast<size_t>((std::numeric_limits<uint32_t>::max)()))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            snapshot.endpoints.emplace_back(FileOperations::PreparationEndpointFact{
                .pluginId   = scope.endpoint.pluginId,
                .instanceId = scope.endpoint.instanceId,
                .profileId  = scope.endpoint.profileId,
                .rootId     = scope.endpoint.rootId,
            });
        }
        snapshot.scopes.emplace_back(FileOperations::PreparationScopeFact{
            .endpointIndex              = static_cast<uint32_t>(endpointIndex),
            .providerPath               = scope.providerPath,
            .access                     = scope.access,
            .exactDeleteAuthority       = scope.exactDeleteAuthority,
            .conservativeIdentityDomain = scope.conservativeIdentityDomain,
        });
    }

    uint64_t retainedBytes = sizeof(PreparationSnapshot);
    retainedBytes += static_cast<uint64_t>(snapshot.endpoints.capacity()) * sizeof(FileOperations::PreparationEndpointFact);
    for (const FileOperations::PreparationEndpointFact& endpoint : snapshot.endpoints)
    {
        retainedBytes +=
            static_cast<uint64_t>(endpoint.pluginId.capacity() + endpoint.instanceId.capacity() + endpoint.profileId.capacity() + endpoint.rootId.capacity()) *
            sizeof(wchar_t);
    }
    retainedBytes += static_cast<uint64_t>(snapshot.scopes.capacity()) * sizeof(FileOperations::PreparationScopeFact);
    for (const FileOperations::PreparationScopeFact& scope : snapshot.scopes)
    {
        retainedBytes += static_cast<uint64_t>(scope.providerPath.capacity()) * sizeof(wchar_t);
    }
    retainedBytes += static_cast<uint64_t>(snapshot.strategies.capacity()) * sizeof(FileOperations::PreparationStrategyFact);
    snapshot.retainedBytes = retainedBytes;

    auto published                                             = std::make_shared<PreparationSnapshot>(std::move(snapshot));
    published->buildSnapshotUs                                 = PerfElapsedUs(startedUs);
    published->selectedRootReadinessUs                         = selectedRootReadinessUs > (std::numeric_limits<uint64_t>::max)() - published->buildSnapshotUs
                                                                     ? (std::numeric_limits<uint64_t>::max)()
                                                                     : selectedRootReadinessUs + published->buildSnapshotUs;
    const std::shared_ptr<const PreparationSnapshot> immutable = std::move(published);
    _preparationSnapshot.store(immutable, std::memory_order_release);
    if (_preparationObserver)
    {
        _preparationObserver(immutable);
    }

    Debug::Perf::Emit(L"FileOps.Preparing.BuildSnapshotUs",
                      OperationToString(_operation),
                      immutable->buildSnapshotUs,
                      immutable->selectedRootCount,
                      static_cast<uint64_t>(immutable->scopes.size()),
                      S_OK);
    Debug::Perf::Emit(L"FileOps.Preparing.SelectedRootReadinessUs",
                      OperationToString(_operation),
                      immutable->selectedRootReadinessUs,
                      immutable->selectedRootCount,
                      static_cast<uint64_t>(immutable->scopes.size()),
                      S_OK);
    Debug::Perf::Emit(
        L"FileOps.Preparing.SelectedRootCount", OperationToString(_operation), 0u, immutable->selectedRootCount, immutable->selectedRootCount, S_OK);
    Debug::Perf::Emit(L"FileOps.Preparing.ScopeFactCount",
                      OperationToString(_operation),
                      0u,
                      static_cast<uint64_t>(immutable->scopes.size()),
                      immutable->selectedRootCount,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.Preparing.StrategyFactCount",
                      OperationToString(_operation),
                      0u,
                      static_cast<uint64_t>(immutable->strategies.size()),
                      immutable->selectedRootCount,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.Preparing.CopyOnlyCount", OperationToString(_operation), 0u, immutable->copyOnlyCount, immutable->selectedRootCount, S_OK);
    Debug::Perf::Emit(L"FileOps.Preparing.RetainedBytes", OperationToString(_operation), 0u, immutable->retainedBytes, 64u * 1024u * 1024u, S_OK);
    return S_OK;
}

HRESULT FolderWindow::FileOperationState::Task::RunPreConsumptionDecisionGate() noexcept
{
    return _preConsumptionDecisionGate ? _preConsumptionDecisionGate() : S_OK;
}

bool FolderWindow::FileOperationState::Task::IsInlineRenameTask() const noexcept
{
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (! plans)
    {
        return false;
    }
    for (const FileOperations::FileOperationPlan& plan : *plans)
    {
        if (const auto* rename = std::get_if<FileOperations::RenamePlan>(&plan);
            rename != nullptr && rename->origin == FileOperations::RenameOrigin::InlineRename)
        {
            return true;
        }
    }
    return false;
}

HRESULT FolderWindow::FileOperationState::Task::PinNativeDeleteIdentity(const FileOperations::QualifiedEndpoint& endpoint,
                                                                        std::wstring_view providerPath) noexcept
{
    wil::com_ptr<IFileSystemIdentityDelete> identityDelete;
    if (! _fileSystem || FAILED(_fileSystem->QueryInterface(__uuidof(IFileSystemIdentityDelete), identityDelete.put_void())) || ! identityDelete)
    {
        return S_OK; // no identity contract: the native path delete stays, and the card says so
    }
    FileSystemOptions options{};
    InitializeFileSystemOptions(options, nullptr);
    FileSystemDeleteIdentity identity{};
    identity.sizeBytes = sizeof(identity);
    const std::wstring path(providerPath);
#ifdef ENABLE_TESTS
    _dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
    const auto dbgCallbackScope = wil::scope_exit([&] noexcept { _dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
#endif
    const HRESULT hr = identityDelete->ResolveDeleteIdentity(path.c_str(), &options, &identity);
    if (FAILED(hr))
    {
        return hr;
    }
    identity.identity[std::size(identity.identity) - 1u] = L'\0';
    if (identity.identity[0] == L'\0')
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    const auto scope = std::ranges::find_if(_mutationInterlockScopes,
                                            [&](const FileOperations::MutationInterlockScope& existing) noexcept
    {
        return FileOperations::QualifiedEndpointsReferToSameRoot(existing.endpoint, endpoint) &&
               EquivalentPath(endpoint.pathIdentity.value(), existing.providerPath, providerPath);
    });
    if (scope == _mutationInterlockScopes.end())
    {
        return E_UNEXPECTED;
    }
    scope->pinnedDeleteIdentity = identity;
    return S_OK;
}

bool FolderWindow::FileOperationState::Task::PermanentDeleteRunsByNameOnly() const noexcept
{
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (! plans)
    {
        return false;
    }
    bool nativeAuthority = false;
    for (const FileOperations::FileOperationPlan& plan : *plans)
    {
        const auto* deletion = std::get_if<FileOperations::DeletePlan>(&plan);
        if (deletion != nullptr && deletion->mode == FileOperations::DeleteMode::Permanent && deletion->nativeAuthority)
        {
            nativeAuthority = true;
        }
    }
    if (! nativeAuthority)
    {
        return false;
    }
    return std::ranges::none_of(_mutationInterlockScopes,
                                [](const FileOperations::MutationInterlockScope& scope) noexcept { return scope.pinnedDeleteIdentity.has_value(); });
}

HRESULT FolderWindow::FileOperationState::Task::RunPermanentDeleteConfirmation() noexcept
{
    // C1: the permanent-delete confirmation that used to be a modal at ingress, before any bind,
    // runs on the card after Preparing has pinned every selected root. The user confirms objects
    // that are pinned, in the words the modal used; a pathname that no longer names its pinned
    // object after the answer refuses, so nothing but what was confirmed is ever deleted.
    if (! _permanentDeleteConfirmationPending)
    {
        return S_OK;
    }
    ConflictAction action = ConflictAction::None;
#ifdef ENABLE_TESTS
    if (_permanentDeleteTestOverride == 1)
    {
        action = ConflictAction::PermanentDelete;
    }
    else if (_permanentDeleteTestOverride == 2)
    {
        action = ConflictAction::Cancel;
    }
#endif
    if (action == ConflictAction::None)
    {
        DeferredConsentRequest request{};
        request.risk          = FileOperations::DeferredConsentRisk::PermanentDelete;
        request.status        = S_OK;
        request.sourcePath    = _permanentDeleteConsentFrom;
        request.consentDetail = _permanentDeleteConsentDetail;
        if (PermanentDeleteRunsByNameOnly())
        {
            // C10: a route with neither object binding nor an identity contract deletes whatever
            // the name refers to when the mutation runs; the card says so before the answer.
            if (! request.consentDetail.empty())
            {
                request.consentDetail += L' ';
            }
            request.consentDetail += LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_PERMANENT_DELETE_BY_NAME);
        }
        request.itemCountKnown             = true;
        request.itemCount                  = _sourcePaths.size();
        const DeferredConsentResult result = RequestDeferredConsent(*this, request);
        if (FAILED(result.status))
        {
            return result.status;
        }
        action = result.action;
    }
    if (action != ConflictAction::PermanentDelete)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
#ifdef ENABLE_TESTS
    g_fileOpsPermanentDeleteBeforeRecheckPausePoint.Pause(10'000ull);
#endif
    // Post-confirmation continuity: every pinned root must still be what its pathname names.
    for (const FileOperations::MutationInterlockScope& scope : _mutationInterlockScopes)
    {
        if (! scope.exactDeleteAuthority || ! scope.root)
        {
            continue;
        }
        constexpr FileSystemBindFlags identityFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
        const FileOperations::ObjectBindingResult current =
            FileOperations::BindObjectAuthority(_fileSystem.get(), scope.providerPath, scope.endpoint.profileId, identityFlags);
        const FileOperations::ProviderIdentitySnapshot& pinned = scope.root->retained.authority.identity;
        if (current.state != FileOperations::ObjectBindingState::Bound || current.authority.identity.pathProfileId != pinned.pathProfileId ||
            current.authority.identity.objectId != pinned.objectId || current.authority.identity.revisionId != pinned.revisionId)
        {
            _permanentDeleteIdentityMismatchCount.fetch_add(1u, std::memory_order_relaxed);
            LogDiagnostic(DiagnosticSeverity::Error,
                          HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH),
                          L"delete.confirmation.identityChanged",
                          L"The selected item changed after it was confirmed; nothing was deleted.",
                          scope.providerPath);
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
    }
    // C10: a native-authority root pinned by identity must still carry that identity.
    for (const FileOperations::MutationInterlockScope& scope : _mutationInterlockScopes)
    {
        if (! scope.pinnedDeleteIdentity.has_value())
        {
            continue;
        }
        wil::com_ptr<IFileSystemIdentityDelete> identityDelete;
        if (! _fileSystem || FAILED(_fileSystem->QueryInterface(__uuidof(IFileSystemIdentityDelete), identityDelete.put_void())) || ! identityDelete)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        FileSystemOptions options{};
        InitializeFileSystemOptions(options, nullptr);
        FileSystemDeleteIdentity current{};
        current.sizeBytes = sizeof(current);
#ifdef ENABLE_TESTS
        _dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const auto dbgCallbackScope = wil::scope_exit([&] noexcept { _dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
#endif
        const HRESULT hr                                   = identityDelete->ResolveDeleteIdentity(scope.providerPath.c_str(), &options, &current);
        current.identity[std::size(current.identity) - 1u] = L'\0';
        if (FAILED(hr) || current.isDirectory != scope.pinnedDeleteIdentity->isDirectory ||
            std::wstring_view(current.identity) != std::wstring_view(scope.pinnedDeleteIdentity->identity))
        {
            _permanentDeleteIdentityMismatchCount.fetch_add(1u, std::memory_order_relaxed);
            LogDiagnostic(DiagnosticSeverity::Error,
                          HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH),
                          L"delete.confirmation.identityChanged",
                          L"The selected item changed after it was confirmed; nothing was deleted.",
                          scope.providerPath);
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
    }
    // Record the consent as the modal did: the kind by origin, the nonce this task.
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (! plans)
    {
        return E_UNEXPECTED;
    }
    auto consented = std::make_shared<FileOperations::FileOperationPlanGroup>(*plans);
    for (FileOperations::FileOperationPlan& plan : *consented)
    {
        auto* deletion = std::get_if<FileOperations::DeletePlan>(&plan);
        if (deletion == nullptr || deletion->mode != FileOperations::DeleteMode::Permanent || deletion->initialConsent.has_value())
        {
            continue;
        }
        const FileOperations::ConsentKind kind =
            deletion->origin == FileOperations::DeleteOrigin::PackCleanup || deletion->origin == FileOperations::DeleteOrigin::UnpackCleanup
                ? FileOperations::ConsentKind::ArchiveDeleteAfter
                : FileOperations::ConsentKind::PermanentDelete;
        deletion->initialConsent = FileOperations::DestructiveConsentReceipt{.kind = kind, .taskNonce = _taskId};
    }
    StorePlans(std::move(consented));
    _permanentDeleteConfirmationPending = false;
    return S_OK;
}

HRESULT FolderWindow::FileOperationState::Task::RunSameHostOverlapAdvisory() noexcept
{
    // Whatever this advisory decides (nothing to warn, Queue after, Run concurrently, Don't start),
    // a peer that publishes its scopes afterwards may name this task; the relation arrays are
    // final before this release store.
    const auto advisoryDecided = wil::scope_exit([this]() noexcept { _overlapAdvisoryDecided.store(true, std::memory_order_release); });
    // Inline F2 keeps the silent interlock: it has no consent surface before its card reveals.
    if (_state == nullptr || IsInlineRenameTask())
    {
        return S_OK;
    }
    const FileOperationState::SameHostOverlapAdvice advice = _state->FindSameHostOverlap(*this);
    if (advice.problem == FileOperations::SameHostOverlapProblem::None)
    {
        return S_OK;
    }
    Debug::Perf::EmitValue(L"FileOps.Overlap.AdvisoryShown", advice.taskCount, S_OK);

    const std::wstring sourcePath      = _sourcePaths.empty() ? std::wstring() : _sourcePaths.front().native();
    const std::wstring destinationPath = GetDestinationFolder().native();
    DeferredConsentRequest request{};
    request.risk                       = FileOperations::DeferredConsentRisk::SameHostOverlap;
    request.status                     = HRESULT_FROM_WIN32(ERROR_BUSY);
    request.sourcePath                 = sourcePath;
    request.destinationPath            = destinationPath;
    request.itemCountKnown             = true;
    request.itemCount                  = advice.taskCount;
    request.overlapProblem             = static_cast<uint8_t>(advice.problem);
    request.overlapTaskId              = advice.firstTaskId;
    request.overlapTaskIds             = advice.taskIds;
    request.overlapTaskIdCount         = advice.relationCount;
    request.allowConcurrentRun         = advice.CanRunConcurrently();
    const DeferredConsentResult result = RequestDeferredConsent(*this, request);
    if (FAILED(result.status) || (result.action != ConflictAction::Proceed && result.action != ConflictAction::RunConcurrently))
    {
        Debug::Perf::EmitValue(L"FileOps.Overlap.DontStart", 1u, HRESULT_FROM_WIN32(ERROR_CANCELLED));
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    if (result.action == ConflictAction::RunConcurrently)
    {
        _concurrentOverlapTaskIds     = advice.taskIds;
        _concurrentOverlapTaskIdCount = advice.relationCount;
        // This explicit per-task choice supersedes the global wait mode only for admission;
        // undisclosed overlapping active tasks still block in EnterOperation.
        _waitForOthers.store(false, std::memory_order_release);
        Debug::Perf::EmitValue(L"FileOps.Overlap.RunConcurrent", advice.taskCount, S_OK);
        return S_OK;
    }

    // Queue after the exact disclosed predecessor set. The fixed IDs are discarded with the task;
    // each edge releases when its prepared/active publication disappears, without a second warning.
    _overlapPredecessorTaskIds     = advice.taskIds;
    _overlapPredecessorTaskIdCount = advice.relationCount;
    // Queue mode admits the queue head only. A task that queues after a younger peer would sit
    // ahead of that peer in key order and wait for it forever, so it takes a fresh key behind
    // every task admitted so far; a task that queues after older peers keeps its own place.
    const bool queuesAfterYoungerPeer = std::ranges::any_of(advice.taskIds.begin(),
                                                            advice.taskIds.begin() + static_cast<std::ptrdiff_t>(advice.relationCount),
                                                            [this](uint64_t id) noexcept { return id > _taskId; });
    if (queuesAfterYoungerPeer)
    {
        _queueOrderKey.store(_state->AllocateLateQueueOrderKey(), std::memory_order_release);
    }
    _overlapQueued.store(true, std::memory_order_release);
    Debug::Perf::EmitValue(L"FileOps.Overlap.QueuedAfter", advice.taskCount, S_OK);
    return S_OK;
}

void FolderWindow::FileOperationState::Task::NoteLiveOutputPublished(std::wstring_view providerPath) noexcept
{
    if (_state != nullptr && ! providerPath.empty())
    {
        _state->NoteLiveOutputPublished(*this, providerPath);
    }
#ifdef ENABLE_TESTS
    g_fileOpsLiveOutputPublishedPausePoint.Pause(10'000ull);
#endif
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::Task::DebugClearConcurrentOverlapRelationsForSelfTest() noexcept
{
    if (_state != nullptr)
    {
        _state->DebugClearConcurrentOverlapRelationsForSelfTest(*this);
    }
}
#endif

FolderWindow::FileOperationState::Task::LiveOutputGuardDisposition FolderWindow::FileOperationState::Task::GuardLiveOutputBeforeInvalidation(
    std::wstring_view providerPath, FileOperations::MutationInterlockAccess access) noexcept
{
    if (_state == nullptr || providerPath.empty())
    {
        return LiveOutputGuardDisposition::Proceed;
    }

    std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations> approvedPublishers{};
    size_t approvedPublisherCount = 0u;
    bool waitedForPublisher       = false;
    for (;;)
    {
        const FileOperationState::LiveOutputConflictAdvice advice =
            _state->FindLiveOutputConflict(*this, providerPath, access, approvedPublishers, approvedPublisherCount);
        if (advice.publisherTaskId == 0u)
        {
            return waitedForPublisher ? LiveOutputGuardDisposition::RetryCurrentMutation : LiveOutputGuardDisposition::Proceed;
        }

        Debug::Perf::EmitValue(L"FileOps.LiveOutput.GateCount", 1u, S_OK);
        DeferredConsentRequest request{};
        request.risk                       = FileOperations::DeferredConsentRisk::SameHostLiveOutput;
        request.status                     = HRESULT_FROM_WIN32(ERROR_BUSY);
        request.sourcePath                 = providerPath;
        request.destinationPath            = providerPath;
        request.overlapTaskId              = advice.publisherTaskId;
        request.itemCountKnown             = true;
        request.itemCount                  = 1u;
        request.allowLiveInvalidation      = ! advice.indexOverflow && approvedPublisherCount < approvedPublishers.size();
        const DeferredConsentResult result = RequestDeferredConsent(*this, request);
        if (FAILED(result.status) || result.action == ConflictAction::Cancel)
        {
            return LiveOutputGuardDisposition::Cancel;
        }
        if (result.action == ConflictAction::Skip)
        {
            _observedSkipAction.store(true, std::memory_order_release);
            return LiveOutputGuardDisposition::Skip;
        }
        if (result.action == ConflictAction::QueueUntilOtherTaskFinishes)
        {
            if (! _state->WaitForLiveOutputPublisher(*this, advice.publisherTaskId, _stopToken))
            {
                return LiveOutputGuardDisposition::Cancel;
            }
            waitedForPublisher = true;
            continue;
        }
        if (result.action != ConflictAction::InvalidateLiveOutput || approvedPublisherCount >= approvedPublishers.size())
        {
            return LiveOutputGuardDisposition::Cancel;
        }
        approvedPublishers[approvedPublisherCount] = advice.publisherTaskId;
        ++approvedPublisherCount;
    }
}

HRESULT FolderWindow::FileOperationState::Task::PrepareTransferDestinationNames(const FileOperations::FileOperationPlanGroup& plans) noexcept
{
    // R0-RC3 (9) / R1d-OR1: every destination leaf of a transfer passes the destination provider's
    // child-name contract before any path is joined, so a name the destination cannot hold fails
    // the task here, with the provider's reason, instead of failing (or being silently altered)
    // at stage creation. This is a provider call, so it runs on the task thread under the
    // synchronous-I/O cancel watch. Only an answered contract that rejects the name refuses; a
    // provider without the contract, or one that cannot answer for this parent, keeps the plain
    // join.
    IFileSystem* const destinationFileSystem = _destinationFileSystem ? _destinationFileSystem.get() : _fileSystem.get();
    if (destinationFileSystem == nullptr)
    {
        return S_OK;
    }
    // C1 (b0): an inline rename published with its join pending gets the provider's answer here.
    for (size_t planIndex = 0u; planIndex < plans.size(); ++planIndex)
    {
        const auto* rename = std::get_if<FileOperations::RenamePlan>(&plans[planIndex]);
        if (rename == nullptr || rename->origin != FileOperations::RenameOrigin::InlineRename || rename->finalMappings.size() != 1u ||
            ! rename->endpoint.pathIdentity.has_value() || ! rename->finalMappings.front().providerJoinedPath.empty())
        {
            continue;
        }
        const FileOperations::RenameStep& step = rename->finalMappings.front();
        std::wstring parentPath;
        if (! TryGetFileSystemParentPath(rename->endpoint.pathIdentity.value(), step.source.providerPath, parentPath))
        {
            return E_INVALIDARG;
        }
        const FileSystemRouteContract::ChildNameContractResult contract =
            FileSystemRouteContract::QueryChildNameContract(_fileSystem.get(), parentPath, step.finalLeafName, FILESYSTEM_RENAME, rename->endpoint.pluginId);
        if (contract.state != FileSystemRouteContract::QueryState::Available || contract.nameStatus != FILESYSTEM_CHILD_NAME_VALID)
        {
            const HRESULT hr = contract.state == FileSystemRouteContract::QueryState::Available && FAILED(contract.failureStatus)
                                   ? contract.failureStatus
                                   : (FAILED(contract.status) ? contract.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
            Debug::Perf::EmitValue(L"fileops.plan.rejection_bucket", static_cast<uint64_t>(FileOperations::PlanRejectionBucket::InvalidRename), hr);
            LogDiagnostic(DiagnosticSeverity::Error,
                          hr,
                          L"preparation.destinationName",
                          std::format(L"The folder cannot hold the name \"{}\"; nothing was renamed.", step.finalLeafName),
                          step.source.providerPath,
                          parentPath);
            return hr;
        }
        auto filled        = std::make_shared<FileOperations::FileOperationPlanGroup>(plans);
        auto* filledRename = std::get_if<FileOperations::RenamePlan>(&(*filled)[planIndex]);
        if (filledRename == nullptr)
        {
            return E_UNEXPECTED;
        }
        filledRename->finalMappings.front().providerJoinedPath   = contract.joinedPath;
        filledRename->finalMappings.front().providerCollisionKey = contract.collisionKey;
        StorePlans(std::move(filled));
    }
    for (const FileOperations::FileOperationPlan& plan : plans)
    {
        const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
        if (transfer == nullptr || ! transfer->sourceEndpoint.pathIdentity.has_value() || ! transfer->destinationEndpoint.pathIdentity.has_value())
        {
            continue;
        }
        for (size_t sourceIndex = 0u; sourceIndex < transfer->selectedItems.size(); ++sourceIndex)
        {
            if (_stopToken.stop_requested() || _cancelled.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            std::wstring destinationLeaf;
            if (! transfer->explicitMappings.empty())
            {
                const auto mapping = std::ranges::find_if(transfer->explicitMappings,
                                                          [sourceIndex](const FileOperations::TransferDestinationMapping& candidate) noexcept
                { return candidate.sourceIndex == sourceIndex; });
                if (mapping == transfer->explicitMappings.end() ||
                    ! TryGetFileSystemLeafName(transfer->destinationEndpoint.pathIdentity.value(), mapping->destinationProviderPath, destinationLeaf))
                {
                    return E_INVALIDARG;
                }
            }
            else if (! TryGetFileSystemLeafName(
                         transfer->sourceEndpoint.pathIdentity.value(), transfer->selectedItems[sourceIndex].providerPath, destinationLeaf))
            {
                return E_INVALIDARG;
            }
            const FileSystemRouteContract::ChildNameContractResult destinationName = FileSystemRouteContract::QueryChildNameContract(
                destinationFileSystem, transfer->destination.providerFolderPath, destinationLeaf, _operation, _destinationPluginId);
            if (destinationName.state == FileSystemRouteContract::QueryState::Available && destinationName.nameStatus == FILESYSTEM_CHILD_NAME_INVALID)
            {
                const HRESULT hr = FAILED(destinationName.failureStatus) ? destinationName.failureStatus : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                Debug::Perf::EmitValue(
                    L"fileops.plan.rejection_bucket", static_cast<uint64_t>(FileOperations::PlanRejectionBucket::InvalidDestinationName), hr);
                LogDiagnostic(DiagnosticSeverity::Error,
                              hr,
                              L"preparation.destinationName",
                              std::format(L"The destination cannot hold the name \"{}\"; nothing was copied or moved.", destinationLeaf),
                              transfer->selectedItems[sourceIndex].providerPath,
                              transfer->destination.providerFolderPath);
                return hr;
            }
        }
    }
    return S_OK;
}

HRESULT FolderWindow::FileOperationState::Task::PrepareForExecution() noexcept
{
    const uint64_t startedUs = PerfNowUs();
    if (_stopToken.stop_requested() || _cancelled.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (_pendingBatchRenameAdmission.has_value())
    {
        HRESULT admissionHr = PrepareBatchRenameAdmission();
        if (admissionHr == S_OK)
        {
            admissionHr = PrepareBatchRenameArtifactGuard();
        }
        if (admissionHr != S_OK)
        {
            return admissionHr == S_FALSE ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : admissionHr;
        }
    }

    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (! plans || plans->empty())
    {
        return E_UNEXPECTED;
    }
    for (const FileOperations::FileOperationPlan& plan : *plans)
    {
        FileOperations::PlanRejectionBucket rejection = FileOperations::PlanRejectionBucket::None;
        const HRESULT validationHr                    = FileOperations::ValidatePlan(plan, &rejection, _permanentDeleteConfirmationPending);
        if (FAILED(validationHr))
        {
            Debug::Perf::EmitValue(L"fileops.plan.rejection_bucket", static_cast<uint64_t>(rejection), validationHr);
            return validationHr;
        }
    }

#ifdef ENABLE_TESTS
    FolderWindowFileOperationsStateInternal::RecordPreparedTransferStrategyForSelfTest(*this);
#endif
    const HRESULT destinationNamesHr = PrepareTransferDestinationNames(*plans);
    if (FAILED(destinationNamesHr))
    {
        return destinationNamesHr;
    }
    if (_stopToken.stop_requested() || _cancelled.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    const HRESULT interlockHr = PrepareMutationInterlockScopes();
    if (FAILED(interlockHr))
    {
        return interlockHr;
    }
    // C1: a Permanent Delete confirms on its card once every selected root is pinned.
    const HRESULT confirmationHr = RunPermanentDeleteConfirmation();
    if (FAILED(confirmationHr))
    {
        return confirmationHr;
    }
    // Publish the immutable scopes before overlap comparison. The queue lock serializes this with
    // every peer's publication, so at least the later of two simultaneous preparations observes
    // the other even though neither task has entered operation yet.
    if (_state != nullptr)
    {
        _state->PublishPreparedMutationInterlock(*this);
    }
    // R4-A02-1: with the scopes known, warn once about an obvious same-host overlap before anything
    // is consumed or mutated; Don't start is a Preparing cancel.
    const HRESULT overlapHr = RunSameHostOverlapAdvisory();
    if (FAILED(overlapHr))
    {
        return overlapHr;
    }
    if (_stopToken.stop_requested() || _cancelled.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    PublishLifecyclePhase(TaskLifecyclePhase::AwaitingAcceptance);
    const HRESULT snapshotHr = BuildPreparationSnapshot(PerfElapsedUs(startedUs));
    if (FAILED(snapshotHr))
    {
        return snapshotHr;
    }
    const HRESULT decisionHr = RunPreConsumptionDecisionGate();
    if (FAILED(decisionHr))
    {
        return decisionHr;
    }
    if (_stopToken.stop_requested() || _cancelled.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return S_OK;
}

void FolderWindow::FileOperationState::Task::ThreadMain(std::stop_token stopToken) noexcept
{
    _stopToken = stopToken;
    // R0f: every synchronous Win32 call this worker makes on a provider's behalf has a quiet point.
    // Once Cancel/Stop is requested, the watch cancels a call still pending after its grace, so a
    // wedged network round trip returns instead of holding the task, its popup, or shutdown.
    const Common::SynchronousIoCancelWatch::Scope cancelWatch(&Task::ShouldCancelSynchronousIo, this);
    [[maybe_unused]] const std::stop_callback stopWake(stopToken,
                                                       [this]() noexcept
    {
        PublishLifecyclePhase(TaskLifecyclePhase::Stopping);
        HRESULT pendingReadiness = E_PENDING;
        if (_selectedRootReadinessStatus.compare_exchange_strong(pendingReadiness, HRESULT_FROM_WIN32(ERROR_CANCELLED), std::memory_order_acq_rel))
        {
            _selectedRootReadinessComplete.store(true, std::memory_order_release);
            _selectedRootReadinessComplete.notify_all();
        }
        _workerReleased.store(true, std::memory_order_release);
        _workerReleased.notify_all();
        WakePauseWaiters();
        {
            // Take the prompt mutex after the stop state is published so a worker between its
            // predicate check and its wait cannot miss this notification.
            std::scoped_lock promptLock(_batchRenameArtifactPromptMutex);
        }
        _batchRenameArtifactPromptCv.notify_all();
        if (_state)
        {
            ++_perf.queueNotifyAllCount;
            _state->NotifyQueueChanged();
        }
    });
    [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();

    const HRESULT preparedHr = PrepareForExecution();
    if (FAILED(preparedHr) && _state != nullptr)
    {
        _state->WithdrawPreparedMutationInterlock(*this);
    }
    HRESULT pendingReadiness = E_PENDING;
    if (_selectedRootReadinessStatus.compare_exchange_strong(pendingReadiness, preparedHr, std::memory_order_acq_rel))
    {
        _selectedRootReadinessComplete.store(true, std::memory_order_release);
        _selectedRootReadinessComplete.notify_all();
    }
    const HRESULT publishedReadinessHr = _selectedRootReadinessStatus.load(std::memory_order_acquire);
    if (_clipboardMoveAdmission)
    {
        if (FAILED(publishedReadinessHr))
        {
            CompleteClipboardMoveAdmission(publishedReadinessHr, false);
        }
        else if (stopToken.stop_requested() || _cancelled.load(std::memory_order_acquire))
        {
            CompleteClipboardMoveAdmission(HRESULT_FROM_WIN32(ERROR_CANCELLED), false);
        }
        else
        {
            auto payload = std::unique_ptr<ClipboardMoveReadyPayload>(new (std::nothrow) ClipboardMoveReadyPayload{});
            if (! payload)
            {
                CompleteClipboardMoveAdmission(E_OUTOFMEMORY, false);
            }
            else
            {
                payload->taskId = _taskId;
                if (! PostMessagePayload(_state->_owner.GetHwnd(), WndMsg::kFileOperationClipboardMoveReady, static_cast<WPARAM>(_taskId), std::move(payload)))
                {
                    CompleteClipboardMoveAdmission(HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS), false);
                }
            }
        }
    }

#ifdef ENABLE_TESTS
    _debugWorkerStartGateWaiting.store(true, std::memory_order_release);
#endif
    _workerReleased.wait(false, std::memory_order_acquire);
#ifdef ENABLE_TESTS
    _debugWorkerStartGateWaiting.store(false, std::memory_order_release);
#endif

    if (! _state)
    {
        return;
    }

    const auto completeAndPost = [&](HRESULT operationStatus) noexcept
    {
        PublishLifecyclePhase(TaskLifecyclePhase::Stopping);
        const HRESULT terminalHr = FinalizeTypedItemResults(operationStatus);
        _resultHr.store(terminalHr, std::memory_order_release);
        _state->PostCompleted(*this);
    };

    if (stopToken.stop_requested() || _cancelled.load(std::memory_order_acquire))
    {
        _state->WithdrawPreparedMutationInterlock(*this);
        completeAndPost(HRESULT_FROM_WIN32(ERROR_CANCELLED));
        return;
    }

    const HRESULT readinessHr = _selectedRootReadinessStatus.load(std::memory_order_acquire);
    if (FAILED(readinessHr))
    {
        _state->WithdrawPreparedMutationInterlock(*this);
        LogDiagnostic(IsCancellationStatus(readinessHr) ? DiagnosticSeverity::Info : DiagnosticSeverity::Error,
                      readinessHr,
                      L"preparation.failed",
                      IsCancellationStatus(readinessHr) ? L"Task preparation was canceled before any mutation was attempted."
                                                        : L"Task preparation failed before any mutation was attempted.");
        completeAndPost(readinessHr);
        return;
    }

    if (_clipboardMoveAdmission)
    {
        const HRESULT consumptionHr = _clipboardMoveConsumptionStatus.load(std::memory_order_acquire);
        const bool consumed         = _clipboardMoveConsumed.load(std::memory_order_acquire);
        if (! ClipboardMutationGateAllows(readinessHr, consumptionHr, consumed))
        {
            _state->WithdrawPreparedMutationInterlock(*this);
            completeAndPost(FAILED(readinessHr) ? readinessHr : consumptionHr);
            return;
        }
    }

    PublishLifecyclePhase(TaskLifecyclePhase::Ready);
    Debug::Perf::Emit(L"FileOps.Preparing.MutationBeforeReadyCount", OperationToString(_operation), 0u, 0u, 0u, S_OK);
    Debug::Perf::Emit(L"FileOps.Preparing.ClipboardBeforeReadyCount", OperationToString(_operation), 0u, 0u, 0u, S_OK);
    Debug::Perf::Emit(L"FileOps.Preparing.BreadcrumbBeforeReadyCount", OperationToString(_operation), 0u, 0u, 0u, S_OK);
    Debug::Perf::Emit(L"FileOps.Preparing.CompletionBeforeReadyCount", OperationToString(_operation), 0u, 0u, 0u, S_OK);

    const HRESULT breadcrumbCreateHr = CreateMoveBreadcrumbAfterAcceptance();
    if (FAILED(breadcrumbCreateHr))
    {
        _state->WithdrawPreparedMutationInterlock(*this);
        LogDiagnostic(DiagnosticSeverity::Error,
                      breadcrumbCreateHr,
                      L"move.breadcrumb.createFailed",
                      L"The accepted Move breadcrumb could not be persisted; no mutation was attempted.");
        completeAndPost(breadcrumbCreateHr);
        return;
    }

    LogDiagnostic(DiagnosticSeverity::Debug,
                  S_OK,
                  L"task.started",
                  std::format(L"Task started (op={}, mode={}, sources={}, flags=0x{:08X}, discoveryAhead=true, waitForOthers={}).",
                              OperationToString(_operation),
                              _executionMode == ExecutionMode::PerItem ? L"perItem" : L"bulkItems",
                              _sourcePaths.size(),
                              static_cast<unsigned long>(static_cast<uint32_t>(_flags)),
                              _waitForOthers.load(std::memory_order_acquire) ? L"true" : L"false"));

    // Mark as waiting in queue before entering (visible to UI while blocked). Use the current
    // desired start-gating state to avoid briefly showing "Waiting" for tasks that will start immediately.
    PublishLifecyclePhase(TaskLifecyclePhase::Waiting);
    SetWaitingInQueue(_waitForOthers.load(std::memory_order_acquire));

    // Enter the queue before starting the shared discovery/execution traversal.
    const bool canStart = _state->EnterOperation(*this, stopToken);

    // No longer waiting in queue (either we got our turn or were cancelled)
    SetWaitingInQueue(false);

    if (! canStart)
    {
        completeAndPost(HRESULT_FROM_WIN32(ERROR_CANCELLED));
        return;
    }

    _enteredOperationTick.store(GetTickCount64(), std::memory_order_release);
    _enteredOperation.store(true, std::memory_order_release);
    PublishLifecyclePhase(TaskLifecyclePhase::Running);

    if (_moveBreadcrumb.has_value())
    {
        const HRESULT breadcrumbHr = _moveBreadcrumb->Advance(FileOperationMoveBreadcrumb::DurablePhase::Executing);
        if (FAILED(breadcrumbHr))
        {
            // The existing admitted record remains conservative: after a crash it still says the
            // Move may have started. Breadcrumb persistence never becomes mutation authority.
            LogDiagnostic(DiagnosticSeverity::Warning,
                          breadcrumbHr,
                          L"move.breadcrumb.advanceFailed",
                          L"The Move breadcrumb could not advance to Executing; the admitted record remains conservative.");
        }
    }

    if (_cancelled.load(std::memory_order_acquire))
    {
        _enteredOperation.store(false, std::memory_order_release);
        _enteredOperationTick.store(0, std::memory_order_release);
        _state->LeaveOperation(*this);
        completeAndPost(HRESULT_FROM_WIN32(ERROR_CANCELLED));
        return;
    }

    // Discovery and mutation now share one bounded traversal. Crossing this point means provider
    // mutation may become observable, so fallback truth must become conservative if no exact item
    // receipt arrives.
    MarkSourceItemsMutationPossible();
    BeginDiscovery();
    if (_cancelled.load(std::memory_order_acquire))
    {
        CloseDiscovery();
        _enteredOperation.store(false, std::memory_order_release);
        _enteredOperationTick.store(0, std::memory_order_release);
        _state->LeaveOperation(*this);
        completeAndPost(HRESULT_FROM_WIN32(ERROR_CANCELLED));
        return;
    }

    const HRESULT executionHr = ExecuteOperation();
    CloseDiscovery();
    // RequestCancel records the timestamp before publishing _cancelled. Treat that durable
    // timestamp as the request's linearization point so a fast provider completion cannot
    // slip through the few instructions between those two atomic stores as a false success.
    const bool cancellationRequested = CancelRequestedDurable();
    // A synchronous call the R0f cancel watch aborted fails with ERROR_OPERATION_ABORTED; under a
    // requested cancel that is the cancellation itself, not a provider fault. Items that carry
    // that status are mapped the same way by FinalizeTypedItemResults, whatever the execution
    // path returned around them.
    const bool abortedByCancelWatch = cancellationRequested && executionHr == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED);
    const HRESULT terminalExecutionHr =
        (SUCCEEDED(executionHr) || abortedByCancelWatch) && cancellationRequested ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : executionHr;
    PublishLifecyclePhase(TaskLifecyclePhase::Stopping);
    const HRESULT hr = FinalizeTypedItemResults(terminalExecutionHr);
    _resultHr.store(hr, std::memory_order_release);
    if (FAILED(hr))
    {
        const PublishedProgressSnapshot progressSnapshot = LoadPublishedProgressSnapshot(*this);
        std::wstring sourcePath;
        std::wstring destinationPath;
        {
            std::scoped_lock lock(_progressPathMutex);
            CopyEffectiveProgressPathsLocked(*this, sourcePath, destinationPath);
        }

        const HRESULT partialCopyHr       = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        const DiagnosticSeverity severity = (hr == partialCopyHr)      ? DiagnosticSeverity::Warning
                                            : IsCancellationStatus(hr) ? DiagnosticSeverity::Info
                                                                       : DiagnosticSeverity::Error;
        std::wstring message;
        if (hr == partialCopyHr)
        {
            if (_operation == FILESYSTEM_MOVE)
            {
                message = std::format(L"Move completed partially: source preserved; partial copy left at destination (items={:L}/{:L}, bytes={:L}/{:L}).",
                                      progressSnapshot.completedItems,
                                      progressSnapshot.totalItems,
                                      progressSnapshot.completedBytes,
                                      progressSnapshot.totalBytes);
            }
            else
            {
                message = std::format(L"Task completed with skipped or partial items (op={}, items={:L}/{:L}, bytes={:L}/{:L}).",
                                      OperationToString(_operation),
                                      progressSnapshot.completedItems,
                                      progressSnapshot.totalItems,
                                      progressSnapshot.completedBytes,
                                      progressSnapshot.totalBytes);
            }
        }
        else if (IsCancellationStatus(hr))
        {
            message = std::format(L"Task was canceled (op={}, items={:L}/{:L}, bytes={:L}/{:L}).",
                                  OperationToString(_operation),
                                  progressSnapshot.completedItems,
                                  progressSnapshot.totalItems,
                                  progressSnapshot.completedBytes,
                                  progressSnapshot.totalBytes);
        }
        else
        {
            const std::wstring statusText = FormatDiagnosticStatusText(hr);
            message                       = std::format(L"Task failed (op={}, hr=0x{:08X}, status='{}', items={:L}/{:L}, bytes={:L}/{:L}).",
                                                        OperationToString(_operation),
                                                        static_cast<unsigned long>(hr),
                                                        statusText,
                                                        progressSnapshot.completedItems,
                                                        progressSnapshot.totalItems,
                                                        progressSnapshot.completedBytes,
                                                        progressSnapshot.totalBytes);
        }
        LogDiagnostic(severity, hr, L"task.result", message, sourcePath, destinationPath);
    }

    {
        const ULONGLONG opStartTick = _operationStartTick.load(std::memory_order_acquire);
        const ULONGLONG endTick     = GetTickCount64();
        const ULONGLONG elapsedMs   = (opStartTick > 0 && endTick >= opStartTick) ? (endTick - opStartTick) : 0;

        const PublishedProgressSnapshot progressSnapshot = LoadPublishedProgressSnapshot(*this);
        std::wstring sourcePath;
        std::wstring destinationPath;
        {
            std::scoped_lock lock(_progressPathMutex);
            CopyEffectiveProgressPathsLocked(*this, sourcePath, destinationPath);
        }

        LogDiagnostic(DiagnosticSeverity::Debug,
                      hr,
                      L"task.operation.result",
                      std::format(L"Operation finished (hr=0x{:08X}, elapsedMs={}, items={:L}/{:L}, bytes={:L}/{:L}, progressCalls={:L}, itemCalls={:L}).",
                                  static_cast<unsigned long>(hr),
                                  elapsedMs,
                                  progressSnapshot.completedItems,
                                  progressSnapshot.totalItems,
                                  progressSnapshot.completedBytes,
                                  progressSnapshot.totalBytes,
                                  progressSnapshot.progressCallbackCount,
                                  progressSnapshot.itemCompletedCallbackCount),
                      sourcePath,
                      destinationPath);
    }

    if (Debug::Perf::IsCaptureEnabled())
    {
        const ULONGLONG opStartTick = _operationStartTick.load(std::memory_order_acquire);
        const ULONGLONG endTick     = GetTickCount64();
        const ULONGLONG elapsedMs   = (opStartTick > 0 && endTick >= opStartTick) ? (endTick - opStartTick) : 0;
        const uint64_t durationUs   = static_cast<uint64_t>(elapsedMs) * 1000ull;

        const PublishedProgressSnapshot progressSnapshot           = LoadPublishedProgressSnapshot(*this);
        const auto& perfStats                                      = _perf;
        const uint64_t bridgeDirectoryEnsureCount                  = _bridgeDirectoryEnsureCount.load(std::memory_order_acquire);
        const uint64_t bridgeFileAdmissionCount                    = _bridgeFileAdmissionCount.load(std::memory_order_acquire);
        const uint64_t bridgeFileStartedBeforeProducerDone         = _bridgeFileStartedBeforeProducerDone.load(std::memory_order_acquire);
        const uint64_t bridgeAdmissionMaxQueueDepth                = _bridgeAdmissionMaxQueueDepth.load(std::memory_order_acquire);
        const uint64_t bridgeTraversalMaxDepth                     = _bridgeTraversalMaxDepth.load(std::memory_order_acquire);
        const uint64_t bridgeTraversalMaxRetainedEntries           = _bridgeTraversalMaxRetainedEntries.load(std::memory_order_acquire);
        const uint64_t bridgeTraversalMaxQueuedPathBytes           = _bridgeTraversalMaxQueuedPathBytes.load(std::memory_order_acquire);
        const uint64_t bridgeTraversalMaxMetadataBytes             = _bridgeTraversalMaxMetadataBytes.load(std::memory_order_acquire);
        const uint64_t bridgeTraversalLimitHitCount                = _bridgeTraversalLimitHitCount.load(std::memory_order_acquire);
        const uint64_t permanentDeleteIdentityMismatchCount        = _permanentDeleteIdentityMismatchCount.load(std::memory_order_acquire);
        const uint64_t permanentDeleteConditionalAttemptCount      = _permanentDeleteConditionalAttemptCount.load(std::memory_order_acquire);
        const uint64_t permanentDeleteRemovedCount                 = _permanentDeleteRemovedCount.load(std::memory_order_acquire);
        const uint64_t permanentDeleteIndeterminateCount           = _permanentDeleteIndeterminateCount.load(std::memory_order_acquire);
        const uint64_t conflictExpectedDestinationBoundCount       = _conflictExpectedDestinationBoundCount.load(std::memory_order_acquire);
        const uint64_t conflictExpectedDestinationUnavailableCount = _conflictExpectedDestinationUnavailableCount.load(std::memory_order_acquire);
        const uint64_t conflictExpectedDestinationReturnedCount    = _conflictExpectedDestinationReturnedCount.load(std::memory_order_acquire);
        const uint64_t discoveryOpenUs                             = perfStats.discoveryOpenUs.load(std::memory_order_acquire);
        const uint64_t discoveryCallbackCount                      = perfStats.discoveryCallbackCount.load(std::memory_order_acquire);
        const uint64_t discoveryCallbackUs                         = perfStats.discoveryCallbackUs.load(std::memory_order_acquire);
        const uint64_t discoveryLockWaitUs                         = perfStats.discoveryLockWaitUs.load(std::memory_order_acquire);
        const ULONGLONG discoverySkipRequestedTick                 = _discoverySkipRequestedTick.load(std::memory_order_acquire);
        const ULONGLONG discoveryReservationReleasedTick           = _discoveryReservationReleasedTick.load(std::memory_order_acquire);
        const uint64_t discoverySkipReleaseUs = discoverySkipRequestedTick != 0u && discoveryReservationReleasedTick >= discoverySkipRequestedTick
                                                    ? (discoveryReservationReleasedTick - discoverySkipRequestedTick) * 1000ull
                                                    : 0u;
        std::array<ProgressStreamPerf, kMaxInFlightFiles> progressStreamPerf{};
        size_t progressStreamPerfCount = 0;
        {
            std::scoped_lock lock(_progressStreamPerfMutex);
            progressStreamPerf      = _progressStreamPerf;
            progressStreamPerfCount = _progressStreamPerfCount;
        }
        uint64_t progressStreamGapCount              = 0;
        uint64_t progressStreamGapMs                 = 0;
        uint64_t progressStreamGapBytes              = 0;
        uint64_t progressStreamMaxGapMs              = 0;
        uint64_t progressStreamMaxGapBytes           = 0;
        uint64_t progressStreamMaxCallbackDeltaBytes = 0;
        for (size_t i = 0; i < progressStreamPerfCount; ++i)
        {
            progressStreamGapCount += progressStreamPerf[i].callbackGapCount;
            progressStreamGapMs += progressStreamPerf[i].callbackGapMs;
            progressStreamGapBytes += progressStreamPerf[i].callbackGapBytes;
            if (progressStreamPerf[i].maxCallbackGapMs > progressStreamMaxGapMs)
            {
                progressStreamMaxGapMs    = progressStreamPerf[i].maxCallbackGapMs;
                progressStreamMaxGapBytes = progressStreamPerf[i].maxCallbackGapBytes;
            }
            else if (progressStreamPerf[i].maxCallbackGapMs == progressStreamMaxGapMs)
            {
                progressStreamMaxGapBytes = std::max(progressStreamMaxGapBytes, progressStreamPerf[i].maxCallbackGapBytes);
            }
            progressStreamMaxCallbackDeltaBytes = (std::max)(progressStreamMaxCallbackDeltaBytes, progressStreamPerf[i].maxCallbackDeltaBytes);
        }

        std::array<ConflictWorkerPerf, kMaxInFlightFiles> conflictWorkerPerf{};
        size_t conflictWorkerPerfCount = 0;
        {
            std::scoped_lock lock(_conflictArbiter.mutex);
            conflictWorkerPerf      = _conflictWorkerPerf;
            conflictWorkerPerfCount = _conflictWorkerPerfCount;
        }

        const uint64_t desired   = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        const uint64_t effective = _effectiveSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

        const size_t sourceCount  = _sourcePaths.size();
        const std::wstring detail = std::format(
            L"id={} op={} desired={} effective={} sources={} items={} queueWaitUs={} schedulerWaitUs={} schedulerWorkUs={} bridgeCopyUs={} bridgeReadUs={} "
            L"bridgeWriteUs={} bridgeReaderWaitUs={} bridgeWriterWaitUs={} bridgeStageCreateUs={} bridgeStageCreateCount={} "
            L"bridgePublicationUs={} bridgePublicationCount={} bridgeStageRetainedCount={} bridgeImmediateDestinationSizeProbeUs={} "
            L"bridgeImmediateDestinationSizeProbeCount={} bridgeCommitSizeProofCount={} bridgeCommitSizeProofFallbackCount={} "
            L"discoveryOpenUs={} discoveryQueueMax={} discoveryStarvation={} progressUs={} firstProgressMs={} maxProgressGapMs={} "
            L"maxProgressGapBytes={} bridgeDirs={} bridgeFiles={} bridgeEarlyFiles={} bridgeQueueMax={} bridgeDepthMax={} "
            L"bridgeRetainedEntriesMax={} bridgeQueuedPathBytesMax={} bridgeMetadataBytesMax={} bridgeTraversalLimitHits={} "
            L"conflictExpectedDestinationBound={} conflictExpectedDestinationUnavailable={} "
            L"conflictExpectedDestinationReturned={} "
            L"itemCompletedUs={} conflictWaitUs={} conflictMetadataUs={} pauseWaitUs={}",
            _taskId,
            OperationToString(_operation),
            desired,
            effective,
            sourceCount,
            progressSnapshot.completedItems,
            perfStats.queueWaitUs,
            perfStats.schedulerWaitUs.load(std::memory_order_acquire),
            perfStats.schedulerWaitForWorkUs.load(std::memory_order_acquire),
            perfStats.bridgeCopyUs.load(std::memory_order_acquire),
            perfStats.bridgeReadUs.load(std::memory_order_acquire),
            perfStats.bridgeWriteUs.load(std::memory_order_acquire),
            perfStats.bridgeReaderWaitUs.load(std::memory_order_acquire),
            perfStats.bridgeWriterWaitUs.load(std::memory_order_acquire),
            perfStats.bridgeStageCreateUs.load(std::memory_order_acquire),
            perfStats.bridgeStageCreateCount.load(std::memory_order_acquire),
            perfStats.bridgePublicationUs.load(std::memory_order_acquire),
            perfStats.bridgePublicationCount.load(std::memory_order_acquire),
            perfStats.bridgeStageRetainedCount.load(std::memory_order_acquire),
            perfStats.bridgeImmediateDestinationSizeProbeUs.load(std::memory_order_acquire),
            perfStats.bridgeImmediateDestinationSizeProbeCount.load(std::memory_order_acquire),
            perfStats.bridgeCommitSizeProofCount.load(std::memory_order_acquire),
            perfStats.bridgeCommitSizeProofFallbackCount.load(std::memory_order_acquire),
            discoveryOpenUs,
            _discoveryMaxQueueDepth.load(std::memory_order_acquire),
            _discoveryStarvationCount.load(std::memory_order_acquire),
            perfStats.progressCallbackUs.load(std::memory_order_acquire),
            perfStats.progressFirstCallbackDelayMs,
            progressStreamMaxGapMs,
            progressStreamMaxGapBytes,
            bridgeDirectoryEnsureCount,
            bridgeFileAdmissionCount,
            bridgeFileStartedBeforeProducerDone,
            bridgeAdmissionMaxQueueDepth,
            bridgeTraversalMaxDepth,
            bridgeTraversalMaxRetainedEntries,
            bridgeTraversalMaxQueuedPathBytes,
            bridgeTraversalMaxMetadataBytes,
            bridgeTraversalLimitHitCount,
            conflictExpectedDestinationBoundCount,
            conflictExpectedDestinationUnavailableCount,
            conflictExpectedDestinationReturnedCount,
            perfStats.itemCompletedCallbackUs,
            perfStats.conflictWaitUs,
            perfStats.conflictMetadataUs.load(std::memory_order_acquire),
            perfStats.pauseWaitUs.load(std::memory_order_acquire));
        Debug::Perf::Emit(L"FileOps.Operation", detail, durationUs, progressSnapshot.completedBytes, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.CopyUs",
                          L"",
                          perfStats.bridgeCopyUs.load(std::memory_order_acquire),
                          progressSnapshot.completedBytes,
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.ReadUs",
                          L"",
                          perfStats.bridgeReadUs.load(std::memory_order_acquire),
                          progressSnapshot.completedBytes,
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.WriteUs",
                          L"",
                          perfStats.bridgeWriteUs.load(std::memory_order_acquire),
                          progressSnapshot.completedBytes,
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.ReaderWaitUs",
                          L"",
                          perfStats.bridgeReaderWaitUs.load(std::memory_order_acquire),
                          progressSnapshot.completedBytes,
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.WriterWaitUs",
                          L"",
                          perfStats.bridgeWriterWaitUs.load(std::memory_order_acquire),
                          progressSnapshot.completedBytes,
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.StageCreateUs",
                          L"",
                          perfStats.bridgeStageCreateUs.load(std::memory_order_acquire),
                          perfStats.bridgeStageCreateCount.load(std::memory_order_acquire),
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.PublicationUs",
                          L"",
                          perfStats.bridgePublicationUs.load(std::memory_order_acquire),
                          perfStats.bridgePublicationCount.load(std::memory_order_acquire),
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.StageRetainedCount",
                          L"",
                          0u,
                          perfStats.bridgeStageRetainedCount.load(std::memory_order_acquire),
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.ImmediateDestinationSizeProbeCount",
                          L"",
                          0u,
                          perfStats.bridgeImmediateDestinationSizeProbeCount.load(std::memory_order_acquire),
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.ImmediateDestinationSizeProbeUs",
                          L"",
                          perfStats.bridgeImmediateDestinationSizeProbeUs.load(std::memory_order_acquire),
                          perfStats.bridgeImmediateDestinationSizeProbeCount.load(std::memory_order_acquire),
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.CommitSizeProofCount",
                          L"",
                          0u,
                          perfStats.bridgeCommitSizeProofCount.load(std::memory_order_acquire),
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.CommitSizeProofFallbackCount",
                          L"",
                          0u,
                          perfStats.bridgeCommitSizeProofFallbackCount.load(std::memory_order_acquire),
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.DirectoryEnsureCount", L"", 0u, bridgeDirectoryEnsureCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.FileAdmissionCount", L"", 0u, bridgeFileAdmissionCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.FileStartedBeforeProducerDone", L"", 0u, bridgeFileStartedBeforeProducerDone, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.AdmissionMaxQueueDepth", L"", 0u, bridgeAdmissionMaxQueueDepth, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.TraversalMaxDepth", L"", 0u, bridgeTraversalMaxDepth, GetFileOpsTraversalMaxDepth(), hr);
        Debug::Perf::Emit(
            L"FileOps.Bridge.TraversalMaxRetainedEntries", L"", 0u, bridgeTraversalMaxRetainedEntries, Common::FileOperations::kTraversalMaxQueuedEntries, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.TraversalMaxQueuedPathBytes",
                          L"",
                          0u,
                          bridgeTraversalMaxQueuedPathBytes,
                          Common::FileOperations::kTraversalMaxQueuedPathBytes,
                          hr);
        Debug::Perf::Emit(L"FileOps.Bridge.TraversalMaxMetadataBytes", L"", 0u, bridgeTraversalMaxMetadataBytes, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.TraversalLimitHitCount", L"", 0u, bridgeTraversalLimitHitCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.PermanentDelete.IdentityMismatchCount", L"", 0u, permanentDeleteIdentityMismatchCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.PermanentDelete.ConditionalAttemptCount", L"", 0u, permanentDeleteConditionalAttemptCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.PermanentDelete.RemovedCount", L"", 0u, permanentDeleteRemovedCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.PermanentDelete.IndeterminateCount", L"", 0u, permanentDeleteIndeterminateCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Conflict.ExpectedDestinationBoundCount", L"", 0u, conflictExpectedDestinationBoundCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Conflict.ExpectedDestinationUnavailableCount", L"", 0u, conflictExpectedDestinationUnavailableCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Conflict.ExpectedDestinationReturnedCount", L"", 0u, conflictExpectedDestinationReturnedCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Discovery.OpenUs", L"", discoveryOpenUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(L"FileOps.Discovery.CallbackCount", L"", 0u, discoveryCallbackCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Discovery.CallbackUs", L"", discoveryCallbackUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(L"FileOps.Discovery.LockWaitUs", L"", discoveryLockWaitUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(L"FileOps.Discovery.MaxQueueDepth",
                          L"",
                          0u,
                          _discoveryMaxQueueDepth.load(std::memory_order_acquire),
                          Common::FileOperations::kDiscoveryMaxTarget,
                          hr);
        Debug::Perf::Emit(L"FileOps.Discovery.StarvationCount", L"", 0u, _discoveryStarvationCount.load(std::memory_order_acquire), 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.Discovery.FirstMutationBeforeClose", L"", 0u, _firstMutationBeforeDiscoveryClosed.load(std::memory_order_acquire) ? 1u : 0u, 1u, hr);
        Debug::Perf::Emit(L"FileOps.Discovery.FirstMutationUs",
                          L"",
                          _discoveryFirstMutationUs.load(std::memory_order_acquire),
                          _discoveryCompletedBytesWhileOpen.load(std::memory_order_acquire),
                          _discoveryCompletedMutationsWhileOpen.load(std::memory_order_acquire),
                          hr);
        Debug::Perf::Emit(L"FileOps.Discovery.BytesCompletedWhileOpen",
                          L"",
                          0u,
                          _discoveryCompletedBytesWhileOpen.load(std::memory_order_acquire),
                          progressSnapshot.completedBytes,
                          hr);
        Debug::Perf::Emit(L"FileOps.Discovery.MutationsCompletedWhileOpen",
                          L"",
                          0u,
                          _discoveryCompletedMutationsWhileOpen.load(std::memory_order_acquire),
                          progressSnapshot.completedItems,
                          hr);
        Debug::Perf::Emit(
            L"FileOps.Discovery.SkipReleaseUs", L"", discoverySkipReleaseUs, _discoverySkipped.load(std::memory_order_acquire) ? 1u : 0u, 50'000u, hr);
        Debug::Perf::Emit(L"FileOps.Discovery.Closed", L"", 0u, _discoveryClosed.load(std::memory_order_acquire) ? 1u : 0u, 1u, hr);
        Debug::Perf::Emit(L"FileOps.Progress.CallbackUs",
                          L"",
                          perfStats.progressCallbackUs.load(std::memory_order_acquire),
                          progressSnapshot.completedBytes,
                          progressSnapshot.progressCallbackCount,
                          hr);
        Debug::Perf::Emit(L"FileOps.Progress.FirstCallbackDelayMs",
                          L"",
                          perfStats.progressFirstCallbackDelayMs,
                          progressSnapshot.completedBytes,
                          progressSnapshot.progressCallbackCount,
                          hr);
        Debug::Perf::Emit(
            L"FileOps.Progress.MaxCallbackGapMs", L"", progressStreamMaxGapMs, progressStreamGapCount, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Progress.CallbackGapMs", L"", progressStreamGapMs, progressStreamGapCount, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(
            L"FileOps.Progress.MaxCallbackGapBytes", L"", progressStreamMaxGapBytes, progressStreamMaxGapMs, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Progress.MaxCallbackDeltaBytes",
                          L"",
                          progressStreamMaxCallbackDeltaBytes,
                          progressStreamPerfCount,
                          progressSnapshot.progressCallbackCount,
                          hr);
        Debug::Perf::Emit(
            L"FileOps.Progress.CallbackGapBytes", L"", progressStreamGapBytes, progressStreamGapCount, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(
            L"FileOps.Progress.LockWaitUs", L"", perfStats.progressLockWaitUs, progressSnapshot.completedBytes, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(
            L"FileOps.Progress.LockHoldUs", L"", perfStats.progressLockHoldUs, progressSnapshot.completedBytes, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Progress.LockContentionCount", L"", 0u, perfStats.progressLockContentionCount, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Progress.PathUpdateBytes", L"", 0u, perfStats.progressPathUpdateBytes, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(
            L"FileOps.Progress.PathUpdateAppliedCount", L"", 0u, perfStats.progressPathUpdateAppliedCount, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(
            L"FileOps.Progress.PathUpdateSkippedCount", L"", 0u, perfStats.progressPathUpdateSkippedCount, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(
            L"FileOps.Progress.PathUpdateThrottledCount", L"", 0u, perfStats.progressPathUpdateThrottledCount, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Progress.InFlightEvictions", L"", 0u, perfStats.progressInFlightEvictions, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Progress.PerItemInFlightEvictions", L"", 0u, perfStats.perItemInFlightEvictions, 0u, hr);
        Debug::Perf::Emit(L"FileOps.ItemCompleted.CallbackUs",
                          L"",
                          perfStats.itemCompletedCallbackUs,
                          progressSnapshot.completedBytes,
                          progressSnapshot.itemCompletedCallbackCount,
                          hr);
        Debug::Perf::Emit(L"FileOps.ItemCompleted.LockWaitUs",
                          L"",
                          perfStats.itemCompletedLockWaitUs,
                          progressSnapshot.completedBytes,
                          progressSnapshot.itemCompletedCallbackCount,
                          hr);
        Debug::Perf::Emit(L"FileOps.ItemCompleted.LockHoldUs",
                          L"",
                          perfStats.itemCompletedLockHoldUs,
                          progressSnapshot.completedBytes,
                          progressSnapshot.itemCompletedCallbackCount,
                          hr);
        Debug::Perf::Emit(
            L"FileOps.ItemCompleted.LockContentionCount", L"", 0u, perfStats.itemCompletedLockContentionCount, progressSnapshot.itemCompletedCallbackCount, hr);
        Debug::Perf::Emit(
            L"FileOps.ItemCompleted.PathUpdateBytes", L"", 0u, perfStats.itemCompletedPathUpdateBytes, progressSnapshot.itemCompletedCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.ItemCompleted.PathUpdateAppliedCount",
                          L"",
                          0u,
                          perfStats.itemCompletedPathUpdateAppliedCount,
                          progressSnapshot.itemCompletedCallbackCount,
                          hr);
        Debug::Perf::Emit(L"FileOps.ItemCompleted.PathUpdateSkippedCount",
                          L"",
                          0u,
                          perfStats.itemCompletedPathUpdateSkippedCount,
                          progressSnapshot.itemCompletedCallbackCount,
                          hr);
        Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", perfStats.queueWaitUs, progressSnapshot.completedBytes, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Queue.EnterCount", L"", 0u, perfStats.queueEnterCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Queue.NotifyAllCount", L"", 0u, perfStats.queueNotifyAllCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Queue.CancelWhileWaiting", L"", 0u, perfStats.queueCancelWhileWaiting, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Queue.DepthOnEnter", L"", 0u, perfStats.queueDepthOnEnter, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Queue.ActiveOperations", L"", 0u, perfStats.queueActiveOperations, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Scheduler.WaitForWorkUs", L"", perfStats.schedulerWaitForWorkUs.load(std::memory_order_acquire), 0u, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Scheduler.ProcessIndexUs", L"", perfStats.schedulerProcessIndexUs.load(std::memory_order_acquire), 0u, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Scheduler.DequeueAttempts", L"", 0u, perfStats.schedulerDequeueAttempts.load(std::memory_order_acquire), 0u, hr);
        Debug::Perf::Emit(L"FileOps.Scheduler.DequeueSuccess", L"", 0u, perfStats.schedulerDequeueSuccess.load(std::memory_order_acquire), 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.Conflict.WaitUs", L"", perfStats.conflictWaitUs, progressSnapshot.completedBytes, progressSnapshot.itemCompletedCallbackCount, hr);
        Debug::Perf::Emit(
            L"FileOps.Conflict.MetadataUs", L"", perfStats.conflictMetadataUs.load(std::memory_order_acquire), perfStats.conflictPromptCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Conflict.ConvergenceWaitUs",
                          L"",
                          perfStats.conflictConvergenceWaitUs,
                          progressSnapshot.completedBytes,
                          progressSnapshot.itemCompletedCallbackCount,
                          hr);
        Debug::Perf::Emit(L"FileOps.Conflict.PromptCount", L"", 0u, perfStats.conflictPromptCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Consent.WaitUs", L"", perfStats.consentWaitUs, progressSnapshot.completedBytes, perfStats.consentPromptCount, hr);
        Debug::Perf::Emit(L"FileOps.Consent.PromptCount", L"", 0u, perfStats.consentPromptCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Pause.WaitUs",
                          L"",
                          perfStats.pauseWaitUs.load(std::memory_order_acquire),
                          progressSnapshot.completedBytes,
                          progressSnapshot.itemCompletedCallbackCount,
                          hr);

        for (size_t i = 0; i < progressStreamPerfCount; ++i)
        {
            const auto& entry = progressStreamPerf[i];
            if (entry.callbackCount == 0 && entry.callbackUs == 0 && entry.lockWaitUs == 0)
            {
                continue;
            }

            const std::wstring streamDetail = std::format(L"id={} op={} stream={} cookie=0x{:X}",
                                                          _taskId,
                                                          OperationToString(_operation),
                                                          entry.progressStreamId,
                                                          static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(entry.cookieKey)));
            Debug::Perf::Emit(L"FileOps.Progress.Stream.CallbackUs", streamDetail, entry.callbackUs, entry.callbackCount, entry.progressStreamId, hr);
            Debug::Perf::Emit(L"FileOps.Progress.Stream.LockWaitUs", streamDetail, entry.lockWaitUs, entry.callbackCount, entry.progressStreamId, hr);
            Debug::Perf::Emit(
                L"FileOps.Progress.Stream.MaxCallbackGapMs", streamDetail, entry.maxCallbackGapMs, entry.callbackGapCount, entry.progressStreamId, hr);
            Debug::Perf::Emit(L"FileOps.Progress.Stream.CallbackGapMs", streamDetail, entry.callbackGapMs, entry.callbackGapCount, entry.progressStreamId, hr);
            Debug::Perf::Emit(
                L"FileOps.Progress.Stream.MaxCallbackGapBytes", streamDetail, entry.maxCallbackGapBytes, entry.maxCallbackGapMs, entry.progressStreamId, hr);
            Debug::Perf::Emit(
                L"FileOps.Progress.Stream.CallbackGapBytes", streamDetail, entry.callbackGapBytes, entry.callbackGapCount, entry.progressStreamId, hr);
        }

        for (size_t i = 0; i < conflictWorkerPerfCount; ++i)
        {
            const auto& entry = conflictWorkerPerf[i];
            if (entry.promptCount == 0 && entry.waitUs == 0)
            {
                continue;
            }

            const std::wstring workerDetail = std::format(L"id={} op={} cookie=0x{:X}",
                                                          _taskId,
                                                          OperationToString(_operation),
                                                          static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(entry.cookieKey)));
            Debug::Perf::Emit(L"FileOps.Conflict.Worker.WaitUs",
                              workerDetail,
                              entry.waitUs,
                              entry.promptCount,
                              static_cast<uint64_t>(reinterpret_cast<uintptr_t>(entry.cookieKey)),
                              hr);
        }

        const ULONGLONG cancelTick = _cancelRequestedTick.load(std::memory_order_acquire);
        if (cancelTick > 0)
        {
            const ULONGLONG cancelMs = (endTick >= cancelTick) ? (endTick - cancelTick) : 0;
            const uint64_t cancelUs  = static_cast<uint64_t>(cancelMs) * 1000ull;
            Debug::Perf::Emit(L"FileOps.CancelLatency", detail, cancelUs, progressSnapshot.completedBytes, progressSnapshot.itemCompletedCallbackCount, hr);
        }
    }

    _enteredOperation.store(false, std::memory_order_release);
    _enteredOperationTick.store(0, std::memory_order_release);
    _state->LeaveOperation(*this);
    _state->PostCompleted(*this);
}

bool FolderWindow::FileOperationState::Task::ShouldCancelSynchronousIo(void* context) noexcept
{
    const auto* task = static_cast<const Task*>(context);
    return task != nullptr && (task->_cancelled.load(std::memory_order_acquire) || task->_stopToken.stop_requested());
}

void FolderWindow::FileOperationState::Task::RequestCancel() noexcept
{
    PublishLifecyclePhase(TaskLifecyclePhase::Stopping);
    {
        ULONGLONG expected = 0;
        _cancelRequestedTick.compare_exchange_strong(expected, GetTickCount64(), std::memory_order_release);
    }
    _cancelled.store(true, std::memory_order_release);
    HRESULT pendingReadiness = E_PENDING;
    if (_selectedRootReadinessStatus.compare_exchange_strong(pendingReadiness, HRESULT_FROM_WIN32(ERROR_CANCELLED), std::memory_order_acq_rel))
    {
        _selectedRootReadinessComplete.store(true, std::memory_order_release);
        _selectedRootReadinessComplete.notify_all();
    }
    _workerReleased.store(true, std::memory_order_release);
    _workerReleased.notify_all();
    {
        std::scoped_lock lock(_pauseMutex);
        _paused.store(false, std::memory_order_release);
    }
    {
        // Take the prompt mutex after _cancelled is published so a worker between its predicate
        // check and its wait cannot miss this notification.
        std::scoped_lock promptLock(_batchRenameArtifactPromptMutex);
    }
    _batchRenameArtifactPromptCv.notify_all();

    if (_conflictArbiter.decisionEvent)
    {
        static_cast<void>(SetEvent(_conflictArbiter.decisionEvent.get()));
    }

    WakePauseWaiters();

    if (_state)
    {
        ++_perf.queueNotifyAllCount;
        _state->NotifyQueueChanged();
    }

    GetPerItemTaskScheduler().NotifyWorkAvailable();
}

void FolderWindow::FileOperationState::Task::SetPaused(bool paused) noexcept
{
    {
        std::scoped_lock lock(_pauseMutex);
        const bool wasPaused = _paused.load(std::memory_order_relaxed);
        if (wasPaused == paused)
        {
            return;
        }

        _paused.store(paused, std::memory_order_release);
    }

    MarkRateSamplingStateChanged();
    if (! paused)
    {
        _pauseCv.notify_all();
    }

    GetPerItemTaskScheduler().NotifyWorkAvailable();
}

void FolderWindow::FileOperationState::Task::TogglePause() noexcept
{
    SetPaused(! _paused.load(std::memory_order_acquire));
}

void FolderWindow::FileOperationState::Task::SetDesiredSpeedLimit(uint64_t bytesPerSecond) noexcept
{
    _desiredSpeedLimitBytesPerSecond.store(bytesPerSecond, std::memory_order_release);
}

void FolderWindow::FileOperationState::Task::InitializeFileSystemOptions(FileSystemOptions& options, void* operationControlCookie) const noexcept
{
    options                                                                   = {};
    options.sizeBytes                                                         = sizeof(FileSystemOptions);
    options.bandwidthLimitBytesPerSecond                                      = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    options.copyMoveMaxConcurrency                                            = 0;
    options.operationControl                                                  = const_cast<Task*>(this);
    options.operationControlCookie                                            = operationControlCookie;
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (plans && ! plans->empty())
    {
        const FileOperations::LinkPolicy capturedPolicy =
            std::visit([](const auto& typedPlan) noexcept { return typedPlan.options.linkPolicy; }, plans->front());
        options.linkPolicy = capturedPolicy == FileOperations::LinkPolicy::Skip ? FILESYSTEM_LINK_SKIP : FILESYSTEM_LINK_PRESERVE;
    }
    if (_executionMode == ExecutionMode::PerItem && (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE))
    {
        options.copyMoveMaxConcurrency = std::max(1u, _perItemMaxConcurrencyBudget);
    }
}

void FolderWindow::FileOperationState::Task::SetWaitForOthers(bool wait) noexcept
{
    FileOperationState* state = _state;
    if (state)
    {
        std::scoped_lock lock(state->_queueMutex);
        if (_enteredOperation.load(std::memory_order_acquire))
        {
            return;
        }

        _waitForOthers.store(wait, std::memory_order_release);
        ++_perf.queueNotifyAllCount;
        state->NotifyQueueChanged();
        return;
    }

    if (_enteredOperation.load(std::memory_order_acquire))
    {
        return;
    }

    _waitForOthers.store(wait, std::memory_order_release);
}

void FolderWindow::FileOperationState::Task::SetWaitingInQueue(bool waiting) noexcept
{
    const bool wasWaiting = _waitingInQueue.load(std::memory_order_acquire);
    if (wasWaiting == waiting)
    {
        return;
    }

    _waitingInQueue.store(waiting, std::memory_order_release);
    MarkRateSamplingStateChanged();
    if (waiting)
    {
        RequestPresentationReveal();
    }
}

bool FolderWindow::FileOperationState::Task::IsPresentationHidden() const noexcept
{
    const TaskPresentationState state = _presentationState.load(std::memory_order_acquire);
    return state == TaskPresentationState::Hidden || state == TaskPresentationState::SuppressedCleanSuccess;
}

bool FolderWindow::FileOperationState::Task::IsPresentationVisible() const noexcept
{
    const TaskPresentationState state = _presentationState.load(std::memory_order_acquire);
    return state == TaskPresentationState::RevealRequested || state == TaskPresentationState::Presented;
}

void FolderWindow::FileOperationState::Task::RequestPresentationReveal() noexcept
{
    TaskPresentationState expected = TaskPresentationState::Hidden;
    if (_presentationState.compare_exchange_strong(expected, TaskPresentationState::RevealRequested, std::memory_order_acq_rel, std::memory_order_acquire) &&
        _state)
    {
        const ULONGLONG nowTick = _state->TaskPresentationNowTick();
        const ULONGLONG admittedTick =
            _presentationDeadlineTick >= FileOperations::kTaskCardRevealDelayMs ? _presentationDeadlineTick - FileOperations::kTaskCardRevealDelayMs : 0u;
        Debug::Perf::Emit(L"FileOps.TaskPresentation.RevealMs",
                          L"non-clean",
                          nowTick >= admittedTick ? nowTick - admittedTick : 0u,
                          FileOperations::kTaskCardRevealDelayMs,
                          0u,
                          S_OK);
        _state->RequestTaskPresentationRefresh();
    }
}

void FolderWindow::FileOperationState::Task::RequestActionablePromptPresentation() noexcept
{
    if (_state)
    {
        _state->RequestActionablePromptPresentation();
    }
}

void FolderWindow::FileOperationState::Task::SetQueuePaused(bool paused) noexcept
{
    {
        std::scoped_lock lock(_pauseMutex);
        const bool wasPaused = _queuePaused.load(std::memory_order_relaxed);
        if (wasPaused == paused)
        {
            return;
        }

        _queuePaused.store(paused, std::memory_order_release);
    }

    MarkRateSamplingStateChanged();
    if (! paused)
    {
        _pauseCv.notify_all();
    }

    GetPerItemTaskScheduler().NotifyWorkAvailable();
}

void FolderWindow::FileOperationState::Task::MarkRateSamplingStateChanged() noexcept
{
    _rateSamplingStateChangeTick.store(GetTickCount64(), std::memory_order_release);
}

void FolderWindow::FileOperationState::Task::ToggleConflictApplyToAllChecked() noexcept
{
    std::scoped_lock lock(_conflictArbiter.mutex);
    if (! _conflictArbiter.prompt.active || _conflictArbiter.prompt.metadataLoading || ! _conflictArbiter.prompt.applyToAllEligible)
    {
        return;
    }

    _conflictArbiter.prompt.applyToAllChecked = ! _conflictArbiter.prompt.applyToAllChecked;
}

void FolderWindow::FileOperationState::Task::SubmitConflictDecision(ConflictAction action, bool applyToAllChecked) noexcept
{
    {
        std::scoped_lock lock(_conflictArbiter.mutex);
        if (! _conflictArbiter.prompt.active || _conflictArbiter.prompt.metadataLoading)
        {
            return;
        }

        const bool actionOffered = std::ranges::any_of(std::span(_conflictArbiter.prompt.actions.data(), _conflictArbiter.prompt.actionCount),
                                                       [action](ConflictAction offered) noexcept { return offered == action; });
        if (! actionOffered)
        {
            action            = ConflictAction::Cancel;
            applyToAllChecked = false;
        }
        const bool explicitSkipAll          = action == ConflictAction::SkipAll;
        _conflictArbiter.decisionAction     = explicitSkipAll ? ConflictAction::Skip : action;
        _conflictArbiter.decisionApplyToAll = _conflictArbiter.prompt.applyToAllEligible &&
                                              (explicitSkipAll || (action != ConflictAction::Retry && action != ConflictAction::Skip && applyToAllChecked));
    }

    if (_conflictArbiter.decisionEvent)
    {
        static_cast<void>(SetEvent(_conflictArbiter.decisionEvent.get()));
    }
}

bool FolderWindow::FileOperationState::Task::HasStarted() const noexcept
{
    return _started.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::HasEnteredOperation() const noexcept
{
    return _enteredOperation.load(std::memory_order_acquire);
}

ULONGLONG FolderWindow::FileOperationState::Task::GetEnteredOperationTick() const noexcept
{
    return _enteredOperationTick.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::IsPaused() const noexcept
{
    return _paused.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::IsWaitingForOthers() const noexcept
{
    return _waitForOthers.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::IsWaitingInQueue() const noexcept
{
    return _waitingInQueue.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::IsOverlapQueued() const noexcept
{
    return _overlapQueued.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::Task::AllowsConcurrentOverlapWith(uint64_t taskId) const noexcept
{
    return std::ranges::find(_concurrentOverlapTaskIds.begin(),
                             _concurrentOverlapTaskIds.begin() + static_cast<std::ptrdiff_t>(_concurrentOverlapTaskIdCount),
                             taskId) != _concurrentOverlapTaskIds.begin() + static_cast<std::ptrdiff_t>(_concurrentOverlapTaskIdCount);
}

bool FolderWindow::FileOperationState::Task::IsQueuePaused() const noexcept
{
    return _queuePaused.load(std::memory_order_acquire);
}

uint64_t FolderWindow::FileOperationState::Task::GetQueueOrderKey() const noexcept
{
    return _queueOrderKey.load(std::memory_order_acquire);
}

void FolderWindow::FileOperationState::Task::SetDestinationFolder(const std::filesystem::path& folder)
{
    if (_started.load(std::memory_order_acquire))
    {
        return;
    }

    std::scoped_lock lock(_operationMutex);
    _destinationFolder = folder;
}

std::filesystem::path FolderWindow::FileOperationState::Task::GetDestinationFolder() const
{
    std::scoped_lock lock(_operationMutex);
    return _destinationFolder;
}

unsigned long FolderWindow::FileOperationState::Task::GetPlannedItemCount() const noexcept
{
    const uint64_t count64 = static_cast<uint64_t>(_sourcePaths.size());
    if (count64 > std::numeric_limits<unsigned long>::max())
    {
        return std::numeric_limits<unsigned long>::max();
    }
    return static_cast<unsigned long>(count64);
}

uint64_t FolderWindow::FileOperationState::Task::GetId() const noexcept
{
    return _taskId;
}

HRESULT FolderWindow::FileOperationState::Task::GetResult() const noexcept
{
    return _resultHr.load(std::memory_order_acquire);
}

void FolderWindow::FileOperationState::Task::InitializeSourceItemResultBuilders() noexcept
{
    std::scoped_lock lock(_sourceItemStatusMutex);
    _sourceItemResultBuilders.clear();
    _sourceItemResultBuilders.resize(_sourcePaths.size());
}

void FolderWindow::FileOperationState::Task::MarkSourceItemsMutationPossible() noexcept
{
    std::scoped_lock lock(_sourceItemStatusMutex);
    if (_sourceItemResultBuilders.size() < _sourcePaths.size())
    {
        _sourceItemResultBuilders.resize(_sourcePaths.size());
    }
    for (SourceItemResultBuilder& builder : _sourceItemResultBuilders)
    {
        builder.phase = SourceItemExecutionPhase::MutationPossible;
    }
}

bool FolderWindow::FileOperationState::Task::ClipboardMutationGateAllows(const HRESULT readinessStatus,
                                                                         const HRESULT consumptionStatus,
                                                                         const bool consumed) noexcept
{
    return SUCCEEDED(readinessStatus) && consumptionStatus == S_OK && consumed;
}

bool FolderWindow::FileOperationState::Task::StoreTypedItemResult(FileOperations::FileOperationItemResult result) noexcept
{
    if (result.sourceIndex >= _sourcePaths.size())
    {
        _invalidTerminalStoreRejectCount.fetch_add(1u, std::memory_order_relaxed);
        Debug::Warning(L"File Operations rejected a terminal item result with an invalid source index (taskId={}, sourceIndex={}, sourceCount={}).",
                       _taskId,
                       result.sourceIndex,
                       _sourcePaths.size());
        return false;
    }

    bool duplicate = false;
    {
        std::scoped_lock lock(_sourceItemStatusMutex);
        if (_sourceItemResultBuilders.size() < _sourcePaths.size())
        {
            _sourceItemResultBuilders.resize(_sourcePaths.size());
        }
        SourceItemResultBuilder& builder = _sourceItemResultBuilders[result.sourceIndex];
        if (builder.terminal.has_value())
        {
            duplicate = true;
        }
        else
        {
            builder.status   = result.status;
            builder.terminal = std::move(result);
        }
    }
    if (duplicate)
    {
        _duplicateTerminalStoreRejectCount.fetch_add(1u, std::memory_order_relaxed);
        Debug::Warning(
            L"File Operations rejected a duplicate terminal item result; first truth was retained (taskId={}, sourceIndex={}).", _taskId, result.sourceIndex);
        return false;
    }
    return true;
}

bool FolderWindow::FileOperationState::Task::CancelRequestedDurable() const noexcept
{
    // RequestCancel records the timestamp before publishing _cancelled; either store, or the
    // worker's stop token, proves the user or host asked for the cancel.
    return _cancelRequestedTick.load(std::memory_order_acquire) != 0 || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested();
}

bool FolderWindow::FileOperationState::Task::IsCancellationOutcome(HRESULT hr) const noexcept
{
    // A synchronous call the R0f cancel watch aborted fails with ERROR_OPERATION_ABORTED; under a
    // requested cancel that is the cancellation itself, not a provider fault, wherever the
    // status is sealed into a result.
    return IsCancellationStatus(hr) || (hr == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED) && CancelRequestedDurable());
}

HRESULT FolderWindow::FileOperationState::Task::FinalizeTypedItemResults(HRESULT operationStatus) noexcept
{
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    std::scoped_lock lock(_sourceItemStatusMutex);
    _sourceItemResultBuilders.resize(_sourcePaths.size());

    const auto strategyForIndex = [&](size_t wantedIndex) noexcept
    {
        if (! plans)
        {
            return FileOperations::OperationStrategy::Copy;
        }
        size_t currentIndex = 0u;
        for (const FileOperations::FileOperationPlan& plan : *plans)
        {
            if (const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan))
            {
                if (wantedIndex < currentIndex + transfer->selectedItems.size())
                {
                    return transfer->strategy;
                }
                currentIndex += transfer->selectedItems.size();
            }
            else if (const auto* rename = std::get_if<FileOperations::RenamePlan>(&plan))
            {
                if (wantedIndex < currentIndex + rename->finalMappings.size())
                {
                    return FileOperations::OperationStrategy::Native;
                }
                currentIndex += rename->finalMappings.size();
            }
            else if (const auto* deletion = std::get_if<FileOperations::DeletePlan>(&plan))
            {
                if (wantedIndex < currentIndex + deletion->selectedItems.size())
                {
                    return FileOperations::OperationStrategy::Native;
                }
                currentIndex += deletion->selectedItems.size();
            }
        }
        return FileOperations::OperationStrategy::Copy;
    };

    const bool taskCanceled = IsCancellationStatus(operationStatus) || CancelRequestedDurable();
    for (size_t index = 0u; index < _sourcePaths.size(); ++index)
    {
        SourceItemResultBuilder& builder = _sourceItemResultBuilders[index];
        if (builder.terminal.has_value())
        {
            continue;
        }

        FileOperations::FileOperationItemResult result{};
        result.sourceIndex     = index;
        result.strategy        = strategyForIndex(index);
        result.finalSourcePath = _sourcePaths[index].native();
        result.verification    = (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE) ? FileOperations::VerificationState::NotRequested
                                                                                                  : FileOperations::VerificationState::NotApplicable;
        result.status          = builder.status.value_or(operationStatus);
        if (result.status == E_PENDING)
        {
            result.status = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (taskCanceled && result.status == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED))
        {
            // The R0f cancel watch aborted this item's pending synchronous call: report the cancel.
            result.status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const std::optional<FileSystemItemMutationResult>& mutation = builder.mutation;
        const bool validMutation                                    = mutation.has_value() && IsValidItemMutationResultPrefix(mutation.value());
        if (validMutation)
        {
            result.ownedStageDisposition = GetOwnedStageDisposition(mutation.value());
        }
        if (builder.phase == SourceItemExecutionPhase::Preparing && ! validMutation)
        {
            result.publication       = FileOperations::PublicationState::NotAttempted;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = IsCancellationStatus(result.status) ? FileOperations::ItemCompletion::Canceled : FileOperations::ItemCompletion::Failed;
        }
        else if (validMutation)
        {
            if (mutation->outcomeKnown == FALSE)
            {
                result.publication =
                    (_operation == FILESYSTEM_DELETE) ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Unknown;
                result.sourceDisposition =
                    (_operation == FILESYSTEM_COPY) ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Unknown;
                result.completion = FileOperations::ItemCompletion::Indeterminate;
                result.status     = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
            }
            else if (mutation->mutationCommitted != FALSE)
            {
                result.publication =
                    (_operation == FILESYSTEM_DELETE) ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Published;
                result.sourceDisposition =
                    mutation->originalStillPresent != FALSE ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Removed;
                // Mutation truth describes what reached storage; it does not turn a canceled or
                // otherwise failed operation into a completed item. A directory copy can publish
                // children before cancellation and must retain both facts.
                if (IsCancellationStatus(result.status))
                {
                    result.completion = FileOperations::ItemCompletion::Canceled;
                }
                else
                {
                    result.completion = FAILED(result.status) ? FileOperations::ItemCompletion::Failed : FileOperations::ItemCompletion::Completed;
                }
            }
            else
            {
                result.publication =
                    (_operation == FILESYSTEM_DELETE) ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::NotPublished;
                result.sourceDisposition = FileOperations::SourceDisposition::Retained;
                result.completion = IsCancellationStatus(result.status) ? FileOperations::ItemCompletion::Canceled : FileOperations::ItemCompletion::Failed;
            }
        }
        else if (taskCanceled || IsCancellationStatus(result.status))
        {
            result.publication       = FileOperations::PublicationState::NotAttempted;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Canceled;
            result.status            = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        else if (result.status == S_FALSE)
        {
            // S_FALSE without a receipt is a root the user skipped before any mutation (a
            // conflict answered Skip): nothing was published and the source is untouched.
            result.publication       = FileOperations::PublicationState::NotAttempted;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Skipped;
        }
        else if (SUCCEEDED(result.status))
        {
            // A root that reached MutationPossible without its own terminal receipt has no
            // publication or source truth even when the task as a whole succeeded. Success
            // without per-item evidence is never reported as Completed.
            result.publication = (_operation == FILESYSTEM_DELETE) ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Unknown;
            result.sourceDisposition =
                (_operation == FILESYSTEM_COPY) ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Unknown;
            result.completion = FileOperations::ItemCompletion::Indeterminate;
            result.status     = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        else
        {
            result.publication = (_operation == FILESYSTEM_DELETE) ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Unknown;
            result.sourceDisposition =
                (_operation == FILESYSTEM_COPY) ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Unknown;
            result.completion = result.status == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) ? FileOperations::ItemCompletion::Indeterminate
                                                                                         : FileOperations::ItemCompletion::Failed;
        }
        if (validMutation && OwnedStageDispositionIsIndeterminate(result.ownedStageDisposition))
        {
            result.completion = FileOperations::ItemCompletion::Indeterminate;
            result.status     = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        builder.terminal = std::move(result);
    }

    bool anyCompleted               = false;
    bool anySkipped                 = false;
    bool anyCanceled                = false;
    bool anyFailed                  = false;
    bool anyIndeterminate           = false;
    bool anyIntentionalSourceKept   = false;
    bool anyVerificationUnavailable = false;
    HRESULT firstFailure            = S_OK;
    size_t failedCount              = 0u;
    for (const SourceItemResultBuilder& builder : _sourceItemResultBuilders)
    {
        const std::optional<FileOperations::FileOperationItemResult>& item = builder.terminal;
        if (! item.has_value())
        {
            anyIndeterminate = true;
            continue;
        }
        // An unknown publication or source axis makes the task indeterminate, except on an item the
        // user or host canceled: the cancel is that item's terminal cause, the unknown axis stays
        // visible on the item and still blocks Retry, and the task ends Canceled.
        anyIndeterminate =
            anyIndeterminate || item->completion == FileOperations::ItemCompletion::Indeterminate ||
            (item->completion != FileOperations::ItemCompletion::Canceled &&
             (item->publication == FileOperations::PublicationState::Unknown || item->sourceDisposition == FileOperations::SourceDisposition::Unknown));
        anyCompleted = anyCompleted || item->completion == FileOperations::ItemCompletion::Completed;
        anySkipped   = anySkipped || item->completion == FileOperations::ItemCompletion::Skipped;
        anyCanceled  = anyCanceled || item->completion == FileOperations::ItemCompletion::Canceled;
        if (item->completion == FileOperations::ItemCompletion::Failed)
        {
            anyFailed = true;
            ++failedCount;
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = item->status;
            }
        }
        anyIntentionalSourceKept =
            anyIntentionalSourceKept || (_operation == FILESYSTEM_MOVE && item->publication == FileOperations::PublicationState::Published &&
                                         item->sourceDisposition == FileOperations::SourceDisposition::Retained);
        anyVerificationUnavailable = anyVerificationUnavailable || item->verification == FileOperations::VerificationState::Unavailable;
    }

    Debug::Perf::Emit(L"FileOps.Result.Items",
                      std::format(L"op={} indeterminate={} sourceKept={}",
                                  OperationToString(_operation),
                                  anyIndeterminate ? L"true" : L"false",
                                  anyIntentionalSourceKept ? L"true" : L"false"),
                      0u,
                      static_cast<uint64_t>(_sourceItemResultBuilders.size()),
                      0u,
                      operationStatus);

    if (IsTraversalResourceLimitStatus(operationStatus))
    {
        // A traversal/resource ceiling is an exact task-terminal cause even when a partially
        // discovered item must retain cautious Unknown publication axes. Do not erase that
        // actionable stop condition with the generic indeterminate-outcome HRESULT.
        return operationStatus;
    }
    if (anyIndeterminate)
    {
        return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
    }
    if (anyFailed)
    {
        return failedCount == 1u && ! anyCompleted && ! anySkipped && ! anyCanceled ? firstFailure : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }
    if (anyCanceled)
    {
        return anyCompleted || anySkipped ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    if (anySkipped || anyIntentionalSourceKept || anyVerificationUnavailable)
    {
        return S_FALSE;
    }
    return S_OK;
}

FileSystemOperation FolderWindow::FileOperationState::Task::GetOperation() const noexcept
{
    return _operation;
}

FolderWindow::Pane FolderWindow::FileOperationState::Task::GetSourcePane() const noexcept
{
    return _sourcePane;
}

std::optional<FolderWindow::Pane> FolderWindow::FileOperationState::Task::GetDestinationPane() const noexcept
{
    return _destinationPane;
}

void FolderWindow::FileOperationState::Task::WaitWhilePaused(const std::atomic<bool>* externalStop) noexcept
{
    const DWORD currentThreadId = GetCurrentThreadId();

    for (;;)
    {
        // Cancellation wakes either condition variable, but queue/conflict state may
        // remain set until another task finishes. Do not re-enter an already-awake wait.
        if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested() ||
            (externalStop != nullptr && externalStop->load(std::memory_order_acquire)))
        {
            return;
        }
        const bool shouldPause     = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
        bool shouldWaitForConflict = false;
        {
            std::scoped_lock lock(_conflictArbiter.mutex);
            shouldWaitForConflict = _conflictArbiter.prompt.active && _conflictArbiter.ownerThreadId != 0 && _conflictArbiter.ownerThreadId != currentThreadId;
        }

        if (! shouldPause && ! shouldWaitForConflict)
        {
            return;
        }

        if (shouldPause)
        {
#ifdef ENABLE_TESTS
            _dbgPauseWaiterCount.fetch_add(1u, std::memory_order_release);
            const auto pauseWaiterScope = wil::scope_exit([&] noexcept { _dbgPauseWaiterCount.fetch_sub(1u, std::memory_order_release); });
#endif
            const uint64_t waitStartUs = PerfNowUs();
            std::unique_lock lock(_pauseMutex);
            _pauseCv.wait(lock,
                          [&]
            {
                const bool stillPaused = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
                return ! stillPaused || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested() ||
                       (externalStop != nullptr && externalStop->load(std::memory_order_acquire));
            });
            _perf.pauseWaitUs.fetch_add(PerfElapsedUs(waitStartUs), std::memory_order_relaxed);
            continue;
        }

        const uint64_t waitStartUs = PerfNowUs();
        std::unique_lock lock(_conflictArbiter.mutex);
        _conflictArbiter.cv.wait(lock,
                                 [&]
        {
            const bool stillPaused = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
            const bool waitingForConflict =
                _conflictArbiter.prompt.active && _conflictArbiter.ownerThreadId != 0 && _conflictArbiter.ownerThreadId != currentThreadId;
            const bool ready = ! waitingForConflict || stillPaused || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested() ||
                               (externalStop != nullptr && externalStop->load(std::memory_order_acquire));
#ifdef ENABLE_TESTS
            if (! ready && _dbgConflictWaitBeforeSleepGate.load(std::memory_order_acquire))
            {
                _dbgConflictWaitBeforeSleepReached.store(true, std::memory_order_release);
                _dbgConflictWaitBeforeSleepGate.wait(true, std::memory_order_acquire);
            }
#endif
            return ready;
        });
        _perf.conflictConvergenceWaitUs += PerfElapsedUs(waitStartUs);
    }
}

void FolderWindow::FileOperationState::Task::WakePauseWaiters() noexcept
{
    {
        // Stop/cancel flags are published before this call. Synchronize with both
        // predicate mutexes so a notification cannot precede a wait that saw false.
        std::scoped_lock lock(_pauseMutex, _conflictArbiter.mutex);
    }
    _pauseCv.notify_all();
    _conflictArbiter.cv.notify_all();
}

bool ShouldRetryPublishedDestinationVerification(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND) ||
           hr == HRESULT_FROM_WIN32(ERROR_BUSY) || hr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED) || hr == HRESULT_FROM_WIN32(ERROR_LOCK_VIOLATION) ||
           hr == HRESULT_FROM_WIN32(ERROR_NETWORK_BUSY) || hr == HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE) || hr == HRESULT_FROM_WIN32(ERROR_RETRY) ||
           hr == HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT) || hr == HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) || hr == HRESULT_FROM_WIN32(ERROR_TIMEOUT) ||
           hr == HRESULT_FROM_WIN32(ERROR_UNEXP_NET_ERR);
}

HRESULT FolderWindow::FileOperationState::Task::ExecuteInlineRename() noexcept
{
    const uint64_t executeStartedUs                                           = PerfNowUs();
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (! plans || plans->size() != 1u || ! _fileSystem)
    {
        return E_UNEXPECTED;
    }
    const auto* rename = std::get_if<FileOperations::RenamePlan>(&plans->front());
    if (rename == nullptr || rename->origin != FileOperations::RenameOrigin::InlineRename || rename->finalMappings.size() != 1u ||
        ! rename->schedule.layers.empty() || ! rename->schedule.cycleOperationIndices.empty() || ! rename->endpoint.pathIdentity.has_value())
    {
        return E_UNEXPECTED;
    }

#ifdef ENABLE_TESTS
    g_fileOpsInlineRenameExecutionThreadId.store(GetCurrentThreadId(), std::memory_order_release);
    g_fileOpsInlineRenameExecutionAttempts.fetch_add(1u, std::memory_order_acq_rel);
#endif

    const FileOperations::RenameStep& step = rename->finalMappings.front();
    const std::wstring& sourcePath         = step.source.providerPath;
    std::wstring parentPath;
    if (! TryGetFileSystemParentPath(rename->endpoint.pathIdentity.value(), sourcePath, parentPath))
    {
        return E_INVALIDARG;
    }
    const auto queryNameContract = [&]() noexcept
    {
        return FileSystemRouteContract::QueryChildNameContract(_fileSystem.get(), parentPath, step.finalLeafName, FILESYSTEM_RENAME, rename->endpoint.pluginId);
    };
    const auto nameContractMatchesAdmission = [&](const FileSystemRouteContract::ChildNameContractResult& contract) noexcept
    {
        return contract.state == FileSystemRouteContract::QueryState::Available && contract.nameStatus == FILESYSTEM_CHILD_NAME_VALID &&
               contract.joinedPath == step.providerJoinedPath && contract.collisionKey == step.providerCollisionKey;
    };
    const FileSystemRouteContract::ChildNameContractResult initialNameContract = queryNameContract();
    if (! nameContractMatchesAdmission(initialNameContract))
    {
        return initialNameContract.state == FileSystemRouteContract::QueryState::Available && FAILED(initialNameContract.failureStatus)
                   ? initialNameContract.failureStatus
                   : (FAILED(initialNameContract.status) ? initialNameContract.status : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH));
    }
    const std::wstring destinationPath = initialNameContract.joinedPath;
    if (destinationPath.empty() || destinationPath == sourcePath)
    {
        return S_FALSE;
    }

    switch (GuardLiveOutputBeforeInvalidation(sourcePath, FileOperations::MutationInterlockAccess::WriteSource))
    {
        case LiveOutputGuardDisposition::Skip: return S_FALSE;
        case LiveOutputGuardDisposition::Cancel: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        case LiveOutputGuardDisposition::RetryCurrentMutation:
        case LiveOutputGuardDisposition::Proceed: break;
    }

    const bool equivalentPath = EquivalentPath(rename->endpoint.pathIdentity.value(), sourcePath, destinationPath);
    if (equivalentPath)
    {
        switch (rename->endpoint.pathIdentity->caseOnlyRename)
        {
            case FileSystemPathCaseOnlyRename::Supported: break;
            case FileSystemPathCaseOnlyRename::NoOp: return S_FALSE;
            case FileSystemPathCaseOnlyRename::Unsupported:
            case FileSystemPathCaseOnlyRename::NotApplicable: return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
    }

    constexpr FileSystemBindFlags sourceBindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_RENAME);
    constexpr FileSystemBindFlags metadataBindFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    constexpr FileSystemBindFlags destinationReplacementBindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION);
    FileOperations::ObjectBindingResult source =
        FileOperations::BindObjectAuthority(_fileSystem.get(), sourcePath, rename->endpoint.profileId, sourceBindFlags);
    if (source.state != FileOperations::ObjectBindingState::Bound)
    {
        return FAILED(source.status) ? source.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const auto parentScope = std::ranges::find_if(_mutationInterlockScopes,
                                                  [&](const FileOperations::MutationInterlockScope& scope) noexcept
    {
        return FileOperations::QualifiedEndpointsReferToSameRoot(scope.endpoint, rename->endpoint) &&
               EquivalentPath(rename->endpoint.pathIdentity.value(), scope.providerPath, parentPath);
    });
    if (parentScope == _mutationInterlockScopes.end() || ! parentScope->root)
    {
        return E_UNEXPECTED;
    }

    const auto recordStatus = [&](HRESULT status, bool completed) noexcept
    {
        FileOperations::FileOperationItemResult result{};
        result.sourceIndex          = 0u;
        result.strategy             = FileOperations::OperationStrategy::Native;
        result.verification         = FileOperations::VerificationState::NotApplicable;
        result.status               = status;
        result.finalSourcePath      = sourcePath;
        result.finalDestinationPath = destinationPath;
        if (completed)
        {
            result.publication       = FileOperations::PublicationState::Published;
            result.sourceDisposition = FileOperations::SourceDisposition::Removed;
            result.completion        = FileOperations::ItemCompletion::Completed;
        }
        else if (IsCancellationOutcome(status))
        {
            result.status            = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            result.publication       = FileOperations::PublicationState::NotAttempted;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Canceled;
        }
        else if (status == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE))
        {
            result.publication       = FileOperations::PublicationState::Unknown;
            result.sourceDisposition = FileOperations::SourceDisposition::Unknown;
            result.completion        = FileOperations::ItemCompletion::Indeterminate;
        }
        else
        {
            result.publication       = FileOperations::PublicationState::NotPublished;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Failed;
        }
        static_cast<void>(StoreTypedItemResult(std::move(result)));
        {
            std::scoped_lock lock(_progressMutex);
            _progressTotalItems     = 1u;
            _progressCompletedItems = completed ? 1u : 0u;
            _lastItemIndex          = 0u;
            _lastItemHr             = status;
            PublishProgressCountersLocked(*this);
        }
        if (completed)
        {
            StorePublishedTopLevelCompletionSnapshot(*this, MarkTopLevelItemCompleted(*this, 0u));
        }
    };
    const auto revalidateBeforeMutation = [&]() noexcept -> HRESULT
    {
        const FileOperations::ObjectRevalidationResult currentSource =
            FileOperations::RevalidateObjectAuthority(_fileSystem.get(), sourcePath, rename->endpoint.profileId, sourceBindFlags, source.authority);
        if (currentSource.state != FileOperations::ObjectRevalidationState::Same)
        {
            return FAILED(currentSource.status) ? currentSource.status : HRESULT_FROM_WIN32(ERROR_RETRY);
        }
        const FileOperations::ObjectRevalidationResult currentParent = FileOperations::RevalidateObjectAuthority(
            _fileSystem.get(), parentPath, rename->endpoint.profileId, metadataBindFlags, parentScope->root->retained.authority);
        if (currentParent.state != FileOperations::ObjectRevalidationState::Same)
        {
            return FAILED(currentParent.status) ? currentParent.status : HRESULT_FROM_WIN32(ERROR_RETRY);
        }
        return S_OK;
    };
    const auto sameBoundObject = [&](const FileOperations::BoundObjectAuthority& other, bool& same) noexcept -> HRESULT
    {
        bool sameRevision = false;
        return FileOperations::CrossCheckBoundObjectIdentity(source.authority, other, same, sameRevision);
    };

    PerItemCallbackCookie cookie{0u};
    cookie.operationDestinationPath = destinationPath;
    FileSystemFlags itemFlags       = FILESYSTEM_FLAG_NONE;
    wil::com_ptr<IFileSystemBoundObject> grantedExpectedDestination;
    unsigned int retryCount = 0u;
    for (;;)
    {
        WaitWhilePaused();
        if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
        {
            recordStatus(HRESULT_FROM_WIN32(ERROR_CANCELLED), false);
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const FileSystemRouteContract::ChildNameContractResult currentNameContract = queryNameContract();
        if (! nameContractMatchesAdmission(currentNameContract) || currentNameContract.joinedPath != destinationPath)
        {
            const HRESULT status = currentNameContract.state == FileSystemRouteContract::QueryState::Available && FAILED(currentNameContract.failureStatus)
                                       ? currentNameContract.failureStatus
                                       : (FAILED(currentNameContract.status) ? currentNameContract.status : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH));
            recordStatus(status, false);
            return status;
        }

        const HRESULT revalidationHr = revalidateBeforeMutation();
        if (FAILED(revalidationHr))
        {
            recordStatus(revalidationHr, false);
            return revalidationHr;
        }

        FileOperations::ObjectBindingResult destination =
            FileOperations::BindObjectAuthority(_fileSystem.get(), destinationPath, rename->endpoint.profileId, destinationReplacementBindFlags);
        bool destinationIsSource = false;
        if (destination.state == FileOperations::ObjectBindingState::Bound)
        {
            const HRESULT sameHr = sameBoundObject(destination.authority, destinationIsSource);
            if (FAILED(sameHr))
            {
                recordStatus(sameHr, false);
                return sameHr;
            }
        }
        else if (destination.state != FileOperations::ObjectBindingState::Missing)
        {
            const HRESULT destinationHr = FAILED(destination.status) ? destination.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            recordStatus(destinationHr, false);
            return destinationHr;
        }

        HRESULT renameHr = S_OK;
        FileSystemConditionalMutationResult mutation{};
        mutation.sizeBytes = sizeof(mutation);
        wil::com_ptr<IFileSystemBoundObject> renamedObject;
        const bool overwriteGranted = (itemFlags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0u;
        if (destination.state == FileOperations::ObjectBindingState::Bound && ! destinationIsSource && ! overwriteGranted)
        {
            renameHr = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        else
        {
#ifdef ENABLE_TESTS
            if (g_fileOpsInlineRenameBeforeMutationPausePoint.enabled.load(std::memory_order_acquire))
            {
                g_fileOpsInlineRenameBeforeMutationPausePoint.entered.store(true, std::memory_order_release);
                const auto clearEntered =
                    wil::scope_exit([]() noexcept { g_fileOpsInlineRenameBeforeMutationPausePoint.entered.store(false, std::memory_order_release); });
                while (g_fileOpsInlineRenameBeforeMutationPausePoint.enabled.load(std::memory_order_acquire) &&
                       ! g_fileOpsInlineRenameBeforeMutationPausePoint.releaseRequested.load(std::memory_order_acquire) &&
                       ! _cancelled.load(std::memory_order_acquire) && ! _stopToken.stop_requested())
                {
                    Sleep(1u);
                }
            }
#endif
            FileSystemOptions options{};
            InitializeFileSystemOptions(options);
            IFileSystemBoundObject* expectedDestination =
                grantedExpectedDestination
                    ? grantedExpectedDestination.get()
                    : (destination.state == FileOperations::ObjectBindingState::Bound && ! destinationIsSource ? destination.authority.boundObject.get()
                                                                                                               : nullptr);
            renameHr = source.authority.boundObject->RenameIfUnchanged(
                destinationPath.c_str(), expectedDestination, itemFlags, &options, &mutation, renamedObject.put());
            if (SUCCEEDED(renameHr))
            {
                if (! renamedObject || mutation.sizeBytes != sizeof(mutation) || mutation.outcomeKnown == FALSE || mutation.mutationCommitted == FALSE ||
                    mutation.originalStillPresent != FALSE)
                {
                    renameHr = E_UNEXPECTED;
                }
                else
                {
                    const FileOperations::ObjectRevalidationResult published = FileOperations::RevalidateObjectAuthority(
                        _fileSystem.get(), destinationPath, rename->endpoint.profileId, metadataBindFlags, source.authority);
                    if (published.state != FileOperations::ObjectRevalidationState::Same)
                    {
                        renameHr = FAILED(published.status) ? published.status : HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
                    }
                }
            }
            else if (mutation.sizeBytes != sizeof(mutation) || mutation.outcomeKnown == FALSE || mutation.mutationCommitted != FALSE)
            {
                // A failed conditional mutation that may already have committed cannot be retried or
                // presented as an ordinary name conflict. The final namespace state is indeterminate.
                renameHr = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
            }
        }

        if (SUCCEEDED(renameHr))
        {
            NoteLiveOutputPublished(destinationPath);
            {
                std::scoped_lock lock(_progressPathMutex);
                _progressSourcePath      = sourcePath;
                _progressDestinationPath = destinationPath;
                PublishDiagnosticPathSnapshotLocked(*this);
            }
            recordStatus(S_OK, true);
            ClearConflictPrompt(*this);
            Debug::Perf::Emit(L"FileOps.InlineRename.ExecuteUs",
                              L"identity-bound",
                              PerfElapsedUs(executeStartedUs),
                              static_cast<uint64_t>(sourcePath.size() + destinationPath.size()) * sizeof(wchar_t),
                              1u,
                              S_OK);
            return S_OK;
        }

        if (renameHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE))
        {
            recordStatus(renameHr, false);
            return renameHr;
        }

        if (renameHr == HRESULT_FROM_WIN32(ERROR_FILE_INVALID) || renameHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) ||
            renameHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
        {
            // The exact destination named by the prior grant disappeared or was replaced. Do not
            // carry that consent to the substitute; a fresh retry must bind and prompt again.
            itemFlags = FILESYSTEM_FLAG_NONE;
            grantedExpectedDestination.reset();
        }

        const wil::com_ptr<IFileSystemIO> fileSystemIo = QueryFileSystemIo(_fileSystem.get());
        ConflictBucket bucket = ClassifyConflictBucket(FILESYSTEM_RENAME, itemFlags, fileSystemIo, renameHr, sourcePath, destinationPath, false);
        const bool allowRetry = IsRetryableConflictBucket(bucket);
        const ConflictPromptBeginResult promptBegin =
            BeginConflictPrompt(*this,
                                &cookie,
                                bucket,
                                renameHr,
                                sourcePath,
                                destinationPath,
                                allowRetry,
                                retryCount,
                                false,
                                destination.state == FileOperations::ObjectBindingState::Bound && ! destinationIsSource);
        ConflictAction action = promptBegin.action;
        if (promptBegin.ownsPrompt)
        {
            action = WaitForConflictDecision(*this, &cookie, promptBegin.decisionScope).first;
        }

        if (action == ConflictAction::Overwrite)
        {
            if (destination.state != FileOperations::ObjectBindingState::Bound || destinationIsSource)
            {
                recordStatus(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), false);
                return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
            }
            grantedExpectedDestination = destination.authority.boundObject;
            itemFlags                  = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
            continue;
        }
        if (action == ConflictAction::ReplaceReadOnly)
        {
            if (destination.state != FileOperations::ObjectBindingState::Bound || destinationIsSource)
            {
                recordStatus(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), false);
                return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
            }
            grantedExpectedDestination = destination.authority.boundObject;
            itemFlags                  = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
            continue;
        }
        if (action == ConflictAction::ReplaceLink)
        {
            if (destination.state != FileOperations::ObjectBindingState::Bound || destinationIsSource)
            {
                recordStatus(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), false);
                return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
            }
            grantedExpectedDestination = destination.authority.boundObject;
            itemFlags                  = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_LINK);
            continue;
        }
        if (action == ConflictAction::Retry)
        {
            ++retryCount;
            continue;
        }
        if (action == ConflictAction::Skip || action == ConflictAction::SkipAll)
        {
            const HRESULT skippedHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            FileOperations::FileOperationItemResult result{};
            result.sourceIndex          = 0u;
            result.strategy             = FileOperations::OperationStrategy::Native;
            result.publication          = FileOperations::PublicationState::NotPublished;
            result.verification         = FileOperations::VerificationState::NotApplicable;
            result.sourceDisposition    = FileOperations::SourceDisposition::Retained;
            result.completion           = FileOperations::ItemCompletion::Skipped;
            result.status               = skippedHr;
            result.finalSourcePath      = sourcePath;
            result.finalDestinationPath = destinationPath;
            static_cast<void>(StoreTypedItemResult(std::move(result)));
            ClearConflictPrompt(*this);
            return skippedHr;
        }

        const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        recordStatus(cancelledHr, false);
        return cancelledHr;
    }
}

HRESULT FolderWindow::FileOperationState::Task::ExecuteBatchRenameMutation(const size_t operationIndex,
                                                                           const std::filesystem::path& sourcePath,
                                                                           const std::filesystem::path& destinationPath,
                                                                           std::vector<FileOperations::BoundObjectAuthority>& authorities) noexcept
{
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (! plans || plans->size() != 1u || ! _fileSystem || operationIndex >= authorities.size())
    {
        return E_UNEXPECTED;
    }
    const auto* rename = std::get_if<FileOperations::RenamePlan>(&plans->front());
    const bool scheduledRename =
        rename != nullptr && (rename->origin == FileOperations::RenameOrigin::BatchRename || rename->origin == FileOperations::RenameOrigin::ChangeCase);
    if (! scheduledRename || ! rename->endpoint.pathIdentity.has_value() || operationIndex >= rename->finalMappings.size())
    {
        return E_UNEXPECTED;
    }

    WaitWhilePaused();
    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    const FileOperations::RenameStep& admittedStep = rename->finalMappings[operationIndex];
    std::wstring admittedParentPath;
    std::wstring admittedSourceLeaf;
    const std::optional<std::wstring> admittedParentKey =
        TryGetFileSystemParentPath(rename->endpoint.pathIdentity.value(), sourcePath.native(), admittedParentPath)
            ? TryMakePathKey(rename->endpoint.pathIdentity.value(), admittedParentPath)
            : std::nullopt;
    if (! admittedParentKey.has_value() || admittedParentKey.value() != admittedStep.providerParentKey ||
        ! TryGetFileSystemLeafName(rename->endpoint.pathIdentity.value(), sourcePath.native(), admittedSourceLeaf) ||
        destinationPath.native() != admittedStep.providerJoinedPath)
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    switch (GuardLiveOutputBeforeInvalidation(sourcePath.native(), FileOperations::MutationInterlockAccess::WriteSource))
    {
        case LiveOutputGuardDisposition::Skip: return S_FALSE;
        case LiveOutputGuardDisposition::Cancel: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        case LiveOutputGuardDisposition::RetryCurrentMutation:
        case LiveOutputGuardDisposition::Proceed: break;
    }
    const auto revalidateProviderName = [&]() noexcept
    {
        const FileSystemRouteContract::ChildNameContractResult source = FileSystemRouteContract::QueryChildNameContract(
            _fileSystem.get(), admittedParentPath, admittedSourceLeaf, FILESYSTEM_RENAME, rename->endpoint.pluginId);
        const FileSystemRouteContract::ChildNameContractResult current = FileSystemRouteContract::QueryChildNameContract(
            _fileSystem.get(), admittedParentPath, admittedStep.finalLeafName, FILESYSTEM_RENAME, rename->endpoint.pluginId);
        if (source.state == FileSystemRouteContract::QueryState::Available && source.nameStatus == FILESYSTEM_CHILD_NAME_VALID &&
            source.collisionKey == admittedStep.providerSourceCollisionKey && current.state == FileSystemRouteContract::QueryState::Available &&
            current.nameStatus == FILESYSTEM_CHILD_NAME_VALID && current.joinedPath == admittedStep.providerJoinedPath &&
            current.collisionKey == admittedStep.providerCollisionKey)
        {
            return S_OK;
        }
        if (source.state == FileSystemRouteContract::QueryState::Available && source.nameStatus == FILESYSTEM_CHILD_NAME_INVALID &&
            FAILED(source.failureStatus))
        {
            return source.failureStatus;
        }
        if (FAILED(source.status))
        {
            return source.status;
        }
        if (current.state == FileSystemRouteContract::QueryState::Available && current.nameStatus == FILESYSTEM_CHILD_NAME_INVALID &&
            FAILED(current.failureStatus))
        {
            return current.failureStatus;
        }
        return FAILED(current.status) ? current.status : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    };

    constexpr FileSystemBindFlags sourceBindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_RENAME);
    constexpr FileSystemBindFlags parentBindFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    constexpr FileSystemBindFlags destinationBindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION);

    FileOperations::BoundObjectAuthority& sourceAuthority = authorities[operationIndex];
    const FileOperations::ObjectRevalidationResult sourceRevalidation =
        FileOperations::RevalidateObjectAuthority(_fileSystem.get(), sourcePath.native(), rename->endpoint.profileId, sourceBindFlags, sourceAuthority);
    if (sourceRevalidation.state != FileOperations::ObjectRevalidationState::Same)
    {
        return FAILED(sourceRevalidation.status) ? sourceRevalidation.status : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }

    std::wstring parentPath;
    if (! TryGetFileSystemParentPath(rename->endpoint.pathIdentity.value(), sourcePath.native(), parentPath))
    {
        return E_INVALIDARG;
    }
    const auto parentScope = std::ranges::find_if(_mutationInterlockScopes,
                                                  [&](const FileOperations::MutationInterlockScope& scope) noexcept
    {
        return FileOperations::QualifiedEndpointsReferToSameRoot(scope.endpoint, rename->endpoint) && scope.root &&
               EquivalentPath(rename->endpoint.pathIdentity.value(), scope.providerPath, parentPath);
    });
    if (parentScope == _mutationInterlockScopes.end() || ! parentScope->root)
    {
        return E_UNEXPECTED;
    }
    const FileOperations::ObjectRevalidationResult parentRevalidation = FileOperations::RevalidateObjectAuthority(
        _fileSystem.get(), parentPath, rename->endpoint.profileId, parentBindFlags, parentScope->root->retained.authority);
    if (parentRevalidation.state != FileOperations::ObjectRevalidationState::Same)
    {
        return FAILED(parentRevalidation.status) ? parentRevalidation.status : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }

    PerItemCallbackCookie cookie{operationIndex};
    cookie.operationDestinationPath = destinationPath.native();
    FileSystemFlags itemFlags       = FILESYSTEM_FLAG_NONE;
    wil::com_ptr<IFileSystemBoundObject> grantedExpectedDestination;
    unsigned int retryCount = 0u;
    for (;;)
    {
        if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        const HRESULT providerNameHr = revalidateProviderName();
        if (FAILED(providerNameHr))
        {
            return providerNameHr;
        }
        FileOperations::ObjectBindingResult destination =
            FileOperations::BindObjectAuthority(_fileSystem.get(), destinationPath.native(), rename->endpoint.profileId, destinationBindFlags);
        bool destinationIsSource = false;
        if (destination.state == FileOperations::ObjectBindingState::Bound)
        {
            bool sameRevision    = false;
            const HRESULT sameHr = FileOperations::CrossCheckBoundObjectIdentity(sourceAuthority, destination.authority, destinationIsSource, sameRevision);
            if (FAILED(sameHr))
            {
                return sameHr;
            }
        }
        else if (destination.state != FileOperations::ObjectBindingState::Missing)
        {
            return FAILED(destination.status) ? destination.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        HRESULT renameHr = S_OK;
        FileSystemConditionalMutationResult mutation{};
        mutation.sizeBytes = sizeof(mutation);
        wil::com_ptr<IFileSystemBoundObject> renamedObject;
        const bool overwriteGranted = (itemFlags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0u;
        if (destination.state == FileOperations::ObjectBindingState::Bound && ! destinationIsSource && ! overwriteGranted)
        {
            renameHr = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        else
        {
            FileSystemOptions options{};
            InitializeFileSystemOptions(options);
            IFileSystemBoundObject* expectedDestination =
                grantedExpectedDestination
                    ? grantedExpectedDestination.get()
                    : (destination.state == FileOperations::ObjectBindingState::Bound && ! destinationIsSource ? destination.authority.boundObject.get()
                                                                                                               : nullptr);
            renameHr = sourceAuthority.boundObject->RenameIfUnchanged(
                destinationPath.c_str(), expectedDestination, itemFlags, &options, &mutation, renamedObject.put());
            if (SUCCEEDED(renameHr))
            {
                if (! renamedObject || mutation.sizeBytes != sizeof(mutation) || mutation.outcomeKnown == FALSE || mutation.mutationCommitted == FALSE ||
                    mutation.originalStillPresent != FALSE)
                {
                    renameHr = E_UNEXPECTED;
                }
                else
                {
                    FileOperations::ObjectBindingResult returned =
                        FileOperations::CaptureReturnedObjectAuthority(std::move(renamedObject), rename->endpoint.profileId);
                    bool sameObject      = false;
                    bool sameRevision    = false;
                    const HRESULT sameHr = returned.state == FileOperations::ObjectBindingState::Bound
                                               ? FileOperations::CrossCheckBoundObjectIdentity(sourceAuthority, returned.authority, sameObject, sameRevision)
                                               : returned.status;
                    if (FAILED(sameHr) || ! sameObject)
                    {
                        // The namespace mutation committed, but the provider did not return the exact
                        // authority needed by the next dependency step. Never turn that into a
                        // retryable name conflict or repeat the mutation.
                        renameHr = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
                    }
                    else
                    {
                        sourceAuthority = std::move(returned.authority);
                    }
                }
            }
            else if (mutation.sizeBytes != sizeof(mutation) || mutation.outcomeKnown == FALSE || mutation.mutationCommitted != FALSE)
            {
                renameHr = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
            }
        }

        if (SUCCEEDED(renameHr))
        {
            NoteLiveOutputPublished(destinationPath.native());
            ClearConflictPrompt(*this);
            return S_OK;
        }
        if (renameHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE))
        {
            ClearConflictPrompt(*this);
            return renameHr;
        }

        const wil::com_ptr<IFileSystemIO> fileSystemIo = QueryFileSystemIo(_fileSystem.get());
        const ConflictBucket bucket =
            ClassifyConflictBucket(FILESYSTEM_RENAME, itemFlags, fileSystemIo, renameHr, sourcePath.native(), destinationPath.native(), false);
        const bool allowRetry = IsRetryableConflictBucket(bucket);
        const ConflictPromptBeginResult promptBegin =
            BeginConflictPrompt(*this,
                                &cookie,
                                bucket,
                                renameHr,
                                sourcePath.native(),
                                destinationPath.native(),
                                allowRetry,
                                retryCount,
                                false,
                                destination.state == FileOperations::ObjectBindingState::Bound && ! destinationIsSource);
        ConflictAction action = promptBegin.action;
        if (promptBegin.ownsPrompt)
        {
            action = WaitForConflictDecision(*this, &cookie, promptBegin.decisionScope).first;
        }

        if (action == ConflictAction::Overwrite || action == ConflictAction::ReplaceReadOnly || action == ConflictAction::ReplaceLink)
        {
            if (destination.state != FileOperations::ObjectBindingState::Bound || destinationIsSource)
            {
                ClearConflictPrompt(*this);
                return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
            }
            grantedExpectedDestination = destination.authority.boundObject;
            itemFlags                  = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
            if (action == ConflictAction::ReplaceReadOnly)
            {
                itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
            }
            else if (action == ConflictAction::ReplaceLink)
            {
                itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_LINK);
            }
            continue;
        }
        if (action == ConflictAction::Retry)
        {
            ++retryCount;
            itemFlags = FILESYSTEM_FLAG_NONE;
            grantedExpectedDestination.reset();
            continue;
        }
        ClearConflictPrompt(*this);
        if (action == ConflictAction::Skip || action == ConflictAction::SkipAll)
        {
            return S_FALSE;
        }
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
}

HRESULT FolderWindow::FileOperationState::Task::ExecuteBatchRename() noexcept
{
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();
    if (! plans || plans->size() != 1u || ! _fileSystem)
    {
        return E_UNEXPECTED;
    }
    const auto* rename = std::get_if<FileOperations::RenamePlan>(&plans->front());
    const bool scheduledRename =
        rename != nullptr && (rename->origin == FileOperations::RenameOrigin::BatchRename || rename->origin == FileOperations::RenameOrigin::ChangeCase);
    if (! scheduledRename || ! rename->endpoint.pathIdentity.has_value() || rename->finalMappings.size() != _sourcePaths.size())
    {
        return E_UNEXPECTED;
    }

    {
        std::scoped_lock lock(_progressMutex);
        _progressTotalItems =
            static_cast<unsigned long>((std::min)(rename->finalMappings.size(), static_cast<size_t>(std::numeric_limits<unsigned long>::max())));
        _progressCompletedItems = 0u;
        PublishProgressCountersLocked(*this);
    }

#ifdef ENABLE_TESTS
    g_fileOpsBatchRenameBeforeExecutionPausePoint.Pause(5'000ull);
#endif
    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    constexpr FileSystemBindFlags sourceBindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_RENAME);
    constexpr FileSystemBindFlags destinationBindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION);
    const auto storePreMutationFailure = [&](const size_t failedIndex, const HRESULT status) noexcept
    {
        for (size_t resultIndex = 0u; resultIndex < rename->finalMappings.size(); ++resultIndex)
        {
            const FileOperations::RenameStep& resultStep = rename->finalMappings[resultIndex];
            FileOperations::FileOperationItemResult result{};
            result.sourceIndex  = resultIndex;
            result.strategy     = FileOperations::OperationStrategy::Native;
            result.publication  = resultIndex == failedIndex ? FileOperations::PublicationState::NotPublished : FileOperations::PublicationState::NotAttempted;
            result.verification = FileOperations::VerificationState::NotApplicable;
            result.sourceDisposition    = FileOperations::SourceDisposition::Retained;
            result.completion           = FileOperations::ItemCompletion::Failed;
            result.status               = resultIndex == failedIndex ? status : HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED);
            result.finalSourcePath      = resultStep.source.providerPath;
            result.finalDestinationPath = resultStep.providerJoinedPath;
            static_cast<void>(StoreTypedItemResult(std::move(result)));
        }
    };
    std::vector<FileOperations::BoundObjectAuthority> authorities;
    std::vector<BatchRenameExecutionOp> operations;
    authorities.reserve(rename->finalMappings.size());
    operations.reserve(rename->finalMappings.size());
    for (size_t index = 0u; index < rename->finalMappings.size(); ++index)
    {
        const FileOperations::RenameStep& step = rename->finalMappings[index];
        if (! step.source.ingressSnapshot.has_value())
        {
            storePreMutationFailure(index, E_UNEXPECTED);
            return E_UNEXPECTED;
        }
        FileOperations::ObjectBindingResult bound =
            FileOperations::BindObjectAuthority(_fileSystem.get(), step.source.providerPath, rename->endpoint.profileId, sourceBindFlags);
        if (bound.state != FileOperations::ObjectBindingState::Bound)
        {
            const HRESULT status = FAILED(bound.status) ? bound.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            storePreMutationFailure(index, status);
            return status;
        }
        const FileOperations::ProviderIdentitySnapshot& ingress = step.source.ingressSnapshot.value();
        if (ingress.pathProfileId != bound.authority.identity.pathProfileId || ingress.objectId != bound.authority.identity.objectId ||
            ingress.revisionId != bound.authority.identity.revisionId)
        {
            storePreMutationFailure(index, HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH));
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        const bool isDirectory = bound.authority.kind == FILESYSTEM_BOUND_DIRECTORY;
        authorities.push_back(std::move(bound.authority));
        const std::filesystem::path sourcePath(step.source.providerPath);
        std::wstring parentPath;
        std::wstring sourceLeaf;
        if (! TryGetFileSystemParentPath(rename->endpoint.pathIdentity.value(), step.source.providerPath, parentPath) ||
            ! TryGetFileSystemLeafName(rename->endpoint.pathIdentity.value(), step.source.providerPath, sourceLeaf))
        {
            storePreMutationFailure(index, E_INVALIDARG);
            return E_INVALIDARG;
        }
        const std::optional<std::wstring> parentKey = TryMakePathKey(rename->endpoint.pathIdentity.value(), parentPath);
        const FileSystemRouteContract::ChildNameContractResult sourceName =
            FileSystemRouteContract::QueryChildNameContract(_fileSystem.get(), parentPath, sourceLeaf, FILESYSTEM_RENAME, rename->endpoint.pluginId);
        const FileSystemRouteContract::ChildNameContractResult finalName =
            FileSystemRouteContract::QueryChildNameContract(_fileSystem.get(), parentPath, step.finalLeafName, FILESYSTEM_RENAME, rename->endpoint.pluginId);
        if (! parentKey.has_value() || sourceName.state != FileSystemRouteContract::QueryState::Available ||
            sourceName.nameStatus != FILESYSTEM_CHILD_NAME_VALID || sourceName.collisionKey.empty() ||
            finalName.state != FileSystemRouteContract::QueryState::Available || finalName.nameStatus != FILESYSTEM_CHILD_NAME_VALID ||
            parentKey.value() != step.providerParentKey || sourceName.collisionKey != step.providerSourceCollisionKey ||
            finalName.joinedPath != step.providerJoinedPath || finalName.collisionKey != step.providerCollisionKey)
        {
            const HRESULT status = sourceName.state == FileSystemRouteContract::QueryState::Available &&
                                           sourceName.nameStatus == FILESYSTEM_CHILD_NAME_INVALID && FAILED(sourceName.failureStatus)
                                       ? sourceName.failureStatus
                                   : FAILED(sourceName.status) ? sourceName.status
                                   : finalName.state == FileSystemRouteContract::QueryState::Available &&
                                           finalName.nameStatus == FILESYSTEM_CHILD_NAME_INVALID && FAILED(finalName.failureStatus)
                                       ? finalName.failureStatus
                                       : (FAILED(finalName.status) ? finalName.status : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH));
            storePreMutationFailure(index, status);
            return status;
        }
        operations.push_back(BatchRenameExecutionOp{
            .originalSource             = sourcePath,
            .finalLeaf                  = step.finalLeafName,
            .providerFinalPath          = step.providerJoinedPath,
            .providerParentKey          = parentKey.value(),
            .providerSourceCollisionKey = sourceName.collisionKey,
            .providerFinalCollisionKey  = finalName.collisionKey,
            .depth                      = PathDepthKey(sourcePath),
            .isDirectory                = isDirectory,
        });
    }

    // Batch Rename has an explicit preview/Run confirmation and a selection-proportional plan.
    // Recheck every final namespace slot once before the first mutation so a source that vanished
    // or an unrelated destination that appeared after preview cannot cause avoidable partial work.
    // Planned-source destinations remain valid dependency/cycle edges and are checked per hop.
    for (size_t index = 0u; index < operations.size(); ++index)
    {
        const BatchRenameExecutionOp& operation = operations[index];
        const bool destinationIsPlannedSource   = std::ranges::any_of(operations, [&](const BatchRenameExecutionOp& candidate) noexcept {
            return operation.providerParentKey == candidate.providerParentKey && operation.providerFinalCollisionKey == candidate.providerSourceCollisionKey;
        });
        if (destinationIsPlannedSource)
        {
            continue;
        }

        FileOperations::ObjectBindingResult destination =
            FileOperations::BindObjectAuthority(_fileSystem.get(), operation.providerFinalPath.native(), rename->endpoint.profileId, destinationBindFlags);
        if (destination.state == FileOperations::ObjectBindingState::Missing)
        {
            continue;
        }
        const HRESULT status = destination.state == FileOperations::ObjectBindingState::Bound
                                   ? HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)
                                   : (FAILED(destination.status) ? destination.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        storePreMutationFailure(index, status);
        return status;
    }

    struct ExecutionContext final
    {
        Task* task                                                     = nullptr;
        std::vector<FileOperations::BoundObjectAuthority>* authorities = nullptr;
    } context{.task = this, .authorities = &authorities};
    const auto mutation =
        [](void* raw, const size_t operationIndex, const std::filesystem::path& sourcePath, const std::filesystem::path& destinationPath) noexcept -> HRESULT
    {
        auto* const execution = static_cast<ExecutionContext*>(raw);
        return execution != nullptr && execution->task != nullptr && execution->authorities != nullptr
                   ? execution->task->ExecuteBatchRenameMutation(operationIndex, sourcePath, destinationPath, *execution->authorities)
                   : E_POINTER;
    };
    const auto progress = [](void* raw, const uint64_t completed, const uint64_t total, bool) noexcept
    {
        auto* const execution = static_cast<ExecutionContext*>(raw);
        if (execution == nullptr || execution->task == nullptr)
        {
            return;
        }
        Task& task                  = *execution->task;
        uint64_t publishedCompleted = 0u;
        uint64_t publishedTotal     = 0u;
        {
            std::scoped_lock lock(task._progressMutex);
            task._progressTotalItems     = static_cast<unsigned long>((std::min)(total, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max())));
            task._progressCompletedItems = static_cast<unsigned long>((std::min)(completed, static_cast<uint64_t>(task._progressTotalItems)));
            PublishProgressCountersLocked(task);
            publishedCompleted = task._progressCompletedItems;
            publishedTotal     = task._progressTotalItems;
        }
        if (task._externalProgressCallback)
        {
            task._externalProgressCallback(publishedCompleted, publishedTotal);
        }
    };

    const auto executionStartedAt        = std::chrono::steady_clock::now();
    BatchRenameExecutionResult execution = RunBatchRenameExecutionEngine(_cancelled,
                                                                         rename->endpoint.pathIdentity.value(),
                                                                         std::move(operations),
                                                                         BatchRenameExecutionOptions{
                                                                             .progressCallback = progress,
                                                                             .progressContext  = &context,
                                                                             .mutationCallback = mutation,
                                                                             .mutationContext  = &context,
                                                                             .schedule         = &rename->schedule,
                                                                         });
    if (rename->origin == FileOperations::RenameOrigin::ChangeCase)
    {
        const uint64_t completedCount = static_cast<uint64_t>(
            std::ranges::count_if(execution.itemResults, [](const BatchRenameExecutionItemResult& item) noexcept { return item.completed; }));
        Debug::Perf::Emit(L"changecase.execute.us",
                          L"central-rename-plan",
                          Debug::Perf::ElapsedUs(executionStartedAt),
                          static_cast<uint64_t>(execution.itemResults.size()),
                          completedCount,
                          execution.hr);
    }

    for (const BatchRenameExecutionItemResult& item : execution.itemResults)
    {
        if (item.operationIndex >= rename->finalMappings.size())
        {
            continue;
        }
        const FileOperations::RenameStep& step = rename->finalMappings[item.operationIndex];
        FileOperations::FileOperationItemResult result{};
        result.sourceIndex          = item.operationIndex;
        result.strategy             = FileOperations::OperationStrategy::Native;
        result.verification         = FileOperations::VerificationState::NotApplicable;
        result.status               = item.status;
        result.finalSourcePath      = step.source.providerPath;
        result.finalDestinationPath = step.providerJoinedPath;
        if (item.completed)
        {
            result.publication       = FileOperations::PublicationState::Published;
            result.sourceDisposition = FileOperations::SourceDisposition::Removed;
            result.completion        = FileOperations::ItemCompletion::Completed;
            StorePublishedTopLevelCompletionSnapshot(*this, MarkTopLevelItemCompleted(*this, item.operationIndex));
        }
        else if (item.skipped)
        {
            result.publication       = FileOperations::PublicationState::NotPublished;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Skipped;
        }
        else if (item.status == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE))
        {
            result.publication       = FileOperations::PublicationState::Unknown;
            result.sourceDisposition = FileOperations::SourceDisposition::Unknown;
            result.completion        = FileOperations::ItemCompletion::Indeterminate;
            result.status            = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        else if (IsCancellationOutcome(item.status))
        {
            result.status            = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            result.publication       = FileOperations::PublicationState::NotAttempted;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Canceled;
        }
        else
        {
            result.publication       = FileOperations::PublicationState::NotPublished;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Failed;
        }
        static_cast<void>(StoreTypedItemResult(std::move(result)));
    }
    return execution.hr;
}

namespace FolderWindowFileOperationsStateInternal
{
// Types the cross-file-system bridge shares with Task::ExecuteOperation. They were local to that
// function; they live here so the bridge can be a namespace-scope type (pure code motion).
using FileOperationState    = FolderWindow::FileOperationState;
using DiagnosticSeverity    = FolderWindow::FileOperationState::DiagnosticSeverity;
using PerItemCallbackCookie = Task::PerItemCallbackCookie;
using OwnedStageDisposition = FileOperations::OwnedStageDisposition;

enum class QualifiedItemFailurePhase : uint8_t
{
    None,
    DestinationParent,
    StageCreate,
    StageIdentity,
    StageWrite,
    StageCommit,
    FinalPublish,
    FinalPublishReconcile,
    ReplacementRollback,
    StageAbort,
    Verification,
    SourceCleanup,
    ProviderNative,
};

// Exact no-follow classification; a profile without object binding falls back to attributes.
[[nodiscard]] bool IsRegularDirectoryObject(IFileSystem& fileSystem, IFileSystemIO& io, std::wstring_view pathProfileId, const std::wstring& path) noexcept
{
    constexpr FileSystemBindFlags bindFlags           = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    const FileOperations::ObjectBindingResult binding = FileOperations::BindObjectAuthority(&fileSystem, path, pathProfileId, bindFlags);
    if (binding.state == FileOperations::ObjectBindingState::Bound)
    {
        return binding.authority.kind == FILESYSTEM_BOUND_DIRECTORY;
    }
    if (binding.state != FileOperations::ObjectBindingState::Unsupported)
    {
        return false;
    }
    unsigned long attributes = 0u;
    return SUCCEEDED(io.GetAttributes(path.c_str(), &attributes)) && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u &&
           (attributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0u;
}

struct CrossFileSystemBridge
{
    enum MutationAttemptFlag : uint32_t
    {
        kStageCreatedFlag          = 1u << 0u,
        kStageCleanupAttemptedFlag = 1u << 1u,
        kFinalPublishAttemptedFlag = 1u << 2u,
        kSourceDeleteAttemptedFlag = 1u << 3u,
    };

    static constexpr DWORD SleepSliceMs() noexcept
    {
        return 50u;
    }
    [[nodiscard]] DWORD ProgressIntervalMs() const noexcept
    {
        return progressPeriodMs;
    }

    struct BridgeCopyPerf final
    {
        uint64_t copyUs        = 0;
        uint64_t readerWaitUs  = 0;
        uint64_t writerWaitUs  = 0;
        uint64_t readUs        = 0;
        uint64_t writeUs       = 0;
        uint64_t progressCalls = 0;
    };

    Task& task;
    IFileSystem& sourceFs;
    IFileSystem& destinationFs;
    IFileSystemIO& sourceIo;
    IFileSystemIO& destinationIo;
    wil::com_ptr<IFileSystemObjectBinding> sourceBinding;
    HRESULT sourceBindingQueryHr = E_NOINTERFACE;
    wil::com_ptr<IFileSystemObjectBinding> destinationBinding;
    HRESULT destinationBindingQueryHr                  = E_NOINTERFACE;
    IFileSystemDirectoryOperations* destinationDirOps  = nullptr;
    unsigned int sourcePluginMaxConcurrencyBudget      = 1;
    unsigned int destinationPluginMaxConcurrencyBudget = 1;
    FileSystemFlags flags                              = FILESYSTEM_FLAG_NONE;
    void* cookie                                       = nullptr;
    DWORD sourceRootAttributesHint                     = 0;
    ReparsePointPolicy reparsePointPolicy              = ReparsePointPolicy::Preserve;
    bool destinationUsesOrdinalIgnoreCaseComponents    = false;
    bool destinationUsesWindowsChildNameRules          = false;
    // Rename merge: a Native directory Move whose destination already holds a regular directory.
    // Every child relocates through one provider Native rename; nothing is copied or staged.
    bool renameMerge = false;
    std::atomic<bool> rootSourceRemoved{false};

    // Total bytes is best-effort: if unknown, keep 0.
    uint64_t totalBytes                        = 0;
    uint64_t completedBytes                    = 0;
    unsigned long skippedDirectoryReparseCount = 0;
    unsigned long skippedFileReparseCount      = 0;
    bool rootDirectoryReparseSkipped           = false;
    bool unsupportedReparseEncountered         = false;
    std::atomic<bool> anyDestinationPublished{false};
    std::atomic<bool> destinationPublicationUnknown{false};
    // Per-file conflicts answered Skip: the file never reached the destination, so for a
    // MOVE the source stays authoritative (no delete) and the transfer ends PARTIAL.
    std::atomic<uint64_t> skippedFileConflictCount{0};

    std::mutex callbackMutex;
    std::mutex throttleMutex;
    std::atomic<uint64_t> bandwidthLimitBytesPerSecond{0};
    std::atomic<bool> managedSourceRetained{false};
    std::atomic<uint64_t> managedSourceRetentionGeneration{0};
    std::atomic<bool> managedCleanupIndeterminate{false};
    std::atomic<bool> ownedStageCleanupIndeterminate{false};
    std::atomic<QualifiedItemFailurePhase> failurePhase{QualifiedItemFailurePhase::None};
    std::atomic<HRESULT> failureStatus{S_OK};
    std::atomic<uint32_t> mutationAttemptFlags{0u};
    std::atomic<OwnedStageDisposition> ownedStageDisposition{OwnedStageDisposition::NotCreated};
    std::wstring sourcePathProfileId;
    std::wstring destinationPathProfileId;
    std::wstring rootSourcePath;
    std::wstring rootDestinationPath;
    FileSystemPathIdentity sourcePathIdentity;
    FileSystemPathIdentity destinationPathIdentity;

    std::mutex traversalBudgetMutex;
    uint64_t traversalRetainedEntries       = 0;
    uint64_t traversalRetainedPathBytes     = 0;
    uint64_t traversalRetainedMetadataBytes = 0;

    struct ManagedSourceCleanupRecord final
    {
        FileOperations::BoundObjectAuthority authority{};
        bool armed              = false;
        unsigned int retryCount = 0;
    };

    void NoteFailure(QualifiedItemFailurePhase phase, HRESULT status) noexcept
    {
        if (! FAILED(status) || phase == QualifiedItemFailurePhase::None)
        {
            return;
        }
        QualifiedItemFailurePhase expected = QualifiedItemFailurePhase::None;
        if (failurePhase.compare_exchange_strong(expected, phase, std::memory_order_acq_rel))
        {
            failureStatus.store(status, std::memory_order_release);
        }
    }

    struct PreparedLinkRecord final
    {
        std::wstring sourcePath;
        std::wstring destinationPath;
        ManagedSourceCleanupRecord cleanupRecord{};
        FileOperations::BoundObjectAuthority copyAuthority{};
        wil::com_ptr<IFileSystemBoundObject> expectedDestination;
        FileSystemBasicInformation basicInformation{};
        bool hasBasicInformation    = false;
        bool overwriteGranted       = false;
        bool replaceReadOnlyGranted = false;
        bool replaceLinkGranted     = false;

        [[nodiscard]] FileOperations::BoundObjectAuthority* ReadAuthority() noexcept
        {
            return cleanupRecord.armed ? &cleanupRecord.authority : &copyAuthority;
        }
    };

    struct ConnectionLimit final
    {
        std::wstring id;
        uint32_t maxCopyMove = 1;
    };
    std::optional<ConnectionLimit> sourceConnectionLimit;
    std::optional<ConnectionLimit> destinationConnectionLimit;
    bool connectionLimitsInitialized = false;

    std::atomic<uint64_t>* completedBytesAtomic = nullptr;

    ULONGLONG startTick = 0;
    FileSystemOptions options{};

    CrossFsBridgeBufferLease bufferBudgetLease;
    std::unique_ptr<std::byte[]> buffer;
    unsigned long bufferBytes            = 0;
    HRESULT bufferAllocationHr           = S_OK;
    DWORD progressPeriodMs               = 200u;
    uint32_t transferLatencyClass        = FILESYSTEM_TRANSFER_LATENCY_UNKNOWN;
    uint32_t transferHintFlags           = FILESYSTEM_TRANSFER_HINT_NONE;
    bool sourceCleanupPermitted          = false;
    bool verificationRequested           = false;
    bool verificationHostReadback        = false;
    bool verificationProviderBlake3Proof = false;
    bool verificationWriterDigestProof   = false; // R3-2
    bool rootVerificationNotApplicable   = false;
    std::atomic<FileOperations::VerificationState> verificationState{FileOperations::VerificationState::NotRequested};
    std::atomic<uint64_t> verificationCompletedBytes{0u};
    std::atomic<uint64_t> verificationFileCount{0u};
    uint64_t discoveredBytes       = 0u;
    uint64_t discoveredFiles       = 0u;
    uint64_t discoveredDirectories = 0u;

    CrossFileSystemBridge(Task& owner,
                          IFileSystem& source,
                          IFileSystem& destination,
                          IFileSystemIO& sourceIoIn,
                          IFileSystemIO& destinationIoIn,
                          IFileSystemDirectoryOperations* destinationDirOpsIn,
                          unsigned int sourcePluginMaxConcurrencyBudgetIn,
                          unsigned int destinationPluginMaxConcurrencyBudgetIn,
                          FileSystemFlags flagsIn,
                          void* cookieIn,
                          uint64_t totalBytesIn,
                          const wchar_t* rootSourcePathIn,
                          const wchar_t* rootDestinationPathIn,
                          DWORD sourceRootAttributesHintIn,
                          ReparsePointPolicy reparsePointPolicyIn,
                          std::wstring_view sourcePathProfileIdIn,
                          std::wstring_view destinationPathProfileIdIn,
                          const FileSystemPathIdentity& sourcePathIdentityIn,
                          const FileSystemPathIdentity& destinationPathIdentityIn,
                          bool sourceCleanupPermittedIn,
                          bool verificationRequestedIn,
                          bool verificationHostReadbackIn,
                          bool verificationProviderBlake3ProofIn,
                          bool verificationWriterDigestProofIn,
                          bool renameMergeIn = false) noexcept
        : task(owner),
          sourceFs(source),
          destinationFs(destination),
          sourceIo(sourceIoIn),
          destinationIo(destinationIoIn),
          destinationDirOps(destinationDirOpsIn),
          sourcePluginMaxConcurrencyBudget(std::max(1u, sourcePluginMaxConcurrencyBudgetIn)),
          destinationPluginMaxConcurrencyBudget(std::max(1u, destinationPluginMaxConcurrencyBudgetIn)),
          flags(flagsIn),
          cookie(cookieIn),
          sourceRootAttributesHint(sourceRootAttributesHintIn),
          reparsePointPolicy(reparsePointPolicyIn),
          renameMerge(renameMergeIn),
          totalBytes(totalBytesIn),
          sourcePathProfileId(sourcePathProfileIdIn),
          destinationPathProfileId(destinationPathProfileIdIn),
          rootSourcePath(rootSourcePathIn ? rootSourcePathIn : L""),
          rootDestinationPath(rootDestinationPathIn ? rootDestinationPathIn : L""),
          sourcePathIdentity(sourcePathIdentityIn),
          destinationPathIdentity(destinationPathIdentityIn),
          sourceCleanupPermitted(sourceCleanupPermittedIn),
          verificationRequested(verificationRequestedIn),
          verificationHostReadback(verificationHostReadbackIn),
          verificationProviderBlake3Proof(verificationProviderBlake3ProofIn),
          verificationWriterDigestProof(verificationWriterDigestProofIn),
          verificationState(verificationRequestedIn ? FileOperations::VerificationState::NotApplicable : FileOperations::VerificationState::NotRequested)
    {
        if (verificationRequested)
        {
            task._verificationRequested.store(true, std::memory_order_release);
        }
        sourceBindingQueryHr                       = sourceFs.QueryInterface(__uuidof(IFileSystemObjectBinding), sourceBinding.put_void());
        destinationBindingQueryHr                  = destinationFs.QueryInterface(__uuidof(IFileSystemObjectBinding), destinationBinding.put_void());
        destinationUsesOrdinalIgnoreCaseComponents = UsesOrdinalIgnoreCaseComponents(destinationFs, rootDestinationPathIn, task._destinationPluginId);
        destinationUsesWindowsChildNameRules       = destinationUsesOrdinalIgnoreCaseComponents ||
                                                     NavigationLocation::EqualsNoCase(task._destinationPluginId, L"builtin/file-system") ||
                                                     NavigationLocation::EqualsNoCase(task._destinationPluginId, L"builtin/file-system-dummy");
        const uint64_t initialBandwidth            = task._desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        task.InitializeFileSystemOptions(options, cookie);
        options.bandwidthLimitBytesPerSecond = initialBandwidth;
        options.copyMoveMaxConcurrency       = std::min(sourcePluginMaxConcurrencyBudget, destinationPluginMaxConcurrencyBudget);
        bandwidthLimitBytesPerSecond.store(initialBandwidth, std::memory_order_release);
        if (renameMerge)
        {
            return;
        }

        const AdaptiveBridgeTuning tuning = ResolveAdaptiveCrossFsBridgeTuning(
            task._crossFsBridgeBufferBytes, sourceFs, rootSourcePathIn, destinationFs, rootDestinationPathIn, task._operation);
        bufferBytes          = tuning.bufferBytes;
        progressPeriodMs     = tuning.progressPeriodMs;
        transferLatencyClass = tuning.latencyClass;
        transferHintFlags    = tuning.flags;
        task._resolvedCrossFsBridgeBufferBytes.store(bufferBytes, std::memory_order_release);
        const uint64_t reservationBytes = static_cast<uint64_t>(bufferBytes) * 2ull;
        if (! bufferBudgetLease.Acquire(reservationBytes, task._cancelled, task._stopToken))
        {
            bufferAllocationHr =
                (task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested()) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : E_OUTOFMEMORY;
            return;
        }
        buffer.reset(new (std::nothrow) std::byte[bufferBytes]);
        if (! buffer)
        {
            bufferBudgetLease.Reset();
            bufferAllocationHr = E_OUTOFMEMORY;
        }
    }

    CrossFileSystemBridge(const CrossFileSystemBridge&)            = delete;
    CrossFileSystemBridge(CrossFileSystemBridge&&)                 = delete;
    CrossFileSystemBridge& operator=(const CrossFileSystemBridge&) = delete;
    CrossFileSystemBridge& operator=(CrossFileSystemBridge&&)      = delete;

    [[nodiscard]] bool TryGetSourceRelativePath(std::wstring_view sourcePath, std::wstring& relativeOut) const noexcept
    {
        return TryGetFileSystemRelativePath(sourcePathIdentity, rootSourcePath, sourcePath, relativeOut);
    }

    [[nodiscard]] bool CancelRequested() const noexcept
    {
        return task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested();
    }

    [[nodiscard]] bool DiscoveryAheadEnabled() noexcept
    {
        FileSystemDiscoveryMode mode = FILESYSTEM_DISCOVERY_JUST_IN_TIME;
        return SUCCEEDED(task.FileSystemGetDiscoveryMode(&mode, cookie)) && mode == FILESYSTEM_DISCOVERY_AHEAD;
    }

    [[nodiscard]] HRESULT ReportDiscovery(uint32_t queuedItems, bool traversalClosed) noexcept
    {
        FileSystemDiscoveryProgress progress{};
        progress.sizeBytes             = sizeof(progress);
        progress.discoveredBytes       = discoveredBytes;
        progress.discoveredFiles       = discoveredFiles;
        progress.discoveredDirectories = discoveredDirectories;
        progress.queuedItems           = queuedItems;
        progress.traversalClosed       = traversalClosed ? TRUE : FALSE;
        return task.FileSystemReportDiscoveryProgress(&progress, cookie);
    }

    [[nodiscard]] static uint64_t RetainedPathBytes(std::wstring_view sourcePath, std::wstring_view destinationPath) noexcept
    {
        constexpr uint64_t kWcharBytes  = sizeof(wchar_t);
        const uint64_t sourceBytes      = static_cast<uint64_t>(sourcePath.size()) * kWcharBytes;
        const uint64_t destinationBytes = static_cast<uint64_t>(destinationPath.size()) * kWcharBytes;
        constexpr uint64_t kMax         = std::numeric_limits<uint64_t>::max();
        if (sourceBytes > kMax - Common::FileOperations::kTraversalRecordOverheadBytes ||
            destinationBytes > kMax - Common::FileOperations::kTraversalRecordOverheadBytes - sourceBytes)
        {
            return kMax;
        }
        return sourceBytes + destinationBytes + Common::FileOperations::kTraversalRecordOverheadBytes;
    }

    [[nodiscard]] static uint64_t AddRetainedBytes(uint64_t left, uint64_t right) noexcept
    {
        return left > std::numeric_limits<uint64_t>::max() - right ? std::numeric_limits<uint64_t>::max() : left + right;
    }

    [[nodiscard]] HRESULT ReportTraversalLimit(std::wstring_view code,
                                               std::wstring_view limitName,
                                               uint64_t limit,
                                               const std::wstring& sourcePath,
                                               const std::wstring& destinationPath,
                                               HRESULT status) noexcept
    {
        task._bridgeTraversalLimitHitCount.fetch_add(1u, std::memory_order_acq_rel);
        task.LogDiagnostic(
            FileOperationState::DiagnosticSeverity::Error,
            status,
            std::wstring(code),
            std::format(L"Traversal stopped at this item after reaching the {} limit ({:L}). Prior completed items remain published.", limitName, limit),
            sourcePath,
            destinationPath);
        return status;
    }

    [[nodiscard]] HRESULT ValidateTraversalDepth(uint64_t depth, const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
    {
        AtomicMax(task._bridgeTraversalMaxDepth, depth);
        const uint64_t maxDepth = GetFileOpsTraversalMaxDepth();
        if (depth <= maxDepth)
        {
            return S_OK;
        }
        return ReportTraversalLimit(
            L"bridge.traversal.depthLimit", L"walk depth", maxDepth, sourcePath, destinationPath, HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW));
    }

    [[nodiscard]] HRESULT ReserveTraversalWork(uint64_t pathBytes, const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
    {
        enum class Rejection : uint8_t
        {
            None,
            EntryCount,
            PathBytes,
        };
        Rejection rejection = Rejection::None;
        {
            std::scoped_lock lock(traversalBudgetMutex);
            if (traversalRetainedEntries >= Common::FileOperations::kTraversalMaxQueuedEntries)
            {
                rejection = Rejection::EntryCount;
            }
            else if (pathBytes > Common::FileOperations::kTraversalMaxQueuedPathBytes ||
                     traversalRetainedPathBytes > Common::FileOperations::kTraversalMaxQueuedPathBytes - pathBytes)
            {
                rejection = Rejection::PathBytes;
            }
            else
            {
                ++traversalRetainedEntries;
                traversalRetainedPathBytes += pathBytes;
                AtomicMax(task._bridgeTraversalMaxRetainedEntries, traversalRetainedEntries);
                AtomicMax(task._bridgeTraversalMaxQueuedPathBytes, traversalRetainedPathBytes);
                return S_OK;
            }
        }

        const bool entryLimit = rejection == Rejection::EntryCount;
        return ReportTraversalLimit(entryLimit ? L"bridge.traversal.entryLimit" : L"bridge.traversal.pathBudget",
                                    entryLimit ? L"queued work entries" : L"queued UTF-16 path bytes",
                                    entryLimit ? Common::FileOperations::kTraversalMaxQueuedEntries : Common::FileOperations::kTraversalMaxQueuedPathBytes,
                                    sourcePath,
                                    destinationPath,
                                    HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY));
    }

    void ReleaseTraversalWork(uint64_t pathBytes) noexcept
    {
        std::scoped_lock lock(traversalBudgetMutex);
        if (traversalRetainedEntries > 0u)
        {
            --traversalRetainedEntries;
        }
        traversalRetainedPathBytes = (traversalRetainedPathBytes >= pathBytes) ? (traversalRetainedPathBytes - pathBytes) : 0u;
    }

    // Per-directory retained state (the provider's listing buffer, the frame's index of child-name
    // views, its path text) is released when the frame pops and reported as a high-water mark; it is
    // not a ceiling (R4-T3): refusing a listing the provider has already produced frees nothing, and
    // the frame stack keeps the live set at O(depth) directories.
    void NoteTraversalMetadataRetained(uint64_t bytes) noexcept
    {
        std::scoped_lock lock(traversalBudgetMutex);
        traversalRetainedMetadataBytes = AddRetainedBytes(traversalRetainedMetadataBytes, bytes);
        AtomicMax(task._bridgeTraversalMaxMetadataBytes, traversalRetainedMetadataBytes);
    }

    void ReleaseTraversalMetadata(uint64_t bytes) noexcept
    {
        std::scoped_lock lock(traversalBudgetMutex);
        traversalRetainedMetadataBytes = (traversalRetainedMetadataBytes >= bytes) ? (traversalRetainedMetadataBytes - bytes) : 0u;
    }

    [[nodiscard]] std::wstring SourceKeptReason(std::wstring_view detail) const
    {
        return std::format(L"{} because {}", renameMerge ? L"Moved; source folder kept" : L"Copied; source kept", detail);
    }

    void MarkManagedSourceRetained(const std::wstring& sourcePath, const std::wstring& destinationPath, HRESULT status, std::wstring_view reason) noexcept
    {
        if (! sourceCleanupPermitted)
        {
            return;
        }
        managedSourceRetained.store(true, std::memory_order_release);
        managedSourceRetentionGeneration.fetch_add(1, std::memory_order_acq_rel);
        task.LogDiagnostic(
            FileOperationState::DiagnosticSeverity::Warning, status, L"bridge.move.sourceRetained", std::wstring(reason), sourcePath, destinationPath);
    }

    [[nodiscard]] DeferredConsentResult RequestBridgeConsent(FileOperations::DeferredConsentRisk risk,
                                                             HRESULT status,
                                                             const std::wstring& sourcePath,
                                                             const std::wstring& destinationPath,
                                                             const FileOperations::BoundObjectAuthority* sourceAuthority,
                                                             const FileSystemMetadataSnapshot* metadataSnapshot) noexcept
    {
        const auto* perItemCookie = static_cast<const Task::PerItemCallbackCookie*>(cookie);
        DeferredConsentRequest request{};
        request.risk               = risk;
        request.status             = status;
        request.perItemCookie      = perItemCookie;
        request.sourcePath         = sourcePath;
        request.destinationPath    = destinationPath;
        request.itemIndex          = perItemCookie ? std::optional<size_t>(perItemCookie->itemIndex) : std::nullopt;
        request.destinationRootId  = rootDestinationPath;
        request.applyToAllEligible = true;
        request.itemCountKnown     = true;
        request.itemCount          = 1u;
        request.bytesKnown         = metadataSnapshot != nullptr && metadataSnapshot->logicalSizeBytes != std::numeric_limits<uint64_t>::max();
        request.bytes              = request.bytesKnown ? metadataSnapshot->logicalSizeBytes : 0u;
        if (sourceAuthority != nullptr)
        {
            request.sourceIdentity = sourceAuthority->identity;
        }
        return RequestDeferredConsent(task, request);
    }

    void ReportMetadataOutcome(const std::wstring& sourcePath, const std::wstring& destinationPath, const FileSystemMetadataTransferResult& result) noexcept
    {
        struct Feature final
        {
            uint32_t mask;
            std::wstring_view name;
        };
        constexpr std::array features{
            Feature{FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES, L"basic"},
            Feature{FILESYSTEM_METADATA_MOTW, L"motw"},
            Feature{FILESYSTEM_METADATA_ALTERNATE_STREAMS, L"alternateStreams"},
            Feature{FILESYSTEM_METADATA_EXTENDED_ATTRIBUTES, L"extendedAttributes"},
            Feature{FILESYSTEM_METADATA_SECURITY, L"security"},
            Feature{FILESYSTEM_METADATA_SPARSE, L"sparse"},
            Feature{FILESYSTEM_METADATA_COMPRESSION, L"compression"},
            Feature{FILESYSTEM_METADATA_EFS, L"efs"},
            Feature{FILESYSTEM_METADATA_PLACEHOLDER, L"placeholder"},
        };
        for (const Feature& feature : features)
        {
            if ((result.lostFeatures & feature.mask) != 0u)
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   FAILED(result.firstFailure) ? result.firstFailure : S_FALSE,
                                   std::format(L"bridge.metadata.{}.lost", feature.name),
                                   std::format(L"Destination did not preserve {} metadata.", feature.name),
                                   sourcePath,
                                   destinationPath);
            }
            if ((result.changedFeatures & feature.mask) != 0u)
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   FAILED(result.firstFailure) ? result.firstFailure : S_FALSE,
                                   std::format(L"bridge.metadata.{}.changed", feature.name),
                                   std::format(L"Destination inherited or changed {} metadata.", feature.name),
                                   sourcePath,
                                   destinationPath);
            }
        }
    }

    // Cleanup seam: retains exact source authority before transfer and is the only path
    // allowed to remove a Managed Move source after publication succeeds.
    [[nodiscard]] HRESULT PrepareManagedSourceAuthority(const std::wstring& sourcePath,
                                                        const std::wstring& destinationPath,
                                                        FileSystemBoundObjectKind expectedKind,
                                                        ManagedSourceCleanupRecord& cleanupRecord) noexcept
    {
        cleanupRecord = {};
        if (! sourceCleanupPermitted)
        {
            return S_FALSE;
        }
        if (renameMerge && ! sourceBinding)
        {
            // No object binding on this profile: the emptied source directory is removed by name.
            return S_FALSE;
        }
        // R3-2: a route without object binding may still prove the published content through its
        // writer; the cleanup record stays armed and the proof decides before any source removal.
        if (! destinationBinding && ! (verificationWriterDigestProof && expectedKind == FILESYSTEM_BOUND_REGULAR_FILE))
        {
            MarkManagedSourceRetained(sourcePath,
                                      destinationPath,
                                      HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                      L"Copied; source kept because exact destination publication authority is unavailable.");
            return S_FALSE;
        }

        FileSystemBindFlags bindFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_DELETE);
        if (expectedKind == FILESYSTEM_BOUND_REGULAR_FILE)
        {
            bindFlags = static_cast<FileSystemBindFlags>(bindFlags | FILESYSTEM_BIND_READ_CONTENT);
        }
        FileOperations::ObjectBindingResult binding = FileOperations::BindObjectAuthority(&sourceFs, sourcePath, sourcePathProfileId, bindFlags);
        if (binding.state == FileOperations::ObjectBindingState::Bound)
        {
            if (binding.authority.kind != expectedKind)
            {
                return HRESULT_FROM_WIN32(ERROR_FILE_INVALID);
            }
            cleanupRecord.authority = std::move(binding.authority);
            cleanupRecord.armed     = true;
            return S_OK;
        }
        if (binding.state == FileOperations::ObjectBindingState::Missing)
        {
            return FAILED(binding.status) ? binding.status : HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        if (binding.state == FileOperations::ObjectBindingState::ProviderContractViolation)
        {
            return FAILED(binding.status) ? binding.status : E_UNEXPECTED;
        }

        MarkManagedSourceRetained(sourcePath,
                                  destinationPath,
                                  FAILED(binding.status) ? binding.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                  L"Copied; source kept because exact source binding is unavailable for this item.");
        return S_FALSE;
    }

    [[nodiscard]] HRESULT FinalizeManagedSourceCleanup(const std::wstring& sourcePath,
                                                       const std::wstring& destinationPath,
                                                       ManagedSourceCleanupRecord& cleanupRecord) noexcept
    {
        if (! cleanupRecord.armed)
        {
            return S_OK;
        }
        if (! cleanupRecord.authority.boundObject)
        {
            return E_UNEXPECTED;
        }

        if (CancelRequested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

#ifdef ENABLE_TESTS
        MaybePauseBeforeBridgeMoveSourceCleanupForSelfTest();
#endif

        for (;;)
        {
            if (CancelRequested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            mutationAttemptFlags.fetch_or(kSourceDeleteAttemptedFlag, std::memory_order_acq_rel);
            FileSystemConditionalMutationResult result{};
            result.sizeBytes = sizeof(result);
            HRESULT deleteHr = E_UNEXPECTED;
#ifdef ENABLE_TESTS
            if (ConsumeBridgeCounterForSelfTest(g_fileOpsManagedCleanupUnknownOutcomeCount, g_fileOpsManagedCleanupUnknownOutcomeAttempts))
            {
                result.outcomeKnown         = FALSE;
                result.mutationCommitted    = FALSE;
                result.originalStillPresent = TRUE;
                deleteHr                    = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
            }
            else if (ConsumeBridgeCounterForSelfTest(g_fileOpsManagedCleanupKnownNonCommitCount, g_fileOpsManagedCleanupKnownNonCommitAttempts))
            {
                result.outcomeKnown         = TRUE;
                result.mutationCommitted    = FALSE;
                result.originalStillPresent = TRUE;
                deleteHr                    = g_fileOpsManagedCleanupKnownNonCommitStatus.load(std::memory_order_acquire);
            }
            else
#endif
            {
                deleteHr = cleanupRecord.authority.boundObject->DeleteIfUnchanged(FILESYSTEM_FLAG_NONE, &options, &result);
            }
            const FileOperations::ManagedCleanupAttemptDisposition disposition = FileOperations::ClassifyManagedCleanupMutation({
                .status               = deleteHr,
                .outcomeKnown         = result.outcomeKnown != FALSE,
                .mutationCommitted    = result.mutationCommitted != FALSE,
                .originalStillPresent = result.originalStillPresent != FALSE,
            });
            switch (disposition)
            {
                case FileOperations::ManagedCleanupAttemptDisposition::Removed:
                    if (FAILED(deleteHr))
                    {
                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                           deleteHr,
                                           L"bridge.move.cleanupCommittedWithError",
                                           L"The provider reported an error but confirmed that the exact source is no longer present.",
                                           sourcePath,
                                           destinationPath);
                    }
                    cleanupRecord.armed     = false;
                    cleanupRecord.authority = {};
                    return S_OK;
                case FileOperations::ManagedCleanupAttemptDisposition::Indeterminate:
                    managedCleanupIndeterminate.store(true, std::memory_order_release);
                    NoteFailure(QualifiedItemFailurePhase::SourceCleanup, deleteHr);
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                       deleteHr,
                                       L"bridge.move.cleanupIndeterminate",
                                       L"Managed Move source cleanup has an unknown outcome; it was not retried.",
                                       sourcePath,
                                       destinationPath);
                    return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
                case FileOperations::ManagedCleanupAttemptDisposition::ProviderContractViolation:
                    managedCleanupIndeterminate.store(true, std::memory_order_release);
                    NoteFailure(QualifiedItemFailurePhase::SourceCleanup, E_UNEXPECTED);
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                       E_UNEXPECTED,
                                       L"bridge.move.cleanupContractViolation",
                                       L"Managed Move source cleanup returned an inconsistent conditional-mutation result.",
                                       sourcePath,
                                       destinationPath);
                    return E_UNEXPECTED;
                case FileOperations::ManagedCleanupAttemptDisposition::Retained:
                default: break;
            }

            const std::optional<DWORD> cleanupError = Win32ErrorFromHRESULT(deleteHr);
            if (cleanupRecord.authority.kind == FILESYSTEM_BOUND_DIRECTORY && cleanupError.has_value() &&
                cleanupError.value() == static_cast<DWORD>(ERROR_DIR_NOT_EMPTY))
            {
                // The directory was enumerated before its published children were cleaned
                // up. A new or otherwise unselected child may now exist. It was never
                // published, so prompting Retry/Skip as if this were a failed leaf delete is
                // misleading and retrying could later remove a directory whose membership
                // changed after discovery. Retain the exact source directory and propagate
                // the task's source-kept outcome through the shared retention generation.
                MarkManagedSourceRetained(
                    sourcePath, destinationPath, deleteHr, SourceKeptReason(L"the source directory gained or retained an unselected child."));
                cleanupRecord.armed     = false;
                cleanupRecord.authority = {};
                return S_OK;
            }

            const ConflictBucket bucket                 = ClassifyManagedCleanupConflictBucket(deleteHr);
            const bool allowRetry                       = IsRetryableConflictBucket(bucket) && cleanupRecord.retryCount == 0u;
            const ConflictPromptBeginResult promptBegin = BeginConflictPrompt(
                task, static_cast<PerItemCallbackCookie*>(cookie), bucket, deleteHr, sourcePath, destinationPath, allowRetry, cleanupRecord.retryCount, false);
            ConflictAction action = promptBegin.action;
            if (promptBegin.ownsPrompt)
            {
                action = WaitForConflictDecision(task, cookie, promptBegin.decisionScope).first;
            }

            if (action == ConflictAction::Retry && allowRetry)
            {
                ++cleanupRecord.retryCount;
                continue;
            }
            if (action == ConflictAction::Skip || action == ConflictAction::SkipAll)
            {
                MarkManagedSourceRetained(
                    sourcePath, destinationPath, deleteHr, L"Copied; source kept because exact conditional source cleanup did not commit.");
                cleanupRecord.armed     = false;
                cleanupRecord.authority = {};
                return S_OK;
            }
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    [[nodiscard]] HRESULT ValidateChildNameForDestination(const std::wstring_view name) const noexcept
    {
        HRESULT hr = ValidateBridgeStructuralChildName(name);
        if (SUCCEEDED(hr) && destinationUsesWindowsChildNameRules)
        {
            hr = ValidateWindowsBridgeChildName(name);
        }
        return hr;
    }

    [[nodiscard]] BridgeChildNameSet MakeChildNameSet() const noexcept
    {
        return BridgeChildNameSet(BridgeChildNameLess{.ignoreCase = destinationUsesOrdinalIgnoreCaseComponents});
    }

    // The second occurrence of a name the destination cannot tell apart (an exact duplicate, or a
    // case-only variant when the destination folds case) is refused. The registered view points
    // into the frame's listing buffer, which outlives the set.
    [[nodiscard]] HRESULT ValidateAndRegisterChildName(const std::wstring_view name, BridgeChildNameSet& registeredNames) noexcept
    {
        const HRESULT hr = ValidateChildNameForDestination(name);
        if (FAILED(hr))
        {
            return hr;
        }
        if (! registeredNames.insert(name).second)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
        return S_OK;
    }

    void NoteInvalidEnumeratedChildName(const std::wstring& sourceParent, const std::wstring& destinationParent) noexcept
    {
        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                           HRESULT_FROM_WIN32(ERROR_INVALID_NAME),
                           L"bridge.source.invalidChildName",
                           L"Source provider returned a child name that is unsafe for the destination component contract.",
                           sourceParent,
                           destinationParent);
    }

    [[nodiscard]] HRESULT MakeTempDestinationPath(std::wstring_view destinationPath, uint64_t progressStreamId, std::wstring& stagePath) const noexcept
    {
        // The name carries 128 bits of CSPRNG entropy so a local attacker cannot pre-create
        // or race the staging file (PID/TID/tick names are predictable); the stream id stays
        // for diagnostic correlation only.
        wchar_t suffix[96]{};
        constexpr size_t suffixMax = (sizeof(suffix) / sizeof(suffix[0])) - 1u;

        uint64_t random[2]{};
#ifdef ENABLE_TESTS
        if (ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeFailNextStageEntropyCount, g_fileOpsBridgeFailNextStageEntropyAttempts))
        {
            return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }
#endif
        const NTSTATUS randomStatus = BCryptGenRandom(nullptr, reinterpret_cast<PUCHAR>(random), sizeof(random), BCRYPT_USE_SYSTEM_PREFERRED_RNG);
        if (! BCRYPT_SUCCESS(randomStatus))
        {
            Debug::Error(L"CrossFileSystemBridge: BCryptGenRandom failed before exclusive stage creation (status=0x{:08X}).",
                         static_cast<unsigned long>(randomStatus));
            return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }

        const auto r         = std::format_to_n(suffix,
                                                suffixMax,
                                                L".rs_tmp_{:016X}{:016X}_{:X}",
                                                static_cast<unsigned long long>(random[0]),
                                                static_cast<unsigned long long>(random[1]),
                                                progressStreamId);
        const size_t written = (r.size < suffixMax) ? static_cast<size_t>(r.size) : suffixMax;
        suffix[written]      = L'\0';

        stagePath.assign(destinationPath);
        stagePath.append(suffix);
        return S_OK;
    }

    [[nodiscard]] static FileSystemFlags BuildPublicationFlags(bool overwriteGranted, bool replaceReadOnlyGranted, bool replaceLinkGranted) noexcept
    {
        uint32_t value = FILESYSTEM_FLAG_NONE;
        if (overwriteGranted)
        {
            value |= FILESYSTEM_FLAG_ALLOW_OVERWRITE;
        }
        if (replaceReadOnlyGranted)
        {
            value |= FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY;
        }
        if (replaceLinkGranted)
        {
            value |= FILESYSTEM_FLAG_ALLOW_REPLACE_LINK;
        }
        return static_cast<FileSystemFlags>(value);
    }

    [[nodiscard]] HRESULT BindExpectedPublicationDestination(const std::wstring& destinationPath,
                                                             bool overwriteGranted,
                                                             wil::com_ptr<IFileSystemBoundObject>& expectedDestination) noexcept
    {
        expectedDestination.reset();
        if (! destinationBinding)
        {
            return destinationBindingQueryHr == E_NOINTERFACE ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                                                              : (FAILED(destinationBindingQueryHr) ? destinationBindingQueryHr : E_UNEXPECTED);
        }
        if (! overwriteGranted)
        {
            return S_OK;
        }

        constexpr FileSystemBindFlags bindFlags =
            static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_PUBLICATION);
        const HRESULT bindHr = destinationBinding->BindObject(destinationPath.c_str(), bindFlags, expectedDestination.put());
        if (IsMissingPathHr(bindHr))
        {
            expectedDestination.reset();
            return S_OK;
        }
        if (FAILED(bindHr))
        {
            return bindHr;
        }
        if (! expectedDestination)
        {
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                               E_UNEXPECTED,
                               L"bridge.binding.nullDestination",
                               L"Destination binding returned success without exact publication authority.",
                               L"",
                               destinationPath);
            return E_UNEXPECTED;
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT CreateOwnedStage(const std::wstring& destinationPath,
                                           uint64_t progressStreamId,
                                           std::wstring& stagePath,
                                           wil::com_ptr<IFileWriter>& writer,
                                           wil::com_ptr<IFileSystemBoundObject>& ownedStage) noexcept
    {
        writer.reset();
        ownedStage.reset();
        stagePath.clear();
        if (! destinationBinding)
        {
            return destinationBindingQueryHr == E_NOINTERFACE ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                                                              : (FAILED(destinationBindingQueryHr) ? destinationBindingQueryHr : E_UNEXPECTED);
        }

        constexpr unsigned int kMaximumStageCreateAttempts = 32u;
        for (unsigned int attempt = 0u; attempt < kMaximumStageCreateAttempts; ++attempt)
        {
            const HRESULT stageNameHr = MakeTempDestinationPath(destinationPath, progressStreamId, stagePath);
            if (FAILED(stageNameHr))
            {
                NoteFailure(QualifiedItemFailurePhase::StageCreate, stageNameHr);
                return stageNameHr;
            }
            const uint64_t createStartUs = PerfNowUs();
            const HRESULT createHr       = destinationBinding->CreateExclusiveWriter(stagePath.c_str(), &options, writer.put(), ownedStage.put());
            task._perf.bridgeStageCreateUs.fetch_add(PerfElapsedUs(createStartUs), std::memory_order_relaxed);
            task._perf.bridgeStageCreateCount.fetch_add(1u, std::memory_order_relaxed);
            if (createHr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) || createHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
            {
                if (writer || ownedStage)
                {
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                       E_UNEXPECTED,
                                       L"bridge.stage.collisionOwnedOutput",
                                       L"Exclusive stage collision returned an output that the host cannot safely adopt.",
                                       L"",
                                       destinationPath);
                    writer.reset();
                    ownedStage.reset();
                    return E_UNEXPECTED;
                }
                continue;
            }
            if (FAILED(createHr))
            {
                NoteFailure(QualifiedItemFailurePhase::StageCreate, createHr);
                writer.reset();
                ownedStage.reset();
                return createHr;
            }
            if (! writer || ! ownedStage)
            {
                NoteFailure(QualifiedItemFailurePhase::StageCreate, E_UNEXPECTED);
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   E_UNEXPECTED,
                                   L"bridge.stage.nullSuccess",
                                   L"Exclusive stage creation returned success without both writer and ownership authority.",
                                   L"",
                                   destinationPath);
                writer.reset();
                ownedStage.reset();
                return E_UNEXPECTED;
            }
#ifdef ENABLE_TESTS
            const HRESULT decorateHr = DecorateBridgeWriterForSelfTest(writer);
            if (FAILED(decorateHr))
            {
                NoteFailure(QualifiedItemFailurePhase::StageCreate, decorateHr);
                writer.reset();
                AbortOwnedStage(*ownedStage, stagePath, L"", destinationPath);
                ownedStage.reset();
                return decorateHr;
            }
#endif
            mutationAttemptFlags.fetch_or(kStageCreatedFlag, std::memory_order_acq_rel);
            ownedStageDisposition.store(OwnedStageDisposition::Owned, std::memory_order_release);
            return S_OK;
        }
        NoteFailure(QualifiedItemFailurePhase::StageCreate, HRESULT_FROM_WIN32(ERROR_FILE_EXISTS));
        return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }

    [[nodiscard]] HRESULT CreateOwnedLinkStage(const std::wstring& destinationPath,
                                               uint64_t progressStreamId,
                                               const FileSystemLinkInformation& information,
                                               std::wstring& stagePath,
                                               wil::com_ptr<IFileSystemBoundObject>& ownedStage) noexcept
    {
        ownedStage.reset();
        stagePath.clear();
        if (! destinationBinding)
        {
            return destinationBindingQueryHr == E_NOINTERFACE ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                                                              : (FAILED(destinationBindingQueryHr) ? destinationBindingQueryHr : E_UNEXPECTED);
        }

        constexpr unsigned int kMaximumStageCreateAttempts = 32u;
        for (unsigned int attempt = 0u; attempt < kMaximumStageCreateAttempts; ++attempt)
        {
            const HRESULT stageNameHr = MakeTempDestinationPath(destinationPath, progressStreamId, stagePath);
            if (FAILED(stageNameHr))
            {
                NoteFailure(QualifiedItemFailurePhase::StageCreate, stageNameHr);
                return stageNameHr;
            }

            const uint64_t createStartUs = PerfNowUs();
            const HRESULT createHr       = destinationBinding->CreateExclusiveLink(stagePath.c_str(), &information, &options, ownedStage.put());
            task._perf.bridgeStageCreateUs.fetch_add(PerfElapsedUs(createStartUs), std::memory_order_relaxed);
            task._perf.bridgeStageCreateCount.fetch_add(1u, std::memory_order_relaxed);
            if (createHr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) || createHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
            {
                if (ownedStage)
                {
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                       E_UNEXPECTED,
                                       L"bridge.linkStage.collisionOwnedOutput",
                                       L"Exclusive link-stage collision returned ownership authority that the host cannot safely adopt.",
                                       L"",
                                       destinationPath);
                    ownedStage.reset();
                    return E_UNEXPECTED;
                }
                continue;
            }
            if (FAILED(createHr))
            {
                NoteFailure(QualifiedItemFailurePhase::StageCreate, createHr);
                ownedStage.reset();
                return createHr;
            }
            if (! ownedStage)
            {
                NoteFailure(QualifiedItemFailurePhase::StageCreate, E_UNEXPECTED);
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   E_UNEXPECTED,
                                   L"bridge.linkStage.nullSuccess",
                                   L"Exclusive link-stage creation returned success without ownership authority.",
                                   L"",
                                   destinationPath);
                return E_UNEXPECTED;
            }
            mutationAttemptFlags.fetch_or(kStageCreatedFlag, std::memory_order_acq_rel);
            ownedStageDisposition.store(OwnedStageDisposition::Owned, std::memory_order_release);
            return S_OK;
        }
        NoteFailure(QualifiedItemFailurePhase::StageCreate, HRESULT_FROM_WIN32(ERROR_FILE_EXISTS));
        return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }

    void AbortOwnedStage(IFileSystemBoundObject& ownedStage,
                         const std::wstring& stagePath,
                         const std::wstring& sourcePath,
                         const std::wstring& destinationPath) noexcept
    {
        mutationAttemptFlags.fetch_or(kStageCleanupAttemptedFlag, std::memory_order_acq_rel);
        FileSystemConditionalMutationResult abortResult{};
        abortResult.sizeBytes                  = sizeof(abortResult);
        const FileSystemOptions cleanupOptions = MakeOwnedStageCleanupOptions(&options);
        const HRESULT abortHr                  = ownedStage.AbortOwnedObject(&cleanupOptions, &abortResult);
        if (abortResult.outcomeKnown == FALSE)
        {
            NoteFailure(QualifiedItemFailurePhase::StageAbort, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE));
            ownedStageDisposition.store(OwnedStageDisposition::Unknown, std::memory_order_release);
            ownedStageCleanupIndeterminate.store(true, std::memory_order_release);
            task._perf.bridgeStageRetainedCount.fetch_add(1u, std::memory_order_relaxed);
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                               HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE),
                               L"bridge.stage.abortUnknown",
                               L"Owned-stage cleanup outcome is unknown; the visible stage is retained as a Possible artifact.",
                               sourcePath,
                               destinationPath);
            return;
        }
        if (FAILED(abortHr) || abortResult.mutationCommitted == FALSE || abortResult.originalStillPresent != FALSE)
        {
            NoteFailure(QualifiedItemFailurePhase::StageAbort, FAILED(abortHr) ? abortHr : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY));
            ownedStageDisposition.store(OwnedStageDisposition::Retained, std::memory_order_release);
            ownedStageCleanupIndeterminate.store(true, std::memory_order_release);
            task._perf.bridgeStageRetainedCount.fetch_add(1u, std::memory_order_relaxed);
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                               FAILED(abortHr) ? abortHr : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                               L"bridge.stage.abortFailed",
                               std::format(L"Failed to remove the exact owned stage '{}'; it remains visible as a Possible artifact.", stagePath),
                               sourcePath,
                               destinationPath);
            return;
        }
        ownedStageDisposition.store(OwnedStageDisposition::Removed, std::memory_order_release);
    }

    void SleepResponsive(DWORD totalMs) noexcept
    {
        while (totalMs > 0)
        {
            if (CancelRequested())
            {
                return;
            }

            task.WaitWhilePaused();

            const DWORD slice = (std::min)(totalMs, SleepSliceMs());
            ::Sleep(slice);
            totalMs -= slice;
        }
    }

    [[nodiscard]] DWORD ComputeThrottleDelay(uint64_t bytesSoFar) noexcept
    {
        const uint64_t bandwidthLimit = bandwidthLimitBytesPerSecond.load(std::memory_order_acquire);
        if (bandwidthLimit == 0)
        {
            return 0u;
        }

        if (startTick == 0)
        {
            startTick = GetTickCount64();
        }

        const ULONGLONG now      = GetTickCount64();
        const uint64_t elapsedMs = static_cast<uint64_t>(now - startTick);

        constexpr uint64_t maxSafeBytes = std::numeric_limits<uint64_t>::max() / 1000u;

        uint64_t desiredMs = 0;
        if (bytesSoFar > 0 && bytesSoFar <= maxSafeBytes)
        {
            desiredMs = (bytesSoFar * 1000u) / bandwidthLimit;
        }
        else if (bytesSoFar > maxSafeBytes)
        {
            desiredMs = std::numeric_limits<uint64_t>::max();
        }

        if (desiredMs > elapsedMs)
        {
            const uint64_t remaining = desiredMs - elapsedMs;
            return remaining > std::numeric_limits<DWORD>::max() ? std::numeric_limits<DWORD>::max() : static_cast<DWORD>(remaining);
        }
        return 0u;
    }

    void ThrottleThreadSafe(uint64_t bytesSoFar) noexcept
    {
        if (bandwidthLimitBytesPerSecond.load(std::memory_order_acquire) == 0)
        {
            return;
        }

        DWORD sleepMs = 0u;
        {
            std::scoped_lock lock(throttleMutex);
            sleepMs = ComputeThrottleDelay(bytesSoFar);
        }
        if (sleepMs > 0u)
        {
            SleepResponsive(sleepMs);
        }
    }

    [[nodiscard]] bool ShouldUseBufferedPipeline(uint64_t fileTotalBytes, unsigned long bufferBytesIn) const noexcept
    {
#ifdef ENABLE_TESTS
        const FileOpsBridgePipelineMode mode = GetBridgePipelineModeOverride();
        if (mode == FileOpsBridgePipelineMode::Disabled)
        {
            return false;
        }
        if (mode == FileOpsBridgePipelineMode::Enabled)
        {
            return bufferBytesIn > 0;
        }
#endif
        return bufferBytesIn > 0 && fileTotalBytes > static_cast<uint64_t>(bufferBytesIn);
    }

    void AccumulateBridgeCopyPerf(
        const BridgeCopyPerf& perf, const std::wstring& sourcePath, const std::wstring& destinationPath, uint64_t transferredBytes, HRESULT hr) noexcept
    {
        task._perf.bridgeCopyUs.fetch_add(perf.copyUs, std::memory_order_relaxed);
        task._perf.bridgeReaderWaitUs.fetch_add(perf.readerWaitUs, std::memory_order_relaxed);
        task._perf.bridgeWriterWaitUs.fetch_add(perf.writerWaitUs, std::memory_order_relaxed);
        task._perf.bridgeReadUs.fetch_add(perf.readUs, std::memory_order_relaxed);
        task._perf.bridgeWriteUs.fetch_add(perf.writeUs, std::memory_order_relaxed);

        const uint64_t throughputBytesPerSecond =
            (perf.copyUs > 0 && transferredBytes > 0) ? static_cast<uint64_t>((transferredBytes * 1000000ull) / perf.copyUs) : 0ull;
        const std::wstring detail = std::format(L"source={} destination={} bytes={} bufferBytes={} progressPeriodMs={} latencyClass={} "
                                                L"hintFlags=0x{:X} progressCalls={} readerWaitUs={} writerWaitUs={} readUs={} writeUs={}",
                                                sourcePath,
                                                destinationPath,
                                                transferredBytes,
                                                bufferBytes,
                                                progressPeriodMs,
                                                transferLatencyClass,
                                                transferHintFlags,
                                                perf.progressCalls,
                                                perf.readerWaitUs,
                                                perf.writerWaitUs,
                                                perf.readUs,
                                                perf.writeUs);
        Debug::Perf::Emit(L"FileOps.Bridge.Copy", detail, perf.copyUs, transferredBytes, throughputBytesPerSecond, hr);
    }

    HRESULT ReportProgress(const std::wstring& currentSourcePath,
                           const std::wstring& currentDestinationPath,
                           uint64_t currentItemTotalBytes,
                           uint64_t currentItemCompletedBytes,
                           uint64_t callCompletedBytes,
                           uint64_t progressStreamId) noexcept
    {
        const uint64_t totalBytesSnapshot   = totalBytes;
        const uint64_t clampedCallCompleted = (totalBytesSnapshot > 0) ? (std::min)(totalBytesSnapshot, callCompletedBytes) : callCompletedBytes;

        std::scoped_lock lock(callbackMutex);
        options.bandwidthLimitBytesPerSecond = bandwidthLimitBytesPerSecond.load(std::memory_order_acquire);
        const HRESULT hr                     = task.FileSystemProgress(task._operation,
                                                                       1,
                                                                       0,
                                                                       totalBytesSnapshot,
                                                                       clampedCallCompleted,
                                                                       currentSourcePath.c_str(),
                                                                       currentDestinationPath.c_str(),
                                                                       currentItemTotalBytes,
                                                                       currentItemCompletedBytes,
                                                                       &options,
                                                                       progressStreamId,
                                                                       cookie);
        bandwidthLimitBytesPerSecond.store(options.bandwidthLimitBytesPerSecond, std::memory_order_release);
        return hr;
    }

    [[nodiscard]] HRESULT PromptDestinationCollision(const std::wstring& sourcePath,
                                                     const std::wstring& destinationPath,
                                                     HRESULT issueStatus,
                                                     bool& overwriteGranted,
                                                     bool& replaceReadOnlyGranted,
                                                     bool& replaceLinkGranted,
                                                     bool& keepBothRequested,
                                                     wil::com_ptr<IFileSystemBoundObject>& expectedDestination) noexcept
    {
        keepBothRequested = false;
        expectedDestination.reset();
        FileSystemIssueAction action = FileSystemIssueAction::Cancel;
        FileSystemOptions issueOptions{};
        HRESULT issueHr = E_UNEXPECTED;
        {
            std::scoped_lock lock(callbackMutex);
            issueOptions        = options;
            auto* perItemCookie = static_cast<PerItemCallbackCookie*>(cookie);
            if (perItemCookie != nullptr)
            {
                perItemCookie->keepBothRequested = false;
            }
            issueOptions.sizeBytes = sizeof(FileSystemOptions);
#ifdef ENABLE_TESTS
            task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
            const auto dbgCallbackScope = wil::scope_exit([&] noexcept { task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
#endif
            issueHr = task.FileSystemIssue(
                task.GetOperation(), sourcePath.c_str(), destinationPath.c_str(), issueStatus, &action, expectedDestination.put(), &issueOptions, cookie);
            if (perItemCookie != nullptr && perItemCookie->keepBothRequested)
            {
                keepBothRequested                = true;
                perItemCookie->keepBothRequested = false;
            }
        }
        if (FAILED(issueHr))
        {
            return issueHr;
        }

        if (keepBothRequested)
        {
            return S_OK;
        }

        switch (action)
        {
            case FileSystemIssueAction::Overwrite:
                if (! expectedDestination && destinationBinding)
                {
                    return E_UNEXPECTED; // a binding route must carry the exact authority
                }
                overwriteGranted = true;
                return S_OK;
            case FileSystemIssueAction::ReplaceLink:
                if (! expectedDestination)
                {
                    return E_UNEXPECTED;
                }
                overwriteGranted   = true;
                replaceLinkGranted = true;
                return S_OK;
            case FileSystemIssueAction::ReplaceReadOnly:
                if (! expectedDestination && destinationBinding)
                {
                    return E_UNEXPECTED;
                }
                overwriteGranted       = true;
                replaceReadOnlyGranted = true;
                return S_OK;
            case FileSystemIssueAction::KeepBoth: keepBothRequested = true; return S_OK;
            case FileSystemIssueAction::Skip: return S_FALSE;
            case FileSystemIssueAction::Retry:
            case FileSystemIssueAction::PermanentDelete:
            case FileSystemIssueAction::Cancel:
            case FileSystemIssueAction::None:
            default: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    void RecordSkippedDestinationCollision(const std::wstring& sourcePath,
                                           const std::wstring& destinationPath,
                                           HRESULT status,
                                           std::wstring_view diagnosticCode,
                                           std::wstring_view message) noexcept
    {
        skippedFileConflictCount.fetch_add(1, std::memory_order_acq_rel);
        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning, status, diagnosticCode, message, sourcePath, destinationPath);
    }

    [[nodiscard]] HRESULT SelectKeepBothDestination(const std::wstring& sourcePath, bool sourceIsDirectory, std::wstring& destinationPath) noexcept
    {
        size_t nextOrdinal = 2u;
        std::wstring keepBothDestination;
        const HRESULT keepBothHr = FindAvailableUniqueSiblingPath(&destinationIo, destinationPath, sourceIsDirectory, nextOrdinal, keepBothDestination);
        if (FAILED(keepBothHr))
        {
            return keepBothHr;
        }

        std::wstring sourceRelativePath;
        if (! TryGetSourceRelativePath(sourcePath, sourceRelativePath))
        {
            return E_UNEXPECTED;
        }
        if (sourceRelativePath.empty())
        {
            // Top-level Keep Both renames the destination root itself.
            rootDestinationPath = keepBothDestination;
        }
        destinationPath = std::move(keepBothDestination);
        return S_OK;
    }

    [[nodiscard]] HRESULT ClassifyExistingDestinationObject(const std::wstring& destinationPath, FileSystemBoundObjectKind& kind) noexcept
    {
        kind                                        = FILESYSTEM_BOUND_OTHER;
        constexpr FileSystemBindFlags bindFlags     = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
        FileOperations::ObjectBindingResult binding = FileOperations::BindObjectAuthority(&destinationFs, destinationPath, destinationPathProfileId, bindFlags);
        if (binding.state == FileOperations::ObjectBindingState::Missing)
        {
            return S_FALSE;
        }
        if (binding.state != FileOperations::ObjectBindingState::Bound)
        {
            return FAILED(binding.status) ? binding.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        kind = binding.authority.kind;
        if (kind != FILESYSTEM_BOUND_REGULAR_FILE && kind != FILESYSTEM_BOUND_DIRECTORY && kind != FILESYSTEM_BOUND_LINK)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT PublishOwnedDirectory(const std::wstring& sourcePath,
                                                const std::wstring& destinationPath,
                                                IFileSystemBoundObject* expectedDestination,
                                                bool replaceReadOnlyGranted,
                                                bool replaceLinkGranted,
                                                wil::com_ptr<IFileSystemBoundObject>& publishedAuthority) noexcept
    {
        publishedAuthority.reset();
        if (! destinationBinding)
        {
            return destinationBindingQueryHr == E_NOINTERFACE ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                                                              : (FAILED(destinationBindingQueryHr) ? destinationBindingQueryHr : E_UNEXPECTED);
        }

        std::wstring stagePath;
        wil::com_ptr<IFileSystemBoundObject> ownedStage;
        constexpr unsigned int kMaximumStageCreateAttempts = 32u;
        HRESULT hr                                         = E_UNEXPECTED;
        for (unsigned int attempt = 0u; attempt < kMaximumStageCreateAttempts; ++attempt)
        {
            hr = MakeTempDestinationPath(destinationPath, 0u, stagePath);
            if (FAILED(hr))
            {
                NoteFailure(QualifiedItemFailurePhase::StageCreate, hr);
                return hr;
            }
            hr = destinationBinding->CreateExclusiveDirectory(stagePath.c_str(), &options, ownedStage.put());
            if (SUCCEEDED(hr))
            {
                break;
            }
            ownedStage.reset();
            if (hr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && hr != HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
            {
                NoteFailure(QualifiedItemFailurePhase::StageCreate, hr);
                return hr;
            }
        }
        if (! ownedStage)
        {
            const HRESULT failure = FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
            NoteFailure(QualifiedItemFailurePhase::StageCreate, failure);
            return failure;
        }
        mutationAttemptFlags.fetch_or(kStageCreatedFlag, std::memory_order_acq_rel);
        ownedStageDisposition.store(OwnedStageDisposition::Owned, std::memory_order_release);

        bool promoted                 = false;
        bool ownedStageCleanupAllowed = true;
        const auto cleanupStage       = wil::scope_exit([&]() noexcept
        {
            if (! promoted && ownedStageCleanupAllowed)
            {
                AbortOwnedStage(*ownedStage, stagePath, sourcePath, destinationPath);
            }
        });

        // R3-3: the directory payload publishes through the shared owner (no payload diagnostics).
        return PublishOwnedStageAs(ownedStage.get(),
                                   {.sourcePath             = sourcePath,
                                    .destinationPath        = destinationPath,
                                    .expectedDestination    = expectedDestination,
                                    .overwriteGranted       = expectedDestination != nullptr,
                                    .replaceReadOnlyGranted = replaceReadOnlyGranted,
                                    .replaceLinkGranted     = replaceLinkGranted},
                                   publishedAuthority,
                                   promoted,
                                   ownedStageCleanupAllowed);
    }

    // Rename merge on a profile without object binding: the provider's own delete of the emptied
    // source directory, by name, which is the authority its Permanent Delete already has.
    [[nodiscard]] HRESULT RemoveEmptiedSourceDirectoryByName(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
    {
        if (CancelRequested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        mutationAttemptFlags.fetch_or(kSourceDeleteAttemptedFlag, std::memory_order_acq_rel);
        FileSystemOptions deleteOptions = options;
        const HRESULT deleteHr          = sourceFs.DeleteItem(sourcePath.c_str(), FILESYSTEM_FLAG_NONE, &deleteOptions, &task, cookie);
        if (SUCCEEDED(deleteHr))
        {
            return S_OK;
        }
        if (deleteHr == HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY))
        {
            MarkManagedSourceRetained(sourcePath, destinationPath, deleteHr, SourceKeptReason(L"the source directory gained or retained an unselected child."));
            return S_OK;
        }
        NoteFailure(QualifiedItemFailurePhase::SourceCleanup, deleteHr);
        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                           deleteHr,
                           L"bridge.renameMerge.cleanupFailed",
                           L"The emptied source directory could not be removed after its children were relocated.",
                           sourcePath,
                           destinationPath);
        return deleteHr;
    }

    HRESULT EnsureDestinationDirectory(const std::wstring& sourcePath,
                                       std::wstring& destinationPath,
                                       bool* createdDestination                               = nullptr,
                                       wil::com_ptr<IFileSystemBoundObject>* createdAuthority = nullptr) noexcept
    {
        if (createdDestination != nullptr)
        {
            *createdDestination = false;
        }
        if (createdAuthority != nullptr)
        {
            createdAuthority->reset();
        }
        if (! destinationDirOps && ! destinationBinding)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        for (unsigned int attempt = 0; attempt < 3u; ++attempt)
        {
            unsigned long attributes = 0;
            const HRESULT hrAttr     = destinationIo.GetAttributes(destinationPath.c_str(), &attributes);
            if (SUCCEEDED(hrAttr))
            {
                FileSystemBoundObjectKind destinationKind =
                    (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u ? FILESYSTEM_BOUND_DIRECTORY : FILESYSTEM_BOUND_REGULAR_FILE;
                if ((attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
                {
                    const HRESULT classificationHr = ClassifyExistingDestinationObject(destinationPath, destinationKind);
                    if (classificationHr == S_FALSE)
                    {
                        continue;
                    }
                    if (FAILED(classificationHr))
                    {
                        return classificationHr;
                    }
                }
                if (destinationKind == FILESYSTEM_BOUND_DIRECTORY)
                {
                    return S_OK;
                }

                const bool destinationIsLink = destinationKind == FILESYSTEM_BOUND_LINK;
                bool overwriteGranted        = false;
                bool replaceReadOnlyGranted  = false;
                bool replaceLinkGranted      = false;
                bool keepBothRequested       = false;
                wil::com_ptr<IFileSystemBoundObject> expectedDestination;
                const HRESULT promptHr =
                    PromptDestinationCollision(sourcePath,
                                               destinationPath,
                                               HRESULT_FROM_WIN32(destinationIsLink ? ERROR_REPARSE_POINT_ENCOUNTERED : ERROR_ALREADY_EXISTS),
                                               overwriteGranted,
                                               replaceReadOnlyGranted,
                                               replaceLinkGranted,
                                               keepBothRequested,
                                               expectedDestination);
                if (keepBothRequested)
                {
                    const HRESULT keepBothHr = SelectKeepBothDestination(sourcePath, true, destinationPath);
                    if (FAILED(keepBothHr))
                    {
                        return keepBothHr;
                    }
                    continue;
                }
                if (promptHr == S_FALSE || FAILED(promptHr))
                {
                    return promptHr;
                }
                if (! destinationIsLink || ! overwriteGranted || ! replaceLinkGranted || ! expectedDestination)
                {
                    // File-vs-directory is a type mismatch and never an overwrite. A
                    // destination link may be replaced only by its typed exact grant.
                    return HRESULT_FROM_WIN32(destinationIsLink ? ERROR_REPARSE_POINT_ENCOUNTERED : ERROR_DATATYPE_MISMATCH);
                }

                wil::com_ptr<IFileSystemBoundObject> publishedDirectory;
                const HRESULT replaceHr =
                    PublishOwnedDirectory(sourcePath, destinationPath, expectedDestination.get(), replaceReadOnlyGranted, true, publishedDirectory);
                if (FAILED(replaceHr))
                {
                    return replaceHr;
                }
                if (createdDestination != nullptr)
                {
                    *createdDestination = true;
                }
                if (createdAuthority != nullptr)
                {
                    *createdAuthority = std::move(publishedDirectory);
                }
                Debug::Perf::EmitCounter(L"FileOps.Link.ReplaceConditionalCount");
                return replaceHr;
            }

#ifdef ENABLE_TESTS
            // FIR-3 reviewed call-site exception: the injected race must occur after the absent-path
            // probe and before CreateDirectory; an interface decorator cannot identify that transition.
            MaybeInjectBridgeCreateDirectoryRaceForSelfTest(destinationIo, destinationPath);
#endif

            wil::com_ptr<IFileSystemBoundObject> publishedDirectory;
            const HRESULT hrCreate = destinationBinding ? PublishOwnedDirectory(sourcePath, destinationPath, nullptr, false, false, publishedDirectory)
                                                        : destinationDirOps->CreateDirectory(destinationPath.c_str());
            if (SUCCEEDED(hrCreate))
            {
                anyDestinationPublished.store(true, std::memory_order_release);
                if (hrCreate == S_OK)
                {
                    task.NoteLiveOutputPublished(destinationPath);
                }
                if (createdDestination != nullptr)
                {
                    *createdDestination = true;
                }
                if (createdAuthority != nullptr)
                {
                    *createdAuthority = std::move(publishedDirectory);
                }
                return S_OK;
            }
            if (hrCreate == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) || hrCreate == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
            {
                continue;
            }
            return hrCreate;
        }

        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    [[nodiscard]] bool CaptureCreatedDirectoryMetadata(const std::wstring& sourcePath,
                                                       const std::wstring& destinationPath,
                                                       bool createdDestination,
                                                       FileSystemBasicInformation& basicInfo) noexcept
    {
        basicInfo           = {};
        basicInfo.sizeBytes = sizeof(basicInfo);
        if (! createdDestination)
        {
            return false;
        }

        const HRESULT hr = sourceIo.GetFileBasicInformation(sourcePath.c_str(), &basicInfo);
        if (SUCCEEDED(hr))
        {
            return true;
        }
        if (hr != E_NOTIMPL && hr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        {
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                               hr,
                               L"bridge.directoryMetadata.read",
                               L"Source directory metadata could not be read; the created destination directory keeps provider defaults.",
                               sourcePath,
                               destinationPath);
        }
        return false;
    }

    void RestoreCreatedDirectoryMetadata(const std::wstring& sourcePath,
                                         const std::wstring& destinationPath,
                                         bool metadataCaptured,
                                         const FileSystemBasicInformation& basicInfo,
                                         IFileSystemBoundObject* createdAuthority) noexcept
    {
        if (! metadataCaptured)
        {
            return;
        }

        if (createdAuthority == nullptr)
        {
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                               HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                               L"bridge.directoryMetadata.authorityUnavailable",
                               L"Created directory metadata was not written because exact no-follow destination authority is unavailable.",
                               sourcePath,
                               destinationPath);
            return;
        }

        const HRESULT hr = createdAuthority->SetBasicInformation(&basicInfo);
        if (FAILED(hr) && hr != E_NOTIMPL && hr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        {
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                               hr,
                               L"bridge.directoryMetadata.write",
                               L"Created destination directory metadata could not be restored after its children completed.",
                               sourcePath,
                               destinationPath);
        }
    }

    [[nodiscard]] HRESULT MarkReparseSkipped(const std::wstring& sourcePath, const std::wstring& destinationPath, bool isDirectory, bool isRoot) noexcept
    {
        if (isDirectory)
        {
            ++skippedDirectoryReparseCount;
        }
        else
        {
            ++skippedFileReparseCount;
        }
        if (isDirectory && isRoot)
        {
            rootDirectoryReparseSkipped = true;
        }
        task._observedSkipAction.store(true, std::memory_order_release);
        if (cookie != nullptr)
        {
            static_cast<PerItemCallbackCookie*>(cookie)->explicitSkipObserved.store(true, std::memory_order_release);
        }

        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                           HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                           L"bridge.reparse.skip",
                           isDirectory ? (isRoot ? L"Skipped root directory reparse point by policy." : L"Skipped directory reparse point by policy.")
                                       : (isRoot ? L"Skipped root file reparse point by policy." : L"Skipped file reparse point by policy."),
                           sourcePath,
                           destinationPath);

        const uint64_t callCompleted = (completedBytesAtomic != nullptr) ? completedBytesAtomic->load(std::memory_order_acquire) : completedBytes;
        return ReportProgress(sourcePath, destinationPath, 0, 0, callCompleted, 0);
    }

    [[nodiscard]] static bool IsMissingPathHr(HRESULT hr) noexcept
    {
        return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    [[nodiscard]] HRESULT VerifyPublishedDestinationSize(const std::wstring& destinationPath,
                                                         uint64_t& destinationSizeBytes,
                                                         wil::com_ptr<IFileReader>& destinationReader) noexcept
    {
        constexpr unsigned int kMaxAttempts = 3u;
        destinationSizeBytes                = 0u;
        HRESULT hr                          = E_UNEXPECTED;
        for (unsigned int attempt = 0u; attempt < kMaxAttempts; ++attempt)
        {
            if (CancelRequested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            const uint64_t probeStartUs = PerfNowUs();
            destinationReader.reset();
            hr = destinationIo.CreateFileReader(destinationPath.c_str(), destinationReader.addressof());
            if (SUCCEEDED(hr) && destinationReader)
            {
                hr = destinationReader->GetSize(&destinationSizeBytes);
            }
            else if (SUCCEEDED(hr))
            {
                hr = E_POINTER;
            }
            task._perf.bridgeImmediateDestinationSizeProbeUs.fetch_add(PerfElapsedUs(probeStartUs), std::memory_order_relaxed);
            task._perf.bridgeImmediateDestinationSizeProbeCount.fetch_add(1u, std::memory_order_relaxed);

            if (SUCCEEDED(hr) || ! ShouldRetryPublishedDestinationVerification(hr) || attempt + 1u == kMaxAttempts)
            {
                return hr;
            }

#ifdef ENABLE_TESTS
            MaybePauseBeforePublishedDestinationRetryBackoffForSelfTest();
#endif
            SleepResponsive(25u * (attempt + 1u));
            if (CancelRequested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
        }

        return hr;
    }

    void StoreVerificationState(FileOperations::VerificationState state) noexcept
    {
        const auto priority = [](FileOperations::VerificationState value) noexcept -> unsigned int
        {
            switch (value)
            {
                case FileOperations::VerificationState::Failed: return 6u;
                case FileOperations::VerificationState::Canceled: return 5u;
                case FileOperations::VerificationState::Unavailable: return 4u;
                case FileOperations::VerificationState::Verified: return 3u;
                case FileOperations::VerificationState::NotApplicable: return 2u;
                case FileOperations::VerificationState::NotRequested: return 1u;
            }
            return 0u;
        };

        FileOperations::VerificationState current = verificationState.load(std::memory_order_acquire);
        while (priority(state) > priority(current) &&
               ! verificationState.compare_exchange_weak(current, state, std::memory_order_acq_rel, std::memory_order_acquire))
        {
        }
    }

    // R3-2: compare the provider's digest of the published object with the bytes streamed to it.
    // nullopt = no usable proof (verification stays Unavailable), false = a different object.
    [[nodiscard]] std::optional<bool> EvaluateWriterContentProof(IFileWriterContentProof& writerProof,
                                                                 std::vector<std::unique_ptr<Common::Crypto::ContentHasher>>& hashers,
                                                                 uint64_t streamedBytes,
                                                                 const std::wstring& sourcePath,
                                                                 const std::wstring& destinationPath) noexcept
    {
        FileSystemContentProof proof{};
        proof.sizeBytes       = sizeof(proof);
        const HRESULT proofHr = writerProof.GetCommittedContentProof(&proof);
        if (FAILED(proofHr) || proof.sizeBytes != sizeof(proof) || proof.contentSizeBytes != streamedBytes)
        {
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                               FAILED(proofHr) ? proofHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                               L"bridge.proof.writerDigestUnavailable",
                               L"The destination did not report a usable digest for the published object.",
                               sourcePath,
                               destinationPath);
            return std::nullopt;
        }
        for (const auto& hasher : hashers)
        {
            if (static_cast<uint32_t>(hasher->Algorithm()) != proof.algorithm)
            {
                continue;
            }
            std::vector<std::byte> digest;
            const size_t expectedBytes = Common::Crypto::ContentDigestBytes(hasher->Algorithm());
            if (! hasher->Finish(digest) || digest.size() != expectedBytes || expectedBytes > sizeof(proof.digest))
            {
                break;
            }
            if (! std::equal(digest.begin(), digest.end(), reinterpret_cast<const std::byte*>(proof.digest)))
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   HRESULT_FROM_WIN32(ERROR_CRC),
                                   L"bridge.proof.writerDigestMismatch",
                                   L"The destination's digest of the published object does not match the bytes sent to it.",
                                   sourcePath,
                                   destinationPath);
                return false;
            }
            task._verificationProviderProofCount.fetch_add(1u, std::memory_order_relaxed);
            task._perf.verificationProviderProofCount.fetch_add(1u, std::memory_order_relaxed);
            return true;
        }
        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                           HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                           L"bridge.proof.writerDigestUnavailable",
                           L"The destination reported a digest the host did not compute for this transfer.",
                           sourcePath,
                           destinationPath);
        return std::nullopt;
    }

    [[nodiscard]] HRESULT VerifyPublishedContent(const std::wstring& sourcePath,
                                                 const std::wstring& destinationPath,
                                                 uint64_t expectedSizeBytes,
                                                 const Common::Crypto::Blake3Digest& sourceDigest,
                                                 IFileSystemBoundObject* publishedAuthority,
                                                 std::byte* verificationBuffer,
                                                 unsigned long verificationBufferBytes,
                                                 ManagedSourceCleanupRecord& managedCleanupRecord,
                                                 std::optional<bool> writerProofMatched = std::nullopt) noexcept
    {
        if (! verificationRequested)
        {
            return S_OK;
        }

        verificationFileCount.fetch_add(1u, std::memory_order_relaxed);
        const uint64_t completedBeforeItem = verificationCompletedBytes.load(std::memory_order_acquire);
        if (completedBeforeItem > (std::numeric_limits<uint64_t>::max)() - expectedSizeBytes)
        {
            StoreVerificationState(FileOperations::VerificationState::Failed);
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        // Discovery owns the frozen denominator. Before it closes, expose at least the
        // currently verifying file without adding it twice when discovery-ahead has
        // already reported the same bytes.
        AtomicMax(task._verificationTotalBytes, completedBeforeItem + expectedSizeBytes);

        const uint64_t verificationStartUs = PerfNowUs();
        const auto finishTiming            = wil::scope_exit([&] noexcept
        {
            task._perf.verificationUs.fetch_add(PerfElapsedUs(verificationStartUs), std::memory_order_relaxed);
            task._verificationActive.store(false, std::memory_order_release);
        });
        if (writerProofMatched.has_value())
        {
            // R3-2: the provider's own digest already answered for the published object.
            if (! writerProofMatched.value())
            {
                StoreVerificationState(FileOperations::VerificationState::Failed);
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   HRESULT_FROM_WIN32(ERROR_CRC),
                                   L"bridge.verification.mismatch",
                                   L"The destination's digest of the published object does not match the bytes read from the source.",
                                   sourcePath,
                                   destinationPath);
                if (managedCleanupRecord.armed)
                {
                    MarkManagedSourceRetained(
                        sourcePath, destinationPath, HRESULT_FROM_WIN32(ERROR_CRC), LoadStringResource(nullptr, IDS_FILEOPS_RETAINED_VERIFICATION_FAILED));
                    managedCleanupRecord = {};
                }
                return HRESULT_FROM_WIN32(ERROR_CRC);
            }
            const uint64_t completedBefore = verificationCompletedBytes.fetch_add(expectedSizeBytes, std::memory_order_acq_rel);
            const uint64_t completedAfter  = completedBefore > (std::numeric_limits<uint64_t>::max)() - expectedSizeBytes
                                                 ? (std::numeric_limits<uint64_t>::max)()
                                                 : completedBefore + expectedSizeBytes;
            StoreVerificationState(FileOperations::VerificationState::Verified);
            static_cast<void>(
                task.ReportVerificationProgress(sourcePath.c_str(), destinationPath.c_str(), expectedSizeBytes, expectedSizeBytes, completedAfter, false));
            Debug::Perf::Emit(L"FileOps.Verification", L"proof=writer-digest", PerfElapsedUs(verificationStartUs), expectedSizeBytes, 0u, S_OK);
            return S_OK;
        }
        HRESULT hr = task.ReportVerificationProgress(
            sourcePath.c_str(), destinationPath.c_str(), expectedSizeBytes, 0u, verificationCompletedBytes.load(std::memory_order_acquire), true);
        if (FAILED(hr))
        {
            StoreVerificationState(FileOperations::VerificationState::Canceled);
            if (managedCleanupRecord.armed)
            {
                MarkManagedSourceRetained(
                    sourcePath, destinationPath, HRESULT_FROM_WIN32(ERROR_CANCELLED), LoadStringResource(nullptr, IDS_FILEOPS_RETAINED_VERIFICATION_CANCELED));
            }
            managedCleanupRecord = {};
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        Common::Crypto::Blake3Digest destinationDigest{};
        bool proofAvailable    = false;
        bool usedProviderProof = false;
#ifdef ENABLE_TESTS
        const bool forceUnavailable =
            ConsumeBridgeCounterForSelfTest(g_fileOpsVerificationForceUnavailableCount, g_fileOpsVerificationForceUnavailableAttempts);
        const bool forceHostReadback =
            ! forceUnavailable && ConsumeBridgeCounterForSelfTest(g_fileOpsVerificationForceHostReadbackCount, g_fileOpsVerificationForceHostReadbackAttempts);
#else
        constexpr bool forceUnavailable  = false;
        constexpr bool forceHostReadback = false;
#endif
        const bool providerProofAllowed = verificationProviderBlake3Proof && ! forceUnavailable && ! forceHostReadback;
        const bool hostReadbackAllowed  = verificationHostReadback && ! forceUnavailable;
        if (providerProofAllowed && publishedAuthority != nullptr)
        {
            wil::com_ptr<IFileSystemBoundContentProof> contentProof;
            const HRESULT queryHr = publishedAuthority->QueryInterface(IID_PPV_ARGS(contentProof.addressof()));
            if (SUCCEEDED(queryHr) && contentProof)
            {
                FileSystemContentProof proof{};
                proof.sizeBytes       = sizeof(proof);
                const HRESULT proofHr = contentProof->GetContentProof(&options, &proof);
                if (SUCCEEDED(proofHr) && proof.sizeBytes == sizeof(proof) && proof.algorithm == FILESYSTEM_CONTENT_PROOF_BLAKE3_256 &&
                    proof.contentSizeBytes == expectedSizeBytes)
                {
                    static_assert(sizeof(proof.digest) == destinationDigest.size());
                    std::memcpy(destinationDigest.data(), proof.digest, destinationDigest.size());
                    proofAvailable    = true;
                    usedProviderProof = true;
                    task._verificationProviderProofCount.fetch_add(1u, std::memory_order_relaxed);
                    task._perf.verificationProviderProofCount.fetch_add(1u, std::memory_order_relaxed);
                }
            }
        }

        if (! proofAvailable && hostReadbackAllowed && publishedAuthority != nullptr && verificationBuffer != nullptr && verificationBufferBytes > 0u)
        {
            wil::com_ptr<IFileReader> destinationReader;
            hr = publishedAuthority->OpenReader(&options, destinationReader.put());
            if (SUCCEEDED(hr) && destinationReader)
            {
                uint64_t destinationSizeBytes = 0u;
                hr                            = destinationReader->GetSize(&destinationSizeBytes);
                if (SUCCEEDED(hr) && destinationSizeBytes != expectedSizeBytes)
                {
                    hr = HRESULT_FROM_WIN32(ERROR_CRC);
                }
                if (SUCCEEDED(hr))
                {
                    Common::Crypto::Blake3Hasher destinationHasher;
                    uint64_t fileVerifiedBytes = 0u;
                    for (;;)
                    {
                        task.WaitWhilePaused();
                        if (CancelRequested())
                        {
                            hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                            break;
                        }
                        unsigned long bytesRead = 0u;
                        const HRESULT readHr    = destinationReader->Read(verificationBuffer, verificationBufferBytes, &bytesRead);
                        task._verificationReadCount.fetch_add(1u, std::memory_order_relaxed);
                        task._perf.verificationReadCalls.fetch_add(1u, std::memory_order_relaxed);
                        if (FAILED(readHr))
                        {
                            hr = readHr;
                            break;
                        }
                        if (bytesRead > verificationBufferBytes || fileVerifiedBytes > (std::numeric_limits<uint64_t>::max)() - bytesRead)
                        {
                            hr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                            break;
                        }
                        if (bytesRead == 0u)
                        {
                            break;
                        }
                        destinationHasher.Update(std::span<const std::byte>(verificationBuffer, bytesRead));
                        fileVerifiedBytes += bytesRead;
                        task._perf.verificationReadBytes.fetch_add(bytesRead, std::memory_order_relaxed);
                        const uint64_t completedBefore = verificationCompletedBytes.fetch_add(bytesRead, std::memory_order_acq_rel);
                        if (completedBefore > (std::numeric_limits<uint64_t>::max)() - bytesRead)
                        {
                            hr = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                            break;
                        }
                        const uint64_t completedAfter = completedBefore + bytesRead;
                        hr                            = task.ReportVerificationProgress(
                            sourcePath.c_str(), destinationPath.c_str(), expectedSizeBytes, fileVerifiedBytes, completedAfter, true);
                        if (FAILED(hr))
                        {
                            break;
                        }
#ifdef ENABLE_TESTS
                        // Pause after publishing real verification progress so UI tests
                        // can observe the Verifying state and cancellation can be injected
                        // at a deterministic readback checkpoint.
                        g_fileOpsVerificationReadbackPausePoint.Pause(5'000ull);
#endif
                        const uint64_t transferredBytes =
                            completedBytesAtomic != nullptr ? completedBytesAtomic->load(std::memory_order_acquire) : completedBytes;
                        ThrottleThreadSafe(transferredBytes + completedAfter);
                    }
                    if (SUCCEEDED(hr) && fileVerifiedBytes != expectedSizeBytes)
                    {
                        hr = HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
                    }
                    if (SUCCEEDED(hr))
                    {
                        destinationDigest = destinationHasher.Finalize();
                        proofAvailable    = true;
                        task._verificationHostReadbackCount.fetch_add(1u, std::memory_order_relaxed);
                        task._perf.verificationHostReadbackCount.fetch_add(1u, std::memory_order_relaxed);
                    }
                }
            }
        }

        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT || CancelRequested())
        {
            StoreVerificationState(FileOperations::VerificationState::Canceled);
            if (managedCleanupRecord.armed)
            {
                MarkManagedSourceRetained(
                    sourcePath, destinationPath, HRESULT_FROM_WIN32(ERROR_CANCELLED), LoadStringResource(nullptr, IDS_FILEOPS_RETAINED_VERIFICATION_CANCELED));
            }
            managedCleanupRecord = {};
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (! proofAvailable)
        {
            StoreVerificationState(FileOperations::VerificationState::Unavailable);
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                               FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                               L"bridge.verification.unavailable",
                               L"Requested content verification could not obtain an exact provider proof or bound-object readback.",
                               sourcePath,
                               destinationPath);
            if (managedCleanupRecord.armed)
            {
                MarkManagedSourceRetained(sourcePath,
                                          destinationPath,
                                          FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                          LoadStringResource(nullptr, IDS_FILEOPS_RETAINED_VERIFICATION_UNAVAILABLE));
                managedCleanupRecord = {};
            }
            return S_FALSE;
        }

        const uint64_t completedNow = verificationCompletedBytes.load(std::memory_order_acquire);
        if (usedProviderProof)
        {
            const uint64_t completedBefore = verificationCompletedBytes.fetch_add(expectedSizeBytes, std::memory_order_acq_rel);
            if (completedBefore > (std::numeric_limits<uint64_t>::max)() - expectedSizeBytes)
            {
                StoreVerificationState(FileOperations::VerificationState::Failed);
                managedCleanupRecord = {};
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            const uint64_t completedAfter   = completedBefore + expectedSizeBytes;
            const uint64_t transferredBytes = completedBytesAtomic != nullptr ? completedBytesAtomic->load(std::memory_order_acquire) : completedBytes;
            // An opaque provider proof may complete in one callback, but its verified byte
            // count still shares the configured data budget with transfer work.
            ThrottleThreadSafe(transferredBytes + completedAfter);
            hr = task.ReportVerificationProgress(sourcePath.c_str(), destinationPath.c_str(), expectedSizeBytes, expectedSizeBytes, completedAfter, true);
            if (FAILED(hr))
            {
                StoreVerificationState(FileOperations::VerificationState::Canceled);
                if (managedCleanupRecord.armed)
                {
                    MarkManagedSourceRetained(sourcePath,
                                              destinationPath,
                                              HRESULT_FROM_WIN32(ERROR_CANCELLED),
                                              LoadStringResource(nullptr, IDS_FILEOPS_RETAINED_VERIFICATION_CANCELED));
                }
                managedCleanupRecord = {};
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
        }

#ifdef ENABLE_TESTS
        if (ConsumeBridgeCounterForSelfTest(g_fileOpsVerificationForceMismatchCount, g_fileOpsVerificationForceMismatchAttempts))
        {
            destinationDigest[0] ^= std::byte{0x01};
        }
#endif
        if (destinationDigest != sourceDigest)
        {
            StoreVerificationState(FileOperations::VerificationState::Failed);
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                               HRESULT_FROM_WIN32(ERROR_CRC),
                               L"bridge.verification.mismatch",
                               L"Published destination BLAKE3 content does not match the bytes read from the source.",
                               sourcePath,
                               destinationPath);
            if (managedCleanupRecord.armed)
            {
                MarkManagedSourceRetained(
                    sourcePath, destinationPath, HRESULT_FROM_WIN32(ERROR_CRC), LoadStringResource(nullptr, IDS_FILEOPS_RETAINED_VERIFICATION_FAILED));
                managedCleanupRecord = {};
            }
            return HRESULT_FROM_WIN32(ERROR_CRC);
        }

        StoreVerificationState(FileOperations::VerificationState::Verified);
        const uint64_t finalCompleted = usedProviderProof ? verificationCompletedBytes.load(std::memory_order_acquire) : completedNow;
        static_cast<void>(
            task.ReportVerificationProgress(sourcePath.c_str(), destinationPath.c_str(), expectedSizeBytes, expectedSizeBytes, finalCompleted, false));
        Debug::Perf::Emit(L"FileOps.Verification",
                          usedProviderProof ? L"proof=provider-blake3" : L"proof=host-readback-blake3",
                          PerfElapsedUs(verificationStartUs),
                          expectedSizeBytes,
                          task._perf.verificationReadCalls.load(std::memory_order_relaxed),
                          S_OK);
        return S_OK;
    }

    void InitializeConnectionLimits(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
    {
        if (connectionLimitsInitialized)
        {
            return;
        }
        connectionLimitsInitialized = true;

        const Common::Settings::Settings* settingsSnapshot = (task._folderWindow != nullptr) ? task._folderWindow->_settings : nullptr;
        if (! settingsSnapshot)
        {
            return;
        }

        const auto initSide = [&](std::wstring_view pluginPath, bool isSource) noexcept
        {
            const auto connNameOpt = ConnectionProfileUtils::TryParseConnNameFromPluginPath(pluginPath);
            if (! connNameOpt.has_value())
            {
                return;
            }

            const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(settingsSnapshot, *connNameOpt);
            if (! profile || profile->id.empty())
            {
                return;
            }

            const uint32_t pluginCap = static_cast<uint32_t>(isSource ? sourcePluginMaxConcurrencyBudget : destinationPluginMaxConcurrencyBudget);

            uint32_t maxEffective      = (std::max)(1u, pluginCap);
            const uint32_t overrideRaw = ConnectionProfileUtils::ExtraGetUInt32(profile->extra, "copyMoveMaxConcurrency").value_or(0);
            if (overrideRaw != 0)
            {
                const uint32_t clamped = std::clamp<uint32_t>(overrideRaw, 1u, 16u);
                maxEffective           = (std::max)(1u, std::min<uint32_t>(maxEffective, clamped));
            }

            ConnectionLimit limit{};
            limit.id          = profile->id;
            limit.maxCopyMove = maxEffective;

            if (isSource)
            {
                sourceConnectionLimit = std::move(limit);
            }
            else
            {
                destinationConnectionLimit = std::move(limit);
            }
        };

        initSide(sourcePath, true);
        initSide(destinationPath, false);
    }

    [[nodiscard]] HRESULT AcquireCopyMovePermits(ConnectionConcurrencyLimiter::Permit& outFirst, ConnectionConcurrencyLimiter::Permit& outSecond) noexcept
    {
        outFirst  = {};
        outSecond = {};

        if (! sourceConnectionLimit.has_value() && ! destinationConnectionLimit.has_value())
        {
            return S_OK;
        }

        if (CancelRequested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        ConnectionConcurrencyLimiter& limiter = GetConnectionConcurrencyLimiter();
        const auto shouldCancel               = [&]() noexcept { return CancelRequested(); };

        if (sourceConnectionLimit.has_value() && destinationConnectionLimit.has_value() && sourceConnectionLimit->id == destinationConnectionLimit->id)
        {
            const uint32_t mergedMax                    = std::min(sourceConnectionLimit->maxCopyMove, destinationConnectionLimit->maxCopyMove);
            ConnectionConcurrencyLimiter::Permit permit = limiter.AcquireCopyMove(sourceConnectionLimit->id, mergedMax, shouldCancel);
            if (! permit)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            outFirst = std::move(permit);
            return S_OK;
        }

        const ConnectionLimit* firstLimit  = sourceConnectionLimit.has_value() ? &*sourceConnectionLimit : nullptr;
        const ConnectionLimit* secondLimit = destinationConnectionLimit.has_value() ? &*destinationConnectionLimit : nullptr;

        if (! firstLimit || ! secondLimit)
        {
            const ConnectionLimit* only                 = firstLimit ? firstLimit : secondLimit;
            ConnectionConcurrencyLimiter::Permit permit = limiter.AcquireCopyMove(only->id, only->maxCopyMove, shouldCancel);
            if (! permit)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            outFirst = std::move(permit);
            return S_OK;
        }

        const bool sourceFirst          = firstLimit->id <= secondLimit->id;
        const ConnectionLimit* acquireA = sourceFirst ? firstLimit : secondLimit;
        const ConnectionLimit* acquireB = sourceFirst ? secondLimit : firstLimit;

        ConnectionConcurrencyLimiter::Permit permitA = limiter.AcquireCopyMove(acquireA->id, acquireA->maxCopyMove, shouldCancel);
        if (! permitA)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        ConnectionConcurrencyLimiter::Permit permitB = limiter.AcquireCopyMove(acquireB->id, acquireB->maxCopyMove, shouldCancel);
        if (! permitB)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        outFirst  = std::move(permitA);
        outSecond = std::move(permitB);
        return S_OK;
    }

    [[nodiscard]] unsigned int ComputeWithinFolderBudget() const noexcept
    {
        if (verificationRequested)
        {
            return 1u;
        }
        const unsigned int taskBudget = std::max(1u, task._perItemMaxConcurrencyBudget);

        size_t activeTopLevelCalls = std::max<size_t>(1u, GetPerItemInFlightCallCountSnapshot(task));

        if (activeTopLevelCalls > static_cast<size_t>((std::numeric_limits<unsigned int>::max)()))
        {
            activeTopLevelCalls = static_cast<size_t>((std::numeric_limits<unsigned int>::max)());
        }

        const unsigned int divisor = static_cast<unsigned int>(activeTopLevelCalls);
        const unsigned int perCall = divisor == 0 ? taskBudget : (taskBudget / divisor);
        return std::max(1u, perCall);
    }

    // Raised before any bytes are read, so answering a conflict never re-transfers an
    // already-copied file. Outputs the per-file overwrite grants the copy should use.
    // S_FALSE = user chose Skip (caller records it and moves on).
    [[nodiscard]] HRESULT ValidateDestinationOverwritePolicy(const std::wstring& sourcePath,
                                                             std::wstring& destinationPath,
                                                             bool& overwriteGranted,
                                                             bool& replaceReadOnlyGranted,
                                                             bool& replaceLinkGranted,
                                                             wil::com_ptr<IFileSystemBoundObject>& expectedDestination,
                                                             std::optional<FileSystemBasicInformation>& replaceExpectation) noexcept
    {
        replaceExpectation.reset();
        const bool requestedOverwrite       = (static_cast<uint32_t>(flags) & static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE)) != 0u;
        const bool requestedReplaceReadOnly = (static_cast<uint32_t>(flags) & static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY)) != 0u;
        overwriteGranted                    = false;
        replaceReadOnlyGranted              = false;
        replaceLinkGranted                  = false;
        expectedDestination.reset();

        unsigned long destinationAttributes = 0;
        const HRESULT hrDestAttr            = destinationIo.GetAttributes(destinationPath.c_str(), &destinationAttributes);
        if (FAILED(hrDestAttr))
        {
            // A compatibility flag is not a replacement receipt for an absent name.
            // Keep exclusive/atomic-final publication enabled so a raced destination is
            // surfaced rather than treated as a pre-authorized overwrite.
            return S_OK;
        }

        overwriteGranted       = requestedOverwrite;
        replaceReadOnlyGranted = requestedReplaceReadOnly;

        FileSystemBoundObjectKind destinationKind =
            (destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u ? FILESYSTEM_BOUND_DIRECTORY : FILESYSTEM_BOUND_REGULAR_FILE;
        if ((destinationAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
        {
            const HRESULT classificationHr = ClassifyExistingDestinationObject(destinationPath, destinationKind);
            if (classificationHr == S_FALSE)
            {
                return S_OK;
            }
            if (FAILED(classificationHr))
            {
                return classificationHr;
            }
        }

        if (destinationKind == FILESYSTEM_BOUND_DIRECTORY)
        {
            bool ignoredOverwriteGrant   = false;
            bool ignoredReadOnlyGrant    = false;
            bool ignoredReplaceLinkGrant = false;
            bool keepBothRequested       = false;
            wil::com_ptr<IFileSystemBoundObject> ignoredExpectedDestination;
            const HRESULT directoryPrompt = PromptDestinationCollision(sourcePath,
                                                                       destinationPath,
                                                                       HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                                                                       ignoredOverwriteGrant,
                                                                       ignoredReadOnlyGrant,
                                                                       ignoredReplaceLinkGrant,
                                                                       keepBothRequested,
                                                                       ignoredExpectedDestination);
            if (keepBothRequested)
            {
                return SelectKeepBothDestination(sourcePath, false, destinationPath);
            }
            if (directoryPrompt == S_FALSE || FAILED(directoryPrompt))
            {
                return directoryPrompt;
            }

            // A file cannot safely replace a directory through the bridge. The local
            // prompt layout suppresses Overwrite for a probeable file-vs-directory
            // collision, but keep this guard for non-local providers where the prompt
            // cannot prove the dead end.
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        const bool readonlyCollision = (destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0;
        const bool linkCollision     = destinationKind == FILESYSTEM_BOUND_LINK;
        const bool alreadyAuthorized = overwriteGranted && (! readonlyCollision || replaceReadOnlyGranted) && (! linkCollision || replaceLinkGranted);
        if (alreadyAuthorized && destinationBinding)
        {
            return BindExpectedPublicationDestination(destinationPath, true, expectedDestination);
        }
        // R3-1: on a route without object binding a compatibility flag is not a receipt for the
        // occupant; the user decides per object and the writer carries what they saw. The occupant
        // is read before the prompt so the decision applies to what the user was shown.
        overwriteGranted       = false;
        replaceReadOnlyGranted = false;
        std::optional<FileSystemBasicInformation> occupantBeforePrompt;
        if (! destinationBinding)
        {
            CaptureReplaceExpectation(destinationPath, occupantBeforePrompt);
        }

        // Raise a PER-FILE conflict instead of failing the whole bridge transfer. The
        // host serializes prompts and caches apply-to-all answers (Fairstream 1C/1D).
        bool keepBothRequested        = false;
        const HRESULT collisionStatus = linkCollision ? HRESULT_FROM_WIN32(ERROR_REPARSE_POINT_ENCOUNTERED)
                                                      : HRESULT_FROM_WIN32(readonlyCollision ? ERROR_ACCESS_DENIED : ERROR_ALREADY_EXISTS);
        const HRESULT promptHr        = PromptDestinationCollision(
            sourcePath, destinationPath, collisionStatus, overwriteGranted, replaceReadOnlyGranted, replaceLinkGranted, keepBothRequested, expectedDestination);
        if (! keepBothRequested)
        {
            if (SUCCEEDED(promptHr) && promptHr != S_FALSE && overwriteGranted && ! expectedDestination)
            {
                replaceExpectation = occupantBeforePrompt;
            }
            return promptHr;
        }
        const HRESULT keepBothHr = SelectKeepBothDestination(sourcePath, false, destinationPath);
        expectedDestination.reset();
        return keepBothHr;
    }

    // R3-1: the occupant the user is about to decide on, read no-follow before the prompt; the
    // atomic-final writer refuses to publish over anything else. Unknown fields stay zero.
    // R0-RC3: a replacement is conditional only on an occupant the host could read. No token means
    // no expectation, and PrepareStage then refuses the granted replacement instead of replacing
    // whatever is at the name at commit time.
    void CaptureReplaceExpectation(const std::wstring& destinationPath, std::optional<FileSystemBasicInformation>& expectation) noexcept
    {
        expectation.reset();
        FileSystemBasicInformation occupant{};
        occupant.sizeBytes = sizeof(occupant);
        HRESULT hr         = destinationIo.GetFileBasicInformation(destinationPath.c_str(), &occupant);
#ifdef ENABLE_TESTS
        if (ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeFailNextDestinationBasicInfoCount, g_fileOpsBridgeFailNextDestinationBasicInfoAttempts))
        {
            hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
#endif
        if (SUCCEEDED(hr))
        {
            expectation = occupant;
        }
    }

    // Publication seam: consumes exact owned-stage authority and preserves the provider's
    // committed/unknown/non-commit result without granting cleanup authority from a path.
    // The file payload's owned-stage publication: the shared owner below with the file diagnostics.
    [[nodiscard]] HRESULT PublishOwnedStage(const std::wstring& sourcePath,
                                            const std::wstring& destinationPath,
                                            IFileSystemBoundObject& ownedStage,
                                            IFileSystemBoundObject* expectedDestination,
                                            bool overwriteGranted,
                                            bool replaceReadOnlyGranted,
                                            bool replaceLinkGranted,
                                            wil::com_ptr<IFileSystemBoundObject>& publishedAuthority,
                                            bool& promoted,
                                            bool& ownedStageCleanupAllowed) noexcept
    {
        return PublishOwnedStageAs(&ownedStage,
                                   {.sourcePath             = sourcePath,
                                    .destinationPath        = destinationPath,
                                    .expectedDestination    = expectedDestination,
                                    .overwriteGranted       = overwriteGranted,
                                    .replaceReadOnlyGranted = replaceReadOnlyGranted,
                                    .replaceLinkGranted     = replaceLinkGranted,
                                    .unknownCategory        = L"bridge.publication.unknown",
                                    .unknownMessage         = L"Publication outcome is unknown; destination and visible stage are retained for reconciliation.",
                                    .partialCategory        = L"bridge.publication.cleanupPartial",
                                    .partialMessage         = L"Destination was published, but exact post-publication cleanup was incomplete."},
                                   publishedAuthority,
                                   promoted,
                                   ownedStageCleanupAllowed);
    }

    // R3-3: the one owner of an owned stage's exact PublishAs outcome, shared by the file, link,
    // and directory payloads. Publication truth (unknown / committed / not committed) sets the
    // bridge flags in one place; the payload only names its diagnostics and failure phase.
    struct OwnedStagePublication final
    {
        std::wstring_view sourcePath;
        std::wstring_view destinationPath;
        IFileSystemBoundObject* expectedDestination = nullptr;
        bool overwriteGranted                       = false;
        bool replaceReadOnlyGranted                 = false;
        bool replaceLinkGranted                     = false;
        std::wstring_view unknownCategory; // empty: no diagnostic for an unknown outcome
        std::wstring_view unknownMessage;
        std::wstring_view partialCategory; // empty: no diagnostic for a partial post-publication cleanup
        std::wstring_view partialMessage;
        QualifiedItemFailurePhase partialPhase = QualifiedItemFailurePhase::FinalPublish;
    };

    [[nodiscard]] HRESULT PublishOwnedStageAs(IFileSystemBoundObject* ownedStage,
                                              const OwnedStagePublication& publication,
                                              wil::com_ptr<IFileSystemBoundObject>& publishedAuthority,
                                              bool& promoted,
                                              bool& ownedStageCleanupAllowed) noexcept
    {
        if (ownedStage == nullptr)
        {
            return E_POINTER;
        }
        const std::wstring destinationPath(publication.destinationPath);
        FileSystemConditionalMutationResult publicationResult{};
        publicationResult.sizeBytes       = sizeof(publicationResult);
        const uint64_t publicationStartUs = PerfNowUs();
        mutationAttemptFlags.fetch_or(kFinalPublishAttemptedFlag, std::memory_order_acq_rel);
        const HRESULT publishHr =
            ownedStage->PublishAs(destinationPath.c_str(),
                                  publication.expectedDestination,
                                  BuildPublicationFlags(publication.overwriteGranted, publication.replaceReadOnlyGranted, publication.replaceLinkGranted),
                                  &options,
                                  &publicationResult,
                                  publishedAuthority.put());
        task._perf.bridgePublicationUs.fetch_add(PerfElapsedUs(publicationStartUs), std::memory_order_relaxed);
        task._perf.bridgePublicationCount.fetch_add(1u, std::memory_order_relaxed);

        if (publicationResult.outcomeKnown == FALSE)
        {
            // Unknown: the destination and the visible stage are retained for reconciliation.
            NoteFailure(QualifiedItemFailurePhase::FinalPublishReconcile, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE));
            ownedStageDisposition.store(OwnedStageDisposition::Unknown, std::memory_order_release);
            destinationPublicationUnknown.store(true, std::memory_order_release);
            ownedStageCleanupAllowed = false;
            task._perf.bridgeStageRetainedCount.fetch_add(1u, std::memory_order_relaxed);
            if (! publication.unknownCategory.empty())
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE),
                                   publication.unknownCategory,
                                   publication.unknownMessage,
                                   publication.sourcePath,
                                   publication.destinationPath);
            }
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (publicationResult.mutationCommitted != FALSE)
        {
            // Committed: the final name is published; the stage is no longer the caller's to abort.
            promoted                 = true;
            ownedStageCleanupAllowed = false;
            anyDestinationPublished.store(true, std::memory_order_release);
            task.NoteLiveOutputPublished(destinationPath);
            ownedStageDisposition.store(OwnedStageDisposition::Published, std::memory_order_release);
            if (! publishedAuthority)
            {
                NoteFailure(QualifiedItemFailurePhase::FinalPublish, E_UNEXPECTED);
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   E_UNEXPECTED,
                                   L"bridge.publication.nullSuccess",
                                   L"Publication committed without returning exact published-object authority.",
                                   publication.sourcePath,
                                   publication.destinationPath);
                return E_UNEXPECTED;
            }
            if (publicationResult.originalStillPresent != FALSE)
            {
                NoteFailure(QualifiedItemFailurePhase::FinalPublish, E_UNEXPECTED);
                return E_UNEXPECTED;
            }
            if (FAILED(publishHr))
            {
                NoteFailure(publication.partialPhase, publishHr);
                if (! publication.partialCategory.empty())
                {
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                       publishHr,
                                       publication.partialCategory,
                                       publication.partialMessage,
                                       publication.sourcePath,
                                       publication.destinationPath);
                }
                return publishHr;
            }
            return S_OK;
        }

        // Known non-commit: nothing is visible under the final name; the stage is aborted by the caller.
        if (SUCCEEDED(publishHr) || publishedAuthority)
        {
            NoteFailure(QualifiedItemFailurePhase::FinalPublish, E_UNEXPECTED);
            return E_UNEXPECTED;
        }
        NoteFailure(QualifiedItemFailurePhase::FinalPublish, publishHr);
        return publishHr;
    }

    // R3-3: one way to open an item's source reader: the exact bound object when the item has one
    // (the cleanup authority for a Managed Move, otherwise the metadata binding), else the legacy
    // pathname reader on routes that cannot bind.
    [[nodiscard]] HRESULT OpenSourceReader(const std::wstring& sourcePath,
                                           const std::wstring& destinationPath,
                                           IFileSystemBoundObject* boundSource,
                                           bool legacyReaderAllowed,
                                           wil::com_ptr<IFileReader>& reader) noexcept
    {
        reader.reset();
        if (boundSource != nullptr)
        {
            const HRESULT hr = boundSource->OpenReader(&options, reader.put());
            if (SUCCEEDED(hr) && ! reader)
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   E_UNEXPECTED,
                                   L"bridge.reader.nullBoundSuccess",
                                   L"Bound source OpenReader returned success without a reader object.",
                                   sourcePath,
                                   destinationPath);
                return E_UNEXPECTED;
            }
            if (FAILED(hr))
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   hr,
                                   L"bridge.reader.openBound",
                                   L"The exact bound source reader could not be opened; the source remains untouched.",
                                   sourcePath,
                                   destinationPath);
                return hr;
            }
            return S_OK;
        }
        if (! legacyReaderAllowed)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        const HRESULT hr = sourceIo.CreateFileReader(sourcePath.c_str(), reader.addressof());
        if (FAILED(hr))
        {
            return hr;
        }
        if (! reader)
        {
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                               E_UNEXPECTED,
                               L"bridge.reader.nullSuccess",
                               L"Source reader returned success without a reader object.",
                               sourcePath,
                               destinationPath);
            return E_UNEXPECTED;
        }
        // R0f-Curl-OR1: a pathname reader carries no options of its own; hand it the task's operation
        // control so a cancel ends an in-flight Read instead of waiting for the next one.
        wil::com_ptr<IFileReaderOperationControl> readerControl;
        if (SUCCEEDED(reader->QueryInterface(IID_PPV_ARGS(readerControl.addressof()))) && readerControl)
        {
            static_cast<void>(readerControl->SetOperationControl(&options));
        }
        return S_OK;
    }

    // R3-3: one write loop for both pump shapes. Writes a read chunk completely, counts the bytes
    // into the item and the task, reports progress, and throttles; a short or oversized write and
    // a failed write are StageWrite failures.
    template <typename ReportProgressFn>
    [[nodiscard]] HRESULT WriteChunk(IFileWriter& writer,
                                     const std::byte* chunk,
                                     unsigned long chunkBytes,
                                     uint64_t fileTotalBytes,
                                     uint64_t& fileCompletedBytes,
                                     std::atomic<uint64_t>& overallCompletedBytes,
                                     uint64_t& writeUs,
                                     ReportProgressFn&& reportProgress) noexcept
    {
        size_t offset = 0;
        while (offset < chunkBytes)
        {
            if (CancelRequested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            unsigned long bytesWritten = 0;
            const unsigned long toWrite =
                static_cast<unsigned long>(std::min(static_cast<size_t>(chunkBytes - offset), static_cast<size_t>(std::numeric_limits<unsigned long>::max())));
            const uint64_t writeStartUs = PerfNowUs();
            const HRESULT hrWrite       = writer.Write(chunk + offset, toWrite, &bytesWritten);
            writeUs += PerfElapsedUs(writeStartUs);
            if (FAILED(hrWrite))
            {
                NoteFailure(QualifiedItemFailurePhase::StageWrite, hrWrite);
                return hrWrite;
            }
            if (bytesWritten == 0)
            {
                const HRESULT writeHr = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
                NoteFailure(QualifiedItemFailurePhase::StageWrite, writeHr);
                return writeHr;
            }
            if (bytesWritten > toWrite)
            {
                const HRESULT writeHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                NoteFailure(QualifiedItemFailurePhase::StageWrite, writeHr);
                return writeHr;
            }

            offset += bytesWritten;

            if (fileCompletedBytes > std::numeric_limits<uint64_t>::max() - static_cast<uint64_t>(bytesWritten))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            fileCompletedBytes += bytesWritten;

            const uint64_t previousOverall = overallCompletedBytes.fetch_add(bytesWritten, std::memory_order_acq_rel);
            if (previousOverall > (std::numeric_limits<uint64_t>::max)() - static_cast<uint64_t>(bytesWritten))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            const uint64_t overallAfter = previousOverall + static_cast<uint64_t>(bytesWritten);
            const bool forceProgress    = fileTotalBytes > 0 && fileCompletedBytes >= fileTotalBytes;

            const HRESULT hrProgress = reportProgress(overallAfter, forceProgress);
            if (FAILED(hrProgress))
            {
                return hrProgress;
            }

            ThrottleThreadSafe(overallAfter);
        }
        return S_OK;
    }

    // R3-3: the one owner of a regular file's publication truth from admission to cleanup. Every
    // field was a local of the former single pump function; the phases below read and write the
    // record in order, and the item's receipt (publication, verification, source disposition,
    // failure phase) follows from the state it reaches.
    struct PublicationTransaction final
    {
        enum class State : uint8_t
        {
            Admitted,    // destination policy answered (grants captured), nothing touched
            SourceBound, // source authority, metadata, reader, and sizes ready; cleanup record armed for Move
            Staged,      // atomic-final writer or owned stage created (expectation and proof handed over)
            Written,     // every source byte accepted by the writer (size and position agree)
            Committed,   // writer Commit succeeded (atomic-final: visible; owned stage: staged only)
            Published,   // final name published with exact authority (or the atomic-final writer promoted)
            Verified,    // requested verification or writer proof answered for the published object
        };

        const std::wstring& sourcePath;
        std::wstring destinationPath; // owned: the destination policy may rename it (Keep Both)
        State state = State::Admitted;

        // Admission grants and expectations.
        bool overwriteGranted       = false;
        bool replaceReadOnlyGranted = false;
        bool replaceLinkGranted     = false;
        wil::com_ptr<IFileSystemBoundObject> expectedDestination;
        std::optional<FileSystemBasicInformation> replaceExpectation;

        // Source authority, metadata, cleanup record, reader, and sizes.
        FileOperations::BoundObjectAuthority sourceMetadataAuthority{};
        bool legacySourceReaderAllowed = false;
        wil::com_ptr<IFileSystemBoundMetadata> sourceMetadata;
        FileSystemMetadataSnapshot sourceMetadataSnapshot{};
        bool hasSourceMetadataSnapshot = false;
        bool retainSourceByConsent     = false;
        ManagedSourceCleanupRecord managedCleanupRecord{};
        wil::com_ptr<IFileReader> reader;
        FileSystemBasicInformation sourceBasicInfo{};
        bool hasSourceBasicInfo     = false;
        uint64_t fileTotalBytes     = 0;
        bool hasKnownFileTotalBytes = false;

        // Writer route.
        AtomicWriterRoute writerRoute{};
        bool identityLess           = false;
        bool writerProofRoute       = false;
        bool identityLessReplace    = false;
        bool useAtomicFinalWriter   = false;
        bool useIdentityOwnedStage  = false;
        FileSystemFlags writerFlags = FILESYSTEM_FLAG_NONE;

        // Stage and writer handles.
        std::wstring stagePath;
        bool promoted                 = false;
        bool ownedStageCleanupAllowed = false;
        wil::com_ptr<IFileSystemBoundObject> ownedStage;
        wil::com_ptr<IFileSystemBoundObject> publishedAuthority;
        wil::com_ptr<IFileWriter> writer;
        wil::com_ptr<IFileWriterContentProof> writerProof;
        std::vector<std::unique_ptr<Common::Crypto::ContentHasher>> writerProofHashers;
        wil::com_ptr<IFileWriterCommitSizeProof> commitSizeProof;
        wil::com_ptr<IFileSystemBoundMetadata> destinationStageMetadata;
        FileSystemMetadataTransferResult metadataResult{};

        // Pump and proofs. `isMove` is captured when the pump starts: a later consent may release
        // the cleanup record, but the post-commit size checks keep the Move shape they were given.
        bool isMove                 = false;
        uint64_t fileCompletedBytes = 0;
        std::optional<Common::Crypto::Blake3Hasher> sourceHasher;
        Common::Crypto::Blake3Digest sourceDigest{};
        BridgeCopyPerf copyPerf{};
        bool committedSizeProven{};
        std::optional<bool> writerProofMatched;

        PublicationTransaction(const std::wstring& sourcePathIn, std::wstring destinationPathIn) noexcept
            : sourcePath(sourcePathIn),
              destinationPath(std::move(destinationPathIn))
        {
            sourceMetadataSnapshot.sizeBytes = sizeof(sourceMetadataSnapshot);
            sourceBasicInfo.sizeBytes        = sizeof(FileSystemBasicInformation);
            metadataResult.sizeBytes         = sizeof(metadataResult);
            metadataResult.firstFailure      = S_OK;
        }
        PublicationTransaction(const PublicationTransaction&)            = delete;
        PublicationTransaction& operator=(const PublicationTransaction&) = delete;
        PublicationTransaction(PublicationTransaction&&)                 = delete;
        PublicationTransaction& operator=(PublicationTransaction&&)      = delete;
    };

    // Pump seam: owns one reader-to-writer transaction, including serialized verification,
    // then delegates exact final-name publication to PublishOwnedStage.
    // R3-1: a replacement refused because the occupant changed after the decision re-raises the
    // collision on the current occupant (bounded), so the user decides on what is really there.
    HRESULT PumpAndPublishFile(const std::wstring& sourcePath,
                               std::wstring destinationPath,
                               std::byte* bufferIn,
                               unsigned long bufferBytesIn,
                               uint64_t progressStreamId,
                               std::atomic<uint64_t>& overallCompletedBytes,
                               bool adoptFileSizeAsTotalWhenUnknown = false) noexcept
    {
        constexpr unsigned int kMaxReplaceRePrompts = 2u;
        for (unsigned int attempt = 0u;; ++attempt)
        {
            bool replaceRefused = false;
            const HRESULT hr    = PumpAndPublishFileOnce(
                sourcePath, destinationPath, bufferIn, bufferBytesIn, progressStreamId, overallCompletedBytes, adoptFileSizeAsTotalWhenUnknown, replaceRefused);
            if (! replaceRefused || attempt >= kMaxReplaceRePrompts || CancelRequested())
            {
                return hr;
            }
        }
    }

    // Phase 1: the destination policy decision (grants and expectations). S_FALSE = user skip.
    [[nodiscard]] HRESULT AdmitDestination(PublicationTransaction& txn) noexcept
    {
        const HRESULT hrDestPolicy = ValidateDestinationOverwritePolicy(txn.sourcePath,
                                                                        txn.destinationPath,
                                                                        txn.overwriteGranted,
                                                                        txn.replaceReadOnlyGranted,
                                                                        txn.replaceLinkGranted,
                                                                        txn.expectedDestination,
                                                                        txn.replaceExpectation);
        if (FAILED(hrDestPolicy))
        {
            return hrDestPolicy;
        }
        if (hrDestPolicy == S_FALSE)
        {
            // User chose Skip for this file: leave the destination untouched, count it so
            // the move keeps the source and the task ends PARTIAL.
            RecordSkippedDestinationCollision(txn.sourcePath,
                                              txn.destinationPath,
                                              HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                                              L"bridge.conflict.skip",
                                              L"Destination already exists; skipped on user request.");
            return S_FALSE;
        }
        txn.state = PublicationTransaction::State::Admitted;
        return S_OK;
    }

    // Phase 2: exact source authority, metadata snapshot and its consents, the Move cleanup
    // record, the source reader, and the sizes the transfer will be held to.
    [[nodiscard]] HRESULT BindSource(PublicationTransaction& txn, bool adoptFileSizeAsTotalWhenUnknown) noexcept
    {
        const std::wstring& sourcePath                                = txn.sourcePath;
        const std::wstring& destinationPath                           = txn.destinationPath;
        FileOperations::BoundObjectAuthority& sourceMetadataAuthority = txn.sourceMetadataAuthority;
        bool& legacySourceReaderAllowed                               = txn.legacySourceReaderAllowed;
        wil::com_ptr<IFileSystemBoundMetadata>& sourceMetadata        = txn.sourceMetadata;
        FileSystemMetadataSnapshot& sourceMetadataSnapshot            = txn.sourceMetadataSnapshot;
        bool& hasSourceMetadataSnapshot                               = txn.hasSourceMetadataSnapshot;
        bool& retainSourceByConsent                                   = txn.retainSourceByConsent;
        ManagedSourceCleanupRecord& managedCleanupRecord              = txn.managedCleanupRecord;

        {
            constexpr FileSystemBindFlags metadataFlags =
                static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_CONTENT | FILESYSTEM_BIND_READ_METADATA);
            FileOperations::ObjectBindingResult metadataBinding =
                FileOperations::BindObjectAuthority(&sourceFs, sourcePath, sourcePathProfileId, metadataFlags);
            if (metadataBinding.state == FileOperations::ObjectBindingState::Bound)
            {
                if (metadataBinding.authority.kind != FILESYSTEM_BOUND_REGULAR_FILE)
                {
                    return HRESULT_FROM_WIN32(ERROR_FILE_INVALID);
                }
                sourceMetadataAuthority = std::move(metadataBinding.authority);
            }
            else if (metadataBinding.state == FileOperations::ObjectBindingState::Unsupported)
            {
                legacySourceReaderAllowed = true;
            }
            else
            {
                return FAILED(metadataBinding.status) ? metadataBinding.status : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }

        if (sourceMetadataAuthority.boundObject && SUCCEEDED(sourceMetadataAuthority.boundObject->QueryInterface(IID_PPV_ARGS(sourceMetadata.addressof()))) &&
            sourceMetadata)
        {
            const uint64_t inspectStartUs = PerfNowUs();
            const HRESULT snapshotHr      = sourceMetadata->GetMetadataSnapshot(&options, &sourceMetadataSnapshot);
            const std::wstring inspectDetail =
                std::format(L"operation={} supported=0x{:X}", task._operation == FILESYSTEM_MOVE ? L"move" : L"copy", sourceMetadataSnapshot.supportedFeatures);
            Debug::Perf::Emit(L"FileOps.Metadata.InspectUs",
                              inspectDetail,
                              PerfElapsedUs(inspectStartUs),
                              sourceMetadataSnapshot.presentFeatures,
                              sourceMetadataSnapshot.logicalSizeBytes,
                              snapshotHr);
            if (SUCCEEDED(snapshotHr))
            {
                hasSourceMetadataSnapshot = true;
            }
            else if (snapshotHr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && snapshotHr != E_NOINTERFACE)
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   snapshotHr,
                                   L"bridge.metadata.inspect",
                                   L"Exact source metadata inspection failed; unsupported metadata is reported rather than assumed preserved.",
                                   sourcePath,
                                   destinationPath);
            }
        }
        if (hasSourceMetadataSnapshot && (sourceMetadataSnapshot.presentFeatures & FILESYSTEM_METADATA_PLACEHOLDER) != 0u)
        {
            const DeferredConsentResult consent = RequestBridgeConsent(FileOperations::DeferredConsentRisk::PlaceholderHydration,
                                                                       HRESULT_FROM_WIN32(ERROR_FILE_OFFLINE),
                                                                       sourcePath,
                                                                       destinationPath,
                                                                       &sourceMetadataAuthority,
                                                                       &sourceMetadataSnapshot);
            if (FAILED(consent.status))
            {
                return consent.status;
            }
            if (consent.action == ConflictAction::RetainSource)
            {
                retainSourceByConsent = true;
            }
            else if (consent.action != ConflictAction::Proceed)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
        }

        HRESULT hr = PrepareManagedSourceAuthority(sourcePath, destinationPath, FILESYSTEM_BOUND_REGULAR_FILE, managedCleanupRecord);
        if (FAILED(hr))
        {
            return hr;
        }
        if (managedCleanupRecord.armed && sourceMetadataAuthority.boundObject)
        {
            bool sameObject   = false;
            bool sameRevision = false;
            hr = FileOperations::CrossCheckBoundObjectIdentity(sourceMetadataAuthority, managedCleanupRecord.authority, sameObject, sameRevision);
            if (FAILED(hr) || ! sameObject || ! sameRevision)
            {
                return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
            }
        }
        if (managedCleanupRecord.armed && ! sourceMetadataAuthority.boundObject)
        {
            sourceMetadataAuthority = managedCleanupRecord.authority;
            static_cast<void>(sourceMetadataAuthority.boundObject->QueryInterface(IID_PPV_ARGS(sourceMetadata.addressof())));
            if (sourceMetadata && ! hasSourceMetadataSnapshot)
            {
                hasSourceMetadataSnapshot = SUCCEEDED(sourceMetadata->GetMetadataSnapshot(&options, &sourceMetadataSnapshot));
            }
        }
        if (retainSourceByConsent && managedCleanupRecord.armed)
        {
            MarkManagedSourceRetained(sourcePath, destinationPath, S_FALSE, L"Copied; source kept by the placeholder-hydration consent decision.");
            managedCleanupRecord = {};
        }

        hr = OpenSourceReader(sourcePath,
                              destinationPath,
                              managedCleanupRecord.armed ? managedCleanupRecord.authority.boundObject.get() : sourceMetadataAuthority.boundObject.get(),
                              legacySourceReaderAllowed,
                              txn.reader);
        if (FAILED(hr))
        {
            return hr;
        }

        const HRESULT hrGetBasic = sourceMetadataAuthority.boundObject ? sourceMetadataAuthority.boundObject->GetBasicInformation(&txn.sourceBasicInfo)
                                                                       : sourceIo.GetFileBasicInformation(sourcePath.c_str(), &txn.sourceBasicInfo);
        if (SUCCEEDED(hrGetBasic))
        {
            txn.hasSourceBasicInfo = true;
        }
        else if (hrGetBasic != E_NOTIMPL && hrGetBasic != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        {
            Debug::Warning(L"CrossFileSystemBridge: GetFileBasicInformation failed for '{}' (hr={:#x})", sourcePath, static_cast<unsigned long>(hrGetBasic));
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                               hrGetBasic,
                               L"bridge.metadata.read",
                               L"GetFileBasicInformation failed for source file.",
                               sourcePath,
                               destinationPath);
        }

        wil::com_ptr<IFileReader>& reader = txn.reader;
        uint64_t& fileTotalBytes          = txn.fileTotalBytes;
        bool& hasKnownFileTotalBytes      = txn.hasKnownFileTotalBytes;
        const HRESULT hrReaderSize        = reader->GetSize(&fileTotalBytes);
        if (SUCCEEDED(hrReaderSize))
        {
            hasKnownFileTotalBytes = true;
        }
        else
        {
            const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            const bool isMove       = task._operation == FILESYSTEM_MOVE;
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                               partialHr,
                               L"bridge.integrity.sourceSizeUnknown",
                               isMove ? L"Cross-filesystem MOVE cannot verify source size; preserving source."
                                      : L"Cross-filesystem COPY cannot verify source size; not committing destination.",
                               sourcePath,
                               destinationPath);
            return partialHr;
        }

        if (adoptFileSizeAsTotalWhenUnknown && totalBytes == 0 && txn.hasKnownFileTotalBytes)
        {
            totalBytes = txn.fileTotalBytes;
        }
        txn.state = PublicationTransaction::State::SourceBound;
        return S_OK;
    }

    // Phase 3: the writer route. Atomic-final publication through the provider writer, an
    // identity-owned stage, or neither (authority unavailable).
    [[nodiscard]] HRESULT RouteWriter(PublicationTransaction& txn) noexcept
    {
        const std::wstring& sourcePath                   = txn.sourcePath;
        const std::wstring& destinationPath              = txn.destinationPath;
        ManagedSourceCleanupRecord& managedCleanupRecord = txn.managedCleanupRecord;
        const bool overwriteGranted                      = txn.overwriteGranted;

        wil::com_ptr<IFileSystemAtomicWriter> atomicWriterCapability;
        static_cast<void>(destinationFs.QueryInterface(IID_PPV_ARGS(atomicWriterCapability.addressof())));
        txn.writerRoute = ResolveAtomicWriterRoute(atomicWriterCapability.get(), destinationPath.c_str(), flags, overwriteGranted, txn.replaceReadOnlyGranted);
        const AtomicWriterRoute& writerRoute = txn.writerRoute;

        txn.identityLess     = ! destinationBinding;
        txn.writerProofRoute = txn.identityLess && verificationWriterDigestProof && writerRoute.useAtomicFinalWriter;
        if (managedCleanupRecord.armed && txn.identityLess && ! txn.writerProofRoute)
        {
            MarkManagedSourceRetained(sourcePath,
                                      destinationPath,
                                      HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                      L"Copied; source kept because the destination cannot retain exact publication authority.");
            managedCleanupRecord = {};
        }
        // R3-1: a granted replacement on a route without object binding publishes through the
        // provider's atomic-final writer, which carries the occupant the user saw.
        // R3-2: on such a route requested verification and Managed Move also use that writer,
        // proved by the provider's own digest of the published object.
        txn.identityLessReplace   = overwriteGranted && ! txn.expectedDestination && txn.identityLess;
        txn.useAtomicFinalWriter  = writerRoute.useAtomicFinalWriter && (! managedCleanupRecord.armed || txn.writerProofRoute) &&
                                    (! overwriteGranted || txn.identityLessReplace) && (! verificationRequested || txn.writerProofRoute);
        txn.useIdentityOwnedStage = ! txn.identityLess && (managedCleanupRecord.armed || ! txn.useAtomicFinalWriter);
        if (! txn.useAtomicFinalWriter && ! txn.useIdentityOwnedStage)
        {
            const HRESULT unsupportedHr = destinationBindingQueryHr == E_NOINTERFACE
                                              ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                                              : (FAILED(destinationBindingQueryHr) ? destinationBindingQueryHr : E_UNEXPECTED);
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                               unsupportedHr,
                               L"bridge.stage.authorityUnavailable",
                               L"Destination cannot provide atomic-final or identity-owned staged publication.",
                               sourcePath,
                               destinationPath);
            return unsupportedHr;
        }

        const bool useAtomicFinalWriter   = txn.useAtomicFinalWriter;
        const FileSystemFlags writerFlags = useAtomicFinalWriter ? writerRoute.effectiveFinalFlags : writerRoute.fallbackSiblingFlags;
        txn.writerFlags                   = writerFlags;
        return S_OK;
    }

    // Phase 4: the stage. An atomic-final writer (with the replace expectation and the proof
    // hashers it can answer) or an identity-owned stage, plus the stage's metadata interface.
    [[nodiscard]] HRESULT PrepareStage(PublicationTransaction& txn, uint64_t progressStreamId, bool& replaceRefused) noexcept
    {
        const std::wstring& sourcePath                                                  = txn.sourcePath;
        const std::wstring& destinationPath                                             = txn.destinationPath;
        ManagedSourceCleanupRecord& managedCleanupRecord                                = txn.managedCleanupRecord;
        wil::com_ptr<IFileWriter>& writer                                               = txn.writer;
        wil::com_ptr<IFileWriterContentProof>& writerProof                              = txn.writerProof;
        std::vector<std::unique_ptr<Common::Crypto::ContentHasher>>& writerProofHashers = txn.writerProofHashers;
        HRESULT hr                                                                      = S_OK;

        if (txn.useAtomicFinalWriter)
        {
            hr = destinationIo.CreateFileWriter(destinationPath.c_str(), txn.writerFlags, writer.put());
            if (FAILED(hr))
            {
                return hr;
            }
            if (! writer)
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   E_UNEXPECTED,
                                   L"bridge.writer.nullSuccess",
                                   L"Destination writer returned success without a writer object.",
                                   sourcePath,
                                   destinationPath);
                return E_UNEXPECTED;
            }
            if (txn.identityLessReplace)
            {
                wil::com_ptr<IFileWriterExpectedReplacement> replacement;
                if (FAILED(writer->QueryInterface(IID_PPV_ARGS(replacement.addressof()))) || ! replacement)
                {
                    const HRESULT unsupportedHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                       unsupportedHr,
                                       L"bridge.replace.expectationUnavailable",
                                       L"The destination writer cannot carry the occupant the user agreed to replace; nothing was written.",
                                       sourcePath,
                                       destinationPath);
                    return unsupportedHr;
                }
                if (! txn.replaceExpectation.has_value())
                {
                    // R0-RC3: the occupant the user agreed to replace could not be read; a replacement
                    // that cannot be conditioned on it is refused before the first byte.
                    const HRESULT unsupportedHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                       unsupportedHr,
                                       L"bridge.replace.expectationUnavailable",
                                       L"The occupant the user agreed to replace could not be read; nothing was written.",
                                       sourcePath,
                                       destinationPath);
                    return unsupportedHr;
                }
                FileSystemBasicInformation expected = txn.replaceExpectation.value();
                expected.sizeBytes                  = sizeof(expected);
                hr                                  = replacement->SetExpectedReplacement(&expected);
                if (FAILED(hr))
                {
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                       hr,
                                       L"bridge.replace.occupantChanged",
                                       L"The destination changed after the replace decision; nothing was replaced.",
                                       sourcePath,
                                       destinationPath);
                    replaceRefused = hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
                    return hr;
                }
            }
            if (txn.writerProofRoute && (verificationRequested || managedCleanupRecord.armed))
            {
                // R3-2: hash the streamed bytes with every digest the writer can prove afterwards.
                static_cast<void>(writer->QueryInterface(IID_PPV_ARGS(writerProof.addressof())));
                uint32_t algorithmMask = 0u;
                if (writerProof && SUCCEEDED(writerProof->GetContentProofAlgorithms(&algorithmMask)))
                {
                    for (const Common::Crypto::ContentDigestAlgorithm algorithm : kWriterProofAlgorithmPreference)
                    {
                        if ((algorithmMask & (1u << static_cast<uint32_t>(algorithm))) == 0u)
                        {
                            continue;
                        }
                        auto hasher = std::make_unique<Common::Crypto::ContentHasher>(algorithm);
                        if (hasher && hasher->Valid())
                        {
                            writerProofHashers.push_back(std::move(hasher));
                        }
                    }
                }
                if (writerProofHashers.empty())
                {
                    writerProof.reset();
                }
            }
        }
        else
        {
            hr = CreateOwnedStage(destinationPath, progressStreamId, txn.stagePath, writer, txn.ownedStage);
            if (FAILED(hr))
            {
                return hr;
            }
            txn.ownedStageCleanupAllowed = true;
        }

        if (txn.ownedStage)
        {
            static_cast<void>(txn.ownedStage->QueryInterface(IID_PPV_ARGS(txn.destinationStageMetadata.addressof())));
        }
        txn.state = PublicationTransaction::State::Staged;
        return S_OK;
    }

    // Phase 5: pre-content metadata (sparse, EFS) on the stage, with the deferred consents that
    // may keep a Move source instead of proceeding.
    [[nodiscard]] HRESULT TransferPreContentMetadata(PublicationTransaction& txn) noexcept
    {
        const std::wstring& sourcePath                                = txn.sourcePath;
        const std::wstring& destinationPath                           = txn.destinationPath;
        ManagedSourceCleanupRecord& managedCleanupRecord              = txn.managedCleanupRecord;
        FileOperations::BoundObjectAuthority& sourceMetadataAuthority = txn.sourceMetadataAuthority;
        FileSystemMetadataSnapshot& sourceMetadataSnapshot            = txn.sourceMetadataSnapshot;
        FileSystemMetadataTransferResult& metadataResult              = txn.metadataResult;
        HRESULT hr                                                    = S_OK;

        if (txn.hasSourceMetadataSnapshot && txn.sourceMetadata)
        {
            // Compression is never transferred (the destination folder's state is inherited,
            // as in File Explorer), so it is never a pre-content loss.
            const uint32_t preContentFeatures = sourceMetadataSnapshot.presentFeatures & (FILESYSTEM_METADATA_SPARSE | FILESYSTEM_METADATA_EFS);
            if (txn.destinationStageMetadata)
            {
                const uint64_t metadataStartUs = PerfNowUs();
                const HRESULT metadataHr =
                    txn.sourceMetadata->TransferMetadataTo(txn.ownedStage.get(), FILESYSTEM_METADATA_TRANSFER_PREPARE_CONTENT, &options, &metadataResult);
                Debug::Perf::Emit(L"FileOps.Metadata.ApplyUs",
                                  L"phase=prepare",
                                  PerfElapsedUs(metadataStartUs),
                                  metadataResult.attemptedFeatures,
                                  metadataResult.lostFeatures,
                                  metadataHr);
                if (FAILED(metadataHr))
                {
                    metadataResult.attemptedFeatures |= preContentFeatures;
                    metadataResult.lostFeatures |= preContentFeatures;
                    metadataResult.firstFailure = metadataHr;
                }
            }
            else
            {
                metadataResult.attemptedFeatures |= preContentFeatures;
                metadataResult.lostFeatures |= preContentFeatures;
                if (preContentFeatures != 0u)
                {
                    metadataResult.firstFailure = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
            }

            const auto applyPreContentConsent =
                [&](uint32_t feature, FileOperations::DeferredConsentRisk risk, std::wstring_view retainedReason) noexcept -> HRESULT
            {
                if ((metadataResult.lostFeatures & feature) == 0u)
                {
                    return S_OK;
                }
                const DeferredConsentResult consent =
                    RequestBridgeConsent(risk,
                                         FAILED(metadataResult.firstFailure) ? metadataResult.firstFailure : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                         sourcePath,
                                         destinationPath,
                                         &sourceMetadataAuthority,
                                         &sourceMetadataSnapshot);
                if (FAILED(consent.status))
                {
                    return consent.status;
                }
                if (consent.action == ConflictAction::RetainSource)
                {
                    if (managedCleanupRecord.armed)
                    {
                        MarkManagedSourceRetained(sourcePath, destinationPath, S_FALSE, retainedReason);
                        managedCleanupRecord = {};
                    }
                    return S_OK;
                }
                return consent.action == ConflictAction::Proceed ? S_OK : HRESULT_FROM_WIN32(ERROR_CANCELLED);
            };
            hr = applyPreContentConsent(FILESYSTEM_METADATA_EFS,
                                        FileOperations::DeferredConsentRisk::EfsPlaintext,
                                        L"Copied; source kept because EFS encryption could not be preserved.");
            if (SUCCEEDED(hr))
            {
                hr = applyPreContentConsent(FILESYSTEM_METADATA_SPARSE,
                                            FileOperations::DeferredConsentRisk::SparseInflation,
                                            L"Copied; source kept because sparse allocation could not be preserved.");
            }
            if (FAILED(hr))
            {
                return hr;
            }
            ReportMetadataOutcome(sourcePath, destinationPath, metadataResult);
        }
        return S_OK;
    }

    // Phase 6: the pump. Every source byte through the bounded buffer(s) into the writer, then the
    // size, digest, and writer-position facts the commit is held to.
    [[nodiscard]] HRESULT Pump(PublicationTransaction& txn,
                               std::byte* bufferIn,
                               unsigned long bufferBytesIn,
                               uint64_t progressStreamId,
                               std::atomic<uint64_t>& overallCompletedBytes) noexcept
    {
        const std::wstring& sourcePath                                                  = txn.sourcePath;
        const std::wstring& destinationPath                                             = txn.destinationPath;
        ManagedSourceCleanupRecord& managedCleanupRecord                                = txn.managedCleanupRecord;
        wil::com_ptr<IFileWriter>& writer                                               = txn.writer;
        wil::com_ptr<IFileReader>& reader                                               = txn.reader;
        const uint64_t fileTotalBytes                                                   = txn.fileTotalBytes;
        const bool hasKnownFileTotalBytes                                               = txn.hasKnownFileTotalBytes;
        uint64_t& fileCompletedBytes                                                    = txn.fileCompletedBytes;
        std::optional<Common::Crypto::Blake3Hasher>& sourceHasher                       = txn.sourceHasher;
        std::vector<std::unique_ptr<Common::Crypto::ContentHasher>>& writerProofHashers = txn.writerProofHashers;
        BridgeCopyPerf& copyPerf                                                        = txn.copyPerf;
        HRESULT hr                                                                      = S_OK;

        const bool isMove         = managedCleanupRecord.armed;
        txn.isMove                = isMove;
        bool queryCommitSizeProof = isMove;
#ifdef ENABLE_TESTS
        constexpr const wchar_t* kDisableCommitSizeProofEnv = L"REDSALAMANDER_FILEOPS_BRIDGE_DISABLE_COMMIT_SIZE_PROOF";
        queryCommitSizeProof                                = queryCommitSizeProof && ! EnvironmentVariables::IsTruthyFlagSet(kDisableCommitSizeProofEnv);
#endif
        if (queryCommitSizeProof)
        {
            static_cast<void>(writer->QueryInterface(IID_PPV_ARGS(txn.commitSizeProof.addressof())));
        }

        wil::com_ptr<IFileWriterExpectedSize> expectedSizeWriter;
        if (SUCCEEDED(writer->QueryInterface(IID_PPV_ARGS(expectedSizeWriter.addressof()))) && expectedSizeWriter)
        {
            hr = expectedSizeWriter->SetExpectedSize(fileTotalBytes);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        if (verificationRequested)
        {
            sourceHasher.emplace();
        }
        hr = ReportProgress(
            sourcePath, destinationPath, fileTotalBytes, fileCompletedBytes, overallCompletedBytes.load(std::memory_order_acquire), progressStreamId);
        if (FAILED(hr))
        {
            return hr;
        }
        copyPerf.progressCalls         = 1;
        ULONGLONG lastProgressTick     = GetTickCount64();
        const auto maybeReportProgress = [&](uint64_t callCompletedBytes, bool force) noexcept -> HRESULT
        {
            if (! force)
            {
                const ULONGLONG nowTick = GetTickCount64();
                if (lastProgressTick != 0 && nowTick >= lastProgressTick && (nowTick - lastProgressTick) < ProgressIntervalMs())
                {
                    return S_OK;
                }
                lastProgressTick = nowTick;
            }
            else
            {
                lastProgressTick = GetTickCount64();
            }

            ++copyPerf.progressCalls;
            return ReportProgress(sourcePath, destinationPath, fileTotalBytes, fileCompletedBytes, callCompletedBytes, progressStreamId);
        };

        const auto copySerial = [&](const wil::com_ptr<IFileReader>& serialReader) noexcept -> HRESULT
        {
            if (! serialReader)
            {
                return E_POINTER;
            }

            for (;;)
            {
                task.WaitWhilePaused();
                if (CancelRequested())
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                unsigned long bytesRead    = 0;
                const uint64_t readStartUs = PerfNowUs();
                const HRESULT hrRead       = serialReader->Read(bufferIn, bufferBytesIn, &bytesRead);
                copyPerf.readUs += PerfElapsedUs(readStartUs);
                if (FAILED(hrRead))
                {
                    return hrRead;
                }
                if (bytesRead > bufferBytesIn)
                {
                    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }

                if (bytesRead == 0)
                {
                    break;
                }

                if (sourceHasher.has_value())
                {
                    sourceHasher.value().Update(std::span<const std::byte>(bufferIn, bytesRead));
                }
                for (const auto& proofHasher : writerProofHashers)
                {
                    static_cast<void>(proofHasher->Update(std::span<const std::byte>(bufferIn, bytesRead)));
                }

                const HRESULT hrWrite =
                    WriteChunk(*writer, bufferIn, bytesRead, fileTotalBytes, fileCompletedBytes, overallCompletedBytes, copyPerf.writeUs, maybeReportProgress);
                if (FAILED(hrWrite))
                {
                    return hrWrite;
                }
            }

            return S_OK;
        };

        const uint64_t copyStartUs = PerfNowUs();
        std::unique_ptr<std::byte[]> secondaryBuffer;
        if (ShouldUseBufferedPipeline(fileTotalBytes, bufferBytesIn))
        {
            secondaryBuffer.reset(new (std::nothrow) std::byte[bufferBytesIn]);
        }

        if (! secondaryBuffer)
        {
            hr = copySerial(reader);
        }
        else
        {
            struct BufferSlot final
            {
                std::byte* buffer       = nullptr;
                unsigned long bytesRead = 0;
                HRESULT readHr          = S_OK;
                bool ready              = false;
                bool eof                = false;
            };

            std::array<BufferSlot, 2> slots{{BufferSlot{bufferIn}, BufferSlot{secondaryBuffer.get()}}};
            std::mutex pipelineMutex;
            std::condition_variable pipelineCv;
            std::atomic<bool> pipelineStop{false};
            std::atomic<bool> readerFinished{false};
            std::atomic<uint64_t> readerWaitUs{0};
            std::atomic<uint64_t> readerReadUs{0};
            uint64_t writerWaitUs  = 0;
            uint64_t writerWriteUs = 0;

            std::jthread readerThread;
            try
            {
                readerThread = std::jthread([&, pipelineReader = reader](std::stop_token) noexcept
                {
                    size_t readIndex = 0;
                    for (;;)
                    {
                        task.WaitWhilePaused(&pipelineStop);
                        if (pipelineStop.load(std::memory_order_acquire) || CancelRequested())
                        {
                            break;
                        }

                        const uint64_t waitStartUs = PerfNowUs();
                        {
                            std::unique_lock lock(pipelineMutex);
                            pipelineCv.wait(lock, [&]() noexcept { return pipelineStop.load(std::memory_order_acquire) || ! slots[readIndex].ready; });
                        }
                        readerWaitUs.fetch_add(PerfElapsedUs(waitStartUs), std::memory_order_relaxed);

                        if (pipelineStop.load(std::memory_order_acquire) || CancelRequested())
                        {
                            break;
                        }

                        unsigned long bytesRead    = 0;
                        const uint64_t readStartUs = PerfNowUs();
                        HRESULT hrRead             = pipelineReader->Read(slots[readIndex].buffer, bufferBytesIn, &bytesRead);
                        readerReadUs.fetch_add(PerfElapsedUs(readStartUs), std::memory_order_relaxed);
                        if (SUCCEEDED(hrRead) && bytesRead > bufferBytesIn)
                        {
                            hrRead    = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                            bytesRead = 0;
                        }

                        {
                            std::scoped_lock lock(pipelineMutex);
                            slots[readIndex].bytesRead = bytesRead;
                            slots[readIndex].readHr    = hrRead;
                            slots[readIndex].eof       = SUCCEEDED(hrRead) && bytesRead == 0;
                            slots[readIndex].ready     = true;
                        }
                        pipelineCv.notify_all();

                        if (FAILED(hrRead) || bytesRead == 0)
                        {
                            break;
                        }

                        readIndex = (readIndex + 1u) % slots.size();
                    }

                    readerFinished.store(true, std::memory_order_release);
                    pipelineCv.notify_all();
                });
            }
            catch (const std::system_error&)
            {
                hr = copySerial(reader);
            }

            if (SUCCEEDED(hr) && readerThread.joinable())
            {
                size_t writeIndex = 0;
                for (;;)
                {
                    task.WaitWhilePaused();
                    if (CancelRequested())
                    {
                        hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        break;
                    }

                    const uint64_t waitStartUs = PerfNowUs();
                    {
                        std::unique_lock lock(pipelineMutex);
                        pipelineCv.wait(lock, [&]() noexcept {
                            return pipelineStop.load(std::memory_order_acquire) || readerFinished.load(std::memory_order_acquire) || slots[writeIndex].ready;
                        });
                    }
                    writerWaitUs += PerfElapsedUs(waitStartUs);

                    HRESULT slotHr              = S_OK;
                    unsigned long slotBytesRead = 0;
                    bool slotEof                = false;
                    bool slotReady              = false;
                    const bool stopped          = pipelineStop.load(std::memory_order_acquire);
                    const bool finished         = readerFinished.load(std::memory_order_acquire);
                    {
                        std::scoped_lock lock(pipelineMutex);
                        slotReady     = slots[writeIndex].ready;
                        slotHr        = slots[writeIndex].readHr;
                        slotBytesRead = slots[writeIndex].bytesRead;
                        slotEof       = slots[writeIndex].eof;
                    }

                    if (! slotReady)
                    {
                        hr = (stopped || CancelRequested()) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : (finished ? S_OK : HRESULT_FROM_WIN32(ERROR_CANCELLED));
                        break;
                    }

                    if (FAILED(slotHr))
                    {
                        hr = slotHr;
                        break;
                    }

                    if (slotEof)
                    {
                        break;
                    }

                    if (sourceHasher.has_value())
                    {
                        sourceHasher.value().Update(std::span<const std::byte>(slots[writeIndex].buffer, slotBytesRead));
                    }
                    for (const auto& proofHasher : writerProofHashers)
                    {
                        static_cast<void>(proofHasher->Update(std::span<const std::byte>(slots[writeIndex].buffer, slotBytesRead)));
                    }

                    hr = WriteChunk(*writer,
                                    slots[writeIndex].buffer,
                                    slotBytesRead,
                                    fileTotalBytes,
                                    fileCompletedBytes,
                                    overallCompletedBytes,
                                    writerWriteUs,
                                    maybeReportProgress);

                    {
                        std::scoped_lock lock(pipelineMutex);
                        slots[writeIndex].bytesRead = 0;
                        slots[writeIndex].readHr    = S_OK;
                        slots[writeIndex].eof       = false;
                        slots[writeIndex].ready     = false;
                    }
                    pipelineCv.notify_all();

                    if (FAILED(hr))
                    {
                        break;
                    }

                    writeIndex = (writeIndex + 1u) % slots.size();
                }

                pipelineStop.store(true, std::memory_order_release);
                task.WakePauseWaiters();
                pipelineCv.notify_all();
                readerThread.join();
                copyPerf.readerWaitUs += readerWaitUs.load(std::memory_order_acquire);
                copyPerf.writerWaitUs += writerWaitUs;
                copyPerf.readUs += readerReadUs.load(std::memory_order_acquire);
                copyPerf.writeUs += writerWriteUs;
            }
        }

        if (FAILED(hr))
        {
            return hr;
        }

        if (hasKnownFileTotalBytes && fileCompletedBytes != fileTotalBytes)
        {
            const HRESULT hrMismatch   = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            const std::wstring message = std::format(L"File copy size mismatch: expected {:L} bytes but wrote {:L} bytes.", fileTotalBytes, fileCompletedBytes);
            task.LogDiagnostic(
                FileOperationState::DiagnosticSeverity::Error, hrMismatch, L"bridge.integrity.sizeMismatch", message, sourcePath, destinationPath);
            NoteFailure(QualifiedItemFailurePhase::StageWrite, hrMismatch);
            return hrMismatch;
        }

        txn.sourceDigest = sourceHasher.has_value() ? sourceHasher.value().Finalize() : Common::Crypto::Blake3Digest{};

        uint64_t writerPositionBytes = 0u;
        const HRESULT positionHr     = writer->GetPosition(&writerPositionBytes);
        if (FAILED(positionHr) || writerPositionBytes != fileCompletedBytes)
        {
            const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            const std::wstring message =
                FAILED(positionHr)
                    ? std::format(L"Destination writer position could not be verified before commit (hr=0x{:08X}).", static_cast<unsigned long>(positionHr))
                    : std::format(L"Destination writer position mismatch before commit: reported {:L} bytes but persisted {:L} bytes.",
                                  fileCompletedBytes,
                                  writerPositionBytes);
            task.LogDiagnostic(
                FileOperationState::DiagnosticSeverity::Error, partialHr, L"bridge.integrity.writerPositionMismatch", message, sourcePath, destinationPath);
            NoteFailure(QualifiedItemFailurePhase::StageWrite, partialHr);
            return partialHr;
        }

        copyPerf.copyUs += PerfElapsedUs(copyStartUs);
        AccumulateBridgeCopyPerf(copyPerf, sourcePath, destinationPath, fileCompletedBytes, S_OK);

        if (hasKnownFileTotalBytes && fileCompletedBytes >= fileTotalBytes)
        {
            constexpr uint64_t kSmallFileCommitIndeterminateThresholdBytes = 1024ull * 1024ull;
            if (fileTotalBytes <= kSmallFileCommitIndeterminateThresholdBytes)
            {
                const uint64_t overallNow = overallCompletedBytes.load(std::memory_order_acquire);
                hr                        = ReportProgress(sourcePath, destinationPath, 0, 0, overallNow, progressStreamId);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
        }
        txn.state = PublicationTransaction::State::Written;
        return S_OK;
    }

    // Phase 7: commit and the proofs that follow it. Writer Commit, post-content metadata with
    // its consent, the committed-size proof, and the writer's content proof (which decides Move
    // cleanup on a route without object binding).
    [[nodiscard]] HRESULT CommitAndProve(PublicationTransaction& txn, bool& replaceRefused) noexcept
    {
        const std::wstring& sourcePath                                = txn.sourcePath;
        const std::wstring& destinationPath                           = txn.destinationPath;
        ManagedSourceCleanupRecord& managedCleanupRecord              = txn.managedCleanupRecord;
        wil::com_ptr<IFileWriter>& writer                             = txn.writer;
        FileOperations::BoundObjectAuthority& sourceMetadataAuthority = txn.sourceMetadataAuthority;
        FileSystemMetadataSnapshot& sourceMetadataSnapshot            = txn.sourceMetadataSnapshot;
        const bool useIdentityOwnedStage                              = txn.useIdentityOwnedStage;
        const bool identityLessReplace                                = txn.identityLessReplace;
        const std::wstring& writerPath                                = txn.useAtomicFinalWriter ? destinationPath : txn.stagePath;
        HRESULT hr                                                    = S_OK;

        hr = writer->Commit();
        if (FAILED(hr))
        {
            // A refused replacement or an unexpected occupant is a known non-commit: nothing was published.
            const bool definitiveNonCommit = hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) ||
                                             hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
            if (! useIdentityOwnedStage && ! definitiveNonCommit)
            {
                destinationPublicationUnknown.store(true, std::memory_order_release);
            }
            if (hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) && identityLessReplace)
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   hr,
                                   L"bridge.replace.occupantChanged",
                                   L"The destination changed after the replace decision; nothing was replaced.",
                                   sourcePath,
                                   destinationPath);
                replaceRefused = true;
                return hr;
            }
            Debug::Warning(L"CrossFileSystemBridge: destination writer Commit failed for '{}' via writer path '{}' (hr={:#x})",
                           destinationPath,
                           writerPath,
                           static_cast<unsigned long>(hr));
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                               hr,
                               L"bridge.commit",
                               L"Destination writer Commit failed while finalizing a staged bridge copy.",
                               sourcePath,
                               destinationPath);
            return hr;
        }
        if (! useIdentityOwnedStage)
        {
            anyDestinationPublished.store(true, std::memory_order_release);
        }
        txn.state = PublicationTransaction::State::Committed;

        if (txn.hasSourceMetadataSnapshot && txn.sourceMetadata)
        {
            FileSystemMetadataTransferResult finalMetadata{};
            finalMetadata.sizeBytes    = sizeof(finalMetadata);
            finalMetadata.firstFailure = S_OK;
            const uint32_t sourcePostContentFeatures =
                sourceMetadataSnapshot.presentFeatures &
                (FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES | FILESYSTEM_METADATA_MOTW | FILESYSTEM_METADATA_ALTERNATE_STREAMS |
                 FILESYSTEM_METADATA_EXTENDED_ATTRIBUTES | FILESYSTEM_METADATA_SECURITY);
            if (txn.destinationStageMetadata && txn.ownedStage)
            {
                const uint64_t metadataStartUs = PerfNowUs();
                const HRESULT metadataHr =
                    txn.sourceMetadata->TransferMetadataTo(txn.ownedStage.get(), FILESYSTEM_METADATA_TRANSFER_FINALIZE, &options, &finalMetadata);
                const std::wstring metadataDetail = std::format(L"phase=finalize changed=0x{:X}", finalMetadata.changedFeatures);
                Debug::Perf::Emit(L"FileOps.Metadata.ApplyUs",
                                  metadataDetail,
                                  PerfElapsedUs(metadataStartUs),
                                  finalMetadata.attemptedFeatures,
                                  finalMetadata.lostFeatures,
                                  metadataHr);
                if (FAILED(metadataHr))
                {
                    finalMetadata.attemptedFeatures |= sourcePostContentFeatures;
                    finalMetadata.lostFeatures |= sourcePostContentFeatures & ~FILESYSTEM_METADATA_SECURITY;
                    finalMetadata.changedFeatures |= sourcePostContentFeatures & FILESYSTEM_METADATA_SECURITY;
                    finalMetadata.firstFailure = metadataHr;
                }
            }
            else
            {
                finalMetadata.attemptedFeatures |= sourcePostContentFeatures;
                finalMetadata.lostFeatures |= sourcePostContentFeatures & ~FILESYSTEM_METADATA_SECURITY;
                finalMetadata.changedFeatures |= sourcePostContentFeatures & FILESYSTEM_METADATA_SECURITY;
                if (sourcePostContentFeatures != 0u)
                {
                    finalMetadata.firstFailure = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
            }
            ReportMetadataOutcome(sourcePath, destinationPath, finalMetadata);

            constexpr uint32_t kSecuritySignificantLoss =
                FILESYSTEM_METADATA_MOTW | FILESYSTEM_METADATA_ALTERNATE_STREAMS | FILESYSTEM_METADATA_EXTENDED_ATTRIBUTES;
            if (managedCleanupRecord.armed && (finalMetadata.lostFeatures & kSecuritySignificantLoss) != 0u)
            {
                const DeferredConsentResult consent = RequestBridgeConsent(FileOperations::DeferredConsentRisk::MetadataLoss,
                                                                           FAILED(finalMetadata.firstFailure) ? finalMetadata.firstFailure : S_FALSE,
                                                                           sourcePath,
                                                                           destinationPath,
                                                                           &sourceMetadataAuthority,
                                                                           &sourceMetadataSnapshot);
                if (FAILED(consent.status))
                {
                    return consent.status;
                }
                if (consent.action == ConflictAction::RetainSource)
                {
                    MarkManagedSourceRetained(
                        sourcePath, destinationPath, S_FALSE, L"Copied; source kept because security-significant metadata was not preserved.");
                    managedCleanupRecord = {};
                }
                else if (consent.action != ConflictAction::Proceed)
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
            }
        }

        const bool isMove                                         = txn.isMove;
        const bool hasKnownFileTotalBytes                         = txn.hasKnownFileTotalBytes;
        const uint64_t fileTotalBytes                             = txn.fileTotalBytes;
        wil::com_ptr<IFileWriterCommitSizeProof>& commitSizeProof = txn.commitSizeProof;
        bool committedSizeProven                                  = false;
        if (hasKnownFileTotalBytes && commitSizeProof)
        {
            uint64_t committedSizeBytes = 0u;
            const HRESULT proofHr       = commitSizeProof->GetCommittedSize(&committedSizeBytes);
            if (SUCCEEDED(proofHr) && committedSizeBytes == fileTotalBytes)
            {
                committedSizeProven = true;
                task._perf.bridgeCommitSizeProofCount.fetch_add(1u, std::memory_order_relaxed);
            }
            else
            {
                Debug::Warning(L"CrossFileSystemBridge: committed-size proof was invalid for '{}' (proofHr={:#x}, expected={}, reported={}); "
                               L"falling back to final-path verification.",
                               destinationPath,
                               static_cast<unsigned long>(proofHr),
                               fileTotalBytes,
                               committedSizeBytes);
            }
        }
        txn.committedSizeProven = committedSizeProven;
        if (isMove && hasKnownFileTotalBytes && ! committedSizeProven)
        {
            task._perf.bridgeCommitSizeProofFallbackCount.fetch_add(1u, std::memory_order_relaxed);
        }
        // R3-2: the provider's own digest of the published object against the bytes streamed.
        if (txn.writerProof && ! txn.writerProofHashers.empty())
        {
            txn.writerProofMatched = EvaluateWriterContentProof(*txn.writerProof, txn.writerProofHashers, txn.fileCompletedBytes, sourcePath, destinationPath);
        }
        txn.writerProof.reset();
        txn.writerProofHashers.clear();
        if (managedCleanupRecord.armed && txn.identityLess && txn.writerProofMatched != true)
        {
            MarkManagedSourceRetained(sourcePath,
                                      destinationPath,
                                      txn.writerProofMatched.has_value() ? HRESULT_FROM_WIN32(ERROR_CRC) : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                      L"Copied; source kept because the destination could not prove the published content.");
            managedCleanupRecord = {};
        }
        writer.reset();
        commitSizeProof.reset();
        return S_OK;
    }

    // Phase 8: publication under the final name (the owned stage's exact PublishAs, or the
    // atomic-final writer already promoted), the post-publication size fact, and basic metadata.
    [[nodiscard]] HRESULT Publish(PublicationTransaction& txn) noexcept
    {
        const std::wstring& sourcePath      = txn.sourcePath;
        const std::wstring& destinationPath = txn.destinationPath;
        const bool isMove                   = txn.isMove;
        const bool hasKnownFileTotalBytes   = txn.hasKnownFileTotalBytes;
        const bool committedSizeProven      = txn.committedSizeProven;
        HRESULT hr                          = S_OK;

        if (txn.useIdentityOwnedStage)
        {
            hr = PublishOwnedStage(sourcePath,
                                   destinationPath,
                                   *txn.ownedStage,
                                   txn.expectedDestination.get(),
                                   txn.overwriteGranted,
                                   txn.replaceReadOnlyGranted,
                                   txn.replaceLinkGranted,
                                   txn.publishedAuthority,
                                   txn.promoted,
                                   txn.ownedStageCleanupAllowed);
            if (FAILED(hr))
            {
                return hr;
            }
        }
        else
        {
            txn.promoted = true;
        }
        txn.state = PublicationTransaction::State::Published;
        task.NoteLiveOutputPublished(destinationPath);

#ifdef ENABLE_TESTS
        // Cinderstar package 20 reviewed call-site exception: replacement must occur after
        // publication and before by-path verification to prove cleanup preserves a new owner.
        MaybeReplacePublishedBridgeDestinationBeforeVerificationForSelfTest(destinationIo, destinationPath);
#endif

        if (hasKnownFileTotalBytes && (! isMove || ! committedSizeProven))
        {
            wil::com_ptr<IFileReader> destinationReader;
            uint64_t destinationSizeBytes = 0;
            HRESULT hrDestinationSize     = S_OK;
            if (txn.publishedAuthority)
            {
                FileSystemBoundObjectSnapshot publishedSnapshot{};
                publishedSnapshot.sizeBytes = sizeof(publishedSnapshot);
                hrDestinationSize           = txn.publishedAuthority->GetSnapshot(&publishedSnapshot);
                if (SUCCEEDED(hrDestinationSize))
                {
                    if (publishedSnapshot.sizeBytes != sizeof(publishedSnapshot) || publishedSnapshot.objectId == nullptr ||
                        publishedSnapshot.objectIdBytes == 0u || publishedSnapshot.objectIdBytes > 64u * 1024u ||
                        (publishedSnapshot.revisionIdBytes != 0u && publishedSnapshot.revisionId == nullptr) ||
                        publishedSnapshot.kind != FILESYSTEM_BOUND_REGULAR_FILE)
                    {
                        hrDestinationSize = E_UNEXPECTED;
                    }
                    else
                    {
                        destinationSizeBytes = publishedSnapshot.committedSizeBytes;
                    }
                }
            }
            else
            {
                hrDestinationSize = VerifyPublishedDestinationSize(destinationPath, destinationSizeBytes, destinationReader);
            }
            if (hrDestinationSize == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hrDestinationSize == E_ABORT)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            if (FAILED(hrDestinationSize) || destinationSizeBytes != txn.fileTotalBytes)
            {
                const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                const std::wstring message =
                    FAILED(hrDestinationSize)
                        ? std::format(L"Cross-filesystem {} could not re-stat destination after promote (hr=0x{:08X}).",
                                      isMove ? L"MOVE" : L"COPY",
                                      static_cast<unsigned long>(hrDestinationSize))
                        : std::format(L"Cross-filesystem {} destination size mismatch after promote: expected {:L} bytes but destination has {:L} bytes.",
                                      isMove ? L"MOVE" : L"COPY",
                                      txn.fileTotalBytes,
                                      destinationSizeBytes);
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   partialHr,
                                   L"bridge.integrity.destinationSizeMismatch",
                                   message,
                                   sourcePath,
                                   destinationPath);
                // Publication has completed. Preserve the exact publication and any
                // concurrent pathname owner; verification never grants cleanup authority.
                return partialHr;
            }
        }

        // Publishing the owned stage deliberately clears its temporary/hidden staging
        // attributes. Reapply the source's ordinary attributes and timestamps through the
        // exact published-object authority after the rename, even when the provider already
        // finalized richer metadata on the stage.
        if (txn.hasSourceBasicInfo)
        {
            const HRESULT hrSetBasic =
                txn.publishedAuthority ? txn.publishedAuthority->SetBasicInformation(&txn.sourceBasicInfo) : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            if (FAILED(hrSetBasic) && hrSetBasic != E_NOTIMPL && hrSetBasic != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
            {
                Debug::Warning(
                    L"CrossFileSystemBridge: SetFileBasicInformation failed for '{}' (hr={:#x})", destinationPath, static_cast<unsigned long>(hrSetBasic));
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   hrSetBasic,
                                   L"bridge.metadata.write",
                                   L"SetFileBasicInformation failed for destination file.",
                                   sourcePath,
                                   destinationPath);
            }
            else if (hrSetBasic == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   hrSetBasic,
                                   L"bridge.metadata.authorityUnavailable",
                                   L"Destination metadata was not written because exact no-follow publication authority is unavailable.",
                                   sourcePath,
                                   destinationPath);
            }
        }
        return S_OK;
    }

    // Phase 9: the final progress fact and the verification of the published object.
    // S_FALSE = verification released the cleanup record (source kept); failures as today.
    [[nodiscard]] HRESULT VerifyPublished(PublicationTransaction& txn,
                                          std::byte* bufferIn,
                                          unsigned long bufferBytesIn,
                                          uint64_t progressStreamId,
                                          std::atomic<uint64_t>& overallCompletedBytes) noexcept
    {
        const std::wstring& sourcePath      = txn.sourcePath;
        const std::wstring& destinationPath = txn.destinationPath;

        const uint64_t overallFinal   = overallCompletedBytes.load(std::memory_order_acquire);
        const uint64_t finalTotal     = txn.hasKnownFileTotalBytes ? txn.fileTotalBytes : txn.fileCompletedBytes;
        const uint64_t finalCompleted = txn.fileCompletedBytes;

        HRESULT hr = ReportProgress(sourcePath, destinationPath, finalTotal, finalCompleted, overallFinal, progressStreamId);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = VerifyPublishedContent(sourcePath,
                                    destinationPath,
                                    txn.fileTotalBytes,
                                    txn.sourceDigest,
                                    txn.publishedAuthority.get(),
                                    bufferIn,
                                    bufferBytesIn,
                                    txn.managedCleanupRecord,
                                    txn.writerProofMatched);
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (FAILED(hr))
        {
            return hr;
        }
        if (hr == S_FALSE)
        {
            return S_FALSE;
        }
        txn.state = PublicationTransaction::State::Verified;
        return S_OK;
    }

    // The exact Move source cleanup after publication.
    [[nodiscard]] HRESULT FinishCleanup(PublicationTransaction& txn) noexcept
    {
        return FinalizeManagedSourceCleanup(txn.sourcePath, txn.destinationPath, txn.managedCleanupRecord);
    }

    // The one regular-file publication: the phases above in order, on one transaction record.
    // Every early return is the same status the former single function returned at that point.
    HRESULT PumpAndPublishFileOnce(const std::wstring& sourcePath,
                                   std::wstring destinationPath,
                                   std::byte* bufferIn,
                                   unsigned long bufferBytesIn,
                                   uint64_t progressStreamId,
                                   std::atomic<uint64_t>& overallCompletedBytes,
                                   bool adoptFileSizeAsTotalWhenUnknown,
                                   bool& replaceRefused) noexcept
    {
        replaceRefused = false;
        if (! bufferIn || bufferBytesIn == 0)
        {
            return E_INVALIDARG;
        }

#ifdef ENABLE_TESTS
        // FIR-2 reviewed call-site exception: this injects a whole-copy worker failure before stream I/O
        // to exercise scheduler/status propagation, not an IFileReader or IFileWriter contract failure.
        if (ConsumeBridgeFailNextFileCopyForSelfTest())
        {
            const HRESULT hrInjected = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                               hrInjected,
                               L"bridge.selftest.filecopy.fail",
                               L"Selftest injected a bridge file-copy failure.",
                               sourcePath,
                               destinationPath);
            return hrInjected;
        }
#endif

        ConnectionConcurrencyLimiter::Permit permit1;
        ConnectionConcurrencyLimiter::Permit permit2;
        const HRESULT hrPermits = AcquireCopyMovePermits(permit1, permit2);
        if (FAILED(hrPermits))
        {
            return hrPermits;
        }

        PublicationTransaction txn(sourcePath, std::move(destinationPath));
        const auto cleanupStage = wil::scope_exit([&] noexcept
        {
            if (! txn.promoted && txn.ownedStageCleanupAllowed && txn.ownedStage)
            {
                AbortOwnedStage(*txn.ownedStage, txn.stagePath, sourcePath, txn.destinationPath);
            }
        });

        HRESULT hr = AdmitDestination(txn);
        if (hr != S_OK)
        {
            return hr;
        }
        hr = BindSource(txn, adoptFileSizeAsTotalWhenUnknown);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = RouteWriter(txn);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = PrepareStage(txn, progressStreamId, replaceRefused);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = TransferPreContentMetadata(txn);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = Pump(txn, bufferIn, bufferBytesIn, progressStreamId, overallCompletedBytes);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = CommitAndProve(txn, replaceRefused);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = Publish(txn);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = VerifyPublished(txn, bufferIn, bufferBytesIn, progressStreamId, overallCompletedBytes);
        if (hr != S_OK)
        {
            return hr;
        }
        return FinishCleanup(txn);
    }

    HRESULT CopyFile(const std::wstring& sourcePath, std::wstring destinationPath) noexcept
    {
        if (! buffer || bufferBytes == 0)
        {
            return FAILED(bufferAllocationHr) ? bufferAllocationHr : E_OUTOFMEMORY;
        }

        if (! connectionLimitsInitialized)
        {
            InitializeConnectionLimits(sourcePath, destinationPath);
        }
        std::atomic<uint64_t> overallCompletedBytes{completedBytes};
        const HRESULT hr = PumpAndPublishFile(sourcePath, destinationPath, buffer.get(), bufferBytes, 0, overallCompletedBytes, true);
        if (SUCCEEDED(hr))
        {
            completedBytes = overallCompletedBytes.load(std::memory_order_acquire);
        }
        return hr;
    }

    struct SemanticLinkPayload final
    {
        FileSystemLinkInformation information{};
        std::vector<wchar_t> targetBuffer;
    };

    [[nodiscard]] HRESULT ReadPreparedLinkPayload(PreparedLinkRecord& record, SemanticLinkPayload& payload) noexcept
    {
        FileOperations::BoundObjectAuthority* authority = record.ReadAuthority();
        if (! sourceBinding || authority == nullptr || ! authority->boundObject)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        FileSystemLinkTransform transform{};
        transform.sizeBytes           = sizeof(transform);
        transform.sourceLinkPath      = record.sourcePath.c_str();
        transform.destinationLinkPath = record.destinationPath.c_str();
        transform.sourceRootPath      = rootSourcePath.c_str();
        transform.destinationRootPath = rootDestinationPath.c_str();

        payload                       = {};
        payload.information.sizeBytes = sizeof(payload.information);
        HRESULT hr                    = sourceBinding->ReadBoundLink(authority->boundObject.get(), &transform, &options, &payload.information);
        if (hr != HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER) || payload.information.targetLengthUtf16 == 0u ||
            payload.information.targetLengthUtf16 > 32'767u)
        {
            return FAILED(hr) ? hr : E_UNEXPECTED;
        }
        // Literal Preserve: the provider reports the stored text with the outside-root mapping and
        // no source-relative component.
        if (payload.information.targetMapping != FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT || payload.information.sourceRelativeTargetLengthUtf16 != 0u)
        {
            return E_UNEXPECTED;
        }

        payload.targetBuffer.assign(static_cast<size_t>(payload.information.targetLengthUtf16) + 1u, L'\0');
        payload.information.targetBuffer        = payload.targetBuffer.data();
        payload.information.targetCapacityUtf16 = static_cast<uint32_t>(payload.targetBuffer.size());
        hr                                      = sourceBinding->ReadBoundLink(authority->boundObject.get(), &transform, &options, &payload.information);
        if (FAILED(hr))
        {
            return hr;
        }
        if (payload.information.targetMapping != FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT || payload.information.targetLengthUtf16 == 0u ||
            payload.information.targetLengthUtf16 >= payload.information.targetCapacityUtf16 ||
            payload.targetBuffer[payload.information.targetLengthUtf16] != L'\0')
        {
            return E_UNEXPECTED;
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT PublishPreparedLink(PreparedLinkRecord& record, SemanticLinkPayload& payload) noexcept
    {
        std::wstring stagePath;
        wil::com_ptr<IFileSystemBoundObject> ownedStage;
        wil::com_ptr<IFileSystemBoundObject> publishedAuthority;
        bool promoted                 = false;
        bool ownedStageCleanupAllowed = false;
        const auto cleanupStage       = wil::scope_exit([&] noexcept
        {
            if (! promoted && ownedStageCleanupAllowed && ownedStage)
            {
                AbortOwnedStage(*ownedStage, stagePath, record.sourcePath, record.destinationPath);
            }
        });

        HRESULT hr = CreateOwnedLinkStage(record.destinationPath, 0u, payload.information, stagePath, ownedStage);
        if (FAILED(hr))
        {
            return hr;
        }
        ownedStageCleanupAllowed = true;

        // R3-3: the link payload publishes through the shared owner; a partial post-publication
        // cleanup keeps its StageCommit phase.
        hr = PublishOwnedStageAs(ownedStage.get(),
                                 {.sourcePath             = record.sourcePath,
                                  .destinationPath        = record.destinationPath,
                                  .expectedDestination    = record.expectedDestination.get(),
                                  .overwriteGranted       = record.overwriteGranted,
                                  .replaceReadOnlyGranted = record.replaceReadOnlyGranted,
                                  .replaceLinkGranted     = record.replaceLinkGranted,
                                  .unknownCategory        = L"bridge.linkPublication.unknown",
                                  .unknownMessage  = L"Link publication outcome is unknown; destination and visible stage are retained for reconciliation.",
                                  .partialCategory = L"bridge.linkPublication.cleanupPartial",
                                  .partialMessage  = L"Link was published, but exact post-publication cleanup was incomplete.",
                                  .partialPhase    = QualifiedItemFailurePhase::StageCommit},
                                 publishedAuthority,
                                 promoted,
                                 ownedStageCleanupAllowed);
        if (FAILED(hr))
        {
            return hr;
        }

        FileSystemBoundObjectSnapshot publishedSnapshot{};
        publishedSnapshot.sizeBytes = sizeof(publishedSnapshot);
        hr                          = publishedAuthority->GetSnapshot(&publishedSnapshot);
        if (FAILED(hr) || publishedSnapshot.sizeBytes != sizeof(publishedSnapshot) || publishedSnapshot.objectId == nullptr ||
            publishedSnapshot.objectIdBytes == 0u || publishedSnapshot.objectIdBytes > 64u * 1024u ||
            (publishedSnapshot.revisionIdBytes != 0u && publishedSnapshot.revisionId == nullptr) || publishedSnapshot.kind != FILESYSTEM_BOUND_LINK)
        {
            return FAILED(hr) ? hr : E_UNEXPECTED;
        }

        if (record.hasBasicInformation)
        {
            const HRESULT basicInfoHr = publishedAuthority->SetBasicInformation(&record.basicInformation);
            if (FAILED(basicInfoHr) && basicInfoHr != E_NOTIMPL && basicInfoHr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
            {
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   basicInfoHr,
                                   L"bridge.linkMetadata.write",
                                   L"SetBasicInformation failed for the published link object.",
                                   record.sourcePath,
                                   record.destinationPath);
            }
        }

        const uint64_t overallCompleted = completedBytesAtomic ? completedBytesAtomic->load(std::memory_order_acquire) : completedBytes;
        hr                              = ReportProgress(record.sourcePath, record.destinationPath, 0u, 0u, overallCompleted, 0u);
        if (FAILED(hr))
        {
            return hr;
        }
        return FinalizeManagedSourceCleanup(record.sourcePath, record.destinationPath, record.cleanupRecord);
    }

    // A walker pushes its own frame for a directory namespace. The parallel producer also
    // requests file routing: its primary buffer is released, so ordinary placeholders must
    // enter the same worker queue as other files instead of copying on the producer.
    HRESULT CopyLink(const std::wstring& sourcePath,
                     std::wstring destinationPath,
                     bool sourceIsDirectory,
                     uint64_t directoryDepth   = 0u,
                     bool* traverseAsDirectory = nullptr,
                     bool* enqueueAsFile       = nullptr) noexcept
    {
        const bool isRoot   = traverseAsDirectory == nullptr;
        const auto skipLink = [&]() noexcept -> HRESULT
        {
            const HRESULT skipHr = MarkReparseSkipped(sourcePath, destinationPath, sourceIsDirectory, isRoot);
            return FAILED(skipHr) ? skipHr : S_FALSE;
        };
        const auto unsupported = [&](HRESULT status, std::wstring_view reason) noexcept
        {
            const HRESULT normalizedStatus = status == E_NOINTERFACE ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) : status;
            if (normalizedStatus == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
            {
                unsupportedReparseEncountered = true;
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                   normalizedStatus,
                                   L"bridge.reparse.unsupported",
                                   std::wstring(reason),
                                   sourcePath,
                                   destinationPath);
            }
            return normalizedStatus;
        };

        if (CancelRequested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (! sourceBinding)
        {
            if (reparsePointPolicy == ReparsePointPolicy::Skip)
            {
                // Without object binding the reparse attribute is the only classification available.
                return skipLink();
            }
            return unsupported(sourceBindingQueryHr == E_NOINTERFACE ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                                                                     : (FAILED(sourceBindingQueryHr) ? sourceBindingQueryHr : E_UNEXPECTED),
                               L"Source provider cannot read an exact no-follow link payload; Preserve cannot continue.");
        }
        constexpr FileSystemBindFlags classificationFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
        FileOperations::ObjectBindingResult classification =
            FileOperations::BindObjectAuthority(&sourceFs, sourcePath, sourcePathProfileId, classificationFlags);
        if (classification.state != FileOperations::ObjectBindingState::Bound)
        {
            return unsupported(FAILED(classification.status) ? classification.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                               L"Source provider could not bind the link object no-follow; Preserve cannot continue.");
        }
        if (classification.authority.kind == FILESYSTEM_BOUND_REGULAR_FILE)
        {
            // Not every reparse tag is a link. Cloud/WOF/dedup placeholders remain their
            // underlying file object and must use the ordinary bound content route.
            if (enqueueAsFile != nullptr)
            {
                *enqueueAsFile = true;
                return S_OK;
            }
            return CopyFile(sourcePath, destinationPath);
        }
        if (classification.authority.kind == FILESYSTEM_BOUND_DIRECTORY)
        {
            // Name-surrogate tags are links; other directory reparse tags remain a
            // directory namespace and are traversed without following another target.
            if (traverseAsDirectory != nullptr)
            {
                *traverseAsDirectory = true;
                return S_OK;
            }
            return CopyDirectorySequential(sourcePath, destinationPath, directoryDepth);
        }
        if (classification.authority.kind != FILESYSTEM_BOUND_LINK)
        {
            return unsupported(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), L"Source reparse kind has no executable Preserve contract.");
        }
        if (reparsePointPolicy == ReparsePointPolicy::Skip)
        {
            // Skip applies to name-surrogate links only; placeholders above stay ordinary objects.
            return skipLink();
        }
        if (! destinationBinding)
        {
            return unsupported(destinationBindingQueryHr == E_NOINTERFACE ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                                                                          : (FAILED(destinationBindingQueryHr) ? destinationBindingQueryHr : E_UNEXPECTED),
                               L"Destination provider cannot create an exact owned link stage; Preserve cannot continue.");
        }

        if (! connectionLimitsInitialized)
        {
            InitializeConnectionLimits(sourcePath, destinationPath);
        }
        ConnectionConcurrencyLimiter::Permit permit1;
        ConnectionConcurrencyLimiter::Permit permit2;
        const HRESULT permitHr = AcquireCopyMovePermits(permit1, permit2);
        if (FAILED(permitHr))
        {
            return permitHr;
        }

        bool overwriteGranted       = false;
        bool replaceReadOnlyGranted = false;
        bool replaceLinkGranted     = false;
        wil::com_ptr<IFileSystemBoundObject> expectedDestination;
        std::optional<FileSystemBasicInformation> linkReplaceExpectation; // links are replaced only with bound authority
        HRESULT hr = ValidateDestinationOverwritePolicy(
            sourcePath, destinationPath, overwriteGranted, replaceReadOnlyGranted, replaceLinkGranted, expectedDestination, linkReplaceExpectation);
        if (FAILED(hr))
        {
            return hr;
        }
        if (hr == S_FALSE)
        {
            RecordSkippedDestinationCollision(sourcePath,
                                              destinationPath,
                                              HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                                              L"bridge.linkConflict.skip",
                                              L"Destination already exists; preserved link was skipped on user request.");
            return S_FALSE;
        }

        PreparedLinkRecord record{};
        record.sourcePath             = sourcePath;
        record.destinationPath        = destinationPath;
        record.expectedDestination    = std::move(expectedDestination);
        record.overwriteGranted       = overwriteGranted;
        record.replaceReadOnlyGranted = replaceReadOnlyGranted;
        record.replaceLinkGranted     = replaceLinkGranted;
        if (sourceCleanupPermitted)
        {
            hr = PrepareManagedSourceAuthority(sourcePath, destinationPath, FILESYSTEM_BOUND_LINK, record.cleanupRecord);
            if (FAILED(hr))
            {
                return hr;
            }
            if (hr != S_OK || ! record.cleanupRecord.armed)
            {
                // Per-item destructive authority may be weaker than the admitted profile.
                // Preserve the link through the exact read authority and keep the source.
                record.copyAuthority = std::move(classification.authority);
            }
        }
        else
        {
            record.copyAuthority = std::move(classification.authority);
        }
        FileOperations::BoundObjectAuthority* sourceAuthority = record.ReadAuthority();
        if (! sourceAuthority || ! sourceAuthority->boundObject)
        {
            return E_UNEXPECTED;
        }

        record.basicInformation.sizeBytes = sizeof(record.basicInformation);
        record.hasBasicInformation        = SUCCEEDED(sourceAuthority->boundObject->GetBasicInformation(&record.basicInformation));

        SemanticLinkPayload payload{};
        hr = ReadPreparedLinkPayload(record, payload);
        if (FAILED(hr))
        {
            return unsupported(hr, L"Source provider failed the bounded literal link payload read.");
        }
        return PublishPreparedLink(record, payload);
    }

    [[nodiscard]] static bool IsProviderCollisionStatus(HRESULT hr) noexcept
    {
        return hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) || hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) ||
               hr == HRESULT_FROM_WIN32(ERROR_REPARSE_POINT_ENCOUNTERED) || hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) ||
               hr == HRESULT_FROM_WIN32(ERROR_DATATYPE_MISMATCH);
    }

    // Rename merge: one provider Native rename per child. A regular directory that meets a regular
    // directory recurses instead. A provider with the conflict callback contract (Local) prompts
    // through the task itself; a provider that only reports the collision gets the bridge's prompt.
    HRESULT RelocateChild(const std::wstring& sourcePath, std::wstring destinationPath, bool isDirectory, bool isReparse, bool& traverseAsDirectory) noexcept
    {
        traverseAsDirectory                 = false;
        const bool sourceIsRegularDirectory = isDirectory && (! isReparse || IsRegularDirectoryObject(sourceFs, sourceIo, sourcePathProfileId, sourcePath));
        FileSystemOptions moveOptions       = options;
        moveOptions.moveMode                = FILESYSTEM_MOVE_NATIVE_ONLY;
        FileSystemFlags moveFlags           = flags;
        for (;;)
        {
            task.WaitWhilePaused();
            if (CancelRequested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            moveOptions.bandwidthLimitBytesPerSecond = bandwidthLimitBytesPerSecond.load(std::memory_order_acquire);
            const HRESULT hr                         = sourceFs.MoveItem(sourcePath.c_str(), destinationPath.c_str(), moveFlags, &moveOptions, &task, cookie);
            if (SUCCEEDED(hr))
            {
                anyDestinationPublished.store(true, std::memory_order_release);
                return S_OK;
            }
            if (task.IsCancellationOutcome(hr))
            {
                return hr;
            }
            if (hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
            {
                RecordSkippedDestinationCollision(
                    sourcePath, destinationPath, hr, L"bridge.renameMerge.skip", L"Skipped at its conflict prompt; the item stays in the source folder.");
                return S_FALSE;
            }
            if (! IsProviderCollisionStatus(hr))
            {
                NoteFailure(QualifiedItemFailurePhase::ProviderNative, hr);
                return hr;
            }
            if (sourceIsRegularDirectory && IsRegularDirectoryObject(destinationFs, destinationIo, destinationPathProfileId, destinationPath))
            {
                traverseAsDirectory = true;
                return S_OK;
            }

            bool overwriteGranted       = false;
            bool replaceReadOnlyGranted = false;
            bool replaceLinkGranted     = false;
            bool keepBothRequested      = false;
            wil::com_ptr<IFileSystemBoundObject> expectedDestination;
            const HRESULT promptHr = PromptDestinationCollision(
                sourcePath, destinationPath, hr, overwriteGranted, replaceReadOnlyGranted, replaceLinkGranted, keepBothRequested, expectedDestination);
            if (keepBothRequested)
            {
                const HRESULT keepBothHr = SelectKeepBothDestination(sourcePath, sourceIsRegularDirectory, destinationPath);
                if (FAILED(keepBothHr))
                {
                    return keepBothHr;
                }
                continue;
            }
            if (promptHr == S_FALSE)
            {
                RecordSkippedDestinationCollision(
                    sourcePath, destinationPath, hr, L"bridge.renameMerge.skip", L"Destination already exists; skipped on user request.");
                return S_FALSE;
            }
            if (FAILED(promptHr))
            {
                return promptHr;
            }
            uint32_t grantedFlags = FILESYSTEM_FLAG_ALLOW_OVERWRITE;
            if (replaceReadOnlyGranted)
            {
                grantedFlags |= FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY;
            }
            if (replaceLinkGranted)
            {
                grantedFlags |= FILESYSTEM_FLAG_ALLOW_REPLACE_LINK;
            }
            moveFlags = static_cast<FileSystemFlags>(static_cast<uint32_t>(moveFlags) | grantedFlags);
        }
    }

    // R4-T1: one directory of the sequential walk. The walker below keeps these frames on an explicit
    // stack instead of the C++ stack, so tree depth is bounded by the retained-metadata budget, not
    // by recursion. Every field mirrors a local of the former recursive function; the exit actions
    // (restore created-directory metadata, release retained bytes)
    // run when the frame pops, in the order the former scope guards ran.
    struct SequentialDirectoryFrame final
    {
        std::wstring sourcePath;
        std::wstring destinationPath;
        uint64_t depth = 0u;
        bool started   = false;

        uint64_t retainedMetadataBytes     = 0u;
        bool metadataReserved              = false;
        bool hadChildFailure               = false;
        bool hadRetainedChild              = false;
        uint64_t retentionGenerationBefore = 0u;
        BridgeChildNameSet registeredChildNames;

        ManagedSourceCleanupRecord managedDirectoryCleanup{};
        bool createdDestination = false;
        wil::com_ptr<IFileSystemBoundObject> createdDestinationAuthority;
        FileSystemBasicInformation directoryBasicInfo{};
        bool directoryMetadataCaptured = false;
        bool directoryMetadataRestored = false;

        wil::com_ptr<IFilesInformation> info; // owns the FileInfo buffer the child views point into
        std::vector<FileInfo*> childEntries;
        std::vector<std::wstring_view> childNames;
        size_t childIndex = 0u;
    };

    void RestoreSequentialDirectoryMetadata(SequentialDirectoryFrame& frame) noexcept
    {
        if (! frame.directoryMetadataRestored)
        {
            RestoreCreatedDirectoryMetadata(
                frame.sourcePath, frame.destinationPath, frame.directoryMetadataCaptured, frame.directoryBasicInfo, frame.createdDestinationAuthority.get());
            frame.directoryMetadataRestored = true;
        }
    }

    // Exit actions of one frame, in the order the former scope guards ran on return.
    void PopSequentialDirectoryFrame(std::vector<std::unique_ptr<SequentialDirectoryFrame>>& frames) noexcept
    {
        SequentialDirectoryFrame& frame = *frames.back();
        if (frame.started)
        {
            RestoreSequentialDirectoryMetadata(frame);
        }
        if (frame.metadataReserved)
        {
            ReleaseTraversalMetadata(frame.retainedMetadataBytes);
        }
        frames.pop_back();
    }

    // Prologue of one directory: everything the recursive function did before its child loop.
    // Returns S_OK when the frame has children to process; any other value is the frame's result.
    HRESULT StartSequentialDirectoryFrame(SequentialDirectoryFrame& frame) noexcept
    {
        const std::wstring& sourcePath = frame.sourcePath;
        std::wstring& destinationPath  = frame.destinationPath; // EnsureDestinationDirectory may rename it (keep both)

        if (CancelRequested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        HRESULT hr = ValidateTraversalDepth(frame.depth, sourcePath, destinationPath);
        if (FAILED(hr))
        {
            return hr;
        }

        frame.retainedMetadataBytes = AddRetainedBytes(RetainedPathBytes(sourcePath, destinationPath), sizeof(FileSystemBasicInformation));
        NoteTraversalMetadataRetained(frame.retainedMetadataBytes);
        frame.metadataReserved     = true;
        frame.registeredChildNames = MakeChildNameSet();

        const bool continueOnError      = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
        frame.retentionGenerationBefore = managedSourceRetentionGeneration.load(std::memory_order_acquire);

        hr = PrepareManagedSourceAuthority(sourcePath, destinationPath, FILESYSTEM_BOUND_DIRECTORY, frame.managedDirectoryCleanup);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = EnsureDestinationDirectory(sourcePath, destinationPath, &frame.createdDestination, std::addressof(frame.createdDestinationAuthority));
        if (hr == S_FALSE)
        {
            return S_FALSE;
        }
        if (FAILED(hr))
        {
            return hr;
        }

        frame.directoryMetadataCaptured = CaptureCreatedDirectoryMetadata(sourcePath, destinationPath, frame.createdDestination, frame.directoryBasicInfo);
        frame.started                   = true; // from here on the pop restores the created directory's metadata

        task._bridgeSourceDirectoryEnumerationCount.fetch_add(1u, std::memory_order_relaxed);
        hr = sourceFs.ReadDirectoryInfo(sourcePath.c_str(), frame.info.addressof());
        if (FAILED(hr))
        {
            return hr;
        }

        FileInfo* entry = nullptr;
        hr              = frame.info->GetBuffer(&entry);
        if (FAILED(hr))
        {
            return hr;
        }
        if (entry == nullptr)
        {
            // An empty directory legitimately yields (nullptr, S_OK) per the IFilesInformation
            // contract (FileSystem.h: "If there are no entries, *ppFileInfo is set to nullptr and
            // S_OK is returned"). The destination directory was already created above, so the
            // copy of this empty directory is complete -- treating it as an error would abort the
            // whole tree copy whenever an empty subdirectory is encountered.
            if (managedSourceRetentionGeneration.load(std::memory_order_acquire) != frame.retentionGenerationBefore)
            {
                return S_FALSE;
            }
            return FinalizeManagedSourceCleanup(sourcePath, destinationPath, frame.managedDirectoryCleanup);
        }

        unsigned long bufferSize = 0;
        hr                       = frame.info->GetBufferSize(&bufferSize);
        if (FAILED(hr) || bufferSize < sizeof(FileInfo))
        {
            return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        NoteTraversalMetadataRetained(bufferSize);
        frame.retainedMetadataBytes += bufferSize;

        std::byte* base = reinterpret_cast<std::byte*>(entry);
        std::byte* end  = base + bufferSize;
#ifdef ENABLE_TESTS
        hr = MaybeInjectHostileBridgeChildNamesForSelfTest(entry, base, end);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = MaybeInjectBridgeFileReparseForSelfTest(entry, base, end);
        if (FAILED(hr))
        {
            return hr;
        }
#endif

        for (;;)
        {
            std::wstring_view name;
            hr = TryGetValidatedFileInfoName(entry, base, end, name);
            if (FAILED(hr))
            {
                if (! continueOnError)
                {
                    return hr;
                }
                frame.hadChildFailure = true;
                break;
            }

            hr = ValidateAndRegisterChildName(name, frame.registeredChildNames);
            if (FAILED(hr))
            {
                NoteInvalidEnumeratedChildName(sourcePath, destinationPath);
                if (! continueOnError)
                {
                    return hr;
                }
                frame.hadChildFailure = true;
            }
            else
            {
                frame.childEntries.push_back(entry);
                frame.childNames.push_back(name);
            }

            FileInfo* nextEntry = nullptr;
            hr                  = AdvanceValidatedFileInfoEntry(entry, base, end, nextEntry);
            if (hr == S_FALSE)
            {
                break;
            }
            if (FAILED(hr))
            {
                if (! continueOnError)
                {
                    return hr;
                }
                frame.hadChildFailure = true;
                break;
            }
            entry = nextEntry;
        }

        const uint64_t traversalIndexBytes = static_cast<uint64_t>(frame.childEntries.capacity()) * sizeof(FileInfo*) +
                                             static_cast<uint64_t>(frame.childNames.capacity()) * sizeof(std::wstring_view);
        NoteTraversalMetadataRetained(traversalIndexBytes);
        frame.retainedMetadataBytes = AddRetainedBytes(frame.retainedMetadataBytes, traversalIndexBytes);

        return S_OK;
    }

    // Epilogue of one directory after every child was processed: the former post-loop block.
    HRESULT FinishSequentialDirectoryFrame(SequentialDirectoryFrame& frame) noexcept
    {
        const std::wstring& sourcePath      = frame.sourcePath;
        const std::wstring& destinationPath = frame.destinationPath;

        if (frame.hadChildFailure)
        {
            MarkManagedSourceRetained(
                sourcePath, destinationPath, HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), SourceKeptReason(L"at least one child did not complete."));
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        if (frame.hadRetainedChild || managedSourceRetentionGeneration.load(std::memory_order_acquire) != frame.retentionGenerationBefore)
        {
            // A retained descendant makes every containing source directory non-removable.
            // Propagate S_FALSE to the parent so it also releases its exact delete authority
            // without issuing a second, misleading directory-not-empty cleanup prompt.
            return S_FALSE;
        }

        RestoreSequentialDirectoryMetadata(frame);
        const bool boundCleanup = frame.managedDirectoryCleanup.armed;
        const HRESULT cleanupHr = boundCleanup ? FinalizeManagedSourceCleanup(sourcePath, destinationPath, frame.managedDirectoryCleanup)
                                               : (renameMerge ? RemoveEmptiedSourceDirectoryByName(sourcePath, destinationPath) : S_OK);
        if (renameMerge && frame.depth == 0u && cleanupHr == S_OK && ! managedSourceRetained.load(std::memory_order_acquire))
        {
            rootSourceRemoved.store(true, std::memory_order_release);
        }
        return cleanupHr;
    }

    // A child's result seen from its parent frame: the former per-child block. Returns S_OK when the
    // parent continues with its next child, otherwise the failure that ends the whole walk.
    HRESULT AcceptSequentialChildResult(SequentialDirectoryFrame& parent, HRESULT hr) noexcept
    {
        const bool continueOnError = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
        if (hr == S_FALSE)
        {
            parent.hadRetainedChild = true;
            return S_OK;
        }
        if (FAILED(hr))
        {
            const bool traversalLimit = hr == HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW) || hr == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
            if (traversalLimit || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT || ! continueOnError)
            {
                return hr;
            }
            parent.hadChildFailure = true;
        }
        return S_OK;
    }

    HRESULT CopyDirectorySequential(const std::wstring& sourcePath, std::wstring destinationPath, uint64_t depth = 0u) noexcept
    {
        std::vector<std::unique_ptr<SequentialDirectoryFrame>> frames;
        const auto pushFrame = [&](const std::wstring& source, std::wstring destination, uint64_t frameDepth) noexcept -> HRESULT
        {
            std::unique_ptr<SequentialDirectoryFrame> frame(new (std::nothrow) SequentialDirectoryFrame{});
            if (! frame)
            {
                return E_OUTOFMEMORY;
            }
            // The frame object is allocated without throwing above; the remaining copies can only
            // fail with std::bad_alloc, which is fatal by policy (AGENTS.md), so nothing is caught here.
            frame->sourcePath      = source;
            frame->destinationPath = std::move(destination);
            frame->depth           = frameDepth;
            frames.push_back(std::move(frame));
            return S_OK;
        };
        // Every remaining frame pops with its exit actions only; the recursive walk returned
        // through each level the same way.
        const auto unwind = [&]() noexcept
        {
            while (! frames.empty())
            {
                PopSequentialDirectoryFrame(frames);
            }
        };

        HRESULT hr = pushFrame(sourcePath, std::move(destinationPath), depth);
        if (FAILED(hr))
        {
            return hr;
        }

        HRESULT result = S_OK;
        while (! frames.empty())
        {
            SequentialDirectoryFrame& frame = *frames.back();
            HRESULT frameResult             = S_OK;
            bool frameDone                  = false;

            if (! frame.started)
            {
                frameResult = StartSequentialDirectoryFrame(frame);
                frameDone   = frameResult != S_OK;
            }

            if (! frameDone && frame.childIndex < frame.childEntries.size())
            {
                task.WaitWhilePaused();
                if (CancelRequested())
                {
                    unwind();
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                const size_t childIndex        = frame.childIndex++;
                const FileInfo* const entry    = frame.childEntries[childIndex];
                const std::wstring_view name   = frame.childNames[childIndex];
                const bool isDirectory         = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                const bool isReparse           = (entry->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                const std::wstring childSource = JoinFolderAndLeaf(frame.sourcePath, name);
                const std::wstring childDest   = JoinFolderAndLeaf(frame.destinationPath, name);

                if (isDirectory)
                {
                    ++discoveredDirectories;
                }
                else
                {
                    ++discoveredFiles;
                    if (entry->EndOfFile > 0)
                    {
                        discoveredBytes = AddRetainedBytes(discoveredBytes, static_cast<uint64_t>(entry->EndOfFile));
                    }
                }
                hr = ReportDiscovery(0u, false);
                if (FAILED(hr))
                {
                    unwind();
                    return hr;
                }

                HRESULT childHr        = S_OK;
                bool traverseChildTree = false;
                if (renameMerge)
                {
                    childHr = RelocateChild(childSource, childDest, isDirectory, isReparse, traverseChildTree);
                }
                else if (isReparse)
                {
                    childHr = CopyLink(childSource, childDest, isDirectory, frame.depth + 1u, &traverseChildTree);
                }
                else if (isDirectory)
                {
                    traverseChildTree = true;
                }
                else
                {
                    childHr = CopyFile(childSource, childDest);
                }

                if (traverseChildTree && SUCCEEDED(childHr))
                {
                    // The child directory becomes the top frame; its result reaches this frame when it pops.
                    hr = pushFrame(childSource, childDest, frame.depth + 1u);
                    if (FAILED(hr))
                    {
                        unwind();
                        return hr;
                    }
                    continue;
                }

                hr = AcceptSequentialChildResult(frame, childHr);
                if (FAILED(hr))
                {
                    unwind();
                    return hr;
                }
                continue;
            }

            if (! frameDone)
            {
                frameResult = FinishSequentialDirectoryFrame(frame);
            }

            // Pop this frame and deliver its result to the parent frame (or return it for the root).
            PopSequentialDirectoryFrame(frames);
            if (frames.empty())
            {
                result = frameResult;
                break;
            }
            hr = AcceptSequentialChildResult(*frames.back(), frameResult);
            if (FAILED(hr))
            {
                unwind();
                return hr;
            }
        }
        return result;
    }

    HRESULT CopyDirectoryParallel(const std::wstring& sourcePath, std::wstring destinationPath, unsigned int withinFolderBudget) noexcept
    {
        if (withinFolderBudget <= 1u)
        {
            return CopyDirectorySequential(sourcePath, destinationPath);
        }

        if (CancelRequested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (! connectionLimitsInitialized)
        {
            InitializeConnectionLimits(sourcePath, destinationPath);
        }

        const bool continueOnError = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;

        struct WorkItem final
        {
            std::wstring source;
            std::wstring destination;
            uint64_t retainedPathBytes = 0;
        };

        std::deque<WorkItem> workItems;
        std::mutex workMutex;
        std::condition_variable workCv;
        size_t activeWorkItems = 0;
        std::atomic<uint64_t> overallCompletedBytes(completedBytes);
        std::atomic<bool> producerDone{false};
        std::atomic<uint64_t> fileStartedBeforeProducerDone{0};
        std::atomic<bool> stopRequested{false};
        std::atomic<bool> hadWorkerFailure{false};
        std::atomic<HRESULT> firstFailure{S_OK};
        uint64_t directoryEnsureCount   = 0;
        uint64_t fileAdmissionCount     = 0;
        uint64_t maxAdmissionQueueDepth = 0;

        const auto workerProc = [&](size_t workerIndex) noexcept -> HRESULT
        {
            CrossFsBridgeBufferLease localBufferBudgetLease;
            const uint64_t reservationBytes = static_cast<uint64_t>(bufferBytes) * 2ull;
            const bool acquiredBudget       = localBufferBudgetLease.Acquire(reservationBytes, task._cancelled, task._stopToken);
            const HRESULT allocationHr =
                acquiredBudget ? S_OK
                               : ((task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested()) ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                                                                        : E_OUTOFMEMORY);
            std::unique_ptr<std::byte[]> localBuffer;
            if (acquiredBudget)
            {
                localBuffer.reset(new (std::nothrow) std::byte[bufferBytes]);
            }
            if (FAILED(allocationHr) || ! localBuffer)
            {
                const HRESULT failure = FAILED(allocationHr) ? allocationHr : E_OUTOFMEMORY;
                hadWorkerFailure.store(true, std::memory_order_release);
                HRESULT expected = S_OK;
                static_cast<void>(firstFailure.compare_exchange_strong(expected, failure));
                stopRequested.store(true, std::memory_order_release);
                workCv.notify_all();
                return failure;
            }

            const unsigned long localBufferBytes = bufferBytes;
            const uint64_t progressStreamId      = static_cast<uint64_t>(workerIndex);
            const auto recordFailure             = [&](HRESULT failure) noexcept
            {
                hadWorkerFailure.store(true, std::memory_order_release);
                const bool cancellation = failure == HRESULT_FROM_WIN32(ERROR_CANCELLED) || failure == E_ABORT;
                if (cancellation || ! continueOnError)
                {
                    HRESULT expected = S_OK;
                    static_cast<void>(firstFailure.compare_exchange_strong(expected, failure));
                    stopRequested.store(true, std::memory_order_release);
                    workCv.notify_all();
                }
            };
            bool starvationReported = false;

            for (;;)
            {
                task.WaitWhilePaused();

                if (CancelRequested())
                {
                    recordFailure(HRESULT_FROM_WIN32(ERROR_CANCELLED));
                    break;
                }

                WorkItem item{};
                {
                    std::unique_lock lock(workMutex);
                    while (! producerDone.load(std::memory_order_acquire) && ! stopRequested.load(std::memory_order_acquire) && ! CancelRequested())
                    {
                        const bool discoveryAhead = DiscoveryAheadEnabled();
                        const size_t activeLimit =
                            Common::FileOperations::DiscoveryWorkerLimit(static_cast<size_t>(withinFolderBudget), workItems.size(), discoveryAhead);
                        if (! workItems.empty() && activeWorkItems < activeLimit)
                        {
                            break;
                        }
                        if (workItems.empty() && ! starvationReported)
                        {
                            task._discoveryStarvationCount.fetch_add(1u, std::memory_order_acq_rel);
                            starvationReported = true;
                        }
                        workCv.wait_for(lock, std::chrono::milliseconds(50));
                    }

                    if (CancelRequested())
                    {
                        recordFailure(HRESULT_FROM_WIN32(ERROR_CANCELLED));
                        break;
                    }

                    if (workItems.empty())
                    {
                        if (producerDone.load(std::memory_order_acquire) || stopRequested.load(std::memory_order_acquire))
                        {
                            break;
                        }
                        continue;
                    }

                    if (stopRequested.load(std::memory_order_acquire))
                    {
                        break;
                    }

                    item = std::move(workItems.front());
                    workItems.pop_front();
                    ++activeWorkItems;
                    starvationReported = false;
                    workCv.notify_all();
                }

                const auto releaseWorkReservation = wil::scope_exit([&]() noexcept
                {
                    ReleaseTraversalWork(item.retainedPathBytes);
                    {
                        std::scoped_lock lock(workMutex);
                        if (activeWorkItems > 0u)
                        {
                            --activeWorkItems;
                        }
                    }
                    workCv.notify_all();
                });

                if (! producerDone.load(std::memory_order_acquire))
                {
                    fileStartedBeforeProducerDone.fetch_add(1u, std::memory_order_acq_rel);
                }
                const HRESULT hrItem =
                    PumpAndPublishFile(item.source, item.destination, localBuffer.get(), localBufferBytes, progressStreamId, overallCompletedBytes);
                if (FAILED(hrItem))
                {
                    recordFailure(hrItem);
                }
            }

            return S_OK;
        };

        auto& scheduler              = GetPerItemTaskScheduler();
        const auto schedulerStart    = scheduler.CapturePerfSnapshot();
        const uint64_t schedulerWall = PerfNowUs();

        ReferencedPerItemExecutionPolicy workerPolicy(workerProc);
        std::atomic<HRESULT> schedulerFirstFailure{S_OK};
        auto job = scheduler.StartJob(
            &task, withinFolderBudget, withinFolderBudget, workerPolicy, schedulerFirstFailure, PerItemTaskScheduler::FailurePolicy::RecordOnly);

        const auto recordProducerFailure = [&](HRESULT failure) noexcept
        {
            hadWorkerFailure.store(true, std::memory_order_release);
            const bool cancellation   = failure == HRESULT_FROM_WIN32(ERROR_CANCELLED) || failure == E_ABORT;
            const bool traversalLimit = failure == HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW) || failure == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
            if (cancellation || traversalLimit || ! continueOnError)
            {
                HRESULT expected = S_OK;
                static_cast<void>(firstFailure.compare_exchange_strong(expected, failure));
                stopRequested.store(true, std::memory_order_release);
            }
            workCv.notify_all();
        };

        const auto enqueueWork = [&](WorkItem item) noexcept -> HRESULT
        {
            using namespace std::chrono_literals;

            item.retainedPathBytes = RetainedPathBytes(item.source, item.destination);
            HRESULT reserveHr      = ReserveTraversalWork(item.retainedPathBytes, item.source, item.destination);
            if (FAILED(reserveHr))
            {
                return reserveHr;
            }
            bool reservationTransferred            = false;
            const auto releaseReservationOnFailure = wil::scope_exit([&]() noexcept
            {
                if (! reservationTransferred)
                {
                    ReleaseTraversalWork(item.retainedPathBytes);
                }
            });

            const auto queueTarget = [&]() noexcept
            { return Common::FileOperations::DiscoveryQueueTarget(static_cast<size_t>(withinFolderBudget), DiscoveryAheadEnabled()); };

            std::unique_lock lock(workMutex);
            while (workItems.size() >= queueTarget() && ! producerDone.load(std::memory_order_acquire) && ! stopRequested.load(std::memory_order_acquire) &&
                   ! CancelRequested())
            {
                workCv.wait_for(lock, 50ms);
            }

            if (stopRequested.load(std::memory_order_acquire) || CancelRequested())
            {
                const HRESULT recordedFailure = firstFailure.load(std::memory_order_acquire);
                if (FAILED(recordedFailure) && recordedFailure != HRESULT_FROM_WIN32(ERROR_CANCELLED) && recordedFailure != E_ABORT)
                {
                    return recordedFailure;
                }
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            workItems.push_back(std::move(item));
            reservationTransferred = true;
            ++fileAdmissionCount;
            maxAdmissionQueueDepth = (std::max)(maxAdmissionQueueDepth, static_cast<uint64_t>(workItems.size()));
            lock.unlock();
            workCv.notify_one();
            return S_OK;
        };

        const auto waitForWorkDrain = [&]() noexcept -> HRESULT
        {
            using namespace std::chrono_literals;
            std::unique_lock lock(workMutex);
            while ((! workItems.empty() || activeWorkItems != 0u) && ! stopRequested.load(std::memory_order_acquire) && ! CancelRequested())
            {
                workCv.wait_for(lock, 50ms);
            }
            if (CancelRequested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            const HRESULT failure = firstFailure.load(std::memory_order_acquire);
            return FAILED(failure) ? failure : S_OK;
        };

        const auto waitForActiveWorkQuiescence = [&]() noexcept
        {
            using namespace std::chrono_literals;
            std::unique_lock lock(workMutex);
            while (activeWorkItems != 0u)
            {
                workCv.wait_for(lock, 50ms);
            }
        };

        // R4-T1: one directory of the parallel producer, kept on an explicit stack instead of the
        // C++ stack. The fields mirror the locals of the former recursive lambda; a frame's exit
        // actions (restore the created directory's metadata after active work quiesces, release the
        // retained bytes) run when it pops, in the order the former scope guards ran.
        struct ProducerDirectoryFrame final
        {
            std::wstring sourcePath;
            std::wstring destinationPath;
            uint64_t depth                 = 0u;
            bool started                   = false; // metadata captured: the pop restores it
            uint64_t retainedMetadataBytes = 0u;
            bool metadataReserved          = false;
            bool createdDestination        = false;
            wil::com_ptr<IFileSystemBoundObject> createdDestinationAuthority;
            FileSystemBasicInformation directoryBasicInfo{};
            bool directoryMetadataCaptured = false;
            bool directoryMetadataRestored = false;
            wil::com_ptr<IFilesInformation> info; // owns the FileInfo buffer
            FileInfo* entry       = nullptr;
            std::byte* base       = nullptr;
            std::byte* end        = nullptr;
            bool entriesExhausted = false;
            BridgeChildNameSet registeredChildNames;
        };
        std::vector<std::unique_ptr<ProducerDirectoryFrame>> producerFrames;

        const auto restoreProducerDirectoryMetadata = [&](ProducerDirectoryFrame& frame) noexcept
        {
            if (! frame.directoryMetadataRestored)
            {
                // A terminal traversal/worker failure can abandon queued work, but any
                // already active child publication must quiesce before post-order metadata.
                waitForActiveWorkQuiescence();
                RestoreCreatedDirectoryMetadata(frame.sourcePath,
                                                frame.destinationPath,
                                                frame.directoryMetadataCaptured,
                                                frame.directoryBasicInfo,
                                                frame.createdDestinationAuthority.get());
                frame.directoryMetadataRestored = true;
            }
        };
        const auto popProducerFrame = [&]() noexcept
        {
            ProducerDirectoryFrame& frame = *producerFrames.back();
            if (frame.started)
            {
                restoreProducerDirectoryMetadata(frame);
            }
            if (frame.metadataReserved)
            {
                ReleaseTraversalMetadata(frame.retainedMetadataBytes);
            }
            producerFrames.pop_back();
        };
        const auto pushProducerFrame = [&](const std::wstring& source, std::wstring destination, uint64_t depth) noexcept -> HRESULT
        {
            std::unique_ptr<ProducerDirectoryFrame> frame(new (std::nothrow) ProducerDirectoryFrame{});
            if (! frame)
            {
                return E_OUTOFMEMORY;
            }
            // Same rule as the sequential walker: only std::bad_alloc can follow the nothrow
            // allocation, and std::bad_alloc is fatal by policy (AGENTS.md).
            frame->sourcePath      = source;
            frame->destinationPath = std::move(destination);
            frame->depth           = depth;
            producerFrames.push_back(std::move(frame));
            return S_OK;
        };

        // Prologue of one directory (the former lambda up to its entry loop). S_OK: entries follow;
        // any other value is the frame's result.
        const auto startProducerFrame = [&](ProducerDirectoryFrame& frame) noexcept -> HRESULT
        {
            const std::wstring& currentSource = frame.sourcePath;
            std::wstring& currentDest         = frame.destinationPath; // EnsureDestinationDirectory may rename it (keep both)

            task.WaitWhilePaused();
            if (CancelRequested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            HRESULT hr = ValidateTraversalDepth(frame.depth, currentSource, currentDest);
            if (FAILED(hr))
            {
                return hr;
            }

            frame.retainedMetadataBytes = AddRetainedBytes(RetainedPathBytes(currentSource, currentDest), sizeof(FileSystemBasicInformation));
            NoteTraversalMetadataRetained(frame.retainedMetadataBytes);
            frame.metadataReserved     = true;
            frame.registeredChildNames = MakeChildNameSet();

            hr = EnsureDestinationDirectory(currentSource, currentDest, &frame.createdDestination, std::addressof(frame.createdDestinationAuthority));
            if (hr == S_FALSE || FAILED(hr))
            {
                return hr;
            }
            ++directoryEnsureCount;

            frame.directoryMetadataCaptured = CaptureCreatedDirectoryMetadata(currentSource, currentDest, frame.createdDestination, frame.directoryBasicInfo);
            frame.started                   = true;

            task._bridgeSourceDirectoryEnumerationCount.fetch_add(1u, std::memory_order_relaxed);
            hr = sourceFs.ReadDirectoryInfo(currentSource.c_str(), frame.info.addressof());
            if (FAILED(hr))
            {
                return hr;
            }

            hr = frame.info->GetBuffer(&frame.entry);
            if (FAILED(hr))
            {
                return hr;
            }
            if (frame.entry == nullptr)
            {
                frame.entriesExhausted = true;
                return waitForWorkDrain();
            }

            unsigned long bufferSize = 0;
            hr                       = frame.info->GetBufferSize(&bufferSize);
            if (FAILED(hr) || bufferSize < sizeof(FileInfo))
            {
                return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            NoteTraversalMetadataRetained(bufferSize);
            frame.retainedMetadataBytes += bufferSize;

            frame.base = reinterpret_cast<std::byte*>(frame.entry);
            frame.end  = frame.base + bufferSize;
#ifdef ENABLE_TESTS
            hr = MaybeInjectHostileBridgeChildNamesForSelfTest(frame.entry, frame.base, frame.end);
            if (FAILED(hr))
            {
                return hr;
            }
            hr = MaybeInjectBridgeFileReparseForSelfTest(frame.entry, frame.base, frame.end);
            if (FAILED(hr))
            {
                return hr;
            }
#endif
            return S_OK;
        };

        // One entry of the top frame (the former loop body). Returns S_OK to continue with the next
        // entry, S_FALSE when a child directory frame was pushed, otherwise the frame's result.
        const auto produceNextEntry = [&](ProducerDirectoryFrame& frame) noexcept -> HRESULT
        {
            const std::wstring& currentSource = frame.sourcePath;
            const std::wstring& currentDest   = frame.destinationPath;
            FileInfo* const entry             = frame.entry;

            task.WaitWhilePaused();
            if (CancelRequested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            if (stopRequested.load(std::memory_order_acquire))
            {
                // A child directory that ended with a stop request ends this frame too, as the
                // former recursion returned right after the child call.
                const HRESULT failure = firstFailure.load(std::memory_order_acquire);
                return FAILED(failure) ? failure : HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            std::wstring_view name;
            HRESULT hr = TryGetValidatedFileInfoName(entry, frame.base, frame.end, name);
            if (FAILED(hr))
            {
                return hr;
            }

            bool pushedChild = false;
            hr               = ValidateAndRegisterChildName(name, frame.registeredChildNames);
            if (FAILED(hr))
            {
                NoteInvalidEnumeratedChildName(currentSource, currentDest);
                recordProducerFailure(hr);
                if (stopRequested.load(std::memory_order_acquire))
                {
                    return hr;
                }
            }
            else
            {
                const bool isDirectory   = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                const bool isReparse     = (entry->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                std::wstring childSource = JoinFolderAndLeaf(currentSource, name);
                std::wstring childDest   = JoinFolderAndLeaf(currentDest, name);

                if (isDirectory)
                {
                    ++discoveredDirectories;
                }
                else
                {
                    ++discoveredFiles;
                    if (entry->EndOfFile > 0)
                    {
                        discoveredBytes = AddRetainedBytes(discoveredBytes, static_cast<uint64_t>(entry->EndOfFile));
                    }
                }
                uint32_t discoveryQueueDepth = 0u;
                {
                    std::scoped_lock lock(workMutex);
                    discoveryQueueDepth =
                        static_cast<uint32_t>((std::min)(workItems.size() + activeWorkItems, static_cast<size_t>(std::numeric_limits<uint32_t>::max())));
                }
                hr = ReportDiscovery(discoveryQueueDepth, false);
                if (FAILED(hr))
                {
                    return hr;
                }

                bool traverseChildTree = false;
                bool enqueueChildFile  = false;
                if (isReparse)
                {
                    const HRESULT linkHr = CopyLink(childSource, childDest, isDirectory, frame.depth + 1u, &traverseChildTree, &enqueueChildFile);
                    if (FAILED(linkHr))
                    {
                        recordProducerFailure(linkHr);
                    }
                }
                else if (isDirectory)
                {
                    traverseChildTree = true;
                }
                else
                {
                    enqueueChildFile = true;
                }
                if (enqueueChildFile)
                {
                    WorkItem item{};
                    item.source             = std::move(childSource);
                    item.destination        = std::move(childDest);
                    const HRESULT enqueueHr = enqueueWork(std::move(item));
                    if (FAILED(enqueueHr))
                    {
                        recordProducerFailure(enqueueHr);
                    }
                }

                if (traverseChildTree)
                {
                    const HRESULT pushHr = pushProducerFrame(childSource, childDest, frame.depth + 1u);
                    if (FAILED(pushHr))
                    {
                        recordProducerFailure(pushHr);
                    }
                    else
                    {
                        pushedChild = true;
                    }
                }
            }

            if (stopRequested.load(std::memory_order_acquire))
            {
                const HRESULT failure = firstFailure.load(std::memory_order_acquire);
                return FAILED(failure) ? failure : HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            FileInfo* nextEntry = nullptr;
            hr                  = AdvanceValidatedFileInfoEntry(entry, frame.base, frame.end, nextEntry);
            if (hr == S_FALSE)
            {
                frame.entriesExhausted = true;
            }
            else if (FAILED(hr))
            {
                return hr;
            }
            else
            {
                frame.entry = nextEntry;
            }
            return pushedChild ? S_FALSE : S_OK;
        };

        // Epilogue of one directory after its last entry (the former lambda tail).
        const auto finishProducerFrame = [&](ProducerDirectoryFrame& frame) noexcept -> HRESULT
        {
#ifdef ENABLE_TESTS
            const unsigned int producerDelayMs = GetBridgeProducerDelayMsForSelfTest();
            if (producerDelayMs > 0)
            {
                SleepResponsive(producerDelayMs);
            }
#endif
            const HRESULT hr = waitForWorkDrain();
            if (FAILED(hr))
            {
                return hr;
            }
            restoreProducerDirectoryMetadata(frame);
            return S_OK;
        };

        HRESULT producerHr = pushProducerFrame(sourcePath, destinationPath, 0u);
        while (SUCCEEDED(producerHr) && ! producerFrames.empty())
        {
            ProducerDirectoryFrame& frame = *producerFrames.back();
            HRESULT frameResult           = S_OK;
            bool frameDone                = false;
            if (! frame.started)
            {
                frameResult = startProducerFrame(frame);
                frameDone   = frameResult != S_OK;
            }
            if (! frameDone && ! frame.entriesExhausted)
            {
                const HRESULT entryHr = produceNextEntry(frame);
                if (entryHr == S_FALSE)
                {
                    continue; // a child directory frame is now on top
                }
                if (entryHr != S_OK)
                {
                    frameResult = entryHr;
                    frameDone   = true;
                }
                else
                {
                    continue;
                }
            }
            if (! frameDone)
            {
                frameResult = finishProducerFrame(frame);
            }
            // The former recursion recorded a failed child directory and continued with the parent's
            // next entry; the root frame's result is the producer's result.
            const bool rootFrame = producerFrames.size() == 1u;
            popProducerFrame();
            if (rootFrame)
            {
                producerHr = frameResult;
            }
            else if (FAILED(frameResult))
            {
                recordProducerFailure(frameResult);
            }
        }
        while (! producerFrames.empty())
        {
            popProducerFrame(); // a failed push leaves frames behind; unwind with their exit actions
        }
        if (FAILED(producerHr))
        {
            recordProducerFailure(producerHr);
        }

        producerDone.store(true, std::memory_order_release);
        workCv.notify_all();
        uint32_t queuedAtTraversalClose = 0u;
        {
            std::scoped_lock lock(workMutex);
            queuedAtTraversalClose =
                static_cast<uint32_t>((std::min)(workItems.size() + activeWorkItems, static_cast<size_t>(std::numeric_limits<uint32_t>::max())));
        }
        const HRESULT discoveryCloseHr = ReportDiscovery(queuedAtTraversalClose, true);
        if (FAILED(discoveryCloseHr))
        {
            recordProducerFailure(discoveryCloseHr);
        }
        scheduler.WaitJob(job);

        const auto schedulerEnd = scheduler.CapturePerfSnapshot();
        task._bridgeDirectoryEnsureCount.fetch_add(directoryEnsureCount, std::memory_order_acq_rel);
        task._bridgeFileAdmissionCount.fetch_add(fileAdmissionCount, std::memory_order_acq_rel);
        task._bridgeFileStartedBeforeProducerDone.fetch_add(fileStartedBeforeProducerDone.load(std::memory_order_acquire), std::memory_order_acq_rel);
        AtomicMax(task._bridgeAdmissionMaxQueueDepth, maxAdmissionQueueDepth);
        task._perf.schedulerWaitUs.fetch_add(PerfElapsedUs(schedulerWall), std::memory_order_relaxed);
        task._perf.schedulerDequeueAttempts.fetch_add(
            (schedulerEnd.dequeueAttempts >= schedulerStart.dequeueAttempts) ? (schedulerEnd.dequeueAttempts - schedulerStart.dequeueAttempts) : 0,
            std::memory_order_relaxed);
        task._perf.schedulerDequeueSuccess.fetch_add(
            (schedulerEnd.dequeueSuccess >= schedulerStart.dequeueSuccess) ? (schedulerEnd.dequeueSuccess - schedulerStart.dequeueSuccess) : 0,
            std::memory_order_relaxed);
        task._perf.schedulerWaitForWorkUs.fetch_add(
            (schedulerEnd.waitForWorkUs >= schedulerStart.waitForWorkUs) ? (schedulerEnd.waitForWorkUs - schedulerStart.waitForWorkUs) : 0,
            std::memory_order_relaxed);
        task._perf.schedulerProcessIndexUs.fetch_add(
            (schedulerEnd.processIndexUs >= schedulerStart.processIndexUs) ? (schedulerEnd.processIndexUs - schedulerStart.processIndexUs) : 0,
            std::memory_order_relaxed);

        completedBytes = overallCompletedBytes.load(std::memory_order_acquire);

        const HRESULT failure = firstFailure.load(std::memory_order_acquire);
        if (CancelRequested() || failure == HRESULT_FROM_WIN32(ERROR_CANCELLED) || failure == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (FAILED(failure))
        {
            return failure;
        }

        if (continueOnError && hadWorkerFailure.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    HRESULT CopyDirectory(const std::wstring& sourcePath, std::wstring destinationPath) noexcept
    {
        const unsigned int withinFolderBudget = ComputeWithinFolderBudget();
        if (sourceCleanupPermitted || withinFolderBudget <= 1u)
        {
            return CopyDirectorySequential(sourcePath, destinationPath);
        }

        // Parallel workers reserve their own two-buffer pipeline allotments. Release the
        // otherwise-idle primary pair first so many simultaneous directory producers
        // cannot consume the whole global budget and then wait on their nested workers.
        buffer.reset();
        bufferBudgetLease.Reset();
        const HRESULT directoryHr = CopyDirectoryParallel(sourcePath, destinationPath, withinFolderBudget);
        if (FAILED(directoryHr) || task._operation != FILESYSTEM_MOVE)
        {
            return directoryHr;
        }

        if (! bufferBudgetLease.Acquire(bufferBytes, task._cancelled, task._stopToken))
        {
            bufferAllocationHr = CancelRequested() ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : E_OUTOFMEMORY;
            return bufferAllocationHr;
        }
        buffer.reset(new (std::nothrow) std::byte[bufferBytes]);
        if (! buffer)
        {
            bufferBudgetLease.Reset();
            bufferAllocationHr = E_OUTOFMEMORY;
            return bufferAllocationHr;
        }
        return directoryHr;
    }

    // Rename-merge entry: the selected root is a regular directory whose destination already holds
    // a regular directory. The sequential walk relocates every child through the provider's Native
    // rename and removes each emptied source directory through the exact cleanup boundary.
    HRESULT RenameMergeDirectory(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
    {
        rootVerificationNotApplicable = true;
        ++discoveredDirectories;
        const auto closeDiscoveryOnExit = wil::scope_exit([&]() noexcept { static_cast<void>(ReportDiscovery(0u, true)); });
        const HRESULT discoveryHr       = ReportDiscovery(0u, false);
        if (FAILED(discoveryHr))
        {
            return discoveryHr;
        }
        const HRESULT dirHr = CopyDirectorySequential(sourcePath, destinationPath);
        if (FAILED(dirHr))
        {
            return dirHr;
        }
        if (dirHr == S_FALSE || skippedFileConflictCount.load(std::memory_order_acquire) > 0)
        {
            // A child stayed behind at its conflict prompt; the tree is not fully relocated.
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
        return S_OK;
    }

    // Traversal seam: classifies the root without following links, then owns discovery,
    // directory walking, deferred-link draining, and traversal-limit completion.
    HRESULT TraverseAndPublishPath(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
    {
        unsigned long attributes = sourceRootAttributesHint;
        const bool haveHint      = attributes != 0;

        if (! connectionLimitsInitialized)
        {
            InitializeConnectionLimits(sourcePath, destinationPath);
        }

        // Hints can be stale, especially for recently-created junctions. Link handling must
        // classify the source object itself, so refresh attributes without following targets
        // and use the admission hint only when the provider cannot refresh it.
        {
            unsigned long refreshed = 0;
            const HRESULT hrAttr    = sourceIo.GetAttributes(sourcePath.c_str(), &refreshed);
            if (SUCCEEDED(hrAttr))
            {
                attributes = refreshed;
            }
            else if (! haveHint)
            {
                return hrAttr;
            }
        }

        const bool isDirectory        = (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const bool isReparse          = (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
        rootVerificationNotApplicable = isDirectory || isReparse;
        if (isDirectory)
        {
            ++discoveredDirectories;
        }
        else
        {
            ++discoveredFiles;
        }
        const auto closeDiscoveryOnExit = wil::scope_exit([&]() noexcept { static_cast<void>(ReportDiscovery(0u, true)); });
        HRESULT discoveryHr             = ReportDiscovery(0u, false);
        if (FAILED(discoveryHr))
        {
            return discoveryHr;
        }
        if (isReparse)
        {
            const HRESULT linkHr = CopyLink(sourcePath, destinationPath, isDirectory);
            if (FAILED(linkHr))
            {
                return linkHr;
            }
            return linkHr == S_FALSE ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
        }
        if (isDirectory)
        {
            const HRESULT dirHr = CopyDirectory(sourcePath, destinationPath);
            if (FAILED(dirHr))
            {
                return dirHr;
            }
            if (skippedDirectoryReparseCount > 0 || skippedFileReparseCount > 0 || skippedFileConflictCount.load(std::memory_order_acquire) > 0)
            {
                // Some child files were skipped at a conflict prompt; the tree is not a
                // full copy. Caller treats PARTIAL as "source preserved" for MOVE.
                return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }
            return dirHr;
        }

        const HRESULT fileHr = CopyFile(sourcePath, destinationPath);
        if (fileHr == S_FALSE)
        {
            // Single top-level file skipped at its conflict prompt.
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
        return fileHr;
    }
};
} // namespace FolderWindowFileOperationsStateInternal

HRESULT FolderWindow::FileOperationState::Task::ExecuteOperation() noexcept
{
    if (! _fileSystem)
    {
        return E_POINTER;
    }

    if (_sourcePaths.empty())
    {
        return S_FALSE;
    }

    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = LoadPlans();

    WaitWhilePaused();
    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    const HRESULT artifactTouchHr = RevalidateArtifactTouchGuard();
    if (FAILED(artifactTouchHr))
    {
        LogDiagnostic(DiagnosticSeverity::Error,
                      artifactTouchHr,
                      L"artifact.touch.revalidateFailed",
                      L"An accepted operation-artifact identity changed or could not be revalidated; no mutation was attempted.");
        return artifactTouchHr;
    }

    _observedSkipAction.store(false, std::memory_order_release);
    _started.store(true, std::memory_order_release);
    _operationStartTick.store(GetTickCount64(), std::memory_order_release);

#ifdef ENABLE_TESTS
    _dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
    const auto dbgCallbackScope = wil::scope_exit([&] noexcept { _dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
#endif

#ifdef ENABLE_TESTS
    _dbgConfiguredMaxConcurrency =
        DeterminePerItemMaxConcurrency(_fileSystem, _sourcePaths, _sourcePluginId, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles));
    _dbgConfiguredMaxConcurrency         = std::max(1u, _dbgConfiguredMaxConcurrency);
    _dbgSingleInFlightStartTick          = 0;
    _dbgLastSingleInFlightWarnTick       = 0;
    _dbgObservedMultipleInFlightFiles    = false;
    _dbgLastPerItemInFlightEvictWarnTick = 0;
#endif

    std::filesystem::path destinationFolder;
    {
        std::scoped_lock lock(_operationMutex);
        destinationFolder = _destinationFolder;
    }

    const bool continueOnError  = (_flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    const bool useResolvedItems = ! _resolvedItems.empty();
    if (useResolvedItems && _resolvedItems.size() != _sourcePaths.size())
    {
        return E_INVALIDARG;
    }

    if (_operation == FILESYSTEM_RENAME)
    {
        if (! plans || plans->size() != 1u)
        {
            return E_UNEXPECTED;
        }
        const auto* rename = std::get_if<FileOperations::RenamePlan>(&plans->front());
        if (rename == nullptr)
        {
            return E_UNEXPECTED;
        }
        return rename->origin == FileOperations::RenameOrigin::InlineRename ? ExecuteInlineRename() : ExecuteBatchRename();
    }

    std::vector<const FileOperations::TransferPlan*> transferPlanBySourceIndex;
    std::vector<size_t> transferPlanItemIndexBySourceIndex;
    std::vector<const FileOperations::DeletePlan*> deletePlanBySourceIndex;
    if (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE)
    {
        if (! plans)
        {
            return E_UNEXPECTED;
        }
        transferPlanBySourceIndex.reserve(_sourcePaths.size());
        transferPlanItemIndexBySourceIndex.reserve(_sourcePaths.size());
        for (const FileOperations::FileOperationPlan& plan : *plans)
        {
            const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
            if (transfer == nullptr)
            {
                return E_UNEXPECTED;
            }
            for (size_t itemIndex = 0u; itemIndex < transfer->selectedItems.size(); ++itemIndex)
            {
                transferPlanBySourceIndex.emplace_back(transfer);
                transferPlanItemIndexBySourceIndex.emplace_back(itemIndex);
            }
        }
        if (transferPlanBySourceIndex.size() != _sourcePaths.size() || transferPlanItemIndexBySourceIndex.size() != _sourcePaths.size())
        {
            return E_UNEXPECTED;
        }

        if (Debug::Perf::IsCaptureEnabled())
        {
            struct StrategyMetricBucket final
            {
                std::wstring_view detail;
                uint64_t planCount = 0u;
                uint64_t itemCount = 0u;
            };
            std::array<StrategyMetricBucket, 8u> buckets{
                StrategyMetricBucket{.detail = L"copy.copy.same-root"},
                StrategyMetricBucket{.detail = L"copy.copy.cross-root"},
                StrategyMetricBucket{.detail = L"move.native.same-root"},
                StrategyMetricBucket{.detail = L"move.native.cross-root"},
                StrategyMetricBucket{.detail = L"move.managed.same-root"},
                StrategyMetricBucket{.detail = L"move.managed.cross-root"},
                StrategyMetricBucket{.detail = L"move.copy-only.same-root"},
                StrategyMetricBucket{.detail = L"move.copy-only.cross-root"},
            };
            for (const FileOperations::FileOperationPlan& plan : *plans)
            {
                const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
                if (transfer == nullptr)
                {
                    return E_UNEXPECTED;
                }

                const size_t topologyOffset =
                    FileOperations::QualifiedEndpointsReferToSameRoot(transfer->sourceEndpoint, transfer->destinationEndpoint) ? 0u : 1u;
                size_t strategyOffset = 0u;
                if (transfer->intent == FileOperations::TransferIntent::Copy)
                {
                    if (transfer->strategy != FileOperations::OperationStrategy::Copy)
                    {
                        return E_UNEXPECTED;
                    }
                }
                else
                {
                    switch (transfer->strategy)
                    {
                        case FileOperations::OperationStrategy::Native: strategyOffset = 2u; break;
                        case FileOperations::OperationStrategy::Managed: strategyOffset = 4u; break;
                        case FileOperations::OperationStrategy::CopyOnly: strategyOffset = 6u; break;
                        case FileOperations::OperationStrategy::Copy: return E_UNEXPECTED;
                    }
                }
                const size_t bucketIndex = strategyOffset + topologyOffset;
                if (bucketIndex >= buckets.size())
                {
                    return E_UNEXPECTED;
                }
                ++buckets[bucketIndex].planCount;
                buckets[bucketIndex].itemCount += static_cast<uint64_t>(transfer->selectedItems.size());
            }
            for (const StrategyMetricBucket& bucket : buckets)
            {
                if (bucket.planCount != 0u)
                {
                    Debug::Perf::Emit(L"fileops.operation.strategy", bucket.detail, 0u, bucket.planCount, bucket.itemCount, S_OK);
                }
            }
        }
    }

    if (_operation == FILESYSTEM_DELETE)
    {
        if (! plans)
        {
            return E_UNEXPECTED;
        }
        deletePlanBySourceIndex.reserve(_sourcePaths.size());
        for (const FileOperations::FileOperationPlan& plan : *plans)
        {
            const auto* deletion = std::get_if<FileOperations::DeletePlan>(&plan);
            if (deletion == nullptr)
            {
                return E_UNEXPECTED;
            }
            for ([[maybe_unused]] const FileOperations::QualifiedSourceItem& item : deletion->selectedItems)
            {
                deletePlanBySourceIndex.emplace_back(deletion);
            }
        }
        if (deletePlanBySourceIndex.size() != _sourcePaths.size())
        {
            return E_UNEXPECTED;
        }
    }

    const bool isTransferOperation = _operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE;
    if (isTransferOperation && _executionMode != ExecutionMode::PerItem)
    {
        LogDiagnostic(DiagnosticSeverity::Error,
                      E_UNEXPECTED,
                      L"strategy.bulkTransferRejected",
                      L"Copy and Move require the identity-guarded per-item executor; provider bulk mutation was rejected before I/O.");
        return E_UNEXPECTED;
    }
    const bool hasManagedMovePlans  = _operation == FILESYSTEM_MOVE && std::ranges::any_of(transferPlanBySourceIndex,
                                                                                           [](const FileOperations::TransferPlan* plan) noexcept
    { return plan != nullptr && plan->strategy == FileOperations::OperationStrategy::Managed; });
    const bool hasVerificationPlans = isTransferOperation && std::ranges::any_of(transferPlanBySourceIndex,
                                                                                 [](const FileOperations::TransferPlan* plan) noexcept
    { return plan != nullptr && plan->options.verifyAfterCopy && plan->strategy != FileOperations::OperationStrategy::Native; });
    if (_operation == FILESYSTEM_DELETE && _executionMode != ExecutionMode::PerItem &&
        std::ranges::any_of(deletePlanBySourceIndex, [](const FileOperations::DeletePlan* plan) noexcept {
        return plan != nullptr && plan->mode == FileOperations::DeleteMode::Permanent;
    }))
    {
        LogDiagnostic(DiagnosticSeverity::Error,
                      E_UNEXPECTED,
                      L"delete.bulkPermanentRejected",
                      L"Permanent Delete requires the exact-authority per-item executor; provider bulk mutation was rejected before I/O.");
        return E_UNEXPECTED;
    }

    const auto storeSourceMutationTruth = [&](size_t sourceIndex, HRESULT status, const std::optional<FileSystemItemMutationResult>& mutationResult) noexcept
    {
        std::scoped_lock lock(_sourceItemStatusMutex);
        if (_sourceItemResultBuilders.size() < _sourcePaths.size())
        {
            _sourceItemResultBuilders.resize(_sourcePaths.size());
        }
        if (sourceIndex < _sourcePaths.size())
        {
            SourceItemResultBuilder& builder = _sourceItemResultBuilders[sourceIndex];
            builder.status                   = status;
            if (mutationResult.has_value())
            {
                builder.mutation = mutationResult.value();
            }
        }
    };

    const auto storeQualifiedItemResult = [&](size_t sourceIndex,
                                              FileOperations::OperationStrategy strategy,
                                              std::wstring_view destinationPath,
                                              HRESULT status,
                                              bool succeeded,
                                              bool skipped,
                                              bool partiallySkipped,
                                              bool partiallyFailed,
                                              std::optional<FileOperations::PublicationState> partialPublication,
                                              bool managedSourceRetained,
                                              bool cleanupIndeterminate,
                                              FileOperations::OwnedStageDisposition ownedStageDisposition,
                                              FileOperations::VerificationState verification,
                                              const FileOperations::ProviderIdentitySnapshot* retainedSourceIdentity,
                                              bool selectedRootVerificationNotApplicable = false) noexcept
    {
        FileOperations::FileOperationItemResult result{};
        result.sourceIndex           = sourceIndex;
        result.strategy              = strategy;
        result.verification          = verification;
        result.ownedStageDisposition = ownedStageDisposition;
        result.status                = status;
        if (sourceIndex < _sourcePaths.size())
        {
            result.finalSourcePath = _sourcePaths[sourceIndex].native();
        }
        result.finalDestinationPath = destinationPath;

        std::optional<FileSystemItemMutationResult> mutation;
        {
            std::scoped_lock lock(_sourceItemStatusMutex);
            if (sourceIndex < _sourceItemResultBuilders.size())
            {
                mutation = _sourceItemResultBuilders[sourceIndex].mutation;
            }
        }
        const bool validMutation = mutation.has_value() && IsValidItemMutationResultPrefix(mutation.value());
        if (validMutation)
        {
            result.ownedStageDisposition = GetOwnedStageDisposition(mutation.value());
        }

        if (cleanupIndeterminate)
        {
            result.publication       = partialPublication.value_or(FileOperations::PublicationState::Unknown);
            result.sourceDisposition = result.publication == FileOperations::PublicationState::Published && _operation == FILESYSTEM_MOVE
                                           ? FileOperations::SourceDisposition::Unknown
                                           : FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Indeterminate;
            result.status            = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        else if (validMutation && mutation->outcomeKnown == FALSE)
        {
            result.publication = _operation == FILESYSTEM_DELETE ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Unknown;
            result.sourceDisposition = _operation == FILESYSTEM_COPY ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Unknown;
            result.completion        = FileOperations::ItemCompletion::Indeterminate;
            result.status            = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        else if (validMutation && mutation->mutationCommitted != FALSE)
        {
            result.publication = _operation == FILESYSTEM_DELETE ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Published;
            result.sourceDisposition =
                mutation->originalStillPresent != FALSE ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Removed;
            if (skipped || partiallySkipped)
            {
                result.completion = FileOperations::ItemCompletion::Skipped;
            }
            else if (IsCancellationOutcome(status))
            {
                result.status     = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                result.completion = FileOperations::ItemCompletion::Canceled;
            }
            else
            {
                result.completion = FAILED(status) ? FileOperations::ItemCompletion::Failed : FileOperations::ItemCompletion::Completed;
            }
        }
        else if (validMutation)
        {
            // A provider-known non-commit is exact retained-source truth, not an indeterminate
            // destructive result. Keep an explicit user Skip neutral; preserve every other failure
            // (including a recursive partial Delete) as Failed so the aggregate cannot become S_FALSE.
            result.publication =
                _operation == FILESYSTEM_DELETE ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::NotPublished;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            if (skipped || partiallySkipped)
            {
                result.completion = FileOperations::ItemCompletion::Skipped;
                result.status     = S_FALSE;
            }
            else if (IsCancellationOutcome(status))
            {
                result.status     = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                result.completion = FileOperations::ItemCompletion::Canceled;
            }
            else
            {
                result.completion = FileOperations::ItemCompletion::Failed;
            }
        }
        else if (partiallyFailed)
        {
            // Continue-on-error and recursive bridge failures are unexpected failures, never user
            // skips. The bridge supplies exact publication truth for its owned stages; providers
            // without that proof remain Unknown rather than being guessed from the HRESULT.
            result.publication       = partialPublication.value_or(FileOperations::PublicationState::Unknown);
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Failed;
            result.status            = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
        else if (succeeded && ! partiallySkipped &&
                 (_operation == FILESYSTEM_DELETE || (_operation == FILESYSTEM_MOVE && strategy == FileOperations::OperationStrategy::Native)))
        {
            // A successful destructive provider call without a valid receipt proves neither
            // which object was consumed nor whether the accepted source still exists. An explicit
            // user Skip answered at the provider's conflict prompt (ERROR_PARTIAL_COPY with the skip
            // observed) is a user decision, not an unproven mutation: it stays Skipped below.
            result.publication = _operation == FILESYSTEM_DELETE ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Unknown;
            result.sourceDisposition = FileOperations::SourceDisposition::Unknown;
            result.completion        = FileOperations::ItemCompletion::Indeterminate;
            result.status            = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        else if (succeeded || partiallySkipped || managedSourceRetained)
        {
            result.publication = isTransferOperation ? FileOperations::PublicationState::Published : FileOperations::PublicationState::NotAttempted;
            if (_operation == FILESYSTEM_COPY || strategy == FileOperations::OperationStrategy::CopyOnly || managedSourceRetained || partiallySkipped)
            {
                result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            }
            else if (_operation == FILESYSTEM_MOVE && strategy == FileOperations::OperationStrategy::Managed && ! partiallySkipped)
            {
                result.sourceDisposition = FileOperations::SourceDisposition::Removed;
            }
            else if (_operation == FILESYSTEM_DELETE)
            {
                result.sourceDisposition = FileOperations::SourceDisposition::Removed;
            }
            else
            {
                result.sourceDisposition = FileOperations::SourceDisposition::Unknown;
            }
            result.completion = partiallySkipped ? FileOperations::ItemCompletion::Skipped : FileOperations::ItemCompletion::Completed;
        }
        else if (skipped)
        {
            // A user decision describes completion, not mutation truth. The provider may have
            // already published (or lost proof of publication) before the prompt was answered.
            result.publication       = partialPublication.value_or(FileOperations::PublicationState::NotPublished);
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Skipped;
            result.status            = S_FALSE;
        }
        else if (IsCancellationOutcome(status))
        {
            // Cancellation can follow publication (for example, metadata-loss consent). Preserve
            // the bridge's exact publication truth so refresh and retry behavior match the disk.
            result.status            = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            result.publication       = partialPublication.value_or(FileOperations::PublicationState::NotAttempted);
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = FileOperations::ItemCompletion::Canceled;
        }
        else
        {
            result.publication       = isTransferOperation ? FileOperations::PublicationState::Unknown : FileOperations::PublicationState::NotAttempted;
            result.sourceDisposition = _operation == FILESYSTEM_COPY ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Unknown;
            result.completion =
                status == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) ? FileOperations::ItemCompletion::Indeterminate : FileOperations::ItemCompletion::Failed;
        }
        if (validMutation && OwnedStageDispositionIsIndeterminate(result.ownedStageDisposition))
        {
            // A provider can know that the final name was not published while independently
            // losing proof that its private stage was removed. Preserve both truths.
            result.completion = FileOperations::ItemCompletion::Indeterminate;
            result.status     = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (verification == FileOperations::VerificationState::Failed || verification == FileOperations::VerificationState::Canceled ||
            verification == FileOperations::VerificationState::Unavailable)
        {
            result.publication       = FileOperations::PublicationState::Published;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            result.completion        = verification == FileOperations::VerificationState::Canceled
                                           ? FileOperations::ItemCompletion::Canceled
                                           : (verification == FileOperations::VerificationState::Failed ? FileOperations::ItemCompletion::Failed
                                                                                                        : FileOperations::ItemCompletion::Completed);
        }
        if (result.sourceDisposition == FileOperations::SourceDisposition::Retained && retainedSourceIdentity != nullptr &&
            ! retainedSourceIdentity->objectId.empty() && ! retainedSourceIdentity->pathProfileId.empty())
        {
            result.retainedSourceIdentity = *retainedSourceIdentity;
        }
        if (selectedRootVerificationNotApplicable)
        {
            // Verification controls the item outcome from nested file proof, but the selected
            // directory itself does not have a file-content verification axis.
            result.verification = FileOperations::VerificationState::NotApplicable;
        }
        static_cast<void>(StoreTypedItemResult(std::move(result)));
    };

    const auto executeRecycleEscalation =
        [&](size_t sourceIndex, PerItemCallbackCookie* perItemCookie, HRESULT recycleStatus, std::wstring_view sourcePath) noexcept -> HRESULT
    {
        if (_operation != FILESYSTEM_DELETE || sourceIndex >= deletePlanBySourceIndex.size() || sourcePath.empty())
        {
            return E_UNEXPECTED;
        }

        std::optional<FileSystemItemMutationResult> recycleMutation;
        {
            std::scoped_lock lock(_sourceItemStatusMutex);
            if (sourceIndex < _sourceItemResultBuilders.size())
            {
                recycleMutation = _sourceItemResultBuilders[sourceIndex].mutation;
            }
        }
        if (! recycleMutation.has_value() || ! IsValidItemMutationResultPrefix(recycleMutation.value()))
        {
            LogDiagnostic(DiagnosticSeverity::Error,
                          HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE),
                          L"delete.recycle.indeterminate",
                          L"Recycle failed without valid per-item mutation truth; permanent deletion was not offered.",
                          sourcePath);
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (recycleMutation->outcomeKnown == FALSE)
        {
            LogDiagnostic(DiagnosticSeverity::Error,
                          HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE),
                          L"delete.recycle.indeterminate",
                          L"Recycle outcome is indeterminate; permanent deletion was not offered.",
                          sourcePath);
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (recycleMutation->mutationCommitted == TRUE && recycleMutation->originalStillPresent == FALSE)
        {
            storeSourceMutationTruth(sourceIndex, S_OK, recycleMutation);
            return S_OK;
        }
        if (recycleMutation->mutationCommitted != FALSE || recycleMutation->originalStillPresent != TRUE)
        {
            LogDiagnostic(DiagnosticSeverity::Error,
                          E_UNEXPECTED,
                          L"delete.recycle.contractViolation",
                          L"Recycle provider returned contradictory mutation truth; permanent deletion was not offered.",
                          sourcePath);
            return E_UNEXPECTED;
        }

#ifdef ENABLE_TESTS
        g_fileOpsRecycleEscalationBeforeBindPausePoint.Pause(10'000ull);
#endif

        switch (GuardLiveOutputBeforeInvalidation(sourcePath, FileOperations::MutationInterlockAccess::WriteSource))
        {
            case LiveOutputGuardDisposition::Skip: storeSourceMutationTruth(sourceIndex, S_FALSE, recycleMutation); return S_FALSE;
            case LiveOutputGuardDisposition::Cancel: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            case LiveOutputGuardDisposition::RetryCurrentMutation:
            case LiveOutputGuardDisposition::Proceed: break;
        }

        const FileOperations::DeletePlan& deletePlan = *deletePlanBySourceIndex[sourceIndex];
        if (! deletePlan.endpoint.pathIdentity.has_value())
        {
            return E_UNEXPECTED;
        }
        const auto retainedScope = std::ranges::find_if(_mutationInterlockScopes,
                                                        [&](const FileOperations::MutationInterlockScope& scope) noexcept
        {
            return scope.access == FileOperations::MutationInterlockAccess::WriteSource && scope.root &&
                   FileOperations::QualifiedEndpointsReferToSameRoot(scope.endpoint, deletePlan.endpoint) &&
                   EquivalentPath(deletePlan.endpoint.pathIdentity.value(), scope.providerPath, sourcePath) &&
                   EquivalentPath(deletePlan.endpoint.pathIdentity.value(), scope.root->retained.providerPath, sourcePath);
        });
        if (retainedScope == _mutationInterlockScopes.end() || ! retainedScope->root || ! retainedScope->root->retained.authority.boundObject)
        {
            LogDiagnostic(DiagnosticSeverity::Error,
                          HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                          L"delete.recycle.identityUnavailable",
                          L"Recycle escalation requires the exact pre-operation source authority.",
                          sourcePath);
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        constexpr FileSystemBindFlags deleteBindFlags =
            static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_DELETE);
        FileOperations::ObjectBindingResult current =
            FileOperations::BindObjectAuthority(_fileSystem.get(), sourcePath, deletePlan.endpoint.profileId, deleteBindFlags);
        if (current.state != FileOperations::ObjectBindingState::Bound)
        {
            return FAILED(current.status) ? current.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        bool sameObject   = false;
        bool sameRevision = false;
        const HRESULT identityHr =
            FileOperations::CrossCheckBoundObjectIdentity(retainedScope->root->retained.authority, current.authority, sameObject, sameRevision);
        if (FAILED(identityHr))
        {
            return identityHr;
        }
        if (! sameObject || ! sameRevision || retainedScope->root->retained.authority.kind != current.authority.kind)
        {
            LogDiagnostic(DiagnosticSeverity::Error,
                          HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH),
                          L"delete.recycle.identityChanged",
                          L"The selected object changed after Recycle failed; permanent deletion was not offered.",
                          sourcePath);
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }

        DeferredConsentRequest request{};
        request.risk                        = FileOperations::DeferredConsentRisk::RecycleEscalation;
        request.status                      = recycleStatus;
        request.perItemCookie               = perItemCookie;
        request.sourcePath                  = sourcePath;
        request.itemIndex                   = sourceIndex;
        request.destinationRootId           = deletePlan.endpoint.rootId;
        request.applyToAllEligible          = false;
        request.itemCountKnown              = true;
        request.itemCount                   = 1u;
        request.bytesKnown                  = current.authority.committedSizeBytes != std::numeric_limits<uint64_t>::max();
        request.bytes                       = request.bytesKnown ? current.authority.committedSizeBytes : 0u;
        request.sourceIdentity              = current.authority.identity;
        const DeferredConsentResult consent = RequestDeferredConsent(*this, request);
        if (FAILED(consent.status))
        {
            return consent.status;
        }
        if (consent.action == ConflictAction::Skip)
        {
            storeSourceMutationTruth(sourceIndex, S_FALSE, recycleMutation);
            return S_FALSE;
        }
        if (consent.action != ConflictAction::PermanentDelete)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        FileSystemOptions deleteOptions{};
        InitializeFileSystemOptions(deleteOptions, perItemCookie);
        FileSystemConditionalMutationResult mutationResult{sizeof(FileSystemConditionalMutationResult), FALSE, TRUE, FALSE};
        // The freshly bound authority owns a no-follow DELETE handle that denies write/delete sharing.
        // Keeping that handle through the prompt is the revalidation boundary: DeleteIfUnchanged
        // mutates the exact retained object and never reopens the pathname after consent.
        const HRESULT deleteHr = current.authority.boundObject->DeleteIfUnchanged(FILESYSTEM_FLAG_NONE, &deleteOptions, &mutationResult);
        const FileOperations::ManagedCleanupAttemptDisposition disposition = FileOperations::ClassifyManagedCleanupMutation({
            .status               = deleteHr,
            .outcomeKnown         = mutationResult.outcomeKnown == TRUE,
            .mutationCommitted    = mutationResult.mutationCommitted == TRUE,
            .originalStillPresent = mutationResult.originalStillPresent == TRUE,
        });

        FileSystemItemMutationResult finalMutation{
            sizeof(FileSystemItemMutationResult), mutationResult.outcomeKnown, mutationResult.mutationCommitted, mutationResult.originalStillPresent};
        switch (disposition)
        {
            case FileOperations::ManagedCleanupAttemptDisposition::Removed:
                storeSourceMutationTruth(sourceIndex, S_OK, finalMutation);
                LogDiagnostic(DiagnosticSeverity::Warning,
                              recycleStatus,
                              L"delete.recycle.escalated",
                              L"Recycle failed; the user granted permanent deletion for this exact item.",
                              sourcePath);
                return S_OK;
            case FileOperations::ManagedCleanupAttemptDisposition::Retained:
                storeSourceMutationTruth(sourceIndex, deleteHr, finalMutation);
                return FAILED(deleteHr) ? deleteHr : E_UNEXPECTED;
            case FileOperations::ManagedCleanupAttemptDisposition::Indeterminate:
                storeSourceMutationTruth(sourceIndex, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE), finalMutation);
                return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
            case FileOperations::ManagedCleanupAttemptDisposition::ProviderContractViolation:
                storeSourceMutationTruth(sourceIndex, E_UNEXPECTED, finalMutation);
                return E_UNEXPECTED;
        }
        return E_UNEXPECTED;
    };

    const auto executePermanentDelete =
        [&](size_t sourceIndex, PerItemCallbackCookie* perItemCookie, std::wstring_view sourcePath, bool& providerMutationAttempted) noexcept -> HRESULT
    {
        if (_operation != FILESYSTEM_DELETE || sourceIndex >= deletePlanBySourceIndex.size() || sourcePath.empty() ||
            deletePlanBySourceIndex[sourceIndex] == nullptr)
        {
            return E_UNEXPECTED;
        }
        const FileOperations::DeletePlan& deletePlan = *deletePlanBySourceIndex[sourceIndex];
        if (deletePlan.mode != FileOperations::DeleteMode::Permanent || ! deletePlan.initialConsent.has_value() ||
            ! deletePlan.endpoint.pathIdentity.has_value())
        {
            return E_UNEXPECTED;
        }

#ifdef ENABLE_TESTS
        g_fileOpsPermanentDeleteBeforeLiveOutputGuardPausePoint.Pause(10'000ull);
#endif
        switch (GuardLiveOutputBeforeInvalidation(sourcePath, FileOperations::MutationInterlockAccess::WriteSource))
        {
            case LiveOutputGuardDisposition::Skip:
                if (perItemCookie != nullptr)
                {
                    perItemCookie->explicitSkipObserved.store(true, std::memory_order_release);
                }
                storeSourceMutationTruth(sourceIndex, S_FALSE, std::nullopt);
                return S_FALSE;
            case LiveOutputGuardDisposition::Cancel: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            case LiveOutputGuardDisposition::RetryCurrentMutation:
            case LiveOutputGuardDisposition::Proceed: break;
        }

        if (deletePlan.nativeAuthority)
        {
            // R0f: the provider has no bound objects. Its contract-tested native delete consumes the
            // selected path and reports its receipt through the item callback; a missing or
            // indeterminate receipt stays Indeterminate in the canonical classifier.
            FileSystemOptions nativeOptions{};
            InitializeFileSystemOptions(nativeOptions, perItemCookie);
            FileSystemFlags nativeFlags = deletePlan.recursive ? FILESYSTEM_FLAG_RECURSIVE : FILESYSTEM_FLAG_NONE;
            if ((_flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0)
            {
                nativeFlags = static_cast<FileSystemFlags>(nativeFlags | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
            }
            const std::wstring nativeSourcePath(sourcePath);
            providerMutationAttempted = true;
            // C10: a root pinned by identity while preparing is deleted only while the live identity
            // still equals it; the provider refuses a different object with ERROR_REVISION_MISMATCH.
            const auto pinnedScope = std::ranges::find_if(_mutationInterlockScopes,
                                                          [&](const FileOperations::MutationInterlockScope& scope) noexcept
            {
                return scope.pinnedDeleteIdentity.has_value() && FileOperations::QualifiedEndpointsReferToSameRoot(scope.endpoint, deletePlan.endpoint) &&
                       EquivalentPath(deletePlan.endpoint.pathIdentity.value(), scope.providerPath, sourcePath);
            });
            if (pinnedScope != _mutationInterlockScopes.end())
            {
                wil::com_ptr<IFileSystemIdentityDelete> identityDelete;
                if (FAILED(_fileSystem->QueryInterface(__uuidof(IFileSystemIdentityDelete), identityDelete.put_void())) || ! identityDelete)
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
                return identityDelete->DeleteIfIdentity(
                    nativeSourcePath.c_str(), &pinnedScope->pinnedDeleteIdentity.value(), nativeFlags, &nativeOptions, this, static_cast<void*>(perItemCookie));
            }
            return _fileSystem->DeleteItem(nativeSourcePath.c_str(), nativeFlags, &nativeOptions, this, static_cast<void*>(perItemCookie));
        }

        const auto retainedScope = std::ranges::find_if(_mutationInterlockScopes,
                                                        [&](const FileOperations::MutationInterlockScope& scope) noexcept
        {
            return scope.exactDeleteAuthority && scope.access == FileOperations::MutationInterlockAccess::WriteSource && scope.root &&
                   FileOperations::QualifiedEndpointsReferToSameRoot(scope.endpoint, deletePlan.endpoint) &&
                   EquivalentPath(deletePlan.endpoint.pathIdentity.value(), scope.providerPath, sourcePath) &&
                   EquivalentPath(deletePlan.endpoint.pathIdentity.value(), scope.root->retained.providerPath, sourcePath);
        });
        if (retainedScope == _mutationInterlockScopes.end() || ! retainedScope->root || ! retainedScope->root->retained.authority.boundObject)
        {
            LogDiagnostic(DiagnosticSeverity::Error,
                          HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                          L"delete.permanent.identityUnavailable",
                          L"Permanent deletion requires the exact object retained from admission through mutation.",
                          sourcePath);
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        FileSystemOptions deleteOptions{};
        InitializeFileSystemOptions(deleteOptions, perItemCookie);
        FileSystemFlags deleteFlags = deletePlan.recursive ? FILESYSTEM_FLAG_RECURSIVE : FILESYSTEM_FLAG_NONE;
        if ((_flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0)
        {
            deleteFlags = static_cast<FileSystemFlags>(deleteFlags | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        }
        FileSystemConditionalMutationResult mutationResult{sizeof(FileSystemConditionalMutationResult), FALSE, TRUE, FALSE};
        _permanentDeleteConditionalAttemptCount.fetch_add(1u, std::memory_order_relaxed);
        HRESULT deleteHr          = E_PENDING;
        providerMutationAttempted = true;
#ifdef ENABLE_TESTS
        if (ConsumeBridgeCounterForSelfTest(g_fileOpsPermanentDeleteKnownNonCommitCount, g_fileOpsPermanentDeleteKnownNonCommitAttempts))
        {
            mutationResult.outcomeKnown         = TRUE;
            mutationResult.mutationCommitted    = FALSE;
            mutationResult.originalStillPresent = TRUE;
            deleteHr                            = g_fileOpsPermanentDeleteKnownNonCommitStatus.load(std::memory_order_acquire);
        }
        else
#endif
        {
            deleteHr = retainedScope->root->retained.authority.boundObject->DeleteIfUnchanged(deleteFlags, &deleteOptions, &mutationResult);
        }
        const FileOperations::ManagedCleanupAttemptDisposition disposition = FileOperations::ClassifyManagedCleanupMutation({
            .status               = deleteHr,
            .outcomeKnown         = mutationResult.outcomeKnown == TRUE,
            .mutationCommitted    = mutationResult.mutationCommitted == TRUE,
            .originalStillPresent = mutationResult.originalStillPresent == TRUE,
        });
        FileSystemItemMutationResult itemMutation{
            sizeof(FileSystemItemMutationResult), mutationResult.outcomeKnown, mutationResult.mutationCommitted, mutationResult.originalStillPresent};
        switch (disposition)
        {
            case FileOperations::ManagedCleanupAttemptDisposition::Removed:
                _permanentDeleteRemovedCount.fetch_add(1u, std::memory_order_relaxed);
                storeSourceMutationTruth(sourceIndex, S_OK, itemMutation);
                return S_OK;
            case FileOperations::ManagedCleanupAttemptDisposition::Retained:
                storeSourceMutationTruth(sourceIndex, deleteHr, itemMutation);
                return FAILED(deleteHr) ? deleteHr : E_UNEXPECTED;
            case FileOperations::ManagedCleanupAttemptDisposition::Indeterminate:
                _permanentDeleteIndeterminateCount.fetch_add(1u, std::memory_order_relaxed);
                storeSourceMutationTruth(sourceIndex, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE), itemMutation);
                return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
            case FileOperations::ManagedCleanupAttemptDisposition::ProviderContractViolation:
                storeSourceMutationTruth(sourceIndex, E_UNEXPECTED, itemMutation);
                return E_UNEXPECTED;
        }
        return E_UNEXPECTED;
    };

    if (_executionMode == ExecutionMode::PerItem)
    {
        wil::com_ptr<IFileSystemIO> fileSystemIo;
        static_cast<void>(_fileSystem->QueryInterface(IID_PPV_ARGS(fileSystemIo.addressof())));
        wil::com_ptr<IFileSystemDirectoryOperations> fileSystemDirOps;
        static_cast<void>(_fileSystem->QueryInterface(IID_PPV_ARGS(fileSystemDirOps.addressof())));

        const bool useCrossFileSystemBridge            = (_destinationFileSystem != nullptr || hasManagedMovePlans || hasVerificationPlans) &&
                                                         (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE);
        IFileSystem* const bridgeDestinationFileSystem = _destinationFileSystem ? _destinationFileSystem.get() : _fileSystem.get();

        wil::com_ptr<IFileSystemIO> destinationFileSystemIo;
        wil::com_ptr<IFileSystemDirectoryOperations> destinationDirOps;
        static_cast<void>(bridgeDestinationFileSystem->QueryInterface(IID_PPV_ARGS(destinationFileSystemIo.addressof())));
        static_cast<void>(bridgeDestinationFileSystem->QueryInterface(IID_PPV_ARGS(destinationDirOps.addressof())));
        if (useCrossFileSystemBridge)
        {
            if (! fileSystemIo || ! destinationFileSystemIo)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }

#ifdef ENABLE_TESTS
            const HRESULT sourceDecorateHr = DecorateBridgeIoForSelfTest(fileSystemIo, SelfTestBridgeIoRole::Source);
            if (FAILED(sourceDecorateHr))
            {
                return sourceDecorateHr;
            }
            const HRESULT destinationDecorateHr = DecorateBridgeIoForSelfTest(destinationFileSystemIo, SelfTestBridgeIoRole::Destination);
            if (FAILED(destinationDecorateHr))
            {
                return destinationDecorateHr;
            }
#endif
        }

        IFileSystem* const effectiveDestinationFileSystem = bridgeDestinationFileSystem;
        const auto safetyCategory                         = [](FileOperations::TransferSafetyState state) noexcept -> std::wstring_view
        {
            switch (state)
            {
                case FileOperations::TransferSafetyState::Ready: return L"identity.ready";
                case FileOperations::TransferSafetyState::SameObject: return L"identity.sameObject";
                case FileOperations::TransferSafetyState::SameFolderMove: return L"identity.sameFolderMove";
                case FileOperations::TransferSafetyState::DestinationInsideSource: return L"identity.destinationInsideSource";
                case FileOperations::TransferSafetyState::AncestryLink: return L"identity.ancestryLink";
                case FileOperations::TransferSafetyState::SourceMissing: return L"identity.sourceMissing";
                case FileOperations::TransferSafetyState::DestinationChanged: return L"identity.destinationChanged";
                case FileOperations::TransferSafetyState::SourceChanged: return L"identity.sourceChanged";
                case FileOperations::TransferSafetyState::Unsupported: return L"identity.unsupported";
                case FileOperations::TransferSafetyState::Indeterminate: return L"identity.indeterminate";
                case FileOperations::TransferSafetyState::ProviderContractViolation: return L"identity.providerContractViolation";
            }
            return L"identity.unknown";
        };
        const auto prepareExactTransferGuard = [&](size_t index,
                                                   std::wstring_view sourcePath,
                                                   std::wstring_view destinationPath,
                                                   FileOperations::TransferMutationGuard& guard,
                                                   bool& samePathCopy) noexcept -> HRESULT
        {
            samePathCopy = false;
            if (index >= transferPlanBySourceIndex.size())
            {
                return E_UNEXPECTED;
            }
            const FileOperations::TransferPlan& plan = *transferPlanBySourceIndex[index];
            guard                                    = FileOperations::PrepareTransferMutationGuard(
                _fileSystem.get(), effectiveDestinationFileSystem, plan.sourceEndpoint, plan.destinationEndpoint, plan.intent, sourcePath, destinationPath);
            if (guard.state == FileOperations::TransferSafetyState::SameObject && plan.intent == FileOperations::TransferIntent::Copy && guard.samePathText)
            {
                samePathCopy = true;
                return S_OK;
            }
            if (guard.state == FileOperations::TransferSafetyState::Unsupported)
            {
                const bool sameEndpoint   = FileOperations::QualifiedEndpointsReferToSameRoot(plan.sourceEndpoint, plan.destinationEndpoint);
                const bool equivalentPath = sameEndpoint && plan.sourceEndpoint.pathIdentity.has_value() &&
                                            EquivalentPath(plan.sourceEndpoint.pathIdentity.value(), sourcePath, destinationPath);
                if (equivalentPath)
                {
                    if (plan.intent == FileOperations::TransferIntent::Copy)
                    {
                        samePathCopy = true;
                        return S_OK;
                    }

                    const HRESULT sameMoveHr = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                    LogDiagnostic(DiagnosticSeverity::Error,
                                  sameMoveHr,
                                  L"identity.sameFolderMove",
                                  L"The provider cannot prove object identity and the Move source and destination paths are equivalent.",
                                  sourcePath,
                                  destinationPath);
                    return sameMoveHr;
                }

                // Binding is optional for honest Copy, a qualified provider-native Move, and
                // Copy-only. It remains mandatory before managed destructive source cleanup.
                if (plan.intent == FileOperations::TransferIntent::Copy || plan.strategy == FileOperations::OperationStrategy::Native ||
                    plan.strategy == FileOperations::OperationStrategy::CopyOnly)
                {
                    return S_FALSE;
                }
                // R3-2: a Managed Move into a route that proves published content through its writer
                // binds no destination object by design; the bridge binds the source itself and
                // removes it only after the provider's proof of the published object matched.
                if (plan.strategy == FileOperations::OperationStrategy::Managed && plan.destinationEndpoint.verificationWriterDigestProof)
                {
                    return S_FALSE;
                }
            }
            if (guard.state != FileOperations::TransferSafetyState::Ready)
            {
                LogDiagnostic(DiagnosticSeverity::Error,
                              guard.status,
                              safetyCategory(guard.state),
                              L"Exact no-follow transfer safety guard rejected the item before provider mutation.",
                              sourcePath,
                              destinationPath);
                return guard.status;
            }
            return S_OK;
        };
        const auto revalidateExactTransferGuard = [&](size_t index,
                                                      std::wstring_view sourcePath,
                                                      std::wstring_view destinationPath,
                                                      const FileOperations::TransferMutationGuard& guard,
                                                      bool allowMissingDestinationDirectoryMerge) noexcept -> HRESULT
        {
            if (guard.state != FileOperations::TransferSafetyState::Ready)
            {
                return S_OK;
            }
            if (index >= transferPlanBySourceIndex.size())
            {
                return E_UNEXPECTED;
            }
            const FileOperations::TransferPlan& plan        = *transferPlanBySourceIndex[index];
            HRESULT status                                  = E_UNEXPECTED;
            const FileOperations::TransferSafetyState state = FileOperations::RevalidateTransferMutationGuard(_fileSystem.get(),
                                                                                                              effectiveDestinationFileSystem,
                                                                                                              plan.sourceEndpoint,
                                                                                                              plan.destinationEndpoint,
                                                                                                              plan.intent,
                                                                                                              sourcePath,
                                                                                                              destinationPath,
                                                                                                              guard,
                                                                                                              status,
                                                                                                              allowMissingDestinationDirectoryMerge);
            if (state != FileOperations::TransferSafetyState::Ready)
            {
                LogDiagnostic(DiagnosticSeverity::Error,
                              status,
                              safetyCategory(state),
                              L"Exact no-follow transfer safety revalidation rejected a changed item at the provider boundary.",
                              sourcePath,
                              destinationPath);
                return status;
            }
            return S_OK;
        };

        // The cross-filesystem bridge is only used for copy/move, so hoisting these capability lookups is safe:
        // per-item conflict handling may tweak `itemFlags`, but delete-only flag-sensitive keys are never consulted here.
        const unsigned int bridgeSourceMaxConcurrencyBudget =
            useCrossFileSystemBridge
                ? DeterminePerItemMaxConcurrency(_fileSystem, _sourcePaths, _sourcePluginId, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles))
                : 1u;
        const unsigned int bridgeDestinationMaxConcurrencyBudget =
            useCrossFileSystemBridge
                ? DeterminePerItemMaxConcurrency(
                      _destinationFileSystem, destinationFolder, _destinationPluginId, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles))
                : 1u;
        const Common::Settings::Settings* settingsSnapshot = (_folderWindow != nullptr) ? _folderWindow->_settings : nullptr;
        const bool sourceUsesAutoConcurrency               = ShouldUseAutoPerItemConcurrency(_fileSystem, _operation, _flags);
        const bool destinationUsesAutoConcurrency = useCrossFileSystemBridge && ShouldUseAutoPerItemConcurrency(_destinationFileSystem, _operation, _flags);
        const AutoConcurrencyResolution sourceAutoResolution =
            sourceUsesAutoConcurrency ? ResolveAutoPerItemMaxConcurrency(_fileSystem, _sourcePaths, _operation, static_cast<unsigned int>(kMaxInFlightFiles))
                                      : AutoConcurrencyResolution{};
        const AutoConcurrencyResolution destinationAutoResolution =
            destinationUsesAutoConcurrency
                ? ResolveAutoPerItemMaxConcurrency(_destinationFileSystem, destinationFolder, _operation, static_cast<unsigned int>(kMaxInFlightFiles))
                : AutoConcurrencyResolution{};

        if (! plans || plans->empty())
        {
            LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                          E_UNEXPECTED,
                          L"Plan",
                          L"The admitted task lost its immutable FileOperationPlan group before execution.");
            return E_UNEXPECTED;
        }
        const FileOperations::OperationOptions planOptions = std::visit([](const auto& typedPlan) noexcept { return typedPlan.options; }, plans->front());
        ReparsePointPolicy reparsePointPolicy =
            planOptions.linkPolicy == FileOperations::LinkPolicy::Skip ? ReparsePointPolicy::Skip : ReparsePointPolicy::Preserve;
#ifdef ENABLE_TESTS
        const int reparsePolicyOverride = g_fileOpsBridgeReparsePolicyOverride.load(std::memory_order_acquire);
        if (reparsePolicyOverride >= static_cast<int>(ReparsePointPolicy::Preserve) && reparsePolicyOverride <= static_cast<int>(ReparsePointPolicy::Skip))
        {
            reparsePointPolicy = static_cast<ReparsePointPolicy>(reparsePolicyOverride);
        }
#endif

        const uint64_t count64 = static_cast<uint64_t>(_sourcePaths.size());
        if (count64 > std::numeric_limits<unsigned long>::max())
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        _perItemTotalItems = static_cast<unsigned long>(count64);
        _perItemMaxConcurrencyBudget =
            DeterminePerItemMaxConcurrency(_fileSystem, _sourcePaths, _sourcePluginId, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles));
        _perItemMaxConcurrencyBudget = std::max(1u, _perItemMaxConcurrencyBudget);
        if (useCrossFileSystemBridge)
        {
            const unsigned int destinationMaxConcurrencyBudget = DeterminePerItemMaxConcurrency(
                _destinationFileSystem, destinationFolder, _destinationPluginId, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles));
            _perItemMaxConcurrencyBudget = std::min(_perItemMaxConcurrencyBudget, destinationMaxConcurrencyBudget);
            _perItemMaxConcurrencyBudget = std::max(1u, _perItemMaxConcurrencyBudget);
        }

        {
            const bool isCopyMove = (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE);
            const bool isDelete   = (_operation == FILESYSTEM_DELETE);

            const char* overrideKey = nullptr;
            uint32_t overrideMin    = 0;
            uint32_t overrideMax    = 0;
            if (isCopyMove)
            {
                overrideKey = "copyMoveMaxConcurrency";
                overrideMin = 1u;
                overrideMax = 16u;
            }
            else if (isDelete)
            {
                overrideKey = "deleteMaxConcurrency";
                overrideMin = 1u;
                overrideMax = 64u;
            }

            if (overrideKey && settingsSnapshot)
            {
                std::optional<uint32_t> minOverride;

                const auto applyOverrideFromPath = [&](std::wstring_view pluginPath) noexcept
                {
                    const auto connNameOpt = ConnectionProfileUtils::TryParseConnNameFromPluginPath(pluginPath);
                    if (! connNameOpt.has_value())
                    {
                        return;
                    }

                    const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(settingsSnapshot, *connNameOpt);
                    if (! profile)
                    {
                        return;
                    }

                    const uint32_t rawValue = ConnectionProfileUtils::ExtraGetUInt32(profile->extra, overrideKey).value_or(0);
                    if (rawValue == 0)
                    {
                        return;
                    }

                    const uint32_t clamped = std::clamp(rawValue, overrideMin, overrideMax);
                    minOverride            = minOverride.has_value() ? std::min<uint32_t>(*minOverride, clamped) : clamped;
                };

                if (isCopyMove)
                {
                    applyOverrideFromPath(destinationFolder.native());
                }
                for (const std::filesystem::path& sourcePath : _sourcePaths)
                {
                    applyOverrideFromPath(sourcePath.native());
                }

                if (minOverride.has_value())
                {
                    _perItemMaxConcurrencyBudget = std::min<unsigned int>(_perItemMaxConcurrencyBudget, *minOverride);
                }
            }
        }
        if (hasVerificationPlans)
        {
            // Verification is a sequential phase of each item in this slice. Discovery may continue
            // ahead, but the next transfer must not start until the exact published object has been
            // proved or the current item has reached a terminal verification state.
            _perItemMaxConcurrencyBudget = 1u;
        }
        _autoConcurrencyUsed.store(false, std::memory_order_release);
        _autoConcurrencyStorageKind.store(FILESYSTEM_STORAGE_UNKNOWN, std::memory_order_release);
        _autoConcurrencyDestinationStorageKind.store(FILESYSTEM_STORAGE_UNKNOWN, std::memory_order_release);
        _autoTunedConcurrency.store(0u, std::memory_order_release);
        if (sourceAutoResolution.HasValue() || destinationAutoResolution.HasValue())
        {
            _autoConcurrencyUsed.store(true, std::memory_order_release);

            if (useCrossFileSystemBridge && destinationAutoResolution.HasValue())
            {
                _autoConcurrencyDestinationStorageKind.store(destinationAutoResolution.storageKind, std::memory_order_release);
            }

            if (sourceAutoResolution.HasValue() && destinationAutoResolution.HasValue())
            {
                if (sourceAutoResolution.concurrency < destinationAutoResolution.concurrency)
                {
                    _autoTunedConcurrency.store(sourceAutoResolution.concurrency, std::memory_order_release);
                    _autoConcurrencyStorageKind.store(sourceAutoResolution.storageKind, std::memory_order_release);
                }
                else if (destinationAutoResolution.concurrency < sourceAutoResolution.concurrency)
                {
                    _autoTunedConcurrency.store(destinationAutoResolution.concurrency, std::memory_order_release);
                    _autoConcurrencyStorageKind.store(destinationAutoResolution.storageKind, std::memory_order_release);
                }
                else
                {
                    _autoTunedConcurrency.store(sourceAutoResolution.concurrency, std::memory_order_release);
                    _autoConcurrencyStorageKind.store(sourceAutoResolution.storageKind, std::memory_order_release);
                    if (sourceAutoResolution.storageKind != destinationAutoResolution.storageKind)
                    {
                        _autoConcurrencyStorageKind.store(FILESYSTEM_STORAGE_UNKNOWN, std::memory_order_release);
                    }
                }
            }
            else if (sourceAutoResolution.HasValue())
            {
                _autoTunedConcurrency.store(sourceAutoResolution.concurrency, std::memory_order_release);
                _autoConcurrencyStorageKind.store(sourceAutoResolution.storageKind, std::memory_order_release);
            }
            else
            {
                _autoTunedConcurrency.store(destinationAutoResolution.concurrency, std::memory_order_release);
                _autoConcurrencyStorageKind.store(destinationAutoResolution.storageKind, std::memory_order_release);
            }
        }

        _perItemMaxConcurrency = std::min<unsigned int>(_perItemMaxConcurrencyBudget, static_cast<unsigned int>(_perItemTotalItems));
        _effectiveConcurrencyBudget.store(_perItemMaxConcurrencyBudget, std::memory_order_release);
        _perItemCompletedItems      = 0;
        _perItemCompletedEntryCount = 0;
        _perItemTotalEntryCount     = 0;
        _perItemCompletedBytes      = 0;
        {
            std::scoped_lock lock(_perItemInFlightCallsMutex);
            _perItemInFlightCallCount      = 0;
            _perItemInFlightCompletedBytes = 0;
            _perItemInFlightCompletedItems = 0;
            _perItemInFlightTotalItems     = 0;
        }
        {
            std::scoped_lock lock(_inFlightFilesMutex);
            _inFlightFileCount = 0;
        }

#ifdef ENABLE_TESTS
        _dbgConfiguredMaxConcurrency = std::max(1u, _perItemMaxConcurrencyBudget);
#endif

        {
            std::scoped_lock lock(_progressMutex);
            if (_operation != FILESYSTEM_DELETE)
            {
                _progressTotalItems = _perItemTotalItems;
            }
            _progressCompletedItems = 0;
            _progressCompletedBytes = 0;
            PublishProgressCountersLocked(*this);
        }

        if ((_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE) && destinationFolder.empty())
        {
            return E_INVALIDARG;
        }

        const std::wstring destinationFolderText = destinationFolder.native();
        const auto resolveTransferDestination =
            [&](const size_t index, const FileOperations::TransferPlan* transferPlan, std::wstring& destinationOut) noexcept -> HRESULT
        {
            if (transferPlan == nullptr || index >= transferPlanItemIndexBySourceIndex.size() ||
                ! FileOperations::TryResolveTransferDestinationProviderPath(*transferPlan, transferPlanItemIndexBySourceIndex[index], destinationOut))
            {
                return E_UNEXPECTED;
            }

            if (useResolvedItems)
            {
                if (! transferPlan->destinationEndpoint.pathIdentity.has_value() ||
                    ! EquivalentPath(transferPlan->destinationEndpoint.pathIdentity.value(), destinationOut, _resolvedItems[index].destinationPath.native()))
                {
                    return E_UNEXPECTED;
                }
            }
            return S_OK;
        };
        std::wstring destinationCircuitBreakerConnectionId;
        std::vector<std::wstring> sourceCircuitBreakerConnectionIds;
        if (settingsSnapshot)
        {
            sourceCircuitBreakerConnectionIds.reserve(_sourcePaths.size());
            if (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE)
            {
                destinationCircuitBreakerConnectionId = ResolveCircuitBreakerConnectionId(settingsSnapshot, destinationFolderText);
            }

            for (const std::filesystem::path& sourcePath : _sourcePaths)
            {
                sourceCircuitBreakerConnectionIds.push_back(ResolveCircuitBreakerConnectionId(settingsSnapshot, sourcePath.native()));
            }
        }

        const auto getSourceCircuitBreakerConnectionId = [&](size_t index) noexcept -> std::wstring_view
        { return index < sourceCircuitBreakerConnectionIds.size() ? std::wstring_view(sourceCircuitBreakerConnectionIds[index]) : std::wstring_view{}; };

        constexpr unsigned int kMaxCachedModifierAttemptsPerBucket = 1u;

        const auto getPerItemInFlightAggregate = [&]() noexcept -> PerItemInFlightAggregate
        {
            std::scoped_lock lock(_perItemInFlightCallsMutex);
            return SummarizePerItemInFlightCallsLocked(*this);
        };

        struct BridgeCallback final : IFileSystemCallback
        {
            Task& task;
            std::mutex* callbackMutex = nullptr;

            explicit BridgeCallback(Task& owner, std::mutex* callbackMutexIn = nullptr) noexcept : task(owner), callbackMutex(callbackMutexIn)
            {
            }

            BridgeCallback(const BridgeCallback&)            = delete;
            BridgeCallback(BridgeCallback&&)                 = delete;
            BridgeCallback& operator=(const BridgeCallback&) = delete;
            BridgeCallback& operator=(BridgeCallback&&)      = delete;

            HRESULT STDMETHODCALLTYPE FileSystemProgress(FileSystemOperation /*operationType*/,
                                                         unsigned long /*totalItems*/,
                                                         unsigned long /*completedItems*/,
                                                         uint64_t /*totalBytes*/,
                                                         uint64_t /*completedBytes*/,
                                                         const wchar_t* /*currentSourcePath*/,
                                                         const wchar_t* /*currentDestinationPath*/,
                                                         uint64_t /*currentItemTotalBytes*/,
                                                         uint64_t /*currentItemCompletedBytes*/,
                                                         FileSystemOptions* /*options*/,
                                                         uint64_t /*progressStreamId*/,
                                                         void* /*cookie*/) noexcept override
            {
                task.WaitWhilePaused();
                if (task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested())
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                return S_OK;
            }

            HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation /*operationType*/,
                                                              unsigned long /*itemIndex*/,
                                                              const wchar_t* /*sourcePath*/,
                                                              const wchar_t* /*destinationPath*/,
                                                              HRESULT /*status*/,
                                                              const FileSystemItemMutationResult* /*mutationResult*/,
                                                              FileSystemOptions* /*options*/,
                                                              void* /*cookie*/) noexcept override
            {
                return S_OK;
            }

            HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, void* cookie) noexcept override
            {
                if (callbackMutex != nullptr)
                {
                    std::scoped_lock lock(*callbackMutex);
                    return task.FileSystemShouldCancel(pCancel, cookie);
                }
                return task.FileSystemShouldCancel(pCancel, cookie);
            }

            HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation operationType,
                                                      const wchar_t* sourcePath,
                                                      const wchar_t* destinationPath,
                                                      HRESULT status,
                                                      FileSystemIssueAction* action,
                                                      IFileSystemBoundObject** expectedDestination,
                                                      FileSystemOptions* options,
                                                      void* cookie) noexcept override
            {
                if (callbackMutex != nullptr)
                {
                    std::scoped_lock lock(*callbackMutex);
                    return task.FileSystemIssue(operationType, sourcePath, destinationPath, status, action, expectedDestination, options, cookie);
                }
                return task.FileSystemIssue(operationType, sourcePath, destinationPath, status, action, expectedDestination, options, cookie);
            }
        };

        using OwnedStageDisposition = FileOperations::OwnedStageDisposition;

        const auto failurePhaseText = [](QualifiedItemFailurePhase phase) noexcept -> std::wstring_view
        {
            switch (phase)
            {
                case QualifiedItemFailurePhase::DestinationParent: return L"DestinationParent";
                case QualifiedItemFailurePhase::StageCreate: return L"StageCreate";
                case QualifiedItemFailurePhase::StageIdentity: return L"StageIdentity";
                case QualifiedItemFailurePhase::StageWrite: return L"StageWrite";
                case QualifiedItemFailurePhase::StageCommit: return L"StageCommit";
                case QualifiedItemFailurePhase::FinalPublish: return L"FinalPublish";
                case QualifiedItemFailurePhase::FinalPublishReconcile: return L"FinalPublishReconcile";
                case QualifiedItemFailurePhase::ReplacementRollback: return L"ReplacementRollback";
                case QualifiedItemFailurePhase::StageAbort: return L"StageAbort";
                case QualifiedItemFailurePhase::Verification: return L"Verification";
                case QualifiedItemFailurePhase::SourceCleanup: return L"SourceCleanup";
                case QualifiedItemFailurePhase::ProviderNative: return L"ProviderNative";
                case QualifiedItemFailurePhase::None:
                default: return L"None";
            }
        };

        const auto stageDispositionText = [](OwnedStageDisposition disposition) noexcept -> std::wstring_view
        {
            switch (disposition)
            {
                case OwnedStageDisposition::Owned: return L"Owned";
                case OwnedStageDisposition::Removed: return L"Removed";
                case OwnedStageDisposition::Published: return L"Published";
                case OwnedStageDisposition::Retained: return L"Retained";
                case OwnedStageDisposition::Unknown: return L"Unknown";
                case OwnedStageDisposition::RetainedIncomplete: return L"RetainedIncomplete";
                case OwnedStageDisposition::NotApplicable: return L"NotApplicable";
                case OwnedStageDisposition::NotCreated:
                default: return L"NotCreated";
            }
        };

        const auto ensureResolvedDirectoryShell = [&]([[maybe_unused]] const std::wstring& sourceText,
                                                      const std::wstring& destinationText,
                                                      const FileSystemPathIdentity& destinationIdentity,
                                                      [[maybe_unused]] FileSystemFlags itemFlags) noexcept -> HRESULT
        {
            const auto isMissingAttributesFailure = [](HRESULT hr) noexcept
            {
                return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ||
                       hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
            };

            IFileSystemIO* targetIo                      = useCrossFileSystemBridge ? destinationFileSystemIo.get() : fileSystemIo.get();
            IFileSystemDirectoryOperations* targetDirOps = useCrossFileSystemBridge ? destinationDirOps.get() : fileSystemDirOps.get();
            if (! targetIo || ! targetDirOps)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }

            const std::optional<std::wstring> destinationShareRoot = TryGetUncShareRootBoundary(destinationText);
            if (IsUncShareRootBoundary(destinationText, destinationShareRoot))
            {
                return S_OK;
            }

            const auto ensureOneDirectory = [&](const std::wstring& pathText) noexcept -> HRESULT
            {
                for (unsigned int attempt = 0; attempt < 3u; ++attempt)
                {
                    unsigned long attributes = 0;
                    const HRESULT hrAttr     = targetIo->GetAttributes(pathText.c_str(), &attributes);
                    if (SUCCEEDED(hrAttr))
                    {
                        if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                        {
                            return S_OK;
                        }

                        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                    }

                    if (! isMissingAttributesFailure(hrAttr))
                    {
                        return hrAttr;
                    }

                    const HRESULT hrCreate = targetDirOps->CreateDirectory(pathText.c_str());
                    if (SUCCEEDED(hrCreate))
                    {
                        return S_OK;
                    }
                    if (hrCreate == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) || hrCreate == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
                    {
                        continue;
                    }
                    return hrCreate;
                }

                LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                              HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                              L"resolved.directoryShell.collision",
                              L"Resolved directory placeholder could not be created.",
                              sourceText,
                              pathText);
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            };

            std::vector<std::wstring> ancestors;
            std::wstring currentPath(destinationText);
            for (size_t depth = 0u; depth < 1024u; ++depth)
            {
                std::wstring parentPath;
                if (! TryGetFileSystemParentPath(destinationIdentity, currentPath, parentPath))
                {
                    currentPath.clear();
                    break;
                }
                if (EquivalentPath(destinationIdentity, currentPath, parentPath))
                {
                    return E_INVALIDARG;
                }
                if (IsUncShareRootBoundary(parentPath, destinationShareRoot))
                {
                    currentPath.clear();
                    break;
                }

                ancestors.emplace_back(parentPath);
                currentPath = std::move(parentPath);
            }
            if (! currentPath.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
            }

            for (auto it = ancestors.rbegin(); it != ancestors.rend(); ++it)
            {
                const HRESULT hrParent = ensureOneDirectory(*it);
                if (FAILED(hrParent))
                {
                    return hrParent;
                }
            }

            return ensureOneDirectory(destinationText);
        };

        const auto ensureResolvedDestinationParent = [&](const std::wstring& sourceText,
                                                         const std::wstring& destinationText,
                                                         const FileSystemPathIdentity& destinationIdentity,
                                                         bool isDirectoryShell,
                                                         FileSystemFlags itemFlags) noexcept -> HRESULT
        {
            if (! useResolvedItems || isDirectoryShell || destinationText.empty())
            {
                return S_OK;
            }

            std::wstring parentPath;
            if (! TryGetFileSystemParentPath(destinationIdentity, destinationText, parentPath))
            {
                return S_OK;
            }
            if (EquivalentPath(destinationIdentity, destinationText, parentPath))
            {
                return E_INVALIDARG;
            }
            const std::optional<std::wstring> destinationShareRoot = TryGetUncShareRootBoundary(destinationText);
            if (IsUncShareRootBoundary(parentPath, destinationShareRoot))
            {
                return S_OK;
            }

            return ensureResolvedDirectoryShell(sourceText, parentPath, destinationIdentity, itemFlags);
        };

        struct QualifiedItemMutationResult final
        {
            HRESULT hr                                       = E_NOTIMPL;
            unsigned long bridgeSkippedDirectoryReparseCount = 0;
            bool bridgeRootDirectoryReparseSkipped           = false;
            bool bridgeUnsupportedReparse                    = false;
            bool providerMutationAttempted                   = false;
            bool managedSourceRetained                       = false;
            bool managedCleanupIndeterminate                 = false;
            bool ownedStageCleanupIndeterminate              = false;
            FileOperations::VerificationState verification   = FileOperations::VerificationState::NotRequested;
            bool verificationNotApplicableForSelectedRoot    = false;
            std::optional<FileOperations::PublicationState> bridgePublication;
            QualifiedItemFailurePhase failurePhase      = QualifiedItemFailurePhase::None;
            HRESULT failureStatus                       = S_OK;
            uint32_t mutationAttemptFlags               = 0u;
            OwnedStageDisposition ownedStageDisposition = OwnedStageDisposition::NotCreated;
        };

        enum class QualifiedItemTerminalDisposition : uint8_t
        {
            ContinueToConflict,
            Success,
            Skipped,
            Partial,
            PartialFailure,
            Canceled,
            TraversalLimit,
            CleanupIndeterminate,
            VerificationFailure,
        };
        const auto classifyQualifiedItemTerminal = [](HRESULT itemHr,
                                                      bool traversalLimitReached,
                                                      bool managedCleanupIndeterminate,
                                                      FileOperations::VerificationState verification,
                                                      bool explicitSkipObserved) noexcept
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || itemHr == E_ABORT)
            {
                return QualifiedItemTerminalDisposition::Canceled;
            }
            if (traversalLimitReached)
            {
                return QualifiedItemTerminalDisposition::TraversalLimit;
            }
            if (itemHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
            {
                return explicitSkipObserved ? QualifiedItemTerminalDisposition::Partial : QualifiedItemTerminalDisposition::PartialFailure;
            }
            if (itemHr == S_FALSE && explicitSkipObserved)
            {
                return QualifiedItemTerminalDisposition::Skipped;
            }
            if (managedCleanupIndeterminate)
            {
                return QualifiedItemTerminalDisposition::CleanupIndeterminate;
            }
            if (verification == FileOperations::VerificationState::Failed)
            {
                return QualifiedItemTerminalDisposition::VerificationFailure;
            }
            return SUCCEEDED(itemHr) ? QualifiedItemTerminalDisposition::Success : QualifiedItemTerminalDisposition::ContinueToConflict;
        };
        const auto traversalLimitReachedForCall = [](HRESULT itemHr, uint64_t generationBefore, uint64_t generationAfter) noexcept
        { return FAILED(itemHr) && (generationAfter != generationBefore || IsTraversalResourceLimitStatus(itemHr)); };

        // Serial and parallel scheduling differ, but strategy dispatch and provider mutation authority do not.
        // Keep the qualified plan as the single owner so a managed source-delete path cannot reappear in only one loop.
        const auto executeQualifiedItemMutation = [&](size_t index,
                                                      const FileOperations::TransferPlan* transferPlan,
                                                      FileOperations::OperationStrategy itemStrategy,
                                                      const std::wstring& sourceText,
                                                      const std::wstring& destinationItemText,
                                                      FileSystemFlags itemFlags,
                                                      bool isDirectoryShell,
                                                      PerItemCallbackCookie& cookie) noexcept -> QualifiedItemMutationResult
        {
            QualifiedItemMutationResult result{};
            ConnectionCircuitBreaker& breaker = GetConnectionCircuitBreaker();
            {
                std::scoped_lock lock(_sourceItemStatusMutex);
                if (index < _sourceItemResultBuilders.size())
                {
                    // Receipts describe one attempt; a Retry must obtain fresh mutation truth.
                    _sourceItemResultBuilders[index].status.reset();
                    _sourceItemResultBuilders[index].mutation.reset();
                }
            }

            if (_operation == FILESYSTEM_MOVE && transferPlan != nullptr && itemStrategy != FileOperations::OperationStrategy::CopyOnly)
            {
                switch (GuardLiveOutputBeforeInvalidation(sourceText, FileOperations::MutationInterlockAccess::WriteSource))
                {
                    case LiveOutputGuardDisposition::Skip:
                        cookie.explicitSkipObserved.store(true, std::memory_order_release);
                        result.hr = S_FALSE;
                        return result;
                    case LiveOutputGuardDisposition::Cancel: result.hr = HRESULT_FROM_WIN32(ERROR_CANCELLED); return result;
                    case LiveOutputGuardDisposition::RetryCurrentMutation:
                    case LiveOutputGuardDisposition::Proceed: break;
                }
            }

            if (isDirectoryShell)
            {
                if (transferPlan == nullptr || ! transferPlan->destinationEndpoint.pathIdentity.has_value())
                {
                    result.hr = E_UNEXPECTED;
                    return result;
                }
                result.hr = RunWithCircuitBreaker(
                    breaker, getSourceCircuitBreakerConnectionId(index), destinationCircuitBreakerConnectionId, [&]() noexcept -> HRESULT {
                    return ensureResolvedDirectoryShell(sourceText, destinationItemText, transferPlan->destinationEndpoint.pathIdentity.value(), itemFlags);
                });
                result.verification = FileOperations::VerificationState::NotApplicable;
                if (SUCCEEDED(result.hr))
                {
                    // A resolved DirectoryShell is a host publication, not a provider mutation: the
                    // destination folder exists and the source folder, which keeps every unlisted
                    // child, is untouched. Record that truth as the item's receipt so the result
                    // reducer never treats the shell as an unproven destructive call.
                    storeSourceMutationTruth(
                        index,
                        result.hr,
                        FileSystemItemMutationResult{
                            .sizeBytes = sizeof(FileSystemItemMutationResult), .outcomeKnown = TRUE, .mutationCommitted = TRUE, .originalStillPresent = TRUE});
                }
                return result;
            }

            const bool copyOnlyMove = _operation == FILESYSTEM_MOVE && transferPlan != nullptr && itemStrategy == FileOperations::OperationStrategy::CopyOnly;
            const bool managedMove  = _operation == FILESYSTEM_MOVE && transferPlan != nullptr && itemStrategy == FileOperations::OperationStrategy::Managed;
            if (_operation == FILESYSTEM_COPY || copyOnlyMove || managedMove)
            {
                if (transferPlan == nullptr)
                {
                    result.hr = E_UNEXPECTED;
                    return result;
                }
                result.hr = RunWithCircuitBreaker(breaker,
                                                  getSourceCircuitBreakerConnectionId(index),
                                                  destinationCircuitBreakerConnectionId,
                                                  [&]() noexcept -> HRESULT
                {
                    if (useCrossFileSystemBridge)
                    {
                        CrossFileSystemBridge bridge(*this,
                                                     *_fileSystem,
                                                     *effectiveDestinationFileSystem,
                                                     *fileSystemIo,
                                                     *destinationFileSystemIo,
                                                     destinationDirOps.get(),
                                                     bridgeSourceMaxConcurrencyBudget,
                                                     bridgeDestinationMaxConcurrencyBudget,
                                                     itemFlags,
                                                     static_cast<void*>(&cookie),
                                                     0u,
                                                     sourceText.c_str(),
                                                     destinationItemText.c_str(),
                                                     (index < _sourcePathAttributesHint.size()) ? _sourcePathAttributesHint[index] : 0,
                                                     reparsePointPolicy,
                                                     transferPlan->sourceEndpoint.profileId,
                                                     transferPlan->destinationEndpoint.profileId,
                                                     transferPlan->sourceEndpoint.pathIdentity.value(),
                                                     transferPlan->destinationEndpoint.pathIdentity.value(),
                                                     managedMove,
                                                     transferPlan->options.verifyAfterCopy,
                                                     transferPlan->destinationEndpoint.verificationHostReadback,
                                                     transferPlan->destinationEndpoint.verificationProviderBlake3Proof,
                                                     transferPlan->destinationEndpoint.verificationWriterDigestProof);
                        const HRESULT bridgeHr                    = bridge.TraverseAndPublishPath(sourceText, destinationItemText);
                        result.bridgeSkippedDirectoryReparseCount = bridge.skippedDirectoryReparseCount + bridge.skippedFileReparseCount;
                        result.bridgeRootDirectoryReparseSkipped  = bridge.rootDirectoryReparseSkipped;
                        result.bridgeUnsupportedReparse           = bridge.unsupportedReparseEncountered;
                        result.managedSourceRetained              = bridge.managedSourceRetained.load(std::memory_order_acquire);
                        result.managedCleanupIndeterminate        = bridge.managedCleanupIndeterminate.load(std::memory_order_acquire);
                        result.ownedStageCleanupIndeterminate     = bridge.ownedStageCleanupIndeterminate.load(std::memory_order_acquire);
                        // Verification is a file axis. A selected directory remains NotApplicable;
                        // nested file verification still contributes progress and telemetry.
                        result.verification                             = bridge.verificationState.load(std::memory_order_acquire);
                        result.verificationNotApplicableForSelectedRoot = bridge.rootVerificationNotApplicable;
                        result.bridgePublication =
                            bridge.destinationPublicationUnknown.load(std::memory_order_acquire)
                                ? FileOperations::PublicationState::Unknown
                                : (bridge.anyDestinationPublished.load(std::memory_order_acquire) ? FileOperations::PublicationState::Published
                                                                                                  : FileOperations::PublicationState::NotPublished);
                        result.failurePhase          = bridge.failurePhase.load(std::memory_order_acquire);
                        result.failureStatus         = bridge.failureStatus.load(std::memory_order_acquire);
                        result.mutationAttemptFlags  = bridge.mutationAttemptFlags.load(std::memory_order_acquire);
                        result.ownedStageDisposition = bridge.ownedStageDisposition.load(std::memory_order_acquire);
                        if (FAILED(bridgeHr) && result.failurePhase == QualifiedItemFailurePhase::None)
                        {
                            result.failurePhase  = QualifiedItemFailurePhase::ProviderNative;
                            result.failureStatus = bridgeHr;
                        }
                        if (result.verification == FileOperations::VerificationState::Failed && result.failurePhase == QualifiedItemFailurePhase::None)
                        {
                            result.failurePhase  = QualifiedItemFailurePhase::Verification;
                            result.failureStatus = bridgeHr;
                        }
                        return bridgeHr;
                    }

                    if (managedMove)
                    {
                        return E_UNEXPECTED;
                    }

                    FileSystemOptions options{};
                    InitializeFileSystemOptions(options, static_cast<void*>(&cookie));
                    result.providerMutationAttempted = true;
                    return _fileSystem->CopyItem(sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                });
                return result;
            }

            if (_operation == FILESYSTEM_MOVE)
            {
                // The Native route is one provider namespace: the endpoint tuple admitted it, and the
                // source provider executes the rename whatever provider object the destination pane holds.
                if (transferPlan == nullptr || itemStrategy != FileOperations::OperationStrategy::Native)
                {
                    result.hr = E_UNEXPECTED;
                    return result;
                }

                result.hr = RunWithCircuitBreaker(breaker,
                                                  getSourceCircuitBreakerConnectionId(index),
                                                  destinationCircuitBreakerConnectionId,
                                                  [&]() noexcept -> HRESULT
                {
#ifdef ENABLE_TESTS
                    MaybeCreateNativeMoveDestinationDirectoryRaceForSelfTest(destinationItemText);
#endif
                    FileSystemOptions options{};
                    InitializeFileSystemOptions(options, static_cast<void*>(&cookie));
                    options.moveMode                 = FILESYSTEM_MOVE_NATIVE_ONLY;
                    result.providerMutationAttempted = true;
                    return _fileSystem->MoveItem(sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                });
                result.verification = FileOperations::VerificationState::NotApplicable;
                const bool directoryMergeNonCommit =
                    result.hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) || result.hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
                if (! directoryMergeNonCommit || ! fileSystemIo ||
                    ! IsRegularDirectoryObject(*_fileSystem, *fileSystemIo, transferPlan->sourceEndpoint.profileId, sourceText) ||
                    ! IsRegularDirectoryObject(*_fileSystem, *fileSystemIo, transferPlan->destinationEndpoint.profileId, destinationItemText))
                {
                    return result;
                }

                // The provider refused the one-mutation rename because a regular directory already
                // occupies the destination. The same item continues as a rename merge: every child
                // relocates through the provider's Native rename and nothing is copied.
                Debug::Perf::EmitCounter(L"fileops.execute.native_directory_race_requalified_after_noncommit");
                result.hr = RunWithCircuitBreaker(breaker,
                                                  getSourceCircuitBreakerConnectionId(index),
                                                  destinationCircuitBreakerConnectionId,
                                                  [&]() noexcept -> HRESULT
                {
                    CrossFileSystemBridge mergeBridge(*this,
                                                      *_fileSystem,
                                                      *_fileSystem,
                                                      *fileSystemIo,
                                                      *fileSystemIo,
                                                      fileSystemDirOps.get(),
                                                      bridgeSourceMaxConcurrencyBudget,
                                                      bridgeDestinationMaxConcurrencyBudget,
                                                      itemFlags,
                                                      static_cast<void*>(&cookie),
                                                      0u,
                                                      sourceText.c_str(),
                                                      destinationItemText.c_str(),
                                                      FILE_ATTRIBUTE_DIRECTORY,
                                                      ReparsePointPolicy::Preserve,
                                                      transferPlan->sourceEndpoint.profileId,
                                                      transferPlan->destinationEndpoint.profileId,
                                                      transferPlan->sourceEndpoint.pathIdentity.value(),
                                                      transferPlan->destinationEndpoint.pathIdentity.value(),
                                                      true,
                                                      false,
                                                      false,
                                                      false,
                                                      false,
                                                      true);
                    const HRESULT mergeHr              = mergeBridge.RenameMergeDirectory(sourceText, destinationItemText);
                    const bool anyChildRelocated       = mergeBridge.anyDestinationPublished.load(std::memory_order_acquire);
                    result.managedSourceRetained       = mergeBridge.managedSourceRetained.load(std::memory_order_acquire);
                    result.managedCleanupIndeterminate = mergeBridge.managedCleanupIndeterminate.load(std::memory_order_acquire);
                    result.bridgePublication = anyChildRelocated ? FileOperations::PublicationState::Published : FileOperations::PublicationState::NotPublished;
                    result.failurePhase      = mergeBridge.failurePhase.load(std::memory_order_acquire);
                    result.failureStatus     = mergeBridge.failureStatus.load(std::memory_order_acquire);
                    result.mutationAttemptFlags = mergeBridge.mutationAttemptFlags.load(std::memory_order_acquire);
                    if (FAILED(mergeHr) && result.failurePhase == QualifiedItemFailurePhase::None)
                    {
                        result.failurePhase  = QualifiedItemFailurePhase::ProviderNative;
                        result.failureStatus = mergeHr;
                    }
                    // The per-child provider receipts under this cookie described children. The merge
                    // supplies the selected root's truth: known, committed once any child relocated,
                    // and the original present unless its emptied directory was removed.
                    storeSourceMutationTruth(
                        index,
                        mergeHr,
                        FileSystemItemMutationResult{.sizeBytes            = sizeof(FileSystemItemMutationResult),
                                                     .outcomeKnown         = TRUE,
                                                     .mutationCommitted    = anyChildRelocated ? TRUE : FALSE,
                                                     .originalStillPresent = mergeBridge.rootSourceRemoved.load(std::memory_order_acquire) ? FALSE : TRUE});
                    return mergeHr;
                });
                return result;
            }

            if (_operation == FILESYSTEM_DELETE)
            {
                if (index >= deletePlanBySourceIndex.size() || deletePlanBySourceIndex[index] == nullptr)
                {
                    result.hr = E_UNEXPECTED;
                    return result;
                }
                if (deletePlanBySourceIndex[index]->mode == FileOperations::DeleteMode::Permanent)
                {
                    result.hr = RunWithCircuitBreaker(breaker, getSourceCircuitBreakerConnectionId(index), {}, [&]() noexcept -> HRESULT {
                        return executePermanentDelete(index, &cookie, sourceText, result.providerMutationAttempted);
                    });
                    return result;
                }
                result.hr = RunWithCircuitBreaker(breaker,
                                                  getSourceCircuitBreakerConnectionId(index),
                                                  {},
                                                  [&]() noexcept -> HRESULT
                {
                    switch (GuardLiveOutputBeforeInvalidation(sourceText, FileOperations::MutationInterlockAccess::WriteSource))
                    {
                        case LiveOutputGuardDisposition::Skip: cookie.explicitSkipObserved.store(true, std::memory_order_release); return S_FALSE;
                        case LiveOutputGuardDisposition::Cancel: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        case LiveOutputGuardDisposition::RetryCurrentMutation:
                        case LiveOutputGuardDisposition::Proceed: break;
                    }
                    FileSystemOptions options{};
                    InitializeFileSystemOptions(options, static_cast<void*>(&cookie));
                    result.providerMutationAttempted = true;
                    return _fileSystem->DeleteItem(sourceText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                });
            }

            return result;
        };

        // One item policy owns mutation retry, conflict decisions, terminal classification, and
        // result publication. Scheduling may remain serial on the task worker or use the shared
        // scheduler, but those modes must not carry separate safety behavior.
        {
            // Per-task multi-item concurrency: run multiple CopyItem/MoveItem/DeleteItem calls concurrently while keeping
            // conflict prompts serialized (one prompt per task at a time).
            std::atomic<bool> hadSkipped{false};
            std::atomic<bool> hadCopyOnly{false};
            std::atomic<bool> hadPartialFailure{false};
            std::atomic<HRESULT> firstFailure{S_OK};

            struct QualifiedItemPolicyContext final
            {
                QualifiedItemPolicyContext(const bool& useResolvedItemsValue,
                                           const bool& isTransferOperationValue,
                                           std::vector<const FileOperations::TransferPlan*>& transferPlans,
                                           decltype(resolveTransferDestination)& resolveDestination,
                                           decltype(prepareExactTransferGuard)& prepareTransferGuard,
                                           const bool& crossFileSystemBridge,
                                           decltype(destinationFileSystemIo)& destinationIo,
                                           decltype(fileSystemIo)& sourceIo,
                                           decltype(ensureResolvedDestinationParent)& ensureDestinationParent,
                                           decltype(revalidateExactTransferGuard)& revalidateTransferGuard,
                                           decltype(executeQualifiedItemMutation)& executeMutation,
                                           std::atomic<bool>& copyOnly,
                                           std::atomic<bool>& skipped,
                                           std::atomic<bool>& partialFailure,
                                           decltype(traversalLimitReachedForCall)& traversalLimitReached,
                                           decltype(classifyQualifiedItemTerminal)& classifyTerminal,
                                           const bool& shouldContinueOnError,
                                           std::vector<const FileOperations::DeletePlan*>& deletePlans,
                                           decltype(getPerItemInFlightAggregate)& getInFlightAggregate,
                                           decltype(executeRecycleEscalation)& recycleEscalation,
                                           decltype(storeQualifiedItemResult)& storeItemResult,
                                           decltype(failurePhaseText)& getFailurePhaseText,
                                           decltype(stageDispositionText)& getStageDispositionText,
                                           unsigned int cachedModifierAttemptLimit) noexcept
                    : useResolvedItems(useResolvedItemsValue),
                      isTransferOperation(isTransferOperationValue),
                      transferPlanBySourceIndex(transferPlans),
                      resolveTransferDestination(resolveDestination),
                      prepareExactTransferGuard(prepareTransferGuard),
                      useCrossFileSystemBridge(crossFileSystemBridge),
                      destinationFileSystemIo(destinationIo),
                      fileSystemIo(sourceIo),
                      ensureResolvedDestinationParent(ensureDestinationParent),
                      revalidateExactTransferGuard(revalidateTransferGuard),
                      executeQualifiedItemMutation(executeMutation),
                      hadCopyOnly(copyOnly),
                      hadSkipped(skipped),
                      hadPartialFailure(partialFailure),
                      traversalLimitReachedForCall(traversalLimitReached),
                      classifyQualifiedItemTerminal(classifyTerminal),
                      continueOnError(shouldContinueOnError),
                      deletePlanBySourceIndex(deletePlans),
                      getPerItemInFlightAggregate(getInFlightAggregate),
                      executeRecycleEscalation(recycleEscalation),
                      storeQualifiedItemResult(storeItemResult),
                      failurePhaseText(getFailurePhaseText),
                      stageDispositionText(getStageDispositionText),
                      maxCachedModifierAttemptsPerBucket(cachedModifierAttemptLimit)
                {
                }

                QualifiedItemPolicyContext(const QualifiedItemPolicyContext&)            = delete;
                QualifiedItemPolicyContext(QualifiedItemPolicyContext&&)                 = delete;
                QualifiedItemPolicyContext& operator=(const QualifiedItemPolicyContext&) = delete;
                QualifiedItemPolicyContext& operator=(QualifiedItemPolicyContext&&)      = delete;

                const bool& useResolvedItems;
                const bool& isTransferOperation;
                std::vector<const FileOperations::TransferPlan*>& transferPlanBySourceIndex;
                decltype(resolveTransferDestination)& resolveTransferDestination;
                decltype(prepareExactTransferGuard)& prepareExactTransferGuard;
                const bool& useCrossFileSystemBridge;
                decltype(destinationFileSystemIo)& destinationFileSystemIo;
                decltype(fileSystemIo)& fileSystemIo;
                decltype(ensureResolvedDestinationParent)& ensureResolvedDestinationParent;
                decltype(revalidateExactTransferGuard)& revalidateExactTransferGuard;
                decltype(executeQualifiedItemMutation)& executeQualifiedItemMutation;
                std::atomic<bool>& hadCopyOnly;
                std::atomic<bool>& hadSkipped;
                std::atomic<bool>& hadPartialFailure;
                decltype(traversalLimitReachedForCall)& traversalLimitReachedForCall;
                decltype(classifyQualifiedItemTerminal)& classifyQualifiedItemTerminal;
                const bool& continueOnError;
                std::vector<const FileOperations::DeletePlan*>& deletePlanBySourceIndex;
                decltype(getPerItemInFlightAggregate)& getPerItemInFlightAggregate;
                decltype(executeRecycleEscalation)& executeRecycleEscalation;
                decltype(storeQualifiedItemResult)& storeQualifiedItemResult;
                decltype(failurePhaseText)& failurePhaseText;
                decltype(stageDispositionText)& stageDispositionText;
                unsigned int maxCachedModifierAttemptsPerBucket;
            };

            class QualifiedItemPolicy final : public PerItemExecutionPolicy
            {
            public:
                QualifiedItemPolicy(Task& owner, QualifiedItemPolicyContext& policyContext) noexcept : task(owner), context(policyContext)
                {
                }

                QualifiedItemPolicy(const QualifiedItemPolicy&)            = delete;
                QualifiedItemPolicy& operator=(const QualifiedItemPolicy&) = delete;

                [[nodiscard]] HRESULT Process(size_t index) noexcept override
                {
                    auto& _bridgeTraversalLimitHitCount = task._bridgeTraversalLimitHitCount;
                    auto& _cancelled                    = task._cancelled;
                    auto& _flags                        = task._flags;
                    auto& _operation                    = task._operation;
                    auto& _perItemCompletedBytes        = task._perItemCompletedBytes;
                    auto& _perItemCompletedEntryCount   = task._perItemCompletedEntryCount;
                    auto& _perItemCompletedItems        = task._perItemCompletedItems;
                    auto& _perItemTotalEntryCount       = task._perItemTotalEntryCount;
                    auto& _progressCompletedBytes       = task._progressCompletedBytes;
                    auto& _progressCompletedItems       = task._progressCompletedItems;
                    auto& _progressMutex                = task._progressMutex;
                    auto& _progressTotalItems           = task._progressTotalItems;
                    auto& _resolvedItems                = task._resolvedItems;
                    auto& _sourcePathAttributesHint     = task._sourcePathAttributesHint;
                    auto& _sourcePaths                  = task._sourcePaths;
                    auto& _stopToken                    = task._stopToken;

                    auto& useResolvedItems                                 = context.useResolvedItems;
                    auto& isTransferOperation                              = context.isTransferOperation;
                    auto& transferPlanBySourceIndex                        = context.transferPlanBySourceIndex;
                    auto& resolveTransferDestination                       = context.resolveTransferDestination;
                    auto& prepareExactTransferGuard                        = context.prepareExactTransferGuard;
                    auto& useCrossFileSystemBridge                         = context.useCrossFileSystemBridge;
                    auto& destinationFileSystemIo                          = context.destinationFileSystemIo;
                    auto& fileSystemIo                                     = context.fileSystemIo;
                    auto& ensureResolvedDestinationParent                  = context.ensureResolvedDestinationParent;
                    auto& revalidateExactTransferGuard                     = context.revalidateExactTransferGuard;
                    auto& executeQualifiedItemMutation                     = context.executeQualifiedItemMutation;
                    auto& hadCopyOnly                                      = context.hadCopyOnly;
                    auto& hadSkipped                                       = context.hadSkipped;
                    auto& hadPartialFailure                                = context.hadPartialFailure;
                    auto& traversalLimitReachedForCall                     = context.traversalLimitReachedForCall;
                    auto& classifyQualifiedItemTerminal                    = context.classifyQualifiedItemTerminal;
                    auto& continueOnError                                  = context.continueOnError;
                    auto& deletePlanBySourceIndex                          = context.deletePlanBySourceIndex;
                    auto& getPerItemInFlightAggregate                      = context.getPerItemInFlightAggregate;
                    auto& executeRecycleEscalation                         = context.executeRecycleEscalation;
                    auto& storeQualifiedItemResult                         = context.storeQualifiedItemResult;
                    auto& failurePhaseText                                 = context.failurePhaseText;
                    auto& stageDispositionText                             = context.stageDispositionText;
                    const unsigned int kMaxCachedModifierAttemptsPerBucket = context.maxCachedModifierAttemptsPerBucket;
                    const auto LogDiagnostic = [this](auto&&... args) noexcept { task.LogDiagnostic(std::forward<decltype(args)>(args)...); };

                    const std::wstring& sourceText                   = _sourcePaths[index].native();
                    const FileOperations::TransferPlan* transferPlan = nullptr;
                    FileOperations::OperationStrategy itemStrategy   = FileOperations::OperationStrategy::Copy;

                    std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> retryCounts{};
                    std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> cachedModifierAttempts{};
                    FileSystemFlags itemFlags = useResolvedItems ? _resolvedItems[index].flags : _flags;

                    bool itemSucceeded        = false;
                    bool itemSkipped          = false;
                    bool itemPartiallySkipped = false;
                    bool itemPartiallyFailed  = false;
                    std::wstring keepBothDestinationOverride;
                    std::wstring finalDestinationPath;
                    HRESULT finalItemStatus                             = E_PENDING;
                    bool finalManagedSourceRetained                     = false;
                    bool finalCleanupIndeterminate                      = false;
                    FileOperations::VerificationState finalVerification = FileOperations::VerificationState::NotRequested;
                    bool finalSelectedRootVerificationNotApplicable     = false;
                    std::optional<FileOperations::PublicationState> finalPartialPublication;
                    std::optional<FileOperations::ProviderIdentitySnapshot> finalRetainedSourceIdentity;
                    QualifiedItemFailurePhase finalFailurePhase      = QualifiedItemFailurePhase::None;
                    HRESULT finalFailureStatus                       = S_OK;
                    uint32_t finalMutationAttemptFlags               = 0u;
                    OwnedStageDisposition finalOwnedStageDisposition = OwnedStageDisposition::NotCreated;
                    size_t keepBothNextOrdinal                       = 2u;
                    bool bridgeUnsupportedReparse                    = false;
                    uint64_t callCompletedBytes                      = 0;
                    uint64_t callCompletedItems                      = 0;
                    uint64_t callTotalItems                          = 0;
                    uint64_t callDiscoveredBytes                     = 0;

                    const HRESULT processHr = [&]() noexcept -> HRESULT
                    {
                        if (sourceText.empty())
                        {
                            return E_INVALIDARG;
                        }
                        const auto closeItemDiscovery = wil::scope_exit([&]() noexcept { task.MarkDiscoveryItemClosed(index); });
                        if (isTransferOperation)
                        {
                            if (index >= transferPlanBySourceIndex.size())
                            {
                                return E_UNEXPECTED;
                            }
                            transferPlan = transferPlanBySourceIndex[index];
                        }
                        itemStrategy = transferPlan != nullptr ? transferPlan->strategy : FileOperations::OperationStrategy::Copy;

                        for (;;)
                        {
                            task.WaitWhilePaused();
                            if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                            {
                                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                            }

                            std::wstring destinationItemText;
                            if (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE)
                            {
                                if (! keepBothDestinationOverride.empty())
                                {
                                    destinationItemText = keepBothDestinationOverride;
                                }
                                else
                                {
                                    const HRESULT destinationHr = resolveTransferDestination(index, transferPlan, destinationItemText);
                                    if (FAILED(destinationHr))
                                    {
                                        finalFailurePhase  = QualifiedItemFailurePhase::DestinationParent;
                                        finalFailureStatus = destinationHr;
                                        return destinationHr;
                                    }
                                }
                            }
                            const bool isDirectoryShell =
                                useResolvedItems && _resolvedItems[index].kind == FolderWindow::ResolvedFileOperationItemKind::DirectoryShell;
                            FileOperations::TransferMutationGuard transferGuard{};
                            bool samePathCopy = false;
                            const HRESULT guardHr =
                                isTransferOperation ? prepareExactTransferGuard(index, sourceText, destinationItemText, transferGuard, samePathCopy) : S_FALSE;
                            if (FAILED(guardHr))
                            {
                                finalFailurePhase  = QualifiedItemFailurePhase::ProviderNative;
                                finalFailureStatus = guardHr;
                                return guardHr;
                            }
                            finalRetainedSourceIdentity =
                                guardHr == S_OK ? std::optional<FileOperations::ProviderIdentitySnapshot>(transferGuard.source.identity) : std::nullopt;
                            if (samePathCopy)
                            {
                                IFileSystemIO* const targetIo = useCrossFileSystemBridge ? destinationFileSystemIo.get() : fileSystemIo.get();
                                const bool sourceIsDirectory  = isDirectoryShell || (index < _sourcePathAttributesHint.size() &&
                                                                                     (_sourcePathAttributesHint[index] & FILE_ATTRIBUTE_DIRECTORY) != 0u);
                                const HRESULT keepBothHr      = FindAvailableUniqueSiblingPath(
                                    targetIo, destinationItemText, sourceIsDirectory, keepBothNextOrdinal, keepBothDestinationOverride);
                                if (FAILED(keepBothHr))
                                {
                                    return keepBothHr;
                                }
                                continue;
                            }
                            HRESULT hrEnsureParent = S_OK;
                            if (isTransferOperation)
                            {
                                if (transferPlan == nullptr || ! transferPlan->destinationEndpoint.pathIdentity.has_value())
                                {
                                    return E_UNEXPECTED;
                                }
                                hrEnsureParent = ensureResolvedDestinationParent(
                                    sourceText, destinationItemText, transferPlan->destinationEndpoint.pathIdentity.value(), isDirectoryShell, itemFlags);
                            }
                            if (FAILED(hrEnsureParent))
                            {
                                finalFailurePhase  = QualifiedItemFailurePhase::DestinationParent;
                                finalFailureStatus = hrEnsureParent;
                                return hrEnsureParent;
                            }
                            if (guardHr == S_OK)
                            {
                                const HRESULT revalidateHr =
                                    revalidateExactTransferGuard(index, sourceText, destinationItemText, transferGuard, isDirectoryShell);
                                if (FAILED(revalidateHr))
                                {
                                    finalFailurePhase  = QualifiedItemFailurePhase::ProviderNative;
                                    finalFailureStatus = revalidateHr;
                                    return revalidateHr;
                                }
                            }
                            PerItemCallbackCookie cookie{index};
                            cookie.operationDestinationPath = destinationItemText;
                            cookie.sourceIsDirectory        = isDirectoryShell || (index < _sourcePathAttributesHint.size() &&
                                                                                   (_sourcePathAttributesHint[index] & FILE_ATTRIBUTE_DIRECTORY) != 0u);

                            const bool copyOnlyMove =
                                _operation == FILESYSTEM_MOVE && transferPlan != nullptr && itemStrategy == FileOperations::OperationStrategy::CopyOnly;
                            const bool managedMove =
                                _operation == FILESYSTEM_MOVE && transferPlan != nullptr && itemStrategy == FileOperations::OperationStrategy::Managed;

                            const PerItemInFlightAggregate inFlightAggregate = BeginPerItemInFlightCall(task, &cookie, GetTickCount64());

                            {
                                std::scoped_lock lock(_progressMutex);
                                _progressCompletedItems = (std::max)(_progressCompletedItems, _perItemCompletedItems);
                                const uint64_t mapped   = _perItemCompletedBytes + inFlightAggregate.completedBytes;
                                _progressCompletedBytes = (std::max)(_progressCompletedBytes, mapped);
                                PublishProgressCountersLocked(task);
                            }

                            callCompletedBytes                               = 0;
                            callCompletedItems                               = 0;
                            callTotalItems                                   = 0;
                            bridgeUnsupportedReparse                         = false;
                            const uint64_t traversalLimitGenerationBefore    = _bridgeTraversalLimitHitCount.load(std::memory_order_acquire);
                            const QualifiedItemMutationResult mutationResult = executeQualifiedItemMutation(
                                index, transferPlan, itemStrategy, sourceText, destinationItemText, itemFlags, isDirectoryShell, cookie);
                            HRESULT itemHr             = mutationResult.hr;
                            bridgeUnsupportedReparse   = mutationResult.bridgeUnsupportedReparse;
                            finalDestinationPath       = destinationItemText;
                            finalManagedSourceRetained = mutationResult.managedSourceRetained;
                            finalCleanupIndeterminate  = mutationResult.managedCleanupIndeterminate || mutationResult.ownedStageCleanupIndeterminate;
                            finalVerification          = mutationResult.verification;
                            finalSelectedRootVerificationNotApplicable = mutationResult.verificationNotApplicableForSelectedRoot;
                            finalPartialPublication                    = mutationResult.bridgePublication;
                            finalFailurePhase                          = mutationResult.failurePhase;
                            finalFailureStatus                         = mutationResult.failureStatus;
                            finalMutationAttemptFlags                  = mutationResult.mutationAttemptFlags;
                            finalOwnedStageDisposition                 = mutationResult.ownedStageDisposition;

                            if (copyOnlyMove && itemHr == S_OK)
                            {
                                LogDiagnostic(
                                    DiagnosticSeverity::Warning,
                                    S_FALSE,
                                    L"move.copyOnly.sourceKept",
                                    L"Copied; source kept. The requested Move used the proven Copy route because no destructive Move strategy was qualified.",
                                    sourceText,
                                    destinationItemText);
                                hadCopyOnly.store(true, std::memory_order_release);
                                itemHr = S_FALSE;
                            }
                            else if (managedMove && mutationResult.managedSourceRetained && SUCCEEDED(itemHr))
                            {
                                LogDiagnostic(DiagnosticSeverity::Warning,
                                              S_FALSE,
                                              L"move.managed.sourceKept",
                                              L"Copied; source kept because exact managed-Move cleanup was unavailable or did not commit.",
                                              sourceText,
                                              destinationItemText);
                                hadCopyOnly.store(true, std::memory_order_release);
                                itemHr = S_FALSE;
                            }
                            else if (itemStrategy == FileOperations::OperationStrategy::Native && mutationResult.managedSourceRetained && SUCCEEDED(itemHr))
                            {
                                LogDiagnostic(DiagnosticSeverity::Warning,
                                              S_FALSE,
                                              L"move.renameMerge.sourceFolderKept",
                                              L"Moved; source folder kept because the folder gained or retained a child after its children were relocated.",
                                              sourceText,
                                              destinationItemText);
                                hadCopyOnly.store(true, std::memory_order_release);
                                itemHr = S_FALSE;
                            }
                            if (cookie.explicitSkipObserved.load(std::memory_order_acquire) && SUCCEEDED(itemHr))
                            {
                                // Recursive providers/bridge walkers may consume a child Skip and finish
                                // the selected root successfully. Preserve that intentional partial truth
                                // on the top-level typed item instead of letting the aggregate become S_OK.
                                itemHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                            }
                            finalItemStatus = itemHr;

                            const PerItemInFlightFinishResult finishedCall = FinishPerItemInFlightCall(task, &cookie);
                            callCompletedItems                             = finishedCall.completedItems;
                            callCompletedBytes                             = finishedCall.completedBytes;
                            callTotalItems                                 = finishedCall.totalItems;
                            callDiscoveredBytes                            = cookie.lastDiscoveredBytes;

                            if (FAILED(mutationResult.hr) && mutationResult.providerMutationAttempted)
                            {
                                std::optional<FileSystemItemMutationResult> receipt;
                                {
                                    std::scoped_lock lock(task._sourceItemStatusMutex);
                                    if (index < task._sourceItemResultBuilders.size())
                                    {
                                        const auto& builder = task._sourceItemResultBuilders[index];
                                        if (builder.status == mutationResult.hr)
                                        {
                                            receipt = builder.mutation;
                                        }
                                    }
                                }
                                const auto classification =
                                    FileSystemRouteContract::ClassifyFailedMutation(true, mutationResult.hr, receipt.has_value() ? &receipt.value() : nullptr);
                                // The IFileSystem directory-merge contract uses PARTIAL_COPY for
                                // consumed child Skips. That is terminal partial work, not permission
                                // to replay the root. An explicit uncertain/invalid receipt still wins.
                                const bool handledPartial =
                                    mutationResult.hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                                    cookie.explicitSkipObserved.load(std::memory_order_acquire) &&
                                    (! receipt.has_value() || classification == FileSystemRouteContract::MutationClassification::RetryableNoCommit ||
                                     classification == FileSystemRouteContract::MutationClassification::FailedKnown);
                                if (! handledPartial && task.IsCancellationOutcome(mutationResult.hr) &&
                                    (! receipt.has_value() || classification == FileSystemRouteContract::MutationClassification::RetryableNoCommit))
                                {
                                    // Cancel is a user/host decision. A proved no-commit receipt keeps
                                    // the item Canceled (no Retry, Keep Both, escalation, or replay).
                                    // No receipt means the provider proved nothing, so publication is
                                    // Unknown rather than NotAttempted. An explicit unknown/committed
                                    // receipt still takes the indeterminate or FailedKnown path below.
                                    if (! receipt.has_value() && ! finalPartialPublication.has_value())
                                    {
                                        finalPartialPublication = FileOperations::PublicationState::Unknown;
                                    }
                                    return itemHr;
                                }
                                if (! handledPartial && (classification == FileSystemRouteContract::MutationClassification::Indeterminate ||
                                                         classification == FileSystemRouteContract::MutationClassification::ContractViolation))
                                {
                                    // Neither a transport error nor Cancel proves non-commit. Stop before
                                    // Keep Both, native requalification, escalation, or a replay/Skip prompt.
                                    // Keep independent known axes (e.g. final not published, stage
                                    // cleanup unknown). Only invalid/mismatched evidence is discarded.
                                    {
                                        std::scoped_lock lock(task._sourceItemStatusMutex);
                                        if (index < task._sourceItemResultBuilders.size())
                                        {
                                            task._sourceItemResultBuilders[index].mutation =
                                                classification == FileSystemRouteContract::MutationClassification::ContractViolation ? std::nullopt : receipt;
                                        }
                                    }
                                    finalItemStatus    = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
                                    finalFailurePhase  = QualifiedItemFailurePhase::ProviderNative;
                                    finalFailureStatus = itemHr;
                                    LogDiagnostic(DiagnosticSeverity::Error,
                                                  itemHr,
                                                  L"item.mutation.indeterminate",
                                                  LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN),
                                                  sourceText,
                                                  destinationItemText);
                                    return finalItemStatus;
                                }
                                if (! handledPartial && classification != FileSystemRouteContract::MutationClassification::RetryableNoCommit)
                                {
                                    // A committed attempt or known retained stage is not a fresh primary retry.
                                    return itemHr;
                                }
                            }

                            if (cookie.keepBothRequested)
                            {
                                IFileSystemIO* const targetIo = useCrossFileSystemBridge ? destinationFileSystemIo.get() : fileSystemIo.get();
                                const HRESULT keepBothHr      = FindAvailableUniqueSiblingPath(
                                    targetIo, cookie.operationDestinationPath, cookie.sourceIsDirectory, keepBothNextOrdinal, keepBothDestinationOverride);
                                if (FAILED(keepBothHr))
                                {
                                    return keepBothHr;
                                }
                                continue;
                            }

                            {
                                std::scoped_lock lock(_progressMutex);
                                if (_operation == FILESYSTEM_DELETE)
                                {
                                    if (callCompletedItems > 0)
                                    {
                                        if (_perItemCompletedEntryCount > std::numeric_limits<uint64_t>::max() - callCompletedItems)
                                        {
                                            _perItemCompletedEntryCount = std::numeric_limits<uint64_t>::max();
                                        }
                                        else
                                        {
                                            _perItemCompletedEntryCount += callCompletedItems;
                                        }
                                    }

                                    if (callTotalItems > 0)
                                    {
                                        if (_perItemTotalEntryCount > std::numeric_limits<uint64_t>::max() - callTotalItems)
                                        {
                                            _perItemTotalEntryCount = std::numeric_limits<uint64_t>::max();
                                        }
                                        else
                                        {
                                            _perItemTotalEntryCount += callTotalItems;
                                        }
                                    }

                                    const uint64_t mappedCompletedItems = _perItemCompletedEntryCount + finishedCall.aggregate.completedItems;
                                    const uint64_t clampedCompleted =
                                        std::min<uint64_t>(mappedCompletedItems, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()));
                                    _progressCompletedItems = (std::max)(_progressCompletedItems, static_cast<unsigned long>(clampedCompleted));

                                    const uint64_t mappedTotalItems = _perItemTotalEntryCount + finishedCall.aggregate.totalItems;
                                    if (mappedTotalItems > 0)
                                    {
                                        const uint64_t clampedTotal =
                                            std::min<uint64_t>(mappedTotalItems, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()));
                                        _progressTotalItems = (std::max)(_progressTotalItems, static_cast<unsigned long>(clampedTotal));
                                    }
                                }

                                const uint64_t mapped   = _perItemCompletedBytes + finishedCall.aggregate.completedBytes;
                                _progressCompletedBytes = (std::max)(_progressCompletedBytes, mapped);
                                PublishProgressCountersLocked(task);
                            }

                            const bool traversalLimitReached = traversalLimitReachedForCall(
                                itemHr, traversalLimitGenerationBefore, _bridgeTraversalLimitHitCount.load(std::memory_order_acquire));
                            const QualifiedItemTerminalDisposition terminal =
                                classifyQualifiedItemTerminal(itemHr,
                                                              traversalLimitReached,
                                                              mutationResult.managedCleanupIndeterminate || mutationResult.ownedStageCleanupIndeterminate,
                                                              mutationResult.verification,
                                                              cookie.explicitSkipObserved.load(std::memory_order_acquire));
                            if (terminal == QualifiedItemTerminalDisposition::Canceled)
                            {
                                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                            }

                            if (terminal == QualifiedItemTerminalDisposition::TraversalLimit)
                            {
                                // Quantitative traversal/resource ceilings stop discovery for the task.
                                // They are not conflict prompts and Continue-on-error cannot resume past them.
                                return itemHr;
                            }

                            if (terminal == QualifiedItemTerminalDisposition::Partial)
                            {
                                itemSucceeded        = true;
                                itemPartiallySkipped = true;
                                hadSkipped.store(true, std::memory_order_release);
                                break;
                            }
                            if (terminal == QualifiedItemTerminalDisposition::Skipped)
                            {
                                itemSkipped = true;
                                hadSkipped.store(true, std::memory_order_release);
                                break;
                            }
                            if (terminal == QualifiedItemTerminalDisposition::PartialFailure)
                            {
                                itemPartiallyFailed = true;
                                hadPartialFailure.store(true, std::memory_order_release);
                                break;
                            }

                            if (terminal == QualifiedItemTerminalDisposition::VerificationFailure)
                            {
                                LogDiagnostic(DiagnosticSeverity::Error,
                                              itemHr,
                                              L"verification.failed",
                                              LoadStringResource(nullptr, IDS_FILEOPS_RETAINED_VERIFICATION_FAILED),
                                              sourceText,
                                              destinationItemText);
                                if (continueOnError)
                                {
                                    itemPartiallyFailed = true;
                                    hadPartialFailure.store(true, std::memory_order_release);
                                    break;
                                }
                                return itemHr;
                            }

                            if (terminal == QualifiedItemTerminalDisposition::CleanupIndeterminate)
                            {
                                const bool ownedStageCleanupUnknown = mutationResult.ownedStageCleanupIndeterminate;
                                LogDiagnostic(
                                    DiagnosticSeverity::Error,
                                    itemHr,
                                    ownedStageCleanupUnknown ? L"bridge.stage.cleanupIndeterminate" : L"move.managed.cleanupIndeterminate",
                                    ownedStageCleanupUnknown
                                        ? LoadStringResource(nullptr, IDS_FILEOPS_OWNED_STAGE_CLEANUP_UNKNOWN)
                                        : L"Destination published; exact source cleanup is indeterminate and the item will not be recopied or retried.",
                                    sourceText,
                                    destinationItemText);
                                if (continueOnError)
                                {
                                    itemPartiallyFailed = true;
                                    hadPartialFailure.store(true, std::memory_order_release);
                                    break;
                                }
                                return itemHr;
                            }

                            if (terminal == QualifiedItemTerminalDisposition::Success)
                            {
                                if (useCrossFileSystemBridge && mutationResult.bridgeRootDirectoryReparseSkipped)
                                {
                                    LogDiagnostic(DiagnosticSeverity::Warning,
                                                  HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                                                  L"bridge.reparse.skip",
                                                  L"Skipped root directory reparse point during bridge operation.",
                                                  sourceText,
                                                  destinationItemText);
                                    itemSkipped = true;
                                    hadSkipped.store(true, std::memory_order_release);
                                    break;
                                }

                                if (useCrossFileSystemBridge && mutationResult.bridgeSkippedDirectoryReparseCount > 0)
                                {
                                    const std::wstring skipMessage = std::format(L"Skipped {:L} directory reparse point{:s} during bridge operation.",
                                                                                 mutationResult.bridgeSkippedDirectoryReparseCount,
                                                                                 mutationResult.bridgeSkippedDirectoryReparseCount == 1ul ? L"" : L"s");
                                    LogDiagnostic(DiagnosticSeverity::Warning,
                                                  HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                                                  L"bridge.reparse.skip",
                                                  skipMessage,
                                                  sourceText,
                                                  destinationItemText);
                                    itemPartiallySkipped = true;
                                    hadSkipped.store(true, std::memory_order_release);
                                }

                                itemSucceeded = true;
                                break;
                            }

                            if (continueOnError)
                            {
                                auto [diagnosticSource, diagnosticDestination] =
                                    GetMostSpecificPathsForDiagnostics(task, &cookie, sourceText, destinationItemText);
                                LogDiagnostic(DiagnosticSeverity::Warning,
                                              itemHr,
                                              L"item.continueOnError",
                                              L"Item failed; continue-on-error preserved the failure and continued with remaining items.",
                                              diagnosticSource,
                                              diagnosticDestination);
                                itemPartiallyFailed = true;
                                hadPartialFailure.store(true, std::memory_order_release);
                                break;
                            }

                            const FileSystemOperation bucketOperation             = _operation;
                            const wil::com_ptr<IFileSystemIO>& bucketFileSystemIo = useCrossFileSystemBridge ? destinationFileSystemIo : fileSystemIo;
                            ConflictBucket bucket                                 = ClassifyConflictBucket(
                                bucketOperation, itemFlags, bucketFileSystemIo, itemHr, sourceText, destinationItemText, bridgeUnsupportedReparse);
                            if (bucket == ConflictBucket::RecycleFailed)
                            {
                                const HRESULT escalationHr = executeRecycleEscalation(index, &cookie, itemHr, sourceText);
                                if (escalationHr == S_OK)
                                {
                                    finalItemStatus    = S_OK;
                                    finalFailurePhase  = QualifiedItemFailurePhase::None;
                                    finalFailureStatus = S_OK;
                                    itemSucceeded      = true;
                                    break;
                                }
                                if (escalationHr == S_FALSE)
                                {
                                    finalItemStatus = S_FALSE;
                                    itemSkipped     = true;
                                    hadSkipped.store(true, std::memory_order_release);
                                    break;
                                }
                                return escalationHr;
                            }

                            const size_t provisionalBucketIndex = static_cast<size_t>(bucket);
                            // C2: only a proved no-commit reaches this prompt (the receipt classifier
                            // above stops committed and indeterminate attempts), so a retryable bucket
                            // stays retryable; each explicit Retry is one attempt, revalidated by the
                            // loop it re-enters, and the prompt says how many have failed so far.
                            const bool allowRetry           = IsRetryableConflictBucket(bucket);
                            const unsigned int attemptCount = provisionalBucketIndex < retryCounts.size() ? retryCounts[provisionalBucketIndex] : 0u;

                            ConflictPromptBeginResult promptBegin = BeginProviderReturnedConflictPrompt(
                                task, &cookie, bucket, itemHr, sourceText, destinationItemText, allowRetry, attemptCount, false);
                            bucket                   = promptBegin.bucket;
                            const size_t bucketIndex = static_cast<size_t>(bucket);
                            if (promptBegin.fromCache && IsModifierConflictAction(promptBegin.action) && bucketIndex < cachedModifierAttempts.size() &&
                                cachedModifierAttempts[bucketIndex] >= kMaxCachedModifierAttemptsPerBucket)
                            {
                                // Only this item falls back to a fresh prompt; the scoped cached decision
                                // stays valid for every other matching item in the task.
                                promptBegin = BeginProviderReturnedConflictPrompt(
                                    task, &cookie, bucket, itemHr, sourceText, destinationItemText, allowRetry, attemptCount, true);
                            }
                            ConflictAction action = promptBegin.action;
                            if (promptBegin.ownsPrompt)
                            {
                                action = WaitForConflictDecision(task, &cookie, promptBegin.decisionScope).first;
                            }

                            if (action == ConflictAction::Overwrite)
                            {
                                if (bucketIndex < cachedModifierAttempts.size())
                                {
                                    ++cachedModifierAttempts[bucketIndex];
                                }
                                LogDiagnostic(DiagnosticSeverity::Error,
                                              HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                              L"conflict.destructiveGrant.unconsumed",
                                              L"The provider returned a destination conflict without consuming the exact conflict authority.",
                                              sourceText,
                                              destinationItemText);
                                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                            }

                            if (action == ConflictAction::ReplaceReadOnly)
                            {
                                if (bucketIndex < cachedModifierAttempts.size())
                                {
                                    ++cachedModifierAttempts[bucketIndex];
                                }
                                LogDiagnostic(DiagnosticSeverity::Error,
                                              HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                              L"conflict.destructiveGrant.unconsumed",
                                              L"The provider returned a read-only conflict without consuming the exact conflict authority.",
                                              sourceText,
                                              destinationItemText);
                                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                            }

                            if (action == ConflictAction::ReplaceLink)
                            {
                                if (bucketIndex < cachedModifierAttempts.size())
                                {
                                    ++cachedModifierAttempts[bucketIndex];
                                }
                                LogDiagnostic(DiagnosticSeverity::Error,
                                              HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                              L"conflict.destructiveGrant.unconsumed",
                                              L"The provider returned a link conflict without consuming the exact conflict authority.",
                                              sourceText,
                                              destinationItemText);
                                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                            }

                            if (action == ConflictAction::KeepBoth)
                            {
                                IFileSystemIO* const targetIo = useCrossFileSystemBridge ? destinationFileSystemIo.get() : fileSystemIo.get();
                                const HRESULT keepBothHr      = FindAvailableUniqueSiblingPath(
                                    targetIo, destinationItemText, cookie.sourceIsDirectory, keepBothNextOrdinal, keepBothDestinationOverride);
                                if (FAILED(keepBothHr))
                                {
                                    return keepBothHr;
                                }
                                continue;
                            }

                            if (action == ConflictAction::Retry)
                            {
                                if (bucketIndex < retryCounts.size())
                                {
                                    ++retryCounts[bucketIndex];
                                }
                                if (bucket == ConflictBucket::SharingViolation)
                                {
                                    Sleep(750);
                                }
                                continue;
                            }

                            if (action == ConflictAction::Skip || action == ConflictAction::SkipAll)
                            {
                                auto [diagnosticSource, diagnosticDestination] =
                                    GetMostSpecificPathsForDiagnostics(task, &cookie, sourceText, destinationItemText);
                                LogDiagnostic(DiagnosticSeverity::Warning,
                                              itemHr,
                                              L"item.conflict.skip",
                                              L"Conflict action Skip item selected.",
                                              diagnosticSource,
                                              diagnosticDestination);
                                itemSkipped = true;
                                hadSkipped.store(true, std::memory_order_release);
                                break;
                            }

                            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        }

                        uint64_t bytesForItem = 0;
                        if (itemSucceeded || itemPartiallyFailed)
                        {
                            if (_operation == FILESYSTEM_DELETE && index < deletePlanBySourceIndex.size() && deletePlanBySourceIndex[index] != nullptr &&
                                deletePlanBySourceIndex[index]->mode == FileOperations::DeleteMode::Permanent)
                            {
                                // Exact bound deletion reports discovery through the operation-control ABI,
                                // not the legacy provider progress callback. Once the conditional mutation
                                // succeeds, those same-walk bytes become completed bytes for this item.
                                callCompletedBytes = (std::max)(callCompletedBytes, callDiscoveredBytes);
                            }
                            bytesForItem = callCompletedBytes;
                        }

                        if (itemSucceeded && _operation == FILESYSTEM_DELETE &&
                            NavigationLocation::EqualsNoCase(task._sourcePluginId, L"builtin/file-system") && index < deletePlanBySourceIndex.size() &&
                            deletePlanBySourceIndex[index] != nullptr && deletePlanBySourceIndex[index]->mode == FileOperations::DeleteMode::Permanent)
                        {
                            // Exact Local deletion has no legacy progress callback. The retained-object
                            // result is the first point where the host knows this selected-root mutation
                            // completed, and the per-item discovery scope is still open here.
                            task.NoteDiscoveryCompletionWhileOpen(PerfNowUs(), bytesForItem, 1u);
                        }

                        const PerItemInFlightAggregate inFlightAggregate = getPerItemInFlightAggregate();
                        StorePublishedTopLevelCompletionSnapshot(task, MarkTopLevelItemCompleted(task, index));

                        {
                            std::scoped_lock lock(_progressMutex);
                            if (itemSucceeded || itemPartiallyFailed)
                            {
                                if (_perItemCompletedBytes > std::numeric_limits<uint64_t>::max() - bytesForItem)
                                {
                                    return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                                }
                                _perItemCompletedBytes += bytesForItem;
                            }

                            if (_perItemCompletedItems < std::numeric_limits<unsigned long>::max())
                            {
                                ++_perItemCompletedItems;
                            }
                            _progressCompletedItems = (std::max)(_progressCompletedItems, _perItemCompletedItems);
                            const uint64_t mapped   = _perItemCompletedBytes + inFlightAggregate.completedBytes;
                            _progressCompletedBytes = (std::max)(_progressCompletedBytes, mapped);
                            PublishProgressCountersLocked(task);
                        }

                        return S_OK;
                    }();

                    const HRESULT typedStatus = FAILED(processHr) ? processHr : (finalItemStatus != E_PENDING ? finalItemStatus : E_UNEXPECTED);
                    if (finalFailurePhase == QualifiedItemFailurePhase::None && FAILED(typedStatus))
                    {
                        finalFailurePhase  = finalVerification == FileOperations::VerificationState::Failed ? QualifiedItemFailurePhase::Verification
                                                                                                            : QualifiedItemFailurePhase::ProviderNative;
                        finalFailureStatus = typedStatus;
                    }
                    storeQualifiedItemResult(index,
                                             itemStrategy,
                                             finalDestinationPath,
                                             itemSkipped || itemPartiallySkipped ? S_FALSE : typedStatus,
                                             itemSucceeded,
                                             itemSkipped,
                                             itemPartiallySkipped,
                                             itemPartiallyFailed,
                                             finalPartialPublication,
                                             finalManagedSourceRetained,
                                             finalCleanupIndeterminate,
                                             finalOwnedStageDisposition,
                                             finalVerification,
                                             finalRetainedSourceIdentity ? &finalRetainedSourceIdentity.value() : nullptr,
                                             finalSelectedRootVerificationNotApplicable);

                    const uint64_t publicationValue = finalPartialPublication.has_value()
                                                          ? static_cast<uint64_t>(finalPartialPublication.value())
                                                          : static_cast<uint64_t>(FileOperations::PublicationState::NotAttempted);
                    Debug::Perf::Emit(L"FileOps.Bridge.FailurePhase",
                                      failurePhaseText(finalFailurePhase),
                                      0u,
                                      finalMutationAttemptFlags,
                                      publicationValue,
                                      FAILED(finalFailureStatus) ? finalFailureStatus : typedStatus);
                    Debug::Perf::Emit(L"FileOps.Bridge.StageDisposition",
                                      stageDispositionText(finalOwnedStageDisposition),
                                      0u,
                                      finalMutationAttemptFlags,
                                      static_cast<uint64_t>(finalFailurePhase),
                                      FAILED(finalFailureStatus) ? finalFailureStatus : typedStatus);

                    return processHr;
                }

            private:
                Task& task;
                QualifiedItemPolicyContext& context;
            };

            QualifiedItemPolicyContext itemPolicyContext{useResolvedItems,
                                                         isTransferOperation,
                                                         transferPlanBySourceIndex,
                                                         resolveTransferDestination,
                                                         prepareExactTransferGuard,
                                                         useCrossFileSystemBridge,
                                                         destinationFileSystemIo,
                                                         fileSystemIo,
                                                         ensureResolvedDestinationParent,
                                                         revalidateExactTransferGuard,
                                                         executeQualifiedItemMutation,
                                                         hadCopyOnly,
                                                         hadSkipped,
                                                         hadPartialFailure,
                                                         traversalLimitReachedForCall,
                                                         classifyQualifiedItemTerminal,
                                                         continueOnError,
                                                         deletePlanBySourceIndex,
                                                         getPerItemInFlightAggregate,
                                                         executeRecycleEscalation,
                                                         storeQualifiedItemResult,
                                                         failurePhaseText,
                                                         stageDispositionText,
                                                         kMaxCachedModifierAttemptsPerBucket};
            QualifiedItemPolicy itemPolicy(*this, itemPolicyContext);

            if (_perItemMaxConcurrency > 1u)
            {
                auto& scheduler                     = GetPerItemTaskScheduler();
                const auto schedulerStart           = scheduler.CapturePerfSnapshot();
                const uint64_t schedulerWallStartUs = PerfNowUs();

                auto job = scheduler.StartJob(
                    this, _perItemMaxConcurrency, _sourcePaths.size(), itemPolicy, firstFailure, PerItemTaskScheduler::FailurePolicy::CancelTask);

                scheduler.WaitJob(job);

                const auto schedulerEnd = scheduler.CapturePerfSnapshot();
                _perf.schedulerWaitUs.fetch_add(PerfElapsedUs(schedulerWallStartUs), std::memory_order_relaxed);
                _perf.schedulerDequeueAttempts.fetch_add(
                    (schedulerEnd.dequeueAttempts >= schedulerStart.dequeueAttempts) ? (schedulerEnd.dequeueAttempts - schedulerStart.dequeueAttempts) : 0,
                    std::memory_order_relaxed);
                _perf.schedulerDequeueSuccess.fetch_add(
                    (schedulerEnd.dequeueSuccess >= schedulerStart.dequeueSuccess) ? (schedulerEnd.dequeueSuccess - schedulerStart.dequeueSuccess) : 0,
                    std::memory_order_relaxed);
                _perf.schedulerWaitForWorkUs.fetch_add(
                    (schedulerEnd.waitForWorkUs >= schedulerStart.waitForWorkUs) ? (schedulerEnd.waitForWorkUs - schedulerStart.waitForWorkUs) : 0,
                    std::memory_order_relaxed);
                _perf.schedulerProcessIndexUs.fetch_add(
                    (schedulerEnd.processIndexUs >= schedulerStart.processIndexUs) ? (schedulerEnd.processIndexUs - schedulerStart.processIndexUs) : 0,
                    std::memory_order_relaxed);
            }
            else
            {
                // Providers that require serialized/task-thread execution keep that thread
                // affinity; they invoke the same policy function without entering the scheduler.
                for (size_t index = 0; index < _sourcePaths.size(); ++index)
                {
                    const HRESULT itemHr = itemPolicy.Process(index);
                    if (FAILED(itemHr))
                    {
                        firstFailure.store(itemHr, std::memory_order_release);
                        break;
                    }
                }
            }

            ClearConflictPrompt(*this);

            const HRESULT hr = firstFailure.load(std::memory_order_acquire);
            if (FAILED(hr))
            {
                return hr;
            }

            if (hadPartialFailure.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }

            if (hadSkipped.load(std::memory_order_acquire) || _observedSkipAction.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }

            if (hadCopyOnly.load(std::memory_order_acquire))
            {
                return S_FALSE;
            }

            return S_OK;
        }
    }

    if (_operation != FILESYSTEM_DELETE)
    {
        return E_NOTIMPL;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** pathArray = nullptr;
    unsigned long count       = 0;
    HRESULT hr                = BuildPathArrayArena(_sourcePaths, arenaOwner, &pathArray, &count);
    if (FAILED(hr))
    {
        return hr;
    }

    if (count == 0)
    {
        return S_FALSE;
    }

    PerItemCallbackCookie cookie{};
    cookie.itemIndex = 0u;
    FileSystemOptions options{};
    InitializeFileSystemOptions(options, static_cast<void*>(&cookie));
    const HRESULT operationHr = _fileSystem->DeleteItems(pathArray, count, _flags, &options, this, static_cast<void*>(&cookie));
    CloseDiscovery();

    if ((_flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0)
    {
        std::vector<std::optional<HRESULT>> itemStatuses;
        {
            std::scoped_lock lock(_sourceItemStatusMutex);
            itemStatuses.reserve(_sourceItemResultBuilders.size());
            for (const SourceItemResultBuilder& builder : _sourceItemResultBuilders)
            {
                itemStatuses.push_back(builder.status);
            }
        }

        HRESULT firstEscalationFailure = S_OK;
        bool hadEscalationSkip         = false;
        for (size_t index = 0u; index < _sourcePaths.size(); ++index)
        {
            if (index >= itemStatuses.size() || ! itemStatuses[index].has_value())
            {
                continue;
            }
            const HRESULT itemStatus = itemStatuses[index].value();
            if (SUCCEEDED(itemStatus))
            {
                continue;
            }
            if (itemStatus == HRESULT_FROM_WIN32(ERROR_CANCELLED) || itemStatus == E_ABORT)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            const HRESULT escalationHr = executeRecycleEscalation(index, nullptr, itemStatus, _sourcePaths[index].native());
            if (escalationHr == S_FALSE)
            {
                hadEscalationSkip = true;
                continue;
            }
            if (FAILED(escalationHr) && SUCCEEDED(firstEscalationFailure))
            {
                firstEscalationFailure = escalationHr;
            }
        }
        if (FAILED(firstEscalationFailure))
        {
            return firstEscalationFailure;
        }
        if (hadEscalationSkip)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        bool everyItemSucceeded = _sourcePaths.size() <= itemStatuses.size();
        if (everyItemSucceeded)
        {
            std::scoped_lock lock(_sourceItemStatusMutex);
            for (size_t index = 0u; index < _sourcePaths.size(); ++index)
            {
                everyItemSucceeded = index < _sourceItemResultBuilders.size() && _sourceItemResultBuilders[index].status.has_value() &&
                                     SUCCEEDED(_sourceItemResultBuilders[index].status.value()) && _sourceItemResultBuilders[index].status.value() != S_FALSE;
                if (! everyItemSucceeded)
                {
                    break;
                }
            }
        }
        if (everyItemSucceeded)
        {
            return S_OK;
        }
    }

    if (operationHr == S_OK && _observedSkipAction.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }
    return operationHr;
}

void FolderWindow::FileOperationState::Task::LogDiagnostic(DiagnosticSeverity severity,
                                                           HRESULT status,
                                                           std::wstring_view category,
                                                           std::wstring_view message,
                                                           std::wstring_view sourcePath,
                                                           std::wstring_view destinationPath) noexcept
{
    if (! _state)
    {
        return;
    }

    std::wstring effectiveSource;
    std::wstring effectiveDestination;

    if (sourcePath.empty() || destinationPath.empty())
    {
        std::scoped_lock lock(_progressPathMutex);
        CopyEffectiveProgressPathsLocked(*this, effectiveSource, effectiveDestination);
    }

    if (! sourcePath.empty())
    {
        effectiveSource = std::wstring(sourcePath);
    }
    if (! destinationPath.empty())
    {
        effectiveDestination = std::wstring(destinationPath);
    }

    _state->RecordTaskDiagnostic(_taskId, _operation, severity, status, category, message, effectiveSource, effectiveDestination);
}

HRESULT FolderWindow::FileOperationState::Task::BuildPathArrayArena(const std::vector<std::filesystem::path>& paths,
                                                                    FileSystemArenaOwner& arenaOwner,
                                                                    const wchar_t*** outPaths,
                                                                    unsigned long* outCount) noexcept
{
    if (! outPaths || ! outCount)
    {
        return E_POINTER;
    }

    *outPaths = nullptr;
    *outCount = 0;

    if (paths.empty())
    {
        return S_OK;
    }

    const uint64_t count64 = static_cast<uint64_t>(paths.size());
    if (count64 > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const uint64_t arrayBytes64 = count64 * static_cast<uint64_t>(sizeof(const wchar_t*));
    if (arrayBytes64 > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    unsigned long totalBytes = static_cast<unsigned long>(arrayBytes64);

    for (const auto& path : paths)
    {
        const std::wstring& text = path.native();
        const size_t length      = text.size();
        if (length > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)) - 1u)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        const unsigned long bytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
        if (totalBytes > std::numeric_limits<unsigned long>::max() - bytes)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        totalBytes += bytes;
    }

    HRESULT hr = arenaOwner.Initialize(totalBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    FileSystemArena* arena = arenaOwner.Get();
    auto* array            = static_cast<const wchar_t**>(
        AllocateFromFileSystemArena(arena, static_cast<unsigned long>(arrayBytes64), static_cast<unsigned long>(alignof(const wchar_t*))));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (size_t index = 0; index < paths.size(); ++index)
    {
        const std::wstring& text  = paths[index].native();
        const size_t length       = text.size();
        const unsigned long bytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
        auto* buffer              = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, bytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! buffer)
        {
            return E_OUTOFMEMORY;
        }

        if (length > 0)
        {
            ::CopyMemory(buffer, text.data(), length * sizeof(wchar_t));
        }
        buffer[length] = L'\0';
        array[index]   = buffer;
    }

    *outPaths = array;
    *outCount = static_cast<unsigned long>(count64);
    return S_OK;
}
