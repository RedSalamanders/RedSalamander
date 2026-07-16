#include "FolderWindow.FileOperationsInternal.h"

#include "ConnectionProfileUtils.h"
#include "FileSystemPathIdentity.h"
#include "FolderWindow.FileOperations.IssuesPane.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "SessionState.h"
#include "SettingsHotReload.h"
#include "SettingsSave.h"
#include "SettingsStore.h"

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
#include <unordered_map>
#include <utility>

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

namespace
{
using Task = FolderWindow::FileOperationState::Task;

[[nodiscard]] HRESULT AdvanceValidatedFileInfoEntry(FileInfo* entry,
                                                    const std::byte* bufferBase,
                                                    const std::byte* bufferEnd,
                                                    FileInfo*& nextOut) noexcept;
[[nodiscard]] HRESULT TryGetValidatedFileInfoName(FileInfo* entry,
                                                  const std::byte* bufferBase,
                                                  const std::byte* bufferEnd,
                                                  std::wstring_view& nameOut) noexcept;

#ifdef ENABLE_TESTS
std::atomic<unsigned int> g_fileOpsBridgePipelineMode{static_cast<unsigned int>(FileOpsBridgePipelineMode::Default)};
std::atomic<unsigned int> g_fileOpsBridgeProducerDelayMs{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextFileCopyCount{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextFileCopyAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextSourceGetSizeCount{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextSourceGetSizeAttempts{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationGetSizeCount{0};
std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationGetSizeAttempts{0};
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
std::atomic<unsigned long> g_fileOpsBridgeMutateDestinationBeforeMoveCleanupAttempts{0};
std::atomic<bool> g_fileOpsPreCalcThreadStartFailure{false};
std::atomic<unsigned long> g_fileOpsPreCalcThreadStartAttempts{0};
std::atomic<bool> g_fileOpsAutoConcurrencyOverrideEnabled{false};
std::atomic<unsigned int> g_fileOpsAutoConcurrencyOverridePreferred{1};
std::atomic<uint32_t> g_fileOpsAutoConcurrencyOverrideStorageKind{FILESYSTEM_STORAGE_UNKNOWN};

struct SelfTestPausePoint final
{
    SelfTestPausePoint()                                     = default;
    SelfTestPausePoint(const SelfTestPausePoint&)            = delete;
    SelfTestPausePoint& operator=(const SelfTestPausePoint&) = delete;
    SelfTestPausePoint(SelfTestPausePoint&&)                 = delete;
    SelfTestPausePoint& operator=(SelfTestPausePoint&&)      = delete;

    std::atomic<bool> enabled{false};
    std::atomic<bool> entered{false};
    std::atomic<bool> releaseRequested{false};

    void Set(bool isEnabled) noexcept
    {
        if (isEnabled)
        {
            entered.store(false, std::memory_order_release);
            releaseRequested.store(false, std::memory_order_release);
            enabled.store(true, std::memory_order_release);
            return;
        }

        enabled.store(false, std::memory_order_release);
        releaseRequested.store(true, std::memory_order_release);
    }

    [[nodiscard]] bool HasEntered() const noexcept
    {
        return entered.load(std::memory_order_acquire);
    }

    void Release() noexcept
    {
        enabled.store(false, std::memory_order_release);
        releaseRequested.store(true, std::memory_order_release);
    }

    void Pause(ULONGLONG bailoutMs) noexcept
    {
        if (! enabled.load(std::memory_order_acquire))
        {
            return;
        }

        entered.store(true, std::memory_order_release);
        const auto clearEntered = wil::scope_exit([&]() noexcept { entered.store(false, std::memory_order_release); });
        const ULONGLONG startTick = GetTickCount64();
        while (! releaseRequested.load(std::memory_order_acquire) && enabled.load(std::memory_order_acquire))
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (nowTick >= startTick && (nowTick - startTick) > bailoutMs)
            {
                break;
            }
            Sleep(1);
        }
    }
};

SelfTestPausePoint g_fileOpsPostFinishedCompletionPausePoint;
SelfTestPausePoint g_fileOpsBridgeMoveSourceCleanupPausePoint;
SelfTestPausePoint g_fileOpsBridgeMoveManifestTakePausePoint;
SelfTestPausePoint g_fileOpsConflictMetadataPausePoint;
std::atomic<ULONGLONG> g_fileOpsConflictMetadataPauseBailoutMs{5'000ull};
std::atomic<uint64_t> g_fileOpsBridgeMoveManifestCurrentEntries{0};
std::atomic<uint64_t> g_fileOpsBridgeMoveManifestPeakEntries{0};

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

[[nodiscard]] bool ConsumeBridgeCounterForSelfTest(std::atomic<unsigned long>& remainingCount,
                                                   std::atomic<unsigned long>& attemptCount) noexcept
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

[[nodiscard]] HRESULT MaybeInjectHostileBridgeChildNamesForSelfTest(FileInfo* head,
                                                                    std::byte* bufferBase,
                                                                    std::byte* bufferEnd) noexcept
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

        const auto* entryBytes = reinterpret_cast<const std::byte*>(entry);
        const size_t available = static_cast<size_t>(bufferEnd - entryBytes);
        const size_t recordBytes = entry->NextEntryOffset != 0u ? static_cast<size_t>(entry->NextEntryOffset) : available;
        constexpr size_t kNameOffset = offsetof(FileInfo, FileName);
        const size_t hostileBytes = hostileName.size() * sizeof(wchar_t);
        if (recordBytes < kNameOffset || hostileBytes > recordBytes - kNameOffset)
        {
            return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
        }

        std::memcpy(entry->FileName, hostileName.data(), hostileBytes);
        entry->FileNameSize = static_cast<unsigned long>(hostileBytes);

        FileInfo* next = nullptr;
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

[[nodiscard]] HRESULT MaybeInjectBridgeFileReparseForSelfTest(FileInfo* head,
                                                              std::byte* bufferBase,
                                                              std::byte* bufferEnd) noexcept
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

struct SelfTestBridgeFileReader final : IFileReader
{
    SelfTestBridgeFileReader(wil::com_ptr<IFileReader> inner, SelfTestBridgeIoRole role) noexcept : _inner(std::move(inner)), _role(role)
    {
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
        return _inner->GetSize(sizeBytes);
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
    SelfTestBridgeIoRole _role = SelfTestBridgeIoRole::Source;
};

struct SelfTestBridgeFileWriter final : IFileWriter
{
    explicit SelfTestBridgeFileWriter(wil::com_ptr<IFileWriter> inner) noexcept : _inner(std::move(inner))
    {
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

        if (bytesToWrite > 0u &&
            ConsumeBridgeCounterForSelfTest(g_fileOpsBridgeUnderConsumeNextWriteCount, g_fileOpsBridgeUnderConsumeNextWriteAttempts))
        {
            unsigned long persistedBytes = 0;
            const HRESULT hr = _inner->Write(buffer, bytesToWrite - 1u, &persistedBytes);
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

private:
    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileWriter> _inner;
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

void MaybePauseBeforeBridgeMoveSourceCleanupForSelfTest() noexcept
{
    g_fileOpsBridgeMoveSourceCleanupPausePoint.Pause(5'000ull);
}

void MaybePauseAfterBridgeMoveManifestTakeForSelfTest() noexcept
{
    g_fileOpsBridgeMoveManifestTakePausePoint.Pause(5'000ull);
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

void MaybeMutateBridgeDestinationBeforeMoveCleanupForSelfTest(IFileSystemIO& destinationIo, const std::wstring& destinationPath) noexcept
{
    constexpr const wchar_t* kMutationPathEnv    = L"REDSALAMANDER_FILEOPS_BRIDGE_MUTATE_DESTINATION_BEFORE_MOVE_CLEANUP_PATH";
    constexpr const wchar_t* kMutationPayloadEnv = L"REDSALAMANDER_FILEOPS_BRIDGE_MUTATE_DESTINATION_BEFORE_MOVE_CLEANUP_PAYLOAD";

    const std::optional<std::wstring> configured = TryReadEnvironmentVariableForSelfTest(kMutationPathEnv);
    auto normalizeForCompare = [](std::wstring value) noexcept
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

    static_cast<void>(SetEnvironmentVariableW(kMutationPathEnv, nullptr));
    const std::optional<std::wstring> payload = TryReadEnvironmentVariableForSelfTest(kMutationPayloadEnv);
    if (! payload.has_value())
    {
        return;
    }

    const std::string bytes = NarrowEnvironmentPayloadForSelfTest(*payload);
    if (bytes.size() > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return;
    }

    g_fileOpsBridgeMutateDestinationBeforeMoveCleanupAttempts.fetch_add(1u, std::memory_order_acq_rel);

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
#endif

enum class ReparsePointPolicy : unsigned char
{
    CopyReparse,
    FollowTargets,
    Skip,
};

enum class FileSystemConcurrencyMode : unsigned char
{
    Auto,
    Manual,
};

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
    if (text == "followTargets")
    {
        return ReparsePointPolicy::FollowTargets;
    }
    if (text == "skip")
    {
        return ReparsePointPolicy::Skip;
    }

    return ReparsePointPolicy::CopyReparse;
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

[[nodiscard]] std::optional<ReparsePointPolicy> TryGetReparsePointPolicyFromFileSystem(const wil::com_ptr<IFileSystem>& fileSystem) noexcept
{
    ParsedFileSystemConfiguration parsed = TryParseFileSystemConfiguration(fileSystem);
    if (! parsed)
    {
        return std::nullopt;
    }

    yyjson_val* policyVal = yyjson_obj_get(parsed.root, "reparsePointPolicy");
    if (! policyVal || ! yyjson_is_str(policyVal))
    {
        return std::nullopt;
    }

    const char* policyText = yyjson_get_str(policyVal);
    if (! policyText || policyText[0] == '\0')
    {
        return std::nullopt;
    }

    return ParseReparsePointPolicy(policyText);
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
        return ReparsePointPolicy::CopyReparse;
    }

    const Common::Settings::JsonValue& config = it->second;
    if (! std::holds_alternative<Common::Settings::JsonValue::ObjectPtr>(config.value))
    {
        return ReparsePointPolicy::CopyReparse;
    }

    const auto obj = std::get<Common::Settings::JsonValue::ObjectPtr>(config.value);
    if (! obj)
    {
        return ReparsePointPolicy::CopyReparse;
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
            return ReparsePointPolicy::CopyReparse;
        }

        const std::string& text = std::get<std::string>(v.value);
        return ParseReparsePointPolicy(text);
    }

    return ReparsePointPolicy::CopyReparse;
}

constexpr std::wstring_view kFileOpsAppId                    = L"RedSalamander";
constexpr std::wstring_view kFileOpsIssuesPaneWindowId       = L"FileOperationsIssuesPane";
constexpr std::wstring_view kFileOpsPopupWindowId            = L"FileOperationsPopup";
constexpr std::wstring_view kFileOpsPopupExpandedWindowId    = L"FileOperationsPopupExpanded";
constexpr std::wstring_view kDiagnosticsLogPrefix            = L"FileOperations-";
constexpr std::wstring_view kDiagnosticsLogExtension         = L".jsonl";
constexpr std::wstring_view kDiagnosticsIssueReportPrefix    = L"FileOperations-Issues-";
constexpr std::wstring_view kDiagnosticsIssueReportExtension = L".txt";
constexpr size_t kMaxCompletedTaskSummaries                  = 24u;
constexpr size_t kMaxTaskIssueDiagnostics                    = 128u;
constexpr size_t kDefaultMaxDiagnosticsInMemory              = 256u;
constexpr size_t kDefaultMaxDiagnosticsPerFlush              = 64u;
constexpr size_t kDefaultMaxDiagnosticsLogFiles              = 14u;
constexpr size_t kDefaultMaxDiagnosticsIssueReportFiles      = 60u;
constexpr ULONGLONG kDefaultDiagnosticsFlushIntervalMs       = 5'000ull;
constexpr ULONGLONG kDefaultDiagnosticsCleanupIntervalMs     = 15ull * 60ull * 1000ull;

struct DiagnosticsSettings
{
    size_t maxDiagnosticsInMemory          = kDefaultMaxDiagnosticsInMemory;
    size_t maxDiagnosticsPerFlush          = kDefaultMaxDiagnosticsPerFlush;
    size_t maxDiagnosticsLogFiles          = kDefaultMaxDiagnosticsLogFiles;
    size_t maxDiagnosticsIssueReportFiles  = kDefaultMaxDiagnosticsIssueReportFiles;
    ULONGLONG diagnosticsFlushIntervalMs   = kDefaultDiagnosticsFlushIntervalMs;
    ULONGLONG diagnosticsCleanupIntervalMs = kDefaultDiagnosticsCleanupIntervalMs;
#if defined(_DEBUG) || defined(DEBUG)
    bool infoEnabled  = true;
    bool debugEnabled = true;
#else
    bool infoEnabled  = false;
    bool debugEnabled = false;
#endif
};

struct PreCalcProgressCookie
{
    std::mutex* totalsMutex                   = nullptr;
    uint64_t* totalBytes                      = nullptr;
    uint64_t* totalFiles                      = nullptr;
    uint64_t* totalDirs                       = nullptr;
    std::vector<uint64_t>* sourceBytesByIndex = nullptr;
    size_t sourceIndex                        = 0;
    std::atomic<bool>* acceptUpdates          = nullptr;
    uint64_t lastBytes                        = 0;
    uint64_t lastFiles                        = 0;
    uint64_t lastDirs                         = 0;
};

void UpdatePreCalcSnapshot(Task& task, uint64_t totalBytes, uint64_t totalFiles, uint64_t totalDirs) noexcept
{
    constexpr uint64_t maxUlong = static_cast<uint64_t>(std::numeric_limits<unsigned long>::max());
    task._preCalcTotalBytes.store(totalBytes, std::memory_order_release);
    task._preCalcFileCount.store(static_cast<unsigned long>(std::min(totalFiles, maxUlong)), std::memory_order_release);
    task._preCalcDirectoryCount.store(static_cast<unsigned long>(std::min(totalDirs, maxUlong)), std::memory_order_release);
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

struct PublishedProgressSnapshot
{
    unsigned long totalItems            = 0;
    unsigned long completedItems        = 0;
    uint64_t totalBytes                 = 0;
    uint64_t completedBytes             = 0;
    uint64_t itemTotalBytes             = 0;
    uint64_t itemCompletedBytes         = 0;
    unsigned long completedFiles        = 0;
    unsigned long completedFolders      = 0;
    uint64_t progressCallbackCount      = 0;
    uint64_t itemCompletedCallbackCount = 0;
};

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

void UpdateInFlightFileProgress(Task& task,
                                const void* cookieKey,
                                uint64_t progressStreamId,
                                const wchar_t* currentSourcePath,
                                uint64_t currentItemTotalBytes,
                                uint64_t currentItemCompletedBytes,
                                ULONGLONG nowTick) noexcept
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
        }
        task._inFlightFiles[found].totalBytes     = currentItemTotalBytes;
        task._inFlightFiles[found].completedBytes = currentItemCompletedBytes;
        task._inFlightFiles[found].lastUpdateTick = nowTick;
        return;
    }

    Task::InFlightFileProgress added{};
    added.cookieKey        = cookieKey;
    added.progressStreamId = streamKey;
    added.sourcePath       = currentSourcePath;
    added.totalBytes       = currentItemTotalBytes;
    added.completedBytes   = currentItemCompletedBytes;
    added.lastUpdateTick   = nowTick;

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

constexpr unsigned int kDefaultPreCalcMaxWorkers         = 4u;
constexpr unsigned int kMaxPreCalcWorkersSetting         = 8u;
constexpr unsigned int kDefaultCrossFsBridgeBufferSizeKB = 4096u;
constexpr unsigned int kMinCrossFsBridgeBufferSizeKB     = 512u;
constexpr unsigned int kMaxCrossFsBridgeBufferSizeKB     = 16384u;
constexpr uint64_t kCrossFsBridgeBufferBudgetBytes       = 256ull * 1024ull * 1024ull;

class CrossFsBridgeBufferBudget final
{
public:
    CrossFsBridgeBufferBudget() = default;
    ~CrossFsBridgeBufferBudget() = default;
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
constexpr uint64_t kDefaultBandwidthLimitBytesPerSecond  = 0;
constexpr size_t kBridgeAdmissionQueueLimit              = 16u;

[[nodiscard]] size_t GetBridgeAdmissionQueueLimit() noexcept
{
    return kBridgeAdmissionQueueLimit;
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
        name.find_first_of(L":*?\"<>|") != std::wstring_view::npos ||
        IsReservedWindowsBridgeChildName(name))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    return S_OK;
}

[[nodiscard]] bool UsesOrdinalIgnoreCaseComponents(IFileSystem& fileSystem, std::wstring_view pluginId) noexcept
{
    const char* jsonUtf8 = nullptr;
    if (FAILED(fileSystem.GetCapabilities(&jsonUtf8)) || jsonUtf8 == nullptr || jsonUtf8[0] == '\0')
    {
        return false;
    }

    const std::optional<FileSystemPathIdentity> identity = TryParseFileSystemPathIdentityContract(jsonUtf8, pluginId);
    return identity.has_value() && identity->pathTextStableIdentity &&
           identity->componentComparison == FileSystemPathComponentComparison::OrdinalIgnoreCase;
}

struct BridgeOrdinalIgnoreCaseLess final
{
    [[nodiscard]] bool operator()(const std::wstring& left, const std::wstring& right) const noexcept
    {
        return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) == CSTR_LESS_THAN;
    }
};

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

[[nodiscard]] bool GetPreCalcEnabledFromSettings(const Common::Settings::Settings* settings) noexcept
{
    if (! settings || ! settings->fileOperations.has_value())
    {
        return true;
    }

    return settings->fileOperations->preCalcEnabled;
}

[[nodiscard]] unsigned int GetPreCalcMaxWorkersFromSettings(const Common::Settings::Settings* settings) noexcept
{
    if (! settings || ! settings->fileOperations.has_value())
    {
        return kDefaultPreCalcMaxWorkers;
    }

    return std::clamp(settings->fileOperations->preCalcMaxWorkers, 1u, kMaxPreCalcWorkersSetting);
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

[[nodiscard]] bool HasHighMetadataCostTransferHint(IFileSystem& fileSystem,
                                                   const std::vector<std::filesystem::path>& paths,
                                                   FileSystemOperation operationType,
                                                   FileSystemTransferEndpoint endpoint) noexcept
{
    for (const std::filesystem::path& path : paths)
    {
        FileSystemTransferHints hints{};
        hints.sizeBytes = sizeof(hints);
        if (SUCCEEDED(fileSystem.GetTransferHints(path.c_str(), operationType, endpoint, &hints)) &&
            (hints.flags & FILESYSTEM_TRANSFER_HINT_HIGH_METADATA_COST) != 0u)
        {
            return true;
        }
    }
    return false;
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
        sawBufferHint                  = true;
        resolvedBytes = (std::max)(resolvedBytes, hintedBytes);
    };

    applyHints(sourceFileSystem, sourcePath, FILESYSTEM_TRANSFER_SOURCE_READ);
    applyHints(destinationFileSystem, destinationPath, FILESYSTEM_TRANSFER_DESTINATION_WRITE);
    if (sawBufferHint)
    {
        // The setting is a fallback. Provider hints remain active even when the user changed it;
        // this avoids silently disabling WAN/cloud tuning for every non-default preference value.
        tuning.bufferBytes = resolvedBytes;
    }
    else if (tuning.latencyClass >= FILESYSTEM_TRANSFER_LATENCY_WAN ||
             (tuning.flags & FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS) != 0u)
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

struct ProcessMemorySnapshot
{
    uint64_t workingSetBytes = 0;
    uint64_t privateBytes    = 0;
};

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
                                                                    FileSystemOperation operation,
                                                                    FileSystemFlags flags,
                                                                    unsigned int uiMax) noexcept
{
    if (! fileSystem || uiMax == 0u)
    {
        return 1u;
    }

    const bool isCopyMove = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
    const bool isDelete   = operation == FILESYSTEM_DELETE;
    if (! isCopyMove && ! isDelete)
    {
        return 1u;
    }

    const char* capabilitiesText = nullptr;
    if (FAILED(fileSystem->GetCapabilities(&capabilitiesText)) || ! capabilitiesText || capabilitiesText[0] == '\0')
    {
        return 1u;
    }

    const std::string_view capabilitiesView(capabilitiesText);
    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc(
        yyjson_read(capabilitiesView.data(), capabilitiesView.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM), &yyjson_doc_free);
    if (! doc)
    {
        return 1u;
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root))
    {
        return 1u;
    }

    yyjson_val* concurrencyObject = yyjson_obj_get(root, "concurrency");
    if (! concurrencyObject || ! yyjson_is_obj(concurrencyObject))
    {
        return 1u;
    }

    const char* key = nullptr;
    if (isCopyMove)
    {
        key = "copyMoveMax";
    }
    else if (isDelete)
    {
        key = (flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0 ? "deleteRecycleBinMax" : "deleteMax";
    }

    if (! key)
    {
        return 1u;
    }

    yyjson_val* valueNode = yyjson_obj_get(concurrencyObject, key);
    if (! valueNode)
    {
        return 1u;
    }

    uint64_t concurrency = 0;
    if (yyjson_is_uint(valueNode))
    {
        concurrency = yyjson_get_uint(valueNode);
    }
    else if (yyjson_is_int(valueNode))
    {
        const int64_t signedValue = yyjson_get_int(valueNode);
        if (signedValue > 0)
        {
            concurrency = static_cast<uint64_t>(signedValue);
        }
    }

    if (concurrency == 0)
    {
        return 1u;
    }

    return std::clamp(static_cast<unsigned int>(std::min<uint64_t>(concurrency, static_cast<uint64_t>(uiMax))), 1u, uiMax);
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

    return DetermineConfiguredPerItemMaxConcurrency(fileSystem, operation, flags, uiMax);
}

[[nodiscard]] unsigned int DeterminePerItemMaxConcurrency(const wil::com_ptr<IFileSystem>& fileSystem,
                                                          const std::filesystem::path& path,
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

    return DetermineConfiguredPerItemMaxConcurrency(fileSystem, operation, flags, uiMax);
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
        return Task::ConflictBucket::RecycleBinFailed;
    }

    const std::optional<DWORD> errorOpt = Win32ErrorFromHRESULT(status);
    const DWORD error                   = errorOpt.value_or(0);

    switch (error)
    {
        case ERROR_ALREADY_EXISTS:
        case ERROR_FILE_EXISTS: return Task::ConflictBucket::Exists;
        // Replacing a non-empty destination directory (e.g. with a reparse point) requires its
        // own explicit consent; the engine raises ERROR_DIR_NOT_EMPTY so the prompt can offer
        // Overwrite for exactly that path.
        case ERROR_DIR_NOT_EMPTY:
            if (IsCopyMoveOperation(operation))
            {
                return Task::ConflictBucket::NonEmptyDirectory;
            }
            break;
        // A junction/symlink occupying the destination is never a silent merge target; the
        // engine raises this so Overwrite consent can replace the LINK with a real directory.
        case ERROR_REPARSE_POINT_ENCOUNTERED:
            if (IsCopyMoveOperation(operation))
            {
                return Task::ConflictBucket::ReparsePoint;
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
                return Task::ConflictBucket::ReadOnly;
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
    return bucket != Task::ConflictBucket::UnsupportedReparse && bucket != Task::ConflictBucket::Exists && bucket != Task::ConflictBucket::NonEmptyDirectory &&
           bucket != Task::ConflictBucket::ReparsePoint && bucket != Task::ConflictBucket::ReadOnly;
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

[[nodiscard]] Task::ConflictPromptState::ItemMetadata ReadConflictItemMetadata(
    IFileSystemIO* io, std::wstring_view path, bool useWin32Metadata) noexcept
{
    Task::ConflictPromptState::ItemMetadata result{};
    if (path.empty())
    {
        return result;
    }

    const std::wstring pathText(path);

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

        result.available     = true;
        result.attributes    = data.dwFileAttributes;
        result.isDirectory   = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const uint64_t lastWriteTime = (static_cast<uint64_t>(data.ftLastWriteTime.dwHighDateTime) << 32) |
                                       static_cast<uint64_t>(data.ftLastWriteTime.dwLowDateTime);
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
    basic.sizeBytes = sizeof(basic);
    const HRESULT basicHr = io->GetFileBasicInformation(pathText.c_str(), &basic);
    if (SUCCEEDED(basicHr))
    {
        result.available     = true;
        result.attributes    = basic.attributes;
        result.isDirectory   = (basic.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        result.lastWriteTime = basic.lastWriteTime;
        return result;
    }

    unsigned long attributes   = 0;
    const HRESULT attributesHr = io->GetAttributes(pathText.c_str(), &attributes);
    if (FAILED(attributesHr))
    {
        return result;
    }

    result.available     = true;
    result.attributes    = attributes;
    result.isDirectory   = (result.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    result.lastWriteTime = 0;
    return result;
}

struct ConflictActionLayout
{
    std::array<ConflictAction, Task::ConflictPromptState::kMaxActions> actions{};
    size_t actionCount = 0;

    void Add(ConflictAction action) noexcept
    {
        if (actionCount < actions.size())
        {
            actions[actionCount] = action;
            ++actionCount;
        }
    }
};

[[nodiscard]] ConflictActionLayout BuildConflictActionLayout(ConflictBucket bucket, bool allowRetry, bool suppressExistsOverwrite) noexcept
{
    ConflictActionLayout layout{};

    switch (bucket)
    {
        case ConflictBucket::Exists:
            if (! suppressExistsOverwrite)
            {
                layout.Add(ConflictAction::Overwrite);
            }
            break;
        case ConflictBucket::NonEmptyDirectory:
        case ConflictBucket::ReparsePoint: layout.Add(ConflictAction::Overwrite); break;
        case ConflictBucket::ReadOnly: layout.Add(ConflictAction::ReplaceReadOnly); break;
        case ConflictBucket::RecycleBinFailed: layout.Add(ConflictAction::PermanentDelete); break;
        case ConflictBucket::AccessDenied:
        case ConflictBucket::SharingViolation:
        case ConflictBucket::DiskFull:
        case ConflictBucket::PathTooLong:
        case ConflictBucket::NetworkOffline:
        case ConflictBucket::UnsupportedReparse:
        case ConflictBucket::Unknown:
        case ConflictBucket::Count:
        default: break;
    }

    if (allowRetry)
    {
        layout.Add(ConflictAction::Retry);
    }

    layout.Add(ConflictAction::Skip);
    layout.Add(ConflictAction::Cancel);
    return layout;
}

[[nodiscard]] bool IsCacheableConflictDecision(ConflictAction action) noexcept
{
    return action != ConflictAction::Retry && action != ConflictAction::Cancel && action != ConflictAction::None;
}

void StoreConflictDecisionInCacheLocked(Task::ConflictArbiter& arbiter, ConflictBucket bucket, ConflictAction action) noexcept
{
    if (! IsCacheableConflictDecision(action))
    {
        return;
    }

    const size_t bucketIndex = static_cast<size_t>(bucket);
    if (bucketIndex >= arbiter.decisionCache.size())
    {
        return;
    }

    arbiter.decisionCache[bucketIndex] = action;
}

[[nodiscard]] std::optional<ConflictAction> LoadConflictDecisionFromCacheLocked(const Task::ConflictArbiter& arbiter, ConflictBucket bucket) noexcept
{
    const size_t bucketIndex = static_cast<size_t>(bucket);
    if (bucketIndex >= arbiter.decisionCache.size())
    {
        return std::nullopt;
    }

    return arbiter.decisionCache[bucketIndex];
}

[[nodiscard]] std::optional<ConflictAction> LoadConflictDecisionFromCache(Task& task, ConflictBucket bucket) noexcept
{
    std::scoped_lock lock(task._conflictArbiter.mutex);
    return LoadConflictDecisionFromCacheLocked(task._conflictArbiter, bucket);
}

void ClearConflictPrompt(Task& task, std::optional<std::pair<ConflictBucket, ConflictAction>> cachedDecision = std::nullopt) noexcept
{
    {
        std::scoped_lock lock(task._conflictArbiter.mutex);
        if (cachedDecision.has_value())
        {
            StoreConflictDecisionInCacheLocked(task._conflictArbiter, cachedDecision->first, cachedDecision->second);
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

void SetConflictPromptLocked(Task& task,
                             ConflictBucket bucket,
                             HRESULT status,
                             std::wstring sourcePath,
                             std::wstring destinationPath,
                             bool allowRetry,
                             bool retryFailed,
                             bool suppressExistsOverwrite) noexcept
{
    if (task._conflictArbiter.decisionEvent)
    {
        static_cast<void>(ResetEvent(task._conflictArbiter.decisionEvent.get()));
    }

    Task::ConflictPromptState prompt{};
    prompt.active            = true;
    prompt.bucket            = bucket;
    prompt.status            = status;
    prompt.sourcePath        = std::move(sourcePath);
    prompt.destinationPath   = std::move(destinationPath);
    prompt.applyToAllChecked = false;
    prompt.retryFailed       = retryFailed;

    const ConflictActionLayout layout = BuildConflictActionLayout(bucket, allowRetry, suppressExistsOverwrite);
    prompt.actions                    = layout.actions;
    prompt.actionCount                = layout.actionCount;

    task._conflictArbiter.prompt        = std::move(prompt);
    task._conflictArbiter.ownerThreadId = GetCurrentThreadId();

    task._conflictArbiter.decisionAction.reset();
    task._conflictArbiter.decisionApplyToAll = false;
}

struct ConflictPromptBeginResult
{
    ConflictAction action = ConflictAction::None;
    bool ownsPrompt       = false;
};

[[nodiscard]] ConflictPromptBeginResult BeginConflictPrompt(Task& task,
                                                            const Task::PerItemCallbackCookie* perItemCookie,
                                                            ConflictBucket bucket,
                                                            HRESULT status,
                                                            std::wstring_view sourcePath,
                                                            std::wstring_view destinationPath,
                                                            bool allowRetry,
                                                            bool retryFailed,
                                                            bool ignoreCachedDecision) noexcept
{
    auto [promptSourcePath, promptDestinationPath] = GetMostSpecificPathsForDiagnostics(task, perItemCookie, sourcePath, destinationPath);
    const DWORD ownerThreadId          = GetCurrentThreadId();
    const bool sourceUsesWin32Metadata = NavigationLocation::EqualsNoCase(task._sourcePluginId, L"builtin/file-system");
    const bool destinationUsesWin32Metadata = NavigationLocation::EqualsNoCase(task._destinationPluginId, L"builtin/file-system") ||
                                              (task._destinationPluginId.empty() && ! task._destinationFileSystem && sourceUsesWin32Metadata);
    const bool deferLocalExistsOverwrite = bucket == ConflictBucket::Exists && sourceUsesWin32Metadata && destinationUsesWin32Metadata;

    std::unique_lock lock(task._conflictArbiter.mutex);

    if (! ignoreCachedDecision)
    {
        if (const std::optional<ConflictAction> cachedDecision = LoadConflictDecisionFromCacheLocked(task._conflictArbiter, bucket); cachedDecision.has_value())
        {
            return {cachedDecision.value(), false};
        }
    }

    task._conflictArbiter.cv.wait(lock, [&]() noexcept {
        return ! task._conflictArbiter.prompt.active || task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested();
    });

    if (task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested())
    {
        return {ConflictAction::Cancel, false};
    }

    if (! ignoreCachedDecision)
    {
        if (const std::optional<ConflictAction> cachedDecision = LoadConflictDecisionFromCacheLocked(task._conflictArbiter, bucket); cachedDecision.has_value())
        {
            return {cachedDecision.value(), false};
        }
    }

    SetConflictPromptLocked(task,
                            bucket,
                            status,
                            promptSourcePath,
                            promptDestinationPath,
                            allowRetry,
                            retryFailed,
                            deferLocalExistsOverwrite);
    lock.unlock();

    task.LogDiagnostic(FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                       status,
                       L"item.conflict.prompt",
                       retryFailed ? L"Conflict prompt shown after retry cap reached." : L"Conflict prompt shown for item.",
                       promptSourcePath,
                       promptDestinationPath);

    const uint64_t metadataStartUs = PerfNowUs();
    const wil::com_ptr<IFileSystemIO> sourceIo = QueryFileSystemIo(task._fileSystem.get());
    wil::com_ptr<IFileSystemIO> destinationIo  = task._destinationFileSystem ? QueryFileSystemIo(task._destinationFileSystem.get()) : sourceIo;
    const Task::ConflictPromptState::ItemMetadata sourceMetadata =
        ReadConflictItemMetadata(sourceIo.get(), promptSourcePath, sourceUsesWin32Metadata);
    const Task::ConflictPromptState::ItemMetadata destinationMetadata =
        ReadConflictItemMetadata(destinationIo.get(), promptDestinationPath, destinationUsesWin32Metadata);
    const uint64_t metadataUs = PerfElapsedUs(metadataStartUs);
    task._perf.conflictMetadataUs.fetch_add(metadataUs, std::memory_order_relaxed);
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Conflict.MetadataPromptUs",
                          L"",
                          metadataUs,
                          sourceMetadata.available ? 1u : 0u,
                          destinationMetadata.available ? 1u : 0u,
                          S_OK);
    }

    lock.lock();
    Task::ConflictPromptState& livePrompt = task._conflictArbiter.prompt;
    if (livePrompt.active && task._conflictArbiter.ownerThreadId == ownerThreadId && livePrompt.bucket == bucket && livePrompt.status == status &&
        livePrompt.sourcePath == promptSourcePath && livePrompt.destinationPath == promptDestinationPath)
    {
        livePrompt.sourceMetadata      = sourceMetadata;
        livePrompt.destinationMetadata = destinationMetadata;
        const bool suppressExistsOverwrite = bucket == ConflictBucket::Exists && sourceUsesWin32Metadata && destinationUsesWin32Metadata &&
                                             (! sourceMetadata.available || ! destinationMetadata.available ||
                                              (! sourceMetadata.isDirectory && destinationMetadata.isDirectory));
        const ConflictActionLayout layout = BuildConflictActionLayout(bucket, allowRetry, suppressExistsOverwrite);
        livePrompt.actions                = layout.actions;
        livePrompt.actionCount            = layout.actionCount;
    }
    return {ConflictAction::None, true};
}

[[nodiscard]] std::pair<ConflictAction, bool> WaitForConflictDecision(Task& task, const void* cookieKey, ConflictBucket bucket) noexcept
{
    const uint64_t perfStartUs   = PerfNowUs();
    const auto perfCallbackScope = wil::scope_exit([&] noexcept
    {
        const uint64_t waitUs = PerfElapsedUs(perfStartUs);
        task._perf.conflictWaitUs += waitUs;
        ++task._perf.conflictPromptCount;
        NoteConflictWorkerWait(task, cookieKey, waitUs);
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Conflict.WaitUs", L"", waitUs, 0u, 0u, S_OK);
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
        ClearConflictPrompt(task, std::pair{bucket, action});
        return {action, true};
    }

    ClearConflictPrompt(task);
    return {action, applyToAll};
}

[[nodiscard]] bool IsModifierConflictAction(ConflictAction action) noexcept
{
    switch (action)
    {
        case ConflictAction::Overwrite:
        case ConflictAction::ReplaceReadOnly:
        case ConflictAction::PermanentDelete: return true;
        case ConflictAction::None:
        case ConflictAction::Retry:
        case ConflictAction::Skip:
        case ConflictAction::Cancel:
        default: return false;
    }
}

class PerItemTaskScheduler final
{
public:
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
        std::function<HRESULT(size_t)> processIndex;
        size_t totalItems           = 0;
        unsigned int maxConcurrency = 1;

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

    JobPtr StartJob(Task* task, unsigned int maxConcurrency, size_t totalItems, std::function<HRESULT(size_t)> processIndex)
    {
        auto job            = std::make_shared<Job>();
        job->task           = task;
        job->totalItems     = totalItems;
        job->processIndex   = std::move(processIndex);
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
                if (job->processIndex)
                {
                    static_cast<void>(job->processIndex(i));
                }
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
        if (job && job->processIndex)
        {
            const uint64_t processStartUs = PerfNowUs();
            static_cast<void>(job->processIndex(index));
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

    void workerMain(std::stop_token stopToken) noexcept
    {
        [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();
        PerItemTaskScheduler* const previousScheduler = s_currentScheduler;
        s_currentScheduler                              = this;
        const auto restoreScheduler = wil::scope_exit([&]() noexcept { s_currentScheduler = previousScheduler; });

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

    const auto job = scheduler.StartJob(&task,
                                        1u,
                                        1u,
                                        [&](size_t) noexcept -> HRESULT
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
    });

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
        SaturationState()                                 = default;
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

    std::vector<PerItemTaskScheduler::JobPtr> outerJobs;
    outerJobs.reserve(workerCount);
    for (unsigned int outerIndex = 0u; outerIndex < workerCount; ++outerIndex)
    {
        outerJobs.push_back(scheduler.StartJob(&task,
                                               1u,
                                               1u,
                                               [&](size_t) noexcept -> HRESULT
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

            const auto nestedJob = scheduler.StartJob(&task,
                                                      1u,
                                                      1u,
                                                      [&](size_t) noexcept -> HRESULT
            {
                saturation.nestedCompleted.fetch_add(1u, std::memory_order_acq_rel);
                return S_OK;
            });
            scheduler.WaitJob(nestedJob);
            return S_OK;
        }));
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
    return true;
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
} // namespace

#ifdef ENABLE_TESTS
bool RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept
{
    return RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTestInternal(state);
}

bool RunFileOpsPerItemSchedulerNestedSaturationSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept
{
    return RunFileOpsPerItemSchedulerNestedSaturationSelfTestForSelfTestInternal(state);
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

bool IsFileOpsCircuitBreakerTransientErrorForSelfTest(DWORD error) noexcept
{
    return IsCircuitBreakerTransientError(error);
}

bool DebugReadFileOpsConflictMetadataForSelfTest(IFileSystemIO* io,
                                                 std::wstring_view path,
                                                 FileOpsConflictMetadataDebugResult& out) noexcept
{
    const Task::ConflictPromptState::ItemMetadata metadata = ReadConflictItemMetadata(io, path, false);
    out.available                                         = metadata.available;
    out.isDirectory                                       = metadata.isDirectory;
    out.attributes                                        = metadata.attributes;
    out.lastWriteTime                                     = metadata.lastWriteTime;
    return metadata.available;
}
#endif

FolderWindow::FileOperationState::Task::Task(FileOperationState& state) noexcept : _state(&state), _folderWindow(&state._owner)
{
    _conflictArbiter.decisionEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
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

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    const ULONGLONG nowTick                   = GetTickCount64();
    const uint64_t perfStartUs                = PerfNowUs();
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
        const uint64_t progressLockHoldStartUs = PerfNowUs();
        const auto progressLockHoldScope       = wil::scope_exit([&] noexcept { _perf.progressLockHoldUs += PerfElapsedUs(progressLockHoldStartUs); });
        trackProgressStreamPerf                = true;
        // 5F early admission: this is a TRANSFER progress callback. If it fires while pre-calc is
        // still running, bytes are moving before the recursive scan finished — the overlap the
        // serial pre-calc-then-execute model could never produce. Latch it as the deterministic proof.
        if (! _transferStartedBeforePreCalcComplete.load(std::memory_order_acquire) && _preCalcInProgress.load(std::memory_order_acquire) &&
            ! _preCalcCompleted.load(std::memory_order_acquire))
        {
            _transferStartedBeforePreCalcComplete.store(true, std::memory_order_release);
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
                const bool precalcTotalAvailable = _preCalcCompleted.load(std::memory_order_acquire) && _progressTotalItems > 0;
                if (! precalcTotalAvailable)
                {
                    const uint64_t mappedTotalItems = _perItemTotalEntryCount + perItemInFlightAggregate.totalItems;
                    if (mappedTotalItems > 0)
                    {
                        const uint64_t clamped = std::min<uint64_t>(mappedTotalItems, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()));
                        _progressTotalItems    = (std::max)(_progressTotalItems, static_cast<unsigned long>(clamped));
                    }
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

        if (_operation != FILESYSTEM_DELETE)
        {
            const unsigned long plannedTopLevelItems = (_executionMode == ExecutionMode::PerItem) ? _perItemTotalItems : GetPlannedItemCount();
            const bool havePreCalcTotals =
                _preCalcCompleted.load(std::memory_order_acquire) && _progressTotalItems > 0 && _progressTotalBytes > 0 && plannedTopLevelItems > 0;

            const bool pluginLikelyReportsTopLevelItems =
                (_executionMode == ExecutionMode::PerItem) ? true : (totalItems == 0 || totalItems <= plannedTopLevelItems);

            if (havePreCalcTotals && pluginLikelyReportsTopLevelItems && _progressTotalItems > plannedTopLevelItems)
            {
                const uint64_t clampedBytes                 = (std::min)(_progressCompletedBytes, _progressTotalBytes);
                const long double ratio                     = static_cast<long double>(clampedBytes) / static_cast<long double>(_progressTotalBytes);
                const long double estimate                  = ratio * static_cast<long double>(_progressTotalItems);
                const long double clampedEstimate           = std::clamp<long double>(estimate, 0.0L, static_cast<long double>(_progressTotalItems));
                const unsigned long estimatedCompletedItems = static_cast<unsigned long>(clampedEstimate);
                _progressCompletedItems                     = (std::max)(_progressCompletedItems, estimatedCompletedItems);
            }
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
        UpdateInFlightFileProgress(*this, cookie, progressStreamId, currentSourcePath, currentItemTotalBytes, currentItemCompletedBytes, nowTick);
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

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::FileSystemItemCompleted(FileSystemOperation operationType,
                                                                                          unsigned long itemIndex,
                                                                                          const wchar_t* sourcePath,
                                                                                          const wchar_t* destinationPath,
                                                                                          HRESULT status,
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

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    const uint64_t perfStartUs                = PerfNowUs();
    const auto perfCallbackScope              = wil::scope_exit([&] noexcept { _perf.itemCompletedCallbackUs += PerfElapsedUs(perfStartUs); });
    unsigned int perItemBandwidthActiveCalls  = 1u;
    const uint64_t itemCompletedCallbackCount = _itemCompletedCallbackCount.fetch_add(1, std::memory_order_relaxed) + 1u;
    PublishedProgressSnapshot publishedProgressSnapshot{};

    size_t sourceIndex = static_cast<size_t>(itemIndex);
    if (_executionMode == ExecutionMode::PerItem && cookie != nullptr)
    {
        sourceIndex = static_cast<PerItemCallbackCookie*>(cookie)->itemIndex;
    }
    if (sourceIndex < _sourcePaths.size())
    {
        std::scoped_lock lock(_sourceItemStatusMutex);
        if (_sourceItemStatuses.size() < _sourcePaths.size())
        {
            _sourceItemStatuses.resize(_sourcePaths.size());
        }
        _sourceItemStatuses[sourceIndex] = status;
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

    UpdateItemCompletedPathState(*this,
                                 (_executionMode == ExecutionMode::PerItem && cookie != nullptr) ? static_cast<PerItemCallbackCookie*>(cookie) : nullptr,
                                 sourcePath,
                                 destinationPath);

    ApplyCallbackBandwidthLimit(*this, options, perItemBandwidthActiveCalls);

    RemoveInFlightFileBySourcePath(*this, sourcePath);

    if (status == S_FALSE)
    {
        _observedSkipAction.store(true, std::memory_order_release);
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
                                                                                  [[maybe_unused]] FileSystemOptions* options,
                                                                                  void* cookie) noexcept
{
#ifdef ENABLE_TESTS
    assert(_dbgCallbackActiveScopeCount.load(std::memory_order_acquire) > 0u);
#endif

    if (! action)
    {
        return E_POINTER;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    *action = FileSystemIssueAction::Cancel;

    WaitWhilePaused();

    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    const std::wstring_view sourceText      = sourcePath ? sourcePath : L"";
    const std::wstring_view destinationText = destinationPath ? destinationPath : L"";

    PerItemCallbackCookie* perItemCookie = nullptr;
    if (_executionMode == ExecutionMode::PerItem && cookie != nullptr)
    {
        perItemCookie = static_cast<PerItemCallbackCookie*>(cookie);
    }

    const ConflictBucket bucket = ClassifyConflictBucket(operationType, _flags, wil::com_ptr<IFileSystemIO>{}, status, sourceText, destinationText, false);
    if (bucket == ConflictBucket::RecycleBinFailed)
    {
        auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, perItemCookie, sourceText, destinationText);
        LogDiagnostic(
            DiagnosticSeverity::Error, status, L"delete.recycleBin.item", L"Recycle Bin delete failed for item.", diagnosticSource, diagnosticDestination);
    }

    const size_t bucketIndex = static_cast<size_t>(bucket);

    ConflictAction decision = LoadConflictDecisionFromCache(*this, bucket).value_or(ConflictAction::None);
    if (decision == ConflictAction::None)
    {
        const bool canRetryBucket = IsRetryableConflictBucket(bucket);
        bool allowRetry           = canRetryBucket;
        bool retryFailed          = false;
        if (perItemCookie != nullptr && bucketIndex < perItemCookie->issueRetryCounts.size())
        {
            allowRetry  = canRetryBucket && perItemCookie->issueRetryCounts[bucketIndex] == 0u;
            retryFailed = canRetryBucket && perItemCookie->issueRetryCounts[bucketIndex] != 0u;
        }

        const ConflictPromptBeginResult promptBegin =
            BeginConflictPrompt(*this, perItemCookie, bucket, status, sourceText, destinationText, allowRetry, retryFailed, false);
        decision = promptBegin.action;
        if (promptBegin.ownsPrompt)
        {
            const auto result = WaitForConflictDecision(*this, cookie, bucket);
            decision          = result.first;
        }
    }

    switch (decision)
    {
        case ConflictAction::Overwrite: *action = FileSystemIssueAction::Overwrite; return S_OK;
        case ConflictAction::ReplaceReadOnly: *action = FileSystemIssueAction::ReplaceReadOnly; return S_OK;
        case ConflictAction::PermanentDelete: *action = FileSystemIssueAction::PermanentDelete; return S_OK;
        case ConflictAction::Retry:
            if (perItemCookie != nullptr && bucketIndex < perItemCookie->issueRetryCounts.size())
            {
                perItemCookie->issueRetryCounts[bucketIndex] = 1u;
            }
            *action = FileSystemIssueAction::Retry;
            return S_OK;
        case ConflictAction::Skip:
        {
            auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, perItemCookie, sourceText, destinationText);
            LogDiagnostic(
                DiagnosticSeverity::Warning, status, L"item.conflict.skip", L"Conflict action Skip item selected.", diagnosticSource, diagnosticDestination);
            _observedSkipAction.store(true, std::memory_order_release);
            *action = FileSystemIssueAction::Skip;
            return S_OK;
        }
        case ConflictAction::Cancel:
        case ConflictAction::None:
        default: *action = FileSystemIssueAction::Cancel; return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::DirectorySizeProgress(
    uint64_t /*scannedEntries*/, uint64_t totalBytes, uint64_t fileCount, uint64_t directoryCount, const wchar_t* currentPath, void* cookie) noexcept
{
#ifdef ENABLE_TESTS
    assert(_dbgCallbackActiveScopeCount.load(std::memory_order_acquire) > 0u);
#endif

    const uint64_t perfStartUs   = PerfNowUs();
    const auto perfCallbackScope = wil::scope_exit([&] noexcept
    {
        _perf.preCalcCallbackCount.fetch_add(1u, std::memory_order_relaxed);
        _perf.preCalcCallbackUs.fetch_add(PerfElapsedUs(perfStartUs), std::memory_order_relaxed);
    });

    WaitWhilePreCalcPaused();

    const bool shouldCancel = _cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire) || _stopToken.stop_requested();
    if (shouldCancel)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    auto* progressCookie = static_cast<PreCalcProgressCookie*>(cookie);
    if (progressCookie && progressCookie->totalsMutex && progressCookie->totalBytes && progressCookie->totalFiles && progressCookie->totalDirs)
    {
        if (progressCookie->acceptUpdates && ! progressCookie->acceptUpdates->load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const uint64_t bytesDelta = (totalBytes >= progressCookie->lastBytes) ? (totalBytes - progressCookie->lastBytes) : totalBytes;
        const uint64_t filesDelta = (fileCount >= progressCookie->lastFiles) ? (fileCount - progressCookie->lastFiles) : fileCount;
        const uint64_t dirsDelta  = (directoryCount >= progressCookie->lastDirs) ? (directoryCount - progressCookie->lastDirs) : directoryCount;
        progressCookie->lastBytes = totalBytes;
        progressCookie->lastFiles = fileCount;
        progressCookie->lastDirs  = directoryCount;

        if (bytesDelta > 0 || filesDelta > 0 || dirsDelta > 0)
        {
            uint64_t snapshotBytes               = 0;
            uint64_t snapshotFiles               = 0;
            uint64_t snapshotDirs                = 0;
            const uint64_t totalsLockWaitStartUs = PerfNowUs();
            {
                std::scoped_lock lock(*progressCookie->totalsMutex);
                if (progressCookie->acceptUpdates && ! progressCookie->acceptUpdates->load(std::memory_order_acquire))
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                if (std::numeric_limits<uint64_t>::max() - *progressCookie->totalBytes < bytesDelta)
                {
                    *progressCookie->totalBytes = std::numeric_limits<uint64_t>::max();
                }
                else
                {
                    *progressCookie->totalBytes += bytesDelta;
                }
                if (progressCookie->sourceBytesByIndex && progressCookie->sourceIndex < progressCookie->sourceBytesByIndex->size())
                {
                    uint64_t& sourceBytes = (*progressCookie->sourceBytesByIndex)[progressCookie->sourceIndex];
                    if (std::numeric_limits<uint64_t>::max() - sourceBytes < bytesDelta)
                    {
                        sourceBytes = std::numeric_limits<uint64_t>::max();
                    }
                    else
                    {
                        sourceBytes += bytesDelta;
                    }
                }
                if (std::numeric_limits<uint64_t>::max() - *progressCookie->totalFiles < filesDelta)
                {
                    *progressCookie->totalFiles = std::numeric_limits<uint64_t>::max();
                }
                else
                {
                    *progressCookie->totalFiles += filesDelta;
                }
                if (std::numeric_limits<uint64_t>::max() - *progressCookie->totalDirs < dirsDelta)
                {
                    *progressCookie->totalDirs = std::numeric_limits<uint64_t>::max();
                }
                else
                {
                    *progressCookie->totalDirs += dirsDelta;
                }

                snapshotBytes = *progressCookie->totalBytes;
                snapshotFiles = *progressCookie->totalFiles;
                snapshotDirs  = *progressCookie->totalDirs;
                _perf.preCalcLockWaitUs.fetch_add(PerfElapsedUs(totalsLockWaitStartUs), std::memory_order_relaxed);
            }
            UpdatePreCalcSnapshot(*this, snapshotBytes, snapshotFiles, snapshotDirs);
        }
    }
    else
    {
        UpdatePreCalcSnapshot(*this, totalBytes, fileCount, directoryCount);
    }

    if (currentPath && currentPath[0] != L'\0')
    {
        const uint64_t progressLockWaitStartUs = PerfNowUs();
        std::scoped_lock lock(_progressPathMutex);
        _perf.preCalcLockWaitUs.fetch_add(PerfElapsedUs(progressLockWaitStartUs), std::memory_order_relaxed);
        UpdateTrackedPathIfPresent(
            _progressSourcePath, currentPath, _perf.progressPathUpdateBytes, _perf.progressPathUpdateAppliedCount, _perf.progressPathUpdateSkippedCount);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FolderWindow::FileOperationState::Task::DirectorySizeShouldCancel(BOOL* pCancel, void* /*cookie*/) noexcept
{
#ifdef ENABLE_TESTS
    assert(_dbgCallbackActiveScopeCount.load(std::memory_order_acquire) > 0u);
#endif

    if (! pCancel)
    {
        return E_POINTER;
    }

    const bool cancel = _cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire) || _stopToken.stop_requested();
    *pCancel          = cancel ? TRUE : FALSE;
    return S_OK;
}

void FolderWindow::FileOperationState::Task::SkipPreCalculation() noexcept
{
    _preCalcSkipped.store(true, std::memory_order_release);
    LogDiagnostic(DiagnosticSeverity::Info, S_FALSE, L"precalc.skip", L"User skipped pre-calculation.");
    _pauseCv.notify_all();
}

void FolderWindow::FileOperationState::Task::RunPreCalculation() noexcept
{
    _preCalcWorkerCountUsed.store(0u, std::memory_order_release);

    if (! _enablePreCalc || (_operation != FILESYSTEM_COPY && _operation != FILESYSTEM_MOVE && _operation != FILESYSTEM_DELETE) ||
        _preCalcSkipped.load(std::memory_order_acquire))
    {
        return;
    }

    if (_sourcePaths.empty())
    {
        return;
    }

    // Query IFileSystemDirectoryOperations interface
    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    if (FAILED(_fileSystem->QueryInterface(IID_PPV_ARGS(&dirOps))) || ! dirOps)
    {
        return; // Interface not supported, proceed without totals
    }

    const uint64_t perfStartUs   = PerfNowUs();
    const auto perfCallbackScope = wil::scope_exit([&] noexcept { _perf.preCalcUs.fetch_add(PerfElapsedUs(perfStartUs), std::memory_order_relaxed); });

#ifdef ENABLE_TESTS
    _dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
    const auto dbgCallbackScope = wil::scope_exit([&] noexcept { _dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
#endif

    _preCalcInProgress.store(true, std::memory_order_release);
    _preCalcStartTick.store(GetTickCount64(), std::memory_order_release);
    _preCalcCompleted.store(false, std::memory_order_release);
    _preCalcTotalBytes.store(0, std::memory_order_release);
    _preCalcFileCount.store(0, std::memory_order_release);
    _preCalcDirectoryCount.store(0, std::memory_order_release);

    _preCalcSourceBytes.clear();
    _preCalcSourceBytes.resize(_sourcePaths.size(), 0);

    std::mutex totalsMutex;
    uint64_t totalBytes = 0;
    uint64_t totalFiles = 0;
    uint64_t totalDirs  = 0;
    std::atomic<bool> acceptUpdates{true};
    std::atomic<bool> preCalcAborted{false};

    const FileSystemFlags sizeFlags = FILESYSTEM_FLAG_RECURSIVE;

    struct PreCalcWorkItem
    {
        size_t sourceIndex = 0;
        std::wstring path;
    };

    constexpr size_t kMaxPendingPreCalcDirectories = 4096u;

    const auto accumulateFinalDirectorySizeResult =
        [&](size_t sourceIndex, const PreCalcProgressCookie& progressCookie, const FileSystemDirectorySizeResult& result) noexcept
    {
        const uint64_t missingBytes = (result.totalBytes >= progressCookie.lastBytes) ? (result.totalBytes - progressCookie.lastBytes) : result.totalBytes;
        const uint64_t missingFiles = (result.fileCount >= progressCookie.lastFiles) ? (result.fileCount - progressCookie.lastFiles) : result.fileCount;
        const uint64_t missingDirs =
            (result.directoryCount >= progressCookie.lastDirs) ? (result.directoryCount - progressCookie.lastDirs) : result.directoryCount;

        if (missingBytes == 0 && missingFiles == 0 && missingDirs == 0)
        {
            return;
        }

        uint64_t snapshotBytes = 0;
        uint64_t snapshotFiles = 0;
        uint64_t snapshotDirs  = 0;
        {
            std::scoped_lock lock(totalsMutex);
            if (! acceptUpdates.load(std::memory_order_acquire))
            {
                return;
            }

            if (std::numeric_limits<uint64_t>::max() - totalBytes < missingBytes)
            {
                totalBytes = std::numeric_limits<uint64_t>::max();
            }
            else
            {
                totalBytes += missingBytes;
            }
            if (sourceIndex < _preCalcSourceBytes.size())
            {
                if (std::numeric_limits<uint64_t>::max() - _preCalcSourceBytes[sourceIndex] < missingBytes)
                {
                    _preCalcSourceBytes[sourceIndex] = std::numeric_limits<uint64_t>::max();
                }
                else
                {
                    _preCalcSourceBytes[sourceIndex] += missingBytes;
                }
            }
            if (std::numeric_limits<uint64_t>::max() - totalFiles < missingFiles)
            {
                totalFiles = std::numeric_limits<uint64_t>::max();
            }
            else
            {
                totalFiles += missingFiles;
            }
            if (std::numeric_limits<uint64_t>::max() - totalDirs < missingDirs)
            {
                totalDirs = std::numeric_limits<uint64_t>::max();
            }
            else
            {
                totalDirs += missingDirs;
            }

            snapshotBytes = totalBytes;
            snapshotFiles = totalFiles;
            snapshotDirs  = totalDirs;
        }

        UpdatePreCalcSnapshot(*this, snapshotBytes, snapshotFiles, snapshotDirs);
    };

    const auto runDirectorySizeForPath = [&](const std::wstring& path, size_t sourceIndex, FileSystemFlags flagsForPath) noexcept -> HRESULT
    {
        if (_cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire))
        {
            acceptUpdates.store(false, std::memory_order_release);
            preCalcAborted.store(true, std::memory_order_release);
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        PreCalcProgressCookie progressCookie{};
        progressCookie.totalsMutex        = &totalsMutex;
        progressCookie.totalBytes         = &totalBytes;
        progressCookie.totalFiles         = &totalFiles;
        progressCookie.totalDirs          = &totalDirs;
        progressCookie.sourceBytesByIndex = &_preCalcSourceBytes;
        progressCookie.sourceIndex        = sourceIndex;
        progressCookie.acceptUpdates      = &acceptUpdates;

        FileSystemDirectorySizeResult result{};
        result.sizeBytes     = sizeof(FileSystemDirectorySizeResult);
        const HRESULT hr     = dirOps->GetDirectorySize(path.c_str(), flagsForPath, this, &progressCookie, &result);
        const HRESULT status = FAILED(hr) ? hr : result.status;

        if (SUCCEEDED(status))
        {
            accumulateFinalDirectorySizeResult(sourceIndex, progressCookie, result);
            return S_OK;
        }

        if (status == HRESULT_FROM_WIN32(ERROR_CANCELLED))
        {
            acceptUpdates.store(false, std::memory_order_release);
            preCalcAborted.store(true, std::memory_order_release);
            return status;
        }

        const std::wstring statusText = FormatDiagnosticStatusText(status);
        LogDiagnostic(DiagnosticSeverity::Warning,
                      status,
                      L"precalc.error",
                      std::format(L"Pre-calculation failed for '{}' (hr=0x{:08X}, status='{}').", path.c_str(), static_cast<unsigned long>(status), statusText),
                      path.c_str());
        return status;
    };

    const auto enumerateChildDirectories = [&](std::wstring_view path, std::vector<std::wstring>& childDirectories) noexcept -> HRESULT
    {
        childDirectories.clear();

        wil::com_ptr<IFilesInformation> files;
        const std::wstring pathText(path);
        const HRESULT readHr = _fileSystem->ReadDirectoryInfo(pathText.c_str(), files.put());
        if (FAILED(readHr) || ! files)
        {
            return readHr;
        }

        FileInfo* head         = nullptr;
        const HRESULT bufferHr = files->GetBuffer(&head);
        if (FAILED(bufferHr))
        {
            return bufferHr;
        }
        if (head == nullptr)
        {
            return S_OK;
        }

        // The FileInfo buffer crosses the plugin trust boundary, so bound every read against its declared
        // size via the shared validators instead of trusting NextEntryOffset/FileNameSize. A malformed entry
        // fails the enumeration, and the caller safely falls back to a recursive whole-subtree size walk.
        unsigned long bufferSize = 0;
        const HRESULT sizeHr     = files->GetBufferSize(&bufferSize);
        if (FAILED(sizeHr) || bufferSize < sizeof(FileInfo))
        {
            return FAILED(sizeHr) ? sizeHr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        std::byte* const base = reinterpret_cast<std::byte*>(head);
        std::byte* const end  = base + bufferSize;

        for (FileInfo* entry = head; entry != nullptr;)
        {
            std::wstring_view name;
            const HRESULT nameHr = TryGetValidatedFileInfoName(entry, base, end, name);
            if (FAILED(nameHr))
            {
                return nameHr;
            }

            const HRESULT structuralNameHr = ValidateBridgeStructuralChildName(name);
            if (FAILED(structuralNameHr))
            {
                return structuralNameHr;
            }

            if ((entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && (entry->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0)
            {
                childDirectories.push_back(JoinFolderAndLeaf(path, name));
            }

            FileInfo* nextEntry = nullptr;
            const HRESULT advHr = AdvanceValidatedFileInfoEntry(entry, base, end, nextEntry);
            if (advHr == S_FALSE)
            {
                break;
            }
            if (FAILED(advHr))
            {
                return advHr;
            }

            entry = nextEntry;
        }

        return S_OK;
    };

    std::deque<PreCalcWorkItem> pendingWork;
    pendingWork.clear();
    for (size_t index = 0; index < _sourcePaths.size(); ++index)
    {
        pendingWork.push_back(PreCalcWorkItem{index, _sourcePaths[index].wstring()});
    }

    std::mutex workMutex;
    std::condition_variable workCv;
    size_t activeWorkers = 0;

    const auto markPreCalcAborted = [&]() noexcept
    {
        acceptUpdates.store(false, std::memory_order_release);
        preCalcAborted.store(true, std::memory_order_release);
        workCv.notify_all();
    };

    const auto dequeueWorkItem = [&](PreCalcWorkItem& workItem) noexcept -> bool
    {
        std::unique_lock lock(workMutex);
        for (;;)
        {
            if (_cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire))
            {
                return false;
            }

            if (! pendingWork.empty())
            {
                workItem = std::move(pendingWork.front());
                pendingWork.pop_front();
                ++activeWorkers;
                return true;
            }

            if (activeWorkers == 0)
            {
                return false;
            }

            workCv.wait(lock);
        }
    };

    const auto completeWorkItem = [&]() noexcept
    {
        std::scoped_lock lock(workMutex);
        if (activeWorkers > 0)
        {
            --activeWorkers;
        }
        workCv.notify_all();
    };

    const auto processWorkItem = [&](const PreCalcWorkItem& workItem) noexcept
    {
        std::vector<std::wstring> childDirectories;
        const HRESULT enumerateHr = enumerateChildDirectories(workItem.path, childDirectories);
        const bool splitChildren  = SUCCEEDED(enumerateHr);
        const HRESULT sizeHr      = runDirectorySizeForPath(workItem.path, workItem.sourceIndex, splitChildren ? FILESYSTEM_FLAG_NONE : sizeFlags);
        if (sizeHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
        {
            markPreCalcAborted();
            return;
        }

        if (! splitChildren || childDirectories.empty())
        {
            return;
        }

        std::vector<std::wstring> overflowChildren;
        overflowChildren.reserve(childDirectories.size());
        {
            std::scoped_lock lock(workMutex);
            for (auto& childPath : childDirectories)
            {
                if (pendingWork.size() < kMaxPendingPreCalcDirectories)
                {
                    pendingWork.push_back(PreCalcWorkItem{workItem.sourceIndex, std::move(childPath)});
                }
                else
                {
                    overflowChildren.push_back(std::move(childPath));
                }
            }
        }
        workCv.notify_all();

        for (const auto& overflowPath : overflowChildren)
        {
            if (_cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire))
            {
                markPreCalcAborted();
                return;
            }

            const HRESULT overflowHr = runDirectorySizeForPath(overflowPath, workItem.sourceIndex, sizeFlags);
            if (overflowHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                markPreCalcAborted();
                return;
            }
        }
    };

    const unsigned int workerCount = std::clamp(_preCalcMaxWorkers, 1u, kMaxPreCalcWorkersSetting);

    const auto workerLoop = [&]() noexcept
    {
        for (;;)
        {
            PreCalcWorkItem workItem{};
            if (! dequeueWorkItem(workItem))
            {
                return;
            }

            processWorkItem(workItem);
            completeWorkItem();
        }
    };

    if (workerCount > 1u)
    {
        std::vector<std::jthread> workers;
        workers.reserve(workerCount);
        for (unsigned int worker = 0; worker < workerCount; ++worker)
        {
            try
            {
                workers.emplace_back([&]() noexcept { workerLoop(); });
            }
            catch (const std::system_error&)
            {
                LogDiagnostic(DiagnosticSeverity::Warning,
                              HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY),
                              L"precalc.worker.start_failed",
                              std::format(L"Pre-calculation worker thread creation failed after starting {} of {} requested workers; continuing degraded.",
                                          workers.size(),
                                          workerCount));
                break;
            }
        }

        if (! workers.empty())
        {
            _preCalcWorkerCountUsed.store(static_cast<unsigned int>(workers.size()), std::memory_order_release);
        }
        else
        {
            _preCalcWorkerCountUsed.store(1u, std::memory_order_release);
            workerLoop();
        }
    }
    else
    {
        _preCalcWorkerCountUsed.store(1u, std::memory_order_release);
        workerLoop();
    }

    _preCalcInProgress.store(false, std::memory_order_release);

    uint64_t finalTotalBytes = 0;
    uint64_t finalTotalFiles = 0;
    uint64_t finalTotalDirs  = 0;
    {
        std::scoped_lock lock(totalsMutex);
        finalTotalBytes = totalBytes;
        finalTotalFiles = totalFiles;
        finalTotalDirs  = totalDirs;
    }
    UpdatePreCalcSnapshot(*this, finalTotalBytes, finalTotalFiles, finalTotalDirs);

    if (! _preCalcSkipped.load(std::memory_order_acquire) && ! _cancelled.load(std::memory_order_acquire) && ! preCalcAborted.load(std::memory_order_acquire))
    {
        // 5F early admission: publish the totals BEFORE marking pre-calc complete, and merge
        // them with std::max so the write is order-independent w.r.t. the concurrently-running
        // ExecuteOperation (which also raises these same totals with std::max under _progressMutex).
        // Setting them first guarantees that the instant a transfer thread observes
        // _preCalcCompleted, the final totals are already visible (no clobber window).
        if (finalTotalBytes > 0 || finalTotalFiles > 0 || finalTotalDirs > 0)
        {
            const unsigned long preCalcItems = static_cast<unsigned long>(std::min(finalTotalFiles + finalTotalDirs, static_cast<uint64_t>(ULONG_MAX)));
            std::scoped_lock lock(_progressMutex);
            _progressTotalBytes = (std::max)(_progressTotalBytes, finalTotalBytes);
            // Only DELETE displays pre-calc's recursive item count as the total. COPY/MOVE track
            // top-level items (PerItem) or plugin-reported items (BulkItems) inside ExecuteOperation,
            // so leave _progressTotalItems to it — matching the serial order, where ExecuteOperation's
            // per-item init overrode pre-calc's recursive count. (Raising it here too would make the
            // displayed item count depend on pre-calc enablement/overlap timing.)
            if (_operation == FILESYSTEM_DELETE)
            {
                _progressTotalItems = (std::max)(_progressTotalItems, preCalcItems);
            }
            PublishProgressCountersLocked(*this);
        }

        _preCalcCompleted.store(true, std::memory_order_release);
    }
}

HRESULT FolderWindow::FileOperationState::Task::TryStartPreCalculationThread(std::jthread& preCalcThread) noexcept
{
#ifdef ENABLE_TESTS
    g_fileOpsPreCalcThreadStartAttempts.fetch_add(1u, std::memory_order_acq_rel);
    if (g_fileOpsPreCalcThreadStartFailure.exchange(false, std::memory_order_acq_rel))
    {
        const HRESULT hr = HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        LogDiagnostic(DiagnosticSeverity::Warning,
                      hr,
                      L"precalc.thread.start_failed.selftest",
                      L"Self-test forced pre-calculation thread creation failure; falling back to serial pre-calculation.");
        return hr;
    }
#endif

    try
    {
        preCalcThread = std::jthread([this]() noexcept
        {
            [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();
            RunPreCalculation();
        });
    }
    catch (const std::system_error&)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        LogDiagnostic(DiagnosticSeverity::Warning,
                      hr,
                      L"precalc.thread.start_failed",
                      L"Pre-calculation thread creation failed; falling back to serial pre-calculation.");
        return hr;
    }

    return S_OK;
}

void FolderWindow::FileOperationState::Task::ThreadMain(std::stop_token stopToken) noexcept
{
    _stopToken                   = stopToken;
    [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();
    [[maybe_unused]] const std::stop_callback stopWake(stopToken,
                                                       [this]() noexcept
    {
        _pauseCv.notify_all();
        _conflictArbiter.cv.notify_all();
        if (_state)
        {
            ++_perf.queueNotifyAllCount;
            _state->NotifyQueueChanged();
        }
    });

    if (! _state)
    {
        return;
    }

    LogDiagnostic(DiagnosticSeverity::Debug,
                  S_OK,
                  L"task.started",
                  std::format(L"Task started (op={}, mode={}, sources={}, flags=0x{:08X}, preCalc={}, highMetadataPreCalcSuppressed={}, waitForOthers={}).",
                              OperationToString(_operation),
                              _executionMode == ExecutionMode::PerItem ? L"perItem" : L"bulkItems",
                              _sourcePaths.size(),
                              static_cast<unsigned long>(static_cast<uint32_t>(_flags)),
                              _enablePreCalc ? L"on" : L"off",
                              _preCalcSuppressedForHighMetadataCrossFs ? L"true" : L"false",
                              _waitForOthers.load(std::memory_order_acquire) ? L"true" : L"false"));

    // Mark as waiting in queue before entering (visible to UI while blocked). Use the current
    // desired start-gating state to avoid briefly showing "Waiting" for tasks that will start immediately.
    SetWaitingInQueue(_waitForOthers.load(std::memory_order_acquire));

    // Enter queue FIRST so both pre-calculation and operation respect Wait/Parallel mode
    const bool canStart = _state->EnterOperation(*this, stopToken);

    // No longer waiting in queue (either we got our turn or were cancelled)
    SetWaitingInQueue(false);

    if (! canStart)
    {
        _resultHr.store(HRESULT_FROM_WIN32(ERROR_CANCELLED), std::memory_order_release);
        _state->PostCompleted(*this);
        return;
    }

    _enteredOperationTick.store(GetTickCount64(), std::memory_order_release);
    _enteredOperation.store(true, std::memory_order_release);

    // 5F early admission: run pre-calculation concurrently with the transfer so bytes start moving
    // immediately instead of after the full recursive scan completes. Totals/ETA stay "estimating"
    // (gated on _preCalcCompleted) until pre-calc publishes them, then reconcile in place (both
    // producers raise the shared totals with std::max under _progressMutex). ONLY COPY qualifies: its
    // source is strictly read-only during the transfer, so concurrent enumeration is safe. MOVE and
    // DELETE modify/remove the source, which would race pre-calc's enumeration (it would size a tree
    // that is being deleted), so they keep the serial pre-calc-then-execute order.
    const bool useEarlyAdmission = (_operation == FILESYSTEM_COPY && _enablePreCalc && ! _preCalcSkipped.load(std::memory_order_acquire));

    const auto cancelBeforeExecute = [&]() noexcept -> bool
    {
        if (! _cancelled.load(std::memory_order_acquire))
        {
            return false;
        }

        _enteredOperation.store(false, std::memory_order_release);
        _enteredOperationTick.store(0, std::memory_order_release);
        _state->LeaveOperation();
        _resultHr.store(HRESULT_FROM_WIN32(ERROR_CANCELLED), std::memory_order_release);
        _state->PostCompleted(*this);
        return true;
    };

    std::jthread preCalcThread;
    if (useEarlyAdmission)
    {
        const HRESULT preCalcThreadHr = TryStartPreCalculationThread(preCalcThread);
        if (FAILED(preCalcThreadHr))
        {
            RunPreCalculation();
            if (cancelBeforeExecute())
            {
                return;
            }
        }
    }
    else
    {
        RunPreCalculation();

        // Serial path: pre-calc is complete. Honor a cancel that arrived during it before touching
        // the source. (For early admission the transfer has already started, so its cancellation
        // flows through ExecuteOperation's result instead of this fast exit.)
        if (cancelBeforeExecute())
        {
            return;
        }
    }

    const HRESULT hr = ExecuteOperation();

    // Pre-calc must finish before we read its results or unwind captured task state.
    if (preCalcThread.joinable())
    {
        preCalcThread.join();
    }
    _resultHr.store(hr, std::memory_order_release);

    const ULONGLONG afterPreCalcTick = GetTickCount64();
    if (Debug::Perf::IsCaptureEnabled())
    {
        const ULONGLONG preStartTick = _preCalcStartTick.load(std::memory_order_acquire);
        if (preStartTick > 0)
        {
            const ULONGLONG elapsedMs      = (afterPreCalcTick >= preStartTick) ? (afterPreCalcTick - preStartTick) : 0;
            const uint64_t preCalcUs       = _perf.preCalcUs.load(std::memory_order_acquire);
            const uint64_t durationUs      = (preCalcUs > 0) ? preCalcUs : static_cast<uint64_t>(elapsedMs) * 1000ull;
            const HRESULT preCalcHr        = _cancelled.load(std::memory_order_acquire) ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                                        : (_preCalcSkipped.load(std::memory_order_acquire) ? S_FALSE : S_OK);
            const uint64_t bytes           = _preCalcTotalBytes.load(std::memory_order_acquire);
            const uint64_t items           = static_cast<uint64_t>(_preCalcFileCount.load(std::memory_order_acquire)) +
                                             static_cast<uint64_t>(_preCalcDirectoryCount.load(std::memory_order_acquire));
            const size_t sourceCount       = _sourcePaths.size();
            const unsigned int workersUsed = _preCalcWorkerCountUsed.load(std::memory_order_acquire);
            const std::wstring detail      = std::format(L"id={} op={} sources={} workers={} configuredWorkers={}",
                                                         _taskId,
                                                         OperationToString(_operation),
                                                         sourceCount,
                                                         workersUsed,
                                                         _preCalcMaxWorkers);
            Debug::Perf::Emit(L"FileOps.PreCalc", detail, durationUs, bytes, items, preCalcHr);
        }
    }

    {
        const ULONGLONG preStartTick = _preCalcStartTick.load(std::memory_order_acquire);
        if (preStartTick > 0)
        {
            const ULONGLONG elapsedMs = (afterPreCalcTick >= preStartTick) ? (afterPreCalcTick - preStartTick) : 0;
            const HRESULT preCalcHr   = _cancelled.load(std::memory_order_acquire) ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                                   : (_preCalcSkipped.load(std::memory_order_acquire) ? S_FALSE : S_OK);
            const uint64_t bytes      = _preCalcTotalBytes.load(std::memory_order_acquire);
            const unsigned long files = _preCalcFileCount.load(std::memory_order_acquire);
            const unsigned long dirs  = _preCalcDirectoryCount.load(std::memory_order_acquire);
            const bool skipped        = _preCalcSkipped.load(std::memory_order_acquire);
            LogDiagnostic(DiagnosticSeverity::Debug,
                          preCalcHr,
                          L"precalc.result",
                          std::format(L"Pre-calculation finished (hr=0x{:08X}, elapsedMs={}, bytes={:L}, files={:L}, dirs={:L}, skipped={}).",
                                      static_cast<unsigned long>(preCalcHr),
                                      elapsedMs,
                                      bytes,
                                      files,
                                      dirs,
                                      skipped ? L"true" : L"false"));
        }
    }

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

        const PublishedProgressSnapshot progressSnapshot   = LoadPublishedProgressSnapshot(*this);
        const auto& perfStats                              = _perf;
        const uint64_t bridgeDirectoryEnsureCount          = _bridgeDirectoryEnsureCount.load(std::memory_order_acquire);
        const uint64_t bridgeFileAdmissionCount            = _bridgeFileAdmissionCount.load(std::memory_order_acquire);
        const uint64_t bridgeFileStartedBeforeProducerDone = _bridgeFileStartedBeforeProducerDone.load(std::memory_order_acquire);
        const uint64_t bridgeAdmissionMaxQueueDepth        = _bridgeAdmissionMaxQueueDepth.load(std::memory_order_acquire);
        const uint64_t preCalcUs                           = perfStats.preCalcUs.load(std::memory_order_acquire);
        const uint64_t preCalcCallbackCount                = perfStats.preCalcCallbackCount.load(std::memory_order_acquire);
        const uint64_t preCalcCallbackUs                   = perfStats.preCalcCallbackUs.load(std::memory_order_acquire);
        const uint64_t preCalcLockWaitUs                   = perfStats.preCalcLockWaitUs.load(std::memory_order_acquire);
        std::array<ProgressStreamPerf, kMaxInFlightFiles> progressStreamPerf{};
        size_t progressStreamPerfCount = 0;
        {
            std::scoped_lock lock(_progressStreamPerfMutex);
            progressStreamPerf      = _progressStreamPerf;
            progressStreamPerfCount = _progressStreamPerfCount;
        }
        uint64_t progressStreamGapCount    = 0;
        uint64_t progressStreamGapMs       = 0;
        uint64_t progressStreamGapBytes    = 0;
        uint64_t progressStreamMaxGapMs    = 0;
        uint64_t progressStreamMaxGapBytes = 0;
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
            L"bridgeWriteUs={} bridgeReaderWaitUs={} bridgeWriterWaitUs={} preCalcUs={} progressUs={} firstProgressMs={} maxProgressGapMs={} "
            L"maxProgressGapBytes={} bridgeDirs={} bridgeFiles={} bridgeEarlyFiles={} bridgeQueueMax={} "
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
            preCalcUs,
            perfStats.progressCallbackUs.load(std::memory_order_acquire),
            perfStats.progressFirstCallbackDelayMs,
            progressStreamMaxGapMs,
            progressStreamMaxGapBytes,
            bridgeDirectoryEnsureCount,
            bridgeFileAdmissionCount,
            bridgeFileStartedBeforeProducerDone,
            bridgeAdmissionMaxQueueDepth,
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
        Debug::Perf::Emit(L"FileOps.Bridge.DirectoryEnsureCount", L"", 0u, bridgeDirectoryEnsureCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.FileAdmissionCount", L"", 0u, bridgeFileAdmissionCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.FileStartedBeforeProducerDone", L"", 0u, bridgeFileStartedBeforeProducerDone, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.AdmissionMaxQueueDepth", L"", 0u, bridgeAdmissionMaxQueueDepth, 0u, hr);
        Debug::Perf::Emit(L"FileOps.PreCalc.TotalUs", L"", preCalcUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(L"FileOps.PreCalc.CallbackCount", L"", 0u, preCalcCallbackCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.PreCalc.CallbackUs", L"", preCalcCallbackUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(L"FileOps.PreCalc.LockWaitUs", L"", preCalcLockWaitUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
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
        Debug::Perf::Emit(
            L"FileOps.Scheduler.WaitForWorkUs", L"", perfStats.schedulerWaitForWorkUs.load(std::memory_order_acquire), 0u, 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.Scheduler.ProcessIndexUs", L"", perfStats.schedulerProcessIndexUs.load(std::memory_order_acquire), 0u, 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.Scheduler.DequeueAttempts", L"", 0u, perfStats.schedulerDequeueAttempts.load(std::memory_order_acquire), 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.Scheduler.DequeueSuccess", L"", 0u, perfStats.schedulerDequeueSuccess.load(std::memory_order_acquire), 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.Conflict.WaitUs", L"", perfStats.conflictWaitUs, progressSnapshot.completedBytes, progressSnapshot.itemCompletedCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Conflict.MetadataUs",
                          L"",
                          perfStats.conflictMetadataUs.load(std::memory_order_acquire),
                          perfStats.conflictPromptCount,
                          0u,
                          hr);
        Debug::Perf::Emit(L"FileOps.Conflict.ConvergenceWaitUs",
                          L"",
                          perfStats.conflictConvergenceWaitUs,
                          progressSnapshot.completedBytes,
                          progressSnapshot.itemCompletedCallbackCount,
                          hr);
        Debug::Perf::Emit(L"FileOps.Conflict.PromptCount", L"", 0u, perfStats.conflictPromptCount, 0u, hr);
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
    _state->LeaveOperation();
    _state->PostCompleted(*this);
}

void FolderWindow::FileOperationState::Task::RequestCancel() noexcept
{
    {
        ULONGLONG expected = 0;
        _cancelRequestedTick.compare_exchange_strong(expected, GetTickCount64(), std::memory_order_release);
    }
    _cancelled.store(true, std::memory_order_release);
    {
        std::scoped_lock lock(_pauseMutex);
        _paused.store(false, std::memory_order_release);
    }
    _pauseCv.notify_all();

    if (_conflictArbiter.decisionEvent)
    {
        static_cast<void>(SetEvent(_conflictArbiter.decisionEvent.get()));
    }

    _conflictArbiter.cv.notify_all();

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

void FolderWindow::FileOperationState::Task::InitializeFileSystemOptions(FileSystemOptions& options) const noexcept
{
    options                              = {};
    options.sizeBytes                    = sizeof(FileSystemOptions);
    options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    options.copyMoveMaxConcurrency       = 0;
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
        if (_started.load(std::memory_order_acquire))
        {
            return;
        }

        _waitForOthers.store(wait, std::memory_order_release);
        ++_perf.queueNotifyAllCount;
        state->NotifyQueueChanged();
        return;
    }

    if (_started.load(std::memory_order_acquire))
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
    if (! _conflictArbiter.prompt.active)
    {
        return;
    }

    _conflictArbiter.prompt.applyToAllChecked = ! _conflictArbiter.prompt.applyToAllChecked;
}

void FolderWindow::FileOperationState::Task::SubmitConflictDecision(ConflictAction action, bool applyToAllChecked) noexcept
{
    {
        std::scoped_lock lock(_conflictArbiter.mutex);
        if (! _conflictArbiter.prompt.active)
        {
            return;
        }

        _conflictArbiter.decisionAction     = action;
        _conflictArbiter.decisionApplyToAll = action != ConflictAction::Retry && applyToAllChecked;
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

bool FolderWindow::FileOperationState::Task::IsQueuePaused() const noexcept
{
    return _queuePaused.load(std::memory_order_acquire);
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
        if (externalStop != nullptr && externalStop->load(std::memory_order_acquire))
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
            return ! waitingForConflict || stillPaused || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested() ||
                   (externalStop != nullptr && externalStop->load(std::memory_order_acquire));
        });
        _perf.conflictConvergenceWaitUs += PerfElapsedUs(waitStartUs);
    }
}

void FolderWindow::FileOperationState::Task::WakePauseWaiters() noexcept
{
    _pauseCv.notify_all();
    _conflictArbiter.cv.notify_all();
}

void FolderWindow::FileOperationState::Task::WaitWhilePreCalcPaused() noexcept
{
    const bool shouldPause = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
    if (! shouldPause)
    {
        return;
    }

    const uint64_t waitStartUs = PerfNowUs();
    std::unique_lock lock(_pauseMutex);
    _pauseCv.wait(lock,
                  [&]
    {
        const bool stillPaused = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
        return ! stillPaused || _cancelled.load(std::memory_order_acquire) || _preCalcSkipped.load(std::memory_order_acquire) || _stopToken.stop_requested();
    });
    _perf.pauseWaitUs.fetch_add(PerfElapsedUs(waitStartUs), std::memory_order_relaxed);
}

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

    WaitWhilePaused();
    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    _observedSkipAction.store(false, std::memory_order_release);
    _started.store(true, std::memory_order_release);
    _operationStartTick.store(GetTickCount64(), std::memory_order_release);

#ifdef ENABLE_TESTS
    _dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
    const auto dbgCallbackScope = wil::scope_exit([&] noexcept { _dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
#endif

#ifdef ENABLE_TESTS
    _dbgConfiguredMaxConcurrency = DeterminePerItemMaxConcurrency(_fileSystem, _sourcePaths, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles));
    _dbgConfiguredMaxConcurrency = std::max(1u, _dbgConfiguredMaxConcurrency);
    _dbgSingleInFlightStartTick  = 0;
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

    if (_executionMode == ExecutionMode::PerItem)
    {
        wil::com_ptr<IFileSystemIO> fileSystemIo;
        static_cast<void>(_fileSystem->QueryInterface(IID_PPV_ARGS(fileSystemIo.addressof())));
        wil::com_ptr<IFileSystemDirectoryOperations> fileSystemDirOps;
        static_cast<void>(_fileSystem->QueryInterface(IID_PPV_ARGS(fileSystemDirOps.addressof())));

        const bool useCrossFileSystemBridge = (_destinationFileSystem != nullptr) && (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE);

        wil::com_ptr<IFileSystemIO> destinationFileSystemIo;
        wil::com_ptr<IFileSystemDirectoryOperations> destinationDirOps;
        if (useCrossFileSystemBridge)
        {
            static_cast<void>(_destinationFileSystem->QueryInterface(IID_PPV_ARGS(destinationFileSystemIo.addressof())));
            static_cast<void>(_destinationFileSystem->QueryInterface(IID_PPV_ARGS(destinationDirOps.addressof())));

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

        // The cross-filesystem bridge is only used for copy/move, so hoisting these capability lookups is safe:
        // per-item conflict handling may tweak `itemFlags`, but delete-only flag-sensitive keys are never consulted here.
        const unsigned int bridgeSourceMaxConcurrencyBudget =
            useCrossFileSystemBridge
                ? DeterminePerItemMaxConcurrency(_fileSystem, _sourcePaths, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles))
                : 1u;
        const unsigned int bridgeDestinationMaxConcurrencyBudget =
            useCrossFileSystemBridge
                ? DeterminePerItemMaxConcurrency(_destinationFileSystem, destinationFolder, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles))
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

        ReparsePointPolicy reparsePointPolicy = ReparsePointPolicy::CopyReparse;
        if (const auto policyOpt = TryGetReparsePointPolicyFromFileSystem(_fileSystem); policyOpt.has_value())
        {
            reparsePointPolicy = policyOpt.value();
        }
        else if (_folderWindow && _folderWindow->_settings)
        {
            const std::wstring& sourcePluginId =
                _sourcePane == FolderWindow::Pane::Left ? _folderWindow->_leftPane.pluginId : _folderWindow->_rightPane.pluginId;
            if (! sourcePluginId.empty())
            {
                reparsePointPolicy = GetReparsePointPolicyFromSettings(*_folderWindow->_settings, sourcePluginId);
            }
        }
#ifdef ENABLE_TESTS
        const int reparsePolicyOverride = g_fileOpsBridgeReparsePolicyOverride.load(std::memory_order_acquire);
        if (reparsePolicyOverride >= static_cast<int>(ReparsePointPolicy::CopyReparse) &&
            reparsePolicyOverride <= static_cast<int>(ReparsePointPolicy::Skip))
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
            DeterminePerItemMaxConcurrency(_fileSystem, _sourcePaths, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles));
        _perItemMaxConcurrencyBudget = std::max(1u, _perItemMaxConcurrencyBudget);
        if (useCrossFileSystemBridge)
        {
            const unsigned int destinationMaxConcurrencyBudget =
                DeterminePerItemMaxConcurrency(_destinationFileSystem, destinationFolder, _operation, _flags, static_cast<unsigned int>(kMaxInFlightFiles));
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

        const bool canUsePreCalcBytes = _preCalcCompleted.load(std::memory_order_acquire) && _preCalcSourceBytes.size() == _sourcePaths.size();

        bool hadSkippedItems = false;

        if ((_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE) && destinationFolder.empty())
        {
            return E_INVALIDARG;
        }

        const std::wstring destinationFolderText = destinationFolder.native();
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
                                                      FileSystemOptions* options,
                                                      void* cookie) noexcept override
            {
                if (callbackMutex != nullptr)
                {
                    std::scoped_lock lock(*callbackMutex);
                    return task.FileSystemIssue(operationType, sourcePath, destinationPath, status, action, options, cookie);
                }
                return task.FileSystemIssue(operationType, sourcePath, destinationPath, status, action, options, cookie);
            }
        };

        struct CrossFileSystemBridge
        {
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
            IFileSystemDirectoryOperations* destinationDirOps  = nullptr;
            unsigned int sourcePluginMaxConcurrencyBudget      = 1;
            unsigned int destinationPluginMaxConcurrencyBudget = 1;
            FileSystemFlags flags                              = FILESYSTEM_FLAG_NONE;
            void* cookie                                       = nullptr;
            DWORD sourceRootAttributesHint                     = 0;
            ReparsePointPolicy reparsePointPolicy              = ReparsePointPolicy::CopyReparse;
            bool destinationUsesOrdinalIgnoreCaseComponents     = false;
            bool destinationUsesWindowsChildNameRules           = false;

            // Total bytes is best-effort: if unknown, keep 0.
            uint64_t totalBytes                         = 0;
            uint64_t completedBytes                     = 0;
            unsigned long skippedDirectoryReparseCount  = 0;
            unsigned long skippedFileReparseCount       = 0;
            bool rootDirectoryReparseSkipped            = false;
            bool unsupportedDirectoryReparseEncountered = false;
            // Per-file conflicts answered Skip: the file never reached the destination, so for a
            // MOVE the source stays authoritative (no delete) and the transfer ends PARTIAL.
            std::atomic<uint64_t> skippedFileConflictCount{0};

            enum class CopiedEntryKind : unsigned char
            {
                File,
                Directory,
            };

            struct CopiedEntry final
            {
                CopiedEntryKind kind = CopiedEntryKind::File;
                std::wstring sourcePath;
                std::wstring destinationPath;
                uint64_t sizeBytes = 0;
                bool hasKnownSize  = false;
                uint64_t contentHash = 0;
                bool hasContentHash  = false;
                bool hasBasicInfo  = false;
                FileSystemBasicInformation basicInfo{};
            };

            struct MoveSourceCleanupStats final
            {
                bool anyDeleted = false;
                bool anySkipped = false;
            };

            std::mutex callbackMutex;
            std::mutex throttleMutex;
            std::mutex copiedEntriesMutex;
            std::unordered_map<std::wstring, CopiedEntry> copiedEntries;
            uint64_t copiedEntriesPeakCount = 0;
            std::atomic<uint64_t> bandwidthLimitBytesPerSecond{0};

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
            unsigned long bufferBytes = 0;
            HRESULT bufferAllocationHr = S_OK;
            DWORD progressPeriodMs    = 200u;
            uint32_t transferLatencyClass = FILESYSTEM_TRANSFER_LATENCY_UNKNOWN;
            uint32_t transferHintFlags    = FILESYSTEM_TRANSFER_HINT_NONE;

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
                                  ReparsePointPolicy reparsePointPolicyIn) noexcept
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
                  totalBytes(totalBytesIn)
            {
                destinationUsesOrdinalIgnoreCaseComponents = UsesOrdinalIgnoreCaseComponents(destinationFs, task._destinationPluginId);
                destinationUsesWindowsChildNameRules = destinationUsesOrdinalIgnoreCaseComponents ||
                                                       NavigationLocation::EqualsNoCase(task._destinationPluginId, L"builtin/file-system") ||
                                                       NavigationLocation::EqualsNoCase(task._destinationPluginId, L"builtin/file-system-dummy");
                const uint64_t initialBandwidth      = task._desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                options.sizeBytes                    = sizeof(FileSystemOptions);
                options.bandwidthLimitBytesPerSecond = initialBandwidth;
                options.copyMoveMaxConcurrency       = std::min(sourcePluginMaxConcurrencyBudget, destinationPluginMaxConcurrencyBudget);
                bandwidthLimitBytesPerSecond.store(initialBandwidth, std::memory_order_release);

                const AdaptiveBridgeTuning tuning = ResolveAdaptiveCrossFsBridgeTuning(
                    task._crossFsBridgeBufferBytes, sourceFs, rootSourcePathIn, destinationFs, rootDestinationPathIn, task._operation);
                bufferBytes          = tuning.bufferBytes;
                progressPeriodMs     = tuning.progressPeriodMs;
                transferLatencyClass = tuning.latencyClass;
                transferHintFlags    = tuning.flags;
                task._resolvedCrossFsBridgeBufferBytes.store(bufferBytes, std::memory_order_release);
#ifdef ENABLE_TESTS
                if (task._operation == FILESYSTEM_MOVE)
                {
                    g_fileOpsBridgeMoveManifestCurrentEntries.store(0u, std::memory_order_release);
                    g_fileOpsBridgeMoveManifestPeakEntries.store(0u, std::memory_order_release);
                }
#endif
                const uint64_t reservationBytes = static_cast<uint64_t>(bufferBytes) * 2ull;
                if (! bufferBudgetLease.Acquire(reservationBytes, task._cancelled, task._stopToken))
                {
                    bufferAllocationHr = (task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested())
                                             ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                             : E_OUTOFMEMORY;
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

            [[nodiscard]] bool CancelRequested() const noexcept
            {
                return task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested();
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

            [[nodiscard]] HRESULT ValidateAndRegisterChildName(const std::wstring_view name,
                                                               std::set<std::wstring>& exactNames,
                                                               std::set<std::wstring, BridgeOrdinalIgnoreCaseLess>& foldedNames) const noexcept
            {
                const HRESULT hr = ValidateChildNameForDestination(name);
                if (FAILED(hr))
                {
                    return hr;
                }

                const std::wstring ownedName(name);
                if (! exactNames.insert(ownedName).second)
                {
                    return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                }
                if (destinationUsesOrdinalIgnoreCaseComponents && ! foldedNames.insert(ownedName).second)
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

            [[nodiscard]] std::wstring MakeTempDestinationPath(std::wstring_view destinationPath, uint64_t progressStreamId) const noexcept
            {
                // Best-effort: keep atomic commit semantics for cross-filesystem transfers by writing to a temp name first.
                // The name carries 128 bits of CSPRNG entropy so a local attacker cannot pre-create
                // or race the staging file (PID/TID/tick names are predictable); the stream id stays
                // for diagnostic correlation only.
                wchar_t suffix[96]{};
                constexpr size_t suffixMax = (sizeof(suffix) / sizeof(suffix[0])) - 1u;

                uint64_t random[2]{};
                if (! BCRYPT_SUCCESS(BCryptGenRandom(nullptr, reinterpret_cast<PUCHAR>(random), sizeof(random), BCRYPT_USE_SYSTEM_PREFERRED_RNG)))
                {
                    // Extremely unlikely; degrade to the legacy name rather than blocking the copy.
                    random[0] = (static_cast<uint64_t>(GetCurrentProcessId()) << 32) | GetCurrentThreadId();
                    random[1] = GetTickCount64();
                }

                const auto r         = std::format_to_n(suffix,
                                                        suffixMax,
                                                        L".rs_tmp_{:016X}{:016X}_{:X}",
                                                        static_cast<unsigned long long>(random[0]),
                                                        static_cast<unsigned long long>(random[1]),
                                                        progressStreamId);
                const size_t written = (r.size < suffixMax) ? static_cast<size_t>(r.size) : suffixMax;
                suffix[written]      = L'\0';

                std::wstring temp(destinationPath);
                temp.append(suffix);
                return temp;
            }

            void BestEffortDeleteTempFile(const std::wstring& tempPath, const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                if (tempPath.empty())
                {
                    return;
                }

                const FileSystemFlags cleanupFlags = static_cast<FileSystemFlags>(static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE) |
                                                                                  static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));

                const HRESULT hrDelete = destinationFs.DeleteItem(tempPath.c_str(), cleanupFlags, nullptr, nullptr, nullptr);
                if (FAILED(hrDelete) && hrDelete != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && hrDelete != HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) &&
                    hrDelete != HRESULT_FROM_WIN32(ERROR_INVALID_NAME))
                {
                    Debug::Warning(L"CrossFileSystemBridge: failed to delete temp file '{}' (hr={:#x})", tempPath, static_cast<unsigned long>(hrDelete));
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                       hrDelete,
                                       L"bridge.temp.cleanup",
                                       L"Failed to remove temporary destination file after a failed transfer.",
                                       sourcePath,
                                       destinationPath);
                }
            }

            [[nodiscard]] HRESULT PromoteTempToFinalPath(const std::wstring& tempPath,
                                                         const std::wstring& destinationPath,
                                                         bool overwriteGranted,
                                                         bool replaceReadOnlyGranted) noexcept
            {
                // The destination conflict was already resolved (per-file or globally); carry the
                // grant into the promote so the final rename can replace the existing file.
                FileSystemFlags promoteFlags = flags;
                if (overwriteGranted)
                {
                    promoteFlags = static_cast<FileSystemFlags>(static_cast<uint32_t>(promoteFlags) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE));
                }
                if (replaceReadOnlyGranted)
                {
                    promoteFlags =
                        static_cast<FileSystemFlags>(static_cast<uint32_t>(promoteFlags) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));
                }

                BridgeCallback callback(task, &callbackMutex);
                return destinationFs.MoveItem(tempPath.c_str(), destinationPath.c_str(), promoteFlags, nullptr, &callback, cookie);
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
                const std::wstring detail =
                    std::format(L"source={} destination={} bytes={} bufferBytes={} progressPeriodMs={} latencyClass={} hintFlags=0x{:X} progressCalls={} readerWaitUs={} writerWaitUs={} readUs={} writeUs={}",
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
                                                             bool& replaceReadOnlyGranted) noexcept
            {
                FileSystemIssueAction action   = FileSystemIssueAction::Cancel;
                FileSystemOptions issueOptions{};
                {
                    std::scoped_lock lock(callbackMutex);
                    issueOptions = options;
                }
                issueOptions.sizeBytes         = sizeof(FileSystemOptions);
#ifdef ENABLE_TESTS
                task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
                const auto dbgCallbackScope = wil::scope_exit([&] noexcept { task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
#endif
                const HRESULT issueHr =
                    task.FileSystemIssue(task.GetOperation(), sourcePath.c_str(), destinationPath.c_str(), issueStatus, &action, &issueOptions, cookie);
                if (FAILED(issueHr))
                {
                    return issueHr;
                }

                switch (action)
                {
                    case FileSystemIssueAction::Overwrite: overwriteGranted = true; return S_OK;
                    case FileSystemIssueAction::ReplaceReadOnly:
                        overwriteGranted       = true;
                        replaceReadOnlyGranted = true;
                        return S_OK;
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

            HRESULT ResolveDirectoryDestinationFileCollision(const std::wstring& sourcePath,
                                                             const std::wstring& destinationPath,
                                                             unsigned long destinationAttributes) noexcept
            {
                bool overwriteGranted       = (static_cast<uint32_t>(flags) & static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE)) != 0u;
                bool replaceReadOnlyGranted = (static_cast<uint32_t>(flags) & static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY)) != 0u;

                const bool readonlyCollision = (destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0;
                if (! overwriteGranted || (readonlyCollision && ! replaceReadOnlyGranted))
                {
                    const HRESULT issueStatus = HRESULT_FROM_WIN32(readonlyCollision ? ERROR_ACCESS_DENIED : ERROR_ALREADY_EXISTS);
                    const HRESULT promptHr    = PromptDestinationCollision(sourcePath, destinationPath, issueStatus, overwriteGranted, replaceReadOnlyGranted);
                    if (promptHr == S_FALSE)
                    {
                        RecordSkippedDestinationCollision(sourcePath,
                                                          destinationPath,
                                                          issueStatus,
                                                          L"bridge.directoryConflict.skip",
                                                          L"Destination blocks directory creation; skipped on user request.");
                        return S_FALSE;
                    }
                    if (FAILED(promptHr))
                    {
                        return promptHr;
                    }
                }

                FileSystemFlags deleteFlags = FILESYSTEM_FLAG_NONE;
                if (replaceReadOnlyGranted)
                {
                    deleteFlags =
                        static_cast<FileSystemFlags>(static_cast<uint32_t>(deleteFlags) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));
                }

                BridgeCallback callback(task, &callbackMutex);
                return destinationFs.DeleteItem(destinationPath.c_str(), deleteFlags, nullptr, &callback, cookie);
            }

            HRESULT EnsureDestinationDirectory(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                if (! destinationDirOps)
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }

                for (unsigned int attempt = 0; attempt < 3u; ++attempt)
                {
                    unsigned long attributes = 0;
                    const HRESULT hrAttr     = destinationIo.GetAttributes(destinationPath.c_str(), &attributes);
                    if (SUCCEEDED(hrAttr))
                    {
                        if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                        {
                            RecordCopiedDirectory(sourcePath, destinationPath);
                            return S_OK;
                        }

                        const HRESULT hrResolve = ResolveDirectoryDestinationFileCollision(sourcePath, destinationPath, attributes);
                        if (hrResolve == S_FALSE || FAILED(hrResolve))
                        {
                            return hrResolve;
                        }
                        continue;
                    }

#ifdef ENABLE_TESTS
                    MaybeInjectBridgeCreateDirectoryRaceForSelfTest(destinationIo, destinationPath);
#endif

                    const HRESULT hrCreate = destinationDirOps->CreateDirectory(destinationPath.c_str());
                    if (SUCCEEDED(hrCreate))
                    {
                        RecordCopiedDirectory(sourcePath, destinationPath);
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

            void MarkReparseSkipped(const std::wstring& sourcePath,
                                    const std::wstring& destinationPath,
                                    bool isDirectory,
                                    bool isRoot) noexcept
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

                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                                   L"bridge.reparse.skip",
                                   isDirectory ? (isRoot ? L"Skipped root directory reparse point by policy." : L"Skipped directory reparse point by policy.")
                                               : (isRoot ? L"Skipped root file reparse point by policy." : L"Skipped file reparse point by policy."),
                                   sourcePath,
                                   destinationPath);

                const uint64_t callCompleted = (completedBytesAtomic != nullptr) ? completedBytesAtomic->load(std::memory_order_acquire) : completedBytes;
                static_cast<void>(ReportProgress(sourcePath, destinationPath, 0, 0, callCompleted, 0));
            }

            [[nodiscard]] static bool IsMissingPathHr(HRESULT hr) noexcept
            {
                return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ||
                       hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
            }

            void RecordCopiedDirectory(const std::wstring& sourcePath, const std::wstring& destinationPath)
            {
                if (task._operation != FILESYSTEM_MOVE)
                {
                    return;
                }

                CopiedEntry entry{};
                entry.kind                = CopiedEntryKind::Directory;
                entry.sourcePath          = sourcePath;
                entry.destinationPath     = destinationPath;
                entry.basicInfo.sizeBytes = sizeof(FileSystemBasicInformation);
                if (SUCCEEDED(sourceIo.GetFileBasicInformation(sourcePath.c_str(), &entry.basicInfo)))
                {
                    entry.hasBasicInfo = true;
                }

                std::scoped_lock lock(copiedEntriesMutex);
                copiedEntries[sourcePath] = std::move(entry);
                copiedEntriesPeakCount    = std::max(copiedEntriesPeakCount, static_cast<uint64_t>(copiedEntries.size()));
#ifdef ENABLE_TESTS
                const uint64_t currentCount = static_cast<uint64_t>(copiedEntries.size());
                g_fileOpsBridgeMoveManifestCurrentEntries.store(currentCount, std::memory_order_release);
                AtomicMax(g_fileOpsBridgeMoveManifestPeakEntries, currentCount);
#endif
            }

            void RecordCopiedFile(const std::wstring& sourcePath,
                                  const std::wstring& destinationPath,
                                  uint64_t sizeBytes,
                                  uint64_t contentHash,
                                  const FileSystemBasicInformation& basicInfo,
                                  bool hasBasicInfo)
            {
                if (task._operation != FILESYSTEM_MOVE)
                {
                    return;
                }

                CopiedEntry entry{};
                entry.kind            = CopiedEntryKind::File;
                entry.sourcePath      = sourcePath;
                entry.destinationPath = destinationPath;
                entry.sizeBytes       = sizeBytes;
                entry.hasKnownSize    = true;
                entry.contentHash     = contentHash;
                entry.hasContentHash  = true;
                entry.basicInfo       = basicInfo;
                entry.hasBasicInfo    = hasBasicInfo;

                std::scoped_lock lock(copiedEntriesMutex);
                copiedEntries[sourcePath] = std::move(entry);
                copiedEntriesPeakCount    = std::max(copiedEntriesPeakCount, static_cast<uint64_t>(copiedEntries.size()));
#ifdef ENABLE_TESTS
                const uint64_t currentCount = static_cast<uint64_t>(copiedEntries.size());
                g_fileOpsBridgeMoveManifestCurrentEntries.store(currentCount, std::memory_order_release);
                AtomicMax(g_fileOpsBridgeMoveManifestPeakEntries, currentCount);
#endif
            }

            [[nodiscard]] bool TakeCopiedEntry(const std::wstring& sourcePath, CopiedEntry& entry) noexcept
            {
                {
                    std::scoped_lock lock(copiedEntriesMutex);
                    auto node = copiedEntries.extract(sourcePath);
                    if (node.empty())
                    {
                        return false;
                    }

                    entry = std::move(node.mapped());
#ifdef ENABLE_TESTS
                    g_fileOpsBridgeMoveManifestCurrentEntries.store(static_cast<uint64_t>(copiedEntries.size()), std::memory_order_release);
#endif
                }

#ifdef ENABLE_TESTS
                MaybePauseAfterBridgeMoveManifestTakeForSelfTest();
#endif
                return true;
            }

            [[nodiscard]] static bool IsTransientCleanupHr(HRESULT hr) noexcept
            {
                return hr == HRESULT_FROM_WIN32(ERROR_BUSY) || hr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED) ||
                       hr == HRESULT_FROM_WIN32(ERROR_LOCK_VIOLATION) || hr == HRESULT_FROM_WIN32(ERROR_NETWORK_BUSY) ||
                       hr == HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE) || hr == HRESULT_FROM_WIN32(ERROR_RETRY) ||
                       hr == HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT) || hr == HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) ||
                       hr == HRESULT_FROM_WIN32(ERROR_TIMEOUT) || hr == HRESULT_FROM_WIN32(ERROR_UNEXP_NET_ERR);
            }

            [[nodiscard]] HRESULT RetryTransientCleanupProbe(const std::function<HRESULT()>& probe) noexcept
            {
                HRESULT hr = S_OK;
                for (unsigned int attempt = 0; attempt < 3u; ++attempt)
                {
                    hr = probe();
                    if (SUCCEEDED(hr) || IsMissingPathHr(hr) || ! IsTransientCleanupHr(hr) || CancelRequested())
                    {
                        return hr;
                    }

                    SleepResponsive(25u * (attempt + 1u));
                }

                return hr;
            }

            [[nodiscard]] bool BasicFileInfoStillMatchesSource(const CopiedEntry& entry) noexcept
            {
                if (! entry.hasBasicInfo)
                {
                    return true;
                }

                FileSystemBasicInformation current{};
                current.sizeBytes = sizeof(FileSystemBasicInformation);
                if (FAILED(sourceIo.GetFileBasicInformation(entry.sourcePath.c_str(), &current)))
                {
                    return false;
                }

                return current.creationTime == entry.basicInfo.creationTime && current.lastWriteTime == entry.basicInfo.lastWriteTime &&
                       current.attributes == entry.basicInfo.attributes;
            }

            static constexpr uint64_t HashOffsetBasis() noexcept
            {
                return 14695981039346656037ull;
            }

            static void HashBytes(uint64_t& hash, const std::byte* bytes, size_t byteCount) noexcept
            {
                constexpr uint64_t kPrime = 1099511628211ull;
                for (size_t index = 0; index < byteCount; ++index)
                {
                    hash ^= static_cast<uint64_t>(std::to_integer<unsigned char>(bytes[index]));
                    hash *= kPrime;
                }
            }

            [[nodiscard]] HRESULT ReaderMatchesHash(IFileReader& reader,
                                                    uint64_t sizeBytes,
                                                    uint64_t expectedHash,
                                                    std::byte* hashBuffer,
                                                    unsigned long hashBufferBytes,
                                                    bool& matches) noexcept
            {
                matches = false;
                if (hashBuffer == nullptr || hashBufferBytes == 0u)
                {
                    return E_INVALIDARG;
                }

                uint64_t position = 0;
                HRESULT hr        = reader.Seek(0, FILE_BEGIN, &position);
                if (FAILED(hr))
                {
                    return hr;
                }

                uint64_t hash = HashOffsetBasis();
                uint64_t readTotal = 0;
                while (readTotal < sizeBytes)
                {
                    if (CancelRequested())
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }

                    const uint64_t remaining = sizeBytes - readTotal;
                    const unsigned long toRead = remaining > hashBufferBytes ? hashBufferBytes : static_cast<unsigned long>(remaining);
                    unsigned long bytesRead = 0;
                    hr = reader.Read(hashBuffer, toRead, &bytesRead);
                    if (FAILED(hr))
                    {
                        return hr;
                    }
                    if (bytesRead > toRead)
                    {
                        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    }
                    if (bytesRead == 0u)
                    {
                        return S_OK;
                    }
                    HashBytes(hash, hashBuffer, bytesRead);
                    readTotal += bytesRead;
                }

                matches = hash == expectedHash;
                return S_OK;
            }

            [[nodiscard]] HRESULT CopiedFileStillMatchesDestination(const CopiedEntry& entry, bool& matches) noexcept
            {
                matches = false;
                if (! BasicFileInfoStillMatchesSource(entry))
                {
                    return S_OK;
                }

                wil::com_ptr<IFileReader> destinationReader;
                HRESULT hr = RetryTransientCleanupProbe([&]() noexcept
                {
                    destinationReader.reset();
                    return destinationIo.CreateFileReader(entry.destinationPath.c_str(), destinationReader.addressof());
                });
                if (FAILED(hr))
                {
                    return IsMissingPathHr(hr) ? S_OK : hr;
                }
                if (! destinationReader)
                {
                    return E_POINTER;
                }

                uint64_t destinationSize = 0;
                hr                       = RetryTransientCleanupProbe([&]() noexcept { return destinationReader->GetSize(&destinationSize); });
                if (FAILED(hr))
                {
                    return hr;
                }

                if (! entry.hasKnownSize || ! entry.hasContentHash || destinationSize != entry.sizeBytes)
                {
                    return S_OK;
                }

                return ReaderMatchesHash(*destinationReader, entry.sizeBytes, entry.contentHash, buffer.get(), bufferBytes, matches);
            }

            [[nodiscard]] HRESULT DeleteCopiedSourcePathNoRecursive(const CopiedEntry& entry) noexcept
            {
                const FileSystemFlags deleteFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
                BridgeCallback callback(task, &callbackMutex);
                const HRESULT hr = sourceFs.DeleteItem(entry.sourcePath.c_str(), deleteFlags, nullptr, &callback, cookie);
                return IsMissingPathHr(hr) ? S_OK : hr;
            }

            void NoteMoveSourceCleanupSkipped(const std::wstring& sourcePath, const std::wstring& destinationPath, std::wstring_view reason) noexcept
            {
                const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   partialHr,
                                   L"bridge.move.cleanup.skip",
                                   std::wstring(reason.data(), reason.size()),
                                   sourcePath,
                                   destinationPath);
            }

            [[nodiscard]] HRESULT DeleteCopiedSourceEntryForMove(const std::wstring& sourcePath,
                                                                 const std::wstring& destinationPath,
                                                                 MoveSourceCleanupStats& stats) noexcept
            {
                if (CancelRequested())
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                unsigned long sourceAttributes = 0;
                HRESULT hr = RetryTransientCleanupProbe([&]() noexcept { return sourceIo.GetAttributes(sourcePath.c_str(), &sourceAttributes); });
                if (IsMissingPathHr(hr))
                {
                    return S_OK;
                }
                if (FAILED(hr))
                {
                    if (IsTransientCleanupHr(hr))
                    {
                        return hr;
                    }
                    stats.anySkipped = true;
                    NoteMoveSourceCleanupSkipped(sourcePath, destinationPath, L"Source cleanup skipped because the source entry could not be re-probed.");
                    return S_OK;
                }

                CopiedEntry entry{};
                if (! TakeCopiedEntry(sourcePath, entry))
                {
                    stats.anySkipped = true;
                    NoteMoveSourceCleanupSkipped(
                        sourcePath, destinationPath, L"Source cleanup skipped because the live source entry was not in the copied manifest.");
                    return S_OK;
                }

                const bool sourceIsDirectory = (sourceAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                const bool sourceIsReparse   = (sourceAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                if (sourceIsDirectory)
                {
                    if (entry.kind != CopiedEntryKind::Directory)
                    {
                        stats.anySkipped = true;
                        NoteMoveSourceCleanupSkipped(sourcePath, destinationPath, L"Source cleanup skipped because the source entry kind changed after copy.");
                        return S_OK;
                    }
                    const bool sourceDirectoryStillMatches = BasicFileInfoStillMatchesSource(entry);
                    if (! sourceDirectoryStillMatches && sourceIsReparse)
                    {
                        stats.anySkipped = true;
                        NoteMoveSourceCleanupSkipped(sourcePath, destinationPath, L"Source cleanup skipped because the source directory changed after copy.");
                        return S_OK;
                    }

                    unsigned long destinationAttributes = 0;
                    hr = RetryTransientCleanupProbe([&]() noexcept { return destinationIo.GetAttributes(destinationPath.c_str(), &destinationAttributes); });
                    if (FAILED(hr) || (destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
                    {
                        if (FAILED(hr) && IsTransientCleanupHr(hr))
                        {
                            return hr;
                        }
                        stats.anySkipped = true;
                        NoteMoveSourceCleanupSkipped(
                            sourcePath, destinationPath, L"Source cleanup skipped because the destination directory could not be verified.");
                        return S_OK;
                    }

                    if (sourceIsReparse)
                    {
                        hr = DeleteCopiedSourcePathNoRecursive(entry);
                        if (FAILED(hr))
                        {
                            if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT || IsTransientCleanupHr(hr))
                            {
                                return hr == E_ABORT ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : hr;
                            }
                            stats.anySkipped = true;
                            NoteMoveSourceCleanupSkipped(
                                sourcePath, destinationPath, L"Source directory reparse cleanup skipped because the link itself could not be removed.");
                            return S_OK;
                        }

                        stats.anyDeleted = true;
                        return S_OK;
                    }

                    if (! sourceDirectoryStillMatches)
                    {
                        stats.anySkipped = true;
                        NoteMoveSourceCleanupSkipped(
                            sourcePath,
                            destinationPath,
                            L"Source directory changed after copy; cleanup will preserve the directory after removing copied children that still match.");
                    }

                    wil::com_ptr<IFilesInformation> info;
                    task._bridgeSourceDirectoryEnumerationCount.fetch_add(1u, std::memory_order_relaxed);
                    hr = RetryTransientCleanupProbe([&]() noexcept
                    {
                        info.reset();
                        return sourceFs.ReadDirectoryInfo(sourcePath.c_str(), info.addressof());
                    });
                    if (IsMissingPathHr(hr))
                    {
                        return S_OK;
                    }
                    if (FAILED(hr) || ! info)
                    {
                        if (FAILED(hr) && IsTransientCleanupHr(hr))
                        {
                            return hr;
                        }
                        stats.anySkipped = true;
                        NoteMoveSourceCleanupSkipped(sourcePath, destinationPath, L"Source cleanup skipped because the source directory could not be listed.");
                        return S_OK;
                    }

                    FileInfo* head = nullptr;
                    hr             = info->GetBuffer(&head);
                    if (FAILED(hr))
                    {
                        stats.anySkipped = true;
                        NoteMoveSourceCleanupSkipped(
                            sourcePath, destinationPath, L"Source cleanup skipped because the source directory buffer was unavailable.");
                        return S_OK;
                    }

                    if (head != nullptr)
                    {
                        unsigned long bufferSize = 0;
                        hr                       = info->GetBufferSize(&bufferSize);
                        if (FAILED(hr) || bufferSize < sizeof(FileInfo))
                        {
                            stats.anySkipped = true;
                            NoteMoveSourceCleanupSkipped(
                                sourcePath, destinationPath, L"Source cleanup skipped because the source directory buffer was invalid.");
                            return S_OK;
                        }

                        std::byte* base   = reinterpret_cast<std::byte*>(head);
                        std::byte* end    = base + bufferSize;
                        FileInfo* current = head;
                        for (;;)
                        {
                            std::wstring_view name;
                            hr = TryGetValidatedFileInfoName(current, base, end, name);
                            if (FAILED(hr))
                            {
                                stats.anySkipped = true;
                                NoteMoveSourceCleanupSkipped(
                                    sourcePath, destinationPath, L"Source cleanup skipped because a source directory entry was invalid.");
                                return S_OK;
                            }

                            hr = ValidateBridgeStructuralChildName(name);
                            if (FAILED(hr))
                            {
                                stats.anySkipped = true;
                                NoteMoveSourceCleanupSkipped(
                                    sourcePath, destinationPath, L"Source cleanup skipped because a source directory entry name was unsafe.");
                                return S_OK;
                            }

                            hr = DeleteCopiedSourceEntryForMove(JoinFolderAndLeaf(sourcePath, name), JoinFolderAndLeaf(destinationPath, name), stats);
                            if (FAILED(hr))
                            {
                                return hr;
                            }

                            FileInfo* next = nullptr;
                            hr             = AdvanceValidatedFileInfoEntry(current, base, end, next);
                            if (hr == S_FALSE)
                            {
                                break;
                            }
                            if (FAILED(hr))
                            {
                                stats.anySkipped = true;
                                NoteMoveSourceCleanupSkipped(
                                    sourcePath, destinationPath, L"Source cleanup skipped because a source directory entry chain was invalid.");
                                return S_OK;
                            }

                            current = next;
                        }
                    }

                    if (! sourceDirectoryStillMatches)
                    {
                        return S_OK;
                    }

                    hr = DeleteCopiedSourcePathNoRecursive(entry);
                    if (FAILED(hr))
                    {
                        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT || IsTransientCleanupHr(hr))
                        {
                            return hr == E_ABORT ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : hr;
                        }
                        stats.anySkipped = true;
                        NoteMoveSourceCleanupSkipped(
                            sourcePath, destinationPath, L"Source directory cleanup skipped because the verified directory could not be removed.");
                        return S_OK;
                    }

                    stats.anyDeleted = true;
                    return S_OK;
                }

                if (entry.kind != CopiedEntryKind::File)
                {
                    stats.anySkipped = true;
                    NoteMoveSourceCleanupSkipped(sourcePath, destinationPath, L"Source cleanup skipped because the source entry kind changed after copy.");
                    return S_OK;
                }

                bool matches = false;
                hr           = CopiedFileStillMatchesDestination(entry, matches);
                if (FAILED(hr))
                {
                    if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }
                    if (IsTransientCleanupHr(hr))
                    {
                        return hr;
                    }
                    stats.anySkipped = true;
                    NoteMoveSourceCleanupSkipped(sourcePath, destinationPath, L"Source cleanup skipped because copied file equality could not be verified.");
                    return S_OK;
                }
                if (! matches)
                {
                    stats.anySkipped = true;
                    NoteMoveSourceCleanupSkipped(
                        sourcePath, destinationPath, L"Source cleanup skipped because the source or destination file changed after copy.");
                    return S_OK;
                }

                hr = DeleteCopiedSourcePathNoRecursive(entry);
                if (FAILED(hr))
                {
                    if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT || IsTransientCleanupHr(hr))
                    {
                        return hr == E_ABORT ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : hr;
                    }
                    stats.anySkipped = true;
                    NoteMoveSourceCleanupSkipped(sourcePath, destinationPath, L"Source cleanup skipped because the verified source file could not be removed.");
                    return S_OK;
                }

                stats.anyDeleted = true;
                return S_OK;
            }

            [[nodiscard]] HRESULT DeleteCopiedSourceForMove(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                MoveSourceCleanupStats stats{};
                const HRESULT hr = DeleteCopiedSourceEntryForMove(sourcePath, destinationPath, stats);
                uint64_t remainingEntries = 0;
                {
                    std::scoped_lock lock(copiedEntriesMutex);
                    remainingEntries = static_cast<uint64_t>(copiedEntries.size());
                }
                const HRESULT result = FAILED(hr) ? hr : (stats.anySkipped ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK);
                Debug::Perf::Emit(L"FileOps.Bridge.MoveManifestPeakEntries", L"", 0u, copiedEntriesPeakCount, remainingEntries, result);
                Debug::Perf::Emit(L"FileOps.Bridge.MoveManifestRemainingEntries", L"", 0u, remainingEntries, copiedEntriesPeakCount, result);
                if (FAILED(hr))
                {
                    if ((hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT) && stats.anyDeleted)
                    {
                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                           HRESULT_FROM_WIN32(ERROR_CANCELLED),
                                           L"bridge.move.cleanup.cancelled",
                                           L"Cross-filesystem MOVE cleanup was cancelled after removing part of the verified source tree.",
                                           sourcePath,
                                           destinationPath);
                    }
                    return hr;
                }
                return result;
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

            [[nodiscard]] HRESULT AcquireCopyMovePermits(ConnectionConcurrencyLimiter::Permit& outFirst,
                                                         ConnectionConcurrencyLimiter::Permit& outSecond) noexcept
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
                                                                     const std::wstring& destinationPath,
                                                                     bool& overwriteGranted,
                                                                     bool& replaceReadOnlyGranted) noexcept
            {
                overwriteGranted       = (static_cast<uint32_t>(flags) & static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE)) != 0u;
                replaceReadOnlyGranted = (static_cast<uint32_t>(flags) & static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY)) != 0u;

                unsigned long destinationAttributes = 0;
                const HRESULT hrDestAttr            = destinationIo.GetAttributes(destinationPath.c_str(), &destinationAttributes);
                if (FAILED(hrDestAttr))
                {
                    return S_OK; // destination free
                }

                if ((destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                {
                    bool ignoredOverwriteGrant    = false;
                    bool ignoredReadOnlyGrant     = false;
                    const HRESULT directoryPrompt = PromptDestinationCollision(
                        sourcePath, destinationPath, HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), ignoredOverwriteGrant, ignoredReadOnlyGrant);
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
                const bool alreadyAuthorized = overwriteGranted && (! readonlyCollision || replaceReadOnlyGranted);
                if (alreadyAuthorized)
                {
                    return S_OK;
                }

                // Raise a PER-FILE conflict instead of failing the whole bridge transfer. The
                // host serializes prompts and caches apply-to-all answers (Fairstream 1C/1D).
                return PromptDestinationCollision(sourcePath,
                                                  destinationPath,
                                                  HRESULT_FROM_WIN32(readonlyCollision ? ERROR_ACCESS_DENIED : ERROR_ALREADY_EXISTS),
                                                  overwriteGranted,
                                                  replaceReadOnlyGranted);
            }

            HRESULT CopyFileWithBuffer(const std::wstring& sourcePath,
                                       const std::wstring& destinationPath,
                                       std::byte* bufferIn,
                                       unsigned long bufferBytesIn,
                                       uint64_t progressStreamId,
                                       std::atomic<uint64_t>& overallCompletedBytes,
                                       bool adoptFileSizeAsTotalWhenUnknown = false) noexcept
            {
                if (! bufferIn || bufferBytesIn == 0)
                {
                    return E_INVALIDARG;
                }

#ifdef ENABLE_TESTS
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

                bool overwriteGranted       = false;
                bool replaceReadOnlyGranted = false;
                const HRESULT hrDestPolicy  = ValidateDestinationOverwritePolicy(sourcePath, destinationPath, overwriteGranted, replaceReadOnlyGranted);
                if (FAILED(hrDestPolicy))
                {
                    return hrDestPolicy;
                }
                if (hrDestPolicy == S_FALSE)
                {
                    // User chose Skip for this file: leave the destination untouched, count it so
                    // the move keeps the source and the task ends PARTIAL.
                    RecordSkippedDestinationCollision(sourcePath,
                                                      destinationPath,
                                                      HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                                                      L"bridge.conflict.skip",
                                                      L"Destination already exists; skipped on user request.");
                    return S_FALSE;
                }

                wil::com_ptr<IFileReader> reader;
                HRESULT hr = sourceIo.CreateFileReader(sourcePath.c_str(), reader.addressof());
                if (FAILED(hr))
                {
                    return hr;
                }

                FileSystemBasicInformation sourceBasicInfo{};
                sourceBasicInfo.sizeBytes = sizeof(FileSystemBasicInformation);
                bool hasSourceBasicInfo   = false;
                const HRESULT hrGetBasic  = sourceIo.GetFileBasicInformation(sourcePath.c_str(), &sourceBasicInfo);
                if (SUCCEEDED(hrGetBasic))
                {
                    hasSourceBasicInfo = true;
                }
                else if (hrGetBasic != E_NOTIMPL && hrGetBasic != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
                {
                    Debug::Warning(
                        L"CrossFileSystemBridge: GetFileBasicInformation failed for '{}' (hr={:#x})", sourcePath, static_cast<unsigned long>(hrGetBasic));
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                       hrGetBasic,
                                       L"bridge.metadata.read",
                                       L"GetFileBasicInformation failed for source file.",
                                       sourcePath,
                                       destinationPath);
                }

                uint64_t fileTotalBytes     = 0;
                bool hasKnownFileTotalBytes = false;
                const HRESULT hrReaderSize  = reader->GetSize(&fileTotalBytes);
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

                if (adoptFileSizeAsTotalWhenUnknown && totalBytes == 0 && hasKnownFileTotalBytes)
                {
                    totalBytes = fileTotalBytes;
                }

                bool useAtomicFinalWriter = false;
                wil::com_ptr<IFileSystemAtomicWriter> atomicWriterCapability;
                if (SUCCEEDED(destinationFs.QueryInterface(IID_PPV_ARGS(atomicWriterCapability.addressof()))) && atomicWriterCapability)
                {
                    BOOL supported = FALSE;
                    const HRESULT capabilityHr = atomicWriterCapability->SupportsAtomicWriterCommit(destinationPath.c_str(), flags, &supported);
                    useAtomicFinalWriter = SUCCEEDED(capabilityHr) && supported == TRUE;
                }

                const std::wstring tempPath = MakeTempDestinationPath(destinationPath, progressStreamId);
                const std::wstring& writerPath = useAtomicFinalWriter ? destinationPath : tempPath;
                bool tempStaged             = false;
                bool promoted               = false;
                const auto cleanupTemp      = wil::scope_exit([&] noexcept
                {
                    if (! promoted && tempStaged)
                    {
                        BestEffortDeleteTempFile(tempPath, sourcePath, destinationPath);
                    }
                });

                const FileSystemFlags tempFlags =
                    static_cast<FileSystemFlags>(static_cast<uint32_t>(flags) & ~(static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE) |
                                                                                  static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY)));
                FileSystemFlags writerFlags = tempFlags;
                if (useAtomicFinalWriter && overwriteGranted)
                {
                    writerFlags = static_cast<FileSystemFlags>(static_cast<uint32_t>(writerFlags) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE));
                }
                if (useAtomicFinalWriter && replaceReadOnlyGranted)
                {
                    writerFlags =
                        static_cast<FileSystemFlags>(static_cast<uint32_t>(writerFlags) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));
                }

                wil::com_ptr<IFileWriter> writer;
                hr = destinationIo.CreateFileWriter(writerPath.c_str(), writerFlags, writer.addressof());
                if (FAILED(hr))
                {
                    return hr;
                }
                tempStaged = ! useAtomicFinalWriter;

                wil::com_ptr<IFileWriterExpectedSize> expectedSizeWriter;
                if (SUCCEEDED(writer->QueryInterface(IID_PPV_ARGS(expectedSizeWriter.addressof()))) && expectedSizeWriter)
                {
                    hr = expectedSizeWriter->SetExpectedSize(fileTotalBytes);
                    if (FAILED(hr))
                    {
                        return hr;
                    }
                }

                uint64_t fileCompletedBytes = 0;
                uint64_t fileContentHash    = HashOffsetBasis();
                hr                          = ReportProgress(
                    sourcePath, destinationPath, fileTotalBytes, fileCompletedBytes, overallCompletedBytes.load(std::memory_order_acquire), progressStreamId);
                if (FAILED(hr))
                {
                    return hr;
                }
                BridgeCopyPerf copyPerf{};
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

                        size_t offset = 0;
                        while (offset < bytesRead)
                        {
                            if (CancelRequested())
                            {
                                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                            }

                            unsigned long bytesWritten  = 0;
                            const unsigned long toWrite = static_cast<unsigned long>(
                                std::min(static_cast<size_t>(bytesRead - offset), static_cast<size_t>(std::numeric_limits<unsigned long>::max())));
                            const uint64_t writeStartUs = PerfNowUs();
                            const HRESULT hrWrite       = writer->Write(bufferIn + offset, toWrite, &bytesWritten);
                            copyPerf.writeUs += PerfElapsedUs(writeStartUs);
                            if (FAILED(hrWrite))
                            {
                                return hrWrite;
                            }
                            if (bytesWritten == 0)
                            {
                                return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
                            }
                            if (bytesWritten > toWrite)
                            {
                                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                            }

                            HashBytes(fileContentHash, bufferIn + offset, bytesWritten);
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

                            const HRESULT hrProgress = maybeReportProgress(overallAfter, forceProgress);
                            if (FAILED(hrProgress))
                            {
                                return hrProgress;
                            }

                            ThrottleThreadSafe(overallAfter);
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
                                    return pipelineStop.load(std::memory_order_acquire) || readerFinished.load(std::memory_order_acquire) ||
                                           slots[writeIndex].ready;
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
                                hr = (stopped || CancelRequested()) ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                    : (finished ? S_OK : HRESULT_FROM_WIN32(ERROR_CANCELLED));
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

                            size_t offset = 0;
                            while (offset < slotBytesRead)
                            {
                                if (CancelRequested())
                                {
                                    hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                                    break;
                                }

                                unsigned long bytesWritten  = 0;
                                const unsigned long toWrite = static_cast<unsigned long>(
                                    std::min(static_cast<size_t>(slotBytesRead - offset), static_cast<size_t>(std::numeric_limits<unsigned long>::max())));
                                const uint64_t writeStartUs = PerfNowUs();
                                const HRESULT hrWrite       = writer->Write(slots[writeIndex].buffer + offset, toWrite, &bytesWritten);
                                writerWriteUs += PerfElapsedUs(writeStartUs);
                                if (FAILED(hrWrite))
                                {
                                    hr = hrWrite;
                                    break;
                                }
                                if (bytesWritten == 0)
                                {
                                    hr = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
                                    break;
                                }
                                if (bytesWritten > toWrite)
                                {
                                    hr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                                    break;
                                }

                                HashBytes(fileContentHash, slots[writeIndex].buffer + offset, bytesWritten);
                                offset += bytesWritten;

                                if (fileCompletedBytes > std::numeric_limits<uint64_t>::max() - static_cast<uint64_t>(bytesWritten))
                                {
                                    hr = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                                    break;
                                }
                                fileCompletedBytes += bytesWritten;

                                const uint64_t previousOverall = overallCompletedBytes.fetch_add(bytesWritten, std::memory_order_acq_rel);
                                if (previousOverall > (std::numeric_limits<uint64_t>::max)() - static_cast<uint64_t>(bytesWritten))
                                {
                                    hr = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                                    break;
                                }
                                const uint64_t overallAfter = previousOverall + static_cast<uint64_t>(bytesWritten);
                                const bool forceProgress    = fileTotalBytes > 0 && fileCompletedBytes >= fileTotalBytes;

                                const HRESULT hrProgress = maybeReportProgress(overallAfter, forceProgress);
                                if (FAILED(hrProgress))
                                {
                                    hr = hrProgress;
                                    break;
                                }

                                ThrottleThreadSafe(overallAfter);
                            }

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
                    const HRESULT hrMismatch = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    const std::wstring message =
                        std::format(L"File copy size mismatch: expected {:L} bytes but wrote {:L} bytes.", fileTotalBytes, fileCompletedBytes);
                    task.LogDiagnostic(
                        FileOperationState::DiagnosticSeverity::Error, hrMismatch, L"bridge.integrity.sizeMismatch", message, sourcePath, destinationPath);
                    return hrMismatch;
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

                hr = writer->Commit();
                if (FAILED(hr))
                {
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
                writer.reset();

                if (! useAtomicFinalWriter)
                {
                    hr = PromoteTempToFinalPath(tempPath, destinationPath, overwriteGranted, replaceReadOnlyGranted);
                    if (FAILED(hr))
                    {
                        Debug::Warning(L"CrossFileSystemBridge: failed to promote temp '{}' to '{}' (hr={:#x})",
                                       tempPath,
                                       destinationPath,
                                       static_cast<unsigned long>(hr));
                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                           hr,
                                           L"bridge.promote",
                                           L"Failed to promote the staged bridge destination into the final path.",
                                           sourcePath,
                                           destinationPath);
                        return hr;
                    }
                }
                promoted = true;

                if (hasKnownFileTotalBytes)
                {
                    wil::com_ptr<IFileReader> destinationReader;
                    HRESULT hrDestinationSize     = destinationIo.CreateFileReader(destinationPath.c_str(), destinationReader.addressof());
                    uint64_t destinationSizeBytes = 0;
                    if (SUCCEEDED(hrDestinationSize) && destinationReader)
                    {
                        hrDestinationSize = destinationReader->GetSize(&destinationSizeBytes);
                    }
                    else if (SUCCEEDED(hrDestinationSize))
                    {
                        hrDestinationSize = E_POINTER;
                    }
                    if (FAILED(hrDestinationSize) || destinationSizeBytes != fileTotalBytes)
                    {
                        const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                        const bool isMove = task._operation == FILESYSTEM_MOVE;
                        const std::wstring message =
                            FAILED(hrDestinationSize)
                                ? std::format(L"Cross-filesystem {} could not re-stat destination after promote (hr=0x{:08X}).",
                                              isMove ? L"MOVE" : L"COPY",
                                              static_cast<unsigned long>(hrDestinationSize))
                                : std::format(
                                      L"Cross-filesystem {} destination size mismatch after promote: expected {:L} bytes but destination has {:L} bytes.",
                                      isMove ? L"MOVE" : L"COPY",
                                      fileTotalBytes,
                                      destinationSizeBytes);
                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                           partialHr,
                                           L"bridge.integrity.destinationSizeMismatch",
                                           message,
                                           sourcePath,
                                           destinationPath);
                        if (! isMove)
                        {
                            BestEffortDeleteTempFile(destinationPath, sourcePath, destinationPath);
                        }
                        return partialHr;
                    }

                    if (task._operation == FILESYSTEM_COPY)
                    {
                        bool hashMatches = false;
                        const HRESULT hashHr = ReaderMatchesHash(*destinationReader, fileTotalBytes, fileContentHash, bufferIn, bufferBytesIn, hashMatches);
                        if (FAILED(hashHr) || ! hashMatches)
                        {
                            const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                            task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                               partialHr,
                                               L"bridge.integrity.destinationHashMismatch",
                                               FAILED(hashHr) ? L"Cross-filesystem COPY could not hash the promoted destination."
                                                              : L"Cross-filesystem COPY destination hash did not match the copied source bytes.",
                                               sourcePath,
                                               destinationPath);
                            BestEffortDeleteTempFile(destinationPath, sourcePath, destinationPath);
                            return partialHr;
                        }
                    }
                }

                if (hasSourceBasicInfo)
                {
                    const HRESULT hrSetBasic = destinationIo.SetFileBasicInformation(destinationPath.c_str(), &sourceBasicInfo);
                    if (FAILED(hrSetBasic) && hrSetBasic != E_NOTIMPL && hrSetBasic != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
                    {
                        Debug::Warning(L"CrossFileSystemBridge: SetFileBasicInformation failed for '{}' (hr={:#x})",
                                       destinationPath,
                                       static_cast<unsigned long>(hrSetBasic));
                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                           hrSetBasic,
                                           L"bridge.metadata.write",
                                           L"SetFileBasicInformation failed for destination file.",
                                           sourcePath,
                                           destinationPath);
                    }
                }

                const uint64_t overallFinal   = overallCompletedBytes.load(std::memory_order_acquire);
                const uint64_t finalTotal     = hasKnownFileTotalBytes ? fileTotalBytes : fileCompletedBytes;
                const uint64_t finalCompleted = fileCompletedBytes;

                hr = ReportProgress(sourcePath, destinationPath, finalTotal, finalCompleted, overallFinal, progressStreamId);
                if (FAILED(hr))
                {
                    return hr;
                }

                RecordCopiedFile(sourcePath, destinationPath, fileTotalBytes, fileContentHash, sourceBasicInfo, hasSourceBasicInfo);
#ifdef ENABLE_TESTS
                if (task._operation == FILESYSTEM_MOVE)
                {
                    MaybeMutateBridgeDestinationBeforeMoveCleanupForSelfTest(destinationIo, destinationPath);
                }
#endif
                return S_OK;
            }

            HRESULT CopyFile(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
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
                const HRESULT hr = CopyFileWithBuffer(sourcePath, destinationPath, buffer.get(), bufferBytes, 0, overallCompletedBytes, true);
                if (SUCCEEDED(hr))
                {
                    completedBytes = overallCompletedBytes.load(std::memory_order_acquire);
                }
                return hr;
            }

            HRESULT CopyDirectorySequential(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                if (CancelRequested())
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                const bool continueOnError = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
                bool hadChildFailure       = false;
                std::set<std::wstring> exactChildNames;
                std::set<std::wstring, BridgeOrdinalIgnoreCaseLess> foldedChildNames;

                HRESULT hr = EnsureDestinationDirectory(sourcePath, destinationPath);
                if (hr == S_FALSE)
                {
                    return S_FALSE;
                }
                if (FAILED(hr))
                {
                    return hr;
                }

                wil::com_ptr<IFilesInformation> info;
                task._bridgeSourceDirectoryEnumerationCount.fetch_add(1u, std::memory_order_relaxed);
                hr = sourceFs.ReadDirectoryInfo(sourcePath.c_str(), info.addressof());
                if (FAILED(hr))
                {
                    return hr;
                }

                FileInfo* entry = nullptr;
                hr              = info->GetBuffer(&entry);
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
                    return S_OK;
                }

                unsigned long bufferSize = 0;
                hr                       = info->GetBufferSize(&bufferSize);
                if (FAILED(hr) || bufferSize < sizeof(FileInfo))
                {
                    return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }

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
                    task.WaitWhilePaused();
                    if (CancelRequested())
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }

                    std::wstring_view name;
                    hr = TryGetValidatedFileInfoName(entry, base, end, name);
                    if (FAILED(hr))
                    {
                        if (! continueOnError)
                        {
                            return hr;
                        }

                        hadChildFailure = true;
                        break;
                    }

                    hr = ValidateAndRegisterChildName(name, exactChildNames, foldedChildNames);
                    if (FAILED(hr))
                    {
                        NoteInvalidEnumeratedChildName(sourcePath, destinationPath);
                        if (! continueOnError)
                        {
                            return hr;
                        }
                        hadChildFailure = true;
                    }
                    else
                    {
                        const bool isDirectory         = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                        const bool isReparse           = (entry->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                        const std::wstring childSource = JoinFolderAndLeaf(sourcePath, name);
                        const std::wstring childDest   = JoinFolderAndLeaf(destinationPath, name);

                        if (isDirectory)
                        {
                            if (isReparse && reparsePointPolicy != ReparsePointPolicy::FollowTargets)
                            {
                                if (reparsePointPolicy == ReparsePointPolicy::Skip)
                                {
                                    MarkReparseSkipped(childSource, childDest, true, false);
                                    hr = S_OK;
                                }
                                else
                                {
                                    // copyReparse requires preserving a link; bridge cannot preserve NTFS reparse payloads.
                                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                                       HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                                       L"bridge.reparse.unsupported",
                                                       L"Cross-filesystem bridge cannot preserve directory reparse payloads.",
                                                       childSource,
                                                       childDest);
                                    unsupportedDirectoryReparseEncountered = true;
                                    hr                                     = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                                }
                            }
                            else
                            {
                                hr = CopyDirectorySequential(childSource, childDest);
                            }
                        }
                        else
                        {
                            if (isReparse && reparsePointPolicy != ReparsePointPolicy::FollowTargets)
                            {
                                if (reparsePointPolicy == ReparsePointPolicy::Skip)
                                {
                                    MarkReparseSkipped(childSource, childDest, false, false);
                                    hr = S_OK;
                                }
                                else
                                {
                                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                                       HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                                       L"bridge.reparse.unsupported",
                                                       L"Cross-filesystem bridge cannot preserve file reparse payloads.",
                                                       childSource,
                                                       childDest);
                                    unsupportedDirectoryReparseEncountered = true;
                                    hr                                     = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                                }
                            }
                            else
                            {
                                hr = CopyFile(childSource, childDest);
                            }
                        }

                        if (hr == S_FALSE)
                        {
                            hr = S_OK;
                        }
                        else if (FAILED(hr))
                        {
                            if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT || ! continueOnError)
                            {
                                return hr;
                            }

                            hadChildFailure = true;
                            hr              = S_OK;
                        }
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

                        hadChildFailure = true;
                        break;
                    }

                    entry = nextEntry;
                }

                if (hadChildFailure)
                {
                    return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                }

                return S_OK;
            }

            HRESULT CopyDirectoryParallel(const std::wstring& sourcePath, const std::wstring& destinationPath, unsigned int withinFolderBudget) noexcept
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

                HRESULT hrRoot = EnsureDestinationDirectory(sourcePath, destinationPath);
                if (hrRoot == S_FALSE)
                {
                    return S_FALSE;
                }
                if (FAILED(hrRoot))
                {
                    return hrRoot;
                }

                struct WorkItem final
                {
                    std::wstring source;
                    std::wstring destination;
                };

                std::deque<WorkItem> workItems;
                std::mutex workMutex;
                std::condition_variable workCv;
                const size_t maxQueuedWorkItems = GetBridgeAdmissionQueueLimit();
                std::atomic<uint64_t> overallCompletedBytes(completedBytes);
                std::atomic<bool> producerDone{false};
                std::atomic<uint64_t> fileStartedBeforeProducerDone{0};
                std::atomic<bool> stopRequested{false};
                std::atomic<bool> hadWorkerFailure{false};
                std::atomic<HRESULT> firstFailure{S_OK};
                uint64_t directoryEnsureCount   = 1;
                uint64_t fileAdmissionCount     = 0;
                uint64_t maxAdmissionQueueDepth = 0;

                const auto workerProc = [&](size_t workerIndex) noexcept -> HRESULT
                {
                    CrossFsBridgeBufferLease localBufferBudgetLease;
                    const uint64_t reservationBytes = static_cast<uint64_t>(bufferBytes) * 2ull;
                    const bool acquiredBudget = localBufferBudgetLease.Acquire(reservationBytes, task._cancelled, task._stopToken);
                    const HRESULT allocationHr = acquiredBudget
                                                     ? S_OK
                                                     : ((task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested())
                                                            ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
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
                        if (! continueOnError)
                        {
                            HRESULT expected = S_OK;
                            static_cast<void>(firstFailure.compare_exchange_strong(expected, failure));
                            stopRequested.store(true, std::memory_order_release);
                            workCv.notify_all();
                        }
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
                            while (workItems.empty() && ! producerDone.load(std::memory_order_acquire) && ! stopRequested.load(std::memory_order_acquire) &&
                                   ! CancelRequested())
                            {
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
                            workCv.notify_all();
                        }

                        if (! producerDone.load(std::memory_order_acquire))
                        {
                            fileStartedBeforeProducerDone.fetch_add(1u, std::memory_order_acq_rel);
                        }
                        const HRESULT hrItem =
                            CopyFileWithBuffer(item.source, item.destination, localBuffer.get(), localBufferBytes, progressStreamId, overallCompletedBytes);
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

                auto job = scheduler.StartJob(&task, withinFolderBudget, withinFolderBudget, workerProc);

                const auto recordProducerFailure = [&](HRESULT failure) noexcept
                {
                    hadWorkerFailure.store(true, std::memory_order_release);
                    const bool cancellation = failure == HRESULT_FROM_WIN32(ERROR_CANCELLED) || failure == E_ABORT;
                    if (cancellation || ! continueOnError)
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

                    std::unique_lock lock(workMutex);
                    while (workItems.size() >= maxQueuedWorkItems && ! producerDone.load(std::memory_order_acquire) &&
                           ! stopRequested.load(std::memory_order_acquire) && ! CancelRequested())
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
                    ++fileAdmissionCount;
                    maxAdmissionQueueDepth = (std::max)(maxAdmissionQueueDepth, static_cast<uint64_t>(workItems.size()));
                    lock.unlock();
                    workCv.notify_one();
                    return S_OK;
                };

                std::vector<std::pair<std::wstring, std::wstring>> stack;
                stack.emplace_back(sourcePath, destinationPath);

                while (! stack.empty() && ! stopRequested.load(std::memory_order_acquire))
                {
                    task.WaitWhilePaused();
                    if (CancelRequested())
                    {
                        recordProducerFailure(HRESULT_FROM_WIN32(ERROR_CANCELLED));
                        break;
                    }

                    auto [currentSource, currentDest] = std::move(stack.back());
                    stack.pop_back();

                    HRESULT hr = EnsureDestinationDirectory(currentSource, currentDest);
                    if (hr == S_FALSE)
                    {
                        continue;
                    }
                    if (FAILED(hr))
                    {
                        recordProducerFailure(hr);
                        if (! continueOnError)
                        {
                            break;
                        }
                        continue;
                    }
                    ++directoryEnsureCount;

                    wil::com_ptr<IFilesInformation> info;
                    task._bridgeSourceDirectoryEnumerationCount.fetch_add(1u, std::memory_order_relaxed);
                    hr = sourceFs.ReadDirectoryInfo(currentSource.c_str(), info.addressof());
                    if (FAILED(hr))
                    {
                        recordProducerFailure(hr);
                        if (! continueOnError)
                        {
                            break;
                        }
                        continue;
                    }

                    FileInfo* entry = nullptr;
                    hr              = info->GetBuffer(&entry);
                    if (FAILED(hr))
                    {
                        recordProducerFailure(hr);
                        if (! continueOnError)
                        {
                            break;
                        }
                        continue;
                    }
                    if (entry == nullptr)
                    {
                        // Empty directory: (nullptr, S_OK) per the IFilesInformation contract. The
                        // destination was already created by EnsureDestinationDirectory above, so there is
                        // nothing to enqueue and this is success, not a producer failure.
                        continue;
                    }

                    unsigned long bufferSize = 0;
                    hr                       = info->GetBufferSize(&bufferSize);
                    if (FAILED(hr) || bufferSize < sizeof(FileInfo))
                    {
                        recordProducerFailure(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
                        if (! continueOnError)
                        {
                            break;
                        }
                        continue;
                    }

                    std::byte* base = reinterpret_cast<std::byte*>(entry);
                    std::byte* end  = base + bufferSize;
#ifdef ENABLE_TESTS
                    hr = MaybeInjectHostileBridgeChildNamesForSelfTest(entry, base, end);
                    if (FAILED(hr))
                    {
                        recordProducerFailure(hr);
                        if (! continueOnError)
                        {
                            break;
                        }
                        continue;
                    }
                    hr = MaybeInjectBridgeFileReparseForSelfTest(entry, base, end);
                    if (FAILED(hr))
                    {
                        recordProducerFailure(hr);
                        if (! continueOnError)
                        {
                            break;
                        }
                        continue;
                    }
#endif
                    std::set<std::wstring> exactChildNames;
                    std::set<std::wstring, BridgeOrdinalIgnoreCaseLess> foldedChildNames;

                    for (;;)
                    {
                        task.WaitWhilePaused();
                        if (CancelRequested())
                        {
                            recordProducerFailure(HRESULT_FROM_WIN32(ERROR_CANCELLED));
                            break;
                        }

                        std::wstring_view name;
                        hr = TryGetValidatedFileInfoName(entry, base, end, name);
                        if (FAILED(hr))
                        {
                            recordProducerFailure(hr);
                            break;
                        }

                        hr = ValidateAndRegisterChildName(name, exactChildNames, foldedChildNames);
                        if (FAILED(hr))
                        {
                            NoteInvalidEnumeratedChildName(currentSource, currentDest);
                            recordProducerFailure(hr);
                            if (! continueOnError)
                            {
                                break;
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
                                if (isReparse && reparsePointPolicy != ReparsePointPolicy::FollowTargets)
                                {
                                    if (reparsePointPolicy == ReparsePointPolicy::Skip)
                                    {
                                        MarkReparseSkipped(childSource, childDest, true, false);
                                    }
                                    else
                                    {
                                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                                           HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                                           L"bridge.reparse.unsupported",
                                                           L"Cross-filesystem bridge cannot preserve directory reparse payloads.",
                                                           childSource,
                                                           childDest);
                                        unsupportedDirectoryReparseEncountered = true;
                                        recordProducerFailure(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
                                        if (! continueOnError)
                                        {
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    stack.emplace_back(std::move(childSource), std::move(childDest));
                                }
                            }
                            else
                            {
                                if (isReparse && reparsePointPolicy != ReparsePointPolicy::FollowTargets)
                                {
                                    if (reparsePointPolicy == ReparsePointPolicy::Skip)
                                    {
                                        MarkReparseSkipped(childSource, childDest, false, false);
                                    }
                                    else
                                    {
                                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                                           HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                                           L"bridge.reparse.unsupported",
                                                           L"Cross-filesystem bridge cannot preserve file reparse payloads.",
                                                           childSource,
                                                           childDest);
                                        unsupportedDirectoryReparseEncountered = true;
                                        recordProducerFailure(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
                                        if (! continueOnError)
                                        {
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    WorkItem item{};
                                    item.source      = std::move(childSource);
                                    item.destination = std::move(childDest);
                                    hr               = enqueueWork(std::move(item));
                                    if (FAILED(hr))
                                    {
                                        recordProducerFailure(hr);
                                        break;
                                    }
                                }
                            }
                        }

                        if (stopRequested.load(std::memory_order_acquire))
                        {
                            break;
                        }

                        FileInfo* nextEntry = nullptr;
                        hr                  = AdvanceValidatedFileInfoEntry(entry, base, end, nextEntry);
                        if (hr == S_FALSE)
                        {
                            break;
                        }
                        if (FAILED(hr))
                        {
                            recordProducerFailure(hr);
                            break;
                        }

                        entry = nextEntry;
                    }

#ifdef ENABLE_TESTS
                    const unsigned int producerDelayMs = GetBridgeProducerDelayMsForSelfTest();
                    if (producerDelayMs > 0)
                    {
                        SleepResponsive(producerDelayMs);
                    }
#endif
                }

                producerDone.store(true, std::memory_order_release);
                workCv.notify_all();
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

            HRESULT CopyDirectory(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                const unsigned int withinFolderBudget = ComputeWithinFolderBudget();
                if (withinFolderBudget <= 1u)
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

            HRESULT CopyPath(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                unsigned long attributes = sourceRootAttributesHint;
                const bool haveHint      = attributes != 0;

                if (! connectionLimitsInitialized)
                {
                    InitializeConnectionLimits(sourcePath, destinationPath);
                }

                // Hints can be stale, especially for recently-created junctions, and the reparse point policy relies on
                // accurate attributes. Prefer refreshing attributes when not following reparse targets; otherwise fall
                // back to the hint if refreshing fails.
                if (attributes == 0 || reparsePointPolicy != ReparsePointPolicy::FollowTargets)
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

                const bool isDirectory = (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                const bool isReparse   = (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                if (isDirectory)
                {
                    if (isReparse && reparsePointPolicy != ReparsePointPolicy::FollowTargets)
                    {
                        if (reparsePointPolicy == ReparsePointPolicy::Skip)
                        {
                            MarkReparseSkipped(sourcePath, destinationPath, true, true);
                            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                        }
                        task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                           HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                           L"bridge.reparse.unsupported",
                                           L"Cross-filesystem bridge cannot preserve root directory reparse payloads.",
                                           sourcePath,
                                           destinationPath);
                        unsupportedDirectoryReparseEncountered = true;
                        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                    }
                    const HRESULT dirHr = CopyDirectory(sourcePath, destinationPath);
                    if (SUCCEEDED(dirHr) && (skippedDirectoryReparseCount > 0 || skippedFileReparseCount > 0 ||
                                             skippedFileConflictCount.load(std::memory_order_acquire) > 0))
                    {
                        // Some child files were skipped at a conflict prompt; the tree is not a
                        // full copy. Caller treats PARTIAL as "source preserved" for MOVE.
                        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    }
                    return dirHr;
                }

                if (isReparse && reparsePointPolicy != ReparsePointPolicy::FollowTargets)
                {
                    if (reparsePointPolicy == ReparsePointPolicy::Skip)
                    {
                        MarkReparseSkipped(sourcePath, destinationPath, false, true);
                        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    }
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                       HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                       L"bridge.reparse.unsupported",
                                       L"Cross-filesystem bridge cannot preserve root file reparse payloads.",
                                       sourcePath,
                                       destinationPath);
                    unsupportedDirectoryReparseEncountered = true;
                    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
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

        const auto ShouldDeleteMoveSourceAfterBridgeCopy =
            [](HRESULT copyHr, unsigned long skippedDirectoryReparseCount, bool rootDirectoryReparseSkipped, uint64_t skippedFileConflictCount) noexcept -> bool
        {
            if (SUCCEEDED(copyHr))
            {
                assert(skippedFileConflictCount == 0);
            }
            if (FAILED(copyHr))
            {
                return false;
            }

            return skippedDirectoryReparseCount == 0 && ! rootDirectoryReparseSkipped && skippedFileConflictCount == 0;
        };

        const auto ensureResolvedDirectoryShell = [&]([[maybe_unused]] const std::wstring& sourceText,
                                                      const std::wstring& destinationText,
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
            std::filesystem::path currentPath(destinationText);
            for (std::filesystem::path parent = currentPath.parent_path();
                 ! parent.empty() && parent != parent.root_path() && parent != currentPath && ! IsUncShareRootBoundary(parent.native(), destinationShareRoot);
                 currentPath = parent, parent = currentPath.parent_path())
            {
                ancestors.push_back(parent.native());
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

        const auto ensureResolvedDestinationParent =
            [&](const std::wstring& sourceText, const std::wstring& destinationText, bool isDirectoryShell, FileSystemFlags itemFlags) noexcept -> HRESULT
        {
            if (! useResolvedItems || isDirectoryShell || destinationText.empty())
            {
                return S_OK;
            }

            const std::filesystem::path destinationPath(destinationText);
            const std::filesystem::path parentPath                 = destinationPath.parent_path();
            const std::optional<std::wstring> destinationShareRoot = TryGetUncShareRootBoundary(destinationText);
            if (parentPath.empty() || parentPath == destinationPath || parentPath == parentPath.root_path() ||
                IsUncShareRootBoundary(parentPath.native(), destinationShareRoot))
            {
                return S_OK;
            }

            return ensureResolvedDirectoryShell(sourceText, parentPath.native(), itemFlags);
        };

        if (_perItemMaxConcurrency > 1u)
        {
            // Per-task multi-item concurrency: run multiple CopyItem/MoveItem/DeleteItem calls concurrently while keeping
            // conflict prompts serialized (one prompt per task at a time).
            std::atomic<bool> hadSkipped{false};
            std::atomic<HRESULT> firstFailure{S_OK};

            const auto processIndex = [&](size_t index) noexcept -> HRESULT
            {
                const std::wstring& sourceText = _sourcePaths[index].native();
                if (sourceText.empty())
                {
                    return E_INVALIDARG;
                }

                const uint64_t preCalcBytesForItem = (canUsePreCalcBytes && index < _preCalcSourceBytes.size()) ? _preCalcSourceBytes[index] : 0;

                std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> retryCounts{};
                std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> cachedModifierAttempts{};
                FileSystemFlags itemFlags = useResolvedItems ? _resolvedItems[index].flags : _flags;

                bool itemSucceeded     = false;
                bool itemSkipped       = false;
                bool moveCopyCompleted = false;
                std::unique_ptr<CrossFileSystemBridge> moveBridge;
                bool bridgeUnsupportedDirectoryReparse = false;
                bool failedDuringMoveDelete            = false;
                uint64_t callCompletedBytes            = 0;
                uint64_t callCompletedItems            = 0;
                uint64_t callTotalItems                = 0;

                for (;;)
                {
                    WaitWhilePaused();
                    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }

                    std::wstring destinationItemText;
                    if (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE)
                    {
                        if (useResolvedItems)
                        {
                            destinationItemText = _resolvedItems[index].destinationPath.native();
                        }
                        else
                        {
                            const std::wstring_view leaf = GetPathLeaf(sourceText);
                            if (leaf.empty())
                            {
                                return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                            }
                            destinationItemText = JoinFolderAndLeaf(destinationFolderText, leaf);
                        }
                    }
                    const bool isDirectoryShell = useResolvedItems && _resolvedItems[index].kind == FolderWindow::ResolvedFileOperationItemKind::DirectoryShell;
                    const HRESULT hrEnsureParent = ensureResolvedDestinationParent(sourceText, destinationItemText, isDirectoryShell, itemFlags);
                    if (FAILED(hrEnsureParent))
                    {
                        return hrEnsureParent;
                    }

                    PerItemCallbackCookie cookie{index};

                    const PerItemInFlightAggregate inFlightAggregate = BeginPerItemInFlightCall(*this, &cookie, GetTickCount64());

                    {
                        std::scoped_lock lock(_progressMutex);
                        _progressCompletedItems = (std::max)(_progressCompletedItems, _perItemCompletedItems);
                        const uint64_t mapped   = _perItemCompletedBytes + inFlightAggregate.completedBytes;
                        _progressCompletedBytes = (std::max)(_progressCompletedBytes, mapped);
                        PublishProgressCountersLocked(*this);
                    }

                    callCompletedBytes                = 0;
                    callCompletedItems                = 0;
                    callTotalItems                    = 0;
                    bridgeUnsupportedDirectoryReparse = false;
                    failedDuringMoveDelete            = false;

                    ConnectionCircuitBreaker& breaker = GetConnectionCircuitBreaker();
                    HRESULT itemHr                    = E_NOTIMPL;
                    if (isDirectoryShell)
                    {
                        itemHr = RunWithCircuitBreaker(
                            breaker, getSourceCircuitBreakerConnectionId(index), destinationCircuitBreakerConnectionId, [&]() noexcept -> HRESULT {
                            return ensureResolvedDirectoryShell(sourceText, destinationItemText, itemFlags);
                        });
                    }
                    else if (_operation == FILESYSTEM_MOVE && useCrossFileSystemBridge)
                    {
                        // Cross-filesystem move is bridge copy + source delete, mirroring the serial
                        // path. Handing the source plugin a destination path from another plugin's
                        // namespace is never valid.
                        if (! moveCopyCompleted)
                        {
                            unsigned long bridgeSkippedDirectoryReparseCount = 0;
                            bool bridgeRootDirectoryReparseSkipped           = false;
                            uint64_t bridgeSkippedFileConflictCount          = 0;
                            itemHr                                           = RunWithCircuitBreaker(breaker,
                                                                                                     getSourceCircuitBreakerConnectionId(index),
                                                                                                     destinationCircuitBreakerConnectionId,
                                                                                                     [&]() noexcept -> HRESULT
                            {
                                std::unique_ptr<CrossFileSystemBridge> bridge(
                                    new (std::nothrow) CrossFileSystemBridge(*this,
                                                                             *_fileSystem,
                                                                             *_destinationFileSystem,
                                                                             *fileSystemIo,
                                                                             *destinationFileSystemIo,
                                                                             destinationDirOps.get(),
                                                                             bridgeSourceMaxConcurrencyBudget,
                                                                             bridgeDestinationMaxConcurrencyBudget,
                                                                             itemFlags,
                                                                             static_cast<void*>(&cookie),
                                                                             preCalcBytesForItem,
                                                                             sourceText.c_str(),
                                                                             destinationItemText.c_str(),
                                                                             (index < _sourcePathAttributesHint.size()) ? _sourcePathAttributesHint[index] : 0,
                                                                             reparsePointPolicy));
                                if (! bridge)
                                {
                                    return E_OUTOFMEMORY;
                                }

                                const HRESULT bridgeHr             = bridge->CopyPath(sourceText, destinationItemText);
                                bridgeSkippedDirectoryReparseCount = bridge->skippedDirectoryReparseCount + bridge->skippedFileReparseCount;
                                bridgeRootDirectoryReparseSkipped  = bridge->rootDirectoryReparseSkipped;
                                bridgeUnsupportedDirectoryReparse  = bridge->unsupportedDirectoryReparseEncountered;
                                bridgeSkippedFileConflictCount     = bridge->skippedFileConflictCount.load(std::memory_order_acquire);
                                if (SUCCEEDED(bridgeHr) &&
                                    ShouldDeleteMoveSourceAfterBridgeCopy(
                                        bridgeHr, bridgeSkippedDirectoryReparseCount, bridgeRootDirectoryReparseSkipped, bridgeSkippedFileConflictCount))
                                {
                                    moveBridge = std::move(bridge);
                                }
                                return bridgeHr;
                            });

                            if (SUCCEEDED(itemHr))
                            {
                                if (ShouldDeleteMoveSourceAfterBridgeCopy(
                                        itemHr, bridgeSkippedDirectoryReparseCount, bridgeRootDirectoryReparseSkipped, bridgeSkippedFileConflictCount))
                                {
                                    moveCopyCompleted = true;
                                }
                                else
                                {
                                    // Skipped reparse content never reached the destination; the
                                    // source stays authoritative and must not be deleted.
                                    itemHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                                }
                            }
                        }
                        else
                        {
                            itemHr = S_OK;
                        }

                        if (SUCCEEDED(itemHr) && moveCopyCompleted)
                        {
                            itemHr = RunWithCircuitBreaker(breaker,
                                                           getSourceCircuitBreakerConnectionId(index),
                                                           destinationCircuitBreakerConnectionId,
                                                           [&]() noexcept -> HRESULT
                            {
#ifdef ENABLE_TESTS
                                MaybePauseBeforeBridgeMoveSourceCleanupForSelfTest();
#endif
                                if (moveBridge)
                                {
                                    moveBridge->cookie = static_cast<void*>(&cookie);
                                }
                                return moveBridge ? moveBridge->DeleteCopiedSourceForMove(sourceText, destinationItemText)
                                                  : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                            });
                            if (FAILED(itemHr))
                            {
                                failedDuringMoveDelete = true;
                            }
                        }
                    }
                    else
                    {
                        itemHr = RunWithCircuitBreaker(breaker,
                                                       getSourceCircuitBreakerConnectionId(index),
                                                       destinationCircuitBreakerConnectionId,
                                                       [&]() noexcept -> HRESULT
                        {
                            if (_operation == FILESYSTEM_COPY)
                            {
                                if (useCrossFileSystemBridge)
                                {
                                    CrossFileSystemBridge bridge(*this,
                                                                 *_fileSystem,
                                                                 *_destinationFileSystem,
                                                                 *fileSystemIo,
                                                                 *destinationFileSystemIo,
                                                                 destinationDirOps.get(),
                                                                 bridgeSourceMaxConcurrencyBudget,
                                                                 bridgeDestinationMaxConcurrencyBudget,
                                                                 itemFlags,
                                                                 static_cast<void*>(&cookie),
                                                                 preCalcBytesForItem,
                                                                 sourceText.c_str(),
                                                                 destinationItemText.c_str(),
                                                                 (index < _sourcePathAttributesHint.size()) ? _sourcePathAttributesHint[index] : 0,
                                                                 reparsePointPolicy);
                                    const HRESULT bridgeHr            = bridge.CopyPath(sourceText, destinationItemText);
                                    bridgeUnsupportedDirectoryReparse = bridge.unsupportedDirectoryReparseEncountered;
                                    return bridgeHr;
                                }

                                FileSystemOptions options{};
                                InitializeFileSystemOptions(options);
                                return _fileSystem->CopyItem(
                                    sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                            }

                            if (_operation == FILESYSTEM_MOVE)
                            {
                                FileSystemOptions options{};
                                InitializeFileSystemOptions(options);
                                return _fileSystem->MoveItem(
                                    sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                            }

                            if (_operation == FILESYSTEM_DELETE)
                            {
                                return _fileSystem->DeleteItem(sourceText.c_str(), itemFlags, nullptr, this, static_cast<void*>(&cookie));
                            }

                            return E_NOTIMPL;
                        });
                    }

                    const PerItemInFlightFinishResult finishedCall = FinishPerItemInFlightCall(*this, &cookie);
                    callCompletedItems                             = finishedCall.completedItems;
                    callCompletedBytes                             = finishedCall.completedBytes;
                    callTotalItems                                 = finishedCall.totalItems;

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

                            const bool precalcTotalAvailable = _preCalcCompleted.load(std::memory_order_acquire) && _progressTotalItems > 0;
                            if (! precalcTotalAvailable)
                            {
                                const uint64_t mappedTotalItems = _perItemTotalEntryCount + finishedCall.aggregate.totalItems;
                                if (mappedTotalItems > 0)
                                {
                                    const uint64_t clampedTotal =
                                        std::min<uint64_t>(mappedTotalItems, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()));
                                    _progressTotalItems = (std::max)(_progressTotalItems, static_cast<unsigned long>(clampedTotal));
                                }
                            }
                        }

                        const uint64_t mapped   = _perItemCompletedBytes + finishedCall.aggregate.completedBytes;
                        _progressCompletedBytes = (std::max)(_progressCompletedBytes, mapped);
                        PublishProgressCountersLocked(*this);
                    }

                    const bool cancelled = itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || itemHr == E_ABORT;
                    if (cancelled)
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }

                    if (itemHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
                    {
                        itemSucceeded = true;
                        hadSkipped.store(true, std::memory_order_release);
                        break;
                    }

                    if (SUCCEEDED(itemHr))
                    {
                        itemSucceeded = true;
                        break;
                    }

                    if (continueOnError)
                    {
                        auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, &cookie, sourceText, destinationItemText);
                        LogDiagnostic(DiagnosticSeverity::Warning,
                                      itemHr,
                                      L"item.continueOnError",
                                      L"Item failed and was skipped due continue-on-error.",
                                      diagnosticSource,
                                      diagnosticDestination);
                        itemSkipped = true;
                        hadSkipped.store(true, std::memory_order_release);
                        break;
                    }

                    const FileSystemOperation bucketOperation = failedDuringMoveDelete ? FILESYSTEM_DELETE : _operation;
                    const wil::com_ptr<IFileSystemIO>& bucketFileSystemIo =
                        failedDuringMoveDelete ? fileSystemIo : (useCrossFileSystemBridge ? destinationFileSystemIo : fileSystemIo);
                    const ConflictBucket bucket = ClassifyConflictBucket(
                        bucketOperation, itemFlags, bucketFileSystemIo, itemHr, sourceText, destinationItemText, bridgeUnsupportedDirectoryReparse);
                    if (bucket == ConflictBucket::RecycleBinFailed)
                    {
                        auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, &cookie, sourceText, destinationItemText);
                        LogDiagnostic(DiagnosticSeverity::Error,
                                      itemHr,
                                      L"delete.recycleBin.item",
                                      L"Recycle Bin delete failed for item.",
                                      diagnosticSource,
                                      diagnosticDestination);
                    }

                    const size_t bucketIndex = static_cast<size_t>(bucket);

                    std::optional<ConflictAction> cached   = LoadConflictDecisionFromCache(*this, bucket);
                    const bool ignoreCachedDecisionForItem = cached.has_value() && IsModifierConflictAction(cached.value()) &&
                                                             bucketIndex < cachedModifierAttempts.size() &&
                                                             cachedModifierAttempts[bucketIndex] >= kMaxCachedModifierAttemptsPerBucket;
                    if (ignoreCachedDecisionForItem)
                    {
                        // Only this item falls back to a fresh prompt; the cached apply-to-all
                        // decision stays valid for every other item in the task.
                        cached.reset();
                    }
                    ConflictAction action = cached.value_or(ConflictAction::None);

                    if (action == ConflictAction::None)
                    {
                        const bool canRetryBucket = IsRetryableConflictBucket(bucket);
                        const bool allowRetry     = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u;
                        const bool retryFailed    = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] != 0u;

                        const ConflictPromptBeginResult promptBegin = BeginConflictPrompt(
                            *this, &cookie, bucket, itemHr, sourceText, destinationItemText, allowRetry, retryFailed, ignoreCachedDecisionForItem);
                        action = promptBegin.action;
                        if (promptBegin.ownsPrompt)
                        {
                            const auto decision = WaitForConflictDecision(*this, &cookie, bucket);
                            action              = decision.first;
                        }
                    }

                    if (action == ConflictAction::Overwrite)
                    {
                        if (bucketIndex < cachedModifierAttempts.size())
                        {
                            ++cachedModifierAttempts[bucketIndex];
                        }
                        itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                        continue;
                    }

                    if (action == ConflictAction::ReplaceReadOnly)
                    {
                        if (bucketIndex < cachedModifierAttempts.size())
                        {
                            ++cachedModifierAttempts[bucketIndex];
                        }
                        itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
                        continue;
                    }

                    if (action == ConflictAction::PermanentDelete)
                    {
                        if (bucketIndex < cachedModifierAttempts.size())
                        {
                            ++cachedModifierAttempts[bucketIndex];
                        }
                        itemFlags = static_cast<FileSystemFlags>(itemFlags & ~FILESYSTEM_FLAG_USE_RECYCLE_BIN);
                        continue;
                    }

                    if (action == ConflictAction::Retry)
                    {
                        if (bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u)
                        {
                            retryCounts[bucketIndex] = 1u;
                            if (bucket == ConflictBucket::SharingViolation)
                            {
                                Sleep(750);
                            }
                            continue;
                        }
                        action = ConflictAction::Skip;
                    }

                    if (action == ConflictAction::Skip)
                    {
                        auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, &cookie, sourceText, destinationItemText);
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

                if (itemSkipped && preCalcBytesForItem > 0)
                {
                    std::scoped_lock lock(_progressMutex);
                    _progressTotalBytes = (_progressTotalBytes >= preCalcBytesForItem) ? (_progressTotalBytes - preCalcBytesForItem) : 0;
                    // If pre-calc bytes were counted into total, and the user later skips the item,
                    // ensure we don't end up reporting "completed > total" (progress > 100%).
                    _progressCompletedBytes = (std::min)(_progressCompletedBytes, _progressTotalBytes);
                    PublishProgressCountersLocked(*this);
                }

                uint64_t bytesForItem = 0;
                if (itemSucceeded)
                {
                    bytesForItem = (preCalcBytesForItem > 0) ? preCalcBytesForItem : callCompletedBytes;
                }

                const PerItemInFlightAggregate inFlightAggregate = getPerItemInFlightAggregate();
                StorePublishedTopLevelCompletionSnapshot(*this, MarkTopLevelItemCompleted(*this, index));

                {
                    std::scoped_lock lock(_progressMutex);
                    if (itemSucceeded)
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
                    PublishProgressCountersLocked(*this);
                }

                return S_OK;
            };

            auto& scheduler                     = GetPerItemTaskScheduler();
            const auto schedulerStart           = scheduler.CapturePerfSnapshot();
            const uint64_t schedulerWallStartUs = PerfNowUs();

            auto job = scheduler.StartJob(this,
                                          _perItemMaxConcurrency,
                                          _sourcePaths.size(),
                                          [&](size_t index) noexcept -> HRESULT
            {
                const HRESULT hrItem = processIndex(index);
                if (FAILED(hrItem))
                {
                    HRESULT expected = S_OK;
                    firstFailure.compare_exchange_strong(expected, hrItem, std::memory_order_acq_rel);
                    RequestCancel();
                }
                return hrItem;
            });

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

            ClearConflictPrompt(*this);

            const HRESULT hr = firstFailure.load(std::memory_order_acquire);
            if (FAILED(hr))
            {
                return hr;
            }

            if (hadSkipped.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }

            return S_OK;
        }

        for (size_t index = 0; index < _sourcePaths.size(); ++index)
        {
            const std::wstring& sourceText = _sourcePaths[index].native();
            if (sourceText.empty())
            {
                return E_INVALIDARG;
            }

            const uint64_t preCalcBytesForItem = (canUsePreCalcBytes && index < _preCalcSourceBytes.size()) ? _preCalcSourceBytes[index] : 0;

            std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> retryCounts{};
            std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> cachedModifierAttempts{};

            bool itemSucceeded        = false;
            bool itemSkipped          = false;
            bool itemPartiallySkipped = false;

            FileSystemFlags itemFlags = _flags;
            if (useResolvedItems)
            {
                itemFlags = _resolvedItems[index].flags;
            }
            uint64_t callCompletedBytes = 0;
            uint64_t callCompletedItems = 0;
            uint64_t callTotalItems     = 0;
            bool moveCopyCompleted      = false;
            uint64_t moveCopiedBytes    = 0;
            std::unique_ptr<CrossFileSystemBridge> moveBridge;

            for (;;)
            {
                WaitWhilePaused();
                if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                {
                    ClearConflictPrompt(*this);
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                callCompletedBytes = 0;
                callCompletedItems = 0;
                callTotalItems     = 0;

                std::wstring destinationItemText;
                if (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE)
                {
                    if (useResolvedItems)
                    {
                        destinationItemText = _resolvedItems[index].destinationPath.native();
                    }
                    else
                    {
                        const std::wstring_view leaf = GetPathLeaf(sourceText);
                        if (leaf.empty())
                        {
                            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                        }
                        destinationItemText = JoinFolderAndLeaf(destinationFolderText, leaf);
                    }
                }
                const bool isDirectoryShell  = useResolvedItems && _resolvedItems[index].kind == FolderWindow::ResolvedFileOperationItemKind::DirectoryShell;
                const HRESULT hrEnsureParent = ensureResolvedDestinationParent(sourceText, destinationItemText, isDirectoryShell, itemFlags);
                if (FAILED(hrEnsureParent))
                {
                    return hrEnsureParent;
                }

                PerItemCallbackCookie cookie{index};

                const PerItemInFlightAggregate inFlightAggregate = ResetPerItemInFlightCalls(*this, &cookie, GetTickCount64());

                {
                    std::scoped_lock lock(_progressMutex);
                    _perItemCompletedItems  = static_cast<unsigned long>(std::min<uint64_t>(static_cast<uint64_t>(index), static_cast<uint64_t>(ULONG_MAX)));
                    _progressCompletedItems = _perItemCompletedItems;
                    const uint64_t mapped   = _perItemCompletedBytes + inFlightAggregate.completedBytes;
                    _progressCompletedBytes = (std::max)(_progressCompletedBytes, mapped);
                    PublishProgressCountersLocked(*this);
                }

                ConnectionCircuitBreaker& breaker = GetConnectionCircuitBreaker();

                HRESULT itemHr                                   = E_NOTIMPL;
                bool failedDuringMoveDelete                      = false;
                unsigned long bridgeSkippedDirectoryReparseCount = 0;
                bool bridgeRootDirectoryReparseSkipped           = false;
                bool bridgeUnsupportedDirectoryReparse           = false;
                uint64_t bridgeSkippedFileConflictCount          = 0;

                if (isDirectoryShell)
                {
                    itemHr = RunWithCircuitBreaker(
                        breaker, getSourceCircuitBreakerConnectionId(index), destinationCircuitBreakerConnectionId, [&]() noexcept -> HRESULT {
                        return ensureResolvedDirectoryShell(sourceText, destinationItemText, itemFlags);
                    });
                }
                else if (_operation == FILESYSTEM_COPY)
                {
                    itemHr = RunWithCircuitBreaker(breaker,
                                                   getSourceCircuitBreakerConnectionId(index),
                                                   destinationCircuitBreakerConnectionId,
                                                   [&]() noexcept -> HRESULT
                    {
                        if (useCrossFileSystemBridge)
                        {
                            CrossFileSystemBridge bridge(*this,
                                                         *_fileSystem,
                                                         *_destinationFileSystem,
                                                         *fileSystemIo,
                                                         *destinationFileSystemIo,
                                                         destinationDirOps.get(),
                                                         bridgeSourceMaxConcurrencyBudget,
                                                         bridgeDestinationMaxConcurrencyBudget,
                                                         itemFlags,
                                                         static_cast<void*>(&cookie),
                                                         preCalcBytesForItem,
                                                         sourceText.c_str(),
                                                         destinationItemText.c_str(),
                                                         (index < _sourcePathAttributesHint.size()) ? _sourcePathAttributesHint[index] : 0,
                                                         reparsePointPolicy);
                            const HRESULT hr                   = bridge.CopyPath(sourceText, destinationItemText);
                            bridgeSkippedDirectoryReparseCount = bridge.skippedDirectoryReparseCount + bridge.skippedFileReparseCount;
                            bridgeRootDirectoryReparseSkipped  = bridge.rootDirectoryReparseSkipped;
                            bridgeUnsupportedDirectoryReparse  = bridge.unsupportedDirectoryReparseEncountered;
                            bridgeSkippedFileConflictCount     = bridge.skippedFileConflictCount.load(std::memory_order_acquire);
                            return hr;
                        }

                        FileSystemOptions options{};
                        InitializeFileSystemOptions(options);
                        return _fileSystem->CopyItem(sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                    });
                }
                else if (_operation == FILESYSTEM_MOVE)
                {
                    if (useCrossFileSystemBridge)
                    {
                        // For cross-filesystem move: copy + delete. If the copy already succeeded and we're retrying due
                        // to a delete failure, skip re-copying (avoid prompting for overwrite again).
                        if (! moveCopyCompleted)
                        {
                            itemHr = RunWithCircuitBreaker(breaker,
                                                           getSourceCircuitBreakerConnectionId(index),
                                                           destinationCircuitBreakerConnectionId,
                                                           [&]() noexcept -> HRESULT
                            {
                                std::unique_ptr<CrossFileSystemBridge> bridge(
                                    new (std::nothrow) CrossFileSystemBridge(*this,
                                                                             *_fileSystem,
                                                                             *_destinationFileSystem,
                                                                             *fileSystemIo,
                                                                             *destinationFileSystemIo,
                                                                             destinationDirOps.get(),
                                                                             bridgeSourceMaxConcurrencyBudget,
                                                                             bridgeDestinationMaxConcurrencyBudget,
                                                                             itemFlags,
                                                                             static_cast<void*>(&cookie),
                                                                             preCalcBytesForItem,
                                                                             sourceText.c_str(),
                                                                             destinationItemText.c_str(),
                                                                             (index < _sourcePathAttributesHint.size()) ? _sourcePathAttributesHint[index] : 0,
                                                                             reparsePointPolicy));
                                if (! bridge)
                                {
                                    return E_OUTOFMEMORY;
                                }

                                const HRESULT hr                   = bridge->CopyPath(sourceText, destinationItemText);
                                bridgeSkippedDirectoryReparseCount = bridge->skippedDirectoryReparseCount + bridge->skippedFileReparseCount;
                                bridgeRootDirectoryReparseSkipped  = bridge->rootDirectoryReparseSkipped;
                                bridgeUnsupportedDirectoryReparse  = bridge->unsupportedDirectoryReparseEncountered;
                                bridgeSkippedFileConflictCount     = bridge->skippedFileConflictCount.load(std::memory_order_acquire);
                                if (SUCCEEDED(hr))
                                {
                                    moveCopyCompleted = ShouldDeleteMoveSourceAfterBridgeCopy(
                                        hr, bridgeSkippedDirectoryReparseCount, bridgeRootDirectoryReparseSkipped, bridgeSkippedFileConflictCount);
                                    if (moveCopyCompleted)
                                    {
                                        moveCopiedBytes = bridge->completedBytes;
                                        moveBridge      = std::move(bridge);
                                    }
                                }
                                return hr;
                            });
                        }
                        else
                        {
                            itemHr = S_OK;
                        }

                        if (SUCCEEDED(itemHr) && moveCopyCompleted)
                        {
                            // Ensure the in-flight call has the best-known completed-bytes snapshot even when we're only deleting.
                            if (moveCopiedBytes > 0)
                            {
                                FileSystemOptions options{};
                                InitializeFileSystemOptions(options);
                                const HRESULT hrProgress = FileSystemProgress(_operation,
                                                                              1,
                                                                              0,
                                                                              preCalcBytesForItem,
                                                                              moveCopiedBytes,
                                                                              sourceText.c_str(),
                                                                              destinationItemText.c_str(),
                                                                              moveCopiedBytes,
                                                                              moveCopiedBytes,
                                                                              &options,
                                                                              0,
                                                                              static_cast<void*>(&cookie));
                                if (FAILED(hrProgress))
                                {
                                    itemHr = hrProgress;
                                }
                            }
                        }

                        if (SUCCEEDED(itemHr) && moveCopyCompleted)
                        {
                            itemHr = RunWithCircuitBreaker(breaker,
                                                           getSourceCircuitBreakerConnectionId(index),
                                                           destinationCircuitBreakerConnectionId,
                                                           [&]() noexcept -> HRESULT
                            {
#ifdef ENABLE_TESTS
                                MaybePauseBeforeBridgeMoveSourceCleanupForSelfTest();
#endif
                                if (moveBridge)
                                {
                                    moveBridge->cookie = static_cast<void*>(&cookie);
                                }
                                return moveBridge ? moveBridge->DeleteCopiedSourceForMove(sourceText, destinationItemText)
                                                  : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                            });

                            if (FAILED(itemHr))
                            {
                                failedDuringMoveDelete = true;
                            }
                        }
                    }
                    else
                    {
                        itemHr = RunWithCircuitBreaker(breaker,
                                                       getSourceCircuitBreakerConnectionId(index),
                                                       destinationCircuitBreakerConnectionId,
                                                       [&]() noexcept -> HRESULT
                        {
                            FileSystemOptions options{};
                            InitializeFileSystemOptions(options);
                            return _fileSystem->MoveItem(
                                sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                        });
                    }
                }
                else if (_operation == FILESYSTEM_DELETE)
                {
                    itemHr = RunWithCircuitBreaker(breaker, getSourceCircuitBreakerConnectionId(index), {}, [&]() noexcept -> HRESULT {
                        return _fileSystem->DeleteItem(sourceText.c_str(), itemFlags, nullptr, this, static_cast<void*>(&cookie));
                    });
                }

                const PerItemInFlightFinishResult finishedCall = FinishPerItemInFlightCall(*this, &cookie);
                callCompletedItems                             = finishedCall.completedItems;
                callCompletedBytes                             = finishedCall.completedBytes;
                callTotalItems                                 = finishedCall.totalItems;

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

                        const bool precalcTotalAvailable = _preCalcCompleted.load(std::memory_order_acquire) && _progressTotalItems > 0;
                        if (! precalcTotalAvailable)
                        {
                            const uint64_t mappedTotalItems = _perItemTotalEntryCount + finishedCall.aggregate.totalItems;
                            if (mappedTotalItems > 0)
                            {
                                const uint64_t clampedTotal =
                                    std::min<uint64_t>(mappedTotalItems, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()));
                                _progressTotalItems = (std::max)(_progressTotalItems, static_cast<unsigned long>(clampedTotal));
                            }
                        }
                    }

                    const uint64_t mapped   = _perItemCompletedBytes + finishedCall.aggregate.completedBytes;
                    _progressCompletedBytes = (std::max)(_progressCompletedBytes, mapped);
                    PublishProgressCountersLocked(*this);
                }

                const bool cancelled = itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || itemHr == E_ABORT;
                if (cancelled)
                {
                    ClearConflictPrompt(*this);
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                if (itemHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
                {
                    itemPartiallySkipped = true;
                    hadSkippedItems      = true;
                    itemSucceeded        = true;
                    break;
                }

                if (SUCCEEDED(itemHr))
                {
                    if (useCrossFileSystemBridge && bridgeRootDirectoryReparseSkipped)
                    {
                        LogDiagnostic(DiagnosticSeverity::Warning,
                                      HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                                      L"bridge.reparse.skip",
                                      L"Skipped root directory reparse point during bridge operation.",
                                      sourceText,
                                      destinationItemText);
                        itemSkipped     = true;
                        hadSkippedItems = true;
                        break;
                    }

                    if (useCrossFileSystemBridge && bridgeSkippedDirectoryReparseCount > 0)
                    {
                        const std::wstring skipMessage = std::format(L"Skipped {:L} directory reparse point{:s} during bridge operation.",
                                                                     bridgeSkippedDirectoryReparseCount,
                                                                     bridgeSkippedDirectoryReparseCount == 1ul ? L"" : L"s");
                        LogDiagnostic(DiagnosticSeverity::Warning,
                                      HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                                      L"bridge.reparse.skip",
                                      skipMessage,
                                      sourceText,
                                      destinationItemText);
                        itemPartiallySkipped = true;
                        hadSkippedItems      = true;
                    }

                    itemSucceeded = true;
                    break;
                }

                // If the caller explicitly requested continue-on-error, preserve legacy behavior.
                if (continueOnError)
                {
                    auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, &cookie, sourceText, destinationItemText);
                    LogDiagnostic(DiagnosticSeverity::Warning,
                                  itemHr,
                                  L"item.continueOnError",
                                  L"Item failed and was skipped due continue-on-error.",
                                  diagnosticSource,
                                  diagnosticDestination);
                    itemSkipped     = true;
                    hadSkippedItems = true;
                    break;
                }

                const FileSystemOperation bucketOperation = failedDuringMoveDelete ? FILESYSTEM_DELETE : _operation;
                const wil::com_ptr<IFileSystemIO>& bucketFileSystemIo =
                    failedDuringMoveDelete ? fileSystemIo : (useCrossFileSystemBridge ? destinationFileSystemIo : fileSystemIo);
                const bool unsupportedReparseHint = bridgeUnsupportedDirectoryReparse;

                const ConflictBucket bucket =
                    ClassifyConflictBucket(bucketOperation, itemFlags, bucketFileSystemIo, itemHr, sourceText, destinationItemText, unsupportedReparseHint);
                if (bucket == ConflictBucket::RecycleBinFailed)
                {
                    auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, &cookie, sourceText, destinationItemText);
                    LogDiagnostic(DiagnosticSeverity::Error,
                                  itemHr,
                                  L"delete.recycleBin.item",
                                  L"Recycle Bin delete failed for item.",
                                  diagnosticSource,
                                  diagnosticDestination);
                }

                const size_t bucketIndex = static_cast<size_t>(bucket);

                std::optional<ConflictAction> cached   = LoadConflictDecisionFromCache(*this, bucket);
                const bool ignoreCachedDecisionForItem = cached.has_value() && IsModifierConflictAction(cached.value()) &&
                                                         bucketIndex < cachedModifierAttempts.size() &&
                                                         cachedModifierAttempts[bucketIndex] >= kMaxCachedModifierAttemptsPerBucket;
                if (ignoreCachedDecisionForItem)
                {
                    // Only this item falls back to a fresh prompt; the cached apply-to-all
                    // decision stays valid for every other item in the task.
                    cached.reset();
                }
                ConflictAction action = cached.value_or(ConflictAction::None);

                if (action == ConflictAction::None)
                {
                    const bool canRetryBucket = IsRetryableConflictBucket(bucket);
                    const bool allowRetry     = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u;
                    const bool retryFailed    = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] != 0u;

                    const ConflictPromptBeginResult promptBegin = BeginConflictPrompt(
                        *this, &cookie, bucket, itemHr, sourceText, destinationItemText, allowRetry, retryFailed, ignoreCachedDecisionForItem);
                    action = promptBegin.action;
                    if (promptBegin.ownsPrompt)
                    {
                        const auto decision = WaitForConflictDecision(*this, &cookie, bucket);
                        action              = decision.first;
                    }
                }

                if (action == ConflictAction::Overwrite)
                {
                    if (bucketIndex < cachedModifierAttempts.size())
                    {
                        ++cachedModifierAttempts[bucketIndex];
                    }
                    itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                    continue;
                }

                if (action == ConflictAction::ReplaceReadOnly)
                {
                    if (bucketIndex < cachedModifierAttempts.size())
                    {
                        ++cachedModifierAttempts[bucketIndex];
                    }
                    itemFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
                    continue;
                }

                if (action == ConflictAction::PermanentDelete)
                {
                    if (bucketIndex < cachedModifierAttempts.size())
                    {
                        ++cachedModifierAttempts[bucketIndex];
                    }
                    itemFlags = static_cast<FileSystemFlags>(itemFlags & ~FILESYSTEM_FLAG_USE_RECYCLE_BIN);
                    continue;
                }

                if (action == ConflictAction::Retry)
                {
                    if (bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u)
                    {
                        retryCounts[bucketIndex] = 1u;

                        if (bucket == ConflictBucket::SharingViolation)
                        {
                            Sleep(750);
                        }

                        continue;
                    }

                    action = ConflictAction::Skip;
                }

                if (action == ConflictAction::Skip)
                {
                    auto [diagnosticSource, diagnosticDestination] = GetMostSpecificPathsForDiagnostics(*this, &cookie, sourceText, destinationItemText);
                    LogDiagnostic(DiagnosticSeverity::Warning,
                                  itemHr,
                                  L"item.conflict.skip",
                                  L"Conflict action Skip item selected.",
                                  diagnosticSource,
                                  diagnosticDestination);
                    itemSkipped     = true;
                    hadSkippedItems = true;
                    break;
                }

                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            if (itemSkipped)
            {
                if (preCalcBytesForItem > 0)
                {
                    std::scoped_lock lock(_progressMutex);
                    _progressTotalBytes = (_progressTotalBytes >= preCalcBytesForItem) ? (_progressTotalBytes - preCalcBytesForItem) : 0;
                    // If pre-calc bytes were counted into total, and the user later skips the item,
                    // ensure we don't end up reporting "completed > total" (progress > 100%).
                    _progressCompletedBytes = (std::min)(_progressCompletedBytes, _progressTotalBytes);
                    PublishProgressCountersLocked(*this);
                }
            }
            else if (itemSucceeded || itemPartiallySkipped)
            {
                const uint64_t bytesForItem = (preCalcBytesForItem > 0) ? preCalcBytesForItem : callCompletedBytes;
                if (_perItemCompletedBytes > std::numeric_limits<uint64_t>::max() - bytesForItem)
                {
                    ClearConflictPrompt(*this);
                    return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                }

                _perItemCompletedBytes += bytesForItem;
            }

            _perItemCompletedItems = static_cast<unsigned long>(std::min<uint64_t>(static_cast<uint64_t>(index + 1u), static_cast<uint64_t>(ULONG_MAX)));

            const PerItemInFlightAggregate inFlightAggregate = getPerItemInFlightAggregate();
            StorePublishedTopLevelCompletionSnapshot(*this, MarkTopLevelItemCompleted(*this, index));

            {
                std::scoped_lock lock(_progressMutex);
                _progressCompletedItems = _perItemCompletedItems;
                const uint64_t mapped   = _perItemCompletedBytes + inFlightAggregate.completedBytes;
                _progressCompletedBytes = (std::max)(_progressCompletedBytes, mapped);
                PublishProgressCountersLocked(*this);
            }
        }

        ClearConflictPrompt(*this);

        if (hadSkippedItems || _observedSkipAction.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    if ((_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE) && _destinationFileSystem)
    {
        // Cross-filesystem bridge is only implemented in per-item mode.
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
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

    if (_operation == FILESYSTEM_COPY)
    {
        if (destinationFolder.empty())
        {
            return E_INVALIDARG;
        }

        FileSystemOptions options{};
        InitializeFileSystemOptions(options);
        const HRESULT operationHr = _fileSystem->CopyItems(pathArray, count, destinationFolder.c_str(), _flags, &options, this, nullptr);
        if (operationHr == S_OK && _observedSkipAction.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
        return operationHr;
    }

    if (_operation == FILESYSTEM_MOVE)
    {
        if (destinationFolder.empty())
        {
            return E_INVALIDARG;
        }

        FileSystemOptions options{};
        InitializeFileSystemOptions(options);
        const HRESULT operationHr = _fileSystem->MoveItems(pathArray, count, destinationFolder.c_str(), _flags, &options, this, nullptr);
        if (operationHr == S_OK && _observedSkipAction.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
        return operationHr;
    }

    if (_operation == FILESYSTEM_DELETE)
    {
        const HRESULT operationHr = _fileSystem->DeleteItems(pathArray, count, _flags, nullptr, this, nullptr);
        if (operationHr == S_OK && _observedSkipAction.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
        return operationHr;
    }

    return E_NOTIMPL;
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

#include "FolderWindow.FileOperations.State.Diagnostics.Part.cpp"
#include "FolderWindow.FileOperations.State.Queue.Part.cpp"
#include "FolderWindow.FileOperations.State.Runtime.Part.cpp"
