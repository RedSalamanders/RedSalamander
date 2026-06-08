#include "FolderWindow.FileOperationsInternal.h"

#include "ConnectionProfileUtils.h"
#include "FolderWindow.FileOperations.IssuesPane.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "SessionState.h"
#include "SettingsHotReload.h"
#include "SettingsSave.h"
#include "SettingsStore.h"

#include <algorithm>
#include <array>
#include <cassert>
#include <chrono>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <deque>
#include <functional>
#include <iterator>
#include <psapi.h>
#include <shellapi.h>
#include <system_error>
#include <thread>
#include <utility>

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

namespace
{
using Task = FolderWindow::FileOperationState::Task;

#ifdef ENABLE_TESTS
std::atomic<unsigned int> g_fileOpsBridgePipelineMode{static_cast<unsigned int>(FileOpsBridgePipelineMode::Default)};
std::atomic<unsigned int> g_fileOpsBridgeProducerDelayMs{0};

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
    std::scoped_lock lock(task._conflictMutex);
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

    std::sort(files.begin(), files.end(), [](const DiagnosticsFileForCleanup& left, const DiagnosticsFileForCleanup& right) {
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

constexpr unsigned int kDefaultPreCalcMaxWorkers         = 4u;
constexpr unsigned int kMaxPreCalcWorkersSetting         = 8u;
constexpr unsigned int kDefaultCrossFsBridgeBufferSizeKB = 4096u;
constexpr unsigned int kMinCrossFsBridgeBufferSizeKB     = 512u;
constexpr unsigned int kMaxCrossFsBridgeBufferSizeKB     = 16384u;
constexpr uint64_t kDefaultBandwidthLimitBytesPerSecond  = 0;

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

[[nodiscard]] unsigned long ResolveAdaptiveCrossFsBridgeBufferBytes(unsigned long configuredBytes,
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
    if (configuredBytes != defaultBytes || sourcePath == nullptr || destinationPath == nullptr)
    {
        return configuredBytes;
    }

    unsigned long resolvedBytes = 0;
    bool sawHint                = false;
    const auto applyHints       = [&](IFileSystem& fileSystem, const wchar_t* path, FileSystemTransferEndpoint endpoint) noexcept
    {
        FileSystemTransferHints hints{};
        hints.sizeBytes  = sizeof(hints);
        const HRESULT hr = fileSystem.GetTransferHints(path, operationType, endpoint, &hints);
        if (FAILED(hr) || hints.preferredBufferBytes == 0)
        {
            return;
        }

        const unsigned long hintedBytes = ClampCrossFsBridgeBufferBytes(hints.preferredBufferBytes);
        if (! sawHint)
        {
            resolvedBytes = hintedBytes;
            sawHint       = true;
            return;
        }

        resolvedBytes = (std::max)(resolvedBytes, hintedBytes);
    };

    applyHints(sourceFileSystem, sourcePath, FILESYSTEM_TRANSFER_SOURCE_READ);
    applyHints(destinationFileSystem, destinationPath, FILESYSTEM_TRANSFER_DESTINATION_WRITE);
    return sawHint ? resolvedBytes : configuredBytes;
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

    if (! settings.fileOperations.has_value())
    {
        return;
    }

    const Common::Settings::FileOperationsSettings defaults{};
    const auto& fileOperations = settings.fileOperations.value();
    const bool hasNonDefault =
        fileOperations.autoDismissSuccess != defaults.autoDismissSuccess || fileOperations.preCalcEnabled != defaults.preCalcEnabled ||
        fileOperations.preCalcMaxWorkers != defaults.preCalcMaxWorkers || fileOperations.crossFsBridgeBufferSizeKB != defaults.crossFsBridgeBufferSizeKB ||
        fileOperations.defaultBandwidthLimitBytesPerSecond != defaults.defaultBandwidthLimitBytesPerSecond ||
        fileOperations.maxDiagnosticsLogFiles != defaults.maxDiagnosticsLogFiles || fileOperations.diagnosticsInfoEnabled != defaults.diagnosticsInfoEnabled ||
        fileOperations.diagnosticsDebugEnabled != defaults.diagnosticsDebugEnabled || fileOperations.maxIssueReportFiles.has_value() ||
        fileOperations.maxDiagnosticsInMemory.has_value() || fileOperations.maxDiagnosticsPerFlush.has_value() ||
        fileOperations.diagnosticsFlushIntervalMs.has_value() || fileOperations.diagnosticsCleanupIntervalMs.has_value() ||
        ! fileOperations.issuesPaneSortColumnId.empty() || fileOperations.issuesPaneSortDescending != defaults.issuesPaneSortDescending ||
        ! fileOperations.issuesPaneGridLayout.empty();
    if (! hasNonDefault)
    {
        settings.fileOperations.reset();
    }
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
        case ERROR_CONNECTION_ABORTED:
        case ERROR_SEM_TIMEOUT:
        case ERROR_TIMEOUT: return true;
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

        std::unique_lock lock(job->doneMutex);
        job->doneCv.wait(lock, [&]() noexcept { return job->done; });
    }

    void NotifyWorkAvailable() noexcept
    {
        _cv.notify_all();
    }

    void Shutdown() noexcept
    {
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
        }

        // Ensure any thread blocked in WaitJob can proceed during teardown.
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

    void workerMain(std::stop_token stopToken) noexcept
    {
        [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();

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
                    job->inFlight -= 1;
                }
                cleanupJobsLocked();
            }

            _cv.notify_all();
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
} // namespace

FolderWindow::FileOperationState::Task::Task(FileOperationState& state) noexcept : _state(&state), _folderWindow(&state._owner)
{
    _conflictDecisionEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
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
        _perf.progressCallbackUs += callbackUs;
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

    const auto clearConflictPrompt = [&](std::optional<std::pair<ConflictBucket, ConflictAction>> cachedDecision = std::nullopt) noexcept
    {
        {
            std::scoped_lock lock(_conflictMutex);
            if (cachedDecision.has_value())
            {
                _conflictDecisionCache[static_cast<size_t>(cachedDecision->first)] = cachedDecision->second;
            }
            _conflictPrompt        = {};
            _conflictOwnerThreadId = 0;
            _conflictDecisionAction.reset();
            _conflictDecisionApplyToAll = false;
        }

        if (_conflictDecisionEvent)
        {
            static_cast<void>(ResetEvent(_conflictDecisionEvent.get()));
        }

        _conflictCv.notify_all();
    };

    const auto getMostSpecificPathsForDiagnostics =
        [&](const PerItemCallbackCookie* perItemCookie, std::wstring_view sourceFallback, std::wstring_view destinationFallback) noexcept
    { return GetMostSpecificPathsForDiagnostics(*this, perItemCookie, sourceFallback, destinationFallback); };

    const auto setCachedDecision = [&](ConflictBucket bucket, ConflictAction decision) noexcept
    {
        std::scoped_lock lock(_conflictMutex);
        _conflictDecisionCache[static_cast<size_t>(bucket)] = decision;
    };

    const auto getCachedDecision = [&](ConflictBucket bucket) noexcept -> std::optional<ConflictAction>
    {
        std::scoped_lock lock(_conflictMutex);
        return _conflictDecisionCache[static_cast<size_t>(bucket)];
    };

    const auto setConflictPromptLocked = [&](const PerItemCallbackCookie* perItemCookie,
                                             ConflictBucket bucket,
                                             HRESULT promptStatus,
                                             std::wstring_view sourceFallback,
                                             std::wstring_view destinationFallback,
                                             bool allowRetry,
                                             bool retryFailed) noexcept
    {
        auto [promptSourcePath, promptDestinationPath] = GetMostSpecificPathsForDiagnostics(*this, perItemCookie, sourceFallback, destinationFallback);

        auto addAction = [&](ConflictAction conflictAction) noexcept
        {
            if (_conflictPrompt.actionCount < _conflictPrompt.actions.size())
            {
                _conflictPrompt.actions[_conflictPrompt.actionCount] = conflictAction;
                ++_conflictPrompt.actionCount;
            }
        };

        if (_conflictDecisionEvent)
        {
            static_cast<void>(ResetEvent(_conflictDecisionEvent.get()));
        }

        _conflictPrompt                   = {};
        _conflictPrompt.active            = true;
        _conflictPrompt.bucket            = bucket;
        _conflictPrompt.status            = promptStatus;
        _conflictPrompt.sourcePath        = std::move(promptSourcePath);
        _conflictPrompt.destinationPath   = std::move(promptDestinationPath);
        _conflictPrompt.applyToAllChecked = false;
        _conflictPrompt.retryFailed       = retryFailed;
        _conflictPrompt.actionCount       = 0;
        _conflictOwnerThreadId            = GetCurrentThreadId();

        LogDiagnostic(DiagnosticSeverity::Warning,
                      promptStatus,
                      L"item.conflict.prompt",
                      retryFailed ? L"Conflict prompt shown after retry cap reached." : L"Conflict prompt shown for item.",
                      _conflictPrompt.sourcePath,
                      _conflictPrompt.destinationPath);

        switch (bucket)
        {
            case ConflictBucket::Exists: addAction(ConflictAction::Overwrite); break;
            case ConflictBucket::ReadOnly: addAction(ConflictAction::ReplaceReadOnly); break;
            case ConflictBucket::RecycleBinFailed: addAction(ConflictAction::PermanentDelete); break;
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
            addAction(ConflictAction::Retry);
        }

        addAction(ConflictAction::Skip);
        addAction(ConflictAction::SkipAll);
        addAction(ConflictAction::Cancel);

        _conflictDecisionAction.reset();
        _conflictDecisionApplyToAll = false;
    };

    const auto waitForConflictDecision = [&](ConflictBucket bucket) noexcept -> std::pair<ConflictAction, bool>
    {
        const uint64_t perfStartUs   = PerfNowUs();
        const auto perfCallbackScope = wil::scope_exit([&] noexcept
        {
            const uint64_t waitUs = PerfElapsedUs(perfStartUs);
            _perf.conflictWaitUs += waitUs;
            ++_perf.conflictPromptCount;
            NoteConflictWorkerWait(*this, cookie, waitUs);
            if (Debug::Perf::IsCaptureEnabled())
            {
                Debug::Perf::Emit(L"FileOps.Conflict.WaitUs", L"", waitUs, 0u, 0u, S_OK);
            }
        });

        if (! _conflictDecisionEvent)
        {
            clearConflictPrompt();
            return {ConflictAction::Cancel, false};
        }

        for (;;)
        {
            if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
            {
                clearConflictPrompt();
                return {ConflictAction::Cancel, false};
            }

            const DWORD wait = WaitForSingleObject(_conflictDecisionEvent.get(), 50);
            if (wait == WAIT_OBJECT_0)
            {
                break;
            }
        }

        ConflictAction decision = ConflictAction::Cancel;
        bool applyToAll         = false;
        {
            std::scoped_lock lock(_conflictMutex);
            decision   = _conflictDecisionAction.value_or(ConflictAction::Cancel);
            applyToAll = _conflictDecisionApplyToAll;
        }

        if (applyToAll && decision != ConflictAction::Retry && decision != ConflictAction::Cancel && decision != ConflictAction::None)
        {
            clearConflictPrompt(std::pair{bucket, decision});
            return {decision, true};
        }

        clearConflictPrompt();
        return {decision, applyToAll};
    };

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
        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(perItemCookie, sourceText, destinationText);
        LogDiagnostic(
            DiagnosticSeverity::Error, status, L"delete.recycleBin.item", L"Recycle Bin delete failed for item.", diagnosticSource, diagnosticDestination);
    }

    const size_t bucketIndex = static_cast<size_t>(bucket);

    ConflictAction decision = getCachedDecision(bucket).value_or(ConflictAction::None);
    if (decision == ConflictAction::None)
    {
        const bool canRetryBucket = bucket != ConflictBucket::UnsupportedReparse;
        bool allowRetry           = canRetryBucket;
        bool retryFailed          = false;
        if (perItemCookie != nullptr && bucketIndex < perItemCookie->issueRetryCounts.size())
        {
            allowRetry  = canRetryBucket && perItemCookie->issueRetryCounts[bucketIndex] == 0u;
            retryFailed = canRetryBucket && perItemCookie->issueRetryCounts[bucketIndex] != 0u;
        }

        {
            std::unique_lock lock(_conflictMutex);
            setConflictPromptLocked(perItemCookie, bucket, status, sourceText, destinationText, allowRetry, retryFailed);
        }

        const auto result = waitForConflictDecision(bucket);
        decision          = result.first;
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
        case ConflictAction::SkipAll:
        {
            auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(perItemCookie, sourceText, destinationText);
            LogDiagnostic(DiagnosticSeverity::Warning,
                          status,
                          L"item.conflict.skipAll",
                          L"Conflict action Skip all similar conflicts selected.",
                          diagnosticSource,
                          diagnosticDestination);
            setCachedDecision(bucket, ConflictAction::Skip);
            _observedSkipAction.store(true, std::memory_order_release);
            *action = FileSystemIssueAction::Skip;
            return S_OK;
        }
        case ConflictAction::Skip:
        {
            auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(perItemCookie, sourceText, destinationText);
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
        ++_perf.preCalcCallbackCount;
        _perf.preCalcCallbackUs += PerfElapsedUs(perfStartUs);
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
                _perf.preCalcLockWaitUs += PerfElapsedUs(totalsLockWaitStartUs);
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
        _perf.preCalcLockWaitUs += PerfElapsedUs(progressLockWaitStartUs);
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
    const auto perfCallbackScope = wil::scope_exit([&] noexcept { _perf.preCalcUs += PerfElapsedUs(perfStartUs); });

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

        for (FileInfo* entry = head; entry;)
        {
            if ((entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && (entry->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0 &&
                entry->FileNameSize >= sizeof(wchar_t))
            {
                const size_t charCount = entry->FileNameSize / sizeof(wchar_t);
                std::wstring_view name(entry->FileName, charCount);
                if (name != L"." && name != L"..")
                {
                    childDirectories.push_back(JoinFolderAndLeaf(path, name));
                }
            }

            if (entry->NextEntryOffset == 0)
            {
                break;
            }

            entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
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
    _preCalcWorkerCountUsed.store(workerCount, std::memory_order_release);

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
            workers.emplace_back([&]() noexcept { workerLoop(); });
        }
    }
    else
    {
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
        _preCalcCompleted.store(true, std::memory_order_release);

        // Update progress totals if we got valid data
        if (finalTotalBytes > 0 || finalTotalFiles > 0 || finalTotalDirs > 0)
        {
            std::scoped_lock lock(_progressMutex);
            _progressTotalBytes = finalTotalBytes;
            _progressTotalItems = static_cast<unsigned long>(std::min(finalTotalFiles + finalTotalDirs, static_cast<uint64_t>(ULONG_MAX)));
            PublishProgressCountersLocked(*this);
        }
    }
}

void FolderWindow::FileOperationState::Task::ThreadMain(std::stop_token stopToken) noexcept
{
    _stopToken                   = stopToken;
    [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();
    [[maybe_unused]] const std::stop_callback stopWake(stopToken,
                                                       [this]() noexcept
    {
        _pauseCv.notify_all();
        _conflictCv.notify_all();
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
                  std::format(L"Task started (op={}, mode={}, sources={}, flags=0x{:08X}, preCalc={}, waitForOthers={}).",
                              OperationToString(_operation),
                              _executionMode == ExecutionMode::PerItem ? L"perItem" : L"bulkItems",
                              _sourcePaths.size(),
                              static_cast<unsigned long>(static_cast<uint32_t>(_flags)),
                              _enablePreCalc ? L"on" : L"off",
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

    // Run pre-calculation phase while holding queue slot
    RunPreCalculation();

    const ULONGLONG afterPreCalcTick = GetTickCount64();
    if (Debug::Perf::IsCaptureEnabled())
    {
        const ULONGLONG preStartTick = _preCalcStartTick.load(std::memory_order_acquire);
        if (preStartTick > 0)
        {
            const ULONGLONG elapsedMs      = (afterPreCalcTick >= preStartTick) ? (afterPreCalcTick - preStartTick) : 0;
            const uint64_t durationUs      = (_perf.preCalcUs > 0) ? _perf.preCalcUs : static_cast<uint64_t>(elapsedMs) * 1000ull;
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

    // Check if cancelled during pre-calc
    if (_cancelled.load(std::memory_order_acquire))
    {
        _enteredOperation.store(false, std::memory_order_release);
        _enteredOperationTick.store(0, std::memory_order_release);
        _state->LeaveOperation();
        _resultHr.store(HRESULT_FROM_WIN32(ERROR_CANCELLED), std::memory_order_release);
        _state->PostCompleted(*this);
        return;
    }

    const HRESULT hr = ExecuteOperation();
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

        const PublishedProgressSnapshot progressSnapshot = LoadPublishedProgressSnapshot(*this);
        auto perfStats                                   = _perf;
        perfStats.bridgeDirectoryEnsureCount             = _bridgeDirectoryEnsureCount.load(std::memory_order_acquire);
        perfStats.bridgeFileAdmissionCount               = _bridgeFileAdmissionCount.load(std::memory_order_acquire);
        perfStats.bridgeFileStartedBeforeProducerDone    = _bridgeFileStartedBeforeProducerDone.load(std::memory_order_acquire);
        perfStats.bridgeAdmissionMaxQueueDepth           = _bridgeAdmissionMaxQueueDepth.load(std::memory_order_acquire);
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
            std::scoped_lock lock(_conflictMutex);
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
            L"itemCompletedUs={} conflictWaitUs={} pauseWaitUs={}",
            _taskId,
            OperationToString(_operation),
            desired,
            effective,
            sourceCount,
            progressSnapshot.completedItems,
            perfStats.queueWaitUs,
            perfStats.schedulerWaitUs,
            perfStats.schedulerWaitForWorkUs,
            perfStats.bridgeCopyUs,
            perfStats.bridgeReadUs,
            perfStats.bridgeWriteUs,
            perfStats.bridgeReaderWaitUs,
            perfStats.bridgeWriterWaitUs,
            perfStats.preCalcUs,
            perfStats.progressCallbackUs,
            perfStats.progressFirstCallbackDelayMs,
            progressStreamMaxGapMs,
            progressStreamMaxGapBytes,
            perfStats.bridgeDirectoryEnsureCount,
            perfStats.bridgeFileAdmissionCount,
            perfStats.bridgeFileStartedBeforeProducerDone,
            perfStats.bridgeAdmissionMaxQueueDepth,
            perfStats.itemCompletedCallbackUs,
            perfStats.conflictWaitUs,
            perfStats.pauseWaitUs);
        Debug::Perf::Emit(L"FileOps.Operation", detail, durationUs, progressSnapshot.completedBytes, progressSnapshot.progressCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.CopyUs", L"", perfStats.bridgeCopyUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.ReadUs", L"", perfStats.bridgeReadUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.WriteUs", L"", perfStats.bridgeWriteUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(
            L"FileOps.Bridge.ReaderWaitUs", L"", perfStats.bridgeReaderWaitUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(
            L"FileOps.Bridge.WriterWaitUs", L"", perfStats.bridgeWriterWaitUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.DirectoryEnsureCount", L"", 0u, perfStats.bridgeDirectoryEnsureCount, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.FileAdmissionCount", L"", 0u, perfStats.bridgeFileAdmissionCount, 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.Bridge.FileStartedBeforeProducerDone", L"", 0u, perfStats.bridgeFileStartedBeforeProducerDone, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Bridge.AdmissionMaxQueueDepth", L"", 0u, perfStats.bridgeAdmissionMaxQueueDepth, 0u, hr);
        Debug::Perf::Emit(L"FileOps.PreCalc.TotalUs", L"", perfStats.preCalcUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(L"FileOps.PreCalc.CallbackCount", L"", 0u, perfStats.preCalcCallbackCount, 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.PreCalc.CallbackUs", L"", perfStats.preCalcCallbackUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(
            L"FileOps.PreCalc.LockWaitUs", L"", perfStats.preCalcLockWaitUs, progressSnapshot.completedBytes, progressSnapshot.completedItems, hr);
        Debug::Perf::Emit(
            L"FileOps.Progress.CallbackUs", L"", perfStats.progressCallbackUs, progressSnapshot.completedBytes, progressSnapshot.progressCallbackCount, hr);
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
        Debug::Perf::Emit(L"FileOps.Scheduler.WaitForWorkUs", L"", perfStats.schedulerWaitForWorkUs, 0u, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Scheduler.ProcessIndexUs", L"", perfStats.schedulerProcessIndexUs, 0u, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Scheduler.DequeueAttempts", L"", 0u, perfStats.schedulerDequeueAttempts, 0u, hr);
        Debug::Perf::Emit(L"FileOps.Scheduler.DequeueSuccess", L"", 0u, perfStats.schedulerDequeueSuccess, 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.Conflict.WaitUs", L"", perfStats.conflictWaitUs, progressSnapshot.completedBytes, progressSnapshot.itemCompletedCallbackCount, hr);
        Debug::Perf::Emit(L"FileOps.Conflict.ConvergenceWaitUs",
                          L"",
                          perfStats.conflictConvergenceWaitUs,
                          progressSnapshot.completedBytes,
                          progressSnapshot.itemCompletedCallbackCount,
                          hr);
        Debug::Perf::Emit(L"FileOps.Conflict.PromptCount", L"", 0u, perfStats.conflictPromptCount, 0u, hr);
        Debug::Perf::Emit(
            L"FileOps.Pause.WaitUs", L"", perfStats.pauseWaitUs, progressSnapshot.completedBytes, progressSnapshot.itemCompletedCallbackCount, hr);

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

    if (_conflictDecisionEvent)
    {
        static_cast<void>(SetEvent(_conflictDecisionEvent.get()));
    }

    _conflictCv.notify_all();

    if (_state)
    {
        ++_perf.queueNotifyAllCount;
        _state->NotifyQueueChanged();
    }

    GetPerItemTaskScheduler().NotifyWorkAvailable();
}

void FolderWindow::FileOperationState::Task::TogglePause() noexcept
{
    const bool nowPaused = ! _paused.load(std::memory_order_acquire);
    _paused.store(nowPaused, std::memory_order_release);
    MarkRateSamplingStateChanged();
    if (! nowPaused)
    {
        _pauseCv.notify_all();
    }

    GetPerItemTaskScheduler().NotifyWorkAvailable();
}

void FolderWindow::FileOperationState::Task::SetDesiredSpeedLimit(uint64_t bytesPerSecond) noexcept
{
    _desiredSpeedLimitBytesPerSecond.store(bytesPerSecond, std::memory_order_release);
}

void FolderWindow::FileOperationState::Task::SetWaitForOthers(bool wait) noexcept
{
    if (_started.load(std::memory_order_acquire))
    {
        return;
    }

    _waitForOthers.store(wait, std::memory_order_release);
    if (_state)
    {
        ++_perf.queueNotifyAllCount;
        _state->NotifyQueueChanged();
    }
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
    const bool wasPaused = _queuePaused.load(std::memory_order_acquire);
    if (wasPaused == paused)
    {
        return;
    }

    _queuePaused.store(paused, std::memory_order_release);
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
    std::scoped_lock lock(_conflictMutex);
    if (! _conflictPrompt.active)
    {
        return;
    }

    _conflictPrompt.applyToAllChecked = ! _conflictPrompt.applyToAllChecked;
}

void FolderWindow::FileOperationState::Task::SubmitConflictDecision(ConflictAction action, bool applyToAllChecked) noexcept
{
    {
        std::scoped_lock lock(_conflictMutex);
        if (! _conflictPrompt.active)
        {
            return;
        }

        _conflictDecisionAction     = action;
        _conflictDecisionApplyToAll = (action == ConflictAction::Retry) ? false : (applyToAllChecked || action == ConflictAction::SkipAll);
    }

    if (_conflictDecisionEvent)
    {
        static_cast<void>(SetEvent(_conflictDecisionEvent.get()));
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

void FolderWindow::FileOperationState::Task::WaitWhilePaused() noexcept
{
    const DWORD currentThreadId = GetCurrentThreadId();

    for (;;)
    {
        const bool shouldPause     = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
        bool shouldWaitForConflict = false;
        {
            std::scoped_lock lock(_conflictMutex);
            shouldWaitForConflict = _conflictPrompt.active && _conflictOwnerThreadId != 0 && _conflictOwnerThreadId != currentThreadId;
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
                return ! stillPaused || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested();
            });
            _perf.pauseWaitUs += PerfElapsedUs(waitStartUs);
            continue;
        }

        const uint64_t waitStartUs = PerfNowUs();
        std::unique_lock lock(_conflictMutex);
        _conflictCv.wait(lock,
                         [&]
        {
            const bool stillPaused        = _paused.load(std::memory_order_acquire) || _queuePaused.load(std::memory_order_acquire);
            const bool waitingForConflict = _conflictPrompt.active && _conflictOwnerThreadId != 0 && _conflictOwnerThreadId != currentThreadId;
            return ! waitingForConflict || stillPaused || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested();
        });
        _perf.conflictConvergenceWaitUs += PerfElapsedUs(waitStartUs);
    }
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
    _perf.pauseWaitUs += PerfElapsedUs(waitStartUs);
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

    const bool continueOnError = (_flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;

    if (_executionMode == ExecutionMode::PerItem)
    {
        wil::com_ptr<IFileSystemIO> fileSystemIo;
        static_cast<void>(_fileSystem->QueryInterface(IID_PPV_ARGS(fileSystemIo.addressof())));

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

        const auto clearConflictPrompt = [&](std::optional<std::pair<ConflictBucket, ConflictAction>> cachedDecision = std::nullopt) noexcept
        {
            {
                std::scoped_lock lock(_conflictMutex);
                if (cachedDecision.has_value())
                {
                    _conflictDecisionCache[static_cast<size_t>(cachedDecision->first)] = cachedDecision->second;
                }
                _conflictPrompt        = {};
                _conflictOwnerThreadId = 0;
                _conflictDecisionAction.reset();
                _conflictDecisionApplyToAll = false;
            }

            if (_conflictDecisionEvent)
            {
                static_cast<void>(ResetEvent(_conflictDecisionEvent.get()));
            }

            _conflictCv.notify_all();
        };

        const auto getMostSpecificPathsForDiagnostics =
            [&](const PerItemCallbackCookie* perItemCookie, std::wstring_view sourceFallback, std::wstring_view destinationFallback) noexcept
        { return GetMostSpecificPathsForDiagnostics(*this, perItemCookie, sourceFallback, destinationFallback); };

        const auto setConflictPromptLocked = [&](const PerItemCallbackCookie* perItemCookie,
                                                 ConflictBucket bucket,
                                                 HRESULT status,
                                                 std::wstring_view sourcePath,
                                                 std::wstring_view destinationPath,
                                                 bool allowRetry,
                                                 bool retryFailed) noexcept
        {
            auto [promptSourcePath, promptDestinationPath] = GetMostSpecificPathsForDiagnostics(*this, perItemCookie, sourcePath, destinationPath);

            auto addAction = [&](ConflictAction action) noexcept
            {
                if (_conflictPrompt.actionCount < _conflictPrompt.actions.size())
                {
                    _conflictPrompt.actions[_conflictPrompt.actionCount] = action;
                    ++_conflictPrompt.actionCount;
                }
            };

            if (_conflictDecisionEvent)
            {
                static_cast<void>(ResetEvent(_conflictDecisionEvent.get()));
            }
            _conflictPrompt                   = {};
            _conflictPrompt.active            = true;
            _conflictPrompt.bucket            = bucket;
            _conflictPrompt.status            = status;
            _conflictPrompt.sourcePath        = std::move(promptSourcePath);
            _conflictPrompt.destinationPath   = std::move(promptDestinationPath);
            _conflictPrompt.applyToAllChecked = false;
            _conflictPrompt.retryFailed       = retryFailed;

            _conflictPrompt.actionCount = 0;
            _conflictOwnerThreadId      = GetCurrentThreadId();

            LogDiagnostic(DiagnosticSeverity::Warning,
                          status,
                          L"item.conflict.prompt",
                          retryFailed ? L"Conflict prompt shown after retry cap reached." : L"Conflict prompt shown for item.",
                          _conflictPrompt.sourcePath,
                          _conflictPrompt.destinationPath);

            switch (bucket)
            {
                case ConflictBucket::Exists: addAction(ConflictAction::Overwrite); break;
                case ConflictBucket::ReadOnly: addAction(ConflictAction::ReplaceReadOnly); break;
                case ConflictBucket::RecycleBinFailed: addAction(ConflictAction::PermanentDelete); break;
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
                addAction(ConflictAction::Retry);
            }
            addAction(ConflictAction::Skip);
            addAction(ConflictAction::SkipAll);
            addAction(ConflictAction::Cancel);

            _conflictDecisionAction.reset();
            _conflictDecisionApplyToAll = false;
        };

        const auto waitForConflictDecision = [&](const void* cookieKey, ConflictBucket bucket) noexcept -> std::pair<ConflictAction, bool>
        {
            const uint64_t perfStartUs   = PerfNowUs();
            const auto perfCallbackScope = wil::scope_exit([&] noexcept
            {
                const uint64_t waitUs = PerfElapsedUs(perfStartUs);
                _perf.conflictWaitUs += waitUs;
                ++_perf.conflictPromptCount;
                NoteConflictWorkerWait(*this, cookieKey, waitUs);
                if (Debug::Perf::IsCaptureEnabled())
                {
                    Debug::Perf::Emit(L"FileOps.Conflict.WaitUs", L"", waitUs, 0u, 0u, S_OK);
                }
            });

            if (! _conflictDecisionEvent)
            {
                clearConflictPrompt();
                return {ConflictAction::Cancel, false};
            }

            for (;;)
            {
                if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                {
                    clearConflictPrompt();
                    return {ConflictAction::Cancel, false};
                }

                const DWORD wait = WaitForSingleObject(_conflictDecisionEvent.get(), 50);
                if (wait == WAIT_OBJECT_0)
                {
                    break;
                }
            }

            ConflictAction action = ConflictAction::Cancel;
            bool applyToAll       = false;
            {
                std::scoped_lock lock(_conflictMutex);
                action     = _conflictDecisionAction.value_or(ConflictAction::Cancel);
                applyToAll = _conflictDecisionApplyToAll;
            }

            if (applyToAll && action != ConflictAction::Retry && action != ConflictAction::Cancel && action != ConflictAction::None)
            {
                clearConflictPrompt(std::pair{bucket, action});
                return {action, true};
            }

            clearConflictPrompt();
            return {action, applyToAll};
        };

        const auto getCachedDecision = [&](ConflictBucket bucket) noexcept -> std::optional<ConflictAction>
        {
            std::scoped_lock lock(_conflictMutex);
            return _conflictDecisionCache[static_cast<size_t>(bucket)];
        };

        const auto setCachedDecision = [&](ConflictBucket bucket, ConflictAction action) noexcept
        {
            if (action == ConflictAction::Retry || action == ConflictAction::Cancel || action == ConflictAction::None)
            {
                return;
            }

            if (action == ConflictAction::SkipAll)
            {
                action = ConflictAction::Skip;
            }

            std::scoped_lock lock(_conflictMutex);
            _conflictDecisionCache[static_cast<size_t>(bucket)] = action;
        };

        const auto clearCachedDecision = [&](ConflictBucket bucket) noexcept
        {
            std::scoped_lock lock(_conflictMutex);
            _conflictDecisionCache[static_cast<size_t>(bucket)].reset();
        };

        const auto isModifierConflictAction = [](ConflictAction action) noexcept
        {
            switch (action)
            {
                case ConflictAction::Overwrite:
                case ConflictAction::ReplaceReadOnly:
                case ConflictAction::PermanentDelete: return true;
                case ConflictAction::None:
                case ConflictAction::Retry:
                case ConflictAction::Skip:
                case ConflictAction::SkipAll:
                case ConflictAction::Cancel:
                default: return false;
            }
        };

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
            static constexpr DWORD ProgressIntervalMs() noexcept
            {
                return 200u;
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

            // Total bytes is best-effort: if unknown, keep 0.
            uint64_t totalBytes                         = 0;
            uint64_t completedBytes                     = 0;
            unsigned long skippedDirectoryReparseCount  = 0;
            bool rootDirectoryReparseSkipped            = false;
            bool unsupportedDirectoryReparseEncountered = false;

            std::mutex callbackMutex;
            std::mutex throttleMutex;
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

            std::unique_ptr<std::byte[]> buffer;
            unsigned long bufferBytes = 0;

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
                const uint64_t initialBandwidth      = task._desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                options.sizeBytes                    = sizeof(FileSystemOptions);
                options.bandwidthLimitBytesPerSecond = initialBandwidth;
                bandwidthLimitBytesPerSecond.store(initialBandwidth, std::memory_order_release);

                bufferBytes = ResolveAdaptiveCrossFsBridgeBufferBytes(
                    task._crossFsBridgeBufferBytes, sourceFs, rootSourcePathIn, destinationFs, rootDestinationPathIn, task._operation);
                task._resolvedCrossFsBridgeBufferBytes.store(bufferBytes, std::memory_order_release);
                buffer.reset(new (std::nothrow) std::byte[bufferBytes]);
            }

            CrossFileSystemBridge(const CrossFileSystemBridge&)            = delete;
            CrossFileSystemBridge(CrossFileSystemBridge&&)                 = delete;
            CrossFileSystemBridge& operator=(const CrossFileSystemBridge&) = delete;
            CrossFileSystemBridge& operator=(CrossFileSystemBridge&&)      = delete;

            [[nodiscard]] bool CancelRequested() const noexcept
            {
                return task._cancelled.load(std::memory_order_acquire) || task._stopToken.stop_requested();
            }

            [[nodiscard]] std::wstring MakeTempDestinationPath(std::wstring_view destinationPath, uint64_t progressStreamId) const noexcept
            {
                // Best-effort: keep atomic commit semantics for cross-filesystem transfers by writing to a temp name first.
                wchar_t suffix[80]{};
                constexpr size_t suffixMax = (sizeof(suffix) / sizeof(suffix[0])) - 1u;

                const DWORD pid    = GetCurrentProcessId();
                const DWORD tid    = GetCurrentThreadId();
                const uint64_t now = GetTickCount64();

                const auto r         = std::format_to_n(suffix,
                                                        suffixMax,
                                                        L".rs_tmp_{:08X}_{:08X}_{:016X}_{:X}",
                                                        static_cast<unsigned long>(pid),
                                                        static_cast<unsigned long>(tid),
                                                        static_cast<unsigned long long>(now),
                                                        progressStreamId);
                const size_t written = (r.size < suffixMax) ? r.size : suffixMax;
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

            [[nodiscard]] HRESULT PromoteTempToFinalPath(const std::wstring& tempPath, const std::wstring& destinationPath) noexcept
            {
                BridgeCallback callback(task, &callbackMutex);
                return destinationFs.MoveItem(tempPath.c_str(), destinationPath.c_str(), flags, nullptr, &callback, cookie);
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

            void Throttle(uint64_t bytesSoFar) noexcept
            {
                const uint64_t bandwidthLimit = bandwidthLimitBytesPerSecond.load(std::memory_order_acquire);
                if (bandwidthLimit == 0)
                {
                    return;
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
                    const DWORD sleepMs = remaining > std::numeric_limits<DWORD>::max() ? std::numeric_limits<DWORD>::max() : static_cast<DWORD>(remaining);
                    if (sleepMs > 0)
                    {
                        SleepResponsive(sleepMs);
                    }
                }
            }

            void ThrottleThreadSafe(uint64_t bytesSoFar) noexcept
            {
                if (bandwidthLimitBytesPerSecond.load(std::memory_order_acquire) == 0)
                {
                    return;
                }

                std::scoped_lock lock(throttleMutex);
                Throttle(bytesSoFar);
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
                task._perf.bridgeCopyUs += perf.copyUs;
                task._perf.bridgeReaderWaitUs += perf.readerWaitUs;
                task._perf.bridgeWriterWaitUs += perf.writerWaitUs;
                task._perf.bridgeReadUs += perf.readUs;
                task._perf.bridgeWriteUs += perf.writeUs;

                const uint64_t throughputBytesPerSecond =
                    (perf.copyUs > 0 && transferredBytes > 0) ? static_cast<uint64_t>((transferredBytes * 1000000ull) / perf.copyUs) : 0ull;
                const std::wstring detail =
                    std::format(L"source={} destination={} bytes={} bufferBytes={} progressCalls={} readerWaitUs={} writerWaitUs={} readUs={} writeUs={}",
                                sourcePath,
                                destinationPath,
                                transferredBytes,
                                bufferBytes,
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

            HRESULT EnsureDestinationDirectory(const std::wstring& destinationPath) noexcept
            {
                if (! destinationDirOps)
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }

                unsigned long attributes = 0;
                const HRESULT hrAttr     = destinationIo.GetAttributes(destinationPath.c_str(), &attributes);
                if (SUCCEEDED(hrAttr))
                {
                    if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                    {
                        return S_OK;
                    }

                    if ((flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) == 0)
                    {
                        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                    }

                    // Replace an existing file with a directory.
                    BridgeCallback callback(task, &callbackMutex);
                    const HRESULT hrDelete = destinationFs.DeleteItem(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
                    if (FAILED(hrDelete))
                    {
                        return hrDelete;
                    }
                }

                const HRESULT hrCreate = destinationDirOps->CreateDirectory(destinationPath.c_str());
                if (SUCCEEDED(hrCreate) || hrCreate == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
                {
                    return S_OK;
                }

                return hrCreate;
            }

            void MarkDirectoryReparseSkipped(const std::wstring& sourcePath, const std::wstring& destinationPath, bool isRoot) noexcept
            {
                ++skippedDirectoryReparseCount;
                if (isRoot)
                {
                    rootDirectoryReparseSkipped = true;
                }

                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Warning,
                                   HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
                                   L"bridge.reparse.skip",
                                   isRoot ? L"Skipped root directory reparse point by policy." : L"Skipped directory reparse point by policy.",
                                   sourcePath,
                                   destinationPath);

                const uint64_t callCompleted = (completedBytesAtomic != nullptr) ? completedBytesAtomic->load(std::memory_order_acquire) : completedBytes;
                static_cast<void>(ReportProgress(sourcePath, destinationPath, 0, 0, callCompleted, 0));
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

            [[nodiscard]] HRESULT ValidateDestinationOverwritePolicy(const std::wstring& destinationPath) noexcept
            {
                unsigned long destinationAttributes = 0;
                const HRESULT hrDestAttr            = destinationIo.GetAttributes(destinationPath.c_str(), &destinationAttributes);
                if (SUCCEEDED(hrDestAttr))
                {
                    if ((destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                    {
                        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                    }

                    const bool allowOverwrite = (static_cast<uint32_t>(flags) & static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE)) != 0u;
                    if (! allowOverwrite)
                    {
                        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                    }

                    const bool replaceReadOnly = (static_cast<uint32_t>(flags) & static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY)) != 0u;
                    if (! replaceReadOnly && (destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0)
                    {
                        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
                    }
                }

                return S_OK;
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

                ConnectionConcurrencyLimiter::Permit permit1;
                ConnectionConcurrencyLimiter::Permit permit2;
                const HRESULT hrPermits = AcquireCopyMovePermits(permit1, permit2);
                if (FAILED(hrPermits))
                {
                    return hrPermits;
                }

                const HRESULT hrDestPolicy = ValidateDestinationOverwritePolicy(destinationPath);
                if (FAILED(hrDestPolicy))
                {
                    return hrDestPolicy;
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

                uint64_t fileTotalBytes = 0;
                static_cast<void>(reader->GetSize(&fileTotalBytes));
                if (adoptFileSizeAsTotalWhenUnknown && totalBytes == 0 && fileTotalBytes > 0)
                {
                    totalBytes = fileTotalBytes;
                }

                const std::wstring tempPath = MakeTempDestinationPath(destinationPath, progressStreamId);
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
                    static_cast<FileSystemFlags>(static_cast<uint32_t>(flags) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE) |
                                                 static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));

                wil::com_ptr<IFileWriter> writer;
                hr = destinationIo.CreateFileWriter(tempPath.c_str(), tempFlags, writer.addressof());
                if (FAILED(hr))
                {
                    return hr;
                }
                tempStaged = true;

                uint64_t fileCompletedBytes = 0;
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
                                task.WaitWhilePaused();
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
                                const HRESULT hrRead       = pipelineReader->Read(slots[readIndex].buffer, bufferBytesIn, &bytesRead);
                                readerReadUs.fetch_add(PerfElapsedUs(readStartUs), std::memory_order_relaxed);

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

                if (fileTotalBytes > 0 && fileCompletedBytes != fileTotalBytes)
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

                if (fileTotalBytes > 0 && fileCompletedBytes >= fileTotalBytes)
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
                    Debug::Warning(L"CrossFileSystemBridge: destination writer Commit failed for '{}' via temp '{}' (hr={:#x})",
                                   destinationPath,
                                   tempPath,
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

                hr = PromoteTempToFinalPath(tempPath, destinationPath);
                if (FAILED(hr))
                {
                    Debug::Warning(
                        L"CrossFileSystemBridge: failed to promote temp '{}' to '{}' (hr={:#x})", tempPath, destinationPath, static_cast<unsigned long>(hr));
                    task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                       hr,
                                       L"bridge.promote",
                                       L"Failed to promote the staged bridge destination into the final path.",
                                       sourcePath,
                                       destinationPath);
                    return hr;
                }
                promoted = true;

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
                const uint64_t finalTotal     = fileTotalBytes > 0 ? fileTotalBytes : fileCompletedBytes;
                const uint64_t finalCompleted = fileCompletedBytes;

                hr = ReportProgress(sourcePath, destinationPath, finalTotal, finalCompleted, overallFinal, progressStreamId);
                if (FAILED(hr))
                {
                    return hr;
                }

                return S_OK;
            }

            HRESULT CopyFile(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
            {
                if (! buffer || bufferBytes == 0)
                {
                    return E_OUTOFMEMORY;
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

                HRESULT hr = EnsureDestinationDirectory(destinationPath);
                if (FAILED(hr))
                {
                    return hr;
                }

                wil::com_ptr<IFilesInformation> info;
                hr = sourceFs.ReadDirectoryInfo(sourcePath.c_str(), info.addressof());
                if (FAILED(hr))
                {
                    return hr;
                }

                FileInfo* entry = nullptr;
                hr              = info->GetBuffer(&entry);
                if (FAILED(hr) || entry == nullptr)
                {
                    return hr;
                }

                unsigned long bufferSize = 0;
                hr                       = info->GetBufferSize(&bufferSize);
                if (FAILED(hr) || bufferSize < sizeof(FileInfo))
                {
                    return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }

                std::byte* base = reinterpret_cast<std::byte*>(entry);
                std::byte* end  = base + bufferSize;

                for (;;)
                {
                    task.WaitWhilePaused();
                    if (CancelRequested())
                    {
                        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }

                    const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
                    const std::wstring_view name(entry->FileName, nameChars);

                    const bool isDot = (name == L"." || name == L"..");
                    if (! name.empty() && ! isDot)
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
                                    MarkDirectoryReparseSkipped(childSource, childDest, false);
                                    continue;
                                }

                                // copyReparse requires preserving a link; bridge cannot preserve NTFS reparse payloads.
                                task.LogDiagnostic(FileOperationState::DiagnosticSeverity::Error,
                                                   HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                                   L"bridge.reparse.unsupported",
                                                   L"Cross-filesystem bridge cannot preserve directory reparse payloads.",
                                                   childSource,
                                                   childDest);
                                unsupportedDirectoryReparseEncountered = true;
                                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                            }
                            hr = CopyDirectorySequential(childSource, childDest);
                        }
                        else
                        {
                            hr = CopyFile(childSource, childDest);
                        }

                        if (FAILED(hr))
                        {
                            return hr;
                        }
                    }

                    if (entry->NextEntryOffset == 0)
                    {
                        break;
                    }

                    if (entry->NextEntryOffset < sizeof(FileInfo))
                    {
                        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    }

                    std::byte* next = reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset;
                    if (next < base || next + sizeof(FileInfo) > end)
                    {
                        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    }

                    entry = reinterpret_cast<FileInfo*>(next);
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

                HRESULT hrRoot = EnsureDestinationDirectory(destinationPath);
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
                    auto localBuffer = std::unique_ptr<std::byte[]>(new (std::nothrow) std::byte[bufferBytes]);
                    if (! localBuffer)
                    {
                        hadWorkerFailure.store(true, std::memory_order_release);
                        if (! continueOnError)
                        {
                            stopRequested.store(true, std::memory_order_release);
                            HRESULT expected = S_OK;
                            static_cast<void>(firstFailure.compare_exchange_strong(expected, E_OUTOFMEMORY));
                        }
                        return E_OUTOFMEMORY;
                    }

                    const unsigned long localBufferBytes = bufferBytes;
                    const uint64_t progressStreamId      = static_cast<uint64_t>(workerIndex);
                    const auto recordFailure = [&](HRESULT failure) noexcept
                    {
                        hadWorkerFailure.store(true, std::memory_order_release);
                        const bool cancellation = failure == HRESULT_FROM_WIN32(ERROR_CANCELLED) || failure == E_ABORT;
                        if (cancellation || ! continueOnError)
                        {
                            stopRequested.store(true, std::memory_order_release);
                            HRESULT expected = S_OK;
                            static_cast<void>(firstFailure.compare_exchange_strong(expected, failure));
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
                        stopRequested.store(true, std::memory_order_release);
                        HRESULT expected = S_OK;
                        static_cast<void>(firstFailure.compare_exchange_strong(expected, failure));
                    }
                    workCv.notify_all();
                };

                const auto enqueueWork = [&](WorkItem item) noexcept
                {
                    std::scoped_lock lock(workMutex);
                    workItems.push_back(std::move(item));
                    ++fileAdmissionCount;
                    maxAdmissionQueueDepth = (std::max)(maxAdmissionQueueDepth, static_cast<uint64_t>(workItems.size()));
                    workCv.notify_one();
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

                    HRESULT hr = EnsureDestinationDirectory(currentDest);
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
                    if (FAILED(hr) || entry == nullptr)
                    {
                        recordProducerFailure(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
                        if (! continueOnError)
                        {
                            break;
                        }
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

                    for (;;)
                    {
                        task.WaitWhilePaused();
                        if (CancelRequested())
                        {
                            recordProducerFailure(HRESULT_FROM_WIN32(ERROR_CANCELLED));
                            break;
                        }

                        const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
                        const std::wstring_view name(entry->FileName, nameChars);

                        const bool isDot = (name == L"." || name == L"..");
                        if (! name.empty() && ! isDot)
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
                                        MarkDirectoryReparseSkipped(childSource, childDest, false);
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
                                WorkItem item{};
                                item.source      = std::move(childSource);
                                item.destination = std::move(childDest);
                                enqueueWork(std::move(item));
                            }
                        }

                        if (stopRequested.load(std::memory_order_acquire) || entry->NextEntryOffset == 0)
                        {
                            break;
                        }

                        if (entry->NextEntryOffset < sizeof(FileInfo))
                        {
                            recordProducerFailure(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
                            break;
                        }

                        std::byte* next = reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset;
                        if (next < base || next + sizeof(FileInfo) > end)
                        {
                            recordProducerFailure(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
                            break;
                        }

                        entry = reinterpret_cast<FileInfo*>(next);
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
                task._perf.schedulerWaitUs += PerfElapsedUs(schedulerWall);
                task._perf.schedulerDequeueAttempts +=
                    (schedulerEnd.dequeueAttempts >= schedulerStart.dequeueAttempts) ? (schedulerEnd.dequeueAttempts - schedulerStart.dequeueAttempts) : 0;
                task._perf.schedulerDequeueSuccess +=
                    (schedulerEnd.dequeueSuccess >= schedulerStart.dequeueSuccess) ? (schedulerEnd.dequeueSuccess - schedulerStart.dequeueSuccess) : 0;
                task._perf.schedulerWaitForWorkUs +=
                    (schedulerEnd.waitForWorkUs >= schedulerStart.waitForWorkUs) ? (schedulerEnd.waitForWorkUs - schedulerStart.waitForWorkUs) : 0;
                task._perf.schedulerProcessIndexUs +=
                    (schedulerEnd.processIndexUs >= schedulerStart.processIndexUs) ? (schedulerEnd.processIndexUs - schedulerStart.processIndexUs) : 0;

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

                return CopyDirectoryParallel(sourcePath, destinationPath, withinFolderBudget);
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
                            MarkDirectoryReparseSkipped(sourcePath, destinationPath, true);
                            return S_OK;
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
                    return CopyDirectory(sourcePath, destinationPath);
                }

                return CopyFile(sourcePath, destinationPath);
            }
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
                FileSystemFlags itemFlags = _flags;

                bool itemSucceeded          = false;
                bool itemSkipped            = false;
                uint64_t callCompletedBytes = 0;
                uint64_t callCompletedItems = 0;
                uint64_t callTotalItems     = 0;

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
                        const std::wstring_view leaf = GetPathLeaf(sourceText);
                        if (leaf.empty())
                        {
                            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                        }
                        destinationItemText = JoinFolderAndLeaf(destinationFolderText, leaf);
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

                    callCompletedBytes = 0;
                    callCompletedItems = 0;
                    callTotalItems     = 0;

                    ConnectionCircuitBreaker& breaker = GetConnectionCircuitBreaker();
                    const HRESULT itemHr              = RunWithCircuitBreaker(breaker,
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
                                return bridge.CopyPath(sourceText, destinationItemText);
                            }

                            FileSystemOptions options{};
                            options.sizeBytes                    = sizeof(FileSystemOptions);
                            options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                            return _fileSystem->CopyItem(
                                sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                        }

                        if (_operation == FILESYSTEM_MOVE)
                        {
                            FileSystemOptions options{};
                            options.sizeBytes                    = sizeof(FileSystemOptions);
                            options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                            return _fileSystem->MoveItem(
                                sourceText.c_str(), destinationItemText.c_str(), itemFlags, &options, this, static_cast<void*>(&cookie));
                        }

                        if (_operation == FILESYSTEM_DELETE)
                        {
                            return _fileSystem->DeleteItem(sourceText.c_str(), itemFlags, nullptr, this, static_cast<void*>(&cookie));
                        }

                        return E_NOTIMPL;
                    });

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
                        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
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

                    const ConflictBucket bucket = ClassifyConflictBucket(_operation, itemFlags, fileSystemIo, itemHr, sourceText, destinationItemText, false);
                    if (bucket == ConflictBucket::RecycleBinFailed)
                    {
                        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                        LogDiagnostic(DiagnosticSeverity::Error,
                                      itemHr,
                                      L"delete.recycleBin.item",
                                      L"Recycle Bin delete failed for item.",
                                      diagnosticSource,
                                      diagnosticDestination);
                    }

                    const size_t bucketIndex = static_cast<size_t>(bucket);

                    std::optional<ConflictAction> cached = getCachedDecision(bucket);
                    if (cached.has_value() && isModifierConflictAction(cached.value()) && bucketIndex < cachedModifierAttempts.size() &&
                        cachedModifierAttempts[bucketIndex] >= kMaxCachedModifierAttemptsPerBucket)
                    {
                        clearCachedDecision(bucket);
                        cached.reset();
                    }
                    ConflictAction action = cached.value_or(ConflictAction::None);

                    if (action == ConflictAction::None)
                    {
                        const bool canRetryBucket = bucket != ConflictBucket::UnsupportedReparse;
                        const bool allowRetry     = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u;
                        const bool retryFailed    = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] != 0u;

                        bool owner = false;
                        {
                            std::unique_lock lock(_conflictMutex);

                            const bool cacheableBucket = bucketIndex < _conflictDecisionCache.size();
                            if (cacheableBucket && _conflictDecisionCache[bucketIndex].has_value())
                            {
                                action = _conflictDecisionCache[bucketIndex].value();
                            }
                            else
                            {
                                _conflictCv.wait(lock, [&]() noexcept {
                                    return ! _conflictPrompt.active || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested();
                                });

                                if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                                {
                                    action = ConflictAction::Cancel;
                                }
                                else if (cacheableBucket && _conflictDecisionCache[bucketIndex].has_value())
                                {
                                    action = _conflictDecisionCache[bucketIndex].value();
                                }
                                else
                                {
                                    setConflictPromptLocked(&cookie, bucket, itemHr, sourceText, destinationItemText, allowRetry, retryFailed);
                                    owner = true;
                                }
                            }
                        }

                        if (owner)
                        {
                            const auto decision = waitForConflictDecision(&cookie, bucket);
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

                    if (action == ConflictAction::SkipAll)
                    {
                        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                        LogDiagnostic(DiagnosticSeverity::Warning,
                                      itemHr,
                                      L"item.conflict.skipAll",
                                      L"Conflict action Skip all similar conflicts selected.",
                                      diagnosticSource,
                                      diagnosticDestination);
                        setCachedDecision(bucket, ConflictAction::Skip);
                        itemSkipped = true;
                        hadSkipped.store(true, std::memory_order_release);
                        break;
                    }

                    if (action == ConflictAction::Skip)
                    {
                        auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
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
            _perf.schedulerWaitUs += PerfElapsedUs(schedulerWallStartUs);
            _perf.schedulerDequeueAttempts +=
                (schedulerEnd.dequeueAttempts >= schedulerStart.dequeueAttempts) ? (schedulerEnd.dequeueAttempts - schedulerStart.dequeueAttempts) : 0;
            _perf.schedulerDequeueSuccess +=
                (schedulerEnd.dequeueSuccess >= schedulerStart.dequeueSuccess) ? (schedulerEnd.dequeueSuccess - schedulerStart.dequeueSuccess) : 0;
            _perf.schedulerWaitForWorkUs +=
                (schedulerEnd.waitForWorkUs >= schedulerStart.waitForWorkUs) ? (schedulerEnd.waitForWorkUs - schedulerStart.waitForWorkUs) : 0;
            _perf.schedulerProcessIndexUs +=
                (schedulerEnd.processIndexUs >= schedulerStart.processIndexUs) ? (schedulerEnd.processIndexUs - schedulerStart.processIndexUs) : 0;

            clearConflictPrompt();

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

            FileSystemFlags itemFlags   = _flags;
            uint64_t callCompletedBytes = 0;
            uint64_t callCompletedItems = 0;
            uint64_t callTotalItems     = 0;
            bool moveCopyCompleted      = false;
            uint64_t moveCopiedBytes    = 0;

            for (;;)
            {
                WaitWhilePaused();
                if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
                {
                    clearConflictPrompt();
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                callCompletedBytes = 0;
                callCompletedItems = 0;
                callTotalItems     = 0;

                std::wstring destinationItemText;
                if (_operation == FILESYSTEM_COPY || _operation == FILESYSTEM_MOVE)
                {
                    const std::wstring_view leaf = GetPathLeaf(sourceText);
                    if (leaf.empty())
                    {
                        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                    }
                    destinationItemText = JoinFolderAndLeaf(destinationFolderText, leaf);
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

                if (_operation == FILESYSTEM_COPY)
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
                            bridgeSkippedDirectoryReparseCount = bridge.skippedDirectoryReparseCount;
                            bridgeRootDirectoryReparseSkipped  = bridge.rootDirectoryReparseSkipped;
                            bridgeUnsupportedDirectoryReparse  = bridge.unsupportedDirectoryReparseEncountered;
                            return hr;
                        }

                        FileSystemOptions options{};
                        options.sizeBytes                    = sizeof(FileSystemOptions);
                        options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
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
                                bridgeSkippedDirectoryReparseCount = bridge.skippedDirectoryReparseCount;
                                bridgeRootDirectoryReparseSkipped  = bridge.rootDirectoryReparseSkipped;
                                bridgeUnsupportedDirectoryReparse  = bridge.unsupportedDirectoryReparseEncountered;
                                if (SUCCEEDED(hr))
                                {
                                    moveCopyCompleted = bridgeSkippedDirectoryReparseCount == 0 && ! bridgeRootDirectoryReparseSkipped;
                                    moveCopiedBytes   = bridge.completedBytes;
                                }
                                return hr;
                            });
                        }

                        if (SUCCEEDED(itemHr) && moveCopyCompleted)
                        {
                            // Ensure the in-flight call has the best-known completed-bytes snapshot even when we're only deleting.
                            if (moveCopiedBytes > 0)
                            {
                                FileSystemOptions options{};
                                options.sizeBytes                    = sizeof(FileSystemOptions);
                                options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                                const HRESULT hrProgress             = FileSystemProgress(_operation,
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
                                                           {},
                                                           [&]() noexcept -> HRESULT
                            {
                                // Cross-filesystem moves already materialized the full destination tree.
                                // Deleting the original source must therefore remove directory roots
                                // predictably even when the caller did not set RECURSIVE explicitly.
                                const FileSystemFlags deleteFlags = static_cast<FileSystemFlags>(itemFlags | FILESYSTEM_FLAG_RECURSIVE);
                                BridgeCallback callback(*this);
                                return _fileSystem->DeleteItem(sourceText.c_str(), deleteFlags, nullptr, &callback, nullptr);
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
                            options.sizeBytes                    = sizeof(FileSystemOptions);
                            options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
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
                    clearConflictPrompt();
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
                    auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
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
                    auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                    LogDiagnostic(DiagnosticSeverity::Error,
                                  itemHr,
                                  L"delete.recycleBin.item",
                                  L"Recycle Bin delete failed for item.",
                                  diagnosticSource,
                                  diagnosticDestination);
                }

                const size_t bucketIndex = static_cast<size_t>(bucket);

                std::optional<ConflictAction> cached = getCachedDecision(bucket);
                if (cached.has_value() && isModifierConflictAction(cached.value()) && bucketIndex < cachedModifierAttempts.size() &&
                    cachedModifierAttempts[bucketIndex] >= kMaxCachedModifierAttemptsPerBucket)
                {
                    clearCachedDecision(bucket);
                    cached.reset();
                }
                ConflictAction action = cached.value_or(ConflictAction::None);

                if (action == ConflictAction::None)
                {
                    const bool canRetryBucket = bucket != ConflictBucket::UnsupportedReparse;
                    const bool allowRetry     = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] == 0u;
                    const bool retryFailed    = canRetryBucket && bucketIndex < retryCounts.size() && retryCounts[bucketIndex] != 0u;

                    {
                        std::unique_lock lock(_conflictMutex);
                        setConflictPromptLocked(&cookie, bucket, itemHr, sourceText, destinationItemText, allowRetry, retryFailed);
                    }
                    const auto decision = waitForConflictDecision(&cookie, bucket);
                    action              = decision.first;
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

                if (action == ConflictAction::SkipAll)
                {
                    auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
                    LogDiagnostic(DiagnosticSeverity::Warning,
                                  itemHr,
                                  L"item.conflict.skipAll",
                                  L"Conflict action Skip all similar conflicts selected.",
                                  diagnosticSource,
                                  diagnosticDestination);
                    setCachedDecision(bucket, ConflictAction::Skip);
                    itemSkipped     = true;
                    hadSkippedItems = true;
                    break;
                }

                if (action == ConflictAction::Skip)
                {
                    auto [diagnosticSource, diagnosticDestination] = getMostSpecificPathsForDiagnostics(&cookie, sourceText, destinationItemText);
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
                    clearConflictPrompt();
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

        clearConflictPrompt();

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
        options.sizeBytes                    = sizeof(FileSystemOptions);
        options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        const HRESULT operationHr            = _fileSystem->CopyItems(pathArray, count, destinationFolder.c_str(), _flags, &options, this, nullptr);
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
        options.sizeBytes                    = sizeof(FileSystemOptions);
        options.bandwidthLimitBytesPerSecond = _desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        const HRESULT operationHr            = _fileSystem->MoveItems(pathArray, count, destinationFolder.c_str(), _flags, &options, this, nullptr);
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
