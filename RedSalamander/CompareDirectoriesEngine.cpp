#include "Framework.h"

#include "CompareDirectoriesEngine.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <cwctype>
#include <deque>
#include <filesystem>
#include <limits>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <utility>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "DirectoryInfoCache.h"
#include "Helpers.h"
#include "NavigationLocation.h"

namespace
{
struct SideEntry
{
    std::wstring name;
    bool isDirectory      = false;
    uint64_t sizeBytes    = 0;
    int64_t lastWriteTime = 0;
    DWORD fileAttributes  = 0;
};

[[nodiscard]] size_t CombineHash(size_t seed, size_t value) noexcept
{
    // 64-bit mix (boost-like).
    constexpr size_t kMagic = 0x9E3779B97F4A7C15ull;
    seed ^= value + kMagic + (seed << 6) + (seed >> 2);
    return seed;
}

[[nodiscard]] bool IsMissingPathError(HRESULT hr) noexcept
{
    if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_DIRECTORY))
    {
        return true;
    }

    if (hr == HRESULT_FROM_WIN32(ERROR_BAD_NETPATH) || hr == HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME) || hr == HRESULT_FROM_WIN32(ERROR_INVALID_DRIVE))
    {
        return true;
    }

    return false;
}

[[nodiscard]] bool AnyChildDifferent(const CompareDirectoriesFolderDecision& decision) noexcept
{
    for (const auto& kv : decision.items)
    {
        if (kv.second.isDifferent)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool AnyChildPending(const CompareDirectoriesFolderDecision& decision) noexcept
{
    if (decision.pendingContentCompareCount != 0u)
    {
        return true;
    }

    for (const auto& kv : decision.items)
    {
        const uint32_t mask = kv.second.differenceMask;
        if (HasFlag(mask, CompareDirectoriesDiffBit::ContentPending) || HasFlag(mask, CompareDirectoriesDiffBit::SubdirPending))
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] std::wstring_view SyncManifestBlockerReasonDetail(CompareSyncManifestBlockerReason reason) noexcept
{
    switch (reason)
    {
        case CompareSyncManifestBlockerReason::None: return L"none";
        case CompareSyncManifestBlockerReason::MissingDecision: return L"missing-decision";
        case CompareSyncManifestBlockerReason::ContentPending: return L"content-pending";
        case CompareSyncManifestBlockerReason::SubdirPending: return L"subdir-pending";
        case CompareSyncManifestBlockerReason::FailedDecision: return L"failed-decision";
        case CompareSyncManifestBlockerReason::Unsupported: return L"unsupported";
        default: break;
    }
    return L"unknown";
}

[[nodiscard]] std::wstring_view SyncManifestStatusDetail(CompareSyncManifestStatus status, CompareSyncManifestBlockerReason reason) noexcept
{
    switch (status)
    {
        case CompareSyncManifestStatus::Ready: return L"ready";
        case CompareSyncManifestStatus::Empty: return L"empty";
        case CompareSyncManifestStatus::NotReady:
            switch (reason)
            {
                case CompareSyncManifestBlockerReason::None: return L"not-ready";
                case CompareSyncManifestBlockerReason::ContentPending: return L"not-ready/content-pending";
                case CompareSyncManifestBlockerReason::SubdirPending: return L"not-ready/subdir-pending";
                case CompareSyncManifestBlockerReason::MissingDecision: return L"not-ready/missing-decision";
                case CompareSyncManifestBlockerReason::FailedDecision: return L"not-ready/failed-decision";
                case CompareSyncManifestBlockerReason::Unsupported: return L"not-ready/unsupported";
                default: return L"not-ready";
            }
        case CompareSyncManifestStatus::Failed: return L"failed";
        case CompareSyncManifestStatus::Unsupported: return L"unsupported";
        default: break;
    }
    return L"unknown";
}

[[nodiscard]] bool IsReparsePairEntry(const CompareDirectoriesItemDecision& item) noexcept
{
    return (item.leftFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0 || (item.rightFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
}

[[nodiscard]] bool IsCacheableDecision(const std::shared_ptr<const CompareDirectoriesFolderDecision>& decision) noexcept
{
    return decision && SUCCEEDED(decision->hr);
}

constexpr uint32_t kMaxPendingUpdateRetryAttempts = 3u;

void ApplyCriteriaDiffAndSelection(CompareDirectoriesItemDecision& item,
                                   const Common::Settings::CompareDirectoriesSettings& settings,
                                   bool canCompareContent,
                                   bool contentDifferent) noexcept
{
    const bool sizeDifferent  = item.leftSizeBytes != item.rightSizeBytes;
    const bool timeDifferent  = item.leftLastWriteTime != item.rightLastWriteTime;
    const bool attrsDifferent = item.leftFileAttributes != item.rightFileAttributes;

    item.differenceMask = 0;
    item.isDifferent    = false;
    item.selectLeft     = false;
    item.selectRight    = false;

    if (settings.compareSize && sizeDifferent)
    {
        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::Size);
    }
    if (settings.compareDateTime && timeDifferent)
    {
        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::DateTime);
    }
    if (settings.compareAttributes && attrsDifferent)
    {
        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::Attributes);
    }
    if (canCompareContent && contentDifferent)
    {
        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::Content);
    }

    const bool anyCriteriaDifferent = (settings.compareSize && sizeDifferent) || (settings.compareDateTime && timeDifferent) ||
                                      (settings.compareAttributes && attrsDifferent) || (canCompareContent && contentDifferent);
    if (! anyCriteriaDifferent)
    {
        return;
    }

    item.isDifferent = true;

    if (settings.compareSize && sizeDifferent)
    {
        if (item.leftSizeBytes > item.rightSizeBytes)
        {
            item.selectLeft = true;
        }
        else
        {
            item.selectRight = true;
        }
    }

    if (settings.compareDateTime && timeDifferent)
    {
        if (item.leftLastWriteTime > item.rightLastWriteTime)
        {
            item.selectLeft = true;
        }
        else
        {
            item.selectRight = true;
        }
    }

    if (settings.compareAttributes && attrsDifferent)
    {
        item.selectLeft  = true;
        item.selectRight = true;
    }

    if (canCompareContent && contentDifferent)
    {
        item.selectLeft  = true;
        item.selectRight = true;
    }
}

template <typename T> void QueueCompareCleanup(std::unique_ptr<T> cleanup, std::wstring_view label) noexcept
{
    if (! cleanup)
    {
        return;
    }

    if (TrySubmitUniqueToThreadpool(cleanup))
    {
        return;
    }

    const DWORD lastError = GetLastError();
    static_cast<void>(cleanup.release());
    Debug::Warning(L"CompareDirectories: {} scheduling failed (gle=0x{:08X}); abandoning deferred cleanup to keep teardown non-blocking.",
                   label,
                   static_cast<unsigned long>(lastError));
}

[[nodiscard]] std::wstring_view TrimWhitespace(std::wstring_view text) noexcept
{
    return StringUtils::TrimWhitespace(text);
}

constexpr size_t kMaxIgnorePatternCount  = 32;
constexpr size_t kMaxIgnorePatternLength = 128;

[[nodiscard]] std::vector<std::wstring> SplitPatternsCapped(std::wstring_view patterns) noexcept
{
    std::vector<std::wstring> result;
    patterns = TrimWhitespace(patterns);
    if (patterns.empty())
    {
        return result;
    }

    result.reserve(kMaxIgnorePatternCount);

    size_t start = 0;
    while (start <= patterns.size() && result.size() < kMaxIgnorePatternCount)
    {
        const size_t sep = patterns.find(L';', start);
        const size_t end = sep == std::wstring_view::npos ? patterns.size() : sep;

        std::wstring_view token = patterns.substr(start, end - start);
        token                   = TrimWhitespace(token);
        if (! token.empty() && token.size() <= kMaxIgnorePatternLength)
        {
            result.emplace_back(token);
        }

        if (sep == std::wstring_view::npos)
        {
            break;
        }
        start = sep + 1;
    }

    return result;
}

[[nodiscard]] wchar_t LowerInvariant(wchar_t ch) noexcept
{
    wchar_t buf[2] = {ch, L'\0'};
    ::CharLowerW(buf);
    return buf[0];
}

[[nodiscard]] bool HasWildcardChars(std::wstring_view pattern) noexcept
{
    return pattern.find_first_of(L"*?") != std::wstring_view::npos;
}

[[nodiscard]] uint32_t WildcardBudgetFor(std::wstring_view text, std::wstring_view pattern) noexcept
{
    constexpr uint32_t kBase    = 64u;
    constexpr uint32_t kPerChar = 8u;
    constexpr uint32_t kMax     = 4096u;

    const uint64_t span = static_cast<uint64_t>(text.size()) + static_cast<uint64_t>(pattern.size());
    const uint64_t raw  = static_cast<uint64_t>(kBase) + static_cast<uint64_t>(kPerChar) * span;
    return static_cast<uint32_t>((std::min<uint64_t>)(raw, kMax));
}

[[nodiscard]] bool WildcardMatchNoCaseBudgeted(std::wstring_view text, std::wstring_view pattern, uint32_t budget) noexcept
{
    // Glob match with '*' and '?', case-insensitive.
    size_t ti = 0;
    size_t pi = 0;

    size_t star  = std::wstring_view::npos;
    size_t match = 0;

    while (ti < text.size())
    {
        if (budget == 0u)
        {
            // Harden against pathological patterns: budget exceeded => treat as no match.
            return false;
        }
        --budget;

        if (pi < pattern.size())
        {
            const wchar_t pch = pattern[pi];
            if (pch == L'?')
            {
                ++ti;
                ++pi;
                continue;
            }
            if (pch == L'*')
            {
                star  = pi++;
                match = ti;
                continue;
            }

            if (LowerInvariant(text[ti]) == LowerInvariant(pch))
            {
                ++ti;
                ++pi;
                continue;
            }
        }

        if (star != std::wstring_view::npos)
        {
            pi = star + 1;
            ++match;
            ti = match;
            continue;
        }

        return false;
    }

    while (pi < pattern.size() && pattern[pi] == L'*')
    {
        if (budget == 0u)
        {
            return false;
        }
        --budget;
        ++pi;
    }

    return pi == pattern.size();
}

[[nodiscard]] bool MatchesAnyPattern(std::wstring_view name, const std::vector<std::wstring>& patterns) noexcept
{
    for (const std::wstring& pat : patterns)
    {
        if (pat.empty())
        {
            continue;
        }

        const std::wstring_view patView(pat);
        if (! HasWildcardChars(patView))
        {
            if (name.size() <= static_cast<size_t>(std::numeric_limits<int>::max()) && patView.size() <= static_cast<size_t>(std::numeric_limits<int>::max()))
            {
                if (wil::compare_string_ordinal(name, patView, true) == wistd::weak_ordering::equivalent)
                {
                    return true;
                }
            }
            else if (name == patView)
            {
                return true;
            }

            continue;
        }

        const uint32_t budget = WildcardBudgetFor(name, patView);
        if (WildcardMatchNoCaseBudgeted(name, patView, budget))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring_view NormalizeEntryNameForCompare(std::wstring_view name) noexcept;

[[nodiscard]] bool ShouldIgnoreEntry(std::wstring_view name,
                                     bool isDirectory,
                                     const Common::Settings::CompareDirectoriesSettings& settings,
                                     const std::vector<std::wstring>& ignoreFilePatterns,
                                     const std::vector<std::wstring>& ignoreDirectoryPatterns) noexcept
{
    if (name.empty())
    {
        return true;
    }

    if (name == L"." || name == L"..")
    {
        return true;
    }

    if (isDirectory)
    {
        return settings.ignoreDirectories && MatchesAnyPattern(name, ignoreDirectoryPatterns);
    }

    return settings.ignoreFiles && MatchesAnyPattern(name, ignoreFilePatterns);
}

[[nodiscard]] bool ShouldIgnoreRelativeFolder(const std::filesystem::path& relativeFolder,
                                              const Common::Settings::CompareDirectoriesSettings& settings,
                                              const std::vector<std::wstring>& ignoreDirectoryPatterns) noexcept
{
    if (! settings.ignoreDirectories || relativeFolder.empty())
    {
        return false;
    }

    const std::filesystem::path normalized = relativeFolder.lexically_normal();
    for (const auto& component : normalized)
    {
        const std::wstring name = component.native();
        if (name.empty() || name == L"." || name == L"..")
        {
            continue;
        }

        if (MatchesAnyPattern(NormalizeEntryNameForCompare(name), ignoreDirectoryPatterns))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool CacheKeyEqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    return OrdinalString::EqualsNoCase(left, right);
}

[[nodiscard]] bool CacheKeyIsSameOrDescendant(std::wstring_view key, std::wstring_view prefix) noexcept
{
    if (prefix == L".")
    {
        return true;
    }

    return CacheKeyEqualsNoCase(key, prefix) || (key.size() > prefix.size() && OrdinalString::StartsWithNoCase(key, prefix) && key[prefix.size()] == L'/');
}

[[nodiscard]] std::filesystem::path ResolveAbsoluteFromRoot(const std::filesystem::path& root, const std::filesystem::path& relativeFolder)
{
    if (NavigationLocation::LooksLikeWindowsAbsolutePath(root.native()))
    {
        if (relativeFolder.empty())
        {
            return root;
        }
        return (root / relativeFolder).lexically_normal();
    }

    const std::wstring rootNorm = NavigationLocation::NormalizePluginPathText(root.native(),
                                                                              NavigationLocation::EmptyPathPolicy::Root,
                                                                              NavigationLocation::LeadingSlashPolicy::Ensure,
                                                                              NavigationLocation::TrailingSlashPolicy::Trim);
    if (relativeFolder.empty())
    {
        return std::filesystem::path(rootNorm);
    }

    std::wstring relativeText = relativeFolder.generic_wstring();
    while (! relativeText.empty() && (relativeText.front() == L'/' || relativeText.front() == L'\\'))
    {
        relativeText.erase(relativeText.begin());
    }

    if (relativeText.empty())
    {
        return std::filesystem::path(rootNorm);
    }

    std::wstring joined = rootNorm;
    if (! joined.empty() && joined.back() != L'/')
    {
        joined.push_back(L'/');
    }
    joined += relativeText;

    return std::filesystem::path(NavigationLocation::NormalizePluginPathText(
        joined, NavigationLocation::EmptyPathPolicy::Root, NavigationLocation::LeadingSlashPolicy::Ensure, NavigationLocation::TrailingSlashPolicy::Preserve));
}
} // namespace

CompareDirectoriesSession::CompareDirectoriesSession(wil::com_ptr<IFileSystem> leftFileSystem,
                                                     wil::com_ptr<IFileSystem> rightFileSystem,
                                                     std::filesystem::path leftRoot,
                                                     std::filesystem::path rightRoot,
                                                     Common::Settings::CompareDirectoriesSettings settings)
    : _leftFileSystem(std::move(leftFileSystem)),
      _rightFileSystem(std::move(rightFileSystem)),
      _leftRoot(std::move(leftRoot)),
      _rightRoot(std::move(rightRoot)),
      _settings(std::move(settings))
{
    _ignoreFilePatterns      = std::make_shared<const std::vector<std::wstring>>(SplitPatternsCapped(_settings.ignoreFilesPatterns));
    _ignoreDirectoryPatterns = std::make_shared<const std::vector<std::wstring>>(SplitPatternsCapped(_settings.ignoreDirectoriesPatterns));

    if (_leftFileSystem)
    {
        wil::com_ptr<IInformations> infos;
        const HRESULT qiHr = _leftFileSystem->QueryInterface(__uuidof(IInformations), infos.put_void());
        if (SUCCEEDED(qiHr) && infos)
        {
            _leftInformations = std::move(infos);
        }

        wil::com_ptr<IFileSystemIO> io;
        const HRESULT qiIo = _leftFileSystem->QueryInterface(__uuidof(IFileSystemIO), io.put_void());
        if (SUCCEEDED(qiIo) && io)
        {
            _leftFileSystemIo = std::move(io);
        }
    }

    if (_rightFileSystem)
    {
        wil::com_ptr<IInformations> infos;
        const HRESULT qiHr = _rightFileSystem->QueryInterface(__uuidof(IInformations), infos.put_void());
        if (SUCCEEDED(qiHr) && infos)
        {
            _rightInformations = std::move(infos);
        }

        wil::com_ptr<IFileSystemIO> io;
        const HRESULT qiIo = _rightFileSystem->QueryInterface(__uuidof(IFileSystemIO), io.put_void());
        if (SUCCEEDED(qiIo) && io)
        {
            _rightFileSystemIo = std::move(io);
        }
    }
}

CompareDirectoriesSession::~CompareDirectoriesSession()
{
    for (auto& worker : _scanWorkers)
    {
        worker.request_stop();
    }

    for (auto& worker : _contentCompareWorkers)
    {
        worker.request_stop();
    }

    {
        std::lock_guard guard(_mutex);

        _scanQueueHigh.clear();
        _scanQueueLow.clear();
        _scanScheduledKeys.clear();
        _scanHighQueuedKeys.clear();
        _scanInFlightKeys.clear();
        _pendingSubdirUpdates.clear();
        _pendingSubdirAggregateUpdates.clear();
        _pendingSubdirAggregateRetryAttempts.clear();
        _scanActiveScans.store(0u, std::memory_order_release);
        _scanFoldersScanned.store(0u, std::memory_order_release);
        _scanEntriesScanned.store(0u, std::memory_order_release);

        _contentCompareQueueHigh.clear();
        _contentCompareQueueLow.clear();
        _contentCompareInFlight.clear();
        _pendingContentCompareUpdates.clear();
        _pendingContentCompareRetryAttempts.clear();
    }

    _scanCv.notify_all();
    _contentCompareCv.notify_all();
    _contentCompareQueueNotFullCv.notify_all();

    const auto joinWorkers = [](std::vector<std::jthread>& workers) noexcept
    {
        for (std::jthread& worker : workers)
        {
            if (worker.joinable())
            {
                worker.join();
            }
        }
    };

    // Do not rely on member destruction order here: scan workers are declared before
    // content-compare state, so implicit jthread cleanup would join them too late.
    joinWorkers(_scanWorkers);
    joinWorkers(_contentCompareWorkers);
}

size_t CompareDirectoriesSession::ContentCompareKeyHash::operator()(const ContentCompareKey& key) const noexcept
{
    size_t hash = std::hash<std::wstring_view>{}(key.relativeFileKey);
    hash        = CombineHash(hash, std::hash<uint64_t>{}(key.leftSizeBytes));
    hash        = CombineHash(hash, std::hash<uint64_t>{}(key.rightSizeBytes));
    hash        = CombineHash(hash, std::hash<int64_t>{}(key.leftLastWriteTime));
    hash        = CombineHash(hash, std::hash<int64_t>{}(key.rightLastWriteTime));
    return hash;
}

bool CompareDirectoriesSession::ContentCompareKeyEq::operator()(const ContentCompareKey& a, const ContentCompareKey& b) const noexcept
{
    return a.leftSizeBytes == b.leftSizeBytes && a.rightSizeBytes == b.rightSizeBytes && a.leftLastWriteTime == b.leftLastWriteTime &&
           a.rightLastWriteTime == b.rightLastWriteTime && a.relativeFileKey == b.relativeFileKey;
}

void CompareDirectoriesSession::SetRoots(std::filesystem::path leftRoot, std::filesystem::path rightRoot)
{
    auto cleanup = std::make_unique<ResetCleanup>();
    {
        std::lock_guard guard(_mutex);
        _leftRoot  = std::move(leftRoot);
        _rightRoot = std::move(rightRoot);
        _version.fetch_add(1u, std::memory_order_relaxed);
        ++_uiVersion;
        ResetCompareStateLocked(*cleanup);
    }
    ScheduleResetCleanup(std::move(cleanup));
    NotifyContentProgress(0u, {}, {}, 0, 0);
}

void CompareDirectoriesSession::SetSettings(Common::Settings::CompareDirectoriesSettings settings)
{
    const auto ignoreFilePatterns      = std::make_shared<const std::vector<std::wstring>>(SplitPatternsCapped(settings.ignoreFilesPatterns));
    const auto ignoreDirectoryPatterns = std::make_shared<const std::vector<std::wstring>>(SplitPatternsCapped(settings.ignoreDirectoriesPatterns));

    bool clearedContentCompare = false;
    std::unique_ptr<ResetCleanup> cleanup;
    {
        std::lock_guard guard(_mutex);

        // Note: keepIdenticalItems is included here because cached decisions are pruned when it is off
        // (keeps memory bounded on large trees). Toggling it requires invalidating decisions.
        const bool comparisonChanged =
            _settings.compareSize != settings.compareSize || _settings.compareDateTime != settings.compareDateTime ||
            _settings.compareAttributes != settings.compareAttributes || _settings.compareContent != settings.compareContent ||
            _settings.compareSubdirectories != settings.compareSubdirectories ||
            _settings.compareSubdirectoryAttributes != settings.compareSubdirectoryAttributes ||
            _settings.selectSubdirsOnlyInOnePane != settings.selectSubdirsOnlyInOnePane || _settings.keepIdenticalItems != settings.keepIdenticalItems ||
            _settings.ignoreFiles != settings.ignoreFiles || _settings.ignoreFilesPatterns != settings.ignoreFilesPatterns ||
            _settings.ignoreDirectories != settings.ignoreDirectories || _settings.ignoreDirectoriesPatterns != settings.ignoreDirectoriesPatterns;

        _settings                = std::move(settings);
        _ignoreFilePatterns      = ignoreFilePatterns;
        _ignoreDirectoryPatterns = ignoreDirectoryPatterns;
        if (comparisonChanged)
        {
            cleanup = std::make_unique<ResetCleanup>();
            _version.fetch_add(1u, std::memory_order_relaxed);
            ++_uiVersion;
            ResetCompareStateLocked(*cleanup, false);
            clearedContentCompare = true;
        }
    }

    if (clearedContentCompare)
    {
        ScheduleResetCleanup(std::move(cleanup));
        NotifyContentProgress(0u, {}, {}, 0, 0);
    }
}

void CompareDirectoriesSession::SetCompareEnabled(bool enabled) noexcept
{
    _compareEnabled.store(enabled, std::memory_order_release);
}

bool CompareDirectoriesSession::IsCompareEnabled() const noexcept
{
    return _compareEnabled.load(std::memory_order_acquire);
}

void CompareDirectoriesSession::SetBackgroundWorkEnabled(bool enabled) noexcept
{
    if (enabled)
    {
        _backgroundWorkEnabled.store(true, std::memory_order_release);
        return;
    }

    _backgroundWorkEnabled.store(false, std::memory_order_release);
    static_cast<void>(_backgroundWorkCancelToken.fetch_add(1u, std::memory_order_acq_rel));

    auto cleanup = std::make_unique<ResetCleanup>();
    {
        std::lock_guard guard(_mutex);
        cleanup->scanQueueHigh.swap(_scanQueueHigh);
        cleanup->scanQueueLow.swap(_scanQueueLow);
        cleanup->scanScheduledKeys.swap(_scanScheduledKeys);
        cleanup->scanHighQueuedKeys.swap(_scanHighQueuedKeys);
        cleanup->scanInFlightKeys.swap(_scanInFlightKeys);
        cleanup->pendingSubdirUpdates.swap(_pendingSubdirUpdates);
        cleanup->pendingSubdirAggregateUpdates.swap(_pendingSubdirAggregateUpdates);
        cleanup->pendingSubdirAggregateRetryAttempts.swap(_pendingSubdirAggregateRetryAttempts);

        _scanActiveScans.store(0u, std::memory_order_release);
        _scanFoldersScanned.store(0u, std::memory_order_release);
        _scanEntriesScanned.store(0u, std::memory_order_release);
        _scanLastNotifyTickMs.store(0u, std::memory_order_release);
        _scanCv.notify_all();

        cleanup->contentCompareInFlight.swap(_contentCompareInFlight);
        cleanup->contentCompareQueueHigh.swap(_contentCompareQueueHigh);
        cleanup->contentCompareQueueLow.swap(_contentCompareQueueLow);
        cleanup->pendingContentCompareUpdates.swap(_pendingContentCompareUpdates);
        cleanup->pendingContentCompareRetryAttempts.swap(_pendingContentCompareRetryAttempts);
        _contentCompareCache.clear();

        _contentComparePendingCompares.store(0u, std::memory_order_release);
        _contentCompareTotalCompares.store(0u, std::memory_order_release);
        _contentCompareCompletedCompares.store(0u, std::memory_order_release);
        _contentCompareTotalBytes.store(0u, std::memory_order_release);
        _contentCompareCompletedBytes.store(0u, std::memory_order_release);

        _contentCompareCv.notify_all();
        _contentCompareQueueNotFullCv.notify_all();
    }

    ScheduleResetCleanup(std::move(cleanup));
    NotifyScanProgress({}, {}, true);
    NotifyContentProgress(0u, {}, {}, 0, 0);
}

bool CompareDirectoriesSession::IsBackgroundWorkEnabled() const noexcept
{
    return _backgroundWorkEnabled.load(std::memory_order_acquire);
}

void CompareDirectoriesSession::Invalidate() noexcept
{
    auto cleanup = std::make_unique<ResetCleanup>();
    {
        std::lock_guard guard(_mutex);
        _version.fetch_add(1u, std::memory_order_relaxed);
        ++_uiVersion;
        ResetCompareStateLocked(*cleanup);
    }
    ScheduleResetCleanup(std::move(cleanup));
    NotifyContentProgress(0u, {}, {}, 0, 0);
}

void CompareDirectoriesSession::InvalidateForRelativePathLocked(const std::filesystem::path& relativePath, bool includeSubtree) noexcept
{
    std::filesystem::path normalizedPath = relativePath.lexically_normal();
    if (normalizedPath == L".")
    {
        normalizedPath.clear();
    }
    EvictContentCompareCacheForRelativePathLocked(normalizedPath, includeSubtree);

    std::filesystem::path folder = normalizedPath;

    if (! includeSubtree)
    {
        folder = folder.parent_path();
    }

    if (! folder.empty())
    {
        folder = folder.lexically_normal();
    }

    const std::wstring subtreePrefix = MakeCacheKey(folder);
    const bool subtreeInvalidatesAll = includeSubtree && folder.empty();
    std::vector<std::wstring> ancestorKeys;

    if (includeSubtree)
    {
        if (folder.empty())
        {
            _cache.clear();
            _decisionCacheLru.clear();
            _decisionCacheMeta.clear();
            _decisionCacheEstimatedBytes = 0;
            _pendingContentCompareUpdates.clear();
            _pendingContentCompareRetryAttempts.clear();
            _pendingSubdirAggregateUpdates.clear();
            _pendingSubdirAggregateRetryAttempts.clear();
        }
        else
        {
            const std::wstring prefix = MakeCacheKey(folder);
            for (auto it = _cache.lower_bound(prefix); it != _cache.end();)
            {
                const std::wstring_view key = it->first;
                if (! OrdinalString::StartsWithNoCase(key, prefix))
                {
                    break;
                }

                if (CacheKeyIsSameOrDescendant(key, prefix))
                {
                    _pendingContentCompareUpdates.erase(it->first);
                    _pendingContentCompareRetryAttempts.erase(it->first);
                    TrackDecisionCacheEraseLocked(key);
                    it = _cache.erase(it);
                }
                else
                {
                    ++it;
                }
            }
        }
    }

    for (std::filesystem::path current = folder;; current = current.parent_path())
    {
        const std::wstring key = MakeCacheKey(current);
        ancestorKeys.emplace_back(key);
        TrackDecisionCacheEraseLocked(key);
        _cache.erase(key);
        _pendingContentCompareUpdates.erase(key);
        _pendingContentCompareRetryAttempts.erase(key);
        _pendingSubdirUpdates.erase(key);
        _pendingSubdirAggregateUpdates.erase(key);
        _pendingSubdirAggregateRetryAttempts.erase(key);
        if (current.empty())
        {
            break;
        }
    }

    const auto decisionKeyAffected = [&](std::wstring_view key) noexcept
    {
        if (subtreeInvalidatesAll)
        {
            return true;
        }
        if (includeSubtree && CacheKeyIsSameOrDescendant(key, subtreePrefix))
        {
            return true;
        }
        return std::any_of(
            ancestorKeys.begin(), ancestorKeys.end(), [&](const std::wstring& ancestorKey) noexcept { return CacheKeyEqualsNoCase(key, ancestorKey); });
    };

    const auto scanKeyAffected = [&](std::wstring_view key) noexcept
    {
        if (subtreeInvalidatesAll)
        {
            return true;
        }
        if (includeSubtree)
        {
            return CacheKeyIsSameOrDescendant(key, subtreePrefix);
        }
        return CacheKeyEqualsNoCase(key, subtreePrefix);
    };

    for (auto it = _pendingSubdirUpdates.begin(); it != _pendingSubdirUpdates.end();)
    {
        if (decisionKeyAffected(*it))
        {
            it = _pendingSubdirUpdates.erase(it);
        }
        else
        {
            ++it;
        }
    }

    for (auto it = _pendingSubdirAggregateUpdates.begin(); it != _pendingSubdirAggregateUpdates.end();)
    {
        if (decisionKeyAffected(it->first))
        {
            _pendingSubdirAggregateRetryAttempts.erase(it->first);
            it = _pendingSubdirAggregateUpdates.erase(it);
        }
        else
        {
            ++it;
        }
    }

    const auto eraseScanQueue = [&](std::deque<FolderScanJob>& queue) noexcept
    {
        for (auto it = queue.begin(); it != queue.end();)
        {
            if (scanKeyAffected(it->key))
            {
                it = queue.erase(it);
            }
            else
            {
                ++it;
            }
        }
    };

    eraseScanQueue(_scanQueueHigh);
    eraseScanQueue(_scanQueueLow);

    for (auto it = _scanHighQueuedKeys.begin(); it != _scanHighQueuedKeys.end();)
    {
        if (scanKeyAffected(*it))
        {
            it = _scanHighQueuedKeys.erase(it);
        }
        else
        {
            ++it;
        }
    }

    size_t removedScheduledScans = 0;
    for (auto it = _scanScheduledKeys.begin(); it != _scanScheduledKeys.end();)
    {
        if (scanKeyAffected(*it))
        {
            it = _scanScheduledKeys.erase(it);
            ++removedScheduledScans;
        }
        else
        {
            ++it;
        }
    }

    for (auto it = _scanInFlightKeys.begin(); it != _scanInFlightKeys.end();)
    {
        if (scanKeyAffected(it->first))
        {
            it = _scanInFlightKeys.erase(it);
        }
        else
        {
            ++it;
        }
    }

    if (removedScheduledScans != 0u)
    {
        const uint32_t activeScans = _scanActiveScans.load(std::memory_order_relaxed);
        if (removedScheduledScans > activeScans)
        {
            Debug::Warning(
                L"CompareDirectories: scan active counter drift during invalidation (active={}, removed scheduled={}).", activeScans, removedScheduledScans);
        }
        const uint32_t removed = static_cast<uint32_t>(std::min<size_t>(removedScheduledScans, activeScans));
        _scanActiveScans.store(activeScans - removed, std::memory_order_release);
    }

    const std::wstring contentPrefix = MakeCacheKey(normalizedPath);
    const auto contentKeyAffected    = [&](std::wstring_view key) noexcept
    {
        if (normalizedPath.empty())
        {
            return includeSubtree;
        }
        return includeSubtree ? CacheKeyIsSameOrDescendant(key, contentPrefix) : CacheKeyEqualsNoCase(key, contentPrefix);
    };

    const auto eraseContentQueue = [&](std::deque<ContentCompareJob>& queue) noexcept
    {
        for (auto it = queue.begin(); it != queue.end();)
        {
            if (contentKeyAffected(it->key.relativeFileKey))
            {
                it = queue.erase(it);
            }
            else
            {
                ++it;
            }
        }
    };

    eraseContentQueue(_contentCompareQueueHigh);
    eraseContentQueue(_contentCompareQueueLow);

    size_t removedContentInFlight = 0;
    for (auto it = _contentCompareInFlight.begin(); it != _contentCompareInFlight.end();)
    {
        if (contentKeyAffected(it->first.relativeFileKey))
        {
            it = _contentCompareInFlight.erase(it);
            ++removedContentInFlight;
        }
        else
        {
            ++it;
        }
    }

    if (removedContentInFlight != 0u)
    {
        const uint64_t pending = _contentComparePendingCompares.load(std::memory_order_relaxed);
        const uint64_t removed = std::min<uint64_t>(static_cast<uint64_t>(removedContentInFlight), pending);
        _contentComparePendingCompares.store(pending - removed, std::memory_order_release);
    }

    _scanCv.notify_all();
    _contentCompareCv.notify_all();
    _contentCompareQueueNotFullCv.notify_all();

    ++_uiVersion;
}

void CompareDirectoriesSession::InvalidateForAbsolutePath(const std::filesystem::path& absolutePath, bool includeSubtree) noexcept
{
    if (absolutePath.empty())
    {
        return;
    }

    const std::optional<std::filesystem::path> relLeft  = TryMakeRelative(ComparePane::Left, absolutePath);
    const std::optional<std::filesystem::path> relRight = TryMakeRelative(ComparePane::Right, absolutePath);

    if (! relLeft.has_value() && ! relRight.has_value())
    {
        return;
    }

    {
        std::lock_guard guard(_mutex);
        if (relLeft.has_value())
        {
            InvalidateForRelativePathLocked(relLeft.value(), includeSubtree);
        }
        if (relRight.has_value())
        {
            InvalidateForRelativePathLocked(relRight.value(), includeSubtree);
        }
    }

    NotifyScanProgress({}, {}, true);
    NotifyContentProgress(0u, {}, {}, 0, 0);
    NotifyDecisionUpdated(true);
}

void CompareDirectoriesSession::SetScanProgressCallback(ScanProgressCallback callback) noexcept
{
    std::shared_ptr<const ScanProgressCallback> stored;
    if (callback)
    {
        stored = std::make_shared<ScanProgressCallback>(std::move(callback));
    }

    _scanProgressCallback.store(std::move(stored), std::memory_order_release);
}

void CompareDirectoriesSession::SetContentProgressCallback(ContentProgressCallback callback) noexcept
{
    std::shared_ptr<const ContentProgressCallback> stored;
    if (callback)
    {
        stored = std::make_shared<ContentProgressCallback>(std::move(callback));
    }

    _contentProgressCallback.store(std::move(stored), std::memory_order_release);
}

void CompareDirectoriesSession::SetDecisionUpdatedCallback(DecisionUpdatedCallback callback) noexcept
{
    std::shared_ptr<const DecisionUpdatedCallback> stored;
    if (callback)
    {
        stored = std::make_shared<DecisionUpdatedCallback>(std::move(callback));
    }

    _decisionUpdatedCallback.store(std::move(stored), std::memory_order_release);
}

void CompareDirectoriesSession::FlushPendingContentCompareUpdates() noexcept
{
    static_cast<void>(FlushPendingContentCompareUpdatesBudgeted(std::numeric_limits<size_t>::max()));
}

bool CompareDirectoriesSession::FlushPendingContentCompareUpdatesBudgeted(size_t maxFoldersToApply) noexcept
{
    if (maxFoldersToApply == 0)
    {
        std::lock_guard guard(_mutex);
        return ! _pendingContentCompareUpdates.empty();
    }

    std::lock_guard guard(_mutex);

    // Apply in a bounded pass. Some folders may not have decisions cached yet (race with scan workers);
    // keep pending updates so they can be applied later once the decision exists.
    std::vector<std::wstring> keysToApply;
    keysToApply.reserve(std::min(maxFoldersToApply, _pendingContentCompareUpdates.size()));
    for (auto it = _pendingContentCompareUpdates.begin(); it != _pendingContentCompareUpdates.end() && keysToApply.size() < maxFoldersToApply; ++it)
    {
        keysToApply.emplace_back(it->first);
    }

    for (const std::wstring& key : keysToApply)
    {
        ApplyPendingContentCompareUpdatesLocked(key);
    }

    return ! _pendingContentCompareUpdates.empty();
}

void CompareDirectoriesSession::StartScan() noexcept
{
    if (! _backgroundWorkEnabled.load(std::memory_order_acquire))
    {
        return;
    }

    {
        std::lock_guard guard(_mutex);
        const uint64_t version     = _version.load(std::memory_order_relaxed);
        const uint64_t cancelToken = _backgroundWorkCancelToken.load(std::memory_order_relaxed);
        EnsureScanWorkersLocked();
        EnqueueScanLocked(std::filesystem::path{}, version, cancelToken, ScanPriority::Low);
    }

    NotifyScanProgress({}, {}, true);
}

void CompareDirectoriesSession::RequestScanForFolder(const std::filesystem::path& relativeFolder) noexcept
{
    if (! _backgroundWorkEnabled.load(std::memory_order_acquire))
    {
        return;
    }

    {
        std::lock_guard guard(_mutex);
        const uint64_t version     = _version.load(std::memory_order_relaxed);
        const uint64_t cancelToken = _backgroundWorkCancelToken.load(std::memory_order_relaxed);
        EnsureScanWorkersLocked();
        EnqueueScanLocked(relativeFolder, version, cancelToken, ScanPriority::High);
    }

    NotifyScanProgress(relativeFolder, {}, true);
}

std::shared_ptr<const CompareDirectoriesFolderDecision> CompareDirectoriesSession::TryGetCachedDecision(const std::filesystem::path& relativeFolder) noexcept
{
    std::lock_guard guard(_mutex);
    const uint64_t version       = _version.load(std::memory_order_relaxed);
    const std::wstring folderKey = MakeCacheKey(relativeFolder);

    const auto it = _cache.find(folderKey);
    if (it != _cache.end() && it->second && it->second->version == version)
    {
        if (IsCacheableDecision(it->second))
        {
            TouchDecisionCacheKeyLocked(folderKey);
            return it->second;
        }

        TrackDecisionCacheEraseLocked(folderKey);
        _cache.erase(it);
    }

    return {};
}

CompareSyncManifestStatus CompareDirectoriesSession::TryBuildSyncManifest(ComparePane sourcePane,
                                                                          const std::vector<std::filesystem::path>& selectedRelativePaths,
                                                                          CompareSyncManifest& out,
                                                                          CompareSyncManifestBlocker& blocker) noexcept
{
    out     = {};
    blocker = {};

    out.sourcePane = sourcePane;
    Debug::Perf::EmitCounter(L"compare.sync.manifest.build_count");
    Debug::Perf::Scope perf(L"compare.sync.manifest.build_us");

    auto pathDepth = [](const std::filesystem::path& path) noexcept
    {
        size_t depth = 0;
        for (const auto& part : path)
        {
            if (! part.empty() && part != L".")
            {
                ++depth;
            }
        }
        return depth;
    };

    const auto finish = [&](CompareSyncManifestStatus status) noexcept
    {
        perf.SetDetail(SyncManifestStatusDetail(status, blocker.reason));
        perf.SetValue0(out.items.size());
        perf.SetValue1(selectedRelativePaths.size());
        perf.SetHr(blocker.hr);

        if (status == CompareSyncManifestStatus::Ready)
        {
            Debug::Perf::EmitValue(L"compare.sync.manifest.items", out.items.size(), blocker.hr);
        }
        else if (status != CompareSyncManifestStatus::Empty)
        {
            Debug::Perf::Emit(
                L"compare.sync.manifest.blocker_count", SyncManifestBlockerReasonDetail(blocker.reason), 0, 1, pathDepth(blocker.relativePath), blocker.hr);
        }

        return status;
    };

    if (selectedRelativePaths.empty())
    {
        blocker.status = CompareSyncManifestStatus::Empty;
        return finish(CompareSyncManifestStatus::Empty);
    }

    const bool sourceIsLeft = sourcePane == ComparePane::Left;

    std::vector<std::filesystem::path> selected;
    selected.reserve(selectedRelativePaths.size());
    for (const auto& path : selectedRelativePaths)
    {
        std::filesystem::path normalized = path.lexically_normal();
        if (normalized == L".")
        {
            normalized.clear();
        }
        selected.push_back(std::move(normalized));
    }

    std::sort(selected.begin(),
              selected.end(),
              [&](const std::filesystem::path& left, const std::filesystem::path& right) noexcept
    {
        const size_t leftDepth  = pathDepth(left);
        const size_t rightDepth = pathDepth(right);
        if (leftDepth != rightDepth)
        {
            return leftDepth < rightDepth;
        }
        return WStringViewNoCaseLess{}(left.generic_wstring(), right.generic_wstring());
    });

    std::lock_guard guard(_mutex);
    const uint64_t version                                      = _version.load(std::memory_order_acquire);
    const Common::Settings::CompareDirectoriesSettings settings = _settings;
    const std::filesystem::path leftRootSnapshot                = _leftRoot;
    const std::filesystem::path rightRootSnapshot               = _rightRoot;
    const std::filesystem::path& sourceRoot                     = sourceIsLeft ? leftRootSnapshot : rightRootSnapshot;
    const std::filesystem::path& destinationRoot                = sourceIsLeft ? rightRootSnapshot : leftRootSnapshot;
    out.version                                                 = version;

    std::set<std::wstring, WStringViewNoCaseLess> emittedKeys;
    std::vector<std::filesystem::path> wholeSubtreeRelativePaths;

    const auto fail =
        [&](CompareSyncManifestStatus status, CompareSyncManifestBlockerReason reason, const std::filesystem::path& relativePath, HRESULT hr) noexcept
    {
        blocker.status       = status;
        blocker.reason       = reason;
        blocker.relativePath = relativePath;
        blocker.hr           = hr;
        out.items.clear();
        return finish(status);
    };

    const auto getDecisionLocked = [&](const std::filesystem::path& relativeFolder,
                                       std::shared_ptr<const CompareDirectoriesFolderDecision>& decision) noexcept -> CompareSyncManifestStatus
    {
        const std::wstring key = MakeCacheKey(relativeFolder);
        const auto it          = _cache.find(key);
        if (it == _cache.end() || ! it->second || it->second->version != version)
        {
            return fail(CompareSyncManifestStatus::NotReady, CompareSyncManifestBlockerReason::MissingDecision, relativeFolder, S_FALSE);
        }

        if (FAILED(it->second->hr))
        {
            return fail(CompareSyncManifestStatus::Failed, CompareSyncManifestBlockerReason::FailedDecision, relativeFolder, it->second->hr);
        }

        decision = it->second;
        return CompareSyncManifestStatus::Ready;
    };

    const auto isSameOrDescendant = [&](const std::filesystem::path& parent, const std::filesystem::path& candidate) noexcept
    { return CacheKeyIsSameOrDescendant(MakeCacheKey(candidate), MakeCacheKey(parent)); };

    const auto isCoveredByWholeSubtree = [&](const std::filesystem::path& relativePath) noexcept
    {
        for (const auto& whole : wholeSubtreeRelativePaths)
        {
            if (isSameOrDescendant(whole, relativePath))
            {
                return true;
            }
        }
        return false;
    };

    const auto sourceExists = [&](const CompareDirectoriesItemDecision& item) noexcept { return sourceIsLeft ? item.existsLeft : item.existsRight; };

    const auto destinationExists = [&](const CompareDirectoriesItemDecision& item) noexcept { return sourceIsLeft ? item.existsRight : item.existsLeft; };

    const auto sourceSelected = [&](const CompareDirectoriesItemDecision& item) noexcept { return sourceIsLeft ? item.selectLeft : item.selectRight; };

    const auto pendingReason = [](const CompareDirectoriesItemDecision& item) noexcept
    {
        if (HasFlag(item.differenceMask, CompareDirectoriesDiffBit::ContentPending))
        {
            return CompareSyncManifestBlockerReason::ContentPending;
        }
        if (HasFlag(item.differenceMask, CompareDirectoriesDiffBit::SubdirPending))
        {
            return CompareSyncManifestBlockerReason::SubdirPending;
        }
        return CompareSyncManifestBlockerReason::None;
    };

    const auto makeEmittedKey = [&](CompareSyncManifestItemKind kind, const std::filesystem::path& relativePath)
    {
        const std::wstring cacheKey = MakeCacheKey(relativePath);
        std::wstring key            = std::to_wstring(static_cast<unsigned int>(kind));
        key.push_back(L':');
        key.append(cacheKey);
        return key;
    };

    const auto emitItem = [&](CompareSyncManifestItemKind kind, const std::filesystem::path& relativePath, FileSystemFlags flags) noexcept
    {
        if (kind != CompareSyncManifestItemKind::DirectoryShell && isCoveredByWholeSubtree(relativePath))
        {
            return;
        }

        const std::wstring key = makeEmittedKey(kind, relativePath);
        if (! emittedKeys.insert(key).second)
        {
            return;
        }

        CompareSyncManifestItem item{};
        item.kind                    = kind;
        item.relativePath            = relativePath.lexically_normal();
        item.sourceAbsolutePath      = ResolveAbsoluteFromRoot(sourceRoot, item.relativePath);
        item.destinationAbsolutePath = ResolveAbsoluteFromRoot(destinationRoot, item.relativePath);
        item.flags                   = flags;
        out.items.push_back(std::move(item));

        if (kind == CompareSyncManifestItemKind::WholeSubtree)
        {
            wholeSubtreeRelativePaths.push_back(relativePath.lexically_normal());
        }
    };

    const auto shouldTransferFile = [&](const CompareDirectoriesItemDecision& item) noexcept
    {
        if (! sourceExists(item))
        {
            return false;
        }

        if (! destinationExists(item))
        {
            return true;
        }

        return item.isDifferent && sourceSelected(item);
    };

    auto collectDirectory = [&](auto&& self, const std::filesystem::path& relativeFolder, bool emitShell) noexcept -> CompareSyncManifestStatus
    {
        if (isCoveredByWholeSubtree(relativeFolder))
        {
            return CompareSyncManifestStatus::Ready;
        }

        std::shared_ptr<const CompareDirectoriesFolderDecision> decision;
        CompareSyncManifestStatus status = getDecisionLocked(relativeFolder, decision);
        if (status != CompareSyncManifestStatus::Ready)
        {
            return status;
        }

        if (decision->pendingContentCompareCount != 0u)
        {
            return fail(CompareSyncManifestStatus::NotReady, CompareSyncManifestBlockerReason::ContentPending, relativeFolder, S_FALSE);
        }

        const size_t startCount = out.items.size();
        if (emitShell)
        {
            emitItem(CompareSyncManifestItemKind::DirectoryShell, relativeFolder, FILESYSTEM_FLAG_NONE);
        }

        for (const auto& [name, child] : decision->items)
        {
            if (! sourceExists(child))
            {
                continue;
            }

            const std::filesystem::path childRelative = (relativeFolder / std::filesystem::path(name)).lexically_normal();
            if (isCoveredByWholeSubtree(childRelative))
            {
                continue;
            }

            const CompareSyncManifestBlockerReason childPendingReason = pendingReason(child);
            if (childPendingReason != CompareSyncManifestBlockerReason::None)
            {
                return fail(CompareSyncManifestStatus::NotReady, childPendingReason, childRelative, S_FALSE);
            }

            if (child.isDirectory)
            {
                if (! destinationExists(child))
                {
                    const bool childIsReparse = IsReparsePairEntry(child);
                    emitItem(childIsReparse ? CompareSyncManifestItemKind::File : CompareSyncManifestItemKind::WholeSubtree,
                             childRelative,
                             childIsReparse ? FILESYSTEM_FLAG_NONE : static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE));
                    continue;
                }

                if (! settings.compareSubdirectories)
                {
                    return fail(CompareSyncManifestStatus::Unsupported,
                                CompareSyncManifestBlockerReason::Unsupported,
                                childRelative,
                                HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
                }

                if (child.isDifferent || sourceSelected(child))
                {
                    status = self(self, childRelative, true);
                    if (status != CompareSyncManifestStatus::Ready)
                    {
                        return status;
                    }
                }
                continue;
            }

            if (shouldTransferFile(child))
            {
                emitItem(CompareSyncManifestItemKind::File, childRelative, FILESYSTEM_FLAG_NONE);
            }
        }

        if (emitShell && out.items.size() == startCount + 1u && out.items.back().kind == CompareSyncManifestItemKind::DirectoryShell &&
            CacheKeyEqualsNoCase(MakeCacheKey(out.items.back().relativePath), MakeCacheKey(relativeFolder)))
        {
            emittedKeys.erase(makeEmittedKey(CompareSyncManifestItemKind::DirectoryShell, relativeFolder));
            out.items.pop_back();
        }

        return CompareSyncManifestStatus::Ready;
    };

    for (const auto& relativePath : selected)
    {
        if (relativePath.empty() || isCoveredByWholeSubtree(relativePath))
        {
            continue;
        }

        const std::filesystem::path parentRelative = relativePath.parent_path().lexically_normal();
        const std::wstring leafName                = relativePath.filename().wstring();
        if (leafName.empty())
        {
            return fail(
                CompareSyncManifestStatus::Unsupported, CompareSyncManifestBlockerReason::Unsupported, relativePath, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        }

        std::shared_ptr<const CompareDirectoriesFolderDecision> parentDecision;
        CompareSyncManifestStatus status = getDecisionLocked(parentRelative, parentDecision);
        if (status != CompareSyncManifestStatus::Ready)
        {
            return status;
        }

        const auto itemIt = parentDecision->items.find(leafName);
        if (itemIt == parentDecision->items.end())
        {
            return fail(CompareSyncManifestStatus::NotReady, CompareSyncManifestBlockerReason::MissingDecision, parentRelative, S_FALSE);
        }

        const CompareDirectoriesItemDecision& item = itemIt->second;
        if (! sourceExists(item))
        {
            continue;
        }

        const CompareSyncManifestBlockerReason itemPendingReason = pendingReason(item);
        if (itemPendingReason != CompareSyncManifestBlockerReason::None)
        {
            return fail(CompareSyncManifestStatus::NotReady, itemPendingReason, relativePath, S_FALSE);
        }

        if (item.isDirectory)
        {
            if (! destinationExists(item))
            {
                const bool itemIsReparse = IsReparsePairEntry(item);
                emitItem(itemIsReparse ? CompareSyncManifestItemKind::File : CompareSyncManifestItemKind::WholeSubtree,
                         relativePath,
                         itemIsReparse ? FILESYSTEM_FLAG_NONE : static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE));
                continue;
            }

            if (! settings.compareSubdirectories)
            {
                return fail(CompareSyncManifestStatus::Unsupported,
                            CompareSyncManifestBlockerReason::Unsupported,
                            relativePath,
                            HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
            }

            status = collectDirectory(collectDirectory, relativePath, true);
            if (status != CompareSyncManifestStatus::Ready)
            {
                return status;
            }
            continue;
        }

        if (shouldTransferFile(item))
        {
            emitItem(CompareSyncManifestItemKind::File, relativePath, FILESYSTEM_FLAG_NONE);
        }
    }

    if (out.items.empty())
    {
        blocker.status = CompareSyncManifestStatus::Empty;
        return finish(CompareSyncManifestStatus::Empty);
    }

    blocker.status = CompareSyncManifestStatus::Ready;
    blocker.reason = CompareSyncManifestBlockerReason::None;
    return finish(CompareSyncManifestStatus::Ready);
}

void CompareDirectoriesSession::SetPinnedFolders(const std::filesystem::path& leftRelativeFolder, const std::filesystem::path& rightRelativeFolder) noexcept
{
    std::lock_guard guard(_mutex);

    _decisionCachePinnedKeys.clear();
    _decisionCachePinnedKeys.insert(L".");

    const auto addAncestors = [&](std::filesystem::path rel) noexcept
    {
        rel = rel.lexically_normal();
        for (;;)
        {
            _decisionCachePinnedKeys.insert(MakeCacheKey(rel));
            if (rel.empty())
            {
                break;
            }
            rel = rel.parent_path();
        }
    };

    addAncestors(leftRelativeFolder);
    addAncestors(rightRelativeFolder);
}

bool CompareDirectoriesSession::FlushPendingSubdirUpdatesBudgeted(size_t maxFoldersToApply) noexcept
{
    if (maxFoldersToApply == 0)
    {
        std::lock_guard guard(_mutex);
        return ! _pendingSubdirUpdates.empty();
    }

    std::lock_guard guard(_mutex);
    const uint64_t currentVersion                               = _version.load(std::memory_order_relaxed);
    const Common::Settings::CompareDirectoriesSettings settings = _settings;

    bool anyChanged    = false;
    size_t appliedKeys = 0;
    std::vector<std::wstring> retryKeys;
    while (! _pendingSubdirUpdates.empty() && appliedKeys < maxFoldersToApply)
    {
        const std::wstring key = *_pendingSubdirUpdates.begin();
        _pendingSubdirUpdates.erase(_pendingSubdirUpdates.begin());

        const SubdirPropagationResult result = PropagateChildAggregateToAncestorsLocked(key, settings, currentVersion);
        if (result.changed)
        {
            anyChanged = true;
        }
        if (! result.retryChildKey.empty())
        {
            retryKeys.emplace_back(result.retryChildKey);
        }
        else
        {
            _pendingSubdirAggregateUpdates.erase(key);
            _pendingSubdirAggregateRetryAttempts.erase(key);
        }
        ++appliedKeys;
    }

    for (const std::wstring& retryKey : retryKeys)
    {
        if (! QueueSubdirAggregateRetryLocked(retryKey, currentVersion))
        {
            _pendingSubdirAggregateUpdates.erase(retryKey);
            _pendingSubdirAggregateRetryAttempts.erase(retryKey);
        }
    }

    if (anyChanged)
    {
        ++_uiVersion;
    }

    MaybeEvictDecisionCacheLocked();

    return ! _pendingSubdirUpdates.empty();
}

Common::Settings::CompareDirectoriesSettings CompareDirectoriesSession::GetSettings() const
{
    std::lock_guard guard(_mutex);
    return _settings;
}

std::filesystem::path CompareDirectoriesSession::GetRoot(ComparePane pane) const
{
    std::lock_guard guard(_mutex);
    return pane == ComparePane::Left ? _leftRoot : _rightRoot;
}

uint64_t CompareDirectoriesSession::GetVersion() const noexcept
{
    return _version.load(std::memory_order_acquire);
}

uint64_t CompareDirectoriesSession::GetUiVersion() const noexcept
{
    std::lock_guard guard(_mutex);
    return _uiVersion;
}

CompareDirectoriesPerfStats CompareDirectoriesSession::GetPerfStats() const noexcept
{
    CompareDirectoriesPerfStats stats{};
    {
        std::lock_guard guard(_mutex);

        stats.version   = _version.load(std::memory_order_relaxed);
        stats.uiVersion = _uiVersion;

        stats.scanActiveScans      = _scanActiveScans.load(std::memory_order_relaxed);
        stats.scanFoldersScanned   = _scanFoldersScanned.load(std::memory_order_relaxed);
        stats.scanEntriesScanned   = _scanEntriesScanned.load(std::memory_order_relaxed);
        stats.scanQueueHighSize    = _scanQueueHigh.size();
        stats.scanQueueLowSize     = _scanQueueLow.size();
        stats.scanQueueSize        = stats.scanQueueHighSize + stats.scanQueueLowSize;
        stats.scanScheduledKeys    = _scanScheduledKeys.size();
        stats.scanInFlightKeys     = _scanInFlightKeys.size();
        stats.pendingSubdirUpdates = _pendingSubdirUpdates.size();

        stats.scanQueueHighWater     = _scanQueueHighWater;
        stats.scanQueueHighHighWater = _scanQueueHighHighWater;
        stats.scanQueueLowHighWater  = _scanQueueLowHighWater;
        stats.scanScheduledHighWater = _scanScheduledHighWater;
        stats.scanInFlightHighWater  = _scanInFlightHighWater;
        stats.pendingSubdirHighWater = _pendingSubdirHighWater;

        stats.contentPendingCompares   = _contentComparePendingCompares.load(std::memory_order_relaxed);
        stats.contentTotalCompares     = _contentCompareTotalCompares.load(std::memory_order_relaxed);
        stats.contentCompletedCompares = _contentCompareCompletedCompares.load(std::memory_order_relaxed);
        stats.contentTotalBytes        = _contentCompareTotalBytes.load(std::memory_order_relaxed);
        stats.contentCompletedBytes    = _contentCompareCompletedBytes.load(std::memory_order_relaxed);
        stats.contentQueueHighSize     = _contentCompareQueueHigh.size();
        stats.contentQueueLowSize      = _contentCompareQueueLow.size();
        stats.contentQueueSize         = stats.contentQueueHighSize + stats.contentQueueLowSize;
        stats.contentInFlightSize      = _contentCompareInFlight.size();
        stats.contentCacheSize         = _contentCompareCache.size();
        stats.pendingContentUpdates    = _pendingContentCompareUpdates.size();

        stats.contentQueueHighWater     = _contentQueueHighWater;
        stats.contentQueueHighHighWater = _contentQueueHighHighWater;
        stats.contentQueueLowHighWater  = _contentQueueLowHighWater;
        stats.contentInFlightHighWater  = _contentInFlightHighWater;
        stats.contentCacheHighWater     = _contentCacheHighWater;
        stats.pendingContentHighWater   = _pendingContentHighWater;

        stats.decisionCacheEntries                 = _cache.size();
        stats.decisionCacheEntriesHighWater        = _decisionCacheEntriesHighWater;
        stats.decisionCacheEstimatedBytes          = _decisionCacheEstimatedBytes;
        stats.decisionCacheEstimatedBytesHighWater = _decisionCacheEstimatedBytesHighWater;
        stats.decisionCacheBudgetBytes             = _decisionCacheBudgetBytes;
    }

    stats.directoryInfoCache = DirectoryInfoCache::GetInstance().GetStats();
    return stats;
}

#ifdef ENABLE_TESTS
void CompareDirectoriesSession::SetDecisionCacheBudgetBytesForSelfTest(uint64_t budgetBytes) noexcept
{
    std::lock_guard guard(_mutex);
    _decisionCacheBudgetBytes = budgetBytes;
    MaybeEvictDecisionCacheLocked();
}
#endif

wil::com_ptr<IFileSystem> CompareDirectoriesSession::GetFileSystem(ComparePane pane) const noexcept
{
    return pane == ComparePane::Left ? _leftFileSystem : _rightFileSystem;
}

wil::com_ptr<IInformations> CompareDirectoriesSession::GetInformations(ComparePane pane) const noexcept
{
    return pane == ComparePane::Left ? _leftInformations : _rightInformations;
}

wil::com_ptr<IFileSystemIO> CompareDirectoriesSession::GetFileSystemIO(ComparePane pane) const noexcept
{
    return pane == ComparePane::Left ? _leftFileSystemIo : _rightFileSystemIo;
}

bool CompareDirectoriesSession::IsContentCompareSupported() const noexcept
{
    return _leftFileSystemIo && _rightFileSystemIo;
}

std::optional<std::filesystem::path> CompareDirectoriesSession::TryMakeRelative(ComparePane pane, const std::filesystem::path& absoluteFolder) const
{
    const std::filesystem::path rootPathRaw = GetRoot(pane);
    if (! NavigationLocation::LooksLikeWindowsAbsolutePath(rootPathRaw.native()))
    {
        const std::wstring rootNorm = NavigationLocation::NormalizePluginPathText(rootPathRaw.native(),
                                                                                  NavigationLocation::EmptyPathPolicy::Root,
                                                                                  NavigationLocation::LeadingSlashPolicy::Ensure,
                                                                                  NavigationLocation::TrailingSlashPolicy::Trim);
        const std::wstring absNorm  = NavigationLocation::NormalizePluginPathText(absoluteFolder.native(),
                                                                                  NavigationLocation::EmptyPathPolicy::Root,
                                                                                  NavigationLocation::LeadingSlashPolicy::Ensure,
                                                                                  NavigationLocation::TrailingSlashPolicy::Trim);

        if (absNorm == rootNorm)
        {
            return std::filesystem::path{};
        }

        std::wstring rootPrefix = rootNorm;
        if (! rootPrefix.empty() && rootPrefix.back() != L'/')
        {
            rootPrefix.push_back(L'/');
        }

        if (absNorm.rfind(rootPrefix, 0) != 0)
        {
            return std::nullopt;
        }

        const std::wstring_view relativeText(absNorm.c_str() + rootPrefix.size());
        std::filesystem::path relative(relativeText);
        relative = relative.lexically_normal();
        return relative;
    }

    const std::filesystem::path rootPath = rootPathRaw.lexically_normal();
    const std::filesystem::path absPath  = absoluteFolder.lexically_normal();

    auto normalizeText = [](std::filesystem::path value, bool lower) noexcept -> std::wstring
    {
        value = value.lexically_normal();
        while (! value.empty() && ! value.has_filename() && value != value.root_path())
        {
            value = value.parent_path();
        }

        std::wstring text = value.wstring();
        std::replace(text.begin(), text.end(), L'/', L'\\');

        if (text.rfind(L"\\\\?\\UNC\\", 0) == 0)
        {
            text.erase(0, 8);
            text.insert(0, L"\\\\");
        }
        else if (text.rfind(L"\\\\?\\", 0) == 0)
        {
            text.erase(0, 4);
        }

        if (lower && ! text.empty())
        {
            ::CharLowerBuffW(text.data(), static_cast<DWORD>(text.size()));
        }
        return text;
    };

    const std::wstring rootNorm      = normalizeText(rootPath, false);
    const std::wstring absNorm       = normalizeText(absPath, false);
    const std::wstring rootNormLower = normalizeText(rootPath, true);
    const std::wstring absNormLower  = normalizeText(absPath, true);

    if (absNormLower == rootNormLower)
    {
        return std::filesystem::path{};
    }

    std::wstring rootPrefix      = rootNorm;
    std::wstring rootPrefixLower = rootNormLower;
    if (! rootPrefix.empty() && rootPrefix.back() != L'\\')
    {
        rootPrefix.push_back(L'\\');
        rootPrefixLower.push_back(L'\\');
    }

    if (absNormLower.rfind(rootPrefixLower, 0) != 0)
    {
        return std::nullopt;
    }

    const std::wstring_view relativeText(absNorm.c_str() + rootPrefix.size());
    std::filesystem::path relative(relativeText);
    relative = relative.lexically_normal();
    return relative;
}

std::filesystem::path CompareDirectoriesSession::ResolveAbsolute(ComparePane pane, const std::filesystem::path& relativeFolder) const
{
    const std::filesystem::path root = GetRoot(pane);
    return ResolveAbsoluteFromRoot(root, relativeFolder);
}

std::wstring CompareDirectoriesSession::MakeCacheKey(const std::filesystem::path& relativeFolder) const
{
    if (relativeFolder.empty())
    {
        return std::wstring(L".");
    }
    return relativeFolder.generic_wstring();
}

uint64_t CompareDirectoriesSession::EstimateDecisionBytes(std::wstring_view folderKey, const CompareDirectoriesFolderDecision& decision) noexcept
{
    // Heuristic budget accounting for decision cache eviction.
    // Over-estimation is intentional: it keeps memory bounded without needing allocator-specific introspection.
    constexpr uint64_t kDecisionOverheadBytes = 512;
    constexpr uint64_t kItemNodeOverheadBytes = 160;

    uint64_t bytes = kDecisionOverheadBytes;
    bytes += static_cast<uint64_t>(folderKey.size()) * sizeof(wchar_t);

    bytes += static_cast<uint64_t>(decision.items.size()) * kItemNodeOverheadBytes;
    for (const auto& [name, item] : decision.items)
    {
        bytes += static_cast<uint64_t>(name.size()) * sizeof(wchar_t);
        bytes += sizeof(item);
    }

    return bytes;
}

void CompareDirectoriesSession::TouchDecisionCacheKeyLocked(std::wstring_view folderKey) noexcept
{
    const auto metaIt = _decisionCacheMeta.find(folderKey);
    if (metaIt == _decisionCacheMeta.end())
    {
        return;
    }

    _decisionCacheLru.splice(_decisionCacheLru.begin(), _decisionCacheLru, metaIt->second.lruIt);
}

void CompareDirectoriesSession::TrackDecisionCacheInsertOrUpdateLocked(std::wstring_view folderKey,
                                                                       const std::shared_ptr<const CompareDirectoriesFolderDecision>& decision) noexcept
{
    if (! decision)
    {
        return;
    }

    const uint64_t estimatedBytes = EstimateDecisionBytes(folderKey, *decision);

    auto metaIt = _decisionCacheMeta.find(folderKey);
    if (metaIt == _decisionCacheMeta.end())
    {
        std::wstring key(folderKey);
        _decisionCacheLru.emplace_front(key);

        DecisionCacheLruEntry entry{};
        entry.estimatedBytes = estimatedBytes;
        entry.lruIt          = _decisionCacheLru.begin();

        metaIt = _decisionCacheMeta.emplace(std::move(key), entry).first;

        _decisionCacheEstimatedBytes += estimatedBytes;
    }
    else
    {
        const uint64_t oldBytes       = metaIt->second.estimatedBytes;
        metaIt->second.estimatedBytes = estimatedBytes;

        if (estimatedBytes >= oldBytes)
        {
            _decisionCacheEstimatedBytes += estimatedBytes - oldBytes;
        }
        else
        {
            const uint64_t delta         = oldBytes - estimatedBytes;
            _decisionCacheEstimatedBytes = (delta >= _decisionCacheEstimatedBytes) ? 0 : (_decisionCacheEstimatedBytes - delta);
        }

        _decisionCacheLru.splice(_decisionCacheLru.begin(), _decisionCacheLru, metaIt->second.lruIt);
    }

    _decisionCacheEstimatedBytesHighWater = std::max(_decisionCacheEstimatedBytesHighWater, _decisionCacheEstimatedBytes);
}

void CompareDirectoriesSession::TrackDecisionCacheEraseLocked(std::wstring_view folderKey) noexcept
{
    const auto metaIt = _decisionCacheMeta.find(folderKey);
    if (metaIt == _decisionCacheMeta.end())
    {
        return;
    }

    const uint64_t oldBytes      = metaIt->second.estimatedBytes;
    _decisionCacheEstimatedBytes = (oldBytes >= _decisionCacheEstimatedBytes) ? 0 : (_decisionCacheEstimatedBytes - oldBytes);

    _decisionCacheLru.erase(metaIt->second.lruIt);
    _decisionCacheMeta.erase(metaIt);
}

void CompareDirectoriesSession::MaybeEvictDecisionCacheLocked() noexcept
{
    if (_decisionCacheEstimatedBytes <= _decisionCacheBudgetBytes)
    {
        return;
    }

    // Eviction can be expensive on large trees; keep it bounded so we don't stall other threads
    // contending on _mutex (including the UI thread).
    constexpr size_t kMaxEvictionsPerCall     = 64;
    constexpr size_t kMaxKeysInspectedPerCall = 16384;

    std::vector<std::wstring> candidates;
    candidates.reserve(kMaxEvictionsPerCall);

    uint64_t projectedBytes = _decisionCacheEstimatedBytes;
    size_t inspectedKeys    = 0;

    for (auto it = _decisionCacheLru.rbegin(); it != _decisionCacheLru.rend() && projectedBytes > _decisionCacheBudgetBytes &&
                                               candidates.size() < kMaxEvictionsPerCall && inspectedKeys < kMaxKeysInspectedPerCall;
         ++it)
    {
        ++inspectedKeys;

        const std::wstring_view keyView = *it;
        if (keyView == L"." || _decisionCachePinnedKeys.find(keyView) != _decisionCachePinnedKeys.end() ||
            _scanInFlightKeys.find(keyView) != _scanInFlightKeys.end())
        {
            continue;
        }

        const auto metaIt             = _decisionCacheMeta.find(keyView);
        const uint64_t estimatedBytes = metaIt != _decisionCacheMeta.end() ? metaIt->second.estimatedBytes : 0u;
        candidates.emplace_back(keyView);
        projectedBytes = estimatedBytes >= projectedBytes ? 0u : projectedBytes - estimatedBytes;
    }

    for (const std::wstring& candidateKey : candidates)
    {
        _cache.erase(candidateKey);
        TrackDecisionCacheEraseLocked(candidateKey);
    }
}

void CompareDirectoriesSession::NotifyScanProgress(const std::filesystem::path& relativeFolder, std::wstring_view currentEntryName, bool force) noexcept
{
    const auto cb = _scanProgressCallback.load(std::memory_order_acquire);
    if (! cb || ! *cb)
    {
        return;
    }

    if (! force)
    {
        const uint64_t now  = GetTickCount64();
        uint64_t lastUpdate = _scanLastNotifyTickMs.load(std::memory_order_relaxed);
        if ((now - lastUpdate) < 80u)
        {
            return;
        }

        if (! _scanLastNotifyTickMs.compare_exchange_strong(lastUpdate, now, std::memory_order_relaxed, std::memory_order_relaxed))
        {
            return;
        }
    }

    const uint64_t scannedFolders             = _scanFoldersScanned.load(std::memory_order_relaxed);
    const uint64_t scannedEntries             = _scanEntriesScanned.load(std::memory_order_relaxed);
    const uint32_t activeScans                = _scanActiveScans.load(std::memory_order_relaxed);
    const uint64_t contentCandidateFileCount  = _contentCompareTotalCompares.load(std::memory_order_relaxed);
    const uint64_t contentCandidateTotalBytes = _contentCompareTotalBytes.load(std::memory_order_relaxed);

    (*cb)(relativeFolder, currentEntryName, scannedFolders, scannedEntries, activeScans, contentCandidateFileCount, contentCandidateTotalBytes);
}

void CompareDirectoriesSession::NotifyContentProgress(
    uint32_t workerIndex, const std::filesystem::path& relativeFolder, std::wstring_view entryName, uint64_t totalBytes, uint64_t completedBytes) noexcept
{
    const auto cb = _contentProgressCallback.load(std::memory_order_acquire);
    if (! cb || ! *cb)
    {
        return;
    }

    const uint64_t pending           = _contentComparePendingCompares.load(std::memory_order_relaxed);
    const uint64_t totalCompares     = _contentCompareTotalCompares.load(std::memory_order_relaxed);
    const uint64_t completedCompares = _contentCompareCompletedCompares.load(std::memory_order_relaxed);

    const uint64_t overallTotalBytes        = _contentCompareTotalBytes.load(std::memory_order_relaxed);
    const uint64_t overallCompletedBytesRaw = _contentCompareCompletedBytes.load(std::memory_order_relaxed);
    const uint64_t overallCompletedBytes    = std::min(overallCompletedBytesRaw, overallTotalBytes);

    (*cb)(workerIndex,
          relativeFolder,
          entryName,
          totalBytes,
          completedBytes,
          overallTotalBytes,
          overallCompletedBytes,
          pending,
          totalCompares,
          completedCompares);
}

void CompareDirectoriesSession::NotifyDecisionUpdated(bool force) noexcept
{
    const auto cb = _decisionUpdatedCallback.load(std::memory_order_acquire);
    if (! cb || ! *cb)
    {
        return;
    }

    if (! force)
    {
        const uint64_t now  = GetTickCount64();
        uint64_t lastUpdate = _decisionUpdatedLastNotifyTickMs.load(std::memory_order_relaxed);
        if ((now - lastUpdate) < 120u)
        {
            return;
        }

        if (! _decisionUpdatedLastNotifyTickMs.compare_exchange_strong(lastUpdate, now, std::memory_order_relaxed, std::memory_order_relaxed))
        {
            return;
        }
    }

    (*cb)();
}

void CompareDirectoriesSession::EnsureScanWorkersLocked() noexcept
{
    if (! _scanWorkers.empty())
    {
        return;
    }

    unsigned int workers = std::thread::hardware_concurrency();
    if (workers == 0u)
    {
        workers = 2u;
    }
    workers = std::clamp(workers, 1u, 4u);

    _scanWorkers.reserve(workers);
    for (unsigned int i = 0; i < workers; ++i)
    {
        const uint32_t workerIndex = i;
        _scanWorkers.emplace_back([this, workerIndex](std::stop_token stopToken) noexcept { ScanWorker(stopToken, workerIndex); });
    }
}

void CompareDirectoriesSession::EnqueueScanLocked(const std::filesystem::path& relativeFolder,
                                                  uint64_t version,
                                                  uint64_t cancelToken,
                                                  ScanPriority priority) noexcept
{
    const std::wstring key      = MakeCacheKey(relativeFolder);
    const bool alreadyScheduled = _scanScheduledKeys.contains(key);

    FolderScanJob job{};
    job.version        = version;
    job.cancelToken    = cancelToken;
    job.relativeFolder = relativeFolder;
    job.key            = key;
    job.priority       = priority;

    if (alreadyScheduled)
    {
        // Visible-first: if a folder is already scheduled as low-priority (background subtree scan),
        // and the UI requests it, queue a single high-priority duplicate to ensure it is serviced soon.
        // The duplicate is safe because ScanWorker gates execution on _scanInFlightKeys and _scanScheduledKeys.
        if (priority == ScanPriority::High && ! _scanHighQueuedKeys.contains(key) && ! _scanInFlightKeys.contains(key))
        {
            job.priority = ScanPriority::High;
            _scanQueueHigh.emplace_back(std::move(job));
            _scanHighQueuedKeys.insert(key);

            _scanQueueHighWater     = std::max(_scanQueueHighWater, _scanQueueHigh.size() + _scanQueueLow.size());
            _scanQueueHighHighWater = std::max(_scanQueueHighHighWater, _scanQueueHigh.size());
            _scanQueueLowHighWater  = std::max(_scanQueueLowHighWater, _scanQueueLow.size());
            _scanScheduledHighWater = std::max(_scanScheduledHighWater, _scanScheduledKeys.size());

            _scanCv.notify_one();
        }
        return;
    }

    _scanScheduledKeys.insert(key);
    if (priority == ScanPriority::High)
    {
        _scanQueueHigh.emplace_back(std::move(job));
        _scanHighQueuedKeys.insert(key);
    }
    else
    {
        _scanQueueLow.emplace_back(std::move(job));
    }
    _scanQueueHighWater     = std::max(_scanQueueHighWater, _scanQueueHigh.size() + _scanQueueLow.size());
    _scanQueueHighHighWater = std::max(_scanQueueHighHighWater, _scanQueueHigh.size());
    _scanQueueLowHighWater  = std::max(_scanQueueLowHighWater, _scanQueueLow.size());
    _scanScheduledHighWater = std::max(_scanScheduledHighWater, _scanScheduledKeys.size());

    const uint32_t activeBefore = _scanActiveScans.fetch_add(1u, std::memory_order_acq_rel);
    if (activeBefore == 0u)
    {
        _scanFoldersScanned.store(0u, std::memory_order_release);
        _scanEntriesScanned.store(0u, std::memory_order_release);
        _scanLastNotifyTickMs.store(0u, std::memory_order_release);
    }

    _scanCv.notify_one();
}

void CompareDirectoriesSession::EnsureContentCompareWorkersLocked() noexcept
{
    if (! IsContentCompareSupported() || ! _contentCompareWorkers.empty())
    {
        return;
    }

    unsigned int workers = 0;
    if (_settings.contentCompareWorkerCount == 0u)
    {
        workers = std::thread::hardware_concurrency();
        if (workers == 0u)
        {
            workers = 2u;
        }
        workers = std::clamp(workers, 1u, 4u);
    }
    else
    {
        workers = std::clamp(static_cast<unsigned int>(_settings.contentCompareWorkerCount), 1u, 4u);
    }

    _contentCompareWorkers.reserve(workers);
    for (unsigned int i = 0; i < workers; ++i)
    {
        const uint32_t workerIndex = i;
        _contentCompareWorkers.emplace_back([this, workerIndex](std::stop_token stopToken) noexcept { ContentCompareWorker(stopToken, workerIndex); });
    }
}

void CompareDirectoriesSession::ScheduleResetCleanup(std::unique_ptr<ResetCleanup> cleanup) noexcept
{
    QueueCompareCleanup(std::move(cleanup), L"reset cleanup");
}

void CompareDirectoriesSession::ResetCompareStateLocked(ResetCleanup& outCleanup, bool clearContentCompareCache) noexcept
{
    outCleanup.cache.swap(_cache);
    outCleanup.decisionCacheLru.swap(_decisionCacheLru);
    outCleanup.decisionCacheMeta.swap(_decisionCacheMeta);
    outCleanup.decisionCachePinnedKeys.swap(_decisionCachePinnedKeys);
    _decisionCacheEstimatedBytes = 0;
    outCleanup.scanQueueHigh.swap(_scanQueueHigh);
    outCleanup.scanQueueLow.swap(_scanQueueLow);
    outCleanup.scanScheduledKeys.swap(_scanScheduledKeys);
    outCleanup.scanHighQueuedKeys.swap(_scanHighQueuedKeys);
    outCleanup.scanInFlightKeys.swap(_scanInFlightKeys);
    outCleanup.pendingSubdirUpdates.swap(_pendingSubdirUpdates);
    outCleanup.pendingSubdirAggregateUpdates.swap(_pendingSubdirAggregateUpdates);
    outCleanup.pendingSubdirAggregateRetryAttempts.swap(_pendingSubdirAggregateRetryAttempts);
    outCleanup.contentCompareInFlight.swap(_contentCompareInFlight);
    outCleanup.contentCompareQueueHigh.swap(_contentCompareQueueHigh);
    outCleanup.contentCompareQueueLow.swap(_contentCompareQueueLow);
    outCleanup.pendingContentCompareUpdates.swap(_pendingContentCompareUpdates);
    outCleanup.pendingContentCompareRetryAttempts.swap(_pendingContentCompareRetryAttempts);
    if (clearContentCompareCache)
    {
        _contentCompareCache.clear();
    }

    _scanActiveScans.store(0u, std::memory_order_release);
    _scanFoldersScanned.store(0u, std::memory_order_release);
    _scanEntriesScanned.store(0u, std::memory_order_release);
    _scanLastNotifyTickMs.store(0u, std::memory_order_release);
    _scanCv.notify_all();

    _contentComparePendingCompares.store(0u, std::memory_order_release);
    _contentCompareTotalCompares.store(0u, std::memory_order_release);
    _contentCompareCompletedCompares.store(0u, std::memory_order_release);
    _contentCompareTotalBytes.store(0u, std::memory_order_release);
    _contentCompareCompletedBytes.store(0u, std::memory_order_release);

    _contentCompareCv.notify_all();
    _contentCompareQueueNotFullCv.notify_all();
}

void CompareDirectoriesSession::ClearContentCompareStateLocked() noexcept
{
    _contentCompareQueueHigh.clear();
    _contentCompareQueueLow.clear();
    _contentCompareInFlight.clear();
    _pendingContentCompareUpdates.clear();
    _pendingContentCompareRetryAttempts.clear();
    _contentCompareCache.clear();
    _contentComparePendingCompares.store(0u, std::memory_order_release);
    _contentCompareTotalCompares.store(0u, std::memory_order_release);
    _contentCompareCompletedCompares.store(0u, std::memory_order_release);
    _contentCompareTotalBytes.store(0u, std::memory_order_release);
    _contentCompareCompletedBytes.store(0u, std::memory_order_release);

    _contentCompareCv.notify_all();
    _contentCompareQueueNotFullCv.notify_all();
}

void CompareDirectoriesSession::EvictContentCompareCacheForRelativePathLocked(const std::filesystem::path& relativePath, bool includeSubtree) noexcept
{
    const std::filesystem::path normalizedPath = relativePath.lexically_normal();
    if (normalizedPath.empty())
    {
        _contentCompareCache.clear();
        return;
    }

    const std::wstring relativeKey = MakeCacheKey(normalizedPath);
    for (auto it = _contentCompareCache.begin(); it != _contentCompareCache.end();)
    {
        const std::wstring_view cachedKey = it->first.relativeFileKey;
        const bool match                  = includeSubtree ? CacheKeyIsSameOrDescendant(cachedKey, relativeKey) : CacheKeyEqualsNoCase(cachedKey, relativeKey);
        if (match)
        {
            it = _contentCompareCache.erase(it);
        }
        else
        {
            ++it;
        }
    }
}

bool CompareDirectoriesSession::QueueSubdirAggregateRetryLocked(std::wstring_view childKey, uint64_t currentVersion) noexcept
{
    if (childKey.empty() || childKey == L"." || ! _backgroundWorkEnabled.load(std::memory_order_relaxed) ||
        _version.load(std::memory_order_relaxed) != currentVersion)
    {
        return false;
    }

    const auto aggregateIt = _pendingSubdirAggregateUpdates.find(childKey);
    if (aggregateIt == _pendingSubdirAggregateUpdates.end() || aggregateIt->second.version != currentVersion)
    {
        return false;
    }

    std::filesystem::path childRel = std::filesystem::path(childKey).lexically_normal();
    if (childRel.empty())
    {
        return false;
    }

    const std::wstring retryKey(childKey);
    uint32_t& attempts = _pendingSubdirAggregateRetryAttempts[retryKey];
    if (attempts >= kMaxPendingUpdateRetryAttempts)
    {
        return false;
    }
    ++attempts;

    const std::filesystem::path parentRel = childRel.parent_path();
    const uint64_t cancelToken            = _backgroundWorkCancelToken.load(std::memory_order_relaxed);
    EnsureScanWorkersLocked();
    EnqueueScanLocked(parentRel, currentVersion, cancelToken, ScanPriority::High);
    _pendingSubdirUpdates.insert(std::wstring(childKey));
    _pendingSubdirHighWater = std::max(_pendingSubdirHighWater, _pendingSubdirUpdates.size());
    return true;
}

bool CompareDirectoriesSession::QueueContentUpdateFolderRetryLocked(std::wstring_view folderKey, uint64_t currentVersion) noexcept
{
    if (folderKey.empty() || ! _backgroundWorkEnabled.load(std::memory_order_relaxed) || _version.load(std::memory_order_relaxed) != currentVersion)
    {
        return false;
    }

    const auto pendingIt = _pendingContentCompareUpdates.find(folderKey);
    if (pendingIt == _pendingContentCompareUpdates.end() || pendingIt->second.empty())
    {
        return false;
    }

    std::filesystem::path folderRel = std::filesystem::path(folderKey).lexically_normal();
    if (folderRel.native() == L".")
    {
        folderRel.clear();
    }

    const std::wstring retryKey(folderKey);
    uint32_t& attempts = _pendingContentCompareRetryAttempts[retryKey];
    if (attempts >= kMaxPendingUpdateRetryAttempts)
    {
        return false;
    }
    ++attempts;

    const uint64_t cancelToken = _backgroundWorkCancelToken.load(std::memory_order_relaxed);
    EnsureScanWorkersLocked();
    EnqueueScanLocked(folderRel, currentVersion, cancelToken, ScanPriority::High);
    return true;
}

CompareDirectoriesSession::SubdirPropagationResult CompareDirectoriesSession::PropagateChildAggregateToAncestorsLocked(
    std::wstring_view childKeyView, const Common::Settings::CompareDirectoriesSettings& settings, uint64_t currentVersion) noexcept
{
    if (! settings.compareSubdirectories)
    {
        return {};
    }

    if (childKeyView.empty() || childKeyView == L".")
    {
        return {};
    }

    std::filesystem::path childRel = std::filesystem::path(childKeyView).lexically_normal();
    if (childRel.empty())
    {
        return {};
    }

    SubdirPropagationResult result{};
    for (;;)
    {
        const std::filesystem::path parentRel = childRel.parent_path();
        const std::wstring parentKey          = MakeCacheKey(parentRel);
        const std::wstring childKey           = MakeCacheKey(childRel);

        const auto parentIt = _cache.find(parentKey);
        if (parentIt == _cache.end() || ! parentIt->second || parentIt->second->version != currentVersion)
        {
            if (_backgroundWorkEnabled.load(std::memory_order_relaxed) && _version.load(std::memory_order_relaxed) == currentVersion)
            {
                result.retryChildKey = std::wstring(childKeyView);
            }
            break;
        }

        HRESULT childHr        = E_FAIL;
        bool childAnyPending   = false;
        bool childAnyDifferent = false;
        const auto childIt     = _cache.find(childKey);
        if (childIt != _cache.end() && childIt->second && childIt->second->version == currentVersion)
        {
            childHr           = childIt->second->hr;
            childAnyPending   = childIt->second->anyPending;
            childAnyDifferent = childIt->second->anyDifferent;
        }
        else
        {
            const auto aggregateIt = _pendingSubdirAggregateUpdates.find(childKey);
            if (aggregateIt == _pendingSubdirAggregateUpdates.end() || aggregateIt->second.version != currentVersion)
            {
                break;
            }

            childHr           = aggregateIt->second.hr;
            childAnyPending   = aggregateIt->second.anyPending;
            childAnyDifferent = aggregateIt->second.anyDifferent;
        }

        const std::wstring childName = childRel.filename().wstring();
        auto updatedParent           = std::make_shared<CompareDirectoriesFolderDecision>(*parentIt->second);

        const auto itemIt = updatedParent->items.find(childName);
        if (itemIt == updatedParent->items.end())
        {
            break;
        }

        CompareDirectoriesItemDecision& item = itemIt->second;
        if (! item.existsLeft || ! item.existsRight)
        {
            break;
        }

        const bool leftIsDir  = (item.leftFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const bool rightIsDir = (item.rightFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (! leftIsDir || ! rightIsDir)
        {
            break;
        }

        // Avoid following directory reparse points (symlinks/junctions).
        if (IsReparsePairEntry(item))
        {
            break;
        }

        const bool childPending   = SUCCEEDED(childHr) && childAnyPending;
        const bool childDifferent = FAILED(childHr) || childAnyDifferent;

        const uint32_t oldMask    = item.differenceMask;
        const bool oldDifferent   = item.isDifferent;
        const bool oldSelectLeft  = item.selectLeft;
        const bool oldSelectRight = item.selectRight;

        const uint32_t subtreeMask =
            static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirContent) | static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirPending);

        uint32_t newMask        = oldMask & ~subtreeMask;
        const uint32_t baseMask = newMask;
        if (childPending)
        {
            newMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirPending);
        }
        if (childDifferent)
        {
            newMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirContent);
        }

        const bool baseDifferent = baseMask != 0u;
        const bool newDifferent  = baseDifferent || childDifferent;

        bool newSelectLeft  = false;
        bool newSelectRight = false;

        if (HasFlag(newMask, CompareDirectoriesDiffBit::OnlyInLeft))
        {
            newSelectLeft = settings.selectSubdirsOnlyInOnePane;
        }
        if (HasFlag(newMask, CompareDirectoriesDiffBit::OnlyInRight))
        {
            newSelectRight = settings.selectSubdirsOnlyInOnePane;
        }
        if (HasFlag(newMask, CompareDirectoriesDiffBit::TypeMismatch) || HasFlag(newMask, CompareDirectoriesDiffBit::SubdirAttributes) ||
            HasFlag(newMask, CompareDirectoriesDiffBit::SubdirContent))
        {
            newSelectLeft  = true;
            newSelectRight = true;
        }

        const bool changed = (newMask != oldMask) || (newDifferent != oldDifferent) || (newSelectLeft != oldSelectLeft) || (newSelectRight != oldSelectRight);
        if (! changed)
        {
            break;
        }

        item.differenceMask = newMask;
        item.isDifferent    = newDifferent;
        item.selectLeft     = newSelectLeft;
        item.selectRight    = newSelectRight;

        // Recompute aggregate flags so the next ancestor iteration can rely on them.
        updatedParent->anyDifferent = AnyChildDifferent(*updatedParent);
        updatedParent->anyPending   = AnyChildPending(*updatedParent);

        const std::shared_ptr<const CompareDirectoriesFolderDecision> storedUpdatedParent = updatedParent;
        _cache[parentKey]                                                                 = std::move(updatedParent);
        TrackDecisionCacheInsertOrUpdateLocked(parentKey, storedUpdatedParent);
        result.changed = true;

        childRel = parentRel;
        if (childRel.empty())
        {
            break;
        }
    }

    if (result.changed)
    {
        MaybeEvictDecisionCacheLocked();
    }

    return result;
}

void CompareDirectoriesSession::ApplyPendingContentCompareUpdatesLocked(const std::wstring& folderKey) noexcept
{
    const auto pendingIt = _pendingContentCompareUpdates.find(folderKey);
    if (pendingIt == _pendingContentCompareUpdates.end())
    {
        return;
    }

    const uint64_t currentVersion = _version.load(std::memory_order_relaxed);

    // Drop stale updates eagerly to avoid accumulating work after resets / version bumps.
    for (auto it = pendingIt->second.begin(); it != pendingIt->second.end();)
    {
        if (it->second.version != currentVersion)
        {
            it = pendingIt->second.erase(it);
            continue;
        }
        ++it;
    }

    if (pendingIt->second.empty())
    {
        _pendingContentCompareRetryAttempts.erase(folderKey);
        _pendingContentCompareUpdates.erase(pendingIt);
        return;
    }

    const auto cacheIt = _cache.find(folderKey);
    if (cacheIt == _cache.end() || ! cacheIt->second || cacheIt->second->version != currentVersion)
    {
        // Decision was not cached yet, or was evicted while content work was in flight.
        // Keep the updates and rescan the folder so the final content diff is not lost.
        if (! QueueContentUpdateFolderRetryLocked(folderKey, currentVersion))
        {
            _pendingContentCompareRetryAttempts.erase(folderKey);
            _pendingContentCompareUpdates.erase(pendingIt);
        }
        return;
    }

    bool anyApplied                                             = false;
    auto updated                                                = std::make_shared<CompareDirectoriesFolderDecision>(*cacheIt->second);
    const Common::Settings::CompareDirectoriesSettings settings = _settings;

    for (const auto& [entryName, update] : pendingIt->second)
    {
        const auto itemIt = updated->items.find(entryName);
        if (itemIt == updated->items.end())
        {
            // In differences-only mode, file-level ContentPending entries can be elided (tracked per-folder).
            if (updated->pendingContentCompareCount != 0u)
            {
                --updated->pendingContentCompareCount;
                anyApplied = true;
            }

            const bool sizeDifferent     = update.leftSizeBytes != update.rightSizeBytes;
            const bool canCompareContent = settings.compareContent && IsContentCompareSupported();
            const bool contentDifferent  = canCompareContent ? (sizeDifferent || ! update.areEqual) : false;

            CompareDirectoriesItemDecision item{};
            item.existsLeft          = true;
            item.existsRight         = true;
            item.isDirectory         = false;
            item.leftSizeBytes       = update.leftSizeBytes;
            item.rightSizeBytes      = update.rightSizeBytes;
            item.leftLastWriteTime   = update.leftLastWriteTime;
            item.rightLastWriteTime  = update.rightLastWriteTime;
            item.leftFileAttributes  = update.leftFileAttributes;
            item.rightFileAttributes = update.rightFileAttributes;

            ApplyCriteriaDiffAndSelection(item, settings, canCompareContent, contentDifferent);
            if (item.isDifferent)
            {
                updated->items.emplace(entryName, std::move(item));
                anyApplied = true;
            }
            continue;
        }

        CompareDirectoriesItemDecision& item = itemIt->second;
        if (! item.existsLeft || ! item.existsRight)
        {
            continue;
        }

        const bool leftIsDir  = (item.leftFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const bool rightIsDir = (item.rightFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (leftIsDir || rightIsDir)
        {
            continue;
        }

        // Skip if metadata no longer matches the queued signature (file changed while the job ran).
        if (item.leftSizeBytes != update.leftSizeBytes || item.rightSizeBytes != update.rightSizeBytes || item.leftLastWriteTime != update.leftLastWriteTime ||
            item.rightLastWriteTime != update.rightLastWriteTime || item.leftFileAttributes != update.leftFileAttributes ||
            item.rightFileAttributes != update.rightFileAttributes)
        {
            continue;
        }

        const bool sizeDifferent     = item.leftSizeBytes != item.rightSizeBytes;
        const bool canCompareContent = settings.compareContent && IsContentCompareSupported();
        const bool contentDifferent  = canCompareContent ? (sizeDifferent || ! update.areEqual) : false;

        ApplyCriteriaDiffAndSelection(item, settings, canCompareContent, contentDifferent);

        anyApplied = true;
    }

    _pendingContentCompareRetryAttempts.erase(folderKey);
    _pendingContentCompareUpdates.erase(pendingIt);

    if (anyApplied)
    {
        // Recompute aggregate flags after applying updates, so ancestor propagation can use them.
        updated->anyDifferent                                                       = AnyChildDifferent(*updated);
        updated->anyPending                                                         = AnyChildPending(*updated);
        const std::shared_ptr<const CompareDirectoriesFolderDecision> storedUpdated = updated;
        _cache[folderKey]                                                           = updated;
        TrackDecisionCacheInsertOrUpdateLocked(folderKey, storedUpdated);

        const SubdirPropagationResult propagation = PropagateChildAggregateToAncestorsLocked(folderKey, settings, currentVersion);
        if (! propagation.retryChildKey.empty())
        {
            if (! QueueSubdirAggregateRetryLocked(propagation.retryChildKey, currentVersion))
            {
                _pendingSubdirAggregateUpdates.erase(propagation.retryChildKey);
                _pendingSubdirAggregateRetryAttempts.erase(propagation.retryChildKey);
            }
        }
        MaybeEvictDecisionCacheLocked();
        ++_uiVersion;
    }
}

namespace
{
[[nodiscard]] std::wstring_view NormalizeEntryNameForCompare(std::wstring_view name) noexcept
{
    // Normalize names to reduce false mismatches across different enumeration backends
    // (e.g. handle-based vs FindFirstFile enumeration) and Win32 vs NT path semantics.
    // In particular, Win32 path parsing treats trailing spaces/dots as insignificant.
    size_t length = 0;
    while (length < name.size() && name[length] != L'\0')
    {
        ++length;
    }

    size_t end = length;
    while (end > 0)
    {
        const wchar_t ch = name[end - 1];
        if (ch == L' ' || ch == L'.')
        {
            --end;
            continue;
        }
        break;
    }

    if (end == 0)
    {
        end = length;
    }

    return name.substr(0, end);
}

void InsertSideEntryPreservingNormalizedCollisions(std::map<std::wstring, SideEntry, WStringViewNoCaseLess>& outEntries,
                                                   std::wstring normalizedKey,
                                                   SideEntry entry)
{
    auto [it, inserted] = outEntries.try_emplace(normalizedKey, std::move(entry));
    if (inserted)
    {
        return;
    }

    if (! OrdinalString::EqualsNoCase(it->first, it->second.name))
    {
        std::wstring restoreKey = it->first;
        auto node               = outEntries.extract(it);
        node.key()              = node.mapped().name;

        auto relocated = outEntries.insert(std::move(node));
        if (! relocated.inserted)
        {
            auto restoreNode  = std::move(relocated.node);
            restoreNode.key() = std::move(restoreKey);
            static_cast<void>(outEntries.insert(std::move(restoreNode)));
        }
    }

    std::wstring originalKey = entry.name;
    static_cast<void>(outEntries.try_emplace(std::move(originalKey), std::move(entry)));
}

[[nodiscard]] bool TryReadDirectoryEntries(const wil::com_ptr<IFileSystem>& baseFs,
                                           const std::filesystem::path& absoluteFolder,
                                           const Common::Settings::CompareDirectoriesSettings& settings,
                                           const std::vector<std::wstring>& ignoreFilePatterns,
                                           const std::vector<std::wstring>& ignoreDirectoryPatterns,
                                           std::map<std::wstring, SideEntry, WStringViewNoCaseLess>& outEntries,
                                           bool& outFolderMissing,
                                           HRESULT& outHr) noexcept
{
    outEntries.clear();
    outFolderMissing = false;
    outHr            = S_OK;

    if (! baseFs)
    {
        outHr = E_POINTER;
        return false;
    }

    // Compare subtree scans can touch a very large number of folders (especially on remote filesystems).
    // Avoid enumerating through DirectoryInfoCache here: it is a global cache sized for interactive browsing
    // and can legitimately grow to GiB scale during large scans.
    wil::com_ptr<IFilesInformation> info;
    Debug::Perf::Scope perf(L"CompareDirectories.ReadDirectoryInfo");
    perf.SetDetail(absoluteFolder.native());
    const HRESULT hr = baseFs->ReadDirectoryInfo(absoluteFolder.c_str(), info.put());
    perf.SetHr(hr);

    if (FAILED(hr))
    {
        if (IsMissingPathError(hr))
        {
            outFolderMissing = true;
            outHr            = S_OK;
            return true;
        }

        outHr = hr;
        return false;
    }

    if (! info)
    {
        outHr = E_UNEXPECTED;
        return false;
    }

    FileInfo* head         = nullptr;
    const HRESULT hrBuffer = info->GetBuffer(&head);
    if (FAILED(hrBuffer))
    {
        outHr = hrBuffer;
        return false;
    }

    unsigned long bufferSizeBytes = 0;
    const HRESULT hrSize          = info->GetBufferSize(&bufferSizeBytes);
    if (FAILED(hrSize))
    {
        outHr = hrSize;
        return false;
    }

    if (head == nullptr)
    {
        outHr = S_OK;
        return true;
    }

    if (bufferSizeBytes < sizeof(FileInfo))
    {
        outHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        return false;
    }

    std::byte* const base = reinterpret_cast<std::byte*>(head);
    std::byte* const end  = base + bufferSizeBytes;

    for (FileInfo* entry = head; entry != nullptr;)
    {
        std::byte* const entryBytes = reinterpret_cast<std::byte*>(entry);
        if (entryBytes < base || entryBytes + sizeof(FileInfo) > end)
        {
            outHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return false;
        }

        const size_t availableBytes  = static_cast<size_t>(end - entryBytes);
        constexpr size_t kNameOffset = offsetof(FileInfo, FileName);
        if (entry->FileNameSize > (availableBytes - kNameOffset) || (entry->FileNameSize % sizeof(wchar_t)) != 0u)
        {
            outHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return false;
        }

        const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
        const std::wstring_view name(entry->FileName, nameChars);
        const std::wstring_view normalizedName = NormalizeEntryNameForCompare(name);

        const bool isDir = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (! ShouldIgnoreEntry(normalizedName, isDir, settings, ignoreFilePatterns, ignoreDirectoryPatterns))
        {
            SideEntry out{};
            out.name           = std::wstring(name);
            out.isDirectory    = isDir;
            out.fileAttributes = entry->FileAttributes;
            out.lastWriteTime  = entry->LastWriteTime;
            out.sizeBytes      = (! isDir && entry->EndOfFile > 0) ? static_cast<uint64_t>(entry->EndOfFile) : 0;
            InsertSideEntryPreservingNormalizedCollisions(outEntries, std::wstring(normalizedName), std::move(out));
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }

        if (entry->NextEntryOffset < sizeof(FileInfo))
        {
            outHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return false;
        }

        if (entry->NextEntryOffset > availableBytes)
        {
            outHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return false;
        }

        std::byte* const next = entryBytes + entry->NextEntryOffset;
        if (next < base || next + sizeof(FileInfo) > end)
        {
            outHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return false;
        }

        entry = reinterpret_cast<FileInfo*>(next);
    }

    outHr = S_OK;
    return true;
}

enum class FileContentCompareResult : uint8_t
{
    Equal,
    Different,
    Cancelled,
};

template <typename ProgressCallback>
[[nodiscard]] FileContentCompareResult CompareFileContent(IFileSystemIO* leftIo,
                                                          IFileSystemIO* rightIo,
                                                          const std::filesystem::path& leftPath,
                                                          const std::filesystem::path& rightPath,
                                                          const std::atomic_uint64_t* versionCounter,
                                                          uint64_t expectedVersion,
                                                          const std::atomic_uint64_t* cancelTokenCounter,
                                                          uint64_t expectedCancelToken,
                                                          std::stop_token stopToken,
                                                          ProgressCallback&& progress) noexcept
{
    const auto isCancelled = [&]() noexcept
    {
        if (stopToken.stop_requested())
        {
            return true;
        }

        if (versionCounter && versionCounter->load(std::memory_order_acquire) != expectedVersion)
        {
            return true;
        }

        return cancelTokenCounter && cancelTokenCounter->load(std::memory_order_acquire) != expectedCancelToken;
    };

    if (isCancelled())
    {
        return FileContentCompareResult::Cancelled;
    }

    if (! leftIo || ! rightIo)
    {
        return FileContentCompareResult::Cancelled;
    }

    // Allow the UI to reflect "starting" even if the underlying reader creation does network/disk work.
    progress(0u, 0u, true);

    wil::com_ptr<IFileReader> left;
    const HRESULT hrLeft = leftIo->CreateFileReader(leftPath.c_str(), left.put());
    if (FAILED(hrLeft) || ! left)
    {
        return FileContentCompareResult::Different;
    }

    wil::com_ptr<IFileReader> right;
    const HRESULT hrRight = rightIo->CreateFileReader(rightPath.c_str(), right.put());
    if (FAILED(hrRight) || ! right)
    {
        return FileContentCompareResult::Different;
    }

    if (isCancelled())
    {
        return FileContentCompareResult::Cancelled;
    }

    uint64_t leftSize         = 0;
    const bool leftSizeKnown  = SUCCEEDED(left->GetSize(&leftSize));
    uint64_t rightSize        = 0;
    const bool rightSizeKnown = SUCCEEDED(right->GetSize(&rightSize));

    const bool sizeKnown = leftSizeKnown && rightSizeKnown;
    if (sizeKnown)
    {
        if (leftSize != rightSize)
        {
            return FileContentCompareResult::Different;
        }

        if (leftSize == 0)
        {
            return FileContentCompareResult::Equal;
        }
    }

    progress(0, sizeKnown ? static_cast<uint64_t>(leftSize) : 0u, true);

    std::array<std::byte, 256 * 1024> leftBuf{};
    std::array<std::byte, 256 * 1024> rightBuf{};

    size_t leftPos  = 0;
    size_t leftHave = 0;
    bool leftEof    = false;

    size_t rightPos  = 0;
    size_t rightHave = 0;
    bool rightEof    = false;

    uint64_t completed             = 0;
    uint64_t lastReportedCompleted = 0;

    const uint64_t expectedTotalBytes = sizeKnown ? static_cast<uint64_t>(leftSize) : 0u;

    auto tryRead =
        [&](IFileReader* reader, std::array<std::byte, 256 * 1024>& buffer, size_t& pos, size_t& have, bool& eof, uint64_t maxBytesToRead) noexcept -> bool
    {
        if (! reader || eof)
        {
            return true;
        }

        if (pos != have)
        {
            return true;
        }

        pos  = 0;
        have = 0;

        const uint64_t want64 = std::min<uint64_t>(static_cast<uint64_t>(buffer.size()), maxBytesToRead);
        if (want64 == 0u)
        {
            return true;
        }

        const DWORD want = static_cast<DWORD>(std::min<uint64_t>(want64, 0xFFFF'FFFFull));

        unsigned long read = 0;
        const HRESULT hr   = reader->Read(buffer.data(), want, &read);
        if (FAILED(hr))
        {
            return false;
        }

        if (read == 0u)
        {
            eof = true;
            return true;
        }

        have = static_cast<size_t>(read);
        return true;
    };

    for (;;)
    {
        if (isCancelled())
        {
            return FileContentCompareResult::Cancelled;
        }

        if (sizeKnown && completed >= expectedTotalBytes)
        {
            break;
        }

        const uint64_t remaining = sizeKnown ? (expectedTotalBytes - completed) : static_cast<uint64_t>((std::numeric_limits<size_t>::max)());

        if (! tryRead(left.get(), leftBuf, leftPos, leftHave, leftEof, remaining))
        {
            return FileContentCompareResult::Different;
        }
        if (! tryRead(right.get(), rightBuf, rightPos, rightHave, rightEof, remaining))
        {
            return FileContentCompareResult::Different;
        }

        const size_t leftAvailable  = leftHave - leftPos;
        const size_t rightAvailable = rightHave - rightPos;

        if (leftAvailable == 0u || rightAvailable == 0u)
        {
            if (! sizeKnown)
            {
                if (leftAvailable == 0u && leftEof && rightAvailable == 0u && rightEof)
                {
                    progress(completed, 0u, true);
                    return FileContentCompareResult::Equal;
                }

                if ((leftAvailable == 0u && leftEof && rightAvailable > 0u) || (rightAvailable == 0u && rightEof && leftAvailable > 0u))
                {
                    return FileContentCompareResult::Different;
                }

                continue;
            }

            return FileContentCompareResult::Different;
        }

        size_t toCompare = std::min(leftAvailable, rightAvailable);
        if (sizeKnown)
        {
            const uint64_t remainingBytes = expectedTotalBytes - completed;
            const size_t remainingSizeT   = remainingBytes > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()) ? (std::numeric_limits<size_t>::max)()
                                                                                                                         : static_cast<size_t>(remainingBytes);
            toCompare                     = std::min(toCompare, remainingSizeT);
        }

        if (toCompare == 0u)
        {
            continue;
        }

        if (std::memcmp(leftBuf.data() + leftPos, rightBuf.data() + rightPos, toCompare) != 0)
        {
            return FileContentCompareResult::Different;
        }

        leftPos += toCompare;
        rightPos += toCompare;
        completed += static_cast<uint64_t>(toCompare);
        if ((completed - lastReportedCompleted) >= (64u * 1024u))
        {
            lastReportedCompleted = completed;
            progress(completed, sizeKnown ? expectedTotalBytes : 0u, false);
        }

        if (leftPos == leftHave)
        {
            leftPos  = 0;
            leftHave = 0;
        }
        if (rightPos == rightHave)
        {
            rightPos  = 0;
            rightHave = 0;
        }
    }

    if (leftPos != leftHave || rightPos != rightHave)
    {
        return FileContentCompareResult::Different;
    }

    unsigned long extraLeft = 0;
    if (FAILED(left->Read(leftBuf.data(), 1, &extraLeft)))
    {
        return FileContentCompareResult::Different;
    }

    unsigned long extraRight = 0;
    if (FAILED(right->Read(rightBuf.data(), 1, &extraRight)))
    {
        return FileContentCompareResult::Different;
    }

    if (extraLeft != 0u || extraRight != 0u)
    {
        return FileContentCompareResult::Different;
    }

    progress(expectedTotalBytes, expectedTotalBytes, true);
    return FileContentCompareResult::Equal;
}

class CompareFilesInformation final : public IFilesInformation
{
public:
    CompareFilesInformation(std::vector<unsigned char> buffer, std::vector<FileInfo*> entries) noexcept
        : _buffer(std::move(buffer)),
          _entries(std::move(entries))
    {
    }

    CompareFilesInformation(const CompareFilesInformation&)            = delete;
    CompareFilesInformation& operator=(const CompareFilesInformation&) = delete;
    CompareFilesInformation(CompareFilesInformation&&)                 = delete;
    CompareFilesInformation& operator=(CompareFilesInformation&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFilesInformation))
        {
            *ppvObject = static_cast<IFilesInformation*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept override
    {
        if (ppFileInfo == nullptr)
        {
            return E_POINTER;
        }

        if (_buffer.empty())
        {
            *ppFileInfo = nullptr;
            return S_OK;
        }

        *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = static_cast<unsigned long>(_buffer.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = static_cast<unsigned long>(_buffer.capacity());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept override
    {
        if (pCount == nullptr)
        {
            return E_POINTER;
        }

        *pCount = static_cast<unsigned long>(_entries.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept override
    {
        if (ppEntry == nullptr)
        {
            return E_POINTER;
        }

        *ppEntry = nullptr;
        if (index >= _entries.size())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_INDEX);
        }

        *ppEntry = _entries[index];
        return S_OK;
    }

private:
    ~CompareFilesInformation() = default;

    std::atomic_ulong _refCount{1};
    std::vector<unsigned char> _buffer;
    std::vector<FileInfo*> _entries;
};

struct OutEntry
{
    std::wstring name;
    DWORD fileAttributes  = 0;
    int64_t lastWriteTime = 0;
    uint64_t sizeBytes    = 0;
};

[[nodiscard]] size_t AlignedFileInfoSizeBytes(size_t nameChars) noexcept
{
    constexpr size_t kAlign = 8;
    const size_t raw        = offsetof(FileInfo, FileName) + (nameChars * sizeof(wchar_t));
    return (raw + (kAlign - 1)) & ~(kAlign - 1);
}

[[nodiscard]] wil::com_ptr<IFilesInformation> BuildFilesInformation(const std::vector<OutEntry>& entries, HRESULT& outHr) noexcept
{
    outHr = S_OK;

    if (entries.empty())
    {
        return wil::com_ptr<IFilesInformation>(new CompareFilesInformation({}, {}));
    }

    size_t totalBytes = 0;
    for (const auto& e : entries)
    {
        totalBytes += AlignedFileInfoSizeBytes(e.name.size());
    }

    if (totalBytes > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
    {
        outHr = E_OUTOFMEMORY;
        return nullptr;
    }

    std::vector<unsigned char> buffer(totalBytes);
    std::vector<FileInfo*> entryPtrs;
    entryPtrs.reserve(entries.size());

    size_t offset = 0;
    for (size_t i = 0; i < entries.size(); ++i)
    {
        const OutEntry& src    = entries[i];
        const size_t entrySize = AlignedFileInfoSizeBytes(src.name.size());
        if (offset + entrySize > buffer.size())
        {
            outHr = E_FAIL;
            return nullptr;
        }

        FileInfo* dst = reinterpret_cast<FileInfo*>(buffer.data() + offset);
        entryPtrs.push_back(dst);

        dst->NextEntryOffset = (i + 1 < entries.size()) ? static_cast<unsigned long>(entrySize) : 0;
        dst->FileIndex       = 0;
        dst->CreationTime    = 0;
        dst->LastAccessTime  = 0;
        dst->LastWriteTime   = src.lastWriteTime;
        dst->ChangeTime      = 0;
        dst->EndOfFile       = static_cast<__int64>(src.sizeBytes);
        dst->AllocationSize  = 0;
        dst->FileAttributes  = src.fileAttributes;
        dst->FileNameSize    = static_cast<unsigned long>(src.name.size() * sizeof(wchar_t));
        dst->EaSize          = 0;

        if (! src.name.empty())
        {
            std::memcpy(dst->FileName, src.name.data(), dst->FileNameSize);
        }

        offset += entrySize;
    }

    return wil::com_ptr<IFilesInformation>(new CompareFilesInformation(std::move(buffer), std::move(entryPtrs)));
}
} // namespace

std::shared_ptr<CompareDirectoriesFolderDecision> CompareDirectoriesSession::ComputeDecisionForFolder(
    const std::filesystem::path& relativeFolder,
    const Common::Settings::CompareDirectoriesSettings& settings,
    const std::vector<std::wstring>& ignoreFilePatterns,
    const std::vector<std::wstring>& ignoreDirectoryPatterns,
    uint64_t version,
    uint64_t cancelToken,
    bool allowBackgroundWork,
    bool reportScanProgress,
    bool forceNotifyFolderStart,
    ScanPriority scanPriority,
    std::stop_token stopToken) noexcept
{
    const std::filesystem::path leftFolder  = ResolveAbsolute(ComparePane::Left, relativeFolder);
    const std::filesystem::path rightFolder = ResolveAbsolute(ComparePane::Right, relativeFolder);

    struct ContentCompareActivation
    {
        std::filesystem::path relativeFolder;
        std::wstring entryName;
        uint64_t totalBytes = 0;
    };

    const auto isCancelled = [&]() noexcept -> bool
    {
        if (stopToken.stop_requested())
        {
            return true;
        }

        if (_version.load(std::memory_order_acquire) != version)
        {
            return true;
        }

        if (_backgroundWorkCancelToken.load(std::memory_order_acquire) != cancelToken)
        {
            return true;
        }

        if (allowBackgroundWork && ! _backgroundWorkEnabled.load(std::memory_order_acquire))
        {
            return true;
        }

        return false;
    };

    if (reportScanProgress)
    {
        static_cast<void>(_scanFoldersScanned.fetch_add(1u, std::memory_order_acq_rel));
        NotifyScanProgress(relativeFolder, {}, forceNotifyFolderStart);
    }

    std::optional<ContentCompareActivation> contentActivated;

    auto decision     = std::make_shared<CompareDirectoriesFolderDecision>();
    decision->version = version;
    decision->hr      = S_OK;

    if (isCancelled())
    {
        return decision;
    }

    if (ShouldIgnoreRelativeFolder(relativeFolder, settings, ignoreDirectoryPatterns))
    {
        return decision;
    }

    std::map<std::wstring, SideEntry, WStringViewNoCaseLess> leftEntries;
    std::map<std::wstring, SideEntry, WStringViewNoCaseLess> rightEntries;

    bool leftMissing = false;
    HRESULT leftHr   = S_OK;
    if (! TryReadDirectoryEntries(_leftFileSystem, leftFolder, settings, ignoreFilePatterns, ignoreDirectoryPatterns, leftEntries, leftMissing, leftHr))
    {
        decision->hr = leftHr;
    }
    decision->leftFolderMissing = leftMissing;

    if (isCancelled())
    {
        return decision;
    }

    bool rightMissing = false;
    HRESULT rightHr   = S_OK;
    if (SUCCEEDED(decision->hr) &&
        ! TryReadDirectoryEntries(_rightFileSystem, rightFolder, settings, ignoreFilePatterns, ignoreDirectoryPatterns, rightEntries, rightMissing, rightHr))
    {
        decision->hr = rightHr;
    }
    decision->rightFolderMissing = rightMissing;

    if (SUCCEEDED(decision->hr))
    {
        // Seed with left entries first (preserves left casing as key when both exist).
        for (const auto& [name, entry] : leftEntries)
        {
            CompareDirectoriesItemDecision item{};
            item.existsLeft         = true;
            item.isDirectory        = entry.isDirectory;
            item.leftSizeBytes      = entry.sizeBytes;
            item.leftLastWriteTime  = entry.lastWriteTime;
            item.leftFileAttributes = entry.fileAttributes;
            decision->items.emplace(name, item);
        }

        for (const auto& [name, entry] : rightEntries)
        {
            auto it = decision->items.find(name);
            if (it == decision->items.end())
            {
                CompareDirectoriesItemDecision item{};
                item.existsRight         = true;
                item.isDirectory         = entry.isDirectory;
                item.rightSizeBytes      = entry.sizeBytes;
                item.rightLastWriteTime  = entry.lastWriteTime;
                item.rightFileAttributes = entry.fileAttributes;
                decision->items.emplace(name, item);
                continue;
            }

            CompareDirectoriesItemDecision& item = it->second;
            item.existsRight                     = true;
            item.isDirectory                     = item.isDirectory || entry.isDirectory;
            item.rightSizeBytes                  = entry.sizeBytes;
            item.rightLastWriteTime              = entry.lastWriteTime;
            item.rightFileAttributes             = entry.fileAttributes;
        }

        for (auto& [name, item] : decision->items)
        {
            if (reportScanProgress)
            {
                const uint64_t scannedEntries = _scanEntriesScanned.fetch_add(1u, std::memory_order_acq_rel) + 1u;
                if ((scannedEntries & 0x3Fu) == 0u)
                {
                    NotifyScanProgress(relativeFolder, name, false);
                    if (isCancelled())
                    {
                        break;
                    }
                }
            }
            else if (isCancelled())
            {
                break;
            }

            item.isDifferent    = false;
            item.selectLeft     = false;
            item.selectRight    = false;
            item.differenceMask = 0;

            if (item.existsLeft != item.existsRight)
            {
                item.isDifferent = true;
                if (item.existsLeft)
                {
                    item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::OnlyInLeft);
                    item.selectLeft = ! item.isDirectory || settings.selectSubdirsOnlyInOnePane;
                }
                if (item.existsRight)
                {
                    item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::OnlyInRight);
                    item.selectRight = ! item.isDirectory || settings.selectSubdirsOnlyInOnePane;
                }
            }
            else if (item.existsLeft && item.existsRight)
            {
                const bool leftIsDir  = (item.leftFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                const bool rightIsDir = (item.rightFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

                if (leftIsDir != rightIsDir)
                {
                    item.isDifferent = true;
                    item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::TypeMismatch);
                    item.selectLeft  = true;
                    item.selectRight = true;
                }
                else if (leftIsDir)
                {
                    bool attrsDifferent = false;
                    if (settings.compareSubdirectoryAttributes && item.leftFileAttributes != item.rightFileAttributes)
                    {
                        attrsDifferent = true;
                        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirAttributes);
                    }

                    if (attrsDifferent)
                    {
                        item.isDifferent = true;
                        item.selectLeft  = true;
                        item.selectRight = true;
                    }
                }
                else
                {
                    const bool sizeDifferent       = item.leftSizeBytes != item.rightSizeBytes;
                    const bool timeDifferent       = item.leftLastWriteTime != item.rightLastWriteTime;
                    const bool attrsDifferent      = item.leftFileAttributes != item.rightFileAttributes;
                    const bool nonContentDifferent = (settings.compareSize && sizeDifferent) || (settings.compareDateTime && timeDifferent) ||
                                                     (settings.compareAttributes && attrsDifferent);

                    bool contentDifferent        = false;
                    bool contentPending          = false;
                    const bool canCompareContent = settings.compareContent && IsContentCompareSupported();
                    if (canCompareContent && ! nonContentDifferent)
                    {
                        if (sizeDifferent)
                        {
                            contentDifferent = true;
                        }
                        else
                        {
                            const std::filesystem::path relativeFile = (relativeFolder / std::filesystem::path(name)).lexically_normal();

                            ContentCompareKey compareKey{};
                            compareKey.relativeFileKey    = MakeCacheKey(relativeFile);
                            compareKey.leftSizeBytes      = item.leftSizeBytes;
                            compareKey.rightSizeBytes     = item.rightSizeBytes;
                            compareKey.leftLastWriteTime  = item.leftLastWriteTime;
                            compareKey.rightLastWriteTime = item.rightLastWriteTime;

                            std::optional<bool> cachedEqual;
                            {
                                std::unique_lock lock(_mutex);
                                if (const auto it = _contentCompareCache.find(compareKey); it != _contentCompareCache.end())
                                {
                                    cachedEqual = it->second;
                                }
                                else if (allowBackgroundWork)
                                {
                                    const auto inflightIt    = _contentCompareInFlight.find(compareKey);
                                    const bool alreadyQueued = (inflightIt != _contentCompareInFlight.end() && inflightIt->second.version == version &&
                                                                inflightIt->second.cancelToken == cancelToken);
                                    if (! alreadyQueued)
                                    {
                                        EnsureContentCompareWorkersLocked();

                                        constexpr size_t kContentCompareQueueMaxHighJobs = 128;
                                        constexpr size_t kContentCompareQueueMaxLowJobs  = 896;

                                        const bool high      = scanPriority == ScanPriority::High;
                                        auto& targetQueue    = high ? _contentCompareQueueHigh : _contentCompareQueueLow;
                                        const size_t maxJobs = high ? kContentCompareQueueMaxHighJobs : kContentCompareQueueMaxLowJobs;

                                        while (targetQueue.size() >= maxJobs && ! isCancelled())
                                        {
                                            _contentCompareQueueNotFullCv.wait(lock, [&]() noexcept { return isCancelled() || targetQueue.size() < maxJobs; });
                                        }

                                        if (! isCancelled() && targetQueue.size() < maxJobs)
                                        {
                                            InFlightContentStamp inFlight{};
                                            inFlight.version                    = version;
                                            inFlight.cancelToken                = cancelToken;
                                            _contentCompareInFlight[compareKey] = inFlight;
                                            _contentInFlightHighWater           = std::max(_contentInFlightHighWater, _contentCompareInFlight.size());

                                            static_cast<void>(_contentCompareTotalCompares.fetch_add(1u, std::memory_order_acq_rel));
                                            static_cast<void>(_contentCompareTotalBytes.fetch_add(item.leftSizeBytes, std::memory_order_acq_rel));

                                            const uint64_t pendingBefore = _contentComparePendingCompares.fetch_add(1u, std::memory_order_acq_rel);
                                            if (pendingBefore == 0u && ! contentActivated.has_value())
                                            {
                                                ContentCompareActivation activation{};
                                                activation.relativeFolder = relativeFolder;
                                                activation.entryName      = name;
                                                activation.totalBytes     = item.leftSizeBytes;
                                                contentActivated.emplace(std::move(activation));
                                            }

                                            ContentCompareJob job{};
                                            job.version             = version;
                                            job.cancelToken         = cancelToken;
                                            job.relativeFolder      = relativeFolder;
                                            job.entryName           = name;
                                            job.key                 = std::move(compareKey);
                                            job.priority            = scanPriority;
                                            job.leftFileAttributes  = item.leftFileAttributes;
                                            job.rightFileAttributes = item.rightFileAttributes;

                                            targetQueue.emplace_back(std::move(job));
                                            _contentQueueHighWater =
                                                std::max(_contentQueueHighWater, _contentCompareQueueHigh.size() + _contentCompareQueueLow.size());
                                            _contentQueueHighHighWater = std::max(_contentQueueHighHighWater, _contentCompareQueueHigh.size());
                                            _contentQueueLowHighWater  = std::max(_contentQueueLowHighWater, _contentCompareQueueLow.size());

                                            _contentCompareCv.notify_one();
                                        }
                                    }
                                }
                            }

                            if (contentActivated.has_value())
                            {
                                NotifyContentProgress(std::numeric_limits<uint32_t>::max(),
                                                      contentActivated->relativeFolder,
                                                      contentActivated->entryName,
                                                      contentActivated->totalBytes,
                                                      0);
                                contentActivated.reset();
                            }

                            if (cachedEqual.has_value())
                            {
                                contentDifferent = ! cachedEqual.value();
                            }
                            else if (allowBackgroundWork && ! isCancelled())
                            {
                                contentPending = true;
                            }
                        }
                    }

                    ApplyCriteriaDiffAndSelection(item, settings, canCompareContent, contentDifferent);

                    if (canCompareContent && contentPending)
                    {
                        item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::ContentPending);
                    }
                }
            }
        }
    }

    if (settings.compareSubdirectories && SUCCEEDED(decision->hr))
    {
        for (auto& [name, item] : decision->items)
        {
            if (isCancelled())
            {
                break;
            }

            if (! item.existsLeft || ! item.existsRight)
            {
                continue;
            }

            const bool leftIsDir  = (item.leftFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            const bool rightIsDir = (item.rightFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            if (! leftIsDir || ! rightIsDir)
            {
                continue;
            }

            // Avoid following directory reparse points (symlinks/junctions).
            if (IsReparsePairEntry(item))
            {
                continue;
            }

            const std::filesystem::path childRel = relativeFolder / std::filesystem::path(name);
            const std::wstring childKey          = MakeCacheKey(childRel);

            std::shared_ptr<const CompareDirectoriesFolderDecision> childDecision;
            {
                std::lock_guard guard(_mutex);
                const auto it = _cache.find(childKey);
                if (it != _cache.end() && it->second && it->second->version == version)
                {
                    childDecision = it->second;
                }
            }

            if (! childDecision)
            {
                if (allowBackgroundWork)
                {
                    item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirPending);
                }
                continue;
            }

            const bool childPending = SUCCEEDED(childDecision->hr) && childDecision->anyPending;
            if (allowBackgroundWork && childPending)
            {
                item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirPending);
            }

            const bool childDifferent = FAILED(childDecision->hr) || childDecision->anyDifferent;
            if (childDifferent)
            {
                item.differenceMask |= static_cast<uint32_t>(CompareDirectoriesDiffBit::SubdirContent);
                item.isDifferent = true;
                item.selectLeft  = true;
                item.selectRight = true;
            }
        }
    }

    if (! settings.keepIdenticalItems && SUCCEEDED(decision->hr))
    {
        for (auto it = decision->items.begin(); it != decision->items.end();)
        {
            const uint32_t diffMask   = it->second.differenceMask;
            const bool pendingSubdir  = HasFlag(diffMask, CompareDirectoriesDiffBit::SubdirPending);
            const bool pendingContent = HasFlag(diffMask, CompareDirectoriesDiffBit::ContentPending);
            const bool anyPendingItem = pendingSubdir || pendingContent;

            if (! it->second.isDifferent && ! pendingSubdir)
            {
                // In differences-only mode, file-level ContentPending entries can explode memory
                // for large folders. Keep the compare running, but elide the per-item entries.
                if (pendingContent)
                {
                    ++decision->pendingContentCompareCount;
                    it = decision->items.erase(it);
                    continue;
                }

                // Drop identical entries to keep cached decisions bounded.
                if (! anyPendingItem && diffMask == 0u)
                {
                    it = decision->items.erase(it);
                    continue;
                }
            }

            ++it;
        }
    }

    decision->anyDifferent = AnyChildDifferent(*decision);
    decision->anyPending   = AnyChildPending(*decision);
    return decision;
}

#ifdef ENABLE_TESTS
std::shared_ptr<const CompareDirectoriesFolderDecision> CompareDirectoriesSession::GetOrComputeDecision(const std::filesystem::path& relativeFolder)
{
    const std::wstring rootKey     = MakeCacheKey(relativeFolder);
    const bool allowBackgroundWork = _backgroundWorkEnabled.load(std::memory_order_acquire);
    const uint64_t cancelToken     = _backgroundWorkCancelToken.load(std::memory_order_acquire);

    Common::Settings::CompareDirectoriesSettings settings{};
    std::shared_ptr<const std::vector<std::wstring>> ignoreFilePatterns;
    std::shared_ptr<const std::vector<std::wstring>> ignoreDirectoryPatterns;

    uint64_t version = 0;
    {
        std::lock_guard guard(_mutex);
        settings                = _settings;
        ignoreFilePatterns      = _ignoreFilePatterns;
        ignoreDirectoryPatterns = _ignoreDirectoryPatterns;
        version                 = _version.load(std::memory_order_relaxed);

        ApplyPendingContentCompareUpdatesLocked(rootKey);

        const auto it = _cache.find(rootKey);
        if (it != _cache.end() && it->second && it->second->version == version)
        {
            if (IsCacheableDecision(it->second))
            {
                return it->second;
            }

            TrackDecisionCacheEraseLocked(rootKey);
            _cache.erase(it);
        }
    }

    if (! ignoreFilePatterns)
    {
        ignoreFilePatterns = std::make_shared<const std::vector<std::wstring>>();
    }
    if (! ignoreDirectoryPatterns)
    {
        ignoreDirectoryPatterns = std::make_shared<const std::vector<std::wstring>>();
    }

    constexpr bool reportScanProgress                                = false;
    const std::shared_ptr<CompareDirectoriesFolderDecision> decision = ComputeDecisionForFolder(relativeFolder,
                                                                                                settings,
                                                                                                *ignoreFilePatterns,
                                                                                                *ignoreDirectoryPatterns,
                                                                                                version,
                                                                                                cancelToken,
                                                                                                allowBackgroundWork,
                                                                                                reportScanProgress,
                                                                                                false,
                                                                                                ScanPriority::Low,
                                                                                                std::stop_token{});

    const std::shared_ptr<const CompareDirectoriesFolderDecision> finalDecision = decision;
    {
        std::lock_guard guard(_mutex);
        if (_version.load(std::memory_order_relaxed) == version && IsCacheableDecision(finalDecision))
        {
            _cache[rootKey]                = finalDecision;
            _decisionCacheEntriesHighWater = std::max(_decisionCacheEntriesHighWater, _cache.size());
            TrackDecisionCacheInsertOrUpdateLocked(rootKey, finalDecision);
            ApplyPendingContentCompareUpdatesLocked(rootKey);
            if (const auto updatedIt = _cache.find(rootKey); updatedIt != _cache.end() && updatedIt->second && updatedIt->second->version == version)
            {
                return updatedIt->second;
            }
            MaybeEvictDecisionCacheLocked();
        }
    }

    return finalDecision ? finalDecision : std::make_shared<CompareDirectoriesFolderDecision>();
}
#endif

void CompareDirectoriesSession::ScanWorker(std::stop_token stopToken, uint32_t workerIndex) noexcept
{
    UNREFERENCED_PARAMETER(workerIndex);

    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(coinitHr))
    {
        Debug::Error(L"CompareDirectories scan worker: CoInitializeEx(COINIT_MULTITHREADED) failed: 0x{:08X}", coinitHr);
        FAIL_FAST_IF_FAILED(coinitHr);
    }
    [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

    if (SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_BELOW_NORMAL) == 0)
    {
        Debug::ErrorWithLastError(L"CompareDirectories scan worker: SetThreadPriority(BELOW_NORMAL) failed");
    }

    while (! stopToken.stop_requested())
    {
        FolderScanJob job{};
        {
            std::unique_lock lock(_mutex);
            _scanCv.wait(lock, [&]() { return stopToken.stop_requested() || ! _scanQueueHigh.empty() || ! _scanQueueLow.empty(); });

            if (stopToken.stop_requested())
            {
                break;
            }

            if (! _scanQueueHigh.empty())
            {
                job = std::move(_scanQueueHigh.front());
                _scanQueueHigh.pop_front();
                _scanHighQueuedKeys.erase(job.key);
            }
            else
            {
                job = std::move(_scanQueueLow.front());
                _scanQueueLow.pop_front();
            }

            // Stale duplicate (upgrade) jobs can remain in the queues after the scheduled work is completed.
            if (! _scanScheduledKeys.contains(job.key))
            {
                continue;
            }

            InFlightScanStamp inFlight{};
            inFlight.version     = job.version;
            inFlight.cancelToken = job.cancelToken;
            const bool inserted  = _scanInFlightKeys.emplace(job.key, inFlight).second;
            if (! inserted)
            {
                continue;
            }
            _scanInFlightHighWater = std::max(_scanInFlightHighWater, _scanInFlightKeys.size());
        }

        auto completeJob = wil::scope_exit([&]
        {
            bool notifyIdle = false;
            {
                std::lock_guard guard(_mutex);
                bool erased = false;
                if (const auto it = _scanInFlightKeys.find(job.key);
                    it != _scanInFlightKeys.end() && it->second.version == job.version && it->second.cancelToken == job.cancelToken)
                {
                    _scanInFlightKeys.erase(it);
                    erased = true;
                }

                if (erased)
                {
                    _scanScheduledKeys.erase(job.key);
                    const uint32_t activeAfter = _scanActiveScans.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
                    notifyIdle                 = activeAfter == 0u;
                }
            }

            if (notifyIdle)
            {
                NotifyScanProgress({}, {}, true);
                NotifyDecisionUpdated(true);
            }
        });

        const uint64_t currentVersion = _version.load(std::memory_order_acquire);
        if (currentVersion != job.version)
        {
            continue;
        }

        const uint64_t currentCancelToken = _backgroundWorkCancelToken.load(std::memory_order_acquire);
        if (! _backgroundWorkEnabled.load(std::memory_order_acquire) || currentCancelToken != job.cancelToken)
        {
            continue;
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        std::shared_ptr<const std::vector<std::wstring>> ignoreFilePatterns;
        std::shared_ptr<const std::vector<std::wstring>> ignoreDirectoryPatterns;
        {
            std::lock_guard guard(_mutex);
            settings                = _settings;
            ignoreFilePatterns      = _ignoreFilePatterns;
            ignoreDirectoryPatterns = _ignoreDirectoryPatterns;
        }

        if (! ignoreFilePatterns)
        {
            ignoreFilePatterns = std::make_shared<const std::vector<std::wstring>>();
        }
        if (! ignoreDirectoryPatterns)
        {
            ignoreDirectoryPatterns = std::make_shared<const std::vector<std::wstring>>();
        }

        std::shared_ptr<const CompareDirectoriesFolderDecision> decisionConst;
        {
            std::lock_guard guard(_mutex);
            const auto it = _cache.find(job.key);
            if (it != _cache.end() && it->second && it->second->version == job.version)
            {
                if (IsCacheableDecision(it->second))
                {
                    decisionConst = it->second;
                }
                else
                {
                    TrackDecisionCacheEraseLocked(job.key);
                    _cache.erase(it);
                }
            }
        }

        std::shared_ptr<CompareDirectoriesFolderDecision> computedDecision;
        if (! decisionConst)
        {
            computedDecision = ComputeDecisionForFolder(job.relativeFolder,
                                                        settings,
                                                        *ignoreFilePatterns,
                                                        *ignoreDirectoryPatterns,
                                                        job.version,
                                                        job.cancelToken,
                                                        true,
                                                        true,
                                                        job.key == L".",
                                                        job.priority,
                                                        stopToken);
            decisionConst    = computedDecision;

            {
                std::lock_guard guard(_mutex);
                const auto inFlightIt = _scanInFlightKeys.find(job.key);
                const bool scanStillCurrent =
                    inFlightIt != _scanInFlightKeys.end() && inFlightIt->second.version == job.version && inFlightIt->second.cancelToken == job.cancelToken;
                if (_version.load(std::memory_order_relaxed) == job.version && scanStillCurrent && IsCacheableDecision(computedDecision))
                {
                    const std::shared_ptr<const CompareDirectoriesFolderDecision> storedDecision = computedDecision;
                    _cache[job.key]                                                              = storedDecision;
                    _decisionCacheEntriesHighWater                                               = std::max(_decisionCacheEntriesHighWater, _cache.size());
                    TrackDecisionCacheInsertOrUpdateLocked(job.key, storedDecision);
                    ApplyPendingContentCompareUpdatesLocked(job.key);
                    if (const auto updatedIt = _cache.find(job.key);
                        updatedIt != _cache.end() && updatedIt->second && updatedIt->second->version == job.version)
                    {
                        decisionConst = updatedIt->second;
                    }
                    MaybeEvictDecisionCacheLocked();
                }
            }
        }
        else
        {
            static_cast<void>(_scanFoldersScanned.fetch_add(1u, std::memory_order_acq_rel));
            NotifyScanProgress(job.relativeFolder, {}, job.key == L".");
        }

        if (! decisionConst)
        {
            continue;
        }

        if (settings.compareSubdirectories && SUCCEEDED(decisionConst->hr))
        {
            std::vector<std::filesystem::path> childFolders;
            childFolders.reserve(decisionConst->items.size());

            for (const auto& [name, item] : decisionConst->items)
            {
                if (! item.existsLeft || ! item.existsRight)
                {
                    continue;
                }

                const bool leftIsDir  = (item.leftFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                const bool rightIsDir = (item.rightFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (! leftIsDir || ! rightIsDir)
                {
                    continue;
                }

                // Avoid following directory reparse points (symlinks/junctions).
                if (IsReparsePairEntry(item))
                {
                    continue;
                }

                childFolders.emplace_back(job.relativeFolder / std::filesystem::path(name));
            }

            if (! childFolders.empty())
            {
                std::lock_guard guard(_mutex);
                const uint64_t nowVersion = _version.load(std::memory_order_relaxed);
                const auto inFlightIt     = _scanInFlightKeys.find(job.key);
                const bool scanStillCurrent =
                    inFlightIt != _scanInFlightKeys.end() && inFlightIt->second.version == job.version && inFlightIt->second.cancelToken == job.cancelToken;
                if (nowVersion == job.version && _backgroundWorkEnabled.load(std::memory_order_relaxed) &&
                    _backgroundWorkCancelToken.load(std::memory_order_relaxed) == job.cancelToken && scanStillCurrent)
                {
                    for (const auto& childRel : childFolders)
                    {
                        EnqueueScanLocked(childRel, job.version, job.cancelToken, ScanPriority::Low);
                    }
                }
            }
        }

        if (settings.compareSubdirectories && job.key != L".")
        {
            std::lock_guard guard(_mutex);
            const auto inFlightIt = _scanInFlightKeys.find(job.key);
            const bool scanStillCurrent =
                inFlightIt != _scanInFlightKeys.end() && inFlightIt->second.version == job.version && inFlightIt->second.cancelToken == job.cancelToken;
            if (_version.load(std::memory_order_relaxed) == job.version && scanStillCurrent)
            {
                PendingSubdirAggregateUpdate update{};
                update.version      = job.version;
                update.hr           = decisionConst->hr;
                update.anyPending   = SUCCEEDED(decisionConst->hr) && decisionConst->anyPending;
                update.anyDifferent = SUCCEEDED(decisionConst->hr) && decisionConst->anyDifferent;

                _pendingSubdirUpdates.insert(job.key);
                _pendingSubdirAggregateUpdates[job.key] = update;
                _pendingSubdirAggregateRetryAttempts.erase(job.key);
                _pendingSubdirHighWater = std::max(_pendingSubdirHighWater, _pendingSubdirUpdates.size());
            }
        }

        NotifyDecisionUpdated(false);
    }
}

void CompareDirectoriesSession::ContentCompareWorker(std::stop_token stopToken, uint32_t workerIndex) noexcept
{
    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(coinitHr))
    {
        Debug::Error(L"CompareDirectories worker: CoInitializeEx(COINIT_MULTITHREADED) failed: 0x{:08X}", coinitHr);
        FAIL_FAST_IF_FAILED(coinitHr);
    }
    [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

    if (SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_BELOW_NORMAL) == 0)
    {
        Debug::ErrorWithLastError(L"CompareDirectories worker: SetThreadPriority(BELOW_NORMAL) failed");
    }

    uint64_t lastProgressNotifyTickMs = 0;
    uint32_t highBurstCount           = 0;

    while (! stopToken.stop_requested())
    {
        ContentCompareJob job{};
        {
            std::unique_lock lock(_mutex);
            _contentCompareCv.wait(lock,
                                   [&]() { return stopToken.stop_requested() || ! _contentCompareQueueHigh.empty() || ! _contentCompareQueueLow.empty(); });

            if (stopToken.stop_requested())
            {
                break;
            }

            const bool haveHigh = ! _contentCompareQueueHigh.empty();
            const bool haveLow  = ! _contentCompareQueueLow.empty();

            // Visible-first: prefer high-priority work, but allow low-priority progress under sustained high load.
            constexpr uint32_t kHighBurstBeforeLow = 8;
            const bool shouldTakeLow               = haveLow && (! haveHigh || highBurstCount >= kHighBurstBeforeLow);

            if (shouldTakeLow)
            {
                job = std::move(_contentCompareQueueLow.front());
                _contentCompareQueueLow.pop_front();
                highBurstCount = 0;
            }
            else
            {
                job = std::move(_contentCompareQueueHigh.front());
                _contentCompareQueueHigh.pop_front();
                ++highBurstCount;
            }
        }

        _contentCompareQueueNotFullCv.notify_all();

        const uint64_t currentVersion = _version.load(std::memory_order_acquire);
        if (currentVersion != job.version)
        {
            bool notifyIdle = false;
            {
                std::lock_guard guard(_mutex);
                bool erased = false;
                if (const auto it = _contentCompareInFlight.find(job.key);
                    it != _contentCompareInFlight.end() && it->second.version == job.version && it->second.cancelToken == job.cancelToken)
                {
                    _contentCompareInFlight.erase(it);
                    erased = true;
                }
                if (erased)
                {
                    const uint64_t pendingAfter = _contentComparePendingCompares.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
                    notifyIdle                  = pendingAfter == 0u;
                }
            }
            if (notifyIdle)
            {
                NotifyContentProgress(workerIndex, {}, {}, 0, 0);
            }
            continue;
        }

        const uint64_t currentCancelToken = _backgroundWorkCancelToken.load(std::memory_order_acquire);
        if (! _backgroundWorkEnabled.load(std::memory_order_acquire) || currentCancelToken != job.cancelToken)
        {
            bool notifyIdle = false;
            {
                std::lock_guard guard(_mutex);
                bool erased = false;
                if (const auto it = _contentCompareInFlight.find(job.key);
                    it != _contentCompareInFlight.end() && it->second.version == job.version && it->second.cancelToken == job.cancelToken)
                {
                    _contentCompareInFlight.erase(it);
                    erased = true;
                }
                if (erased)
                {
                    const uint64_t pendingAfter = _contentComparePendingCompares.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
                    notifyIdle                  = pendingAfter == 0u;
                }
            }
            if (notifyIdle)
            {
                NotifyContentProgress(workerIndex, {}, {}, 0, 0);
            }
            continue;
        }

        const auto progress = [&](uint64_t completedBytes, uint64_t totalBytes, bool force) noexcept
        {
            if (! force)
            {
                const uint64_t now = GetTickCount64();
                if (now >= lastProgressNotifyTickMs && (now - lastProgressNotifyTickMs) < 80u)
                {
                    return;
                }
                lastProgressNotifyTickMs = now;
            }
            NotifyContentProgress(workerIndex, job.relativeFolder, job.entryName, totalBytes, completedBytes);
        };

        const std::filesystem::path relativeFile = (job.relativeFolder / std::filesystem::path(job.entryName)).lexically_normal();
        const std::filesystem::path leftPath     = ResolveAbsolute(ComparePane::Left, relativeFile);
        const std::filesystem::path rightPath    = ResolveAbsolute(ComparePane::Right, relativeFile);

        const FileContentCompareResult compareResult = CompareFileContent(_leftFileSystemIo.get(),
                                                                          _rightFileSystemIo.get(),
                                                                          leftPath,
                                                                          rightPath,
                                                                          &_version,
                                                                          job.version,
                                                                          &_backgroundWorkCancelToken,
                                                                          job.cancelToken,
                                                                          stopToken,
                                                                          progress);
        if (compareResult == FileContentCompareResult::Cancelled)
        {
            bool notifyIdle = false;
            {
                std::lock_guard guard(_mutex);
                bool erased = false;
                if (const auto it = _contentCompareInFlight.find(job.key);
                    it != _contentCompareInFlight.end() && it->second.version == job.version && it->second.cancelToken == job.cancelToken)
                {
                    _contentCompareInFlight.erase(it);
                    erased = true;
                }
                if (erased)
                {
                    const uint64_t pendingAfter = _contentComparePendingCompares.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
                    notifyIdle                  = pendingAfter == 0u;
                }
            }
            if (notifyIdle)
            {
                NotifyContentProgress(workerIndex, {}, {}, 0, 0);
            }
            continue;
        }

        const bool areEqual = compareResult == FileContentCompareResult::Equal;

        bool shouldNotify         = false;
        bool forceNotifyFinal     = false;
        bool erased               = false;
        uint64_t pendingAfter     = 0;
        bool pendingAfterComputed = false;
        {
            std::lock_guard guard(_mutex);

            if (const auto it = _contentCompareInFlight.find(job.key);
                it != _contentCompareInFlight.end() && it->second.version == job.version && it->second.cancelToken == job.cancelToken)
            {
                _contentCompareInFlight.erase(it);
                erased = true;
            }

            if (erased)
            {
                // Only keep results from the currently active run; stale completions must not warm the cache.
                constexpr size_t kContentCacheMaxEntries = 16384;
                constexpr size_t kContentCacheEvictBatch = 4096;
                if (_contentCompareCache.size() >= kContentCacheMaxEntries)
                {
                    size_t evicted = 0;
                    for (auto it = _contentCompareCache.begin(); it != _contentCompareCache.end() && evicted < kContentCacheEvictBatch;)
                    {
                        it = _contentCompareCache.erase(it);
                        ++evicted;
                    }
                }
                static_cast<void>(_contentCompareCache.emplace(job.key, areEqual));
                _contentCacheHighWater = std::max(_contentCacheHighWater, _contentCompareCache.size());

                const std::wstring folderKey = MakeCacheKey(job.relativeFolder);

                PendingContentCompareUpdate update{};
                update.version             = job.version;
                update.leftSizeBytes       = job.key.leftSizeBytes;
                update.rightSizeBytes      = job.key.rightSizeBytes;
                update.leftLastWriteTime   = job.key.leftLastWriteTime;
                update.rightLastWriteTime  = job.key.rightLastWriteTime;
                update.leftFileAttributes  = job.leftFileAttributes;
                update.rightFileAttributes = job.rightFileAttributes;
                update.areEqual            = areEqual;

                _pendingContentCompareUpdates[folderKey][job.entryName] = update;
                _pendingContentCompareRetryAttempts.erase(folderKey);
                _pendingContentHighWater = std::max(_pendingContentHighWater, _pendingContentCompareUpdates.size());
                shouldNotify             = true;
                forceNotifyFinal         = _contentCompareQueueHigh.empty() && _contentCompareQueueLow.empty() && _contentCompareInFlight.empty();
                pendingAfter             = _contentComparePendingCompares.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
                pendingAfterComputed     = true;
            }
        }

        if (erased)
        {
            static_cast<void>(_contentCompareCompletedCompares.fetch_add(1u, std::memory_order_acq_rel));
            static_cast<void>(_contentCompareCompletedBytes.fetch_add(job.key.leftSizeBytes, std::memory_order_acq_rel));

            if (pendingAfterComputed && pendingAfter == 0u)
            {
                NotifyContentProgress(workerIndex, {}, {}, 0, 0);
            }
            else if (pendingAfterComputed)
            {
                NotifyContentProgress(workerIndex, job.relativeFolder, job.entryName, job.key.leftSizeBytes, job.key.leftSizeBytes);
            }
        }

        if (shouldNotify)
        {
            NotifyDecisionUpdated(false);
        }

        if (forceNotifyFinal)
        {
            NotifyDecisionUpdated(true);
        }
    }
}

class CompareDirectoriesFileSystem final : public IFileSystem, public IInformations
{
public:
    CompareDirectoriesFileSystem(ComparePane pane, std::shared_ptr<CompareDirectoriesSession> session) noexcept : _pane(pane), _session(std::move(session))
    {
        if (_session)
        {
            _baseFs    = _session->GetFileSystem(_pane);
            _baseInfos = _session->GetInformations(_pane);
        }
    }

    CompareDirectoriesFileSystem(const CompareDirectoriesFileSystem&)            = delete;
    CompareDirectoriesFileSystem& operator=(const CompareDirectoriesFileSystem&) = delete;
    CompareDirectoriesFileSystem(CompareDirectoriesFileSystem&&)                 = delete;
    CompareDirectoriesFileSystem& operator=(CompareDirectoriesFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        if (riid == __uuidof(IInformations))
        {
            *ppvObject = static_cast<IInformations*>(this);
            AddRef();
            return S_OK;
        }

        if (_baseFs)
        {
            return _baseFs->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    // IInformations
    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->GetMetaData(metaData);
    }

    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->GetConfigurationSchema(schemaJsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->SetConfiguration(configurationJsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->GetConfiguration(configurationJsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override
    {
        if (! _baseInfos)
        {
            return E_NOINTERFACE;
        }
        return _baseInfos->SomethingToSave(pSomethingToSave);
    }

    // IFileSystem
    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (ppFilesInformation == nullptr)
        {
            return E_POINTER;
        }

        *ppFilesInformation = nullptr;

        if (! _session || ! _baseFs)
        {
            return E_POINTER;
        }

        if (! _session->IsCompareEnabled())
        {
            return _baseFs->ReadDirectoryInfo(path, ppFilesInformation);
        }

        const std::filesystem::path absolute(path ? path : L"");
        const auto relOpt = _session->TryMakeRelative(_pane, absolute);
        if (! relOpt.has_value())
        {
            // Path outside compare roots: allow independent browsing by delegating to the base filesystem.
            return _baseFs->ReadDirectoryInfo(path, ppFilesInformation);
        }

        auto decision = _session->TryGetCachedDecision(relOpt.value());
        if (! decision)
        {
            _session->RequestScanForFolder(relOpt.value());

            HRESULT infoHr                       = S_OK;
            wil::com_ptr<IFilesInformation> info = BuildFilesInformation({}, infoHr);
            if (FAILED(infoHr) || ! info)
            {
                return FAILED(infoHr) ? infoHr : E_FAIL;
            }

            *ppFilesInformation = info.detach();
            return S_OK;
        }

        if (FAILED(decision->hr))
        {
            return decision->hr;
        }

        const Common::Settings::CompareDirectoriesSettings settings = _session->GetSettings();
        const bool showIdentical                                    = settings.showIdenticalItems;

        std::vector<OutEntry> out;
        out.reserve(decision->items.size());

        const bool isLeft = _pane == ComparePane::Left;
        for (const auto& [name, item] : decision->items)
        {
            const uint32_t diffMask = item.differenceMask;
            const bool pending = HasFlag(diffMask, CompareDirectoriesDiffBit::ContentPending) || HasFlag(diffMask, CompareDirectoriesDiffBit::SubdirPending);
            const bool include = isLeft ? (item.existsLeft && (showIdentical || item.isDifferent || pending))
                                        : (item.existsRight && (showIdentical || item.isDifferent || pending));
            if (! include)
            {
                continue;
            }

            OutEntry e{};
            e.name = name;
            if (isLeft)
            {
                e.fileAttributes = item.leftFileAttributes;
                e.lastWriteTime  = item.leftLastWriteTime;
                e.sizeBytes      = item.leftSizeBytes;
            }
            else
            {
                e.fileAttributes = item.rightFileAttributes;
                e.lastWriteTime  = item.rightLastWriteTime;
                e.sizeBytes      = item.rightSizeBytes;
            }

            out.emplace_back(std::move(e));
        }

        HRESULT infoHr                       = S_OK;
        wil::com_ptr<IFilesInformation> info = BuildFilesInformation(out, infoHr);
        if (FAILED(infoHr) || ! info)
        {
            return FAILED(infoHr) ? infoHr : E_FAIL;
        }

        *ppFilesInformation = info.detach();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* itemPath, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->DeleteItem(itemPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->DeleteItems(paths, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _baseFs ? _baseFs->RenameItems(items, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _baseFs ? _baseFs->GetCapabilities(jsonUtf8) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        return _baseFs ? _baseFs->GetTransferHints(path, operationType, endpoint, hints) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        return _baseFs ? _baseFs->GetStorageCharacteristics(path, characteristics) : E_POINTER;
    }

private:
    ~CompareDirectoriesFileSystem() = default;

    std::atomic_ulong _refCount{1};
    ComparePane _pane = ComparePane::Left;
    std::shared_ptr<CompareDirectoriesSession> _session;
    wil::com_ptr<IFileSystem> _baseFs;
    wil::com_ptr<IInformations> _baseInfos;
};

wil::com_ptr<IFileSystem> CreateCompareDirectoriesFileSystem(ComparePane pane, std::shared_ptr<CompareDirectoriesSession> session) noexcept
{
    if (! session)
    {
        return nullptr;
    }

    wil::com_ptr<IFileSystem> fs;
    fs.attach(static_cast<IFileSystem*>(new CompareDirectoriesFileSystem(pane, std::move(session))));
    return fs;
}
