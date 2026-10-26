#include "ViewerSpace.h"
#include "FileMetadataFormatting.h"
#include "FileSystemRouteContract.h"
#include "UnicodeClipboard.h"
#include "ViewerDxUiTheme.h"
#include "ViewerSpace.ListingValidation.h"
#include <windowsx.h>

#include "LocalizationManager.h"
#include "NaturalSort.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <exception>
#include <filesystem>
#include <format>
#include <iterator>
#include <limits>
#include <memory>
#include <mutex>
#include <numeric>
#include <psapi.h>
#include <set>
#include <shlwapi.h>
#include <string_view>
#include <system_error>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include "FluentIcons.h"
#include "FolderViewMenuIds.h"
#include "Helpers.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#include "resource.h"
#include <DxUi/DxUi.h>
#include <DxUi/Typography.h>

#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dxgi.lib")
#pragma comment(lib, "shlwapi.lib")

extern HINSTANCE g_hInstance;

namespace Typography = DxUi::Typography;

namespace
{
constexpr UINT_PTR kTimerAnimationId = 1;
constexpr UINT kAnimationIntervalMs  = 16;

constexpr int kHostFolderViewItemContextMenuResourceId = IDR_FOLDERVIEW_ITEM_CONTEXT;

constexpr UINT kCmdTreemapContextFocusInPane = 0xC100u;
constexpr UINT kCmdTreemapContextZoomIn      = 0xC101u;
constexpr UINT kCmdTreemapContextZoomOut     = 0xC102u;

constexpr UINT kCmdFolderViewContextOpen       = IDM_FOLDERVIEW_CONTEXT_OPEN;
constexpr UINT kCmdFolderViewContextOpenWith   = IDM_FOLDERVIEW_CONTEXT_OPEN_WITH;
constexpr UINT kCmdFolderViewContextDelete     = IDM_FOLDERVIEW_CONTEXT_DELETE;
constexpr UINT kCmdFolderViewContextRename     = IDM_FOLDERVIEW_CONTEXT_RENAME;
constexpr UINT kCmdFolderViewContextCopy       = IDM_FOLDERVIEW_CONTEXT_COPY;
constexpr UINT kCmdFolderViewContextPaste      = IDM_FOLDERVIEW_CONTEXT_PASTE;
constexpr UINT kCmdFolderViewContextProperties = IDM_FOLDERVIEW_CONTEXT_PROPERTIES;
constexpr UINT kCmdFolderViewContextMove       = IDM_FOLDERVIEW_CONTEXT_MOVE;
constexpr UINT kCmdFolderViewContextViewSpace  = IDM_FOLDERVIEW_CONTEXT_VIEW_SPACE;
constexpr UINT kCmdFolderViewContextCut        = IDM_FOLDERVIEW_CONTEXT_CUT;

constexpr UINT kFolderViewDebugCommandIdBase = 60000u;

void UpdateHostFolderViewContextCommandState(HMENU menu, const bool canExecuteLeafCommands, const bool isDirectory, const bool canExecutePaste) noexcept
{
    if (! menu)
    {
        return;
    }

    const auto setEnabled = [menu](const UINT commandId, const bool enabled) noexcept
    { static_cast<void>(EnableMenuItem(menu, commandId, static_cast<UINT>(MF_BYCOMMAND | (enabled ? MF_ENABLED : MF_GRAYED)))); };

    setEnabled(kCmdFolderViewContextOpen, canExecuteLeafCommands);
    setEnabled(kCmdFolderViewContextOpenWith, canExecuteLeafCommands && ! isDirectory);
    setEnabled(kCmdFolderViewContextDelete, canExecuteLeafCommands);
    setEnabled(kCmdFolderViewContextMove, canExecuteLeafCommands);
    setEnabled(kCmdFolderViewContextRename, canExecuteLeafCommands);
    setEnabled(kCmdFolderViewContextCut, canExecuteLeafCommands);
    setEnabled(kCmdFolderViewContextCopy, canExecuteLeafCommands);
    setEnabled(kCmdFolderViewContextProperties, canExecuteLeafCommands);
    setEnabled(kCmdFolderViewContextPaste, canExecutePaste);
}

constexpr float kHeaderHeightDip      = 72.0f;
constexpr float kNavigationHeightDip  = 36.0f;
constexpr float kHeaderButtonWidthDip = 52.0f;
constexpr float kPaddingDip           = 8.0f;
constexpr float kItemGapDip           = 1.0f;
constexpr float kMinHitAreaDip2       = 4.0f * 4.0f;
constexpr float kTooltipMaxWidthDip   = 420.0f;
constexpr float kTooltipMaxHeightDip  = 320.0f;

constexpr double kAnimationDurationSeconds          = 0.18;
constexpr double kScanCompletedOverlaySeconds       = 1.35;
constexpr double kTooltipHoverStabilityDelaySeconds = 0.12;

constexpr size_t kDefaultLayoutItems                   = 4096;
constexpr size_t kMaxNavigationHistory                 = 64;
constexpr size_t kMaxLayoutItems                       = 50000;
constexpr float kMinVisibleTileAreaDip2                = 64.0f;
constexpr uint32_t kAdaptiveFileCandidateCeiling       = 20000u;
constexpr uint64_t kMaxScanCacheSnapshotBytes          = 64ull * 1024ull * 1024ull;
constexpr uint64_t kMaxScanCacheResidentBytes          = 64ull * 1024ull * 1024ull;
constexpr uint64_t kViewerSpaceMemoryBudgetBytes       = 192ull * 1024ull * 1024ull;
constexpr uint64_t kViewerSpaceDerivedReserveBytes     = 16ull * 1024ull * 1024ull;
constexpr uint64_t kMaxAggregateActivePathBytes        = 32ull * 1024ull * 1024ull;
constexpr uint64_t kMaxAggregateValidationScratchBytes = 32ull * 1024ull * 1024ull;
constexpr uint64_t kMaxAggregateWorkerStagingBytes     = 32ull * 1024ull * 1024ull;
#ifdef ENABLE_TESTS
constexpr wchar_t kTestRefreshBudgetEnvVar[]              = L"REDSALAMANDER_VIEWERSPACE_FORCE_SMALL_REFRESH_BUDGET";
constexpr wchar_t kTestScanThreadsEnvVar[]                = L"REDSALAMANDER_VIEWERSPACE_TEST_SCAN_THREADS";
constexpr wchar_t kForceSmallScanPolicyEnvVar[]           = L"REDSALAMANDER_VIEWERSPACE_FORCE_SMALL_SCAN_POLICY";
constexpr wchar_t kTestMaxOutstandingEnvVar[]             = L"REDSALAMANDER_VIEWERSPACE_TEST_MAX_OUTSTANDING_DIRECTORIES";
constexpr wchar_t kTestMaxRetainedDirectoriesEnvVar[]     = L"REDSALAMANDER_VIEWERSPACE_TEST_MAX_RETAINED_DIRECTORIES";
constexpr wchar_t kTestNextModelIdEnvVar[]                = L"REDSALAMANDER_VIEWERSPACE_TEST_NEXT_MODEL_ID";
constexpr wchar_t kTestCacheAggregateCapEnvVar[]          = L"REDSALAMANDER_VIEWERSPACE_TEST_CACHE_AGGREGATE_BYTES";
constexpr wchar_t kTestCorruptCacheNameOffsetEnvVar[]     = L"REDSALAMANDER_VIEWERSPACE_TEST_CORRUPT_CACHE_NAME_OFFSET";
constexpr wchar_t kTestCorruptCacheRootStateEnvVar[]      = L"REDSALAMANDER_VIEWERSPACE_TEST_CORRUPT_CACHE_ROOT_STATE";
constexpr wchar_t kTestCorruptCacheChildOwnerEnvVar[]     = L"REDSALAMANDER_VIEWERSPACE_TEST_CORRUPT_CACHE_CHILD_OWNER";
constexpr wchar_t kTestCorruptCacheMissingChildEnvVar[]   = L"REDSALAMANDER_VIEWERSPACE_TEST_CORRUPT_CACHE_MISSING_CHILD";
constexpr wchar_t kTestValidationScratchBudgetEnvVar[]    = L"REDSALAMANDER_VIEWERSPACE_TEST_VALIDATION_SCRATCH_BYTES";
constexpr wchar_t kTestWorkerStagingBudgetEnvVar[]        = L"REDSALAMANDER_VIEWERSPACE_TEST_WORKER_STAGING_BYTES";
constexpr wchar_t kTestActivePathBudgetEnvVar[]           = L"REDSALAMANDER_VIEWERSPACE_TEST_ACTIVE_PATH_BYTES";
constexpr wchar_t kTestPublicationPhaseTimingEnvVar[]     = L"REDSALAMANDER_VIEWERSPACE_PUBLICATION_PHASES";
constexpr wchar_t kTestSummaryFaultMissingEndEnvVar[]     = L"REDSALAMANDER_VIEWERSPACE_TEST_SUMMARY_MISSING_END";
constexpr wchar_t kTestSummaryFaultDuplicateChunkEnvVar[] = L"REDSALAMANDER_VIEWERSPACE_TEST_SUMMARY_DUPLICATE_CHUNK";
constexpr wchar_t kTestSummaryFaultOutOfOrderEnvVar[]     = L"REDSALAMANDER_VIEWERSPACE_TEST_SUMMARY_OUT_OF_ORDER";
constexpr wchar_t kTestSummaryCopyBudgetEnvVar[]          = L"REDSALAMANDER_VIEWERSPACE_TEST_SUMMARY_COPY_BUDGET";
constexpr uint32_t kRendererFailureD3DDevice              = 1u;
constexpr uint32_t kRendererFailureD2DContext             = 2u;
constexpr uint32_t kRendererFailureSwapChain              = 3u;
constexpr uint32_t kRendererFailureResizeBuffers          = 4u;
constexpr uint32_t kRendererFailureGetBuffer              = 5u;
constexpr uint32_t kRendererFailureTargetBitmap           = 6u;
constexpr uint32_t kRendererFailureRenderTarget           = 7u;
constexpr uint32_t kRendererDiscardDestroy                = 20u;
constexpr uint32_t kRendererDiscardDpiChanged             = 21u;
constexpr uint32_t kRendererDiscardStaticDeviceLost       = 22u;
constexpr uint32_t kRendererDiscardDeviceLost             = 23u;
constexpr uint32_t kRendererDiscardTheme                  = 25u;
constexpr uint32_t kRendererDiscardUnknown                = 26u;
#endif

[[nodiscard]] uint64_t MaxScanCacheSnapshotBytes() noexcept
{
#ifdef ENABLE_TESTS
    std::array<wchar_t, 32> value{};
    const DWORD chars = GetEnvironmentVariableW(L"REDSALAMANDER_VIEWERSPACE_TEST_CACHE_CAP_BYTES", value.data(), static_cast<DWORD>(value.size()));
    if (chars > 0 && chars < value.size())
    {
        wchar_t* end                    = nullptr;
        const unsigned long long parsed = wcstoull(value.data(), &end, 10);
        if (parsed > 0 && end != value.data())
        {
            return std::min<uint64_t>(static_cast<uint64_t>(parsed), kMaxScanCacheSnapshotBytes);
        }
    }
#endif
    return kMaxScanCacheSnapshotBytes;
}

[[nodiscard]] uint64_t MaxScanCacheResidentBytes() noexcept
{
#ifdef ENABLE_TESTS
    std::array<wchar_t, 32> value{};
    const DWORD chars = GetEnvironmentVariableW(kTestCacheAggregateCapEnvVar, value.data(), static_cast<DWORD>(value.size()));
    if (chars > 0 && chars < value.size())
    {
        wchar_t* end                    = nullptr;
        const unsigned long long parsed = wcstoull(value.data(), &end, 10);
        if (parsed > 0 && end != value.data())
        {
            return std::min<uint64_t>(static_cast<uint64_t>(parsed), kMaxScanCacheResidentBytes);
        }
    }
#endif
    return kMaxScanCacheResidentBytes;
}

[[nodiscard]] bool FitsViewerSpaceCacheWorkingSet(uint64_t viewerOwnedBytes, uint64_t additionalCacheBytes) noexcept
{
    if (viewerOwnedBytes > kViewerSpaceMemoryBudgetBytes - kViewerSpaceDerivedReserveBytes)
    {
        return false;
    }
    return additionalCacheBytes <= kViewerSpaceMemoryBudgetBytes - kViewerSpaceDerivedReserveBytes - viewerOwnedBytes;
}

[[nodiscard]] bool FitsViewerSpaceRefreshWorkingSet(uint64_t viewerOwnedBytes, uint64_t additionalBytes) noexcept
{
    constexpr uint64_t scanScratchReserve = kMaxAggregateActivePathBytes + kMaxAggregateValidationScratchBytes + kMaxAggregateWorkerStagingBytes;
    uint64_t available                    = kViewerSpaceMemoryBudgetBytes - kViewerSpaceDerivedReserveBytes - scanScratchReserve;
#ifdef ENABLE_TESTS
    std::array<wchar_t, 32> value{};
    const DWORD chars = GetEnvironmentVariableW(kTestRefreshBudgetEnvVar, value.data(), static_cast<DWORD>(value.size()));
    if (chars > 0u && chars < value.size())
    {
        wchar_t* end                    = nullptr;
        const unsigned long long parsed = wcstoull(value.data(), &end, 10);
        if (parsed > 0u && end != value.data() && *end == L'\0')
        {
            available = std::min<uint64_t>(available, static_cast<uint64_t>(parsed));
        }
    }
#endif
    if (viewerOwnedBytes > available)
    {
        return false;
    }
    return additionalBytes <= available - viewerOwnedBytes;
}

#ifdef ENABLE_TESTS
[[nodiscard]] std::optional<uint32_t> ReadPositiveUint32Environment(std::wstring_view variableName) noexcept
{
    std::array<wchar_t, 32> value{};
    const DWORD chars = GetEnvironmentVariableW(variableName.data(), value.data(), static_cast<DWORD>(value.size()));
    if (chars == 0u || chars >= value.size())
    {
        return std::nullopt;
    }

    // Test budgets require unsigned decimal values; wcstoul also accepts signs and saturates overflow.
    uint32_t parsed = 0u;
    for (DWORD index = 0u; index < chars; ++index)
    {
        const wchar_t digit = value[index];
        if (digit < L'0' || digit > L'9')
        {
            return std::nullopt;
        }
        const uint32_t increment = static_cast<uint32_t>(digit - L'0');
        if (parsed > (std::numeric_limits<uint32_t>::max() - increment) / 10u)
        {
            return std::nullopt;
        }
        parsed = parsed * 10u + increment;
    }
    return parsed == 0u ? std::nullopt : std::optional<uint32_t>{parsed};
}
#endif

[[nodiscard]] ViewerSpaceScan::ResourcePolicy ActiveScanResourcePolicy() noexcept
{
    ViewerSpaceScan::ResourcePolicy policy = ViewerSpaceScan::kProductionResourcePolicy;
#ifdef ENABLE_TESTS
    if (EnvironmentVariables::IsTruthyFlagSet(kForceSmallScanPolicyEnvVar))
    {
        policy.maxRetainedDirectories      = 4u;
        policy.maxRetainedFileRecords      = 3u;
        policy.maxChildReferences          = 8u;
        policy.maxChildArenaSlots          = 32u;
        policy.maxRetainedNameBytes        = 512u;
        policy.maxProviderBufferBytes      = 64u * 1024u;
        policy.maxProviderEntriesPerFolder = 32u;
        policy.maxProviderNameBytes        = 128u;
        policy.maxTraversalDepth           = 4u;
        policy.maxTraversedDirectories     = 64u;
        policy.maxOutstandingDirectories   = 32u;
        policy.maxOutstandingPathBytes     = 8u * 1024u;
        policy.maxPathChars                = 256u;
    }
    if (const std::optional<uint32_t> value = ReadPositiveUint32Environment(kTestMaxOutstandingEnvVar); value.has_value())
    {
        policy.maxOutstandingDirectories = value.value();
    }
    if (const std::optional<uint32_t> value = ReadPositiveUint32Environment(kTestMaxRetainedDirectoriesEnvVar); value.has_value())
    {
        policy.maxRetainedDirectories = value.value();
    }
#endif
    return policy;
}

[[nodiscard]] uint64_t ActiveValidationScratchBudgetBytes() noexcept
{
#ifdef ENABLE_TESTS
    if (const std::optional<uint32_t> value = ReadPositiveUint32Environment(kTestValidationScratchBudgetEnvVar); value.has_value())
    {
        return std::min<uint64_t>(value.value(), kMaxAggregateValidationScratchBytes);
    }
#endif
    return kMaxAggregateValidationScratchBytes;
}

[[nodiscard]] uint64_t ActiveWorkerStagingBudgetBytes() noexcept
{
#ifdef ENABLE_TESTS
    if (const auto value = ReadPositiveUint32Environment(kTestWorkerStagingBudgetEnvVar); value.has_value())
        return std::min<uint64_t>(value.value(), kMaxAggregateWorkerStagingBytes);
#endif
    return kMaxAggregateWorkerStagingBytes;
}

[[nodiscard]] uint64_t ActivePathBudgetBytes(const ViewerSpaceScan::ResourcePolicy& policy) noexcept
{
    uint64_t budget = std::min(kMaxAggregateActivePathBytes, policy.maxOutstandingPathBytes);
#ifdef ENABLE_TESTS
    if (const std::optional<uint32_t> value = ReadPositiveUint32Environment(kTestActivePathBudgetEnvVar); value.has_value())
    {
        budget = std::min<uint64_t>(budget, value.value());
    }
#endif
    return budget;
}

[[nodiscard]] constexpr uint64_t ConservativePathAllocationBytes(uint64_t pathChars) noexcept
{
    // MSVC basic_string rounds a single reserve by at most seven wchar_t slots. Include those slots and the
    // terminator in the admitted charge; scanner paths are bounded by ResourcePolicy::maxPathChars.
    constexpr uint64_t kMsvcWcharReserveSlackAndTerminator = 8u;
    const uint64_t maxChars                                = (std::numeric_limits<uint64_t>::max)() / sizeof(wchar_t);
    if (pathChars > maxChars - kMsvcWcharReserveSlackAndTerminator)
    {
        return (std::numeric_limits<uint64_t>::max)();
    }
    return (pathChars + kMsvcWcharReserveSlackAndTerminator) * sizeof(wchar_t);
}

[[nodiscard]] size_t ActiveDirectorySummaryCopyBudget() noexcept
{
#ifdef ENABLE_TESTS
    if (const std::optional<uint32_t> value = ReadPositiveUint32Environment(kTestSummaryCopyBudgetEnvVar); value.has_value())
    {
        return std::min<size_t>(value.value(), 2048u);
    }
#endif
    return 2048u;
}

[[nodiscard]] constexpr std::wstring_view ValidationErrorDetail(ViewerSpaceScan::ValidationError error) noexcept
{
    using ViewerSpaceScan::ValidationError;
    switch (error)
    {
        case ValidationError::None: return L"none";
        case ValidationError::NullBuffer: return L"null-buffer";
        case ValidationError::UsedExceedsAllocated: return L"used-exceeds-allocated";
        case ValidationError::BufferLimit: return L"buffer-limit";
        case ValidationError::AllocatedBufferLimit: return L"allocated-buffer-limit";
        case ValidationError::EntryCountLimit: return L"entry-count-limit";
        case ValidationError::CountBufferMismatch: return L"count-buffer-mismatch";
        case ValidationError::MisalignedEntry: return L"misaligned-entry";
        case ValidationError::TruncatedHeader: return L"truncated-header";
        case ValidationError::InvalidNameLength: return L"invalid-name-length";
        case ValidationError::NameLimit: return L"name-limit";
        case ValidationError::TruncatedName: return L"truncated-name";
        case ValidationError::InvalidNextOffset: return L"invalid-next-offset";
        case ValidationError::UnsafeName: return L"unsafe-name";
        case ValidationError::DuplicateName: return L"duplicate-name";
        case ValidationError::DuplicateId: return L"duplicate-id";
        // Retired: FileIndex is not a tree-wide identity, so no scan reports it. The value keeps later errors' numbers stable.
        case ValidationError::AncestorCycle: return L"ancestor-cycle";
        case ValidationError::CountMismatch: return L"count-mismatch";
        case ValidationError::DepthLimit: return L"depth-limit";
        case ValidationError::TraversalLimit: return L"traversal-limit";
        case ValidationError::OutstandingLimit: return L"outstanding-limit";
        case ValidationError::PathLimit: return L"path-limit";
        case ValidationError::ItemIdLimit: return L"item-id-limit";
        default: return L"unknown";
    }
}

uint64_t SaturatingAtomicAdd(std::atomic_uint64_t& value, uint64_t delta) noexcept
{
    uint64_t current = value.load(std::memory_order_relaxed);
    for (;;)
    {
        const uint64_t next = ViewerSpaceScan::SaturatingAdd(current, delta);
        if (value.compare_exchange_weak(current, next, std::memory_order_relaxed, std::memory_order_relaxed))
        {
            return next;
        }
    }
}

#ifdef ENABLE_TESTS
[[nodiscard]] uint64_t SampleCurrentWorkingSetBytes() noexcept
{
    PROCESS_MEMORY_COUNTERS_EX counters{};
    counters.cb = static_cast<DWORD>(sizeof(counters));
    if (K32GetProcessMemoryInfo(GetCurrentProcess(), reinterpret_cast<PROCESS_MEMORY_COUNTERS*>(&counters), counters.cb) == 0)
    {
        return 0u;
    }

    return static_cast<uint64_t>(counters.WorkingSetSize);
}
#endif

[[nodiscard]] uint32_t DefaultWin32ScanThreads() noexcept
{
    const uint32_t hardware = std::max<uint32_t>(1u, std::thread::hardware_concurrency());
    return std::clamp<uint32_t>(hardware, 1u, 8u);
}

[[nodiscard]] uint32_t PixelDrivenLayoutItemLimit(float projectedAreaDip2) noexcept
{
    if (projectedAreaDip2 <= 0.0f || ! std::isfinite(projectedAreaDip2))
    {
        return static_cast<uint32_t>(kDefaultLayoutItems);
    }

    const auto pixelBudget        = static_cast<uint32_t>(std::ceil(projectedAreaDip2 / kMinVisibleTileAreaDip2));
    const auto conservativeBudget = std::max<uint32_t>(static_cast<uint32_t>(kDefaultLayoutItems), pixelBudget);
    return std::clamp<uint32_t>(conservativeBudget, 32u, static_cast<uint32_t>(kMaxLayoutItems));
}

[[nodiscard]] size_t CountOwnerDrawMenuItems(HMENU menu) noexcept
{
    if (! menu)
    {
        return 0u;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return 0u;
    }

    size_t ownerDrawCount = 0u;
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_SUBMENU;
        if (GetMenuItemInfoW(menu, position, TRUE, &itemInfo) == 0)
        {
            continue;
        }

        if ((itemInfo.fType & MFT_OWNERDRAW) != 0)
        {
            ++ownerDrawCount;
        }

        if (itemInfo.hSubMenu)
        {
            ownerDrawCount += CountOwnerDrawMenuItems(itemInfo.hSubMenu);
        }
    }

    return ownerDrawCount;
}

[[nodiscard]] std::wstring StripMenuMnemonicMarkers(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());

    for (size_t index = 0; index < text.size(); ++index)
    {
        if (text[index] != L'&')
        {
            result.push_back(text[index]);
            continue;
        }

        if ((index + 1u) < text.size() && text[index + 1u] == L'&')
        {
            result.push_back(L'&');
            ++index;
        }
    }

    return result;
}

[[nodiscard]] std::wstring GetViewerSpaceContextMenuIconGlyph(UINT commandId) noexcept
{
    switch (commandId)
    {
        case kCmdFolderViewContextOpen: return std::wstring(1u, FluentIcons::kOpenFile);
        case kCmdFolderViewContextCut: return std::wstring(1u, FluentIcons::kCut);
        case kCmdFolderViewContextCopy: return std::wstring(1u, FluentIcons::kCopy);
        case kCmdFolderViewContextPaste: return std::wstring(1u, FluentIcons::kPaste);
        case kCmdFolderViewContextDelete: return std::wstring(1u, FluentIcons::kDelete);
        case kCmdFolderViewContextRename: return std::wstring(1u, FluentIcons::kRename);
        case kCmdFolderViewContextProperties: return std::wstring(1u, FluentIcons::kInfo);
        default: return {};
    }
}

[[nodiscard]] std::vector<DxUi::MenuFlyoutItem> ConvertHMenuToDxFlyoutItems(HMENU menu) noexcept
{
    using DxUi::MenuFlyoutItem;
    using DxUi::MenuItemKind;

    std::vector<MenuFlyoutItem> items;
    if (! menu)
    {
        return items;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return items;
    }

    items.reserve(static_cast<size_t>(itemCount));
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, position, TRUE, &itemInfo))
        {
            continue;
        }

        MenuFlyoutItem item{};
        if ((itemInfo.fType & MFT_SEPARATOR) != 0)
        {
            item.kind = MenuItemKind::Separator;
            items.push_back(std::move(item));
            continue;
        }

        std::array<wchar_t, 512> buffer{};
        const int textLength = GetMenuStringW(menu, position, buffer.data(), static_cast<int>(buffer.size()), MF_BYPOSITION);
        if (textLength > 0)
        {
            std::wstring_view fullText(buffer.data(), static_cast<size_t>(textLength));
            const size_t tabPos               = fullText.find(L'\t');
            const std::wstring_view labelText = tabPos == std::wstring_view::npos ? fullText : fullText.substr(0, tabPos);
            item.text                         = StripMenuMnemonicMarkers(labelText);
            if (tabPos != std::wstring_view::npos && (tabPos + 1u) < fullText.size())
            {
                item.acceleratorText.assign(fullText.substr(tabPos + 1u));
            }
        }

        item.commandId = static_cast<int>(itemInfo.wID);
        item.enabled   = (itemInfo.fState & MFS_GRAYED) == 0;
        item.checked   = (itemInfo.fState & MFS_CHECKED) != 0;

        if ((itemInfo.fType & MFT_RADIOCHECK) != 0)
        {
            item.kind = MenuItemKind::Radio;
        }
        else if (item.checked)
        {
            item.kind = MenuItemKind::Toggle;
        }

        item.iconGlyph = GetViewerSpaceContextMenuIconGlyph(itemInfo.wID);

        if (itemInfo.hSubMenu)
        {
            item.children = ConvertHMenuToDxFlyoutItems(itemInfo.hSubMenu);
        }

        items.push_back(std::move(item));
    }

    return items;
}

struct ViewerSpaceClassBackgroundBrushState
{
    ViewerSpaceClassBackgroundBrushState()                                                       = default;
    ViewerSpaceClassBackgroundBrushState(const ViewerSpaceClassBackgroundBrushState&)            = delete;
    ViewerSpaceClassBackgroundBrushState& operator=(const ViewerSpaceClassBackgroundBrushState&) = delete;
    ViewerSpaceClassBackgroundBrushState(ViewerSpaceClassBackgroundBrushState&&)                 = delete;
    ViewerSpaceClassBackgroundBrushState& operator=(ViewerSpaceClassBackgroundBrushState&&)      = delete;

    wil::unique_hbrush activeBrush;
    COLORREF activeColor = CLR_INVALID;

    wil::unique_hbrush pendingBrush;
    COLORREF pendingColor = CLR_INVALID;

    bool classRegistered = false;
};

ViewerSpaceClassBackgroundBrushState g_viewerSpaceClassBackgroundBrush;
std::atomic_uint32_t g_viewerSpaceWindowCount{0};
ATOM g_viewerSpaceClassAtom = 0;

void DestroyViewerSpaceClassBackgroundBrushState() noexcept
{
    g_viewerSpaceClassBackgroundBrush.classRegistered = false;
    g_viewerSpaceClassBackgroundBrush.pendingBrush.reset();
    g_viewerSpaceClassBackgroundBrush.pendingColor = CLR_INVALID;
    g_viewerSpaceClassBackgroundBrush.activeBrush.reset();
    g_viewerSpaceClassBackgroundBrush.activeColor = CLR_INVALID;
}

HBRUSH GetActiveViewerSpaceClassBackgroundBrush() noexcept
{
    if (g_viewerSpaceClassBackgroundBrush.pendingBrush)
    {
        return g_viewerSpaceClassBackgroundBrush.pendingBrush.get();
    }

    if (! g_viewerSpaceClassBackgroundBrush.activeBrush)
    {
        const COLORREF fallback                       = GetSysColor(COLOR_WINDOW);
        g_viewerSpaceClassBackgroundBrush.activeColor = fallback;
        g_viewerSpaceClassBackgroundBrush.activeBrush.reset(CreateSolidBrush(fallback));
    }

    return g_viewerSpaceClassBackgroundBrush.activeBrush.get();
}

void RequestViewerSpaceClassBackgroundColor(COLORREF color) noexcept
{
    if (color == CLR_INVALID)
    {
        return;
    }

    if (g_viewerSpaceClassBackgroundBrush.pendingBrush && g_viewerSpaceClassBackgroundBrush.pendingColor == color)
    {
        return;
    }

    wil::unique_hbrush brush(CreateSolidBrush(color));
    if (! brush)
    {
        return;
    }

    g_viewerSpaceClassBackgroundBrush.pendingColor = color;
    g_viewerSpaceClassBackgroundBrush.pendingBrush = std::move(brush);
}

void ApplyPendingViewerSpaceClassBackgroundBrush(HWND hwnd) noexcept
{
    if (! hwnd || ! g_viewerSpaceClassBackgroundBrush.pendingBrush || ! g_viewerSpaceClassBackgroundBrush.classRegistered)
    {
        return;
    }

    const LONG_PTR newBrush = reinterpret_cast<LONG_PTR>(g_viewerSpaceClassBackgroundBrush.pendingBrush.get());
    SetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND, newBrush);

    g_viewerSpaceClassBackgroundBrush.activeBrush  = std::move(g_viewerSpaceClassBackgroundBrush.pendingBrush);
    g_viewerSpaceClassBackgroundBrush.activeColor  = g_viewerSpaceClassBackgroundBrush.pendingColor;
    g_viewerSpaceClassBackgroundBrush.pendingColor = CLR_INVALID;
}

constexpr char kViewerSpaceSchemaJson[] = R"json({
    "version": 1,
    "title": "Space Viewer",
    "fields": [
        {
            "key": "viewMode",
            "type": "option",
            "label": "Default layout",
            "default": "explorer",
            "options": [
                {
                    "value": "explorer",
                    "label": "Panel right"
                },
                {
                    "value": "analyst",
                    "label": "Panel bottom"
                }
            ]
        },
        {
            "key": "analysisPanelVisible",
            "type": "option",
            "label": "Show analysis",
            "default": "1",
            "options": [
                {
                    "value": "0",
                    "label": "Off"
                },
                {
                    "value": "1",
                    "label": "On"
                }
            ]
        },
        {
            "key": "compactRows",
            "type": "option",
            "label": "Compact rows",
            "default": "0",
            "options": [
                {
                    "value": "0",
                    "label": "Off"
                },
                {
                    "value": "1",
                    "label": "On"
                }
            ]
        },
        {
            "key": "fileCountWeight",
            "type": "option",
            "label": "Default map area",
            "description": "Size matches Salamander. File count becomes available after bounded retained metadata is prepared.",
            "default": "0",
            "options": [
                {
                    "value": "0",
                    "label": "Size"
                },
                {
                    "value": "1",
                    "label": "File count"
                }
            ]
        },
        {
            "key": "colorMode",
            "type": "option",
            "label": "Default colors",
            "default": "folder",
            "options": [
                {
                    "value": "folder",
                    "label": "Folder groups"
                },
                {
                    "value": "type",
                    "label": "File type"
                },
                {
                    "value": "age",
                    "label": "Age: Modified date"
                },
                {
                    "value": "newest",
                    "label": "Age: newest known contained file"
                }
            ]
        },
        {
            "key": "topFilesPerDirectory",
            "type": "value",
            "label": "Top files per directory",
            "description": "Compatibility floor for largest files retained per directory; the renderer may retain more candidates when pixel/detail budgets allow it.",
            "default": 96,
            "min": 0,
            "max": 20000,
            "x-ui-section": "Advanced"
        },
        {
            "key": "scanThreads",
            "type": "value",
            "label": "Scan threads",
            "description": "Number of background threads used to scan subfolders in parallel.",
            "default": 1,
            "min": 1,
            "max": 16,
            "x-ui-section": "Advanced"
        },
        {
            "key": "maxConcurrentScansPerVolume",
            "type": "value",
            "label": "Max concurrent scans per volume",
            "description": "Limits how many Space viewers scan the same drive at once (reduces disk thrash when opening multiple viewers).",
            "default": 1,
            "min": 1,
            "max": 8,
            "x-ui-section": "Advanced"
        },
        {
            "key": "cacheEnabled",
            "type": "option",
            "label": "Scan cache",
            "description": "Reuse recent scan results when opening another Space viewer on the same root.",
            "default": "1",
            "options": [
                {
                    "value": "0",
                    "label": "Off"
                },
                {
                    "value": "1",
                    "label": "On"
                }
            ],
            "x-ui-section": "Advanced"
        },
        {
            "key": "cacheTtlSeconds",
            "type": "value",
            "label": "Cache TTL (seconds)",
            "description": "How long an unused scan result remains reusable; 0 keeps it until it is replaced or evicted.",
            "default": 60,
            "min": 0,
            "max": 3600,
            "x-ui-section": "Advanced"
        },
        {
            "key": "cacheMaxEntries",
            "type": "value",
            "label": "Cache max entries",
            "description": "Maximum number of cached roots kept in memory.",
            "default": 1,
            "min": 0,
            "max": 16,
            "x-ui-section": "Advanced"
        }
    ]
})json";

[[nodiscard]] const char* GetViewerSpaceStaticConfigurationSchemaImpl() noexcept
{
    return kViewerSpaceSchemaJson;
}

std::atomic_uint32_t g_maxConcurrentScansPerVolume{1};
std::atomic_bool g_cacheEnabled{true};
std::atomic_uint32_t g_cacheTtlSeconds{60};
std::atomic_uint32_t g_cacheMaxEntries{1};

std::wstring_view CopyToArena(std::pmr::monotonic_buffer_resource& arena, std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const size_t count = text.size();
    const size_t bytes = (count + 1u) * sizeof(wchar_t);
    void* memory       = arena.allocate(bytes, alignof(wchar_t));
    auto* buffer       = static_cast<wchar_t*>(memory);
    std::copy_n(text.data(), count, buffer);
    buffer[count] = L'\0';
    return std::wstring_view(buffer, count);
}

uint32_t GetMaxConcurrentScansPerVolume() noexcept
{
    const uint32_t value = g_maxConcurrentScansPerVolume.load(std::memory_order_acquire);
    return std::clamp(value, 1u, 8u);
}

using Common::Colors::BlendColorRefTruncate;
using Common::Colors::ColorRefFromArgb;
using DxUi::ColorFromArgb;

COLORREF ChooseContrastingTextColor(COLORREF background) noexcept
{
    const float r   = static_cast<float>(GetRValue(background)) / 255.0f;
    const float g   = static_cast<float>(GetGValue(background)) / 255.0f;
    const float b   = static_cast<float>(GetBValue(background)) / 255.0f;
    const float lum = 0.2126f * r + 0.7152f * g + 0.0722f * b;
    return lum > 0.60f ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

D2D1::ColorF ColorFFromHsv(double hue01, double saturation, double value, float alpha) noexcept;

bool IsAsciiAlpha(wchar_t ch) noexcept
{
    return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z');
}

wchar_t DeterminePreferredPathSeparator(std::wstring_view path, bool fileSystemIsWin32) noexcept;

bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

bool LooksLikeWin32Path(std::wstring_view path) noexcept
{
    if (path.size() >= 2 && IsAsciiAlpha(path[0]) && path[1] == L':')
    {
        return true;
    }

    if (path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\')
    {
        return true;
    }

    return false;
}

float MeasureTextWidthDip(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text) noexcept
{
    if (factory == nullptr || format == nullptr || text.empty())
    {
        return 0.0f;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    static constexpr float kMeasureWidthDip  = 8192.0f;
    static constexpr float kMeasureHeightDip = 256.0f;
    const HRESULT hr = factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, kMeasureWidthDip, kMeasureHeightDip, layout.put());
    if (FAILED(hr) || ! layout)
    {
        return 0.0f;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return 0.0f;
    }

    return metrics.widthIncludingTrailingWhitespace;
}

bool FitsTextWidthDip(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text, float maxWidthDip) noexcept
{
    if (maxWidthDip <= 0.0f)
    {
        return false;
    }

    const float width = MeasureTextWidthDip(factory, format, text);
    return width <= maxWidthDip + 0.25f;
}

std::wstring BuildTailEllipsisText(std::wstring_view text, IDWriteFactory* factory, IDWriteTextFormat* format, float maxWidthDip) noexcept
{
    if (text.empty() || factory == nullptr || format == nullptr)
    {
        return std::wstring(text);
    }

    if (FitsTextWidthDip(factory, format, text, maxWidthDip))
    {
        return std::wstring(text);
    }

    static constexpr wchar_t kEllipsis[] = L"\u2026";
    if (FitsTextWidthDip(factory, format, kEllipsis, maxWidthDip))
    {
        std::wstring best = kEllipsis;

        size_t low  = 0;
        size_t high = text.size();
        while (low < high)
        {
            const size_t mid = low + (high - low + 1) / 2;
            std::wstring cand;
            cand.reserve(1u + mid);
            cand.append(kEllipsis);
            cand.append(text.substr(text.size() - mid));
            if (FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                best = std::move(cand);
                low  = mid;
            }
            else
            {
                high = mid - 1;
            }
        }

        return best;
    }

    return kEllipsis;
}

std::wstring BuildTrailingEllipsisText(std::wstring_view text, IDWriteFactory* factory, IDWriteTextFormat* format, float maxWidthDip) noexcept
{
    if (text.empty() || factory == nullptr || format == nullptr)
    {
        return std::wstring(text);
    }

    if (FitsTextWidthDip(factory, format, text, maxWidthDip))
    {
        return std::wstring(text);
    }

    static constexpr wchar_t kEllipsis[] = L"\u2026";
    if (FitsTextWidthDip(factory, format, kEllipsis, maxWidthDip))
    {
        std::wstring best = kEllipsis;

        size_t low  = 0;
        size_t high = text.size();
        while (low < high)
        {
            const size_t mid = low + (high - low + 1) / 2;
            size_t prefixLen = mid;

            if (prefixLen > 0 && prefixLen < text.size())
            {
                const wchar_t last             = text[prefixLen - 1];
                const wchar_t next             = text[prefixLen];
                const bool lastIsHighSurrogate = (last >= 0xD800 && last <= 0xDBFF);
                const bool nextIsLowSurrogate  = (next >= 0xDC00 && next <= 0xDFFF);
                if (lastIsHighSurrogate && nextIsLowSurrogate)
                {
                    prefixLen -= 1;
                }
            }

            std::wstring cand;
            cand.reserve(prefixLen + 1u);
            cand.append(text.substr(0, prefixLen));
            cand.append(kEllipsis);
            if (FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                best = std::move(cand);
                low  = mid;
            }
            else
            {
                high = mid - 1;
            }
        }

        return best;
    }

    return kEllipsis;
}

struct PathEllipsisParts final
{
    std::wstring_view root;
    std::vector<std::wstring_view> segments;
    wchar_t separator = L'\\';
};

PathEllipsisParts SplitPathForEllipsis(std::wstring_view path, bool fileSystemIsWin32) noexcept
{
    PathEllipsisParts parts;
    parts.separator = DeterminePreferredPathSeparator(path, fileSystemIsWin32);

    size_t rootLen = 0;
    if (path.size() >= 2 && IsPathSeparator(path[0]) && IsPathSeparator(path[1]))
    {
        // UNC: \\server\share\...
        auto findSep = [&](size_t start) noexcept -> size_t
        {
            for (size_t i = start; i < path.size(); ++i)
            {
                if (IsPathSeparator(path[i]))
                {
                    return i;
                }
            }
            return std::wstring_view::npos;
        };

        const size_t serverEnd = findSep(2);
        if (serverEnd == std::wstring_view::npos)
        {
            parts.root = path;
            return parts;
        }

        const size_t shareEnd = findSep(serverEnd + 1);
        if (shareEnd == std::wstring_view::npos)
        {
            parts.root = path;
            return parts;
        }

        rootLen    = shareEnd + 1;
        parts.root = path.substr(0, rootLen);
    }
    else if (path.size() >= 2 && IsAsciiAlpha(path[0]) && path[1] == L':')
    {
        // Drive root: C:\...
        rootLen = 2;
        if (path.size() >= 3 && IsPathSeparator(path[2]))
        {
            rootLen = 3;
        }
        parts.root = path.substr(0, rootLen);
    }
    else if (! path.empty() && IsPathSeparator(path[0]))
    {
        rootLen    = 1;
        parts.root = path.substr(0, rootLen);
    }

    size_t pos = rootLen;
    while (pos < path.size() && IsPathSeparator(path[pos]))
    {
        pos += 1;
    }

    while (pos < path.size())
    {
        const size_t start = pos;
        while (pos < path.size() && ! IsPathSeparator(path[pos]))
        {
            pos += 1;
        }
        if (pos > start)
        {
            parts.segments.push_back(path.substr(start, pos - start));
        }

        while (pos < path.size() && IsPathSeparator(path[pos]))
        {
            pos += 1;
        }
    }

    return parts;
}

std::wstring BuildMiddleEllipsisPathText(
    std::wstring_view fullText, bool fileSystemIsWin32, IDWriteFactory* factory, IDWriteTextFormat* format, float maxWidthDip) noexcept
{
    if (fullText.empty() || factory == nullptr || format == nullptr)
    {
        return std::wstring(fullText);
    }

    if (FitsTextWidthDip(factory, format, fullText, maxWidthDip))
    {
        return std::wstring(fullText);
    }

    const bool hasSeparator = fullText.find_first_of(L"\\/") != std::wstring_view::npos;
    if (! hasSeparator)
    {
        return BuildTailEllipsisText(fullText, factory, format, maxWidthDip);
    }

    const PathEllipsisParts parts = SplitPathForEllipsis(fullText, fileSystemIsWin32);
    if (parts.segments.empty())
    {
        return BuildTailEllipsisText(fullText, factory, format, maxWidthDip);
    }

    auto buildCandidate = [&](std::wstring_view root, size_t prefixCount, size_t suffixCount) -> std::wstring
    {
        static constexpr wchar_t kEllipsis[] = L"\u2026";

        std::wstring out;
        out.reserve(fullText.size() + 4);
        out.append(root);

        auto appendSeg = [&](std::wstring_view seg) noexcept
        {
            if (seg.empty())
            {
                return;
            }
            if (! out.empty() && ! IsPathSeparator(out.back()))
            {
                out.push_back(parts.separator);
            }
            out.append(seg.data(), seg.size());
        };

        for (size_t i = 0; i < prefixCount; ++i)
        {
            appendSeg(parts.segments[i]);
        }

        if (prefixCount > 0 && ! out.empty() && ! IsPathSeparator(out.back()))
        {
            out.push_back(parts.separator);
        }

        out.append(kEllipsis);

        if (suffixCount > 0)
        {
            if (! out.empty() && ! IsPathSeparator(out.back()))
            {
                out.push_back(parts.separator);
            }

            const size_t start = parts.segments.size() - suffixCount;
            for (size_t i = start; i < parts.segments.size(); ++i)
            {
                if (i != start)
                {
                    out.push_back(parts.separator);
                }
                out.append(parts.segments[i].data(), parts.segments[i].size());
            }
        }

        return out;
    };

    auto bestForRoot = [&](std::wstring_view root) -> std::wstring
    {
        const size_t total = parts.segments.size();

        size_t suffixCount = 1;
        std::wstring best  = buildCandidate(root, 0, suffixCount);
        if (! FitsTextWidthDip(factory, format, best, maxWidthDip))
        {
            return {};
        }

        while (suffixCount + 1 <= total)
        {
            const std::wstring cand = buildCandidate(root, 0, suffixCount + 1);
            if (! FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                break;
            }
            suffixCount += 1;
            best = cand;
        }

        size_t prefixCount = 0;
        while (prefixCount + 1 + suffixCount <= total)
        {
            const std::wstring cand = buildCandidate(root, prefixCount + 1, suffixCount);
            if (! FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                break;
            }
            prefixCount += 1;
            best = cand;
        }

        return best;
    };

    std::wstring best = bestForRoot(parts.root);
    if (! best.empty())
    {
        return best;
    }

    best = bestForRoot({});
    if (! best.empty())
    {
        return best;
    }

    const std::wstring_view leaf = parts.segments.back();

    auto tryTrimLeaf = [&](std::wstring_view root) noexcept -> std::wstring
    {
        static constexpr wchar_t kEllipsis[] = L"\u2026";

        std::wstring prefix;
        prefix.reserve(root.size() + 4);
        prefix.append(root);
        prefix.append(kEllipsis);
        prefix.push_back(parts.separator);

        const float prefixWidth = MeasureTextWidthDip(factory, format, prefix);
        if (prefixWidth <= 0.0f || prefixWidth >= maxWidthDip)
        {
            return {};
        }

        float leafWidthDip = maxWidthDip - prefixWidth;
        if (leafWidthDip <= 1.0f)
        {
            return {};
        }

        for (int attempt = 0; attempt < 4 && leafWidthDip > 0.0f; ++attempt)
        {
            const std::wstring trimmedLeaf = BuildTrailingEllipsisText(leaf, factory, format, leafWidthDip);
            std::wstring cand;
            cand.reserve(prefix.size() + trimmedLeaf.size());
            cand.append(prefix);
            cand.append(trimmedLeaf);

            if (FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                return cand;
            }

            leafWidthDip -= 1.0f;
        }

        return {};
    };

    best = tryTrimLeaf(parts.root);
    if (! best.empty())
    {
        return best;
    }

    best = tryTrimLeaf({});
    if (! best.empty())
    {
        return best;
    }

    if (FitsTextWidthDip(factory, format, leaf, maxWidthDip))
    {
        return std::wstring(leaf);
    }

    return BuildTrailingEllipsisText(leaf, factory, format, maxWidthDip);
}

std::wstring_view TrimTrailingPathSeparators(std::wstring_view path) noexcept
{
    while (path.size() > 1)
    {
        const wchar_t last = path.back();
        if (last != L'/' && last != L'\\')
        {
            break;
        }
        path.remove_suffix(1);
    }
    return path;
}

std::optional<std::wstring> TryGetParentPathForNavigationGeneric(std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return std::nullopt;
    }

    std::wstring_view trimmed = TrimTrailingPathSeparators(path);
    if (trimmed.empty() || trimmed == L"/" || trimmed == L"\\")
    {
        return std::nullopt;
    }

    const size_t lastSep = trimmed.find_last_of(L"/\\");
    if (lastSep == std::wstring_view::npos)
    {
        return std::nullopt;
    }

    if (lastSep == 0)
    {
        return std::wstring(trimmed.substr(0, 1));
    }

    if (lastSep > 0 && trimmed[lastSep - 1] == L':')
    {
        // For plugin paths like "sftp:/home", the parent should be "sftp:/", not "sftp:" (which has special meaning).
        return std::wstring(trimmed.substr(0, lastSep + 1));
    }

    return std::wstring(trimmed.substr(0, lastSep));
}

wchar_t DeterminePreferredPathSeparator(std::wstring_view path, bool fileSystemIsWin32) noexcept
{
    if (fileSystemIsWin32)
    {
        return L'\\';
    }

    const bool hasForward = path.find(L'/') != std::wstring_view::npos;
    const bool hasBack    = path.find(L'\\') != std::wstring_view::npos;

    if (hasForward && ! hasBack)
    {
        return L'/';
    }

    if (hasBack && ! hasForward)
    {
        return L'\\';
    }

    if (LooksLikeWin32Path(path))
    {
        return L'\\';
    }

    return L'/';
}

std::wstring JoinPath(std::wstring_view parent, std::wstring_view leaf, wchar_t separator) noexcept
{
    const bool addSeparator = ! parent.empty() && parent.back() != L'/' && parent.back() != L'\\';
    std::wstring result;
    result.reserve(parent.size() + leaf.size() + (addSeparator ? 1u : 0u));
    if (! parent.empty())
    {
        result.append(parent.data(), parent.size());
    }
    if (addSeparator)
    {
        result.push_back(separator);
    }
    if (! leaf.empty())
    {
        result.append(leaf.data(), leaf.size());
    }
    return result;
}

std::optional<std::filesystem::path> TryGetParentPathForNavigation(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return std::nullopt;
    }

    const std::filesystem::path normalized = path.lexically_normal();
    const std::filesystem::path root       = normalized.root_path();
    if (! root.empty() && normalized == root)
    {
        return std::nullopt;
    }

    const std::filesystem::path parent = normalized.parent_path();
    if (parent.empty() || parent == normalized)
    {
        return std::nullopt;
    }

    if (! parent.has_root_directory())
    {
        // For drive roots, parent_path() can produce "C:" which is not a navigable folder.
        return std::nullopt;
    }

    return parent;
}

D2D1::ColorF Mix(const D2D1::ColorF& a, const D2D1::ColorF& b, float t) noexcept
{
    const float clamped = std::clamp(t, 0.0f, 1.0f);
    return D2D1::ColorF(a.r + (b.r - a.r) * clamped, a.g + (b.g - a.g) * clamped, a.b + (b.b - a.b) * clamped, a.a + (b.a - a.a) * clamped);
}

double Fract(double value) noexcept
{
    return value - std::floor(value);
}

double EaseOutCubic(double t) noexcept
{
    const double clamped = std::clamp(t, 0.0, 1.0);
    const double inv     = 1.0 - clamped;
    return 1.0 - (inv * inv * inv);
}

uint32_t HashU32(uint32_t value) noexcept
{
    // SplitMix32-style mixing (stable across runs).
    value ^= value >> 16;
    value *= 0x7feb352du;
    value ^= value >> 15;
    value *= 0x846ca68bu;
    value ^= value >> 16;
    return value;
}

struct CompactBytesText final
{
    std::array<wchar_t, 64> buffer{};
    UINT32 length = 0;
};

CompactBytesText FormatBytesCompactInline(uint64_t bytes) noexcept
{
    CompactBytesText text;

    double value       = static_cast<double>(bytes);
    size_t suffixIndex = 0;
#pragma warning(push)
// 5246: the initialization of a subobject should be wrapped in braces
#pragma warning(disable : 5246)
    static constexpr std::array<std::wstring_view, 5> suffixes = {
        L"B",
        L"KB",
        L"MB",
        L"GB",
        L"TB",
    };
#pragma warning(pop)

    while (value >= 1024.0 && (suffixIndex + 1) < suffixes.size())
    {
        value /= 1024.0;
        suffixIndex += 1;
    }

    const auto& locale   = LocaleFormatting::GetFormatLocale();
    wchar_t* const begin = text.buffer.data();
    constexpr size_t max = (sizeof(text.buffer) / sizeof(text.buffer[0])) - 1;

    if (suffixIndex == 0)
    {
        const auto r         = std::format_to_n(begin, max, locale, L"{:L} {}", bytes, suffixes[suffixIndex]);
        const size_t written = (r.size < max) ? r.size : max;
        begin[written]       = L'\0';
        text.length          = static_cast<UINT32>(written);
    }
    else
    {
        int decimals = 0;
        if (value < 10.0)
        {
            decimals = (value >= 9.995) ? 1 : 2;
        }
        else if (value < 100.0)
        {
            decimals = (value >= 99.95) ? 0 : 1;
        }

        const auto r         = std::format_to_n(begin, max, locale, L"{:.{}Lf} {}", value, decimals, suffixes[suffixIndex]);
        const size_t written = (r.size < max) ? r.size : max;
        begin[written]       = L'\0';
        text.length          = static_cast<UINT32>(written);
    }

    return text;
}

D2D1::ColorF ColorFFromHsv(double hue01, double saturation, double value, float alpha) noexcept
{
    const double h   = Fract(hue01) * 6.0;
    const int sector = static_cast<int>(std::floor(h));
    const double f   = h - static_cast<double>(sector);

    const double s = std::clamp(saturation, 0.0, 1.0);
    const double v = std::clamp(value, 0.0, 1.0);

    const double p = v * (1.0 - s);
    const double q = v * (1.0 - s * f);
    const double t = v * (1.0 - s * (1.0 - f));

    double r = v;
    double g = t;
    double b = p;

    switch (sector)
    {
        case 0:
            r = v;
            g = t;
            b = p;
            break;
        case 1:
            r = q;
            g = v;
            b = p;
            break;
        case 2:
            r = p;
            g = v;
            b = t;
            break;
        case 3:
            r = p;
            g = q;
            b = v;
            break;
        case 4:
            r = t;
            g = p;
            b = v;
            break;
        default:
            r = v;
            g = p;
            b = q;
            break;
    }

    return D2D1::ColorF(static_cast<float>(r), static_cast<float>(g), static_cast<float>(b), alpha);
}

float RectArea(const D2D1_RECT_F& rc) noexcept
{
    const float w = std::max(0.0f, rc.right - rc.left);
    const float h = std::max(0.0f, rc.bottom - rc.top);
    return w * h;
}

class ScanScheduler final
{
public:
    ScanScheduler()  = default;
    ~ScanScheduler() = default;

    ScanScheduler(const ScanScheduler&)             = delete;
    ScanScheduler(const ScanScheduler&&)            = delete;
    ScanScheduler& operator=(const ScanScheduler&)  = delete;
    ScanScheduler& operator=(const ScanScheduler&&) = delete;

    struct VolumeEntry final
    {
        VolumeEntry()  = default;
        ~VolumeEntry() = default;

        VolumeEntry(const VolumeEntry&)            = delete;
        VolumeEntry(VolumeEntry&&)                 = delete;
        VolumeEntry& operator=(const VolumeEntry&) = delete;
        VolumeEntry& operator=(VolumeEntry&&)      = delete;

        std::mutex mutex;
        std::condition_variable cv;
        uint32_t inUse = 0;
    };

    class Permit final
    {
    public:
        Permit() = default;

        explicit Permit(std::shared_ptr<VolumeEntry> entry) noexcept : _entry(std::move(entry))
        {
        }

        Permit(const Permit&)            = delete;
        Permit& operator=(const Permit&) = delete;

        Permit(Permit&& other) noexcept
        {
            _entry = std::move(other._entry);
        }

        Permit& operator=(Permit&& other) noexcept
        {
            if (this != &other)
            {
                Release();
                _entry = std::move(other._entry);
            }
            return *this;
        }

        ~Permit()
        {
            Release();
        }

        explicit operator bool() const noexcept
        {
            return static_cast<bool>(_entry);
        }

    private:
        void Release() noexcept
        {
            if (! _entry)
            {
                return;
            }

            {
                std::scoped_lock lock(_entry->mutex);
                if (_entry->inUse > 0)
                {
                    _entry->inUse -= 1;
                }
            }

            _entry->cv.notify_one();
            _entry.reset();
        }

        std::shared_ptr<VolumeEntry> _entry;
    };

    Permit AcquireForPath(const std::filesystem::path& path, std::stop_token stopToken) noexcept
    {
        const std::wstring volumeKey = TryGetVolumeKey(path);
        return AcquireForKey(volumeKey.empty() ? L"*" : volumeKey, stopToken);
    }

    Permit AcquireForKey(const std::wstring& key, std::stop_token stopToken) noexcept
    {
        if (_shutdownRequested.load(std::memory_order_acquire))
        {
            return {};
        }

        std::shared_ptr<VolumeEntry> entry;
        {
            std::scoped_lock lock(_mutex);
            if (_shutdownRequested.load(std::memory_order_relaxed))
            {
                return {};
            }
            auto it = _byVolume.find(key);
            if (it == _byVolume.end())
            {
                entry = std::make_shared<VolumeEntry>();
                _byVolume.emplace(key, entry);
            }
            else
            {
                entry = it->second;
            }
        }

        if (! entry)
        {
            return {};
        }

        std::unique_lock lock(entry->mutex);
        while (! stopToken.stop_requested() && ! _shutdownRequested.load(std::memory_order_acquire))
        {
            const uint32_t limit = GetMaxConcurrentScansPerVolume();
            if (entry->inUse < limit)
            {
                entry->inUse += 1;
                return Permit(entry);
            }

            entry->cv.wait_for(lock, std::chrono::milliseconds(50));
        }

        return {};
    }

    void Shutdown() noexcept
    {
        _shutdownRequested.store(true, std::memory_order_release);
        std::vector<std::shared_ptr<VolumeEntry>> entries;
        {
            std::scoped_lock lock(_mutex);
            entries.reserve(_byVolume.size());
            for (const auto& item : _byVolume)
            {
                if (item.second)
                {
                    entries.push_back(item.second);
                }
            }
            _byVolume.clear();
        }

        for (const auto& entry : entries)
        {
            entry->cv.notify_all();
        }
    }

private:
    static std::wstring TryGetVolumeKey(const std::filesystem::path& path) noexcept
    {
        if (path.empty())
        {
            return {};
        }

        std::array<wchar_t, 1024> buffer{};
        if (GetVolumePathNameW(path.c_str(), buffer.data(), static_cast<DWORD>(buffer.size())) != 0)
        {
            return std::wstring(buffer.data());
        }

        const std::wstring pathText = path.wstring();
        if (pathText.size() >= 2u && pathText[1] == L':')
        {
            wchar_t drive[4]{};
            drive[0] = pathText[0];
            drive[1] = L':';
            drive[2] = L'\\';
            drive[3] = L'\0';
            return std::wstring(drive);
        }

        return {};
    }

private:
    std::mutex _mutex;
    std::unordered_map<std::wstring, std::shared_ptr<VolumeEntry>> _byVolume;
    std::atomic_bool _shutdownRequested{false};
};

std::mutex g_scanSchedulerMutex;
std::shared_ptr<ScanScheduler> g_scanScheduler;
bool g_scanSchedulerShutdown = false;
std::atomic_uint64_t g_scanWorkerUnloadGates{0u};
constexpr int kViewerSpaceModuleAnchor = 0;

std::shared_ptr<ScanScheduler> GetScanScheduler() noexcept
{
    std::scoped_lock lock(g_scanSchedulerMutex);
    if (g_scanSchedulerShutdown)
    {
        return {};
    }
    if (! g_scanScheduler)
    {
        g_scanScheduler = std::make_shared<ScanScheduler>();
    }
    return g_scanScheduler;
}

void ShutdownScanScheduler() noexcept
{
    std::shared_ptr<ScanScheduler> scheduler;
    {
        std::scoped_lock lock(g_scanSchedulerMutex);
        g_scanSchedulerShutdown = true;
        scheduler               = std::move(g_scanScheduler);
    }

    if (scheduler)
    {
        scheduler->Shutdown();
    }
}

struct ScanResultCacheKey final
{
    std::wstring rootKey;
    uint32_t topFilesPerDirectory = 0;
};

struct ScanResultCacheNode final
{
    uint32_t id                   = 0;
    uint32_t parentId             = 0;
    bool isDirectory              = false;
    bool isSynthetic              = false;
    uint8_t scanState             = 0;
    uint8_t flags                 = 0u;
    uint64_t totalBytes           = 0;
    int64_t lastWriteTime         = 0;
    uint64_t observedOwnFileBytes = 0;
    uint64_t observedChildBytes   = 0;
    uint32_t childrenStart        = 0;
    uint32_t childrenCount        = 0;
    uint32_t childrenCapacity     = 0;
    uint64_t aggregateBytes       = 0;
    uint64_t aggregateFolders     = 0;
    uint64_t aggregateFiles       = 0;
    uint32_t nameOffset           = 0;
    uint32_t nameLength           = 0;
};

constexpr uint8_t kScanResultCacheNodeContentNotMeasured = 0x01u;
constexpr uint8_t kScanResultCacheNodeKnownFlags         = kScanResultCacheNodeContentNotMeasured;

struct ScanResultCacheFileRecord final
{
    uint32_t id           = 0;
    uint32_t parentId     = 0;
    uint64_t bytes        = 0;
    int64_t lastWriteTime = 0;
    uint32_t nameOffset   = 0;
    uint32_t nameLength   = 0;
};

struct ScanResultSnapshot final
{
    uint32_t topFilesPerDirectory     = 0;
    uint64_t scannedBytes             = 0;
    uint64_t scannedFolders           = 0;
    uint64_t scannedFiles             = 0;
    uint64_t failedFolders            = 0;
    uint32_t firstPreviewPriorityJobs = 0u;
    std::vector<wchar_t> names;
    std::vector<ScanResultCacheNode> nodes;
    std::vector<ScanResultCacheFileRecord> fileRecords;
    std::vector<uint32_t> childrenArena;

    [[nodiscard]] uint64_t AllocatedCapacityBytes() const noexcept
    {
        uint64_t bytes         = static_cast<uint64_t>(sizeof(ScanResultSnapshot));
        const auto addCapacity = [&bytes](size_t capacity, size_t elementSize) noexcept
        {
            if (capacity > (std::numeric_limits<uint64_t>::max() - bytes) / elementSize)
            {
                bytes = std::numeric_limits<uint64_t>::max();
                return;
            }
            bytes += static_cast<uint64_t>(capacity) * static_cast<uint64_t>(elementSize);
        };
        addCapacity(names.capacity(), sizeof(wchar_t));
        addCapacity(nodes.capacity(), sizeof(ScanResultCacheNode));
        addCapacity(fileRecords.capacity(), sizeof(ScanResultCacheFileRecord));
        addCapacity(childrenArena.capacity(), sizeof(uint32_t));
        return bytes;
    }
};

enum class ScanCacheRestorePhase : uint8_t
{
    None,
    ValidateNodes,
    ValidateFiles,
    ValidateChildren,
    ValidateCompleteness,
    PrepareChildren,
    PrepareNodes,
    PrepareFiles,
    Children,
    Nodes,
    Files,
};

[[nodiscard]] bool IsValidSnapshotNameExtent(const ScanResultSnapshot& snapshot, uint32_t offset, uint32_t length) noexcept
{
    const size_t nameOffset = static_cast<size_t>(offset);
    const size_t nameLength = static_cast<size_t>(length);
    return nameOffset <= snapshot.names.size() && nameLength <= snapshot.names.size() - nameOffset;
}

[[nodiscard]] std::wstring_view ScanSnapshotName(const ScanResultSnapshot& snapshot, uint32_t offset, uint32_t length) noexcept
{
    if (length == 0u)
    {
        return {};
    }
    return std::wstring_view(snapshot.names.data() + offset, length);
}

std::wstring NormalizeRootPathForScanCache(const std::filesystem::path& rootPath) noexcept
{
    if (rootPath.empty())
    {
        return {};
    }

    std::error_code ec;
    const std::filesystem::path absolutePath   = std::filesystem::absolute(rootPath, ec);
    const std::filesystem::path normalizedPath = (ec ? rootPath : absolutePath).lexically_normal();

    std::wstring key = normalizedPath.wstring();
    if (key.empty())
    {
        key = rootPath.wstring();
    }

    std::replace(key.begin(), key.end(), L'/', L'\\');

    while (key.size() > 3 && key.back() == L'\\')
    {
        key.pop_back();
    }

    if (key.size() >= 2 && key[1] == L':' && key[0] >= L'a' && key[0] <= L'z')
    {
        key[0] = static_cast<wchar_t>(key[0] - L'a' + L'A');
    }

    return key;
}

class ScanResultCache final
{
public:
    ScanResultCache()  = default;
    ~ScanResultCache() = default;

    ScanResultCache& operator=(ScanResultCache&)  = delete;
    ScanResultCache(ScanResultCache&)             = delete;
    ScanResultCache& operator=(ScanResultCache&&) = delete;
    ScanResultCache(ScanResultCache&&)            = delete;

    void Clear() noexcept
    {
        std::scoped_lock lock(_mutex);
        _entries.clear();
    }

    void Shutdown() noexcept
    {
        Clear();
    }

    void TrimTo(uint32_t maxEntries) noexcept
    {
        const size_t limit = static_cast<size_t>(std::min<uint32_t>(maxEntries, 64u));
        std::scoped_lock lock(_mutex);
        while (_entries.size() > limit)
        {
            _entries.pop_back();
        }
    }

    std::shared_ptr<const ScanResultSnapshot> TryGet(const ScanResultCacheKey& key) noexcept
    {
        if (! g_cacheEnabled.load(std::memory_order_acquire))
        {
            return {};
        }

        const uint32_t maxEntries = g_cacheMaxEntries.load(std::memory_order_acquire);
        if (maxEntries == 0)
        {
            return {};
        }

        const uint32_t ttlSeconds = g_cacheTtlSeconds.load(std::memory_order_acquire);
        const auto now            = std::chrono::steady_clock::now();

        std::scoped_lock lock(_mutex);
        PurgeExpiredLocked(now, ttlSeconds);

        for (size_t i = 0; i < _entries.size(); ++i)
        {
            // A result retained with at least the requested file-candidate budget is reusable: the budget follows the window area,
            // so matching it exactly would turn every resize into a miss.
            Entry& entry = _entries[i];
            if (entry.key.topFilesPerDirectory >= key.topFilesPerDirectory && entry.key.rootKey == key.rootKey)
            {
                entry.lastUsed                                     = now;
                std::shared_ptr<const ScanResultSnapshot> snapshot = entry.snapshot;

                if (i != 0)
                {
                    std::rotate(_entries.begin(), _entries.begin() + static_cast<ptrdiff_t>(i), _entries.begin() + static_cast<ptrdiff_t>(i) + 1);
                }

                return snapshot;
            }
        }

        return {};
    }

    bool Store(ScanResultCacheKey key, std::shared_ptr<ScanResultSnapshot> snapshot) noexcept
    {
        if (! snapshot)
        {
            return false;
        }
        if (! g_cacheEnabled.load(std::memory_order_acquire))
        {
            return false;
        }

        const uint32_t maxEntries = g_cacheMaxEntries.load(std::memory_order_acquire);
        if (maxEntries == 0)
        {
            return false;
        }

        const uint32_t ttlSeconds = g_cacheTtlSeconds.load(std::memory_order_acquire);
        const auto now            = std::chrono::steady_clock::now();

        std::scoped_lock lock(_mutex);
        PurgeExpiredLocked(now, ttlSeconds);

        for (size_t i = 0; i < _entries.size(); ++i)
        {
            if (_entries[i].key.rootKey == key.rootKey)
            {
                _entries.erase(_entries.begin() + static_cast<ptrdiff_t>(i));
                break;
            }
        }

        const size_t limit = static_cast<size_t>(std::min<uint32_t>(maxEntries, 64u));
        if (_entries.capacity() < limit)
        {
            _entries.reserve(limit);
        }

        const uint64_t snapshotBytes = snapshot->AllocatedCapacityBytes();
        const size_t keyCapacity     = key.rootKey.capacity();
        if (keyCapacity == std::numeric_limits<size_t>::max())
        {
            return false;
        }
        const uint64_t keyBytes = static_cast<uint64_t>(keyCapacity + 1u) * static_cast<uint64_t>(sizeof(wchar_t));
        if (keyBytes > std::numeric_limits<uint64_t>::max() - snapshotBytes)
        {
            return false;
        }
        const uint64_t additionalBytes  = snapshotBytes + keyBytes;
        const uint64_t maxResidentBytes = MaxScanCacheResidentBytes();
        if (additionalBytes > maxResidentBytes)
        {
            return false;
        }

        while (_entries.size() >= limit || ResidentBytesLocked() > maxResidentBytes - additionalBytes)
        {
            const auto unpinned = std::find_if(_entries.rbegin(), _entries.rend(), [](const Entry& candidate) { return candidate.snapshot.use_count() == 1; });
            if (unpinned == _entries.rend())
            {
                return false;
            }

            const size_t index = static_cast<size_t>(std::distance(_entries.begin(), std::prev(unpinned.base())));
            _entries.erase(_entries.begin() + static_cast<ptrdiff_t>(index));
        }

        if (additionalBytes > maxResidentBytes || ResidentBytesLocked() > maxResidentBytes - additionalBytes)
        {
            return false;
        }

        Entry entry;
        entry.key      = std::move(key);
        entry.snapshot = std::move(snapshot);
        entry.lastUsed = now;
        _entries.push_back(std::move(entry));
        std::rotate(_entries.begin(), _entries.end() - 1, _entries.end());
        return true;
    }

    void DiscardIfMatches(const ScanResultCacheKey& key, const std::shared_ptr<const ScanResultSnapshot>& snapshot) noexcept
    {
        if (! snapshot)
        {
            return;
        }

        std::scoped_lock lock(_mutex);
        const auto entry = std::find_if(
            _entries.begin(), _entries.end(), [&](const Entry& candidate) { return candidate.key.rootKey == key.rootKey && candidate.snapshot == snapshot; });
        if (entry != _entries.end())
            _entries.erase(entry);
    }

    [[nodiscard]] uint64_t ResidentBytes() noexcept
    {
        std::scoped_lock lock(_mutex);
        return ResidentBytesLocked();
    }

private:
    struct Entry final
    {
        ScanResultCacheKey key;
        std::shared_ptr<const ScanResultSnapshot> snapshot;
        std::chrono::steady_clock::time_point lastUsed;
    };

    [[nodiscard]] uint64_t ResidentBytesLocked() const noexcept
    {
        uint64_t bytes      = static_cast<uint64_t>(_entries.capacity()) * static_cast<uint64_t>(sizeof(Entry));
        const auto addBytes = [&bytes](uint64_t amount) noexcept
        {
            if (amount > std::numeric_limits<uint64_t>::max() - bytes)
            {
                bytes = std::numeric_limits<uint64_t>::max();
                return;
            }
            bytes += amount;
        };
        for (const Entry& entry : _entries)
        {
            const size_t keyCapacity = entry.key.rootKey.capacity();
            if (keyCapacity == std::numeric_limits<size_t>::max())
            {
                return std::numeric_limits<uint64_t>::max();
            }
            addBytes(static_cast<uint64_t>(keyCapacity + 1u) * static_cast<uint64_t>(sizeof(wchar_t)));
            if (entry.snapshot)
            {
                addBytes(entry.snapshot->AllocatedCapacityBytes());
            }
        }
        return bytes;
    }

    void PurgeExpiredLocked(const std::chrono::steady_clock::time_point now, uint32_t ttlSeconds)
    {
        if (ttlSeconds == 0)
        {
            return;
        }

        const auto ttl = std::chrono::seconds(ttlSeconds);
        for (size_t i = 0; i < _entries.size();)
        {
            if ((now - _entries[i].lastUsed) > ttl)
            {
                _entries.erase(_entries.begin() + static_cast<ptrdiff_t>(i));
                continue;
            }
            i += 1;
        }
    }

    std::mutex _mutex;
    std::vector<Entry> _entries;
};

std::mutex g_scanResultCacheMutex;
std::unique_ptr<ScanResultCache> g_scanResultCache;

ScanResultCache& GetScanResultCache() noexcept
{
    std::scoped_lock lock(g_scanResultCacheMutex);
    if (! g_scanResultCache)
    {
        g_scanResultCache = std::make_unique<ScanResultCache>();
    }
    return *g_scanResultCache;
}

void ShutdownScanResultCache() noexcept
{
    std::scoped_lock lock(g_scanResultCacheMutex);
    if (! g_scanResultCache)
    {
        return;
    }

    g_scanResultCache->Shutdown();
    g_scanResultCache.reset();
}

void ShutdownViewerSpaceModuleStateImpl() noexcept
{
    ShutdownScanResultCache();
    ShutdownScanScheduler();
    DestroyViewerSpaceClassBackgroundBrushState();
}

} // namespace

const char* GetViewerSpaceStaticConfigurationSchema() noexcept
{
    return GetViewerSpaceStaticConfigurationSchemaImpl();
}

void ShutdownViewerSpaceModuleState() noexcept
{
    ShutdownViewerSpaceModuleStateImpl();
}

bool CanUnloadViewerSpaceModuleNow() noexcept
{
    return g_scanWorkerUnloadGates.load(std::memory_order_acquire) == 0u && ViewerSpaceWorkspaceExportsIdle();
}

ViewerSpace::ViewerSpace()
{
    _displayModel = std::make_unique<ModelGeneration>();
    _model        = _displayModel.get();

    _metaId                      = L"builtin/viewer-space";
    _metaShortId                 = L"viewspace";
    _metaName                    = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_NAME);
    _metaDescription             = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_DESCRIPTION);
    _scanInProgressWatermarkText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_WATERMARK_IN_PROGRESS);
    _scanIncompleteWatermarkText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_WATERMARK_SCAN_INCOMPLETE);

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    // Instance defaults precede the host's persisted settings; they must not
    // trim other viewers' shared cache or change their admission policy.
    static_cast<void>(ApplyConfiguration(nullptr, false));
}

ViewerSpace::~ViewerSpace()
{
    AbandonScanWorkers();
    CancelScanCacheBuild();
    for (DeferredDisposalWorker& worker : _deferredDisposalWorkers)
    {
        if (! worker.submitted && worker.storage != nullptr)
        {
            worker.storage->ReleaseStorage();
            SubtractAtomicFloor(_deferredDisposalBytes, worker.bytes);
        }
    }
    _deferredDisposalWorkers.clear();

    _fileSystem.reset();
    _hostPaneExecute.reset();
}

void ViewerSpace::SetHost(IHost* host) noexcept
{
    _hostPaneExecute.reset();
    _hostAlerts.reset();
    _hostViewerActions.reset();

    if (! host)
    {
        return;
    }

    static_cast<void>(host->QueryInterface(__uuidof(IHostAlerts), _hostAlerts.put_void()));
    static_cast<void>(host->QueryInterface(__uuidof(IHostViewerActions), _hostViewerActions.put_void()));

    wil::com_ptr<IHostPaneExecute> paneExecute;
    const HRESULT hr = host->QueryInterface(__uuidof(IHostPaneExecute), paneExecute.put_void());
    if (SUCCEEDED(hr) && paneExecute)
    {
        _hostPaneExecute = std::move(paneExecute);
    }
}

HRESULT STDMETHODCALLTYPE ViewerSpace::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IViewer))
    {
        *ppvObject = static_cast<IViewer*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ViewerSpace::AddRef() noexcept
{
    return _refCount.fetch_add(1) + 1;
}

ULONG STDMETHODCALLTYPE ViewerSpace::Release() noexcept
{
    const ULONG remaining = _refCount.fetch_sub(1) - 1;
    if (remaining == 0)
    {
        delete this;
    }
    return remaining;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = GetViewerSpaceStaticConfigurationSchema();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    return ApplyConfiguration(configurationJsonUtf8, true);
}

HRESULT ViewerSpace::ApplyConfiguration(const char* configurationJsonUtf8, bool publishSharedPolicy) noexcept
{
    uint32_t topFilesPerDirectory        = 96;
    uint32_t scanThreads                 = 1;
    uint32_t maxConcurrentScansPerVolume = 1;
    bool cacheEnabled                    = true;
    uint32_t cacheTtlSeconds             = 60;
    uint32_t cacheMaxEntries             = 1;
    ViewMode viewMode                    = ViewMode::Explorer;
    bool analysisPanelVisible            = true;
    float analysisPanelRatio             = 0.33f;
    float analysisBottomRatio            = 0.62f;
    bool inspectorVisible                = true;
    float inspectorRatio                 = 0.38f;
    bool compactRows                     = false;
    uint32_t colorMode                   = 0u;
    bool fileCountWeight                 = false;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        const std::string_view utf8(configurationJsonUtf8);
        if (! utf8.empty())
        {
            yyjson_doc* doc = yyjson_read(utf8.data(), utf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
            if (doc)
            {
                auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

                yyjson_val* root = yyjson_doc_get_root(doc);
                if (root && yyjson_is_obj(root))
                {
                    if (const auto value = yyjson_obj_get(root, "viewMode"); yyjson_is_str(value))
                    {
                        viewMode = std::string_view(yyjson_get_str(value)) == "analyst" ? ViewMode::Analyst : ViewMode::Explorer;
                    }
                    auto readBoolean = [](yyjson_val* value, bool fallback) noexcept
                    {
                        if (yyjson_is_bool(value))
                        {
                            return yyjson_get_bool(value);
                        }
                        if (yyjson_is_str(value))
                        {
                            const std::string_view text(yyjson_get_str(value));
                            if (text == "1" || text == "true" || text == "on")
                            {
                                return true;
                            }
                            if (text == "0" || text == "false" || text == "off")
                            {
                                return false;
                            }
                        }
                        return fallback;
                    };
                    fileCountWeight = readBoolean(yyjson_obj_get(root, "fileCountWeight"), false);
                    if (const auto value = yyjson_obj_get(root, "colorMode"); yyjson_is_str(value))
                    {
                        const std::string_view mode(yyjson_get_str(value));
                        colorMode = mode == "type" ? 1u : mode == "age" ? 2u : mode == "newest" ? 3u : 0u;
                    }
                    analysisPanelVisible = readBoolean(yyjson_obj_get(root, "analysisPanelVisible"), analysisPanelVisible);
                    if (const auto value = yyjson_obj_get(root, "analysisPanelRatio"); yyjson_is_num(value))
                    {
                        const double ratio = yyjson_get_num(value);
                        if (std::isfinite(ratio))
                        {
                            analysisPanelRatio = static_cast<float>(std::clamp(ratio, 0.2, 0.65));
                        }
                    }
                    const auto readRatio = [root](const char* key, float fallback, double maximum) noexcept
                    {
                        const auto value   = yyjson_obj_get(root, key);
                        const double ratio = yyjson_is_num(value) ? yyjson_get_num(value) : static_cast<double>(fallback);
                        return std::isfinite(ratio) ? static_cast<float>(std::clamp(ratio, 0.20, maximum)) : fallback;
                    };
                    analysisBottomRatio  = readRatio("analysisBottomRatio", analysisBottomRatio, 0.80);
                    inspectorRatio       = readRatio("inspectorRatio", inspectorRatio, 0.70);
                    inspectorVisible     = readBoolean(yyjson_obj_get(root, "inspectorVisible"), inspectorVisible);
                    compactRows          = readBoolean(yyjson_obj_get(root, "compactRows"), compactRows);
                    yyjson_val* topFiles = yyjson_obj_get(root, "topFilesPerDirectory");
                    if (topFiles && yyjson_is_int(topFiles))
                    {
                        const int64_t value = yyjson_get_int(topFiles);
                        if (value >= 0)
                        {
                            topFilesPerDirectory = static_cast<uint32_t>(std::min<int64_t>(value, kAdaptiveFileCandidateCeiling));
                        }
                    }

                    yyjson_val* threads = yyjson_obj_get(root, "scanThreads");
                    if (threads && yyjson_is_int(threads))
                    {
                        const int64_t value = yyjson_get_int(threads);
                        if (value > 0)
                        {
                            scanThreads = static_cast<uint32_t>(std::clamp<int64_t>(value, 1, 16));
                        }
                    }

                    yyjson_val* concurrent = yyjson_obj_get(root, "maxConcurrentScansPerVolume");
                    if (concurrent && yyjson_is_int(concurrent))
                    {
                        const int64_t value = yyjson_get_int(concurrent);
                        if (value > 0)
                        {
                            maxConcurrentScansPerVolume = static_cast<uint32_t>(std::clamp<int64_t>(value, 1, 8));
                        }
                    }

                    yyjson_val* cache = yyjson_obj_get(root, "cacheEnabled");
                    if (cache)
                    {
                        if (yyjson_is_str(cache))
                        {
                            const char* value = yyjson_get_str(cache);
                            if (value != nullptr)
                            {
                                cacheEnabled = (strcmp(value, "1") == 0) || (strcmp(value, "true") == 0) || (strcmp(value, "on") == 0);
                            }
                        }
                        else if (yyjson_is_bool(cache))
                        {
                            cacheEnabled = yyjson_get_bool(cache);
                        }
                    }

                    yyjson_val* ttl = yyjson_obj_get(root, "cacheTtlSeconds");
                    if (ttl && yyjson_is_int(ttl))
                    {
                        const int64_t value = yyjson_get_int(ttl);
                        if (value >= 0)
                        {
                            cacheTtlSeconds = static_cast<uint32_t>(std::min<int64_t>(value, 3600));
                        }
                    }

                    yyjson_val* maxEntries = yyjson_obj_get(root, "cacheMaxEntries");
                    if (maxEntries && yyjson_is_int(maxEntries))
                    {
                        const int64_t value = yyjson_get_int(maxEntries);
                        if (value >= 0)
                        {
                            cacheMaxEntries = static_cast<uint32_t>(std::min<int64_t>(value, 16));
                        }
                    }
                }
            }
        }
    }

    _analysisHost.CancelResize();
    _config.topFilesPerDirectory        = topFilesPerDirectory;
    _config.scanThreads                 = scanThreads;
    _config.maxConcurrentScansPerVolume = maxConcurrentScansPerVolume;
    _config.cacheEnabled                = cacheEnabled;
    _config.cacheTtlSeconds             = cacheTtlSeconds;
    _config.cacheMaxEntries             = cacheMaxEntries;
    _config.viewMode                    = viewMode;
    _config.analysisPanelVisible        = analysisPanelVisible;
    _config.analysisPanelRatio          = analysisPanelRatio;
    _config.analysisBottomRatio         = analysisBottomRatio;
    _config.inspectorVisible            = inspectorVisible;
    _config.inspectorRatio              = inspectorRatio;
    if (_hWnd)
    {
        _navigationHost.SetMode(viewMode == ViewMode::Analyst);
        _analysisHost.SetCompactRows(compactRows);
        _analysisHost.SetInspectorLayout(inspectorVisible, inspectorRatio);
        InvalidateAnalysis();
        LayoutAnalysisHost();
        CancelLayoutBuild();
        _layoutDirty = true;
        InvalidateStaticTreemapCache();
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    _config.compactRows     = compactRows;
    _config.colorMode       = colorMode;
    _config.fileCountWeight = fileCountWeight;
    if (_hWnd)
        UpdateWorkspaceScope();
    _layoutCandidateCache.clear();

    if (publishSharedPolicy)
    {
        g_maxConcurrentScansPerVolume.store(_config.maxConcurrentScansPerVolume, std::memory_order_release);
        g_cacheEnabled.store(_config.cacheEnabled, std::memory_order_release);
        g_cacheTtlSeconds.store(_config.cacheTtlSeconds, std::memory_order_release);
        g_cacheMaxEntries.store(_config.cacheMaxEntries, std::memory_order_release);

        if (! _config.cacheEnabled || _config.cacheMaxEntries == 0)
        {
            GetScanResultCache().Clear();
        }
        else
        {
            GetScanResultCache().TrimTo(_config.cacheMaxEntries);
        }
    }

    UpdateConfigurationJson();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    if (_configurationJson.empty())
    {
        *configurationJsonUtf8 = nullptr;
        return S_OK;
    }

    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    const bool isDefault = _config.topFilesPerDirectory == 96u && _config.scanThreads == 1u && _config.maxConcurrentScansPerVolume == 1u &&
                           _config.cacheEnabled && _config.cacheTtlSeconds == 60u && _config.cacheMaxEntries == 1u && _config.viewMode == ViewMode::Explorer &&
                           _config.analysisPanelVisible && _config.analysisPanelRatio == 0.33f && _config.analysisBottomRatio == 0.62f &&
                           _config.inspectorVisible && _config.inspectorRatio == 0.38f && ! _config.compactRows && ! _config.colorMode &&
                           ! _config.fileCountWeight;
    *pSomethingToSave    = isDefault ? FALSE : TRUE;
    return S_OK;
}

ATOM ViewerSpace::RegisterWndClass(HINSTANCE instance) noexcept
{
    if (g_viewerSpaceClassAtom)
    {
        g_viewerSpaceClassBackgroundBrush.classRegistered = true;
        return g_viewerSpaceClassAtom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize              = sizeof(wc);
    wc.lpfnWndProc         = &ViewerSpace::WndProcThunk;
    wc.hInstance           = instance;
    wc.lpszClassName       = kClassName;
    wc.hCursor             = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground       = GetActiveViewerSpaceClassBackgroundBrush();
    wc.style               = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    g_viewerSpaceClassAtom = RegisterClassExW(&wc);
    if (g_viewerSpaceClassAtom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            g_viewerSpaceClassAtom = 1;
        }
    }
    if (g_viewerSpaceClassAtom)
    {
        g_viewerSpaceClassBackgroundBrush.classRegistered = true;
    }
    return g_viewerSpaceClassAtom;
}

void ViewerSpace::UnregisterWndClassIfIdle() noexcept
{
    if (g_viewerSpaceWindowCount.load(std::memory_order_acquire) != 0 || ! g_viewerSpaceClassAtom)
    {
        return;
    }

    if (UnregisterClassW(kClassName, g_hInstance) == 0)
    {
        return;
    }

    g_viewerSpaceClassAtom                            = 0;
    g_viewerSpaceClassBackgroundBrush.classRegistered = false;
    g_viewerSpaceClassBackgroundBrush.pendingBrush.reset();
    g_viewerSpaceClassBackgroundBrush.pendingColor = CLR_INVALID;
    g_viewerSpaceClassBackgroundBrush.activeBrush.reset();
    g_viewerSpaceClassBackgroundBrush.activeColor = CLR_INVALID;
}

LRESULT CALLBACK ViewerSpace::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<ViewerSpace*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            InitPostedPayloadWindow(hwnd);
            g_viewerSpaceWindowCount.fetch_add(1, std::memory_order_acq_rel);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    auto* self = reinterpret_cast<ViewerSpace*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerSpace::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WndMsg::kViewerSpaceCacheRestoreContinue:
        {
            auto continuation = TakeMessagePayload<ScanCacheRestoreContinuation>(hwnd, msg, lp);
            if (! continuation)
            {
                return 0;
            }
            auto clearContinuationPending = wil::scope_exit([this]() noexcept { _scanCacheRestoreContinuationPending = false; });
            if (! _scanCacheRestoreSnapshot || continuation->generation != _scanCacheRestoreGeneration ||
                continuation->generation != _scanGeneration.load(std::memory_order_acquire))
            {
                Debug::Perf::EmitCounter(L"viewer.space.model.cache_restore_continuation_stale");
                return 0;
            }
            Debug::Perf::EmitCounter(L"viewer.space.model.cache_restore_continuation_dispatched");
            ContinueScanCacheRestore();
            return 0;
        }
#ifdef ENABLE_TESTS
        case WndMsg::kViewerDebugGetMenuAttachmentSnapshot:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerMenuAttachmentDebugSnapshot*>(lp);
            if (! snapshot)
                return FALSE;
            *snapshot = {_menuAttachResult.inputMenu,
                         _menuAttachResult.menuAfterAttach,
                         static_cast<unsigned>(_menuAttachResult.outcome),
                         _menuAttachResult.fallbackError,
                         _menuAttachResult.attempted,
                         _menuAttachResult.attached};
            return TRUE;
        }
        case WndMsg::kViewerDebugGetNativeMenuModelSnapshot:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerNativeMenuModelDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            const UINT queryCommandId    = snapshot->queryCommandId;
            *snapshot                    = {};
            snapshot->queryCommandId     = queryCommandId;
            snapshot->hasHiddenMenuModel = _menuHandle != nullptr;
            snapshot->ownerDrawItemCount = CountOwnerDrawMenuItems(_menuHandle.get());
            if (queryCommandId != 0u)
            {
                wil::unique_hmenu rootMenu(Localization::LoadMenuResource(GetModuleHandleW(nullptr), kHostFolderViewItemContextMenuResourceId));
                HMENU menu = rootMenu ? GetSubMenu(rootMenu.get(), 0) : nullptr;
                UpdateHostFolderViewContextCommandState(menu, false, false, false);

                MENUITEMINFOW itemInfo{};
                itemInfo.cbSize = sizeof(itemInfo);
                itemInfo.fMask  = MIIM_STATE;
                if (menu && GetMenuItemInfoW(menu, queryCommandId, FALSE, &itemInfo) != FALSE)
                {
                    snapshot->queryCommandPresent = true;
                    snapshot->queryCommandEnabled = (itemInfo.fState & MFS_DISABLED) == 0u;
                    snapshot->queryCommandChecked = (itemInfo.fState & MFS_CHECKED) != 0u;
                }
            }
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugGetTooltipSnapshot:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerSpaceTooltipDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot                        = {};
            snapshot->tooltipNodeId          = _tooltipNodeId;
            snapshot->tooltipTextLength      = _tooltipText.size();
            snapshot->tooltipAnchorXDip      = _tooltipAnchorDip.x;
            snapshot->tooltipAnchorYDip      = _tooltipAnchorDip.y;
            snapshot->tooltipMaxWidthDip     = kTooltipMaxWidthDip;
            snapshot->tooltipMaxHeightDip    = kTooltipMaxHeightDip;
            snapshot->tooltipPaintCount      = _tooltipPaintCount;
            snapshot->tooltipPaintedBoundsPx = {PxFromDip(_lastTooltipPaintedBoundsDip.left),
                                                PxFromDip(_lastTooltipPaintedBoundsDip.top),
                                                PxFromDip(_lastTooltipPaintedBoundsDip.right),
                                                PxFromDip(_lastTooltipPaintedBoundsDip.bottom)};
            const auto tooltipMapBounds      = GetTreemapBounds();
            snapshot->treemapBoundsPx        = {
                PxFromDip(tooltipMapBounds.left), PxFromDip(tooltipMapBounds.top), PxFromDip(tooltipMapBounds.right), PxFromDip(tooltipMapBounds.bottom)};
            snapshot->hasPaintedTooltipBounds       = _tooltipNodeId != 0u && RectArea(_lastTooltipPaintedBoundsDip) > 0.0f;
            snapshot->hasRenderTarget               = _renderTarget != nullptr;
            snapshot->hasTooltipFormat              = _tooltipFormat != nullptr;
            snapshot->hoverNodeId                   = _hoverNodeId;
            snapshot->hasLastMouseMoveClientPoint   = _hasLastMouseMoveClientPoint;
            snapshot->lastMouseMoveClientX          = _lastMouseMoveClientPoint.x;
            snapshot->lastMouseMoveClientY          = _lastMouseMoveClientPoint.y;
            snapshot->hasLastContextMenuScreenPoint = _debugHasLastContextMenuScreenPoint;
            snapshot->lastContextMenuScreenX        = _debugLastContextMenuScreenPoint.x;
            snapshot->lastContextMenuScreenY        = _debugLastContextMenuScreenPoint.y;
            snapshot->lastContextMenuHitNodeId      = _debugLastContextMenuHitNodeId;
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugCompareHitTesting:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerSpaceHitTestDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot                                    = {};
            static constexpr uint32_t kMaxHitTestSamples = 256u;
            const size_t stride = _drawItems.size() > kMaxHitTestSamples ? ((_drawItems.size() + kMaxHitTestSamples - 1u) / kMaxHitTestSamples) : 1u;
            for (size_t index = 0; index < _drawItems.size() && snapshot->sampleCount < kMaxHitTestSamples; index += stride)
            {
                const DrawItem& item = _drawItems[index];
                if (item.lod == DrawItem::Lod::Culled)
                {
                    continue;
                }

                float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
                gap       = std::clamp(gap, 0.5f, kItemGapDip);

                D2D1_RECT_F rc = item.currentRect;
                rc.left += gap;
                rc.top += gap;
                rc.right -= gap;
                rc.bottom -= gap;
                if (rc.right <= rc.left || rc.bottom <= rc.top || RectArea(rc) < kMinHitAreaDip2)
                {
                    continue;
                }

                const float xDip                        = (rc.left + rc.right) * 0.5f;
                const float yDip                        = (rc.top + rc.bottom) * 0.5f;
                const std::optional<uint32_t> gridHit   = HitTestTreemap(xDip, yDip);
                const uint32_t gridCandidates           = _lastHitTestCandidatesChecked;
                uint32_t linearCandidates               = 0u;
                const std::optional<uint32_t> linearHit = HitTestTreemapLinear(xDip, yDip, linearCandidates);

                snapshot->sampleCount += 1u;
                snapshot->maxGridCandidatesChecked   = std::max(snapshot->maxGridCandidatesChecked, gridCandidates);
                snapshot->maxLinearCandidatesChecked = std::max(snapshot->maxLinearCandidatesChecked, linearCandidates);
                if (gridHit.has_value())
                {
                    snapshot->gridHitCount += 1u;
                }
                if (linearHit.has_value())
                {
                    snapshot->linearHitCount += 1u;
                }

                const bool mismatch =
                    gridHit.has_value() != linearHit.has_value() || (gridHit.has_value() && linearHit.has_value() && gridHit.value() != linearHit.value());
                if (mismatch)
                {
                    snapshot->mismatchCount += 1u;
                    snapshot->lastMismatchGridNodeId   = gridHit.value_or(0u);
                    snapshot->lastMismatchLinearNodeId = linearHit.value_or(0u);
                }
            }

            Debug::Perf::Emit(L"viewer.space.hit_test.debug_samples",
                              L"grid-vs-linear",
                              0u,
                              snapshot->sampleCount,
                              snapshot->mismatchCount,
                              snapshot->mismatchCount == 0u ? S_OK : E_FAIL);
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugGetPerfSnapshot:
        case WndMsg::kViewerSpaceDebugForcePerfSample:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerSpacePerfDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot              = {};
            snapshot->rendererMode = _renderer.d2dContext
                                         ? WndMsg::ViewerSpacePerfRendererMode::DeviceContext
                                         : (_renderTarget ? WndMsg::ViewerSpacePerfRendererMode::HwndRenderTarget : WndMsg::ViewerSpacePerfRendererMode::None);
            snapshot->scanState    = [&]() noexcept
            {
                switch (_overallState)
                {
                    case ScanState::Queued: return WndMsg::ViewerSpacePerfScanState::Queued;
                    case ScanState::Scanning: return WndMsg::ViewerSpacePerfScanState::Scanning;
                    case ScanState::Done: return WndMsg::ViewerSpacePerfScanState::Done;
                    case ScanState::Error: return WndMsg::ViewerSpacePerfScanState::Error;
                    case ScanState::Canceled: return WndMsg::ViewerSpacePerfScanState::Canceled;
                    case ScanState::NotStarted:
                    default: return WndMsg::ViewerSpacePerfScanState::NotStarted;
                }
            }();
            snapshot->hasRenderTarget = _renderTarget != nullptr && (_renderer.d2dContext == nullptr || _renderer.targetBitmap != nullptr);

            for (const Node& node : _model->nodes)
            {
                if (node.id == 0)
                {
                    continue;
                }

                if (node.isSynthetic)
                {
                    snapshot->syntheticCount += 1u;
                    snapshot->syntheticBytes += node.totalBytes;
                }
                else if (node.isDirectory)
                {
                    snapshot->realDirectoryCount += 1u;
                    snapshot->aggregateBytes   = ViewerSpaceScan::SaturatingAdd(snapshot->aggregateBytes, node.aggregateBytes);
                    snapshot->aggregateFolders = ViewerSpaceScan::SaturatingAdd(snapshot->aggregateFolders, node.aggregateFolders);
                    if (node.id == _model->rootNodeId)
                    {
                        for (const uint32_t childId : GetRealNodeChildren(node))
                        {
                            const Node* child = TryGetRealNode(childId);
                            if (child != nullptr && child->isDirectory)
                            {
                                snapshot->firstRootDirectoryTotalBytes = child->totalBytes;
                                break;
                            }
                        }
                    }
                    snapshot->aggregateFiles      = ViewerSpaceScan::SaturatingAdd(snapshot->aggregateFiles, node.aggregateFiles);
                    snapshot->childReferenceCount = ViewerSpaceScan::SaturatingAdd(snapshot->childReferenceCount, static_cast<uint64_t>(node.childrenCount));
                }
            }
            for (const auto& entry : _syntheticNodes)
            {
                const Node& node = entry.second;
                snapshot->syntheticCount += 1u;
                snapshot->syntheticBytes = ViewerSpaceScan::SaturatingAdd(snapshot->syntheticBytes, node.totalBytes);
            }
            snapshot->fileCandidateCount = _model->fileRecords.size();
            for (const FileRecord& record : _model->fileRecords)
            {
                snapshot->fileCandidateBytes += record.bytes;
            }
            if (const Node* root = TryGetRealNode(_model->rootNodeId))
            {
                snapshot->rootTotalBytes = root->totalBytes;
            }
            {
                std::scoped_lock lock(_updateMutex);
                snapshot->pendingQueueCount              = _pendingUpdates.size();
                snapshot->pendingQueueBytes              = _pendingUpdateBytes.load(std::memory_order_relaxed);
                snapshot->pendingQueuePeakBytes          = _pendingUpdatePeakBytes.load(std::memory_order_relaxed);
                snapshot->queuePostedCount               = _pendingUpdatePostedCount.load(std::memory_order_relaxed);
                snapshot->queueCoalescedCount            = _pendingUpdateCoalescedCount.load(std::memory_order_relaxed);
                snapshot->queueLockHoldCount             = _pendingUpdateLockHoldCount.load(std::memory_order_relaxed);
                snapshot->queueDrainedCount              = _pendingUpdateDrainedCount.load(std::memory_order_relaxed);
                snapshot->providerBufferCurrentBytes     = _providerBufferCurrentBytes.load(std::memory_order_relaxed);
                snapshot->providerBufferPeakBytes        = _providerBufferPeakBytes.load(std::memory_order_relaxed);
                snapshot->validationScratchCurrentBytes  = _validationScratchCurrentBytes.load(std::memory_order_relaxed);
                snapshot->validationScratchPeakBytes     = _validationScratchPeakBytes.load(std::memory_order_relaxed);
                snapshot->validationScratchRejectedCount = _validationScratchRejectedCount.load(std::memory_order_relaxed);
                snapshot->workerStagingCurrentBytes      = _workerStagingCurrentBytes.load(std::memory_order_relaxed);
                snapshot->workerStagingPeakBytes         = _workerStagingPeakBytes.load(std::memory_order_relaxed);
                snapshot->workerStagingRejectedCount     = _workerStagingRejectedCount.load(std::memory_order_relaxed);
                snapshot->scannerBookkeepingCurrentBytes = _scannerBookkeepingCurrentBytes.load(std::memory_order_relaxed);
                snapshot->scannerBookkeepingPeakBytes    = _scannerBookkeepingPeakBytes.load(std::memory_order_relaxed);
                snapshot->activePathCurrentBytes         = _activePathCurrentBytes.load(std::memory_order_relaxed);
                snapshot->activePathPeakBytes            = _activePathPeakBytes.load(std::memory_order_relaxed);
                snapshot->activePathRejectedCount        = _activePathRejectedCount.load(std::memory_order_relaxed);
            }

            snapshot->drawItemCount = _drawItems.size();
            for (const DrawItem& item : _drawItems)
            {
                if (item.lod == DrawItem::Lod::Culled)
                {
                    snapshot->culledTileCount += 1u;
                }
            }
            snapshot->visibleTileCount             = snapshot->drawItemCount - snapshot->culledTileCount;
            snapshot->staticCacheGeneration        = _staticTreemapCache.generation;
            snapshot->staticCacheHits              = _staticTreemapCache.hits;
            snapshot->staticCacheMisses            = _staticTreemapCache.misses;
            snapshot->staticCacheRecordCount       = _staticTreemapCache.recordCount;
            snapshot->staticCacheBytes             = _staticTreemapCache.bytes;
            snapshot->scanCacheSnapshotBytes       = _lastScanCacheSnapshotBytes;
            snapshot->scanCacheRestoreBytes        = _lastScanCacheRestoreBytes;
            snapshot->scanCacheResidentBytes       = _lastScanCacheResidentBytes;
            snapshot->scanCacheRestoreUs           = _lastScanCacheRestoreUs;
            snapshot->cacheRestorePhase            = _scanCacheRestorePhase;
            snapshot->viewerOwnedBytes             = ViewerOwnedBytes();
            snapshot->layoutTurnCount              = _layoutTurnCount;
            snapshot->layoutMaxTurnUs              = _layoutMaxTurnUs;
            snapshot->layoutBuildActive            = _layoutBuild != nullptr;
            snapshot->layoutCancellationPending    = _cancelledLayoutBuild != nullptr;
            snapshot->layoutCancellationOwnedBytes = _cancelledLayoutBuild != nullptr ? LayoutBuildOwnedBytes(*_cancelledLayoutBuild) : 0u;
            snapshot->layoutDirty                  = _layoutDirty;
            snapshot->summaryStageBytes            = _model->directorySummaryStageBytes;
            snapshot->summaryCommittedCount        = _directorySummaryCommittedCount;
            snapshot->summaryRejectedCount         = _directorySummaryRejectedCount;
            snapshot->summaryChunkMaxBytes         = _directorySummaryChunkMaxBytes.load(std::memory_order_relaxed);
            snapshot->scanPreviewPriorityJobs      = _scanPreviewPriorityJobs;
            snapshot->summaryRootDoneDeferred      = _model->deferredDirectoryDoneStates.contains(_model->rootNodeId);
            snapshot->scanCacheBuildBytes =
                _scanCacheBuildSnapshot ? std::static_pointer_cast<ScanResultSnapshot>(_scanCacheBuildSnapshot)->AllocatedCapacityBytes() : 0u;
            snapshot->analysisRowCount     = _analysisRows.size();
            snapshot->analysisGeneration   = _analysisGeneration;
            snapshot->selectedItemId       = _selectedItemId;
            snapshot->analysisCellRequests = _analysisCellRequests;
            snapshot->analysisMaxTurnUs    = _analysisMaxTurnUs;
            snapshot->analysisDirty        = _analysisDirty;
            snapshot->analysisBuildActive  = _analysisBuild != nullptr;
            snapshot->analysisHwnd         = _analysisHost.GetHwnd();
            snapshot->navigationHwnd       = _navigationHost.GetHwnd();
            snapshot->analysisDarkTheme    = _analysisHost.GetTheme().dark;
            snapshot->navigationDarkTheme  = _navigationHost.GetTheme().dark;
            const auto copyHostDiagnostics =
                [](WndMsg::ViewerSpacePerfDebugSnapshot::ChildHost& target, const ViewerSpaceHosts::HostDiagnostics& source) noexcept
            {
                target.clientWidthPx               = source.clientWidthPx;
                target.clientHeightPx              = source.clientHeightPx;
                target.windowBackgroundArgb        = source.windowBackgroundArgb;
                target.textArgb                    = source.textArgb;
                target.renderCount                 = source.renderCount;
                target.resizeCount                 = source.resizeCount;
                target.resizeFailureCount          = source.resizeFailureCount;
                target.presentFailureCount         = source.presentFailureCount;
                target.prepareD2DFlushFailureCount = source.prepareD2DFlushFailureCount;
                target.invalidateCount             = source.invalidateCount;
                target.presentationMode            = source.presentationMode;
                target.diagnosticsAvailable        = source.diagnosticsAvailable;
                target.hasWindow                   = source.hasWindow;
                target.visible                     = source.visible;
                target.hasRoot                     = source.hasRoot;
                target.hasD2DContext               = source.hasD2DContext;
                target.dark                        = source.dark;
                target.highContrast                = source.highContrast;
            };
            copyHostDiagnostics(snapshot->analysisHost, _analysisHost.GetDiagnostics());
            copyHostDiagnostics(snapshot->navigationHost, _navigationHost.GetDiagnostics());
            const auto analysisRootBounds   = _analysisHost.GetRootBounds();
            snapshot->analysisRootWidthDip  = analysisRootBounds.right - analysisRootBounds.left;
            snapshot->analysisRootHeightDip = analysisRootBounds.bottom - analysisRootBounds.top;
            snapshot->analysisBoundsPx      = {
                PxFromDip(_analysisBounds.left), PxFromDip(_analysisBounds.top), PxFromDip(_analysisBounds.right), PxFromDip(_analysisBounds.bottom)};
            snapshot->inputPresentationCount  = _inputPresentationCount;
            snapshot->lastInputPresentationUs = _lastInputPresentationUs;
            snapshot->lastRenderPresentHr     = _lastRenderPresentHr;
            snapshot->headerStatusResourceId  = _headerStatusId;
            snapshot->refreshPending          = _refreshPending;
            snapshot->refreshRefused          = _refreshRefused;
            snapshot->refreshFailed           = _refreshFailed;
            snapshot->previousModelBytes      = _refreshPending && _model ? _model->AllocatedBytes() : 0u;
            snapshot->candidateModelBytes     = _workingModel ? _workingModel->AllocatedBytes() : 0u;
            snapshot->displayGeneration       = _displayGeneration;
            snapshot->refreshDisposePending   = static_cast<uint32_t>(std::min<size_t>(_deferredDisposalWorkers.size(), std::numeric_limits<uint32_t>::max()));
            snapshot->refreshDisposeMaxUs     = _deferredDisposalSubmitMaxUs.load(std::memory_order_relaxed);
            snapshot->retiredRendererDeviceCount       = RetiredRendererDeviceCount();
            snapshot->lastStaticCacheRecordUs          = _staticTreemapCache.lastRecordUs;
            snapshot->lastPaintUs                      = _lastPaintUs;
            snapshot->lastLayoutUs                     = _lastLayoutUs;
            snapshot->lastDrainUs                      = _lastDrainUs;
            _lastWorkingSetBytes                       = SampleCurrentWorkingSetBytes();
            snapshot->lastWorkingSetBytes              = _lastWorkingSetBytes;
            snapshot->lastTileDrawCount                = _lastTileDrawCount;
            snapshot->lastTextDrawCount                = _lastTextDrawCount;
            snapshot->lastHitTestUs                    = _lastHitTestUs;
            snapshot->lastHitTestCandidatesChecked     = _lastHitTestCandidatesChecked;
            snapshot->effectiveFileCandidateBudget     = _scanTopFilesPerDirectory;
            snapshot->rendererDeviceCreateCount        = _rendererDeviceCreateCount;
            snapshot->swapChainResizeCount             = _swapChainResizeCount;
            snapshot->rendererBrushCreateCount         = _rendererBrushCreateCount;
            snapshot->rendererTextFormatCreateCount    = _rendererTextFormatCreateCount;
            snapshot->swapChainWidthPx                 = _renderer.swapChainWidthPx;
            snapshot->swapChainHeightPx                = _renderer.swapChainHeightPx;
            snapshot->rendererFailureStage             = _rendererFailureStage;
            snapshot->rendererFailureHr                = _rendererFailureHr;
            snapshot->layoutGeneration                 = _layoutGeneration;
            snapshot->presentedLayoutGeneration        = _presentedLayoutGeneration;
            snapshot->firstMeaningfulPaintUs           = _firstMeaningfulPaintUs;
            snapshot->firstPositiveProjectionAttempted = _firstPositiveProjectionAttempted;
            snapshot->hasPublishedPositiveGeometry     = _hasPublishedPositiveGeometry;
            snapshot->hitGridCellCount                 = static_cast<uint32_t>(std::min<size_t>(_hitGrid.cells.size(), std::numeric_limits<uint32_t>::max()));
            snapshot->hitGridMaxCandidatesPerCell      = _hitGrid.maxCandidatesPerCell;
            const D2D1_RECT_F treemapBounds            = GetTreemapBounds();
            snapshot->treemapWidthDip                  = std::max(0.0f, treemapBounds.right - treemapBounds.left);
            snapshot->treemapHeightDip                 = std::max(0.0f, treemapBounds.bottom - treemapBounds.top);
            snapshot->scannedFolders                   = _scanProgressFolders;
            snapshot->scannedFiles                     = _scanProgressFiles;
            snapshot->scannedBytes                     = _scanProgressBytes;
            snapshot->scanFailedFolders                = _scanFailedFolders;
            snapshot->scanEffectiveThreads             = _scanEffectiveThreads;
            snapshot->localeInvalidations              = _debugLocaleInvalidations;
            snapshot->animationTimerArmed              = _animationTimerArmed;
            snapshot->viewNodeId                       = _model->viewNodeId;
            if (const Node* viewNode = TryGetRealNode(_model->viewNodeId); viewNode != nullptr)
            {
                snapshot->viewNodeParentId       = viewNode->parentId;
                snapshot->viewContentNotMeasured = viewNode->contentNotMeasured;
            }
            snapshot->lastHitTestPath                  = _lastHitTestPath;
            snapshot->layoutAnimationActive            = HasActiveLayoutAnimation(NowSeconds());
            snapshot->lastLButtonDownHitNodeId         = _lastLButtonDownHitNodeId;
            snapshot->childArenaSlots                  = _model->childrenArena.size();
            snapshot->childArenaFreeSlots              = ChildArenaFreeSlots();
            snapshot->modelAcceptedEntries             = _lastModelAcceptedEntries;
            snapshot->modelRejectedEntries             = _lastModelRejectedEntries;
            snapshot->modelCappedDirectories           = _lastModelCappedDirectories;
            snapshot->modelCappedFiles                 = _lastModelCappedFiles;
            snapshot->modelRetainedNameBytes           = _lastModelRetainedNameBytes;
            snapshot->modelRetainedChildReferences     = _lastModelRetainedChildReferences;
            snapshot->modelTraversedDirectories        = _lastModelTraversedDirectories;
            snapshot->modelRetainedDirectoryLimit      = _scanResourcePolicy.maxRetainedDirectories;
            snapshot->modelRetainedFileLimit           = _scanResourcePolicy.maxRetainedFileRecords;
            snapshot->modelChildReferenceLimit         = _scanResourcePolicy.maxChildReferences;
            snapshot->modelChildArenaSlotLimit         = _scanResourcePolicy.maxChildArenaSlots;
            snapshot->modelValidationError             = static_cast<uint32_t>(_lastModelValidationError);
            snapshot->postUpdateInnerGenerationRejects = _debugPostUpdateInnerGenerationRejects.load(std::memory_order_acquire);

            Debug::Perf::EmitValue(L"viewer.space.model.directory_count", static_cast<uint64_t>(snapshot->realDirectoryCount));
            Debug::Perf::EmitValue(L"viewer.space.model.file_candidate_count", static_cast<uint64_t>(snapshot->fileCandidateCount));
            Debug::Perf::EmitValue(L"viewer.space.model.synthetic_count", static_cast<uint64_t>(snapshot->syntheticCount));
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugForceRendererFault:
        {
            const auto mode = static_cast<WndMsg::ViewerSpaceRendererFaultDebugMode>(wp);
            if (mode != WndMsg::ViewerSpaceRendererFaultDebugMode::D2DRecreateTarget && mode != WndMsg::ViewerSpaceRendererFaultDebugMode::DxgiDeviceRemoved)
            {
                return FALSE;
            }

            _debugForcedRendererFault = static_cast<uint32_t>(mode);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
            }
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugSetLayoutWorkUnits: _layoutWorkUnitsPerTurn = std::clamp<uint32_t>(static_cast<uint32_t>(wp), 1u, 32768u); return TRUE;
        case WndMsg::kViewerSpaceDebugGetAnalysisRow:
        {
            auto* query = reinterpret_cast<WndMsg::ViewerSpaceAnalysisRowDebugQuery*>(lp);
            if (! query || query->index >= _analysisRows.size())
            {
                return FALSE;
            }
            const auto item = ResolveAnalysisItem(_analysisRows[query->index]);
            if (! item.has_value())
            {
                return FALSE;
            }
            query->token                 = AnalysisRowToken(query->index);
            query->itemId                = item.value().id;
            query->generation            = _analysisGeneration;
            query->logicalBytes          = item.value().totalBytes;
            query->shownBytes            = ShownBytes(item.value().id, item.value().totalBytes);
            query->excluded              = IsItemExcluded(item.value().id);
            query->cleanupMarked         = std::find(_cleanupMarks.begin(), _cleanupMarks.end(), item.value().id) != _cleanupMarks.end();
            query->directory             = item.value().isDirectory;
            query->other                 = item.value().isSynthetic;
            query->contentNotMeasured    = item.value().contentNotMeasured;
            query->scanState             = static_cast<WndMsg::ViewerSpacePerfScanState>(item.value().scanState);
            const std::wstring_view name = item.value().name;
            const size_t count           = std::min(name.size(), std::size(query->name) - 1u);
            std::copy_n(name.begin(), count, query->name);
            query->name[count] = L'\0';
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugNavigationFolder: return DebugNavigationFolder(reinterpret_cast<WndMsg::ViewerSpaceNavigationFolderDebugQuery*>(lp));
        case WndMsg::kViewerSpaceDebugExport: return DebugQueueWorkspaceExport(reinterpret_cast<const WndMsg::ViewerSpaceExportDebugRequest*>(lp));
        case WndMsg::kViewerSpaceDebugGetWorkspace:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerSpaceWorkspaceDebugSnapshot*>(lp);
            if (! snapshot || ! _model)
                return FALSE;
            *snapshot                   = {};
            snapshot->generation        = _displayGeneration;
            snapshot->discoveredBytes   = _scanProgressBytes;
            const Node* view            = TryGetRealNode(_model->viewNodeId);
            snapshot->originalViewBytes = view ? view->totalBytes : 0u;
            snapshot->shownViewBytes    = view ? ShownBytes(view->id, view->totalBytes) : 0u;
            snapshot->ownedBytes        = WorkspaceOwnedBytes();
            snapshot->exclusions        = _excludedPaths.size();
            snapshot->unresolved        = _workspaceUnresolved;
            snapshot->cleanupMarks      = _cleanupMarks.size();
            snapshot->rebinding         = _workspaceRebinding;
            snapshot->reviewing         = _reviewingCleanup;
            snapshot->exportBusy        = ! ViewerSpaceWorkspaceExportsIdle();
            snapshot->factsReady        = _workspaceFactsReady && _workspaceFactsGeneration == _displayGeneration;
            snapshot->fileCountWeight   = _config.fileCountWeight;
            snapshot->shownFileCount    = snapshot->factsReady && view ? ShownFileCount(view->id) : 0u;
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugDetachAnalysis: _analysisHost.Detach(); return TRUE;
        case WndMsg::kViewerSpaceDebugRefreshAnalysisAccessibility:
            _analysisHost.NotifyRowsChanged(_selectedItemId &&
                                                    FindAnalysisRow((static_cast<uint64_t>(_analysisGeneration) << 32u) | _selectedItemId).has_value()
                                                ? std::optional<uint64_t>((static_cast<uint64_t>(_analysisGeneration) << 32u) | _selectedItemId)
                                                : std::nullopt);
            return TRUE;
        case WndMsg::kViewerSpaceDebugSetAnalysisSort:
            _analysisSortColumn    = std::min<size_t>(static_cast<size_t>(wp), 4u);
            _analysisSortAscending = lp != 0;
            _analysisHost.SetSortSpec({_analysisSortColumn, _analysisSortAscending ? DxUi::SortDirection::Ascending : DxUi::SortDirection::Descending});
            InvalidateAnalysis();
            return TRUE;
        case WndMsg::kViewerSpaceDebugSetAnalysisFilter:
        {
            const auto* filter = reinterpret_cast<const wchar_t*>(lp);
            _analysisFilter    = filter ? filter : L"";
            _analysisHost.SetFilterText(_analysisFilter);
            _analysisFilterDueSeconds = NowSeconds() + 0.15;
            InvalidateAnalysis();
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugPauseNextPostUpdate:
        {
            const auto* control = reinterpret_cast<const WndMsg::ViewerSpacePostUpdatePauseDebugControl*>(lp);
            if (! control || ! control->enteredEvent || ! control->releaseEvent)
            {
                return FALSE;
            }
            _debugPostUpdateEnteredEvent.store(control->enteredEvent, std::memory_order_relaxed);
            _debugPostUpdateReleaseEvent.store(control->releaseEvent, std::memory_order_relaxed);
            _debugPauseNextPostUpdate.store(true, std::memory_order_release);
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugPauseAfterValidationScratch:
        {
            const auto* control = reinterpret_cast<const WndMsg::ViewerSpaceValidationScratchPauseDebugControl*>(lp);
            if (! control || ! control->enteredEvent || ! control->releaseEvent)
            {
                return FALSE;
            }
            _debugValidationScratchEnteredEvent.store(control->enteredEvent, std::memory_order_relaxed);
            _debugValidationScratchReleaseEvent.store(control->releaseEvent, std::memory_order_relaxed);
            _debugPauseAfterValidationScratch.store(true, std::memory_order_release);
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugPauseAfterWorkerStaging:
        {
            const auto* control = reinterpret_cast<const WndMsg::ViewerSpaceValidationScratchPauseDebugControl*>(lp);
            if (! control)
                return FALSE;
            _debugWorkerStagingEnteredEvent.store(control->enteredEvent, std::memory_order_relaxed);
            _debugWorkerStagingReleaseEvent.store(control->releaseEvent, std::memory_order_relaxed);
            _debugPauseAfterWorkerStaging.store(true, std::memory_order_release);
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugGetDrawItem:
        {
            auto* snapshot     = reinterpret_cast<WndMsg::ViewerSpaceDrawItemDebugSnapshot*>(lp);
            const size_t index = static_cast<size_t>(wp);
            if (! snapshot || index >= _drawItems.size())
            {
                return FALSE;
            }

            const DrawItem& item                   = _drawItems[index];
            const std::optional<ResolvedItem> node = ResolveTreemapItem(item.nodeId);
            *snapshot                              = {};
            snapshot->drawItemCount                = _drawItems.size();
            snapshot->dpi                          = _dpi;
            snapshot->headerBottomDip              = GetHeaderBottomDip();
            snapshot->nodeId                       = item.nodeId;
            if (node.has_value())
            {
                snapshot->parentId           = node.value().parentId;
                snapshot->logicalBytes       = node.value().totalBytes;
                snapshot->contentNotMeasured = node.value().contentNotMeasured;
                const size_t count           = std::min(node.value().name.size(), std::size(snapshot->name) - 1u);
                std::copy_n(node.value().name.begin(), count, snapshot->name);
                snapshot->name[count] = L'\0';
            }
            snapshot->depth          = item.depth;
            snapshot->isDirectory    = node.has_value() && node->isDirectory;
            snapshot->isSynthetic    = node.has_value() && node->isSynthetic;
            snapshot->labelHeightDip = item.labelHeightDip;
            snapshot->current        = {item.currentRect.left, item.currentRect.top, item.currentRect.right, item.currentRect.bottom};
            snapshot->target         = {item.targetRect.left, item.targetRect.top, item.targetRect.right, item.targetRect.bottom};
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugShowTooltipOverlay:
        {
            _tooltipNodeId    = 1u;
            _tooltipText      = L"ViewerSpace debug tooltip overlay text that wraps through Direct2D.";
            _tooltipAnchorDip = D2D1::Point2F(40.0f, GetHeaderBottomDip() + 24.0f);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
                UpdateWindow(_hWnd.get());
            }
            return TRUE;
        }
#endif
        case kWorkspaceExportCompleted: HandleWorkspaceExportCompletion(hwnd, lp); return 0;
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_SIZE: OnSize(static_cast<UINT>(LOWORD(lp)), static_cast<UINT>(HIWORD(lp))); return 0;
        case WM_MOVE:
            // A parent can reposition an embedded viewer beneath a stationary
            // cursor without another WM_MOUSEMOVE. Reproject its saved screen point.
            if (_trackingMouse && _hasLastMouseMoveClientPoint)
            {
                EnsureAnimationTimer();
            }
            return 0;
        case WM_DPICHANGED: OnDpiChanged(hwnd, static_cast<UINT>(HIWORD(wp)), reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_DPICHANGED_AFTERPARENT: OnDpiChanged(hwnd, GetDpiForWindow(hwnd), nullptr); return 0;
        case WM_SETTINGCHANGE: OnSettingChange(reinterpret_cast<const wchar_t*>(lp)); return 0;
        case WM_GETMINMAXINFO:
            if (! _embeddedMode)
            {
                if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
                {
                    Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 520, 360);
                }
            }
            return 0;
        case WM_PAINT: OnPaint(); return 0;
        case WM_ERASEBKGND: return _allowEraseBkgnd ? DefWindowProcW(hwnd, msg, wp, lp) : 1;
        case WM_CLOSE: static_cast<void>(Close()); return 0;
        case WM_COMMAND: OnCommand(hwnd, LOWORD(wp)); return 0;
        case WM_SYSKEYDOWN:
            if (! _embeddedMode && (wp == VK_F10 || wp == VK_MENU) && _menuBarHost.FocusFirstItem())
            {
                return 0;
            }
            break;
        case WM_SYSCHAR:
            if (! _embeddedMode && wp >= 0x20u && _menuBarHost.ActivateMnemonic(static_cast<wchar_t>(wp)))
            {
                return 0;
            }
            break;
        case WM_KEYDOWN:
        {
            const bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;
            OnKeyDown(wp, alt);
            return 0;
        }
        case WM_MOUSEMOVE:
        {
            const POINTS pt = {static_cast<short>(LOWORD(lp)), static_cast<short>(HIWORD(lp))};
            OnMouseMove(pt.x, pt.y);
            return 0;
        }
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_LBUTTONDOWN:
        {
            const POINTS pt = {static_cast<short>(LOWORD(lp)), static_cast<short>(HIWORD(lp))};
            OnLButtonDown(pt.x, pt.y);
            return 0;
        }
        case WM_LBUTTONDBLCLK:
        {
            const POINTS pt = {static_cast<short>(LOWORD(lp)), static_cast<short>(HIWORD(lp))};
            OnLButtonDblClk(pt.x, pt.y);
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            if ((GET_KEYSTATE_WPARAM(wp) & MK_CONTROL) == 0u)
            {
                break;
            }
            POINT point{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
            if (ScreenToClient(hwnd, &point) == FALSE)
            {
                return 0;
            }
            const auto hit = HitTestTreemap(DipFromPx(point.x), DipFromPx(point.y));
            if (GET_WHEEL_DELTA_WPARAM(wp) > 0 && hit.has_value())
            {
                const auto item = ResolveTreemapItem(hit.value());
                if (item.has_value() && item.value().isDirectory && ! item.value().isSynthetic)
                {
                    NavigateTo(item.value().id);
                }
            }
            else if (GET_WHEEL_DELTA_WPARAM(wp) < 0)
            {
                NavigateUp();
            }
            return 0;
        }
        case WM_CONTEXTMENU:
        {
            POINT screenPt{static_cast<LONG>(static_cast<short>(LOWORD(lp))), static_cast<LONG>(static_cast<short>(HIWORD(lp)))};
            const bool keyboardContext = screenPt.x == -1 && screenPt.y == -1;
            if (keyboardContext && _selectedGeneration == _displayGeneration && ResolveTreemapItem(_selectedItemId).has_value())
            {
                _contextItemIdOverride = _selectedItemId;
            }
            const auto restoreContextItem = wil::scope_exit([this]() noexcept { _contextItemIdOverride = 0u; });
            if (screenPt.x == -1 && screenPt.y == -1)
            {
                RECT client{};
                if (GetClientRect(hwnd, &client) != 0)
                {
                    screenPt.x = client.left + ((client.right - client.left) / 2);
                    screenPt.y = client.top + ((client.bottom - client.top) / 2);
                    ClientToScreen(hwnd, &screenPt);
                }
                else
                {
                    screenPt = {};
                }
            }
            OnContextMenu(hwnd, screenPt);
            return 0;
        }
        case WM_TIMER: OnTimer(static_cast<UINT_PTR>(wp)); return 0;
        case WM_NCACTIVATE: OnNcActivate(hwnd, wp != FALSE); return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_NCDESTROY: return OnNcDestroy(hwnd, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void ViewerSpace::OnCreate(HWND hwnd)
{
    if (! _workspaceLifetime)
        _workspaceLifetime = std::make_shared<const uint8_t>(uint8_t{0});
    _deferredDisposalSubmitMaxUs.store(0u, std::memory_order_relaxed);
    _deferredDisposalSubmitCount     = 0u;
    _deferredDisposalSubmitTotalUs   = 0u;
    _deferredDisposalSubmitBytes     = 0u;
    _deferredDisposalSubmitOwnerPeak = 0u;

    _allowEraseBkgnd = true;
    _dpi             = static_cast<float>(GetDpiForWindow(hwnd));
#ifdef ENABLE_TESTS
    _debugPublicationPhaseTiming = ReadPositiveUint32Environment(kTestPublicationPhaseTimingEnvVar).has_value();
#endif

    if (! _embeddedMode && GetMenu(hwnd))
    {
        _menuBarHost.SetTheme(_hasTheme ? RedSalamander::MakeThemePaletteFromViewerTheme(_theme) : DxUi::MakeDefaultThemePalette(false));
        _menuBarHost.SetRefreshMenuStateCallback([this]
        {
            if (_hWnd)
            {
                UpdateMenuState(_hWnd.get(), false);
            }
        });
        _menuBarHost.SetOnTabBoundary([this](bool) noexcept
        {
            if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
            {
                SetFocus(_hWnd.get());
            }
            return true;
        });
        _menuBarHost.SetOnEscape([this]() noexcept
        {
            if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
            {
                SetFocus(_hWnd.get());
            }
            return true;
        });
    }

    _animationTimerArmed = SetTimer(hwnd, kTimerAnimationId, kAnimationIntervalMs, nullptr) != 0;
    ApplyThemeToWindow(hwnd);
    ApplyPendingViewerSpaceClassBackgroundBrush(hwnd);
    CreateAnalysisHost(hwnd);
    CreateNavigationHost(hwnd);
    LayoutNavigationHost();
    UpdateNavigationHostState();
    UpdateMenuState(hwnd);
    if (! _embeddedMode)
    {
        _menuBarHost.UpdateLayout();
    }
}

void ViewerSpace::OnNcActivate(HWND hwnd, bool windowActive) noexcept
{
    ApplyTitleBarTheme(hwnd, windowActive);
}

void ViewerSpace::OnDestroy()
{
    const auto destructionStartedAt = std::chrono::steady_clock::now();
    std::array<uint64_t, 9u> destructionStages{};
    const auto markDestructionStage = [&](size_t index) noexcept { destructionStages[index] = Debug::Perf::ElapsedUs(destructionStartedAt); };
    _workspaceLifetime.reset();
    _excludedPaths.clear();
    _excludedBelow.clear();
    _excludedRoots.clear();
    _workspaceNextExcludedBelow.clear();
    _workspaceNextExcludedRoots.clear();
    _workspaceBindingParts.clear();
    _workspaceGeneration        = 0u;
    _workspaceBindingGeneration = 0u;
    _workspaceRebinding         = false;
    _workspaceViewExcluded      = false;
    _workspaceUnresolved        = 0u;
    _workspaceResetBindings     = false;
    _cleanupMarks.clear();
    _cleanupGeneration = 0u;
    _reviewingCleanup  = false;
    _showingProblems   = false;
    _problemRows.clear();
    _problemCursor = 0u;
    _workspaceFacts.clear();
    _workspaceFactsNext.clear();
    _workspaceFactsReady       = false;
    _workspaceFactsBuilding    = false;
    _workspaceFactsGeneration  = 0u;
    _typeFilter                = 0u;
    _ageFilter                 = 0u;
    _exportProjectionRequested = false;
    SubmitWorkspaceExport(false);
    if (_analysisResizeStartRatio.has_value())
    {
        (_analysisSidePanel ? _config.analysisPanelRatio : _config.analysisBottomRatio) = _analysisResizeStartRatio.value();
        _analysisResizeStartRatio.reset();
    }
    _pendingInputPresentation    = {};
    _trackingMouse               = false;
    _hasLastMouseMoveClientPoint = false;
    _lastMouseMoveScreenPoint    = {};
    ++_deferredStartScanGeneration;
    _deferredStartScan.reset();
    _workingModelRetirementPending = false;
    markDestructionStage(0u);
    _navigationHost.Detach();
    markDestructionStage(1u);
    _analysisHost.Detach();
    markDestructionStage(2u);
    CancelScanCacheRestore();
    AbandonScanWorkers();
    markDestructionStage(3u);
    EmitQueueDiagnostics(false, true);
    markDestructionStage(4u);
    DiscardDirectorySummaryStages();
    CancelScanCacheBuild();
    // Hold only the free-threaded D3D device. Every HWND, immediate context,
    // swap chain and single-threaded D2D object is still released here.
    auto closingRendererDevice = CaptureClosingRendererDevice();
#ifdef ENABLE_TESTS
    _rendererFailureStage = kRendererDiscardDestroy;
    _rendererFailureHr    = 0u;
#endif
    DiscardDirect2D();
    _otherStrokeStyle.reset();
    _dogEarFlapGeometry.reset();
    _d2dFactory.reset();
    _dwriteFactory.reset();
    _fileSystem.reset();
    _fileSystemName.clear();
    _fileSystemShortId.clear();
    _fileSystemIsWin32 = true;
    markDestructionStage(5u);
    RetireDisplayedStorage(std::move(_displayModel), std::move(_workingModel), true, true, std::move(closingRendererDevice));
    _model = nullptr;
    markDestructionStage(6u);
    AbandonDeferredDisposals();
    _refreshDisplayState.reset();
    _refreshPending           = false;
    _refreshAdoptionPending   = false;
    _refreshPathLookupStarted = false;
    markDestructionStage(7u);
    NotifyViewerClosed();
    _hostPaneExecute.reset();
    markDestructionStage(8u);
    Debug::Perf::Emit(L"viewer.space.close.destroy_stages_us",
                      std::format(L"clear={};navigation={};analysis={};workers={};queueDiagnostics={};renderer={};retire={};abandonDisposals={};notify={}",
                                  destructionStages[0],
                                  destructionStages[1] - destructionStages[0],
                                  destructionStages[2] - destructionStages[1],
                                  destructionStages[3] - destructionStages[2],
                                  destructionStages[4] - destructionStages[3],
                                  destructionStages[5] - destructionStages[4],
                                  destructionStages[6] - destructionStages[5],
                                  destructionStages[7] - destructionStages[6],
                                  destructionStages[8] - destructionStages[7]),
                      destructionStages[8],
                      0u,
                      0u,
                      S_OK);
}

void ViewerSpace::OnSize(UINT width, UINT height) noexcept
{
    _analysisHost.CancelResize();
    _clientSize.cx = static_cast<LONG>(width);
    _clientSize.cy = static_cast<LONG>(height);

    // The map and hover geometry are about to change. Require a newly delivered
    // client point instead of carrying the old hover through the resize.
    OnMouseLeave();

    if (! _embeddedMode)
    {
        _menuBarHost.UpdateLayout();
    }
    LayoutNavigationHost();
    LayoutAnalysisHost();
    CancelLayoutBuild();
    _layoutDirty = true;
    _layoutCandidateCache.clear();
    InvalidateStaticTreemapCache();
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void ViewerSpace::OnDpiChanged(HWND hwnd, UINT dpi, const RECT* suggestedRect) noexcept
{
    _dpi = static_cast<float>(dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi);

    if (suggestedRect != nullptr && ! _embeddedMode)
    {
        SetWindowPos(hwnd,
                     nullptr,
                     suggestedRect->left,
                     suggestedRect->top,
                     suggestedRect->right - suggestedRect->left,
                     suggestedRect->bottom - suggestedRect->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

#ifdef ENABLE_TESTS
    _rendererFailureStage = kRendererDiscardDpiChanged;
    _rendererFailureHr    = 0u;
#endif
    DiscardDirect2D();
    ApplyThemeToWindow(hwnd);
    ApplyPendingViewerSpaceClassBackgroundBrush(hwnd);
    if (! _embeddedMode)
    {
        _menuBarHost.UpdateLayout();
    }
    _layoutDirty = true;
    _layoutCandidateCache.clear();
    InvalidateRect(hwnd, nullptr, FALSE);
    LayoutNavigationHost();
    LayoutAnalysisHost();
    CancelLayoutBuild();
}

void ViewerSpace::CreateNavigationHost(HWND hwnd)
{
    using Host = ViewerSpaceNavigation::NavigationHost;
    auto text  = [](UINT id) { return LoadStringResource(g_hInstance, id); };
    Host::Captions captions;
    captions.back    = text(IDS_VIEWERSPACE_NAV_BACK);
    captions.forward = text(IDS_VIEWERSPACE_NAV_FORWARD);

    captions.path      = text(IDS_VIEWERSPACE_ANALYSIS_FIELD_PATH);
    captions.ancestors = text(IDS_VIEWERSPACE_NAV_ANCESTORS);
    captions.folders   = text(IDS_VIEWERSPACE_NAV_FOLDERS);
    captions.mode      = text(IDS_VIEWERSPACE_ANALYSIS_MODE);
    captions.explorer  = text(IDS_VIEWERSPACE_MODE_EXPLORER);
    captions.analyst   = text(IDS_VIEWERSPACE_MODE_ANALYST);
    captions.hidePanel = text(IDS_VIEWERSPACE_HIDE_PANEL);
    captions.showPanel = text(IDS_VIEWERSPACE_SHOW_PANEL);
    captions.mapOnly   = text(IDS_VIEWERSPACE_LAYOUT_MAP_ONLY);

    Host::Callbacks callbacks;
    callbacks.navigate = [this](uint64_t token)
    {
        const uint32_t generation = static_cast<uint32_t>(token >> 32u);
        const uint32_t nodeId     = static_cast<uint32_t>(token & 0xFFFFFFFFu);
        if (nodeId != 0u && generation == _displayGeneration)
        {
            NavigateTo(nodeId);
        }
    };
    callbacks.back    = [this] { NavigateHistory(false); };
    callbacks.forward = [this] { NavigateHistory(true); };

    callbacks.toggleAnalysis = [this] { ToggleAnalysisPanel(); };
    callbacks.modeChanged    = [this](bool analyst) { SetViewMode(analyst ? ViewMode::Analyst : ViewMode::Explorer); };
    callbacks.folders        = [this](uint64_t token) { return GetNavigationFolders(token); };
    callbacks.focusBoundary  = [this](bool reverse)
    {
        if (reverse)
        {
            const HWND menu = _menuBarHost.GetHwnd();
            if (! _embeddedMode && menu != nullptr && IsWindowVisible(menu) != FALSE)
            {
                SetFocus(menu);
            }
            else if (_hWnd)
            {
                SetFocus(_hWnd.get());
            }
            return true;
        }

        const HWND analysis = _analysisHost.GetHwnd();
        if (analysis != nullptr && IsWindowVisible(analysis) != FALSE && _analysisHost.Focus(false))
        {
            return true;
        }
        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
        return true;
    };

    const HRESULT result = _navigationHost.Attach(hwnd, captions, std::move(callbacks));
    if (FAILED(result))
    {
        Debug::Error(L"ViewerSpace navigation host failed: 0x{:08X}", static_cast<unsigned long>(result));
        return;
    }
    _navigationHost.ApplyTheme(_hasTheme ? RedSalamander::MakeThemePaletteFromViewerTheme(_theme) : DxUi::MakeDefaultThemePalette(false));
    LayoutNavigationHost();
    UpdateNavigationHostState();
}

std::vector<ViewerSpaceNavigation::NavigationHost::Crumb> ViewerSpace::GetNavigationFolders(uint64_t token) const
{
    using Host = ViewerSpaceNavigation::NavigationHost;
    std::vector<Host::Crumb> folders;
    if (static_cast<uint32_t>(token >> 32u) != _displayGeneration)
        return folders;
    const Node* parent = TryGetRealNode(static_cast<uint32_t>(token));
    if (parent == nullptr)
        return folders;
    for (uint32_t id : GetRealNodeChildren(*parent))
    {
        const Node* child = TryGetRealNode(id);
        if (child != nullptr && child->isDirectory && ! child->isSynthetic)
        {
            folders.push_back({(static_cast<uint64_t>(_displayGeneration) << 32u) | id, std::wstring(child->name)});
            if (folders.size() == 512u)
                break;
        }
    }
    std::sort(folders.begin(),
              folders.end(),
              [](const Host::Crumb& a, const Host::Crumb& b)
    {
        // Presentation collation must not change the literal names or generation-qualified targets.
        return NaturalSort::Less(a.text, b.text);
    });
    return folders;
}

#ifdef ENABLE_TESTS
LRESULT ViewerSpace::DebugNavigationFolder(WndMsg::ViewerSpaceNavigationFolderDebugQuery* query)
{
    if (query == nullptr)
        return FALSE;
    const auto folders = GetNavigationFolders(query->parentToken);
    query->count       = folders.size();
    query->token       = 0u;
    if (query->index >= folders.size())
        return FALSE;
    const auto& folder = folders[query->index];
    query->token       = folder.id;
    const size_t count = std::min(folder.text.size(), std::size(query->name) - 1u);
    std::copy_n(folder.text.begin(), count, query->name);
    query->name[count] = L'\0';
    if (query->navigate)
        NavigateTo(static_cast<uint32_t>(folder.id));
    return TRUE;
}
#endif

void ViewerSpace::LayoutNavigationHost() noexcept
{
    if (_navigationHost.GetHwnd() == nullptr)
    {
        return;
    }
    const int top    = PxFromDip(GetHeaderTopDip() + kHeaderHeightDip);
    const int bottom = PxFromDip(GetHeaderTopDip() + kHeaderHeightDip + kNavigationHeightDip);
    _navigationHost.SetBounds({0, top, _clientSize.cx, std::max(top, bottom)});
}

void ViewerSpace::UpdateNavigationHostState() noexcept
{
    _workspaceViewExcluded = _model && IsItemExcluded(_model->viewNodeId);
    UpdateHeaderTextCache();
    if (_navigationHost.GetHwnd() == nullptr || _model == nullptr)
    {
        return;
    }

    std::vector<ViewerSpaceNavigation::NavigationHost::Crumb> crumbs;
    uint32_t nodeId = _model->viewNodeId;
    for (size_t depth = 0u; nodeId != 0u && depth < 512u; ++depth)
    {
        const Node* const node = TryGetRealNode(nodeId);
        if (node == nullptr)
        {
            break;
        }
        const uint64_t token = (static_cast<uint64_t>(_displayGeneration) << 32u) | node->id;
        crumbs.push_back({token, std::wstring(node->name)});
        nodeId = node->parentId;
    }
    std::reverse(crumbs.begin(), crumbs.end());
    _navigationHost.SetState(crumbs, ! _navStack.empty(), ! _navForwardStack.empty());
}

void ViewerSpace::OnPaint()
{
    EnsureAnimationTimer();
    const auto paintStartedAt = std::chrono::steady_clock::now();

    if (! _hWnd)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd.get(), &ps);
    _allowEraseBkgnd          = false;
    static_cast<void>(hdc);

    const uint32_t firstRestoreGeneration = _scanCacheRestoreSnapshot && ! _scanCacheRestorePaintSliceStarted ? _scanCacheRestoreGeneration : 0u;
    const auto queueRestoreContinuation   = wil::scope_exit([this, firstRestoreGeneration]() noexcept
    {
        if (firstRestoreGeneration != 0u)
        {
            QueueScanCacheRestoreContinuation(firstRestoreGeneration);
        }
    });

    // Use one existing bounded restore turn before the first renderer preparation. Timer turns
    // retain their cadence; subsequent paints cannot add more restore work to the input path.
    if (_scanCacheRestoreSnapshot && ! _scanCacheRestorePaintSliceStarted)
    {
        _scanCacheRestorePaintSliceStarted = true;
        ContinueScanCacheRestore();
    }

    const bool prepareRequired  = ! _renderTarget;
    const auto prepareStartedAt = std::chrono::steady_clock::now();
    const bool prepared         = EnsureDirect2D(_hWnd.get());
    Debug::Perf::EmitDurationUs(
        L"viewer.space.render.prepare_us", Debug::Perf::ElapsedUs(prepareStartedAt), prepareRequired ? 1u : 0u, 0u, prepared ? S_OK : E_FAIL);
    if (! prepared)
    {
        return;
    }

    EnsureLayoutForView();

    HRESULT drawHr = S_OK;
    _renderTarget->BeginDraw();
    auto endDraw = wil::scope_exit([&] { drawHr = _renderTarget->EndDraw(); });

    const D2D1::ColorF bg          = _hasTheme ? ColorFromArgb(_theme.backgroundArgb) : D2D1::ColorF(D2D1::ColorF::White);
    const D2D1::ColorF textColor   = _hasTheme ? ColorFromArgb(_theme.textArgb) : D2D1::ColorF(D2D1::ColorF::Black);
    const D2D1::ColorF accentColor = _hasTheme ? ColorFromArgb(_theme.accentArgb) : D2D1::ColorF(D2D1::ColorF::CornflowerBlue);
    const D2D1::ColorF selectionBg = _hasTheme ? ColorFromArgb(_theme.selectionBackgroundArgb) : accentColor;
    const D2D1::ColorF errorTint   = _hasTheme ? ColorFromArgb(_theme.alertErrorBackgroundArgb) : D2D1::ColorF(D2D1::ColorF::IndianRed);

    const bool highContrast = _hasTheme && _theme.highContrast != FALSE;
    const bool rainbowMode  = _hasTheme && _theme.rainbowMode != FALSE;
    const bool darkMode     = _hasTheme && _theme.darkMode != FALSE;

    const double nowSeconds = NowSeconds();

    _renderTarget->Clear(bg);

#ifdef ENABLE_TESTS
    auto applyForcedRendererFault = [&](HRESULT& forcedDrawHr, HRESULT& forcedPresentHr) noexcept
    {
        const uint32_t mode = _debugForcedRendererFault;
        if (mode == 0u)
        {
            return;
        }

        _debugForcedRendererFault = 0u;
        if (mode == static_cast<uint32_t>(WndMsg::ViewerSpaceRendererFaultDebugMode::D2DRecreateTarget))
        {
            forcedDrawHr = D2DERR_RECREATE_TARGET;
        }
        else if (mode == static_cast<uint32_t>(WndMsg::ViewerSpaceRendererFaultDebugMode::DxgiDeviceRemoved))
        {
            forcedDrawHr    = S_OK;
            forcedPresentHr = DXGI_ERROR_DEVICE_REMOVED;
        }
    };
#endif

    const float headerTopDip    = GetHeaderTopDip();
    const float headerBottomDip = GetHeaderBottomDip();
    const D2D1_RECT_F headerRc  = D2D1::RectF(0.0f, headerTopDip, DipFromPx(_clientSize.cx), headerBottomDip);

    const bool scanActive              = _overallState == ScanState::Queued || _overallState == ScanState::Scanning;
    const bool completionOverlayActive = ! scanActive && _overallState == ScanState::Done && _scanCompletedSinceSeconds > 0.0 &&
                                         (nowSeconds - _scanCompletedSinceSeconds) >= 0.0 &&
                                         (nowSeconds - _scanCompletedSinceSeconds) < kScanCompletedOverlaySeconds;
    uint64_t culledTileCount           = 0u;
    uint64_t textDrawCandidateCount    = 0u;
    for (const DrawItem& item : _drawItems)
    {
        if (item.lod == DrawItem::Lod::Culled)
        {
            culledTileCount += 1u;
        }
        if (item.lod >= DrawItem::Lod::Medium)
        {
            textDrawCandidateCount += 1u;
        }
    }
    const uint64_t tileDrawCount = static_cast<uint64_t>(_drawItems.size()) - culledTileCount;
#ifdef ENABLE_TESTS
    _lastTileDrawCount = tileDrawCount;
    _lastTextDrawCount = textDrawCandidateCount;
#endif

    if (CanReplayStaticTreemapCache(scanActive, completionOverlayActive))
    {
        _staticTreemapCache.hits += 1u;
        Debug::Perf::EmitCounter(L"viewer.space.render.static_cache_hit");
        _renderTarget->DrawBitmap(_staticTreemapCache.bitmap.get(), nullptr, 1.0f, D2D1_BITMAP_INTERPOLATION_MODE_NEAREST_NEIGHBOR, nullptr);
        PaintHoverOverlay(accentColor);
        PaintTooltipOverlay();

        endDraw.reset();

        HRESULT presentHr = S_OK;
        if (SUCCEEDED(drawHr) && _renderer.swapChain)
        {
            // DWM composes this window; the 16-ms animation timer already paces redraws.
            // Waiting for vblank here delays input and the next bounded restore turn.
            const auto presentStartedAt = std::chrono::steady_clock::now();
            presentHr                   = _renderer.swapChain->Present(0u, 0u);
            Debug::Perf::EmitDurationUs(L"viewer.space.render.present_us", Debug::Perf::ElapsedUs(presentStartedAt), 0u, _drawItems.size(), presentHr);
        }
#ifdef ENABLE_TESTS
        applyForcedRendererFault(drawHr, presentHr);
#endif
        const HRESULT renderHr = FAILED(drawHr) ? drawHr : presentHr;
        if (_renderer.swapChain)
            CompletePendingInputPresentation(renderHr, tileDrawCount);

        const uint64_t paintUs = Debug::Perf::ElapsedUs(paintStartedAt);
#ifdef ENABLE_TESTS
        _lastPaintUs = paintUs;
#endif
        Debug::Perf::Emit(
            L"viewer.space.render.paint_us", L"static-cache-hit", paintUs, static_cast<uint64_t>(_drawItems.size()), _staticTreemapCache.hits, renderHr);
        Debug::Perf::Emit(
            L"viewer.space.render.static_cache_replay_us", L"bitmap", paintUs, static_cast<uint64_t>(_drawItems.size()), _staticTreemapCache.bytes, renderHr);
        Debug::Perf::Emit(L"viewer.space.render.tile_draw_count", L"static-cache-hit", 0u, 0u, culledTileCount, renderHr);
        Debug::Perf::Emit(L"viewer.space.render.text_draw_count", L"static-cache-hit", 0u, 0u, textDrawCandidateCount, renderHr);
        Debug::Perf::Emit(L"viewer.space.layout.visible_tiles", L"static-cache-hit", 0u, tileDrawCount, culledTileCount, renderHr);

        if (! _openToFirstPaintEmitted && _openStartedAt != std::chrono::steady_clock::time_point{})
        {
            _openToFirstPaintEmitted = true;
            Debug::Perf::Emit(L"viewer.space.open_to_first_paint_us",
                              L"static-cache-hit",
                              Debug::Perf::ElapsedUs(_openStartedAt),
                              static_cast<uint64_t>(_drawItems.size()),
                              0u,
                              renderHr);
        }

        if (drawHr == D2DERR_RECREATE_TARGET || presentHr == DXGI_ERROR_DEVICE_REMOVED || presentHr == DXGI_ERROR_DEVICE_RESET)
        {
#ifdef ENABLE_TESTS
            _rendererFailureStage = kRendererDiscardStaticDeviceLost;
            _rendererFailureHr    = static_cast<uint32_t>(FAILED(drawHr) ? drawHr : presentHr);
#endif
            DiscardDirect2D();
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
            }
        }
        else if (FAILED(drawHr))
        {
            Debug::Warning(L"ViewerSpace: EndDraw failed during static cache replay: 0x{:08X}", static_cast<unsigned long>(drawHr));
        }
        else if (FAILED(presentHr))
        {
            Debug::Warning(L"ViewerSpace: Present failed during static cache replay: 0x{:08X}", static_cast<unsigned long>(presentHr));
        }
        return;
    }

    if (_brushAccent)
    {
        const float barHeight = 5.0f;
        const D2D1_RECT_F track =
            D2D1::RectF(headerRc.left + kPaddingDip, headerRc.bottom - barHeight - 4.0f, headerRc.right - kPaddingDip, headerRc.bottom - 4.0f);
        if (track.right > track.left && track.bottom > track.top)
        {
            const float radius      = barHeight * 0.5f;
            const float trackAlpha  = highContrast ? 0.32f : (scanActive ? (darkMode ? 0.22f : 0.16f) : 0.10f);
            D2D1::ColorF trackColor = accentColor;
            if (highContrast)
            {
                trackColor = textColor;
            }
            else if (rainbowMode)
            {
                const double hue01 = Fract((nowSeconds - _animationStartSeconds) * 0.08);
                trackColor         = ColorFFromHsv(hue01, 0.90, darkMode ? 0.98 : 0.90, 1.0f);
            }

            _brushAccent->SetColor(D2D1::ColorF(trackColor.r, trackColor.g, trackColor.b, trackAlpha));
            _renderTarget->FillRoundedRectangle(D2D1::RoundedRect(track, radius, radius), _brushAccent.get());

            if (scanActive)
            {
                const float t           = static_cast<float>(Fract((nowSeconds - _animationStartSeconds) * 0.72));
                const float w           = std::max(42.0f, (track.right - track.left) * 0.26f);
                const float x           = track.left - w + t * ((track.right - track.left) + w * 2.0f);
                const D2D1_RECT_F chunk = D2D1::RectF(x, track.top, x + w, track.bottom);

                _renderTarget->PushAxisAlignedClip(track, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                const float chunkAlpha = highContrast ? 0.85f : 0.70f;

                auto chunkColorFor = [&](double offset) noexcept -> D2D1::ColorF
                {
                    if (highContrast)
                    {
                        return textColor;
                    }

                    if (rainbowMode)
                    {
                        const double hue01 = Fract((nowSeconds - _animationStartSeconds) * 0.08 + static_cast<double>(t) * 0.24 + offset);
                        return ColorFFromHsv(hue01, 0.95, darkMode ? 0.99 : 0.92, 1.0f);
                    }

                    return accentColor;
                };

                const float inset1 = std::min(14.0f, w * 0.22f);
                const float inset2 = std::min(24.0f, w * 0.36f);

                auto fillRounded = [&](D2D1_RECT_F rc, float alpha, double hueOffset) noexcept
                {
                    if (rc.right <= rc.left || rc.bottom <= rc.top)
                    {
                        return;
                    }

                    const D2D1::ColorF c = chunkColorFor(hueOffset);
                    _brushAccent->SetColor(D2D1::ColorF(c.r, c.g, c.b, alpha));
                    const float r = std::min(radius, std::min((rc.right - rc.left) * 0.5f, (rc.bottom - rc.top) * 0.5f));
                    _renderTarget->FillRoundedRectangle(D2D1::RoundedRect(rc, r, r), _brushAccent.get());
                };

                fillRounded(chunk, chunkAlpha * 0.35f, 0.0);

                D2D1_RECT_F mid = chunk;
                mid.left += inset1;
                mid.right -= inset1;
                fillRounded(mid, chunkAlpha * 0.60f, 0.05);

                D2D1_RECT_F core = chunk;
                core.left += inset2;
                core.right -= inset2;
                fillRounded(core, chunkAlpha * 0.92f, 0.10);
                _renderTarget->PopAxisAlignedClip();
            }
        }
    }

    if (_brushText)
    {
        if (_headerStatusText.empty())
        {
            UpdateHeaderTextCache();
        }

        const std::wstring& status = _headerStatusText;

        const float buttonSide           = kHeaderButtonWidthDip;
        const bool showCancel            = scanActive;
        const D2D1_RECT_F cancelButtonRc = showCancel ? D2D1::RectF(headerRc.right - buttonSide, headerTopDip, headerRc.right, headerBottomDip)
                                                      : D2D1::RectF(headerRc.right, headerTopDip, headerRc.right, headerBottomDip);

        if (_brushBackground)
        {
            const float hoverAlpha = 0.18f;
            if (_hoverHeaderHit == HeaderHit::Cancel)
            {
                _brushBackground->SetColor(D2D1::ColorF(selectionBg.r, selectionBg.g, selectionBg.b, hoverAlpha));
                _renderTarget->FillRectangle(cancelButtonRc, _brushBackground.get());
            }
        }

        if (_headerIconFormat)
        {
            if (showCancel)
            {
                const wchar_t cancelGlyph[] = L"×";
                _renderTarget->DrawTextW(cancelGlyph, 1, _headerIconFormat.get(), cancelButtonRc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }
        }

        const std::wstring_view pathText = _headerSizeText;

        const float contentTop = headerTopDip + 4.0f;
        // The navigation child owns the following band, including cached paint.
        const float contentBottom = headerTopDip + kHeaderHeightDip - 4.0f;
        D2D1_RECT_F contentRc     = headerRc;
        contentRc.left += kPaddingDip;
        contentRc.right -= (showCancel ? buttonSide : 0.0f) + kPaddingDip;
        if (contentRc.right > contentRc.left && contentBottom > contentTop)
        {
            const float lineHeight  = (contentBottom - contentTop) / 3.0f;
            const D2D1_RECT_F line1 = D2D1::RectF(contentRc.left, contentTop, contentRc.right, contentTop + lineHeight);
            const D2D1_RECT_F line2 = D2D1::RectF(contentRc.left, line1.bottom, contentRc.right, line1.bottom + lineHeight);
            const D2D1_RECT_F line3 = D2D1::RectF(contentRc.left, line2.bottom, contentRc.right, contentBottom);

            const float availableWidth = std::max(0.0f, contentRc.right - contentRc.left);
            const float columnGapDip   = 12.0f;

            float statusWidthDip = 0.0f;
            if (_dwriteFactory && _headerStatusFormatRight && ! status.empty())
            {
                statusWidthDip = MeasureTextWidthDip(_dwriteFactory.get(), _headerStatusFormatRight.get(), status);
            }

            float statusColumnWidthDip = 0.0f;
            if (statusWidthDip > 0.0f)
            {
                statusColumnWidthDip = statusWidthDip + 10.0f;
            }
            else
            {
                statusColumnWidthDip = std::min(160.0f, availableWidth * 0.35f);
            }
            statusColumnWidthDip = std::clamp(statusColumnWidthDip, 48.0f, availableWidth);

            D2D1_RECT_F line1Right = line1;
            line1Right.left        = std::max(line1Right.left, line1Right.right - statusColumnWidthDip);

            _headerProblemsBounds = _scanFailedFolders ? line1Right : D2D1_RECT_F{};
            D2D1_RECT_F line1Left = line1;
            line1Left.right       = std::max(line1Left.left, line1Right.left - columnGapDip);

            const float line2RightWidth = std::min(320.0f, availableWidth * 0.40f);
            D2D1_RECT_F line2Left       = line2;
            line2Left.right             = std::max(line2Left.left, line2Left.right - line2RightWidth);
            D2D1_RECT_F line2Right      = line2;
            line2Right.left             = line2Left.right;

            if (_headerFormat && ! pathText.empty())
            {
                std::wstring_view pathToDraw = pathText;
                const float maxWidthDip      = std::max(0.0f, line1Left.right - line1Left.left);
                if (_dwriteFactory && maxWidthDip > 1.0f)
                {
                    if (_headerPathSourceText != pathText || std::abs(maxWidthDip - _headerPathDisplayMaxWidthDip) > 0.5f)
                    {
                        _headerPathSourceText.assign(pathText.data(), pathText.size());
                        _headerPathDisplayMaxWidthDip = maxWidthDip;
                        _headerPathDisplayText =
                            BuildMiddleEllipsisPathText(_headerPathSourceText, _fileSystemIsWin32, _dwriteFactory.get(), _headerFormat.get(), maxWidthDip);
                    }

                    if (! _headerPathDisplayText.empty())
                    {
                        pathToDraw = _headerPathDisplayText;
                    }
                }

                _renderTarget->DrawTextW(
                    pathToDraw.data(), static_cast<UINT32>(pathToDraw.size()), _headerFormat.get(), line1Left, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (_headerStatusFormatRight && ! status.empty())
            {
                _renderTarget->DrawTextW(status.c_str(),
                                         static_cast<UINT32>(status.size()),
                                         _headerStatusFormatRight.get(),
                                         line1Right,
                                         _brushText.get(),
                                         D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (scanActive && _headerInfoFormatRight && ! _headerProcessingText.empty())
            {
                _renderTarget->DrawTextW(_headerProcessingText.c_str(),
                                         static_cast<UINT32>(_headerProcessingText.size()),
                                         _headerInfoFormatRight.get(),
                                         line2Right,
                                         _brushText.get(),
                                         D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (_headerInfoFormat)
            {
                if (! _headerCountsText.empty() || ! _headerSizeText.empty())
                {
                    if (! _headerCountsText.empty())
                    {
                        _renderTarget->DrawTextW(_headerCountsText.c_str(),
                                                 static_cast<UINT32>(_headerCountsText.size()),
                                                 _headerInfoFormat.get(),
                                                 line2Left,
                                                 _brushText.get(),
                                                 D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }

                    if (! _headerViewText.empty())
                    {
                        _renderTarget->DrawTextW(_headerViewText.c_str(),
                                                 static_cast<UINT32>(_headerViewText.size()),
                                                 _headerInfoFormat.get(),
                                                 line3,
                                                 _brushText.get(),
                                                 D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }
                }
            }
        }
    }

    const D2D1_RECT_F treemapRc                                = GetTreemapBounds();
    const D2D1_RECT_F treemapLayoutRc                          = treemapRc;
    const float treemapLayoutAreaDip2                          = RectArea(treemapLayoutRc);
    static constexpr float kLargeTileAreaFractionWhileScanning = 0.10f;
    static constexpr float kLargeTileAreaFractionIdle          = 0.10f;
    const float largeTileAreaFractionThreshold                 = scanActive ? kLargeTileAreaFractionWhileScanning : kLargeTileAreaFractionIdle;

    uint64_t viewBytes   = 0;
    const Node* viewNode = TryGetRealNode(_model->viewNodeId);
    if (viewNode != nullptr)
    {
        viewBytes = ShownBytes(viewNode->id, viewNode->totalBytes);
    }

    auto resolveItem = [&](uint32_t itemId) noexcept -> std::optional<ResolvedItem> { return ResolveTreemapItem(itemId); };

    auto rectForItem = [&](const DrawItem& item) noexcept -> D2D1_RECT_F
    {
        float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
        gap       = std::clamp(gap, 0.5f, kItemGapDip);

        D2D1_RECT_F rc = item.currentRect;
        rc.left += gap;
        rc.top += gap;
        rc.right -= gap;
        rc.bottom -= gap;
        return rc;
    };

    // Item ids follow scan completion order, which differs between runs; colours come from names so a rescan looks the same.
    auto stableColorKey = [&](const ResolvedItem& item) noexcept -> uint32_t
    {
        const auto hashName = [](std::wstring_view name) noexcept -> uint32_t
        {
            const size_t value = std::hash<std::wstring_view>{}(name);
            return static_cast<uint32_t>(static_cast<uint64_t>(value) ^ (static_cast<uint64_t>(value) >> 32u));
        };

        uint32_t key = item.isSynthetic ? 0x4f746872u : hashName(item.name);
        if (const Node* parent = TryGetRealNode(item.parentId); parent != nullptr)
        {
            key ^= hashName(parent->name) * 0x9e3779b1u;
        }
        return HashU32(key);
    };

    auto baseColorForNode = [&](const ResolvedItem& nodeRef) noexcept -> D2D1::ColorF
    {
        double ratio = 0.0;
        if (viewBytes > 0)
        {
            ratio = static_cast<double>(ShownBytes(nodeRef.id, nodeRef.totalBytes)) / static_cast<double>(viewBytes);
            ratio = std::clamp(ratio, 0.0, 1.0);
        }

        const double sizeFactor = std::sqrt(ratio);
        const uint32_t hashed   = stableColorKey(nodeRef);

        D2D1::ColorF base = bg;
        if (highContrast)
        {
            base = Mix(bg, textColor, nodeRef.isDirectory ? 0.18f : 0.10f);
        }
        else if (rainbowMode)
        {
            const double hue01  = static_cast<double>(hashed) / 4294967296.0;
            const double jitter = static_cast<double>(HashU32(hashed ^ 0x68bc21ebu)) / 4294967296.0;

            double saturation = nodeRef.isDirectory ? 0.95 : 0.88;
            double value      = darkMode ? 0.92 : 0.80;
            if (! nodeRef.isDirectory)
            {
                value = darkMode ? 0.86 : 0.74;
            }

            saturation += (jitter - 0.5) * 0.08;
            value += (jitter - 0.5) * 0.06;

            const D2D1::ColorF rainbow = ColorFFromHsv(hue01, saturation, value, 1.0f);

            float mixT = nodeRef.isDirectory ? 0.82f : 0.74f;
            mixT += darkMode ? 0.08f : 0.04f;
            mixT += static_cast<float>(0.10 * sizeFactor);
            mixT = std::clamp(mixT, 0.55f, 1.0f);
            base = Mix(bg, rainbow, mixT);
        }
        else
        {
            const double variant01 = static_cast<double>(HashU32(hashed ^ 0x9e3779b9u)) / 4294967296.0;
            float accentMix        = nodeRef.isDirectory ? 0.55f : 0.35f;
            accentMix += static_cast<float>((variant01 - 0.5) * 0.30);
            accentMix = std::clamp(accentMix, 0.0f, 1.0f);

            const D2D1::ColorF nodeAccent = Mix(accentColor, selectionBg, accentMix);

            float mixT = nodeRef.isDirectory ? 0.24f : 0.14f;
            mixT += static_cast<float>((nodeRef.isDirectory ? 0.34 : 0.22) * sizeFactor);
            if (darkMode)
            {
                mixT += nodeRef.isDirectory ? 0.06f : 0.03f;
            }
            mixT = std::clamp(mixT, 0.0f, 0.90f);
            base = Mix(bg, nodeAccent, mixT);

            const double shade01      = static_cast<double>(HashU32(hashed ^ 0x85ebca6bu)) / 4294967296.0;
            const float shadeSigned   = static_cast<float>((shade01 - 0.5) * 2.0);
            const float shadeStrength = nodeRef.isDirectory ? 0.16f : 0.12f;

            const D2D1::ColorF lighter = darkMode ? Mix(base, textColor, 0.28f) : Mix(base, bg, 0.22f);
            const D2D1::ColorF darker  = darkMode ? Mix(base, bg, 0.22f) : Mix(base, textColor, 0.22f);

            if (shadeSigned >= 0.0f)
            {
                base = Mix(base, lighter, shadeSigned * shadeStrength);
            }
            else
            {
                base = Mix(base, darker, (-shadeSigned) * shadeStrength);
            }
        }

        if (! highContrast && _config.colorMode)
        {
            D2D1_COLOR_F mapped{};
            if (_config.colorMode == 1u)
                mapped = CategoryColor(ItemCategory(nodeRef));
            else
            {
                constexpr int64_t day = 86400ll * 10000000ll;
                const int64_t date    = ItemAgeDate(nodeRef);
                const int64_t age     = date > 0 && date <= _workspaceNow ? (_workspaceNow - date) / day : -1;
                mapped                = D2D1::ColorF(age < 0 ? 0x858B95u : age < 7 ? 0x4AA67Bu : age < 30 ? 0xBDAC3Cu : age < 365 ? 0xC98439u : 0xBA5961u);
            }
            base = Mix(bg, D2D1::ColorF(mapped.r, mapped.g, mapped.b, mapped.a), darkMode ? 0.84f : 0.68f);
        }
        return base;
    };

    const float labelAreaThresholdDip2 = scanActive ? (32.0f * 32.0f) : kMinHitAreaDip2;

    static constexpr double kMiniSpinnerPhaseSpeed = 1.8;
    static constexpr double kMiniSpinnerHueSpeed   = 0.10;
    static constexpr double kBigSpinnerPhaseSpeed  = 0.95;
    static constexpr double kBigSpinnerHueSpeed    = 0.06;

    auto drawMiniSpinner = [&](const D2D1_POINT_2F& center, float radius, uint32_t seed, double phaseSpeed, double hueSpeed) noexcept
    {
        if (! _brushAccent || radius <= 1.0f)
        {
            return;
        }

        static constexpr int kSegments = 12;
        static constexpr double kTwoPi = 6.28318530717958647692;

        const double seed01  = static_cast<double>(HashU32(seed)) / 4294967296.0;
        const double phase01 = Fract((nowSeconds - _animationStartSeconds) * phaseSpeed + seed01);
        const int head       = static_cast<int>(phase01 * static_cast<double>(kSegments));

        const float baseAlpha = highContrast ? 1.0f : 0.92f;
        const float stroke    = std::clamp(radius * 0.20f, 1.1f, 1.8f);

        for (int i = 0; i < kSegments; ++i)
        {
            const int dist    = (head - i + kSegments) % kSegments;
            const float fade  = 1.0f - static_cast<float>(dist) / static_cast<float>(kSegments);
            const float alpha = baseAlpha * (0.15f + 0.85f * fade);

            D2D1::ColorF c = accentColor;
            if (highContrast)
            {
                c = textColor;
            }
            else if (rainbowMode)
            {
                const double hue01 = Fract((nowSeconds - _animationStartSeconds) * hueSpeed + seed01 + static_cast<double>(i) / static_cast<double>(kSegments));
                c                  = ColorFFromHsv(hue01, 0.95, darkMode ? 0.99 : 0.92, 1.0f);
            }

            _brushAccent->SetColor(D2D1::ColorF(c.r, c.g, c.b, alpha));

            const double ang = (static_cast<double>(i) / static_cast<double>(kSegments)) * kTwoPi;
            const float cx   = static_cast<float>(std::cos(ang));
            const float cy   = static_cast<float>(std::sin(ang));

            const float r0 = radius * 0.52f;
            const float r1 = radius;

            const D2D1_POINT_2F p0 = D2D1::Point2F(center.x + cx * r0, center.y + cy * r0);
            const D2D1_POINT_2F p1 = D2D1::Point2F(center.x + cx * r1, center.y + cy * r1);
            _renderTarget->DrawLine(p0, p1, _brushAccent.get(), stroke);
        }
    };

    // Fill pass: parents first, children later.
    for (const auto& item : _drawItems)
    {
        const std::optional<ResolvedItem> node = resolveItem(item.nodeId);
        if (! node.has_value())
        {
            continue;
        }

        const ResolvedItem& nodeRef = node.value();

        D2D1_RECT_F rc = rectForItem(item);

        if (rc.right <= treemapRc.left || rc.left >= treemapRc.right || rc.bottom <= treemapRc.top || rc.top >= treemapRc.bottom)
        {
            continue;
        }

        const float area = RectArea(rc);
        if (item.lod == DrawItem::Lod::Culled)
        {
            continue;
        }

        D2D1::ColorF base              = baseColorForNode(nodeRef);
        const bool isOtherBucket       = nodeRef.isSynthetic;
        const bool isRealDirectory     = nodeRef.isDirectory && ! nodeRef.isSynthetic;
        const bool incompleteDirectory = isRealDirectory && (nodeRef.scanState == ScanState::NotStarted || nodeRef.scanState == ScanState::Queued ||
                                                             nodeRef.scanState == ScanState::Scanning);
        if (incompleteDirectory && ! highContrast)
        {
            float dimT = darkMode ? 0.38f : 0.28f;
            if (nodeRef.scanState == ScanState::Queued || nodeRef.scanState == ScanState::NotStarted)
            {
                dimT = darkMode ? 0.52f : 0.40f;
            }
            base = Mix(base, bg, dimT);
        }
        if (isRealDirectory && nodeRef.scanState == ScanState::Error && ! highContrast)
        {
            // A folder that could not be read keeps its own tile, tinted with the theme's error colour so it is not mistaken for Done.
            base = Mix(base, errorTint, 0.45f);
        }

        if (_brushBackground)
        {
            float fillAlpha = highContrast ? 1.0f : 0.96f;
            if (! nodeRef.isDirectory)
            {
                fillAlpha -= 0.04f;
            }
            if (isOtherBucket && ! highContrast)
            {
                fillAlpha = std::clamp(fillAlpha - (darkMode ? 0.05f : 0.07f), 0.55f, 0.96f);
            }
            _brushBackground->SetColor(D2D1::ColorF(base.r, base.g, base.b, fillAlpha));
            _renderTarget->FillRectangle(rc, _brushBackground.get());
        }

        if (_brushShading && ! highContrast && ! scanActive)
        {
            _brushShading->SetStartPoint(D2D1::Point2F(rc.left, rc.top));
            _brushShading->SetEndPoint(D2D1::Point2F(rc.right, rc.bottom));
            _renderTarget->FillRectangle(rc, _brushShading.get());
        }

        if (isOtherBucket && _brushOutline && ! highContrast)
        {
            const float w    = std::max(0.0f, rc.right - rc.left);
            const float h    = std::max(0.0f, rc.bottom - rc.top);
            const float side = std::min(w, h);
            if (side >= 18.0f)
            {
                float spacing         = std::clamp(side * 0.12f, 8.0f, 18.0f);
                const float thickness = std::clamp(side * 0.010f, 0.9f, 1.4f);
                float alpha           = darkMode ? 0.11f : 0.09f;
                if (rainbowMode)
                {
                    alpha += 0.04f;
                }

                const D2D1::ColorF hatch = rainbowMode ? textColor : Mix(accentColor, textColor, 0.55f);
                _brushOutline->SetColor(D2D1::ColorF(hatch.r, hatch.g, hatch.b, alpha));

                _renderTarget->PushAxisAlignedClip(rc, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                const float diag = h;
                for (float x0 = rc.left - diag; x0 < rc.right; x0 += spacing)
                {
                    const D2D1_POINT_2F p0 = D2D1::Point2F(x0, rc.bottom);
                    const D2D1_POINT_2F p1 = D2D1::Point2F(x0 + diag, rc.top);
                    _renderTarget->DrawLine(p0, p1, _brushOutline.get(), thickness);
                }
                _renderTarget->PopAxisAlignedClip();
            }
        }

        const bool expanded = isRealDirectory && item.labelHeightDip > 0.0f;

        float directoryHeaderHeightDip = 0.0f;
        if (isRealDirectory)
        {
            const float areaFraction   = (treemapLayoutAreaDip2 > 1.0f) ? (area / treemapLayoutAreaDip2) : 0.0f;
            const bool wantsTallHeader = incompleteDirectory && areaFraction >= largeTileAreaFractionThreshold;

            const float h                   = std::max(0.0f, rc.bottom - rc.top);
            const float baseHeaderHeightDip = std::clamp(24.0f - static_cast<float>(item.depth) * 2.0f, 20.0f, 24.0f);
            float desiredHeaderHeightDip    = expanded ? std::max(item.labelHeightDip, baseHeaderHeightDip) : baseHeaderHeightDip;
            if (wantsTallHeader)
            {
                const float twoLine    = std::clamp(38.0f - static_cast<float>(item.depth) * 2.0f, 30.0f, 38.0f);
                desiredHeaderHeightDip = std::max(desiredHeaderHeightDip, twoLine);
            }
            const float maxHeaderHeightDip = wantsTallHeader ? 44.0f : 24.0f;
            desiredHeaderHeightDip         = std::clamp(desiredHeaderHeightDip, 20.0f, maxHeaderHeightDip);

            if (h >= desiredHeaderHeightDip)
            {
                directoryHeaderHeightDip = desiredHeaderHeightDip;
            }
        }

        if (directoryHeaderHeightDip > 0.0f && _brushBackground)
        {
            D2D1_RECT_F headerBar = rc;
            headerBar.bottom      = std::min(rc.bottom, rc.top + directoryHeaderHeightDip);

            D2D1::ColorF headerColor = Mix(base, bg, darkMode ? 0.10f : 0.22f);
            float headerAlpha        = 0.96f;
            if (rainbowMode)
            {
                headerColor = Mix(base, bg, darkMode ? 0.18f : 0.12f);
                headerAlpha = 0.96f;
            }
            _brushBackground->SetColor(D2D1::ColorF(headerColor.r, headerColor.g, headerColor.b, headerAlpha));
            _renderTarget->FillRectangle(headerBar, _brushBackground.get());

            if (_brushOutline && ! highContrast)
            {
                D2D1_RECT_F sep         = headerBar;
                sep.top                 = std::max(headerBar.top, headerBar.bottom - 1.0f);
                sep.bottom              = headerBar.bottom;
                const D2D1::ColorF line = Mix(base, textColor, 0.20f);
                _brushOutline->SetColor(D2D1::ColorF(line.r, line.g, line.b, 0.55f));
                _renderTarget->FillRectangle(sep, _brushOutline.get());
            }
        }
    }

    // Border + labels pass: draw children first, then parents (so parent borders stay visible).
    for (size_t idx = _drawItems.size(); idx-- > 0;)
    {
        const DrawItem& item = _drawItems[idx];

        const std::optional<ResolvedItem> node = resolveItem(item.nodeId);
        if (! node.has_value())
        {
            continue;
        }

        const ResolvedItem& nodeRef = node.value();

        D2D1_RECT_F rc = rectForItem(item);

        if (rc.right <= treemapRc.left || rc.left >= treemapRc.right || rc.bottom <= treemapRc.top || rc.top >= treemapRc.bottom)
        {
            continue;
        }

        const float area = RectArea(rc);
        if (item.lod == DrawItem::Lod::Culled || item.lod == DrawItem::Lod::Tiny)
        {
            continue;
        }

        const float tileW                = std::max(0.0f, rc.right - rc.left);
        const float tileH                = std::max(0.0f, rc.bottom - rc.top);
        const bool canShowAtLeastOneLine = tileW >= 28.0f && tileH >= 20.0f;
        const bool canShowTileLabels     = item.lod >= DrawItem::Lod::Medium && (area >= labelAreaThresholdDip2 || canShowAtLeastOneLine);

        D2D1::ColorF base         = baseColorForNode(nodeRef);
        const bool isScanningTile = nodeRef.isDirectory && ! nodeRef.isSynthetic && nodeRef.scanState == ScanState::Scanning;
        if (isScanningTile && ! highContrast)
        {
            const float dimT = darkMode ? 0.38f : 0.28f;
            base             = Mix(base, bg, dimT);
        }
        const bool isRealDirectory = nodeRef.isDirectory && ! nodeRef.isSynthetic;
        if (isRealDirectory && nodeRef.scanState == ScanState::Error && ! highContrast)
        {
            base = Mix(base, errorTint, 0.45f);
        }
        const bool expanded      = isRealDirectory && item.labelHeightDip > 0.0f;
        const bool isOtherBucket = nodeRef.isSynthetic;
        const bool incomplete    = isRealDirectory && (nodeRef.scanState == ScanState::NotStarted || nodeRef.scanState == ScanState::Queued ||
                                                       nodeRef.scanState == ScanState::Scanning);

        const float areaFraction   = (treemapLayoutAreaDip2 > 1.0f) ? (area / treemapLayoutAreaDip2) : 0.0f;
        const bool wantsTallHeader = incomplete && areaFraction >= largeTileAreaFractionThreshold;

        float directoryHeaderHeightDip = 0.0f;
        if (isRealDirectory)
        {
            const float h                   = std::max(0.0f, rc.bottom - rc.top);
            const float baseHeaderHeightDip = std::clamp(24.0f - static_cast<float>(item.depth) * 2.0f, 20.0f, 24.0f);
            float desiredHeaderHeightDip    = expanded ? std::max(item.labelHeightDip, baseHeaderHeightDip) : baseHeaderHeightDip;
            if (wantsTallHeader)
            {
                const float twoLine    = std::clamp(38.0f - static_cast<float>(item.depth) * 2.0f, 30.0f, 38.0f);
                desiredHeaderHeightDip = std::max(desiredHeaderHeightDip, twoLine);
            }
            const float maxHeaderHeightDip = wantsTallHeader ? 44.0f : 24.0f;
            desiredHeaderHeightDip         = std::clamp(desiredHeaderHeightDip, 20.0f, maxHeaderHeightDip);

            if (h >= desiredHeaderHeightDip)
            {
                directoryHeaderHeightDip = desiredHeaderHeightDip;
            }
        }
        const bool hasDirectoryHeader = directoryHeaderHeightDip > 0.0f;

        const std::wstring_view watermarkText = scanActive ? std::wstring_view(_scanInProgressWatermarkText) : std::wstring_view(_scanIncompleteWatermarkText);

        const bool showStatusInHeader = wantsTallHeader && hasDirectoryHeader && directoryHeaderHeightDip >= 32.0f && ! watermarkText.empty();

        if (incomplete && _watermarkFormat && _brushWatermark && ! watermarkText.empty() && ! showStatusInHeader)
        {
            const float w    = std::max(0.0f, rc.right - rc.left);
            const float h    = std::max(0.0f, rc.bottom - rc.top);
            const float side = std::min(w, h);

            if (side >= 72.0f && area >= 3200.0f)
            {
                const D2D1_POINT_2F center = D2D1::Point2F((rc.left + rc.right) * 0.5f, (rc.top + rc.bottom) * 0.5f);
                const float diag           = std::sqrt(w * w + h * h);
                const float textW          = std::max(0.0f, diag - 12.0f);
                const float textH          = std::clamp(side * 0.24f, 24.0f, 56.0f);
                const D2D1_RECT_F textRc   = D2D1::RectF(center.x - textW * 0.5f, center.y - textH * 0.5f, center.x + textW * 0.5f, center.y + textH * 0.5f);

                float alpha = highContrast ? 0.78f : (darkMode ? 0.26f : 0.20f);
                if (rainbowMode)
                {
                    alpha += 0.04f;
                }
                alpha = std::clamp(alpha, 0.14f, 0.88f);

                D2D1::ColorF watermarkColor = textColor;
                if (! highContrast && ! rainbowMode)
                {
                    watermarkColor = Mix(base, textColor, darkMode ? 0.86f : 0.76f);
                }
                _brushWatermark->SetColor(D2D1::ColorF(watermarkColor.r, watermarkColor.g, watermarkColor.b, alpha));

                _renderTarget->PushAxisAlignedClip(rc, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                auto popClip = wil::scope_exit([&] { _renderTarget->PopAxisAlignedClip(); });

                D2D1_MATRIX_3X2_F oldTransform{};
                _renderTarget->GetTransform(&oldTransform);
                _renderTarget->SetTransform(D2D1::Matrix3x2F::Rotation(-35.0f, center) * oldTransform);
                auto restoreTransform = wil::scope_exit([&] { _renderTarget->SetTransform(oldTransform); });

                _renderTarget->DrawTextW(watermarkText.data(),
                                         static_cast<UINT32>(watermarkText.size()),
                                         _watermarkFormat.get(),
                                         textRc,
                                         _brushWatermark.get(),
                                         D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }
        }

        if (_brushOutline)
        {
            float stroke = nodeRef.isDirectory ? 1.25f : 1.0f;
            stroke -= static_cast<float>(item.depth) * 0.06f;
            stroke = std::clamp(stroke, 0.85f, 1.35f);

            ID2D1StrokeStyle* strokeStyle = nullptr;
            D2D1::ColorF outline          = highContrast ? textColor : Mix(base, textColor, nodeRef.isDirectory ? 0.20f : 0.14f);
            if (isOtherBucket)
            {
                stroke      = std::max(stroke, 1.85f);
                strokeStyle = (! highContrast && _otherStrokeStyle) ? _otherStrokeStyle.get() : nullptr;

                if (highContrast)
                {
                    outline = textColor;
                }
                else if (rainbowMode)
                {
                    outline = textColor;
                }
                else
                {
                    outline = Mix(accentColor, textColor, 0.30f);
                }
            }
            _brushOutline->SetColor(D2D1::ColorF(outline.r, outline.g, outline.b, 1.0f));

            const bool isFileTile = ! nodeRef.isDirectory && ! nodeRef.isSynthetic;
            if (isFileTile && canShowTileLabels && item.lod >= DrawItem::Lod::Large)
            {
                const float side    = std::min(tileW, tileH);
                const float dogSize = std::clamp(side * 0.18f, 8.0f, 14.0f);
                float cut           = dogSize + 2.0f;
                cut                 = std::clamp(cut, 6.0f, std::max(6.0f, side - 1.0f));

                if (cut > 1.0f && tileW > cut + 1.0f && tileH > cut + 1.0f)
                {
                    const D2D1_POINT_2F tl = D2D1::Point2F(rc.left, rc.top);
                    const D2D1_POINT_2F tr = D2D1::Point2F(rc.right, rc.top);
                    const D2D1_POINT_2F br = D2D1::Point2F(rc.right, rc.bottom);
                    const D2D1_POINT_2F bl = D2D1::Point2F(rc.left, rc.bottom);

                    const D2D1_POINT_2F cutA = D2D1::Point2F(rc.right - cut, rc.top);
                    const D2D1_POINT_2F cutB = D2D1::Point2F(rc.right, rc.top + cut);

                    _renderTarget->DrawLine(tl, cutA, _brushOutline.get(), stroke, strokeStyle);
                    _renderTarget->DrawLine(cutA, cutB, _brushOutline.get(), stroke, strokeStyle);
                    _renderTarget->DrawLine(cutB, br, _brushOutline.get(), stroke, strokeStyle);
                    _renderTarget->DrawLine(br, bl, _brushOutline.get(), stroke, strokeStyle);
                    _renderTarget->DrawLine(bl, tl, _brushOutline.get(), stroke, strokeStyle);
                }
                else
                {
                    _renderTarget->DrawRectangle(rc, _brushOutline.get(), stroke, strokeStyle);
                }
            }
            else
            {
                _renderTarget->DrawRectangle(rc, _brushOutline.get(), stroke, strokeStyle);
            }
        }

        if (_textFormat && _brushText && canShowTileLabels)
        {
            std::wstring fallbackName;
            std::wstring_view nameView = nodeRef.name;
            if (nameView.empty() && ! nodeRef.isSynthetic)
            {
                fallbackName = BuildItemPathText(nodeRef.id);
                nameView     = fallbackName;
            }

            D2D1_RECT_F labelRc = rc;
            labelRc.left += 6.0f;
            labelRc.top += 4.0f;
            labelRc.right -= 6.0f;
            labelRc.bottom -= 4.0f;

            D2D1_RECT_F spinnerRc{};
            bool showSpinner = false;

            if (hasDirectoryHeader)
            {
                const float headerBottom = std::min(rc.bottom, rc.top + directoryHeaderHeightDip);
                labelRc.bottom           = std::min(labelRc.bottom, headerBottom - 2.0f);
                labelRc.bottom           = std::max(labelRc.bottom, labelRc.top);

                if (incomplete)
                {
                    constexpr float kSpinnerBoxDip = 20.0f;
                    spinnerRc                      = labelRc;
                    spinnerRc.left                 = std::max(labelRc.left, labelRc.right - kSpinnerBoxDip);
                    const float spinnerW           = std::max(0.0f, spinnerRc.right - spinnerRc.left);
                    const float spinnerH           = std::max(0.0f, spinnerRc.bottom - spinnerRc.top);
                    showSpinner                    = spinnerW >= 8.0f && spinnerH >= 8.0f;
                    if (showSpinner)
                    {
                        labelRc.right = std::max(labelRc.left, spinnerRc.left - 2.0f);
                    }
                }
            }
            else
            {
                if (incomplete)
                {
                    constexpr float kSpinnerBoxDip = 20.0f;
                    spinnerRc                      = labelRc;
                    spinnerRc.left                 = std::max(labelRc.left, labelRc.right - kSpinnerBoxDip);
                    spinnerRc.bottom               = std::min(labelRc.bottom, labelRc.top + kSpinnerBoxDip);
                    const float spinnerW           = std::max(0.0f, spinnerRc.right - spinnerRc.left);
                    const float spinnerH           = std::max(0.0f, spinnerRc.bottom - spinnerRc.top);
                    showSpinner                    = spinnerW >= 8.0f && spinnerH >= 8.0f;
                    if (showSpinner)
                    {
                        labelRc.right = std::max(labelRc.left, spinnerRc.left - 2.0f);
                    }
                }

                if (rainbowMode && _brushBackground)
                {
                    const float scrimAlpha = darkMode ? 0.34f : 0.54f;
                    D2D1_RECT_F scrimRc    = labelRc;
                    scrimRc.bottom         = std::min(scrimRc.bottom, scrimRc.top + 34.0f);
                    _brushBackground->SetColor(D2D1::ColorF(bg.r, bg.g, bg.b, scrimAlpha));
                    _renderTarget->FillRectangle(scrimRc, _brushBackground.get());
                }
            }

            D2D1_RECT_F dogEarRc{};
            bool showDogEar = false;
            if (! nodeRef.isDirectory && ! nodeRef.isSynthetic && _brushOutline)
            {
                const float side = std::min(tileW, tileH);
                const float size = std::clamp(side * 0.18f, 8.0f, 14.0f);

                dogEarRc = rc;
                dogEarRc.top += 2.0f;
                dogEarRc.right -= 2.0f;
                dogEarRc.left   = std::max(dogEarRc.left, dogEarRc.right - size);
                dogEarRc.bottom = std::min(dogEarRc.bottom, dogEarRc.top + size);

                const float dogW = std::max(0.0f, dogEarRc.right - dogEarRc.left);
                const float dogH = std::max(0.0f, dogEarRc.bottom - dogEarRc.top);
                showDogEar       = dogW >= 6.0f && dogH >= 6.0f;
                if (showDogEar)
                {
                    labelRc.right = std::max(labelRc.left, std::min(labelRc.right, dogEarRc.left - 2.0f));

                    if (_brushBackground)
                    {
                        D2D1::ColorF revealFill = bg;
                        if (const std::optional<ResolvedItem> parent = ResolveTreemapItem(nodeRef.parentId); parent.has_value())
                        {
                            revealFill = baseColorForNode(parent.value());

                            const bool parentRealDir    = parent->isDirectory && ! parent->isSynthetic;
                            const bool parentIncomplete = parentRealDir && (parent->scanState == ScanState::NotStarted ||
                                                                            parent->scanState == ScanState::Queued || parent->scanState == ScanState::Scanning);
                            if (parentIncomplete && ! highContrast)
                            {
                                float dimT = darkMode ? 0.38f : 0.28f;
                                if (parent->scanState == ScanState::Queued || parent->scanState == ScanState::NotStarted)
                                {
                                    dimT = darkMode ? 0.52f : 0.40f;
                                }
                                revealFill = Mix(revealFill, bg, dimT);
                            }
                        }

                        const float revealAlpha = highContrast ? 1.0f : 0.96f;
                        _brushBackground->SetColor(D2D1::ColorF(revealFill.r, revealFill.g, revealFill.b, revealAlpha));
                        _renderTarget->FillRectangle(dogEarRc, _brushBackground.get());

                        if (_dogEarFlapGeometry)
                        {
                            const D2D1::ColorF flapFill = Mix(base, bg, darkMode ? 0.08f : 0.18f);
                            const float flapAlpha       = highContrast ? 1.0f : 0.88f;
                            _brushBackground->SetColor(D2D1::ColorF(flapFill.r, flapFill.g, flapFill.b, flapAlpha));

                            D2D1_MATRIX_3X2_F oldTransform{};
                            _renderTarget->GetTransform(&oldTransform);

                            const D2D1_MATRIX_3X2_F dogTransform =
                                D2D1::Matrix3x2F::Scale(dogW, dogH) * D2D1::Matrix3x2F::Translation(dogEarRc.left, dogEarRc.top) * oldTransform;
                            _renderTarget->SetTransform(dogTransform);
                            _renderTarget->FillGeometry(_dogEarFlapGeometry.get(), _brushBackground.get());
                            _renderTarget->SetTransform(oldTransform);
                        }
                    }

                    const D2D1::ColorF line = highContrast ? textColor : Mix(base, textColor, 0.52f);
                    _brushOutline->SetColor(D2D1::ColorF(line.r, line.g, line.b, highContrast ? 1.0f : 0.92f));
                    _renderTarget->DrawLine(
                        D2D1::Point2F(dogEarRc.left, dogEarRc.top), D2D1::Point2F(dogEarRc.right, dogEarRc.bottom), _brushOutline.get(), 1.1f);
                }
            }

            D2D1_RECT_F nameRc        = labelRc;
            D2D1_RECT_F statusRc      = {};
            D2D1_RECT_F sizeRc        = labelRc;
            const bool showStatusLine = showStatusInHeader && nodeRef.isDirectory;
            bool showSize             = false;

            if (showStatusLine)
            {
                const float availableH = std::max(0.0f, labelRc.bottom - labelRc.top);
                if (availableH >= 26.0f)
                {
                    const float nameH = std::clamp(availableH * 0.60f, 14.0f, availableH);
                    nameRc.bottom     = std::clamp(labelRc.top + nameH, labelRc.top, labelRc.bottom);

                    statusRc       = labelRc;
                    statusRc.top   = std::min(labelRc.bottom, nameRc.bottom + 1.0f);
                    statusRc.top   = std::min(statusRc.top, statusRc.bottom);
                    statusRc.left  = nameRc.left;
                    statusRc.right = nameRc.right;
                }
            }

            if (! nodeRef.isDirectory && item.lod >= DrawItem::Lod::Large)
            {
                static constexpr float kSizeLineHeightDip = 14.0f;
                const float availableH                    = std::max(0.0f, labelRc.bottom - labelRc.top);
                if (availableH >= 32.0f)
                {
                    showSize      = true;
                    sizeRc.top    = std::max(labelRc.top, labelRc.bottom - kSizeLineHeightDip);
                    nameRc.bottom = std::min(nameRc.bottom, sizeRc.top);
                }
            }

            if (! nameView.empty() && nameRc.right > nameRc.left && nameRc.bottom > nameRc.top)
            {
                _renderTarget->DrawTextW(
                    nameView.data(), static_cast<UINT32>(nameView.size()), _textFormat.get(), nameRc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (showStatusLine && statusRc.right > statusRc.left && statusRc.bottom > statusRc.top && ! watermarkText.empty())
            {
                const D2D1_COLOR_F old = _brushText->GetColor();
                const float alpha      = highContrast ? 1.0f : 0.72f;
                _brushText->SetColor(D2D1::ColorF(old.r, old.g, old.b, old.a * alpha));

                _renderTarget->DrawTextW(watermarkText.data(),
                                         static_cast<UINT32>(watermarkText.size()),
                                         _textFormat.get(),
                                         statusRc,
                                         _brushText.get(),
                                         D2D1_DRAW_TEXT_OPTIONS_CLIP);

                _brushText->SetColor(old);
            }

            if (showSize && sizeRc.right > sizeRc.left && sizeRc.bottom > sizeRc.top)
            {
                const CompactBytesText sizeText = FormatBytesCompactInline(ShownBytes(nodeRef.id, nodeRef.totalBytes));
                if (sizeText.length > 0)
                {
                    _renderTarget->DrawTextW(sizeText.buffer.data(), sizeText.length, _textFormat.get(), sizeRc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                }
            }

            if (showSpinner)
            {
                const float w              = std::max(0.0f, spinnerRc.right - spinnerRc.left);
                const float h              = std::max(0.0f, spinnerRc.bottom - spinnerRc.top);
                const float radius         = std::clamp(std::min(w, h) * 0.34f, 3.0f, 7.0f);
                const D2D1_POINT_2F center = D2D1::Point2F((spinnerRc.left + spinnerRc.right) * 0.5f, (spinnerRc.top + spinnerRc.bottom) * 0.5f);
                drawMiniSpinner(center, radius, nodeRef.id, kMiniSpinnerPhaseSpeed, kMiniSpinnerHueSpeed);
            }
        }
    }

    PaintHoverOverlay(accentColor);

    if (scanActive && _brushBackground && _brushAccent)
    {
        struct SpinnerCandidate final
        {
            uint32_t nodeId = 0;
            D2D1_RECT_F rc{};
            float area = 0.0f;
        };

        static constexpr size_t kMaxBigSpinners      = 48;
        static constexpr float kMinBigSpinnerSideDip = 36.0f;

        std::array<SpinnerCandidate, kMaxBigSpinners> spinners{};
        size_t spinnerCount = 0;

        auto tryInsertSpinner = [&](uint32_t nodeId, const D2D1_RECT_F& rc, float area) noexcept
        {
            if (spinnerCount < spinners.size())
            {
                SpinnerCandidate cand;
                cand.nodeId            = nodeId;
                cand.rc                = rc;
                cand.area              = area;
                spinners[spinnerCount] = cand;
                spinnerCount += 1;
                return;
            }

            size_t smallest    = 0;
            float smallestArea = spinners[0].area;
            for (size_t i = 1; i < spinners.size(); ++i)
            {
                if (spinners[i].area < smallestArea)
                {
                    smallestArea = spinners[i].area;
                    smallest     = i;
                }
            }

            if (area <= smallestArea)
            {
                return;
            }

            spinners[smallest].nodeId = nodeId;
            spinners[smallest].rc     = rc;
            spinners[smallest].area   = area;
        };

        for (const DrawItem& item : _drawItems)
        {
            if (item.depth != 0)
            {
                continue;
            }

            const std::optional<ResolvedItem> node = resolveItem(item.nodeId);
            if (! node.has_value())
            {
                continue;
            }

            const ResolvedItem& nodeRef = node.value();
            if (! nodeRef.isDirectory || nodeRef.isSynthetic)
            {
                continue;
            }

            if (nodeRef.scanState != ScanState::Scanning)
            {
                continue;
            }

            const D2D1_RECT_F rc = rectForItem(item);
            const float w        = std::max(0.0f, rc.right - rc.left);
            const float h        = std::max(0.0f, rc.bottom - rc.top);
            const float side     = std::min(w, h);
            if (side < kMinBigSpinnerSideDip)
            {
                continue;
            }

            const float area = w * h;
            if (area <= 1.0f)
            {
                continue;
            }

            tryInsertSpinner(nodeRef.id, rc, area);
        }

        std::sort(spinners.begin(), spinners.begin() + static_cast<ptrdiff_t>(spinnerCount), [](const SpinnerCandidate& a, const SpinnerCandidate& b) noexcept {
            return a.area > b.area;
        });

        const size_t maxWanted = std::clamp<size_t>(static_cast<size_t>(_scanEffectiveThreads), 1u, kMaxBigSpinners);

        bool drewSpinner = false;
        for (size_t i = 0; i < std::min(spinnerCount, maxWanted); ++i)
        {
            const SpinnerCandidate& cand = spinners[i];
            const float w                = std::max(0.0f, cand.rc.right - cand.rc.left);
            const float h                = std::max(0.0f, cand.rc.bottom - cand.rc.top);
            const float side             = std::min(w, h);
            if (side < kMinBigSpinnerSideDip)
            {
                continue;
            }

            const D2D1_POINT_2F center = D2D1::Point2F((cand.rc.left + cand.rc.right) * 0.5f, (cand.rc.top + cand.rc.bottom) * 0.5f);
            const float radius         = std::clamp(side * 0.12f, 10.0f, 42.0f);
            const float scrimAlpha     = highContrast ? 0.86f : (darkMode ? 0.76f : 0.62f);
            _brushBackground->SetColor(D2D1::ColorF(bg.r, bg.g, bg.b, scrimAlpha));
            _renderTarget->FillEllipse(D2D1::Ellipse(center, radius * 1.55f, radius * 1.55f), _brushBackground.get());

            drawMiniSpinner(center, radius, cand.nodeId ^ 0x9e3779b9u, kBigSpinnerPhaseSpeed, kBigSpinnerHueSpeed);
            drewSpinner = true;
        }

        if (! drewSpinner)
        {
            const D2D1_RECT_F spinnerHostRc = treemapRc;

            const float hostW    = std::max(0.0f, spinnerHostRc.right - spinnerHostRc.left);
            const float hostH    = std::max(0.0f, spinnerHostRc.bottom - spinnerHostRc.top);
            const float hostSide = std::min(hostW, hostH);
            if (hostSide >= 26.0f)
            {
                const D2D1_POINT_2F center =
                    D2D1::Point2F((spinnerHostRc.left + spinnerHostRc.right) * 0.5f, (spinnerHostRc.top + spinnerHostRc.bottom) * 0.5f);

                const float radius     = std::clamp(hostSide * 0.10f, 12.0f, 42.0f);
                const float scrimAlpha = highContrast ? 0.86f : (darkMode ? 0.76f : 0.62f);
                _brushBackground->SetColor(D2D1::ColorF(bg.r, bg.g, bg.b, scrimAlpha));
                _renderTarget->FillEllipse(D2D1::Ellipse(center, radius * 1.55f, radius * 1.55f), _brushBackground.get());

                const uint32_t seed = _scanGeneration.load(std::memory_order_acquire);
                drawMiniSpinner(center, radius, seed ^ 0x9e3779b9u, kBigSpinnerPhaseSpeed, kBigSpinnerHueSpeed);

                if (_brushText && _headerInfoFormat && (! _headerCountsText.empty() || ! _headerSizeText.empty()))
                {
                    const D2D1_COLOR_F old = _brushText->GetColor();
                    const float alpha      = highContrast ? 0.95f : (darkMode ? 0.82f : 0.78f);
                    _brushText->SetColor(D2D1::ColorF(old.r, old.g, old.b, old.a * alpha));

                    auto drawCenteredLine = [&](std::wstring_view text, float top, float height) noexcept
                    {
                        if (text.empty() || ! _dwriteFactory)
                        {
                            return;
                        }

                        const float w     = MeasureTextWidthDip(_dwriteFactory.get(), _headerInfoFormat.get(), text);
                        const float maxW  = std::max(0.0f, spinnerHostRc.right - spinnerHostRc.left);
                        const float lineW = std::clamp(w, 0.0f, maxW);

                        D2D1_RECT_F rc{};
                        rc.left   = std::clamp(center.x - lineW * 0.5f, spinnerHostRc.left, spinnerHostRc.right);
                        rc.right  = std::clamp(rc.left + lineW, spinnerHostRc.left, spinnerHostRc.right);
                        rc.top    = top;
                        rc.bottom = top + height;

                        _renderTarget->DrawTextW(
                            text.data(), static_cast<UINT32>(text.size()), _headerInfoFormat.get(), rc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    };

                    const float lineHeight = std::clamp(radius * 0.85f, 14.0f, 18.0f);
                    float textTop          = center.y + radius * 2.05f;
                    if (textTop + lineHeight <= spinnerHostRc.bottom)
                    {
                        drawCenteredLine(_headerCountsText, textTop, lineHeight);
                        textTop += lineHeight;
                        if (textTop + lineHeight <= spinnerHostRc.bottom)
                        {
                            drawCenteredLine(_headerSizeText, textTop, lineHeight);
                        }
                    }

                    _brushText->SetColor(old);
                }
            }
        }
    }

    if (! scanActive && _overallState == ScanState::Done && _scanCompletedSinceSeconds > 0.0 && _watermarkFormat && _brushText && _brushBackground)
    {
        const double elapsed = nowSeconds - _scanCompletedSinceSeconds;
        if (elapsed >= 0.0 && elapsed < kScanCompletedOverlaySeconds)
        {
            const float t    = static_cast<float>(elapsed / kScanCompletedOverlaySeconds);
            const float fade = 1.0f - std::clamp(t, 0.0f, 1.0f);

            const std::wstring overlayText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_OVERLAY_SCAN_COMPLETED);
            if (! overlayText.empty() && _dwriteFactory)
            {
                const float textW = MeasureTextWidthDip(_dwriteFactory.get(), _watermarkFormat.get(), overlayText);
                const float maxW  = std::max(0.0f, treemapRc.right - treemapRc.left);
                const float boxW  = std::clamp(textW + 28.0f, 120.0f, maxW);

                const D2D1_POINT_2F center = D2D1::Point2F((treemapRc.left + treemapRc.right) * 0.5f, (treemapRc.top + treemapRc.bottom) * 0.5f);

                D2D1_RECT_F boxRc{};
                boxRc.left   = center.x - boxW * 0.5f;
                boxRc.right  = center.x + boxW * 0.5f;
                boxRc.top    = center.y - 22.0f;
                boxRc.bottom = center.y + 22.0f;

                const float scrimAlpha = (highContrast ? 0.88f : (darkMode ? 0.74f : 0.62f)) * fade;
                _brushBackground->SetColor(D2D1::ColorF(bg.r, bg.g, bg.b, scrimAlpha));
                _renderTarget->FillRoundedRectangle(D2D1::RoundedRect(boxRc, 10.0f, 10.0f), _brushBackground.get());

                const D2D1_COLOR_F oldTextColor = _brushText->GetColor();
                const float textAlpha           = (highContrast ? 1.0f : 0.92f) * fade;
                _brushText->SetColor(D2D1::ColorF(oldTextColor.r, oldTextColor.g, oldTextColor.b, oldTextColor.a * textAlpha));

                _renderTarget->DrawTextW(
                    overlayText.c_str(), static_cast<UINT32>(overlayText.size()), _watermarkFormat.get(), boxRc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                _brushText->SetColor(oldTextColor);
            }
        }
    }

    PaintTooltipOverlay();

    const bool shouldRecordStaticCache = CanRecordStaticTreemapCache(scanActive, completionOverlayActive, nowSeconds);
    const bool staticCacheWasEmpty     = _staticTreemapCache.bitmap == nullptr;

    endDraw.reset();

    if (SUCCEEDED(drawHr) && shouldRecordStaticCache)
    {
        if (staticCacheWasEmpty)
        {
            _staticTreemapCache.misses += 1u;
            Debug::Perf::EmitCounter(L"viewer.space.render.static_cache_miss");
        }

        const HRESULT cacheHr = RecordStaticTreemapCacheFromCurrentTarget();
        if (FAILED(cacheHr))
        {
            Debug::Warning(L"ViewerSpace: static bitmap cache record failed: 0x{:08X}", static_cast<unsigned long>(cacheHr));
        }
    }

    HRESULT presentHr = S_OK;
    if (SUCCEEDED(drawHr) && _renderer.swapChain)
    {
        // Submit without a UI-thread vblank wait; the compositor and animation timer pace display.
        const auto presentStartedAt = std::chrono::steady_clock::now();
        presentHr                   = _renderer.swapChain->Present(0u, 0u);
        Debug::Perf::EmitDurationUs(L"viewer.space.render.present_us", Debug::Perf::ElapsedUs(presentStartedAt), 0u, _drawItems.size(), presentHr);
    }
#ifdef ENABLE_TESTS
    applyForcedRendererFault(drawHr, presentHr);
#endif
    const HRESULT renderHr = FAILED(drawHr) ? drawHr : presentHr;
    if (_renderer.swapChain)
        CompletePendingInputPresentation(renderHr, tileDrawCount);

    const uint64_t paintUs = Debug::Perf::ElapsedUs(paintStartedAt);
#ifdef ENABLE_TESTS
    _lastPaintUs = paintUs;
#endif
    Debug::Perf::Emit(
        L"viewer.space.render.paint_us", L"device-context", paintUs, static_cast<uint64_t>(_drawItems.size()), _scanActive.load() ? 1u : 0u, renderHr);
    Debug::Perf::Emit(L"viewer.space.render.tile_draw_count", L"lod", 0u, tileDrawCount, culledTileCount, renderHr);
    Debug::Perf::Emit(L"viewer.space.render.text_draw_count", L"lod", 0u, textDrawCandidateCount, tileDrawCount, renderHr);
    Debug::Perf::Emit(L"viewer.space.layout.visible_tiles", L"lod", 0u, tileDrawCount, culledTileCount, renderHr);

    if (! _openToFirstPaintEmitted && _openStartedAt != std::chrono::steady_clock::time_point{})
    {
        _openToFirstPaintEmitted = true;
        Debug::Perf::Emit(L"viewer.space.open_to_first_paint_us",
                          L"phase0-current",
                          Debug::Perf::ElapsedUs(_openStartedAt),
                          static_cast<uint64_t>(_drawItems.size()),
                          _scanActive.load() ? 1u : 0u,
                          renderHr);
    }

    if (drawHr == D2DERR_RECREATE_TARGET || presentHr == DXGI_ERROR_DEVICE_REMOVED || presentHr == DXGI_ERROR_DEVICE_RESET)
    {
#ifdef ENABLE_TESTS
        _rendererFailureStage = kRendererDiscardDeviceLost;
        _rendererFailureHr    = static_cast<uint32_t>(FAILED(drawHr) ? drawHr : presentHr);
#endif
        DiscardDirect2D();
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
    else if (FAILED(drawHr))
    {
        Debug::Warning(L"ViewerSpace: EndDraw failed: 0x{:08X}", static_cast<unsigned long>(drawHr));
    }
    else if (FAILED(presentHr))
    {
        Debug::Warning(L"ViewerSpace: Present failed: 0x{:08X}", static_cast<unsigned long>(presentHr));
    }
}

void ViewerSpace::BeginInputPresentation(InputPresentationKind kind, std::chrono::steady_clock::time_point startedAt) noexcept
{
    auto& pending = _pendingInputPresentation[static_cast<size_t>(kind)];
    if (pending == std::chrono::steady_clock::time_point{})
        pending = startedAt;
    else
        Debug::Perf::EmitCounter(L"viewer.space.input.coalesced_count");
}

void ViewerSpace::CompletePendingInputPresentation(HRESULT result, uint64_t visibleTileCount) noexcept
{
    _lastRenderPresentHr = static_cast<uint32_t>(result);
    if (result != S_OK)
        return;
    if (! _layoutDirty && ! _layoutBuild && ! _drawItems.empty() && visibleTileCount > 0u)
    {
#ifdef ENABLE_TESTS
        _presentedLayoutGeneration = _layoutGeneration;
#endif
        if (_firstMeaningfulPaintUs == 0u && _viewerOpenStartedAt != std::chrono::steady_clock::time_point{})
        {
            _firstMeaningfulPaintUs = Debug::Perf::ElapsedUs(_viewerOpenStartedAt);
            Debug::Perf::Emit(L"viewer.space.open_to_first_meaningful_paint_us",
                              L"published-map",
                              _firstMeaningfulPaintUs,
                              static_cast<uint64_t>(_drawItems.size()),
                              _scanActive.load() ? 1u : 0u,
                              S_OK);
        }
    }
    constexpr std::array<std::wstring_view, 3u> kinds{{L"hover", L"selection", L"navigation"}};
    for (size_t index = 0u; index < _pendingInputPresentation.size(); ++index)
    {
        auto& pending = _pendingInputPresentation[index];
        if (pending == std::chrono::steady_clock::time_point{})
            continue;
        _lastInputPresentationUs = Debug::Perf::ElapsedUs(pending);
        pending                  = {};
        ++_inputPresentationCount;
        // This ends at successful Present submission, not physical display or layout completion.
        Debug::Perf::Emit(
            L"viewer.space.input.handler_to_next_present_us", kinds[index], _lastInputPresentationUs, static_cast<uint64_t>(_drawItems.size()), index, S_OK);
    }
}

void ViewerSpace::OnCommand([[maybe_unused]] HWND hwnd, UINT commandId) noexcept
{
    if (commandId >= IDM_VIEWERSPACE_TYPE_ALL && commandId <= IDM_VIEWERSPACE_AGE_YEAR)
    {
        if (commandId <= IDM_VIEWERSPACE_TYPE_LAST)
            _typeFilter = commandId - IDM_VIEWERSPACE_TYPE_ALL;
        else
            _ageFilter = commandId - IDM_VIEWERSPACE_AGE_ANY;
        UpdateWorkspaceScope();
        InvalidateAnalysis();
        InvalidateStaticTreemapCache();
        if (_hWnd)
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        return;
    }
    switch (commandId)
    {
        case IDM_VIEWERSPACE_COPY_PATH: CopyItemPath(); break;
        case IDM_VIEWERSPACE_EXCLUDE:
            if (_selectedGeneration == _displayGeneration)
                ToggleExcludedItem(_selectedItemId);
            break;
        case IDM_VIEWERSPACE_RESTORE_ALL: RestoreExcludedItems(); break;
        case IDM_VIEWERSPACE_MARK_CLEANUP:
            if (_selectedGeneration == _displayGeneration)
                ToggleCleanupMark(_selectedItemId);
            break;
        case IDM_VIEWERSPACE_MARK_MATCHING: MarkMatchingItems(); break;
        case IDM_VIEWERSPACE_REVIEW_CLEANUP: ReviewCleanup(); break;
        case IDM_VIEWERSPACE_SEND_RECYCLE: SendReviewedCleanup(true); break;
        case IDM_VIEWERSPACE_SEND_DELETE: SendReviewedCleanup(false); break;
        case IDM_VIEWERSPACE_CLEAR_MARKS:
            _cleanupMarks.clear();
            _reviewingCleanup = false;
            UpdateWorkspaceScope();
            InvalidateAnalysis();
            break;
        case IDM_VIEWERSPACE_OPEN_OTHER_PANE: OpenFolderInOtherPane(); break;
        case IDM_VIEWERSPACE_EXPORT: ExportCurrentView(false); break;
        case IDM_VIEWERSPACE_PROBLEMS: ShowScanProblems(); break;
        case IDM_VIEWERSPACE_FILTERS: ShowWorkspaceFilters(false); break;
        case IDM_VIEWERSPACE_TYPE_BREAKDOWN: ShowWorkspaceFilters(true); break;
        case IDM_VIEWERSPACE_WEIGHT_DISK: ShowWorkspaceMessage(IDS_VIEWERSPACE_DISK_UNAVAILABLE); break;
        case IDM_VIEWERSPACE_WEIGHT_SIZE:
        case IDM_VIEWERSPACE_WEIGHT_COUNT:
            if (commandId == IDM_VIEWERSPACE_WEIGHT_COUNT && ! _workspaceFactsReady)
            {
                ShowWorkspaceMessage(IDS_VIEWERSPACE_FACTS_PENDING);
                break;
            }
            _config.fileCountWeight = commandId == IDM_VIEWERSPACE_WEIGHT_COUNT;
            UpdateConfigurationJson();
            UpdateWorkspaceScope();
            WorkspaceChanged();
            break;
        case IDM_VIEWERSPACE_COLOR_FOLDER:
        case IDM_VIEWERSPACE_COLOR_TYPE:
        case IDM_VIEWERSPACE_COLOR_AGE:
        case IDM_VIEWERSPACE_COLOR_NEWEST:
            _config.colorMode = commandId - IDM_VIEWERSPACE_COLOR_FOLDER;
            UpdateConfigurationJson();
            UpdateWorkspaceScope();
            InvalidateAnalysis();
            InvalidateStaticTreemapCache();
            break;
        case IDM_VIEWERSPACE_FILE_REFRESH: RefreshCurrent(); break;
        case IDM_VIEWERSPACE_FILE_REBUILD: RebuildCurrent(); break;
        case IDM_VIEWERSPACE_NAV_UP: NavigateUp(); break;
        case IDM_VIEWERSPACE_NAV_BACK: NavigateHistory(false); break;
        case IDM_VIEWERSPACE_NAV_FORWARD: NavigateHistory(true); break;
        case IDM_VIEWERSPACE_VIEW_EXPLORER: SetViewMode(ViewMode::Explorer); break;
        case IDM_VIEWERSPACE_VIEW_ANALYST: SetViewMode(ViewMode::Analyst); break;
        case IDM_VIEWERSPACE_VIEW_SHOW_ANALYSIS: ToggleAnalysisPanel(); break;
        case IDM_VIEWERSPACE_FILE_EXIT: static_cast<void>(Close()); break;
    }
}

void ViewerSpace::CopyItemPath() noexcept
{
    if (! _model)
        return;
    if (_selectedItemId && _selectedGeneration != _displayGeneration)
    {
        ShowWorkspaceMessage(IDS_VIEWERSPACE_TARGET_UNAVAILABLE);
        return;
    }
    const uint32_t id = _selectedItemId != 0u && _selectedGeneration == _displayGeneration ? _selectedItemId : _model->viewNodeId;
    const auto item   = ResolveAnalysisItem(id);
    if (! item.has_value() || item.value().isSynthetic)
        return;
    const auto path = BuildItemPathText(id);
    if (! Common::Clipboard::TrySetUnicodeText(_hWnd.get(), path, Common::Clipboard::EmptyUnicodeTextPolicy::Reject))
    {
        Debug::Warning(L"ViewerSpace Copy path could not publish Unicode text.");
        ShowWorkspaceMessage(IDS_VIEWERSPACE_COPY_FAILED);
    }
}

void ViewerSpace::OnKeyDown(WPARAM vk, bool alt) noexcept
{
    if (vk == VK_ESCAPE && (_reviewingCleanup || _showingProblems))
    {
        _reviewingCleanup = _showingProblems = false;
        UpdateWorkspaceScope();
        InvalidateAnalysis();
        return;
    }
    if (vk == VK_ESCAPE && (_typeFilter || _ageFilter))
    {
        _typeFilter = _ageFilter = 0u;
        UpdateWorkspaceScope();
        InvalidateAnalysis();
        InvalidateStaticTreemapCache();
        return;
    }
    if (alt && vk == VK_INSERT)
    {
        CopyItemPath();
        return;
    }
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0 && (vk == '1' || vk == '2'))
    {
        SetViewMode(vk == '1' ? ViewMode::Explorer : ViewMode::Analyst);
        return;
    }
    if (alt && (vk == VK_LEFT || vk == VK_RIGHT))
    {
        NavigateHistory(vk == VK_RIGHT);
        return;
    }
    if (vk == VK_RETURN && _selectedItemId)
    {
        ActivateItem(_selectedItemId);
        return;
    }
    if (vk == VK_TAB)
    {
        static_cast<void>(_navigationHost.Focus((GetKeyState(VK_SHIFT) & 0x8000) != 0));
        return;
    }
    if (! alt && (vk == VK_LEFT || vk == VK_RIGHT || vk == VK_UP || vk == VK_DOWN))
    {
        const auto current = std::find_if(_drawItems.begin(), _drawItems.end(), [this](const DrawItem& item) { return item.nodeId == _selectedItemId; });
        const bool reverse = vk == VK_LEFT || vk == VK_UP;
        if (! _drawItems.empty())
        {
            size_t index = current == _drawItems.end() ? 0u : static_cast<size_t>(current - _drawItems.begin());
            if (reverse && index > 0u)
            {
                --index;
            }
            else if (! reverse && index + 1u < _drawItems.size())
            {
                ++index;
            }
            SelectItem(_drawItems[index].nodeId);
        }
        return;
    }
    if (vk == VK_ESCAPE)
    {
        if (_deferredStartScan.has_value() || _refreshPending || _refreshAdoptionPending || _overallState == ScanState::Queued ||
            _overallState == ScanState::Scanning)
        {
            CancelScanByUser();
            return;
        }

        if (_hWnd)
        {
            PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0);
        }
        return;
    }

    if (vk == VK_F5)
    {
        RefreshCurrent();
        return;
    }

    if (vk == VK_BACK || (alt && vk == VK_UP))
    {
        NavigateUp();
        return;
    }
}

void ViewerSpace::OnMouseMove(int x, int y) noexcept
{
    const auto handlerStartedAt = std::chrono::steady_clock::now();
    if (! _hWnd)
    {
        return;
    }

    _lastMouseMoveClientPoint = POINT{x, y};
    POINT screenPoint{x, y};
    _hasLastMouseMoveClientPoint = ClientToScreen(_hWnd.get(), &screenPoint) != FALSE;
    if (_hasLastMouseMoveClientPoint)
    {
        _lastMouseMoveScreenPoint = screenPoint;
    }

    if (! _trackingMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize     = sizeof(tme);
        tme.dwFlags    = TME_LEAVE;
        tme.hwndTrack  = _hWnd.get();
        _trackingMouse = TrackMouseEvent(&tme) != FALSE;
    }

    const float xDip         = DipFromPx(x);
    const float yDip         = DipFromPx(y);
    const float headerTop    = GetHeaderTopDip();
    const float headerBottom = GetHeaderBottomDip();

    HeaderHit newHeaderHit = HeaderHit::None;
    if (yDip >= headerTop && yDip <= headerBottom)
    {
        if ((_refreshPending || _refreshAdoptionPending || _overallState == ScanState::Queued || _overallState == ScanState::Scanning) &&
            xDip >= DipFromPx(_clientSize.cx) - kHeaderButtonWidthDip)
        {
            newHeaderHit = HeaderHit::Cancel;
        }
    }

    if (newHeaderHit != _hoverHeaderHit)
    {
        BeginInputPresentation(InputPresentationKind::Hover, handlerStartedAt);
        _hoverHeaderHit = newHeaderHit;
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    const std::optional<uint32_t> hit = HitTestTreemap(xDip, yDip);
    const uint32_t newHover           = hit.value_or(0);
    if (newHover != _hoverNodeId)
    {
        BeginInputPresentation(InputPresentationKind::Hover, handlerStartedAt);
        _hoverNodeId = newHover;
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    const uint32_t tooltipNodeId = yDip >= headerBottom ? newHover : 0u;

    if (tooltipNodeId == 0)
    {
        _tooltipCandidateNodeId       = 0;
        _tooltipCandidateSinceSeconds = 0.0;
        UpdateTooltipForHit(0);
        return;
    }

    if (tooltipNodeId == _tooltipNodeId)
    {
        _tooltipCandidateNodeId       = tooltipNodeId;
        _tooltipCandidateSinceSeconds = 0.0;
        UpdateTooltipPosition(x, y);
        return;
    }

    if (_tooltipNodeId != 0)
    {
        UpdateTooltipForHit(0);
    }

    const double nowSeconds = NowSeconds();
    if (tooltipNodeId != _tooltipCandidateNodeId)
    {
        _tooltipCandidateNodeId       = tooltipNodeId;
        _tooltipCandidateSinceSeconds = nowSeconds;
        EnsureAnimationTimer();
        return;
    }

    if (_tooltipCandidateSinceSeconds <= 0.0)
    {
        _tooltipCandidateSinceSeconds = nowSeconds;
        EnsureAnimationTimer();
        return;
    }

    if ((nowSeconds - _tooltipCandidateSinceSeconds) >= kTooltipHoverStabilityDelaySeconds)
    {
        UpdateTooltipForHit(tooltipNodeId);
        UpdateTooltipPosition(x, y);
    }
}

void ViewerSpace::OnMouseLeave() noexcept
{
    _trackingMouse                = false;
    _hasLastMouseMoveClientPoint  = false;
    _lastMouseMoveScreenPoint     = {};
    _hoverHeaderHit               = HeaderHit::None;
    _tooltipCandidateNodeId       = 0;
    _tooltipCandidateSinceSeconds = 0.0;
    UpdateTooltipForHit(0);
    if (_hoverNodeId != 0 && _hWnd)
    {
        _hoverNodeId = 0;
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::OnLButtonDown(int x, int y) noexcept
{
    const auto handlerStartedAt = std::chrono::steady_clock::now();
    _lastLButtonDownNavigated   = false;
#ifdef ENABLE_TESTS
    _lastLButtonDownHitNodeId = 0u;
#endif

    const float xDip         = DipFromPx(x);
    const float yDip         = DipFromPx(y);
    const float headerTop    = GetHeaderTopDip();
    const float headerBottom = GetHeaderBottomDip();
    if (yDip >= headerTop && yDip <= headerBottom)
    {
        if ((_refreshPending || _refreshAdoptionPending || _overallState == ScanState::Queued || _overallState == ScanState::Scanning) &&
            xDip >= DipFromPx(_clientSize.cx) - kHeaderButtonWidthDip)
        {
            CancelScanByUser();
        }
        else if (_scanFailedFolders && xDip >= _headerProblemsBounds.left && xDip < _headerProblemsBounds.right && yDip >= _headerProblemsBounds.top &&
                 yDip < _headerProblemsBounds.bottom)
            ShowScanProblems();
        return;
    }

    const std::optional<uint32_t> hit = HitTestTreemap(xDip, yDip);
#ifdef ENABLE_TESTS
    _lastLButtonDownHitNodeId = hit.value_or(0u);
#endif
    if (! hit.has_value())
    {
        return;
    }

    const uint32_t nodeId = hit.value();
    BeginInputPresentation(InputPresentationKind::Selection, handlerStartedAt);
    SelectItem(nodeId);
    const std::optional<ResolvedItem> node = ResolveTreemapItem(nodeId);
    if (node.has_value() && node->isDirectory && ! node->isSynthetic && ! node->isFileRecord)
    {
        const uint32_t previousViewNodeId = _model->viewNodeId;
        NavigateTo(nodeId);
        _lastLButtonDownNavigated = _model->viewNodeId != previousViewNodeId;
    }
}

void ViewerSpace::OnLButtonDblClk(int x, int y) noexcept
{
    const float xDip         = DipFromPx(x);
    const float yDip         = DipFromPx(y);
    const float headerTop    = GetHeaderTopDip();
    const float headerBottom = GetHeaderBottomDip();
    if (yDip >= headerTop && yDip <= headerBottom)
    {
        OnLButtonDown(x, y);
        return;
    }

    // Windows delivers WM_LBUTTONDOWN before WM_LBUTTONDBLCLK. When that first click already drilled into a folder, the double
    // click has done its job; hit-testing again would target whatever tile now lies under the pointer in the new layout.
    const bool firstClickNavigated = std::exchange(_lastLButtonDownNavigated, false);
    if (firstClickNavigated)
    {
        return;
    }

    const std::optional<uint32_t> hit = HitTestTreemap(xDip, yDip);
    if (! hit.has_value())
    {
        return;
    }

    const uint32_t nodeId                  = hit.value();
    const std::optional<ResolvedItem> node = ResolveTreemapItem(nodeId);
    if (! node.has_value())
    {
        return;
    }

    if (node->isDirectory && ! node->isSynthetic && ! node->isFileRecord)
    {
        NavigateTo(nodeId);
        return;
    }

    if (! node->isSynthetic)
    {
        return;
    }

    const uint32_t parentId = node->parentId;
    if (parentId == 0)
    {
        return;
    }

    // An "Other" bucket is synthetic, never a model node: raise its parent's layout budget and show the parent. Files that the
    // scan aggregated cannot be itemized by layout alone; a subtree rescan with a higher budget belongs to the incremental model.
    uint32_t currentLimit = static_cast<uint32_t>(kDefaultLayoutItems);
    const auto limitIt    = _layoutMaxItemsByNode.find(parentId);
    if (limitIt != _layoutMaxItemsByNode.end())
    {
        currentLimit = limitIt->second;
    }

    uint32_t nextLimit = currentLimit;
    if (nextLimit < kMaxLayoutItems)
    {
        nextLimit = nextLimit == 0 ? static_cast<uint32_t>(kDefaultLayoutItems) : std::min(static_cast<uint32_t>(kMaxLayoutItems), nextLimit * 2u);
    }

    if (nextLimit != currentLimit)
    {
        _layoutMaxItemsByNode[parentId] = nextLimit;
    }

    if (parentId != _model->viewNodeId)
    {
        NavigateTo(parentId);
        return;
    }

    _layoutDirty = true;
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::OnContextMenu(HWND hwnd, POINT screenPt) noexcept
{
#if defined(ENABLE_TESTS)
    _debugHasLastContextMenuScreenPoint = true;
    _debugLastContextMenuScreenPoint    = screenPt;
    _debugLastContextMenuHitNodeId      = 0;
#endif

    if (! hwnd)
    {
        return;
    }

    if (! _menuHandle && ! GetMenu(hwnd))
    {
        _menuHandle.reset(Localization::LoadMenuResource(g_hInstance, IDR_VIEWERSPACE_MENU));
    }

    HMENU viewerMenu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
    if (viewerMenu)
    {
        UpdateMenuState(hwnd, false);
    }

    static constexpr std::array<int, 1> kPreviewContextMenuExcludedCommandIds{{
        IDM_VIEWERSPACE_FILE_EXIT,
    }};
    DxUi::NativeMenuFlyoutOptions previewMenuOptions{};
    previewMenuOptions.includeAcceleratorText = false;
    previewMenuOptions.omitEmptySubmenus      = true;
    previewMenuOptions.trimSeparators         = true;
    previewMenuOptions.excludedCommandIds     = kPreviewContextMenuExcludedCommandIds;

    POINT clientPt = screenPt;
    if (ScreenToClient(hwnd, &clientPt) == 0)
    {
        return;
    }

    const float xDip = DipFromPx(clientPt.x);
    const float yDip = DipFromPx(clientPt.y);

    const std::optional<uint32_t> hit = _contextItemIdOverride ? std::optional<uint32_t>(_contextItemIdOverride) : HitTestTreemap(xDip, yDip);
#if defined(ENABLE_TESTS)
    _debugLastContextMenuHitNodeId = hit.value_or(0);
#endif
    if (! hit.has_value())
    {
        HMENU menu = viewerMenu;
        if (! menu)
        {
            return;
        }

        const auto result =
            DxUi::ShowNativeHMenuContextMenu(hwnd,
                                             screenPt,
                                             menu,
                                             _hasTheme ? RedSalamander::MakeThemePaletteFromViewerTheme(_theme) : DxUi::MakeDefaultThemePalette(false),
                                             previewMenuOptions);
        if (result.has_value() && result.value() > 0)
        {
            OnCommand(hwnd, static_cast<UINT>(result.value()));
        }
        return;
    }

    const uint32_t nodeId = hit.value();

    const std::optional<ResolvedItem> node = ResolveTreemapItem(nodeId);
    if (! node.has_value())
    {
        return;
    }

    const uint32_t contextGeneration = _displayGeneration;
    SelectItem(nodeId, true);
    const std::wstring focusText   = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_CONTEXT_FOCUS_IN_PANE);
    const std::wstring zoomInText  = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_CONTEXT_ZOOM_IN);
    const std::wstring zoomOutText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_CONTEXT_ZOOM_OUT);

    wil::unique_hmenu rootMenu(Localization::LoadMenuResource(GetModuleHandleW(nullptr), kHostFolderViewItemContextMenuResourceId));
    if (! rootMenu)
    {
        return;
    }

    HMENU menu = GetSubMenu(rootMenu.get(), 0);
    if (! menu)
    {
        return;
    }

    // Remove View Space + any debug-only items.
    static_cast<void>(DeleteMenu(menu, kCmdFolderViewContextViewSpace, MF_BYCOMMAND));

    auto menuContainsDebugCommands = [&](HMENU m, auto&& self) noexcept -> bool
    {
        if (! m)
        {
            return false;
        }

        const int count = GetMenuItemCount(m);
        if (count <= 0)
        {
            return false;
        }

        for (int pos = 0; pos < count; ++pos)
        {
            const UINT id = GetMenuItemID(m, pos);
            if (id != static_cast<UINT>(-1) && id >= kFolderViewDebugCommandIdBase)
            {
                return true;
            }

            HMENU sub = GetSubMenu(m, pos);
            if (sub && self(sub, self))
            {
                return true;
            }
        }

        return false;
    };

    const int topCount = GetMenuItemCount(menu);
    if (topCount > 0)
    {
        for (int pos = 0; pos < topCount; ++pos)
        {
            HMENU sub = GetSubMenu(menu, pos);
            if (! sub)
            {
                continue;
            }

            if (! menuContainsDebugCommands(sub, menuContainsDebugCommands))
            {
                continue;
            }

            // Remove the debug popup and any separator directly above it.
            DeleteMenu(menu, static_cast<UINT>(pos), MF_BYPOSITION);

            const int sepPos = pos - 1;
            if (sepPos >= 0)
            {
                MENUITEMINFOW info{};
                info.cbSize = sizeof(info);
                info.fMask  = MIIM_FTYPE;
                if (GetMenuItemInfoW(menu, static_cast<UINT>(sepPos), TRUE, &info) != 0 && (info.fType & MFT_SEPARATOR) != 0)
                {
                    RemoveMenu(menu, static_cast<UINT>(sepPos), MF_BYPOSITION);
                }
            }
            break;
        }
    }

    // Insert our treemap commands at the top (reverse order).
    InsertMenuW(menu, 0, MF_BYPOSITION | MF_SEPARATOR, 0, nullptr);
    InsertMenuW(menu, 0, MF_BYPOSITION | MF_STRING, kCmdTreemapContextZoomOut, zoomOutText.c_str());
    InsertMenuW(menu, 0, MF_BYPOSITION | MF_STRING, kCmdTreemapContextZoomIn, zoomInText.c_str());
    InsertMenuW(menu, 0, MF_BYPOSITION | MF_STRING, kCmdTreemapContextFocusInPane, focusText.c_str());

    const bool markedForCleanup =
        _cleanupGeneration == _displayGeneration && std::find(_cleanupMarks.begin(), _cleanupMarks.end(), nodeId) != _cleanupMarks.end();
    const std::array<std::pair<UINT, UINT>, 4u> workspaceCommands{
        {{IDM_VIEWERSPACE_OPEN_OTHER_PANE, IDS_VIEWERSPACE_OPEN_OTHER_PANE},
         {IDM_VIEWERSPACE_MARK_CLEANUP, markedForCleanup ? IDS_VIEWERSPACE_UNMARK_CLEANUP : IDS_VIEWERSPACE_MARK_CLEANUP},
         {IDM_VIEWERSPACE_EXCLUDE, IsItemExcluded(nodeId) ? IDS_VIEWERSPACE_RESTORE_ITEM : IDS_VIEWERSPACE_EXCLUDE},
         {IDM_VIEWERSPACE_COPY_PATH, IDS_VIEWERSPACE_COPY_PATH}}};
    InsertMenuW(menu, 0, MF_BYPOSITION | MF_SEPARATOR, 0, nullptr);
    for (const auto& [command, label] : workspaceCommands)
    {
        const std::wstring text = LoadStringResource(g_hInstance, label);
        InsertMenuW(menu, 0, MF_BYPOSITION | MF_STRING, command, text.c_str());
    }
    const bool real = ! node->isSynthetic;
    EnableMenuItem(menu, IDM_VIEWERSPACE_COPY_PATH, static_cast<UINT>(MF_BYCOMMAND | (real ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(
        menu,
        IDM_VIEWERSPACE_EXCLUDE,
        static_cast<UINT>(MF_BYCOMMAND | (real && ! _workspaceRebinding && _model && _model->workspacePathIdentity.has_value() ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu,
                   IDM_VIEWERSPACE_MARK_CLEANUP,
                   static_cast<UINT>(MF_BYCOMMAND | (real && (markedForCleanup || ! IsItemExcluded(nodeId)) ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(
        menu, IDM_VIEWERSPACE_OPEN_OTHER_PANE, static_cast<UINT>(MF_BYCOMMAND | (real && node->isDirectory && _hostViewerActions ? MF_ENABLED : MF_GRAYED)));

    // Resolve host navigation target.
    std::wstring folderPath;
    std::wstring folderPathForCommand;
    std::wstring focusItemForCommand;

    const std::wstring nodePath = BuildItemPathText(nodeId);
    const bool nodeHasPath      = ! nodePath.empty();

    if (nodeHasPath)
    {
        if (_fileSystemIsWin32)
        {
            const std::filesystem::path full(nodePath);
            const std::wstring leaf = full.filename().wstring();
            if (const std::optional<std::filesystem::path> parent = TryGetParentPathForNavigation(full); parent.has_value() && ! leaf.empty())
            {
                folderPathForCommand = parent.value().wstring();
                focusItemForCommand  = leaf;
            }
            else if (node->isDirectory)
            {
                folderPath = nodePath;
            }
        }
        else
        {
            std::wstring_view trimmed = TrimTrailingPathSeparators(nodePath);
            const size_t lastSep      = trimmed.find_last_of(L"/\\");
            if (lastSep != std::wstring_view::npos && (lastSep + 1) < trimmed.size())
            {
                const std::wstring_view leaf = trimmed.substr(lastSep + 1);
                if (const std::optional<std::wstring> parent = TryGetParentPathForNavigationGeneric(trimmed); parent.has_value())
                {
                    folderPathForCommand = parent.value();
                    focusItemForCommand.assign(leaf.data(), leaf.size());
                }
                else if (node->isDirectory)
                {
                    folderPath = nodePath;
                }
            }
            else if (node->isDirectory)
            {
                folderPath = nodePath;
            }
        }

        if (folderPath.empty())
        {
            folderPath = folderPathForCommand;
        }
    }

    const bool canZoomIn  = node->isDirectory && ! node->isSynthetic && ! node->isFileRecord;
    const bool canZoomOut = CanNavigateUp();

    const bool hostAvailable = _fileSystemShortId == L"file" && static_cast<bool>(_hostPaneExecute);
    const bool canFocusInPane =
        ! node->isSynthetic && ! folderPath.empty() && (static_cast<bool>(_hostViewerActions) || (_fileSystemShortId == L"file" && hostAvailable));
    const bool canExecuteLeafCmds = hostAvailable && ! folderPathForCommand.empty() && ! focusItemForCommand.empty();
    const bool canExecutePaste    = hostAvailable && ! folderPath.empty();

    EnableMenuItem(menu, kCmdTreemapContextZoomIn, static_cast<UINT>(MF_BYCOMMAND | (canZoomIn ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, kCmdTreemapContextZoomOut, static_cast<UINT>(MF_BYCOMMAND | (canZoomOut ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, kCmdTreemapContextFocusInPane, static_cast<UINT>(MF_BYCOMMAND | (canFocusInPane ? MF_ENABLED : MF_GRAYED)));

    const bool isDirectory = node->isDirectory;

    UpdateHostFolderViewContextCommandState(menu, canExecuteLeafCmds, isDirectory, canExecutePaste);

    std::vector<DxUi::MenuFlyoutItem> popupItems = ConvertHMenuToDxFlyoutItems(menu);
    if (viewerMenu)
    {
        std::vector<DxUi::MenuFlyoutItem> viewerItems = DxUi::ConvertNativeHMenuToFlyoutItems(viewerMenu, previewMenuOptions);
        if (! viewerItems.empty())
        {
            if (! popupItems.empty())
            {
                DxUi::MenuFlyoutItem separator{};
                separator.kind = DxUi::MenuItemKind::Separator;
                popupItems.push_back(std::move(separator));
            }

            for (auto& viewerItem : viewerItems)
            {
                popupItems.push_back(std::move(viewerItem));
            }
        }
    }

    if (popupItems.empty())
    {
        return;
    }

    const auto result = DxUi::ContextMenu::Show(
        hwnd, screenPt, popupItems, _hasTheme ? RedSalamander::MakeThemePaletteFromViewerTheme(_theme) : DxUi::MakeDefaultThemePalette(false));
    if (! result.has_value())
    {
        return;
    }
    const UINT commandId = static_cast<UINT>(result.value());
    if (_displayGeneration != contextGeneration)
        return;
    if (commandId >= IDM_VIEWERSPACE_FILE_REFRESH && commandId <= IDM_VIEWERSPACE_AGE_YEAR)
    {
        OnCommand(hwnd, commandId);
        return;
    }

    if (commandId == kCmdTreemapContextZoomIn)
    {
        if (canZoomIn)
        {
            NavigateTo(nodeId);
        }
        return;
    }

    if (commandId == kCmdTreemapContextZoomOut)
    {
        if (canZoomOut)
        {
            NavigateUp();
        }
        return;
    }

    if (commandId == kCmdTreemapContextFocusInPane)
    {
        if (! canFocusInPane)
        {
            return;
        }

        FocusItemInPane(nodeId);
        return;
    }

    switch (commandId)
    {
        case IDM_VIEWERSPACE_FILE_REFRESH:
        case IDM_VIEWERSPACE_NAV_UP:
        case IDM_VIEWERSPACE_FILE_EXIT: OnCommand(hwnd, commandId); return;
        default: break;
    }

    if (! canExecutePaste && commandId == kCmdFolderViewContextPaste)
    {
        return;
    }

    if (! canExecuteLeafCmds && commandId != kCmdFolderViewContextPaste)
    {
        return;
    }

    if (! _hostPaneExecute)
    {
        return;
    }

    HostPaneExecuteRequest request{};
    request.sizeBytes            = sizeof(request);
    request.flags                = HOST_PANE_EXECUTE_FLAG_ACTIVATE_WINDOW;
    request.folderPath           = (commandId == kCmdFolderViewContextPaste) ? folderPath.c_str() : folderPathForCommand.c_str();
    request.focusItemDisplayName = (commandId == kCmdFolderViewContextPaste || focusItemForCommand.empty()) ? nullptr : focusItemForCommand.c_str();
    request.folderViewCommandId  = commandId;

    static_cast<void>(_hostPaneExecute->ExecuteInActivePane(&request));
}

LRESULT ViewerSpace::OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    std::array<uint64_t, 5u> stages{};
    static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
    _scanCacheRestoreContinuationPending = false;
    stages[0]                            = Debug::Perf::ElapsedUs(startedAt);
    _menuBarHost.Detach();
    stages[1] = Debug::Perf::ElapsedUs(startedAt);
    _menuHandle.reset();
    _embeddedMode = false;
    _hWnd.release();
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    stages[2] = Debug::Perf::ElapsedUs(startedAt);

    const LRESULT result               = DefWindowProcW(hwnd, WM_NCDESTROY, wp, lp);
    stages[3]                          = Debug::Perf::ElapsedUs(startedAt);
    const uint32_t previousWindowCount = g_viewerSpaceWindowCount.load(std::memory_order_acquire);
    if (previousWindowCount != 0 && g_viewerSpaceWindowCount.fetch_sub(1, std::memory_order_acq_rel) == 1u)
    {
        UnregisterWndClassIfIdle();
    }

    stages[4] = Debug::Perf::ElapsedUs(startedAt);
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"viewer.space.close.nc_destroy_stages_us",
                          std::format(L"drain={};menuHost={};windowState={};defaultProc={};unregister={}",
                                      stages[0],
                                      stages[1] - stages[0],
                                      stages[2] - stages[1],
                                      stages[3] - stages[2],
                                      stages[4] - stages[3]),
                          stages[4],
                          0u,
                          0u,
                          S_OK);
    }
    Release();
    return result;
}

void ViewerSpace::OnTimer(UINT_PTR timerId) noexcept
{
    if (timerId != kTimerAnimationId || ! _hWnd)
    {
        return;
    }

    ReapFinishedScanWorkers();
    ReapFinishedDeferredDisposals();
    ContinueDeferredDisposalSubmissions();
    ContinueDeferredStartScan();
    ContinueDeferredModelRetirement();
    DrainUpdates();
    ContinueScanCacheRestore();
    ContinueRefreshAdoption();
    ContinueWorkspaceBindings();
    ContinueWorkspaceFacts();
    MaybeRebuildLayout();
    ContinueAnalysisBuild();
    ContinueWorkspaceExport();

    const double now = NowSeconds();

    // A stationary pointer sends no second WM_MOUSEMOVE. Finish its stability
    // delay on the existing timer, using current painted hit geometry and the
    // last delivered client point while native leave tracking remains armed.
    // Keep the delayed tooltip separate from the immediate-input latency sample.
    if (_trackingMouse && _hasLastMouseMoveClientPoint && (_tooltipCandidateNodeId != 0u || _tooltipNodeId != 0u))
    {
        POINT currentClientPoint = _lastMouseMoveScreenPoint;
        if (! _hWnd || ScreenToClient(_hWnd.get(), &currentClientPoint) == FALSE)
        {
            _hasLastMouseMoveClientPoint  = false;
            _lastMouseMoveScreenPoint     = {};
            _tooltipCandidateNodeId       = 0u;
            _tooltipCandidateSinceSeconds = 0.0;
            UpdateTooltipForHit(0u);
        }
        else
        {
            const auto hit = HitTestTreemap(DipFromPx(currentClientPoint.x), DipFromPx(currentClientPoint.y));
            if (_tooltipNodeId != 0u)
            {
                if (hit.has_value() && hit.value() == _tooltipNodeId)
                {
                    UpdateTooltipPosition(currentClientPoint.x, currentClientPoint.y);
                }
                else
                {
                    UpdateTooltipForHit(0u);
                }
            }
            if (_tooltipCandidateNodeId != 0u && _tooltipCandidateNodeId != _tooltipNodeId && _hoverNodeId == _tooltipCandidateNodeId &&
                _tooltipCandidateSinceSeconds > 0.0 && now - _tooltipCandidateSinceSeconds >= kTooltipHoverStabilityDelaySeconds && hit.has_value() &&
                hit.value() == _tooltipCandidateNodeId)
            {
                UpdateTooltipForHit(hit.value());
                UpdateTooltipPosition(currentClientPoint.x, currentClientPoint.y);
            }
        }
    }

    if (_scanActive.load() || _overallState == ScanState::Queued || _overallState == ScanState::Scanning || _refreshAdoptionPending)
    {
        static constexpr double kMinInvalidateIntervalSecondsWhileScanning = 1.0 / 30.0;

        const double sinceLast = now - _lastScanInvalidateSeconds;
        if (_lastScanInvalidateSeconds <= 0.0 || sinceLast >= kMinInvalidateIntervalSecondsWhileScanning)
        {
            _lastScanInvalidateSeconds = now;
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        return;
    }

    if (! _layoutDirty && ! _refreshPending)
    {
        ContinueScanCacheBuild();
    }

    if (_scanCompletedSinceSeconds > 0.0)
    {
        const double elapsed = now - _scanCompletedSinceSeconds;
        if (elapsed < kScanCompletedOverlaySeconds)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
            return;
        }

        _scanCompletedSinceSeconds = 0.0;
        // Remove the final completion overlay before going idle. Its last paint
        // cannot record the static bitmap, so expiry must schedule a clean frame.
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
        return;
    }

    // Keep invalidating until paint has actually snapped every rectangle to its
    // target. Age alone can expire between paints and leave the last frame
    // permanently short of the final geometry.
    if (HasActiveLayoutAnimation(now))
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
        return;
    }

    // Nothing left to drain, lay out, animate or cache: stop the 16 ms timer so an idle viewer does not wake the CPU.
    // Every paint re-arms it (EnsureAnimationTimer), and every model or view change invalidates the window.
    if (CanStopAnimationTimer(now))
    {
        KillTimer(_hWnd.get(), kTimerAnimationId);
        _animationTimerArmed = false;
    }
}

void ViewerSpace::EnsureAnimationTimer() noexcept
{
    // SetTimer on a live id restarts its countdown, so re-arm only a stopped timer.
    if (_animationTimerArmed || ! _hWnd)
    {
        return;
    }
    _animationTimerArmed = SetTimer(_hWnd.get(), kTimerAnimationId, kAnimationIntervalMs, nullptr) != 0;
}

bool ViewerSpace::CanStopAnimationTimer(double nowSeconds) noexcept
{
    // Keep only the bounded stability wait alive. An expired candidate that is
    // no longer under the actual pointer must not create permanent idle wakeups.
    // Use the same timestamp as tooltip promotion above: a second clock read
    // could cross the deadline and stop the timer before its promotion turn.
    if (_tooltipCandidateNodeId != 0u && _tooltipCandidateNodeId != _tooltipNodeId && _tooltipCandidateSinceSeconds > 0.0 &&
        nowSeconds - _tooltipCandidateSinceSeconds < kTooltipHoverStabilityDelaySeconds)
    {
        return false;
    }
    if (_exportProjectionRequested || _workspaceExport || _workspaceFactsBuilding || _workspaceRebinding || _analysisBuild || _analysisDirty || _layoutBuild ||
        _cancelledLayoutBuild || _scanCacheRestoreSnapshot || _refreshAdoptionPending || _deferredStartScan.has_value() || _workingModelRetirementPending ||
        _layoutDirty || _scanActive.load() || _overallState == ScanState::Queued || _overallState == ScanState::Scanning || _scanCompletedSinceSeconds > 0.0)
    {
        return false;
    }

    if (_scanWorker.thread.joinable() || ! _retiredScanWorkers.empty())
    {
        return false;
    }

    if (! _deferredDisposalWorkers.empty())
    {
        return false;
    }

    {
        std::scoped_lock lock(_updateMutex);
        if (! _pendingUpdates.empty())
        {
            return false;
        }
    }

    // Mirrors ContinueScanCacheBuild: a completed Win32 scan is copied into the scan cache a slice per tick.
    const bool cacheApplicable = _config.cacheEnabled && ! _model->scanRootPath.empty() && _fileSystemIsWin32 &&
                                 g_cacheEnabled.load(std::memory_order_acquire) && g_cacheMaxEntries.load(std::memory_order_acquire) != 0u &&
                                 _overallState == ScanState::Done;
    if (_scanCacheBuildSnapshot || (cacheApplicable && _scanCacheLastStoredGeneration != _scanGeneration.load(std::memory_order_acquire)))
    {
        return false;
    }

    return true;
}

bool ViewerSpace::EnsureDirect2D(HWND hwnd) noexcept
{
    RECT rc{};
    GetClientRect(hwnd, &rc);
    const UINT32 width  = static_cast<UINT32>(std::max<LONG>(1, rc.right - rc.left));
    const UINT32 height = static_cast<UINT32>(std::max<LONG>(1, rc.bottom - rc.top));

    if (! _d2dFactory || ! _dwriteFactory)
    {
        const auto factoriesStartedAt = std::chrono::steady_clock::now();
        HRESULT factoriesHr           = S_OK;
        auto emitFactories            = wil::scope_exit([&]() noexcept
        { Debug::Perf::EmitDurationUs(L"viewer.space.render.factories_us", Debug::Perf::ElapsedUs(factoriesStartedAt), 0u, 0u, factoriesHr); });

        if (! _d2dFactory)
        {
            D2D1_FACTORY_OPTIONS options{};
            factoriesHr =
                D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, __uuidof(ID2D1Factory1), &options, reinterpret_cast<void**>(_d2dFactory.addressof()));
            if (FAILED(factoriesHr))
            {
                Debug::Warning(L"ViewerSpace: Failed to create D2D factory: 0x{:08X}", static_cast<unsigned long>(factoriesHr));
                return false;
            }
        }

        if (! _dwriteFactory)
        {
            factoriesHr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.put()));
            if (FAILED(factoriesHr))
            {
                return false;
            }
        }
    }

    if (! _otherStrokeStyle && _d2dFactory)
    {
        D2D1_STROKE_STYLE_PROPERTIES props = D2D1::StrokeStyleProperties();
        props.startCap                     = D2D1_CAP_STYLE_ROUND;
        props.endCap                       = D2D1_CAP_STYLE_ROUND;
        props.dashCap                      = D2D1_CAP_STYLE_ROUND;
        props.lineJoin                     = D2D1_LINE_JOIN_ROUND;
        props.dashStyle                    = D2D1_DASH_STYLE_DASH;
        if (FAILED(_d2dFactory->CreateStrokeStyle(props, nullptr, 0, _otherStrokeStyle.put())))
        {
            Debug::Warning(L"ViewerSpace: Failed to create stroke style");
        }
    }

    if (! _dogEarFlapGeometry && _d2dFactory)
    {
        wil::com_ptr<ID2D1PathGeometry> geom;
        if (SUCCEEDED(_d2dFactory->CreatePathGeometry(geom.put())) && geom)
        {
            wil::com_ptr<ID2D1GeometrySink> sink;
            if (SUCCEEDED(geom->Open(sink.put())) && sink)
            {
                sink->BeginFigure(D2D1::Point2F(0.0f, 0.0f), D2D1_FIGURE_BEGIN_FILLED);
                const D2D1_POINT_2F points[] = {D2D1::Point2F(0.0f, 1.0f), D2D1::Point2F(1.0f, 1.0f)};
                sink->AddLines(points, static_cast<UINT32>(std::size(points)));
                sink->EndFigure(D2D1_FIGURE_END_CLOSED);

                if (SUCCEEDED(sink->Close()))
                {
                    _dogEarFlapGeometry = std::move(geom);
                }
            }
        }
    }

    if (_renderTarget && _renderer.swapChainWidthPx == width && _renderer.swapChainHeightPx == height)
    {
        return true;
    }
#ifdef ENABLE_TESTS
    auto rememberRendererFailure = [&](uint32_t stage, HRESULT hr) noexcept
    {
        _rendererFailureStage = stage;
        _rendererFailureHr    = static_cast<uint32_t>(hr);
    };
    auto clearRendererFailure = [&]() noexcept
    {
        _rendererFailureStage = 0u;
        _rendererFailureHr    = 0u;
    };
#endif

    if (! _renderer.d3dDevice || ! _renderer.d3dContext)
    {
        UINT creationFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
#if defined(_DEBUG)
        creationFlags |= D3D11_CREATE_DEVICE_DEBUG;
#endif
        static constexpr D3D_FEATURE_LEVEL kFeatureLevels[] = {D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL_10_1};
        // The SDK debug layer is an optional Windows component (absent on native ARM64 CI runners): DxUi retries without it
        // instead of failing the device, and falls back from hardware to WARP.
        D3D_DRIVER_TYPE driverType    = D3D_DRIVER_TYPE_UNKNOWN;
        const auto d3dDeviceStartedAt = std::chrono::steady_clock::now();
        const HRESULT d3dHr           = DxUi::CreateD3D11DeviceWithWarpFallback(
            creationFlags, kFeatureLevels, false, _renderer.d3dDevice.put(), &_renderer.featureLevel, _renderer.d3dContext.put(), &driverType);
        Debug::Perf::EmitDurationUs(L"viewer.space.render.d3d_device_us", Debug::Perf::ElapsedUs(d3dDeviceStartedAt), 0u, 0u, d3dHr);
        if (FAILED(d3dHr) || ! _renderer.d3dDevice || ! _renderer.d3dContext)
        {
            Debug::Warning(L"ViewerSpace: D3D11CreateDevice failed: 0x{:08X}", static_cast<unsigned long>(d3dHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureD3DDevice, d3dHr);
#endif
            return false;
        }
        if (driverType == D3D_DRIVER_TYPE_WARP)
        {
            Debug::Warning(L"ViewerSpace: Hardware D3D11 device unavailable; rendering through WARP.");
        }

        wil::com_ptr<IDXGIDevice> dxgiDevice;
        wil::com_ptr<IDXGIAdapter> adapter;
        const auto dxgiFactoryStartedAt  = std::chrono::steady_clock::now();
        const HRESULT dxgiHr             = _renderer.d3dDevice->QueryInterface(IID_PPV_ARGS(dxgiDevice.addressof()));
        const HRESULT adapterHr          = SUCCEEDED(dxgiHr) && dxgiDevice ? dxgiDevice->GetAdapter(adapter.addressof()) : E_FAIL;
        const HRESULT factoryHr          = SUCCEEDED(adapterHr) && adapter ? adapter->GetParent(IID_PPV_ARGS(_renderer.dxgiFactory.addressof())) : E_FAIL;
        const HRESULT dxgiFactoryStageHr = FAILED(dxgiHr) ? dxgiHr : FAILED(adapterHr) ? adapterHr : factoryHr;
        Debug::Perf::EmitDurationUs(L"viewer.space.render.dxgi_factory_us", Debug::Perf::ElapsedUs(dxgiFactoryStartedAt), 0u, 0u, dxgiFactoryStageHr);
        const auto d2dContextStartedAt = std::chrono::steady_clock::now();
        const HRESULT d2dDeviceHr  = SUCCEEDED(dxgiHr) && dxgiDevice ? _d2dFactory->CreateDevice(dxgiDevice.get(), _renderer.d2dDevice.addressof()) : E_FAIL;
        const HRESULT d2dContextHr = SUCCEEDED(d2dDeviceHr) && _renderer.d2dDevice
                                         ? _renderer.d2dDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE, _renderer.d2dContext.addressof())
                                         : E_FAIL;
        const HRESULT d2dContextStageHr = FAILED(d2dDeviceHr) ? d2dDeviceHr : d2dContextHr;
        Debug::Perf::EmitDurationUs(L"viewer.space.render.d2d_context_us", Debug::Perf::ElapsedUs(d2dContextStartedAt), 0u, 0u, d2dContextStageHr);
        if (FAILED(dxgiHr) || ! dxgiDevice || FAILED(adapterHr) || ! adapter || FAILED(factoryHr) || ! _renderer.dxgiFactory || FAILED(d2dDeviceHr) ||
            ! _renderer.d2dDevice || FAILED(d2dContextHr) || ! _renderer.d2dContext)
        {
            Debug::Warning(L"ViewerSpace: Device-context renderer creation failed ({:08X}, {:08X}, {:08X}, {:08X}, {:08X})",
                           static_cast<unsigned long>(dxgiHr),
                           static_cast<unsigned long>(adapterHr),
                           static_cast<unsigned long>(factoryHr),
                           static_cast<unsigned long>(d2dDeviceHr),
                           static_cast<unsigned long>(d2dContextHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureD2DContext, FAILED(d2dContextHr) ? d2dContextHr : E_FAIL);
#endif
            DiscardDirect2D();
            return false;
        }

        _renderer.d2dContext->SetUnitMode(D2D1_UNIT_MODE_DIPS);
#ifdef ENABLE_TESTS
        _rendererDeviceCreateCount += 1u;
#endif
    }

    if (! _renderer.swapChain)
    {
        DXGI_SWAP_CHAIN_DESC1 desc{};
        desc.Width            = width;
        desc.Height           = height;
        desc.Format           = DXGI_FORMAT_B8G8R8A8_UNORM;
        desc.SampleDesc.Count = 1;
        desc.BufferUsage      = DXGI_USAGE_RENDER_TARGET_OUTPUT;
        desc.BufferCount      = 2;
        desc.Scaling          = DXGI_SCALING_STRETCH;
        desc.SwapEffect       = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
        desc.AlphaMode        = DXGI_ALPHA_MODE_IGNORE;

        const auto swapChainCreateStartedAt = std::chrono::steady_clock::now();
        const HRESULT swapHr =
            _renderer.dxgiFactory
                ? _renderer.dxgiFactory->CreateSwapChainForHwnd(_renderer.d3dDevice.get(), hwnd, &desc, nullptr, nullptr, _renderer.swapChain.addressof())
                : E_FAIL;
        Debug::Perf::EmitDurationUs(L"viewer.space.render.swapchain_create_us", Debug::Perf::ElapsedUs(swapChainCreateStartedAt), 0u, 0u, swapHr);
        if (FAILED(swapHr) || ! _renderer.swapChain)
        {
            Debug::Warning(L"ViewerSpace: CreateSwapChainForHwnd failed: 0x{:08X}", static_cast<unsigned long>(swapHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureSwapChain, swapHr);
#endif
            DiscardDirect2D();
            return false;
        }
        _renderer.swapChainWidthPx  = width;
        _renderer.swapChainHeightPx = height;
        static_cast<void>(_renderer.dxgiFactory->MakeWindowAssociation(hwnd, DXGI_MWA_NO_ALT_ENTER | DXGI_MWA_NO_WINDOW_CHANGES));
    }

    const bool sizeChanged = _renderer.swapChainWidthPx != width || _renderer.swapChainHeightPx != height;
    if (sizeChanged)
    {
        InvalidateStaticTreemapCache();
        if (_renderer.d2dContext)
        {
            _renderer.d2dContext->SetTarget(nullptr);
            static_cast<void>(_renderer.d2dContext->Flush());
        }
        _renderTarget.reset();
        _renderer.targetBitmap.reset();
        _headerPathSourceText.clear();
        _headerPathDisplayText.clear();
        _headerPathDisplayMaxWidthDip = 0.0f;
        if (_renderer.d3dContext)
        {
            _renderer.d3dContext->ClearState();
            _renderer.d3dContext->Flush();
        }

        const HRESULT resizeHr = _renderer.swapChain->ResizeBuffers(0u, width, height, DXGI_FORMAT_UNKNOWN, 0u);
        if (FAILED(resizeHr))
        {
            Debug::Warning(L"ViewerSpace: ResizeBuffers failed: 0x{:08X}", static_cast<unsigned long>(resizeHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureResizeBuffers, resizeHr);
#endif
            DiscardDirect2D();
            return false;
        }
        _renderer.swapChainWidthPx  = width;
        _renderer.swapChainHeightPx = height;
#ifdef ENABLE_TESTS
        _swapChainResizeCount += 1u;
#endif
    }

    if (! _renderer.targetBitmap)
    {
        const auto targetBitmapStartedAt = std::chrono::steady_clock::now();
        HRESULT targetBitmapHr           = S_OK;
        auto emitTargetBitmap            = wil::scope_exit([&]() noexcept
        { Debug::Perf::EmitDurationUs(L"viewer.space.render.target_bitmap_us", Debug::Perf::ElapsedUs(targetBitmapStartedAt), 0u, 0u, targetBitmapHr); });
        wil::com_ptr<IDXGISurface> surface;
        const HRESULT bufferHr = _renderer.swapChain ? _renderer.swapChain->GetBuffer(0u, IID_PPV_ARGS(surface.addressof())) : E_FAIL;
        targetBitmapHr         = bufferHr;
        if (FAILED(bufferHr) || ! surface)
        {
            Debug::Warning(L"ViewerSpace: IDXGISwapChain1::GetBuffer failed: 0x{:08X}", static_cast<unsigned long>(bufferHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureGetBuffer, bufferHr);
#endif
            DiscardDirect2D();
            return false;
        }

        const D2D1_BITMAP_PROPERTIES1 props = D2D1::BitmapProperties1(
            D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE), _dpi, _dpi);
        const HRESULT bitmapHr = _renderer.d2dContext->CreateBitmapFromDxgiSurface(surface.get(), &props, _renderer.targetBitmap.addressof());
        targetBitmapHr         = bitmapHr;
        if (FAILED(bitmapHr) || ! _renderer.targetBitmap)
        {
            Debug::Warning(L"ViewerSpace: CreateBitmapFromDxgiSurface failed: 0x{:08X}", static_cast<unsigned long>(bitmapHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureTargetBitmap, bitmapHr);
#endif
            DiscardDirect2D();
            return false;
        }
        _renderer.d2dContext->SetTarget(_renderer.targetBitmap.get());
    }

    _renderer.d2dContext->SetDpi(_dpi, _dpi);
    _renderer.d2dContext->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    const HRESULT renderTargetHr = _renderTarget ? S_OK : _renderer.d2dContext->QueryInterface(IID_PPV_ARGS(_renderTarget.addressof()));
    if (FAILED(renderTargetHr) || ! _renderTarget)
    {
        Debug::Warning(L"ViewerSpace: Device context did not expose ID2D1RenderTarget.");
#ifdef ENABLE_TESTS
        rememberRendererFailure(kRendererFailureRenderTarget, renderTargetHr);
#endif
        DiscardDirect2D();
        return false;
    }

    const D2D1::ColorF bg  = _hasTheme ? ColorFromArgb(_theme.backgroundArgb) : D2D1::ColorF(D2D1::ColorF::White);
    const D2D1::ColorF txt = _hasTheme ? ColorFromArgb(_theme.textArgb) : D2D1::ColorF(D2D1::ColorF::Black);
    const D2D1::ColorF acc = _hasTheme ? ColorFromArgb(_theme.accentArgb) : D2D1::ColorF(D2D1::ColorF::CornflowerBlue);

    const bool resourcesNeedCreation = ! _brushBackground || ! _brushText || ! _brushOutline || ! _brushAccent || ! _brushWatermark || ! _shadingStops ||
                                       ! _brushShading || ! _textFormat || ! _headerFormat || ! _headerStatusFormatRight || ! _headerInfoFormat ||
                                       ! _headerInfoFormatRight || ! _headerIconFormat || ! _watermarkFormat || ! _tooltipFormat;
    const auto resourcesStartedAt    = resourcesNeedCreation ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{};
    HRESULT resourcesHr              = S_OK;
    auto emitResources               = wil::scope_exit([&]() noexcept
    {
        if (resourcesNeedCreation)
        {
            Debug::Perf::EmitDurationUs(L"viewer.space.render.resources_us", Debug::Perf::ElapsedUs(resourcesStartedAt), 0u, 0u, resourcesHr);
        }
    });

    if (! _brushBackground)
    {
        const HRESULT brushHr = _renderTarget->CreateSolidColorBrush(bg, _brushBackground.put());
        if (FAILED(brushHr))
        {
            resourcesHr = brushHr;
            Debug::Warning(L"ViewerSpace: Failed to create background brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }
    if (! _brushText)
    {
        const HRESULT brushHr = _renderTarget->CreateSolidColorBrush(txt, _brushText.put());
        if (FAILED(brushHr))
        {
            resourcesHr = brushHr;
            Debug::Warning(L"ViewerSpace: Failed to create text brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }
    if (! _brushOutline)
    {
        const HRESULT brushHr = _renderTarget->CreateSolidColorBrush(txt, _brushOutline.put());
        if (FAILED(brushHr))
        {
            resourcesHr = brushHr;
            Debug::Warning(L"ViewerSpace: Failed to create outline brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }
    if (! _brushAccent)
    {
        const HRESULT brushHr = _renderTarget->CreateSolidColorBrush(acc, _brushAccent.put());
        if (FAILED(brushHr))
        {
            resourcesHr = brushHr;
            Debug::Warning(L"ViewerSpace: Failed to create accent brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }
    if (! _brushWatermark)
    {
        const HRESULT brushHr = _renderTarget->CreateSolidColorBrush(txt, _brushWatermark.put());
        if (FAILED(brushHr))
        {
            resourcesHr = brushHr;
            Debug::Warning(L"ViewerSpace: Failed to create watermark brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }

    const bool darkMode     = _hasTheme && _theme.darkMode != FALSE;
    const bool highContrast = _hasTheme && _theme.highContrast != FALSE;
    const bool rainbowMode  = _hasTheme && _theme.rainbowMode != FALSE;

    float highlightAlpha = darkMode ? 0.14f : 0.08f;
    float shadowAlpha    = darkMode ? 0.22f : 0.14f;
    if (rainbowMode)
    {
        highlightAlpha *= 0.45f;
        shadowAlpha *= 0.45f;
    }
    if (highContrast)
    {
        highlightAlpha = 0.0f;
        shadowAlpha    = 0.0f;
    }

    std::array<D2D1_GRADIENT_STOP, 3> stops{};
    stops[0].position = 0.0f;
    stops[0].color    = D2D1::ColorF(1.0f, 1.0f, 1.0f, highlightAlpha);
    stops[1].position = 0.55f;
    stops[1].color    = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.0f);
    stops[2].position = 1.0f;
    stops[2].color    = D2D1::ColorF(0.0f, 0.0f, 0.0f, shadowAlpha);

    if (! _shadingStops)
    {
        const HRESULT shadingStopsHr = _renderTarget->CreateGradientStopCollection(stops.data(), static_cast<UINT32>(stops.size()), _shadingStops.put());
        if (FAILED(shadingStopsHr))
        {
            resourcesHr = shadingStopsHr;
            Debug::Warning(L"ViewerSpace: Failed to create gradient stop collection");
            return false;
        }
    }
    if (! _brushShading)
    {
        const HRESULT shadingBrushHr = _renderTarget->CreateLinearGradientBrush(
            D2D1::LinearGradientBrushProperties(D2D1::Point2F(0, 0), D2D1::Point2F(1, 1)), _shadingStops.get(), _brushShading.put());
        if (FAILED(shadingBrushHr))
        {
            resourcesHr = shadingBrushHr;
            Debug::Warning(L"ViewerSpace: Failed to create shading brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }

    HRESULT textHr = S_OK;
    if (! _textFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(12.0f), _textFormat.put());
        if (FAILED(textHr))
        {
            resourcesHr = textHr;
        }
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _textFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _textFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
        }
    }

    if (! _headerFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(13.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _headerFormat.put());
        if (FAILED(textHr))
        {
            resourcesHr = textHr;
        }
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        }
    }

    if (! _headerStatusFormatRight)
    {
        textHr =
            Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(13.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _headerStatusFormatRight.put());
        if (FAILED(textHr))
        {
            resourcesHr = textHr;
        }
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerStatusFormatRight->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerStatusFormatRight->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerStatusFormatRight->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
        }
    }

    if (! _headerInfoFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(11.0f), _headerInfoFormat.put());
        if (FAILED(textHr))
        {
            resourcesHr = textHr;
        }
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerInfoFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerInfoFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerInfoFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        }
    }

    if (! _headerInfoFormatRight)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(11.0f), _headerInfoFormatRight.put());
        if (FAILED(textHr))
        {
            resourcesHr = textHr;
        }
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerInfoFormatRight->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerInfoFormatRight->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerInfoFormatRight->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
        }
    }

    if (! _headerIconFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(18.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _headerIconFormat.put());
        if (FAILED(textHr))
        {
            resourcesHr = textHr;
        }
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerIconFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerIconFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerIconFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        }
    }

    if (! _watermarkFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(18.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _watermarkFormat.put());
        if (FAILED(textHr))
        {
            resourcesHr = textHr;
        }
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _watermarkFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _watermarkFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _watermarkFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        }
    }

    if (! _tooltipFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(12.0f), _tooltipFormat.put());
        if (FAILED(textHr))
        {
            resourcesHr = textHr;
        }
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _tooltipFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
            _tooltipFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
            _tooltipFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        }
    }

    auto applyEllipsisTrimming = [&](IDWriteTextFormat* format) noexcept
    {
        if (format == nullptr || ! _dwriteFactory)
        {
            return;
        }

        DWRITE_TRIMMING trimming{};
        trimming.granularity = DWRITE_TRIMMING_GRANULARITY_CHARACTER;

        wil::com_ptr<IDWriteInlineObject> ellipsis;
        if (SUCCEEDED(_dwriteFactory->CreateEllipsisTrimmingSign(format, ellipsis.put())))
        {
            format->SetTrimming(&trimming, ellipsis.get());
        }
    };

    applyEllipsisTrimming(_textFormat.get());
    applyEllipsisTrimming(_headerFormat.get());
    applyEllipsisTrimming(_headerStatusFormatRight.get());
    applyEllipsisTrimming(_headerInfoFormat.get());
    applyEllipsisTrimming(_headerInfoFormatRight.get());
    applyEllipsisTrimming(_watermarkFormat.get());

#ifdef ENABLE_TESTS
    clearRendererFailure();
#endif
    return true;
}

void ViewerSpace::DiscardDirect2D() noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    std::array<uint64_t, 8u> stages{};
#ifdef ENABLE_TESTS
    if (_rendererFailureStage == 0u)
    {
        _rendererFailureStage = kRendererDiscardUnknown;
        _rendererFailureHr    = 0u;
    }
#endif
    InvalidateStaticTreemapCache();
    if (_renderer.d2dContext)
    {
        _renderer.d2dContext->SetTarget(nullptr);
    }
    _renderTarget.reset();
    _renderer.targetBitmap.reset();
    stages[0] = Debug::Perf::ElapsedUs(startedAt);
    _renderer.swapChain.reset();
    stages[1] = Debug::Perf::ElapsedUs(startedAt);
    _renderer.d2dContext.reset();
    _renderer.d2dDevice.reset();
    stages[2] = Debug::Perf::ElapsedUs(startedAt);
    _renderer.dxgiFactory.reset();
    stages[3] = Debug::Perf::ElapsedUs(startedAt);
    _renderer.d3dContext.reset();
    stages[4] = Debug::Perf::ElapsedUs(startedAt);
    _renderer.d3dDevice.reset();
    stages[5]                   = Debug::Perf::ElapsedUs(startedAt);
    _renderer.swapChainWidthPx  = 0;
    _renderer.swapChainHeightPx = 0;
    _brushBackground.reset();
    _brushText.reset();
    _brushOutline.reset();
    _brushAccent.reset();
    _brushWatermark.reset();
    _brushShading.reset();
    _shadingStops.reset();
    stages[6] = Debug::Perf::ElapsedUs(startedAt);
    _textFormat.reset();
    _headerFormat.reset();
    _headerStatusFormatRight.reset();
    _headerInfoFormat.reset();
    _headerInfoFormatRight.reset();
    _headerIconFormat.reset();
    _watermarkFormat.reset();
    _tooltipFormat.reset();
    _headerPathSourceText.clear();
    _headerPathDisplayText.clear();
    _headerPathDisplayMaxWidthDip = 0.0f;
    stages[7]                     = Debug::Perf::ElapsedUs(startedAt);
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"viewer.space.render.discard_stages_us",
                          std::format(L"target={};swapChain={};d2d={};dxgiFactory={};d3dContext={};d3dDevice={};brushes={};text={}",
                                      stages[0],
                                      stages[1] - stages[0],
                                      stages[2] - stages[1],
                                      stages[3] - stages[2],
                                      stages[4] - stages[3],
                                      stages[5] - stages[4],
                                      stages[6] - stages[5],
                                      stages[7] - stages[6]),
                          stages[7],
                          0u,
                          0u,
                          S_OK);
    }
}

void ViewerSpace::InvalidateStaticTreemapCache() noexcept
{
    _staticTreemapCache.bitmap.reset();
    _staticTreemapCache.widthPx  = 0u;
    _staticTreemapCache.heightPx = 0u;
    _staticTreemapCache.bytes    = 0u;
    _staticTreemapCache.generation += 1u;
}

bool ViewerSpace::HasActiveLayoutAnimation(double nowSeconds) const noexcept
{
    static constexpr float kSettledRectToleranceDip = 0.25f;
    for (const DrawItem& item : _drawItems)
    {
        if ((nowSeconds - item.animationStartSeconds) < kAnimationDurationSeconds)
        {
            return true;
        }

        if (std::fabs(item.currentRect.left - item.targetRect.left) > kSettledRectToleranceDip ||
            std::fabs(item.currentRect.top - item.targetRect.top) > kSettledRectToleranceDip ||
            std::fabs(item.currentRect.right - item.targetRect.right) > kSettledRectToleranceDip ||
            std::fabs(item.currentRect.bottom - item.targetRect.bottom) > kSettledRectToleranceDip)
        {
            return true;
        }
    }

    return false;
}

bool ViewerSpace::CanReplayStaticTreemapCache(bool scanActive, bool completionOverlayActive) const noexcept
{
    if (! _staticTreemapCache.bitmap || ! _renderTarget)
    {
        return false;
    }

    if (_staticTreemapCache.widthPx != _renderer.swapChainWidthPx || _staticTreemapCache.heightPx != _renderer.swapChainHeightPx)
    {
        return false;
    }

    if (_layoutDirty || scanActive || completionOverlayActive || _hoverHeaderHit != HeaderHit::None)
    {
        return false;
    }

    return true;
}

bool ViewerSpace::CanRecordStaticTreemapCache(bool scanActive, bool completionOverlayActive, double nowSeconds) const noexcept
{
    if (! _renderTarget || _renderer.swapChainWidthPx == 0u || _renderer.swapChainHeightPx == 0u)
    {
        return false;
    }

    if (_layoutDirty || scanActive || completionOverlayActive || _hoverHeaderHit != HeaderHit::None || _hoverNodeId != 0 || _tooltipNodeId != 0)
    {
        return false;
    }

    return ! HasActiveLayoutAnimation(nowSeconds);
}

HRESULT ViewerSpace::EnsureStaticTreemapCacheBitmap() noexcept
{
    if (! _renderer.d2dContext)
    {
        return E_FAIL;
    }

    const UINT width  = _renderer.swapChainWidthPx;
    const UINT height = _renderer.swapChainHeightPx;
    if (width == 0u || height == 0u)
    {
        return E_INVALIDARG;
    }

    if (_staticTreemapCache.bitmap && _staticTreemapCache.widthPx == width && _staticTreemapCache.heightPx == height)
    {
        return S_OK;
    }

    _staticTreemapCache.bitmap.reset();

    D2D1_BITMAP_PROPERTIES1 props =
        D2D1::BitmapProperties1(D2D1_BITMAP_OPTIONS_NONE, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE), _dpi, _dpi);

    const HRESULT hr = _renderer.d2dContext->CreateBitmap(D2D1::SizeU(width, height), nullptr, 0u, &props, _staticTreemapCache.bitmap.put());
    if (FAILED(hr))
    {
        _staticTreemapCache.widthPx  = 0u;
        _staticTreemapCache.heightPx = 0u;
        _staticTreemapCache.bytes    = 0u;
        return hr;
    }

    _staticTreemapCache.widthPx  = width;
    _staticTreemapCache.heightPx = height;
    _staticTreemapCache.bytes    = static_cast<uint64_t>(width) * static_cast<uint64_t>(height) * 4u;
    return S_OK;
}

HRESULT ViewerSpace::RecordStaticTreemapCacheFromCurrentTarget() noexcept
{
    const auto recordStartedAt = std::chrono::steady_clock::now();
    HRESULT hr                 = EnsureStaticTreemapCacheBitmap();
    if (FAILED(hr))
    {
        return hr;
    }

    const D2D1_RECT_U sourceRect = D2D1::RectU(0u, 0u, _staticTreemapCache.widthPx, _staticTreemapCache.heightPx);
    hr                           = _staticTreemapCache.bitmap->CopyFromRenderTarget(nullptr, _renderTarget.get(), &sourceRect);
    if (FAILED(hr))
    {
        _staticTreemapCache.bitmap.reset();
        _staticTreemapCache.bytes = 0u;
        return hr;
    }

    _staticTreemapCache.recordCount += 1u;
    _staticTreemapCache.lastRecordUs = Debug::Perf::ElapsedUs(recordStartedAt);
    Debug::Perf::Emit(L"viewer.space.render.static_cache_record_us",
                      L"bitmap",
                      _staticTreemapCache.lastRecordUs,
                      static_cast<uint64_t>(_drawItems.size()),
                      _staticTreemapCache.bytes,
                      S_OK);
    return S_OK;
}

void ViewerSpace::PaintHoverOverlay(const D2D1::ColorF& accentColor) noexcept
{
    if ((_hoverNodeId == 0 && _selectedItemId == 0 && _analysisFilter.empty() && ! _typeFilter && ! _ageFilter) || ! _brushOutline || ! _renderTarget)
    {
        return;
    }

    for (const auto& item : _drawItems)
    {
        const bool selected = item.nodeId == _selectedItemId && _selectedGeneration == _displayGeneration;
        // Filtering highlights retained names in the map without changing geometry, totals, or cached pixels.
        const bool filterMatch = AnalysisFilterMatches(item.nodeId);
        if (item.nodeId != _hoverNodeId && ! selected && ! filterMatch)
        {
            continue;
        }

        float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
        gap       = std::clamp(gap, 0.5f, kItemGapDip);

        D2D1_RECT_F rc = item.currentRect;
        rc.left += gap;
        rc.top += gap;
        rc.right -= gap;
        rc.bottom -= gap;
        if (RectArea(rc) < 1.0f)
        {
            continue;
        }

        _brushOutline->SetColor(D2D1::ColorF(accentColor.r, accentColor.g, accentColor.b, selected ? 1.0f : filterMatch ? 0.55f : 0.9f));
        const float stroke                     = selected ? 3.25f : filterMatch ? 1.25f : 2.25f;
        const std::optional<ResolvedItem> node = ResolveTreemapItem(item.nodeId);
        const bool isFileTile                  = node.has_value() && ! node->isDirectory && ! node->isSynthetic;
        if (isFileTile)
        {
            const float tileW   = std::max(0.0f, rc.right - rc.left);
            const float tileH   = std::max(0.0f, rc.bottom - rc.top);
            const float side    = std::min(tileW, tileH);
            const float dogSize = std::clamp(side * 0.18f, 8.0f, 14.0f);
            float cut           = dogSize + 2.0f;
            cut                 = std::clamp(cut, 6.0f, std::max(6.0f, side - 1.0f));

            if (cut > 1.0f && tileW > cut + 1.0f && tileH > cut + 1.0f)
            {
                const D2D1_POINT_2F tl = D2D1::Point2F(rc.left, rc.top);
                const D2D1_POINT_2F br = D2D1::Point2F(rc.right, rc.bottom);
                const D2D1_POINT_2F bl = D2D1::Point2F(rc.left, rc.bottom);

                const D2D1_POINT_2F cutA = D2D1::Point2F(rc.right - cut, rc.top);
                const D2D1_POINT_2F cutB = D2D1::Point2F(rc.right, rc.top + cut);

                _renderTarget->DrawLine(tl, cutA, _brushOutline.get(), stroke, nullptr);
                _renderTarget->DrawLine(cutA, cutB, _brushOutline.get(), stroke, nullptr);
                _renderTarget->DrawLine(cutB, br, _brushOutline.get(), stroke, nullptr);
                _renderTarget->DrawLine(br, bl, _brushOutline.get(), stroke, nullptr);
                _renderTarget->DrawLine(bl, tl, _brushOutline.get(), stroke, nullptr);
            }
            else
            {
                _renderTarget->DrawRectangle(rc, _brushOutline.get(), stroke);
            }
        }
        else
        {
            _renderTarget->DrawRectangle(rc, _brushOutline.get(), stroke);
        }
    }
}

[[nodiscard]] bool TryReserveAtomicBudget(std::atomic_uint64_t& currentBytes, uint64_t requestedBytes, uint64_t byteLimit, uint64_t& admittedBytes) noexcept
{
    uint64_t observed = currentBytes.load(std::memory_order_relaxed);
    for (;;)
    {
        if (observed > byteLimit || requestedBytes > byteLimit - observed)
        {
            return false;
        }
        const uint64_t desired = observed + requestedBytes;
        if (currentBytes.compare_exchange_weak(observed, desired, std::memory_order_acq_rel, std::memory_order_relaxed))
        {
            admittedBytes = desired;
            return true;
        }
    }
}

[[nodiscard]] uint64_t ReleaseAtomicBudget(std::atomic_uint64_t& currentBytes, uint64_t releasedBytes) noexcept
{
    uint64_t observed = currentBytes.load(std::memory_order_relaxed);
    for (;;)
    {
        const uint64_t desired = observed >= releasedBytes ? observed - releasedBytes : 0u;
        if (currentBytes.compare_exchange_weak(observed, desired, std::memory_order_acq_rel, std::memory_order_relaxed))
        {
            return desired;
        }
    }
}

void UpdateAtomicPeak(std::atomic_uint64_t& peak, uint64_t value) noexcept
{
    uint64_t observed = peak.load(std::memory_order_relaxed);
    while (value > observed && ! peak.compare_exchange_weak(observed, value, std::memory_order_relaxed, std::memory_order_relaxed))
    {
    }
}

void ViewerSpace::ApplyThemeToWindow(HWND hwnd) noexcept
{
    if (! _hasTheme)
    {
        return;
    }

    const bool dark = _theme.darkMode != FALSE;
    MessageBoxCenteringDetail::ApplyImmersiveDarkMode(hwnd, dark);

    const bool windowActive = GetActiveWindow() == hwnd;
    ApplyTitleBarTheme(hwnd, windowActive);
    ApplyMenuTheme(hwnd);
    _analysisHost.ApplyTheme(_hasTheme ? RedSalamander::MakeThemePaletteFromViewerTheme(_theme) : DxUi::MakeDefaultThemePalette(false));
    _navigationHost.ApplyTheme(_hasTheme ? RedSalamander::MakeThemePaletteFromViewerTheme(_theme) : DxUi::MakeDefaultThemePalette(false));
}

void ViewerSpace::ApplyTitleBarTheme(HWND hwnd, bool windowActive) noexcept
{
    if (! _hasTheme || ! hwnd)
    {
        return;
    }

    using DwmSetWindowAttributeFunc          = HRESULT(WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
    static DwmSetWindowAttributeFunc setAttr = []() noexcept -> DwmSetWindowAttributeFunc
    {
        HMODULE dwm = LoadLibraryW(L"dwmapi.dll");
        if (! dwm)
        {
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // 'reinterpret_cast': unsafe conversion from 'FARPROC'
        return reinterpret_cast<DwmSetWindowAttributeFunc>(GetProcAddress(dwm, "DwmSetWindowAttribute"));
#pragma warning(pop)
    }();

    if (! setAttr)
    {
        return;
    }

    static constexpr DWORD kDwmwaBorderColor  = 34u;
    static constexpr DWORD kDwmwaCaptionColor = 35u;
    static constexpr DWORD kDwmwaTextColor    = 36u;
    static constexpr DWORD kDwmColorDefault   = 0xFFFFFFFFu;

    DWORD borderValue  = kDwmColorDefault;
    DWORD captionValue = kDwmColorDefault;
    DWORD textValue    = kDwmColorDefault;

    if (_theme.highContrast == FALSE && _theme.rainbowMode != FALSE)
    {
        COLORREF accent = ColorRefFromArgb(_theme.accentArgb);
        if (! windowActive)
        {
            static constexpr uint8_t kInactiveTitleBlendAlpha = 223u; // ~7/8 toward background
            const COLORREF bg                                 = ColorRefFromArgb(_theme.backgroundArgb);
            accent                                            = BlendColorRefTruncate(accent, bg, kInactiveTitleBlendAlpha);
        }

        const COLORREF text = ChooseContrastingTextColor(accent);
        borderValue         = static_cast<DWORD>(accent);
        captionValue        = static_cast<DWORD>(accent);
        textValue           = static_cast<DWORD>(text);
    }

    setAttr(hwnd, kDwmwaBorderColor, &borderValue, sizeof(borderValue));
    setAttr(hwnd, kDwmwaCaptionColor, &captionValue, sizeof(captionValue));
    setAttr(hwnd, kDwmwaTextColor, &textValue, sizeof(textValue));
}

void ViewerSpace::UpdateWindowTitle(HWND hwnd) noexcept
{
    std::wstring pathText = _viewPathText;

    std::wstring title;
    if (! pathText.empty())
    {
        title = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TITLE_FORMAT, pathText);
        if (title.empty())
        {
            title = pathText;
        }
    }
    else
    {
        title = _metaName;
    }

    if (! title.empty())
    {
        SetWindowTextW(hwnd, title.c_str());
    }
}

void ViewerSpace::ApplyMenuTheme(HWND hwnd) noexcept
{
    static_cast<void>(hwnd);
    _menuBarHost.SetTheme(_hasTheme ? RedSalamander::MakeThemePaletteFromViewerTheme(_theme) : DxUi::MakeDefaultThemePalette(false));
    if (_menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
    else if (hwnd)
    {
        DrawMenuBar(hwnd);
    }
}

void ViewerSpace::UpdateHeaderTextCache() noexcept
{
    bool visibleHeaderChanged = false;
    const auto assignChanged  = [&visibleHeaderChanged](std::wstring& target, std::wstring value) noexcept
    {
        if (target != value)
        {
            target               = std::move(value);
            visibleHeaderChanged = true;
        }
    };
    const auto clearChanged = [&visibleHeaderChanged](std::wstring& target) noexcept
    {
        if (! target.empty())
        {
            target.clear();
            visibleHeaderChanged = true;
        }
    };

    UINT statusId = IDS_VIEWERSPACE_STATUS_DONE;
    switch (_overallState)
    {
        case ScanState::NotStarted: statusId = IDS_VIEWERSPACE_STATUS_NOT_STARTED; break;
        case ScanState::Queued: statusId = IDS_VIEWERSPACE_STATUS_QUEUED; break;
        case ScanState::Scanning: statusId = IDS_VIEWERSPACE_STATUS_SCANNING; break;
        case ScanState::Done: statusId = IDS_VIEWERSPACE_STATUS_DONE; break;
        case ScanState::Error: statusId = IDS_VIEWERSPACE_STATUS_ERROR; break;
        case ScanState::Canceled: statusId = IDS_VIEWERSPACE_STATUS_CANCELED; break;
    }

    // A completed scan with unreadable folders says so instead of plain "Done"; the failed folders keep their own Error tiles.
    if (_deferredStartScan.has_value() || _workingModelRetirementPending)
    {
        statusId = IDS_VIEWERSPACE_WAITING_FOR_PREVIOUS_VIEW;
    }
    else if (_refreshFailed)
    {
        statusId = IDS_VIEWERSPACE_REFRESH_FAILED;
    }
    else if (_refreshRefused)
    {
        statusId = IDS_VIEWERSPACE_REBUILD_REQUIRED;
    }
    else if (_refreshPending)
    {
        statusId = IDS_VIEWERSPACE_ANALYSIS_STATE_REFRESHING;
    }
    const uint64_t statusFailedFolders = _scanFailedFolders;
    if (_headerStatusId != statusId || _headerStatusFailedFolders != statusFailedFolders)
    {
        _headerStatusId            = statusId;
        _headerStatusFailedFolders = statusFailedFolders;
        assignChanged(_headerStatusText,
                      statusFailedFolders > 0u
                          ? FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_PROBLEMS, LoadStringResource(g_hInstance, statusId), statusFailedFolders)
                          : LoadStringResource(g_hInstance, statusId));
    }

    const bool scanActive = _overallState == ScanState::Queued || _overallState == ScanState::Scanning;

    const uint64_t items = static_cast<uint64_t>(_scanProgressFolders) + static_cast<uint64_t>(_scanProgressFiles);
    if (items > 0 || scanActive)
    {
        assignChanged(_headerCountsText,
                      FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_COUNTS_FORMAT, items, _scanProgressFolders, _scanProgressFiles));

        const CompactBytesText sizeText = FormatBytesCompactInline(_scanProgressBytes);
        const std::wstring_view sizeView(sizeText.buffer.data(), sizeText.length);

        assignChanged(_headerSizeText, FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_DISCOVERED, sizeView));
        const Node* viewed = _model ? TryGetRealNode(_model->viewNodeId) : nullptr;
        std::wstring viewText;
        if (viewed && ! _excludedPaths.empty())
            viewText = FormatStringResource(
                g_hInstance, IDS_VIEWERSPACE_EXCLUSION_SUMMARY, FormatBytesCompact(ShownBytes(viewed->id, viewed->totalBytes)), _excludedPaths.size());
        else if (viewed && viewed->id != _model->rootNodeId)
            viewText = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_SHOWN, FormatBytesCompact(viewed->totalBytes));
        if (_workspaceRebinding || _workspaceUnresolved)
            viewText.append(L" · ").append(
                FormatStringResource(g_hInstance, IDS_VIEWERSPACE_EXCLUSION_PENDING, _workspaceRebinding ? _excludedPaths.size() : _workspaceUnresolved));
        if (_config.fileCountWeight && viewed)
        {
            if (! viewText.empty())
                viewText.append(L" · ");
            viewText.append(LoadStringResource(g_hInstance, IDS_VIEWERSPACE_WEIGHT_COUNT))
                .append(L": ")
                .append(_workspaceFactsReady ? FormatStringResource(g_hInstance, IDS_VIEWERSPACE_FILE_COUNT_FORMAT, ShownFileCount(viewed->id))
                                             : LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ALLOCATION_UNAVAILABLE));
        }
        assignChanged(_headerViewText, std::move(viewText));
    }
    else
    {
        clearChanged(_headerCountsText);
        assignChanged(_headerSizeText, FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_DISCOVERED, FormatBytesCompact(0u)));
        clearChanged(_headerViewText);
    }

    if (scanActive && ! _scanProcessingFolderName.empty())
    {
        if (_workspaceScanStartedSeconds <= 0.0)
            _workspaceScanStartedSeconds = NowSeconds();
        std::wstring processing = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_PROCESSING_FORMAT, _scanProcessingFolderName);
        const double elapsed    = NowSeconds() - _workspaceScanStartedSeconds;
        if (elapsed >= 1.0 && items)
            processing.append(L" · ").append(
                FormatStringResource(g_hInstance, IDS_VIEWERSPACE_SCAN_RATE, static_cast<uint64_t>(static_cast<double>(items) / elapsed)));
        assignChanged(_headerProcessingText, std::move(processing));
    }
    else
    {
        clearChanged(_headerProcessingText);
        if (! scanActive)
            _workspaceScanStartedSeconds = 0.0;
    }

    // The static bitmap is a full-client capture, including these header strings.
    if (visibleHeaderChanged)
    {
        InvalidateStaticTreemapCache();
    }
}

bool ViewerSpace::CanAdmitWorkspaceBytes(uint64_t additional) const noexcept
{
    const uint64_t owned = ViewerOwnedBytes();
    return owned <= kViewerSpaceMemoryBudgetBytes && additional <= kViewerSpaceMemoryBudgetBytes - owned;
}

void ViewerSpace::UpdateTooltipForHit(uint32_t nodeId) noexcept
{
    if (! _hWnd)
    {
        _tooltipNodeId = 0;
        _tooltipText.clear();
        return;
    }

    if (nodeId == 0)
    {
        if (_tooltipNodeId != 0)
        {
            _tooltipNodeId = 0;
            _tooltipText.clear();
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        return;
    }

    if (nodeId != _tooltipNodeId)
    {
        _tooltipNodeId = nodeId;
        _tooltipText   = BuildTooltipText(nodeId);
        if (_tooltipText.empty())
        {
            _tooltipNodeId = 0;
        }
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::RefreshTooltipText() noexcept
{
    // The hovered item keeps growing while a scan runs; rebuild its text so the size and state shown are current.
    if (! _hWnd || _tooltipNodeId == 0)
    {
        return;
    }

    std::wstring text = BuildTooltipText(_tooltipNodeId);
    if (text.empty())
    {
        _tooltipNodeId = 0;
        _tooltipText.clear();
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
        return;
    }

    if (text != _tooltipText)
    {
        _tooltipText = std::move(text);
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::OnSettingChange(const wchar_t* area) noexcept
{
    // Regional-format changes arrive as WM_SETTINGCHANGE("intl"). The format-locale cache is per module, so the host's own
    // invalidation does not reach this plugin.
    const bool regionalChange = area == nullptr || CompareStringOrdinal(area, -1, L"intl", -1, TRUE) == CSTR_EQUAL;
    if (! regionalChange)
    {
        return;
    }

    LocaleFormatting::InvalidateFormatLocaleCache();
#ifdef ENABLE_TESTS
    _debugLocaleInvalidations += 1u;
#endif
    _headerStatusId = 0;
    _headerStatusText.clear();
    UpdateHeaderTextCache();
    RefreshTooltipText();
    InvalidateAnalysis();
    UpdateInspector();
    InvalidateStaticTreemapCache();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::UpdateTooltipPosition(int x, int y) noexcept
{
    if (! _hWnd || _tooltipNodeId == 0)
    {
        return;
    }

    const D2D1_POINT_2F newAnchor = D2D1::Point2F(DipFromPx(x) + 14.0f, DipFromPx(y) + 18.0f);
    if (std::abs(newAnchor.x - _tooltipAnchorDip.x) <= 0.5f && std::abs(newAnchor.y - _tooltipAnchorDip.y) <= 0.5f)
    {
        return;
    }

    _tooltipAnchorDip = newAnchor;
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void ViewerSpace::PaintTooltipOverlay() noexcept
{
#ifdef ENABLE_TESTS
    _lastTooltipPaintedBoundsDip = {};
#endif
    if (_tooltipNodeId == 0 || _tooltipText.empty() || ! _renderTarget || ! _dwriteFactory || ! _tooltipFormat || ! _brushBackground || ! _brushText)
    {
        return;
    }

    constexpr float kPadXDip          = 10.0f;
    constexpr float kPadYDip          = 7.0f;
    constexpr float kScreenGapDip     = 8.0f;
    constexpr float kMinTextWidthDip  = 48.0f;
    constexpr float kMinTextHeightDip = 16.0f;
    const D2D1_RECT_F mapBounds       = GetTreemapBounds();
    const float maxWidthDip           = std::min(kTooltipMaxWidthDip, mapBounds.right - mapBounds.left - 2.0f * (kScreenGapDip + kPadXDip));
    const float maxHeightDip          = std::min(kTooltipMaxHeightDip, mapBounds.bottom - mapBounds.top - 2.0f * (kScreenGapDip + kPadYDip));
    // This overlay is clipped by the canvas. Fit its text and padding to that actual
    // surface, and omit it when the map cannot hold a readable box.
    if (maxWidthDip < kMinTextWidthDip || maxHeightDip < kMinTextHeightDip)
    {
        return;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    const HRESULT layoutHr = _dwriteFactory->CreateTextLayout(
        _tooltipText.c_str(), static_cast<UINT32>(_tooltipText.size()), _tooltipFormat.get(), maxWidthDip, maxHeightDip, layout.put());
    if (FAILED(layoutHr) || ! layout)
    {
        return;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return;
    }

    const float textWidthDip  = std::clamp(metrics.widthIncludingTrailingWhitespace, kMinTextWidthDip, maxWidthDip);
    const float textHeightDip = std::clamp(metrics.height, kMinTextHeightDip, maxHeightDip);
    const float boxWidthDip   = textWidthDip + (kPadXDip * 2.0f);
    const float boxHeightDip  = textHeightDip + (kPadYDip * 2.0f);
    const float minLeft       = mapBounds.left + kScreenGapDip;
    const float minTop        = mapBounds.top + kScreenGapDip;
    const float maxLeft       = std::max(minLeft, mapBounds.right - boxWidthDip - kScreenGapDip);
    const float maxTop        = std::max(minTop, mapBounds.bottom - boxHeightDip - kScreenGapDip);

    float left = _tooltipAnchorDip.x;
    float top  = _tooltipAnchorDip.y;
    if (left + boxWidthDip > mapBounds.right - kScreenGapDip)
    {
        left = _tooltipAnchorDip.x - boxWidthDip - 24.0f;
    }
    if (top + boxHeightDip > mapBounds.bottom - kScreenGapDip)
    {
        top = _tooltipAnchorDip.y - boxHeightDip - 24.0f;
    }
    left = std::clamp(left, minLeft, maxLeft);
    top  = std::clamp(top, minTop, maxTop);

    const D2D1_RECT_F boxRc  = D2D1::RectF(left, top, left + boxWidthDip, top + boxHeightDip);
    const D2D1_POINT_2F text = D2D1::Point2F(boxRc.left + kPadXDip, boxRc.top + kPadYDip);

    const bool highContrast = _hasTheme && _theme.highContrast != FALSE;
    const bool darkMode     = _hasTheme && _theme.darkMode != FALSE;
    D2D1::ColorF bg         = _hasTheme ? ColorFromArgb(_theme.selectionBackgroundArgb)
                                        : D2D1::ColorF(GetRValue(GetSysColor(COLOR_INFOBK)) / 255.0f,
                                                       GetGValue(GetSysColor(COLOR_INFOBK)) / 255.0f,
                                                       GetBValue(GetSysColor(COLOR_INFOBK)) / 255.0f,
                                                       1.0f);
    D2D1::ColorF fg         = _hasTheme ? ColorFromArgb(_theme.selectionTextArgb)
                                        : D2D1::ColorF(GetRValue(GetSysColor(COLOR_INFOTEXT)) / 255.0f,
                                                       GetGValue(GetSysColor(COLOR_INFOTEXT)) / 255.0f,
                                                       GetBValue(GetSysColor(COLOR_INFOTEXT)) / 255.0f,
                                                       1.0f);

    bg.a = highContrast ? 1.0f : 0.96f;
    fg.a = 1.0f;

    const D2D1_COLOR_F oldBg   = _brushBackground->GetColor();
    const D2D1_COLOR_F oldText = _brushText->GetColor();
    _brushBackground->SetColor(bg);
    _renderTarget->FillRoundedRectangle(D2D1::RoundedRect(boxRc, 4.0f, 4.0f), _brushBackground.get());

    if (_brushOutline)
    {
        const D2D1_COLOR_F oldOutline = _brushOutline->GetColor();
        const float alpha             = highContrast ? 1.0f : (darkMode ? 0.45f : 0.30f);
        _brushOutline->SetColor(D2D1::ColorF(fg.r, fg.g, fg.b, alpha));
        _renderTarget->DrawRoundedRectangle(D2D1::RoundedRect(boxRc, 4.0f, 4.0f), _brushOutline.get(), 1.0f);
        _brushOutline->SetColor(oldOutline);
    }

    _brushText->SetColor(fg);
    _renderTarget->DrawTextLayout(text, layout.get(), _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
#ifdef ENABLE_TESTS
    _lastTooltipPaintedBoundsDip = boxRc;
    ++_tooltipPaintCount;
#endif
    _brushText->SetColor(oldText);
    _brushBackground->SetColor(oldBg);
}

std::wstring ViewerSpace::BuildTooltipText(uint32_t nodeId) const
{
    if (nodeId == 0)
    {
        return {};
    }

    const std::optional<ResolvedItem> resolved = ResolveTreemapItem(nodeId);
    if (! resolved.has_value())
    {
        return {};
    }
    const ResolvedItem& node = resolved.value();

    const auto withModified = [&node](std::wstring text)
    {
        std::wstring modified =
            Common::FileMetadata::FormatDisplayFields({node.lastWriteTime, 0u}, Common::FileMetadata::DisplayProfile::CompactDetails).localTime;
        if (modified.empty())
        {
            modified = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_STATE_UNAVAILABLE);
        }
        return FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_MODIFIED_FORMAT, text, modified);
    };

    const std::wstring pathText = BuildItemPathText(nodeId);

    std::wstring name(node.name.data(), node.name.size());
    if (name.empty())
    {
        name = pathText;
    }

    const std::wstring sizeText = FormatBytesCompact(node.totalBytes);

    uint64_t viewBytes = 0;
    const Node* view   = TryGetRealNode(_model->viewNodeId);
    if (view != nullptr)
    {
        viewBytes = MapWeight(view->id, view->totalBytes);
    }

    std::wstring shareText = LoadEmbeddedStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_SHARE_UNKNOWN);
    if (viewBytes > 0 && ! (_config.fileCountWeight && node.isSynthetic))
    {
        const double percent = std::clamp(static_cast<double>(MapWeight(node.id, node.totalBytes)) * 100.0 / static_cast<double>(viewBytes), 0.0, 100.0);
        shareText            = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_PERCENT_FORMAT, percent);
        if (shareText.empty())
        {
            shareText = LoadEmbeddedStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_SHARE_UNKNOWN);
        }
    }

    if (! pathText.empty())
    {
        if (node.contentNotMeasured || node.scanState != ScanState::Done)
        {
            std::wstring stateText;
            if (node.contentNotMeasured)
            {
                stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_CONTENTS_NOT_MEASURED);
            }
            else
                switch (node.scanState)
                {
                    case ScanState::NotStarted: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_NOT_STARTED); break;
                    case ScanState::Queued: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_QUEUED); break;
                    case ScanState::Scanning: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_SCANNING); break;
                    case ScanState::Done: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_DONE); break;
                    case ScanState::Error: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_ERROR); break;
                    case ScanState::Canceled: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_CANCELED); break;
                }

            return withModified(FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_FORMAT_WITH_PATH, name, pathText, sizeText, shareText, stateText));
        }

        return withModified(FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_FORMAT_WITH_PATH_NO_STATE, name, pathText, sizeText, shareText));
    }

    if (node.contentNotMeasured || node.scanState != ScanState::Done)
    {
        std::wstring stateText;
        if (node.contentNotMeasured)
        {
            stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_CONTENTS_NOT_MEASURED);
        }
        else
            switch (node.scanState)
            {
                case ScanState::NotStarted: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_NOT_STARTED); break;
                case ScanState::Queued: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_QUEUED); break;
                case ScanState::Scanning: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_SCANNING); break;
                case ScanState::Done: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_DONE); break;
                case ScanState::Error: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_ERROR); break;
                case ScanState::Canceled: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_CANCELED); break;
            }

        return withModified(FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_FORMAT_NO_PATH, name, sizeText, shareText, stateText));
    }

    return withModified(FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_FORMAT_NO_PATH_NO_STATE, name, sizeText, shareText));
}

void ViewerSpace::StartScan(std::wstring_view rootPath, bool allowCache, bool retainPrevious)
{
    const bool preserveDisplayedModel = retainPrevious && _displayModel != nullptr;
    if (_workingModelRetirementPending && HasDeferredDisposalCapacity())
    {
        _workingModelRetirementPending = false;
        RetireDisplayedStorage({}, std::move(_workingModel), false);
    }

    const bool requiresRetirement =
        (_workingModelRetirementPending && preserveDisplayedModel) || (! preserveDisplayedModel && (_displayModel != nullptr || _workingModel != nullptr));
    if (requiresRetirement && ! HasDeferredDisposalCapacity())
    {
        const bool interruptedActiveScan =
            _scanActive.load(std::memory_order_acquire) || _overallState == ScanState::Queued || _overallState == ScanState::Scanning;
        if (_workingModel != nullptr && _refreshPending)
        {
            DiscardWorkingModel(false);
        }

        _scanGeneration.fetch_add(1u, std::memory_order_acq_rel);
        CancelScan();
        CancelScanCacheBuild();
        CancelScanCacheRestore();
        {
            std::scoped_lock lock(_updateMutex);
            _pendingUpdates.clear();
            _pendingUpdateBytes.store(0u, std::memory_order_relaxed);
        }
        _scanActive.store(false, std::memory_order_release);
        if (_workingModel == nullptr && interruptedActiveScan)
        {
            if (Node* root = TryGetRealNode(_model->rootNodeId); root != nullptr)
            {
                root->scanState = ScanState::Canceled;
            }
            _overallState = ScanState::Canceled;
        }

        DeferredStartScan request;
        request.rootPath          = std::wstring(rootPath);
        request.requestGeneration = ++_deferredStartScanGeneration;
        request.allowCache        = allowCache;
        request.retainPrevious    = retainPrevious;
        _deferredStartScan        = std::move(request);
        Debug::Perf::EmitCounter(L"viewer.space.refresh.start_deferred");
        Debug::Perf::EmitValue(L"viewer.space.refresh.start_deferred_generation", _deferredStartScanGeneration);
        UpdateHeaderTextCache();
        EnsureAnimationTimer();
        if (_hWnd)
        {
            UpdateWindowTitle(_hWnd.get());
            UpdateMenuState(_hWnd.get());
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        return;
    }

    if (_deferredStartScan.has_value())
    {
        ++_deferredStartScanGeneration;
        _deferredStartScan.reset();
    }

    if (preserveDisplayedModel && ! _refreshDisplayState.has_value())
    {
        RefreshDisplayState state;
        state.overallState              = _overallState;
        state.completedSinceSeconds     = _scanCompletedSinceSeconds;
        state.progressBytes             = _scanProgressBytes;
        state.progressFolders           = _scanProgressFolders;
        state.progressFiles             = _scanProgressFiles;
        state.failedFolders             = _scanFailedFolders;
        state.effectiveThreads          = _scanEffectiveThreads;
        state.topFilesPerDirectory      = _scanTopFilesPerDirectory;
        state.cacheLastStoredGeneration = _scanCacheLastStoredGeneration;
        state.cacheSnapshotBytes        = _lastScanCacheSnapshotBytes;
        state.cacheRestoreBytes         = _lastScanCacheRestoreBytes;
        state.cacheResidentBytes        = _lastScanCacheResidentBytes;
        state.cacheRestoreUs            = _lastScanCacheRestoreUs;
        state.processingNodeId          = _scanProcessingNodeId;
        state.processingFolderName      = _scanProcessingFolderName;
        state.acceptedEntries           = _lastModelAcceptedEntries;
        state.rejectedEntries           = _lastModelRejectedEntries;
        state.cappedDirectories         = _lastModelCappedDirectories;
        state.cappedFiles               = _lastModelCappedFiles;
        state.retainedNameBytes         = _lastModelRetainedNameBytes;
        state.retainedChildReferences   = _lastModelRetainedChildReferences;
        state.traversedDirectories      = _lastModelTraversedDirectories;
        state.summaryCommittedCount     = _directorySummaryCommittedCount;
        state.summaryRejectedCount      = _directorySummaryRejectedCount;
        state.summaryChunkMaxBytes      = _directorySummaryChunkMaxBytes.load(std::memory_order_relaxed);
        state.validationError           = _lastModelValidationError;
        state.viewNodeId                = _model->viewNodeId;
        state.selectedItemId            = _selectedItemId;
        state.selectionFromAnalysis     = _selectionFromAnalysis;
        state.viewPathText              = _viewPathText;
        state.navStack                  = _navStack;
        state.navForwardStack           = _navForwardStack;
        if (_selectedItemId != 0u)
        {
            if (const FileRecord* file = TryGetFileRecord(_selectedItemId))
            {
                state.selectionPathIsFile = true;
                state.selectedRelativePath.emplace_back(file->name);
                uint32_t parentId = file->parentId;
                while (parentId != 0u && parentId != _model->viewNodeId)
                {
                    const Node* parent = TryGetRealNode(parentId);
                    if (parent == nullptr || parent->isSynthetic || parent->name.empty())
                    {
                        state.selectedRelativePath.clear();
                        break;
                    }
                    state.selectedRelativePath.emplace_back(parent->name);
                    parentId = parent->parentId;
                }
                if (parentId != _model->viewNodeId)
                {
                    state.selectedRelativePath.clear();
                }
            }
            else if (const Node* node = TryGetRealNode(_selectedItemId); node != nullptr && ! node->isSynthetic)
            {
                uint32_t currentId = node->id;
                while (currentId != 0u && currentId != _model->viewNodeId)
                {
                    const Node* current = TryGetRealNode(currentId);
                    if (current == nullptr || current->name.empty())
                    {
                        state.selectedRelativePath.clear();
                        break;
                    }
                    state.selectedRelativePath.emplace_back(current->name);
                    currentId = current->parentId;
                }
                if (currentId != _model->viewNodeId)
                {
                    state.selectedRelativePath.clear();
                }
            }
            std::reverse(state.selectedRelativePath.begin(), state.selectedRelativePath.end());
        }
        _refreshDisplayState = std::move(state);
    }

    if (preserveDisplayedModel)
    {
        const uint64_t minimalCandidateBytes = sizeof(ModelGeneration) + static_cast<uint64_t>(rootPath.size() + 1u) * sizeof(wchar_t);
        uint64_t baseBytes                   = ViewerOwnedBytes();
        if (_workingModel != nullptr)
        {
            const uint64_t currentCandidateBytes = _workingModel->AllocatedBytes();
            baseBytes                            = currentCandidateBytes > baseBytes ? 0u : baseBytes - currentCandidateBytes;
        }
        if (! FitsViewerSpaceRefreshWorkingSet(baseBytes, minimalCandidateBytes))
        {
#ifdef ENABLE_TESTS
            Debug::Perf::Emit(L"viewer.space.refresh.admission_refused", L"start-scan", 0u, baseBytes, minimalCandidateBytes, E_OUTOFMEMORY);
#endif
            _refreshPending         = false;
            _refreshRefused         = true;
            _refreshFailed          = false;
            _refreshAdoptionPending = false;
            _refreshDisplayState.reset();
            UpdateHeaderTextCache();
            if (_hWnd)
            {
                UpdateMenuState(_hWnd.get());
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
            }
            return;
        }
    }

    const uint32_t generation = _scanGeneration.fetch_add(1) + 1;
    CancelScan();
    ReapFinishedScanWorkers();
    CancelScanCacheBuild();
    CancelScanCacheRestore();
    EnsureAnimationTimer();
    _scanCacheLastStoredGeneration    = 0;
    _scanCompletedSinceSeconds        = 0.0;
    _lastScanCacheSnapshotBytes       = 0u;
    _lastScanCacheRestoreBytes        = 0u;
    _lastScanCacheResidentBytes       = 0u;
    _lastScanCacheRestoreUs           = 0u;
    _lastModelAcceptedEntries         = 0u;
    _lastModelRejectedEntries         = 0u;
    _lastModelCappedDirectories       = 0u;
    _lastModelCappedFiles             = 0u;
    _lastModelRetainedNameBytes       = 0u;
    _lastModelRetainedChildReferences = 0u;
    _lastModelTraversedDirectories    = 0u;
    _lastModelValidationError         = ViewerSpaceScan::ValidationError::None;
    _scanResourcePolicy               = ActiveScanResourcePolicy();
#ifdef ENABLE_TESTS
    _debugPostUpdateInnerGenerationRejects.store(0u, std::memory_order_release);
#endif
    if (preserveDisplayedModel)
    {
        _workingModel           = std::make_unique<ModelGeneration>();
        _model                  = _workingModel.get();
        _model->generation      = generation;
        _refreshPending         = true;
        _refreshRefused         = false;
        _refreshFailed          = false;
        _refreshAdoptionPending = false;
        _refreshPathCandidates.clear();
        _refreshNextPathCandidates.clear();
        _refreshPathCandidateIndex  = 0u;
        _refreshPathChildIndex      = 0u;
        _refreshResolvedSelectionId = 0u;
        _refreshPathLookupStarted   = false;
        _workingScanGeneration      = generation;
    }
    else
    {
        _workingModelRetirementPending = false;
        RetireDisplayedStorage(std::move(_displayModel), std::move(_workingModel));
        _displayModel           = std::make_unique<ModelGeneration>();
        _model                  = _displayModel.get();
        _model->generation      = generation;
        _displayGeneration      = generation;
        _refreshPending         = false;
        _refreshRefused         = false;
        _refreshFailed          = false;
        _refreshAdoptionPending = false;
        _refreshDisplayState.reset();
        _workingScanGeneration = 0u;
        _selectedItemId        = 0u;
        _selectedGeneration    = generation;
        _selectionFromAnalysis = false;
        // Rebind the model before synchronous Grid callbacks; disposal may have freed the old owner.
        _analysisHost.NotifyRowsChanged(std::nullopt);
    }

    if (! preserveDisplayedModel)
    {
        _lastRectsByNode.clear();
    }
    {
        std::scoped_lock lock(_updateMutex);
        _pendingUpdatePostedCount.store(0u, std::memory_order_relaxed);
        _pendingUpdateCoalescedCount.store(0u, std::memory_order_relaxed);
        _pendingUpdateBytes.store(0u, std::memory_order_relaxed);
        _pendingUpdateAcceptedBytes.store(0u, std::memory_order_relaxed);
        _pendingUpdatePeakBytes.store(0u, std::memory_order_relaxed);
        _pendingUpdateLockHoldCount.store(0u, std::memory_order_relaxed);
        _pendingUpdateLockHoldTotalUs.store(0u, std::memory_order_relaxed);
        _pendingUpdateLockHoldMaxUs.store(0u, std::memory_order_relaxed);
        _pendingUpdateDrainedCount.store(0u, std::memory_order_relaxed);
        _pendingUpdateDrainCount.store(0u, std::memory_order_relaxed);
        _pendingUpdateDrainTotalUs.store(0u, std::memory_order_relaxed);
        _pendingUpdateDrainMaxUs.store(0u, std::memory_order_relaxed);
        _lastQueueDiagnosticsEmitAt      = {};
        _queueDiagnosticsTerminalEmitted = false;
    }
    _providerBufferPeakBytes.store(_providerBufferCurrentBytes.load(std::memory_order_relaxed), std::memory_order_relaxed);
    _validationScratchPeakBytes.store(_validationScratchCurrentBytes.load(std::memory_order_relaxed), std::memory_order_relaxed);
    _validationScratchRejectedCount.store(0u, std::memory_order_relaxed);
    _activePathPeakBytes.store(_activePathCurrentBytes.load(std::memory_order_relaxed), std::memory_order_relaxed);
    _activePathRejectedCount.store(0u, std::memory_order_relaxed);
    if (! preserveDisplayedModel)
    {
        InvalidateStaticTreemapCache();
        _staticTreemapCache.hits         = 0u;
        _staticTreemapCache.misses       = 0u;
        _staticTreemapCache.recordCount  = 0u;
        _staticTreemapCache.lastRecordUs = 0u;
    }
    _openStartedAt           = std::chrono::steady_clock::now();
    _openToFirstPaintEmitted = false;
#ifdef ENABLE_TESTS
    _lastPaintUs         = 0u;
    _lastLayoutUs        = 0u;
    _lastDrainUs         = 0u;
    _lastWorkingSetBytes = 0u;
    _lastTileDrawCount   = 0u;
    _lastTextDrawCount   = 0u;
    if (! preserveDisplayedModel)
    {
        _layoutGeneration = 0u;
    }
    _lastHitTestUs                = 0u;
    _lastHitTestCandidatesChecked = 0u;
#endif

    std::wstring scanRootPath(rootPath);
    _model->scanRootPath = scanRootPath;
    _model->scanRootParentPath.reset();

    const bool rootLooksWin32 = LooksLikeWin32Path(_model->scanRootPath);
    if (_fileSystemIsWin32 || rootLooksWin32)
    {
        const std::optional<std::filesystem::path> parent = TryGetParentPathForNavigation(std::filesystem::path(_model->scanRootPath));
        if (parent.has_value())
        {
            _model->scanRootParentPath = parent.value().wstring();
        }
    }
    else
    {
        _model->scanRootParentPath = TryGetParentPathForNavigationGeneric(_model->scanRootPath);
    }

    if (! preserveDisplayedModel)
    {
        _syntheticNodes.clear();
        _layoutMaxItemsByNode.clear();
        _autoExpandedOtherByNode.clear();
        _layoutCandidateCache.clear();
        _layoutCandidateCacheHits   = 0u;
        _layoutCandidateCacheMisses = 0u;
        _analysisBuild.reset();
        _analysisRows.clear();
        _publishedAnalysisOwner.reset();
        _analysisDirty = true;
        _drawItems.clear();
        _hitGrid.Clear();
        _navStack.clear();
        _navForwardStack.clear();
        _layoutDirty                  = true;
        _hoverNodeId                  = 0;
        _tooltipNodeId                = 0;
        _tooltipCandidateNodeId       = 0;
        _tooltipCandidateSinceSeconds = 0.0;
    }
    _model->nodes.resize(2u);
    if (! preserveDisplayedModel)
    {
        UpdateTooltipForHit(0);
    }

    {
        std::scoped_lock lock(_updateMutex);
        _pendingUpdates.clear();
        _pendingUpdateBytes.store(0u, std::memory_order_relaxed);
    }

    _scanProgressBytes       = 0;
    _scanPreviewPriorityJobs = 0u;
    _scanProgressFolders     = 0;
    _scanProgressFiles       = 0;
    _scanFailedFolders       = 0;
    _scanProcessingNodeId    = 0;
    _scanProcessingFolderName.clear();
    _headerStatusId            = 0;
    _headerStatusFailedFolders = 0;
    _headerStatusText.clear();
    _headerCountsText.clear();
    _headerSizeText.clear();
    _headerViewText.clear();
    _headerProcessingText.clear();

    const uint32_t effectiveTopFilesPerDirectory = ComputeAdaptiveFileCandidateBudget();
    _scanTopFilesPerDirectory                    = effectiveTopFilesPerDirectory;
    const size_t topFilesPerDirectory            = static_cast<size_t>(effectiveTopFilesPerDirectory);
    uint32_t scanThreads =
        _fileSystemIsWin32 ? std::clamp<uint32_t>(std::max(_config.scanThreads, DefaultWin32ScanThreads()), 1u, 16u) : std::clamp(_config.scanThreads, 1u, 16u);
#ifdef ENABLE_TESTS
    // The focused serial/parallel golden test must exercise a true one-worker Win32 scan without changing production defaults.
    if (const std::optional<uint32_t> testScanThreads = ReadPositiveUint32Environment(kTestScanThreadsEnvVar); testScanThreads.has_value())
    {
        scanThreads = std::clamp(testScanThreads.value(), 1u, 16u);
    }
#endif
    // Spinners and forced-visible scanning tiles follow the workers that actually run, not the configured floor.
    _scanEffectiveThreads = scanThreads;

    if (allowCache && _config.cacheEnabled && _fileSystemIsWin32)
    {
        ScanResultCacheKey cacheKey;
        cacheKey.rootKey              = NormalizeRootPathForScanCache(std::filesystem::path(_model->scanRootPath));
        cacheKey.topFilesPerDirectory = effectiveTopFilesPerDirectory;

        if (! cacheKey.rootKey.empty())
        {
            const std::shared_ptr<const ScanResultSnapshot> snapshot = GetScanResultCache().TryGet(cacheKey);
#ifdef ENABLE_TESTS
            if (snapshot && ReadPositiveUint32Environment(kTestCorruptCacheNameOffsetEnvVar).has_value() && snapshot->nodes.size() > 1u)
            {
                // Test-only fault injection: force the next reopened cache entry through the fail-closed extent validator.
                std::const_pointer_cast<ScanResultSnapshot>(snapshot)->nodes[1].nameOffset = std::numeric_limits<uint32_t>::max();
            }
            if (snapshot && ReadPositiveUint32Environment(kTestCorruptCacheRootStateEnvVar).has_value() && snapshot->nodes.size() > 1u)
            {
                // Test-only fault injection: completed cache entries cannot restore a nonterminal root state.
                std::const_pointer_cast<ScanResultSnapshot>(snapshot)->nodes[1].scanState = static_cast<uint8_t>(ScanState::Scanning);
            }
            if (snapshot && ReadPositiveUint32Environment(kTestCorruptCacheChildOwnerEnvVar).has_value())
            {
                // Test-only fault injection: replace the root's first valid child reference with the root itself.
                auto& mutableSnapshot = *std::const_pointer_cast<ScanResultSnapshot>(snapshot);
                if (mutableSnapshot.nodes.size() > 1u)
                {
                    const ScanResultCacheNode& rootNode = mutableSnapshot.nodes[1u];
                    const size_t childStart             = static_cast<size_t>(rootNode.childrenStart);
                    const size_t childCount             = static_cast<size_t>(rootNode.childrenCount);
                    if (rootNode.id == 1u && childCount != 0u && childStart <= mutableSnapshot.childrenArena.size() &&
                        childCount <= mutableSnapshot.childrenArena.size() - childStart)
                    {
                        mutableSnapshot.childrenArena[childStart] = rootNode.id;
                        Debug::Perf::EmitCounter(L"viewer.space.scan.test.cache_child_owner_fault_applied");
                    }
                    else
                    {
                        Debug::Perf::EmitCounter(L"viewer.space.scan.test.cache_child_owner_fault_unavailable");
                    }
                }
                else
                {
                    Debug::Perf::EmitCounter(L"viewer.space.scan.test.cache_child_owner_fault_unavailable");
                }
            }
            if (snapshot && ReadPositiveUint32Environment(kTestCorruptCacheMissingChildEnvVar).has_value())
            {
                // Test-only fault injection: omit the final root child while retaining its active record.
                auto& mutableSnapshot = *std::const_pointer_cast<ScanResultSnapshot>(snapshot);
                if (mutableSnapshot.nodes.size() > 1u)
                {
                    ScanResultCacheNode& rootNode = mutableSnapshot.nodes[1u];
                    const size_t childStart       = static_cast<size_t>(rootNode.childrenStart);
                    const size_t childCount       = static_cast<size_t>(rootNode.childrenCount);
                    if (rootNode.id == 1u && childCount != 0u && childStart <= mutableSnapshot.childrenArena.size() &&
                        childCount <= mutableSnapshot.childrenArena.size() - childStart)
                    {
                        --rootNode.childrenCount;
                        Debug::Perf::EmitCounter(L"viewer.space.scan.test.cache_missing_child_fault_applied");
                    }
                    else
                    {
                        Debug::Perf::EmitCounter(L"viewer.space.scan.test.cache_missing_child_fault_unavailable");
                    }
                }
                else
                {
                    Debug::Perf::EmitCounter(L"viewer.space.scan.test.cache_missing_child_fault_unavailable");
                }
            }
#endif
            if (snapshot && snapshot->nodes.size() > 1 && snapshot->nodes[1].id == 1)
            {
                _scanCacheRestoreSnapshot             = snapshot;
                _scanCacheRestoreRootKey              = cacheKey.rootKey;
                _scanCacheRestoreTopFilesPerDirectory = snapshot->topFilesPerDirectory;
                _scanCacheRestoreGeneration           = generation;
                _scanCacheRestorePhase                = static_cast<uint8_t>(ScanCacheRestorePhase::ValidateNodes);
                _scanCacheRestoreValidationNext       = 1u;
                _scanCacheRestoreValidationChildNext  = 0u;
                _scanCacheRestoreChildrenNext         = 0u;
                _scanCacheRestoreNodesNext            = 0u;
                _scanCacheRestoreFileRecordsNext      = 0u;
                _scanCacheRestoreStartedAtSeconds     = NowSeconds();
                _overallState                         = ScanState::Queued;
                _model->rootNodeId                    = 1u;
                _model->viewNodeId                    = 1u;
                _scanActive.store(true);
                _animationStartSeconds    = NowSeconds();
                _lastLayoutRebuildSeconds = 0.0;
                if (! preserveDisplayedModel)
                {
                    _firstPositiveProjectionAttempted = false;
                    _hasPublishedPositiveGeometry     = false;
                    UpdateViewPathText();
                }
                UpdateHeaderTextCache();

                if (preserveDisplayedModel)
                {
                    _model = _displayModel.get();
                }
                else
                {
                    UpdateInspector();
                }

                if (_hWnd)
                {
                    UpdateWindowTitle(_hWnd.get());
                    UpdateMenuState(_hWnd.get());
                    InvalidateRect(_hWnd.get(), nullptr, FALSE);
                }
                return;
            }
            if (snapshot)
            {
                GetScanResultCache().DiscardIfMatches(cacheKey, snapshot);
                Debug::Perf::EmitCounter(L"viewer.space.scan.cache_invalid_snapshot");
            }
            Debug::Perf::EmitCounter(L"viewer.space.scan.cache_miss");
        }
    }

    _overallState = ScanState::Queued;
    _scanActive.store(true);
    _animationStartSeconds    = NowSeconds();
    _lastLayoutRebuildSeconds = 0.0;
    if (! preserveDisplayedModel)
    {
        _firstPositiveProjectionAttempted = false;
        _hasPublishedPositiveGeometry     = false;
    }

    Node root;
    root.id          = 1;
    root.parentId    = 0;
    root.isDirectory = true;
    root.scanState   = ScanState::Queued;
    std::wstring rootName;
    if (_fileSystemIsWin32 || rootLooksWin32)
    {
        const std::filesystem::path rootFs(_model->scanRootPath);
        rootName = rootFs.filename().wstring();
        if (rootName.empty())
        {
            rootName = rootFs.wstring();
        }
    }
    else
    {
        const std::wstring_view trimmed = TrimTrailingPathSeparators(_model->scanRootPath);
        const size_t lastSep            = trimmed.find_last_of(L"/\\");
        if (lastSep != std::wstring_view::npos && lastSep + 1 < trimmed.size())
        {
            rootName = std::wstring(trimmed.substr(lastSep + 1));
        }
        if (rootName.empty())
        {
            rootName = std::wstring(trimmed);
        }
        if (rootName.empty())
        {
            rootName = L"/";
        }
    }

    root.name = CopyToArena(_model->nameArena, rootName);
    if (root.name.empty() && ! _model->scanRootPath.empty())
    {
        root.name = CopyToArena(_model->nameArena, _model->scanRootPath);
    }

    _model->nodes[root.id] = std::move(root);
    _model->rootNodeId     = 1;
    _model->viewNodeId     = 1;
    if (! preserveDisplayedModel)
    {
        UpdateViewPathText();
    }

    UpdateHeaderTextCache();

    if (preserveDisplayedModel)
    {
        _model = _displayModel.get();
    }
    else
    {
        UpdateInspector();
    }

    auto done                                                = std::make_shared<std::atomic_bool>(false);
    wil::com_ptr<IFileSystem> scanFileSystem                 = _fileSystem;
    const bool fileSystemIsWin32                             = _fileSystemIsWin32;
    const ViewerSpaceScan::ResourcePolicy scanResourcePolicy = _scanResourcePolicy;
    uint32_t initialModelNodeId                              = 2u;
#ifdef ENABLE_TESTS
    if (const std::optional<uint32_t> value = ReadPositiveUint32Environment(kTestNextModelIdEnvVar); value.has_value())
    {
        initialModelNodeId = value.value();
    }
#endif
    ViewerSpace* const workerOwner = this;
    workerOwner->AddRef();
    auto releaseWorkerReferenceOnFailure = wil::scope_exit([workerOwner]() noexcept { workerOwner->Release(); });
    g_scanWorkerUnloadGates.fetch_add(1u, std::memory_order_acq_rel);
    auto releaseUnloadGateOnFailure = wil::scope_exit([]() noexcept { g_scanWorkerUnloadGates.fetch_sub(1u, std::memory_order_acq_rel); });

    ScanWorker newWorker;
    newWorker.done   = done;
    newWorker.thread = std::jthread([workerOwner,
                                     generation,
                                     done,
                                     scanFileSystem = std::move(scanFileSystem),
                                     fileSystemIsWin32,
                                     scanRootPath = std::move(scanRootPath),
                                     topFilesPerDirectory,
                                     scanThreads,
                                     scanResourcePolicy,
                                     initialModelNodeId](std::stop_token st) mutable noexcept
    {
        auto markDone               = wil::scope_exit([done]() noexcept { done->store(true, std::memory_order_release); });
        auto releaseWorkerReference = wil::scope_exit([workerOwner]() noexcept { workerOwner->Release(); });
        workerOwner->ScanMain(st,
                              generation,
                              std::move(scanFileSystem),
                              fileSystemIsWin32,
                              std::move(scanRootPath),
                              1,
                              initialModelNodeId,
                              topFilesPerDirectory,
                              scanThreads,
                              scanResourcePolicy);
        Debug::Perf::EmitCounter(L"viewer.space.scan.worker_completed");
    });
    releaseUnloadGateOnFailure.release();
    releaseWorkerReferenceOnFailure.release();
    _scanWorker = std::move(newWorker);

    if (_hWnd)
    {
        UpdateWindowTitle(_hWnd.get());
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::CancelScan() noexcept
{
    if (_scanWorker.thread.joinable())
    {
        _scanWorker.thread.request_stop();
        _retiredScanWorkers.emplace_back(std::move(_scanWorker));
        _scanWorker = {};
    }
    _scanActive.store(false);
}

void ViewerSpace::CancelScanByUser() noexcept
{
    if (_deferredStartScan.has_value())
    {
        CancelDeferredStartScan();
        if (_workingModel != nullptr && _refreshPending)
        {
            DiscardWorkingModel(false);
        }
        return;
    }

    if (_workingModel != nullptr && _refreshPending)
    {
        DiscardWorkingModel(false);
        return;
    }

    if (! _scanWorker.thread.joinable() && ! _scanCacheRestoreSnapshot)
    {
        return;
    }

    _scanGeneration.fetch_add(1);

    CancelScan();
    CancelScanCacheBuild();
    CancelScanCacheRestore();
    _scanCacheLastStoredGeneration = 0;

    {
        std::scoped_lock lock(_updateMutex);
        _pendingUpdates.clear();
        _pendingUpdateBytes.store(0u, std::memory_order_relaxed);
    }

    Node* root = TryGetRealNode(_model->rootNodeId);
    if (root != nullptr)
    {
        root->scanState = ScanState::Canceled;
    }
    _overallState = ScanState::Canceled;
    _scanActive.store(false);
    UpdateHeaderTextCache();

    if (_hWnd)
    {
        UpdateWindowTitle(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::AbandonScanWorkers() noexcept
{
    const auto abandonStartedAt = std::chrono::steady_clock::now();
    _scanGeneration.fetch_add(1, std::memory_order_acq_rel);
    _scanActive.store(false, std::memory_order_release);

    if (_scanWorker.thread.joinable())
    {
        _scanWorker.thread.request_stop();
    }

    for (auto& worker : _retiredScanWorkers)
    {
        if (worker.thread.joinable())
        {
            worker.thread.request_stop();
        }
    }

    uint64_t reapedWorkers    = 0u;
    uint64_t abandonedWorkers = 0u;
    const auto finishWorker   = [&reapedWorkers, &abandonedWorkers](ScanWorker& worker) noexcept
    {
        if (! worker.thread.joinable())
        {
            worker = {};
            return;
        }

        if (worker.done != nullptr && worker.done->load(std::memory_order_acquire))
        {
            worker.thread.join();
            g_scanWorkerUnloadGates.fetch_sub(1u, std::memory_order_acq_rel);
            reapedWorkers += 1u;
        }
        else
        {
            // ReadDirectoryInfo is an untrusted provider boundary with no forced-cancel contract.
            // The worker owns a ViewerSpace reference. Intentionally retain its unload gate after
            // detach: there is no safe point inside a std::jthread lambda at which the last module
            // pin could be released after the callback has returned. Runtime reload is therefore
            // quarantined for this process after an uncooperative provider forces abandonment.
            worker.thread.detach();
            abandonedWorkers += 1u;
        }
        worker = {};
    };

    finishWorker(_scanWorker);
    for (ScanWorker& worker : _retiredScanWorkers)
    {
        finishWorker(worker);
    }
    _retiredScanWorkers.clear();

    {
        std::scoped_lock lock(_updateMutex);
        _pendingUpdates.clear();
        _pendingUpdateBytes.store(0u, std::memory_order_relaxed);
    }

    if (reapedWorkers != 0u || abandonedWorkers != 0u)
    {
        Debug::Perf::EmitValue(L"viewer.space.scan.worker_reaped", reapedWorkers);
        Debug::Perf::EmitValue(L"viewer.space.scan.worker_abandoned", abandonedWorkers);
        Debug::Perf::EmitDurationUs(L"viewer.space.close.cancel_us", Debug::Perf::ElapsedUs(abandonStartedAt), abandonedWorkers, reapedWorkers);
    }
}

void ViewerSpace::ReapFinishedScanWorkers() noexcept
{
    if (_scanWorker.thread.joinable())
    {
        if (_scanWorker.done != nullptr && _scanWorker.done->load(std::memory_order_acquire))
        {
            _scanWorker.thread.join();
            g_scanWorkerUnloadGates.fetch_sub(1u, std::memory_order_acq_rel);
            _scanWorker = {};
        }
    }

    for (size_t i = 0; i < _retiredScanWorkers.size();)
    {
        ScanWorker& worker = _retiredScanWorkers[i];
        if (worker.done != nullptr && worker.done->load(std::memory_order_acquire) && worker.thread.joinable())
        {
            worker.thread.join();
            g_scanWorkerUnloadGates.fetch_sub(1u, std::memory_order_acq_rel);
            _retiredScanWorkers.erase(_retiredScanWorkers.begin() + static_cast<ptrdiff_t>(i));
            continue;
        }
        i += 1;
    }
}

bool ViewerSpace::SubmitDeferredDisposal(size_t index) noexcept
{
    if (index >= _deferredDisposalWorkers.size())
    {
        return false;
    }
    DeferredDisposalWorker& worker = _deferredDisposalWorkers[index];
    if (worker.submitted || worker.storage == nullptr || worker.done == nullptr)
    {
        return false;
    }

    auto context       = std::make_unique<DeferredDisposalContext>();
    context->storage   = worker.storage;
    context->done      = worker.done;
    context->bytes     = worker.bytes;
    context->modulePin = AcquireModuleReferenceFromAddress(&kViewerSpaceModuleAnchor);
    if (! context->modulePin)
    {
        worker.retryAfterSeconds = NowSeconds() + 0.25;
        return false;
    }
    context->owner = this;

    const auto submitStartedAt          = std::chrono::steady_clock::now();
    DeferredDisposalContext* rawContext = context.release();
    if (TrySubmitThreadpoolCallback(&ViewerSpace::DeferredDisposalCallback, rawContext, nullptr) == FALSE)
    {
        context.reset(rawContext);
        worker.retryAfterSeconds = NowSeconds() + 0.25;
        Debug::Perf::EmitValue(L"viewer.space.refresh.dispose_submit_failed", 1u);
        return false;
    }

    worker.submitted        = true;
    const uint64_t submitUs = Debug::Perf::ElapsedUs(submitStartedAt);
    UpdateAtomicPeak(_deferredDisposalSubmitMaxUs, submitUs);
    _deferredDisposalSubmitCount     = ViewerSpaceScan::SaturatingAdd(_deferredDisposalSubmitCount, uint64_t{1u});
    _deferredDisposalSubmitTotalUs   = ViewerSpaceScan::SaturatingAdd(_deferredDisposalSubmitTotalUs, submitUs);
    _deferredDisposalSubmitBytes     = ViewerSpaceScan::SaturatingAdd(_deferredDisposalSubmitBytes, worker.bytes);
    _deferredDisposalSubmitOwnerPeak = std::max(_deferredDisposalSubmitOwnerPeak, static_cast<uint64_t>(_deferredDisposalWorkers.size()));
    return true;
}

VOID CALLBACK ViewerSpace::DeferredDisposalCallback(PTP_CALLBACK_INSTANCE instance, void* rawContext) noexcept
{
    std::unique_ptr<DeferredDisposalContext> context(static_cast<DeferredDisposalContext*>(rawContext));
    if (context == nullptr)
    {
        return;
    }

    TransferModulePinToCallbackReturn(instance, context->modulePin);
    ViewerSpace* const owner           = context->owner.get();
    const bool layoutOnly              = context->storage != nullptr && context->storage->layoutOnly;
    const bool layoutCancellation      = context->storage != nullptr && context->storage->layoutCancellation;
    const uint64_t layoutGeometryCount = context->storage == nullptr ? 0u : context->storage->layoutGeometryCount;
    const uint32_t layoutGeneration    = context->storage == nullptr ? 0u : context->storage->layoutGeneration;
#ifdef ENABLE_TESTS
    const ULONGLONG holdStartedAt = GetTickCount64();
    bool holdEntered              = false;
    for (;;)
    {
        std::array<wchar_t, 2> hold{};
        if (GetEnvironmentVariableW(L"REDSALAMANDER_VIEWERSPACE_TEST_HOLD_DISPOSAL_CALLBACK", hold.data(), static_cast<DWORD>(hold.size())) == 0u)
        {
            break;
        }
        if (! holdEntered)
        {
            holdEntered = true;
            Debug::Perf::EmitCounter(layoutOnly ? L"viewer.space.layout.dispose_hold_entered" : L"viewer.space.refresh.dispose_hold_entered");
        }
        if (GetTickCount64() - holdStartedAt >= 30000u)
        {
            Debug::Warning(L"ViewerSpace test disposal hold expired; releasing retired storage.");
            break;
        }
        Sleep(1u);
    }
#endif
    const auto startedAt = std::chrono::steady_clock::now();
    if (context->storage != nullptr)
    {
        context->storage->ReleaseStorage();
    }
    SubtractAtomicFloor(owner->_deferredDisposalBytes, context->bytes);
    const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
    UpdateAtomicPeak(owner->_deferredDisposalMaxUs, elapsedUs);
    if (layoutOnly)
    {
        Debug::Perf::EmitDurationUs(L"viewer.space.layout.dispose_us", elapsedUs, context->bytes, 0u, S_OK);
        Debug::Perf::EmitCounter(L"viewer.space.layout.dispose_completed");
        if (layoutCancellation)
        {
            Debug::Perf::Emit(L"viewer.space.layout.cancel_cleanup_us", L"candidate-retired", elapsedUs, layoutGeometryCount, layoutGeneration, S_FALSE);
        }
    }
    else
    {
        Debug::Perf::EmitDurationUs(L"viewer.space.refresh.dispose_us", elapsedUs, context->bytes, 0u, S_OK);
        Debug::Perf::EmitCounter(L"viewer.space.refresh.dispose_completed");
    }
    if (context->done != nullptr)
    {
        context->done->store(true, std::memory_order_release);
    }
    context->owner.reset();
}

void ViewerSpace::ContinueDeferredDisposalSubmissions() noexcept
{
    const double now = NowSeconds();
    for (size_t index = 0u; index < _deferredDisposalWorkers.size(); ++index)
    {
        const DeferredDisposalWorker& worker = _deferredDisposalWorkers[index];
        if (! worker.submitted && now >= worker.retryAfterSeconds)
        {
            static_cast<void>(SubmitDeferredDisposal(index));
        }
    }
}

void ViewerSpace::ReapFinishedDeferredDisposals() noexcept
{
    for (size_t index = 0u; index < _deferredDisposalWorkers.size();)
    {
        const DeferredDisposalWorker& worker = _deferredDisposalWorkers[index];
        if (worker.submitted && worker.done != nullptr && worker.done->load(std::memory_order_acquire))
        {
            _deferredDisposalWorkers.erase(_deferredDisposalWorkers.begin() + static_cast<ptrdiff_t>(index));
            continue;
        }
        ++index;
    }
}

void ViewerSpace::AbandonDeferredDisposals() noexcept
{
    ContinueDeferredDisposalSubmissions();
    for (DeferredDisposalWorker& worker : _deferredDisposalWorkers)
    {
        if (! worker.submitted && worker.storage != nullptr)
        {
            // While the window lives, failed submissions stay charged and retry on the timer. At final
            // teardown no retry timer remains, so this is the bounded OS-threadpool-submission-failure
            // exception that releases at most the three retained owners inline; submitted callbacks
            // retain their own storage, module pin, and object reference and are never joined here.
            const auto fallbackStartedAt = std::chrono::steady_clock::now();
            worker.storage->ReleaseStorage();
            SubtractAtomicFloor(_deferredDisposalBytes, worker.bytes);
            Debug::Perf::EmitDurationUs(L"viewer.space.refresh.dispose_inline_teardown_submit_failure_us",
                                        Debug::Perf::ElapsedUs(fallbackStartedAt),
                                        worker.bytes,
                                        static_cast<uint64_t>(_deferredDisposalWorkers.size()),
                                        S_OK);
        }
    }
    _deferredDisposalWorkers.clear();
    // Keep synchronous publication free of per-handoff diagnostic writes. The submission
    // endpoint still measures TrySubmitThreadpoolCallback; its count/sum/max and bytes survive.
    Debug::Perf::Emit(L"viewer.space.refresh.dispose_submit_us",
                      L"cumulative-max-since-window-open",
                      _deferredDisposalSubmitMaxUs.load(std::memory_order_relaxed),
                      _deferredDisposalSubmitCount,
                      _deferredDisposalSubmitTotalUs,
                      S_OK);
    Debug::Perf::EmitValue(L"viewer.space.refresh.dispose_submit_bytes_total", _deferredDisposalSubmitBytes);
    Debug::Perf::EmitValue(L"viewer.space.refresh.dispose_submit_owner_peak", _deferredDisposalSubmitOwnerPeak);
}

void ViewerSpace::ScanMain(std::stop_token stopToken,
                           uint32_t generation,
                           wil::com_ptr<IFileSystem> fileSystem,
                           bool fileSystemIsWin32,
                           std::wstring rootPath,
                           uint32_t rootNodeId,
                           uint32_t nextNodeId,
                           size_t topFilesPerDirectory,
                           uint32_t scanThreads,
                           ViewerSpaceScan::ResourcePolicy resourcePolicy) noexcept
{
    static constexpr size_t kProgressUpdateStride = 384;
    static constexpr auto kProgressUpdateInterval = std::chrono::milliseconds(150);

    const uint32_t threadCount = std::clamp(scanThreads, 1u, 16u);

    struct ScanResourceTracker final
    {
        explicit ScanResourceTracker(ViewerSpaceScan::ResourcePolicy scanPolicy, uint64_t rootNameBytes) noexcept
            : policy(scanPolicy),
              retainedNameBytes(std::min(rootNameBytes, scanPolicy.maxRetainedNameBytes))
        {
        }
        ~ScanResourceTracker() = default;

        ScanResourceTracker(const ScanResourceTracker&)            = delete;
        ScanResourceTracker& operator=(const ScanResourceTracker&) = delete;
        ScanResourceTracker(ScanResourceTracker&&)                 = delete;
        ScanResourceTracker& operator=(ScanResourceTracker&&)      = delete;

        [[nodiscard]] bool TryRetainDirectory(uint64_t nameBytes) noexcept
        {
            std::scoped_lock lock(mutex);
            if (retainedDirectories >= policy.maxRetainedDirectories || retainedChildReferences >= policy.maxChildReferences ||
                nameBytes > policy.maxRetainedNameBytes - retainedNameBytes)
            {
                cappedDirectories.fetch_add(1u, std::memory_order_relaxed);
                return false;
            }
            retainedDirectories += 1u;
            retainedChildReferences += 1u;
            retainedNameBytes += nameBytes;
            return true;
        }

        [[nodiscard]] bool TryRetainFile(uint64_t nameBytes) noexcept
        {
            std::scoped_lock lock(mutex);
            if (retainedFiles >= policy.maxRetainedFileRecords || retainedChildReferences >= policy.maxChildReferences ||
                nameBytes > policy.maxRetainedNameBytes - retainedNameBytes)
            {
                cappedFiles.fetch_add(1u, std::memory_order_relaxed);
                return false;
            }
            retainedFiles += 1u;
            retainedChildReferences += 1u;
            retainedNameBytes += nameBytes;
            return true;
        }

        [[nodiscard]] bool TryTraverseDirectories(uint64_t count) noexcept
        {
            if (count == 0u)
            {
                return true;
            }
            uint64_t current = traversedDirectories.load(std::memory_order_relaxed);
            for (;;)
            {
                if (current > policy.maxTraversedDirectories || count > policy.maxTraversedDirectories - current)
                {
                    return false;
                }
                if (traversedDirectories.compare_exchange_weak(current, current + count, std::memory_order_relaxed, std::memory_order_relaxed))
                {
                    return true;
                }
            }
        }

        void RecordValidationFailure(ViewerSpaceScan::ValidationError error) noexcept
        {
            rejectedEntries.fetch_add(1u, std::memory_order_relaxed);
            uint32_t expected = static_cast<uint32_t>(ViewerSpaceScan::ValidationError::None);
            static_cast<void>(validationError.compare_exchange_strong(expected, static_cast<uint32_t>(error), std::memory_order_relaxed));
        }

        void Snapshot(uint64_t& outDirectories, uint64_t& outFiles, uint64_t& outChildReferences, uint64_t& outNameBytes) noexcept
        {
            std::scoped_lock lock(mutex);
            outDirectories     = retainedDirectories;
            outFiles           = retainedFiles;
            outChildReferences = retainedChildReferences;
            outNameBytes       = retainedNameBytes;
        }

        ViewerSpaceScan::ResourcePolicy policy{};
        std::mutex mutex;
        uint64_t retainedDirectories     = 1u;
        uint64_t retainedFiles           = 0u;
        uint64_t retainedChildReferences = 0u;
        uint64_t retainedNameBytes       = 0u;
        std::atomic_uint64_t acceptedEntries{0u};
        std::atomic_uint64_t rejectedEntries{0u};
        std::atomic_uint64_t cappedDirectories{0u};
        std::atomic_uint64_t cappedFiles{0u};
        std::atomic_uint64_t traversedDirectories{1u};
        std::atomic_uint32_t validationError{static_cast<uint32_t>(ViewerSpaceScan::ValidationError::None)};
    };

    struct ValidationScratchBudget final
    {
        enum class Admission : uint8_t
        {
            Acquired,
            Canceled,
            TooLarge,
        };

        ValidationScratchBudget(uint64_t byteLimit, std::atomic_uint64_t& currentBytes, std::atomic_uint64_t& peakBytes) noexcept
            : limit(byteLimit),
              current(currentBytes),
              peak(peakBytes)
        {
        }

        ValidationScratchBudget(const ValidationScratchBudget&)            = delete;
        ValidationScratchBudget& operator=(const ValidationScratchBudget&) = delete;
        ValidationScratchBudget(ValidationScratchBudget&&)                 = delete;
        ValidationScratchBudget& operator=(ValidationScratchBudget&&)      = delete;

        [[nodiscard]] Admission Acquire(uint64_t bytes, std::stop_token stopToken) noexcept
        {
            if (bytes > limit)
            {
                return Admission::TooLarge;
            }
            for (;;)
            {
                if (stopToken.stop_requested())
                {
                    return Admission::Canceled;
                }
                uint64_t admittedBytes = 0u;
                if (TryReserveAtomicBudget(current, bytes, limit, admittedBytes))
                {
                    UpdateAtomicPeak(peak, admittedBytes);
                    return Admission::Acquired;
                }
                std::this_thread::sleep_for(std::chrono::milliseconds(25));
            }
        }

        void Release(uint64_t bytes) noexcept
        {
            static_cast<void>(ReleaseAtomicBudget(current, bytes));
        }

        const uint64_t limit;
        std::atomic_uint64_t& current;
        std::atomic_uint64_t& peak;
    };

    const uint64_t rootNameBytes = rootPath.size() > ((std::numeric_limits<uint64_t>::max)() / sizeof(wchar_t))
                                       ? (std::numeric_limits<uint64_t>::max)()
                                       : static_cast<uint64_t>(rootPath.size()) * sizeof(wchar_t);
    ScanResourceTracker resources(resourcePolicy, rootNameBytes);
    ValidationScratchBudget validationScratchBudget(ActiveValidationScratchBudgetBytes(), _validationScratchCurrentBytes, _validationScratchPeakBytes);
    ValidationScratchBudget workerStagingBudget(ActiveWorkerStagingBudgetBytes(), _workerStagingCurrentBytes, _workerStagingPeakBytes);
    uint32_t nextModelId = nextNodeId;
    std::atomic_uint64_t nextTraversalId(2u);
    std::atomic_uint64_t scannedFolders(0);
    std::atomic_uint64_t scannedFiles(0);
    std::atomic_uint64_t scannedBytes(0);
    std::atomic_uint32_t firstPreviewPriorityJobs(0u);
    std::atomic_uint64_t failedFolders(0);
    std::atomic_uint64_t stagingPreflightCount(0u);
    std::atomic_uint64_t stagingPreflightTotalUs(0u);
    std::atomic_uint64_t stagingPreflightMaxUs(0u);
    std::atomic_uint64_t stagingPreflightReservationPeakBytes(0u);
    std::atomic_uint64_t stagingPreflightRecordTotal(0u);
    std::atomic_uint32_t cpuPriorityAppliedCount(0u);
    std::atomic_uint32_t cpuPriorityRestoredCount(0u);
    std::atomic_long cpuPriorityFailure(S_OK);
#ifdef ENABLE_TESTS
    const bool captureScanPhases = ReadPositiveUint32Environment(L"REDSALAMANDER_VIEWERSPACE_SCAN_PHASES").has_value();
    // Unlike Debug::Perf::Scope, these worker timers aggregate atomically and emit only once after all workers join.
    struct ScanPhaseAggregate final
    {
        ScanPhaseAggregate()                                     = default;
        ScanPhaseAggregate(const ScanPhaseAggregate&)            = delete;
        ScanPhaseAggregate& operator=(const ScanPhaseAggregate&) = delete;
        ScanPhaseAggregate(ScanPhaseAggregate&&)                 = delete;
        ScanPhaseAggregate& operator=(ScanPhaseAggregate&&)      = delete;

        std::atomic_uint64_t count{0u};
        std::atomic_uint64_t totalUs{0u};
        std::atomic_uint64_t maxUs{0u};
    };
    struct ScanPhaseTimer final
    {
        ScanPhaseTimer(bool enabled, ScanPhaseAggregate& aggregate) noexcept : _enabled(enabled), _aggregate(aggregate)
        {
            if (_enabled)
            {
                _startedAt = std::chrono::steady_clock::now();
                _running   = true;
            }
        }
        ~ScanPhaseTimer() noexcept
        {
            Stop();
        }
        void Stop() noexcept
        {
            if (! _enabled || _recorded)
            {
                return;
            }
            if (_running)
            {
                _accumulatedUs = ViewerSpaceScan::SaturatingAdd(_accumulatedUs, Debug::Perf::ElapsedUs(_startedAt));
                _running       = false;
            }
            SaturatingAtomicAdd(_aggregate.count, 1u);
            SaturatingAtomicAdd(_aggregate.totalUs, _accumulatedUs);
            UpdateAtomicPeak(_aggregate.maxUs, _accumulatedUs);
            _recorded = true;
        }
        void Pause() noexcept
        {
            if (_enabled && _running)
            {
                _accumulatedUs = ViewerSpaceScan::SaturatingAdd(_accumulatedUs, Debug::Perf::ElapsedUs(_startedAt));
                _running       = false;
            }
        }
        void Resume() noexcept
        {
            if (_enabled && ! _running && ! _recorded)
            {
                _startedAt = std::chrono::steady_clock::now();
                _running   = true;
            }
        }
        ScanPhaseTimer(const ScanPhaseTimer&)            = delete;
        ScanPhaseTimer& operator=(const ScanPhaseTimer&) = delete;

        bool _enabled  = false;
        bool _running  = false;
        bool _recorded = false;
        ScanPhaseAggregate& _aggregate;
        std::chrono::steady_clock::time_point _startedAt{};
        uint64_t _accumulatedUs = 0u;
    };
    ScanPhaseAggregate providerReadPhase;
    ScanPhaseAggregate listingValidateAdmitPhase;
    ScanPhaseAggregate workerStagingAdmitPhase;
    ScanPhaseAggregate candidateBuildPhase;
    ScanPhaseAggregate modelRecordRetainPhase;
    ScanPhaseAggregate summaryEncodePhase;
#endif
#ifdef ENABLE_TESTS
    std::atomic_bool summaryFaultInjected(false);
#endif

    const auto scanStartedAt = std::chrono::steady_clock::now();
    ScanState scanFinalState = ScanState::Done;
    HRESULT scanFinalHr      = S_OK;
    auto emitScanSummary     = wil::scope_exit([&]() noexcept
    {
        std::wstring_view detail = L"done";
        switch (scanFinalState)
        {
            case ScanState::Canceled: detail = L"canceled"; break;
            case ScanState::Error: detail = L"error"; break;
            case ScanState::Queued: detail = L"queued"; break;
            case ScanState::Scanning: detail = L"scanning"; break;
            case ScanState::Done: detail = L"done"; break;
            case ScanState::NotStarted:
            default: detail = L"not-started"; break;
        }

        const uint64_t folders = scannedFolders.load(std::memory_order_relaxed);
        const uint64_t files   = scannedFiles.load(std::memory_order_relaxed);
        const uint64_t bytes   = scannedBytes.load(std::memory_order_relaxed);
        Debug::Perf::Emit(L"viewer.space.scan.total_us", detail, Debug::Perf::ElapsedUs(scanStartedAt), files, folders, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.files", files, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.folders", folders, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.bytes", bytes, scanFinalHr);
        Debug::Perf::Emit(L"viewer.space.scan.cpu_priority",
                          L"below-normal-cpu-normal-io-restored",
                          0u,
                          cpuPriorityAppliedCount.load(std::memory_order_relaxed),
                          cpuPriorityRestoredCount.load(std::memory_order_relaxed),
                          cpuPriorityFailure.load(std::memory_order_relaxed));
#ifdef ENABLE_TESTS
        if (captureScanPhases)
        {
            const auto emitPhase = [&](std::wstring_view detail, const ScanPhaseAggregate& phase) noexcept
            {
                Debug::Perf::Emit(L"viewer.space.scan.phase_us",
                                  detail,
                                  phase.maxUs.load(std::memory_order_relaxed),
                                  phase.count.load(std::memory_order_relaxed),
                                  phase.totalUs.load(std::memory_order_relaxed),
                                  scanFinalHr);
            };
            emitPhase(L"provider-read-cumulative-max-since-scan-start", providerReadPhase);
            emitPhase(L"listing-validate-admit-cumulative-max-since-scan-start", listingValidateAdmitPhase);
            emitPhase(L"worker-staging-admit-cumulative-max-since-scan-start", workerStagingAdmitPhase);
            emitPhase(L"candidate-build-cumulative-max-since-scan-start", candidateBuildPhase);
            emitPhase(L"model-record-retain-cumulative-max-since-scan-start", modelRecordRetainPhase);
            emitPhase(L"summary-encode-cumulative-max-since-scan-start", summaryEncodePhase);
        }
#endif
        Debug::Perf::EmitValue(L"viewer.space.scan.first_preview_priority_jobs", firstPreviewPriorityJobs.load(std::memory_order_relaxed), scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.active_workers", threadCount, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.provider_buffer_peak_bytes", _providerBufferPeakBytes.load(std::memory_order_relaxed), scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.validation_scratch_peak_bytes", _validationScratchPeakBytes.load(std::memory_order_relaxed), scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.worker_staging_peak_bytes", _workerStagingPeakBytes.load(std::memory_order_relaxed), scanFinalHr);
        Debug::Perf::Emit(L"viewer.space.scan.worker_staging_preflight_us",
                          L"cumulative-max-since-scan-start",
                          stagingPreflightMaxUs.load(std::memory_order_relaxed),
                          stagingPreflightCount.load(std::memory_order_relaxed),
                          stagingPreflightTotalUs.load(std::memory_order_relaxed),
                          scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.worker_staging_preflight_reservation_peak_bytes",
                               stagingPreflightReservationPeakBytes.load(std::memory_order_relaxed),
                               scanFinalHr);
        Debug::Perf::EmitValue(
            L"viewer.space.scan.worker_staging_preflight_record_total", stagingPreflightRecordTotal.load(std::memory_order_relaxed), scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.worker_staging_rejected_count", _workerStagingRejectedCount.load(std::memory_order_relaxed), scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.bookkeeping_peak_bytes", _scannerBookkeepingPeakBytes.load(std::memory_order_relaxed), scanFinalHr);
        Debug::Perf::EmitValue(
            L"viewer.space.scan.validation_scratch_rejected_count", _validationScratchRejectedCount.load(std::memory_order_relaxed), scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.active_path_peak_bytes", _activePathPeakBytes.load(std::memory_order_relaxed), scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.active_path_rejected_count", _activePathRejectedCount.load(std::memory_order_relaxed), scanFinalHr);

        uint64_t retainedDirectories = 0u;
        uint64_t retainedFiles       = 0u;
        uint64_t retainedChildRefs   = 0u;
        uint64_t retainedNameBytes   = 0u;
        resources.Snapshot(retainedDirectories, retainedFiles, retainedChildRefs, retainedNameBytes);
        const uint64_t acceptedEntries   = resources.acceptedEntries.load(std::memory_order_relaxed);
        const uint64_t rejectedEntries   = resources.rejectedEntries.load(std::memory_order_relaxed);
        const uint64_t cappedDirectories = resources.cappedDirectories.load(std::memory_order_relaxed);
        const uint64_t cappedFiles       = resources.cappedFiles.load(std::memory_order_relaxed);
        const uint64_t traversedDirs     = resources.traversedDirectories.load(std::memory_order_relaxed);
        const auto validationError       = static_cast<ViewerSpaceScan::ValidationError>(resources.validationError.load(std::memory_order_relaxed));

        PendingUpdate resourceUp;
        resourceUp.kind                    = PendingUpdate::Kind::ResourceSummary;
        resourceUp.generation              = generation;
        resourceUp.acceptedEntries         = acceptedEntries;
        resourceUp.rejectedEntries         = rejectedEntries;
        resourceUp.cappedDirectories       = cappedDirectories;
        resourceUp.cappedFiles             = cappedFiles;
        resourceUp.retainedNameBytes       = retainedNameBytes;
        resourceUp.retainedChildReferences = retainedChildRefs;
        resourceUp.traversedDirectories    = traversedDirs;
        resourceUp.validationError         = validationError;
        PostUpdate(std::move(resourceUp));

        Debug::Perf::EmitValue(L"viewer.space.model.accepted_entries", acceptedEntries, scanFinalHr);
        Debug::Perf::Emit(L"viewer.space.model.rejected_entries", ValidationErrorDetail(validationError), 0u, rejectedEntries, 0u, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.model.capped_directories", cappedDirectories, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.model.capped_files", cappedFiles, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.model.retained_directories", retainedDirectories, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.model.retained_files", retainedFiles, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.model.child_references", retainedChildRefs, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.model.retained_name_bytes", retainedNameBytes, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.model.traversed_directories", traversedDirs, scanFinalHr);
    });

    struct ProgressState final
    {
        std::chrono::steady_clock::time_point lastProgress = std::chrono::steady_clock::now();
        uint32_t lastProgressNodeId                        = 0;
    };

    const wchar_t pathSeparator = DeterminePreferredPathSeparator(rootPath, fileSystemIsWin32);

    auto leafNameFromPath = [&](std::wstring_view path) noexcept -> std::wstring
    {
        if (path.empty())
        {
            return {};
        }

        const std::wstring_view trimmed = TrimTrailingPathSeparators(path);
        if (trimmed.size() == 2 && IsAsciiAlpha(trimmed[0]) && trimmed[1] == L':' && path.size() >= 3)
        {
            return std::wstring(path);
        }

        const size_t lastSep = trimmed.find_last_of(L"/\\");
        if (lastSep != std::wstring_view::npos && lastSep + 1 < trimmed.size())
        {
            return std::wstring(trimmed.substr(lastSep + 1));
        }

        return std::wstring(trimmed);
    };

    auto postProgress = [&](ProgressState& progress, uint32_t nodeId, std::wstring_view currentPath, bool force = false) noexcept
    {
        const auto now = std::chrono::steady_clock::now();
        if (! force && nodeId == progress.lastProgressNodeId && (now - progress.lastProgress) < kProgressUpdateInterval)
        {
            return;
        }
        progress.lastProgress       = now;
        progress.lastProgressNodeId = nodeId;

        PendingUpdate up;
        up.kind                     = PendingUpdate::Kind::Progress;
        up.generation               = generation;
        up.nodeId                   = nodeId;
        up.bytes                    = scannedBytes.load(std::memory_order_relaxed);
        up.scannedFolders           = scannedFolders.load(std::memory_order_relaxed);
        up.scannedFiles             = scannedFiles.load(std::memory_order_relaxed);
        up.failedFolders            = failedFolders.load(std::memory_order_relaxed);
        up.firstPreviewPriorityJobs = firstPreviewPriorityJobs.load(std::memory_order_relaxed);
        up.name                     = leafNameFromPath(currentPath);
        PostUpdate(std::move(up));
    };

    PendingUpdate queuedUp;
    queuedUp.kind       = PendingUpdate::Kind::UpdateState;
    queuedUp.generation = generation;
    queuedUp.nodeId     = rootNodeId;
    queuedUp.state      = ScanState::Queued;
    PostUpdate(std::move(queuedUp));

    if (! fileSystem)
    {
        scanFinalState = ScanState::Error;
        scanFinalHr    = E_POINTER;
        PendingUpdate errorUp;
        errorUp.kind       = PendingUpdate::Kind::UpdateState;
        errorUp.generation = generation;
        errorUp.nodeId     = rootNodeId;
        errorUp.state      = ScanState::Error;
        PostUpdate(std::move(errorUp));
        return;
    }
    if (rootPath.empty() || rootPath.size() > resourcePolicy.maxPathChars)
    {
        resources.RecordValidationFailure(ViewerSpaceScan::ValidationError::PathLimit);
        scanFinalState = ScanState::Error;
        scanFinalHr    = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        PendingUpdate errorUp;
        errorUp.kind       = PendingUpdate::Kind::UpdateState;
        errorUp.generation = generation;
        errorUp.nodeId     = rootNodeId;
        errorUp.state      = ScanState::Error;
        PostUpdate(std::move(errorUp));
        return;
    }

    const std::shared_ptr<ScanScheduler> scheduler = GetScanScheduler();
    if (! scheduler)
    {
        scanFinalState = ScanState::Canceled;
        scanFinalHr    = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        return;
    }
    if (! fileSystemIsWin32)
    {
        // Query provider path semantics once on the existing contained worker. Input, paint,
        // exclusion toggles and refresh rebinding only consume this copied typed contract.
        // The shared validator requires the exact metadata provider ID, including its
        // namespace. A display name or short ID is not authority for path equivalence.
        wil::com_ptr<IInformations> information;
        const PluginMetaData* metadata = nullptr;
        std::wstring providerId;
        if (fileSystem.try_query_to(information.put()) && information && SUCCEEDED(information->GetMetaData(&metadata)) && metadata && metadata->id)
        {
            const size_t length = wcsnlen_s(metadata->id, 4097u);
            if (length && length <= 4096u)
                providerId.assign(metadata->id, length);
        }
        auto route = FileSystemRouteContract::Query(fileSystem.get(), rootPath, FILESYSTEM_COPY, providerId);
        if (route.state == FileSystemRouteContract::QueryState::Available && route.snapshot.pathIdentity.has_value() &&
            route.snapshot.pathIdentity.value().pathTextStableIdentity)
        {
            PendingUpdate identity;
            identity.kind         = PendingUpdate::Kind::PathIdentity;
            identity.generation   = generation;
            identity.pathIdentity = std::move(route.snapshot.pathIdentity);
            PostUpdate(std::move(identity));
        }
    }
    ScanScheduler::Permit permit;
    if (fileSystemIsWin32)
    {
        permit = scheduler->AcquireForPath(std::filesystem::path(rootPath), stopToken);
    }
    else
    {
        wchar_t key[32]{};
        constexpr size_t keyMax                  = (sizeof(key) / sizeof(key[0])) - 1;
        const auto r                             = std::format_to_n(key, keyMax, L"fs:{:p}", static_cast<const void*>(fileSystem.get()));
        key[(r.size < keyMax) ? r.size : keyMax] = L'\0';
        permit                                   = scheduler->AcquireForKey(key, stopToken);
    }
    if (! permit)
    {
        scanFinalState = ScanState::Canceled;
        scanFinalHr    = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        PendingUpdate canceled;
        canceled.kind       = PendingUpdate::Kind::UpdateState;
        canceled.generation = generation;
        canceled.nodeId     = rootNodeId;
        canceled.state      = ScanState::Canceled;
        PostUpdate(std::move(canceled));
        return;
    }

    PendingUpdate scanningUp;
    scanningUp.kind       = PendingUpdate::Kind::UpdateState;
    scanningUp.generation = generation;
    scanningUp.nodeId     = rootNodeId;
    scanningUp.state      = ScanState::Scanning;
    PostUpdate(std::move(scanningUp));

    struct ChildDir final
    {
        uint64_t traversalId    = 0u;
        uint32_t modelNodeId    = 0u;
        uint32_t aggregateOwner = 0u;
        bool aggregateRoot      = false;
        bool descend            = true;
        std::wstring name;
        int64_t lastWriteTime = 0;
    };

    struct StackItem final
    {
        uint64_t traversalId    = 0u;
        uint32_t modelNodeId    = 0u;
        uint32_t aggregateOwner = 0u;
        bool aggregateRoot      = false;
        uint32_t depth          = 0u;
        std::wstring path;
        uint64_t bytes          = 0;
        bool enumerated         = false;
        bool failed             = false;
        bool propagateFailure   = true;
        bool contentNotMeasured = false;
        size_t processedEntries = 0;
        std::vector<ChildDir> children;
        std::vector<FileSummaryItem> topFiles;
        uint64_t stagingReservationBytes = 0u;
        uint64_t fileCount               = 0u;
        uint64_t nonDescendingFolders    = 0u;
        uint64_t otherBytes              = 0u;
        uint64_t otherFolders            = 0u;
        uint64_t otherFiles              = 0u;
    };

    auto minHeapByBytes = [](const FileSummaryItem& a, const FileSummaryItem& b) noexcept { return a.bytes > b.bytes; };

    auto enumerate = [&](StackItem& item, ProgressState& progress, uint32_t workerIndex) noexcept -> void
    {
        const auto enumerateStartedAt = std::chrono::steady_clock::now();
        HRESULT enumerateHr           = S_OK;
        uint64_t candidateNameCopies  = 0u;
        auto emitEnumerateMetric      = wil::scope_exit([&]() noexcept
        {
            Debug::Perf::Emit(L"viewer.space.scan.enumerate_us",
                              L"bounded-provider",
                              Debug::Perf::ElapsedUs(enumerateStartedAt),
                              item.processedEntries,
                              workerIndex,
                              enumerateHr);
            if (candidateNameCopies < item.fileCount)
            {
                Debug::Perf::EmitValue(L"viewer.space.scan.candidate_name_copies", candidateNameCopies, enumerateHr);
            }
        });

        const uint32_t progressNodeId = item.modelNodeId != 0u ? item.modelNodeId : item.aggregateOwner;
        auto failProvider             = [&](ViewerSpaceScan::ValidationError error, HRESULT hr, bool propagateFailure) noexcept
        {
            enumerateHr           = hr;
            item.failed           = true;
            item.propagateFailure = propagateFailure;
            resources.RecordValidationFailure(error);
            Debug::Warning(
                L"ViewerSpace: Rejected provider directory '{}' ({}, HRESULT: {:#x})", item.path, ValidationErrorDetail(error), static_cast<uint32_t>(hr));
            if (item.modelNodeId != 0u)
            {
                PendingUpdate errorUp;
                errorUp.kind       = PendingUpdate::Kind::UpdateState;
                errorUp.generation = generation;
                errorUp.nodeId     = item.modelNodeId;
                errorUp.state      = ScanState::Error;
                PostUpdate(std::move(errorUp));
            }
        };

        item.enumerated = true;
        // A provider can block inside ReadDirectoryInfo. Publish the active read
        // before entering it so the linked row does not stay misleadingly Queued.
        if (item.modelNodeId != 0u)
        {
            PendingUpdate stateUp;
            stateUp.kind       = PendingUpdate::Kind::UpdateState;
            stateUp.generation = generation;
            stateUp.nodeId     = item.modelNodeId;
            stateUp.state      = ScanState::Scanning;
            PostUpdate(std::move(stateUp));
        }
        postProgress(progress, progressNodeId, item.path);

        // Child timestamps arrive in their parent's listing. Only the scan root needs an optional
        // metadata read, on the scan worker; selection and hover consume retained values exclusively.
        if (item.modelNodeId == rootNodeId && ! stopToken.stop_requested())
        {
            wil::com_ptr<IFileSystemIO> io;
            FileSystemBasicInformation info{};
            info.sizeBytes = sizeof(info);
            if (SUCCEEDED(fileSystem->QueryInterface(__uuidof(IFileSystemIO), io.put_void())) && io &&
                SUCCEEDED(io->GetFileBasicInformation(item.path.c_str(), &info)) && info.sizeBytes == sizeof(info))
            {
                PendingUpdate metadata;
                metadata.kind          = PendingUpdate::Kind::UpdateMetadata;
                metadata.generation    = generation;
                metadata.nodeId        = rootNodeId;
                metadata.lastWriteTime = info.lastWriteTime;
                PostUpdate(std::move(metadata));
            }
        }

        wil::com_ptr<IFilesInformation> filesInformation;
#ifdef ENABLE_TESTS
        HRESULT enumHr = E_FAIL;
        {
            ScanPhaseTimer phaseTimer(captureScanPhases, providerReadPhase);
            enumHr = fileSystem->ReadDirectoryInfo(item.path.c_str(), filesInformation.put());
        }
#else
        const HRESULT enumHr = fileSystem->ReadDirectoryInfo(item.path.c_str(), filesInformation.put());
#endif
        enumerateHr = enumHr;
        if (FAILED(enumHr) || ! filesInformation)
        {
            const HRESULT failureHr = FAILED(enumHr) ? enumHr : E_FAIL;
            const bool accessDenied = failureHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            if (accessDenied)
            {
                // Access-denied directories are a normal zero-sized leaf in a best-effort space scan.
                item.contentNotMeasured = true;
                return;
            }
            // A folder that cannot be read is an Error in its own subtree only: its ancestors and the root still complete as Done,
            // and the count is reported. Malformed provider buffers and scan-wide ceilings keep propagating (they remain whole-scan errors).
            SaturatingAtomicAdd(failedFolders, 1u);
            failProvider(ViewerSpaceScan::ValidationError::NullBuffer, failureHr, false);
            return;
        }

        FileInfo* buffer              = nullptr;
        unsigned long bufferSize      = 0u;
        unsigned long allocatedSize   = 0u;
        unsigned long advertisedCount = 0u;
        const HRESULT bufferHr        = filesInformation->GetBuffer(&buffer);
        const HRESULT sizeHr          = filesInformation->GetBufferSize(&bufferSize);
        const HRESULT allocatedHr     = filesInformation->GetAllocatedSize(&allocatedSize);
        const HRESULT countHr         = filesInformation->GetCount(&advertisedCount);
        if (FAILED(bufferHr) || FAILED(sizeHr) || FAILED(allocatedHr) || FAILED(countHr))
        {
            const HRESULT failingHr = FAILED(bufferHr) ? bufferHr : (FAILED(sizeHr) ? sizeHr : (FAILED(allocatedHr) ? allocatedHr : countHr));
            failProvider(ViewerSpaceScan::ValidationError::CountBufferMismatch, failingHr, true);
            return;
        }

        // This is the provider-reported opaque allocation. It is measured separately and is not debited from our local validation-scratch budget.
        const uint64_t activeProviderBytes = SaturatingAtomicAdd(_providerBufferCurrentBytes, allocatedSize);
        UpdateAtomicPeak(_providerBufferPeakBytes, activeProviderBytes);
        auto releaseProviderBufferAccounting =
            wil::scope_exit([&]() noexcept { static_cast<void>(SubtractAtomicFloor(_providerBufferCurrentBytes, allocatedSize)); });

        const size_t fileCandidateLimit = std::min(topFilesPerDirectory, static_cast<size_t>(resourcePolicy.maxRetainedFileRecords));
        const auto* firstByte           = reinterpret_cast<const std::byte*>(buffer);
        const std::span<const std::byte> providerBuffer(firstByte, static_cast<size_t>(bufferSize));
#ifdef ENABLE_TESTS
        ScanPhaseTimer validationPhaseTimer(captureScanPhases, listingValidateAdmitPhase);
#endif
        const ViewerSpaceScan::ValidationError bufferContractError =
            ViewerSpaceScan::ValidateProviderBufferContract(buffer, providerBuffer.size(), allocatedSize, advertisedCount, resourcePolicy);
        if (bufferContractError != ViewerSpaceScan::ValidationError::None)
        {
#ifdef ENABLE_TESTS
            validationPhaseTimer.Stop();
#endif
            failProvider(bufferContractError, HRESULT_FROM_WIN32(ERROR_INVALID_DATA), true);
            return;
        }
        // Conservative admission covers offsets, name-order indexes, duplicate-ID storage, and both merge-sort scratch buffers.
        constexpr uint64_t kValidationScratchBytesPerPossibleEntry = 64u;
        constexpr uint64_t kValidationScratchFixedBytes            = 64u * 1024u;
        const size_t recordHeaderBytes                             = offsetof(FileInfo, FileName);
        const size_t maximumPossibleEntries =
            recordHeaderBytes == 0u ? 0u : std::min(static_cast<size_t>(advertisedCount), providerBuffer.size() / recordHeaderBytes);
        const uint64_t estimatedScratchBytes =
            ViewerSpaceScan::SaturatingAdd(kValidationScratchFixedBytes,
                                           maximumPossibleEntries > ((std::numeric_limits<uint64_t>::max)() / kValidationScratchBytesPerPossibleEntry)
                                               ? (std::numeric_limits<uint64_t>::max)()
                                               : static_cast<uint64_t>(maximumPossibleEntries) * kValidationScratchBytesPerPossibleEntry);
        const ValidationScratchBudget::Admission scratchAdmission = validationScratchBudget.Acquire(estimatedScratchBytes, stopToken);
        if (scratchAdmission == ValidationScratchBudget::Admission::Canceled)
        {
            return;
        }
        if (scratchAdmission == ValidationScratchBudget::Admission::TooLarge)
        {
#ifdef ENABLE_TESTS
            validationPhaseTimer.Stop();
#endif
            _validationScratchRejectedCount.fetch_add(1u, std::memory_order_relaxed);
            failProvider(ViewerSpaceScan::ValidationError::OutstandingLimit, HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY), true);
            return;
        }
        auto releaseValidationScratch = wil::scope_exit([&]() noexcept { validationScratchBudget.Release(estimatedScratchBytes); });
#ifdef ENABLE_TESTS
        validationPhaseTimer.Pause(); // Exclude the deliberate test gate, while retaining real scratch-budget admission time.
        if (_debugPauseAfterValidationScratch.exchange(false, std::memory_order_acq_rel))
        {
            const HANDLE enteredEvent = _debugValidationScratchEnteredEvent.load(std::memory_order_relaxed);
            const HANDLE releaseEvent = _debugValidationScratchReleaseEvent.load(std::memory_order_relaxed);
            if (enteredEvent && releaseEvent)
            {
                static_cast<void>(SetEvent(enteredEvent));
                static_cast<void>(WaitForSingleObject(releaseEvent, 10000u));
            }
        }
        validationPhaseTimer.Resume();
#endif
        if (stopToken.stop_requested())
        {
            return;
        }
        ViewerSpaceScan::ValidatedListing validatedListing =
            ViewerSpaceScan::ValidateProviderListing(providerBuffer, allocatedSize, advertisedCount, resourcePolicy, fileSystemIsWin32, stopToken);
#ifdef ENABLE_TESTS
        validationPhaseTimer.Stop();
#endif
        if (validatedListing.canceled)
        {
            return;
        }
        if (validatedListing.error != ViewerSpaceScan::ValidationError::None)
        {
            failProvider(validatedListing.error, HRESULT_FROM_WIN32(ERROR_INVALID_DATA), true);
            return;
        }

        // Reserve before any owning name copy. The validated provider extent and entry ceiling bound all
        // arithmetic below well inside uint64_t. This constant-time estimate charges every directory name,
        // the per-directory top-K heap plus old/new replacement names, and conservative vector/summary
        // metadata. Charging every discarded file would reject supported large flat directories with a
        // small top-K. This lease remains through processJob, including moves into unposted chunks.
        const auto stagingPreflightStartedAt = std::chrono::steady_clock::now();
        // Validation gathered these byte-safe shape statistics while decoding the provider records, avoiding a second walk.
        const uint64_t directoryCount       = validatedListing.directoryCount;
        const uint64_t directoryNameBytes   = validatedListing.directoryNameBytes;
        const uint64_t fileCount            = validatedListing.fileCount;
        const uint64_t maximumFileNameBytes = validatedListing.maximumFileNameBytes;
        const uint64_t candidateCount       = item.modelNodeId != 0u ? std::min<uint64_t>(fileCount, fileCandidateLimit) : 0u;
        const uint64_t stagingBytes =
            64u * 1024u + 2u * (directoryNameBytes + (candidateCount + 2u) * maximumFileNameBytes) + 512u * (directoryCount + candidateCount + 2u);
        const uint64_t stagingPreflightUs = Debug::Perf::ElapsedUs(stagingPreflightStartedAt);
        SaturatingAtomicAdd(stagingPreflightCount, 1u);
        SaturatingAtomicAdd(stagingPreflightTotalUs, stagingPreflightUs);
        UpdateAtomicPeak(stagingPreflightMaxUs, stagingPreflightUs);
        UpdateAtomicPeak(stagingPreflightReservationPeakBytes, stagingBytes);
        SaturatingAtomicAdd(stagingPreflightRecordTotal, static_cast<uint64_t>(validatedListing.entryOffsets.size()));
#ifdef ENABLE_TESTS
        ValidationScratchBudget::Admission stagingAdmission = ValidationScratchBudget::Admission::Canceled;
        {
            ScanPhaseTimer phaseTimer(captureScanPhases, workerStagingAdmitPhase);
            stagingAdmission = workerStagingBudget.Acquire(stagingBytes, stopToken);
        }
#else
        const auto stagingAdmission = workerStagingBudget.Acquire(stagingBytes, stopToken);
#endif
        if (stagingAdmission == ValidationScratchBudget::Admission::Canceled)
            return;
        if (stagingAdmission == ValidationScratchBudget::Admission::TooLarge)
        {
            _workerStagingRejectedCount.fetch_add(1u, std::memory_order_relaxed);
            failProvider(ViewerSpaceScan::ValidationError::OutstandingLimit, HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY), true);
            return;
        }
        item.stagingReservationBytes = stagingBytes;

        uint64_t localFolderCount = 0u;
        uint64_t localFileCount   = 0u;
#ifdef ENABLE_TESTS
        ScanPhaseTimer candidateBuildTimer(captureScanPhases, candidateBuildPhase);
#endif
        if (! validatedListing.entryOffsets.empty())
        {
            uint64_t parsed = 0u;

            for (const uint32_t recordOffset : validatedListing.entryOffsets)
            {
                if (stopToken.stop_requested())
                {
                    return;
                }

                const auto* entryPointer = reinterpret_cast<const FileInfo*>(providerBuffer.data() + static_cast<size_t>(recordOffset));
                const FileInfo& entry    = *entryPointer;
                const std::wstring_view name(entry.FileName, static_cast<size_t>(entry.FileNameSize / sizeof(wchar_t)));
                parsed = ViewerSpaceScan::SaturatingAdd(parsed, uint64_t{1u});
                item.processedEntries += 1u;
                const bool isDirectory = (entry.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
                const bool isReparse   = (entry.FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u;
                if (isDirectory)
                {
                    localFolderCount = ViewerSpaceScan::SaturatingAdd(localFolderCount, uint64_t{1u});
                    ChildDir childDir;
                    childDir.descend = ! isReparse;
                    childDir.name.assign(name);
                    childDir.lastWriteTime = entry.LastWriteTime;
                    if (! childDir.descend)
                    {
                        item.nonDescendingFolders = ViewerSpaceScan::SaturatingAdd(item.nonDescendingFolders, uint64_t{1u});
                    }
                    item.children.push_back(std::move(childDir));
                }
                else
                {
                    localFileCount           = ViewerSpaceScan::SaturatingAdd(localFileCount, uint64_t{1u});
                    const uint64_t fileBytes = entry.EndOfFile > 0 ? static_cast<uint64_t>(entry.EndOfFile) : 0u;
                    item.bytes               = ViewerSpaceScan::SaturatingAdd(item.bytes, fileBytes);

                    auto aggregateFile = [&](uint64_t bytes, bool alreadyCounted) noexcept
                    {
                        item.otherBytes = ViewerSpaceScan::SaturatingAdd(item.otherBytes, bytes);
                        item.otherFiles = ViewerSpaceScan::SaturatingAdd(item.otherFiles, uint64_t{1u});
                        if (! alreadyCounted)
                        {
                            SaturatingAtomicAdd(resources.cappedFiles, 1u);
                        }
                    };

                    if (item.modelNodeId == 0u || fileCandidateLimit == 0u)
                    {
                        aggregateFile(fileBytes, false);
                    }
                    else if (item.topFiles.size() < fileCandidateLimit)
                    {
                        FileSummaryItem candidate;
                        candidate.bytes         = fileBytes;
                        candidate.lastWriteTime = entry.LastWriteTime;
                        candidate.name.assign(name);
                        candidateNameCopies = ViewerSpaceScan::SaturatingAdd(candidateNameCopies, uint64_t{1u});
                        item.topFiles.push_back(std::move(candidate));
                        std::push_heap(item.topFiles.begin(), item.topFiles.end(), minHeapByBytes);
                    }
                    else if (! item.topFiles.empty() && fileBytes > item.topFiles.front().bytes)
                    {
                        std::pop_heap(item.topFiles.begin(), item.topFiles.end(), minHeapByBytes);
                        FileSummaryItem dropped = std::move(item.topFiles.back());
                        FileSummaryItem candidate;
                        candidate.bytes         = fileBytes;
                        candidate.lastWriteTime = entry.LastWriteTime;
                        candidate.name.assign(name);
                        candidateNameCopies  = ViewerSpaceScan::SaturatingAdd(candidateNameCopies, uint64_t{1u});
                        item.topFiles.back() = std::move(candidate);
                        std::push_heap(item.topFiles.begin(), item.topFiles.end(), minHeapByBytes);
                        aggregateFile(dropped.bytes, false);
                    }
                    else
                    {
                        aggregateFile(fileBytes, false);
                    }

                    if (item.modelNodeId != 0u && (item.processedEntries % kProgressUpdateStride) == 0u)
                    {
                        PendingUpdate progressUp;
                        progressUp.kind       = PendingUpdate::Kind::UpdateSize;
                        progressUp.generation = generation;
                        progressUp.nodeId     = item.modelNodeId;
                        progressUp.bytes      = item.bytes;
                        PostUpdate(std::move(progressUp));
                        postProgress(progress, progressNodeId, item.path);
                    }
                }
            }

            if (parsed != advertisedCount)
            {
                item.children.clear();
                item.topFiles.clear();
                item.bytes      = 0u;
                item.otherBytes = 0u;
                item.otherFiles = 0u;
#ifdef ENABLE_TESTS
                candidateBuildTimer.Stop();
#endif
                failProvider(ViewerSpaceScan::ValidationError::CountMismatch, HRESULT_FROM_WIN32(ERROR_INVALID_DATA), true);
                return;
            }

            SaturatingAtomicAdd(resources.acceptedEntries, parsed);
        }

        item.fileCount = localFileCount;
        SaturatingAtomicAdd(scannedFolders, localFolderCount);
        SaturatingAtomicAdd(scannedFiles, localFileCount);
        SaturatingAtomicAdd(scannedBytes, item.bytes);
#ifdef ENABLE_TESTS
        candidateBuildTimer.Stop();
#endif

        postProgress(progress, progressNodeId, item.path, true);
        std::vector<uint32_t>().swap(validatedListing.entryOffsets);
        validationScratchBudget.Release(estimatedScratchBytes);
        releaseValidationScratch.release();
#ifdef ENABLE_TESTS
        if (_debugPauseAfterWorkerStaging.exchange(false, std::memory_order_acq_rel))
        {
            const HANDLE entered = _debugWorkerStagingEnteredEvent.load(std::memory_order_relaxed);
            const HANDLE release = _debugWorkerStagingReleaseEvent.load(std::memory_order_relaxed);
            if (entered && release)
            {
                static_cast<void>(SetEvent(entered));
                static_cast<void>(WaitForSingleObject(release, 10000u));
            }
        }
#endif
    };

    struct ScanJob final
    {
        uint64_t traversalId       = 0u;
        uint64_t parentTraversalId = 0u;
        uint32_t modelNodeId       = 0u;
        uint32_t aggregateOwner    = 0u;
        bool aggregateRoot         = false;
        bool pathBudgetLease       = false;
        uint32_t depth             = 0u;
        uint64_t pathBytes         = 0u;
        bool pathBudgetFailed      = false;
        std::wstring path;
    };

    struct DirectoryCompletion final
    {
        uint64_t parentTraversalId = 0u;
        uint32_t modelNodeId       = 0u;
        uint32_t aggregateOwner    = 0u;
        uint64_t bytes             = 0u;
        uint64_t folderCount       = 0u;
        uint64_t fileCount         = 0u;
        uint32_t pendingChildren   = 0u;
        bool aggregateRoot         = false;
        bool enumerated            = false;
        bool completed             = false;
        bool displayError          = false;
        bool failed                = false;
        bool contentNotMeasured    = false;
    };

    std::mutex frontierMutex;
    std::condition_variable frontierCv;
    ViewerSpaceDetail::ConcurrentCountingMemoryResource scannerBookkeeping(_scannerBookkeepingCurrentBytes, _scannerBookkeepingPeakBytes);
    std::pmr::deque<ScanJob> frontier(&scannerBookkeeping);
    std::pmr::unordered_map<uint64_t, DirectoryCompletion> completions(&scannerBookkeeping);
    completions.reserve(std::min<size_t>(resourcePolicy.maxOutstandingDirectories, 4096u));

    const uint64_t rootPathBytes = (static_cast<uint64_t>(rootPath.capacity()) + 1u) * sizeof(wchar_t);
    DirectoryCompletion rootCompletion;
    rootCompletion.modelNodeId    = rootNodeId;
    rootCompletion.aggregateOwner = rootNodeId;
    completions.emplace(1u, rootCompletion);

    const uint64_t activePathBudgetBytes = ActivePathBudgetBytes(resourcePolicy);
    ScanJob rootJob;
    rootJob.traversalId            = 1u;
    rootJob.modelNodeId            = rootNodeId;
    rootJob.aggregateOwner         = rootNodeId;
    rootJob.pathBytes              = rootPathBytes;
    uint64_t admittedRootPathBytes = 0u;
    rootJob.pathBudgetLease =
        rootPathBytes != 0u && TryReserveAtomicBudget(_activePathCurrentBytes, rootPathBytes, activePathBudgetBytes, admittedRootPathBytes);
    rootJob.pathBudgetFailed = rootPathBytes > activePathBudgetBytes || (rootPathBytes != 0u && ! rootJob.pathBudgetLease);
    if (rootJob.pathBudgetFailed)
    {
        _activePathRejectedCount.fetch_add(1u, std::memory_order_relaxed);
    }
    rootJob.path = std::move(rootPath);
    frontier.push_back(std::move(rootJob));
    uint64_t livePathBytes     = rootPathBytes;
    uint64_t reservedPathBytes = 0u;
    if (rootPathBytes != 0u && ! frontier.front().pathBudgetFailed)
    {
        UpdateAtomicPeak(_activePathPeakBytes, admittedRootPathBytes);
    }

    bool frontierFinished = false;
    std::atomic_bool rootFailed(false);
    std::atomic_uint32_t activeWorkers(0);
    std::atomic_uint32_t idleWorkerWaits(0);
    std::atomic_uint32_t frontierDepthPeak(1u);

    auto updateFrontierDepthPeak = [&](uint32_t depth) noexcept
    {
        uint32_t observed = frontierDepthPeak.load(std::memory_order_relaxed);
        while (depth > observed && ! frontierDepthPeak.compare_exchange_weak(observed, depth, std::memory_order_relaxed))
        {
        }
    };

    auto postDirectoryCompletion = [&](uint32_t modelNodeId, uint32_t aggregateOwner, uint64_t bytes, ScanState state, bool contentNotMeasured) noexcept
    {
        if (modelNodeId == 0u && (! contentNotMeasured || aggregateOwner == 0u))
        {
            return;
        }
        if (modelNodeId != 0u)
        {
            PendingUpdate sizeUp;
            sizeUp.kind       = PendingUpdate::Kind::UpdateSize;
            sizeUp.generation = generation;
            sizeUp.nodeId     = modelNodeId;
            sizeUp.bytes      = bytes;
            PostUpdate(std::move(sizeUp));
        }

        PendingUpdate doneUp;
        doneUp.kind               = PendingUpdate::Kind::UpdateState;
        doneUp.generation         = generation;
        doneUp.nodeId             = modelNodeId != 0u ? modelNodeId : aggregateOwner;
        doneUp.state              = modelNodeId != 0u ? state : ScanState::Scanning;
        doneUp.contentNotMeasured = contentNotMeasured;
        PostUpdate(std::move(doneUp));
    };

    auto appendDirectoryFilesSummary = [&](std::vector<PendingUpdate>& updates,
                                           uint32_t nodeId,
                                           uint64_t revision,
                                           uint64_t otherBytes,
                                           uint64_t otherFolders,
                                           uint64_t otherFiles,
                                           std::vector<FileSummaryItem>& files) noexcept -> bool
    {
        static constexpr size_t kOrdinaryChunkBytes   = 64u * 1024u;
        static constexpr size_t kSingleNameChunkBytes = 128u * 1024u;
        static constexpr size_t kMaxFilesPerChunk     = 64u;

        uint64_t nameBytes = 0u;
        for (const FileSummaryItem& file : files)
        {
            nameBytes = ViewerSpaceScan::SaturatingAdd(nameBytes, static_cast<uint64_t>(file.name.size()) * sizeof(wchar_t));
        }

        auto initialize = [&](PendingUpdate& update, PendingUpdate::Kind kind) noexcept
        {
            update.kind                 = kind;
            update.generation           = generation;
            update.nodeId               = nodeId;
            update.summaryRevision      = revision;
            update.summaryExpectedFiles = static_cast<uint32_t>(files.size());
            update.summaryNameBytes     = nameBytes;
        };

        std::vector<PendingUpdate> chunks;
        chunks.reserve((files.size() + kMaxFilesPerChunk - 1u) / kMaxFilesPerChunk);
        PendingUpdate chunk;
        initialize(chunk, PendingUpdate::Kind::DirectorySummaryChunk);
        chunk.topFiles.reserve(kMaxFilesPerChunk);
        auto flushChunk = [&]() noexcept
        {
            if (chunk.topFiles.empty())
            {
                return;
            }
            chunks.push_back(std::move(chunk));
            chunk = {};
            initialize(chunk, PendingUpdate::Kind::DirectorySummaryChunk);
            chunk.topFiles.reserve(kMaxFilesPerChunk);
        };

        for (FileSummaryItem& file : files)
        {
            if (chunk.topFiles.size() == kMaxFilesPerChunk)
            {
                flushChunk();
            }
            chunk.topFiles.push_back(std::move(file));
            const size_t chunkBytes = EstimatePendingUpdateBytes(chunk);
            if (chunkBytes > kOrdinaryChunkBytes)
            {
                FileSummaryItem oversized = std::move(chunk.topFiles.back());
                chunk.topFiles.pop_back();
                if (chunk.topFiles.empty())
                {
                    PendingUpdate rejected;
                    initialize(rejected, PendingUpdate::Kind::DirectorySummaryReject);
                    updates.push_back(std::move(rejected));
                    return false;
                }
                flushChunk();
                chunk.topFiles.push_back(std::move(oversized));
                const size_t oneNameBytes = EstimatePendingUpdateBytes(chunk);
                if (oneNameBytes > kSingleNameChunkBytes || chunk.topFiles.size() != 1u)
                {
                    PendingUpdate rejected;
                    initialize(rejected, PendingUpdate::Kind::DirectorySummaryReject);
                    updates.push_back(std::move(rejected));
                    return false;
                }
                flushChunk();
            }
        }
        flushChunk();
        if (chunks.size() > (std::numeric_limits<uint32_t>::max)())
        {
            PendingUpdate rejected;
            initialize(rejected, PendingUpdate::Kind::DirectorySummaryReject);
            updates.push_back(std::move(rejected));
            return false;
        }

        bool omitEndForTest = false;
#ifdef ENABLE_TESTS
        bool duplicateChunkForTest = false;
        bool outOfOrderForTest     = false;
        if (nodeId == rootNodeId && files.size() > 1u && ! summaryFaultInjected.exchange(true, std::memory_order_acq_rel))
        {
            if (ReadPositiveUint32Environment(kTestSummaryFaultMissingEndEnvVar).has_value())
            {
                omitEndForTest = true;
            }
            else if (ReadPositiveUint32Environment(kTestSummaryFaultDuplicateChunkEnvVar).has_value() && ! chunks.empty())
            {
                duplicateChunkForTest = true;
            }
            else if (ReadPositiveUint32Environment(kTestSummaryFaultOutOfOrderEnvVar).has_value() && chunks.size() > 1u)
            {
                outOfOrderForTest = true;
            }
        }
#endif

        PendingUpdate begin;
        initialize(begin, PendingUpdate::Kind::DirectorySummaryBegin);
        begin.otherBytes            = otherBytes;
        begin.otherFolders          = otherFolders;
        begin.otherFiles            = otherFiles;
        begin.summaryExpectedChunks = static_cast<uint32_t>(chunks.size());
        updates.push_back(std::move(begin));
#ifdef ENABLE_TESTS
        const size_t firstChunkOffset = updates.size();
#endif
        for (size_t index = 0u; index < chunks.size(); ++index)
        {
            PendingUpdate& next        = chunks[index];
            next.summarySequence       = static_cast<uint32_t>(index);
            next.summaryExpectedChunks = static_cast<uint32_t>(chunks.size());
            updates.push_back(std::move(next));
        }
#ifdef ENABLE_TESTS
        // Mutate the sequenced packets: reordering before sequence assignment would remain valid traffic.
        if (duplicateChunkForTest)
        {
            PendingUpdate duplicate = updates[firstChunkOffset];
            updates.insert(updates.begin() + static_cast<ptrdiff_t>(firstChunkOffset + 1u), std::move(duplicate));
            Debug::Perf::EmitCounter(L"viewer.space.scan.test.summary_duplicate_chunk_fault_applied");
        }
        if (outOfOrderForTest)
        {
            std::swap(updates[firstChunkOffset], updates[firstChunkOffset + 1u]);
            Debug::Perf::EmitCounter(L"viewer.space.scan.test.summary_out_of_order_fault_applied");
        }
        if (omitEndForTest)
        {
            Debug::Perf::EmitCounter(L"viewer.space.scan.test.summary_missing_end_fault_applied");
        }
#endif
        PendingUpdate end;
        initialize(end, PendingUpdate::Kind::DirectorySummaryEnd);
        end.summaryExpectedChunks = static_cast<uint32_t>(chunks.size());
        end.summarySequence       = static_cast<uint32_t>(chunks.size());
        if (! omitEndForTest)
        {
            updates.push_back(std::move(end));
        }
        return true;
    };

#ifdef ENABLE_TESTS
    auto tryRetainDirectoryProfiled = [&](uint64_t nameBytes) noexcept
    {
        if (! captureScanPhases)
        {
            return resources.TryRetainDirectory(nameBytes);
        }
        ScanPhaseTimer phaseTimer(true, modelRecordRetainPhase);
        return resources.TryRetainDirectory(nameBytes);
    };
    auto tryRetainFileProfiled = [&](uint64_t nameBytes) noexcept
    {
        if (! captureScanPhases)
        {
            return resources.TryRetainFile(nameBytes);
        }
        ScanPhaseTimer phaseTimer(true, modelRecordRetainPhase);
        return resources.TryRetainFile(nameBytes);
    };
#endif

    auto completeReadyDirectories = [&](uint64_t initialTraversalId) noexcept -> void
    {
        uint64_t traversalId = initialTraversalId;
        for (;;)
        {
            uint64_t parentTraversalId = 0u;
            uint32_t modelNodeId       = 0u;
            uint32_t aggregateOwner    = 0u;
            uint64_t bytes             = 0u;
            uint64_t folderCount       = 0u;
            uint64_t fileCount         = 0u;
            bool aggregateRoot         = false;
            bool displayError          = false;
            bool failed                = false;
            bool contentNotMeasured    = false;
            bool completedRoot         = false;

            {
                std::scoped_lock lock(frontierMutex);
                auto it = completions.find(traversalId);
                if (it == completions.end())
                {
                    return;
                }

                DirectoryCompletion& completion = it->second;
                if (! completion.enumerated || completion.pendingChildren != 0u || completion.completed)
                {
                    return;
                }

                completion.completed = true;
                parentTraversalId    = completion.parentTraversalId;
                modelNodeId          = completion.modelNodeId;
                aggregateOwner       = completion.aggregateOwner;
                bytes                = completion.bytes;
                folderCount          = completion.folderCount;
                fileCount            = completion.fileCount;
                aggregateRoot        = completion.aggregateRoot;
                displayError         = completion.displayError;
                failed               = completion.failed;
                contentNotMeasured   = completion.contentNotMeasured;
                completedRoot        = traversalId == 1u;
                if (completedRoot)
                {
                    // An unreadable root has only a local failure (displayError); the scan as a whole still failed.
                    frontierFinished = true;
                    rootFailed.store(failed || displayError, std::memory_order_relaxed);
                }
                else
                {
                    completions.erase(it);
                }
            }

            if (completedRoot)
            {
                PendingUpdate finalProgress;
                finalProgress.kind                     = PendingUpdate::Kind::Progress;
                finalProgress.generation               = generation;
                finalProgress.nodeId                   = rootNodeId;
                finalProgress.bytes                    = scannedBytes.load(std::memory_order_relaxed);
                finalProgress.scannedFolders           = scannedFolders.load(std::memory_order_relaxed);
                finalProgress.scannedFiles             = scannedFiles.load(std::memory_order_relaxed);
                finalProgress.failedFolders            = failedFolders.load(std::memory_order_relaxed);
                finalProgress.firstPreviewPriorityJobs = firstPreviewPriorityJobs.load(std::memory_order_relaxed);
                PostUpdate(std::move(finalProgress));
            }

            postDirectoryCompletion(modelNodeId, aggregateOwner, bytes, displayError ? ScanState::Error : ScanState::Done, contentNotMeasured);

            // The parent roll-up below adds these bytes unconditionally, so the "Other" summary must be posted even for a failed subtree;
            // otherwise the owner's total exceeds the sum of its tiles.
            if (aggregateRoot && aggregateOwner != 0u)
            {
                std::vector<PendingUpdate> aggregateUpdates;
                std::vector<FileSummaryItem> noFiles;
#ifdef ENABLE_TESTS
                {
                    ScanPhaseTimer phaseTimer(captureScanPhases, summaryEncodePhase);
                    static_cast<void>(appendDirectoryFilesSummary(aggregateUpdates, aggregateOwner, traversalId, bytes, folderCount, fileCount, noFiles));
                }
#else
                static_cast<void>(appendDirectoryFilesSummary(aggregateUpdates, aggregateOwner, traversalId, bytes, folderCount, fileCount, noFiles));
#endif
                for (PendingUpdate& aggregateUpdate : aggregateUpdates)
                {
                    PostUpdate(std::move(aggregateUpdate));
                }
            }

            if (completedRoot)
            {
                frontierCv.notify_all();
                return;
            }

            uint64_t parentBytes       = 0;
            uint32_t parentModelNodeId = 0u;
            bool parentReady           = false;
            {
                std::scoped_lock lock(frontierMutex);
                auto parentIt = completions.find(parentTraversalId);
                if (parentIt == completions.end())
                {
                    return;
                }

                DirectoryCompletion& parent = parentIt->second;
                parent.bytes                = ViewerSpaceScan::SaturatingAdd(parent.bytes, bytes);
                parent.folderCount          = ViewerSpaceScan::SaturatingAdd(parent.folderCount, folderCount);
                parent.fileCount            = ViewerSpaceScan::SaturatingAdd(parent.fileCount, fileCount);
                parent.displayError         = parent.displayError || failed;
                parent.failed               = parent.failed || failed;
                parent.contentNotMeasured   = parent.contentNotMeasured || contentNotMeasured;
                parentBytes                 = parent.bytes;
                parentModelNodeId           = parent.modelNodeId;
                if (parent.pendingChildren > 0u)
                {
                    parent.pendingChildren -= 1u;
                }
                parentReady = parent.enumerated && parent.pendingChildren == 0u && ! parent.completed;
            }

            if (parentModelNodeId != 0u)
            {
                PendingUpdate parentSizeUp;
                parentSizeUp.kind       = PendingUpdate::Kind::UpdateSize;
                parentSizeUp.generation = generation;
                parentSizeUp.nodeId     = parentModelNodeId;
                parentSizeUp.bytes      = parentBytes;
                PostUpdate(std::move(parentSizeUp));
            }

            if (! parentReady)
            {
                return;
            }

            traversalId = parentTraversalId;
        }
    };

    auto processJob = [&](ScanJob job, uint32_t workerIndex) noexcept
    {
        ProgressState progress;

        StackItem item;
        auto releaseWorkerStaging = wil::scope_exit([&]() noexcept
        {
            // Destroy the charged owning arrays before releasing admission to another generation.
            std::vector<ChildDir>().swap(item.children);
            std::vector<FileSummaryItem>().swap(item.topFiles);
            if (item.stagingReservationBytes != 0u)
                workerStagingBudget.Release(item.stagingReservationBytes);
        });
        item.traversalId          = job.traversalId;
        item.modelNodeId          = job.modelNodeId;
        item.aggregateOwner       = job.aggregateOwner;
        item.aggregateRoot        = job.aggregateRoot;
        item.depth                = job.depth;
        item.path                 = std::move(job.path);
        if (job.pathBudgetFailed)
        {
            item.failed = true;
            resources.RecordValidationFailure(ViewerSpaceScan::ValidationError::OutstandingLimit);
        }
        else if (item.depth > resourcePolicy.maxTraversalDepth)
        {
            item.failed = true;
            resources.RecordValidationFailure(ViewerSpaceScan::ValidationError::DepthLimit);
        }
        else
        {
            enumerate(item, progress, workerIndex);
        }

        if (stopToken.stop_requested())
        {
            return;
        }

        std::vector<ScanJob> childJobs;
        uint64_t reservedChildPathBytes       = 0u;
        bool pathReservationHeld              = false;
        uint64_t reservedGlobalChildPathBytes = 0u;
        bool globalPathReservationHeld        = false;
        auto releasePathReservation           = wil::scope_exit([&]() noexcept
        {
            if (pathReservationHeld)
            {
                {
                    std::scoped_lock lock(frontierMutex);
                    reservedPathBytes = reservedChildPathBytes >= reservedPathBytes ? 0u : reservedPathBytes - reservedChildPathBytes;
                }
                frontierCv.notify_all();
            }
            if (globalPathReservationHeld)
            {
                static_cast<void>(ReleaseAtomicBudget(_activePathCurrentBytes, reservedGlobalChildPathBytes));
            }
        });
        if (! item.failed)
        {
            uint64_t prospectivePathBytes = 0u;
            for (const ChildDir& child : item.children)
            {
                if (stopToken.stop_requested())
                {
                    return;
                }
                if (! child.descend)
                {
                    continue;
                }
                const bool addSeparator  = ! item.path.empty() && item.path.back() != L'/' && item.path.back() != L'\\';
                const uint64_t pathChars = ViewerSpaceScan::SaturatingAdd(
                    ViewerSpaceScan::SaturatingAdd(static_cast<uint64_t>(item.path.size()), static_cast<uint64_t>(child.name.size())),
                    addSeparator ? uint64_t{1u} : uint64_t{0u});
                if (pathChars > resourcePolicy.maxPathChars)
                {
                    item.failed = true;
                    resources.RecordValidationFailure(ViewerSpaceScan::ValidationError::PathLimit);
                    break;
                }
                // JoinPath reserves once, so there is no second parent-copy buffer during construction.
                prospectivePathBytes = ViewerSpaceScan::SaturatingAdd(prospectivePathBytes, ConservativePathAllocationBytes(pathChars));
            }
            if (! item.failed && prospectivePathBytes != 0u)
            {
                std::scoped_lock lock(frontierMutex);
                const uint64_t admittedBytes = ViewerSpaceScan::SaturatingAdd(livePathBytes, reservedPathBytes);
                uint64_t aggregatePathBytes  = 0u;
                if (prospectivePathBytes > activePathBudgetBytes - std::min(admittedBytes, activePathBudgetBytes) ||
                    ! TryReserveAtomicBudget(_activePathCurrentBytes, prospectivePathBytes, activePathBudgetBytes, aggregatePathBytes))
                {
                    item.failed = true;
                    _activePathRejectedCount.fetch_add(1u, std::memory_order_relaxed);
                    resources.RecordValidationFailure(ViewerSpaceScan::ValidationError::OutstandingLimit);
                }
                else
                {
                    reservedChildPathBytes       = prospectivePathBytes;
                    reservedPathBytes            = ViewerSpaceScan::SaturatingAdd(reservedPathBytes, prospectivePathBytes);
                    pathReservationHeld          = true;
                    reservedGlobalChildPathBytes = prospectivePathBytes;
                    globalPathReservationHeld    = true;
                    UpdateAtomicPeak(_activePathPeakBytes, aggregatePathBytes);
                }
            }
        }
        if (! item.failed)
        {
            childJobs.reserve(item.children.size());
            for (const ChildDir& child : item.children)
            {
                if (stopToken.stop_requested())
                {
                    return;
                }
                if (! child.descend)
                {
                    continue;
                }
                ScanJob childJob;
                childJob.parentTraversalId = item.traversalId;
                childJob.depth             = ViewerSpaceScan::SaturatingAdd(item.depth, 1u);
                childJob.path              = JoinPath(item.path, child.name, pathSeparator);
                if (childJob.depth > resourcePolicy.maxTraversalDepth || childJob.path.size() > resourcePolicy.maxPathChars)
                {
                    item.failed = true;
                    resources.RecordValidationFailure(childJob.depth > resourcePolicy.maxTraversalDepth ? ViewerSpaceScan::ValidationError::DepthLimit
                                                                                                        : ViewerSpaceScan::ValidationError::PathLimit);
                    childJobs.clear();
                    break;
                }
                const bool addSeparator  = ! item.path.empty() && item.path.back() != L'/' && item.path.back() != L'\\';
                const uint64_t pathChars = ViewerSpaceScan::SaturatingAdd(
                    ViewerSpaceScan::SaturatingAdd(static_cast<uint64_t>(item.path.size()), static_cast<uint64_t>(child.name.size())),
                    addSeparator ? uint64_t{1u} : uint64_t{0u});
                childJob.pathBytes = ConservativePathAllocationBytes(pathChars);
                if (static_cast<uint64_t>(childJob.path.capacity()) + 1u > childJob.pathBytes / sizeof(wchar_t))
                {
                    item.failed = true;
                    resources.RecordValidationFailure(ViewerSpaceScan::ValidationError::OutstandingLimit);
                    childJobs.clear();
                    break;
                }
                childJob.pathBudgetLease = true;
                childJobs.push_back(std::move(childJob));
            }
        }

        // Admission and bookkeeping happen under frontierMutex; the model records are posted only after it is released, because
        // PostUpdate can sleep for backpressure and a throttled worker must not stall every other worker on the frontier lock.
        // Child jobs are published to the frontier only after these records are posted, so no worker can report on a child node
        // before its AddChild record is queued.
        std::vector<PendingUpdate> deferredUpdates;
        uint64_t frontierLockWaitUs = 0u;
        {
            const auto frontierLockRequestedAt = std::chrono::steady_clock::now();
            std::scoped_lock lock(frontierMutex);
            frontierLockWaitUs      = Debug::Perf::ElapsedUs(frontierLockRequestedAt);
            const auto completionIt = completions.find(item.traversalId);
            if (completionIt == completions.end())
            {
                return;
            }
            DirectoryCompletion& completion = completionIt->second;
            completion.bytes                = ViewerSpaceScan::SaturatingAdd(completion.bytes, item.bytes);
            completion.folderCount          = ViewerSpaceScan::SaturatingAdd(completion.folderCount, item.nonDescendingFolders);
            completion.fileCount            = ViewerSpaceScan::SaturatingAdd(completion.fileCount, item.fileCount);
            completion.enumerated           = true;
            completion.displayError         = completion.displayError || item.failed;
            completion.failed               = completion.failed || (item.failed && item.propagateFailure);
            completion.contentNotMeasured   = completion.contentNotMeasured || item.contentNotMeasured;

            uint64_t childPathBytes = 0u;
            for (const ScanJob& childJob : childJobs)
            {
                childPathBytes = ViewerSpaceScan::SaturatingAdd(childPathBytes, childJob.pathBytes);
            }
            if (! item.failed && (childJobs.size() > resourcePolicy.maxOutstandingDirectories ||
                                  completions.size() > resourcePolicy.maxOutstandingDirectories - childJobs.size() ||
                                  childPathBytes > activePathBudgetBytes - std::min(livePathBytes, activePathBudgetBytes)))
            {
                completion.failed       = true;
                completion.displayError = true;
                item.failed             = true;
                resources.RecordValidationFailure(ViewerSpaceScan::ValidationError::OutstandingLimit);
                childJobs.clear();
            }
            if (! item.failed)
            {
                if (item.modelNodeId != 0u &&
                    (nextModelId == 0u || nextModelId >= kFileRecordIdBase || item.children.size() > static_cast<size_t>(kFileRecordIdBase - nextModelId)))
                {
                    completion.failed       = true;
                    completion.displayError = true;
                    item.failed             = true;
                    resources.RecordValidationFailure(ViewerSpaceScan::ValidationError::ItemIdLimit);
                    childJobs.clear();
                }
                else if (! resources.TryTraverseDirectories(static_cast<uint64_t>(childJobs.size())))
                {
                    completion.failed       = true;
                    completion.displayError = true;
                    item.failed             = true;
                    resources.RecordValidationFailure(ViewerSpaceScan::ValidationError::TraversalLimit);
                    childJobs.clear();
                }
            }

            if (! item.failed)
            {
                if (pathReservationHeld)
                {
                    reservedPathBytes   = childPathBytes >= reservedPathBytes ? 0u : reservedPathBytes - childPathBytes;
                    pathReservationHeld = false;
                    livePathBytes       = ViewerSpaceScan::SaturatingAdd(livePathBytes, childPathBytes);
                }
                size_t descendingIndex = 0u;
                for (ChildDir& child : item.children)
                {
                    child.aggregateOwner = item.modelNodeId != 0u ? item.modelNodeId : item.aggregateOwner;
                    child.aggregateRoot  = item.modelNodeId != 0u && child.descend;

                    const uint64_t nameBytes = static_cast<uint64_t>(child.name.size()) * sizeof(wchar_t);
#ifdef ENABLE_TESTS
                    if (item.modelNodeId != 0u && tryRetainDirectoryProfiled(nameBytes))
#else
                    if (item.modelNodeId != 0u && resources.TryRetainDirectory(nameBytes))
#endif
                    {
                        child.modelNodeId    = nextModelId;
                        child.aggregateOwner = nextModelId;
                        child.aggregateRoot  = false;
                        nextModelId += 1u;
                    }
                    else if (item.modelNodeId == 0u)
                    {
                        SaturatingAtomicAdd(resources.cappedDirectories, 1u);
                        child.aggregateRoot = false;
                    }
                    else if (! child.descend)
                    {
                        item.otherFolders = ViewerSpaceScan::SaturatingAdd(item.otherFolders, uint64_t{1u});
                    }

                    if (child.descend)
                    {
                        ScanJob& childJob       = childJobs[descendingIndex++];
                        child.traversalId       = nextTraversalId.fetch_add(1u, std::memory_order_relaxed);
                        childJob.traversalId    = child.traversalId;
                        childJob.modelNodeId    = child.modelNodeId;
                        childJob.aggregateOwner = child.aggregateOwner;
                        childJob.aggregateRoot  = child.aggregateRoot;
                    }
                }

                std::vector<FileSummaryItem> retainedFiles;
                retainedFiles.reserve(item.topFiles.size());
                for (FileSummaryItem& file : item.topFiles)
                {
                    const uint64_t nameBytes = static_cast<uint64_t>(file.name.size()) * sizeof(wchar_t);
#ifdef ENABLE_TESTS
                    if (tryRetainFileProfiled(nameBytes))
#else
                    if (resources.TryRetainFile(nameBytes))
#endif
                    {
                        retainedFiles.push_back(std::move(file));
                    }
                    else
                    {
                        item.otherBytes = ViewerSpaceScan::SaturatingAdd(item.otherBytes, file.bytes);
                        item.otherFiles = ViewerSpaceScan::SaturatingAdd(item.otherFiles, uint64_t{1u});
                    }
                }
                item.topFiles = std::move(retainedFiles);

                if (item.modelNodeId != 0u)
                {
                    std::sort(item.topFiles.begin(),
                              item.topFiles.end(),
                              [](const FileSummaryItem& left, const FileSummaryItem& right) noexcept
                    {
                        if (left.bytes != right.bytes)
                        {
                            return left.bytes > right.bytes;
                        }
                        return NaturalSort::Less(left.name, right.name);
                    });

                    PendingUpdate sizeUp;
                    sizeUp.kind       = PendingUpdate::Kind::UpdateSize;
                    sizeUp.generation = generation;
                    sizeUp.nodeId     = item.modelNodeId;
                    sizeUp.bytes      = item.bytes;
                    deferredUpdates.push_back(std::move(sizeUp));

#ifdef ENABLE_TESTS
                    {
                        ScanPhaseTimer phaseTimer(captureScanPhases, summaryEncodePhase);
                        static_cast<void>(appendDirectoryFilesSummary(
                            deferredUpdates, item.modelNodeId, item.traversalId, item.otherBytes, item.otherFolders, item.otherFiles, item.topFiles));
                    }
#else
                    static_cast<void>(appendDirectoryFilesSummary(
                        deferredUpdates, item.modelNodeId, item.traversalId, item.otherBytes, item.otherFolders, item.otherFiles, item.topFiles));
#endif
                }

                for (const ChildDir& child : item.children)
                {
                    if (child.modelNodeId == 0u)
                    {
                        continue;
                    }

                    PendingUpdate addDir;
                    addDir.kind          = PendingUpdate::Kind::AddChild;
                    addDir.generation    = generation;
                    addDir.parentId      = item.modelNodeId;
                    addDir.nodeId        = child.modelNodeId;
                    addDir.name          = child.name;
                    addDir.isDirectory   = true;
                    addDir.lastWriteTime = child.lastWriteTime;
                    deferredUpdates.push_back(std::move(addDir));

                    if (! child.descend)
                    {
                        PendingUpdate leafSizeUp;
                        leafSizeUp.kind       = PendingUpdate::Kind::UpdateSize;
                        leafSizeUp.generation = generation;
                        leafSizeUp.nodeId     = child.modelNodeId;
                        leafSizeUp.bytes      = 0u;
                        deferredUpdates.push_back(std::move(leafSizeUp));

                        PendingUpdate leafDoneUp;
                        leafDoneUp.kind       = PendingUpdate::Kind::UpdateState;
                        leafDoneUp.generation = generation;
                        leafDoneUp.nodeId     = child.modelNodeId;
                        leafDoneUp.state      = ScanState::Done;
                        deferredUpdates.push_back(std::move(leafDoneUp));
                    }
                }
            }

            completion.pendingChildren = static_cast<uint32_t>(childJobs.size());

            for (auto& childJob : childJobs)
            {
                DirectoryCompletion childCompletion;
                childCompletion.parentTraversalId = item.traversalId;
                childCompletion.modelNodeId       = childJob.modelNodeId;
                childCompletion.aggregateOwner    = childJob.aggregateOwner;
                childCompletion.aggregateRoot     = childJob.aggregateRoot;
                childCompletion.folderCount       = 1u;
                completions.emplace(childJob.traversalId, childCompletion);
            }
        }

#ifdef ENABLE_TESTS
        if (ReadPositiveUint32Environment(kTestSummaryCopyBudgetEnvVar).has_value())
        {
            const auto isSummaryPacket = [](const PendingUpdate& update) noexcept
            {
                return update.kind == PendingUpdate::Kind::DirectorySummaryBegin || update.kind == PendingUpdate::Kind::DirectorySummaryChunk ||
                       update.kind == PendingUpdate::Kind::DirectorySummaryEnd || update.kind == PendingUpdate::Kind::DirectorySummaryReject;
            };
            const auto firstSummary = std::find_if(deferredUpdates.begin(), deferredUpdates.end(), isSummaryPacket);
            if (firstSummary != deferredUpdates.end())
            {
                const auto summaryEnd = std::find_if_not(firstSummary, deferredUpdates.end(), isSummaryPacket);
                if (summaryEnd != deferredUpdates.end() && std::find_if(summaryEnd, deferredUpdates.end(), isSummaryPacket) == deferredUpdates.end())
                {
                    // Test-only ordering makes the parent child block nonempty before the summary transaction begins.
                    std::rotate(firstSummary, summaryEnd, deferredUpdates.end());
                }
            }
        }
#endif
        Debug::Perf::EmitDurationUs(L"viewer.space.scan.frontier_lock_wait_us", frontierLockWaitUs, static_cast<uint64_t>(childJobs.size()), workerIndex);
        for (PendingUpdate& update : deferredUpdates)
        {
            PostUpdate(std::move(update));
        }

        size_t frontierDepth = 0u;
        if (! childJobs.empty())
        {
            std::scoped_lock lock(frontierMutex);
            bool prioritizeFirstChild = scannedBytes.load(std::memory_order_relaxed) == 0u && firstPreviewPriorityJobs.load(std::memory_order_relaxed) < 16u;
            for (ScanJob& childJob : childJobs)
            {
                if (prioritizeFirstChild)
                {
                    frontier.push_front(std::move(childJob));
                    firstPreviewPriorityJobs.fetch_add(1u, std::memory_order_relaxed);
                    prioritizeFirstChild = false;
                }
                else
                {
                    frontier.push_back(std::move(childJob));
                }
            }
            if (globalPathReservationHeld)
            {
                globalPathReservationHeld    = false;
                reservedGlobalChildPathBytes = 0u;
            }
            frontierDepth = frontier.size();
        }

        updateFrontierDepthPeak(static_cast<uint32_t>(std::min<size_t>(frontierDepth, std::numeric_limits<uint32_t>::max())));
        if (! childJobs.empty())
        {
            Debug::Perf::EmitValue(L"viewer.space.scan.frontier_depth", static_cast<uint64_t>(frontierDepth));
            frontierCv.notify_all();
        }

        completeReadyDirectories(item.traversalId);
    };

    auto runWorker = [&](uint32_t workerIndex) noexcept
    {
        // A user-requested scan keeps ordinary I/O and memory priority. Lower CPU priority alone lets UI work preempt it
        // without background-mode I/O throttling. This policy is local to ViewerSpace, not a generic thread helper.
        const int previousPriority = GetThreadPriority(GetCurrentThread());
        const BOOL priorityApplied = previousPriority != THREAD_PRIORITY_ERROR_RETURN ? SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_BELOW_NORMAL) : 0;
        if (priorityApplied != 0)
        {
            cpuPriorityAppliedCount.fetch_add(1u, std::memory_order_relaxed);
        }
        else
        {
            const DWORD error = GetLastError();
            cpuPriorityFailure.store(HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_INVALID_FUNCTION : error), std::memory_order_relaxed);
            Debug::Warning(L"ViewerSpace could not lower scan CPU priority. error={0}", error);
        }
        auto restoreThreadPriority = wil::scope_exit([&, previousPriority, priorityApplied]
        {
            if (priorityApplied != 0)
            {
                if (SetThreadPriority(GetCurrentThread(), previousPriority))
                {
                    cpuPriorityRestoredCount.fetch_add(1u, std::memory_order_relaxed);
                }
                else
                {
                    const DWORD error = GetLastError();
                    cpuPriorityFailure.store(HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_INVALID_FUNCTION : error), std::memory_order_relaxed);
                    Debug::Warning(L"ViewerSpace could not restore scan CPU priority. error={0}", error);
                }
            }
        });

        for (;;)
        {
            ScanJob job;
            uint32_t activeAfterStart    = 0u;
            size_t frontierDepthAfterPop = 0u;
            {
                std::unique_lock lock(frontierMutex);
                while (frontier.empty() && ! frontierFinished && ! stopToken.stop_requested())
                {
                    idleWorkerWaits.fetch_add(1u, std::memory_order_relaxed);
                    frontierCv.wait_for(lock, std::chrono::milliseconds(25));
                }

                if (stopToken.stop_requested() || (frontier.empty() && frontierFinished))
                {
                    return;
                }

                if (frontier.empty())
                {
                    continue;
                }

                job = std::move(frontier.front());
                frontier.pop_front();
                frontierDepthAfterPop = frontier.size();
                activeAfterStart      = activeWorkers.fetch_add(1u, std::memory_order_relaxed) + 1u;
            }

            Debug::Perf::EmitValue(L"viewer.space.scan.active_workers", activeAfterStart);
            Debug::Perf::EmitValue(L"viewer.space.scan.frontier_depth", static_cast<uint64_t>(frontierDepthAfterPop));

            const uint64_t completedPathBytes = job.pathBytes;
            const bool pathBudgetLease        = job.pathBudgetLease;
            processJob(std::move(job), workerIndex);
            if (pathBudgetLease)
            {
                static_cast<void>(ReleaseAtomicBudget(_activePathCurrentBytes, completedPathBytes));
            }

            const uint32_t activeAfterFinish = activeWorkers.fetch_sub(1u, std::memory_order_relaxed) - 1u;
            Debug::Perf::EmitValue(L"viewer.space.scan.active_workers", activeAfterFinish);
            {
                std::scoped_lock lock(frontierMutex);
                livePathBytes = completedPathBytes >= livePathBytes ? 0u : livePathBytes - completedPathBytes;
                if (frontier.empty() && activeAfterFinish == 0u)
                {
                    frontierCv.notify_all();
                }
            }
        }
    };

    {
        const size_t extraWorkers = threadCount > 0u ? (static_cast<size_t>(threadCount) - 1u) : 0u;

        std::vector<std::jthread> workers;
        workers.reserve(extraWorkers);
        for (size_t workerIndex = 0; workerIndex < extraWorkers; ++workerIndex)
        {
            const uint32_t metricWorkerIndex = static_cast<uint32_t>(workerIndex + 1u);
            workers.emplace_back([&runWorker, metricWorkerIndex](std::stop_token) noexcept { runWorker(metricWorkerIndex); });
        }

        runWorker(0u);
    }

    if (stopToken.stop_requested())
    {
        uint64_t abandonedPathBytes = 0u;
        {
            std::scoped_lock lock(frontierMutex);
            for (const ScanJob& job : frontier)
            {
                if (job.pathBudgetLease)
                {
                    abandonedPathBytes = ViewerSpaceScan::SaturatingAdd(abandonedPathBytes, job.pathBytes);
                }
            }
            frontier.clear();
        }
        if (abandonedPathBytes != 0u)
        {
            static_cast<void>(ReleaseAtomicBudget(_activePathCurrentBytes, abandonedPathBytes));
        }
    }

    if (stopToken.stop_requested())
    {
        scanFinalState = ScanState::Canceled;
        scanFinalHr    = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        PendingUpdate canceled;
        canceled.kind       = PendingUpdate::Kind::UpdateState;
        canceled.generation = generation;
        canceled.nodeId     = rootNodeId;
        canceled.state      = ScanState::Canceled;
        PostUpdate(std::move(canceled));
        return;
    }

    Debug::Perf::EmitValue(L"viewer.space.scan.frontier_depth_peak", frontierDepthPeak.load(std::memory_order_relaxed), scanFinalHr);
    Debug::Perf::EmitValue(L"viewer.space.scan.idle_worker_waits", idleWorkerWaits.load(std::memory_order_relaxed), scanFinalHr);
    if (rootFailed.load(std::memory_order_relaxed))
    {
        scanFinalState = ScanState::Error;
        scanFinalHr    = E_FAIL;
    }
}

void ViewerSpace::PostUpdate(PendingUpdate&& update) noexcept
{
    if (update.generation != _scanGeneration.load())
    {
        return;
    }

#ifdef ENABLE_TESTS
    if (_debugPauseNextPostUpdate.exchange(false, std::memory_order_acq_rel))
    {
        const HANDLE enteredEvent = _debugPostUpdateEnteredEvent.load(std::memory_order_relaxed);
        const HANDLE releaseEvent = _debugPostUpdateReleaseEvent.load(std::memory_order_relaxed);
        if (enteredEvent && releaseEvent)
        {
            static_cast<void>(SetEvent(enteredEvent));
            static_cast<void>(WaitForSingleObject(releaseEvent, 5000u));
        }
    }
#endif

    static constexpr size_t kMaxOrdinarySummaryChunkBytes = 64u * 1024u;
    // This transport ceiling is immutable: workers must not read the UI-owned per-scan policy here.
    static constexpr size_t kTransportMaxProviderNameBytes  = ViewerSpaceScan::kProductionResourcePolicy.maxProviderNameBytes;
    static constexpr uint64_t kPendingUpdateQueueLimitBytes = 8ull * 1024ull * 1024ull;

    if (update.kind == PendingUpdate::Kind::DirectorySummaryChunk)
    {
        const size_t chunkBytes                 = EstimatePendingUpdateBytes(update);
        const size_t maxNameBytes               = kTransportMaxProviderNameBytes;
        const size_t boundedSingleNameException = maxNameBytes > (std::numeric_limits<size_t>::max)() - kMaxOrdinarySummaryChunkBytes
                                                      ? (std::numeric_limits<size_t>::max)()
                                                      : kMaxOrdinarySummaryChunkBytes + maxNameBytes;
        const bool validSingleNameException =
            update.topFiles.size() == 1u &&
            ViewerSpaceScan::ValidateProviderChildName(update.topFiles.front().name) == ViewerSpaceScan::ValidationError::None &&
            update.topFiles.front().name.size() <= maxNameBytes / sizeof(wchar_t) && chunkBytes <= boundedSingleNameException;
        if (chunkBytes > kMaxOrdinarySummaryChunkBytes && ! validSingleNameException)
        {
            update.kind = PendingUpdate::Kind::DirectorySummaryReject;
            std::vector<FileSummaryItem>().swap(update.topFiles);
            update.summaryNameBytes      = 0u;
            update.summarySequence       = 0u;
            update.summaryExpectedChunks = 0u;
            update.summaryExpectedFiles  = 0u;
        }
    }

    size_t updateBytes = EstimatePendingUpdateBytes(update);
    if (updateBytes > kPendingUpdateQueueLimitBytes)
    {
        // Every producer payload is bounded by the scan policy. A packet beyond the queue
        // cap indicates malformed accounting/input; do not admit it or stall a worker forever.
        if (update.kind == PendingUpdate::Kind::DirectorySummaryBegin || update.kind == PendingUpdate::Kind::DirectorySummaryChunk ||
            update.kind == PendingUpdate::Kind::DirectorySummaryEnd)
        {
            update.kind = PendingUpdate::Kind::DirectorySummaryReject;
            std::vector<FileSummaryItem>().swap(update.topFiles);
            std::wstring().swap(update.name);
            update.summaryNameBytes      = 0u;
            update.summarySequence       = 0u;
            update.summaryExpectedChunks = 0u;
            update.summaryExpectedFiles  = 0u;
        }
        updateBytes = EstimatePendingUpdateBytes(update);
        if (updateBytes > kPendingUpdateQueueLimitBytes)
        {
            return;
        }
    }

    size_t pendingCount   = 0u;
    uint64_t pendingBytes = 0u;
    bool coalesced        = false;
    for (;;)
    {
        coalesced = false;
        {
            std::scoped_lock lock(_updateMutex);
            const auto lockAcquiredAt = std::chrono::steady_clock::now();
            bool generationAccepted   = false;
            auto recordLockHold       = wil::scope_exit([this, lockAcquiredAt, &generationAccepted]() noexcept
            {
                if (! generationAccepted)
                {
                    return;
                }
                const uint64_t lockHoldUs = Debug::Perf::ElapsedUs(lockAcquiredAt);
                SaturatingAtomicAdd(_pendingUpdateLockHoldCount, 1u);
                SaturatingAtomicAdd(_pendingUpdateLockHoldTotalUs, lockHoldUs);
                UpdateAtomicPeak(_pendingUpdateLockHoldMaxUs, lockHoldUs);
            });
            if (update.generation != _scanGeneration.load(std::memory_order_acquire))
            {
#ifdef ENABLE_TESTS
                _debugPostUpdateInnerGenerationRejects.fetch_add(1u, std::memory_order_acq_rel);
#endif
                return;
            }
            generationAccepted = true;
            if (update.kind == PendingUpdate::Kind::UpdateSize || update.kind == PendingUpdate::Kind::Progress)
            {
                size_t examined = 0u;
                for (auto it = _pendingUpdates.rbegin(); it != _pendingUpdates.rend() && examined < 64u; ++it, ++examined)
                {
                    if (it->kind != update.kind || it->generation != update.generation || it->nodeId != update.nodeId)
                    {
                        continue;
                    }

                    const size_t previousBytes    = EstimatePendingUpdateBytes(*it);
                    const size_t replacementBytes = updateBytes;
                    const uint64_t queuedBytes    = _pendingUpdateBytes.load(std::memory_order_relaxed);
                    const uint64_t reducedBytes   = queuedBytes > previousBytes ? queuedBytes - previousBytes : 0u;
                    if (replacementBytes <= kPendingUpdateQueueLimitBytes - std::min<uint64_t>(reducedBytes, kPendingUpdateQueueLimitBytes))
                    {
                        *it = std::move(update);
                        if (replacementBytes >= previousBytes)
                        {
                            pendingBytes = _pendingUpdateBytes.fetch_add(static_cast<uint64_t>(replacementBytes - previousBytes), std::memory_order_relaxed) +
                                           static_cast<uint64_t>(replacementBytes - previousBytes);
                        }
                        else
                        {
                            pendingBytes = SubtractAtomicFloor(_pendingUpdateBytes, static_cast<uint64_t>(previousBytes - replacementBytes));
                        }
                        coalesced    = true;
                        pendingCount = _pendingUpdates.size();
                    }
                    break;
                }
            }

            if (! coalesced)
            {
                const uint64_t queuedBytes = _pendingUpdateBytes.load(std::memory_order_relaxed);
                if (updateBytes <= kPendingUpdateQueueLimitBytes - std::min<uint64_t>(queuedBytes, kPendingUpdateQueueLimitBytes))
                {
                    _pendingUpdates.emplace_back(std::move(update));
                    pendingBytes =
                        _pendingUpdateBytes.fetch_add(static_cast<uint64_t>(updateBytes), std::memory_order_relaxed) + static_cast<uint64_t>(updateBytes);
                    pendingCount = _pendingUpdates.size();
                }
            }
            UpdateAtomicPeak(_pendingUpdatePeakBytes, pendingBytes);
            if (coalesced || pendingCount != 0u)
            {
                SaturatingAtomicAdd(_pendingUpdatePostedCount, 1u);
                SaturatingAtomicAdd(_pendingUpdateAcceptedBytes, static_cast<uint64_t>(updateBytes));
                if (coalesced)
                {
                    SaturatingAtomicAdd(_pendingUpdateCoalescedCount, 1u);
                }
            }
        }

        if (coalesced || pendingCount != 0u)
        {
            break;
        }
        if (update.generation != _scanGeneration.load(std::memory_order_acquire))
        {
            return;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(2));
    }

    if (pendingBytes > (kPendingUpdateQueueLimitBytes / 2ull))
    {
        std::this_thread::yield();
    }
}

bool ViewerSpace::CanBeginScanProjection(bool hasPublishedProjection) const noexcept
{
    // A builder borrows the model between turns. Once a preview exists, give a
    // backed-up transport priority instead of freezing it for another full projection.
    // The first preview remains eligible; in-flight builders still finish coherently.
    if (_scanActive.load() && hasPublishedProjection && _pendingUpdateBytes.load(std::memory_order_relaxed) > 4ull * 1024ull * 1024ull)
    {
        Debug::Perf::EmitCounter(L"viewer.space.projection.backlog_deferred_count");
        return false;
    }
    return true;
}

bool ViewerSpace::IsQueueTerminalReady() noexcept
{
    if ((_overallState != ScanState::Done && _overallState != ScanState::Error && _overallState != ScanState::Canceled) || _model == nullptr)
    {
        return false;
    }
    const Node* root = TryGetRealNode(_model->rootNodeId);
    if (root == nullptr || (root->scanState != ScanState::Done && root->scanState != ScanState::Error && root->scanState != ScanState::Canceled))
    {
        return false;
    }

    ReapFinishedScanWorkers();
    const auto isComplete = [](const ScanWorker& worker) noexcept
    { return worker.done != nullptr ? worker.done->load(std::memory_order_acquire) : ! worker.thread.joinable(); };
    if (! isComplete(_scanWorker))
    {
        return false;
    }
    for (const ScanWorker& worker : _retiredScanWorkers)
    {
        if (! isComplete(worker))
        {
            return false;
        }
    }

    std::scoped_lock lock(_updateMutex);
    return _pendingUpdates.empty() && _pendingUpdateBytes.load(std::memory_order_relaxed) == 0u;
}

void ViewerSpace::EmitQueueDiagnostics(bool terminalSnapshot, bool force) noexcept
{
    static constexpr auto kMinimumEmitInterval = std::chrono::milliseconds(250);
    const auto now                             = std::chrono::steady_clock::now();
    if (! force)
    {
        if (terminalSnapshot)
        {
            if (_queueDiagnosticsTerminalEmitted)
            {
                return;
            }
            _queueDiagnosticsTerminalEmitted = true;
        }
        else if (_lastQueueDiagnosticsEmitAt.time_since_epoch().count() != 0 && now - _lastQueueDiagnosticsEmitAt < kMinimumEmitInterval)
        {
            return;
        }
    }
    _lastQueueDiagnosticsEmitAt = now;

    // Counts, accepted bytes, and timing totals/maxima are cumulative since StartScan; pending_bytes is the current queue gauge. Legacy duration rows carry the
    // exact max sample, not a percentile.
    uint64_t postedCount      = 0u;
    uint64_t coalescedCount   = 0u;
    uint64_t pendingBytes     = 0u;
    uint64_t pendingPeakBytes = 0u;
    uint64_t acceptedBytes    = 0u;
    uint64_t lockHoldCount    = 0u;
    uint64_t lockHoldTotalUs  = 0u;
    uint64_t lockHoldMaxUs    = 0u;
    uint64_t drainedCount     = 0u;
    uint64_t drainCount       = 0u;
    uint64_t drainTotalUs     = 0u;
    uint64_t drainMaxUs       = 0u;
    {
        std::scoped_lock lock(_updateMutex);
        postedCount      = _pendingUpdatePostedCount.load(std::memory_order_relaxed);
        coalescedCount   = _pendingUpdateCoalescedCount.load(std::memory_order_relaxed);
        pendingBytes     = _pendingUpdateBytes.load(std::memory_order_relaxed);
        pendingPeakBytes = _pendingUpdatePeakBytes.load(std::memory_order_relaxed);
        acceptedBytes    = _pendingUpdateAcceptedBytes.load(std::memory_order_relaxed);
        lockHoldCount    = _pendingUpdateLockHoldCount.load(std::memory_order_relaxed);
        lockHoldTotalUs  = _pendingUpdateLockHoldTotalUs.load(std::memory_order_relaxed);
        lockHoldMaxUs    = _pendingUpdateLockHoldMaxUs.load(std::memory_order_relaxed);
        drainedCount     = _pendingUpdateDrainedCount.load(std::memory_order_relaxed);
        drainCount       = _pendingUpdateDrainCount.load(std::memory_order_relaxed);
        drainTotalUs     = _pendingUpdateDrainTotalUs.load(std::memory_order_relaxed);
        drainMaxUs       = _pendingUpdateDrainMaxUs.load(std::memory_order_relaxed);
    }

    Debug::Perf::EmitValue(L"viewer.space.queue.posted_count", postedCount);
    Debug::Perf::EmitValue(L"viewer.space.queue.coalesced_count", coalescedCount);
    Debug::Perf::EmitValue(L"viewer.space.queue.pending_bytes", pendingBytes);
    Debug::Perf::EmitValue(L"viewer.space.queue.pending_peak_bytes", pendingPeakBytes);
    Debug::Perf::EmitValue(L"viewer.space.queue.posted_bytes", acceptedBytes);
    Debug::Perf::EmitValue(L"viewer.space.queue.lock_hold_count", lockHoldCount);
    Debug::Perf::EmitValue(L"viewer.space.queue.lock_hold_total_us", lockHoldTotalUs);
    Debug::Perf::EmitValue(L"viewer.space.queue.lock_hold_max_us", lockHoldMaxUs);
    Debug::Perf::Emit(L"viewer.space.queue.lock_hold_us", L"cumulative-max-since-start-scan", lockHoldMaxUs, lockHoldCount, lockHoldTotalUs, S_OK);
    Debug::Perf::EmitValue(L"viewer.space.queue.drained_count", drainedCount);
    Debug::Perf::EmitValue(L"viewer.space.queue.drain_count", drainCount);
    Debug::Perf::EmitValue(L"viewer.space.queue.drain_total_us", drainTotalUs);
    Debug::Perf::EmitValue(L"viewer.space.queue.drain_max_us", drainMaxUs);
    Debug::Perf::Emit(L"viewer.space.queue.drain_us", L"cumulative-max-since-start-scan", drainMaxUs, drainCount, drainTotalUs, S_OK);

    // Summary telemetry is a sampled state snapshot, not a per-packet event stream. Keep these
    // reads on the UI owner alongside queue diagnostics; only the chunk high-water mark is atomic.
    Debug::Perf::EmitValue(L"viewer.space.scan.summary_stage_bytes", _model == nullptr ? 0u : _model->directorySummaryStageBytes);
    Debug::Perf::EmitValue(L"viewer.space.scan.summary_committed_count", _directorySummaryCommittedCount);
    Debug::Perf::EmitValue(L"viewer.space.scan.summary_rejected_count", _directorySummaryRejectedCount);
    Debug::Perf::EmitValue(L"viewer.space.scan.summary_chunk_max_bytes", _directorySummaryChunkMaxBytes.load(std::memory_order_relaxed));
}

void ViewerSpace::DrainUpdates() noexcept
{
    // Layout borrows generation-owned views. Defer mutations for its bounded preparation turns;
    // the transport remains byte bounded and producers retain cancellation/backpressure.
    const bool applyingRefreshCandidate = _workingModel != nullptr && _refreshPending;
    if (! applyingRefreshCandidate && (_layoutBuild || _analysisBuild))
    {
        return;
    }

    ModelGeneration* const priorModel = _model;
    if (applyingRefreshCandidate)
    {
        _model = _workingModel.get();
        if (! FitsViewerSpaceRefreshWorkingSet(ViewerOwnedBytes(), 0u))
        {
#ifdef ENABLE_TESTS
            Debug::Perf::Emit(L"viewer.space.refresh.admission_refused", L"before-drain", 0u, ViewerOwnedBytes(), 0u, E_OUTOFMEMORY);
#endif
            _model = priorModel;
            DiscardWorkingModel(true);
            return;
        }
    }
    auto restoreModel = wil::scope_exit([this, priorModel]() noexcept { _model = priorModel; });

    static constexpr double kDrainBudgetSecondsWhileScanning = 0.002;
    static constexpr double kDrainBudgetSecondsIdle          = 0.004;

    const bool scanning                             = _scanActive.load();
    const double budgetSeconds                      = scanning ? kDrainBudgetSecondsWhileScanning : kDrainBudgetSecondsIdle;
    const size_t maxUpdatesPerDrain                 = scanning ? 512u : 2048u;
    static constexpr size_t kMaxQueuedBytesPerDrain = 8u * 1024u * 1024u;

    const auto drainStartedAt  = std::chrono::steady_clock::now();
    const double startSeconds  = NowSeconds();
    size_t processed           = 0;
    size_t drainedUpdates      = 0;
    uint64_t pendingBytesAfter = _pendingUpdateBytes.load(std::memory_order_relaxed);

    bool layoutChanged   = false;
    bool headerTextDirty = false;
    bool inspectorDirty  = false;

    std::deque<PendingUpdate> batch;
    std::deque<PendingUpdate> deferredAddChildren;
    size_t batchBytes = 0u;
    {
        std::scoped_lock lock(_updateMutex);
        while (batch.size() < maxUpdatesPerDrain && ! _pendingUpdates.empty())
        {
            const size_t updateBytes = EstimatePendingUpdateBytes(_pendingUpdates.front());
            if (updateBytes > kMaxQueuedBytesPerDrain - std::min(batchBytes, kMaxQueuedBytesPerDrain))
            {
                break;
            }
            batchBytes += updateBytes;
            batch.emplace_back(std::move(_pendingUpdates.front()));
            _pendingUpdates.pop_front();
        }

        pendingBytesAfter = _pendingUpdateBytes.load(std::memory_order_relaxed);
    }

    size_t batchIndex = 0u;
    while (batchIndex < batch.size())
    {
        PendingUpdate update = std::move(batch[batchIndex]);
        batchIndex += 1u;
        processed += 1;
        const size_t updateBytes = EstimatePendingUpdateBytes(update);
        if (applyingRefreshCandidate)
        {
            const uint64_t updateAllowance = updateBytes > (std::numeric_limits<uint64_t>::max() - 4096u) / 2u
                                                 ? std::numeric_limits<uint64_t>::max()
                                                 : static_cast<uint64_t>(updateBytes) * 2u + 4096u;
            if (! FitsViewerSpaceRefreshWorkingSet(ViewerOwnedBytes(), updateAllowance))
            {
#ifdef ENABLE_TESTS
                Debug::Perf::Emit(L"viewer.space.refresh.admission_refused", L"before-update", 0u, ViewerOwnedBytes(), updateAllowance, E_OUTOFMEMORY);
#endif
                _model = priorModel;
                DiscardWorkingModel(true);
                return;
            }
        }
        bool keepQueuedBytes = false;
        if (update.kind == PendingUpdate::Kind::AddChild && HasPendingDirectorySummaryChildBlock(update.generation, update.parentId))
        {
            deferredAddChildren.emplace_back(std::move(update));
            keepQueuedBytes = true;
        }
        else
            switch (update.kind)
            {
                case PendingUpdate::Kind::PathIdentity:
                    _model->workspacePathIdentity = std::move(update.pathIdentity);
                    headerTextDirty               = true;
                    break;
                case PendingUpdate::Kind::AddChild:
                {
                    if (update.nodeId == 0u || update.nodeId >= kFileRecordIdBase || update.parentId == 0u ||
                        update.nodeId > _scanResourcePolicy.maxRetainedDirectories ||
                        update.name.size() > (_scanResourcePolicy.maxProviderNameBytes / sizeof(wchar_t)))
                    {
                        _overallState = ScanState::Error;
                        break;
                    }

                    Node node;
                    node.id            = update.nodeId;
                    node.parentId      = update.parentId;
                    node.isDirectory   = update.isDirectory;
                    node.isSynthetic   = update.isSynthetic;
                    node.name          = CopyToArena(_model->nameArena, update.name);
                    node.totalBytes    = 0u;
                    node.lastWriteTime = update.lastWriteTime;
                    node.scanState     = update.isDirectory ? ScanState::Queued : ScanState::Done;

                    const size_t requiredSize = static_cast<size_t>(node.id) + 1u;
                    if (requiredSize > static_cast<size_t>(_scanResourcePolicy.maxRetainedDirectories) + 1u)
                    {
                        _overallState = ScanState::Error;
                        break;
                    }
                    if (_model->nodes.size() < requiredSize)
                    {
                        _model->nodes.resize(requiredSize);
                    }

                    _model->nodes[node.id] = std::move(node);

                    Node* parent = TryGetRealNode(update.parentId);
                    if (parent == nullptr || ! AddRealNodeChild(*parent, update.nodeId))
                    {
                        _model->nodes[node.id] = {};
                        _overallState          = ScanState::Error;
                        break;
                    }

                    static_cast<void>(RaiseNodeObservedBytes(update.nodeId, update.bytes));
                    layoutChanged = true;
                    break;
                }
                case PendingUpdate::Kind::UpdateMetadata:
                {
                    if (Node* node = TryGetRealNode(update.nodeId))
                    {
                        node->lastWriteTime = update.lastWriteTime;
                        inspectorDirty      = inspectorDirty || _selectedItemId == node->id;
                    }
                    break;
                }
                case PendingUpdate::Kind::UpdateSize:
                {
                    Node* node = TryGetRealNode(update.nodeId);
                    if (node != nullptr)
                    {
                        // UpdateSize carries either partial local progress or a complete subtree total.
                        // Both are monotonic lower bounds; RaiseNodeObservedBytes merges them with committed contributions.
                        layoutChanged = RaiseNodeObservedBytes(node->id, update.bytes) || layoutChanged;
                    }
                    break;
                }
                case PendingUpdate::Kind::UpdateState:
                {
                    Node* node = TryGetRealNode(update.nodeId);
                    if (node != nullptr)
                    {
                        if (update.contentNotMeasured && ! node->contentNotMeasured)
                        {
                            node->contentNotMeasured = true;
                            layoutChanged            = true;
                            inspectorDirty =
                                _selectedItemId == node->id || (IsSyntheticNodeId(_selectedItemId) && (_selectedItemId & kTreemapItemIdMask) == node->id);
                        }
                        if (update.state == ScanState::Done && update.nodeId == _model->rootNodeId && ! _model->rejectedDirectorySummaryNodes.empty())
                        {
                            while (! _model->directorySummaryStages.empty())
                            {
                                const DirectorySummaryStageKey pendingKey = _model->directorySummaryStages.begin()->first;
                                static_cast<void>(RejectDirectorySummary(pendingKey, true));
                            }
                            node->scanState = ScanState::Error;
                            _model->deferredDirectoryDoneStates.clear();
                            layoutChanged = true;
                        }
                        else if (update.state == ScanState::Done && update.nodeId == _model->rootNodeId && ! _model->directorySummaryStages.empty())
                        {
                            const bool allStagesEnded = std::all_of(_model->directorySummaryStages.begin(),
                                                                    _model->directorySummaryStages.end(),
                                                                    [](const auto& entry) { return entry.second.endReceived; });
                            if (allStagesEnded)
                            {
                                // A terminal packet may overtake the UI's bounded child-span copy work.
                                // End proves the payload is complete; hold Done until all atomic commits finish.
                                _model->deferredDirectoryDoneStates[_model->rootNodeId] = ScanState::Done;
                            }
                            else
                            {
                                while (! _model->directorySummaryStages.empty())
                                {
                                    const DirectorySummaryStageKey pendingKey = _model->directorySummaryStages.begin()->first;
                                    static_cast<void>(RejectDirectorySummary(pendingKey, true));
                                }
                                node->scanState = ScanState::Error;
                                _model->deferredDirectoryDoneStates.clear();
                                layoutChanged = true;
                            }
                        }
                        else if (update.state == ScanState::Done && _model->rejectedDirectorySummaryNodes.contains(update.nodeId))
                        {
                            node->scanState = ScanState::Error;
                            layoutChanged   = true;
                        }
                        else if (update.state == ScanState::Done && HasDirectorySummaryStage(update.generation, update.nodeId))
                        {
                            bool everyStageEnded = true;
                            for (const auto& [key, stage] : _model->directorySummaryStages)
                            {
                                if (key.generation == update.generation && key.nodeId == update.nodeId && ! stage.endReceived)
                                {
                                    everyStageEnded = false;
                                    break;
                                }
                            }
                            if (everyStageEnded)
                            {
                                _model->deferredDirectoryDoneStates[update.nodeId] = update.state;
                            }
                            else
                            {
                                while (HasDirectorySummaryStage(update.generation, update.nodeId))
                                {
                                    const auto pendingStage =
                                        std::find_if(_model->directorySummaryStages.begin(), _model->directorySummaryStages.end(), [&](const auto& entry) {
                                        return entry.first.generation == update.generation && entry.first.nodeId == update.nodeId;
                                    });
                                    if (pendingStage == _model->directorySummaryStages.end())
                                    {
                                        break;
                                    }
                                    static_cast<void>(RejectDirectorySummary(pendingStage->first, true));
                                }
                                node->scanState = ScanState::Error;
                                layoutChanged   = true;
                            }
                        }
                        else if (update.state == ScanState::Canceled)
                        {
                            while (HasDirectorySummaryStage(update.generation, update.nodeId))
                            {
                                const auto pendingStage =
                                    std::find_if(_model->directorySummaryStages.begin(), _model->directorySummaryStages.end(), [&](const auto& entry) {
                                    return entry.first.generation == update.generation && entry.first.nodeId == update.nodeId;
                                });
                                if (pendingStage == _model->directorySummaryStages.end())
                                {
                                    break;
                                }
                                static_cast<void>(RejectDirectorySummary(pendingStage->first, false));
                            }
                            if (update.nodeId == _model->rootNodeId)
                            {
                                DiscardDirectorySummaryStages();
                            }
                            node->scanState = update.state;
                        }
                        else
                        {
                            node->scanState = update.state;
                        }
                        if (update.nodeId == _model->rootNodeId)
                        {
                            headerTextDirty = true;
                        }
                    }
                    break;
                }
                case PendingUpdate::Kind::DirectorySummaryBegin:
                {
                    const uint64_t rejectedBefore = _directorySummaryRejectedCount;
                    static_cast<void>(BeginDirectorySummary(update));
                    layoutChanged = layoutChanged || _directorySummaryRejectedCount != rejectedBefore;
                    break;
                }
                case PendingUpdate::Kind::DirectorySummaryChunk:
                {
                    const uint64_t rejectedBefore = _directorySummaryRejectedCount;
                    static_cast<void>(ApplyDirectorySummaryChunk(update));
                    layoutChanged = layoutChanged || _directorySummaryRejectedCount != rejectedBefore;
                    break;
                }
                case PendingUpdate::Kind::DirectorySummaryEnd:
                {
                    const uint64_t committedBefore = _directorySummaryCommittedCount;
                    const uint64_t rejectedBefore  = _directorySummaryRejectedCount;
                    static_cast<void>(EndDirectorySummary(update));
                    layoutChanged = layoutChanged || _directorySummaryCommittedCount != committedBefore;
                    layoutChanged = layoutChanged || _directorySummaryRejectedCount != rejectedBefore;
                    break;
                }
                case PendingUpdate::Kind::DirectorySummaryReject:
                {
                    const bool rejected = RejectDirectorySummary(DirectorySummaryStageKey{update.generation, update.nodeId, update.summaryRevision}, true);
                    layoutChanged       = layoutChanged || rejected;
                    if (update.nodeId == _model->rootNodeId && rejected)
                    {
                        headerTextDirty = true;
                    }
                    break;
                }
                case PendingUpdate::Kind::Progress:
                {
                    _scanProgressBytes        = std::max(_scanProgressBytes, update.bytes);
                    _scanProgressFolders      = std::max(_scanProgressFolders, update.scannedFolders);
                    _scanProgressFiles        = std::max(_scanProgressFiles, update.scannedFiles);
                    _scanFailedFolders        = std::max(_scanFailedFolders, update.failedFolders);
                    _scanPreviewPriorityJobs  = std::max(_scanPreviewPriorityJobs, update.firstPreviewPriorityJobs);
                    _scanProcessingNodeId     = update.nodeId;
                    _scanProcessingFolderName = std::move(update.name);

                    Node* root = TryGetRealNode(_model->rootNodeId);
                    if (root != nullptr && root->scanState != ScanState::Done)
                    {
                        // Progress updates can lag directory-completion updates in parallel scans.
                        // Keep the root total monotonic so a stale progress packet cannot downgrade
                        // the final authoritative size just before the root reaches Done.
                        static_cast<void>(RaiseNodeObservedBytes(root->id, update.bytes));
                    }

                    headerTextDirty = true;
                    break;
                }
                case PendingUpdate::Kind::ResourceSummary:
                {
                    _lastModelAcceptedEntries         = update.acceptedEntries;
                    _lastModelRejectedEntries         = update.rejectedEntries;
                    _lastModelCappedDirectories       = update.cappedDirectories;
                    _lastModelCappedFiles             = update.cappedFiles;
                    _lastModelRetainedNameBytes       = update.retainedNameBytes;
                    _lastModelRetainedChildReferences = update.retainedChildReferences;
                    _lastModelTraversedDirectories    = update.traversedDirectories;
                    _lastModelValidationError         = update.validationError;
                    Debug::Perf::EmitValue(L"viewer.space.model.child_arena_slots", _model->childrenArena.size());
                    Debug::Perf::EmitValue(L"viewer.space.model.child_arena_free_slots", ChildArenaFreeSlots());
                    uint64_t aggregateBytes = 0u;
                    for (const Node& node : _model->nodes)
                    {
                        aggregateBytes = ViewerSpaceScan::SaturatingAdd(aggregateBytes, node.aggregateBytes);
                    }
                    Debug::Perf::EmitValue(L"viewer.space.model.aggregate_bytes", aggregateBytes);
                    break;
                }
            }

        if (! keepQueuedBytes)
        {
            pendingBytesAfter = SubtractAtomicFloor(_pendingUpdateBytes, static_cast<uint64_t>(updateBytes));
            drainedUpdates += 1u;
        }

        if (processed >= 128u && (processed % 64u) == 0u)
        {
            const double elapsed = NowSeconds() - startSeconds;
            if (elapsed >= budgetSeconds)
            {
                break;
            }
        }
    }

    if (batchIndex < batch.size())
    {
        {
            std::scoped_lock lock(_updateMutex);
            for (size_t index = batch.size(); index-- > batchIndex;)
            {
                _pendingUpdates.emplace_front(std::move(batch[index]));
            }
            for (PendingUpdate& update : deferredAddChildren)
            {
                _pendingUpdates.emplace_back(std::move(update));
            }
        }
    }
    else if (! deferredAddChildren.empty())
    {
        std::scoped_lock lock(_updateMutex);
        for (PendingUpdate& update : deferredAddChildren)
        {
            _pendingUpdates.emplace_back(std::move(update));
        }
    }
    pendingBytesAfter = _pendingUpdateBytes.load(std::memory_order_relaxed);

    {
        const uint64_t committedBefore  = _directorySummaryCommittedCount;
        const uint64_t rejectedBefore   = _directorySummaryRejectedCount;
        const Node* rootBefore          = TryGetRealNode(_model->rootNodeId);
        const ScanState rootStateBefore = rootBefore == nullptr ? ScanState::NotStarted : rootBefore->scanState;
        static_cast<void>(AdvanceDirectorySummaryStages(ActiveDirectorySummaryCopyBudget()));
        layoutChanged = layoutChanged || _directorySummaryCommittedCount != committedBefore;
        layoutChanged = layoutChanged || _directorySummaryRejectedCount != rejectedBefore;
        ApplyDeferredDirectoryDoneStates();
        if (const Node* root = TryGetRealNode(_model->rootNodeId); root != nullptr && root->scanState != rootStateBefore)
        {
            headerTextDirty = true;
        }
    }

    if (applyingRefreshCandidate)
    {
        const Node* candidateRoot = TryGetRealNode(_model->rootNodeId);
        bool noQueuedUpdates      = false;
        {
            std::scoped_lock lock(_updateMutex);
            noQueuedUpdates = _pendingUpdates.empty() && batchIndex == batch.size() && deferredAddChildren.empty();
        }
        if (candidateRoot != nullptr && candidateRoot->scanState == ScanState::Done && noQueuedUpdates && _model->directorySummaryStages.empty())
        {
            _refreshAdoptionPending = true;
            headerTextDirty         = true;
        }
        else if (candidateRoot != nullptr && (candidateRoot->scanState == ScanState::Error || candidateRoot->scanState == ScanState::Canceled))
        {
            const bool failed = candidateRoot->scanState == ScanState::Error;
            _model            = priorModel;
            DiscardWorkingModel(false);
            _refreshFailed = failed;
            UpdateHeaderTextCache();
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
            }
            return;
        }
    }

    if (processed == 0u && ! layoutChanged && ! headerTextDirty && ! _refreshAdoptionPending)
    {
        if (IsQueueTerminalReady())
        {
            EmitQueueDiagnostics(true, false);
        }
        return;
    }

    const uint64_t drainUs = Debug::Perf::ElapsedUs(drainStartedAt);
    SaturatingAtomicAdd(_pendingUpdateDrainedCount, static_cast<uint64_t>(drainedUpdates));
    SaturatingAtomicAdd(_pendingUpdateDrainCount, 1u);
    SaturatingAtomicAdd(_pendingUpdateDrainTotalUs, drainUs);
    uint64_t observedDrainMaxUs = _pendingUpdateDrainMaxUs.load(std::memory_order_relaxed);
    while (drainUs > observedDrainMaxUs &&
           ! _pendingUpdateDrainMaxUs.compare_exchange_weak(observedDrainMaxUs, drainUs, std::memory_order_relaxed, std::memory_order_relaxed))
    {
    }
#ifdef ENABLE_TESTS
    _lastDrainUs = drainUs;
#endif

    const ScanState previousOverallState = _overallState;

    const bool wasScanActive = _scanActive.load();
    bool isScanActiveNow     = false;

    const Node* root = TryGetRealNode(_model->rootNodeId);
    if (root != nullptr)
    {
        _overallState = root->scanState;
        if (root->scanState == ScanState::Queued || root->scanState == ScanState::Scanning || root->scanState == ScanState::NotStarted)
        {
            isScanActiveNow = true;
        }
        else
        {
            isScanActiveNow = false;
        }
    }
    else
    {
        _overallState   = ScanState::NotStarted;
        isScanActiveNow = false;
    }

    _scanActive.store(isScanActiveNow);
    if (previousOverallState != _overallState)
    {
        if (_overallState == ScanState::Done)
        {
            _scanCompletedSinceSeconds = NowSeconds();
        }
        else
        {
            _scanCompletedSinceSeconds = 0.0;
        }
    }

    const bool terminalQueueSnapshot =
        ! applyingRefreshCandidate && (_overallState == ScanState::Done || _overallState == ScanState::Error || _overallState == ScanState::Canceled);
    if (terminalQueueSnapshot && IsQueueTerminalReady())
    {
        EmitQueueDiagnostics(true, false);
    }
    else
    {
        EmitQueueDiagnostics(false, false);
    }
    if (! applyingRefreshCandidate && wasScanActive != isScanActiveNow)
    {
        _layoutDirty              = true;
        _lastLayoutRebuildSeconds = 0.0;
        headerTextDirty           = true;
    }

    if (layoutChanged && ! applyingRefreshCandidate)
    {
        ++_modelRevision;
        InvalidateAnalysis();
        _layoutDirty = true;
        _layoutCandidateCache.clear();
    }

    if (! applyingRefreshCandidate && (layoutChanged || headerTextDirty || previousOverallState != _overallState || wasScanActive != isScanActiveNow))
    {
        InvalidateStaticTreemapCache();
    }

    if (headerTextDirty)
    {
        UpdateHeaderTextCache();
    }
    if (inspectorDirty && ! applyingRefreshCandidate)
    {
        UpdateInspector();
    }

    if (! applyingRefreshCandidate && (processed > 0u || headerTextDirty))
    {
        RefreshTooltipText();
    }

    if (_hWnd && (layoutChanged || headerTextDirty || wasScanActive != isScanActiveNow || _refreshAdoptionPending))
    {
        if (! isScanActiveNow || wasScanActive != isScanActiveNow)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
}

void ViewerSpace::CancelScanCacheBuild() noexcept
{
    _scanCacheBuildSnapshot.reset();
    _scanCacheBuildRootKey.clear();
    _scanCacheBuildTopFilesPerDirectory = 0;
    _scanCacheBuildGeneration           = 0;
    _scanCacheBuildPreflightNodesNext   = 0;
    _scanCacheBuildPreflightFilesNext   = 0;
    _scanCacheBuildNameChars            = 0;
    _scanCacheBuildChildrenNext         = 0;
    _scanCacheBuildNodesNext            = 0;
    _scanCacheBuildFileRecordsNext      = 0;
}

void ViewerSpace::ContinueScanCacheBuild() noexcept
{
    if (_refreshPending || _workingModel != nullptr)
    {
        CancelScanCacheBuild();
        return;
    }

    if (! _config.cacheEnabled || _model->scanRootPath.empty() || ! _fileSystemIsWin32)
    {
        CancelScanCacheBuild();
        return;
    }

    if (! g_cacheEnabled.load(std::memory_order_acquire) || g_cacheMaxEntries.load(std::memory_order_acquire) == 0)
    {
        CancelScanCacheBuild();
        return;
    }

    const uint32_t currentGeneration = _scanGeneration.load(std::memory_order_acquire);
    if (_overallState != ScanState::Done)
    {
        CancelScanCacheBuild();
        return;
    }

    if (_scanCacheBuildRootKey.empty() && _scanCacheLastStoredGeneration != currentGeneration)
    {
        ScanResultCacheKey cacheKey;
        cacheKey.rootKey              = NormalizeRootPathForScanCache(std::filesystem::path(_model->scanRootPath));
        cacheKey.topFilesPerDirectory = _scanTopFilesPerDirectory;

        if (! cacheKey.rootKey.empty())
        {
            _scanCacheBuildRootKey              = std::move(cacheKey.rootKey);
            _scanCacheBuildTopFilesPerDirectory = cacheKey.topFilesPerDirectory;
            _scanCacheBuildGeneration           = currentGeneration;
            _scanCacheBuildPreflightNodesNext   = 0;
            _scanCacheBuildPreflightFilesNext   = 0;
            _scanCacheBuildNameChars            = 0;
        }
    }

    if (_scanCacheBuildRootKey.empty())
    {
        return;
    }

    if (_scanCacheBuildGeneration != currentGeneration)
    {
        CancelScanCacheBuild();
        return;
    }

    static constexpr double kCacheBuildBudgetSeconds = 0.0012;
    const double startSeconds                        = NowSeconds();
    const uint64_t maxSnapshotBytes                  = MaxScanCacheSnapshotBytes();

    const auto fixedBytes = [this]() noexcept -> uint64_t
    {
        const uint64_t nodeBytes  = static_cast<uint64_t>(_model->nodes.size()) * static_cast<uint64_t>(sizeof(ScanResultCacheNode));
        const uint64_t fileBytes  = static_cast<uint64_t>(_model->fileRecords.size()) * static_cast<uint64_t>(sizeof(ScanResultCacheFileRecord));
        const uint64_t childBytes = static_cast<uint64_t>(_model->childrenArena.size()) * static_cast<uint64_t>(sizeof(uint32_t));
        if ((_model->nodes.size() != 0u && nodeBytes / sizeof(ScanResultCacheNode) != _model->nodes.size()) ||
            (_model->fileRecords.size() != 0u && fileBytes / sizeof(ScanResultCacheFileRecord) != _model->fileRecords.size()) ||
            (_model->childrenArena.size() != 0u && childBytes / sizeof(uint32_t) != _model->childrenArena.size()) ||
            nodeBytes > std::numeric_limits<uint64_t>::max() - fileBytes || childBytes > std::numeric_limits<uint64_t>::max() - sizeof(ScanResultSnapshot) ||
            nodeBytes + fileBytes > std::numeric_limits<uint64_t>::max() - sizeof(ScanResultSnapshot) - childBytes)
        {
            return std::numeric_limits<uint64_t>::max();
        }
        return nodeBytes + fileBytes + childBytes + static_cast<uint64_t>(sizeof(ScanResultSnapshot));
    };
    const auto abandonLargeSnapshot = [this, currentGeneration](uint64_t estimatedBytes) noexcept
    {
        _lastScanCacheSnapshotBytes = estimatedBytes;
        Debug::Perf::EmitValue(L"viewer.space.model.cache_skipped_large", estimatedBytes);
        _scanCacheLastStoredGeneration = currentGeneration;
        CancelScanCacheBuild();
    };
    const auto abandonViewerBudget = [this, currentGeneration](uint64_t estimatedBytes) noexcept
    {
        _lastScanCacheSnapshotBytes = estimatedBytes;
        Debug::Perf::EmitValue(L"viewer.space.model.cache_skipped_viewer_budget", estimatedBytes);
        _scanCacheLastStoredGeneration = currentGeneration;
        CancelScanCacheBuild();
    };

    if (! _scanCacheBuildSnapshot)
    {
        uint64_t estimate       = fixedBytes();
        const auto addNameChars = [&](size_t count) noexcept -> bool
        {
            if (count > std::numeric_limits<uint32_t>::max() - _scanCacheBuildNameChars)
            {
                estimate = std::numeric_limits<uint64_t>::max();
                return false;
            }
            _scanCacheBuildNameChars += count;
            const uint64_t nameBytes = static_cast<uint64_t>(_scanCacheBuildNameChars) * static_cast<uint64_t>(sizeof(wchar_t));
            if (nameBytes > std::numeric_limits<uint64_t>::max() - estimate)
            {
                estimate = std::numeric_limits<uint64_t>::max();
                return false;
            }
            estimate = fixedBytes() + nameBytes;
            return estimate <= maxSnapshotBytes;
        };

        while (_scanCacheBuildPreflightNodesNext < _model->nodes.size())
        {
            if (! addNameChars(_model->nodes[_scanCacheBuildPreflightNodesNext].name.size()))
            {
                abandonLargeSnapshot(estimate);
                return;
            }
            ++_scanCacheBuildPreflightNodesNext;
            if ((_scanCacheBuildPreflightNodesNext % 256u) == 0u && (NowSeconds() - startSeconds) >= kCacheBuildBudgetSeconds)
            {
                return;
            }
        }

        while (_scanCacheBuildPreflightFilesNext < _model->fileRecords.size())
        {
            if (! addNameChars(_model->fileRecords[_scanCacheBuildPreflightFilesNext].name.size()))
            {
                abandonLargeSnapshot(estimate);
                return;
            }
            ++_scanCacheBuildPreflightFilesNext;
            if ((_scanCacheBuildPreflightFilesNext % 256u) == 0u && (NowSeconds() - startSeconds) >= kCacheBuildBudgetSeconds)
            {
                return;
            }
        }

        const uint64_t estimatedSnapshotBytes = fixedBytes() + static_cast<uint64_t>(_scanCacheBuildNameChars) * static_cast<uint64_t>(sizeof(wchar_t));
        const size_t keyCapacity              = _scanCacheBuildRootKey.capacity();
        const uint64_t keyBytes               = keyCapacity == std::numeric_limits<size_t>::max()
                                                    ? std::numeric_limits<uint64_t>::max()
                                                    : static_cast<uint64_t>(keyCapacity + 1u) * static_cast<uint64_t>(sizeof(wchar_t));
        if (keyBytes > std::numeric_limits<uint64_t>::max() - estimatedSnapshotBytes ||
            ! FitsViewerSpaceCacheWorkingSet(ViewerOwnedBytes(), estimatedSnapshotBytes + keyBytes))
        {
            abandonViewerBudget(estimatedSnapshotBytes);
            return;
        }

        auto snapshot                      = std::make_shared<ScanResultSnapshot>();
        snapshot->topFilesPerDirectory     = _scanCacheBuildTopFilesPerDirectory;
        snapshot->scannedBytes             = _scanProgressBytes;
        snapshot->scannedFolders           = _scanProgressFolders;
        snapshot->scannedFiles             = _scanProgressFiles;
        snapshot->failedFolders            = _scanFailedFolders;
        snapshot->firstPreviewPriorityJobs = _scanPreviewPriorityJobs;
        snapshot->names.reserve(_scanCacheBuildNameChars);
        snapshot->nodes.reserve(_model->nodes.size());
        snapshot->fileRecords.reserve(_model->fileRecords.size());
        snapshot->childrenArena.reserve(_model->childrenArena.size());
        const uint64_t actualSnapshotBytes = snapshot->AllocatedCapacityBytes();
        if (keyBytes > std::numeric_limits<uint64_t>::max() - actualSnapshotBytes ||
            ! FitsViewerSpaceCacheWorkingSet(ViewerOwnedBytes(), actualSnapshotBytes + keyBytes))
        {
            abandonViewerBudget(actualSnapshotBytes);
            return;
        }
        _scanCacheBuildSnapshot        = std::move(snapshot);
        _scanCacheBuildChildrenNext    = 0;
        _scanCacheBuildNodesNext       = 0;
        _scanCacheBuildFileRecordsNext = 0;
    }

    auto snapshot = std::static_pointer_cast<ScanResultSnapshot>(_scanCacheBuildSnapshot);
    if (! snapshot)
    {
        CancelScanCacheBuild();
        return;
    }

    const size_t keyCapacity            = _scanCacheBuildRootKey.capacity();
    const uint64_t keyBytes             = keyCapacity == std::numeric_limits<size_t>::max()
                                              ? std::numeric_limits<uint64_t>::max()
                                              : static_cast<uint64_t>(keyCapacity + 1u) * static_cast<uint64_t>(sizeof(wchar_t));
    const uint64_t currentSnapshotBytes = snapshot->AllocatedCapacityBytes();
    if (keyBytes > std::numeric_limits<uint64_t>::max() - currentSnapshotBytes ||
        ! FitsViewerSpaceCacheWorkingSet(ViewerOwnedBytes(), currentSnapshotBytes + keyBytes))
    {
        abandonViewerBudget(currentSnapshotBytes);
        return;
    }

    const size_t childCount = _model->childrenArena.size();
    while (_scanCacheBuildChildrenNext < childCount)
    {
        snapshot->childrenArena.push_back(_model->childrenArena[_scanCacheBuildChildrenNext]);
        _scanCacheBuildChildrenNext += 1;

        if ((_scanCacheBuildChildrenNext % 256u) == 0u)
        {
            if ((NowSeconds() - startSeconds) >= kCacheBuildBudgetSeconds)
            {
                return;
            }
        }
    }

    const size_t nodeCount = _model->nodes.size();
    while (_scanCacheBuildNodesNext < nodeCount)
    {
        const Node& node = _model->nodes[_scanCacheBuildNodesNext];

        ScanResultCacheNode cachedNode;
        if (node.id != 0)
        {
            cachedNode.id                   = node.id;
            cachedNode.parentId             = node.parentId;
            cachedNode.isDirectory          = node.isDirectory;
            cachedNode.isSynthetic          = node.isSynthetic;
            cachedNode.scanState            = static_cast<uint8_t>(node.scanState);
            cachedNode.flags                = node.contentNotMeasured ? kScanResultCacheNodeContentNotMeasured : 0u;
            cachedNode.totalBytes           = node.totalBytes;
            cachedNode.lastWriteTime        = node.lastWriteTime;
            cachedNode.observedOwnFileBytes = node.observedOwnFileBytes;
            cachedNode.observedChildBytes   = node.observedChildBytes;
            cachedNode.childrenStart        = node.childrenStart;
            cachedNode.childrenCount        = node.childrenCount;
            cachedNode.childrenCapacity     = node.childrenCapacity;
            cachedNode.aggregateBytes       = node.aggregateBytes;
            cachedNode.aggregateFolders     = node.aggregateFolders;
            cachedNode.aggregateFiles       = node.aggregateFiles;
            if (snapshot->names.size() > std::numeric_limits<uint32_t>::max() ||
                node.name.size() > std::numeric_limits<uint32_t>::max() - snapshot->names.size())
            {
                abandonLargeSnapshot(std::numeric_limits<uint64_t>::max());
                return;
            }
            cachedNode.nameOffset = static_cast<uint32_t>(snapshot->names.size());
            cachedNode.nameLength = static_cast<uint32_t>(node.name.size());
            snapshot->names.insert(snapshot->names.end(), node.name.begin(), node.name.end());
        }

        snapshot->nodes.push_back(std::move(cachedNode));
        _scanCacheBuildNodesNext += 1;

        if ((_scanCacheBuildNodesNext % 64u) == 0u)
        {
            if ((NowSeconds() - startSeconds) >= kCacheBuildBudgetSeconds)
            {
                return;
            }
        }
    }

    if (snapshot->nodes.size() != nodeCount)
    {
        CancelScanCacheBuild();
        return;
    }

    const size_t fileRecordCount = _model->fileRecords.size();
    while (_scanCacheBuildFileRecordsNext < fileRecordCount)
    {
        const FileRecord& file = _model->fileRecords[_scanCacheBuildFileRecordsNext];

        ScanResultCacheFileRecord cachedFile;
        if (file.id != 0)
        {
            cachedFile.id            = file.id;
            cachedFile.parentId      = file.parentId;
            cachedFile.bytes         = file.bytes;
            cachedFile.lastWriteTime = file.lastWriteTime;
            if (snapshot->names.size() > std::numeric_limits<uint32_t>::max() ||
                file.name.size() > std::numeric_limits<uint32_t>::max() - snapshot->names.size())
            {
                abandonLargeSnapshot(std::numeric_limits<uint64_t>::max());
                return;
            }
            cachedFile.nameOffset = static_cast<uint32_t>(snapshot->names.size());
            cachedFile.nameLength = static_cast<uint32_t>(file.name.size());
            snapshot->names.insert(snapshot->names.end(), file.name.begin(), file.name.end());
        }

        snapshot->fileRecords.push_back(std::move(cachedFile));
        _scanCacheBuildFileRecordsNext += 1;

        if ((_scanCacheBuildFileRecordsNext % 64u) == 0u)
        {
            if ((NowSeconds() - startSeconds) >= kCacheBuildBudgetSeconds)
            {
                return;
            }
        }
    }

    if (snapshot->fileRecords.size() != fileRecordCount)
    {
        CancelScanCacheBuild();
        return;
    }

    ScanResultCacheKey cacheKey;
    cacheKey.rootKey              = _scanCacheBuildRootKey;
    cacheKey.topFilesPerDirectory = _scanCacheBuildTopFilesPerDirectory;

    uint64_t snapshotBytes      = snapshot->AllocatedCapacityBytes();
    _lastScanCacheSnapshotBytes = snapshotBytes;
    Debug::Perf::EmitValue(L"viewer.space.model.cache_snapshot_bytes", snapshotBytes);
    if (snapshotBytes > MaxScanCacheSnapshotBytes())
    {
        Debug::Perf::EmitValue(L"viewer.space.model.cache_skipped_large", snapshotBytes);
        _scanCacheLastStoredGeneration = _scanCacheBuildGeneration;
        CancelScanCacheBuild();
        return;
    }

    ScanResultCache& cache = GetScanResultCache();
    if (! cache.Store(std::move(cacheKey), std::move(snapshot)) && g_cacheEnabled.load(std::memory_order_acquire) &&
        g_cacheMaxEntries.load(std::memory_order_acquire) != 0u)
    {
        Debug::Perf::EmitValue(L"viewer.space.model.cache_skipped_aggregate", snapshotBytes);
    }
    _lastScanCacheResidentBytes = cache.ResidentBytes();
    Debug::Perf::EmitValue(L"viewer.space.model.cache_resident_bytes", _lastScanCacheResidentBytes);
    _scanCacheLastStoredGeneration = _scanCacheBuildGeneration;
    CancelScanCacheBuild();
}

void ViewerSpace::CancelScanCacheRestore() noexcept
{
    _scanCacheRestoreSnapshot.reset();
    _scanCacheRestoreRootKey.clear();
    _scanCacheRestoreTopFilesPerDirectory  = 0u;
    _scanCacheRestoreGeneration            = 0u;
    _scanCacheRestorePhase                 = static_cast<uint8_t>(ScanCacheRestorePhase::None);
    _scanCacheRestorePaintSliceStarted     = false;
    _scanCacheRestoreContinuationAttempted = false;
    // A canceled generation keeps its single queued token owned until dispatch or HWND drain.
    // Keeping pending set prevents a replacement generation from accumulating another token.
    _scanCacheRestoreChildrenNext        = 0u;
    _scanCacheRestoreNodesNext           = 0u;
    _scanCacheRestoreFileRecordsNext     = 0u;
    _scanCacheRestoreValidationNext      = 0u;
    _scanCacheRestoreValidationChildNext = 0u;
    std::vector<uint64_t>().swap(_scanCacheRestoreSeenChildren);
    _scanCacheRestoreStartedAtSeconds = 0.0;
}

void ViewerSpace::QueueScanCacheRestoreContinuation(uint32_t generation) noexcept
{
    if (! _hWnd || ! _scanCacheRestoreSnapshot || _scanCacheRestoreContinuationAttempted || _scanCacheRestoreContinuationPending ||
        generation != _scanCacheRestoreGeneration || generation != _scanGeneration.load(std::memory_order_acquire) ||
        ! FitsViewerSpaceCacheWorkingSet(ViewerOwnedBytes(), sizeof(ScanCacheRestoreContinuation)))
    {
        return;
    }
    // Only one extra dispatch is admitted per restore. Later turns retain the normal timer;
    // never self-post a loop that could hold posted-message priority over real input.
    _scanCacheRestoreContinuationAttempted = true;
    auto continuation                      = std::make_unique<ScanCacheRestoreContinuation>();
    continuation->generation               = generation;
    _scanCacheRestoreContinuationPending   = PostMessagePayload(_hWnd.get(), WndMsg::kViewerSpaceCacheRestoreContinue, 0u, std::move(continuation));
    Debug::Perf::EmitCounter(_scanCacheRestoreContinuationPending ? L"viewer.space.model.cache_restore_continuation_posted"
                                                                  : L"viewer.space.model.cache_restore_continuation_post_failed");
}

void ViewerSpace::ContinueScanCacheRestore() noexcept
{
    if (! _scanCacheRestoreSnapshot)
    {
        return;
    }

    static constexpr double kCacheRestoreBudgetSeconds = 0.0012;
    static constexpr uint32_t kCacheRestorePhaseCount  = static_cast<uint32_t>(ScanCacheRestorePhase::Files);
    const double startSeconds                          = NowSeconds();
    const uint32_t currentGeneration                   = _scanGeneration.load(std::memory_order_acquire);
    auto snapshot                                      = std::static_pointer_cast<const ScanResultSnapshot>(_scanCacheRestoreSnapshot);
    uint32_t completedPhasesThisTurn                   = 0u;
    uint32_t phaseLimitThisTurn                        = kCacheRestorePhaseCount;
    std::wstring_view turnOutcome                      = L"yield";
    HRESULT turnResult                                 = S_OK;
    const bool captureTurnTelemetry                    = Debug::Perf::IsCaptureEnabled();
#ifdef ENABLE_TESTS
    if (const auto testPhaseLimit = ReadPositiveUint32Environment(L"REDSALAMANDER_VIEWERSPACE_TEST_CACHE_RESTORE_PHASES_PER_TURN"); testPhaseLimit.has_value())
    {
        phaseLimitThisTurn = std::min(testPhaseLimit.value(), kCacheRestorePhaseCount);
    }
#endif
    const auto emitTurnTelemetry = wil::scope_exit([&]() noexcept
    {
        if (! captureTurnTelemetry)
        {
            return;
        }
        const uint64_t elapsedUs = static_cast<uint64_t>(std::max(0.0, (NowSeconds() - startSeconds) * 1'000'000.0));
        Debug::Perf::Emit(L"viewer.space.model.cache_restore_turn_us",
                          turnOutcome,
                          elapsedUs,
                          completedPhasesThisTurn,
                          static_cast<uint64_t>(_scanCacheRestorePhase),
                          turnResult);
        if (completedPhasesThisTurn > 1u)
        {
            Debug::Perf::EmitValue(L"viewer.space.model.cache_restore_batched_turn", completedPhasesThisTurn);
        }
    });
    if (! snapshot || _scanCacheRestoreGeneration != currentGeneration || _overallState != ScanState::Queued)
    {
        turnOutcome = L"stale";
        turnResult  = E_ABORT;
        CancelScanCacheRestore();
        return;
    }

    const auto setPhase   = [this](ScanCacheRestorePhase next) noexcept { _scanCacheRestorePhase = static_cast<uint8_t>(next); };
    const auto failClosed = [this, &snapshot, &turnOutcome, &turnResult]() noexcept
    {
        if (turnOutcome == L"yield")
        {
            turnOutcome = L"invalid-snapshot";
        }
        turnResult = E_FAIL;
        ScanResultCacheKey key;
        key.rootKey              = _scanCacheRestoreRootKey;
        key.topFilesPerDirectory = _scanCacheRestoreTopFilesPerDirectory;
        GetScanResultCache().DiscardIfMatches(key, snapshot);
        Debug::Perf::EmitCounter(L"viewer.space.scan.cache_invalid_snapshot");
        const std::wstring rootPath = _model->scanRootPath;
        CancelScanCacheRestore();
        snapshot.reset();
        StartScan(rootPath, false);
    };

    const auto isTimeBudgetExhausted    = [this, startSeconds]() noexcept { return (NowSeconds() - startSeconds) >= kCacheRestoreBudgetSeconds; };
    const size_t restoreKeyCapacity     = _scanCacheRestoreRootKey.capacity();
    const uint64_t restoreKeyBytes      = restoreKeyCapacity == std::numeric_limits<size_t>::max()
                                              ? std::numeric_limits<uint64_t>::max()
                                              : static_cast<uint64_t>(restoreKeyCapacity + 1u) * static_cast<uint64_t>(sizeof(wchar_t));
    const uint64_t compactSnapshotBytes = snapshot->AllocatedCapacityBytes();
    if (restoreKeyBytes > std::numeric_limits<uint64_t>::max() - compactSnapshotBytes)
    {
        failClosed();
        return;
    }
    const uint64_t sourceSnapshotBytes = compactSnapshotBytes + restoreKeyBytes;
    const auto sourceSnapshotFits = [this, sourceSnapshotBytes]() noexcept { return FitsViewerSpaceCacheWorkingSet(ViewerOwnedBytes(), sourceSnapshotBytes); };
    if (! sourceSnapshotFits())
    {
        failClosed();
        return;
    }

    for (;;)
    {
        if (completedPhasesThisTurn >= phaseLimitThisTurn || (completedPhasesThisTurn > 0u && isTimeBudgetExhausted()))
        {
            return;
        }

        const auto phase = static_cast<ScanCacheRestorePhase>(_scanCacheRestorePhase);
        switch (phase)
        {
            case ScanCacheRestorePhase::ValidateNodes:
            {
                const uint64_t maximumNodes = static_cast<uint64_t>(_scanResourcePolicy.maxRetainedDirectories) + 1u;
                if (snapshot->nodes.size() < 2u || static_cast<uint64_t>(snapshot->nodes.size()) > maximumNodes ||
                    snapshot->childrenArena.size() > _scanResourcePolicy.maxChildArenaSlots)
                {
                    failClosed();
                    return;
                }

                size_t checked = 0u;
                while (_scanCacheRestoreValidationNext < snapshot->nodes.size())
                {
                    const size_t index              = _scanCacheRestoreValidationNext;
                    const ScanResultCacheNode& node = snapshot->nodes[index];
                    if (index == 0u)
                    {
                        if (node.id != 0u || node.parentId != 0u || node.isDirectory || node.isSynthetic || node.scanState != 0u || node.flags != 0u ||
                            node.totalBytes != 0u || node.childrenStart != 0u || node.childrenCount != 0u || node.childrenCapacity != 0u ||
                            node.aggregateBytes != 0u || node.aggregateFolders != 0u || node.aggregateFiles != 0u || node.nameOffset != 0u ||
                            node.nameLength != 0u)
                        {
                            failClosed();
                            return;
                        }
                    }
                    else if (node.id == 0u
                                 ? (node.parentId != 0u || node.isDirectory || node.isSynthetic || node.scanState != 0u || node.flags != 0u ||
                                    node.totalBytes != 0u || node.childrenStart != 0u || node.childrenCount != 0u || node.childrenCapacity != 0u ||
                                    node.aggregateBytes != 0u || node.aggregateFolders != 0u || node.aggregateFiles != 0u || node.nameOffset != 0u ||
                                    node.nameLength != 0u)
                                 : (static_cast<size_t>(node.id) != index || node.parentId >= index || node.parentId >= snapshot->nodes.size() ||
                                    node.scanState > static_cast<uint8_t>(ScanState::Canceled) || (node.flags & ~kScanResultCacheNodeKnownFlags) != 0u ||
                                    ! IsValidSnapshotNameExtent(*snapshot, node.nameOffset, node.nameLength) ||
                                    node.childrenStart > snapshot->childrenArena.size() ||
                                    node.childrenCapacity > snapshot->childrenArena.size() - node.childrenStart || node.childrenCount > node.childrenCapacity))
                    {
                        failClosed();
                        return;
                    }
                    else if (node.id != 0u &&
                             (index == 1u ? (! node.isDirectory || node.parentId != 0u || node.scanState != static_cast<uint8_t>(ScanState::Done))
                                          : (node.parentId == 0u || ! snapshot->nodes[node.parentId].isDirectory ||
                                             snapshot->nodes[node.parentId].id != node.parentId)))
                    {
                        failClosed();
                        return;
                    }

                    ++_scanCacheRestoreValidationNext;
                    if (++checked >= 64u && isTimeBudgetExhausted())
                    {
                        return;
                    }
                }
                _scanCacheRestoreValidationNext = 0u;
                setPhase(ScanCacheRestorePhase::ValidateFiles);
                ++completedPhasesThisTurn;
                continue;
            }

            case ScanCacheRestorePhase::ValidateFiles:
            {
                if (snapshot->fileRecords.size() > _scanResourcePolicy.maxRetainedFileRecords ||
                    snapshot->fileRecords.size() > static_cast<size_t>(kTreemapItemIdMask) + 1u)
                {
                    failClosed();
                    return;
                }

                size_t checked = 0u;
                while (_scanCacheRestoreValidationNext < snapshot->fileRecords.size())
                {
                    const size_t index                    = _scanCacheRestoreValidationNext;
                    const ScanResultCacheFileRecord& file = snapshot->fileRecords[index];
                    if (file.id == 0u ? (file.parentId != 0u || file.bytes != 0u || file.nameOffset != 0u || file.nameLength != 0u)
                                      : (file.id != FileRecordIdFromIndex(index) || file.parentId == 0u || file.parentId >= snapshot->nodes.size() ||
                                         ! snapshot->nodes[file.parentId].isDirectory || snapshot->nodes[file.parentId].id != file.parentId ||
                                         ! IsValidSnapshotNameExtent(*snapshot, file.nameOffset, file.nameLength)))
                    {
                        failClosed();
                        return;
                    }

                    ++_scanCacheRestoreValidationNext;
                    if (++checked >= 64u && isTimeBudgetExhausted())
                    {
                        return;
                    }
                }
                _scanCacheRestoreValidationNext      = 1u;
                _scanCacheRestoreValidationChildNext = 0u;
                setPhase(ScanCacheRestorePhase::ValidateChildren);
                ++completedPhasesThisTurn;
                continue;
            }

            case ScanCacheRestorePhase::ValidateChildren:
            {
                if (_scanCacheRestoreSeenChildren.empty())
                {
                    if (snapshot->nodes.size() > std::numeric_limits<size_t>::max() - snapshot->fileRecords.size())
                    {
                        failClosed();
                        return;
                    }
                    const size_t bitCount = snapshot->nodes.size() + snapshot->fileRecords.size();
                    if (bitCount > std::numeric_limits<size_t>::max() - 63u)
                    {
                        failClosed();
                        return;
                    }
                    const size_t wordCount = (bitCount + 63u) / 64u;
                    if (wordCount > std::numeric_limits<uint64_t>::max() / sizeof(uint64_t))
                    {
                        failClosed();
                        return;
                    }
                    const uint64_t bitmapBytes = static_cast<uint64_t>(wordCount) * sizeof(uint64_t);
                    if (sourceSnapshotBytes > std::numeric_limits<uint64_t>::max() - bitmapBytes ||
                        ! FitsViewerSpaceCacheWorkingSet(ViewerOwnedBytes(), sourceSnapshotBytes + bitmapBytes))
                    {
                        failClosed();
                        return;
                    }
                    _scanCacheRestoreSeenChildren.assign(wordCount, 0u);
                    if (! sourceSnapshotFits())
                    {
                        failClosed();
                        return;
                    }
                }

                size_t checked = 0u;
                while (_scanCacheRestoreValidationNext < snapshot->nodes.size())
                {
                    const ScanResultCacheNode& node = snapshot->nodes[_scanCacheRestoreValidationNext];
                    while (_scanCacheRestoreValidationChildNext < node.childrenCount)
                    {
                        const size_t childOffset  = static_cast<size_t>(node.childrenStart) + _scanCacheRestoreValidationChildNext;
                        const uint32_t childId    = snapshot->childrenArena[childOffset];
                        const bool validNodeChild = childId != 0u && childId < snapshot->nodes.size() && snapshot->nodes[childId].id == childId;
                        const size_t fileIndex = childId >= kFileRecordIdBase ? static_cast<size_t>(childId - kFileRecordIdBase) : snapshot->fileRecords.size();
                        const bool validFileChild = fileIndex < snapshot->fileRecords.size() && snapshot->fileRecords[fileIndex].id == childId;
                        if (node.id == 0u || (! validNodeChild && ! validFileChild) || (validNodeChild && snapshot->nodes[childId].parentId != node.id) ||
                            (validFileChild && snapshot->fileRecords[fileIndex].parentId != node.id))
                        {
                            failClosed();
                            return;
                        }
                        const size_t seenSlot = validNodeChild ? static_cast<size_t>(childId) : snapshot->nodes.size() + fileIndex;
                        const size_t seenWord = seenSlot / 64u;
                        if (seenWord >= _scanCacheRestoreSeenChildren.size())
                        {
                            failClosed();
                            return;
                        }
                        const uint64_t seenMask = uint64_t{1} << (seenSlot % 64u);
                        if ((_scanCacheRestoreSeenChildren[seenWord] & seenMask) != 0u)
                        {
                            failClosed();
                            return;
                        }
                        _scanCacheRestoreSeenChildren[seenWord] |= seenMask;
                        ++_scanCacheRestoreValidationChildNext;
                        if (++checked >= 128u && isTimeBudgetExhausted())
                        {
                            return;
                        }
                    }
                    _scanCacheRestoreValidationChildNext = 0u;
                    ++_scanCacheRestoreValidationNext;
                    // Empty directories still consume validation work; bound this phase by nodes as well as child references.
                    if (++checked >= 128u && isTimeBudgetExhausted())
                    {
                        return;
                    }
                }
                _scanCacheRestoreValidationNext = 0u;
                setPhase(ScanCacheRestorePhase::ValidateCompleteness);
                ++completedPhasesThisTurn;
                continue;
            }

            case ScanCacheRestorePhase::ValidateCompleteness:
            {
                const size_t nodeCount   = snapshot->nodes.size();
                const size_t recordCount = nodeCount + snapshot->fileRecords.size(); // Addition was checked when the bitmap was allocated.
                size_t checked           = 0u;
                while (_scanCacheRestoreValidationNext < recordCount)
                {
                    const size_t slot = _scanCacheRestoreValidationNext++;
                    if (slot <= 1u)
                    {
                        continue; // The sentinel and root have no incoming child references.
                    }
                    const bool isNode = slot < nodeCount;
                    const bool active = isNode ? snapshot->nodes[slot].id != 0u : snapshot->fileRecords[slot - nodeCount].id != 0u;
                    if (active)
                    {
                        const size_t seenWord   = slot / 64u;
                        const uint64_t seenMask = uint64_t{1} << (slot % 64u);
                        if (seenWord >= _scanCacheRestoreSeenChildren.size() || (_scanCacheRestoreSeenChildren[seenWord] & seenMask) == 0u)
                        {
                            failClosed();
                            return;
                        }
                    }
                    if (++checked >= 128u && isTimeBudgetExhausted())
                    {
                        return;
                    }
                }
                _scanCacheRestoreValidationNext = 0u;
                setPhase(ScanCacheRestorePhase::PrepareChildren);
                ++completedPhasesThisTurn;
                continue;
            }

            case ScanCacheRestorePhase::PrepareChildren:
            {
                const uint64_t targetNameUnits = static_cast<uint64_t>(snapshot->names.size()) + static_cast<uint64_t>(snapshot->nodes.size()) +
                                                 static_cast<uint64_t>(snapshot->fileRecords.size());
                const uint64_t targetBytes     = static_cast<uint64_t>(snapshot->nodes.size()) * static_cast<uint64_t>(sizeof(Node)) +
                                                 static_cast<uint64_t>(snapshot->fileRecords.size()) * static_cast<uint64_t>(sizeof(FileRecord)) +
                                                 static_cast<uint64_t>(snapshot->childrenArena.size()) * static_cast<uint64_t>(sizeof(uint32_t)) +
                                                 targetNameUnits * static_cast<uint64_t>(sizeof(wchar_t));
                if (sourceSnapshotBytes > std::numeric_limits<uint64_t>::max() - targetBytes ||
                    ! FitsViewerSpaceCacheWorkingSet(ViewerOwnedBytes(), sourceSnapshotBytes + targetBytes))
                {
                    failClosed();
                    return;
                }
                _model->childrenArena.reserve(snapshot->childrenArena.size());
                setPhase(ScanCacheRestorePhase::PrepareNodes);
                ++completedPhasesThisTurn;
                continue;
            }

            case ScanCacheRestorePhase::PrepareNodes:
                _model->nodes.reserve(std::max<size_t>(snapshot->nodes.size(), 2u));
                if (! sourceSnapshotFits())
                {
                    failClosed();
                    return;
                }
                setPhase(ScanCacheRestorePhase::PrepareFiles);
                ++completedPhasesThisTurn;
                continue;

            case ScanCacheRestorePhase::PrepareFiles:
                _model->fileRecords.reserve(snapshot->fileRecords.size());
                if (! sourceSnapshotFits())
                {
                    failClosed();
                    return;
                }
                _scanCacheRestoreChildrenNext = 0u;
                setPhase(ScanCacheRestorePhase::Children);
                ++completedPhasesThisTurn;
                continue;

            case ScanCacheRestorePhase::Children:
            {
                size_t checked = 0u;
                while (_scanCacheRestoreChildrenNext < snapshot->childrenArena.size())
                {
                    _model->childrenArena.push_back(snapshot->childrenArena[_scanCacheRestoreChildrenNext]);
                    ++_scanCacheRestoreChildrenNext;
                    if (++checked >= 256u && isTimeBudgetExhausted())
                    {
                        if (! sourceSnapshotFits())
                        {
                            failClosed();
                            return;
                        }
                        return;
                    }
                }
                _scanCacheRestoreNodesNext = 0u;
                setPhase(ScanCacheRestorePhase::Nodes);
                ++completedPhasesThisTurn;
                continue;
            }

            case ScanCacheRestorePhase::Nodes:
            {
                size_t checked = 0u;
                while (_scanCacheRestoreNodesNext < snapshot->nodes.size())
                {
                    const size_t index                    = _scanCacheRestoreNodesNext;
                    const ScanResultCacheNode& cachedNode = snapshot->nodes[index];
                    if (index >= _model->nodes.size())
                    {
                        _model->nodes.emplace_back();
                    }
                    if (cachedNode.id != 0u)
                    {
                        const std::wstring_view name = ScanSnapshotName(*snapshot, cachedNode.nameOffset, cachedNode.nameLength);
                        Node node;
                        node.id                   = cachedNode.id;
                        node.parentId             = cachedNode.parentId;
                        node.isDirectory          = cachedNode.isDirectory;
                        node.isSynthetic          = cachedNode.isSynthetic;
                        node.contentNotMeasured   = (cachedNode.flags & kScanResultCacheNodeContentNotMeasured) != 0u;
                        node.scanState            = cachedNode.id == _model->rootNodeId ? ScanState::Queued : static_cast<ScanState>(cachedNode.scanState);
                        node.name                 = CopyToArena(_model->nameArena, name);
                        node.totalBytes           = cachedNode.totalBytes;
                        node.lastWriteTime        = cachedNode.lastWriteTime;
                        node.observedOwnFileBytes = cachedNode.observedOwnFileBytes;
                        node.observedChildBytes   = cachedNode.observedChildBytes;
                        node.childrenStart        = cachedNode.childrenStart;
                        node.childrenCount        = cachedNode.childrenCount;
                        node.childrenCapacity     = cachedNode.childrenCapacity;
                        node.aggregateBytes       = cachedNode.aggregateBytes;
                        node.aggregateFolders     = cachedNode.aggregateFolders;
                        node.aggregateFiles       = cachedNode.aggregateFiles;
                        _model->nodes[index]      = std::move(node);
                    }
                    ++_scanCacheRestoreNodesNext;
                    if (++checked >= 32u && isTimeBudgetExhausted())
                    {
                        if (! sourceSnapshotFits())
                        {
                            failClosed();
                            return;
                        }
                        return;
                    }
                }
                _scanCacheRestoreFileRecordsNext = 0u;
                setPhase(ScanCacheRestorePhase::Files);
                ++completedPhasesThisTurn;
                continue;
            }

            case ScanCacheRestorePhase::Files:
            {
                size_t checked = 0u;
                while (_scanCacheRestoreFileRecordsNext < snapshot->fileRecords.size())
                {
                    const size_t index                          = _scanCacheRestoreFileRecordsNext;
                    const ScanResultCacheFileRecord& cachedFile = snapshot->fileRecords[index];
                    if (index >= _model->fileRecords.size())
                    {
                        _model->fileRecords.emplace_back();
                    }
                    FileRecord record;
                    record.id                  = cachedFile.id;
                    record.parentId            = cachedFile.parentId;
                    record.bytes               = cachedFile.bytes;
                    record.lastWriteTime       = cachedFile.lastWriteTime;
                    record.name                = CopyToArena(_model->nameArena, ScanSnapshotName(*snapshot, cachedFile.nameOffset, cachedFile.nameLength));
                    _model->fileRecords[index] = std::move(record);
                    ++_scanCacheRestoreFileRecordsNext;
                    if (++checked >= 32u && isTimeBudgetExhausted())
                    {
                        if (! sourceSnapshotFits())
                        {
                            failClosed();
                            return;
                        }
                        return;
                    }
                }

                _scanTopFilesPerDirectory = std::max(_scanTopFilesPerDirectory, _scanCacheRestoreTopFilesPerDirectory);
                _scanProgressBytes        = snapshot->scannedBytes;
                _scanProgressFolders      = snapshot->scannedFolders;
                _scanProgressFiles        = snapshot->scannedFiles;
                _scanFailedFolders        = snapshot->failedFolders;
                _scanPreviewPriorityJobs  = snapshot->firstPreviewPriorityJobs;
                UpdateViewPathText();
                _model->nodes[1].scanState = ScanState::Done;
                _overallState              = ScanState::Done;
                ++_modelRevision;
                _layoutDirty = true;
                InvalidateAnalysis();
                _scanActive.store(false, std::memory_order_release);
                _animationStartSeconds      = NowSeconds();
                _lastLayoutRebuildSeconds   = 0.0;
                _lastScanCacheRestoreBytes  = static_cast<uint64_t>(_model->nodes.size()) * static_cast<uint64_t>(sizeof(Node)) +
                                              static_cast<uint64_t>(_model->fileRecords.size()) * static_cast<uint64_t>(sizeof(FileRecord)) +
                                              static_cast<uint64_t>(_model->childrenArena.size()) * static_cast<uint64_t>(sizeof(uint32_t)) +
                                              static_cast<uint64_t>(snapshot->names.size()) * static_cast<uint64_t>(sizeof(wchar_t));
                _lastScanCacheRestoreUs     = static_cast<uint64_t>(std::max(0.0, (NowSeconds() - _scanCacheRestoreStartedAtSeconds) * 1'000'000.0));
                _lastScanCacheResidentBytes = GetScanResultCache().ResidentBytes();
                Debug::Perf::EmitCounter(L"viewer.space.scan.cache_hit");
                Debug::Perf::EmitValue(L"viewer.space.model.cache_restore_bytes", _lastScanCacheRestoreBytes);
                Debug::Perf::EmitValue(L"viewer.space.model.cache_restore_us", _lastScanCacheRestoreUs);
                Debug::Perf::EmitValue(L"viewer.space.model.cache_resident_bytes", _lastScanCacheResidentBytes);
                _scanCacheLastStoredGeneration = _scanCacheRestoreGeneration;
                ++completedPhasesThisTurn;
                turnOutcome = L"complete";
                CancelScanCacheRestore();
                UpdateHeaderTextCache();
                if (_hWnd)
                {
                    UpdateWindowTitle(_hWnd.get());
                    UpdateMenuState(_hWnd.get());
                    InvalidateRect(_hWnd.get(), nullptr, FALSE);
                }
                return;
            }

            case ScanCacheRestorePhase::None:
            default:
                turnOutcome = L"invalid-phase";
                turnResult  = E_FAIL;
                failClosed();
                return;
        }
    }
}

size_t ViewerSpace::EstimatePendingUpdateBytes(const PendingUpdate& update) noexcept
{
    const auto addSaturated = [](size_t left, size_t right) noexcept
    { return right > (std::numeric_limits<size_t>::max)() - left ? (std::numeric_limits<size_t>::max)() : left + right; };
    const auto stringCapacityBytes = [](size_t capacity) noexcept
    { return capacity >= (std::numeric_limits<size_t>::max)() / sizeof(wchar_t) ? (std::numeric_limits<size_t>::max)() : (capacity + 1u) * sizeof(wchar_t); };

    size_t bytes = addSaturated(sizeof(PendingUpdate), stringCapacityBytes(update.name.capacity()));
    if (update.pathIdentity.has_value())
        bytes = addSaturated(bytes, stringCapacityBytes(update.pathIdentity.value().acceptedSeparators.capacity()));
    const size_t itemCapacityBytes = update.topFiles.capacity() > (std::numeric_limits<size_t>::max)() / sizeof(FileSummaryItem)
                                         ? (std::numeric_limits<size_t>::max)()
                                         : update.topFiles.capacity() * sizeof(FileSummaryItem);
    bytes                          = addSaturated(bytes, itemCapacityBytes);
    for (const FileSummaryItem& file : update.topFiles)
    {
        bytes = addSaturated(bytes, stringCapacityBytes(file.name.capacity()));
    }
    return bytes;
}

uint64_t ViewerSpace::SubtractAtomicFloor(std::atomic_uint64_t& value, uint64_t delta) noexcept
{
    uint64_t current = value.load(std::memory_order_relaxed);
    for (;;)
    {
        const uint64_t next = current > delta ? current - delta : 0u;
        if (value.compare_exchange_weak(current, next, std::memory_order_relaxed, std::memory_order_relaxed))
        {
            return next;
        }
    }
}

const ViewerSpace::Node* ViewerSpace::TryGetRealNode(uint32_t nodeId) const noexcept
{
    if (nodeId == 0)
    {
        return nullptr;
    }

    const size_t index = static_cast<size_t>(nodeId);
    if (index >= _model->nodes.size())
    {
        return nullptr;
    }

    const Node& node = _model->nodes[index];
    if (node.id != nodeId)
    {
        return nullptr;
    }

    return &node;
}

ViewerSpace::Node* ViewerSpace::TryGetRealNode(uint32_t nodeId) noexcept
{
    if (nodeId == 0)
    {
        return nullptr;
    }

    const size_t index = static_cast<size_t>(nodeId);
    if (index >= _model->nodes.size())
    {
        return nullptr;
    }

    Node& node = _model->nodes[index];
    if (node.id != nodeId)
    {
        return nullptr;
    }

    return &node;
}

bool ViewerSpace::IsFileRecordId(uint32_t itemId) noexcept
{
    return itemId >= kFileRecordIdBase && itemId < kSyntheticNodeIdBase;
}

bool ViewerSpace::IsSyntheticNodeId(uint32_t itemId) noexcept
{
    return itemId >= kSyntheticNodeIdBase;
}

uint32_t ViewerSpace::FileRecordIdFromIndex(size_t index) noexcept
{
    if (index > static_cast<size_t>(kTreemapItemIdMask))
    {
        return 0u;
    }

    return kFileRecordIdBase + static_cast<uint32_t>(index);
}

uint32_t ViewerSpace::SyntheticOtherBucketIdForParent(uint32_t parentId) noexcept
{
    return kSyntheticNodeIdBase | (parentId & kTreemapItemIdMask);
}

const ViewerSpace::FileRecord* ViewerSpace::TryGetFileRecord(uint32_t itemId) const noexcept
{
    if (! IsFileRecordId(itemId))
    {
        return nullptr;
    }

    const size_t index = static_cast<size_t>(itemId - kFileRecordIdBase);
    if (index >= _model->fileRecords.size())
    {
        return nullptr;
    }

    const FileRecord& file = _model->fileRecords[index];
    if (file.id != itemId)
    {
        return nullptr;
    }

    return &file;
}

std::optional<ViewerSpace::ResolvedItem> ViewerSpace::ResolveTreemapItem(uint32_t itemId) const noexcept
{
    if (itemId == 0)
    {
        return std::nullopt;
    }

    if (const FileRecord* file = TryGetFileRecord(itemId))
    {
        return ResolvedItem{
            .id                 = file->id,
            .parentId           = file->parentId,
            .isDirectory        = false,
            .isSynthetic        = false,
            .isFileRecord       = true,
            .contentNotMeasured = false,
            .scanState          = ScanState::Done,
            .name               = file->name,
            .totalBytes         = file->bytes,
            .lastWriteTime      = file->lastWriteTime,
            .aggregateFolders   = 0,
            .aggregateFiles     = 1,
        };
    }

    if (const Node* node = TryGetRealNode(itemId))
    {
        return ResolvedItem{
            .id                 = node->id,
            .parentId           = node->parentId,
            .isDirectory        = node->isDirectory,
            .isSynthetic        = node->isSynthetic,
            .isFileRecord       = false,
            .contentNotMeasured = node->contentNotMeasured,
            .scanState          = node->scanState,
            .name               = node->name,
            .totalBytes         = node->totalBytes,
            .lastWriteTime      = node->lastWriteTime,
            .aggregateFolders   = node->aggregateFolders,
            .aggregateFiles     = node->aggregateFiles,
        };
    }

    const auto synthIt = _syntheticNodes.find(itemId);
    if (synthIt != _syntheticNodes.end())
    {
        const Node& node = synthIt->second;
        return ResolvedItem{
            .id                 = node.id,
            .parentId           = node.parentId,
            .isDirectory        = node.isDirectory,
            .isSynthetic        = node.isSynthetic,
            .isFileRecord       = false,
            .contentNotMeasured = node.contentNotMeasured,
            .scanState          = node.scanState,
            .name               = node.name,
            .totalBytes         = node.totalBytes,
            .aggregateFolders   = node.aggregateFolders,
            .aggregateFiles     = node.aggregateFiles,
        };
    }

    return std::nullopt;
}

std::span<const uint32_t> ViewerSpace::GetRealNodeChildren(const Node& node) const noexcept
{
    if (node.childrenCount == 0)
    {
        return {};
    }

    const size_t start = static_cast<size_t>(node.childrenStart);
    const size_t count = static_cast<size_t>(node.childrenCount);
    if (start >= _model->childrenArena.size())
    {
        return {};
    }

    size_t end = 0u;
    if (! ViewerSpaceScan::TryAddSize(start, count, end) || end > _model->childrenArena.size())
    {
        return {};
    }

    return std::span<const uint32_t>(_model->childrenArena.data() + start, count);
}

std::optional<uint32_t> ViewerSpace::AllocateChildArenaBlock(uint32_t capacity) noexcept
{
    if (capacity == 0u || capacity > _scanResourcePolicy.maxChildArenaSlots)
    {
        return std::nullopt;
    }

    size_t bestIndex      = _model->childrenArenaFreeBlocks.size();
    uint32_t bestCapacity = (std::numeric_limits<uint32_t>::max)();
    for (size_t index = 0u; index < _model->childrenArenaFreeBlocks.size(); ++index)
    {
        const ChildArenaFreeBlock& block = _model->childrenArenaFreeBlocks[index];
        if (block.capacity >= capacity && block.capacity < bestCapacity)
        {
            bestIndex    = index;
            bestCapacity = block.capacity;
        }
    }
    if (bestIndex < _model->childrenArenaFreeBlocks.size())
    {
        ChildArenaFreeBlock& block = _model->childrenArenaFreeBlocks[bestIndex];
        const uint32_t start       = block.start;
        if (block.capacity == capacity)
        {
            _model->childrenArenaFreeBlocks.erase(_model->childrenArenaFreeBlocks.begin() + static_cast<ptrdiff_t>(bestIndex));
        }
        else
        {
            block.start += capacity;
            block.capacity -= capacity;
        }
        return start;
    }

    const size_t previousSize = _model->childrenArena.size();
    if (previousSize > _scanResourcePolicy.maxChildArenaSlots || capacity > static_cast<uint32_t>(_scanResourcePolicy.maxChildArenaSlots - previousSize) ||
        previousSize > static_cast<size_t>((std::numeric_limits<uint32_t>::max)()))
    {
        return std::nullopt;
    }

    _model->childrenArena.resize(previousSize + capacity);
    return static_cast<uint32_t>(previousSize);
}

void ViewerSpace::ReleaseChildArenaBlock(uint32_t start, uint32_t capacity) noexcept
{
    if (capacity == 0u || start >= _model->childrenArena.size() || capacity > _model->childrenArena.size() - start)
    {
        return;
    }

    _model->childrenArenaFreeBlocks.push_back(ChildArenaFreeBlock{start, capacity});
    std::sort(_model->childrenArenaFreeBlocks.begin(),
              _model->childrenArenaFreeBlocks.end(),
              [](const ChildArenaFreeBlock& left, const ChildArenaFreeBlock& right) noexcept { return left.start < right.start; });

    size_t writeIndex = 0u;
    for (const ChildArenaFreeBlock& block : _model->childrenArenaFreeBlocks)
    {
        if (writeIndex == 0u)
        {
            _model->childrenArenaFreeBlocks[writeIndex++] = block;
            continue;
        }

        ChildArenaFreeBlock& previous = _model->childrenArenaFreeBlocks[writeIndex - 1u];
        const uint64_t previousEnd    = static_cast<uint64_t>(previous.start) + previous.capacity;
        if (previousEnd == block.start)
        {
            previous.capacity = ViewerSpaceScan::SaturatingAdd(previous.capacity, block.capacity);
        }
        else
        {
            _model->childrenArenaFreeBlocks[writeIndex++] = block;
        }
    }
    _model->childrenArenaFreeBlocks.resize(writeIndex);

    while (! _model->childrenArenaFreeBlocks.empty())
    {
        const ChildArenaFreeBlock& tail = _model->childrenArenaFreeBlocks.back();
        if (static_cast<uint64_t>(tail.start) + tail.capacity != _model->childrenArena.size())
        {
            break;
        }
        _model->childrenArena.resize(tail.start);
        _model->childrenArenaFreeBlocks.pop_back();
    }
}

uint64_t ViewerSpace::ChildArenaFreeSlots() const noexcept
{
    uint64_t freeSlots = 0u;
    for (const ChildArenaFreeBlock& block : _model->childrenArenaFreeBlocks)
    {
        freeSlots = ViewerSpaceScan::SaturatingAdd(freeSlots, static_cast<uint64_t>(block.capacity));
    }
    return freeSlots;
}

std::optional<size_t> ViewerSpace::AllocateFileRecordRange(size_t count) noexcept
{
    if (count == 0u)
    {
        return _model->fileRecords.size();
    }

    size_t bestIndex = _model->fileRecordFreeBlocks.size();
    size_t bestCount = (std::numeric_limits<size_t>::max)();
    for (size_t index = 0u; index < _model->fileRecordFreeBlocks.size(); ++index)
    {
        const FileRecordFreeBlock& block = _model->fileRecordFreeBlocks[index];
        if (block.count >= count && block.count < bestCount)
        {
            bestIndex = index;
            bestCount = block.count;
        }
    }
    if (bestIndex < _model->fileRecordFreeBlocks.size())
    {
        FileRecordFreeBlock& block = _model->fileRecordFreeBlocks[bestIndex];
        const size_t start         = block.start;
        if (block.count == count)
        {
            _model->fileRecordFreeBlocks.erase(_model->fileRecordFreeBlocks.begin() + static_cast<ptrdiff_t>(bestIndex));
        }
        else
        {
            block.start += count;
            block.count -= count;
        }
        return start;
    }

    const size_t maximumRecords = static_cast<size_t>(_scanResourcePolicy.maxRetainedFileRecords);
    if (_model->fileRecords.size() > maximumRecords || count > maximumRecords - _model->fileRecords.size())
    {
        return std::nullopt;
    }
    const size_t start = _model->fileRecords.size();
    _model->fileRecords.resize(start + count);
    return start;
}

void ViewerSpace::ReleaseFileRecordRange(size_t start, size_t count) noexcept
{
    if (count == 0u || start >= _model->fileRecords.size() || count > _model->fileRecords.size() - start)
    {
        return;
    }

    std::fill_n(_model->fileRecords.begin() + static_cast<ptrdiff_t>(start), count, FileRecord{});
    _model->fileRecordFreeBlocks.push_back(FileRecordFreeBlock{start, count});
    std::sort(_model->fileRecordFreeBlocks.begin(),
              _model->fileRecordFreeBlocks.end(),
              [](const FileRecordFreeBlock& left, const FileRecordFreeBlock& right) noexcept { return left.start < right.start; });

    size_t writeIndex = 0u;
    for (const FileRecordFreeBlock& block : _model->fileRecordFreeBlocks)
    {
        if (writeIndex == 0u)
        {
            _model->fileRecordFreeBlocks[writeIndex++] = block;
            continue;
        }
        FileRecordFreeBlock& previous = _model->fileRecordFreeBlocks[writeIndex - 1u];
        if (previous.start + previous.count >= block.start)
        {
            const size_t blockEnd = block.start + block.count;
            previous.count        = std::max(previous.count, blockEnd - previous.start);
        }
        else
        {
            _model->fileRecordFreeBlocks[writeIndex++] = block;
        }
    }
    _model->fileRecordFreeBlocks.resize(writeIndex);

    while (! _model->fileRecordFreeBlocks.empty())
    {
        const FileRecordFreeBlock& tail = _model->fileRecordFreeBlocks.back();
        if (tail.start + tail.count != _model->fileRecords.size())
        {
            break;
        }
        _model->fileRecords.resize(tail.start);
        _model->fileRecordFreeBlocks.pop_back();
    }
}

bool ViewerSpace::BeginDirectorySummary(const PendingUpdate& update) noexcept
{
    const DirectorySummaryStageKey key{update.generation, update.nodeId, update.summaryRevision};
    if (update.generation != _scanGeneration.load(std::memory_order_acquire))
    {
        return false;
    }
    if (key.nodeId == 0u || key.revision == 0u || _model->directorySummaryStages.contains(key) || _model->rejectedDirectorySummaryNodes.contains(key.nodeId) ||
        update.summaryExpectedFiles > _scanResourcePolicy.maxRetainedFileRecords || update.summaryNameBytes > _scanResourcePolicy.maxRetainedNameBytes ||
        ((update.summaryExpectedFiles == 0u) != (update.summaryExpectedChunks == 0u)) || update.summaryExpectedChunks > update.summaryExpectedFiles)
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }

    Node* const parent = TryGetRealNode(key.nodeId);
    if (parent == nullptr || ! parent->isDirectory)
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }
    const std::span<const uint32_t> existingChildren = GetRealNodeChildren(*parent);
    if (existingChildren.size() != parent->childrenCount)
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }

    const uint64_t childCapacity64 = static_cast<uint64_t>(parent->childrenCount) + update.summaryExpectedFiles;
    if (update.summaryExpectedFiles != 0u &&
        (childCapacity64 > _scanResourcePolicy.maxChildArenaSlots || childCapacity64 > (std::numeric_limits<uint32_t>::max)()))
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }

    uint64_t recordBytes = static_cast<uint64_t>(update.summaryExpectedFiles) * static_cast<uint64_t>(sizeof(FileRecord));
    uint64_t childBytes  = update.summaryExpectedFiles == 0u ? 0u : childCapacity64 * sizeof(uint32_t);
    uint64_t stageBytes =
        ViewerSpaceScan::SaturatingAdd(static_cast<uint64_t>(sizeof(DirectorySummaryStageKey) + sizeof(DirectorySummaryStage)),
                                       ViewerSpaceScan::SaturatingAdd(recordBytes, ViewerSpaceScan::SaturatingAdd(childBytes, update.summaryNameBytes)));
    uint64_t stageBudget = _scanResourcePolicy.maxRetainedNameBytes;
    stageBudget          = ViewerSpaceScan::SaturatingAdd(stageBudget, static_cast<uint64_t>(_scanResourcePolicy.maxRetainedFileRecords) * sizeof(FileRecord));
    stageBudget          = ViewerSpaceScan::SaturatingAdd(stageBudget, static_cast<uint64_t>(_scanResourcePolicy.maxChildArenaSlots) * sizeof(uint32_t));
    stageBudget          = ViewerSpaceScan::SaturatingAdd(stageBudget,
                                                          static_cast<uint64_t>(_scanResourcePolicy.maxOutstandingDirectories) *
                                                              static_cast<uint64_t>(sizeof(DirectorySummaryStageKey) + sizeof(DirectorySummaryStage)));
    if (stageBytes > stageBudget || _model->directorySummaryStageBytes > stageBudget - stageBytes)
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }

    const std::optional<size_t> recordStart = AllocateFileRecordRange(update.summaryExpectedFiles);
    if (! recordStart.has_value())
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }

    DirectorySummaryStage stage;
    stage.otherBytes            = update.otherBytes;
    stage.otherFolders          = update.otherFolders;
    stage.otherFiles            = update.otherFiles;
    stage.expectedNameBytes     = update.summaryNameBytes;
    stage.fileRecordStart       = recordStart.value();
    stage.expectedChunks        = update.summaryExpectedChunks;
    stage.expectedFiles         = update.summaryExpectedFiles;
    stage.previousChildStart    = parent->childrenStart;
    stage.previousChildCount    = parent->childrenCount;
    stage.previousChildCapacity = parent->childrenCapacity;
    if (update.summaryExpectedFiles == 0u)
    {
        stage.copiedChildren = stage.previousChildCount;
    }
    stage.accountedBytes = stageBytes;

    if (update.summaryExpectedFiles != 0u)
    {
        const std::optional<uint32_t> childStart = AllocateChildArenaBlock(static_cast<uint32_t>(childCapacity64));
        if (! childStart.has_value())
        {
            ReleaseFileRecordRange(recordStart.value(), update.summaryExpectedFiles);
            static_cast<void>(RejectDirectorySummary(key, true));
            return false;
        }
        stage.hasChildArena      = true;
        stage.childArenaStart    = childStart.value();
        stage.childArenaCapacity = static_cast<uint32_t>(childCapacity64);
    }

    const auto [insertedAt, inserted] = _model->directorySummaryStages.emplace(key, stage);
    if (! inserted)
    {
        if (stage.hasChildArena)
        {
            ReleaseChildArenaBlock(stage.childArenaStart, stage.childArenaCapacity);
        }
        ReleaseFileRecordRange(stage.fileRecordStart, stage.expectedFiles);
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }
    _model->directorySummaryStageBytes = ViewerSpaceScan::SaturatingAdd(_model->directorySummaryStageBytes, insertedAt->second.accountedBytes);
    if (insertedAt->second.hasChildArena && insertedAt->second.previousChildCount != 0u)
    {
        _model->directorySummaryCopyQueue.push_back(key);
    }
    return true;
}

bool ViewerSpace::ApplyDirectorySummaryChunk(const PendingUpdate& update) noexcept
{
    const DirectorySummaryStageKey key{update.generation, update.nodeId, update.summaryRevision};
    if (update.generation != _scanGeneration.load(std::memory_order_acquire))
    {
        return false;
    }
    auto stageIt = _model->directorySummaryStages.find(key);
    if (stageIt == _model->directorySummaryStages.end())
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }
    DirectorySummaryStage& stage                          = stageIt->second;
    const size_t chunkBytes                               = EstimatePendingUpdateBytes(update);
    static constexpr size_t kMaxOrdinarySummaryChunkBytes = 64u * 1024u;
    const size_t maxNameBytes                             = static_cast<size_t>(_scanResourcePolicy.maxProviderNameBytes);
    const size_t boundedSingleNameException               = maxNameBytes > (std::numeric_limits<size_t>::max)() - kMaxOrdinarySummaryChunkBytes
                                                                ? (std::numeric_limits<size_t>::max)()
                                                                : kMaxOrdinarySummaryChunkBytes + maxNameBytes;
    const bool validSingleNameException = update.topFiles.size() == 1u &&
                                          ViewerSpaceScan::ValidateProviderChildName(update.topFiles.front().name) == ViewerSpaceScan::ValidationError::None &&
                                          update.topFiles.front().name.size() <= maxNameBytes / sizeof(wchar_t) && chunkBytes <= boundedSingleNameException;
    if (stage.endReceived || update.summarySequence != stage.receivedChunks || update.summaryExpectedChunks != stage.expectedChunks ||
        update.summaryExpectedFiles != stage.expectedFiles || update.summaryNameBytes != stage.expectedNameBytes || update.topFiles.empty() ||
        (chunkBytes > kMaxOrdinarySummaryChunkBytes && ! validSingleNameException) ||
        update.topFiles.size() > stage.expectedFiles - std::min(stage.receivedFiles, stage.expectedFiles))
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }

    uint64_t chunkNameBytes = 0u;
    for (const FileSummaryItem& file : update.topFiles)
    {
        const size_t maxNameChars = static_cast<size_t>(_scanResourcePolicy.maxProviderNameBytes) / sizeof(wchar_t);
        if (file.name.empty() || file.name.size() > maxNameChars ||
            ViewerSpaceScan::ValidateProviderChildName(file.name) != ViewerSpaceScan::ValidationError::None)
        {
            static_cast<void>(RejectDirectorySummary(key, true));
            return false;
        }
        const uint64_t nameBytes = static_cast<uint64_t>(file.name.size()) * sizeof(wchar_t);
        chunkNameBytes           = ViewerSpaceScan::SaturatingAdd(chunkNameBytes, nameBytes);
    }
    if (chunkNameBytes > stage.expectedNameBytes - std::min(stage.stagedNameBytes, stage.expectedNameBytes))
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }

    for (const FileSummaryItem& file : update.topFiles)
    {
        const size_t fileIndex = stage.fileRecordStart + stage.receivedFiles;
        if (fileIndex >= _model->fileRecords.size())
        {
            static_cast<void>(RejectDirectorySummary(key, true));
            return false;
        }
        const uint32_t fileId = FileRecordIdFromIndex(fileIndex);
        if (fileId == 0u)
        {
            static_cast<void>(RejectDirectorySummary(key, true));
            return false;
        }
        FileRecord record;
        record.id                      = fileId;
        record.parentId                = key.nodeId;
        record.name                    = CopyToArena(_model->nameArena, file.name);
        record.bytes                   = file.bytes;
        record.lastWriteTime           = file.lastWriteTime;
        _model->fileRecords[fileIndex] = record;
        if (stage.hasChildArena)
        {
            const uint64_t childOffset = static_cast<uint64_t>(stage.previousChildCount) + stage.receivedFiles;
            if (childOffset >= stage.childArenaCapacity)
            {
                static_cast<void>(RejectDirectorySummary(key, true));
                return false;
            }
            _model->childrenArena[static_cast<size_t>(stage.childArenaStart) + static_cast<size_t>(childOffset)] = fileId;
        }
        stage.receivedFiles += 1u;
        stage.stagedFileBytes = ViewerSpaceScan::SaturatingAdd(stage.stagedFileBytes, file.bytes);
        stage.stagedNameBytes = ViewerSpaceScan::SaturatingAdd(stage.stagedNameBytes, static_cast<uint64_t>(file.name.size()) * sizeof(wchar_t));
    }
    stage.receivedChunks += 1u;

    uint64_t observed = _directorySummaryChunkMaxBytes.load(std::memory_order_relaxed);
    while (chunkBytes > observed &&
           ! _directorySummaryChunkMaxBytes.compare_exchange_weak(observed, static_cast<uint64_t>(chunkBytes), std::memory_order_relaxed))
    {
    }
    return true;
}

bool ViewerSpace::EndDirectorySummary(const PendingUpdate& update) noexcept
{
    const DirectorySummaryStageKey key{update.generation, update.nodeId, update.summaryRevision};
    if (update.generation != _scanGeneration.load(std::memory_order_acquire))
    {
        return false;
    }
    const auto stageIt = _model->directorySummaryStages.find(key);
    if (stageIt == _model->directorySummaryStages.end())
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }
    DirectorySummaryStage& stage = stageIt->second;
    if (stage.endReceived || update.summarySequence != stage.expectedChunks || update.summaryExpectedChunks != stage.expectedChunks ||
        update.summaryExpectedFiles != stage.expectedFiles || update.summaryNameBytes != stage.expectedNameBytes || ! update.topFiles.empty() ||
        stage.receivedChunks != stage.expectedChunks || stage.receivedFiles != stage.expectedFiles || stage.stagedNameBytes != stage.expectedNameBytes)
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }
    stage.endReceived = true;
    if (stage.copiedChildren == stage.previousChildCount)
    {
        return CommitDirectorySummary(key);
    }
    return true;
}

bool ViewerSpace::CommitDirectorySummary(const DirectorySummaryStageKey& key) noexcept
{
    const auto stageIt = _model->directorySummaryStages.find(key);
    if (stageIt == _model->directorySummaryStages.end())
    {
        return false;
    }
    const DirectorySummaryStage stage = stageIt->second;
    if (! stage.endReceived || stage.copiedChildren != stage.previousChildCount)
    {
        return false;
    }

    Node* const parent = TryGetRealNode(key.nodeId);
    if (parent == nullptr || parent->childrenStart != stage.previousChildStart || parent->childrenCount != stage.previousChildCount ||
        parent->childrenCapacity != stage.previousChildCapacity)
    {
        static_cast<void>(RejectDirectorySummary(key, true));
        return false;
    }
    if (stage.hasChildArena)
    {
        parent->childrenStart    = stage.childArenaStart;
        parent->childrenCount    = stage.childArenaCapacity;
        parent->childrenCapacity = stage.childArenaCapacity;
        ReleaseChildArenaBlock(stage.previousChildStart, stage.previousChildCapacity);
    }
    parent->observedOwnFileBytes = ViewerSpaceScan::SaturatingAdd(parent->observedOwnFileBytes, stage.stagedFileBytes);
    parent->aggregateBytes       = ViewerSpaceScan::SaturatingAdd(parent->aggregateBytes, stage.otherBytes);
    parent->aggregateFolders     = ViewerSpaceScan::SaturatingAdd(parent->aggregateFolders, stage.otherFolders);
    parent->aggregateFiles       = ViewerSpaceScan::SaturatingAdd(parent->aggregateFiles, stage.otherFiles);
    _model->directorySummaryStageBytes =
        _model->directorySummaryStageBytes > stage.accountedBytes ? _model->directorySummaryStageBytes - stage.accountedBytes : 0u;
    _model->directorySummaryStages.erase(stageIt);
    static_cast<void>(RaiseNodeObservedBytes(parent->id, 0u));
    _directorySummaryCommittedCount = ViewerSpaceScan::SaturatingAdd(_directorySummaryCommittedCount, uint64_t{1u});
    ApplyDeferredDirectoryDoneStates();
    return true;
}

bool ViewerSpace::RejectDirectorySummary(const DirectorySummaryStageKey& key, bool markRejected) noexcept
{
    const auto stageIt  = _model->directorySummaryStages.find(key);
    const bool hadStage = stageIt != _model->directorySummaryStages.end();
    bool rejected       = false;
    if (hadStage)
    {
        const DirectorySummaryStage stage = stageIt->second;
        if (stage.hasChildArena)
        {
            ReleaseChildArenaBlock(stage.childArenaStart, stage.childArenaCapacity);
        }
        ReleaseFileRecordRange(stage.fileRecordStart, stage.expectedFiles);
        _model->directorySummaryStageBytes =
            _model->directorySummaryStageBytes > stage.accountedBytes ? _model->directorySummaryStageBytes - stage.accountedBytes : 0u;
        _model->directorySummaryStages.erase(stageIt);
        rejected = true;
    }
    if (markRejected && key.generation == _scanGeneration.load(std::memory_order_acquire) && key.nodeId != 0u)
    {
        rejected = _model->rejectedDirectorySummaryNodes.insert(key.nodeId).second || rejected;
        if (rejected)
        {
            _directorySummaryRejectedCount = ViewerSpaceScan::SaturatingAdd(_directorySummaryRejectedCount, uint64_t{1u});
            if (Node* node = TryGetRealNode(key.nodeId))
            {
                node->scanState = ScanState::Error;
            }
        }
    }
    return rejected;
}

size_t ViewerSpace::AdvanceDirectorySummaryStages(size_t copyBudget) noexcept
{
    size_t copied        = 0u;
    size_t visitedStages = 0u;
    while (! _model->directorySummaryCopyQueue.empty() && copied < copyBudget && visitedStages < 64u)
    {
        const DirectorySummaryStageKey key = _model->directorySummaryCopyQueue.front();
        _model->directorySummaryCopyQueue.pop_front();
        ++visitedStages;
        auto stageIt = _model->directorySummaryStages.find(key);
        if (stageIt == _model->directorySummaryStages.end())
        {
            continue;
        }
        DirectorySummaryStage& stage = stageIt->second;
        if (stage.copiedChildren >= stage.previousChildCount)
        {
            continue;
        }
        Node* const parent = TryGetRealNode(key.nodeId);
        if (parent == nullptr || parent->childrenStart != stage.previousChildStart || parent->childrenCount != stage.previousChildCount ||
            parent->childrenCapacity != stage.previousChildCapacity)
        {
            static_cast<void>(RejectDirectorySummary(key, true));
            continue;
        }
        const size_t remaining = static_cast<size_t>(stage.previousChildCount - stage.copiedChildren);
        const size_t amount    = std::min({remaining, copyBudget - copied, size_t{128u}});
        const size_t source    = static_cast<size_t>(stage.previousChildStart) + stage.copiedChildren;
        const size_t dest      = static_cast<size_t>(stage.childArenaStart) + stage.copiedChildren;
        std::copy_n(_model->childrenArena.begin() + static_cast<ptrdiff_t>(source), amount, _model->childrenArena.begin() + static_cast<ptrdiff_t>(dest));
        stage.copiedChildren += static_cast<uint32_t>(amount);
        copied += amount;
        if (stage.copiedChildren < stage.previousChildCount)
        {
            _model->directorySummaryCopyQueue.push_back(key);
            continue;
        }
        if (stage.endReceived)
        {
            static_cast<void>(CommitDirectorySummary(key));
        }
    }
    return copied;
}

bool ViewerSpace::HasDirectorySummaryStage(uint32_t generation, uint32_t nodeId) const noexcept
{
    return std::any_of(_model->directorySummaryStages.begin(), _model->directorySummaryStages.end(), [&](const auto& entry) {
        return entry.first.generation == generation && entry.first.nodeId == nodeId;
    });
}

bool ViewerSpace::HasPendingDirectorySummaryChildBlock(uint32_t generation, uint32_t nodeId) const noexcept
{
    return std::any_of(_model->directorySummaryStages.begin(), _model->directorySummaryStages.end(), [&](const auto& entry) {
        return entry.first.generation == generation && entry.first.nodeId == nodeId;
    });
}

void ViewerSpace::ApplyDeferredDirectoryDoneStates() noexcept
{
    bool rootChanged = false;
    for (auto it = _model->deferredDirectoryDoneStates.begin(); it != _model->deferredDirectoryDoneStates.end();)
    {
        const uint32_t nodeId = it->first;
        if ((nodeId == _model->rootNodeId && ! _model->directorySummaryStages.empty()) ||
            HasDirectorySummaryStage(_scanGeneration.load(std::memory_order_acquire), nodeId))
        {
            ++it;
            continue;
        }
        Node* node = TryGetRealNode(nodeId);
        if (node != nullptr)
        {
            node->scanState =
                _model->rejectedDirectorySummaryNodes.contains(nodeId) || (nodeId == _model->rootNodeId && ! _model->rejectedDirectorySummaryNodes.empty())
                    ? ScanState::Error
                    : it->second;
            rootChanged = rootChanged || nodeId == _model->rootNodeId;
        }
        it = _model->deferredDirectoryDoneStates.erase(it);
    }
    if (rootChanged)
    {
        UpdateHeaderTextCache();
    }
}

void ViewerSpace::DiscardDirectorySummaryStages() noexcept
{
    for (const auto& [key, stage] : _model->directorySummaryStages)
    {
        static_cast<void>(key);
        if (stage.hasChildArena)
        {
            ReleaseChildArenaBlock(stage.childArenaStart, stage.childArenaCapacity);
        }
        ReleaseFileRecordRange(stage.fileRecordStart, stage.expectedFiles);
    }
    _model->directorySummaryStages.clear();
    _model->directorySummaryStages.rehash(0u);
    std::pmr::deque<DirectorySummaryStageKey> emptyCopyQueue{&_model->memory};
    _model->directorySummaryCopyQueue.swap(emptyCopyQueue);
    _model->deferredDirectoryDoneStates.clear();
    _model->deferredDirectoryDoneStates.rehash(0u);
    _model->rejectedDirectorySummaryNodes.clear();
    _model->rejectedDirectorySummaryNodes.rehash(0u);
    std::pmr::vector<FileRecordFreeBlock> emptyFileRecordFreeBlocks{&_model->memory};
    _model->fileRecordFreeBlocks.swap(emptyFileRecordFreeBlocks);
    _model->directorySummaryStageBytes = 0u;
    _directorySummaryCommittedCount    = 0u;
    _directorySummaryRejectedCount     = 0u;
    _directorySummaryChunkMaxBytes.store(0u, std::memory_order_relaxed);
}

bool ViewerSpace::AddRealNodeChild(Node& parent, uint32_t childItemId) noexcept
{
    static constexpr uint32_t kInitialCapacity = 2;

    if (parent.childrenCapacity == 0)
    {
        const std::optional<uint32_t> start = AllocateChildArenaBlock(kInitialCapacity);
        if (! start.has_value())
        {
            return false;
        }

        parent.childrenStart    = start.value();
        parent.childrenCount    = 0;
        parent.childrenCapacity = kInitialCapacity;
    }

    if (parent.childrenCount >= parent.childrenCapacity)
    {
        const uint32_t oldStart    = parent.childrenStart;
        const uint32_t oldCount    = parent.childrenCount;
        const uint32_t oldCapacity = parent.childrenCapacity;

        uint32_t nextCapacity = oldCapacity;
        if (nextCapacity < kInitialCapacity)
        {
            nextCapacity = kInitialCapacity;
        }
        else if (nextCapacity > (std::numeric_limits<uint32_t>::max() / 2u))
        {
            nextCapacity = std::numeric_limits<uint32_t>::max();
        }
        else
        {
            nextCapacity = nextCapacity * 2u;
        }

        nextCapacity = std::min(nextCapacity, _scanResourcePolicy.maxChildArenaSlots);
        if (nextCapacity <= oldCapacity)
        {
            return false;
        }

        const uint64_t oldEnd = static_cast<uint64_t>(oldStart) + oldCapacity;
        const uint32_t growth = nextCapacity - oldCapacity;
        if (oldEnd == _model->childrenArena.size() && _model->childrenArena.size() <= _scanResourcePolicy.maxChildArenaSlots &&
            growth <= static_cast<size_t>(_scanResourcePolicy.maxChildArenaSlots) - _model->childrenArena.size())
        {
            _model->childrenArena.resize(_model->childrenArena.size() + growth, 0u);
            parent.childrenCapacity = nextCapacity;
        }
        else
        {
            const std::optional<uint32_t> newStart = AllocateChildArenaBlock(nextCapacity);
            if (! newStart.has_value())
            {
                return false;
            }

            if (oldCount > 0)
            {
                const size_t oldStartIndex = static_cast<size_t>(oldStart);
                const size_t oldCountIndex = static_cast<size_t>(oldCount);
                size_t oldEndIndex         = 0u;
                if (oldStartIndex < _model->childrenArena.size() && ViewerSpaceScan::TryAddSize(oldStartIndex, oldCountIndex, oldEndIndex) &&
                    oldEndIndex <= _model->childrenArena.size())
                {
                    std::copy_n(_model->childrenArena.data() + oldStartIndex, oldCountIndex, _model->childrenArena.data() + newStart.value());
                }
                else
                {
                    ReleaseChildArenaBlock(newStart.value(), nextCapacity);
                    return false;
                }
            }

            parent.childrenStart    = newStart.value();
            parent.childrenCapacity = nextCapacity;
            parent.childrenCount    = oldCount;
            ReleaseChildArenaBlock(oldStart, oldCapacity);
        }
    }

    const size_t slot = static_cast<size_t>(parent.childrenStart) + static_cast<size_t>(parent.childrenCount);
    if (slot >= _model->childrenArena.size())
    {
        return false;
    }

    _model->childrenArena[slot] = childItemId;
    parent.childrenCount += 1;
    return true;
}

std::wstring ViewerSpace::BuildNodePathText(uint32_t nodeId) const
{
    if (_model == nullptr || _model->scanRootPath.empty() || _model->rootNodeId == 0)
    {
        return {};
    }

    if (nodeId == _model->rootNodeId)
    {
        // The admitted root path precedes node publication during cache validation.
        // Refresh/Rebuild must still be able to replace that uncommitted generation.
        return _model->scanRootPath;
    }

    const Node* node = TryGetRealNode(nodeId);
    if (node == nullptr || node->isSynthetic)
    {
        return {};
    }

    std::vector<const Node*> segments;
    segments.reserve(24);

    uint32_t currentId = nodeId;
    while (currentId != 0 && currentId != _model->rootNodeId)
    {
        const Node* current = TryGetRealNode(currentId);
        if (current == nullptr)
        {
            return {};
        }

        if (current->name.empty())
        {
            return {};
        }

        segments.push_back(current);
        currentId = current->parentId;
    }

    if (currentId != _model->rootNodeId)
    {
        return {};
    }

    const wchar_t separator = DeterminePreferredPathSeparator(_model->scanRootPath, _fileSystemIsWin32);

    std::wstring pathText = _model->scanRootPath;
    if (! pathText.empty() && pathText.back() != L'/' && pathText.back() != L'\\')
    {
        pathText.push_back(separator);
    }

    for (auto it = segments.rbegin(); it != segments.rend(); ++it)
    {
        const Node* segment = *it;
        if (segment == nullptr)
        {
            return {};
        }

        pathText.append(segment->name.data(), segment->name.size());
        if (it + 1 != segments.rend())
        {
            pathText.push_back(separator);
        }
    }

    return pathText;
}

std::wstring ViewerSpace::BuildItemPathText(uint32_t itemId) const
{
    if (const FileRecord* file = TryGetFileRecord(itemId))
    {
        std::wstring parentPath = BuildNodePathText(file->parentId);
        if (parentPath.empty() || file->name.empty())
        {
            return {};
        }

        const wchar_t separator = DeterminePreferredPathSeparator(parentPath, _fileSystemIsWin32);
        if (! parentPath.empty() && parentPath.back() != L'/' && parentPath.back() != L'\\')
        {
            parentPath.push_back(separator);
        }
        parentPath.append(file->name.data(), file->name.size());
        return parentPath;
    }

    return BuildNodePathText(itemId);
}

void ViewerSpace::UpdateViewPathText() noexcept
{
    _viewPathText = BuildNodePathText(_model->viewNodeId);
    _headerPathSourceText.clear();
    _headerPathDisplayText.clear();
    _headerPathDisplayMaxWidthDip = 0.0f;
    UpdateNavigationHostState();
}

uint32_t ViewerSpace::ComputeAdaptiveFileCandidateBudget() const noexcept
{
    const float widthDip        = DipFromPx(_clientSize.cx);
    const float heightDip       = DipFromPx(_clientSize.cy);
    const float headerBottomDip = GetHeaderBottomDip();
    const D2D1_RECT_F viewRc    = D2D1::RectF(kPaddingDip,
                                              headerBottomDip + kPaddingDip,
                                              std::max(kPaddingDip, widthDip - kPaddingDip),
                                              std::max(headerBottomDip + kPaddingDip, heightDip - kPaddingDip));

    const uint32_t pixelDrivenBudget = std::min<uint32_t>(PixelDrivenLayoutItemLimit(RectArea(viewRc)), kAdaptiveFileCandidateCeiling);
    return std::max<uint32_t>(_config.topFilesPerDirectory, pixelDrivenBudget);
}

void ViewerSpace::EnsureLayoutForView() noexcept
{
    const double now = NowSeconds();
    for (auto& item : _drawItems)
    {
        const double dt         = now - item.animationStartSeconds;
        const double t          = std::clamp(dt / kAnimationDurationSeconds, 0.0, 1.0);
        const double eased      = EaseOutCubic(t);
        item.currentRect.left   = static_cast<float>(item.startRect.left + (item.targetRect.left - item.startRect.left) * eased);
        item.currentRect.top    = static_cast<float>(item.startRect.top + (item.targetRect.top - item.startRect.top) * eased);
        item.currentRect.right  = static_cast<float>(item.startRect.right + (item.targetRect.right - item.startRect.right) * eased);
        item.currentRect.bottom = static_cast<float>(item.startRect.bottom + (item.targetRect.bottom - item.startRect.bottom) * eased);
    }
}

ViewerSpace::DrawItem::Lod ViewerSpace::ClassifyDrawItemLod(const D2D1_RECT_F& itemRc, uint8_t depth) noexcept
{
    float gap = kItemGapDip - static_cast<float>(depth) * 0.15f;
    gap       = std::clamp(gap, 0.5f, kItemGapDip);

    const float w    = std::max(0.0f, (itemRc.right - itemRc.left) - (gap * 2.0f));
    const float h    = std::max(0.0f, (itemRc.bottom - itemRc.top) - (gap * 2.0f));
    const float area = w * h;

    if (w < 1.0f || h < 1.0f || area < 1.0f)
    {
        return DrawItem::Lod::Culled;
    }

    if (w >= 110.0f && h >= 80.0f)
    {
        return DrawItem::Lod::Hero;
    }

    if (w >= 56.0f && h >= 36.0f && area >= 1800.0f)
    {
        return DrawItem::Lod::Large;
    }

    if (w >= 24.0f && h >= 12.0f && area >= 400.0f)
    {
        return DrawItem::Lod::Medium;
    }

    if (std::min(w, h) >= 3.0f && area >= 16.0f)
    {
        return DrawItem::Lod::Small;
    }

    return DrawItem::Lod::Tiny;
}

void ViewerSpace::BuildHitGrid(const D2D1_RECT_F& bounds) noexcept
{
    _hitGrid.Clear();
    const float boundsW = std::max(0.0f, bounds.right - bounds.left);
    const float boundsH = std::max(0.0f, bounds.bottom - bounds.top);
    if (boundsW <= 1.0f || boundsH <= 1.0f || _drawItems.size() < 256u)
    {
        return;
    }

    const uint32_t columns = static_cast<uint32_t>(std::clamp(std::ceil(static_cast<double>(boundsW) / 96.0), 8.0, 64.0));
    const uint32_t rows    = static_cast<uint32_t>(std::clamp(std::ceil(static_cast<double>(boundsH) / 96.0), 8.0, 64.0));
    _hitGrid.bounds        = bounds;
    _hitGrid.columns       = columns;
    _hitGrid.rows          = rows;
    _hitGrid.cells.resize(static_cast<size_t>(columns) * static_cast<size_t>(rows));

    auto cellIndex = [columns](uint32_t x, uint32_t y) noexcept -> size_t
    { return static_cast<size_t>(y) * static_cast<size_t>(columns) + static_cast<size_t>(x); };

    for (size_t itemIndex = 0; itemIndex < _drawItems.size(); ++itemIndex)
    {
        const DrawItem& item = _drawItems[itemIndex];
        if (item.lod == DrawItem::Lod::Culled)
        {
            continue;
        }

        float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
        gap       = std::clamp(gap, 0.5f, kItemGapDip);

        D2D1_RECT_F rc = item.targetRect;
        rc.left += gap;
        rc.top += gap;
        rc.right -= gap;
        rc.bottom -= gap;
        if (rc.right <= rc.left || rc.bottom <= rc.top || RectArea(rc) < kMinHitAreaDip2)
        {
            continue;
        }

        const auto clampCellX = [&](float x) noexcept -> uint32_t
        {
            const float t = (x - bounds.left) / std::max(1.0f, boundsW);
            return static_cast<uint32_t>(std::clamp(std::floor(static_cast<double>(t) * static_cast<double>(columns)), 0.0, static_cast<double>(columns - 1u)));
        };
        const auto clampCellY = [&](float y) noexcept -> uint32_t
        {
            const float t = (y - bounds.top) / std::max(1.0f, boundsH);
            return static_cast<uint32_t>(std::clamp(std::floor(static_cast<double>(t) * static_cast<double>(rows)), 0.0, static_cast<double>(rows - 1u)));
        };

        const uint32_t leftCell   = clampCellX(rc.left);
        const uint32_t rightCell  = clampCellX(rc.right);
        const uint32_t topCell    = clampCellY(rc.top);
        const uint32_t bottomCell = clampCellY(rc.bottom);

        for (uint32_t y = topCell; y <= bottomCell; ++y)
        {
            for (uint32_t x = leftCell; x <= rightCell; ++x)
            {
                std::vector<uint32_t>& cell = _hitGrid.cells[cellIndex(x, y)];
                cell.push_back(static_cast<uint32_t>(itemIndex));
                _hitGrid.maxCandidatesPerCell = std::max<uint32_t>(_hitGrid.maxCandidatesPerCell, static_cast<uint32_t>(cell.size()));
            }
        }
    }

    Debug::Perf::EmitValue(L"viewer.space.hit_grid.cells", static_cast<uint64_t>(_hitGrid.cells.size()));
    Debug::Perf::EmitValue(L"viewer.space.hit_grid.max_candidates", _hitGrid.maxCandidatesPerCell);
}

std::optional<uint32_t> ViewerSpace::HitTestTreemap(float xDip, float yDip) const noexcept
{
    const auto hitStartedAt     = std::chrono::steady_clock::now();
    uint32_t candidatesChecked  = 0u;
    std::wstring_view hitDetail = L"linear";
    auto emitHitMetric          = wil::scope_exit([&]() noexcept
    {
        const uint64_t hitUs = Debug::Perf::ElapsedUs(hitStartedAt);
#ifdef ENABLE_TESTS
        _lastHitTestUs                = hitUs;
        _lastHitTestCandidatesChecked = candidatesChecked;
        _lastHitTestPath              = hitDetail == L"grid" ? 1u : hitDetail == L"linear-animation" ? 2u : hitDetail == L"grid-fallback-linear" ? 4u : 3u;
#endif
        Debug::Perf::Emit(L"viewer.space.hit_test_us", hitDetail, hitUs, static_cast<uint64_t>(_drawItems.size()), candidatesChecked, S_OK);
    });

    if (yDip < GetHeaderBottomDip())
    {
        return std::nullopt;
    }

    if (HasActiveLayoutAnimation(NowSeconds()))
    {
        hitDetail = L"linear-animation";
        return HitTestTreemapLinear(xDip, yDip, candidatesChecked);
    }

    if (_hitGrid.columns > 0u && _hitGrid.rows > 0u && ! _hitGrid.cells.empty() && xDip >= _hitGrid.bounds.left && xDip <= _hitGrid.bounds.right &&
        yDip >= _hitGrid.bounds.top && yDip <= _hitGrid.bounds.bottom)
    {
        hitDetail            = L"grid";
        const float boundsW  = std::max(1.0f, _hitGrid.bounds.right - _hitGrid.bounds.left);
        const float boundsH  = std::max(1.0f, _hitGrid.bounds.bottom - _hitGrid.bounds.top);
        const uint32_t cellX = static_cast<uint32_t>(
            std::clamp(std::floor((static_cast<double>(xDip - _hitGrid.bounds.left) / static_cast<double>(boundsW)) * static_cast<double>(_hitGrid.columns)),
                       0.0,
                       static_cast<double>(_hitGrid.columns - 1u)));
        const uint32_t cellY = static_cast<uint32_t>(
            std::clamp(std::floor((static_cast<double>(yDip - _hitGrid.bounds.top) / static_cast<double>(boundsH)) * static_cast<double>(_hitGrid.rows)),
                       0.0,
                       static_cast<double>(_hitGrid.rows - 1u)));
        const size_t index = static_cast<size_t>(cellY) * static_cast<size_t>(_hitGrid.columns) + static_cast<size_t>(cellX);
        if (index < _hitGrid.cells.size())
        {
            const std::vector<uint32_t>& candidates = _hitGrid.cells[index];
            for (size_t i = candidates.size(); i-- > 0;)
            {
                const uint32_t itemIndex = candidates[i];
                if (itemIndex >= _drawItems.size())
                {
                    continue;
                }

                const DrawItem& item = _drawItems[itemIndex];
                candidatesChecked += 1u;

                float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
                gap       = std::clamp(gap, 0.5f, kItemGapDip);

                D2D1_RECT_F rc = item.currentRect;
                rc.left += gap;
                rc.top += gap;
                rc.right -= gap;
                rc.bottom -= gap;
                if (rc.right <= rc.left || rc.bottom <= rc.top)
                {
                    continue;
                }

                if (xDip >= rc.left && xDip <= rc.right && yDip >= rc.top && yDip <= rc.bottom)
                {
                    const float area = RectArea(rc);
                    if (area >= kMinHitAreaDip2)
                    {
                        return item.nodeId;
                    }
                }
            }

            hitDetail                         = L"grid-fallback-linear";
            uint32_t linearCandidates         = 0u;
            std::optional<uint32_t> linearHit = HitTestTreemapLinear(xDip, yDip, linearCandidates);
            candidatesChecked += linearCandidates;
            return linearHit;
        }
    }

    std::optional<uint32_t> linearHit = HitTestTreemapLinear(xDip, yDip, candidatesChecked);
    return linearHit;
}

std::optional<uint32_t> ViewerSpace::HitTestTreemapLinear(float xDip, float yDip, uint32_t& candidatesChecked) const noexcept
{
    candidatesChecked = 0u;
    if (yDip < GetHeaderBottomDip())
    {
        return std::nullopt;
    }

    for (size_t i = _drawItems.size(); i-- > 0;)
    {
        const DrawItem& item = _drawItems[i];
        candidatesChecked += 1u;

        float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
        gap       = std::clamp(gap, 0.5f, kItemGapDip);

        D2D1_RECT_F rc = item.currentRect;
        rc.left += gap;
        rc.top += gap;
        rc.right -= gap;
        rc.bottom -= gap;
        if (rc.right <= rc.left || rc.bottom <= rc.top)
        {
            continue;
        }

        if (xDip >= rc.left && xDip <= rc.right && yDip >= rc.top && yDip <= rc.bottom)
        {
            const float area = RectArea(rc);
            if (area >= kMinHitAreaDip2)
            {
                return item.nodeId;
            }
        }
    }

    return std::nullopt;
}

void ViewerSpace::NavigateTo(uint32_t nodeId) noexcept
{
    const auto handlerStartedAt = std::chrono::steady_clock::now();
    CancelDeferredStartScan();
    if (nodeId == 0u || _model == nullptr || nodeId == _model->viewNodeId)
    {
        return;
    }
    if (TryGetRealNode(nodeId) == nullptr)
    {
        return;
    }
    if (_workingModel != nullptr)
    {
        DiscardWorkingModel(false);
        if (_model == nullptr || nodeId == _model->viewNodeId || TryGetRealNode(nodeId) == nullptr)
        {
            UpdateNavigationHostState();
            if (_hWnd)
            {
                UpdateMenuState(_hWnd.get());
            }
            return;
        }
    }
    _refreshFailed  = false;
    _refreshRefused = false;

    BeginInputPresentation(InputPresentationKind::Navigation, handlerStartedAt);
    // _navStack is the view history kept for Back/Forward; Up does not use it (Up always means the structural parent).
    if (_navStack.size() >= kMaxNavigationHistory)
    {
        _navStack.erase(_navStack.begin());
    }
    _navStack.push_back(_model->viewNodeId);
    _navForwardStack.clear();
    _model->viewNodeId                = nodeId;
    _firstPositiveProjectionAttempted = false;
    _hasPublishedPositiveGeometry     = false;
    UpdateViewPathText();
    _layoutDirty = true;
    CancelLayoutBuild();
    InvalidateAnalysis();
    _layoutCandidateCache.clear();
    InvalidateStaticTreemapCache();

    if (_hWnd)
    {
        UpdateWindowTitle(_hWnd.get());
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool ViewerSpace::CanNavigateUp() const noexcept
{
    const Node* view = TryGetRealNode(_model->viewNodeId);
    if (view != nullptr && view->parentId != 0)
    {
        return true;
    }

    return _model->scanRootParentPath.has_value();
}

void ViewerSpace::UpdateMenuState(HWND hwnd, bool syncDxMenuBar) noexcept
{
    HMENU menu = _menuHandle ? _menuHandle.get() : (hwnd ? GetMenu(hwnd) : nullptr);
    if (! menu)
    {
        return;
    }

    CheckMenuRadioItem(menu,
                       IDM_VIEWERSPACE_WEIGHT_SIZE,
                       IDM_VIEWERSPACE_WEIGHT_COUNT,
                       _config.fileCountWeight ? IDM_VIEWERSPACE_WEIGHT_COUNT : IDM_VIEWERSPACE_WEIGHT_SIZE,
                       MF_BYCOMMAND);
    CheckMenuRadioItem(menu, IDM_VIEWERSPACE_COLOR_FOLDER, IDM_VIEWERSPACE_COLOR_NEWEST, IDM_VIEWERSPACE_COLOR_FOLDER + _config.colorMode, MF_BYCOMMAND);
    const auto selected      = _selectedGeneration == _displayGeneration ? ResolveAnalysisItem(_selectedItemId) : std::nullopt;
    const bool realSelection = selected.has_value() && ! selected.value().isSynthetic;
    EnableMenuItem(menu, IDM_VIEWERSPACE_COPY_PATH, static_cast<UINT>(MF_BYCOMMAND | ((realSelection || ! _selectedItemId) ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu,
                   IDM_VIEWERSPACE_EXCLUDE,
                   static_cast<UINT>(MF_BYCOMMAND |
                                     (realSelection && ! _workspaceRebinding && _model && _model->workspacePathIdentity.has_value() ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWERSPACE_RESTORE_ALL, static_cast<UINT>(MF_BYCOMMAND | (_excludedPaths.empty() ? MF_GRAYED : MF_ENABLED)));
    const std::wstring exclusionLabel =
        LoadStringResource(g_hInstance, realSelection && IsItemExcluded(_selectedItemId) ? IDS_VIEWERSPACE_RESTORE_ITEM : IDS_VIEWERSPACE_EXCLUDE);
    ModifyMenuW(menu, IDM_VIEWERSPACE_EXCLUDE, MF_BYCOMMAND | MF_STRING, IDM_VIEWERSPACE_EXCLUDE, exclusionLabel.c_str());
    const bool markedForCleanup =
        _cleanupGeneration == _displayGeneration && std::find(_cleanupMarks.begin(), _cleanupMarks.end(), _selectedItemId) != _cleanupMarks.end();
    const std::wstring markLabel = LoadStringResource(g_hInstance, markedForCleanup ? IDS_VIEWERSPACE_UNMARK_CLEANUP : IDS_VIEWERSPACE_MARK_CLEANUP);
    ModifyMenuW(menu, IDM_VIEWERSPACE_MARK_CLEANUP, MF_BYCOMMAND | MF_STRING, IDM_VIEWERSPACE_MARK_CLEANUP, markLabel.c_str());
    EnableMenuItem(menu,
                   IDM_VIEWERSPACE_MARK_CLEANUP,
                   static_cast<UINT>(MF_BYCOMMAND | (realSelection && (markedForCleanup || ! IsItemExcluded(_selectedItemId)) ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu,
                   IDM_VIEWERSPACE_OPEN_OTHER_PANE,
                   static_cast<UINT>(MF_BYCOMMAND | (realSelection && selected.value().isDirectory && _hostViewerActions ? MF_ENABLED : MF_GRAYED)));
    for (UINT command : std::array<UINT, 2u>{{IDM_VIEWERSPACE_SEND_RECYCLE, IDM_VIEWERSPACE_SEND_DELETE}})
        EnableMenuItem(
            menu, command, static_cast<UINT>(MF_BYCOMMAND | (_reviewingCleanup && ! _cleanupMarks.empty() && _hostViewerActions ? MF_ENABLED : MF_GRAYED)));
    CheckMenuItem(menu, IDM_VIEWERSPACE_REVIEW_CLEANUP, static_cast<UINT>(MF_BYCOMMAND | (_reviewingCleanup ? MF_CHECKED : MF_UNCHECKED)));
    const UINT state = CanNavigateUp() ? MF_ENABLED : MF_GRAYED;
    EnableMenuItem(menu, IDM_VIEWERSPACE_NAV_UP, MF_BYCOMMAND | state);
    EnableMenuItem(menu, IDM_VIEWERSPACE_NAV_BACK, static_cast<UINT>(MF_BYCOMMAND | (_navStack.empty() ? MF_GRAYED : MF_ENABLED)));
    EnableMenuItem(menu, IDM_VIEWERSPACE_NAV_FORWARD, static_cast<UINT>(MF_BYCOMMAND | (_navForwardStack.empty() ? MF_GRAYED : MF_ENABLED)));
    CheckMenuRadioItem(menu,
                       IDM_VIEWERSPACE_VIEW_EXPLORER,
                       IDM_VIEWERSPACE_VIEW_ANALYST,
                       ! _config.analysisPanelVisible           ? 0u
                       : _config.viewMode == ViewMode::Explorer ? IDM_VIEWERSPACE_VIEW_EXPLORER
                                                                : IDM_VIEWERSPACE_VIEW_ANALYST,
                       MF_BYCOMMAND);
    CheckMenuItem(menu, IDM_VIEWERSPACE_VIEW_SHOW_ANALYSIS, static_cast<UINT>(MF_BYCOMMAND | (! _config.analysisPanelVisible ? MF_CHECKED : MF_UNCHECKED)));
    if (syncDxMenuBar && _menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
    else if (hwnd)
    {
        DrawMenuBar(hwnd);
    }
}

float ViewerSpace::GetMenuBarHeightDip() const noexcept
{
    if (! _menuBarHost.GetHwnd())
    {
        return 0.0f;
    }

    const float dpi = _dpi > 0.0f ? _dpi : static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    return static_cast<float>(_menuBarHost.GetVisibleHeightPx()) * static_cast<float>(USER_DEFAULT_SCREEN_DPI) / dpi;
}

float ViewerSpace::GetHeaderTopDip() const noexcept
{
    return GetMenuBarHeightDip();
}

float ViewerSpace::GetHeaderBottomDip() const noexcept
{
    return GetHeaderTopDip() + kHeaderHeightDip + kNavigationHeightDip;
}

void ViewerSpace::NavigateUp() noexcept
{
    CancelDeferredStartScan();
    if (! CanNavigateUp())
    {
        return;
    }

    const auto destinationNode = [this]() noexcept -> uint32_t
    {
        const Node* const view = TryGetRealNode(_model->viewNodeId);
        return (view != nullptr && view->parentId != 0u) ? view->parentId : 0u;
    };
    uint32_t nextNode = destinationNode();
    if (nextNode != 0u && TryGetRealNode(nextNode) == nullptr)
    {
        return;
    }
    if (nextNode == 0u && ! _model->scanRootParentPath.has_value())
    {
        return;
    }

    if (_workingModel != nullptr)
    {
        DiscardWorkingModel(false);
        _refreshFailed  = false;
        _refreshRefused = false;
        if (_model == nullptr)
        {
            return;
        }
        nextNode = destinationNode();
        if ((nextNode != 0u && TryGetRealNode(nextNode) == nullptr) || (nextNode == 0u && ! _model->scanRootParentPath.has_value()))
        {
            UpdateNavigationHostState();
            if (_hWnd)
            {
                UpdateMenuState(_hWnd.get());
            }
            return;
        }
    }

    if (nextNode == 0u)
    {
        const std::wstring parentPath = _model->scanRootParentPath.value_or(std::wstring{});
        if (! parentPath.empty())
        {
            _refreshFailed  = false;
            _refreshRefused = false;
            StartScan(parentPath);
        }
        return;
    }
    if (nextNode == _model->viewNodeId)
    {
        return;
    }

    _refreshFailed  = false;
    _refreshRefused = false;
    NavigateTo(nextNode);
}

float ViewerSpace::DipFromPx(int px) const noexcept
{
    return Common::WindowSizing::PixelToDip(static_cast<float>(px), _dpi);
}

int ViewerSpace::PxFromDip(float dip) const noexcept
{
    return Common::WindowSizing::DipToPixelRounded(dip, static_cast<UINT>(_dpi));
}

double ViewerSpace::NowSeconds() const noexcept
{
    using clock             = std::chrono::steady_clock;
    static const auto start = clock::now();
    const auto now          = clock::now();
    return std::chrono::duration<double>(now - start).count();
}

HRESULT STDMETHODCALLTYPE ViewerSpace::Open(const ViewerOpenContext* context) noexcept
{
    if (context == nullptr || context->sizeBytes < sizeof(ViewerOpenContext) || context->fileSystem == nullptr || context->focusedPath == nullptr ||
        context->focusedPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const bool embeddedMode   = IsEmbeddedOpen(*context);
    const HWND embeddedParent = embeddedMode ? context->ownerWindow : nullptr;
    if (embeddedMode && (! embeddedParent || IsWindow(embeddedParent) == FALSE))
    {
        return E_INVALIDARG;
    }
    if (ShouldRecreateViewerWindow(embeddedMode, embeddedParent))
    {
        static_cast<void>(Close());
    }

    // The new content clock includes this viewer's controls and window creation. The historical
    // first-paint clock remains scan-start based for comparable existing measurements.
    _viewerOpenStartedAt    = std::chrono::steady_clock::now();
    _firstMeaningfulPaintUs = 0u;
#ifdef ENABLE_TESTS
    _presentedLayoutGeneration = 0u;
#endif

    // Close retires the prior model. Reopening the same public IViewer must restore
    // the constructor's empty-model invariant before synchronous window callbacks.
    if (_model == nullptr)
    {
        _displayModel = std::make_unique<ModelGeneration>();
        _model        = _displayModel.get();
        UpdateViewPathText();
    }

    if (_fileSystem)
    {
        wil::com_ptr<IUnknown> previousIdentity;
        wil::com_ptr<IUnknown> nextIdentity;
        if (! _fileSystem.try_query_to(previousIdentity.put()) || FAILED(context->fileSystem->QueryInterface(IID_PPV_ARGS(nextIdentity.put()))) ||
            previousIdentity.get() != nextIdentity.get())
        {
            RestoreExcludedItems();
            _cleanupMarks.clear();
            _reviewingCleanup = false;
            _showingProblems  = false;
        }
    }
    _fileSystem = context->fileSystem;

    _fileSystemName = context->fileSystemName ? context->fileSystemName : L"";
    _fileSystemShortId.clear();
    _fileSystemIsWin32 = true;

    wil::com_ptr<IInformations> fileSystemInfo;
    if (_fileSystem.try_query_to(fileSystemInfo.put()) && fileSystemInfo)
    {
        const PluginMetaData* metaData = nullptr;
        if (SUCCEEDED(fileSystemInfo->GetMetaData(&metaData)) && metaData != nullptr && metaData->shortId != nullptr)
        {
            _fileSystemShortId = metaData->shortId;
            _fileSystemIsWin32 = _fileSystemShortId == L"file";
        }
    }

    if (! _hWnd)
    {
        if (! RegisterWndClass(g_hInstance))
        {
            return E_FAIL;
        }

        HWND ownerWindow = context->ownerWindow;
        RECT ownerRc{};
        int x = CW_USEDEFAULT;
        int y = CW_USEDEFAULT;
        int w = 1000;
        int h = 700;
        if (embeddedMode)
        {
            RECT client{};
            if (GetClientRect(embeddedParent, &client) != 0)
            {
                x = 0;
                y = 0;
                w = std::max(1L, client.right - client.left);
                h = std::max(1L, client.bottom - client.top);
            }
        }
        else if (ownerWindow && GetWindowRect(ownerWindow, &ownerRc) != 0)
        {
            x = ownerRc.left;
            y = ownerRc.top;
            w = std::max(1L, ownerRc.right - ownerRc.left);
            h = std::max(1L, ownerRc.bottom - ownerRc.top);
        }

        HMENU loadedMenu = embeddedMode ? nullptr : Localization::LoadMenuResource(g_hInstance, IDR_VIEWERSPACE_MENU);
        wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(loadedMenu);
        const DWORD style               = embeddedMode ? (WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS) : (WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN);
        const bool previousEmbeddedMode = _embeddedMode;
        _embeddedMode                   = embeddedMode;
        const std::wstring initialTitle =
            embeddedMode ? std::wstring{} : (_metaName.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERSPACE_NAME) : _metaName);
        HWND window =
            CreateWindowExW(0, kClassName, initialTitle.c_str(), style, x, y, w, h, embeddedMode ? embeddedParent : nullptr, menu.get(), g_hInstance, this);
        if (! window)
        {
            _embeddedMode = previousEmbeddedMode;
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (! embeddedMode)
        {
            menu.release();
        }
        _hWnd.reset(window);

        ApplyThemeToWindow(_hWnd.get());
        AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
        if (! _embeddedMode)
        {
            _menuAttachResult = RedSalamander::ViewerMenuHost::Attach(g_hInstance, _hWnd.get(), _menuBarHost, _menuHandle);
        }
        ShowWindow(_hWnd.get(), embeddedMode ? SW_SHOWNA : SW_SHOWNORMAL);
        if (! embeddedMode)
        {
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }
    else
    {
        if (embeddedMode)
        {
            RECT client{};
            if (GetClientRect(embeddedParent, &client) != 0)
            {
                SetWindowPos(_hWnd.get(),
                             nullptr,
                             0,
                             0,
                             std::max(1L, client.right - client.left),
                             std::max(1L, client.bottom - client.top),
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            ShowWindow(_hWnd.get(), SW_SHOWNA);
        }
        else
        {
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }

    Debug::Perf::EmitDurationUs(L"viewer.space.open.setup_us", Debug::Perf::ElapsedUs(_viewerOpenStartedAt));
    StartScan(context->focusedPath);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::Close() noexcept
{
    const auto closeStartedAt = std::chrono::steady_clock::now();
    AddRef();
    const auto releaseSelf = wil::scope_exit([&]() noexcept { Release(); });
    _hWnd.reset();
    Debug::Perf::EmitDurationUs(L"viewer.space.close.destroy_hwnd_us", Debug::Perf::ElapsedUs(closeStartedAt));
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::SetTheme(const ViewerTheme* theme) noexcept
{
    if (theme == nullptr || theme->sizeBytes < sizeof(ViewerTheme))
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;
    // A live theme may carry default/stale DPI; the owned HWND is authoritative
    // for geometry and must agree with the hosted DxUi controls.
    _dpi = static_cast<float>(_hWnd ? GetDpiForWindow(_hWnd.get()) : (_theme.dpi == 0 ? USER_DEFAULT_SCREEN_DPI : _theme.dpi));

    RequestViewerSpaceClassBackgroundColor(ColorRefFromArgb(_theme.backgroundArgb));
    ApplyPendingViewerSpaceClassBackgroundBrush(_hWnd.get());

    _layoutDirty = true;
    _layoutCandidateCache.clear();
#ifdef ENABLE_TESTS
    _rendererFailureStage = kRendererDiscardTheme;
    _rendererFailureHr    = 0u;
#endif
    DiscardDirect2D();

    if (_hWnd)
    {
        ApplyThemeToWindow(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    return S_OK;
}
