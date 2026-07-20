#include "Framework.h"

#include "FindFilesWindow.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <condition_variable>
#include <cstdint>
#include <exception>
#include <filesystem>
#include <format>
#include <iterator>
#include <limits>
#include <memory>
#include <mutex>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include <shellapi.h>
#include <shlobj.h>
#include <windowsx.h>

#include "CommandRegistry.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "FileSystemPluginManager.h"
#include "FluentIcons.h"
#include "FolderWindow.h"
#include "Helpers.h"
#include "HostServices.h"
#include "IconCache.h"
#include "NavigationLocation.h"
#include "NavigationView.h"
#include "PlugInterfaces/Factory.h"
#include "SearchFallbackEngine.h"
#include "SearchServiceBroker.h"
#include "ShortcutDefaults.h"
#include "ShortcutManager.h"
#include "UiMetrics.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "WindowSizing.h"
#include "resource.h"

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ButtonVariant;
using RedSalamander::DxUi::Checkbox;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::ContextMenu;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridCellKind;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::GridVisualMode;
using RedSalamander::DxUi::IDxGridDelegate;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::IsContextMenuDiagnosticsEnabled;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::MenuFlyoutItem;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::StatusStrip;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::TraceContextMenuDiagnostics;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kFindFilesWindowClassName[] = L"RedSalamander.FindFilesWindow";
constexpr wchar_t kFindFilesWindowId[]        = L"FindFilesWindow";

constexpr uint32_t kMaxSnippetCharacters              = 160u;
constexpr uint64_t kMaxContentBytes                   = 64ull * 1024ull * 1024ull;
constexpr size_t kMaxRecentEntries                    = 10u;
constexpr size_t kBatchSize                           = 32u;
constexpr size_t kProgressFlushMinBatchSize           = 8u;
constexpr auto kProgressFlushMaxAge                   = std::chrono::milliseconds(40);
constexpr UINT_PTR kStatusRefreshTimerId              = 1u;
constexpr UINT_PTR kDeferredResultsRefreshTimerId     = 2u;
constexpr UINT kStatusRefreshTimerIntervalMs          = 250u;
constexpr uint64_t kStatusStallThresholdMs            = 5000u;
constexpr size_t kResultsDrainMaxMessages             = 8u;
constexpr size_t kResultsDrainMaxRecords              = 256u;
constexpr size_t kInteractiveResultsRefreshRecords    = 256u;
constexpr uint64_t kInteractiveResultsRefreshMaxAgeMs = 100u;
constexpr uint64_t kPendingResultRemovalMaxAgeMs      = 24ull * 60ull * 60ull * 1000ull;
constexpr uint32_t kFindShortcutCtrl                  = 0x1u;
constexpr uint32_t kFindShortcutShift                 = 0x2u;
constexpr uint32_t kFindShortcutAlt                   = 0x4u;
constexpr int kFindResultMenuClickedCommandBase       = 0x5200;
constexpr int kFindResultMenuSelectionCommandBase     = 0x5300;

enum class FindResultMenuTarget : uint8_t
{
    ClickedItem,
    Selection,
};

enum class FindResultMenuAction : uint8_t
{
    Open = 1u,
    GoToFolder,
    View,
    AlternateView,
    Edit,
    AlternateEdit,
    ClipboardCopy,
    ClipboardCut,
    CopyToDestination,
    MoveToDestination,
    Delete,
    PermanentDelete,
};

struct FindResultMenuCommand final
{
    FindResultMenuTarget target = FindResultMenuTarget::ClickedItem;
    FindResultMenuAction action = FindResultMenuAction::Open;
};

[[nodiscard]] int EncodeFindResultMenuCommand(FindResultMenuTarget target, FindResultMenuAction action) noexcept
{
    const int base = target == FindResultMenuTarget::ClickedItem ? kFindResultMenuClickedCommandBase : kFindResultMenuSelectionCommandBase;
    return base + static_cast<int>(action);
}

[[nodiscard]] std::optional<FindResultMenuCommand> DecodeFindResultMenuCommand(int commandId) noexcept
{
    if (commandId > kFindResultMenuClickedCommandBase && commandId < kFindResultMenuClickedCommandBase + 100)
    {
        return FindResultMenuCommand{.target = FindResultMenuTarget::ClickedItem,
                                     .action = static_cast<FindResultMenuAction>(commandId - kFindResultMenuClickedCommandBase)};
    }
    if (commandId > kFindResultMenuSelectionCommandBase && commandId < kFindResultMenuSelectionCommandBase + 100)
    {
        return FindResultMenuCommand{.target = FindResultMenuTarget::Selection,
                                     .action = static_cast<FindResultMenuAction>(commandId - kFindResultMenuSelectionCommandBase)};
    }
    return std::nullopt;
}

[[nodiscard]] std::wstring FindVkToMenuShortcutText(uint32_t vk) noexcept
{
    vk &= 0xFFu;

    if (vk >= VK_F1 && vk <= VK_F24)
    {
        return std::format(L"F{}", static_cast<unsigned>(vk - VK_F1 + 1u));
    }

    if ((vk >= static_cast<uint32_t>('0') && vk <= static_cast<uint32_t>('9')) || (vk >= static_cast<uint32_t>('A') && vk <= static_cast<uint32_t>('Z')))
    {
        return std::wstring(1u, static_cast<wchar_t>(vk));
    }

    UINT scanCode = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (scanCode == 0u)
    {
        return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
    }

    bool extended = false;
    switch (vk)
    {
        case VK_LEFT:
        case VK_UP:
        case VK_RIGHT:
        case VK_DOWN:
        case VK_PRIOR:
        case VK_NEXT:
        case VK_END:
        case VK_HOME:
        case VK_INSERT:
        case VK_DELETE: extended = true; break;
        default: break;
    }

    LPARAM lParam = static_cast<LPARAM>(scanCode) << 16;
    if (extended)
    {
        lParam |= (1 << 24);
    }

    wchar_t keyName[64]{};
    const int length = GetKeyNameTextW(static_cast<LONG>(lParam), keyName, static_cast<int>(std::size(keyName)));
    if (length > 0)
    {
        return std::wstring(keyName, static_cast<size_t>(length));
    }

    return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
}

[[nodiscard]] std::wstring FormatFindMenuChordText(uint32_t vk, uint32_t modifiers) noexcept
{
    std::wstring result;
    const auto appendPart = [&](std::wstring_view part)
    {
        if (part.empty())
        {
            return;
        }
        if (! result.empty())
        {
            result.append(L"+");
        }
        result.append(part);
    };

    const uint32_t maskedMods = modifiers & 0x7u;
    if ((maskedMods & ShortcutManager::kModCtrl) != 0u)
    {
        appendPart(LoadEmbeddedStringResource(nullptr, IDS_MOD_CTRL));
    }
    if ((maskedMods & ShortcutManager::kModAlt) != 0u)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_ALT));
    }
    if ((maskedMods & ShortcutManager::kModShift) != 0u)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }

    appendPart(FindVkToMenuShortcutText(vk));
    return result;
}

[[nodiscard]] uint32_t ToShortcutManagerModifiers(uint32_t findModifiers) noexcept
{
    uint32_t modifiers = 0u;
    if ((findModifiers & kFindShortcutCtrl) != 0u)
    {
        modifiers |= ShortcutManager::kModCtrl;
    }
    if ((findModifiers & kFindShortcutAlt) != 0u)
    {
        modifiers |= ShortcutManager::kModAlt;
    }
    if ((findModifiers & kFindShortcutShift) != 0u)
    {
        modifiers |= ShortcutManager::kModShift;
    }
    return modifiers;
}

#ifdef ENABLE_TESTS
constexpr UINT kFindFilesWindowDebugMessage = WM_APP + 0x74u;

enum class FindFilesWindowDebugCommand : WPARAM
{
    Configure = 1u,
    SetOptions,
    StartSearch,
    FocusTarget,
    SetComboText,
    GetSnapshot,
    ResizeVisibleResultColumn,
    ApplyResultsLayoutFromSettings,
    SelectResults,
    PostStaleSearchPayloads,
};

struct FindFilesWindowDebugConfigurePayload final
{
    std::wstring rootPath;
    std::wstring namePattern;
    std::wstring contentPattern;
    Common::Settings::SearchNameMode nameMode       = Common::Settings::SearchNameMode::Wildcard;
    Common::Settings::SearchContentMode contentMode = Common::Settings::SearchContentMode::Disabled;
    bool result                                     = false;
};

struct FindFilesWindowDebugSetOptionsPayload final
{
    bool recursive          = false;
    bool includeFiles       = false;
    bool includeDirectories = false;
    bool preferIndex        = false;
    bool wantSnippets       = false;
    bool result             = false;
};

struct FindFilesWindowDebugStartSearchPayload final
{
    FindFilesDebugOperation operation = FindFilesDebugOperation::Find;
    bool result                       = false;
};

struct FindFilesWindowDebugFocusTargetPayload final
{
    FindFilesDebugFocusTarget target = FindFilesDebugFocusTarget::None;
    bool result                      = false;
};

struct FindFilesWindowDebugSetComboTextPayload final
{
    FindFilesDebugFocusTarget target = FindFilesDebugFocusTarget::None;
    std::wstring text;
    bool result = false;
};

struct FindFilesWindowDebugSnapshotPayload final
{
    FindFilesDebugSnapshot* snapshot = nullptr;
    bool result                      = false;
};

struct FindFilesWindowDebugResizeVisibleResultColumnPayload final
{
    size_t visibleIndex = 0u;
    float deltaDip      = 0.0f;
    bool result         = false;
};

struct FindFilesWindowDebugApplyResultsLayoutFromSettingsPayload final
{
    bool result = false;
};

struct FindFilesWindowDebugSelectResultsPayload final
{
    std::vector<std::wstring> fullPaths;
    bool result = false;
};

struct FindFilesWindowDebugPostStaleSearchPayloadsPayload final
{
    std::wstring fullPath;
    bool result = false;
};

std::atomic<bool> g_debugSearchRunBlockEnabled{false};
std::atomic<bool> g_debugSearchRunBlocked{false};
std::atomic<bool> g_debugSearchRunRelease{false};

void DebugMaybeBlockFindFilesWindowSearchRun() noexcept
{
    if (! g_debugSearchRunBlockEnabled.load(std::memory_order_acquire))
    {
        return;
    }

    g_debugSearchRunBlocked.store(true, std::memory_order_release);
    while (g_debugSearchRunBlockEnabled.load(std::memory_order_acquire) && ! g_debugSearchRunRelease.load(std::memory_order_acquire))
    {
        ::Sleep(10);
    }
}
#endif
constexpr std::wstring_view kBuiltinLocalFileSystemId = L"builtin/file-system";
constexpr std::wstring_view kStatusSpinnerFrames[]    = {L"|", L"/", L"-", L"\\"};

using SteadyClock = std::chrono::steady_clock;

[[nodiscard]] std::wstring TruncateMiddleText(std::wstring_view text, const size_t maxCharacters) noexcept
{
    if (text.size() <= maxCharacters)
    {
        return std::wstring(text);
    }

    if (maxCharacters <= 3u)
    {
        return std::wstring(text.substr(0u, maxCharacters));
    }

    constexpr std::wstring_view kEllipsis = L"...";
    const size_t availableCharacters      = maxCharacters - kEllipsis.size();
    const size_t prefixCharacters         = availableCharacters / 2u;
    const size_t suffixCharacters         = availableCharacters - prefixCharacters;

    std::wstring truncated;
    truncated.reserve(maxCharacters);
    truncated.append(text.substr(0u, prefixCharacters));
    truncated.append(kEllipsis);
    truncated.append(text.substr(text.size() - suffixCharacters));
    return truncated;
}

inline void EmitPerfCount(std::wstring_view name, uint64_t value = 1u) noexcept
{
    Debug::Perf::Emit(name, L"", 0u, value, 0u, S_OK);
}

[[nodiscard]] uint64_t ElapsedUsSince(SteadyClock::time_point start) noexcept
{
    const auto elapsed = SteadyClock::now() - start;
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count());
}

using CreateFactoryFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

enum class SearchOperation : uint8_t
{
    Find,
    Append,
    Intersect,
    Subtract,
};

[[nodiscard]] const wchar_t* TraceSearchOperationName(SearchOperation operation) noexcept
{
    switch (operation)
    {
        case SearchOperation::Find: return L"Find";
        case SearchOperation::Append: return L"Append";
        case SearchOperation::Intersect: return L"Intersect";
        case SearchOperation::Subtract: return L"Subtract";
    }
    return L"unknown";
}

[[nodiscard]] const wchar_t* TraceFindWindowMessageName(UINT message) noexcept
{
    switch (message)
    {
        case WM_MOUSEMOVE: return L"WM_MOUSEMOVE";
        case WM_LBUTTONDOWN: return L"WM_LBUTTONDOWN";
        case WM_LBUTTONUP: return L"WM_LBUTTONUP";
        case WM_LBUTTONDBLCLK: return L"WM_LBUTTONDBLCLK";
        case WM_RBUTTONDOWN: return L"WM_RBUTTONDOWN";
        case WM_RBUTTONUP: return L"WM_RBUTTONUP";
        case WM_MOUSEACTIVATE: return L"WM_MOUSEACTIVATE";
        case WM_SETCURSOR: return L"WM_SETCURSOR";
        case WM_NCLBUTTONDOWN: return L"WM_NCLBUTTONDOWN";
        case WM_NCLBUTTONUP: return L"WM_NCLBUTTONUP";
        case WM_NCLBUTTONDBLCLK: return L"WM_NCLBUTTONDBLCLK";
        case WM_CAPTURECHANGED: return L"WM_CAPTURECHANGED";
        case WM_CANCELMODE: return L"WM_CANCELMODE";
        case WM_SETFOCUS: return L"WM_SETFOCUS";
        case WM_KILLFOCUS: return L"WM_KILLFOCUS";
        case WM_ACTIVATE: return L"WM_ACTIVATE";
        case WM_NCACTIVATE: return L"WM_NCACTIVATE";
        case WM_ENTERMENULOOP: return L"WM_ENTERMENULOOP";
        case WM_EXITMENULOOP: return L"WM_EXITMENULOOP";
        default: return L"message";
    }
}

[[nodiscard]] bool ShouldTraceFindWindowMessage(UINT message) noexcept
{
    switch (message)
    {
        case WM_MOUSEMOVE:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_MOUSEACTIVATE:
        case WM_SETCURSOR:
        case WM_NCLBUTTONDOWN:
        case WM_NCLBUTTONUP:
        case WM_NCLBUTTONDBLCLK:
        case WM_CAPTURECHANGED:
        case WM_CANCELMODE:
        case WM_SETFOCUS:
        case WM_KILLFOCUS:
        case WM_ACTIVATE:
        case WM_NCACTIVATE:
        case WM_ENTERMENULOOP:
        case WM_EXITMENULOOP: return true;
        default: return false;
    }
}

[[nodiscard]] const wchar_t* TraceFindQueueMessageName(UINT message) noexcept
{
    switch (message)
    {
        case WndMsg::kFindSearchResults: return L"kFindSearchResults";
        case WndMsg::kFindSearchProgress: return L"kFindSearchProgress";
        case WndMsg::kFindSearchComplete: return L"kFindSearchComplete";
        case WndMsg::kFindSearchDeferredRefresh: return L"kFindSearchDeferredRefresh";
        case WndMsg::kFindShowActionMenu: return L"kFindShowActionMenu";
        default: return TraceFindWindowMessageName(message);
    }
}

[[nodiscard]] bool IsNextThreadQueueMessage(HWND targetHwnd, UINT targetMessage, MSG* queuedMessage = nullptr) noexcept
{
    MSG nextMessage{};
    if (PeekMessageW(&nextMessage, nullptr, 0, 0, PM_NOREMOVE) == 0)
    {
        return false;
    }

    if (queuedMessage)
    {
        *queuedMessage = nextMessage;
    }

    return nextMessage.hwnd == targetHwnd && nextMessage.message == targetMessage;
}

[[nodiscard]] bool ShouldResolveScreenHitWindowForFindTrace(UINT message) noexcept
{
    switch (message)
    {
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_MOUSEACTIVATE:
        case WM_NCLBUTTONDOWN:
        case WM_NCLBUTTONUP:
        case WM_NCLBUTTONDBLCLK: return true;
        default: return false;
    }
}

template <typename... Args> void TraceFindContextMenuDiagnostics(std::wstring_view eventName, std::wformat_string<Args...> format, Args&&... args) noexcept
{
    if (! IsContextMenuDiagnosticsEnabled())
    {
        return;
    }

    try
    {
        TraceContextMenuDiagnostics(eventName, std::format(format, std::forward<Args>(args)...));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::format_error&)
    {
        // Menu tracing is diagnostic only; formatting failure must not change Find window input behavior.
        TraceContextMenuDiagnostics(eventName, L"formatting failed");
    }
}

enum ControlId : int
{
    kRootLabelId = 4100,
    kRootComboId,
    kNameLabelId,
    kNameComboId,
    kNameModeComboId,
    kContentLabelId,
    kContentComboId,
    kContentModeComboId,
    kRecursiveCheckId,
    kIncludeFilesCheckId,
    kIncludeDirectoriesCheckId,
    kFollowSymlinksCheckId,
    kMatchCaseNameCheckId,
    kMatchCaseContentCheckId,
    kPreferIndexCheckId,
    kWantSnippetsCheckId,
    kFindButtonId,
    kAppendButtonId,
    kIntersectButtonId,
    kSubtractButtonId,
    kCancelButtonId,
    kOpenButtonId,
    kParentButtonId,
    kStatusTextId,
    kResultsListId,
};

enum ColumnIndex : int
{
    kColumnName = 0,
    kColumnPath,
    kColumnSize,
    kColumnModified,
    kColumnAttributes,
    kColumnSnippet,
};

enum FindActionMenuCommand : int
{
    kFindActionMenuFind = 1,
    kFindActionMenuIntersect,
    kFindActionMenuSubtract,
    kFindActionMenuAppend,
};

struct FindResultRecord
{
    uint64_t arrivalOrdinal = 0u;
    uint64_t stableRowId    = 0u;
    std::wstring key;
    std::wstring pluginId;
    std::wstring pluginShortId;
    std::wstring instanceContext;
    std::wstring fullPath;
    std::wstring relativePath;
    std::wstring displayPath;
    std::wstring displayName;
    std::wstring previewText;
    uint32_t folderViewRainbowHash32 = 0u;
    int iconIndex                    = -1;
    unsigned long fileAttributes     = 0;
    int64_t lastWriteTime            = 0;
    int64_t endOfFile                = 0;
    uint32_t matchedBy               = 0;
};

struct FindSearchResultsPayload
{
    std::vector<FindResultRecord> results;
    uint64_t epoch = 0u;
    SteadyClock::time_point enqueuedAt{};
};

struct SearchServiceStatusSnapshot
{
    bool available                                              = false;
    HRESULT hr                                                  = S_OK;
    LocalSearchIndexCore::StoreState storeState                 = LocalSearchIndexCore::StoreState::Unknown;
    LocalSearchIndexCore::SyncPhase syncPhase                   = LocalSearchIndexCore::SyncPhase::Idle;
    LocalSearchIndexCore::QueryExecutionMode queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::Unknown;
    LocalSearchIndexCore::FallbackReason fallbackReason         = LocalSearchIndexCore::FallbackReason::None;
    uint64_t completedRoots                                     = 0u;
    uint64_t totalRoots                                         = 0u;
    std::wstring activeRoot;
};

struct FindSearchProgressPayload
{
    uint64_t epoch                  = 0u;
    FileSystemSearchPhase phase     = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    FileSystemSearchBackend backend = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
    uint32_t warningFlags           = FILESYSTEM_SEARCH_WARNING_NONE;
    HRESULT statusHint              = S_OK;
    uint64_t scannedDirectories     = 0;
    uint64_t scannedFiles           = 0;
    uint64_t candidateFiles         = 0;
    uint64_t matchedEntries         = 0;
    std::wstring currentPath;
    SearchServiceStatusSnapshot serviceStatus;
    SteadyClock::time_point enqueuedAt{};
};

struct FindSearchCompletePayload
{
    uint64_t epoch                  = 0u;
    HRESULT hr                      = S_OK;
    FileSystemSearchBackend backend = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
    uint32_t warningFlags           = FILESYSTEM_SEARCH_WARNING_NONE;
    uint64_t scannedDirectories     = 0;
    uint64_t scannedFiles           = 0;
    uint64_t candidateFiles         = 0;
    uint64_t matchedEntries         = 0;
};

struct SearchRequest
{
    FindFilesPaneContext context;
    std::shared_ptr<wil::unique_hmodule> contextModulePin;
    std::wstring rootPath;
    std::wstring namePattern;
    std::wstring contentPattern;
    FileSystemSearchFlags flags             = FILESYSTEM_SEARCH_NONE;
    FileSystemSearchNameMode nameMode       = FILESYSTEM_SEARCH_NAME_DISABLED;
    FileSystemSearchContentMode contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;
    uint64_t maxResults                     = 0;
    uint64_t maxContentBytesPerFile         = 0;
    uint32_t maxSnippetCharacters           = 0;
};

[[nodiscard]] std::vector<RedSalamander::DxUi::GridColumnLayoutEntry> ConvertColumnLayout(const std::vector<Common::Settings::GridColumnLayoutEntry>& layout)
{
    std::vector<RedSalamander::DxUi::GridColumnLayoutEntry> converted;
    converted.reserve(layout.size());
    for (const auto& entry : layout)
    {
        if (entry.columnId.empty())
        {
            continue;
        }

        converted.push_back(RedSalamander::DxUi::GridColumnLayoutEntry{
            .columnId     = entry.columnId,
            .displayIndex = entry.displayIndex,
            .widthDip     = entry.widthDip,
        });
    }
    return converted;
}

[[nodiscard]] std::vector<Common::Settings::GridColumnLayoutEntry> ConvertColumnLayout(const std::vector<RedSalamander::DxUi::GridColumnLayoutEntry>& layout)
{
    std::vector<Common::Settings::GridColumnLayoutEntry> converted;
    converted.reserve(layout.size());
    for (const auto& entry : layout)
    {
        if (entry.columnId.empty())
        {
            continue;
        }

        converted.push_back(Common::Settings::GridColumnLayoutEntry{
            .columnId     = entry.columnId,
            .displayIndex = static_cast<uint32_t>(entry.displayIndex),
            .widthDip     = entry.widthDip,
        });
    }
    return converted;
}

struct ResultListMutation
{
    size_t index  = 0;
    bool inserted = false;
};


[[nodiscard]] uint64_t MakeResultStableId(const FindResultRecord& record) noexcept
{
    return (static_cast<uint64_t>(StableHash32(record.key)) << 32u) ^ static_cast<uint64_t>(StableHash32(record.relativePath));
}

[[nodiscard]] std::wstring_view GetResultRelativeFolderPath(std::wstring_view relativePath) noexcept
{
    const size_t separatorIndex = relativePath.find_last_of(L"\\/");
    if (separatorIndex == std::wstring_view::npos)
    {
        return {};
    }

    return relativePath.substr(0u, separatorIndex);
}

[[nodiscard]] uint32_t MakeFindResultFolderViewRainbowHash32(const FindResultRecord& record)
{
    std::wstring containingFolder;
    std::wstring fallbackDisplayName;

    if (! record.fullPath.empty())
    {
        const std::filesystem::path fullPath(record.fullPath);
        containingFolder    = fullPath.parent_path().native();
        fallbackDisplayName = fullPath.filename().native();
    }

    if (containingFolder.empty() && ! record.displayPath.empty())
    {
        containingFolder = record.displayPath;
    }
    if (containingFolder.empty() && ! record.relativePath.empty())
    {
        containingFolder = std::wstring(GetResultRelativeFolderPath(record.relativePath));
    }
    if (containingFolder.empty())
    {
        containingFolder = record.instanceContext.empty() ? record.pluginId : record.instanceContext;
    }

    std::wstring_view displayName = record.displayName;
    if (displayName.empty())
    {
        displayName = fallbackDisplayName;
    }
    if (displayName.empty())
    {
        displayName = record.key;
    }

    return FolderItemStableHash32(containingFolder, displayName);
}

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

[[nodiscard]] std::wstring_view TrimTrailingPathSeparators(std::wstring_view path) noexcept
{
    size_t end = path.size();
    while (end > 0u && IsPathSeparator(path[end - 1u]))
    {
        if (end == 3u && path[1u] == L':')
        {
            break;
        }
        --end;
    }
    return path.substr(0u, end);
}

[[nodiscard]] bool StartsWithNoCase(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (prefix.size() > text.size() || prefix.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }

    return ::CompareStringOrdinal(text.data(), static_cast<int>(prefix.size()), prefix.data(), static_cast<int>(prefix.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] std::wstring_view BuildRelativePathFromRoot(std::wstring_view rootPath, std::wstring_view fullPath) noexcept
{
    const std::wstring_view root = TrimTrailingPathSeparators(rootPath);
    if (root.empty() || fullPath.empty() || ! StartsWithNoCase(fullPath, root))
    {
        return {};
    }

    if (fullPath.size() == root.size())
    {
        return {};
    }

    if (IsPathSeparator(root.back()))
    {
        return fullPath.substr(root.size());
    }

    if (! IsPathSeparator(fullPath[root.size()]))
    {
        return {};
    }

    return fullPath.substr(root.size() + 1u);
}

[[nodiscard]] std::wstring BuildResultDisplayPath(std::wstring_view relativePath, std::wstring_view rootPath, std::wstring_view fullPath)
{
    std::wstring_view folderPath = GetResultRelativeFolderPath(relativePath);
    if (! folderPath.empty())
    {
        return std::wstring(folderPath);
    }

    const std::wstring_view fallbackRelativePath = BuildRelativePathFromRoot(rootPath, fullPath);
    folderPath                                   = GetResultRelativeFolderPath(fallbackRelativePath);
    return std::wstring(folderPath);
}

[[nodiscard]] std::wstring BuildResultIconText(const FindResultRecord& record)
{
    const wchar_t glyph = (record.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u ? FluentIcons::kFolder : FluentIcons::kDocument;
    return std::wstring(1u, glyph);
}

[[nodiscard]] int ResolveFindResultIconIndex(const FindResultRecord& record) noexcept
{
    const bool isDirectory = (record.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
    auto& iconCache        = IconCache::GetInstance();
    if (isDirectory)
    {
        if (NavigationLocation::LooksLikeWindowsAbsolutePath(record.fullPath) && IconCache::IsSpecialFolder(record.fullPath))
        {
            const auto pathIcon = iconCache.QuerySysIconIndexForPath(record.fullPath.c_str(), record.fileAttributes, true);
            if (pathIcon.has_value())
            {
                return pathIcon.value();
            }
        }

        const auto folderIcon = iconCache.GetOrQueryIconIndexByExtension(L"<directory>", FILE_ATTRIBUTE_DIRECTORY);
        return folderIcon.value_or(-1);
    }

    const std::filesystem::path fullPath(record.fullPath);
    const std::wstring extension = fullPath.extension().wstring();
    if (iconCache.RequiresPerFileLookup(extension) && NavigationLocation::LooksLikeWindowsAbsolutePath(record.fullPath))
    {
        const auto pathIcon = iconCache.QuerySysIconIndexForPath(record.fullPath.c_str(), record.fileAttributes, true);
        if (pathIcon.has_value())
        {
            return pathIcon.value();
        }
    }

    const auto extensionIcon = iconCache.GetOrQueryIconIndexByExtension(extension, FILE_ATTRIBUTE_NORMAL);
    if (extensionIcon.has_value())
    {
        return extensionIcon.value();
    }

    if (NavigationLocation::LooksLikeWindowsAbsolutePath(record.fullPath))
    {
        const auto pathIcon = iconCache.QuerySysIconIndexForPath(record.fullPath.c_str(), record.fileAttributes, true);
        if (pathIcon.has_value())
        {
            return pathIcon.value();
        }
    }

    return -1;
}

[[nodiscard]] std::wstring FormatFileSize(int64_t sizeBytes) noexcept;
[[nodiscard]] std::wstring FormatFileTimeValue(int64_t fileTimeTicks) noexcept;
[[nodiscard]] std::wstring FormatAttributes(unsigned long attributes) noexcept;

class FindResultsGridModel final : public IDxGridModel
{
public:
    explicit FindResultsGridModel(const AppTheme& theme) : _theme(&theme)
    {
        RebuildColumns();
    }

    void SetRows(const std::vector<FindResultRecord>* rows) noexcept
    {
        _rows = rows;
    }

    void SetShowSnippetColumn(bool showSnippet) noexcept
    {
        if (_showSnippet == showSnippet)
        {
            return;
        }

        _showSnippet = showSnippet;
        RebuildColumns();
    }

    [[nodiscard]] bool ShowsSnippetColumn() const noexcept
    {
        return _showSnippet;
    }

    void ApplyColumnLayoutDefaults(std::span<const RedSalamander::DxUi::GridColumnLayoutEntry> layout) noexcept
    {
        if (_columns.empty() || layout.empty())
        {
            return;
        }

        for (const auto& entry : layout)
        {
            if (entry.columnId.empty() || ! std::isfinite(entry.widthDip) || entry.widthDip <= 0.0f)
            {
                continue;
            }

            const auto it = std::find_if(_columns.begin(), _columns.end(), [&](const GridColumnDesc& column) noexcept { return column.id == entry.columnId; });
            if (it == _columns.end())
            {
                continue;
            }

            it->widthDip = std::max(it->minWidthDip, entry.widthDip);
        }
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows ? _rows->size() : 0u;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return _columns.at(columnIndex);
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (! _rows || rowIndex >= _rows->size())
        {
            return;
        }

        const FindResultRecord& record = _rows->at(rowIndex);
        switch (columnIndex)
        {
            case kColumnName:
                outCell.kind      = GridCellKind::IconText;
                outCell.iconText  = BuildResultIconText(record);
                outCell.iconIndex = record.iconIndex;
                outCell.text      = record.displayName;
                break;
            case kColumnPath: outCell.text = record.displayPath; break;
            case kColumnSize:
                outCell.text          = FormatFileSize(record.endOfFile);
                outCell.textAlignment = DWRITE_TEXT_ALIGNMENT_TRAILING;
                break;
            case kColumnModified: outCell.text = FormatFileTimeValue(record.lastWriteTime); break;
            case kColumnAttributes: outCell.text = FormatAttributes(record.fileAttributes); break;
            case kColumnSnippet:
                outCell.text      = record.previewText;
                outCell.multiline = true;
                break;
            default: break;
        }
    }

    [[nodiscard]] GridRowStyle GetRowStyle(size_t rowIndex) const override
    {
        GridRowStyle style{};
        if (_theme && _theme->menu.rainbowMode && _rows && rowIndex < _rows->size())
        {
            const FindResultRecord& record = _rows->at(rowIndex);
            style.rainbowSeed              = record.key;
            style.folderViewRainbowHash32  = record.folderViewRainbowHash32;
        }
        return style;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return (_rows && rowIndex < _rows->size()) ? _rows->at(rowIndex).stableRowId : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (! _rows)
        {
            return std::nullopt;
        }

        for (size_t rowIndex = 0; rowIndex < _rows->size(); ++rowIndex)
        {
            if (_rows->at(rowIndex).stableRowId == rowId)
            {
                return rowIndex;
            }
        }
        return std::nullopt;
    }

private:
    void RebuildColumns()
    {
        _columns = {
            {L"name", LoadStringResource(nullptr, IDS_FIND_COLUMN_NAME), 220.0f, 120.0f, RedSalamander::DxUi::GridColumnKind::Text, true, false},
            {L"path", LoadStringResource(nullptr, IDS_FIND_COLUMN_PATH), 360.0f, 180.0f, RedSalamander::DxUi::GridColumnKind::Text, true, false},
            {L"size",
             LoadStringResource(nullptr, IDS_FIND_COLUMN_SIZE),
             110.0f,
             96.0f,
             RedSalamander::DxUi::GridColumnKind::Text,
             true,
             false,
             DWRITE_TEXT_ALIGNMENT_TRAILING},
            {L"modified", LoadStringResource(nullptr, IDS_FIND_COLUMN_MODIFIED), 150.0f, 120.0f, RedSalamander::DxUi::GridColumnKind::Text, true, false},
            {L"attributes", LoadStringResource(nullptr, IDS_FIND_COLUMN_ATTR), 80.0f, 72.0f, RedSalamander::DxUi::GridColumnKind::Text, true, false},
        };

        if (_showSnippet)
        {
            GridColumnDesc snippetColumn{};
            snippetColumn.id          = L"snippet";
            snippetColumn.title       = LoadStringResource(nullptr, IDS_FIND_COLUMN_SNIPPET);
            snippetColumn.widthDip    = 260.0f;
            snippetColumn.minWidthDip = 140.0f;
            snippetColumn.sortable    = true;
            snippetColumn.multiline   = true;
            _columns.push_back(std::move(snippetColumn));
        }
    }

    const AppTheme* _theme                     = nullptr;
    const std::vector<FindResultRecord>* _rows = nullptr;
    bool _showSnippet                          = false;
    std::vector<GridColumnDesc> _columns;
};

[[nodiscard]] uint64_t GetTickAgeMs(uint64_t tickMs) noexcept
{
    if (tickMs == 0u)
    {
        return 0u;
    }

    const ULONGLONG now = ::GetTickCount64();
    return now >= tickMs ? static_cast<uint64_t>(now - tickMs) : 0u;
}

[[nodiscard]] std::wstring_view GetStatusSpinnerFrame() noexcept
{
    constexpr size_t kFrameCount = sizeof(kStatusSpinnerFrames) / sizeof(kStatusSpinnerFrames[0]);
    const size_t frameIndex      = static_cast<size_t>((::GetTickCount64() / kStatusRefreshTimerIntervalMs) % kFrameCount);
    return kStatusSpinnerFrames[frameIndex];
}

[[nodiscard]] std::wstring FormatStatusAgeText(uint64_t elapsedMs) noexcept
{
    if (elapsedMs < 1000u)
    {
        return std::format(L"{} ms", elapsedMs);
    }

    const uint64_t totalSeconds = elapsedMs / 1000u;
    if (totalSeconds < 60u)
    {
        return std::format(L"{} s", totalSeconds);
    }

    const uint64_t minutes = totalSeconds / 60u;
    const uint64_t seconds = totalSeconds % 60u;
    if (minutes < 60u)
    {
        return std::format(L"{}m {:02}s", minutes, seconds);
    }

    const uint64_t hours            = minutes / 60u;
    const uint64_t remainingMinutes = minutes % 60u;
    return std::format(L"{}h {:02}m", hours, remainingMinutes);
}

[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindFileSystemPluginById(std::wstring_view pluginId) noexcept
{
    return FileSystemPluginManager::GetInstance().FindPluginById(pluginId);
}

[[nodiscard]] bool TryCreateIsolatedFileSystem(std::wstring_view pluginId,
                                               wil::com_ptr<IFileSystem>& outFileSystem,
                                               std::wstring& outPluginId,
                                               std::wstring& outPluginShortId,
                                               std::shared_ptr<wil::unique_hmodule>& outModulePin) noexcept
{
    outFileSystem.reset();
    outPluginId.clear();
    outPluginShortId.clear();
    outModulePin.reset();

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(pluginId);
    if (! entry || entry->id.empty() || entry->disabled || ! entry->loadable || entry->path.empty())
    {
        return false;
    }

    wil::unique_hmodule module(::LoadLibraryExW(entry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    if (! module)
    {
        const DWORD error = ::GetLastError();
        Debug::Warning(L"FindFiles: LoadLibraryExW failed for '{}' (pluginId={} error={}).", entry->path.wstring(), entry->id, error);
        return false;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // FARPROC to typed factory function
    const auto createFactory = reinterpret_cast<CreateFactoryFunc>(::GetProcAddress(module.get(), "RedSalamanderCreate"));
#pragma warning(pop)
    if (! createFactory)
    {
        Debug::Warning(L"FindFiles: Missing export RedSalamanderCreate in '{}' (pluginId={}).", entry->path.wstring(), entry->id);
        return false;
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr<IFileSystem> fileSystem;
    const std::wstring requestedPluginId = entry->factoryPluginId.empty() ? entry->id : entry->factoryPluginId;
    if (requestedPluginId.empty())
    {
        return false;
    }
    const HRESULT createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), requestedPluginId.c_str(), fileSystem.put_void());

    if (FAILED(createHr) || ! fileSystem)
    {
        Debug::Warning(L"FindFiles: RedSalamanderCreate failed for '{}' (pluginId={} hr=0x{:08X}).",
                       entry->path.wstring(),
                       requestedPluginId,
                       static_cast<unsigned long>(createHr));
        return false;
    }

    if (entry->informations)
    {
        wil::com_ptr<IInformations> informations;
        const HRESULT qiInfos = fileSystem->QueryInterface(__uuidof(IInformations), informations.put_void());
        if (FAILED(qiInfos) || ! informations)
        {
            Debug::Warning(L"FindFiles: IInformations not supported by '{}' (pluginId={} hr=0x{:08X}).",
                           entry->path.wstring(),
                           entry->id,
                           static_cast<unsigned long>(qiInfos));
            return false;
        }

        const char* configuration = nullptr;
        static_cast<void>(entry->informations->GetConfiguration(&configuration));
        if (configuration && configuration[0] != '\0')
        {
            static_cast<void>(informations->SetConfiguration(configuration));
        }
    }

    auto modulePin = std::shared_ptr<wil::unique_hmodule>(new (std::nothrow) wil::unique_hmodule());
    if (! modulePin)
    {
        return false;
    }
    *modulePin = std::move(module);

    outFileSystem    = std::move(fileSystem);
    outPluginId      = entry->id;
    outPluginShortId = entry->shortId;
    outModulePin     = std::move(modulePin);
    return true;
}

[[nodiscard]] bool IsExplicitLocalSearchContext(const FindFilesPaneContext& context) noexcept
{
    if (! context.fileSystem)
    {
        return false;
    }

    return OrdinalString::EqualsNoCase(context.pluginId, kBuiltinLocalFileSystemId) || OrdinalString::EqualsNoCase(context.pluginShortId, L"file");
}

[[nodiscard]] bool TryResolveSearchContextForRoot(SearchRequest& request) noexcept
{
    if (! NavigationLocation::LooksLikeWindowsAbsolutePath(request.rootPath))
    {
        return static_cast<bool>(request.context.fileSystem);
    }

    if (IsExplicitLocalSearchContext(request.context))
    {
        return true;
    }

    wil::com_ptr<IFileSystem> localFileSystem;
    std::wstring pluginId;
    std::wstring pluginShortId;
    std::shared_ptr<wil::unique_hmodule> modulePin;
    if (! TryCreateIsolatedFileSystem(kBuiltinLocalFileSystemId, localFileSystem, pluginId, pluginShortId, modulePin))
    {
        return static_cast<bool>(request.context.fileSystem);
    }

    request.context.fileSystem    = std::move(localFileSystem);
    request.context.pluginId      = std::move(pluginId);
    request.context.pluginShortId = std::move(pluginShortId);
    request.context.instanceContext.clear();
    request.context.rootPluginPath = std::filesystem::path(request.rootPath);
    request.contextModulePin       = std::move(modulePin);
    return true;
}

[[nodiscard]] std::wstring ToLowerCopy(std::wstring_view value) noexcept
{
    return OrdinalString::FoldCaseInvariant(value);
}

[[nodiscard]] std::wstring MakeResultKey(std::wstring_view pluginId, std::wstring_view instanceContext, std::wstring_view fullPath) noexcept
{
    std::wstring key = ToLowerCopy(pluginId);
    key.push_back(L'|');
    key.append(ToLowerCopy(instanceContext));
    key.push_back(L'|');
    key.append(ToLowerCopy(fullPath));
    return key;
}

[[nodiscard]] std::wstring FormatSearchStatusHint(HRESULT hr) noexcept
{
    return FormatHResultMessage(hr);
}

void UpdateRecentValue(std::vector<std::wstring>& history, std::wstring value) noexcept
{
    if (value.empty())
    {
        return;
    }

    const auto it =
        std::find_if(history.begin(), history.end(), [&](const std::wstring& existing) noexcept { return OrdinalString::EqualsNoCase(existing, value); });
    if (it != history.end())
    {
        history.erase(it);
    }

    history.insert(history.begin(), std::move(value));
    if (history.size() > kMaxRecentEntries)
    {
        history.resize(kMaxRecentEntries);
    }
}

[[nodiscard]] std::wstring StripSingleLineControlCharacters(std::wstring_view text) noexcept
{
    std::wstring sanitized;
    sanitized.reserve(text.size());
    for (const wchar_t ch : text)
    {
        if (std::iswcntrl(static_cast<wint_t>(ch)) == 0)
        {
            sanitized.push_back(ch);
        }
    }
    return sanitized;
}

[[nodiscard]] bool OpenClipboardWithRetriesForFind(HWND ownerWindow) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + 250ms;
    while (std::chrono::steady_clock::now() < deadline)
    {
        if (OpenClipboard(ownerWindow) != 0)
        {
            return true;
        }

        if (GetOpenClipboardWindow() == nullptr)
        {
            break;
        }

        std::this_thread::sleep_for(10ms);
    }

    return OpenClipboard(ownerWindow) != 0;
}

[[nodiscard]] wil::unique_hglobal BuildUnicodeTextClipboardHGlobal(std::wstring_view text) noexcept
{
    const size_t bytes = (text.size() + 1u) * sizeof(wchar_t);
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! memory)
    {
        return nullptr;
    }

    auto* data = static_cast<wchar_t*>(GlobalLock(memory.get()));
    if (! data)
    {
        return nullptr;
    }

    std::copy(text.begin(), text.end(), data);
    data[text.size()] = L'\0';
    GlobalUnlock(memory.get());
    return memory;
}

[[nodiscard]] wil::unique_hglobal BuildFileDropClipboardHGlobal(const std::vector<std::filesystem::path>& paths) noexcept
{
    size_t totalChars = 1u;
    for (const auto& path : paths)
    {
        totalChars += path.native().size() + 1u;
    }

    const size_t bytes = sizeof(DROPFILES) + totalChars * sizeof(wchar_t);
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! memory)
    {
        return nullptr;
    }

    auto* drop = static_cast<DROPFILES*>(GlobalLock(memory.get()));
    if (! drop)
    {
        return nullptr;
    }

    drop->pFiles = sizeof(DROPFILES);
    drop->pt     = POINT{};
    drop->fNC    = FALSE;
    drop->fWide  = TRUE;

    auto* cursor = reinterpret_cast<wchar_t*>(reinterpret_cast<BYTE*>(drop) + drop->pFiles);
    for (const auto& path : paths)
    {
        const std::wstring text = path.native();
        std::copy(text.begin(), text.end(), cursor);
        cursor += text.size();
        *cursor++ = L'\0';
    }
    *cursor = L'\0';
    GlobalUnlock(memory.get());

    return memory;
}

[[nodiscard]] wil::unique_hglobal BuildPreferredDropEffectClipboardHGlobal(DWORD preferredEffect) noexcept
{
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD)));
    if (! memory)
    {
        return nullptr;
    }

    auto* effect = static_cast<DWORD*>(GlobalLock(memory.get()));
    if (! effect)
    {
        return nullptr;
    }

    *effect = preferredEffect;
    GlobalUnlock(memory.get());
    return memory;
}

[[nodiscard]] UINT PreferredDropEffectClipboardFormatForFind() noexcept
{
    static const UINT format = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
    return format;
}

[[nodiscard]] bool SetFindResultsClipboard(HWND ownerWindow,
                                           const std::vector<std::filesystem::path>& paths,
                                           DWORD preferredEffect,
                                           std::wstring_view text) noexcept
{
    if (paths.empty())
    {
        return false;
    }

    wil::unique_hglobal drop = BuildFileDropClipboardHGlobal(paths);
    if (! drop)
    {
        return false;
    }

    wil::unique_hglobal effect = BuildPreferredDropEffectClipboardHGlobal(preferredEffect);
    if (! effect)
    {
        return false;
    }

    wil::unique_hglobal textMemory;
    if (! text.empty())
    {
        textMemory = BuildUnicodeTextClipboardHGlobal(text);
        if (! textMemory)
        {
            return false;
        }
    }

    if (! OpenClipboardWithRetriesForFind(ownerWindow))
    {
        Debug::Warning(L"FindFiles: OpenClipboard failed for result file-drop copy (error={}).", GetLastError());
        return false;
    }
    const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });

    if (EmptyClipboard() == 0)
    {
        Debug::Warning(L"FindFiles: EmptyClipboard failed for result file-drop copy (error={}).", GetLastError());
        return false;
    }

    if (SetClipboardData(CF_HDROP, drop.get()) == nullptr)
    {
        Debug::Warning(L"FindFiles: SetClipboardData(CF_HDROP) failed (error={}).", GetLastError());
        return false;
    }
    drop.release();

    const UINT preferredDropEffectFormat = PreferredDropEffectClipboardFormatForFind();
    if (preferredDropEffectFormat == 0u || SetClipboardData(preferredDropEffectFormat, effect.get()) == nullptr)
    {
        Debug::Warning(L"FindFiles: SetClipboardData(Preferred DropEffect) failed (format={}, error={}).", preferredDropEffectFormat, GetLastError());
        return false;
    }
    effect.release();

    if (textMemory && SetClipboardData(CF_UNICODETEXT, textMemory.get()) != nullptr)
    {
        textMemory.release();
    }

    return true;
}

[[nodiscard]] bool IsSearchUnsupported(HRESULT hr) noexcept
{
    return hr == E_NOTIMPL || hr == HRESULT_FROM_WIN32(ERROR_CALL_NOT_IMPLEMENTED) || hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

[[nodiscard]] uint32_t ToComboIndex(Common::Settings::SearchNameMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::SearchNameMode::Wildcard: return 0u;
        case Common::Settings::SearchNameMode::Literal: return 1u;
        case Common::Settings::SearchNameMode::Regex: return 2u;
    }

    return 0u;
}

[[nodiscard]] Common::Settings::SearchNameMode FromNameModeComboIndex(int index) noexcept
{
    switch (index)
    {
        case 1: return Common::Settings::SearchNameMode::Literal;
        case 2: return Common::Settings::SearchNameMode::Regex;
        default: return Common::Settings::SearchNameMode::Wildcard;
    }
}

[[nodiscard]] uint32_t ToComboIndex(Common::Settings::SearchContentMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::SearchContentMode::Disabled: return 0u;
        case Common::Settings::SearchContentMode::TextLiteral: return 1u;
        case Common::Settings::SearchContentMode::TextRegex: return 2u;
    }

    return 0u;
}

[[nodiscard]] Common::Settings::SearchContentMode FromContentModeComboIndex(int index) noexcept
{
    switch (index)
    {
        case 1: return Common::Settings::SearchContentMode::TextLiteral;
        case 2: return Common::Settings::SearchContentMode::TextRegex;
        default: return Common::Settings::SearchContentMode::Disabled;
    }
}

[[nodiscard]] FileSystemSearchNameMode ToAbiNameMode(Common::Settings::SearchNameMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::SearchNameMode::Wildcard: return FILESYSTEM_SEARCH_NAME_WILDCARD;
        case Common::Settings::SearchNameMode::Literal: return FILESYSTEM_SEARCH_NAME_LITERAL;
        case Common::Settings::SearchNameMode::Regex: return FILESYSTEM_SEARCH_NAME_REGEX;
    }

    return FILESYSTEM_SEARCH_NAME_WILDCARD;
}

[[nodiscard]] FileSystemSearchContentMode ToAbiContentMode(Common::Settings::SearchContentMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::SearchContentMode::Disabled: return FILESYSTEM_SEARCH_CONTENT_DISABLED;
        case Common::Settings::SearchContentMode::TextLiteral: return FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;
        case Common::Settings::SearchContentMode::TextRegex: return FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX;
    }

    return FILESYSTEM_SEARCH_CONTENT_DISABLED;
}

[[nodiscard]] std::wstring FormatFileSize(int64_t sizeBytes) noexcept
{
    if (sizeBytes <= 0)
    {
        return {};
    }

    return std::format(L"{:L}", sizeBytes);
}

[[nodiscard]] std::wstring FormatFileTimeValue(int64_t fileTimeTicks) noexcept
{
    if (fileTimeTicks <= 0)
    {
        return {};
    }

    FILETIME ft{};
    ft.dwLowDateTime  = static_cast<DWORD>(static_cast<uint64_t>(fileTimeTicks) & 0xFFFFFFFFull);
    ft.dwHighDateTime = static_cast<DWORD>((static_cast<uint64_t>(fileTimeTicks) >> 32u) & 0xFFFFFFFFull);

    SYSTEMTIME utc{};
    SYSTEMTIME local{};
    if (FileTimeToSystemTime(&ft, &utc) == FALSE || SystemTimeToTzSpecificLocalTime(nullptr, &utc, &local) == FALSE)
    {
        return {};
    }

    return std::format(L"{:04}-{:02}-{:02} {:02}:{:02}", local.wYear, local.wMonth, local.wDay, local.wHour, local.wMinute);
}

[[nodiscard]] std::wstring FormatAttributes(unsigned long attributes) noexcept
{
    std::wstring value;
    if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        value.push_back(L'D');
    }
    if ((attributes & FILE_ATTRIBUTE_ARCHIVE) != 0)
    {
        value.push_back(L'A');
    }
    if ((attributes & FILE_ATTRIBUTE_READONLY) != 0)
    {
        value.push_back(L'R');
    }
    if ((attributes & FILE_ATTRIBUTE_HIDDEN) != 0)
    {
        value.push_back(L'H');
    }
    if ((attributes & FILE_ATTRIBUTE_SYSTEM) != 0)
    {
        value.push_back(L'S');
    }

    return value;
}

class FindFilesWindow;

class SearchSessionController
{
public:
    SearchSessionController()                                          = default;
    SearchSessionController(const SearchSessionController&)            = delete;
    SearchSessionController& operator=(const SearchSessionController&) = delete;

    ~SearchSessionController() noexcept;

    [[nodiscard]] bool Start(FindFilesWindow& owner, SearchRequest request, uint64_t epoch) noexcept;
    void Cancel() noexcept;
    void Shutdown() noexcept;
    void NotifyUiSettled() noexcept;
    [[nodiscard]] bool IsActive() const noexcept;
#ifdef ENABLE_TESTS
    [[nodiscard]] bool IsUiSettled() const noexcept;
    [[nodiscard]] bool WaitForIdle(uint32_t timeoutMs) noexcept;
#endif

private:
    void JoinCompletedWorker() noexcept;
    void Run(SearchRequest request, uint64_t epoch) noexcept;
    void MarkIdle() noexcept;

    std::jthread _worker;
    std::atomic<bool> _cancelRequested{false};
    std::atomic<bool> _active{false};
    std::atomic<bool> _uiSettled{true};
    HWND _ownerHwnd         = nullptr;
    FindFilesWindow* _owner = nullptr;
    mutable std::mutex _mutex;
    std::condition_variable _idleCv;
};

class FindFilesWindow final : public IDxGridDelegate
{
public:
    using IDxGridDelegate::OnGridContextMenu;
    using IDxGridDelegate::OnGridRowActivated;
    using IDxGridDelegate::OnGridSelectionChanged;

    FindFilesWindow(
        HWND owner, FolderWindow& applicationFolderWindow, Common::Settings::Settings& settings, AppTheme theme, FindFilesPaneContext context) noexcept;

    ~FindFilesWindow() noexcept override
    {
        // Lets Create() observe a self-delete performed by the window
        // procedure while CreateWindowExW was still on the stack.
        if (_destructionObserver)
        {
            *_destructionObserver = true;
        }
    }

    FindFilesWindow(const FindFilesWindow&)            = delete;
    FindFilesWindow& operator=(const FindFilesWindow&) = delete;

    [[nodiscard]] bool Create() noexcept;
    void UpdateTheme(const AppTheme& theme) noexcept;
    void UpdateOwnerWindow(HWND owner) noexcept;
    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hWnd.get();
    }
    [[nodiscard]] bool IsSearchActive() const noexcept
    {
        return _session.IsActive();
    }
#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugConfigure(std::wstring rootPath,
                                      std::wstring namePattern,
                                      std::wstring contentPattern,
                                      Common::Settings::SearchNameMode nameMode,
                                      Common::Settings::SearchContentMode contentMode) noexcept;
    [[nodiscard]] bool DebugSetOptions(bool recursive, bool includeFiles, bool includeDirectories, bool preferIndex, bool wantSnippets) noexcept;
    [[nodiscard]] bool DebugSetComboText(FindFilesDebugFocusTarget target, std::wstring text) noexcept;
    [[nodiscard]] bool DebugSetDestinationPath(std::wstring path) noexcept;
    [[nodiscard]] bool DebugStartSearch(FindFilesDebugOperation operation) noexcept;
    [[nodiscard]] bool DebugCancelSearch() noexcept;
    [[nodiscard]] bool DebugGetSnapshot(FindFilesDebugSnapshot& out) noexcept;
    [[nodiscard]] bool DebugHitTestResultsGrid(D2D1_POINT_2F pointDip, FindFilesDebugGridHit& out) const noexcept;
    [[nodiscard]] bool DebugFocusTarget(FindFilesDebugFocusTarget target) noexcept;
    [[nodiscard]] bool DebugGetTargetClientRect(FindFilesDebugFocusTarget target, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugApplyResultsLayoutFromSettings() noexcept;
    [[nodiscard]] bool DebugSetResultSort(size_t columnIndex, bool descending) noexcept;
    [[nodiscard]] bool DebugReorderVisibleResultColumn(size_t fromVisibleIndex, size_t targetVisibleIndex) noexcept;
    [[nodiscard]] bool DebugResizeVisibleResultColumn(size_t visibleIndex, float deltaDip) noexcept;
    [[nodiscard]] bool DebugSelectResult(std::wstring fullPath) noexcept;
    [[nodiscard]] bool DebugSelectResults(std::vector<std::wstring> fullPaths) noexcept;
    [[nodiscard]] bool DebugActivateSelectedResult() noexcept;
    [[nodiscard]] bool DebugOpenSelectedResultParent() noexcept;
    [[nodiscard]] bool DebugGetSelectedOpenDisposition(bool parentOnly, FindFilesDebugOpenDisposition& out) const noexcept;
    [[nodiscard]] bool DebugScrollResultsByWheelDetents(int detents) noexcept;
    [[nodiscard]] bool DebugWaitForIdle(uint32_t timeoutMs) noexcept;
    [[nodiscard]] bool DebugPostStaleSearchPayloads(std::wstring fullPath) noexcept;
#endif

    LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;

private:
    void TraceRawWindowMessage(UINT message, WPARAM wParam, LPARAM lParam, std::wstring_view phase, bool dxHandled = false) const noexcept;
    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept;
    [[nodiscard]] bool OnClose() noexcept;
    LRESULT OnNcDestroy() noexcept;
    void BuildUi() noexcept;
    [[nodiscard]] bool CreateRootNavigation(HWND parent) noexcept;
    [[nodiscard]] bool CreateDestinationNavigation(HWND parent) noexcept;
    void Layout() noexcept;
    void ApplyTheme() noexcept;
    void ApplyResultsGridMetrics() noexcept;
    void PopulateFromSettings() noexcept;
    void PopulateModeCombos() noexcept;
    void PopulateHistoryCombos() noexcept;
    void ApplyResultsSortFromSettings() noexcept;
    void ApplyResultsGridLayoutFromSettings() noexcept;
    void UpdateOptionDependencies() noexcept;
    void UpdateActionButtons() noexcept;
    void ShowFindActionMenu(POINT screenPoint) noexcept;
    [[nodiscard]] std::wstring ResolveResultMenuShortcutText(std::wstring_view commandId) const noexcept;
    [[nodiscard]] MenuFlyoutItem BuildResultMenuItem(FindResultMenuTarget target, FindResultMenuAction action, bool enabled) const;
    void AppendResultMenuActions(std::vector<MenuFlyoutItem>& items, FindResultMenuTarget target, bool includeOpenActions, bool enabled) const;
    void ShowResultContextMenu(size_t clickedRowIndex, POINT screenPoint) noexcept;
    [[nodiscard]] bool DispatchResultContextMenuCommand(const FindResultMenuCommand& command, size_t clickedRowIndex) noexcept;
    void ShowResultActionsHelp() noexcept;
    void PersistUiState(bool updateHistory) noexcept;
    [[nodiscard]] bool IsIndexedPreferenceAvailableForCurrentRoot() const noexcept;
    [[nodiscard]] bool CanHandleResultCommands() const noexcept;
    [[nodiscard]] std::optional<std::wstring> ResolveResultShortcutCommand(UINT message, WPARAM wParam) const noexcept;
    [[nodiscard]] bool HandleResultShortcut(UINT message, WPARAM wParam) noexcept;
    [[nodiscard]] bool HandleResultCommandId(unsigned int commandId) noexcept;
    [[nodiscard]] bool HandleResultCommand(std::wstring_view commandId) noexcept;
    [[nodiscard]] bool FocusRootNavigation(bool editMode) noexcept;
    [[nodiscard]] bool IsRootNavigationFocused() const noexcept;
    [[nodiscard]] bool HandleRootNavigationTabBridge(UINT message, WPARAM wParam) noexcept;
    [[nodiscard]] bool HandleRootMnemonic(UINT message, WPARAM wParam) noexcept;
    void UpdateKeyboardModifierState(UINT message, WPARAM wParam) noexcept;
    [[nodiscard]] uint32_t GetEffectiveKeyboardModifiers() const noexcept;
    [[nodiscard]] std::vector<size_t> CollectSelectedResultIndices() const;
    struct SelectedResultsFileOperationContext final
    {
        std::vector<std::filesystem::path> paths;
        std::vector<std::wstring> resultKeys;
        std::wstring pluginId;
        std::wstring instanceContext;
    };
    [[nodiscard]] std::optional<SelectedResultsFileOperationContext> CollectSelectedResultsFileOperationContext(
        std::wstring_view operationLabel) const noexcept;
    [[nodiscard]] bool CopySelectedResultsToClipboard(DWORD preferredEffect) noexcept;
    [[nodiscard]] bool LaunchSelectedResultFileAction(unsigned int commandId) noexcept;
    [[nodiscard]] bool CopyOrMoveSelectedResultsToOtherPane(FileSystemOperation operation) noexcept;
    [[nodiscard]] bool DeleteSelectedResults(bool permanent) noexcept;
    struct SearchTextOverride final
    {
        std::wstring rootPath;
        std::wstring namePattern;
        std::wstring contentPattern;
    };
#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugStartSearchWithTextOverride(FindFilesDebugOperation operation, SearchTextOverride textOverride) noexcept;
    void DebugRecordIncrementalResultRefresh() noexcept;
#endif
    [[nodiscard]] bool BeginSearch(SearchOperation operation, const SearchTextOverride* textOverride = nullptr) noexcept;
    void OnSearchStarted(SearchOperation operation, const SearchRequest& request) noexcept;
    void OnSearchResults(WPARAM operationKey, LPARAM lParam) noexcept;
    void OnSearchProgress(WPARAM operationKey, LPARAM lParam) noexcept;
    void OnSearchComplete(std::unique_ptr<FindSearchCompletePayload> payload) noexcept;
    void CompleteDeferredCloseIfReady() noexcept;
    void ApplyDeferredSetOperation(SearchOperation operation) noexcept;
    void ClearResults() noexcept;
    void RebuildResultsList() noexcept;
    [[nodiscard]] ResultListMutation AddOrUpdateVisibleResult(FindResultRecord result) noexcept;
    void RemoveKeysFromResults(const std::unordered_set<std::wstring>& keys) noexcept;
    void KeepOnlyKeysInResults(const std::unordered_set<std::wstring>& keys) noexcept;
    [[nodiscard]] std::optional<size_t> FindResultColumnIndexById(std::wstring_view columnId) const noexcept;
    [[nodiscard]] std::optional<size_t> GetSelectedResultIndex() const noexcept;
    enum class ResultOpenDisposition : uint8_t
    {
        None,
        NavigateToResult,
        NavigateToParent,
        NavigateToParentAndOpen,
        DefaultOpenFile,
    };
    struct ResultOpenPlan final
    {
        ResultOpenDisposition disposition = ResultOpenDisposition::None;
        std::filesystem::path targetFolder;
        std::wstring focusName;
        unsigned int commandId = 0u;
    };
    [[nodiscard]] static ResultOpenPlan BuildOpenPlan(const FindResultRecord& record, bool parentOnly) noexcept;
    [[nodiscard]] bool OpenFileWithDefaultApplication(const FindResultRecord& record) noexcept;
    void OpenSelectedResult(bool parentOnly) noexcept;
    void SetStatusText(std::wstring text) noexcept;
    void RefreshRootNavigationPath() noexcept;
    void RefreshRootNavigationHistory() noexcept;
    void OnRootNavigationPathChanged(const std::optional<std::filesystem::path>& path) noexcept;
    void RefreshDestinationStatusText() noexcept;
    void RefreshDestinationNavigationPath() noexcept;
    void RefreshDestinationNavigationHistory(const std::optional<std::filesystem::path>& destination) noexcept;
    void OnDestinationNavigationPathChanged(const std::optional<std::filesystem::path>& path) noexcept;
    [[nodiscard]] std::optional<std::filesystem::path> ResolveDestinationFolderForDisplay() const noexcept;
    [[nodiscard]] std::wstring BuildDestinationStatusText() const;
    void RefreshStatusText() noexcept;
    void UpdateStatusRefreshTimer() noexcept;
    void OnStatusRefreshTimer() noexcept;
    [[nodiscard]] std::wstring BuildStatusText() const noexcept;
    [[nodiscard]] std::wstring BuildRunningStatusText() const noexcept;
    [[nodiscard]] std::wstring BuildBackendStatusText() const noexcept;
    [[nodiscard]] std::wstring BuildWarningSummary(uint32_t warningFlags) const noexcept;
    [[nodiscard]] UINT BackendStringId(FileSystemSearchBackend backend) const noexcept;
    [[nodiscard]] UINT PhaseStringId(FileSystemSearchPhase phase) const noexcept;
    [[nodiscard]] HWND ResolveRestoreFolderViewWindow() const noexcept;
    [[nodiscard]] std::optional<SearchRequest> BuildSearchRequest(const SearchTextOverride* textOverride = nullptr) noexcept;
    [[nodiscard]] std::wstring GetComboText(const ComboBox* combo) const noexcept;
    void RebuildResultIndexByKey() noexcept;
    void SortResults() noexcept;
    void RefreshResultsView(bool fullRebuild) noexcept;
    void RestorePendingSelectionIfAvailable() noexcept;
    void ScheduleResultsRefresh(bool fullRebuild, size_t changedCount) noexcept;
    void ApplyPendingResultsRefresh() noexcept;
#ifdef ENABLE_TESTS
    [[nodiscard]] FindFilesDebugFocusTarget ResolveDebugFocusTarget() const noexcept;
    [[nodiscard]] static bool IsComboPopupOpen(const ComboBox* combo) noexcept;
#endif

    void OnGridSortRequested(const GridSortSpec& sortSpec) override;
    void OnGridSelectionChanged() override;
    void OnGridRowActivated(size_t rowIndex) override;
    void OnGridContextMenu(Grid& sender, size_t rowIndex, POINT screenPoint) override;
    [[nodiscard]] wil::com_ptr<ID2D1Bitmap1> GetGridIconBitmap(const Grid& sender, int iconIndex, float targetDipSize, ID2D1DeviceContext* d2dContext) override;

    HWND _ownerWindow                     = nullptr;
    FolderWindow* _applicationFolderWindow = nullptr;
    HWND _restoreFocusWindow              = nullptr;
    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    FindFilesPaneContext _context;
    size_t _dispatchDepth      = 0u;
    bool _deletePending        = false;
    bool* _destructionObserver = nullptr;

    wil::unique_hwnd _hWnd;
    WindowHost _dxHost;
    std::unique_ptr<Panel> _rootStorage;
    Panel* _root                       = nullptr;
    Label* _rootLabel                  = nullptr;
    ComboBox* _rootCombo               = nullptr;
    Label* _nameLabel                  = nullptr;
    ComboBox* _nameCombo               = nullptr;
    ComboBox* _nameModeCombo           = nullptr;
    Label* _contentLabel               = nullptr;
    ComboBox* _contentCombo            = nullptr;
    ComboBox* _contentModeCombo        = nullptr;
    Checkbox* _recursiveCheck          = nullptr;
    Checkbox* _includeFilesCheck       = nullptr;
    Checkbox* _includeDirectoriesCheck = nullptr;
    Checkbox* _followSymlinksCheck     = nullptr;
    Checkbox* _matchCaseNameCheck      = nullptr;
    Checkbox* _matchCaseContentCheck   = nullptr;
    Checkbox* _preferIndexCheck        = nullptr;
    Checkbox* _wantSnippetsCheck       = nullptr;
    Button* _findButton                = nullptr;
    Button* _appendButton              = nullptr;
    Button* _intersectButton           = nullptr;
    Button* _subtractButton            = nullptr;
    Button* _cancelButton              = nullptr;
    Button* _openButton                = nullptr;
    Button* _parentButton              = nullptr;
    Button* _helpButton                = nullptr;
    StatusStrip* _statusText           = nullptr;
    Label* _destinationLabel           = nullptr;
    Grid* _resultsList                 = nullptr;
    NavigationView _rootNavigation;
    NavigationView _destinationNavigation;
    FindResultsGridModel _resultsModel{_theme};
    GridSortSpec _resultSortSpec{};
    uint64_t _nextResultOrdinal = 1u;
    std::optional<POINT> _pendingFindActionMenuPoint;
    std::optional<std::vector<size_t>> _resultCommandIndexOverride;
    uint32_t _keyboardModifiers = 0u;
#ifdef ENABLE_TESTS
    uint64_t _debugResultActionFocusRestoreRequestCount = 0u;
#endif

    SearchSessionController _session;
    SearchOperation _activeOperation      = SearchOperation::Find;
    FileSystemSearchBackend _lastBackend  = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
    FileSystemSearchPhase _lastPhase      = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    uint32_t _lastWarningFlags            = FILESYSTEM_SEARCH_WARNING_NONE;
    uint64_t _lastScannedDirectories      = 0;
    uint64_t _lastScannedFiles            = 0;
    uint64_t _lastCandidateFiles          = 0;
    uint64_t _lastMatchedEntries          = 0;
    HRESULT _lastStatusHint               = S_OK;
    bool _cancelRequestedUi               = false;
    bool _closeRequested                  = false;
    uint64_t _activeSearchEpoch           = 0u;
    uint64_t _searchStartedTickMs         = 0u;
    uint64_t _lastProgressTickMs          = 0u;
    uint64_t _lastBackendStatusTickMs     = 0u;
    uint64_t _lastBackendStatusPollTickMs = 0u;
    std::wstring _status;
    std::wstring _destinationStatus;
    std::optional<std::filesystem::path> _explicitDestinationFolder;

    // Result rows disappear only for source indices with a known S_OK terminal outcome. Failed,
    // skipped, cancelled, and unreported indices remain visible and truthful.
    struct PendingResultRemoval final
    {
        uint64_t taskId               = 0;
        FileSystemOperation operation = FILESYSTEM_COPY;
        uint64_t createdTickMs        = 0;
        std::vector<std::filesystem::path> sourcePaths;
        std::vector<std::wstring> resultKeys;
    };
    std::vector<PendingResultRemoval> _pendingResultRemovals;
    uint64_t _fileOperationCompletedCallbackToken                 = 0;
    std::shared_ptr<void> _fileOperationCompletedCallbackLifetime = std::make_shared<int>(0);

    void EnsureFileOperationCompletedSubscription() noexcept;
    void ReapExpiredPendingResultRemovals(uint64_t nowTickMs) noexcept;
    void OnFolderWindowFileOperationCompleted(const FolderWindow::FileOperationCompletedEvent& e) noexcept;
    std::wstring _lastCurrentPath;
    std::wstring _lastSearchRootPath;
    std::wstring _lastSubmittedRootPath;
    std::wstring _lastSubmittedNamePattern;
    std::wstring _lastSubmittedContentPattern;
    std::wstring _lastBeginRootPath;
    std::wstring _lastBeginNamePattern;
    std::wstring _lastBeginContentPattern;
    std::wstring _lastDebugStartRootPath;
    std::wstring _lastDebugStartNamePattern;
    std::wstring _lastDebugStartContentPattern;
    std::wstring _lastBuiltRootPath;
    std::wstring _lastBuiltNamePattern;
    std::wstring _lastBuiltContentPattern;
    SearchServiceBroker::ServiceStatus _lastServiceStatus;
    HRESULT _lastServiceStatusHr = E_FAIL;
    bool _hasServiceStatus       = false;
    std::vector<FindResultRecord> _results;
    std::unordered_map<std::wstring, size_t> _resultIndexByKey;
    std::unordered_set<std::wstring> _deferredKeys;
    std::wstring _pendingSelectionFullPath;
    bool _resultsRefreshPending               = false;
    bool _resultsRefreshTimerArmed            = false;
    bool _resultsRefreshFullRebuild           = false;
    size_t _resultsRefreshBatchCount          = 0u;
    size_t _resultsRefreshRecordCount         = 0u;
    uint64_t _resultsRefreshFirstQueuedTickMs = 0u;
#ifdef ENABLE_TESTS
    uint32_t _debugResultListFullRebuildCount           = 0;
    uint32_t _debugIncrementalResultRefreshCount        = 0;
    uint32_t _debugIncrementalVisibleResultRefreshCount = 0;
    float _debugResizeBeforeWidthDip                    = 0.0f;
    float _debugResizeTargetWidthDip                    = 0.0f;
    float _debugResizeObservedWidthDip                  = 0.0f;
    bool _debugResizeSucceeded                          = false;
    FindFilesDebugFocusTarget _debugLastSetComboTarget  = FindFilesDebugFocusTarget::None;
    std::wstring _debugLastSetComboRequestedText;
    std::wstring _debugLastSetComboObservedText;
    std::optional<std::wstring> _debugRootTextOverride;
    std::optional<std::wstring> _debugNameTextOverride;
    std::optional<std::wstring> _debugContentTextOverride;
#endif
};

FindFilesWindow* g_findFilesWindow = nullptr;
std::vector<HWND> g_findFilesWindows;

[[nodiscard]] std::wstring CopySizedUtf16(const wchar_t* value, unsigned long sizeBytes) noexcept
{
    if (! value || sizeBytes == 0)
    {
        return {};
    }

    size_t charCount = sizeBytes / sizeof(wchar_t);
    while (charCount > 0 && value[charCount - 1u] == L'\0')
    {
        --charCount;
    }

    return std::wstring(value, value + charCount);
}

struct SearchCallbacks final : IFileSystemSearchCallback
{
    explicit SearchCallbacks(HWND hwnd, const SearchRequest& request, std::atomic<bool>& cancelRequested, uint64_t epoch) noexcept
        : _hwnd(hwnd),
          _request(request),
          _cancelRequested(cancelRequested),
          _epoch(epoch)
    {
        _hostExtensions.sizeBytes             = sizeof(_hostExtensions);
        _hostExtensions.version               = FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1;
        _hostExtensions.callbackCookie        = nullptr;
        _hostExtensions.serviceStatusCallback = &SearchCallbacks::FileSystemSearchServiceStatus;
        _hostExtensions.serviceStatusCookie   = this;
    }

    SearchCallbacks(const SearchCallbacks&)            = delete;
    SearchCallbacks& operator=(const SearchCallbacks&) = delete;
    SearchCallbacks(SearchCallbacks&&)                 = delete;
    SearchCallbacks& operator=(SearchCallbacks&&)      = delete;

    [[nodiscard]] const FileSystemSearchHostExtensions* GetHostExtensions() const noexcept
    {
        return &_hostExtensions;
    }

    void ApplyServiceStatusInference(bool readyForQueryCutover, bool startupWarmupRunning) noexcept
    {
        if (_latestServiceStatus.queryExecutionMode != LocalSearchIndexCore::QueryExecutionMode::Unknown)
        {
            return;
        }

        if (_latestServiceStatus.fallbackReason != LocalSearchIndexCore::FallbackReason::None)
        {
            _latestServiceStatus.queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback;
            return;
        }

        if (startupWarmupRunning || (_latestServiceStatus.totalRoots != 0u && _latestServiceStatus.completedRoots < _latestServiceStatus.totalRoots))
        {
            _latestServiceStatus.queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback;
            _latestServiceStatus.fallbackReason     = LocalSearchIndexCore::FallbackReason::WarmupRunning;
            return;
        }

        if (! readyForQueryCutover)
        {
            _latestServiceStatus.queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback;
            _latestServiceStatus.fallbackReason     = LocalSearchIndexCore::FallbackReason::CutoverBlocked;
        }
    }

    void SetServiceStatus(const SearchServiceBroker::ServiceStatus& status) noexcept
    {
        _latestServiceStatus.available          = true;
        _latestServiceStatus.hr                 = S_OK;
        _latestServiceStatus.storeState         = status.storeState;
        _latestServiceStatus.syncPhase          = status.syncPhase;
        _latestServiceStatus.queryExecutionMode = status.queryExecutionMode;
        _latestServiceStatus.fallbackReason     = status.fallbackReason;
        _latestServiceStatus.completedRoots =
            status.completedRoots != 0u || status.totalRoots != 0u ? status.completedRoots : status.startupWarmupCompletedRoots;
        _latestServiceStatus.totalRoots = status.completedRoots != 0u || status.totalRoots != 0u ? status.totalRoots : status.startupWarmupTotalRoots;
        _latestServiceStatus.activeRoot = ! status.activeRoot.empty() ? status.activeRoot : status.startupWarmupCurrentRoot;
        ApplyServiceStatusInference(status.readyForQueryCutover, status.startupWarmupRunning);
        EmitPerfCount(L"find.service_status.snapshot_count");
    }

    void SetServiceStatusUnavailable(HRESULT hr) noexcept
    {
        _latestServiceStatus.available          = false;
        _latestServiceStatus.hr                 = hr;
        _latestServiceStatus.storeState         = LocalSearchIndexCore::StoreState::Unknown;
        _latestServiceStatus.syncPhase          = LocalSearchIndexCore::SyncPhase::Idle;
        _latestServiceStatus.queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::Unknown;
        _latestServiceStatus.fallbackReason     = LocalSearchIndexCore::FallbackReason::None;
        _latestServiceStatus.completedRoots     = 0u;
        _latestServiceStatus.totalRoots         = 0u;
        _latestServiceStatus.activeRoot.clear();
        EmitPerfCount(L"find.service_status.snapshot_count");
    }

    static HRESULT STDMETHODCALLTYPE FileSystemSearchServiceStatus(const ::FileSystemSearchServiceStatus* status, void* cookie) noexcept
    {
        if (! status || status->sizeBytes != sizeof(::FileSystemSearchServiceStatus) || ! cookie)
        {
            return E_INVALIDARG;
        }

        SearchCallbacks& self                        = *static_cast<SearchCallbacks*>(cookie);
        self._latestServiceStatus.available          = true;
        self._latestServiceStatus.hr                 = S_OK;
        self._latestServiceStatus.storeState         = static_cast<LocalSearchIndexCore::StoreState>(status->storeState);
        self._latestServiceStatus.syncPhase          = static_cast<LocalSearchIndexCore::SyncPhase>(status->syncPhase);
        self._latestServiceStatus.queryExecutionMode = static_cast<LocalSearchIndexCore::QueryExecutionMode>(status->queryExecutionMode);
        self._latestServiceStatus.fallbackReason     = static_cast<LocalSearchIndexCore::FallbackReason>(status->fallbackReason);
        self._latestServiceStatus.completedRoots     = status->completedRoots;
        self._latestServiceStatus.totalRoots         = status->totalRoots;
        self._latestServiceStatus.activeRoot         = CopySizedUtf16(status->activeRoot, status->activeRootSize);
        self.ApplyServiceStatusInference(true, false);
        EmitPerfCount(L"find.service_status.snapshot_count");
        return S_OK;
    }

    [[nodiscard]] bool FlushResults() noexcept
    {
        if (_batch.empty())
        {
            return true;
        }

        auto payload = std::unique_ptr<FindSearchResultsPayload>(new (std::nothrow) FindSearchResultsPayload{});
        if (! payload)
        {
            return false;
        }

        const auto now            = SteadyClock::now();
        const uint64_t batchSize  = static_cast<uint64_t>(_batch.size());
        const uint64_t batchAgeUs = _batchFirstQueuedAt != SteadyClock::time_point{}
                                        ? static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(now - _batchFirstQueuedAt).count())
                                        : 0u;
        EmitPerfCount(L"find.results.batch_size", batchSize);
        Debug::Perf::Emit(L"find.results.batch_age_us", L"", batchAgeUs, batchSize, 0u, S_OK);
        payload->results    = std::move(_batch);
        payload->epoch      = _epoch;
        payload->enqueuedAt = now;
        _batch.clear();
        _batchFirstQueuedAt = {};
        const WPARAM operationKey = static_cast<WPARAM>(payload->epoch);
        if (! PostMessagePayload(_hwnd, WndMsg::kFindSearchResults, operationKey, std::move(payload)))
        {
            return false;
        }

        EmitPerfCount(L"find.results.flush_count");
        return true;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchMatch(const ::FileSystemSearchMatch* match, void* /*cookie*/) noexcept override
    {
        if (! match || match->sizeBytes != sizeof(::FileSystemSearchMatch))
        {
            return E_INVALIDARG;
        }

        EmitPerfCount(L"find.results.match_callbacks");
        const Debug::Perf::Scope matchPerf(L"find.results.match_callback_ms");
        FindResultRecord record;
        record.pluginId                = _request.context.pluginId;
        record.pluginShortId           = _request.context.pluginShortId;
        record.instanceContext         = _request.context.instanceContext;
        record.fullPath                = CopySizedUtf16(match->fullPath, match->fullPathSize);
        record.relativePath            = CopySizedUtf16(match->relativePath, match->relativePathSize);
        record.displayPath             = BuildResultDisplayPath(record.relativePath, _request.rootPath, record.fullPath);
        record.displayName             = CopySizedUtf16(match->displayName, match->displayNameSize);
        record.previewText             = CopySizedUtf16(match->previewText, match->previewTextSize);
        record.fileAttributes          = match->fileAttributes;
        record.lastWriteTime           = match->lastWriteTime;
        record.endOfFile               = match->endOfFile;
        record.matchedBy               = match->matchedBy;
        record.key                     = MakeResultKey(record.pluginId, record.instanceContext, record.fullPath);
        record.folderViewRainbowHash32 = MakeFindResultFolderViewRainbowHash32(record);
        record.stableRowId             = MakeResultStableId(record);

        if (_batch.empty())
        {
            _batchFirstQueuedAt = SteadyClock::now();
        }
        _batch.push_back(std::move(record));
        const bool flushByCapacity = _batch.size() >= kBatchSize;
        if (flushByCapacity)
        {
            EmitPerfCount(L"find.results.flush_trigger.capacity");
        }

        if (flushByCapacity && ! FlushResults())
        {
            return E_FAIL;
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchProgress(const ::FileSystemSearchProgress* progress, void* /*cookie*/) noexcept override
    {
        if (! progress || progress->sizeBytes != sizeof(::FileSystemSearchProgress))
        {
            return E_INVALIDARG;
        }

        EmitPerfCount(L"find.progress.callback_count");
        const Debug::Perf::Scope progressPerf(L"find.progress.callback_ms");
        EmitPerfCount(L"find.results.flush_trigger.progress");
        if (! _batch.empty())
        {
            const bool flushByAge  = _batchFirstQueuedAt != SteadyClock::time_point{} && (SteadyClock::now() - _batchFirstQueuedAt) >= kProgressFlushMaxAge;
            const bool shouldFlush = _batch.size() >= kProgressFlushMinBatchSize || flushByAge;
            if (shouldFlush)
            {
                EmitPerfCount(L"find.results.flush_progress_executed");
                if (! FlushResults())
                {
                    return E_FAIL;
                }
            }
            else
            {
                EmitPerfCount(L"find.results.flush_progress_skipped");
                EmitPerfCount(L"find.results.flush_progress_skipped_batch_size", static_cast<uint64_t>(_batch.size()));
                const uint64_t pendingAgeUs =
                    _batchFirstQueuedAt != SteadyClock::time_point{}
                        ? static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SteadyClock::now() - _batchFirstQueuedAt).count())
                        : 0u;
                Debug::Perf::Emit(L"find.results.flush_progress_skipped_age_us", L"", pendingAgeUs, static_cast<uint64_t>(_batch.size()), 0u, S_OK);
            }
        }

        auto payload = std::unique_ptr<FindSearchProgressPayload>(new (std::nothrow) FindSearchProgressPayload{});
        if (! payload)
        {
            return E_OUTOFMEMORY;
        }

        payload->phase              = progress->phase;
        payload->epoch              = _epoch;
        payload->backend            = progress->backend;
        payload->warningFlags       = progress->warningFlags;
        payload->statusHint         = progress->statusHint;
        payload->scannedDirectories = progress->scannedDirectories;
        payload->scannedFiles       = progress->scannedFiles;
        payload->candidateFiles     = progress->candidateFiles;
        payload->matchedEntries     = progress->matchedEntries;
        payload->currentPath        = CopySizedUtf16(progress->currentPath, progress->currentPathSize);
        payload->serviceStatus      = _latestServiceStatus;
        payload->enqueuedAt         = SteadyClock::now();
        if (payload->backend == FILESYSTEM_SEARCH_BACKEND_SERVICE && payload->warningFlags == FILESYSTEM_SEARCH_WARNING_NONE &&
            payload->serviceStatus.available && payload->serviceStatus.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback)
        {
            payload->warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
        }
        _latestProgress = *payload;
        EmitPerfCount(L"find.progress.current_path_chars", static_cast<uint64_t>(payload->currentPath.size()));

        const WPARAM operationKey = static_cast<WPARAM>(payload->epoch);
        return PostMessagePayload(_hwnd, WndMsg::kFindSearchProgress, operationKey, std::move(payload)) ? S_OK : E_FAIL;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchShouldCancel(BOOL* pCancel, void* /*cookie*/) noexcept override
    {
        if (! pCancel)
        {
            return E_POINTER;
        }

        *pCancel = _cancelRequested.load(std::memory_order_acquire) ? TRUE : FALSE;
        return S_OK;
    }

    HWND _hwnd = nullptr;
    const SearchRequest& _request;
    std::atomic<bool>& _cancelRequested;
    uint64_t _epoch = 0u;
    FileSystemSearchHostExtensions _hostExtensions{};
    std::vector<FindResultRecord> _batch;
    SteadyClock::time_point _batchFirstQueuedAt{};
    SearchServiceStatusSnapshot _latestServiceStatus;
    FindSearchProgressPayload _latestProgress;
};

SearchSessionController::~SearchSessionController() noexcept
{
    Shutdown();
}

bool SearchSessionController::Start(FindFilesWindow& owner, SearchRequest request, uint64_t epoch) noexcept
{
    JoinCompletedWorker();

    bool expected = false;
    if (! _active.compare_exchange_strong(expected, true, std::memory_order_acq_rel))
    {
        return false;
    }

    _cancelRequested.store(false, std::memory_order_release);
    _uiSettled.store(false, std::memory_order_release);
    _ownerHwnd = owner.GetHwnd();
    _owner     = &owner;
    try
    {
        _worker = std::jthread([this, request = std::move(request), epoch]() mutable noexcept { Run(std::move(request), epoch); });
        EmitPerfCount(L"find.session.start_count");
    }
    catch (const std::system_error&)
    {
        // Thread creation is required for the modeless Find dialog; fall back to a clean start failure.
        _owner     = nullptr;
        _ownerHwnd = nullptr;
        _uiSettled.store(true, std::memory_order_release);
        MarkIdle();
        return false;
    }
    return true;
}

void SearchSessionController::JoinCompletedWorker() noexcept
{
    if (_worker.joinable() && ! _active.load(std::memory_order_acquire))
    {
        _worker.join();
    }
}

void SearchSessionController::Cancel() noexcept
{
    _cancelRequested.store(true, std::memory_order_release);
}

void SearchSessionController::Shutdown() noexcept
{
    Cancel();
    if (_worker.joinable())
    {
        _worker.join();
    }

    _owner     = nullptr;
    _ownerHwnd = nullptr;
    _uiSettled.store(true, std::memory_order_release);
    MarkIdle();
}

void SearchSessionController::NotifyUiSettled() noexcept
{
    _uiSettled.store(true, std::memory_order_release);
    _idleCv.notify_all();
}

bool SearchSessionController::IsActive() const noexcept
{
    return _active.load(std::memory_order_acquire);
}

#ifdef ENABLE_TESTS
bool SearchSessionController::WaitForIdle(uint32_t timeoutMs) noexcept
{
    std::unique_lock lock(_mutex);
    return _idleCv.wait_for(lock, std::chrono::milliseconds(timeoutMs), [this]() noexcept { return ! _active.load(std::memory_order_acquire); });
}

bool SearchSessionController::IsUiSettled() const noexcept
{
    return _uiSettled.load(std::memory_order_acquire);
}
#endif

void SearchSessionController::Run(SearchRequest request, uint64_t epoch) noexcept
{
    const Debug::Perf::Scope runPerf(L"find.session.run_total_ms");
    auto initialPayload = std::unique_ptr<FindSearchProgressPayload>(new (std::nothrow) FindSearchProgressPayload{});
    if (initialPayload)
    {
        initialPayload->epoch      = epoch;
        initialPayload->phase      = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
        initialPayload->backend    = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
        initialPayload->statusHint = S_OK;
        initialPayload->enqueuedAt = SteadyClock::now();
        const WPARAM operationKey  = static_cast<WPARAM>(initialPayload->epoch);
        static_cast<void>(PostMessagePayload(_ownerHwnd, WndMsg::kFindSearchProgress, operationKey, std::move(initialPayload)));
    }

    SearchCallbacks callbacks(_ownerHwnd, request, _cancelRequested, epoch);
#ifdef ENABLE_TESTS
    DebugMaybeBlockFindFilesWindowSearchRun();
#endif

    FileSystemSearchQuery query{};
    query.sizeBytes                     = sizeof(query);
    query.rootPath                      = request.rootPath.c_str();
    query.namePattern                   = request.nameMode == FILESYSTEM_SEARCH_NAME_DISABLED ? nullptr : request.namePattern.c_str();
    query.contentPattern                = request.contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED ? nullptr : request.contentPattern.c_str();
    query.flags                         = request.flags;
    query.nameMode                      = request.nameMode;
    query.contentMode                   = request.contentMode;
    query.maxResults                    = request.maxResults;
    query.maxContentBytesPerFile        = request.maxContentBytesPerFile;
    query.maxSnippetCharacters          = request.maxSnippetCharacters;
    query.reserved                      = 0;
    FileSystemSearchQuery fallbackQuery = query;

    HRESULT hr = E_NOINTERFACE;
    if (request.context.fileSystem)
    {
        void* searchCookie = nullptr;
        const bool isBuiltinLocalFileSystem =
            CompareStringOrdinal(
                request.context.pluginId.c_str(), -1, kBuiltinLocalFileSystemId.data(), static_cast<int>(kBuiltinLocalFileSystemId.size()), TRUE) == CSTR_EQUAL;
        if (isBuiltinLocalFileSystem)
        {
            query.reserved = FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1;
            searchCookie   = const_cast<FileSystemSearchHostExtensions*>(callbacks.GetHostExtensions());

            SearchServiceBroker::ServiceStatus status{};
            const HRESULT statusHr = SearchServiceBroker::GetStatus(status);
            if (SUCCEEDED(statusHr))
            {
                callbacks.SetServiceStatus(status);
            }
            else
            {
                callbacks.SetServiceStatusUnavailable(statusHr);
            }
        }

        wil::com_ptr<IFileSystemSearch> nativeSearch;
        if (SUCCEEDED(request.context.fileSystem->QueryInterface(IID_PPV_ARGS(nativeSearch.put()))) && nativeSearch)
        {
            EmitPerfCount(L"find.session.backend.native_search_calls");
            {
                const Debug::Perf::Scope nativeSearchPerf(L"find.session.backend.native_search_ms");
                hr = nativeSearch->Search(&query, &callbacks, searchCookie);
            }
            if (IsSearchUnsupported(hr) && ! _cancelRequested.load(std::memory_order_acquire))
            {
                EmitPerfCount(L"find.session.backend.unsupported_fallbacks");
                {
                    const Debug::Perf::Scope fallbackPerf(L"find.session.backend.fallback_engine_ms");
                    hr = SearchFallbackEngine::Execute(request.context.fileSystem.get(), &fallbackQuery, &callbacks, nullptr);
                }
            }
        }
        else
        {
            {
                const Debug::Perf::Scope fallbackPerf(L"find.session.backend.fallback_engine_ms");
                hr = SearchFallbackEngine::Execute(request.context.fileSystem.get(), &fallbackQuery, &callbacks, nullptr);
            }
        }
    }

    if (! callbacks._batch.empty())
    {
        EmitPerfCount(L"find.results.flush_trigger.complete");
    }
    static_cast<void>(callbacks.FlushResults());

    auto complete = std::unique_ptr<FindSearchCompletePayload>(new (std::nothrow) FindSearchCompletePayload{});
    if (complete)
    {
        complete->hr                 = hr;
        complete->epoch              = epoch;
        complete->backend            = callbacks._latestProgress.backend;
        complete->warningFlags       = callbacks._latestProgress.warningFlags;
        complete->scannedDirectories = callbacks._latestProgress.scannedDirectories;
        complete->scannedFiles       = callbacks._latestProgress.scannedFiles;
        complete->candidateFiles     = callbacks._latestProgress.candidateFiles;
        complete->matchedEntries     = callbacks._latestProgress.matchedEntries;
    }

    MarkIdle();

    bool completionQueued = false;
    if (complete)
    {
        const WPARAM operationKey = static_cast<WPARAM>(complete->epoch);
        completionQueued = PostMessagePayload(_ownerHwnd, WndMsg::kFindSearchComplete, operationKey, std::move(complete));
    }
    if (! completionQueued)
    {
        NotifyUiSettled();
        // The payload either could not be allocated or could not be posted. Post a bare (payload-less)
        // completion so the UI thread still runs OnSearchComplete -> CompleteDeferredCloseIfReady and a
        // deferred window close does not leak the hidden window. Best-effort: if even this post fails the
        // window survives until app teardown, but no UI-thread state is corrupted.
        static_cast<void>(PostMessagePayload(
            _ownerHwnd, WndMsg::kFindSearchComplete, static_cast<WPARAM>(epoch), std::unique_ptr<FindSearchCompletePayload>{}));
    }
}

void SearchSessionController::MarkIdle() noexcept
{
    {
        std::lock_guard lock(_mutex);
        _active.store(false, std::memory_order_release);
    }
    _idleCv.notify_all();
}

FindFilesWindow::FindFilesWindow(
    HWND owner, FolderWindow& applicationFolderWindow, Common::Settings::Settings& settings, AppTheme theme, FindFilesPaneContext context) noexcept
    : _applicationFolderWindow(&applicationFolderWindow),
      _settings(&settings),
      _theme(std::move(theme)),
      _context(std::move(context))
{
    UpdateOwnerWindow(owner);
}

bool FindFilesWindow::Create() noexcept
{
    // Create() owns failure cleanup: on any failure path the instance is
    // deleted exactly once, either here or by the window procedure tearing
    // down a half-created window. Callers must not delete after Create().
    HINSTANCE instance = GetModuleHandleW(nullptr);

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = &FindFilesWindow::WndProc;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon         = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
    wc.hIconSm       = LoadIconW(instance, MAKEINTRESOURCEW(IDI_SMALL));
    wc.lpszClassName = kFindFilesWindowClassName;
    if (RegisterClassExW(&wc) == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
    {
        delete this;
        return false;
    }

    const std::wstring title   = LoadStringResource(nullptr, IDS_FIND_TITLE);
    bool destroyedDuringCreate = false;
    _destructionObserver       = &destroyedDuringCreate;
    const HWND created         = CreateWindowExW(0,
                                                 kFindFilesWindowClassName,
                                                 title.c_str(),
                                                 WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                                 CW_USEDEFAULT,
                                                 CW_USEDEFAULT,
                                                 UiMetrics::ScaleDip(96u, 1120),
                                                 UiMetrics::ScaleDip(96u, 760),
                                                 nullptr,
                                                 nullptr,
                                                 instance,
                                                 this);
    if (destroyedDuringCreate)
    {
        // WM_CREATE failed: CreateWindowExW destroyed the half-created window
        // and the window procedure already deleted this instance.
        return false;
    }
    _destructionObserver = nullptr;
    if (! created)
    {
        delete this;
        return false;
    }

    const bool hasPlacement = _settings && _settings->windows.contains(std::wstring(kFindFilesWindowId));
    const int showCmd       = hasPlacement ? WindowPlacementPersistence::Restore(*_settings, kFindFilesWindowId, created) : SW_SHOWNORMAL;
    ShowWindow(created, showCmd);
    SetForegroundWindow(created);
    return true;
}

void FindFilesWindow::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    ApplyTheme();
    Layout();
}

void FindFilesWindow::UpdateOwnerWindow(HWND owner) noexcept
{
    _ownerWindow        = (owner && IsWindow(owner) != FALSE) ? GetAncestor(owner, GA_ROOT) : nullptr;
    _restoreFocusWindow = ResolveRestoreFolderViewWindow();
}

bool FindFilesWindow::OnCreate(HWND hwnd) noexcept
{
    if (! _dxHost.Attach(hwnd))
    {
        Debug::Error(L"FindFiles: failed to attach DxUi host.");
        return false;
    }

    BuildUi();
    if (! CreateRootNavigation(hwnd))
    {
        Debug::Error(L"FindFiles: failed to create root navigation bar.");
        return false;
    }
    if (! CreateDestinationNavigation(hwnd))
    {
        Debug::Error(L"FindFiles: failed to create destination navigation bar.");
        return false;
    }
    PopulateModeCombos();
    PopulateHistoryCombos();
    PopulateFromSettings();
    ApplyTheme();
    UpdateOptionDependencies();
    ApplyResultsSortFromSettings();
    ApplyResultsGridLayoutFromSettings();
    UpdateActionButtons();
    SetStatusText({});
    Layout();
    return true;
}

void FindFilesWindow::BuildUi() noexcept
{
    if (_root != nullptr)
    {
        return;
    }

    _rootStorage = std::make_unique<Panel>();
    _root        = _rootStorage.get();

    _rootLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_FIND_LABEL_ROOT));
    _rootLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
    _rootLabel->SetMnemonic(L'L');

    _rootCombo = _root->AddChild<ComboBox>();
    _rootCombo->SetEditable(true);
    _rootCombo->SetVariant(ComboBoxVariant::Edit);
    _rootCombo->SetVisible(false);
    _rootCombo->SetOnTextChanged([this](std::wstring_view)
    {
        UpdateOptionDependencies();
        PersistUiState(false);
    });
    _rootCombo->SetOnSubmitted([this] { static_cast<void>(BeginSearch(SearchOperation::Find)); });
    _rootLabel->SetMnemonicTarget(_rootCombo);

    _nameLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_FIND_LABEL_NAME));
    _nameLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
    _nameLabel->SetMnemonic(L'N');

    _nameCombo = _root->AddChild<ComboBox>();
    _nameCombo->SetEditable(true);
    _nameCombo->SetVariant(ComboBoxVariant::Edit);
    _nameCombo->SetOnTextChanged([this](std::wstring_view)
    {
        PersistUiState(false);
        UpdateActionButtons();
    });
    _nameCombo->SetOnSubmitted([this] { static_cast<void>(BeginSearch(SearchOperation::Find)); });
    _nameLabel->SetMnemonicTarget(_nameCombo);

    _nameModeCombo = _root->AddChild<ComboBox>();
    _nameModeCombo->SetVariant(ComboBoxVariant::Window);
    _nameModeCombo->SetOnSelectionChanged([this](size_t)
    {
        PersistUiState(false);
        UpdateActionButtons();
    });

    _contentLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_FIND_LABEL_CONTENT));
    _contentLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
    _contentLabel->SetMnemonic(L'C');

    _contentCombo = _root->AddChild<ComboBox>();
    _contentCombo->SetEditable(true);
    _contentCombo->SetVariant(ComboBoxVariant::Edit);
    _contentCombo->SetOnTextChanged([this](std::wstring_view)
    {
        PersistUiState(false);
        UpdateActionButtons();
    });
    _contentCombo->SetOnSubmitted([this] { static_cast<void>(BeginSearch(SearchOperation::Find)); });
    _contentLabel->SetMnemonicTarget(_contentCombo);

    _contentModeCombo = _root->AddChild<ComboBox>();
    _contentModeCombo->SetVariant(ComboBoxVariant::Window);
    _contentModeCombo->SetOnSelectionChanged([this](size_t)
    {
        UpdateOptionDependencies();
        PersistUiState(false);
        UpdateActionButtons();
    });

    const auto bindToggle = [this](Checkbox*& target, UINT resourceId)
    {
        target = _root->AddChild<Checkbox>(LoadStringResource(nullptr, resourceId));
        target->SetOnToggled([this](bool)
        {
            UpdateOptionDependencies();
            PersistUiState(false);
            UpdateActionButtons();
        });
    };

    bindToggle(_recursiveCheck, IDS_FIND_OPTION_RECURSIVE);
    bindToggle(_includeFilesCheck, IDS_FIND_OPTION_INCLUDE_FILES);
    bindToggle(_includeDirectoriesCheck, IDS_FIND_OPTION_INCLUDE_DIRECTORIES);
    bindToggle(_followSymlinksCheck, IDS_FIND_OPTION_FOLLOW_SYMLINKS);
    bindToggle(_matchCaseNameCheck, IDS_FIND_OPTION_MATCH_CASE_NAME);
    bindToggle(_matchCaseContentCheck, IDS_FIND_OPTION_MATCH_CASE_CONTENT);
    bindToggle(_preferIndexCheck, IDS_FIND_OPTION_PREFER_INDEX);
    bindToggle(_wantSnippetsCheck, IDS_FIND_OPTION_WANT_SNIPPETS);

    const auto bindAction = [this](Button*& target, UINT resourceId, SearchOperation operation)
    {
        target = _root->AddChild<Button>(LoadStringResource(nullptr, resourceId));
        target->SetOnClick([this, operation] { static_cast<void>(BeginSearch(operation)); });
    };

    bindAction(_findButton, IDS_FIND_ACTION_FIND, SearchOperation::Find);
    _findButton->SetPrimary(true);
    _findButton->SetVariant(ButtonVariant::Split);
    _findButton->SetOnDropDownClick([this]
    {
        if (! _findButton)
        {
            return;
        }

        const D2D1_RECT_F bounds    = _findButton->GetBounds();
        const POINT screenPoint     = _dxHost.DipPointToScreenPoint(D2D1::Point2F(bounds.left, bounds.bottom));
        _pendingFindActionMenuPoint = screenPoint;
        if (_hWnd)
        {
            const BOOL posted = PostMessageW(_hWnd.get(), WndMsg::kFindShowActionMenu, 0, 0);
            TraceFindContextMenuDiagnostics(L"find.action-menu-dropdown-click",
                                            L"hwnd={:#x} postMessage={} lastError={} point=({}, {}) bounds=({:.1f}, {:.1f}, {:.1f}, {:.1f}) "
                                            L"capture={:#x} focus={:#x}",
                                            reinterpret_cast<uintptr_t>(_hWnd.get()),
                                            posted != FALSE ? L"true" : L"false",
                                            posted != FALSE ? 0ul : static_cast<unsigned long>(GetLastError()),
                                            screenPoint.x,
                                            screenPoint.y,
                                            bounds.left,
                                            bounds.top,
                                            bounds.right,
                                            bounds.bottom,
                                            reinterpret_cast<uintptr_t>(GetCapture()),
                                            reinterpret_cast<uintptr_t>(GetFocus()));
        }
        else
        {
            if (IsContextMenuDiagnosticsEnabled())
            {
                TraceContextMenuDiagnostics(L"find.action-menu-dropdown-click", L"hwnd=null postMessage=false");
            }
        }
    });
    _findButton->SetMnemonic(L'F');
    bindAction(_appendButton, IDS_FIND_ACTION_APPEND, SearchOperation::Append);
    _appendButton->SetVisible(false);
    _appendButton->SetMnemonic(L'A');
    bindAction(_intersectButton, IDS_FIND_ACTION_INTERSECT, SearchOperation::Intersect);
    _intersectButton->SetVisible(false);
    _intersectButton->SetMnemonic(L'I');
    bindAction(_subtractButton, IDS_FIND_ACTION_SUBTRACT, SearchOperation::Subtract);
    _subtractButton->SetVisible(false);
    _subtractButton->SetMnemonic(L'S');

    _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
    _cancelButton->SetOnClick([this]
    {
        _cancelRequestedUi = true;
        _session.Cancel();
        RefreshStatusText();
        UpdateActionButtons();
    });

    _openButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_FIND_ACTION_OPEN));
    _openButton->SetMnemonic(L'O');
    _openButton->SetOnClick([this] { OpenSelectedResult(false); });

    _parentButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_FIND_ACTION_PARENT));
    _parentButton->SetMnemonic(L'P');
    _parentButton->SetOnClick([this] { OpenSelectedResult(true); });

    _helpButton = _root->AddChild<Button>(LoadEmbeddedStringResource(nullptr, IDS_FIND_ACTION_HELP));
    _helpButton->SetTooltipText(LoadStringResource(nullptr, IDS_FIND_ACTION_HELP_TOOLTIP));
    _helpButton->SetAccessibleName(LoadStringResource(nullptr, IDS_FIND_ACTION_HELP_TOOLTIP));
    _helpButton->SetAccessibleHelpText(LoadStringResource(nullptr, IDS_FIND_RESULT_ACTIONS_HELP_TEXT));
    _helpButton->SetOnClick([this] { ShowResultActionsHelp(); });

    _statusText = _root->AddChild<StatusStrip>();
    _statusText->SetFontRole(RedSalamander::DxUi::FontRole::Small);
    _statusText->SetBlendWithWindowBackground(true);
    _statusText->SetSections({StatusStrip::Section{
        .text            = _status,
        .widthDip        = 0.0f,
        .alignment       = DWRITE_TEXT_ALIGNMENT_TRAILING,
        .leadingEllipsis = true,
    }});

    _destinationLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_FIND_LABEL_DESTINATION));
    _destinationLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

    _resultsList = _root->AddChild<Grid>();
    _resultsList->SetDelegate(this);
    _resultsList->SetModel(&_resultsModel);
    _resultsList->SetSelectionMode(GridSelectionMode::Extended);
    _resultsList->SetVisualMode(GridVisualMode::FolderView);
    ApplyResultsGridMetrics();

    _resultsModel.SetRows(&_results);
    _dxHost.SetRoot(std::move(_rootStorage));
    _dxHost.SetDefaultButton(_findButton);
    _dxHost.SetCancelButton(_cancelButton);
    _dxHost.SetOnTabBoundary([this](bool) noexcept { return FocusRootNavigation(false); });
    _dxHost.SetOnEscape([this]() noexcept
    {
        if (_session.IsActive() || ! _hWnd || _dxHost.GetFocusControl() == nullptr)
        {
            return false;
        }

        return PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0) != FALSE;
    });
}

bool FindFilesWindow::CreateRootNavigation(HWND parent) noexcept
{
    _rootNavigation.SetSettings(_settings);
    _rootNavigation.SetFileSystem(_context.fileSystem);
    _rootNavigation.SetTheme(_theme);
    _rootNavigation.SetEmbeddedDestinationMode(true);
    _rootNavigation.SetPaneFocused(false);
    _rootNavigation.SetPathChangedCallback([this](const std::optional<std::filesystem::path>& path) noexcept { OnRootNavigationPathChanged(path); });
    _rootNavigation.SetRequestFolderViewFocusCallback([this]
    {
        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
    });

    const HWND hwnd = _rootNavigation.Create(parent, 0, 0, 0, 0);
    if (! hwnd)
    {
        return false;
    }

    RefreshRootNavigationPath();
    RefreshRootNavigationHistory();
    return true;
}

bool FindFilesWindow::CreateDestinationNavigation(HWND parent) noexcept
{
    _destinationNavigation.SetSettings(_settings);
    _destinationNavigation.SetFileSystem(_context.fileSystem);
    _destinationNavigation.SetTheme(_theme);
    _destinationNavigation.SetEmbeddedDestinationMode(true);
    _destinationNavigation.SetPaneFocused(false);
    _destinationNavigation.SetPathChangedCallback([this](const std::optional<std::filesystem::path>& path) noexcept
    { OnDestinationNavigationPathChanged(path); });
    _destinationNavigation.SetRequestFolderViewFocusCallback([this]
    {
        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
    });

    const HWND hwnd = _destinationNavigation.Create(parent, 0, 0, 0, 0);
    if (! hwnd)
    {
        return false;
    }

    TraceFindContextMenuDiagnostics(
        L"find.destination-navigation.create", L"find={:#x} nav={:#x}", reinterpret_cast<uintptr_t>(_hWnd.get()), reinterpret_cast<uintptr_t>(hwnd));
    RefreshDestinationNavigationPath();
    return true;
}

void FindFilesWindow::PopulateModeCombos() noexcept
{
    if (_nameModeCombo)
    {
        _nameModeCombo->SetItems({ComboBox::Item{L"wildcard", LoadStringResource(nullptr, IDS_FIND_NAME_MODE_WILDCARD)},
                                  ComboBox::Item{L"literal", LoadStringResource(nullptr, IDS_FIND_NAME_MODE_LITERAL)},
                                  ComboBox::Item{L"regex", LoadStringResource(nullptr, IDS_FIND_NAME_MODE_REGEX)}});
    }

    if (_contentModeCombo)
    {
        _contentModeCombo->SetItems({ComboBox::Item{L"disabled", LoadStringResource(nullptr, IDS_FIND_CONTENT_MODE_DISABLED)},
                                     ComboBox::Item{L"literal", LoadStringResource(nullptr, IDS_FIND_CONTENT_MODE_LITERAL)},
                                     ComboBox::Item{L"regex", LoadStringResource(nullptr, IDS_FIND_CONTENT_MODE_REGEX)}});
    }
}

void FindFilesWindow::PopulateHistoryCombos() noexcept
{
    const auto settings    = _settings && _settings->search.has_value() ? _settings->search.value() : Common::Settings::SearchDialogSettings{};
    const auto loadHistory = [](ComboBox* combo, const std::vector<std::wstring>& values) noexcept
    {
        if (! combo)
        {
            return;
        }

        const std::wstring currentText = std::wstring(combo->GetText());
        std::vector<ComboBox::Item> items;
        items.reserve(values.size());
        for (const auto& value : values)
        {
            items.push_back(ComboBox::Item{value, value});
        }
        combo->SetItems(std::move(items));
        combo->SetText(currentText);
    };

    loadHistory(_rootCombo, settings.recentRoots);
    loadHistory(_nameCombo, settings.recentNamePatterns);
    loadHistory(_contentCombo, settings.recentContentPatterns);
    RefreshRootNavigationHistory();
}

void FindFilesWindow::PopulateFromSettings() noexcept
{
    const Common::Settings::SearchDialogSettings defaults{};
    const auto settings = _settings && _settings->search.has_value() ? _settings->search.value() : defaults;

    const std::wstring contextRoot = _context.rootPluginPath.native();
    const std::wstring initialRoot = ! contextRoot.empty() ? contextRoot : settings.lastRoot;
    if (_rootCombo)
    {
        _rootCombo->SetText(initialRoot);
    }
    RefreshRootNavigationPath();
    RefreshRootNavigationHistory();
    if (_nameCombo)
    {
        _nameCombo->SetText(settings.lastNamePattern);
    }
    if (_contentCombo)
    {
        _contentCombo->SetText(settings.lastContentPattern);
    }
    if (_nameModeCombo)
    {
        _nameModeCombo->SetSelectedIndex(static_cast<size_t>(ToComboIndex(settings.nameMode)));
    }
    if (_contentModeCombo)
    {
        _contentModeCombo->SetSelectedIndex(static_cast<size_t>(ToComboIndex(settings.contentMode)));
    }
    if (_recursiveCheck)
    {
        _recursiveCheck->SetChecked(settings.recursive);
    }
    if (_includeFilesCheck)
    {
        _includeFilesCheck->SetChecked(settings.includeFiles);
    }
    if (_includeDirectoriesCheck)
    {
        _includeDirectoriesCheck->SetChecked(settings.includeDirectories);
    }
    if (_followSymlinksCheck)
    {
        _followSymlinksCheck->SetChecked(settings.followSymlinks);
    }
    if (_matchCaseNameCheck)
    {
        _matchCaseNameCheck->SetChecked(settings.matchCaseName);
    }
    if (_matchCaseContentCheck)
    {
        _matchCaseContentCheck->SetChecked(settings.matchCaseContent);
    }
    if (_preferIndexCheck)
    {
        _preferIndexCheck->SetChecked(settings.preferIndex);
    }
    if (_wantSnippetsCheck)
    {
        _wantSnippetsCheck->SetChecked(settings.wantSnippets);
    }
}

void FindFilesWindow::ApplyTheme() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    _dxHost.SetTheme(MakeFolderContentDxPalette(_theme));
    _rootNavigation.SetTheme(_theme);
    _destinationNavigation.SetTheme(_theme);
    ApplyTitleBarTheme(_hWnd.get(), _theme, GetActiveWindow() == _hWnd.get());
    ApplyWindowBackdropTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool);
    _dxHost.Invalidate();
}

void FindFilesWindow::ApplyResultsGridMetrics() noexcept
{
    if (! _resultsList)
    {
        return;
    }

    const bool compact      = _theme.compactMode;
    const bool showSnippets = _resultsModel.ShowsSnippetColumn();
    _resultsList->SetHeaderHeightDip(compact ? 26.0f : 32.0f);
    _resultsList->SetRowHeightDip(showSnippets ? (compact ? 38.0f : 46.0f) : (compact ? 24.0f : 28.0f));
    _resultsList->SetLineClamp(showSnippets ? (compact ? 2u : 3u) : 1u);
}

void FindFilesWindow::ApplyResultsGridLayoutFromSettings() noexcept
{
    if (! _resultsList || ! _settings || ! _settings->search.has_value())
    {
        return;
    }

    const auto layout = ConvertColumnLayout(_settings->search->resultsGridLayout);
    if (! layout.empty())
    {
        _resultsModel.ApplyColumnLayoutDefaults(layout);
        _resultsList->ApplyColumnLayout(layout);
    }
}

bool FindFilesWindow::IsIndexedPreferenceAvailableForCurrentRoot() const noexcept
{
    const std::wstring root = StripSingleLineControlCharacters(GetComboText(_rootCombo));
    if (root.empty() || ! NavigationLocation::LooksLikeWindowsAbsolutePath(root))
    {
        return false;
    }

    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::SupportInfo support{};
    const HRESULT hr = repository.ProbePath(root, support);
    return SUCCEEDED(hr) && support.indexable;
}

void FindFilesWindow::UpdateKeyboardModifierState(UINT message, WPARAM wParam) noexcept
{
    const bool keyDown = message == WM_KEYDOWN || message == WM_SYSKEYDOWN;
    const bool keyUp   = message == WM_KEYUP || message == WM_SYSKEYUP;
    if (! keyDown && ! keyUp)
    {
        return;
    }

    const auto setModifier = [&](uint32_t flag) noexcept
    {
        if (keyDown)
        {
            _keyboardModifiers |= flag;
        }
        else
        {
            _keyboardModifiers &= ~flag;
        }
    };

    switch (wParam)
    {
        case VK_CONTROL:
        case VK_LCONTROL:
        case VK_RCONTROL: setModifier(kFindShortcutCtrl); break;
        case VK_SHIFT:
        case VK_LSHIFT:
        case VK_RSHIFT: setModifier(kFindShortcutShift); break;
        case VK_MENU:
        case VK_LMENU:
        case VK_RMENU: setModifier(kFindShortcutAlt); break;
        default: break;
    }
}

uint32_t FindFilesWindow::GetEffectiveKeyboardModifiers() const noexcept
{
    uint32_t modifiers = _keyboardModifiers;
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        modifiers |= kFindShortcutCtrl;
    }
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        modifiers |= kFindShortcutShift;
    }
    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        modifiers |= kFindShortcutAlt;
    }
    return modifiers;
}

std::vector<size_t> FindFilesWindow::CollectSelectedResultIndices() const
{
    std::vector<size_t> indices;
    if (_resultCommandIndexOverride.has_value())
    {
        indices.reserve(_resultCommandIndexOverride->size());
        for (const size_t index : _resultCommandIndexOverride.value())
        {
            if (index < _results.size() && std::ranges::find(indices, index) == indices.end())
            {
                indices.push_back(index);
            }
        }
        return indices;
    }

    if (! _resultsList)
    {
        return indices;
    }

    const auto selectedIds = _resultsList->GetSelectionModel().GetOrderedSelection();
    indices.reserve(selectedIds.size());
    for (uint64_t stableId : selectedIds)
    {
        const std::optional<size_t> row = _resultsModel.FindRowByStableId(stableId);
        if (row.has_value() && row.value() < _results.size())
        {
            indices.push_back(row.value());
        }
    }

    if (indices.empty())
    {
        const std::optional<size_t> row = GetSelectedResultIndex();
        if (row.has_value())
        {
            indices.push_back(row.value());
        }
    }

    return indices;
}

bool FindFilesWindow::CopySelectedResultsToClipboard(DWORD preferredEffect) noexcept
{
    const std::vector<size_t> indices = CollectSelectedResultIndices();
    if (indices.empty())
    {
        return false;
    }

    std::vector<std::filesystem::path> paths;
    paths.reserve(indices.size());
    for (const size_t index : indices)
    {
        if (index >= _results.size())
        {
            continue;
        }

        const FindResultRecord& result = _results[index];
        if (! NavigationLocation::LooksLikeWindowsAbsolutePath(result.fullPath))
        {
            SetStatusText(LoadStringResource(nullptr, IDS_MSG_CLIPBOARD_WRITE_FAILED));
            return true;
        }
        paths.emplace_back(result.fullPath);
    }

    if (paths.empty())
    {
        return false;
    }

    std::wstring text;
    if (preferredEffect == DROPEFFECT_MOVE)
    {
        for (size_t index = 0; index < paths.size(); ++index)
        {
            if (index != 0u)
            {
                text.append(L"\r\n");
            }
            text.append(paths[index].native());
        }
    }
    else
    {
        text = _resultsList ? _resultsList->BuildSelectionTsv() : std::wstring{};
    }
    if (text.empty())
    {
        for (size_t index = 0; index < paths.size(); ++index)
        {
            if (index != 0u)
            {
                text.append(L"\r\n");
            }
            text.append(paths[index].native());
        }
    }

    if (! SetFindResultsClipboard(_hWnd.get(), paths, preferredEffect, text))
    {
        SetStatusText(LoadStringResource(nullptr, IDS_MSG_CLIPBOARD_WRITE_FAILED));
    }
    return true;
}

bool FindFilesWindow::LaunchSelectedResultFileAction(unsigned int commandId) noexcept
{
    const std::optional<size_t> selected = GetSelectedResultIndex();
    if (! selected.has_value() || selected.value() >= _results.size())
    {
        return false;
    }

    const FindResultRecord& result = _results[selected.value()];
    if (! NavigationLocation::LooksLikeWindowsAbsolutePath(result.fullPath))
    {
        OpenSelectedResult(false);
        return true;
    }

    if ((result.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
    {
        OpenSelectedResult(false);
        return true;
    }

    const std::filesystem::path fullPath(result.fullPath);
    std::vector<std::filesystem::path> selectedPaths;
    const std::vector<size_t> selectedIndices = CollectSelectedResultIndices();
    selectedPaths.reserve(selectedIndices.size());
    for (const size_t selectedIndex : selectedIndices)
    {
        if (selectedIndex >= _results.size())
        {
            continue;
        }

        const FindResultRecord& selectedResult = _results[selectedIndex];
        if ((selectedResult.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u || ! NavigationLocation::LooksLikeWindowsAbsolutePath(selectedResult.fullPath))
        {
            continue;
        }
        if (CompareStringOrdinal(selectedResult.pluginId.c_str(), -1, result.pluginId.c_str(), -1, TRUE) != CSTR_EQUAL ||
            ! NavigationLocation::EqualsNoCase(selectedResult.instanceContext, result.instanceContext))
        {
            continue;
        }

        selectedPaths.emplace_back(selectedResult.fullPath);
    }

    std::vector<std::filesystem::path> displayedFilePaths = selectedPaths;
    const bool focusedPathListed = std::find_if(displayedFilePaths.begin(), displayedFilePaths.end(), [&](const std::filesystem::path& path) noexcept {
        return OrdinalString::EqualsNoCasePath(path, fullPath);
    }) != displayedFilePaths.end();
    if (! focusedPathListed)
    {
        displayedFilePaths.push_back(fullPath);
    }

    const bool launched = _applicationFolderWindow->TryLaunchResolvedFileAction(
        result.pluginId, result.instanceContext, fullPath, std::move(selectedPaths), std::move(displayedFilePaths), commandId, _hWnd.get());
    if (! launched)
    {
        Debug::Warning(L"FindFiles: result file action command {} was not handled for '{}'.", commandId, result.fullPath);
    }
    return true;
}

std::optional<FindFilesWindow::SelectedResultsFileOperationContext> FindFilesWindow::CollectSelectedResultsFileOperationContext(
    std::wstring_view operationLabel) const noexcept
{
    const std::vector<size_t> indices = CollectSelectedResultIndices();
    if (indices.empty())
    {
        return std::nullopt;
    }

    SelectedResultsFileOperationContext context;
    context.paths.reserve(indices.size());
    context.resultKeys.reserve(indices.size());

    for (const size_t index : indices)
    {
        if (index >= _results.size())
        {
            continue;
        }

        const FindResultRecord& result = _results[index];
        if (result.fullPath.empty())
        {
            continue;
        }

        if (context.pluginId.empty())
        {
            context.pluginId        = result.pluginId;
            context.instanceContext = result.instanceContext;
        }
        else if (CompareStringOrdinal(context.pluginId.c_str(), -1, result.pluginId.c_str(), -1, TRUE) != CSTR_EQUAL ||
                 ! NavigationLocation::EqualsNoCase(context.instanceContext, result.instanceContext))
        {
            Debug::Warning(L"FindFiles: refusing {} for a mixed plugin/context result selection.", operationLabel);
            return std::nullopt;
        }

        context.paths.emplace_back(result.fullPath);
        context.resultKeys.push_back(result.key);
    }

    if (context.paths.empty() || context.pluginId.empty())
    {
        return std::nullopt;
    }

    return context;
}

template <typename OutcomeRange>
[[nodiscard]] std::unordered_set<std::wstring> CollectKnownCompletedResultKeys(std::span<const std::wstring> resultKeys,
                                                                               const OutcomeRange& outcomes,
                                                                               HRESULT overallStatus)
{
    std::unordered_set<std::wstring> completedKeys;
    completedKeys.reserve(resultKeys.size());

    if (outcomes.empty() && overallStatus == S_OK)
    {
        completedKeys.insert(resultKeys.begin(), resultKeys.end());
        return completedKeys;
    }

    for (const auto& outcome : outcomes)
    {
        if (outcome.status == S_OK && outcome.sourceIndex < resultKeys.size())
        {
            completedKeys.insert(resultKeys[outcome.sourceIndex]);
        }
    }
    return completedKeys;
}

bool FindFilesWindow::CopyOrMoveSelectedResultsToOtherPane(FileSystemOperation operation) noexcept
{
    const bool isMove = operation == FILESYSTEM_MOVE;
    auto context      = CollectSelectedResultsFileOperationContext(isMove ? L"move" : L"copy");
    if (! context.has_value())
    {
        return false;
    }

    std::vector<std::filesystem::path> sourcePathsForCompletion = context->paths;

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
    std::optional<std::filesystem::path> destinationFolder;
    HRESULT hr      = S_OK;
    uint64_t taskId = 0;
    if (_explicitDestinationFolder.has_value() && ! _explicitDestinationFolder->empty())
    {
        destinationFolder = _explicitDestinationFolder.value();
        hr                = _applicationFolderWindow->StartFileOperationForResolvedPathsToDestination(
            context->pluginId, context->instanceContext, operation, std::move(context->paths), destinationFolder.value(), flags, &taskId);
    }
    else
    {
        hr = _applicationFolderWindow->StartFileOperationForResolvedPathsToOtherPane(
            context->pluginId, context->instanceContext, operation, std::move(context->paths), flags, &destinationFolder, &taskId);
    }
    if (FAILED(hr))
    {
        Debug::Warning(L"FindFiles: result {}-to-other-pane command failed (hr={:#010x}).", isMove ? L"move" : L"copy", static_cast<unsigned long>(hr));
    }
    else if (hr == S_OK && isMove && taskId != 0)
    {
        // Rows disappear when the move SUCCEEDS, not when it merely starts.
        EnsureFileOperationCompletedSubscription();
        const uint64_t nowTickMs = GetTickCount64();
        ReapExpiredPendingResultRemovals(nowTickMs);
        _pendingResultRemovals.push_back(PendingResultRemoval{.taskId        = taskId,
                                                              .operation     = operation,
                                                              .createdTickMs = nowTickMs,
                                                              .sourcePaths   = std::move(sourcePathsForCompletion),
                                                              .resultKeys    = std::move(context->resultKeys)});
    }
    if (hr == S_OK && destinationFolder.has_value())
    {
        SetStatusText(FormatStringResource(
            nullptr, isMove ? IDS_FIND_STATUS_MOVE_TO_OTHER_STARTED_FMT : IDS_FIND_STATUS_COPY_TO_OTHER_STARTED_FMT, destinationFolder->wstring()));
    }

    return true;
}

bool FindFilesWindow::DeleteSelectedResults(bool permanent) noexcept
{
    auto context = CollectSelectedResultsFileOperationContext(permanent ? L"permanent delete" : L"delete");
    if (! context.has_value())
    {
        return false;
    }

    std::vector<std::filesystem::path> sourcePathsForCompletion = context->paths;

    const FileSystemFlags flags = permanent ? static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE)
                                            : static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);
    uint64_t taskId             = 0;
    const HRESULT hr            = _applicationFolderWindow->StartFileOperationForResolvedPaths(
        context->pluginId, context->instanceContext, FILESYSTEM_DELETE, std::move(context->paths), flags, permanent, &taskId);
    if (FAILED(hr))
    {
        Debug::Warning(L"FindFiles: result delete command failed (permanent={} hr={:#010x}).", permanent, static_cast<unsigned long>(hr));
    }
    else if (hr == S_OK && taskId != 0)
    {
        // Rows disappear when the delete SUCCEEDS, not when it merely starts.
        EnsureFileOperationCompletedSubscription();
        const uint64_t nowTickMs = GetTickCount64();
        ReapExpiredPendingResultRemovals(nowTickMs);
        _pendingResultRemovals.push_back(PendingResultRemoval{.taskId        = taskId,
                                                              .operation     = FILESYSTEM_DELETE,
                                                              .createdTickMs = nowTickMs,
                                                              .sourcePaths   = std::move(sourcePathsForCompletion),
                                                              .resultKeys    = std::move(context->resultKeys)});
    }

    return true;
}

void FindFilesWindow::EnsureFileOperationCompletedSubscription() noexcept
{
    if (_fileOperationCompletedCallbackToken != 0)
    {
        return;
    }
    _fileOperationCompletedCallbackToken = _applicationFolderWindow->AddFileOperationCompletedCallback([this](const FolderWindow::FileOperationCompletedEvent& e) noexcept
    { OnFolderWindowFileOperationCompleted(e); },
                                                                                            _fileOperationCompletedCallbackLifetime);
}

void FindFilesWindow::ReapExpiredPendingResultRemovals(uint64_t nowTickMs) noexcept
{
    std::erase_if(_pendingResultRemovals, [nowTickMs](const PendingResultRemoval& pending) noexcept {
        return pending.createdTickMs != 0 && nowTickMs >= pending.createdTickMs && (nowTickMs - pending.createdTickMs) > kPendingResultRemovalMaxAgeMs;
    });
}

void FindFilesWindow::OnFolderWindowFileOperationCompleted(const FolderWindow::FileOperationCompletedEvent& e) noexcept
{
    ReapExpiredPendingResultRemovals(GetTickCount64());
    for (auto it = _pendingResultRemovals.begin(); it != _pendingResultRemovals.end(); ++it)
    {
        if (it->taskId != e.taskId)
        {
            continue;
        }

        const std::unordered_set<std::wstring> completedKeys =
            CollectKnownCompletedResultKeys(std::span<const std::wstring>(it->resultKeys), e.itemOutcomes, e.hr);
        if (! completedKeys.empty())
        {
            RemoveKeysFromResults(completedKeys);
            RefreshResultsView(true);
        }
        _pendingResultRemovals.erase(it);
        return;
    }
}

std::optional<std::wstring> FindFilesWindow::ResolveResultShortcutCommand(UINT message, WPARAM wParam) const noexcept
{
    if (message != WM_KEYDOWN && message != WM_SYSKEYDOWN)
    {
        return std::nullopt;
    }

    const uint32_t vk        = static_cast<uint32_t>(wParam);
    const uint32_t modifiers = ToShortcutManagerModifiers(GetEffectiveKeyboardModifiers());

    ShortcutManager shortcuts;
    Common::Settings::ShortcutsSettings defaultShortcuts;
    if (_settings && _settings->shortcuts.has_value())
    {
        shortcuts.Load(_settings->shortcuts.value());
    }
    else
    {
        defaultShortcuts = ShortcutDefaults::CreateDefaultShortcuts();
        shortcuts.Load(defaultShortcuts);
    }

    const std::optional<std::wstring_view> command =
        (vk >= VK_F1 && vk <= VK_F12) ? shortcuts.FindFunctionBarCommand(vk, modifiers) : shortcuts.FindFolderViewCommand(vk, modifiers);
    if (! command.has_value())
    {
        return std::nullopt;
    }

    return std::wstring(CanonicalizeCommandId(command.value()));
}

bool FindFilesWindow::CanHandleResultCommands() const noexcept
{
    return ! _session.IsActive() && _resultsList && _dxHost.GetFocusControl() == _resultsList;
}

bool FindFilesWindow::HandleResultCommandId(unsigned int commandId) noexcept
{
    switch (commandId)
    {
        case IDM_PANE_CLIPBOARD_COPY: return CopySelectedResultsToClipboard(DROPEFFECT_COPY);
        case IDM_PANE_CLIPBOARD_CUT: return CopySelectedResultsToClipboard(DROPEFFECT_MOVE);
        case IDM_PANE_COPY_TO_OTHER: return CopyOrMoveSelectedResultsToOtherPane(FILESYSTEM_COPY);
        case IDM_PANE_MOVE_TO_OTHER: return CopyOrMoveSelectedResultsToOtherPane(FILESYSTEM_MOVE);
        case IDM_PANE_VIEW: return LaunchSelectedResultFileAction(IDM_PANE_VIEW);
        case IDM_PANE_ALTERNATE_VIEW: return LaunchSelectedResultFileAction(IDM_PANE_ALTERNATE_VIEW);
        case IDM_PANE_EDIT: return LaunchSelectedResultFileAction(IDM_PANE_EDIT);
        case IDM_PANE_ALTERNATE_EDIT: return LaunchSelectedResultFileAction(IDM_PANE_ALTERNATE_EDIT);
        case IDM_PANE_DELETE:
        case IDM_PANE_MOVE_TO_RECYCLE_BIN: return DeleteSelectedResults(false);
        case IDM_PANE_PERMANENT_DELETE: return DeleteSelectedResults(true);
        default: break;
    }

    return false;
}

bool FindFilesWindow::HandleResultCommand(std::wstring_view commandId) noexcept
{
    const std::wstring_view canonicalCommandId = CanonicalizeCommandId(commandId);
    if (ShortcutIds::IsUnassignedCommandId(canonicalCommandId))
    {
        return true;
    }

    const std::optional<unsigned int> wmCommandId = TryGetWmCommandId(canonicalCommandId);
    if (! wmCommandId.has_value())
    {
        return false;
    }

    return HandleResultCommandId(wmCommandId.value());
}

bool FindFilesWindow::HandleResultShortcut(UINT message, WPARAM wParam) noexcept
{
    if (message != WM_KEYDOWN && message != WM_SYSKEYDOWN)
    {
        return false;
    }

    if (! CanHandleResultCommands())
    {
        return false;
    }

    const std::optional<std::wstring> commandId = ResolveResultShortcutCommand(message, wParam);
    if (! commandId.has_value())
    {
        return false;
    }

    return HandleResultCommand(commandId.value());
}

bool FindFilesWindow::FocusRootNavigation(bool editMode) noexcept
{
    const HWND rootNavigationHwnd = _rootNavigation.GetHwnd();
    if (! rootNavigationHwnd || IsWindow(rootNavigationHwnd) == FALSE)
    {
        return false;
    }

    _dxHost.SetFocusControl(nullptr);
    _rootNavigation.SetFocusRegion(NavigationView::FocusRegion::Path);
    if (editMode)
    {
        _rootNavigation.FocusAddressBar();
    }
    else
    {
        SetFocus(rootNavigationHwnd);
    }

    return IsRootNavigationFocused();
}

bool FindFilesWindow::IsRootNavigationFocused() const noexcept
{
    const HWND rootNavigationHwnd = _rootNavigation.GetHwnd();
    const HWND focused            = GetFocus();
    return rootNavigationHwnd && focused && (focused == rootNavigationHwnd || IsChild(rootNavigationHwnd, focused) != FALSE);
}

bool FindFilesWindow::HandleRootNavigationTabBridge(UINT message, WPARAM wParam) noexcept
{
    if (message != WM_KEYDOWN || wParam != VK_TAB || ! IsRootNavigationFocused())
    {
        return false;
    }

    const bool reverse = (GetEffectiveKeyboardModifiers() & kFindShortcutShift) != 0u;
    if (_hWnd)
    {
        SetFocus(_hWnd.get());
    }
    _dxHost.SetFocusControl(reverse ? static_cast<RedSalamander::DxUi::Control*>(_resultsList) : static_cast<RedSalamander::DxUi::Control*>(_nameCombo));
    return true;
}

bool FindFilesWindow::HandleRootMnemonic(UINT message, WPARAM wParam) noexcept
{
    if (message != WM_SYSCHAR)
    {
        return false;
    }

    const wchar_t mnemonic = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(wParam)));
    if (mnemonic != L'L')
    {
        return false;
    }

    return FocusRootNavigation(true);
}

void FindFilesWindow::UpdateOptionDependencies() noexcept
{
    const bool contentEnabled = _contentModeCombo && _contentModeCombo->GetSelectedIndex().value_or(0u) > 0u;
    if (_contentCombo)
    {
        _contentCombo->SetEnabled(true);
    }

    if (_matchCaseContentCheck)
    {
        _matchCaseContentCheck->SetEnabled(contentEnabled);
        if (! contentEnabled)
        {
            _matchCaseContentCheck->SetChecked(false);
        }
    }

    if (_wantSnippetsCheck)
    {
        _wantSnippetsCheck->SetEnabled(contentEnabled);
        if (! contentEnabled)
        {
            _wantSnippetsCheck->SetChecked(false);
        }
    }

    if (contentEnabled && _includeFilesCheck && ! _includeFilesCheck->IsChecked())
    {
        _includeFilesCheck->SetChecked(true);
    }

    const bool indexedAvailable = IsIndexedPreferenceAvailableForCurrentRoot();
    if (_preferIndexCheck)
    {
        _preferIndexCheck->SetEnabled(indexedAvailable);
        if (! indexedAvailable)
        {
            _preferIndexCheck->SetChecked(false);
        }
    }

    _resultsModel.SetShowSnippetColumn(contentEnabled);
    if (_resultSortSpec.direction != SortDirection::None && _resultSortSpec.columnIndex >= _resultsModel.GetColumnCount())
    {
        _resultSortSpec = {};
    }
    if (_resultsList)
    {
        ApplyResultsGridMetrics();
        _resultsList->SetSortSpec(_resultSortSpec);
        _resultsList->NotifyDataChanged();
        ApplyResultsGridLayoutFromSettings();
    }
}

void FindFilesWindow::UpdateActionButtons() noexcept
{
    const bool active       = _session.IsActive();
    const bool hasSelection = GetSelectedResultIndex().has_value();
    const bool hasResults   = ! _results.empty();
    bool layoutNeeded       = false;

    if (_findButton)
    {
        _findButton->SetEnabled(! active);
    }
    if (_appendButton)
    {
        _appendButton->SetEnabled(! active && hasResults);
    }
    if (_intersectButton)
    {
        _intersectButton->SetEnabled(! active && hasResults);
    }
    if (_subtractButton)
    {
        _subtractButton->SetEnabled(! active && hasResults);
    }
    if (_cancelButton)
    {
        _cancelButton->SetEnabled(active);
        if (_cancelButton->IsVisible() != active)
        {
            _cancelButton->SetVisible(active);
            layoutNeeded = true;
        }
    }
    if (_openButton)
    {
        _openButton->SetEnabled(! active && hasSelection);
    }
    if (_parentButton)
    {
        _parentButton->SetEnabled(! active && hasSelection);
    }
    if (_helpButton)
    {
        _helpButton->SetEnabled(true);
    }
    if (_resultsList)
    {
        _resultsList->SetEnabled(! _results.empty());
        _resultsList->SetHeaderBusy(active);
    }
    if (_rootCombo)
    {
        _rootCombo->SetEnabled(! active);
    }
    if (_nameCombo)
    {
        _nameCombo->SetEnabled(! active);
    }
    if (_nameModeCombo)
    {
        _nameModeCombo->SetEnabled(! active);
    }
    if (_contentCombo)
    {
        _contentCombo->SetEnabled(! active);
    }
    if (_contentModeCombo)
    {
        _contentModeCombo->SetEnabled(! active);
    }
    RefreshDestinationStatusText();
    if (layoutNeeded)
    {
        Layout();
    }
}

std::wstring FindFilesWindow::ResolveResultMenuShortcutText(std::wstring_view commandId) const noexcept
{
    if (commandId.empty())
    {
        return {};
    }

    ShortcutManager shortcuts;
    Common::Settings::ShortcutsSettings defaultShortcuts;
    if (_settings && _settings->shortcuts.has_value())
    {
        shortcuts.Load(_settings->shortcuts.value());
    }
    else
    {
        defaultShortcuts = ShortcutDefaults::CreateDefaultShortcuts();
        shortcuts.Load(defaultShortcuts);
    }

    const std::optional<ShortcutManager::ShortcutChord> chord = shortcuts.TryGetShortcutForCommand(commandId);
    if (! chord.has_value())
    {
        return {};
    }

    return FormatFindMenuChordText(chord->vk, chord->modifiers);
}

MenuFlyoutItem FindFilesWindow::BuildResultMenuItem(FindResultMenuTarget target, FindResultMenuAction action, bool enabled) const
{
    UINT textResourceId = 0u;
    std::wstring_view commandId;
    switch (action)
    {
        case FindResultMenuAction::Open:
            textResourceId = IDS_FIND_ACTION_OPEN;
            commandId      = L"cmd/pane/executeOpen";
            break;
        case FindResultMenuAction::GoToFolder: textResourceId = IDS_FIND_ACTION_PARENT; break;
        case FindResultMenuAction::View:
            textResourceId = IDS_CMD_VIEW;
            commandId      = L"cmd/pane/view";
            break;
        case FindResultMenuAction::AlternateView:
            textResourceId = IDS_CMD_ALTERNATE_VIEW;
            commandId      = L"cmd/pane/alternateView";
            break;
        case FindResultMenuAction::Edit:
            textResourceId = IDS_CMD_EDIT;
            commandId      = L"cmd/pane/edit";
            break;
        case FindResultMenuAction::AlternateEdit:
            textResourceId = IDS_CMD_ALTERNATE_EDIT;
            commandId      = L"cmd/pane/alternateEdit";
            break;
        case FindResultMenuAction::ClipboardCopy:
            textResourceId = IDS_CMD_CLIPBOARD_COPY;
            commandId      = L"cmd/pane/clipboardCopy";
            break;
        case FindResultMenuAction::ClipboardCut:
            textResourceId = IDS_CMD_CLIPBOARD_CUT;
            commandId      = L"cmd/pane/clipboardCut";
            break;
        case FindResultMenuAction::CopyToDestination:
            textResourceId = IDS_FIND_RESULT_MENU_COPY_TO_DESTINATION;
            commandId      = L"cmd/pane/copyToOtherPane";
            break;
        case FindResultMenuAction::MoveToDestination:
            textResourceId = IDS_FIND_RESULT_MENU_MOVE_TO_DESTINATION;
            commandId      = L"cmd/pane/moveToOtherPane";
            break;
        case FindResultMenuAction::Delete:
            textResourceId = IDS_CMD_MOVE_TO_RECYCLE_BIN;
            commandId      = L"cmd/pane/moveToRecycleBin";
            break;
        case FindResultMenuAction::PermanentDelete:
            textResourceId = IDS_CMD_PERMANENT_DELETE;
            commandId      = L"cmd/pane/permanentDelete";
            break;
    }

    return MenuFlyoutItem{
        .text            = textResourceId != 0u ? LoadStringResource(nullptr, textResourceId) : std::wstring{},
        .acceleratorText = ResolveResultMenuShortcutText(commandId),
        .enabled         = enabled,
        .commandId       = EncodeFindResultMenuCommand(target, action),
    };
}

void FindFilesWindow::AppendResultMenuActions(std::vector<MenuFlyoutItem>& items, FindResultMenuTarget target, bool includeOpenActions, bool enabled) const
{
    if (includeOpenActions)
    {
        items.push_back(BuildResultMenuItem(target, FindResultMenuAction::Open, enabled));
        items.push_back(BuildResultMenuItem(target, FindResultMenuAction::GoToFolder, enabled));
        items.push_back(MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Separator});
    }

    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::View, enabled));
    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::AlternateView, enabled));
    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::Edit, enabled));
    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::AlternateEdit, enabled));
    items.push_back(MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Separator});
    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::ClipboardCopy, enabled));
    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::ClipboardCut, enabled));
    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::CopyToDestination, enabled));
    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::MoveToDestination, enabled));
    items.push_back(MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Separator});
    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::Delete, enabled));
    items.push_back(BuildResultMenuItem(target, FindResultMenuAction::PermanentDelete, enabled));
}

void FindFilesWindow::ShowFindActionMenu(POINT screenPoint) noexcept
{
    TraceFindContextMenuDiagnostics(L"find.action-menu-show-request",
                                    L"hwnd={:#x} sessionActive={} point=({}, {})",
                                    reinterpret_cast<uintptr_t>(_hWnd.get()),
                                    _session.IsActive() ? L"true" : L"false",
                                    screenPoint.x,
                                    screenPoint.y);
    if (! _hWnd || _session.IsActive())
    {
        TraceFindContextMenuDiagnostics(L"find.action-menu-show-blocked",
                                        L"hwnd={:#x} sessionActive={}",
                                        reinterpret_cast<uintptr_t>(_hWnd.get()),
                                        _session.IsActive() ? L"true" : L"false");
        return;
    }

    const bool hasResults = ! _results.empty();
    std::vector<MenuFlyoutItem> items;
    items.reserve(4u);
    items.push_back(MenuFlyoutItem{
        .text      = LoadStringResource(nullptr, IDS_FIND_ACTION_FIND),
        .enabled   = true,
        .commandId = kFindActionMenuFind,
    });
    items.push_back(MenuFlyoutItem{
        .text      = LoadStringResource(nullptr, IDS_FIND_ACTION_REFINE_INTERSECT),
        .enabled   = hasResults,
        .commandId = kFindActionMenuIntersect,
    });
    items.push_back(MenuFlyoutItem{
        .text      = LoadStringResource(nullptr, IDS_FIND_ACTION_REFINE_SUBTRACT),
        .enabled   = hasResults,
        .commandId = kFindActionMenuSubtract,
    });
    items.push_back(MenuFlyoutItem{
        .text      = LoadStringResource(nullptr, IDS_FIND_ACTION_APPEND_TO_FOUND),
        .enabled   = hasResults,
        .commandId = kFindActionMenuAppend,
    });
    TraceFindContextMenuDiagnostics(L"find.action-menu-show-begin",
                                    L"hwnd={:#x} point=({}, {}) hasResults={} results={} items={}",
                                    reinterpret_cast<uintptr_t>(_hWnd.get()),
                                    screenPoint.x,
                                    screenPoint.y,
                                    hasResults ? L"true" : L"false",
                                    _results.size(),
                                    items.size());

    if (_findButton)
    {
        _findButton->SetPressedVisual(true);
        _dxHost.Invalidate();
    }
    const auto clearFindButtonPressed = wil::scope_exit([&]() noexcept
    {
        if (_findButton)
        {
            _findButton->SetPressedVisual(false);
            _dxHost.Invalidate();
        }
    });

    const std::optional<int> command = ContextMenu::Show(_hWnd.get(), screenPoint, items, _dxHost.GetTheme());
    TraceFindContextMenuDiagnostics(
        L"find.action-menu-show-result", L"hasCommand={} command={}", command.has_value() ? L"true" : L"false", command.value_or(-1));
    if (! command.has_value())
    {
        return;
    }

    std::optional<SearchOperation> operation;
    switch (command.value())
    {
        case kFindActionMenuFind: operation = SearchOperation::Find; break;
        case kFindActionMenuAppend: operation = SearchOperation::Append; break;
        case kFindActionMenuIntersect: operation = SearchOperation::Intersect; break;
        case kFindActionMenuSubtract: operation = SearchOperation::Subtract; break;
        default: TraceFindContextMenuDiagnostics(L"find.action-menu-command-ignored", L"command={}", command.value()); break;
    }
    if (operation.has_value())
    {
        const bool started = BeginSearch(operation.value());
        TraceFindContextMenuDiagnostics(L"find.action-menu-command-dispatched",
                                        L"command={} operation={} beginSearch={}",
                                        command.value(),
                                        TraceSearchOperationName(operation.value()),
                                        started ? L"true" : L"false");
    }
}

void FindFilesWindow::ShowResultContextMenu(size_t clickedRowIndex, POINT screenPoint) noexcept
{
    if (! _hWnd || _session.IsActive() || clickedRowIndex >= _results.size())
    {
        return;
    }

    const std::vector<size_t> selectedIndices = CollectSelectedResultIndices();
    const bool clickedIsSelected              = std::ranges::find(selectedIndices, clickedRowIndex) != selectedIndices.end();
    const bool hasMultiSelection              = clickedIsSelected && selectedIndices.size() > 1u;
    const bool hasSelection                   = ! selectedIndices.empty();

    std::vector<MenuFlyoutItem> items;
    if (hasMultiSelection)
    {
        items.reserve(30u);
        items.push_back(MenuFlyoutItem{
            .kind = RedSalamander::DxUi::MenuItemKind::Header, .text = LoadStringResource(nullptr, IDS_FIND_RESULT_MENU_THIS_ITEM), .enabled = false});
        AppendResultMenuActions(items, FindResultMenuTarget::ClickedItem, true, true);
        items.push_back(MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Separator});
        items.push_back(MenuFlyoutItem{.kind    = RedSalamander::DxUi::MenuItemKind::Header,
                                       .text    = FormatStringResource(nullptr, IDS_FIND_RESULT_MENU_SELECTION_FMT, selectedIndices.size()),
                                       .enabled = false});
        AppendResultMenuActions(items, FindResultMenuTarget::Selection, false, hasSelection);
    }
    else
    {
        items.reserve(14u);
        AppendResultMenuActions(items, FindResultMenuTarget::ClickedItem, true, true);
    }

    const std::optional<int> command = ContextMenu::Show(_hWnd.get(), screenPoint, items, _dxHost.GetTheme());
    if (! command.has_value())
    {
        return;
    }

    const std::optional<FindResultMenuCommand> resultCommand = DecodeFindResultMenuCommand(command.value());
    if (! resultCommand.has_value())
    {
        return;
    }

    static_cast<void>(DispatchResultContextMenuCommand(resultCommand.value(), clickedRowIndex));
}

bool FindFilesWindow::DispatchResultContextMenuCommand(const FindResultMenuCommand& command, size_t clickedRowIndex) noexcept
{
    if (clickedRowIndex >= _results.size())
    {
        return false;
    }

    std::optional<std::vector<size_t>> previousOverride = std::move(_resultCommandIndexOverride);
    if (command.target == FindResultMenuTarget::ClickedItem)
    {
        _resultCommandIndexOverride = std::vector<size_t>{clickedRowIndex};
    }
    else
    {
        _resultCommandIndexOverride.reset();
    }
    const auto restoreOverride = wil::scope_exit([&]() noexcept { _resultCommandIndexOverride = std::move(previousOverride); });

    switch (command.action)
    {
        case FindResultMenuAction::Open: OpenSelectedResult(false); return true;
        case FindResultMenuAction::GoToFolder: OpenSelectedResult(true); return true;
        case FindResultMenuAction::View: return HandleResultCommandId(IDM_PANE_VIEW);
        case FindResultMenuAction::AlternateView: return HandleResultCommandId(IDM_PANE_ALTERNATE_VIEW);
        case FindResultMenuAction::Edit: return HandleResultCommandId(IDM_PANE_EDIT);
        case FindResultMenuAction::AlternateEdit: return HandleResultCommandId(IDM_PANE_ALTERNATE_EDIT);
        case FindResultMenuAction::ClipboardCopy: return HandleResultCommandId(IDM_PANE_CLIPBOARD_COPY);
        case FindResultMenuAction::ClipboardCut: return HandleResultCommandId(IDM_PANE_CLIPBOARD_CUT);
        case FindResultMenuAction::CopyToDestination: return HandleResultCommandId(IDM_PANE_COPY_TO_OTHER);
        case FindResultMenuAction::MoveToDestination: return HandleResultCommandId(IDM_PANE_MOVE_TO_OTHER);
        case FindResultMenuAction::Delete: return HandleResultCommandId(IDM_PANE_MOVE_TO_RECYCLE_BIN);
        case FindResultMenuAction::PermanentDelete: return HandleResultCommandId(IDM_PANE_PERMANENT_DELETE);
    }

    return false;
}

void FindFilesWindow::ShowResultActionsHelp() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_FIND_RESULT_ACTIONS_HELP_TITLE);
    std::wstring destinationText;
    if (const auto destination = ResolveDestinationFolderForDisplay())
    {
        destinationText = destination->wstring();
    }
    else
    {
        destinationText = LoadStringResource(nullptr, IDS_FIND_STATUS_DESTINATION_UNAVAILABLE);
    }
    std::wstring message = FormatStringResource(nullptr, IDS_FIND_RESULT_ACTIONS_HELP_TEXT_FMT, destinationText);
    if (message.empty())
    {
        message = LoadStringResource(nullptr, IDS_FIND_RESULT_ACTIONS_HELP_TEXT);
    }
    if (message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1u;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODAL;
    request.severity     = HOST_ALERT_INFO;
    request.targetWindow = _hWnd.get();
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    const HRESULT hr = HostShowAlert(request);
    if (FAILED(hr))
    {
        Debug::Warning(L"FindFiles: unable to show result-actions help (hr=0x{:08X}).", static_cast<unsigned long>(hr));
    }
}

void FindFilesWindow::ApplyResultsSortFromSettings() noexcept
{
    _resultSortSpec = {};
    if (! _settings || ! _settings->search.has_value())
    {
        return;
    }

    const auto& settings = _settings->search.value();
    if (settings.sortColumnId.empty())
    {
        return;
    }

    const auto columnIndex = FindResultColumnIndexById(settings.sortColumnId);
    if (! columnIndex.has_value())
    {
        return;
    }

    _resultSortSpec.columnIndex = columnIndex.value();
    _resultSortSpec.direction   = settings.sortDescending ? SortDirection::Descending : SortDirection::Ascending;
    if (_resultsList)
    {
        _resultsList->SetSortSpec(_resultSortSpec);
    }
}

void FindFilesWindow::Layout() noexcept
{
    if (! _hWnd || ! _root)
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(_hWnd.get(), &client))
    {
        return;
    }

    const float widthDip          = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.right - client.left)));
    const float heightDip         = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.bottom - client.top)));
    const bool compact            = _theme.compactMode;
    const float margin            = compact ? 10.0f : 12.0f;
    const float gap               = compact ? 6.0f : 8.0f;
    const float labelWidth        = compact ? 72.0f : 76.0f;
    const float rowHeight         = compact ? 28.0f : 30.0f;
    const float optionHeight      = compact ? 22.0f : 24.0f;
    const float buttonHeight      = compact ? 30.0f : 32.0f;
    const float buttonWidth       = compact ? 104.0f : 108.0f;
    const float findButtonWidth   = compact ? 118.0f : 124.0f;
    const float helpButtonWidth   = compact ? 34.0f : 36.0f;
    const float modeWidth         = compact ? 142.0f : 150.0f;
    const float destinationHeight = static_cast<float>(NavigationView::kHeight);
    const auto snapDip            = [this](float dip) noexcept { return _dxHost.PixelsToDip(std::round(_dxHost.DipsToPixels(dip))); };
    const auto rect               = [&snapDip](float left, float top, float rightEdge, float bottom) noexcept
    { return D2D1::RectF(snapDip(left), snapDip(top), snapDip(rightEdge), snapDip(bottom)); };
    const auto moveChildToDip = [this, &rect](HWND hwnd, float left, float top, float rightEdge, float bottom) noexcept
    {
        if (! hwnd)
        {
            return;
        }

        const D2D1_RECT_F childRect = rect(left, top, rightEdge, bottom);
        const int x                 = static_cast<int>(std::lround(_dxHost.DipsToPixels(childRect.left)));
        const int y                 = static_cast<int>(std::lround(_dxHost.DipsToPixels(childRect.top)));
        const int width             = std::max(0, static_cast<int>(std::lround(_dxHost.DipsToPixels(childRect.right - childRect.left))));
        const int height            = std::max(0, static_cast<int>(std::lround(_dxHost.DipsToPixels(childRect.bottom - childRect.top))));
        SetWindowPos(hwnd, nullptr, x, y, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    };

    _root->SetBounds(rect(0.0f, 0.0f, widthDip, heightDip));

    float y                    = margin;
    const float right          = std::max(margin, widthDip - margin);
    const float contentWidth   = std::max(100.0f, widthDip - (2.0f * margin));
    const float comboWidth     = std::max(100.0f, contentWidth - labelWidth - gap - modeWidth - gap);
    const float fullComboWidth = std::max(100.0f, contentWidth - labelWidth - gap);

    _rootLabel->SetBounds(rect(margin, y, margin + labelWidth, y + rowHeight));
    _rootCombo->SetBounds(rect(margin + labelWidth + gap, y, margin + labelWidth + gap, y + rowHeight));
    moveChildToDip(_rootNavigation.GetHwnd(), margin + labelWidth + gap, y, margin + labelWidth + gap + fullComboWidth, y + rowHeight);
    y += rowHeight + gap;

    _nameLabel->SetBounds(rect(margin, y, margin + labelWidth, y + rowHeight));
    _nameCombo->SetBounds(rect(margin + labelWidth + gap, y, margin + labelWidth + gap + comboWidth, y + rowHeight));
    _nameModeCombo->SetBounds(rect(right - modeWidth, y, right, y + rowHeight));
    y += rowHeight + gap;

    _contentLabel->SetBounds(rect(margin, y, margin + labelWidth, y + rowHeight));
    _contentCombo->SetBounds(rect(margin + labelWidth + gap, y, margin + labelWidth + gap + comboWidth, y + rowHeight));
    _contentModeCombo->SetBounds(rect(right - modeWidth, y, right, y + rowHeight));
    y += rowHeight + gap;

    const float optionColumnWidth = std::max(120.0f, (contentWidth - (3.0f * gap)) / 4.0f);
    const float optionX0          = margin;
    const float optionX1          = optionX0 + optionColumnWidth + gap;
    const float optionX2          = optionX1 + optionColumnWidth + gap;
    const float optionX3          = optionX2 + optionColumnWidth + gap;

    _recursiveCheck->SetBounds(rect(optionX0, y, optionX0 + optionColumnWidth, y + optionHeight));
    _includeFilesCheck->SetBounds(rect(optionX1, y, optionX1 + optionColumnWidth, y + optionHeight));
    _includeDirectoriesCheck->SetBounds(rect(optionX2, y, optionX2 + optionColumnWidth, y + optionHeight));
    _followSymlinksCheck->SetBounds(rect(optionX3, y, optionX3 + optionColumnWidth, y + optionHeight));
    y += optionHeight + gap;

    _matchCaseNameCheck->SetBounds(rect(optionX0, y, optionX0 + optionColumnWidth, y + optionHeight));
    _matchCaseContentCheck->SetBounds(rect(optionX1, y, optionX1 + optionColumnWidth, y + optionHeight));
    _preferIndexCheck->SetBounds(rect(optionX2, y, optionX2 + optionColumnWidth, y + optionHeight));
    _wantSnippetsCheck->SetBounds(rect(optionX3, y, optionX3 + optionColumnWidth, y + optionHeight));
    y += optionHeight + gap;

    float buttonX = margin;
    _findButton->SetBounds(rect(buttonX, y, buttonX + findButtonWidth, y + buttonHeight));
    buttonX += findButtonWidth + gap;
    _appendButton->SetBounds(rect(buttonX, y, buttonX, y + buttonHeight));
    _intersectButton->SetBounds(rect(buttonX, y, buttonX, y + buttonHeight));
    _subtractButton->SetBounds(rect(buttonX, y, buttonX, y + buttonHeight));
    _openButton->SetBounds(rect(buttonX, y, buttonX + buttonWidth, y + buttonHeight));
    buttonX += buttonWidth + gap;
    _parentButton->SetBounds(rect(buttonX, y, buttonX + buttonWidth, y + buttonHeight));

    const float helpButtonLeft = std::max(margin, right - helpButtonWidth);
    _helpButton->SetBounds(rect(helpButtonLeft, y, helpButtonLeft + helpButtonWidth, y + buttonHeight));
    float rightActionLeft = helpButtonLeft;
    if (_session.IsActive())
    {
        const float cancelLeft = std::max(margin, helpButtonLeft - gap - buttonWidth);
        _cancelButton->SetBounds(rect(cancelLeft, y, cancelLeft + buttonWidth, y + buttonHeight));
        rightActionLeft = cancelLeft;
    }
    else
    {
        _cancelButton->SetBounds(rect(helpButtonLeft, y, helpButtonLeft, y + buttonHeight));
    }

    const float statusLeft  = buttonX + buttonWidth + gap;
    const float statusRight = std::max(statusLeft, rightActionLeft - gap);
    _statusText->SetBounds(rect(statusLeft, y, statusRight, y + buttonHeight));
    y += buttonHeight + gap;

    const float minResultsHeight       = 120.0f;
    const float destinationBottomInset = compact ? 4.0f : 6.0f;
    const float destinationTop         = std::max(y + minResultsHeight + gap, heightDip - destinationBottomInset - destinationHeight);
    _destinationLabel->SetBounds(rect(margin, destinationTop, margin + labelWidth, destinationTop + destinationHeight));
    moveChildToDip(_destinationNavigation.GetHwnd(), margin + labelWidth + gap, destinationTop, right, destinationTop + destinationHeight);

    ApplyResultsGridMetrics();
    _resultsList->SetBounds(rect(margin, y, right, std::max(y + minResultsHeight, destinationTop - (compact ? 4.0f : 6.0f))));
    _dxHost.Invalidate();
}

std::wstring FindFilesWindow::GetComboText(const ComboBox* combo) const noexcept
{
    return combo ? StripSingleLineControlCharacters(combo->GetText()) : std::wstring{};
}

void FindFilesWindow::PersistUiState(bool updateHistory) noexcept
{
    if (! _settings)
    {
        return;
    }

    _dxHost.CommitFocusedTextInput();
    Common::Settings::SearchDialogSettings settings = _settings->search.value_or(Common::Settings::SearchDialogSettings{});
    settings.lastRoot                               = GetComboText(_rootCombo);
    settings.lastNamePattern                        = GetComboText(_nameCombo);
    settings.lastContentPattern                     = GetComboText(_contentCombo);
    settings.recursive                              = _recursiveCheck && _recursiveCheck->IsChecked();
    settings.includeFiles                           = _includeFilesCheck && _includeFilesCheck->IsChecked();
    settings.includeDirectories                     = _includeDirectoriesCheck && _includeDirectoriesCheck->IsChecked();
    settings.followSymlinks                         = _followSymlinksCheck && _followSymlinksCheck->IsChecked();
    settings.matchCaseName                          = _matchCaseNameCheck && _matchCaseNameCheck->IsChecked();
    settings.matchCaseContent                       = _matchCaseContentCheck && _matchCaseContentCheck->IsChecked();
    settings.preferIndex                            = _preferIndexCheck && _preferIndexCheck->IsChecked();
    settings.wantSnippets                           = _wantSnippetsCheck && _wantSnippetsCheck->IsChecked();
    settings.nameMode    = FromNameModeComboIndex(static_cast<int>(_nameModeCombo ? _nameModeCombo->GetSelectedIndex().value_or(0u) : 0u));
    settings.contentMode = FromContentModeComboIndex(static_cast<int>(_contentModeCombo ? _contentModeCombo->GetSelectedIndex().value_or(0u) : 0u));
    settings.maxResults  = 0;
    settings.sortColumnId.clear();
    settings.sortDescending = false;
    if (_resultSortSpec.direction != SortDirection::None && _resultSortSpec.columnIndex < _resultsModel.GetColumnCount())
    {
        settings.sortColumnId   = _resultsModel.GetColumn(_resultSortSpec.columnIndex).id;
        settings.sortDescending = _resultSortSpec.direction == SortDirection::Descending;
    }
    settings.resultsGridLayout.clear();
    if (_resultsList)
    {
        settings.resultsGridLayout = ConvertColumnLayout(_resultsList->CaptureColumnLayout());
    }

    if (updateHistory)
    {
        UpdateRecentValue(settings.recentRoots, settings.lastRoot);
        UpdateRecentValue(settings.recentNamePatterns, settings.lastNamePattern);
        if (! settings.lastContentPattern.empty())
        {
            UpdateRecentValue(settings.recentContentPatterns, settings.lastContentPattern);
        }
    }

    _settings->search = std::move(settings);
}

HWND FindFilesWindow::ResolveRestoreFolderViewWindow() const noexcept
{
    HWND focusedFolderView = _applicationFolderWindow->GetFocusedFolderViewHwnd();
    if (! focusedFolderView)
    {
        focusedFolderView = _applicationFolderWindow->GetFolderViewHwnd(_applicationFolderWindow->GetFocusedPane());
    }

    if (! focusedFolderView || IsWindow(focusedFolderView) == FALSE)
    {
        return nullptr;
    }

    return focusedFolderView;
}

std::optional<SearchRequest> FindFilesWindow::BuildSearchRequest(const SearchTextOverride* textOverride) noexcept
{
    if (! textOverride)
    {
        _dxHost.CommitFocusedTextInput();
    }
    std::wstring rootPath       = textOverride ? textOverride->rootPath : GetComboText(_rootCombo);
    std::wstring namePattern    = textOverride ? textOverride->namePattern : GetComboText(_nameCombo);
    std::wstring contentPattern = textOverride ? textOverride->contentPattern : GetComboText(_contentCombo);
    rootPath                    = StripSingleLineControlCharacters(rootPath);
    namePattern                 = StripSingleLineControlCharacters(namePattern);
    contentPattern              = StripSingleLineControlCharacters(contentPattern);
    _lastBuiltRootPath          = rootPath;
    _lastBuiltNamePattern       = namePattern;
    _lastBuiltContentPattern    = contentPattern;

    if (rootPath.empty())
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_ROOT_REQUIRED));
        return std::nullopt;
    }

    const bool includeFiles = _includeFilesCheck && _includeFilesCheck->IsChecked();
    const bool includeDirs  = _includeDirectoriesCheck && _includeDirectoriesCheck->IsChecked();
    if (! includeFiles && ! includeDirs)
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_TARGETS_REQUIRED));
        return std::nullopt;
    }

    const auto configuredNameMode    = FromNameModeComboIndex(static_cast<int>(_nameModeCombo ? _nameModeCombo->GetSelectedIndex().value_or(0u) : 0u));
    const auto configuredContentMode = FromContentModeComboIndex(static_cast<int>(_contentModeCombo ? _contentModeCombo->GetSelectedIndex().value_or(0u) : 0u));
    const FileSystemSearchNameMode nameMode       = namePattern.empty() ? FILESYSTEM_SEARCH_NAME_DISABLED : ToAbiNameMode(configuredNameMode);
    const FileSystemSearchContentMode contentMode = contentPattern.empty() || configuredContentMode == Common::Settings::SearchContentMode::Disabled
                                                        ? FILESYSTEM_SEARCH_CONTENT_DISABLED
                                                        : ToAbiContentMode(configuredContentMode);

    if (configuredContentMode != Common::Settings::SearchContentMode::Disabled && contentPattern.empty())
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_CONTENT_REQUIRED));
        return std::nullopt;
    }

    if (nameMode == FILESYSTEM_SEARCH_NAME_DISABLED && contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_CRITERIA_REQUIRED));
        return std::nullopt;
    }

    SearchRequest request;
    request.context                = _context;
    request.rootPath               = std::move(rootPath);
    request.namePattern            = std::move(namePattern);
    request.contentPattern         = std::move(contentPattern);
    request.nameMode               = nameMode;
    request.contentMode            = contentMode;
    request.maxResults             = 0;
    request.maxContentBytesPerFile = kMaxContentBytes;
    request.maxSnippetCharacters   = (_wantSnippetsCheck && _wantSnippetsCheck->IsChecked()) ? kMaxSnippetCharacters : 0u;

    if (_recursiveCheck && _recursiveCheck->IsChecked())
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_RECURSIVE);
    }
    if (includeFiles || contentMode != FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_INCLUDE_FILES);
    }
    if (includeDirs)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES);
    }
    if (_followSymlinksCheck && _followSymlinksCheck->IsChecked())
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_FOLLOW_SYMLINKS);
    }
    if (_matchCaseNameCheck && _matchCaseNameCheck->IsChecked())
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_MATCH_CASE_NAME);
    }
    if (_matchCaseContentCheck && _matchCaseContentCheck->IsChecked())
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_MATCH_CASE_CONTENT);
    }
    if (request.maxSnippetCharacters != 0)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_WANT_SNIPPETS);
    }

    if (! TryResolveSearchContextForRoot(request))
    {
        SetStatusText(LoadStringResource(nullptr, IDS_FIND_ERROR_NO_FILESYSTEM));
        return std::nullopt;
    }

    const bool preferIndex = _preferIndexCheck && _preferIndexCheck->IsChecked() && IsIndexedPreferenceAvailableForCurrentRoot();
    if (preferIndex)
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_PREFER_INDEX);
    }
    else if (IsExplicitLocalSearchContext(request.context))
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_FORCE_SCAN);
    }

    return request;
}

bool FindFilesWindow::OnClose() noexcept
{
    if (_closeRequested)
    {
        return ! _session.IsActive();
    }

    _closeRequested = true;
    if (_hWnd)
    {
        static_cast<void>(::KillTimer(_hWnd.get(), kStatusRefreshTimerId));
    }
    PersistUiState(false);
    _cancelRequestedUi = true;
    _session.Cancel();
    if (_session.IsActive())
    {
        if (_hWnd)
        {
            static_cast<void>(::EnableWindow(_hWnd.get(), FALSE));
            static_cast<void>(::ShowWindow(_hWnd.get(), SW_HIDE));
        }
        EmitPerfCount(L"find.ui.close_deferred_count");
        return false;
    }

    _session.Shutdown();
    return true;
}

LRESULT FindFilesWindow::OnNcDestroy() noexcept
{
    _fileOperationCompletedCallbackLifetime.reset();
    if (_fileOperationCompletedCallbackToken != 0)
    {
        _applicationFolderWindow->RemoveFileOperationCompletedCallback(std::exchange(_fileOperationCompletedCallbackToken, 0ull));
    }
    if (_hWnd)
    {
        static_cast<void>(::KillTimer(_hWnd.get(), kStatusRefreshTimerId));
    }
    PersistUiState(false);
    if (_settings && _hWnd)
    {
        WindowPlacementPersistence::Save(*_settings, kFindFilesWindowId, _hWnd.get());
    }

    const HWND hwnd         = _hWnd.get();
    const HWND restoreOwner = (_ownerWindow && IsWindow(_ownerWindow) != FALSE) ? _ownerWindow : nullptr;
    const HWND restoreFocus = (_restoreFocusWindow && IsWindow(_restoreFocusWindow) != FALSE) ? _restoreFocusWindow : ResolveRestoreFolderViewWindow();
    _rootNavigation.Destroy();
    _destinationNavigation.Destroy();
    _dxHost.Detach();
    _root                    = nullptr;
    _rootLabel               = nullptr;
    _rootCombo               = nullptr;
    _nameLabel               = nullptr;
    _nameCombo               = nullptr;
    _nameModeCombo           = nullptr;
    _contentLabel            = nullptr;
    _contentCombo            = nullptr;
    _contentModeCombo        = nullptr;
    _recursiveCheck          = nullptr;
    _includeFilesCheck       = nullptr;
    _includeDirectoriesCheck = nullptr;
    _followSymlinksCheck     = nullptr;
    _matchCaseNameCheck      = nullptr;
    _matchCaseContentCheck   = nullptr;
    _preferIndexCheck        = nullptr;
    _wantSnippetsCheck       = nullptr;
    _findButton              = nullptr;
    _appendButton            = nullptr;
    _intersectButton         = nullptr;
    _subtractButton          = nullptr;
    _cancelButton            = nullptr;
    _openButton              = nullptr;
    _parentButton            = nullptr;
    _helpButton              = nullptr;
    _statusText              = nullptr;
    _destinationLabel        = nullptr;
    _resultsList             = nullptr;
    _hWnd.release();
    g_findFilesWindows.erase(std::remove(g_findFilesWindows.begin(), g_findFilesWindows.end(), hwnd), g_findFilesWindows.end());
    if (g_findFilesWindow == this)
    {
        g_findFilesWindow = nullptr;
        for (auto it = g_findFilesWindows.rbegin(); it != g_findFilesWindows.rend(); ++it)
        {
            const HWND candidate = *it;
            if (candidate && IsWindow(candidate) != FALSE)
            {
                g_findFilesWindow = reinterpret_cast<FindFilesWindow*>(GetWindowLongPtrW(candidate, GWLP_USERDATA));
                if (g_findFilesWindow)
                {
                    break;
                }
            }
        }
    }
    _deletePending = true;

    if (restoreOwner)
    {
        static_cast<void>(SetActiveWindow(restoreOwner));
        static_cast<void>(SetForegroundWindow(restoreOwner));
    }
    if (restoreFocus)
    {
        _applicationFolderWindow->RequestRestoreFolderViewFocus(restoreFocus);
    }
    if (_dispatchDepth == 0u)
    {
        delete this;
    }
    return 0;
}

void FindFilesWindow::ClearResults() noexcept
{
    _results.clear();
    _resultIndexByKey.clear();
    _resultsRefreshTimerArmed        = false;
    _resultsRefreshPending           = false;
    _resultsRefreshFullRebuild       = false;
    _resultsRefreshBatchCount        = 0u;
    _resultsRefreshRecordCount       = 0u;
    _resultsRefreshFirstQueuedTickMs = 0u;
    _nextResultOrdinal               = 1u;
    if (_hWnd)
    {
        static_cast<void>(::KillTimer(_hWnd.get(), kDeferredResultsRefreshTimerId));
    }
    if (_resultsList)
    {
        _resultsList->NotifyDataChanged();
        _dxHost.Invalidate();
    }
}

void FindFilesWindow::RebuildResultsList() noexcept
{
#ifdef ENABLE_TESTS
    ++_debugResultListFullRebuildCount;
#endif
    RefreshResultsView(true);
}

ResultListMutation FindFilesWindow::AddOrUpdateVisibleResult(FindResultRecord result) noexcept
{
    if (result.key.empty())
    {
        result.key = MakeResultKey(result.pluginId, result.instanceContext, result.fullPath);
    }
    if (result.displayPath.empty())
    {
        result.displayPath = BuildResultDisplayPath(result.relativePath, _lastSubmittedRootPath, result.fullPath);
    }
    if (result.iconIndex < 0)
    {
        result.iconIndex = ResolveFindResultIconIndex(result);
    }
    result.folderViewRainbowHash32 = MakeFindResultFolderViewRainbowHash32(result);
    result.stableRowId             = MakeResultStableId(result);

    const auto existingIt = _resultIndexByKey.find(result.key);
    if (existingIt != _resultIndexByKey.end())
    {
        result.arrivalOrdinal        = _results[existingIt->second].arrivalOrdinal;
        _results[existingIt->second] = std::move(result);
        return {existingIt->second, false};
    }

    result.arrivalOrdinal = _nextResultOrdinal++;
    const size_t index    = _results.size();
    _resultIndexByKey.emplace(result.key, index);
    _results.push_back(std::move(result));
    return {index, true};
}

void FindFilesWindow::RemoveKeysFromResults(const std::unordered_set<std::wstring>& keys) noexcept
{
    if (keys.empty())
    {
        return;
    }

    std::vector<FindResultRecord> filtered;
    filtered.reserve(_results.size());
    for (auto& record : _results)
    {
        if (! keys.contains(record.key))
        {
            filtered.push_back(std::move(record));
        }
    }

    _results = std::move(filtered);
    RebuildResultIndexByKey();
}

void FindFilesWindow::KeepOnlyKeysInResults(const std::unordered_set<std::wstring>& keys) noexcept
{
    std::vector<FindResultRecord> filtered;
    filtered.reserve(_results.size());
    for (auto& record : _results)
    {
        if (keys.contains(record.key))
        {
            filtered.push_back(std::move(record));
        }
    }

    _results = std::move(filtered);
    RebuildResultIndexByKey();
}

std::optional<size_t> FindFilesWindow::GetSelectedResultIndex() const noexcept
{
    if (_resultCommandIndexOverride.has_value())
    {
        for (const size_t index : _resultCommandIndexOverride.value())
        {
            if (index < _results.size())
            {
                return index;
            }
        }
        return std::nullopt;
    }

    if (! _resultsList)
    {
        return std::nullopt;
    }

    const auto selection = _resultsList->GetSelectionModel().GetOrderedSelection();
    if (selection.empty())
    {
        return std::nullopt;
    }

    const auto selectedRow = _resultsModel.FindRowByStableId(selection.back());
    if (! selectedRow || selectedRow.value() >= _results.size())
    {
        return std::nullopt;
    }
    return selectedRow;
}

std::optional<size_t> FindFilesWindow::FindResultColumnIndexById(std::wstring_view columnId) const noexcept
{
    if (columnId.empty())
    {
        return std::nullopt;
    }

    const size_t columnCount = _resultsModel.GetColumnCount();
    for (size_t columnIndex = 0; columnIndex < columnCount; ++columnIndex)
    {
        if (_resultsModel.GetColumn(columnIndex).id == columnId)
        {
            return columnIndex;
        }
    }

    return std::nullopt;
}

FindFilesWindow::ResultOpenPlan FindFilesWindow::BuildOpenPlan(const FindResultRecord& record, bool parentOnly) noexcept
{
    ResultOpenPlan plan{};
    if (record.fullPath.empty())
    {
        return plan;
    }

    const bool isDirectory = (record.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
    const std::filesystem::path fullPath(record.fullPath);
    if (parentOnly)
    {
        plan.disposition  = ResultOpenDisposition::NavigateToParent;
        plan.targetFolder = fullPath.parent_path();
        plan.focusName    = record.displayName;
        return plan;
    }

    if (isDirectory)
    {
        plan.disposition  = ResultOpenDisposition::NavigateToResult;
        plan.targetFolder = fullPath;
        return plan;
    }

    if (NavigationLocation::LooksLikeWindowsAbsolutePath(record.fullPath))
    {
        plan.disposition = ResultOpenDisposition::DefaultOpenFile;
        return plan;
    }

    plan.disposition  = ResultOpenDisposition::NavigateToParentAndOpen;
    plan.targetFolder = fullPath.parent_path();
    plan.focusName    = record.displayName;
    plan.commandId    = IDM_FOLDERVIEW_CONTEXT_OPEN;
    return plan;
}

bool FindFilesWindow::OpenFileWithDefaultApplication(const FindResultRecord& record) noexcept
{
    if ((record.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u || ! NavigationLocation::LooksLikeWindowsAbsolutePath(record.fullPath))
    {
        return false;
    }

    const std::filesystem::path fullPath(record.fullPath);
    const std::filesystem::path workingDirectory = fullPath.parent_path();
    const HINSTANCE result =
        ShellExecuteW(_hWnd.get(), L"open", fullPath.c_str(), nullptr, workingDirectory.empty() ? nullptr : workingDirectory.c_str(), SW_SHOWNORMAL);
    const auto resultCode = reinterpret_cast<INT_PTR>(result);
    if (resultCode <= 32)
    {
        Debug::Warning(L"FindFiles: failed to default-open result file '{}' (ShellExecute code={}).", record.fullPath, resultCode);
        return false;
    }

    return true;
}

void FindFilesWindow::OpenSelectedResult(bool parentOnly) noexcept
{
    const auto index = GetSelectedResultIndex();
    if (! index.has_value() || index.value() >= _results.size())
    {
        return;
    }

    const FindResultRecord& record = _results[index.value()];
    const ResultOpenPlan plan      = BuildOpenPlan(record, parentOnly);
    if (plan.disposition == ResultOpenDisposition::None)
    {
        return;
    }

    if (plan.disposition == ResultOpenDisposition::DefaultOpenFile)
    {
        static_cast<void>(OpenFileWithDefaultApplication(record));
        return;
    }

    const std::filesystem::path historyPath = NavigationLocation::FormatHistoryPath(record.pluginShortId, record.instanceContext, plan.targetFolder);
    static_cast<void>(_applicationFolderWindow->ExecuteInActivePane(historyPath, plan.focusName, plan.commandId, true));
}

void FindFilesWindow::SetStatusText(std::wstring text) noexcept
{
    _status = std::move(text);
    if (_statusText)
    {
        if (_statusText->GetSectionCount() >= 1u)
        {
            _statusText->SetSectionText(0u, _status);
        }
        else
        {
            _statusText->SetText(_status);
        }
    }
    _dxHost.Invalidate();
}

void FindFilesWindow::RefreshRootNavigationPath() noexcept
{
    if (! _rootNavigation.GetHwnd())
    {
        return;
    }

    const std::wstring rootPath = GetComboText(_rootCombo);
    if (rootPath.empty())
    {
        _rootNavigation.SetPath(std::nullopt);
        return;
    }

    _rootNavigation.SetPath(std::filesystem::path(rootPath));
}

void FindFilesWindow::RefreshRootNavigationHistory() noexcept
{
    if (! _rootNavigation.GetHwnd())
    {
        return;
    }

    std::vector<std::filesystem::path> history;
    const std::wstring currentRoot = GetComboText(_rootCombo);
    if (! currentRoot.empty())
    {
        history.emplace_back(currentRoot);
    }

    if (_settings && _settings->search.has_value())
    {
        history.reserve(_settings->search->recentRoots.size() + history.size());
        for (const std::wstring& entry : _settings->search->recentRoots)
        {
            if (! entry.empty())
            {
                history.emplace_back(entry);
            }
        }
    }

    _rootNavigation.SetHistory(history);
}

void FindFilesWindow::OnRootNavigationPathChanged(const std::optional<std::filesystem::path>& path) noexcept
{
    const std::wstring nextRoot = (path.has_value() && ! path->empty()) ? path->native() : std::wstring{};
    if (_rootCombo && GetComboText(_rootCombo) != nextRoot)
    {
        _rootCombo->SetText(nextRoot);
    }

    UpdateOptionDependencies();
    PersistUiState(false);
    RefreshRootNavigationHistory();
}

std::optional<std::filesystem::path> FindFilesWindow::ResolveDestinationFolderForDisplay() const noexcept
{
    if (_explicitDestinationFolder.has_value() && ! _explicitDestinationFolder->empty())
    {
        return _explicitDestinationFolder.value();
    }

    const std::vector<size_t> selectedIndices = CollectSelectedResultIndices();
    std::vector<std::filesystem::path> sourcePaths;
    std::wstring pluginId;
    std::wstring instanceContext;
    sourcePaths.reserve(selectedIndices.size());

    for (const size_t index : selectedIndices)
    {
        if (index >= _results.size())
        {
            continue;
        }

        const FindResultRecord& result = _results[index];
        if (result.fullPath.empty())
        {
            continue;
        }

        if (pluginId.empty())
        {
            pluginId        = result.pluginId;
            instanceContext = result.instanceContext;
        }
        else if (CompareStringOrdinal(pluginId.c_str(), -1, result.pluginId.c_str(), -1, TRUE) != CSTR_EQUAL ||
                 ! NavigationLocation::EqualsNoCase(instanceContext, result.instanceContext))
        {
            return std::nullopt;
        }

        sourcePaths.emplace_back(result.fullPath);
    }

    if (! pluginId.empty() && ! sourcePaths.empty())
    {
        if (auto destination = _applicationFolderWindow->GetOtherPaneDestinationForResolvedPaths(pluginId, instanceContext, sourcePaths))
        {
            return destination;
        }
    }

    const FolderWindow::Pane focusedPane     = _applicationFolderWindow->GetFocusedPane();
    const FolderWindow::Pane destinationPane = focusedPane == FolderWindow::Pane::Left ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
    return _applicationFolderWindow->GetCurrentPath(destinationPane);
}

std::wstring FindFilesWindow::BuildDestinationStatusText() const
{
    std::wstring destinationText;
    if (const auto destination = ResolveDestinationFolderForDisplay())
    {
        destinationText = destination->wstring();
    }
    else
    {
        destinationText = LoadStringResource(nullptr, IDS_FIND_STATUS_DESTINATION_UNAVAILABLE);
    }

    std::wstring formatted = FormatStringResource(nullptr, IDS_FIND_STATUS_DESTINATION_FMT, destinationText);
    if (formatted.empty())
    {
        formatted = destinationText;
    }
    return formatted;
}

void FindFilesWindow::RefreshDestinationStatusText() noexcept
{
    std::wstring next = BuildDestinationStatusText();
    if (next.empty())
    {
        next = LoadStringResource(nullptr, IDS_FIND_STATUS_DESTINATION_UNAVAILABLE);
    }

    if (_destinationStatus == next)
    {
        return;
    }

    _destinationStatus = std::move(next);
    RefreshDestinationNavigationPath();
    _dxHost.Invalidate();
}

void FindFilesWindow::RefreshDestinationNavigationPath() noexcept
{
    if (! _destinationNavigation.GetHwnd())
    {
        return;
    }

    const std::optional<std::filesystem::path> destination = ResolveDestinationFolderForDisplay();
    RefreshDestinationNavigationHistory(destination);
    if (destination.has_value())
    {
        _destinationNavigation.SetPath(destination.value());
    }
    else
    {
        _destinationNavigation.SetPath(std::nullopt);
    }
}

void FindFilesWindow::RefreshDestinationNavigationHistory(const std::optional<std::filesystem::path>& destination) noexcept
{
    std::vector<std::filesystem::path> history;
    const size_t maxItems = static_cast<size_t>((std::clamp)(_applicationFolderWindow->GetFolderHistoryMax(), 1u, 50u));
    history.reserve(maxItems);

    const auto appendUnique = [&](const std::filesystem::path& path) noexcept
    {
        if (path.empty() || history.size() >= maxItems)
        {
            return;
        }

        const auto existing = std::find_if(
            history.begin(), history.end(), [&](const std::filesystem::path& entry) noexcept { return OrdinalString::EqualsNoCasePath(entry, path); });
        if (existing == history.end())
        {
            history.push_back(path);
        }
    };

    if (destination.has_value())
    {
        appendUnique(destination.value());
    }

    for (const auto& entry : _applicationFolderWindow->GetFolderHistory())
    {
        appendUnique(entry);
        if (history.size() >= maxItems)
        {
            break;
        }
    }

    _destinationNavigation.SetHistory(history);
    TraceFindContextMenuDiagnostics(L"find.destination-navigation.history",
                                    L"find={:#x} nav={:#x} destination='{}' count={}",
                                    reinterpret_cast<uintptr_t>(_hWnd.get()),
                                    reinterpret_cast<uintptr_t>(_destinationNavigation.GetHwnd()),
                                    destination.has_value() ? destination->wstring() : std::wstring{},
                                    history.size());
}

void FindFilesWindow::OnDestinationNavigationPathChanged(const std::optional<std::filesystem::path>& path) noexcept
{
    if (path.has_value() && ! path->empty())
    {
        _explicitDestinationFolder = path.value();
    }
    else
    {
        _explicitDestinationFolder.reset();
    }

    RefreshDestinationStatusText();
}

void FindFilesWindow::RebuildResultIndexByKey() noexcept
{
    _resultIndexByKey.clear();
    for (size_t index = 0; index < _results.size(); ++index)
    {
        _resultIndexByKey.emplace(_results[index].key, index);
    }
}

void FindFilesWindow::SortResults() noexcept
{
    const auto sortStartedAt = SteadyClock::now();
    const auto emitSortPerf  = wil::scope_exit([&]() noexcept
    {
        const uint64_t durationUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SteadyClock::now() - sortStartedAt).count());
        Debug::Perf::Emit(L"find.ui.sort_results_ms", L"", durationUs, static_cast<uint64_t>(_results.size()), 0u, S_OK);
    });
    if (_resultSortSpec.direction == SortDirection::None)
    {
        std::stable_sort(_results.begin(), _results.end(), [](const FindResultRecord& lhs, const FindResultRecord& rhs) noexcept {
            return lhs.arrivalOrdinal < rhs.arrivalOrdinal;
        });
        RebuildResultIndexByKey();
        return;
    }

    const auto compareText = [](std::wstring_view lhs, std::wstring_view rhs) noexcept { return OrdinalString::Compare(lhs, rhs, true); };

    std::stable_sort(_results.begin(),
                     _results.end(),
                     [&](const FindResultRecord& lhs, const FindResultRecord& rhs) noexcept
    {
        int compare = 0;
        switch (_resultSortSpec.columnIndex)
        {
            case kColumnName: compare = compareText(lhs.displayName, rhs.displayName); break;
            case kColumnPath: compare = compareText(GetResultRelativeFolderPath(lhs.relativePath), GetResultRelativeFolderPath(rhs.relativePath)); break;
            case kColumnSize: compare = (lhs.endOfFile < rhs.endOfFile) ? -1 : ((lhs.endOfFile > rhs.endOfFile) ? 1 : 0); break;
            case kColumnModified: compare = (lhs.lastWriteTime < rhs.lastWriteTime) ? -1 : ((lhs.lastWriteTime > rhs.lastWriteTime) ? 1 : 0); break;
            case kColumnAttributes: compare = compareText(FormatAttributes(lhs.fileAttributes), FormatAttributes(rhs.fileAttributes)); break;
            case kColumnSnippet: compare = compareText(lhs.previewText, rhs.previewText); break;
            default: compare = compareText(lhs.displayName, rhs.displayName); break;
        }

        if (compare == 0)
        {
            return lhs.arrivalOrdinal < rhs.arrivalOrdinal;
        }
        return _resultSortSpec.direction == SortDirection::Ascending ? (compare < 0) : (compare > 0);
    });
    RebuildResultIndexByKey();
}

void FindFilesWindow::RefreshResultsView(bool fullRebuild) noexcept
{
    const auto refreshStartedAt = SteadyClock::now();
    uint64_t refreshCoreUs      = 0u;
    const auto emitRefreshPerf  = wil::scope_exit([&]() noexcept
    {
        const uint64_t durationUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SteadyClock::now() - refreshStartedAt).count());
        Debug::Perf::Emit(L"find.ui.refresh_core_us", L"", refreshCoreUs, static_cast<uint64_t>(_results.size()), fullRebuild ? 1u : 0u, S_OK);
        Debug::Perf::Emit(L"find.ui.refresh_results_view_ms", L"", durationUs, static_cast<uint64_t>(_results.size()), fullRebuild ? 1u : 0u, S_OK);
    });
    uint64_t setRowsUs          = 0u;
    uint64_t gridUpdateUs       = 0u;
    uint64_t actionButtonsUs    = 0u;
    uint64_t hostInvalidateUs   = 0u;
    bool emitFullRebuildCount   = false;
    bool emitFastRefreshCount   = false;
    bool emitGridNotifyCount    = false;
    if (fullRebuild)
    {
        emitFullRebuildCount = true;
        SortResults();
    }

    const auto setRowsStartedAt = SteadyClock::now();
    _resultsModel.SetRows(&_results);
    setRowsUs = ElapsedUsSince(setRowsStartedAt);
    if (_resultsList)
    {
        const auto gridUpdateStartedAt = SteadyClock::now();
        _resultsList->SetSortSpec(_resultSortSpec);
        const bool useFastRefresh = ! fullRebuild && _resultSortSpec.direction == SortDirection::None;
        if (useFastRefresh)
        {
            emitFastRefreshCount = true;
        }
        else
        {
            emitGridNotifyCount = true;
            _resultsList->NotifyDataChanged();
        }
        RestorePendingSelectionIfAvailable();
        gridUpdateUs = ElapsedUsSince(gridUpdateStartedAt);
    }

    const auto actionButtonsStartedAt = SteadyClock::now();
    UpdateActionButtons();
    actionButtonsUs = ElapsedUsSince(actionButtonsStartedAt);

    const auto invalidateStartedAt = SteadyClock::now();
    _dxHost.Invalidate();
    hostInvalidateUs = ElapsedUsSince(invalidateStartedAt);
    refreshCoreUs    = ElapsedUsSince(refreshStartedAt);
    if (emitFullRebuildCount)
    {
        EmitPerfCount(L"find.ui.results_full_rebuild_count");
    }
    if (emitFastRefreshCount)
    {
        EmitPerfCount(L"find.ui.results_fast_refresh_count");
    }
    if (emitGridNotifyCount)
    {
        EmitPerfCount(L"find.ui.results_grid_notify_count");
    }
    Debug::Perf::Emit(L"find.ui.refresh_set_rows_us", L"", setRowsUs, static_cast<uint64_t>(_results.size()), fullRebuild ? 1u : 0u, S_OK);
    Debug::Perf::Emit(L"find.ui.refresh_grid_update_us", L"", gridUpdateUs, emitFastRefreshCount ? 1u : 0u, fullRebuild ? 1u : 0u, S_OK);
    Debug::Perf::Emit(
        L"find.ui.refresh_action_buttons_us", L"", actionButtonsUs, _session.IsActive() ? 1u : 0u, GetSelectedResultIndex().has_value() ? 1u : 0u, S_OK);
    Debug::Perf::Emit(L"find.ui.refresh_host_invalidate_us", L"", hostInvalidateUs, static_cast<uint64_t>(_results.size()), fullRebuild ? 1u : 0u, S_OK);
}

void FindFilesWindow::RestorePendingSelectionIfAvailable() noexcept
{
    if (! _resultsList || _pendingSelectionFullPath.empty())
    {
        return;
    }

    if (GetSelectedResultIndex().has_value())
    {
        _pendingSelectionFullPath.clear();
        return;
    }

    const auto it = std::find_if(
        _results.begin(), _results.end(), [this](const FindResultRecord& record) noexcept { return record.fullPath == _pendingSelectionFullPath; });
    if (it == _results.end())
    {
        return;
    }

    const size_t rowIndex = static_cast<size_t>(std::distance(_results.begin(), it));
    _resultsList->GetSelectionModel().SetSingle(_resultsModel.GetStableRowId(rowIndex));
    _pendingSelectionFullPath.clear();
}

void FindFilesWindow::ScheduleResultsRefresh(bool fullRebuild, size_t changedCount) noexcept
{
    if (_resultsRefreshBatchCount == 0u)
    {
        _resultsRefreshFirstQueuedTickMs = ::GetTickCount64();
    }
    _resultsRefreshFullRebuild = _resultsRefreshFullRebuild || fullRebuild;
    _resultsRefreshBatchCount += 1u;
    _resultsRefreshRecordCount += changedCount;

    if (_resultsRefreshPending)
    {
        EmitPerfCount(L"find.ui.deferred_refresh_coalesced_count");
        return;
    }

    _resultsRefreshPending = true;
    if (! _hWnd)
    {
        ApplyPendingResultsRefresh();
        return;
    }

    if (::PostMessageW(_hWnd.get(), WndMsg::kFindSearchDeferredRefresh, 0, 0) == 0)
    {
        ApplyPendingResultsRefresh();
        return;
    }

    _resultsRefreshTimerArmed = true;
    EmitPerfCount(L"find.ui.deferred_refresh_post_count");
}

void FindFilesWindow::ApplyPendingResultsRefresh() noexcept
{
    if (! _resultsRefreshPending && _resultsRefreshBatchCount == 0u)
    {
        return;
    }

    const bool fullRebuild     = _resultsRefreshFullRebuild;
    const uint64_t batchCount  = static_cast<uint64_t>(_resultsRefreshBatchCount);
    const uint64_t recordCount = static_cast<uint64_t>(_resultsRefreshRecordCount);

    _resultsRefreshTimerArmed        = false;
    _resultsRefreshPending           = false;
    _resultsRefreshFullRebuild       = false;
    _resultsRefreshBatchCount        = 0u;
    _resultsRefreshRecordCount       = 0u;
    _resultsRefreshFirstQueuedTickMs = 0u;

    EmitPerfCount(L"find.ui.deferred_refresh_batch_count", batchCount);
    EmitPerfCount(L"find.ui.deferred_refresh_result_count", recordCount);
    const auto refreshStartedAt = SteadyClock::now();
    RefreshResultsView(fullRebuild);
    const uint64_t durationUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SteadyClock::now() - refreshStartedAt).count());
    Debug::Perf::Emit(L"find.ui.deferred_refresh_handler_ms", L"", durationUs, batchCount, recordCount, S_OK);
}

void FindFilesWindow::OnGridSortRequested(const GridSortSpec& sortSpec)
{
    _resultSortSpec = sortSpec;
    RebuildResultsList();
    PersistUiState(false);
}

void FindFilesWindow::OnGridSelectionChanged()
{
    UpdateActionButtons();
}

void FindFilesWindow::OnGridRowActivated(size_t rowIndex)
{
    if (rowIndex >= _results.size())
    {
        return;
    }

    OpenSelectedResult(false);
}

void FindFilesWindow::OnGridContextMenu(Grid& sender, size_t rowIndex, POINT screenPoint)
{
    if (&sender != _resultsList)
    {
        return;
    }

    ShowResultContextMenu(rowIndex, screenPoint);
}

wil::com_ptr<ID2D1Bitmap1> FindFilesWindow::GetGridIconBitmap(const Grid& sender, int iconIndex, float targetDipSize, ID2D1DeviceContext* d2dContext)
{
    if (&sender != _resultsList || iconIndex < 0 || ! d2dContext)
    {
        return nullptr;
    }

    return IconCache::GetInstance().GetIconBitmap(iconIndex, d2dContext, targetDipSize);
}

std::wstring FindFilesWindow::BuildWarningSummary(uint32_t warningFlags) const noexcept
{
    std::vector<std::wstring> warnings;
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_NO_INDEX));
    }
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_NO_CONTENT));
    }
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_ACCESS_DENIED));
    }
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_OVERFLOW) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_OVERFLOW));
    }
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_SERVICE_UNAVAILABLE));
    }
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_ROOT_REJECTED) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_SERVICE_ROOT_REJECTED));
    }
    if ((warningFlags & FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED) != 0)
    {
        warnings.push_back(LoadStringResource(nullptr, IDS_FIND_WARNING_REGEX_REJECTED));
    }

    std::wstring result;
    for (size_t i = 0; i < warnings.size(); ++i)
    {
        if (i != 0)
        {
            result.append(L", ");
        }
        result.append(warnings[i]);
    }
    return result;
}

UINT FindFilesWindow::BackendStringId(FileSystemSearchBackend backend) const noexcept
{
    switch (backend)
    {
        case FILESYSTEM_SEARCH_BACKEND_UNKNOWN: return IDS_FIND_BACKEND_UNKNOWN;
        case FILESYSTEM_SEARCH_BACKEND_SCAN: return IDS_FIND_BACKEND_SCAN;
        case FILESYSTEM_SEARCH_BACKEND_INDEX: return IDS_FIND_BACKEND_INDEX;
        case FILESYSTEM_SEARCH_BACKEND_SERVICE: return IDS_FIND_BACKEND_SERVICE;
        default: return IDS_FIND_BACKEND_UNKNOWN;
    }
}

UINT FindFilesWindow::PhaseStringId(FileSystemSearchPhase phase) const noexcept
{
    switch (phase)
    {
        case FILESYSTEM_SEARCH_PHASE_INITIALIZING: return static_cast<UINT>(IDS_FIND_PHASE_INITIALIZING);
        case FILESYSTEM_SEARCH_PHASE_ENUMERATING: return static_cast<UINT>(IDS_FIND_PHASE_ENUMERATING);
        case FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP: return static_cast<UINT>(IDS_FIND_PHASE_INDEX_LOOKUP);
        case FILESYSTEM_SEARCH_PHASE_CONTENT_SCAN: return static_cast<UINT>(IDS_FIND_PHASE_CONTENT_SCAN);
        case FILESYSTEM_SEARCH_PHASE_COMPLETED: return static_cast<UINT>(IDS_FIND_PHASE_COMPLETED);
        default: return static_cast<UINT>(IDS_FIND_PHASE_INITIALIZING);
    }
}

std::wstring FindFilesWindow::BuildRunningStatusText() const noexcept
{
    const std::wstring backendText = LoadStringResource(nullptr, BackendStringId(_lastBackend));
    const std::wstring phaseText   = LoadStringResource(nullptr, PhaseStringId(_lastPhase));
    const uint64_t activityTickMs  = _lastProgressTickMs != 0u ? _lastProgressTickMs : _searchStartedTickMs;
    const uint64_t activityAgeMs   = GetTickAgeMs(activityTickMs);
    const bool stalled             = _lastPhase != FILESYSTEM_SEARCH_PHASE_COMPLETED && activityAgeMs >= kStatusStallThresholdMs;
    const std::wstring ageText     = FormatStatusAgeText(activityAgeMs);
    const std::wstring pathText    = _lastCurrentPath.empty() ? _lastSearchRootPath : _lastCurrentPath;
    const std::wstring spinnerText(GetStatusSpinnerFrame());

    const UINT formatId = pathText.empty() ? static_cast<UINT>(stalled ? IDS_FIND_STATUS_RUNNING_STALLED_FMT : IDS_FIND_STATUS_RUNNING_ACTIVE_FMT)
                                           : static_cast<UINT>(stalled ? IDS_FIND_STATUS_RUNNING_STALLED_PATH_FMT : IDS_FIND_STATUS_RUNNING_ACTIVE_PATH_FMT);

    std::wstring running;
    if (pathText.empty())
    {
        running = FormatStringResource(
            nullptr, formatId, spinnerText, backendText, phaseText, ageText, _lastScannedDirectories, _lastScannedFiles, _lastMatchedEntries);
    }
    else
    {
        running = FormatStringResource(
            nullptr, formatId, spinnerText, backendText, phaseText, ageText, _lastScannedDirectories, _lastScannedFiles, _lastMatchedEntries, pathText);
    }

    const std::wstring warnings = BuildWarningSummary(_lastWarningFlags);
    if (! warnings.empty())
    {
        running.append(L" [");
        running.append(warnings);
        running.push_back(L']');
    }

    const std::wstring backendStatus = BuildBackendStatusText();
    if (! backendStatus.empty())
    {
        running.append(L" | ");
        running.append(backendStatus);
    }
    return running;
}

std::wstring FindFilesWindow::BuildBackendStatusText() const noexcept
{
    if (_lastBackend == FILESYSTEM_SEARCH_BACKEND_SERVICE)
    {
        if (! _hasServiceStatus)
        {
            return FAILED(_lastServiceStatusHr) ? L"db status unavailable" : L"db status pending";
        }

        std::wstring text = std::format(L"db {} / {}",
                                        LocalSearchIndexCore::GetStoreStateText(_lastServiceStatus.storeState),
                                        LocalSearchIndexCore::GetSyncPhaseText(_lastServiceStatus.syncPhase));
        if (_lastServiceStatus.totalRoots != 0u)
        {
            text.append(std::format(L" {}/{}", _lastServiceStatus.completedRoots, _lastServiceStatus.totalRoots));
        }
        if (! _lastServiceStatus.activeRoot.empty())
        {
            text.append(L" @ ");
            text.append(TruncateMiddleText(_lastServiceStatus.activeRoot, 36u));
        }

        text.append(std::format(L" | search {}", LocalSearchIndexCore::GetQueryExecutionModeText(_lastServiceStatus.queryExecutionMode)));
        if (_lastServiceStatus.fallbackReason != LocalSearchIndexCore::FallbackReason::None)
        {
            text.append(std::format(L" ({})", LocalSearchIndexCore::GetFallbackReasonText(_lastServiceStatus.fallbackReason)));
        }
        return text;
    }

    if (_lastBackend == FILESYSTEM_SEARCH_BACKEND_INDEX)
    {
        return std::format(L"index {}", LoadStringResource(nullptr, PhaseStringId(_lastPhase)));
    }

    return {};
}

std::wstring FindFilesWindow::BuildStatusText() const noexcept
{
    if (_session.IsActive())
    {
        if (_cancelRequestedUi)
        {
            return LoadStringResource(nullptr, IDS_FIND_STATUS_CANCELLING);
        }

        return BuildRunningStatusText();
    }

    if (_lastStatusHint == HRESULT_FROM_WIN32(ERROR_CANCELLED))
    {
        return FormatStringResource(
            nullptr, IDS_FIND_STATUS_CANCELLED_FMT, LoadStringResource(nullptr, BackendStringId(_lastBackend)), _lastMatchedEntries, _lastScannedFiles);
    }

    if (FAILED(_lastStatusHint))
    {
        std::wstring failed         = FormatStringResource(nullptr,
                                                           IDS_FIND_STATUS_FAILED_FMT,
                                                           LoadStringResource(nullptr, BackendStringId(_lastBackend)),
                                                           FormatSearchStatusHint(_lastStatusHint),
                                                           static_cast<unsigned long>(_lastStatusHint));
        const std::wstring warnings = BuildWarningSummary(_lastWarningFlags);
        if (! warnings.empty())
        {
            failed.append(L" [");
            failed.append(warnings);
            failed.push_back(L']');
        }
        return failed;
    }

    if (_lastBackend != FILESYSTEM_SEARCH_BACKEND_UNKNOWN || _lastMatchedEntries != 0 || _lastScannedFiles != 0)
    {
        std::wstring complete = FormatStringResource(
            nullptr, IDS_FIND_STATUS_COMPLETE_FMT, LoadStringResource(nullptr, BackendStringId(_lastBackend)), _lastMatchedEntries, _lastScannedFiles);
        const std::wstring warnings = BuildWarningSummary(_lastWarningFlags);
        if (! warnings.empty())
        {
            complete.append(L" [");
            complete.append(warnings);
            complete.push_back(L']');
        }
        if (_lastBackend == FILESYSTEM_SEARCH_BACKEND_SERVICE)
        {
            const std::wstring backendStatus = BuildBackendStatusText();
            if (! backendStatus.empty())
            {
                complete.append(L" | ");
                complete.append(backendStatus);
            }
        }
        return complete;
    }

    return {};
}

void FindFilesWindow::RefreshStatusText() noexcept
{
    EmitPerfCount(L"find.ui.status_refresh_count");
    const Debug::Perf::Scope statusPerf(L"find.ui.status_refresh_ms");
    SetStatusText(BuildStatusText());
}

void FindFilesWindow::UpdateStatusRefreshTimer() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (_session.IsActive())
    {
        static_cast<void>(::SetTimer(_hWnd.get(), kStatusRefreshTimerId, kStatusRefreshTimerIntervalMs, nullptr));
    }
    else
    {
        static_cast<void>(::KillTimer(_hWnd.get(), kStatusRefreshTimerId));
    }
}

void FindFilesWindow::OnStatusRefreshTimer() noexcept
{
    if (! _session.IsActive())
    {
        UpdateStatusRefreshTimer();
        return;
    }

    RefreshStatusText();
}

void FindFilesWindow::OnSearchStarted(SearchOperation operation, const SearchRequest& request) noexcept
{
    _activeOperation             = operation;
    _cancelRequestedUi           = false;
    _lastBackend                 = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
    _lastPhase                   = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    _lastWarningFlags            = FILESYSTEM_SEARCH_WARNING_NONE;
    _lastScannedDirectories      = 0;
    _lastScannedFiles            = 0;
    _lastCandidateFiles          = 0;
    _lastMatchedEntries          = 0;
    _lastStatusHint              = S_OK;
    _lastCurrentPath             = request.rootPath;
    _lastSearchRootPath          = request.rootPath;
    _lastSubmittedRootPath       = request.rootPath;
    _lastSubmittedNamePattern    = request.namePattern;
    _lastSubmittedContentPattern = request.contentPattern;
    _searchStartedTickMs         = ::GetTickCount64();
    _lastProgressTickMs          = _searchStartedTickMs;
    _lastBackendStatusTickMs     = 0u;
    _lastBackendStatusPollTickMs = 0u;
    _lastServiceStatus           = {};
    _lastServiceStatusHr         = S_OK;
    _hasServiceStatus            = false;
    _deferredKeys.clear();
#ifdef ENABLE_TESTS
    _debugIncrementalResultRefreshCount        = 0u;
    _debugIncrementalVisibleResultRefreshCount = 0u;
#endif

    if (operation == SearchOperation::Find)
    {
        _pendingSelectionFullPath.clear();
        if (const auto selectedResult = GetSelectedResultIndex(); selectedResult.has_value() && selectedResult.value() < _results.size())
        {
            _pendingSelectionFullPath = _results[selectedResult.value()].fullPath;
        }
        ClearResults();
    }

    RefreshStatusText();
    UpdateStatusRefreshTimer();
    UpdateActionButtons();
}

bool FindFilesWindow::BeginSearch(SearchOperation operation, const SearchTextOverride* textOverride) noexcept
{
    TraceFindContextMenuDiagnostics(L"find.begin-search-request",
                                    L"operation={} sessionActive={} textOverride={} hwnd={:#x}",
                                    TraceSearchOperationName(operation),
                                    _session.IsActive() ? L"true" : L"false",
                                    textOverride ? L"true" : L"false",
                                    reinterpret_cast<uintptr_t>(_hWnd.get()));
    if (_session.IsActive())
    {
        TraceFindContextMenuDiagnostics(L"find.begin-search-blocked", L"reason=session-active operation={}", TraceSearchOperationName(operation));
        return false;
    }

    _lastBeginRootPath       = textOverride ? textOverride->rootPath : GetComboText(_rootCombo);
    _lastBeginNamePattern    = textOverride ? textOverride->namePattern : GetComboText(_nameCombo);
    _lastBeginContentPattern = textOverride ? textOverride->contentPattern : GetComboText(_contentCombo);

    std::optional<SearchRequest> request = BuildSearchRequest(textOverride);
    if (! request.has_value())
    {
        TraceFindContextMenuDiagnostics(L"find.begin-search-blocked", L"reason=build-request operation={}", TraceSearchOperationName(operation));
        return false;
    }

    PersistUiState(true);
    PopulateHistoryCombos();
    const uint64_t searchEpoch = ++_activeSearchEpoch;
    OnSearchStarted(operation, request.value());

    if (! _session.Start(*this, std::move(request.value()), searchEpoch))
    {
        TraceFindContextMenuDiagnostics(L"find.begin-search-blocked", L"reason=session-start operation={}", TraceSearchOperationName(operation));
        _lastStatusHint = E_FAIL;
        RefreshStatusText();
        UpdateStatusRefreshTimer();
        UpdateActionButtons();
        return false;
    }

    // OnSearchStarted prepares the UI before the worker flips the session active bit.
    // Refresh the action surface again after Start succeeds so the enabled state
    // reflects the now-active session immediately instead of waiting for progress.
    RefreshStatusText();
    UpdateStatusRefreshTimer();
    UpdateActionButtons();
    TraceFindContextMenuDiagnostics(L"find.begin-search-started", L"operation={}", TraceSearchOperationName(operation));
    return true;
}

void FindFilesWindow::OnSearchResults(WPARAM operationKey, LPARAM lParam) noexcept
{
    uint64_t drainedResultsRecordCount = 0u;
    auto drained = TakeAndCoalesceContiguousPostedPayloads<FindSearchResultsPayload>(
        _hWnd.get(),
        WndMsg::kFindSearchResults,
        operationKey,
        lParam,
        [](const FindSearchResultsPayload& current, uint64_t drainedPayloadCount) noexcept
        { return drainedPayloadCount < kResultsDrainMaxMessages && current.results.size() < kResultsDrainMaxRecords; },
        [&](std::unique_ptr<FindSearchResultsPayload>& current, std::unique_ptr<FindSearchResultsPayload> newer) noexcept
        {
            if (newer->epoch != current->epoch)
            {
                return;
            }
            if (newer->enqueuedAt != SteadyClock::time_point{})
            {
                Debug::Perf::Emit(L"find.ui.results_to_visible_latency_ms",
                                  L"",
                                  ElapsedUsSince(newer->enqueuedAt),
                                  static_cast<uint64_t>(newer->results.size()),
                                  0u,
                                  S_OK);
            }
            drainedResultsRecordCount += static_cast<uint64_t>(newer->results.size());
            current->results.insert(
                current->results.end(), std::make_move_iterator(newer->results.begin()), std::make_move_iterator(newer->results.end()));
        });
    auto payload = std::move(drained.payload);
    if (! payload)
    {
        return;
    }
    if (payload->epoch != _activeSearchEpoch || static_cast<WPARAM>(payload->epoch) != operationKey)
    {
        return;
    }

    const auto handlerStartedAt = SteadyClock::now();
    const Debug::Perf::Scope resultsPerf(L"find.ui.results_handler_ms");
    EmitPerfCount(L"find.ui.results_message_count");
    TraceFindContextMenuDiagnostics(L"find.ui.results-handler-enter",
                                    L"hwnd={:#x} initialRecords={} pendingRefresh={} timerArmed={}",
                                    reinterpret_cast<uintptr_t>(_hWnd.get()),
                                    payload->results.size(),
                                    _resultsRefreshPending ? 1 : 0,
                                    _resultsRefreshTimerArmed ? 1 : 0);
    if (payload->enqueuedAt != SteadyClock::time_point{})
    {
        Debug::Perf::Emit(
            L"find.ui.results_to_visible_latency_ms", L"", ElapsedUsSince(payload->enqueuedAt), static_cast<uint64_t>(payload->results.size()), 0u, S_OK);
    }

    const uint64_t drainedResultsMessageCount = drained.drainedPayloadCount;
    if (drained.stoppedAtQueuedMessage)
    {
        const MSG& queuedMessage = drained.queuedMessage;
        TraceFindContextMenuDiagnostics(L"find.ui.results-drain-stop",
                                        L"reason=queue-head hwnd={:#x} msg={} msgId=0x{:x} expected={} expectedId=0x{:x}",
                                        reinterpret_cast<uintptr_t>(queuedMessage.hwnd),
                                        TraceFindQueueMessageName(queuedMessage.message),
                                        queuedMessage.message,
                                        TraceFindQueueMessageName(WndMsg::kFindSearchResults),
                                        WndMsg::kFindSearchResults);
    }
    if (drainedResultsMessageCount >= kResultsDrainMaxMessages || payload->results.size() >= kResultsDrainMaxRecords)
    {
        TraceFindContextMenuDiagnostics(
            L"find.ui.results-drain-stop", L"reason=budget messages={} records={}", drainedResultsMessageCount, payload->results.size());
    }
    if (drainedResultsMessageCount != 0u)
    {
        EmitPerfCount(L"find.ui.results_messages_coalesced_count");
        EmitPerfCount(L"find.ui.results_messages_drained", drainedResultsMessageCount);
        EmitPerfCount(L"find.ui.results_records_drained", drainedResultsRecordCount);
    }
    EmitPerfCount(L"find.ui.results_message_batch_size", static_cast<uint64_t>(payload->results.size()));

    _lastMatchedEntries += static_cast<uint64_t>(payload->results.size());

    if (_activeOperation == SearchOperation::Intersect || _activeOperation == SearchOperation::Subtract)
    {
        for (auto& record : payload->results)
        {
            _deferredKeys.insert(record.key);
        }
        RefreshStatusText();
        const uint64_t elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SteadyClock::now() - handlerStartedAt).count());
        TraceFindContextMenuDiagnostics(L"find.ui.results-handler-exit",
                                        L"hwnd={:#x} setOperation=1 records={} drainedMessages={} drainedRecords={} elapsedUs={}",
                                        reinterpret_cast<uintptr_t>(_hWnd.get()),
                                        payload->results.size(),
                                        drainedResultsMessageCount,
                                        drainedResultsRecordCount,
                                        elapsedUs);
        return;
    }

    const size_t previousResultCount = _results.size();
    bool changed                     = false;
    size_t insertedCount             = 0u;
    size_t updatedCount              = 0u;
    const auto applyStartedAt        = SteadyClock::now();
    for (auto& record : payload->results)
    {
        const ResultListMutation mutation = AddOrUpdateVisibleResult(std::move(record));
        insertedCount += mutation.inserted ? 1u : 0u;
        updatedCount += mutation.inserted ? 0u : 1u;
        changed = true;
    }
    const uint64_t applyUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SteadyClock::now() - applyStartedAt).count());
    Debug::Perf::Emit(
        L"find.ui.results_apply_records_us", L"", applyUs, static_cast<uint64_t>(payload->results.size()), static_cast<uint64_t>(previousResultCount), S_OK);

    if (! changed)
    {
        RefreshStatusText();
        const uint64_t elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SteadyClock::now() - handlerStartedAt).count());
        TraceFindContextMenuDiagnostics(L"find.ui.results-handler-exit",
                                        L"hwnd={:#x} changed=0 records={} drainedMessages={} drainedRecords={} elapsedUs={}",
                                        reinterpret_cast<uintptr_t>(_hWnd.get()),
                                        payload->results.size(),
                                        drainedResultsMessageCount,
                                        drainedResultsRecordCount,
                                        elapsedUs);
        return;
    }

    EmitPerfCount(L"find.ui.results_inserted_count", static_cast<uint64_t>(insertedCount));
    EmitPerfCount(L"find.ui.results_updated_count", static_cast<uint64_t>(updatedCount));

    if (_resultSortSpec.direction == SortDirection::None)
    {
        RefreshResultsView(false);
#ifdef ENABLE_TESTS
        DebugRecordIncrementalResultRefresh();
#endif
    }
    else
    {
        // Coalesce sorted result refreshes so progress-driven micro-batches do not force repeated full rebuilds.
        ScheduleResultsRefresh(true, payload->results.size());
        const bool firstVisibleBatch = previousResultCount == 0u;
        const bool pendingBatchLarge = _resultsRefreshRecordCount >= kInteractiveResultsRefreshRecords;
        const bool pendingBatchOld =
            _resultsRefreshFirstQueuedTickMs != 0u && GetTickAgeMs(_resultsRefreshFirstQueuedTickMs) >= kInteractiveResultsRefreshMaxAgeMs;
        if (firstVisibleBatch || pendingBatchLarge || pendingBatchOld)
        {
            EmitPerfCount(L"find.ui.deferred_refresh_forced_count");
            ApplyPendingResultsRefresh();
#ifdef ENABLE_TESTS
            DebugRecordIncrementalResultRefresh();
#endif
        }
    }
    RefreshStatusText();
    const uint64_t elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SteadyClock::now() - handlerStartedAt).count());
    TraceFindContextMenuDiagnostics(L"find.ui.results-handler-exit",
                                    L"hwnd={:#x} changed=1 records={} inserted={} updated={} drainedMessages={} drainedRecords={} applyUs={} elapsedUs={}",
                                    reinterpret_cast<uintptr_t>(_hWnd.get()),
                                    payload->results.size(),
                                    insertedCount,
                                    updatedCount,
                                    drainedResultsMessageCount,
                                    drainedResultsRecordCount,
                                    applyUs,
                                    elapsedUs);
}

void FindFilesWindow::OnSearchProgress(WPARAM operationKey, LPARAM lParam) noexcept
{
    auto drained = TakeAndCoalesceContiguousPostedPayloads<FindSearchProgressPayload>(
        _hWnd.get(),
        WndMsg::kFindSearchProgress,
        operationKey,
        lParam,
        [](const FindSearchProgressPayload&, uint64_t) noexcept { return true; },
        [](std::unique_ptr<FindSearchProgressPayload>& current, std::unique_ptr<FindSearchProgressPayload> newer) noexcept
        {
            if (newer->epoch == current->epoch)
            {
                current = std::move(newer);
            }
        });
    auto payload = std::move(drained.payload);
    if (! payload)
    {
        return;
    }
    if (payload->epoch != _activeSearchEpoch || static_cast<WPARAM>(payload->epoch) != operationKey)
    {
        return;
    }

    const auto handlerStartedAt = SteadyClock::now();
    const Debug::Perf::Scope progressPerf(L"find.ui.progress_handler_ms");
    TraceFindContextMenuDiagnostics(L"find.ui.progress-handler-enter",
                                    L"hwnd={:#x} phase={} scannedDirs={} scannedFiles={} matches={}",
                                    reinterpret_cast<uintptr_t>(_hWnd.get()),
                                    static_cast<uint32_t>(payload->phase),
                                    payload->scannedDirectories,
                                    payload->scannedFiles,
                                    payload->matchedEntries);
    const uint64_t drainedProgressCount = drained.drainedPayloadCount;
    if (drained.stoppedAtQueuedMessage)
    {
        const MSG& queuedMessage = drained.queuedMessage;
        TraceFindContextMenuDiagnostics(L"find.ui.progress-drain-stop",
                                        L"reason=queue-head hwnd={:#x} msg={} msgId=0x{:x} expected={} expectedId=0x{:x}",
                                        reinterpret_cast<uintptr_t>(queuedMessage.hwnd),
                                        TraceFindQueueMessageName(queuedMessage.message),
                                        queuedMessage.message,
                                        TraceFindQueueMessageName(WndMsg::kFindSearchProgress),
                                        WndMsg::kFindSearchProgress);
    }

    EmitPerfCount(L"find.ui.progress_message_count");
    if (drainedProgressCount != 0u)
    {
        EmitPerfCount(L"find.ui.progress_messages_coalesced_count");
        EmitPerfCount(L"find.ui.progress_messages_drained", drainedProgressCount);
    }
    if (payload->enqueuedAt != SteadyClock::time_point{})
    {
        Debug::Perf::Emit(
            L"find.ui.progress_to_visible_latency_ms", L"", ElapsedUsSince(payload->enqueuedAt), static_cast<uint64_t>(payload->currentPath.size()), 0u, S_OK);
    }

    _lastBackend            = payload->backend;
    _lastPhase              = payload->phase;
    _lastWarningFlags       = payload->warningFlags;
    _lastStatusHint         = payload->statusHint;
    _lastScannedDirectories = payload->scannedDirectories;
    _lastScannedFiles       = payload->scannedFiles;
    _lastCandidateFiles     = payload->candidateFiles;
    _lastMatchedEntries     = std::max(_lastMatchedEntries, payload->matchedEntries);
    if (! payload->currentPath.empty())
    {
        _lastCurrentPath = payload->currentPath;
    }
    else if (_lastCurrentPath.empty())
    {
        _lastCurrentPath = _lastSearchRootPath;
    }

    if (payload->serviceStatus.available)
    {
        _lastServiceStatus.storeState         = payload->serviceStatus.storeState;
        _lastServiceStatus.syncPhase          = payload->serviceStatus.syncPhase;
        _lastServiceStatus.queryExecutionMode = payload->serviceStatus.queryExecutionMode;
        _lastServiceStatus.fallbackReason     = payload->serviceStatus.fallbackReason;
        _lastServiceStatus.completedRoots     = payload->serviceStatus.completedRoots;
        _lastServiceStatus.totalRoots         = payload->serviceStatus.totalRoots;
        _lastServiceStatus.activeRoot         = payload->serviceStatus.activeRoot;
        _lastServiceStatusHr                  = payload->serviceStatus.hr;
        _lastBackendStatusTickMs              = ::GetTickCount64();
        _hasServiceStatus                     = true;
    }

    _lastProgressTickMs = ::GetTickCount64();
    UpdateStatusRefreshTimer();
    RefreshStatusText();
    const uint64_t elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SteadyClock::now() - handlerStartedAt).count());
    TraceFindContextMenuDiagnostics(L"find.ui.progress-handler-exit",
                                    L"hwnd={:#x} drained={} phase={} scannedDirs={} scannedFiles={} matches={} elapsedUs={}",
                                    reinterpret_cast<uintptr_t>(_hWnd.get()),
                                    drainedProgressCount,
                                    static_cast<uint32_t>(_lastPhase),
                                    _lastScannedDirectories,
                                    _lastScannedFiles,
                                    _lastMatchedEntries,
                                    elapsedUs);
}

void FindFilesWindow::ApplyDeferredSetOperation(SearchOperation operation) noexcept
{
    switch (operation)
    {
        case SearchOperation::Find:
        case SearchOperation::Append: break;
        case SearchOperation::Intersect: KeepOnlyKeysInResults(_deferredKeys); break;
        case SearchOperation::Subtract: RemoveKeysFromResults(_deferredKeys); break;
        default: break;
    }

    _deferredKeys.clear();
    RebuildResultsList();
}

void FindFilesWindow::OnSearchComplete(std::unique_ptr<FindSearchCompletePayload> payload) noexcept
{
    // A deferred window close (OnClose hid the window while a search was active) can only be finished on
    // the UI thread here. Run the check on every exit path -- including the null-payload fallback post and
    // the stale-epoch early return -- so a pending close always completes once the session is idle.
    auto completeDeferredClose = wil::scope_exit([this]() noexcept { CompleteDeferredCloseIfReady(); });

    if (! payload)
    {
        return;
    }
    if (payload->epoch != _activeSearchEpoch)
    {
        return;
    }

    const Debug::Perf::Scope completePerf(L"find.ui.complete_handler_ms");
    if (payload->backend != FILESYSTEM_SEARCH_BACKEND_UNKNOWN || _lastBackend == FILESYSTEM_SEARCH_BACKEND_UNKNOWN)
    {
        _lastBackend = payload->backend;
    }
    _lastPhase                   = FILESYSTEM_SEARCH_PHASE_COMPLETED;
    _lastWarningFlags            = payload->warningFlags;
    _lastScannedDirectories      = payload->scannedDirectories;
    _lastScannedFiles            = payload->scannedFiles;
    _lastCandidateFiles          = payload->candidateFiles;
    _lastMatchedEntries          = payload->matchedEntries;
    const bool uiCancelRequested = _cancelRequestedUi;
    const HRESULT effectiveHr    = uiCancelRequested && SUCCEEDED(payload->hr) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : payload->hr;
    _lastStatusHint              = effectiveHr;
    _cancelRequestedUi           = false;

    if (FAILED(effectiveHr) && effectiveHr != HRESULT_FROM_WIN32(ERROR_CANCELLED))
    {
        Debug::Warning(
            L"FindFiles: search completed with failure backend={} hr=0x{:08X} scannedDirs={} scannedFiles={} candidates={} matches={} warnings=0x{:08X}",
            static_cast<uint32_t>(_lastBackend),
            static_cast<unsigned long>(effectiveHr),
            _lastScannedDirectories,
            _lastScannedFiles,
            _lastCandidateFiles,
            _lastMatchedEntries,
            static_cast<unsigned long>(_lastWarningFlags));
    }

    if (_activeOperation == SearchOperation::Intersect || _activeOperation == SearchOperation::Subtract)
    {
        if (effectiveHr == S_OK && ! uiCancelRequested)
        {
            ApplyDeferredSetOperation(_activeOperation);
        }
        else
        {
            _deferredKeys.clear();
        }
    }

    ApplyPendingResultsRefresh();
    if (_activeOperation == SearchOperation::Find)
    {
        _pendingSelectionFullPath.clear();
    }
    _session.NotifyUiSettled();
    UpdateStatusRefreshTimer();
    RefreshStatusText();
    UpdateActionButtons();
    // CompleteDeferredCloseIfReady() runs unconditionally via the scope_exit at the top of this function.
}

void FindFilesWindow::CompleteDeferredCloseIfReady() noexcept
{
    if (! _closeRequested || ! _hWnd || _session.IsActive())
    {
        return;
    }

    EmitPerfCount(L"find.ui.close_deferred_complete_count");
    _hWnd.reset();
}

#ifdef ENABLE_TESTS
void FindFilesWindow::DebugRecordIncrementalResultRefresh() noexcept
{
    if (_session.IsUiSettled() || _results.empty())
    {
        return;
    }

    ++_debugIncrementalResultRefreshCount;
    if (_resultsList)
    {
        const auto metrics = _resultsList->GetVisibleWorkMetrics();
        if (metrics.visibleRowCount != 0u)
        {
            ++_debugIncrementalVisibleResultRefreshCount;
        }
    }
}

bool FindFilesWindow::DebugConfigure(std::wstring rootPath,
                                     std::wstring namePattern,
                                     std::wstring contentPattern,
                                     Common::Settings::SearchNameMode nameMode,
                                     Common::Settings::SearchContentMode contentMode) noexcept
{
    if (! _hWnd)
    {
        return false;
    }

    if (! DebugSetComboText(FindFilesDebugFocusTarget::RootCombo, std::move(rootPath)))
    {
        return false;
    }
    if (! DebugSetComboText(FindFilesDebugFocusTarget::NameCombo, std::move(namePattern)))
    {
        return false;
    }
    if (! DebugSetComboText(FindFilesDebugFocusTarget::ContentCombo, std::move(contentPattern)))
    {
        return false;
    }
    _nameModeCombo->SetSelectedIndex(static_cast<size_t>(ToComboIndex(nameMode)));
    _contentModeCombo->SetSelectedIndex(static_cast<size_t>(ToComboIndex(contentMode)));
    UpdateOptionDependencies();
    PersistUiState(false);
    return true;
}

bool FindFilesWindow::DebugSetOptions(bool recursive, bool includeFiles, bool includeDirectories, bool preferIndex, bool wantSnippets) noexcept
{
    if (! _hWnd)
    {
        return false;
    }

    _recursiveCheck->SetChecked(recursive);
    _includeFilesCheck->SetChecked(includeFiles);
    _includeDirectoriesCheck->SetChecked(includeDirectories);
    _preferIndexCheck->SetChecked(preferIndex);
    _wantSnippetsCheck->SetChecked(wantSnippets);
    UpdateOptionDependencies();
    PersistUiState(false);
    return true;
}

bool FindFilesWindow::DebugSetComboText(FindFilesDebugFocusTarget target, std::wstring text) noexcept
{
    if (! _hWnd)
    {
        return false;
    }

    ComboBox* combo = nullptr;
    switch (target)
    {
        case FindFilesDebugFocusTarget::RootCombo: combo = _rootCombo; break;
        case FindFilesDebugFocusTarget::NameCombo: combo = _nameCombo; break;
        case FindFilesDebugFocusTarget::ContentCombo: combo = _contentCombo; break;
        case FindFilesDebugFocusTarget::ContentModeCombo: return false;
        case FindFilesDebugFocusTarget::None:
        case FindFilesDebugFocusTarget::NameModeCombo:
        case FindFilesDebugFocusTarget::RecursiveCheck:
        case FindFilesDebugFocusTarget::IncludeFilesCheck:
        case FindFilesDebugFocusTarget::IncludeDirectoriesCheck:
        case FindFilesDebugFocusTarget::FollowSymlinksCheck:
        case FindFilesDebugFocusTarget::MatchCaseNameCheck:
        case FindFilesDebugFocusTarget::MatchCaseContentCheck:
        case FindFilesDebugFocusTarget::PreferIndexCheck:
        case FindFilesDebugFocusTarget::WantSnippetsCheck:
        case FindFilesDebugFocusTarget::FindButton:
        case FindFilesDebugFocusTarget::AppendButton:
        case FindFilesDebugFocusTarget::IntersectButton:
        case FindFilesDebugFocusTarget::SubtractButton:
        case FindFilesDebugFocusTarget::CancelButton:
        case FindFilesDebugFocusTarget::OpenButton:
        case FindFilesDebugFocusTarget::ParentButton:
        case FindFilesDebugFocusTarget::HelpButton:
        case FindFilesDebugFocusTarget::ResultsGrid: return false;
        default: return false;
    }

    if (! combo)
    {
        return false;
    }

    const std::wstring expectedText         = text;
    const std::wstring expectedObservedText = StripSingleLineControlCharacters(expectedText);
    _dxHost.CommitFocusedTextInput();
    combo->SetText(std::move(text));
    SetFocus(_hWnd.get());
    _dxHost.SetFocusControl(combo);
    _dxHost.SyncTextInput(combo);
    if (target == FindFilesDebugFocusTarget::RootCombo)
    {
        UpdateOptionDependencies();
        RefreshRootNavigationPath();
        RefreshRootNavigationHistory();
    }
    UpdateActionButtons();
    PersistUiState(false);
    _dxHost.SetFocusControl(combo);
    _dxHost.SyncTextInput(combo);
    _debugLastSetComboTarget        = target;
    _debugLastSetComboRequestedText = expectedText;
    _debugLastSetComboObservedText  = GetComboText(combo);
    if (target == FindFilesDebugFocusTarget::RootCombo)
    {
        _debugRootTextOverride = expectedText;
    }
    else if (target == FindFilesDebugFocusTarget::NameCombo)
    {
        _debugNameTextOverride = expectedText;
    }
    else if (target == FindFilesDebugFocusTarget::ContentCombo)
    {
        _debugContentTextOverride = expectedText;
    }
    return _debugLastSetComboObservedText == expectedObservedText;
}

bool FindFilesWindow::DebugSetDestinationPath(std::wstring path) noexcept
{
    if (! _hWnd)
    {
        return false;
    }

    if (path.empty())
    {
        _explicitDestinationFolder.reset();
    }
    else
    {
        _explicitDestinationFolder = std::filesystem::path(std::move(path));
    }

    RefreshDestinationStatusText();
    return true;
}

bool FindFilesWindow::DebugStartSearch(FindFilesDebugOperation operation) noexcept
{
    _dxHost.CommitFocusedTextInput();

    SearchTextOverride textOverride{
        .rootPath       = GetComboText(_rootCombo),
        .namePattern    = GetComboText(_nameCombo),
        .contentPattern = GetComboText(_contentCombo),
    };
    if (_debugRootTextOverride.has_value())
    {
        textOverride.rootPath = _debugRootTextOverride.value();
    }
    if (_debugNameTextOverride.has_value())
    {
        textOverride.namePattern = _debugNameTextOverride.value();
    }
    if (_debugContentTextOverride.has_value())
    {
        textOverride.contentPattern = _debugContentTextOverride.value();
    }
    if (_debugLastSetComboObservedText == _debugLastSetComboRequestedText)
    {
        if (_debugLastSetComboTarget == FindFilesDebugFocusTarget::RootCombo)
        {
            textOverride.rootPath = _debugLastSetComboObservedText;
        }
        else if (_debugLastSetComboTarget == FindFilesDebugFocusTarget::NameCombo)
        {
            textOverride.namePattern = _debugLastSetComboObservedText;
        }
        else if (_debugLastSetComboTarget == FindFilesDebugFocusTarget::ContentCombo)
        {
            textOverride.contentPattern = _debugLastSetComboObservedText;
        }
    }

    return DebugStartSearchWithTextOverride(operation, std::move(textOverride));
}

bool FindFilesWindow::DebugStartSearchWithTextOverride(FindFilesDebugOperation operation, SearchTextOverride textOverride) noexcept
{
    _lastDebugStartRootPath       = textOverride.rootPath;
    _lastDebugStartNamePattern    = textOverride.namePattern;
    _lastDebugStartContentPattern = textOverride.contentPattern;

    switch (operation)
    {
        case FindFilesDebugOperation::Find: return BeginSearch(SearchOperation::Find, &textOverride);
        case FindFilesDebugOperation::Append: return BeginSearch(SearchOperation::Append, &textOverride);
        case FindFilesDebugOperation::Intersect: return BeginSearch(SearchOperation::Intersect, &textOverride);
        case FindFilesDebugOperation::Subtract: return BeginSearch(SearchOperation::Subtract, &textOverride);
    }

    return false;
}

bool FindFilesWindow::DebugCancelSearch() noexcept
{
    if (! _session.IsActive())
    {
        return true;
    }

    _cancelRequestedUi = true;
    _session.Cancel();
    RefreshStatusText();
    UpdateActionButtons();
    return true;
}

bool FindFilesWindow::DebugPostStaleSearchPayloads(std::wstring fullPath) noexcept
{
    if (! _hWnd)
    {
        return false;
    }

    const uint64_t staleEpoch = _activeSearchEpoch == 0u ? 1u : _activeSearchEpoch - 1u;
    if (staleEpoch == _activeSearchEpoch)
    {
        return false;
    }

    const std::filesystem::path path(fullPath);
    const std::wstring displayName = path.filename().native();
    const std::wstring displayPath = path.parent_path().native();

    auto results  = std::unique_ptr<FindSearchResultsPayload>(new (std::nothrow) FindSearchResultsPayload{});
    auto progress = std::unique_ptr<FindSearchProgressPayload>(new (std::nothrow) FindSearchProgressPayload{});
    auto complete = std::unique_ptr<FindSearchCompletePayload>(new (std::nothrow) FindSearchCompletePayload{});
    if (! results || ! progress || ! complete)
    {
        return false;
    }

    FindResultRecord staleRecord{};
    staleRecord.key            = MakeResultKey(kBuiltinLocalFileSystemId, L"", fullPath);
    staleRecord.pluginId       = std::wstring(kBuiltinLocalFileSystemId);
    staleRecord.pluginShortId  = L"file";
    staleRecord.fullPath       = std::move(fullPath);
    staleRecord.relativePath   = displayName;
    staleRecord.displayPath    = displayPath;
    staleRecord.displayName    = displayName;
    staleRecord.fileAttributes = FILE_ATTRIBUTE_ARCHIVE;
    staleRecord.matchedBy      = FILESYSTEM_SEARCH_MATCH_SOURCE_NAME;
    results->results.push_back(std::move(staleRecord));
    results->epoch      = staleEpoch;
    results->enqueuedAt = SteadyClock::now();

    progress->epoch              = staleEpoch;
    progress->phase              = FILESYSTEM_SEARCH_PHASE_CONTENT_SCAN;
    progress->backend            = FILESYSTEM_SEARCH_BACKEND_SCAN;
    progress->warningFlags       = FILESYSTEM_SEARCH_WARNING_OVERFLOW;
    progress->statusHint         = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    progress->scannedDirectories = 9001u;
    progress->scannedFiles       = 9002u;
    progress->candidateFiles     = 9003u;
    progress->matchedEntries     = 9004u;
    progress->currentPath        = displayPath;
    progress->enqueuedAt         = SteadyClock::now();

    complete->epoch              = staleEpoch;
    complete->hr                 = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    complete->backend            = FILESYSTEM_SEARCH_BACKEND_SCAN;
    complete->warningFlags       = FILESYSTEM_SEARCH_WARNING_OVERFLOW;
    complete->scannedDirectories = 9001u;
    complete->scannedFiles       = 9002u;
    complete->candidateFiles     = 9003u;
    complete->matchedEntries     = 9004u;

    const WPARAM operationKey = static_cast<WPARAM>(staleEpoch);
    return PostMessagePayload(_hWnd.get(), WndMsg::kFindSearchResults, operationKey, std::move(results)) &&
           PostMessagePayload(_hWnd.get(), WndMsg::kFindSearchProgress, operationKey, std::move(progress)) &&
           PostMessagePayload(_hWnd.get(), WndMsg::kFindSearchComplete, operationKey, std::move(complete));
}

FindFilesDebugFocusTarget FindFilesWindow::ResolveDebugFocusTarget() const noexcept
{
    const RedSalamander::DxUi::Control* const focused = _dxHost.GetFocusControl();
    const HWND focusedHwnd                            = GetFocus();
    const HWND rootNavigationHwnd                     = _rootNavigation.GetHwnd();
    if (rootNavigationHwnd && focusedHwnd && (focusedHwnd == rootNavigationHwnd || IsChild(rootNavigationHwnd, focusedHwnd) != FALSE))
    {
        return FindFilesDebugFocusTarget::RootCombo;
    }
    if (focused == _rootCombo)
    {
        return FindFilesDebugFocusTarget::RootCombo;
    }
    if (focused == _nameCombo)
    {
        return FindFilesDebugFocusTarget::NameCombo;
    }
    if (focused == _nameModeCombo)
    {
        return FindFilesDebugFocusTarget::NameModeCombo;
    }
    if (focused == _contentCombo)
    {
        return FindFilesDebugFocusTarget::ContentCombo;
    }
    if (focused == _contentModeCombo)
    {
        return FindFilesDebugFocusTarget::ContentModeCombo;
    }
    if (focused == _recursiveCheck)
    {
        return FindFilesDebugFocusTarget::RecursiveCheck;
    }
    if (focused == _includeFilesCheck)
    {
        return FindFilesDebugFocusTarget::IncludeFilesCheck;
    }
    if (focused == _includeDirectoriesCheck)
    {
        return FindFilesDebugFocusTarget::IncludeDirectoriesCheck;
    }
    if (focused == _followSymlinksCheck)
    {
        return FindFilesDebugFocusTarget::FollowSymlinksCheck;
    }
    if (focused == _matchCaseNameCheck)
    {
        return FindFilesDebugFocusTarget::MatchCaseNameCheck;
    }
    if (focused == _matchCaseContentCheck)
    {
        return FindFilesDebugFocusTarget::MatchCaseContentCheck;
    }
    if (focused == _preferIndexCheck)
    {
        return FindFilesDebugFocusTarget::PreferIndexCheck;
    }
    if (focused == _wantSnippetsCheck)
    {
        return FindFilesDebugFocusTarget::WantSnippetsCheck;
    }
    if (focused == _findButton)
    {
        return FindFilesDebugFocusTarget::FindButton;
    }
    if (focused == _appendButton)
    {
        return FindFilesDebugFocusTarget::AppendButton;
    }
    if (focused == _intersectButton)
    {
        return FindFilesDebugFocusTarget::IntersectButton;
    }
    if (focused == _subtractButton)
    {
        return FindFilesDebugFocusTarget::SubtractButton;
    }
    if (focused == _cancelButton)
    {
        return FindFilesDebugFocusTarget::CancelButton;
    }
    if (focused == _openButton)
    {
        return FindFilesDebugFocusTarget::OpenButton;
    }
    if (focused == _parentButton)
    {
        return FindFilesDebugFocusTarget::ParentButton;
    }
    if (focused == _helpButton)
    {
        return FindFilesDebugFocusTarget::HelpButton;
    }
    if (focused == _resultsList)
    {
        return FindFilesDebugFocusTarget::ResultsGrid;
    }
    return FindFilesDebugFocusTarget::None;
}

bool FindFilesWindow::IsComboPopupOpen(const ComboBox* combo) noexcept
{
    if (! combo)
    {
        return false;
    }

    const D2D1_RECT_F bounds    = combo->GetBounds();
    const D2D1_RECT_F hitBounds = combo->GetHitBounds();
    return hitBounds.top < bounds.top || hitBounds.bottom > bounds.bottom || hitBounds.left < bounds.left || hitBounds.right > bounds.right;
}

bool FindFilesWindow::DebugGetSnapshot(FindFilesDebugSnapshot& out) noexcept
{
    out                                       = {};
    out.searchActive                          = _session.IsActive();
    out.usesDxUiHost                          = _dxHost.GetHwnd() == _hWnd.get();
    out.findButtonEnabled                     = _findButton ? _findButton->IsEnabled() : false;
    out.findButtonPressed                     = _findButton ? _findButton->DebugIsPressed() : false;
    out.appendButtonEnabled                   = _appendButton ? _appendButton->IsEnabled() : false;
    out.intersectButtonEnabled                = _intersectButton ? _intersectButton->IsEnabled() : false;
    out.subtractButtonEnabled                 = _subtractButton ? _subtractButton->IsEnabled() : false;
    out.cancelButtonEnabled                   = _cancelButton ? _cancelButton->IsEnabled() : false;
    out.openButtonEnabled                     = _openButton ? _openButton->IsEnabled() : false;
    out.parentButtonEnabled                   = _parentButton ? _parentButton->IsEnabled() : false;
    out.helpButtonEnabled                     = _helpButton ? _helpButton->IsEnabled() : false;
    out.rootComboEnabled                      = _rootCombo ? _rootCombo->IsEnabled() : false;
    out.nameComboEnabled                      = _nameCombo ? _nameCombo->IsEnabled() : false;
    out.nameModeComboEnabled                  = _nameModeCombo ? _nameModeCombo->IsEnabled() : false;
    out.contentComboEnabled                   = _contentCombo ? _contentCombo->IsEnabled() : false;
    out.contentModeComboEnabled               = _contentModeCombo ? _contentModeCombo->IsEnabled() : false;
    out.matchCaseContentEnabled               = _matchCaseContentCheck ? _matchCaseContentCheck->IsEnabled() : false;
    out.preferIndexEnabled                    = _preferIndexCheck ? _preferIndexCheck->IsEnabled() : false;
    out.preferIndexChecked                    = _preferIndexCheck ? _preferIndexCheck->IsChecked() : false;
    out.wantSnippetsEnabled                   = _wantSnippetsCheck ? _wantSnippetsCheck->IsEnabled() : false;
    out.recursiveChecked                      = _recursiveCheck ? _recursiveCheck->IsChecked() : false;
    out.resultCount                           = _results.size();
    out.selectedResultCount                   = _resultsList ? _resultsList->GetSelectionModel().GetCount() : 0u;
    out.visibleChildWindowCount               = 0u;
    out.hasStatusStrip                        = _statusText != nullptr;
    out.statusStripVisible                    = _statusText ? _statusText->IsVisible() : false;
    out.statusStripSectionCount               = _statusText ? static_cast<uint32_t>(_statusText->GetSectionCount()) : 0u;
    out.statusStripBlendsWithWindowBackground = _statusText ? _statusText->GetBlendWithWindowBackground() : false;
    const HWND focusedWindow                  = GetFocus();
    if (_statusText)
    {
        const D2D1_RECT_F statusBounds = _statusText->GetBounds();
        out.statusStripHeightDip       = (statusBounds.bottom > statusBounds.top) ? (statusBounds.bottom - statusBounds.top) : 0.0f;
    }
    if (const HWND rootNavigationHwnd = _rootNavigation.GetHwnd(); rootNavigationHwnd && IsWindow(rootNavigationHwnd) != FALSE)
    {
        out.rootNavigationHwnd          = rootNavigationHwnd;
        out.rootNavigationVisible       = IsWindowVisible(rootNavigationHwnd) != FALSE;
        out.rootNavigationHasWin32Focus = focusedWindow && (focusedWindow == rootNavigationHwnd || IsChild(rootNavigationHwnd, focusedWindow) != FALSE);
        if (const auto path = _rootNavigation.GetPath(); path.has_value())
        {
            out.rootNavigationText = path->wstring();
        }

        RECT rootNavigationRectPx{};
        if (_hWnd && GetWindowRect(rootNavigationHwnd, &rootNavigationRectPx) != FALSE)
        {
            static_cast<void>(MapWindowPoints(nullptr, _hWnd.get(), reinterpret_cast<POINT*>(&rootNavigationRectPx), 2));
            out.rootNavigationRect = D2D1::RectF(_dxHost.PixelsToDip(static_cast<float>(rootNavigationRectPx.left)),
                                                 _dxHost.PixelsToDip(static_cast<float>(rootNavigationRectPx.top)),
                                                 _dxHost.PixelsToDip(static_cast<float>(rootNavigationRectPx.right)),
                                                 _dxHost.PixelsToDip(static_cast<float>(rootNavigationRectPx.bottom)));
        }

#ifdef ENABLE_TESTS
        NavigationViewDebugSnapshot rootNavigationSnapshot{};
        if (_rootNavigation.DebugGetSnapshot(rootNavigationSnapshot))
        {
            out.rootNavigationEmbedded      = rootNavigationSnapshot.embeddedDestinationMode;
            out.rootNavigationEditMode      = rootNavigationSnapshot.editMode;
            out.rootNavigationHistoryCount  = rootNavigationSnapshot.historyCount;
            out.rootNavigationEditHostHwnd  = rootNavigationSnapshot.currentEditHostHwnd;
            out.rootNavigationEditInputHwnd = rootNavigationSnapshot.currentEditInputHwnd;
            if (! rootNavigationSnapshot.currentEditText.empty())
            {
                out.rootNavigationText = rootNavigationSnapshot.currentEditText;
            }
            else if (! rootNavigationSnapshot.currentPathText.empty())
            {
                out.rootNavigationText = rootNavigationSnapshot.currentPathText;
            }
        }
#endif
    }
    if (const HWND destinationHwnd = _destinationNavigation.GetHwnd(); destinationHwnd && IsWindow(destinationHwnd) != FALSE)
    {
        out.destinationNavigationVisible = IsWindowVisible(destinationHwnd) != FALSE;
        if (const auto path = _destinationNavigation.GetPath(); path.has_value())
        {
            out.destinationNavigationText = path->wstring();
        }

        RECT destinationRectPx{};
        if (_hWnd && GetWindowRect(destinationHwnd, &destinationRectPx) != FALSE)
        {
            static_cast<void>(MapWindowPoints(nullptr, _hWnd.get(), reinterpret_cast<POINT*>(&destinationRectPx), 2));
            out.destinationNavigationRect = D2D1::RectF(_dxHost.PixelsToDip(static_cast<float>(destinationRectPx.left)),
                                                        _dxHost.PixelsToDip(static_cast<float>(destinationRectPx.top)),
                                                        _dxHost.PixelsToDip(static_cast<float>(destinationRectPx.right)),
                                                        _dxHost.PixelsToDip(static_cast<float>(destinationRectPx.bottom)));
        }

#ifdef ENABLE_TESTS
        NavigationViewDebugSnapshot destinationSnapshot{};
        if (_destinationNavigation.DebugGetSnapshot(destinationSnapshot))
        {
            out.destinationNavigationEmbedded                 = destinationSnapshot.embeddedDestinationMode;
            out.destinationNavigationEditMode                 = destinationSnapshot.editMode;
            out.destinationNavigationMenuHovered              = destinationSnapshot.menuButtonHovered;
            out.destinationNavigationHistoryHovered           = destinationSnapshot.historyButtonHovered;
            out.destinationNavigationDiskHovered              = destinationSnapshot.diskInfoHovered;
            out.destinationNavigationHoveredSegmentIndex      = destinationSnapshot.hoveredSegmentIndex;
            out.destinationNavigationHoveredSeparatorIndex    = destinationSnapshot.hoveredSeparatorIndex;
            out.destinationNavigationHistoryCount             = destinationSnapshot.historyCount;
            out.destinationNavigationHistoryDropdownVisible   = destinationSnapshot.historyDropdownVisible;
            out.destinationNavigationHistoryDropdownOpenCount = destinationSnapshot.historyDropdownOpenCount;
            if (destinationSnapshot.historyRegionRect.right > destinationSnapshot.historyRegionRect.left &&
                destinationSnapshot.historyRegionRect.bottom > destinationSnapshot.historyRegionRect.top)
            {
                std::array<POINT, 2> mappedPoints{{
                    POINT{destinationSnapshot.historyRegionRect.left, destinationSnapshot.historyRegionRect.top},
                    POINT{destinationSnapshot.historyRegionRect.right, destinationSnapshot.historyRegionRect.bottom},
                }};
                if (MapWindowPoints(destinationHwnd, _hWnd.get(), mappedPoints.data(), static_cast<UINT>(mappedPoints.size())) != 0)
                {
                    out.destinationNavigationHistoryRect = RECT{mappedPoints[0].x, mappedPoints[0].y, mappedPoints[1].x, mappedPoints[1].y};
                }
            }
            if (! destinationSnapshot.currentEditText.empty())
            {
                out.destinationNavigationText = destinationSnapshot.currentEditText;
            }
            else if (! destinationSnapshot.currentPathText.empty())
            {
                out.destinationNavigationText = destinationSnapshot.currentPathText;
            }
        }
#endif
    }
    out.resultColumnIds.clear();
    out.resultColumnWidthsDip.clear();
    out.rootPopupOpen        = IsComboPopupOpen(_rootCombo);
    out.nameModePopupOpen    = IsComboPopupOpen(_nameModeCombo);
    out.contentModePopupOpen = IsComboPopupOpen(_contentModeCombo);
    if (_hWnd)
    {
        out.hasWin32Focus = focusedWindow == _hWnd.get() || (focusedWindow && IsChild(_hWnd.get(), focusedWindow) != FALSE);

        const HWND foreground     = GetForegroundWindow();
        const HWND foregroundRoot = foreground ? GetAncestor(foreground, GA_ROOT) : nullptr;
        const HWND findRoot       = GetAncestor(_hWnd.get(), GA_ROOT);
        out.isForegroundWindow    = foregroundRoot && findRoot && foregroundRoot == findRoot;
    }
    out.nameModeSelectedIndex          = _nameModeCombo ? _nameModeCombo->GetSelectedIndex() : std::nullopt;
    out.contentModeSelectedIndex       = _contentModeCombo ? _contentModeCombo->GetSelectedIndex() : std::nullopt;
    out.focusTarget                    = ResolveDebugFocusTarget();
    out.rootText                       = GetComboText(_rootCombo);
    out.namePatternText                = GetComboText(_nameCombo);
    out.contentPatternText             = GetComboText(_contentCombo);
    out.beginRootText                  = _lastBeginRootPath;
    out.beginNamePatternText           = _lastBeginNamePattern;
    out.beginContentPatternText        = _lastBeginContentPattern;
    out.debugStartRootText             = _lastDebugStartRootPath;
    out.debugStartNamePatternText      = _lastDebugStartNamePattern;
    out.debugStartContentPatternText   = _lastDebugStartContentPattern;
    out.debugLastSetComboTarget        = _debugLastSetComboTarget;
    out.debugLastSetComboRequestedText = _debugLastSetComboRequestedText;
    out.debugLastSetComboObservedText  = _debugLastSetComboObservedText;
    out.builtRootText                  = _lastBuiltRootPath;
    out.builtNamePatternText           = _lastBuiltNamePattern;
    out.builtContentPatternText        = _lastBuiltContentPattern;
    out.submittedRootText              = _lastSubmittedRootPath;
    out.submittedNamePatternText       = _lastSubmittedNamePattern;
    out.submittedContentPatternText    = _lastSubmittedContentPattern;
    out.lastStatusHint                 = _lastStatusHint;
    out.warningFlags                   = _lastWarningFlags;
    out.backend                        = static_cast<uint32_t>(_lastBackend);
    out.phase                          = static_cast<uint32_t>(_lastPhase);
    out.hasServiceStatus               = _hasServiceStatus;
    out.themeCompactMode               = _theme.compactMode;
    out.themeDark                      = _theme.dark;
    out.themeHighContrast              = _theme.highContrast;
    out.themeRainbow                   = _theme.menu.rainbowMode;
#ifdef ENABLE_TESTS
    out.resultListFullRebuildCount                = _debugResultListFullRebuildCount;
    out.incrementalResultRefreshCount             = _debugIncrementalResultRefreshCount;
    out.incrementalVisibleResultRefreshCount      = _debugIncrementalVisibleResultRefreshCount;
    out.debugResultActionFocusRestoreRequestCount = _debugResultActionFocusRestoreRequestCount;
    out.dxRenderCount                             = _dxHost.DebugGetRenderCount();
    out.resultGridPaintCount                      = _resultsList ? _resultsList->DebugGetPaintCount() : 0u;
    out.dxResizeCount                             = _dxHost.DebugGetResizeCount();
    out.dxResizeFailureCount                      = _dxHost.DebugGetResizeFailureCount();
    out.debugResizeBeforeWidthDip                 = _debugResizeBeforeWidthDip;
    out.debugResizeTargetWidthDip                 = _debugResizeTargetWidthDip;
    out.debugResizeObservedWidthDip               = _debugResizeObservedWidthDip;
    out.debugResizeSucceeded                      = _debugResizeSucceeded;
#else
    out.resultListFullRebuildCount = 0;
#endif
    out.statusText            = _status;
    out.destinationStatusText = _destinationStatus;
    out.backendStatusText     = BuildBackendStatusText();
    if (_resultsList)
    {
        const auto metrics                 = _resultsList->GetVisibleWorkMetrics();
        out.visibleResultRowCount          = metrics.visibleRowCount;
        out.visibleResultColumnCount       = metrics.visibleColumnCount;
        out.visibleResultCellCount         = metrics.visibleCellCount;
        out.visibleResultIconCellCount     = metrics.visibleIconCellCount;
        out.resultListHasVerticalScrollbar = metrics.hasVerticalScrollbar;
        out.resultsGridFolderViewMode      = _resultsList->GetVisualMode() == GridVisualMode::FolderView;
        const auto layout                  = _resultsList->CaptureColumnLayout();
        out.resultColumnIds.reserve(layout.size());
        out.resultColumnWidthsDip.reserve(layout.size());
        for (const auto& entry : layout)
        {
            out.resultColumnIds.push_back(entry.columnId);
            out.resultColumnWidthsDip.push_back(entry.widthDip);
        }
        if (_settings && _settings->search && ! _settings->search->resultsGridLayout.empty())
        {
            out.debugSettingsFirstWidthDip = _settings->search->resultsGridLayout.front().widthDip;
        }
        if (! layout.empty())
        {
            if (const auto firstVisibleModelIndex = FindResultColumnIndexById(layout[0].columnId); firstVisibleModelIndex.has_value())
            {
                out.firstResultHeaderRect = _resultsList->GetVisibleColumnHeaderRect(firstVisibleModelIndex.value()).value_or(D2D1::RectF());
            }
        }
        if (layout.size() >= 2u)
        {
            if (const auto secondVisibleModelIndex = FindResultColumnIndexById(layout[1].columnId); secondVisibleModelIndex.has_value())
            {
                out.secondResultHeaderRect = _resultsList->GetVisibleColumnHeaderRect(secondVisibleModelIndex.value()).value_or(D2D1::RectF());
            }
        }

        const std::optional<size_t> selectedRow = _resultsList->GetPrimarySelectedRow();
        if (selectedRow.has_value())
        {
            if (selectedRow.value() < _results.size())
            {
                out.selectedResultFullPath = _results[selectedRow.value()].fullPath;
            }
            out.selectedResultRowRect = _resultsList->GetVisibleRowRect(selectedRow.value()).value_or(D2D1::RectF());
            RedSalamander::DxUi::GridDebugRowVisualState rowVisualState{};
            if (_resultsList->DebugGetRowVisualState(_dxHost.GetTheme(), selectedRow.value(), rowVisualState))
            {
                out.selectedResultRowFillArgb    = rowVisualState.fillArgb;
                out.selectedResultRowTextArgb    = rowVisualState.textArgb;
                out.selectedResultRowUsesRainbow = rowVisualState.usesRainbow;
            }
        }
    }
    out.fullPaths.clear();
    out.fullPaths.reserve(_results.size());
    out.resultPathTexts.clear();
    out.resultPathTexts.reserve(_results.size());
    out.resultIconIndices.clear();
    out.resultIconIndices.reserve(_results.size());
    for (const auto& record : _results)
    {
        out.fullPaths.push_back(record.fullPath);
        out.resultPathTexts.push_back(record.displayPath);
        out.resultIconIndices.push_back(record.iconIndex);
    }

    if (_hWnd)
    {
        struct VisibleChildCounter
        {
            size_t count               = 0u;
            HWND rootNavigation        = nullptr;
            HWND destinationNavigation = nullptr;
        } counter{};
        counter.rootNavigation        = _rootNavigation.GetHwnd();
        counter.destinationNavigation = _destinationNavigation.GetHwnd();

        static_cast<void>(::EnumChildWindows(_hWnd.get(),
                                             [](HWND child, LPARAM lParam) noexcept -> BOOL
        {
            auto& counterRef = *reinterpret_cast<VisibleChildCounter*>(lParam);
            if ((counterRef.rootNavigation && (child == counterRef.rootNavigation || IsChild(counterRef.rootNavigation, child) != FALSE)) ||
                (counterRef.destinationNavigation && (child == counterRef.destinationNavigation || IsChild(counterRef.destinationNavigation, child) != FALSE)))
            {
                return TRUE;
            }
            if (::IsWindowVisible(child) != FALSE)
            {
                ++counterRef.count;
            }
            return TRUE;
        },
                                             reinterpret_cast<LPARAM>(&counter)));
        out.visibleChildWindowCount = counter.count;
    }

    return true;
}

bool FindFilesWindow::DebugHitTestResultsGrid(D2D1_POINT_2F pointDip, FindFilesDebugGridHit& out) const noexcept
{
    out = {};
    if (! _resultsList)
    {
        return false;
    }

    RedSalamander::DxUi::Grid::GridDebugHitInfo hit{};
    if (! _resultsList->DebugHitTestPoint(RedSalamander::DxUi::MakePointDip(pointDip), hit))
    {
        return false;
    }

    out.zone           = hit.zone;
    out.columnIndex    = hit.columnIndex;
    out.rectDip        = hit.rectDip;
    out.isHeaderResize = hit.isHeaderResize;
    return true;
}

bool FindFilesWindow::DebugFocusTarget(FindFilesDebugFocusTarget target) noexcept
{
    if (target == FindFilesDebugFocusTarget::RootCombo && _rootNavigation.GetHwnd())
    {
        return FocusRootNavigation(false) && ResolveDebugFocusTarget() == target;
    }

    RedSalamander::DxUi::Control* control = nullptr;
    switch (target)
    {
        case FindFilesDebugFocusTarget::None: control = nullptr; break;
        case FindFilesDebugFocusTarget::RootCombo: control = _rootCombo; break;
        case FindFilesDebugFocusTarget::NameCombo: control = _nameCombo; break;
        case FindFilesDebugFocusTarget::NameModeCombo: control = _nameModeCombo; break;
        case FindFilesDebugFocusTarget::ContentCombo: control = _contentCombo; break;
        case FindFilesDebugFocusTarget::ContentModeCombo: control = _contentModeCombo; break;
        case FindFilesDebugFocusTarget::RecursiveCheck: control = _recursiveCheck; break;
        case FindFilesDebugFocusTarget::IncludeFilesCheck: control = _includeFilesCheck; break;
        case FindFilesDebugFocusTarget::IncludeDirectoriesCheck: control = _includeDirectoriesCheck; break;
        case FindFilesDebugFocusTarget::FollowSymlinksCheck: control = _followSymlinksCheck; break;
        case FindFilesDebugFocusTarget::MatchCaseNameCheck: control = _matchCaseNameCheck; break;
        case FindFilesDebugFocusTarget::MatchCaseContentCheck: control = _matchCaseContentCheck; break;
        case FindFilesDebugFocusTarget::PreferIndexCheck: control = _preferIndexCheck; break;
        case FindFilesDebugFocusTarget::WantSnippetsCheck: control = _wantSnippetsCheck; break;
        case FindFilesDebugFocusTarget::FindButton: control = _findButton; break;
        case FindFilesDebugFocusTarget::AppendButton: control = _appendButton; break;
        case FindFilesDebugFocusTarget::IntersectButton: control = _intersectButton; break;
        case FindFilesDebugFocusTarget::SubtractButton: control = _subtractButton; break;
        case FindFilesDebugFocusTarget::CancelButton: control = _cancelButton; break;
        case FindFilesDebugFocusTarget::OpenButton: control = _openButton; break;
        case FindFilesDebugFocusTarget::ParentButton: control = _parentButton; break;
        case FindFilesDebugFocusTarget::HelpButton: control = _helpButton; break;
        case FindFilesDebugFocusTarget::ResultsGrid: control = _resultsList; break;
    }

    if (target != FindFilesDebugFocusTarget::None && ! control)
    {
        return false;
    }

    if (_hWnd && target != FindFilesDebugFocusTarget::None)
    {
        SetFocus(_hWnd.get());
    }
    _dxHost.SetFocusControl(control);
    return ResolveDebugFocusTarget() == target;
}

bool FindFilesWindow::DebugGetTargetClientRect(FindFilesDebugFocusTarget target, RECT& outRect) const noexcept
{
    outRect = RECT{};
    if (target == FindFilesDebugFocusTarget::RootCombo)
    {
        const HWND rootNavigationHwnd = _rootNavigation.GetHwnd();
        if (! _hWnd || ! rootNavigationHwnd || IsWindow(rootNavigationHwnd) == FALSE || GetWindowRect(rootNavigationHwnd, &outRect) == FALSE)
        {
            outRect = RECT{};
            return false;
        }

        static_cast<void>(MapWindowPoints(nullptr, _hWnd.get(), reinterpret_cast<POINT*>(&outRect), 2));
        return outRect.right > outRect.left && outRect.bottom > outRect.top;
    }

    const RedSalamander::DxUi::Control* control = nullptr;
    switch (target)
    {
        case FindFilesDebugFocusTarget::None: return false;
        case FindFilesDebugFocusTarget::RootCombo: control = _rootCombo; break;
        case FindFilesDebugFocusTarget::NameCombo: control = _nameCombo; break;
        case FindFilesDebugFocusTarget::NameModeCombo: control = _nameModeCombo; break;
        case FindFilesDebugFocusTarget::ContentCombo: control = _contentCombo; break;
        case FindFilesDebugFocusTarget::ContentModeCombo: control = _contentModeCombo; break;
        case FindFilesDebugFocusTarget::RecursiveCheck: control = _recursiveCheck; break;
        case FindFilesDebugFocusTarget::IncludeFilesCheck: control = _includeFilesCheck; break;
        case FindFilesDebugFocusTarget::IncludeDirectoriesCheck: control = _includeDirectoriesCheck; break;
        case FindFilesDebugFocusTarget::FollowSymlinksCheck: control = _followSymlinksCheck; break;
        case FindFilesDebugFocusTarget::MatchCaseNameCheck: control = _matchCaseNameCheck; break;
        case FindFilesDebugFocusTarget::MatchCaseContentCheck: control = _matchCaseContentCheck; break;
        case FindFilesDebugFocusTarget::PreferIndexCheck: control = _preferIndexCheck; break;
        case FindFilesDebugFocusTarget::WantSnippetsCheck: control = _wantSnippetsCheck; break;
        case FindFilesDebugFocusTarget::FindButton: control = _findButton; break;
        case FindFilesDebugFocusTarget::AppendButton: control = _appendButton; break;
        case FindFilesDebugFocusTarget::IntersectButton: control = _intersectButton; break;
        case FindFilesDebugFocusTarget::SubtractButton: control = _subtractButton; break;
        case FindFilesDebugFocusTarget::CancelButton: control = _cancelButton; break;
        case FindFilesDebugFocusTarget::OpenButton: control = _openButton; break;
        case FindFilesDebugFocusTarget::ParentButton: control = _parentButton; break;
        case FindFilesDebugFocusTarget::HelpButton: control = _helpButton; break;
        case FindFilesDebugFocusTarget::ResultsGrid: control = _resultsList; break;
    }

    if (! control)
    {
        return false;
    }

    const D2D1_RECT_F bounds = control->GetBounds();
    outRect.left             = static_cast<LONG>(std::lround(_dxHost.DipsToPixels(bounds.left)));
    outRect.top              = static_cast<LONG>(std::lround(_dxHost.DipsToPixels(bounds.top)));
    outRect.right            = static_cast<LONG>(std::lround(_dxHost.DipsToPixels(bounds.right)));
    outRect.bottom           = static_cast<LONG>(std::lround(_dxHost.DipsToPixels(bounds.bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool FindFilesWindow::DebugApplyResultsLayoutFromSettings() noexcept
{
    if (! _resultsList)
    {
        return false;
    }

    ApplyResultsGridLayoutFromSettings();
    _dxHost.Invalidate();
    return true;
}

bool FindFilesWindow::DebugSetResultSort(size_t columnIndex, bool descending) noexcept
{
    if (! _resultsList || columnIndex >= _resultsModel.GetColumnCount())
    {
        return false;
    }

    GridSortSpec sortSpec{};
    sortSpec.columnIndex = columnIndex;
    sortSpec.direction   = descending ? SortDirection::Descending : SortDirection::Ascending;
    OnGridSortRequested(sortSpec);
    return _resultSortSpec.columnIndex == sortSpec.columnIndex && _resultSortSpec.direction == sortSpec.direction;
}

bool FindFilesWindow::DebugReorderVisibleResultColumn(size_t fromVisibleIndex, size_t targetVisibleIndex) noexcept
{
    if (! _resultsList || fromVisibleIndex == targetVisibleIndex)
    {
        return false;
    }

    const auto beforeLayout = _resultsList->CaptureColumnLayout();
    if (fromVisibleIndex >= beforeLayout.size() || targetVisibleIndex >= beforeLayout.size())
    {
        return false;
    }

    auto targetLayout                                     = beforeLayout;
    const auto movedIt                                    = targetLayout.begin() + static_cast<std::ptrdiff_t>(fromVisibleIndex);
    RedSalamander::DxUi::GridColumnLayoutEntry movedEntry = *movedIt;
    targetLayout.erase(movedIt);
    targetLayout.insert(targetLayout.begin() + static_cast<std::ptrdiff_t>(targetVisibleIndex), movedEntry);
    for (size_t index = 0; index < targetLayout.size(); ++index)
    {
        targetLayout[index].displayIndex = index;
    }

    const std::wstring movedColumnId = targetLayout[targetVisibleIndex].columnId;
    _resultsList->ApplyColumnLayout(targetLayout);
    _dxHost.Invalidate();

    using namespace std::chrono_literals;
    const auto deadline = std::chrono::steady_clock::now() + 300ms;
    while (std::chrono::steady_clock::now() < deadline)
    {
        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        const auto currentLayout = _resultsList->CaptureColumnLayout();
        if (targetVisibleIndex < currentLayout.size() && currentLayout[targetVisibleIndex].columnId == movedColumnId)
        {
            std::this_thread::sleep_for(10ms);
            continue;
        }

        _resultsList->ApplyColumnLayout(targetLayout);
        _dxHost.Invalidate();
        std::this_thread::sleep_for(10ms);
    }

    const auto finalLayout = _resultsList->CaptureColumnLayout();
    if (_settings && _settings->search && targetVisibleIndex < finalLayout.size() && finalLayout[targetVisibleIndex].columnId == movedColumnId)
    {
        _settings->search->resultsGridLayout = ConvertColumnLayout(finalLayout);
    }
    return targetVisibleIndex < finalLayout.size() && finalLayout[targetVisibleIndex].columnId == movedColumnId;
}

bool FindFilesWindow::DebugResizeVisibleResultColumn(size_t visibleIndex, float deltaDip) noexcept
{
    if (! _resultsList || ! std::isfinite(deltaDip) || deltaDip == 0.0f)
    {
        return false;
    }

    const auto beforeLayout = _resultsList->CaptureColumnLayout();
    if (visibleIndex >= beforeLayout.size())
    {
        return false;
    }

    const float beforeWidthDip  = beforeLayout[visibleIndex].widthDip;
    const std::wstring columnId = beforeLayout[visibleIndex].columnId;
    const std::vector<uint64_t> previousSelection(_resultsList->GetSelectionModel().GetOrderedSelection().begin(),
                                                  _resultsList->GetSelectionModel().GetOrderedSelection().end());
    auto targetLayout                   = beforeLayout;
    targetLayout[visibleIndex].widthDip = std::max(32.0f, beforeWidthDip + deltaDip);
#ifdef ENABLE_TESTS
    _debugResizeBeforeWidthDip   = beforeWidthDip;
    _debugResizeTargetWidthDip   = targetLayout[visibleIndex].widthDip;
    _debugResizeObservedWidthDip = beforeWidthDip;
    _debugResizeSucceeded        = false;
#endif
    _resultsModel.ApplyColumnLayoutDefaults(targetLayout);
    _resultsList->SetModel(&_resultsModel);
    _resultsList->ApplyColumnLayout(targetLayout);
    if (previousSelection.size() == 1u)
    {
        _resultsList->GetSelectionModel().SetSingle(previousSelection.front());
    }
    _dxHost.Invalidate();

    using namespace std::chrono_literals;
    const auto deadline = std::chrono::steady_clock::now() + 300ms;
    while (std::chrono::steady_clock::now() < deadline)
    {
        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        const auto currentLayout = _resultsList->CaptureColumnLayout();
        const auto it = std::find_if(currentLayout.begin(), currentLayout.end(), [&](const RedSalamander::DxUi::GridColumnLayoutEntry& entry) noexcept {
            return entry.columnId == columnId;
        });
        if (it != currentLayout.end())
        {
#ifdef ENABLE_TESTS
            _debugResizeObservedWidthDip = it->widthDip;
#endif
            if (deltaDip > 0.0f ? it->widthDip >= beforeWidthDip + 20.0f : it->widthDip <= beforeWidthDip - 20.0f)
            {
                std::this_thread::sleep_for(10ms);
                continue;
            }
        }

        _resultsList->ApplyColumnLayout(targetLayout);
        _dxHost.Invalidate();
        std::this_thread::sleep_for(10ms);
    }

    const auto finalLayout = _resultsList->CaptureColumnLayout();
    const auto it          = std::find_if(
        finalLayout.begin(), finalLayout.end(), [&](const RedSalamander::DxUi::GridColumnLayoutEntry& entry) noexcept { return entry.columnId == columnId; });
    if (it == finalLayout.end())
    {
        return false;
    }

#ifdef ENABLE_TESTS
    _debugResizeObservedWidthDip = it->widthDip;
    _debugResizeSucceeded        = deltaDip > 0.0f ? it->widthDip >= beforeWidthDip + 20.0f : it->widthDip <= beforeWidthDip - 20.0f;
#endif
    if (_settings && _settings->search && (deltaDip > 0.0f ? it->widthDip >= beforeWidthDip + 20.0f : it->widthDip <= beforeWidthDip - 20.0f))
    {
        _resultsModel.ApplyColumnLayoutDefaults(finalLayout);
        _settings->search->resultsGridLayout = ConvertColumnLayout(finalLayout);
    }
    return deltaDip > 0.0f ? it->widthDip >= beforeWidthDip + 20.0f : it->widthDip <= beforeWidthDip - 20.0f;
}

bool FindFilesWindow::DebugSelectResult(std::wstring fullPath) noexcept
{
    if (! _resultsList || fullPath.empty())
    {
        return false;
    }

    const auto findMatchingResult = [&]() noexcept -> std::vector<FindResultRecord>::iterator
    {
        const std::wstring requestedLeaf       = std::filesystem::path(fullPath).filename().native();
        const std::wstring requestedParentLeaf = std::filesystem::path(fullPath).parent_path().filename().native();
        auto exactMatch                        = std::find_if(_results.begin(), _results.end(), [&](const FindResultRecord& record) noexcept {
            return OrdinalString::EqualsNoCasePath(std::filesystem::path(record.fullPath), std::filesystem::path(fullPath));
        });
        if (exactMatch != _results.end())
        {
            return exactMatch;
        }

        auto leafMatch = _results.end();
        if (! requestedLeaf.empty())
        {
            const std::wstring requestedLeafLower       = ToLowerCopy(requestedLeaf);
            const std::wstring requestedParentLeafLower = ToLowerCopy(requestedParentLeaf);
            for (auto candidate = _results.begin(); candidate != _results.end(); ++candidate)
            {
                const std::wstring candidateLeaf = std::filesystem::path(candidate->fullPath).filename().native();
                bool matchesRequestedLeaf        = OrdinalString::EqualsNoCase(candidateLeaf, requestedLeaf);
                if (! matchesRequestedLeaf)
                {
                    matchesRequestedLeaf = OrdinalString::EqualsNoCase(candidate->displayName, requestedLeaf);
                }
                if (! matchesRequestedLeaf)
                {
                    continue;
                }

                if (! requestedParentLeafLower.empty())
                {
                    const std::wstring candidateRelativeLower = ToLowerCopy(candidate->relativePath);
                    if (candidateRelativeLower.find(requestedParentLeafLower) == std::wstring::npos)
                    {
                        continue;
                    }
                }

                if (leafMatch != _results.end())
                {
                    return _results.end();
                }

                leafMatch = candidate;
            }
        }

        return leafMatch;
    };

    using namespace std::chrono_literals;
    const auto deadline = std::chrono::steady_clock::now() + 500ms;
    auto it             = _results.end();
    while (std::chrono::steady_clock::now() < deadline)
    {
        it = findMatchingResult();
        if (it != _results.end())
        {
            break;
        }

        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        std::this_thread::sleep_for(10ms);
    }

    if (it == _results.end())
    {
        return false;
    }

    const size_t rowIndex = static_cast<size_t>(std::distance(_results.begin(), it));
    _resultsList->GetSelectionModel().SetSingle(_resultsModel.GetStableRowId(rowIndex));
    _dxHost.SetFocusControl(_resultsList);
    _resultsList->NotifyDataChanged();
    UpdateActionButtons();
    _dxHost.Invalidate();
    return true;
}

bool FindFilesWindow::DebugSelectResults(std::vector<std::wstring> fullPaths) noexcept
{
    if (! _resultsList || fullPaths.empty())
    {
        return false;
    }

    std::vector<uint64_t> stableIds;
    stableIds.reserve(fullPaths.size());
    for (const std::wstring& fullPath : fullPaths)
    {
        if (fullPath.empty())
        {
            return false;
        }

        const auto it = std::find_if(_results.begin(), _results.end(), [&](const FindResultRecord& record) noexcept {
            return OrdinalString::EqualsNoCasePath(std::filesystem::path(record.fullPath), std::filesystem::path(fullPath));
        });
        if (it == _results.end())
        {
            return false;
        }

        const size_t rowIndex = static_cast<size_t>(std::distance(_results.begin(), it));
        stableIds.push_back(_resultsModel.GetStableRowId(rowIndex));
    }

    auto& selection = _resultsList->GetSelectionModel();
    selection.Clear();
    for (const uint64_t stableId : stableIds)
    {
        selection.Toggle(stableId);
    }

    _dxHost.SetFocusControl(_resultsList);
    _resultsList->NotifyDataChanged();
    UpdateActionButtons();
    _dxHost.Invalidate();
    return selection.GetCount() == fullPaths.size();
}

bool FindFilesWindow::DebugActivateSelectedResult() noexcept
{
    const auto index = GetSelectedResultIndex();
    if (! index.has_value())
    {
        return false;
    }

    OnGridRowActivated(index.value());
    return true;
}

bool FindFilesWindow::DebugOpenSelectedResultParent() noexcept
{
    const auto index = GetSelectedResultIndex();
    if (! index.has_value())
    {
        return false;
    }

    OpenSelectedResult(true);
    return true;
}

bool FindFilesWindow::DebugGetSelectedOpenDisposition(bool parentOnly, FindFilesDebugOpenDisposition& out) const noexcept
{
    out              = FindFilesDebugOpenDisposition::None;
    const auto index = GetSelectedResultIndex();
    if (! index.has_value() || index.value() >= _results.size())
    {
        return false;
    }

    const FindResultRecord& record = _results[index.value()];
    switch (BuildOpenPlan(record, parentOnly).disposition)
    {
        case ResultOpenDisposition::None: out = FindFilesDebugOpenDisposition::None; break;
        case ResultOpenDisposition::NavigateToResult: out = FindFilesDebugOpenDisposition::NavigateToResult; break;
        case ResultOpenDisposition::NavigateToParent: out = FindFilesDebugOpenDisposition::NavigateToParent; break;
        case ResultOpenDisposition::NavigateToParentAndOpen: out = FindFilesDebugOpenDisposition::NavigateToParentAndOpen; break;
        case ResultOpenDisposition::DefaultOpenFile: out = FindFilesDebugOpenDisposition::DefaultOpenFile; break;
    }
    return true;
}

bool FindFilesWindow::DebugScrollResultsByWheelDetents(int detents) noexcept
{
    if (! _resultsList || detents == 0)
    {
        return detents == 0;
    }

    _dxHost.SetFocusControl(_resultsList);
    const float wheelDelta = detents > 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
    const int stepCount    = detents > 0 ? detents : -detents;
    for (int remaining = stepCount; remaining > 0; --remaining)
    {
        if (! _resultsList->OnMouseWheel(_dxHost, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0))
        {
            return false;
        }
    }

    return true;
}

bool FindFilesWindow::DebugWaitForIdle(uint32_t timeoutMs) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(timeoutMs);
    while (std::chrono::steady_clock::now() < deadline)
    {
        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        const bool workerIdle  = _session.WaitForIdle(0u);
        const bool refreshIdle = ! _resultsRefreshPending && ! _resultsRefreshTimerArmed;
        if (workerIdle && _session.IsUiSettled() && refreshIdle)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return _session.WaitForIdle(0u) && _session.IsUiSettled() && ! _resultsRefreshPending && ! _resultsRefreshTimerArmed;
}
#endif

void FindFilesWindow::TraceRawWindowMessage(UINT message, WPARAM wParam, LPARAM lParam, std::wstring_view phase, bool dxHandled) const noexcept
{
    if (! IsContextMenuDiagnosticsEnabled() || ! ShouldTraceFindWindowMessage(message))
    {
        return;
    }

    const HWND hwnd = _hWnd.get();

    POINT screenPt{};
    bool haveScreenPt = false;
    if (message >= WM_NCMOUSEMOVE && message <= WM_NCXBUTTONDBLCLK)
    {
        screenPt     = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        haveScreenPt = true;
    }
    else if (message >= WM_MOUSEFIRST && message <= WM_MOUSELAST)
    {
        POINT clientPt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        screenPt = clientPt;
        if (hwnd && ClientToScreen(hwnd, &screenPt) != FALSE)
        {
            haveScreenPt = true;
        }
    }
    else if (GetCursorPos(&screenPt) != FALSE) // getcursorpos-allow: diagnostic-only
    {
        haveScreenPt = true;
    }

    POINT clientPt    = screenPt;
    bool haveClientPt = haveScreenPt && hwnd && ScreenToClient(hwnd, &clientPt) != FALSE;

    RECT clientRect{};
    const bool haveClientRect = hwnd && GetClientRect(hwnd, &clientRect) != FALSE;
    const bool inClient       = haveClientPt && haveClientRect && PtInRect(&clientRect, clientPt) != FALSE;
    const HWND windowAtPoint  = (haveScreenPt && ShouldResolveScreenHitWindowForFindTrace(message)) ? WindowFromPoint(screenPt) : nullptr;
    const HWND rootAtPoint    = windowAtPoint ? GetAncestor(windowAtPoint, GA_ROOT) : nullptr;
    const HWND childAtPoint   = (haveClientPt && hwnd) ? ChildWindowFromPointEx(hwnd, clientPt, CWP_SKIPINVISIBLE) : nullptr;
    const HWND destinationNav = _destinationNavigation.GetHwnd();

    TraceFindContextMenuDiagnostics(L"find.wndproc.raw",
                                    L"phase={} hwnd={:#x} message={} msg=0x{:x} wParam={:#x} lParam={:#x} clientPt=({}, {}) haveClient={} "
                                    L"screenPt=({}, {}) haveScreen={} inClient={} windowAtPoint={:#x} rootAtPoint={:#x} childAtPoint={:#x} "
                                    L"destinationNav={:#x} dxHost={:#x} dxHandled={} focus={:#x} active={:#x} foreground={:#x} capture={:#x}",
                                    phase,
                                    reinterpret_cast<uintptr_t>(hwnd),
                                    TraceFindWindowMessageName(message),
                                    message,
                                    static_cast<uintptr_t>(wParam),
                                    static_cast<uintptr_t>(lParam),
                                    clientPt.x,
                                    clientPt.y,
                                    haveClientPt ? 1 : 0,
                                    screenPt.x,
                                    screenPt.y,
                                    haveScreenPt ? 1 : 0,
                                    inClient ? 1 : 0,
                                    reinterpret_cast<uintptr_t>(windowAtPoint),
                                    reinterpret_cast<uintptr_t>(rootAtPoint),
                                    reinterpret_cast<uintptr_t>(childAtPoint),
                                    reinterpret_cast<uintptr_t>(destinationNav),
                                    reinterpret_cast<uintptr_t>(_dxHost.GetHwnd()),
                                    dxHandled ? 1 : 0,
                                    reinterpret_cast<uintptr_t>(GetFocus()),
                                    reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                    reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                    reinterpret_cast<uintptr_t>(GetCapture()));
}

LRESULT FindFilesWindow::WindowProc(UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_KILLFOCUS || (message == WM_ACTIVATE && LOWORD(wParam) == WA_INACTIVE))
    {
        _keyboardModifiers = 0u;
    }
    UpdateKeyboardModifierState(message, wParam);
    TraceRawWindowMessage(message, wParam, lParam, L"enter");
    if (HandleResultShortcut(message, wParam))
    {
        return 0;
    }
    if (HandleRootMnemonic(message, wParam))
    {
        return 0;
    }
    if (HandleRootNavigationTabBridge(message, wParam))
    {
        return 0;
    }
    if (message == WM_COMMAND && CanHandleResultCommands() && HandleResultCommandId(LOWORD(wParam)))
    {
        return 0;
    }

    bool dxHandled         = false;
    const LRESULT dxResult = _dxHost.HandleMessage(_hWnd.get(), message, wParam, lParam, dxHandled);
    TraceRawWindowMessage(message, wParam, lParam, dxHandled ? L"after-dx-handled" : L"after-dx-unhandled", dxHandled);
    if (dxHandled)
    {
        if (message == WM_SIZE)
        {
            Layout();
        }
        return dxResult;
    }

    switch (message)
    {
        case WM_CREATE: return OnCreate(_hWnd.get()) ? 0 : -1;
        case WM_SIZE: Layout(); return 0;
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lParam))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(_hWnd.get(), *info, 760, 480);
            }
            return 0;
        case WM_TIMER:
            if (wParam == kStatusRefreshTimerId)
            {
                OnStatusRefreshTimer();
                return 0;
            }
            if (wParam == kDeferredResultsRefreshTimerId)
            {
                ApplyPendingResultsRefresh();
                return 0;
            }
            break;
        case WM_DPICHANGED:
        {
            if (const RECT* suggested = reinterpret_cast<const RECT*>(lParam))
            {
                SetWindowPos(_hWnd.get(),
                             nullptr,
                             suggested->left,
                             suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            // The hosted NavigationViews are native children that re-render fonts/icons from their
            // own DPI state; forward the change explicitly like FolderWindow does for its panes
            // (the system's WM_DPICHANGED_AFTERPARENT does not reliably refresh them).
            if (_rootNavigation.GetHwnd())
            {
                _rootNavigation.OnDpiChanged(static_cast<float>(HIWORD(wParam)));
            }
            if (_destinationNavigation.GetHwnd())
            {
                _destinationNavigation.OnDpiChanged(static_cast<float>(HIWORD(wParam)));
            }
            ApplyTheme();
            Layout();
            return 0;
        }
        case WM_NCACTIVATE: ApplyTitleBarTheme(_hWnd.get(), _theme, wParam != FALSE); return DefWindowProcW(_hWnd.get(), message, wParam, lParam);
        case WM_CLOSE:
            if (OnClose())
            {
                _hWnd.reset();
            }
            return 0;
        case WndMsg::kFindSearchResults: OnSearchResults(wParam, lParam); return 0;
        case WndMsg::kFindSearchProgress: OnSearchProgress(wParam, lParam); return 0;
        case WndMsg::kFindSearchComplete: OnSearchComplete(TakeMessagePayload<FindSearchCompletePayload>(lParam)); return 0;
        case WndMsg::kFindSearchDeferredRefresh: ApplyPendingResultsRefresh(); return 0;
        case WndMsg::kFindShowActionMenu:
        {
            const std::optional<POINT> screenPoint = _pendingFindActionMenuPoint;
            _pendingFindActionMenuPoint.reset();
            TraceFindContextMenuDiagnostics(L"find.action-menu-posted-message",
                                            L"hwnd={:#x} hasPoint={} point=({}, {})",
                                            reinterpret_cast<uintptr_t>(_hWnd.get()),
                                            screenPoint.has_value() ? L"true" : L"false",
                                            screenPoint.has_value() ? screenPoint->x : 0,
                                            screenPoint.has_value() ? screenPoint->y : 0);
            if (screenPoint.has_value())
            {
                ShowFindActionMenu(screenPoint.value());
            }
            return 0;
        }
#ifdef ENABLE_TESTS
        case kFindFilesWindowDebugMessage:
            switch (static_cast<FindFilesWindowDebugCommand>(wParam))
            {
                case FindFilesWindowDebugCommand::Configure:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugConfigurePayload*>(lParam))
                    {
                        payload->result = DebugConfigure(std::move(payload->rootPath),
                                                         std::move(payload->namePattern),
                                                         std::move(payload->contentPattern),
                                                         payload->nameMode,
                                                         payload->contentMode);
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
                case FindFilesWindowDebugCommand::SetOptions:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugSetOptionsPayload*>(lParam))
                    {
                        payload->result = DebugSetOptions(
                            payload->recursive, payload->includeFiles, payload->includeDirectories, payload->preferIndex, payload->wantSnippets);
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
                case FindFilesWindowDebugCommand::StartSearch:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugStartSearchPayload*>(lParam))
                    {
                        payload->result = DebugStartSearch(payload->operation);
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
                case FindFilesWindowDebugCommand::FocusTarget:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugFocusTargetPayload*>(lParam))
                    {
                        payload->result = DebugFocusTarget(payload->target);
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
                case FindFilesWindowDebugCommand::SetComboText:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugSetComboTextPayload*>(lParam))
                    {
                        payload->result = DebugSetComboText(payload->target, std::move(payload->text));
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
                case FindFilesWindowDebugCommand::GetSnapshot:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugSnapshotPayload*>(lParam))
                    {
                        payload->result = payload->snapshot && DebugGetSnapshot(*payload->snapshot);
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
                case FindFilesWindowDebugCommand::ResizeVisibleResultColumn:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugResizeVisibleResultColumnPayload*>(lParam))
                    {
                        payload->result = DebugResizeVisibleResultColumn(payload->visibleIndex, payload->deltaDip);
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
                case FindFilesWindowDebugCommand::ApplyResultsLayoutFromSettings:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugApplyResultsLayoutFromSettingsPayload*>(lParam))
                    {
                        payload->result = DebugApplyResultsLayoutFromSettings();
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
                case FindFilesWindowDebugCommand::SelectResults:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugSelectResultsPayload*>(lParam))
                    {
                        payload->result = DebugSelectResults(std::move(payload->fullPaths));
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
                case FindFilesWindowDebugCommand::PostStaleSearchPayloads:
                    if (auto* payload = reinterpret_cast<FindFilesWindowDebugPostStaleSearchPayloadsPayload*>(lParam))
                    {
                        payload->result = DebugPostStaleSearchPayloads(std::move(payload->fullPath));
                        return payload->result ? TRUE : FALSE;
                    }
                    return FALSE;
            }
            return FALSE;
#endif
        case WM_NCDESTROY:
            SetWindowLongPtrW(_hWnd.get(), GWLP_USERDATA, 0);
            static_cast<void>(DrainPostedPayloadsForWindow(_hWnd.get()));
            return OnNcDestroy();
    }

    return DefWindowProcW(_hWnd.get(), message, wParam, lParam);
}

LRESULT CALLBACK FindFilesWindow::WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    FindFilesWindow* self = reinterpret_cast<FindFilesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (message == WM_NCCREATE)
    {
        auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
        self         = create ? reinterpret_cast<FindFilesWindow*>(create->lpCreateParams) : nullptr;
        if (! self)
        {
            return FALSE;
        }

        self->_hWnd.reset(hwnd);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        InitPostedPayloadWindow(hwnd);
        g_findFilesWindow = self;
        g_findFilesWindows.push_back(hwnd);
    }

    if (! self)
    {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    ++self->_dispatchDepth;
    const auto finishDispatch = wil::scope_exit([self]() noexcept
    {
        if (self->_dispatchDepth > 0u)
        {
            --self->_dispatchDepth;
        }
        if (self->_dispatchDepth == 0u && self->_deletePending)
        {
            delete self;
        }
    });

    return self->WindowProc(message, wParam, lParam);
}

} // namespace

bool ShowFindFilesWindow(HWND owner,
                         FolderWindow& applicationFolderWindow,
                         Common::Settings::Settings& settings,
                         const AppTheme& theme,
                         FindFilesPaneContext context) noexcept
{
    auto* window = new (std::nothrow) FindFilesWindow(owner, applicationFolderWindow, settings, theme, std::move(context));
    if (! window)
    {
        return false;
    }

    // Create() owns failure cleanup: when WM_CREATE fails, the window
    // procedure already deleted the instance while CreateWindowExW was
    // unwinding, so a caller-side delete here would be a double delete.
    return window->Create();
}

void UpdateFindFilesWindowsTheme(const AppTheme& theme) noexcept
{
    const auto windows = g_findFilesWindows;
    for (HWND hwnd : windows)
    {
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            continue;
        }

        if (auto* self = reinterpret_cast<FindFilesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA)))
        {
            self->UpdateTheme(theme);
        }
    }
}

HWND GetFindFilesWindowHandle() noexcept
{
    return g_findFilesWindow && g_findFilesWindow->GetHwnd() && IsWindow(g_findFilesWindow->GetHwnd()) ? g_findFilesWindow->GetHwnd() : nullptr;
}

bool IsFindFilesWindowHandle(HWND hwnd) noexcept
{
    return hwnd && IsWindow(hwnd) != FALSE && std::find(g_findFilesWindows.begin(), g_findFilesWindows.end(), hwnd) != g_findFilesWindows.end();
}

#ifdef ENABLE_TESTS
size_t DebugGetFindFilesWindowCount() noexcept
{
    size_t count = 0u;
    for (const HWND hwnd : g_findFilesWindows)
    {
        if (hwnd && IsWindow(hwnd) != FALSE)
        {
            ++count;
        }
    }
    return count;
}

std::wstring DebugMakeFindFilesResultKeyForTests(std::wstring_view pluginId, std::wstring_view instanceContext, std::wstring_view fullPath) noexcept
{
    return MakeResultKey(pluginId, instanceContext, fullPath);
}

std::vector<size_t> DebugSelectKnownCompletedFindFilesSourceIndicesForTests(size_t sourceCount,
                                                                            std::span<const FindFilesDebugSourceOutcome> outcomes,
                                                                            HRESULT overallStatus)
{
    std::vector<std::wstring> keys;
    keys.reserve(sourceCount);
    for (size_t index = 0; index < sourceCount; ++index)
    {
        keys.push_back(std::to_wstring(index));
    }

    const std::unordered_set<std::wstring> completedKeys =
        CollectKnownCompletedResultKeys(std::span<const std::wstring>(keys), outcomes, overallStatus);
    std::vector<size_t> completedIndices;
    completedIndices.reserve(completedKeys.size());
    for (size_t index = 0; index < keys.size(); ++index)
    {
        if (completedKeys.contains(keys[index]))
        {
            completedIndices.push_back(index);
        }
    }
    return completedIndices;
}

bool DebugConfigureFindFilesWindow(std::wstring rootPath,
                                   std::wstring namePattern,
                                   std::wstring contentPattern,
                                   Common::Settings::SearchNameMode nameMode,
                                   Common::Settings::SearchContentMode contentMode) noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugConfigurePayload payload{
        .rootPath       = std::move(rootPath),
        .namePattern    = std::move(namePattern),
        .contentPattern = std::move(contentPattern),
        .nameMode       = nameMode,
        .contentMode    = contentMode,
    };
    return SendMessageW(
               findWindow, kFindFilesWindowDebugMessage, static_cast<WPARAM>(FindFilesWindowDebugCommand::Configure), reinterpret_cast<LPARAM>(&payload)) !=
               FALSE &&
           payload.result;
}

bool DebugSetFindFilesWindowOptions(bool recursive, bool includeFiles, bool includeDirectories, bool preferIndex, bool wantSnippets) noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugSetOptionsPayload payload{
        .recursive          = recursive,
        .includeFiles       = includeFiles,
        .includeDirectories = includeDirectories,
        .preferIndex        = preferIndex,
        .wantSnippets       = wantSnippets,
    };
    return SendMessageW(
               findWindow, kFindFilesWindowDebugMessage, static_cast<WPARAM>(FindFilesWindowDebugCommand::SetOptions), reinterpret_cast<LPARAM>(&payload)) !=
               FALSE &&
           payload.result;
}

bool DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget target, std::wstring text) noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugSetComboTextPayload payload{
        .target = target,
        .text   = std::move(text),
    };
    return SendMessageW(
               findWindow, kFindFilesWindowDebugMessage, static_cast<WPARAM>(FindFilesWindowDebugCommand::SetComboText), reinterpret_cast<LPARAM>(&payload)) !=
               FALSE &&
           payload.result;
}

bool DebugSetFindFilesWindowDestinationPath(std::wstring path) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugSetDestinationPath(std::move(path)) : false;
}

bool DebugStartFindFilesWindowSearch(FindFilesDebugOperation operation) noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugStartSearchPayload payload{
        .operation = operation,
    };
    return SendMessageW(
               findWindow, kFindFilesWindowDebugMessage, static_cast<WPARAM>(FindFilesWindowDebugCommand::StartSearch), reinterpret_cast<LPARAM>(&payload)) !=
               FALSE &&
           payload.result;
}

bool DebugCancelFindFilesWindowSearch() noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugCancelSearch() : false;
}

bool DebugPostFindFilesWindowStaleSearchPayloads(std::wstring fullPath) noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugPostStaleSearchPayloadsPayload payload{
        .fullPath = std::move(fullPath),
    };
    return SendMessageW(findWindow,
                        kFindFilesWindowDebugMessage,
                        static_cast<WPARAM>(FindFilesWindowDebugCommand::PostStaleSearchPayloads),
                        reinterpret_cast<LPARAM>(&payload)) != FALSE &&
           payload.result;
}

bool DebugGetFindFilesWindowSnapshot(FindFilesDebugSnapshot& out) noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugSnapshotPayload payload{
        .snapshot = &out,
    };
    return SendMessageW(
               findWindow, kFindFilesWindowDebugMessage, static_cast<WPARAM>(FindFilesWindowDebugCommand::GetSnapshot), reinterpret_cast<LPARAM>(&payload)) !=
               FALSE &&
           payload.result;
}

bool DebugHitTestFindFilesWindowResultsGrid(float xDip, float yDip, FindFilesDebugGridHit& out) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugHitTestResultsGrid(D2D1::Point2F(xDip, yDip), out) : false;
}

bool DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget target) noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugFocusTargetPayload payload{
        .target = target,
    };
    return SendMessageW(
               findWindow, kFindFilesWindowDebugMessage, static_cast<WPARAM>(FindFilesWindowDebugCommand::FocusTarget), reinterpret_cast<LPARAM>(&payload)) !=
               FALSE &&
           payload.result;
}

bool DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget target, RECT& outRect) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugGetTargetClientRect(target, outRect) : false;
}

bool DebugSetFindFilesWindowResultSort(size_t columnIndex, bool descending) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugSetResultSort(columnIndex, descending) : false;
}

bool DebugReorderFindFilesWindowVisibleResultColumn(size_t fromVisibleIndex, size_t targetVisibleIndex) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugReorderVisibleResultColumn(fromVisibleIndex, targetVisibleIndex) : false;
}

bool DebugResizeFindFilesWindowVisibleResultColumn(size_t visibleIndex, float deltaDip) noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugResizeVisibleResultColumnPayload payload{
        .visibleIndex = visibleIndex,
        .deltaDip     = deltaDip,
    };
    return SendMessageW(findWindow,
                        kFindFilesWindowDebugMessage,
                        static_cast<WPARAM>(FindFilesWindowDebugCommand::ResizeVisibleResultColumn),
                        reinterpret_cast<LPARAM>(&payload)) != FALSE &&
           payload.result;
}

bool DebugApplyFindFilesWindowResultsLayoutFromSettings() noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugApplyResultsLayoutFromSettingsPayload payload{};
    return SendMessageW(findWindow,
                        kFindFilesWindowDebugMessage,
                        static_cast<WPARAM>(FindFilesWindowDebugCommand::ApplyResultsLayoutFromSettings),
                        reinterpret_cast<LPARAM>(&payload)) != FALSE &&
           payload.result;
}

bool DebugSelectFindFilesWindowResult(std::wstring fullPath) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugSelectResult(std::move(fullPath)) : false;
}

bool DebugSelectFindFilesWindowResults(std::vector<std::wstring> fullPaths) noexcept
{
    const HWND findWindow = GetFindFilesWindowHandle();
    if (! g_findFilesWindow || ! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesWindowDebugSelectResultsPayload payload{.fullPaths = std::move(fullPaths)};
    return SendMessageW(
               findWindow, kFindFilesWindowDebugMessage, static_cast<WPARAM>(FindFilesWindowDebugCommand::SelectResults), reinterpret_cast<LPARAM>(&payload)) !=
               FALSE &&
           payload.result;
}

bool DebugActivateSelectedFindFilesWindowResult() noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugActivateSelectedResult() : false;
}

bool DebugOpenSelectedFindFilesWindowResultParent() noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugOpenSelectedResultParent() : false;
}

bool DebugGetSelectedFindFilesWindowOpenDisposition(bool parentOnly, FindFilesDebugOpenDisposition& out) noexcept
{
    out = FindFilesDebugOpenDisposition::None;
    return g_findFilesWindow ? g_findFilesWindow->DebugGetSelectedOpenDisposition(parentOnly, out) : false;
}

bool DebugScrollFindFilesWindowResultsByWheelDetents(int detents) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugScrollResultsByWheelDetents(detents) : false;
}

bool DebugWaitForFindFilesWindowIdle(uint32_t timeoutMs) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugWaitForIdle(timeoutMs) : true;
}

bool DebugFindFilesIsNextQueuedMessage(HWND targetHwnd, UINT targetMessage) noexcept
{
    return IsNextThreadQueueMessage(targetHwnd, targetMessage);
}

void DebugConfigureFindFilesWindowSearchRunBlocker(bool enabled) noexcept
{
    g_debugSearchRunBlocked.store(false, std::memory_order_release);
    g_debugSearchRunRelease.store(! enabled, std::memory_order_release);
    g_debugSearchRunBlockEnabled.store(enabled, std::memory_order_release);
}

void DebugReleaseFindFilesWindowSearchRunBlocker() noexcept
{
    g_debugSearchRunRelease.store(true, std::memory_order_release);
}

bool DebugWaitForFindFilesWindowSearchRunBlocked(uint32_t timeoutMs) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(timeoutMs);
    while (std::chrono::steady_clock::now() < deadline)
    {
        if (g_debugSearchRunBlocked.load(std::memory_order_acquire))
        {
            return true;
        }

        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        std::this_thread::sleep_for(10ms);
    }

    return g_debugSearchRunBlocked.load(std::memory_order_acquire);
}
#endif
