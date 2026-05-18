#include "Framework.h"

#include "FindFilesWindow.h"

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cmath>
#include <condition_variable>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <format>
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

#include "DxUi/DxUi.h"
#include "FileSystemPluginManager.h"
#include "FolderWindow.h"
#include "Helpers.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "PlugInterfaces/Factory.h"
#include "SearchFallbackEngine.h"
#include "SearchServiceBroker.h"
#include "UiMetrics.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "WindowSizing.h"
#include "resource.h"

extern FolderWindow g_folderWindow;

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::Checkbox;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::IDxGridDelegate;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::StatusStrip;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kFindFilesWindowClassName[] = L"RedSalamander.FindFilesWindow";
constexpr wchar_t kFindFilesWindowId[]        = L"FindFilesWindow";

constexpr uint32_t kMaxSnippetCharacters          = 160u;
constexpr uint64_t kMaxContentBytes               = 64ull * 1024ull * 1024ull;
constexpr size_t kMaxRecentEntries                = 10u;
constexpr size_t kBatchSize                       = 32u;
constexpr size_t kProgressFlushMinBatchSize       = 8u;
constexpr auto kProgressFlushMaxAge               = std::chrono::milliseconds(40);
constexpr UINT_PTR kStatusRefreshTimerId          = 1u;
constexpr UINT_PTR kDeferredResultsRefreshTimerId = 2u;
constexpr UINT kStatusRefreshTimerIntervalMs      = 250u;
constexpr UINT kDeferredResultsRefreshDelayMs     = 25u;
constexpr uint64_t kStatusStallThresholdMs        = 5000u;
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
    std::wstring displayName;
    std::wstring previewText;
    unsigned long fileAttributes = 0;
    int64_t lastWriteTime        = 0;
    int64_t endOfFile            = 0;
    uint32_t matchedBy           = 0;
};

struct FindSearchResultsPayload
{
    std::vector<FindResultRecord> results;
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

[[nodiscard]] ThemePalette MakeDxPalette(const AppTheme& theme) noexcept
{
    const auto mix = [](const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, float t) noexcept
    {
        const float clamped = std::clamp(t, 0.0f, 1.0f);
        return D2D1::ColorF(a.r + ((b.r - a.r) * clamped), a.g + ((b.g - a.g) * clamped), a.b + ((b.b - a.b) * clamped), a.a + ((b.a - a.a) * clamped));
    };

    ThemePalette palette          = RedSalamander::DxUi::MakeDefaultThemePalette(theme.dark);
    palette.dark                  = theme.dark;
    palette.highContrast          = theme.highContrast;
    palette.rainbowMode           = theme.menu.rainbowMode;
    palette.accent                = theme.accent;
    palette.windowBackground      = ColorFromCOLORREF(theme.windowBackground);
    palette.surfaceBackground     = theme.folderView.backgroundColor;
    palette.headerBackground      = ColorFromCOLORREF(theme.menu.background);
    palette.headerHovered         = mix(palette.headerBackground, palette.accent, theme.dark ? 0.22f : 0.10f);
    palette.headerPressed         = mix(palette.headerBackground, palette.accent, theme.dark ? 0.32f : 0.18f);
    palette.border                = ColorFromCOLORREF(theme.menu.border);
    palette.gridLine              = theme.folderView.gridLines;
    palette.text                  = theme.folderView.textNormal;
    palette.subduedText           = ColorFromCOLORREF(theme.menu.shortcutText);
    palette.selectionFill         = ColorFromCOLORREF(theme.menu.selectionBg);
    palette.selectionText         = ColorFromCOLORREF(theme.menu.selectionText);
    palette.selectionInactiveFill = D2D1::ColorF(palette.selectionFill.r, palette.selectionFill.g, palette.selectionFill.b, theme.highContrast ? 1.0f : 0.55f);
    palette.focusStroke           = theme.folderView.focusBorder;
    palette.hoverFill             = D2D1::ColorF(palette.accent.r, palette.accent.g, palette.accent.b, theme.dark ? 0.18f : 0.10f);
    palette.buttonFill            = ColorFromCOLORREF(theme.menu.background);
    palette.buttonBorder          = ColorFromCOLORREF(theme.menu.border);
    palette.buttonHotFill         = palette.headerHovered;
    palette.buttonPressedFill     = palette.headerPressed;
    palette.inputFill             = theme.folderView.backgroundColor;
    palette.inputBorder           = ColorFromCOLORREF(theme.menu.border);
    palette.scrollbarTrack        = theme.fileOperations.scrollbarTrack;
    palette.scrollbarThumb        = theme.fileOperations.scrollbarThumb;
    palette.scrollbarThumbHot =
        D2D1::ColorF(palette.scrollbarThumb.r, palette.scrollbarThumb.g, palette.scrollbarThumb.b, std::min(1.0f, palette.scrollbarThumb.a + 0.10f));
    palette.infoFill    = theme.folderView.infoBackground;
    palette.infoText    = theme.folderView.infoText;
    palette.warningFill = theme.folderView.warningBackground;
    palette.warningText = theme.folderView.warningText;
    palette.errorFill   = theme.folderView.errorBackground;
    palette.errorText   = theme.folderView.errorText;
    return palette;
}

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
            case kColumnName: outCell.text = record.displayName; break;
            case kColumnPath: outCell.text = GetResultRelativeFolderPath(record.relativePath); break;
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
            style.rainbowSeed = _rows->at(rowIndex).key;
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
    if (pluginId.empty())
    {
        return nullptr;
    }

    const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
    for (const FileSystemPluginManager::PluginEntry& entry : plugins)
    {
        if (entry.id.empty())
        {
            continue;
        }

        if (CompareStringOrdinal(entry.id.c_str(), -1, pluginId.data(), static_cast<int>(pluginId.size()), TRUE) == CSTR_EQUAL)
        {
            return &entry;
        }
    }

    return nullptr;
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
    std::wstring result(value);
    std::transform(
        result.begin(), result.end(), result.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))); });
    return result;
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

    [[nodiscard]] bool Start(FindFilesWindow& owner, SearchRequest request) noexcept;
    void Cancel() noexcept;
    void Shutdown() noexcept;
    void NotifyUiSettled() noexcept;
    [[nodiscard]] bool IsActive() const noexcept;
#ifdef ENABLE_TESTS
    [[nodiscard]] bool IsUiSettled() const noexcept;
    [[nodiscard]] bool WaitForIdle(uint32_t timeoutMs) noexcept;
#endif

private:
    void Run(SearchRequest request) noexcept;
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
    using IDxGridDelegate::OnGridRowActivated;
    using IDxGridDelegate::OnGridSelectionChanged;

    FindFilesWindow(HWND owner, Common::Settings::Settings& settings, AppTheme theme, FindFilesPaneContext context) noexcept;
    FindFilesWindow(const FindFilesWindow&)            = delete;
    FindFilesWindow& operator=(const FindFilesWindow&) = delete;

    [[nodiscard]] bool Create() noexcept;
    void UpdateTheme(const AppTheme& theme) noexcept;
    void UpdateOwnerWindow(HWND owner) noexcept;
    void UpdateContext(FindFilesPaneContext context) noexcept;
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
    [[nodiscard]] bool DebugActivateSelectedResult() noexcept;
    [[nodiscard]] bool DebugOpenSelectedResultParent() noexcept;
    [[nodiscard]] bool DebugScrollResultsByWheelDetents(int detents) noexcept;
    [[nodiscard]] bool DebugWaitForIdle(uint32_t timeoutMs) noexcept;
#endif

    LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;

private:
    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept;
    void OnClose() noexcept;
    LRESULT OnNcDestroy() noexcept;
    void BuildUi() noexcept;
    void Layout() noexcept;
    void ApplyTheme() noexcept;
    void PopulateFromSettings() noexcept;
    void PopulateModeCombos() noexcept;
    void PopulateHistoryCombos() noexcept;
    void ApplyResultsSortFromSettings() noexcept;
    void ApplyResultsGridLayoutFromSettings() noexcept;
    void UpdateOptionDependencies() noexcept;
    void UpdateActionButtons() noexcept;
    void PersistUiState(bool updateHistory) noexcept;
    struct SearchTextOverride final
    {
        std::wstring rootPath;
        std::wstring namePattern;
        std::wstring contentPattern;
    };
#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugStartSearchWithTextOverride(FindFilesDebugOperation operation, SearchTextOverride textOverride) noexcept;
#endif
    [[nodiscard]] bool BeginSearch(SearchOperation operation, const SearchTextOverride* textOverride = nullptr) noexcept;
    void OnSearchStarted(SearchOperation operation, const SearchRequest& request) noexcept;
    void OnSearchResults(std::unique_ptr<FindSearchResultsPayload> payload) noexcept;
    void OnSearchProgress(std::unique_ptr<FindSearchProgressPayload> payload) noexcept;
    void OnSearchComplete(std::unique_ptr<FindSearchCompletePayload> payload) noexcept;
    void ApplyDeferredSetOperation(SearchOperation operation) noexcept;
    void ClearResults() noexcept;
    void RebuildResultsList() noexcept;
    [[nodiscard]] ResultListMutation AddOrUpdateVisibleResult(FindResultRecord result) noexcept;
    void RemoveKeysFromResults(const std::unordered_set<std::wstring>& keys) noexcept;
    void KeepOnlyKeysInResults(const std::unordered_set<std::wstring>& keys) noexcept;
    [[nodiscard]] std::optional<size_t> FindResultColumnIndexById(std::wstring_view columnId) const noexcept;
    [[nodiscard]] std::optional<size_t> GetSelectedResultIndex() const noexcept;
    void OpenSelectedResult(bool parentOnly) noexcept;
    void SetStatusText(std::wstring text) noexcept;
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

    HWND _ownerWindow                     = nullptr;
    HWND _restoreFocusWindow              = nullptr;
    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    FindFilesPaneContext _context;
    size_t _dispatchDepth = 0u;
    bool _deletePending   = false;

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
    StatusStrip* _statusText           = nullptr;
    Grid* _resultsList                 = nullptr;
    FindResultsGridModel _resultsModel{_theme};
    GridSortSpec _resultSortSpec{};
    uint64_t _nextResultOrdinal = 1u;

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
    uint64_t _searchStartedTickMs         = 0u;
    uint64_t _lastProgressTickMs          = 0u;
    uint64_t _lastBackendStatusTickMs     = 0u;
    uint64_t _lastBackendStatusPollTickMs = 0u;
    std::wstring _status;
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
    bool _resultsRefreshPending       = false;
    bool _resultsRefreshTimerArmed    = false;
    bool _resultsRefreshFullRebuild   = false;
    size_t _resultsRefreshBatchCount  = 0u;
    size_t _resultsRefreshRecordCount = 0u;
#ifdef ENABLE_TESTS
    uint32_t _debugResultListFullRebuildCount          = 0;
    float _debugResizeBeforeWidthDip                   = 0.0f;
    float _debugResizeTargetWidthDip                   = 0.0f;
    float _debugResizeObservedWidthDip                 = 0.0f;
    bool _debugResizeSucceeded                         = false;
    FindFilesDebugFocusTarget _debugLastSetComboTarget = FindFilesDebugFocusTarget::None;
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
    explicit SearchCallbacks(HWND hwnd, const SearchRequest& request, std::atomic<bool>& cancelRequested) noexcept
        : _hwnd(hwnd),
          _request(request),
          _cancelRequested(cancelRequested)
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
        payload->enqueuedAt = now;
        _batch.clear();
        _batchFirstQueuedAt = {};
        if (! PostMessagePayload(_hwnd, WndMsg::kFindSearchResults, 0, std::move(payload)))
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
        record.pluginId        = _request.context.pluginId;
        record.pluginShortId   = _request.context.pluginShortId;
        record.instanceContext = _request.context.instanceContext;
        record.fullPath        = CopySizedUtf16(match->fullPath, match->fullPathSize);
        record.relativePath    = CopySizedUtf16(match->relativePath, match->relativePathSize);
        record.displayName     = CopySizedUtf16(match->displayName, match->displayNameSize);
        record.previewText     = CopySizedUtf16(match->previewText, match->previewTextSize);
        record.fileAttributes  = match->fileAttributes;
        record.lastWriteTime   = match->lastWriteTime;
        record.endOfFile       = match->endOfFile;
        record.matchedBy       = match->matchedBy;
        record.key             = MakeResultKey(record.pluginId, record.instanceContext, record.fullPath);
        record.stableRowId     = MakeResultStableId(record);

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

        return PostMessagePayload(_hwnd, WndMsg::kFindSearchProgress, 0, std::move(payload)) ? S_OK : E_FAIL;
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

bool SearchSessionController::Start(FindFilesWindow& owner, SearchRequest request) noexcept
{
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
        _worker = std::jthread([this, request = std::move(request)]() mutable noexcept { Run(std::move(request)); });
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

void SearchSessionController::Run(SearchRequest request) noexcept
{
    const Debug::Perf::Scope runPerf(L"find.session.run_total_ms");
    auto initialPayload = std::unique_ptr<FindSearchProgressPayload>(new (std::nothrow) FindSearchProgressPayload{});
    if (initialPayload)
    {
        initialPayload->phase      = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
        initialPayload->backend    = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
        initialPayload->statusHint = S_OK;
        initialPayload->enqueuedAt = SteadyClock::now();
        static_cast<void>(PostMessagePayload(_ownerHwnd, WndMsg::kFindSearchProgress, 0, std::move(initialPayload)));
    }

    SearchCallbacks callbacks(_ownerHwnd, request, _cancelRequested);

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

    bool completionQueued = false;
    auto complete         = std::unique_ptr<FindSearchCompletePayload>(new (std::nothrow) FindSearchCompletePayload{});
    if (complete)
    {
        complete->hr                 = hr;
        complete->backend            = callbacks._latestProgress.backend;
        complete->warningFlags       = callbacks._latestProgress.warningFlags;
        complete->scannedDirectories = callbacks._latestProgress.scannedDirectories;
        complete->scannedFiles       = callbacks._latestProgress.scannedFiles;
        complete->candidateFiles     = callbacks._latestProgress.candidateFiles;
        complete->matchedEntries     = callbacks._latestProgress.matchedEntries;
        completionQueued             = PostMessagePayload(_ownerHwnd, WndMsg::kFindSearchComplete, 0, std::move(complete));
    }

    if (! completionQueued)
    {
        NotifyUiSettled();
    }

    MarkIdle();
}

void SearchSessionController::MarkIdle() noexcept
{
    {
        std::lock_guard lock(_mutex);
        _active.store(false, std::memory_order_release);
    }
    _idleCv.notify_all();
}

FindFilesWindow::FindFilesWindow(HWND owner, Common::Settings::Settings& settings, AppTheme theme, FindFilesPaneContext context) noexcept
    : _settings(&settings),
      _theme(std::move(theme)),
      _context(std::move(context))
{
    UpdateOwnerWindow(owner);
}

bool FindFilesWindow::Create() noexcept
{
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
        return false;
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_FIND_TITLE);
    const HWND created       = CreateWindowExW(0,
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
    if (! created)
    {
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

void FindFilesWindow::UpdateContext(FindFilesPaneContext context) noexcept
{
    const bool pluginChanged =
        ! OrdinalString::EqualsNoCase(_context.pluginId, context.pluginId) || ! OrdinalString::EqualsNoCase(_context.instanceContext, context.instanceContext);
    _context = std::move(context);
    if (pluginChanged)
    {
        _deferredKeys.clear();
        if (! _session.IsActive())
        {
            ClearResults();
        }

        const std::wstring rootText = (! _settings || ! _settings->search.has_value() || _settings->search->lastRoot.empty()) ? _context.rootPluginPath.native()
                                                                                                                              : _settings->search->lastRoot;
        if (_rootCombo)
        {
            _rootCombo->SetText(rootText);
        }
    }
}

bool FindFilesWindow::OnCreate(HWND hwnd) noexcept
{
    if (! _dxHost.Attach(hwnd))
    {
        Debug::Error(L"FindFiles: failed to attach DxUi host.");
        return false;
    }

    BuildUi();
    PopulateModeCombos();
    PopulateHistoryCombos();
    PopulateFromSettings();
    ApplyTheme();
    UpdateOptionDependencies();
    ApplyResultsSortFromSettings();
    ApplyResultsGridLayoutFromSettings();
    UpdateActionButtons();
    SetStatusText(LoadStringResource(nullptr, IDS_FIND_STATUS_READY));
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
    _rootCombo->SetOnTextChanged([this](std::wstring_view) { PersistUiState(false); });
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
    _findButton->SetMnemonic(L'F');
    bindAction(_appendButton, IDS_FIND_ACTION_APPEND, SearchOperation::Append);
    _appendButton->SetMnemonic(L'A');
    bindAction(_intersectButton, IDS_FIND_ACTION_INTERSECT, SearchOperation::Intersect);
    _intersectButton->SetMnemonic(L'I');
    bindAction(_subtractButton, IDS_FIND_ACTION_SUBTRACT, SearchOperation::Subtract);
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

    _statusText = _root->AddChild<StatusStrip>(LoadStringResource(nullptr, IDS_FIND_STATUS_READY));
    _statusText->SetFontRole(RedSalamander::DxUi::FontRole::Small);

    _resultsList = _root->AddChild<Grid>();
    _resultsList->SetDelegate(this);
    _resultsList->SetModel(&_resultsModel);
    _resultsList->SetSelectionMode(GridSelectionMode::Single);
    _resultsList->SetHeaderHeightDip(30.0f);
    _resultsList->SetRowHeightDip(46.0f);
    _resultsList->SetLineClamp(3u);

    _resultsModel.SetRows(&_results);
    _dxHost.SetRoot(std::move(_rootStorage));
    _dxHost.SetDefaultButton(_findButton);
    _dxHost.SetCancelButton(_cancelButton);
    _dxHost.SetOnEscape([this]() noexcept
    {
        if (_session.IsActive() || ! _hWnd || _dxHost.GetFocusControl() == nullptr)
        {
            return false;
        }

        return PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0) != FALSE;
    });
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
}

void FindFilesWindow::PopulateFromSettings() noexcept
{
    const Common::Settings::SearchDialogSettings defaults{};
    const auto settings = _settings && _settings->search.has_value() ? _settings->search.value() : defaults;

    const std::wstring initialRoot = settings.lastRoot.empty() ? _context.rootPluginPath.native() : settings.lastRoot;
    if (_rootCombo)
    {
        _rootCombo->SetText(initialRoot);
    }
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

    _dxHost.SetTheme(MakeDxPalette(_theme));
    ApplyTitleBarTheme(_hWnd.get(), _theme, GetActiveWindow() == _hWnd.get());
    ApplyWindowBackdropTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool);
    _dxHost.Invalidate();
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

    _resultsModel.SetShowSnippetColumn(contentEnabled);
    if (_resultSortSpec.direction != SortDirection::None && _resultSortSpec.columnIndex >= _resultsModel.GetColumnCount())
    {
        _resultSortSpec = {};
    }
    if (_resultsList)
    {
        _resultsList->SetSortSpec(_resultSortSpec);
        _resultsList->NotifyDataChanged();
        ApplyResultsGridLayoutFromSettings();
    }
}

void FindFilesWindow::UpdateActionButtons() noexcept
{
    const bool active       = _session.IsActive();
    const bool hasSelection = GetSelectedResultIndex().has_value();

    if (_findButton)
    {
        _findButton->SetEnabled(! active);
    }
    if (_appendButton)
    {
        _appendButton->SetEnabled(! active);
    }
    if (_intersectButton)
    {
        _intersectButton->SetEnabled(! active);
    }
    if (_subtractButton)
    {
        _subtractButton->SetEnabled(! active);
    }
    if (_cancelButton)
    {
        _cancelButton->SetEnabled(active);
    }
    if (_openButton)
    {
        _openButton->SetEnabled(! active && hasSelection);
    }
    if (_parentButton)
    {
        _parentButton->SetEnabled(! active && hasSelection);
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

    const float widthDip     = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.right - client.left)));
    const float heightDip    = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.bottom - client.top)));
    const bool compact       = _theme.compactMode;
    const float margin       = compact ? 10.0f : 12.0f;
    const float gap          = compact ? 6.0f : 8.0f;
    const float labelWidth   = compact ? 72.0f : 76.0f;
    const float rowHeight    = compact ? 28.0f : 30.0f;
    const float optionHeight = compact ? 22.0f : 24.0f;
    const float buttonHeight = compact ? 30.0f : 32.0f;
    const float buttonWidth  = compact ? 104.0f : 108.0f;
    const float modeWidth    = compact ? 142.0f : 150.0f;
    const float statusHeight = compact ? 20.0f : 22.0f;
    const auto snapDip       = [this](float dip) noexcept { return _dxHost.PixelsToDip(std::round(_dxHost.DipsToPixels(dip))); };
    const auto rect          = [&snapDip](float left, float top, float rightEdge, float bottom) noexcept
    { return D2D1::RectF(snapDip(left), snapDip(top), snapDip(rightEdge), snapDip(bottom)); };

    _root->SetBounds(rect(0.0f, 0.0f, widthDip, heightDip));

    float y                    = margin;
    const float right          = std::max(margin, widthDip - margin);
    const float contentWidth   = std::max(100.0f, widthDip - (2.0f * margin));
    const float comboWidth     = std::max(100.0f, contentWidth - labelWidth - gap - modeWidth - gap);
    const float fullComboWidth = std::max(100.0f, contentWidth - labelWidth - gap);

    _rootLabel->SetBounds(rect(margin, y, margin + labelWidth, y + rowHeight));
    _rootCombo->SetBounds(rect(margin + labelWidth + gap, y, margin + labelWidth + gap + fullComboWidth, y + rowHeight));
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
    _findButton->SetBounds(rect(buttonX, y, buttonX + buttonWidth, y + buttonHeight));
    buttonX += buttonWidth + gap;
    _appendButton->SetBounds(rect(buttonX, y, buttonX + buttonWidth, y + buttonHeight));
    buttonX += buttonWidth + gap;
    _intersectButton->SetBounds(rect(buttonX, y, buttonX + buttonWidth, y + buttonHeight));
    buttonX += buttonWidth + gap;
    _subtractButton->SetBounds(rect(buttonX, y, buttonX + buttonWidth, y + buttonHeight));
    buttonX += buttonWidth + gap;
    _cancelButton->SetBounds(rect(buttonX, y, buttonX + buttonWidth, y + buttonHeight));
    buttonX += buttonWidth + gap;
    _openButton->SetBounds(rect(buttonX, y, buttonX + buttonWidth, y + buttonHeight));
    buttonX += buttonWidth + gap;
    _parentButton->SetBounds(rect(buttonX, y, buttonX + buttonWidth, y + buttonHeight));
    y += buttonHeight + gap;

    _statusText->SetBounds(rect(margin, y, right, y + statusHeight));
    y += statusHeight + gap;

    _resultsList->SetHeaderHeightDip(compact ? 28.0f : 30.0f);
    _resultsList->SetRowHeightDip(compact ? 42.0f : 46.0f);
    _resultsList->SetLineClamp(compact ? 2u : 3u);
    _resultsList->SetBounds(rect(margin, y, right, std::max(y + 120.0f, heightDip - margin)));
    _dxHost.Invalidate();
}

std::wstring FindFilesWindow::GetComboText(const ComboBox* combo) const noexcept
{
    return combo ? std::wstring(combo->GetText()) : std::wstring{};
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
    HWND focusedFolderView = g_folderWindow.GetFocusedFolderViewHwnd();
    if (! focusedFolderView)
    {
        focusedFolderView = g_folderWindow.GetFolderViewHwnd(g_folderWindow.GetFocusedPane());
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
    if (_preferIndexCheck && _preferIndexCheck->IsChecked())
    {
        request.flags = static_cast<FileSystemSearchFlags>(request.flags | FILESYSTEM_SEARCH_PREFER_INDEX);
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

    return request;
}

void FindFilesWindow::OnClose() noexcept
{
    if (_closeRequested)
    {
        return;
    }

    _closeRequested = true;
    if (_hWnd)
    {
        static_cast<void>(::KillTimer(_hWnd.get(), kStatusRefreshTimerId));
    }
    PersistUiState(false);
    _session.Shutdown();
}

LRESULT FindFilesWindow::OnNcDestroy() noexcept
{
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
    _statusText              = nullptr;
    _resultsList             = nullptr;
    _hWnd.release();
    g_findFilesWindows.erase(std::remove(g_findFilesWindows.begin(), g_findFilesWindows.end(), hwnd), g_findFilesWindows.end());
    if (g_findFilesWindow == this)
    {
        g_findFilesWindow = nullptr;
    }
    _deletePending = true;

    if (restoreOwner)
    {
        static_cast<void>(SetActiveWindow(restoreOwner));
        static_cast<void>(SetForegroundWindow(restoreOwner));
    }
    if (restoreFocus)
    {
        g_folderWindow.RequestRestoreFolderViewFocus(restoreFocus);
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
    _resultsRefreshTimerArmed  = false;
    _resultsRefreshPending     = false;
    _resultsRefreshFullRebuild = false;
    _resultsRefreshBatchCount  = 0u;
    _resultsRefreshRecordCount = 0u;
    _nextResultOrdinal         = 1u;
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
    result.stableRowId = MakeResultStableId(result);

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

void FindFilesWindow::OpenSelectedResult(bool parentOnly) noexcept
{
    const auto index = GetSelectedResultIndex();
    if (! index.has_value())
    {
        return;
    }

    const FindResultRecord& record = _results[index.value()];
    const bool isDirectory         = (record.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    const std::filesystem::path fullPath(record.fullPath);

    std::filesystem::path targetFolder;
    std::wstring focusName;
    unsigned int commandId = 0u;

    if (parentOnly)
    {
        targetFolder = fullPath.parent_path();
        focusName    = record.displayName;
    }
    else if (isDirectory)
    {
        targetFolder = fullPath;
    }
    else
    {
        targetFolder = fullPath.parent_path();
        focusName    = record.displayName;
        commandId    = IDM_FOLDERVIEW_CONTEXT_OPEN;
    }

    const std::filesystem::path historyPath = NavigationLocation::FormatHistoryPath(record.pluginShortId, record.instanceContext, targetFolder);
    static_cast<void>(g_folderWindow.ExecuteInActivePane(historyPath, focusName, commandId, true));
}

void FindFilesWindow::SetStatusText(std::wstring text) noexcept
{
    _status = std::move(text);
    if (_statusText)
    {
        _statusText->SetText(_status);
    }
    _dxHost.Invalidate();
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
        // Keep the user-selected result-grid layout stable across late refreshes
        // triggered by search completion or rerun bookkeeping.
        ApplyResultsGridLayoutFromSettings();
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

    if (::SetTimer(_hWnd.get(), kDeferredResultsRefreshTimerId, kDeferredResultsRefreshDelayMs, nullptr) == 0)
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

    if (_resultsRefreshTimerArmed && _hWnd)
    {
        static_cast<void>(::KillTimer(_hWnd.get(), kDeferredResultsRefreshTimerId));
    }
    _resultsRefreshTimerArmed  = false;
    _resultsRefreshPending     = false;
    _resultsRefreshFullRebuild = false;
    _resultsRefreshBatchCount  = 0u;
    _resultsRefreshRecordCount = 0u;

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

    return LoadStringResource(nullptr, IDS_FIND_STATUS_READY);
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
    if (_session.IsActive())
    {
        return false;
    }

    _lastBeginRootPath       = textOverride ? textOverride->rootPath : GetComboText(_rootCombo);
    _lastBeginNamePattern    = textOverride ? textOverride->namePattern : GetComboText(_nameCombo);
    _lastBeginContentPattern = textOverride ? textOverride->contentPattern : GetComboText(_contentCombo);

    std::optional<SearchRequest> request = BuildSearchRequest(textOverride);
    if (! request.has_value())
    {
        return false;
    }

    PersistUiState(true);
    PopulateHistoryCombos();
    OnSearchStarted(operation, request.value());

    if (! _session.Start(*this, std::move(request.value())))
    {
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
    return true;
}

void FindFilesWindow::OnSearchResults(std::unique_ptr<FindSearchResultsPayload> payload) noexcept
{
    if (! payload)
    {
        return;
    }

    const Debug::Perf::Scope resultsPerf(L"find.ui.results_handler_ms");
    EmitPerfCount(L"find.ui.results_message_count");
    EmitPerfCount(L"find.ui.results_message_batch_size", static_cast<uint64_t>(payload->results.size()));
    if (payload->enqueuedAt != SteadyClock::time_point{})
    {
        Debug::Perf::Emit(
            L"find.ui.results_to_visible_latency_ms", L"", ElapsedUsSince(payload->enqueuedAt), static_cast<uint64_t>(payload->results.size()), 0u, S_OK);
    }

    if (_activeOperation == SearchOperation::Intersect || _activeOperation == SearchOperation::Subtract)
    {
        for (auto& record : payload->results)
        {
            _deferredKeys.insert(record.key);
        }
        return;
    }

    bool changed         = false;
    size_t insertedCount = 0u;
    size_t updatedCount  = 0u;
    for (auto& record : payload->results)
    {
        const ResultListMutation mutation = AddOrUpdateVisibleResult(std::move(record));
        insertedCount += mutation.inserted ? 1u : 0u;
        updatedCount += mutation.inserted ? 0u : 1u;
        changed = true;
    }

    if (! changed)
    {
        return;
    }

    EmitPerfCount(L"find.ui.results_inserted_count", static_cast<uint64_t>(insertedCount));
    EmitPerfCount(L"find.ui.results_updated_count", static_cast<uint64_t>(updatedCount));

    if (_resultSortSpec.direction == SortDirection::None)
    {
        RefreshResultsView(false);
    }
    else
    {
        // Coalesce sorted result refreshes so progress-driven micro-batches do not force repeated full rebuilds.
        ScheduleResultsRefresh(true, payload->results.size());
    }
}

void FindFilesWindow::OnSearchProgress(std::unique_ptr<FindSearchProgressPayload> payload) noexcept
{
    if (! payload)
    {
        return;
    }

    const Debug::Perf::Scope progressPerf(L"find.ui.progress_handler_ms");
    uint64_t drainedProgressCount = 0u;
    while (_hWnd)
    {
        MSG queuedMessage{};
        if (! ::PeekMessageW(&queuedMessage, _hWnd.get(), 0, 0, PM_NOREMOVE))
        {
            break;
        }
        if (queuedMessage.message != WndMsg::kFindSearchProgress)
        {
            break;
        }
        if (! ::PeekMessageW(&queuedMessage, _hWnd.get(), WndMsg::kFindSearchProgress, WndMsg::kFindSearchProgress, PM_REMOVE))
        {
            break;
        }

        auto newerPayload = TakeMessagePayload<FindSearchProgressPayload>(queuedMessage.lParam);
        if (! newerPayload)
        {
            continue;
        }

        payload = std::move(newerPayload);
        ++drainedProgressCount;
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
    _lastMatchedEntries     = payload->matchedEntries;
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

    RebuildResultsList();
}

void FindFilesWindow::OnSearchComplete(std::unique_ptr<FindSearchCompletePayload> payload) noexcept
{
    if (! payload)
    {
        return;
    }

    const Debug::Perf::Scope completePerf(L"find.ui.complete_handler_ms");
    if (payload->backend != FILESYSTEM_SEARCH_BACKEND_UNKNOWN || _lastBackend == FILESYSTEM_SEARCH_BACKEND_UNKNOWN)
    {
        _lastBackend = payload->backend;
    }
    _lastPhase              = FILESYSTEM_SEARCH_PHASE_COMPLETED;
    _lastWarningFlags       = payload->warningFlags;
    _lastScannedDirectories = payload->scannedDirectories;
    _lastScannedFiles       = payload->scannedFiles;
    _lastCandidateFiles     = payload->candidateFiles;
    _lastMatchedEntries     = payload->matchedEntries;
    _lastStatusHint         = payload->hr;
    _cancelRequestedUi      = false;

    if (FAILED(payload->hr) && payload->hr != HRESULT_FROM_WIN32(ERROR_CANCELLED))
    {
        Debug::Warning(
            L"FindFiles: search completed with failure backend={} hr=0x{:08X} scannedDirs={} scannedFiles={} candidates={} matches={} warnings=0x{:08X}",
            static_cast<uint32_t>(_lastBackend),
            static_cast<unsigned long>(payload->hr),
            _lastScannedDirectories,
            _lastScannedFiles,
            _lastCandidateFiles,
            _lastMatchedEntries,
            static_cast<unsigned long>(_lastWarningFlags));
    }

    if (_activeOperation == SearchOperation::Intersect || _activeOperation == SearchOperation::Subtract)
    {
        ApplyDeferredSetOperation(_activeOperation);
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
}

#ifdef ENABLE_TESTS
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
        case FindFilesDebugFocusTarget::ResultsGrid: return false;
        default: return false;
    }

    if (! combo)
    {
        return false;
    }

    const std::wstring expectedText = text;
    _dxHost.CommitFocusedTextInput();
    combo->SetText(std::move(text));
    SetFocus(_hWnd.get());
    _dxHost.SetFocusControl(combo);
    _dxHost.SyncTextInput(combo);
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
    return _debugLastSetComboObservedText == expectedText;
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

FindFilesDebugFocusTarget FindFilesWindow::ResolveDebugFocusTarget() const noexcept
{
    const RedSalamander::DxUi::Control* const focused = _dxHost.GetFocusControl();
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
    out                         = {};
    out.searchActive            = _session.IsActive();
    out.usesDxUiHost            = _dxHost.GetHwnd() == _hWnd.get();
    out.findButtonEnabled       = _findButton ? _findButton->IsEnabled() : false;
    out.appendButtonEnabled     = _appendButton ? _appendButton->IsEnabled() : false;
    out.intersectButtonEnabled  = _intersectButton ? _intersectButton->IsEnabled() : false;
    out.subtractButtonEnabled   = _subtractButton ? _subtractButton->IsEnabled() : false;
    out.cancelButtonEnabled     = _cancelButton ? _cancelButton->IsEnabled() : false;
    out.openButtonEnabled       = _openButton ? _openButton->IsEnabled() : false;
    out.parentButtonEnabled     = _parentButton ? _parentButton->IsEnabled() : false;
    out.rootComboEnabled        = _rootCombo ? _rootCombo->IsEnabled() : false;
    out.nameComboEnabled        = _nameCombo ? _nameCombo->IsEnabled() : false;
    out.nameModeComboEnabled    = _nameModeCombo ? _nameModeCombo->IsEnabled() : false;
    out.contentComboEnabled     = _contentCombo ? _contentCombo->IsEnabled() : false;
    out.contentModeComboEnabled = _contentModeCombo ? _contentModeCombo->IsEnabled() : false;
    out.matchCaseContentEnabled = _matchCaseContentCheck ? _matchCaseContentCheck->IsEnabled() : false;
    out.wantSnippetsEnabled     = _wantSnippetsCheck ? _wantSnippetsCheck->IsEnabled() : false;
    out.recursiveChecked        = _recursiveCheck ? _recursiveCheck->IsChecked() : false;
    out.resultCount             = _results.size();
    out.selectedResultCount     = _resultsList ? _resultsList->GetSelectionModel().GetCount() : 0u;
    out.visibleChildWindowCount = 0u;
    out.hasStatusStrip          = _statusText != nullptr;
    out.statusStripVisible      = _statusText ? _statusText->IsVisible() : false;
    out.statusStripSectionCount = _statusText ? static_cast<uint32_t>(_statusText->GetSectionCount()) : 0u;
    if (_statusText)
    {
        const D2D1_RECT_F statusBounds = _statusText->GetBounds();
        out.statusStripHeightDip       = (statusBounds.bottom > statusBounds.top) ? (statusBounds.bottom - statusBounds.top) : 0.0f;
    }
    out.resultColumnIds.clear();
    out.resultColumnWidthsDip.clear();
    out.rootPopupOpen                  = IsComboPopupOpen(_rootCombo);
    out.nameModePopupOpen              = IsComboPopupOpen(_nameModeCombo);
    out.contentModePopupOpen           = IsComboPopupOpen(_contentModeCombo);
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
    out.resultListFullRebuildCount  = _debugResultListFullRebuildCount;
    out.dxRenderCount               = _dxHost.DebugGetRenderCount();
    out.resultGridPaintCount        = _resultsList ? _resultsList->DebugGetPaintCount() : 0u;
    out.dxResizeCount               = _dxHost.DebugGetResizeCount();
    out.dxResizeFailureCount        = _dxHost.DebugGetResizeFailureCount();
    out.debugResizeBeforeWidthDip   = _debugResizeBeforeWidthDip;
    out.debugResizeTargetWidthDip   = _debugResizeTargetWidthDip;
    out.debugResizeObservedWidthDip = _debugResizeObservedWidthDip;
    out.debugResizeSucceeded        = _debugResizeSucceeded;
#else
    out.resultListFullRebuildCount = 0;
#endif
    out.statusText        = _status;
    out.backendStatusText = BuildBackendStatusText();
    if (_resultsList)
    {
        const auto metrics                 = _resultsList->GetVisibleWorkMetrics();
        out.visibleResultRowCount          = metrics.visibleRowCount;
        out.visibleResultColumnCount       = metrics.visibleColumnCount;
        out.visibleResultCellCount         = metrics.visibleCellCount;
        out.resultListHasVerticalScrollbar = metrics.hasVerticalScrollbar;
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
    for (const auto& record : _results)
    {
        out.fullPaths.push_back(record.fullPath);
    }

    if (_hWnd)
    {
        struct VisibleChildCounter
        {
            size_t count = 0u;
        } counter{};

        static_cast<void>(::EnumChildWindows(_hWnd.get(),
                                             [](HWND child, LPARAM lParam) noexcept -> BOOL
        {
            auto& counterRef = *reinterpret_cast<VisibleChildCounter*>(lParam);
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
    outRect                                     = RECT{};
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

LRESULT FindFilesWindow::WindowProc(UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    bool dxHandled         = false;
    const LRESULT dxResult = _dxHost.HandleMessage(_hWnd.get(), message, wParam, lParam, dxHandled);
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
            ApplyTheme();
            Layout();
            return 0;
        }
        case WM_NCACTIVATE: ApplyTitleBarTheme(_hWnd.get(), _theme, wParam != FALSE); return DefWindowProcW(_hWnd.get(), message, wParam, lParam);
        case WM_CLOSE:
            OnClose();
            _hWnd.reset();
            return 0;
        case WndMsg::kFindSearchResults: OnSearchResults(TakeMessagePayload<FindSearchResultsPayload>(lParam)); return 0;
        case WndMsg::kFindSearchProgress: OnSearchProgress(TakeMessagePayload<FindSearchProgressPayload>(lParam)); return 0;
        case WndMsg::kFindSearchComplete: OnSearchComplete(TakeMessagePayload<FindSearchCompletePayload>(lParam)); return 0;
        case WndMsg::kFindSearchDeferredRefresh: ApplyPendingResultsRefresh(); return 0;
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

bool ShowFindFilesWindow(HWND owner, Common::Settings::Settings& settings, const AppTheme& theme, FindFilesPaneContext context) noexcept
{
    if (g_findFilesWindow && g_findFilesWindow->GetHwnd() && IsWindow(g_findFilesWindow->GetHwnd()))
    {
        g_findFilesWindow->UpdateOwnerWindow(owner);
        g_findFilesWindow->UpdateTheme(theme);
        g_findFilesWindow->UpdateContext(std::move(context));
        ShowWindow(g_findFilesWindow->GetHwnd(), IsIconic(g_findFilesWindow->GetHwnd()) ? SW_RESTORE : SW_SHOWNORMAL);
        SetForegroundWindow(g_findFilesWindow->GetHwnd());
        return true;
    }

    auto window = std::unique_ptr<FindFilesWindow>(new (std::nothrow) FindFilesWindow(owner, settings, theme, std::move(context)));
    if (! window)
    {
        return false;
    }

    if (! window->Create())
    {
        return false;
    }

    static_cast<void>(window.release());
    return true;
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

#ifdef ENABLE_TESTS
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

bool DebugActivateSelectedFindFilesWindowResult() noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugActivateSelectedResult() : false;
}

bool DebugOpenSelectedFindFilesWindowResultParent() noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugOpenSelectedResultParent() : false;
}

bool DebugScrollFindFilesWindowResultsByWheelDetents(int detents) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugScrollResultsByWheelDetents(detents) : false;
}

bool DebugWaitForFindFilesWindowIdle(uint32_t timeoutMs) noexcept
{
    return g_findFilesWindow ? g_findFilesWindow->DebugWaitForIdle(timeoutMs) : true;
}
#endif
