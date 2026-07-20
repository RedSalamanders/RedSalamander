#include "ChangeCase.h"
#include "ConnectionManagerWindow.h"
#include "ConnectionSecrets.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "FileActionLauncher.h"
#include "FileActionResolver.h"
#include "FolderWindow.FileSystem.Private.h"
#include "FolderWindowInternal.h"
#include "Helpers.h"
#include "HostServices.h"
#include "LocalFileTransaction.h"
#include "MaskSyntax.h"
#include "NavigationLocation.h"
#include "PathUtils.h"

#include "SettingsStore.h"
#include "ViewerPluginManager.h"
#include "Win32CallbackHelpers.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#endif
#include "UiMetrics.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstddef>
#include <cstdlib>
#include <cstring>
#include <cwctype>
#include <fstream>
#include <limits>
#include <map>
#include <span>
#include <string>
#include <system_error>
#include <unordered_set>
#include <utility>
#include <vector>

#include <commdlg.h>
#include <lm.h>
#include <oleauto.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <winnetwk.h>

#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "Netapi32.lib")
#pragma comment(lib, "Shlwapi.lib")

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>

#pragma comment(lib, "OleAut32.lib")
#pragma warning(pop)

namespace FolderWindowFileSystemInternal
{
using OrdinalString::EqualsNoCase;
using OrdinalString::StartsWithNoCase;

bool IsFilePluginShortId(std::wstring_view pluginShortId) noexcept
{
    return EqualsNoCase(pluginShortId, L"file");
}

std::filesystem::path NormalizeLocalNavigationPath(std::filesystem::path path)
{
    path = path.lexically_normal();
    while (! path.empty() && ! path.has_filename() && path != path.root_path())
    {
        path = path.parent_path();
    }
    return path;
}

[[nodiscard]] bool IsSameLocalNavigationPathText(const std::filesystem::path& left, const std::filesystem::path& right)
{
    // Byte-wise comparison on purpose: a case-only navigation ('C:\foo' -> 'C:\FOO') must re-run the
    // navigation pipeline so the displayed casing refreshes.
    return NormalizeLocalNavigationPath(left).native() == NormalizeLocalNavigationPath(right).native();
}

[[nodiscard]] HWND GetClipboardOwnerWindow(HWND window) noexcept
{
    HWND ownerWindow = window ? GetAncestor(window, GA_ROOT) : nullptr;
    return ownerWindow ? ownerWindow : window;
}

[[nodiscard]] std::wstring NormalizeUncCandidatePath(std::wstring_view nativePath) noexcept
{
    static constexpr std::wstring_view kExtendedUncPrefix  = LR"(\\?\UNC\)";
    static constexpr std::wstring_view kExtendedPathPrefix = LR"(\\?\)";

    std::wstring normalized(nativePath);
    for (wchar_t& ch : normalized)
    {
        if (ch == L'/')
        {
            ch = L'\\';
        }
    }

    if (StartsWithNoCase(normalized, kExtendedUncPrefix))
    {
        std::wstring uncPath = LR"(\\)";
        uncPath.append(normalized.substr(kExtendedUncPrefix.size()));
        return uncPath;
    }

    if (StartsWithNoCase(normalized, kExtendedPathPrefix) && normalized.size() >= 6u && std::iswalpha(static_cast<wint_t>(normalized[4])) != 0 &&
        normalized[5] == L':')
    {
        return normalized.substr(kExtendedPathPrefix.size());
    }

    return normalized;
}

[[nodiscard]] std::wstring TryGetProviderUniversalPath(std::wstring_view normalizedPath) noexcept
{
    if (normalizedPath.empty())
    {
        return {};
    }

    const std::wstring path(normalizedPath);
    DWORD bufferBytes = 0;
    DWORD result      = WNetGetUniversalNameW(path.c_str(), UNIVERSAL_NAME_INFO_LEVEL, nullptr, &bufferBytes);
    if (result != ERROR_MORE_DATA || bufferBytes < sizeof(UNIVERSAL_NAME_INFOW))
    {
        return {};
    }

    std::vector<char> buffer(bufferBytes);
    auto* info = reinterpret_cast<UNIVERSAL_NAME_INFOW*>(buffer.data());
    result     = WNetGetUniversalNameW(path.c_str(), UNIVERSAL_NAME_INFO_LEVEL, info, &bufferBytes);
    if (result != NO_ERROR || info->lpUniversalName == nullptr || info->lpUniversalName[0] == L'\0')
    {
        return {};
    }

    return info->lpUniversalName;
}

[[nodiscard]] bool IsAbsoluteDrivePath(std::wstring_view path) noexcept
{
    return path.size() >= 3u && std::iswalpha(static_cast<wint_t>(path[0])) != 0 && path[1] == L':' && path[2] == L'\\';
}

[[nodiscard]] std::wstring TryGetLocalMachineAdministrativeSharePath(std::wstring_view normalizedPath) noexcept
{
    if (! IsAbsoluteDrivePath(normalizedPath))
    {
        return {};
    }

    const std::wstring driveRoot(normalizedPath.substr(0, 3));
    const UINT driveType = GetDriveTypeW(driveRoot.c_str());
    if (driveType == DRIVE_REMOTE || driveType == DRIVE_NO_ROOT_DIR || driveType == DRIVE_UNKNOWN)
    {
        return {};
    }

    wchar_t computerName[MAX_COMPUTERNAME_LENGTH + 1]{};
    DWORD computerNameLength = static_cast<DWORD>(std::size(computerName));
    if (GetComputerNameW(computerName, &computerNameLength) == FALSE || computerNameLength == 0u)
    {
        return {};
    }

    std::wstring uncPath = LR"(\\)";
    uncPath.append(computerName, computerNameLength);
    uncPath.push_back(L'\\');
    uncPath.push_back(static_cast<wchar_t>(std::towupper(static_cast<wint_t>(normalizedPath[0]))));
    uncPath.push_back(L'$');
    uncPath.append(normalizedPath.substr(2));
    return uncPath;
}

[[nodiscard]] std::wstring GetUniversalPathOrOriginal(std::wstring_view nativePath) noexcept
{
    if (nativePath.empty())
    {
        return {};
    }

    const std::wstring normalizedPath = NormalizeUncCandidatePath(nativePath);

    if (StartsWithNoCase(normalizedPath, LR"(\\)"))
    {
        return normalizedPath;
    }

    if (std::wstring providerUncPath = TryGetProviderUniversalPath(normalizedPath); ! providerUncPath.empty())
    {
        return providerUncPath;
    }

    if (std::wstring administrativeSharePath = TryGetLocalMachineAdministrativeSharePath(normalizedPath); ! administrativeSharePath.empty())
    {
        return administrativeSharePath;
    }

    return normalizedPath;
}

constexpr uint32_t kFolderHistoryMaxMax = 50u;

void NormalizeFolderHistory(std::vector<std::filesystem::path>& history, size_t maxItems)
{
    std::vector<std::filesystem::path> normalized;
    normalized.reserve(std::min(history.size(), maxItems));

    for (const auto& entry : history)
    {
        if (entry.empty())
        {
            continue;
        }

        const std::wstring_view entryText = entry.native();
        const bool exists                 = std::find_if(normalized.begin(), normalized.end(), [&](const std::filesystem::path& existing) {
            return OrdinalString::EqualsNoCasePath(existing, entryText);
        }) != normalized.end();
        if (exists)
        {
            continue;
        }

        normalized.push_back(entry);
        if (normalized.size() >= maxItems)
        {
            break;
        }
    }

    history = std::move(normalized);
}

void AddToFolderHistory(std::vector<std::filesystem::path>& history, size_t maxItems, const std::filesystem::path& entry)
{
    if (entry.empty() || maxItems == 0)
    {
        return;
    }

    const std::wstring_view entryText = entry.native();
    auto it                           = std::find_if(
        history.begin(), history.end(), [&](const std::filesystem::path& existing) { return OrdinalString::EqualsNoCasePath(existing, entryText); });

    if (it != history.end())
    {
        if (it == history.begin())
        {
            return;
        }

        std::filesystem::path moved = std::move(*it);
        history.erase(it);
        history.insert(history.begin(), std::move(moved));
        return;
    }

    history.insert(history.begin(), entry);
    if (history.size() > maxItems)
    {
        history.resize(maxItems);
    }
}

[[nodiscard]] FolderView::NameFilterState GetFolderHistoryFilterState(const Common::Settings::FoldersSettings* folders,
                                                                      const std::filesystem::path& displayPath)
{
    FolderView::NameFilterState result{};

    if (! folders || folders->historyFilters.empty() || displayPath.empty())
    {
        return result;
    }

    const auto it =
        std::find_if(folders->historyFilters.begin(), folders->historyFilters.end(), [&](const Common::Settings::FolderHistoryFilterState& state) noexcept {
        return ! state.path.empty() && OrdinalString::EqualsNoCasePath(state.path, displayPath);
    });
    if (it == folders->historyFilters.end())
    {
        return result;
    }

    result.enabled = it->enabled;
    result.text    = it->text;
    return result;
}

void SetFolderHistoryFilterState(Common::Settings::FoldersSettings& folders,
                                 const std::filesystem::path& displayPath,
                                 const FolderView::NameFilterState& filter)
{
    if (displayPath.empty())
    {
        return;
    }

    auto it = std::find_if(folders.historyFilters.begin(), folders.historyFilters.end(), [&](const Common::Settings::FolderHistoryFilterState& state) noexcept {
        return ! state.path.empty() && OrdinalString::EqualsNoCasePath(state.path, displayPath);
    });

    if (! filter.enabled && filter.text.empty())
    {
        if (it != folders.historyFilters.end())
        {
            folders.historyFilters.erase(it);
        }
        return;
    }

    if (it != folders.historyFilters.end())
    {
        it->path    = displayPath;
        it->enabled = filter.enabled;
        it->text    = filter.text;
        return;
    }

    Common::Settings::FolderHistoryFilterState state{};
    state.path    = displayPath;
    state.enabled = filter.enabled;
    state.text    = filter.text;
    folders.historyFilters.push_back(std::move(state));
}

void PruneFolderHistoryFilters(Common::Settings::FoldersSettings& folders, const std::vector<std::filesystem::path>& history, size_t maxItems)
{
    if (maxItems == 0 || history.empty() || folders.historyFilters.empty())
    {
        folders.historyFilters.clear();
        return;
    }

    std::vector<Common::Settings::FolderHistoryFilterState> pruned;
    pruned.reserve(std::min(folders.historyFilters.size(), maxItems));

    for (const auto& historyPath : history)
    {
        if (historyPath.empty())
        {
            continue;
        }

        const auto it =
            std::find_if(folders.historyFilters.begin(), folders.historyFilters.end(), [&](const Common::Settings::FolderHistoryFilterState& state) noexcept {
            return ! state.path.empty() && OrdinalString::EqualsNoCasePath(state.path, historyPath);
        });
        if (it == folders.historyFilters.end())
        {
            continue;
        }

        if (! it->enabled && it->text.empty())
        {
            continue;
        }

        Common::Settings::FolderHistoryFilterState state = *it;
        state.path                                       = historyPath;
        pruned.push_back(std::move(state));

        if (pruned.size() >= maxItems)
        {
            break;
        }
    }

    folders.historyFilters = std::move(pruned);
}

bool LooksLikeWindowsAbsolutePath(std::wstring_view text) noexcept
{
    const Common::Paths::WindowsPathClass pathClass = Common::Paths::ClassifyWindowsPath(text);
    return pathClass != Common::Paths::WindowsPathClass::Relative && pathClass != Common::Paths::WindowsPathClass::Rooted;
}

std::filesystem::path GetDefaultFileSystemRoot() noexcept
{
    wchar_t buffer[MAX_PATH] = {};
    const UINT bufferSize    = static_cast<UINT>(ARRAYSIZE(buffer));
    const UINT length        = GetWindowsDirectoryW(buffer, bufferSize);
    if (length > 0 && length < bufferSize)
    {
        const std::filesystem::path root = std::filesystem::path(buffer).root_path();
        if (! root.empty())
        {
            return root;
        }
    }

    return std::filesystem::path(L"C:\\");
}

bool IsValidPluginIdPrefix(std::wstring_view prefix) noexcept
{
    if (prefix.empty())
    {
        return false;
    }

    for (wchar_t ch : prefix)
    {
        if (std::iswalnum(ch) == 0)
        {
            return false;
        }
    }

    return true;
}

bool TryParsePluginPrefix(std::wstring_view text, std::wstring& outPluginId, std::wstring& outRemainder) noexcept
{
    outPluginId.clear();
    outRemainder.clear();

    if (text.empty())
    {
        return false;
    }

    const size_t colon = text.find(L':');
    if (colon == std::wstring_view::npos || colon < 1)
    {
        return false;
    }

    if (colon == 1u && std::iswalpha(static_cast<wint_t>(text[0])) != 0)
    {
        // Avoid treating Windows drive-letter paths ("C:\...") as plugin prefixes.
        return false;
    }

    const size_t sep = text.find_first_of(L"\\/");
    if (sep != std::wstring_view::npos && sep < colon)
    {
        return false;
    }

    const std::wstring_view prefix = text.substr(0, colon);
    if (! IsValidPluginIdPrefix(prefix))
    {
        return false;
    }

    outPluginId.assign(prefix);
    outRemainder.assign(text.substr(colon + 1));
    return true;
}

const FileSystemPluginManager::PluginEntry* FindPluginByShortId(const std::vector<FileSystemPluginManager::PluginEntry>& plugins,
                                                                std::wstring_view shortId) noexcept
{
    if (shortId.empty())
    {
        return nullptr;
    }

    const size_t idSize = shortId.size();
    if (idSize > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return nullptr;
    }

    for (const auto& entry : plugins)
    {
        if (entry.shortId.empty())
        {
            continue;
        }

        if (entry.shortId.size() != idSize)
        {
            continue;
        }

        if (EqualsNoCase(entry.shortId, shortId))
        {
            return &entry;
        }
    }

    return nullptr;
}

const FileSystemPluginManager::PluginEntry* FindPluginById(const std::vector<FileSystemPluginManager::PluginEntry>& plugins,
                                                           std::wstring_view pluginId) noexcept
{
    if (pluginId.empty())
    {
        return nullptr;
    }

    const size_t idSize = pluginId.size();
    if (idSize > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return nullptr;
    }

    for (const auto& entry : plugins)
    {
        if (entry.id.empty())
        {
            continue;
        }

        if (entry.id.size() != idSize)
        {
            continue;
        }

        if (EqualsNoCase(entry.id, pluginId))
        {
            return &entry;
        }
    }

    return nullptr;
}

HWND GetOwnerWindowOrSelf(HWND window) noexcept
{
    if (! window)
    {
        return nullptr;
    }

    HWND rootWindow = GetAncestor(window, GA_ROOT);
    if (rootWindow)
    {
        return rootWindow;
    }

    return window;
}

[[nodiscard]] bool IsActuallyVisibleChildWindow(HWND hwnd) noexcept
{
    if (! hwnd || IsWindowVisible(hwnd) == FALSE)
    {
        return false;
    }

    // DxUi text bridges stay WS_VISIBLE for IME routing, but an empty region keeps them off-screen.
    wil::unique_hrgn region(CreateRectRgn(0, 0, 0, 0));
    if (region)
    {
        const int rgnType = GetWindowRgn(hwnd, region.get());
        if (rgnType == NULLREGION)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] size_t CountVisibleChildWindowsLocal(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    size_t count = 0u;
    EnumChildWindows(hwnd,
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto* countPtr = reinterpret_cast<size_t*>(lParam);
        if (countPtr && IsActuallyVisibleChildWindow(child))
        {
            ++(*countPtr);
        }
        return TRUE;
    },
                     reinterpret_cast<LPARAM>(&count));
    return count;
}

#ifdef ENABLE_TESTS

[[nodiscard]] std::wstring GetWindowClassNameLocal(HWND hwnd)
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return {};
    }

    std::array<wchar_t, 128> className{};
    const int length = GetClassNameW(hwnd, className.data(), static_cast<int>(className.size()));
    if (length <= 0)
    {
        return {};
    }

    return std::wstring(className.data(), static_cast<size_t>(length));
}

[[nodiscard]] bool IsNativeDialogTemplateControlClass(std::wstring_view className) noexcept
{
    return _wcsicmp(className.data(), L"#32770") == 0 || _wcsicmp(className.data(), L"Button") == 0 || _wcsicmp(className.data(), L"Edit") == 0 ||
           _wcsicmp(className.data(), L"Static") == 0 || _wcsicmp(className.data(), L"ComboBox") == 0;
}

[[nodiscard]] size_t CountVisibleNativeChildControlWindowsLocal(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    size_t count = 0u;
    EnumChildWindows(hwnd,
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto* countPtr = reinterpret_cast<size_t*>(lParam);
        if (! countPtr || ! IsActuallyVisibleChildWindow(child))
        {
            return TRUE;
        }

        std::array<wchar_t, 64> className{};
        const int length = GetClassNameW(child, className.data(), static_cast<int>(className.size()));
        if (length > 0 && IsNativeDialogTemplateControlClass(std::wstring_view(className.data(), static_cast<size_t>(length))))
        {
            ++(*countPtr);
        }
        return TRUE;
    },
                     reinterpret_cast<LPARAM>(&count));
    return count;
}
#endif

[[nodiscard]] int ScalePanePromptForDpi(const UINT dpi, const int dip) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi == 0u ? 96u : dpi), 96);
}

constexpr float kPaneFilterPromptMarginDip       = 16.0f;
constexpr float kPaneFilterPromptGapDip          = 8.0f;
constexpr float kPaneFilterPromptRowHeightDip    = 34.0f;
constexpr float kPaneFilterPromptLabelHeightDip  = 22.0f;
constexpr float kPaneFilterPromptToggleWidthDip  = 112.0f;
constexpr float kPaneFilterPromptHintHeightDip   = 28.0f;
constexpr float kPaneFilterPromptHelpHeightDip   = 112.0f;
constexpr float kPaneFilterPromptButtonWidthDip  = 96.0f;
constexpr float kPaneFilterPromptButtonHeightDip = 34.0f;

[[nodiscard]] constexpr int ResolvePaneFilterPromptClientHeightDip(bool helpExpanded) noexcept
{
    float height = kPaneFilterPromptMarginDip + kPaneFilterPromptRowHeightDip + kPaneFilterPromptGapDip + kPaneFilterPromptRowHeightDip +
                   kPaneFilterPromptGapDip + kPaneFilterPromptHintHeightDip + kPaneFilterPromptGapDip;
    if (helpExpanded)
    {
        height += kPaneFilterPromptHelpHeightDip + kPaneFilterPromptGapDip;
    }
    height += kPaneFilterPromptButtonHeightDip + kPaneFilterPromptMarginDip;
    return static_cast<int>(height);
}

constexpr wchar_t kFolderViewPaneFilterPromptClassName[]      = L"RedSalamander.FolderView.PaneFilterPrompt";
constexpr wchar_t kFolderViewSelectionMaskPromptClassName[]   = L"RedSalamander.FolderView.SelectionMaskPrompt";
constexpr wchar_t kFolderViewChangeCasePromptClassName[]      = L"RedSalamander.FolderView.ChangeCasePrompt";
constexpr wchar_t kFolderViewCreateDirectoryPromptClassName[] = L"RedSalamander.FolderView.CreateDirectoryPrompt";
constexpr wchar_t kFolderViewEditNewPromptClassName[]         = L"RedSalamander.FolderView.EditNewPrompt";

#ifdef ENABLE_TESTS
enum class FolderViewPaneFilterPromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetEnabled,
    SetText,
    SetTextAndNotify,
    SetHelpExpanded,
    Confirm,
    Cancel,
};

enum class FolderViewSelectionMaskPromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetText,
    Confirm,
    Cancel,
};

enum class FolderViewChangeCasePromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetSelections,
    Confirm,
    Cancel,
};

enum class FolderViewCreateDirectoryPromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetText,
    Confirm,
    Cancel,
};

enum class FolderViewEditNewPromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetText,
    SelectEditor,
    Confirm,
    Cancel,
};

[[nodiscard]] UINT GetFolderViewCreateDirectoryPromptDebugMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"RedSalamander.FolderView.CreateDirectoryPrompt.Debug");
    return message;
}

[[nodiscard]] UINT GetFolderViewEditNewPromptDebugMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"RedSalamander.FolderView.EditNewPrompt.Debug");
    return message;
}

struct FolderViewChangeCasePromptDebugSelectionPayload final
{
    size_t styleIndex   = 0u;
    size_t targetIndex  = 0u;
    bool includeSubdirs = false;
};

[[nodiscard]] constexpr LPARAM PackFolderViewChangeCasePromptSelections(size_t styleIndex, size_t targetIndex, bool includeSubdirs) noexcept
{
    return static_cast<LPARAM>((static_cast<uint64_t>(styleIndex) & 0xffffull) | ((static_cast<uint64_t>(targetIndex) & 0xffffull) << 16) |
                               ((includeSubdirs ? 1ull : 0ull) << 32));
}

[[nodiscard]] constexpr FolderViewChangeCasePromptDebugSelectionPayload UnpackFolderViewChangeCasePromptSelections(const LPARAM packed) noexcept
{
    FolderViewChangeCasePromptDebugSelectionPayload payload{};
    const uint64_t value   = static_cast<uint64_t>(packed);
    payload.styleIndex     = static_cast<size_t>(value & 0xffffull);
    payload.targetIndex    = static_cast<size_t>((value >> 16) & 0xffffull);
    payload.includeSubdirs = ((value >> 32) & 0x1ull) != 0;
    return payload;
}

std::mutex g_folderViewPaneFilterPromptDebugMutex;
std::optional<FolderViewPaneFilterPromptDebugSnapshot> g_folderViewPaneFilterPromptDebugSnapshot;
std::wstring g_folderViewPaneFilterPromptDebugText;
#endif

[[nodiscard]] std::vector<std::wstring> BuildPromptHistoryEntries(const std::vector<std::wstring>& history) noexcept
{
    std::vector<std::wstring> entries;
    entries.reserve(history.size());
    for (const std::wstring& rawEntry : history)
    {
        std::wstring entry = StringUtils::TrimWhitespaceCopy(rawEntry);
        if (entry.empty())
        {
            continue;
        }

        const bool duplicate = std::ranges::any_of(entries, [&](const std::wstring& existing) noexcept { return existing == entry; });
        if (! duplicate)
        {
            entries.push_back(std::move(entry));
        }
    }

    return entries;
}

template <typename OnSelected>
void TryShowPromptHistoryMenu(HWND ownerWindow,
                              RedSalamander::DxUi::WindowHost& host,
                              const RedSalamander::DxUi::ThemePalette& palette,
                              const D2D1_RECT_F& anchorBounds,
                              const std::vector<std::wstring>& history,
                              std::wstring_view currentText,
                              OnSelected&& onSelected) noexcept
{
    if (! ownerWindow || IsWindow(ownerWindow) == FALSE)
    {
        return;
    }

    std::vector<std::wstring> entries = BuildPromptHistoryEntries(history);
    if (entries.empty())
    {
        return;
    }

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(entries.size());
    for (size_t index = 0; index < entries.size(); ++index)
    {
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.text      = entries[index];
        item.checked   = entries[index] == currentText;
        item.commandId = static_cast<int>(index) + 1;
        items.push_back(std::move(item));
    }

    const POINT screenPoint = host.DipPointToScreenPoint(D2D1::Point2F(anchorBounds.left, anchorBounds.bottom));
    const auto result       = RedSalamander::DxUi::ContextMenu::Show(ownerWindow, screenPoint, items, palette);
    if (! result.has_value() || result.value() <= 0)
    {
        return;
    }

    const size_t selectedIndex = static_cast<size_t>(result.value() - 1);
    if (selectedIndex >= entries.size())
    {
        return;
    }

    onSelected(entries[selectedIndex]);
}

class FolderViewPaneFilterPromptWindow final
{
public:
    FolderViewPaneFilterPromptWindow(const FolderViewPaneFilterPromptWindow&)            = delete;
    FolderViewPaneFilterPromptWindow& operator=(const FolderViewPaneFilterPromptWindow&) = delete;
    FolderViewPaneFilterPromptWindow(FolderViewPaneFilterPromptWindow&&)                 = delete;
    FolderViewPaneFilterPromptWindow& operator=(FolderViewPaneFilterPromptWindow&&)      = delete;

    FolderViewPaneFilterPromptWindow(HWND ownerWindow,
                                     std::vector<std::wstring> history,
                                     const AppTheme& theme,
                                     const FolderView::NameFilterState& initial) noexcept
        : _ownerWindow(GetOwnerWindowOrSelf(ownerWindow)),
          _history(std::move(history)),
          _theme(theme),
          _enabled(initial.enabled),
          _initialText(StringUtils::TrimWhitespaceCopy(initial.text))
    {
    }

    [[nodiscard]] std::optional<FolderView::NameFilterState> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 452), ResolveWindowClientHeightPx(dpi)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
        const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled] noexcept
        {
            if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
            {
                EnableWindow(_ownerWindow, TRUE);
                SetActiveWindow(_ownerWindow);
            }
        });

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kFolderViewPaneFilterPromptClassName,
                                          LoadStringResource(nullptr, IDS_CAPTION_PANE_FILTER).c_str(),
                                          style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(_hWnd.get(), _ownerWindow));
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return std::nullopt;
            }
            if (getMessageResult == 0)
            {
                _done = true;
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<FolderViewPaneFilterPromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<FolderViewPaneFilterPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            if (message == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                self->_done = true;
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
#ifdef ENABLE_TESTS
            case WndMsg::kFolderViewPaneFilterPromptDebug: return self->OnDebugCommand(static_cast<FolderViewPaneFilterPromptDebugCommand>(wParam), lParam);
#endif
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = FolderViewPaneFilterPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kFolderViewPaneFilterPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] int ResolveWindowClientHeightPx(const UINT dpi) const noexcept
    {
        return ScalePanePromptForDpi(dpi, ResolvePaneFilterPromptClientHeightDip(_helpExpanded));
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        ApplyTheme();
        UpdateHintUi();
        if (_filterCombo)
        {
            _filterCombo->SetText(_initialText);
        }
        Layout();
        _dxHost.SetFocusControl(_filterCombo);
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi() noexcept
    {
        if (_root != nullptr)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _useLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER_USE_FILTER));
        _useLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _toggle = _root->AddChild<Toggle>();
        _toggle->SetChecked(_enabled);
        _toggle->SetAccessibleName(LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER_USE_FILTER));
        _toggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
        _toggle->SetOnToggled([this](bool checked)
        {
            _enabled = checked;
            if (_toggle)
            {
                _toggle->SetChecked(checked);
            }
        });

        std::vector<ComboBox::Item> historyItems;
        const std::vector<std::wstring> historyEntries = BuildPromptHistoryEntries(_history);
        historyItems.reserve(historyEntries.size());
        for (const std::wstring& entry : historyEntries)
        {
            historyItems.push_back(ComboBox::Item{entry, entry});
        }

        _filterCombo = _root->AddChild<ComboBox>();
        _filterCombo->SetEditable(true);
        _filterCombo->SetVariant(ComboBoxVariant::Edit);
        _filterCombo->SetAutoOpenOnTextInput(false);
        _filterCombo->SetMaxVisibleItems(MaskSyntax::kWildcardMaskHistoryMaxItems);
        _filterCombo->SetItems(std::move(historyItems));
        _filterCombo->SetPlaceholder(LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER));
        _filterCombo->SetAccessibleName(LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER));
        _filterCombo->SetOnTextChanged([this](std::wstring_view text) noexcept { SyncEnabledForFilterText(text); });
        _filterCombo->SetOnSubmitted([this] { Confirm(); });

        _hintButton = _root->AddChild<Button>(L"");
        _hintButton->SetOnClick([this]
        {
            _helpExpanded = ! _helpExpanded;
            UpdateHintUi();
            ResizeForHelpState();
        });

        _helpLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_SELECTION_MASK_HELP_TEXT));
        _helpLabel->SetMultiline(true);
        _helpLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_hWnd)
        {
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
        }
    }

    void UpdateHintUi() noexcept
    {
        if (_hintButton)
        {
            _hintButton->SetText(LoadStringResource(nullptr, _helpExpanded ? IDS_SELECTION_MASK_HINT_EXPANDED : IDS_SELECTION_MASK_HINT_COLLAPSED));
        }
        if (_helpLabel)
        {
            _helpLabel->SetVisible(_helpExpanded);
        }
    }

    void ResizeForHelpState() noexcept
    {
        if (! _hWnd || IsWindow(_hWnd.get()) == FALSE)
        {
            return;
        }

        RECT windowRect{};
        RECT clientRect{};
        if (GetWindowRect(_hWnd.get(), &windowRect) == FALSE || GetClientRect(_hWnd.get(), &clientRect) == FALSE)
        {
            return;
        }

        const UINT dpi                = GetDpiForWindow(_hWnd.get());
        const int currentClientHeight = std::max(0l, clientRect.bottom - clientRect.top);
        const int nonClientHeight     = std::max<int>(0, static_cast<int>((windowRect.bottom - windowRect.top) - currentClientHeight));
        const int targetWindowHeight  = ResolveWindowClientHeightPx(dpi) + nonClientHeight;

        SetWindowPos(_hWnd.get(), nullptr, 0, 0, windowRect.right - windowRect.left, targetWindowHeight, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        Layout();
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        const float left         = client.left + kPaneFilterPromptMarginDip;
        const float right        = std::max(left, client.right - kPaneFilterPromptMarginDip);
        const float contentWidth = std::max(0.0f, right - left);
        float y                  = client.top + kPaneFilterPromptMarginDip;

        if (_useLabel)
        {
            _useLabel->SetBounds(D2D1::RectF(
                left, y + 6.0f, std::max(left, right - kPaneFilterPromptToggleWidthDip - kPaneFilterPromptGapDip), y + 6.0f + kPaneFilterPromptLabelHeightDip));
        }
        if (_toggle)
        {
            _toggle->SetBounds(D2D1::RectF(std::max(left, right - kPaneFilterPromptToggleWidthDip), y, right, y + kPaneFilterPromptRowHeightDip));
        }
        y += kPaneFilterPromptRowHeightDip + kPaneFilterPromptGapDip;

        if (_filterCombo)
        {
            _filterCombo->SetBounds(D2D1::RectF(left, y, right, y + kPaneFilterPromptRowHeightDip));
        }
        y += kPaneFilterPromptRowHeightDip + kPaneFilterPromptGapDip;

        if (_hintButton)
        {
            _hintButton->SetBounds(D2D1::RectF(left, y, std::min(right, left + std::min(contentWidth, 160.0f)), y + kPaneFilterPromptHintHeightDip));
        }
        y += kPaneFilterPromptHintHeightDip + kPaneFilterPromptGapDip;

        if (_helpLabel)
        {
            _helpLabel->SetBounds(D2D1::RectF(left, y, right, y + (_helpExpanded ? kPaneFilterPromptHelpHeightDip : 0.0f)));
        }
        if (_helpExpanded)
        {
            y += kPaneFilterPromptHelpHeightDip + kPaneFilterPromptGapDip;
        }

        const float buttonsTop = std::max(y, client.bottom - kPaneFilterPromptMarginDip - kPaneFilterPromptButtonHeightDip);
        const float cancelLeft = std::max(left, right - kPaneFilterPromptButtonWidthDip);
        const float okLeft     = std::max(left, cancelLeft - kPaneFilterPromptGapDip - kPaneFilterPromptButtonWidthDip);

        if (_okButton)
        {
            _okButton->SetBounds(D2D1::RectF(okLeft, buttonsTop, okLeft + kPaneFilterPromptButtonWidthDip, buttonsTop + kPaneFilterPromptButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(
                D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kPaneFilterPromptButtonWidthDip, buttonsTop + kPaneFilterPromptButtonHeightDip));
        }
    }

    void SetFilterEnabled(bool enabled) noexcept
    {
        _enabled = enabled;
        if (_toggle)
        {
            _toggle->SetChecked(enabled);
        }
    }

    void SyncEnabledForFilterText(std::wstring_view text) noexcept
    {
        SetFilterEnabled(! StringUtils::TrimWhitespaceCopy(std::wstring(text)).empty());
    }

    void Confirm() noexcept
    {
        const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(_filterCombo ? std::wstring(_filterCombo->GetText()) : std::wstring{});
        if (trimmed.empty())
        {
            SetFilterEnabled(false);
        }

        if (_enabled)
        {
            const MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(trimmed);
            const bool hasMask                  = ! mask.includePatterns.empty() || ! mask.excludePatterns.empty();
            if (! hasMask)
            {
                MessageBeep(MB_ICONWARNING);
                if (_filterCombo)
                {
                    _dxHost.SetFocusControl(_filterCombo);
                }
                return;
            }
        }

        _result = FolderView::NameFilterState{_enabled, trimmed};
        _done   = true;
        ClosePromptWindow();
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        ClosePromptWindow();
    }

    void ClosePromptWindow() noexcept
    {
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _dxHost.SetFocusControl(nullptr);
            _hWnd.reset();
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(FolderViewPaneFilterPromptDebugCommand command, LPARAM lParam) noexcept
    {
        switch (command)
        {
            case FolderViewPaneFilterPromptDebugCommand::GetSnapshot:
            {
                FolderViewPaneFilterPromptDebugSnapshot snapshot{};
                snapshot.usesDxUiHost            = _dxHost.GetRoot() != nullptr;
                snapshot.visibleChildWindowCount = CountVisibleChildWindowsLocal(_hWnd.get());
                snapshot.enabled                 = _enabled;
                snapshot.helpExpanded            = _helpExpanded;
                snapshot.historyComboVisible     = _filterCombo && _filterCombo->IsVisible();
                snapshot.historyButtonVisible    = false;
                snapshot.historyItemCount        = BuildPromptHistoryEntries(_history).size();
                snapshot.text                    = _filterCombo ? std::wstring(_filterCombo->GetText()) : std::wstring{};
                const D2D1_RECT_F client         = _dxHost.GetClientBoundsDip();
                snapshot.clientBottomDip         = client.bottom;
                if (_okButton)
                {
                    snapshot.okButtonBottomDip = _okButton->GetBounds().bottom;
                }
                if (_cancelButton)
                {
                    snapshot.cancelButtonBottomDip = _cancelButton->GetBounds().bottom;
                }
                snapshot.commandButtonsFitInClient =
                    snapshot.okButtonBottomDip <= snapshot.clientBottomDip && snapshot.cancelButtonBottomDip <= snapshot.clientBottomDip;
                {
                    const std::scoped_lock lock(g_folderViewPaneFilterPromptDebugMutex);
                    g_folderViewPaneFilterPromptDebugSnapshot = std::move(snapshot);
                }
                return TRUE;
            }
            case FolderViewPaneFilterPromptDebugCommand::SetEnabled:
                _enabled = lParam != 0;
                if (_toggle)
                {
                    _toggle->SetChecked(_enabled);
                }
                return TRUE;
            case FolderViewPaneFilterPromptDebugCommand::SetText:
            {
                if (! _filterCombo)
                {
                    return FALSE;
                }

                std::wstring text;
                {
                    const std::scoped_lock lock(g_folderViewPaneFilterPromptDebugMutex);
                    text = g_folderViewPaneFilterPromptDebugText;
                }

                _filterCombo->SetText(std::move(text));
                _dxHost.SetFocusControl(_filterCombo);
                return TRUE;
            }
            case FolderViewPaneFilterPromptDebugCommand::SetTextAndNotify:
            {
                if (! _filterCombo)
                {
                    return FALSE;
                }

                std::wstring text;
                {
                    const std::scoped_lock lock(g_folderViewPaneFilterPromptDebugMutex);
                    text = g_folderViewPaneFilterPromptDebugText;
                }

                _filterCombo->SetTextAndNotify(std::move(text));
                _dxHost.SetFocusControl(_filterCombo);
                return TRUE;
            }
            case FolderViewPaneFilterPromptDebugCommand::SetHelpExpanded:
                _helpExpanded = lParam != 0;
                UpdateHintUi();
                ResizeForHelpState();
                return TRUE;
            case FolderViewPaneFilterPromptDebugCommand::Confirm: Confirm(); return TRUE;
            case FolderViewPaneFilterPromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

    HWND _ownerWindow = nullptr;
    std::vector<std::wstring> _history;
    AppTheme _theme{};
    bool _enabled      = false;
    bool _helpExpanded = false;
    std::wstring _initialText;
    RedSalamander::DxUi::ThemePalette _palette{};
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root           = nullptr;
    RedSalamander::DxUi::Label* _useLabel       = nullptr;
    RedSalamander::DxUi::Toggle* _toggle        = nullptr;
    RedSalamander::DxUi::ComboBox* _filterCombo = nullptr;
    RedSalamander::DxUi::Button* _hintButton    = nullptr;
    RedSalamander::DxUi::Label* _helpLabel      = nullptr;
    RedSalamander::DxUi::Button* _okButton      = nullptr;
    RedSalamander::DxUi::Button* _cancelButton  = nullptr;
    bool _done                                  = false;
    std::optional<FolderView::NameFilterState> _result;
};

class FolderViewSelectionMaskPromptWindow final
{
public:
    FolderViewSelectionMaskPromptWindow(const FolderViewSelectionMaskPromptWindow&)            = delete;
    FolderViewSelectionMaskPromptWindow& operator=(const FolderViewSelectionMaskPromptWindow&) = delete;
    FolderViewSelectionMaskPromptWindow(FolderViewSelectionMaskPromptWindow&&)                 = delete;
    FolderViewSelectionMaskPromptWindow& operator=(FolderViewSelectionMaskPromptWindow&&)      = delete;

    FolderViewSelectionMaskPromptWindow(
        HWND ownerWindow, std::vector<std::wstring> history, const AppTheme& theme, std::wstring captionText, std::wstring labelText) noexcept
        : _ownerWindow(GetOwnerWindowOrSelf(ownerWindow)),
          _history(std::move(history)),
          _theme(theme),
          _captionText(std::move(captionText)),
          _labelText(std::move(labelText))
    {
    }

    [[nodiscard]] std::optional<std::wstring> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 420), ResolveWindowClientHeightPx(dpi)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
        const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled] noexcept
        {
            if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
            {
                EnableWindow(_ownerWindow, TRUE);
                SetActiveWindow(_ownerWindow);
            }
        });

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kFolderViewSelectionMaskPromptClassName,
                                          _captionText.c_str(),
                                          style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(_hWnd.get(), _ownerWindow));
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return std::nullopt;
            }
            if (getMessageResult == 0)
            {
                _done = true;
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<FolderViewSelectionMaskPromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<FolderViewSelectionMaskPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            if (message == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                self->_done = true;
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
#ifdef ENABLE_TESTS
            case WndMsg::kFolderViewSelectionMaskPromptDebug:
                return self->OnDebugCommand(static_cast<FolderViewSelectionMaskPromptDebugCommand>(wParam), lParam);
#endif
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = FolderViewSelectionMaskPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kFolderViewSelectionMaskPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] int ResolveWindowClientHeightPx(const UINT dpi) const noexcept
    {
        return ScalePanePromptForDpi(dpi, _helpExpanded ? 246 : 174);
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        ApplyTheme();
        UpdateHintUi();
        if (_textField)
        {
            if (! _history.empty() && ! _history.front().empty())
            {
                _textField->SetText(StringUtils::TrimWhitespaceCopy(_history.front()));
            }
        }
        Layout();
        if (_textField)
        {
            _dxHost.SetFocusControl(_textField);
        }
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi() noexcept
    {
        if (_root != nullptr)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _label = _root->AddChild<Label>(_labelText);
        _label->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _textField = _root->AddChild<TextField>();
        _textField->SetMultiline(false);
        _textField->SetAccessibleName(_labelText);
        _textField->SetOnSubmitted([this] { Confirm(); });

        _historyButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_HISTORY));
        _historyButton->SetVisible(! BuildPromptHistoryEntries(_history).empty());
        _historyButton->SetOnClick([this] { ShowHistoryMenu(); });

        _hintButton = _root->AddChild<Button>(L"");
        _hintButton->SetOnClick([this]
        {
            _helpExpanded = ! _helpExpanded;
            UpdateHintUi();
            ResizeForHelpState();
        });

        _helpLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_SELECTION_MASK_HELP_TEXT));
        _helpLabel->SetMultiline(true);
        _helpLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_hWnd)
        {
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
        }
    }

    void UpdateHintUi() noexcept
    {
        if (_hintButton)
        {
            _hintButton->SetText(LoadStringResource(nullptr, _helpExpanded ? IDS_SELECTION_MASK_HINT_EXPANDED : IDS_SELECTION_MASK_HINT_COLLAPSED));
        }
        if (_helpLabel)
        {
            _helpLabel->SetVisible(_helpExpanded);
        }
    }

    void ResizeForHelpState() noexcept
    {
        if (! _hWnd || IsWindow(_hWnd.get()) == FALSE)
        {
            return;
        }

        RECT windowRect{};
        RECT clientRect{};
        if (GetWindowRect(_hWnd.get(), &windowRect) == FALSE || GetClientRect(_hWnd.get(), &clientRect) == FALSE)
        {
            return;
        }

        const UINT dpi                = GetDpiForWindow(_hWnd.get());
        const int currentClientHeight = std::max(0l, clientRect.bottom - clientRect.top);
        const int nonClientHeight     = std::max<int>(0, static_cast<int>((windowRect.bottom - windowRect.top) - currentClientHeight));
        const int targetWindowHeight  = ResolveWindowClientHeightPx(dpi) + nonClientHeight;

        SetWindowPos(_hWnd.get(), nullptr, 0, 0, windowRect.right - windowRect.left, targetWindowHeight, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        Layout();
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip       = 16.0f;
        constexpr float kGapDip          = 8.0f;
        constexpr float kRowHeightDip    = 34.0f;
        constexpr float kLabelHeightDip  = 22.0f;
        constexpr float kHistoryWidthDip = 88.0f;
        constexpr float kHintHeightDip   = 28.0f;
        constexpr float kHelpHeightDip   = 100.0f;
        constexpr float kButtonWidthDip  = 96.0f;
        constexpr float kButtonHeightDip = 34.0f;

        const float left         = client.left + kMarginDip;
        const float right        = std::max(left, client.right - kMarginDip);
        const float contentWidth = std::max(0.0f, right - left);
        float y                  = client.top + kMarginDip;

        if (_label)
        {
            _label->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + kGapDip;

        const bool showHistoryButton = _historyButton && _historyButton->IsVisible();
        const float historyLeft      = std::max(left, right - kHistoryWidthDip);
        const float fieldRight       = showHistoryButton ? std::max(left, historyLeft - kGapDip) : right;

        if (_textField)
        {
            _textField->SetBounds(D2D1::RectF(left, y, fieldRight, y + kRowHeightDip));
        }
        if (_historyButton)
        {
            _historyButton->SetVisible(showHistoryButton);
            if (showHistoryButton)
            {
                _historyButton->SetBounds(D2D1::RectF(historyLeft, y, right, y + kRowHeightDip));
            }
        }
        y += kRowHeightDip + kGapDip;

        if (_hintButton)
        {
            _hintButton->SetBounds(D2D1::RectF(left, y, std::min(right, left + std::min(contentWidth, 160.0f)), y + kHintHeightDip));
        }
        y += kHintHeightDip + kGapDip;

        if (_helpLabel)
        {
            _helpLabel->SetBounds(D2D1::RectF(left, y, right, y + (_helpExpanded ? kHelpHeightDip : 0.0f)));
        }
        if (_helpExpanded)
        {
            y += kHelpHeightDip + kGapDip;
        }

        const float buttonsTop = std::max(y, client.bottom - kMarginDip - kButtonHeightDip);
        const float cancelLeft = std::max(left, right - kButtonWidthDip);
        const float okLeft     = std::max(left, cancelLeft - kGapDip - kButtonWidthDip);

        if (_okButton)
        {
            _okButton->SetBounds(D2D1::RectF(okLeft, buttonsTop, okLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
    }

    void ShowHistoryMenu() noexcept
    {
        if (! _textField || ! _historyButton || ! _historyButton->IsVisible() || ! _hWnd || IsWindow(_hWnd.get()) == FALSE)
        {
            return;
        }

        _dxHost.CommitFocusedTextInput();
        TryShowPromptHistoryMenu(_hWnd.get(),
                                 _dxHost,
                                 _palette,
                                 _historyButton->GetBounds(),
                                 _history,
                                 _textField->GetText(),
                                 [this](std::wstring_view value) noexcept
        {
            if (! _textField)
            {
                return;
            }

            _textField->SetText(std::wstring(value));
            _dxHost.SetFocusControl(_textField);
        });
    }

    void Confirm() noexcept
    {
        std::wstring trimmed = StringUtils::TrimWhitespaceCopy(_textField ? std::wstring(_textField->GetText()) : std::wstring{});
        if (trimmed.empty())
        {
            MessageBeep(MB_ICONWARNING);
            if (_textField)
            {
                _dxHost.SetFocusControl(_textField);
            }
            return;
        }

        _result = std::move(trimmed);
        _done   = true;
        ClosePromptWindow();
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        ClosePromptWindow();
    }

    void ClosePromptWindow() noexcept
    {
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _dxHost.SetFocusControl(nullptr);
            _hWnd.reset();
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(FolderViewSelectionMaskPromptDebugCommand command, LPARAM lParam) noexcept
    {
        switch (command)
        {
            case FolderViewSelectionMaskPromptDebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<FolderViewSelectionMaskPromptDebugSnapshot*>(lParam);
                if (! snapshot)
                {
                    return FALSE;
                }

                snapshot->usesDxUiHost            = _dxHost.GetRoot() != nullptr;
                snapshot->visibleChildWindowCount = CountVisibleChildWindowsLocal(_hWnd.get());
                snapshot->title                   = _captionText;
                snapshot->text                    = _textField ? std::wstring(_textField->GetText()) : std::wstring{};
                return TRUE;
            }
            case FolderViewSelectionMaskPromptDebugCommand::SetText:
            {
                const auto* text = reinterpret_cast<const std::wstring*>(lParam);
                if (! text || ! _textField)
                {
                    return FALSE;
                }
                _textField->SetTextAndNotify(*text);
                _dxHost.SetFocusControl(_textField);
                _dxHost.SyncTextInput(_textField);
                _dxHost.Invalidate();
                return TRUE;
            }
            case FolderViewSelectionMaskPromptDebugCommand::Confirm: Confirm(); return TRUE;
            case FolderViewSelectionMaskPromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

    HWND _ownerWindow = nullptr;
    std::vector<std::wstring> _history;
    AppTheme _theme{};
    std::wstring _captionText;
    std::wstring _labelText;
    bool _helpExpanded = false;
    RedSalamander::DxUi::ThemePalette _palette{};
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root           = nullptr;
    RedSalamander::DxUi::Label* _label          = nullptr;
    RedSalamander::DxUi::TextField* _textField  = nullptr;
    RedSalamander::DxUi::Button* _historyButton = nullptr;
    RedSalamander::DxUi::Button* _hintButton    = nullptr;
    RedSalamander::DxUi::Label* _helpLabel      = nullptr;
    RedSalamander::DxUi::Button* _okButton      = nullptr;
    RedSalamander::DxUi::Button* _cancelButton  = nullptr;
    bool _done                                  = false;
    std::optional<std::wstring> _result;
};

std::optional<std::filesystem::path> TryResolveInstanceContextToWindowsPath(std::wstring_view instanceContext) noexcept
{
    if (instanceContext.empty())
    {
        return std::nullopt;
    }

    std::wstring text = StringUtils::TrimWhitespaceCopy(instanceContext);
    if (text.empty())
    {
        return std::nullopt;
    }

    if (text.size() >= 2u && text.front() == L'"' && text.back() == L'"')
    {
        text.erase(text.begin());
        text.pop_back();
        text = StringUtils::TrimWhitespaceCopy(text);
        if (text.empty())
        {
            return std::nullopt;
        }
    }

    if (LooksLikeWindowsAbsolutePath(text))
    {
        return std::filesystem::path(text);
    }

    std::wstring prefix;
    std::wstring remainder;
    if (! TryParsePluginPrefix(text, prefix, remainder))
    {
        return std::nullopt;
    }

    std::wstring_view remainderView = remainder;
    const size_t bar                = remainderView.find(L'|');
    if (bar != std::wstring_view::npos)
    {
        remainderView = remainderView.substr(0, bar);
    }

    if (! LooksLikeWindowsAbsolutePath(remainderView))
    {
        return std::nullopt;
    }

    return std::filesystem::path(remainderView);
}

bool ContainsPathSeparators(std::wstring_view name) noexcept
{
    return name.find_first_of(L"\\/") != std::wstring_view::npos;
}

[[nodiscard]] std::optional<UINT> ResolveCreateDirectoryValidationMessageId(std::wstring_view trimmed) noexcept
{
    if (trimmed.empty())
    {
        return IDS_MSG_PANE_CREATE_DIR_EMPTY_NAME;
    }
    if (trimmed == L"." || trimmed == L"..")
    {
        return IDS_MSG_PANE_CREATE_DIR_DOT_NAME;
    }
    if (ContainsPathSeparators(trimmed))
    {
        return IDS_MSG_PANE_CREATE_DIR_INVALID_CHARS;
    }

    constexpr std::wstring_view kInvalidNameChars = L":*?\"<>|";
    if (trimmed.find_first_of(kInvalidNameChars) != std::wstring_view::npos)
    {
        return IDS_MSG_PANE_CREATE_DIR_INVALID_CHARS;
    }
    if (trimmed.find_first_of(L"\r\n\t") != std::wstring_view::npos)
    {
        return IDS_MSG_PANE_CREATE_DIR_INVALID_WHITESPACE;
    }
    return std::nullopt;
}

[[nodiscard]] bool IsReservedWindowsDeviceName(std::wstring_view name) noexcept
{
    std::wstring_view stem = name;
    if (const size_t dot = stem.find(L'.'); dot != std::wstring_view::npos)
    {
        stem = stem.substr(0, dot);
    }

    while (! stem.empty() && (stem.back() == L' ' || stem.back() == L'.'))
    {
        stem.remove_suffix(1);
    }

    if (stem.empty())
    {
        return false;
    }

    if (EqualsNoCase(stem, L"CON") || EqualsNoCase(stem, L"PRN") || EqualsNoCase(stem, L"AUX") || EqualsNoCase(stem, L"NUL"))
    {
        return true;
    }

    if (stem.size() == 4u && (StartsWithNoCase(stem, L"COM") || StartsWithNoCase(stem, L"LPT")) && stem[3] >= L'1' && stem[3] <= L'9')
    {
        return true;
    }

    return false;
}

[[nodiscard]] std::optional<UINT> ResolveEditNewValidationMessageId(std::wstring_view trimmed, const std::filesystem::path& targetFolder) noexcept
{
    if (trimmed.empty())
    {
        return IDS_MSG_PANE_EDIT_NEW_EMPTY_NAME;
    }
    if (trimmed == L"." || trimmed == L"..")
    {
        return IDS_MSG_PANE_EDIT_NEW_DOT_NAME;
    }
    if (LooksLikeWindowsAbsolutePath(trimmed) || ContainsPathSeparators(trimmed))
    {
        return IDS_MSG_PANE_EDIT_NEW_INVALID_CHARS;
    }

    constexpr std::wstring_view kInvalidNameChars = L":*?\"<>|";
    if (trimmed.find_first_of(kInvalidNameChars) != std::wstring_view::npos)
    {
        return IDS_MSG_PANE_EDIT_NEW_INVALID_CHARS;
    }
    if (trimmed.find_first_of(L"\r\n\t") != std::wstring_view::npos)
    {
        return IDS_MSG_PANE_EDIT_NEW_INVALID_WHITESPACE;
    }
    if (IsReservedWindowsDeviceName(trimmed))
    {
        return IDS_MSG_PANE_EDIT_NEW_RESERVED_NAME;
    }

    if (! targetFolder.empty())
    {
        std::error_code ec;
        const bool exists = std::filesystem::exists(targetFolder / std::filesystem::path(trimmed), ec);
        if (! ec && exists)
        {
            return IDS_MSG_PANE_EDIT_NEW_EXISTS;
        }
    }

    return std::nullopt;
}

std::wstring GetComputerNameTextForFileActions() noexcept
{
    wchar_t computerName[MAX_COMPUTERNAME_LENGTH + 1]{};
    DWORD computerNameLength = static_cast<DWORD>(std::size(computerName));
    if (GetComputerNameW(computerName, &computerNameLength) == FALSE || computerNameLength == 0u)
    {
        return {};
    }
    return std::wstring(computerName, computerNameLength);
}

class FolderViewCreateDirectoryPromptWindow final
{
public:
    FolderViewCreateDirectoryPromptWindow(const FolderViewCreateDirectoryPromptWindow&)            = delete;
    FolderViewCreateDirectoryPromptWindow& operator=(const FolderViewCreateDirectoryPromptWindow&) = delete;
    FolderViewCreateDirectoryPromptWindow(FolderViewCreateDirectoryPromptWindow&&)                 = delete;
    FolderViewCreateDirectoryPromptWindow& operator=(FolderViewCreateDirectoryPromptWindow&&)      = delete;

    FolderViewCreateDirectoryPromptWindow(HWND ownerWindow, std::wstring createInPath, std::wstring initialName, const AppTheme& theme) noexcept
        : _ownerWindow(GetOwnerWindowOrSelf(ownerWindow)),
          _createInPath(std::move(createInPath)),
          _initialName(std::move(initialName)),
          _captionText(LoadStringResource(nullptr, IDS_CMD_MAKE_DIRECTORY)),
          _theme(theme)
    {
    }

    [[nodiscard]] std::optional<std::wstring> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 480), ScalePanePromptForDpi(dpi, 244)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kFolderViewCreateDirectoryPromptClassName,
                                          _captionText.c_str(),
                                          style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(_hWnd.get(), _ownerWindow));
        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return std::nullopt;
            }
            if (getMessageResult == 0)
            {
                _done = true;
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<FolderViewCreateDirectoryPromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<FolderViewCreateDirectoryPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            if (message == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                self->_done = true;
            }
            return dxResult;
        }

#ifdef ENABLE_TESTS
        if (message == GetFolderViewCreateDirectoryPromptDebugMessage())
        {
            return self->OnDebugCommand(static_cast<FolderViewCreateDirectoryPromptDebugCommand>(wParam), lParam);
        }
#endif

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, self->_theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = FolderViewCreateDirectoryPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kFolderViewCreateDirectoryPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        if (! _captionText.empty())
        {
            SetWindowTextW(hwnd, _captionText.c_str());
        }
        ApplyTheme();
        if (_nameField)
        {
            _nameField->SetText(_initialName);
        }
        _currentText = _initialName;
        UpdateValidationUi();
        Layout();
        if (_nameField)
        {
            _dxHost.SetFocusControl(_nameField);
            _nameField->SetSelectionRange(0u, _initialName.size());
            _dxHost.SyncTextInput(_nameField);
        }
        _dxHost.SetDefaultButton(_createButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi()
    {
        if (_root != nullptr)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _pathCaptionLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_LABEL_CREATE_DIR_IN));
        _pathCaptionLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _pathLabel = _root->AddChild<Label>(_createInPath);
        _pathLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _pathLabel->SetMultiline(true);

        _nameCaptionLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_LABEL_CREATE_DIR_NAME));
        _nameCaptionLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _nameField = _root->AddChild<TextField>(_initialName);
        _nameField->SetMultiline(false);
        _nameField->SetOnTextChanged([this](std::wstring_view text)
        {
            _currentText.assign(text);
            if (_validationMessageId.has_value())
            {
                _validationMessageId = ResolveCreateDirectoryValidationMessageId(StringUtils::TrimWhitespaceCopy(_currentText));
            }
            UpdateValidationUi();
            Layout();
        });

        _validationLabel = _root->AddChild<Label>();
        _validationLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _validationLabel->SetMultiline(true);

        _createButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BUTTON_CREATE));
        _createButton->SetPrimary(true);
        _createButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_validationLabel)
        {
            _validationLabel->SetTextColor(_palette.errorText);
        }
        if (_hWnd)
        {
            if (! _captionText.empty())
            {
                SetWindowTextW(_hWnd.get(), _captionText.c_str());
            }
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
            static_cast<void>(RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW));
        }
    }

    void UpdateValidationUi() noexcept
    {
        if (! _validationLabel)
        {
            return;
        }

        _validationText.clear();
        if (_validationMessageId.has_value())
        {
            _validationText = LoadStringResource(nullptr, _validationMessageId.value());
        }
        _validationLabel->SetText(_validationText);
        _validationLabel->SetVisible(! _validationText.empty());
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip           = 16.0f;
        constexpr float kGapDip              = 8.0f;
        constexpr float kCaptionHeightDip    = 22.0f;
        constexpr float kPathHeightDip       = 58.0f;
        constexpr float kRowHeightDip        = 34.0f;
        constexpr float kValidationHeightDip = 40.0f;
        constexpr float kButtonWidthDip      = 96.0f;
        constexpr float kButtonHeightDip     = 34.0f;

        const float left  = client.left + kMarginDip;
        const float right = std::max(left, client.right - kMarginDip);
        float y           = client.top + kMarginDip;

        if (_pathCaptionLabel)
        {
            _pathCaptionLabel->SetBounds(D2D1::RectF(left, y, right, y + kCaptionHeightDip));
        }
        y += kCaptionHeightDip + 4.0f;

        if (_pathLabel)
        {
            _pathLabel->SetBounds(D2D1::RectF(left, y, right, y + kPathHeightDip));
        }
        y += kPathHeightDip + kGapDip;

        if (_nameCaptionLabel)
        {
            _nameCaptionLabel->SetBounds(D2D1::RectF(left, y, right, y + kCaptionHeightDip));
        }
        y += kCaptionHeightDip + 4.0f;

        if (_nameField)
        {
            _nameField->SetBounds(D2D1::RectF(left, y, right, y + kRowHeightDip));
        }
        y += kRowHeightDip + kGapDip;

        const bool showValidation = ! _validationText.empty();
        if (_validationLabel)
        {
            _validationLabel->SetBounds(D2D1::RectF(left, y, right, y + (showValidation ? kValidationHeightDip : 0.0f)));
        }
        if (showValidation)
        {
            y += kValidationHeightDip + kGapDip;
        }

        const float buttonsTop = std::max(y, client.bottom - kMarginDip - kButtonHeightDip);
        const float cancelLeft = std::max(left, right - kButtonWidthDip);
        const float createLeft = std::max(left, cancelLeft - kGapDip - kButtonWidthDip);

        if (_createButton)
        {
            _createButton->SetBounds(D2D1::RectF(createLeft, buttonsTop, createLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
    }

    void Confirm() noexcept
    {
        const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(_nameField ? std::wstring(_nameField->GetText()) : std::wstring{});
        _validationMessageId       = ResolveCreateDirectoryValidationMessageId(trimmed);
        UpdateValidationUi();
        Layout();
        if (_validationMessageId.has_value())
        {
            MessageBeep(MB_ICONWARNING);
            if (_nameField)
            {
                _dxHost.SetFocusControl(_nameField);
            }
            return;
        }

        _result = trimmed;
        _done   = true;
        ClosePromptWindow();
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        ClosePromptWindow();
    }

    void ClosePromptWindow() noexcept
    {
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _dxHost.SetFocusControl(nullptr);
            _hWnd.reset();
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(FolderViewCreateDirectoryPromptDebugCommand command, LPARAM lParam) noexcept
    {
        switch (command)
        {
            case FolderViewCreateDirectoryPromptDebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<FolderViewCreateDirectoryPromptDebugSnapshot*>(lParam);
                if (! snapshot)
                {
                    return FALSE;
                }

                snapshot->usesDxUiHost            = _dxHost.GetRoot() != nullptr;
                snapshot->visibleChildWindowCount = CountVisibleChildWindowsLocal(_hWnd.get());
                snapshot->createInPath            = _createInPath;
                snapshot->text                    = _nameField ? std::wstring(_nameField->GetText()) : std::wstring{};
                snapshot->validationText          = _validationText;
                snapshot->nameFieldFocused        = _dxHost.GetFocusControl() == _nameField;
                snapshot->selectionStart          = snapshot->text.size();
                snapshot->selectionEnd            = snapshot->text.size();
                if (_nameField)
                {
                    if (const auto selectionRange = _nameField->GetSelectionRange(); selectionRange.has_value())
                    {
                        snapshot->selectionStart = selectionRange->first;
                        snapshot->selectionEnd   = selectionRange->second;
                    }
                }
                return TRUE;
            }
            case FolderViewCreateDirectoryPromptDebugCommand::SetText:
            {
                const auto* text = reinterpret_cast<const std::wstring*>(lParam);
                if (! text || ! _nameField)
                {
                    return FALSE;
                }
                _nameField->SetText(*text);
                _currentText = *text;
                _validationMessageId.reset();
                UpdateValidationUi();
                Layout();
                _dxHost.SetFocusControl(_nameField);
                return TRUE;
            }
            case FolderViewCreateDirectoryPromptDebugCommand::Confirm: Confirm(); return TRUE;
            case FolderViewCreateDirectoryPromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

private:
    HWND _ownerWindow = nullptr;
    std::wstring _createInPath;
    std::wstring _initialName;
    std::wstring _currentText;
    std::wstring _validationText;
    std::wstring _captionText;
    AppTheme _theme{};
    RedSalamander::DxUi::ThemePalette _palette{};
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root             = nullptr;
    RedSalamander::DxUi::Label* _pathCaptionLabel = nullptr;
    RedSalamander::DxUi::Label* _pathLabel        = nullptr;
    RedSalamander::DxUi::Label* _nameCaptionLabel = nullptr;
    RedSalamander::DxUi::TextField* _nameField    = nullptr;
    RedSalamander::DxUi::Label* _validationLabel  = nullptr;
    RedSalamander::DxUi::Button* _createButton    = nullptr;
    RedSalamander::DxUi::Button* _cancelButton    = nullptr;
    std::optional<UINT> _validationMessageId;
    bool _done = false;
    std::optional<std::wstring> _result;
};

std::wstring TryGetFileSystemPluginDisplayName(const std::vector<FileSystemPluginManager::PluginEntry>& plugins,
                                               std::wstring_view pluginId,
                                               std::wstring_view pluginShortId) noexcept
{
    const FileSystemPluginManager::PluginEntry* entry = FindPluginById(plugins, pluginId);
    if (! entry)
    {
        entry = FindPluginByShortId(plugins, pluginShortId);
    }

    if (entry && ! entry->name.empty())
    {
        return entry->name;
    }

    if (! pluginShortId.empty())
    {
        return std::wstring(pluginShortId);
    }

    if (! pluginId.empty())
    {
        return std::wstring(pluginId);
    }

    return {};
}

std::optional<std::wstring> PromptForCreateDirectoryName(HWND ownerWindow, std::wstring_view createInPath, std::wstring_view initialName, const AppTheme& theme)
{
    auto prompt = std::make_unique<FolderViewCreateDirectoryPromptWindow>(ownerWindow, std::wstring(createInPath), std::wstring(initialName), theme);
    return prompt ? prompt->ShowModal() : std::nullopt;
}

class FolderViewEditNewPromptWindow final
{
public:
    FolderViewEditNewPromptWindow(const FolderViewEditNewPromptWindow&)            = delete;
    FolderViewEditNewPromptWindow& operator=(const FolderViewEditNewPromptWindow&) = delete;
    FolderViewEditNewPromptWindow(FolderViewEditNewPromptWindow&&)                 = delete;
    FolderViewEditNewPromptWindow& operator=(FolderViewEditNewPromptWindow&&)      = delete;

    FolderViewEditNewPromptWindow(HWND ownerWindow,
                                  std::filesystem::path targetFolder,
                                  std::wstring displayPath,
                                  const Common::Settings::EditorFileActionsSettings* editorSettings,
                                  std::wstring computerName,
                                  const AppTheme& theme,
                                  std::wstring initialFileName = {},
                                  std::wstring captionText     = {},
                                  bool showEditorControls      = true) noexcept
        : _ownerWindow(GetOwnerWindowOrSelf(ownerWindow)),
          _targetFolder(std::move(targetFolder)),
          _displayPath(std::move(displayPath)),
          _editorSettings(editorSettings),
          _computerName(std::move(computerName)),
          _currentText(std::move(initialFileName)),
          _captionText(captionText.empty() ? LoadStringResource(nullptr, IDS_CAPTION_EDIT_NEW) : std::move(captionText)),
          _theme(theme),
          _showEditorControls(showEditorControls)
    {
    }

    [[nodiscard]] std::optional<EditNewPromptResult> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 520), ScalePanePromptForDpi(dpi, 306)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kFolderViewEditNewPromptClassName,
                                          _captionText.c_str(),
                                          style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(_hWnd.get(), _ownerWindow));
        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return std::nullopt;
            }
            if (getMessageResult == 0)
            {
                _done = true;
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<FolderViewEditNewPromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<FolderViewEditNewPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            if (message == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                self->_done = true;
            }
            return dxResult;
        }

#ifdef ENABLE_TESTS
        if (message == GetFolderViewEditNewPromptDebugMessage())
        {
            return self->OnDebugCommand(static_cast<FolderViewEditNewPromptDebugCommand>(wParam), lParam);
        }
#endif

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, self->_theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = FolderViewEditNewPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kFolderViewEditNewPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        if (! _captionText.empty())
        {
            SetWindowTextW(hwnd, _captionText.c_str());
        }
        if (_nameField && ! _currentText.empty())
        {
            _nameField->SetText(_currentText);
        }
        ApplyTheme();
        UpdateEditorChoices();
        UpdateValidationUi();
        Layout();
        if (_nameField)
        {
            _dxHost.SetFocusControl(_nameField);
            const size_t selectionEnd = _currentText.empty() ? 0u : _currentText.size();
            _nameField->SetSelectionRange(0u, selectionEnd);
            _dxHost.SyncTextInput(_nameField);
        }
        _dxHost.SetDefaultButton(_createButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi()
    {
        if (_root != nullptr)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _pathCaptionLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_LABEL_EDIT_NEW_IN));
        _pathCaptionLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _pathLabel = _root->AddChild<Label>(_displayPath);
        _pathLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _pathLabel->SetMultiline(true);

        _nameCaptionLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_LABEL_EDIT_NEW_NAME));
        _nameCaptionLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _nameField = _root->AddChild<TextField>();
        _nameField->SetMultiline(false);
        _nameField->SetOnTextChanged([this](std::wstring_view text)
        {
            _currentText.assign(text);
            if (_validationMessageId.has_value())
            {
                _validationMessageId = ResolveEditNewValidationMessageId(StringUtils::TrimWhitespaceCopy(_currentText), _targetFolder);
            }
            UpdateEditorChoices();
            UpdateValidationUi();
            Layout();
        });

        if (_showEditorControls)
        {
            _editorCaptionLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_LABEL_EDIT_NEW_EDITOR));
            _editorCaptionLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

            _editorCombo = _root->AddChild<ComboBox>();
            _editorCombo->SetVariant(ComboBoxVariant::Window);
            _editorCombo->SetOnSelectionChanged([this](size_t index) noexcept
            {
                if (index < _editorActionIds.size())
                {
                    _selectedEditorActionId = _editorActionIds[index];
                }
            });
        }

        _validationLabel = _root->AddChild<Label>();
        _validationLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _validationLabel->SetMultiline(true);

        _createButton = _root->AddChild<Button>(LoadStringResource(nullptr, _showEditorControls ? IDS_BUTTON_CREATE_AND_EDIT : IDS_BUTTON_CREATE));
        _createButton->SetPrimary(true);
        _createButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_validationLabel)
        {
            _validationLabel->SetTextColor(_palette.errorText);
        }
        if (_hWnd)
        {
            if (! _captionText.empty())
            {
                SetWindowTextW(_hWnd.get(), _captionText.c_str());
            }
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
            static_cast<void>(RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW));
        }
    }

    void UpdateValidationUi() noexcept
    {
        if (! _validationLabel)
        {
            return;
        }

        _validationText.clear();
        if (_validationMessageId.has_value())
        {
            _validationText = LoadStringResource(nullptr, _validationMessageId.value());
        }
        _validationLabel->SetText(_validationText);
        _validationLabel->SetVisible(! _validationText.empty());
    }

    void UpdateEditorChoices() noexcept
    {
        if (! _editorCombo)
        {
            return;
        }

        const auto startedAt = std::chrono::steady_clock::now();
        _editorActionIds.clear();
        _editorDisplayNames.clear();
        std::vector<RedSalamander::DxUi::ComboBox::Item> items;

        const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(_currentText);
        std::optional<FileActionResolver::Request> request;
        if (_editorSettings && ! trimmed.empty())
        {
            request = FileActionResolver::Request{
                .command      = FileActionResolver::Command::EditNew,
                .filePath     = _targetFolder / std::filesystem::path(trimmed),
                .computerName = _computerName,
            };
            const std::vector<const Common::Settings::FileActionDefinition*> actions =
                FileActionResolver::CollectAssociatedEditorActions(*_editorSettings, request.value());
            items.reserve(actions.size());

            for (const Common::Settings::FileActionDefinition* action : actions)
            {
                if (! action || action->kind != Common::Settings::FileActionKind::ExternalProgram || action->id.empty() ||
                    StringUtils::TrimWhitespace(action->executablePath).empty())
                {
                    continue;
                }

                const std::wstring displayName = action->displayName.empty() ? action->id : action->displayName;
                _editorActionIds.push_back(action->id);
                _editorDisplayNames.push_back(displayName);
                items.push_back(RedSalamander::DxUi::ComboBox::Item{action->id, displayName});
            }
        }

        std::optional<size_t> selectedIndex;
        if (! _editorActionIds.empty() && request.has_value())
        {
            const FileActionResolver::Resolution preferred =
                _editorSettings ? FileActionResolver::ResolveEditorAction(*_editorSettings, request.value()) : FileActionResolver::Resolution{};
            if (preferred.action)
            {
                for (size_t index = 0; index < _editorActionIds.size(); ++index)
                {
                    if (EqualsNoCase(_editorActionIds[index], preferred.action->id))
                    {
                        selectedIndex = index;
                        break;
                    }
                }
            }
            if (! selectedIndex.has_value())
            {
                selectedIndex = 0u;
            }
        }

        _editorCombo->SetItems(std::move(items));
        _editorCombo->SetEnabled(! _editorActionIds.empty());
        _editorCombo->SetSelectedIndex(selectedIndex);
        _selectedEditorActionId = selectedIndex.has_value() ? _editorActionIds[selectedIndex.value()] : std::wstring{};

        Debug::Perf::Emit(L"fileaction.editnew.editor_combo_us",
                          L"",
                          Debug::Perf::ElapsedUs(startedAt),
                          static_cast<uint64_t>(_editorActionIds.size()),
                          selectedIndex.has_value() ? static_cast<uint64_t>(selectedIndex.value()) : 0u,
                          S_OK);
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip           = 16.0f;
        constexpr float kGapDip              = 8.0f;
        constexpr float kCaptionHeightDip    = 22.0f;
        constexpr float kPathHeightDip       = 54.0f;
        constexpr float kRowHeightDip        = 34.0f;
        constexpr float kValidationHeightDip = 40.0f;
        constexpr float kButtonWidthDip      = 120.0f;
        constexpr float kButtonHeightDip     = 34.0f;

        const float left  = client.left + kMarginDip;
        const float right = std::max(left, client.right - kMarginDip);
        float y           = client.top + kMarginDip;

        if (_pathCaptionLabel)
        {
            _pathCaptionLabel->SetBounds(D2D1::RectF(left, y, right, y + kCaptionHeightDip));
        }
        y += kCaptionHeightDip + 4.0f;

        if (_pathLabel)
        {
            _pathLabel->SetBounds(D2D1::RectF(left, y, right, y + kPathHeightDip));
        }
        y += kPathHeightDip + kGapDip;

        if (_nameCaptionLabel)
        {
            _nameCaptionLabel->SetBounds(D2D1::RectF(left, y, right, y + kCaptionHeightDip));
        }
        y += kCaptionHeightDip + 4.0f;

        if (_nameField)
        {
            _nameField->SetBounds(D2D1::RectF(left, y, right, y + kRowHeightDip));
        }
        y += kRowHeightDip + kGapDip;

        if (_showEditorControls)
        {
            if (_editorCaptionLabel)
            {
                _editorCaptionLabel->SetBounds(D2D1::RectF(left, y, right, y + kCaptionHeightDip));
            }
            y += kCaptionHeightDip + 4.0f;

            if (_editorCombo)
            {
                _editorCombo->SetBounds(D2D1::RectF(left, y, right, y + kRowHeightDip));
            }
            y += kRowHeightDip + kGapDip;
        }

        const bool showValidation = ! _validationText.empty();
        if (_validationLabel)
        {
            _validationLabel->SetBounds(D2D1::RectF(left, y, right, y + (showValidation ? kValidationHeightDip : 0.0f)));
        }
        if (showValidation)
        {
            y += kValidationHeightDip + kGapDip;
        }

        const float buttonsTop = std::max(y, client.bottom - kMarginDip - kButtonHeightDip);
        const float cancelLeft = std::max(left, right - kButtonWidthDip);
        const float createLeft = std::max(left, cancelLeft - kGapDip - kButtonWidthDip);

        if (_createButton)
        {
            _createButton->SetBounds(D2D1::RectF(createLeft, buttonsTop, createLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
    }

    void Confirm() noexcept
    {
        const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(_nameField ? std::wstring(_nameField->GetText()) : std::wstring{});
        _validationMessageId       = ResolveEditNewValidationMessageId(trimmed, _targetFolder);
        UpdateValidationUi();
        Layout();
        if (_validationMessageId.has_value())
        {
            MessageBeep(MB_ICONWARNING);
            if (_nameField)
            {
                _dxHost.SetFocusControl(_nameField);
            }
            return;
        }

        _result = EditNewPromptResult{.fileName = trimmed, .editorActionId = _selectedEditorActionId};
        _done   = true;
        ClosePromptWindow();
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        ClosePromptWindow();
    }

    void ClosePromptWindow() noexcept
    {
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _dxHost.SetFocusControl(nullptr);
            _hWnd.reset();
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(FolderViewEditNewPromptDebugCommand command, LPARAM lParam) noexcept
    {
        switch (command)
        {
            case FolderViewEditNewPromptDebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<FolderViewEditNewPromptDebugSnapshot*>(lParam);
                if (! snapshot)
                {
                    return FALSE;
                }

                snapshot->usesDxUiHost            = _dxHost.GetRoot() != nullptr;
                snapshot->visibleChildWindowCount = CountVisibleChildWindowsLocal(_hWnd.get());
                snapshot->createInPath            = _displayPath;
                snapshot->fileNameText            = _nameField ? std::wstring(_nameField->GetText()) : std::wstring{};
                snapshot->validationText          = _validationText;
                snapshot->nameFieldFocused        = _dxHost.GetFocusControl() == _nameField;
                snapshot->selectionStart          = snapshot->fileNameText.size();
                snapshot->selectionEnd            = snapshot->fileNameText.size();
                if (_nameField)
                {
                    if (const auto selectionRange = _nameField->GetSelectionRange(); selectionRange.has_value())
                    {
                        snapshot->selectionStart = selectionRange->first;
                        snapshot->selectionEnd   = selectionRange->second;
                    }
                }
                snapshot->editorComboEnabled     = _editorCombo && _editorCombo->IsEnabled();
                snapshot->selectedEditorActionId = _selectedEditorActionId;
                snapshot->editorActionIds        = _editorActionIds;
                snapshot->editorDisplayNames     = _editorDisplayNames;
                return TRUE;
            }
            case FolderViewEditNewPromptDebugCommand::SetText:
            {
                const auto* text = reinterpret_cast<const std::wstring*>(lParam);
                if (! text || ! _nameField)
                {
                    return FALSE;
                }
                _nameField->SetText(*text);
                _currentText = *text;
                _validationMessageId.reset();
                UpdateEditorChoices();
                UpdateValidationUi();
                Layout();
                _dxHost.SetFocusControl(_nameField);
                return TRUE;
            }
            case FolderViewEditNewPromptDebugCommand::SelectEditor:
            {
                const auto* actionId = reinterpret_cast<const std::wstring*>(lParam);
                if (! actionId || ! _editorCombo)
                {
                    return FALSE;
                }
                for (size_t index = 0; index < _editorActionIds.size(); ++index)
                {
                    if (EqualsNoCase(_editorActionIds[index], *actionId))
                    {
                        _editorCombo->SetSelectedIndex(index);
                        _selectedEditorActionId = _editorActionIds[index];
                        return TRUE;
                    }
                }
                return FALSE;
            }
            case FolderViewEditNewPromptDebugCommand::Confirm: Confirm(); return TRUE;
            case FolderViewEditNewPromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

private:
    HWND _ownerWindow = nullptr;
    std::filesystem::path _targetFolder;
    std::wstring _displayPath;
    const Common::Settings::EditorFileActionsSettings* _editorSettings = nullptr;
    std::wstring _computerName;
    std::wstring _currentText;
    std::wstring _selectedEditorActionId;
    std::vector<std::wstring> _editorActionIds;
    std::vector<std::wstring> _editorDisplayNames;
    std::wstring _validationText;
    std::wstring _captionText;
    AppTheme _theme{};
    RedSalamander::DxUi::ThemePalette _palette{};
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root               = nullptr;
    RedSalamander::DxUi::Label* _pathCaptionLabel   = nullptr;
    RedSalamander::DxUi::Label* _pathLabel          = nullptr;
    RedSalamander::DxUi::Label* _nameCaptionLabel   = nullptr;
    RedSalamander::DxUi::TextField* _nameField      = nullptr;
    RedSalamander::DxUi::Label* _editorCaptionLabel = nullptr;
    RedSalamander::DxUi::ComboBox* _editorCombo     = nullptr;
    RedSalamander::DxUi::Label* _validationLabel    = nullptr;
    RedSalamander::DxUi::Button* _createButton      = nullptr;
    RedSalamander::DxUi::Button* _cancelButton      = nullptr;
    std::optional<UINT> _validationMessageId;
    bool _showEditorControls = true;
    bool _done               = false;
    std::optional<EditNewPromptResult> _result;
};

std::optional<EditNewPromptResult> PromptForEditNewFile(HWND ownerWindow,
                                                        const std::filesystem::path& targetFolder,
                                                        std::wstring_view displayPath,
                                                        const Common::Settings::EditorFileActionsSettings* editorSettings,
                                                        std::wstring_view computerName,
                                                        const AppTheme& theme,
                                                        std::wstring_view initialFileName,
                                                        std::wstring_view captionText,
                                                        bool showEditorControls)
{
    auto prompt = std::make_unique<FolderViewEditNewPromptWindow>(ownerWindow,
                                                                  targetFolder,
                                                                  std::wstring(displayPath),
                                                                  editorSettings,
                                                                  std::wstring(computerName),
                                                                  theme,
                                                                  std::wstring(initialFileName),
                                                                  std::wstring(captionText),
                                                                  showEditorControls);
    return prompt ? prompt->ShowModal() : std::nullopt;
}

class FolderViewChangeCasePromptWindow final
{
public:
    FolderViewChangeCasePromptWindow(const FolderViewChangeCasePromptWindow&)            = delete;
    FolderViewChangeCasePromptWindow& operator=(const FolderViewChangeCasePromptWindow&) = delete;
    FolderViewChangeCasePromptWindow(FolderViewChangeCasePromptWindow&&)                 = delete;
    FolderViewChangeCasePromptWindow& operator=(FolderViewChangeCasePromptWindow&&)      = delete;

    FolderViewChangeCasePromptWindow(HWND ownerWindow, const AppTheme& theme, bool allowSubdirs) noexcept
        : _ownerWindow(GetOwnerWindowOrSelf(ownerWindow)),
          _restoreFocusWindow(ownerWindow && IsWindow(ownerWindow) != FALSE ? ownerWindow : nullptr),
          _theme(theme),
          _allowSubdirs(allowSubdirs)
    {
        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            const HWND focused = GetFocus();
            if (focused && IsWindow(focused) != FALSE && (focused == _ownerWindow || IsChild(_ownerWindow, focused) != FALSE))
            {
                _restoreFocusWindow = focused;
            }
            else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE ||
                     (_restoreFocusWindow != _ownerWindow && IsChild(_ownerWindow, _restoreFocusWindow) == FALSE))
            {
                _restoreFocusWindow = _ownerWindow;
            }
        }
        else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE)
        {
            _restoreFocusWindow = nullptr;
        }
    }

    [[nodiscard]] std::optional<ChangeCase::Options> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 448), ScalePanePromptForDpi(dpi, 304)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
        const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled] noexcept
        {
            if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
            {
                EnableWindow(_ownerWindow, TRUE);
                SetActiveWindow(_ownerWindow);

                const HWND restoreFocus = (_restoreFocusWindow && IsWindow(_restoreFocusWindow) != FALSE &&
                                           (_restoreFocusWindow == _ownerWindow || IsChild(_ownerWindow, _restoreFocusWindow) != FALSE))
                                              ? _restoreFocusWindow
                                              : _ownerWindow;
                SetFocus(restoreFocus);
            }
        });

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kFolderViewChangeCasePromptClassName,
                                          LoadStringResource(nullptr, IDS_CMD_CHANGE_CASE).c_str(),
                                          style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(_hWnd.get(), _ownerWindow));
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return std::nullopt;
            }
            if (getMessageResult == 0)
            {
                _done = true;
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<FolderViewChangeCasePromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<FolderViewChangeCasePromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            if (message == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                self->_done = true;
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
#ifdef ENABLE_TESTS
            case WndMsg::kFolderViewChangeCasePromptDebug: return self->OnDebugCommand(static_cast<FolderViewChangeCasePromptDebugCommand>(wParam), lParam);
#endif
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    struct StyleChoice final
    {
        ChangeCase::CaseStyle style;
        const wchar_t* label;
    };

    struct TargetChoice final
    {
        ChangeCase::ChangeTarget target;
        const wchar_t* label;
    };

    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = FolderViewChangeCasePromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kFolderViewChangeCasePromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        ApplyTheme();
        Layout();
        if (_styleCombo)
        {
            _dxHost.SetFocusControl(_styleCombo);
        }
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi() noexcept
    {
        if (_root)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _styleLabel = _root->AddChild<Label>(L"Change case to");
        _styleLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _styleCombo = _root->AddChild<ComboBox>();
        _styleCombo->SetVariant(ComboBoxVariant::Window);
        _styleCombo->SetEditable(false);

        std::vector<ComboBox::Item> styleItems;
        styleItems.reserve(std::size(kStyleChoices));
        for (const auto& choice : kStyleChoices)
        {
            styleItems.push_back(ComboBox::Item{std::wstring(choice.label), std::wstring(choice.label)});
        }
        _styleCombo->SetItems(std::move(styleItems));
        _styleCombo->SetSelectedIndex(0u);
        _styleCombo->SetOnSelectionChanged([this](size_t) noexcept { UpdateExampleLabel(); });
        _styleCombo->SetOnSubmitted([this] { Confirm(); });

        _targetLabel = _root->AddChild<Label>(L"Change");
        _targetLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _targetCombo = _root->AddChild<ComboBox>();
        _targetCombo->SetVariant(ComboBoxVariant::Window);
        _targetCombo->SetEditable(false);

        std::vector<ComboBox::Item> targetItems;
        targetItems.reserve(std::size(kTargetChoices));
        for (const auto& choice : kTargetChoices)
        {
            targetItems.push_back(ComboBox::Item{std::wstring(choice.label), std::wstring(choice.label)});
        }
        _targetCombo->SetItems(std::move(targetItems));
        _targetCombo->SetSelectedIndex(0u);
        _targetCombo->SetOnSelectionChanged([this](size_t) noexcept { UpdateExampleLabel(); });

        _exampleLabel = _root->AddChild<Label>(BuildExampleText());
        _exampleLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _exampleLabel->SetFontRole(FontRole::Body);

        _includeSubdirsToggle = _root->AddChild<Toggle>(L"Include subdirectories");
        _includeSubdirsToggle->SetChecked(false);
        _includeSubdirsToggle->SetEnabled(_allowSubdirs);
        _includeSubdirsToggle->SetOnToggled([this](bool checked) noexcept
        {
            _includeSubdirs = checked;
            if (_includeSubdirsToggle)
            {
                _includeSubdirsToggle->SetChecked(checked);
            }
        });

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_hWnd)
        {
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
        }
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip        = 16.0f;
        constexpr float kGapDip           = 8.0f;
        constexpr float kLabelHeightDip   = 22.0f;
        constexpr float kRowHeightDip     = 34.0f;
        constexpr float kExampleHeightDip = 24.0f;
        constexpr float kToggleHeightDip  = 36.0f;
        constexpr float kButtonWidthDip   = 96.0f;
        constexpr float kButtonHeightDip  = 34.0f;

        const float left  = client.left + kMarginDip;
        const float right = std::max(left, client.right - kMarginDip);
        float y           = client.top + kMarginDip;

        if (_styleLabel)
        {
            _styleLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + kGapDip;

        if (_styleCombo)
        {
            _styleCombo->SetBounds(D2D1::RectF(left, y, right, y + kRowHeightDip));
        }
        y += kRowHeightDip + (kGapDip * 1.5f);

        if (_targetLabel)
        {
            _targetLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + kGapDip;

        if (_targetCombo)
        {
            _targetCombo->SetBounds(D2D1::RectF(left, y, right, y + kRowHeightDip));
        }
        y += kRowHeightDip + kGapDip;

        if (_exampleLabel)
        {
            _exampleLabel->SetBounds(D2D1::RectF(left, y, right, y + kExampleHeightDip));
        }
        y += kExampleHeightDip + (kGapDip * 1.5f);

        if (_includeSubdirsToggle)
        {
            _includeSubdirsToggle->SetBounds(D2D1::RectF(left, y, right, y + kToggleHeightDip));
        }

        const float buttonsTop = std::max(y + kToggleHeightDip + kGapDip, client.bottom - kMarginDip - kButtonHeightDip);
        const float cancelLeft = std::max(left, right - kButtonWidthDip);
        const float okLeft     = std::max(left, cancelLeft - kGapDip - kButtonWidthDip);

        if (_okButton)
        {
            _okButton->SetBounds(D2D1::RectF(okLeft, buttonsTop, okLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
    }

    void Confirm() noexcept
    {
        _result = BuildOptionsFromSelections();
        _done   = true;
        ClosePromptWindow();
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        ClosePromptWindow();
    }

    void ClosePromptWindow() noexcept
    {
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _dxHost.SetFocusControl(nullptr);
            _hWnd.reset();
        }
    }

    [[nodiscard]] size_t GetStyleIndex() const noexcept
    {
        return _styleCombo ? _styleCombo->GetSelectedIndex().value_or(0u) : 0u;
    }

    [[nodiscard]] size_t GetTargetIndex() const noexcept
    {
        return _targetCombo ? _targetCombo->GetSelectedIndex().value_or(0u) : 0u;
    }

    void SetSelections(size_t styleIndex, size_t targetIndex, bool includeSubdirs) noexcept
    {
        if (_styleCombo)
        {
            _styleCombo->SetSelectedIndex(std::min(styleIndex, std::size(kStyleChoices) - 1u));
        }
        if (_targetCombo)
        {
            _targetCombo->SetSelectedIndex(std::min(targetIndex, std::size(kTargetChoices) - 1u));
        }
        _includeSubdirs = _allowSubdirs && includeSubdirs;
        if (_includeSubdirsToggle)
        {
            _includeSubdirsToggle->SetChecked(_includeSubdirs);
        }
        UpdateExampleLabel();
    }

    [[nodiscard]] ChangeCase::Options BuildOptionsFromSelections() const noexcept
    {
        ChangeCase::Options options{};
        options.style          = kStyleChoices[std::min(GetStyleIndex(), std::size(kStyleChoices) - 1u)].style;
        options.target         = kTargetChoices[std::min(GetTargetIndex(), std::size(kTargetChoices) - 1u)].target;
        options.includeSubdirs = _allowSubdirs && _includeSubdirs;
        return options;
    }

    [[nodiscard]] std::wstring BuildExampleText() const
    {
        constexpr std::wstring_view kSampleName = L"Sample File.TXT";

        ChangeCase::Options options = BuildOptionsFromSelections();
        options.includeSubdirs      = false;

        const std::wstring before(kSampleName);
        const std::wstring after = ChangeCase::TransformLeafName(before, options);
        std::wstring text        = FormatStringResource(nullptr, IDS_CHANGE_CASE_EXAMPLE_FMT, before, after);
        if (text.empty())
        {
            text = before + L" -> " + after;
        }
        return text;
    }

    void UpdateExampleLabel() noexcept
    {
        if (_exampleLabel)
        {
            _exampleLabel->SetText(BuildExampleText());
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(FolderViewChangeCasePromptDebugCommand command, LPARAM lParam) noexcept
    {
        switch (command)
        {
            case FolderViewChangeCasePromptDebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<FolderViewChangeCasePromptDebugSnapshot*>(lParam);
                if (! snapshot)
                {
                    return FALSE;
                }

                *snapshot                         = FolderViewChangeCasePromptDebugSnapshot{};
                snapshot->usesDxUiHost            = _dxHost.GetHwnd() != nullptr;
                snapshot->visibleChildWindowCount = CountVisibleChildWindowsLocal(_hWnd.get());
                snapshot->includeSubdirsEnabled   = _allowSubdirs;
                snapshot->includeSubdirsChecked   = _includeSubdirs;
                snapshot->styleIndex              = GetStyleIndex();
                snapshot->targetIndex             = GetTargetIndex();
                snapshot->exampleText             = _exampleLabel ? std::wstring(_exampleLabel->GetText()) : std::wstring{};
                return TRUE;
            }
            case FolderViewChangeCasePromptDebugCommand::SetSelections:
            {
                const auto payload = UnpackFolderViewChangeCasePromptSelections(lParam);
                SetSelections(payload.styleIndex, payload.targetIndex, payload.includeSubdirs);
                return TRUE;
            }
            case FolderViewChangeCasePromptDebugCommand::Confirm: Confirm(); return TRUE;
            case FolderViewChangeCasePromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

    static constexpr std::array<StyleChoice, 4> kStyleChoices{{
        {ChangeCase::CaseStyle::Lower, L"Lower case"},
        {ChangeCase::CaseStyle::Upper, L"Upper case"},
        {ChangeCase::CaseStyle::PartiallyMixed, L"Partially mixed case"},
        {ChangeCase::CaseStyle::Mixed, L"Mixed case"},
    }};

    static constexpr std::array<TargetChoice, 3> kTargetChoices{{
        {ChangeCase::ChangeTarget::WholeFilename, L"Whole filename"},
        {ChangeCase::ChangeTarget::OnlyName, L"Only name"},
        {ChangeCase::ChangeTarget::OnlyExtension, L"Only extension"},
    }};

    HWND _ownerWindow        = nullptr;
    HWND _restoreFocusWindow = nullptr;
    AppTheme _theme{};
    bool _allowSubdirs   = false;
    bool _includeSubdirs = false;
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    RedSalamander::DxUi::ThemePalette _palette{};
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root                  = nullptr;
    RedSalamander::DxUi::Label* _styleLabel            = nullptr;
    RedSalamander::DxUi::ComboBox* _styleCombo         = nullptr;
    RedSalamander::DxUi::Label* _targetLabel           = nullptr;
    RedSalamander::DxUi::ComboBox* _targetCombo        = nullptr;
    RedSalamander::DxUi::Label* _exampleLabel          = nullptr;
    RedSalamander::DxUi::Toggle* _includeSubdirsToggle = nullptr;
    RedSalamander::DxUi::Button* _okButton             = nullptr;
    RedSalamander::DxUi::Button* _cancelButton         = nullptr;
    bool _done                                         = false;
    std::optional<ChangeCase::Options> _result;
};

std::optional<ChangeCase::Options> PromptForChangeCase(HWND ownerWindow, const AppTheme& theme, bool allowSubdirs) noexcept
{
    FolderViewChangeCasePromptWindow prompt(ownerWindow, theme, allowSubdirs);
    return prompt.ShowModal();
}

std::optional<std::wstring> PromptForSelectionMask(
    HWND ownerWindow, const std::vector<std::wstring>& history, const AppTheme& theme, UINT captionId, UINT labelId)
{
    FolderViewSelectionMaskPromptWindow prompt(ownerWindow, history, theme, LoadStringResource(nullptr, captionId), LoadStringResource(nullptr, labelId));
    return prompt.ShowModal();
}

std::optional<FolderView::NameFilterState> PromptForPaneFilter(HWND ownerWindow,
                                                               const std::vector<std::wstring>& history,
                                                               const AppTheme& theme,
                                                               const FolderView::NameFilterState& initial)
{
    FolderViewPaneFilterPromptWindow prompt(ownerWindow, history, theme, initial);
    return prompt.ShowModal();
}

FolderView::SortDirection DefaultSortDirectionFor(FolderView::SortBy sortBy) noexcept
{
    switch (sortBy)
    {
        case FolderView::SortBy::Time:
        case FolderView::SortBy::Size: return FolderView::SortDirection::Descending;
        case FolderView::SortBy::Name:
        case FolderView::SortBy::Extension:
        case FolderView::SortBy::Attributes:
        case FolderView::SortBy::None: return FolderView::SortDirection::Ascending;
    }
    return FolderView::SortDirection::Ascending;
}
} // namespace FolderWindowFileSystemInternal

using namespace FolderWindowFileSystemInternal;

LRESULT FolderWindow::OnChangeCaseTaskUpdate(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<ChangeCaseTaskPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    return static_cast<LRESULT>(CreateOrUpdateInformationalTask(payload->update));
}

LRESULT FolderWindow::OnChangeCaseCompleted(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<ChangeCaseCompletedPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    PaneState& state = payload->pane == Pane::Left ? _leftPane : _rightPane;
    if (state.changeCaseThread.joinable())
    {
        state.changeCaseThread = {};
    }

    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    if (FAILED(payload->hr) && payload->hr != cancelledHr && payload->hr != E_ABORT)
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message = FormatStringResource(nullptr, IDS_FMT_PANE_CHANGE_CASE_FAILED, static_cast<unsigned long>(payload->hr));
        state.folderView.ShowAlertOverlay(
            FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message), payload->hr);
        MessageBeep(MB_ICONERROR);
        return 0;
    }

    if (SUCCEEDED(payload->hr))
    {
        state.folderView.ForceRefresh();
    }

    return 0;
}

HRESULT FolderWindow::EnsurePaneFileSystem(Pane pane, std::wstring_view pluginId) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const PaneState& other = pane == Pane::Left ? _rightPane : _leftPane;

    FileSystemPluginManager& plugins                  = FileSystemPluginManager::GetInstance();
    const auto& allPlugins                            = plugins.GetPlugins();
    const FileSystemPluginManager::PluginEntry* entry = FindPluginById(allPlugins, pluginId);

    if (pluginId.empty())
    {
        state.folderView.CancelPendingEnumeration();

        wil::unique_hmodule previousModule = std::move(state.fileSystemModule);
        wil::com_ptr<IFileSystem> previous = std::move(state.fileSystem);

        state.fileSystem = nullptr;
        state.fileSystemModule.reset();
        state.pluginId.clear();
        state.pluginShortId.clear();
        state.instanceContext.clear();

        state.folderView.SetFileSystem(state.fileSystem);
        state.folderView.SetFileSystemContext(state.pluginId, state.instanceContext);
        state.navigationView.SetFileSystem(state.fileSystem);

        if (previous && (! other.fileSystem || other.fileSystem.get() != previous.get()))
        {
            DirectoryInfoCache::GetInstance().UnregisterProvider(previous.get());
        }

        previous.reset(); // release before module unload
        state.folderView.ForceRefresh();
        return S_FALSE;
    }

    if (! entry || entry->id.empty() || entry->disabled || ! entry->loadable || ! entry->fileSystem)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (state.fileSystem && EqualsNoCase(state.pluginId, pluginId))
    {
        state.pluginShortId = entry->shortId;
        DirectoryInfoCache::GetInstance().RegisterProvider(state.fileSystem.get(), state.pluginId, state.pluginShortId, state.instanceContext);

        wil::com_ptr<IInformations> informationsInstance;
        const HRESULT qiInfos = state.fileSystem->QueryInterface(__uuidof(IInformations), informationsInstance.put_void());
        if (SUCCEEDED(qiInfos) && informationsInstance && entry->informations)
        {
            const char* configuration = nullptr;
            static_cast<void>(entry->informations->GetConfiguration(&configuration));
            static_cast<void>(informationsInstance->SetConfiguration(configuration));
        }
        return S_OK;
    }

    if (entry->path.empty())
    {
        return E_FAIL;
    }

    wil::unique_hmodule keepAlive(LoadLibrary(entry->path.c_str()));
    if (! keepAlive)
    {
        const DWORD lastError = Debug::ErrorWithLastError(L"FolderWindow: Failed to LoadLibrary '{}' for keep-alive", entry->path.c_str());
        return HRESULT_FROM_WIN32(lastError);
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(keepAlive.get(), "RedSalamanderCreate"));
#pragma warning(pop)
    if (! createFactory)
    {
        DWORD lastError = GetLastError();
        if (lastError == ERROR_SUCCESS)
        {
            lastError = ERROR_PROC_NOT_FOUND;
        }
        Debug::Error(L"FolderWindow: Missing export RedSalamanderCreate in '{}'", entry->path.c_str());
        return HRESULT_FROM_WIN32(lastError);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr<IFileSystem> fileSystemInstance;
    const std::wstring requestedPluginId = entry->factoryPluginId.empty() ? entry->id : entry->factoryPluginId;
    if (requestedPluginId.empty())
    {
        Debug::Error(L"FolderWindow: Missing logical plugin id for '{}'", entry->path.c_str());
        return E_INVALIDARG;
    }
    const HRESULT createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), requestedPluginId.c_str(), fileSystemInstance.put_void());
    if (FAILED(createHr) || ! fileSystemInstance)
    {
        Debug::Error(L"FolderWindow: RedSalamanderCreate failed for '{}' (pluginId={} hr=0x{:08X})",
                     entry->path.c_str(),
                     requestedPluginId,
                     static_cast<unsigned long>(createHr));
        return FAILED(createHr) ? createHr : E_FAIL;
    }

    wil::com_ptr<IInformations> informationsInstance;
    const HRESULT qiInfos = fileSystemInstance->QueryInterface(__uuidof(IInformations), informationsInstance.put_void());
    if (FAILED(qiInfos) || ! informationsInstance)
    {
        Debug::Error(L"FolderWindow: IInformations not supported by '{}' (hr=0x{:08X})", entry->path.c_str(), static_cast<unsigned long>(qiInfos));
        return FAILED(qiInfos) ? qiInfos : E_NOINTERFACE;
    }

    const char* configuration = nullptr;
    if (entry->informations)
    {
        static_cast<void>(entry->informations->GetConfiguration(&configuration));
    }
    if (configuration && configuration[0] != '\0')
    {
        static_cast<void>(informationsInstance->SetConfiguration(configuration));
    }

    state.folderView.CancelPendingEnumeration();

    wil::unique_hmodule previousModule = std::move(state.fileSystemModule);
    wil::com_ptr<IFileSystem> previous = std::move(state.fileSystem);

    state.fileSystem       = std::move(fileSystemInstance);
    state.fileSystemModule = std::move(keepAlive);
    state.pluginId         = entry->id;
    state.pluginShortId    = entry->shortId;
    state.instanceContext.clear();

    state.folderView.SetFileSystem(state.fileSystem);
    state.folderView.SetFileSystemContext(state.pluginId, state.instanceContext);
    state.navigationView.SetFileSystem(state.fileSystem);
    DirectoryInfoCache::GetInstance().RegisterProvider(state.fileSystem.get(), state.pluginId, state.pluginShortId, state.instanceContext);

    if (previous && previous.get() != state.fileSystem.get() && (! other.fileSystem || other.fileSystem.get() != previous.get()))
    {
        DirectoryInfoCache::GetInstance().UnregisterProvider(previous.get());
    }

    previous.reset(); // release before module unload
    return S_OK;
}

HRESULT FolderWindow::ReloadFileSystemPlugins() noexcept
{
    const std::wstring_view defaultPluginId = FileSystemPluginManager::GetInstance().GetActivePluginId();

    if (_leftPane.pluginId.empty())
    {
        _leftPane.pluginId = std::wstring(defaultPluginId);
    }
    if (_rightPane.pluginId.empty())
    {
        _rightPane.pluginId = std::wstring(defaultPluginId);
    }

    const HRESULT leftHr  = EnsurePaneFileSystem(Pane::Left, _leftPane.pluginId);
    const HRESULT rightHr = EnsurePaneFileSystem(Pane::Right, _rightPane.pluginId);

    if (FAILED(leftHr) && ! defaultPluginId.empty())
    {
        static_cast<void>(SetFileSystemPluginForPane(Pane::Left, defaultPluginId));
    }
    if (FAILED(rightHr) && ! defaultPluginId.empty())
    {
        static_cast<void>(SetFileSystemPluginForPane(Pane::Right, defaultPluginId));
    }
    return S_OK;
}

void FolderWindow::ReleaseFileSystemPluginsForRefresh() noexcept
{
    std::array<IFileSystem*, 2> providers{_leftPane.fileSystem.get(), _rightPane.fileSystem.get()};
    for (size_t i = 0; i < providers.size(); ++i)
    {
        IFileSystem* provider = providers[i];
        if (! provider)
        {
            continue;
        }

        bool alreadySeen = false;
        for (size_t j = 0; j < i; ++j)
        {
            if (providers[j] == provider)
            {
                alreadySeen = true;
                break;
            }
        }
        if (! alreadySeen)
        {
            DirectoryInfoCache::GetInstance().UnregisterProvider(provider);
        }
    }

    auto releasePane = [](PaneState& state) noexcept
    {
        wil::com_ptr<IFileSystem> emptyFileSystem;
        state.folderView.SetFileSystem(emptyFileSystem);
        state.navigationView.SetFileSystem(emptyFileSystem);
        state.fileSystem.reset();
        // Release the pane's explicit LoadLibrary pin before plugin-manager refresh unloads/reloads DLLs.
        state.fileSystemModule.reset();
    };

    releasePane(_leftPane);
    releasePane(_rightPane);
}

HRESULT FolderWindow::SetFileSystemPluginForPane(Pane pane, std::wstring_view pluginId) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (state.fileSystem && ! state.pluginId.empty() && EqualsNoCase(state.pluginId, pluginId))
    {
        state.navigationView.SetHistory(_folderHistory);
        if (state.currentPath.has_value())
        {
            state.navigationView.SetPath(state.currentPath);
        }
        return S_FALSE;
    }

    const HRESULT hr = EnsurePaneFileSystem(pane, pluginId);
    if (FAILED(hr))
    {
        return hr;
    }

    const bool isFile = IsFilePluginShortId(state.pluginShortId);
    if (isFile)
    {
        const std::optional<std::filesystem::path> current = state.folderView.GetFolderPath();
        if (current && LooksLikeWindowsAbsolutePath(current.value().wstring()))
        {
            SetFolderPath(pane, current.value());
        }
        else
        {
            SetFolderPath(pane, GetDefaultFileSystemRoot());
        }
        return S_OK;
    }

    SetFolderPath(pane, std::filesystem::path(std::wstring(state.pluginShortId) + L":/"));
    return S_OK;
}

std::wstring_view FolderWindow::GetFileSystemPluginId(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.pluginId;
}

std::wstring_view FolderWindow::GetFileSystemPluginShortId(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.pluginShortId;
}

std::wstring_view FolderWindow::GetFileSystemInstanceContext(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.instanceContext;
}

wil::com_ptr<IFileSystem> FolderWindow::GetFileSystem(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.fileSystem;
}

HRESULT FolderWindow::SetFileSystemInstanceForPane(
    Pane pane, wil::com_ptr<IFileSystem> fileSystem, std::wstring pluginId, std::wstring pluginShortId, std::wstring instanceContext) noexcept
{
    PaneState& state       = pane == Pane::Left ? _leftPane : _rightPane;
    const PaneState& other = pane == Pane::Left ? _rightPane : _leftPane;

    state.folderView.CancelPendingEnumeration();

    wil::unique_hmodule previousModule = std::move(state.fileSystemModule);
    wil::com_ptr<IFileSystem> previous = std::move(state.fileSystem);

    state.fileSystem = std::move(fileSystem);
    state.fileSystemModule.reset();
    state.pluginId        = std::move(pluginId);
    state.pluginShortId   = std::move(pluginShortId);
    state.instanceContext = std::move(instanceContext);
    state.currentPath.reset();
    state.updatingPath = false;

    state.folderView.SetFileSystem(state.fileSystem);
    state.folderView.SetFileSystemContext(state.pluginId, state.instanceContext);
    state.navigationView.SetFileSystem(state.fileSystem);
    DirectoryInfoCache::GetInstance().RegisterProvider(state.fileSystem.get(), state.pluginId, state.pluginShortId, state.instanceContext);

    if (previous && previous.get() != state.fileSystem.get() && (! other.fileSystem || other.fileSystem.get() != previous.get()))
    {
        DirectoryInfoCache::GetInstance().UnregisterProvider(previous.get());
    }

    previous.reset(); // release before module unload
    return S_OK;
}

HRESULT FolderWindow::ExecuteInActivePane(const std::filesystem::path& folderPath,
                                          std::wstring_view focusItemDisplayName,
                                          unsigned int folderViewCommandId,
                                          bool activateWindow,
                                          std::wstring_view navigateToPaintMetricName,
                                          std::chrono::steady_clock::time_point inputStartedAt) noexcept
{
    return ExecuteInPane(_activePane, folderPath, focusItemDisplayName, folderViewCommandId, activateWindow, navigateToPaintMetricName, inputStartedAt);
}

HRESULT FolderWindow::ExecuteInPane(Pane pane,
                                    const std::filesystem::path& folderPath,
                                    std::wstring_view focusItemDisplayName,
                                    unsigned int folderViewCommandId,
                                    bool activateWindow,
                                    std::wstring_view navigateToPaintMetricName,
                                    std::chrono::steady_clock::time_point inputStartedAt) noexcept
{
    if (folderPath.empty())
    {
        return E_INVALIDARG;
    }

    Debug::Perf::Scope perf(pane == Pane::Left ? L"FolderWindow.ExecuteInPane.Left" : L"FolderWindow.ExecuteInPane.Right");
    perf.SetDetail(folderPath.native());

    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const auto activateRootWindow = [&]() noexcept
    {
        if (! activateWindow)
        {
            return;
        }

        const HWND root = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
        const HWND wnd  = root ? root : _hWnd.get();
        if (wnd)
        {
            if (IsIconic(wnd))
            {
                ShowWindow(wnd, SW_RESTORE);
            }
            else
            {
                ShowWindow(wnd, SW_SHOWNORMAL);
            }

            SetForegroundWindow(wnd);
        }
    };

    const auto focusFolderView = [&]() noexcept
    {
        if (state.hFolderView && IsWindow(state.hFolderView.get()))
        {
            SetFocus(state.hFolderView.get());
        }
    };

    const auto flushPaneVisuals = [&]() noexcept
    {
        if (state.hNavigationView && IsWindow(state.hNavigationView.get()))
        {
            UpdateWindow(state.hNavigationView.get());
        }
        if (state.hFolderView && IsWindow(state.hFolderView.get()))
        {
            UpdateWindow(state.hFolderView.get());
        }
    };

    const std::optional<std::filesystem::path> currentFolder = state.folderView.GetFolderPath();

    bool sameFolder = false;
    if (currentFolder.has_value())
    {
        const std::wstring_view currentText = currentFolder.value().native();
        const std::wstring_view targetText  = folderPath.native();

        if (IsFilePluginShortId(state.pluginShortId))
        {
            sameFolder = EqualsNoCase(currentText, targetText);
        }
        else
        {
            sameFolder = currentText == targetText;
        }
    }

    if (sameFolder)
    {
        state.pendingNavigationToPaintMetric.reset();
        bool ready = true;
        if (! focusItemDisplayName.empty())
        {
            ready = state.folderView.PrepareForExternalCommand(focusItemDisplayName);
        }

        if (ready && folderViewCommandId != 0u && state.hFolderView)
        {
            activateRootWindow();
            focusFolderView();
            PostMessageW(state.hFolderView.get(), WM_COMMAND, MAKEWPARAM(folderViewCommandId, 0), 0);
            return S_OK;
        }

        if (ready && folderViewCommandId == 0u)
        {
            activateRootWindow();
            focusFolderView();
            flushPaneVisuals();
            return S_OK;
        }

        if (! focusItemDisplayName.empty())
        {
            state.folderView.RememberFocusedItemForFolder(folderPath, focusItemDisplayName);
        }
        if (folderViewCommandId != 0u)
        {
            state.folderView.QueueCommandAfterNextEnumeration(folderViewCommandId, folderPath, focusItemDisplayName);
        }

        state.folderView.ForceRefresh();
        activateRootWindow();
        focusFolderView();
        flushPaneVisuals();
        return S_OK;
    }

    if (! focusItemDisplayName.empty())
    {
        state.folderView.RememberFocusedItemForFolder(folderPath, focusItemDisplayName);
    }
    if (folderViewCommandId != 0u)
    {
        state.folderView.QueueCommandAfterNextEnumeration(folderViewCommandId, folderPath, focusItemDisplayName);
    }

    if (! navigateToPaintMetricName.empty() && inputStartedAt != std::chrono::steady_clock::time_point{})
    {
        state.pendingNavigationToPaintMetric = PaneState::PendingNavigationToPaintMetric{
            .targetFolder = folderPath,
            .startedAt    = inputStartedAt,
            .metricName   = std::wstring(navigateToPaintMetricName),
            .detail       = folderPath.native(),
            .value0       = focusItemDisplayName.empty() ? 0u : 1u,
            .value1       = folderViewCommandId,
        };
    }
    else
    {
        state.pendingNavigationToPaintMetric.reset();
    }

    SetFolderPath(pane, folderPath);
    activateRootWindow();
    focusFolderView();
    flushPaneVisuals();
    return S_OK;
}

HRESULT FolderWindow::ExecuteInPaneLocation(Pane pane,
                                            std::wstring_view pluginId,
                                            std::wstring_view pluginShortId,
                                            std::wstring_view instanceContext,
                                            const std::filesystem::path& folderPath,
                                            std::wstring_view focusItemDisplayName,
                                            unsigned int folderViewCommandId,
                                            bool activateWindow) noexcept
{
    if (pluginId.empty() || pluginShortId.empty() || folderPath.empty())
    {
        return E_INVALIDARG;
    }

    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    const std::wstring previousPluginId(state.pluginId);
    const std::wstring previousPluginShortId(state.pluginShortId);
    const std::wstring previousInstanceContext(state.instanceContext);
    const std::optional<std::filesystem::path> previousFolderPath = state.folderView.GetFolderPath();

    auto restorePreviousLocation = [&]() noexcept
    {
        if (previousPluginId.empty())
        {
            return;
        }

        if (! NavigationLocation::EqualsNoCase(state.pluginId, previousPluginId))
        {
            const HRESULT restoreHr = SetFileSystemPluginForPane(pane, previousPluginId);
            if (FAILED(restoreHr))
            {
                Debug::Error(L"File Operations completed action could not restore pane provider: 0x{:08X}", restoreHr);
                return;
            }
        }

        if (! previousPluginShortId.empty() && previousFolderPath.has_value())
        {
            SetFolderPath(pane, NavigationLocation::FormatHistoryPath(previousPluginShortId, previousInstanceContext, previousFolderPath.value()));
        }
    };

    const std::filesystem::path qualifiedPath = NavigationLocation::FormatHistoryPath(pluginShortId, instanceContext, folderPath);
    SetFolderPath(pane, qualifiedPath);
    if (! NavigationLocation::EqualsNoCase(state.pluginId, pluginId) || ! NavigationLocation::EqualsNoCase(state.pluginShortId, pluginShortId) ||
        ! NavigationLocation::EqualsNoCase(state.instanceContext, instanceContext))
    {
        restorePreviousLocation();
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    return ExecuteInPane(pane, folderPath, focusItemDisplayName, folderViewCommandId, activateWindow);
}

void FolderWindow::SetFolderPath(const std::filesystem::path& path)
{
    SetFolderPath(_activePane, path);
}

void FolderWindow::SetFolderPath(Pane pane, const std::filesystem::path& path)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.updatingPath)
    {
        return;
    }

    const std::optional<std::filesystem::path> previousPluginPath = state.folderView.GetFolderPath();

    FileSystemPluginManager& pluginManager  = FileSystemPluginManager::GetInstance();
    const auto& plugins                     = pluginManager.GetPlugins();
    const std::wstring_view defaultPluginId = pluginManager.GetActivePluginId();

    std::wstring pluginId;
    std::wstring pluginShortId;
    std::wstring remainder;
    std::wstring instanceContext;
    bool instanceContextSpecified = false;

    const std::wstring text = path.wstring();

    Debug::Perf::Scope perf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left" : L"FolderWindow.SetFolderPath.Right");
    perf.SetDetail(text);

    auto tryResolveConnectionNameToTarget = [&](std::wstring_view connectionName, std::wstring_view overridePluginPath, std::wstring& outTarget) -> bool
    {
        outTarget.clear();

        if (! _settings || connectionName.empty())
        {
            return false;
        }

        Common::Settings::ConnectionProfile quick{};
        const Common::Settings::ConnectionProfile* profile = nullptr;

        if (RedSalamander::Connections::IsQuickConnectConnectionName(connectionName))
        {
            const std::wstring_view preferredPluginId = defaultPluginId.empty() ? pluginManager.GetActivePluginId() : defaultPluginId;
            RedSalamander::Connections::EnsureQuickConnectProfile(preferredPluginId);
            RedSalamander::Connections::GetQuickConnectProfile(quick);
            profile = &quick;
        }
        else if (_settings->connections)
        {
            const auto& conns = _settings->connections->items;
            const auto it     = std::find_if(conns.begin(), conns.end(), [&](const Common::Settings::ConnectionProfile& c) noexcept {
                return ! c.name.empty() && EqualsNoCase(c.name, connectionName);
            });
            if (it != conns.end())
            {
                profile = &(*it);
            }
        }

        if (! profile || profile->pluginId.empty())
        {
            return false;
        }

        const FileSystemPluginManager::PluginEntry* navEntry = FindPluginById(plugins, profile->pluginId);
        if (! navEntry || navEntry->shortId.empty())
        {
            return false;
        }

        std::wstring initial = profile->initialPath.empty() ? L"/" : profile->initialPath;
        if (! initial.empty() && initial.front() != L'/')
        {
            initial.insert(initial.begin(), L'/');
        }

        std::wstring_view pluginPath = initial;
        if (! overridePluginPath.empty())
        {
            pluginPath = overridePluginPath;
        }

        std::wstring normalized = NavigationLocation::NormalizePluginPathText(pluginPath,
                                                                              NavigationLocation::EmptyPathPolicy::Root,
                                                                              NavigationLocation::LeadingSlashPolicy::Ensure,
                                                                              NavigationLocation::TrailingSlashPolicy::Preserve);
        if (normalized.empty())
        {
            normalized = L"/";
        }

        outTarget.reserve(navEntry->shortId.size() + 16u + profile->name.size() + normalized.size());
        outTarget.append(navEntry->shortId);
        outTarget.append(L":/@conn:");
        outTarget.append(profile->name);
        outTarget.append(normalized);
        return true;
    };

    auto openConnectionManagerAndNavigate = [&](std::wstring_view filterPluginId) noexcept
    {
        if (! _settings)
        {
            return;
        }

        static_cast<void>(ShowConnectionManagerWindow(_hWnd.get(), *this, L"RedSalamander", *_settings, _theme, filterPluginId, static_cast<uint8_t>(pane)));
    };

    auto parseNavConnectionName = [&](std::wstring_view rawNavText, std::wstring& outConnectionName, std::wstring& outPathOverride) -> bool
    {
        outConnectionName.clear();
        outPathOverride.clear();

        std::wstring_view name = rawNavText;
        while (! name.empty() && std::iswspace(name.front()))
        {
            name.remove_prefix(1);
        }
        while (! name.empty() && std::iswspace(name.back()))
        {
            name.remove_suffix(1);
        }

        if (name.size() >= 2u && name[0] == L'/' && name[1] == L'/')
        {
            name.remove_prefix(2u);
        }
        else if (! name.empty() && name.front() == L'/')
        {
            name.remove_prefix(1u);
        }

        const size_t slash               = name.find_first_of(L"/\\");
        const std::wstring_view connName = slash == std::wstring_view::npos ? name : name.substr(0, slash);
        const std::wstring_view pathPart = slash == std::wstring_view::npos ? std::wstring_view{} : name.substr(slash);

        if (! connName.empty())
        {
            outConnectionName.assign(connName);
        }

        if (! pathPart.empty())
        {
            outPathOverride.assign(pathPart);
        }

        return true;
    };

    if (StartsWithNoCase(text, L"nav:") || StartsWithNoCase(text, L"@conn:"))
    {
        const bool isConnPrefix        = StartsWithNoCase(text, L"@conn:");
        const std::wstring_view suffix = isConnPrefix ? std::wstring_view(text).substr(6) : std::wstring_view(text).substr(4);

        std::wstring connectionName;
        std::wstring pathOverride;
        static_cast<void>(parseNavConnectionName(suffix, connectionName, pathOverride));

        if (connectionName.empty())
        {
            openConnectionManagerAndNavigate({});
            return;
        }

        std::wstring target;
        if (tryResolveConnectionNameToTarget(connectionName, pathOverride, target))
        {
            SetFolderPath(pane, std::filesystem::path(std::move(target)));
            return;
        }

        HostAlertRequest request{};
        request.version      = 1;
        request.sizeBytes    = sizeof(request);
        request.scope        = HOST_ALERT_SCOPE_APPLICATION;
        request.modality     = HOST_ALERT_MODELESS;
        request.severity     = HOST_ALERT_ERROR;
        request.targetWindow = nullptr;
        request.title        = nullptr;
        request.message      = L"Connection not found.";
        request.closable     = TRUE;
        static_cast<void>(HostShowAlert(request));
        return;
    }

    bool hasPluginPrefix = false;
    std::filesystem::path pluginPath;
    {
        Debug::Perf::Scope parsePerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.Parse" : L"FolderWindow.SetFolderPath.Right.Parse");
        parsePerf.SetDetail(text);

        hasPluginPrefix = TryParsePluginPrefix(text, pluginShortId, remainder);
        parsePerf.SetValue0(hasPluginPrefix ? 1u : 0u);

        if (hasPluginPrefix)
        {
            const bool supportsConnections =
                (EqualsNoCase(pluginShortId, L"ftp") || EqualsNoCase(pluginShortId, L"sftp") || EqualsNoCase(pluginShortId, L"scp") ||
                 EqualsNoCase(pluginShortId, L"imap") || EqualsNoCase(pluginShortId, L"gdrive") || EqualsNoCase(pluginShortId, L"onedrive") ||
                 EqualsNoCase(pluginShortId, L"onedrive-pro") || EqualsNoCase(pluginShortId, L"sharepoint") || EqualsNoCase(pluginShortId, L"s3") ||
                 EqualsNoCase(pluginShortId, L"s3table"));

            const auto openProtocolFilteredConnectionManager = [&]
            {
                const FileSystemPluginManager::PluginEntry* shortEntry = FindPluginByShortId(plugins, pluginShortId);
                if (shortEntry && ! shortEntry->id.empty())
                {
                    openConnectionManagerAndNavigate(shortEntry->id);
                }
            };

            if (supportsConnections)
            {
                // Treat protocol roots like `ftp:` / `gdrive:` and `ftp://@conn` / `gdrive://@conn`
                // as explicit Connection Manager entry points.
                std::wstring_view check = remainder;
                if (check.empty())
                {
                    openProtocolFilteredConnectionManager();
                    return;
                }

                const auto tryStripConnAuthority = [&](std::wstring_view value, std::wstring_view& outRest) noexcept -> bool
                {
                    outRest = {};
                    if (value.size() < 7u)
                    {
                        return false;
                    }

                    // Accept both `//@conn` and `\\@conn` (depending on how the path string is formed).
                    if (! ((value[0] == L'/' || value[0] == L'\\') && (value[1] == L'/' || value[1] == L'\\')))
                    {
                        return false;
                    }

                    std::wstring_view afterSlashes         = value.substr(2);
                    constexpr std::wstring_view kAuthority = L"@conn";
                    if (afterSlashes.size() < kAuthority.size() || ! EqualsNoCase(afterSlashes.substr(0, kAuthority.size()), kAuthority))
                    {
                        return false;
                    }

                    if (afterSlashes.size() == kAuthority.size() || afterSlashes[kAuthority.size()] == L'/' || afterSlashes[kAuthority.size()] == L'\\')
                    {
                        outRest = afterSlashes.substr(kAuthority.size());
                        return true;
                    }

                    return false;
                };

                std::wstring_view restAfterAuthority;
                if (tryStripConnAuthority(check, restAfterAuthority))
                {
                    std::wstring_view rest = restAfterAuthority;
                    while (! rest.empty() && (rest.front() == L'/' || rest.front() == L'\\'))
                    {
                        rest.remove_prefix(1u);
                    }

                    const size_t slash                     = rest.find_first_of(L"/\\");
                    const std::wstring_view connectionName = slash == std::wstring_view::npos ? rest : rest.substr(0, slash);
                    const std::wstring_view remotePart     = slash == std::wstring_view::npos ? std::wstring_view{} : rest.substr(slash);

                    if (connectionName.empty())
                    {
                        openProtocolFilteredConnectionManager();
                        return;
                    }

                    std::wstring target;
                    target.reserve(pluginShortId.size() + 16u + connectionName.size() + remotePart.size());
                    target.append(pluginShortId);
                    target.append(L":/@conn:");
                    target.append(connectionName);
                    if (remotePart.empty())
                    {
                        target.append(L"/");
                    }
                    else
                    {
                        std::wstring normalized = NavigationLocation::NormalizePluginPathText(remotePart,
                                                                                              NavigationLocation::EmptyPathPolicy::Root,
                                                                                              NavigationLocation::LeadingSlashPolicy::Ensure,
                                                                                              NavigationLocation::TrailingSlashPolicy::Preserve);
                        if (normalized.empty())
                        {
                            normalized = L"/";
                        }
                        target.append(normalized);
                    }

                    SetFolderPath(pane, std::filesystem::path(std::move(target)));
                    return;
                }
            }

            std::wstring_view pluginPathText = remainder;
            const size_t bar                 = remainder.find(L'|');
            if (bar != std::wstring::npos)
            {
                instanceContextSpecified = true;
                instanceContext          = remainder.substr(0, bar);
                pluginPathText           = std::wstring_view(remainder).substr(bar + 1);
            }
            else if (EqualsNoCase(pluginShortId, L"7z") && ! pluginPathText.empty() && pluginPathText.front() != L'/' && pluginPathText.front() != L'\\')
            {
                // Shorthand mount syntax: "7z:<zipPath>" mounts <zipPath> and opens "/".
                instanceContextSpecified = true;
                instanceContext          = std::wstring(pluginPathText);
                pluginPathText           = L"/";

                if (! LooksLikeWindowsAbsolutePath(instanceContext))
                {
                    const std::optional<std::filesystem::path> baseFolder = state.folderView.GetFolderPath();
                    if (baseFolder.has_value() && IsFilePluginShortId(state.pluginShortId))
                    {
                        std::filesystem::path resolved = baseFolder.value() / std::filesystem::path(instanceContext);
                        resolved                       = resolved.lexically_normal();
                        instanceContext                = resolved.wstring();
                    }
                }
            }

            if (IsFilePluginShortId(pluginShortId))
            {
                std::filesystem::path parsed;
                if (NavigationLocation::TryParseFileUriRemainder(pluginPathText, parsed))
                {
                    pluginPath = std::move(parsed);
                }
                else
                {
                    std::wstring win(pluginPathText);
                    for (wchar_t& ch : win)
                    {
                        if (ch == L'/')
                        {
                            ch = L'\\';
                        }
                    }
                    pluginPath = std::filesystem::path(std::move(win));
                }
            }
            else
            {
                pluginPath = NavigationLocation::NormalizePluginPath(pluginPathText);
            }
        }
        else
        {
            if (LooksLikeWindowsAbsolutePath(text))
            {
                pluginShortId = L"file";
            }
            else if (! state.pluginId.empty())
            {
                pluginId        = state.pluginId;
                pluginShortId   = state.pluginShortId;
                instanceContext = state.instanceContext;
            }
            else if (! defaultPluginId.empty())
            {
                pluginId = std::wstring(defaultPluginId);
            }
            else
            {
                pluginShortId = L"file";
            }

            pluginPath = path;
        }
    }

    const auto isUsable = [](const FileSystemPluginManager::PluginEntry* candidate) noexcept
    { return candidate && ! candidate->id.empty() && candidate->loadable && ! candidate->disabled && candidate->fileSystem; };

    {
        Debug::Perf::Scope resolvePerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.ResolvePlugin"
                                                          : L"FolderWindow.SetFolderPath.Right.ResolvePlugin");
        resolvePerf.SetDetail(text);

        const FileSystemPluginManager::PluginEntry* entry = nullptr;
        if (! pluginShortId.empty())
        {
            entry = FindPluginByShortId(plugins, pluginShortId);
        }

        if (! isUsable(entry))
        {
            entry = nullptr;
        }

        if (! entry && ! pluginId.empty())
        {
            entry = FindPluginById(plugins, pluginId);
        }

        if (! isUsable(entry))
        {
            entry = nullptr;
        }

        if (! entry && ! defaultPluginId.empty())
        {
            entry = FindPluginById(plugins, defaultPluginId);
        }

        if (! isUsable(entry))
        {
            entry = nullptr;
        }

        if (! entry)
        {
            return;
        }

        pluginId      = entry->id;
        pluginShortId = entry->shortId;
        resolvePerf.SetDetail(pluginId);

        if (! IsFilePluginShortId(pluginShortId))
        {
            pluginPath = NavigationLocation::NormalizePluginPath(pluginPath.wstring());
        }

        if (IsFilePluginShortId(pluginShortId) && ! LooksLikeWindowsAbsolutePath(pluginPath.native()))
        {
            pluginPath = GetDefaultFileSystemRoot();
        }
    }

    // Skip the navigation pipeline only when the pane already shows the identical local path AND the
    // last enumeration of that folder succeeded. A failed navigation (access denied, offline share,
    // ejected media) leaves `_currentFolder` pointing at the failed path, so retrying it must re-enumerate.
    if (state.fileSystem && IsFilePluginShortId(state.pluginShortId) && IsFilePluginShortId(pluginShortId) && EqualsNoCase(state.pluginId, pluginId) &&
        state.instanceContext.empty() && instanceContext.empty() && previousPluginPath.has_value() &&
        IsSameLocalNavigationPathText(previousPluginPath.value(), pluginPath) && state.folderView.IsCurrentFolderEnumerated())
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(std::format(L"FolderWindow::SetFolderPath skipped same local path pane={} previous='{}' requested='{}'",
                                                  pane == Pane::Left ? L"left" : L"right",
                                                  previousPluginPath->wstring(),
                                                  pluginPath.wstring()));
#endif
        return;
    }

    Debug::Info(L"FolderWindow::SetFolderPath resolved input='{}' pluginId='{}' pluginShortId='{}' instanceContext='{}' pluginPath='{}'",
                text,
                pluginId,
                pluginShortId,
                instanceContext.empty() ? std::wstring_view(L"<none>") : std::wstring_view(instanceContext),
                pluginPath.native());

    {
        Debug::Perf::Scope ensurePerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.EnsurePaneFileSystem"
                                                         : L"FolderWindow.SetFolderPath.Right.EnsurePaneFileSystem");
        ensurePerf.SetDetail(pluginId);

        HRESULT pluginHr = EnsurePaneFileSystem(pane, pluginId);
        if (FAILED(pluginHr) && ! defaultPluginId.empty() && ! EqualsNoCase(pluginId, defaultPluginId))
        {
            const FileSystemPluginManager::PluginEntry* fallback = FindPluginById(plugins, defaultPluginId);
            if (isUsable(fallback))
            {
                pluginId      = fallback->id;
                pluginShortId = fallback->shortId;

                if (IsFilePluginShortId(pluginShortId))
                {
                    pluginPath = GetDefaultFileSystemRoot();
                }
                else
                {
                    pluginPath = std::filesystem::path(L"/");
                }

                ensurePerf.SetDetail(pluginId);
                pluginHr = EnsurePaneFileSystem(pane, pluginId);
            }
        }

        ensurePerf.SetHr(pluginHr);

        if (FAILED(pluginHr))
        {
            Debug::Error(L"FolderWindow::SetFolderPath: Failed to ensure pane file system `{}`", pluginId);
            return;
        }
    }

    {
        Debug::Perf::Scope initPerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.InitializeFileSystem"
                                                       : L"FolderWindow.SetFolderPath.Right.InitializeFileSystem");
        initPerf.SetDetail(pluginId);

        if (state.fileSystem)
        {
            wil::com_ptr<IFileSystemInitialize> initializer;
            const HRESULT initQi = state.fileSystem->QueryInterface(__uuidof(IFileSystemInitialize), initializer.put_void());
            if (SUCCEEDED(initQi) && initializer)
            {
                if (! instanceContextSpecified && instanceContext.empty())
                {
                    instanceContext = state.instanceContext;
                }

                const bool contextSame = EqualsNoCase(state.instanceContext, instanceContext);
                if (! instanceContext.empty() && ! contextSame)
                {
                    state.instanceContext = instanceContext;
                    DirectoryInfoCache::GetInstance().RegisterProvider(state.fileSystem.get(), state.pluginId, state.pluginShortId, state.instanceContext);
                    static_cast<void>(initializer->Initialize(state.instanceContext.c_str(), nullptr));
                }
                else if (instanceContextSpecified && instanceContext.empty() && ! state.instanceContext.empty())
                {
                    state.instanceContext.clear();
                    DirectoryInfoCache::GetInstance().RegisterProvider(state.fileSystem.get(), state.pluginId, state.pluginShortId, std::wstring_view{});
                }
            }
            else
            {
                state.instanceContext.clear();
            }
        }
    }

    // Keep FolderView informed so it can include mount context in internal drag/drop formats.
    state.folderView.SetFileSystemContext(state.pluginId, state.instanceContext);
    if (state.fileSystem)
    {
        DirectoryInfoCache::GetInstance().RegisterProvider(state.fileSystem.get(), state.pluginId, state.pluginShortId, state.instanceContext);
    }

    const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, pluginPath);

    {
        Debug::Perf::Scope updatePerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.UpdateViews" : L"FolderWindow.SetFolderPath.Right.UpdateViews");
        updatePerf.SetDetail(displayPath.native());

        state.updatingPath = true;
        state.currentPath  = displayPath;

        Debug::Perf::Scope navPerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.UpdateViews.NavigationView.SetPath"
                                                      : L"FolderWindow.SetFolderPath.Right.UpdateViews.NavigationView.SetPath");
        navPerf.SetDetail(displayPath.native());
        state.navigationView.SetPath(displayPath);

        if (state.hFolderView)
        {
            Debug::Perf::Scope viewPerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.UpdateViews.FolderView.SetFolderPath"
                                                           : L"FolderWindow.SetFolderPath.Right.UpdateViews.FolderView.SetFolderPath");
            viewPerf.SetDetail(pluginPath.native());

            const Common::Settings::FoldersSettings* folders = (_settings && _settings->folders.has_value()) ? &_settings->folders.value() : nullptr;
            const FolderView::NameFilterState filter         = GetFolderHistoryFilterState(folders, displayPath);
            state.folderView.SetNameFilterState(filter, false /* refresh */);
            UpdatePaneFilterBar(pane);
            state.folderView.SetFolderPath(pluginPath);
        }

        state.updatingPath = false;
    }

    {
        Debug::Perf::Scope historyPerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.UpdateHistory"
                                                          : L"FolderWindow.SetFolderPath.Right.UpdateHistory");
        historyPerf.SetDetail(displayPath.native());

        RecordNavigationHistory(state, displayPath);
        AddToFolderHistory(_folderHistory, static_cast<size_t>(_folderHistoryMax), displayPath);
        _leftPane.navigationView.SetHistory(_folderHistory);
        _rightPane.navigationView.SetHistory(_folderHistory);

        if (_settings)
        {
            Common::Settings::FoldersSettings& folders = _settings->folders.has_value() ? _settings->folders.value() : _settings->folders.emplace();
            PruneFolderHistoryFilters(folders, _folderHistory, static_cast<size_t>(_folderHistoryMax));
        }
    }

    if (_panePathChangedCallback)
    {
        const bool changed = ! previousPluginPath.has_value() || previousPluginPath->native() != pluginPath.native();
        if (changed)
        {
            _panePathChangedCallback(pane, pluginPath);
        }
    }
}

bool FolderWindow::TryOpenFileAsVirtualFileSystem(Pane pane, const std::filesystem::path& path) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! IsFilePluginShortId(state.pluginShortId))
    {
        return true;
    }

    if (! _settings)
    {
        return false;
    }

    std::wstring extension = path.extension().wstring();
    if (extension.empty())
    {
        return false;
    }

    std::transform(
        extension.begin(), extension.end(), extension.begin(), [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))); });

    const auto it = _settings->extensions.openWithFileSystemByExtension.find(extension);
    if (it == _settings->extensions.openWithFileSystemByExtension.end())
    {
        return false;
    }

    const std::wstring_view pluginId = it->second;
    if (pluginId.empty())
    {
        return false;
    }

    FileSystemPluginManager& pluginManager            = FileSystemPluginManager::GetInstance();
    const auto& plugins                               = pluginManager.GetPlugins();
    const FileSystemPluginManager::PluginEntry* entry = FindPluginById(plugins, pluginId);

    const auto isUsable = [](const FileSystemPluginManager::PluginEntry* candidate) noexcept
    { return candidate && ! candidate->id.empty() && candidate->loadable && ! candidate->disabled && candidate->fileSystem && ! candidate->shortId.empty(); };
    if (! isUsable(entry))
    {
        return false;
    }

    const std::wstring filePath = path.wstring();
    if (filePath.empty())
    {
        return false;
    }

    std::wstring mountPath;
    mountPath.reserve(entry->shortId.size() + 1u + filePath.size() + 2u);
    mountPath.append(entry->shortId);
    mountPath.push_back(L':');
    mountPath.append(filePath);
    mountPath.append(L"|/");

    SetFolderPath(pane, std::filesystem::path(mountPath));
    return true;
}

std::optional<std::filesystem::path> FolderWindow::GetCurrentPath() const
{
    return GetCurrentPath(_activePane);
}

std::optional<std::filesystem::path> FolderWindow::GetCurrentPluginPath() const
{
    return GetCurrentPluginPath(_activePane);
}

std::optional<std::filesystem::path> FolderWindow::GetCurrentPath(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.currentPath;
}

std::optional<std::filesystem::path> FolderWindow::GetCurrentPluginPath(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetFolderPath();
}

std::optional<std::filesystem::path> FolderWindow::GetFocusedItemPath(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetFocusedPath();
}

std::vector<std::filesystem::path> FolderWindow::GetSelectedOrFocusedPaths(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetSelectedOrFocusedPaths();
}

std::vector<std::filesystem::path> FolderWindow::GetFolderHistory() const
{
    return _folderHistory;
}

std::vector<std::filesystem::path> FolderWindow::GetFolderHistory(Pane pane) const
{
    static_cast<void>(pane);
    return _folderHistory;
}

void FolderWindow::SetFolderHistory(const std::vector<std::filesystem::path>& history)
{
    _folderHistory = history;
    NormalizeFolderHistory(_folderHistory, static_cast<size_t>(_folderHistoryMax));

    _leftPane.navigationView.SetHistory(_folderHistory);
    _rightPane.navigationView.SetHistory(_folderHistory);

    if (_settings)
    {
        Common::Settings::FoldersSettings& folders = _settings->folders.has_value() ? _settings->folders.value() : _settings->folders.emplace();
        PruneFolderHistoryFilters(folders, _folderHistory, static_cast<size_t>(_folderHistoryMax));
    }
}

void FolderWindow::SetFolderHistory(Pane pane, const std::vector<std::filesystem::path>& history)
{
    static_cast<void>(pane);
    SetFolderHistory(history);
}

uint32_t FolderWindow::GetFolderHistoryMax() const noexcept
{
    return _folderHistoryMax;
}

void FolderWindow::SetFolderHistoryMax(uint32_t maxItems)
{
    _folderHistoryMax = std::clamp(maxItems, 1u, kFolderHistoryMaxMax);
    NormalizeFolderHistory(_folderHistory, static_cast<size_t>(_folderHistoryMax));

    _leftPane.navigationView.SetHistory(_folderHistory);
    _rightPane.navigationView.SetHistory(_folderHistory);

    if (_settings)
    {
        Common::Settings::FoldersSettings& folders = _settings->folders.has_value() ? _settings->folders.value() : _settings->folders.emplace();
        PruneFolderHistoryFilters(folders, _folderHistory, static_cast<size_t>(_folderHistoryMax));
    }

    TrimNavigationHistory(_leftPane);
    TrimNavigationHistory(_rightPane);
}

void FolderWindow::TrimNavigationHistory(PaneState& state)
{
    const size_t maxItems = static_cast<size_t>(_folderHistoryMax);
    if (maxItems == 0)
    {
        state.navigationHistory.clear();
        state.navigationHistoryIndex = 0;
        return;
    }

    if (state.navigationHistory.empty())
    {
        state.navigationHistoryIndex = 0;
        return;
    }

    if (state.navigationHistoryIndex >= state.navigationHistory.size())
    {
        state.navigationHistoryIndex = state.navigationHistory.size() - 1;
    }

    if (state.navigationHistory.size() <= maxItems)
    {
        return;
    }

    const size_t trimCount = state.navigationHistory.size() - maxItems;
    state.navigationHistory.erase(state.navigationHistory.begin(), state.navigationHistory.begin() + static_cast<std::ptrdiff_t>(trimCount));

    if (state.navigationHistoryIndex >= trimCount)
    {
        state.navigationHistoryIndex -= trimCount;
    }
    else
    {
        state.navigationHistoryIndex = 0;
    }
}

void FolderWindow::RecordNavigationHistory(PaneState& state, const std::filesystem::path& displayPath)
{
    if (state.navigationHistorySuspendRecord)
    {
        return;
    }

    if (displayPath.empty())
    {
        return;
    }

    if (state.navigationHistory.empty())
    {
        state.navigationHistory.push_back(displayPath);
        state.navigationHistoryIndex = 0;
        TrimNavigationHistory(state);
        return;
    }

    if (state.navigationHistoryIndex >= state.navigationHistory.size())
    {
        state.navigationHistoryIndex = state.navigationHistory.size() - 1;
    }

    const std::wstring_view entryText   = displayPath.native();
    const std::wstring_view currentText = state.navigationHistory[state.navigationHistoryIndex].native();
    if (EqualsNoCase(currentText, entryText))
    {
        return;
    }

    const size_t nextIndex = state.navigationHistoryIndex + 1;
    if (nextIndex < state.navigationHistory.size())
    {
        state.navigationHistory.erase(state.navigationHistory.begin() + static_cast<std::ptrdiff_t>(nextIndex), state.navigationHistory.end());
    }

    state.navigationHistory.push_back(displayPath);
    state.navigationHistoryIndex = state.navigationHistory.size() - 1;
    TrimNavigationHistory(state);
}

void FolderWindow::SetDisplayMode(Pane pane, FolderView::DisplayMode mode)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetDisplayMode(mode);
}

FolderView::DisplayMode FolderWindow::GetDisplayMode(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetDisplayMode();
}

void FolderWindow::SetSort(Pane pane, FolderView::SortBy sortBy, FolderView::SortDirection direction)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetSort(sortBy, direction);
    UpdatePaneStatusBar(pane);
}

void FolderWindow::CycleSortBy(Pane pane, FolderView::SortBy sortBy)
{
    const FolderView::SortBy currentBy         = GetSortBy(pane);
    const FolderView::SortDirection currentDir = GetSortDirection(pane);
    const FolderView::SortDirection defaultDir = DefaultSortDirectionFor(sortBy);

    if (currentBy != sortBy)
    {
        SetSort(pane, sortBy, defaultDir);
        return;
    }

    if (currentDir == defaultDir)
    {
        const FolderView::SortDirection flipped =
            defaultDir == FolderView::SortDirection::Ascending ? FolderView::SortDirection::Descending : FolderView::SortDirection::Ascending;
        SetSort(pane, sortBy, flipped);
        return;
    }

    SetSort(pane, sortBy, defaultDir);
}

FolderView::SortBy FolderWindow::GetSortBy(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetSortBy();
}

FolderView::SortDirection FolderWindow::GetSortDirection(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetSortDirection();
}

void FolderWindow::SetShowHiddenFiles(bool show)
{
    if (_showHiddenFiles == show)
    {
        return;
    }

    _showHiddenFiles = show;
    _leftPane.folderView.SetShowHiddenFiles(show);
    _rightPane.folderView.SetShowHiddenFiles(show);
}

bool FolderWindow::GetShowHiddenFiles() const noexcept
{
    return _showHiddenFiles;
}

void FolderWindow::SetShowSystemFiles(bool show)
{
    if (_showSystemFiles == show)
    {
        return;
    }

    _showSystemFiles = show;
    _leftPane.folderView.SetShowSystemFiles(show);
    _rightPane.folderView.SetShowSystemFiles(show);
}

bool FolderWindow::GetShowSystemFiles() const noexcept
{
    return _showSystemFiles;
}

#ifdef ENABLE_TESTS
HWND FindDebugPromptWindowForCurrentProcess(const wchar_t* className) noexcept
{
    if (! className || *className == L'\0')
    {
        return nullptr;
    }

    struct SearchState
    {
        DWORD processId          = 0;
        const wchar_t* className = nullptr;
        HWND hwnd                = nullptr;
    } state{GetCurrentProcessId(), className, nullptr};

    EnumWindows(
        [](HWND hwnd, LPARAM lParam) noexcept -> BOOL
    {
        auto* state = reinterpret_cast<SearchState*>(lParam);
        if (! state || ! hwnd || IsWindow(hwnd) == FALSE || IsWindowVisible(hwnd) == FALSE)
        {
            return TRUE;
        }

        DWORD processId = 0;
        GetWindowThreadProcessId(hwnd, &processId);
        if (processId != state->processId)
        {
            return TRUE;
        }

        wchar_t windowClass[128]{};
        if (GetClassNameW(hwnd, windowClass, static_cast<int>(std::size(windowClass))) == 0)
        {
            return TRUE;
        }

        if (wcscmp(windowClass, state->className) != 0)
        {
            return TRUE;
        }

        state->hwnd = hwnd;
        return FALSE;
    },
        reinterpret_cast<LPARAM>(&state));

    return state.hwnd;
}

HWND GetFolderViewPaneFilterPromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewPaneFilterPromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetFolderViewPaneFilterPromptSnapshot(FolderViewPaneFilterPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    const bool ok =
        SendMessageW(hwnd, WndMsg::kFolderViewPaneFilterPromptDebug, static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::GetSnapshot), 0) != FALSE;
    if (ok)
    {
        const std::scoped_lock lock(g_folderViewPaneFilterPromptDebugMutex);
        if (! g_folderViewPaneFilterPromptDebugSnapshot.has_value())
        {
            return false;
        }
        out = g_folderViewPaneFilterPromptDebugSnapshot.value();
    }
    return ok;
}

bool DebugSetFolderViewPaneFilterPromptEnabled(bool enabled) noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    return hwnd && SendMessageW(hwnd,
                                WndMsg::kFolderViewPaneFilterPromptDebug,
                                static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::SetEnabled),
                                enabled ? 1 : 0) != FALSE;
}

bool DebugSetFolderViewPaneFilterPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    {
        const std::scoped_lock lock(g_folderViewPaneFilterPromptDebugMutex);
        g_folderViewPaneFilterPromptDebugText.assign(text);
    }
    return SendMessageW(hwnd, WndMsg::kFolderViewPaneFilterPromptDebug, static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::SetText), 0) != FALSE;
}

bool DebugSetFolderViewPaneFilterPromptTextAndNotify(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    {
        const std::scoped_lock lock(g_folderViewPaneFilterPromptDebugMutex);
        g_folderViewPaneFilterPromptDebugText.assign(text);
    }
    return SendMessageW(hwnd, WndMsg::kFolderViewPaneFilterPromptDebug, static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::SetTextAndNotify), 0) !=
           FALSE;
}

bool DebugSetFolderViewPaneFilterPromptHelpExpanded(bool expanded) noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    return hwnd && SendMessageW(hwnd,
                                WndMsg::kFolderViewPaneFilterPromptDebug,
                                static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::SetHelpExpanded),
                                expanded ? 1 : 0) != FALSE;
}

bool DebugConfirmFolderViewPaneFilterPrompt() noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    return PostDxUiPromptCloseDebugCommand(
        hwnd, WndMsg::kFolderViewPaneFilterPromptDebug, static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::Confirm));
}

bool DebugCancelFolderViewPaneFilterPrompt() noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    return PostDxUiPromptCloseDebugCommand(hwnd, WndMsg::kFolderViewPaneFilterPromptDebug, static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::Cancel));
}

HWND GetFolderViewSelectionMaskPromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewSelectionMaskPromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

HWND GetFolderViewCreateDirectoryPromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewCreateDirectoryPromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

HWND GetFolderViewEditNewPromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewEditNewPromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetFolderViewSelectionMaskPromptSnapshot(FolderViewSelectionMaskPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewSelectionMaskPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FolderViewSelectionMaskPromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 WndMsg::kFolderViewSelectionMaskPromptDebug,
                                 static_cast<WPARAM>(FolderViewSelectionMaskPromptDebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFolderViewSelectionMaskPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewSelectionMaskPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(text);
    return SendMessageW(hwnd,
                        WndMsg::kFolderViewSelectionMaskPromptDebug,
                        static_cast<WPARAM>(FolderViewSelectionMaskPromptDebugCommand::SetText),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugConfirmFolderViewSelectionMaskPrompt() noexcept
{
    const HWND hwnd = GetFolderViewSelectionMaskPromptHandle();
    return PostDxUiPromptCloseDebugCommand(
        hwnd, WndMsg::kFolderViewSelectionMaskPromptDebug, static_cast<WPARAM>(FolderViewSelectionMaskPromptDebugCommand::Confirm));
}

bool DebugCancelFolderViewSelectionMaskPrompt() noexcept
{
    const HWND hwnd = GetFolderViewSelectionMaskPromptHandle();
    return PostDxUiPromptCloseDebugCommand(
        hwnd, WndMsg::kFolderViewSelectionMaskPromptDebug, static_cast<WPARAM>(FolderViewSelectionMaskPromptDebugCommand::Cancel));
}

bool DebugGetFolderViewCreateDirectoryPromptSnapshot(FolderViewCreateDirectoryPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewCreateDirectoryPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FolderViewCreateDirectoryPromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 GetFolderViewCreateDirectoryPromptDebugMessage(),
                                 static_cast<WPARAM>(FolderViewCreateDirectoryPromptDebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFolderViewCreateDirectoryPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewCreateDirectoryPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(text);
    return SendMessageW(hwnd,
                        GetFolderViewCreateDirectoryPromptDebugMessage(),
                        static_cast<WPARAM>(FolderViewCreateDirectoryPromptDebugCommand::SetText),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugConfirmFolderViewCreateDirectoryPrompt() noexcept
{
    const HWND hwnd = GetFolderViewCreateDirectoryPromptHandle();
    return PostDxUiPromptCloseDebugCommand(
        hwnd, GetFolderViewCreateDirectoryPromptDebugMessage(), static_cast<WPARAM>(FolderViewCreateDirectoryPromptDebugCommand::Confirm));
}

bool DebugCancelFolderViewCreateDirectoryPrompt() noexcept
{
    const HWND hwnd = GetFolderViewCreateDirectoryPromptHandle();
    return PostDxUiPromptCloseDebugCommand(
        hwnd, GetFolderViewCreateDirectoryPromptDebugMessage(), static_cast<WPARAM>(FolderViewCreateDirectoryPromptDebugCommand::Cancel));
}

bool DebugGetFolderViewEditNewPromptSnapshot(FolderViewEditNewPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FolderViewEditNewPromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 GetFolderViewEditNewPromptDebugMessage(),
                                 static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFolderViewEditNewPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(text);
    return SendMessageW(hwnd,
                        GetFolderViewEditNewPromptDebugMessage(),
                        static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::SetText),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugSelectFolderViewEditNewPromptEditor(std::wstring_view actionId) noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(actionId);
    return SendMessageW(hwnd,
                        GetFolderViewEditNewPromptDebugMessage(),
                        static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::SelectEditor),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugConfirmFolderViewEditNewPrompt() noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    return PostDxUiPromptCloseDebugCommand(hwnd, GetFolderViewEditNewPromptDebugMessage(), static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::Confirm));
}

bool DebugCancelFolderViewEditNewPrompt() noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    return PostDxUiPromptCloseDebugCommand(hwnd, GetFolderViewEditNewPromptDebugMessage(), static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::Cancel));
}

HWND GetFolderViewChangeCasePromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewChangeCasePromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetFolderViewChangeCasePromptSnapshot(FolderViewChangeCasePromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewChangeCasePromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FolderViewChangeCasePromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 WndMsg::kFolderViewChangeCasePromptDebug,
                                 static_cast<WPARAM>(FolderViewChangeCasePromptDebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFolderViewChangeCasePromptSelections(size_t styleIndex, size_t targetIndex, bool includeSubdirs) noexcept
{
    const HWND hwnd = GetFolderViewChangeCasePromptHandle();
    if (! hwnd)
    {
        return false;
    }

    return SendMessageW(hwnd,
                        WndMsg::kFolderViewChangeCasePromptDebug,
                        static_cast<WPARAM>(FolderViewChangeCasePromptDebugCommand::SetSelections),
                        PackFolderViewChangeCasePromptSelections(styleIndex, targetIndex, includeSubdirs)) != FALSE;
}

bool DebugConfirmFolderViewChangeCasePrompt() noexcept
{
    const HWND hwnd = GetFolderViewChangeCasePromptHandle();
    return PostDxUiPromptCloseDebugCommand(
        hwnd, WndMsg::kFolderViewChangeCasePromptDebug, static_cast<WPARAM>(FolderViewChangeCasePromptDebugCommand::Confirm));
}

bool DebugCancelFolderViewChangeCasePrompt() noexcept
{
    const HWND hwnd = GetFolderViewChangeCasePromptHandle();
    return PostDxUiPromptCloseDebugCommand(hwnd, WndMsg::kFolderViewChangeCasePromptDebug, static_cast<WPARAM>(FolderViewChangeCasePromptDebugCommand::Cancel));
}
#endif
