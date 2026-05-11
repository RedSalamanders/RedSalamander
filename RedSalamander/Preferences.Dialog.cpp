#include "Framework.h"

#include "D2DHdcPaint.h"
#include "DxUi/DxUi.Typography.h"
#include "Preferences.Advanced.h"
#include "Preferences.CompareDirectories.h"
#include "Preferences.Dialog.h"
#include "Preferences.Editors.h"
#include "Preferences.FileOperations.h"
#include "Preferences.General.h"
#include "Preferences.HotPaths.h"
#include "Preferences.Internal.h"
#include "Preferences.Keyboard.h"
#include "Preferences.Mouse.h"
#include "Preferences.Panes.h"
#include "Preferences.Plugins.h"
#include "Preferences.Themes.h"
#include "Preferences.UserMenu.h"
#include "Preferences.Viewers.h"
#include "Preferences.h"

#include <algorithm>
#include <array>
#include <charconv>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cwchar>
#include <cwctype>
#include <filesystem>
#include <format>
#include <functional>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <tuple>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include <uxtheme.h>
#include <windowsx.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include "CommandRegistry.h"
#include "DxUi/DxUi.Win32Hooks.h"
#include "DxUi/DxUi.h"
#include "FluentIcons.h"
#include "Helpers.h"
#include "HostServices.h"
#include "SettingsHotReload.h"
#include "SettingsSave.h"
#include "ShortcutDefaults.h"
#include "ShortcutManager.h"
#include "ShortcutText.h"
#include "UiMetrics.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "resource.h"
#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#endif

namespace
{
constexpr UINT kPrefsDeferredCloseMessage                   = WM_APP + 0x43u;
thread_local bool g_preferencesWheelMessageForwarded        = false;
constexpr wchar_t kPrefsWheelRouteOriginalWndProcProp[]     = L"RedSalamander.Preferences.WheelRouteOriginalWndProc";
constexpr wchar_t kPrefsDxCategoryHostOriginalWndProcProp[] = L"RedSalamander.Preferences.DxCategoryHostOriginalWndProc";
constexpr wchar_t kPrefsDxCategoryHostStateProp[]           = L"RedSalamander.Preferences.DxCategoryHostState";
constexpr wchar_t kPrefsDxShellHostOriginalWndProcProp[]    = L"RedSalamander.Preferences.DxShellHostOriginalWndProc";
constexpr wchar_t kPrefsDxShellHostProp[]                   = L"RedSalamander.Preferences.DxShellHost";
constexpr wchar_t kPrefsDxDiagnosticsProp[]                 = L"RedSalamander.Preferences.DxDiagnostics";

using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::CallStoredWndProc;
using RedSalamander::DxUi::Control;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::GetStoredWndProc;
using RedSalamander::DxUi::IDxTreeDelegate;
using RedSalamander::DxUi::IDxTreeModel;
using RedSalamander::DxUi::InstallWndProcHook;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::ScrollPanel;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Tree;
using RedSalamander::DxUi::TreeItemData;
using RedSalamander::DxUi::WindowHost;

[[maybe_unused]] [[nodiscard]] uint32_t PackColorArgb(const D2D1_COLOR_F& color) noexcept
{
    const auto pack = [](const float value) noexcept -> uint8_t
    {
        const float scaled = std::clamp(value, 0.0f, 1.0f) * 255.0f;
        return static_cast<uint8_t>(std::lround(scaled));
    };

    return (static_cast<uint32_t>(pack(color.a)) << 24u) | (static_cast<uint32_t>(pack(color.r)) << 16u) | (static_cast<uint32_t>(pack(color.g)) << 8u) |
           static_cast<uint32_t>(pack(color.b));
}

void UpdatePageText(HWND dlg, PreferencesDialogState& state, PrefCategory category) noexcept;
void RefreshActivePreferencesPage(HWND host, PreferencesDialogState& state) noexcept;
void RebindActivePreferencesPage(HWND dlg, PreferencesDialogState& state) noexcept;
bool RequestPreferencesDialogClose(HWND dlg) noexcept;
bool HandleDeferredPaneAction(HWND dlg, PreferencesDialogState& state, PreferencesDeferredActionPayload&& payload) noexcept;
struct PreferencesPageHostSurfaceControl;
LRESULT CALLBACK PreferencesPageHostWindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
void InstallWheelRoutingHooks(HWND dlg) noexcept;
[[nodiscard]] PreferencesDialogState* GetState(HWND dlg) noexcept;
void RestoreWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp) noexcept;

[[nodiscard]] const wchar_t* GetPrefCategoryDebugName(PrefCategory category) noexcept;
void SyncPreferencesPageHostDxSize(HWND host, PreferencesDialogState& state, const wchar_t* reason) noexcept;
void LogPreferencesPageHostState(HWND host, const PreferencesDialogState& state, const wchar_t* reason) noexcept;
[[nodiscard]] bool IsPreferencesDxDiagnosticsEnabled(HWND hwnd) noexcept;

struct CategoryInfo
{
    PrefCategory id{};
    UINT labelId{};
    UINT descriptionId{};
};

struct ScopedWindowRedrawBlock final
{
    explicit ScopedWindowRedrawBlock(HWND hwnd) noexcept : _hwnd(hwnd)
    {
        if (_hwnd)
        {
            const auto currentCount = reinterpret_cast<ULONG_PTR>(GetPropW(_hwnd, PrefsUi::kPrefsTreeRedrawBlockProp));
            SetPropW(_hwnd, PrefsUi::kPrefsTreeRedrawBlockProp, reinterpret_cast<HANDLE>(currentCount + 1u));
            SendMessageW(_hwnd, WM_SETREDRAW, FALSE, 0);
        }
    }

    ScopedWindowRedrawBlock(const ScopedWindowRedrawBlock&)            = delete;
    ScopedWindowRedrawBlock& operator=(const ScopedWindowRedrawBlock&) = delete;

    ~ScopedWindowRedrawBlock()
    {
        Enable();
    }

    void Enable() noexcept
    {
        if (! _hwnd)
        {
            return;
        }

        const HWND hwnd         = _hwnd;
        _hwnd                   = nullptr;
        const auto currentCount = reinterpret_cast<ULONG_PTR>(GetPropW(hwnd, PrefsUi::kPrefsTreeRedrawBlockProp));
        bool enableRedraw       = false;
        if (currentCount <= 1u)
        {
            RemovePropW(hwnd, PrefsUi::kPrefsTreeRedrawBlockProp);
            enableRedraw = true;
        }
        else
        {
            SetPropW(hwnd, PrefsUi::kPrefsTreeRedrawBlockProp, reinterpret_cast<HANDLE>(currentCount - 1u));
        }

        if (enableRedraw)
        {
            SendMessageW(hwnd, WM_SETREDRAW, TRUE, 0);
        }
    }

    void EnableAndRedraw(const UINT flags) noexcept
    {
        if (! _hwnd)
        {
            return;
        }

        const HWND hwnd = _hwnd;
        Enable();
        RedrawWindow(hwnd, nullptr, nullptr, flags);
    }

private:
    HWND _hwnd = nullptr;
};

constexpr std::array<CategoryInfo, 13> kCategories = {{
    {PrefCategory::General, IDS_PREFS_CAT_GENERAL, IDS_PREFS_CAT_GENERAL_DESC},
    {PrefCategory::Panes, IDS_PREFS_CAT_PANES, IDS_PREFS_CAT_PANES_DESC},
    {PrefCategory::Viewers, IDS_PREFS_CAT_VIEWERS, IDS_PREFS_CAT_VIEWERS_DESC},
    {PrefCategory::Editors, IDS_PREFS_CAT_EDITORS, IDS_PREFS_CAT_EDITORS_DESC},
    {PrefCategory::UserMenu, IDS_PREFS_CAT_USER_MENU, IDS_PREFS_CAT_USER_MENU_DESC},
    {PrefCategory::Keyboard, IDS_PREFS_CAT_KEYBOARD, IDS_PREFS_CAT_KEYBOARD_DESC},
    {PrefCategory::Mouse, IDS_PREFS_CAT_MOUSE, IDS_PREFS_CAT_MOUSE_DESC},
    {PrefCategory::Themes, IDS_PREFS_CAT_THEMES, IDS_PREFS_CAT_THEMES_DESC},
    {PrefCategory::Plugins, IDS_PREFS_CAT_PLUGINS, IDS_PREFS_CAT_PLUGINS_DESC},
    {PrefCategory::FileOperations, IDS_PREFS_CAT_FILE_OPERATIONS, IDS_PREFS_CAT_FILE_OPERATIONS_DESC},
    {PrefCategory::CompareDirectories, IDS_PREFS_CAT_COMPARE_DIRECTORIES, IDS_PREFS_CAT_COMPARE_DIRECTORIES_DESC},
    {PrefCategory::HotPaths, IDS_PREFS_CAT_HOT_PATHS, IDS_PREFS_CAT_HOT_PATHS_DESC},
    {PrefCategory::Advanced, IDS_PREFS_CAT_ADVANCED, IDS_PREFS_CAT_ADVANCED_DESC},
}};

struct PreferencesDialogHost final : PreferencesDialogState
{
    PreferencesDialogHost()                                        = default;
    PreferencesDialogHost(const PreferencesDialogHost&)            = delete;
    PreferencesDialogHost& operator=(const PreferencesDialogHost&) = delete;

    GeneralPane _generalPane;
    PanesPane _panesPane;
    ViewersPane _viewersPane;
    EditorsPane _editorsPane;
    UserMenuPane _userMenuPane;
    KeyboardPane _keyboardPane;
    MousePane _mousePane;
    ThemesPane _themesPane;
    PluginsPane _pluginsPane;
    FileOperationsPane _fileOperationsPane;
    CompareDirectoriesPane _compareDirectoriesPane;
    HotPathsPane _hotPathsPane;
    AdvancedPane _advancedPane;
    WindowHost _categoryTreeHost;
    Tree* _categoryTreeControl = nullptr;
    WindowHost _shellHost;
    WindowHost _pageHostHost;
    Panel* _pageHostRootControl                                = nullptr;
    Panel* _pageHostContentRootControl                         = nullptr;
    PreferencesPageHostSurfaceControl* _pageHostSurfaceControl = nullptr;
    Panel* _pageHostEmptyStatePanel                            = nullptr;
    Label* _pageHostEmptyStateIconControl                      = nullptr;
    Label* _pageHostEmptyStateTitleControl                     = nullptr;
    Label* _pageHostEmptyStateBodyControl                      = nullptr;
    Label* _pageHostEmptyStateCaptionControl                   = nullptr;
    HWND _shellHostHwnd                                        = nullptr;
    Label* _pageTitleControl                                   = nullptr;
    Label* _pageDescriptionControl                             = nullptr;
    Button* _okButtonControl                                   = nullptr;
    Button* _cancelButtonControl                               = nullptr;
    Button* _applyButtonControl                                = nullptr;
    Button* _resetAllButtonControl                             = nullptr;

    struct PreferencesCategoryTreeModel final : IDxTreeModel
    {
        void Rebuild() noexcept
        {
            _visibleItems.clear();
            _pluginItems.clear();
            _visibleItems.reserve(kCategories.size());

            for (const CategoryInfo& category : kCategories)
            {
                TreeItemData item{};
                item.id   = EncodeCategoryNodeId(category.id);
                item.text = LoadStringResource(nullptr, category.labelId);
                if (item.text.empty())
                {
                    item.text = L"?";
                }
                if (category.id == PrefCategory::Plugins)
                {
                    std::vector<PrefsPluginListItem> plugins;
                    PrefsPlugins::BuildListItems(plugins);
                    _pluginItems     = std::move(plugins);
                    item.hasChildren = ! _pluginItems.empty();
                    item.expanded    = _pluginsExpanded;
                    item.badgeTone   = RedSalamander::DxUi::AdornmentTone::Info;
                    if (! _pluginItems.empty())
                    {
                        item.badgeText = std::format(L"{}", _pluginItems.size());
                    }
                }
                _visibleItems.push_back(std::move(item));

                if (category.id == PrefCategory::Plugins && _pluginsExpanded)
                {
                    for (const PrefsPluginListItem& plugin : _pluginItems)
                    {
                        const std::wstring_view displayName = PrefsPlugins::GetDisplayName(plugin);
                        if (displayName.empty())
                        {
                            continue;
                        }

                        TreeItemData child{};
                        child.id        = EncodePluginNodeId(plugin.type, plugin.index);
                        child.parentId  = EncodeCategoryNodeId(PrefCategory::Plugins);
                        child.text      = std::wstring(displayName);
                        child.depth     = 1u;
                        child.badgeTone = RedSalamander::DxUi::AdornmentTone::Accent;
                        _visibleItems.push_back(std::move(child));
                    }
                }
            }
        }

        void SetPluginsExpanded(const bool expanded) noexcept
        {
            _pluginsExpanded = expanded;
        }

        [[nodiscard]] bool IsPluginsExpanded() const noexcept
        {
            return _pluginsExpanded;
        }

        [[nodiscard]] size_t GetPluginItemCount() const noexcept
        {
            return _pluginItems.size();
        }

        [[nodiscard]] size_t GetVisibleItemCount() const noexcept override
        {
            return _visibleItems.size();
        }

        void GetVisibleItem(const size_t visibleIndex, TreeItemData& outItem) const override
        {
            outItem = {};
            if (visibleIndex >= _visibleItems.size())
            {
                return;
            }
            outItem = _visibleItems[visibleIndex];
        }

        [[nodiscard]] static constexpr uint64_t EncodeCategoryNodeId(const PrefCategory category) noexcept
        {
            return 0x1000ull + static_cast<uint64_t>(category);
        }

        [[nodiscard]] static constexpr uint64_t EncodePluginNodeId(const PrefsPluginType type, const size_t index) noexcept
        {
            return (1ull << 63u) | (static_cast<uint64_t>(type) << 48u) | static_cast<uint64_t>(index);
        }

        [[nodiscard]] static bool TryDecodePluginNodeId(const uint64_t nodeId, PrefsPluginListItem& outItem) noexcept
        {
            if ((nodeId & (1ull << 63u)) == 0u)
            {
                return false;
            }

            outItem.type  = static_cast<PrefsPluginType>((nodeId >> 48u) & 0x7FFFu);
            outItem.index = static_cast<size_t>(nodeId & 0x0000FFFFFFFFFFFFull);
            return true;
        }

        [[nodiscard]] static std::optional<PrefCategory> TryDecodeCategoryNodeId(const uint64_t nodeId) noexcept
        {
            if (nodeId < 0x1000ull || (nodeId & (1ull << 63u)) != 0u)
            {
                return std::nullopt;
            }

            const auto category = static_cast<PrefCategory>(nodeId - 0x1000ull);
            for (const CategoryInfo& info : kCategories)
            {
                if (info.id == category)
                {
                    return category;
                }
            }
            return std::nullopt;
        }

    private:
        bool _pluginsExpanded = true;
        std::vector<TreeItemData> _visibleItems;
        std::vector<PrefsPluginListItem> _pluginItems;
    } _categoryTreeModel;

    struct PreferencesCategoryTreeDelegate final : IDxTreeDelegate
    {
        void Attach(HWND dialog, PreferencesDialogState* state, PreferencesCategoryTreeModel* model, Tree* tree) noexcept
        {
            _dialog = dialog;
            _state  = state;
            _model  = model;
            _tree   = tree;
        }

        void OnTreeSelectionChanged(const uint64_t itemId) override
        {
            if (! _dialog || ! _state || ! _model)
            {
                Debug::Warning(L"Preferences: OnTreeSelectionChanged ignored (null dialog/state/model)");
                return;
            }

            PrefsPluginListItem pluginItem{};
            if (PreferencesCategoryTreeModel::TryDecodePluginNodeId(itemId, pluginItem))
            {
                _state->pluginsSelectedPlugin = pluginItem;
                _state->pluginsSelectedPluginId.assign(PrefsPlugins::GetId(pluginItem));
                _state->pluginsRetainedSelectedPluginId.assign(PrefsPlugins::GetId(pluginItem));
                _state->pluginsDetailsActive = true;
                UpdatePageText(_dialog, *_state, PrefCategory::Plugins);
                return;
            }

            if (const std::optional<PrefCategory> category = PreferencesCategoryTreeModel::TryDecodeCategoryNodeId(itemId))
            {
                _state->pluginsSelectedPlugin.reset();
                _state->pluginsSelectedPluginId.clear();
                _state->pluginsRetainedSelectedPluginId.clear();
                _state->pluginsDetailsActive = false;
                UpdatePageText(_dialog, *_state, category.value());
            }
            else
            {
                Debug::Warning(L"Preferences: OnTreeSelectionChanged with unknown itemId={:#x}", itemId);
            }
        }

        void OnTreeToggleExpanded(const uint64_t itemId, const bool expanded) override
        {
            if (! _state || ! _model || ! _tree || itemId != PreferencesCategoryTreeModel::EncodeCategoryNodeId(PrefCategory::Plugins))
            {
                return;
            }

            _model->SetPluginsExpanded(expanded);
            if (! expanded)
            {
                _state->pluginsSelectedPlugin.reset();
                _state->pluginsSelectedPluginId.clear();
                _state->pluginsRetainedSelectedPluginId.clear();
                _state->pluginsDetailsActive = false;
                _tree->SetSelectedItemId(itemId);
            }
            _model->Rebuild();
            _tree->NotifyDataChanged();
            if (_dialog && _state)
            {
                UpdatePageText(_dialog, *_state, PrefCategory::Plugins);
            }
        }

    private:
        HWND _dialog                         = nullptr;
        PreferencesDialogState* _state       = nullptr;
        PreferencesCategoryTreeModel* _model = nullptr;
        Tree* _tree                          = nullptr;
    } _categoryTreeDelegate;
};

constexpr wchar_t kPrefsPageHostClassName[]    = L"RedSalamanderPrefsPageHost";
constexpr wchar_t kPrefsDxShellHostClassName[] = L"RedSalamanderPrefsDxShellHost";
constexpr wchar_t kPreferencesWindowId[]       = L"PreferencesWindow";

struct PreferencesPageHostSurfaceControl final : RedSalamander::DxUi::Control
{
    explicit PreferencesPageHostSurfaceControl(const PreferencesDialogState* state) noexcept : _state(state)
    {
    }

    void Paint(WindowHost& host) const override
    {
        auto* dc = host.GetDeviceContext();
        if (! dc || ! _state)
        {
            return;
        }

        const D2D1_RECT_F bounds = GetBounds();
        if (auto* backgroundBrush = host.GetSolidBrush(ColorFromCOLORREF(_state->theme.windowBackground)))
        {
            dc->FillRectangle(bounds, backgroundBrush);
        }

        if (_state->theme.systemHighContrast || _state->pageSettingCards.empty())
        {
            return;
        }

        const COLORREF surfaceColor = UiMetrics::GetControlSurfaceColor(_state->theme);
        const COLORREF borderColor  = UiMetrics::BlendColor(surfaceColor, _state->theme.menu.text, _state->theme.dark ? 40 : 30, 255);
        auto* fillBrush             = host.GetSolidBrush(ColorFromCOLORREF(surfaceColor));
        auto* borderBrush           = host.GetSolidBrush(ColorFromCOLORREF(borderColor));
        if (! fillBrush || ! borderBrush)
        {
            return;
        }

        const float scrollDip     = host.PixelsToDip(static_cast<float>(_state->pageScrollY));
        constexpr float radiusDip = 6.0f;
        for (const RECT& baseCard : _state->pageSettingCards)
        {
            D2D1_RECT_F card = D2D1::RectF(host.PixelsToDip(static_cast<float>(baseCard.left)),
                                           host.PixelsToDip(static_cast<float>(baseCard.top)) - scrollDip,
                                           host.PixelsToDip(static_cast<float>(baseCard.right)),
                                           host.PixelsToDip(static_cast<float>(baseCard.bottom)) - scrollDip);
            if (card.right <= card.left || card.bottom <= card.top)
            {
                continue;
            }
            if (card.bottom <= bounds.top || card.top >= bounds.bottom)
            {
                continue;
            }

            const D2D1_ROUNDED_RECT rounded = D2D1::RoundedRect(card, radiusDip, radiusDip);
            dc->FillRoundedRectangle(&rounded, fillBrush);
            dc->DrawRoundedRectangle(&rounded, borderBrush, 1.0f);
        }
    }

protected:
    [[nodiscard]] RedSalamander::DxUi::Control* HitTest(D2D1_POINT_2F /*point*/) override
    {
        return nullptr;
    }

    [[nodiscard]] const RedSalamander::DxUi::Control* HitTest(D2D1_POINT_2F /*point*/) const override
    {
        return nullptr;
    }

private:
    const PreferencesDialogState* _state = nullptr;
};

struct PreferencesShellSurfaceControl final : RedSalamander::DxUi::Control
{
    void Paint(WindowHost& host) const override
    {
        auto* dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        if (auto* brush = host.GetSolidBrush(host.GetTheme().windowBackground))
        {
            dc->FillRectangle(host.GetClientBoundsDip(), brush);
        }
    }

protected:
    [[nodiscard]] RedSalamander::DxUi::Control* HitTest(D2D1_POINT_2F /*point*/) override
    {
        return nullptr;
    }

    [[nodiscard]] const RedSalamander::DxUi::Control* HitTest(D2D1_POINT_2F /*point*/) const override
    {
        return nullptr;
    }
};

[[nodiscard]] const wchar_t* GetPrefCategoryDebugName(const PrefCategory category) noexcept
{
    switch (category)
    {
        case PrefCategory::General: return L"General";
        case PrefCategory::Panes: return L"Panes";
        case PrefCategory::Viewers: return L"Viewers";
        case PrefCategory::Editors: return L"Editors";
        case PrefCategory::UserMenu: return L"UserMenu";
        case PrefCategory::Keyboard: return L"Keyboard";
        case PrefCategory::Mouse: return L"Mouse";
        case PrefCategory::Themes: return L"Themes";
        case PrefCategory::Plugins: return L"Plugins";
        case PrefCategory::FileOperations: return L"FileOperations";
        case PrefCategory::CompareDirectories: return L"CompareDirectories";
        case PrefCategory::HotPaths: return L"HotPaths";
        case PrefCategory::Advanced: return L"Advanced";
        default: return L"Unknown";
    }
}

[[nodiscard]] bool IsPreferencesDxDiagnosticsEnabled(HWND hwnd) noexcept
{
    return hwnd && GetPropW(hwnd, kPrefsDxDiagnosticsProp) != nullptr;
}

[[nodiscard]] PreferencesEmptyStateSpec GetCurrentPreferencesSharedEmptyState(const PreferencesDialogState& state) noexcept
{
    switch (state.currentCategory)
    {
        case PrefCategory::Editors:
            return PreferencesEmptyStateSpec{
                .iconGlyph         = FluentIcons::kOpenFile,
                .fallbackIconGlyph = FluentIcons::kFallbackChevronRight,
                .title             = LoadStringResource(nullptr, IDS_PREFS_EDITORS_EMPTY_TITLE),
                .body              = LoadStringResource(nullptr, IDS_PREFS_EDITORS_EMPTY_BODY),
                .caption           = LoadStringResource(nullptr, IDS_PREFS_EDITORS_EMPTY_CAPTION),
            };
        case PrefCategory::Mouse:
            return PreferencesEmptyStateSpec{
                .iconGlyph         = FluentIcons::kPreview,
                .fallbackIconGlyph = FluentIcons::kFallbackChevronRight,
                .title             = LoadStringResource(nullptr, IDS_PREFS_MOUSE_EMPTY_TITLE),
                .body              = LoadStringResource(nullptr, IDS_PREFS_MOUSE_EMPTY_BODY),
                .caption           = LoadStringResource(nullptr, IDS_PREFS_MOUSE_EMPTY_CAPTION),
            };
        case PrefCategory::General:
        case PrefCategory::Panes:
        case PrefCategory::Viewers:
        case PrefCategory::Keyboard:
        case PrefCategory::UserMenu:
        case PrefCategory::Themes:
        case PrefCategory::Plugins:
        case PrefCategory::Advanced:
        case PrefCategory::CompareDirectories:
        case PrefCategory::HotPaths:
        case PrefCategory::FileOperations: return {};
        default: return {};
    }
}

[[nodiscard]] std::wstring GetEmptyStateIconText(const WindowHost* host, const PreferencesEmptyStateSpec& spec) noexcept
{
    const bool useFluentIcon = host && host->HasFluentIconFont() && spec.iconGlyph != L'\0';
    const wchar_t glyph      = useFluentIcon ? spec.iconGlyph : spec.fallbackIconGlyph;
    return (glyph == L'\0') ? std::wstring{} : std::wstring(1, glyph);
}

void ApplySharedEmptyStatePalette(PreferencesDialogState& state, const PreferencesEmptyStateSpec* spec = nullptr) noexcept
{
    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    if (! hostState._pageHostEmptyStateIconControl || ! hostState._pageHostEmptyStateTitleControl || ! hostState._pageHostEmptyStateBodyControl ||
        ! hostState._pageHostEmptyStateCaptionControl)
    {
        return;
    }

    const ThemePalette palette = PrefsUi::MakeDxPalette(state.theme);
    D2D1_COLOR_F iconColor     = palette.subduedText;
    iconColor.a                = state.theme.dark ? 0.28f : 0.20f;

    if (spec)
    {
        switch (spec->tone)
        {
            case PreferencesEmptyStateSpec::Tone::Accent:
            case PreferencesEmptyStateSpec::Tone::Warning:
            case PreferencesEmptyStateSpec::Tone::Error:
            case PreferencesEmptyStateSpec::Tone::Info:
                iconColor   = palette.accent;
                iconColor.a = state.theme.dark ? 0.96f : 0.92f;
                break;
            case PreferencesEmptyStateSpec::Tone::Neutral:
            default: break;
        }
    }

    D2D1_COLOR_F bodyColor = palette.subduedText;
    bodyColor.a            = state.theme.dark ? 0.96f : 0.90f;

    D2D1_COLOR_F captionColor = palette.subduedText;
    captionColor.a            = state.theme.dark ? 0.72f : 0.76f;

    hostState._pageHostEmptyStateIconControl->SetTextColor(iconColor);
    hostState._pageHostEmptyStateTitleControl->SetTextColor(palette.text);
    hostState._pageHostEmptyStateBodyControl->SetTextColor(bodyColor);
    hostState._pageHostEmptyStateCaptionControl->SetTextColor(captionColor);
}

[[nodiscard]] HWND GetActivePreferencesPageRootWindow(const PreferencesDialogState& state) noexcept
{
    return (state.pageHostWindow && IsWindow(state.pageHostWindow) != FALSE) ? state.pageHostWindow : nullptr;
}

void InitializePreferencesPageControls(const PrefCategory category, PreferencesDialogState& state) noexcept
{
    UNREFERENCED_PARAMETER(category);
    UNREFERENCED_PARAMETER(state);
}

[[nodiscard]] bool EnsurePreferencesPageInitialized(HWND pageHostWindow, PreferencesDialogState& state, const PrefCategory category) noexcept
{
    if (! pageHostWindow)
    {
        return false;
    }

    auto& hostState               = static_cast<PreferencesDialogHost&>(state);
    const size_t catIndex         = PrefCategoryIndex(category);
    const bool needsFirstTimeInit = ! state.paneFirstCreateDone[catIndex];

    auto ensurePage = [&](auto& pane) noexcept
    {
        // Create a per-pane wrapper Panel so controls survive across switches.
        if (! state.paneWrapperPanels[catIndex] && hostState._pageHostContentRootControl)
        {
            auto* wrapper = hostState._pageHostContentRootControl->AddChild<Panel>();
            wrapper->SetBounds(hostState._pageHostContentRootControl->GetBounds());
            state.paneWrapperPanels[catIndex] = wrapper;
        }

        // Show only the active wrapper — hide all others.
        for (size_t i = 0; i < kPrefCategoryCount; ++i)
        {
            if (state.paneWrapperPanels[i])
            {
                state.paneWrapperPanels[i]->SetVisible(i == catIndex);
            }
        }

        if (state.pageHostDxHost)
        {
            state.pageHostDxHost->ResetInteractionState();
        }

        // Redirect shared content root to the active wrapper so pane EnsureDxHosts
        // stores it as _pageContentRoot, enabling per-pane control caching/reuse.
        state.pageHostDxContentRootControl = state.paneWrapperPanels[catIndex];

        if (needsFirstTimeInit)
        {
            pane.InitializePage(pageHostWindow, state);
            InitializePreferencesPageControls(category, state);
            InstallWheelRoutingHooks(pageHostWindow);
            state.paneFirstCreateDone[catIndex] = true;
        }
        return true;
    };

    switch (category)
    {
        case PrefCategory::General: return ensurePage(hostState._generalPane);
        case PrefCategory::Panes: return ensurePage(hostState._panesPane);
        case PrefCategory::Viewers: return ensurePage(hostState._viewersPane);
        case PrefCategory::Editors: return ensurePage(hostState._editorsPane);
        case PrefCategory::UserMenu: return ensurePage(hostState._userMenuPane);
        case PrefCategory::Keyboard: return ensurePage(hostState._keyboardPane);
        case PrefCategory::Mouse: return ensurePage(hostState._mousePane);
        case PrefCategory::Themes: return ensurePage(hostState._themesPane);
        case PrefCategory::Plugins: return ensurePage(hostState._pluginsPane);
        case PrefCategory::FileOperations: return ensurePage(hostState._fileOperationsPane);
        case PrefCategory::CompareDirectories: return ensurePage(hostState._compareDirectoriesPane);
        case PrefCategory::HotPaths: return ensurePage(hostState._hotPathsPane);
        case PrefCategory::Advanced: return ensurePage(hostState._advancedPane);
        default: return false;
    }
}

[[nodiscard]] bool EnsureActivePreferencesPageInitialized(HWND pageHostWindow, PreferencesDialogState& state) noexcept
{
    return EnsurePreferencesPageInitialized(pageHostWindow, state, state.currentCategory);
}

void ResetPreferencesSharedPageSurface(PreferencesDialogState& state) noexcept
{
    if (state.pageHostDxHost)
    {
        state.pageHostDxHost->ResetInteractionState();
    }
    if (state.pageHostDxContentRootControl)
    {
        state.pageHostDxContentRootControl->ClearChildren();
    }
    PrefsUi::HideSharedPageEmptyState(state);

    state.pageSettingCards.clear();
    state.pageHostDirectContentBottomPx = 0;
}

void DestroyInactivePreferencesPageState(PreferencesDialogState& state, const PrefCategory category) noexcept
{
    // Don't destroy panes — preserve DxUI state and pane data (keyboard rows,
    // theme items, etc.) for fast reactivation.  EnsureDxHosts will recreate
    // shared-root controls when the pane is next activated.
    // Only cancel keyboard capture when leaving the Keyboard pane.
    if (category == PrefCategory::Keyboard)
    {
        state.keyboardCaptureActive = false;
        state.keyboardCaptureCommandId.clear();
        state.keyboardCaptureBindingIndex.reset();
        state.keyboardCapturePendingVk.reset();
        state.keyboardCapturePendingModifiers = 0;
        state.keyboardCaptureConflictCommandId.clear();
        state.keyboardCaptureConflictBindingIndex.reset();
        state.keyboardCaptureConflictMultiple = false;
    }
}

[[nodiscard]] bool EnsurePrefsPageHostClassRegistered() noexcept
{
    const HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSEXW existing{};
    existing.cbSize = sizeof(existing);
    if (GetClassInfoExW(instance, kPrefsPageHostClassName, &existing) != 0)
    {
        return true;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_DBLCLKS;
    wc.lpfnWndProc   = PreferencesPageHostWindowProc;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = kPrefsPageHostClassName;

    const ATOM atom = RegisterClassExW(&wc);
    return atom != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

[[nodiscard]] bool EnsurePrefsDxShellHostClassRegistered() noexcept
{
    const HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSEXW existing{};
    existing.cbSize = sizeof(existing);
    if (GetClassInfoExW(instance, kPrefsDxShellHostClassName, &existing) != 0)
    {
        return true;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_DBLCLKS;
    wc.lpfnWndProc   = DefWindowProcW;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = kPrefsDxShellHostClassName;

    const ATOM atom = RegisterClassExW(&wc);
    return atom != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

[[nodiscard]] HWND FindWheelTargetFromPoint(HWND root, const PreferencesDialogState& state, POINT ptScreen) noexcept
{
    if (! root)
    {
        return nullptr;
    }

    HWND target = WindowFromPoint(ptScreen);
    if (! target || GetAncestor(target, GA_ROOT) != root)
    {
        return nullptr;
    }

    {
        const auto& hostState = static_cast<const PreferencesDialogHost&>(state);
        if (state.pageHostWindow && state.pageScrollMaxY > 0 && hostState._shellHostHwnd)
        {
            bool targetIsShellHost = false;
            for (HWND walk = target; walk && walk != root; walk = GetParent(walk))
            {
                if (walk == hostState._shellHostHwnd)
                {
                    targetIsShellHost = true;
                    break;
                }
            }

            if (targetIsShellHost)
            {
                RECT shellRect{};
                RECT pageHostRect{};
                if (GetWindowRect(hostState._shellHostHwnd, &shellRect) != FALSE && GetWindowRect(state.pageHostWindow, &pageHostRect) != FALSE &&
                    ptScreen.x >= pageHostRect.left && ptScreen.x < pageHostRect.right && ptScreen.y >= shellRect.top && ptScreen.y < pageHostRect.top)
                {
                    return state.pageHostWindow;
                }
            }
        }
    }

    while (target && target != root)
    {
        if (target == state.categoryTreeWindow || target == state.pageHostWindow)
        {
            return target;
        }

        const LONG_PTR style = GetWindowLongPtrW(target, GWL_STYLE);
        if ((style & WS_VSCROLL) != 0)
        {
            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_RANGE | SIF_PAGE;
            if (GetScrollInfo(target, SB_VERT, &si))
            {
                const int range = std::max(0, (si.nMax - si.nMin) + 1);
                if (range <= static_cast<int>(si.nPage))
                {
                    target = GetParent(target);
                    continue;
                }
            }

            std::array<wchar_t, 16> className{};
            const int len = GetClassNameW(target, className.data(), static_cast<int>(className.size()));
            if (len > 0 && _wcsicmp(className.data(), L"ComboBox") == 0)
            {
                if (SendMessageW(target, CB_GETDROPPEDSTATE, 0, 0) == 0)
                {
                    target = GetParent(target);
                    continue;
                }
            }
            return target;
        }
        target = GetParent(target);
    }

    return nullptr;
}

[[nodiscard]] bool HandlePageHostMouseWheel(HWND host, PreferencesDialogState& state, WPARAM wp) noexcept
{
#ifdef ENABLE_TESTS
    state.debugLastWheelFallbackCalled   = true;
    state.debugLastWheelFallbackHandled  = false;
    state.debugLastWheelDelta            = static_cast<int>(GET_WHEEL_DELTA_WPARAM(wp));
    state.debugLastWheelBeforeY          = state.pageScrollY;
    state.debugLastWheelBeforeMaxY       = state.pageScrollMaxY;
    state.debugLastWheelAfterY           = state.pageScrollY;
    state.debugLastWheelAfterMaxY        = state.pageScrollMaxY;
#endif

    if (! host || state.pageScrollMaxY <= 0)
    {
        return false;
    }

    const int delta = GET_WHEEL_DELTA_WPARAM(wp);
    if (delta == 0)
    {
#ifdef ENABLE_TESTS
        state.debugLastWheelFallbackHandled = true;
#endif
        return true;
    }

    state.pageWheelDeltaRemainder += delta;
    const int steps = state.pageWheelDeltaRemainder / WHEEL_DELTA;
    if (steps == 0)
    {
#ifdef ENABLE_TESTS
        state.debugLastWheelFallbackHandled = true;
#endif
        return true;
    }
    state.pageWheelDeltaRemainder -= steps * WHEEL_DELTA;

    UINT linesPerNotch = 3;
    SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &linesPerNotch, 0);
    if (linesPerNotch == 0)
    {
#ifdef ENABLE_TESTS
        state.debugLastWheelFallbackHandled = true;
#endif
        return true;
    }

    const UINT dpi     = GetDpiForWindow(host);
    const int lineStep = std::max(1, UiMetrics::ScaleDip(dpi, 24));

    int scrollDelta = 0;
    if (linesPerNotch == WHEEL_PAGESCROLL)
    {
        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_PAGE;
        GetScrollInfo(host, SB_VERT, &si);
        scrollDelta = steps * static_cast<int>(si.nPage);
    }
    else
    {
        scrollDelta = steps * lineStep * static_cast<int>(linesPerNotch);
    }

    const int newPos = state.pageScrollY - scrollDelta;
    PrefsPageHost::ScrollTo(host, state, newPos);
#ifdef ENABLE_TESTS
    state.debugLastWheelFallbackHandled = true;
    state.debugLastWheelAfterY          = state.pageScrollY;
    state.debugLastWheelAfterMaxY       = state.pageScrollMaxY;
#endif
    return true;
}

[[nodiscard]] int ColorLuma(COLORREF c) noexcept
{
    // Approximate ITU-R BT.601 luma in 0..255.
    const int r = static_cast<int>(GetRValue(c));
    const int g = static_cast<int>(GetGValue(c));
    const int b = static_cast<int>(GetBValue(c));
    return (299 * r + 587 * g + 114 * b) / 1000;
}

[[nodiscard]] COLORREF GetDisabledTextColor(const PreferencesDialogState& state, COLORREF background) noexcept
{
    COLORREF candidate = state.theme.menu.disabledText;
    if (state.theme.highContrast)
    {
        return candidate;
    }

    const COLORREF normal   = state.theme.menu.text;
    const int minBgDiff     = 80;
    const int minNormalDiff = 36;

    const auto isReadableAndDim = [&](COLORREF color) noexcept
    {
        const int bgDiff     = std::abs(ColorLuma(color) - ColorLuma(background));
        const int normalDiff = std::abs(ColorLuma(color) - ColorLuma(normal));
        return bgDiff >= minBgDiff && normalDiff >= minNormalDiff;
    };

    auto blended = UiMetrics::BlendColor(background, normal, state.theme.dark ? 140 : 90, 255);
    if (std::abs(ColorLuma(blended) - ColorLuma(background)) < minBgDiff)
    {
        blended = UiMetrics::BlendColor(background, normal, state.theme.dark ? 170 : 120, 255);
    }

    if (isReadableAndDim(candidate))
    {
        const int candNormalDiff  = std::abs(ColorLuma(candidate) - ColorLuma(normal));
        const int blendNormalDiff = std::abs(ColorLuma(blended) - ColorLuma(normal));
        if (candNormalDiff >= blendNormalDiff)
        {
            return candidate;
        }
    }

    return blended;
}

LRESULT CALLBACK PreferencesWheelRouteWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_MOUSEWHEEL:
        {
            if (g_preferencesWheelMessageForwarded)
            {
                return CallStoredWndProc(hwnd, kPrefsWheelRouteOriginalWndProcProp, msg, wp, lp);
            }

            const HWND dlg = GetAncestor(hwnd, GA_ROOT);
            if (! dlg)
            {
                return 0;
            }

            auto* state = GetState(dlg);
            if (! state)
            {
                return CallStoredWndProc(hwnd, kPrefsWheelRouteOriginalWndProcProp, msg, wp, lp);
            }

            POINT ptScreen{};
            ptScreen.x = GET_X_LPARAM(lp);
            ptScreen.y = GET_Y_LPARAM(lp);

            HWND target = FindWheelTargetFromPoint(dlg, *state, ptScreen);
#ifdef ENABLE_TESTS
            const HWND windowFromPoint                           = WindowFromPoint(ptScreen);
            state->debugLastWheelRouteSeen                      = true;
            state->debugLastWheelRouteForwarded                 = target && target != hwnd;
            state->debugLastWheelRouteTargetWasPageHost         = target && target == state->pageHostWindow;
            state->debugLastWheelRouteTargetWasCategoryTree     = target && target == state->categoryTreeWindow;
            state->debugLastWheelRouteTargetHadVerticalScroll   = target && ((GetWindowLongPtrW(target, GWL_STYLE) & WS_VSCROLL) != 0);
            state->debugLastWheelWindowFromPointWasPageHost     = windowFromPoint && windowFromPoint == state->pageHostWindow;
            state->debugLastWheelWindowFromPointWasCategoryTree = windowFromPoint && windowFromPoint == state->categoryTreeWindow;
            state->debugLastWheelWndProcSeen                    = false;
            state->debugLastWheelDxHandled                      = false;
            state->debugLastWheelFallbackCalled                 = false;
            state->debugLastWheelFallbackHandled                = false;
            state->debugLastWheelDelta                          = static_cast<int>(GET_WHEEL_DELTA_WPARAM(wp));
            state->debugLastWheelBeforeY                        = state->pageScrollY;
            state->debugLastWheelBeforeMaxY                     = state->pageScrollMaxY;
            state->debugLastWheelAfterY                         = state->pageScrollY;
            state->debugLastWheelAfterMaxY                      = state->pageScrollMaxY;
            POINT pageHostClientPoint                           = ptScreen;
            if (state->pageHostWindow)
            {
                ScreenToClient(state->pageHostWindow, &pageHostClientPoint);
            }
            state->debugLastWheelClientX = pageHostClientPoint.x;
            state->debugLastWheelClientY = pageHostClientPoint.y;
#endif
            if (! target)
            {
                // Don't scroll the dialog when the user is wheeling outside it.
                return 0;
            }

            if (target == hwnd)
            {
                break;
            }

            const bool previousForwardedState  = g_preferencesWheelMessageForwarded;
            g_preferencesWheelMessageForwarded = true;
            SendMessageW(target, msg, wp, lp);
            g_preferencesWheelMessageForwarded = previousForwardedState;
            return 0;
        }
        case WM_NCDESTROY:
        {
            RemovePropW(hwnd, kPrefsWheelRouteOriginalWndProcProp);
            break;
        }
    }

    return CallStoredWndProc(hwnd, kPrefsWheelRouteOriginalWndProcProp, msg, wp, lp);
}

void InstallWheelRoutingHooks(HWND dlg) noexcept
{
    if (! dlg)
    {
        return;
    }

    const auto installHook = [](HWND hwnd, LPARAM) noexcept -> BOOL
    {
        static_cast<void>(InstallWndProcHook(hwnd, kPrefsWheelRouteOriginalWndProcProp, PreferencesWheelRouteWndProc));
        return TRUE;
    };

    static_cast<void>(installHook(dlg, 0));
    EnumChildWindows(dlg, installHook, 0);
}

void PaintPageHostBackgroundAndCards(HDC hdc, HWND host, const PreferencesDialogState& state) noexcept
{
    if (! hdc || ! host)
    {
        return;
    }

    RECT rc{};
    GetClientRect(host, &rc);

    D2DHdcPaint::Session paint;
    if (! paint.Begin(hdc, rc))
    {
        HBRUSH brush = state.backgroundBrush ? state.backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(NULL_BRUSH));
        FillRect(hdc, &rc, brush);
        return;
    }

    paint.FillRectangle(rc, state.theme.windowBackground);

    if (state.theme.systemHighContrast || state.pageSettingCards.empty())
    {
        return;
    }

    const UINT dpi         = GetDpiForWindow(host);
    const float radius     = static_cast<float>(UiMetrics::ScaleDip(dpi, 6));
    const COLORREF surface = UiMetrics::GetControlSurfaceColor(state.theme);
    const COLORREF border  = UiMetrics::BlendColor(surface, state.theme.menu.text, state.theme.dark ? 40 : 30, 255);

    for (const RECT& baseCard : state.pageSettingCards)
    {
        RECT card = baseCard;
        OffsetRect(&card, 0, -state.pageScrollY);
        if (card.right <= card.left || card.bottom <= card.top)
        {
            continue;
        }
        if (card.bottom <= rc.top || card.top >= rc.bottom)
        {
            continue;
        }
        paint.FillRoundedRectangle(card, radius, surface, border);
    }
}

wil::unique_hwnd g_preferencesDialog;

bool RequestPreferencesDialogClose(HWND dlg) noexcept
{
    if (! dlg)
    {
        return false;
    }

    if (PostMessageW(dlg, kPrefsDeferredCloseMessage, 0, 0) == 0)
    {
        static_cast<void>(Debug::ErrorWithLastError(L"Preferences: PostMessageW failed for deferred close"));
        return false;
    }

    return true;
}

void RestorePreviewAppliedPreferencesOnCancel(PreferencesDialogState& state) noexcept
{
    if (! state.previewApplied || ! state.settings)
    {
        return;
    }

    Common::Settings::Settings restored = *state.settings;
    restored.theme                      = state.baselineSettings.theme;
    *state.settings                    = std::move(restored);

    if (state.owner)
    {
        PostMessageW(state.owner, WndMsg::kSettingsApplied, 0, 0);
    }
}

[[nodiscard]] const CategoryInfo* FindCategoryInfo(PrefCategory id) noexcept
{
    for (const auto& c : kCategories)
    {
        if (c.id == id)
        {
            return &c;
        }
    }
    return nullptr;
}

[[nodiscard]] PreferencesDialogState* GetState(HWND dlg) noexcept
{
    return reinterpret_cast<PreferencesDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
}

void SetState(HWND dlg, PreferencesDialogState* state) noexcept
{
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
}

void RestoreWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    if (const auto originalWndProc = GetStoredWndProc(hwnd, originalWndProcProp))
    {
        RemovePropW(hwnd, originalWndProcProp);
        static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(originalWndProc)));
    }
}

[[maybe_unused]] [[nodiscard]] size_t CountVisibleChildWindows(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    struct VisibleChildCounter
    {
        size_t count = 0u;
    } counter{};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& counterRef = *reinterpret_cast<VisibleChildCounter*>(lParam);
        if (PrefsUi::IsActuallyVisibleChildWindow(child))
        {
            ++counterRef.count;
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&counter)));

    return counter.count;
}

[[maybe_unused]] [[nodiscard]] size_t CountVisibleChildWindowsByClass(HWND hwnd, const wchar_t* className) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE || ! className || *className == L'\0')
    {
        return 0u;
    }

    struct VisibleClassCounter
    {
        const wchar_t* className = nullptr;
        size_t count             = 0u;
    } counter{className, 0u};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& counterRef = *reinterpret_cast<VisibleClassCounter*>(lParam);
        if (! PrefsUi::IsActuallyVisibleChildWindow(child))
        {
            return TRUE;
        }

        std::array<wchar_t, 64> actualClassName{};
        if (GetClassNameW(child, actualClassName.data(), static_cast<int>(actualClassName.size())) > 0 &&
            _wcsicmp(actualClassName.data(), counterRef.className) == 0)
        {
            ++counterRef.count;
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&counter)));

    return counter.count;
}

LRESULT CALLBACK PreferencesDxCategoryHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* state = reinterpret_cast<PreferencesDialogState*>(GetPropW(hwnd, kPrefsDxCategoryHostStateProp));
    if (! state)
    {
        return CallStoredWndProc(hwnd, kPrefsDxCategoryHostOriginalWndProcProp, msg, wp, lp);
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);

    if (msg == WM_GETDLGCODE)
    {
        return CallStoredWndProc(hwnd, kPrefsDxCategoryHostOriginalWndProcProp, msg, wp, lp) | DLGC_WANTARROWS | DLGC_WANTCHARS | DLGC_WANTALLKEYS;
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kPrefsDxCategoryHostOriginalWndProcProp);
        RemovePropW(hwnd, kPrefsDxCategoryHostStateProp);
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kPrefsDxCategoryHostOriginalWndProcProp, PreferencesDxCategoryHostWndProc);
        hostState._categoryTreeHost.ReleaseMouseCapture();
        hostState._categoryTreeControl = nullptr;
        return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    bool handled           = false;
    const LRESULT dxResult = hostState._categoryTreeHost.HandleMessage(hwnd, msg, wp, lp, handled);
    if (handled)
    {
        return dxResult;
    }

    return CallStoredWndProc(hwnd, kPrefsDxCategoryHostOriginalWndProcProp, msg, wp, lp);
}

LRESULT CALLBACK PreferencesDxShellHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* host = reinterpret_cast<WindowHost*>(GetPropW(hwnd, kPrefsDxShellHostProp));
    if (! host)
    {
        return CallStoredWndProc(hwnd, kPrefsDxShellHostOriginalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_GETDLGCODE)
    {
        return CallStoredWndProc(hwnd, kPrefsDxShellHostOriginalWndProcProp, msg, wp, lp) | DLGC_WANTARROWS | DLGC_WANTCHARS | DLGC_WANTALLKEYS;
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kPrefsDxShellHostOriginalWndProcProp);
        RemovePropW(hwnd, kPrefsDxShellHostProp);
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kPrefsDxShellHostOriginalWndProcProp, PreferencesDxShellHostWndProc);
        host->ReleaseMouseCapture();
        return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    bool handled           = false;
    const LRESULT dxResult = host->HandleMessage(hwnd, msg, wp, lp, handled);
    if (handled)
    {
        return dxResult;
    }

    return CallStoredWndProc(hwnd, kPrefsDxShellHostOriginalWndProcProp, msg, wp, lp);
}

void UpdateDxShellText(PreferencesDialogHost& hostState, std::wstring_view title, std::wstring_view description) noexcept
{
    if (hostState._pageTitleControl)
    {
        hostState._pageTitleControl->SetText(std::wstring(title));
    }

    if (hostState._pageDescriptionControl)
    {
        hostState._pageDescriptionControl->SetText(std::wstring(description));
    }

    hostState._shellHost.Invalidate();
    hostState._pageHostHost.Invalidate();
}

void UpdateDxShellButtons(HWND dlg, PreferencesDialogHost& hostState, const PreferencesDialogState& state) noexcept
{
    UNREFERENCED_PARAMETER(dlg);

    if (hostState._okButtonControl)
    {
        hostState._okButtonControl->SetText(LoadStringResource(nullptr, IDS_BTN_OK));
        hostState._okButtonControl->SetMnemonic(L'O');
        hostState._okButtonControl->SetPrimary(true);
    }
    if (hostState._cancelButtonControl)
    {
        hostState._cancelButtonControl->SetText(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        hostState._cancelButtonControl->SetMnemonic(L'C');
    }
    if (hostState._applyButtonControl)
    {
        hostState._applyButtonControl->SetText(LoadStringResource(nullptr, IDS_BTN_APPLY));
        hostState._applyButtonControl->SetMnemonic(L'A');
        hostState._applyButtonControl->SetEnabled(state.dirty);
    }
    if (hostState._resetAllButtonControl)
    {
        hostState._resetAllButtonControl->SetText(LoadStringResource(nullptr, IDS_PREFS_BUTTON_RESET_ALL));
        hostState._resetAllButtonControl->SetMnemonic(L'R');
    }

    hostState._shellHost.Invalidate();
}

void ApplyDxShellTheme(PreferencesDialogHost& hostState, const AppTheme& theme) noexcept
{
    const ThemePalette palette = PrefsUi::MakeDxPalette(theme);
    hostState._shellHost.SetTheme(palette);
}

[[nodiscard]] HWND CreatePreferencesShellHostWindow(HWND dlg, bool tabStop) noexcept
{
    if (! EnsurePrefsDxShellHostClassRegistered())
    {
        Debug::ErrorWithLastError(L"Preferences: failed to register DX shell host class.");
        return nullptr;
    }

    const DWORD style = WS_CHILD | SS_NOTIFY | (tabStop ? WS_TABSTOP : 0u);
    return CreateWindowExW(0, kPrefsDxShellHostClassName, L"", style, 0, 0, 10, 10, dlg, nullptr, GetModuleHandleW(nullptr), nullptr);
}

[[nodiscard]] bool AttachPreferencesShellHost(HWND dlg, PreferencesDialogHost& hostState) noexcept
{
    if (! hostState._shellHostHwnd || ! hostState._shellHost.Attach(hostState._shellHostHwnd))
    {
        return false;
    }

    SetPropW(hostState._shellHostHwnd, kPrefsDxShellHostProp, &hostState._shellHost);
    if (! InstallWndProcHook(hostState._shellHostHwnd, kPrefsDxShellHostOriginalWndProcProp, PreferencesDxShellHostWndProc))
    {
        Debug::ErrorWithLastError(L"Preferences: failed to install DX shell host window proc hook.");
        RemovePropW(hostState._shellHostHwnd, kPrefsDxShellHostProp);
        hostState._shellHost.Detach();
        return false;
    }

    auto root = std::make_unique<RedSalamander::DxUi::Panel>();
    root->AddChild<PreferencesShellSurfaceControl>();

    const auto addButton = [&](const UINT commandId, const bool primary) noexcept
    {
        auto* button = root->AddChild<Button>();
        button->SetPrimary(primary);
        button->SetOnClick([dlg, commandId]
        {
            if (dlg && IsWindow(dlg) != FALSE)
            {
                PostMessageW(dlg, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
            }
        });
        return button;
    };

    hostState._resetAllButtonControl  = addButton(IDC_PREFS_RESET_ALL, false);
    hostState._okButtonControl        = addButton(IDOK, true);
    hostState._cancelButtonControl    = addButton(IDCANCEL, false);
    hostState._applyButtonControl     = addButton(IDC_PREFS_APPLY, false);

    hostState._shellHost.SetRoot(std::move(root));
    hostState._shellHost.SetDefaultButton(hostState._okButtonControl);
    hostState._shellHost.SetCancelButton(hostState._cancelButtonControl);
    hostState._shellHost.SetOnTabBoundary([dlg, shellHostHwnd = hostState._shellHostHwnd](const bool reverse) noexcept
    {
        if (! dlg || IsWindow(dlg) == FALSE || ! shellHostHwnd || IsWindow(shellHostHwnd) == FALSE)
        {
            return false;
        }

        if (HWND target = GetNextDlgTabItem(dlg, shellHostHwnd, reverse ? TRUE : FALSE); target && target != shellHostHwnd)
        {
            SetFocus(target);
            return true;
        }

        return false;
    });
    return true;
}

void CreatePreferencesShellHosts(HWND dlg, PreferencesDialogHost& hostState) noexcept
{
    hostState._shellHostHwnd = CreatePreferencesShellHostWindow(dlg, true);

    if (! AttachPreferencesShellHost(dlg, hostState))
    {
        Debug::Error(L"Preferences: failed to initialize DX shell hosts.");
        return;
    }

    UpdateDxShellButtons(dlg, hostState, hostState);
    ApplyDxShellTheme(hostState, hostState.theme);

    const auto hideWindow = [](HWND hwnd) noexcept
    {
        if (hwnd)
        {
            ShowWindow(hwnd, SW_HIDE);
        }
    };

    hideWindow(GetDlgItem(dlg, IDOK));
    hideWindow(GetDlgItem(dlg, IDCANCEL));
    hideWindow(GetDlgItem(dlg, IDC_PREFS_APPLY));
}

void OnPageHostScrollPanelOffsetChanged(PreferencesDialogState& state, const float scrollOffsetDip) noexcept
{
    if (state.pageHostSyncingScrollPanel || ! state.pageHostWindow || IsWindow(state.pageHostWindow) == FALSE)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(state.pageHostWindow);
    if (dpi == 0u)
    {
        return;
    }

    const int newScrollY = std::clamp(static_cast<int>(std::lround((scrollOffsetDip * static_cast<float>(dpi)) / 96.0f)), 0, state.pageScrollMaxY);
    if (newScrollY == state.pageScrollY)
    {
        return;
    }

    const int oldScrollY = state.pageScrollY;
    state.pageScrollY    = newScrollY;
    state.pageHostScrollApplyPending = false;

    PrefsPageHost::ApplyScrollDelta(state.pageHostWindow, oldScrollY - state.pageScrollY, false);
    RedrawWindow(state.pageHostWindow, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN);
}

void AttachPreferencesPageHostDxSurface(PreferencesDialogHost& hostState) noexcept
{
    hostState.pageHostUsesDxUi                  = false;
    hostState._pageHostRootControl              = nullptr;
    hostState.pageHostDxScrollPanelControl      = nullptr;
    hostState._pageHostContentRootControl       = nullptr;
    hostState._pageHostSurfaceControl           = nullptr;
    hostState._pageTitleControl                 = nullptr;
    hostState._pageDescriptionControl           = nullptr;
    hostState._pageHostEmptyStatePanel          = nullptr;
    hostState._pageHostEmptyStateIconControl    = nullptr;
    hostState._pageHostEmptyStateTitleControl   = nullptr;
    hostState._pageHostEmptyStateBodyControl    = nullptr;
    hostState._pageHostEmptyStateCaptionControl = nullptr;
    hostState.pageHostDxHost                    = nullptr;
    hostState.pageHostDxRootControl             = nullptr;
    hostState.pageHostDxScrollPanelControl      = nullptr;
    hostState.pageHostDxContentRootControl      = nullptr;
    hostState.pageHostDxNoteControl             = nullptr;
    if (! hostState.pageHostWindow)
    {
        return;
    }

    if (hostState._pageHostHost.GetHwnd() != hostState.pageHostWindow)
    {
        hostState._pageHostHost.Detach();
        if (! hostState._pageHostHost.Attach(hostState.pageHostWindow))
        {
            Debug::Error(L"Preferences: failed to attach DxUi host for page host surface.");
            return;
        }
    }

    hostState._pageHostHost.SetTheme(PrefsUi::MakeDxPalette(hostState.theme));
    auto root               = std::make_unique<Panel>();
    auto* rawRoot           = root.get();
    auto* surface           = rawRoot->AddChild<PreferencesPageHostSurfaceControl>(&hostState);
    auto* contentRoot       = rawRoot->AddChild<ScrollPanel>();
    auto* title             = contentRoot->AddChild<Label>();
    auto* description       = contentRoot->AddChild<Label>();
    auto* emptyStatePanel   = rawRoot->AddChild<Panel>();
    auto* emptyStateIcon    = rawRoot->AddChild<Label>();
    auto* emptyStateTitle   = rawRoot->AddChild<Label>();
    auto* emptyStateBody    = rawRoot->AddChild<Label>();
    auto* emptyStateCaption = rawRoot->AddChild<Label>();

    contentRoot->SetInternalScrollbarEnabled(true);
    contentRoot->SetOnScrollChanged([state = static_cast<PreferencesDialogState*>(&hostState)](const float scrollOffsetDip) noexcept
    {
        if (state)
        {
            OnPageHostScrollPanelOffsetChanged(*state, scrollOffsetDip);
        }
    });
    title->SetMultiline(false);
    title->SetFontRole(FontRole::Title);
    description->SetMultiline(true);
    description->SetFontRole(FontRole::Body);
    emptyStatePanel->SetVisible(false);

    emptyStateIcon->SetMultiline(false);
    emptyStateIcon->SetFontRole(FontRole::HeroIcon);
    emptyStateIcon->SetAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
    emptyStateIcon->SetVisible(false);

    emptyStateTitle->SetMultiline(true);
    emptyStateTitle->SetFontRole(FontRole::Title);
    emptyStateTitle->SetAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
    emptyStateTitle->SetVisible(false);

    emptyStateBody->SetMultiline(true);
    emptyStateBody->SetFontRole(FontRole::Body);
    emptyStateBody->SetAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
    emptyStateBody->SetVisible(false);

    emptyStateCaption->SetMultiline(true);
    emptyStateCaption->SetFontRole(FontRole::Small);
    emptyStateCaption->SetAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
    emptyStateCaption->SetVisible(false);

    hostState._pageHostRootControl              = rawRoot;
    hostState._pageHostContentRootControl       = contentRoot;
    hostState._pageHostSurfaceControl           = surface;
    hostState._pageTitleControl                 = title;
    hostState._pageDescriptionControl           = description;
    hostState._pageHostEmptyStatePanel          = emptyStatePanel;
    hostState._pageHostEmptyStateIconControl    = emptyStateIcon;
    hostState._pageHostEmptyStateTitleControl   = emptyStateTitle;
    hostState._pageHostEmptyStateBodyControl    = emptyStateBody;
    hostState._pageHostEmptyStateCaptionControl = emptyStateCaption;
    hostState.pageHostDxHost                    = &hostState._pageHostHost;
    hostState.pageHostDxRootControl             = rawRoot;
    hostState.pageHostDxScrollPanelControl      = contentRoot;
    hostState.pageHostDxContentRootControl      = contentRoot;
    hostState.pageHostDxNoteControl             = emptyStatePanel;
    hostState._pageHostHost.SetRoot(std::move(root));
    hostState.pageHostUsesDxUi = true;
    ApplySharedEmptyStatePalette(hostState);
}

void ShowDialogAlert(HWND dlg, HostAlertSeverity severity, const std::wstring& title, const std::wstring& message) noexcept
{
    if (! dlg || message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = severity;
    request.targetWindow = dlg;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(HostShowAlert(request));
}

[[nodiscard]] Common::Settings::MainMenuState GetMainMenu(const Common::Settings::Settings& settings) noexcept
{
    if (settings.mainMenu.has_value())
    {
        return settings.mainMenu.value();
    }
    return {};
}

[[nodiscard]] const Common::Settings::StartupSettings& GetStartupSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::StartupSettings kDefaults{};
    if (settings.startup.has_value())
    {
        return settings.startup.value();
    }
    return kDefaults;
}

[[nodiscard]] const Common::Settings::UiSettings& GetUiSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::UiSettings kDefaults{};
    if (settings.ui.has_value())
    {
        return settings.ui.value();
    }
    return kDefaults;
}

[[nodiscard]] const Common::Settings::MonitorSettings& GetMonitorSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::MonitorSettings kDefaults{};
    if (settings.monitor.has_value())
    {
        return settings.monitor.value();
    }
    return kDefaults;
}

[[nodiscard]] const Common::Settings::CacheSettings& GetCacheSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::CacheSettings kDefaults{};
    if (settings.cache.has_value())
    {
        return settings.cache.value();
    }
    return kDefaults;
}

[[nodiscard]] const Common::Settings::FileOperationsSettings& GetFileOperationsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::FileOperationsSettings kDefaults{};
    if (settings.fileOperations.has_value())
    {
        return settings.fileOperations.value();
    }
    return kDefaults;
}

[[nodiscard]] bool AreEquivalentShortcutBindings(const std::vector<Common::Settings::ShortcutBinding>& a,
                                                 const std::vector<Common::Settings::ShortcutBinding>& b) noexcept
{
    using Key = std::tuple<uint32_t, uint32_t, std::wstring>;

    auto normalize = [](const std::vector<Common::Settings::ShortcutBinding>& bindings)
    {
        std::vector<Key> keys;
        keys.reserve(bindings.size());
        for (const auto& binding : bindings)
        {
            if (binding.commandId.empty())
            {
                continue;
            }
            keys.emplace_back(binding.vk, binding.modifiers & 0x7u, binding.commandId);
        }

        std::sort(keys.begin(), keys.end());
        keys.erase(std::unique(keys.begin(), keys.end()), keys.end());
        return keys;
    };

    return normalize(a) == normalize(b);
}

[[nodiscard]] bool AreEquivalentShortcuts(const std::optional<Common::Settings::ShortcutsSettings>& a,
                                          const std::optional<Common::Settings::ShortcutsSettings>& b) noexcept
{
    if (! a.has_value() && ! b.has_value())
    {
        return true;
    }

    if (a.has_value() && b.has_value())
    {
        return AreEquivalentShortcutBindings(a.value().functionBar, b.value().functionBar) &&
               AreEquivalentShortcutBindings(a.value().folderView, b.value().folderView);
    }

    const Common::Settings::ShortcutsSettings defaults = ShortcutDefaults::CreateDefaultShortcuts();
    const Common::Settings::ShortcutsSettings& aValue  = a.has_value() ? a.value() : defaults;
    const Common::Settings::ShortcutsSettings& bValue  = b.has_value() ? b.value() : defaults;

    return AreEquivalentShortcutBindings(aValue.functionBar, bValue.functionBar) && AreEquivalentShortcutBindings(aValue.folderView, bValue.folderView);
}

[[nodiscard]] bool AreEquivalentThemeDefinition(const Common::Settings::ThemeDefinition& a, const Common::Settings::ThemeDefinition& b) noexcept
{
    if (a.id != b.id || a.name != b.name || a.baseThemeId != b.baseThemeId)
    {
        return false;
    }

    if (a.colors.size() != b.colors.size())
    {
        return false;
    }

    for (const auto& [key, value] : a.colors)
    {
        const auto it = b.colors.find(key);
        if (it == b.colors.end() || it->second != value)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool AreEquivalentThemeSettings(const Common::Settings::ThemeSettings& a, const Common::Settings::ThemeSettings& b) noexcept
{
    if (a.currentThemeId != b.currentThemeId)
    {
        return false;
    }

    if (a.themes.size() != b.themes.size())
    {
        return false;
    }

    for (const auto& theme : a.themes)
    {
        const auto it =
            std::find_if(b.themes.begin(), b.themes.end(), [&](const Common::Settings::ThemeDefinition& other) noexcept { return other.id == theme.id; });
        if (it == b.themes.end() || ! AreEquivalentThemeDefinition(theme, *it))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool AreEquivalentJsonValue(const Common::Settings::JsonValue& a, const Common::Settings::JsonValue& b) noexcept
{
    return a.value == b.value;
}

[[nodiscard]] bool AreEquivalentConnectionProfile(const Common::Settings::ConnectionProfile& a, const Common::Settings::ConnectionProfile& b) noexcept
{
    return a.id == b.id && a.name == b.name && a.pluginId == b.pluginId && a.host == b.host && a.port == b.port && a.initialPath == b.initialPath &&
           a.userName == b.userName && a.authMode == b.authMode && a.savePassword == b.savePassword && a.requireWindowsHello == b.requireWindowsHello &&
           AreEquivalentJsonValue(a.extra, b.extra);
}

[[nodiscard]] bool AreEquivalentConnectionsSettings(const std::optional<Common::Settings::ConnectionsSettings>& a,
                                                    const std::optional<Common::Settings::ConnectionsSettings>& b) noexcept
{
    if (! a.has_value() && ! b.has_value())
    {
        return true;
    }

    const Common::Settings::ConnectionsSettings defaults{};
    const Common::Settings::ConnectionsSettings& left  = a.has_value() ? a.value() : defaults;
    const Common::Settings::ConnectionsSettings& right = b.has_value() ? b.value() : defaults;

    return left.bypassWindowsHello == right.bypassWindowsHello && left.allowInsecureTlsInAutomation == right.allowInsecureTlsInAutomation &&
           left.windowsHelloReauthTimeoutMinute == right.windowsHelloReauthTimeoutMinute && left.items.size() == right.items.size() &&
           std::ranges::equal(left.items, right.items, [](const auto& lhs, const auto& rhs) noexcept { return AreEquivalentConnectionProfile(lhs, rhs); });
}

[[nodiscard]] bool AreEquivalentPluginsDisabledIds(const std::vector<std::wstring>& a, const std::vector<std::wstring>& b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        const std::wstring_view id = a[i];
        if (id.empty())
        {
            return false;
        }

        for (size_t j = 0; j < i; ++j)
        {
            if (a[j] == a[i])
            {
                return false;
            }
        }

        if (std::find(b.begin(), b.end(), id) == b.end())
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool AreEquivalentPluginsSettings(const Common::Settings::PluginsSettings& a, const Common::Settings::PluginsSettings& b) noexcept
{
    if (a.currentFileSystemPluginId != b.currentFileSystemPluginId)
    {
        return false;
    }
    if (a.customPluginPaths != b.customPluginPaths)
    {
        return false;
    }
    if (! AreEquivalentPluginsDisabledIds(a.disabledPluginIds, b.disabledPluginIds))
    {
        return false;
    }

    if (a.configurationByPluginId.size() != b.configurationByPluginId.size())
    {
        return false;
    }

    for (const auto& [id, value] : a.configurationByPluginId)
    {
        const auto it = b.configurationByPluginId.find(id);
        if (it == b.configurationByPluginId.end())
        {
            return false;
        }
        if (! AreEquivalentJsonValue(value, it->second))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool AreEquivalentCompareDirectoriesSettings(const Common::Settings::CompareDirectoriesSettings& a,
                                                           const Common::Settings::CompareDirectoriesSettings& b) noexcept
{
    return a.compareSize == b.compareSize && a.compareDateTime == b.compareDateTime && a.compareAttributes == b.compareAttributes &&
           a.compareContent == b.compareContent && a.compareSubdirectories == b.compareSubdirectories &&
           a.compareSubdirectoryAttributes == b.compareSubdirectoryAttributes && a.selectSubdirsOnlyInOnePane == b.selectSubdirsOnlyInOnePane &&
           a.ignoreFiles == b.ignoreFiles && a.ignoreFilesPatterns == b.ignoreFilesPatterns && a.ignoreDirectories == b.ignoreDirectories &&
           a.ignoreDirectoriesPatterns == b.ignoreDirectoriesPatterns && a.keepIdenticalItems == b.keepIdenticalItems &&
           a.showIdenticalItems == b.showIdenticalItems && a.contentCompareWorkerCount == b.contentCompareWorkerCount;
}

[[nodiscard]] bool IsEmptyHotPathSlot(const Common::Settings::HotPathSlot& slot) noexcept
{
    return slot.path.empty() && slot.label.empty() && ! slot.showInMenu;
}

[[nodiscard]] bool AreEquivalentHotPathSlots(const std::optional<Common::Settings::HotPathSlot>& a,
                                             const std::optional<Common::Settings::HotPathSlot>& b) noexcept
{
    const Common::Settings::HotPathSlot* aSlot = (a.has_value() && ! IsEmptyHotPathSlot(a.value())) ? &a.value() : nullptr;
    const Common::Settings::HotPathSlot* bSlot = (b.has_value() && ! IsEmptyHotPathSlot(b.value())) ? &b.value() : nullptr;
    if (! aSlot && ! bSlot)
    {
        return true;
    }
    if (! aSlot || ! bSlot)
    {
        return false;
    }

    return aSlot->path == bSlot->path && aSlot->label == bSlot->label && aSlot->showInMenu == bSlot->showInMenu;
}

[[nodiscard]] bool AreEquivalentHotPathsSettings(const Common::Settings::HotPathsSettings& a, const Common::Settings::HotPathsSettings& b) noexcept
{
    if (a.openPrefsOnAssign != b.openPrefsOnAssign)
    {
        return false;
    }

    for (size_t i = 0; i < a.slots.size(); ++i)
    {
        if (! AreEquivalentHotPathSlots(a.slots[i], b.slots[i]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool IsDirty(const PreferencesDialogState& state) noexcept
{
    const Common::Settings::MainMenuState baseline = GetMainMenu(state.baselineSettings);
    const Common::Settings::MainMenuState working  = GetMainMenu(state.workingSettings);
    if (baseline.menuBarVisible != working.menuBarVisible)
    {
        return true;
    }
    if (baseline.functionBarVisible != working.functionBarVisible)
    {
        return true;
    }
    {
        const auto& baselineStartup = GetStartupSettingsOrDefault(state.baselineSettings);
        const auto& workingStartup  = GetStartupSettingsOrDefault(state.workingSettings);
        if (baselineStartup.showSplash != workingStartup.showSplash)
        {
            return true;
        }
    }
    {
        const auto& baselineUi = GetUiSettingsOrDefault(state.baselineSettings);
        const auto& workingUi  = GetUiSettingsOrDefault(state.workingSettings);
        if (baselineUi != workingUi)
        {
            return true;
        }
    }
    if (! AreEquivalentShortcuts(state.baselineSettings.shortcuts, state.workingSettings.shortcuts))
    {
        return true;
    }
    if (! AreEquivalentThemeSettings(state.baselineSettings.theme, state.workingSettings.theme))
    {
        return true;
    }
    if (! PrefsFolders::AreEquivalentFolderPreferences(state.baselineSettings, state.workingSettings))
    {
        return true;
    }
    {
        const auto& baselineMonitor = GetMonitorSettingsOrDefault(state.baselineSettings);
        const auto& workingMonitor  = GetMonitorSettingsOrDefault(state.workingSettings);
        if (baselineMonitor.menu.toolbarVisible != workingMonitor.menu.toolbarVisible ||
            baselineMonitor.menu.lineNumbersVisible != workingMonitor.menu.lineNumbersVisible ||
            baselineMonitor.menu.alwaysOnTop != workingMonitor.menu.alwaysOnTop || baselineMonitor.menu.showIds != workingMonitor.menu.showIds ||
            baselineMonitor.menu.autoScroll != workingMonitor.menu.autoScroll || baselineMonitor.filter.mask != workingMonitor.filter.mask ||
            baselineMonitor.filter.preset != workingMonitor.filter.preset)
        {
            return true;
        }
    }
    {
        const auto& baselineCache = GetCacheSettingsOrDefault(state.baselineSettings);
        const auto& workingCache  = GetCacheSettingsOrDefault(state.workingSettings);
        if (baselineCache.directoryInfo.maxBytes != workingCache.directoryInfo.maxBytes ||
            baselineCache.directoryInfo.maxWatchers != workingCache.directoryInfo.maxWatchers ||
            baselineCache.directoryInfo.mruWatched != workingCache.directoryInfo.mruWatched)
        {
            return true;
        }
    }
    {
        const auto& baselineFileOperations = GetFileOperationsSettingsOrDefault(state.baselineSettings);
        const auto& workingFileOperations  = GetFileOperationsSettingsOrDefault(state.workingSettings);
        if (baselineFileOperations.autoDismissSuccess != workingFileOperations.autoDismissSuccess ||
            baselineFileOperations.preCalcEnabled != workingFileOperations.preCalcEnabled ||
            baselineFileOperations.preCalcMaxWorkers != workingFileOperations.preCalcMaxWorkers ||
            baselineFileOperations.defaultBandwidthLimitBytesPerSecond != workingFileOperations.defaultBandwidthLimitBytesPerSecond ||
            baselineFileOperations.crossFsBridgeBufferSizeKB != workingFileOperations.crossFsBridgeBufferSizeKB ||
            baselineFileOperations.maxDiagnosticsLogFiles != workingFileOperations.maxDiagnosticsLogFiles ||
            baselineFileOperations.diagnosticsInfoEnabled != workingFileOperations.diagnosticsInfoEnabled ||
            baselineFileOperations.diagnosticsDebugEnabled != workingFileOperations.diagnosticsDebugEnabled ||
            baselineFileOperations.maxIssueReportFiles != workingFileOperations.maxIssueReportFiles ||
            baselineFileOperations.maxDiagnosticsInMemory != workingFileOperations.maxDiagnosticsInMemory ||
            baselineFileOperations.maxDiagnosticsPerFlush != workingFileOperations.maxDiagnosticsPerFlush ||
            baselineFileOperations.diagnosticsFlushIntervalMs != workingFileOperations.diagnosticsFlushIntervalMs ||
            baselineFileOperations.diagnosticsCleanupIntervalMs != workingFileOperations.diagnosticsCleanupIntervalMs)
        {
            return true;
        }
    }
    if (! AreEquivalentConnectionsSettings(state.baselineSettings.connections, state.workingSettings.connections))
    {
        return true;
    }
    if (state.baselineSettings.fileActions != state.workingSettings.fileActions)
    {
        return true;
    }
    {
        const auto& baselineCompare = PrefsCompareDirectories::GetCompareDirectoriesSettingsOrDefault(state.baselineSettings);
        const auto& workingCompare  = PrefsCompareDirectories::GetCompareDirectoriesSettingsOrDefault(state.workingSettings);
        if (! AreEquivalentCompareDirectoriesSettings(baselineCompare, workingCompare))
        {
            return true;
        }
    }
    {
        const auto& baselineHotPaths = PrefsHotPaths::GetHotPathsSettingsOrDefault(state.baselineSettings);
        const auto& workingHotPaths  = PrefsHotPaths::GetHotPathsSettingsOrDefault(state.workingSettings);
        if (! AreEquivalentHotPathsSettings(baselineHotPaths, workingHotPaths))
        {
            return true;
        }
    }
    if (! AreEquivalentPluginsSettings(state.baselineSettings.plugins, state.workingSettings.plugins))
    {
        return true;
    }
    return false;
}

void UpdateApplyButton(HWND dlg, const PreferencesDialogState& state) noexcept
{
    HWND apply = dlg ? GetDlgItem(dlg, IDC_PREFS_APPLY) : nullptr;
    if (apply)
    {
        EnableWindow(apply, state.dirty ? TRUE : FALSE);
    }

    if (dlg)
    {
        auto* mutableState = GetState(dlg);
        if (mutableState)
        {
            auto& hostState = static_cast<PreferencesDialogHost&>(*mutableState);
            if (hostState._applyButtonControl)
            {
                hostState._applyButtonControl->SetEnabled(state.dirty);
                hostState._shellHost.Invalidate();
            }
        }
    }
}

} // namespace

void SetDirty(HWND dlg, PreferencesDialogState& state) noexcept
{
    state.dirty = IsDirty(state);
    UpdateApplyButton(dlg, state);
}

void PrefsUi::HideSharedPageEmptyState(PreferencesDialogState& state) noexcept
{
    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    if (hostState._pageHostEmptyStatePanel)
    {
        hostState._pageHostEmptyStatePanel->SetVisible(false);
        hostState._pageHostEmptyStatePanel->SetBounds(D2D1::RectF());
    }

    if (hostState._pageHostEmptyStateIconControl)
    {
        hostState._pageHostEmptyStateIconControl->SetText(L"");
        hostState._pageHostEmptyStateIconControl->SetBounds(D2D1::RectF());
    }
    if (hostState._pageHostEmptyStateTitleControl)
    {
        hostState._pageHostEmptyStateTitleControl->SetText(L"");
        hostState._pageHostEmptyStateTitleControl->SetBounds(D2D1::RectF());
    }
    if (hostState._pageHostEmptyStateBodyControl)
    {
        hostState._pageHostEmptyStateBodyControl->SetText(L"");
        hostState._pageHostEmptyStateBodyControl->SetBounds(D2D1::RectF());
    }
    if (hostState._pageHostEmptyStateCaptionControl)
    {
        hostState._pageHostEmptyStateCaptionControl->SetText(L"");
        hostState._pageHostEmptyStateCaptionControl->SetBounds(D2D1::RectF());
    }
}

int PrefsUi::ShowSharedPageEmptyState(HWND host,
                                      PreferencesDialogState& state,
                                      const PreferencesEmptyStateSpec& spec,
                                      int x,
                                      int y,
                                      int width,
                                      const PreferencesTypographyContext& typography) noexcept
{
    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    if (! host || width <= 0 || ! hostState._pageHostEmptyStatePanel || ! hostState._pageHostEmptyStateIconControl ||
        ! hostState._pageHostEmptyStateTitleControl || ! hostState._pageHostEmptyStateBodyControl || ! hostState._pageHostEmptyStateCaptionControl)
    {
        return 0;
    }

    const UINT dpi              = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const int cardPaddingX      = UiMetrics::ScaleDip(dpi, 20);
    const int cardPaddingTop    = UiMetrics::ScaleDip(dpi, 18);
    const int cardPaddingBottom = UiMetrics::ScaleDip(dpi, 18);
    const int iconHeight        = UiMetrics::ScaleDip(dpi, 72);
    const int gapIconTitle      = UiMetrics::ScaleDip(dpi, 8);
    const int gapTitleBody      = UiMetrics::ScaleDip(dpi, 6);
    const int gapBodyCaption    = UiMetrics::ScaleDip(dpi, 8);
    const int textWidth         = std::max(0, width - 2 * cardPaddingX);

    const int titleHeight   = spec.title.empty() ? 0 : PrefsUi::MeasureWrappedTextHeightPx(typography, typography.title, textWidth, spec.title);
    const int bodyHeight    = spec.body.empty() ? 0 : PrefsUi::MeasureWrappedTextHeightPx(typography, typography.body, textWidth, spec.body);
    const int captionHeight = spec.caption.empty() ? 0 : PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, spec.caption);

    int contentHeight = iconHeight;
    if (titleHeight > 0)
    {
        contentHeight += gapIconTitle + titleHeight;
    }
    if (bodyHeight > 0)
    {
        contentHeight += gapTitleBody + bodyHeight;
    }
    if (captionHeight > 0)
    {
        contentHeight += gapBodyCaption + captionHeight;
    }

    const int cardHeight = cardPaddingTop + contentHeight + cardPaddingBottom;
    const auto pxToDip   = [dpi](const int value) noexcept { return (static_cast<float>(value) * 96.0f) / static_cast<float>((std::max)(1u, dpi)); };

    ApplySharedEmptyStatePalette(state, &spec);

    const std::wstring iconText = GetEmptyStateIconText(state.pageHostDxHost, spec);
    hostState._pageHostEmptyStatePanel->SetVisible(true);
    hostState._pageHostEmptyStatePanel->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + cardHeight)));

    int currentY = y + cardPaddingTop;

    hostState._pageHostEmptyStateIconControl->SetText(iconText);
    hostState._pageHostEmptyStateIconControl->SetVisible(! iconText.empty());
    hostState._pageHostEmptyStateIconControl->SetBounds(
        D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(currentY), pxToDip(x + width - cardPaddingX), pxToDip(currentY + iconHeight)));
    currentY += iconHeight;

    if (titleHeight > 0)
    {
        currentY += gapIconTitle;
        hostState._pageHostEmptyStateTitleControl->SetText(spec.title);
        hostState._pageHostEmptyStateTitleControl->SetVisible(true);
        hostState._pageHostEmptyStateTitleControl->SetBounds(
            D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(currentY), pxToDip(x + width - cardPaddingX), pxToDip(currentY + titleHeight)));
        currentY += titleHeight;
    }
    else
    {
        hostState._pageHostEmptyStateTitleControl->SetVisible(false);
        hostState._pageHostEmptyStateTitleControl->SetBounds(D2D1::RectF());
        hostState._pageHostEmptyStateTitleControl->SetText(L"");
    }

    if (bodyHeight > 0)
    {
        currentY += gapTitleBody;
        hostState._pageHostEmptyStateBodyControl->SetText(spec.body);
        hostState._pageHostEmptyStateBodyControl->SetVisible(true);
        hostState._pageHostEmptyStateBodyControl->SetBounds(
            D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(currentY), pxToDip(x + width - cardPaddingX), pxToDip(currentY + bodyHeight)));
        currentY += bodyHeight;
    }
    else
    {
        hostState._pageHostEmptyStateBodyControl->SetVisible(false);
        hostState._pageHostEmptyStateBodyControl->SetBounds(D2D1::RectF());
        hostState._pageHostEmptyStateBodyControl->SetText(L"");
    }

    if (captionHeight > 0)
    {
        currentY += gapBodyCaption;
        hostState._pageHostEmptyStateCaptionControl->SetText(spec.caption);
        hostState._pageHostEmptyStateCaptionControl->SetVisible(true);
        hostState._pageHostEmptyStateCaptionControl->SetBounds(
            D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(currentY), pxToDip(x + width - cardPaddingX), pxToDip(currentY + captionHeight)));
    }
    else
    {
        hostState._pageHostEmptyStateCaptionControl->SetVisible(false);
        hostState._pageHostEmptyStateCaptionControl->SetBounds(D2D1::RectF());
        hostState._pageHostEmptyStateCaptionControl->SetText(L"");
    }

    return cardHeight;
}

namespace
{
[[nodiscard]] HRESULT SaveSettingsFromDialog(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (state.appId.empty())
    {
        return E_INVALIDARG;
    }

    if (state.owner && IsWindow(state.owner))
    {
        SendMessageW(state.owner, WndMsg::kPreferencesRequestSettingsSnapshot, 0, 0);
    }

    Common::Settings::Settings merged;
    merged = state.settings ? *state.settings : state.workingSettings;

    const Common::Settings::MainMenuState baselineMenu = GetMainMenu(state.baselineSettings);
    const Common::Settings::MainMenuState workingMenu  = GetMainMenu(state.workingSettings);
    // Always preserve mainMenu if it exists in working settings or if values differ from baseline
    // This ensures defaults are explicitly saved rather than relying on implicit defaults
    if (state.workingSettings.mainMenu.has_value() || baselineMenu.menuBarVisible != workingMenu.menuBarVisible ||
        baselineMenu.functionBarVisible != workingMenu.functionBarVisible)
    {
        merged.mainMenu = workingMenu;
    }

    {
        const auto& baselineStartup = GetStartupSettingsOrDefault(state.baselineSettings);
        const auto& workingStartup  = GetStartupSettingsOrDefault(state.workingSettings);
        if (baselineStartup.showSplash != workingStartup.showSplash)
        {
            merged.startup = workingStartup;
        }
    }
    {
        const auto& baselineUi = GetUiSettingsOrDefault(state.baselineSettings);
        const auto& workingUi  = GetUiSettingsOrDefault(state.workingSettings);
        if (baselineUi != workingUi)
        {
            merged.ui = workingUi;
        }
    }

    if (! AreEquivalentShortcuts(state.baselineSettings.shortcuts, state.workingSettings.shortcuts))
    {
        merged.shortcuts = state.workingSettings.shortcuts;
    }
    if (! AreEquivalentThemeSettings(state.baselineSettings.theme, state.workingSettings.theme))
    {
        merged.theme = state.workingSettings.theme;
    }
    if (! PrefsFolders::AreEquivalentFolderPreferences(state.baselineSettings, state.workingSettings))
    {
        const PrefsFolders::FolderPanePreferences left  = PrefsFolders::GetFolderPanePreferences(state.workingSettings, PrefsFolders::kLeftPaneSlot);
        const PrefsFolders::FolderPanePreferences right = PrefsFolders::GetFolderPanePreferences(state.workingSettings, PrefsFolders::kRightPaneSlot);
        const uint32_t historyMax                       = PrefsFolders::GetFolderHistoryMax(state.workingSettings);
        const bool showHiddenFiles                      = PrefsFolders::GetFolderShowHiddenFiles(state.workingSettings);
        const bool showSystemFiles                      = PrefsFolders::GetFolderShowSystemFiles(state.workingSettings);

        auto* folders = PrefsFolders::EnsureWorkingFoldersSettings(merged);
        if (folders)
        {
            folders->historyMax      = historyMax;
            folders->showHiddenFiles = showHiddenFiles;
            folders->showSystemFiles = showSystemFiles;

            if (auto* pane = PrefsFolders::EnsureWorkingFolderPane(merged, PrefsFolders::kLeftPaneSlot))
            {
                pane->view.display          = left.display;
                pane->view.sortBy           = left.sortBy;
                pane->view.sortDirection    = left.sortDirection;
                pane->view.statusBarVisible = left.statusBarVisible;
            }
            if (auto* pane = PrefsFolders::EnsureWorkingFolderPane(merged, PrefsFolders::kRightPaneSlot))
            {
                pane->view.display          = right.display;
                pane->view.sortBy           = right.sortBy;
                pane->view.sortDirection    = right.sortDirection;
                pane->view.statusBarVisible = right.statusBarVisible;
            }
        }
    }

    {
        const auto& baselineMonitor = GetMonitorSettingsOrDefault(state.baselineSettings);
        const auto& workingMonitor  = GetMonitorSettingsOrDefault(state.workingSettings);
        if (baselineMonitor.menu.toolbarVisible != workingMonitor.menu.toolbarVisible ||
            baselineMonitor.menu.lineNumbersVisible != workingMonitor.menu.lineNumbersVisible ||
            baselineMonitor.menu.alwaysOnTop != workingMonitor.menu.alwaysOnTop || baselineMonitor.menu.showIds != workingMonitor.menu.showIds ||
            baselineMonitor.menu.autoScroll != workingMonitor.menu.autoScroll || baselineMonitor.filter.mask != workingMonitor.filter.mask ||
            baselineMonitor.filter.preset != workingMonitor.filter.preset)
        {
            merged.monitor = workingMonitor;
        }
    }
    {
        const auto& baselineCache = GetCacheSettingsOrDefault(state.baselineSettings);
        const auto& workingCache  = GetCacheSettingsOrDefault(state.workingSettings);
        if (baselineCache.directoryInfo.maxBytes != workingCache.directoryInfo.maxBytes ||
            baselineCache.directoryInfo.maxWatchers != workingCache.directoryInfo.maxWatchers ||
            baselineCache.directoryInfo.mruWatched != workingCache.directoryInfo.mruWatched)
        {
            merged.cache = state.workingSettings.cache;
        }
    }
    {
        const auto& baselineFileOperations = GetFileOperationsSettingsOrDefault(state.baselineSettings);
        const auto& workingFileOperations  = GetFileOperationsSettingsOrDefault(state.workingSettings);
        if (baselineFileOperations.autoDismissSuccess != workingFileOperations.autoDismissSuccess ||
            baselineFileOperations.preCalcEnabled != workingFileOperations.preCalcEnabled ||
            baselineFileOperations.preCalcMaxWorkers != workingFileOperations.preCalcMaxWorkers ||
            baselineFileOperations.defaultBandwidthLimitBytesPerSecond != workingFileOperations.defaultBandwidthLimitBytesPerSecond ||
            baselineFileOperations.crossFsBridgeBufferSizeKB != workingFileOperations.crossFsBridgeBufferSizeKB ||
            baselineFileOperations.maxDiagnosticsLogFiles != workingFileOperations.maxDiagnosticsLogFiles ||
            baselineFileOperations.diagnosticsInfoEnabled != workingFileOperations.diagnosticsInfoEnabled ||
            baselineFileOperations.diagnosticsDebugEnabled != workingFileOperations.diagnosticsDebugEnabled ||
            baselineFileOperations.maxIssueReportFiles != workingFileOperations.maxIssueReportFiles ||
            baselineFileOperations.maxDiagnosticsInMemory != workingFileOperations.maxDiagnosticsInMemory ||
            baselineFileOperations.maxDiagnosticsPerFlush != workingFileOperations.maxDiagnosticsPerFlush ||
            baselineFileOperations.diagnosticsFlushIntervalMs != workingFileOperations.diagnosticsFlushIntervalMs ||
            baselineFileOperations.diagnosticsCleanupIntervalMs != workingFileOperations.diagnosticsCleanupIntervalMs)
        {
            merged.fileOperations = state.workingSettings.fileOperations;
        }
    }
    if (! AreEquivalentConnectionsSettings(state.baselineSettings.connections, state.workingSettings.connections))
    {
        merged.connections = state.workingSettings.connections;
    }
    if (state.baselineSettings.fileActions != state.workingSettings.fileActions)
    {
        merged.fileActions = state.workingSettings.fileActions;
    }
    {
        const auto& baselineCompare = PrefsCompareDirectories::GetCompareDirectoriesSettingsOrDefault(state.baselineSettings);
        const auto& workingCompare  = PrefsCompareDirectories::GetCompareDirectoriesSettingsOrDefault(state.workingSettings);
        if (! AreEquivalentCompareDirectoriesSettings(baselineCompare, workingCompare))
        {
            merged.compareDirectories = state.workingSettings.compareDirectories;
        }
    }
    {
        const auto& baselineHotPaths = PrefsHotPaths::GetHotPathsSettingsOrDefault(state.baselineSettings);
        const auto& workingHotPaths  = PrefsHotPaths::GetHotPathsSettingsOrDefault(state.workingSettings);
        if (! AreEquivalentHotPathsSettings(baselineHotPaths, workingHotPaths))
        {
            merged.hotPaths = state.workingSettings.hotPaths;
        }
    }
    if (! AreEquivalentPluginsSettings(state.baselineSettings.plugins, state.workingSettings.plugins))
    {
        merged.plugins = state.workingSettings.plugins;
    }

    Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(merged);

    const HRESULT hr = SettingsHotReload::SaveSettingsAndSchema(state.appId, merged);
    if (FAILED(hr))
    {
        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(state.appId);
        const std::wstring title                 = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        const std::wstring message = FormatStringResource(nullptr, IDS_FMT_SETTINGS_SAVE_FAILED, settingsPath.wstring(), static_cast<unsigned long>(hr));
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, message);
        return hr;
    }

    state.workingSettings = std::move(settingsToSave);

    return S_OK;
}

void RefreshPreferencesDialogThemeImpl(HWND dlg, PreferencesDialogState& state) noexcept;
void PopulateCategoryTree(HWND dlg, PreferencesDialogState& state) noexcept;
void UpdatePageText(HWND dlg, PreferencesDialogState& state, PrefCategory category) noexcept;
void LayoutPreferencesPageHost(HWND host, PreferencesDialogState& state) noexcept;

void ReloadPreferencesDialogFromDisk(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! state.settings)
    {
        return;
    }

    state.previewApplied          = false;
    state.staleFromExternalReload = false;
    state.baselineSettings        = *state.settings;
    state.workingSettings         = *state.settings;

    PopulateCategoryTree(dlg, state);
    RefreshPreferencesDialogThemeImpl(dlg, state);
    UpdatePageText(dlg, state, state.currentCategory);
    SetDirty(dlg, state);
}

void ResetAllPreferencesToDefaults(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

    const auto& hostState      = static_cast<PreferencesDialogHost&>(state);
    const std::wstring title   = LoadStringResource(nullptr, IDS_PREFS_BUTTON_RESET_ALL);
    const std::wstring message = LoadStringResource(nullptr, IDS_PREFS_RESET_ALL_CONFIRM);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences reset-all: begin");
#endif

    HostPromptRequest request{};
    request.version       = 1u;
    request.sizeBytes     = sizeof(request);
    request.scope         = HOST_ALERT_SCOPE_WINDOW;
    request.severity      = HOST_ALERT_WARNING;
    request.buttons       = HOST_PROMPT_BUTTONS_YES_NO;
    request.targetWindow  = (state.pageHostWindow && IsWindow(state.pageHostWindow) != FALSE)
                                ? state.pageHostWindow
                                : ((hostState._shellHostHwnd && IsWindow(hostState._shellHostHwnd) != FALSE) ? hostState._shellHostHwnd : dlg);
    request.title         = title.c_str();
    request.message       = message.c_str();
    request.defaultResult = HOST_PROMPT_RESULT_NO;
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences reset-all: prompt request prepared");
#endif

    Debug::Info(L"Preferences: reset-all prompt begin target={:#x} pageHost={:#x} shellHost={:#x}",
                reinterpret_cast<uintptr_t>(request.targetWindow),
                reinterpret_cast<uintptr_t>(state.pageHostWindow),
                reinterpret_cast<uintptr_t>(hostState._shellHostHwnd));

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(request, nullptr, &promptResult);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(
        L"Preferences reset-all: prompt completed hr=0x{:08X} result={}", static_cast<unsigned int>(hrPrompt), static_cast<unsigned int>(promptResult)));
#endif
    Debug::Info(L"Preferences: reset-all prompt end hr=0x{:08X} result={}", static_cast<unsigned int>(hrPrompt), static_cast<unsigned int>(promptResult));
    if (FAILED(hrPrompt))
    {
        Debug::Warning(L"Preferences: reset-all confirmation prompt failed; hr=0x{:08X}.", static_cast<unsigned int>(hrPrompt));
        return;
    }

    if (promptResult != HOST_PROMPT_RESULT_YES)
    {
        return;
    }

    // Reset all working settings to defaults.
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences reset-all: resetting working settings");
#endif
    state.workingSettings = Common::Settings::Settings{};

    // Shortcuts require explicit default initialization via the factory.
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences reset-all: restoring default shortcuts");
#endif
    state.workingSettings.shortcuts.emplace(ShortcutDefaults::CreateDefaultShortcuts());

    // Clear cached pane-local UI state that may reference old values.
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences reset-all: clearing cached pane-local state");
#endif
    state.viewersSelectedExtensionText.clear();
    state.pluginsSelectedCustomPathText.clear();
    state.keyboardCaptureActive = false;
    state.keyboardCaptureCommandId.clear();
    state.keyboardCaptureBindingIndex.reset();
    state.keyboardCapturePendingVk.reset();
    state.keyboardCapturePendingModifiers = 0;
    state.keyboardCaptureConflictCommandId.clear();
    state.keyboardCaptureConflictBindingIndex.reset();
    state.keyboardCaptureConflictMultiple = false;

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences reset-all: SetDirty");
#endif
    SetDirty(dlg, state);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences reset-all: RefreshPreferencesDialogTheme");
#endif
    RefreshPreferencesDialogThemeImpl(dlg, state);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences reset-all: RebindActivePreferencesPage");
#endif
    RebindActivePreferencesPage(dlg, state);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences reset-all: complete");
#endif
}

void RebindActivePreferencesPage(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences rebind: begin");
#endif
    ResetPreferencesSharedPageSurface(state);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences rebind: shared page surface reset");
#endif
    if (state.pageHostWindow)
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"Preferences rebind: ensuring active page initialized");
#endif
        static_cast<void>(EnsureActivePreferencesPageInitialized(state.pageHostWindow, state));
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"Preferences rebind: refreshing active page");
#endif
        RefreshActivePreferencesPage(state.pageHostWindow, state);
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"Preferences rebind: laying out page host");
#endif
        LayoutPreferencesPageHost(state.pageHostWindow, state);
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences rebind: updating page text");
#endif
    UpdatePageText(dlg, state, state.currentCategory);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences rebind: complete");
#endif
}

[[nodiscard]] bool ResolvePreferencesStaleSaveConflict(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! state.staleFromExternalReload)
    {
        return true;
    }

    SettingsHotReload::StaleSaveChoice choice = SettingsHotReload::StaleSaveChoice::Cancel;
    const HRESULT promptHr                    = SettingsHotReload::PromptStaleSaveConflict(dlg, LoadStringResource(nullptr, IDS_PREFS_CAPTION), choice);
    if (FAILED(promptHr))
    {
        Debug::Warning(L"Preferences: failed to prompt for stale save conflict (hr=0x{:08X})", static_cast<unsigned long>(promptHr));
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::ReloadFromDisk)
    {
        ReloadPreferencesDialogFromDisk(dlg, state);
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::Cancel)
    {
        return false;
    }

    state.staleFromExternalReload = false;
    return true;
}

INT_PTR OnSettingsReloadedFromDisk(HWND dlg, PreferencesDialogState& state) noexcept
{
    state.previewApplied = false;
    RefreshPreferencesDialogThemeImpl(dlg, state);

    if (state.dirty)
    {
        SettingsHotReload::ExternalReloadChoice choice = SettingsHotReload::ExternalReloadChoice::KeepEditing;
        const HRESULT promptHr = SettingsHotReload::PromptExternalReloadConflict(dlg, LoadStringResource(nullptr, IDS_PREFS_CAPTION), choice);
        if (FAILED(promptHr))
        {
            Debug::Warning(L"Preferences: failed to prompt for external reload conflict (hr=0x{:08X})", static_cast<unsigned long>(promptHr));
            return TRUE;
        }

        if (choice == SettingsHotReload::ExternalReloadChoice::KeepEditing)
        {
            state.staleFromExternalReload = true;
            return TRUE;
        }
    }

    ReloadPreferencesDialogFromDisk(dlg, state);
    return TRUE;
}

void CommitAndApply(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! dlg || ! state.settings)
    {
        return;
    }

    if (! ResolvePreferencesStaleSaveConflict(dlg, state))
    {
        return;
    }

    const HRESULT saveHr = SaveSettingsFromDialog(dlg, state);
    if (FAILED(saveHr))
    {
        return;
    }

    const bool pluginsChanged = ! AreEquivalentPluginsSettings(state.baselineSettings.plugins, state.workingSettings.plugins);

    *state.settings               = state.workingSettings;
    state.baselineSettings        = state.workingSettings;
    state.previewApplied          = false;
    state.staleFromExternalReload = false;

    state.appliedOnce = true;
    SetDirty(dlg, state);

    if (state.owner && IsWindow(state.owner) != FALSE)
    {
        PostMessageW(state.owner, WndMsg::kSettingsApplied, 0, 0);
    }
    if (pluginsChanged && state.owner && IsWindow(state.owner) != FALSE)
    {
        PostMessageW(state.owner, WndMsg::kPluginsChanged, 0, 0);
    }

    RefreshPreferencesDialogThemeImpl(dlg, state);
}

[[nodiscard]] HostPromptResult PromptSaveDirtyPreferencesBeforeClose(HWND dlg, PreferencesDialogState& state) noexcept
{
    const auto& hostState = static_cast<PreferencesDialogHost&>(state);

    const std::wstring title   = LoadStringResource(nullptr, IDS_PREFS_CAPTION);
    const std::wstring message = LoadStringResource(nullptr, IDS_PREFS_CONFIRM_SAVE_CHANGES);

    HostPromptRequest request{};
    request.version       = 1u;
    request.sizeBytes     = sizeof(request);
    request.scope         = HOST_ALERT_SCOPE_WINDOW;
    request.severity      = HOST_ALERT_WARNING;
    request.buttons       = HOST_PROMPT_BUTTONS_YES_NO_CANCEL;
    request.targetWindow  = (state.pageHostWindow && IsWindow(state.pageHostWindow) != FALSE)
                                ? state.pageHostWindow
                                : ((hostState._shellHostHwnd && IsWindow(hostState._shellHostHwnd) != FALSE) ? hostState._shellHostHwnd : dlg);
    request.title         = title.c_str();
    request.message       = message.c_str();
    request.defaultResult = HOST_PROMPT_RESULT_CANCEL;

#ifdef ENABLE_TESTS
    const bool forceDiscardForAutoAccept =
        HostGetAutoAcceptPrompts() && HostGetTestPromptResultOverride() == HOST_PROMPT_RESULT_NONE;
    if (forceDiscardForAutoAccept)
    {
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_NO);
    }
    const auto restorePromptOverride = wil::scope_exit([forceDiscardForAutoAccept]() noexcept
    {
        if (forceDiscardForAutoAccept)
        {
            HostClearTestPromptResultOverride();
        }
    });
#endif

    HostPromptResult result = HOST_PROMPT_RESULT_CANCEL;
    const HRESULT hrPrompt  = HostShowPrompt(request, nullptr, &result);
    if (FAILED(hrPrompt))
    {
        Debug::Warning(L"Preferences: failed to prompt for dirty close confirmation (hr=0x{:08X})", static_cast<unsigned int>(hrPrompt));
        return HOST_PROMPT_RESULT_CANCEL;
    }

    return result;
}

[[nodiscard]] bool RequestPreferencesDialogCancelClose(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! state.dirty)
    {
        RestorePreviewAppliedPreferencesOnCancel(state);
        return RequestPreferencesDialogClose(dlg);
    }

    switch (PromptSaveDirtyPreferencesBeforeClose(dlg, state))
    {
        case HOST_PROMPT_RESULT_YES:
            CommitAndApply(dlg, state);
            if (state.dirty)
            {
                return true;
            }
            return RequestPreferencesDialogClose(dlg);
        case HOST_PROMPT_RESULT_NO:
            RestorePreviewAppliedPreferencesOnCancel(state);
            return RequestPreferencesDialogClose(dlg);
        case HOST_PROMPT_RESULT_NONE:
        case HOST_PROMPT_RESULT_OK:
        case HOST_PROMPT_RESULT_CANCEL:
        default: return true;
    }
}

void LayoutPreferencesDialog(HWND dlg, PreferencesDialogState& state) noexcept;
void LayoutPreferencesPageHost(HWND host, PreferencesDialogState& state) noexcept;

void RefreshPreferencesDialogThemeImpl(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! dlg || ! state.settings)
    {
        return;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    ApplyWindowChromeTheme(dlg, state.theme, WindowBackdropTarget::Tool, GetActiveWindow() == dlg);

    // No ScopedWindowRedrawBlock — see comment in UpdatePageText.
    if (state.categoryTreeUsesDxUi)
    {
        hostState._categoryTreeHost.SetTheme(PrefsUi::MakeDxPalette(state.theme));
    }
    ApplyDxShellTheme(hostState, state.theme);
    UpdateDxShellButtons(dlg, hostState, state);

    // Update page title/description for current category.
    {
        const CategoryInfo* info = FindCategoryInfo(state.currentCategory);
        std::wstring title       = info ? LoadStringResource(nullptr, info->labelId) : std::wstring{};
        std::wstring description = info ? LoadStringResource(nullptr, info->descriptionId) : std::wstring{};
        UpdateDxShellText(hostState, title, description);
    }

    if (state.pageHostUsesDxUi)
    {
        hostState._pageHostHost.SetTheme(PrefsUi::MakeDxPalette(state.theme));
    }

    // Set pageHostIgnoreSize BEFORE creating pane controls so that any WM_SIZE
    // messages generated during EnsureCreated / CreateControls do not trigger
    // LayoutPreferencesPageHost with partially-initialized state.
    state.pageHostIgnoreSize = true;
    const auto clearIgnore   = wil::scope_exit([&]() noexcept { state.pageHostIgnoreSize = false; });

    if (state.pageHostWindow)
    {
        static_cast<void>(EnsureActivePreferencesPageInitialized(state.pageHostWindow, state));
        RefreshActivePreferencesPage(state.pageHostWindow, state);
    }

    LayoutPreferencesDialog(dlg, state);

    if (state.pageHostWindow)
    {
        LayoutPreferencesPageHost(state.pageHostWindow, state);
    }

    // Force a single synchronous repaint — see comment in UpdatePageText.
    RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
}
[[nodiscard]] int MeasurePageHostContentHeightPx(HWND host, const PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return 0;
    }

    const HWND paneRoot = GetActivePreferencesPageRootWindow(state);
    if (! paneRoot || IsWindow(paneRoot) == FALSE)
    {
        return 0;
    }

    int maxBottomPx = 0;

    HWND current = (paneRoot == host) ? GetWindow(paneRoot, GW_CHILD) : paneRoot;
    while (current)
    {
        if (PrefsUi::IsActuallyVisibleChildWindow(current))
        {
            RECT rc{};
            if (GetWindowRect(current, &rc))
            {
                MapWindowPoints(nullptr, host, reinterpret_cast<POINT*>(&rc), 2);
                maxBottomPx = std::max(maxBottomPx, static_cast<int>(rc.bottom));
            }
        }

        HWND next = GetWindow(current, GW_CHILD);
        if (next)
        {
            current = next;
            continue;
        }

        while (current)
        {
            next = GetWindow(current, GW_HWNDNEXT);
            if (next && GetParent(next) == GetParent(current))
            {
                current = next;
                break;
            }

            current = GetParent(current);
            if (! current || current == paneRoot)
            {
                current = nullptr;
                break;
            }
        }
    }

    const auto measureDxBranchBottomPx =
        [dpi = GetDpiForWindow(host)](const RedSalamander::DxUi::Control* control, bool includeSelf, const auto& self) noexcept -> int
    {
        if (! control || ! control->IsVisible() || dpi == 0u)
        {
            return 0;
        }

        int maxBottom = 0;
        if (includeSelf)
        {
            const D2D1_RECT_F bounds = control->GetBounds();
            maxBottom                = static_cast<int>(std::lround((bounds.bottom * static_cast<float>(dpi)) / 96.0f));
        }
        if (const auto* panel = dynamic_cast<const Panel*>(control))
        {
            for (const auto& child : panel->GetChildren())
            {
                maxBottom = std::max(maxBottom, self(child.get(), true, self));
            }
        }
        return maxBottom;
    };

    const int directDxBottomPx = std::max(measureDxBranchBottomPx(state.pageHostDxContentRootControl, false, measureDxBranchBottomPx),
                                          measureDxBranchBottomPx(state.pageHostDxNoteControl, true, measureDxBranchBottomPx));

    return std::max({0, maxBottomPx, state.pageHostDirectContentBottomPx, directDxBottomPx});
}

void SyncPageHostDxContentRoot(HWND host, PreferencesDialogState& state, int contentHeightPx) noexcept
{
    if (! host)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(host);
    if (dpi == 0u)
    {
        return;
    }

    RECT client{};
    if (GetClientRect(host, &client) == FALSE)
    {
        return;
    }

    const int clientWidthPx  = std::max(0l, client.right - client.left);
    const int clientHeightPx = std::max(0l, client.bottom - client.top);
    const auto pxToDip       = [dpi](int px) noexcept { return (static_cast<float>(px) * 96.0f) / static_cast<float>(dpi); };

    const float viewportWidthDip  = pxToDip(clientWidthPx);
    const float viewportHeightDip = pxToDip(clientHeightPx);
    const float contentHeightDip  = (std::max)(viewportHeightDip, pxToDip(std::max(contentHeightPx, clientHeightPx)));
    const float scrollDip         = pxToDip(state.pageScrollY);

    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    if (hostState._pageHostContentRootControl)
    {
        hostState._pageHostContentRootControl->SetBounds(D2D1::RectF(0.0f, 0.0f, viewportWidthDip, viewportHeightDip));
        if (auto* scrollPanel = dynamic_cast<ScrollPanel*>(hostState._pageHostContentRootControl))
        {
            state.pageHostSyncingScrollPanel = true;
            const auto clearSync             = wil::scope_exit([&]() noexcept { state.pageHostSyncingScrollPanel = false; });
            scrollPanel->SetInternalScrollbarEnabled(true);
            scrollPanel->SetContentHeight(contentHeightDip);
            scrollPanel->SetScrollOffset(scrollDip);
        }
    }

    const D2D1_RECT_F wrapperBounds = D2D1::RectF(0.0f, 0.0f, viewportWidthDip, contentHeightDip);
    for (auto* wrapper : state.paneWrapperPanels)
    {
        if (wrapper)
        {
            wrapper->SetBounds(wrapperBounds);
        }
    }
}

void UpdatePageHostScrollInfo(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    RECT client{};
    GetClientRect(host, &client);
    const int clientHeight = std::max(0l, client.bottom - client.top);

    const UINT dpi          = GetDpiForWindow(host);
    const int paddingBottom = UiMetrics::ScaleDip(dpi, 12);
    int contentHeight       = MeasurePageHostContentHeightPx(host, state);
    for (const RECT& card : state.pageSettingCards)
    {
        contentHeight = std::max(contentHeight, static_cast<int>(card.bottom));
    }
    contentHeight        = std::max(0, contentHeight);
    const int overflowPx = std::max(0, contentHeight - clientHeight);
    if (overflowPx > paddingBottom)
    {
        contentHeight += paddingBottom;
    }
    else if (overflowPx > 0)
    {
        contentHeight = clientHeight;
    }

    const int maxScroll     = std::max(0, contentHeight - clientHeight);
    state.pageScrollMaxY = maxScroll;
    state.pageScrollY    = std::clamp(state.pageScrollY, 0, maxScroll);

    const LONG_PTR styleNow = GetWindowLongPtrW(host, GWL_STYLE);
    LONG_PTR styleWanted    = styleNow;
    styleWanted &= ~WS_HSCROLL;
    styleWanted &= ~WS_VSCROLL;

    if (styleWanted != styleNow)
    {
        state.pageHostIgnoreSize = true;
        auto clearIgnore         = wil::scope_exit([&]() noexcept { state.pageHostIgnoreSize = false; });

        SetWindowLongPtrW(host, GWL_STYLE, styleWanted);
        SetWindowPos(host, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        SendMessageW(host, WM_THEMECHANGED, 0, 0);
        // Avoid RDW_UPDATENOW — this runs inside ScopedWindowRedrawBlock which
        // already defers painting.  Synchronous paint here added 500ms+ per switch.
        RedrawWindow(host, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME);
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = (contentHeight > 0) ? (contentHeight - 1) : 0;
    si.nPage  = static_cast<UINT>(clientHeight);
    si.nPos   = state.pageScrollY;
    SetScrollInfo(host, SB_VERT, &si, TRUE);

    SyncPageHostDxContentRoot(host, state, contentHeight);
}

void ApplyPageHostScrollFromLayout(HWND host, const PreferencesDialogState& state) noexcept
{
    if (! host || state.pageScrollY == 0)
    {
        return;
    }

    PrefsPageHost::ApplyScrollDelta(host, -state.pageScrollY);
    // Avoid RDW_UPDATENOW — caller (UpdatePageText) already manages redraw
    // blocks and will coalesce invalidations into a single WM_PAINT pass.
    RedrawWindow(host, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME);
}

void FinalizePreferencesPageHostLayout(HWND host, PreferencesDialogState& state, int margin, int layoutWidth) noexcept
{
    if (! host)
    {
        return;
    }

    UpdatePageHostScrollInfo(host, state);
    ApplyPageHostScrollFromLayout(host, state);

    if (state.pageHostRelayoutInProgress)
    {
        return;
    }

    RECT client{};
    GetClientRect(host, &client);
    const int clientWidth = std::max(0l, client.right - client.left);
    const UINT dpi        = GetDpiForWindow(host);
    const int scrollbarGap = state.pageScrollMaxY > 0 ? UiMetrics::ScaleDip(dpi, 6) : 0;
    const int widthNow     = std::max(0, clientWidth - 2 * margin - scrollbarGap);
    if (widthNow == layoutWidth)
    {
        return;
    }

    state.pageHostRelayoutInProgress = true;
    LayoutPreferencesPageHost(host, state);
    state.pageHostRelayoutInProgress = false;
}

void SyncPreferencesPageHostDxSize(HWND host, PreferencesDialogState& state, const wchar_t* reason) noexcept
{
    if (! host || ! state.pageHostUsesDxUi)
    {
        return;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    RECT rc{};
    GetClientRect(host, &rc);
    const int widthPx  = std::max(0l, rc.right - rc.left);
    const int heightPx = std::max(0l, rc.bottom - rc.top);
    bool handled       = false;
    hostState._pageHostHost.HandleMessage(host, WM_SIZE, SIZE_RESTORED, MAKELPARAM(widthPx, heightPx), handled);

    Debug::Info(
        L"Preferences: synced page-host DxUi size reason={} category={} client={}x{} dx={}x{} handled={} renderCount={} resizeCount={} resizeFailures={}",
        reason ? reason : L"(null)",
        GetPrefCategoryDebugName(state.currentCategory),
        widthPx,
        heightPx,
        GetDxHostDebugWidthPx(hostState._pageHostHost),
        GetDxHostDebugHeightPx(hostState._pageHostHost),
        handled ? L"true" : L"false",
        GetDxHostDebugRenderCount(hostState._pageHostHost),
        GetDxHostDebugResizeCount(hostState._pageHostHost),
        GetDxHostDebugResizeFailureCount(hostState._pageHostHost));
}

void LogPreferencesPageHostState(HWND host, const PreferencesDialogState& state, const wchar_t* reason) noexcept
{
    if (! host)
    {
        return;
    }

    RECT rc{};
    GetClientRect(host, &rc);
    const int widthPx  = std::max(0l, rc.right - rc.left);
    const int heightPx = std::max(0l, rc.bottom - rc.top);

    const auto& hostState  = static_cast<const PreferencesDialogHost&>(state);
    size_t wrapperChildren = 0u;
    if (const auto* wrapperPanel = dynamic_cast<const Panel*>(state.pageHostDxContentRootControl))
    {
        wrapperChildren = wrapperPanel->GetChildren().size();
    }

    Debug::Info(
        L"Preferences: page-host state reason={} category={} client={}x{} scroll={}/{} dx={}x{} focus={} bridge={} wrapperChildren={} pageHostUsesDxUi={}",
        reason ? reason : L"(null)",
        GetPrefCategoryDebugName(state.currentCategory),
        widthPx,
        heightPx,
        state.pageScrollY,
        state.pageScrollMaxY,
        GetDxHostDebugWidthPx(hostState._pageHostHost),
        GetDxHostDebugHeightPx(hostState._pageHostHost),
        static_cast<const void*>(hostState._pageHostHost.GetFocusControl()),
        hostState._pageHostHost.HasActiveTextInputBridge() ? L"true" : L"false",
        wrapperChildren,
        state.pageHostUsesDxUi ? L"true" : L"false");
}

[[nodiscard]] HWND FindFirstOrLastTabStopChild(HWND host, bool forward) noexcept
{
    if (! host)
    {
        return nullptr;
    }

    const HWND dlg = GetParent(host);
    if (! dlg)
    {
        return nullptr;
    }

    const BOOL previous = forward ? FALSE : TRUE;
    const HWND start    = GetNextDlgTabItem(dlg, nullptr, previous);
    if (! start)
    {
        return nullptr;
    }

    HWND item = start;
    do
    {
        if (IsChild(host, item) && PrefsUi::IsActuallyVisibleChildWindow(item) && IsWindowEnabled(item))
        {
            const LONG_PTR style = GetWindowLongPtrW(item, GWL_STYLE);
            if ((style & WS_TABSTOP) != 0)
            {
                return item;
            }
        }

        item = GetNextDlgTabItem(dlg, item, previous);
    } while (item && item != start);

    return nullptr;
}

void LayoutPreferencesDialog(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(state);

    HWND list     = state.categoryTreeWindow ? state.categoryTreeWindow : GetDlgItem(dlg, IDC_PREFS_CATEGORY_LIST);
    HWND host     = state.pageHostWindow ? state.pageHostWindow : GetDlgItem(dlg, IDC_PREFS_PAGE_HOST);
    HWND ok       = GetDlgItem(dlg, IDOK);
    HWND cancel   = GetDlgItem(dlg, IDCANCEL);
    HWND apply    = GetDlgItem(dlg, IDC_PREFS_APPLY);
    HWND resetAll = GetDlgItem(dlg, IDC_PREFS_RESET_ALL);
    if (! list || ! host || ! ok || ! cancel || ! apply)
    {
        return;
    }

    RECT client{};
    GetClientRect(dlg, &client);

    const UINT dpi   = GetDpiForWindow(dlg);
    const int margin = UiMetrics::ScaleDip(dpi, 8);
    const int gapX   = UiMetrics::ScaleDip(dpi, 8);

    RECT okRect{};
    RECT cancelRect{};
    RECT applyRect{};
    GetWindowRect(ok, &okRect);
    GetWindowRect(cancel, &cancelRect);
    GetWindowRect(apply, &applyRect);

    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&okRect), 2);
    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&cancelRect), 2);
    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&applyRect), 2);

    const int okWidthDesired     = std::max(0l, okRect.right - okRect.left);
    const int cancelWidthDesired = std::max(0l, cancelRect.right - cancelRect.left);
    const int applyWidthDesired  = std::max(0l, applyRect.right - applyRect.left);

    const int measuredButtonHeight = static_cast<int>(std::max(
        {0l, std::max(0l, okRect.bottom - okRect.top), std::max(0l, cancelRect.bottom - cancelRect.top), std::max(0l, applyRect.bottom - applyRect.top)}));
    const int buttonHeight         = std::max(UiMetrics::ScaleDip(dpi, 26), measuredButtonHeight);

    const int buttonPadX                               = UiMetrics::ScaleDip(dpi, 12);
    const int minGapX                                  = UiMetrics::ScaleDip(dpi, 4);
    const PreferencesTypographyContext shellTypography = PrefsUi::MakeTypographyContext(dlg);

    const auto measureButtonMinWidth = [&](HWND button) noexcept
    {
        if (! button)
        {
            return 0;
        }

        const std::wstring text = PrefsUi::GetWindowTextString(button);
        const int textW         = PrefsUi::MeasureSingleLineTextWidthPx(shellTypography, shellTypography.body, text);
        return std::max(UiMetrics::ScaleDip(dpi, 60), textW + 2 * buttonPadX);
    };

    const int okWidthMin       = measureButtonMinWidth(ok);
    const int cancelWidthMin   = measureButtonMinWidth(cancel);
    const int applyWidthMin    = measureButtonMinWidth(apply);
    const int resetAllWidthMin = resetAll ? measureButtonMinWidth(resetAll) : 0;

    const int clientWidth         = std::max(0l, client.right - client.left);
    const int groupAvailableWidth = std::max(0, clientWidth - 2 * margin);
    int gapUsed                   = gapX;
    int minGroupWidth             = okWidthMin + cancelWidthMin + applyWidthMin + 2 * gapUsed;
    if (minGroupWidth > groupAvailableWidth)
    {
        gapUsed       = minGapX;
        minGroupWidth = okWidthMin + cancelWidthMin + applyWidthMin + 2 * gapUsed;
    }

    int okWidth     = okWidthDesired;
    int cancelWidth = cancelWidthDesired;
    int applyWidth  = applyWidthDesired;

    const int desiredGroupWidth = okWidth + cancelWidth + applyWidth + 2 * gapUsed;
    if (desiredGroupWidth > groupAvailableWidth)
    {
        okWidth     = okWidthMin;
        cancelWidth = cancelWidthMin;
        applyWidth  = applyWidthMin;

        int remaining   = std::max(0, groupAvailableWidth - 2 * gapUsed - (okWidth + cancelWidth + applyWidth));
        const auto grow = [&](int& width, int desired) noexcept
        {
            const int target = std::max(width, desired);
            const int add    = std::min(remaining, target - width);
            if (add > 0)
            {
                width += add;
                remaining -= add;
            }
        };

        grow(applyWidth, applyWidthDesired);
        grow(cancelWidth, cancelWidthDesired);
        grow(okWidth, okWidthDesired);
    }

    // Last-resort safety: avoid overlap if the window was resized smaller than the computed minimum.
    int finalGroupWidth = okWidth + cancelWidth + applyWidth + 2 * gapUsed;
    if (finalGroupWidth > groupAvailableWidth)
    {
        gapUsed                 = minGapX;
        int availableForButtons = std::max(0, groupAvailableWidth - 2 * gapUsed);
        if (availableForButtons < 3)
        {
            gapUsed             = 0;
            availableForButtons = std::max(0, groupAvailableWidth);
        }

        const int baseWidth = std::max(1, availableForButtons / 3);
        okWidth             = baseWidth;
        cancelWidth         = baseWidth;
        applyWidth          = baseWidth;

        const int remainder = std::max(0, availableForButtons - (baseWidth * 3));
        applyWidth += remainder;
        finalGroupWidth = okWidth + cancelWidth + applyWidth + 2 * gapUsed;
    }

    const int applyLeft  = static_cast<int>(client.right) - margin - applyWidth;
    const int cancelLeft = applyLeft - gapUsed - cancelWidth;
    const int okLeft     = cancelLeft - gapUsed - okWidth;
    const int buttonsTop = std::max(0, static_cast<int>(client.bottom) - margin - buttonHeight);

    // Reset All button is left-aligned, independent of the right-aligned OK/Cancel/Apply group.
    const int resetAllLeft  = margin;
    const int resetAllWidth = resetAllWidthMin;

    const UINT moveFlags       = SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS;
    const UINT dxShowMoveFlags = moveFlags | SWP_SHOWWINDOW;

    const int contentTop    = margin;
    const int contentBottom = std::max(contentTop, buttonsTop - margin);
    const int footerPad     = UiMetrics::ScaleDip(dpi, 10);
    const int footerTop     = std::max(contentTop, buttonsTop - footerPad);
    const int footerHeight  = std::max(0, std::min(static_cast<int>(client.bottom) - footerTop, buttonHeight + (2 * footerPad)));
    const int listBottom    = std::max(contentTop, static_cast<int>(client.bottom) - margin);
    const int listHeight    = std::max(0, listBottom - contentTop);

    const int listDesiredWidth = state.categoryListWidthPx > 0 ? state.categoryListWidthPx : UiMetrics::ScaleDip(dpi, 120);
    const int listMinWidth     = UiMetrics::ScaleDip(dpi, 72);
    const int hostMinWidth     = UiMetrics::ScaleDip(dpi, 140);

    const int availableForList = std::max(0, groupAvailableWidth - gapX - hostMinWidth);
    const int listMaxWidth     = std::max(listMinWidth, availableForList);
    const int listWidth        = std::clamp(listDesiredWidth, listMinWidth, listMaxWidth);

    const int hostLeft  = std::max(0, margin + listWidth + gapX);
    const int hostWidth = std::max(0, static_cast<int>(client.right) - margin - hostLeft);

    const int hostTop    = contentTop;
    const int hostHeight = std::max(0, contentBottom - hostTop);

    const auto needsWindowMove = [&](HWND hwnd, int left, int top, int width, int height, UINT flags) noexcept
    {
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            return false;
        }

        RECT windowRect{};
        if (! GetWindowRect(hwnd, &windowRect))
        {
            return true;
        }

        POINT corners[2] = {
            {windowRect.left, windowRect.top},
            {windowRect.right, windowRect.bottom},
        };

        if (const HWND parent = GetParent(hwnd); parent && IsWindow(parent) != FALSE)
        {
            MapWindowPoints(nullptr, parent, corners, 2);
        }

        const int currentWidth  = corners[1].x - corners[0].x;
        const int currentHeight = corners[1].y - corners[0].y;
        const bool sameRect     = corners[0].x == left && corners[0].y == top && currentWidth == width && currentHeight == height;

        if (! sameRect)
        {
            return true;
        }

        const bool visible = PrefsUi::IsActuallyVisibleChildWindow(hwnd);
        if ((flags & SWP_SHOWWINDOW) != 0u && ! visible)
        {
            return true;
        }
        if ((flags & SWP_HIDEWINDOW) != 0u && visible)
        {
            return true;
        }

        return false;
    };

    bool moved = false;
    if (HDWP hdwp = BeginDeferWindowPos(13))
    {
        auto deferWindow = [&](HWND hwnd, int left, int top, int width, int height, const UINT flags) noexcept
        {
            if (! hwnd || ! hdwp)
            {
                return;
            }
            if (! needsWindowMove(hwnd, left, top, width, height, flags))
            {
                return;
            }
            hdwp = DeferWindowPos(hdwp, hwnd, nullptr, left, top, width, height, flags);
        };

        deferWindow(apply, applyLeft, buttonsTop, applyWidth, buttonHeight, moveFlags);
        deferWindow(cancel, cancelLeft, buttonsTop, cancelWidth, buttonHeight, moveFlags);
        deferWindow(ok, okLeft, buttonsTop, okWidth, buttonHeight, moveFlags);
        if (resetAll)
        {
            deferWindow(resetAll, resetAllLeft, buttonsTop, resetAllWidth, buttonHeight, moveFlags);
        }
        deferWindow(list, margin, contentTop, listWidth, listHeight, moveFlags);
        if (hostState._shellHostHwnd)
        {
            const int shellWidth  = hostWidth;
            const int shellHeight = footerHeight;
            deferWindow(hostState._shellHostHwnd, hostLeft, footerTop, shellWidth, shellHeight, dxShowMoveFlags);
        }
        deferWindow(host, hostLeft, hostTop, hostWidth, hostHeight, moveFlags);

        if (hdwp && EndDeferWindowPos(hdwp))
        {
            moved = true;
        }
    }

    if (! moved)
    {
        if (needsWindowMove(apply, applyLeft, buttonsTop, applyWidth, buttonHeight, moveFlags))
        {
            SetWindowPos(apply, nullptr, applyLeft, buttonsTop, applyWidth, buttonHeight, moveFlags);
        }
        if (needsWindowMove(cancel, cancelLeft, buttonsTop, cancelWidth, buttonHeight, moveFlags))
        {
            SetWindowPos(cancel, nullptr, cancelLeft, buttonsTop, cancelWidth, buttonHeight, moveFlags);
        }
        if (needsWindowMove(ok, okLeft, buttonsTop, okWidth, buttonHeight, moveFlags))
        {
            SetWindowPos(ok, nullptr, okLeft, buttonsTop, okWidth, buttonHeight, moveFlags);
        }
        if (resetAll && needsWindowMove(resetAll, resetAllLeft, buttonsTop, resetAllWidth, buttonHeight, moveFlags))
        {
            SetWindowPos(resetAll, nullptr, resetAllLeft, buttonsTop, resetAllWidth, buttonHeight, moveFlags);
        }
        if (needsWindowMove(list, margin, contentTop, listWidth, listHeight, moveFlags))
        {
            SetWindowPos(list, nullptr, margin, contentTop, listWidth, listHeight, moveFlags);
        }
        if (hostState._shellHostHwnd)
        {
            const int shellWidth  = hostWidth;
            const int shellHeight = footerHeight;
            if (needsWindowMove(hostState._shellHostHwnd, hostLeft, footerTop, shellWidth, shellHeight, dxShowMoveFlags))
            {
                SetWindowPos(hostState._shellHostHwnd, nullptr, hostLeft, footerTop, shellWidth, shellHeight, dxShowMoveFlags);
            }
        }
        if (needsWindowMove(host, hostLeft, hostTop, hostWidth, hostHeight, moveFlags))
        {
            SetWindowPos(host, nullptr, hostLeft, hostTop, hostWidth, hostHeight, moveFlags);
        }
    }

    if (hostState._shellHostHwnd)
    {
        static_cast<void>(SetWindowRgn(hostState._shellHostHwnd, nullptr, FALSE));

        // Shell host DxUi controls expect DIP coordinates, not raw pixels.
        const auto shellToDip = [&](float px) noexcept { return hostState._shellHost.PixelsToDip(px); };
        const int buttonTopInShell = buttonsTop - footerTop;

        if (hostState._okButtonControl)
        {
            hostState._okButtonControl->SetVisible(true);
            hostState._okButtonControl->SetBounds(D2D1::RectF(shellToDip(static_cast<float>(okLeft - hostLeft)),
                                                              shellToDip(static_cast<float>(buttonTopInShell)),
                                                              shellToDip(static_cast<float>(okLeft - hostLeft + okWidth)),
                                                              shellToDip(static_cast<float>(buttonTopInShell + buttonHeight))));
        }
        if (hostState._cancelButtonControl)
        {
            hostState._cancelButtonControl->SetVisible(true);
            hostState._cancelButtonControl->SetBounds(D2D1::RectF(shellToDip(static_cast<float>(cancelLeft - hostLeft)),
                                                                  shellToDip(static_cast<float>(buttonTopInShell)),
                                                                  shellToDip(static_cast<float>(cancelLeft - hostLeft + cancelWidth)),
                                                                  shellToDip(static_cast<float>(buttonTopInShell + buttonHeight))));
        }
        if (hostState._applyButtonControl)
        {
            hostState._applyButtonControl->SetVisible(true);
            hostState._applyButtonControl->SetBounds(D2D1::RectF(shellToDip(static_cast<float>(applyLeft - hostLeft)),
                                                                 shellToDip(static_cast<float>(buttonTopInShell)),
                                                                 shellToDip(static_cast<float>(applyLeft - hostLeft + applyWidth)),
                                                                 shellToDip(static_cast<float>(buttonTopInShell + buttonHeight))));
        }
        if (hostState._resetAllButtonControl && resetAll)
        {
            hostState._resetAllButtonControl->SetVisible(true);
            hostState._resetAllButtonControl->SetBounds(D2D1::RectF(shellToDip(static_cast<float>(resetAllLeft - hostLeft)),
                                                                    shellToDip(static_cast<float>(buttonTopInShell)),
                                                                    shellToDip(static_cast<float>(resetAllLeft - hostLeft + resetAllWidth)),
                                                                    shellToDip(static_cast<float>(buttonTopInShell + buttonHeight))));
        }
        hostState._shellHost.Invalidate();
        RedrawWindow(hostState._shellHostHwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    }
}

void LayoutPreferencesPageHost(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    // NOTE: EnsureActivePreferencesPageInitialized is NOT called here.
    // It is already called by UpdatePageText (line 3686) before layout.
    // Calling it again doubled the cost of CreateControls on every switch.

    RECT client{};
    GetClientRect(host, &client);

    auto& hostState = static_cast<PreferencesDialogHost&>(state);

    const UINT dpi     = GetDpiForWindow(host);
    const int margin   = UiMetrics::ScaleDip(dpi, 12);
    const int gapY     = UiMetrics::ScaleDip(dpi, 6);
    const int sectionY = UiMetrics::ScaleDip(dpi, 14);

    const int scrollbarGap = state.pageScrollMaxY > 0 ? UiMetrics::ScaleDip(dpi, 6) : 0;
    const int width = std::max(0l, client.right - client.left - 2 * margin - scrollbarGap);
    int x           = margin;
    int y           = margin;

    const PreferencesTypographyContext pageTypography = PrefsUi::MakeTypographyContext(host);
    state.pageHostDirectContentBottomPx               = 0;

    // Page host DxUi controls expect DIP coordinates, not raw pixels.
    const auto pageToDip      = [&](float px) noexcept { return hostState._pageHostHost.PixelsToDip(px); };
    const float pageWidthDip  = pageToDip(static_cast<float>(std::max(0l, client.right - client.left)));
    const float pageHeightDip = pageToDip(static_cast<float>(std::max(0l, client.bottom - client.top)));

    if (hostState._pageHostRootControl)
    {
        hostState._pageHostRootControl->SetBounds(D2D1::RectF(0.0f, 0.0f, pageWidthDip, pageHeightDip));
    }
    if (hostState._pageHostSurfaceControl)
    {
        hostState._pageHostSurfaceControl->SetBounds(D2D1::RectF(0.0f, 0.0f, pageWidthDip, pageHeightDip));
    }
    if (hostState._pageHostContentRootControl)
    {
        hostState._pageHostContentRootControl->SetBounds(D2D1::RectF(0.0f, 0.0f, pageWidthDip, pageHeightDip));
        if (auto* scrollPanel = dynamic_cast<ScrollPanel*>(hostState._pageHostContentRootControl))
        {
            state.pageHostSyncingScrollPanel = true;
            const auto clearSync             = wil::scope_exit([&]() noexcept { state.pageHostSyncingScrollPanel = false; });
            scrollPanel->SetInternalScrollbarEnabled(true);
            scrollPanel->SetContentHeight((std::max)(pageHeightDip, scrollPanel->GetContentHeight()));
            scrollPanel->SetScrollOffset(hostState._pageHostHost.PixelsToDip(static_cast<float>(state.pageScrollY)));
        }
    }
    for (auto* wrapper : state.paneWrapperPanels)
    {
        if (wrapper)
        {
            wrapper->SetBounds(D2D1::RectF(0.0f, 0.0f, pageWidthDip, pageHeightDip));
        }
    }
    PrefsUi::HideSharedPageEmptyState(state);

    const auto setLabelBounds = [&](Label* label, const int left, const int top, const int right, const int bottom) noexcept
    {
        if (label)
        {
            label->SetBounds(D2D1::RectF(pageToDip(static_cast<float>(left)),
                                         pageToDip(static_cast<float>(top)),
                                         pageToDip(static_cast<float>(right)),
                                         pageToDip(static_cast<float>(bottom))));
        }
    };

    const int headerGapY     = gapY;
    const int headerSectionY = sectionY;
    const std::wstring titleText = hostState._pageTitleControl ? std::wstring(hostState._pageTitleControl->GetText()) : std::wstring{};
    if (! titleText.empty())
    {
        const int measuredTitleHeight = PrefsUi::MeasureWrappedTextHeightPx(pageTypography, pageTypography.title, width, titleText);
        const int titleHeight         = std::max(UiMetrics::ScaleDip(dpi, 40), std::max(0, measuredTitleHeight));
        setLabelBounds(hostState._pageTitleControl, x, y, x + width, y + titleHeight);
        y += titleHeight + headerGapY;
    }
    else
    {
        setLabelBounds(hostState._pageTitleControl, x, y, x + width, y);
    }

    const std::wstring descText = hostState._pageDescriptionControl ? std::wstring(hostState._pageDescriptionControl->GetText()) : std::wstring{};
    if (! descText.empty())
    {
        const int descHeight = std::max(0, PrefsUi::MeasureWrappedTextHeightPx(pageTypography, pageTypography.body, width, descText));
        setLabelBounds(hostState._pageDescriptionControl, x, y, x + width, y + descHeight);
        y += descHeight + headerSectionY;
    }
    else
    {
        setLabelBounds(hostState._pageDescriptionControl, x, y, x + width, y);
    }
    state.pageHostDirectContentBottomPx = (std::max)(state.pageHostDirectContentBottomPx, y);

    const bool showGeneral              = state.currentCategory == PrefCategory::General;
    const bool showPanes                = state.currentCategory == PrefCategory::Panes;
    const bool showViewers              = state.currentCategory == PrefCategory::Viewers;
    const bool showEditors              = state.currentCategory == PrefCategory::Editors;
    const bool showUserMenu             = state.currentCategory == PrefCategory::UserMenu;
    const bool showKeyboard             = state.currentCategory == PrefCategory::Keyboard;
    const bool showMouse                = state.currentCategory == PrefCategory::Mouse;
    const bool showThemes               = state.currentCategory == PrefCategory::Themes;
    const bool showPlugins              = state.currentCategory == PrefCategory::Plugins;
    const bool showFileOperations       = state.currentCategory == PrefCategory::FileOperations;
    const bool showCompareDirectories   = state.currentCategory == PrefCategory::CompareDirectories;
    const bool showHotPaths             = state.currentCategory == PrefCategory::HotPaths;
    const bool showAdvanced             = state.currentCategory == PrefCategory::Advanced;
    const bool notePageSkipsHostTabStop = showMouse;

    if (host)
    {
        const LONG_PTR style        = GetWindowLongPtrW(host, GWL_STYLE);
        const LONG_PTR desiredStyle = notePageSkipsHostTabStop ? (style & ~static_cast<LONG_PTR>(WS_TABSTOP)) : (style | static_cast<LONG_PTR>(WS_TABSTOP));
        if (desiredStyle != style)
        {
            SetWindowLongPtrW(host, GWL_STYLE, desiredStyle);
            SetWindowPos(host, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        }
    }

    state.pageSettingCards.clear();

    // Win32 title/description/buttons are hidden by CreatePreferencesShellHosts;
    // DxUi replacements are always active — no runtime visibility toggling needed.

    auto setVisible = [&](const auto& hwndLike, bool visible) noexcept
    {
        HWND hwnd = nullptr;
        if constexpr (requires { hwndLike.get(); })
        {
            hwnd = hwndLike.get();
        }
        else
        {
            hwnd = hwndLike;
        }

        if (hwnd)
        {
            ShowWindow(hwnd, visible ? SW_SHOW : SW_HIDE);
        }
    };

    hostState._generalPane.OnVisibilityChanged(showGeneral);
    hostState._panesPane.OnVisibilityChanged(showPanes);
    hostState._viewersPane.OnVisibilityChanged(showViewers);
    hostState._editorsPane.OnVisibilityChanged(showEditors);
    hostState._userMenuPane.OnVisibilityChanged(showUserMenu);
    hostState._keyboardPane.OnVisibilityChanged(showKeyboard);
    hostState._mousePane.OnVisibilityChanged(showMouse);
    hostState._themesPane.OnVisibilityChanged(showThemes);
    hostState._pluginsPane.OnVisibilityChanged(showPlugins);
    hostState._fileOperationsPane.OnVisibilityChanged(showFileOperations);
    hostState._compareDirectoriesPane.OnVisibilityChanged(showCompareDirectories);
    hostState._hotPathsPane.OnVisibilityChanged(showHotPaths);
    hostState._advancedPane.OnVisibilityChanged(showAdvanced);

    if (showPanes)
    {
        const PreferencesTypographyContext typography = PrefsUi::MakeTypographyContext(host);
        hostState._panesPane.LayoutPage(host, state, x, y, width, margin, gapY, sectionY, typography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showViewers)
    {
        const PreferencesTypographyContext typography = PrefsUi::MakeTypographyContext(host);
        hostState._viewersPane.LayoutPage(host, state, x, y, width, margin, gapY, typography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showEditors)
    {
        static_cast<PreferencesDialogHost&>(state)._editorsPane.LayoutPage(host, state, x, y, width, margin, gapY, sectionY, pageTypography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showUserMenu)
    {
        static_cast<PreferencesDialogHost&>(state)._userMenuPane.LayoutPage(host, state, x, y, width, margin, gapY, sectionY, pageTypography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showMouse)
    {
        static_cast<PreferencesDialogHost&>(state)._mousePane.LayoutPage(host, state, x, y, width, margin, gapY, sectionY, pageTypography);
        const PreferencesEmptyStateSpec spec = GetCurrentPreferencesSharedEmptyState(state);
        if (! spec.title.empty() || ! spec.body.empty() || ! spec.caption.empty())
        {
            const int cardHeight = PrefsUi::ShowSharedPageEmptyState(host, state, spec, x, y, width, pageTypography);
            if (cardHeight > 0)
            {
                RECT card{x, y, x + width, y + cardHeight};
                PrefsUi::TryPushCard(state.pageSettingCards, card);
                state.pageHostDirectContentBottomPx = (std::max)(state.pageHostDirectContentBottomPx, static_cast<int>(card.bottom));
            }
        }
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showThemes)
    {
        hostState._themesPane.LayoutPage(host, state, x, y, width, margin, gapY, sectionY, pageTypography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showPlugins)
    {
        hostState._pluginsPane.LayoutPage(host, state, x, y, width, margin, gapY, sectionY, pageTypography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showFileOperations)
    {
        hostState._fileOperationsPane.LayoutPage(host, state, x, y, width, margin, gapY, pageTypography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showCompareDirectories)
    {
        hostState._compareDirectoriesPane.LayoutPage(host, state, x, y, width, margin, gapY, pageTypography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showHotPaths)
    {
        hostState._hotPathsPane.LayoutPage(host, state, x, y, width, margin, gapY, pageTypography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showAdvanced)
    {
        hostState._advancedPane.LayoutPage(host, state, x, y, width, margin, gapY, pageTypography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showKeyboard)
    {
        KeyboardPane::LayoutPage(host, state, x, y, width, margin, gapY, sectionY, pageTypography);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showGeneral)
    {
        const PreferencesTypographyContext typography = PrefsUi::MakeTypographyContext(host);
        static_cast<PreferencesDialogHost&>(state)._generalPane.LayoutPage(host, state, x, y, width, typography);
    }

    FinalizePreferencesPageHostLayout(host, state, margin, width);
}

void RefreshAdvancedPage(HWND host, PreferencesDialogState& state) noexcept
{
    static_cast<PreferencesDialogHost&>(state)._advancedPane.Refresh(host, state);
}

void RefreshActivePreferencesPage(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    switch (state.currentCategory)
    {
        case PrefCategory::General: static_cast<PreferencesDialogHost&>(state)._generalPane.Refresh(host, state); break;
        case PrefCategory::Keyboard: KeyboardPane::Refresh(host, state); break;
        case PrefCategory::Panes: static_cast<PreferencesDialogHost&>(state)._panesPane.Refresh(host, state); break;
        case PrefCategory::Viewers: static_cast<PreferencesDialogHost&>(state)._viewersPane.Refresh(host, state); break;
        case PrefCategory::Editors: static_cast<PreferencesDialogHost&>(state)._editorsPane.Refresh(host, state); break;
        case PrefCategory::UserMenu: static_cast<PreferencesDialogHost&>(state)._userMenuPane.Refresh(host, state); break;
        case PrefCategory::Themes: static_cast<PreferencesDialogHost&>(state)._themesPane.Refresh(host, state); break;
        case PrefCategory::Plugins: static_cast<PreferencesDialogHost&>(state)._pluginsPane.Refresh(host, state); break;
        case PrefCategory::FileOperations: static_cast<PreferencesDialogHost&>(state)._fileOperationsPane.Refresh(host, state); break;
        case PrefCategory::CompareDirectories: static_cast<PreferencesDialogHost&>(state)._compareDirectoriesPane.Refresh(host, state); break;
        case PrefCategory::HotPaths: static_cast<PreferencesDialogHost&>(state)._hotPathsPane.Refresh(host, state); break;
        case PrefCategory::Advanced: RefreshAdvancedPage(host, state); break;
        case PrefCategory::Mouse: break;
    }
}

void UpdatePageText(HWND dlg, PreferencesDialogState& state, PrefCategory category) noexcept
{
    if (state.updatingPageText)
    {
        Debug::Warning(
            L"Preferences: re-entrant UpdatePageText blocked (current={}, requested={})", static_cast<int>(state.currentCategory), static_cast<int>(category));
        return;
    }
    state.updatingPageText = true;
    const auto clearGuard  = wil::scope_exit([&]() noexcept { state.updatingPageText = false; });

    const PrefCategory previousCategory = state.currentCategory;
    const bool categoryChanged          = previousCategory != category;
    if (categoryChanged)
    {
        state.retainedPageScrollYByCategory[PrefCategoryIndex(previousCategory)] = state.pageScrollY;
    }
    state.currentCategory = category;
    if (categoryChanged)
    {
        state.pageScrollY = std::max(0, state.retainedPageScrollYByCategory[PrefCategoryIndex(category)]);
    }
    state.pageScrollMaxY          = 0;
    state.pageWheelDeltaRemainder = 0;

    const CategoryInfo* info = FindCategoryInfo(category);
    std::wstring title       = info ? LoadStringResource(nullptr, info->labelId) : std::wstring{};
    std::wstring description = info ? LoadStringResource(nullptr, info->descriptionId) : std::wstring{};

    if (category == PrefCategory::Plugins)
    {
        std::optional<PrefsPluginListItem> selectedPlugin;
        if (! state.pluginsSelectedPluginId.empty())
        {
            selectedPlugin = PrefsPlugins::FindItemById(state.pluginsSelectedPluginId);
            if (selectedPlugin.has_value())
            {
                state.pluginsSelectedPlugin = selectedPlugin;
                state.pluginsRetainedSelectedPluginId.assign(state.pluginsSelectedPluginId);
            }
            else
            {
                state.pluginsSelectedPlugin.reset();
                state.pluginsSelectedPluginId.clear();
                state.pluginsDetailsActive = false;
            }
        }
        else if (state.pluginsSelectedPlugin.has_value())
        {
            selectedPlugin                   = state.pluginsSelectedPlugin;
            const std::wstring_view pluginId = PrefsPlugins::GetId(state.pluginsSelectedPlugin.value());
            if (! pluginId.empty())
            {
                state.pluginsSelectedPluginId.assign(pluginId);
                state.pluginsRetainedSelectedPluginId.assign(pluginId);
            }
            else
            {
                selectedPlugin.reset();
                state.pluginsSelectedPlugin.reset();
                state.pluginsDetailsActive = false;
            }
        }

        if (state.pluginsDetailsActive && selectedPlugin.has_value())
        {
            const std::wstring_view pluginName = PrefsPlugins::GetDisplayName(selectedPlugin.value());
            if (! pluginName.empty())
            {
                title = std::wstring(pluginName);
            }

            const std::wstring_view pluginDescription = PrefsPlugins::GetDescription(selectedPlugin.value());
            if (! pluginDescription.empty())
            {
                description = std::wstring(pluginDescription);
            }
        }
    }
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_PREFS_CAPTION);
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(state);

    // No ScopedWindowRedrawBlock on the dialog.  WM_SETREDRAW FALSE/TRUE
    // toggles WS_VISIBLE, causing the DWM to clear the composition surface.
    // Between WM_SETREDRAW TRUE and the first Present1/GDI-BitBlt, the
    // compositor shows a blank white frame — the "white flash" bug.
    //
    // This is safe because WM_PAINT is a low-priority message: it's only
    // dispatched when the message loop has no other messages.  Since we're
    // executing inside a message handler (tree selection change), all the
    // InvalidateRect calls below accumulate silently and are coalesced into
    // a single paint after this function returns and the final
    // RedrawWindow(RDW_UPDATENOW) forces synchronous painting.

    UpdateDxShellText(hostState, title, description);

    if (categoryChanged)
    {
        DestroyInactivePreferencesPageState(state, previousCategory);

        // Reset per-pane layout state only — wrapper panels and their controls persist.
        if (state.pageHostDxHost)
        {
            state.pageHostDxHost->ResetInteractionState();
        }
        state.pageSettingCards.clear();
        state.pageHostDirectContentBottomPx = 0;
    }

    // Set pageHostIgnoreSize BEFORE creating pane controls so that any WM_SIZE
    // messages generated during EnsureCreated / CreateControls do not trigger
    // LayoutPreferencesPageHost with partially-initialized state.
    state.pageHostIgnoreSize = true;
    const auto clearIgnore   = wil::scope_exit([&]() noexcept { state.pageHostIgnoreSize = false; });

    if (state.pageHostWindow)
    {
        static_cast<void>(EnsureActivePreferencesPageInitialized(state.pageHostWindow, state));
    }

    if (state.pageHostWindow)
    {
        RefreshActivePreferencesPage(state.pageHostWindow, state);
    }

    if (dlg)
    {
        LayoutPreferencesDialog(dlg, state);
    }

    if (state.pageHostWindow)
    {
        LayoutPreferencesPageHost(state.pageHostWindow, state);
    }

    // After layout, forward the actual page host size to the DxUi host.
    // LayoutPreferencesDialog may have resized the page host window, but
    // WM_SIZE was blocked by pageHostIgnoreSize.  Without this explicit
    // sync the DxUi swap chain keeps the stale template-size dimensions
    // from Attach(), so the D2D render is clipped or stretched.
    if (state.pageHostWindow && state.pageHostUsesDxUi)
    {
        SyncPreferencesPageHostDxSize(state.pageHostWindow, state, L"page-switch");
    }

    if (state.pageHostWindow)
    {
        LogPreferencesPageHostState(state.pageHostWindow, state, L"page-switch");
    }

    // Force a single synchronous repaint of the dialog and all children.
    // All intermediate InvalidateRect calls during the update phase are
    // coalesced into this one paint pass — no intermediate frames.
    RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
}

void PopulateCategoryTree(HWND dlg, PreferencesDialogState& state) noexcept
{
    auto& hostState          = static_cast<PreferencesDialogHost&>(state);
    state.categoryTreeWindow = GetDlgItem(dlg, IDC_PREFS_CATEGORY_LIST);
    if (! state.categoryTreeWindow)
    {
        return;
    }

    state.categoryTreeUsesDxUi = true;

    RECT rc{};
    if (GetWindowRect(state.categoryTreeWindow, &rc))
    {
        state.categoryListWidthPx = std::max(0l, rc.right - rc.left);
    }

    if (hostState._categoryTreeHost.GetHwnd() != state.categoryTreeWindow)
    {
        hostState._categoryTreeHost.Detach();
        if (! hostState._categoryTreeHost.Attach(state.categoryTreeWindow))
        {
            Debug::Error(L"Preferences: failed to attach DxUi host for category tree.");
            state.categoryTreeUsesDxUi = false;
            return;
        }

        RestoreWndProcHook(state.categoryTreeWindow, kPrefsDxCategoryHostOriginalWndProcProp);
        SetPropW(state.categoryTreeWindow, kPrefsDxCategoryHostStateProp, &state);
        if (! InstallWndProcHook(state.categoryTreeWindow, kPrefsDxCategoryHostOriginalWndProcProp, PreferencesDxCategoryHostWndProc))
        {
            Debug::ErrorWithLastError(L"Preferences: failed to install category tree host window proc hook.");
            RemovePropW(state.categoryTreeWindow, kPrefsDxCategoryHostStateProp);
            hostState._categoryTreeHost.Detach();
            state.categoryTreeUsesDxUi = false;
            return;
        }

        auto tree                      = std::make_unique<Tree>();
        hostState._categoryTreeControl = tree.get();
        hostState._categoryTreeControl->SetModel(&hostState._categoryTreeModel);
        hostState._categoryTreeControl->SetDelegate(&hostState._categoryTreeDelegate);
        hostState._categoryTreeControl->SetRowHeightDip(24.0f);
        hostState._categoryTreeControl->SetIndentDip(14.0f);
        hostState._categoryTreeHost.SetRoot(std::move(tree));
    }

    const LONG_PTR categoryHostStyle = GetWindowLongPtrW(state.categoryTreeWindow, GWL_STYLE);
    if ((categoryHostStyle & SS_NOTIFY) == 0)
    {
        SetWindowLongPtrW(state.categoryTreeWindow, GWL_STYLE, categoryHostStyle | SS_NOTIFY);
    }

    hostState._categoryTreeDelegate.Attach(dlg, &state, &hostState._categoryTreeModel, hostState._categoryTreeControl);
    hostState._categoryTreeModel.Rebuild();
    hostState._categoryTreeHost.SetTheme(PrefsUi::MakeDxPalette(state.theme));
    hostState._categoryTreeHost.SetOnTabBoundary([dlg, categoryHost = state.categoryTreeWindow, hostStatePtr = &hostState](const bool reverse) noexcept
    {
        if (! dlg || IsWindow(dlg) == FALSE || ! categoryHost || IsWindow(categoryHost) == FALSE)
        {
            return false;
        }

        if (HWND target = GetNextDlgTabItem(dlg, categoryHost, reverse ? TRUE : FALSE); target && target != categoryHost)
        {
            if (reverse && hostStatePtr && target == hostStatePtr->_shellHostHwnd)
            {
                if (hostStatePtr->_applyButtonControl && hostStatePtr->_applyButtonControl->IsEnabled())
                {
                    hostStatePtr->_shellHost.SetFocusControl(hostStatePtr->_applyButtonControl);
                }
                else if (hostStatePtr->_cancelButtonControl && hostStatePtr->_cancelButtonControl->IsEnabled())
                {
                    hostStatePtr->_shellHost.SetFocusControl(hostStatePtr->_cancelButtonControl);
                }
                else if (hostStatePtr->_okButtonControl && hostStatePtr->_okButtonControl->IsEnabled())
                {
                    hostStatePtr->_shellHost.SetFocusControl(hostStatePtr->_okButtonControl);
                }
                else if (hostStatePtr->_resetAllButtonControl && hostStatePtr->_resetAllButtonControl->IsEnabled())
                {
                    hostStatePtr->_shellHost.SetFocusControl(hostStatePtr->_resetAllButtonControl);
                }
            }

            SetFocus(target);
            return true;
        }

        return false;
    });
    if (hostState._categoryTreeControl)
    {
        std::optional<uint64_t> selectedItemId;
        if (state.pluginsSelectedPlugin.has_value())
        {
            selectedItemId =
                PreferencesDialogHost::PreferencesCategoryTreeModel::EncodePluginNodeId(state.pluginsSelectedPlugin->type, state.pluginsSelectedPlugin->index);
        }
        else
        {
            selectedItemId = PreferencesDialogHost::PreferencesCategoryTreeModel::EncodeCategoryNodeId(state.currentCategory);
        }
        hostState._categoryTreeControl->SetSelectedItemId(selectedItemId);
        hostState._categoryTreeControl->NotifyDataChanged();
    }
}

void SelectCategory(HWND dlg, PreferencesDialogState& state, PrefCategory category) noexcept
{
    state.initialCategory = category;
    state.pluginsSelectedPlugin.reset();
    state.pluginsSelectedPluginId.clear();
    state.pluginsDetailsActive = false;

    if (! dlg || ! state.categoryTreeWindow)
    {
        return;
    }

    if (state.categoryTreeUsesDxUi)
    {
        auto& hostState = static_cast<PreferencesDialogHost&>(state);
        if (hostState._categoryTreeControl)
        {
            hostState._categoryTreeControl->SetSelectedItemId(PreferencesDialogHost::PreferencesCategoryTreeModel::EncodeCategoryNodeId(category));
            hostState._categoryTreeControl->NotifyDataChanged();
        }
        UpdatePageText(dlg, state, category);
        return;
    }
}

void CreatePageControls(HWND dlg, PreferencesDialogState& state) noexcept
{
    state.pageHostWindow = GetDlgItem(dlg, IDC_PREFS_PAGE_HOST);
    if (! state.pageHostWindow)
    {
        return;
    }

    InitPostedPayloadWindow(state.pageHostWindow);

    LONG_PTR exStyle = GetWindowLongPtrW(state.pageHostWindow, GWL_EXSTYLE);
    if ((exStyle & WS_EX_CONTROLPARENT) == 0)
    {
        exStyle |= WS_EX_CONTROLPARENT;
        SetWindowLongPtrW(state.pageHostWindow, GWL_EXSTYLE, exStyle);
    }

    LONG_PTR style    = GetWindowLongPtrW(state.pageHostWindow, GWL_STYLE);
    LONG_PTR newStyle = style;
    // Prevent the host from painting over its pane windows (avoids "blank until hover" artifacts).
    // Each pane paints its own themed background/cards.
    newStyle |= WS_CLIPCHILDREN;
    newStyle &= ~WS_HSCROLL;
    newStyle &= ~WS_VSCROLL;
    if (newStyle != style)
    {
        SetWindowLongPtrW(state.pageHostWindow, GWL_STYLE, newStyle);
        SetWindowPos(state.pageHostWindow, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    AttachPreferencesPageHostDxSurface(hostState);

    CreatePreferencesShellHosts(dlg, hostState);

    static_cast<void>(EnsureActivePreferencesPageInitialized(state.pageHostWindow, state));
    LayoutPreferencesPageHost(state.pageHostWindow, state);
}

LRESULT CALLBACK PreferencesPageHostWindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* state = reinterpret_cast<PreferencesDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! state)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);

    if (state->pageHostUsesDxUi)
    {
        switch (msg)
        {
            case WM_NCHITTEST:
            case WM_NCCALCSIZE:
            case WM_NCPAINT:
            case WM_NCLBUTTONDOWN:
            case WM_NCLBUTTONUP:
            case WM_NCLBUTTONDBLCLK:
            case WM_NCMOUSEMOVE: break;
            case WM_SIZE:
                if (state->pageHostIgnoreSize)
                {
                    if (IsPreferencesDxDiagnosticsEnabled(hwnd))
                    {
                        Debug::Info(L"Preferences: skipped forwarding ignored WM_SIZE to page-host DxUi hwnd={:#x} category={} lp=0x{:08X}",
                                    reinterpret_cast<uintptr_t>(hwnd),
                                    GetPrefCategoryDebugName(state->currentCategory),
                                    static_cast<unsigned long>(lp));
                    }
                    break;
                }
                [[fallthrough]];
            default:
            {
                if (msg == WM_NCDESTROY)
                {
                    hostState._pageHostHost.ReleaseMouseCapture();
                    state->pageHostUsesDxUi               = false;
                    state->pageHostDxHost                 = nullptr;
                    state->pageHostDxRootControl          = nullptr;
                    state->pageHostDxScrollPanelControl   = nullptr;
                    state->pageHostDxContentRootControl   = nullptr;
                    state->pageHostDxNoteControl          = nullptr;
                    return 0;
                }

                bool handled = false;
#ifdef ENABLE_TESTS
                const bool debugWheelMessage = msg == WM_MOUSEWHEEL;
                if (debugWheelMessage)
                {
                    state->debugLastWheelWndProcSeen      = true;
                    state->debugLastWheelDxHandled        = false;
                    state->debugLastWheelFallbackCalled   = false;
                    state->debugLastWheelFallbackHandled  = false;
                    state->debugLastWheelDelta            = static_cast<int>(GET_WHEEL_DELTA_WPARAM(wp));
                    state->debugLastWheelBeforeY          = state->pageScrollY;
                    state->debugLastWheelBeforeMaxY       = state->pageScrollMaxY;
                    state->debugLastWheelAfterY           = state->pageScrollY;
                    state->debugLastWheelAfterMaxY        = state->pageScrollMaxY;
                    POINT pageHostClientPoint{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
                    ScreenToClient(hwnd, &pageHostClientPoint);
                    state->debugLastWheelClientX = pageHostClientPoint.x;
                    state->debugLastWheelClientY = pageHostClientPoint.y;
                }
#endif
                const LRESULT dxResult = hostState._pageHostHost.HandleMessage(hwnd, msg, wp, lp, handled);
#ifdef ENABLE_TESTS
                if (debugWheelMessage)
                {
                    state->debugLastWheelDxHandled = handled;
                    state->debugLastWheelAfterY    = state->pageScrollY;
                    state->debugLastWheelAfterMaxY = state->pageScrollMaxY;
                }
#endif
                if (handled)
                {
                    return dxResult;
                }
                break;
            }
        }
    }

    if (IsPreferencesDxDiagnosticsEnabled(hwnd))
    {
        switch (msg)
        {
            case WM_MOUSEACTIVATE:
            case WM_LBUTTONDOWN:
            case WM_LBUTTONUP:
            case WM_RBUTTONDOWN:
            case WM_RBUTTONUP:
            case WM_MOUSEWHEEL:
            case WM_SETFOCUS:
            case WM_KILLFOCUS:
                Debug::Info(L"Preferences: page-host wndproc msg=0x{:04X} hwnd={:#x} category={} focus={:#x} pageHostFocus={:#x}",
                            msg,
                            reinterpret_cast<uintptr_t>(hwnd),
                            GetPrefCategoryDebugName(state->currentCategory),
                            reinterpret_cast<uintptr_t>(GetFocus()),
                            reinterpret_cast<uintptr_t>(state->pageHostWindow));
                break;
            default: break;
        }
    }

    switch (msg)
    {
        case WM_NCHITTEST:
        {
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_NCCALCSIZE:
        case WM_NCPAINT:
        {
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_NCLBUTTONDOWN:
        {
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_NCLBUTTONUP:
        case WM_NCLBUTTONDBLCLK:
        case WM_NCMOUSEMOVE:
        {
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_ERASEBKGND: return 1;
        case WM_SETFOCUS:
        {
            const bool forward = (GetKeyState(VK_SHIFT) & 0x8000) == 0;
            if (HWND target = FindFirstOrLastTabStopChild(hwnd, forward))
            {
                SetFocus(target);
                return 0;
            }

            // Note-only pages share the page host but do not expose tab-stop children.
            // When the host is reached through dialog tab traversal, skip the inert host
            // and continue to the next dialog-level focus target.
            if (HWND dlg = GetParent(hwnd))
            {
                const BOOL previous = forward ? FALSE : TRUE;
                if (HWND target = GetNextDlgTabItem(dlg, hwnd, previous); target && target != hwnd)
                {
                    SetFocus(target);
                    return 0;
                }
            }
            break;
        }
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (! hdc)
            {
                return 0;
            }

            PaintPageHostBackgroundAndCards(hdc.get(), hwnd, *state);

            return 0;
        }
        case WM_PRINTCLIENT:
        {
            HDC hdc = reinterpret_cast<HDC>(wp);
            if (! hdc)
            {
                break;
            }

            PaintPageHostBackgroundAndCards(hdc, hwnd, *state);
            return 0;
        }
        case WndMsg::kPreferencesSelectPluginDetails:
        {
            const PrefsPluginType pluginType = static_cast<PrefsPluginType>(wp);
            const size_t pluginIndex         = static_cast<size_t>(lp);
            const PrefsPluginListItem pluginItem{pluginType, pluginIndex};
            const std::optional<PrefsPluginListItem> resolved = PrefsPlugins::FindItemById(PrefsPlugins::GetId(pluginItem));
            const PrefsPluginListItem effectiveItem           = resolved.has_value() ? resolved.value() : pluginItem;
            const std::wstring_view pluginId                  = PrefsPlugins::GetId(effectiveItem);
            if (pluginId.empty())
            {
                return 0;
            }

            state->initialCategory       = PrefCategory::Plugins;
            state->currentCategory       = PrefCategory::Plugins;
            state->pluginsSelectedPlugin = effectiveItem;
            state->pluginsSelectedPluginId.assign(pluginId);
            state->pluginsRetainedSelectedPluginId.assign(pluginId);
            state->pluginsDetailsActive = true;

            auto& dialogHostState = static_cast<PreferencesDialogHost&>(*state);
            if (dialogHostState._categoryTreeControl)
            {
                dialogHostState._categoryTreeControl->SetSelectedItemId(
                    PreferencesDialogHost::PreferencesCategoryTreeModel::EncodePluginNodeId(effectiveItem.type, effectiveItem.index));
                dialogHostState._categoryTreeControl->NotifyDataChanged();
            }

            UpdatePageText(hwnd, *state, PrefCategory::Plugins);
            return 0;
        }
        case WM_CTLCOLORSTATIC:
        {
            HDC hdc      = reinterpret_cast<HDC>(wp);
            HWND control = reinterpret_cast<HWND>(lp);
            if (! hdc)
            {
                break;
            }

            const bool visuallyDisabled = control && GetPropW(control, kPrefsVisuallyDisabledProp) != nullptr;
            bool enabled                = true;
            if (control)
            {
                enabled = IsWindowEnabled(control) != FALSE;

                // Combo box selection fields sometimes paint via a child static control; match the input background.
                if (const HWND parent = GetParent(control))
                {
                    std::array<wchar_t, 32> className{};
                    const int len = GetClassNameW(parent, className.data(), static_cast<int>(className.size()));
                    if (len > 0 && _wcsicmp(className.data(), L"ComboBox") == 0)
                    {
                        const bool comboEnabled   = IsWindowEnabled(parent) != FALSE;
                        const bool focused        = comboEnabled && (GetFocus() == parent || SendMessageW(parent, CB_GETDROPPEDSTATE, 0, 0) != 0);
                        const bool themedInputs   = state->inputBrush.get() != nullptr;
                        const COLORREF background = themedInputs ? (comboEnabled ? (focused ? state->inputFocusedBackgroundColor : state->inputBackgroundColor)
                                                                                 : state->inputDisabledBackgroundColor)
                                                                 : state->theme.windowBackground;
                        HBRUSH brush              = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));
                        if (themedInputs)
                        {
                            if (! comboEnabled)
                            {
                                brush = state->inputDisabledBrush ? state->inputDisabledBrush.get() : state->inputBrush.get();
                            }
                            else if (focused && state->inputFocusedBrush)
                            {
                                brush = state->inputFocusedBrush.get();
                            }
                            else
                            {
                                brush = state->inputBrush.get();
                            }
                        }

                        COLORREF textColor = (comboEnabled && ! visuallyDisabled) ? state->theme.menu.text : GetDisabledTextColor(*state, background);
                        if (comboEnabled && ! visuallyDisabled && ! state->theme.highContrast && ! state->theme.systemHighContrast)
                        {
                            constexpr int kMinTextLumaDiff = 80;
                            if (std::abs(ColorLuma(textColor) - ColorLuma(background)) < kMinTextLumaDiff)
                            {
                                textColor = ChooseContrastingTextColor(background);
                            }
                        }
                        SetBkMode(hdc, OPAQUE);
                        SetBkColor(hdc, background);
                        SetTextColor(hdc, textColor);
                        if (! state->backgroundBrush)
                        {
                            SetDCBrushColor(hdc, background);
                        }
                        return reinterpret_cast<LRESULT>(brush);
                    }
                }
            }

            const COLORREF windowBackground = state->theme.windowBackground;
            COLORREF background             = windowBackground;
            HBRUSH brush                    = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));

            if (! state->theme.systemHighContrast && state->cardBrush && ! state->pageSettingCards.empty() && control)
            {
                RECT rcControl{};
                if (GetWindowRect(control, &rcControl))
                {
                    MapWindowPoints(nullptr, hwnd, reinterpret_cast<POINT*>(&rcControl), 2);
                    POINT center{};
                    center.x = (rcControl.left + rcControl.right) / 2;
                    center.y = (rcControl.top + rcControl.bottom) / 2;

                    for (const RECT& baseCard : state->pageSettingCards)
                    {
                        RECT card = baseCard;
                        OffsetRect(&card, 0, -state->pageScrollY);
                        if (PtInRect(&card, center) != FALSE)
                        {
                            background = state->cardBackgroundColor;
                            brush      = state->cardBrush.get();
                            break;
                        }
                    }
                }
            }

            const bool enabledForText = enabled && ! visuallyDisabled;
            COLORREF textColor        = enabledForText ? state->theme.menu.text : GetDisabledTextColor(*state, background);
            if (enabledForText && ! state->theme.highContrast && ! state->theme.systemHighContrast)
            {
                constexpr int kMinTextLumaDiff = 80;
                if (std::abs(ColorLuma(textColor) - ColorLuma(background)) < kMinTextLumaDiff)
                {
                    textColor = ChooseContrastingTextColor(background);
                }
            }
            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, background);
            SetTextColor(hdc, textColor);
            if (! state->backgroundBrush)
            {
                SetDCBrushColor(hdc, background);
            }
            return reinterpret_cast<LRESULT>(brush);
        }
        case WM_CTLCOLOREDIT:
        {
            HDC hdc = reinterpret_cast<HDC>(wp);
            if (! hdc)
            {
                break;
            }

            const HWND control      = reinterpret_cast<HWND>(lp);
            const bool enabled      = ! control || IsWindowEnabled(control) != FALSE;
            const bool focused      = enabled && control && GetFocus() == control;
            const bool themedInputs = state->inputBrush.get() != nullptr;
            const COLORREF background =
                themedInputs ? (enabled ? (focused ? state->inputFocusedBackgroundColor : state->inputBackgroundColor) : state->inputDisabledBackgroundColor)
                             : state->theme.windowBackground;
            COLORREF textColor = enabled ? state->theme.menu.text : GetDisabledTextColor(*state, background);
            if (enabled && ! state->theme.highContrast && ! state->theme.systemHighContrast)
            {
                constexpr int kMinTextLumaDiff = 80;
                if (std::abs(ColorLuma(textColor) - ColorLuma(background)) < kMinTextLumaDiff)
                {
                    textColor = ChooseContrastingTextColor(background);
                }
            }
            HBRUSH brush = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));
            if (themedInputs)
            {
                if (! enabled)
                {
                    brush = state->inputDisabledBrush ? state->inputDisabledBrush.get() : state->inputBrush.get();
                }
                else if (focused && state->inputFocusedBrush)
                {
                    brush = state->inputFocusedBrush.get();
                }
                else
                {
                    brush = state->inputBrush.get();
                }
            }
            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, background);
            SetTextColor(hdc, textColor);
            if (! state->backgroundBrush)
            {
                SetDCBrushColor(hdc, background);
            }
            return reinterpret_cast<LRESULT>(brush);
        }
        case WM_CTLCOLORBTN:
        {
            HDC hdc      = reinterpret_cast<HDC>(wp);
            HWND control = reinterpret_cast<HWND>(lp);
            if (! hdc)
            {
                break;
            }

            const COLORREF windowBackground = state->theme.windowBackground;
            COLORREF background             = windowBackground;
            HBRUSH brush                    = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));

            if (! state->theme.systemHighContrast && state->cardBrush && ! state->pageSettingCards.empty() && control)
            {
                RECT rcControl{};
                if (GetWindowRect(control, &rcControl))
                {
                    MapWindowPoints(nullptr, hwnd, reinterpret_cast<POINT*>(&rcControl), 2);
                    POINT center{};
                    center.x = (rcControl.left + rcControl.right) / 2;
                    center.y = (rcControl.top + rcControl.bottom) / 2;

                    for (const RECT& baseCard : state->pageSettingCards)
                    {
                        RECT card = baseCard;
                        OffsetRect(&card, 0, -state->pageScrollY);
                        if (PtInRect(&card, center) != FALSE)
                        {
                            background = state->cardBackgroundColor;
                            brush      = state->cardBrush.get();
                            break;
                        }
                    }
                }
            }

            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, background);
            SetTextColor(hdc, state->theme.menu.text);
            if (! state->backgroundBrush)
            {
                SetDCBrushColor(hdc, background);
            }
            return reinterpret_cast<LRESULT>(brush);
        }
        case WM_CTLCOLORLISTBOX:
        {
            HDC hdc = reinterpret_cast<HDC>(wp);
            if (! hdc)
            {
                break;
            }

            const HWND control      = reinterpret_cast<HWND>(lp);
            const bool enabled      = ! control || IsWindowEnabled(control) != FALSE;
            const bool themedInputs = state->inputBrush.get() != nullptr;
            const COLORREF background =
                themedInputs ? (enabled ? state->inputBackgroundColor : state->inputDisabledBackgroundColor) : state->theme.windowBackground;
            const COLORREF textColor = enabled ? state->theme.menu.text : GetDisabledTextColor(*state, background);
            HBRUSH brush             = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));
            if (themedInputs)
            {
                brush = enabled ? state->inputBrush.get() : (state->inputDisabledBrush ? state->inputDisabledBrush.get() : state->inputBrush.get());
            }
            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, background);
            SetTextColor(hdc, textColor);
            if (! state->backgroundBrush)
            {
                SetDCBrushColor(hdc, background);
            }
            return reinterpret_cast<LRESULT>(brush);
        }
        case WM_MEASUREITEM: break;
        case WM_DRAWITEM:
        {
            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis)
            {
                break;
            }

            {
                const LRESULT handled = ThemesPane::OnDrawColorSwatch(dis, *state);
                if (handled != 0)
                {
                    return handled;
                }
            }

            break;
        }
        case WndMsg::kPreferencesDeferredPaneAction:
            if (auto payload = TakeMessagePayload<PreferencesDeferredActionPayload>(lp))
            {
                return HandleDeferredPaneAction(hwnd, *state, std::move(*payload)) ? 0 : FALSE;
            }
            return FALSE;
        case WndMsg::kPreferencesApplyPageHostScroll:
            PrefsPageHost::FlushPendingScroll(hwnd, *state);
            return 0;
        case WM_COMMAND:
        {
            const UINT notify = HIWORD(wp);
            if (notify == BN_SETFOCUS || notify == EN_SETFOCUS || notify == CBN_SETFOCUS)
            {
                HWND control = reinterpret_cast<HWND>(lp);
                if (control)
                {
                    PrefsPageHost::EnsureControlVisible(hwnd, *state, control);
                    InvalidateRect(control, nullptr, TRUE);
                }
            }
            if (notify == BN_KILLFOCUS || notify == EN_KILLFOCUS || notify == CBN_KILLFOCUS)
            {
                HWND control = reinterpret_cast<HWND>(lp);
                if (control)
                {
                    InvalidateRect(control, nullptr, TRUE);
                }
            }

            break;
        }
        case WM_VSCROLL:
        {
            if (! state || state->pageScrollMaxY <= 0)
            {
                break;
            }

            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_ALL;
            if (! GetScrollInfo(hwnd, SB_VERT, &si))
            {
                break;
            }

            int newPos         = state->pageScrollY;
            const UINT dpi     = GetDpiForWindow(hwnd);
            const int lineStep = std::max(1, UiMetrics::ScaleDip(dpi, 24));
            const UINT scrollCode = LOWORD(wp);

            switch (scrollCode)
            {
                case SB_LINEUP: newPos -= lineStep; break;
                case SB_LINEDOWN: newPos += lineStep; break;
                case SB_PAGEUP: newPos -= static_cast<int>(si.nPage); break;
                case SB_PAGEDOWN: newPos += static_cast<int>(si.nPage); break;
                case SB_TOP: newPos = 0; break;
                case SB_BOTTOM: newPos = state->pageScrollMaxY; break;
                case SB_THUMBPOSITION:
                case SB_THUMBTRACK:
                {
                    newPos = si.nTrackPos;
                    const int packedTrackPos = static_cast<int>(HIWORD(wp));
                    if (packedTrackPos != 0 && newPos == si.nPos)
                    {
                        newPos = packedTrackPos;
                    }
                    break;
                }
                case SB_ENDSCROLL:
                    PrefsPageHost::FlushPendingScroll(hwnd, *state);
                    return 0;
                default: break;
            }

            PrefsPageHost::ScrollTo(hwnd, *state, newPos);
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            if (! state)
            {
                break;
            }

            if (HandlePageHostMouseWheel(hwnd, *state, wp))
            {
                return 0;
            }
            break;
        }
        case WM_SIZE:
        {
            if (state->pageHostIgnoreSize)
            {
                return DefWindowProcW(hwnd, msg, wp, lp);
            }

            // Suppress intermediate repaints while repositioning child
            // windows during layout.  This is safe for DxUI children
            // because the WindowHost no longer discards its swap chain
            // on WM_SHOWWINDOW FALSE (triggered by WM_SETREDRAW FALSE).
            ScopedWindowRedrawBlock pageHostRedraw(hwnd);
            const LRESULT result = DefWindowProcW(hwnd, msg, wp, lp);
            if (state->pageHostUsesDxUi)
            {
                bool handled = false;
                static_cast<void>(hostState._pageHostHost.HandleMessage(hwnd, msg, wp, lp, handled));
            }
            LayoutPreferencesPageHost(hwnd, *state);
            if (state->pageHostUsesDxUi)
            {
                RECT rc{};
                GetClientRect(hwnd, &rc);
                bool handled = false;
                static_cast<void>(hostState._pageHostHost.HandleMessage(
                    hwnd, WM_SIZE, SIZE_RESTORED, MAKELPARAM(std::max(0l, rc.right - rc.left), std::max(0l, rc.bottom - rc.top)), handled));
            }
            LogPreferencesPageHostState(hwnd, *state, L"pagehost-wm-size");
            pageHostRedraw.EnableAndRedraw(RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
            return result;
        }
        case WM_NCDESTROY:
        {
            static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
            if (state->pageHostUsesDxUi)
            {
                hostState._pageHostHost.Detach();
                state->pageHostUsesDxUi               = false;
                state->pageHostDxHost                 = nullptr;
                state->pageHostDxRootControl          = nullptr;
                state->pageHostDxScrollPanelControl   = nullptr;
                state->pageHostDxContentRootControl   = nullptr;
                state->pageHostDxNoteControl          = nullptr;
            }
            RemovePropW(hwnd, kPrefsDxDiagnosticsProp);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            break;
        }
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

INT_PTR OnInitDialog(HWND dlg, PreferencesDialogState* state)
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    SetState(dlg, state);
    InitPostedPayloadWindow(dlg);
    SettingsHotReload::RegisterParticipant(dlg);

    SetWindowTextW(dlg, LoadStringResource(nullptr, IDS_PREFS_CAPTION).c_str());
    if (HWND ok = GetDlgItem(dlg, IDOK))
    {
        SetWindowTextW(ok, LoadStringResource(nullptr, IDS_BTN_OK).c_str());
    }
    if (HWND cancel = GetDlgItem(dlg, IDCANCEL))
    {
        SetWindowTextW(cancel, LoadStringResource(nullptr, IDS_BTN_CANCEL).c_str());
    }
    if (HWND apply = GetDlgItem(dlg, IDC_PREFS_APPLY))
    {
        SetWindowTextW(apply, LoadStringResource(nullptr, IDS_BTN_APPLY).c_str());
    }

    ApplyWindowChromeTheme(dlg, state->theme, WindowBackdropTarget::Tool, GetActiveWindow() == dlg);

    state->backgroundBrush.reset(CreateSolidBrush(state->theme.windowBackground));
    state->cardBackgroundColor = UiMetrics::GetControlSurfaceColor(state->theme);

    state->inputBackgroundColor         = UiMetrics::BlendColor(state->cardBackgroundColor, state->theme.windowBackground, state->theme.dark ? 50 : 30, 255);
    state->inputFocusedBackgroundColor  = UiMetrics::BlendColor(state->inputBackgroundColor, state->theme.menu.text, state->theme.dark ? 20 : 16, 255);
    state->inputDisabledBackgroundColor = UiMetrics::BlendColor(state->theme.windowBackground, state->inputBackgroundColor, state->theme.dark ? 70 : 40, 255);
    state->cardBrush.reset();
    state->inputBrush.reset();
    state->inputFocusedBrush.reset();
    state->inputDisabledBrush.reset();
    if (! state->theme.systemHighContrast)
    {
        state->cardBrush.reset(CreateSolidBrush(state->cardBackgroundColor));
        state->inputBrush.reset(CreateSolidBrush(state->inputBackgroundColor));
        state->inputFocusedBrush.reset(CreateSolidBrush(state->inputFocusedBackgroundColor));
        state->inputDisabledBrush.reset(CreateSolidBrush(state->inputDisabledBackgroundColor));
    }

    RECT initial{};
    if (GetWindowRect(dlg, &initial))
    {
        state->minTrackSizePx.cx = std::max(0l, initial.right - initial.left);
        state->minTrackSizePx.cy = std::max(0l, initial.bottom - initial.top);

        HWND ok     = GetDlgItem(dlg, IDOK);
        HWND cancel = GetDlgItem(dlg, IDCANCEL);
        HWND apply  = GetDlgItem(dlg, IDC_PREFS_APPLY);

        RECT client{};
        GetClientRect(dlg, &client);
        const int windowWidth     = std::max(0, static_cast<int>(state->minTrackSizePx.cx));
        const int clientWidth     = std::max(0l, client.right - client.left);
        const int windowHeight    = std::max(0, static_cast<int>(state->minTrackSizePx.cy));
        const int clientHeight    = std::max(0l, client.bottom - client.top);
        const int nonClientWidth  = std::max(0, windowWidth - clientWidth);
        const int nonClientHeight = std::max(0, windowHeight - clientHeight);

        if (ok && cancel && apply)
        {
            const UINT dpi                                     = GetDpiForWindow(dlg);
            const int margin                                   = UiMetrics::ScaleDip(dpi, 8);
            const int gapX                                     = UiMetrics::ScaleDip(dpi, 8);
            const int minGapX                                  = UiMetrics::ScaleDip(dpi, 4);
            const int buttonPadX                               = UiMetrics::ScaleDip(dpi, 12);
            const PreferencesTypographyContext shellTypography = PrefsUi::MakeTypographyContext(dlg);

            const auto measureButtonMinWidth = [&](HWND button) noexcept
            {
                const std::wstring text = PrefsUi::GetWindowTextString(button);
                const int textW         = PrefsUi::MeasureSingleLineTextWidthPx(shellTypography, shellTypography.body, text);
                return std::max(UiMetrics::ScaleDip(dpi, 60), textW + 2 * buttonPadX);
            };

            const int okMin     = measureButtonMinWidth(ok);
            const int cancelMin = measureButtonMinWidth(cancel);
            const int applyMin  = measureButtonMinWidth(apply);

            const int minButtonsClientWidth = std::max(0, (2 * margin) + okMin + cancelMin + applyMin + (2 * minGapX));

            const int listMinWidth          = UiMetrics::ScaleDip(dpi, 72);
            const int hostMinWidth          = UiMetrics::ScaleDip(dpi, 140);
            const int minContentClientWidth = std::max(0, (2 * margin) + listMinWidth + gapX + hostMinWidth);

            const int minClientWidth = std::max(minButtonsClientWidth, minContentClientWidth);
            state->minTrackSizePx.cx = std::max(0, minClientWidth + nonClientWidth);

            RECT okRect{};
            RECT cancelRect{};
            RECT applyRect{};
            GetWindowRect(ok, &okRect);
            GetWindowRect(cancel, &cancelRect);
            GetWindowRect(apply, &applyRect);

            const int okHeight     = std::max(0l, okRect.bottom - okRect.top);
            const int cancelHeight = std::max(0l, cancelRect.bottom - cancelRect.top);
            const int applyHeight  = std::max(0l, applyRect.bottom - applyRect.top);
            int buttonHeight       = std::max({okHeight, cancelHeight, applyHeight});
            if (buttonHeight <= 0)
            {
                buttonHeight = UiMetrics::ScaleDip(dpi, 26);
            }

            // Content area = left list + page host (scrolls vertically). Keep the minimum height small enough
            // to allow the user to shrink the dialog while still keeping the buttons reachable.
            const int minContentClientHeight = UiMetrics::ScaleDip(dpi, 160);
            const int minClientHeight        = std::max(0, minContentClientHeight + buttonHeight + 3 * margin);
            state->minTrackSizePx.cy         = std::max(0, minClientHeight + nonClientHeight);
        }
    }

    PopulateCategoryTree(dlg, *state);

    UpdateApplyButton(dlg, *state);

    CreatePageControls(dlg, *state);
    ApplyThemeToPreferencesDialog(dlg, *state, state->theme);

    // DxUi shell chrome (title label + button bar) needs more vertical space than the
    // original Win32 statics. Grow the dialog once to accommodate and update min size.
    {
        auto& hostInit = static_cast<PreferencesDialogHost&>(*state);
        if (hostInit._shellHostHwnd)
        {
            const UINT initDpi      = GetDpiForWindow(dlg);
            const int chromeExtraPx = UiMetrics::ScaleDip(initDpi, 50);
            RECT wr{};
            if (GetWindowRect(dlg, &wr))
            {
                const int newH = std::max(0l, wr.bottom - wr.top) + chromeExtraPx;
                SetWindowPos(dlg, nullptr, 0, 0, wr.right - wr.left, newH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                state->minTrackSizePx.cy += chromeExtraPx;
            }
        }
    }

    if (state->pageHostWindow)
    {
        SetWindowLongPtrW(state->pageHostWindow, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
    }

    InstallWheelRoutingHooks(dlg);

    SelectCategory(dlg, *state, state->initialCategory);
    return TRUE;
}

INT_PTR OnCtlColorDialog(PreferencesDialogState* state)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCtlColorStatic(PreferencesDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    COLORREF textColor = state->theme.menu.text;
    if (control)
    {
        if (IsWindowEnabled(control) == FALSE)
        {
            textColor = GetDisabledTextColor(*state, state->theme.windowBackground);
        }
    }

    if (! state->theme.systemHighContrast)
    {
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, textColor);
        return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
    }

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, state->theme.windowBackground);
    SetTextColor(hdc, textColor);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCtlColorListBox(PreferencesDialogState* state, HDC hdc, HWND listBox)
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    const bool isCategoryTree = listBox && state->categoryTreeWindow && listBox == state->categoryTreeWindow;
    const bool useInputBrush  = ! isCategoryTree && state->inputBrush && ! state->theme.systemHighContrast;

    const COLORREF background = useInputBrush ? state->inputBackgroundColor : state->theme.windowBackground;
    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, background);
    SetTextColor(hdc, state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(useInputBrush ? state->inputBrush.get() : state->backgroundBrush.get());
}

INT_PTR OnCommand(HWND dlg, PreferencesDialogState* state, UINT commandId, [[maybe_unused]] UINT notifyCode, [[maybe_unused]] HWND hwndCtl)
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    switch (commandId)
    {
        case IDOK:
            if (state->dirty)
            {
                CommitAndApply(dlg, *state);
                if (state->dirty)
                {
                    return TRUE;
                }
            }
            return RequestPreferencesDialogClose(dlg) ? TRUE : FALSE;
        case IDC_PREFS_APPLY:
            if (state->dirty)
            {
                CommitAndApply(dlg, *state);
            }
            return TRUE;
        case IDC_PREFS_RESET_ALL: ResetAllPreferencesToDefaults(dlg, *state); return TRUE;
        case IDCANCEL: return RequestPreferencesDialogCancelClose(dlg, *state) ? TRUE : FALSE;
    }

    return FALSE;
}

bool HandleDeferredPaneAction(HWND dlg, PreferencesDialogState& state, PreferencesDeferredActionPayload&& payload) noexcept
{
    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    const HWND host = (state.pageHostWindow && IsWindow(state.pageHostWindow) != FALSE) ? state.pageHostWindow : dlg;

    switch (payload.kind)
    {
        case PreferencesDeferredActionKind::ViewersSearchChanged: return hostState._viewersPane.HandleDeferredAction(host, state, payload.kind);
        case PreferencesDeferredActionKind::KeyboardSearchChanged:
        case PreferencesDeferredActionKind::KeyboardScopeChanged:
        case PreferencesDeferredActionKind::KeyboardAssign:
        case PreferencesDeferredActionKind::KeyboardRemove:
        case PreferencesDeferredActionKind::KeyboardReset:
        case PreferencesDeferredActionKind::KeyboardImport:
        case PreferencesDeferredActionKind::KeyboardExport: return hostState._keyboardPane.HandleDeferredAction(host, state, payload.kind);
        case PreferencesDeferredActionKind::ThemesThemeChanged:
        case PreferencesDeferredActionKind::ThemesBaseChanged:
        case PreferencesDeferredActionKind::ThemesNameBlur:
        case PreferencesDeferredActionKind::ThemesSearchChanged: return hostState._themesPane.HandleDeferredAction(host, state, payload.kind);
        case PreferencesDeferredActionKind::PluginsSearchChanged:
        case PreferencesDeferredActionKind::PluginsConfigure:
        case PreferencesDeferredActionKind::PluginsTest:
        case PreferencesDeferredActionKind::PluginsTestAll:
        {
            const bool handled = hostState._pluginsPane.HandleDeferredAction(host, state, payload.kind);
            if (handled && host)
            {
                LayoutPreferencesPageHost(host, state);
                InvalidateRect(host, nullptr, FALSE);
            }
            return handled;
        }
        case PreferencesDeferredActionKind::FileOperationsBandwidthPresetChanged:
        {
            const bool handled = hostState._fileOperationsPane.HandleDeferredAction(host, state, payload.kind);
            if (handled && host)
            {
                LayoutPreferencesPageHost(host, state);
                InvalidateRect(host, nullptr, FALSE);
                hostState._fileOperationsPane.PostLayoutFocusCustomBandwidthEdit();
            }
            return handled;
        }
        case PreferencesDeferredActionKind::CompareDirectoriesIgnoreToggleChanged:
        {
            const bool handled = hostState._compareDirectoriesPane.HandleDeferredAction(host, state, payload.kind);
            if (handled && host)
            {
                LayoutPreferencesPageHost(host, state);
                InvalidateRect(host, nullptr, FALSE);
            }
            return handled;
        }
        default: return false;
    }
}

INT_PTR CALLBACK PreferencesDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* state = GetState(dlg);

    switch (msg)
    {
        case WM_INITDIALOG: return OnInitDialog(dlg, reinterpret_cast<PreferencesDialogState*>(lp));
        case WM_CLOSE: return OnCommand(dlg, state, IDCANCEL, 0, nullptr);
        case kPrefsDeferredCloseMessage:
            if (g_preferencesDialog.get() == dlg)
            {
                const HWND restoreOwner = (state && state->owner && IsWindow(state->owner) != FALSE) ? state->owner : nullptr;
                if (restoreOwner)
                {
                    static_cast<void>(PostMessageW(restoreOwner, WndMsg::kPaneRestoreFolderFocus, 0, 0));
                }
                g_preferencesDialog.reset();
                return TRUE;
            }
            return FALSE;
        case WndMsg::kPreferencesDeferredPaneAction:
            if (auto payload = TakeMessagePayload<PreferencesDeferredActionPayload>(lp))
            {
                return state ? (HandleDeferredPaneAction(dlg, *state, std::move(*payload)) ? TRUE : FALSE) : FALSE;
            }
            return FALSE;
        case WM_ERASEBKGND:
            if (state && state->backgroundBrush && wp)
            {
                RECT rc{};
                if (GetClientRect(dlg, &rc))
                {
                    FillRect(reinterpret_cast<HDC>(wp), &rc, state->backgroundBrush.get());
                    return TRUE;
                }
            }
            break;
        case WM_CTLCOLORDLG: return OnCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLORLISTBOX: return OnCtlColorListBox(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_ACTIVATE:
            if (state)
            {
                if (state->categoryTreeWindow)
                {
                    InvalidateRect(state->categoryTreeWindow, nullptr, FALSE);
                }
                if (state->pageHostWindow)
                {
                    RedrawWindow(state->pageHostWindow, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
                }
            }
            return FALSE;
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->theme, wp != FALSE);
            }
            return FALSE;
        case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lp);
            if (! info)
            {
                break;
            }

            bool handled = false;
            if (state && state->minTrackSizePx.cx > 0 && state->minTrackSizePx.cy > 0)
            {
                info->ptMinTrackSize.x = state->minTrackSizePx.cx;
                info->ptMinTrackSize.y = state->minTrackSizePx.cy;
                handled                = true;
            }

            // Custom "maximize vertically": keep the current width, but expand to the monitor work-area height.
            MONITORINFO mi{};
            mi.cbSize              = sizeof(mi);
            const HMONITOR monitor = MonitorFromWindow(dlg, MONITOR_DEFAULTTONEAREST);
            if (monitor && GetMonitorInfoW(monitor, &mi))
            {
                RECT windowRc{};
                if (GetWindowRect(dlg, &windowRc))
                {
                    const int workWidth    = std::max(0l, mi.rcWork.right - mi.rcWork.left);
                    const int workHeight   = std::max(0l, mi.rcWork.bottom - mi.rcWork.top);
                    const int currentWidth = std::max(0l, windowRc.right - windowRc.left);
                    const int desiredWidth = std::clamp(currentWidth, 0, workWidth);
                    const int maxLeft      = static_cast<int>(mi.rcWork.right) - desiredWidth;
                    const int desiredLeft  = std::clamp(static_cast<int>(windowRc.left), static_cast<int>(mi.rcWork.left), maxLeft);

                    info->ptMaxSize.x     = desiredWidth;
                    info->ptMaxSize.y     = workHeight;
                    info->ptMaxPosition.x = desiredLeft - mi.rcMonitor.left;
                    info->ptMaxPosition.y = mi.rcWork.top - mi.rcMonitor.top;
                    handled               = true;
                }
            }

            if (handled)
            {
                return TRUE;
            }
            break;
        }
        case WM_DPICHANGED:
            if (state)
            {
                const RECT* const suggested = reinterpret_cast<const RECT*>(lp);
                if (suggested)
                {
                    const int width  = static_cast<int>(std::max(0l, suggested->right - suggested->left));
                    const int height = static_cast<int>(std::max(0l, suggested->bottom - suggested->top));
                    SetWindowPos(dlg, nullptr, suggested->left, suggested->top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
                }

                LayoutPreferencesDialog(dlg, *state);
                if (state->pageHostWindow)
                {
                    LayoutPreferencesPageHost(state->pageHostWindow, *state);
                    SyncPreferencesPageHostDxSize(state->pageHostWindow, *state, L"dialog-dpi-changed");
                    LogPreferencesPageHostState(state->pageHostWindow, *state, L"dialog-dpi-changed");
                }
                RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
            }
            return TRUE;
        case WM_SIZE:
            if (state)
            {
                LayoutPreferencesDialog(dlg, *state);
                if (state->pageHostWindow)
                {
                    LayoutPreferencesPageHost(state->pageHostWindow, *state);
                    SyncPreferencesPageHostDxSize(state->pageHostWindow, *state, L"dialog-wm-size");
                    LogPreferencesPageHostState(state->pageHostWindow, *state, L"dialog-wm-size");
                }
                InvalidateRect(dlg, nullptr, TRUE);
            }
            return TRUE;
        case WndMsg::kSettingsReloadedFromDisk:
            if (state)
            {
                return OnSettingsReloadedFromDisk(dlg, *state);
            }
            return TRUE;
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED:
        case WM_DWMCOLORIZATIONCOLORCHANGED:
        case WM_SYSCOLORCHANGE:
            if (state)
            {
                RefreshPreferencesDialogThemeImpl(dlg, *state);
            }
            return TRUE;
        case WM_EXITSIZEMOVE:
            if (state)
            {
                LayoutPreferencesDialog(dlg, *state);
                if (state->pageHostWindow)
                {
                    LayoutPreferencesPageHost(state->pageHostWindow, *state);
                    SyncPreferencesPageHostDxSize(state->pageHostWindow, *state, L"dialog-exitsizemove");
                    LogPreferencesPageHostState(state->pageHostWindow, *state, L"dialog-exitsizemove");
                }
                RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            }
            return TRUE;
        case WM_COMMAND: return OnCommand(dlg, state, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp));
        case WM_NCDESTROY:
        {
            if (state)
            {
                const HWND restoreOwner = (state->owner && IsWindow(state->owner) != FALSE) ? state->owner : nullptr;
                std::unique_ptr<PreferencesDialogHost> stateOwner;
                stateOwner.reset(static_cast<PreferencesDialogHost*>(state));
                SettingsHotReload::UnregisterParticipant(dlg);
                auto& hostState = static_cast<PreferencesDialogHost&>(*state);

                if (state->settings)
                {
                    WindowPlacementPersistence::Save(*state->settings, kPreferencesWindowId, dlg);

                    const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(state->appId, *state->settings);
                    if (FAILED(saveHr))
                    {
                        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(state->appId);
                        Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
                    }
                }

                if (state->pageHostWindow)
                {
                    if (state->pageHostUsesDxUi)
                    {
                        hostState._pageHostHost.Detach();
                        state->pageHostUsesDxUi               = false;
                        state->pageHostDxHost                 = nullptr;
                        state->pageHostDxRootControl          = nullptr;
                        state->pageHostDxScrollPanelControl   = nullptr;
                        state->pageHostDxContentRootControl   = nullptr;
                        state->pageHostDxNoteControl          = nullptr;
                    }
                    SetWindowLongPtrW(state->pageHostWindow, GWLP_USERDATA, 0);
                }

                if (hostState._shellHostHwnd && IsWindow(hostState._shellHostHwnd) != FALSE)
                {
                    RestoreWndProcHook(hostState._shellHostHwnd, kPrefsDxShellHostOriginalWndProcProp);
                    RemovePropW(hostState._shellHostHwnd, kPrefsDxShellHostProp);
                    RemovePropW(hostState._shellHostHwnd, kPrefsDxShellHostOriginalWndProcProp);
                }
                hostState._shellHost.Detach();
                hostState._shellHostHwnd = nullptr;

                if (state->categoryTreeUsesDxUi)
                {
                    if (state->categoryTreeWindow && IsWindow(state->categoryTreeWindow) != FALSE)
                    {
                        RestoreWndProcHook(state->categoryTreeWindow, kPrefsDxCategoryHostOriginalWndProcProp);
                        RemovePropW(state->categoryTreeWindow, kPrefsDxCategoryHostStateProp);
                        RemovePropW(state->categoryTreeWindow, kPrefsDxCategoryHostOriginalWndProcProp);
                    }
                    hostState._categoryTreeHost.Detach();
                    hostState._categoryTreeControl = nullptr;
                    state->categoryTreeUsesDxUi    = false;
                }

                RestoreWndProcHook(dlg, kPrefsWheelRouteOriginalWndProcProp);
                EnumChildWindows(dlg,
                                 [](HWND child, LPARAM) noexcept -> BOOL
                {
                    RestoreWndProcHook(child, kPrefsWheelRouteOriginalWndProcProp);
                    return TRUE;
                },
                                 0);

                SetState(dlg, nullptr);
                if (g_preferencesDialog.get() == dlg)
                {
                    g_preferencesDialog.release();
                }

                if (restoreOwner)
                {
                    static_cast<void>(SetActiveWindow(restoreOwner));
                    static_cast<void>(SetForegroundWindow(restoreOwner));
                    static_cast<void>(SetFocus(restoreOwner));
                    static_cast<void>(PostMessageW(restoreOwner, WndMsg::kPaneRestoreFolderFocus, 0, 0));
                }
            }
            static_cast<void>(DrainPostedPayloadsForWindow(dlg));
            return FALSE;
        }
    }

    return FALSE;
}
} // namespace

void RefreshPreferencesDialogTheme(HWND dlg, PreferencesDialogState& state) noexcept
{
    RefreshPreferencesDialogThemeImpl(dlg, state);
}

[[nodiscard]] bool PreferencesDialog::Show(
    HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme, PrefCategory initialCategory) noexcept
{
    if (const HWND existing = g_preferencesDialog.get())
    {
        if (! IsWindow(existing))
        {
            g_preferencesDialog.release();
        }
        else
        {
            if (IsIconic(existing))
            {
                ShowWindow(existing, SW_RESTORE);
            }
            else
            {
                ShowWindow(existing, SW_SHOW);
            }
            SetForegroundWindow(existing);
            if (auto* state = GetState(existing))
            {
                HWND effectiveOwner = owner;
                if (effectiveOwner && IsWindow(effectiveOwner))
                {
                    effectiveOwner = GetAncestor(effectiveOwner, GA_ROOT);
                }
                else
                {
                    effectiveOwner = nullptr;
                }

                state->owner = effectiveOwner;
                SelectCategory(existing, *state, initialCategory);
            }
            return true;
        }
    }

    auto statePtr = std::make_unique<PreferencesDialogHost>();
    auto* state   = statePtr.get();

    HWND effectiveOwner = owner;
    if (effectiveOwner && IsWindow(effectiveOwner))
    {
        effectiveOwner = GetAncestor(effectiveOwner, GA_ROOT);
    }
    else
    {
        effectiveOwner = nullptr;
    }

    state->owner           = effectiveOwner;
    state->settings        = &settings;
    state->appId           = std::wstring(appId);
    state->theme           = theme;
    state->initialCategory = initialCategory;

    if (! EnsurePrefsPageHostClassRegistered())
    {
        return false;
    }

    state->baselineSettings = settings;
    state->workingSettings  = settings;

    // Ensure mainMenu is explicitly set with defaults if not present
    // This prevents function bar from being reset when applying preferences
    if (! state->workingSettings.mainMenu.has_value())
    {
        state->workingSettings.mainMenu = Common::Settings::MainMenuState{};
    }

    SetDirty(nullptr, *state);

    const HWND dlg = RedSalamander::Win32Callback::CreateDialogParamResourceNoThrow(
        GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PREFERENCES), nullptr, PreferencesDialogProc, reinterpret_cast<LPARAM>(state));

    if (! dlg)
    {
        return false;
    }

    g_preferencesDialog.reset(dlg);
    static_cast<void>(statePtr.release());
    const int showCmd = WindowPlacementPersistence::Restore(settings, kPreferencesWindowId, dlg);
    static_cast<void>(ShowWindow(dlg, showCmd));
    static_cast<void>(SetForegroundWindow(dlg));
    return true;
}

HWND PreferencesDialog::GetHandle() noexcept
{
    if (const HWND dlg = g_preferencesDialog.get(); dlg && IsWindow(dlg))
    {
        return dlg;
    }
    return nullptr;
}

void UpdatePreferencesWindowsTheme(const AppTheme& theme) noexcept
{
    const HWND dlg = PreferencesDialog::GetHandle();
    if (! dlg)
    {
        return;
    }

    if (auto* state = GetState(dlg))
    {
        state->theme = theme;
        RefreshPreferencesDialogTheme(dlg, *state);
        InvalidateRect(dlg, nullptr, FALSE);
    }
}

#ifdef ENABLE_TESTS
bool PreferencesDialog::DebugGetSnapshot(::PreferencesDebugSnapshot& out) noexcept
{
    out = {};

    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    RECT dialogClient{};
    if (GetClientRect(dlg, &dialogClient))
    {
        out.dialogClientBottomPx = static_cast<int>(dialogClient.bottom);
    }

    out.categoryTreeUsesDxUiHost      = state->categoryTreeUsesDxUi;
    out.shellUsesDxUiHost             = true;
    out.pageHostUsesDxUiHost          = state->pageHostUsesDxUi;
    out.pageHostDxHostAttachedToPageHost =
        state->pageHostWindow && static_cast<PreferencesDialogHost*>(state)->_pageHostHost.GetHwnd() == state->pageHostWindow;
    out.pageHostDxContentRootUsesScrollPanel =
        dynamic_cast<const ScrollPanel*>(static_cast<const PreferencesDialogHost*>(state)->_pageHostContentRootControl) != nullptr;
    if (const auto* scrollPanel = dynamic_cast<const ScrollPanel*>(static_cast<const PreferencesDialogHost*>(state)->_pageHostContentRootControl))
    {
        out.pageHostDxInternalScrollbarEnabled = scrollPanel->IsInternalScrollbarEnabled();
        const UINT dpi = state->pageHostWindow ? GetDpiForWindow(state->pageHostWindow) : USER_DEFAULT_SCREEN_DPI;
        out.pageHostDxScrollOffsetPx = static_cast<int>(std::lround((scrollPanel->GetScrollOffset() * static_cast<float>(dpi)) / 96.0f));
        const auto dipToPx = [dpi](const float value) noexcept
        {
            return static_cast<LONG>(std::lround((value * static_cast<float>(dpi)) / 96.0f));
        };
        if (state->pageHostWindow)
        {
            RECT pageClient{};
            if (GetClientRect(state->pageHostWindow, &pageClient))
            {
                const int clientWidthPx = std::max(0l, pageClient.right - pageClient.left);
                const int thicknessPx =
                    static_cast<int>(std::lround((scrollPanel->GetScrollbarThickness() * static_cast<float>(dpi)) / 96.0f));
                out.pageHostScrollbarTrackLeftPx = std::max(0, clientWidthPx - std::max(0, thicknessPx));
            }
        }
        D2D1_RECT_F thumbDip{};
        if (scrollPanel->DebugGetScrollbarThumbHitRect(thumbDip))
        {
            out.pageHostDxScrollbarThumbHitRectPx = RECT{dipToPx(thumbDip.left), dipToPx(thumbDip.top), dipToPx(thumbDip.right), dipToPx(thumbDip.bottom)};
            const D2D1_POINT_2F centerDip = D2D1::Point2F((thumbDip.left + thumbDip.right) * 0.5f, (thumbDip.top + thumbDip.bottom) * 0.5f);
            if (const Control* hitControl = static_cast<PreferencesDialogHost*>(state)->_pageHostHost.DebugHitTestControl(centerDip))
            {
                out.pageHostDxThumbCenterHitAnyControl  = true;
                out.pageHostDxThumbCenterHitScrollPanel = hitControl == scrollPanel;
            }
        }
    }
    out.categoryTreeFocused           = state->categoryTreeWindow && GetFocus() == state->categoryTreeWindow;
    out.pluginItemSelected            = state->pluginsSelectedPlugin.has_value();
    out.pluginsDetailsActive          = state->pluginsDetailsActive;
    out.currentCategory               = state->currentCategory;
    out.visibleChildWindowCount       = CountVisibleChildWindows(dlg);
    out.visibleLegacyTreeViewCount    = CountVisibleChildWindowsByClass(dlg, L"SysTreeView32");
    out.currentPageCardCount          = state->pageSettingCards.size();
    for (const RECT& card : state->pageSettingCards)
    {
        out.pageHostRightmostCardRightPx = std::max(out.pageHostRightmostCardRightPx, static_cast<int>(card.right));
    }
    if (out.pageHostScrollbarTrackLeftPx > 0 && out.pageHostRightmostCardRightPx > 0)
    {
        out.pageHostCardToScrollbarGapPx = out.pageHostScrollbarTrackLeftPx - out.pageHostRightmostCardRightPx;
    }
    out.visibleLegacyShellStaticCount = 0u;
    out.visibleLegacyFooterButtonCount =
        ((GetDlgItem(dlg, IDOK) && PrefsUi::IsActuallyVisibleChildWindow(GetDlgItem(dlg, IDOK))) ? 1u : 0u) +
        ((GetDlgItem(dlg, IDCANCEL) && PrefsUi::IsActuallyVisibleChildWindow(GetDlgItem(dlg, IDCANCEL))) ? 1u : 0u) +
        ((GetDlgItem(dlg, IDC_PREFS_APPLY) && PrefsUi::IsActuallyVisibleChildWindow(GetDlgItem(dlg, IDC_PREFS_APPLY))) ? 1u : 0u);
    if (state->pageHostWindow)
    {
        out.pageHostShowsVerticalScroll = state->pageScrollMaxY > 0;
        out.pageHostHasNativeVerticalScroll = (GetWindowLongPtrW(state->pageHostWindow, GWL_STYLE) & WS_VSCROLL) != 0;
        RECT pageHostRect{};
        if (GetWindowRect(state->pageHostWindow, &pageHostRect))
        {
            out.pageHostTopPx = static_cast<int>(pageHostRect.top);
        }
    }
    if (state->categoryTreeWindow)
    {
        RECT categoryTreeRect{};
        if (GetWindowRect(state->categoryTreeWindow, &categoryTreeRect))
        {
            out.categoryTreeTopPx = static_cast<int>(categoryTreeRect.top);
            POINT categoryClientCorners[2]{{categoryTreeRect.left, categoryTreeRect.top}, {categoryTreeRect.right, categoryTreeRect.bottom}};
            MapWindowPoints(nullptr, dlg, categoryClientCorners, 2);
            out.categoryTreeTopClientPx    = categoryClientCorners[0].y;
            out.categoryTreeBottomClientPx = categoryClientCorners[1].y;
            if (out.dialogClientBottomPx > 0)
            {
                out.categoryTreeBottomGapPx = out.dialogClientBottomPx - out.categoryTreeBottomClientPx;
            }
        }
    }
    if (auto& hostState = static_cast<PreferencesDialogHost&>(*state); hostState._shellHostHwnd)
    {
        RECT shellHostRect{};
        if (GetWindowRect(hostState._shellHostHwnd, &shellHostRect))
        {
            out.shellHostTopPx = static_cast<int>(shellHostRect.top);
        }
        RECT shellClient{};
        if (GetClientRect(hostState._shellHostHwnd, &shellClient))
        {
            out.shellHostClientWidthPx  = std::max(0l, shellClient.right - shellClient.left);
            out.shellHostClientHeightPx = std::max(0l, shellClient.bottom - shellClient.top);
        }

        const auto controlBoundsToPx = [&](const Control* control) noexcept
        {
            if (! control)
            {
                return RECT{};
            }

            const D2D1_RECT_F bounds = control->GetBounds();
            return RECT{static_cast<LONG>(std::lround(hostState._shellHost.DipsToPixels(bounds.left))),
                        static_cast<LONG>(std::lround(hostState._shellHost.DipsToPixels(bounds.top))),
                        static_cast<LONG>(std::lround(hostState._shellHost.DipsToPixels(bounds.right))),
                        static_cast<LONG>(std::lround(hostState._shellHost.DipsToPixels(bounds.bottom)))};
        };

        out.shellOkButtonBoundsPx     = controlBoundsToPx(hostState._okButtonControl);
        out.shellCancelButtonBoundsPx = controlBoundsToPx(hostState._cancelButtonControl);
        out.shellApplyButtonBoundsPx  = controlBoundsToPx(hostState._applyButtonControl);

        const auto rectInsideHost = [&](const RECT& rect) noexcept
        {
            return rect.right > rect.left && rect.bottom > rect.top && rect.left >= 0 && rect.top >= 0 &&
                   rect.right <= out.shellHostClientWidthPx && rect.bottom <= out.shellHostClientHeightPx;
        };
        out.shellFooterButtonsInsideHost = rectInsideHost(out.shellOkButtonBoundsPx) && rectInsideHost(out.shellCancelButtonBoundsPx) &&
                                           rectInsideHost(out.shellApplyButtonBoundsPx);

        wil::unique_hrgn clipRegion(CreateRectRgn(0, 0, 0, 0));
        const int clipType = clipRegion ? GetWindowRgn(hostState._shellHostHwnd, clipRegion.get()) : ERROR;
        const auto buttonCenterInsideClip = [&](const RECT& rect) noexcept
        {
            if (clipType == ERROR)
            {
                return true;
            }
            if (rect.right <= rect.left || rect.bottom <= rect.top)
            {
                return false;
            }
            const int x = rect.left + ((rect.right - rect.left) / 2);
            const int y = rect.top + ((rect.bottom - rect.top) / 2);
            return PtInRegion(clipRegion.get(), x, y) != FALSE;
        };
        out.shellFooterButtonsInsideClip = buttonCenterInsideClip(out.shellOkButtonBoundsPx) &&
                                           buttonCenterInsideClip(out.shellCancelButtonBoundsPx) &&
                                           buttonCenterInsideClip(out.shellApplyButtonBoundsPx);

        RedSalamander::DxUi::WindowHostBitmapCapture capture{};
        if (hostState._shellHost.DebugCaptureBitmap(capture))
        {
            const auto samplePixel = [&](const int x, const int y, uint32_t& outPixel) noexcept
            {
                if (x < 0 || y < 0 || static_cast<UINT>(x) >= capture.widthPx || static_cast<UINT>(y) >= capture.heightPx)
                {
                    return false;
                }
                const size_t base =
                    (static_cast<size_t>(y) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(x)) * 4u;
                if ((base + 3u) >= capture.bgraPixels.size())
                {
                    return false;
                }
                outPixel = static_cast<uint32_t>(capture.bgraPixels[base + 0u]) |
                           (static_cast<uint32_t>(capture.bgraPixels[base + 1u]) << 8u) |
                           (static_cast<uint32_t>(capture.bgraPixels[base + 2u]) << 16u) |
                           (static_cast<uint32_t>(capture.bgraPixels[base + 3u]) << 24u);
                return true;
            };

            const auto isWhite = [](const uint32_t bgra) noexcept
            {
                const uint8_t b = static_cast<uint8_t>(bgra & 0xFFu);
                const uint8_t g = static_cast<uint8_t>((bgra >> 8u) & 0xFFu);
                const uint8_t r = static_cast<uint8_t>((bgra >> 16u) & 0xFFu);
                return r >= 240u && g >= 240u && b >= 240u;
            };

            const auto colorComponent = [](const float value) noexcept
            {
                return static_cast<uint8_t>(std::clamp(std::lround(value * 255.0f), 0l, 255l));
            };
            const ThemePalette expectedPalette = PrefsUi::MakeDxPalette(state->theme);
            const uint8_t expectedB            = colorComponent(expectedPalette.windowBackground.b);
            const uint8_t expectedG            = colorComponent(expectedPalette.windowBackground.g);
            const uint8_t expectedR            = colorComponent(expectedPalette.windowBackground.r);
            const auto matchesExpectedBackground = [&](const uint32_t bgra) noexcept
            {
                const int b = static_cast<int>(bgra & 0xFFu);
                const int g = static_cast<int>((bgra >> 8u) & 0xFFu);
                const int r = static_cast<int>((bgra >> 16u) & 0xFFu);
                return std::abs(b - static_cast<int>(expectedB)) <= 4 && std::abs(g - static_cast<int>(expectedG)) <= 4 &&
                       std::abs(r - static_cast<int>(expectedR)) <= 4;
            };

            if (out.shellOkButtonBoundsPx.right > out.shellOkButtonBoundsPx.left && out.shellOkButtonBoundsPx.bottom > out.shellOkButtonBoundsPx.top)
            {
                const int okWidth = out.shellOkButtonBoundsPx.right - out.shellOkButtonBoundsPx.left;
                const int sampleX = out.shellOkButtonBoundsPx.left + std::clamp(okWidth / 4, 2, std::max(2, okWidth - 2));
                const int sampleY = out.shellOkButtonBoundsPx.top + ((out.shellOkButtonBoundsPx.bottom - out.shellOkButtonBoundsPx.top) / 2);
                out.shellOkButtonInteriorSampled = samplePixel(sampleX, sampleY, out.shellOkButtonInteriorBgra);
                if (out.shellOkButtonInteriorSampled)
                {
                    out.shellOkButtonInteriorLooksPainted = ! isWhite(out.shellOkButtonInteriorBgra);
                }
            }

            if (out.shellHostClientWidthPx > 0 && out.shellOkButtonBoundsPx.bottom > out.shellOkButtonBoundsPx.top)
            {
                const int sampleX = out.shellHostClientWidthPx / 2;
                const int sampleY = out.shellOkButtonBoundsPx.top + ((out.shellOkButtonBoundsPx.bottom - out.shellOkButtonBoundsPx.top) / 2);
                out.shellFooterBackgroundSampled = samplePixel(sampleX, sampleY, out.shellFooterBackgroundBgra);
                if (out.shellFooterBackgroundSampled)
                {
                    out.shellFooterBackgroundLooksThemed = matchesExpectedBackground(out.shellFooterBackgroundBgra);
                }
            }
        }
    }
    out.pageScrollY    = state->pageScrollY;
    out.pageScrollMaxY = state->pageScrollMaxY;
    out.pageHostLastWheelRouteSeen                      = state->debugLastWheelRouteSeen;
    out.pageHostLastWheelRouteForwarded                 = state->debugLastWheelRouteForwarded;
    out.pageHostLastWheelRouteTargetWasPageHost         = state->debugLastWheelRouteTargetWasPageHost;
    out.pageHostLastWheelRouteTargetWasCategoryTree     = state->debugLastWheelRouteTargetWasCategoryTree;
    out.pageHostLastWheelRouteTargetHadVerticalScroll   = state->debugLastWheelRouteTargetHadVerticalScroll;
    out.pageHostLastWheelWindowFromPointWasPageHost     = state->debugLastWheelWindowFromPointWasPageHost;
    out.pageHostLastWheelWindowFromPointWasCategoryTree = state->debugLastWheelWindowFromPointWasCategoryTree;
    out.pageHostLastWheelWndProcSeen                    = state->debugLastWheelWndProcSeen;
    out.pageHostLastWheelDxHandled                      = state->debugLastWheelDxHandled;
    out.pageHostLastWheelFallbackCalled                 = state->debugLastWheelFallbackCalled;
    out.pageHostLastWheelFallbackHandled                = state->debugLastWheelFallbackHandled;
    out.pageHostLastWheelDelta                          = state->debugLastWheelDelta;
    out.pageHostLastWheelClientX                        = state->debugLastWheelClientX;
    out.pageHostLastWheelClientY                        = state->debugLastWheelClientY;
    out.pageHostLastWheelBeforeY                        = state->debugLastWheelBeforeY;
    out.pageHostLastWheelBeforeMaxY                     = state->debugLastWheelBeforeMaxY;
    out.pageHostLastWheelAfterY                         = state->debugLastWheelAfterY;
    out.pageHostLastWheelAfterMaxY                      = state->debugLastWheelAfterMaxY;
    out.pageHostScrollRequestCount          = state->pageHostScrollRequestCount;
    out.pageHostScrollCoalescedRequestCount = state->pageHostScrollCoalescedRequestCount;
    out.pageHostScrollApplyCount            = state->pageHostScrollApplyCount;
    out.pageHostScrollMovedChildCountTotal  = state->pageHostScrollMovedChildCountTotal;
    out.pageHostDxScrollMovedControlCountTotal = state->pageHostDxScrollMovedControlCountTotal;
    out.pageHostDxScrollLastMovedControlCount  = state->pageHostDxScrollLastMovedControlCount;
    out.pageHostScrollLastApplyUs           = state->pageHostScrollLastApplyUs;
    out.pageHostScrollApplyPending          = state->pageHostScrollApplyPending;

    if (state->categoryTreeUsesDxUi)
    {
        const auto& hostState                   = static_cast<const PreferencesDialogHost&>(*state);
        out.categoryTreeDxHostRenderCount       = GetDxHostDebugRenderCount(hostState._categoryTreeHost);
        out.categoryTreeDxHostHasResizeFailures = GetDxHostDebugResizeFailureCount(hostState._categoryTreeHost) != 0u;
        if (hostState._categoryTreeControl)
        {
            const auto treeScrollbarState        = hostState._categoryTreeControl->DebugGetScrollbarVisualState(hostState._categoryTreeHost.GetTheme());
            out.categoryTreeHasVerticalScrollbar = treeScrollbarState.hasVerticalScrollbar;
            out.categoryTreeFirstVisibleIndex    = hostState._categoryTreeControl->DebugGetFirstVisibleIndex();
            if (const auto selectedVisibleIndex = hostState._categoryTreeControl->DebugGetSelectedVisibleIndex())
            {
                out.categoryTreeHasSelectedItem      = true;
                out.categoryTreeSelectedVisibleIndex = selectedVisibleIndex.value();
            }
            out.categoryTreeVerticalScrollDip = hostState._categoryTreeControl->DebugGetVerticalScrollDip();
        }
        out.pageHostDxHostRenderCount = GetDxHostDebugRenderCount(hostState._pageHostHost);
        out.createdPaneWindowCount    = 0u;
        out.visiblePaneWindowCount    = 0u;
        out.generalPaneVisible        = state->currentCategory == PrefCategory::General;
        out.pluginsPaneVisible        = state->currentCategory == PrefCategory::Plugins;
        if (hostState._pageTitleControl)
        {
            out.pageTitle = std::wstring(hostState._pageTitleControl->GetText());
        }
        if (hostState._pageDescriptionControl)
        {
            out.pageDescription = std::wstring(hostState._pageDescriptionControl->GetText());
        }

        const auto accumulateShellHost = [&](HWND hwnd, const WindowHost& dxHost) noexcept
        {
            if (! hwnd || ! PrefsUi::IsActuallyVisibleChildWindow(hwnd))
            {
                return;
            }
            if (GetDxHostDebugRenderCount(dxHost) != 0u)
            {
                ++out.visibleShellRenderedDxHostCount;
            }
            if (GetDxHostDebugResizeFailureCount(dxHost) != 0u)
            {
                ++out.shellDxHostResizeFailureCount;
            }
        };

        accumulateShellHost(hostState._shellHostHwnd, hostState._shellHost);

        const auto getCurrentPageWindow = [&]() noexcept -> HWND { return GetActivePreferencesPageRootWindow(*state); };

        const HWND currentPageWindow = getCurrentPageWindow();
        const bool currentPageUsesSharedHost =
            currentPageWindow && state->pageHostWindow && currentPageWindow == state->pageHostWindow && IsWindow(currentPageWindow) != FALSE;

        if (currentPageUsesSharedHost)
        {
            out.visibleCurrentPageChildWindowCount  = PrefsUi::IsActuallyVisibleChildWindow(currentPageWindow) ? 1u : 0u;
            out.currentPageRenderedDxHostCount      = GetDxHostDebugRenderCount(hostState._pageHostHost) != 0u ? 1u : 0u;
            out.currentPageDxHostResizeFailureCount = GetDxHostDebugResizeFailureCount(hostState._pageHostHost) != 0u ? 1u : 0u;
            out.currentPageDxHostRenderCountTotal   = GetDxHostDebugRenderCount(hostState._pageHostHost);
        }
        else
        {
            if (! PrefsDxHost::TryGetDirectHostMetrics(
                    currentPageWindow, out.visibleCurrentPageChildWindowCount, out.currentPageDxHostResizeFailureCount, out.currentPageDxHostRenderCountTotal))
            {
                if (currentPageWindow && state->pageHostWindow && currentPageWindow != state->pageHostWindow &&
                    PrefsUi::IsActuallyVisibleChildWindow(currentPageWindow))
                {
                    out.visibleCurrentPageChildWindowCount  = 1u;
                    out.currentPageDxHostResizeFailureCount = 0u;
                    out.currentPageDxHostRenderCountTotal   = 1u;
                }
                else
                {
                    out.visibleCurrentPageChildWindowCount  = CountVisibleChildWindows(currentPageWindow);
                    out.currentPageDxHostResizeFailureCount = PrefsDxHost::CountVisibleHostsWithResizeFailures(currentPageWindow);
                    out.currentPageDxHostRenderCountTotal   = PrefsDxHost::SumVisibleRenderedHostRenderCounts(currentPageWindow);
                }
            }
            out.currentPageRenderedDxHostCount = out.currentPageDxHostRenderCountTotal != 0u ? 1u : 0u;
        }
        out.pluginsExpanded                   = hostState._categoryTreeModel.IsPluginsExpanded();
        out.pluginsTreeChildCount             = hostState._categoryTreeModel.GetPluginItemCount();
        out.themesListRowCount                = hostState._themesPane.DebugListRowCount();
        const auto themesListMetrics          = hostState._themesPane.DebugListVisibleWorkMetrics();
        out.themesListVisibleRowCount         = static_cast<size_t>(themesListMetrics.visibleRowCount);
        out.themesListVisibleColumnCount      = static_cast<size_t>(themesListMetrics.visibleColumnCount);
        out.themesListVisibleCellCount        = static_cast<size_t>(themesListMetrics.visibleCellCount);
        out.themesListHasVerticalScrollbar    = themesListMetrics.hasVerticalScrollbar;
        out.themesListVerticalScrollDip       = themesListMetrics.verticalScrollDip;
        out.themesListRenderCount             = hostState._themesPane.DebugListRenderCount();
        out.themesListResizeCount             = hostState._themesPane.DebugListResizeCount();
        out.themesListResizeFailureCount      = hostState._themesPane.DebugListResizeFailureCount();
        out.themesSearchText                  = state->themesSearchText;
        out.themesSelectedThemeIdText         = state->workingSettings.theme.currentThemeId;
        out.themesSelectedColorKeyText        = state->themesSelectedColorKey;
        out.themesColorText                   = state->themesColorText;
        out.themesSelectedColorOverrideActive = false;
        out.generalFocusTarget                = hostState._generalPane.DebugGetFocusTarget();
        out.generalUsesDxUiTypographyContext  = hostState._generalPane.DebugUsesDxUiTypographyContext();
        out.generalUsesDxUiTypographyMetrics  = hostState._generalPane.DebugUsesDxUiTypographyMetrics();
        out.panesFocusTarget                  = hostState._panesPane.DebugGetFocusTarget();
        out.panesUsesDxUiTypographyContext    = hostState._panesPane.DebugUsesDxUiTypographyContext();
        out.panesUsesDxUiTypographyMetrics    = hostState._panesPane.DebugUsesDxUiTypographyMetrics();
        out.viewersUsesDxUiTypographyContext  = hostState._viewersPane.DebugUsesDxUiTypographyContext();
        out.viewersUsesDxUiTypographyMetrics  = hostState._viewersPane.DebugUsesDxUiTypographyMetrics();
        out.hotPathsFocusTarget               = hostState._hotPathsPane.DebugGetFocusTarget();
        out.advancedFocusTarget               = hostState._advancedPane.DebugGetFocusTarget();
        out.fileOperationsFocusTarget         = hostState._fileOperationsPane.DebugGetFocusTarget();
        out.compareDirectoriesFocusTarget     = hostState._compareDirectoriesPane.DebugGetFocusTarget();
        out.themesFocusTarget                 = hostState._themesPane.DebugGetFocusTarget();
        if (const auto* focusControl = hostState._shellHost.GetFocusControl(); focusControl != nullptr)
        {
            if (focusControl == hostState._resetAllButtonControl)
            {
                out.shellFocusTarget = PreferencesShellDebugFocusTarget::ResetAllButton;
            }
            else if (focusControl == hostState._okButtonControl)
            {
                out.shellFocusTarget = PreferencesShellDebugFocusTarget::OkButton;
            }
            else if (focusControl == hostState._cancelButtonControl)
            {
                out.shellFocusTarget = PreferencesShellDebugFocusTarget::CancelButton;
            }
            else if (focusControl == hostState._applyButtonControl)
            {
                out.shellFocusTarget = PreferencesShellDebugFocusTarget::ApplyButton;
            }
        }
        const ThemePalette palette = PrefsUi::MakeDxPalette(state->theme);
        out.previewApplied         = state->previewApplied;
        out.themeDark              = state->theme.dark;
        out.themeHighContrast      = state->theme.highContrast;
        out.themeRainbow           = state->theme.menu.rainbowMode;
        out.themeCompactMode       = state->theme.compactMode;
        out.themeReducedMotion     = palette.reducedMotion;
        out.generalUiLanguage      = GetUiSettingsOrDefault(state->workingSettings).language;
        if (out.generalUiLanguage.empty())
        {
            out.generalUiLanguage = L"system";
        }
        out.themeOverlayBackgroundArgb = PackColorArgb(palette.overlayBackground);
        out.themePrimaryBackdrop       = state->theme.primaryWindowBackdrop;
        out.themeToolBackdrop          = state->theme.toolWindowBackdrop;
        static_cast<void>(hostState._generalPane.DebugGetCompactModeToggleHeightDip(out.generalCompactToggleHeightDip));
        if (! state->themesSelectedColorKey.empty())
        {
            for (const auto& def : state->workingSettings.theme.themes)
            {
                if (def.id == state->workingSettings.theme.currentThemeId)
                {
                    out.themesSelectedColorOverrideActive = def.colors.contains(state->themesSelectedColorKey);
                    break;
                }
            }
        }
        out.viewersListRowCount                        = hostState._viewersPane.DebugListRowCount();
        const auto viewersListMetrics                  = hostState._viewersPane.DebugListVisibleWorkMetrics();
        out.viewersListVisibleRowCount                 = static_cast<size_t>(viewersListMetrics.visibleRowCount);
        out.viewersListVisibleColumnCount              = static_cast<size_t>(viewersListMetrics.visibleColumnCount);
        out.viewersListVisibleCellCount                = static_cast<size_t>(viewersListMetrics.visibleCellCount);
        out.viewersListHasVerticalScrollbar            = viewersListMetrics.hasVerticalScrollbar;
        out.viewersListVerticalScrollDip               = viewersListMetrics.verticalScrollDip;
        out.viewersListRenderCount                     = hostState._viewersPane.DebugListRenderCount();
        out.viewersListResizeCount                     = hostState._viewersPane.DebugListResizeCount();
        out.viewersListResizeFailureCount              = hostState._viewersPane.DebugListResizeFailureCount();
        out.viewersActionCount                         = state->workingSettings.fileActions.viewers.actions.size();
        out.viewersActionRowCount                      = hostState._viewersPane.DebugActionRowCount();
        out.viewersPrimaryActionIdText.clear();
        out.viewersAlternateActionIdText.clear();
        for (const auto& rule : state->workingSettings.fileActions.viewers.associations)
        {
            if (rule.match.kind == Common::Settings::FileActionMatchKind::Default && rule.computerName.empty())
            {
                out.viewersPrimaryActionIdText   = rule.viewActionId;
                out.viewersAlternateActionIdText = rule.alternateViewActionId;
                break;
            }
        }
        out.viewersSearchText                          = state->viewersSearchText;
        out.viewersSelectedExtensionText               = state->viewersSelectedExtensionText;
        out.viewersPreviewActionIdText                 = hostState._viewersPane.DebugPreviewActionId();
        out.viewersPreviewReasonText                   = hostState._viewersPane.DebugPreviewReason();
        out.viewersFocusTarget                         = hostState._viewersPane.DebugGetFocusTarget();
        out.editorsActionCount                         = state->workingSettings.fileActions.editors.actions.size();
        out.editorsAssociationRowCount                 = hostState._editorsPane.DebugAssociationRowCount();
        const auto editorsAssociationMetrics           = hostState._editorsPane.DebugAssociationVisibleWorkMetrics();
        out.editorsAssociationVisibleRowCount          = static_cast<size_t>(editorsAssociationMetrics.visibleRowCount);
        out.editorsAssociationVisibleColumnCount       = static_cast<size_t>(editorsAssociationMetrics.visibleColumnCount);
        out.editorsAssociationVisibleCellCount         = static_cast<size_t>(editorsAssociationMetrics.visibleCellCount);
        out.editorsAssociationHasVerticalScrollbar     = editorsAssociationMetrics.hasVerticalScrollbar;
        out.editorsAssociationVerticalScrollDip        = editorsAssociationMetrics.verticalScrollDip;
        out.editorsActionRowCount                      = hostState._editorsPane.DebugActionRowCount();
        out.editorsPrimaryActionIdText.clear();
        out.editorsAlternateActionIdText.clear();
        out.editorsEditNewActionIdText.clear();
        for (const auto& rule : state->workingSettings.fileActions.editors.associations)
        {
            if (rule.match.kind == Common::Settings::FileActionMatchKind::Default && rule.computerName.empty())
            {
                out.editorsPrimaryActionIdText   = rule.editActionId;
                out.editorsAlternateActionIdText = rule.alternateEditActionId;
                out.editorsEditNewActionIdText   = rule.editNewActionId;
                break;
            }
        }
        out.editorsPreviewActionIdText                 = hostState._editorsPane.DebugPreviewActionId();
        out.editorsPreviewReasonText                   = hostState._editorsPane.DebugPreviewReason();
        out.userMenuActionCount                        = state->workingSettings.userMenu.actions.size();
        out.keyboardListRowCount                       = hostState._keyboardPane.DebugListRowCount();
        const auto keyboardListMetrics                 = hostState._keyboardPane.DebugListVisibleWorkMetrics();
        out.keyboardListVisibleRowCount                = static_cast<size_t>(keyboardListMetrics.visibleRowCount);
        out.keyboardListVisibleColumnCount             = static_cast<size_t>(keyboardListMetrics.visibleColumnCount);
        out.keyboardListVisibleCellCount               = static_cast<size_t>(keyboardListMetrics.visibleCellCount);
        out.keyboardListHasVerticalScrollbar           = keyboardListMetrics.hasVerticalScrollbar;
        out.keyboardListVerticalScrollDip              = keyboardListMetrics.verticalScrollDip;
        out.keyboardListRenderCount                    = hostState._keyboardPane.DebugListRenderCount();
        out.keyboardListResizeCount                    = hostState._keyboardPane.DebugListResizeCount();
        out.keyboardListResizeFailureCount             = hostState._keyboardPane.DebugListResizeFailureCount();
        out.keyboardSearchText                         = state->keyboardSearchText;
        out.keyboardHintText                           = state->keyboardHintText;
        PreferencesKeyboardDebugSnapshot keyboardSnapshot{};
        if (hostState._keyboardPane.DebugGetSnapshot(keyboardSnapshot))
        {
            out.keyboardListColumnLayoutText = keyboardSnapshot.keyboardListColumnLayoutText;
        }
        out.keyboardFocusTarget                        = hostState._keyboardPane.DebugGetFocusTarget();
        out.keyboardCaptureActive                      = state->keyboardCaptureActive;
        out.pluginsMainListRowCount                    = hostState._pluginsPane.DebugMainListRowCount();
        const auto pluginsMainListMetrics              = hostState._pluginsPane.DebugMainListVisibleWorkMetrics();
        out.pluginsMainListVisibleRowCount             = static_cast<size_t>(pluginsMainListMetrics.visibleRowCount);
        out.pluginsMainListVisibleColumnCount          = static_cast<size_t>(pluginsMainListMetrics.visibleColumnCount);
        out.pluginsMainListVisibleCellCount            = static_cast<size_t>(pluginsMainListMetrics.visibleCellCount);
        out.pluginsMainListHasVerticalScrollbar        = pluginsMainListMetrics.hasVerticalScrollbar;
        out.pluginsMainListVerticalScrollDip           = pluginsMainListMetrics.verticalScrollDip;
        out.pluginsMainListRenderCount                 = hostState._pluginsPane.DebugMainListRenderCount();
        out.pluginsMainListResizeCount                 = hostState._pluginsPane.DebugMainListResizeCount();
        out.pluginsMainListResizeFailureCount          = hostState._pluginsPane.DebugMainListResizeFailureCount();
        out.pluginsCustomPathsListRowCount             = hostState._pluginsPane.DebugCustomPathsListRowCount();
        const auto pluginsCustomPathsListMetrics       = hostState._pluginsPane.DebugCustomPathsListVisibleWorkMetrics();
        out.pluginsCustomPathsListVisibleRowCount      = static_cast<size_t>(pluginsCustomPathsListMetrics.visibleRowCount);
        out.pluginsCustomPathsListVisibleColumnCount   = static_cast<size_t>(pluginsCustomPathsListMetrics.visibleColumnCount);
        out.pluginsCustomPathsListVisibleCellCount     = static_cast<size_t>(pluginsCustomPathsListMetrics.visibleCellCount);
        out.pluginsCustomPathsListHasVerticalScrollbar = pluginsCustomPathsListMetrics.hasVerticalScrollbar;
        out.pluginsCustomPathsListVerticalScrollDip    = pluginsCustomPathsListMetrics.verticalScrollDip;
        out.pluginsCustomPathsListRenderCount          = hostState._pluginsPane.DebugCustomPathsListRenderCount();
        out.pluginsCustomPathsListResizeCount          = hostState._pluginsPane.DebugCustomPathsListResizeCount();
        out.pluginsCustomPathsListResizeFailureCount   = hostState._pluginsPane.DebugCustomPathsListResizeFailureCount();
        out.pluginsCustomPathsEmptyPlaceholderVisible =
            state->currentCategory == PrefCategory::Plugins && ! state->pluginsDetailsActive && state->workingSettings.plugins.customPluginPaths.empty();
        out.pluginsSearchText                  = state->pluginsSearchText;
        out.pluginsSelectedPluginIdText        = state->pluginsSelectedPluginId;
        out.pluginsSelectedCustomPathText      = state->pluginsSelectedCustomPathText;
        out.pluginsStatusTitleText             = state->pluginsStatusTitleText;
        out.pluginsStatusBodyText              = state->pluginsStatusBodyText;
        out.pluginsDetailsConfigErrorText      = state->pluginsDetailsConfigErrorText;
        out.pluginsDetailsConfigEmptyStateText = state->pluginsDetailsConfigEmptyStateText;
        out.pluginsDetailsConfigFieldCount     = state->pluginsDetailsConfigFields.size();
        out.pluginsDetailsVisibleConfigFieldCount =
            static_cast<size_t>(std::count_if(state->pluginsDetailsConfigFields.begin(),
                                              state->pluginsDetailsConfigFields.end(),
                                              [](const PrefsPluginConfigFieldControls& controls) noexcept { return ! controls.field.uiHidden; }));
        if (const auto* configPanel = state->pluginsDetailsConfigDxPanel)
        {
            out.pluginsDetailsConfigDxPanelVisible = configPanel->IsVisible();
            out.pluginsDetailsConfigDxChildCount   = configPanel->GetChildren().size();
        }
        out.pluginsFocusTarget = hostState._pluginsPane.DebugGetFocusTarget();
    }

    return true;
}

bool PreferencesDialog::DebugGetKeyboardSnapshot(::PreferencesKeyboardDebugSnapshot& out) noexcept
{
    out = {};

    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state || ! state->categoryTreeUsesDxUi)
    {
        return false;
    }

    const auto& hostState = static_cast<const PreferencesDialogHost&>(*state);
    if (! hostState._keyboardPane.DebugGetSnapshot(out))
    {
        return false;
    }

    out.currentCategory = state->currentCategory;
    return true;
}

bool PreferencesDialog::DebugSelectCategory(const PrefCategory category) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    SelectCategory(dlg, *state, category);

    PreferencesDebugSnapshot snapshot{};
    return DebugGetSnapshot(snapshot) && snapshot.currentCategory == category;
}

bool PreferencesDialog::DebugSelectPluginsTreeChild(const size_t childIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    if (! hostState._categoryTreeControl)
    {
        return false;
    }

    hostState._categoryTreeModel.SetPluginsExpanded(true);
    hostState._categoryTreeModel.Rebuild();
    hostState._categoryTreeControl->NotifyDataChanged();

    size_t matchedPluginChildCount = 0u;
    const size_t visibleItemCount  = hostState._categoryTreeModel.GetVisibleItemCount();
    for (size_t visibleIndex = 0u; visibleIndex < visibleItemCount; ++visibleIndex)
    {
        TreeItemData item{};
        hostState._categoryTreeModel.GetVisibleItem(visibleIndex, item);

        PrefsPluginListItem pluginItem{};
        if (! PreferencesDialogHost::PreferencesCategoryTreeModel::TryDecodePluginNodeId(item.id, pluginItem))
        {
            continue;
        }

        if (matchedPluginChildCount != childIndex)
        {
            ++matchedPluginChildCount;
            continue;
        }

        if (! hostState._categoryTreeControl->RequestSelectVisibleItem(visibleIndex))
        {
            return false;
        }

        PreferencesDebugSnapshot snapshot{};
        return DebugGetSnapshot(snapshot) && snapshot.currentCategory == PrefCategory::Plugins && snapshot.pluginItemSelected;
    }

    return false;
}

namespace
{
[[nodiscard]] HWND GetDebugActivePageKeyboardInputHost(PreferencesDialogState& state) noexcept
{
    const HWND currentPageRoot = GetActivePreferencesPageRootWindow(state);
    const bool usesSharedHost  = currentPageRoot && state.pageHostWindow && currentPageRoot == state.pageHostWindow && IsWindow(currentPageRoot) != FALSE;
    if (usesSharedHost)
    {
        auto& hostState   = static_cast<PreferencesDialogHost&>(state);
        const HWND dxHost = hostState._pageHostHost.GetHwnd();
        return (dxHost && PrefsUi::IsActuallyVisibleChildWindow(dxHost)) ? dxHost : nullptr;
    }

    return (currentPageRoot && PrefsUi::IsActuallyVisibleChildWindow(currentPageRoot)) ? currentPageRoot : nullptr;
}
} // namespace

HWND PreferencesDialog::DebugGetActivePageHandle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return nullptr;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return nullptr;
    }

    const HWND pane = GetActivePreferencesPageRootWindow(*state);
    return (pane && PrefsUi::IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
}

HWND PreferencesDialog::DebugGetActivePageDxHostHandle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return nullptr;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return nullptr;
    }

    return GetDebugActivePageKeyboardInputHost(*state);
}

bool PreferencesDialog::DebugCaptureKeyboardShortcut(const uint32_t vk, const uint32_t modifiers) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state || ! state->categoryTreeUsesDxUi || state->currentCategory != PrefCategory::Keyboard)
    {
        return false;
    }

    const HWND host = GetDebugActivePageKeyboardInputHost(*state);
    return KeyboardPane::DebugApplyCapturedShortcut(host, *state, vk, modifiers);
}

HWND PreferencesDialog::DebugGetShellHostHandle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return nullptr;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return nullptr;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return (hostState._shellHostHwnd && PrefsUi::IsActuallyVisibleChildWindow(hostState._shellHostHwnd)) ? hostState._shellHostHwnd : nullptr;
}

bool PreferencesDialog::DebugScrollCategoryTreeByWheelDetents(const int detents) noexcept
{
    if (detents == 0)
    {
        return false;
    }

    const int direction = detents < 0 ? -1 : 1;
    const int steps     = std::abs(detents);
    for (int index = 0; index < steps; ++index)
    {
        if (! DebugScrollCategoryTreeByWheelDelta(direction * WHEEL_DELTA))
        {
            return false;
        }
    }

    return true;
}

bool PreferencesDialog::DebugScrollCategoryTreeByWheelDelta(const int wheelDelta) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg || wheelDelta == 0)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state || ! state->categoryTreeUsesDxUi)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    if (! hostState._categoryTreeControl)
    {
        return false;
    }

    hostState._categoryTreeHost.SetFocusControl(hostState._categoryTreeControl);

    const D2D1_RECT_F hitBounds = hostState._categoryTreeControl->GetHitBounds();
    if (hitBounds.right <= hitBounds.left || hitBounds.bottom <= hitBounds.top)
    {
        return false;
    }

    const D2D1_POINT_2F point = D2D1::Point2F((hitBounds.left + hitBounds.right) * 0.5f, (hitBounds.top + hitBounds.bottom) * 0.5f);
    return hostState._categoryTreeControl->OnMouseWheel(hostState._categoryTreeHost, point, static_cast<float>(wheelDelta), 0u);
}

bool PreferencesDialog::DebugDragPageHostDxScrollbarThumb(const int distancePx, const int moveCount) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg || distancePx <= 0 || moveCount <= 0)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state || ! state->pageHostUsesDxUi)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    auto* scrollPanel = dynamic_cast<ScrollPanel*>(hostState._pageHostContentRootControl);
    if (! scrollPanel || ! scrollPanel->IsInternalScrollbarEnabled() || state->pageScrollMaxY <= 0)
    {
        return false;
    }

    D2D1_RECT_F thumb{};
    if (! scrollPanel->DebugGetScrollbarThumbHitRect(thumb))
    {
        return false;
    }

    const float x = (thumb.left + thumb.right) * 0.5f;
    const float y = (thumb.top + thumb.bottom) * 0.5f;
    const float targetY = y + hostState._pageHostHost.PixelsToDip(static_cast<float>(distancePx));

    if (! scrollPanel->OnMouseDown(hostState._pageHostHost, D2D1::Point2F(x, y), false, MK_LBUTTON))
    {
        return false;
    }

    for (int index = 1; index <= moveCount; ++index)
    {
        const float stepY = y + ((targetY - y) * static_cast<float>(index)) / static_cast<float>(moveCount);
        scrollPanel->OnMouseMove(hostState._pageHostHost, D2D1::Point2F(x, stepY), MK_LBUTTON);
    }
    scrollPanel->OnMouseUp(hostState._pageHostHost, D2D1::Point2F(x, targetY), false, 0u);
    return true;
}

bool PreferencesDialog::DebugSelectPluginsMainListRow(const size_t rowIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugSelectMainListRow(rowIndex);
}

bool PreferencesDialog::DebugClickPluginsMainListRow(const size_t rowIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugClickMainListRow(rowIndex);
}

bool PreferencesDialog::DebugFindToggleablePluginsMainListRow(size_t& outRowIndex, bool& outEnabled) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugFindToggleableMainListRow(outRowIndex, outEnabled);
}

bool PreferencesDialog::DebugFindLoadablePluginsMainListRow(size_t& outRowIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugFindLoadableMainListRow(outRowIndex);
}

bool PreferencesDialog::DebugTogglePluginsMainListCheckbox(const size_t rowIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugToggleMainListCheckbox(rowIndex);
}

bool PreferencesDialog::DebugGetPluginsMainListRowEnabled(const size_t rowIndex, bool& outEnabled) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugGetMainListRowEnabled(rowIndex, outEnabled);
}

bool PreferencesDialog::DebugGetPluginsMainListCheckboxClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugGetMainListCheckboxClientRect(rowIndex, outRect);
}

bool PreferencesDialog::DebugGetPluginsMainListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugGetMainListHeaderClientRect(columnIndex, outRect);
}

bool PreferencesDialog::DebugSelectPluginsCustomPathsListRow(const size_t rowIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugSelectCustomPathsListRow(rowIndex);
}

bool PreferencesDialog::DebugClearPluginsCustomPaths() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugClearCustomPaths();
}

bool PreferencesDialog::DebugGetPluginsCustomPathsListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugGetCustomPathsListHeaderClientRect(columnIndex, outRect);
}

bool PreferencesDialog::DebugFocusPluginsSearchField() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugFocusSearchField();
}

bool PreferencesDialog::DebugFocusPluginsMainList() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg || IsWindow(dlg) == FALSE)
    {
        return false;
    }

    auto* state = PrefsUi::GetDialogState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugFocusMainList();
}

bool PreferencesDialog::DebugSelectKeyboardListRow(const size_t rowIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugSelectListRow(rowIndex);
}

bool PreferencesDialog::DebugFindKeyboardListRowByCommandId(std::wstring_view commandId, size_t& outRowIndex) noexcept
{
    outRowIndex = 0u;

    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state || ! state->categoryTreeUsesDxUi)
    {
        return false;
    }

    const auto& hostState = static_cast<const PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugFindListRowByCommandId(commandId, outRowIndex);
}

bool PreferencesDialog::DebugGetKeyboardVisibleRowChordByCommandId(std::wstring_view commandId, std::wstring& outChordText) noexcept
{
    outChordText.clear();

    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state || ! state->categoryTreeUsesDxUi)
    {
        return false;
    }

    const auto& hostState = static_cast<const PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugGetVisibleRowChordByCommandId(commandId, outChordText);
}

bool PreferencesDialog::DebugGetKeyboardListRowClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugGetListRowClientRect(rowIndex, outRect);
}

bool PreferencesDialog::DebugGetKeyboardListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugGetListHeaderClientRect(columnIndex, outRect);
}

bool PreferencesDialog::DebugHitTestKeyboardListClientPoint(
    const POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugHitTestListClientPoint(clientPoint, outZone, outColumnIndex, outHeaderResize, outHostHitsList);
}

bool PreferencesDialog::DebugGetKeyboardListPointerState(::PreferencesGridPointerDebugState& outState) noexcept
{
    outState       = {};
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugGetListPointerState(outState);
}

bool PreferencesDialog::DebugSelectViewersListRow(const size_t rowIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugSelectListRow(rowIndex);
}

bool PreferencesDialog::DebugGetViewersListRowClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugGetListRowClientRect(rowIndex, outRect);
}

bool PreferencesDialog::DebugGetViewersListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugGetListHeaderClientRect(columnIndex, outRect);
}

bool PreferencesDialog::DebugHitTestViewersListClientPoint(
    const POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugHitTestListClientPoint(clientPoint, outZone, outColumnIndex, outHeaderResize, outHostHitsList);
}

bool PreferencesDialog::DebugGetViewersListPointerState(::PreferencesGridPointerDebugState& outState) noexcept
{
    outState = {};
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugGetListPointerState(outState);
}

bool PreferencesDialog::DebugGetViewersTabClientRect(const size_t tabIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugGetTabClientRect(tabIndex, outRect);
}

bool PreferencesDialog::DebugGetViewersSelectedTabIndex(size_t& outIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugGetSelectedTabIndex(outIndex);
}

bool PreferencesDialog::DebugSelectThemesListRow(const size_t rowIndex) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._themesPane.DebugSelectListRow(rowIndex);
}

bool PreferencesDialog::DebugGetThemesListRowClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._themesPane.DebugGetListRowClientRect(rowIndex, outRect);
}

bool PreferencesDialog::DebugGetThemesListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._themesPane.DebugGetListHeaderClientRect(columnIndex, outRect);
}

bool PreferencesDialog::DebugSetPluginsSearchText(std::wstring_view text) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugSetSearchText(text);
}

bool PreferencesDialog::DebugSetViewersSearchText(std::wstring_view text) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugSetSearchText(text);
}

bool PreferencesDialog::DebugSelectViewersDefaultAction(const bool alternate, std::wstring_view actionId) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugSelectDefaultAction(alternate, actionId);
}

bool PreferencesDialog::DebugSelectEditorsDefaultAction(const bool alternate, std::wstring_view actionId) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._editorsPane.DebugSelectDefaultAction(alternate, actionId);
}

bool PreferencesDialog::DebugSelectEditorsDefaultEditNewAction(std::wstring_view actionId) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._editorsPane.DebugSelectDefaultEditNewAction(actionId);
}

bool PreferencesDialog::DebugSetKeyboardSearchText(std::wstring_view text) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugSetSearchText(text);
}

bool PreferencesDialog::DebugSetKeyboardFunctionBarScope() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugSetFunctionBarScope();
}

bool PreferencesDialog::DebugFocusViewersSearchField() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugFocusSearchField();
}

bool PreferencesDialog::DebugFocusKeyboardSearchField() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugFocusSearchField();
}

bool PreferencesDialog::DebugSetThemesSearchText(std::wstring_view text) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._themesPane.DebugSetSearchText(text);
}

bool PreferencesDialog::DebugFocusGeneralMenuBarToggle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._generalPane.DebugFocusMenuBarToggle();
}

bool PreferencesDialog::DebugGetGeneralMenuBarToggleChecked(bool& outChecked) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._generalPane.DebugGetMenuBarToggleChecked(outChecked);
}

bool PreferencesDialog::DebugSetGeneralCompactMode(bool checked) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._generalPane.DebugSetCompactMode(checked);
}

bool PreferencesDialog::DebugSelectGeneralLanguage(std::wstring_view displayText) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._generalPane.DebugSelectLanguageByText(displayText);
}

bool PreferencesDialog::DebugSelectGeneralReducedMotion(std::wstring_view displayText) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._generalPane.DebugSelectReducedMotionByText(displayText);
}

bool PreferencesDialog::DebugSelectGeneralWindowBackdrop(std::wstring_view displayText) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._generalPane.DebugSelectWindowBackdropByText(displayText);
}

bool PreferencesDialog::DebugFocusPanesLeftDisplayToggle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._panesPane.DebugFocusLeftDisplayToggle();
}

bool PreferencesDialog::DebugSelectPanesLeftDisplay(std::wstring_view displayText) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._panesPane.DebugSelectLeftDisplayByText(displayText);
}

bool PreferencesDialog::DebugFocusPanesLeftStatusBarToggle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._panesPane.DebugFocusLeftStatusBarToggle();
}

bool PreferencesDialog::DebugGetPanesLeftStatusBarToggleChecked(bool& outChecked) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._panesPane.DebugGetLeftStatusBarToggleChecked(outChecked);
}

bool PreferencesDialog::DebugFocusHotPathsFirstPathField() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._hotPathsPane.DebugFocusFirstPathField();
}

bool PreferencesDialog::DebugGetHotPathsFirstPathText(std::wstring& outText) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._hotPathsPane.DebugGetFirstPathText(outText);
}

bool PreferencesDialog::DebugSetHotPathsFirstPathText(std::wstring_view text) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._hotPathsPane.DebugSetFirstPathText(text);
}

bool PreferencesDialog::DebugFocusHotPathsOpenPrefsToggle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._hotPathsPane.DebugFocusOpenPrefsToggle();
}

bool PreferencesDialog::DebugGetHotPathsOpenPrefsToggleChecked(bool& outChecked) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._hotPathsPane.DebugGetOpenPrefsToggleChecked(outChecked);
}

bool PreferencesDialog::DebugFocusAdvancedBypassHelloToggle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._advancedPane.DebugFocusBypassHelloToggle();
}

bool PreferencesDialog::DebugSelectAdvancedFilterPreset(std::wstring_view displayText) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._advancedPane.DebugSelectFilterPresetByText(displayText);
}

bool PreferencesDialog::DebugSelectFileOperationsBandwidthPreset(std::wstring_view displayText) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState     = static_cast<PreferencesDialogHost&>(*state);
    const bool selected = hostState._fileOperationsPane.DebugSelectBandwidthPresetByText(displayText);
    if (selected && state->pageHostWindow && IsWindow(state->pageHostWindow) != FALSE)
    {
        LayoutPreferencesPageHost(state->pageHostWindow, *state);
        InvalidateRect(state->pageHostWindow, nullptr, FALSE);
        hostState._fileOperationsPane.PostLayoutFocusCustomBandwidthEdit();
    }

    return selected;
}

bool PreferencesDialog::DebugFocusFileOperationsPreCalcEnabledToggle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._fileOperationsPane.DebugFocusPreCalcEnabledToggle();
}

bool PreferencesDialog::DebugGetFileOperationsPreCalcEnabledToggleChecked(bool& outChecked) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._fileOperationsPane.DebugGetPreCalcEnabledToggleChecked(outChecked);
}

bool PreferencesDialog::DebugSelectCompareDirectoriesContentWorkers(std::wstring_view displayText) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._compareDirectoriesPane.DebugSelectContentWorkersByText(displayText);
}

bool PreferencesDialog::DebugSetFileOperationsBridgeBufferText(std::wstring_view text) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._fileOperationsPane.DebugSetBridgeBufferText(text);
}

bool PreferencesDialog::DebugCancelDialog() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    return RequestPreferencesDialogClose(dlg);
}

bool PreferencesDialog::DebugResetAllToDefaults(const bool confirm) noexcept
{
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"Preferences debug reset-all: begin confirm={}", static_cast<int>(confirm)));
#endif
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    if (! confirm)
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"Preferences debug reset-all: no-op cancel path");
#endif
        return true;
    }

    const bool autoAcceptPromptsBefore = HostGetAutoAcceptPrompts();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"Preferences debug reset-all: auto-accept before={}", static_cast<int>(autoAcceptPromptsBefore)));
#endif
    HostSetAutoAcceptPrompts(true);
    const auto restoreAutoAcceptPrompts = wil::scope_exit([&]() noexcept { HostSetAutoAcceptPrompts(autoAcceptPromptsBefore); });

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences debug reset-all: calling ResetAllPreferencesToDefaults");
#endif
    ResetAllPreferencesToDefaults(dlg, *state);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"Preferences debug reset-all: returned from ResetAllPreferencesToDefaults");
#endif
    return true;
}

bool PreferencesDialog::DebugFocusCompareDirectoriesSubdirectoriesToggle() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._compareDirectoriesPane.DebugFocusCompareSubdirectoriesToggle();
}

bool PreferencesDialog::DebugFocusCompareDirectoriesTarget(const PreferencesCompareDirectoriesDebugFocusTarget target) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._compareDirectoriesPane.DebugFocusTarget(target);
}

bool PreferencesDialog::DebugGetCompareDirectoriesToggleChecked(const PreferencesCompareDirectoriesDebugFocusTarget target, bool& outChecked) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._compareDirectoriesPane.DebugGetToggleChecked(target, outChecked);
}

bool PreferencesDialog::DebugFocusThemesSearchField() noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._themesPane.DebugFocusSearchField();
}

bool PreferencesDialog::DebugScrollPluginsMainListByWheelDetents(const int detents) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugScrollMainListByWheelDetents(detents);
}

bool PreferencesDialog::DebugScrollPluginsCustomPathsListByWheelDetents(const int detents) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._pluginsPane.DebugScrollCustomPathsListByWheelDetents(detents);
}

bool PreferencesDialog::DebugScrollKeyboardListByWheelDetents(const int detents) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._keyboardPane.DebugScrollListByWheelDetents(detents);
}

bool PreferencesDialog::DebugScrollViewersListByWheelDetents(const int detents) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._viewersPane.DebugScrollListByWheelDetents(detents);
}

bool PreferencesDialog::DebugScrollThemesListByWheelDetents(const int detents) noexcept
{
    const HWND dlg = GetHandle();
    if (! dlg)
    {
        return false;
    }

    auto* state = GetState(dlg);
    if (! state)
    {
        return false;
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(*state);
    return hostState._themesPane.DebugScrollListByWheelDetents(detents);
}
#endif
