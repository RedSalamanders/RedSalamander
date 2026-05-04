#include "Framework.h"

#include "Preferences.UserMenu.h"

#include <algorithm>
#include <string>

#include "Helpers.h"
#include "UiMetrics.h"
#include "resource.h"

using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::WindowHost;

namespace
{
[[nodiscard]] std::wstring FileActionDisplayName(const Common::Settings::FileActionDefinition& action)
{
    if (! action.displayName.empty())
    {
        return action.displayName;
    }
    return action.id;
}

[[nodiscard]] std::wstring BuildUserMenuActionListText(const Common::Settings::UserMenuSettings& settings)
{
    if (settings.actions.empty())
    {
        return LoadStringResource(nullptr, IDS_PREFS_USER_MENU_EMPTY_BODY);
    }

    std::wstring text;
    for (const Common::Settings::FileActionDefinition& action : settings.actions)
    {
        if (action.id.empty())
        {
            continue;
        }

        if (! text.empty())
        {
            text.append(L"\r\n");
        }
        text.append(FileActionDisplayName(action));
        text.append(L"  ");
        text.push_back(L'(');
        text.append(action.id);
        text.push_back(L')');
    }

    if (text.empty())
    {
        return LoadStringResource(nullptr, IDS_PREFS_USER_MENU_EMPTY_BODY);
    }
    return text;
}
} // namespace

void UserMenuPane::OnVisibilityChanged(bool visible) noexcept
{
    if (! visible && _pageHost)
    {
        _pageHost->ResetInteractionState();
    }
}

void UserMenuPane::Destroy(PreferencesDialogState& state) noexcept
{
    static_cast<void>(state);
    DetachDxPageHost();
}

bool UserMenuPane::EnsureDxPageHost(HWND parent, PreferencesDialogState& state) noexcept
{
    static_cast<void>(parent);

    _pageHost        = state.pageHostDxHost;
    _pageContentRoot = state.pageHostDxContentRootControl;
    if (! _pageHost || ! _pageContentRoot)
    {
        return false;
    }

    if (PrefsUi::HasRetainedDxChildren(_pageContentRoot) && _title && _actions && _hint)
    {
        return true;
    }

    _pageHost->ResetInteractionState();
    _pageContentRoot->ClearChildren();

    _title   = _pageContentRoot->AddChild<Label>();
    _actions = _pageContentRoot->AddChild<Label>();
    _hint    = _pageContentRoot->AddChild<Label>();

    _title->SetFontRole(RedSalamander::DxUi::FontRole::Header);
    _actions->SetMultiline(true);
    _hint->SetMultiline(true);
    _hint->SetFontRole(RedSalamander::DxUi::FontRole::Small);

    ApplyTheme(state);
    SyncFromState(state);
    return true;
}

void UserMenuPane::DetachDxPageHost() noexcept
{
    if (_pageContentRoot && _pageHost)
    {
        _pageHost->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }

    _pageHost        = nullptr;
    _pageContentRoot = nullptr;
    _title           = nullptr;
    _actions         = nullptr;
    _hint            = nullptr;
    _state           = nullptr;
    _hostWindow      = nullptr;
}

void UserMenuPane::ApplyTheme(const PreferencesDialogState& state) noexcept
{
    if (_pageHost)
    {
        _pageHost->SetTheme(PrefsUi::MakeDxPalette(state.theme));
    }
}

void UserMenuPane::SyncFromState(PreferencesDialogState& state) noexcept
{
    const bool hasActions = ! state.workingSettings.userMenu.actions.empty();
    if (_title)
    {
        _title->SetText(LoadStringResource(nullptr, hasActions ? IDS_PREFS_USER_MENU_ACTIONS_TITLE : IDS_PREFS_USER_MENU_EMPTY_TITLE));
        _title->SetFontRole(RedSalamander::DxUi::FontRole::Header);
    }
    if (_actions)
    {
        _actions->SetText(BuildUserMenuActionListText(state.workingSettings.userMenu));
    }
    if (_hint)
    {
        _hint->SetText(LoadStringResource(nullptr, IDS_PREFS_USER_MENU_HINT));
        _hint->SetFontRole(RedSalamander::DxUi::FontRole::Small);
    }

    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void UserMenuPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _state      = &state;
    _hostWindow = parent;

    if (state.currentCategory != PrefCategory::UserMenu)
    {
        return;
    }

    if (! EnsureDxPageHost(parent, state))
    {
        Debug::Error(L"Preferences.UserMenu: Failed to initialize DxUi hosts.");
        DetachDxPageHost();
    }
}

void UserMenuPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    _hostWindow = host;
    _state      = &state;
    if (state.currentCategory == PrefCategory::UserMenu && ! EnsureDxPageHost(host, state))
    {
        Debug::Error(L"Preferences.UserMenu: Failed to refresh DxUi hosts.");
        return;
    }

    ApplyTheme(state);
    SyncFromState(state);
}

void UserMenuPane::LayoutPage(HWND host,
                              PreferencesDialogState& state,
                              int x,
                              int& y,
                              int width,
                              int margin,
                              int gapY,
                              int /*sectionY*/,
                              const PreferencesTypographyContext& typography) noexcept
{
    if (! host || ! EnsureDxPageHost(host, state))
    {
        return;
    }

    _hostWindow = host;
    _state      = &state;

    Debug::Perf::Scope layoutPerf(L"preferences.ui.usermenu_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(state.workingSettings.userMenu.actions.size()));
    layoutPerf.SetValue1(typography.dpi);

    const UINT dpi     = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const int rowGap   = std::max(1, gapY / 2);
    const int titleHeight = std::max(UiMetrics::ScaleDip(dpi, 24), PrefsUi::MeasureWrappedTextHeightPx(
                                                                 typography, typography.strong, width, _title ? std::wstring(_title->GetText()) : std::wstring{}));
    const auto pxToDip = [dpi](const int pixels) noexcept { return (static_cast<float>(pixels) * 96.0f) / static_cast<float>(dpi); };

    if (_title)
    {
        _title->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + titleHeight)));
    }
    y += titleHeight + rowGap;

    const std::wstring actionsText = _actions ? std::wstring(_actions->GetText()) : std::wstring{};
    const int actionsHeight = std::max(UiMetrics::ScaleDip(dpi, 40), PrefsUi::MeasureWrappedTextHeightPx(typography, typography.body, width, actionsText));
    if (_actions)
    {
        _actions->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + actionsHeight)));
    }
    y += actionsHeight + gapY;

    const std::wstring hintText = _hint ? std::wstring(_hint->GetText()) : LoadStringResource(nullptr, IDS_PREFS_USER_MENU_HINT);
    const int hintHeight = std::max(UiMetrics::ScaleDip(dpi, 40), PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, width, hintText));
    if (_hint)
    {
        _hint->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + hintHeight)));
    }
    y += hintHeight + margin;

    state.pageHostDirectContentBottomPx = std::max(state.pageHostDirectContentBottomPx, y);
    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}
