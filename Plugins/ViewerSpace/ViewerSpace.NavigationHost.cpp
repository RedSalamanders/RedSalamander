#include "ViewerSpace.NavigationHost.h"

#include <DxUi/Typography.h>
#include <algorithm>
#include <atomic>
#include <chrono>
#include <exception>
#include <format>
#include <initializer_list>
#include <new>

#include "Helpers.h"

extern HINSTANCE g_hInstance;

namespace ViewerSpaceNavigation
{
namespace
{
constexpr float kMarginDip              = 6.0f;
constexpr float kGapDip                 = 4.0f;
constexpr float kMinimumCrumbWidthDip   = 88.0f;
constexpr float kOverflowWidthDip       = 32.0f;
constexpr float kRegularButtonWidthDip  = 36.0f;
constexpr wchar_t kBackGlyph            = L'\uE76B'; // Segoe Fluent Icons chevron-left.
constexpr wchar_t kForwardGlyph         = L'\uE76C'; // Segoe Fluent Icons chevron-right.
constexpr wchar_t kFallbackBackGlyph    = L'\u2039'; // ‹
constexpr wchar_t kFallbackForwardGlyph = L'\u203A'; // ›

[[nodiscard]] D2D1_RECT_F Rect(float left, float top, float right, float bottom) noexcept
{
    return D2D1::RectF(left, top, right, bottom);
}
} // namespace

NavigationHost::NavigationHost() noexcept = default;

NavigationHost::~NavigationHost() noexcept
{
    Detach();
}

HRESULT NavigationHost::Attach(HWND parent, const Captions& captions, Callbacks callbacks) noexcept
{
    if (parent == nullptr || IsWindow(parent) == FALSE)
    {
        return E_INVALIDARG;
    }
    if (_hwnd || _host.GetHwnd() != nullptr)
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
    }

    try
    {
        _captions  = captions;
        _callbacks = std::move(callbacks);
        _root      = std::make_unique<DxUi::Panel>();
        BuildControls();
        if (! RegisterChildClass())
        {
            const DWORD error = GetLastError();
            Detach();
            return HRESULT_FROM_WIN32((error == ERROR_SUCCESS) ? ERROR_CANNOT_MAKE : error);
        }

        _hwnd.reset(
            CreateWindowExW(0u, _windowClassName.c_str(), L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, 0, 0, 0, 0, parent, nullptr, g_hInstance, this));
        if (! _hwnd)
        {
            const DWORD error = GetLastError();
            Detach();
            return HRESULT_FROM_WIN32((error == ERROR_SUCCESS) ? ERROR_CANNOT_MAKE : error);
        }
        if (! _host.Attach(_hwnd.get()))
        {
            Detach();
            return E_FAIL;
        }

        _host.SetRoot(std::move(_root));
        // IconOnly buttons use Segoe Fluent Icons. Plain U+2190/U+2191/U+2192 arrows
        // are not glyphs in that font and render as tofu; use PUA Fluent glyphs and
        // switch to ordinary Unicode stand-ins if the optional font is unavailable.
        const bool hasFluentIcons = _host.HasFluentIconFont();
        _backButton->SetText(std::wstring(1u, hasFluentIcons ? kBackGlyph : kFallbackBackGlyph));
        _forwardButton->SetText(std::wstring(1u, hasFluentIcons ? kForwardGlyph : kFallbackForwardGlyph));
        _host.SetOnTabBoundary([this](bool reverse) noexcept
        {
            bool handled = false;
            InvokeFocusBoundary(reverse, handled);
            return handled;
        });
        SetState(_crumbs, _canBack, _canForward);
        SetBounds(_boundsPx);
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"ViewerSpace navigation host setup failed.");
        Detach();
        return E_FAIL;
    }
}

void NavigationHost::Detach() noexcept
{
    _callbacks = {};
    for (DxUi::Button* button : _folderButtons)
    {
        if (button != nullptr)
            button->SetOnClick({});
    }
    if (_analysisButton != nullptr)
        _analysisButton->SetOnClick({});
    if (_modeCombo != nullptr)
    {
        _modeCombo->SetOnSelectionChanged({});
        _modeCombo->SetOnPopupRequested({});
    }
    if (_backButton != nullptr)
    {
        _backButton->SetOnClick({});
    }
    if (_forwardButton != nullptr)
    {
        _forwardButton->SetOnClick({});
    }
    if (_ancestorCombo != nullptr)
    {
        _ancestorCombo->SetOnSelectionChanged({});
        _ancestorCombo->SetOnPopupRequested({});
    }
    for (DxUi::Button* button : _crumbButtons)
    {
        if (button != nullptr)
        {
            button->SetOnClick({});
        }
    }

    _host.SetOnTabBoundary({});
    _host.SetRoot(std::unique_ptr<DxUi::Control>{});
    _host.Detach();
    _hwnd.reset();
    if (_windowClassAtom != 0u && ! _windowClassName.empty())
    {
        static_cast<void>(UnregisterClassW(_windowClassName.c_str(), g_hInstance));
    }
    _windowClassAtom = 0u;
    _windowClassName.clear();
    _root.reset();
    _backButton    = nullptr;
    _forwardButton = nullptr;

    _analysisButton = nullptr;
    _modeCombo      = nullptr;
    _folderButtons.fill(nullptr);
    _ancestorCombo = nullptr;
    _crumbButtons.fill(nullptr);
    _visibleCrumbTokens.fill(0u);
    _crumbs.clear();
    _ancestorTokens.clear();
    _canBack         = false;
    _canForward      = false;
    _analysisVisible = true;

    _callbackFailureLogged = false;
}

void NavigationHost::SetBounds(const RECT& boundsPx) noexcept
{
    _boundsPx = boundsPx;
    if (_hwnd)
    {
        const int width  = static_cast<int>(std::max(0L, boundsPx.right - boundsPx.left));
        const int height = static_cast<int>(std::max(0L, boundsPx.bottom - boundsPx.top));
        static_cast<void>(SetWindowPos(_hwnd.get(), nullptr, boundsPx.left, boundsPx.top, width, height, SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOOWNERZORDER));
    }
}

void NavigationHost::ApplyTheme(const DxUi::ThemePalette& palette) noexcept
{
    auto breadcrumbPalette = palette;
    // The pinned Hyperlink uses selectionFill for its text. Breadcrumbs use ordinary
    // path text; a light selection background is not a readable link foreground.
    breadcrumbPalette.selectionFill = palette.text;
    breadcrumbPalette.selectionText = palette.windowBackground;
    _host.SetTheme(breadcrumbPalette);
}

void NavigationHost::SetState(std::span<const Crumb> crumbs, bool canBack, bool canForward)
{
    _canBack    = canBack;
    _canForward = canForward;

    const size_t count = std::min(crumbs.size(), kMaxCrumbs);
    std::vector<Crumb> nextCrumbs;
    nextCrumbs.reserve(count);
    for (size_t index = 0u; index < count; ++index)
    {
        nextCrumbs.push_back(crumbs[index]);
    }

    std::vector<uint64_t> nextAncestorTokens;
    std::vector<DxUi::ComboBox::Item> items;
    if (nextCrumbs.size() > 1u)
    {
        const size_t ancestorCount = nextCrumbs.size() - 1u;
        nextAncestorTokens.reserve(ancestorCount);
        items.reserve(ancestorCount + 1u);
        // The native noneditable ComboBox does not paint Placeholder; this sentinel keeps its closed face an ellipsis.
        items.push_back({L"", L"\u2026"});
        for (size_t index = 0u; index < ancestorCount; ++index)
        {
            nextAncestorTokens.push_back(nextCrumbs[index].id);
            items.push_back({std::to_wstring(index + 1u), nextCrumbs[index].text});
        }
    }
    if (_ancestorCombo != nullptr)
    {
        _ancestorCombo->SetItems(std::move(items));
        _ancestorCombo->SetSelectedIndex(nextAncestorTokens.empty() ? std::nullopt : std::optional<size_t>(0u));
    }
    _crumbs         = std::move(nextCrumbs);
    _ancestorTokens = std::move(nextAncestorTokens);
    if (_backButton != nullptr)
    {
        _backButton->SetEnabled(canBack);
    }
    if (_forwardButton != nullptr)
    {
        _forwardButton->SetEnabled(canForward);
    }
    if (_ancestorCombo != nullptr)
    {
        _ancestorCombo->SetEnabled(! _ancestorTokens.empty());
    }
    Layout();
    _host.RefreshAccessibilitySnapshot();
    _host.Invalidate();
}

bool NavigationHost::Focus(bool reverse) noexcept
{
    if (_host.GetHwnd() == nullptr)
    {
        return false;
    }
    DxUi::Control* target = nullptr;
    if (! reverse)
    {
        for (DxUi::Button* button : {_backButton, _forwardButton})
        {
            if (button != nullptr && button->IsVisible() && button->IsEnabled())
            {
                target = button;
                break;
            }
        }
        if (target == nullptr && _crumbButtons[0] != nullptr && _crumbButtons[0]->IsVisible() && _crumbButtons[0]->IsEnabled())
        {
            target = _crumbButtons[0];
        }
        if (target == nullptr && _ancestorCombo != nullptr && _ancestorCombo->IsVisible() && _ancestorCombo->IsEnabled())
        {
            target = _ancestorCombo;
        }
        if (target == nullptr)
        {
            for (size_t slot = 1u; slot < _crumbButtons.size(); ++slot)
            {
                DxUi::Button* const button = _crumbButtons[slot];
                if (button != nullptr && button->IsVisible() && button->IsEnabled())
                {
                    target = button;
                    break;
                }
            }
        }
        if (target == nullptr && _ancestorCombo != nullptr && _ancestorCombo->IsVisible() && _ancestorCombo->IsEnabled())
        {
            target = _ancestorCombo;
        }
    }
    else
    {
        if (_analysisButton != nullptr && _analysisButton->IsVisible())
            target = _analysisButton;
        for (size_t slot = _crumbButtons.size(); slot > 1u; --slot)
        {
            if (target != nullptr)
                break;
            DxUi::Button* const button = _crumbButtons[slot - 1u];
            if (button != nullptr && button->IsVisible() && button->IsEnabled())
            {
                target = button;
                break;
            }
        }
        if (target == nullptr && _ancestorCombo != nullptr && _ancestorCombo->IsVisible() && _ancestorCombo->IsEnabled())
        {
            target = _ancestorCombo;
        }
        if (target == nullptr && _crumbButtons[0] != nullptr && _crumbButtons[0]->IsVisible() && _crumbButtons[0]->IsEnabled())
        {
            target = _crumbButtons[0];
        }
        if (target == nullptr)
        {
            for (DxUi::Button* button : {_forwardButton, _backButton})
            {
                if (button != nullptr && button->IsVisible() && button->IsEnabled())
                {
                    target = button;
                    break;
                }
            }
        }
    }
    if (target == nullptr)
    {
        return false;
    }
    _host.SetFocusControl(target);
    return GetFocus() == _host.GetHwnd() || (GetFocus() != nullptr && IsChild(_host.GetHwnd(), GetFocus()) != FALSE);
}

HWND NavigationHost::GetHwnd() const noexcept
{
    return _hwnd.get();
}

void NavigationHost::SetAnalysisVisible(bool visible) noexcept
{
    _analysisVisible = visible;
    SetMode(_analystMode);
}

void NavigationHost::SetMode(bool analystMode) noexcept
{
    _analystMode = analystMode;
    if (_analysisButton == nullptr)
        return;
    const auto& caption = ! _analysisVisible ? _captions.mapOnly : analystMode ? _captions.analyst : _captions.explorer;
    // Segoe Fluent Icons: BackToWindow, DockBottom and DockRight. One control
    // exposes placement and visibility; it never targets another physical monitor.
    _analysisButton->SetText(std::wstring(1u, ! _host.HasFluentIconFont() ? L'\u25A6' : ! _analysisVisible ? L'\uE73F' : analystMode ? L'\uE90E' : L'\uE90D'));
    _analysisButton->SetAccessibleName(caption);
    _analysisButton->SetTooltipText(caption);
}

void NavigationHost::InvokeModeChanged(bool analystMode) noexcept
{
    try
    {
        if (_callbacks.modeChanged)
            _callbacks.modeChanged(analystMode);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // DxUi invokes this on the UI thread; keep exceptions inside the native callback boundary.
        LogCallbackFailureOnce();
    }
}

LRESULT CALLBACK NavigationHost::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_NCCREATE)
    {
        const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
        if (create == nullptr || create->lpCreateParams == nullptr)
        {
            return FALSE;
        }
        auto* self = static_cast<NavigationHost*>(create->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    }

    auto* self = reinterpret_cast<NavigationHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self != nullptr)
    {
        return self->HandleWindowMessage(hwnd, message, wParam, lParam);
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

LRESULT NavigationHost::HandleWindowMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    const ViewerSpaceHosts::ScopedHostDiagnostics diagnostics;
    if (message == WM_KEYDOWN || message == WM_SYSKEYDOWN)
    {
        const bool control = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool alt     = (GetKeyState(VK_MENU) & 0x8000) != 0;
        if (_host.GetFocusControl() == _modeCombo && ! control && ! alt && (wParam == VK_UP || wParam == VK_DOWN))
        {
            InvokeModeChanged(wParam == VK_DOWN);
            return 0;
        }
        const bool comboPopupOpen = (_modeCombo != nullptr && _modeCombo->IsPopupOpen()) || (_ancestorCombo != nullptr && _ancestorCombo->IsPopupOpen());
        const bool dismissCombo   = comboPopupOpen && (wParam == VK_ESCAPE || (alt && wParam == VK_UP));
        if (! dismissCombo && (wParam == VK_ESCAPE || wParam == VK_F5 || wParam == VK_BACK || (control && (wParam == '1' || wParam == '2')) ||
                               (alt && (wParam == VK_LEFT || wParam == VK_RIGHT || wParam == VK_UP))))
        {
            return SendMessageW(GetParent(hwnd), WM_KEYDOWN, wParam, lParam);
        }
    }
    bool handled         = false;
    const LRESULT result = _host.HandleMessage(hwnd, message, wParam, lParam, handled);
    if (message == WM_SIZE || message == WM_DPICHANGED_AFTERPARENT)
    {
        Layout();
    }
    if (message == WM_NCDESTROY)
    {
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    }
    return handled ? result : DefWindowProcW(hwnd, message, wParam, lParam);
}

bool NavigationHost::RegisterChildClass() noexcept
{
    static std::atomic<uint64_t> nextClassId{1u};
    try
    {
        const uint64_t classId = nextClassId.fetch_add(1u, std::memory_order_relaxed);
        _windowClassName       = std::format(L"RedSalamander.ViewerSpace.NavigationHost.{:016X}", classId);
        WNDCLASSEXW windowClass{};
        windowClass.cbSize        = sizeof(windowClass);
        windowClass.hInstance     = g_hInstance;
        windowClass.lpszClassName = _windowClassName.c_str();
        windowClass.lpfnWndProc   = WindowProc;
        windowClass.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        windowClass.style         = CS_HREDRAW | CS_VREDRAW;
        _windowClassAtom          = RegisterClassExW(&windowClass);
        return _windowClassAtom != 0u;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
        return false;
    }
}

void NavigationHost::BuildControls()
{
    _root->SetAccessibleName(_captions.path);
    _root->SetAccessibleAutomationId(L"ViewerSpace.Navigation.Toolbar");

    _backButton = _root->AddChild<DxUi::Button>();
    _backButton->SetVariant(DxUi::ButtonVariant::IconOnly);
    _backButton->SetAccessibleName(_captions.back);
    _backButton->SetAccessibleAutomationId(L"ViewerSpace.Navigation.Back");
    _backButton->SetTooltipText(_captions.back);
    _backButton->SetOnClick([this] { InvokeBack(); });

    _forwardButton = _root->AddChild<DxUi::Button>();
    _forwardButton->SetVariant(DxUi::ButtonVariant::IconOnly);
    _forwardButton->SetAccessibleName(_captions.forward);
    _forwardButton->SetAccessibleAutomationId(L"ViewerSpace.Navigation.Forward");
    _forwardButton->SetTooltipText(_captions.forward);
    _forwardButton->SetOnClick([this] { InvokeForward(); });

    const auto createCrumbButton = [this](size_t slot)
    {
        _crumbButtons[slot] = _root->AddChild<DxUi::Button>();
        _crumbButtons[slot]->SetVariant(DxUi::ButtonVariant::Hyperlink);
        _crumbButtons[slot]->SetAccessibleAutomationId(std::format(L"ViewerSpace.Navigation.Crumb{}", slot));
        _crumbButtons[slot]->SetAccessibleHelpText(_captions.path);
        _crumbButtons[slot]->SetVisible(false);
        _crumbButtons[slot]->SetOnClick([this, slot] { InvokeNavigate(_visibleCrumbTokens[slot]); });
        _folderButtons[slot] = _root->AddChild<DxUi::Button>();
        _folderButtons[slot]->SetVariant(DxUi::ButtonVariant::Disclosure);
        _folderButtons[slot]->SetDisclosureCollapsedDirection(DxUi::ChevronDirection::Right);
        _folderButtons[slot]->SetDisclosureExpanded(false);
        _folderButtons[slot]->SetAccessibleAutomationId(std::format(L"ViewerSpace.Navigation.Folders{}", slot));
        _folderButtons[slot]->SetAccessibleName(_captions.folders);
        _folderButtons[slot]->SetVisible(false);
        _folderButtons[slot]->SetOnClick([this, slot] { ShowFolders(slot); });
    };
    createCrumbButton(0u);

    _ancestorCombo = _root->AddChild<DxUi::ComboBox>();
    _ancestorCombo->SetEditable(false);
    _ancestorCombo->SetAccessibleName(_captions.ancestors);
    _ancestorCombo->SetAccessibleHelpText(_captions.path);
    _ancestorCombo->SetAccessibleAutomationId(L"ViewerSpace.Navigation.Ancestors");
    _ancestorCombo->SetOnPopupRequested([this]
    {
        ShowAncestorMenu();
        return true;
    });
    _ancestorCombo->SetOnSelectionChanged([this](size_t selected)
    {
        if (selected > 0u && selected - 1u < _ancestorTokens.size())
        {
            InvokeNavigate(_ancestorTokens[selected - 1u]);
        }
    });
    for (size_t slot = 1u; slot < _crumbButtons.size(); ++slot)
    {
        createCrumbButton(slot);
    }
    // DxUi traverses focusable children in insertion order. Presentation controls
    // follow the breadcrumbs; the panel toggle remains the final visible action.
    _analysisButton = _root->AddChild<DxUi::Button>();
    _analysisButton->SetVariant(DxUi::ButtonVariant::IconOnly);
    _analysisButton->SetAccessibleAutomationId(L"ViewerSpace.Navigation.Layout");
    _analysisButton->SetOnClick([this] { ShowModeMenu(); });
    SetMode(_analystMode);
}

void NavigationHost::Layout() noexcept
{
    auto* const root = static_cast<DxUi::Panel*>(_host.GetRoot());
    if (root == nullptr || _host.GetHwnd() == nullptr)
    {
        return;
    }

    try
    {
        RECT client{};
        if (GetClientRect(_host.GetHwnd(), &client) == 0)
        {
            return;
        }
        const float widthDip  = _host.PixelsToDip(static_cast<float>(std::max(0L, client.right - client.left)));
        const float heightDip = _host.PixelsToDip(static_cast<float>(std::max(0L, client.bottom - client.top)));
        const bool compact    = widthDip < 360.0f;
        const float margin    = compact ? 4.0f : kMarginDip;
        const float gap       = compact ? 3.0f : kGapDip;
        const float navWidth  = compact ? 28.0f : kRegularButtonWidthDip;
        const float navHeight = std::min(30.0f, std::max(24.0f, heightDip - 4.0f));
        const float navTop    = std::max(0.0f, (heightDip - navHeight) * 0.5f);
        root->SetBounds(Rect(0.0f, 0.0f, widthDip, heightDip));

        float left = margin;
        for (DxUi::Button* button : {_backButton, _forwardButton})
        {
            if (button != nullptr)
            {
                button->SetBounds(Rect(left, navTop, left + navWidth, navTop + navHeight));
                left += navWidth + gap;
            }
        }
        left                    = (left > margin) ? left - gap + gap : margin;
        const float layoutWidth = 30.0f;
        _analysisButton->SetBounds(Rect(widthDip - margin - layoutWidth, navTop, widthDip - margin, navTop + navHeight));
        const float right     = std::max(left, widthDip - margin - layoutWidth - gap);
        const float pathWidth = right - left;

        size_t visibleCount  = 0u;
        const size_t count   = _crumbs.size();
        const size_t maximum = std::min(count, kVisibleCrumbButtons);
        for (size_t candidate = 1u; candidate <= maximum; ++candidate)
        {
            const bool overflow = count > candidate;
            const float needed  = static_cast<float>(candidate) * kMinimumCrumbWidthDip +
                                  static_cast<float>((candidate > 0u ? candidate - 1u : 0u) + (overflow ? 1u : 0u)) * gap +
                                  (overflow ? kOverflowWidthDip : 0.0f);
            if (needed <= pathWidth)
            {
                visibleCount = candidate;
            }
        }
        if (count > 0u && visibleCount == 0u)
        {
            visibleCount = 1u;
        }

        const bool overflow        = count > visibleCount;
        size_t crumbSlot           = 0u;
        float crumbLeft            = left;
        const float overflowWidth  = overflow ? std::min(kOverflowWidthDip, pathWidth) : 0.0f;
        const float gaps           = static_cast<float>((visibleCount > 0u ? visibleCount - 1u : 0u) + (overflow ? 1u : 0u)) * gap;
        const float eachCrumbWidth = (visibleCount > 0u) ? std::max(1.0f, (pathWidth - overflowWidth - gaps) / static_cast<float>(visibleCount)) : 0.0f;

        const auto placeCrumb = [this, &crumbSlot, &crumbLeft, eachCrumbWidth, navTop, navHeight, gap](size_t crumbIndex)
        {
            if (crumbSlot >= _crumbButtons.size() || crumbIndex >= _crumbs.size())
            {
                return;
            }
            DxUi::Button* const button     = _crumbButtons[crumbSlot];
            const Crumb& crumb             = _crumbs[crumbIndex];
            _visibleCrumbTokens[crumbSlot] = crumb.id;
            button->SetVisible(true);
            button->SetEnabled(crumb.id != 0u);
            button->SetText(crumb.text);
            button->SetAccessibleName(crumb.text);
            button->SetTooltipText(crumb.text);
            const int textWidthPx    = DxUi::Typography::MeasureSingleLineTextWidthPx(_host.GetHwnd(), DxUi::FontRole::Body, crumb.text);
            const float textWidth    = _host.PixelsToDip(static_cast<float>(std::max(0, textWidthPx))) + 12.0f;
            const float segmentWidth = std::min(eachCrumbWidth, std::max(32.0f, textWidth + 22.0f));
            button->SetBounds(Rect(crumbLeft, navTop, crumbLeft + std::max(1.0f, segmentWidth - 22.0f), navTop + navHeight));
            _folderButtons[crumbSlot]->SetVisible(true);
            _folderButtons[crumbSlot]->SetBounds(Rect(crumbLeft + std::max(1.0f, segmentWidth - 22.0f), navTop, crumbLeft + segmentWidth, navTop + navHeight));
            _folderButtons[crumbSlot]->SetAccessibleHelpText(crumb.text);
            crumbLeft += segmentWidth + gap;
            ++crumbSlot;
        };

        _visibleCrumbTokens.fill(0u);
        for (size_t slot = 0u; slot < _crumbButtons.size(); ++slot)
        {
            _crumbButtons[slot]->SetVisible(false);
            _folderButtons[slot]->SetVisible(false);
        }
        if (count > 0u && overflow && visibleCount > 1u)
        {
            placeCrumb(0u);
            _ancestorCombo->SetVisible(true);
            _ancestorCombo->SetBounds(Rect(crumbLeft, navTop, crumbLeft + overflowWidth, navTop + navHeight));
            crumbLeft += overflowWidth + gap;
            const size_t suffixCount = visibleCount - 1u;
            const size_t suffixStart = count - suffixCount;
            for (size_t index = suffixStart; index < count; ++index)
            {
                placeCrumb(index);
            }
        }
        else if (count > 0u && overflow)
        {
            placeCrumb(count - 1u);
            _ancestorCombo->SetVisible(true);
            _ancestorCombo->SetBounds(Rect(crumbLeft, navTop, crumbLeft + overflowWidth, navTop + navHeight));
        }
        else
        {
            _ancestorCombo->SetVisible(false);
            for (size_t index = 0u; index < visibleCount; ++index)
            {
                placeCrumb(index);
            }
        }
        if (count == 0u)
        {
            _ancestorCombo->SetVisible(false);
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    }
}

void NavigationHost::ShowFolders(size_t slot) noexcept
{
    if (slot >= _folderButtons.size() || ! _callbacks.folders || ! _hwnd)
        return;
    // The menu snapshots retained model tokens; the owner rechecks the display generation on invocation.
    // No provider I/O is allowed on this UI-thread breadcrumb path.
    try
    {
        const auto started = std::chrono::steady_clock::now();
        const auto folders = _callbacks.folders(_visibleCrumbTokens[slot]);
        Debug::Perf::EmitDurationUs(L"viewer.space.navigation.folder_menu_us", Debug::Perf::ElapsedUs(started), folders.size());
        if (const auto chosen = ShowChoiceMenu(*_folderButtons[slot], folders))
            InvokeNavigate(chosen.value());
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    } // UI callback boundary; leave the breadcrumb available.
}

void NavigationHost::ShowModeMenu() noexcept
{
    const std::array<Crumb, 3u> choices{{{1u, _captions.explorer}, {2u, _captions.analyst}, {3u, _captions.mapOnly}}};
    const uint64_t selected = ! _analysisVisible ? 3u : _analystMode ? 2u : 1u;
    if (const auto chosen = ShowChoiceMenu(*_analysisButton, choices, selected))
    {
        if (chosen.value() == 3u)
        {
            if (_analysisVisible && _callbacks.toggleAnalysis)
                _callbacks.toggleAnalysis();
        }
        else
            InvokeModeChanged(chosen.value() == 2u);
    }
}

void NavigationHost::ShowAncestorMenu() noexcept
{
    if (_crumbs.size() > 1u)
    {
        if (const auto chosen = ShowChoiceMenu(*_ancestorCombo, std::span<const Crumb>(_crumbs).first(_crumbs.size() - 1u)))
            InvokeNavigate(chosen.value());
    }
}

std::optional<uint64_t> NavigationHost::ShowChoiceMenu(DxUi::Control& anchor, std::span<const Crumb> choices, std::optional<uint64_t> selected) noexcept
{
    try
    {
        const auto started = std::chrono::steady_clock::now();
        // Product-private token projection: the modal popup can pump a new model
        // generation. Snapshot IDs before showing it, then let the owner reject stale IDs.
        std::vector<uint64_t> tokens;
        std::vector<DxUi::MenuFlyoutItem> items;
        tokens.reserve(choices.size());
        items.reserve(choices.size());
        for (size_t index = 0u; index < choices.size(); ++index)
        {
            DxUi::MenuFlyoutItem item;
            // Popup labels interpret ampersands as mnemonics. Escape only at this
            // presentation boundary; model names and breadcrumb text stay literal.
            for (wchar_t character : choices[index].text)
            {
                item.text.push_back(character);
                if (character == L'&')
                    item.text.push_back(character);
            }
            item.commandId = static_cast<int>(index + 1u);
            if (selected.has_value())
            {
                item.kind    = DxUi::MenuItemKind::Radio;
                item.checked = selected.value() == choices[index].id;
            }
            tokens.push_back(choices[index].id);
            items.push_back(std::move(item));
        }
        Debug::Perf::EmitDurationUs(L"viewer.space.navigation.choice_menu_us", Debug::Perf::ElapsedUs(started), items.size());
        if (items.empty())
            return std::nullopt;
        const auto bounds = anchor.GetBounds();
        POINT point{static_cast<LONG>(_host.DipsToPixels(bounds.left)), static_cast<LONG>(_host.DipsToPixels(bounds.bottom))};
        if (! ClientToScreen(_hwnd.get(), &point))
            return std::nullopt;
        const auto chosen = DxUi::ContextMenu::Show(_hwnd.get(), point, items, _host.GetTheme());
        if (chosen.has_value() && chosen.value() > 0 && static_cast<size_t>(chosen.value()) <= tokens.size())
            return tokens[static_cast<size_t>(chosen.value() - 1)];
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    } // UI callback boundary; leave the selector available.
    return std::nullopt;
}

void NavigationHost::InvokeNavigate(uint64_t token) noexcept
{
    if (token == 0u || ! _callbacks.navigate)
    {
        return;
    }
    try
    {
        _callbacks.navigate(token);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    }
}

void NavigationHost::InvokeBack() noexcept
{
    if (! _callbacks.back)
    {
        return;
    }
    try
    {
        _callbacks.back();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    }
}

void NavigationHost::InvokeForward() noexcept
{
    if (! _callbacks.forward)
    {
        return;
    }
    try
    {
        _callbacks.forward();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    }
}

void NavigationHost::InvokeFocusBoundary(bool reverse, bool& handled) noexcept
{
    if (! _callbacks.focusBoundary)
    {
        return;
    }
    try
    {
        handled = _callbacks.focusBoundary(reverse);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        LogCallbackFailureOnce();
    }
}

void NavigationHost::LogCallbackFailureOnce() noexcept
{
    if (_callbackFailureLogged)
    {
        return;
    }
    _callbackFailureLogged = true;
    Debug::Error(L"ViewerSpace navigation host callback failed.");
}
} // namespace ViewerSpaceNavigation
