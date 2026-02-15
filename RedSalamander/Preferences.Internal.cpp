// Preferences.Internal.cpp

#include "Framework.h"

#include "Preferences.Internal.h"

#include <algorithm>
#include <array>
#include <cwchar>
#include <cwctype>
#include <fstream>
#include <limits>
#include <string>

#include <commctrl.h>

#include "FileSystemPluginManager.h"
#include "ThemedControls.h"
#include "ThemedInputFrames.h"
#include "ViewerPluginManager.h"

namespace
{
constexpr UINT_PTR kPrefsPaneForwardSubclassId  = 1u;
constexpr UINT_PTR kPrefsCenteredEditSubclassId = 3u;

void CenterMultilineEditTextVertically(HWND edit) noexcept
{
    ThemedControls::CenterEditTextVertically(edit);
}

LRESULT CALLBACK PrefsCenteredEditSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR uIdSubclass, DWORD_PTR /*refData*/) noexcept
{
    switch (msg)
    {
        case WM_SIZE:
        case WM_SETFONT: CenterMultilineEditTextVertically(hwnd); break;
        case WM_CHAR:
        {
            if (wp == L'\r' || wp == L'\n')
            {
                return 0;
            }

            const LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
            if ((style & ES_NUMBER) != 0)
            {
                if (wp < 0x20 || wp == 0x7F) // allow control chars (backspace, etc)
                {
                    break;
                }

                if (wp < L'0' || wp > L'9')
                {
                    return 0;
                }
            }
            break;
        }
        case WM_PASTE:
        {
            const LRESULT result = DefSubclassProc(hwnd, msg, wp, lp);

            const int length = GetWindowTextLengthW(hwnd);
            if (length <= 0)
            {
                CenterMultilineEditTextVertically(hwnd);
                return result;
            }

            std::wstring buffer;
            buffer.resize(static_cast<size_t>(length) + 1u);
            GetWindowTextW(hwnd, buffer.data(), length + 1);
            buffer.resize(static_cast<size_t>(length));

            buffer.erase(std::remove(buffer.begin(), buffer.end(), L'\r'), buffer.end());
            buffer.erase(std::remove(buffer.begin(), buffer.end(), L'\n'), buffer.end());
            buffer.erase(std::remove(buffer.begin(), buffer.end(), L'\t'), buffer.end());

            SetWindowTextW(hwnd, buffer.c_str());
            CenterMultilineEditTextVertically(hwnd);
            return result;
        }
        case WM_NCDESTROY:
        {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
            RemoveWindowSubclass(hwnd, PrefsCenteredEditSubclassProc, uIdSubclass);
#pragma warning(pop)
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

LRESULT CALLBACK PrefsPaneForwardSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR uIdSubclass, DWORD_PTR /*dwRefData*/) noexcept
{
    const HWND pageHost = GetParent(hwnd);

    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_PRINTCLIENT:
        {
            const HDC hdc = reinterpret_cast<HDC>(wp);
            if (! hdc || ! pageHost)
            {
                break;
            }

            const HWND dlg = GetParent(pageHost);
            auto* state    = dlg ? reinterpret_cast<PreferencesDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER)) : nullptr;
            if (! state)
            {
                break;
            }

            RECT rc{};
            GetClientRect(hwnd, &rc);

            HBRUSH brush = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(NULL_BRUSH));
            FillRect(hdc, &rc, brush);

            if (! state->theme.systemHighContrast && ! state->pageSettingCards.empty())
            {
                const UINT dpi         = GetDpiForWindow(hwnd);
                const int radius       = ThemedControls::ScaleDip(dpi, 6);
                const COLORREF surface = ThemedControls::GetControlSurfaceColor(state->theme);
                const COLORREF border  = ThemedControls::BlendColor(surface, state->theme.menu.text, state->theme.dark ? 40 : 30, 255);

                wil::unique_hbrush cardBrush(CreateSolidBrush(surface));
                wil::unique_hpen cardPen(CreatePen(PS_SOLID, 1, border));
                if (cardBrush && cardPen)
                {
                    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc, cardBrush.get());
                    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc, cardPen.get());

                    for (const RECT& card : state->pageSettingCards)
                    {
                        RoundRect(hdc, card.left, card.top, card.right, card.bottom, radius, radius);
                    }
                }
            }

            return 0;
        }
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (! hdc || ! pageHost)
            {
                return 0;
            }

            const HWND dlg = GetParent(pageHost);
            auto* state    = dlg ? reinterpret_cast<PreferencesDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER)) : nullptr;
            if (! state)
            {
                return 0;
            }

            RECT rc{};
            GetClientRect(hwnd, &rc);

            HBRUSH brush = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(NULL_BRUSH));
            FillRect(hdc.get(), &rc, brush);

            if (! state->theme.systemHighContrast && ! state->pageSettingCards.empty())
            {
                const UINT dpi         = GetDpiForWindow(hwnd);
                const int radius       = ThemedControls::ScaleDip(dpi, 6);
                const COLORREF surface = ThemedControls::GetControlSurfaceColor(state->theme);
                const COLORREF border  = ThemedControls::BlendColor(surface, state->theme.menu.text, state->theme.dark ? 40 : 30, 255);

                wil::unique_hbrush cardBrush(CreateSolidBrush(surface));
                wil::unique_hpen cardPen(CreatePen(PS_SOLID, 1, border));
                if (cardBrush && cardPen)
                {
                    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc.get(), cardBrush.get());
                    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc.get(), cardPen.get());

                    for (const RECT& card : state->pageSettingCards)
                    {
                        RoundRect(hdc.get(), card.left, card.top, card.right, card.bottom, radius, radius);
                    }
                }
            }

            return 0;
        }
        case WM_COMMAND:
        case WM_NOTIFY:
        case WM_DRAWITEM:
        case WM_MEASUREITEM:
        case WM_COMPAREITEM:
        case WM_DELETEITEM:
        case WM_VKEYTOITEM:
        case WM_CHARTOITEM:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORLISTBOX:
        case WM_CTLCOLORBTN:
            if (pageHost)
            {
                return SendMessageW(pageHost, msg, wp, lp);
            }
            break;
        case WM_NCDESTROY:
        {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
            RemoveWindowSubclass(hwnd, PrefsPaneForwardSubclassProc, uIdSubclass);
#pragma warning(pop)
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

LRESULT CALLBACK PrefsInputControlSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept
{
    return ThemedInputFrames::InputControlSubclassProc(hwnd, msg, wp, lp, subclassId, refData);
}

LRESULT CALLBACK PrefsInputFrameSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept
{
    auto* state = reinterpret_cast<PreferencesDialogState*>(refData);
    if (! state)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    ThemedInputFrames::FrameStyle frameStyle{};
    frameStyle.theme                        = &state->theme;
    frameStyle.backdropBrush                = state->cardBrush ? state->cardBrush.get() : state->backgroundBrush.get();
    frameStyle.inputBackgroundColor         = state->inputBackgroundColor;
    frameStyle.inputFocusedBackgroundColor  = state->inputFocusedBackgroundColor;
    frameStyle.inputDisabledBackgroundColor = state->inputDisabledBackgroundColor;

    return ThemedInputFrames::InputFrameSubclassProc(hwnd, msg, wp, lp, subclassId, reinterpret_cast<DWORD_PTR>(&frameStyle));
}
} // namespace

namespace PrefsPaneHost
{
[[nodiscard]] bool EnsureCreated(HWND pageHost, wil::unique_hwnd& paneHwnd) noexcept
{
    if (paneHwnd && IsWindow(paneHwnd.get()))
    {
        return true;
    }

    paneHwnd.reset();
    if (! pageHost)
    {
        return false;
    }

    constexpr DWORD style   = WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    constexpr DWORD exStyle = WS_EX_CONTROLPARENT;

    paneHwnd.reset(CreateWindowExW(exStyle, L"Static", L"", style, 0, 0, 10, 10, pageHost, nullptr, GetModuleHandleW(nullptr), nullptr));

    if (paneHwnd)
    {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(paneHwnd.get(), PrefsPaneForwardSubclassProc, kPrefsPaneForwardSubclassId, 0);
#pragma warning(pop)
    }

    return paneHwnd != nullptr;
}

void ResizeToHostClient(HWND pageHost, HWND paneHwnd) noexcept
{
    if (! paneHwnd || ! pageHost)
    {
        return;
    }

    RECT rc{};
    if (! GetClientRect(pageHost, &rc))
    {
        return;
    }

    const int width  = std::max(0l, rc.right - rc.left);
    const int height = std::max(0l, rc.bottom - rc.top);
    SetWindowPos(paneHwnd, nullptr, 0, 0, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
}

void Show(HWND paneHwnd, bool visible) noexcept
{
    if (! paneHwnd)
    {
        return;
    }

    ShowWindow(paneHwnd, visible ? SW_SHOW : SW_HIDE);
}

void ApplyScrollDelta(HWND pageHost, int dy) noexcept
{
    if (! pageHost || dy == 0)
    {
        return;
    }

    int childCount = 0;
    for (HWND child = GetWindow(pageHost, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT))
    {
        ++childCount;
    }

    HDWP hdwp = BeginDeferWindowPos(std::max(1, childCount));
    if (! hdwp)
    {
        for (HWND child = GetWindow(pageHost, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT))
        {
            RECT rc{};
            if (! GetWindowRect(child, &rc))
            {
                continue;
            }
            MapWindowPoints(nullptr, pageHost, reinterpret_cast<POINT*>(&rc), 2);
            SetWindowPos(child, nullptr, rc.left, rc.top + dy, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);
        }
        return;
    }

    for (HWND child = GetWindow(pageHost, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT))
    {
        RECT rc{};
        if (! GetWindowRect(child, &rc))
        {
            continue;
        }
        MapWindowPoints(nullptr, pageHost, reinterpret_cast<POINT*>(&rc), 2);
        hdwp = DeferWindowPos(hdwp, child, nullptr, rc.left, rc.top + dy, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);
        if (! hdwp)
        {
            break;
        }
    }

    if (hdwp)
    {
        static_cast<void>(EndDeferWindowPos(hdwp));
    }
}

void ScrollTo(HWND pageHost, PreferencesDialogState& state, int newScrollY) noexcept
{
    if (! pageHost)
    {
        return;
    }

    newScrollY = std::clamp(newScrollY, 0, state.pageScrollMaxY);
    if (newScrollY == state.pageScrollY)
    {
        return;
    }

    const int oldScrollY = state.pageScrollY;
    state.pageScrollY    = newScrollY;

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_POS;
    si.nPos   = state.pageScrollY;
    SetScrollInfo(pageHost, SB_VERT, &si, TRUE);

    const int dy = oldScrollY - state.pageScrollY;
    ApplyScrollDelta(pageHost, dy);
    RedrawWindow(pageHost, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
}

void EnsureControlVisible(HWND pageHost, PreferencesDialogState& state, HWND control) noexcept
{
    if (! pageHost || ! control || state.pageScrollMaxY <= 0)
    {
        return;
    }

    RECT rc{};
    if (! GetWindowRect(control, &rc))
    {
        return;
    }

    MapWindowPoints(nullptr, pageHost, reinterpret_cast<POINT*>(&rc), 2);

    RECT client{};
    GetClientRect(pageHost, &client);

    const UINT dpi          = GetDpiForWindow(pageHost);
    const int padY          = ThemedControls::ScaleDip(dpi, 10);
    const int desiredTop    = client.top + padY;
    const int desiredBottom = client.bottom - padY;

    int newScrollY = state.pageScrollY;
    if (rc.top < desiredTop)
    {
        newScrollY = state.pageScrollY + (rc.top - desiredTop);
    }
    else if (rc.bottom > desiredBottom)
    {
        newScrollY = state.pageScrollY + (rc.bottom - desiredBottom);
    }

    ScrollTo(pageHost, state, newScrollY);
}
} // namespace PrefsPaneHost

namespace PrefsInput
{
void CreateFramedComboBox(PreferencesDialogState& state, HWND parent, HWND& outFrame, HWND& outCombo, int controlId) noexcept
{
    const bool customFrames = ! state.theme.systemHighContrast;
    outFrame                = nullptr;
    outCombo                = nullptr;

    if (customFrames)
    {
        constexpr DWORD frameStyle = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS;
        outFrame                   = CreateWindowExW(0, L"Static", L"", frameStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr);

        if (outFrame)
        {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
            SetWindowSubclass(outFrame, PrefsInputFrameSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(&state));
#pragma warning(pop)
        }
    }

    if (customFrames)
    {
        outCombo = ThemedControls::CreateModernComboBox(parent, controlId, &state.theme);
    }
    else
    {
        outCombo = CreateWindowExW(WS_EX_CLIENTEDGE,
                                   L"ComboBox",
                                   L"",
                                   WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
                                   0,
                                   0,
                                   10,
                                   10,
                                   parent,
                                   reinterpret_cast<HMENU>(static_cast<INT_PTR>(controlId)),
                                   GetModuleHandleW(nullptr),
                                   nullptr);
    }

    if (outFrame && outCombo)
    {
        SetWindowLongPtrW(outFrame, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(outCombo));
    }

    if (outCombo)
    {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(outCombo, PrefsInputControlSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(outFrame));
#pragma warning(pop)
    }
}

void CreateFramedComboBox(PreferencesDialogState& state, HWND parent, wil::unique_hwnd& outFrame, wil::unique_hwnd& outCombo, int controlId) noexcept
{
    HWND frame = nullptr;
    HWND combo = nullptr;
    CreateFramedComboBox(state, parent, frame, combo, controlId);
    outFrame.reset(frame);
    outCombo.reset(combo);
}

void CreateFramedEditBox(PreferencesDialogState& state, HWND parent, HWND& outFrame, HWND& outEdit, int controlId, DWORD style) noexcept
{
    const bool customFrames = ! state.theme.systemHighContrast;
    outFrame                = nullptr;
    outEdit                 = nullptr;

    if (customFrames)
    {
        constexpr DWORD frameStyle = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS;
        outFrame                   = CreateWindowExW(0, L"Static", L"", frameStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr);

        if (outFrame)
        {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
            SetWindowSubclass(outFrame, PrefsInputFrameSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(&state));
#pragma warning(pop)
        }
    }

    const bool wantsCentering = (style & ES_MULTILINE) == 0;
    DWORD editStyle           = style;
    if (wantsCentering)
    {
        editStyle |= ES_MULTILINE;
        editStyle &= ~ES_WANTRETURN;
    }

    const DWORD editExStyle = customFrames ? 0 : WS_EX_CLIENTEDGE;
    outEdit                 = CreateWindowExW(editExStyle,
                              L"Edit",
                              L"",
                              editStyle,
                              0,
                              0,
                              10,
                              10,
                              parent,
                              reinterpret_cast<HMENU>(static_cast<INT_PTR>(controlId)),
                              GetModuleHandleW(nullptr),
                              nullptr);

    if (outFrame && outEdit)
    {
        SetWindowLongPtrW(outFrame, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(outEdit));
    }

    if (outEdit)
    {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(outEdit, PrefsInputControlSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(outFrame));
#pragma warning(pop)

        const UINT dpi       = GetDpiForWindow(outEdit);
        const int textMargin = ThemedControls::ScaleDip(dpi, 6);
        SendMessageW(outEdit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(textMargin, textMargin));

        if (wantsCentering)
        {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
            SetWindowSubclass(outEdit, PrefsCenteredEditSubclassProc, kPrefsCenteredEditSubclassId, 0);
#pragma warning(pop)
            CenterMultilineEditTextVertically(outEdit);
        }
    }
}

void CreateFramedEditBox(PreferencesDialogState& state, HWND parent, wil::unique_hwnd& outFrame, wil::unique_hwnd& outEdit, int controlId, DWORD style) noexcept
{
    HWND frame = nullptr;
    HWND edit  = nullptr;
    CreateFramedEditBox(state, parent, frame, edit, controlId, style);
    outFrame.reset(frame);
    outEdit.reset(edit);
}

void EnableMouseWheelForwarding(HWND control) noexcept
{
    if (! control)
    {
        return;
    }

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    constexpr UINT_PTR kPrefsMouseWheelForwardSubclassId = 2u;
    SetWindowSubclass(control, PrefsInputControlSubclassProc, kPrefsMouseWheelForwardSubclassId, 0);
#pragma warning(pop)
}

void EnableMouseWheelForwarding(const wil::unique_hwnd& control) noexcept
{
    EnableMouseWheelForwarding(control.get());
}
} // namespace PrefsInput

namespace PrefsPlugins
{
void BuildListItems(std::vector<PrefsPluginListItem>& out) noexcept
{
    out.clear();

    const auto& fsPlugins = FileSystemPluginManager::GetInstance().GetPlugins();
    out.reserve(fsPlugins.size());
    for (size_t i = 0; i < fsPlugins.size(); ++i)
    {
        if (! fsPlugins[i].id.empty())
        {
            out.push_back(PrefsPluginListItem{PrefsPluginType::FileSystem, i});
        }
    }

    const auto& viewerPlugins = ViewerPluginManager::GetInstance().GetPlugins();
    out.reserve(out.size() + viewerPlugins.size());
    for (size_t i = 0; i < viewerPlugins.size(); ++i)
    {
        if (! viewerPlugins[i].id.empty())
        {
            out.push_back(PrefsPluginListItem{PrefsPluginType::Viewer, i});
        }
    }

    std::sort(out.begin(),
              out.end(),
              [](const PrefsPluginListItem& a, const PrefsPluginListItem& b) noexcept
              {
                  if (a.type != b.type)
                  {
                      return a.type < b.type;
                  }

                  const int aOrigin = GetOriginOrder(a);
                  const int bOrigin = GetOriginOrder(b);
                  if (aOrigin != bOrigin)
                  {
                      return aOrigin < bOrigin;
                  }

                  const std::wstring_view aName = GetDisplayName(a);
                  const std::wstring_view bName = GetDisplayName(b);
                  if (aName.empty() || bName.empty())
                  {
                      return aName < bName;
                  }
                  return _wcsicmp(aName.data(), bName.data()) < 0;
              });
}

[[nodiscard]] std::wstring_view GetId(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        return (item.index < plugins.size()) ? std::wstring_view(plugins[item.index].id) : std::wstring_view{};
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    return (item.index < plugins.size()) ? std::wstring_view(plugins[item.index].id) : std::wstring_view{};
}

[[nodiscard]] std::wstring_view GetDisplayName(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        if (item.index >= plugins.size())
        {
            return {};
        }

        const auto& plugin = plugins[item.index];
        return plugin.name.empty() ? std::wstring_view(plugin.id) : std::wstring_view(plugin.name);
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    if (item.index >= plugins.size())
    {
        return {};
    }

    const auto& plugin = plugins[item.index];
    return plugin.name.empty() ? std::wstring_view(plugin.id) : std::wstring_view(plugin.name);
}

[[nodiscard]] std::wstring_view GetDescription(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        if (item.index >= plugins.size())
        {
            return {};
        }

        return std::wstring_view(plugins[item.index].description);
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    if (item.index >= plugins.size())
    {
        return {};
    }

    return std::wstring_view(plugins[item.index].description);
}

[[nodiscard]] std::wstring_view GetShortIdOrId(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        if (item.index >= plugins.size())
        {
            return {};
        }

        const auto& plugin = plugins[item.index];
        return plugin.shortId.empty() ? std::wstring_view(plugin.id) : std::wstring_view(plugin.shortId);
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    if (item.index >= plugins.size())
    {
        return {};
    }

    const auto& plugin = plugins[item.index];
    return plugin.shortId.empty() ? std::wstring_view(plugin.id) : std::wstring_view(plugin.shortId);
}

[[nodiscard]] bool IsLoadable(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        return (item.index < plugins.size()) ? plugins[item.index].loadable : false;
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    return (item.index < plugins.size()) ? plugins[item.index].loadable : false;
}

[[nodiscard]] int GetOriginOrder(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        return (item.index < plugins.size()) ? static_cast<int>(plugins[item.index].origin) : 0;
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    return (item.index < plugins.size()) ? static_cast<int>(plugins[item.index].origin) : 0;
}
} // namespace PrefsPlugins

namespace PrefsUi
{
std::wstring GetWindowTextString(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return {};
    }

    const int length = GetWindowTextLengthW(hwnd);
    if (length <= 0)
    {
        return {};
    }

    std::wstring text;
    text.resize(static_cast<size_t>(length) + 1u);
    const int copied = GetWindowTextW(hwnd, text.data(), length + 1);
    if (copied <= 0)
    {
        return {};
    }
    text.resize(static_cast<size_t>(copied));
    return text;
}

int MeasureStaticTextHeight(HWND referenceWindow, HFONT font, int width, std::wstring_view text) noexcept
{
    if (! referenceWindow || ! font || width <= 0 || text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return 0;
    }

    auto hdc = wil::GetDC(referenceWindow);
    if (! hdc)
    {
        return 0;
    }

    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), font);

    RECT rc{};
    rc.left   = 0;
    rc.top    = 0;
    rc.right  = width;
    rc.bottom = 0;

    DrawTextW(hdc.get(), text.data(), static_cast<int>(text.size()), &rc, DT_LEFT | DT_WORDBREAK | DT_NOPREFIX | DT_CALCRECT);

    const UINT dpi     = GetDpiForWindow(referenceWindow);
    const int paddingY = ThemedControls::ScaleDip(dpi, 6);
    return static_cast<int>(std::max(0l, rc.bottom - rc.top) + std::max(1, paddingY));
}

std::wstring_view TrimWhitespace(std::wstring_view text) noexcept
{
    const auto isSpace = [](wchar_t ch) noexcept { return std::iswspace(static_cast<wint_t>(ch)) != 0; };
    while (! text.empty() && isSpace(text.front()))
    {
        text.remove_prefix(1);
    }
    while (! text.empty() && isSpace(text.back()))
    {
        text.remove_suffix(1);
    }
    return text;
}

bool ContainsCaseInsensitive(std::wstring_view haystack, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }

    const auto it = std::search(haystack.begin(),
                                haystack.end(),
                                needle.begin(),
                                needle.end(),
                                [](wchar_t a, wchar_t b) noexcept { return std::towlower(static_cast<wint_t>(a)) == std::towlower(static_cast<wint_t>(b)); });
    return it != haystack.end();
}

void InvalidateComboBox(HWND combo) noexcept
{
    if (! combo)
    {
        return;
    }

    InvalidateRect(combo, nullptr, TRUE);

    COMBOBOXINFO cbi{};
    cbi.cbSize = sizeof(cbi);
    if (GetComboBoxInfo(combo, &cbi) && cbi.hwndItem)
    {
        InvalidateRect(cbi.hwndItem, nullptr, TRUE);
    }
}

void SelectComboItemByData(HWND combo, LPARAM data) noexcept
{
    if (! combo)
    {
        return;
    }

    const LRESULT count = SendMessageW(combo, CB_GETCOUNT, 0, 0);
    if (count == CB_ERR)
    {
        return;
    }

    for (LRESULT index = 0; index < count; ++index)
    {
        const LRESULT itemData = SendMessageW(combo, CB_GETITEMDATA, static_cast<WPARAM>(index), 0);
        if (itemData != CB_ERR && itemData == data)
        {
            SendMessageW(combo, CB_SETCURSEL, static_cast<WPARAM>(index), 0);
            InvalidateComboBox(combo);
            return;
        }
    }
}

std::optional<LPARAM> TryGetSelectedComboItemData(HWND combo) noexcept
{
    if (! combo)
    {
        return std::nullopt;
    }

    const LRESULT sel = SendMessageW(combo, CB_GETCURSEL, 0, 0);
    if (sel == CB_ERR)
    {
        return std::nullopt;
    }

    const LRESULT data = SendMessageW(combo, CB_GETITEMDATA, static_cast<WPARAM>(sel), 0);
    if (data == CB_ERR)
    {
        return std::nullopt;
    }

    return static_cast<LPARAM>(data);
}

void SetTwoStateToggleState(HWND toggle, bool highContrast, bool toggledOn) noexcept
{
    if (! toggle)
    {
        return;
    }

    const LONG_PTR style = GetWindowLongPtrW(toggle, GWL_STYLE);
    const UINT type      = static_cast<UINT>(style & BS_TYPEMASK);
    bool useBmCheck      = highContrast;
    if (type == BS_OWNERDRAW)
    {
        useBmCheck = false;
    }
    else if (type == BS_CHECKBOX || type == BS_AUTOCHECKBOX || type == BS_3STATE || type == BS_AUTO3STATE || type == BS_RADIOBUTTON ||
             type == BS_AUTORADIOBUTTON)
    {
        useBmCheck = true;
    }

    if (useBmCheck)
    {
        SendMessageW(toggle, BM_SETCHECK, toggledOn ? BST_CHECKED : BST_UNCHECKED, 0);
        return;
    }

    SetWindowLongPtrW(toggle, GWLP_USERDATA, toggledOn ? 1 : 0);
    InvalidateRect(toggle, nullptr, TRUE);
}

bool GetTwoStateToggleState(HWND toggle, bool highContrast) noexcept
{
    if (! toggle)
    {
        return false;
    }

    const LONG_PTR style = GetWindowLongPtrW(toggle, GWL_STYLE);
    const UINT type      = static_cast<UINT>(style & BS_TYPEMASK);
    bool useBmCheck      = highContrast;
    if (type == BS_OWNERDRAW)
    {
        useBmCheck = false;
    }
    else if (type == BS_CHECKBOX || type == BS_AUTOCHECKBOX || type == BS_3STATE || type == BS_AUTO3STATE || type == BS_RADIOBUTTON ||
             type == BS_AUTORADIOBUTTON)
    {
        useBmCheck = true;
    }

    if (useBmCheck)
    {
        return SendMessageW(toggle, BM_GETCHECK, 0, 0) == BST_CHECKED;
    }

    return GetWindowLongPtrW(toggle, GWLP_USERDATA) != 0;
}

std::optional<uint32_t> TryParseUInt32(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return std::nullopt;
    }

    uint32_t value = 0;
    for (const wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return std::nullopt;
        }

        const uint32_t digit = static_cast<uint32_t>(ch - L'0');
        if (value > (std::numeric_limits<uint32_t>::max() - digit) / 10u)
        {
            return std::nullopt;
        }

        value = (value * 10u) + digit;
    }

    return value;
}

std::optional<uint64_t> TryParseUInt64(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return std::nullopt;
    }

    uint64_t value = 0;
    for (const wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return std::nullopt;
        }

        const uint64_t digit = static_cast<uint64_t>(ch - L'0');
        if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10ull)
        {
            return std::nullopt;
        }

        value = (value * 10ull) + digit;
    }

    return value;
}

bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        const wchar_t ca = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(a[i])));
        const wchar_t cb = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(b[i])));
        if (ca != cb)
        {
            return false;
        }
    }

    return true;
}

} // namespace PrefsUi

namespace PrefsUi
{

// Schema-driven UI helper functions
[[nodiscard]] HWND CreateSchemaToggle(HWND parent,
                                      const SettingsSchemaParser::SettingField& field,
                                      PreferencesDialogState& state,
                                      int x,
                                      int& y,
                                      int width,
                                      int margin,
                                      int gapY,
                                      HFONT font) noexcept
{
    const UINT dpi         = GetDpiForWindow(parent);
    const int rowHeight    = ThemedControls::ScaleDip(dpi, 32);
    const int toggleWidth  = ThemedControls::ScaleDip(dpi, 40);
    const int toggleHeight = ThemedControls::ScaleDip(dpi, 20);

    // Create label
    const int labelHeight = ThemedControls::ScaleDip(dpi, 20);
    HWND label            = CreateWindowExW(0,
                                 L"STATIC",
                                 field.title.c_str(),
                                 WS_CHILD | WS_VISIBLE | SS_LEFT | SS_CENTERIMAGE,
                                 x + margin,
                                 y,
                                 width - margin - toggleWidth - ThemedControls::ScaleDip(dpi, 12),
                                 labelHeight,
                                 parent,
                                 nullptr,
                                 GetModuleHandleW(nullptr),
                                 nullptr);

    if (label && font)
    {
        SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
    }

    // Create toggle button (will be owner-drawn)
    HWND toggle = CreateWindowExW(0,
                                  L"BUTTON",
                                  L"",
                                  WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                                  x + width - toggleWidth - margin,
                                  y + (rowHeight - toggleHeight) / 2,
                                  toggleWidth,
                                  toggleHeight,
                                  parent,
                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(10000 + state.schemaFields.size())), // Unique ID
                                  GetModuleHandleW(nullptr),
                                  nullptr);

    if (toggle && font)
    {
        SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
    }

    y += rowHeight + gapY;
    return toggle;
}

[[nodiscard]] HWND CreateSchemaEdit(HWND parent,
                                    const SettingsSchemaParser::SettingField& field,
                                    PreferencesDialogState& state,
                                    int x,
                                    int& y,
                                    int width,
                                    int margin,
                                    int gapY,
                                    HFONT font) noexcept
{
    const UINT dpi        = GetDpiForWindow(parent);
    const int labelHeight = ThemedControls::ScaleDip(dpi, 20);
    const int rowSpacing  = ThemedControls::ScaleDip(dpi, 4);

    // Create label
    HWND label = CreateWindowExW(0,
                                 L"STATIC",
                                 field.title.c_str(),
                                 WS_CHILD | WS_VISIBLE | SS_LEFT,
                                 x + margin,
                                 y,
                                 width - margin * 2,
                                 labelHeight,
                                 parent,
                                 nullptr,
                                 GetModuleHandleW(nullptr),
                                 nullptr);

    if (label && font)
    {
        SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
    }

    y += labelHeight + rowSpacing;

    // Create framed edit control using existing helper for consistent styling
    HWND frame = nullptr;
    HWND edit  = nullptr;
    PrefsInput::CreateFramedEditBox(
        state, parent, frame, edit, 10000 + static_cast<int>(state.schemaFields.size()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_LEFT | ES_AUTOHSCROLL);

    if (edit)
    {
        if (font)
        {
            SendMessageW(edit, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
        }

        // Set initial value
        SetWindowTextW(edit, field.defaultValue.c_str());
    }

    const int editHeight = ThemedControls::ScaleDip(dpi, 28);
    y += editHeight + gapY;
    return edit;
}

[[nodiscard]] HWND CreateSchemaNumber(HWND parent,
                                      const SettingsSchemaParser::SettingField& field,
                                      PreferencesDialogState& state,
                                      int x,
                                      int& y,
                                      int width,
                                      int margin,
                                      int gapY,
                                      HFONT font) noexcept
{
    const UINT dpi        = GetDpiForWindow(parent);
    const int labelHeight = ThemedControls::ScaleDip(dpi, 20);
    const int rowSpacing  = ThemedControls::ScaleDip(dpi, 4);

    // Create label
    HWND label = CreateWindowExW(0,
                                 L"STATIC",
                                 field.title.c_str(),
                                 WS_CHILD | WS_VISIBLE | SS_LEFT,
                                 x + margin,
                                 y,
                                 width - margin * 2,
                                 labelHeight,
                                 parent,
                                 nullptr,
                                 GetModuleHandleW(nullptr),
                                 nullptr);

    if (label && font)
    {
        SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
    }

    y += labelHeight + rowSpacing;

    // Create framed number edit control using existing helper
    HWND frame = nullptr;
    HWND edit  = nullptr;
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    frame,
                                    edit,
                                    10000 + static_cast<int>(state.schemaFields.size()),
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_LEFT | ES_AUTOHSCROLL | ES_NUMBER);

    if (edit)
    {
        if (font)
        {
            SendMessageW(edit, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
        }

        // Set initial value
        SetWindowTextW(edit, field.defaultValue.c_str());
    }

    const int editHeight = ThemedControls::ScaleDip(dpi, 28);
    y += editHeight + gapY;
    return edit;
}

[[nodiscard]] HWND CreateSchemaControl(HWND parent,
                                       const SettingsSchemaParser::SettingField& field,
                                       PreferencesDialogState& state,
                                       int x,
                                       int& y,
                                       int width,
                                       int margin,
                                       int gapY,
                                       HFONT font) noexcept
{
    // Route to appropriate control type
    if (field.controlType == L"toggle")
    {
        return CreateSchemaToggle(parent, field, state, x, y, width, margin, gapY, font);
    }
    else if (field.controlType == L"number")
    {
        return CreateSchemaNumber(parent, field, state, x, y, width, margin, gapY, font);
    }
    else if (field.controlType == L"edit")
    {
        return CreateSchemaEdit(parent, field, state, x, y, width, margin, gapY, font);
    }

    // Default to edit
    return CreateSchemaEdit(parent, field, state, x, y, width, margin, gapY, font);
}

void PositionControl(HWND hwnd, int x, int y, int width, int height) noexcept
{
    if (hwnd)
    {
        SetWindowPos(hwnd, nullptr, x, y, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    }
}

void PositionAndSetFont(HWND hwnd, HFONT font, int x, int y, int width, int height) noexcept
{
    if (hwnd)
    {
        SetWindowPos(hwnd, nullptr, x, y, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }
}

void SetControlText(HWND hwnd, std::wstring_view text) noexcept
{
    if (hwnd && ! text.empty())
    {
        SetWindowTextW(hwnd, text.data());
    }
}

int CalculateCardHeight(int rowHeight, int titleHeight, int cardPaddingY, int cardGapY, int descHeight) noexcept
{
    const int contentHeight = std::max(0, titleHeight + (descHeight > 0 ? (cardGapY + descHeight) : 0));
    return std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);
}

void TryPushCard(std::vector<RECT>& cards, const RECT& card) noexcept
{
    cards.push_back(card);
}

} // namespace PrefsUi

namespace PrefsFile
{
bool TryReadFileToString(const std::filesystem::path& path, std::string& out) noexcept
{
    out.clear();

    std::ifstream file(path, std::ios::binary);
    if (! file)
    {
        return false;
    }

    file.seekg(0, std::ios::end);
    const std::streamoff end = file.tellg();
    if (end < 0)
    {
        return false;
    }

    file.seekg(0, std::ios::beg);
    out.resize(static_cast<size_t>(end));
    if (end > 0)
    {
        file.read(out.data(), static_cast<std::streamsize>(end));
        if (! file)
        {
            return false;
        }
    }
    return true;
}

bool TryWriteFileFromString(const std::filesystem::path& path, std::string_view text) noexcept
{
    std::ofstream file(path, std::ios::binary | std::ios::trunc);
    if (! file)
    {
        return false;
    }

    if (! text.empty())
    {
        file.write(text.data(), static_cast<std::streamsize>(text.size()));
        if (! file)
        {
            return false;
        }
    }

    file.flush();
    return static_cast<bool>(file);
}
} // namespace PrefsFile

namespace PrefsListView
{
int GetSingleLineRowHeightPx(HWND list, HDC hdc) noexcept
{
    if (! list)
    {
        return 26;
    }

    const UINT dpi      = GetDpiForWindow(list);
    const int minHeight = std::max(1, ThemedControls::ScaleDip(dpi, 26));
    const int paddingY  = std::max(1, ThemedControls::ScaleDip(dpi, 3));

    if (! hdc)
    {
        return minHeight;
    }

    TEXTMETRICW tm{};
    if (! GetTextMetricsW(hdc, &tm))
    {
        return minHeight;
    }

    const int lineHeight = std::max(1, static_cast<int>(tm.tmHeight + tm.tmExternalLeading));
    return std::max(minHeight, (paddingY * 2) + lineHeight);
}

LRESULT DrawThemedTwoColumnListRow(DRAWITEMSTRUCT* dis, PreferencesDialogState& state, HWND list, UINT expectedCtlId, bool secondColumnRightAlign) noexcept
{
    if (! dis || dis->CtlType != ODT_LISTVIEW || dis->CtlID != expectedCtlId)
    {
        return 0;
    }

    if (! list || ! dis->hDC)
    {
        return 1;
    }

    const int itemIndex = static_cast<int>(dis->itemID);
    if (itemIndex < 0)
    {
        return 1;
    }

    RECT rc = dis->rcItem;
    if (rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return 1;
    }

    wchar_t seedText[256]{};
    ListView_GetItemText(list, itemIndex, 0, seedText, static_cast<int>(std::size(seedText)));
    const std::wstring_view seed = std::wstring_view(seedText, std::wcslen(seedText));

    const bool selected    = (dis->itemState & ODS_SELECTED) != 0;
    const bool focused     = (dis->itemState & ODS_FOCUS) != 0;
    const bool listFocused = GetFocus() == list;

    const HWND root         = GetAncestor(list, GA_ROOT);
    const bool windowActive = root && GetActiveWindow() == root;

    const bool systemHighContrast = state.theme.systemHighContrast;
    COLORREF bg                   = systemHighContrast ? GetSysColor(COLOR_WINDOW) : state.theme.windowBackground;
    COLORREF textColor            = systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : state.theme.menu.text;

    if (selected)
    {
        COLORREF selBg = systemHighContrast ? GetSysColor(COLOR_HIGHLIGHT) : state.theme.menu.selectionBg;
        if (! state.theme.highContrast && state.theme.menu.rainbowMode && ! seed.empty())
        {
            selBg = RainbowMenuSelectionColor(seed, state.theme.menu.darkBase);
        }

        COLORREF selText = systemHighContrast ? GetSysColor(COLOR_HIGHLIGHTTEXT) : state.theme.menu.selectionText;
        if (! state.theme.highContrast && state.theme.menu.rainbowMode)
        {
            selText = ChooseContrastingTextColor(selBg);
        }

        if (windowActive && listFocused)
        {
            bg        = selBg;
            textColor = selText;
        }
        else if (! state.theme.highContrast)
        {
            const int denom = state.theme.menu.darkBase ? 2 : 3;
            bg              = ThemedControls::BlendColor(state.theme.windowBackground, selBg, 1, denom);
            textColor       = ChooseContrastingTextColor(bg);
        }
        else
        {
            bg        = selBg;
            textColor = selText;
        }
    }
    else if (! state.theme.highContrast && ((itemIndex % 2) == 1))
    {
        const COLORREF tint =
            (state.theme.menu.rainbowMode && ! seed.empty()) ? RainbowMenuSelectionColor(seed, state.theme.menu.darkBase) : state.theme.menu.selectionBg;
        const int denom = state.theme.menu.darkBase ? 6 : 8;
        bg              = ThemedControls::BlendColor(bg, tint, 1, denom);
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(bg));
    if (bgBrush)
    {
        FillRect(dis->hDC, &rc, bgBrush.get());
    }

    if (! state.theme.highContrast && textColor == bg)
    {
        textColor = ChooseContrastingTextColor(bg);
    }

    const UINT dpi     = GetDpiForWindow(list);
    const int paddingX = ThemedControls::ScaleDip(dpi, 8);

    const int col0W = std::max(0, ListView_GetColumnWidth(list, 0));
    const int col1W = std::max(0, ListView_GetColumnWidth(list, 1));

    RECT col0Rect  = rc;
    col0Rect.right = std::min(rc.right, rc.left + col0W);

    RECT col1Rect  = rc;
    col1Rect.left  = col0Rect.right;
    col1Rect.right = (col1W > 0) ? std::min(rc.right, col1Rect.left + col1W) : rc.right;

    wchar_t text0[256]{};
    ListView_GetItemText(list, itemIndex, 0, text0, static_cast<int>(std::size(text0)));
    wchar_t text1[512]{};
    ListView_GetItemText(list, itemIndex, 1, text1, static_cast<int>(std::size(text1)));

    HFONT fontToUse = reinterpret_cast<HFONT>(SendMessageW(list, WM_GETFONT, 0, 0));
    if (! fontToUse)
    {
        fontToUse = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    [[maybe_unused]] auto oldFont = wil::SelectObject(dis->hDC, fontToUse);

    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, textColor);

    RECT textRect0  = col0Rect;
    textRect0.left  = std::min(textRect0.right, textRect0.left + paddingX);
    textRect0.right = std::max(textRect0.left, textRect0.right - paddingX);

    DrawTextW(dis->hDC, text0, static_cast<int>(std::wcslen(text0)), &textRect0, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);

    RECT textRect1  = col1Rect;
    textRect1.left  = std::min(textRect1.right, textRect1.left + paddingX);
    textRect1.right = std::max(textRect1.left, textRect1.right - paddingX);

    UINT flags = DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX;
    flags |= secondColumnRightAlign ? DT_RIGHT : DT_LEFT;

    DrawTextW(dis->hDC, text1, static_cast<int>(std::wcslen(text1)), &textRect1, flags);

    if (focused)
    {
        RECT focusRc = rc;
        InflateRect(&focusRc, -ThemedControls::ScaleDip(dpi, 2), -ThemedControls::ScaleDip(dpi, 2));

        COLORREF focusTint = state.theme.menu.selectionBg;
        if (! state.theme.highContrast && state.theme.menu.rainbowMode && ! seed.empty())
        {
            focusTint = RainbowMenuSelectionColor(seed, state.theme.menu.darkBase);
        }

        const int weight          = (windowActive && listFocused) ? (state.theme.dark ? 70 : 55) : (state.theme.dark ? 55 : 40);
        const COLORREF focusColor = systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : ThemedControls::BlendColor(bg, focusTint, weight, 255);

        wil::unique_hpen focusPen(CreatePen(PS_SOLID, 1, focusColor));
        if (focusPen)
        {
            [[maybe_unused]] auto oldBrush2 = wil::SelectObject(dis->hDC, GetStockObject(NULL_BRUSH));
            [[maybe_unused]] auto oldPen2   = wil::SelectObject(dis->hDC, focusPen.get());
            Rectangle(dis->hDC, focusRc.left, focusRc.top, focusRc.right, focusRc.bottom);
        }
    }

    return 1;
}
} // namespace PrefsListView

namespace PrefsFolders
{
FolderPanePreferences GetFolderPanePreferences(const Common::Settings::Settings& settings, std::wstring_view slot) noexcept
{
    FolderPanePreferences prefs{};
    if (! settings.folders.has_value())
    {
        return prefs;
    }

    for (const auto& pane : settings.folders->items)
    {
        if (pane.slot != slot)
        {
            continue;
        }

        prefs.display          = pane.view.display;
        prefs.sortBy           = pane.view.sortBy;
        prefs.sortDirection    = pane.view.sortDirection;
        prefs.statusBarVisible = pane.view.statusBarVisible;
        break;
    }

    return prefs;
}

uint32_t GetFolderHistoryMax(const Common::Settings::Settings& settings) noexcept
{
    if (! settings.folders.has_value())
    {
        return Common::Settings::FoldersSettings{}.historyMax;
    }

    return std::clamp(settings.folders->historyMax, 1u, 50u);
}

bool AreEquivalentFolderPreferences(const Common::Settings::Settings& a, const Common::Settings::Settings& b) noexcept
{
    if (GetFolderHistoryMax(a) != GetFolderHistoryMax(b))
    {
        return false;
    }

    const FolderPanePreferences leftA  = GetFolderPanePreferences(a, kLeftPaneSlot);
    const FolderPanePreferences leftB  = GetFolderPanePreferences(b, kLeftPaneSlot);
    const FolderPanePreferences rightA = GetFolderPanePreferences(a, kRightPaneSlot);
    const FolderPanePreferences rightB = GetFolderPanePreferences(b, kRightPaneSlot);

    if (leftA.display != leftB.display || leftA.sortBy != leftB.sortBy || leftA.sortDirection != leftB.sortDirection ||
        leftA.statusBarVisible != leftB.statusBarVisible)
    {
        return false;
    }

    if (rightA.display != rightB.display || rightA.sortBy != rightB.sortBy || rightA.sortDirection != rightB.sortDirection ||
        rightA.statusBarVisible != rightB.statusBarVisible)
    {
        return false;
    }

    return true;
}

Common::Settings::FolderSortDirection DefaultFolderSortDirection(Common::Settings::FolderSortBy sortBy) noexcept
{
    switch (sortBy)
    {
        case Common::Settings::FolderSortBy::Time:
        case Common::Settings::FolderSortBy::Size: return Common::Settings::FolderSortDirection::Descending;
        case Common::Settings::FolderSortBy::Name:
        case Common::Settings::FolderSortBy::Extension:
        case Common::Settings::FolderSortBy::Attributes:
        case Common::Settings::FolderSortBy::None: return Common::Settings::FolderSortDirection::Ascending;
    }
    return Common::Settings::FolderSortDirection::Ascending;
}

Common::Settings::FoldersSettings* EnsureWorkingFoldersSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.folders.has_value())
    {
        return &settings.folders.value();
    }

    settings.folders.emplace();
    return &settings.folders.value();
}

Common::Settings::FolderPane* EnsureWorkingFolderPane(Common::Settings::Settings& settings, std::wstring_view slot) noexcept
{
    Common::Settings::FoldersSettings* folders = settings.folders.has_value() ? &settings.folders.value() : EnsureWorkingFoldersSettings(settings);
    if (! folders)
    {
        return nullptr;
    }

    for (auto& pane : folders->items)
    {
        if (pane.slot == slot)
        {
            return &pane;
        }
    }

    Common::Settings::FolderPane pane{};
    pane.slot.assign(slot);
    folders->items.push_back(std::move(pane));
    return &folders->items.back();
}
} // namespace PrefsFolders

namespace PrefsMonitor
{
const Common::Settings::MonitorSettings& GetMonitorSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::MonitorSettings kDefaults{};
    if (settings.monitor.has_value())
    {
        return settings.monitor.value();
    }
    return kDefaults;
}

Common::Settings::MonitorSettings* EnsureWorkingMonitorSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.monitor.has_value())
    {
        return &settings.monitor.value();
    }

    settings.monitor.emplace();
    return &settings.monitor.value();
}
} // namespace PrefsMonitor

namespace PrefsCache
{
const Common::Settings::CacheSettings& GetCacheSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::CacheSettings kDefaults{};
    if (settings.cache.has_value())
    {
        return settings.cache.value();
    }
    return kDefaults;
}

Common::Settings::CacheSettings* EnsureWorkingCacheSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.cache.has_value())
    {
        return &settings.cache.value();
    }

    settings.cache.emplace();
    return &settings.cache.value();
}

void MaybeResetWorkingCacheSettingsIfEmpty(Common::Settings::Settings& settings) noexcept
{
    if (! settings.cache.has_value())
    {
        return;
    }

    const auto& directoryInfo     = settings.cache->directoryInfo;
    const bool wroteDirectoryInfo = (directoryInfo.maxBytes.has_value() && directoryInfo.maxBytes.value() > 0) || directoryInfo.maxWatchers.has_value() ||
                                    directoryInfo.mruWatched.has_value();
    if (! wroteDirectoryInfo)
    {
        settings.cache.reset();
    }
}

std::optional<uint64_t> TryParseCacheBytes(std::wstring_view text) noexcept
{
    std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
    if (trimmed.empty())
    {
        return std::nullopt;
    }

    size_t digitCount = 0;
    while (digitCount < trimmed.size() && trimmed[digitCount] >= L'0' && trimmed[digitCount] <= L'9')
    {
        ++digitCount;
    }
    if (digitCount == 0)
    {
        return std::nullopt;
    }

    const auto valueOpt = PrefsUi::TryParseUInt64(trimmed.substr(0, digitCount));
    if (! valueOpt.has_value())
    {
        return std::nullopt;
    }

    const uint64_t value   = valueOpt.value();
    std::wstring_view unit = PrefsUi::TrimWhitespace(trimmed.substr(digitCount));
    uint64_t multiplier    = 1024ull;

    if (! unit.empty())
    {
        if (PrefsUi::EqualsNoCase(unit, L"kb") || PrefsUi::EqualsNoCase(unit, L"kib"))
        {
            multiplier = 1024ull;
        }
        else if (PrefsUi::EqualsNoCase(unit, L"mb") || PrefsUi::EqualsNoCase(unit, L"mib"))
        {
            multiplier = 1024ull * 1024ull;
        }
        else if (PrefsUi::EqualsNoCase(unit, L"gb") || PrefsUi::EqualsNoCase(unit, L"gib"))
        {
            multiplier = 1024ull * 1024ull * 1024ull;
        }
        else
        {
            return std::nullopt;
        }
    }

    if (value == 0 || multiplier == 0)
    {
        return 0ull;
    }

    if (value > std::numeric_limits<uint64_t>::max() / multiplier)
    {
        return std::nullopt;
    }

    return value * multiplier;
}

std::wstring FormatCacheBytes(uint64_t bytes) noexcept
{
    constexpr uint64_t kKiB = 1024ull;
    constexpr uint64_t kMiB = 1024ull * 1024ull;
    constexpr uint64_t kGiB = 1024ull * 1024ull * 1024ull;

    if (bytes == 0)
    {
        return {};
    }

    if (bytes % kGiB == 0)
    {
        std::wstring text = std::to_wstring(bytes / kGiB);
        text.append(L" GB");
        return text;
    }
    if (bytes % kMiB == 0)
    {
        std::wstring text = std::to_wstring(bytes / kMiB);
        text.append(L" MB");
        return text;
    }
    if (bytes % kKiB == 0)
    {
        std::wstring text = std::to_wstring(bytes / kKiB);
        text.append(L" KB");
        return text;
    }

    return std::to_wstring(bytes);
}
} // namespace PrefsCache

namespace PrefsConnections
{
const Common::Settings::ConnectionsSettings& GetConnectionsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::ConnectionsSettings kDefaults{};
    if (settings.connections.has_value())
    {
        return settings.connections.value();
    }
    return kDefaults;
}

Common::Settings::ConnectionsSettings* EnsureWorkingConnectionsSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.connections.has_value())
    {
        return &settings.connections.value();
    }

    settings.connections.emplace();
    return &settings.connections.value();
}

void MaybeResetWorkingConnectionsSettingsIfEmpty(Common::Settings::Settings& settings) noexcept
{
    if (! settings.connections.has_value())
    {
        return;
    }

    if (! settings.connections->items.empty())
    {
        return;
    }

    const Common::Settings::ConnectionsSettings defaults{};
    const bool wroteGlobals = settings.connections->bypassWindowsHello != defaults.bypassWindowsHello ||
                              settings.connections->windowsHelloReauthTimeoutMinute != defaults.windowsHelloReauthTimeoutMinute;
    if (! wroteGlobals)
    {
        settings.connections.reset();
    }
}
} // namespace PrefsConnections

namespace PrefsFileOperations
{
const Common::Settings::FileOperationsSettings& GetFileOperationsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::FileOperationsSettings kDefaults{};
    if (settings.fileOperations.has_value())
    {
        return settings.fileOperations.value();
    }
    return kDefaults;
}

Common::Settings::FileOperationsSettings* EnsureWorkingFileOperationsSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.fileOperations.has_value())
    {
        return &settings.fileOperations.value();
    }

    settings.fileOperations.emplace();
    return &settings.fileOperations.value();
}

void MaybeResetWorkingFileOperationsSettingsIfEmpty(Common::Settings::Settings& settings) noexcept
{
    if (! settings.fileOperations.has_value())
    {
        return;
    }

    const Common::Settings::FileOperationsSettings defaults{};
    const auto& fileOperations = settings.fileOperations.value();
    const bool hasNonDefault   = fileOperations.autoDismissSuccess != defaults.autoDismissSuccess ||
                               fileOperations.maxDiagnosticsLogFiles != defaults.maxDiagnosticsLogFiles ||
                               fileOperations.diagnosticsInfoEnabled != defaults.diagnosticsInfoEnabled ||
                               fileOperations.diagnosticsDebugEnabled != defaults.diagnosticsDebugEnabled || fileOperations.maxIssueReportFiles.has_value() ||
                               fileOperations.maxDiagnosticsInMemory.has_value() || fileOperations.maxDiagnosticsPerFlush.has_value() ||
                               fileOperations.diagnosticsFlushIntervalMs.has_value() || fileOperations.diagnosticsCleanupIntervalMs.has_value();

    if (! hasNonDefault)
    {
        settings.fileOperations.reset();
    }
}
} // namespace PrefsFileOperations
