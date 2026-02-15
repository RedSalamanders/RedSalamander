#include "Framework.h"

#include "ThemedControls.h"
#include "WindowMessages.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cwchar>
#include <cwctype>
#include <limits>
#include <memory>
#include <mutex>
#include <string>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

#include <commctrl.h>
#include <uxtheme.h>

namespace
{
constexpr UINT_PTR kThemedButtonHoverSubclassId = 1u;
constexpr wchar_t kThemedButtonHotProp[]        = L"ThemedControlsHot";

LRESULT CALLBACK ThemedButtonHoverSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept
{
    UNREFERENCED_PARAMETER(subclassId);
    UNREFERENCED_PARAMETER(refData);

    switch (msg)
    {
        case WM_MOUSEMOVE:
        {
            if (GetPropW(hwnd, kThemedButtonHotProp) == nullptr)
            {
                SetPropW(hwnd, kThemedButtonHotProp, reinterpret_cast<HANDLE>(1));

                TRACKMOUSEEVENT tme{};
                tme.cbSize    = sizeof(tme);
                tme.dwFlags   = TME_LEAVE;
                tme.hwndTrack = hwnd;
                TrackMouseEvent(&tme);

                InvalidateRect(hwnd, nullptr, TRUE);
            }
            break;
        }
        case WM_MOUSELEAVE:
            if (GetPropW(hwnd, kThemedButtonHotProp) != nullptr)
            {
                RemovePropW(hwnd, kThemedButtonHotProp);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            break;
        case WM_ENABLE:
            if (GetPropW(hwnd, kThemedButtonHotProp) != nullptr)
            {
                RemovePropW(hwnd, kThemedButtonHotProp);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            break;
        case WM_NCDESTROY: RemovePropW(hwnd, kThemedButtonHotProp);
#pragma warning(push)
#pragma warning(disable : 5039) // passing potentially-throwing callback to extern "C" Win32 API under -EHc
            RemoveWindowSubclass(hwnd, ThemedButtonHoverSubclassProc, kThemedButtonHoverSubclassId);
#pragma warning(pop)
            break;
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}
} // namespace

namespace ThemedControls
{
COLORREF BlendColor(COLORREF base, COLORREF overlay, int overlayWeight, int denom) noexcept
{
    if (denom <= 0)
    {
        return base;
    }

    overlayWeight        = std::clamp(overlayWeight, 0, denom);
    const int baseWeight = denom - overlayWeight;

    const int r = (static_cast<int>(GetRValue(base)) * baseWeight + static_cast<int>(GetRValue(overlay)) * overlayWeight) / denom;
    const int g = (static_cast<int>(GetGValue(base)) * baseWeight + static_cast<int>(GetGValue(overlay)) * overlayWeight) / denom;
    const int b = (static_cast<int>(GetBValue(base)) * baseWeight + static_cast<int>(GetBValue(overlay)) * overlayWeight) / denom;
    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}

int ScaleDip(UINT dpi, int dip) noexcept
{
    const int useDpi = (dpi > 0) ? static_cast<int>(dpi) : USER_DEFAULT_SCREEN_DPI;
    return (std::max)(0, MulDiv(dip, useDpi, USER_DEFAULT_SCREEN_DPI));
}

void EnableOwnerDrawButton(HWND dlg, int controlId) noexcept
{
    HWND button = dlg ? GetDlgItem(dlg, controlId) : nullptr;
    if (! button)
    {
        return;
    }

    LONG_PTR style = GetWindowLongPtrW(button, GWL_STYLE);
    if ((style & BS_OWNERDRAW) == BS_OWNERDRAW)
    {
        SetWindowSubclass(button, ThemedButtonHoverSubclassProc, kThemedButtonHoverSubclassId, 0);
        return;
    }

    style |= BS_OWNERDRAW;
    SetWindowLongPtrW(button, GWL_STYLE, style);
    SetWindowPos(button, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    InvalidateRect(button, nullptr, TRUE);

    SetWindowSubclass(button, ThemedButtonHoverSubclassProc, kThemedButtonHoverSubclassId, 0);
}

int MeasureTextWidth(HWND hwnd, HFONT font, std::wstring_view text) noexcept
{
    if (! hwnd || text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return 0;
    }

    auto hdc = wil::GetDC(hwnd);
    if (! hdc)
    {
        return 0;
    }

    HFONT useFont                 = font ? font : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), useFont);

    SIZE sz{};
    if (! GetTextExtentPoint32W(hdc.get(), text.data(), static_cast<int>(text.size()), &sz))
    {
        return 0;
    }

    return (std::max)(0l, sz.cx);
}

COLORREF GetControlSurfaceColor(const AppTheme& theme) noexcept
{
    if (theme.systemHighContrast)
    {
        return GetSysColor(COLOR_WINDOW);
    }

    const int weight = theme.dark ? 18 : 10;
    return BlendColor(theme.windowBackground, theme.menu.text, weight, 255);
}

void CenterEditTextVertically(HWND edit) noexcept
{
    if (! edit)
    {
        return;
    }

    RECT client{};
    if (GetClientRect(edit, &client) == 0)
    {
        return;
    }

    const int clientHeight = std::max(0l, client.bottom - client.top);
    if (clientHeight <= 0)
    {
        return;
    }

    RECT formatRect{};
    SendMessageW(edit, EM_GETRECT, 0, reinterpret_cast<LPARAM>(&formatRect));

    // Recompute vertical centering from the current client size each time (avoid drift across resizes).
    formatRect.top    = client.top;
    formatRect.bottom = client.bottom;

    const int availableHeight = std::max(0l, formatRect.bottom - formatRect.top);
    if (availableHeight <= 0)
    {
        return;
    }

    const auto hdc = wil::GetDC(edit);
    if (! hdc)
    {
        return;
    }

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(edit, WM_GETFONT, 0, 0));
    if (! font)
    {
        font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    [[maybe_unused]] auto restoreFont = wil::SelectObject(hdc.get(), font);

    TEXTMETRICW tm{};
    if (GetTextMetricsW(hdc.get(), &tm) == 0)
    {
        return;
    }

    const int lineHeight = std::clamp(static_cast<int>(tm.tmHeight), 1, availableHeight);
    if (lineHeight >= availableHeight)
    {
        SendMessageW(edit, EM_SETRECTNP, 0, reinterpret_cast<LPARAM>(&formatRect));
        InvalidateRect(edit, nullptr, FALSE);
        return;
    }

    const int desiredTop = client.top + (availableHeight - lineHeight) / 2;
    formatRect.top       = desiredTop;
    formatRect.bottom    = desiredTop + lineHeight;
    SendMessageW(edit, EM_SETRECTNP, 0, reinterpret_cast<LPARAM>(&formatRect));
    InvalidateRect(edit, nullptr, FALSE);
}

namespace
{
constexpr wchar_t kModernComboBoxClassName[]         = L"RedSalamanderModernComboBox";
constexpr wchar_t kModernComboPopupClassName[]       = L"RedSalamanderModernComboPopup";
constexpr UINT_PTR kModernComboListSubclassId        = 1u;
constexpr int kModernComboMaxVisibleItems            = 8;
constexpr UINT_PTR kModernComboTypeResetTimerId      = 1u;
constexpr UINT kModernComboTypeResetMs               = 1200u;
struct ModernComboItem
{
    std::wstring text;
    LPARAM data = 0;
};

struct ModernComboState
{
    ModernComboState() = default;

    ModernComboState(const ModernComboState&)            = delete;
    ModernComboState& operator=(const ModernComboState&) = delete;

    ModernComboState(ModernComboState&&)            = delete;
    ModernComboState& operator=(ModernComboState&&) = delete;

    const AppTheme* theme = nullptr;
    std::vector<ModernComboItem> items;

    int selectedIndex = -1;
    int openedIndex   = -1;

    int droppedWidthPx = 0;
    int itemHeightPx   = 0;

    wil::unique_hwnd popup;
    HWND list = nullptr;

    wil::unique_hbrush listBackgroundBrush;
    COLORREF listBackgroundColor = RGB(0, 0, 0);

    bool mouseDown          = false;
    bool pressedVisual      = false;
    bool buttonHot          = false;
    bool trackingMouseLeave = false;

    bool mouseSelectionArmed        = false;
    bool selectionChangedDuringDrop = false;
    bool closingPopup               = false;
    bool closeOutsideAccept         = true;
    bool dropDownPreferBelow        = false;
    int pinnedIndex                 = -1;
    bool compactMode                = false;
    bool useMiddleEllipsis          = false;

    HFONT font = nullptr;
    UINT dpi   = USER_DEFAULT_SCREEN_DPI;

    std::wstring typeBuffer;
    ULONGLONG lastTypeTick = 0;
};

[[nodiscard]] std::wstring MakeMiddleEllipsisText(HDC hdc, std::wstring_view text, int maxWidthPx) noexcept
{
    if (! hdc || text.empty() || maxWidthPx <= 0 || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    SIZE full{};
    if (GetTextExtentPoint32W(hdc, text.data(), static_cast<int>(text.size()), &full) != 0 && full.cx <= maxWidthPx)
    {
        return std::wstring(text);
    }

    constexpr wchar_t kEllipsis[] = L"\u2026";
    SIZE ellipsisSz{};
    if (GetTextExtentPoint32W(hdc, kEllipsis, 1, &ellipsisSz) == 0)
    {
        return {};
    }
    if (ellipsisSz.cx > maxWidthPx)
    {
        return {};
    }

    const bool looksLikePath =
        text.find(L'\\') != std::wstring_view::npos || text.find(L'/') != std::wstring_view::npos || text.find(L':') != std::wstring_view::npos;
    const double rightShare = looksLikePath ? 0.60 : 0.50;

    int low  = 0;
    int high = static_cast<int>(text.size());
    while (low < high)
    {
        const int kept      = (low + high + 1) / 2;
        const int rightKept = (kept > 0) ? std::clamp(static_cast<int>(std::ceil(static_cast<double>(kept) * rightShare)), 0, kept) : 0;
        const int leftKept  = kept - rightKept;

        std::wstring candidate;
        candidate.reserve(static_cast<size_t>(leftKept) + 1u + static_cast<size_t>(rightKept));
        candidate.append(text.substr(0, static_cast<size_t>(leftKept)));
        candidate.append(kEllipsis);
        if (rightKept > 0)
        {
            candidate.append(text.substr(text.size() - static_cast<size_t>(rightKept), static_cast<size_t>(rightKept)));
        }

        SIZE sz{};
        if (GetTextExtentPoint32W(hdc, candidate.data(), static_cast<int>(candidate.size()), &sz) != 0 && sz.cx <= maxWidthPx)
        {
            low = kept;
        }
        else
        {
            high = kept - 1;
        }
    }

    if (low <= 1)
    {
        return std::wstring(kEllipsis);
    }

    const int kept      = low;
    const int rightKept = std::clamp(static_cast<int>(std::ceil(static_cast<double>(kept) * rightShare)), 0, kept);
    const int leftKept  = kept - rightKept;

    std::wstring out;
    out.reserve(static_cast<size_t>(leftKept) + 1u + static_cast<size_t>(rightKept));
    out.append(text.substr(0, static_cast<size_t>(leftKept)));
    out.append(kEllipsis);
    if (rightKept > 0)
    {
        out.append(text.substr(text.size() - static_cast<size_t>(rightKept), static_cast<size_t>(rightKept)));
    }

    return out;
}

[[nodiscard]] bool IsModernComboItemBlank(std::wstring_view text) noexcept
{
    for (const wchar_t ch : text)
    {
        if (std::iswspace(static_cast<wint_t>(ch)) == 0)
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] ModernComboState* GetModernComboState(HWND hwnd) noexcept
{
    return reinterpret_cast<ModernComboState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
}

[[nodiscard]] HWND GetComboOwnerWindow(HWND combo) noexcept
{
    if (! combo)
    {
        return nullptr;
    }

    HWND root = GetAncestor(combo, GA_ROOT);
    return root ? root : GetParent(combo);
}

void NotifyCombo(HWND combo, UINT notifyCode) noexcept
{
    if (! combo)
    {
        return;
    }

    const HWND parent = GetParent(combo);
    if (! parent)
    {
        return;
    }

    const int id = GetDlgCtrlID(combo);
    SendMessageW(parent, WM_COMMAND, MAKEWPARAM(id, notifyCode), reinterpret_cast<LPARAM>(combo));
}

void EnsureModernComboItemHeight(HWND combo, ModernComboState& state) noexcept
{
    if (state.itemHeightPx > 0)
    {
        return;
    }

    state.dpi = GetDpiForWindow(combo);

    HFONT font = state.font;
    if (! font)
    {
        font = reinterpret_cast<HFONT>(SendMessageW(combo, WM_GETFONT, 0, 0));
    }
    if (! font)
    {
        font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    const auto hdc = wil::GetDC(combo);
    if (! hdc)
    {
        state.itemHeightPx = std::max(1, ScaleDip(state.dpi, 24));
        return;
    }

    [[maybe_unused]] auto restoreFont = wil::SelectObject(hdc.get(), font);

    TEXTMETRICW tm{};
    if (GetTextMetricsW(hdc.get(), &tm) == 0)
    {
        state.itemHeightPx = std::max(1, ScaleDip(state.dpi, 24));
        return;
    }

    const bool compact = state.compactMode;
    const int paddingY = ScaleDip(state.dpi, compact ? 6 : 8);
    const int minItemH = ScaleDip(state.dpi, compact ? 34 : 40);
    const int textH    = static_cast<int>(tm.tmHeight) + static_cast<int>(tm.tmExternalLeading);
    state.itemHeightPx = std::max(1, std::max(minItemH, textH + (2 * paddingY)));
}

void EnsureListBoxItemVisible(HWND list, int index, int itemHeightPx) noexcept
{
    if (! list || index < 0 || itemHeightPx <= 0)
    {
        return;
    }

    const LRESULT topIndexRes = SendMessageW(list, LB_GETTOPINDEX, 0, 0);
    if (topIndexRes == LB_ERR)
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(list, &client))
    {
        return;
    }

    const int clientHeight = std::max(0l, client.bottom - client.top);
    const int visible      = std::max(1, clientHeight / itemHeightPx);
    const int topIndex     = static_cast<int>(topIndexRes);

    int newTop = topIndex;
    if (index < topIndex)
    {
        newTop = index;
    }
    else if (index >= topIndex + visible)
    {
        newTop = index - visible + 1;
    }

    if (newTop != topIndex)
    {
        SendMessageW(list, LB_SETTOPINDEX, static_cast<WPARAM>(newTop), 0);
    }
}

void ModernComboCloseDropDown(HWND combo, ModernComboState& state, bool accept) noexcept;

void ModernComboSyncListSelection(HWND combo, ModernComboState& state) noexcept
{
    if (state.list && IsWindow(state.list))
    {
        const int index = state.selectedIndex;
        SendMessageW(state.list, LB_SETCURSEL, static_cast<WPARAM>(index), 0);
        EnsureListBoxItemVisible(state.list, index, state.itemHeightPx);
        InvalidateRect(state.list, nullptr, TRUE);

        if (HWND owner = GetParent(state.list))
        {
            InvalidateRect(owner, nullptr, TRUE);
        }
    }

    InvalidateRect(combo, nullptr, TRUE);
}

void ModernComboSetSelection(HWND combo, ModernComboState& state, int index, bool notify) noexcept
{
    const int count = static_cast<int>(state.items.size());
    if (count <= 0)
    {
        if (state.selectedIndex != -1)
        {
            state.selectedIndex = -1;
            ModernComboSyncListSelection(combo, state);
            if (notify)
            {
                NotifyCombo(combo, CBN_SELCHANGE);
            }
        }
        return;
    }

    if (index < 0)
    {
        // Comboboxes in this UI are dropdownlist-style and always keep a valid selection; ignore invalid indices.
        return;
    }

    int clamped = std::clamp(index, 0, count - 1);

    // Avoid selecting blank/whitespace-only items (treat as separators/non-selectable).
    if (IsModernComboItemBlank(state.items[static_cast<size_t>(clamped)].text))
    {
        int next = -1;
        for (int i = clamped + 1; i < count; ++i)
        {
            if (! IsModernComboItemBlank(state.items[static_cast<size_t>(i)].text))
            {
                next = i;
                break;
            }
        }
        if (next < 0)
        {
            for (int i = clamped - 1; i >= 0; --i)
            {
                if (! IsModernComboItemBlank(state.items[static_cast<size_t>(i)].text))
                {
                    next = i;
                    break;
                }
            }
        }
        if (next < 0)
        {
            return;
        }
        clamped = next;
    }

    if (clamped == state.selectedIndex)
    {
        if (state.list && IsWindow(state.list))
        {
            const int listSel = static_cast<int>(SendMessageW(state.list, LB_GETCURSEL, 0, 0));
            if (listSel != state.selectedIndex)
            {
                ModernComboSyncListSelection(combo, state);
            }
        }
        return;
    }

    state.selectedIndex = clamped;
    ModernComboSyncListSelection(combo, state);

    if (state.popup)
    {
        state.selectionChangedDuringDrop = true;
        notify                           = false;
    }

    if (notify)
    {
        NotifyCombo(combo, CBN_SELCHANGE);
    }
}

[[nodiscard]] bool ModernComboPointInWindow(HWND hwnd, POINT ptScreen) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    RECT rc{};
    if (! GetWindowRect(hwnd, &rc))
    {
        return false;
    }

    return PtInRect(&rc, ptScreen) != FALSE;
}

LRESULT CALLBACK ModernComboListSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR uIdSubclass, DWORD_PTR refData) noexcept
{
    const HWND combo = reinterpret_cast<HWND>(refData);
    auto* state      = combo ? GetModernComboState(combo) : nullptr;

    switch (msg)
    {
        case WM_MOUSEMOVE:
        {
            if (! combo || ! state || ! state->popup)
            {
                break;
            }

            const DWORD hit    = static_cast<DWORD>(SendMessageW(hwnd, LB_ITEMFROMPOINT, 0, lp));
            const int hitIndex = static_cast<int>(LOWORD(hit));
            const bool outside = HIWORD(hit) != 0;

            const int count = static_cast<int>(state->items.size());
            if (! outside && hitIndex >= 0 && hitIndex < count && ! IsModernComboItemBlank(state->items[static_cast<size_t>(hitIndex)].text))
            {
                ModernComboSetSelection(combo, *state, hitIndex, true);
            }

            break;
        }
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
        case WM_MBUTTONDOWN:
        {
            if (! combo || ! state || ! state->popup)
            {
                break;
            }

            state->mouseSelectionArmed = false;

            POINT pt{};
            pt.x = static_cast<short>(LOWORD(lp));
            pt.y = static_cast<short>(HIWORD(lp));
            ClientToScreen(hwnd, &pt);

            const bool inCombo = ModernComboPointInWindow(combo, pt);
            const bool inPopup = state->popup && ModernComboPointInWindow(state->popup.get(), pt);
            const bool inList  = ModernComboPointInWindow(hwnd, pt);

            if (! inPopup)
            {
                ModernComboCloseDropDown(combo, *state, state->closeOutsideAccept);
                return 0;
            }

            if (! inList)
            {
                if (inCombo)
                {
                    ModernComboCloseDropDown(combo, *state, true);
                    return 0;
                }
                return 0;
            }

            const DWORD hit    = static_cast<DWORD>(SendMessageW(hwnd, LB_ITEMFROMPOINT, 0, lp));
            const int hitIndex = static_cast<int>(LOWORD(hit));
            const bool outside = HIWORD(hit) != 0;
            const int count    = static_cast<int>(state->items.size());
            const bool selectable =
                ! outside && hitIndex >= 0 && hitIndex < count && ! IsModernComboItemBlank(state->items[static_cast<size_t>(hitIndex)].text);
            state->mouseSelectionArmed = selectable;
            if (selectable)
            {
                ModernComboSetSelection(combo, *state, hitIndex, true);
            }
            else if (! outside && hitIndex >= 0 && hitIndex < count && IsModernComboItemBlank(state->items[static_cast<size_t>(hitIndex)].text))
            {
                return 0;
            }
            break;
        }
        case WM_LBUTTONUP:
        {
            if (combo && state && state->mouseSelectionArmed)
            {
                state->mouseSelectionArmed = false;
                ModernComboCloseDropDown(combo, *state, true);
                return 0;
            }
            if (state)
            {
                state->mouseSelectionArmed = false;
            }
            break;
        }
        case WM_KEYDOWN:
        {
            if (! combo || ! state)
            {
                break;
            }

            if (wp == VK_ESCAPE)
            {
                ModernComboCloseDropDown(combo, *state, false);
                return 0;
            }
            if (wp == VK_RETURN)
            {
                ModernComboCloseDropDown(combo, *state, true);
                return 0;
            }
            break;
        }
        case WM_KILLFOCUS:
        {
            if (! combo || ! state || ! state->popup)
            {
                break;
            }

            const HWND newFocus = reinterpret_cast<HWND>(wp);
            const HWND popup    = state->popup.get();
            if (newFocus && (newFocus == combo || newFocus == popup || IsChild(popup, newFocus) != FALSE))
            {
                break;
            }

            ModernComboCloseDropDown(combo, *state, state->closeOutsideAccept);
            break;
        }
        case WM_CAPTURECHANGED:
        {
            if (combo && state && state->popup && ! state->mouseSelectionArmed)
            {
                ModernComboCloseDropDown(combo, *state, state->closeOutsideAccept);
            }
            break;
        }
        case WM_NCDESTROY:
        {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
            RemoveWindowSubclass(hwnd, ModernComboListSubclassProc, uIdSubclass);
#pragma warning(pop)
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

LRESULT CALLBACK ModernComboPopupWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    const HWND combo = reinterpret_cast<HWND>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    auto* state      = combo ? GetModernComboState(combo) : nullptr;

    switch (msg)
    {
        case WM_NCCREATE:
        {
            auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
            if (cs && cs->lpCreateParams)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(cs->lpCreateParams));
            }
            break;
        }
        case WM_MOUSEACTIVATE: return MA_NOACTIVATE;
        case WM_ERASEBKGND: return 1;
        case WM_CTLCOLORLISTBOX:
        {
            if (! state || ! state->theme)
            {
                break;
            }

            HDC hdc = reinterpret_cast<HDC>(wp);
            if (! hdc)
            {
                break;
            }

            COLORREF bg        = state->theme->systemHighContrast ? GetSysColor(COLOR_WINDOW) : state->theme->menu.background;
            COLORREF textColor = state->theme->menu.text;
            if (state->theme->systemHighContrast)
            {
                textColor = GetSysColor(COLOR_WINDOWTEXT);
            }

            SetBkColor(hdc, bg);
            SetTextColor(hdc, textColor);

            if (! state->listBackgroundBrush || state->listBackgroundColor != bg)
            {
                state->listBackgroundBrush.reset(CreateSolidBrush(bg));
                state->listBackgroundColor = bg;
            }

            if (state->listBackgroundBrush)
            {
                return reinterpret_cast<LRESULT>(state->listBackgroundBrush.get());
            }

            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
        }
        case WM_THEMECHANGED:
            InvalidateRect(hwnd, nullptr, TRUE);
            if (state && state->list)
            {
                InvalidateRect(state->list, nullptr, TRUE);
            }
            return 0;
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (! hdc || ! state || ! state->theme)
            {
                return 0;
            }

            RECT rc{};
            GetClientRect(hwnd, &rc);

            COLORREF surface = state->theme->systemHighContrast ? GetSysColor(COLOR_WINDOW) : state->theme->menu.background;

            wil::unique_hbrush bgBrush(CreateSolidBrush(surface));
            if (bgBrush)
            {
                FillRect(hdc.get(), &rc, bgBrush.get());
            }

            if (! state->theme->systemHighContrast)
            {
                const COLORREF border = BlendColor(surface, state->theme->menu.text, state->theme->dark ? 60 : 40, 255);
                wil::unique_hpen pen(CreatePen(PS_SOLID, 1, border));
                if (pen)
                {
                    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc.get(), GetStockObject(NULL_BRUSH));
                    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc.get(), pen.get());
                    const UINT dpi                 = GetDpiForWindow(hwnd);
                    const int radius               = ScaleDip(dpi, 8);
                    const int rightEdge            = std::max(rc.left + 1, rc.right - 1);
                    const int bottomEdge           = std::max(rc.top + 1, rc.bottom - 1);
                    RoundRect(hdc.get(), rc.left, rc.top, rightEdge, bottomEdge, radius, radius);
                }
            }

            return 0;
        }
        case WM_MEASUREITEM:
        {
            auto* mis = reinterpret_cast<MEASUREITEMSTRUCT*>(lp);
            if (! mis || ! state)
            {
                break;
            }

            EnsureModernComboItemHeight(combo, *state);
            mis->itemHeight = static_cast<UINT>(std::max(1, state->itemHeightPx));
            return TRUE;
        }
        case WM_DRAWITEM:
        {
            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis || ! state || ! state->theme)
            {
                break;
            }

            const int itemIndex = static_cast<int>(dis->itemID);
            if (itemIndex < 0 || itemIndex >= static_cast<int>(state->items.size()))
            {
                break;
            }
            const size_t itemIdx = static_cast<size_t>(itemIndex);

            const bool selected = (dis->itemState & ODS_SELECTED) != 0;
            const bool enabled  = (dis->itemState & ODS_DISABLED) == 0;

            const COLORREF surface = state->theme->systemHighContrast ? GetSysColor(COLOR_WINDOW) : state->theme->menu.background;
            COLORREF textColor     = enabled ? state->theme->menu.text : state->theme->menu.disabledText;
            if (state->theme->systemHighContrast)
            {
                textColor = enabled ? GetSysColor(COLOR_WINDOWTEXT) : GetSysColor(COLOR_GRAYTEXT);
            }
            COLORREF bg = surface;
            if (selected && state->theme->highContrast)
            {
                if (state->theme->systemHighContrast)
                {
                    bg        = GetSysColor(COLOR_HIGHLIGHT);
                    textColor = GetSysColor(COLOR_HIGHLIGHTTEXT);
                }
                else
                {
                    bg        = state->theme->menu.selectionBg;
                    textColor = state->theme->menu.selectionText;
                }
            }

            wil::unique_hbrush bgBrush(CreateSolidBrush(bg));
            if (bgBrush)
            {
                FillRect(dis->hDC, &dis->rcItem, bgBrush.get());
            }

            const UINT dpi            = GetDpiForWindow(hwnd);
            const bool compact        = state->compactMode;
            const int highlightInsetX = ScaleDip(dpi, compact ? 4 : 6);
            const int highlightInsetY = ScaleDip(dpi, compact ? 1 : 2);
            const int textInsetX      = ScaleDip(dpi, compact ? 18 : 22);

            RECT highlightRc = dis->rcItem;
            InflateRect(&highlightRc, -highlightInsetX, -highlightInsetY);

            if (selected && ! state->theme->highContrast)
            {
                const int highlightWeight = state->theme->dark ? 30 : 18;
                const COLORREF highlight  = BlendColor(surface, state->theme->menu.text, highlightWeight, 255);
                wil::unique_hbrush highlightBrush(CreateSolidBrush(highlight));
                if (highlightBrush)
                {
                    [[maybe_unused]] auto oldBrush = wil::SelectObject(dis->hDC, highlightBrush.get());
                    [[maybe_unused]] auto oldPen   = wil::SelectObject(dis->hDC, GetStockObject(NULL_PEN));
                    const int radius               = ScaleDip(dpi, compact ? 6 : 8);
                    RoundRect(dis->hDC, highlightRc.left, highlightRc.top, highlightRc.right, highlightRc.bottom, radius, radius);
                }
            }

            const bool hasPinnedIndex = state->pinnedIndex >= 0;
            const bool showCurrentBar = ! state->theme->highContrast && (hasPinnedIndex ? (state->pinnedIndex == itemIndex) : selected);
            if (showCurrentBar)
            {
                RECT barRc          = highlightRc;
                const int barWidth  = ScaleDip(dpi, 5);
                const int barInsetX = ScaleDip(dpi, compact ? 3 : 4);
                const int barInsetY = ScaleDip(dpi, compact ? 3 : 4);
                barRc.left          = std::min(barRc.right, barRc.left + barInsetX);
                barRc.right         = std::min(barRc.right, barRc.left + barWidth);
                barRc.top           = std::min(barRc.bottom, barRc.top + barInsetY);
                barRc.bottom        = std::max(barRc.top, barRc.bottom - barInsetY);

                wil::unique_hbrush accentBrush(CreateSolidBrush(state->theme->menu.selectionBg));
                if (accentBrush)
                {
                    [[maybe_unused]] auto oldBrush = wil::SelectObject(dis->hDC, accentBrush.get());
                    [[maybe_unused]] auto oldPen   = wil::SelectObject(dis->hDC, GetStockObject(NULL_PEN));
                    const int radius               = ScaleDip(dpi, compact ? 3 : 4);
                    RoundRect(dis->hDC, barRc.left, barRc.top, barRc.right, barRc.bottom, radius, radius);
                }
            }

            HFONT font = state->font;
            if (! font)
            {
                font = reinterpret_cast<HFONT>(SendMessageW(combo, WM_GETFONT, 0, 0));
            }
            if (! font)
            {
                font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            }

            [[maybe_unused]] auto oldFont = wil::SelectObject(dis->hDC, font);
            SetBkMode(dis->hDC, TRANSPARENT);
            SetTextColor(dis->hDC, textColor);

            RECT textRc = dis->rcItem;
            InflateRect(&textRc, -textInsetX, 0);
            if (state->useMiddleEllipsis && ! state->items[itemIdx].text.empty())
            {
                const int availableW        = std::max(0l, textRc.right - textRc.left);
                const std::wstring elided   = MakeMiddleEllipsisText(dis->hDC, state->items[itemIdx].text, availableW);
                const std::wstring_view str = elided.empty() ? std::wstring_view(state->items[itemIdx].text) : std::wstring_view(elided);
                DrawTextW(dis->hDC, str.data(), static_cast<int>(str.size()), &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
            }
            else
            {
                DrawTextW(dis->hDC,
                          state->items[itemIdx].text.c_str(),
                          static_cast<int>(state->items[itemIdx].text.size()),
                          &textRc,
                          DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
            }
            return TRUE;
        }
        case WM_COMMAND:
        {
            if (! state || reinterpret_cast<HWND>(lp) != state->list)
            {
                break;
            }

            const UINT notifyCode = HIWORD(wp);
            if (notifyCode == LBN_SELCHANGE)
            {
                if (! combo)
                {
                    return 0;
                }

                const int sel   = static_cast<int>(SendMessageW(state->list, LB_GETCURSEL, 0, 0));
                const int count = static_cast<int>(state->items.size());
                if (sel >= 0 && sel < count && ! IsModernComboItemBlank(state->items[static_cast<size_t>(sel)].text))
                {
                    ModernComboSetSelection(combo, *state, sel, true);
                }
                return 0;
            }
            break;
        }
        case WM_LBUTTONDOWN:
        {
            if (! combo || ! state)
            {
                break;
            }

            POINT pt{};
            pt.x = static_cast<short>(LOWORD(lp));
            pt.y = static_cast<short>(HIWORD(lp));
            ClientToScreen(hwnd, &pt);
            if (! ModernComboPointInWindow(hwnd, pt) && ! ModernComboPointInWindow(combo, pt))
            {
                ModernComboCloseDropDown(combo, *state, state->closeOutsideAccept);
                return 0;
            }
            break;
        }
        case WM_KILLFOCUS:
        {
            if (combo && state)
            {
                const HWND newFocus = reinterpret_cast<HWND>(wp);
                const HWND popup    = state->popup ? state->popup.get() : nullptr;
                if (newFocus && (newFocus == combo || newFocus == popup || (popup && IsChild(popup, newFocus) != FALSE)))
                {
                    break;
                }
                ModernComboCloseDropDown(combo, *state, state->closeOutsideAccept);
                return 0;
            }
            break;
        }
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void ModernComboOpenDropDown(HWND combo, ModernComboState& state) noexcept
{
    if (! combo || state.popup)
    {
        return;
    }

    const int count = static_cast<int>(state.items.size());
    if (count <= 0)
    {
        return;
    }

    int selectedIndex = state.selectedIndex;
    if (selectedIndex < 0 || selectedIndex >= count || IsModernComboItemBlank(state.items[static_cast<size_t>(selectedIndex)].text))
    {
        selectedIndex = -1;
        for (int i = 0; i < count; ++i)
        {
            if (! IsModernComboItemBlank(state.items[static_cast<size_t>(i)].text))
            {
                selectedIndex = i;
                break;
            }
        }
        if (selectedIndex < 0)
        {
            return;
        }
    }

    state.selectedIndex              = selectedIndex;
    state.openedIndex                = selectedIndex;
    state.mouseSelectionArmed        = false;
    state.selectionChangedDuringDrop = false;

    const HWND owner = GetComboOwnerWindow(combo);
    if (! owner)
    {
        return;
    }

    EnsureModernComboItemHeight(combo, state);

    RECT comboRc{};
    if (! GetWindowRect(combo, &comboRc))
    {
        return;
    }

    const UINT dpi         = GetDpiForWindow(combo);
    const int border       = std::max(1, ScaleDip(dpi, 1));
    const int popupExtraX  = std::max(0, ScaleDip(dpi, 6));
    const int itemHeightPx = std::max(1, state.itemHeightPx);

    const bool preferBelow = state.dropDownPreferBelow;

    int visibleItems = std::clamp(count, 1, preferBelow ? count : kModernComboMaxVisibleItems);

    RECT limitRc{};
    bool hasLimitRc = false;
    if (preferBelow)
    {
        RECT ownerClient{};
        if (GetClientRect(owner, &ownerClient))
        {
            POINT tl{ownerClient.left, ownerClient.top};
            POINT br{ownerClient.right, ownerClient.bottom};
            ClientToScreen(owner, &tl);
            ClientToScreen(owner, &br);
            limitRc    = RECT{tl.x, tl.y, br.x, br.y};
            hasLimitRc = true;
        }

        HMONITOR monitor = MonitorFromWindow(owner, MONITOR_DEFAULTTONEAREST);
        MONITORINFO mi{};
        mi.cbSize = sizeof(mi);
        if (monitor && GetMonitorInfoW(monitor, &mi))
        {
            const RECT work = mi.rcWork;
            if (hasLimitRc)
            {
                RECT intersect{};
                if (IntersectRect(&intersect, &limitRc, &work))
                {
                    limitRc = intersect;
                }
                else
                {
                    limitRc = work;
                }
            }
            else
            {
                limitRc    = work;
                hasLimitRc = true;
            }
        }

        if (hasLimitRc)
        {
            const int maxHeightBelow = std::max(0L, limitRc.bottom - comboRc.bottom);
            const int availableListH = std::max(0, maxHeightBelow - (2 * border));
            const int maxVisible     = availableListH / itemHeightPx;
            if (maxVisible > 0)
            {
                visibleItems = std::clamp(maxVisible, 1, count);
            }
        }
    }

    const int controlW   = std::max(0l, comboRc.right - comboRc.left);
    const int preferredW = MeasureComboBoxPreferredWidth(combo, dpi);
    const int baseW      = std::max(controlW, preferredW);
    const int listWidth  = (state.droppedWidthPx > 0) ? std::max(controlW, state.droppedWidthPx) : baseW;
    const int listHeight = std::max(1, visibleItems * itemHeightPx);
    const int width      = std::max(1, listWidth + (2 * border) + popupExtraX);
    const int height     = std::max(1, listHeight + (2 * border));

    RECT popupRc{};
    popupRc.left  = comboRc.left - (popupExtraX / 2);
    popupRc.right = popupRc.left + width;

    const int maxTopIndex = std::max(0, count - visibleItems);
    const int anchorRow   = (visibleItems - 1) / 2;
    const int topIndex    = std::clamp(selectedIndex - anchorRow, 0, maxTopIndex);
    const int rowIndex    = std::clamp(selectedIndex - topIndex, 0, std::max(0, visibleItems - 1));

    if (preferBelow && hasLimitRc)
    {
        const int limitW = std::max(0l, limitRc.right - limitRc.left);
        const int popupW = (limitW > 0) ? std::min(width, limitW) : width;

        int left      = static_cast<int>(popupRc.left);
        left          = std::clamp(left, static_cast<int>(limitRc.left), std::max(static_cast<int>(limitRc.left), static_cast<int>(limitRc.right) - popupW));
        popupRc.left  = left;
        popupRc.right = left + popupW;

        popupRc.top    = comboRc.bottom;
        popupRc.bottom = popupRc.top + height;
    }
    else
    {
        const int comboCenterY            = static_cast<int>(comboRc.top + ((comboRc.bottom - comboRc.top) / 2));
        const int selectedRowCenterOffset = border + (rowIndex * itemHeightPx) + (itemHeightPx / 2);

        popupRc.top    = comboCenterY - selectedRowCenterOffset;
        popupRc.bottom = popupRc.top + height;

        HMONITOR monitor = MonitorFromWindow(owner, MONITOR_DEFAULTTONEAREST);
        MONITORINFO mi{};
        mi.cbSize = sizeof(mi);
        if (monitor && GetMonitorInfoW(monitor, &mi))
        {
            const RECT work  = mi.rcWork;
            const int workW  = std::max(0l, work.right - work.left);
            const int popupW = (workW > 0) ? std::min(width, workW) : width;

            int left      = static_cast<int>(popupRc.left);
            left          = std::clamp(left, static_cast<int>(work.left), std::max(static_cast<int>(work.left), static_cast<int>(work.right) - popupW));
            popupRc.left  = left;
            popupRc.right = left + popupW;

            const int workH  = std::max(0l, work.bottom - work.top);
            const int popupH = (workH > 0) ? std::min(height, workH) : height;
            const int maxTop = std::max(static_cast<int>(work.top), static_cast<int>(work.bottom) - popupH);
            popupRc.top      = std::clamp(static_cast<int>(popupRc.top), static_cast<int>(work.top), maxTop);
            popupRc.bottom   = popupRc.top + popupH;
        }
    }

    state.popup.reset(CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
                                      kModernComboPopupClassName,
                                      L"",
                                      WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                      popupRc.left,
                                      popupRc.top,
                                      std::max(1l, popupRc.right - popupRc.left),
                                      std::max(1l, popupRc.bottom - popupRc.top),
                                      owner,
                                      nullptr,
                                      GetModuleHandleW(nullptr),
                                      combo));
    if (! state.popup)
    {
        return;
    }

    const int cornerRadius = ScaleDip(dpi, 8);
    const int popupW       = std::max(1, static_cast<int>(popupRc.right - popupRc.left));
    const int popupH       = std::max(1, static_cast<int>(popupRc.bottom - popupRc.top));
    wil::unique_hrgn popupRgn(CreateRoundRectRgn(0, 0, popupW + 1, popupH + 1, cornerRadius, cornerRadius));
    if (popupRgn)
    {
        SetWindowRgn(state.popup.get(), popupRgn.release(), TRUE);
    }

    DWORD listStyle = WS_CHILD | WS_VISIBLE | LBS_NOTIFY | LBS_OWNERDRAWFIXED | LBS_NOINTEGRALHEIGHT;
    if (count > visibleItems)
    {
        listStyle |= WS_VSCROLL;
    }

    state.list = CreateWindowExW(0,
                                 L"ListBox",
                                 L"",
                                 listStyle,
                                 border,
                                 border,
                                 std::max(1, static_cast<int>(popupRc.right - popupRc.left) - (2 * border)),
                                 std::max(1, static_cast<int>(popupRc.bottom - popupRc.top) - (2 * border)),
                                 state.popup.get(),
                                 nullptr,
                                 GetModuleHandleW(nullptr),
                                 nullptr);

    if (state.list)
    {
        if (state.theme && ! state.theme->highContrast)
        {
            // Ensure the scrollbar track uses the same background color as the menu/popup surface.
            InitializeFlatSB(state.list);
            const COLORREF bg = state.theme->menu.background;
            static_cast<void>(FlatSB_SetScrollProp(state.list, WSB_PROP_VBKGCOLOR, static_cast<INT_PTR>(bg), TRUE));
            static_cast<void>(FlatSB_SetScrollProp(state.list, WSB_PROP_HBKGCOLOR, static_cast<INT_PTR>(bg), TRUE));
            RedrawWindow(state.list, nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME);
        }

        if (state.theme)
        {
            if (state.theme->highContrast)
            {
                SetWindowTheme(state.list, L"", nullptr);
            }
            else
            {
                const wchar_t* listTheme = state.theme->dark ? L"DarkMode_Explorer" : L"Explorer";
                SetWindowTheme(state.list, listTheme, nullptr);
            }

            SendMessageW(state.list, WM_THEMECHANGED, 0, 0);
        }

        if (state.font)
        {
            SendMessageW(state.list, WM_SETFONT, reinterpret_cast<WPARAM>(state.font), FALSE);
        }

        SendMessageW(state.list, LB_SETITEMHEIGHT, 0, static_cast<LPARAM>(std::max(1, state.itemHeightPx)));

        for (int i = 0; i < count; ++i)
        {
            const auto& text  = state.items[static_cast<size_t>(i)].text;
            const LRESULT idx = SendMessageW(state.list, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(text.c_str()));
            if (idx != LB_ERR && idx != LB_ERRSPACE)
            {
                SendMessageW(state.list, LB_SETITEMDATA, static_cast<WPARAM>(idx), static_cast<LPARAM>(i));
            }
        }

        SendMessageW(state.list, LB_SETCURSEL, static_cast<WPARAM>(state.selectedIndex), 0);
        SendMessageW(state.list, LB_SETTOPINDEX, static_cast<WPARAM>(topIndex), 0);

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(state.list, ModernComboListSubclassProc, kModernComboListSubclassId, reinterpret_cast<DWORD_PTR>(combo));
#pragma warning(pop)
    }

    ShowWindow(state.popup.get(), SW_SHOWNOACTIVATE);
    if (state.list)
    {
        SetCapture(state.list);
    }
    NotifyCombo(combo, CBN_DROPDOWN);
    InvalidateRect(combo, nullptr, TRUE);
}

void ModernComboCloseDropDown(HWND combo, ModernComboState& state, bool accept) noexcept
{
    if (! combo || ! state.popup)
    {
        return;
    }

    if (state.closingPopup)
    {
        return;
    }

    state.closingPopup = true;
    auto resetClosing  = wil::scope_exit([&state]() noexcept { state.closingPopup = false; });

    state.mouseSelectionArmed = false;

    if (state.list && GetCapture() == state.list)
    {
        ReleaseCapture();
    }

    if (! accept)
    {
        ModernComboSetSelection(combo, state, state.openedIndex, true);
    }

    NotifyCombo(combo, accept ? CBN_SELENDOK : CBN_SELENDCANCEL);
    state.list = nullptr;
    state.popup.reset();
    NotifyCombo(combo, CBN_CLOSEUP);
    if (accept && state.selectionChangedDuringDrop && state.selectedIndex != state.openedIndex)
    {
        NotifyCombo(combo, CBN_SELCHANGE);
    }
    state.selectionChangedDuringDrop = false;
    InvalidateRect(combo, nullptr, TRUE);
}

void ModernComboResetTypeBuffer(HWND combo, ModernComboState& state) noexcept
{
    state.typeBuffer.clear();
    KillTimer(combo, kModernComboTypeResetTimerId);
}

void ModernComboHandleTypeChar(HWND combo, ModernComboState& state, wchar_t ch) noexcept
{
    if (ch < 0x20 || ch == 0x7F)
    {
        return;
    }

    const ULONGLONG now = GetTickCount64();
    if (! state.typeBuffer.empty() && now - state.lastTypeTick > kModernComboTypeResetMs)
    {
        state.typeBuffer.clear();
    }
    state.lastTypeTick = now;

    state.typeBuffer.push_back(static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))));

    const int count = static_cast<int>(state.items.size());
    if (count <= 0)
    {
        return;
    }

    const int start = (state.selectedIndex >= 0) ? (state.selectedIndex + 1) : 0;
    for (int offset = 0; offset < count; ++offset)
    {
        const int index              = (start + offset) % count;
        const std::wstring_view text = state.items[static_cast<size_t>(index)].text;
        if (text.size() < state.typeBuffer.size())
        {
            continue;
        }

        bool match = true;
        for (size_t i = 0; i < state.typeBuffer.size(); ++i)
        {
            if (std::towlower(static_cast<wint_t>(text[i])) != state.typeBuffer[i])
            {
                match = false;
                break;
            }
        }
        if (match)
        {
            ModernComboSetSelection(combo, state, index, true);
            break;
        }
    }

    SetTimer(combo, kModernComboTypeResetTimerId, kModernComboTypeResetMs, nullptr);
}

LRESULT CALLBACK ModernComboWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ModernComboState* state = GetModernComboState(hwnd);

    switch (msg)
    {
        case WM_NCCREATE:
        {
            auto* cs    = reinterpret_cast<CREATESTRUCTW*>(lp);
            auto init   = std::make_unique<ModernComboState>();
            init->theme = cs ? reinterpret_cast<const AppTheme*>(cs->lpCreateParams) : nullptr;
            init->dpi   = GetDpiForWindow(hwnd);
            init->font  = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(init.release()));
            return TRUE;
        }
        case WM_NCDESTROY:
        {
            std::unique_ptr<ModernComboState> owned(state);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            if (owned)
            {
                if (GetCapture() == hwnd)
                {
                    ReleaseCapture();
                }
                if (owned->list && GetCapture() == owned->list)
                {
                    ReleaseCapture();
                }
                owned->list = nullptr;
                owned->popup.reset();
                KillTimer(hwnd, kModernComboTypeResetTimerId);
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_ERASEBKGND: return 1;
        case WM_THEMECHANGED:
            if (state && state->popup)
            {
                SendMessageW(state->popup.get(), msg, wp, lp);
            }
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        case WndMsg::kModernComboSetCloseOutsideAccept:
            if (state)
            {
                state->closeOutsideAccept = wp != FALSE;
            }
            return 0;
        case WndMsg::kModernComboSetDropDownPreferBelow:
            if (state)
            {
                state->dropDownPreferBelow = wp != FALSE;
            }
            return 0;
        case WndMsg::kModernComboSetPinnedIndex:
            if (state)
            {
                state->pinnedIndex = static_cast<int>(lp);
            }
            return 0;
        case WndMsg::kModernComboSetCompactMode:
            if (state)
            {
                state->compactMode  = wp != FALSE;
                state->itemHeightPx = 0;
                EnsureModernComboItemHeight(hwnd, *state);
                InvalidateRect(hwnd, nullptr, TRUE);
                if (state->popup && state->list)
                {
                    SendMessageW(state->list, LB_SETITEMHEIGHT, 0, static_cast<LPARAM>(std::max(1, state->itemHeightPx)));
                    InvalidateRect(state->list, nullptr, TRUE);
                }
            }
            return 0;
        case WndMsg::kModernComboSetUseMiddleEllipsis:
            if (state)
            {
                state->useMiddleEllipsis = wp != FALSE;
                InvalidateRect(hwnd, nullptr, TRUE);
                if (state->popup && state->list)
                {
                    InvalidateRect(state->list, nullptr, TRUE);
                }
            }
            return 0;
        case WM_CANCELMODE:
            if (state)
            {
                if (GetCapture() == hwnd)
                {
                    ReleaseCapture();
                }
                state->mouseDown     = false;
                state->pressedVisual = false;
                state->buttonHot     = false;
                ModernComboCloseDropDown(hwnd, *state, state->closeOutsideAccept);
                return 0;
            }
            break;
        case WM_ENABLE:
            if (state && state->popup && wp == FALSE)
            {
                ModernComboCloseDropDown(hwnd, *state, state->closeOutsideAccept);
            }
            if (state && wp == FALSE)
            {
                if (GetCapture() == hwnd)
                {
                    ReleaseCapture();
                }
                state->mouseDown     = false;
                state->pressedVisual = false;
                state->buttonHot     = false;
            }
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        case WM_SETFONT:
            if (state)
            {
                state->font         = reinterpret_cast<HFONT>(wp);
                state->itemHeightPx = 0;
                EnsureModernComboItemHeight(hwnd, *state);
                InvalidateRect(hwnd, nullptr, TRUE);
                if (state->popup && state->list)
                {
                    SendMessageW(state->list, WM_SETFONT, wp, FALSE);
                    SendMessageW(state->list, LB_SETITEMHEIGHT, 0, static_cast<LPARAM>(std::max(1, state->itemHeightPx)));
                    InvalidateRect(state->list, nullptr, TRUE);
                }
            }
            return 0;
        case WM_GETFONT: return state ? reinterpret_cast<LRESULT>(state->font) : 0;
        case WM_TIMER:
            if (state && wp == kModernComboTypeResetTimerId)
            {
                ModernComboResetTypeBuffer(hwnd, *state);
                return 0;
            }
            break;
        case WM_LBUTTONDOWN:
            if (state)
            {
                if (IsWindowEnabled(hwnd) == FALSE)
                {
                    break;
                }

                SetFocus(hwnd);
                state->mouseDown     = true;
                state->pressedVisual = true;
                SetCapture(hwnd);
                InvalidateRect(hwnd, nullptr, TRUE);
                return 0;
            }
            break;
        case WM_LBUTTONUP:
            if (state && state->mouseDown)
            {
                if (GetCapture() == hwnd)
                {
                    ReleaseCapture();
                }

                state->mouseDown     = false;
                state->pressedVisual = false;
                InvalidateRect(hwnd, nullptr, TRUE);

                POINT pt{};
                pt.x = static_cast<short>(LOWORD(lp));
                pt.y = static_cast<short>(HIWORD(lp));

                RECT rc{};
                GetClientRect(hwnd, &rc);
                const bool inside = PtInRect(&rc, pt) != FALSE;
                if (! inside || IsWindowEnabled(hwnd) == FALSE)
                {
                    return 0;
                }

                if (state->popup)
                {
                    ModernComboCloseDropDown(hwnd, *state, true);
                }
                else
                {
                    ModernComboOpenDropDown(hwnd, *state);
                }
                return 0;
            }
            break;
        case WM_CAPTURECHANGED:
            if (state && state->mouseDown)
            {
                state->mouseDown     = false;
                state->pressedVisual = false;
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            break;
        case WM_MOUSEMOVE:
            if (state)
            {
                const UINT dpi = GetDpiForWindow(hwnd);

                POINT pt{};
                pt.x = static_cast<short>(LOWORD(lp));
                pt.y = static_cast<short>(HIWORD(lp));

                RECT rc{};
                GetClientRect(hwnd, &rc);

                const bool inControl = PtInRect(&rc, pt) != FALSE;
                if (state->mouseDown)
                {
                    if (state->pressedVisual != inControl)
                    {
                        state->pressedVisual = inControl;
                        InvalidateRect(hwnd, nullptr, TRUE);
                    }
                }

                const int arrowW = GetSystemMetricsForDpi(SM_CXVSCROLL, dpi);
                RECT arrowRc     = rc;
                arrowRc.left     = std::max(arrowRc.left, arrowRc.right - arrowW);

                const bool hot = PtInRect(&arrowRc, pt) != FALSE;
                if (state->buttonHot != hot)
                {
                    state->buttonHot = hot;
                    InvalidateRect(hwnd, nullptr, TRUE);
                }

                if (! state->trackingMouseLeave)
                {
                    TRACKMOUSEEVENT tme{};
                    tme.cbSize    = sizeof(tme);
                    tme.dwFlags   = TME_LEAVE;
                    tme.hwndTrack = hwnd;
                    TrackMouseEvent(&tme);
                    state->trackingMouseLeave = true;
                }

                return 0;
            }
            break;
        case WM_MOUSELEAVE:
            if (state)
            {
                state->trackingMouseLeave = false;
                if (state->buttonHot)
                {
                    state->buttonHot = false;
                    InvalidateRect(hwnd, nullptr, TRUE);
                }
                return 0;
            }
            break;
        case WM_MOUSEWHEEL:
            if (state && state->popup && state->list)
            {
                SendMessageW(state->list, msg, wp, lp);
                return 0;
            }
            break;
        case WM_KEYDOWN:
            if (state)
            {
                const bool dropped = state->popup != nullptr;
                const int count    = static_cast<int>(state->items.size());
                if (dropped && wp == VK_ESCAPE)
                {
                    ModernComboCloseDropDown(hwnd, *state, false);
                    return 0;
                }
                if (dropped && (wp == VK_RETURN || wp == VK_TAB))
                {
                    ModernComboCloseDropDown(hwnd, *state, true);
                    if (wp == VK_RETURN)
                    {
                        return 0;
                    }
                    break;
                }

                if (! dropped && wp == VK_RETURN)
                {
                    ModernComboOpenDropDown(hwnd, *state);
                    return 0;
                }

                if (! dropped && (wp == VK_F4 || (wp == VK_DOWN && (GetKeyState(VK_MENU) & 0x8000) != 0)))
                {
                    ModernComboOpenDropDown(hwnd, *state);
                    return 0;
                }

                if (count > 0 && (wp == VK_UP || wp == VK_DOWN))
                {
                    const int step = (wp == VK_UP) ? -1 : 1;

                    const auto findSelectableFrom = [&](int start, int direction) noexcept -> int
                    {
                        for (int i = start; i >= 0 && i < count; i += direction)
                        {
                            if (! IsModernComboItemBlank(state->items[static_cast<size_t>(i)].text))
                            {
                                return i;
                            }
                        }
                        return -1;
                    };

                    const int current = state->selectedIndex;
                    if (current < 0 || current >= count || IsModernComboItemBlank(state->items[static_cast<size_t>(current)].text))
                    {
                        const int fallback = (step < 0) ? findSelectableFrom(count - 1, -1) : findSelectableFrom(0, 1);
                        if (fallback >= 0)
                        {
                            ModernComboSetSelection(hwnd, *state, fallback, true);
                        }
                        return 0;
                    }

                    const int next = findSelectableFrom(current + step, step);
                    if (next >= 0)
                    {
                        ModernComboSetSelection(hwnd, *state, next, true);
                    }
                    else
                    {
                        ModernComboSyncListSelection(hwnd, *state);
                        InvalidateRect(hwnd, nullptr, TRUE);
                    }
                    return 0;
                }

                if (dropped && count > 0)
                {
                    if (wp == VK_HOME)
                    {
                        int next = 0;
                        while (next < count && IsModernComboItemBlank(state->items[static_cast<size_t>(next)].text))
                        {
                            ++next;
                        }
                        if (next < count)
                        {
                            ModernComboSetSelection(hwnd, *state, next, true);
                        }
                        return 0;
                    }
                    if (wp == VK_END)
                    {
                        int next = count - 1;
                        while (next >= 0 && IsModernComboItemBlank(state->items[static_cast<size_t>(next)].text))
                        {
                            --next;
                        }
                        if (next >= 0)
                        {
                            ModernComboSetSelection(hwnd, *state, next, true);
                        }
                        return 0;
                    }
                    if (wp == VK_PRIOR || wp == VK_NEXT)
                    {
                        EnsureModernComboItemHeight(hwnd, *state);
                        RECT rcList{};
                        if (state->list && GetClientRect(state->list, &rcList))
                        {
                            const int page  = std::max(1, std::max(0, static_cast<int>(rcList.bottom - rcList.top)) / std::max(1, state->itemHeightPx));
                            const int delta = (wp == VK_PRIOR) ? -page : page;
                            const int base  = (state->selectedIndex >= 0) ? state->selectedIndex : 0;
                            const int step  = (delta < 0) ? -1 : 1;
                            int next        = std::clamp(base + delta, 0, count - 1);
                            while (next >= 0 && next < count && IsModernComboItemBlank(state->items[static_cast<size_t>(next)].text))
                            {
                                next += step;
                            }
                            if (next >= 0 && next < count)
                            {
                                ModernComboSetSelection(hwnd, *state, next, true);
                            }
                            return 0;
                        }
                    }
                }

                if (wp == VK_CONTROL || wp == VK_LCONTROL || wp == VK_RCONTROL || wp == VK_SHIFT || wp == VK_LSHIFT || wp == VK_RSHIFT)
                {
                    InvalidateRect(hwnd, nullptr, TRUE);
                }
            }
            break;
        case WM_CHAR:
            if (state)
            {
                ModernComboHandleTypeChar(hwnd, *state, static_cast<wchar_t>(wp));
                return 0;
            }
            break;
        case WM_GETDLGCODE:
        {
            DWORD code = DLGC_WANTARROWS | DLGC_WANTCHARS;
            if (wp == VK_RETURN)
            {
                code |= DLGC_WANTMESSAGE;
            }
            if (state && state->popup && wp == VK_ESCAPE)
            {
                // Prevent the dialog manager from treating Enter/Esc as default/cancel while the dropdown is open.
                code |= DLGC_WANTMESSAGE;
            }
            return static_cast<LRESULT>(code);
        }
        case WM_KILLFOCUS:
            if (state)
            {
                const HWND newFocus = reinterpret_cast<HWND>(wp);
                const HWND popup    = state->popup ? state->popup.get() : nullptr;
                if (popup && newFocus && (newFocus == popup || IsChild(popup, newFocus) != FALSE))
                {
                    break;
                }
                ModernComboCloseDropDown(hwnd, *state, true);
            }
            break;
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (! hdc || ! state || ! state->theme)
            {
                return 0;
            }

            RECT rc{};
            GetClientRect(hwnd, &rc);

            const bool enabled = IsWindowEnabled(hwnd) != FALSE;
            COLORREF fill      = GetControlSurfaceColor(*state->theme);
            if (! enabled && ! state->theme->highContrast)
            {
                fill = BlendColor(state->theme->windowBackground, fill, state->theme->dark ? 70 : 40, 255);
            }

            wil::unique_hbrush fillBrush(CreateSolidBrush(fill));
            if (fillBrush)
            {
                FillRect(hdc.get(), &rc, fillBrush.get());
            }

            COLORREF textColor = enabled ? state->theme->menu.text : state->theme->menu.disabledText;
            if (state->theme->systemHighContrast)
            {
                textColor = enabled ? GetSysColor(COLOR_WINDOWTEXT) : GetSysColor(COLOR_GRAYTEXT);
            }

            HFONT font                    = state->font ? state->font : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), font);

            SetBkMode(hdc.get(), TRANSPARENT);
            SetTextColor(hdc.get(), textColor);

            const UINT dpi   = GetDpiForWindow(hwnd);
            const int padX   = ScaleDip(dpi, 8);
            const int arrowW = GetSystemMetricsForDpi(SM_CXVSCROLL, dpi);

            const bool dropped       = state->popup != nullptr;
            const bool buttonHot     = state->buttonHot;
            const bool buttonPressed = state->pressedVisual;

            RECT arrowRc = rc;
            arrowRc.left = std::max(arrowRc.left, arrowRc.right - arrowW);

            if (enabled && ! state->theme->highContrast)
            {
                COLORREF buttonFill = fill;
                bool paintButtonBg  = false;
                if (buttonPressed || dropped)
                {
                    paintButtonBg    = true;
                    const int weight = state->theme->dark ? 40 : 24;
                    buttonFill       = BlendColor(fill, state->theme->menu.text, weight, 255);
                }
                else if (buttonHot)
                {
                    paintButtonBg    = true;
                    const int weight = state->theme->dark ? 26 : 14;
                    buttonFill       = BlendColor(fill, state->theme->menu.text, weight, 255);
                }

                if (paintButtonBg)
                {
                    wil::unique_hbrush buttonBrush(CreateSolidBrush(buttonFill));
                    if (buttonBrush)
                    {
                        const int saved = SaveDC(hdc.get());
                        if (saved != 0)
                        {
                            auto restore = wil::scope_exit([&hdc, saved]() noexcept { RestoreDC(hdc.get(), saved); });
                            IntersectClipRect(hdc.get(), arrowRc.left, arrowRc.top, arrowRc.right, arrowRc.bottom);
                            [[maybe_unused]] auto oldBrush2 = wil::SelectObject(hdc.get(), buttonBrush.get());
                            [[maybe_unused]] auto oldPen2   = wil::SelectObject(hdc.get(), GetStockObject(NULL_PEN));
                            const int radius                = ScaleDip(dpi, 8);
                            RoundRect(hdc.get(), rc.left, rc.top, rc.right, rc.bottom, radius, radius);
                        }
                        else
                        {
                            FillRect(hdc.get(), &arrowRc, buttonBrush.get());
                        }
                    }
                }
            }

            const bool focused       = GetFocus() == hwnd;
            const bool focusViaMouse = GetPropW(hwnd, L"FocusViaMouse") != nullptr;
            if (focused && enabled && ! focusViaMouse && ! state->theme->highContrast)
            {
                const int barWidth  = ScaleDip(dpi, 5);
                const int barInsetX = ScaleDip(dpi, 4);
                const int barInsetY = ScaleDip(dpi, 4);
                RECT barRc          = rc;
                barRc.left          = std::min(barRc.right, barRc.left + barInsetX);
                barRc.right         = std::min(barRc.right, barRc.left + barWidth);
                barRc.top           = std::min(barRc.bottom, barRc.top + barInsetY);
                barRc.bottom        = std::max(barRc.top, barRc.bottom - barInsetY);

                wil::unique_hbrush accentBrush(CreateSolidBrush(state->theme->menu.selectionBg));
                if (accentBrush)
                {
                    [[maybe_unused]] auto oldBrush2 = wil::SelectObject(hdc.get(), accentBrush.get());
                    [[maybe_unused]] auto oldPen2   = wil::SelectObject(hdc.get(), GetStockObject(NULL_PEN));
                    const int radius                = std::min(barWidth, ScaleDip(dpi, 4));
                    RoundRect(hdc.get(), barRc.left, barRc.top, barRc.right, barRc.bottom, radius, radius);
                }
            }

            RECT textRc   = rc;
            int leftInset = 2 * padX;

            textRc.left  = std::min(textRc.right, textRc.left + leftInset);
            textRc.right = std::max(textRc.left, textRc.right - arrowW);

            std::wstring_view text;
            if (state->selectedIndex >= 0 && state->selectedIndex < static_cast<int>(state->items.size()))
            {
                text = state->items[static_cast<size_t>(state->selectedIndex)].text;
            }

            if (! text.empty())
            {
                DrawTextW(hdc.get(), text.data(), static_cast<int>(text.size()), &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
            }

            const int cx   = (arrowRc.left + arrowRc.right) / 2;
            const int cy   = (arrowRc.top + arrowRc.bottom) / 2;
            const int size = ScaleDip(dpi, 5);

            const int pressOffset = buttonPressed ? 1 : 0;
            POINT pts[3]          = {
                POINT{cx - size + pressOffset, cy - 1 + pressOffset},
                POINT{cx + size + pressOffset, cy - 1 + pressOffset},
                POINT{cx + pressOffset, cy + size + pressOffset},
            };

            wil::unique_hbrush arrowBrush(CreateSolidBrush(textColor));
            if (arrowBrush)
            {
                [[maybe_unused]] auto oldBrush2 = wil::SelectObject(hdc.get(), arrowBrush.get());
                [[maybe_unused]] auto oldPen2   = wil::SelectObject(hdc.get(), GetStockObject(NULL_PEN));
                Polygon(hdc.get(), pts, 3);
            }

            return 0;
        }
        case WM_GETTEXT:
        {
            if (! state || lp == 0)
            {
                return 0;
            }

            const int cch = static_cast<int>(wp);
            if (cch <= 0)
            {
                return 0;
            }

            std::wstring_view text;
            if (state->selectedIndex >= 0 && state->selectedIndex < static_cast<int>(state->items.size()))
            {
                text = state->items[static_cast<size_t>(state->selectedIndex)].text;
            }

            auto* out      = reinterpret_cast<wchar_t*>(lp);
            const int copy = std::min(static_cast<int>(text.size()), cch - 1);
            if (copy > 0)
            {
                std::wmemcpy(out, text.data(), static_cast<size_t>(copy));
            }
            out[copy] = L'\0';
            return copy;
        }
        case WM_GETTEXTLENGTH:
            if (state && state->selectedIndex >= 0 && state->selectedIndex < static_cast<int>(state->items.size()))
            {
                return static_cast<LRESULT>(state->items[static_cast<size_t>(state->selectedIndex)].text.size());
            }
            return 0;
        case CB_GETCOUNT: return state ? static_cast<LRESULT>(state->items.size()) : 0;
        case CB_RESETCONTENT:
            if (state)
            {
                state->items.clear();
                state->selectedIndex = -1;
                state->openedIndex   = -1;
                state->pinnedIndex   = -1;
                ModernComboResetTypeBuffer(hwnd, *state);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        case CB_ADDSTRING:
            if (state && lp)
            {
                state->items.push_back(ModernComboItem{std::wstring(reinterpret_cast<const wchar_t*>(lp)), 0});
                InvalidateRect(hwnd, nullptr, TRUE);
                return static_cast<LRESULT>(state->items.size() - 1);
            }
            return CB_ERR;
        case CB_GETLBTEXTLEN:
            if (state)
            {
                const int index = static_cast<int>(wp);
                if (index >= 0 && index < static_cast<int>(state->items.size()))
                {
                    return static_cast<LRESULT>(state->items[static_cast<size_t>(index)].text.size());
                }
            }
            return CB_ERR;
        case CB_GETLBTEXT:
            if (state && lp)
            {
                const int index = static_cast<int>(wp);
                if (index >= 0 && index < static_cast<int>(state->items.size()))
                {
                    const auto& text = state->items[static_cast<size_t>(index)].text;
                    wcscpy_s(reinterpret_cast<wchar_t*>(lp), text.size() + 1u, text.c_str());
                    return static_cast<LRESULT>(text.size());
                }
            }
            return CB_ERR;
        case CB_SETCURSEL:
            if (state)
            {
                ModernComboSetSelection(hwnd, *state, static_cast<int>(wp), false);
                return static_cast<LRESULT>(state->selectedIndex);
            }
            return CB_ERR;
        case CB_GETCURSEL: return state ? static_cast<LRESULT>(state->selectedIndex) : CB_ERR;
        case CB_SETITEMDATA:
            if (state)
            {
                const int index = static_cast<int>(wp);
                if (index >= 0 && index < static_cast<int>(state->items.size()))
                {
                    state->items[static_cast<size_t>(index)].data = static_cast<LPARAM>(lp);
                    return TRUE;
                }
            }
            return CB_ERR;
        case CB_GETITEMDATA:
            if (state)
            {
                const int index = static_cast<int>(wp);
                if (index >= 0 && index < static_cast<int>(state->items.size()))
                {
                    return static_cast<LRESULT>(state->items[static_cast<size_t>(index)].data);
                }
            }
            return CB_ERR;
        case CB_SETDROPPEDWIDTH:
            if (state)
            {
                state->droppedWidthPx = static_cast<int>(wp);
                return TRUE;
            }
            return FALSE;
        case CB_GETDROPPEDWIDTH: return state ? static_cast<LRESULT>(state->droppedWidthPx) : 0;
        case CB_GETDROPPEDSTATE: return (state && state->popup) ? TRUE : FALSE;
        case CB_SHOWDROPDOWN:
            if (state)
            {
                const bool show = wp != FALSE;
                if (show)
                {
                    ModernComboOpenDropDown(hwnd, *state);
                }
                else
                {
                    ModernComboCloseDropDown(hwnd, *state, true);
                }
                return 0;
            }
            break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void EnsureModernComboClassesRegistered() noexcept
{
    static std::once_flag s_once;
    std::call_once(s_once,
                   []() noexcept
                   {
                       WNDCLASSEXW wc{};
                       wc.cbSize        = sizeof(wc);
                       wc.style         = CS_HREDRAW | CS_VREDRAW;
                       wc.lpfnWndProc   = ModernComboWndProc;
                       wc.hInstance     = GetModuleHandleW(nullptr);
                       wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
                       wc.lpszClassName = kModernComboBoxClassName;
                       RegisterClassExW(&wc);

                       WNDCLASSEXW popup{};
                       popup.cbSize        = sizeof(popup);
                       popup.style         = CS_HREDRAW | CS_VREDRAW;
                       popup.lpfnWndProc   = ModernComboPopupWndProc;
                       popup.hInstance     = GetModuleHandleW(nullptr);
                       popup.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
                       popup.lpszClassName = kModernComboPopupClassName;
                       RegisterClassExW(&popup);
                   });
}

[[nodiscard]] bool IsModernComboClass(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    std::array<wchar_t, 64> name{};
    const int len = GetClassNameW(hwnd, name.data(), static_cast<int>(name.size()));
    if (len <= 0)
    {
        return false;
    }

    return _wcsicmp(name.data(), kModernComboBoxClassName) == 0;
}
} // namespace

HWND CreateModernComboBox(HWND parent, int controlId, const AppTheme* theme) noexcept
{
    EnsureModernComboClassesRegistered();

    return CreateWindowExW(0,
                           kModernComboBoxClassName,
                           L"",
                           WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                           0,
                           0,
                           10,
                           10,
                           parent,
                           reinterpret_cast<HMENU>(static_cast<INT_PTR>(controlId)),
                           GetModuleHandleW(nullptr),
                           const_cast<AppTheme*>(theme));
}

bool IsModernComboBox(HWND hwnd) noexcept
{
    return IsModernComboClass(hwnd);
}

void SetModernComboCloseOnOutsideAccept(HWND combo, bool accept) noexcept
{
    if (! combo || ! IsModernComboBox(combo))
    {
        return;
    }

    SendMessageW(combo, WndMsg::kModernComboSetCloseOutsideAccept, accept ? TRUE : FALSE, 0);
}

void SetModernComboDropDownPreferBelow(HWND combo, bool preferBelow) noexcept
{
    if (! combo || ! IsModernComboBox(combo))
    {
        return;
    }

    SendMessageW(combo, WndMsg::kModernComboSetDropDownPreferBelow, preferBelow ? TRUE : FALSE, 0);
}

void SetModernComboPinnedIndex(HWND combo, int index) noexcept
{
    if (! combo || ! IsModernComboBox(combo))
    {
        return;
    }

    SendMessageW(combo, WndMsg::kModernComboSetPinnedIndex, 0, static_cast<LPARAM>(index));
}

void SetModernComboCompactMode(HWND combo, bool compact) noexcept
{
    if (! combo || ! IsModernComboBox(combo))
    {
        return;
    }

    SendMessageW(combo, WndMsg::kModernComboSetCompactMode, compact ? TRUE : FALSE, 0);
}

void SetModernComboUseMiddleEllipsis(HWND combo, bool enable) noexcept
{
    if (! combo || ! IsModernComboBox(combo))
    {
        return;
    }

    SendMessageW(combo, WndMsg::kModernComboSetUseMiddleEllipsis, enable ? TRUE : FALSE, 0);
}

void ApplyThemeToComboBox(HWND combo, const AppTheme& theme) noexcept
{
    if (! combo)
    {
        return;
    }

    if (IsModernComboBox(combo))
    {
        SendMessageW(combo, WM_THEMECHANGED, 0, 0);
        InvalidateRect(combo, nullptr, TRUE);
        return;
    }

    if (theme.systemHighContrast)
    {
        SetWindowTheme(combo, L"", nullptr);
        SendMessageW(combo, WM_THEMECHANGED, 0, 0);
        return;
    }

    const bool darkBackground = ChooseContrastingTextColor(theme.windowBackground) == RGB(255, 255, 255);
    const wchar_t* fieldTheme = darkBackground ? L"DarkMode_CFD" : L"Explorer";
    const wchar_t* listTheme  = darkBackground ? L"DarkMode_Explorer" : L"Explorer";

    SetWindowTheme(combo, fieldTheme, nullptr);
    SendMessageW(combo, WM_THEMECHANGED, 0, 0);

    COMBOBOXINFO cbi{};
    cbi.cbSize = sizeof(cbi);
    if (GetComboBoxInfo(combo, &cbi))
    {
        if (cbi.hwndItem)
        {
            SetWindowTheme(cbi.hwndItem, fieldTheme, nullptr);
            SendMessageW(cbi.hwndItem, WM_THEMECHANGED, 0, 0);
            InvalidateRect(cbi.hwndItem, nullptr, TRUE);
        }
        if (cbi.hwndList)
        {
            SetWindowTheme(cbi.hwndList, listTheme, nullptr);
            SendMessageW(cbi.hwndList, WM_THEMECHANGED, 0, 0);
            InvalidateRect(cbi.hwndList, nullptr, TRUE);
        }
    }

    InvalidateRect(combo, nullptr, TRUE);
}

void ApplyThemeToComboBoxDropDown(HWND combo, const AppTheme& theme) noexcept
{
    if (! combo)
    {
        return;
    }

    if (IsModernComboBox(combo))
    {
        return;
    }

    COMBOBOXINFO cbi{};
    cbi.cbSize = sizeof(cbi);
    if (! GetComboBoxInfo(combo, &cbi))
    {
        return;
    }

    if (! cbi.hwndList)
    {
        return;
    }

    if (theme.systemHighContrast)
    {
        SetWindowTheme(cbi.hwndList, L"", nullptr);
        SendMessageW(cbi.hwndList, WM_THEMECHANGED, 0, 0);
        InvalidateRect(cbi.hwndList, nullptr, TRUE);
        return;
    }

    const bool darkBackground = ChooseContrastingTextColor(theme.windowBackground) == RGB(255, 255, 255);
    const wchar_t* listTheme  = darkBackground ? L"DarkMode_Explorer" : L"Explorer";

    SetWindowTheme(cbi.hwndList, listTheme, nullptr);
    SendMessageW(cbi.hwndList, WM_THEMECHANGED, 0, 0);
    InvalidateRect(cbi.hwndList, nullptr, TRUE);
}

void ApplyThemeToListView(HWND listView, const AppTheme& theme) noexcept
{
    if (! listView)
    {
        return;
    }

    const COLORREF background = theme.systemHighContrast ? GetSysColor(COLOR_WINDOW) : theme.windowBackground;
    const COLORREF textColor  = theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : theme.menu.text;

    ListView_SetBkColor(listView, background);
    ListView_SetTextBkColor(listView, background);
    ListView_SetTextColor(listView, textColor);

    if (theme.systemHighContrast)
    {
        SetWindowTheme(listView, L"", nullptr);
    }
    else
    {
        const wchar_t* listTheme = theme.dark ? L"DarkMode_Explorer" : L"Explorer";
        SetWindowTheme(listView, listTheme, nullptr);
        if (const HWND header = ListView_GetHeader(listView))
        {
            SetWindowTheme(header, listTheme, nullptr);
            SendMessageW(header, WM_THEMECHANGED, 0, 0);
        }
        if (const HWND tooltips = ListView_GetToolTips(listView))
        {
            SetWindowTheme(tooltips, listTheme, nullptr);
            SendMessageW(tooltips, WM_THEMECHANGED, 0, 0);
        }
    }

    EnsureListViewHeaderThemed(listView, theme);
    SendMessageW(listView, WM_THEMECHANGED, 0, 0);
    InvalidateRect(listView, nullptr, TRUE);
}

namespace
{
constexpr UINT_PTR kListViewHeaderSubclassId = 1u;

void PaintListViewHeader(HWND header, const AppTheme& theme) noexcept
{
    if (! header)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(header, &ps);
    if (! hdc)
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(header, &client))
    {
        return;
    }

    const HWND root         = GetAncestor(header, GA_ROOT);
    const bool windowActive = root && GetActiveWindow() == root;

    const COLORREF bg  = BlendColor(theme.windowBackground, theme.menu.separator, 1, 12);
    COLORREF textColor = windowActive ? theme.menu.headerText : theme.menu.headerTextDisabled;
    if (textColor == bg)
    {
        textColor = ChooseContrastingTextColor(bg);
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(bg));
    if (bgBrush)
    {
        FillRect(hdc.get(), &ps.rcPaint, bgBrush.get());
    }

    HFONT fontToUse = reinterpret_cast<HFONT>(SendMessageW(header, WM_GETFONT, 0, 0));
    if (! fontToUse)
    {
        fontToUse = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), fontToUse);

    SetBkMode(hdc.get(), TRANSPARENT);
    SetTextColor(hdc.get(), textColor);

    const int dpi      = GetDeviceCaps(hdc.get(), LOGPIXELSX);
    const int paddingX = MulDiv(8, dpi, USER_DEFAULT_SCREEN_DPI);

    const COLORREF lineColor = theme.menu.separator;
    wil::unique_hbrush lineBrush(CreateSolidBrush(lineColor));

    const int count = Header_GetItemCount(header);
    for (int i = 0; i < count; ++i)
    {
        RECT rc{};
        if (Header_GetItemRect(header, i, &rc) == FALSE)
        {
            continue;
        }

        if (! IntersectRect(&rc, &rc, &client))
        {
            continue;
        }

        wchar_t buf[128]{};
        HDITEMW item{};
        item.mask       = HDI_TEXT | HDI_FORMAT;
        item.pszText    = buf;
        item.cchTextMax = static_cast<int>(_countof(buf));
        if (Header_GetItem(header, i, &item) == FALSE)
        {
            continue;
        }

        RECT textRect  = rc;
        textRect.left  = (std::min)(textRect.right, textRect.left + paddingX);
        textRect.right = (std::max)(textRect.left, textRect.right - paddingX);

        UINT flags = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX;
        if ((item.fmt & HDF_RIGHT) != 0)
        {
            flags |= DT_RIGHT;
        }
        else if ((item.fmt & HDF_CENTER) != 0)
        {
            flags |= DT_CENTER;
        }
        else
        {
            flags |= DT_LEFT;
        }

        DrawTextW(hdc.get(), buf, static_cast<int>(std::wcslen(buf)), &textRect, flags);

        if (lineBrush)
        {
            RECT rightLine = rc;
            rightLine.left = (std::max)(rightLine.left, rightLine.right - 1);
            FillRect(hdc.get(), &rightLine, lineBrush.get());
        }
    }

    if (lineBrush)
    {
        RECT bottomLine = client;
        bottomLine.top  = (std::max)(bottomLine.top, bottomLine.bottom - 1);
        FillRect(hdc.get(), &bottomLine, lineBrush.get());
    }
}

LRESULT CALLBACK ListViewHeaderSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) noexcept
{
    const auto* theme = reinterpret_cast<const AppTheme*>(dwRefData);
    if (! theme)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: PaintListViewHeader(hwnd, *theme); return 0;
        case WM_NCDESTROY:
        {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
            RemoveWindowSubclass(hwnd, ListViewHeaderSubclassProc, uIdSubclass);
#pragma warning(pop)
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}
} // namespace

void EnsureListViewHeaderThemed(HWND listView, const AppTheme& theme) noexcept
{
    const HWND header = listView ? ListView_GetHeader(listView) : nullptr;
    if (! header)
    {
        return;
    }

    if (theme.highContrast)
    {
        SetWindowTheme(header, L"", nullptr);
        SendMessageW(header, WM_THEMECHANGED, 0, 0);
        RemoveWindowSubclass(header, ListViewHeaderSubclassProc, kListViewHeaderSubclassId);
        InvalidateRect(header, nullptr, TRUE);
        return;
    }

    const bool darkBackground = ChooseContrastingTextColor(theme.windowBackground) == RGB(255, 255, 255);
    const wchar_t* listTheme  = darkBackground ? L"DarkMode_Explorer" : L"Explorer";
    SetWindowTheme(header, listTheme, nullptr);
    SendMessageW(header, WM_THEMECHANGED, 0, 0);

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    SetWindowSubclass(header, ListViewHeaderSubclassProc, kListViewHeaderSubclassId, reinterpret_cast<DWORD_PTR>(&theme));
#pragma warning(pop)
    InvalidateRect(header, nullptr, TRUE);
}

int MeasureComboBoxPreferredWidth(HWND combo, UINT dpi) noexcept
{
    if (! combo)
    {
        return 0;
    }

    const LRESULT count = SendMessageW(combo, CB_GETCOUNT, 0, 0);
    if (count == CB_ERR || count <= 0)
    {
        return 0;
    }

    const HFONT font = reinterpret_cast<HFONT>(SendMessageW(combo, WM_GETFONT, 0, 0));

    int maxTextWidth = 0;
    for (int i = 0; i < static_cast<int>(count); ++i)
    {
        const LRESULT len = SendMessageW(combo, CB_GETLBTEXTLEN, static_cast<WPARAM>(i), 0);
        if (len == CB_ERR || len <= 0 || len > static_cast<LRESULT>(std::numeric_limits<int>::max() - 1))
        {
            continue;
        }

        std::wstring text;
        text.resize(static_cast<size_t>(len) + 1u);

        const LRESULT copied = SendMessageW(combo, CB_GETLBTEXT, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(text.data()));
        if (copied == CB_ERR)
        {
            continue;
        }

        text.resize(static_cast<size_t>(std::wcslen(text.c_str())));
        maxTextWidth = (std::max)(maxTextWidth, MeasureTextWidth(combo, font, text));
    }

    const int paddingX = ScaleDip(dpi, 10);
    const int arrowW   = GetSystemMetricsForDpi(SM_CXVSCROLL, dpi);
    const int extra    = ScaleDip(dpi, 12);

    return (std::max)(0, maxTextWidth + (2 * paddingX) + arrowW + extra);
}

void EnsureComboBoxDroppedWidth(HWND combo, UINT dpi) noexcept
{
    if (! combo)
    {
        return;
    }

    const int preferred = MeasureComboBoxPreferredWidth(combo, dpi);

    RECT rc{};
    if (! GetWindowRect(combo, &rc))
    {
        return;
    }

    const int controlW = (std::max)(0l, rc.right - rc.left);
    const int droppedW = (std::max)(controlW, preferred);
    if (droppedW <= 0)
    {
        return;
    }

    SendMessageW(combo, CB_SETDROPPEDWIDTH, static_cast<WPARAM>(droppedW), 0);
}

void DrawThemedPushButton(const DRAWITEMSTRUCT& dis, const AppTheme& theme) noexcept
{
    if (! dis.hwndItem || ! dis.hDC)
    {
        return;
    }

    const UINT dpi          = GetDpiForWindow(dis.hwndItem);
    const int cornerRadius  = ScaleDip(dpi, 4);
    const int borderInsetPx = ScaleDip(dpi, 1);

    const bool enabled   = (dis.itemState & ODS_DISABLED) == 0;
    const bool pressed   = (dis.itemState & ODS_SELECTED) != 0;
    const bool isDefault = (dis.itemState & ODS_DEFAULT) != 0;
    const bool focused   = (dis.itemState & ODS_FOCUS) != 0;
    const bool hot       = (dis.itemState & ODS_HOTLIGHT) != 0 || GetPropW(dis.hwndItem, kThemedButtonHotProp) != nullptr;

    const int controlId  = GetDlgCtrlID(dis.hwndItem);
    const bool isPrimary = enabled && (controlId == IDOK || isDefault);

    const COLORREF surface = GetControlSurfaceColor(theme);

    // Modern flat design - subtle border only for default buttons or when focused
    COLORREF fill = surface;
    if (isPrimary)
    {
        fill = BlendColor(surface, theme.menu.selectionBg, theme.dark ? 110 : 90, 255);
    }
    if (hot && enabled && ! pressed)
    {
        fill = BlendColor(fill, theme.menu.text, theme.dark ? 18 : 12, 255);
    }
    if (pressed)
    {
        fill = BlendColor(fill, theme.menu.text, theme.dark ? 24 : 16, 255);
    }
    if (! enabled)
    {
        fill = BlendColor(theme.windowBackground, surface, theme.dark ? 70 : 40, 255);
    }

    COLORREF textColor = theme.menu.text;
    if (! enabled)
    {
        textColor = theme.menu.disabledText;
    }
    else if (isPrimary)
    {
        textColor = ChooseContrastingTextColor(fill);
    }

    RECT rc = dis.rcItem;

    // Fill the full rect to avoid dirty corners around the rounded button. Prefer the parent's CTLCOLORBTN
    // brush so buttons inside cards blend correctly with their backdrop.
    COLORREF backdrop    = theme.windowBackground;
    HBRUSH backdropBrush = nullptr;
    if (const HWND parent = GetParent(dis.hwndItem))
    {
        const LRESULT lr = SendMessageW(parent, WM_CTLCOLORBTN, reinterpret_cast<WPARAM>(dis.hDC), reinterpret_cast<LPARAM>(dis.hwndItem));
        if (lr != 0)
        {
            const HBRUSH candidateBrush = reinterpret_cast<HBRUSH>(lr);

            const bool isDefaultBrush = candidateBrush == GetSysColorBrush(COLOR_BTNFACE) || candidateBrush == GetSysColorBrush(COLOR_3DFACE) ||
                                        candidateBrush == GetSysColorBrush(COLOR_WINDOW) || candidateBrush == GetSysColorBrush(COLOR_MENU) ||
                                        candidateBrush == GetStockObject(WHITE_BRUSH) || candidateBrush == GetStockObject(LTGRAY_BRUSH);

            if (! isDefaultBrush)
            {
                backdropBrush = candidateBrush;
                backdrop      = GetBkColor(dis.hDC);
            }
        }
    }

    if (! backdropBrush)
    {
        wil::unique_hbrush backgroundBrush(CreateSolidBrush(backdrop));
        if (backgroundBrush)
        {
            FillRect(dis.hDC, &rc, backgroundBrush.get());
        }
    }
    else
    {
        if (backdropBrush == GetStockObject(DC_BRUSH))
        {
            SetDCBrushColor(dis.hDC, backdrop);
        }
        FillRect(dis.hDC, &rc, backdropBrush);
    }

    wil::unique_hbrush brush(CreateSolidBrush(fill));
    if (! brush)
    {
        return;
    }

    // Flat modern style - no visible border for regular buttons
    const bool showBorder = isPrimary || focused || hot;
    if (showBorder)
    {
        const COLORREF border =
            isPrimary ? theme.menu.selectionBg : BlendColor(theme.windowBackground, theme.menu.text, theme.dark ? (hot ? 70 : 50) : (hot ? 45 : 35), 255);
        wil::unique_hpen pen(CreatePen(PS_SOLID, 1, border));
        if (pen)
        {
            [[maybe_unused]] auto oldBrush = wil::SelectObject(dis.hDC, brush.get());
            [[maybe_unused]] auto oldPen   = wil::SelectObject(dis.hDC, pen.get());
            RoundRect(dis.hDC, rc.left, rc.top, rc.right, rc.bottom, cornerRadius, cornerRadius);
        }
    }
    else
    {
        [[maybe_unused]] auto oldBrush = wil::SelectObject(dis.hDC, brush.get());
        [[maybe_unused]] auto oldPen   = wil::SelectObject(dis.hDC, GetStockObject(NULL_PEN));
        RoundRect(dis.hDC, rc.left, rc.top, rc.right, rc.bottom, cornerRadius, cornerRadius);
    }

    std::wstring label;
    const int length = GetWindowTextLengthW(dis.hwndItem);
    if (length > 0)
    {
        label.resize(static_cast<size_t>(length) + 1u);
        GetWindowTextW(dis.hwndItem, label.data(), length + 1);
        label.resize(static_cast<size_t>(length));
    }

    HFONT font                    = reinterpret_cast<HFONT>(SendMessageW(dis.hwndItem, WM_GETFONT, 0, 0));
    [[maybe_unused]] auto oldFont = wil::SelectObject(dis.hDC, font);

    SetBkMode(dis.hDC, TRANSPARENT);
    SetTextColor(dis.hDC, textColor);

    RECT textRc = rc;
    InflateRect(&textRc, -ScaleDip(dpi, 10), -ScaleDip(dpi, 4));
    if (pressed)
    {
        OffsetRect(&textRc, 1, 1);
    }

    DrawTextW(dis.hDC, label.c_str(), static_cast<int>(label.size()), &textRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    if (focused && enabled)
    {
        RECT focusRc = rc;
        InflateRect(&focusRc, -borderInsetPx - ScaleDip(dpi, 2), -borderInsetPx - ScaleDip(dpi, 2));
        const COLORREF focusColor = BlendColor(fill, theme.menu.selectionBg, theme.dark ? 70 : 55, 255);
        wil::unique_hpen focusPen(CreatePen(PS_SOLID, 1, focusColor));
        if (focusPen)
        {
            [[maybe_unused]] auto oldBrush2 = wil::SelectObject(dis.hDC, GetStockObject(NULL_BRUSH));
            [[maybe_unused]] auto oldPen2   = wil::SelectObject(dis.hDC, focusPen.get());
            const int radius                = (std::max)(1, cornerRadius - ScaleDip(dpi, 1));
            RoundRect(dis.hDC, focusRc.left, focusRc.top, focusRc.right, focusRc.bottom, radius, radius);
        }
    }
}

void DrawThemedSwitchToggle(const DRAWITEMSTRUCT& dis,
                            const AppTheme& theme,
                            COLORREF surface,
                            HFONT boldFont,
                            std::wstring_view onLabel,
                            std::wstring_view offLabel,
                            bool toggledOn) noexcept
{
    if (! dis.hwndItem || ! dis.hDC)
    {
        return;
    }

    const UINT dpi     = GetDpiForWindow(dis.hwndItem);
    const int paddingX = ScaleDip(dpi, 6);
    const int paddingY = ScaleDip(dpi, 4);
    const int gapX     = ScaleDip(dpi, 8);

    const bool enabled       = (dis.itemState & ODS_DISABLED) == 0;
    const bool focused       = (dis.itemState & ODS_FOCUS) != 0;
    const bool focusViaMouse = GetPropW(dis.hwndItem, L"FocusViaMouse") != nullptr;
    const bool showFocus     = focused && enabled && ! focusViaMouse;

    const COLORREF accent = theme.menu.selectionBg;

    RECT rc = dis.rcItem;

    wil::unique_hbrush bgBrush(CreateSolidBrush(surface));
    if (bgBrush)
    {
        FillRect(dis.hDC, &rc, bgBrush.get());
    }

    HFONT baseFont = reinterpret_cast<HFONT>(SendMessageW(dis.hwndItem, WM_GETFONT, 0, 0));
    if (! baseFont)
    {
        baseFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    if (! boldFont)
    {
        boldFont = baseFont;
    }

    const std::wstring_view stateText = toggledOn ? onLabel : offLabel;

    SetBkMode(dis.hDC, TRANSPARENT);
    SetTextColor(dis.hDC, enabled ? theme.menu.text : theme.menu.disabledText);

    RECT contentRc = rc;
    InflateRect(&contentRc, -paddingX, -paddingY);

    const int trackHeight = ScaleDip(dpi, 18);
    const int trackWidth  = ScaleDip(dpi, 34);

    const int stateTextWidth = MeasureTextWidth(dis.hwndItem, boldFont, stateText);
    const int groupWidth     = (std::max)(0, stateTextWidth) + gapX + trackWidth;

    const int contentLeft  = static_cast<int>(contentRc.left);
    const int contentRight = static_cast<int>(contentRc.right);

    const int groupRight = contentRight;
    const int groupLeft  = (std::max)(contentLeft, groupRight - groupWidth);

    RECT trackRc{};
    trackRc.right  = groupRight;
    trackRc.left   = (std::max)(contentRc.left, trackRc.right - trackWidth);
    trackRc.top    = contentRc.top + ((std::max)(0l, contentRc.bottom - contentRc.top) - trackHeight) / 2;
    trackRc.bottom = trackRc.top + trackHeight;

    RECT textRc  = contentRc;
    textRc.left  = groupLeft;
    textRc.right = (std::max)(textRc.left, trackRc.left - gapX);

    [[maybe_unused]] auto oldFont = wil::SelectObject(dis.hDC, boldFont);
    DrawTextW(dis.hDC, stateText.data(), static_cast<int>(stateText.size()), &textRc, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);

    COLORREF trackFill   = toggledOn ? accent : BlendColor(surface, theme.menu.text, theme.dark ? 40 : 30, 255);
    COLORREF trackBorder = theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : BlendColor(surface, theme.menu.text, theme.dark ? 90 : 70, 255);

    if (! enabled && ! theme.highContrast)
    {
        trackFill   = BlendColor(surface, trackFill, theme.dark ? 130 : 110, 255);
        trackBorder = BlendColor(surface, trackBorder, theme.dark ? 130 : 110, 255);
    }

    wil::unique_hbrush trackBrush(CreateSolidBrush(trackFill));
    wil::unique_hpen trackPen(CreatePen(PS_SOLID, 1, trackBorder));
    if (trackBrush && trackPen)
    {
        [[maybe_unused]] auto oldBrush = wil::SelectObject(dis.hDC, trackBrush.get());
        [[maybe_unused]] auto oldPen   = wil::SelectObject(dis.hDC, trackPen.get());
        RoundRect(dis.hDC, trackRc.left, trackRc.top, trackRc.right, trackRc.bottom, trackHeight, trackHeight);
    }

    const int knobInset = ScaleDip(dpi, 2);
    const int knobSize  = (std::max)(1, trackHeight - 2 * knobInset);
    const int knobX     = toggledOn ? (trackRc.right - knobInset - knobSize) : (trackRc.left + knobInset);
    const int knobY     = trackRc.top + knobInset;

    COLORREF knobFill = RGB(0, 0, 0);
    if (! enabled && ! theme.highContrast)
    {
        knobFill = BlendColor(trackFill, knobFill, theme.dark ? 120 : 100, 255);
    }

    const COLORREF knobBorder = theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : BlendColor(trackFill, theme.menu.text, theme.dark ? 60 : 45, 255);

    wil::unique_hbrush knobBrush(CreateSolidBrush(knobFill));
    wil::unique_hpen knobPen(CreatePen(PS_SOLID, 1, knobBorder));
    if (knobBrush && knobPen)
    {
        [[maybe_unused]] auto oldBrush = wil::SelectObject(dis.hDC, knobBrush.get());
        [[maybe_unused]] auto oldPen   = wil::SelectObject(dis.hDC, knobPen.get());
        Ellipse(dis.hDC, knobX, knobY, knobX + knobSize, knobY + knobSize);
    }

    if (showFocus)
    {
        RECT focusRc = rc;
        InflateRect(&focusRc, -ScaleDip(dpi, 2), -ScaleDip(dpi, 2));

        const COLORREF focusColor = BlendColor(surface, accent, theme.dark ? 55 : 45, 255);
        wil::unique_hpen focusPen(CreatePen(PS_SOLID, 1, focusColor));
        if (focusPen)
        {
            [[maybe_unused]] auto oldBrush = wil::SelectObject(dis.hDC, GetStockObject(NULL_BRUSH));
            [[maybe_unused]] auto oldPen   = wil::SelectObject(dis.hDC, focusPen.get());

            const int radius = ScaleDip(dpi, 4);
            RoundRect(dis.hDC, focusRc.left, focusRc.top, focusRc.right, focusRc.bottom, radius, radius);
        }
    }
}

void ApplyModernEditStyle(HWND edit, [[maybe_unused]] const AppTheme& theme) noexcept
{
    if (! edit)
    {
        return;
    }

    // Remove the 3D border effect (WS_EX_CLIENTEDGE)
    const LONG_PTR exStyle = GetWindowLongPtrW(edit, GWL_EXSTYLE);
    SetWindowLongPtrW(edit, GWL_EXSTYLE, exStyle & ~WS_EX_CLIENTEDGE);

    // Set margins for better text padding (creates internal spacing)
    const UINT dpi             = GetDpiForWindow(edit);
    const int horizontalMargin = ScaleDip(dpi, 8);

    SendMessageW(edit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(horizontalMargin, horizontalMargin));

    // Increase height slightly for better vertical centering appearance
    RECT rc{};
    if (GetWindowRect(edit, &rc))
    {
        const int width           = rc.right - rc.left;
        const int currentHeight   = rc.bottom - rc.top;
        const int preferredHeight = ScaleDip(dpi, 28); // Modern standard height

        if (currentHeight < preferredHeight)
        {
            const HWND parent = GetParent(edit);
            if (parent)
            {
                POINT pt{rc.left, rc.top};
                ScreenToClient(parent, &pt);
                SetWindowPos(edit, nullptr, pt.x, pt.y, width, preferredHeight, SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOMOVE);
            }
        }
    }

    // Force redraw with new styles
    SetWindowPos(edit, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
    InvalidateRect(edit, nullptr, TRUE);
}

void ApplyModernComboStyle(HWND combo, const AppTheme& theme) noexcept
{
    if (! combo)
    {
        return;
    }

    // Apply base theme
    ApplyThemeToComboBox(combo, theme);

    // Remove the 3D border effect (WS_EX_CLIENTEDGE)
    const LONG_PTR exStyle = GetWindowLongPtrW(combo, GWL_EXSTYLE);
    SetWindowLongPtrW(combo, GWL_EXSTYLE, exStyle & ~WS_EX_CLIENTEDGE);

    // Force redraw with new styles
    SetWindowPos(combo, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
    InvalidateRect(combo, nullptr, TRUE);
}
} // namespace ThemedControls
