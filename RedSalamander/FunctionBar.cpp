#include "FunctionBar.h"

#include <algorithm>
#include <cwchar>
#include <windowsx.h>

#include "CommandRegistry.h"
#include "Helpers.h"
#include "ShortcutManager.h"
#include "WindowMessages.h"
#include "resource.h"

namespace
{
constexpr int kFunctionKeyCount         = 12;
constexpr int kZonePaddingXDip          = 6;
constexpr int kZonePaddingYDip          = 2;
constexpr int kGlyphPaddingXDip         = 3;
constexpr int kGlyphPaddingYDip         = 0;
constexpr int kGlyphCornerRadiusDip     = 2;
constexpr int kMinZoneWidthForModifiers = 50;
constexpr int kKeyFontHeightDip         = 7;
constexpr int kTextFontHeightDip        = 11;
constexpr int kModifiersGapDip          = 6;

[[nodiscard]] COLORREF BlendColor(COLORREF base, COLORREF overlay, int overlayWeight, int denom) noexcept
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

[[nodiscard]] std::wstring BuildModifierText(uint32_t modifiers) noexcept
{
    std::wstring result;

    auto append = [&](UINT stringId) noexcept
    {
        const std::wstring text = LoadStringResource(nullptr, stringId);
        if (text.empty())
        {
            return;
        }

        if (! result.empty())
        {
            result.push_back(L'+');
        }

        result.append(text);
    };

    if ((modifiers & ShortcutManager::kModCtrl) != 0)
    {
        append(IDS_MOD_CTRL);
    }
    if ((modifiers & ShortcutManager::kModAlt) != 0)
    {
        append(IDS_MOD_ALT);
    }
    if ((modifiers & ShortcutManager::kModShift) != 0)
    {
        append(IDS_MOD_SHIFT);
    }

    return result;
}

void DrawTextEllipsis(HDC hdc, const std::wstring& text, RECT rect, UINT flags) noexcept
{
    if (! hdc || text.empty())
    {
        return;
    }

    DrawTextW(hdc, text.c_str(), static_cast<int>(text.size()), &rect, flags | DT_END_ELLIPSIS | DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX);
}
} // namespace

FunctionBar::FunctionBar() = default;

FunctionBar::~FunctionBar()
{
    Destroy();
}

ATOM FunctionBar::RegisterWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

HWND FunctionBar::Create(HWND parent, int x, int y, int width, int height)
{
    _hInstance = GetModuleHandleW(nullptr);

    if (! RegisterWndClass(_hInstance))
    {
        return nullptr;
    }

    CreateWindowExW(0, kClassName, L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, x, y, width, height, parent, nullptr, _hInstance, this);

    return _hWnd.get();
}

void FunctionBar::Destroy() noexcept
{
    _backgroundBrush.reset();
    _pressedBrush.reset();
    _hoverBrush.reset();
    _glyphPen.reset();
    _separatorPen.reset();
    _keyFont.reset();
    _textFont.reset();
    _hWnd.reset();
}

void FunctionBar::SetTheme(const AppTheme& theme)
{
    _theme = theme;

    _backgroundBrush.reset(CreateSolidBrush(_theme.menu.background));
    _pressedBrush.reset(CreateSolidBrush(_theme.menu.selectionBg));
    const COLORREF hoverColor = _theme.highContrast ? _theme.menu.selectionBg : BlendColor(_theme.menu.background, _theme.menu.selectionBg, 1, 3);
    _hoverBrush.reset(CreateSolidBrush(hoverColor));
    _glyphPen.reset(CreatePen(PS_SOLID, 1, _theme.menu.separator));
    _separatorPen.reset(CreatePen(PS_SOLID, 1, _theme.menu.separator));

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void FunctionBar::SetShortcutManager(const ShortcutManager* shortcuts) noexcept
{
    _shortcuts = shortcuts;
    RecomputeLabels();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void FunctionBar::SetDpi(UINT dpi) noexcept
{
    _dpi = dpi;
    EnsureKeyFont();
    EnsureTextFont();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void FunctionBar::SetModifiers(uint32_t modifiers) noexcept
{
    modifiers = modifiers & 0x7u;
    if (_modifiers == modifiers)
    {
        return;
    }

    _modifiers = modifiers;
    RecomputeLabels();
    RecomputeModifierText();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void FunctionBar::SetPressedFunctionKey(std::optional<uint32_t> vk) noexcept
{
    if (_pressedKey == vk)
    {
        return;
    }

    _pressedKey = vk;
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

LRESULT CALLBACK FunctionBar::WndProcThunk(HWND hWindow, UINT msg, WPARAM wp, LPARAM lp)
{
    FunctionBar* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<FunctionBar*>(cs->lpCreateParams);
        SetWindowLongPtrW(hWindow, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->_hWnd.reset(hWindow);
    }
    else
    {
        self = reinterpret_cast<FunctionBar*>(GetWindowLongPtrW(hWindow, GWLP_USERDATA));
    }

    if (self)
    {
        return self->WndProc(hWindow, msg, wp, lp);
    }

    return DefWindowProcW(hWindow, msg, wp, lp);
}

LRESULT FunctionBar::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: OnPaint(); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_MOUSEMOVE: OnMouseMove({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_LBUTTONUP: OnLButtonUp({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void FunctionBar::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);
    EnsureKeyFont();
    EnsureTextFont();
    RecomputeLabels();
    RecomputeModifierText();
}

void FunctionBar::OnDestroy() noexcept
{
}

void FunctionBar::OnSize(UINT width, UINT height) noexcept
{
    _clientSize = {static_cast<LONG>(width), static_cast<LONG>(height)};
}

void FunctionBar::OnMouseMove(POINT pt) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (! _trackingMouseLeave)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = _hWnd.get();
        if (TrackMouseEvent(&tme))
        {
            _trackingMouseLeave = true;
        }
    }

    const std::optional<uint32_t> key = HitTestFunctionKey(pt);
    if (_hoveredKey == key)
    {
        return;
    }

    _hoveredKey = key;
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void FunctionBar::OnMouseLeave() noexcept
{
    _trackingMouseLeave = false;
    if (! _hoveredKey.has_value() || ! _hWnd)
    {
        _hoveredKey.reset();
        return;
    }

    _hoveredKey.reset();
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void FunctionBar::OnLButtonUp(POINT pt) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const std::optional<uint32_t> vkOpt = HitTestFunctionKey(pt);
    if (! vkOpt.has_value())
    {
        return;
    }

    uint32_t modifiers = 0;
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModCtrl;
    }
    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModAlt;
    }
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModShift;
    }

    const HWND owner = GetAncestor(_hWnd.get(), GA_ROOT);
    if (! owner)
    {
        return;
    }

    SendMessageW(owner, WndMsg::kFunctionBarInvoke, static_cast<WPARAM>(vkOpt.value()), static_cast<LPARAM>(modifiers & 0x7u));
}

std::optional<uint32_t> FunctionBar::HitTestFunctionKey(POINT pt) const noexcept
{
    if (! _hWnd)
    {
        return std::nullopt;
    }

    const int width  = static_cast<int>(_clientSize.cx);
    const int height = static_cast<int>(_clientSize.cy);
    if (width <= 0 || height <= 0)
    {
        return std::nullopt;
    }

    if (pt.x < 0 || pt.y < 0 || pt.x >= width || pt.y >= height)
    {
        return std::nullopt;
    }

    wil::unique_hdc_window hdc(GetDC(_hWnd.get()));
    if (! hdc)
    {
        return std::nullopt;
    }

    const HFONT fallbackFont      = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    const HFONT textFont          = _textFont ? _textFont.get() : fallbackFont;
    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), textFont);

    const int paddingX          = PxFromDip(kZonePaddingXDip);
    const int modifiersGap      = std::max(1, PxFromDip(kModifiersGapDip) / 2);
    const int modifiersPaddingX = std::max(1, paddingX / 2);

    int reservedModifiersWidth = 0;
    {
        const std::wstring maxModifiersText = BuildModifierText(ShortcutManager::kModCtrl | ShortcutManager::kModAlt | ShortcutManager::kModShift);
        if (! maxModifiersText.empty())
        {
            SIZE modSize{};
            GetTextExtentPoint32W(hdc.get(), maxModifiersText.c_str(), static_cast<int>(maxModifiersText.size()), &modSize);
            reservedModifiersWidth = modSize.cx + (modifiersPaddingX * 2) + modifiersGap;

            const int minZoneWidthPx = PxFromDip(kMinZoneWidthForModifiers);
            if ((width - reservedModifiersWidth) / kFunctionKeyCount < minZoneWidthPx)
            {
                reservedModifiersWidth = 0;
            }
        }
    }

    const int zonesWidth = std::max(0, width - reservedModifiersWidth);
    if (zonesWidth <= 0 || pt.x >= zonesWidth)
    {
        return std::nullopt;
    }

    const int index = std::clamp((static_cast<int>(pt.x) * kFunctionKeyCount) / zonesWidth, 0, kFunctionKeyCount - 1);
    return static_cast<uint32_t>(VK_F1 + index);
}

void FunctionBar::OnPaint() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    PAINTSTRUCT ps;
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd.get(), &ps);

    RECT client{};
    if (! GetClientRect(_hWnd.get(), &client))
    {
        return;
    }

    const HBRUSH bg = _backgroundBrush ? _backgroundBrush.get() : static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH));
    FillRect(hdc.get(), &client, bg);

    const int width  = std::max(0l, client.right - client.left);
    const int height = std::max(0l, client.bottom - client.top);
    if (width <= 0 || height <= 0)
    {
        return;
    }

    const HFONT fallbackFont      = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    const HFONT textFont          = _textFont ? _textFont.get() : fallbackFont;
    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), textFont);
    SetBkMode(hdc.get(), TRANSPARENT);

    const int paddingX          = PxFromDip(kZonePaddingXDip);
    const int paddingY          = PxFromDip(kZonePaddingYDip);
    const int glyphPadX         = PxFromDip(kGlyphPaddingXDip);
    const int glyphPadY         = PxFromDip(kGlyphPaddingYDip);
    const int modifiersGap      = std::max(1, PxFromDip(kModifiersGapDip) / 2);
    const int modifiersPaddingX = std::max(1, paddingX / 2);

    int reservedModifiersWidth = 0;
    {
        const std::wstring maxModifiersText = BuildModifierText(ShortcutManager::kModCtrl | ShortcutManager::kModAlt | ShortcutManager::kModShift);
        if (! maxModifiersText.empty())
        {
            SIZE modSize{};
            GetTextExtentPoint32W(hdc.get(), maxModifiersText.c_str(), static_cast<int>(maxModifiersText.size()), &modSize);
            reservedModifiersWidth = modSize.cx + (modifiersPaddingX * 2) + modifiersGap;

            const int minZoneWidthPx = PxFromDip(kMinZoneWidthForModifiers);
            if ((width - reservedModifiersWidth) / kFunctionKeyCount < minZoneWidthPx)
            {
                reservedModifiersWidth = 0;
            }
        }
    }

    const bool showModifiers = reservedModifiersWidth > 0 && ! _modifierText.empty();
    const int zonesWidth     = std::max(0, width - reservedModifiersWidth);

    const COLORREF textColorNormal  = _theme.menu.text;
    const COLORREF textColorPressed = _theme.menu.selectionText;

    HFONT keyFont = _keyFont ? _keyFont.get() : textFont;

    for (int i = 0; i < kFunctionKeyCount; ++i)
    {
        const int left  = (i * zonesWidth) / kFunctionKeyCount;
        const int right = ((i + 1) * zonesWidth) / kFunctionKeyCount;

        RECT zone{};
        zone.left   = left;
        zone.top    = 0;
        zone.right  = right;
        zone.bottom = height;

        const uint32_t vk  = static_cast<uint32_t>(VK_F1 + i);
        const bool pressed = _pressedKey.has_value() && _pressedKey.value() == vk;
        const bool hovered = _hoveredKey.has_value() && _hoveredKey.value() == vk;
        if (pressed && _pressedBrush)
        {
            FillRect(hdc.get(), &zone, _pressedBrush.get());
        }
        else if (hovered && _hoverBrush)
        {
            FillRect(hdc.get(), &zone, _hoverBrush.get());
        }

        if (_separatorPen && i > 0)
        {
            [[maybe_unused]] auto oldPen = wil::SelectObject(hdc.get(), _separatorPen.get());
            MoveToEx(hdc.get(), zone.left, zone.top, nullptr);
            LineTo(hdc.get(), zone.left, zone.bottom);
        }

        wchar_t keyLabel[4]{};
        keyLabel[0]           = L'F';
        const unsigned number = static_cast<unsigned>(i + 1);
        if (number < 10)
        {
            keyLabel[1] = static_cast<wchar_t>(L'0' + number);
            keyLabel[2] = L'\0';
        }
        else
        {
            keyLabel[1] = static_cast<wchar_t>(L'0' + (number / 10));
            keyLabel[2] = static_cast<wchar_t>(L'0' + (number % 10));
            keyLabel[3] = L'\0';
        }
        SIZE keySize{};
        const int keyLabelLength = static_cast<int>(std::wcslen(keyLabel));
        {
            [[maybe_unused]] auto oldKeyFont = wil::SelectObject(hdc.get(), keyFont);
            GetTextExtentPoint32W(hdc.get(), keyLabel, keyLabelLength, &keySize);
        }

        const int availableHeight    = std::max(1, height - (paddingY * 2));
        const int desiredGlyphHeight = keySize.cy + (glyphPadY * 2);
        const int glyphHeight        = std::clamp(desiredGlyphHeight, 1, availableHeight);
        const int glyphWidth         = std::min(zone.right - zone.left - (paddingX * 2), keySize.cx + (glyphPadX * 2));
        const int glyphTop           = zone.top + (height - glyphHeight) / 2;

        RECT glyph{};
        glyph.left   = zone.left + paddingX;
        glyph.top    = glyphTop;
        glyph.right  = glyph.left + std::max(1, glyphWidth);
        glyph.bottom = glyph.top + glyphHeight;

        if (_glyphPen)
        {
            [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc.get(), _glyphPen.get());
            [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc.get(), GetStockObject(HOLLOW_BRUSH));
            const int radius               = std::clamp(PxFromDip(kGlyphCornerRadiusDip), 1, std::max(1, glyphHeight / 2));
            RoundRect(hdc.get(), glyph.left, glyph.top, glyph.right, glyph.bottom, radius, radius);
        }

        SetTextColor(hdc.get(), pressed ? textColorPressed : textColorNormal);
        RECT keyTextRect  = glyph;
        keyTextRect.left  = std::min(keyTextRect.right, keyTextRect.left + glyphPadX);
        keyTextRect.right = std::max(keyTextRect.left, keyTextRect.right - glyphPadX);
        {
            [[maybe_unused]] auto oldKeyFont = wil::SelectObject(hdc.get(), keyFont);
            DrawTextW(hdc.get(), keyLabel, keyLabelLength, &keyTextRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        }

        RECT textRect{};
        textRect.left   = std::min(zone.right, glyph.right + paddingX);
        textRect.top    = zone.top;
        textRect.right  = std::max(textRect.left, zone.right - paddingX);
        textRect.bottom = zone.bottom;

        if (i >= 0 && i < static_cast<int>(_labels.size()))
        {
            DrawTextEllipsis(hdc.get(), _labels[static_cast<size_t>(i)], textRect, DT_LEFT);
        }
    }

    if (showModifiers)
    {
        RECT modRect{};
        modRect.left   = zonesWidth;
        modRect.top    = 0;
        modRect.right  = width;
        modRect.bottom = height;

        SetTextColor(hdc.get(), textColorNormal);
        RECT modTextRect  = modRect;
        modTextRect.left  = std::min(modTextRect.right, modTextRect.left + modifiersPaddingX);
        modTextRect.right = std::max(modTextRect.left, modTextRect.right - modifiersPaddingX);
        DrawTextEllipsis(hdc.get(), _modifierText, modTextRect, DT_RIGHT);
    }
}

void FunctionBar::RecomputeLabels()
{
    for (auto& label : _labels)
    {
        label.clear();
    }

    if (! _shortcuts)
    {
        return;
    }

    for (int i = 0; i < kFunctionKeyCount; ++i)
    {
        const uint32_t vk                                 = static_cast<uint32_t>(VK_F1 + i);
        const std::optional<std::wstring_view> commandOpt = _shortcuts->FindFunctionBarCommand(vk, _modifiers);
        if (! commandOpt.has_value())
        {
            continue;
        }

        const std::optional<unsigned int> displayNameIdOpt = TryGetCommandDisplayNameStringId(commandOpt.value());
        if (displayNameIdOpt.has_value())
        {
            _labels[static_cast<size_t>(i)] = LoadStringResource(nullptr, displayNameIdOpt.value());
        }
    }
}

void FunctionBar::RecomputeModifierText()
{
    _modifierText = BuildModifierText(_modifiers);
}

void FunctionBar::EnsureKeyFont() noexcept
{
    _keyFont.reset();

    LOGFONTW lf{};
    const HFONT baseFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    if (! baseFont)
    {
        return;
    }

    if (GetObjectW(baseFont, sizeof(lf), &lf) != sizeof(lf))
    {
        return;
    }

    lf.lfHeight = -MulDiv(kKeyFontHeightDip, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI);
    lf.lfWeight = FW_NORMAL;
    _keyFont.reset(CreateFontIndirectW(&lf));
}

void FunctionBar::EnsureTextFont() noexcept
{
    _textFont.reset();

    LOGFONTW lf{};
    const HFONT baseFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    if (! baseFont)
    {
        return;
    }

    if (GetObjectW(baseFont, sizeof(lf), &lf) != sizeof(lf))
    {
        return;
    }

    lf.lfHeight = -MulDiv(kTextFontHeightDip, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI);
    lf.lfWeight = FW_NORMAL;
    _textFont.reset(CreateFontIndirectW(&lf));
}

int FunctionBar::PxFromDip(int dip) const noexcept
{
    return MulDiv(dip, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI);
}
