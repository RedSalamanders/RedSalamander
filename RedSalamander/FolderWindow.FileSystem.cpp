#include "ChangeCase.h"
#include "ConnectionManagerDialog.h"
#include "ConnectionSecrets.h"
#include "FolderWindowInternal.h"
#include "Helpers.h"
#include "HostServices.h"
#include "MaskSyntax.h"
#include "NavigationLocation.h"

#include "SettingsStore.h"
#include "ThemedControls.h"
#include "ThemedInputFrames.h"

#include <cstring>
#include <limits>
#include <map>
#include <unordered_set>

#include <commctrl.h>
#include <shellapi.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

namespace
{
using OrdinalString::EqualsNoCase;
using OrdinalString::StartsWithNoCase;

bool IsFilePluginShortId(std::wstring_view pluginShortId) noexcept
{
    return EqualsNoCase(pluginShortId, L"file");
}

[[nodiscard]] bool SetClipboardUnicodeText(HWND ownerWindow, std::wstring_view text) noexcept
{
    if (! ownerWindow)
    {
        return false;
    }

    if (text.size() >= (std::numeric_limits<size_t>::max() / sizeof(wchar_t)))
    {
        return false;
    }

    const SIZE_T bytes = (text.size() + 1u) * sizeof(wchar_t);
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! memory)
    {
        return false;
    }

    auto* out = static_cast<wchar_t*>(GlobalLock(memory.get()));
    if (! out)
    {
        return false;
    }

    if (! text.empty())
    {
        memcpy(out, text.data(), text.size() * sizeof(wchar_t));
    }
    out[text.size()] = L'\0';
    GlobalUnlock(memory.get());

    if (OpenClipboard(ownerWindow) == 0)
    {
        return false;
    }
    const auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });

    if (EmptyClipboard() == 0)
    {
        return false;
    }

    if (SetClipboardData(CF_UNICODETEXT, memory.get()) == nullptr)
    {
        return false;
    }

    static_cast<void>(memory.release());
    return true;
}

struct ChangeCaseDialogState
{
    ChangeCaseDialogState()                                        = default;
    ChangeCaseDialogState(const ChangeCaseDialogState&)            = delete;
    ChangeCaseDialogState& operator=(const ChangeCaseDialogState&) = delete;
    ChangeCaseDialogState(ChangeCaseDialogState&&)                 = delete;
    ChangeCaseDialogState& operator=(ChangeCaseDialogState&&)      = delete;
    ~ChangeCaseDialogState()                                       = default;

    struct ToggleCard final
    {
        HWND title       = nullptr;
        HWND description = nullptr;
        HWND toggle      = nullptr;
    };

    struct Ui final
    {
        ToggleCard includeSubdirs{};

        HWND headerChangeCaseTo = nullptr;
        ToggleCard lower{};
        ToggleCard upper{};
        ToggleCard partiallyMixed{};
        ToggleCard mixed{};

        HWND headerChange = nullptr;
        ToggleCard whole{};
        ToggleCard onlyName{};
        ToggleCard onlyExtension{};
    };

    AppTheme theme{};
    wil::unique_hbrush backgroundBrush;
    wil::unique_hbrush cardBrush;
    std::vector<RECT> cards;
    bool useTwoColumns            = false;
    int twoColumnSeparatorX       = -1;
    int twoColumnSeparatorYTop    = 0;
    int twoColumnSeparatorYBottom = 0;
    wil::unique_hfont boldFont;
    wil::unique_hfont italicFont;
    UINT fontsDpi = 0;
    std::wstring toggleOnLabel;
    std::wstring toggleOffLabel;

    bool allowSubdirs = false;
    ChangeCase::Options options{};
    bool accepted = false;

    Ui ui{};
};

void CenterWindowOnOwner(HWND window, HWND owner) noexcept;

[[nodiscard]] int MeasureStaticTextHeight(HWND referenceWindow, HFONT font, int width, std::wstring_view text) noexcept
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

void EnsureChangeCaseDialogFonts(HWND dlg, ChangeCaseDialogState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(dlg);
    if (state.fontsDpi != dpi)
    {
        state.fontsDpi = dpi;
        state.boldFont.reset();
        state.italicFont.reset();
    }

    HFONT baseFont = reinterpret_cast<HFONT>(SendMessageW(dlg, WM_GETFONT, 0, 0));
    if (! baseFont)
    {
        baseFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    LOGFONTW lf{};
    if (GetObjectW(baseFont, sizeof(lf), &lf) != sizeof(lf))
    {
        return;
    }

    if (! state.boldFont)
    {
        LOGFONTW bold = lf;
        bold.lfWeight = std::max<LONG>(FW_SEMIBOLD, lf.lfWeight);
        state.boldFont.reset(CreateFontIndirectW(&bold));
    }
    if (! state.italicFont)
    {
        LOGFONTW italic = lf;
        italic.lfItalic = TRUE;
        state.italicFont.reset(CreateFontIndirectW(&italic));
    }
}

LRESULT CALLBACK ChangeCaseToggleCheckSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR) noexcept
{
    switch (msg)
    {
        case BM_GETCHECK: return (GetWindowLongPtrW(hwnd, GWLP_USERDATA) != 0) ? BST_CHECKED : BST_UNCHECKED;
        case BM_SETCHECK:
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, (wp == BST_CHECKED) ? 1 : 0);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        case WM_NCDESTROY: RemoveWindowSubclass(hwnd, ChangeCaseToggleCheckSubclassProc, subclassId); break;
        default: break;
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

void ApplyChangeCaseDialogTheme(HWND dlg, const AppTheme& theme) noexcept
{
    if (! dlg)
    {
        return;
    }

    const bool darkBackground = ChooseContrastingTextColor(theme.windowBackground) == RGB(255, 255, 255);
    const wchar_t* themeName  = theme.highContrast ? L"" : (darkBackground ? L"DarkMode_Explorer" : L"Explorer");

    SetWindowTheme(dlg, themeName, nullptr);
    SendMessageW(dlg, WM_THEMECHANGED, 0, 0);

    EnumChildWindows(dlg,
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        const wchar_t* themeName = reinterpret_cast<const wchar_t*>(lParam);
        if (! child)
        {
            return TRUE;
        }

        const wchar_t* appliedTheme = themeName ? themeName : L"";
        if (themeName)
        {
            wchar_t className[32]{};
            const int classLen = GetClassNameW(child, className, static_cast<int>(_countof(className)));
            if (classLen > 0)
            {
                if (_wcsicmp(className, L"Static") == 0)
                {
                    appliedTheme = L"";
                }
                else if (_wcsicmp(className, L"Button") == 0)
                {
                    const LONG_PTR style = GetWindowLongPtrW(child, GWL_STYLE);
                    const LONG_PTR type  = style & BS_TYPEMASK;
                    if (type == BS_GROUPBOX || type == BS_PUSHBUTTON || type == BS_DEFPUSHBUTTON)
                    {
                        appliedTheme = L"";
                    }
                }
            }
        }

        SetWindowTheme(child, appliedTheme, nullptr);
        SendMessageW(child, WM_THEMECHANGED, 0, 0);
        return TRUE;
    },
                     reinterpret_cast<LPARAM>(themeName));

    RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
}

void EnsureChangeCaseDialogControlsCreated(HWND dlg, ChangeCaseDialogState& state) noexcept
{
    if (! dlg || state.ui.headerChangeCaseTo)
    {
        return;
    }

    const HINSTANCE instance = GetModuleHandleW(nullptr);
    const bool highContrast  = state.theme.highContrast || state.theme.systemHighContrast;

    constexpr DWORD baseStaticStyle = static_cast<DWORD>(WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX);
    constexpr DWORD wrapStaticStyle = static_cast<DWORD>(WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL);

    const DWORD toggleStyle = static_cast<DWORD>(WS_CHILD | WS_VISIBLE | WS_TABSTOP) | static_cast<DWORD>(highContrast ? BS_AUTOCHECKBOX : 0);

    const auto makeStatic = [&](DWORD style) noexcept -> HWND
    { return CreateWindowExW(0, L"Static", L"", style, 0, 0, 10, 10, dlg, nullptr, instance, nullptr); };

    const auto makeToggle = [&](int id) noexcept -> HWND
    {
        const HWND toggle =
            CreateWindowExW(0, L"Button", L"", toggleStyle, 0, 0, 10, 10, dlg, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), instance, nullptr);
        if (toggle && ! highContrast)
        {
            ThemedControls::EnableOwnerDrawButton(dlg, id);
            SetWindowSubclass(toggle, ChangeCaseToggleCheckSubclassProc, 1u, 0);
        }
        return toggle;
    };

    state.ui.includeSubdirs.title       = makeStatic(baseStaticStyle);
    state.ui.includeSubdirs.description = makeStatic(wrapStaticStyle);
    state.ui.includeSubdirs.toggle      = makeToggle(IDC_CHANGE_CASE_INCLUDE_SUBDIRS);

    state.ui.headerChangeCaseTo         = makeStatic(baseStaticStyle);
    state.ui.lower.title                = makeStatic(baseStaticStyle);
    state.ui.lower.description          = makeStatic(wrapStaticStyle);
    state.ui.lower.toggle               = makeToggle(IDC_CHANGE_CASE_LOWER);
    state.ui.upper.title                = makeStatic(baseStaticStyle);
    state.ui.upper.description          = makeStatic(wrapStaticStyle);
    state.ui.upper.toggle               = makeToggle(IDC_CHANGE_CASE_UPPER);
    state.ui.partiallyMixed.title       = makeStatic(baseStaticStyle);
    state.ui.partiallyMixed.description = makeStatic(wrapStaticStyle);
    state.ui.partiallyMixed.toggle      = makeToggle(IDC_CHANGE_CASE_PARTIALLY_MIXED);
    state.ui.mixed.title                = makeStatic(baseStaticStyle);
    state.ui.mixed.description          = makeStatic(wrapStaticStyle);
    state.ui.mixed.toggle               = makeToggle(IDC_CHANGE_CASE_MIXED);

    state.ui.headerChange              = makeStatic(baseStaticStyle);
    state.ui.whole.title               = makeStatic(baseStaticStyle);
    state.ui.whole.description         = makeStatic(wrapStaticStyle);
    state.ui.whole.toggle              = makeToggle(IDC_CHANGE_CASE_WHOLE);
    state.ui.onlyName.title            = makeStatic(baseStaticStyle);
    state.ui.onlyName.description      = makeStatic(wrapStaticStyle);
    state.ui.onlyName.toggle           = makeToggle(IDC_CHANGE_CASE_ONLY_NAME);
    state.ui.onlyExtension.title       = makeStatic(baseStaticStyle);
    state.ui.onlyExtension.description = makeStatic(wrapStaticStyle);
    state.ui.onlyExtension.toggle      = makeToggle(IDC_CHANGE_CASE_ONLY_EXTENSION);

    state.toggleOnLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    state.toggleOffLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
}

void PaintChangeCaseDialogBackgroundAndCards(HDC hdc, HWND dlg, const ChangeCaseDialogState& state) noexcept
{
    if (! hdc || ! dlg)
    {
        return;
    }

    RECT rc{};
    GetClientRect(dlg, &rc);

    if (state.backgroundBrush)
    {
        FillRect(hdc, &rc, state.backgroundBrush.get());
    }

    if (state.theme.systemHighContrast || state.theme.highContrast || state.cards.empty())
    {
        return;
    }

    const UINT dpi         = GetDpiForWindow(dlg);
    const int radius       = ThemedControls::ScaleDip(dpi, 6);
    const COLORREF surface = ThemedControls::GetControlSurfaceColor(state.theme);
    const COLORREF border  = ThemedControls::BlendColor(surface, state.theme.menu.text, state.theme.dark ? 40 : 30, 255);

    wil::unique_hbrush cardBrushFallback;
    HBRUSH cardBrush = state.cardBrush.get();
    if (! cardBrush)
    {
        cardBrushFallback.reset(CreateSolidBrush(surface));
        cardBrush = cardBrushFallback.get();
    }
    wil::unique_hpen cardPen(CreatePen(PS_SOLID, 1, border));
    if (! cardBrush || ! cardPen)
    {
        return;
    }

    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc, cardBrush);
    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc, cardPen.get());

    for (const RECT& card : state.cards)
    {
        RoundRect(hdc, card.left, card.top, card.right, card.bottom, radius, radius);
    }

    if (state.useTwoColumns && state.twoColumnSeparatorX > rc.left && state.twoColumnSeparatorX < rc.right)
    {
        const int rcTop    = static_cast<int>(rc.top);
        const int rcBottom = static_cast<int>(rc.bottom);
        const int yTop     = std::clamp(state.twoColumnSeparatorYTop, rcTop, rcBottom);
        const int yBottom  = std::clamp(state.twoColumnSeparatorYBottom, rcTop, rcBottom);
        if (yBottom > yTop)
        {
            const COLORREF separator = ThemedControls::BlendColor(surface, state.theme.menu.text, state.theme.dark ? 28 : 20, 255);
            wil::unique_hpen sepPen(CreatePen(PS_SOLID, 1, separator));
            if (sepPen)
            {
                [[maybe_unused]] auto oldSepPen = wil::SelectObject(hdc, sepPen.get());
                MoveToEx(hdc, state.twoColumnSeparatorX, yTop, nullptr);
                LineTo(hdc, state.twoColumnSeparatorX, yBottom);
            }
        }
    }
}

void LayoutChangeCaseDialogControls(HWND dlg, ChangeCaseDialogState& state) noexcept
{
    EnsureChangeCaseDialogControlsCreated(dlg, state);

    if (! dlg || ! state.ui.headerChangeCaseTo)
    {
        return;
    }

    RECT rcDlg{};
    if (! GetClientRect(dlg, &rcDlg))
    {
        return;
    }

    const int dlgW = std::max(0l, rcDlg.right - rcDlg.left);
    const int dlgH = std::max(0l, rcDlg.bottom - rcDlg.top);

    const UINT dpi = GetDpiForWindow(dlg);

    const int margin       = ThemedControls::ScaleDip(dpi, 16);
    const int gapX         = ThemedControls::ScaleDip(dpi, 12);
    const int gapY         = ThemedControls::ScaleDip(dpi, 12);
    const int rowHeight    = std::max(1, ThemedControls::ScaleDip(dpi, 26));
    const int titleHeight  = std::max(1, ThemedControls::ScaleDip(dpi, 18));
    const int headerHeight = std::max(1, ThemedControls::ScaleDip(dpi, 20));

    const int cardPaddingX   = ThemedControls::ScaleDip(dpi, 12);
    const int cardPaddingY   = ThemedControls::ScaleDip(dpi, 8);
    const int cardGapY       = ThemedControls::ScaleDip(dpi, 2);
    const int cardGapX       = ThemedControls::ScaleDip(dpi, 12);
    const int cardSpacingY   = ThemedControls::ScaleDip(dpi, 8);
    const int sectionSpacing = ThemedControls::ScaleDip(dpi, 16);

    HFONT dialogFont = reinterpret_cast<HFONT>(SendMessageW(dlg, WM_GETFONT, 0, 0));
    if (! dialogFont)
    {
        dialogFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    EnsureChangeCaseDialogFonts(dlg, state);
    const HFONT headerFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const HFONT infoFont   = state.italicFont ? state.italicFont.get() : dialogFont;

    const auto getWindowText = [](HWND hwnd) noexcept -> std::wstring
    {
        if (! hwnd)
        {
            return {};
        }
        const int len = GetWindowTextLengthW(hwnd);
        if (len <= 0)
        {
            return {};
        }
        std::wstring text(static_cast<size_t>(len) + 1u, L'\0');
        const int copied = GetWindowTextW(hwnd, text.data(), len + 1);
        if (copied <= 0)
        {
            return {};
        }
        text.resize(static_cast<size_t>(copied));
        return text;
    };

    const int buttonPadX = ThemedControls::ScaleDip(dpi, 16);
    const int minBtnW    = ThemedControls::ScaleDip(dpi, 80);

    const auto measureButtonWidth = [&](HWND btn) noexcept -> int
    {
        const std::wstring text = getWindowText(btn);
        const int textW         = ThemedControls::MeasureTextWidth(dlg, dialogFont, text);
        return std::max(minBtnW, (2 * buttonPadX) + textW);
    };

    const HWND okBtn     = GetDlgItem(dlg, IDOK);
    const HWND cancelBtn = GetDlgItem(dlg, IDCANCEL);

    const int okW      = measureButtonWidth(okBtn);
    const int cancelW  = measureButtonWidth(cancelBtn);
    const int buttonsY = std::max(0, dlgH - margin - rowHeight);

    const UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;

    int nextRight = std::max(0, dlgW - margin);
    if (cancelBtn)
    {
        nextRight -= cancelW;
        SetWindowPos(cancelBtn, nullptr, nextRight, buttonsY, cancelW, rowHeight, flags);
        SendMessageW(cancelBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        nextRight -= gapX;
    }
    if (okBtn)
    {
        nextRight -= okW;
        SetWindowPos(okBtn, nullptr, nextRight, buttonsY, okW, rowHeight, flags);
        SendMessageW(okBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    const int contentX = margin;
    const int contentY = margin;
    const int contentW = std::max(0, dlgW - 2 * margin);

    const int minToggleWidth  = ThemedControls::ScaleDip(dpi, 90);
    const int onWidth         = ThemedControls::MeasureTextWidth(dlg, headerFont, state.toggleOnLabel);
    const int offWidth        = ThemedControls::MeasureTextWidth(dlg, headerFont, state.toggleOffLabel);
    const int togglePaddingX  = ThemedControls::ScaleDip(dpi, 6);
    const int toggleGapX      = ThemedControls::ScaleDip(dpi, 8);
    const int toggleTrackW    = ThemedControls::ScaleDip(dpi, 34);
    const int stateTextW      = std::max(onWidth, offWidth);
    const int measuredToggleW = std::max(minToggleWidth, (2 * togglePaddingX) + stateTextW + toggleGapX + toggleTrackW);

    const auto computeToggleWidth = [&](int availableW) noexcept -> int { return std::min(std::max(0, availableW - 2 * cardPaddingX), measuredToggleW); };

    const int columnSeparatorAreaW = ThemedControls::ScaleDip(dpi, 28);
    const int minLeftColumnW       = ThemedControls::ScaleDip(dpi, 320);
    const int minRightColumnW      = ThemedControls::ScaleDip(dpi, 240);

    state.cards.clear();
    state.useTwoColumns             = false;
    state.twoColumnSeparatorX       = -1;
    state.twoColumnSeparatorYTop    = 0;
    state.twoColumnSeparatorYBottom = 0;

    auto pushCard = [&](const RECT& card) noexcept { state.cards.push_back(card); };

    auto showToggleCardControls = [&](const ChangeCaseDialogState::ToggleCard& card, bool visible) noexcept
    {
        const int cmd = visible ? SW_SHOW : SW_HIDE;
        ShowWindow(card.title, cmd);
        ShowWindow(card.description, cmd);
        ShowWindow(card.toggle, cmd);
    };

    auto layoutSectionHeaderAt = [&](HWND header, std::wstring_view text, int x, int width, int& y) noexcept
    {
        if (! header)
        {
            return;
        }

        SetWindowTextW(header, text.data());
        ShowWindow(header, SW_SHOW);
        SetWindowPos(header, nullptr, x, y, width, headerHeight, flags);
        SendMessageW(header, WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    };

    auto layoutToggleCardAt = [&](const ChangeCaseDialogState::ToggleCard& card,
                                  std::wstring_view titleText,
                                  std::wstring_view descText,
                                  bool visible,
                                  int x,
                                  int width,
                                  int toggleW,
                                  bool addBottomSpacing,
                                  int& y) noexcept
    {
        showToggleCardControls(card, visible);
        if (! visible)
        {
            return;
        }

        const int textW = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH = descText.empty() ? 0 : MeasureStaticTextHeight(dlg, infoFont, textW, descText);

        const int contentH = std::max(0, titleHeight + (descH > 0 ? (cardGapY + descH) : 0));
        const int cardH    = std::max(rowHeight + 2 * cardPaddingY, contentH + 2 * cardPaddingY);

        RECT cardRc{};
        cardRc.left   = x;
        cardRc.top    = y;
        cardRc.right  = x + width;
        cardRc.bottom = y + cardH;
        pushCard(cardRc);

        SetWindowTextW(card.title, titleText.data());
        SetWindowPos(card.title, nullptr, cardRc.left + cardPaddingX, cardRc.top + cardPaddingY, textW, titleHeight, flags);
        SendMessageW(card.title, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        if (descH > 0)
        {
            SetWindowTextW(card.description, descText.data());
            ShowWindow(card.description, SW_SHOW);
            SetWindowPos(
                card.description, nullptr, cardRc.left + cardPaddingX, cardRc.top + cardPaddingY + titleHeight + cardGapY, textW, std::max(0, descH), flags);
            SendMessageW(card.description, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        }
        else
        {
            ShowWindow(card.description, SW_HIDE);
        }

        if (card.toggle)
        {
            SetWindowPos(card.toggle, nullptr, cardRc.right - cardPaddingX - toggleW, cardRc.top + (cardH - rowHeight) / 2, toggleW, rowHeight, flags);
            SendMessageW(card.toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        y += cardH + (addBottomSpacing ? cardSpacingY : 0);
    };

    int y = contentY;

    const int toggleWFull = computeToggleWidth(contentW);

    layoutToggleCardAt(
        state.ui.includeSubdirs, L"Include subdirectories", L"Apply to selected folders recursively.", true, contentX, contentW, toggleWFull, true, y);

    y += sectionSpacing;

    const bool useTwoColumnsLayout = contentW >= (minLeftColumnW + minRightColumnW + columnSeparatorAreaW);

    if (useTwoColumnsLayout)
    {
        const int availableW = std::max(0, contentW - columnSeparatorAreaW);
        const int leftW      = std::clamp(availableW / 2, minLeftColumnW, availableW - minRightColumnW);
        const int rightW     = std::max(0, availableW - leftW);
        const int leftX      = contentX;
        const int rightX     = contentX + leftW + columnSeparatorAreaW;

        const int toggleWLeft  = computeToggleWidth(leftW);
        const int toggleWRight = computeToggleWidth(rightW);

        int leftY  = y;
        int rightY = y;

        if (! state.theme.systemHighContrast && ! state.theme.highContrast)
        {
            state.useTwoColumns          = true;
            state.twoColumnSeparatorX    = contentX + leftW + (columnSeparatorAreaW / 2);
            state.twoColumnSeparatorYTop = y;
        }

        layoutSectionHeaderAt(state.ui.headerChangeCaseTo, L"Change case to", leftX, leftW, leftY);
        layoutToggleCardAt(state.ui.lower, L"Lower case", {}, true, leftX, leftW, toggleWLeft, true, leftY);
        layoutToggleCardAt(state.ui.upper, L"Upper case", {}, true, leftX, leftW, toggleWLeft, true, leftY);
        layoutToggleCardAt(
            state.ui.partiallyMixed, L"Partially mixed case", L"Name in Mixed Case, extension in lower case.", true, leftX, leftW, toggleWLeft, true, leftY);
        layoutToggleCardAt(state.ui.mixed, L"Mixed case", L"Title case (first letter of each word).", true, leftX, leftW, toggleWLeft, false, leftY);

        layoutSectionHeaderAt(state.ui.headerChange, L"Change", rightX, rightW, rightY);
        layoutToggleCardAt(state.ui.whole, L"Whole filename", {}, true, rightX, rightW, toggleWRight, true, rightY);
        layoutToggleCardAt(state.ui.onlyName, L"Only name", {}, true, rightX, rightW, toggleWRight, true, rightY);
        layoutToggleCardAt(state.ui.onlyExtension, L"Only extension", {}, true, rightX, rightW, toggleWRight, false, rightY);

        y = std::max(leftY, rightY);
        if (state.useTwoColumns)
        {
            state.twoColumnSeparatorYBottom = y;
        }
    }
    else
    {
        layoutSectionHeaderAt(state.ui.headerChangeCaseTo, L"Change case to", contentX, contentW, y);
        layoutToggleCardAt(state.ui.lower, L"Lower case", {}, true, contentX, contentW, toggleWFull, true, y);
        layoutToggleCardAt(state.ui.upper, L"Upper case", {}, true, contentX, contentW, toggleWFull, true, y);
        layoutToggleCardAt(
            state.ui.partiallyMixed, L"Partially mixed case", L"Name in Mixed Case, extension in lower case.", true, contentX, contentW, toggleWFull, true, y);
        layoutToggleCardAt(state.ui.mixed, L"Mixed case", L"Title case (first letter of each word).", true, contentX, contentW, toggleWFull, false, y);

        y += sectionSpacing;
        layoutSectionHeaderAt(state.ui.headerChange, L"Change", contentX, contentW, y);
        layoutToggleCardAt(state.ui.whole, L"Whole filename", {}, true, contentX, contentW, toggleWFull, true, y);
        layoutToggleCardAt(state.ui.onlyName, L"Only name", {}, true, contentX, contentW, toggleWFull, true, y);
        layoutToggleCardAt(state.ui.onlyExtension, L"Only extension", {}, true, contentX, contentW, toggleWFull, false, y);
    }

    if (state.ui.includeSubdirs.toggle)
    {
        EnableWindow(state.ui.includeSubdirs.toggle, state.allowSubdirs ? TRUE : FALSE);
        InvalidateRect(state.ui.includeSubdirs.toggle, nullptr, TRUE);
    }

    InvalidateRect(dlg, nullptr, TRUE);
}

void EnsureChangeCaseDialogAllOptionsVisible(HWND dlg, ChangeCaseDialogState& state, bool allowShrink) noexcept
{
    if (! dlg || state.cards.empty())
    {
        return;
    }

    const HWND okBtn     = GetDlgItem(dlg, IDOK);
    const HWND cancelBtn = GetDlgItem(dlg, IDCANCEL);
    if (! okBtn && ! cancelBtn)
    {
        return;
    }

    const auto mapWindowTopToClient = [&](HWND window) noexcept -> std::optional<int>
    {
        if (! window)
        {
            return std::nullopt;
        }

        RECT rc{};
        if (GetWindowRect(window, &rc) == 0)
        {
            return std::nullopt;
        }

        MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&rc), 2);
        return static_cast<int>(rc.top);
    };

    int buttonsTop = std::numeric_limits<int>::max();
    if (const auto okTop = mapWindowTopToClient(okBtn); okTop.has_value())
    {
        buttonsTop = std::min(buttonsTop, okTop.value());
    }
    if (const auto cancelTop = mapWindowTopToClient(cancelBtn); cancelTop.has_value())
    {
        buttonsTop = std::min(buttonsTop, cancelTop.value());
    }
    if (buttonsTop == std::numeric_limits<int>::max())
    {
        return;
    }

    int contentBottom = 0;
    for (const RECT& card : state.cards)
    {
        contentBottom = std::max(contentBottom, static_cast<int>(card.bottom));
    }

    const UINT dpi          = GetDpiForWindow(dlg);
    const int gapY          = ThemedControls::ScaleDip(dpi, 12);
    const int desiredBottom = contentBottom + gapY;
    const int delta         = desiredBottom - buttonsTop;
    if (delta == 0 || (! allowShrink && delta < 0))
    {
        return;
    }

    RECT windowRc{};
    RECT clientRc{};
    if (! GetWindowRect(dlg, &windowRc) || ! GetClientRect(dlg, &clientRc))
    {
        return;
    }

    const int windowW = std::max(1l, windowRc.right - windowRc.left);
    const int windowH = std::max(1l, windowRc.bottom - windowRc.top);

    int newWindowH = windowH + delta;

    if (delta > 0)
    {
        HMONITOR monitor = MonitorFromWindow(dlg, MONITOR_DEFAULTTONEAREST);
        MONITORINFO mi{};
        mi.cbSize = sizeof(mi);
        if (monitor && GetMonitorInfoW(monitor, &mi))
        {
            const int workH = static_cast<int>(mi.rcWork.bottom - mi.rcWork.top);
            newWindowH      = std::min(newWindowH, workH);
        }
    }

    newWindowH = std::max(1, newWindowH);

    if (newWindowH == windowH)
    {
        return;
    }

    SetWindowPos(dlg, nullptr, 0, 0, windowW, newWindowH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

    if (HWND owner = GetParent(dlg))
    {
        CenterWindowOnOwner(dlg, owner);
    }
}

INT_PTR OnChangeCaseDialogInit(HWND dlg, ChangeCaseDialogState* state) noexcept
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    ApplyTitleBarTheme(dlg, state->theme, GetActiveWindow() == dlg);
    state->backgroundBrush.reset(CreateSolidBrush(state->theme.windowBackground));
    state->cardBrush.reset();
    if (! state->theme.systemHighContrast && ! state->theme.highContrast)
    {
        const COLORREF surface = ThemedControls::GetControlSurfaceColor(state->theme);
        state->cardBrush.reset(CreateSolidBrush(surface));
    }
    EnsureChangeCaseDialogControlsCreated(dlg, *state);
    ApplyChangeCaseDialogTheme(dlg, state->theme);

    const bool highContrast = state->theme.highContrast || state->theme.systemHighContrast;
    if (! highContrast)
    {
        ThemedControls::EnableOwnerDrawButton(dlg, IDOK);
        ThemedControls::EnableOwnerDrawButton(dlg, IDCANCEL);
    }

    CheckRadioButton(dlg, IDC_CHANGE_CASE_LOWER, IDC_CHANGE_CASE_MIXED, IDC_CHANGE_CASE_LOWER);
    CheckRadioButton(dlg, IDC_CHANGE_CASE_WHOLE, IDC_CHANGE_CASE_ONLY_EXTENSION, IDC_CHANGE_CASE_WHOLE);

    if (HWND include = GetDlgItem(dlg, IDC_CHANGE_CASE_INCLUDE_SUBDIRS))
    {
        EnableWindow(include, state->allowSubdirs ? TRUE : FALSE);
        CheckDlgButton(dlg, IDC_CHANGE_CASE_INCLUDE_SUBDIRS, BST_UNCHECKED);
    }

    LayoutChangeCaseDialogControls(dlg, *state);
    EnsureChangeCaseDialogAllOptionsVisible(dlg, *state, true);

    return TRUE;
}

INT_PTR OnChangeCaseDialogDpiChanged(HWND dlg, ChangeCaseDialogState* state, UINT dpi, const RECT* suggested) noexcept
{
    UNREFERENCED_PARAMETER(dpi);

    if (! dlg || ! state)
    {
        return FALSE;
    }

    if (suggested)
    {
        const int width  = static_cast<int>(std::max(0l, suggested->right - suggested->left));
        const int height = static_cast<int>(std::max(0l, suggested->bottom - suggested->top));
        SetWindowPos(dlg, nullptr, suggested->left, suggested->top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    state->fontsDpi = 0;
    state->boldFont.reset();
    state->italicFont.reset();

    LayoutChangeCaseDialogControls(dlg, *state);
    EnsureChangeCaseDialogAllOptionsVisible(dlg, *state, true);

    RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
    return TRUE;
}

INT_PTR OnChangeCaseDialogCtlColorDialog(ChangeCaseDialogState* state) noexcept
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnChangeCaseDialogCtlColorStatic(ChangeCaseDialogState* state, HDC hdc, HWND control) noexcept
{
    if (! state || ! state->backgroundBrush || ! hdc)
    {
        return FALSE;
    }

    const bool enabled       = ! control || IsWindowEnabled(control) != FALSE;
    const COLORREF textColor = enabled ? state->theme.menu.text : state->theme.menu.disabledText;

    COLORREF background = state->theme.windowBackground;
    HBRUSH brush        = state->backgroundBrush.get();

    if (! state->theme.systemHighContrast && ! state->theme.highContrast && control && state->cardBrush && ! state->cards.empty())
    {
        RECT rcControl{};
        if (GetWindowRect(control, &rcControl) != 0)
        {
            const HWND root = GetAncestor(control, GA_ROOT);
            if (root)
            {
                MapWindowPoints(nullptr, root, reinterpret_cast<POINT*>(&rcControl), 2);
                POINT center{};
                center.x = (rcControl.left + rcControl.right) / 2;
                center.y = (rcControl.top + rcControl.bottom) / 2;

                for (const RECT& card : state->cards)
                {
                    if (PtInRect(&card, center) != FALSE)
                    {
                        background = ThemedControls::GetControlSurfaceColor(state->theme);
                        brush      = state->cardBrush.get();
                        break;
                    }
                }
            }
        }
    }

    if (! state->theme.systemHighContrast)
    {
        SetBkMode(hdc, TRANSPARENT);
        SetBkColor(hdc, background);
        SetTextColor(hdc, textColor);
        return reinterpret_cast<INT_PTR>(brush);
    }

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, background);
    SetTextColor(hdc, textColor);
    return reinterpret_cast<INT_PTR>(brush);
}

INT_PTR OnChangeCaseDialogCtlColorButton(ChangeCaseDialogState* state, HDC hdc, HWND control) noexcept
{
    if (! state || ! state->backgroundBrush || ! hdc)
    {
        return FALSE;
    }

    const bool enabled       = ! control || IsWindowEnabled(control) != FALSE;
    const COLORREF textColor = enabled ? state->theme.menu.text : state->theme.menu.disabledText;

    COLORREF background = state->theme.windowBackground;
    HBRUSH brush        = state->backgroundBrush.get();

    if (! state->theme.systemHighContrast && ! state->theme.highContrast && control && state->cardBrush && ! state->cards.empty())
    {
        RECT rcControl{};
        if (GetWindowRect(control, &rcControl) != 0)
        {
            const HWND root = GetAncestor(control, GA_ROOT);
            if (root)
            {
                MapWindowPoints(nullptr, root, reinterpret_cast<POINT*>(&rcControl), 2);
                POINT center{};
                center.x = (rcControl.left + rcControl.right) / 2;
                center.y = (rcControl.top + rcControl.bottom) / 2;

                for (const RECT& card : state->cards)
                {
                    if (PtInRect(&card, center) != FALSE)
                    {
                        background = ThemedControls::GetControlSurfaceColor(state->theme);
                        brush      = state->cardBrush.get();
                        break;
                    }
                }
            }
        }
    }

    if (! state->theme.systemHighContrast)
    {
        SetBkMode(hdc, TRANSPARENT);
        SetBkColor(hdc, background);
        SetTextColor(hdc, textColor);
        return reinterpret_cast<INT_PTR>(brush);
    }

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, background);
    SetTextColor(hdc, textColor);
    return reinterpret_cast<INT_PTR>(brush);
}

INT_PTR OnChangeCaseDialogCommand(HWND dlg, ChangeCaseDialogState* state, WORD commandId) noexcept
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    switch (commandId)
    {
        case IDOK:
        {
            ChangeCase::Options options{};

            if (IsDlgButtonChecked(dlg, IDC_CHANGE_CASE_UPPER) == BST_CHECKED)
            {
                options.style = ChangeCase::CaseStyle::Upper;
            }
            else if (IsDlgButtonChecked(dlg, IDC_CHANGE_CASE_PARTIALLY_MIXED) == BST_CHECKED)
            {
                options.style = ChangeCase::CaseStyle::PartiallyMixed;
            }
            else if (IsDlgButtonChecked(dlg, IDC_CHANGE_CASE_MIXED) == BST_CHECKED)
            {
                options.style = ChangeCase::CaseStyle::Mixed;
            }
            else
            {
                options.style = ChangeCase::CaseStyle::Lower;
            }

            if (IsDlgButtonChecked(dlg, IDC_CHANGE_CASE_ONLY_NAME) == BST_CHECKED)
            {
                options.target = ChangeCase::ChangeTarget::OnlyName;
            }
            else if (IsDlgButtonChecked(dlg, IDC_CHANGE_CASE_ONLY_EXTENSION) == BST_CHECKED)
            {
                options.target = ChangeCase::ChangeTarget::OnlyExtension;
            }
            else
            {
                options.target = ChangeCase::ChangeTarget::WholeFilename;
            }

            options.includeSubdirs = state->allowSubdirs && (IsDlgButtonChecked(dlg, IDC_CHANGE_CASE_INCLUDE_SUBDIRS) == BST_CHECKED);

            state->options  = options;
            state->accepted = true;
            EndDialog(dlg, IDOK);
            return TRUE;
        }
        case IDCANCEL: EndDialog(dlg, IDCANCEL); return TRUE;
    }

    return FALSE;
}

INT_PTR CALLBACK ChangeCaseDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<ChangeCaseDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnChangeCaseDialogInit(dlg, reinterpret_cast<ChangeCaseDialogState*>(lp));
        case WM_SIZE:
            if (state)
            {
                LayoutChangeCaseDialogControls(dlg, *state);
                EnsureChangeCaseDialogAllOptionsVisible(dlg, *state, false);
            }
            return TRUE;
        case WM_DPICHANGED: return OnChangeCaseDialogDpiChanged(dlg, state, static_cast<UINT>(HIWORD(wp)), reinterpret_cast<const RECT*>(lp));
        case WM_ERASEBKGND:
            // Avoid flicker; paint background and cards in WM_PAINT.
            return TRUE;
        case WM_PAINT:
        {
            if (! state)
            {
                break;
            }

            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(dlg, &ps);
            if (! hdc)
            {
                return TRUE;
            }

            RECT client{};
            GetClientRect(dlg, &client);
            const int width  = std::max(0l, client.right - client.left);
            const int height = std::max(0l, client.bottom - client.top);

            wil::unique_hdc memDc;
            wil::unique_hbitmap memBmp;
            if (width > 0 && height > 0)
            {
                memDc.reset(CreateCompatibleDC(hdc.get()));
                memBmp.reset(CreateCompatibleBitmap(hdc.get(), width, height));
            }

            if (memDc && memBmp)
            {
                [[maybe_unused]] auto oldBmp = wil::SelectObject(memDc.get(), memBmp.get());
                PaintChangeCaseDialogBackgroundAndCards(memDc.get(), dlg, *state);
                BitBlt(hdc.get(), 0, 0, width, height, memDc.get(), 0, 0, SRCCOPY);
            }
            else
            {
                PaintChangeCaseDialogBackgroundAndCards(hdc.get(), dlg, *state);
            }

            return TRUE;
        }
        case WM_CTLCOLORDLG: return OnChangeCaseDialogCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnChangeCaseDialogCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLORBTN: return OnChangeCaseDialogCtlColorButton(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->theme, wp != FALSE);
            }
            return FALSE;
        case WM_DRAWITEM:
        {
            if (! state || state->theme.highContrast || state->theme.systemHighContrast)
            {
                break;
            }

            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis || dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            const UINT id       = dis->CtlID;
            const bool isToggle = id == IDC_CHANGE_CASE_INCLUDE_SUBDIRS || id == IDC_CHANGE_CASE_LOWER || id == IDC_CHANGE_CASE_UPPER ||
                                  id == IDC_CHANGE_CASE_PARTIALLY_MIXED || id == IDC_CHANGE_CASE_MIXED || id == IDC_CHANGE_CASE_WHOLE ||
                                  id == IDC_CHANGE_CASE_ONLY_NAME || id == IDC_CHANGE_CASE_ONLY_EXTENSION;
            if (isToggle)
            {
                const bool toggledOn   = dis->hwndItem && SendMessageW(dis->hwndItem, BM_GETCHECK, 0, 0) == BST_CHECKED;
                const COLORREF surface = ThemedControls::GetControlSurfaceColor(state->theme);
                const HFONT boldFont   = state->boldFont ? state->boldFont.get() : reinterpret_cast<HFONT>(SendMessageW(dlg, WM_GETFONT, 0, 0));
                ThemedControls::DrawThemedSwitchToggle(*dis, state->theme, surface, boldFont, state->toggleOnLabel, state->toggleOffLabel, toggledOn);
                return TRUE;
            }

            ThemedControls::DrawThemedPushButton(*dis, state->theme);
            return TRUE;
        }
        case WM_COMMAND:
        {
            const WORD id         = LOWORD(wp);
            const WORD notifyCode = HIWORD(wp);
            HWND hwndCtl          = reinterpret_cast<HWND>(lp);

            if (state && notifyCode == BN_CLICKED && hwndCtl)
            {
                switch (id)
                {
                    case IDC_CHANGE_CASE_LOWER:
                    case IDC_CHANGE_CASE_UPPER:
                    case IDC_CHANGE_CASE_PARTIALLY_MIXED:
                    case IDC_CHANGE_CASE_MIXED: CheckRadioButton(dlg, IDC_CHANGE_CASE_LOWER, IDC_CHANGE_CASE_MIXED, id); return TRUE;
                    case IDC_CHANGE_CASE_WHOLE:
                    case IDC_CHANGE_CASE_ONLY_NAME:
                    case IDC_CHANGE_CASE_ONLY_EXTENSION: CheckRadioButton(dlg, IDC_CHANGE_CASE_WHOLE, IDC_CHANGE_CASE_ONLY_EXTENSION, id); return TRUE;
                    case IDC_CHANGE_CASE_INCLUDE_SUBDIRS:
                    {
                        const LONG_PTR style = GetWindowLongPtrW(hwndCtl, GWL_STYLE);
                        if ((style & BS_TYPEMASK) == BS_OWNERDRAW)
                        {
                            const bool checked = IsDlgButtonChecked(dlg, IDC_CHANGE_CASE_INCLUDE_SUBDIRS) == BST_CHECKED;
                            CheckDlgButton(dlg, IDC_CHANGE_CASE_INCLUDE_SUBDIRS, checked ? BST_UNCHECKED : BST_CHECKED);
                        }
                        break;
                    }
                    default: break;
                }
            }

            return OnChangeCaseDialogCommand(dlg, state, id);
        }
    }

    return FALSE;
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

bool LooksLikeWindowsDrivePath(std::wstring_view text) noexcept
{
    if (text.size() < 2)
    {
        return false;
    }

    const wchar_t first = text[0];
    if (! ((first >= L'A' && first <= L'Z') || (first >= L'a' && first <= L'z')))
    {
        return false;
    }

    return text[1] == L':';
}

bool LooksLikeUncPath(std::wstring_view text) noexcept
{
    return text.rfind(L"\\\\", 0) == 0 || text.rfind(L"//", 0) == 0;
}

bool LooksLikeExtendedPath(std::wstring_view text) noexcept
{
    return text.rfind(L"\\\\?\\", 0) == 0 || text.rfind(L"\\\\.\\", 0) == 0;
}

bool LooksLikeWindowsAbsolutePath(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return false;
    }

    if (LooksLikeExtendedPath(text))
    {
        return true;
    }

    if (LooksLikeUncPath(text))
    {
        return true;
    }

    return LooksLikeWindowsDrivePath(text);
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

void CenterWindowOnOwner(HWND window, HWND owner) noexcept
{
    if (! window || ! owner)
    {
        return;
    }

    RECT ownerRc{};
    RECT windowRc{};
    if (! GetWindowRect(owner, &ownerRc) || ! GetWindowRect(window, &windowRc))
    {
        return;
    }

    const int ownerW  = ownerRc.right - ownerRc.left;
    const int ownerH  = ownerRc.bottom - ownerRc.top;
    const int windowW = windowRc.right - windowRc.left;
    const int windowH = windowRc.bottom - windowRc.top;

    const int x = ownerRc.left + (ownerW - windowW) / 2;
    const int y = ownerRc.top + (ownerH - windowH) / 2;
    SetWindowPos(window, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

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

constexpr int kCreateDirectoryPathMaxLines            = 3;
constexpr UINT_PTR kCreateDirectoryNameEditSubclassId = 1;

bool IsPathBreakChar(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/' || ch == L'|';
}

int MeasureTextWidthPx(HDC hdc, std::wstring_view text) noexcept
{
    if (! hdc || text.empty())
    {
        return 0;
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        text = text.substr(0, static_cast<size_t>(std::numeric_limits<int>::max()));
    }

    SIZE extent{};
    if (GetTextExtentPoint32W(hdc, text.data(), static_cast<int>(text.size()), &extent) == 0)
    {
        return 0;
    }

    return extent.cx;
}

int FitTextChars(HDC hdc, std::wstring_view text, int widthPx) noexcept
{
    if (! hdc || text.empty() || widthPx <= 0)
    {
        return 0;
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        text = text.substr(0, static_cast<size_t>(std::numeric_limits<int>::max()));
    }

    int fitChars = 0;
    SIZE extent{};
    if (GetTextExtentExPointW(hdc, text.data(), static_cast<int>(text.size()), widthPx, &fitChars, nullptr, &extent) == 0)
    {
        return 0;
    }

    if (fitChars < 0)
    {
        return 0;
    }

    return fitChars;
}

size_t FindBreakAfterSeparator(std::wstring_view text, size_t start, size_t maxExclusive) noexcept
{
    if (start >= text.size() || maxExclusive <= start)
    {
        return std::wstring_view::npos;
    }

    maxExclusive = std::min(maxExclusive, text.size());
    for (size_t i = maxExclusive; i > start; --i)
    {
        if (IsPathBreakChar(text[i - 1]))
        {
            return i;
        }
    }

    return std::wstring_view::npos;
}

std::wstring FormatMiddleEllipsisLine(HDC hdc, std::wstring_view text, int widthPx) noexcept
{
    if (! hdc)
    {
        return std::wstring(text);
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        text = text.substr(0, static_cast<size_t>(std::numeric_limits<int>::max()));
    }

    static constexpr wchar_t kEllipsis = L'\u2026';
    const int ellipsisWidth            = MeasureTextWidthPx(hdc, std::wstring_view(&kEllipsis, 1));
    if (ellipsisWidth <= 0 || widthPx <= ellipsisWidth)
    {
        return std::wstring(1, kEllipsis);
    }

    if (MeasureTextWidthPx(hdc, text) <= widthPx)
    {
        return std::wstring(text);
    }

    const int availableForParts = widthPx - ellipsisWidth;

    size_t bestSuffixStart = std::wstring_view::npos;
    int bestSuffixWidth    = 0;

    std::vector<size_t> candidates;
    candidates.reserve(text.size() / 4);
    candidates.push_back(0);
    for (size_t i = 0; i < text.size(); ++i)
    {
        if (IsPathBreakChar(text[i]))
        {
            candidates.push_back(i);
        }
    }

    for (auto it = candidates.rbegin(); it != candidates.rend(); ++it)
    {
        const size_t candidateStart = *it;
        const int suffixWidth       = MeasureTextWidthPx(hdc, text.substr(candidateStart));
        if (suffixWidth <= availableForParts)
        {
            bestSuffixStart = candidateStart;
            bestSuffixWidth = suffixWidth;
        }
    }

    if (bestSuffixStart == std::wstring_view::npos)
    {
        for (size_t start = 0; start < text.size(); ++start)
        {
            const int suffixWidth = MeasureTextWidthPx(hdc, text.substr(start));
            if (suffixWidth <= availableForParts)
            {
                bestSuffixStart = start;
                bestSuffixWidth = suffixWidth;
                break;
            }
        }
    }

    if (bestSuffixStart == std::wstring_view::npos)
    {
        return std::wstring(1, kEllipsis);
    }

    int prefixWidthLimit = availableForParts - bestSuffixWidth;
    if (prefixWidthLimit <= 0)
    {
        std::wstring result;
        result.push_back(kEllipsis);
        result.append(text.substr(bestSuffixStart));
        return result;
    }

    int prefixChars = FitTextChars(hdc, text, prefixWidthLimit);
    if (prefixChars <= 0)
    {
        std::wstring result;
        result.push_back(kEllipsis);
        result.append(text.substr(bestSuffixStart));
        return result;
    }

    size_t prefixLen = std::min(static_cast<size_t>(prefixChars), text.size());
    prefixLen        = std::min(prefixLen, bestSuffixStart);

    const size_t breakPos = FindBreakAfterSeparator(text, 0, prefixLen);
    if (breakPos != std::wstring_view::npos && breakPos <= bestSuffixStart)
    {
        prefixLen = breakPos;
    }

    std::wstring result;
    result.append(text.substr(0, prefixLen));
    result.push_back(kEllipsis);
    result.append(text.substr(bestSuffixStart));
    return result;
}

struct WrappedPathLayout
{
    std::wstring text;
    int lineCount    = 1;
    int lineHeightPx = 0;
    bool truncated   = false;
};

WrappedPathLayout BuildWrappedPathLayout(HWND control, std::wstring_view path, int maxLines)
{
    WrappedPathLayout layout{};

    if (! control || path.empty())
    {
        layout.text = std::wstring(path);
        return layout;
    }

    RECT rc{};
    if (GetClientRect(control, &rc) == 0)
    {
        layout.text = std::wstring(path);
        return layout;
    }

    const int widthPx = rc.right - rc.left;
    if (widthPx <= 0)
    {
        layout.text = std::wstring(path);
        return layout;
    }

    const auto hdc = wil::GetDC(control);
    if (! hdc)
    {
        layout.text = std::wstring(path);
        return layout;
    }

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(control, WM_GETFONT, 0, 0));
    if (! font)
    {
        font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    [[maybe_unused]] auto restoreFont = wil::SelectObject(hdc.get(), font);

    TEXTMETRICW tm{};
    if (GetTextMetricsW(hdc.get(), &tm) == 0)
    {
        layout.text = std::wstring(path);
        return layout;
    }

    layout.lineHeightPx = tm.tmHeight + tm.tmExternalLeading;
    if (layout.lineHeightPx <= 0)
    {
        layout.text = std::wstring(path);
        return layout;
    }

    if (maxLines < 1)
    {
        maxLines = 1;
    }

    if (path.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        path = path.substr(0, static_cast<size_t>(std::numeric_limits<int>::max()));
    }

    std::wstring resultText;
    size_t start  = 0;
    int lineCount = 0;

    for (int line = 0; line < maxLines && start < path.size(); ++line)
    {
        const std::wstring_view remaining = path.substr(start);
        if (MeasureTextWidthPx(hdc.get(), remaining) <= widthPx)
        {
            resultText.append(remaining);
            ++lineCount;
            break;
        }

        if (line == (maxLines - 1))
        {
            resultText.append(FormatMiddleEllipsisLine(hdc.get(), remaining, widthPx));
            layout.truncated = true;
            ++lineCount;
            break;
        }

        int fitChars = FitTextChars(hdc.get(), remaining, widthPx);
        if (fitChars <= 0)
        {
            fitChars = 1;
        }

        const size_t limit = std::min(start + static_cast<size_t>(fitChars), path.size());
        size_t breakPos    = FindBreakAfterSeparator(path, start, limit);
        if (breakPos == std::wstring_view::npos || breakPos <= start)
        {
            breakPos = limit;
        }

        resultText.append(path.substr(start, breakPos - start));
        start = breakPos;
        ++lineCount;

        if (start < path.size())
        {
            resultText.append(L"\r\n");
        }
    }

    if (lineCount < 1)
    {
        lineCount = 1;
    }

    layout.text      = std::move(resultText);
    layout.lineCount = lineCount;
    return layout;
}

struct CreateDirectoryDialogState
{
    CreateDirectoryDialogState()                                             = default;
    CreateDirectoryDialogState(const CreateDirectoryDialogState&)            = delete;
    CreateDirectoryDialogState& operator=(const CreateDirectoryDialogState&) = delete;
    CreateDirectoryDialogState(CreateDirectoryDialogState&&)                 = delete;
    CreateDirectoryDialogState& operator=(CreateDirectoryDialogState&&)      = delete;
    ~CreateDirectoryDialogState()                                            = default;

    HWND centerOnWindow = nullptr;
    std::wstring createInPath;
    std::wstring initialName;
    std::wstring folderName;
    AppTheme theme{};
    wil::unique_hbrush backgroundBrush;
    COLORREF inputBackgroundColor         = RGB(255, 255, 255);
    COLORREF inputFocusedBackgroundColor  = RGB(255, 255, 255);
    COLORREF inputDisabledBackgroundColor = RGB(255, 255, 255);
    wil::unique_hbrush inputBrush;
    wil::unique_hbrush inputFocusedBrush;
    wil::unique_hbrush inputDisabledBrush;
    ThemedInputFrames::FrameStyle inputFrameStyle{};
    wil::unique_hwnd nameFrame;
    bool showingValidationMessage = false;
};

COLORREF ColorRefFromColorF(const D2D1::ColorF& color) noexcept
{
    const auto toByte = [](float v) noexcept
    {
        const float clamped = std::clamp(v, 0.0f, 1.0f);
        const float scaled  = (clamped * 255.0f) + 0.5f;
        const int asInt     = static_cast<int>(scaled);
        const int bounded   = std::clamp(asInt, 0, 255);
        return static_cast<BYTE>(bounded);
    };

    return RGB(toByte(color.r), toByte(color.g), toByte(color.b));
}

void ClearCreateDirectoryDialogValidation(HWND dlg, CreateDirectoryDialogState* state) noexcept
{
    if (! dlg || ! state)
    {
        return;
    }

    state->showingValidationMessage = false;
    const HWND validation           = GetDlgItem(dlg, IDC_PANE_CREATE_DIR_VALIDATION);
    if (! validation)
    {
        return;
    }

    SetWindowTextW(validation, L"");
    ShowWindow(validation, SW_HIDE);
}

void ShowCreateDirectoryDialogValidation(HWND dlg, CreateDirectoryDialogState* state, UINT messageId) noexcept
{
    if (! dlg || ! state)
    {
        return;
    }

    const HWND validation = GetDlgItem(dlg, IDC_PANE_CREATE_DIR_VALIDATION);
    if (! validation)
    {
        return;
    }

    const std::wstring message = LoadStringResource(nullptr, messageId);
    SetWindowTextW(validation, message.c_str());
    ShowWindow(validation, SW_SHOW);
    state->showingValidationMessage = true;
    InvalidateRect(validation, nullptr, TRUE);
}

void FocusCreateDirectoryNameEdit(HWND dlg) noexcept
{
    const HWND edit = dlg ? GetDlgItem(dlg, IDC_PANE_CREATE_DIR_NAME_EDIT) : nullptr;
    if (! edit)
    {
        return;
    }

    SetFocus(edit);
    SendMessageW(edit, EM_SETSEL, 0, -1);
}

void UpdateCreateDirectoryDialogValidationForInput(HWND dlg, CreateDirectoryDialogState* state) noexcept
{
    if (! dlg || ! state)
    {
        return;
    }

    wchar_t buffer[MAX_PATH] = {};
    GetDlgItemTextW(dlg, IDC_PANE_CREATE_DIR_NAME_EDIT, buffer, static_cast<int>(std::size(buffer)));

    const std::wstring_view raw(buffer);
    if (raw.empty())
    {
        ClearCreateDirectoryDialogValidation(dlg, state);
        return;
    }

    const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(raw);
    if (trimmed.empty())
    {
        ShowCreateDirectoryDialogValidation(dlg, state, IDS_MSG_PANE_CREATE_DIR_EMPTY_NAME);
        return;
    }

    if (trimmed == L"." || trimmed == L"..")
    {
        ShowCreateDirectoryDialogValidation(dlg, state, IDS_MSG_PANE_CREATE_DIR_DOT_NAME);
        return;
    }

    if (raw.find_first_of(L"\r\n\t") != std::wstring_view::npos)
    {
        ShowCreateDirectoryDialogValidation(dlg, state, IDS_MSG_PANE_CREATE_DIR_INVALID_WHITESPACE);
        return;
    }

    static constexpr std::wstring_view kInvalidNameChars = L":*?\"<>|";
    if (ContainsPathSeparators(raw) || raw.find_first_of(kInvalidNameChars) != std::wstring_view::npos)
    {
        ShowCreateDirectoryDialogValidation(dlg, state, IDS_MSG_PANE_CREATE_DIR_INVALID_CHARS);
        return;
    }

    ClearCreateDirectoryDialogValidation(dlg, state);
}

void CenterMultilineEditTextVertically(HWND edit) noexcept
{
    ThemedControls::CenterEditTextVertically(edit);
}

void PrepareFlatEditControl(HWND control) noexcept
{
    if (! control)
    {
        return;
    }

    const LONG_PTR exStyle    = GetWindowLongPtrW(control, GWL_EXSTYLE);
    const LONG_PTR style      = GetWindowLongPtrW(control, GWL_STYLE);
    const LONG_PTR newExStyle = exStyle & ~static_cast<LONG_PTR>(WS_EX_CLIENTEDGE);
    const LONG_PTR newStyle   = style & ~static_cast<LONG_PTR>(WS_BORDER);

    if (newExStyle != exStyle)
    {
        SetWindowLongPtrW(control, GWL_EXSTYLE, newExStyle);
    }
    if (newStyle != style)
    {
        SetWindowLongPtrW(control, GWL_STYLE, newStyle);
    }

    SetWindowPos(control, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    InvalidateRect(control, nullptr, TRUE);
}

void PrepareEditMargins(HWND edit) noexcept
{
    if (! edit)
    {
        return;
    }

    const UINT dpi       = GetDpiForWindow(edit);
    const int textMargin = ThemedControls::ScaleDip(dpi, 6);
    SendMessageW(edit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(textMargin, textMargin));
}

LRESULT OnCreateDirectoryNameEditPaste(HWND hwnd, WPARAM wParam, LPARAM lParam)
{
    const LRESULT result = DefSubclassProc(hwnd, WM_PASTE, wParam, lParam);

    const int length = GetWindowTextLengthW(hwnd);
    if (length <= 0)
    {
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
    SendMessageW(hwnd, EM_SETSEL, 0, -1);
    return result;
}

LRESULT CALLBACK CreateDirectoryNameEditSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR /*uIdSubclass*/, DWORD_PTR /*dwRefData*/)
{
    switch (msg)
    {
        case WM_KEYDOWN:
            if (wParam == VK_RETURN)
            {
                SendMessageW(GetParent(hwnd), WM_COMMAND, IDOK, 0);
                return 0;
            }
            break;
        case WM_CHAR:
            if (wParam == L'\r' || wParam == L'\n')
            {
                return 0;
            }
            break;
        case WM_PASTE: return OnCreateDirectoryNameEditPaste(hwnd, wParam, lParam);
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

void UpdateCreateDirectoryDialogLayout(HWND dlg, CreateDirectoryDialogState* state) noexcept
{
    if (! dlg || ! state)
    {
        return;
    }

    const HWND pathControl = GetDlgItem(dlg, IDC_PANE_CREATE_DIR_PATH);
    if (! pathControl)
    {
        return;
    }

    RECT pathRect{};
    if (GetWindowRect(pathControl, &pathRect) == 0)
    {
        return;
    }

    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&pathRect), 2);

    const int controlWidth = pathRect.right - pathRect.left;
    const int oldHeight    = pathRect.bottom - pathRect.top;

    WrappedPathLayout layout = BuildWrappedPathLayout(pathControl, state->createInPath, kCreateDirectoryPathMaxLines);
    SetWindowTextW(pathControl, layout.text.c_str());

    if (layout.lineHeightPx <= 0)
    {
        return;
    }

    int desiredLines  = std::max(1, std::min(layout.lineCount, kCreateDirectoryPathMaxLines));
    int desiredHeight = desiredLines * layout.lineHeightPx;
    desiredHeight += 2;

    const int maxHeight = (kCreateDirectoryPathMaxLines * layout.lineHeightPx) + 2;
    desiredHeight       = std::max(layout.lineHeightPx + 2, std::min(desiredHeight, maxHeight));

    if (desiredHeight == oldHeight)
    {
        return;
    }

    const int delta = desiredHeight - oldHeight;

    SetWindowPos(pathControl, nullptr, pathRect.left, pathRect.top, controlWidth, desiredHeight, SWP_NOZORDER | SWP_NOACTIVATE);

    const std::array<int, 5> moveIds = {
        IDC_PANE_CREATE_DIR_NAME_LABEL,
        IDC_PANE_CREATE_DIR_NAME_EDIT,
        IDC_PANE_CREATE_DIR_VALIDATION,
        IDOK,
        IDCANCEL,
    };

    for (const int id : moveIds)
    {
        const HWND control = GetDlgItem(dlg, id);
        if (! control)
        {
            continue;
        }

        RECT rect{};
        if (GetWindowRect(control, &rect) == 0)
        {
            continue;
        }

        MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&rect), 2);
        SetWindowPos(control, nullptr, rect.left, rect.top + delta, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    }

    if (state->nameFrame)
    {
        RECT frameRect{};
        if (GetWindowRect(state->nameFrame.get(), &frameRect) != 0)
        {
            MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&frameRect), 2);
            SetWindowPos(state->nameFrame.get(), nullptr, frameRect.left, frameRect.top + delta, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }

    RECT dialogRect{};
    if (GetWindowRect(dlg, &dialogRect) == 0)
    {
        return;
    }

    const int dialogWidth  = dialogRect.right - dialogRect.left;
    const int dialogHeight = dialogRect.bottom - dialogRect.top;
    SetWindowPos(dlg, nullptr, 0, 0, dialogWidth, dialogHeight + delta, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
}

INT_PTR OnCreateDirectoryDialogCtlColorDialog(CreateDirectoryDialogState* state)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCreateDirectoryDialogCtlColorStatic(CreateDirectoryDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    COLORREF textColor = state->theme.menu.text;
    if (control && state->showingValidationMessage)
    {
        const int controlId = GetDlgCtrlID(control);
        if (controlId == IDC_PANE_CREATE_DIR_VALIDATION)
        {
            textColor = ColorRefFromColorF(state->theme.folderView.errorText);
        }
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCreateDirectoryDialogCtlColorEdit(CreateDirectoryDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    const bool highContrast = state->theme.highContrast || state->theme.systemHighContrast;
    const bool enabled      = ! control || IsWindowEnabled(control) != FALSE;
    const bool focused      = enabled && control && GetFocus() == control;

    const COLORREF bg = enabled ? (focused ? state->inputFocusedBackgroundColor : state->inputBackgroundColor) : state->inputDisabledBackgroundColor;

    SetBkColor(hdc, bg);
    SetTextColor(hdc, enabled ? state->theme.menu.text : state->theme.menu.disabledText);

    if (highContrast)
    {
        return state->backgroundBrush ? reinterpret_cast<INT_PTR>(state->backgroundBrush.get()) : FALSE;
    }

    HBRUSH brush = nullptr;
    if (! enabled)
    {
        brush = state->inputDisabledBrush.get();
    }
    else
    {
        brush = (focused && state->inputFocusedBrush) ? state->inputFocusedBrush.get() : state->inputBrush.get();
    }

    if (brush)
    {
        return reinterpret_cast<INT_PTR>(brush);
    }

    return state->backgroundBrush ? reinterpret_cast<INT_PTR>(state->backgroundBrush.get()) : FALSE;
}

INT_PTR OnCreateDirectoryDialogInit(HWND dlg, CreateDirectoryDialogState* state)
{
    if (! state)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

    ApplyTitleBarTheme(dlg, state->theme, GetActiveWindow() == dlg);
    state->backgroundBrush.reset(CreateSolidBrush(state->theme.windowBackground));

    const bool highContrast = state->theme.highContrast || state->theme.systemHighContrast;
    if (! highContrast)
    {
        ThemedControls::EnableOwnerDrawButton(dlg, IDOK);
        ThemedControls::EnableOwnerDrawButton(dlg, IDCANCEL);
    }

    const COLORREF surface             = ThemedControls::GetControlSurfaceColor(state->theme);
    state->inputBackgroundColor        = ThemedControls::BlendColor(surface, state->theme.windowBackground, state->theme.dark ? 50 : 30, 255);
    state->inputFocusedBackgroundColor = ThemedControls::BlendColor(state->inputBackgroundColor, state->theme.menu.text, state->theme.dark ? 20 : 16, 255);
    state->inputDisabledBackgroundColor =
        ThemedControls::BlendColor(state->theme.windowBackground, state->inputBackgroundColor, state->theme.dark ? 70 : 40, 255);

    state->inputBrush.reset();
    state->inputFocusedBrush.reset();
    state->inputDisabledBrush.reset();
    if (! highContrast)
    {
        state->inputBrush.reset(CreateSolidBrush(state->inputBackgroundColor));
        state->inputFocusedBrush.reset(CreateSolidBrush(state->inputFocusedBackgroundColor));
        state->inputDisabledBrush.reset(CreateSolidBrush(state->inputDisabledBackgroundColor));
    }

    state->inputFrameStyle.theme                        = &state->theme;
    state->inputFrameStyle.backdropBrush                = state->backgroundBrush.get();
    state->inputFrameStyle.inputBackgroundColor         = state->inputBackgroundColor;
    state->inputFrameStyle.inputFocusedBackgroundColor  = state->inputFocusedBackgroundColor;
    state->inputFrameStyle.inputDisabledBackgroundColor = state->inputDisabledBackgroundColor;

    const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_CREATE_DIR);
    if (! caption.empty())
    {
        SetWindowTextW(dlg, caption.c_str());
    }

    SetDlgItemTextW(dlg, IDC_PANE_CREATE_DIR_PATH_LABEL, LoadStringResource(nullptr, IDS_LABEL_CREATE_DIR_IN).c_str());
    SetDlgItemTextW(dlg, IDC_PANE_CREATE_DIR_NAME_LABEL, LoadStringResource(nullptr, IDS_LABEL_CREATE_DIR_NAME).c_str());
    SetDlgItemTextW(dlg, IDOK, LoadStringResource(nullptr, IDS_BUTTON_CREATE).c_str());
    SetDlgItemTextW(dlg, IDCANCEL, LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL).c_str());

    ClearCreateDirectoryDialogValidation(dlg, state);
    UpdateCreateDirectoryDialogLayout(dlg, state);

    const HWND nameEdit = GetDlgItem(dlg, IDC_PANE_CREATE_DIR_NAME_EDIT);
    if (nameEdit)
    {
        SetWindowTextW(nameEdit, state->initialName.c_str());
        SendMessageW(nameEdit, EM_SETSEL, 0, -1);

        if (! highContrast)
        {
            const bool darkBackground = ChooseContrastingTextColor(state->theme.windowBackground) == RGB(255, 255, 255);
            SetWindowTheme(nameEdit, darkBackground ? L"DarkMode_Explorer" : L"Explorer", nullptr);
            SendMessageW(nameEdit, WM_THEMECHANGED, 0, 0);

            PrepareFlatEditControl(nameEdit);
            PrepareEditMargins(nameEdit);

            RECT frameRc{};
            if (GetWindowRect(nameEdit, &frameRc))
            {
                MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&frameRc), 2);

                const int frameW = std::max(0l, frameRc.right - frameRc.left);
                const int frameH = std::max(0l, frameRc.bottom - frameRc.top);

                const UINT dpi         = GetDpiForWindow(dlg);
                const int framePadding = std::max(1, ThemedControls::ScaleDip(dpi, 2));

                const int editW = std::max(1, frameW - 2 * framePadding);
                const int editH = std::max(1, frameH - 2 * framePadding);
                SetWindowPos(nameEdit, nullptr, frameRc.left + framePadding, frameRc.top + framePadding, editW, editH, SWP_NOZORDER | SWP_NOACTIVATE);

                state->nameFrame.reset(CreateWindowExW(0,
                                                       L"Static",
                                                       L"",
                                                       WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
                                                       frameRc.left,
                                                       frameRc.top,
                                                       frameW,
                                                       frameH,
                                                       dlg,
                                                       nullptr,
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));
                if (state->nameFrame)
                {
                    SetWindowPos(state->nameFrame.get(), nameEdit, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                    ThemedInputFrames::InstallFrame(state->nameFrame.get(), nameEdit, &state->inputFrameStyle);
                }
            }
        }

        CenterMultilineEditTextVertically(nameEdit);
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(nameEdit, CreateDirectoryNameEditSubclassProc, kCreateDirectoryNameEditSubclassId, 0);
#pragma warning(pop)
    }

    CenterWindowOnOwner(dlg, state->centerOnWindow);
    return TRUE;
}

INT_PTR OnCreateDirectoryDialogCommand(HWND dlg, CreateDirectoryDialogState* state, UINT commandId, UINT notifyCode)
{
    if (commandId == IDC_PANE_CREATE_DIR_NAME_EDIT && notifyCode == EN_CHANGE)
    {
        UpdateCreateDirectoryDialogValidationForInput(dlg, state);
        return TRUE;
    }

    if (commandId == IDCANCEL)
    {
        EndDialog(dlg, IDCANCEL);
        return TRUE;
    }

    if (commandId != IDOK)
    {
        return FALSE;
    }

    if (! state)
    {
        return FALSE;
    }

    ClearCreateDirectoryDialogValidation(dlg, state);

    wchar_t buffer[MAX_PATH] = {};
    GetDlgItemTextW(dlg, IDC_PANE_CREATE_DIR_NAME_EDIT, buffer, static_cast<int>(std::size(buffer)));

    std::wstring trimmed = StringUtils::TrimWhitespaceCopy(buffer);
    if (trimmed.empty())
    {
        MessageBeep(MB_ICONWARNING);
        ShowCreateDirectoryDialogValidation(dlg, state, IDS_MSG_PANE_CREATE_DIR_EMPTY_NAME);
        FocusCreateDirectoryNameEdit(dlg);
        return TRUE;
    }

    if (trimmed == L"." || trimmed == L"..")
    {
        MessageBeep(MB_ICONWARNING);
        ShowCreateDirectoryDialogValidation(dlg, state, IDS_MSG_PANE_CREATE_DIR_DOT_NAME);
        FocusCreateDirectoryNameEdit(dlg);
        return TRUE;
    }

    if (ContainsPathSeparators(trimmed))
    {
        MessageBeep(MB_ICONWARNING);
        ShowCreateDirectoryDialogValidation(dlg, state, IDS_MSG_PANE_CREATE_DIR_INVALID_CHARS);
        FocusCreateDirectoryNameEdit(dlg);
        return TRUE;
    }

    constexpr std::wstring_view kInvalidNameChars = L":*?\"<>|";
    if (trimmed.find_first_of(kInvalidNameChars) != std::wstring::npos)
    {
        MessageBeep(MB_ICONWARNING);
        ShowCreateDirectoryDialogValidation(dlg, state, IDS_MSG_PANE_CREATE_DIR_INVALID_CHARS);
        FocusCreateDirectoryNameEdit(dlg);
        return TRUE;
    }

    if (trimmed.find_first_of(L"\r\n\t") != std::wstring::npos)
    {
        MessageBeep(MB_ICONWARNING);
        ShowCreateDirectoryDialogValidation(dlg, state, IDS_MSG_PANE_CREATE_DIR_INVALID_WHITESPACE);
        FocusCreateDirectoryNameEdit(dlg);
        return TRUE;
    }

    state->folderName = std::move(trimmed);
    EndDialog(dlg, IDOK);
    return TRUE;
}

INT_PTR CALLBACK CreateDirectoryDialogProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    auto* state = reinterpret_cast<CreateDirectoryDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnCreateDirectoryDialogInit(dlg, reinterpret_cast<CreateDirectoryDialogState*>(lParam));
        case WM_CTLCOLORDLG: return OnCreateDirectoryDialogCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnCreateDirectoryDialogCtlColorStatic(state, reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_CTLCOLOREDIT: return OnCreateDirectoryDialogCtlColorEdit(state, reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->theme, wParam != FALSE);
            }
            return FALSE;
        case WM_DRAWITEM:
        {
            if (! state || state->theme.highContrast || state->theme.systemHighContrast)
            {
                break;
            }

            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (! dis || dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            ThemedControls::DrawThemedPushButton(*dis, state->theme);
            return TRUE;
        }
        case WM_COMMAND: return OnCreateDirectoryDialogCommand(dlg, state, LOWORD(wParam), HIWORD(wParam));
    }
    return FALSE;
}

std::optional<std::wstring> PromptForCreateDirectoryName(HWND ownerWindow, std::wstring_view createInPath, std::wstring_view initialName, const AppTheme& theme)
{
    CreateDirectoryDialogState state{};
    state.centerOnWindow = ownerWindow;
    state.createInPath   = std::wstring(createInPath);
    state.initialName    = std::wstring(initialName);
    state.theme          = theme;

#pragma warning(push)
    // pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
    INT_PTR result = DialogBoxParamW(
        GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PANE_CREATE_DIR), ownerWindow, CreateDirectoryDialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)

    if (result == IDOK && ! state.folderName.empty())
    {
        return state.folderName;
    }

    return std::nullopt;
}

struct MaskDialogCommonState
{
    MaskDialogCommonState()                                        = default;
    MaskDialogCommonState(const MaskDialogCommonState&)            = delete;
    MaskDialogCommonState& operator=(const MaskDialogCommonState&) = delete;
    MaskDialogCommonState(MaskDialogCommonState&&)                 = delete;
    MaskDialogCommonState& operator=(MaskDialogCommonState&&)      = delete;
    ~MaskDialogCommonState()                                       = default;

    HWND centerOnWindow                      = nullptr;
    const std::vector<std::wstring>* history = nullptr;

    AppTheme theme{};
    wil::unique_hbrush backgroundBrush;

    COLORREF inputBackgroundColor         = RGB(255, 255, 255);
    COLORREF inputFocusedBackgroundColor  = RGB(255, 255, 255);
    COLORREF inputDisabledBackgroundColor = RGB(255, 255, 255);
    wil::unique_hbrush inputBrush;
    wil::unique_hbrush inputFocusedBrush;
    wil::unique_hbrush inputDisabledBrush;

    ThemedInputFrames::FrameStyle inputFrameStyle{};
    wil::unique_hwnd inputFrame;

    wil::unique_hfont boldFont;
    UINT fontsDpi = 0;

    std::wstring hintCollapsed;
    std::wstring hintExpanded;
    std::wstring helpText;

    bool helpExpanded     = false;
    bool layoutInProgress = false;
};

void EnsureMaskDialogFonts(HWND dlg, MaskDialogCommonState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(dlg);
    if (state.fontsDpi != dpi)
    {
        state.fontsDpi = dpi;
        state.boldFont.reset();
    }

    HFONT baseFont = reinterpret_cast<HFONT>(SendMessageW(dlg, WM_GETFONT, 0, 0));
    if (! baseFont)
    {
        baseFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    LOGFONTW lf{};
    if (GetObjectW(baseFont, sizeof(lf), &lf) != sizeof(lf))
    {
        return;
    }

    if (! state.boldFont)
    {
        LOGFONTW bold = lf;
        bold.lfWeight = std::max<LONG>(FW_SEMIBOLD, lf.lfWeight);
        state.boldFont.reset(CreateFontIndirectW(&bold));
    }
}

struct SelectionMaskDialogState
{
    SelectionMaskDialogState()                                           = default;
    SelectionMaskDialogState(const SelectionMaskDialogState&)            = delete;
    SelectionMaskDialogState& operator=(const SelectionMaskDialogState&) = delete;
    SelectionMaskDialogState(SelectionMaskDialogState&&)                 = delete;
    SelectionMaskDialogState& operator=(SelectionMaskDialogState&&)      = delete;
    ~SelectionMaskDialogState()                                          = default;

    MaskDialogCommonState common{};
    std::wstring captionText;
    std::wstring labelText;
    std::wstring maskText;
};

void EnsureSelectionMaskDialogFonts(HWND dlg, SelectionMaskDialogState& state) noexcept
{
    EnsureMaskDialogFonts(dlg, state.common);
}

void UpdateSelectionMaskDialogLayout(HWND dlg, SelectionMaskDialogState* state, bool allowResize) noexcept
{
    if (! dlg || ! state || state->common.layoutInProgress)
    {
        return;
    }

    state->common.layoutInProgress = true;
    auto clearGuard                = wil::scope_exit([&] { state->common.layoutInProgress = false; });

    RECT rcClient{};
    if (GetClientRect(dlg, &rcClient) == 0)
    {
        return;
    }

    const int clientW = std::max(0l, rcClient.right - rcClient.left);
    const int clientH = std::max(0l, rcClient.bottom - rcClient.top);
    const UINT dpi    = GetDpiForWindow(dlg);

    const int margin    = ThemedControls::ScaleDip(dpi, 16);
    const int gapX      = ThemedControls::ScaleDip(dpi, 12);
    const int gapY      = ThemedControls::ScaleDip(dpi, 10);
    const int rowHeight = std::max(1, ThemedControls::ScaleDip(dpi, 26));

    const int contentX = margin;
    const int contentW = std::max(0, clientW - 2 * margin);

    HFONT dialogFont = reinterpret_cast<HFONT>(SendMessageW(dlg, WM_GETFONT, 0, 0));
    if (! dialogFont)
    {
        dialogFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    EnsureSelectionMaskDialogFonts(dlg, *state);
    const HFONT headerFont = state->common.boldFont ? state->common.boldFont.get() : dialogFont;

    const int buttonPadX = ThemedControls::ScaleDip(dpi, 16);
    const int minBtnW    = ThemedControls::ScaleDip(dpi, 80);

    const auto measureButtonWidth = [&](HWND btn) noexcept -> int
    {
        const std::wstring text = Win32Text::GetWindowTextString(btn);
        const int textW         = ThemedControls::MeasureTextWidth(dlg, dialogFont, text);
        return std::max(minBtnW, (2 * buttonPadX) + textW);
    };

    const HWND okBtn     = GetDlgItem(dlg, IDOK);
    const HWND cancelBtn = GetDlgItem(dlg, IDCANCEL);

    const int okW     = measureButtonWidth(okBtn);
    const int cancelW = measureButtonWidth(cancelBtn);

    const UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;

    int y = margin;

    const HWND label = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_LABEL);
    if (label)
    {
        const int labelH = std::max(1, ThemedControls::ScaleDip(dpi, 20));
        SetWindowPos(label, nullptr, contentX, y, contentW, labelH, flags);
        SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += labelH + gapY;
    }

    const HWND combo = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_COMBO);
    const HWND frame = state->common.inputFrame.get();
    if (combo)
    {
        const int framePadding = std::max(1, ThemedControls::ScaleDip(dpi, 2));

        if (frame)
        {
            SetWindowPos(frame, nullptr, contentX, y, contentW, rowHeight, flags);
        }

        SetWindowPos(combo,
                     nullptr,
                     contentX + framePadding,
                     y + framePadding,
                     std::max(1, contentW - 2 * framePadding),
                     std::max(1, rowHeight - 2 * framePadding),
                     flags);
        if (frame)
        {
            SetWindowPos(frame, combo, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        }

        ThemedControls::EnsureComboBoxDroppedWidth(combo, dpi);
        y += rowHeight + gapY;
    }

    const HWND hint = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_HINT);
    if (hint)
    {
        const int hintH = std::max(1, ThemedControls::ScaleDip(dpi, 18));
        SetWindowPos(hint, nullptr, contentX, y, contentW, hintH, flags);
        SendMessageW(hint, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        y += hintH + gapY;
    }

    const HWND help = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_HELP);
    if (help && state->common.helpExpanded && ! state->common.helpText.empty())
    {
        const int helpH = MeasureStaticTextHeight(dlg, dialogFont, contentW, state->common.helpText);
        SetWindowPos(help, nullptr, contentX, y, contentW, std::max(1, helpH), flags);
        SendMessageW(help, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        ShowWindow(help, SW_SHOW);
        y += std::max(1, helpH) + gapY;
    }
    else if (help)
    {
        ShowWindow(help, SW_HIDE);
    }

    const int buttonsY       = y;
    const int desiredClientH = buttonsY + rowHeight + margin;

    if (allowResize && desiredClientH > 0 && desiredClientH != clientH)
    {
        RECT rcWindow{};
        if (GetWindowRect(dlg, &rcWindow) != 0)
        {
            const int windowW    = rcWindow.right - rcWindow.left;
            const int windowH    = rcWindow.bottom - rcWindow.top;
            const int nonClientH = std::max(0, windowH - clientH);
            const int newWindowH = desiredClientH + nonClientH;
            SetWindowPos(dlg, nullptr, 0, 0, windowW, newWindowH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }

    int nextRight = std::max(0, clientW - margin);
    if (cancelBtn)
    {
        nextRight -= cancelW;
        SetWindowPos(cancelBtn, nullptr, nextRight, buttonsY, cancelW, rowHeight, flags);
        SendMessageW(cancelBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        nextRight -= gapX;
    }
    if (okBtn)
    {
        nextRight -= okW;
        SetWindowPos(okBtn, nullptr, nextRight, buttonsY, okW, rowHeight, flags);
        SendMessageW(okBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    InvalidateRect(dlg, nullptr, TRUE);
}

INT_PTR OnSelectionMaskDialogCtlColorDialog(SelectionMaskDialogState* state)
{
    if (! state || ! state->common.backgroundBrush)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->common.backgroundBrush.get());
}

INT_PTR OnSelectionMaskDialogCtlColorStatic(SelectionMaskDialogState* state, HDC hdc, HWND)
{
    if (! state || ! state->common.backgroundBrush)
    {
        return FALSE;
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, state->common.theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->common.backgroundBrush.get());
}

INT_PTR OnSelectionMaskDialogCtlColorEdit(SelectionMaskDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    const bool highContrast = state->common.theme.highContrast || state->common.theme.systemHighContrast;
    const bool enabled      = ! control || IsWindowEnabled(control) != FALSE;
    const bool focused      = enabled && control && GetFocus() == control;

    const COLORREF bg =
        enabled ? (focused ? state->common.inputFocusedBackgroundColor : state->common.inputBackgroundColor) : state->common.inputDisabledBackgroundColor;

    SetBkColor(hdc, bg);
    SetTextColor(hdc, enabled ? state->common.theme.menu.text : state->common.theme.menu.disabledText);

    if (highContrast)
    {
        return state->common.backgroundBrush ? reinterpret_cast<INT_PTR>(state->common.backgroundBrush.get()) : FALSE;
    }

    HBRUSH brush = nullptr;
    if (! enabled)
    {
        brush = state->common.inputDisabledBrush.get();
    }
    else
    {
        brush = (focused && state->common.inputFocusedBrush) ? state->common.inputFocusedBrush.get() : state->common.inputBrush.get();
    }

    if (brush)
    {
        return reinterpret_cast<INT_PTR>(brush);
    }

    return state->common.backgroundBrush ? reinterpret_cast<INT_PTR>(state->common.backgroundBrush.get()) : FALSE;
}

INT_PTR OnSelectionMaskDialogInit(HWND dlg, SelectionMaskDialogState* state)
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

    ApplyTitleBarTheme(dlg, state->common.theme, GetActiveWindow() == dlg);
    state->common.backgroundBrush.reset(CreateSolidBrush(state->common.theme.windowBackground));

    const bool highContrast = state->common.theme.highContrast || state->common.theme.systemHighContrast;
    if (! highContrast)
    {
        ThemedControls::EnableOwnerDrawButton(dlg, IDOK);
        ThemedControls::EnableOwnerDrawButton(dlg, IDCANCEL);
    }

    const COLORREF surface             = ThemedControls::GetControlSurfaceColor(state->common.theme);
    state->common.inputBackgroundColor = ThemedControls::BlendColor(surface, state->common.theme.windowBackground, state->common.theme.dark ? 50 : 30, 255);
    state->common.inputFocusedBackgroundColor =
        ThemedControls::BlendColor(state->common.inputBackgroundColor, state->common.theme.menu.text, state->common.theme.dark ? 20 : 16, 255);
    state->common.inputDisabledBackgroundColor =
        ThemedControls::BlendColor(state->common.theme.windowBackground, state->common.inputBackgroundColor, state->common.theme.dark ? 70 : 40, 255);

    state->common.inputBrush.reset();
    state->common.inputFocusedBrush.reset();
    state->common.inputDisabledBrush.reset();
    if (! highContrast)
    {
        state->common.inputBrush.reset(CreateSolidBrush(state->common.inputBackgroundColor));
        state->common.inputFocusedBrush.reset(CreateSolidBrush(state->common.inputFocusedBackgroundColor));
        state->common.inputDisabledBrush.reset(CreateSolidBrush(state->common.inputDisabledBackgroundColor));
    }

    state->common.inputFrameStyle.theme                        = &state->common.theme;
    state->common.inputFrameStyle.backdropBrush                = state->common.backgroundBrush.get();
    state->common.inputFrameStyle.inputBackgroundColor         = state->common.inputBackgroundColor;
    state->common.inputFrameStyle.inputFocusedBackgroundColor  = state->common.inputFocusedBackgroundColor;
    state->common.inputFrameStyle.inputDisabledBackgroundColor = state->common.inputDisabledBackgroundColor;

    if (! state->captionText.empty())
    {
        SetWindowTextW(dlg, state->captionText.c_str());
    }

    if (! state->labelText.empty())
    {
        SetDlgItemTextW(dlg, IDC_PANE_SELECTION_MASK_LABEL, state->labelText.c_str());
    }

    SetDlgItemTextW(dlg, IDOK, LoadStringResource(nullptr, IDS_BTN_OK).c_str());
    SetDlgItemTextW(dlg, IDCANCEL, LoadStringResource(nullptr, IDS_BTN_CANCEL).c_str());

    state->common.hintCollapsed = LoadStringResource(nullptr, IDS_SELECTION_MASK_HINT_COLLAPSED);
    state->common.hintExpanded  = LoadStringResource(nullptr, IDS_SELECTION_MASK_HINT_EXPANDED);
    state->common.helpText      = LoadStringResource(nullptr, IDS_SELECTION_MASK_HELP_TEXT);

    if (const HWND hint = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_HINT))
    {
        const std::wstring& text = state->common.helpExpanded ? state->common.hintExpanded : state->common.hintCollapsed;
        SetWindowTextW(hint, text.c_str());
    }

    if (const HWND help = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_HELP))
    {
        SetWindowTextW(help, state->common.helpText.c_str());
        ShowWindow(help, state->common.helpExpanded ? SW_SHOW : SW_HIDE);
    }

    const HWND combo = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_COMBO);
    if (combo)
    {
        SendMessageW(combo, CB_RESETCONTENT, 0, 0);
        if (state->common.history)
        {
            for (const auto& entry : *state->common.history)
            {
                if (entry.empty())
                {
                    continue;
                }
                SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(entry.c_str()));
            }

            if (! state->common.history->empty() && ! state->common.history->front().empty())
            {
                SetWindowTextW(combo, state->common.history->front().c_str());
                SendMessageW(combo, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
            }
        }

        ThemedControls::ApplyThemeToComboBox(combo, state->common.theme);

        if (! highContrast)
        {
            PrepareFlatEditControl(combo);

            COMBOBOXINFO cbi{};
            cbi.cbSize = sizeof(cbi);
            if (GetComboBoxInfo(combo, &cbi) && cbi.hwndItem)
            {
                PrepareEditMargins(cbi.hwndItem);
            }

            RECT frameRc{};
            if (GetWindowRect(combo, &frameRc))
            {
                MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&frameRc), 2);

                const int frameW = std::max(0l, frameRc.right - frameRc.left);
                const int frameH = std::max(0l, frameRc.bottom - frameRc.top);

                const UINT frameDpi    = GetDpiForWindow(dlg);
                const int framePadding = std::max(1, ThemedControls::ScaleDip(frameDpi, 2));
                const int innerW       = std::max(1, frameW - 2 * framePadding);
                const int innerH       = std::max(1, frameH - 2 * framePadding);

                SetWindowPos(combo, nullptr, frameRc.left + framePadding, frameRc.top + framePadding, innerW, innerH, SWP_NOZORDER | SWP_NOACTIVATE);

                state->common.inputFrame.reset(CreateWindowExW(0,
                                                               L"Static",
                                                               L"",
                                                               WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
                                                               frameRc.left,
                                                               frameRc.top,
                                                               frameW,
                                                               frameH,
                                                               dlg,
                                                               nullptr,
                                                               GetModuleHandleW(nullptr),
                                                               nullptr));
                if (state->common.inputFrame)
                {
                    SetWindowPos(state->common.inputFrame.get(), combo, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                    ThemedInputFrames::InstallFrame(state->common.inputFrame.get(), combo, &state->common.inputFrameStyle);
                }
            }
        }
    }

    UpdateSelectionMaskDialogLayout(dlg, state, true);
    CenterWindowOnOwner(dlg, state->common.centerOnWindow);

    if (combo)
    {
        SetFocus(combo);
        SendMessageW(combo, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
        return FALSE;
    }

    return TRUE;
}

INT_PTR OnSelectionMaskDialogCommand(HWND dlg, SelectionMaskDialogState* state, UINT commandId, UINT notifyCode)
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    if (commandId == IDC_PANE_SELECTION_MASK_HINT && notifyCode == STN_CLICKED)
    {
        state->common.helpExpanded = ! state->common.helpExpanded;

        const HWND hint = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_HINT);
        if (hint)
        {
            const std::wstring& text = state->common.helpExpanded ? state->common.hintExpanded : state->common.hintCollapsed;
            SetWindowTextW(hint, text.c_str());
        }

        UpdateSelectionMaskDialogLayout(dlg, state, true);
        return TRUE;
    }

    if (commandId == IDC_PANE_SELECTION_MASK_COMBO && notifyCode == CBN_DROPDOWN)
    {
        if (const HWND combo = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_COMBO))
        {
            ThemedControls::ApplyThemeToComboBoxDropDown(combo, state->common.theme);
        }
        return TRUE;
    }

    if (commandId == IDCANCEL)
    {
        EndDialog(dlg, IDCANCEL);
        return TRUE;
    }

    if (commandId != IDOK)
    {
        return FALSE;
    }

    const HWND combo = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_COMBO);
    if (! combo)
    {
        return TRUE;
    }

    const std::wstring text = Win32Text::GetWindowTextString(combo);
    std::wstring trimmed    = StringUtils::TrimWhitespaceCopy(text);
    if (trimmed.empty())
    {
        MessageBeep(MB_ICONWARNING);
        SetFocus(combo);
        return TRUE;
    }

    state->maskText = std::move(trimmed);
    EndDialog(dlg, IDOK);
    return TRUE;
}

INT_PTR CALLBACK SelectionMaskDialogProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    auto* state = reinterpret_cast<SelectionMaskDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnSelectionMaskDialogInit(dlg, reinterpret_cast<SelectionMaskDialogState*>(lParam));
        case WM_CTLCOLORDLG: return OnSelectionMaskDialogCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnSelectionMaskDialogCtlColorStatic(state, reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORLISTBOX: return OnSelectionMaskDialogCtlColorEdit(state, reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->common.theme, wParam != FALSE);
            }
            return FALSE;
        case WM_SETCURSOR:
            if (state)
            {
                const HWND hover = reinterpret_cast<HWND>(wParam);
                const HWND hint  = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_HINT);
                if (hover && hint && hover == hint)
                {
                    SetCursor(LoadCursorW(nullptr, IDC_HAND));
                    return TRUE;
                }
            }
            break;
        case WM_SIZE:
            if (state)
            {
                UpdateSelectionMaskDialogLayout(dlg, state, false);
                return TRUE;
            }
            break;
        case WM_DPICHANGED:
            if (state)
            {
                const RECT* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    const int w = std::max(1l, suggested->right - suggested->left);
                    const int h = std::max(1l, suggested->bottom - suggested->top);
                    SetWindowPos(dlg, nullptr, suggested->left, suggested->top, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
                }
                UpdateSelectionMaskDialogLayout(dlg, state, true);
                return TRUE;
            }
            break;
        case WM_DRAWITEM:
        {
            if (! state || state->common.theme.highContrast || state->common.theme.systemHighContrast)
            {
                break;
            }

            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (! dis || dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            ThemedControls::DrawThemedPushButton(*dis, state->common.theme);
            return TRUE;
        }
        case WM_COMMAND: return OnSelectionMaskDialogCommand(dlg, state, LOWORD(wParam), HIWORD(wParam));
    }
    return FALSE;
}

std::optional<std::wstring> PromptForSelectionMask(
    HWND ownerWindow, const std::vector<std::wstring>& history, const AppTheme& theme, UINT captionId, UINT labelId)
{
    SelectionMaskDialogState state{};
    state.common.centerOnWindow = ownerWindow;
    state.common.history        = &history;
    state.common.theme          = theme;
    state.common.helpExpanded   = false;
    state.captionText           = LoadStringResource(nullptr, captionId);
    state.labelText             = LoadStringResource(nullptr, labelId);

#pragma warning(push)
// pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
    const INT_PTR result = DialogBoxParamW(
        GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PANE_SELECTION_MASK), ownerWindow, SelectionMaskDialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)

    if (result == IDOK && ! state.maskText.empty())
    {
        return state.maskText;
    }

    return std::nullopt;
}

struct PaneFilterDialogState
{
    PaneFilterDialogState()                                        = default;
    PaneFilterDialogState(const PaneFilterDialogState&)            = delete;
    PaneFilterDialogState& operator=(const PaneFilterDialogState&) = delete;
    PaneFilterDialogState(PaneFilterDialogState&&)                 = delete;
    PaneFilterDialogState& operator=(PaneFilterDialogState&&)      = delete;
    ~PaneFilterDialogState()                                       = default;

    MaskDialogCommonState common{};
    std::wstring captionText;
    std::wstring useLabelText;
    std::wstring labelText;

    std::wstring toggleOnLabel;
    std::wstring toggleOffLabel;

    FolderView::NameFilterState initial;
    FolderView::NameFilterState result;
};

void EnsurePaneFilterDialogFonts(HWND dlg, PaneFilterDialogState& state) noexcept
{
    EnsureMaskDialogFonts(dlg, state.common);
}

void UpdatePaneFilterDialogLayout(HWND dlg, PaneFilterDialogState* state, bool allowResize) noexcept
{
    if (! dlg || ! state || state->common.layoutInProgress)
    {
        return;
    }

    state->common.layoutInProgress = true;
    auto clearGuard                = wil::scope_exit([&] { state->common.layoutInProgress = false; });

    RECT rcClient{};
    if (GetClientRect(dlg, &rcClient) == 0)
    {
        return;
    }

    const int clientW = std::max(0l, rcClient.right - rcClient.left);
    const int clientH = std::max(0l, rcClient.bottom - rcClient.top);
    const UINT dpi    = GetDpiForWindow(dlg);

    const int margin    = ThemedControls::ScaleDip(dpi, 16);
    const int gapX      = ThemedControls::ScaleDip(dpi, 12);
    const int gapY      = ThemedControls::ScaleDip(dpi, 10);
    const int rowHeight = std::max(1, ThemedControls::ScaleDip(dpi, 26));

    const int contentX = margin;
    const int contentW = std::max(0, clientW - 2 * margin);

    HFONT dialogFont = reinterpret_cast<HFONT>(SendMessageW(dlg, WM_GETFONT, 0, 0));
    if (! dialogFont)
    {
        dialogFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    EnsurePaneFilterDialogFonts(dlg, *state);
    const HFONT headerFont = state->common.boldFont ? state->common.boldFont.get() : dialogFont;

    const int buttonPadX = ThemedControls::ScaleDip(dpi, 16);
    const int minBtnW    = ThemedControls::ScaleDip(dpi, 80);

    const auto measureButtonWidth = [&](HWND btn) noexcept -> int
    {
        const std::wstring text = Win32Text::GetWindowTextString(btn);
        const int textW         = ThemedControls::MeasureTextWidth(dlg, dialogFont, text);
        return std::max(minBtnW, (2 * buttonPadX) + textW);
    };

    const HWND okBtn     = GetDlgItem(dlg, IDOK);
    const HWND cancelBtn = GetDlgItem(dlg, IDCANCEL);

    const int okW     = measureButtonWidth(okBtn);
    const int cancelW = measureButtonWidth(cancelBtn);

    const UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;

    int y = margin;

    // Use Filter row.
    const HWND useLabel = GetDlgItem(dlg, IDC_PANE_FILTER_USE_LABEL);
    const HWND toggle   = GetDlgItem(dlg, IDC_PANE_FILTER_USE_TOGGLE);
    if (useLabel && toggle)
    {
        const int labelH       = std::max(1, ThemedControls::ScaleDip(dpi, 20));
        const int minToggleW   = ThemedControls::ScaleDip(dpi, 90);
        const int paddingX     = ThemedControls::ScaleDip(dpi, 6);
        const int toggleGapX   = ThemedControls::ScaleDip(dpi, 8);
        const int trackWidth   = ThemedControls::ScaleDip(dpi, 34);
        const int stateTextW   = std::max(ThemedControls::MeasureTextWidth(dlg, headerFont, state->toggleOnLabel),
                                        ThemedControls::MeasureTextWidth(dlg, headerFont, state->toggleOffLabel));
        const int measuredW    = std::max(minToggleW, (2 * paddingX) + stateTextW + toggleGapX + trackWidth);
        const int toggleW      = std::min(std::max(0, contentW), measuredW);
        const int labelW       = std::max(0, contentW - toggleW - gapX);
        const int labelYOffset = (rowHeight - labelH) / 2;

        SetWindowPos(useLabel, nullptr, contentX, y + std::max(0, labelYOffset), labelW, labelH, flags);
        SendMessageW(useLabel, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        SetWindowPos(toggle, nullptr, contentX + contentW - toggleW, y, toggleW, rowHeight, flags);
        SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        y += rowHeight + gapY;
    }

    const HWND label = GetDlgItem(dlg, IDC_PANE_FILTER_LABEL);
    if (label)
    {
        const int labelH = std::max(1, ThemedControls::ScaleDip(dpi, 20));
        SetWindowPos(label, nullptr, contentX, y, contentW, labelH, flags);
        SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += labelH + gapY;
    }

    const HWND combo = GetDlgItem(dlg, IDC_PANE_FILTER_COMBO);
    const HWND frame = state->common.inputFrame.get();
    if (combo)
    {
        const int framePadding = std::max(1, ThemedControls::ScaleDip(dpi, 2));

        if (frame)
        {
            SetWindowPos(frame, nullptr, contentX, y, contentW, rowHeight, flags);
        }

        SetWindowPos(combo,
                     nullptr,
                     contentX + framePadding,
                     y + framePadding,
                     std::max(1, contentW - 2 * framePadding),
                     std::max(1, rowHeight - 2 * framePadding),
                     flags);
        if (frame)
        {
            SetWindowPos(frame, combo, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        }

        ThemedControls::EnsureComboBoxDroppedWidth(combo, dpi);
        y += rowHeight + gapY;
    }

    const HWND hint = GetDlgItem(dlg, IDC_PANE_FILTER_HINT);
    if (hint)
    {
        const int hintH = std::max(1, ThemedControls::ScaleDip(dpi, 18));
        SetWindowPos(hint, nullptr, contentX, y, contentW, hintH, flags);
        SendMessageW(hint, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        y += hintH + gapY;
    }

    const HWND help = GetDlgItem(dlg, IDC_PANE_FILTER_HELP);
    if (help && state->common.helpExpanded && ! state->common.helpText.empty())
    {
        const int helpH = MeasureStaticTextHeight(dlg, dialogFont, contentW, state->common.helpText);
        SetWindowPos(help, nullptr, contentX, y, contentW, std::max(1, helpH), flags);
        SendMessageW(help, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        ShowWindow(help, SW_SHOW);
        y += std::max(1, helpH) + gapY;
    }
    else if (help)
    {
        ShowWindow(help, SW_HIDE);
    }

    const int buttonsY       = y;
    const int desiredClientH = buttonsY + rowHeight + margin;

    if (allowResize && desiredClientH > 0 && desiredClientH != clientH)
    {
        RECT rcWindow{};
        if (GetWindowRect(dlg, &rcWindow) != 0)
        {
            const int windowW    = rcWindow.right - rcWindow.left;
            const int windowH    = rcWindow.bottom - rcWindow.top;
            const int nonClientH = std::max(0, windowH - clientH);
            const int newWindowH = desiredClientH + nonClientH;
            SetWindowPos(dlg, nullptr, 0, 0, windowW, newWindowH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }

    int nextRight = std::max(0, clientW - margin);
    if (cancelBtn)
    {
        nextRight -= cancelW;
        SetWindowPos(cancelBtn, nullptr, nextRight, buttonsY, cancelW, rowHeight, flags);
        SendMessageW(cancelBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        nextRight -= gapX;
    }
    if (okBtn)
    {
        nextRight -= okW;
        SetWindowPos(okBtn, nullptr, nextRight, buttonsY, okW, rowHeight, flags);
        SendMessageW(okBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    InvalidateRect(dlg, nullptr, TRUE);
}

INT_PTR OnPaneFilterDialogCtlColorDialog(PaneFilterDialogState* state)
{
    if (! state || ! state->common.backgroundBrush)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->common.backgroundBrush.get());
}

INT_PTR OnPaneFilterDialogCtlColorStatic(PaneFilterDialogState* state, HDC hdc, HWND)
{
    if (! state || ! state->common.backgroundBrush)
    {
        return FALSE;
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, state->common.theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->common.backgroundBrush.get());
}

INT_PTR OnPaneFilterDialogCtlColorEdit(PaneFilterDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    const bool highContrast = state->common.theme.highContrast || state->common.theme.systemHighContrast;
    const bool enabled      = ! control || IsWindowEnabled(control) != FALSE;
    const bool focused      = enabled && control && GetFocus() == control;

    const COLORREF bg =
        enabled ? (focused ? state->common.inputFocusedBackgroundColor : state->common.inputBackgroundColor) : state->common.inputDisabledBackgroundColor;

    SetBkColor(hdc, bg);
    SetTextColor(hdc, enabled ? state->common.theme.menu.text : state->common.theme.menu.disabledText);

    if (highContrast)
    {
        return state->common.backgroundBrush ? reinterpret_cast<INT_PTR>(state->common.backgroundBrush.get()) : FALSE;
    }

    HBRUSH brush = nullptr;
    if (! enabled)
    {
        brush = state->common.inputDisabledBrush.get();
    }
    else
    {
        brush = (focused && state->common.inputFocusedBrush) ? state->common.inputFocusedBrush.get() : state->common.inputBrush.get();
    }

    if (brush)
    {
        return reinterpret_cast<INT_PTR>(brush);
    }

    return state->common.backgroundBrush ? reinterpret_cast<INT_PTR>(state->common.backgroundBrush.get()) : FALSE;
}

INT_PTR OnPaneFilterDialogInit(HWND dlg, PaneFilterDialogState* state)
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

    ApplyTitleBarTheme(dlg, state->common.theme, GetActiveWindow() == dlg);
    const COLORREF surface = ThemedControls::GetControlSurfaceColor(state->common.theme);

    const bool highContrast         = state->common.theme.highContrast || state->common.theme.systemHighContrast;
    const COLORREF dialogBackground = highContrast ? state->common.theme.windowBackground : surface;
    state->common.backgroundBrush.reset(CreateSolidBrush(dialogBackground));
    if (! highContrast)
    {
        ThemedControls::EnableOwnerDrawButton(dlg, IDOK);
        ThemedControls::EnableOwnerDrawButton(dlg, IDCANCEL);
        ThemedControls::EnableOwnerDrawButton(dlg, IDC_PANE_FILTER_USE_TOGGLE);
    }

    state->common.inputBackgroundColor = ThemedControls::BlendColor(surface, state->common.theme.windowBackground, state->common.theme.dark ? 50 : 30, 255);
    state->common.inputFocusedBackgroundColor =
        ThemedControls::BlendColor(state->common.inputBackgroundColor, state->common.theme.menu.text, state->common.theme.dark ? 20 : 16, 255);
    state->common.inputDisabledBackgroundColor =
        ThemedControls::BlendColor(dialogBackground, state->common.inputBackgroundColor, state->common.theme.dark ? 70 : 40, 255);

    state->common.inputBrush.reset();
    state->common.inputFocusedBrush.reset();
    state->common.inputDisabledBrush.reset();
    if (! highContrast)
    {
        state->common.inputBrush.reset(CreateSolidBrush(state->common.inputBackgroundColor));
        state->common.inputFocusedBrush.reset(CreateSolidBrush(state->common.inputFocusedBackgroundColor));
        state->common.inputDisabledBrush.reset(CreateSolidBrush(state->common.inputDisabledBackgroundColor));
    }

    state->common.inputFrameStyle.theme                        = &state->common.theme;
    state->common.inputFrameStyle.backdropBrush                = state->common.backgroundBrush.get();
    state->common.inputFrameStyle.inputBackgroundColor         = state->common.inputBackgroundColor;
    state->common.inputFrameStyle.inputFocusedBackgroundColor  = state->common.inputFocusedBackgroundColor;
    state->common.inputFrameStyle.inputDisabledBackgroundColor = state->common.inputDisabledBackgroundColor;

    if (! state->captionText.empty())
    {
        SetWindowTextW(dlg, state->captionText.c_str());
    }

    if (! state->useLabelText.empty())
    {
        SetDlgItemTextW(dlg, IDC_PANE_FILTER_USE_LABEL, state->useLabelText.c_str());
    }

    if (! state->labelText.empty())
    {
        SetDlgItemTextW(dlg, IDC_PANE_FILTER_LABEL, state->labelText.c_str());
    }

    SetDlgItemTextW(dlg, IDOK, LoadStringResource(nullptr, IDS_BTN_OK).c_str());
    SetDlgItemTextW(dlg, IDCANCEL, LoadStringResource(nullptr, IDS_BTN_CANCEL).c_str());

    state->toggleOnLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    state->toggleOffLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    const HWND toggle = GetDlgItem(dlg, IDC_PANE_FILTER_USE_TOGGLE);
    if (toggle)
    {
        if (! highContrast)
        {
            SetWindowSubclass(toggle, ChangeCaseToggleCheckSubclassProc, 1u, 0);
        }
        CheckDlgButton(dlg, IDC_PANE_FILTER_USE_TOGGLE, state->initial.enabled ? BST_CHECKED : BST_UNCHECKED);
    }

    state->common.hintCollapsed = LoadStringResource(nullptr, IDS_SELECTION_MASK_HINT_COLLAPSED);
    state->common.hintExpanded  = LoadStringResource(nullptr, IDS_SELECTION_MASK_HINT_EXPANDED);
    state->common.helpText      = LoadStringResource(nullptr, IDS_SELECTION_MASK_HELP_TEXT);

    if (const HWND hint = GetDlgItem(dlg, IDC_PANE_FILTER_HINT))
    {
        const std::wstring& text = state->common.helpExpanded ? state->common.hintExpanded : state->common.hintCollapsed;
        SetWindowTextW(hint, text.c_str());
    }

    if (const HWND help = GetDlgItem(dlg, IDC_PANE_FILTER_HELP))
    {
        SetWindowTextW(help, state->common.helpText.c_str());
        ShowWindow(help, state->common.helpExpanded ? SW_SHOW : SW_HIDE);
    }

    const HWND combo = GetDlgItem(dlg, IDC_PANE_FILTER_COMBO);
    if (combo)
    {
        SendMessageW(combo, CB_RESETCONTENT, 0, 0);
        if (state->common.history)
        {
            for (const auto& entry : *state->common.history)
            {
                if (entry.empty())
                {
                    continue;
                }
                SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(entry.c_str()));
            }
        }

        std::wstring initialText = StringUtils::TrimWhitespaceCopy(state->initial.text);
        if (initialText.empty() && state->common.history && ! state->common.history->empty())
        {
            initialText = state->common.history->front();
        }
        if (! initialText.empty())
        {
            SetWindowTextW(combo, initialText.c_str());
            SendMessageW(combo, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
        }

        ThemedControls::ApplyThemeToComboBox(combo, state->common.theme);

        if (! highContrast)
        {
            PrepareFlatEditControl(combo);

            COMBOBOXINFO cbi{};
            cbi.cbSize = sizeof(cbi);
            if (GetComboBoxInfo(combo, &cbi) && cbi.hwndItem)
            {
                PrepareEditMargins(cbi.hwndItem);
            }

            RECT frameRc{};
            if (GetWindowRect(combo, &frameRc))
            {
                MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&frameRc), 2);

                const int frameW = std::max(0l, frameRc.right - frameRc.left);
                const int frameH = std::max(0l, frameRc.bottom - frameRc.top);

                const UINT frameDpi    = GetDpiForWindow(dlg);
                const int framePadding = std::max(1, ThemedControls::ScaleDip(frameDpi, 2));
                const int innerW       = std::max(1, frameW - 2 * framePadding);
                const int innerH       = std::max(1, frameH - 2 * framePadding);

                SetWindowPos(combo, nullptr, frameRc.left + framePadding, frameRc.top + framePadding, innerW, innerH, SWP_NOZORDER | SWP_NOACTIVATE);

                state->common.inputFrame.reset(CreateWindowExW(0,
                                                               L"Static",
                                                               L"",
                                                               WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
                                                               frameRc.left,
                                                               frameRc.top,
                                                               frameW,
                                                               frameH,
                                                               dlg,
                                                               nullptr,
                                                               GetModuleHandleW(nullptr),
                                                               nullptr));
                if (state->common.inputFrame)
                {
                    SetWindowPos(state->common.inputFrame.get(), combo, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                    ThemedInputFrames::InstallFrame(state->common.inputFrame.get(), combo, &state->common.inputFrameStyle);
                }
            }
        }
    }

    UpdatePaneFilterDialogLayout(dlg, state, true);
    CenterWindowOnOwner(dlg, state->common.centerOnWindow);

    if (combo)
    {
        SetFocus(combo);
        SendMessageW(combo, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
        return FALSE;
    }

    return TRUE;
}

INT_PTR OnPaneFilterDialogCommand(HWND dlg, PaneFilterDialogState* state, UINT commandId, UINT notifyCode)
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    if (commandId == IDC_PANE_FILTER_HINT && notifyCode == STN_CLICKED)
    {
        state->common.helpExpanded = ! state->common.helpExpanded;

        const HWND hint = GetDlgItem(dlg, IDC_PANE_FILTER_HINT);
        if (hint)
        {
            const std::wstring& text = state->common.helpExpanded ? state->common.hintExpanded : state->common.hintCollapsed;
            SetWindowTextW(hint, text.c_str());
        }

        UpdatePaneFilterDialogLayout(dlg, state, true);
        return TRUE;
    }

    if (commandId == IDC_PANE_FILTER_COMBO && notifyCode == CBN_DROPDOWN)
    {
        if (const HWND combo = GetDlgItem(dlg, IDC_PANE_FILTER_COMBO))
        {
            ThemedControls::ApplyThemeToComboBoxDropDown(combo, state->common.theme);
        }
        return TRUE;
    }

    if (commandId == IDC_PANE_FILTER_USE_TOGGLE && notifyCode == BN_CLICKED)
    {
        if (const HWND toggle = GetDlgItem(dlg, IDC_PANE_FILTER_USE_TOGGLE))
        {
            const LONG_PTR style = GetWindowLongPtrW(toggle, GWL_STYLE);
            if ((style & BS_TYPEMASK) == BS_OWNERDRAW)
            {
                const bool checked = IsDlgButtonChecked(dlg, IDC_PANE_FILTER_USE_TOGGLE) == BST_CHECKED;
                CheckDlgButton(dlg, IDC_PANE_FILTER_USE_TOGGLE, checked ? BST_UNCHECKED : BST_CHECKED);
            }

            InvalidateRect(toggle, nullptr, TRUE);
        }
        return TRUE;
    }

    if (commandId == IDCANCEL)
    {
        EndDialog(dlg, IDCANCEL);
        return TRUE;
    }

    if (commandId != IDOK)
    {
        return FALSE;
    }

    const bool enabled = IsDlgButtonChecked(dlg, IDC_PANE_FILTER_USE_TOGGLE) == BST_CHECKED;

    const HWND combo = GetDlgItem(dlg, IDC_PANE_FILTER_COMBO);
    if (! combo)
    {
        state->result.enabled = enabled;
        state->result.text.clear();
        EndDialog(dlg, IDOK);
        return TRUE;
    }

    std::wstring trimmed = StringUtils::TrimWhitespaceCopy(Win32Text::GetWindowTextString(combo));

    if (enabled)
    {
        const MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(trimmed);
        const bool hasMask                  = ! mask.includePatterns.empty() || ! mask.excludePatterns.empty();
        if (! hasMask)
        {
            MessageBeep(MB_ICONWARNING);
            SetFocus(combo);
            return TRUE;
        }
    }

    state->result.enabled = enabled;
    state->result.text    = std::move(trimmed);
    EndDialog(dlg, IDOK);
    return TRUE;
}

INT_PTR CALLBACK PaneFilterDialogProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    auto* state = reinterpret_cast<PaneFilterDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnPaneFilterDialogInit(dlg, reinterpret_cast<PaneFilterDialogState*>(lParam));
        case WM_CTLCOLORDLG: return OnPaneFilterDialogCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnPaneFilterDialogCtlColorStatic(state, reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORLISTBOX: return OnPaneFilterDialogCtlColorEdit(state, reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->common.theme, wParam != FALSE);
            }
            return FALSE;
        case WM_SETCURSOR:
            if (state)
            {
                const HWND hover = reinterpret_cast<HWND>(wParam);
                const HWND hint  = GetDlgItem(dlg, IDC_PANE_FILTER_HINT);
                if (hover && hint && hover == hint)
                {
                    SetCursor(LoadCursorW(nullptr, IDC_HAND));
                    return TRUE;
                }
            }
            break;
        case WM_SIZE:
            if (state)
            {
                UpdatePaneFilterDialogLayout(dlg, state, false);
                return TRUE;
            }
            break;
        case WM_DPICHANGED:
            if (state)
            {
                const RECT* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    const int w = std::max(1l, suggested->right - suggested->left);
                    const int h = std::max(1l, suggested->bottom - suggested->top);
                    SetWindowPos(dlg, nullptr, suggested->left, suggested->top, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
                }
                UpdatePaneFilterDialogLayout(dlg, state, true);
                return TRUE;
            }
            break;
        case WM_DRAWITEM:
        {
            if (! state || state->common.theme.highContrast || state->common.theme.systemHighContrast)
            {
                break;
            }

            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (! dis || dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            if (dis->CtlID == IDC_PANE_FILTER_USE_TOGGLE)
            {
                const bool toggledOn   = dis->hwndItem && SendMessageW(dis->hwndItem, BM_GETCHECK, 0, 0) == BST_CHECKED;
                const COLORREF surface = ThemedControls::GetControlSurfaceColor(state->common.theme);
                const HFONT boldFont   = state->common.boldFont ? state->common.boldFont.get() : nullptr;
                ThemedControls::DrawThemedSwitchToggle(*dis, state->common.theme, surface, boldFont, state->toggleOnLabel, state->toggleOffLabel, toggledOn);
                return TRUE;
            }

            ThemedControls::DrawThemedPushButton(*dis, state->common.theme);
            return TRUE;
        }
        case WM_COMMAND: return OnPaneFilterDialogCommand(dlg, state, LOWORD(wParam), HIWORD(wParam));
    }
    return FALSE;
}

std::optional<FolderView::NameFilterState> PromptForPaneFilter(HWND ownerWindow,
                                                               const std::vector<std::wstring>& history,
                                                               const AppTheme& theme,
                                                               const FolderView::NameFilterState& initial)
{
    PaneFilterDialogState state{};
    state.common.centerOnWindow = ownerWindow;
    state.common.history        = &history;
    state.common.theme          = theme;
    state.common.helpExpanded   = false;
    state.captionText           = LoadStringResource(nullptr, IDS_CAPTION_PANE_FILTER);
    state.useLabelText          = LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER_USE_FILTER);
    state.labelText             = LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER);
    state.initial               = initial;

#pragma warning(push)
// pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
    const INT_PTR result =
        DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PANE_FILTER), ownerWindow, PaneFilterDialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)

    if (result == IDOK)
    {
        return state.result;
    }

    return std::nullopt;
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
} // namespace

struct ChangeCaseTaskPayload final
{
    FolderWindow::InformationalTaskUpdate update{};
};

struct ChangeCaseCompletedPayload final
{
    FolderWindow::Pane pane = FolderWindow::Pane::Left;
    HRESULT hr              = S_OK;
};

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
            DirectoryInfoCache::GetInstance().ClearForFileSystem(previous.get());
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
    const auto createFactory   = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(keepAlive.get(), "RedSalamanderCreate"));
    const auto createFactoryEx = reinterpret_cast<CreateFactoryExFunc>(GetProcAddress(keepAlive.get(), "RedSalamanderCreateEx"));
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
    HRESULT createHr = E_FAIL;
    if (entry->factoryPluginId.empty())
    {
        createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), fileSystemInstance.put_void());
    }
    else if (createFactoryEx)
    {
        createHr = createFactoryEx(__uuidof(IFileSystem), &options, GetHostServices(), entry->factoryPluginId.c_str(), fileSystemInstance.put_void());
    }
    else
    {
        Debug::Error(L"FolderWindow: Missing export RedSalamanderCreateEx in '{}' for multi-plugin DLL", entry->path.c_str());
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }
    if (FAILED(createHr) || ! fileSystemInstance)
    {
        Debug::Error(L"FolderWindow: RedSalamanderCreate failed for '{}' (hr=0x{:08X})", entry->path.c_str(), static_cast<unsigned long>(createHr));
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

    if (previous && previous.get() != state.fileSystem.get() && (! other.fileSystem || other.fileSystem.get() != previous.get()))
    {
        DirectoryInfoCache::GetInstance().ClearForFileSystem(previous.get());
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

HRESULT FolderWindow::SetFileSystemPluginForPane(Pane pane, std::wstring_view pluginId) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! state.pluginId.empty() && EqualsNoCase(state.pluginId, pluginId))
    {
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

    if (previous && previous.get() != state.fileSystem.get() && (! other.fileSystem || other.fileSystem.get() != previous.get()))
    {
        DirectoryInfoCache::GetInstance().ClearForFileSystem(previous.get());
    }

    previous.reset(); // release before module unload
    return S_OK;
}

HRESULT FolderWindow::ExecuteInActivePane(const std::filesystem::path& folderPath,
                                          std::wstring_view focusItemDisplayName,
                                          unsigned int folderViewCommandId,
                                          bool activateWindow) noexcept
{
    if (folderPath.empty())
    {
        return E_INVALIDARG;
    }

    const Pane pane  = _activePane;
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (activateWindow)
    {
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
    }

    if (state.hFolderView && IsWindow(state.hFolderView.get()))
    {
        SetFocus(state.hFolderView.get());
    }

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
        bool ready = true;
        if (! focusItemDisplayName.empty())
        {
            ready = state.folderView.PrepareForExternalCommand(focusItemDisplayName);
        }

        if (ready && folderViewCommandId != 0u && state.hFolderView)
        {
            PostMessageW(state.hFolderView.get(), WM_COMMAND, MAKEWPARAM(folderViewCommandId, 0), 0);
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

    SetFolderPath(pane, folderPath);
    return S_OK;
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

        static_cast<void>(ShowConnectionManagerWindow(_hWnd.get(), L"RedSalamander", *_settings, _theme, filterPluginId, static_cast<uint8_t>(pane)));
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
                 EqualsNoCase(pluginShortId, L"imap") || EqualsNoCase(pluginShortId, L"s3") || EqualsNoCase(pluginShortId, L"s3table"));

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
                // Treat `ftp:` and `ftp://@conn` as explicit Connection Manager entry points.
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
                    DirectoryInfoCache::GetInstance().ClearForFileSystem(state.fileSystem.get());
                    state.instanceContext = instanceContext;
                    static_cast<void>(initializer->Initialize(state.instanceContext.c_str(), nullptr));
                }
                else if (instanceContextSpecified && instanceContext.empty() && ! state.instanceContext.empty())
                {
                    DirectoryInfoCache::GetInstance().ClearForFileSystem(state.fileSystem.get());
                    state.instanceContext.clear();
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

    const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, pluginPath);

    {
        Debug::Perf::Scope updatePerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.UpdateViews" : L"FolderWindow.SetFolderPath.Right.UpdateViews");
        updatePerf.SetDetail(displayPath.native());

        state.updatingPath = true;
        state.currentPath  = displayPath;

        if (state.hNavigationView)
        {
            Debug::Perf::Scope navPerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.UpdateViews.NavigationView.SetPath"
                                                          : L"FolderWindow.SetFolderPath.Right.UpdateViews.NavigationView.SetPath");
            navPerf.SetDetail(displayPath.native());
            state.navigationView.SetPath(displayPath);
        }

        if (state.hFolderView)
        {
            Debug::Perf::Scope viewPerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.UpdateViews.FolderView.SetFolderPath"
                                                           : L"FolderWindow.SetFolderPath.Right.UpdateViews.FolderView.SetFolderPath");
            viewPerf.SetDetail(pluginPath.native());

            const Common::Settings::FoldersSettings* folders = (_settings && _settings->folders.has_value()) ? &_settings->folders.value() : nullptr;
            const FolderView::NameFilterState filter         = GetFolderHistoryFilterState(folders, displayPath);
            state.folderView.SetNameFilterState(filter, false /* refresh */);
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

void FolderWindow::CommandCreateDirectory(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.fileSystem)
    {
        return;
    }

    HWND ownerWindow = GetOwnerWindowOrSelf(_hWnd.get());
    std::wstring pluginName;
    if (ownerWindow)
    {
        FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
        const auto& plugins                    = pluginManager.GetPlugins();
        pluginName                             = TryGetFileSystemPluginDisplayName(plugins, state.pluginId, state.pluginShortId);
    }

    const auto folder = state.folderView.GetFolderPath();
    if (! folder)
    {
        return;
    }

    const std::filesystem::path base = folder.value();

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryOperations), dirOps.put_void());

    const bool canUseWin32 = IsFilePluginShortId(state.pluginShortId) && LooksLikeWindowsAbsolutePath(base.wstring());
    if (! dirOps && ! canUseWin32)
    {
        std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message;
        if (! pluginName.empty())
        {
            message = FormatStringResource(nullptr, IDS_FMT_PANE_CREATE_DIR_UNSUPPORTED_PLUGIN, pluginName);
        }
        if (message.empty())
        {
            message = LoadStringResource(nullptr, IDS_MSG_PANE_CREATE_DIR_UNSUPPORTED);
        }

        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
        return;
    }

    std::wstring defaultName = LoadStringResource(nullptr, IDS_NEW_FOLDER_DEFAULT_NAME);
    if (defaultName.empty())
    {
        return;
    }

    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, base);
    const auto folderName                   = PromptForCreateDirectoryName(ownerWindow, displayPath.wstring(), defaultName, _theme);
    if (! folderName.has_value())
    {
        return;
    }

    const std::wstring requestedName = folderName.value();
    const bool autoSuffix            = requestedName == defaultName;

    const int maxAttempts = autoSuffix ? 1000 : 1;
    for (int attempt = 0; attempt < maxAttempts; ++attempt)
    {
        std::wstring candidateName = requestedName;
        if (autoSuffix && attempt > 0)
        {
            candidateName = std::format(L"{} ({})", requestedName, attempt + 1);
        }

        const std::filesystem::path newFolderPath = base / std::filesystem::path(candidateName);
        if (newFolderPath.empty())
        {
            continue;
        }

        HRESULT hr = S_OK;
        if (dirOps)
        {
            hr = dirOps->CreateDirectory(newFolderPath.c_str());
        }
        else
        {
            if (::CreateDirectoryW(newFolderPath.c_str(), nullptr) == 0)
            {
                const DWORD error = GetLastError();
                hr                = HRESULT_FROM_WIN32(error);
            }
        }

        if (SUCCEEDED(hr))
        {
            const std::wstring focusName = newFolderPath.filename().wstring();
            if (! focusName.empty())
            {
                state.folderView.RememberFocusedItemForFolder(base, focusName);
            }

            DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
            state.folderView.ForceRefresh();

            const Pane otherPane   = pane == Pane::Left ? Pane::Right : Pane::Left;
            PaneState& otherState  = otherPane == Pane::Left ? _leftPane : _rightPane;
            const auto otherFolder = otherState.folderView.GetFolderPath();
            if (otherState.fileSystem && otherFolder.has_value() && OrdinalString::EqualsNoCasePath(otherFolder.value(), base) &&
                EqualsNoCase(otherState.pluginId, state.pluginId) && EqualsNoCase(otherState.instanceContext, state.instanceContext) &&
                ! cache.IsFolderWatched(otherState.fileSystem.get(), base))
            {
                otherState.folderView.ForceRefresh();
            }
            return;
        }

        if (hr == E_NOTIMPL)
        {
            std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
            std::wstring message;
            if (! pluginName.empty())
            {
                message = FormatStringResource(nullptr, IDS_FMT_PANE_CREATE_DIR_UNSUPPORTED_PLUGIN, pluginName);
            }
            if (message.empty())
            {
                message = LoadStringResource(nullptr, IDS_MSG_PANE_CREATE_DIR_UNSUPPORTED);
            }

            state.folderView.ShowAlertOverlay(
                FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
            return;
        }

        constexpr HRESULT alreadyExistsHr = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        constexpr HRESULT fileExistsHr    = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        if (autoSuffix && (hr == alreadyExistsHr || hr == fileExistsHr))
        {
            continue;
        }

        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message = FormatStringResource(nullptr, IDS_FMT_PANE_CREATE_DIR_FAILED, newFolderPath.wstring(), static_cast<unsigned long>(hr));
        state.folderView.ShowAlertOverlay(
            FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message), hr);
        return;
    }
}

void FolderWindow::CommandRefresh(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ForceRefresh();
}

#ifdef _DEBUG
uint64_t FolderWindow::DebugGetForceRefreshCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetForceRefreshCount();
}

std::wstring_view FolderWindow::DebugGetFocusedItemDisplayName(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetFocusedDisplayName();
}

bool FolderWindow::DebugHasItemDisplayName(Pane pane, std::wstring_view displayName) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugHasItemDisplayName(displayName);
}

bool FolderWindow::DebugIsItemSelected(Pane pane, std::wstring_view displayName) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugIsItemSelectedByDisplayName(displayName);
}

size_t FolderWindow::DebugGetSelectedCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetSelectedItemCount();
}

bool FolderWindow::DebugIsEmptyFolderStateActive(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugIsEmptyFolderStateActive();
}

std::wstring_view FolderWindow::DebugGetEmptyFolderFunMessage(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetEmptyFolderFunMessage();
}

bool FolderWindow::DebugFocusItemByDisplayName(Pane pane, std::wstring_view displayName) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.PrepareForExternalCommand(displayName);
}

FolderView::NameFilterState FolderWindow::DebugGetNameFilterState(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetNameFilterState();
}

bool FolderWindow::DebugIsNameFilterActive(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.IsNameFilterActive();
}

FolderView::FilterWatermarkVisualMode FolderWindow::DebugGetFilterWatermarkVisualMode(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetFilterWatermarkVisualMode();
}
#endif

void FolderWindow::CommandCalculateDirectorySizes(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::optional<std::filesystem::path> folderPath = state.folderView.GetFolderPath();
    if (! folderPath.has_value() || folderPath.value().empty())
    {
        return;
    }

    if (! TryViewSpaceWithViewer(pane, folderPath.value()))
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_PANE_CALCULATE_DIRECTORY_SIZES_FAILED);

        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
    }
}

void FolderWindow::CommandSelectionSelectDialog(Pane pane)
{
    SetActivePane(pane);

    std::vector<std::wstring> history;
    if (_settings && _settings->selectionMasks.has_value())
    {
        history = _settings->selectionMasks->selectHistory;
    }
    MaskSyntax::NormalizeWildcardMaskHistory(history, MaskSyntax::kWildcardMaskHistoryMaxItems);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::optional<std::wstring> maskTextOpt =
        PromptForSelectionMask(ownerWindow, history, _theme, IDS_CAPTION_SELECTION_MASK_SELECT, IDS_LABEL_SELECTION_MASK_SELECT);
    if (! maskTextOpt.has_value())
    {
        return;
    }

    const std::wstring& maskText = maskTextOpt.value();

    if (_settings)
    {
        Common::Settings::SelectionMasksSettings& masks =
            _settings->selectionMasks.has_value() ? _settings->selectionMasks.value() : _settings->selectionMasks.emplace();

        MaskSyntax::AddToWildcardMaskHistory(masks.selectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems, maskText);
        MaskSyntax::NormalizeWildcardMaskHistory(masks.selectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems);
    }

    MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(maskText);

    SetPaneSelectionByDisplayNamePredicate(pane, [mask = std::move(mask)](std::wstring_view displayName) noexcept {
        return MaskSyntax::MatchesWildcardMask(displayName, mask);
    }, false /* clearExistingSelection */);
}

void FolderWindow::CommandSelectionUnselectDialog(Pane pane)
{
    SetActivePane(pane);

    std::vector<std::wstring> history;
    if (_settings && _settings->selectionMasks.has_value())
    {
        history = _settings->selectionMasks->unselectHistory;
    }
    MaskSyntax::NormalizeWildcardMaskHistory(history, MaskSyntax::kWildcardMaskHistoryMaxItems);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::optional<std::wstring> maskTextOpt =
        PromptForSelectionMask(ownerWindow, history, _theme, IDS_CAPTION_SELECTION_MASK_UNSELECT, IDS_LABEL_SELECTION_MASK_UNSELECT);
    if (! maskTextOpt.has_value())
    {
        return;
    }

    const std::wstring& maskText = maskTextOpt.value();

    if (_settings)
    {
        Common::Settings::SelectionMasksSettings& masks =
            _settings->selectionMasks.has_value() ? _settings->selectionMasks.value() : _settings->selectionMasks.emplace();

        MaskSyntax::AddToWildcardMaskHistory(masks.unselectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems, maskText);
        MaskSyntax::NormalizeWildcardMaskHistory(masks.unselectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems);
    }

    MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(maskText);

    ClearPaneSelectionByDisplayNamePredicate(
        pane, [mask = std::move(mask)](std::wstring_view displayName) noexcept { return MaskSyntax::MatchesWildcardMask(displayName, mask); });
}

void FolderWindow::CommandSelectionInvert(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.InvertSelection();
}

bool FolderWindow::HasSavedSelection() const noexcept
{
    return _savedSelection.has_value() && ! _savedSelection->displayNames.empty();
}

void FolderWindow::CommandSelectionSave(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::optional<std::filesystem::path> folderOpt = state.folderView.GetFolderPath();
    if (! folderOpt.has_value() || folderOpt.value().empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    std::vector<std::wstring> names = state.folderView.GetSelectedOrFocusedDisplayNames();
    if (names.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    SavedSelection saved{};
    saved.sourcePluginId        = std::wstring(state.folderView.GetFileSystemPluginId());
    saved.sourceInstanceContext = std::wstring(state.folderView.GetFileSystemInstanceContext());
    saved.sourceFolder          = folderOpt.value();
    saved.displayNames          = std::move(names);
    _savedSelection             = std::move(saved);

    std::wstring clipboardText;
    {
        const std::wstring folderText = folderOpt.value().native();
        size_t reserveChars           = folderText.size() + 2u;
        for (const auto& name : _savedSelection->displayNames)
        {
            reserveChars += name.size() + 2u;
        }

        clipboardText.reserve(reserveChars);
        clipboardText.append(folderText);
        for (const auto& name : _savedSelection->displayNames)
        {
            clipboardText.append(L"\r\n");
            clipboardText.append(name);
        }
        clipboardText.append(L"\r\n");
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    if (! SetClipboardUnicodeText(ownerWindow, clipboardText))
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_SELECTION_SAVE_CLIPBOARD_FAILED);
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
    }
}

void FolderWindow::CommandCopyPathAndFileNameAsText(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    Debug::Info(L"event=share_copy_path_and_file_name_clicked count={}", static_cast<uint64_t>(paths.size()));

    std::wstring clipboardText;
    {
        size_t reserveChars = 2u;
        for (const auto& path : paths)
        {
            reserveChars += path.native().size() + 2u;
        }
        clipboardText.reserve(reserveChars);

        bool first = true;
        for (const auto& path : paths)
        {
            const std::wstring_view native = path.native();
            if (native.empty())
            {
                continue;
            }

            if (! first)
            {
                clipboardText.append(L"\r\n");
            }
            clipboardText.append(native);
            first = false;
        }
        clipboardText.append(L"\r\n");
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    if (! SetClipboardUnicodeText(ownerWindow, clipboardText))
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_SELECTION_SAVE_CLIPBOARD_FAILED);
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
        return;
    }

    Debug::Info(L"event=share_copy_path_and_file_name_succeeded count={}", static_cast<uint64_t>(paths.size()));

    std::wstring title = LoadStringResource(nullptr, IDS_CMD_COPY_PATH_AND_FILE_NAME);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_INFORMATION);
    }

    const unsigned long long count = static_cast<unsigned long long>(paths.size());
    const std::wstring_view suffix = count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s");
    std::wstring message           = FormatStringResource(nullptr, IDS_FMT_COPY_PATH_AND_FILE_NAME_COPIED, count, suffix);
    if (message.empty())
    {
        message = std::format(L"Copied {} path{} to clipboard.", count, suffix);
    }

    ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Information, std::move(title), std::move(message), S_OK, false, false);
}

void FolderWindow::CommandSelectionRestore(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! HasSavedSelection())
    {
        MessageBeep(MB_ICONWARNING);
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_SELECTION_RESTORE_NO_SAVED);
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
        return;
    }

    const SavedSelection& saved = _savedSelection.value();

    std::unordered_set<std::wstring_view> remaining;
    remaining.reserve(saved.displayNames.size());
    for (const auto& name : saved.displayNames)
    {
        if (! name.empty())
        {
            remaining.emplace(name);
        }
    }

    SetPaneSelectionByDisplayNamePredicate(pane,
                                           [&](std::wstring_view displayName) noexcept
    {
        const auto it = remaining.find(displayName);
        if (it == remaining.end())
        {
            return false;
        }
        remaining.erase(it);
        return true;
    },
                                           true /* clearExistingSelection */);

    if (! remaining.empty())
    {
        struct WStringViewNoCaseLess final
        {
            bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
            {
                return OrdinalString::Compare(left, right, true) < 0;
            }
        };

        std::map<std::wstring_view, std::vector<std::wstring_view>, WStringViewNoCaseLess> remainingNoCase;
        for (const auto& name : remaining)
        {
            remainingNoCase[name].push_back(name);
        }

        SetPaneSelectionByDisplayNamePredicate(pane,
                                               [&](std::wstring_view displayName) noexcept
        {
            const auto it = remainingNoCase.find(displayName);
            if (it == remainingNoCase.end() || it->second.empty())
            {
                return false;
            }

            const std::wstring_view matched = it->second.back();
            it->second.pop_back();
            remaining.erase(matched);
            if (it->second.empty())
            {
                remainingNoCase.erase(it);
            }
            return true;
        },
                                               false /* clearExistingSelection */);
    }

    if (remaining.empty())
    {
        return;
    }

    std::wstring missingLines;
    for (const auto& name : saved.displayNames)
    {
        if (name.empty())
        {
            continue;
        }

        if (remaining.find(std::wstring_view(name)) == remaining.end())
        {
            continue;
        }

        missingLines.append(L"- ");
        missingLines.append(name);
        missingLines.append(L"\r\n");
    }

    std::wstring filterNote;
    if (state.folderView.IsNameFilterActive())
    {
        const std::wstring noteText = LoadStringResource(nullptr, IDS_MSG_SELECTION_RESTORE_FILTER_NOTE);
        if (! noteText.empty())
        {
            filterNote.reserve(4u + noteText.size());
            filterNote.append(L"\r\n\r\n");
            filterNote.append(noteText);
        }
    }

    std::wstring title = LoadStringResource(nullptr, IDS_CMD_SELECTION_RESTORE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_INFORMATION);
    }

    std::wstring message = FormatStringResource(nullptr, IDS_FMT_SELECTION_RESTORE_INCOMPLETE, missingLines, filterNote);
    ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Information, std::move(title), std::move(message), S_OK, true, false);
}

void FolderWindow::CommandSelectionSelectSameExtension(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SelectSameExtension();
}

void FolderWindow::CommandSelectionUnselectSameExtension(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.UnselectSameExtension();
}

void FolderWindow::CommandSelectionHideSelectedNames(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.HideSelectedNames();
}

void FolderWindow::CommandSelectionHideUnselectedNames(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.HideUnselectedNames();
}

void FolderWindow::CommandSelectionShowHiddenNames(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ShowHiddenNames();
}

bool FolderWindow::CanShowHiddenNames(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.HasHiddenNames();
}

void FolderWindow::CommandSelectionGoToPreviousSelectedName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    static_cast<void>(state.folderView.GoToPreviousSelectedName());
}

void FolderWindow::CommandSelectionGoToNextSelectedName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    static_cast<void>(state.folderView.GoToNextSelectedName());
}

void FolderWindow::CommandChangeCase(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! state.fileSystem)
    {
        return;
    }

    if (state.changeCaseThread.joinable())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    ChangeCaseDialogState dialogState{};
    dialogState.theme        = _theme;
    dialogState.allowSubdirs = true;

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

#pragma warning(push)
// C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
    const INT_PTR dlgResult = DialogBoxParamW(
        GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PANE_CHANGE_CASE), ownerWindow, ChangeCaseDialogProc, reinterpret_cast<LPARAM>(&dialogState));
#pragma warning(pop)

    if (dlgResult != IDOK || ! dialogState.accepted)
    {
        return;
    }

    const std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    const HWND ownerHwnd                 = _hWnd.get();
    wil::com_ptr<IFileSystem> fileSystem = state.fileSystem;
    const ChangeCase::Options options    = dialogState.options;
    const std::wstring title             = LoadStringResource(nullptr, IDS_CMD_CHANGE_CASE);

    state.changeCaseThread = std::jthread([ownerHwnd, pane, fileSystem, paths, options, title](std::stop_token stopToken) noexcept
    {
        const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(coinitHr))
        {
            Debug::Error(L"ChangeCase task: CoInitializeEx(COINIT_MULTITHREADED) failed: 0x{:08X}", coinitHr);
            FAIL_FAST_IF_FAILED(coinitHr);
        }
        [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

        struct ProgressState final
        {
            HWND hwnd               = nullptr;
            FolderWindow::Pane pane = FolderWindow::Pane::Left;
            std::wstring title;
            ULONGLONG startTick      = 0;
            ULONGLONG lastPostedTick = 0;
            uint64_t infoTaskId      = 0;
            ChangeCase::ProgressUpdate last{};

            void PostTaskUpdate(bool finished, HRESULT hr) noexcept
            {
                if (! hwnd || IsWindow(hwnd) == FALSE || infoTaskId == 0)
                {
                    return;
                }

                FolderWindow::InformationalTaskUpdate info{};
                info.kind                       = FolderWindow::InformationalTaskUpdate::Kind::ChangeCase;
                info.taskId                     = infoTaskId;
                info.title                      = title;
                info.changeCaseCurrentPath      = last.currentPath;
                info.changeCaseScannedFolders   = last.scannedFolders;
                info.changeCaseScannedEntries   = last.scannedEntries;
                info.changeCasePlannedRenames   = last.plannedRenames;
                info.changeCaseCompletedRenames = last.completedRenames;
                info.changeCaseEnumerating      = ! finished && last.phase == ChangeCase::ProgressUpdate::Phase::Enumerating;
                info.changeCaseRenaming         = ! finished && last.phase == ChangeCase::ProgressUpdate::Phase::Renaming;
                info.finished                   = finished;
                info.resultHr                   = hr;

                auto payload    = std::make_unique<ChangeCaseTaskPayload>();
                payload->update = std::move(info);
                static_cast<void>(PostMessagePayload(hwnd, WndMsg::kChangeCaseTaskUpdate, 0, std::move(payload)));
            }

            void EnsureTaskVisibleAfterThreshold() noexcept
            {
                if (! hwnd || IsWindow(hwnd) == FALSE || infoTaskId != 0)
                {
                    return;
                }

                const ULONGLONG nowTick = GetTickCount64();
                if (startTick == 0 || nowTick < startTick || (nowTick - startTick) < 700ull)
                {
                    return;
                }

                FolderWindow::InformationalTaskUpdate info{};
                info.kind                       = FolderWindow::InformationalTaskUpdate::Kind::ChangeCase;
                info.title                      = title;
                info.changeCaseCurrentPath      = last.currentPath;
                info.changeCaseScannedFolders   = last.scannedFolders;
                info.changeCaseScannedEntries   = last.scannedEntries;
                info.changeCasePlannedRenames   = last.plannedRenames;
                info.changeCaseCompletedRenames = last.completedRenames;
                info.changeCaseEnumerating      = last.phase == ChangeCase::ProgressUpdate::Phase::Enumerating;
                info.changeCaseRenaming         = last.phase == ChangeCase::ProgressUpdate::Phase::Renaming;

                auto payload    = std::make_unique<ChangeCaseTaskPayload>();
                payload->update = std::move(info);

                ChangeCaseTaskPayload* raw = payload.release();
                DWORD_PTR result           = 0;
                const LRESULT sendOk =
                    SendMessageTimeoutW(hwnd, WndMsg::kChangeCaseTaskUpdate, 0, reinterpret_cast<LPARAM>(raw), SMTO_ABORTIFHUNG | SMTO_BLOCK, 1000, &result);
                if (sendOk == 0)
                {
                    delete raw;
                    return;
                }

                infoTaskId     = static_cast<uint64_t>(result);
                lastPostedTick = nowTick;
            }
        };

        ProgressState progressState{};
        progressState.hwnd      = ownerHwnd;
        progressState.pane      = pane;
        progressState.title     = title;
        progressState.startTick = GetTickCount64();

        const auto onProgress = [](const ChangeCase::ProgressUpdate& update, void* cookie) noexcept
        {
            auto* state = static_cast<ProgressState*>(cookie);
            if (! state)
            {
                return;
            }

            state->last = update;
            state->EnsureTaskVisibleAfterThreshold();

            if (state->infoTaskId != 0)
            {
                const ULONGLONG nowTick = GetTickCount64();
                if (state->lastPostedTick != 0 && nowTick >= state->lastPostedTick && (nowTick - state->lastPostedTick) < 100ull)
                {
                    return;
                }

                state->lastPostedTick = nowTick;
                state->PostTaskUpdate(false, S_OK);
            }
        };

        const HRESULT operationHr = ChangeCase::ApplyToPaths(*fileSystem, paths, options, stopToken, onProgress, &progressState);

        progressState.EnsureTaskVisibleAfterThreshold();
        progressState.PostTaskUpdate(true, operationHr);

        if (ownerHwnd && IsWindow(ownerHwnd) != FALSE)
        {
            auto completed  = std::make_unique<ChangeCaseCompletedPayload>();
            completed->pane = pane;
            completed->hr   = operationHr;
            static_cast<void>(PostMessagePayload(ownerHwnd, WndMsg::kChangeCaseCompleted, 0, std::move(completed)));
        }
    });
}

void FolderWindow::CommandChangeDirectory(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.OpenChangeDirectoryFromCommand();
}

void FolderWindow::CommandFocusAddressBar(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.FocusAddressBar();
}

void FolderWindow::CommandOpenDriveMenu(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.OpenDriveMenuFromCommand();
}

void FolderWindow::CommandShowFolderHistory(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.OpenHistoryDropdownFromKeyboard();
}

void FolderWindow::CommandFilter(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    std::vector<std::wstring> history;
    if (_settings && _settings->selectionMasks.has_value())
    {
        history = _settings->selectionMasks->filterHistory;
    }
    MaskSyntax::NormalizeWildcardMaskHistory(history, MaskSyntax::kWildcardMaskHistoryMaxItems);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const FolderView::NameFilterState initial                  = state.folderView.GetNameFilterState();
    const std::optional<FolderView::NameFilterState> resultOpt = PromptForPaneFilter(ownerWindow, history, _theme, initial);
    if (! resultOpt.has_value())
    {
        return;
    }

    const FolderView::NameFilterState result = resultOpt.value();

    if (_settings && ! result.text.empty())
    {
        Common::Settings::SelectionMasksSettings& masks =
            _settings->selectionMasks.has_value() ? _settings->selectionMasks.value() : _settings->selectionMasks.emplace();

        MaskSyntax::AddToWildcardMaskHistory(masks.filterHistory, MaskSyntax::kWildcardMaskHistoryMaxItems, result.text);
        MaskSyntax::NormalizeWildcardMaskHistory(masks.filterHistory, MaskSyntax::kWildcardMaskHistoryMaxItems);
    }

    state.folderView.SetNameFilterState(result);

    if (_settings && state.currentPath.has_value() && ! state.currentPath.value().empty())
    {
        Common::Settings::FoldersSettings& folders = _settings->folders.has_value() ? _settings->folders.value() : _settings->folders.emplace();
        SetFolderHistoryFilterState(folders, state.currentPath.value(), result);
        PruneFolderHistoryFilters(folders, _folderHistory, static_cast<size_t>(_folderHistoryMax));
    }
}

void FolderWindow::CommandGoRootDirectory(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (IsFilePluginShortId(state.pluginShortId))
    {
        const std::optional<std::filesystem::path> pluginPathOpt = state.folderView.GetFolderPath();
        if (! pluginPathOpt.has_value() || pluginPathOpt.value().empty())
        {
            return;
        }

        std::filesystem::path root = pluginPathOpt.value().root_path();
        if (root.empty())
        {
            root = GetDefaultFileSystemRoot();
        }

        SetFolderPath(pane, root);
        return;
    }

    std::filesystem::path rootPluginPath(L"/");
    {
        const std::optional<std::filesystem::path> pluginPathOpt = state.folderView.GetFolderPath();
        if (pluginPathOpt.has_value() && ! pluginPathOpt.value().empty())
        {
            std::wstring normalized = NavigationLocation::NormalizePluginPathText(pluginPathOpt.value().wstring(),
                                                                                  NavigationLocation::EmptyPathPolicy::Root,
                                                                                  NavigationLocation::LeadingSlashPolicy::Ensure,
                                                                                  NavigationLocation::TrailingSlashPolicy::Trim);

            constexpr std::wstring_view kConnPrefix = L"/@conn:";
            if (StartsWithNoCase(normalized, kConnPrefix))
            {
                const size_t nameStart = kConnPrefix.size();
                if (nameStart < normalized.size())
                {
                    const size_t nextSlash = normalized.find(L'/', nameStart);
                    const size_t end       = nextSlash == std::wstring::npos ? normalized.size() : nextSlash;
                    if (end > 0u)
                    {
                        std::wstring rootText = normalized.substr(0, end);
                        if (! rootText.empty() && rootText.back() != L'/')
                        {
                            rootText.push_back(L'/');
                        }
                        rootPluginPath = std::filesystem::path(std::move(rootText));
                    }
                }
            }
        }
    }

    const std::filesystem::path rootDisplayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, rootPluginPath);
    SetFolderPath(pane, rootDisplayPath);
}

void FolderWindow::CommandSetPathFromOtherPane(Pane pane)
{
    SetActivePane(pane);

    const Pane otherPane                                 = pane == Pane::Left ? Pane::Right : Pane::Left;
    const std::optional<std::filesystem::path> otherPath = GetCurrentPath(otherPane);
    if (! otherPath.has_value() || otherPath.value().empty())
    {
        return;
    }

    SetFolderPath(pane, otherPath.value());
}

bool FolderWindow::CanHistoryBack(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return ! state.navigationHistory.empty() && state.navigationHistoryIndex > 0 && state.navigationHistoryIndex < state.navigationHistory.size();
}

bool FolderWindow::CanHistoryForward(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return ! state.navigationHistory.empty() && (state.navigationHistoryIndex + 1) < state.navigationHistory.size();
}

void FolderWindow::CommandHistoryBack(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! CanHistoryBack(pane))
    {
        return;
    }

    const std::optional<std::filesystem::path> before = state.currentPath;
    const size_t previousIndex                        = state.navigationHistoryIndex;
    const size_t targetIndex                          = previousIndex - 1;
    const std::filesystem::path target                = state.navigationHistory[targetIndex];

    state.navigationHistoryIndex = targetIndex;

    struct SuspendNavigationHistoryRecord final
    {
        PaneState& state;
        bool previous = false;

        explicit SuspendNavigationHistoryRecord(PaneState& s) noexcept : state(s), previous(s.navigationHistorySuspendRecord)
        {
            state.navigationHistorySuspendRecord = true;
        }

        SuspendNavigationHistoryRecord(const SuspendNavigationHistoryRecord&)            = delete;
        SuspendNavigationHistoryRecord& operator=(const SuspendNavigationHistoryRecord&) = delete;
        SuspendNavigationHistoryRecord(SuspendNavigationHistoryRecord&&)                 = delete;
        SuspendNavigationHistoryRecord& operator=(SuspendNavigationHistoryRecord&&)      = delete;

        ~SuspendNavigationHistoryRecord()
        {
            state.navigationHistorySuspendRecord = previous;
        }
    };

    {
        SuspendNavigationHistoryRecord guard(state);
        SetFolderPath(pane, target);
    }

    const bool changed =
        state.currentPath.has_value() && (! before.has_value() || ! OrdinalString::EqualsNoCasePath(before.value(), state.currentPath.value()));
    if (! changed)
    {
        state.navigationHistoryIndex = previousIndex;
        return;
    }

    if (state.navigationHistoryIndex < state.navigationHistory.size() && state.currentPath.has_value() && ! state.currentPath.value().empty())
    {
        state.navigationHistory[state.navigationHistoryIndex] = state.currentPath.value();
    }
}

void FolderWindow::CommandHistoryForward(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! CanHistoryForward(pane))
    {
        return;
    }

    const std::optional<std::filesystem::path> before = state.currentPath;
    const size_t previousIndex                        = state.navigationHistoryIndex;
    const size_t targetIndex                          = previousIndex + 1;
    const std::filesystem::path target                = state.navigationHistory[targetIndex];

    state.navigationHistoryIndex = targetIndex;

    struct SuspendNavigationHistoryRecord final
    {
        PaneState& state;
        bool previous = false;

        explicit SuspendNavigationHistoryRecord(PaneState& s) noexcept : state(s), previous(s.navigationHistorySuspendRecord)
        {
            state.navigationHistorySuspendRecord = true;
        }

        SuspendNavigationHistoryRecord(const SuspendNavigationHistoryRecord&)            = delete;
        SuspendNavigationHistoryRecord& operator=(const SuspendNavigationHistoryRecord&) = delete;
        SuspendNavigationHistoryRecord(SuspendNavigationHistoryRecord&&)                 = delete;
        SuspendNavigationHistoryRecord& operator=(SuspendNavigationHistoryRecord&&)      = delete;

        ~SuspendNavigationHistoryRecord()
        {
            state.navigationHistorySuspendRecord = previous;
        }
    };

    {
        SuspendNavigationHistoryRecord guard(state);
        SetFolderPath(pane, target);
    }

    const bool changed =
        state.currentPath.has_value() && (! before.has_value() || ! OrdinalString::EqualsNoCasePath(before.value(), state.currentPath.value()));
    if (! changed)
    {
        state.navigationHistoryIndex = previousIndex;
        return;
    }

    if (state.navigationHistoryIndex < state.navigationHistory.size() && state.currentPath.has_value() && ! state.currentPath.value().empty())
    {
        state.navigationHistory[state.navigationHistoryIndex] = state.currentPath.value();
    }
}

void FolderWindow::PrepareForNetworkDriveDisconnect(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.CancelPendingEnumeration();
    if (state.fileSystem)
    {
        DirectoryInfoCache::GetInstance().ClearForFileSystem(state.fileSystem.get());
    }
}

void FolderWindow::CommandOpenCommandShell(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    std::filesystem::path workingDir;
    if (IsFilePluginShortId(state.pluginShortId))
    {
        const std::optional<std::filesystem::path> folderPath = state.folderView.GetFolderPath();
        if (folderPath.has_value() && LooksLikeWindowsAbsolutePath(folderPath.value().wstring()))
        {
            workingDir = folderPath.value();
        }
    }
    else if (! state.instanceContext.empty() && LooksLikeWindowsAbsolutePath(state.instanceContext))
    {
        std::filesystem::path contextPath(state.instanceContext);
        DWORD attrs = GetFileAttributesW(contextPath.c_str());
        if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            workingDir = std::move(contextPath);
        }
        else
        {
            workingDir = contextPath.parent_path();
        }
    }

    if (workingDir.empty())
    {
        workingDir = GetDefaultFileSystemRoot();
    }

    std::wstring workingDirText = workingDir.wstring();
    if (workingDirText.rfind(L"\\\\?\\UNC\\", 0) == 0 && workingDirText.size() > 8u)
    {
        workingDirText = std::wstring(L"\\\\") + workingDirText.substr(8u);
    }
    else if (workingDirText.rfind(L"\\\\?\\", 0) == 0 && workingDirText.size() > 4u)
    {
        workingDirText = workingDirText.substr(4u);
    }

    std::wstring comSpec;
    const DWORD comSpecLen = GetEnvironmentVariableW(L"ComSpec", nullptr, 0);
    if (comSpecLen > 0)
    {
        comSpec.resize(static_cast<size_t>(comSpecLen));
        const DWORD copied = GetEnvironmentVariableW(L"ComSpec", comSpec.data(), comSpecLen);
        if (copied > 0)
        {
            comSpec.resize(static_cast<size_t>(copied));
        }
        else
        {
            comSpec.clear();
        }
    }

    if (comSpec.empty())
    {
        comSpec = L"cmd.exe";
    }

    std::wstring parameters;
    std::wstring directory;

    const bool isUncPath = LooksLikeUncPath(workingDirText);
    const bool isCmd =
        (comSpec.size() >= 7u && wil::compare_string_ordinal(comSpec.substr(comSpec.size() - 7u), L"cmd.exe", true) == wistd::weak_ordering::equivalent);

    if (isUncPath && isCmd)
    {
        directory  = GetDefaultFileSystemRoot().wstring();
        parameters = std::format(L"/K pushd \"{}\"", workingDirText);
    }
    else
    {
        directory = std::move(workingDirText);
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    static_cast<void>(ShellExecuteW(ownerWindow,
                                    L"open",
                                    comSpec.c_str(),
                                    parameters.empty() ? nullptr : parameters.c_str(),
                                    directory.empty() ? nullptr : directory.c_str(),
                                    SW_SHOWNORMAL));
}

void FolderWindow::SwapPanes()
{
    CancelSelectionSizeComputation(Pane::Left);
    CancelSelectionSizeComputation(Pane::Right);

    _leftPane.folderView.CancelPendingEnumeration();
    _rightPane.folderView.CancelPendingEnumeration();

    const auto leftPluginPath  = _leftPane.folderView.GetFolderPath();
    const auto rightPluginPath = _rightPane.folderView.GetFolderPath();

    std::swap(_leftPane.fileSystemModule, _rightPane.fileSystemModule);
    std::swap(_leftPane.fileSystem, _rightPane.fileSystem);
    std::swap(_leftPane.pluginId, _rightPane.pluginId);
    std::swap(_leftPane.pluginShortId, _rightPane.pluginShortId);
    std::swap(_leftPane.instanceContext, _rightPane.instanceContext);

    _leftPane.folderView.SetFileSystem(_leftPane.fileSystem);
    _leftPane.folderView.SetFileSystemContext(_leftPane.pluginId, _leftPane.instanceContext);
    _leftPane.navigationView.SetFileSystem(_leftPane.fileSystem);
    _rightPane.folderView.SetFileSystem(_rightPane.fileSystem);
    _rightPane.folderView.SetFileSystemContext(_rightPane.pluginId, _rightPane.instanceContext);
    _rightPane.navigationView.SetFileSystem(_rightPane.fileSystem);

    auto applyPaneState = [&](PaneState& state, const std::optional<std::filesystem::path>& pluginPath)
    {
        std::optional<std::filesystem::path> displayPath;
        if (pluginPath.has_value())
        {
            displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, pluginPath.value());
        }

        FolderView::NameFilterState filter;
        if (displayPath.has_value())
        {
            const Common::Settings::FoldersSettings* folders = (_settings && _settings->folders.has_value()) ? &_settings->folders.value() : nullptr;
            filter                                           = GetFolderHistoryFilterState(folders, displayPath.value());
        }
        state.folderView.SetNameFilterState(filter, false /* refresh */);

        state.updatingPath = true;
        state.currentPath  = displayPath;
        state.navigationView.SetPath(displayPath);
        state.folderView.SetFolderPath(pluginPath);
        state.currentPath  = state.navigationView.GetPath();
        state.updatingPath = false;
    };

    applyPaneState(_leftPane, rightPluginPath);
    applyPaneState(_rightPane, leftPluginPath);

    _leftPane.selectionStats  = {};
    _rightPane.selectionStats = {};
    UpdatePaneStatusBar(Pane::Left);
    UpdatePaneStatusBar(Pane::Right);

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderWindow::OnNavigationPathChanged(Pane pane, const std::optional<std::filesystem::path>& path)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.updatingPath)
    {
        return;
    }

    if (! path)
    {
        state.updatingPath = true;
        state.currentPath.reset();
        state.folderView.SetNameFilterState(FolderView::NameFilterState{}, false /* refresh */);
        state.folderView.SetFolderPath(std::nullopt);
        state.updatingPath = false;
        if (_panePathChangedCallback)
        {
            _panePathChangedCallback(pane, std::nullopt);
        }
        return;
    }

    SetFolderPath(pane, path.value());
}

void FolderWindow::OnFolderViewPathChanged(Pane pane, const std::optional<std::filesystem::path>& path)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.updatingPath)
    {
        return;
    }

    if (! path)
    {
        state.updatingPath = true;
        state.currentPath.reset();
        state.navigationView.SetPath(std::nullopt);
        state.updatingPath = false;
        if (_panePathChangedCallback)
        {
            _panePathChangedCallback(pane, std::nullopt);
        }
        return;
    }

    FileSystemPluginManager& manager = FileSystemPluginManager::GetInstance();
    const std::wstring_view pluginId = state.pluginId.empty() ? manager.GetActivePluginId() : std::wstring_view(state.pluginId);

    std::wstring shortId = state.pluginShortId;
    if (shortId.empty())
    {
        const auto* entry = FindPluginById(manager.GetPlugins(), pluginId);
        if (entry)
        {
            shortId = entry->shortId;
        }
    }

    const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(shortId, state.instanceContext, path.value());

    const Common::Settings::FoldersSettings* folders = (_settings && _settings->folders.has_value()) ? &_settings->folders.value() : nullptr;
    const FolderView::NameFilterState filter         = GetFolderHistoryFilterState(folders, displayPath);
    state.folderView.SetNameFilterState(filter, false /* refresh */);

    state.updatingPath = true;
    state.currentPath  = displayPath;

    if (state.hNavigationView)
    {
        state.navigationView.SetPath(displayPath);
    }

    state.updatingPath = false;

    RecordNavigationHistory(state, displayPath);
    AddToFolderHistory(_folderHistory, static_cast<size_t>(_folderHistoryMax), displayPath);
    _leftPane.navigationView.SetHistory(_folderHistory);
    _rightPane.navigationView.SetHistory(_folderHistory);

    if (_settings)
    {
        Common::Settings::FoldersSettings& foldersSettings = _settings->folders.has_value() ? _settings->folders.value() : _settings->folders.emplace();
        PruneFolderHistoryFilters(foldersSettings, _folderHistory, static_cast<size_t>(_folderHistoryMax));
    }

    if (_panePathChangedCallback)
    {
        _panePathChangedCallback(pane, path);
    }
}

void FolderWindow::OnFolderViewNavigateUpFromRoot(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.updatingPath)
    {
        return;
    }

    if (state.instanceContext.empty())
    {
        return;
    }

    if (IsFilePluginShortId(state.pluginShortId))
    {
        return;
    }

    const std::optional<std::filesystem::path> pluginPathOpt = state.folderView.GetFolderPath();
    if (! pluginPathOpt.has_value())
    {
        return;
    }

    const std::filesystem::path pluginPath   = pluginPathOpt.value();
    const std::filesystem::path pluginParent = pluginPath.parent_path();
    if (! pluginParent.empty() && pluginParent != pluginPath)
    {
        return;
    }

    const std::optional<std::filesystem::path> mountPointOpt = TryResolveInstanceContextToWindowsPath(state.instanceContext);
    if (! mountPointOpt.has_value())
    {
        return;
    }

    std::filesystem::path mountPoint = mountPointOpt.value().lexically_normal();
    if (! mountPoint.has_filename())
    {
        const std::filesystem::path trimmed = mountPoint.parent_path();
        if (! trimmed.empty())
        {
            mountPoint = trimmed;
        }
    }

    std::filesystem::path mountParent = mountPoint.parent_path();
    if (mountParent.empty())
    {
        mountParent = GetDefaultFileSystemRoot();
    }

    const std::wstring focusName = mountPoint.filename().wstring();
    if (! focusName.empty())
    {
        state.folderView.RememberFocusedItemForFolder(mountParent, focusName);
    }

    SetFolderPath(pane, mountParent);
}
