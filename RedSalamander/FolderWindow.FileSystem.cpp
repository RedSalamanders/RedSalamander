#include "ConnectionManagerDialog.h"
#include "ConnectionSecrets.h"
#include "ChangeCase.h"
#include "FolderWindowInternal.h"
#include "HostServices.h"
#include "NavigationLocation.h"

#include "SettingsStore.h"
#include "ThemedControls.h"

#include <shellapi.h>

namespace
{
bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    const int len = static_cast<int>(a.size());
    return CompareStringOrdinal(a.data(), len, b.data(), len, TRUE) == CSTR_EQUAL;
}

bool IsFilePluginShortId(std::wstring_view pluginShortId) noexcept
{
    return EqualsNoCase(pluginShortId, L"file");
}

[[nodiscard]] bool StartsWithNoCase(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (text.size() < prefix.size())
    {
        return false;
    }

    return EqualsNoCase(text.substr(0, prefix.size()), prefix);
}

 struct ChangeCaseDialogState
 {
    ChangeCaseDialogState()                                         = default;
    ChangeCaseDialogState(const ChangeCaseDialogState&)             = delete;
    ChangeCaseDialogState& operator=(const ChangeCaseDialogState&)  = delete;
    ChangeCaseDialogState(ChangeCaseDialogState&&)                  = delete;
    ChangeCaseDialogState& operator=(ChangeCaseDialogState&&)       = delete;
    ~ChangeCaseDialogState()                                        = default;

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
    bool useTwoColumns               = false;
    int twoColumnSeparatorX          = -1;
    int twoColumnSeparatorYTop       = 0;
    int twoColumnSeparatorYBottom    = 0;
    wil::unique_hfont boldFont;
    wil::unique_hfont italicFont;
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
        case WM_NCDESTROY:
            RemoveWindowSubclass(hwnd, ChangeCaseToggleCheckSubclassProc, subclassId);
            break;
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

    EnumChildWindows(
        dlg,
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

    const DWORD toggleStyle =
        static_cast<DWORD>(WS_CHILD | WS_VISIBLE | WS_TABSTOP) | static_cast<DWORD>(highContrast ? BS_AUTOCHECKBOX : 0);

    const auto makeStatic = [&](DWORD style) noexcept -> HWND
    { return CreateWindowExW(0, L"Static", L"", style, 0, 0, 10, 10, dlg, nullptr, instance, nullptr); };

    const auto makeToggle = [&](int id) noexcept -> HWND
    {
        const HWND toggle = CreateWindowExW(
            0, L"Button", L"", toggleStyle, 0, 0, 10, 10, dlg, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), instance, nullptr);
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

    state.ui.headerChangeCaseTo  = makeStatic(baseStaticStyle);
    state.ui.lower.title         = makeStatic(baseStaticStyle);
    state.ui.lower.description   = makeStatic(wrapStaticStyle);
    state.ui.lower.toggle        = makeToggle(IDC_CHANGE_CASE_LOWER);
    state.ui.upper.title         = makeStatic(baseStaticStyle);
    state.ui.upper.description   = makeStatic(wrapStaticStyle);
    state.ui.upper.toggle        = makeToggle(IDC_CHANGE_CASE_UPPER);
    state.ui.partiallyMixed.title       = makeStatic(baseStaticStyle);
    state.ui.partiallyMixed.description = makeStatic(wrapStaticStyle);
    state.ui.partiallyMixed.toggle      = makeToggle(IDC_CHANGE_CASE_PARTIALLY_MIXED);
    state.ui.mixed.title         = makeStatic(baseStaticStyle);
    state.ui.mixed.description   = makeStatic(wrapStaticStyle);
    state.ui.mixed.toggle        = makeToggle(IDC_CHANGE_CASE_MIXED);

    state.ui.headerChange        = makeStatic(baseStaticStyle);
    state.ui.whole.title         = makeStatic(baseStaticStyle);
    state.ui.whole.description   = makeStatic(wrapStaticStyle);
    state.ui.whole.toggle        = makeToggle(IDC_CHANGE_CASE_WHOLE);
    state.ui.onlyName.title      = makeStatic(baseStaticStyle);
    state.ui.onlyName.description = makeStatic(wrapStaticStyle);
    state.ui.onlyName.toggle     = makeToggle(IDC_CHANGE_CASE_ONLY_NAME);
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
        std::wstring text(static_cast<size_t>(len), L'\0');
        const int copied = GetWindowTextW(hwnd, text.data(), len + 1);
        if (copied <= 0)
        {
            return {};
        }
        if (static_cast<size_t>(copied) < text.size())
        {
            text.resize(static_cast<size_t>(copied));
        }
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

    const int okW     = measureButtonWidth(okBtn);
    const int cancelW = measureButtonWidth(cancelBtn);
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

    const auto computeToggleWidth = [&](int availableW) noexcept -> int
    { return std::min(std::max(0, availableW - 2 * cardPaddingX), measuredToggleW); };

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
            SetWindowPos(card.description,
                         nullptr,
                         cardRc.left + cardPaddingX,
                         cardRc.top + cardPaddingY + titleHeight + cardGapY,
                         textW,
                         std::max(0, descH),
                         flags);
            SendMessageW(card.description, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        }
        else
        {
            ShowWindow(card.description, SW_HIDE);
        }

        if (card.toggle)
        {
            SetWindowPos(card.toggle,
                         nullptr,
                         cardRc.right - cardPaddingX - toggleW,
                         cardRc.top + (cardH - rowHeight) / 2,
                         toggleW,
                         rowHeight,
                         flags);
            SendMessageW(card.toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        y += cardH + (addBottomSpacing ? cardSpacingY : 0);
    };

    int y = contentY;

    const int toggleWFull = computeToggleWidth(contentW);

    layoutToggleCardAt(state.ui.includeSubdirs,
                       L"Include subdirectories",
                       L"Apply to selected folders recursively.",
                       true,
                       contentX,
                       contentW,
                       toggleWFull,
                       true,
                       y);

    y += sectionSpacing;

     const bool useTwoColumnsLayout = contentW >= (minLeftColumnW + minRightColumnW + columnSeparatorAreaW);

     if (useTwoColumnsLayout)
     {
         const int availableW = std::max(0, contentW - columnSeparatorAreaW);
         const int minCombined = minLeftColumnW + minRightColumnW;
         const int extra       = std::max(0, availableW - minCombined);
 
         const int leftExtra = (extra * 2) / 3;
         const int leftW     = std::max(0, minLeftColumnW + leftExtra);
         const int rightW    = std::max(0, availableW - leftW);
         const int leftX  = contentX;
         const int rightX = contentX + leftW + columnSeparatorAreaW;

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

 void EnsureChangeCaseDialogAllOptionsVisible(HWND dlg, ChangeCaseDialogState& state) noexcept
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
 
     const UINT dpi  = GetDpiForWindow(dlg);
     const int gapY  = ThemedControls::ScaleDip(dpi, 12);
     const int desiredBottom = contentBottom + gapY;
     if (desiredBottom <= buttonsTop)
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
 
     int newWindowH       = windowH + (desiredBottom - buttonsTop);
 
     HMONITOR monitor = MonitorFromWindow(dlg, MONITOR_DEFAULTTONEAREST);
     MONITORINFO mi{};
     mi.cbSize = sizeof(mi);
     if (monitor && GetMonitorInfoW(monitor, &mi))
     {
         const int workH = static_cast<int>(mi.rcWork.bottom - mi.rcWork.top);
         newWindowH      = std::min(newWindowH, workH);
     }
 
     if (newWindowH <= windowH)
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
     EnsureChangeCaseDialogAllOptionsVisible(dlg, *state);

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

    const bool enabled = ! control || IsWindowEnabled(control) != FALSE;
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

    const bool enabled = ! control || IsWindowEnabled(control) != FALSE;
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
        case IDCANCEL:
            EndDialog(dlg, IDCANCEL);
            return TRUE;
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
            }
            return TRUE;
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

            const UINT id = dis->CtlID;
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
                    case IDC_CHANGE_CASE_MIXED:
                        CheckRadioButton(dlg, IDC_CHANGE_CASE_LOWER, IDC_CHANGE_CASE_MIXED, id);
                        return TRUE;
                    case IDC_CHANGE_CASE_WHOLE:
                    case IDC_CHANGE_CASE_ONLY_NAME:
                    case IDC_CHANGE_CASE_ONLY_EXTENSION:
                        CheckRadioButton(dlg, IDC_CHANGE_CASE_WHOLE, IDC_CHANGE_CASE_ONLY_EXTENSION, id);
                        return TRUE;
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
        const bool exists                 = std::find_if(normalized.begin(),
                                         normalized.end(),
                                         [&](const std::filesystem::path& existing) { return EqualsNoCase(existing.native(), entryText); }) != normalized.end();
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
    auto it = std::find_if(history.begin(), history.end(), [&](const std::filesystem::path& existing) { return EqualsNoCase(existing.native(), entryText); });

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

        const int len = static_cast<int>(idSize);
        if (CompareStringOrdinal(entry.shortId.c_str(), len, shortId.data(), len, TRUE) == CSTR_EQUAL)
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

        const int len = static_cast<int>(idSize);
        if (CompareStringOrdinal(entry.id.c_str(), len, pluginId.data(), len, TRUE) == CSTR_EQUAL)
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

std::wstring TrimWhitespace(std::wstring_view text)
{
    std::wstring result(text);
    result.erase(result.begin(), std::find_if(result.begin(), result.end(), [](wchar_t ch) { return iswspace(ch) == 0; }));
    result.erase(std::find_if(result.rbegin(), result.rend(), [](wchar_t ch) { return iswspace(ch) == 0; }).base(), result.end());
    return result;
}

std::optional<std::filesystem::path> TryResolveInstanceContextToWindowsPath(std::wstring_view instanceContext) noexcept
{
    if (instanceContext.empty())
    {
        return std::nullopt;
    }

    std::wstring text = TrimWhitespace(instanceContext);
    if (text.empty())
    {
        return std::nullopt;
    }

    if (text.size() >= 2u && text.front() == L'"' && text.back() == L'"')
    {
        text.erase(text.begin());
        text.pop_back();
        text = TrimWhitespace(text);
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

    const std::wstring trimmed = TrimWhitespace(raw);
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

INT_PTR OnCreateDirectoryDialogCtlColorEdit(CreateDirectoryDialogState* state, HDC hdc)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    SetBkColor(hdc, state->theme.windowBackground);
    SetTextColor(hdc, state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
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

    std::wstring trimmed = TrimWhitespace(buffer);
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
        case WM_CTLCOLOREDIT: return OnCreateDirectoryDialogCtlColorEdit(state, reinterpret_cast<HDC>(wParam));
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->theme, wParam != FALSE);
            }
            return FALSE;
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

    if (state.fileSystem && state.fileSystemModule && EqualsNoCase(state.pluginId, pluginId))
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
            const auto it =
                std::find_if(conns.begin(),
                             conns.end(),
                             [&](const Common::Settings::ConnectionProfile& c) noexcept { return ! c.name.empty() && EqualsNoCase(c.name, connectionName); });
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
            state.folderView.SetFolderPath(pluginPath);
        }

        state.updatingPath = false;
    }

    {
        Debug::Perf::Scope historyPerf(pane == Pane::Left ? L"FolderWindow.SetFolderPath.Left.UpdateHistory"
                                                          : L"FolderWindow.SetFolderPath.Right.UpdateHistory");
        historyPerf.SetDetail(displayPath.native());

        AddToFolderHistory(_folderHistory, static_cast<size_t>(_folderHistoryMax), displayPath);
        _leftPane.navigationView.SetHistory(_folderHistory);
        _rightPane.navigationView.SetHistory(_folderHistory);
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
            if (otherState.fileSystem && otherFolder.has_value() && EqualsNoCase(otherFolder.value().native(), base.native()) &&
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

        state.folderView.ShowAlertOverlay(
            FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
    }
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

    const HWND ownerHwnd = _hWnd.get();
    wil::com_ptr<IFileSystem> fileSystem = state.fileSystem;
    const ChangeCase::Options options    = dialogState.options;
    const std::wstring title             = LoadStringResource(nullptr, IDS_CMD_CHANGE_CASE);

    state.changeCaseThread = std::jthread(
        [ownerHwnd, pane, fileSystem, paths, options, title](std::stop_token stopToken) noexcept
        {
            [[maybe_unused]] auto coInit = wil::CoInitializeEx();

            struct ProgressState final
            {
                HWND hwnd                    = nullptr;
                FolderWindow::Pane pane      = FolderWindow::Pane::Left;
                std::wstring title;
                ULONGLONG startTick          = 0;
                ULONGLONG lastPostedTick     = 0;
                uint64_t infoTaskId          = 0;
                ChangeCase::ProgressUpdate last{};

                void PostTaskUpdate(bool finished, HRESULT hr) noexcept
                {
                    if (! hwnd || IsWindow(hwnd) == FALSE || infoTaskId == 0)
                    {
                        return;
                    }

                    FolderWindow::InformationalTaskUpdate info{};
                    info.kind  = FolderWindow::InformationalTaskUpdate::Kind::ChangeCase;
                    info.taskId = infoTaskId;
                    info.title = title;
                    info.changeCaseCurrentPath      = last.currentPath;
                    info.changeCaseScannedFolders   = last.scannedFolders;
                    info.changeCaseScannedEntries   = last.scannedEntries;
                    info.changeCasePlannedRenames   = last.plannedRenames;
                    info.changeCaseCompletedRenames = last.completedRenames;
                    info.changeCaseEnumerating      = ! finished && last.phase == ChangeCase::ProgressUpdate::Phase::Enumerating;
                    info.changeCaseRenaming         = ! finished && last.phase == ChangeCase::ProgressUpdate::Phase::Renaming;
                    info.finished                   = finished;
                    info.resultHr                   = hr;

                    auto payload     = std::make_unique<ChangeCaseTaskPayload>();
                    payload->update  = std::move(info);
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
                    info.kind  = FolderWindow::InformationalTaskUpdate::Kind::ChangeCase;
                    info.title = title;
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
                        SendMessageTimeoutW(hwnd,
                                            WndMsg::kChangeCaseTaskUpdate,
                                            0,
                                            reinterpret_cast<LPARAM>(raw),
                                            SMTO_ABORTIFHUNG | SMTO_BLOCK,
                                            1000,
                                            &result);
                    if (sendOk == 0)
                    {
                        delete raw;
                        return;
                    }

                    infoTaskId      = static_cast<uint64_t>(result);
                    lastPostedTick  = nowTick;
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
                auto completed       = std::make_unique<ChangeCaseCompletedPayload>();
                completed->pane      = pane;
                completed->hr        = operationHr;
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

    state.updatingPath = true;
    state.currentPath  = displayPath;

    if (state.hNavigationView)
    {
        state.navigationView.SetPath(displayPath);
    }

    state.updatingPath = false;

    AddToFolderHistory(_folderHistory, static_cast<size_t>(_folderHistoryMax), displayPath);
    _leftPane.navigationView.SetHistory(_folderHistory);
    _rightPane.navigationView.SetHistory(_folderHistory);

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
