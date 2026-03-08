#include "Framework.h"

#include "CompareDirectoriesWindow.Internal.h"
#include "SettingsHotReload.h"

namespace CompareDirectoriesWindowInternal
{
namespace
{
constexpr wchar_t kSettingsAppId[] = L"RedSalamander";

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
[[nodiscard]] bool GetTwoStateToggleState(HWND toggle, bool highContrast) noexcept
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
} // namespace

void CompareDirectoriesWindow::ApplyOptionsDialogTheme() noexcept
{
    if (! _optionsDlg)
    {
        return;
    }

    const bool darkBackground = ChooseContrastingTextColor(_theme.windowBackground) == RGB(255, 255, 255);
    const wchar_t* themeName  = _theme.highContrast ? L"" : (darkBackground ? L"DarkMode_Explorer" : L"Explorer");

    SetWindowTheme(_optionsDlg.get(), themeName, nullptr);
    SendMessageW(_optionsDlg.get(), WM_THEMECHANGED, 0, 0);

    const HFONT font = _uiFont.get();
    if (font)
    {
        SendMessageW(_optionsDlg.get(), WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
    }

    struct EnumData
    {
        HFONT font               = nullptr;
        const wchar_t* themeName = nullptr;
        HWND optionsHost         = nullptr;
    };

    EnumData data{};
    data.font        = font;
    data.themeName   = themeName;
    data.optionsHost = _optionsUi.host;

    EnumChildWindows(_optionsDlg.get(),
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto* data = reinterpret_cast<const EnumData*>(lParam);
        if (! data || ! child)
        {
            return TRUE;
        }

        if (data->font)
        {
            SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(data->font), FALSE);
        }

        if (data->themeName)
        {
            std::array<wchar_t, 32> className{};
            const int classLen = GetClassNameW(child, className.data(), static_cast<int>(className.size()));

            const wchar_t* appliedTheme = data->themeName;
            if (classLen > 0)
            {
                if (_wcsicmp(className.data(), L"Static") == 0)
                {
                    appliedTheme = child == data->optionsHost ? data->themeName : L"";
                }
                else if (_wcsicmp(className.data(), L"Button") == 0)
                {
                    const LONG_PTR style = GetWindowLongPtrW(child, GWL_STYLE);
                    const LONG_PTR type  = style & BS_TYPEMASK;
                    if (type == BS_GROUPBOX || type == BS_PUSHBUTTON || type == BS_DEFPUSHBUTTON)
                    {
                        appliedTheme = L"";
                    }
                }
            }

            SetWindowTheme(child, appliedTheme, nullptr);
            SendMessageW(child, WM_THEMECHANGED, 0, 0);
        }

        return TRUE;
    },
                     reinterpret_cast<LPARAM>(&data));

    if (data.optionsHost && data.themeName)
    {
        EnumChildWindows(data.optionsHost,
                         [](HWND child, LPARAM lParam) noexcept -> BOOL
        {
            auto* data = reinterpret_cast<const EnumData*>(lParam);
            if (! data || ! child)
            {
                return TRUE;
            }

            if (data->font)
            {
                SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(data->font), FALSE);
            }

            std::array<wchar_t, 32> className{};
            const int classLen = GetClassNameW(child, className.data(), static_cast<int>(className.size()));

            const wchar_t* appliedTheme = data->themeName;
            if (classLen > 0)
            {
                if (_wcsicmp(className.data(), L"Static") == 0)
                {
                    appliedTheme = L"";
                }
                else if (_wcsicmp(className.data(), L"Button") == 0)
                {
                    appliedTheme = L"";
                }
            }

            SetWindowTheme(child, appliedTheme, nullptr);
            SendMessageW(child, WM_THEMECHANGED, 0, 0);
            return TRUE;
        },
                         reinterpret_cast<LPARAM>(&data));
    }

    RedrawWindow(_optionsDlg.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
}

INT_PTR CALLBACK CompareDirectoriesWindow::OptionsDlgProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam) noexcept
{
    if (msg == WM_INITDIALOG)
    {
        auto* self = reinterpret_cast<CompareDirectoriesWindow*>(lParam);
        SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(self));
        return self ? self->OnOptionsInitDialog(dlg) : TRUE;
    }

    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(GetWindowLongPtrW(dlg, DWLP_USER));
    if (! self)
    {
        return FALSE;
    }

    switch (msg)
    {
        case WM_ERASEBKGND: return self->OnOptionsEraseBkgnd(dlg, reinterpret_cast<HDC>(wParam));
        case WM_COMMAND: return self->OnOptionsCommand(dlg, wParam, lParam);
        case WndMsg::kSettingsReloadedFromDisk: return self->OnOptionsSettingsReloadedFromDisk(dlg);
        case WM_DRAWITEM: return self->OnOptionsDrawItem(reinterpret_cast<const DRAWITEMSTRUCT*>(lParam));
        case WM_CTLCOLOREDIT: return self->OnOptionsCtlColorEdit(reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_CTLCOLORDLG: return self->OnOptionsCtlColorDlg(reinterpret_cast<HDC>(wParam));
        case WM_CTLCOLORSTATIC: return self->OnOptionsCtlColorStatic(reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_CTLCOLORBTN: return self->OnOptionsCtlColorBtn(reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_NCDESTROY: SettingsHotReload::UnregisterParticipant(dlg); return FALSE;
        default: break;
    }

    return FALSE;
}

INT_PTR CompareDirectoriesWindow::OnOptionsInitDialog(HWND dlg) noexcept
{
    const bool darkBackground = ChooseContrastingTextColor(_theme.windowBackground) == RGB(255, 255, 255);
    const wchar_t* themeName  = _theme.highContrast ? L"" : (darkBackground ? L"DarkMode_Explorer" : L"Explorer");

    SetWindowTheme(dlg, themeName, nullptr);
    SendMessageW(dlg, WM_THEMECHANGED, 0, 0);

    if (! _theme.highContrast)
    {
        ThemedControls::EnableOwnerDrawButton(dlg, IDOK);
        ThemedControls::EnableOwnerDrawButton(dlg, IDCANCEL);
    }

    SettingsHotReload::RegisterParticipant(dlg);
    EnsureOptionsControlsCreated(dlg);
    return TRUE;
}

INT_PTR CompareDirectoriesWindow::OnOptionsEraseBkgnd(HWND dlg, HDC hdc) noexcept
{
    if (! _optionsBackgroundBrush)
    {
        return FALSE;
    }

    RECT rc{};
    GetClientRect(dlg, &rc);
    FillRect(hdc, &rc, _optionsBackgroundBrush.get());
    return TRUE;
}

INT_PTR CompareDirectoriesWindow::OnOptionsSettingsReloadedFromDisk(HWND dlg) noexcept
{
    if (IsOptionsDialogDirty())
    {
        SettingsHotReload::ExternalReloadChoice choice = SettingsHotReload::ExternalReloadChoice::KeepEditing;
        const HRESULT promptHr = SettingsHotReload::PromptExternalReloadConflict(dlg, LoadStringResource(nullptr, IDS_COMPARE_DIRECTORIES_TITLE), choice);
        if (FAILED(promptHr))
        {
            Debug::Warning(L"CompareDirectories: failed to prompt for external reload conflict (hr=0x{:08X})", static_cast<unsigned long>(promptHr));
            return TRUE;
        }

        if (choice == SettingsHotReload::ExternalReloadChoice::KeepEditing)
        {
            _optionsStaleFromExternalReload = true;
            return TRUE;
        }
    }

    ReloadOptionsDialogFromDisk();
    return TRUE;
}

INT_PTR CompareDirectoriesWindow::OnOptionsCommand([[maybe_unused]] HWND dlg, WPARAM wParam, LPARAM lParam) noexcept
{
    const UINT controlId  = LOWORD(wParam);
    const UINT notifyCode = HIWORD(wParam);
    HWND hwndCtl          = reinterpret_cast<HWND>(lParam);

    if (notifyCode == BN_CLICKED && hwndCtl)
    {
        const LONG_PTR style = GetWindowLongPtrW(hwndCtl, GWL_STYLE);
        if ((style & BS_TYPEMASK) == BS_OWNERDRAW)
        {
            switch (controlId)
            {
                case IDC_CMP_SIZE:
                case IDC_CMP_DATETIME:
                case IDC_CMP_ATTRIBUTES:
                case IDC_CMP_CONTENT:
                case IDC_CMP_SUBDIRECTORIES:
                case IDC_CMP_SUBDIR_ATTRIBUTES:
                case IDC_CMP_SELECT_SUBDIRS_ONLY_ONE_PANE:
                case IDC_CMP_KEEP_IDENTICAL:
                case IDC_CMP_IGNORE_FILES:
                case IDC_CMP_IGNORE_DIRECTORIES:
                {
                    const bool toggledOn = GetTwoStateToggleState(hwndCtl, false);
                    SetTwoStateToggleState(hwndCtl, false, ! toggledOn);
                    break;
                }
                default: break;
            }
        }
    }

    switch (controlId)
    {
        case IDOK:
            if (! ResolveOptionsStaleSaveConflict(dlg))
            {
                return TRUE;
            }

            SaveOptionsControlsToSettings();
            if (_settings)
            {
                const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(kSettingsAppId, *_settings);
                if (FAILED(saveHr))
                {
                    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kSettingsAppId);
                    const std::wstring title                 = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
                    const std::wstring message =
                        FormatStringResource(nullptr, IDS_FMT_SETTINGS_SAVE_FAILED, settingsPath.wstring(), static_cast<unsigned long>(saveHr));
                    ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, message);
                    return TRUE;
                }
            }
            UpdateViewMenuChecks();
            ShowOptionsPanel(false);
            ScheduleBeginOrRescanCompare();
            return TRUE;
        case IDCANCEL:
            if (! _compareStarted)
            {
                PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0);
                return TRUE;
            }
            ShowOptionsPanel(false);
            return TRUE;
        case IDC_CMP_IGNORE_FILES:
        case IDC_CMP_IGNORE_DIRECTORIES: UpdateOptionsVisibility(); return TRUE;
        default: break;
    }

    return FALSE;
}

INT_PTR CompareDirectoriesWindow::OnOptionsDrawItem(const DRAWITEMSTRUCT* dis) noexcept
{
    if (! dis || dis->CtlType != ODT_BUTTON)
    {
        return FALSE;
    }

    const LONG_PTR style = dis->hwndItem ? GetWindowLongPtrW(dis->hwndItem, GWL_STYLE) : 0;
    if ((style & BS_TYPEMASK) == BS_OWNERDRAW)
    {
        const UINT id       = dis->CtlID;
        const bool isToggle = id == IDC_CMP_SIZE || id == IDC_CMP_DATETIME || id == IDC_CMP_ATTRIBUTES || id == IDC_CMP_CONTENT ||
                              id == IDC_CMP_SUBDIRECTORIES || id == IDC_CMP_SUBDIR_ATTRIBUTES || id == IDC_CMP_SELECT_SUBDIRS_ONLY_ONE_PANE ||
                              id == IDC_CMP_KEEP_IDENTICAL || id == IDC_CMP_IGNORE_FILES || id == IDC_CMP_IGNORE_DIRECTORIES;
        if (isToggle)
        {
            const bool toggledOn        = GetWindowLongPtrW(dis->hwndItem, GWLP_USERDATA) != 0;
            const COLORREF surface      = ThemedControls::GetControlSurfaceColor(_theme);
            const HFONT boldFont        = _uiBoldFont ? _uiBoldFont.get() : nullptr;
            const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
            const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
            ThemedControls::DrawThemedSwitchToggle(*dis, _theme, surface, boldFont, onLabel, offLabel, toggledOn);
            return TRUE;
        }
    }

    ThemedControls::DrawThemedPushButton(*dis, _theme);
    return TRUE;
}

INT_PTR CompareDirectoriesWindow::OnOptionsCtlColorEdit(HDC hdc, HWND control) noexcept
{
    if (! _optionsInputBrush)
    {
        return FALSE;
    }

    const bool enabled = ! control || IsWindowEnabled(control) != FALSE;
    const bool focused = enabled && control && GetFocus() == control;
    const COLORREF bg  = enabled ? (focused ? _optionsInputFocusedBackgroundColor : _optionsInputBackgroundColor) : _optionsInputDisabledBackgroundColor;

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, bg);
    SetTextColor(hdc, enabled ? _theme.menu.text : _theme.menu.disabledText);

    if (_theme.highContrast)
    {
        return reinterpret_cast<INT_PTR>(_optionsBackgroundBrush.get());
    }

    if (! enabled)
    {
        return reinterpret_cast<INT_PTR>(_optionsInputDisabledBrush.get());
    }

    return reinterpret_cast<INT_PTR>(focused && _optionsInputFocusedBrush ? _optionsInputFocusedBrush.get() : _optionsInputBrush.get());
}

INT_PTR CompareDirectoriesWindow::OnOptionsCtlColorDlg(HDC hdc) noexcept
{
    if (! _optionsBackgroundBrush)
    {
        return FALSE;
    }

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, _theme.windowBackground);
    SetTextColor(hdc, _theme.menu.text);
    return reinterpret_cast<INT_PTR>(_optionsBackgroundBrush.get());
}

INT_PTR CompareDirectoriesWindow::OnOptionsCtlColorStatic(HDC hdc, HWND control) noexcept
{
    if (! _optionsBackgroundBrush)
    {
        return FALSE;
    }

    COLORREF textColor = _theme.menu.text;
    if (control && IsWindowEnabled(control) == FALSE)
    {
        textColor = _theme.menu.disabledText;
    }

    if (_theme.systemHighContrast || _theme.highContrast)
    {
        SetBkMode(hdc, OPAQUE);
        SetBkColor(hdc, _theme.windowBackground);
        SetTextColor(hdc, textColor);
        return reinterpret_cast<INT_PTR>(_optionsBackgroundBrush.get());
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    SetBkColor(hdc, _theme.windowBackground);

    HBRUSH brush = _optionsBackgroundBrush.get();
    if (control && _optionsUi.host && _optionsCardBrush && ! _optionsCards.empty())
    {
        RECT rc{};
        if (GetWindowRect(control, &rc) != FALSE)
        {
            MapWindowPoints(nullptr, _optionsUi.host, reinterpret_cast<POINT*>(&rc), 2);
            for (const RECT& card : _optionsCards)
            {
                RECT intersect{};
                if (IntersectRect(&intersect, &card, &rc) != FALSE)
                {
                    brush = _optionsCardBrush.get();
                    break;
                }
            }
        }
    }

    return reinterpret_cast<INT_PTR>(brush);
}

INT_PTR CompareDirectoriesWindow::OnOptionsCtlColorBtn(HDC hdc, HWND control) noexcept
{
    if (! _optionsBackgroundBrush)
    {
        return FALSE;
    }

    const LONG_PTR style = control ? GetWindowLongPtrW(control, GWL_STYLE) : 0;
    const LONG_PTR type  = style & BS_TYPEMASK;

    const bool themed = type == BS_CHECKBOX || type == BS_AUTOCHECKBOX || type == BS_RADIOBUTTON || type == BS_AUTORADIOBUTTON || type == BS_3STATE ||
                        type == BS_AUTO3STATE || type == BS_GROUPBOX;
    if (! themed)
    {
        return FALSE;
    }

    const bool enabled = ! control || IsWindowEnabled(control) != FALSE;
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, enabled ? _theme.menu.text : _theme.menu.disabledText);
    SetBkColor(hdc, _theme.windowBackground);
    return reinterpret_cast<INT_PTR>(_optionsBackgroundBrush.get());
}

LRESULT CALLBACK CompareOptionsHostSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept
{
    auto* self     = reinterpret_cast<CompareDirectoriesWindow*>(refData);
    const HWND dlg = GetParent(hwnd);

    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_PRINTCLIENT:
        {
            if (self)
            {
                self->PaintOptionsHostBackgroundAndCards(reinterpret_cast<HDC>(wp), hwnd);
            }
            return 0;
        }
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (! hdc)
            {
                return 0;
            }

            RECT client{};
            GetClientRect(hwnd, &client);
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
                if (self)
                {
                    self->PaintOptionsHostBackgroundAndCards(memDc.get(), hwnd);
                }
                BitBlt(hdc.get(), 0, 0, width, height, memDc.get(), 0, 0, SRCCOPY);
            }
            else if (self)
            {
                self->PaintOptionsHostBackgroundAndCards(hdc.get(), hwnd);
            }
            return 0;
        }
        case WM_VSCROLL:
        {
            if (! self || self->_optionsScrollMax <= 0)
            {
                break;
            }

            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_ALL;
            GetScrollInfo(hwnd, SB_VERT, &si);

            const UINT dpi  = GetDpiForWindow(hwnd);
            const int lineY = ThemedControls::ScaleDip(dpi, 24);

            int newPos = self->_optionsScrollOffset;
            switch (LOWORD(wp))
            {
                case SB_TOP: newPos = 0; break;
                case SB_BOTTOM: newPos = self->_optionsScrollMax; break;
                case SB_LINEUP: newPos -= lineY; break;
                case SB_LINEDOWN: newPos += lineY; break;
                case SB_PAGEUP: newPos -= static_cast<int>(si.nPage); break;
                case SB_PAGEDOWN: newPos += static_cast<int>(si.nPage); break;
                case SB_THUMBTRACK: newPos = si.nTrackPos; break;
                case SB_THUMBPOSITION: newPos = si.nPos; break;
                default: break;
            }

            newPos = std::clamp(newPos, 0, self->_optionsScrollMax);
            if (newPos != self->_optionsScrollOffset)
            {
                self->_optionsScrollOffset = newPos;
                self->LayoutOptionsControls();
            }
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            if (! self || self->_optionsScrollMax <= 0)
            {
                break;
            }

            const int delta = GET_WHEEL_DELTA_WPARAM(wp);
            if (delta == 0)
            {
                return 0;
            }

            self->_optionsWheelRemainder += delta;
            const int notches = self->_optionsWheelRemainder / WHEEL_DELTA;
            self->_optionsWheelRemainder -= notches * WHEEL_DELTA;
            if (notches == 0)
            {
                return 0;
            }

            UINT linesPerNotch = 3;
            SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &linesPerNotch, 0);
            if (linesPerNotch == 0)
            {
                return 0;
            }

            const UINT dpi  = GetDpiForWindow(hwnd);
            const int lineY = ThemedControls::ScaleDip(dpi, 32);

            int scrollDelta = 0;
            if (linesPerNotch == WHEEL_PAGESCROLL)
            {
                SCROLLINFO si{};
                si.cbSize = sizeof(si);
                si.fMask  = SIF_PAGE;
                GetScrollInfo(hwnd, SB_VERT, &si);
                scrollDelta = notches * static_cast<int>(si.nPage);
            }
            else
            {
                scrollDelta = notches * lineY * static_cast<int>(linesPerNotch);
            }

            const int newPos = std::clamp(self->_optionsScrollOffset - scrollDelta, 0, self->_optionsScrollMax);
            if (newPos != self->_optionsScrollOffset)
            {
                self->_optionsScrollOffset = newPos;
                self->LayoutOptionsControls();
            }
            return 0;
        }
        case WM_COMMAND:
        case WM_NOTIFY:
        case WM_DRAWITEM:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORBTN:
            if (dlg)
            {
                return SendMessageW(dlg, msg, wp, lp);
            }
            break;
        case WM_NCDESTROY:
        {
            RemoveWindowSubclass(hwnd, CompareOptionsHostSubclassProc, subclassId);
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}
LRESULT CALLBACK CompareOptionsWheelRouteSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept
{
    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(refData);
    if (! self)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_MOUSEWHEEL:
        {
            if (! self->_optionsDlg || ! self->_optionsUi.host)
            {
                return 0;
            }

            static thread_local bool s_routingWheel = false;
            if (s_routingWheel && hwnd != self->_optionsUi.host)
            {
                return 0;
            }

            POINT ptScreen{};
            ptScreen.x = GET_X_LPARAM(lp);
            ptScreen.y = GET_Y_LPARAM(lp);

            RECT dlgRect{};
            if (GetWindowRect(self->_optionsDlg.get(), &dlgRect) == 0 || PtInRect(&dlgRect, ptScreen) == FALSE)
            {
                // Don't scroll the options dialog when the user is wheeling outside it.
                return 0;
            }

            RECT hostRect{};
            if (GetWindowRect(self->_optionsUi.host, &hostRect) == 0 || PtInRect(&hostRect, ptScreen) == FALSE)
            {
                // Only scroll when the wheel is over the options host area.
                return 0;
            }

            if (self->_optionsScrollMax <= 0)
            {
                // Nothing to scroll; swallow the wheel to avoid comctl32 forwarding loops.
                return 0;
            }

            if (hwnd != self->_optionsUi.host)
            {
                s_routingWheel = true;
                auto reset     = wil::scope_exit([&] { s_routingWheel = false; });
                SendMessageW(self->_optionsUi.host, msg, wp, lp);
                return 0;
            }

            break;
        }
        case WM_NCDESTROY:
        {
            RemoveWindowSubclass(hwnd, CompareOptionsWheelRouteSubclassProc, subclassId);
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}
void CompareDirectoriesWindow::EnsureOptionsControlsCreated(HWND dlg) noexcept
{
    if (! dlg || _optionsUi.host)
    {
        return;
    }

    const HINSTANCE instance = GetModuleHandleW(nullptr);

    _optionsUi.host = CreateWindowExW(WS_EX_CONTROLPARENT, L"Static", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP, 0, 0, 10, 10, dlg, nullptr, instance, nullptr);
    if (_optionsUi.host)
    {
        const wchar_t* hostTheme = _theme.highContrast ? L"" : (_theme.dark ? L"DarkMode_Explorer" : L"Explorer");
        SetWindowTheme(_optionsUi.host, hostTheme, nullptr);
        SendMessageW(_optionsUi.host, WM_THEMECHANGED, 0, 0);

#pragma warning(push)
#pragma warning(disable : 5039) // passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(_optionsUi.host, CompareOptionsHostSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(this));
#pragma warning(pop)
    }

    if (! _optionsUi.host)
    {
        return;
    }

    constexpr DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    constexpr DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;

    const DWORD toggleStyle = static_cast<DWORD>(WS_CHILD | WS_VISIBLE | WS_TABSTOP | (_theme.highContrast ? BS_AUTOCHECKBOX : BS_OWNERDRAW));

    const auto makeStatic = [&](DWORD style) noexcept -> HWND
    { return CreateWindowExW(0, L"Static", L"", style, 0, 0, 10, 10, _optionsUi.host, nullptr, instance, nullptr); };

    const auto makeToggle = [&](int id) noexcept -> HWND
    {
        const HWND toggle = CreateWindowExW(
            0, L"Button", L"", toggleStyle, 0, 0, 10, 10, _optionsUi.host, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), instance, nullptr);
        if (toggle && ! _theme.highContrast)
        {
            ThemedControls::EnableOwnerDrawButton(_optionsUi.host, id);
        }
        return toggle;
    };

    const auto makeFramedEdit = [&](HWND& outFrame, HWND& outEdit, int editId) noexcept
    {
        outFrame = nullptr;
        outEdit  = nullptr;

        const bool customFrames = ! _theme.highContrast;
        if (customFrames)
        {
            outFrame = CreateWindowExW(0, L"Static", L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, 0, 0, 10, 10, _optionsUi.host, nullptr, instance, nullptr);
        }

        DWORD editStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL;
        editStyle |= ES_MULTILINE;
        editStyle &= ~ES_WANTRETURN;

        const DWORD editExStyle = customFrames ? 0 : WS_EX_CLIENTEDGE;
        outEdit                 = CreateWindowExW(
            editExStyle, L"Edit", L"", editStyle, 0, 0, 10, 10, _optionsUi.host, reinterpret_cast<HMENU>(static_cast<INT_PTR>(editId)), instance, nullptr);

        if (outEdit)
        {
            const wchar_t* themeName = (_theme.highContrast || _theme.systemHighContrast) ? L"" : ((_theme.dark) ? L"DarkMode_Explorer" : L"Explorer");
            SetWindowTheme(outEdit, themeName, nullptr);
            SendMessageW(outEdit, WM_THEMECHANGED, 0, 0);
        }

        if (customFrames && outFrame && outEdit)
        {
            ThemedInputFrames::InstallFrame(outFrame, outEdit, &_optionsFrameStyle);
        }

        if (outEdit)
        {
            const UINT dpi       = GetDpiForWindow(outEdit);
            const int textMargin = ThemedControls::ScaleDip(dpi, 6);
            SendMessageW(outEdit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(textMargin, textMargin));
        }
    };

    _optionsUi.headerCompare  = makeStatic(baseStaticStyle);
    _optionsUi.headerSubdirs  = makeStatic(baseStaticStyle);
    _optionsUi.headerAdvanced = makeStatic(baseStaticStyle);
    _optionsUi.headerIgnore   = makeStatic(baseStaticStyle);

    _optionsUi.compareSize.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareSize.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareSize.toggle      = makeToggle(IDC_CMP_SIZE);

    _optionsUi.compareDateTime.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareDateTime.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareDateTime.toggle      = makeToggle(IDC_CMP_DATETIME);

    _optionsUi.compareAttributes.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareAttributes.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareAttributes.toggle      = makeToggle(IDC_CMP_ATTRIBUTES);

    _optionsUi.compareContent.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareContent.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareContent.toggle      = makeToggle(IDC_CMP_CONTENT);

    _optionsUi.compareSubdirectories.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareSubdirectories.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareSubdirectories.toggle      = makeToggle(IDC_CMP_SUBDIRECTORIES);

    _optionsUi.compareSubdirAttributes.title       = makeStatic(baseStaticStyle);
    _optionsUi.compareSubdirAttributes.description = makeStatic(wrapStaticStyle);
    _optionsUi.compareSubdirAttributes.toggle      = makeToggle(IDC_CMP_SUBDIR_ATTRIBUTES);

    _optionsUi.selectSubdirsOnlyInOnePane.title       = makeStatic(baseStaticStyle);
    _optionsUi.selectSubdirsOnlyInOnePane.description = makeStatic(wrapStaticStyle);
    _optionsUi.selectSubdirsOnlyInOnePane.toggle      = makeToggle(IDC_CMP_SELECT_SUBDIRS_ONLY_ONE_PANE);

    _optionsUi.keepIdenticalItems.title       = makeStatic(baseStaticStyle);
    _optionsUi.keepIdenticalItems.description = makeStatic(wrapStaticStyle);
    _optionsUi.keepIdenticalItems.toggle      = makeToggle(IDC_CMP_KEEP_IDENTICAL);

    _optionsUi.ignoreFiles.title       = makeStatic(baseStaticStyle);
    _optionsUi.ignoreFiles.description = makeStatic(wrapStaticStyle);
    _optionsUi.ignoreFiles.toggle      = makeToggle(IDC_CMP_IGNORE_FILES);
    makeFramedEdit(_optionsUi.ignoreFiles.frame, _optionsUi.ignoreFiles.edit, IDC_CMP_IGNORE_FILES_PATTERNS);

    _optionsUi.ignoreDirectories.title       = makeStatic(baseStaticStyle);
    _optionsUi.ignoreDirectories.description = makeStatic(wrapStaticStyle);
    _optionsUi.ignoreDirectories.toggle      = makeToggle(IDC_CMP_IGNORE_DIRECTORIES);
    makeFramedEdit(_optionsUi.ignoreDirectories.frame, _optionsUi.ignoreDirectories.edit, IDC_CMP_IGNORE_DIRECTORIES_PATTERNS);

#pragma warning(push)
#pragma warning(disable : 5039) // passing potentially-throwing callback to extern "C" Win32 API under -EHc
    SetWindowSubclass(dlg, CompareOptionsWheelRouteSubclassProc, 2u, reinterpret_cast<DWORD_PTR>(this));
    EnumChildWindows(dlg,
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto* self = reinterpret_cast<CompareDirectoriesWindow*>(lParam);
        if (! self)
        {
            return TRUE;
        }
        SetWindowSubclass(child, CompareOptionsWheelRouteSubclassProc, 2u, reinterpret_cast<DWORD_PTR>(self));
        return TRUE;
    },
                     reinterpret_cast<LPARAM>(this));
#pragma warning(pop)
}
void CompareDirectoriesWindow::PaintOptionsHostBackgroundAndCards(HDC hdc, HWND host) noexcept
{
    if (! hdc || ! host)
    {
        return;
    }

    RECT rc{};
    GetClientRect(host, &rc);

    if (_optionsBackgroundBrush)
    {
        FillRect(hdc, &rc, _optionsBackgroundBrush.get());
    }

    if (_theme.systemHighContrast || _theme.highContrast || _optionsCards.empty())
    {
        return;
    }

    const UINT dpi         = GetDpiForWindow(host);
    const int radius       = ThemedControls::ScaleDip(dpi, 6);
    const COLORREF surface = ThemedControls::GetControlSurfaceColor(_theme);
    const COLORREF border  = ThemedControls::BlendColor(surface, _theme.menu.text, _theme.dark ? 40 : 30, 255);

    wil::unique_hpen cardPen(CreatePen(PS_SOLID, 1, border));
    if (! _optionsCardBrush || ! cardPen)
    {
        return;
    }

    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc, _optionsCardBrush.get());
    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc, cardPen.get());

    for (const RECT& card : _optionsCards)
    {
        RoundRect(hdc, card.left, card.top, card.right, card.bottom, radius, radius);
    }

    if (_optionsUseTwoColumns && _optionsTwoColumnSeparatorX > rc.left && _optionsTwoColumnSeparatorX < rc.right)
    {
        const COLORREF separator = ThemedControls::BlendColor(surface, _theme.menu.text, _theme.dark ? 28 : 20, 255);
        wil::unique_hpen sepPen(CreatePen(PS_SOLID, 1, separator));
        if (sepPen)
        {
            [[maybe_unused]] auto oldSepPen = wil::SelectObject(hdc, sepPen.get());
            MoveToEx(hdc, _optionsTwoColumnSeparatorX, rc.top, nullptr);
            LineTo(hdc, _optionsTwoColumnSeparatorX, rc.bottom);
        }
    }
}
void CompareDirectoriesWindow::LayoutOptionsControls() noexcept
{
    if (! _optionsDlg || ! _optionsUi.host)
    {
        return;
    }

    EnsureCompareSession();
    const bool contentCompareSupported = ! _session || _session->IsContentCompareSupported();
    const UINT contentCompareDescId    = contentCompareSupported ? IDS_COMPARE_OPTIONS_CONTENT_DESC : IDS_COMPARE_OPTIONS_CONTENT_DESC_UNSUPPORTED;

    if (_optionsUi.compareContent.toggle)
    {
        EnableWindow(_optionsUi.compareContent.toggle, contentCompareSupported ? TRUE : FALSE);
    }

    RECT rcDlg{};
    if (! GetClientRect(_optionsDlg.get(), &rcDlg))
    {
        return;
    }

    const int dlgW = std::max(0l, rcDlg.right - rcDlg.left);
    const int dlgH = std::max(0l, rcDlg.bottom - rcDlg.top);

    const UINT dpi = GetDpiForWindow(_optionsDlg.get());

    const int margin       = ThemedControls::ScaleDip(dpi, 16);
    const int gapX         = ThemedControls::ScaleDip(dpi, 12);
    const int gapY         = ThemedControls::ScaleDip(dpi, 12);
    const int rowHeight    = std::max(1, ThemedControls::ScaleDip(dpi, 26));
    const int titleHeight  = std::max(1, ThemedControls::ScaleDip(dpi, 18));
    const int headerHeight = std::max(1, ThemedControls::ScaleDip(dpi, 20));

    const int cardPaddingX         = ThemedControls::ScaleDip(dpi, 12);
    const int cardPaddingY         = ThemedControls::ScaleDip(dpi, 8);
    const int cardGapY             = ThemedControls::ScaleDip(dpi, 2);
    const int cardGapX             = ThemedControls::ScaleDip(dpi, 12);
    const int cardSpacingY         = ThemedControls::ScaleDip(dpi, 8);
    const int sectionSpacing       = ThemedControls::ScaleDip(dpi, 16);
    const int framePadding         = ThemedControls::ScaleDip(dpi, 2);
    const int minToggleWidth       = ThemedControls::ScaleDip(dpi, 90);
    const int columnSeparatorAreaW = ThemedControls::ScaleDip(dpi, 28);
    const int minColumnW           = ThemedControls::ScaleDip(dpi, 360);

    const HFONT dialogFont = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    const HFONT headerFont = _uiBoldFont ? _uiBoldFont.get() : dialogFont;
    const HFONT infoFont   = _uiItalicFont ? _uiItalicFont.get() : dialogFont;

    const HWND okBtn     = GetDlgItem(_optionsDlg.get(), IDOK);
    const HWND cancelBtn = GetDlgItem(_optionsDlg.get(), IDCANCEL);

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
        const int textW         = ThemedControls::MeasureTextWidth(_optionsDlg.get(), dialogFont, text);
        return std::max(minBtnW, (2 * buttonPadX) + textW);
    };

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

    const int hostX = margin;
    const int hostY = margin;
    const int hostW = std::max(0, dlgW - 2 * margin);
    const int hostH = std::max(0, buttonsY - gapY - hostY);
    SetWindowPos(_optionsUi.host, nullptr, hostX, hostY, hostW, hostH, flags);

    RECT hostClient{};
    if (! GetClientRect(_optionsUi.host, &hostClient))
    {
        return;
    }

    const auto computeToggleWidth = [&](int contentW) noexcept -> int
    {
        const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
        const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

        const int onWidth  = ThemedControls::MeasureTextWidth(_optionsUi.host, headerFont, onLabel);
        const int offWidth = ThemedControls::MeasureTextWidth(_optionsUi.host, headerFont, offLabel);

        const int togglePaddingX = ThemedControls::ScaleDip(dpi, 6);
        const int toggleGapX     = ThemedControls::ScaleDip(dpi, 8);
        const int toggleTrackW   = ThemedControls::ScaleDip(dpi, 34);
        const int stateTextW     = std::max(onWidth, offWidth);

        const int measured = std::max(minToggleWidth, (2 * togglePaddingX) + stateTextW + toggleGapX + toggleTrackW);
        return std::min(std::max(0, contentW - 2 * cardPaddingX), measured);
    };

    auto computeToggleCardHeight = [&](int contentW, std::wstring_view descText, int toggleW) noexcept -> int
    {
        const int textW    = std::max(0, contentW - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH    = MeasureStaticTextHeight(_optionsUi.host, infoFont, textW, descText);
        const int contentH = std::max(0, titleHeight + cardGapY + descH);
        const int cardH    = std::max(rowHeight + 2 * cardPaddingY, contentH + 2 * cardPaddingY);
        return cardH;
    };

    auto computeIgnoreCardHeight = [&](int contentW, std::wstring_view descText, int toggleW, bool showEdit) noexcept -> int
    {
        const int textW = std::max(0, contentW - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH = MeasureStaticTextHeight(_optionsUi.host, infoFont, textW, descText);

        int contentH = std::max(0, titleHeight + cardGapY + descH);
        if (showEdit)
        {
            contentH += cardGapY + rowHeight;
        }
        const int cardH = std::max(rowHeight + 2 * cardPaddingY, contentH + 2 * cardPaddingY);
        return cardH;
    };

    auto computeContentHeight = [&](int contentW) noexcept -> int
    {
        const bool ignoreFilesOn = GetTwoStateToggleState(_optionsUi.ignoreFiles.toggle, _theme.highContrast);
        const bool ignoreDirsOn  = GetTwoStateToggleState(_optionsUi.ignoreDirectories.toggle, _theme.highContrast);

        const bool canTwoColumnLayout = (! _theme.systemHighContrast && ! _theme.highContrast) && (contentW >= (2 * minColumnW + columnSeparatorAreaW));
        if (! canTwoColumnLayout)
        {
            const int toggleW = computeToggleWidth(contentW);

            int y = 0;

            y += headerHeight + gapY;
            y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SUBDIRS_DESC), toggleW) + cardSpacingY;

            y += sectionSpacing;
            y += headerHeight + gapY;
            y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SIZE_DESC), toggleW) + cardSpacingY;
            y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_DATETIME_DESC), toggleW) + cardSpacingY;
            y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC), toggleW) + cardSpacingY;
            y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, contentCompareDescId), toggleW) + cardSpacingY;

            y += sectionSpacing;
            y += headerHeight + gapY;
            y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC), toggleW) + cardSpacingY;
            y += computeToggleCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC), toggleW) + cardSpacingY;

            y += sectionSpacing;
            y += headerHeight + gapY;
            y += computeIgnoreCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC), toggleW, ignoreFilesOn) +
                 cardSpacingY;
            y += computeIgnoreCardHeight(contentW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC), toggleW, ignoreDirsOn) +
                 cardSpacingY;

            return y;
        }

        const int leftW  = std::max(0, (contentW - columnSeparatorAreaW) / 2);
        const int rightX = leftW + columnSeparatorAreaW;
        const int rightW = std::max(0, contentW - rightX);

        const int toggleWLeft  = computeToggleWidth(leftW);
        const int toggleWRight = computeToggleWidth(rightW);

        int leftY = 0;
        leftY += headerHeight + gapY;
        leftY += computeToggleCardHeight(leftW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SUBDIRS_DESC), toggleWLeft) + cardSpacingY;

        leftY += sectionSpacing;
        leftY += headerHeight + gapY;
        leftY += computeToggleCardHeight(leftW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SIZE_DESC), toggleWLeft) + cardSpacingY;
        leftY += computeToggleCardHeight(leftW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_DATETIME_DESC), toggleWLeft) + cardSpacingY;
        leftY += computeToggleCardHeight(leftW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC), toggleWLeft) + cardSpacingY;
        leftY += computeToggleCardHeight(leftW, LoadStringResourceView(nullptr, contentCompareDescId), toggleWLeft) + cardSpacingY;

        int rightY = 0;
        rightY += headerHeight + gapY;
        rightY += computeToggleCardHeight(rightW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC), toggleWRight) + cardSpacingY;
        rightY += computeToggleCardHeight(rightW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC), toggleWRight) + cardSpacingY;

        rightY += sectionSpacing;
        rightY += headerHeight + gapY;
        rightY +=
            computeIgnoreCardHeight(rightW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC), toggleWRight, ignoreFilesOn) + cardSpacingY;
        rightY += computeIgnoreCardHeight(rightW, LoadStringResourceView(nullptr, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC), toggleWRight, ignoreDirsOn) +
                  cardSpacingY;

        return std::max(leftY, rightY);
    };

    const int viewportW = std::max(0l, hostClient.right - hostClient.left);
    const int viewportH = std::max(0l, hostClient.bottom - hostClient.top);

    int contentHeight       = computeContentHeight(viewportW);
    const bool wantsVScroll = viewportH > 0 && contentHeight > viewportH;

    LONG_PTR styleNow    = GetWindowLongPtrW(_optionsUi.host, GWL_STYLE);
    LONG_PTR styleWanted = styleNow;
    styleWanted |= WS_TABSTOP;
    styleWanted &= ~WS_HSCROLL;
    if (wantsVScroll)
    {
        styleWanted |= WS_VSCROLL;
    }
    else
    {
        styleWanted &= ~WS_VSCROLL;
    }

    const bool styleChanged = styleWanted != styleNow;
    if (styleChanged)
    {
        SetWindowLongPtrW(_optionsUi.host, GWL_STYLE, styleWanted);
        SetWindowPos(_optionsUi.host, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);

        const wchar_t* hostTheme = _theme.highContrast ? L"" : (_theme.dark ? L"DarkMode_Explorer" : L"Explorer");
        SetWindowTheme(_optionsUi.host, hostTheme, nullptr);
        SendMessageW(_optionsUi.host, WM_THEMECHANGED, 0, 0);
    }

    GetClientRect(_optionsUi.host, &hostClient);
    const int viewportW2 = std::max(0l, hostClient.right - hostClient.left);
    const int viewportH2 = std::max(0l, hostClient.bottom - hostClient.top);

    contentHeight        = computeContentHeight(viewportW2);
    _optionsScrollMax    = (viewportH2 > 0) ? std::max(0, contentHeight - viewportH2) : 0;
    _optionsScrollOffset = std::clamp(_optionsScrollOffset, 0, _optionsScrollMax);
    if (_optionsScrollMax <= 0)
    {
        _optionsScrollOffset = 0;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = std::max(0, contentHeight - 1);
    si.nPage  = (viewportH2 > 0) ? static_cast<UINT>(viewportH2) : 0u;
    si.nPos   = _optionsScrollOffset;
    SetScrollInfo(_optionsUi.host, SB_VERT, &si, TRUE);

    _optionsCards.clear();

    const int scrollOffset   = _optionsScrollOffset;
    const bool useTwoColumns = (! _theme.systemHighContrast && ! _theme.highContrast) && (viewportW2 >= (2 * minColumnW + columnSeparatorAreaW));

    _optionsUseTwoColumns       = useTwoColumns;
    _optionsTwoColumnSeparatorX = useTwoColumns ? (std::max(0, (viewportW2 - columnSeparatorAreaW) / 2) + (columnSeparatorAreaW / 2)) : -1;

    const auto positionScrollable = [&](HWND hwnd, int x, int y, int w, int h) noexcept
    {
        if (! hwnd)
        {
            return;
        }

        SetWindowPos(hwnd, nullptr, x, y - scrollOffset, w, h, flags);
    };

    auto pushCard = [&](int x, int top, int width, int height) noexcept
    {
        RECT card{};
        card.left   = x;
        card.top    = top - scrollOffset;
        card.right  = x + width;
        card.bottom = top + height - scrollOffset;
        _optionsCards.push_back(card);
    };

    const auto showToggleCardControls = [&](const OptionsToggleCard& card, bool visible) noexcept
    {
        const int cmd = visible ? SW_SHOW : SW_HIDE;
        ShowWindow(card.title, cmd);
        ShowWindow(card.description, cmd);
        ShowWindow(card.toggle, cmd);
    };

    const auto showIgnoreCardControls = [&](const OptionsIgnoreCard& card, bool visible, bool showEdit) noexcept
    {
        const int cmd = visible ? SW_SHOW : SW_HIDE;
        ShowWindow(card.title, cmd);
        ShowWindow(card.description, cmd);
        ShowWindow(card.toggle, cmd);
        if (card.frame)
        {
            ShowWindow(card.frame, (visible && showEdit) ? SW_SHOW : SW_HIDE);
        }
        if (card.edit)
        {
            ShowWindow(card.edit, (visible && showEdit) ? SW_SHOW : SW_HIDE);
        }
    };

    const auto layoutSectionHeader = [&](HWND header, UINT textId, int contentX, int contentW, int& y) noexcept
    {
        if (! header)
        {
            return;
        }

        const std::wstring text = LoadStringResource(nullptr, textId);
        SetWindowTextW(header, text.c_str());
        ShowWindow(header, SW_SHOW);
        positionScrollable(header, contentX + cardPaddingX, y, std::max(0, contentW - 2 * cardPaddingX), headerHeight);
        SendMessageW(header, WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    };

    const auto layoutToggleCard =
        [&](const OptionsToggleCard& card, UINT titleId, UINT descId, bool visible, int contentX, int contentW, int toggleW, int& y) noexcept
    {
        showToggleCardControls(card, visible);
        if (! visible)
        {
            return;
        }

        const std::wstring titleText = LoadStringResource(nullptr, titleId);
        const std::wstring descText  = LoadStringResource(nullptr, descId);

        const int textW = std::max(0, contentW - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH = MeasureStaticTextHeight(_optionsUi.host, infoFont, textW, descText);
        const int cardH = computeToggleCardHeight(contentW, descText, toggleW);

        pushCard(contentX, y, contentW, cardH);

        SetWindowTextW(card.title, titleText.c_str());
        positionScrollable(card.title, contentX + cardPaddingX, y + cardPaddingY, textW, titleHeight);
        SendMessageW(card.title, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        SetWindowTextW(card.description, descText.c_str());
        positionScrollable(card.description, contentX + cardPaddingX, y + cardPaddingY + titleHeight + cardGapY, textW, std::max(0, descH));
        SendMessageW(card.description, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);

        positionScrollable(card.toggle, contentX + contentW - cardPaddingX - toggleW, y + (cardH - rowHeight) / 2, toggleW, rowHeight);
        SendMessageW(card.toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        y += cardH + cardSpacingY;
    };

    const auto layoutIgnoreCard =
        [&](const OptionsIgnoreCard& card, UINT titleId, UINT descId, bool visible, bool showEdit, int contentX, int contentW, int toggleW, int& y) noexcept
    {
        showIgnoreCardControls(card, visible, showEdit);
        if (! visible)
        {
            return;
        }

        const std::wstring titleText = LoadStringResource(nullptr, titleId);
        const std::wstring descText  = LoadStringResource(nullptr, descId);

        const int textW = std::max(0, contentW - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH = MeasureStaticTextHeight(_optionsUi.host, infoFont, textW, descText);
        const int cardH = computeIgnoreCardHeight(contentW, descText, toggleW, showEdit);

        pushCard(contentX, y, contentW, cardH);

        SetWindowTextW(card.title, titleText.c_str());
        positionScrollable(card.title, contentX + cardPaddingX, y + cardPaddingY, textW, titleHeight);
        SendMessageW(card.title, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        SetWindowTextW(card.description, descText.c_str());
        positionScrollable(card.description, contentX + cardPaddingX, y + cardPaddingY + titleHeight + cardGapY, textW, std::max(0, descH));
        SendMessageW(card.description, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);

        positionScrollable(card.toggle, contentX + contentW - cardPaddingX - toggleW, y + cardPaddingY, toggleW, rowHeight);
        SendMessageW(card.toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        if (showEdit && card.frame && card.edit)
        {
            const int editX = contentX + cardPaddingX;
            const int editW = std::max(0, contentW - 2 * cardPaddingX);

            const int contentTop    = y + cardPaddingY;
            const int contentBottom = contentTop + titleHeight + cardGapY + descH;
            const int editTop       = contentBottom + cardGapY;

            const int innerPadding = (! _theme.highContrast && card.frame) ? framePadding : 0;

            positionScrollable(card.frame, editX, editTop, editW, rowHeight);
            positionScrollable(
                card.edit, editX + innerPadding, editTop + innerPadding, std::max(1, editW - 2 * innerPadding), std::max(1, rowHeight - 2 * innerPadding));
            SendMessageW(card.edit, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            ThemedControls::CenterEditTextVertically(card.edit);
        }

        y += cardH + cardSpacingY;
    };

    const bool ignoreFilesOn = GetTwoStateToggleState(_optionsUi.ignoreFiles.toggle, _theme.highContrast);
    const bool ignoreDirsOn  = GetTwoStateToggleState(_optionsUi.ignoreDirectories.toggle, _theme.highContrast);

    if (! useTwoColumns)
    {
        const int toggleW = computeToggleWidth(viewportW2);

        int y = 0;

        layoutSectionHeader(_optionsUi.headerSubdirs, IDS_COMPARE_OPTIONS_SECTION_SUBDIRS, 0, viewportW2, y);
        layoutToggleCard(
            _optionsUi.compareSubdirectories, IDS_COMPARE_OPTIONS_SUBDIRS_TITLE, IDS_COMPARE_OPTIONS_SUBDIRS_DESC, true, 0, viewportW2, toggleW, y);

        y += sectionSpacing;
        layoutSectionHeader(_optionsUi.headerCompare, IDS_COMPARE_OPTIONS_SECTION_COMPARE, 0, viewportW2, y);
        layoutToggleCard(_optionsUi.compareSize, IDS_COMPARE_OPTIONS_SIZE_TITLE, IDS_COMPARE_OPTIONS_SIZE_DESC, true, 0, viewportW2, toggleW, y);
        layoutToggleCard(_optionsUi.compareDateTime, IDS_COMPARE_OPTIONS_DATETIME_TITLE, IDS_COMPARE_OPTIONS_DATETIME_DESC, true, 0, viewportW2, toggleW, y);
        layoutToggleCard(
            _optionsUi.compareAttributes, IDS_COMPARE_OPTIONS_ATTRIBUTES_TITLE, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC, true, 0, viewportW2, toggleW, y);
        layoutToggleCard(_optionsUi.compareContent, IDS_COMPARE_OPTIONS_CONTENT_TITLE, contentCompareDescId, true, 0, viewportW2, toggleW, y);

        y += sectionSpacing;
        layoutSectionHeader(_optionsUi.headerAdvanced, IDS_COMPARE_OPTIONS_SECTION_ADVANCED, 0, viewportW2, y);
        layoutToggleCard(_optionsUi.compareSubdirAttributes,
                         IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_TITLE,
                         IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC,
                         true,
                         0,
                         viewportW2,
                         toggleW,
                         y);
        layoutToggleCard(_optionsUi.selectSubdirsOnlyInOnePane,
                         IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_TITLE,
                         IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC,
                         true,
                         0,
                         viewportW2,
                         toggleW,
                         y);
        layoutToggleCard(
            _optionsUi.keepIdenticalItems, IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_TITLE, IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_DESC, true, 0, viewportW2, toggleW, y);

        y += sectionSpacing;
        layoutSectionHeader(_optionsUi.headerIgnore, IDS_COMPARE_OPTIONS_SECTION_IGNORE, 0, viewportW2, y);
        layoutIgnoreCard(_optionsUi.ignoreFiles,
                         IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE,
                         IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC,
                         true,
                         ignoreFilesOn,
                         0,
                         viewportW2,
                         toggleW,
                         y);
        layoutIgnoreCard(_optionsUi.ignoreDirectories,
                         IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_TITLE,
                         IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC,
                         true,
                         ignoreDirsOn,
                         0,
                         viewportW2,
                         toggleW,
                         y);
    }
    else
    {
        const int leftW  = std::max(0, (viewportW2 - columnSeparatorAreaW) / 2);
        const int rightX = leftW + columnSeparatorAreaW;
        const int rightW = std::max(0, viewportW2 - rightX);

        const int toggleWLeft  = computeToggleWidth(leftW);
        const int toggleWRight = computeToggleWidth(rightW);

        int leftY = 0;
        layoutSectionHeader(_optionsUi.headerSubdirs, IDS_COMPARE_OPTIONS_SECTION_SUBDIRS, 0, leftW, leftY);
        layoutToggleCard(
            _optionsUi.compareSubdirectories, IDS_COMPARE_OPTIONS_SUBDIRS_TITLE, IDS_COMPARE_OPTIONS_SUBDIRS_DESC, true, 0, leftW, toggleWLeft, leftY);

        leftY += sectionSpacing;
        layoutSectionHeader(_optionsUi.headerCompare, IDS_COMPARE_OPTIONS_SECTION_COMPARE, 0, leftW, leftY);
        layoutToggleCard(_optionsUi.compareSize, IDS_COMPARE_OPTIONS_SIZE_TITLE, IDS_COMPARE_OPTIONS_SIZE_DESC, true, 0, leftW, toggleWLeft, leftY);
        layoutToggleCard(_optionsUi.compareDateTime, IDS_COMPARE_OPTIONS_DATETIME_TITLE, IDS_COMPARE_OPTIONS_DATETIME_DESC, true, 0, leftW, toggleWLeft, leftY);
        layoutToggleCard(
            _optionsUi.compareAttributes, IDS_COMPARE_OPTIONS_ATTRIBUTES_TITLE, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC, true, 0, leftW, toggleWLeft, leftY);
        layoutToggleCard(_optionsUi.compareContent, IDS_COMPARE_OPTIONS_CONTENT_TITLE, contentCompareDescId, true, 0, leftW, toggleWLeft, leftY);

        int rightY = 0;
        layoutSectionHeader(_optionsUi.headerAdvanced, IDS_COMPARE_OPTIONS_SECTION_ADVANCED, rightX, rightW, rightY);
        layoutToggleCard(_optionsUi.compareSubdirAttributes,
                         IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_TITLE,
                         IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC,
                         true,
                         rightX,
                         rightW,
                         toggleWRight,
                         rightY);
        layoutToggleCard(_optionsUi.selectSubdirsOnlyInOnePane,
                         IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_TITLE,
                         IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC,
                         true,
                         rightX,
                         rightW,
                         toggleWRight,
                         rightY);
        layoutToggleCard(_optionsUi.keepIdenticalItems,
                         IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_TITLE,
                         IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_DESC,
                         true,
                         rightX,
                         rightW,
                         toggleWRight,
                         rightY);

        rightY += sectionSpacing;
        layoutSectionHeader(_optionsUi.headerIgnore, IDS_COMPARE_OPTIONS_SECTION_IGNORE, rightX, rightW, rightY);
        layoutIgnoreCard(_optionsUi.ignoreFiles,
                         IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE,
                         IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC,
                         true,
                         ignoreFilesOn,
                         rightX,
                         rightW,
                         toggleWRight,
                         rightY);
        layoutIgnoreCard(_optionsUi.ignoreDirectories,
                         IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_TITLE,
                         IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC,
                         true,
                         ignoreDirsOn,
                         rightX,
                         rightW,
                         toggleWRight,
                         rightY);
    }

    InvalidateRect(_optionsUi.host, nullptr, TRUE);
}
void CompareDirectoriesWindow::ShowOptionsPanel(bool show) noexcept
{
    if (! _optionsDlg)
    {
        return;
    }

    ShowWindow(_optionsDlg.get(), show ? SW_SHOW : SW_HIDE);
    if (show)
    {
        LoadOptionsControlsFromSettings();
        if (const HWND fw = _folderWindow.GetHwnd())
        {
            ShowWindow(fw, SW_HIDE);
        }

        Layout();
        ApplyOptionsDialogTheme();
        SetWindowPos(_optionsDlg.get(), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        RedrawWindow(_optionsDlg.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
        SetFocus(_optionsDlg.get());
        UpdateCompareWatermark();
        UpdateViewMenuChecks();
        return;
    }

    if (_compareStarted)
    {
        if (const HWND fw = _folderWindow.GetHwnd())
        {
            ShowWindow(fw, SW_SHOW);
        }

        Layout();
        const HWND focus = GetFocus();
        if (! focus || (focus == _optionsDlg.get()) || IsChild(_optionsDlg.get(), focus))
        {
            SetFocus(_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left));
        }
    }

    UpdateCompareWatermark();
    UpdateViewMenuChecks();
}
Common::Settings::CompareDirectoriesSettings CompareDirectoriesWindow::GetEffectiveCompareSettings() const noexcept
{
    Common::Settings::CompareDirectoriesSettings s{};
    if (_settings && _settings->compareDirectories.has_value())
    {
        s = _settings->compareDirectories.value();
    }

    if (! s.keepIdenticalItems)
    {
        s.showIdenticalItems = false;
    }

    return s;
}
Common::Settings::CompareDirectoriesSettings CompareDirectoriesWindow::ReadOptionsControlsToSettings() const noexcept
{
    Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();
    if (! _optionsDlg || ! _optionsUi.host)
    {
        return s;
    }

    s.compareSize                   = GetTwoStateToggleState(_optionsUi.compareSize.toggle, _theme.highContrast);
    s.compareDateTime               = GetTwoStateToggleState(_optionsUi.compareDateTime.toggle, _theme.highContrast);
    s.compareAttributes             = GetTwoStateToggleState(_optionsUi.compareAttributes.toggle, _theme.highContrast);
    s.compareContent                = GetTwoStateToggleState(_optionsUi.compareContent.toggle, _theme.highContrast);
    s.compareSubdirectories         = GetTwoStateToggleState(_optionsUi.compareSubdirectories.toggle, _theme.highContrast);
    s.compareSubdirectoryAttributes = GetTwoStateToggleState(_optionsUi.compareSubdirAttributes.toggle, _theme.highContrast);
    s.selectSubdirsOnlyInOnePane    = GetTwoStateToggleState(_optionsUi.selectSubdirsOnlyInOnePane.toggle, _theme.highContrast);
    s.keepIdenticalItems            = GetTwoStateToggleState(_optionsUi.keepIdenticalItems.toggle, _theme.highContrast);
    s.ignoreFiles                   = GetTwoStateToggleState(_optionsUi.ignoreFiles.toggle, _theme.highContrast);
    s.ignoreDirectories             = GetTwoStateToggleState(_optionsUi.ignoreDirectories.toggle, _theme.highContrast);
    s.ignoreFilesPatterns = _optionsUi.ignoreFiles.edit ? Win32Text::GetDlgItemTextString(_optionsUi.host, IDC_CMP_IGNORE_FILES_PATTERNS) : std::wstring{};
    s.ignoreDirectoriesPatterns =
        _optionsUi.ignoreDirectories.edit ? Win32Text::GetDlgItemTextString(_optionsUi.host, IDC_CMP_IGNORE_DIRECTORIES_PATTERNS) : std::wstring{};

    if (! s.keepIdenticalItems)
    {
        s.showIdenticalItems = false;
    }

    return s;
}

bool CompareDirectoriesWindow::IsOptionsDialogDirty() const noexcept
{
    return ! AreEquivalentCompareDirectoriesSettings(ReadOptionsControlsToSettings(), GetEffectiveCompareSettings());
}

void CompareDirectoriesWindow::LoadOptionsControlsFromSettings() noexcept
{
    if (! _optionsDlg || ! _optionsUi.host)
    {
        return;
    }

    const Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();

    SetTwoStateToggleState(_optionsUi.compareSize.toggle, _theme.highContrast, s.compareSize);
    SetTwoStateToggleState(_optionsUi.compareDateTime.toggle, _theme.highContrast, s.compareDateTime);
    SetTwoStateToggleState(_optionsUi.compareAttributes.toggle, _theme.highContrast, s.compareAttributes);
    SetTwoStateToggleState(_optionsUi.compareContent.toggle, _theme.highContrast, s.compareContent);

    SetTwoStateToggleState(_optionsUi.compareSubdirectories.toggle, _theme.highContrast, s.compareSubdirectories);

    SetTwoStateToggleState(_optionsUi.compareSubdirAttributes.toggle, _theme.highContrast, s.compareSubdirectoryAttributes);
    SetTwoStateToggleState(_optionsUi.selectSubdirsOnlyInOnePane.toggle, _theme.highContrast, s.selectSubdirsOnlyInOnePane);
    SetTwoStateToggleState(_optionsUi.keepIdenticalItems.toggle, _theme.highContrast, s.keepIdenticalItems);

    SetTwoStateToggleState(_optionsUi.ignoreFiles.toggle, _theme.highContrast, s.ignoreFiles);
    SetTwoStateToggleState(_optionsUi.ignoreDirectories.toggle, _theme.highContrast, s.ignoreDirectories);
    if (_optionsUi.ignoreFiles.edit)
    {
        SetWindowTextW(_optionsUi.ignoreFiles.edit, s.ignoreFilesPatterns.c_str());
    }
    if (_optionsUi.ignoreDirectories.edit)
    {
        SetWindowTextW(_optionsUi.ignoreDirectories.edit, s.ignoreDirectoriesPatterns.c_str());
    }

    UpdateOptionsVisibility();
}
void CompareDirectoriesWindow::SaveOptionsControlsToSettings() noexcept
{
    if (! _optionsDlg || ! _settings || ! _optionsUi.host)
    {
        return;
    }

    _settings->compareDirectories   = ReadOptionsControlsToSettings();
    _optionsStaleFromExternalReload = false;
}

bool CompareDirectoriesWindow::ResolveOptionsStaleSaveConflict(HWND dlg) noexcept
{
    if (! _optionsStaleFromExternalReload)
    {
        return true;
    }

    SettingsHotReload::StaleSaveChoice choice = SettingsHotReload::StaleSaveChoice::Cancel;
    const HRESULT promptHr = SettingsHotReload::PromptStaleSaveConflict(dlg, LoadStringResource(nullptr, IDS_COMPARE_DIRECTORIES_TITLE), choice);
    if (FAILED(promptHr))
    {
        Debug::Warning(L"CompareDirectories: failed to prompt for stale save conflict (hr=0x{:08X})", static_cast<unsigned long>(promptHr));
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::ReloadFromDisk)
    {
        ReloadOptionsDialogFromDisk();
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::Cancel)
    {
        return false;
    }

    _optionsStaleFromExternalReload = false;
    return true;
}

void CompareDirectoriesWindow::ReloadOptionsDialogFromDisk() noexcept
{
    _optionsStaleFromExternalReload = false;
    LoadOptionsControlsFromSettings();
    UpdateViewMenuChecks();
}
void CompareDirectoriesWindow::UpdateOptionsVisibility() noexcept
{
    if (! _optionsDlg || ! _optionsUi.host)
    {
        return;
    }

    LayoutOptionsControls();
    RedrawWindow(_optionsUi.host, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
}

} // namespace CompareDirectoriesWindowInternal
