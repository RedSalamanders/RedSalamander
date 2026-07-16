#include "Framework.h"

#include "CompareDirectoriesWindow.Internal.h"
#include "D2DHdcPaint.h"
#include "DxUi/DxUi.Typography.h"
#include "DxUiThemePalette.h"
#include "SettingsHotReload.h"

namespace CompareDirectoriesWindowInternal
{
namespace
{
using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;
using RedSalamander::DxUi::WindowHost;
constexpr wchar_t kSettingsAppId[]                               = L"RedSalamander";
constexpr wchar_t kCompareOptionsHostOriginalWndProcProp[]       = L"RS.CompareOptionsHostOriginalWndProc";
constexpr wchar_t kCompareOptionsHostStateProp[]                 = L"RS.CompareOptionsHostState";
constexpr wchar_t kCompareOptionsWheelRouteOriginalWndProcProp[] = L"RS.CompareOptionsWheelRouteOriginalWndProc";
constexpr wchar_t kCompareOptionsWheelRouteStateProp[]           = L"RS.CompareOptionsWheelRouteState";
constexpr wchar_t kCompareOptionsDxHostOriginalWndProcProp[]     = L"RS.CompareOptionsDxHostOriginalWndProc";
constexpr wchar_t kCompareOptionsDxHostStateProp[]               = L"RS.CompareOptionsDxHostState";
constexpr wchar_t kCompareOptionsDxHostClassName[]               = L"RedSalamander.CompareOptions.DxHost";

[[nodiscard]] bool EnsureCompareOptionsDxHostClassRegistered() noexcept
{
    const HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSEXW existing{};
    existing.cbSize = sizeof(existing);
    if (GetClassInfoExW(instance, kCompareOptionsDxHostClassName, &existing) != 0)
    {
        return true;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = DefWindowProcW;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kCompareOptionsDxHostClassName;

    const ATOM atom = RegisterClassExW(&wc);
    return atom != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

[[nodiscard]] HWND CreateCompareOptionsDxHostWindow(HWND parent, DWORD style, DWORD exStyle = 0) noexcept
{
    if (! EnsureCompareOptionsDxHostClassRegistered())
    {
        return nullptr;
    }

    return CreateWindowExW(exStyle, kCompareOptionsDxHostClassName, L"", style, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr);
}

[[nodiscard]] std::wstring NormalizeDialogButtonCaption(std::wstring_view text)
{
    std::wstring normalized;
    normalized.reserve(text.size());

    for (size_t index = 0; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch != L'&')
        {
            normalized.push_back(ch);
            continue;
        }

        if ((index + 1u) < text.size() && text[index + 1u] == L'&')
        {
            normalized.push_back(L'&');
            ++index;
        }
    }

    return normalized;
}

[[nodiscard]] wchar_t ExtractDialogButtonMnemonic(std::wstring_view text) noexcept
{
    for (size_t index = 0; index < text.size(); ++index)
    {
        if (text[index] != L'&')
        {
            continue;
        }

        if ((index + 1u) >= text.size())
        {
            break;
        }

        const wchar_t candidate = text[index + 1u];
        if (candidate == L'&')
        {
            ++index;
            continue;
        }

        return candidate;
    }

    return L'\0';
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

struct CompareOptionsTypographyFormats final
{
    IDWriteFactory* dwriteFactory = nullptr;
    UINT dpi                      = USER_DEFAULT_SCREEN_DPI;
    wil::com_ptr<IDWriteTextFormat> bodyFormat;
    wil::com_ptr<IDWriteTextFormat> headerFormat;
    wil::com_ptr<IDWriteTextFormat> infoFormat;

    [[nodiscard]] bool IsReady() const noexcept
    {
        return dwriteFactory && bodyFormat && headerFormat && infoFormat;
    }
};

[[nodiscard]] float PixelExtentToDip(int valuePx, UINT dpi) noexcept
{
    return (static_cast<float>(std::max(0, valuePx)) * static_cast<float>(USER_DEFAULT_SCREEN_DPI)) /
           static_cast<float>(std::max<UINT>(dpi, USER_DEFAULT_SCREEN_DPI));
}

void ConfigureCompareOptionsTextFormat(IDWriteTextFormat* format, DWRITE_WORD_WRAPPING wrapping) noexcept
{
    if (! format)
    {
        return;
    }

    static_cast<void>(format->SetWordWrapping(wrapping));
    static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
    static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
}

[[nodiscard]] bool CreateCompareOptionsTextFormat(IDWriteFactory* factory,
                                                  const RedSalamander::DxUi::Typography::TypographySpec& spec,
                                                  DWRITE_FONT_STYLE style,
                                                  DWRITE_WORD_WRAPPING wrapping,
                                                  IDWriteTextFormat** outFormat) noexcept
{
    if (! factory || ! outFormat)
    {
        return false;
    }

    wchar_t preferredFamilyBuffer[RedSalamander::DxUi::Typography::kMaxDWriteFamilyNameLength + 1u]{};
    if (! RedSalamander::DxUi::Typography::CopyNullTerminated(spec.familyName, preferredFamilyBuffer))
    {
        return false;
    }

    const HRESULT hr =
        RedSalamander::DxUi::Typography::CreateTextFormatWithStyle(factory, preferredFamilyBuffer, spec.weight, style, spec.sizeDip, outFormat, L"");
    if (FAILED(hr) || ! *outFormat)
    {
        return false;
    }

    ConfigureCompareOptionsTextFormat(*outFormat, wrapping);
    return true;
}

[[nodiscard]] CompareOptionsTypographyFormats CreateCompareOptionsTypographyFormats(HWND referenceWindow) noexcept
{
    CompareOptionsTypographyFormats formats{};
    formats.dpi           = RedSalamander::DxUi::Typography::GetEffectiveDpi(referenceWindow);
    formats.dwriteFactory = RedSalamander::DxUi::Typography::GetSharedMeasurementFactory();
    if (! formats.dwriteFactory)
    {
        return formats;
    }

    const auto bodySpec   = RedSalamander::DxUi::Typography::MakeUiTextSpec(13.0f);
    const auto headerSpec = RedSalamander::DxUi::Typography::MakeUiTextSpec(13.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD);
    const auto infoSpec   = RedSalamander::DxUi::Typography::MakeUiTextSpec(13.0f);

    static_cast<void>(
        CreateCompareOptionsTextFormat(formats.dwriteFactory, bodySpec, DWRITE_FONT_STYLE_NORMAL, DWRITE_WORD_WRAPPING_NO_WRAP, formats.bodyFormat.put()));
    static_cast<void>(
        CreateCompareOptionsTextFormat(formats.dwriteFactory, headerSpec, DWRITE_FONT_STYLE_NORMAL, DWRITE_WORD_WRAPPING_NO_WRAP, formats.headerFormat.put()));
    static_cast<void>(
        CreateCompareOptionsTextFormat(formats.dwriteFactory, infoSpec, DWRITE_FONT_STYLE_ITALIC, DWRITE_WORD_WRAPPING_WRAP, formats.infoFormat.put()));

    return formats;
}

[[nodiscard]] int MeasureCompareOptionsTextWidth(const CompareOptionsTypographyFormats& formats, IDWriteTextFormat* format, std::wstring_view text) noexcept
{
    return RedSalamander::DxUi::Typography::MeasureSingleLineTextMetrics(formats.dwriteFactory, format, formats.dpi, text).widthPx;
}

[[nodiscard]] int MeasureCompareOptionsWrappedTextHeight(const CompareOptionsTypographyFormats& formats,
                                                         IDWriteTextFormat* format,
                                                         int widthPx,
                                                         std::wstring_view text) noexcept
{
    if (! formats.dwriteFactory || ! format || widthPx <= 0 || text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
    {
        return 0;
    }

    const float layoutWidthDip = std::max(1.0f, PixelExtentToDip(widthPx, formats.dpi));
    wil::com_ptr<IDWriteTextLayout> layout;
    const HRESULT hr = formats.dwriteFactory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, layoutWidthDip, 4096.0f, layout.put());
    if (FAILED(hr) || ! layout)
    {
        return 0;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return 0;
    }

    const int paddingY = UiMetrics::ScaleDip(formats.dpi, 6);
    return RedSalamander::DxUi::Typography::DipExtentToPixels(metrics.height, formats.dpi) + std::max(1, paddingY);
}

[[nodiscard]] bool MoveDialogTabFocusFromHost(HWND hostHwnd, bool reverse) noexcept
{
    const HWND dlg = GetAncestor(hostHwnd, GA_ROOT);
    if (! dlg)
    {
        return false;
    }

    HWND next = GetNextDlgTabItem(dlg, hostHwnd, reverse ? TRUE : FALSE);
    if (! next || next == hostHwnd)
    {
        return false;
    }

    SetFocus(next);
    return true;
}

[[nodiscard]] WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept
{
    return RedSalamander::Win32Callback::GetStoredWndProc(hwnd, propName);
}

[[nodiscard]] bool InstallWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp, WNDPROC hookWndProc) noexcept
{
    if (! hwnd || ! originalWndProcProp || ! hookWndProc)
    {
        return false;
    }

    if (GetStoredWndProc(hwnd, originalWndProcProp))
    {
        return true;
    }

    const auto originalWndProc = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(hwnd, GWLP_WNDPROC));
    if (! originalWndProc)
    {
        return false;
    }

    if (! RedSalamander::Win32Callback::SetPropNoThrow(hwnd, originalWndProcProp, reinterpret_cast<HANDLE>(originalWndProc)))
    {
        return false;
    }

    const auto previousWndProc =
        reinterpret_cast<WNDPROC>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(hookWndProc)));
    if (previousWndProc != originalWndProc)
    {
        RemovePropW(hwnd, originalWndProcProp);
        if (previousWndProc)
        {
            static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(previousWndProc)));
        }
        return false;
    }

    return true;
}

void RestoreWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp) noexcept
{
    if (! hwnd || ! originalWndProcProp)
    {
        return;
    }

    if (const auto originalWndProc = GetStoredWndProc(hwnd, originalWndProcProp))
    {
        RemovePropW(hwnd, originalWndProcProp);
        static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(originalWndProc)));
    }
}

[[nodiscard]] LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* originalWndProcProp, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (const auto originalWndProc = GetStoredWndProc(hwnd, originalWndProcProp))
    {
        return RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}
} // namespace

void CompareDirectoriesWindow::ApplyOptionsDialogTheme() noexcept
{
    if (! _optionsPanel.dlg)
    {
        return;
    }

    const bool darkBackground = ChooseContrastingTextColor(_theme.windowBackground) == RGB(255, 255, 255);
    const wchar_t* themeName  = _theme.highContrast ? L"" : (darkBackground ? L"DarkMode_Explorer" : L"Explorer");

    SetWindowTheme(_optionsPanel.dlg.get(), themeName, nullptr);
    SendMessageW(_optionsPanel.dlg.get(), WM_THEMECHANGED, 0, 0);

    struct EnumData
    {
        const wchar_t* themeName = nullptr;
        HWND optionsHost         = nullptr;
    };

    EnumData data{};
    data.themeName   = themeName;
    data.optionsHost = _optionsPanel.ui.host;

    EnumChildWindows(_optionsPanel.dlg.get(),
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto* data = reinterpret_cast<const EnumData*>(lParam);
        if (! data || ! child)
        {
            return TRUE;
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

    ApplyOptionsDxStaticTheme();
    ApplyOptionsDxButtonTheme();
    ApplyOptionsDxToggleTheme();
    ApplyOptionsDxEditTheme();
    SyncOptionsDxButtons();
    SyncOptionsDxToggles();
    SyncOptionsDxEdits();
    RedrawWindow(_optionsPanel.dlg.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
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
        case WM_SYSCHAR: return self->HandleOptionsDxMnemonic(static_cast<wchar_t>(wParam)) ? TRUE : FALSE;
        case WndMsg::kSettingsReloadedFromDisk: return self->OnOptionsSettingsReloadedFromDisk(dlg);
        case WM_CTLCOLOREDIT: return self->OnOptionsCtlColorEdit(reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_CTLCOLORDLG: return self->OnOptionsCtlColorDlg(reinterpret_cast<HDC>(wParam));
        case WM_CTLCOLORSTATIC: return self->OnOptionsCtlColorStatic(reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_CTLCOLORBTN: return self->OnOptionsCtlColorBtn(reinterpret_cast<HDC>(wParam), reinterpret_cast<HWND>(lParam));
        case WM_NCDESTROY: self->UnregisterOptionsReloadParticipant(); return FALSE;
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

    EnsureOptionsControlsCreated(dlg);
    return TRUE;
}

void CompareDirectoriesWindow::RegisterOptionsReloadParticipant() noexcept
{
    if (_optionsPanel.reloadParticipantRegistered || ! _optionsPanel.dlg || IsWindow(_optionsPanel.dlg.get()) == FALSE)
    {
        return;
    }

    SettingsHotReload::RegisterParticipant(_optionsPanel.dlg.get());
    _optionsPanel.reloadParticipantRegistered = true;
}

void CompareDirectoriesWindow::UnregisterOptionsReloadParticipant() noexcept
{
    if (! _optionsPanel.reloadParticipantRegistered)
    {
        return;
    }

    if (_optionsPanel.dlg)
    {
        SettingsHotReload::UnregisterParticipant(_optionsPanel.dlg.get());
    }
    _optionsPanel.reloadParticipantRegistered = false;
}

bool CompareDirectoriesWindow::EnsureOptionsDxStaticHosts() noexcept
{
    DetachOptionsDxStaticHosts();

    if (! _optionsPanel.ui.host)
    {
        return false;
    }

    auto dxState = std::make_unique<OptionsDxUiState>();

    wil::unique_hwnd hwnd(CreateCompareOptionsDxHostWindow(_optionsPanel.ui.host, WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP | SS_NOTIFY));
    if (! hwnd)
    {
        return false;
    }

    if (! RedSalamander::Win32Callback::SetPropNoThrow(hwnd.get(), kCompareOptionsDxHostStateProp, reinterpret_cast<HANDLE>(this)) ||
        ! InstallWndProcHook(hwnd.get(), kCompareOptionsDxHostOriginalWndProcProp, CompareOptionsDxHostWndProc) || ! dxState->body.host.Attach(hwnd.get()))
    {
        RemovePropW(hwnd.get(), kCompareOptionsDxHostStateProp);
        return false;
    }

    auto root  = std::make_unique<Panel>();
    auto& body = dxState->body;

    const auto makeHeader = [&](Label*& out) noexcept
    {
        out = root->AddChild<Label>();
        out->SetFontRole(FontRole::Header);
    };

    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);

    const auto makeToggleCard = [&](OptionsToggleCardDx& out, const UINT commandId) noexcept
    {
        out.card  = root->AddChild<CardPanel>();
        out.title = root->AddChild<Label>();
        out.title->SetFontRole(FontRole::Body);
        out.description = root->AddChild<Label>();
        out.description->SetFontRole(FontRole::Small);
        out.description->SetMultiline(true);
        out.toggle = root->AddChild<Toggle>();
        out.toggle->SetStateLabels(offLabel, onLabel);
        out.toggle->SetOnToggled([this, commandId](bool checked) noexcept { UpdateOptionsDraftToggle(commandId, checked); });
    };

    const auto makeIgnoreCard = [&](OptionsIgnoreCardDx& out, const UINT commandId, const UINT editId) noexcept
    {
        out.card  = root->AddChild<CardPanel>();
        out.title = root->AddChild<Label>();
        out.title->SetFontRole(FontRole::Body);
        out.description = root->AddChild<Label>();
        out.description->SetFontRole(FontRole::Small);
        out.description->SetMultiline(true);
        out.toggle = root->AddChild<Toggle>();
        out.toggle->SetStateLabels(offLabel, onLabel);
        out.toggle->SetOnToggled([this, commandId](bool checked) noexcept { UpdateOptionsDraftToggle(commandId, checked); });
        out.edit = root->AddChild<TextField>();
        out.title->SetMnemonicTarget(out.toggle);
        out.edit->SetOnTextChanged([this, editId](std::wstring_view text) noexcept { UpdateOptionsDraftText(editId, text); });
    };

    // Keep retained children in the same order as the visible dialog so keyboard Tab traversal
    // follows the on-screen flow instead of the raw member declaration order.
    makeHeader(body.headerSubdirs);
    makeToggleCard(body.compareSubdirectories, IDC_CMP_SUBDIRECTORIES);

    makeHeader(body.headerCompare);
    makeToggleCard(body.compareSize, IDC_CMP_SIZE);
    makeToggleCard(body.compareDateTime, IDC_CMP_DATETIME);
    makeToggleCard(body.compareAttributes, IDC_CMP_ATTRIBUTES);
    makeToggleCard(body.compareContent, IDC_CMP_CONTENT);

    makeHeader(body.headerAdvanced);
    makeToggleCard(body.compareSubdirAttributes, IDC_CMP_SUBDIR_ATTRIBUTES);
    makeToggleCard(body.selectSubdirsOnlyInOnePane, IDC_CMP_SELECT_SUBDIRS_ONLY_ONE_PANE);
    makeToggleCard(body.keepIdenticalItems, IDC_CMP_KEEP_IDENTICAL);

    makeHeader(body.headerIgnore);
    makeIgnoreCard(body.ignoreFiles, IDC_CMP_IGNORE_FILES, IDC_CMP_IGNORE_FILES_PATTERNS);
    makeIgnoreCard(body.ignoreDirectories, IDC_CMP_IGNORE_DIRECTORIES, IDC_CMP_IGNORE_DIRECTORIES_PATTERNS);
    body.ignoreFiles.title->SetMnemonic(L'F');
    body.ignoreDirectories.title->SetMnemonic(L'D');

    dxState->body.host.SetRoot(std::move(root));
    dxState->body.host.SetOnTabBoundary([this, hostHwnd = hwnd.get()](bool reverse) noexcept
    {
        if (! _optionsPanel.dxUi)
        {
            return MoveDialogTabFocusFromHost(hostHwnd, reverse);
        }

        const auto focusButtonHost = [](OptionsButtonDx& target) noexcept
        {
            if (! target.hostHwnd || IsWindow(target.hostHwnd.get()) == FALSE || ! target.button)
            {
                return false;
            }

            ::SetFocus(target.hostHwnd.get());
            target.host.SetFocusControl(target.button);
            return target.host.GetFocusControl() == target.button;
        };

        if (reverse)
        {
            return focusButtonHost(_optionsPanel.dxUi->cancelButton);
        }

        _optionsPanel.dxUi->body.lastFooterReturnTarget = _optionsPanel.dxUi->body.host.GetFocusControl();
        if (! _optionsPanel.dxUi->body.lastFooterReturnTarget)
        {
            _optionsPanel.dxUi->body.lastFooterReturnTarget = _optionsPanel.dxUi->body.compareSubdirectories.toggle;
        }
        return focusButtonHost(_optionsPanel.dxUi->okButton);
    });
    dxState->body.host.SetOnFocusChanged([this](RedSalamander::DxUi::Control* control) noexcept
    { static_cast<void>(EnsureOptionsDxBodyControlVisible(control)); });
    dxState->body.hostHwnd       = std::move(hwnd);
    dxState->usesDxUiStatics     = true;
    dxState->usesDxUiToggles     = true;
    dxState->usesDxUiEdits       = true;
    dxState->usesDxUiButtons     = false;
    _optionsPanel.dxUi           = std::move(dxState);
    _optionsPanel.syncingDxDraft = false;
    ApplyOptionsDxStaticTheme();
    SyncOptionsDxStatics();
    SyncOptionsDxToggles();
    SyncOptionsDxEdits();
    return true;
}

void CompareDirectoriesWindow::DetachOptionsDxStaticHosts() noexcept
{
    if (! _optionsPanel.dxUi)
    {
        return;
    }

    if (_optionsPanel.dxUi->body.hostHwnd)
    {
        RemovePropW(_optionsPanel.dxUi->body.hostHwnd.get(), kCompareOptionsDxHostStateProp);
        RestoreWndProcHook(_optionsPanel.dxUi->body.hostHwnd.get(), kCompareOptionsDxHostOriginalWndProcProp);
    }

    _optionsPanel.dxUi->Detach();
    _optionsPanel.dxUi.reset();
    _optionsPanel.syncingDxDraft = false;
}

bool CompareDirectoriesWindow::EnsureOptionsDxButtonHosts() noexcept
{
    const HWND dialog = _optionsPanel.dlg ? _optionsPanel.dlg.get() : (_optionsPanel.ui.host ? GetParent(_optionsPanel.ui.host) : nullptr);
    if (! _optionsPanel.dxUi || ! dialog)
    {
        return false;
    }

    auto attachButtonHost = [&](OptionsButtonDx& slot, const int commandId, const bool primary) noexcept
    {
        slot.Detach();

        const HWND legacyButton = GetDlgItem(dialog, commandId);
        if (! legacyButton)
        {
            slot.attachFailureStage = 1;
            return false;
        }

        wil::unique_hwnd hwnd(CreateCompareOptionsDxHostWindow(dialog, WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP | SS_NOTIFY));
        if (! hwnd)
        {
            slot.attachFailureStage = 2;
            return false;
        }

        if (! RedSalamander::Win32Callback::SetPropNoThrow(hwnd.get(), kCompareOptionsDxHostStateProp, reinterpret_cast<HANDLE>(this)))
        {
            slot.attachFailureStage = 3;
            return false;
        }

        if (! InstallWndProcHook(hwnd.get(), kCompareOptionsDxHostOriginalWndProcProp, CompareOptionsDxHostWndProc))
        {
            slot.attachFailureStage = 4;
            RemovePropW(hwnd.get(), kCompareOptionsDxHostStateProp);
            return false;
        }

        if (! slot.host.Attach(hwnd.get()))
        {
            slot.attachFailureStage = 5;
            RemovePropW(hwnd.get(), kCompareOptionsDxHostStateProp);
            RestoreWndProcHook(hwnd.get(), kCompareOptionsDxHostOriginalWndProcProp);
            return false;
        }

        auto root   = std::make_unique<Panel>();
        slot.button = root->AddChild<RedSalamander::DxUi::Button>();
        slot.button->SetPrimary(primary);
        slot.button->SetOnClick([this, legacyButton, commandId]() noexcept
        {
            const HWND dialogHwnd = _optionsPanel.dlg ? _optionsPanel.dlg.get() : (_optionsPanel.ui.host ? GetParent(_optionsPanel.ui.host) : nullptr);
            if (! dialogHwnd || ! legacyButton || IsWindow(dialogHwnd) == FALSE || IsWindow(legacyButton) == FALSE)
            {
                return;
            }

            PostMessageW(dialogHwnd, WM_COMMAND, MAKEWPARAM(commandId, BN_CLICKED), reinterpret_cast<LPARAM>(legacyButton));
        });
        slot.host.SetRoot(std::move(root));
        slot.host.SetOnTabBoundary([this, commandId, hostHwnd = hwnd.get()](bool reverse) noexcept
        {
            if (! _optionsPanel.dxUi)
            {
                return MoveDialogTabFocusFromHost(hostHwnd, reverse);
            }

            const auto focusButtonHost = [](OptionsButtonDx& target) noexcept
            {
                if (! target.hostHwnd || IsWindow(target.hostHwnd.get()) == FALSE || ! target.button)
                {
                    return false;
                }

                ::SetFocus(target.hostHwnd.get());
                target.host.SetFocusControl(target.button);
                return target.host.GetFocusControl() == target.button;
            };

            const auto focusBodyHost = [this](RedSalamander::DxUi::Control* preferredTarget) noexcept
            {
                if (! _optionsPanel.dxUi || ! _optionsPanel.dxUi->body.hostHwnd || IsWindow(_optionsPanel.dxUi->body.hostHwnd.get()) == FALSE)
                {
                    return false;
                }

                ::SetFocus(_optionsPanel.dxUi->body.hostHwnd.get());
                auto* target = preferredTarget;
                if (! target)
                {
                    target = _optionsPanel.dxUi->body.host.GetFocusControl();
                }
                if (! target)
                {
                    target = _optionsPanel.dxUi->body.compareSubdirectories.toggle;
                }
                if (! target)
                {
                    return false;
                }

                _optionsPanel.dxUi->body.host.SetFocusControl(target);
                static_cast<void>(EnsureOptionsDxBodyControlVisible(target));
                return _optionsPanel.dxUi->body.host.GetFocusControl() == target;
            };

            if (commandId == IDOK)
            {
                return reverse ? focusBodyHost(_optionsPanel.dxUi->body.lastFooterReturnTarget) : focusButtonHost(_optionsPanel.dxUi->cancelButton);
            }
            if (commandId == IDCANCEL)
            {
                return reverse ? focusButtonHost(_optionsPanel.dxUi->okButton) : focusBodyHost(_optionsPanel.dxUi->body.compareSubdirectories.toggle);
            }

            return MoveDialogTabFocusFromHost(hostHwnd, reverse);
        });
        slot.hostHwnd           = std::move(hwnd);
        slot.attachFailureStage = 0;
        return true;
    };

    const bool okAttached               = attachButtonHost(_optionsPanel.dxUi->okButton, IDOK, true);
    const bool cancelAttached           = attachButtonHost(_optionsPanel.dxUi->cancelButton, IDCANCEL, false);
    _optionsPanel.dxUi->usesDxUiButtons = okAttached && cancelAttached;
    if (_optionsPanel.dxUi->usesDxUiButtons)
    {
        ApplyOptionsDxButtonTheme();
        SyncOptionsDxButtons();
    }
    return _optionsPanel.dxUi->usesDxUiButtons;
}

void CompareDirectoriesWindow::DetachOptionsDxButtonHosts() noexcept
{
    if (_optionsPanel.dxUi)
    {
        _optionsPanel.dxUi->usesDxUiButtons = false;
        if (_optionsPanel.dxUi->okButton.hostHwnd)
        {
            RemovePropW(_optionsPanel.dxUi->okButton.hostHwnd.get(), kCompareOptionsDxHostStateProp);
            RestoreWndProcHook(_optionsPanel.dxUi->okButton.hostHwnd.get(), kCompareOptionsDxHostOriginalWndProcProp);
        }
        if (_optionsPanel.dxUi->cancelButton.hostHwnd)
        {
            RemovePropW(_optionsPanel.dxUi->cancelButton.hostHwnd.get(), kCompareOptionsDxHostStateProp);
            RestoreWndProcHook(_optionsPanel.dxUi->cancelButton.hostHwnd.get(), kCompareOptionsDxHostOriginalWndProcProp);
        }
        _optionsPanel.dxUi->okButton.Detach();
        _optionsPanel.dxUi->cancelButton.Detach();
    }
}

bool CompareDirectoriesWindow::EnsureOptionsDxToggleHosts() noexcept
{
    if (! _optionsPanel.dxUi)
    {
        return false;
    }

    _optionsPanel.dxUi->usesDxUiToggles = true;
    ApplyOptionsDxToggleTheme();
    SyncOptionsDxToggles();
    return true;
}

void CompareDirectoriesWindow::DetachOptionsDxToggleHosts() noexcept
{
    if (_optionsPanel.dxUi)
    {
        _optionsPanel.dxUi->usesDxUiToggles = false;
    }
}

void CompareDirectoriesWindow::ApplyOptionsDxStaticTheme() noexcept
{
    if (! _optionsPanel.dxUi || ! _optionsPanel.dxUi->usesDxUiStatics)
    {
        return;
    }

    _optionsPanel.dxUi->body.host.SetTheme(MakeAppThemeDxPalette(_theme));
    _optionsPanel.dxUi->body.host.Invalidate();
}

void CompareDirectoriesWindow::ApplyOptionsDxButtonTheme() noexcept
{
    if (! _optionsPanel.dxUi || ! _optionsPanel.dxUi->usesDxUiButtons)
    {
        return;
    }

    const ThemePalette palette = MakeAppThemeDxPalette(_theme);
    _optionsPanel.dxUi->okButton.host.SetTheme(palette);
    _optionsPanel.dxUi->cancelButton.host.SetTheme(palette);
    _optionsPanel.dxUi->okButton.host.Invalidate();
    _optionsPanel.dxUi->cancelButton.host.Invalidate();
}

void CompareDirectoriesWindow::ApplyOptionsDxToggleTheme() noexcept
{
    ApplyOptionsDxStaticTheme();
}

void CompareDirectoriesWindow::SyncOptionsDxStatics() noexcept
{
    if (! _optionsPanel.dxUi || ! _optionsPanel.dxUi->usesDxUiStatics)
    {
        return;
    }

    EnsureCompareSession();
    const bool contentCompareSupported = ! _session || _session->IsContentCompareSupported();
    const UINT contentCompareDescId    = contentCompareSupported ? IDS_COMPARE_OPTIONS_CONTENT_DESC : IDS_COMPARE_OPTIONS_CONTENT_DESC_UNSUPPORTED;

    auto& body = _optionsPanel.dxUi->body;

    const auto syncSection = [](Label* section, const UINT textId) noexcept
    {
        if (section)
        {
            section->SetText(LoadStringResource(nullptr, textId));
        }
    };

    const auto syncCard = [](OptionsToggleCardDx& card, const UINT titleId, const UINT descId) noexcept
    {
        if (card.title)
        {
            card.title->SetText(LoadStringResource(nullptr, titleId));
        }
        if (card.description)
        {
            card.description->SetText(LoadStringResource(nullptr, descId));
        }
    };

    const auto syncIgnoreCard = [](OptionsIgnoreCardDx& card, const UINT titleId, const UINT descId) noexcept
    {
        if (card.title)
        {
            card.title->SetText(LoadStringResource(nullptr, titleId));
        }
        if (card.description)
        {
            card.description->SetText(LoadStringResource(nullptr, descId));
        }
    };

    syncSection(body.headerCompare, IDS_COMPARE_OPTIONS_SECTION_COMPARE);
    syncSection(body.headerSubdirs, IDS_COMPARE_OPTIONS_SECTION_SUBDIRS);
    syncSection(body.headerAdvanced, IDS_COMPARE_OPTIONS_SECTION_ADVANCED);
    syncSection(body.headerIgnore, IDS_COMPARE_OPTIONS_SECTION_IGNORE);

    syncCard(body.compareSize, IDS_COMPARE_OPTIONS_SIZE_TITLE, IDS_COMPARE_OPTIONS_SIZE_DESC);
    syncCard(body.compareDateTime, IDS_COMPARE_OPTIONS_DATETIME_TITLE, IDS_COMPARE_OPTIONS_DATETIME_DESC);
    syncCard(body.compareAttributes, IDS_COMPARE_OPTIONS_ATTRIBUTES_TITLE, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC);
    syncCard(body.compareContent, IDS_COMPARE_OPTIONS_CONTENT_TITLE, contentCompareDescId);
    syncCard(body.compareSubdirectories, IDS_COMPARE_OPTIONS_SUBDIRS_TITLE, IDS_COMPARE_OPTIONS_SUBDIRS_DESC);
    syncCard(body.compareSubdirAttributes, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_TITLE, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC);
    syncCard(body.selectSubdirsOnlyInOnePane, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_TITLE, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC);
    syncCard(body.keepIdenticalItems, IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_TITLE, IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_DESC);
    syncIgnoreCard(body.ignoreFiles, IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE, IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC);
    syncIgnoreCard(body.ignoreDirectories, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_TITLE, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC);
    body.host.Invalidate();
}

void CompareDirectoriesWindow::SyncOptionsDxButtons() noexcept
{
    if (! _optionsPanel.dxUi || ! _optionsPanel.dxUi->usesDxUiButtons || ! _optionsPanel.dlg)
    {
        return;
    }

    const auto syncButton = [&](OptionsButtonDx& slot, const int commandId, const bool primary) noexcept
    {
        const HWND legacyButton = GetDlgItem(_optionsPanel.dlg.get(), commandId);
        if (! slot.button || ! legacyButton)
        {
            return;
        }

        const std::wstring legacyCaption = Win32Text::GetWindowTextString(legacyButton);
        slot.button->SetPrimary(primary);
        slot.button->SetText(NormalizeDialogButtonCaption(legacyCaption));
        const wchar_t mnemonic = commandId == IDOK ? L'O' : (commandId == IDCANCEL ? L'C' : ExtractDialogButtonMnemonic(legacyCaption));
        slot.button->SetMnemonic(mnemonic);
        slot.button->SetEnabled(IsWindowEnabled(legacyButton) != FALSE);
        slot.host.Invalidate();
    };

    syncButton(_optionsPanel.dxUi->okButton, IDOK, true);
    syncButton(_optionsPanel.dxUi->cancelButton, IDCANCEL, false);
}

void CompareDirectoriesWindow::SyncOptionsDxToggles() noexcept
{
    if (! _optionsPanel.dxUi || ! _optionsPanel.dxUi->usesDxUiToggles)
    {
        return;
    }

    EnsureCompareSession();
    const bool contentCompareSupported = ! _session || _session->IsContentCompareSupported();

    const auto syncToggle = [](Toggle* toggle, const bool checked, const bool enabled = true) noexcept
    {
        if (! toggle)
        {
            return;
        }

        toggle->SetChecked(checked);
        toggle->SetEnabled(enabled);
    };

    auto& body = _optionsPanel.dxUi->body;
    syncToggle(body.compareSize.toggle, _optionsPanel.draft.compareSize);
    syncToggle(body.compareDateTime.toggle, _optionsPanel.draft.compareDateTime);
    syncToggle(body.compareAttributes.toggle, _optionsPanel.draft.compareAttributes);
    syncToggle(body.compareContent.toggle, contentCompareSupported && _optionsPanel.draft.compareContent, contentCompareSupported);
    syncToggle(body.compareSubdirectories.toggle, _optionsPanel.draft.compareSubdirectories);
    syncToggle(body.compareSubdirAttributes.toggle, _optionsPanel.draft.compareSubdirectoryAttributes);
    syncToggle(body.selectSubdirsOnlyInOnePane.toggle, _optionsPanel.draft.selectSubdirsOnlyInOnePane);
    syncToggle(body.keepIdenticalItems.toggle, _optionsPanel.draft.keepIdenticalItems);
    syncToggle(body.ignoreFiles.toggle, _optionsPanel.draft.ignoreFiles);
    syncToggle(body.ignoreDirectories.toggle, _optionsPanel.draft.ignoreDirectories);
    body.host.Invalidate();
}

bool CompareDirectoriesWindow::EnsureOptionsDxEditHosts() noexcept
{
    if (! _optionsPanel.dxUi)
    {
        return false;
    }

    _optionsPanel.dxUi->usesDxUiEdits = true;
    ApplyOptionsDxEditTheme();
    SyncOptionsDxEdits();
    return true;
}

void CompareDirectoriesWindow::DetachOptionsDxEditHosts() noexcept
{
    if (_optionsPanel.dxUi)
    {
        _optionsPanel.dxUi->usesDxUiEdits = false;
    }
}

void CompareDirectoriesWindow::ApplyOptionsDxEditTheme() noexcept
{
    ApplyOptionsDxStaticTheme();
}

void CompareDirectoriesWindow::SyncOptionsDxEdits() noexcept
{
    if (! _optionsPanel.dxUi || ! _optionsPanel.dxUi->usesDxUiEdits)
    {
        return;
    }

    const auto syncEdit = [&](TextField* edit, const std::wstring& text, const bool enabled = true) noexcept
    {
        if (! edit)
        {
            return;
        }

        _optionsPanel.syncingDxDraft = true;
        edit->SetText(text);
        edit->SetEnabled(enabled);
        _optionsPanel.syncingDxDraft = false;
    };

    auto& body = _optionsPanel.dxUi->body;
    syncEdit(body.ignoreFiles.edit, _optionsPanel.draft.ignoreFilesPatterns);
    syncEdit(body.ignoreDirectories.edit, _optionsPanel.draft.ignoreDirectoriesPatterns);
    body.host.Invalidate();
}

bool CompareDirectoriesWindow::HandleOptionsDxMnemonic(const wchar_t mnemonic) noexcept
{
    if (! _optionsPanel.dxUi)
    {
        return false;
    }

    const auto focusBodyControl = [&](RedSalamander::DxUi::Control* control) noexcept
    {
        if (! control || ! control->IsVisible() || ! control->IsEnabled())
        {
            return false;
        }

        if (_optionsPanel.dxUi->body.hostHwnd && IsWindow(_optionsPanel.dxUi->body.hostHwnd.get()) != FALSE)
        {
            ::SetFocus(_optionsPanel.dxUi->body.hostHwnd.get());
        }

        _optionsPanel.dxUi->body.host.SetFocusControl(control);
        static_cast<void>(EnsureOptionsDxBodyControlVisible(control));
        _optionsPanel.dxUi->body.host.Invalidate();
        return _optionsPanel.dxUi->body.host.GetFocusControl() == control;
    };

    switch (std::towlower(mnemonic))
    {
        case L'f':
        {
            auto* const target = _optionsPanel.dxUi->body.ignoreFiles.edit && _optionsPanel.dxUi->body.ignoreFiles.edit->IsVisible()
                                     ? static_cast<RedSalamander::DxUi::Control*>(_optionsPanel.dxUi->body.ignoreFiles.edit)
                                     : static_cast<RedSalamander::DxUi::Control*>(_optionsPanel.dxUi->body.ignoreFiles.toggle);
            return focusBodyControl(target);
        }
        case L'd':
        {
            auto* const target = _optionsPanel.dxUi->body.ignoreDirectories.edit && _optionsPanel.dxUi->body.ignoreDirectories.edit->IsVisible()
                                     ? static_cast<RedSalamander::DxUi::Control*>(_optionsPanel.dxUi->body.ignoreDirectories.edit)
                                     : static_cast<RedSalamander::DxUi::Control*>(_optionsPanel.dxUi->body.ignoreDirectories.toggle);
            return focusBodyControl(target);
        }
        default: break;
    }

    const auto tryHost = [mnemonic](auto& slot) noexcept
    { return slot.hostHwnd && IsWindowVisible(slot.hostHwnd.get()) != FALSE && slot.host.HandleMnemonic(mnemonic); };

    return tryHost(_optionsPanel.dxUi->body) || tryHost(_optionsPanel.dxUi->okButton) || tryHost(_optionsPanel.dxUi->cancelButton);
}

bool CompareDirectoriesWindow::EnsureOptionsDxBodyControlVisible(RedSalamander::DxUi::Control* control) noexcept
{
    if (! _optionsPanel.dxUi || ! _optionsPanel.dxUi->body.hostHwnd || ! _optionsPanel.ui.host || ! control || ! control->IsVisible())
    {
        return false;
    }

    const HWND bodyHwnd            = _optionsPanel.dxUi->body.hostHwnd.get();
    const HWND focusedBeforeLayout = GetFocus();
    const bool restoreBodyFocus =
        _optionsPanel.dxUi->body.host.GetFocusControl() == control && (focusedBeforeLayout == bodyHwnd || IsChild(bodyHwnd, focusedBeforeLayout) != FALSE);

    RECT hostClient{};
    if (! GetClientRect(_optionsPanel.ui.host, &hostClient))
    {
        return false;
    }

    const UINT dpi           = GetDpiForWindow(_optionsPanel.ui.host);
    const int paddingY       = UiMetrics::ScaleDip(dpi, 8);
    const int viewportTop    = hostClient.top + paddingY;
    const int viewportBottom = (std::max)(viewportTop, static_cast<int>(hostClient.bottom) - paddingY);
    const auto bounds        = control->GetBounds();
    const int controlTop     = static_cast<int>(std::lround(_optionsPanel.dxUi->body.host.DipsToPixels(bounds.top)));
    const int controlBottom  = static_cast<int>(std::lround(_optionsPanel.dxUi->body.host.DipsToPixels(bounds.bottom)));

    int newOffset = _optionsPanel.scrollOffset;
    if (controlTop < viewportTop)
    {
        newOffset += controlTop - viewportTop;
    }
    else if (controlBottom > viewportBottom)
    {
        newOffset += controlBottom - viewportBottom;
    }

    newOffset = std::clamp(newOffset, 0, _optionsPanel.scrollMax);
    if (newOffset == _optionsPanel.scrollOffset)
    {
        return false;
    }

    _optionsPanel.scrollOffset = newOffset;
    LayoutOptionsControls();
    if (restoreBodyFocus && _optionsPanel.dxUi && _optionsPanel.dxUi->body.hostHwnd && IsWindow(_optionsPanel.dxUi->body.hostHwnd.get()) != FALSE)
    {
        const HWND currentFocus = GetFocus();
        if (currentFocus != _optionsPanel.dxUi->body.hostHwnd.get() && IsChild(_optionsPanel.dxUi->body.hostHwnd.get(), currentFocus) == FALSE)
        {
            ::SetFocus(_optionsPanel.dxUi->body.hostHwnd.get());
        }
        if (_optionsPanel.dxUi->body.host.GetFocusControl() != control)
        {
            _optionsPanel.dxUi->body.host.SetFocusControl(control);
        }
    }
    return true;
}

LRESULT CompareDirectoriesWindow::HandleOptionsDxHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept
{
    if (! _optionsPanel.dxUi)
    {
        handled = false;
        return 0;
    }

    if (hwnd == _optionsPanel.dxUi->body.hostHwnd.get() && msg == WM_KEYDOWN && _optionsPanel.dlg && IsWindow(_optionsPanel.dlg.get()) != FALSE)
    {
        const bool altDown  = (GetKeyState(VK_MENU) & 0x8000) != 0;
        const bool ctrlDown = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        if (! altDown && ! ctrlDown)
        {
            if (wp == VK_RETURN)
            {
                if (dynamic_cast<RedSalamander::DxUi::TextField*>(_optionsPanel.dxUi->body.host.GetFocusControl()) != nullptr)
                {
                    handled = true;
                    PostMessageW(
                        _optionsPanel.dlg.get(), WM_COMMAND, MAKEWPARAM(IDOK, BN_CLICKED), reinterpret_cast<LPARAM>(GetDlgItem(_optionsPanel.dlg.get(), IDOK)));
                    return 0;
                }
            }
            else if (wp == VK_ESCAPE)
            {
                handled = true;
                PostMessageW(_optionsPanel.dlg.get(),
                             WM_COMMAND,
                             MAKEWPARAM(IDCANCEL, BN_CLICKED),
                             reinterpret_cast<LPARAM>(GetDlgItem(_optionsPanel.dlg.get(), IDCANCEL)));
                return 0;
            }
        }
    }

    const auto dispatch = [&](wil::unique_hwnd& expectedHwnd, WindowHost& host) noexcept -> std::optional<LRESULT>
    {
        if (hwnd != expectedHwnd.get() || ! expectedHwnd)
        {
            return std::nullopt;
        }

        if (msg == WM_NCDESTROY)
        {
            handled = true;
            host.ReleaseMouseCapture();
            expectedHwnd.release();
            return 0;
        }
        return host.HandleMessage(hwnd, msg, wp, lp, handled);
    };

    if (const auto result = dispatch(_optionsPanel.dxUi->body.hostHwnd, _optionsPanel.dxUi->body.host))
    {
        if (handled && (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN || msg == WM_CHAR))
        {
            static_cast<void>(EnsureOptionsDxBodyControlVisible(_optionsPanel.dxUi->body.host.GetFocusControl()));
        }
        return result.value();
    }
    if (const auto result = dispatch(_optionsPanel.dxUi->okButton.hostHwnd, _optionsPanel.dxUi->okButton.host))
    {
        return result.value();
    }
    if (const auto result = dispatch(_optionsPanel.dxUi->cancelButton.hostHwnd, _optionsPanel.dxUi->cancelButton.host))
    {
        return result.value();
    }

    handled = false;
    return 0;
}

INT_PTR CompareDirectoriesWindow::OnOptionsEraseBkgnd(HWND dlg, HDC hdc) noexcept
{
    if (! _optionsPanel.backgroundBrush)
    {
        return FALSE;
    }

    RECT rc{};
    GetClientRect(dlg, &rc);
    FillRect(hdc, &rc, _optionsPanel.backgroundBrush.get());
    return TRUE;
}

INT_PTR CompareDirectoriesWindow::OnOptionsSettingsReloadedFromDisk(HWND dlg) noexcept
{
    if (! dlg || IsWindowVisible(dlg) == FALSE)
    {
        ReloadOptionsDialogFromDisk();
        return TRUE;
    }

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
            _optionsPanel.staleFromExternalReload = true;
            return TRUE;
        }
    }

    ReloadOptionsDialogFromDisk();
    return TRUE;
}

INT_PTR CompareDirectoriesWindow::OnOptionsCommand([[maybe_unused]] HWND dlg, WPARAM wParam, [[maybe_unused]] LPARAM lParam) noexcept
{
    const UINT controlId = LOWORD(wParam);

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
        default: break;
    }

    return FALSE;
}

INT_PTR CompareDirectoriesWindow::OnOptionsCtlColorEdit(HDC hdc, HWND control) noexcept
{
    if (! _optionsPanel.inputBrush)
    {
        return FALSE;
    }

    const bool enabled = ! control || IsWindowEnabled(control) != FALSE;
    const bool focused = enabled && control && GetFocus() == control;
    const COLORREF bg =
        enabled ? (focused ? _optionsPanel.inputFocusedBackgroundColor : _optionsPanel.inputBackgroundColor) : _optionsPanel.inputDisabledBackgroundColor;

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, bg);
    SetTextColor(hdc, enabled ? _theme.menu.text : _theme.menu.disabledText);

    if (_theme.highContrast)
    {
        return reinterpret_cast<INT_PTR>(_optionsPanel.backgroundBrush.get());
    }

    if (! enabled)
    {
        return reinterpret_cast<INT_PTR>(_optionsPanel.inputDisabledBrush.get());
    }

    return reinterpret_cast<INT_PTR>(focused && _optionsPanel.inputFocusedBrush ? _optionsPanel.inputFocusedBrush.get() : _optionsPanel.inputBrush.get());
}

INT_PTR CompareDirectoriesWindow::OnOptionsCtlColorDlg(HDC hdc) noexcept
{
    if (! _optionsPanel.backgroundBrush)
    {
        return FALSE;
    }

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, _theme.windowBackground);
    SetTextColor(hdc, _theme.menu.text);
    return reinterpret_cast<INT_PTR>(_optionsPanel.backgroundBrush.get());
}

INT_PTR CompareDirectoriesWindow::OnOptionsCtlColorStatic(HDC hdc, HWND control) noexcept
{
    if (! _optionsPanel.backgroundBrush)
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
        return reinterpret_cast<INT_PTR>(_optionsPanel.backgroundBrush.get());
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    SetBkColor(hdc, _theme.windowBackground);

    HBRUSH brush = _optionsPanel.backgroundBrush.get();
    if (control && _optionsPanel.ui.host && _optionsPanel.cardBrush && ! _optionsPanel.cards.empty())
    {
        RECT rc{};
        if (GetWindowRect(control, &rc) != FALSE)
        {
            MapWindowPoints(nullptr, _optionsPanel.ui.host, reinterpret_cast<POINT*>(&rc), 2);
            for (const RECT& card : _optionsPanel.cards)
            {
                RECT intersect{};
                if (IntersectRect(&intersect, &card, &rc) != FALSE)
                {
                    brush = _optionsPanel.cardBrush.get();
                    break;
                }
            }
        }
    }

    return reinterpret_cast<INT_PTR>(brush);
}

INT_PTR CompareDirectoriesWindow::OnOptionsCtlColorBtn(HDC hdc, HWND control) noexcept
{
    if (! _optionsPanel.backgroundBrush)
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
    return reinterpret_cast<INT_PTR>(_optionsPanel.backgroundBrush.get());
}

LRESULT CALLBACK CompareOptionsHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self     = reinterpret_cast<CompareDirectoriesWindow*>(GetPropW(hwnd, kCompareOptionsHostStateProp));
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

            if (self)
            {
                self->PaintOptionsHostBackgroundAndCards(hdc.get(), hwnd);
            }
            return 0;
        }
        case WM_VSCROLL:
        {
            if (! self || self->_optionsPanel.scrollMax <= 0)
            {
                break;
            }

            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_ALL;
            GetScrollInfo(hwnd, SB_VERT, &si);

            const UINT dpi  = GetDpiForWindow(hwnd);
            const int lineY = UiMetrics::ScaleDip(dpi, 24);

            int newPos = self->_optionsPanel.scrollOffset;
            switch (LOWORD(wp))
            {
                case SB_TOP: newPos = 0; break;
                case SB_BOTTOM: newPos = self->_optionsPanel.scrollMax; break;
                case SB_LINEUP: newPos -= lineY; break;
                case SB_LINEDOWN: newPos += lineY; break;
                case SB_PAGEUP: newPos -= static_cast<int>(si.nPage); break;
                case SB_PAGEDOWN: newPos += static_cast<int>(si.nPage); break;
                case SB_THUMBTRACK: newPos = si.nTrackPos; break;
                case SB_THUMBPOSITION: newPos = si.nPos; break;
                default: break;
            }

            newPos = std::clamp(newPos, 0, self->_optionsPanel.scrollMax);
            if (newPos != self->_optionsPanel.scrollOffset)
            {
                self->_optionsPanel.scrollOffset = newPos;
                self->LayoutOptionsControls();
            }
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            if (! self || self->_optionsPanel.scrollMax <= 0)
            {
                if (self)
                {
                    self->_optionsPanel.wheelRemainder = 0;
                }
                break;
            }

            const int delta = GET_WHEEL_DELTA_WPARAM(wp);
            if (delta == 0)
            {
                return 0;
            }

            self->_optionsPanel.wheelRemainder += delta;
            const int notches = self->_optionsPanel.wheelRemainder / WHEEL_DELTA;
            self->_optionsPanel.wheelRemainder -= notches * WHEEL_DELTA;
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
            const int lineY = UiMetrics::ScaleDip(dpi, 32);

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

            const int newPos = std::clamp(self->_optionsPanel.scrollOffset - scrollDelta, 0, self->_optionsPanel.scrollMax);
            if (newPos != self->_optionsPanel.scrollOffset)
            {
                self->_optionsPanel.scrollOffset = newPos;
                self->LayoutOptionsControls();
            }
            return 0;
        }
        case WM_COMMAND:
        case WM_NOTIFY:
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
            const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kCompareOptionsHostOriginalWndProcProp);
            RemovePropW(hwnd, kCompareOptionsHostStateProp);
            RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kCompareOptionsHostOriginalWndProcProp, CompareOptionsHostWndProc);
            return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp)
                                   : DefWindowProcW(hwnd, msg, wp, lp);
        }
    }

    return CallStoredWndProc(hwnd, kCompareOptionsHostOriginalWndProcProp, msg, wp, lp);
}
LRESULT CALLBACK CompareOptionsWheelRouteWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(GetPropW(hwnd, kCompareOptionsWheelRouteStateProp));
    if (! self)
    {
        return CallStoredWndProc(hwnd, kCompareOptionsWheelRouteOriginalWndProcProp, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_MOUSEWHEEL:
        {
            if (! self->_optionsPanel.dlg || ! self->_optionsPanel.ui.host)
            {
                return 0;
            }

            static thread_local bool s_routingWheel = false;
            if (s_routingWheel && hwnd != self->_optionsPanel.ui.host)
            {
                return 0;
            }

            POINT ptScreen{};
            ptScreen.x = GET_X_LPARAM(lp);
            ptScreen.y = GET_Y_LPARAM(lp);

            RECT dlgRect{};
            if (GetWindowRect(self->_optionsPanel.dlg.get(), &dlgRect) == 0 || PtInRect(&dlgRect, ptScreen) == FALSE)
            {
                // Don't scroll the options dialog when the user is wheeling outside it.
                return 0;
            }

            RECT hostRect{};
            if (GetWindowRect(self->_optionsPanel.ui.host, &hostRect) == 0 || PtInRect(&hostRect, ptScreen) == FALSE)
            {
                // Only scroll when the wheel is over the options host area.
                return 0;
            }

            if (self->_optionsPanel.scrollMax <= 0)
            {
                // Nothing to scroll; swallow the wheel to avoid comctl32 forwarding loops.
                self->_optionsPanel.wheelRemainder = 0;
                return 0;
            }

            if (hwnd != self->_optionsPanel.ui.host)
            {
                s_routingWheel = true;
                auto reset     = wil::scope_exit([&] { s_routingWheel = false; });
                SendMessageW(self->_optionsPanel.ui.host, msg, wp, lp);
                return 0;
            }

            break;
        }
        case WM_NCDESTROY:
        {
            const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kCompareOptionsWheelRouteOriginalWndProcProp);
            RemovePropW(hwnd, kCompareOptionsWheelRouteStateProp);
            RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kCompareOptionsWheelRouteOriginalWndProcProp, CompareOptionsWheelRouteWndProc);
            return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp)
                                   : DefWindowProcW(hwnd, msg, wp, lp);
        }
    }

    return CallStoredWndProc(hwnd, kCompareOptionsWheelRouteOriginalWndProcProp, msg, wp, lp);
}

LRESULT CALLBACK CompareOptionsDxHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(GetPropW(hwnd, kCompareOptionsDxHostStateProp));
    if (! self)
    {
        return CallStoredWndProc(hwnd, kCompareOptionsDxHostOriginalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kCompareOptionsDxHostOriginalWndProcProp);
        RemovePropW(hwnd, kCompareOptionsDxHostStateProp);
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kCompareOptionsDxHostOriginalWndProcProp, CompareOptionsDxHostWndProc);

        bool handled = false;
        static_cast<void>(self->HandleOptionsDxHostMessage(hwnd, msg, wp, lp, handled));

        return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    if (msg == WM_MOUSEWHEEL && self->_optionsPanel.dxUi && self->_optionsPanel.dxUi->body.hostHwnd && self->_optionsPanel.dxUi->body.hostHwnd.get() == hwnd)
    {
        if (! self->_optionsPanel.dlg || ! self->_optionsPanel.ui.host)
        {
            return 0;
        }

        static thread_local bool s_routingWheel = false;
        if (! s_routingWheel)
        {
            POINT ptScreen{};
            ptScreen.x = GET_X_LPARAM(lp);
            ptScreen.y = GET_Y_LPARAM(lp);

            RECT dlgRect{};
            RECT hostRect{};
            if (GetWindowRect(self->_optionsPanel.dlg.get(), &dlgRect) != 0 && PtInRect(&dlgRect, ptScreen) != FALSE &&
                GetWindowRect(self->_optionsPanel.ui.host, &hostRect) != 0 && PtInRect(&hostRect, ptScreen) != FALSE)
            {
                if (self->_optionsPanel.scrollMax <= 0)
                {
                    self->_optionsPanel.wheelRemainder = 0;
                    return 0;
                }

                s_routingWheel = true;
                auto reset     = wil::scope_exit([&] { s_routingWheel = false; });
                SendMessageW(self->_optionsPanel.ui.host, msg, wp, lp);
                return 0;
            }
        }
    }

    bool handled         = false;
    const LRESULT result = self->HandleOptionsDxHostMessage(hwnd, msg, wp, lp, handled);
    if (handled)
    {
        return result;
    }

    return CallStoredWndProc(hwnd, kCompareOptionsDxHostOriginalWndProcProp, msg, wp, lp);
}

void CompareDirectoriesWindow::EnsureOptionsControlsCreated(HWND dlg) noexcept
{
    if (! dlg || _optionsPanel.ui.host)
    {
        return;
    }

    _optionsPanel.ui.host = CreateCompareOptionsDxHostWindow(dlg, WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, WS_EX_CONTROLPARENT);
    if (_optionsPanel.ui.host)
    {
        const wchar_t* hostTheme = _theme.highContrast ? L"" : (_theme.dark ? L"DarkMode_Explorer" : L"Explorer");
        SetWindowTheme(_optionsPanel.ui.host, hostTheme, nullptr);
        SendMessageW(_optionsPanel.ui.host, WM_THEMECHANGED, 0, 0);

        if (! SetPropW(_optionsPanel.ui.host, kCompareOptionsHostStateProp, reinterpret_cast<HANDLE>(this)) ||
            ! InstallWndProcHook(_optionsPanel.ui.host, kCompareOptionsHostOriginalWndProcProp, CompareOptionsHostWndProc))
        {
            RemovePropW(_optionsPanel.ui.host, kCompareOptionsHostStateProp);
        }
    }

    if (! _optionsPanel.ui.host)
    {
        return;
    }

    static_cast<void>(EnsureOptionsDxStaticHosts());
    static_cast<void>(EnsureOptionsDxButtonHosts());
    static_cast<void>(EnsureOptionsDxToggleHosts());
    static_cast<void>(EnsureOptionsDxEditHosts());

    static_cast<void>(SetPropW(dlg, kCompareOptionsWheelRouteStateProp, reinterpret_cast<HANDLE>(this)));
    static_cast<void>(InstallWndProcHook(dlg, kCompareOptionsWheelRouteOriginalWndProcProp, CompareOptionsWheelRouteWndProc));
    EnumChildWindows(dlg,
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto* self = reinterpret_cast<CompareDirectoriesWindow*>(lParam);
        if (! self || child == self->_optionsPanel.ui.host)
        {
            return TRUE;
        }
        static_cast<void>(SetPropW(child, kCompareOptionsWheelRouteStateProp, reinterpret_cast<HANDLE>(self)));
        static_cast<void>(InstallWndProcHook(child, kCompareOptionsWheelRouteOriginalWndProcProp, CompareOptionsWheelRouteWndProc));
        return TRUE;
    },
                     reinterpret_cast<LPARAM>(this));
}
void CompareDirectoriesWindow::PaintOptionsHostBackgroundAndCards(HDC hdc, HWND host) noexcept
{
    if (! hdc || ! host)
    {
        return;
    }

    RECT rc{};
    GetClientRect(host, &rc);

    D2DHdcPaint::Session paint;
    if (paint.Begin(hdc, rc))
    {
        paint.FillRectangle(rc, _theme.windowBackground);

        if (_theme.systemHighContrast || _theme.highContrast || _optionsPanel.cards.empty())
        {
            return;
        }

        const UINT dpi         = GetDpiForWindow(host);
        const float radius     = static_cast<float>(UiMetrics::ScaleDip(dpi, 6));
        const COLORREF surface = UiMetrics::GetControlSurfaceColor(_theme);
        const COLORREF border  = UiMetrics::BlendColorRefWeightedTruncate(surface, _theme.menu.text, _theme.dark ? 40 : 30, 255);

        for (const RECT& card : _optionsPanel.cards)
        {
            paint.FillRoundedRectangle(card, radius, surface, border);
        }

        if (_optionsPanel.useTwoColumns && _optionsPanel.twoColumnSeparatorX > rc.left && _optionsPanel.twoColumnSeparatorX < rc.right)
        {
            const COLORREF separator = UiMetrics::BlendColorRefWeightedTruncate(surface, _theme.menu.text, _theme.dark ? 28 : 20, 255);
            paint.DrawLine(static_cast<float>(_optionsPanel.twoColumnSeparatorX),
                           static_cast<float>(rc.top),
                           static_cast<float>(_optionsPanel.twoColumnSeparatorX),
                           static_cast<float>(rc.bottom),
                           separator);
        }
        return;
    }

    if (_optionsPanel.backgroundBrush)
    {
        FillRect(hdc, &rc, _optionsPanel.backgroundBrush.get());
    }

    if (_theme.systemHighContrast || _theme.highContrast || _optionsPanel.cards.empty())
    {
        return;
    }
}
void CompareDirectoriesWindow::LayoutOptionsControls() noexcept
{
    if (! _optionsPanel.dlg || ! _optionsPanel.ui.host)
    {
        return;
    }

    Debug::Perf::Scope layoutPerf(L"compare.ui.options_layout_us");

    EnsureCompareSession();
    ApplyOptionsDraftInterlocks();
    const bool contentCompareSupported = ! _session || _session->IsContentCompareSupported();
    const UINT contentCompareDescId    = contentCompareSupported ? IDS_COMPARE_OPTIONS_CONTENT_DESC : IDS_COMPARE_OPTIONS_CONTENT_DESC_UNSUPPORTED;
    SyncOptionsDxStatics();
    SyncOptionsDxToggles();

    RECT rcDlg{};
    if (! GetClientRect(_optionsPanel.dlg.get(), &rcDlg))
    {
        return;
    }

    const int dlgW = std::max(0l, rcDlg.right - rcDlg.left);
    const int dlgH = std::max(0l, rcDlg.bottom - rcDlg.top);
    layoutPerf.SetValue0(static_cast<uint64_t>(dlgW));

    const UINT dpi = GetDpiForWindow(_optionsPanel.dlg.get());

    const int margin       = UiMetrics::ScaleDip(dpi, 16);
    const int gapY         = UiMetrics::ScaleDip(dpi, 12);
    const int rowHeight    = std::max(1, UiMetrics::ScaleDip(dpi, 26));
    const int titleHeight  = std::max(1, UiMetrics::ScaleDip(dpi, 18));
    const int headerHeight = std::max(1, UiMetrics::ScaleDip(dpi, 20));

    const int cardPaddingX                           = UiMetrics::ScaleDip(dpi, 12);
    const int cardPaddingY                           = UiMetrics::ScaleDip(dpi, 8);
    const int cardGapY                               = UiMetrics::ScaleDip(dpi, 2);
    const int cardGapX                               = UiMetrics::ScaleDip(dpi, 12);
    const int cardSpacingY                           = UiMetrics::ScaleDip(dpi, 8);
    const int sectionSpacing                         = UiMetrics::ScaleDip(dpi, 16);
    const int minToggleWidth                         = UiMetrics::ScaleDip(dpi, 90);
    const int columnSeparatorAreaW                   = UiMetrics::ScaleDip(dpi, 28);
    const int minColumnW                             = UiMetrics::ScaleDip(dpi, 360);
    const CompareOptionsTypographyFormats typography = CreateCompareOptionsTypographyFormats(_optionsPanel.dlg.get());
    _optionsPanel.usesDxUiTypographyMetrics          = typography.IsReady();
    layoutPerf.SetDetail(_optionsPanel.usesDxUiTypographyMetrics ? L"dx-typography" : L"fallback-typography");
    layoutPerf.SetValue1(_optionsPanel.usesDxUiTypographyMetrics ? 1u : 0u);

    const HWND okBtn     = GetDlgItem(_optionsPanel.dlg.get(), IDOK);
    const HWND cancelBtn = GetDlgItem(_optionsPanel.dlg.get(), IDCANCEL);

    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
    const int onWidth           = MeasureCompareOptionsTextWidth(typography, typography.headerFormat.get(), onLabel);
    const int offWidth          = MeasureCompareOptionsTextWidth(typography, typography.headerFormat.get(), offLabel);

    const int togglePaddingX      = UiMetrics::ScaleDip(dpi, 6);
    const int toggleGapX          = UiMetrics::ScaleDip(dpi, 8);
    const int toggleTrackW        = UiMetrics::ScaleDip(dpi, 34);
    const int toggleStateTextW    = std::max(onWidth, offWidth);
    const int measuredToggleWidth = std::max(minToggleWidth, (2 * togglePaddingX) + toggleStateTextW + toggleGapX + toggleTrackW);

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

    const int buttonPadX = UiMetrics::ScaleDip(dpi, 12);
    const int minBtnW    = UiMetrics::ScaleDip(dpi, 120);
    const int buttonGapX = UiMetrics::ScaleDip(dpi, 8);

    const auto measureButtonWidth = [&](HWND btn) noexcept -> int
    {
        const std::wstring text = NormalizeDialogButtonCaption(getWindowText(btn));
        const int textW         = MeasureCompareOptionsTextWidth(typography, typography.bodyFormat.get(), text);
        return std::max(minBtnW, (2 * buttonPadX) + textW);
    };

    const auto measureWindowHeight = [](HWND hwnd) noexcept -> int
    {
        if (! hwnd)
        {
            return 0;
        }

        RECT rc{};
        if (GetWindowRect(hwnd, &rc) == FALSE)
        {
            return 0;
        }

        return std::max(0l, rc.bottom - rc.top);
    };

    const int buttonHeight      = std::max(UiMetrics::ScaleDip(dpi, 32), std::max(measureWindowHeight(okBtn), measureWindowHeight(cancelBtn)));
    const int footerButtonWidth = std::max(measureButtonWidth(okBtn), measureButtonWidth(cancelBtn));
    const int okW               = footerButtonWidth;
    const int cancelW           = footerButtonWidth;

    const int buttonsY = std::max(0, dlgH - margin - buttonHeight);

    const UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;

    int nextRight = std::max(0, dlgW - margin);
    if (cancelBtn)
    {
        nextRight -= cancelW;
        if (_optionsPanel.dxUi && _optionsPanel.dxUi->usesDxUiButtons && _optionsPanel.dxUi->cancelButton.hostHwnd)
        {
            ShowWindow(cancelBtn, SW_HIDE);
            if (_optionsPanel.dxUi->cancelButton.button)
            {
                _optionsPanel.dxUi->cancelButton.button->SetBounds(
                    D2D1::RectF(0.0f,
                                0.0f,
                                _optionsPanel.dxUi->cancelButton.host.PixelsToDip(static_cast<float>(cancelW)),
                                _optionsPanel.dxUi->cancelButton.host.PixelsToDip(static_cast<float>(buttonHeight))));
            }
            SetWindowPos(_optionsPanel.dxUi->cancelButton.hostHwnd.get(), nullptr, nextRight, buttonsY, cancelW, buttonHeight, flags | SWP_SHOWWINDOW);
            _optionsPanel.dxUi->cancelButton.host.SetCancelButton(_optionsPanel.dxUi->cancelButton.button);
            _optionsPanel.dxUi->cancelButton.host.Invalidate();
        }
        else
        {
            SetWindowPos(cancelBtn, nullptr, nextRight, buttonsY, cancelW, buttonHeight, flags);
        }
        nextRight -= buttonGapX;
    }
    if (okBtn)
    {
        nextRight -= okW;
        if (_optionsPanel.dxUi && _optionsPanel.dxUi->usesDxUiButtons && _optionsPanel.dxUi->okButton.hostHwnd)
        {
            ShowWindow(okBtn, SW_HIDE);
            if (_optionsPanel.dxUi->okButton.button)
            {
                _optionsPanel.dxUi->okButton.button->SetBounds(D2D1::RectF(0.0f,
                                                                           0.0f,
                                                                           _optionsPanel.dxUi->okButton.host.PixelsToDip(static_cast<float>(okW)),
                                                                           _optionsPanel.dxUi->okButton.host.PixelsToDip(static_cast<float>(buttonHeight))));
            }
            SetWindowPos(_optionsPanel.dxUi->okButton.hostHwnd.get(), nullptr, nextRight, buttonsY, okW, buttonHeight, flags | SWP_SHOWWINDOW);
            _optionsPanel.dxUi->okButton.host.SetDefaultButton(_optionsPanel.dxUi->okButton.button);
            _optionsPanel.dxUi->okButton.host.Invalidate();
        }
        else
        {
            SetWindowPos(okBtn, nullptr, nextRight, buttonsY, okW, buttonHeight, flags);
        }
    }

    const int hostX = margin;
    const int hostY = margin;
    const int hostW = std::max(0, dlgW - 2 * margin);
    const int hostH = std::max(0, buttonsY - gapY - hostY);
    SetWindowPos(_optionsPanel.ui.host, nullptr, hostX, hostY, hostW, hostH, flags);

    RECT hostClient{};
    if (! GetClientRect(_optionsPanel.ui.host, &hostClient))
    {
        return;
    }

    const auto computeToggleWidth = [&](int contentW) noexcept -> int
    {
        const int maxAvailable = std::max(0, contentW - 2 * cardPaddingX);
        return std::min(maxAvailable, measuredToggleWidth);
    };

    auto computeToggleCardHeight = [&](int contentW, std::wstring_view descText, int toggleW) noexcept -> int
    {
        const int textW    = std::max(0, contentW - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH    = MeasureCompareOptionsWrappedTextHeight(typography, typography.infoFormat.get(), textW, descText);
        const int contentH = std::max(0, titleHeight + cardGapY + descH);
        const int cardH    = std::max(rowHeight + 2 * cardPaddingY, contentH + 2 * cardPaddingY);
        return cardH;
    };

    auto computeIgnoreCardHeight = [&](int contentW, std::wstring_view descText, int toggleW, bool showEdit) noexcept -> int
    {
        const int textW = std::max(0, contentW - 2 * cardPaddingX - cardGapX - toggleW);
        const int descH = MeasureCompareOptionsWrappedTextHeight(typography, typography.infoFormat.get(), textW, descText);

        int contentH = std::max(0, titleHeight + cardGapY + descH);
        if (showEdit)
        {
            contentH += cardGapY + rowHeight;
        }
        const int cardH = std::max(rowHeight + 2 * cardPaddingY, contentH + 2 * cardPaddingY);
        return cardH;
    };

    struct OptionsSectionPlacement final
    {
        int x     = 0;
        int y     = 0;
        int width = 0;
    };

    struct OptionsCardPlacement final
    {
        int x                 = 0;
        int y                 = 0;
        int width             = 0;
        int height            = 0;
        int toggleWidth       = 0;
        int textWidth         = 0;
        int descriptionHeight = 0;
        bool showEdit         = false;
    };

    struct OptionsLayoutMeasurement final
    {
        bool useTwoColumns = false;
        int separatorX     = -1;
        int contentHeight  = 0;

        OptionsSectionPlacement headerSubdirs;
        OptionsSectionPlacement headerCompare;
        OptionsSectionPlacement headerAdvanced;
        OptionsSectionPlacement headerIgnore;

        OptionsCardPlacement compareSubdirectories;
        OptionsCardPlacement compareSize;
        OptionsCardPlacement compareDateTime;
        OptionsCardPlacement compareAttributes;
        OptionsCardPlacement compareContent;
        OptionsCardPlacement compareSubdirAttributes;
        OptionsCardPlacement selectSubdirsOnlyInOnePane;
        OptionsCardPlacement keepIdenticalItems;
        OptionsCardPlacement ignoreFiles;
        OptionsCardPlacement ignoreDirectories;
    };

    const auto measureLayout = [&](int contentW) noexcept -> OptionsLayoutMeasurement
    {
        OptionsLayoutMeasurement layout{};
        layout.useTwoColumns = (! _theme.systemHighContrast && ! _theme.highContrast) && (contentW >= (2 * minColumnW + columnSeparatorAreaW));
        layout.separatorX    = layout.useTwoColumns ? (std::max(0, (contentW - columnSeparatorAreaW) / 2) + (columnSeparatorAreaW / 2)) : -1;

        const bool ignoreFilesOn = _optionsPanel.draft.ignoreFiles;
        const bool ignoreDirsOn  = _optionsPanel.draft.ignoreDirectories;

        const auto measureSectionHeader = [&](OptionsSectionPlacement& placement, int contentX, int availableWidth, int& y) noexcept
        {
            placement.x     = contentX + cardPaddingX;
            placement.y     = y;
            placement.width = std::max(0, availableWidth - 2 * cardPaddingX);
            y += headerHeight + gapY;
        };

        const auto measureToggleCard =
            [&](OptionsCardPlacement& placement, int contentX, int availableWidth, int toggleW, std::wstring_view descText, int& y) noexcept
        {
            placement.x                 = contentX;
            placement.y                 = y;
            placement.width             = availableWidth;
            placement.toggleWidth       = toggleW;
            placement.textWidth         = std::max(0, availableWidth - 2 * cardPaddingX - cardGapX - toggleW);
            placement.descriptionHeight = MeasureCompareOptionsWrappedTextHeight(typography, typography.infoFormat.get(), placement.textWidth, descText);
            placement.height            = computeToggleCardHeight(availableWidth, descText, toggleW);
            placement.showEdit          = false;
            y += placement.height + cardSpacingY;
        };

        const auto measureIgnoreCard =
            [&](OptionsCardPlacement& placement, int contentX, int availableWidth, int toggleW, std::wstring_view descText, bool showEdit, int& y) noexcept
        {
            placement.x                 = contentX;
            placement.y                 = y;
            placement.width             = availableWidth;
            placement.toggleWidth       = toggleW;
            placement.textWidth         = std::max(0, availableWidth - 2 * cardPaddingX - cardGapX - toggleW);
            placement.descriptionHeight = MeasureCompareOptionsWrappedTextHeight(typography, typography.infoFormat.get(), placement.textWidth, descText);
            placement.height            = computeIgnoreCardHeight(availableWidth, descText, toggleW, showEdit);
            placement.showEdit          = showEdit;
            y += placement.height + cardSpacingY;
        };

        if (! layout.useTwoColumns)
        {
            const int toggleW = computeToggleWidth(contentW);
            int y             = 0;

            measureSectionHeader(layout.headerSubdirs, 0, contentW, y);
            measureToggleCard(layout.compareSubdirectories, 0, contentW, toggleW, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIRS_DESC), y);

            y += sectionSpacing;
            measureSectionHeader(layout.headerCompare, 0, contentW, y);
            measureToggleCard(layout.compareSize, 0, contentW, toggleW, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SIZE_DESC), y);
            measureToggleCard(layout.compareDateTime, 0, contentW, toggleW, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_DATETIME_DESC), y);
            measureToggleCard(layout.compareAttributes, 0, contentW, toggleW, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC), y);
            measureToggleCard(layout.compareContent, 0, contentW, toggleW, LoadStringResource(nullptr, contentCompareDescId), y);

            y += sectionSpacing;
            measureSectionHeader(layout.headerAdvanced, 0, contentW, y);
            measureToggleCard(layout.compareSubdirAttributes, 0, contentW, toggleW, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC), y);
            measureToggleCard(layout.selectSubdirsOnlyInOnePane, 0, contentW, toggleW, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC), y);
            measureToggleCard(layout.keepIdenticalItems, 0, contentW, toggleW, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_DESC), y);

            y += sectionSpacing;
            measureSectionHeader(layout.headerIgnore, 0, contentW, y);
            measureIgnoreCard(layout.ignoreFiles, 0, contentW, toggleW, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC), ignoreFilesOn, y);
            measureIgnoreCard(
                layout.ignoreDirectories, 0, contentW, toggleW, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC), ignoreDirsOn, y);

            layout.contentHeight = y;
            return layout;
        }

        const int leftW        = std::max(0, (contentW - columnSeparatorAreaW) / 2);
        const int rightX       = leftW + columnSeparatorAreaW;
        const int rightW       = std::max(0, contentW - rightX);
        const int toggleWLeft  = computeToggleWidth(leftW);
        const int toggleWRight = computeToggleWidth(rightW);

        int leftY = 0;
        measureSectionHeader(layout.headerSubdirs, 0, leftW, leftY);
        measureToggleCard(layout.compareSubdirectories, 0, leftW, toggleWLeft, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIRS_DESC), leftY);

        leftY += sectionSpacing;
        measureSectionHeader(layout.headerCompare, 0, leftW, leftY);
        measureToggleCard(layout.compareSize, 0, leftW, toggleWLeft, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SIZE_DESC), leftY);
        measureToggleCard(layout.compareDateTime, 0, leftW, toggleWLeft, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_DATETIME_DESC), leftY);
        measureToggleCard(layout.compareAttributes, 0, leftW, toggleWLeft, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC), leftY);
        measureToggleCard(layout.compareContent, 0, leftW, toggleWLeft, LoadStringResource(nullptr, contentCompareDescId), leftY);

        int rightY = 0;
        measureSectionHeader(layout.headerAdvanced, rightX, rightW, rightY);
        measureToggleCard(
            layout.compareSubdirAttributes, rightX, rightW, toggleWRight, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC), rightY);
        measureToggleCard(
            layout.selectSubdirsOnlyInOnePane, rightX, rightW, toggleWRight, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC), rightY);
        measureToggleCard(
            layout.keepIdenticalItems, rightX, rightW, toggleWRight, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_DESC), rightY);

        rightY += sectionSpacing;
        measureSectionHeader(layout.headerIgnore, rightX, rightW, rightY);
        measureIgnoreCard(
            layout.ignoreFiles, rightX, rightW, toggleWRight, LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC), ignoreFilesOn, rightY);
        measureIgnoreCard(layout.ignoreDirectories,
                          rightX,
                          rightW,
                          toggleWRight,
                          LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC),
                          ignoreDirsOn,
                          rightY);

        layout.contentHeight = std::max(leftY, rightY);
        return layout;
    };

    const int viewportW = std::max(0l, hostClient.right - hostClient.left);
    const int viewportH = std::max(0l, hostClient.bottom - hostClient.top);

    const OptionsLayoutMeasurement initialLayout = measureLayout(viewportW);
    int contentHeight                            = initialLayout.contentHeight;
    const bool wantsVScroll                      = viewportH > 0 && contentHeight > viewportH;

    LONG_PTR styleNow    = GetWindowLongPtrW(_optionsPanel.ui.host, GWL_STYLE);
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
        SetWindowLongPtrW(_optionsPanel.ui.host, GWL_STYLE, styleWanted);
        SetWindowPos(_optionsPanel.ui.host, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);

        const wchar_t* hostTheme = _theme.highContrast ? L"" : (_theme.dark ? L"DarkMode_Explorer" : L"Explorer");
        SetWindowTheme(_optionsPanel.ui.host, hostTheme, nullptr);
        SendMessageW(_optionsPanel.ui.host, WM_THEMECHANGED, 0, 0);
    }

    GetClientRect(_optionsPanel.ui.host, &hostClient);
    const int viewportW2                          = std::max(0l, hostClient.right - hostClient.left);
    const int viewportH2                          = std::max(0l, hostClient.bottom - hostClient.top);
    const OptionsLayoutMeasurement measuredLayout = measureLayout(viewportW2);

    _optionsPanel.useTwoColumns       = measuredLayout.useTwoColumns;
    _optionsPanel.twoColumnSeparatorX = measuredLayout.separatorX;

    contentHeight                   = measuredLayout.contentHeight;
    _optionsPanel.bodyContentHeight = contentHeight;
    _optionsPanel.scrollMax         = (viewportH2 > 0) ? std::max(0, contentHeight - viewportH2) : 0;
    _optionsPanel.scrollOffset      = std::clamp(_optionsPanel.scrollOffset, 0, _optionsPanel.scrollMax);
    if (_optionsPanel.scrollMax <= 0)
    {
        _optionsPanel.scrollOffset   = 0;
        _optionsPanel.wheelRemainder = 0;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = std::max(0, contentHeight - 1);
    si.nPage  = (viewportH2 > 0) ? static_cast<UINT>(viewportH2) : 0u;
    si.nPos   = _optionsPanel.scrollOffset;
    SetScrollInfo(_optionsPanel.ui.host, SB_VERT, &si, TRUE);

    _optionsPanel.cards.clear();

    auto pushCard = [&](int x, int top, int width, int height) noexcept
    {
        RECT card{};
        card.left   = x;
        card.top    = top - _optionsPanel.scrollOffset;
        card.right  = x + width;
        card.bottom = top + height - _optionsPanel.scrollOffset;
        _optionsPanel.cards.push_back(card);
    };

    if (_optionsPanel.dxUi && _optionsPanel.dxUi->usesDxUiStatics && _optionsPanel.dxUi->body.hostHwnd)
    {
        auto& bodyDx       = _optionsPanel.dxUi->body;
        const auto pxToDip = [dpi](int px) noexcept { return (static_cast<float>(px) * 96.0f) / static_cast<float>(std::max<UINT>(1u, dpi)); };

        const int scrollOffset = _optionsPanel.scrollOffset;
        SetWindowPos(bodyDx.hostHwnd.get(), nullptr, 0, 0, viewportW2, viewportH2, flags | SWP_SHOWWINDOW);

        const auto setHeaderVisible = [](Label* label, bool visible) noexcept
        {
            if (label)
            {
                label->SetVisible(visible);
            }
        };

        const auto setToggleCardVisible = [](OptionsToggleCardDx& card, bool visible) noexcept
        {
            if (card.card)
            {
                card.card->SetVisible(visible);
            }
            if (card.title)
            {
                card.title->SetVisible(visible);
            }
            if (card.description)
            {
                card.description->SetVisible(visible);
            }
            if (card.toggle)
            {
                card.toggle->SetVisible(visible);
            }
        };

        const auto setIgnoreCardVisible = [](OptionsIgnoreCardDx& card, bool visible, bool showEdit) noexcept
        {
            if (card.card)
            {
                card.card->SetVisible(visible);
            }
            if (card.title)
            {
                card.title->SetVisible(visible);
            }
            if (card.description)
            {
                card.description->SetVisible(visible);
            }
            if (card.toggle)
            {
                card.toggle->SetVisible(visible);
            }
            if (card.edit)
            {
                card.edit->SetVisible(visible && showEdit);
            }
        };

        const auto layoutSectionHeaderDx = [&](Label* header, const UINT textId, const OptionsSectionPlacement& placement) noexcept
        {
            if (! header)
            {
                return;
            }

            const bool visible = placement.width > 0;
            header->SetText(LoadStringResource(nullptr, textId));
            header->SetVisible(visible);
            if (! visible)
            {
                return;
            }

            const int headerTop = placement.y - scrollOffset;
            header->SetBounds(D2D1::RectF(pxToDip(placement.x), pxToDip(headerTop), pxToDip(placement.x + placement.width), pxToDip(headerTop + headerHeight)));
        };

        const auto layoutToggleCardDx = [&](OptionsToggleCardDx& card, const UINT titleId, const UINT descId, const OptionsCardPlacement& placement) noexcept
        {
            const bool visible = placement.width > 0 && placement.height > 0;
            setToggleCardVisible(card, visible);
            if (! visible)
            {
                return;
            }

            const std::wstring titleText = LoadStringResource(nullptr, titleId);
            const std::wstring descText  = LoadStringResource(nullptr, descId);
            const int textX              = placement.x + cardPaddingX;
            const int cardTop            = placement.y - scrollOffset;
            const int textY              = cardTop + cardPaddingY;
            const int toggleX            = placement.x + placement.width - cardPaddingX - placement.toggleWidth;
            const int toggleY            = cardTop + (placement.height - rowHeight) / 2;

            pushCard(placement.x, placement.y, placement.width, placement.height);
            card.card->SetBounds(
                D2D1::RectF(pxToDip(placement.x), pxToDip(cardTop), pxToDip(placement.x + placement.width), pxToDip(cardTop + placement.height)));
            card.title->SetText(titleText);
            card.title->SetBounds(D2D1::RectF(pxToDip(textX), pxToDip(textY), pxToDip(textX + placement.textWidth), pxToDip(textY + titleHeight)));
            card.description->SetText(descText);
            card.description->SetBounds(D2D1::RectF(pxToDip(textX),
                                                    pxToDip(textY + titleHeight + cardGapY),
                                                    pxToDip(textX + placement.textWidth),
                                                    pxToDip(textY + titleHeight + cardGapY + placement.descriptionHeight)));
            card.toggle->SetBounds(D2D1::RectF(pxToDip(toggleX), pxToDip(toggleY), pxToDip(toggleX + placement.toggleWidth), pxToDip(toggleY + rowHeight)));
        };

        const auto layoutIgnoreCardDx = [&](OptionsIgnoreCardDx& card, const UINT titleId, const UINT descId, const OptionsCardPlacement& placement) noexcept
        {
            const bool visible  = placement.width > 0 && placement.height > 0;
            const bool showEdit = placement.showEdit;
            setIgnoreCardVisible(card, visible, showEdit);
            if (! visible)
            {
                return;
            }

            const std::wstring titleText = LoadStringResource(nullptr, titleId);
            const std::wstring descText  = LoadStringResource(nullptr, descId);
            const int textX              = placement.x + cardPaddingX;
            const int cardTop            = placement.y - scrollOffset;
            const int textY              = cardTop + cardPaddingY;
            const int toggleX            = placement.x + placement.width - cardPaddingX - placement.toggleWidth;
            const int toggleY            = cardTop + cardPaddingY;

            pushCard(placement.x, placement.y, placement.width, placement.height);
            card.card->SetBounds(
                D2D1::RectF(pxToDip(placement.x), pxToDip(cardTop), pxToDip(placement.x + placement.width), pxToDip(cardTop + placement.height)));
            card.title->SetText(titleText);
            card.title->SetMnemonicTarget(showEdit ? static_cast<RedSalamander::DxUi::Control*>(card.edit)
                                                   : static_cast<RedSalamander::DxUi::Control*>(card.toggle));
            card.title->SetBounds(D2D1::RectF(pxToDip(textX), pxToDip(textY), pxToDip(textX + placement.textWidth), pxToDip(textY + titleHeight)));
            card.description->SetText(descText);
            card.description->SetBounds(D2D1::RectF(pxToDip(textX),
                                                    pxToDip(textY + titleHeight + cardGapY),
                                                    pxToDip(textX + placement.textWidth),
                                                    pxToDip(textY + titleHeight + cardGapY + placement.descriptionHeight)));
            card.toggle->SetBounds(D2D1::RectF(pxToDip(toggleX), pxToDip(toggleY), pxToDip(toggleX + placement.toggleWidth), pxToDip(toggleY + rowHeight)));
            if (showEdit)
            {
                const int editX   = placement.x + cardPaddingX;
                const int editW   = std::max(0, placement.width - 2 * cardPaddingX);
                const int editTop = cardTop + cardPaddingY + titleHeight + cardGapY + placement.descriptionHeight + cardGapY;
                card.edit->SetBounds(D2D1::RectF(pxToDip(editX), pxToDip(editTop), pxToDip(editX + editW), pxToDip(editTop + rowHeight)));
            }
        };

        layoutSectionHeaderDx(bodyDx.headerSubdirs, IDS_COMPARE_OPTIONS_SECTION_SUBDIRS, measuredLayout.headerSubdirs);
        layoutToggleCardDx(
            bodyDx.compareSubdirectories, IDS_COMPARE_OPTIONS_SUBDIRS_TITLE, IDS_COMPARE_OPTIONS_SUBDIRS_DESC, measuredLayout.compareSubdirectories);
        layoutSectionHeaderDx(bodyDx.headerCompare, IDS_COMPARE_OPTIONS_SECTION_COMPARE, measuredLayout.headerCompare);
        layoutToggleCardDx(bodyDx.compareSize, IDS_COMPARE_OPTIONS_SIZE_TITLE, IDS_COMPARE_OPTIONS_SIZE_DESC, measuredLayout.compareSize);
        layoutToggleCardDx(bodyDx.compareDateTime, IDS_COMPARE_OPTIONS_DATETIME_TITLE, IDS_COMPARE_OPTIONS_DATETIME_DESC, measuredLayout.compareDateTime);
        layoutToggleCardDx(
            bodyDx.compareAttributes, IDS_COMPARE_OPTIONS_ATTRIBUTES_TITLE, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC, measuredLayout.compareAttributes);
        layoutToggleCardDx(bodyDx.compareContent, IDS_COMPARE_OPTIONS_CONTENT_TITLE, contentCompareDescId, measuredLayout.compareContent);
        layoutSectionHeaderDx(bodyDx.headerAdvanced, IDS_COMPARE_OPTIONS_SECTION_ADVANCED, measuredLayout.headerAdvanced);
        layoutToggleCardDx(bodyDx.compareSubdirAttributes,
                           IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_TITLE,
                           IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC,
                           measuredLayout.compareSubdirAttributes);
        layoutToggleCardDx(bodyDx.selectSubdirsOnlyInOnePane,
                           IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_TITLE,
                           IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC,
                           measuredLayout.selectSubdirsOnlyInOnePane);
        layoutToggleCardDx(
            bodyDx.keepIdenticalItems, IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_TITLE, IDS_COMPARE_OPTIONS_KEEP_IDENTICAL_DESC, measuredLayout.keepIdenticalItems);
        layoutSectionHeaderDx(bodyDx.headerIgnore, IDS_COMPARE_OPTIONS_SECTION_IGNORE, measuredLayout.headerIgnore);
        layoutIgnoreCardDx(bodyDx.ignoreFiles, IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE, IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC, measuredLayout.ignoreFiles);
        layoutIgnoreCardDx(bodyDx.ignoreDirectories,
                           IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_TITLE,
                           IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC,
                           measuredLayout.ignoreDirectories);

        bodyDx.host.Invalidate();
        InvalidateRect(_optionsPanel.ui.host, nullptr, TRUE);
        return;
    }

    InvalidateRect(_optionsPanel.ui.host, nullptr, TRUE);
}
void CompareDirectoriesWindow::ShowOptionsPanel(bool show) noexcept
{
    if (! _optionsPanel.dlg)
    {
        return;
    }

    if (show)
    {
        _optionsPanel.staleFromExternalReload = false;
        LoadOptionsControlsFromSettings();
        ShowWindow(_optionsPanel.dlg.get(), SW_SHOW);
        RegisterOptionsReloadParticipant();
        if (const HWND fw = _folderWindow.GetHwnd())
        {
            ShowWindow(fw, SW_HIDE);
        }

        Layout();
        ApplyOptionsDialogTheme();
        SetWindowPos(_optionsPanel.dlg.get(), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        RedrawWindow(_optionsPanel.dlg.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
        SetFocus(_optionsPanel.dlg.get());
        UpdateCompareWatermark();
        UpdateViewMenuChecks();
        return;
    }

    UnregisterOptionsReloadParticipant();
    _optionsPanel.staleFromExternalReload = false;
    ShowWindow(_optionsPanel.dlg.get(), SW_HIDE);

    if (_compareStarted)
    {
        if (const HWND fw = _folderWindow.GetHwnd())
        {
            ShowWindow(fw, SW_SHOW);
        }

        Layout();
        const HWND focus = GetFocus();
        if (! focus || (focus == _optionsPanel.dlg.get()) || IsChild(_optionsPanel.dlg.get(), focus))
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
    Common::Settings::CompareDirectoriesSettings s = _optionsPanel.draftLoaded ? _optionsPanel.draft : GetEffectiveCompareSettings();
    if (! _optionsPanel.dlg || ! _optionsPanel.ui.host || ! _optionsPanel.draftLoaded)
    {
        return s;
    }

    const bool contentCompareSupported = ! _session || _session->IsContentCompareSupported();
    if (! contentCompareSupported)
    {
        s.compareContent = false;
    }

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

void CompareDirectoriesWindow::ApplyOptionsDraftInterlocks() noexcept
{
    EnsureCompareSession();
    if (_session && ! _session->IsContentCompareSupported())
    {
        _optionsPanel.draft.compareContent = false;
    }

    if (! _optionsPanel.draft.keepIdenticalItems)
    {
        _optionsPanel.draft.showIdenticalItems = false;
    }
}

void CompareDirectoriesWindow::UpdateOptionsDraftToggle(const UINT controlId, const bool checked) noexcept
{
    if (_optionsPanel.syncingDxDraft)
    {
        return;
    }

    bool requiresLayout = false;
    switch (controlId)
    {
        case IDC_CMP_SIZE: _optionsPanel.draft.compareSize = checked; break;
        case IDC_CMP_DATETIME: _optionsPanel.draft.compareDateTime = checked; break;
        case IDC_CMP_ATTRIBUTES: _optionsPanel.draft.compareAttributes = checked; break;
        case IDC_CMP_CONTENT: _optionsPanel.draft.compareContent = checked; break;
        case IDC_CMP_SUBDIRECTORIES: _optionsPanel.draft.compareSubdirectories = checked; break;
        case IDC_CMP_SUBDIR_ATTRIBUTES: _optionsPanel.draft.compareSubdirectoryAttributes = checked; break;
        case IDC_CMP_SELECT_SUBDIRS_ONLY_ONE_PANE: _optionsPanel.draft.selectSubdirsOnlyInOnePane = checked; break;
        case IDC_CMP_KEEP_IDENTICAL: _optionsPanel.draft.keepIdenticalItems = checked; break;
        case IDC_CMP_IGNORE_FILES:
            _optionsPanel.draft.ignoreFiles = checked;
            requiresLayout                  = true;
            break;
        case IDC_CMP_IGNORE_DIRECTORIES:
            _optionsPanel.draft.ignoreDirectories = checked;
            requiresLayout                        = true;
            break;
        default: return;
    }

    _optionsPanel.draftLoaded = true;
    ApplyOptionsDraftInterlocks();
    SyncOptionsDxToggles();
    if (requiresLayout)
    {
        UpdateOptionsVisibility();
    }
}

void CompareDirectoriesWindow::UpdateOptionsDraftText(const UINT controlId, std::wstring_view text) noexcept
{
    if (_optionsPanel.syncingDxDraft)
    {
        return;
    }

    switch (controlId)
    {
        case IDC_CMP_IGNORE_FILES_PATTERNS: _optionsPanel.draft.ignoreFilesPatterns = text; break;
        case IDC_CMP_IGNORE_DIRECTORIES_PATTERNS: _optionsPanel.draft.ignoreDirectoriesPatterns = text; break;
        default: return;
    }

    _optionsPanel.draftLoaded = true;
}

#ifdef ENABLE_TESTS
[[nodiscard]] bool WindowOwnsFocus(HWND hwnd, HWND focused) noexcept
{
    return hwnd && focused && (hwnd == focused || IsChild(hwnd, focused) != FALSE);
}

bool CompareDirectoriesWindow::DebugGetOptionsSnapshot(::CompareDirectoriesOptionsDebugSnapshot& out) const noexcept
{
    out                                    = {};
    out.optionsDialogVisible               = _optionsPanel.dlg && IsWindowVisible(_optionsPanel.dlg.get()) != FALSE;
    out.optionsUsesDxUiStatics             = _optionsPanel.dxUi && _optionsPanel.dxUi->usesDxUiStatics;
    out.optionsUsesDxUiButtons             = _optionsPanel.dxUi && _optionsPanel.dxUi->usesDxUiButtons;
    out.optionsUsesDxUiToggles             = _optionsPanel.dxUi && _optionsPanel.dxUi->usesDxUiToggles;
    out.optionsUsesDxUiEdits               = _optionsPanel.dxUi && _optionsPanel.dxUi->usesDxUiEdits;
    out.usesDxUiTypographyMetrics          = _optionsPanel.usesDxUiTypographyMetrics;
    out.optionsReloadParticipantRegistered = _optionsPanel.reloadParticipantRegistered;
    out.optionsStaleFromExternalReload     = _optionsPanel.staleFromExternalReload;
    out.optionsDialogDirty                 = IsOptionsDialogDirty();
    out.compareSubdirectoriesChecked       = _optionsPanel.draft.compareSubdirectories;
    out.themeDark                          = _theme.dark;
    out.themeHighContrast                  = _theme.highContrast;
    out.themeRainbow                       = _theme.menu.rainbowMode;

    const auto isActuallyVisibleChildWindow = [](HWND hwnd) noexcept
    {
        if (! hwnd || IsWindowVisible(hwnd) == FALSE)
        {
            return false;
        }

        // DxUi text bridges stay WS_VISIBLE for IME routing, but an empty region keeps them off-screen.
        wil::unique_hrgn region(CreateRectRgn(0, 0, 0, 0));
        if (region)
        {
            const int regionType = GetWindowRgn(hwnd, region.get());
            if (regionType == NULLREGION)
            {
                return false;
            }
        }

        return true;
    };

    const auto countIfVisible = [&](HWND hwnd, size_t& count) noexcept
    {
        if (isActuallyVisibleChildWindow(hwnd))
        {
            ++count;
        }
    };

    const auto countHiddenOwnerDrawButton = [](HWND hwnd, size_t& count) noexcept
    {
        if (! hwnd || IsWindowVisible(hwnd) != FALSE)
        {
            return;
        }

        const LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
        if ((style & BS_TYPEMASK) == BS_OWNERDRAW)
        {
            ++count;
        }
    };

    const auto countDxControlIfVisible = [](const RedSalamander::DxUi::Control* control, size_t& count) noexcept
    {
        if (! control || ! control->IsVisible())
        {
            return;
        }

        const D2D1_RECT_F bounds = control->GetBounds();
        if (bounds.right > bounds.left && bounds.bottom > bounds.top)
        {
            ++count;
        }
    };

    if (_optionsPanel.dxUi && _optionsPanel.dxUi->body.hostHwnd && IsWindowVisible(_optionsPanel.dxUi->body.hostHwnd.get()) != FALSE)
    {
        RECT bodyClient{};
        if (GetClientRect(_optionsPanel.dxUi->body.hostHwnd.get(), &bodyClient) != FALSE)
        {
            out.bodyDxHostWidth  = std::max(0l, bodyClient.right - bodyClient.left);
            out.bodyDxHostHeight = std::max(0l, bodyClient.bottom - bodyClient.top);
        }

        if (_optionsPanel.dxUi->body.host.DebugGetRenderCount() != 0u)
        {
            out.visibleBodyRenderedDxHostCount = 1u;
        }
        out.bodyDxHostResizeFailureCount  = static_cast<size_t>(_optionsPanel.dxUi->body.host.DebugGetResizeFailureCount());
        out.bodyDxHostPresentFailureCount = static_cast<size_t>(_optionsPanel.dxUi->body.host.DebugGetPresentFailureCount());
        out.bodyContentHeight             = _optionsPanel.bodyContentHeight;
        out.bodyScrollOffset              = _optionsPanel.scrollOffset;
        out.bodyScrollMax                 = _optionsPanel.scrollMax;
        out.bodyUsesTwoColumns            = _optionsPanel.useTwoColumns;

        const auto& body = _optionsPanel.dxUi->body;
        countDxControlIfVisible(body.headerCompare, out.visibleDxBodyHeaderCount);
        countDxControlIfVisible(body.headerSubdirs, out.visibleDxBodyHeaderCount);
        countDxControlIfVisible(body.headerAdvanced, out.visibleDxBodyHeaderCount);
        countDxControlIfVisible(body.headerIgnore, out.visibleDxBodyHeaderCount);
        countDxControlIfVisible(body.compareSize.card, out.visibleDxBodyCardCount);
        countDxControlIfVisible(body.compareDateTime.card, out.visibleDxBodyCardCount);
        countDxControlIfVisible(body.compareAttributes.card, out.visibleDxBodyCardCount);
        countDxControlIfVisible(body.compareContent.card, out.visibleDxBodyCardCount);
        countDxControlIfVisible(body.compareSubdirectories.card, out.visibleDxBodyCardCount);
        countDxControlIfVisible(body.compareSubdirAttributes.card, out.visibleDxBodyCardCount);
        countDxControlIfVisible(body.selectSubdirsOnlyInOnePane.card, out.visibleDxBodyCardCount);
        countDxControlIfVisible(body.keepIdenticalItems.card, out.visibleDxBodyCardCount);
        countDxControlIfVisible(body.ignoreFiles.card, out.visibleDxBodyCardCount);
        countDxControlIfVisible(body.ignoreDirectories.card, out.visibleDxBodyCardCount);
    }

    const auto countDxButtonHost = [&](const auto& slot) noexcept
    {
        if (slot.hostHwnd)
        {
            ++out.dxFooterButtonHostCount;
            countIfVisible(slot.hostHwnd.get(), out.visibleDxFooterButtonHostCount);
        }
    };

    if (_optionsPanel.dxUi)
    {
        countDxButtonHost(_optionsPanel.dxUi->okButton);
        countDxButtonHost(_optionsPanel.dxUi->cancelButton);
        out.okFooterAttachFailureStage     = _optionsPanel.dxUi->okButton.attachFailureStage;
        out.cancelFooterAttachFailureStage = _optionsPanel.dxUi->cancelButton.attachFailureStage;
    }

    countIfVisible(GetDlgItem(_optionsPanel.dlg.get(), IDOK), out.visibleLegacyFooterButtonCount);
    countIfVisible(GetDlgItem(_optionsPanel.dlg.get(), IDCANCEL), out.visibleLegacyFooterButtonCount);
    countHiddenOwnerDrawButton(GetDlgItem(_optionsPanel.dlg.get(), IDOK), out.hiddenLegacyOwnerDrawFooterButtonCount);
    countHiddenOwnerDrawButton(GetDlgItem(_optionsPanel.dlg.get(), IDCANCEL), out.hiddenLegacyOwnerDrawFooterButtonCount);

    out.visibleNativeBodyControlCount =
        out.visibleLegacyStaticCount + out.visibleLegacyFooterButtonCount + out.visibleLegacyToggleCount + out.visibleLegacyEditCount;

    const HWND focused = GetFocus();
    if (_optionsPanel.dxUi && focused)
    {
        const auto* bodyFocused     = _optionsPanel.dxUi->body.host.GetFocusControl();
        const bool bodyOwnsFocus    = _optionsPanel.dxUi->body.hostHwnd && WindowOwnsFocus(_optionsPanel.dxUi->body.hostHwnd.get(), focused);
        const auto assignBodyTarget = [&](const auto& slot, const CompareDirectoriesOptionsDebugFocusTarget target) noexcept
        {
            if (bodyOwnsFocus && slot == bodyFocused)
            {
                out.focusTarget = target;
                return true;
            }

            return false;
        };

        const auto& body = _optionsPanel.dxUi->body;
        if (assignBodyTarget(body.compareSubdirectories.toggle, CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle) ||
            assignBodyTarget(body.compareSize.toggle, CompareDirectoriesOptionsDebugFocusTarget::CompareSizeToggle) ||
            assignBodyTarget(body.compareDateTime.toggle, CompareDirectoriesOptionsDebugFocusTarget::CompareDateTimeToggle) ||
            assignBodyTarget(body.compareAttributes.toggle, CompareDirectoriesOptionsDebugFocusTarget::CompareAttributesToggle) ||
            assignBodyTarget(body.compareContent.toggle, CompareDirectoriesOptionsDebugFocusTarget::CompareContentToggle) ||
            assignBodyTarget(body.compareSubdirAttributes.toggle, CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirAttributesToggle) ||
            assignBodyTarget(body.selectSubdirsOnlyInOnePane.toggle, CompareDirectoriesOptionsDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle) ||
            assignBodyTarget(body.keepIdenticalItems.toggle, CompareDirectoriesOptionsDebugFocusTarget::KeepIdenticalItemsToggle) ||
            assignBodyTarget(body.ignoreFiles.toggle, CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle) ||
            assignBodyTarget(body.ignoreFiles.edit, CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit) ||
            assignBodyTarget(body.ignoreDirectories.toggle, CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle) ||
            assignBodyTarget(body.ignoreDirectories.edit, CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesEdit))
        {
            return true;
        }

        if (_optionsPanel.dxUi->okButton.button && WindowOwnsFocus(_optionsPanel.dxUi->okButton.hostHwnd.get(), focused) &&
            _optionsPanel.dxUi->okButton.host.GetFocusControl() == _optionsPanel.dxUi->okButton.button)
        {
            out.focusTarget = CompareDirectoriesOptionsDebugFocusTarget::OkButton;
            return true;
        }

        if (_optionsPanel.dxUi->cancelButton.button && WindowOwnsFocus(_optionsPanel.dxUi->cancelButton.hostHwnd.get(), focused) &&
            _optionsPanel.dxUi->cancelButton.host.GetFocusControl() == _optionsPanel.dxUi->cancelButton.button)
        {
            out.focusTarget = CompareDirectoriesOptionsDebugFocusTarget::CancelButton;
            return true;
        }
    }

    return true;
}

bool CompareDirectoriesWindow::DebugFocusOptionsFirstControl() noexcept
{
    if (! _optionsPanel.dxUi)
    {
        return false;
    }

    auto* const toggle = _optionsPanel.dxUi->body.compareSubdirectories.toggle;
    if (! toggle || ! toggle->IsVisible() || ! toggle->IsEnabled())
    {
        return false;
    }

    if (_optionsPanel.dxUi->body.hostHwnd && IsWindow(_optionsPanel.dxUi->body.hostHwnd.get()) != FALSE)
    {
        ::SetFocus(_optionsPanel.dxUi->body.hostHwnd.get());
    }
    _optionsPanel.dxUi->body.host.SetFocusControl(toggle);
    static_cast<void>(EnsureOptionsDxBodyControlVisible(toggle));
    return _optionsPanel.dxUi->body.host.GetFocusControl() == toggle;
}

bool CompareDirectoriesWindow::DebugFocusOptionsTarget(const ::CompareDirectoriesOptionsDebugFocusTarget target) noexcept
{
    if (! _optionsPanel.dxUi)
    {
        return false;
    }

    RedSalamander::DxUi::Control* focusControl = nullptr;
    const auto& body                           = _optionsPanel.dxUi->body;
    switch (target)
    {
        case CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle: focusControl = body.compareSubdirectories.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::CompareSizeToggle: focusControl = body.compareSize.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::CompareDateTimeToggle: focusControl = body.compareDateTime.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::CompareAttributesToggle: focusControl = body.compareAttributes.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::CompareContentToggle: focusControl = body.compareContent.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirAttributesToggle: focusControl = body.compareSubdirAttributes.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle: focusControl = body.selectSubdirsOnlyInOnePane.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::KeepIdenticalItemsToggle: focusControl = body.keepIdenticalItems.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle: focusControl = body.ignoreFiles.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit: focusControl = body.ignoreFiles.edit; break;
        case CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle: focusControl = body.ignoreDirectories.toggle; break;
        case CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesEdit: focusControl = body.ignoreDirectories.edit; break;
        case CompareDirectoriesOptionsDebugFocusTarget::OkButton:
            if (_optionsPanel.dxUi->okButton.button && _optionsPanel.dxUi->okButton.button->IsVisible() && _optionsPanel.dxUi->okButton.button->IsEnabled())
            {
                if (_optionsPanel.dxUi->okButton.hostHwnd && IsWindow(_optionsPanel.dxUi->okButton.hostHwnd.get()) != FALSE)
                {
                    ::SetFocus(_optionsPanel.dxUi->okButton.hostHwnd.get());
                }
                _optionsPanel.dxUi->okButton.host.SetFocusControl(_optionsPanel.dxUi->okButton.button);
                return _optionsPanel.dxUi->okButton.host.GetFocusControl() == _optionsPanel.dxUi->okButton.button;
            }
            return false;
        case CompareDirectoriesOptionsDebugFocusTarget::CancelButton:
            if (_optionsPanel.dxUi->cancelButton.button && _optionsPanel.dxUi->cancelButton.button->IsVisible() &&
                _optionsPanel.dxUi->cancelButton.button->IsEnabled())
            {
                if (_optionsPanel.dxUi->cancelButton.hostHwnd && IsWindow(_optionsPanel.dxUi->cancelButton.hostHwnd.get()) != FALSE)
                {
                    ::SetFocus(_optionsPanel.dxUi->cancelButton.hostHwnd.get());
                }
                _optionsPanel.dxUi->cancelButton.host.SetFocusControl(_optionsPanel.dxUi->cancelButton.button);
                return _optionsPanel.dxUi->cancelButton.host.GetFocusControl() == _optionsPanel.dxUi->cancelButton.button;
            }
            return false;
        case CompareDirectoriesOptionsDebugFocusTarget::None:
            if (_optionsPanel.dxUi->body.hostHwnd && IsWindow(_optionsPanel.dxUi->body.hostHwnd.get()) != FALSE)
            {
                ::SetFocus(_optionsPanel.dxUi->body.hostHwnd.get());
            }
            _optionsPanel.dxUi->body.host.SetFocusControl(nullptr);
            _optionsPanel.dxUi->okButton.host.SetFocusControl(nullptr);
            _optionsPanel.dxUi->cancelButton.host.SetFocusControl(nullptr);
            return _optionsPanel.dxUi->body.host.GetFocusControl() == nullptr && _optionsPanel.dxUi->okButton.host.GetFocusControl() == nullptr &&
                   _optionsPanel.dxUi->cancelButton.host.GetFocusControl() == nullptr;
        default: return false;
    }

    if (! focusControl || ! focusControl->IsVisible() || ! focusControl->IsEnabled())
    {
        return false;
    }

    if (_optionsPanel.dxUi->body.hostHwnd && IsWindow(_optionsPanel.dxUi->body.hostHwnd.get()) != FALSE)
    {
        ::SetFocus(_optionsPanel.dxUi->body.hostHwnd.get());
    }
    _optionsPanel.dxUi->body.host.SetFocusControl(focusControl);
    static_cast<void>(EnsureOptionsDxBodyControlVisible(focusControl));
    return _optionsPanel.dxUi->body.host.GetFocusControl() == focusControl;
}

bool CompareDirectoriesWindow::DebugSetOptionsIgnoreFilesEnabled(const bool enabled) noexcept
{
    if (! _optionsPanel.dlg || ! _optionsPanel.dxUi)
    {
        return false;
    }

    _optionsPanel.draft.ignoreFiles = enabled;
    _optionsPanel.draftLoaded       = true;
    ApplyOptionsDraftInterlocks();
    UpdateOptionsVisibility();

    CompareDirectoriesOptionsDebugSnapshot snapshot{};
    return DebugGetOptionsSnapshot(snapshot) && snapshot.optionsDialogVisible;
}

bool CompareDirectoriesWindow::DebugSetOptionsIgnoreDirectoriesEnabled(const bool enabled) noexcept
{
    if (! _optionsPanel.dlg || ! _optionsPanel.dxUi)
    {
        return false;
    }

    _optionsPanel.draft.ignoreDirectories = enabled;
    _optionsPanel.draftLoaded             = true;
    ApplyOptionsDraftInterlocks();
    UpdateOptionsVisibility();

    CompareDirectoriesOptionsDebugSnapshot snapshot{};
    return DebugGetOptionsSnapshot(snapshot) && snapshot.optionsDialogVisible;
}

HWND CompareDirectoriesWindow::DebugGetOptionsDialogHandle() const noexcept
{
    if (! _optionsPanel.dlg || IsWindow(_optionsPanel.dlg.get()) == FALSE)
    {
        return nullptr;
    }

    return _optionsPanel.dlg.get();
}

bool CompareDirectoriesWindow::DebugScrollOptionsBodyPages(const int pageDelta) noexcept
{
    if (! _optionsPanel.ui.host || IsWindow(_optionsPanel.ui.host) == FALSE || pageDelta == 0 || _optionsPanel.scrollMax <= 0)
    {
        return false;
    }

    RECT hostClient{};
    if (! GetClientRect(_optionsPanel.ui.host, &hostClient))
    {
        return false;
    }

    const int pageHeight = std::max(1l, hostClient.bottom - hostClient.top);
    const int newOffset  = std::clamp(_optionsPanel.scrollOffset + (pageDelta * pageHeight), 0, _optionsPanel.scrollMax);
    if (newOffset == _optionsPanel.scrollOffset)
    {
        return false;
    }

    _optionsPanel.scrollOffset = newOffset;
    LayoutOptionsControls();
    return true;
}

bool CompareDirectoriesWindow::DebugGetOptionsTargetHostAndClientRect(const ::CompareDirectoriesOptionsDebugFocusTarget target,
                                                                      HWND& outHost,
                                                                      RECT& outRect) const noexcept
{
    outHost = nullptr;
    outRect = {};

    if (! _optionsPanel.dxUi)
    {
        return false;
    }

    const auto assignBodyRect = [&](RedSalamander::DxUi::Control* control) noexcept
    {
        if (! control || ! _optionsPanel.dxUi->body.hostHwnd || IsWindow(_optionsPanel.dxUi->body.hostHwnd.get()) == FALSE || ! control->IsVisible())
        {
            return false;
        }

        const auto bounds = control->GetBounds();
        outHost           = _optionsPanel.dxUi->body.hostHwnd.get();
        outRect.left      = static_cast<LONG>(std::lround(_optionsPanel.dxUi->body.host.DipsToPixels(bounds.left)));
        outRect.top       = static_cast<LONG>(std::lround(_optionsPanel.dxUi->body.host.DipsToPixels(bounds.top)));
        outRect.right     = static_cast<LONG>(std::lround(_optionsPanel.dxUi->body.host.DipsToPixels(bounds.right)));
        outRect.bottom    = static_cast<LONG>(std::lround(_optionsPanel.dxUi->body.host.DipsToPixels(bounds.bottom)));
        return true;
    };

    const auto assignButtonRect = [&](const OptionsButtonDx& buttonDx) noexcept
    {
        if (! buttonDx.button || ! buttonDx.hostHwnd || IsWindow(buttonDx.hostHwnd.get()) == FALSE || ! buttonDx.button->IsVisible())
        {
            return false;
        }

        const auto bounds = buttonDx.button->GetBounds();
        outHost           = buttonDx.hostHwnd.get();
        outRect.left      = static_cast<LONG>(std::lround(buttonDx.host.DipsToPixels(bounds.left)));
        outRect.top       = static_cast<LONG>(std::lround(buttonDx.host.DipsToPixels(bounds.top)));
        outRect.right     = static_cast<LONG>(std::lround(buttonDx.host.DipsToPixels(bounds.right)));
        outRect.bottom    = static_cast<LONG>(std::lround(buttonDx.host.DipsToPixels(bounds.bottom)));
        return true;
    };

    const auto& body = _optionsPanel.dxUi->body;
    switch (target)
    {
        case CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle: return assignBodyRect(body.compareSubdirectories.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::CompareSizeToggle: return assignBodyRect(body.compareSize.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::CompareDateTimeToggle: return assignBodyRect(body.compareDateTime.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::CompareAttributesToggle: return assignBodyRect(body.compareAttributes.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::CompareContentToggle: return assignBodyRect(body.compareContent.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirAttributesToggle: return assignBodyRect(body.compareSubdirAttributes.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle: return assignBodyRect(body.selectSubdirsOnlyInOnePane.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::KeepIdenticalItemsToggle: return assignBodyRect(body.keepIdenticalItems.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle: return assignBodyRect(body.ignoreFiles.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit: return assignBodyRect(body.ignoreFiles.edit);
        case CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle: return assignBodyRect(body.ignoreDirectories.toggle);
        case CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesEdit: return assignBodyRect(body.ignoreDirectories.edit);
        case CompareDirectoriesOptionsDebugFocusTarget::OkButton: return assignButtonRect(_optionsPanel.dxUi->okButton);
        case CompareDirectoriesOptionsDebugFocusTarget::CancelButton: return assignButtonRect(_optionsPanel.dxUi->cancelButton);
        case CompareDirectoriesOptionsDebugFocusTarget::None:
        default: return false;
    }
}
#endif

void CompareDirectoriesWindow::LoadOptionsControlsFromSettings() noexcept
{
    if (! _optionsPanel.dlg || ! _optionsPanel.ui.host)
    {
        return;
    }

    _optionsPanel.draft       = GetEffectiveCompareSettings();
    _optionsPanel.draftLoaded = true;
    ApplyOptionsDraftInterlocks();

    SyncOptionsDxToggles();
    SyncOptionsDxEdits();
    UpdateOptionsVisibility();
}
void CompareDirectoriesWindow::SaveOptionsControlsToSettings() noexcept
{
    if (! _optionsPanel.dlg || ! _settings || ! _optionsPanel.ui.host)
    {
        return;
    }

    _settings->compareDirectories         = ReadOptionsControlsToSettings();
    _optionsPanel.staleFromExternalReload = false;
}

bool CompareDirectoriesWindow::ResolveOptionsStaleSaveConflict(HWND dlg) noexcept
{
    if (! _optionsPanel.staleFromExternalReload)
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

    _optionsPanel.staleFromExternalReload = false;
    return true;
}

void CompareDirectoriesWindow::ReloadOptionsDialogFromDisk() noexcept
{
    _optionsPanel.staleFromExternalReload = false;
    LoadOptionsControlsFromSettings();
    UpdateViewMenuChecks();
}
void CompareDirectoriesWindow::UpdateOptionsVisibility() noexcept
{
    if (! _optionsPanel.dlg || ! _optionsPanel.ui.host)
    {
        return;
    }

    LayoutOptionsControls();
    if (_optionsPanel.dxUi && _optionsPanel.dxUi->body.hostHwnd && IsWindow(_optionsPanel.dxUi->body.hostHwnd.get()) != FALSE)
    {
        _optionsPanel.dxUi->body.host.RefreshAccessibilitySnapshot();
    }
    RedrawWindow(_optionsPanel.ui.host, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
}

} // namespace CompareDirectoriesWindowInternal
