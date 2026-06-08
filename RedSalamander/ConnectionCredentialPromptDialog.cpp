#include "Framework.h"

#include "ConnectionCredentialPromptDialog.h"

#include <algorithm>
#include <atomic>
#include <cwctype>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/resource.h>
#pragma warning(pop)

#include "DxUi/DxUi.h"
#include "Helpers.h"
#include "UiMetrics.h"
#include "WindowMessages.h"
#include "resource.h"

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::Control;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kPromptWindowClassName[] = L"RedSalamander.ConnectionCredentialPromptWindow";
constexpr UINT kPromptDeferredCloseMessage = WM_APP + 1u;
std::atomic<HWND> g_connectionCredentialPromptWindow{nullptr};

#ifdef ENABLE_TESTS
enum class PromptDebugCommand : WPARAM
{
    GetSnapshot = 1,
    SetUserName,
    SetSecret,
    GetToggleSecretButtonRect,
    ToggleSecretVisibility,
    Confirm,
};

struct PromptDebugText
{
    const wchar_t* text = nullptr;
    size_t length       = 0u;
};

struct PromptDebugHostRect
{
    HWND host = nullptr;
    RECT rect{};
};
#endif

[[nodiscard]] HWND NormalizeOwnerWindow(HWND ownerWindow) noexcept
{
    if (! ownerWindow || IsWindow(ownerWindow) == FALSE)
    {
        return nullptr;
    }
    return GetAncestor(ownerWindow, GA_ROOT);
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0u;
    while (start < text.size() && std::iswspace(static_cast<wint_t>(text[start])) != 0)
    {
        ++start;
    }

    size_t end = text.size();
    while (end > start && std::iswspace(static_cast<wint_t>(text[end - 1u])) != 0)
    {
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

void CenterWindowOnOwner(HWND window, HWND owner) noexcept
{
    if (! window || IsWindow(window) == FALSE || ! owner || IsWindow(owner) == FALSE)
    {
        return;
    }

    RECT ownerRect{};
    RECT windowRect{};
    if (GetWindowRect(owner, &ownerRect) == FALSE || GetWindowRect(window, &windowRect) == FALSE)
    {
        return;
    }

    const int x = ownerRect.left + (((ownerRect.right - ownerRect.left) - (windowRect.right - windowRect.left)) / 2);
    const int y = ownerRect.top + (((ownerRect.bottom - ownerRect.top) - (windowRect.bottom - windowRect.top)) / 2);
    SetWindowPos(window, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
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
        if (const auto isActuallyVisibleChildWindow =
                [](HWND window) noexcept
        {
            if (! window || IsWindowVisible(window) == FALSE)
            {
                return false;
            }

            wil::unique_hrgn region(CreateRectRgn(0, 0, 0, 0));
            if (region)
            {
                const int regionType = GetWindowRgn(window, region.get());
                if (regionType == NULLREGION)
                {
                    return false;
                }
            }

            return true;
        };
            isActuallyVisibleChildWindow(child))
        {
            ++counterRef.count;
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&counter)));

    return counter.count;
}

[[nodiscard]] ThemePalette MakeDxPalette(const AppTheme& theme) noexcept
{
    const auto mix = [](const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, float t) noexcept
    {
        const float clamped = std::clamp(t, 0.0f, 1.0f);
        return D2D1::ColorF(a.r + ((b.r - a.r) * clamped), a.g + ((b.g - a.g) * clamped), a.b + ((b.b - a.b) * clamped), a.a + ((b.a - a.a) * clamped));
    };

    ThemePalette palette          = RedSalamander::DxUi::MakeDefaultThemePalette(theme.dark);
    palette.dark                  = theme.dark;
    palette.highContrast          = theme.highContrast;
    palette.accent                = theme.accent;
    palette.windowBackground      = ColorFromCOLORREF(theme.windowBackground);
    palette.surfaceBackground     = ColorFromCOLORREF(UiMetrics::GetControlSurfaceColor(theme));
    palette.headerBackground      = ColorFromCOLORREF(theme.menu.background);
    palette.headerHovered         = mix(palette.headerBackground, palette.accent, theme.dark ? 0.22f : 0.10f);
    palette.headerPressed         = mix(palette.headerBackground, palette.accent, theme.dark ? 0.30f : 0.16f);
    palette.border                = ColorFromCOLORREF(theme.menu.border);
    palette.gridLine              = ColorFromCOLORREF(theme.menu.border);
    palette.text                  = ColorFromCOLORREF(theme.menu.text);
    palette.subduedText           = ColorFromCOLORREF(theme.menu.shortcutText);
    palette.disabledText          = ColorFromCOLORREF(theme.menu.disabledText);
    palette.selectionFill         = ColorFromCOLORREF(theme.menu.selectionBg);
    palette.selectionText         = ColorFromCOLORREF(theme.menu.selectionText);
    palette.selectionInactiveFill = D2D1::ColorF(palette.selectionFill.r, palette.selectionFill.g, palette.selectionFill.b, theme.highContrast ? 1.0f : 0.55f);
    palette.focusStroke           = theme.folderView.focusBorder;
    palette.hoverFill             = D2D1::ColorF(palette.accent.r, palette.accent.g, palette.accent.b, theme.dark ? 0.18f : 0.10f);
    palette.buttonFill            = ColorFromCOLORREF(theme.menu.background);
    palette.buttonBorder          = ColorFromCOLORREF(theme.menu.border);
    palette.buttonHotFill         = palette.headerHovered;
    palette.buttonPressedFill     = palette.headerPressed;
    palette.inputFill             = ColorFromCOLORREF(UiMetrics::GetControlSurfaceColor(theme));
    palette.inputBorder           = ColorFromCOLORREF(theme.menu.border);
    palette.scrollbarTrack        = theme.fileOperations.scrollbarTrack;
    palette.scrollbarThumb        = theme.fileOperations.scrollbarThumb;
    palette.scrollbarThumbHot =
        D2D1::ColorF(palette.scrollbarThumb.r, palette.scrollbarThumb.g, palette.scrollbarThumb.b, std::min(1.0f, palette.scrollbarThumb.a + 0.10f));
    palette.infoFill    = theme.folderView.infoBackground;
    palette.infoText    = theme.folderView.infoText;
    palette.warningFill = theme.folderView.warningBackground;
    palette.warningText = theme.folderView.warningText;
    palette.errorFill   = theme.folderView.errorBackground;
    palette.errorText   = theme.folderView.errorText;
    return palette;
}

[[nodiscard]] HRESULT EnsurePromptWindowClass() noexcept;

class ConnectionCredentialPromptWindow final
{
public:
    ConnectionCredentialPromptWindow(HWND ownerWindow,
                                     const AppTheme& theme,
                                     std::wstring caption,
                                     std::wstring message,
                                     std::wstring secretLabel,
                                     bool showUserName,
                                     bool allowEmptySecret,
                                     std::wstring initialUserName) noexcept;
    ConnectionCredentialPromptWindow(const ConnectionCredentialPromptWindow&)            = delete;
    ConnectionCredentialPromptWindow& operator=(const ConnectionCredentialPromptWindow&) = delete;
    ConnectionCredentialPromptWindow(ConnectionCredentialPromptWindow&&)                 = delete;
    ConnectionCredentialPromptWindow& operator=(ConnectionCredentialPromptWindow&&)      = delete;

    [[nodiscard]] HRESULT ShowModal(std::wstring& userNameOut, std::wstring& secretOut) noexcept;
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;

private:
    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept;
    void BuildUi() noexcept;
    void ApplyTheme() noexcept;
    void Layout() noexcept;
    void ClearValidation() noexcept;
    void ShowValidation(UINT resourceId) noexcept;
    [[maybe_unused]] [[nodiscard]] void ToggleSecretVisibility() noexcept;
    void SetSecretVisibility(bool visible) noexcept;
    void CloseDeferred() noexcept;
    void Confirm() noexcept;
    void Cancel() noexcept;
    [[nodiscard]] LRESULT WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
#ifdef ENABLE_TESTS
    [[nodiscard]] LRESULT OnDebugMessage(WPARAM command, LPARAM payload) noexcept;
#endif

    HWND _ownerWindow = nullptr;
    AppTheme _theme{};
    ThemePalette _palette{};
    std::wstring _caption;
    std::wstring _message;
    std::wstring _secretLabelText;
    std::wstring _initialUserName;
    std::wstring _validationText;
    std::wstring _acceptedUserName;
    std::wstring _acceptedSecret;
    wil::unique_hwnd _hWnd;
    WindowHost _dxHost;
    std::unique_ptr<Panel> _rootStorage;
    Panel* _root                = nullptr;
    Label* _messageLabel        = nullptr;
    Label* _userLabel           = nullptr;
    TextField* _userField       = nullptr;
    Label* _secretLabel         = nullptr;
    TextField* _secretField     = nullptr;
    Toggle* _toggleSecretButton = nullptr;
    Label* _validationLabel     = nullptr;
    Button* _okButton           = nullptr;
    Button* _cancelButton       = nullptr;
    bool _showUserName          = false;
    bool _allowEmptySecret      = false;
    bool _secretVisible         = false;
    bool _closing               = false;
    bool _done                  = false;
    HRESULT _result             = E_FAIL;
};

HRESULT EnsurePromptWindowClass() noexcept
{
    static ATOM atom = 0;
    if (atom != 0)
    {
        return S_OK;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = ConnectionCredentialPromptWindow::WndProc;
    wc.hInstance     = GetModuleHandleW(nullptr);
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = kPromptWindowClassName;
    wc.style         = CS_DBLCLKS;

    atom = RegisterClassExW(&wc);
    if (atom != 0)
    {
        return S_OK;
    }

    const DWORD lastError = GetLastError();
    if (lastError == ERROR_CLASS_ALREADY_EXISTS)
    {
        atom = 1;
        return S_OK;
    }

    Debug::ErrorWithLastError(L"ConnectionCredentialPrompt: RegisterClassExW failed.");
    return HRESULT_FROM_WIN32(lastError);
}

ConnectionCredentialPromptWindow::ConnectionCredentialPromptWindow(HWND ownerWindow,
                                                                   const AppTheme& theme,
                                                                   std::wstring caption,
                                                                   std::wstring message,
                                                                   std::wstring secretLabel,
                                                                   bool showUserName,
                                                                   bool allowEmptySecret,
                                                                   std::wstring initialUserName) noexcept
    : _ownerWindow(NormalizeOwnerWindow(ownerWindow)),
      _theme(theme),
      _caption(std::move(caption)),
      _message(std::move(message)),
      _secretLabelText(std::move(secretLabel)),
      _initialUserName(std::move(initialUserName)),
      _showUserName(showUserName),
      _allowEmptySecret(allowEmptySecret)
{
}

HRESULT ConnectionCredentialPromptWindow::ShowModal(std::wstring& userNameOut, std::wstring& secretOut) noexcept
{
    userNameOut.clear();
    secretOut.clear();

    const HRESULT classHr = EnsurePromptWindowClass();
    if (FAILED(classHr))
    {
        return classHr;
    }

    const DWORD style        = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    const DWORD exStyle      = WS_EX_DLGMODALFRAME;
    const UINT dpi           = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();
    const int clientWidthPx  = UiMetrics::ScaleDip(dpi, 520);
    const int clientHeightPx = UiMetrics::ScaleDip(dpi, _showUserName ? 320 : 262);

    RECT bounds{0, 0, clientWidthPx, clientHeightPx};
    if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
    {
        const DWORD lastError = Debug::ErrorWithLastError(L"ConnectionCredentialPrompt: AdjustWindowRectExForDpi failed.");
        return HRESULT_FROM_WIN32(lastError);
    }

    const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
    if (restoreOwnerEnabled)
    {
        EnableWindow(_ownerWindow, FALSE);
    }
    const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled] noexcept
    {
        if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            EnableWindow(_ownerWindow, TRUE);
            SetActiveWindow(_ownerWindow);
        }
    });

    const HWND hwnd = CreateWindowExW(exStyle,
                                      kPromptWindowClassName,
                                      _caption.c_str(),
                                      style,
                                      CW_USEDEFAULT,
                                      CW_USEDEFAULT,
                                      bounds.right - bounds.left,
                                      bounds.bottom - bounds.top,
                                      _ownerWindow,
                                      nullptr,
                                      GetModuleHandleW(nullptr),
                                      this);
    if (! hwnd)
    {
        const DWORD lastError = Debug::ErrorWithLastError(L"ConnectionCredentialPrompt: CreateWindowExW failed.");
        return HRESULT_FROM_WIN32(lastError);
    }
    if (! _hWnd)
    {
        _hWnd.reset(hwnd);
    }

    CenterWindowOnOwner(_hWnd.get(), _ownerWindow);
    ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
    UpdateWindow(_hWnd.get());
    SetForegroundWindow(_hWnd.get());

    MSG msg{};
    while (! _done)
    {
        const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
        if (getMessageResult == -1)
        {
            const DWORD lastError = Debug::ErrorWithLastError(L"ConnectionCredentialPrompt: GetMessageW failed.");
            return HRESULT_FROM_WIN32(lastError);
        }
        if (getMessageResult == 0)
        {
            _done   = true;
            _result = S_FALSE;
            PostQuitMessage(static_cast<int>(msg.wParam));
            break;
        }

        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    if (_result == S_OK)
    {
        userNameOut = _acceptedUserName;
        secretOut   = _acceptedSecret;
    }

    return _result;
}

bool ConnectionCredentialPromptWindow::OnCreate(HWND hwnd) noexcept
{
    if (! _dxHost.Attach(hwnd))
    {
        Debug::Error(L"ConnectionCredentialPrompt: failed to attach DxUi host.");
        return false;
    }

    BuildUi();
    ApplyTheme();
    Layout();
    _dxHost.SetFocusControl(_showUserName ? static_cast<Control*>(_userField) : static_cast<Control*>(_secretField));
    return true;
}

void ConnectionCredentialPromptWindow::BuildUi() noexcept
{
    if (_root != nullptr)
    {
        return;
    }

    _rootStorage = std::make_unique<Panel>();
    _root        = _rootStorage.get();

    _messageLabel = _root->AddChild<Label>(_message);
    _messageLabel->SetMultiline(true);
    _messageLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

    _userLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_USER));
    _userLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
    _userLabel->SetMnemonic(L'U');
    _userLabel->SetVisible(_showUserName);

    _userField = _root->AddChild<TextField>(_initialUserName);
    _userField->SetVisible(_showUserName);
    _userField->SetOnTextChanged([this](std::wstring_view) { ClearValidation(); });
    _userLabel->SetMnemonicTarget(_userField);

    _secretLabel = _root->AddChild<Label>(_secretLabelText);
    _secretLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
    _secretLabel->SetMnemonic(L'P');

    _secretField = _root->AddChild<TextField>();
    _secretField->SetMasked(true);
    _secretField->SetOnTextChanged([this](std::wstring_view) { ClearValidation(); });
    _secretLabel->SetMnemonicTarget(_secretField);

    _toggleSecretButton = _root->AddChild<Toggle>(LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_SHOW_SECRET));
    _toggleSecretButton->SetChecked(_secretVisible);
    _toggleSecretButton->SetOnToggled([this](bool checked) { SetSecretVisibility(checked); });

    _validationLabel = _root->AddChild<Label>();
    _validationLabel->SetMultiline(true);
    _validationLabel->SetFontRole(FontRole::Small);
    _validationLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

    _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
    _okButton->SetPrimary(true);
    _okButton->SetMnemonic(L'O');
    _okButton->SetOnClick([this] { Confirm(); });

    _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
    _cancelButton->SetMnemonic(L'C');
    _cancelButton->SetOnClick([this] { Cancel(); });

    _dxHost.SetRoot(std::move(_rootStorage));
    _dxHost.SetDefaultButton(_okButton);
    _dxHost.SetCancelButton(_cancelButton);
}

void ConnectionCredentialPromptWindow::ApplyTheme() noexcept
{
    _palette = MakeDxPalette(_theme);
    _dxHost.SetTheme(_palette);
    if (_validationLabel)
    {
        _validationLabel->SetTextColor(_palette.errorText);
    }
    if (_hWnd)
    {
        ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
    }
}

void ConnectionCredentialPromptWindow::Layout() noexcept
{
    if (! _root)
    {
        return;
    }

    const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
    _root->SetBounds(client);

    constexpr float kMarginDip            = 16.0f;
    constexpr float kFieldHeightDip       = 32.0f;
    constexpr float kButtonHeightDip      = 34.0f;
    constexpr float kButtonWidthDip       = 104.0f;
    constexpr float kSecretToggleWidthDip = 88.0f;
    constexpr float kLabelHeightDip       = 20.0f;
    constexpr float kMessageHeightDip     = 42.0f;
    constexpr float kGapDip               = 10.0f;
    constexpr float kButtonGapDip         = 8.0f;

    const float left  = client.left + kMarginDip;
    const float right = client.right - kMarginDip;
    float y           = client.top + kMarginDip;

    _messageLabel->SetBounds(D2D1::RectF(left, y, right, y + kMessageHeightDip));
    y += kMessageHeightDip + kGapDip;

    if (_showUserName)
    {
        _userLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        y += kLabelHeightDip + 4.0f;
        _userField->SetBounds(D2D1::RectF(left, y, right, y + kFieldHeightDip));
        y += kFieldHeightDip + kGapDip;
    }

    _secretLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
    y += kLabelHeightDip + 4.0f;

    const float toggleLeft = right - kSecretToggleWidthDip;
    _secretField->SetBounds(D2D1::RectF(left, y, toggleLeft - kButtonGapDip, y + kFieldHeightDip));
    _toggleSecretButton->SetBounds(D2D1::RectF(toggleLeft, y, right, y + kFieldHeightDip));

    const float buttonsTop = client.bottom - kMarginDip - kButtonHeightDip;
    _validationLabel->SetBounds(D2D1::RectF(left, y + kFieldHeightDip + 8.0f, right, buttonsTop - 8.0f));

    const float cancelLeft = right - kButtonWidthDip;
    const float okLeft     = cancelLeft - kButtonGapDip - kButtonWidthDip;
    _okButton->SetBounds(D2D1::RectF(okLeft, buttonsTop, okLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
    _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
}

void ConnectionCredentialPromptWindow::ClearValidation() noexcept
{
    if (_validationText.empty())
    {
        return;
    }

    _validationText.clear();
    if (_validationLabel)
    {
        _validationLabel->SetText({});
    }
    _dxHost.Invalidate();
}

void ConnectionCredentialPromptWindow::ShowValidation(UINT resourceId) noexcept
{
    _validationText = LoadStringResource(nullptr, resourceId);
    if (_validationLabel)
    {
        _validationLabel->SetText(_validationText);
    }
    _dxHost.Invalidate();
    MessageBeep(MB_ICONWARNING);
}

void ConnectionCredentialPromptWindow::ToggleSecretVisibility() noexcept
{
    SetSecretVisibility(! _secretVisible);
}

void ConnectionCredentialPromptWindow::SetSecretVisibility(const bool visible) noexcept
{
    _secretVisible = visible;
    if (_secretField)
    {
        _secretField->SetMasked(! _secretVisible);
    }
    if (_toggleSecretButton)
    {
        _toggleSecretButton->SetChecked(_secretVisible);
    }
    _dxHost.Invalidate();
}

void ConnectionCredentialPromptWindow::Confirm() noexcept
{
    if (_closing)
    {
        return;
    }

    ClearValidation();

    std::wstring userName;
    if (_showUserName)
    {
        userName = TrimWhitespace(_userField ? _userField->GetText() : std::wstring_view{});
        if (userName.empty())
        {
            ShowValidation(IDS_CONNECTIONS_ERR_PROMPT_USER_REQUIRED);
            _dxHost.SetFocusControl(_userField);
            return;
        }
    }

    const std::wstring secret = _secretField ? std::wstring(_secretField->GetText()) : std::wstring{};
    if (! _allowEmptySecret && secret.empty())
    {
        ShowValidation(IDS_CONNECTIONS_ERR_PROMPT_PASSWORD_REQUIRED);
        _dxHost.SetFocusControl(_secretField);
        return;
    }

    _acceptedUserName = std::move(userName);
    _acceptedSecret   = secret;
    _result           = S_OK;
    _closing          = true;
    CloseDeferred();
}

void ConnectionCredentialPromptWindow::Cancel() noexcept
{
    if (_closing)
    {
        return;
    }

    _acceptedUserName.clear();
    _acceptedSecret.clear();
    _result  = S_FALSE;
    _closing = true;
    CloseDeferred();
}

void ConnectionCredentialPromptWindow::CloseDeferred() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (PostMessageW(_hWnd.get(), kPromptDeferredCloseMessage, 0, 0) == 0)
    {
        Debug::ErrorWithLastError(L"Connection credential prompt failed to post deferred close.");
        _hWnd.reset();
    }
}

#ifdef ENABLE_TESTS
LRESULT ConnectionCredentialPromptWindow::OnDebugMessage(WPARAM command, LPARAM payload) noexcept
{
    switch (static_cast<PromptDebugCommand>(command))
    {
        case PromptDebugCommand::GetSnapshot:
        {
            auto* snapshot = reinterpret_cast<ConnectionCredentialPromptDebugSnapshot*>(payload);
            if (! snapshot)
            {
                return FALSE;
            }

            snapshot->usesDxUiHost            = _dxHost.GetRoot() == _root;
            snapshot->showUserName            = _showUserName;
            snapshot->allowEmptySecret        = _allowEmptySecret;
            snapshot->secretVisible           = _secretVisible;
            snapshot->themeDark               = _theme.dark;
            snapshot->themeHighContrast       = _theme.highContrast;
            snapshot->themeRainbow            = _theme.menu.rainbowMode;
            snapshot->secretLength            = _secretField ? _secretField->GetText().size() : 0u;
            snapshot->visibleChildWindowCount = CountVisibleChildWindows(_hWnd.get());
            snapshot->toggleSecretChecked     = _toggleSecretButton ? _toggleSecretButton->IsChecked() : false;
            snapshot->toggleSecretPressed     = _toggleSecretButton ? _toggleSecretButton->DebugIsPressed() : false;
            snapshot->hostHasCapture          = _hWnd && GetCapture() == _hWnd.get();
            snapshot->userNameText            = _userField ? std::wstring(_userField->GetText()) : std::wstring{};
            snapshot->validationText          = _validationText;

            const Control* focusControl = _dxHost.GetFocusControl();
            snapshot->focusTarget       = ConnectionCredentialPromptDebugFocusTarget::None;
            if (focusControl == _userField)
            {
                snapshot->focusTarget = ConnectionCredentialPromptDebugFocusTarget::UserField;
            }
            else if (focusControl == _secretField)
            {
                snapshot->focusTarget = ConnectionCredentialPromptDebugFocusTarget::SecretField;
            }
            else if (focusControl == _toggleSecretButton)
            {
                snapshot->focusTarget = ConnectionCredentialPromptDebugFocusTarget::ToggleSecretButton;
            }
            else if (focusControl == _okButton)
            {
                snapshot->focusTarget = ConnectionCredentialPromptDebugFocusTarget::OkButton;
            }
            else if (focusControl == _cancelButton)
            {
                snapshot->focusTarget = ConnectionCredentialPromptDebugFocusTarget::CancelButton;
            }
            return TRUE;
        }
        case PromptDebugCommand::SetUserName:
        {
            if (! _showUserName || ! _userField)
            {
                return FALSE;
            }
            const auto* text = reinterpret_cast<const PromptDebugText*>(payload);
            if (! text)
            {
                return FALSE;
            }

            _userField->SetText(text->text ? std::wstring(text->text, text->length) : std::wstring{});
            ClearValidation();
            _dxHost.SetFocusControl(_userField);
            _dxHost.Invalidate();
            return TRUE;
        }
        case PromptDebugCommand::SetSecret:
        {
            if (! _secretField)
            {
                return FALSE;
            }
            const auto* text = reinterpret_cast<const PromptDebugText*>(payload);
            if (! text)
            {
                return FALSE;
            }

            _secretField->SetText(text->text ? std::wstring(text->text, text->length) : std::wstring{});
            ClearValidation();
            _dxHost.SetFocusControl(_secretField);
            _dxHost.Invalidate();
            return TRUE;
        }
        case PromptDebugCommand::GetToggleSecretButtonRect:
        {
            auto* hostRect = reinterpret_cast<PromptDebugHostRect*>(payload);
            if (! hostRect)
            {
                return FALSE;
            }

            hostRect->host = nullptr;
            hostRect->rect = {};
            if (! _toggleSecretButton || ! _toggleSecretButton->IsVisible() || ! _toggleSecretButton->IsEnabled() || ! _hWnd)
            {
                return FALSE;
            }

            hostRect->host           = _hWnd.get();
            const D2D1_RECT_F bounds = _toggleSecretButton->GetBounds();
            hostRect->rect.left      = static_cast<LONG>(std::lround(_dxHost.DipsToPixels(bounds.left)));
            hostRect->rect.top       = static_cast<LONG>(std::lround(_dxHost.DipsToPixels(bounds.top)));
            hostRect->rect.right     = static_cast<LONG>(std::lround(_dxHost.DipsToPixels(bounds.right)));
            hostRect->rect.bottom    = static_cast<LONG>(std::lround(_dxHost.DipsToPixels(bounds.bottom)));
            return TRUE;
        }
        case PromptDebugCommand::ToggleSecretVisibility:
            ToggleSecretVisibility();
            _dxHost.SetFocusControl(_toggleSecretButton);
            return TRUE;
        case PromptDebugCommand::Confirm: Confirm(); return TRUE;
    }

    return FALSE;
}
#endif

LRESULT ConnectionCredentialPromptWindow::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_NCDESTROY)
    {
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
        _dxHost.ReleaseMouseCapture();
        _dxHost.Detach();
        if (_hWnd.get() == hwnd)
        {
            static_cast<void>(_hWnd.release());
        }
        g_connectionCredentialPromptWindow.store(nullptr, std::memory_order_release);
        if (_result == E_FAIL)
        {
            _result = S_FALSE;
        }
        _closing = false;
        _done    = true;
        return 0;
    }

    bool dxHandled         = false;
    const LRESULT dxResult = _dxHost.HandleMessage(hwnd, message, wParam, lParam, dxHandled);
    if (dxHandled)
    {
        if (message == WM_SIZE)
        {
            Layout();
        }
        return dxResult;
    }

    switch (message)
    {
        case WM_CREATE: return OnCreate(hwnd) ? 0 : -1;
        case WM_SIZE: Layout(); return 0;
        case WM_DPICHANGED:
        {
            if (const RECT* suggested = reinterpret_cast<const RECT*>(lParam))
            {
                SetWindowPos(hwnd,
                             nullptr,
                             suggested->left,
                             suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            ApplyTheme();
            Layout();
            return 0;
        }
        case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, _theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
        case WndMsg::kConnectionCredentialPromptApplyTheme:
        {
            const auto* theme = reinterpret_cast<const AppTheme*>(lParam);
            if (! theme)
            {
                return FALSE;
            }

            _theme = *theme;
            ApplyTheme();
            InvalidateRect(hwnd, nullptr, FALSE);
            return TRUE;
        }
        case kPromptDeferredCloseMessage:
            if (_hWnd.get() == hwnd)
            {
                _hWnd.reset();
            }
            return 0;
        case WM_CLOSE:
            if (_closing)
            {
                return 0;
            }
            Cancel();
            return 0;
#ifdef ENABLE_TESTS
        case WndMsg::kConnectionCredentialPromptDebug: return OnDebugMessage(wParam, lParam);
#endif
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

LRESULT CALLBACK ConnectionCredentialPromptWindow::WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    auto* self = reinterpret_cast<ConnectionCredentialPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (message == WM_NCCREATE)
    {
        const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
        self               = create ? reinterpret_cast<ConnectionCredentialPromptWindow*>(create->lpCreateParams) : nullptr;
        if (! self)
        {
            return FALSE;
        }

        self->_hWnd.reset(hwnd);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        g_connectionCredentialPromptWindow.store(hwnd, std::memory_order_release);
    }

    if (! self)
    {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    return self->WindowProc(hwnd, message, wParam, lParam);
}

} // namespace

HRESULT PromptForConnectionSecret(HWND ownerWindow,
                                  const AppTheme& theme,
                                  std::wstring_view caption,
                                  std::wstring_view message,
                                  std::wstring_view secretLabel,
                                  bool allowEmptySecret,
                                  std::wstring& secretOut) noexcept
{
    std::wstring userName;
    secretOut.clear();

    ConnectionCredentialPromptWindow window(
        ownerWindow, theme, std::wstring(caption), std::wstring(message), std::wstring(secretLabel), false, allowEmptySecret, {});
    const HRESULT hr = window.ShowModal(userName, secretOut);
    if (hr != S_OK)
    {
        secretOut.clear();
        return hr;
    }

    return S_OK;
}

HRESULT PromptForConnectionUserAndPassword(HWND ownerWindow,
                                           const AppTheme& theme,
                                           std::wstring_view caption,
                                           std::wstring_view message,
                                           std::wstring_view initialUserName,
                                           std::wstring& userNameOut,
                                           std::wstring& passwordOut) noexcept
{
    userNameOut.clear();
    passwordOut.clear();

    ConnectionCredentialPromptWindow window(ownerWindow,
                                            theme,
                                            std::wstring(caption),
                                            std::wstring(message),
                                            LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PASSWORD),
                                            true,
                                            false,
                                            std::wstring(initialUserName));
    const HRESULT hr = window.ShowModal(userNameOut, passwordOut);
    if (hr != S_OK)
    {
        userNameOut.clear();
        passwordOut.clear();
        return hr;
    }

    return S_OK;
}

HWND GetConnectionCredentialPromptDialogHandle() noexcept
{
    const HWND hwnd = g_connectionCredentialPromptWindow.load(std::memory_order_acquire);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

void UpdateConnectionCredentialPromptWindowsTheme(const AppTheme& theme) noexcept
{
    const HWND hwnd = GetConnectionCredentialPromptDialogHandle();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    static_cast<void>(SendMessageW(hwnd, WndMsg::kConnectionCredentialPromptApplyTheme, 0, reinterpret_cast<LPARAM>(&theme)));
}

#ifdef ENABLE_TESTS
bool DebugGetConnectionCredentialPromptSnapshot(ConnectionCredentialPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetConnectionCredentialPromptDialogHandle();
    return hwnd &&
           SendMessageW(hwnd, WndMsg::kConnectionCredentialPromptDebug, static_cast<WPARAM>(PromptDebugCommand::GetSnapshot), reinterpret_cast<LPARAM>(&out)) !=
               FALSE;
}

bool DebugSetConnectionCredentialPromptUserName(std::wstring_view text) noexcept
{
    const HWND hwnd = GetConnectionCredentialPromptDialogHandle();
    if (! hwnd)
    {
        return false;
    }

    const PromptDebugText payload{.text = text.data(), .length = text.size()};
    return SendMessageW(
               hwnd, WndMsg::kConnectionCredentialPromptDebug, static_cast<WPARAM>(PromptDebugCommand::SetUserName), reinterpret_cast<LPARAM>(&payload)) !=
           FALSE;
}

bool DebugSetConnectionCredentialPromptSecret(std::wstring_view text) noexcept
{
    const HWND hwnd = GetConnectionCredentialPromptDialogHandle();
    if (! hwnd)
    {
        return false;
    }

    const PromptDebugText payload{.text = text.data(), .length = text.size()};
    return SendMessageW(
               hwnd, WndMsg::kConnectionCredentialPromptDebug, static_cast<WPARAM>(PromptDebugCommand::SetSecret), reinterpret_cast<LPARAM>(&payload)) != FALSE;
}

bool DebugGetConnectionCredentialPromptToggleSecretButtonHostAndClientRect(HWND& outHost, RECT& outRect) noexcept
{
    outHost = nullptr;
    outRect = {};

    PromptDebugHostRect payload{};
    const HWND hwnd = GetConnectionCredentialPromptDialogHandle();
    if (! hwnd)
    {
        return false;
    }

    if (SendMessageW(hwnd,
                     WndMsg::kConnectionCredentialPromptDebug,
                     static_cast<WPARAM>(PromptDebugCommand::GetToggleSecretButtonRect),
                     reinterpret_cast<LPARAM>(&payload)) == FALSE)
    {
        return false;
    }

    outHost = payload.host;
    outRect = payload.rect;
    return outHost != nullptr;
}

bool DebugToggleConnectionCredentialPromptSecretVisibility() noexcept
{
    const HWND hwnd = GetConnectionCredentialPromptDialogHandle();
    return hwnd && SendMessageW(hwnd, WndMsg::kConnectionCredentialPromptDebug, static_cast<WPARAM>(PromptDebugCommand::ToggleSecretVisibility), 0) != FALSE;
}

bool DebugConfirmConnectionCredentialPrompt() noexcept
{
    const HWND hwnd = GetConnectionCredentialPromptDialogHandle();
    return hwnd && SendMessageW(hwnd, WndMsg::kConnectionCredentialPromptDebug, static_cast<WPARAM>(PromptDebugCommand::Confirm), 0) != FALSE;
}
#endif
