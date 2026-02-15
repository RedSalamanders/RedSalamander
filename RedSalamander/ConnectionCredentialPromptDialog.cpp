#include "Framework.h"

#include "ConnectionCredentialPromptDialog.h"

#include <algorithm>
#include <commctrl.h>
#include <cwctype>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "ThemedControls.h"
#include "ThemedInputFrames.h"
#include "resource.h"

namespace
{
constexpr UINT_PTR kSecretEditSubclassId = 1u;

struct DialogState
{
    DialogState()                              = default;
    DialogState(const DialogState&)            = delete;
    DialogState& operator=(const DialogState&) = delete;
    DialogState(DialogState&&)                 = delete;
    DialogState& operator=(DialogState&&)      = delete;
    ~DialogState()                             = default;

    AppTheme theme{};
    wil::unique_hbrush backgroundBrush;

    COLORREF inputBackgroundColor         = RGB(255, 255, 255);
    COLORREF inputFocusedBackgroundColor  = RGB(255, 255, 255);
    COLORREF inputDisabledBackgroundColor = RGB(255, 255, 255);
    wil::unique_hbrush inputBrush;
    wil::unique_hbrush inputFocusedBrush;
    wil::unique_hbrush inputDisabledBrush;

    ThemedInputFrames::FrameStyle inputFrameStyle{};
    wil::unique_hwnd userFrame;
    wil::unique_hwnd secretFrame;

    bool showUserName      = false;
    bool allowEmptySecret  = false;
    bool secretVisible     = false;
    bool showingValidation = false;

    std::wstring caption;
    std::wstring message;
    std::wstring secretLabel;
    std::wstring initialUserName;

    std::wstring userNameOut;
    std::wstring secretOut;
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

void CenterWindowOnOwner(HWND window, HWND owner) noexcept
{
    if (! window || ! owner)
    {
        return;
    }

    RECT ownerRect{};
    RECT windowRect{};
    if (GetWindowRect(owner, &ownerRect) == 0 || GetWindowRect(window, &windowRect) == 0)
    {
        return;
    }

    const int ownerW  = ownerRect.right - ownerRect.left;
    const int ownerH  = ownerRect.bottom - ownerRect.top;
    const int windowW = windowRect.right - windowRect.left;
    const int windowH = windowRect.bottom - windowRect.top;

    const int x = ownerRect.left + (ownerW - windowW) / 2;
    const int y = ownerRect.top + (ownerH - windowH) / 2;
    SetWindowPos(window, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0;
    while (start < text.size() && std::iswspace(static_cast<wint_t>(text[start])) != 0)
    {
        ++start;
    }

    size_t end = text.size();
    while (end > start && std::iswspace(static_cast<wint_t>(text[end - 1])) != 0)
    {
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] std::wstring GetDlgItemTextString(HWND dlg, int controlId) noexcept
{
    HWND control = dlg ? GetDlgItem(dlg, controlId) : nullptr;
    if (! control)
    {
        return {};
    }

    const int len = GetWindowTextLengthW(control);
    if (len <= 0)
    {
        return {};
    }

    std::wstring buffer(static_cast<size_t>(len) + 1u, L'\0');
    const int written = GetWindowTextW(control, buffer.data(), len + 1);
    if (written <= 0)
    {
        return {};
    }
    if (written < static_cast<int>(buffer.size()))
    {
        buffer.resize(static_cast<size_t>(written));
    }
    else if (! buffer.empty())
    {
        buffer.resize(buffer.size() - 1u);
    }
    return buffer;
}

void PrepareFlatControl(HWND control) noexcept
{
    if (! control)
    {
        return;
    }

    const LONG_PTR exStyle = GetWindowLongPtrW(control, GWL_EXSTYLE);
    if ((exStyle & WS_EX_CLIENTEDGE) == 0)
    {
        return;
    }

    SetWindowLongPtrW(control, GWL_EXSTYLE, exStyle & ~WS_EX_CLIENTEDGE);
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

void ClearValidation(HWND dlg, DialogState* state) noexcept
{
    if (! dlg || ! state)
    {
        return;
    }
    state->showingValidation = false;
    SetDlgItemTextW(dlg, IDC_CONNECTION_CRED_PROMPT_VALIDATION, L"");
    InvalidateRect(GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_VALIDATION), nullptr, TRUE);
}

void ShowValidation(HWND dlg, DialogState* state, std::wstring_view text) noexcept
{
    if (! dlg || ! state)
    {
        return;
    }

    state->showingValidation = true;
    SetDlgItemTextW(dlg, IDC_CONNECTION_CRED_PROMPT_VALIDATION, std::wstring(text).c_str());
    InvalidateRect(GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_VALIDATION), nullptr, TRUE);
    MessageBeep(MB_ICONWARNING);
}

INT_PTR OnCtlColorDialog(DialogState* state) noexcept
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCtlColorStatic(DialogState* state, HDC hdc, HWND control) noexcept
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    COLORREF textColor = state->theme.menu.text;
    if (control && state->showingValidation)
    {
        const int controlId = GetDlgCtrlID(control);
        if (controlId == IDC_CONNECTION_CRED_PROMPT_VALIDATION)
        {
            textColor = ColorRefFromColorF(state->theme.folderView.errorText);
        }
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCtlColorEdit(DialogState* state, HDC hdc, HWND control) noexcept
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    const bool enabled = ! control || IsWindowEnabled(control) != FALSE;
    const bool focused = enabled && control && GetFocus() == control;
    const COLORREF bg  = enabled ? (focused ? state->inputFocusedBackgroundColor : state->inputBackgroundColor) : state->inputDisabledBackgroundColor;

    SetBkColor(hdc, bg);
    SetTextColor(hdc, enabled ? state->theme.menu.text : state->theme.menu.disabledText);

    if (state->theme.highContrast)
    {
        return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
    }

    if (! enabled)
    {
        return reinterpret_cast<INT_PTR>(state->inputDisabledBrush.get());
    }
    return reinterpret_cast<INT_PTR>(focused && state->inputFocusedBrush ? state->inputFocusedBrush.get() : state->inputBrush.get());
}

void UpdateSecretVisibility(HWND dlg, DialogState* state) noexcept
{
    if (! dlg || ! state)
    {
        return;
    }

    const HWND secretEdit = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_EDIT);
    if (! secretEdit)
    {
        return;
    }

    DWORD selStart = 0;
    DWORD selEnd   = 0;
    SendMessageW(secretEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selStart), reinterpret_cast<LPARAM>(&selEnd));

    LONG_PTR style = GetWindowLongPtrW(secretEdit, GWL_STYLE);
    if (state->secretVisible)
    {
        style &= ~static_cast<LONG_PTR>(ES_PASSWORD);
        SetWindowLongPtrW(secretEdit, GWL_STYLE, style);
        SendMessageW(secretEdit, EM_SETPASSWORDCHAR, 0, 0);
    }
    else
    {
        style |= ES_PASSWORD;
        SetWindowLongPtrW(secretEdit, GWL_STYLE, style);
        SendMessageW(secretEdit, EM_SETPASSWORDCHAR, static_cast<WPARAM>(L'\u2022'), 0);
    }

    SetWindowPos(secretEdit, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    SendMessageW(secretEdit, EM_SETSEL, selStart, selEnd);
    InvalidateRect(secretEdit, nullptr, TRUE);

    if (const HWND showBtn = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_SHOW_SECRET))
    {
        const UINT labelId       = state->secretVisible ? IDS_CONNECTIONS_BTN_HIDE_SECRET : IDS_CONNECTIONS_BTN_SHOW_SECRET;
        const std::wstring label = LoadStringResource(nullptr, labelId);
        if (! label.empty())
        {
            SetWindowTextW(showBtn, label.c_str());
        }
    }
}

void MoveControlY(HWND dlg, int controlId, int deltaY) noexcept
{
    const HWND control = dlg ? GetDlgItem(dlg, controlId) : nullptr;
    if (! control || deltaY == 0)
    {
        return;
    }

    RECT rect{};
    if (GetWindowRect(control, &rect) == 0)
    {
        return;
    }

    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&rect), 2);
    SetWindowPos(control, nullptr, rect.left, rect.top + deltaY, rect.right - rect.left, rect.bottom - rect.top, SWP_NOZORDER | SWP_NOACTIVATE);
}

void CompactLayoutIfNoUser(HWND dlg, DialogState* state) noexcept
{
    if (! dlg || ! state || state->showUserName)
    {
        return;
    }

    const HWND userLabel   = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_USER_LABEL);
    const HWND secretLabel = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_LABEL);
    if (! userLabel || ! secretLabel)
    {
        return;
    }

    RECT userRect{};
    RECT secretRect{};
    if (GetWindowRect(userLabel, &userRect) == 0 || GetWindowRect(secretLabel, &secretRect) == 0)
    {
        return;
    }

    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&userRect), 2);
    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&secretRect), 2);

    const int deltaY = userRect.top - secretRect.top;
    if (deltaY == 0)
    {
        return;
    }

    ShowWindow(userLabel, SW_HIDE);
    if (const HWND userEdit = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_USER_EDIT))
    {
        ShowWindow(userEdit, SW_HIDE);
    }

    MoveControlY(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_LABEL, deltaY);
    MoveControlY(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_EDIT, deltaY);
    MoveControlY(dlg, IDC_CONNECTION_CRED_PROMPT_SHOW_SECRET, deltaY);
    MoveControlY(dlg, IDC_CONNECTION_CRED_PROMPT_VALIDATION, deltaY);
    MoveControlY(dlg, IDOK, deltaY);
    MoveControlY(dlg, IDCANCEL, deltaY);

    RECT windowRect{};
    if (GetWindowRect(dlg, &windowRect) == 0)
    {
        return;
    }

    const int height    = std::max(0l, windowRect.bottom - windowRect.top);
    const int newHeight = std::max(0, height + deltaY);

    SetWindowPos(dlg, nullptr, 0, 0, std::max(0l, windowRect.right - windowRect.left), newHeight, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
}

LRESULT CALLBACK SecretEditSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept
{
    UNREFERENCED_PARAMETER(subclassId);

    auto* state = reinterpret_cast<DialogState*>(refData);
    if (! state)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    if (msg == WM_KEYDOWN && wp == VK_ESCAPE)
    {
        HWND dlg = GetParent(hwnd);
        if (dlg)
        {
            SendMessageW(dlg, WM_COMMAND, MAKEWPARAM(IDCANCEL, BN_CLICKED), 0);
            return 0;
        }
    }

    if (msg == WM_NCDESTROY)
    {
#pragma warning(push)
#pragma warning(disable : 5039) // passing potentially-throwing callback to extern "C" Win32 API under -EHc
        RemoveWindowSubclass(hwnd, SecretEditSubclassProc, kSecretEditSubclassId);
#pragma warning(pop)
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

INT_PTR OnInitDialog(HWND dlg, DialogState* state) noexcept
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

    ApplyTitleBarTheme(dlg, state->theme, GetActiveWindow() == dlg);
    state->backgroundBrush.reset(CreateSolidBrush(state->theme.windowBackground));

    const COLORREF surface             = ThemedControls::GetControlSurfaceColor(state->theme);
    state->inputBackgroundColor        = ThemedControls::BlendColor(surface, state->theme.windowBackground, state->theme.dark ? 50 : 30, 255);
    state->inputFocusedBackgroundColor = ThemedControls::BlendColor(state->inputBackgroundColor, state->theme.menu.text, state->theme.dark ? 20 : 16, 255);
    state->inputDisabledBackgroundColor =
        ThemedControls::BlendColor(state->theme.windowBackground, state->inputBackgroundColor, state->theme.dark ? 70 : 40, 255);

    state->inputBrush.reset();
    state->inputFocusedBrush.reset();
    state->inputDisabledBrush.reset();
    if (! state->theme.highContrast)
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

    if (! state->caption.empty())
    {
        SetWindowTextW(dlg, state->caption.c_str());
    }

    SetDlgItemTextW(dlg, IDC_CONNECTION_CRED_PROMPT_MESSAGE, state->message.c_str());
    SetDlgItemTextW(dlg, IDC_CONNECTION_CRED_PROMPT_USER_LABEL, LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_USER).c_str());
    SetDlgItemTextW(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_LABEL, state->secretLabel.c_str());

    SetDlgItemTextW(dlg, IDOK, LoadStringResource(nullptr, IDS_BTN_OK).c_str());
    SetDlgItemTextW(dlg, IDCANCEL, LoadStringResource(nullptr, IDS_BTN_CANCEL).c_str());

    if (! state->theme.highContrast)
    {
        ThemedControls::EnableOwnerDrawButton(dlg, IDOK);
        ThemedControls::EnableOwnerDrawButton(dlg, IDCANCEL);
        ThemedControls::EnableOwnerDrawButton(dlg, IDC_CONNECTION_CRED_PROMPT_SHOW_SECRET);
    }

    if (const HWND userEdit = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_USER_EDIT))
    {
        SetWindowTextW(userEdit, state->initialUserName.c_str());
        PrepareFlatControl(userEdit);
        PrepareEditMargins(userEdit);
        ThemedControls::CenterEditTextVertically(userEdit);
    }

    if (const HWND secretEdit = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_EDIT))
    {
        PrepareFlatControl(secretEdit);
        PrepareEditMargins(secretEdit);
        ThemedControls::CenterEditTextVertically(secretEdit);
#pragma warning(push)
#pragma warning(disable : 5039) // passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(secretEdit, SecretEditSubclassProc, kSecretEditSubclassId, reinterpret_cast<DWORD_PTR>(state));
#pragma warning(pop)
    }

    state->secretVisible = false;
    UpdateSecretVisibility(dlg, state);
    ClearValidation(dlg, state);

    CompactLayoutIfNoUser(dlg, state);

    if (! state->theme.highContrast)
    {
        auto createFrame = [&](wil::unique_hwnd& frameOut, HWND input) noexcept
        {
            if (! input)
            {
                return;
            }

            frameOut.reset(
                CreateWindowExW(0, L"Static", L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, 0, 0, 10, 10, dlg, nullptr, GetModuleHandleW(nullptr), nullptr));
            if (! frameOut)
            {
                return;
            }

            SetWindowPos(frameOut.get(), input, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            ThemedInputFrames::InstallFrame(frameOut.get(), input, &state->inputFrameStyle);
        };

        if (state->showUserName)
        {
            createFrame(state->userFrame, GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_USER_EDIT));
        }
        createFrame(state->secretFrame, GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_EDIT));
    }

    CenterWindowOnOwner(dlg, GetParent(dlg));

    if (state->showUserName)
    {
        if (const HWND userEdit = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_USER_EDIT))
        {
            SendMessageW(userEdit, EM_SETSEL, 0, -1);
            SetFocus(userEdit);
            return FALSE;
        }
    }

    if (const HWND secretEdit = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_EDIT))
    {
        SetFocus(secretEdit);
        return FALSE;
    }

    return TRUE;
}

INT_PTR OnCommand(HWND dlg, DialogState* state, int controlId, int notifyCode) noexcept
{
    if (! dlg)
    {
        return FALSE;
    }

    if (controlId == IDC_CONNECTION_CRED_PROMPT_SHOW_SECRET && notifyCode == BN_CLICKED && state)
    {
        state->secretVisible = ! state->secretVisible;
        UpdateSecretVisibility(dlg, state);
        return TRUE;
    }

    if (state && (notifyCode == EN_SETFOCUS || notifyCode == EN_KILLFOCUS) &&
        (controlId == IDC_CONNECTION_CRED_PROMPT_USER_EDIT || controlId == IDC_CONNECTION_CRED_PROMPT_SECRET_EDIT))
    {
        if (const HWND edit = GetDlgItem(dlg, controlId))
        {
            InvalidateRect(edit, nullptr, TRUE);
        }

        if (! state->theme.highContrast)
        {
            if (controlId == IDC_CONNECTION_CRED_PROMPT_USER_EDIT && state->userFrame)
            {
                InvalidateRect(state->userFrame.get(), nullptr, TRUE);
            }
            if (controlId == IDC_CONNECTION_CRED_PROMPT_SECRET_EDIT && state->secretFrame)
            {
                InvalidateRect(state->secretFrame.get(), nullptr, TRUE);
            }
        }

        return FALSE;
    }

    if (controlId == IDCANCEL)
    {
        EndDialog(dlg, IDCANCEL);
        return TRUE;
    }

    if (controlId != IDOK)
    {
        return FALSE;
    }

    if (! state)
    {
        return FALSE;
    }

    ClearValidation(dlg, state);

    std::wstring userName;
    if (state->showUserName)
    {
        userName = TrimWhitespace(GetDlgItemTextString(dlg, IDC_CONNECTION_CRED_PROMPT_USER_EDIT));
        if (userName.empty())
        {
            ShowValidation(dlg, state, LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_PROMPT_USER_REQUIRED));
            if (const HWND userEdit = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_USER_EDIT))
            {
                SetFocus(userEdit);
            }
            return TRUE;
        }
    }

    std::wstring secret = GetDlgItemTextString(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_EDIT);
    if (! state->allowEmptySecret)
    {
        if (secret.empty())
        {
            ShowValidation(dlg, state, LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_PROMPT_PASSWORD_REQUIRED));
            if (const HWND secretEdit = GetDlgItem(dlg, IDC_CONNECTION_CRED_PROMPT_SECRET_EDIT))
            {
                SetFocus(secretEdit);
            }
            return TRUE;
        }
    }

    state->userNameOut = std::move(userName);
    state->secretOut   = std::move(secret);
    EndDialog(dlg, IDOK);
    return TRUE;
}

INT_PTR CALLBACK DialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<DialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnInitDialog(dlg, reinterpret_cast<DialogState*>(lp));
        case WM_ERASEBKGND:
            if (state && state->backgroundBrush && wp)
            {
                RECT rc{};
                if (GetClientRect(dlg, &rc))
                {
                    FillRect(reinterpret_cast<HDC>(wp), &rc, state->backgroundBrush.get());
                    return TRUE;
                }
            }
            break;
        case WM_CTLCOLORDLG: return OnCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLOREDIT: return OnCtlColorEdit(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->theme, wp != FALSE);
            }
            return FALSE;
        case WM_DRAWITEM:
        {
            if (! state || state->theme.highContrast)
            {
                break;
            }

            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis || dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            ThemedControls::DrawThemedPushButton(*dis, state->theme);
            return TRUE;
        }
        case WM_COMMAND: return OnCommand(dlg, state, LOWORD(wp), HIWORD(wp));
    }

    return FALSE;
}

[[nodiscard]] HRESULT ShowPromptDialog(HWND ownerWindow, DialogState& state) noexcept
{
#pragma warning(push)
    // pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
    const INT_PTR result =
        DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_CONNECTION_CRED_PROMPT), ownerWindow, DialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)

    if (result == IDCANCEL)
    {
        state.userNameOut.clear();
        state.secretOut.clear();
        return S_FALSE;
    }

    if (result != IDOK)
    {
        state.userNameOut.clear();
        state.secretOut.clear();
        return E_FAIL;
    }

    return S_OK;
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
    secretOut.clear();

    DialogState state{};
    state.theme            = theme;
    state.showUserName     = false;
    state.allowEmptySecret = allowEmptySecret;
    state.caption          = std::wstring(caption);
    state.message          = std::wstring(message);
    state.secretLabel      = std::wstring(secretLabel);

    const HRESULT hr = ShowPromptDialog(ownerWindow, state);
    if (hr != S_OK)
    {
        return hr;
    }

    secretOut = std::move(state.secretOut);
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

    DialogState state{};
    state.theme            = theme;
    state.showUserName     = true;
    state.allowEmptySecret = false;
    state.caption          = std::wstring(caption);
    state.message          = std::wstring(message);
    state.secretLabel      = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PASSWORD);
    state.initialUserName  = std::wstring(initialUserName);

    const HRESULT hr = ShowPromptDialog(ownerWindow, state);
    if (hr != S_OK)
    {
        return hr;
    }

    userNameOut = std::move(state.userNameOut);
    passwordOut = std::move(state.secretOut);
    return S_OK;
}
