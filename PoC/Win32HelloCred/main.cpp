#include <strsafe.h>
#include <wincred.h>
#include <windows.h>

#include "../../Common/WindowMessages.h"

#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Security.Credentials.UI.h>
#include <winrt/base.h>

#include <userconsentverifierinterop.h> // IUserConsentVerifierInterop

#pragma comment(lib, "Advapi32.lib")
#pragma comment(lib, "RuntimeObject.lib")
#pragma comment(lib, "WindowsApp.lib")

using namespace winrt;
using namespace winrt::Windows::Foundation;
using namespace winrt::Windows::Security::Credentials::UI;

// -------------------------
// WinCred helpers
// -------------------------
static bool SaveGenericCredential(const std::wstring& targetName, const std::wstring& userName, const std::wstring& secret)
{
    if (targetName.empty() || secret.empty())
        return false;

    CREDENTIALW cred{};
    cred.Type       = CRED_TYPE_GENERIC;
    cred.TargetName = const_cast<LPWSTR>(targetName.c_str());
    cred.UserName   = const_cast<LPWSTR>(userName.c_str());
    cred.Persist    = CRED_PERSIST_LOCAL_MACHINE;

    cred.CredentialBlobSize = static_cast<DWORD>((secret.size() + 1) * sizeof(wchar_t));
    cred.CredentialBlob     = (LPBYTE)secret.c_str();

    return CredWriteW(&cred, 0) == TRUE;
}

static bool LoadGenericCredential(const std::wstring& targetName, std::wstring& userNameOut, std::wstring& secretOut)
{
    PCREDENTIALW pcred = nullptr;
    if (! CredReadW(targetName.c_str(), CRED_TYPE_GENERIC, 0, &pcred))
        return false;

    userNameOut = pcred->UserName ? pcred->UserName : L"";

    // We stored UTF-16 chars (wchar_t) in CredentialBlob.
    const wchar_t* blob = reinterpret_cast<const wchar_t*>(pcred->CredentialBlob);
    secretOut           = blob ? blob : L"";

    CredFree(pcred);
    return true;
}

static bool DeleteGenericCredential(const std::wstring& targetName)
{
    if (targetName.empty())
        return false;
    return CredDeleteW(targetName.c_str(), CRED_TYPE_GENERIC, 0) == TRUE;
}

// -------------------------
// UI helpers
// -------------------------
static std::wstring GetWindowTextString(HWND h)
{
    int len = GetWindowTextLengthW(h);
    std::wstring s;
    s.resize(static_cast<size_t>(len));
    if (len > 0)
        GetWindowTextW(h, s.data(), len + 1);
    return s;
}

static void ShowLastError(HWND hwnd, const wchar_t* prefix)
{
    DWORD err = GetLastError();
    wchar_t buf[512]{};
    FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, err, 0, buf, (DWORD)std::size(buf), nullptr);

    wchar_t msg[1024]{};
    StringCchPrintfW(msg, std::size(msg), L"%s\n\nGetLastError=%lu\n%s", prefix, err, buf);
    MessageBoxW(hwnd, msg, L"Win32HelloCred", MB_ICONERROR);
}

static void ShowHresult(HWND hwnd, const wchar_t* prefix, HRESULT hr)
{
    wchar_t msg[512]{};
    StringCchPrintfW(msg, std::size(msg), L"%s\n\nHRESULT=0x%08X", prefix, (unsigned)hr);
    MessageBoxW(hwnd, msg, L"Win32HelloCred", MB_ICONERROR);
}

// -------------------------
// Controls / IDs
// -------------------------
enum : int
{
    IDC_EDIT_TARGET = 1001,
    IDC_EDIT_USER   = 1002,
    IDC_EDIT_SECRET = 1003,
    IDC_BTN_SAVE    = 1101,
    IDC_BTN_LOAD    = 1102,
    IDC_BTN_DELETE  = 1103,
};

static HWND g_hEditTarget = nullptr;
static HWND g_hEditUser   = nullptr;
static HWND g_hEditSecret = nullptr;

static std::wstring g_pendingTarget; // read after Hello verification

// -------------------------
// Windows Hello verification (non-blocking)
// -------------------------
static void BeginHelloVerification(HWND hwnd)
{
    // Step 1: Check availability
    auto availOp = UserConsentVerifier::CheckAvailabilityAsync();
    availOp.Completed(
        [hwnd](auto const& op, winrt::Windows::Foundation::AsyncStatus status)
        {
            if (status != winrt::Windows::Foundation::AsyncStatus::Completed)
            {
                PostMessageW(hwnd, WndMsg::kWin32HelloCredHelloResult, (WPARAM)UserConsentVerificationResult::Canceled, 0);
                return;
            }

            UserConsentVerifierAvailability avail{};
            try
            {
                avail = op.GetResults();
            }
            catch (...)
            {
                PostMessageW(hwnd, WndMsg::kWin32HelloCredHelloResult, (WPARAM)UserConsentVerificationResult::DeviceNotPresent, 0);
                return;
            }

            if (avail != UserConsentVerifierAvailability::Available)
            {
                // Map "not available" to a representative result.
                PostMessageW(hwnd, WndMsg::kWin32HelloCredHelloResult, (WPARAM)UserConsentVerificationResult::DeviceNotPresent, 0);
                return;
            }

            // Step 2: Request verification attached to our window (HWND)
            try
            {
                winrt::hstring message = L"Verify with Windows Hello to access the stored credential";

                winrt::com_ptr<IUserConsentVerifierInterop> interop = winrt::get_activation_factory<UserConsentVerifier, IUserConsentVerifierInterop>();

                IAsyncOperation<UserConsentVerificationResult> verifyOp{nullptr};
                HRESULT hr = interop->RequestVerificationForWindowAsync(
                    hwnd, (HSTRING)winrt::get_abi(message), winrt::guid_of<IAsyncOperation<UserConsentVerificationResult>>(), winrt::put_abi(verifyOp));

                if (FAILED(hr))
                {
                    PostMessageW(hwnd, WndMsg::kWin32HelloCredHelloResult, (WPARAM)UserConsentVerificationResult::Canceled, (LPARAM)hr);
                    return;
                }

                verifyOp.Completed(
                    [hwnd](auto const& vop, winrt::Windows::Foundation::AsyncStatus vstatus)
                    {
                        UserConsentVerificationResult res = UserConsentVerificationResult::Canceled;
                        if (vstatus == winrt::Windows::Foundation::AsyncStatus::Completed)
                        {
                            try
                            {
                                res = vop.GetResults();
                            }
                            catch (...)
                            {
                                res = UserConsentVerificationResult::Canceled;
                            }
                        }
                        PostMessageW(hwnd, WndMsg::kWin32HelloCredHelloResult, (WPARAM)res, 0);
                    });
            }
            catch (...)
            {
                PostMessageW(hwnd, WndMsg::kWin32HelloCredHelloResult, (WPARAM)UserConsentVerificationResult::Canceled, 0);
            }
        });
}

// -------------------------
// Win32 window procedure
// -------------------------
static void LayoutControls(HWND hwnd)
{
    RECT rc{};
    GetClientRect(hwnd, &rc);
    int x = 14, y = 14;
    int labelW = 90;
    int editW  = (rc.right - rc.left) - x - labelW - 24;
    int rowH   = 26;
    int gapY   = 10;

    auto move = [](HWND h, int X, int Y, int W, int H) { MoveWindow(h, X, Y, W, H, TRUE); };

    // We created the static labels just before edits in WM_CREATE.
    // Find them by walking siblings is messy; easiest is keep handles.
    // For this sample we layout only the edits + buttons; labels are positioned relative.

    // Target label is HWND of previous sibling; ignore.
    move(g_hEditTarget, x + labelW, y, editW, rowH);
    y += rowH + gapY;
    move(g_hEditUser, x + labelW, y, editW, rowH);
    y += rowH + gapY;
    move(g_hEditSecret, x + labelW, y, editW, rowH);
    y += rowH + gapY + 6;

    int btnW = 160, btnH = 32;
    int btnX = x + labelW;
    int btnY = y;

    move(GetDlgItem(hwnd, IDC_BTN_SAVE), btnX, btnY, btnW, btnH);
    move(GetDlgItem(hwnd, IDC_BTN_LOAD), btnX + btnW + 10, btnY, btnW, btnH);
    move(GetDlgItem(hwnd, IDC_BTN_DELETE), btnX, btnY + btnH + 10, btnW, btnH);
}

static LRESULT OnCreateMainWindow(HWND hwnd)
{
    CreateWindowW(L"STATIC", L"Target:", WS_CHILD | WS_VISIBLE, 14, 14, 80, 20, hwnd, nullptr, nullptr, nullptr);
    g_hEditTarget = CreateWindowExW(
        WS_EX_CLIENTEDGE, L"EDIT", L"MyApp:demo", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 110, 14, 300, 24, hwnd, (HMENU)IDC_EDIT_TARGET, nullptr, nullptr);

    CreateWindowW(L"STATIC", L"Username:", WS_CHILD | WS_VISIBLE, 14, 50, 80, 20, hwnd, nullptr, nullptr, nullptr);
    g_hEditUser = CreateWindowExW(
        WS_EX_CLIENTEDGE, L"EDIT", L"user@example.com", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 110, 50, 300, 24, hwnd, (HMENU)IDC_EDIT_USER, nullptr, nullptr);

    CreateWindowW(L"STATIC", L"Secret:", WS_CHILD | WS_VISIBLE, 14, 86, 80, 20, hwnd, nullptr, nullptr, nullptr);
    g_hEditSecret = CreateWindowExW(
        WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_PASSWORD, 110, 86, 300, 24, hwnd, (HMENU)IDC_EDIT_SECRET, nullptr, nullptr);

    CreateWindowW(
        L"BUTTON", L"Save to Credential Manager", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 110, 130, 180, 32, hwnd, (HMENU)IDC_BTN_SAVE, nullptr, nullptr);

    CreateWindowW(L"BUTTON", L"Load (requires Hello)", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 300, 130, 180, 32, hwnd, (HMENU)IDC_BTN_LOAD, nullptr, nullptr);

    CreateWindowW(L"BUTTON", L"Delete credential", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 110, 172, 180, 32, hwnd, (HMENU)IDC_BTN_DELETE, nullptr, nullptr);

    LayoutControls(hwnd);
    return 0;
}

static LRESULT OnCommandMainWindow(HWND hwnd, WORD commandId)
{
    switch (commandId)
    {
        case IDC_BTN_SAVE:
        {
            std::wstring target = GetWindowTextString(g_hEditTarget);
            std::wstring user   = GetWindowTextString(g_hEditUser);
            std::wstring secret = GetWindowTextString(g_hEditSecret);

            if (! SaveGenericCredential(target, user, secret))
                ShowLastError(hwnd, L"CredWriteW failed.");
            else
                MessageBoxW(hwnd, L"Saved to Windows Credential Manager.", L"Win32HelloCred", MB_OK | MB_ICONINFORMATION);

            // Best-effort wipe of the secret edit control.
            SetWindowTextW(g_hEditSecret, L"");
            SecureZeroMemory(secret.data(), secret.size() * sizeof(wchar_t));
            return 0;
        }

        case IDC_BTN_LOAD:
        {
            g_pendingTarget = GetWindowTextString(g_hEditTarget);
            if (g_pendingTarget.empty())
            {
                MessageBoxW(hwnd, L"Target is empty.", L"Win32HelloCred", MB_OK | MB_ICONWARNING);
                return 0;
            }

            BeginHelloVerification(hwnd);
            return 0;
        }

        case IDC_BTN_DELETE:
        {
            std::wstring target = GetWindowTextString(g_hEditTarget);
            if (! DeleteGenericCredential(target))
                ShowLastError(hwnd, L"CredDeleteW failed (or credential not found).");
            else
                MessageBoxW(hwnd, L"Deleted credential.", L"Win32HelloCred", MB_OK | MB_ICONINFORMATION);
            return 0;
        }
    }

    return 0;
}

static LRESULT OnHelloResultMainWindow(HWND hwnd, UserConsentVerificationResult res)
{
    if (res != UserConsentVerificationResult::Verified)
    {
        // Basic feedback for common outcomes.
        const wchar_t* why = L"Verification failed or was canceled.";
        switch (res)
        {
            case UserConsentVerificationResult::Canceled: why = L"Canceled."; break;
            case UserConsentVerificationResult::DeviceNotPresent: why = L"Windows Hello is not available on this device/profile."; break;
            case UserConsentVerificationResult::NotConfiguredForUser: why = L"Windows Hello is not configured for this user."; break;
            case UserConsentVerificationResult::DisabledByPolicy: why = L"Disabled by policy."; break;
            case UserConsentVerificationResult::DeviceBusy: why = L"Device busy."; break;
            case UserConsentVerificationResult::RetriesExhausted: why = L"Too many failed attempts."; break;
            default: break;
        }

        MessageBoxW(hwnd, why, L"Win32HelloCred", MB_OK | MB_ICONWARNING);
        return 0;
    }

    // Verified -> read credential and show it (for demo purposes).
    std::wstring user;
    std::wstring secret;
    if (! LoadGenericCredential(g_pendingTarget, user, secret))
    {
        ShowLastError(hwnd, L"CredReadW failed.");
        return 0;
    }

    wchar_t msg[1024]{};
    StringCchPrintfW(msg,
                     std::size(msg),
                     L"Target: %s\nUsername: %s\nSecret: %s\n\n(Showing the secret is only for demo.)",
                     g_pendingTarget.c_str(),
                     user.c_str(),
                     secret.c_str());

    MessageBoxW(hwnd, msg, L"Credential loaded", MB_OK | MB_ICONINFORMATION);

    // Best-effort memory wipe.
    SecureZeroMemory(secret.data(), secret.size() * sizeof(wchar_t));
    SecureZeroMemory(user.data(), user.size() * sizeof(wchar_t));
    return 0;
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
        case WM_CREATE: return OnCreateMainWindow(hwnd);
        case WM_SIZE: LayoutControls(hwnd); return 0;
        case WM_COMMAND: return OnCommandMainWindow(hwnd, LOWORD(wParam));
        case WndMsg::kWin32HelloCredHelloResult: return OnHelloResultMainWindow(hwnd, static_cast<UserConsentVerificationResult>(wParam));
        case WM_DESTROY: PostQuitMessage(0); return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// -------------------------
// Entry point
// -------------------------
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow)
{
    // Initialize WinRT for this thread (needed for UserConsentVerifier).
    winrt::init_apartment(winrt::apartment_type::single_threaded);

    const wchar_t CLASS_NAME[] = L"Win32HelloCredWindow";

    WNDCLASSW wc{};
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);

    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(0,
                                CLASS_NAME,
                                L"Win32HelloCred (Credential Manager + Windows Hello)",
                                WS_OVERLAPPEDWINDOW,
                                CW_USEDEFAULT,
                                CW_USEDEFAULT,
                                720,
                                320,
                                nullptr,
                                nullptr,
                                hInstance,
                                nullptr);

    if (! hwnd)
        return 0;

    ShowWindow(hwnd, nCmdShow);

    MSG m{};
    while (GetMessageW(&m, nullptr, 0, 0))
    {
        TranslateMessage(&m);
        DispatchMessageW(&m);
    }
    return 0;
}
