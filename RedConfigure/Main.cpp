#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "DxUi.h"
#include "MinimumOsVersion.h"
#include "RedConfigureApp.h"
#include "RedConfigureSession.h"
#include "RedConfigureRoot.h"
#include "SettingsStore.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <iterator>
#include <limits>
#include <memory>
#include <optional>
#include <set>
#include <span>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026, C5027, C4820
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace
{
constexpr wchar_t kWindowClassName[] = L"RedConfigure.MainWindow";

[[nodiscard]] std::wstring LoadAppString(HINSTANCE instance, UINT resourceId)
{
    wchar_t buffer[1024]{};
    const int length = ::LoadStringW(instance, resourceId, buffer, static_cast<int>(std::size(buffer)));
    if (length <= 0)
    {
        return {};
    }

    return std::wstring(buffer, static_cast<size_t>(length));
}

class MainWindow final
{
public:
    explicit MainWindow(HINSTANCE instance) noexcept : _instance(instance) {}

    MainWindow(const MainWindow&)            = delete;
    MainWindow& operator=(const MainWindow&) = delete;
    MainWindow(MainWindow&&)                 = delete;
    MainWindow& operator=(MainWindow&&)      = delete;

    [[nodiscard]] HWND Create(int showCommand) noexcept
    {
        HICON icon = ::LoadIconW(_instance, MAKEINTRESOURCEW(IDI_REDCONFIGURE_APP));
        if (! icon)
        {
            icon = ::LoadIconW(nullptr, IDI_APPLICATION);
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
        wc.lpfnWndProc   = &MainWindow::WndProc;
        wc.hInstance     = _instance;
        wc.hCursor       = ::LoadCursorW(nullptr, IDC_ARROW);
        wc.hIcon         = icon;
        wc.hIconSm       = icon;
        wc.lpszClassName = kWindowClassName;

        if (::RegisterClassExW(&wc) == 0 && ::GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
        {
            return nullptr;
        }

        wil::unique_hmenu menu(::LoadMenuW(_instance, MAKEINTRESOURCEW(IDR_REDCONFIGURE_MAINMENU)));
        const std::wstring title = LoadAppString(_instance, IDS_REDCONFIGURE_APP_TITLE);

        HWND hwnd = ::CreateWindowExW(0,
                                      kWindowClassName,
                                      title.c_str(),
                                      WS_OVERLAPPEDWINDOW,
                                      CW_USEDEFAULT,
                                      CW_USEDEFAULT,
                                      1280,
                                      860,
                                      nullptr,
                                      menu.get(),
                                      _instance,
                                      this);
        if (! hwnd)
        {
            return nullptr;
        }

        menu.release();
        ::ShowWindow(hwnd, showCommand);
        ::UpdateWindow(hwnd);
        return hwnd;
    }

private:
    [[nodiscard]] static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        MainWindow* window = nullptr;
        if (message == WM_NCCREATE)
        {
            const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            window             = static_cast<MainWindow*>(create ? create->lpCreateParams : nullptr);
            ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(window));
            if (window)
            {
                window->_hwnd = hwnd;
            }
        }
        else
        {
            window = reinterpret_cast<MainWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        }

        if (! window)
        {
            return ::DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool dxHandled         = false;
        const LRESULT dxResult = window->_dxHost.HandleMessage(hwnd, message, wParam, lParam, dxHandled);
        if (dxHandled)
        {
            if (message == WM_SIZE)
            {
                window->Layout();
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return window->OnCreate() ? 0 : -1;
            case WM_SIZE:
                window->Layout();
                return 0;
            case WM_GETMINMAXINFO:
                if (auto* info = reinterpret_cast<MINMAXINFO*>(lParam))
                {
                    info->ptMinTrackSize.x = 860;
                    info->ptMinTrackSize.y = 620;
                }
                return 0;
            case WM_COMMAND:
                if (LOWORD(wParam) == IDM_REDCONFIGURE_EXIT)
                {
                    ::DestroyWindow(hwnd);
                    return 0;
                }
                break;
            case WM_DESTROY:
                ::PostQuitMessage(0);
                return 0;
            case WM_NCDESTROY:
                window->_dxHost.Detach();
                window->_root = nullptr;
                window->_rootController = nullptr;
                ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                window->_hwnd = nullptr;
                return ::DefWindowProcW(hwnd, message, wParam, lParam);
            default: break;
        }

        return ::DefWindowProcW(hwnd, message, wParam, lParam);
    }

    [[nodiscard]] bool OnCreate()
    {
        if (! _dxHost.Attach(_hwnd))
        {
            return false;
        }

        _dxHost.SetTheme(RedSalamander::DxUi::MakeDefaultThemePalette(false));

        std::filesystem::path root;
        std::array<wchar_t, 32768> modulePath{};
        const DWORD moduleLength = ::GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
        if (moduleLength > 0u && moduleLength < modulePath.size())
        {
            root = RedConfigure::ResolveWorkspaceRootForLaunchPath(std::filesystem::path(modulePath.data()).parent_path());
        }
        else
        {
            std::error_code ec;
            root = std::filesystem::current_path(ec);
            if (ec)
            {
                root = std::filesystem::path{};
            }
            root = RedConfigure::ResolveWorkspaceRootForLaunchPath(root);
        }

        auto ui = RedConfigure::Ui::CreateRedConfigureRoot(_instance, _session, root);
        _root = ui.control.get();
        _rootController = ui.controller;
        _dxHost.SetRoot(std::move(ui.control));
        if (_rootController)
        {
            _rootController->ReloadWorkspaceFromFields();
        }
        Layout();
        return _dxHost.PrimeForShow();
    }

    void Layout() noexcept
    {
        if (_root)
        {
            _root->SetBounds(_dxHost.GetClientBoundsDip());
            _dxHost.Invalidate();
        }
    }

    HINSTANCE _instance = nullptr;
    HWND _hwnd          = nullptr;
    RedSalamander::DxUi::WindowHost _dxHost;
    RedConfigure::RedConfigureSession _session;
    RedSalamander::DxUi::Panel* _root = nullptr;
    RedConfigure::Ui::RedConfigureRootController* _rootController = nullptr;
};
} // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int showCommand)
{
    if (! Common::MinimumOsVersion::EnsureCurrentWindowsVersionSupported(nullptr))
    {
        return 1;
    }

    MainWindow window(instance);
    wil::unique_hwnd hwnd(window.Create(showCommand));
    if (! hwnd)
    {
        const std::wstring title   = LoadAppString(instance, IDS_REDCONFIGURE_APP_TITLE);
        const std::wstring message = LoadAppString(instance, IDS_REDCONFIGURE_ERROR_CREATE_WINDOW);
        if (! message.empty())
        {
            ::MessageBoxW(nullptr, message.c_str(), title.c_str(), MB_OK | MB_ICONERROR);
        }
        return 1;
    }

    MSG msg{};
    while (::GetMessageW(&msg, nullptr, 0, 0) > 0)
    {
        ::TranslateMessage(&msg);
        ::DispatchMessageW(&msg);
    }

    if (hwnd && ::IsWindow(hwnd.get()) == FALSE)
    {
        hwnd.release();
    }

    return static_cast<int>(msg.wParam);
}
