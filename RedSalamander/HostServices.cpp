#include "HostServices.h"

#include <algorithm>
#include <atomic>
#include <cstring>
#include <cwctype>
#include <filesystem>
#include <format>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>

#include <objbase.h>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include "ConnectionCredentialPromptDialog.h"
#include "ConnectionManagerWindow.h"
#include "ConnectionProfileUtils.h"
#include "ConnectionSecrets.h"
#include "FolderWindow.h"
#include "Helpers.h"
#include "SettingsHotReload.h"
#include "SettingsStore.h"
#include "Ui/AlertOverlayWindow.h"
#include "WindowMessages.h"
#include "WindowsHello.h"
#include "resource.h"

// Globals owned by RedSalamander.cpp.
namespace
{
std::atomic<FolderWindow*> g_hostFolderWindow{nullptr};
std::atomic<std::atomic<HWND>*> g_hostFolderWindowHwnd{nullptr};
std::atomic<Common::Settings::Settings*> g_hostSettings{nullptr};

template <typename T> [[nodiscard]] T& RequireHostDependency(const std::atomic<T*>& dependency) noexcept
{
    T* value = dependency.load(std::memory_order_acquire);
    if (! value)
    {
        std::terminate();
    }
    return *value;
}

[[nodiscard]] FolderWindow& HostFolderWindow() noexcept
{
    return RequireHostDependency(g_hostFolderWindow);
}

[[nodiscard]] std::atomic<HWND>& HostFolderWindowHwnd() noexcept
{
    return RequireHostDependency(g_hostFolderWindowHwnd);
}

[[nodiscard]] Common::Settings::Settings& HostSettings() noexcept
{
    return RequireHostDependency(g_hostSettings);
}
} // namespace

namespace
{
struct PendingAlert
{
    HostAlertRequest request{};
    void* cookie = nullptr;
    std::wstring title;
    std::wstring message;
};

struct PendingClearAlert
{
    HostAlertScope scope = HOST_ALERT_SCOPE_APPLICATION;
    void* cookie         = nullptr;
};

struct PendingPrompt
{
    HostPromptRequest request{};
    void* cookie             = nullptr;
    HostPromptResult* result = nullptr;
    std::wstring title;
    std::wstring message;
};

struct PendingConnectionManager
{
    HostConnectionManagerRequest request{};
    HostConnectionManagerResult* result = nullptr;
    std::wstring filterPluginId;
};

struct PendingConnectionSecret
{
    PendingConnectionSecret()                                          = default;
    PendingConnectionSecret(const PendingConnectionSecret&)            = delete;
    PendingConnectionSecret& operator=(const PendingConnectionSecret&) = delete;
    PendingConnectionSecret(PendingConnectionSecret&&)                 = delete;
    PendingConnectionSecret& operator=(PendingConnectionSecret&&)      = delete;

    std::wstring connectionName;
    HostConnectionSecretKind kind = HOST_CONNECTION_SECRET_PASSWORD;
    HWND ownerWindow              = nullptr;
    wil::unique_cotaskmem_string secret;
};

struct PendingSetConnectionSecret
{
    std::wstring connectionName;
    HostConnectionSecretKind kind = HOST_CONNECTION_SECRET_PASSWORD;
    std::wstring secret;
    BOOL persist = FALSE;
};

struct PendingDeleteConnectionSecret
{
    std::wstring connectionName;
    HostConnectionSecretKind kind = HOST_CONNECTION_SECRET_PASSWORD;
    BOOL deletePersisted          = FALSE;
};

struct PendingConnectionJson
{
    PendingConnectionJson()                                        = default;
    PendingConnectionJson(const PendingConnectionJson&)            = delete;
    PendingConnectionJson& operator=(const PendingConnectionJson&) = delete;
    PendingConnectionJson(PendingConnectionJson&&)                 = delete;
    PendingConnectionJson& operator=(PendingConnectionJson&&)      = delete;

    std::wstring connectionName;
    wil::unique_cotaskmem_ptr<char> json;
};

struct PendingClearConnectionSecretCache
{
    std::wstring connectionName;
    HostConnectionSecretKind kind = HOST_CONNECTION_SECRET_PASSWORD;
};

struct PendingUpgradeFtpAnonymousToPassword
{
    std::wstring connectionName;
    HWND ownerWindow = nullptr;
};

struct PendingExecuteInPane
{
    HostPaneExecuteFlags flags = HOST_PANE_EXECUTE_FLAG_NONE;
    std::wstring folderPath;
    std::wstring focusItemDisplayName;
    unsigned int folderViewCommandId = 0;
};

struct PendingOpenViewer
{
    std::wstring pluginId;
    HWND ownerWindow = nullptr;
    wil::com_ptr<IFileSystem> fileSystem;
    std::wstring fileSystemName;
    std::wstring focusedPath;
    std::vector<std::wstring> selectionStorage;
    std::vector<const wchar_t*> selectionPointers;
    std::vector<std::wstring> otherFilesStorage;
    std::vector<const wchar_t*> otherFilePointers;
    unsigned long focusedOtherFileIndex = 0;
    uint32_t viewerFlags                = 0;
};

[[nodiscard]] bool IsAutomationModeFromCommandLine() noexcept
{
    static std::once_flag initOnce;
    static bool cached = false;

    std::call_once(initOnce,
                   []() noexcept
    {
        bool result = false;

        int argc = 0;
        wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
        if (argv && argc > 1)
        {
            constexpr std::wstring_view kArgs[] = {
                L"--selftest",
                L"--compare-selftest",
                L"--commands-selftest",
                L"--fileops-selftest",
            };

            for (int i = 1; i < argc && ! result; ++i)
            {
                const wchar_t* arg = argv.get()[i];
                if (! arg || arg[0] == L'\0')
                {
                    continue;
                }

                const std::wstring_view argView(arg);
                for (const std::wstring_view needle : kArgs)
                {
                    if (OrdinalString::EqualsNoCase(argView, needle))
                    {
                        result = true;
                        break;
                    }
                }
            }
        }

        cached = result;
    });

    return cached;
}

FolderView::OverlaySeverity ToFolderOverlaySeverity(HostAlertSeverity severity) noexcept
{
    switch (severity)
    {
        case HOST_ALERT_ERROR: return FolderView::OverlaySeverity::Error;
        case HOST_ALERT_WARNING: return FolderView::OverlaySeverity::Warning;
        case HOST_ALERT_INFO: return FolderView::OverlaySeverity::Information;
        case HOST_ALERT_BUSY: return FolderView::OverlaySeverity::Busy;
        default: return FolderView::OverlaySeverity::Error;
    }
}

[[nodiscard]] HostPromptResult DefaultPromptResultForButtons(HostPromptButtons buttons) noexcept
{
    switch (buttons)
    {
        case HOST_PROMPT_BUTTONS_OK: return HOST_PROMPT_RESULT_OK;
        case HOST_PROMPT_BUTTONS_OK_CANCEL: return HOST_PROMPT_RESULT_OK;
        case HOST_PROMPT_BUTTONS_YES_NO: return HOST_PROMPT_RESULT_YES;
        case HOST_PROMPT_BUTTONS_YES_NO_CANCEL: return HOST_PROMPT_RESULT_YES;
        default: return HOST_PROMPT_RESULT_OK;
    }
}

[[nodiscard]] HostPromptResult EscapePromptResultForButtons(HostPromptButtons buttons) noexcept
{
    switch (buttons)
    {
        case HOST_PROMPT_BUTTONS_OK: return HOST_PROMPT_RESULT_OK;
        case HOST_PROMPT_BUTTONS_OK_CANCEL: return HOST_PROMPT_RESULT_CANCEL;
        case HOST_PROMPT_BUTTONS_YES_NO: return HOST_PROMPT_RESULT_NO;
        case HOST_PROMPT_BUTTONS_YES_NO_CANCEL: return HOST_PROMPT_RESULT_CANCEL;
        default: return HOST_PROMPT_RESULT_CANCEL;
    }
}

[[nodiscard]] bool PromptButtonsSupportResult(HostPromptButtons buttons, HostPromptResult result) noexcept
{
    switch (buttons)
    {
        case HOST_PROMPT_BUTTONS_OK: return result == HOST_PROMPT_RESULT_OK;
        case HOST_PROMPT_BUTTONS_OK_CANCEL: return result == HOST_PROMPT_RESULT_OK || result == HOST_PROMPT_RESULT_CANCEL;
        case HOST_PROMPT_BUTTONS_YES_NO: return result == HOST_PROMPT_RESULT_YES || result == HOST_PROMPT_RESULT_NO;
        case HOST_PROMPT_BUTTONS_YES_NO_CANCEL:
            return result == HOST_PROMPT_RESULT_YES || result == HOST_PROMPT_RESULT_NO || result == HOST_PROMPT_RESULT_CANCEL;
        default: return false;
    }
}

[[nodiscard]] HWND GetInitializedHostWindow() noexcept
{
    const HWND hostWindow = HostFolderWindowHwnd().load(std::memory_order_acquire);
    if (! hostWindow || ! IsWindow(hostWindow))
    {
        return nullptr;
    }

    return hostWindow;
}

[[nodiscard]] bool IsCurrentThreadWindowThread(HWND window) noexcept
{
    if (! window)
    {
        return false;
    }

    const DWORD windowThreadId = GetWindowThreadProcessId(window, nullptr);
    return windowThreadId != 0 && windowThreadId == GetCurrentThreadId();
}

[[nodiscard]] HRESULT EnsureHostUiThreadReady(HWND& hostWindow) noexcept
{
    hostWindow = GetInitializedHostWindow();
    if (! hostWindow)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    if (! IsCurrentThreadWindowThread(hostWindow))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_THREAD_ID);
    }

    return S_OK;
}

[[nodiscard]] HRESULT EnsureHostUiThreadReady() noexcept
{
    HWND hostWindow = nullptr;
    return EnsureHostUiThreadReady(hostWindow);
}

[[nodiscard]] std::optional<FolderWindow::Pane> ResolveHostPaneCookie(void* cookie) noexcept
{
    switch (reinterpret_cast<uintptr_t>(cookie))
    {
        case HOST_PANE_COOKIE_LEFT: return FolderWindow::Pane::Left;
        case HOST_PANE_COOKIE_RIGHT: return FolderWindow::Pane::Right;
        default: return std::nullopt;
    }
}

[[nodiscard]] FolderWindow::Pane ResolvePaneScopeTarget(void* cookie) noexcept
{
    return ResolveHostPaneCookie(cookie).value_or(HostFolderWindow().GetFocusedPane());
}

HWND ResolvePromptTargetWindow(const HostPromptRequest& request, void* cookie) noexcept
{
    if (request.scope == HOST_ALERT_SCOPE_WINDOW && request.targetWindow && IsWindow(request.targetWindow))
    {
        return request.targetWindow;
    }

    const HWND hostWindow = GetInitializedHostWindow();

    if (request.scope == HOST_ALERT_SCOPE_PANE_CONTENT || request.scope == HOST_ALERT_SCOPE_PANE)
    {
        if (hostWindow && IsCurrentThreadWindowThread(hostWindow))
        {
            const FolderWindow::Pane pane = ResolvePaneScopeTarget(cookie);
            const HWND folderView         = HostFolderWindow().GetFolderViewHwnd(pane);
            if (folderView && IsWindow(folderView))
            {
                return folderView;
            }
        }
    }

    return hostWindow;
}

RedSalamander::Ui::AlertSeverity ToUiAlertSeverity(HostAlertSeverity severity) noexcept
{
    switch (severity)
    {
        case HOST_ALERT_WARNING: return RedSalamander::Ui::AlertSeverity::Warning;
        case HOST_ALERT_INFO: return RedSalamander::Ui::AlertSeverity::Info;
        case HOST_ALERT_BUSY: return RedSalamander::Ui::AlertSeverity::Busy;
        case HOST_ALERT_ERROR:
        default: return RedSalamander::Ui::AlertSeverity::Error;
    }
}

RedSalamander::Ui::AlertTheme BuildHostAlertTheme() noexcept
{
    AppTheme theme = ResolveAppTheme(ThemeMode::System, L"HostServices");
    if (const HWND hostWindow = GetInitializedHostWindow(); hostWindow && IsCurrentThreadWindowThread(hostWindow))
    {
        theme = HostFolderWindow().GetTheme();
    }

    const FolderViewTheme& fv = theme.folderView;

    RedSalamander::Ui::AlertTheme alertTheme{};
    alertTheme.background          = fv.backgroundColor;
    alertTheme.text                = fv.textNormal;
    alertTheme.accent              = fv.focusBorder;
    alertTheme.selectionBackground = fv.itemBackgroundSelected;
    alertTheme.selectionText       = fv.textSelected;
    alertTheme.errorBackground     = fv.errorBackground;
    alertTheme.errorText           = fv.errorText;
    alertTheme.warningBackground   = fv.warningBackground;
    alertTheme.warningText         = fv.warningText;
    alertTheme.infoBackground      = fv.infoBackground;
    alertTheme.infoText            = fv.infoText;
    alertTheme.darkBase            = fv.darkBase;
    alertTheme.highContrast        = theme.highContrast;
    return alertTheme;
}

} // namespace

class HostServices final : public IHost, public IHostAlerts, public IHostPrompts, public IHostConnections, public IHostPaneExecute, public IHostViewers
{
public:
    HostServices() noexcept : _refCount(1)
    {
    }

    ~HostServices() noexcept
    {
        clearAllSessionSecrets();
    }

    void ClearConnectionSessionState() noexcept
    {
        clearAllSessionSecrets();
    }

    HostServices(const HostServices&)            = delete;
    HostServices(HostServices&&)                 = delete;
    HostServices& operator=(const HostServices&) = delete;
    HostServices& operator=(HostServices&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        if (riid == IID_IUnknown || riid == __uuidof(IHost))
        {
            *ppvObject = static_cast<IHost*>(this);
            AddRef();
            return S_OK;
        }

        if (riid == __uuidof(IHostAlerts))
        {
            *ppvObject = static_cast<IHostAlerts*>(this);
            AddRef();
            return S_OK;
        }

        if (riid == __uuidof(IHostPrompts))
        {
            *ppvObject = static_cast<IHostPrompts*>(this);
            AddRef();
            return S_OK;
        }

        if (riid == __uuidof(IHostConnections))
        {
            *ppvObject = static_cast<IHostConnections*>(this);
            AddRef();
            return S_OK;
        }

        if (riid == __uuidof(IHostPaneExecute))
        {
            *ppvObject = static_cast<IHostPaneExecute*>(this);
            AddRef();
            return S_OK;
        }

        if (riid == __uuidof(IHostViewers))
        {
            *ppvObject = static_cast<IHostViewers*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        // Process-lifetime singleton; never delete.
        return --_refCount;
    }

    HRESULT STDMETHODCALLTYPE ShowAlert(const HostAlertRequest* request, void* cookie) noexcept override
    {
        if (! request)
        {
            return E_POINTER;
        }

        if (request->version != 1 || request->sizeBytes < sizeof(HostAlertRequest))
        {
            return E_INVALIDARG;
        }

        if (! request->message || request->message[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data     = std::make_unique<PendingAlert>();
            data->request = *request;
            data->cookie  = cookie;
            if (request->title)
            {
                data->title = request->title;
            }
            data->message = request->message;

            data->request.title   = data->title.empty() ? nullptr : data->title.c_str();
            data->request.message = data->message.c_str();

            if (! PostMessagePayload(hostWindow, WndMsg::kHostShowAlert, 0, std::move(data)))
            {
                const DWORD lastError = GetLastError();
                return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
            }

            return S_OK;
        }

        return ShowAlertOnUiThread(*request, cookie);
    }

    HRESULT STDMETHODCALLTYPE ClearAlert(HostAlertScope scope, void* cookie) noexcept override
    {
        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data    = std::make_unique<PendingClearAlert>();
            data->scope  = scope;
            data->cookie = cookie;

            if (! PostMessagePayload(hostWindow, WndMsg::kHostClearAlert, 0, std::move(data)))
            {
                const DWORD lastError = GetLastError();
                return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
            }

            return S_OK;
        }

        return ClearAlertOnUiThread(scope, cookie);
    }

    HRESULT STDMETHODCALLTYPE ShowPrompt(const HostPromptRequest* request, void* cookie, HostPromptResult* result) noexcept override
    {
        if (! request || ! result)
        {
            return E_POINTER;
        }

        if (request->version != 1 || request->sizeBytes < sizeof(HostPromptRequest))
        {
            return E_INVALIDARG;
        }

        if (! request->message || request->message[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data     = std::make_unique<PendingPrompt>();
            data->request = *request;
            data->cookie  = cookie;
            data->result  = result;
            if (request->title)
            {
                data->title = request->title;
            }
            data->message = request->message;

            data->request.title   = data->title.empty() ? nullptr : data->title.c_str();
            data->request.message = data->message.c_str();

            auto* raw               = data.release();
            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostShowPrompt, 0, reinterpret_cast<LPARAM>(raw));
            if (msgResult != 0)
            {
                // Failure; WndProc returns HRESULT in msgResult.
                return static_cast<HRESULT>(msgResult);
            }

            return S_OK;
        }

        return ShowPromptOnUiThread(*request, cookie, result);
    }

    HRESULT STDMETHODCALLTYPE ShowConnectionManager(const HostConnectionManagerRequest* request, HostConnectionManagerResult* result) noexcept override
    {
        if (! request || ! result)
        {
            return E_POINTER;
        }

        if (request->version != 1 || request->sizeBytes < sizeof(HostConnectionManagerRequest))
        {
            return E_INVALIDARG;
        }

        result->version        = 1;
        result->sizeBytes      = sizeof(HostConnectionManagerResult);
        result->connectionName = nullptr;

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data     = std::make_unique<PendingConnectionManager>();
            data->request = *request;
            data->result  = result;
            if (request->filterPluginId)
            {
                data->filterPluginId = request->filterPluginId;
            }
            data->request.filterPluginId = data->filterPluginId.empty() ? nullptr : data->filterPluginId.c_str();

            auto* raw               = data.release();
            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostShowConnectionManager, 0, reinterpret_cast<LPARAM>(raw));
            if (msgResult != 0)
            {
                return static_cast<HRESULT>(msgResult);
            }
            return result->connectionName ? S_OK : S_FALSE;
        }

        return ShowConnectionManagerOnUiThread(*request, result);
    }

    HRESULT STDMETHODCALLTYPE GetConnectionJsonUtf8(const wchar_t* connectionName, char** jsonUtf8) noexcept override
    {
        if (! jsonUtf8)
        {
            return E_POINTER;
        }
        *jsonUtf8 = nullptr;

        if (! connectionName || connectionName[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data            = std::make_unique<PendingConnectionJson>();
            data->connectionName = connectionName;

            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostGetConnectionJsonUtf8, 0, reinterpret_cast<LPARAM>(data.get()));
            if (msgResult != 0)
            {
                return static_cast<HRESULT>(msgResult);
            }

            if (! data->json)
            {
                return E_FAIL;
            }

            *jsonUtf8 = data->json.release();
            return S_OK;
        }

        return BuildConnectionJsonUtf8(connectionName, jsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE GetConnectionSecret(const wchar_t* connectionName,
                                                  HostConnectionSecretKind kind,
                                                  HWND ownerWindow,
                                                  wchar_t** secretOut) noexcept override
    {
        if (! secretOut)
        {
            return E_POINTER;
        }
        *secretOut = nullptr;

        if (! connectionName || connectionName[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data            = std::make_unique<PendingConnectionSecret>();
            data->connectionName = connectionName;
            data->kind           = kind;
            data->ownerWindow    = ownerWindow;

            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostGetConnectionSecret, 0, reinterpret_cast<LPARAM>(data.get()));
            if (msgResult != 0)
            {
                return static_cast<HRESULT>(msgResult);
            }

            if (! data->secret)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
            }

            *secretOut = data->secret.release();
            return S_OK;
        }

        return GetConnectionSecretOnUiThread(connectionName, kind, ownerWindow, secretOut);
    }

    HRESULT STDMETHODCALLTYPE SetConnectionSecret(const wchar_t* connectionName,
                                                  HostConnectionSecretKind kind,
                                                  const wchar_t* secret,
                                                  BOOL persist) noexcept override
    {
        if (! connectionName || connectionName[0] == L'\0' || ! secret)
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data            = std::make_unique<PendingSetConnectionSecret>();
            data->connectionName = connectionName;
            data->kind           = kind;
            data->secret         = secret;
            data->persist        = persist;

            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostSetConnectionSecret, 0, reinterpret_cast<LPARAM>(data.get()));
            return static_cast<HRESULT>(msgResult);
        }

        return SetConnectionSecretOnUiThread(connectionName, kind, secret, persist);
    }

    HRESULT STDMETHODCALLTYPE DeleteConnectionSecret(const wchar_t* connectionName, HostConnectionSecretKind kind, BOOL deletePersisted) noexcept override
    {
        if (! connectionName || connectionName[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data             = std::make_unique<PendingDeleteConnectionSecret>();
            data->connectionName  = connectionName;
            data->kind            = kind;
            data->deletePersisted = deletePersisted;

            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostDeleteConnectionSecret, 0, reinterpret_cast<LPARAM>(data.get()));
            return static_cast<HRESULT>(msgResult);
        }

        return DeleteConnectionSecretOnUiThread(connectionName, kind, deletePersisted);
    }

    HRESULT STDMETHODCALLTYPE PromptForConnectionSecret(const wchar_t* connectionName,
                                                        HostConnectionSecretKind kind,
                                                        HWND ownerWindow,
                                                        wchar_t** secretOut) noexcept override
    {
        if (! secretOut)
        {
            return E_POINTER;
        }
        *secretOut = nullptr;

        if (! connectionName || connectionName[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data            = std::make_unique<PendingConnectionSecret>();
            data->connectionName = connectionName;
            data->kind           = kind;
            data->ownerWindow    = ownerWindow;

            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostPromptConnectionSecret, 0, reinterpret_cast<LPARAM>(data.get()));
            const HRESULT hr        = static_cast<HRESULT>(msgResult);
            if (FAILED(hr) || hr == S_FALSE)
            {
                return hr;
            }

            if (! data->secret)
            {
                return E_FAIL;
            }

            *secretOut = data->secret.release();
            return S_OK;
        }

        return PromptForConnectionSecretOnUiThread(connectionName, kind, ownerWindow, secretOut);
    }

    HRESULT STDMETHODCALLTYPE ClearCachedConnectionSecret(const wchar_t* connectionName, HostConnectionSecretKind kind) noexcept override
    {
        if (! connectionName || connectionName[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data            = std::make_unique<PendingClearConnectionSecretCache>();
            data->connectionName = connectionName;
            data->kind           = kind;

            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostClearCachedConnectionSecret, 0, reinterpret_cast<LPARAM>(data.get()));
            return static_cast<HRESULT>(msgResult);
        }

        return ClearCachedConnectionSecretOnUiThread(connectionName, kind);
    }

    HRESULT STDMETHODCALLTYPE UpgradeFtpAnonymousToPassword(const wchar_t* connectionName, HWND ownerWindow) noexcept override
    {
        if (! connectionName || connectionName[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data            = std::make_unique<PendingUpgradeFtpAnonymousToPassword>();
            data->connectionName = connectionName;
            data->ownerWindow    = ownerWindow;

            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostUpgradeFtpAnonymousToPassword, 0, reinterpret_cast<LPARAM>(data.get()));
            return static_cast<HRESULT>(msgResult);
        }

        return UpgradeFtpAnonymousToPasswordOnUiThread(connectionName, ownerWindow);
    }

    HRESULT STDMETHODCALLTYPE ExecuteInActivePane(const HostPaneExecuteRequest* request) noexcept override
    {
        if (! request)
        {
            return E_POINTER;
        }

        if (request->version != 1 || request->sizeBytes < sizeof(HostPaneExecuteRequest))
        {
            return E_INVALIDARG;
        }

        if (! request->folderPath || request->folderPath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        std::wstring_view focusName;
        if (request->focusItemDisplayName)
        {
            focusName = request->focusItemDisplayName;
            if (focusName.empty())
            {
                return E_INVALIDARG;
            }
            if (focusName.find_first_of(L"/\\") != std::wstring_view::npos)
            {
                return E_INVALIDARG;
            }
        }

        if (request->folderViewCommandId > 0xFFFFu)
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data                 = std::make_unique<PendingExecuteInPane>();
            data->flags               = request->flags;
            data->folderViewCommandId = request->folderViewCommandId;
            data->folderPath          = request->folderPath;
            if (! focusName.empty())
            {
                data->focusItemDisplayName.assign(focusName);
            }

            if (! PostMessagePayload(hostWindow, WndMsg::kHostExecuteInPane, 0, std::move(data)))
            {
                const DWORD lastError = GetLastError();
                return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
            }

            return S_OK;
        }

        const bool activateWindow = (request->flags & HOST_PANE_EXECUTE_FLAG_ACTIVATE_WINDOW) != 0;
        return HostFolderWindow().ExecuteInActivePane(std::filesystem::path(request->folderPath), focusName, request->folderViewCommandId, activateWindow);
    }

    HRESULT STDMETHODCALLTYPE OpenViewer(const HostViewerOpenRequest* request) noexcept override
    {
        if (! request)
        {
            return E_POINTER;
        }

        if (request->version != 1 || request->sizeBytes < sizeof(HostViewerOpenRequest))
        {
            return E_INVALIDARG;
        }

        if (! request->pluginId || request->pluginId[0] == L'\0' || ! request->fileSystem || ! request->focusedPath || request->focusedPath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        if ((request->selectionCount > 0 && ! request->selectionPaths) || (request->otherFileCount > 0 && ! request->otherFiles))
        {
            return E_INVALIDARG;
        }

        const HWND hostWindow = GetInitializedHostWindow();
        if (! hostWindow)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        if (! IsCurrentThreadWindowThread(hostWindow))
        {
            auto data                   = std::make_unique<PendingOpenViewer>();
            data->pluginId              = request->pluginId;
            data->ownerWindow           = request->ownerWindow;
            data->fileSystem            = request->fileSystem;
            data->fileSystemName        = request->fileSystemName ? request->fileSystemName : L"";
            data->focusedPath           = request->focusedPath;
            data->focusedOtherFileIndex = request->focusedOtherFileIndex;
            data->viewerFlags           = request->viewerFlags;

            if (request->selectionPaths && request->selectionCount > 0)
            {
                data->selectionStorage.reserve(request->selectionCount);
                for (unsigned long i = 0; i < request->selectionCount; ++i)
                {
                    const wchar_t* path = request->selectionPaths[i];
                    if (path && path[0] != L'\0')
                    {
                        data->selectionStorage.emplace_back(path);
                    }
                }

                data->selectionPointers.reserve(data->selectionStorage.size());
                for (const auto& path : data->selectionStorage)
                {
                    data->selectionPointers.push_back(path.c_str());
                }
            }

            if (request->otherFiles && request->otherFileCount > 0)
            {
                data->otherFilesStorage.reserve(request->otherFileCount);
                for (unsigned long i = 0; i < request->otherFileCount; ++i)
                {
                    const wchar_t* path = request->otherFiles[i];
                    if (path && path[0] != L'\0')
                    {
                        data->otherFilesStorage.emplace_back(path);
                    }
                }

                data->otherFilePointers.reserve(data->otherFilesStorage.size());
                for (const auto& path : data->otherFilesStorage)
                {
                    data->otherFilePointers.push_back(path.c_str());
                }
            }

            const LRESULT msgResult = SendMessageW(hostWindow, WndMsg::kHostOpenViewer, 0, reinterpret_cast<LPARAM>(data.get()));
            return static_cast<HRESULT>(msgResult);
        }

        ViewerOpenContext context{};
        context.ownerWindow           = request->ownerWindow;
        context.fileSystem            = request->fileSystem;
        context.fileSystemName        = request->fileSystemName;
        context.focusedPath           = request->focusedPath;
        context.selectionPaths        = request->selectionPaths;
        context.selectionCount        = request->selectionCount;
        context.otherFiles            = request->otherFiles;
        context.otherFileCount        = request->otherFileCount;
        context.focusedOtherFileIndex = request->focusedOtherFileIndex;
        context.flags                 = static_cast<ViewerOpenFlags>(request->viewerFlags);
        return HostFolderWindow().OpenViewerWithPlugin(request->pluginId, context);
    }

    bool TryHandleMessage(UINT message, WPARAM /*wParam*/, LPARAM lParam, LRESULT& result) noexcept
    {
        if (message == WndMsg::kHostShowAlert)
        {
            auto data = TakeMessagePayload<PendingAlert>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            static_cast<void>(ShowAlertOnUiThread(data->request, data->cookie));
            result = 0;
            return true;
        }

        if (message == WndMsg::kHostClearAlert)
        {
            auto data = TakeMessagePayload<PendingClearAlert>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            static_cast<void>(ClearAlertOnUiThread(data->scope, data->cookie));
            result = 0;
            return true;
        }

        if (message == WndMsg::kHostShowPrompt)
        {
            auto data = TakeMessagePayload<PendingPrompt>(lParam);
            if (! data || ! data->result)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            const HRESULT hr = ShowPromptOnUiThread(data->request, data->cookie, data->result);
            result           = static_cast<LRESULT>(hr);
            return true;
        }

        if (message == WndMsg::kHostShowConnectionManager)
        {
            auto data = TakeMessagePayload<PendingConnectionManager>(lParam);
            if (! data || ! data->result)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            const HRESULT hr = ShowConnectionManagerOnUiThread(data->request, data->result);
            result           = static_cast<LRESULT>(FAILED(hr) ? hr : 0);
            return true;
        }

        if (message == WndMsg::kHostGetConnectionJsonUtf8)
        {
            auto* data = reinterpret_cast<PendingConnectionJson*>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            char* rawJson    = nullptr;
            const HRESULT hr = BuildConnectionJsonUtf8(data->connectionName.c_str(), &rawJson);
            if (SUCCEEDED(hr))
            {
                data->json.reset(rawJson);
                if (! data->json)
                {
                    result = static_cast<LRESULT>(E_FAIL);
                    return true;
                }
            }

            result = static_cast<LRESULT>(FAILED(hr) ? hr : 0);
            return true;
        }

        if (message == WndMsg::kHostGetConnectionSecret)
        {
            auto* data = reinterpret_cast<PendingConnectionSecret*>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            const HRESULT hr = GetConnectionSecretOnUiThread(data->connectionName.c_str(), data->kind, data->ownerWindow, data->secret.put());
            result           = static_cast<LRESULT>(FAILED(hr) ? hr : 0);
            return true;
        }

        if (message == WndMsg::kHostSetConnectionSecret)
        {
            auto* data = reinterpret_cast<PendingSetConnectionSecret*>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            const HRESULT hr = SetConnectionSecretOnUiThread(data->connectionName.c_str(), data->kind, data->secret.c_str(), data->persist);
            result           = static_cast<LRESULT>(hr);
            return true;
        }

        if (message == WndMsg::kHostDeleteConnectionSecret)
        {
            auto* data = reinterpret_cast<PendingDeleteConnectionSecret*>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            const HRESULT hr = DeleteConnectionSecretOnUiThread(data->connectionName.c_str(), data->kind, data->deletePersisted);
            result           = static_cast<LRESULT>(hr);
            return true;
        }

        if (message == WndMsg::kHostPromptConnectionSecret)
        {
            auto* data = reinterpret_cast<PendingConnectionSecret*>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            const HRESULT hr = PromptForConnectionSecretOnUiThread(data->connectionName.c_str(), data->kind, data->ownerWindow, data->secret.put());
            result           = static_cast<LRESULT>(hr);
            return true;
        }

        if (message == WndMsg::kHostClearCachedConnectionSecret)
        {
            auto* data = reinterpret_cast<PendingClearConnectionSecretCache*>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            const HRESULT hr = ClearCachedConnectionSecretOnUiThread(data->connectionName.c_str(), data->kind);
            result           = static_cast<LRESULT>(hr);
            return true;
        }

        if (message == WndMsg::kHostUpgradeFtpAnonymousToPassword)
        {
            auto* data = reinterpret_cast<PendingUpgradeFtpAnonymousToPassword*>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            const HRESULT hr = UpgradeFtpAnonymousToPasswordOnUiThread(data->connectionName.c_str(), data->ownerWindow);
            result           = static_cast<LRESULT>(hr);
            return true;
        }

        if (message == WndMsg::kHostExecuteInPane)
        {
            auto data = TakeMessagePayload<PendingExecuteInPane>(lParam);
            if (! data)
            {
                result = static_cast<LRESULT>(E_POINTER);
                return true;
            }

            const bool activateWindow = (data->flags & HOST_PANE_EXECUTE_FLAG_ACTIVATE_WINDOW) != 0;
            const HRESULT hr          = HostFolderWindow().ExecuteInActivePane(
                std::filesystem::path(data->folderPath), std::wstring_view(data->focusItemDisplayName), data->folderViewCommandId, activateWindow);

            result = static_cast<LRESULT>(hr);
            return true;
        }

        if (message == WndMsg::kHostOpenViewer)
        {
            auto* data = reinterpret_cast<PendingOpenViewer*>(lParam);
            if (! data || ! data->fileSystem || data->pluginId.empty() || data->focusedPath.empty())
            {
                result = static_cast<LRESULT>(E_INVALIDARG);
                return true;
            }

            ViewerOpenContext context{};
            context.ownerWindow           = data->ownerWindow;
            context.fileSystem            = data->fileSystem.get();
            context.fileSystemName        = data->fileSystemName.empty() ? nullptr : data->fileSystemName.c_str();
            context.focusedPath           = data->focusedPath.c_str();
            context.selectionPaths        = data->selectionPointers.empty() ? nullptr : data->selectionPointers.data();
            context.selectionCount        = static_cast<unsigned long>(data->selectionPointers.size());
            context.otherFiles            = data->otherFilePointers.empty() ? nullptr : data->otherFilePointers.data();
            context.otherFileCount        = static_cast<unsigned long>(data->otherFilePointers.size());
            context.focusedOtherFileIndex = data->focusedOtherFileIndex;
            context.flags                 = static_cast<ViewerOpenFlags>(data->viewerFlags);

            const HRESULT hr = HostFolderWindow().OpenViewerWithPlugin(data->pluginId, context);
            result           = static_cast<LRESULT>(hr);
            return true;
        }

        return false;
    }

private:
    struct SessionSecretEntry
    {
        bool present = false;
        std::wstring secret;
    };

    [[nodiscard]] static bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept
    {
        if (a.size() != b.size())
        {
            return false;
        }

        for (size_t i = 0; i < a.size(); ++i)
        {
            if (std::towlower(a[i]) != std::towlower(b[i]))
            {
                return false;
            }
        }

        return true;
    }

    [[nodiscard]] static std::string Utf8FromUtf16(std::wstring_view text) noexcept
    {
        return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
    }

    [[nodiscard]] static const Common::Settings::ConnectionProfile* FindConnectionProfile(std::wstring_view connectionName) noexcept
    {
        if (connectionName.empty() || ! HostSettings().connections)
        {
            return nullptr;
        }

        const auto& items = HostSettings().connections->items;

        const auto byName = std::find_if(items.begin(), items.end(), [&](const Common::Settings::ConnectionProfile& c) noexcept {
            return ! c.name.empty() && EqualsIgnoreCase(c.name, connectionName);
        });
        if (byName != items.end())
        {
            return &(*byName);
        }

        return nullptr;
    }

    [[nodiscard]] static Common::Settings::ConnectionProfile* FindConnectionProfileMutable(std::wstring_view connectionName) noexcept
    {
        if (connectionName.empty() || ! HostSettings().connections)
        {
            return nullptr;
        }

        auto& items = HostSettings().connections->items;

        const auto byName = std::find_if(items.begin(), items.end(), [&](Common::Settings::ConnectionProfile& c) noexcept {
            return ! c.name.empty() && EqualsIgnoreCase(c.name, connectionName);
        });
        if (byName != items.end())
        {
            return &(*byName);
        }

        return nullptr;
    }

    [[nodiscard]] static const wchar_t* AuthModeToString(Common::Settings::ConnectionAuthMode mode) noexcept
    {
        switch (mode)
        {
            case Common::Settings::ConnectionAuthMode::Anonymous: return L"anonymous";
            case Common::Settings::ConnectionAuthMode::SshKey: return L"sshKey";
            case Common::Settings::ConnectionAuthMode::OAuth2Pkce: return L"oauth2Pkce";
            case Common::Settings::ConnectionAuthMode::Password:
            default: return L"password";
        }
    }

    [[nodiscard]] static const wchar_t* ConnectionSecretKindToString(HostConnectionSecretKind kind) noexcept
    {
        switch (kind)
        {
            case HOST_CONNECTION_SECRET_PASSWORD: return L"password";
            case HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE: return L"sshKeyPassphrase";
            case HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN: return L"refreshToken";
        }
        return L"password";
    }

    [[nodiscard]] static RedSalamander::Connections::SecretKind ToLocalSecretKind(HostConnectionSecretKind kind) noexcept
    {
        switch (kind)
        {
            case HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE: return RedSalamander::Connections::SecretKind::SshKeyPassphrase;
            case HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN: return RedSalamander::Connections::SecretKind::RefreshToken;
            case HOST_CONNECTION_SECRET_PASSWORD:
            default: return RedSalamander::Connections::SecretKind::Password;
        }
    }

    static void SecureClear(std::wstring& text) noexcept
    {
        if (! text.empty())
        {
            SecureZeroMemory(text.data(), text.size() * sizeof(wchar_t));
        }
        text.clear();
    }

    void clearAllSessionSecrets() noexcept
    {
        for (auto& [id, entry] : _sessionPasswordByConnectionId)
        {
            SecureClear(entry.secret);
        }
        _sessionPasswordByConnectionId.clear();

        for (auto& [id, entry] : _sessionPassphraseByConnectionId)
        {
            SecureClear(entry.secret);
        }
        _sessionPassphraseByConnectionId.clear();

        for (auto& [id, entry] : _sessionRefreshTokenByConnectionId)
        {
            SecureClear(entry.secret);
        }
        _sessionRefreshTokenByConnectionId.clear();
        RedSalamander::Connections::ClearAllSecretAccessAuthorizations();
    }

    HRESULT ShowConnectionManagerOnUiThread(const HostConnectionManagerRequest& request, HostConnectionManagerResult* result) noexcept
    {
        if (! result)
        {
            return E_POINTER;
        }

        HWND hostWindow = nullptr;

        result->version        = 1;
        result->sizeBytes      = sizeof(HostConnectionManagerResult);
        result->connectionName = nullptr;

        HWND owner = request.ownerWindow;
        if (! owner || ! IsWindow(owner))
        {
            owner = hostWindow;
        }

        std::wstring selectedName;
        const std::wstring_view filter = (request.filterPluginId && request.filterPluginId[0] != L'\0') ? request.filterPluginId : L"";
        const HRESULT hr =
            ShowConnectionManagerDialog(owner, HostFolderWindow(), L"RedSalamander", HostSettings(), HostFolderWindow().GetTheme(), filter, selectedName);
        if (hr == S_FALSE)
        {
            return S_FALSE;
        }
        if (FAILED(hr))
        {
            return hr;
        }

        if (selectedName.empty())
        {
            return E_FAIL;
        }

        const size_t bytes = (selectedName.size() + 1u) * sizeof(wchar_t);
        wil::unique_cotaskmem_string mem(static_cast<wchar_t*>(CoTaskMemAlloc(bytes)));
        if (! mem)
        {
            return E_OUTOFMEMORY;
        }
        memcpy(mem.get(), selectedName.c_str(), bytes);
        result->connectionName = mem.release();
        return S_OK;
    }

    HRESULT BuildConnectionJsonUtf8(const wchar_t* connectionName, char** jsonUtf8) noexcept
    {
        if (! jsonUtf8)
        {
            return E_POINTER;
        }
        *jsonUtf8 = nullptr;

        const HRESULT hrHostReady = EnsureHostUiThreadReady();
        if (FAILED(hrHostReady))
        {
            return hrHostReady;
        }

        const Common::Settings::ConnectionProfile* profile = FindConnectionProfile(connectionName);
        if (! profile)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }

        if (IsAutomationModeFromCommandLine() && ConnectionProfileUtils::ConnectionProfileUsesInsecureTls(*profile))
        {
            const bool allowInAutomation = HostSettings().connections ? HostSettings().connections->allowInsecureTlsInAutomation : false;
            const bool acked             = ConnectionProfileUtils::ExtraGetBool(profile->extra, "insecureTlsAck").value_or(false);
            if (! allowInAutomation || ! acked)
            {
                Debug::Warning(L"Blocked insecure TLS connection in automation: connection='{}' pluginId='{}' allowInsecureTlsInAutomation={} acked={}",
                               profile->name,
                               profile->pluginId,
                               allowInAutomation ? 1 : 0,
                               acked ? 1 : 0);
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }
        }

        yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
        if (! doc)
        {
            return E_OUTOFMEMORY;
        }

        auto freeDoc         = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });
        yyjson_mut_val* root = yyjson_mut_obj(doc);
        if (! root)
        {
            return E_OUTOFMEMORY;
        }
        yyjson_mut_doc_set_root(doc, root);

        auto addWideString = [&](const char* key, std::wstring_view value) -> HRESULT
        {
            const std::string utf8 = Utf8FromUtf16(value);
            if (utf8.empty() && ! value.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
            }

            yyjson_mut_val* str = yyjson_mut_strncpy(doc, utf8.c_str(), utf8.size());
            if (! str)
            {
                return E_OUTOFMEMORY;
            }

            yyjson_mut_obj_add_val(doc, root, key, str);
            return S_OK;
        };

        HRESULT hr = addWideString("id", profile->id);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = addWideString("name", profile->name);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = addWideString("pluginId", profile->pluginId);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = addWideString("host", profile->host);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = addWideString("initialPath", profile->initialPath.empty() ? L"/" : profile->initialPath);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = addWideString("userName", profile->userName);
        if (FAILED(hr))
        {
            return hr;
        }

        yyjson_mut_obj_add_uint(doc, root, "port", profile->port);
        yyjson_mut_obj_add_bool(doc, root, "savePassword", profile->savePassword);
        yyjson_mut_obj_add_bool(doc, root, "requireWindowsHello", profile->requireWindowsHello);

        const wchar_t* authModeWide = AuthModeToString(profile->authMode);
        if (authModeWide)
        {
            const std::string authModeUtf8 = Utf8FromUtf16(authModeWide);
            if (! authModeUtf8.empty())
            {
                yyjson_mut_obj_add_strcpy(doc, root, "authMode", authModeUtf8.c_str());
            }
        }

        if (const auto keyPath = Common::Settings::GetWString(profile->extra, "sshPrivateKey"); keyPath.has_value() && ! keyPath->empty())
        {
            static_cast<void>(addWideString("sshPrivateKey", *keyPath));
        }
        if (const auto knownHosts = Common::Settings::GetWString(profile->extra, "sshKnownHosts"); knownHosts.has_value() && ! knownHosts->empty())
        {
            static_cast<void>(addWideString("sshKnownHosts", *knownHosts));
        }
        if (profile->pluginId == L"builtin/file-system-imap")
        {
            if (const auto ignoreSslTrust = ConnectionProfileUtils::ExtraGetBool(profile->extra, "ignoreSslTrust"); ignoreSslTrust.has_value())
            {
                yyjson_mut_obj_add_bool(doc, root, "ignoreSslTrust", ignoreSslTrust.value());
            }
        }

        // Full plugin-specific extra payload (best-effort; intended for plugins and advanced settings).
        {
            std::string extraJson;
            if (SUCCEEDED(Common::Settings::SerializeJsonValue(profile->extra, extraJson)) && ! extraJson.empty())
            {
                yyjson_read_err extraErr{};
                yyjson_doc* extraDoc = yyjson_read_opts(extraJson.data(), extraJson.size(), YYJSON_READ_ALLOW_BOM, nullptr, &extraErr);
                if (extraDoc)
                {
                    auto freeExtraDoc = wil::scope_exit([&] { yyjson_doc_free(extraDoc); });
                    if (yyjson_val* extraRoot = yyjson_doc_get_root(extraDoc); extraRoot && yyjson_is_obj(extraRoot))
                    {
                        if (yyjson_mut_val* extraCopy = yyjson_val_mut_copy(doc, extraRoot))
                        {
                            yyjson_mut_obj_add_val(doc, root, "extra", extraCopy);
                        }
                    }
                }
            }
        }

        yyjson_write_err writeErr{};
        size_t jsonLen = 0;
        char* json     = yyjson_mut_write_opts(doc, YYJSON_WRITE_NOFLAG, nullptr, &jsonLen, &writeErr);
        if (! json)
        {
            return writeErr.code == YYJSON_WRITE_ERROR_MEMORY_ALLOCATION ? E_OUTOFMEMORY : E_FAIL;
        }

        auto freeJson = wil::scope_exit([&] { free(json); });

        wil::unique_cotaskmem_ptr<char> out(static_cast<char*>(CoTaskMemAlloc(jsonLen + 1u)));
        if (! out)
        {
            return E_OUTOFMEMORY;
        }

        memcpy(out.get(), json, jsonLen);
        out.get()[jsonLen] = '\0';
        *jsonUtf8          = out.release();
        return S_OK;
    }

    HRESULT GetConnectionSecretOnUiThread(const wchar_t* connectionName, HostConnectionSecretKind kind, HWND ownerWindow, wchar_t** secretOut) noexcept
    {
        if (! secretOut)
        {
            return E_POINTER;
        }
        *secretOut = nullptr;

        HWND hostWindow           = nullptr;
        const HRESULT hrHostReady = EnsureHostUiThreadReady(hostWindow);
        if (FAILED(hrHostReady))
        {
            return hrHostReady;
        }

        const Common::Settings::ConnectionProfile* profile = FindConnectionProfile(connectionName);
        if (! profile)
        {
            Debug::Error(L"GetConnectionSecret failed: connection not found: '{}'", connectionName ? connectionName : L"(null)");
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }

        const wchar_t* kindText = ConnectionSecretKindToString(kind);
        auto* sessionMap        = &_sessionPasswordByConnectionId;
        if (kind == HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE)
        {
            sessionMap = &_sessionPassphraseByConnectionId;
        }
        else if (kind == HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN)
        {
            sessionMap = &_sessionRefreshTokenByConnectionId;
        }

        // Session cache takes priority (allows ephemeral secrets even when not persisted).
        if (! profile->id.empty())
        {
            auto& map = *sessionMap;
            if (const auto it = map.find(profile->id); it != map.end())
            {
                if (! it->second.present)
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
                }
                const std::wstring_view cached = it->second.secret;
                const size_t bytes             = (cached.size() + 1u) * sizeof(wchar_t);
                wil::unique_cotaskmem_string mem(static_cast<wchar_t*>(CoTaskMemAlloc(bytes)));
                if (! mem)
                {
                    return E_OUTOFMEMORY;
                }

                if (! cached.empty())
                {
                    memcpy(mem.get(), cached.data(), cached.size() * sizeof(wchar_t));
                }
                mem.get()[cached.size()] = L'\0';
                *secretOut               = mem.release();
                return S_OK;
            }
        }

        HWND owner = ownerWindow;

        const Common::Settings::ConnectionsSettings defaults{};
        bool bypassWindowsHello                  = false;
        uint32_t windowsHelloReauthTimeoutMinute = defaults.windowsHelloReauthTimeoutMinute;
        if (HostSettings().connections)
        {
            bypassWindowsHello              = HostSettings().connections->bypassWindowsHello;
            windowsHelloReauthTimeoutMinute = HostSettings().connections->windowsHelloReauthTimeoutMinute;
        }

        const RedSalamander::Connections::SecretKind secretKind = ToLocalSecretKind(kind);

        const bool isQuickConnect = RedSalamander::Connections::IsQuickConnectConnectionId(profile->id);

        if (isQuickConnect)
        {
            std::wstring secret;
            const HRESULT loadHr = RedSalamander::Connections::LoadQuickConnectSecret(secretKind, secret);
            if (FAILED(loadHr))
            {
                Debug::Error(L"GetConnectionSecret failed: LoadQuickConnectSecret failed for connection '{}' (id={}) hr=0x{:08X}",
                             profile->name,
                             profile->id,
                             static_cast<unsigned long>(loadHr));
                return loadHr;
            }

            Debug::Info(
                L"GetConnectionSecret loaded connection='{}' id='{}' kind='{}' secretPresent={}", profile->name, profile->id, kindText, secret.empty() ? 0 : 1);

            const size_t bytes = (secret.size() + 1u) * sizeof(wchar_t);
            wil::unique_cotaskmem_string mem(static_cast<wchar_t*>(CoTaskMemAlloc(bytes)));
            if (! mem)
            {
                return E_OUTOFMEMORY;
            }

            memcpy(mem.get(), secret.c_str(), bytes);
            *secretOut = mem.release();
            return S_OK;
        }

        if (! profile->savePassword)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }

        Debug::Info(L"GetConnectionSecret request connection='{}' id='{}' kind='{}' requireHello={} bypassHello={} reauthTimeoutMinute={}",
                    profile->name,
                    profile->id,
                    kindText,
                    profile->requireWindowsHello ? 1 : 0,
                    bypassWindowsHello ? 1 : 0,
                    windowsHelloReauthTimeoutMinute);

        if (profile->requireWindowsHello && ! bypassWindowsHello)
        {
            const bool shouldPrompt = ! RedSalamander::Connections::IsSecretAccessAuthorized(
                profile->id, secretKind, RedSalamander::Connections::SecretAccessPurpose::Background, 0u);

            if (shouldPrompt)
            {
                if (! owner || ! IsWindow(owner))
                {
                    HWND fallbackWindow     = nullptr;
                    const HRESULT hostReady = EnsureHostUiThreadReady(fallbackWindow);
                    if (SUCCEEDED(hostReady))
                    {
                        hostWindow = fallbackWindow;
                        owner      = hostWindow;
                    }
                }

                const HRESULT helloHr =
                    RedSalamander::Security::VerifyWindowsHelloForWindow(owner, LoadStringResource(nullptr, IDS_CONNECTIONS_HELLO_PROMPT_CREDENTIAL));
                if (FAILED(helloHr))
                {
                    Debug::Warning(L"GetConnectionSecret: Windows Hello verification failed for connection '{}' (id={}) hr=0x{:08X}",
                                   profile->name,
                                   profile->id,
                                   static_cast<unsigned long>(helloHr));
                    return helloHr;
                }

                if (! profile->id.empty())
                {
                    RedSalamander::Connections::NoteSecretAccessAuthorized(profile->id, secretKind);
                }
            }
        }

        std::wstring targetName;
        targetName = RedSalamander::Connections::BuildCredentialTargetName(profile->id, secretKind);

        std::wstring userName;
        std::wstring secret;
        const HRESULT loadHr = RedSalamander::Connections::LoadGenericCredential(targetName, userName, secret);
        if (FAILED(loadHr))
        {
            Debug::Error(L"GetConnectionSecret failed: LoadGenericCredential failed for connection '{}' (id={}) target='{}' hr=0x{:08X}",
                         profile->name,
                         profile->id,
                         targetName,
                         static_cast<unsigned long>(loadHr));
            return loadHr;
        }

        Debug::Info(
            L"GetConnectionSecret loaded connection='{}' id='{}' kind='{}' secretPresent={}", profile->name, profile->id, kindText, secret.empty() ? 0 : 1);

        const size_t bytes = (secret.size() + 1u) * sizeof(wchar_t);
        wil::unique_cotaskmem_string mem(static_cast<wchar_t*>(CoTaskMemAlloc(bytes)));
        if (! mem)
        {
            return E_OUTOFMEMORY;
        }

        memcpy(mem.get(), secret.c_str(), bytes);
        *secretOut = mem.release();
        return S_OK;
    }

    HRESULT PromptForConnectionSecretOnUiThread(const wchar_t* connectionName, HostConnectionSecretKind kind, HWND ownerWindow, wchar_t** secretOut) noexcept;

    HRESULT SetConnectionSecretOnUiThread(const wchar_t* connectionName, HostConnectionSecretKind kind, const wchar_t* secret, BOOL persist) noexcept;

    HRESULT DeleteConnectionSecretOnUiThread(const wchar_t* connectionName, HostConnectionSecretKind kind, BOOL deletePersisted) noexcept;

    HRESULT ClearCachedConnectionSecretOnUiThread(const wchar_t* connectionName, HostConnectionSecretKind kind) noexcept;

    HRESULT UpgradeFtpAnonymousToPasswordOnUiThread(const wchar_t* connectionName, HWND ownerWindow) noexcept;

    HRESULT ShowAlertOnUiThread(const HostAlertRequest& request, void* cookie) noexcept
    {
        const bool blocksInput = request.modality == HOST_ALERT_MODAL;
        const bool closable    = request.closable != FALSE;

        std::wstring title;
        if (request.title && request.title[0] != L'\0')
        {
            title = request.title;
        }

        std::wstring message;
        if (request.message)
        {
            message = request.message;
        }

        if (message.empty())
        {
            return E_INVALIDARG;
        }

        HWND hostWindow = nullptr;
        if (request.scope != HOST_ALERT_SCOPE_WINDOW)
        {
            const HRESULT hrHostReady = EnsureHostUiThreadReady(hostWindow);
            if (FAILED(hrHostReady))
            {
                return hrHostReady;
            }
        }

        const RedSalamander::Ui::AlertTheme theme = BuildHostAlertTheme();

        RedSalamander::Ui::AlertModel model{};
        model.severity = ToUiAlertSeverity(request.severity);
        model.title    = std::move(title);
        model.message  = std::move(message);
        model.closable = closable;

        if (request.scope == HOST_ALERT_SCOPE_APPLICATION)
        {
            if (! _applicationOverlay)
            {
                _applicationOverlay = std::make_unique<RedSalamander::Ui::AlertOverlayWindow>();
            }
            return _applicationOverlay->ShowForParentClient(hostWindow, theme, std::move(model), blocksInput);
        }

        if (request.scope == HOST_ALERT_SCOPE_WINDOW)
        {
            if (! request.targetWindow || ! IsWindow(request.targetWindow))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
            }

            auto& overlay = _windowOverlays[request.targetWindow];
            if (! overlay)
            {
                overlay = std::make_unique<RedSalamander::Ui::AlertOverlayWindow>();
            }

            return overlay->ShowForAnchor(request.targetWindow, theme, std::move(model), blocksInput);
        }

        const FolderView::OverlaySeverity folderSeverity = ToFolderOverlaySeverity(request.severity);
        const FolderWindow::Pane pane                    = ResolvePaneScopeTarget(cookie);
        HostFolderWindow().ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, folderSeverity, std::move(model.title), std::move(model.message), S_OK, closable, blocksInput);
        return S_OK;
    }

    HRESULT ClearAlertOnUiThread(HostAlertScope scope, void* cookie) noexcept
    {
        if (scope == HOST_ALERT_SCOPE_APPLICATION)
        {
            if (_applicationOverlay)
            {
                _applicationOverlay->Hide();
            }
            return S_OK;
        }

        if (scope == HOST_ALERT_SCOPE_WINDOW)
        {
            HWND targetWindow = reinterpret_cast<HWND>(cookie);
            if (! targetWindow)
            {
                return E_INVALIDARG;
            }

            const auto it = _windowOverlays.find(targetWindow);
            if (it != _windowOverlays.end() && it->second)
            {
                it->second->Hide();
            }
            return S_OK;
        }

        HWND hostWindow           = nullptr;
        const HRESULT hrHostReady = EnsureHostUiThreadReady(hostWindow);
        if (FAILED(hrHostReady))
        {
            return hrHostReady;
        }

        const FolderWindow::Pane pane = ResolvePaneScopeTarget(cookie);
        HostFolderWindow().DismissPaneAlertOverlay(pane);
        return S_OK;
    }

    HRESULT ShowPromptOnUiThread(const HostPromptRequest& request, void* cookie, HostPromptResult* result) noexcept
    {
        if (! result)
        {
            return E_POINTER;
        }

        HWND targetWindow = ResolvePromptTargetWindow(request, cookie);
        if (! targetWindow || ! IsWindow(targetWindow))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        }

        std::wstring caption;
        if (request.title && request.title[0] != L'\0')
        {
            caption = request.title;
        }

        std::wstring text = request.message ? request.message : L"";
        if (text.empty())
        {
            return E_INVALIDARG;
        }

        HostPromptResult primaryResult = DefaultPromptResultForButtons(request.buttons);
        if (request.defaultResult != HOST_PROMPT_RESULT_NONE && PromptButtonsSupportResult(request.buttons, request.defaultResult))
        {
            primaryResult = request.defaultResult;
        }

        const HostPromptResult escapeResult = EscapePromptResultForButtons(request.buttons);

        RedSalamander::Ui::AlertModel model{};
        model.severity = ToUiAlertSeverity(request.severity);
        model.title    = std::move(caption);
        model.message  = std::move(text);
        model.closable = true;

        const std::wstring labelOk     = LoadStringResource(nullptr, IDS_BTN_OK);
        const std::wstring labelCancel = LoadStringResource(nullptr, IDS_BTN_CANCEL);
        const std::wstring labelYes    = LoadStringResource(nullptr, IDS_BTN_YES);
        const std::wstring labelNo     = LoadStringResource(nullptr, IDS_BTN_NO);

        auto addButton = [&](uint32_t id, const std::wstring& label, HostPromptResult thisResult) -> void
        {
            RedSalamander::Ui::AlertButton button{};
            button.id      = id;
            button.label   = label;
            button.primary = thisResult == primaryResult;
            model.buttons.push_back(std::move(button));
        };

        switch (request.buttons)
        {
            case HOST_PROMPT_BUTTONS_OK: addButton(static_cast<uint32_t>(IDOK), labelOk, HOST_PROMPT_RESULT_OK); break;
            case HOST_PROMPT_BUTTONS_OK_CANCEL:
                addButton(static_cast<uint32_t>(IDOK), labelOk, HOST_PROMPT_RESULT_OK);
                addButton(static_cast<uint32_t>(IDCANCEL), labelCancel, HOST_PROMPT_RESULT_CANCEL);
                break;
            case HOST_PROMPT_BUTTONS_YES_NO:
                addButton(static_cast<uint32_t>(IDYES), labelYes, HOST_PROMPT_RESULT_YES);
                addButton(static_cast<uint32_t>(IDNO), labelNo, HOST_PROMPT_RESULT_NO);
                break;
            case HOST_PROMPT_BUTTONS_YES_NO_CANCEL:
                addButton(static_cast<uint32_t>(IDYES), labelYes, HOST_PROMPT_RESULT_YES);
                addButton(static_cast<uint32_t>(IDNO), labelNo, HOST_PROMPT_RESULT_NO);
                addButton(static_cast<uint32_t>(IDCANCEL), labelCancel, HOST_PROMPT_RESULT_CANCEL);
                break;
            default: addButton(static_cast<uint32_t>(IDOK), labelOk, HOST_PROMPT_RESULT_OK); break;
        }

        struct PromptState
        {
            HostPromptResult chosen                        = HOST_PROMPT_RESULT_NONE;
            bool completed                                 = false;
            RedSalamander::Ui::AlertOverlayWindow* overlay = nullptr;
        };

        PromptState state{};
        state.chosen = escapeResult;

        auto onPromptButton = [](void* context, uint32_t buttonId) noexcept
        {
            auto* self = static_cast<PromptState*>(context);
            if (! self)
            {
                return;
            }

            self->chosen    = static_cast<HostPromptResult>(buttonId);
            self->completed = true;

            if (self->overlay)
            {
                self->overlay->Hide();
            }
        };

        RedSalamander::Ui::AlertOverlayWindow overlayWindow{};
        state.overlay = &overlayWindow;

        RedSalamander::Ui::AlertOverlayWindowCallbacks callbacks{};
        callbacks.context  = &state;
        callbacks.onButton = onPromptButton;
        overlayWindow.SetCallbacks(callbacks);
        overlayWindow.SetKeyBindings(static_cast<uint32_t>(primaryResult), static_cast<uint32_t>(escapeResult));

        const RedSalamander::Ui::AlertTheme theme = BuildHostAlertTheme();

        HRESULT hrShow = S_OK;
        if (request.scope == HOST_ALERT_SCOPE_APPLICATION)
        {
            HWND hostWindow           = nullptr;
            const HRESULT hrHostReady = EnsureHostUiThreadReady(hostWindow);
            if (FAILED(hrHostReady))
            {
                return hrHostReady;
            }

            hrShow = overlayWindow.ShowForParentClient(hostWindow, theme, std::move(model), true);
        }
        else
        {
            hrShow = overlayWindow.ShowForAnchor(targetWindow, theme, std::move(model), true);
        }

        if (FAILED(hrShow))
        {
            return hrShow;
        }

        while (! state.completed && overlayWindow.IsVisible())
        {
            MSG msg{};
            bool sawMessage = false;
            while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
            {
                sawMessage = true;
                if (msg.message == WM_QUIT)
                {
                    PostQuitMessage(static_cast<int>(msg.wParam));
                    state.completed = true;
                    break;
                }

                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }

            if (state.completed)
            {
                break;
            }

            if (! overlayWindow.IsVisible())
            {
                break;
            }

            if (! sawMessage)
            {
                WaitMessage();
            }
        }

        *result = state.chosen;
        return S_OK;
    }

    std::atomic<ULONG> _refCount;
    std::unique_ptr<RedSalamander::Ui::AlertOverlayWindow> _applicationOverlay;
    std::unordered_map<HWND, std::unique_ptr<RedSalamander::Ui::AlertOverlayWindow>> _windowOverlays;
    std::unordered_map<std::wstring, SessionSecretEntry> _sessionPasswordByConnectionId;
    std::unordered_map<std::wstring, SessionSecretEntry> _sessionPassphraseByConnectionId;
    std::unordered_map<std::wstring, SessionSecretEntry> _sessionRefreshTokenByConnectionId;
};

HRESULT
HostServices::PromptForConnectionSecretOnUiThread(const wchar_t* connectionName, HostConnectionSecretKind kind, HWND ownerWindow, wchar_t** secretOut) noexcept
{
    if (! secretOut)
    {
        return E_POINTER;
    }
    *secretOut = nullptr;

    HWND hostWindow           = nullptr;
    const HRESULT hrHostReady = EnsureHostUiThreadReady(hostWindow);
    if (FAILED(hrHostReady))
    {
        return hrHostReady;
    }

    Common::Settings::ConnectionProfile* profile = FindConnectionProfileMutable(connectionName);
    if (! profile)
    {
        Debug::Error(L"PromptForConnectionSecret failed: connection not found: '{}'", connectionName ? connectionName : L"(null)");
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    HWND owner = ownerWindow;
    if (! owner || ! IsWindow(owner))
    {
        owner = hostWindow;
    }

    const AppTheme& theme = HostFolderWindow().GetTheme();

    if (kind == HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool passphrase = (kind == HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE);
    const UINT captionId  = passphrase ? IDS_CONNECTIONS_PROMPT_PASSPHRASE_CAPTION : IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION;
    const UINT messageId  = passphrase ? IDS_CONNECTIONS_PROMPT_PASSPHRASE_MESSAGE_FMT : IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT;
    const UINT labelId    = passphrase ? IDS_CONNECTIONS_LABEL_PASSPHRASE : IDS_CONNECTIONS_LABEL_PASSWORD;

    const std::wstring caption = LoadStringResource(nullptr, captionId);
    std::wstring quickConnectLabel;
    std::wstring_view displayName = profile->name.empty() ? std::wstring_view(L"(unnamed)") : std::wstring_view(profile->name);
    if (RedSalamander::Connections::IsQuickConnectConnectionId(profile->id))
    {
        quickConnectLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
        if (quickConnectLabel.empty())
        {
            quickConnectLabel = L"<Quick Connect>";
        }
        displayName = quickConnectLabel;
    }

    std::wstring message           = FormatStringResource(nullptr, messageId, displayName);
    const std::wstring secretLabel = LoadStringResource(nullptr, labelId);

    if (const std::wstring url = ConnectionProfileUtils::BuildConnectionDisplayUrl(*profile); ! url.empty())
    {
        message = std::format(L"{}\n{}", message, url);
    }

    const bool promptForUserAndPassword = ! passphrase && profile->authMode == Common::Settings::ConnectionAuthMode::Password && profile->userName.empty();

    std::wstring userName;
    std::wstring secret;
    auto clearSecret       = wil::scope_exit([&] { SecureClear(secret); });
    const HRESULT promptHr = promptForUserAndPassword ? ::PromptForConnectionUserAndPassword(owner, theme, caption, message, {}, userName, secret)
                                                      : ::PromptForConnectionSecret(owner, theme, caption, message, secretLabel, passphrase, secret);
    if (FAILED(promptHr) || promptHr == S_FALSE)
    {
        return promptHr;
    }

    if (promptForUserAndPassword && ! userName.empty())
    {
        profile->userName = userName;

        if (RedSalamander::Connections::IsQuickConnectConnectionId(profile->id))
        {
            RedSalamander::Connections::SetQuickConnectProfile(*profile);
        }
        else
        {
            const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(L"RedSalamander", HostSettings());
            if (FAILED(saveHr))
            {
                const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(L"RedSalamander");
                Debug::Warning(
                    L"PromptForConnectionSecret: SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
            }
        }
    }

    if (! profile->id.empty())
    {
        auto& map                 = passphrase ? _sessionPassphraseByConnectionId : _sessionPasswordByConnectionId;
        SessionSecretEntry& entry = map[profile->id];
        if (entry.present)
        {
            SecureClear(entry.secret);
        }
        entry.present = true;
        entry.secret.assign(secret);
        RedSalamander::Connections::NoteSecretAccessAuthorized(profile->id, ToLocalSecretKind(kind));
    }

    const size_t bytes = (secret.size() + 1u) * sizeof(wchar_t);
    wil::unique_cotaskmem_string mem(static_cast<wchar_t*>(CoTaskMemAlloc(bytes)));
    if (! mem)
    {
        return E_OUTOFMEMORY;
    }

    if (! secret.empty())
    {
        memcpy(mem.get(), secret.data(), secret.size() * sizeof(wchar_t));
    }
    mem.get()[secret.size()] = L'\0';
    *secretOut               = mem.release();
    return S_OK;
}

HRESULT HostServices::SetConnectionSecretOnUiThread(const wchar_t* connectionName, HostConnectionSecretKind kind, const wchar_t* secret, BOOL persist) noexcept
{
    const HRESULT hrHostReady = EnsureHostUiThreadReady();
    if (FAILED(hrHostReady))
    {
        return hrHostReady;
    }

    if (! secret)
    {
        return E_INVALIDARG;
    }

    const Common::Settings::ConnectionProfile* profile = FindConnectionProfile(connectionName);
    if (! profile)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (profile->id.empty())
    {
        return E_INVALIDARG;
    }

    auto* sessionMap = &_sessionPasswordByConnectionId;
    if (kind == HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE)
    {
        sessionMap = &_sessionPassphraseByConnectionId;
    }
    else if (kind == HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN)
    {
        sessionMap = &_sessionRefreshTokenByConnectionId;
    }

    const std::wstring_view secretView(secret);
    const RedSalamander::Connections::SecretKind secretKind = ToLocalSecretKind(kind);
    if (RedSalamander::Connections::IsQuickConnectConnectionId(profile->id))
    {
        if (persist != FALSE || ! secretView.empty())
        {
            RedSalamander::Connections::SetQuickConnectSecret(secretKind, secretView);
        }
        else
        {
            RedSalamander::Connections::ClearQuickConnectSecret(secretKind);
        }

        SessionSecretEntry& entry = (*sessionMap)[profile->id];
        if (entry.present)
        {
            SecureClear(entry.secret);
        }
        entry.present = false;
        if (secretView.empty())
        {
            RedSalamander::Connections::ClearSecretAccessAuthorization(profile->id, secretKind);
        }
        else
        {
            entry.present = true;
            entry.secret.assign(secretView);
            RedSalamander::Connections::NoteSecretAccessAuthorized(profile->id, secretKind);
        }
        return S_OK;
    }

    SessionSecretEntry& entry = (*sessionMap)[profile->id];
    if (entry.present)
    {
        SecureClear(entry.secret);
    }
    entry.present = false;
    if (secretView.empty())
    {
        RedSalamander::Connections::ClearSecretAccessAuthorization(profile->id, secretKind);
    }
    else
    {
        entry.present = true;
        entry.secret.assign(secretView);
        RedSalamander::Connections::NoteSecretAccessAuthorized(profile->id, secretKind);
    }

    if (persist != FALSE)
    {
        const std::wstring targetName = RedSalamander::Connections::BuildCredentialTargetName(profile->id, secretKind);
        if (targetName.empty())
        {
            Debug::Error(L"HostServices: failed to build the credential target for connection '{}'.", profile->id);
            return E_FAIL;
        }

        if (secretView.empty())
        {
            const HRESULT deleteHr = RedSalamander::Connections::DeleteGenericCredential(targetName);
            if (FAILED(deleteHr) && deleteHr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
            {
                Debug::Error(L"HostServices: failed to delete the persisted credential for connection '{}' (hr=0x{:08X}).",
                             profile->id,
                             static_cast<unsigned long>(deleteHr));
                return deleteHr;
            }
        }
        else
        {
            const HRESULT saveHr = RedSalamander::Connections::SaveGenericCredential(targetName, profile->userName, secretView);
            if (FAILED(saveHr))
            {
                Debug::Error(L"HostServices: failed to save the persisted credential for connection '{}' (hr=0x{:08X}).",
                             profile->id,
                             static_cast<unsigned long>(saveHr));
                return saveHr;
            }
        }
    }

    return S_OK;
}

HRESULT HostServices::DeleteConnectionSecretOnUiThread(const wchar_t* connectionName, HostConnectionSecretKind kind, BOOL deletePersisted) noexcept
{
    const HRESULT hrHostReady = EnsureHostUiThreadReady();
    if (FAILED(hrHostReady))
    {
        return hrHostReady;
    }

    if (deletePersisted == FALSE)
    {
        const HRESULT clearHr = ClearCachedConnectionSecretOnUiThread(connectionName, kind);
        return clearHr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND) ? S_OK : clearHr;
    }

    const Common::Settings::ConnectionProfile* profile = FindConnectionProfile(connectionName);
    if (! profile)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (profile->id.empty())
    {
        return E_INVALIDARG;
    }

    const RedSalamander::Connections::SecretKind secretKind = ToLocalSecretKind(kind);
    if (RedSalamander::Connections::IsQuickConnectConnectionId(profile->id))
    {
        RedSalamander::Connections::ClearQuickConnectSecret(secretKind);
        const HRESULT clearHr = ClearCachedConnectionSecretOnUiThread(connectionName, kind);
        return clearHr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND) ? S_OK : clearHr;
    }

    return SetConnectionSecretOnUiThread(connectionName, kind, L"", TRUE);
}

HRESULT HostServices::ClearCachedConnectionSecretOnUiThread(const wchar_t* connectionName, HostConnectionSecretKind kind) noexcept
{
    const HRESULT hrHostReady = EnsureHostUiThreadReady();
    if (FAILED(hrHostReady))
    {
        return hrHostReady;
    }

    const Common::Settings::ConnectionProfile* profile = FindConnectionProfile(connectionName);
    if (! profile)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (profile->id.empty())
    {
        return S_OK;
    }

    auto* sessionMap = &_sessionPasswordByConnectionId;
    if (kind == HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE)
    {
        sessionMap = &_sessionPassphraseByConnectionId;
    }
    else if (kind == HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN)
    {
        sessionMap = &_sessionRefreshTokenByConnectionId;
    }
    auto& map = *sessionMap;

    const auto it = map.find(profile->id);
    if (it != map.end())
    {
        SecureClear(it->second.secret);
        map.erase(it);
    }

    RedSalamander::Connections::ClearSecretAccessAuthorization(profile->id, ToLocalSecretKind(kind));

    if (RedSalamander::Connections::IsQuickConnectConnectionId(profile->id))
    {
        RedSalamander::Connections::ClearQuickConnectSecret(ToLocalSecretKind(kind));
    }

    return S_OK;
}

HRESULT HostServices::UpgradeFtpAnonymousToPasswordOnUiThread(const wchar_t* connectionName, HWND ownerWindow) noexcept
{
    HWND hostWindow           = nullptr;
    const HRESULT hrHostReady = EnsureHostUiThreadReady(hostWindow);
    if (FAILED(hrHostReady))
    {
        return hrHostReady;
    }

    Common::Settings::ConnectionProfile* profile = FindConnectionProfileMutable(connectionName);
    if (! profile)
    {
        Debug::Error(L"UpgradeFtpAnonymousToPassword failed: connection not found: '{}'", connectionName ? connectionName : L"(null)");
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (profile->pluginId != L"builtin/file-system-ftp")
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (profile->authMode != Common::Settings::ConnectionAuthMode::Anonymous)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }

    HWND owner = ownerWindow;
    if (! owner || ! IsWindow(owner))
    {
        owner = hostWindow;
    }

    const AppTheme& theme = HostFolderWindow().GetTheme();

    const std::wstring caption = LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_FTP_CREDENTIALS_CAPTION);
    std::wstring message =
        FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_FTP_ANON_REJECTED_MESSAGE_FMT, profile->name.empty() ? L"(unnamed)" : profile->name);

    if (const std::wstring url = ConnectionProfileUtils::BuildConnectionDisplayUrl(*profile); ! url.empty())
    {
        message = std::format(L"{}\n{}", message, url);
    }

    std::wstring userName;
    std::wstring password;

    const std::wstring initialUser = (! profile->userName.empty() && profile->userName != L"anonymous") ? profile->userName : L"";
    const HRESULT promptHr         = ::PromptForConnectionUserAndPassword(owner, theme, caption, message, initialUser, userName, password);
    if (FAILED(promptHr) || promptHr == S_FALSE)
    {
        SecureClear(password);
        return promptHr;
    }

    profile->authMode = Common::Settings::ConnectionAuthMode::Password;
    profile->userName = userName;

    const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(L"RedSalamander", HostSettings());
    if (FAILED(saveHr))
    {
        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(L"RedSalamander");
        Debug::Error(L"UpgradeFtpAnonymousToPassword: SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
        SecureClear(password);
        return saveHr;
    }

    if (! profile->id.empty())
    {
        SessionSecretEntry& entry = _sessionPasswordByConnectionId[profile->id];
        if (entry.present)
        {
            SecureClear(entry.secret);
        }
        entry.present = true;
        entry.secret.assign(password);
        RedSalamander::Connections::NoteSecretAccessAuthorized(profile->id, RedSalamander::Connections::SecretKind::Password);
    }

    SecureClear(password);
    return S_OK;
}

HostServices& GetHostServicesImpl() noexcept
{
    static HostServices instance;
    return instance;
}

void ConfigureHostServices(FolderWindow& folderWindow, std::atomic<HWND>& folderWindowHwnd, Common::Settings::Settings& settings) noexcept
{
    g_hostFolderWindow.store(&folderWindow, std::memory_order_release);
    g_hostFolderWindowHwnd.store(&folderWindowHwnd, std::memory_order_release);
    g_hostSettings.store(&settings, std::memory_order_release);
}

IHost* GetHostServices() noexcept
{
    return &GetHostServicesImpl();
}

void HostClearConnectionSessionState() noexcept
{
    GetHostServicesImpl().ClearConnectionSessionState();
}

HRESULT HostShowAlert(const HostAlertRequest& request, void* cookie) noexcept
{
    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hrQuery = GetHostServices()->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (FAILED(hrQuery) || ! alerts)
    {
        return FAILED(hrQuery) ? hrQuery : E_NOINTERFACE;
    }

    return alerts->ShowAlert(&request, cookie);
}

HRESULT HostClearAlert(HostAlertScope scope, void* cookie) noexcept
{
    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hrQuery = GetHostServices()->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (FAILED(hrQuery) || ! alerts)
    {
        return FAILED(hrQuery) ? hrQuery : E_NOINTERFACE;
    }

    return alerts->ClearAlert(scope, cookie);
}

#ifdef ENABLE_TESTS
namespace
{
std::atomic<int> g_testPromptResultOverride{static_cast<int>(HOST_PROMPT_RESULT_NONE)};
std::atomic<uint64_t> g_testPromptRequestCount{0u};
} // namespace
#endif

HRESULT HostShowPrompt(const HostPromptRequest& request, void* cookie, HostPromptResult* result) noexcept
{
    if (! result)
    {
        return E_POINTER;
    }

#ifdef ENABLE_TESTS
    g_testPromptRequestCount.fetch_add(1u, std::memory_order_acq_rel);
    const auto testPromptResultOverride = static_cast<HostPromptResult>(g_testPromptResultOverride.load(std::memory_order_acquire));
    if (testPromptResultOverride != HOST_PROMPT_RESULT_NONE && PromptButtonsSupportResult(request.buttons, testPromptResultOverride))
    {
        *result = testPromptResultOverride;
        return S_OK;
    }

    if (HostGetAutoAcceptPrompts())
    {
        const HostPromptResult accept = DefaultPromptResultForButtons(request.buttons);
        *result                       = PromptButtonsSupportResult(request.buttons, accept) ? accept : request.defaultResult;
        if (*result == HOST_PROMPT_RESULT_NONE)
        {
            *result = accept;
        }
        return S_OK;
    }
#endif

    wil::com_ptr<IHostPrompts> prompts;
    const HRESULT hrQuery = GetHostServices()->QueryInterface(__uuidof(IHostPrompts), prompts.put_void());
    if (FAILED(hrQuery) || ! prompts)
    {
        return FAILED(hrQuery) ? hrQuery : E_NOINTERFACE;
    }

    return prompts->ShowPrompt(&request, cookie, result);
}

namespace
{
std::atomic<bool> g_autoAcceptPrompts{false};
} // namespace

void HostSetAutoAcceptPrompts(bool enabled) noexcept
{
    g_autoAcceptPrompts.store(enabled, std::memory_order_release);
}

bool HostGetAutoAcceptPrompts() noexcept
{
    return g_autoAcceptPrompts.load(std::memory_order_acquire);
}

#ifdef ENABLE_TESTS
void HostSetTestPromptResultOverride(HostPromptResult result) noexcept
{
    g_testPromptResultOverride.store(static_cast<int>(result), std::memory_order_release);
}

HostPromptResult HostGetTestPromptResultOverride() noexcept
{
    return static_cast<HostPromptResult>(g_testPromptResultOverride.load(std::memory_order_acquire));
}

void HostClearTestPromptResultOverride() noexcept
{
    g_testPromptResultOverride.store(static_cast<int>(HOST_PROMPT_RESULT_NONE), std::memory_order_release);
}

void HostResetTestPromptRequestCount() noexcept
{
    g_testPromptRequestCount.store(0u, std::memory_order_release);
}

uint64_t HostGetTestPromptRequestCount() noexcept
{
    return g_testPromptRequestCount.load(std::memory_order_acquire);
}
#endif

bool TryHandleHostServicesWindowMessage(UINT message, WPARAM wParam, LPARAM lParam, LRESULT& result) noexcept
{
    return GetHostServicesImpl().TryHandleMessage(message, wParam, lParam, result);
}
