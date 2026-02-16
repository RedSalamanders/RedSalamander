#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <commdlg.h>

#include <array>
#include <atomic>
#include <cstdint>
#include <filesystem>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820) // WIL: deleted copy/move operators and padding
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <WebView2.h>

#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"

enum class ViewerWebKind : uint8_t
{
    Web,
    Json,
    Markdown,
};

class ViewerWeb final : public IViewer, public IInformations
{
public:
    explicit ViewerWeb(ViewerWebKind kind) noexcept;
    ~ViewerWeb();

    void SetHost(IHost* host) noexcept;

    ViewerWeb(const ViewerWeb&)            = delete;
    ViewerWeb(ViewerWeb&&)                 = delete;
    ViewerWeb& operator=(const ViewerWeb&) = delete;
    ViewerWeb& operator=(ViewerWeb&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override;

    HRESULT STDMETHODCALLTYPE Open(const ViewerOpenContext* context) noexcept override;
    HRESULT STDMETHODCALLTYPE Close() noexcept override;
    HRESULT STDMETHODCALLTYPE SetTheme(const ViewerTheme* theme) noexcept override;
    HRESULT STDMETHODCALLTYPE SetCallback(IViewerCallback* callback, void* cookie) noexcept override;

private:
    enum class JsonViewMode : uint8_t
    {
        Pretty,
        Tree,
    };

    struct ViewerWebConfig
    {
        uint32_t maxDocumentMiB      = 32;
        bool allowExternalNavigation = true;
        bool devToolsEnabled         = false;
        JsonViewMode jsonViewMode    = JsonViewMode::Pretty;
    };

    struct MenuItemData
    {
        UINT id = 0;
        std::wstring text;
        std::wstring shortcut;
        bool separator  = false;
        bool topLevel   = false;
        bool hasSubMenu = false;
    };

    struct AsyncLoadResult
    {
        ViewerWeb* viewer  = nullptr;
        HWND hwnd          = nullptr;
        uint64_t requestId = 0;
        HRESULT hr         = E_FAIL;
        std::wstring path;
        std::wstring title;
        std::string utf8;
        std::wstring statusMessage;
        std::optional<std::filesystem::path> extractedWin32Path;
    };

    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kClassName[] = L"RedSalamander.ViewerWeb";

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnCreate(HWND hwnd);
    void OnDestroy() noexcept;
    void OnSize(UINT width, UINT height) noexcept;
    void OnCommand(HWND hwnd, UINT commandId, UINT code, HWND control) noexcept;
    void OnKeyDown(HWND hwnd, UINT vk) noexcept;
    void OnPaint(HWND hwnd) noexcept;
    LRESULT OnEraseBkgnd(HWND hwnd, HDC hdc) noexcept;
    void OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggested) noexcept;
    LRESULT OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept;
    void OnFindMessage(const FINDREPLACEW* findReplace) noexcept;
    LRESULT OnMeasureItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept;
    LRESULT OnDrawItem(HWND hwnd, DRAWITEMSTRUCT* draw) noexcept;
    void OnAsyncLoadComplete(std::unique_ptr<AsyncLoadResult> result) noexcept;

    void Layout(HWND hwnd) noexcept;
    void ComputeLayoutRects(HWND hwnd) noexcept;

    void ApplyTheme(HWND hwnd) noexcept;
    void ApplyTitleBarTheme(bool windowActive) noexcept;
    void ApplyMenuTheme(HWND hwnd) noexcept;
    void PrepareMenuTheme(HMENU menu, bool topLevel, std::vector<MenuItemData>& outItems) noexcept;

    HRESULT EnsureWebView2(HWND hwnd) noexcept;
    void DiscardWebView2() noexcept;
    void UpdateWebViewTheme() noexcept;

    HRESULT OpenPath(HWND hwnd, const std::wstring& path, bool updateOtherFiles) noexcept;
    void RefreshFileCombo(HWND hwnd) noexcept;

    HRESULT StartAsyncLoad(HWND hwnd, const std::wstring& path) noexcept;
    static void AsyncLoadProc(AsyncLoadResult* payload) noexcept;

    HRESULT CommandSaveAs(HWND hwnd) noexcept;
    void CommandFind(HWND hwnd) noexcept;
    void CommandFindNext(HWND hwnd) noexcept;
    void CommandFindPrevious(HWND hwnd) noexcept;
    void CommandCopyUrl(HWND hwnd) noexcept;
    void CommandOpenExternal(HWND hwnd) noexcept;
    void CommandZoom(double factor) noexcept;
    void CommandZoomIn() noexcept;
    void CommandZoomOut() noexcept;
    void CommandZoomReset() noexcept;
    void CommandToggleDevTools() noexcept;
    void CommandJsonExpandAll() noexcept;
    void CommandJsonCollapseAll() noexcept;
    void CommandMarkdownToggleSource() noexcept;

    void ShowHostAlert(HWND targetWindow, HostAlertSeverity severity, const std::wstring& message) noexcept;

private:
    ViewerWebKind _kind = ViewerWebKind::Web;

    std::atomic<ULONG> _refCount{1};
    IHost* _host = nullptr;
    wil::com_ptr<IHostAlerts> _hostAlerts;

    IViewerCallback* _callback = nullptr;
    void* _callbackCookie      = nullptr;

    PluginMetaData _metaData{};
    std::wstring _metaId;
    std::wstring _metaShortId;
    std::wstring _metaName;
    std::wstring _metaDescription;

    ViewerWebConfig _config{};
    std::string _configurationJson;

    bool _hasTheme = false;
    ViewerTheme _theme{};

    wil::com_ptr<IFileSystem> _fileSystem;
    std::wstring _fileSystemName;
    std::vector<std::wstring> _otherFiles;
    size_t _otherIndex = 0;
    std::wstring _currentPath;

    wil::unique_hwnd _hWnd;
    wil::unique_hwnd _hFileCombo;
    HWND _hFileComboItem = nullptr;
    HWND _hFileComboList = nullptr;

    RECT _headerRect{};
    RECT _contentRect{};

    wil::unique_hfont _uiFont;
    wil::unique_hbrush _headerBrush;

    std::vector<MenuItemData> _menuThemeItems;

    uint64_t _openRequestId = 0;
    std::wstring _statusMessage;

    std::optional<std::wstring> _pendingPath;
    std::optional<std::wstring> _pendingWebContent;
    std::optional<std::string> _pendingDocumentUtf8;
    std::optional<std::filesystem::path> _tempExtractedPath;
    bool _markdownShowSource    = false;
    bool _webViewInitInProgress = false;

    wil::com_ptr<ICoreWebView2Environment> _webViewEnvironment;
    wil::com_ptr<ICoreWebView2Controller> _webViewController;
    wil::com_ptr<ICoreWebView2> _webView;

    EventRegistrationToken _navStartingToken{};
    EventRegistrationToken _navCompletedToken{};
    EventRegistrationToken _accelToken{};

    wil::unique_hwnd _hFindDialog;
    std::array<wchar_t, 256> _findBuffer{};
    FINDREPLACEW _findReplace{};
    std::wstring _findQuery;
};
