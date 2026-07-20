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
#pragma warning(disable : 4625 4626 5026 5027 4820 28182) // WIL: deleted copy/move operators and padding
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <WebView2.h>

#include "DxUi/DxUi.h"
#include "DxUi/DxUiNativeMenuInterop.h"
#include "EmbeddedViewerBase.h"
#include "Helpers.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"
#include "ViewerWebSecurity.h"

enum class ViewerWebKind : uint8_t
{
    Web,
    Json,
    Markdown,
};

[[nodiscard]] const char* GetViewerWebStaticConfigurationSchema(ViewerWebKind kind) noexcept;

class ViewerWeb final : public EmbeddedViewerBase<ViewerWeb>, public IInformations
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

    void CancelPendingWebView2Initialization() noexcept;

    LRESULT HandleFileComboHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept;
    void FocusMainSurfaceFromFileCombo(HWND hwnd) noexcept;

private:
    enum class JsonViewMode : uint8_t
    {
        Pretty,
        Tree,
        JsonLines,
    };

    struct ViewerWebConfig
    {
        uint32_t maxDocumentMiB      = ViewerWebSecurity::kDefaultMaxDocumentMiB;
        bool allowExternalNavigation = ViewerWebSecurity::kDefaultAllowExternalNavigation;
        bool devToolsEnabled         = false;
        JsonViewMode jsonViewMode    = JsonViewMode::Pretty;
    };

    struct AsyncLoadResult
    {
        AsyncLoadResult()                                  = default;
        AsyncLoadResult(const AsyncLoadResult&)            = delete;
        AsyncLoadResult(AsyncLoadResult&&)                 = delete;
        AsyncLoadResult& operator=(const AsyncLoadResult&) = delete;
        AsyncLoadResult& operator=(AsyncLoadResult&&)      = delete;

        ViewerWeb* viewer  = nullptr;
        HWND hwnd          = nullptr;
        uint64_t requestId = 0;
        HRESULT hr         = E_FAIL;
        std::wstring path;
        std::wstring title;
        std::string utf8;
        std::wstring statusMessage;
        std::optional<std::filesystem::path> extractedWin32Path;
        ViewerWebSecurity::DocumentRoute documentRoute = ViewerWebSecurity::DocumentRoute::None;
        uint64_t loadedSourceBytes                     = 0u;
        uint64_t generatedOutputBytes                  = 0u;
        uint64_t generatedOutputLimit                  = 0u;
        bool generatedOutputRejected                   = false;
        bool jsonExpandCollapseAvailable               = false;
        bool offerTextViewerFallback                   = false;

        // Snapshot of UI-thread state captured in StartAsyncLoad so the worker never reads live members.
        ViewerWebKind kindSnapshot{};
        ViewerWebConfig configSnapshot{};
        bool hasThemeSnapshot = false;
        ViewerTheme themeSnapshot{};
        bool markdownShowSourceSnapshot = false;
        wil::com_ptr<IFileSystem> fileSystemSnapshot;
        std::wstring metaIdSnapshot;
        std::wstring metaNameSnapshot;
    };

    struct AsyncSaveWorkItem
    {
        AsyncSaveWorkItem()                                    = default;
        AsyncSaveWorkItem(const AsyncSaveWorkItem&)            = delete;
        AsyncSaveWorkItem(AsyncSaveWorkItem&&)                 = delete;
        AsyncSaveWorkItem& operator=(const AsyncSaveWorkItem&) = delete;
        AsyncSaveWorkItem& operator=(AsyncSaveWorkItem&&)      = delete;

        ViewerWeb* viewer      = nullptr;
        HWND hwnd              = nullptr;
        uint64_t requestId     = 0u;
        uint64_t maxBytes      = 0u;
        uint32_t testFaultMask = 0u;
        HRESULT hr             = E_FAIL;
        std::wstring sourcePath;
        std::wstring destinationPath;
        std::wstring statusMessage;
        wil::com_ptr<IFileSystem> fileSystemSnapshot;
        wil::unique_hmodule moduleKeepAlive;
    };

    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kClassName[] = L"RedSalamander.ViewerWeb";

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnCreate(HWND hwnd);
    void OnDestroy() noexcept;
    void OnSize(UINT width, UINT height) noexcept;
    void OnCommand(HWND hwnd, UINT commandId, UINT code, HWND control) noexcept;
    void OnContextMenu(HWND hwnd, POINT screenPt) noexcept;
    void OnKeyDown(HWND hwnd, UINT vk) noexcept;
    void OnPaint(HWND hwnd) noexcept;
    LRESULT OnEraseBkgnd(HWND hwnd, HDC hdc) noexcept;
    void OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggested) noexcept;
    LRESULT OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept;
    void OnFindMessage(const FINDREPLACEW* findReplace) noexcept;
    void OnAsyncLoadComplete(std::unique_ptr<AsyncLoadResult> result) noexcept;
    void OnAsyncSaveComplete(std::unique_ptr<AsyncSaveWorkItem> result) noexcept;
    void OnAsyncPostFailure(UINT operationKind) noexcept;

    void Layout(HWND hwnd) noexcept;
    void ComputeLayoutRects(HWND hwnd) noexcept;

    void ApplyTheme(HWND hwnd) noexcept;
    void ApplyTitleBarTheme(bool windowActive) noexcept;
    void ApplyMenuTheme(HWND hwnd) noexcept;
    void UpdateMenuState(HWND hwnd, bool syncDxMenuBar = true) noexcept;

    HRESULT EnsureWebView2(HWND hwnd) noexcept;
    HRESULT CreateControllerFromEnvironment(HWND hwnd, ICoreWebView2Environment* environment, uint64_t sharedEnvironmentGeneration) noexcept;
    void DiscardWebView2() noexcept;
    [[nodiscard]] HRESULT ConfigureWebViewSettings() noexcept;
    HRESULT HandleNavigationStarting(ICoreWebView2NavigationStartingEventArgs* args, ViewerWebSecurity::NavigationSurface surface) noexcept;
    HRESULT HandleNewWindowRequested(ICoreWebView2NewWindowRequestedEventArgs* args) noexcept;
    void ApplyWebViewThemeScript() noexcept;
    HRESULT NavigatePendingContent(HWND hwnd) noexcept;

    HRESULT OpenPath(HWND hwnd, const std::wstring& path, bool updateOtherFiles) noexcept;
    void RefreshFileCombo(HWND hwnd) noexcept;

    HRESULT StartAsyncLoad(HWND hwnd, const std::wstring& path) noexcept;
    static void AsyncLoadProc(AsyncLoadResult* payload) noexcept;
    static void AsyncSaveProc(AsyncSaveWorkItem* payload) noexcept;

    HRESULT CommandSaveAs(HWND hwnd) noexcept;
    HRESULT StartAsyncSave(HWND hwnd, const std::filesystem::path& destination, uint32_t testFaultMask = 0u) noexcept;
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
    bool OfferTextViewerFallbackPrompt() noexcept;
    HRESULT OpenCurrentDocumentInTextViewer() noexcept;

    void ShowHostAlert(HWND targetWindow, HostAlertSeverity severity, const std::wstring& message) noexcept;

private:
    ViewerWebKind _kind = ViewerWebKind::Web;

    std::atomic<ULONG> _refCount{1};
    IHost* _host = nullptr;
    wil::com_ptr<IHostAlerts> _hostAlerts;

    PluginMetaData _metaData{};
    std::wstring _metaId;
    std::wstring _metaShortId;
    std::wstring _metaName;
    std::wstring _metaDescription;

    ViewerWebConfig _config{};
    std::string _configurationJson;

    wil::com_ptr<IFileSystem> _fileSystem;
    std::wstring _fileSystemName;
    std::vector<std::wstring> _otherFiles;
    size_t _otherIndex = 0;
    std::wstring _currentPath;
    bool _syncingFileCombo = false;

    wil::unique_hmenu _menuHandle;
    RedSalamander::DxUi::NativeMenuBarHost _menuBarHost;
    wil::unique_hwnd _hFileComboHost;
    RedSalamander::DxUi::WindowHost _fileComboHost;
    RedSalamander::DxUi::ComboBox* _fileComboControl = nullptr;
    bool _fileComboHostPreExpandPopup                = false;

    RECT _headerRect{};
    RECT _contentRect{};

    wil::unique_hbrush _headerBrush;

    uint64_t _openRequestId = 0;
    std::atomic<uint64_t> _saveRequestId{0u};
    std::atomic<uint64_t> _asyncLoadPostFailureRequestId{0u};
    std::atomic<uint64_t> _asyncSavePostFailureRequestId{0u};
    std::atomic<HRESULT> _asyncSavePostFailureHr{E_FAIL};
    std::atomic<uint64_t> _asyncLoadPostFailureCount{0u};
    std::atomic<uint64_t> _asyncSavePostFailureCount{0u};
    std::wstring _statusMessage;
    bool _loadPostFailureTerminal = false;
    bool _saveInProgress          = false;

    std::optional<std::wstring> _pendingPath;
    std::optional<std::wstring> _pendingWebContent;
    std::optional<std::string> _pendingDocumentUtf8;
    std::optional<std::wstring> _internalDocumentUrl;
    std::optional<std::filesystem::path> _tempExtractedPath;
    std::wstring _allowedDocumentUrl;
    ViewerWebSecurity::DocumentRoute _documentRoute = ViewerWebSecurity::DocumentRoute::None;
    uint64_t _loadedSourceBytes                     = 0u;
    bool _documentScriptsEnabled                    = false;
    bool _navigationCompleted                       = false;
    bool _navigationSucceeded                       = false;
    bool _generatedOutputRejected                   = false;
    uint64_t _generatedOutputBytes                  = 0u;
    uint64_t _generatedOutputLimit                  = 0u;
    bool _markdownShowSource                        = false;
    bool _jsonExpandCollapseAvailable               = false;
    bool _webViewInitInProgress                     = false;

    wil::com_ptr<ICoreWebView2Controller> _webViewController;
    wil::com_ptr<ICoreWebView2> _webView;

    EventRegistrationToken _navStartingToken{};
    EventRegistrationToken _frameNavStartingToken{};
    EventRegistrationToken _newWindowRequestedToken{};
    EventRegistrationToken _navCompletedToken{};
    EventRegistrationToken _accelToken{};
    EventRegistrationToken _webResourceRequestedToken{};

    wil::unique_hwnd _hFindDialog;
    std::array<wchar_t, 256> _findBuffer{};
    FINDREPLACEW _findReplace{};
    std::wstring _findQuery;
};
