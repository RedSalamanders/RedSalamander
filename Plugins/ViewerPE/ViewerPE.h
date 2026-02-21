#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <cstdint>
#include <filesystem>
#include <memory>
#include <string>
#include <thread>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182) // WIL: deleted copy/move operators and padding
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"

struct ID2D1Factory;
struct ID2D1HwndRenderTarget;
struct ID2D1SolidColorBrush;
struct IDWriteFactory;
struct IDWriteTextFormat;
struct IDWriteTextLayout;

class ViewerPE final : public IViewer, public IInformations
{
public:
    ViewerPE();
    ~ViewerPE();

    void SetHost(IHost* host) noexcept;

    ViewerPE(const ViewerPE&)            = delete;
    ViewerPE(ViewerPE&&)                 = delete;
    ViewerPE& operator=(const ViewerPE&) = delete;
    ViewerPE& operator=(ViewerPE&&)      = delete;

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
    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kClassName[] = L"RedSalamander.ViewerPE";

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnCreate(HWND hwnd) noexcept;
    void OnSize(UINT width, UINT height) noexcept;
    void OnDpiChanged(UINT dpi, const RECT& suggestedRect) noexcept;
    void OnPaint(HWND hwnd) noexcept;
    void OnMouseWheel(short delta) noexcept;
    void OnVScroll(WORD request, WORD position) noexcept;
    void OnKeyDown(UINT vk) noexcept;
    void OnCommand(HWND hwnd, UINT commandId, UINT notifyCode, HWND control) noexcept;
    LRESULT OnMeasureItem(MEASUREITEMSTRUCT* measure) noexcept;
    LRESULT OnDrawItem(DRAWITEMSTRUCT* draw) noexcept;
    LRESULT OnCtlColor(UINT msg, HDC hdc, HWND control) noexcept;

    void ApplyTitleBarTheme(bool windowActive) noexcept;
    void ResetDeviceResources() noexcept;
    void EnsureDeviceResources(HWND hwnd) noexcept;
    void EnsureTextLayout(float viewportWidthDip, float viewportHeightDip) noexcept;
    void UpdateScrollBars(HWND hwnd, float viewportHeightDip) noexcept;

    void Layout(HWND hwnd) noexcept;
    void RefreshFileCombo(HWND hwnd) noexcept;
    void SyncFileComboSelection() noexcept;
    void UpdateMenuState(HWND hwnd) noexcept;

    void SetScrollDip(HWND hwnd, float scrollDip) noexcept;
    void ScrollByDip(HWND hwnd, float deltaDip) noexcept;

    void CommandExit() noexcept;
    void CommandRefresh(HWND hwnd) noexcept;
    void CommandOtherNext(HWND hwnd) noexcept;
    void CommandOtherPrevious(HWND hwnd) noexcept;
    void CommandOtherFirst(HWND hwnd) noexcept;
    void CommandOtherLast(HWND hwnd) noexcept;
    void CommandExportText(HWND hwnd) noexcept;
    void CommandExportMarkdown(HWND hwnd) noexcept;

    struct AsyncParseResult
    {
        uint64_t requestId = 0;
        HRESULT hr         = E_FAIL;
        std::wstring title;
        std::wstring subtitle;
        std::wstring body;
        std::wstring markdown;
    };

    void StartAsyncParse(HWND hwnd, wil::com_ptr<IFileSystem> fileSystem, std::wstring path) noexcept;
    void OnAsyncParseComplete(std::unique_ptr<AsyncParseResult> result) noexcept;

private:
    std::atomic_ulong _refCount{1};

    PluginMetaData _metaData{};
    std::wstring _metaId;
    std::wstring _metaShortId;
    std::wstring _metaName;
    std::wstring _metaDescription;

    std::string _configurationJson;

    ViewerTheme _theme{};
    bool _hasTheme  = false;
    bool _isLoading = false;

    IViewerCallback* _callback = nullptr;
    void* _callbackCookie      = nullptr;

    wil::com_ptr<IHostAlerts> _hostAlerts;

    wil::com_ptr<IFileSystem> _fileSystem;
    std::wstring _currentPath;
    std::vector<std::wstring> _otherFiles;
    size_t _otherIndex     = 0;
    bool _syncingFileCombo = false;

    wil::unique_hwnd _hWnd;
    wil::unique_hwnd _hFileCombo;
    HWND _hFileComboList = nullptr;
    HWND _hFileComboItem = nullptr;
    UINT _dpi            = 96;

    wil::unique_hfont _uiFont;
    wil::unique_hbrush _headerBrush;

    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<IDWriteFactory> _writeFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _renderTarget;

    wil::com_ptr<ID2D1SolidColorBrush> _bgBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _cardBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _cardBorderBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _textBrush;

    wil::com_ptr<IDWriteTextFormat> _baseTextFormat;
    wil::com_ptr<IDWriteTextLayout> _textLayout;

    float _scrollDip         = 0.0f;
    float _contentHeightDip  = 0.0f;
    float _layoutWidthDip    = 0.0f;
    float _viewportHeightDip = 0.0f;
    float _headerHeightDip   = 0.0f;

    std::wstring _titleText;
    std::wstring _subtitleText;
    std::wstring _bodyText;
    std::wstring _markdownText;

    std::atomic_uint64_t _parseRequestId{0};
    std::jthread _worker;
};
