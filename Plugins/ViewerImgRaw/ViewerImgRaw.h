#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "WindowMessages.h"

#include <d2d1.h>
#include <dwrite.h>

#include <atomic>
#include <cstdint>
#include <filesystem>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820) // WIL: deleted copy/move operators and padding
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"

struct ID2D1Bitmap;
struct ID2D1Factory;
struct ID2D1HwndRenderTarget;
struct ID2D1SolidColorBrush;
struct IDWriteFactory;
struct IDWriteTextFormat;

class ViewerImgRaw final : public IViewer, public IInformations
{
public:
    ViewerImgRaw();
    ~ViewerImgRaw();

    void SetHost(IHost* host) noexcept;

    ViewerImgRaw(const ViewerImgRaw&)            = delete;
    ViewerImgRaw(ViewerImgRaw&&)                 = delete;
    ViewerImgRaw& operator=(const ViewerImgRaw&) = delete;
    ViewerImgRaw& operator=(ViewerImgRaw&&)      = delete;

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
    enum class DisplayMode : uint8_t
    {
        Raw       = 0,
        Thumbnail = 1,
    };

    struct ExifInfo
    {
        std::wstring camera;
        std::wstring lens;
        std::wstring dateTime;
        float iso            = 0.0f;
        float shutterSeconds = 0.0f;
        float aperture       = 0.0f;
        float focalLengthMm  = 0.0f;
        uint16_t orientation = 1; // EXIF orientation (1..8)
        bool valid           = false;
    };

    struct Config
    {
        bool halfSize               = true;
        bool useCameraWb            = true;
        bool autoWb                 = false;
        bool preferThumbnail        = true;
        uint32_t zoomOnClickPercent = 50;
        uint32_t prevCache          = 1;
        uint32_t nextCache          = 1;

        // Export encoder options (WIC)
        uint32_t exportJpegQualityPercent  = 90;
        uint32_t exportJpegSubsampling     = 0; // WICJpegYCrCbSubsamplingOption
        uint32_t exportPngFilter           = 0; // WICPngFilterOption
        bool exportPngInterlace            = false;
        uint32_t exportTiffCompression     = 0; // WICTiffCompressionOption
        bool exportBmpUseV5Header32bppBGRA = true;
        bool exportGifInterlace            = false;
        uint32_t exportWmpQualityPercent   = 90;
        bool exportWmpLossless             = false;

        bool operator==(const Config&) const = default;
    };

    struct CachedImage final
    {
        enum class ThumbSource : uint8_t
        {
            None        = 0,
            Embedded    = 1,
            SidecarJpeg = 2,
        };

        uint32_t rawWidth       = 0;
        uint32_t rawHeight      = 0;
        uint16_t rawOrientation = 1; // orientation to apply when displaying RAW frame (1..8)
        std::vector<uint8_t> rawBgra;

        uint32_t thumbWidth       = 0;
        uint32_t thumbHeight      = 0;
        uint16_t thumbOrientation = 1; // orientation to apply when displaying thumbnail frame (1..8)
        std::vector<uint8_t> thumbBgra;

        bool thumbAvailable     = false;
        bool thumbDecoded       = false;
        ThumbSource thumbSource = ThumbSource::None;

        ExifInfo exif;
    };

    struct OtherItem final
    {
        std::wstring primaryPath;
        std::wstring sidecarJpegPath;
        std::wstring label;
        bool isRaw = false;
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

    struct AsyncOpenResult
    {
        ViewerImgRaw* viewer = nullptr;
        uint64_t requestId   = 0;
        HRESULT hr           = E_FAIL;
        std::wstring path;
        bool updateOtherFiles                = false;
        uint32_t configSignature             = 0;
        DisplayMode frameMode                = DisplayMode::Raw;
        bool isFinal                         = true;
        bool thumbAvailable                  = false;
        CachedImage::ThumbSource thumbSource = CachedImage::ThumbSource::None;

        uint32_t width  = 0;
        uint32_t height = 0;
        std::vector<uint8_t> bgra;

        std::wstring statusMessage;
        ExifInfo exif;
    };

    struct AsyncExportResult
    {
        ViewerImgRaw* viewer = nullptr;
        HRESULT hr           = E_FAIL;
        std::wstring outputPath;
        std::wstring statusMessage;
    };

    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kClassName[] = L"RedSalamander.ViewerImgRaw";

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnCreate(HWND hwnd);
    void OnDestroy();
    void OnTimer(UINT_PTR timerId) noexcept;
    void OnSize(UINT width, UINT height);
    void OnPaint();
    LRESULT OnEraseBkgnd(HWND hwnd, HDC hdc) noexcept;
    void OnCommand(HWND hwnd, UINT id, UINT code, HWND control);
    void OnKeyDown(HWND hwnd, UINT vk) noexcept;
    LRESULT OnMouseWheel(HWND hwnd, short delta, UINT keyState, int x, int y) noexcept;
    void OnHScroll(HWND hwnd, UINT code) noexcept;
    void OnVScroll(HWND hwnd, UINT code) noexcept;
    void OnLButtonDown(HWND hwnd, int x, int y) noexcept;
    void OnLButtonDblClick(HWND hwnd, int x, int y) noexcept;
    void OnLButtonUp(HWND hwnd) noexcept;
    void OnMouseMove(HWND hwnd, int x, int y) noexcept;
    void OnCaptureChanged() noexcept;
    void OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggested) noexcept;
    void OnAsyncOpenComplete(std::unique_ptr<AsyncOpenResult> result) noexcept;
    void OnAsyncProgress(int stage, int percent) noexcept;
    void OnAsyncExportComplete(std::unique_ptr<AsyncExportResult> result) noexcept;
    void OnNcActivate(bool windowActive) noexcept;
    LRESULT OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept;
    LRESULT OnInputLangChange(HWND hwnd, WPARAM wp, LPARAM lp) noexcept;

    void Layout(HWND hwnd) noexcept;
    void ComputeLayoutRects(HWND hwnd) noexcept;
    void UpdateScrollBars(HWND hwnd) noexcept;
    void ApplyTheme(HWND hwnd) noexcept;
    void ApplyTitleBarTheme(bool windowActive) noexcept;
    void UpdateMenuChecks(HWND hwnd) noexcept;
    void ApplyMenuTheme(HWND hwnd) noexcept;
    void UpdateMenuShortcutTextForKeyboardLayout() noexcept;
    void PrepareMenuTheme(HMENU menu, bool topLevel, std::vector<MenuItemData>& outItems) noexcept;
    void OnMeasureMenuItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept;
    void OnDrawMenuItem(DRAWITEMSTRUCT* draw) noexcept;
    LRESULT OnMeasureItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept;
    LRESULT OnDrawItem(HWND hwnd, DRAWITEMSTRUCT* draw) noexcept;

    void RefreshFileCombo(HWND hwnd) noexcept;
    void SyncFileComboSelection() noexcept;
    void BeginExport(HWND hwnd);
    void StartAsyncOpen(HWND hwnd, std::wstring_view path, bool updateOtherFiles) noexcept;

    void BeginLoadingUi() noexcept;
    void EndLoadingUi() noexcept;
    void UpdateLoadingSpinner() noexcept;
    void DrawLoadingOverlay(ID2D1HwndRenderTarget* target, ID2D1SolidColorBrush* brush) noexcept;
    void DrawExifOverlay(ID2D1HwndRenderTarget* target, ID2D1SolidColorBrush* brush) noexcept;

    bool EnsureDirect2D(HWND hwnd) noexcept;
    void DiscardDirect2D() noexcept;
    bool EnsureImageBitmap() noexcept;
    bool ComputeImageLayoutPx(float& outZoom, float& outX, float& outY, float& outDrawW, float& outDrawH) noexcept;
    void ApplyZoom(HWND hwnd, float newZoom, std::optional<POINT> anchorClientPt) noexcept;
    void ClearImageCache() noexcept;
    void UpdateNeighborCache(uint64_t requestId) noexcept;
    void StartPrefetchNeighbors(uint64_t requestId) noexcept;
    bool TryUseCachedImage(HWND hwnd, const std::wstring& path, bool& outContinueDecoding) noexcept;
    bool HasDisplayImage() const noexcept;
    bool IsDisplayingThumbnail() const noexcept;
    void SetDisplayMode(DisplayMode mode) noexcept;
    void UpdateOrientationState() noexcept;
    void RebuildExifOverlayText() noexcept;
    std::wstring BuildStatusBarText(bool drewImage, float displayedZoom) const;
    std::wstring BuildStatusBarRightText(bool drewImage, float displayedZoom) const;

    std::atomic_uint32_t _refCount{1};

    // Host services
    wil::com_ptr<IHostAlerts> _hostAlerts;

    // Plugin metadata / configuration
    PluginMetaData _metaData{};
    std::wstring _metaId;
    std::wstring _metaShortId;
    std::wstring _metaName;
    std::wstring _metaDescription;

    Config _config;
    std::string _configJson;

    // Viewer state
    wil::unique_hwnd _hWnd;
    wil::unique_hwnd _hFileCombo;
    HWND _hFileComboList = nullptr;
    HWND _hFileComboItem = nullptr;
    wil::unique_hfont _uiFont;

    RECT _headerRect{};
    RECT _contentRect{};
    RECT _statusRect{};

    wil::unique_hbrush _menuHeaderBrush;
    std::vector<MenuItemData> _menuThemeItems;

    wil::com_ptr<IFileSystem> _fileSystem;
    std::wstring _fileSystemName;

    std::wstring _currentPath;
    std::wstring _currentSidecarJpegPath;
    std::wstring _currentLabel;
    std::vector<OtherItem> _otherItems;
    size_t _otherIndex     = 0;
    bool _syncingFileCombo = false;

    bool _hasTheme = false;
    ViewerTheme _theme{};

    COLORREF _uiBg       = RGB(255, 255, 255);
    COLORREF _uiText     = RGB(0, 0, 0);
    COLORREF _uiHeaderBg = RGB(240, 240, 240);
    COLORREF _uiStatusBg = RGB(240, 240, 240);

    bool _allowEraseBkgnd = true;

    // Loading/image state
    std::atomic_uint64_t _openRequestId{0};
    bool _isLoading = false;
    std::wstring _statusMessage;
    bool _alertVisible = false;

    bool _showLoadingOverlay            = false;
    float _loadingSpinnerAngleDeg       = 0.0f;
    ULONGLONG _loadingSpinnerLastTickMs = 0;
    wil::com_ptr<IDWriteTextFormat> _loadingOverlayFormat;
    wil::com_ptr<IDWriteTextFormat> _loadingOverlaySubFormat;
    wil::com_ptr<IDWriteTextFormat> _exifOverlayFormat;

    DisplayMode _displayMode   = DisplayMode::Raw;
    DisplayMode _displayedMode = DisplayMode::Raw;
    bool _showExifOverlay      = false;
    std::wstring _exifOverlayText;

    int _rawProgressPercent = -1;
    int _rawProgressStage   = -1;
    std::wstring _rawProgressStageText;

    std::mutex _cacheMutex;
    std::unordered_map<std::wstring, std::unique_ptr<CachedImage>> _imageCache;
    std::unordered_set<std::wstring> _inflightDecodes;
    std::unique_ptr<CachedImage> _currentImageOwned;
    CachedImage* _currentImage = nullptr;
    std::wstring _currentImageKey;

    bool _fitToWindow        = true;
    float _manualZoom        = 1.0f;
    float _panOffsetXPx      = 0.0f;
    float _panOffsetYPx      = 0.0f;
    bool _panning            = false;
    bool _hScrollVisible     = false;
    bool _vScrollVisible     = false;
    bool _updatingScrollBars = false;
    POINT _panStartPoint{};
    float _panStartOffsetXPx          = 0.0f;
    float _panStartOffsetYPx          = 0.0f;
    bool _transientZoomActive         = false;
    bool _transientSavedFitToWindow   = true;
    float _transientSavedManualZoom   = 1.0f;
    float _transientSavedPanOffsetXPx = 0.0f;
    float _transientSavedPanOffsetYPx = 0.0f;

    uint16_t _baseOrientation     = 1;
    uint16_t _userOrientation     = 1;
    uint16_t _viewOrientation     = 1;
    bool _orientationUserModified = false;

    float _brightness = 0.0f; // -1..1
    float _contrast   = 1.0f; // 0.1..3
    float _gamma      = 1.0f; // 0.1..5
    bool _grayscale   = false;
    bool _negative    = false;

    std::vector<uint8_t> _adjustedBgra;

    // Direct2D resources
    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _d2dTarget;
    wil::com_ptr<ID2D1SolidColorBrush> _solidBrush;
    wil::com_ptr<IDWriteTextFormat> _uiTextFormat;
    wil::com_ptr<IDWriteTextFormat> _uiTextFormatRight;
    wil::com_ptr<ID2D1Bitmap> _imageBitmap;

    // Callback (weak)
    IViewerCallback* _callback = nullptr;
    void* _callbackCookie      = nullptr;
};

inline constexpr UINT kAsyncOpenCompleteMessage   = WndMsg::kViewerImgRawAsyncOpenComplete;
inline constexpr UINT kAsyncProgressMessage       = WndMsg::kViewerImgRawAsyncProgress;
inline constexpr UINT kAsyncExportCompleteMessage = WndMsg::kViewerImgRawAsyncExportComplete;

extern HINSTANCE g_hInstance;
