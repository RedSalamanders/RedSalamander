#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <cstdint>
#include <filesystem>
#include <memory>
#include <string>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182) // WIL: deleted copy/move operators and padding
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include "EmbeddedViewerBase.h"
#include "Helpers.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"

struct VlcState;
struct ID2D1Factory;
struct ID2D1HwndRenderTarget;
struct IDWriteFactory;
struct IDWriteTextFormat;

[[nodiscard]] const char* GetViewerVlcStaticConfigurationSchema() noexcept;

class ViewerVLC final : public EmbeddedViewerBase<ViewerVLC>, public IInformations
{
public:
    ViewerVLC();
    ~ViewerVLC();

    void SetHost(IHost* host) noexcept;

    ViewerVLC(const ViewerVLC&)            = delete;
    ViewerVLC(ViewerVLC&&)                 = delete;
    ViewerVLC& operator=(const ViewerVLC&) = delete;
    ViewerVLC& operator=(ViewerVLC&&)      = delete;

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

private:
    enum class HudPart : uint8_t;
    struct VlcLoadSpec;
    struct VlcAsyncLoadWork;
    struct VlcAsyncLoadResult;

    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kClassName[] = L"RedSalamander.ViewerVLC";

    static ATOM RegisterVideoClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kVideoClassName[] = L"RedSalamander.ViewerVLC.Video";

    static ATOM RegisterHudClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kHudClassName[] = L"RedSalamander.ViewerVLC.Hud";

    static ATOM RegisterOverlayClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kOverlayClassName[] = L"RedSalamander.ViewerVLC.Overlay";

    static ATOM RegisterSeekPreviewClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kSeekPreviewClassName[] = L"RedSalamander.ViewerVLC.SeekPreview";

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    static LRESULT CALLBACK VideoProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT VideoProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    static LRESULT CALLBACK HudProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT HudProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    static LRESULT CALLBACK OverlayProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT OverlayProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    static LRESULT CALLBACK SeekPreviewProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT SeekPreviewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    static LRESULT CALLBACK WheelForwardProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnCreate(HWND hwnd) noexcept;
    void OnDestroy() noexcept;
    LRESULT OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept;
    void OnSize(UINT width, UINT height) noexcept;
    void OnTimer(UINT_PTR timerId) noexcept;
    LRESULT OnNotify(const NMHDR* hdr) noexcept;
    bool TryRestoreFocusFromEmbeddedSetFocus(WPARAM previousFocusParam) noexcept;

    void Layout(HWND hwnd, UINT width, UINT height) noexcept;

    void UpdatePlaybackUi() noexcept;
    void SetMissingUiVisible(bool visible, std::wstring_view details) noexcept;
    void SetLoadingUiVisible(bool visible) noexcept;

    bool EnsureHudDirect2D(HWND hwnd) noexcept;
    void DiscardHudRenderTarget() noexcept;
    void UpdateHudOpacityTarget(HWND hwnd, bool forceInvalidate) noexcept;
    void CycleHudFocus(bool backwards) noexcept;
    HudPart HitTestHud(HWND hwnd, POINT pt) const noexcept;
    void UpdateHudSeekDrag(HWND hwnd, POINT pt) noexcept;
    void UpdateHudVolumeDrag(HWND hwnd, POINT pt) noexcept;
    void OnHudPaint(HWND hwnd) noexcept;
    void OnHudSize(HWND hwnd, UINT width, UINT height) noexcept;
    void OnHudTimer(HWND hwnd) noexcept;
    void OnHudMouseMove(HWND hwnd, POINT pt) noexcept;
    void OnHudMouseLeave(HWND hwnd) noexcept;
    void OnHudLButtonDown(HWND hwnd, POINT pt) noexcept;
    void OnHudLButtonUp(HWND hwnd, POINT pt) noexcept;
    void OnHudKeyDown(HWND hwnd, UINT vkey) noexcept;
    void OnHudMouseWheel(HWND hwnd, WPARAM wheelParam) noexcept;
    void OnHudKillFocus(HWND hwnd) noexcept;

    [[nodiscard]] bool BuildVlcLoadSpec(VlcLoadSpec& spec, std::wstring& outError) const noexcept;
    [[nodiscard]] static std::unique_ptr<VlcState> LoadVlcState(const VlcLoadSpec& spec, std::wstring& outError) noexcept;
    bool StartPlayback(const std::filesystem::path& path) noexcept;
    bool StartPlaybackWithLoadedVlc(const std::filesystem::path& path) noexcept;
    void BeginAsyncVlcLoad(const std::filesystem::path& path, VlcLoadSpec&& spec) noexcept;
    void OnAsyncVlcLoadComplete(std::unique_ptr<VlcAsyncLoadResult> result) noexcept;
    void StopPlayback() noexcept;
    void RetireCurrentPlayerAsync(bool updateUi) noexcept;
    void RetireVlcStateAsync(std::unique_ptr<VlcState> state, bool updateUi) noexcept;

    void TakeSnapshot() noexcept;

    void TogglePlayPause() noexcept;
    void StopCommand() noexcept;
    void SeekAbsoluteMs(int64_t timeMs) noexcept;
    void SeekRelativeMs(int64_t deltaMs) noexcept;
    void SetVolume(int volume) noexcept;
    void ApplyVolumeToPlayer() noexcept;
    void ToggleMute() noexcept;
    void PersistVolumeState() noexcept;
    void SeekByWheel(int wheelDelta, bool shift, bool ctrl) noexcept;
    void OnSurfaceMouseWheel(WPARAM wheelParam) noexcept;
    [[nodiscard]] wchar_t GetHudPlayPauseGlyph() const noexcept;
    [[nodiscard]] wchar_t GetHudVolumeGlyph() const noexcept;
    [[nodiscard]] int64_t ComputeWheelSeekDeltaMs(bool shift, bool ctrl) const noexcept;
    void RefreshVideoChildWheelForwarding() noexcept;
    void RemoveVideoChildWheelForwarding() noexcept;
    bool InstallVideoChildWheelForwarding(HWND child) noexcept;
#ifdef ENABLE_TESTS
    bool DebugEnsureWheelVideoChild() noexcept;
#endif
    [[nodiscard]] SIZE ComputeSnapshotSize() const noexcept;
    void SetPlaybackRate(float rate) noexcept;
    void StepPlaybackRate(int deltaSteps) noexcept;
    void ToggleFullscreen() noexcept;
    void SetFullscreen(bool enabled) noexcept;

    void ApplyTitleBarTheme(bool windowActive) noexcept;
    void CreateOrUpdateWindowBackgroundBrush() noexcept;

    bool EnsureOverlayDirect2D(HWND hwnd) noexcept;
    void DiscardOverlayRenderTarget() noexcept;
    void OnOverlaySize(HWND hwnd, UINT width, UINT height) noexcept;
    void OnOverlayPaint(HWND hwnd) noexcept;
    void OnOverlayMouseMove(HWND hwnd, POINT pt) noexcept;
    void OnOverlayMouseLeave(HWND hwnd) noexcept;
    void OnOverlayLButtonUp(HWND hwnd, POINT pt) noexcept;
    LRESULT OnOverlaySetCursor(HWND hwnd) noexcept;

    bool EnsureSeekPreviewDirect2D(HWND hwnd) noexcept;
    void DiscardSeekPreviewRenderTarget() noexcept;
    void OnSeekPreviewSize(HWND hwnd, UINT width, UINT height) noexcept;
    void OnSeekPreviewPaint(HWND hwnd) noexcept;
    void UpdateSeekPreviewLayout() noexcept;
    void UpdateSeekPreviewTargetTimeMs(int64_t timeMs) noexcept;
    void ClearSeekPreview() noexcept;

    [[nodiscard]] std::wstring GetOverlayTitleText() const;
    [[nodiscard]] std::wstring GetOverlayBodyText() const;
    [[nodiscard]] std::wstring GetOverlayLinkLabelText() const;
    [[nodiscard]] std::wstring GetOverlayLinkUrl() const;
    [[nodiscard]] HRESULT RebuildConfigurationJson() noexcept;
    [[nodiscard]] bool ConfigurationDiffersFromDefaults() const noexcept;

    struct ViewerVlcConfig
    {
        std::filesystem::path vlcInstallPath;
        bool autoDetectVlc                  = true;
        bool quiet                          = true;
        uint32_t fileCachingMs              = 300;
        uint32_t networkCachingMs           = 1000;
        uint32_t defaultPlaybackRatePercent = 100;
        std::string avcodecHw               = "any";
        std::string videoOutput;
        std::string audioOutput;
        std::string audioVisualization = "visual";
        std::string extraArgs;
        uint32_t lastVolumePercent = 100;
        bool muted                 = false;
    };

    std::atomic_ulong _refCount{1};

    PluginMetaData _metaData{};
    std::wstring _metaId;
    std::wstring _metaShortId;
    std::wstring _metaName;
    std::wstring _metaDescription;

    std::string _configurationJson;

    wil::com_ptr<IHostAlerts> _hostAlerts;

    ViewerVlcConfig _config;

    HWND _embeddedFocusReturnWindow = nullptr;
    wil::unique_hwnd _hVideo;
    wil::unique_hwnd _hHud;

    wil::unique_hwnd _hMissingOverlay;
    wil::unique_hwnd _hSeekPreview;
    RECT _overlayLinkRect{};
    bool _overlayLinkHot       = false;
    bool _overlayTrackingMouse = false;
    std::wstring _overlayDetails;

    wil::unique_hbrush _backgroundBrush;
    COLORREF _backgroundColor = CLR_INVALID;

    enum class HudPart : uint8_t
    {
        None,
        PlayPause,
        Stop,
        Snapshot,
        Seek,
        Speed,
        VolumeMute,
        Volume,
    };

    HudPart _hudHot     = HudPart::None;
    HudPart _hudPressed = HudPart::None;
    HudPart _hudFocus   = HudPart::PlayPause;

    bool _hudTrackingMouse   = false;
    bool _hudSeekDragging    = false;
    bool _hudVolumeDragging  = false;
    float _hudOpacity        = 1.0f;
    float _hudTargetOpacity  = 1.0f;
    UINT_PTR _hudAnimTimerId = 0;

    int _hudVolumeValue            = 100;
    int _hudPreviousVolumeValue    = 100;
    bool _hudMuted                 = false;
    int64_t _hudTimeMs             = 0;
    int64_t _hudLengthMs           = 0;
    bool _hudPlaying               = false;
    int64_t _hudDragTimeMs         = 0;
    float _hudRate                 = 1.0f;
    ULONGLONG _hudLastActivityTick = 0;
    bool _isAudioFile              = false;

    ULONGLONG _videoLastClickTick = 0;
    POINT _videoLastClickPos{};
    std::vector<HWND> _wheelForwardedChildren;

    wil::com_ptr<ID2D1Factory> _hudD2DFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _hudRenderTarget;
    wil::com_ptr<IDWriteFactory> _hudDWriteFactory;
    wil::com_ptr<IDWriteTextFormat> _hudTextFormat;
    wil::com_ptr<IDWriteTextFormat> _hudMonoFormat;
    wil::com_ptr<IDWriteTextFormat> _hudIconFormat;
    UINT _hudTextDpi = 0;

    wil::com_ptr<ID2D1HwndRenderTarget> _overlayRenderTarget;
    wil::com_ptr<IDWriteTextFormat> _overlayTitleFormat;
    wil::com_ptr<IDWriteTextFormat> _overlayBodyFormat;
    wil::com_ptr<IDWriteTextFormat> _overlayLinkFormat;
    UINT _overlayTextDpi = 0;

    wil::com_ptr<ID2D1HwndRenderTarget> _seekPreviewRenderTarget;
    wil::com_ptr<IDWriteTextFormat> _seekPreviewTextFormat;
    UINT _seekPreviewTextDpi = 0;

    int64_t _seekPreviewTargetTimeMs = -1;
    bool _seekDragWasPlaying         = false;

    bool _isFullscreen = false;
    WINDOWPLACEMENT _restorePlacement{sizeof(WINDOWPLACEMENT)};
    DWORD _restoreStyle   = 0;
    DWORD _restoreExStyle = 0;

    UINT_PTR _uiTimerId           = 0;
    bool _missingUiVisible        = false;
    UINT_PTR _loadingTimerId      = 0;
    bool _loadingUiActive         = false;
    bool _loadingUiVisible        = false;
    ULONGLONG _loadingStartedTick = 0;
    uint64_t _asyncOpenGeneration = 0;

    std::unique_ptr<VlcState> _vlc;

    std::filesystem::path _currentPath;

#ifdef ENABLE_TESTS
    wil::unique_hwnd _debugWheelVideoChild;
    uint32_t _debugStopDelayMs = 0;
#endif
};
