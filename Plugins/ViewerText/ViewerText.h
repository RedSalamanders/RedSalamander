#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <array>
#include <atomic>
#include <cstdint>
#include <filesystem>
#include <memory>
#include <optional>
#include <string>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182) // WIL: deleted copy/move operators and padding
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
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

class ViewerText final : public IViewer, public IInformations
{
public:
    static constexpr size_t kHexBytesPerLine = 16;

    ViewerText();
    ~ViewerText();

    void SetHost(IHost* host) noexcept;

    ViewerText(const ViewerText&)            = delete;
    ViewerText(ViewerText&&)                 = delete;
    ViewerText& operator=(const ViewerText&) = delete;
    ViewerText& operator=(ViewerText&&)      = delete;

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
    struct ViewerTextConfig
    {
        uint32_t textBufferMiB = 16;
        uint32_t hexBufferMiB  = 8;
        bool showLineNumbers   = false;
        bool wrapText          = true;
    };

    struct MenuItemData
    {
        std::wstring text;
        std::wstring shortcut;
        bool separator  = false;
        bool topLevel   = false;
        bool hasSubMenu = false;
    };

    struct ByteSpan
    {
        size_t start  = 0;
        size_t length = 0;
    };

    struct SaveAsResult
    {
        std::filesystem::path path;
        UINT encodingSelection = 0;
    };

    enum class ViewMode : uint8_t
    {
        Text,
        Hex,
    };

public:
    enum class FileEncoding : uint8_t
    {
        Utf8,
        Utf16LE,
        Utf16BE,
        Utf32LE,
        Utf32BE,
        Unknown,
    };

private:
    struct AsyncOpenResult
    {
        ViewerText* viewer = nullptr;
        uint64_t requestId = 0;
        HRESULT hr         = E_FAIL;
        std::filesystem::path path;
        bool updateOtherFiles = false;
        ViewMode viewMode     = ViewMode::Text;
        std::wstring title;

        wil::com_ptr<IFileReader> fileReader;
        uint64_t fileSize     = 0;
        FileEncoding encoding = FileEncoding::Unknown;
        uint64_t bomBytes     = 0;

        UINT displayEncodingMenuSelection = 0;
        uint32_t detectedCodePage         = 0;
        bool detectedCodePageValid        = false;
        bool detectedCodePageIsGuess      = false;

        std::wstring statusMessage;

        uint64_t textStreamSkipBytes   = 0;
        uint64_t textStreamStartOffset = 0;
        uint64_t textStreamEndOffset   = 0;
        bool textStreamActive          = false;

        std::wstring textBuffer;
        std::vector<uint32_t> textLineStarts;
        std::vector<uint32_t> textLineEnds;
        uint32_t textMaxLineLength = 0;

        std::vector<uint8_t> hexBytes;
        bool hasHexCache = false;
        std::vector<uint8_t> hexCache;
        uint64_t hexCacheOffset = 0;
        size_t hexCacheValid    = 0;
    };

    enum class HexColumnMode : uint8_t
    {
        Byte,
        Word,
        Dword,
        Qword,
    };

    enum class HexOffsetMode : uint8_t
    {
        Hex,
        Decimal,
    };

    enum class HexTextMode : uint8_t
    {
        Ansi,
        Utf8,
        Utf16,
    };

    enum class HexHeaderHit : uint8_t
    {
        None,
        Offset,
        Data,
        Text,
    };

    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kClassName[] = L"RedSalamander.ViewerText";

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    static ATOM RegisterTextViewClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kTextViewClassName[] = L"RedSalamander.ViewerText.TextView";
    static LRESULT CALLBACK TextViewProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT TextViewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    static ATOM RegisterHexViewClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kHexViewClassName[] = L"RedSalamander.ViewerText.HexView";
    static LRESULT CALLBACK HexViewProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT HexViewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnCreate(HWND hwnd);
    void OnDestroy();
    void OnTimer(UINT_PTR timerId) noexcept;
    void OnAsyncOpenComplete(std::unique_ptr<AsyncOpenResult> result) noexcept;
    void OnSize(UINT width, UINT height);
    void OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggested) noexcept;
    void OnPaint();
    LRESULT OnTextViewPaint(HWND hwnd) noexcept;
    LRESULT OnTextViewSize(HWND hwnd, UINT32 width, UINT32 height) noexcept;
    LRESULT OnTextViewVScroll(HWND hwnd, UINT scrollCode) noexcept;
    LRESULT OnTextViewHScroll(HWND hwnd, UINT scrollCode) noexcept;
    LRESULT OnTextViewMouseWheel(HWND hwnd, int delta) noexcept;
    LRESULT OnTextViewLButtonDown(HWND hwnd, POINT pt) noexcept;
    LRESULT OnTextViewMouseMove(HWND hwnd, POINT pt) noexcept;
    LRESULT OnTextViewLButtonUp(HWND hwnd) noexcept;
    LRESULT OnTextViewKeyDown(HWND hwnd, WPARAM vk, LPARAM lParam) noexcept;
    LRESULT OnTextViewSetFocus(HWND hwnd) noexcept;
    LRESULT OnTextViewKillFocus(HWND hwnd) noexcept;
    LRESULT OnHexViewPaint(HWND hwnd) noexcept;
    LRESULT OnHexViewSize(HWND hwnd, UINT32 width, UINT32 height) noexcept;
    LRESULT OnHexViewVScroll(HWND hwnd, UINT scrollCode) noexcept;
    LRESULT OnHexViewMouseWheel(HWND hwnd, int delta) noexcept;
    LRESULT OnHexViewMouseMove(HWND hwnd, POINT pt) noexcept;
    LRESULT OnHexViewMouseLeave(HWND hwnd) noexcept;
    LRESULT OnHexViewSetCursor(HWND hwnd, LPARAM lParam) noexcept;
    LRESULT OnHexViewLButtonDown(HWND hwnd, POINT pt) noexcept;
    LRESULT OnHexViewKeyDown(HWND hwnd, WPARAM vk, LPARAM lParam) noexcept;
    LRESULT OnHexViewSetFocus(HWND hwnd) noexcept;
    LRESULT OnHexViewKillFocus(HWND hwnd) noexcept;
    void OnCommand(HWND hwnd, UINT commandId, UINT notifyCode, HWND control) noexcept;
    LRESULT OnNotify(const NMHDR* header);
    LRESULT OnMeasureItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept;
    LRESULT OnDrawItem(DRAWITEMSTRUCT* draw) noexcept;
    LRESULT OnCtlColor(UINT msg, HDC hdc, HWND control) noexcept;
    void OnMouseMove(int x, int y) noexcept;
    void OnMouseLeave() noexcept;
    void OnLButtonDown(int x, int y) noexcept;
    void OnLButtonUp(int x, int y) noexcept;
    bool OnSetCursor(HWND hwnd, LPARAM lParam) noexcept;
    void OnNcActivate(bool windowActive) noexcept;
    LRESULT OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept;

    void Layout(HWND hwnd) noexcept;
    void RefreshFileCombo(HWND hwnd) noexcept;
    void SyncFileComboSelection() noexcept;
    void OnMeasureFileComboItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept;
    void OnDrawFileComboItem(DRAWITEMSTRUCT* draw) noexcept;
    void UpdateMenuChecks(HWND hwnd) noexcept;
    void ApplyTheme(HWND hwnd) noexcept;
    void ApplyTitleBarTheme(bool windowActive) noexcept;
    void UpdateStatusText(HWND hwnd) noexcept;

    void SetViewMode(HWND hwnd, ViewMode mode) noexcept;
    void SetWrap(HWND hwnd, bool wrap) noexcept;
    void SetShowLineNumbers(HWND hwnd, bool showLineNumbers) noexcept;

    void CommandOpen(HWND hwnd);
    void CommandSaveAs(HWND hwnd);
    void CommandRefresh(HWND hwnd);
    void CommandExit(HWND hwnd) noexcept;

    void CommandOtherNext(HWND hwnd);
    void CommandOtherPrevious(HWND hwnd);
    void CommandOtherFirst(HWND hwnd);
    void CommandOtherLast(HWND hwnd);

    void CommandFind(HWND hwnd);
    void CommandFindNext(HWND hwnd, bool backward);
    void CommandFindNextHex(HWND hwnd, bool backward);
    void UpdateSearchHighlights() noexcept;
    void CommandGoToOffset(HWND hwnd);
    void CommandGoToTop(HWND hwnd, bool extendSelection) noexcept;
    void CommandGoToBottom(HWND hwnd, bool extendSelection) noexcept;
    void CommandGoToOffsetValue(HWND hwnd, uint64_t offset);

    HRESULT OpenPath(HWND hwnd, const std::filesystem::path& path, bool updateOtherFiles) noexcept;
    void StartAsyncOpen(HWND hwnd, const std::filesystem::path& path, bool updateOtherFiles, UINT displayEncodingMenuSelection) noexcept;
    void BeginLoadingUi() noexcept;
    void EndLoadingUi() noexcept;
    void UpdateLoadingSpinner() noexcept;
    void DrawLoadingOverlay(ID2D1HwndRenderTarget* target, ID2D1SolidColorBrush* brush, float widthDip, float heightDip) noexcept;

    enum class InlineAlertSeverity : uint8_t
    {
        Error,
        Warning,
        Info,
    };

    void ShowInlineAlert(InlineAlertSeverity severity, UINT titleId, UINT messageId) noexcept;
    void ClearInlineAlert() noexcept;

    std::optional<std::filesystem::path> ShowOpenDialog(HWND hwnd) noexcept;
    std::optional<SaveAsResult> ShowSaveAsDialog(HWND hwnd) noexcept;

    HRESULT DetectEncodingAndSize(const std::filesystem::path& path, FileEncoding& encoding, uint64_t& bomBytes, uint64_t& fileSize) noexcept;
    std::wstring EncodingLabel() const;

    HRESULT LoadTextToEdit(HWND hwnd, uint64_t startOffset, bool scrollToEnd) noexcept;
    void UpdateTextStreamTotalLineCountAfterLoad() noexcept;
    void RebuildTextLineIndex() noexcept;
    void RebuildTextVisualLines(HWND hwnd) noexcept;
    void UpdateTextViewScrollBars(HWND hwnd) noexcept;
    bool TryNavigateTextStream(HWND hwnd, bool backward) noexcept;
    uint64_t TextStreamChunkBytes() const noexcept;
    uint64_t AlignTextStreamOffset(uint64_t offset) const noexcept;
    HRESULT LoadHexData(HWND hwnd) noexcept;
    void UpdateHexViewScrollBars(HWND hwnd) noexcept;
    void ResetHexState() noexcept;

    void UpdateHexColumns(HWND hwnd) noexcept;
    void UpdateHexItemCount(HWND hwnd) noexcept;
    void FormatHexLine(uint64_t offset,
                       std::wstring& outOffset,
                       std::wstring& outHex,
                       std::wstring& outAscii,
                       std::array<ByteSpan, kHexBytesPerLine>& hexSpans,
                       std::array<ByteSpan, kHexBytesPerLine>& textSpans,
                       size_t& validBytes) noexcept;
    void EnsureHexLineCache(int item) noexcept;
    size_t ReadHexBytes(uint64_t offset, uint8_t* dest, size_t destSize) noexcept;
    HRESULT RefillHexCache(uint64_t offset) noexcept;
    std::wstring FormatFileOffset(uint64_t offset) const;
    std::wstring BuildStatusText() const;
    void UpdateHexTextColumnHeader() noexcept;
    void UpdateHexColumnHeader() noexcept;
    size_t HexGroupSize() const noexcept;
    void CycleHexColumnMode() noexcept;
    void CycleHexOffsetMode() noexcept;
    void CycleHexTextMode() noexcept;
    bool HexBigEndian() const noexcept;
    void OnHexMouseDown(HWND hwnd, int x, int y) noexcept;
    void OnHexMouseMove(HWND hwnd, int x, int y) noexcept;
    void OnHexMouseUp(HWND hwnd) noexcept;
    void CopyHexCsvToClipboard(HWND hwnd) noexcept;

    void SetDisplayEncodingMenuSelection(HWND hwnd, UINT commandId, bool reload) noexcept;
    void SetSaveEncodingMenuSelection(HWND hwnd, UINT commandId) noexcept;
    void CommandCycleDisplayEncoding(HWND hwnd, bool backward) noexcept;
    bool IsEncodingMenuSelectionValid(UINT commandId) const noexcept;
    bool IsSaveEncodingMenuSelectionValid(UINT commandId) const noexcept;
    UINT EffectiveDisplayEncodingMenuSelection() const noexcept;
    UINT EffectiveSaveEncodingMenuSelection() const noexcept;
    uint64_t BytesToSkipForDisplayEncoding() const noexcept;
    bool DisplayEncodingUsesUnicodeStream() const noexcept;
    FileEncoding DisplayEncodingFileEncoding() const noexcept;
    UINT DisplayEncodingCodePage() const noexcept;
    UINT CodePageForMenuSelection(UINT commandId) const noexcept;

    bool HandleShortcutKey(HWND hwnd, WPARAM vk) noexcept;
    bool EnsureDirect2D(HWND hwnd) noexcept;
    void DiscardDirect2D() noexcept;
    void ApplyMenuTheme(HWND hwnd) noexcept;
    void PrepareMenuTheme(HMENU menu, bool topLevel, std::vector<MenuItemData>& outItems) noexcept;
    void OnMeasureMenuItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept;
    void OnDrawMenuItem(DRAWITEMSTRUCT* draw) noexcept;

    bool EnsureTextViewDirect2D(HWND hwnd) noexcept;
    void DiscardTextViewDirect2D() noexcept;
    bool EnsureHexViewDirect2D(HWND hwnd) noexcept;
    void DiscardHexViewDirect2D() noexcept;

private:
    static constexpr UINT_PTR kLoadingDelayTimerId = 3;
    static constexpr UINT_PTR kLoadingAnimTimerId  = 4;

    std::atomic_ulong _refCount{1};

    PluginMetaData _metaData{};
    std::wstring _metaId;
    std::wstring _metaShortId;
    std::wstring _metaName;
    std::wstring _metaDescription;

    std::string _configurationJson;
    ViewerTextConfig _config;

    IViewerCallback* _callback = nullptr;
    void* _callbackCookie      = nullptr;

    ViewerTheme _theme{};
    bool _hasTheme                = false;
    bool _allowEraseBkgnd         = true;
    bool _allowEraseBkgndTextView = true;
    bool _allowEraseBkgndHexView  = true;

    wil::unique_hmodule _msftEditModule;

    wil::unique_hwnd _hWnd;
    wil::unique_hwnd _hEdit;
    wil::unique_hwnd _hHex;
    wil::unique_hwnd _hFileCombo;

    wil::unique_hicon _windowIconSmall;
    wil::unique_hicon _windowIconBig;

    wil::unique_hbrush _backgroundBrush;
    wil::unique_hbrush _headerBrush;
    wil::unique_hbrush _statusBrush;
    wil::unique_hfont _uiFont;
    wil::unique_hfont _monoFont;

    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _d2dTarget;
    wil::com_ptr<IDWriteTextFormat> _headerFormat;
    wil::com_ptr<IDWriteTextFormat> _headerFormatRight;
    wil::com_ptr<IDWriteTextFormat> _modeButtonFormat;
    wil::com_ptr<IDWriteTextFormat> _statusFormat;
    wil::com_ptr<IDWriteTextFormat> _watermarkFormat;
    wil::com_ptr<IDWriteTextFormat> _loadingOverlayFormat;
    wil::com_ptr<ID2D1SolidColorBrush> _d2dBrush;

    wil::com_ptr<ID2D1HwndRenderTarget> _textViewTarget;
    wil::com_ptr<ID2D1SolidColorBrush> _textViewBrush;
    wil::com_ptr<IDWriteTextFormat> _textViewFormat;
    wil::com_ptr<IDWriteTextFormat> _textViewFormatRight;
    float _textCharWidthDip  = 0.0f;
    float _textLineHeightDip = 0.0f;

    wil::com_ptr<ID2D1HwndRenderTarget> _hexViewTarget;
    wil::com_ptr<ID2D1SolidColorBrush> _hexViewBrush;
    wil::com_ptr<IDWriteTextFormat> _hexViewFormat;
    wil::com_ptr<IDWriteTextFormat> _hexViewFormatRight;
    float _hexCharWidthDip  = 0.0f;
    float _hexLineHeightDip = 0.0f;

    COLORREF _uiBackground = RGB(255, 255, 255);
    COLORREF _uiText       = RGB(0, 0, 0);
    COLORREF _uiHeaderBg   = RGB(255, 255, 255);
    COLORREF _uiStatusBg   = RGB(255, 255, 255);

    HWND _hFileComboList = nullptr;
    HWND _hFileComboItem = nullptr;

    RECT _modeButtonRect{};
    bool _modeButtonHot      = false;
    bool _modeButtonPressed  = false;
    bool _trackingMouseLeave = false;

    RECT _headerRect{};
    RECT _contentRect{};
    RECT _statusRect{};

    ViewMode _viewMode           = ViewMode::Text;
    bool _wrap                   = true;
    HexColumnMode _hexColumnMode = HexColumnMode::Byte;
    HexOffsetMode _hexOffsetMode = HexOffsetMode::Hex;
    HexTextMode _hexTextMode     = HexTextMode::Ansi;

    std::filesystem::path _currentPath;
    uint64_t _fileSize     = 0;
    FileEncoding _encoding = FileEncoding::Unknown;
    uint64_t _bomBytes     = 0;

    wil::com_ptr<IFileSystem> _fileSystem;
    std::wstring _fileSystemName;
    wil::com_ptr<IFileReader> _fileReader;

    std::vector<std::filesystem::path> _otherFiles;
    size_t _otherIndex     = 0;
    bool _syncingFileCombo = false;
    std::vector<std::filesystem::path> _selection;
    std::wstring _searchQuery;
    std::vector<size_t> _searchMatchStarts;
    std::vector<uint8_t> _hexSearchNeedle;
    bool _hexSearchNeedleValid = false;
    std::wstring _statusMessage;

    std::atomic_uint64_t _asyncOpenRequestId{0};
    uint64_t _activeAsyncOpenRequestId  = 0;
    bool _isLoading                     = false;
    bool _showLoadingOverlay            = false;
    float _loadingSpinnerAngleDeg       = 0.0f;
    ULONGLONG _loadingSpinnerLastTickMs = 0;

    wil::com_ptr<IHostAlerts> _hostAlerts;

    std::wstring _textBuffer;
    std::vector<uint32_t> _textLineStarts;
    std::vector<uint32_t> _textLineEnds;
    std::vector<uint32_t> _textVisualLineStarts;
    std::vector<uint32_t> _textVisualLineLogical;
    uint32_t _textTopVisualLine = 0;
    uint32_t _textLeftColumn    = 0;
    uint32_t _textMaxLineLength = 0;
    uint32_t _textWrapColumns   = 0;
    size_t _textCaretIndex      = 0;
    size_t _textSelAnchor       = 0;
    size_t _textSelActive       = 0;
    size_t _textPreferredColumn = 0;
    bool _textSelecting         = false;

    bool _textStreamActive          = false;
    uint64_t _textStreamSkipBytes   = 0;
    uint64_t _textStreamStartOffset = 0;
    uint64_t _textStreamEndOffset   = 0;
    std::optional<uint64_t> _textTotalLineCount;
    uint64_t _textStreamLineCountedEndOffset = 0;
    uint64_t _textStreamLineCountedNewlines  = 0;
    bool _textStreamLineCountLastWasCR       = false;

    UINT _displayEncodingMenuSelection = 0;
    UINT _saveEncodingMenuSelection    = 0;
    UINT _detectedCodePage             = 0;
    bool _detectedCodePageValid        = false;
    bool _detectedCodePageIsGuess      = false;

    std::vector<uint8_t> _hexBytes;
    std::optional<uint64_t> _hexSelectionAnchorOffset;
    std::optional<uint64_t> _hexSelectedOffset;
    uint64_t _hexTopLine        = 0;
    bool _hexSelecting          = false;
    HexHeaderHit _hexHeaderHot  = HexHeaderHit::None;
    bool _hexTrackingMouseLeave = false;

    std::vector<uint8_t> _hexCache;
    uint64_t _hexCacheOffset       = 0;
    size_t _hexCacheValid          = 0;
    int _hexLineCacheItem          = -1;
    size_t _hexLineCacheValidBytes = 0;
    std::wstring _hexLineCacheOffsetText;
    std::wstring _hexLineCacheHexText;
    std::wstring _hexLineCacheAsciiText;
    std::array<ByteSpan, kHexBytesPerLine> _hexLineCacheHexSpans{};
    std::array<ByteSpan, kHexBytesPerLine> _hexLineCacheTextSpans{};

    std::vector<MenuItemData> _menuThemeItems;
};
