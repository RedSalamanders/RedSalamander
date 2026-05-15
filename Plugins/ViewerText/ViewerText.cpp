#include "ViewerText.h"

#include "ViewerText.ThemeHelpers.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cerrno>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <format>
#include <functional>
#include <iterator>
#include <limits>
#include <memory>
#include <new>
#include <string_view>
#include <thread>
#include <unordered_map>

#include <commdlg.h>
#include <d2d1.h>
#include <d2d1helper.h>
#include <dwmapi.h>
#include <dwrite.h>
#include <richedit.h>
#include <shobjidl_core.h>
#include <uxtheme.h>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#pragma comment(lib, "d2d1")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "uxtheme")

#include "DxUi/DxUi.Typography.h"
#include "Helpers.h"
#include "WindowMessages.h"
#include "WindowSizing.h"

#include "resource.h"

extern HINSTANCE g_hInstance;

using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::MakeDefaultThemePalette;
using RedSalamander::DxUi::MakeThemePaletteFromViewerTheme;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;
namespace Typography = RedSalamander::DxUi::Typography;

enum class ParsedDiffLineKind : uint8_t
{
    Context,
    Added,
    Removed,
    NoNewlineMarker,
};

struct ParsedDiffLine
{
    ParsedDiffLineKind kind = ParsedDiffLineKind::Context;
    std::wstring text;
    uint32_t oldLine = 0u;
    uint32_t newLine = 0u;
    bool hasOldLine  = false;
    bool hasNewLine  = false;
};

struct ParsedDiffHunk
{
    std::wstring header;
    uint32_t oldStart = 0u;
    uint32_t newStart = 0u;
    std::vector<ParsedDiffLine> lines;
};

struct ParsedDiffFileSection
{
    std::vector<std::wstring> metadataLines;
    std::wstring leftDisplayPath;
    std::wstring rightDisplayPath;
    std::vector<ParsedDiffHunk> hunks;
};

struct ParsedDiffDocument
{
    std::vector<ParsedDiffFileSection> files;
};

struct ResolvedDiffTextFile
{
    bool available                 = false;
    bool existsInCurrentFileSystem = false;
    std::wstring reason;
    std::filesystem::path resolvedPath;
    wil::com_ptr<IFileReader> reader;
    ViewerText::FileEncoding encoding = ViewerText::FileEncoding::Unknown;
    UINT codePage                     = CP_ACP;
    uint64_t fileSize                 = 0u;
    size_t bomBytes                   = 0u;
    uint64_t nextReadOffset           = 0u;
    uint64_t bytesRead                = 0u;
    bool loadComplete                 = false;
    bool hadDecodeFailure             = false;
    std::vector<uint8_t> pendingBytes;
    std::vector<std::wstring> lines;
};

struct DiffReferenceCache
{
    std::unordered_map<std::wstring, std::shared_ptr<ResolvedDiffTextFile>> files;
};

namespace
{
constexpr int kHeaderHeightDip                        = 28;
constexpr int kStatusHeightDip                        = 22;
constexpr float kWatermarkAngleDegrees                = -22.0f;
constexpr float kWatermarkFontSizeDip                 = 56.0f;
constexpr uint64_t kMaxHexLoadBytes                   = 128u * 1024u * 1024u; // 128 MiB
constexpr UINT kAsyncOpenCompleteMessage              = WndMsg::kViewerTextAsyncOpenComplete;
constexpr UINT kLoadingDelayMs                        = 500u;
constexpr UINT kLoadingAnimIntervalMs                 = 16u;
constexpr float kLoadingSpinnerDegPerSec              = 90.0f;
constexpr uint64_t kMaxFullyBufferedParsedDiffBytes   = 16u * 1024u * 1024u;
constexpr uint64_t kMaxBoundedStreamedDiffIndexBytes  = 64u * 1024u * 1024u;
constexpr uint64_t kMaxReferencedDiffFileBytes        = 2u * 1024u * 1024u;
constexpr size_t kReferencedDiffProbeBytes            = 16u * 1024u;
constexpr size_t kReferencedDiffChunkBytes            = 16u * 1024u;
constexpr size_t kMaxBoundedStreamedDiffSections      = 4096u;
constexpr size_t kViewerComboPopupMaxVisibleItems     = 8u;
constexpr wchar_t kFileComboHostOriginalWndProcProp[] = L"RS.ViewerText.FileComboHostOriginalWndProc";
constexpr wchar_t kFileComboHostStateProp[]           = L"RS.ViewerText.FileComboHostState";
constexpr wchar_t kViewerTextPromptWindowClassName[]  = L"RedSalamander.ViewerText.Prompt";

[[nodiscard]] size_t CountOwnerDrawMenuItems(HMENU menu) noexcept
{
    if (! menu)
    {
        return 0u;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return 0u;
    }

    size_t ownerDrawCount = 0u;
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_SUBMENU;
        if (GetMenuItemInfoW(menu, position, TRUE, &itemInfo) == 0)
        {
            continue;
        }

        if ((itemInfo.fType & MFT_OWNERDRAW) != 0)
        {
            ++ownerDrawCount;
        }

        if (itemInfo.hSubMenu)
        {
            ownerDrawCount += CountOwnerDrawMenuItems(itemInfo.hSubMenu);
        }
    }

    return ownerDrawCount;
}

static const int kViewerTextModuleAnchor = 0;

struct DiffVariantBuildResult
{
    ViewerText::DiffTextVariant variant;
};

[[nodiscard]] const char* DiffDefaultLayoutToConfigString(ViewerText::DiffDefaultLayout value) noexcept;
[[nodiscard]] const char* DiffContextModeToConfigString(ViewerText::DiffContextMode value) noexcept;
[[nodiscard]] const char* DiffAutoOpenModeToConfigString(ViewerText::DiffAutoOpenMode value) noexcept;
[[nodiscard]] ViewerText::DiffDefaultLayout ParseDiffDefaultLayout(std::string_view value) noexcept;
[[nodiscard]] ViewerText::DiffContextMode ParseDiffContextMode(std::string_view value) noexcept;
[[nodiscard]] ViewerText::DiffAutoOpenMode ParseDiffAutoOpenMode(std::string_view value) noexcept;
[[nodiscard]] bool HasDiffLikeExtension(const std::filesystem::path& path) noexcept;
[[nodiscard]] bool LooksLikeUnifiedDiffText(std::wstring_view text) noexcept;
[[nodiscard]] bool ParseUnifiedDiffDocument(std::wstring_view text, ParsedDiffDocument& outDocument) noexcept;
[[nodiscard]] std::wstring TrimDiffPathLabel(std::wstring_view text);
[[nodiscard]] std::wstring StripGitDiffPrefix(std::wstring_view path);
[[nodiscard]] std::wstring FormatDiffSummaryPath(std::wstring_view path);
[[nodiscard]] std::wstring BuildDiffSectionNavigationLabel(const ParsedDiffFileSection& file, size_t fileIndex);
[[nodiscard]] std::wstring BuildDiffSectionNavigationLabel(std::wstring_view leftPath, std::wstring_view rightPath, size_t fileIndex);
[[nodiscard]] std::wstring BuildDiffHunkNavigationLabel(const ParsedDiffFileSection& file,
                                                        size_t fileIndex,
                                                        const ParsedDiffHunk& hunk,
                                                        size_t hunkIndexInFile);
[[nodiscard]] std::wstring BuildHiddenContextBannerLabel(uint32_t oldHiddenLineCount, uint32_t newHiddenLineCount);
[[nodiscard]] bool ParseUnifiedRange(std::wstring_view text, uint32_t& startOut, uint32_t& countOut, size_t& consumedChars) noexcept;
[[nodiscard]] bool ParseUnifiedHunkHeader(std::wstring_view line, uint32_t& oldStartOut, uint32_t& newStartOut) noexcept;
[[nodiscard]] std::wstring DecodeBytesToWide(std::span<const uint8_t> bytes, ViewerText::FileEncoding encoding, UINT codePage, HRESULT& hrOut) noexcept;
[[nodiscard]] std::wstring DecodeBytesToWide(const std::vector<uint8_t>& bytes, ViewerText::FileEncoding encoding, UINT codePage, HRESULT& hrOut) noexcept;
[[nodiscard]] bool BuildBoundedStreamedDiffSectionIndex(IFileReader* reader,
                                                        uint64_t fileSize,
                                                        uint64_t streamSkipBytes,
                                                        ViewerText::FileEncoding encoding,
                                                        UINT codePage,
                                                        std::vector<ViewerText::StreamedDiffSectionEntry>& outSections) noexcept;
[[nodiscard]] std::shared_ptr<ResolvedDiffTextFile> OpenReferencedDiffTextFile(IFileSystemIO* fileIo, const std::filesystem::path& path) noexcept;
[[nodiscard]] std::shared_ptr<ResolvedDiffTextFile> ResolveDiffReference(
    IFileSystemIO* fileIo,
    const std::filesystem::path& diffPath,
    std::wstring_view label,
    UINT missingFileMessageId,
    std::unordered_map<std::wstring, std::shared_ptr<ResolvedDiffTextFile>>& cache) noexcept;
[[nodiscard]] bool EnsureReferencedDiffLinesLoaded(ResolvedDiffTextFile& file, uint32_t lineNumberInclusive) noexcept;
[[nodiscard]] std::wstring_view TryGetReferencedDiffLine(const ResolvedDiffTextFile& file, uint32_t lineNumber) noexcept;
[[nodiscard]] DiffVariantBuildResult BuildInlineDiffText(const ParsedDiffDocument& document,
                                                         ViewerText::DiffContextMode contextMode,
                                                         const std::filesystem::path& diffPath,
                                                         IFileSystemIO* fileIo,
                                                         DiffReferenceCache* referenceCache                                = nullptr,
                                                         std::optional<size_t> hydratedSectionIndex                        = std::nullopt,
                                                         std::optional<std::pair<uint32_t, uint32_t>> hydratedLogicalRange = std::nullopt) noexcept;
[[nodiscard]] DiffVariantBuildResult BuildSideBySideDiffText(const ParsedDiffDocument& document,
                                                             ViewerText::DiffContextMode contextMode,
                                                             const std::filesystem::path& diffPath,
                                                             IFileSystemIO* fileIo,
                                                             DiffReferenceCache* referenceCache                                = nullptr,
                                                             std::optional<size_t> hydratedSectionIndex                        = std::nullopt,
                                                             std::optional<std::pair<uint32_t, uint32_t>> hydratedLogicalRange = std::nullopt) noexcept;
[[nodiscard]] bool ShouldHydrateExpandedDiffSection(ViewerText::DiffContextMode contextMode,
                                                    std::optional<size_t> hydratedSectionIndex,
                                                    size_t fileIndex) noexcept;
[[nodiscard]] bool ShouldHydrateExpandedDiffLogicalLine(std::optional<std::pair<uint32_t, uint32_t>> hydratedLogicalRange, uint32_t logicalLine) noexcept;
[[nodiscard]] std::wstring MaskDeferredDiffText(std::wstring_view text);
void AppendLine(std::wstring& target, std::wstring_view line);
[[nodiscard]] size_t DecimalDigits(uint32_t value) noexcept;

[[nodiscard]] WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept;
[[nodiscard]] bool InstallWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp, WNDPROC hookWndProc) noexcept;
[[nodiscard]] LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* originalWndProcProp, UINT msg, WPARAM wp, LPARAM lp) noexcept;
void UnhookFileComboHostWindow(HWND hwnd) noexcept;
[[nodiscard]] bool MessageMayOpenWindowComboPopup(UINT msg, WPARAM wp) noexcept;
[[nodiscard]] int ComputeWindowComboPopupHeightPx(size_t itemCount, UINT dpi) noexcept;

[[nodiscard]] HWND NormalizeOwnerWindow(HWND ownerWindow) noexcept
{
    if (! ownerWindow || IsWindow(ownerWindow) == FALSE)
    {
        return nullptr;
    }

    return GetAncestor(ownerWindow, GA_ROOT);
}

void CenterWindowOnOwner(HWND window, HWND owner) noexcept
{
    if (! window || IsWindow(window) == FALSE || ! owner || IsWindow(owner) == FALSE)
    {
        return;
    }

    RECT ownerRect{};
    RECT windowRect{};
    if (GetWindowRect(owner, &ownerRect) == FALSE || GetWindowRect(window, &windowRect) == FALSE)
    {
        return;
    }

    const int x = ownerRect.left + (((ownerRect.right - ownerRect.left) - (windowRect.right - windowRect.left)) / 2);
    const int y = ownerRect.top + (((ownerRect.bottom - ownerRect.top) - (windowRect.bottom - windowRect.top)) / 2);
    SetWindowPos(window, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

[[nodiscard]] int ScaleForDpi(const UINT dpi, const int dip) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi == 0u ? 96u : dpi), 96);
}

class ViewerTextPromptWindow final
{
public:
    ViewerTextPromptWindow(const ViewerTextPromptWindow&)            = delete;
    ViewerTextPromptWindow& operator=(const ViewerTextPromptWindow&) = delete;
    ViewerTextPromptWindow(ViewerTextPromptWindow&&)                 = delete;
    ViewerTextPromptWindow& operator=(ViewerTextPromptWindow&&)      = delete;

    ViewerTextPromptWindow(HWND ownerWindow, const ViewerTheme* theme, std::wstring caption, std::wstring label, std::wstring initialText) noexcept
        : _ownerWindow(NormalizeOwnerWindow(ownerWindow)),
          _caption(std::move(caption)),
          _labelText(std::move(label)),
          _initialText(std::move(initialText))
    {
        if (theme)
        {
            _theme    = *theme;
            _hasTheme = true;
        }
    }

    [[nodiscard]] HRESULT ShowModal(std::wstring& textOut) noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return classHr;
        }

        const DWORD style        = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle      = WS_EX_DLGMODALFRAME;
        const UINT dpi           = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();
        const int clientWidthPx  = ScaleForDpi(dpi, 420);
        const int clientHeightPx = ScaleForDpi(dpi, 164);

        RECT bounds{0, 0, clientWidthPx, clientHeightPx};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
        const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled] noexcept
        {
            if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
            {
                EnableWindow(_ownerWindow, TRUE);
                SetActiveWindow(_ownerWindow);
            }
        });

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kViewerTextPromptWindowClassName,
                                          _caption.c_str(),
                                          style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          g_hInstance,
                                          this);
        if (! hwnd)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        CenterWindowOnOwner(_hWnd.get(), _ownerWindow);
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            if (getMessageResult == 0)
            {
                _done   = true;
                _result = S_FALSE;
                PostQuitMessage(static_cast<int>(msg.wParam));
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        if (_result == S_OK)
        {
            textOut = _acceptedText;
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<ViewerTextPromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
        }

        auto* self = reinterpret_cast<ViewerTextPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            else if (message == WM_NCDESTROY)
            {
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                if (! self->_done)
                {
                    self->_done   = true;
                    self->_result = S_FALSE;
                }
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND:
            {
                HDC hdc = reinterpret_cast<HDC>(wParam);
                if (! hdc)
                {
                    return 1;
                }

                RECT client{};
                GetClientRect(hwnd, &client);
                const COLORREF bg       = self->_hasTheme ? ColorRefFromArgb(self->_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
                const COLORREF oldColor = SetDCBrushColor(hdc, bg);
                FillRect(hdc, &client, reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
                SetDCBrushColor(hdc, oldColor);
                return 1;
            }
            case WM_CLOSE: self->Cancel(); return 0;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = ViewerTextPromptWindow::WndProc;
        wc.hInstance     = g_hInstance;
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kViewerTextPromptWindowClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        ApplyTheme();
        Layout();
        _dxHost.SetFocusControl(_field);
        return true;
    }

    void BuildUi()
    {
        if (_root != nullptr)
        {
            return;
        }

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _label = _root->AddChild<Label>(_labelText);
        _label->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _field = _root->AddChild<TextField>(_initialText);

        _okButton = _root->AddChild<Button>(LoadStringResource(g_hInstance, IDS_VIEWERTEXT_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(g_hInstance, IDS_VIEWERTEXT_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
    }

    void ApplyTheme() noexcept
    {
        _palette = _hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false);
        _dxHost.SetTheme(_palette);
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip       = 16.0f;
        constexpr float kLabelHeightDip  = 22.0f;
        constexpr float kFieldHeightDip  = 32.0f;
        constexpr float kButtonHeightDip = 34.0f;
        constexpr float kButtonWidthDip  = 96.0f;
        constexpr float kGapDip          = 8.0f;

        const float left  = client.left + kMarginDip;
        const float right = std::max(left, client.right - kMarginDip);
        float y           = client.top + kMarginDip;

        if (_label)
        {
            _label->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + kGapDip;

        if (_field)
        {
            _field->SetBounds(D2D1::RectF(left, y, right, y + kFieldHeightDip));
        }

        const float buttonsTop = std::max(y + kFieldHeightDip + kGapDip, client.bottom - kMarginDip - kButtonHeightDip);
        const float cancelLeft = std::max(left, right - kButtonWidthDip);
        const float okLeft     = std::max(left, cancelLeft - kGapDip - kButtonWidthDip);

        if (_okButton)
        {
            _okButton->SetBounds(D2D1::RectF(okLeft, buttonsTop, okLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
    }

    void Confirm() noexcept
    {
        _acceptedText = _field ? _field->GetText() : std::wstring{};
        _result       = S_OK;
        _done         = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            DestroyWindow(_hWnd.get());
        }
    }

    void Cancel() noexcept
    {
        _result = S_FALSE;
        _done   = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            DestroyWindow(_hWnd.get());
        }
    }

    HWND _ownerWindow = nullptr;
    ViewerTheme _theme{};
    bool _hasTheme = false;
    ThemePalette _palette{};
    std::wstring _caption;
    std::wstring _labelText;
    std::wstring _initialText;
    std::wstring _acceptedText;
    wil::unique_hwnd _hWnd;
    WindowHost _dxHost;
    std::unique_ptr<Panel> _rootStorage;
    Panel* _root          = nullptr;
    Label* _label         = nullptr;
    TextField* _field     = nullptr;
    Button* _okButton     = nullptr;
    Button* _cancelButton = nullptr;
    bool _done            = false;
    HRESULT _result       = E_ABORT;
};

[[nodiscard]] HRESULT ShowViewerTextPromptDialog(
    HWND ownerWindow, const ViewerTheme* theme, const UINT captionId, const UINT labelId, std::wstring initialText, std::wstring& resultText) noexcept
{
    ViewerTextPromptWindow window(
        ownerWindow, theme, LoadStringResource(g_hInstance, captionId), LoadStringResource(g_hInstance, labelId), std::move(initialText));
    return window.ShowModal(resultText);
}

LRESULT CALLBACK FileComboHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<ViewerText*>(GetPropW(hwnd, kFileComboHostStateProp));
    if (! self)
    {
        return CallStoredWndProc(hwnd, kFileComboHostOriginalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kFileComboHostOriginalWndProcProp);
        RemovePropW(hwnd, kFileComboHostStateProp);
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kFileComboHostOriginalWndProcProp, FileComboHostWndProc);

        bool handled = false;
        static_cast<void>(self->HandleFileComboHostMessage(hwnd, msg, wp, lp, handled));

        return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    bool handled           = false;
    const LRESULT dxResult = self->HandleFileComboHostMessage(hwnd, msg, wp, lp, handled);
    if (handled)
    {
        return dxResult;
    }

    if (msg == WM_KEYUP && (wp == VK_ESCAPE || wp == VK_TAB))
    {
        self->FocusMainSurfaceFromFileCombo(GetAncestor(hwnd, GA_ROOT));
        return 0;
    }

    if (msg == WM_KEYDOWN && wp == VK_ESCAPE)
    {
        const HWND root = GetAncestor(hwnd, GA_ROOT);
        if (root)
        {
            PostMessageW(root, WM_CLOSE, 0, 0);
            return 0;
        }
    }

    return CallStoredWndProc(hwnd, kFileComboHostOriginalWndProcProp, msg, wp, lp);
}

int HexNibbleValue(wchar_t ch) noexcept
{
    if (ch >= L'0' && ch <= L'9')
    {
        return static_cast<int>(ch - L'0');
    }
    if (ch >= L'a' && ch <= L'f')
    {
        return 10 + static_cast<int>(ch - L'a');
    }
    if (ch >= L'A' && ch <= L'F')
    {
        return 10 + static_cast<int>(ch - L'A');
    }
    return -1;
}

bool TryParseHexSearchNeedle(std::wstring_view query, std::vector<uint8_t>& outBytes) noexcept
{
    outBytes.clear();

    std::wstring digits;
    digits.reserve(query.size());

    for (size_t i = 0; i < query.size(); ++i)
    {
        const wchar_t ch = query[i];
        if (ch == L'0' && (i + 1) < query.size() && (query[i + 1] == L'x' || query[i + 1] == L'X'))
        {
            i += 1;
            continue;
        }

        if (iswspace(ch) != 0 || ch == L',' || ch == L';' || ch == L':' || ch == L'_')
        {
            continue;
        }

        if (HexNibbleValue(ch) >= 0)
        {
            digits.push_back(ch);
            continue;
        }

        return false;
    }

    if (digits.empty())
    {
        return false;
    }

    if ((digits.size() % 2u) == 1u)
    {
        digits.insert(digits.begin(), L'0');
    }

    outBytes.reserve(digits.size() / 2u);
    for (size_t i = 0; i + 1 < digits.size(); i += 2)
    {
        const int hi = HexNibbleValue(digits[i]);
        const int lo = HexNibbleValue(digits[i + 1]);
        if (hi < 0 || lo < 0)
        {
            outBytes.clear();
            return false;
        }

        outBytes.push_back(static_cast<uint8_t>((hi << 4u) | lo));
    }

    return ! outBytes.empty();
}

constexpr char kViewerTextSchemaJson[] = R"json({
    "version": 1,
    "title": "Text Viewer",
    "fields": [
        {
            "key": "textBufferMiB",
            "type": "value",
            "label": "Text buffer (MiB)",
            "description": "Approximate in-memory read buffer used by the streaming text renderer.",
            "default": 16,
            "min": 1,
            "max": 256
        },
        {
            "key": "hexBufferMiB",
            "type": "value",
            "label": "Hex buffer (MiB)",
            "description": "Approximate in-memory read buffer used by the streaming hex renderer.",
            "default": 8,
            "min": 1,
            "max": 256
        },
        {
            "key": "showLineNumbers",
            "type": "option",
            "label": "Line numbers",
            "description": "Show logical line numbers (newline-delimited).",
            "default": "0",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        },
        {
            "key": "wrapText",
            "type": "option",
            "label": "Wrap",
            "description": "Wrap long lines in text mode.",
            "default": "1",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        },
        {
            "key": "hexByteColorMode",
            "type": "option",
            "label": "Hex byte colors",
            "description": "Color-code visible bytes in hex mode to make binary patterns easier to spot.",
            "default": "leadingNibble",
            "options": [
                { "value": "leadingNibble", "label": "Leading nibble (00/FF emphasized)" },
                { "value": "off", "label": "Off" }
            ]
        },
        {
            "key": "diffDefaultLayout",
            "type": "option",
            "label": "Default diff layout",
            "description": "Choose how parsed diff documents open by default.",
            "default": "sideBySide",
            "options": [
                { "value": "sideBySide", "label": "Side by side" },
                { "value": "inline", "label": "Inline" }
            ]
        },
        {
            "key": "diffContextMode",
            "type": "option",
            "label": "Unchanged text",
            "description": "Show only changed hunks or expand unchanged text when the referenced files are available.",
            "default": "hunksOnly",
            "options": [
                { "value": "hunksOnly", "label": "Changed hunks only" },
                { "value": "fullFileWhenAvailable", "label": "Show unchanged text when referenced files are available" }
            ]
        },
        {
            "key": "diffAutoOpenMode",
            "type": "option",
            "label": "Open recognized diff files as",
            "description": "Choose whether .diff, .patch, and .rej files open as parsed diffs or raw text by default.",
            "default": "parsed",
            "options": [
                { "value": "parsed", "label": "Parsed diff" },
                { "value": "rawText", "label": "Raw text" }
            ]
        }
    ]
})json";

[[nodiscard]] const char* GetViewerTextStaticConfigurationSchemaImpl() noexcept
{
    return kViewerTextSchemaJson;
}

bool IsValidUtf8(const uint8_t* data, size_t size) noexcept
{
    if (! data || size == 0)
    {
        return true;
    }

    size_t i = 0;
    while (i < size)
    {
        const uint8_t b0 = data[i];
        if (b0 <= 0x7Fu)
        {
            i += 1;
            continue;
        }

        if (b0 < 0xC2u)
        {
            return false;
        }

        if (b0 <= 0xDFu)
        {
            if ((i + 1) >= size)
            {
                return true;
            }

            const uint8_t b1 = data[i + 1];
            if ((b1 & 0xC0u) != 0x80u)
            {
                return false;
            }

            i += 2;
            continue;
        }

        if (b0 <= 0xEFu)
        {
            if ((i + 2) >= size)
            {
                return true;
            }

            const uint8_t b1 = data[i + 1];
            const uint8_t b2 = data[i + 2];

            if ((b1 & 0xC0u) != 0x80u || (b2 & 0xC0u) != 0x80u)
            {
                return false;
            }

            if (b0 == 0xE0u && b1 < 0xA0u)
            {
                return false;
            }
            if (b0 == 0xEDu && b1 >= 0xA0u)
            {
                return false;
            }

            i += 3;
            continue;
        }

        if (b0 <= 0xF4u)
        {
            if ((i + 3) >= size)
            {
                return true;
            }

            const uint8_t b1 = data[i + 1];
            const uint8_t b2 = data[i + 2];
            const uint8_t b3 = data[i + 3];

            if ((b1 & 0xC0u) != 0x80u || (b2 & 0xC0u) != 0x80u || (b3 & 0xC0u) != 0x80u)
            {
                return false;
            }

            if (b0 == 0xF0u && b1 < 0x90u)
            {
                return false;
            }
            if (b0 == 0xF4u && b1 >= 0x90u)
            {
                return false;
            }

            i += 4;
            continue;
        }

        return false;
    }

    return true;
}

bool LooksLikeBinaryData(const uint8_t* data, size_t size) noexcept
{
    if (! data || size == 0)
    {
        return false;
    }

    constexpr size_t kMaxProbeBytes = 64u * 1024u;
    const size_t probeSize          = std::min(size, kMaxProbeBytes);

    size_t suspiciousControls = 0;
    for (size_t i = 0; i < probeSize; ++i)
    {
        const uint8_t b = data[i];
        if (b == 0)
        {
            return true;
        }

        if (b < 0x20u)
        {
            // Allow common whitespace/control used in text files.
            if (b == 0x09u || b == 0x0Au || b == 0x0Cu || b == 0x0Du)
            {
                continue;
            }
            suspiciousControls += 1;
            continue;
        }

        if (b == 0x7Fu)
        {
            suspiciousControls += 1;
            continue;
        }
    }

    const double ratio = static_cast<double>(suspiciousControls) / static_cast<double>(probeSize);
    return ratio > 0.25;
}

ViewerText::FileEncoding DisplayEncodingFileEncodingForSelection(UINT selection) noexcept
{
    switch (selection)
    {
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM: return ViewerText::FileEncoding::Utf8;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM: return ViewerText::FileEncoding::Utf16BE;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM: return ViewerText::FileEncoding::Utf16LE;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM: return ViewerText::FileEncoding::Utf32BE;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM: return ViewerText::FileEncoding::Utf32LE;
        default: return ViewerText::FileEncoding::Unknown;
    }
}

UINT CodePageForSelection(UINT selection) noexcept
{
    switch (selection)
    {
        case IDM_VIEWER_ENCODING_DISPLAY_ANSI: return CP_ACP;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF7: return 65000u;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM: return CP_UTF8;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM: return CP_ACP;
        default: break;
    }

    return selection;
}

uint64_t BytesToSkipForDisplayEncoding(UINT selection, ViewerText::FileEncoding encoding, uint64_t bomBytes) noexcept
{
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM && encoding == ViewerText::FileEncoding::Utf8 && bomBytes == 3)
    {
        return 3u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM && encoding == ViewerText::FileEncoding::Utf16LE && bomBytes == 2)
    {
        return 2u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM && encoding == ViewerText::FileEncoding::Utf16BE && bomBytes == 2)
    {
        return 2u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM && encoding == ViewerText::FileEncoding::Utf32LE && bomBytes == 4)
    {
        return 4u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM && encoding == ViewerText::FileEncoding::Utf32BE && bomBytes == 4)
    {
        return 4u;
    }

    return 0;
}

uint64_t TextStreamChunkBytes(uint32_t textBufferMiB, ViewerText::FileEncoding displayEncoding) noexcept
{
    uint64_t bytes = static_cast<uint64_t>(textBufferMiB) * 1024u * 1024u;
    bytes          = std::clamp<uint64_t>(bytes, 256u * 1024u, 256u * 1024u * 1024u);

    if (displayEncoding == ViewerText::FileEncoding::Utf16LE || displayEncoding == ViewerText::FileEncoding::Utf16BE)
    {
        bytes &= ~static_cast<uint64_t>(1);
        bytes = std::max<uint64_t>(bytes, 2u);
    }
    else if (displayEncoding == ViewerText::FileEncoding::Utf32LE || displayEncoding == ViewerText::FileEncoding::Utf32BE)
    {
        bytes &= ~static_cast<uint64_t>(3);
        bytes = std::max<uint64_t>(bytes, 4u);
    }

    return bytes;
}

size_t Utf8IncompleteTailSize(const uint8_t* data, size_t size) noexcept
{
    if (! data || size == 0)
    {
        return 0;
    }

    size_t start = size;
    for (size_t i = size; i > 0; --i)
    {
        const uint8_t b = data[i - 1];
        if ((b & 0xC0u) != 0x80u)
        {
            start = i - 1;
            break;
        }
    }

    if (start >= size)
    {
        return 0;
    }

    const uint8_t lead = data[start];
    size_t expected    = 1;
    if (lead <= 0x7Fu)
    {
        expected = 1;
    }
    else if (lead >= 0xC2u && lead <= 0xDFu)
    {
        expected = 2;
    }
    else if (lead >= 0xE0u && lead <= 0xEFu)
    {
        expected = 3;
    }
    else if (lead >= 0xF0u && lead <= 0xF4u)
    {
        expected = 4;
    }

    const size_t available = size - start;
    if (expected > 1 && available < expected)
    {
        return available;
    }

    return 0;
}

void BuildTextLineIndex(std::wstring_view text, std::vector<uint32_t>& outLineStarts, std::vector<uint32_t>& outLineEnds, uint32_t& outMaxLineLength) noexcept
{
    outLineStarts.clear();
    outLineEnds.clear();
    outMaxLineLength = 0;

    const size_t size = text.size();
    size_t start      = 0;

    for (;;)
    {
        size_t pos = start;
        while (pos < size)
        {
            const wchar_t ch = text[pos];
            if (ch == L'\n' || ch == L'\r')
            {
                break;
            }
            pos += 1;
        }

        const uint32_t start32 =
            start > static_cast<size_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max() : static_cast<uint32_t>(start);
        const uint32_t end32 =
            pos > static_cast<size_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max() : static_cast<uint32_t>(pos);

        outLineStarts.push_back(start32);
        outLineEnds.push_back(end32);

        if (end32 >= start32)
        {
            outMaxLineLength = std::max(outMaxLineLength, end32 - start32);
        }

        if (pos >= size)
        {
            break;
        }

        if (text[pos] == L'\r' && (pos + 1) < size && text[pos + 1] == L'\n')
        {
            start = pos + 2;
        }
        else
        {
            start = pos + 1;
        }

        if (start > size)
        {
            start = size;
        }
    }

    if (outLineStarts.empty())
    {
        outLineStarts.push_back(0);
        outLineEnds.push_back(0);
    }
}

const char* DiffDefaultLayoutToConfigString(ViewerText::DiffDefaultLayout value) noexcept
{
    return value == ViewerText::DiffDefaultLayout::Inline ? "inline" : "sideBySide";
}

const char* DiffContextModeToConfigString(ViewerText::DiffContextMode value) noexcept
{
    return value == ViewerText::DiffContextMode::FullFileWhenAvailable ? "fullFileWhenAvailable" : "hunksOnly";
}

const char* DiffAutoOpenModeToConfigString(ViewerText::DiffAutoOpenMode value) noexcept
{
    return value == ViewerText::DiffAutoOpenMode::RawText ? "rawText" : "parsed";
}

ViewerText::DiffDefaultLayout ParseDiffDefaultLayout(std::string_view value) noexcept
{
    return value == "inline" ? ViewerText::DiffDefaultLayout::Inline : ViewerText::DiffDefaultLayout::SideBySide;
}

ViewerText::DiffContextMode ParseDiffContextMode(std::string_view value) noexcept
{
    return value == "fullFileWhenAvailable" ? ViewerText::DiffContextMode::FullFileWhenAvailable : ViewerText::DiffContextMode::HunksOnly;
}

ViewerText::DiffAutoOpenMode ParseDiffAutoOpenMode(std::string_view value) noexcept
{
    return value == "rawText" ? ViewerText::DiffAutoOpenMode::RawText : ViewerText::DiffAutoOpenMode::Parsed;
}

bool HasDiffLikeExtension(const std::filesystem::path& path) noexcept
{
    std::wstring extension = path.extension().wstring();
    std::transform(extension.begin(), extension.end(), extension.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(std::towlower(ch)); });
    return extension == L".diff" || extension == L".patch" || extension == L".rej";
}

bool LooksLikeUnifiedDiffText(std::wstring_view text) noexcept
{
    size_t inspectedLines = 0u;
    size_t start          = 0u;
    while (start <= text.size() && inspectedLines < 32u)
    {
        size_t end = start;
        while (end < text.size() && text[end] != L'\r' && text[end] != L'\n')
        {
            end += 1u;
        }

        const std::wstring_view line = text.substr(start, end - start);
        if (line.starts_with(L"diff --git ") || line.starts_with(L"@@ -") || line.starts_with(L"--- ") || line.starts_with(L"+++ "))
        {
            return true;
        }

        if (end >= text.size())
        {
            break;
        }

        if (text[end] == L'\r' && (end + 1u) < text.size() && text[end + 1u] == L'\n')
        {
            start = end + 2u;
        }
        else
        {
            start = end + 1u;
        }
        inspectedLines += 1u;
    }

    return false;
}

void AppendLine(std::wstring& target, std::wstring_view line)
{
    target.append(line);
    target.push_back(L'\n');
}

size_t DecimalDigits(uint32_t value) noexcept
{
    size_t digits = 1u;
    while (value >= 10u)
    {
        value /= 10u;
        digits += 1u;
    }
    return digits;
}

std::wstring TrimDiffPathLabel(std::wstring_view text)
{
    while (! text.empty() && std::iswspace(text.front()))
    {
        text.remove_prefix(1u);
    }

    const size_t tabPos = text.find(L'\t');
    if (tabPos != std::wstring_view::npos)
    {
        text = text.substr(0, tabPos);
    }

    while (! text.empty() && std::iswspace(text.back()))
    {
        text.remove_suffix(1u);
    }

    if (text.size() >= 2u && text.front() == L'"' && text.back() == L'"')
    {
        text.remove_prefix(1u);
        text.remove_suffix(1u);
    }

    return std::wstring(text);
}

std::wstring StripGitDiffPrefix(std::wstring_view path)
{
    if (path.size() > 2u && (path[0] == L'a' || path[0] == L'b') && path[1] == L'/')
    {
        return std::wstring(path.substr(2u));
    }

    return std::wstring(path);
}

std::wstring FormatDiffSummaryPath(std::wstring_view path)
{
    std::wstring normalized = StripGitDiffPrefix(TrimDiffPathLabel(path));
    if (normalized.empty())
    {
        normalized = L"(unknown)";
    }
    return normalized;
}

std::wstring BuildDiffSectionNavigationLabel(const ParsedDiffFileSection& file, size_t fileIndex)
{
    return BuildDiffSectionNavigationLabel(file.leftDisplayPath, file.rightDisplayPath, fileIndex);
}

std::wstring BuildDiffSectionNavigationLabel(std::wstring_view leftPathLabel, std::wstring_view rightPathLabel, size_t fileIndex)
{
    const std::wstring leftPath  = FormatDiffSummaryPath(leftPathLabel);
    const std::wstring rightPath = FormatDiffSummaryPath(rightPathLabel);
    if (leftPathLabel == L"/dev/null")
    {
        return std::format(L"add: {}", rightPath);
    }
    if (rightPathLabel == L"/dev/null")
    {
        return std::format(L"delete: {}", leftPath);
    }
    if (! leftPath.empty() && ! rightPath.empty() && leftPath != rightPath)
    {
        return std::format(L"{} -> {}", leftPath, rightPath);
    }
    if (! rightPath.empty())
    {
        return rightPath;
    }
    if (! leftPath.empty())
    {
        return leftPath;
    }

    return std::format(L"section {}", fileIndex + 1u);
}

std::wstring BuildDiffHunkNavigationLabel(const ParsedDiffFileSection& file, size_t fileIndex, const ParsedDiffHunk& hunk, size_t hunkIndexInFile)
{
    return std::format(L"{} / hunk {} {}", BuildDiffSectionNavigationLabel(file, fileIndex), hunkIndexInFile + 1u, hunk.header);
}

std::wstring BuildHiddenContextBannerLabel(uint32_t oldHiddenLineCount, uint32_t newHiddenLineCount)
{
    const auto lineLabel = [](uint32_t count) noexcept -> const wchar_t* { return count == 1u ? L"line" : L"lines"; };

    if (oldHiddenLineCount == 0u)
    {
        return std::format(L"Show {} hidden new {}", newHiddenLineCount, lineLabel(newHiddenLineCount));
    }
    if (newHiddenLineCount == 0u)
    {
        return std::format(L"Show {} hidden old {}", oldHiddenLineCount, lineLabel(oldHiddenLineCount));
    }
    if (oldHiddenLineCount == newHiddenLineCount)
    {
        return std::format(L"Show {} hidden {}", oldHiddenLineCount, lineLabel(oldHiddenLineCount));
    }

    return std::format(
        L"Show {} hidden old {} and {} hidden new {}", oldHiddenLineCount, lineLabel(oldHiddenLineCount), newHiddenLineCount, lineLabel(newHiddenLineCount));
}

bool ParseUnifiedRange(std::wstring_view text, uint32_t& startOut, uint32_t& countOut, size_t& consumedChars) noexcept
{
    consumedChars = 0u;
    if (text.empty() || ! std::iswdigit(text.front()))
    {
        return false;
    }

    uint64_t start = 0u;
    size_t pos     = 0u;
    while (pos < text.size() && std::iswdigit(text[pos]))
    {
        start = (start * 10u) + static_cast<uint64_t>(text[pos] - L'0');
        pos += 1u;
    }

    uint64_t count = 1u;
    if (pos < text.size() && text[pos] == L',')
    {
        pos += 1u;
        if (pos >= text.size() || ! std::iswdigit(text[pos]))
        {
            return false;
        }

        count = 0u;
        while (pos < text.size() && std::iswdigit(text[pos]))
        {
            count = (count * 10u) + static_cast<uint64_t>(text[pos] - L'0');
            pos += 1u;
        }
    }

    startOut      = start > static_cast<uint64_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max() : static_cast<uint32_t>(start);
    countOut      = count > static_cast<uint64_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max() : static_cast<uint32_t>(count);
    consumedChars = pos;
    return true;
}

bool ParseUnifiedHunkHeader(std::wstring_view line, uint32_t& oldStartOut, uint32_t& newStartOut) noexcept
{
    if (! line.starts_with(L"@@ -"))
    {
        return false;
    }

    size_t pos        = 4u;
    uint32_t oldCount = 0u;
    uint32_t newCount = 0u;
    size_t consumed   = 0u;
    if (! ParseUnifiedRange(line.substr(pos), oldStartOut, oldCount, consumed))
    {
        return false;
    }
    pos += consumed;

    if (pos >= line.size() || line[pos] != L' ')
    {
        return false;
    }
    pos += 1u;

    if (pos >= line.size() || line[pos] != L'+')
    {
        return false;
    }
    pos += 1u;

    if (! ParseUnifiedRange(line.substr(pos), newStartOut, newCount, consumed))
    {
        return false;
    }
    pos += consumed;

    return line.find(L"@@", pos) != std::wstring_view::npos;
}

std::wstring DecodeBytesToWide(std::span<const uint8_t> bytes, ViewerText::FileEncoding encoding, UINT codePage, HRESULT& hrOut) noexcept
{
    hrOut = S_OK;
    std::wstring text;

    if (bytes.empty())
    {
        return text;
    }

    if (encoding == ViewerText::FileEncoding::Utf16LE || encoding == ViewerText::FileEncoding::Utf16BE)
    {
        if ((bytes.size() % 2u) != 0u)
        {
            hrOut = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return {};
        }

        text.resize(bytes.size() / 2u);
        memcpy(text.data(), bytes.data(), bytes.size_bytes());
        if (encoding == ViewerText::FileEncoding::Utf16BE)
        {
            for (wchar_t& ch : text)
            {
                const uint16_t value = static_cast<uint16_t>(ch);
                ch                   = static_cast<wchar_t>((value >> 8) | (value << 8));
            }
        }
        return text;
    }

    if (encoding == ViewerText::FileEncoding::Utf32LE || encoding == ViewerText::FileEncoding::Utf32BE)
    {
        if ((bytes.size() % 4u) != 0u)
        {
            hrOut = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return {};
        }

        const bool bigEndian = (encoding == ViewerText::FileEncoding::Utf32BE);
        for (size_t i = 0; i + 3u < bytes.size(); i += 4u)
        {
            uint32_t cp = 0u;
            if (bigEndian)
            {
                cp = (static_cast<uint32_t>(bytes[i]) << 24) | (static_cast<uint32_t>(bytes[i + 1u]) << 16) | (static_cast<uint32_t>(bytes[i + 2u]) << 8) |
                     static_cast<uint32_t>(bytes[i + 3u]);
            }
            else
            {
                cp = static_cast<uint32_t>(bytes[i]) | (static_cast<uint32_t>(bytes[i + 1u]) << 8) | (static_cast<uint32_t>(bytes[i + 2u]) << 16) |
                     (static_cast<uint32_t>(bytes[i + 3u]) << 24);
            }

            if (cp <= 0xFFFFu)
            {
                text.push_back(static_cast<wchar_t>(cp));
            }
            else if (cp <= 0x10FFFFu)
            {
                const uint32_t value = cp - 0x10000u;
                text.push_back(static_cast<wchar_t>(0xD800u + (value >> 10)));
                text.push_back(static_cast<wchar_t>(0xDC00u + (value & 0x3FFu)));
            }
            else
            {
                text.push_back(static_cast<wchar_t>(0xFFFDu));
            }
        }

        return text;
    }

    if (bytes.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        hrOut = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        return {};
    }

    const int srcLen       = static_cast<int>(bytes.size());
    const int requiredWide = MultiByteToWideChar(codePage, 0, reinterpret_cast<LPCCH>(bytes.data()), srcLen, nullptr, 0);
    if (requiredWide <= 0)
    {
        const DWORD lastError = GetLastError();
        hrOut                 = HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_INVALID_DATA);
        return {};
    }

    text.resize(static_cast<size_t>(requiredWide));
    const int written = MultiByteToWideChar(codePage, 0, reinterpret_cast<LPCCH>(bytes.data()), srcLen, text.data(), requiredWide);
    if (written <= 0)
    {
        const DWORD lastError = GetLastError();
        hrOut                 = HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_INVALID_DATA);
        return {};
    }

    text.resize(static_cast<size_t>(written));
    return text;
}

std::wstring DecodeBytesToWide(const std::vector<uint8_t>& bytes, ViewerText::FileEncoding encoding, UINT codePage, HRESULT& hrOut) noexcept
{
    return DecodeBytesToWide(std::span<const uint8_t>(bytes.data(), bytes.size()), encoding, codePage, hrOut);
}

bool BuildBoundedStreamedDiffSectionIndex(IFileReader* reader,
                                          uint64_t fileSize,
                                          uint64_t streamSkipBytes,
                                          ViewerText::FileEncoding encoding,
                                          UINT codePage,
                                          std::vector<ViewerText::StreamedDiffSectionEntry>& outSections) noexcept
{
    outSections.clear();

    if (! reader || fileSize <= streamSkipBytes)
    {
        return false;
    }

    const uint64_t availableBytes = fileSize - streamSkipBytes;
    if (availableBytes == 0u || availableBytes > kMaxBoundedStreamedDiffIndexBytes ||
        availableBytes > static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
    {
        return false;
    }

    if (streamSkipBytes > static_cast<uint64_t>(std::numeric_limits<__int64>::max()))
    {
        return false;
    }

    uint64_t newPosition = 0u;
    const HRESULT seekHr = reader->Seek(static_cast<__int64>(streamSkipBytes), FILE_BEGIN, &newPosition);
    if (FAILED(seekHr))
    {
        return false;
    }

    std::vector<uint8_t> bytes(static_cast<size_t>(availableBytes));
    size_t totalRead = 0u;
    while (totalRead < bytes.size())
    {
        const unsigned long want = static_cast<unsigned long>(std::min<size_t>(bytes.size() - totalRead, 256u * 1024u));
        unsigned long read       = 0u;
        const HRESULT readHr     = reader->Read(bytes.data() + totalRead, want, &read);
        if (FAILED(readHr))
        {
            return false;
        }
        if (read == 0u)
        {
            break;
        }

        totalRead += static_cast<size_t>(read);
    }
    bytes.resize(totalRead);
    if (bytes.empty())
    {
        return false;
    }

    struct IndexedSectionState
    {
        uint64_t startOffset = 0u;
        std::wstring leftPath;
        std::wstring rightPath;
        bool sawHunk = false;
    };

    std::vector<IndexedSectionState> indexedSections;
    indexedSections.reserve(16u);

    IndexedSectionState* currentSection = nullptr;
    auto startSection                   = [&](uint64_t startOffset) noexcept -> IndexedSectionState&
    {
        indexedSections.push_back(IndexedSectionState{startOffset});
        currentSection = &indexedSections.back();
        return *currentSection;
    };
    auto ensureSection = [&](uint64_t startOffset) noexcept -> IndexedSectionState&
    {
        if (! currentSection)
        {
            return startSection(startOffset);
        }
        return *currentSection;
    };

    auto processLine = [&](uint64_t lineStartOffset, std::wstring_view line) noexcept -> bool
    {
        if (line.empty() && ! currentSection)
        {
            return true;
        }

        const bool startsDiffGit      = line.starts_with(L"diff --git ");
        const bool startsIndex        = line.starts_with(L"Index: ");
        const bool startsLegacyHeader = line.starts_with(L"--- ") && currentSection && currentSection->sawHunk;
        if (startsDiffGit || startsIndex || startsLegacyHeader)
        {
            startSection(lineStartOffset);
            if (indexedSections.size() >= kMaxBoundedStreamedDiffSections)
            {
                return false;
            }
        }

        if (line.starts_with(L"@@ -"))
        {
            ensureSection(lineStartOffset).sawHunk = true;
        }

        if (line.starts_with(L"--- "))
        {
            ensureSection(lineStartOffset).leftPath = TrimDiffPathLabel(line.substr(4u));
        }
        else if (line.starts_with(L"+++ "))
        {
            ensureSection(lineStartOffset).rightPath = TrimDiffPathLabel(line.substr(4u));
        }

        return indexedSections.size() < kMaxBoundedStreamedDiffSections;
    };

    auto decodeAndProcessLine = [&](uint64_t lineStartOffset, size_t startIndex, size_t endIndex) noexcept -> bool
    {
        if (endIndex < startIndex || endIndex > bytes.size())
        {
            return false;
        }

        HRESULT decodeHr = S_OK;
        const std::wstring line =
            DecodeBytesToWide(std::span<const uint8_t>(bytes.data() + static_cast<ptrdiff_t>(startIndex), endIndex - startIndex), encoding, codePage, decodeHr);
        return SUCCEEDED(decodeHr) && processLine(lineStartOffset, line);
    };

    if (encoding == ViewerText::FileEncoding::Utf16LE || encoding == ViewerText::FileEncoding::Utf16BE)
    {
        if ((bytes.size() % 2u) != 0u)
        {
            bytes.pop_back();
        }

        const auto readUnit = [&](size_t index) noexcept -> uint16_t
        {
            const uint16_t lo = static_cast<uint16_t>(bytes[index]);
            const uint16_t hi = static_cast<uint16_t>(bytes[index + 1u]);
            if (encoding == ViewerText::FileEncoding::Utf16BE)
            {
                return static_cast<uint16_t>((lo << 8) | hi);
            }
            return static_cast<uint16_t>(lo | (hi << 8));
        };

        size_t lineStart = 0u;
        while (lineStart + 1u < bytes.size())
        {
            size_t lineEnd = lineStart;
            while (lineEnd + 1u < bytes.size())
            {
                const uint16_t ch = readUnit(lineEnd);
                if (ch == L'\r' || ch == L'\n')
                {
                    break;
                }
                lineEnd += 2u;
            }

            if (! decodeAndProcessLine(streamSkipBytes + static_cast<uint64_t>(lineStart), lineStart, lineEnd))
            {
                break;
            }

            if (lineEnd + 1u >= bytes.size())
            {
                break;
            }

            size_t nextStart = lineEnd + 2u;
            if (readUnit(lineEnd) == L'\r' && nextStart + 1u < bytes.size() && readUnit(nextStart) == L'\n')
            {
                nextStart += 2u;
            }
            lineStart = nextStart;
        }
    }
    else if (encoding == ViewerText::FileEncoding::Utf32LE || encoding == ViewerText::FileEncoding::Utf32BE)
    {
        const size_t remainder = bytes.size() % 4u;
        if (remainder != 0u)
        {
            bytes.resize(bytes.size() - remainder);
        }

        const auto readUnit = [&](size_t index) noexcept -> uint32_t
        {
            const uint32_t b0 = static_cast<uint32_t>(bytes[index]);
            const uint32_t b1 = static_cast<uint32_t>(bytes[index + 1u]);
            const uint32_t b2 = static_cast<uint32_t>(bytes[index + 2u]);
            const uint32_t b3 = static_cast<uint32_t>(bytes[index + 3u]);
            if (encoding == ViewerText::FileEncoding::Utf32BE)
            {
                return (b0 << 24) | (b1 << 16) | (b2 << 8) | b3;
            }
            return b0 | (b1 << 8) | (b2 << 16) | (b3 << 24);
        };

        size_t lineStart = 0u;
        while (lineStart + 3u < bytes.size())
        {
            size_t lineEnd = lineStart;
            while (lineEnd + 3u < bytes.size())
            {
                const uint32_t ch = readUnit(lineEnd);
                if (ch == L'\r' || ch == L'\n')
                {
                    break;
                }
                lineEnd += 4u;
            }

            if (! decodeAndProcessLine(streamSkipBytes + static_cast<uint64_t>(lineStart), lineStart, lineEnd))
            {
                break;
            }

            if (lineEnd + 3u >= bytes.size())
            {
                break;
            }

            size_t nextStart = lineEnd + 4u;
            if (readUnit(lineEnd) == L'\r' && nextStart + 3u < bytes.size() && readUnit(nextStart) == L'\n')
            {
                nextStart += 4u;
            }
            lineStart = nextStart;
        }
    }
    else
    {
        size_t lineStart = 0u;
        while (lineStart < bytes.size())
        {
            size_t lineEnd = lineStart;
            while (lineEnd < bytes.size() && bytes[lineEnd] != 0x0Du && bytes[lineEnd] != 0x0Au)
            {
                lineEnd += 1u;
            }

            if (! decodeAndProcessLine(streamSkipBytes + static_cast<uint64_t>(lineStart), lineStart, lineEnd))
            {
                break;
            }

            if (lineEnd >= bytes.size())
            {
                break;
            }

            size_t nextStart = lineEnd + 1u;
            if (bytes[lineEnd] == 0x0Du && nextStart < bytes.size() && bytes[nextStart] == 0x0Au)
            {
                nextStart += 1u;
            }
            lineStart = nextStart;
        }
    }

    if (indexedSections.empty())
    {
        return false;
    }

    outSections.reserve(indexedSections.size());
    for (size_t i = 0; i < indexedSections.size(); ++i)
    {
        const IndexedSectionState& section = indexedSections[i];
        outSections.push_back(
            ViewerText::StreamedDiffSectionEntry{BuildDiffSectionNavigationLabel(section.leftPath, section.rightPath, i), section.startOffset});
    }

    return ! outSections.empty();
}

[[nodiscard]] bool TryFindReferencedDiffLineBreak(std::span<const uint8_t> bytes,
                                                  ViewerText::FileEncoding encoding,
                                                  size_t& contentBytesOut,
                                                  size_t& consumedBytesOut) noexcept
{
    if (bytes.empty())
    {
        return false;
    }

    switch (encoding)
    {
        case ViewerText::FileEncoding::Utf16LE:
        case ViewerText::FileEncoding::Utf16BE:
        {
            const size_t usable = bytes.size() - (bytes.size() % 2u);
            for (size_t i = 0u; i + 1u < usable; i += 2u)
            {
                const uint16_t codeUnit = (encoding == ViewerText::FileEncoding::Utf16LE)
                                              ? static_cast<uint16_t>(bytes[i] | (static_cast<uint16_t>(bytes[i + 1u]) << 8u))
                                              : static_cast<uint16_t>((static_cast<uint16_t>(bytes[i]) << 8u) | bytes[i + 1u]);
                if (codeUnit != L'\r' && codeUnit != L'\n')
                {
                    continue;
                }

                contentBytesOut  = i;
                consumedBytesOut = i + 2u;
                if (codeUnit == L'\r' && (i + 3u) < usable)
                {
                    const uint16_t nextCodeUnit = (encoding == ViewerText::FileEncoding::Utf16LE)
                                                      ? static_cast<uint16_t>(bytes[i + 2u] | (static_cast<uint16_t>(bytes[i + 3u]) << 8u))
                                                      : static_cast<uint16_t>((static_cast<uint16_t>(bytes[i + 2u]) << 8u) | bytes[i + 3u]);
                    if (nextCodeUnit == L'\n')
                    {
                        consumedBytesOut += 2u;
                    }
                }
                return true;
            }
            return false;
        }
        case ViewerText::FileEncoding::Utf32LE:
        case ViewerText::FileEncoding::Utf32BE:
        {
            const size_t usable = bytes.size() - (bytes.size() % 4u);
            for (size_t i = 0u; i + 3u < usable; i += 4u)
            {
                const uint32_t codePoint = (encoding == ViewerText::FileEncoding::Utf32LE)
                                               ? static_cast<uint32_t>(bytes[i]) | (static_cast<uint32_t>(bytes[i + 1u]) << 8u) |
                                                     (static_cast<uint32_t>(bytes[i + 2u]) << 16u) | (static_cast<uint32_t>(bytes[i + 3u]) << 24u)
                                               : (static_cast<uint32_t>(bytes[i]) << 24u) | (static_cast<uint32_t>(bytes[i + 1u]) << 16u) |
                                                     (static_cast<uint32_t>(bytes[i + 2u]) << 8u) | static_cast<uint32_t>(bytes[i + 3u]);
                if (codePoint != L'\r' && codePoint != L'\n')
                {
                    continue;
                }

                contentBytesOut  = i;
                consumedBytesOut = i + 4u;
                if (codePoint == L'\r' && (i + 7u) < usable)
                {
                    const uint32_t nextCodePoint = (encoding == ViewerText::FileEncoding::Utf32LE)
                                                       ? static_cast<uint32_t>(bytes[i + 4u]) | (static_cast<uint32_t>(bytes[i + 5u]) << 8u) |
                                                             (static_cast<uint32_t>(bytes[i + 6u]) << 16u) | (static_cast<uint32_t>(bytes[i + 7u]) << 24u)
                                                       : (static_cast<uint32_t>(bytes[i + 4u]) << 24u) | (static_cast<uint32_t>(bytes[i + 5u]) << 16u) |
                                                             (static_cast<uint32_t>(bytes[i + 6u]) << 8u) | static_cast<uint32_t>(bytes[i + 7u]);
                    if (nextCodePoint == L'\n')
                    {
                        consumedBytesOut += 4u;
                    }
                }
                return true;
            }
            return false;
        }
        case ViewerText::FileEncoding::Utf8: break;
        case ViewerText::FileEncoding::Unknown:
        default: break;
    }

    for (size_t i = 0u; i < bytes.size(); ++i)
    {
        const uint8_t value = bytes[i];
        if (value != '\r' && value != '\n')
        {
            continue;
        }

        contentBytesOut  = i;
        consumedBytesOut = i + 1u;
        if (value == '\r' && (i + 1u) < bytes.size() && bytes[i + 1u] == '\n')
        {
            consumedBytesOut += 1u;
        }
        return true;
    }

    return false;
}

[[nodiscard]] bool DecodeReferencedDiffLineBytes(ResolvedDiffTextFile& file, std::span<const uint8_t> bytes) noexcept
{
    HRESULT decodeHr  = S_OK;
    std::wstring line = DecodeBytesToWide(bytes, file.encoding, file.codePage, decodeHr);
    if (FAILED(decodeHr))
    {
        file.available        = false;
        file.hadDecodeFailure = true;
        file.reason           = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_BINARY);
        file.reader.reset();
        file.pendingBytes.clear();
        return false;
    }

    file.lines.push_back(std::move(line));
    return true;
}

[[nodiscard]] bool ProcessReferencedDiffPendingBytes(ResolvedDiffTextFile& file, bool finalChunk) noexcept
{
    if (! file.available)
    {
        return false;
    }

    size_t contentBytes  = 0u;
    size_t consumedBytes = 0u;
    while (TryFindReferencedDiffLineBreak(file.pendingBytes, file.encoding, contentBytes, consumedBytes))
    {
        if (! DecodeReferencedDiffLineBytes(file, std::span<const uint8_t>(file.pendingBytes.data(), contentBytes)))
        {
            return false;
        }

        file.pendingBytes.erase(file.pendingBytes.begin(), file.pendingBytes.begin() + static_cast<ptrdiff_t>(consumedBytes));
    }

    if (! finalChunk)
    {
        return true;
    }

    if (! file.pendingBytes.empty())
    {
        const bool hasIncompleteCodeUnit = ((file.encoding == ViewerText::FileEncoding::Utf16LE || file.encoding == ViewerText::FileEncoding::Utf16BE) &&
                                            ((file.pendingBytes.size() % 2u) != 0u)) ||
                                           ((file.encoding == ViewerText::FileEncoding::Utf32LE || file.encoding == ViewerText::FileEncoding::Utf32BE) &&
                                            ((file.pendingBytes.size() % 4u) != 0u));
        if (hasIncompleteCodeUnit)
        {
            file.available        = false;
            file.hadDecodeFailure = true;
            file.reason           = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_BINARY);
            file.reader.reset();
            file.pendingBytes.clear();
            return false;
        }

        if (! DecodeReferencedDiffLineBytes(file, file.pendingBytes))
        {
            return false;
        }
        file.pendingBytes.clear();
    }

    file.loadComplete = true;
    file.reader.reset();
    return true;
}

std::shared_ptr<ResolvedDiffTextFile> OpenReferencedDiffTextFile(IFileSystemIO* fileIo, const std::filesystem::path& path) noexcept
{
    auto result = std::make_shared<ResolvedDiffTextFile>();
    if (! result)
    {
        return {};
    }

    if (! fileIo || path.empty())
    {
        result->reason = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_MISSING_FILE);
        return result;
    }

    wil::com_ptr<IFileReader> reader;
    const HRESULT openHr = fileIo->CreateFileReader(path.c_str(), reader.put());
    if (FAILED(openHr) || ! reader)
    {
        result->reason = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_MISSING_FILE);
        return result;
    }

    result->existsInCurrentFileSystem = true;
    result->resolvedPath              = path;

    uint64_t fileSize    = 0u;
    const HRESULT sizeHr = reader->GetSize(&fileSize);
    if (FAILED(sizeHr))
    {
        result->reason = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_MISSING_FILE);
        return result;
    }

    if (fileSize > kMaxReferencedDiffFileBytes || fileSize > static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
    {
        result->reason = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_TOO_LARGE);
        return result;
    }

    const size_t probeSize = static_cast<size_t>(std::min<uint64_t>(fileSize, kReferencedDiffProbeBytes));
    std::vector<uint8_t> probe(probeSize);
    if (! probe.empty())
    {
        unsigned long probeRead = 0u;
        const HRESULT readHr    = reader->Read(probe.data(), static_cast<unsigned long>(probe.size()), &probeRead);
        if (FAILED(readHr))
        {
            result->reason = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_MISSING_FILE);
            return result;
        }
        probe.resize(static_cast<size_t>(probeRead));
        result->bytesRead = probeRead;
    }

    ViewerText::FileEncoding encoding = ViewerText::FileEncoding::Unknown;
    size_t skipBytes                  = 0u;
    if (probe.size() >= 4u && probe[0] == 0xFFu && probe[1] == 0xFEu && probe[2] == 0x00u && probe[3] == 0x00u)
    {
        encoding  = ViewerText::FileEncoding::Utf32LE;
        skipBytes = 4u;
    }
    else if (probe.size() >= 4u && probe[0] == 0x00u && probe[1] == 0x00u && probe[2] == 0xFEu && probe[3] == 0xFFu)
    {
        encoding  = ViewerText::FileEncoding::Utf32BE;
        skipBytes = 4u;
    }
    else if (probe.size() >= 3u && probe[0] == 0xEFu && probe[1] == 0xBBu && probe[2] == 0xBFu)
    {
        encoding  = ViewerText::FileEncoding::Utf8;
        skipBytes = 3u;
    }
    else if (probe.size() >= 2u && probe[0] == 0xFFu && probe[1] == 0xFEu)
    {
        encoding  = ViewerText::FileEncoding::Utf16LE;
        skipBytes = 2u;
    }
    else if (probe.size() >= 2u && probe[0] == 0xFEu && probe[1] == 0xFFu)
    {
        encoding  = ViewerText::FileEncoding::Utf16BE;
        skipBytes = 2u;
    }

    const size_t payloadStart = std::min(skipBytes, probe.size());
    const uint8_t* textBytes  = payloadStart < probe.size() ? (probe.data() + payloadStart) : nullptr;
    const size_t textSize     = probe.size() - payloadStart;
    if (encoding == ViewerText::FileEncoding::Unknown && textSize > 0u && LooksLikeBinaryData(textBytes, textSize))
    {
        result->reason = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_BINARY);
        return result;
    }

    UINT codePage = GetACP();
    if (encoding == ViewerText::FileEncoding::Utf8 || (encoding == ViewerText::FileEncoding::Unknown && textSize > 0u && IsValidUtf8(textBytes, textSize)))
    {
        codePage = CP_UTF8;
    }

    result->available = true;
    result->reason.clear();
    result->reader         = std::move(reader);
    result->encoding       = encoding;
    result->codePage       = codePage;
    result->fileSize       = fileSize;
    result->bomBytes       = skipBytes;
    result->nextReadOffset = probe.size();
    if (payloadStart < probe.size())
    {
        result->pendingBytes.insert(result->pendingBytes.end(), probe.begin() + static_cast<ptrdiff_t>(payloadStart), probe.end());
    }
    if (result->nextReadOffset >= result->fileSize && result->pendingBytes.empty())
    {
        result->loadComplete = true;
        result->reader.reset();
    }

    return result;
}

bool EnsureReferencedDiffLinesLoaded(ResolvedDiffTextFile& file, uint32_t lineNumberInclusive) noexcept
{
    if (! file.available)
    {
        return false;
    }

    if (lineNumberInclusive == 0u || file.lines.size() >= static_cast<size_t>(lineNumberInclusive) || file.loadComplete)
    {
        return true;
    }

    while (file.lines.size() < static_cast<size_t>(lineNumberInclusive) && ! file.loadComplete)
    {
        if (! ProcessReferencedDiffPendingBytes(file, false))
        {
            return false;
        }
        if (file.lines.size() >= static_cast<size_t>(lineNumberInclusive) || file.loadComplete)
        {
            break;
        }

        if (! file.reader)
        {
            file.loadComplete = true;
            break;
        }

        const uint64_t remaining = (file.fileSize > file.nextReadOffset) ? (file.fileSize - file.nextReadOffset) : 0u;
        if (remaining == 0u)
        {
            return ProcessReferencedDiffPendingBytes(file, true);
        }

        const size_t wantSize     = static_cast<size_t>(std::min<uint64_t>(remaining, kReferencedDiffChunkBytes));
        const size_t previousSize = file.pendingBytes.size();
        file.pendingBytes.resize(previousSize + wantSize);

        unsigned long read   = 0u;
        const HRESULT readHr = file.reader->Read(file.pendingBytes.data() + previousSize, static_cast<unsigned long>(wantSize), &read);
        if (FAILED(readHr))
        {
            file.available        = false;
            file.hadDecodeFailure = true;
            file.reason           = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_MISSING_FILE);
            file.pendingBytes.resize(previousSize);
            file.reader.reset();
            return false;
        }

        file.pendingBytes.resize(previousSize + static_cast<size_t>(read));
        file.nextReadOffset += read;
        file.bytesRead += read;
        if (read == 0u)
        {
            return ProcessReferencedDiffPendingBytes(file, true);
        }
    }

    if (file.lines.size() < static_cast<size_t>(lineNumberInclusive) && ! file.loadComplete && file.nextReadOffset >= file.fileSize)
    {
        return ProcessReferencedDiffPendingBytes(file, true);
    }

    return true;
}

std::wstring_view TryGetReferencedDiffLine(const ResolvedDiffTextFile& file, uint32_t lineNumber) noexcept
{
    if (lineNumber == 0u)
    {
        return {};
    }

    const size_t index = static_cast<size_t>(lineNumber - 1u);
    if (index >= file.lines.size())
    {
        return {};
    }

    return file.lines[index];
}

std::shared_ptr<ResolvedDiffTextFile> ResolveDiffReference(IFileSystemIO* fileIo,
                                                           const std::filesystem::path& diffPath,
                                                           std::wstring_view label,
                                                           UINT missingFileMessageId,
                                                           std::unordered_map<std::wstring, std::shared_ptr<ResolvedDiffTextFile>>& cache) noexcept
{
    const std::wstring trimmed = TrimDiffPathLabel(label);
    if (trimmed.empty() || trimmed == L"/dev/null")
    {
        auto result = std::make_shared<ResolvedDiffTextFile>();
        if (result)
        {
            result->reason = LoadStringResource(g_hInstance, missingFileMessageId);
        }
        return result;
    }

    std::vector<std::filesystem::path> candidates;
    const std::filesystem::path exactPath(trimmed);
    const std::wstring stripped = StripGitDiffPrefix(trimmed);
    const std::filesystem::path strippedPath(stripped);

    if (exactPath.is_absolute())
    {
        candidates.push_back(exactPath);
        if (strippedPath.is_absolute() && strippedPath != exactPath)
        {
            candidates.push_back(strippedPath);
        }
    }
    else
    {
        const std::filesystem::path baseDir = diffPath.parent_path();
        candidates.push_back((baseDir / exactPath).lexically_normal());
        if (strippedPath != exactPath)
        {
            candidates.push_back((baseDir / strippedPath).lexically_normal());
        }
    }

    std::shared_ptr<ResolvedDiffTextFile> bestFailure;
    for (const std::filesystem::path& candidate : candidates)
    {
        const std::wstring key = candidate.lexically_normal().wstring();
        auto [it, inserted]    = cache.try_emplace(key);
        if (inserted || ! it->second)
        {
            it->second = OpenReferencedDiffTextFile(fileIo, candidate);
        }

        if (it->second && it->second->available)
        {
            return it->second;
        }

        if (it->second && it->second->existsInCurrentFileSystem)
        {
            bestFailure = it->second;
        }
    }

    if (bestFailure && ! bestFailure->reason.empty())
    {
        return bestFailure;
    }

    auto missing = std::make_shared<ResolvedDiffTextFile>();
    if (missing)
    {
        missing->reason = LoadStringResource(g_hInstance, missingFileMessageId);
    }
    return missing;
}

bool ParseUnifiedDiffDocument(std::wstring_view text, ParsedDiffDocument& outDocument) noexcept
{
    outDocument.files.clear();

    ParsedDiffFileSection* currentFile = nullptr;
    ParsedDiffHunk* currentHunk        = nullptr;
    uint32_t oldCursor                 = 0u;
    uint32_t newCursor                 = 0u;

    auto ensureFileSection = [&]() noexcept -> ParsedDiffFileSection&
    {
        if (! currentFile)
        {
            outDocument.files.emplace_back();
            currentFile = &outDocument.files.back();
        }
        return *currentFile;
    };

    size_t start = 0u;
    while (start <= text.size())
    {
        size_t end = start;
        while (end < text.size() && text[end] != L'\r' && text[end] != L'\n')
        {
            end += 1u;
        }

        const std::wstring_view line = text.substr(start, end - start);

        const bool startsDiffGit       = line.starts_with(L"diff --git ");
        const bool startsIndex         = line.starts_with(L"Index: ");
        const bool startsLegacySection = line.starts_with(L"--- ") && currentFile && ! currentFile->hunks.empty() && currentHunk == nullptr;
        if (startsDiffGit || startsIndex || startsLegacySection)
        {
            outDocument.files.emplace_back();
            currentFile = &outDocument.files.back();
            currentHunk = nullptr;
            if (startsDiffGit || startsIndex)
            {
                currentFile->metadataLines.emplace_back(line);
            }
        }

        if (line.starts_with(L"@@ -"))
        {
            ParsedDiffFileSection& file = ensureFileSection();

            uint32_t oldStart = 0u;
            uint32_t newStart = 0u;
            if (! ParseUnifiedHunkHeader(line, oldStart, newStart))
            {
                return false;
            }

            file.hunks.emplace_back();
            currentHunk           = &file.hunks.back();
            currentHunk->header   = std::wstring(line);
            currentHunk->oldStart = oldStart;
            currentHunk->newStart = newStart;
            oldCursor             = oldStart;
            newCursor             = newStart;
        }
        else if (currentHunk && ! line.empty())
        {
            const wchar_t prefix = line.front();
            if (prefix == L' ' || prefix == L'+' || prefix == L'-' || prefix == L'\\')
            {
                ParsedDiffLine parsedLine{};
                parsedLine.text = std::wstring(prefix == L'\\' ? line : line.substr(1u));

                switch (prefix)
                {
                    case L' ':
                        parsedLine.kind       = ParsedDiffLineKind::Context;
                        parsedLine.oldLine    = oldCursor;
                        parsedLine.newLine    = newCursor;
                        parsedLine.hasOldLine = true;
                        parsedLine.hasNewLine = true;
                        oldCursor += 1u;
                        newCursor += 1u;
                        break;
                    case L'+':
                        parsedLine.kind       = ParsedDiffLineKind::Added;
                        parsedLine.newLine    = newCursor;
                        parsedLine.hasNewLine = true;
                        newCursor += 1u;
                        break;
                    case L'-':
                        parsedLine.kind       = ParsedDiffLineKind::Removed;
                        parsedLine.oldLine    = oldCursor;
                        parsedLine.hasOldLine = true;
                        oldCursor += 1u;
                        break;
                    case L'\\': parsedLine.kind = ParsedDiffLineKind::NoNewlineMarker; break;
                    default: break;
                }

                currentHunk->lines.push_back(std::move(parsedLine));
            }
            else
            {
                currentHunk = nullptr;
            }
        }

        if (! currentHunk)
        {
            ParsedDiffFileSection* fileForMetadata = currentFile;
            if (! fileForMetadata && (line.starts_with(L"--- ") || line.starts_with(L"+++ ")))
            {
                fileForMetadata = &ensureFileSection();
            }

            if (fileForMetadata)
            {
                if (line.starts_with(L"--- "))
                {
                    fileForMetadata->leftDisplayPath = TrimDiffPathLabel(line.substr(4u));
                    fileForMetadata->metadataLines.emplace_back(line);
                }
                else if (line.starts_with(L"+++ "))
                {
                    fileForMetadata->rightDisplayPath = TrimDiffPathLabel(line.substr(4u));
                    fileForMetadata->metadataLines.emplace_back(line);
                }
                else if (! line.empty() && ! line.starts_with(L"@@ -") && ! startsDiffGit && ! startsIndex)
                {
                    fileForMetadata->metadataLines.emplace_back(line);
                }
            }
        }

        if (end >= text.size())
        {
            break;
        }

        if (text[end] == L'\r' && (end + 1u) < text.size() && text[end + 1u] == L'\n')
        {
            start = end + 2u;
        }
        else
        {
            start = end + 1u;
        }
    }

    bool hasHunks = false;
    for (const ParsedDiffFileSection& file : outDocument.files)
    {
        if (! file.hunks.empty())
        {
            hasHunks = true;
            break;
        }
    }

    return ! outDocument.files.empty() && hasHunks;
}

DiffVariantBuildResult BuildInlineDiffText(const ParsedDiffDocument& document,
                                           ViewerText::DiffContextMode contextMode,
                                           const std::filesystem::path& diffPath,
                                           IFileSystemIO* fileIo,
                                           DiffReferenceCache* referenceCache,
                                           std::optional<size_t> hydratedSectionIndex,
                                           std::optional<std::pair<uint32_t, uint32_t>> hydratedLogicalRange) noexcept
{
    DiffVariantBuildResult result{};
    result.variant.fileSectionCount        = document.files.size();
    result.variant.referencedFilesResolved = (contextMode == ViewerText::DiffContextMode::FullFileWhenAvailable);
    if (hydratedLogicalRange.has_value())
    {
        result.variant.hydratedLogicalLineStart        = hydratedLogicalRange->first;
        result.variant.hydratedLogicalLineEndExclusive = hydratedLogicalRange->second;
    }

    std::unordered_map<std::wstring, std::shared_ptr<ResolvedDiffTextFile>> localCache;
    auto& cache                 = referenceCache ? referenceCache->files : localCache;
    uint32_t currentLogicalLine = 0u;
    const auto recordRowStyle   = [&](const ViewerText::DiffTextVariant::LogicalRowStyleEntry& style) noexcept
    {
        const auto countKind = [&](ViewerText::DiffTextVariant::SemanticRowKind kind) noexcept
        {
            switch (kind)
            {
                case ViewerText::DiffTextVariant::SemanticRowKind::Context: result.variant.contextRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::Added: result.variant.addedRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::Removed: result.variant.removedRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::FileHeader: result.variant.headerRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::HunkHeader:
                case ViewerText::DiffTextVariant::SemanticRowKind::HiddenContextBanner: result.variant.bannerRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::Placeholder: result.variant.placeholderRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::None: break;
            }
        };

        bool styled = false;
        if (style.fullRow != ViewerText::DiffTextVariant::SemanticRowKind::None)
        {
            countKind(style.fullRow);
            styled = true;
        }
        else
        {
            if (style.leftPane != ViewerText::DiffTextVariant::SemanticRowKind::None)
            {
                countKind(style.leftPane);
                styled = true;
            }
            if (style.rightPane != ViewerText::DiffTextVariant::SemanticRowKind::None)
            {
                countKind(style.rightPane);
                styled = true;
            }
        }

        if (styled)
        {
            result.variant.styledRowCount += 1u;
        }
    };
    const auto appendOutputLine =
        [&](std::wstring_view line,
            ViewerText::DiffTextVariant::LogicalRowStyleEntry style       = ViewerText::DiffTextVariant::LogicalRowStyleEntry{},
            ViewerText::DiffTextVariant::LogicalRowRenderEntry renderInfo = ViewerText::DiffTextVariant::LogicalRowRenderEntry{}) noexcept
    {
        const uint32_t lineStartIndex = static_cast<uint32_t>(std::min<size_t>(result.variant.text.size(), std::numeric_limits<uint32_t>::max()));
        if (renderInfo.fullMarkerIndex != std::numeric_limits<uint32_t>::max())
        {
            renderInfo.fullMarkerIndex += lineStartIndex;
        }
        if (renderInfo.leftMarkerIndex != std::numeric_limits<uint32_t>::max())
        {
            renderInfo.leftMarkerIndex += lineStartIndex;
        }
        if (renderInfo.rightMarkerIndex != std::numeric_limits<uint32_t>::max())
        {
            renderInfo.rightMarkerIndex += lineStartIndex;
        }
        AppendLine(result.variant.text, line);
        result.variant.logicalRowStyles.push_back(style);
        result.variant.logicalRowRenderInfo.push_back(renderInfo);
        result.variant.logicalRowPaneLayouts.push_back({});
        recordRowStyle(style);
        currentLogicalLine += 1u;
    };
    const auto appendPlaceholderBand = [&](ViewerText::DiffTextVariant::PlaceholderBandPlacement placement) noexcept
    {
        result.variant.placeholderBands.push_back(ViewerText::DiffTextVariant::PlaceholderBandEntry{currentLogicalLine, placement});
        result.variant.placeholderBandCount = result.variant.placeholderBands.size();
    };

    for (size_t fileIndex = 0u; fileIndex < document.files.size(); ++fileIndex)
    {
        const ParsedDiffFileSection& file = document.files[fileIndex];
        const bool hydrateExpandedContext = ShouldHydrateExpandedDiffSection(contextMode, hydratedSectionIndex, fileIndex);
        if (fileIndex != 0u)
        {
            appendOutputLine(L"");
        }

        result.variant.sectionNavigation.push_back(
            ViewerText::DiffTextVariant::SectionNavigationEntry{BuildDiffSectionNavigationLabel(file, fileIndex), currentLogicalLine});
        appendOutputLine(std::format(L"old path: {}", FormatDiffSummaryPath(file.leftDisplayPath)), {ViewerText::DiffTextVariant::SemanticRowKind::FileHeader});
        appendOutputLine(std::format(L"new path: {}", FormatDiffSummaryPath(file.rightDisplayPath)),
                         {ViewerText::DiffTextVariant::SemanticRowKind::FileHeader});
        for (const std::wstring& metadataLine : file.metadataLines)
        {
            appendOutputLine(metadataLine, {ViewerText::DiffTextVariant::SemanticRowKind::FileHeader});
        }

        std::shared_ptr<ResolvedDiffTextFile> leftFile;
        std::shared_ptr<ResolvedDiffTextFile> rightFile;
        if (hydrateExpandedContext)
        {
            const UINT leftMissingId =
                (file.leftDisplayPath == L"/dev/null") ? IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_ADDED_FILE : IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_MISSING_FILE;
            const UINT rightMissingId =
                (file.rightDisplayPath == L"/dev/null") ? IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_DELETED_FILE : IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_MISSING_FILE;
            leftFile  = ResolveDiffReference(fileIo, diffPath, file.leftDisplayPath, leftMissingId, cache);
            rightFile = ResolveDiffReference(fileIo, diffPath, file.rightDisplayPath, rightMissingId, cache);
        }

        uint32_t maxLineNumber = 1u;
        for (const ParsedDiffHunk& hunk : file.hunks)
        {
            for (const ParsedDiffLine& line : hunk.lines)
            {
                if (line.hasOldLine)
                {
                    maxLineNumber = std::max(maxLineNumber, line.oldLine);
                }
                if (line.hasNewLine)
                {
                    maxLineNumber = std::max(maxLineNumber, line.newLine);
                }
            }
        }

        if (hydrateExpandedContext && leftFile && leftFile->lines.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
        {
            maxLineNumber = std::max(maxLineNumber, static_cast<uint32_t>(leftFile->lines.size()));
        }
        if (hydrateExpandedContext && rightFile && rightFile->lines.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
        {
            maxLineNumber = std::max(maxLineNumber, static_cast<uint32_t>(rightFile->lines.size()));
        }

        const size_t digits = DecimalDigits(maxLineNumber);
        auto formatNumber   = [digits](bool hasValue, uint32_t value) -> std::wstring
        {
            if (! hasValue)
            {
                return std::wstring(digits, L' ');
            }
            return std::format(L"{:>{}}", value, digits);
        };

        auto appendInlineRow = [&](bool hasOld, uint32_t oldLine, bool hasNew, uint32_t newLine, wchar_t marker, std::wstring_view text) noexcept
        {
            ViewerText::DiffTextVariant::LogicalRowStyleEntry style{};
            if (marker == L'+')
            {
                style.fullRow = ViewerText::DiffTextVariant::SemanticRowKind::Added;
            }
            else if (marker == L'-')
            {
                style.fullRow = ViewerText::DiffTextVariant::SemanticRowKind::Removed;
            }
            else if (marker == L' ' && hasOld && hasNew)
            {
                style.fullRow = ViewerText::DiffTextVariant::SemanticRowKind::Context;
            }

            const std::wstring oldNumber = formatNumber(hasOld, oldLine);
            const std::wstring newNumber = formatNumber(hasNew, newLine);
            const std::wstring lineText  = std::format(L"{} {} {} {}", oldNumber, newNumber, marker, std::wstring(text));
            ViewerText::DiffTextVariant::LogicalRowRenderEntry renderInfo{};
            if (marker == L'+' || marker == L'-')
            {
                renderInfo.fullMarkerIndex = static_cast<uint32_t>(oldNumber.size() + 1u + newNumber.size() + 1u);
            }
            appendOutputLine(lineText, style, renderInfo);
        };

        auto appendPlaceholder = [&](std::wstring_view message) noexcept
        {
            result.variant.hasPlaceholderRows = true;
            appendPlaceholderBand(ViewerText::DiffTextVariant::PlaceholderBandPlacement::FullRow);
            appendOutputLine(std::format(L"{} {} ! {}", formatNumber(false, 0u), formatNumber(false, 0u), std::wstring(message)),
                             {ViewerText::DiffTextVariant::SemanticRowKind::Placeholder});
        };

        auto appendGap = [&](uint32_t oldStartInclusive, uint32_t oldEndExclusive, uint32_t newStartInclusive, uint32_t newEndExclusive) noexcept
        {
            const uint32_t oldGap = oldEndExclusive > oldStartInclusive ? (oldEndExclusive - oldStartInclusive) : 0u;
            const uint32_t newGap = newEndExclusive > newStartInclusive ? (newEndExclusive - newStartInclusive) : 0u;
            if (oldGap == 0u && newGap == 0u)
            {
                return;
            }

            if (! hydrateExpandedContext)
            {
                appendOutputLine(BuildHiddenContextBannerLabel(oldGap, newGap),
                                 {ViewerText::DiffTextVariant::SemanticRowKind::HiddenContextBanner},
                                 {.clickableBanner = true});
                return;
            }

            if (leftFile && rightFile && ! leftFile->available && ! rightFile->available)
            {
                result.variant.referencedFilesResolved = false;
                appendPlaceholder(! leftFile->reason.empty() ? leftFile->reason : rightFile->reason);
                return;
            }

            if (! leftFile || ! rightFile || ! leftFile->available || ! rightFile->available)
            {
                result.variant.referencedFilesResolved = false;
            }

            const uint32_t rowCount = std::max(oldGap, newGap);
            if (hydratedLogicalRange.has_value() && rowCount > 0u)
            {
                const uint32_t gapLogicalStart = currentLogicalLine;
                const uint32_t gapLogicalEnd   = gapLogicalStart + rowCount;
                const uint32_t hydratedStart   = std::max<uint32_t>(gapLogicalStart, hydratedLogicalRange->first);
                const uint32_t hydratedEnd     = std::min<uint32_t>(gapLogicalEnd, hydratedLogicalRange->second);
                if (hydratedStart < hydratedEnd)
                {
                    const uint32_t rowEndExclusive = hydratedEnd - gapLogicalStart;
                    if (leftFile && leftFile->available && oldGap > 0u)
                    {
                        const uint32_t neededOldRows = std::min<uint32_t>(oldGap, rowEndExclusive);
                        if (neededOldRows > 0u)
                        {
                            static_cast<void>(EnsureReferencedDiffLinesLoaded(*leftFile, oldStartInclusive + neededOldRows - 1u));
                        }
                    }
                    if (rightFile && rightFile->available && newGap > 0u)
                    {
                        const uint32_t neededNewRows = std::min<uint32_t>(newGap, rowEndExclusive);
                        if (neededNewRows > 0u)
                        {
                            static_cast<void>(EnsureReferencedDiffLinesLoaded(*rightFile, newStartInclusive + neededNewRows - 1u));
                        }
                    }
                }
            }

            for (uint32_t row = 0u; row < rowCount; ++row)
            {
                const bool hasOld             = row < oldGap;
                const bool hasNew             = row < newGap;
                const bool hydrateLogicalLine = ShouldHydrateExpandedDiffLogicalLine(hydratedLogicalRange, currentLogicalLine);

                std::wstring_view lineText;
                if (hasNew && rightFile && rightFile->available)
                {
                    lineText = TryGetReferencedDiffLine(*rightFile, newStartInclusive + row);
                }
                if (lineText.empty() && hasOld && leftFile && leftFile->available)
                {
                    lineText = TryGetReferencedDiffLine(*leftFile, oldStartInclusive + row);
                }

                std::wstring deferredLineText;
                if (! hydrateLogicalLine)
                {
                    deferredLineText = MaskDeferredDiffText(lineText);
                    lineText         = deferredLineText;
                    result.variant.deferredContextRowCount += 1u;
                }

                appendInlineRow(hasOld, oldStartInclusive + row, hasNew, newStartInclusive + row, L' ', lineText);
                result.variant.hasExpandedContext = true;
            }
        };

        uint32_t nextOldLine = 1u;
        uint32_t nextNewLine = 1u;
        for (size_t hunkIndex = 0u; hunkIndex < file.hunks.size(); ++hunkIndex)
        {
            const ParsedDiffHunk& hunk = file.hunks[hunkIndex];
            appendGap(nextOldLine, hunk.oldStart, nextNewLine, hunk.newStart);
            result.variant.hunkNavigation.push_back(ViewerText::DiffTextVariant::HunkNavigationEntry{
                BuildDiffHunkNavigationLabel(file, fileIndex, hunk, hunkIndex), currentLogicalLine, fileIndex});

            uint32_t oldCursorLine = hunk.oldStart;
            uint32_t newCursorLine = hunk.newStart;
            for (const ParsedDiffLine& line : hunk.lines)
            {
                switch (line.kind)
                {
                    case ParsedDiffLineKind::Context:
                        appendInlineRow(true, line.oldLine, true, line.newLine, L' ', line.text);
                        oldCursorLine += 1u;
                        newCursorLine += 1u;
                        break;
                    case ParsedDiffLineKind::Added:
                        appendInlineRow(false, 0u, true, line.newLine, L'+', line.text);
                        newCursorLine += 1u;
                        break;
                    case ParsedDiffLineKind::Removed:
                        appendInlineRow(true, line.oldLine, false, 0u, L'-', line.text);
                        oldCursorLine += 1u;
                        break;
                    case ParsedDiffLineKind::NoNewlineMarker: appendPlaceholder(line.text); break;
                }
            }

            nextOldLine = oldCursorLine;
            nextNewLine = newCursorLine;
        }

        if (hydrateExpandedContext)
        {
            uint32_t requestedTailRows = 0u;
            if (hydratedLogicalRange.has_value() && hydratedLogicalRange->second > currentLogicalLine)
            {
                requestedTailRows = hydratedLogicalRange->second - currentLogicalLine;
            }

            if (requestedTailRows > 0u)
            {
                if (leftFile && leftFile->available)
                {
                    static_cast<void>(EnsureReferencedDiffLinesLoaded(*leftFile, nextOldLine + requestedTailRows - 1u));
                }
                if (rightFile && rightFile->available)
                {
                    static_cast<void>(EnsureReferencedDiffLinesLoaded(*rightFile, nextNewLine + requestedTailRows - 1u));
                }
            }

            uint32_t oldTail = nextOldLine;
            uint32_t newTail = nextNewLine;
            if (leftFile && leftFile->available && leftFile->lines.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
            {
                oldTail = std::max<uint32_t>(nextOldLine, static_cast<uint32_t>(leftFile->lines.size()) + 1u);
                if (! leftFile->loadComplete)
                {
                    result.variant.hasExpandableTail = true;
                }
            }
            if (rightFile && rightFile->available && rightFile->lines.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
            {
                newTail = std::max<uint32_t>(nextNewLine, static_cast<uint32_t>(rightFile->lines.size()) + 1u);
                if (! rightFile->loadComplete)
                {
                    result.variant.hasExpandableTail = true;
                }
            }
            appendGap(nextOldLine, oldTail, nextNewLine, newTail);
        }
    }

    result.variant.builtLogicalLineCount = currentLogicalLine;
    if (! result.variant.text.empty())
    {
        result.variant.logicalRowStyles.push_back({});
        result.variant.logicalRowPaneLayouts.push_back({});
    }

    return result;
}

DiffVariantBuildResult BuildSideBySideDiffText(const ParsedDiffDocument& document,
                                               ViewerText::DiffContextMode contextMode,
                                               const std::filesystem::path& diffPath,
                                               IFileSystemIO* fileIo,
                                               DiffReferenceCache* referenceCache,
                                               std::optional<size_t> hydratedSectionIndex,
                                               std::optional<std::pair<uint32_t, uint32_t>> hydratedLogicalRange) noexcept
{
    DiffVariantBuildResult result{};
    result.variant.fileSectionCount        = document.files.size();
    result.variant.referencedFilesResolved = (contextMode == ViewerText::DiffContextMode::FullFileWhenAvailable);
    if (hydratedLogicalRange.has_value())
    {
        result.variant.hydratedLogicalLineStart        = hydratedLogicalRange->first;
        result.variant.hydratedLogicalLineEndExclusive = hydratedLogicalRange->second;
    }

    std::unordered_map<std::wstring, std::shared_ptr<ResolvedDiffTextFile>> localCache;
    auto& cache                 = referenceCache ? referenceCache->files : localCache;
    uint32_t currentLogicalLine = 0u;
    const auto recordRowStyle   = [&](const ViewerText::DiffTextVariant::LogicalRowStyleEntry& style) noexcept
    {
        const auto countKind = [&](ViewerText::DiffTextVariant::SemanticRowKind kind) noexcept
        {
            switch (kind)
            {
                case ViewerText::DiffTextVariant::SemanticRowKind::Context: result.variant.contextRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::Added: result.variant.addedRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::Removed: result.variant.removedRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::FileHeader: result.variant.headerRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::HunkHeader:
                case ViewerText::DiffTextVariant::SemanticRowKind::HiddenContextBanner: result.variant.bannerRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::Placeholder: result.variant.placeholderRowCount += 1u; break;
                case ViewerText::DiffTextVariant::SemanticRowKind::None: break;
            }
        };

        bool styled = false;
        if (style.fullRow != ViewerText::DiffTextVariant::SemanticRowKind::None)
        {
            countKind(style.fullRow);
            styled = true;
        }
        else
        {
            if (style.leftPane != ViewerText::DiffTextVariant::SemanticRowKind::None)
            {
                countKind(style.leftPane);
                styled = true;
            }
            if (style.rightPane != ViewerText::DiffTextVariant::SemanticRowKind::None)
            {
                countKind(style.rightPane);
                styled = true;
            }
        }

        if (styled)
        {
            result.variant.styledRowCount += 1u;
        }
    };
    const auto appendOutputLine =
        [&](std::wstring_view line,
            ViewerText::DiffTextVariant::LogicalRowStyleEntry style       = ViewerText::DiffTextVariant::LogicalRowStyleEntry{},
            ViewerText::DiffTextVariant::LogicalRowRenderEntry renderInfo = ViewerText::DiffTextVariant::LogicalRowRenderEntry{}) noexcept
    {
        const uint32_t lineStartIndex = static_cast<uint32_t>(std::min<size_t>(result.variant.text.size(), std::numeric_limits<uint32_t>::max()));
        if (renderInfo.fullMarkerIndex != std::numeric_limits<uint32_t>::max())
        {
            renderInfo.fullMarkerIndex += lineStartIndex;
        }
        if (renderInfo.leftMarkerIndex != std::numeric_limits<uint32_t>::max())
        {
            renderInfo.leftMarkerIndex += lineStartIndex;
        }
        if (renderInfo.rightMarkerIndex != std::numeric_limits<uint32_t>::max())
        {
            renderInfo.rightMarkerIndex += lineStartIndex;
        }
        AppendLine(result.variant.text, line);
        result.variant.logicalRowStyles.push_back(style);
        result.variant.logicalRowRenderInfo.push_back(renderInfo);
        result.variant.logicalRowPaneLayouts.push_back({});
        recordRowStyle(style);
        currentLogicalLine += 1u;
    };
    const auto appendPlaceholderBand = [&](ViewerText::DiffTextVariant::PlaceholderBandPlacement placement) noexcept
    {
        result.variant.placeholderBands.push_back(ViewerText::DiffTextVariant::PlaceholderBandEntry{currentLogicalLine, placement});
        result.variant.placeholderBandCount = result.variant.placeholderBands.size();
    };

    for (size_t fileIndex = 0u; fileIndex < document.files.size(); ++fileIndex)
    {
        struct PendingSideBySideRow
        {
            bool splitRow = false;
            std::wstring fullText;
            std::wstring leftCell;
            std::wstring rightCell;
            ViewerText::DiffTextVariant::LogicalRowStyleEntry style{};
            ViewerText::DiffTextVariant::LogicalRowRenderEntry renderInfo{};
            ViewerText::DiffTextVariant::SideBySidePaneLayoutEntry paneLayout{};
        };

        const ParsedDiffFileSection& file = document.files[fileIndex];
        const bool hydrateExpandedContext = ShouldHydrateExpandedDiffSection(contextMode, hydratedSectionIndex, fileIndex);
        if (fileIndex != 0u)
        {
            appendOutputLine(L"");
        }

        result.variant.sectionNavigation.push_back(
            ViewerText::DiffTextVariant::SectionNavigationEntry{BuildDiffSectionNavigationLabel(file, fileIndex), currentLogicalLine});
        appendOutputLine(std::format(L"old path: {}", FormatDiffSummaryPath(file.leftDisplayPath)), {ViewerText::DiffTextVariant::SemanticRowKind::FileHeader});
        appendOutputLine(std::format(L"new path: {}", FormatDiffSummaryPath(file.rightDisplayPath)),
                         {ViewerText::DiffTextVariant::SemanticRowKind::FileHeader});
        for (const std::wstring& metadataLine : file.metadataLines)
        {
            appendOutputLine(metadataLine, {ViewerText::DiffTextVariant::SemanticRowKind::FileHeader});
        }

        std::shared_ptr<ResolvedDiffTextFile> leftFile;
        std::shared_ptr<ResolvedDiffTextFile> rightFile;
        if (hydrateExpandedContext)
        {
            const UINT leftMissingId =
                (file.leftDisplayPath == L"/dev/null") ? IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_ADDED_FILE : IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_MISSING_FILE;
            const UINT rightMissingId =
                (file.rightDisplayPath == L"/dev/null") ? IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_DELETED_FILE : IDS_VIEWERTEXT_MSG_DIFF_UNCHANGED_MISSING_FILE;
            leftFile  = ResolveDiffReference(fileIo, diffPath, file.leftDisplayPath, leftMissingId, cache);
            rightFile = ResolveDiffReference(fileIo, diffPath, file.rightDisplayPath, rightMissingId, cache);
        }

        uint32_t maxLineNumber = 1u;
        for (const ParsedDiffHunk& hunk : file.hunks)
        {
            for (const ParsedDiffLine& line : hunk.lines)
            {
                if (line.hasOldLine)
                {
                    maxLineNumber = std::max(maxLineNumber, line.oldLine);
                }
                if (line.hasNewLine)
                {
                    maxLineNumber = std::max(maxLineNumber, line.newLine);
                }
            }
        }

        if (hydrateExpandedContext && leftFile && leftFile->lines.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
        {
            maxLineNumber = std::max(maxLineNumber, static_cast<uint32_t>(leftFile->lines.size()));
        }
        if (hydrateExpandedContext && rightFile && rightFile->lines.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
        {
            maxLineNumber = std::max(maxLineNumber, static_cast<uint32_t>(rightFile->lines.size()));
        }

        const size_t digits = DecimalDigits(maxLineNumber);
        auto formatNumber   = [digits](bool hasValue, uint32_t value) -> std::wstring
        {
            if (! hasValue)
            {
                return std::wstring(digits, L' ');
            }
            return std::format(L"{:>{}}", value, digits);
        };
        std::vector<PendingSideBySideRow> pendingRows;
        const auto emitQueuedRows = [&]() noexcept
        {
            for (const PendingSideBySideRow& row : pendingRows)
            {
                auto renderInfo               = row.renderInfo;
                const uint32_t lineStartIndex = static_cast<uint32_t>(std::min<size_t>(result.variant.text.size(), std::numeric_limits<uint32_t>::max()));
                if (renderInfo.fullMarkerIndex != std::numeric_limits<uint32_t>::max())
                {
                    renderInfo.fullMarkerIndex += lineStartIndex;
                }
                if (renderInfo.leftMarkerIndex != std::numeric_limits<uint32_t>::max())
                {
                    renderInfo.leftMarkerIndex += lineStartIndex;
                }
                if (renderInfo.rightMarkerIndex != std::numeric_limits<uint32_t>::max())
                {
                    renderInfo.rightMarkerIndex += lineStartIndex;
                }
                if (row.splitRow)
                {
                    AppendLine(result.variant.text, std::format(L"{}{}{}", row.leftCell, std::wstring(row.paneLayout.separatorColumns, L' '), row.rightCell));
                }
                else
                {
                    AppendLine(result.variant.text, row.fullText);
                }
                result.variant.logicalRowStyles.push_back(row.style);
                result.variant.logicalRowRenderInfo.push_back(renderInfo);
                result.variant.logicalRowPaneLayouts.push_back(row.paneLayout);
                recordRowStyle(row.style);
            }
            pendingRows.clear();
        };
        const auto queueFullRow =
            [&](std::wstring_view line,
                ViewerText::DiffTextVariant::LogicalRowStyleEntry style       = ViewerText::DiffTextVariant::LogicalRowStyleEntry{},
                ViewerText::DiffTextVariant::LogicalRowRenderEntry renderInfo = ViewerText::DiffTextVariant::LogicalRowRenderEntry{}) noexcept
        {
            pendingRows.push_back(PendingSideBySideRow{
                .splitRow   = false,
                .fullText   = std::wstring(line),
                .style      = style,
                .renderInfo = renderInfo,
                .paneLayout = {},
            });
            currentLogicalLine += 1u;
        };

        auto appendRow = [&](bool hasLeftNumber,
                             uint32_t leftNumber,
                             wchar_t leftMarker,
                             std::wstring_view leftText,
                             bool hasRightNumber,
                             uint32_t rightNumber,
                             wchar_t rightMarker,
                             std::wstring_view rightText,
                             ViewerText::DiffTextVariant::LogicalRowStyleEntry style = ViewerText::DiffTextVariant::LogicalRowStyleEntry{}) noexcept
        {
            if (style.fullRow == ViewerText::DiffTextVariant::SemanticRowKind::None && style.leftPane == ViewerText::DiffTextVariant::SemanticRowKind::None &&
                style.rightPane == ViewerText::DiffTextVariant::SemanticRowKind::None)
            {
                if (leftMarker == L'!')
                {
                    style.leftPane = ViewerText::DiffTextVariant::SemanticRowKind::Placeholder;
                }
                else if (leftMarker == L'-')
                {
                    style.leftPane = ViewerText::DiffTextVariant::SemanticRowKind::Removed;
                }
                else if (leftMarker == L'+')
                {
                    style.leftPane = ViewerText::DiffTextVariant::SemanticRowKind::Added;
                }
                else if (hasLeftNumber)
                {
                    style.leftPane = ViewerText::DiffTextVariant::SemanticRowKind::Context;
                }

                if (rightMarker == L'!')
                {
                    style.rightPane = ViewerText::DiffTextVariant::SemanticRowKind::Placeholder;
                }
                else if (rightMarker == L'-')
                {
                    style.rightPane = ViewerText::DiffTextVariant::SemanticRowKind::Removed;
                }
                else if (rightMarker == L'+')
                {
                    style.rightPane = ViewerText::DiffTextVariant::SemanticRowKind::Added;
                }
                else if (hasRightNumber)
                {
                    style.rightPane = ViewerText::DiffTextVariant::SemanticRowKind::Context;
                }
            }

            std::wstring leftCell  = std::format(L"{} {} {}", formatNumber(hasLeftNumber, leftNumber), leftMarker, std::wstring(leftText));
            std::wstring rightCell = std::format(L"{} {} {}", formatNumber(hasRightNumber, rightNumber), rightMarker, std::wstring(rightText));
            ViewerText::DiffTextVariant::SideBySidePaneLayoutEntry paneLayout{};
            paneLayout.splitRow         = true;
            paneLayout.leftTextColumns  = static_cast<uint32_t>(std::min<size_t>(leftCell.size(), std::numeric_limits<uint32_t>::max()));
            paneLayout.separatorColumns = 3u;
            paneLayout.rightTextColumns = static_cast<uint32_t>(std::min<size_t>(rightCell.size(), std::numeric_limits<uint32_t>::max()));
            ViewerText::DiffTextVariant::LogicalRowRenderEntry renderInfo{};
            if (leftMarker == L'+' || leftMarker == L'-')
            {
                renderInfo.leftMarkerIndex = static_cast<uint32_t>(formatNumber(hasLeftNumber, leftNumber).size() + 1u);
            }
            if (rightMarker == L'+' || rightMarker == L'-')
            {
                renderInfo.rightMarkerIndex =
                    static_cast<uint32_t>(paneLayout.leftTextColumns + paneLayout.separatorColumns + formatNumber(hasRightNumber, rightNumber).size() + 1u);
            }
            renderInfo.leftPaneAbsent  = ! hasLeftNumber && leftMarker == L' ' && leftText.empty();
            renderInfo.rightPaneAbsent = ! hasRightNumber && rightMarker == L' ' && rightText.empty();
            pendingRows.push_back(PendingSideBySideRow{
                .splitRow   = true,
                .leftCell   = std::move(leftCell),
                .rightCell  = std::move(rightCell),
                .style      = style,
                .renderInfo = renderInfo,
                .paneLayout = paneLayout,
            });
            currentLogicalLine += 1u;
        };

        auto appendGap = [&](uint32_t oldStartInclusive, uint32_t oldEndExclusive, uint32_t newStartInclusive, uint32_t newEndExclusive) noexcept
        {
            const uint32_t oldGap = oldEndExclusive > oldStartInclusive ? (oldEndExclusive - oldStartInclusive) : 0u;
            const uint32_t newGap = newEndExclusive > newStartInclusive ? (newEndExclusive - newStartInclusive) : 0u;
            if (oldGap == 0u && newGap == 0u)
            {
                return;
            }

            if (! hydrateExpandedContext)
            {
                queueFullRow(BuildHiddenContextBannerLabel(oldGap, newGap),
                             {ViewerText::DiffTextVariant::SemanticRowKind::HiddenContextBanner},
                             {.clickableBanner = true});
                return;
            }

            if (! leftFile || ! rightFile || ! leftFile->available || ! rightFile->available)
            {
                result.variant.referencedFilesResolved = false;
            }

            if (leftFile && rightFile && ! leftFile->available && ! rightFile->available)
            {
                result.variant.hasPlaceholderRows = true;
                appendPlaceholderBand(ViewerText::DiffTextVariant::PlaceholderBandPlacement::FullRow);
                appendRow(false,
                          0u,
                          L'!',
                          ! leftFile->reason.empty() ? leftFile->reason : rightFile->reason,
                          false,
                          0u,
                          L'!',
                          ! rightFile->reason.empty() ? rightFile->reason : leftFile->reason,
                          {ViewerText::DiffTextVariant::SemanticRowKind::Placeholder});
                return;
            }

            const uint32_t rowCount = std::max(oldGap, newGap);
            if (hydratedLogicalRange.has_value() && rowCount > 0u)
            {
                const uint32_t gapLogicalStart = currentLogicalLine;
                const uint32_t gapLogicalEnd   = gapLogicalStart + rowCount;
                const uint32_t hydratedStart   = std::max<uint32_t>(gapLogicalStart, hydratedLogicalRange->first);
                const uint32_t hydratedEnd     = std::min<uint32_t>(gapLogicalEnd, hydratedLogicalRange->second);
                if (hydratedStart < hydratedEnd)
                {
                    const uint32_t rowEndExclusive = hydratedEnd - gapLogicalStart;
                    if (leftFile && leftFile->available && oldGap > 0u)
                    {
                        const uint32_t neededOldRows = std::min<uint32_t>(oldGap, rowEndExclusive);
                        if (neededOldRows > 0u)
                        {
                            static_cast<void>(EnsureReferencedDiffLinesLoaded(*leftFile, oldStartInclusive + neededOldRows - 1u));
                        }
                    }
                    if (rightFile && rightFile->available && newGap > 0u)
                    {
                        const uint32_t neededNewRows = std::min<uint32_t>(newGap, rowEndExclusive);
                        if (neededNewRows > 0u)
                        {
                            static_cast<void>(EnsureReferencedDiffLinesLoaded(*rightFile, newStartInclusive + neededNewRows - 1u));
                        }
                    }
                }
            }

            for (uint32_t row = 0u; row < rowCount; ++row)
            {
                const bool hasLeftRow         = row < oldGap;
                const bool hasRightRow        = row < newGap;
                const bool hydrateLogicalLine = ShouldHydrateExpandedDiffLogicalLine(hydratedLogicalRange, currentLogicalLine);

                std::wstring_view leftText;
                if (hasLeftRow && leftFile && leftFile->available)
                {
                    leftText = TryGetReferencedDiffLine(*leftFile, oldStartInclusive + row);
                }

                std::wstring_view rightText;
                if (hasRightRow && rightFile && rightFile->available)
                {
                    rightText = TryGetReferencedDiffLine(*rightFile, newStartInclusive + row);
                }

                if (leftFile && ! leftFile->available && row == 0u)
                {
                    leftText                          = leftFile->reason;
                    result.variant.hasPlaceholderRows = true;
                    appendPlaceholderBand(ViewerText::DiffTextVariant::PlaceholderBandPlacement::LeftPane);
                }
                if (rightFile && ! rightFile->available && row == 0u)
                {
                    rightText                         = rightFile->reason;
                    result.variant.hasPlaceholderRows = true;
                    appendPlaceholderBand(ViewerText::DiffTextVariant::PlaceholderBandPlacement::RightPane);
                }

                std::wstring deferredLeftText;
                std::wstring deferredRightText;
                if (leftFile && rightFile && leftFile->available && rightFile->available && ! hydrateLogicalLine)
                {
                    deferredLeftText  = MaskDeferredDiffText(leftText);
                    deferredRightText = MaskDeferredDiffText(rightText);
                    leftText          = deferredLeftText;
                    rightText         = deferredRightText;
                    result.variant.deferredContextRowCount += 1u;
                }

                appendRow(hasLeftRow, oldStartInclusive + row, L' ', leftText, hasRightRow, newStartInclusive + row, L' ', rightText);
                result.variant.hasExpandedContext = true;
            }
        };

        uint32_t nextOldLine = 1u;
        uint32_t nextNewLine = 1u;
        for (size_t hunkIndex = 0u; hunkIndex < file.hunks.size(); ++hunkIndex)
        {
            const ParsedDiffHunk& hunk = file.hunks[hunkIndex];
            appendGap(nextOldLine, hunk.oldStart, nextNewLine, hunk.newStart);
            result.variant.hunkNavigation.push_back(ViewerText::DiffTextVariant::HunkNavigationEntry{
                BuildDiffHunkNavigationLabel(file, fileIndex, hunk, hunkIndex), currentLogicalLine, fileIndex});

            std::vector<const ParsedDiffLine*> removedRun;
            std::vector<const ParsedDiffLine*> addedRun;
            auto flushRuns = [&]() noexcept
            {
                const size_t rowCount = std::max(removedRun.size(), addedRun.size());
                for (size_t row = 0u; row < rowCount; ++row)
                {
                    const ParsedDiffLine* removed = row < removedRun.size() ? removedRun[row] : nullptr;
                    const ParsedDiffLine* added   = row < addedRun.size() ? addedRun[row] : nullptr;
                    appendRow(removed != nullptr && removed->hasOldLine,
                              removed ? removed->oldLine : 0u,
                              removed ? L'-' : L' ',
                              removed ? std::wstring_view(removed->text) : std::wstring_view{},
                              added != nullptr && added->hasNewLine,
                              added ? added->newLine : 0u,
                              added ? L'+' : L' ',
                              added ? std::wstring_view(added->text) : std::wstring_view{});
                }
                removedRun.clear();
                addedRun.clear();
            };

            uint32_t oldCursorLine = hunk.oldStart;
            uint32_t newCursorLine = hunk.newStart;
            for (const ParsedDiffLine& line : hunk.lines)
            {
                if (line.kind == ParsedDiffLineKind::Removed)
                {
                    removedRun.push_back(&line);
                    oldCursorLine += 1u;
                    continue;
                }
                if (line.kind == ParsedDiffLineKind::Added)
                {
                    addedRun.push_back(&line);
                    newCursorLine += 1u;
                    continue;
                }

                flushRuns();
                switch (line.kind)
                {
                    case ParsedDiffLineKind::Context:
                        appendRow(true, line.oldLine, L' ', line.text, true, line.newLine, L' ', line.text);
                        oldCursorLine += 1u;
                        newCursorLine += 1u;
                        break;
                    case ParsedDiffLineKind::NoNewlineMarker:
                        result.variant.hasPlaceholderRows = true;
                        appendPlaceholderBand(ViewerText::DiffTextVariant::PlaceholderBandPlacement::FullRow);
                        appendRow(false, 0u, L'!', line.text, false, 0u, L'!', line.text, {ViewerText::DiffTextVariant::SemanticRowKind::Placeholder});
                        break;
                    case ParsedDiffLineKind::Added:
                    case ParsedDiffLineKind::Removed: break;
                }
            }

            flushRuns();
            nextOldLine = oldCursorLine;
            nextNewLine = newCursorLine;
        }

        if (hydrateExpandedContext)
        {
            uint32_t requestedTailRows = 0u;
            if (hydratedLogicalRange.has_value() && hydratedLogicalRange->second > currentLogicalLine)
            {
                requestedTailRows = hydratedLogicalRange->second - currentLogicalLine;
            }

            if (requestedTailRows > 0u)
            {
                if (leftFile && leftFile->available)
                {
                    static_cast<void>(EnsureReferencedDiffLinesLoaded(*leftFile, nextOldLine + requestedTailRows - 1u));
                }
                if (rightFile && rightFile->available)
                {
                    static_cast<void>(EnsureReferencedDiffLinesLoaded(*rightFile, nextNewLine + requestedTailRows - 1u));
                }
            }

            uint32_t oldTail = nextOldLine;
            uint32_t newTail = nextNewLine;
            if (leftFile && leftFile->available && leftFile->lines.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
            {
                oldTail = std::max<uint32_t>(nextOldLine, static_cast<uint32_t>(leftFile->lines.size()) + 1u);
                if (! leftFile->loadComplete)
                {
                    result.variant.hasExpandableTail = true;
                }
            }
            if (rightFile && rightFile->available && rightFile->lines.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
            {
                newTail = std::max<uint32_t>(nextNewLine, static_cast<uint32_t>(rightFile->lines.size()) + 1u);
                if (! rightFile->loadComplete)
                {
                    result.variant.hasExpandableTail = true;
                }
            }
            appendGap(nextOldLine, oldTail, nextNewLine, newTail);
        }

        emitQueuedRows();
    }

    result.variant.builtLogicalLineCount = currentLogicalLine;
    if (! result.variant.text.empty())
    {
        result.variant.logicalRowStyles.push_back({});
        result.variant.logicalRowPaneLayouts.push_back({});
    }
    return result;
}

uint32_t MakeBgra(uint8_t r, uint8_t g, uint8_t b, uint8_t a) noexcept
{
    return static_cast<uint32_t>(b) | (static_cast<uint32_t>(g) << 8) | (static_cast<uint32_t>(r) << 16) | (static_cast<uint32_t>(a) << 24);
}

bool PointInRoundedRect(int x, int y, int left, int top, int right, int bottom, int radius) noexcept
{
    if (x < left || x >= right || y < top || y >= bottom)
    {
        return false;
    }

    const int r = std::max(0, radius);
    if (r == 0)
    {
        return true;
    }

    const int innerLeft   = left + r;
    const int innerTop    = top + r;
    const int innerRight  = right - r;
    const int innerBottom = bottom - r;

    if (x >= innerLeft && x < innerRight)
    {
        return true;
    }
    if (y >= innerTop && y < innerBottom)
    {
        return true;
    }

    const auto inCorner = [&](int cx, int cy) noexcept
    {
        const int dx = x - cx;
        const int dy = y - cy;
        return (dx * dx + dy * dy) <= (r * r);
    };

    if (x < innerLeft && y < innerTop)
    {
        return inCorner(innerLeft, innerTop);
    }
    if (x >= innerRight && y < innerTop)
    {
        return inCorner(innerRight - 1, innerTop);
    }
    if (x < innerLeft && y >= innerBottom)
    {
        return inCorner(innerLeft, innerBottom - 1);
    }
    if (x >= innerRight && y >= innerBottom)
    {
        return inCorner(innerRight - 1, innerBottom - 1);
    }

    return true;
}

wil::unique_hicon CreateViewerTextIcon(int sizePx) noexcept
{
    if (sizePx <= 0 || sizePx > 256)
    {
        return {};
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = sizePx;
    bmi.bmiHeader.biHeight      = -sizePx;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hbitmap color(CreateDIBSection(nullptr, &bmi, DIB_RGB_COLORS, &bits, nullptr, 0));
    if (! color || ! bits)
    {
        return {};
    }

    auto* pixels            = static_cast<uint32_t*>(bits);
    const size_t pixelCount = static_cast<size_t>(sizePx) * static_cast<size_t>(sizePx);
    std::fill_n(pixels, pixelCount, 0u);

    const COLORREF baseRef   = RGB(0, 120, 215);
    const COLORREF borderRef = RGB(0, 90, 160);
    const COLORREF lineRef   = RGB(255, 255, 255);

    const uint32_t lineRgb   = static_cast<uint32_t>(lineRef);
    const uint8_t lineR      = static_cast<uint8_t>(lineRgb & 0xFFu);
    const uint8_t lineG      = static_cast<uint8_t>((lineRgb >> 8) & 0xFFu);
    const uint8_t lineB      = static_cast<uint8_t>((lineRgb >> 16) & 0xFFu);
    const uint32_t linePixel = MakeBgra(lineR, lineG, lineB, 255u);

    const int margin = std::max(1, sizePx / 8);
    const int left   = margin;
    const int top    = margin;
    const int right  = sizePx - margin;
    const int bottom = sizePx - margin;
    const int radius = std::max(2, sizePx / 6);

    const int border      = std::max(1, sizePx / 16);
    const int innerLeft   = left + border;
    const int innerTop    = top + border;
    const int innerRight  = right - border;
    const int innerBottom = bottom - border;
    const int innerRadius = std::max(0, radius - border);

    for (int y = 0; y < sizePx; ++y)
    {
        for (int x = 0; x < sizePx; ++x)
        {
            if (! PointInRoundedRect(x, y, left, top, right, bottom, radius))
            {
                continue;
            }

            const bool inInner = PointInRoundedRect(x, y, innerLeft, innerTop, innerRight, innerBottom, innerRadius);
            const COLORREF c   = inInner ? baseRef : borderRef;
            pixels[static_cast<size_t>(y) * static_cast<size_t>(sizePx) + static_cast<size_t>(x)] =
                MakeBgra(static_cast<uint8_t>(GetRValue(c)), static_cast<uint8_t>(GetGValue(c)), static_cast<uint8_t>(GetBValue(c)), 255u);
        }
    }

    const int lineLeft   = innerLeft + std::max(1, sizePx / 8);
    const int lineRight  = innerRight - std::max(1, sizePx / 8);
    const int lineHeight = std::max(1, sizePx / 14);
    const int lineGap    = std::max(1, sizePx / 10);
    const int firstLineY = innerTop + std::max(1, sizePx / 6);

    for (int i = 0; i < 3; ++i)
    {
        const int y0 = firstLineY + i * (lineHeight + lineGap);
        for (int y = y0; y < y0 + lineHeight; ++y)
        {
            if (y < innerTop || y >= innerBottom)
            {
                continue;
            }

            for (int x = lineLeft; x < lineRight; ++x)
            {
                if (x < innerLeft || x >= innerRight)
                {
                    continue;
                }

                pixels[static_cast<size_t>(y) * static_cast<size_t>(sizePx) + static_cast<size_t>(x)] = linePixel;
            }
        }
    }

    const size_t maskStride = static_cast<size_t>(((sizePx + 31) / 32) * 4);
    std::vector<uint8_t> maskBits(maskStride * static_cast<size_t>(sizePx), 0u);
    wil::unique_hbitmap mask(CreateBitmap(sizePx, sizePx, 1, 1, maskBits.data()));
    if (! mask)
    {
        return {};
    }

    ICONINFO ii{};
    ii.fIcon    = TRUE;
    ii.hbmColor = color.get();
    ii.hbmMask  = mask.get();

    return wil::unique_hicon(CreateIconIndirect(&ii));
}

uint32_t StableHash32(std::wstring_view text) noexcept
{
    uint32_t hash = 2166136261u;
    for (wchar_t ch : text)
    {
        hash ^= static_cast<uint32_t>(ch);
        hash *= 16777619u;
    }
    return hash;
}

COLORREF ColorFromHSV(float hueDegrees, float saturation, float value) noexcept
{
    const float h = std::fmod(std::max(0.0f, hueDegrees), 360.0f);
    const float s = std::clamp(saturation, 0.0f, 1.0f);
    const float v = std::clamp(value, 0.0f, 1.0f);

    const float c = v * s;
    const float x = c * (1.0f - std::fabs(std::fmod(h / 60.0f, 2.0f) - 1.0f));
    const float m = v - c;

    float rf = 0.0f;
    float gf = 0.0f;
    float bf = 0.0f;

    if (h < 60.0f)
    {
        rf = c;
        gf = x;
        bf = 0.0f;
    }
    else if (h < 120.0f)
    {
        rf = x;
        gf = c;
        bf = 0.0f;
    }
    else if (h < 180.0f)
    {
        rf = 0.0f;
        gf = c;
        bf = x;
    }
    else if (h < 240.0f)
    {
        rf = 0.0f;
        gf = x;
        bf = c;
    }
    else if (h < 300.0f)
    {
        rf = x;
        gf = 0.0f;
        bf = c;
    }
    else
    {
        rf = c;
        gf = 0.0f;
        bf = x;
    }

    const auto toByte = [](float v01) noexcept
    {
        const float scaled = std::clamp(v01 * 255.0f, 0.0f, 255.0f);
        return static_cast<BYTE>(std::lround(scaled));
    };

    const BYTE r = toByte(rf + m);
    const BYTE g = toByte(gf + m);
    const BYTE b = toByte(bf + m);
    return RGB(r, g, b);
}

COLORREF ResolveAccentColor(const ViewerTheme& theme, std::wstring_view seed) noexcept
{
    if (theme.rainbowMode)
    {
        const uint32_t h = StableHash32(seed);
        const float hue  = static_cast<float>(h % 360u);
        const float sat  = theme.darkBase ? 0.70f : 0.55f;
        const float val  = theme.darkBase ? 0.95f : 0.85f;
        return ColorFromHSV(hue, sat, val);
    }

    return ColorRefFromArgb(theme.accentArgb);
}

uint32_t ArgbFromColorRef(COLORREF rgb, uint8_t alpha = 0xFFu) noexcept
{
    const uint32_t r = static_cast<uint32_t>(GetRValue(rgb));
    const uint32_t g = static_cast<uint32_t>(GetGValue(rgb));
    const uint32_t b = static_cast<uint32_t>(GetBValue(rgb));
    return (static_cast<uint32_t>(alpha) << 24) | (r << 16) | (g << 8) | b;
}

void PopulateViewerDiffThemeDefaults(ViewerTheme& theme) noexcept
{
    const bool dark         = theme.darkMode != FALSE;
    const bool highContrast = theme.highContrast != FALSE;
    const bool rainbowMode  = theme.rainbowMode != FALSE;
    const COLORREF bg       = ColorRefFromArgb(theme.backgroundArgb);
    const COLORREF accent   = ResolveAccentColor(theme, L"viewer-diff");

    const auto ensureArgb = [&](uint32_t& target, COLORREF rgb, uint8_t alpha) noexcept
    {
        if (target == 0u)
        {
            target = ArgbFromColorRef(rgb, alpha);
        }
    };

    if (highContrast)
    {
        ensureArgb(theme.diffAddedBackgroundArgb, RGB(56, 198, 96), dark ? 90u : 72u);
        ensureArgb(theme.diffRemovedBackgroundArgb, RGB(224, 84, 84), dark ? 90u : 72u);
        ensureArgb(theme.diffContextBackgroundArgb, accent, dark ? 48u : 36u);
        ensureArgb(theme.diffHeaderBackgroundArgb, accent, dark ? 76u : 60u);
        ensureArgb(theme.diffBannerBackgroundArgb, accent, dark ? 96u : 76u);
        ensureArgb(theme.diffPlaceholderBackgroundArgb, accent, dark ? 88u : 70u);
        ensureArgb(theme.diffDividerArgb, dark ? RGB(255, 255, 255) : RGB(0, 0, 0), dark ? 230u : 216u);
        return;
    }

    if (rainbowMode)
    {
        ensureArgb(theme.diffAddedBackgroundArgb, ResolveAccentColor(theme, L"viewer-diff-added"), dark ? 62u : 44u);
        ensureArgb(theme.diffRemovedBackgroundArgb, ResolveAccentColor(theme, L"viewer-diff-removed"), dark ? 62u : 44u);
        ensureArgb(theme.diffContextBackgroundArgb, accent, dark ? 28u : 18u);
        ensureArgb(theme.diffHeaderBackgroundArgb, accent, dark ? 46u : 30u);
        ensureArgb(theme.diffBannerBackgroundArgb, accent, dark ? 64u : 46u);
        ensureArgb(theme.diffPlaceholderBackgroundArgb, accent, dark ? 52u : 38u);
        ensureArgb(theme.diffDividerArgb, BlendColor(bg, accent, dark ? 40u : 28u), dark ? 232u : 210u);
        return;
    }

    ensureArgb(theme.diffAddedBackgroundArgb, RGB(46, 160, 67), dark ? 56u : 36u);
    ensureArgb(theme.diffRemovedBackgroundArgb, RGB(204, 51, 51), dark ? 56u : 36u);
    ensureArgb(theme.diffContextBackgroundArgb, accent, dark ? 20u : 12u);
    ensureArgb(theme.diffHeaderBackgroundArgb, accent, dark ? 40u : 24u);
    ensureArgb(theme.diffBannerBackgroundArgb, accent, dark ? 56u : 36u);
    ensureArgb(theme.diffPlaceholderBackgroundArgb, accent, dark ? 46u : 30u);
    ensureArgb(theme.diffDividerArgb, BlendColor(bg, accent, dark ? 28u : 18u), dark ? 220u : 200u);
}

int PxFromDip(int dip, UINT dpi) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
}

float DipsFromPixels(int px, UINT dpi) noexcept
{
    if (dpi == 0)
    {
        return static_cast<float>(px);
    }

    return static_cast<float>(px) * 96.0f / static_cast<float>(dpi);
}

D2D1_RECT_F RectFFromPixels(const RECT& rc, UINT dpi) noexcept
{
    const float left   = DipsFromPixels(static_cast<int>(rc.left), dpi);
    const float top    = DipsFromPixels(static_cast<int>(rc.top), dpi);
    const float right  = DipsFromPixels(static_cast<int>(rc.right), dpi);
    const float bottom = DipsFromPixels(static_cast<int>(rc.bottom), dpi);
    return D2D1::RectF(left, top, right, bottom);
}

D2D1_COLOR_F ColorFFromColorRef(COLORREF color, float alpha = 1.0f) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;
    return D2D1::ColorF(r, g, b, alpha);
}

void ClampRectNonNegative(RECT& rc) noexcept
{
    if (rc.right < rc.left)
    {
        rc.right = rc.left;
    }
    if (rc.bottom < rc.top)
    {
        rc.bottom = rc.top;
    }
}

bool TryParseOffset(std::wstring_view text, uint64_t& value) noexcept;

bool TryParseOffset(std::wstring_view text, uint64_t& value) noexcept
{
    std::wstring t(text);
    if (t.empty())
    {
        return false;
    }

    wchar_t* start = t.data();
    while (*start != L'\0' && std::iswspace(static_cast<wint_t>(*start)) != 0)
    {
        ++start;
    }

    if (*start == L'\0')
    {
        return false;
    }

    errno                           = 0;
    wchar_t* end                    = nullptr;
    const unsigned long long parsed = wcstoull(start, &end, 0);
    if (end == start)
    {
        return false;
    }
    if (errno == ERANGE)
    {
        return false;
    }

    while (*end != L'\0' && std::iswspace(static_cast<wint_t>(*end)) != 0)
    {
        ++end;
    }

    if (*end != L'\0')
    {
        return false;
    }

    value = static_cast<uint64_t>(parsed);
    return true;
}

HRESULT WriteAllHandle(HANDLE file, const void* data, size_t size) noexcept
{
    if (! file)
    {
        return E_INVALIDARG;
    }

    const uint8_t* bytes = reinterpret_cast<const uint8_t*>(data);
    size_t offset        = 0;
    while (offset < size)
    {
        const DWORD want = static_cast<DWORD>(std::min<size_t>(size - offset, std::numeric_limits<DWORD>::max()));
        DWORD written    = 0;
        if (WriteFile(file, bytes + offset, want, &written, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (written == 0)
        {
            return E_FAIL;
        }
        offset += written;
    }

    return S_OK;
}
} // namespace

const char* GetViewerTextStaticConfigurationSchema() noexcept
{
    return GetViewerTextStaticConfigurationSchemaImpl();
}

ViewerText::ViewerText()
{
    _metaId          = L"builtin/viewer-text";
    _metaShortId     = L"read";
    _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_NAME);
    _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DESCRIPTION);

    _displayEncodingMenuSelection = IDM_VIEWER_ENCODING_DISPLAY_ANSI;
    _saveEncodingMenuSelection    = IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL;

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    static_cast<void>(SetConfiguration(nullptr));
}

ViewerText::~ViewerText()
{
    if (! _hFileComboHost)
    {
        return;
    }

    _fileComboControl            = nullptr;
    _fileComboHostPreExpandPopup = false;
    _lastSyncedFileComboIndex.reset();
    UnhookFileComboHostWindow(_hFileComboHost.get());
    _fileComboHost.Detach();
    _hFileComboHost.reset();
}

void ViewerText::SetHost(IHost* host) noexcept
{
    _hostAlerts = nullptr;

    if (! host)
    {
        return;
    }

    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hr = host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (SUCCEEDED(hr) && alerts)
    {
        _hostAlerts = std::move(alerts);
    }
}

HRESULT STDMETHODCALLTYPE ViewerText::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IViewer))
    {
        *ppvObject = static_cast<IViewer*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ViewerText::AddRef() noexcept
{
    return _refCount.fetch_add(1) + 1;
}

ULONG STDMETHODCALLTYPE ViewerText::Release() noexcept
{
    const ULONG remaining = _refCount.fetch_sub(1) - 1;
    if (remaining == 0)
    {
        delete this;
    }
    return remaining;
}

HRESULT STDMETHODCALLTYPE ViewerText::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = GetViewerTextStaticConfigurationSchema();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    uint32_t textBufferMiB              = 16;
    uint32_t hexBufferMiB               = 8;
    bool showLineNumbers                = false;
    bool wrapText                       = true;
    HexByteColorMode hexByteColorMode   = HexByteColorMode::LeadingNibble;
    DiffDefaultLayout diffDefaultLayout = DiffDefaultLayout::SideBySide;
    DiffContextMode diffContextMode     = DiffContextMode::HunksOnly;
    DiffAutoOpenMode diffAutoOpenMode   = DiffAutoOpenMode::Parsed;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        const std::string_view utf8(configurationJsonUtf8);
        if (! utf8.empty())
        {
            yyjson_doc* doc = yyjson_read(utf8.data(), utf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
            if (doc)
            {
                auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

                yyjson_val* root = yyjson_doc_get_root(doc);
                if (root && yyjson_is_obj(root))
                {
                    yyjson_val* textBuf = yyjson_obj_get(root, "textBufferMiB");
                    if (textBuf && yyjson_is_int(textBuf))
                    {
                        const int64_t value = yyjson_get_int(textBuf);
                        if (value > 0)
                        {
                            textBufferMiB = static_cast<uint32_t>(std::min<int64_t>(value, 256));
                        }
                    }

                    yyjson_val* hexBuf = yyjson_obj_get(root, "hexBufferMiB");
                    if (hexBuf && yyjson_is_int(hexBuf))
                    {
                        const int64_t value = yyjson_get_int(hexBuf);
                        if (value > 0)
                        {
                            hexBufferMiB = static_cast<uint32_t>(std::min<int64_t>(value, 256));
                        }
                    }

                    yyjson_val* lineNums = yyjson_obj_get(root, "showLineNumbers");
                    if (lineNums && yyjson_is_str(lineNums))
                    {
                        const char* value = yyjson_get_str(lineNums);
                        if (value != nullptr)
                        {
                            showLineNumbers = (strcmp(value, "1") == 0) || (strcmp(value, "true") == 0) || (strcmp(value, "on") == 0);
                        }
                    }

                    yyjson_val* wrap = yyjson_obj_get(root, "wrapText");
                    if (wrap && yyjson_is_str(wrap))
                    {
                        const char* value = yyjson_get_str(wrap);
                        if (value != nullptr)
                        {
                            wrapText = (strcmp(value, "1") == 0) || (strcmp(value, "true") == 0) || (strcmp(value, "on") == 0);
                        }
                    }

                    yyjson_val* hexByteColors = yyjson_obj_get(root, "hexByteColorMode");
                    if (hexByteColors && yyjson_is_str(hexByteColors))
                    {
                        const char* value = yyjson_get_str(hexByteColors);
                        if (value != nullptr && strcmp(value, "off") == 0)
                        {
                            hexByteColorMode = HexByteColorMode::Off;
                        }
                        else if (value != nullptr && strcmp(value, "leadingNibble") == 0)
                        {
                            hexByteColorMode = HexByteColorMode::LeadingNibble;
                        }
                    }

                    yyjson_val* diffLayout = yyjson_obj_get(root, "diffDefaultLayout");
                    if (diffLayout && yyjson_is_str(diffLayout))
                    {
                        const char* value = yyjson_get_str(diffLayout);
                        if (value != nullptr)
                        {
                            diffDefaultLayout = ParseDiffDefaultLayout(value);
                        }
                    }

                    yyjson_val* diffContext = yyjson_obj_get(root, "diffContextMode");
                    if (diffContext && yyjson_is_str(diffContext))
                    {
                        const char* value = yyjson_get_str(diffContext);
                        if (value != nullptr)
                        {
                            diffContextMode = ParseDiffContextMode(value);
                        }
                    }

                    yyjson_val* diffOpenMode = yyjson_obj_get(root, "diffAutoOpenMode");
                    if (diffOpenMode && yyjson_is_str(diffOpenMode))
                    {
                        const char* value = yyjson_get_str(diffOpenMode);
                        if (value != nullptr)
                        {
                            diffAutoOpenMode = ParseDiffAutoOpenMode(value);
                        }
                    }
                }
            }
        }
    }

    _config.textBufferMiB     = textBufferMiB;
    _config.hexBufferMiB      = hexBufferMiB;
    _config.showLineNumbers   = showLineNumbers;
    _config.wrapText          = wrapText;
    _config.hexByteColorMode  = hexByteColorMode;
    _config.diffDefaultLayout = diffDefaultLayout;
    _config.diffContextMode   = diffContextMode;
    _config.diffAutoOpenMode  = diffAutoOpenMode;
    _wrap                     = wrapText;

    RefreshConfigurationJson();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    if (_configurationJson.empty())
    {
        *configurationJsonUtf8 = nullptr;
        return S_OK;
    }

    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    const bool isDefault = _config.textBufferMiB == 16u && _config.hexBufferMiB == 8u && ! _config.showLineNumbers && _config.wrapText &&
                           _config.hexByteColorMode == HexByteColorMode::LeadingNibble && _config.diffDefaultLayout == DiffDefaultLayout::SideBySide &&
                           _config.diffContextMode == DiffContextMode::HunksOnly && _config.diffAutoOpenMode == DiffAutoOpenMode::Parsed;
    *pSomethingToSave    = isDefault ? FALSE : TRUE;
    return S_OK;
}

void ViewerText::RefreshConfigurationJson() noexcept
{
    const char* const hexByteColorMode = (_config.hexByteColorMode == HexByteColorMode::LeadingNibble) ? "leadingNibble" : "off";
    _configurationJson = std::format("{{\"textBufferMiB\":{},\"hexBufferMiB\":{},\"showLineNumbers\":\"{}\",\"wrapText\":\"{}\",\"hexByteColorMode\":\"{}\","
                                     "\"diffDefaultLayout\":\"{}\",\"diffContextMode\":\"{}\",\"diffAutoOpenMode\":\"{}\"}}",
                                     _config.textBufferMiB,
                                     _config.hexBufferMiB,
                                     _config.showLineNumbers ? "1" : "0",
                                     _config.wrapText ? "1" : "0",
                                     hexByteColorMode,
                                     DiffDefaultLayoutToConfigString(_config.diffDefaultLayout),
                                     DiffContextModeToConfigString(_config.diffContextMode),
                                     DiffAutoOpenModeToConfigString(_config.diffAutoOpenMode));
}

namespace
{
struct ViewerTextClassBackgroundBrushState
{
    ViewerTextClassBackgroundBrushState()                                                      = default;
    ViewerTextClassBackgroundBrushState(const ViewerTextClassBackgroundBrushState&)            = delete;
    ViewerTextClassBackgroundBrushState& operator=(const ViewerTextClassBackgroundBrushState&) = delete;
    ViewerTextClassBackgroundBrushState(ViewerTextClassBackgroundBrushState&&)                 = delete;
    ViewerTextClassBackgroundBrushState& operator=(ViewerTextClassBackgroundBrushState&&)      = delete;

    wil::unique_hbrush activeBrush;
    COLORREF activeColor = CLR_INVALID;

    wil::unique_hbrush pendingBrush;
    COLORREF pendingColor = CLR_INVALID;

    bool viewerClassRegistered   = false;
    bool textViewClassRegistered = false;
    bool hexViewClassRegistered  = false;
};

ViewerTextClassBackgroundBrushState g_viewerTextClassBackgroundBrush;

HBRUSH GetActiveViewerTextClassBackgroundBrush() noexcept
{
    if (g_viewerTextClassBackgroundBrush.pendingBrush)
    {
        return g_viewerTextClassBackgroundBrush.pendingBrush.get();
    }

    if (! g_viewerTextClassBackgroundBrush.activeBrush)
    {
        const COLORREF fallback                      = GetSysColor(COLOR_WINDOW);
        g_viewerTextClassBackgroundBrush.activeColor = fallback;
        g_viewerTextClassBackgroundBrush.activeBrush.reset(CreateSolidBrush(fallback));
    }

    return g_viewerTextClassBackgroundBrush.activeBrush.get();
}

void RequestViewerTextClassBackgroundColor(COLORREF color) noexcept
{
    if (color == CLR_INVALID)
    {
        return;
    }

    if (g_viewerTextClassBackgroundBrush.pendingBrush && g_viewerTextClassBackgroundBrush.pendingColor == color)
    {
        return;
    }

    wil::unique_hbrush brush(CreateSolidBrush(color));
    if (! brush)
    {
        return;
    }

    g_viewerTextClassBackgroundBrush.pendingColor = color;
    g_viewerTextClassBackgroundBrush.pendingBrush = std::move(brush);
}

void ApplyPendingViewerTextClassBackgroundBrush(HWND viewerHwnd, HWND textViewHwnd, HWND hexViewHwnd) noexcept
{
    if (! g_viewerTextClassBackgroundBrush.pendingBrush)
    {
        return;
    }

    const bool needViewer = g_viewerTextClassBackgroundBrush.viewerClassRegistered;
    const bool needText   = g_viewerTextClassBackgroundBrush.textViewClassRegistered;
    const bool needHex    = g_viewerTextClassBackgroundBrush.hexViewClassRegistered;

    if (! needViewer && ! needText && ! needHex)
    {
        return;
    }

    if ((needViewer && ! viewerHwnd) || (needText && ! textViewHwnd) || (needHex && ! hexViewHwnd))
    {
        return;
    }

    const LONG_PTR newBrush = reinterpret_cast<LONG_PTR>(g_viewerTextClassBackgroundBrush.pendingBrush.get());
    if (needViewer)
    {
        SetClassLongPtrW(viewerHwnd, GCLP_HBRBACKGROUND, newBrush);
    }
    if (needText)
    {
        SetClassLongPtrW(textViewHwnd, GCLP_HBRBACKGROUND, newBrush);
    }
    if (needHex)
    {
        SetClassLongPtrW(hexViewHwnd, GCLP_HBRBACKGROUND, newBrush);
    }

    g_viewerTextClassBackgroundBrush.activeBrush  = std::move(g_viewerTextClassBackgroundBrush.pendingBrush);
    g_viewerTextClassBackgroundBrush.activeColor  = g_viewerTextClassBackgroundBrush.pendingColor;
    g_viewerTextClassBackgroundBrush.pendingColor = CLR_INVALID;
}
} // namespace

ATOM ViewerText::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        g_viewerTextClassBackgroundBrush.viewerClassRegistered = true;
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = GetActiveViewerTextClassBackgroundBrush();
    wc.lpszClassName = kClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom                                                   = 1;
            g_viewerTextClassBackgroundBrush.viewerClassRegistered = true;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerText: RegisterClassExW failed.");
        }
    }
    else
    {
        g_viewerTextClassBackgroundBrush.viewerClassRegistered = true;
    }
    return atom;
}

ATOM ViewerText::RegisterTextViewClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        g_viewerTextClassBackgroundBrush.textViewClassRegistered = true;
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = TextViewProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_IBEAM);
    wc.hbrBackground = GetActiveViewerTextClassBackgroundBrush();
    wc.lpszClassName = kTextViewClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom                                                     = 1;
            g_viewerTextClassBackgroundBrush.textViewClassRegistered = true;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerText: RegisterClassExW failed for text view class.");
        }
    }
    else
    {
        g_viewerTextClassBackgroundBrush.textViewClassRegistered = true;
    }
    return atom;
}

ATOM ViewerText::RegisterHexViewClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        g_viewerTextClassBackgroundBrush.hexViewClassRegistered = true;
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = HexViewProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_IBEAM);
    wc.hbrBackground = GetActiveViewerTextClassBackgroundBrush();
    wc.lpszClassName = kHexViewClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom                                                    = 1;
            g_viewerTextClassBackgroundBrush.hexViewClassRegistered = true;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerText: RegisterClassExW failed for hex view class.");
        }
    }
    else
    {
        g_viewerTextClassBackgroundBrush.hexViewClassRegistered = true;
    }
    return atom;
}

LRESULT CALLBACK ViewerText::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<ViewerText*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            InitPostedPayloadWindow(hwnd);
        }
    }

    auto* self = reinterpret_cast<ViewerText*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK ViewerText::TextViewProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<ViewerText*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    auto* self = reinterpret_cast<ViewerText*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->TextViewProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK ViewerText::HexViewProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<ViewerText*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    auto* self = reinterpret_cast<ViewerText*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->HexViewProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerText::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
#ifdef ENABLE_TESTS
        case WndMsg::kViewerDebugGetNativeMenuModelSnapshot:
        {
            auto* menuSnapshot = reinterpret_cast<WndMsg::ViewerNativeMenuModelDebugSnapshot*>(lp);
            if (! menuSnapshot)
            {
                return FALSE;
            }

            *menuSnapshot                    = {};
            menuSnapshot->hasHiddenMenuModel = _menuHandle != nullptr;
            menuSnapshot->ownerDrawItemCount = CountOwnerDrawMenuItems(_menuHandle.get());
            return TRUE;
        }
#endif
#ifdef _DEBUG
        case WndMsg::kViewerTextDebugGetSnapshot:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerTextDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot          = {};
            snapshot->viewMode = (_viewMode == ViewMode::Hex) ? WndMsg::ViewerTextDebugViewMode::Hex : WndMsg::ViewerTextDebugViewMode::Text;
            snapshot->documentKind =
                (_documentKind == DocumentKind::Diff) ? WndMsg::ViewerTextDebugDocumentKind::Diff : WndMsg::ViewerTextDebugDocumentKind::PlainText;
            snapshot->diffPresentation = WndMsg::ViewerTextDebugDiffPresentation::None;
            if (_documentKind == DocumentKind::Diff)
            {
                switch (_diffPresentation)
                {
                    case DiffPresentationMode::RawText: snapshot->diffPresentation = WndMsg::ViewerTextDebugDiffPresentation::RawText; break;
                    case DiffPresentationMode::Inline: snapshot->diffPresentation = WndMsg::ViewerTextDebugDiffPresentation::Inline; break;
                    case DiffPresentationMode::SideBySide: snapshot->diffPresentation = WndMsg::ViewerTextDebugDiffPresentation::SideBySide; break;
                }
            }
            snapshot->hexByteColorMode = (_config.hexByteColorMode == HexByteColorMode::LeadingNibble) ? WndMsg::ViewerTextDebugHexByteColorMode::LeadingNibble
                                                                                                       : WndMsg::ViewerTextDebugHexByteColorMode::Off;
            snapshot->diffParsedAvailable           = _diffParsedAvailable;
            snapshot->fileComboUsesDiffSections     = UseDiffSectionFileCombo();
            snapshot->fileComboEntryCount           = ActiveFileComboEntryCount();
            snapshot->activeDiffSectionIndex        = CurrentDiffSectionIndex();
            snapshot->activeDiffHunkIndex           = CurrentDiffHunkIndex();
            snapshot->themeRainbow                  = _theme.rainbowMode != FALSE;
            snapshot->diffAddedBackgroundArgb       = _theme.diffAddedBackgroundArgb;
            snapshot->diffRemovedBackgroundArgb     = _theme.diffRemovedBackgroundArgb;
            snapshot->diffContextBackgroundArgb     = _theme.diffContextBackgroundArgb;
            snapshot->diffHeaderBackgroundArgb      = _theme.diffHeaderBackgroundArgb;
            snapshot->diffBannerBackgroundArgb      = _theme.diffBannerBackgroundArgb;
            snapshot->diffPlaceholderBackgroundArgb = _theme.diffPlaceholderBackgroundArgb;
            snapshot->diffDividerArgb               = _theme.diffDividerArgb;
            if (const auto* variant = CurrentDiffVariant())
            {
                snapshot->fileSectionCount                = variant->fileSectionCount;
                snapshot->diffHunkCount                   = variant->hunkNavigation.size();
                snapshot->styledRowCount                  = variant->styledRowCount;
                snapshot->contextRowCount                 = variant->contextRowCount;
                snapshot->addedRowCount                   = variant->addedRowCount;
                snapshot->removedRowCount                 = variant->removedRowCount;
                snapshot->headerRowCount                  = variant->headerRowCount;
                snapshot->bannerRowCount                  = variant->bannerRowCount;
                snapshot->placeholderRowCount             = variant->placeholderRowCount;
                snapshot->placeholderBandCount            = variant->placeholderBandCount;
                snapshot->deferredContextRowCount         = variant->deferredContextRowCount;
                snapshot->diffExpandedContext             = variant->hasExpandedContext;
                snapshot->diffHasPlaceholderRows          = variant->hasPlaceholderRows;
                snapshot->diffReferencedFilesResolved     = variant->referencedFilesResolved;
                snapshot->hydratedLogicalLineStart        = variant->hydratedLogicalLineStart;
                snapshot->hydratedLogicalLineEndExclusive = variant->hydratedLogicalLineEndExclusive;
                snapshot->builtLogicalLineCount           = variant->builtLogicalLineCount;
                snapshot->diffHasExpandableTail           = variant->hasExpandableTail;
                for (size_t logicalLine = 0; logicalLine < variant->logicalRowStyles.size() && logicalLine < variant->logicalRowRenderInfo.size();
                     ++logicalLine)
                {
                    if (variant->logicalRowStyles[logicalLine].fullRow == DiffTextVariant::SemanticRowKind::HiddenContextBanner &&
                        variant->logicalRowRenderInfo[logicalLine].clickableBanner)
                    {
                        snapshot->firstClickableBannerLogicalLine = logicalLine;
                        break;
                    }
                }
            }
            else if (_documentKind == DocumentKind::Diff && ! _diffParsedAvailable && _textStreamActive && ! _diffStreamSections.empty())
            {
                snapshot->fileSectionCount = _diffStreamSections.size();
            }
            if (_diffReferenceCache)
            {
                uint64_t referencedBytesRead = 0u;
                for (const auto& [key, value] : _diffReferenceCache->files)
                {
                    (void)key;
                    if (value)
                    {
                        referencedBytesRead += value->bytesRead;
                    }
                }
                snapshot->referencedBytesRead = referencedBytesRead;
            }
            snapshot->renderCount                      = (_viewMode == ViewMode::Hex) ? _debugHexRenderCount : _debugTextRenderCount;
            snapshot->legacyVisibleGdiTextSurfaceCount = 0u;
            snapshot->legacyVisibleHfontSurfaceCount   = 0u;
            snapshot->diffParseCount                   = _debugDiffParseCount;
            snapshot->visibleRowCount                  = (_viewMode == ViewMode::Hex) ? _debugHexVisibleRowCount : _debugTextVisibleRowCount;
            snapshot->textLeftColumn                   = _textLeftColumn;
            snapshot->visibleStyledRowCount            = _debugTextVisibleStyledRowCount;
            snapshot->visibleContextRowCount           = _debugTextVisibleContextRowCount;
            snapshot->visibleAddedRowCount             = _debugTextVisibleAddedRowCount;
            snapshot->visibleRemovedRowCount           = _debugTextVisibleRemovedRowCount;
            snapshot->visibleHeaderRowCount            = _debugTextVisibleHeaderRowCount;
            snapshot->visibleBannerRowCount            = _debugTextVisibleBannerRowCount;
            snapshot->visibleGapHatchCount             = _debugTextVisibleGapHatchCount;
            snapshot->visibleSplitRowCount             = _debugTextVisibleSplitRowCount;
            snapshot->textLastPaintUs                  = _debugTextLastPaintUs;
            snapshot->paneLocalSideBySideLayout        = HasPaneLocalSideBySideVisualLayout();
            snapshot->sideBySideLeftPaneColumns        = _textSideBySideLeftPaneColumns;
            snapshot->sideBySideRightPaneColumns       = _textSideBySideRightPaneColumns;
            snapshot->sideBySideSeparatorColumns       = _textSideBySideSeparatorColumns;
            snapshot->diffContextUsesBaseBackground    = _debugDiffContextUsesBaseBackground;
            snapshot->diffMarkerArgb                   = _debugDiffMarkerArgb;
            snapshot->diffGapHatchArgb                 = _debugDiffGapHatchArgb;
            snapshot->visibleByteCount                 = _debugHexVisibleByteCount;
            snapshot->visibleColorizedByteCount        = _debugHexColorizedByteCount;
            snapshot->visibleUniqueColorBucketCount    = _debugHexUniqueColorBucketCount;
            snapshot->highContrastFallback             = _debugHexHighContrastFallback;
            if (_viewMode == ViewMode::Text)
            {
                const auto copyPreviewText = [](std::wstring_view source, auto& destination) noexcept
                {
                    destination[0]         = L'\0';
                    const size_t copyCount = std::min(source.size(), std::size(destination) - 1u);
                    std::copy_n(source.begin(), copyCount, destination);
                    destination[copyCount] = L'\0';
                };

                const auto copyPreviewLine = [&](wchar_t(&destination)[WndMsg::kViewerTextDebugTextPreviewChars], size_t lineIndex) noexcept
                {
                    destination[0] = L'\0';
                    if (lineIndex >= _textLineStarts.size() || lineIndex >= _textLineEnds.size())
                    {
                        return;
                    }

                    const size_t start = static_cast<size_t>(_textLineStarts[lineIndex]);
                    const size_t end   = static_cast<size_t>(_textLineEnds[lineIndex]);
                    if (start > end || end > _textBuffer.size())
                    {
                        return;
                    }

                    const std::wstring_view line(_textBuffer.data() + static_cast<ptrdiff_t>(start), end - start);
                    const size_t copyCount = std::min(line.size(), std::size(destination) - 1u);
                    std::copy_n(line.begin(), copyCount, destination);
                    destination[copyCount] = L'\0';
                };

                copyPreviewLine(snapshot->firstTextLine, 0u);
                copyPreviewLine(snapshot->secondTextLine, 1u);
                if (! _textVisualLineLogical.empty())
                {
                    const size_t topVisual          = std::min<size_t>(_textTopVisualLine, _textVisualLineLogical.size() - 1u);
                    snapshot->topVisibleLogicalLine = _textVisualLineLogical[topVisual];
                    copyPreviewLine(snapshot->topVisibleTextLine, snapshot->topVisibleLogicalLine);
                    if (snapshot->topVisibleLogicalLine < _textLineStarts.size() && topVisual < _textVisualLineLayouts.size())
                    {
                        const uint32_t logicalStart = _textLineStarts[snapshot->topVisibleLogicalLine];
                        const auto& layout          = _textVisualLineLayouts[topVisual];
                        uint32_t segmentStart       = layout.segmentStartIndex;
                        const uint32_t segmentEnd   = layout.segmentEndIndex;
                        if (! _wrap && ! layout.splitPanes && segmentEnd >= segmentStart && _textLeftColumn != 0u)
                        {
                            const uint32_t skip = std::min<uint32_t>(_textLeftColumn, segmentEnd - segmentStart);
                            segmentStart += skip;
                        }
                        snapshot->topVisibleSegmentColumnStart = segmentStart >= logicalStart ? (segmentStart - logicalStart) : 0u;
                        if (layout.splitPanes)
                        {
                            snapshot->topVisibleLeftPaneColumnStart  = layout.leftStartIndex >= logicalStart ? (layout.leftStartIndex - logicalStart) : 0u;
                            snapshot->topVisibleRightPaneColumnStart = layout.rightStartIndex >= logicalStart ? (layout.rightStartIndex - logicalStart) : 0u;
                        }
                    }

                    const size_t maxVisibleVisual = std::min(_textVisualLineLayouts.size(), topVisual + std::max<size_t>(1u, _debugTextVisibleRowCount));
                    for (size_t visual = topVisual; visual < maxVisibleVisual; ++visual)
                    {
                        const auto& layout = _textVisualLineLayouts[visual];
                        if (! layout.splitPanes)
                        {
                            continue;
                        }

                        const size_t logicalLine = _textVisualLineLogical[visual];
                        if (logicalLine >= _textLineStarts.size())
                        {
                            break;
                        }

                        const uint32_t logicalStart                     = _textLineStarts[logicalLine];
                        snapshot->firstVisibleSplitLeftPaneColumnStart  = layout.leftStartIndex >= logicalStart ? (layout.leftStartIndex - logicalStart) : 0u;
                        snapshot->firstVisibleSplitRightPaneColumnStart = layout.rightStartIndex >= logicalStart ? (layout.rightStartIndex - logicalStart) : 0u;
                        break;
                    }
                }
                copyPreviewText(_textBuffer, snapshot->textPreview);
            }
            return TRUE;
        }
        case WndMsg::kViewerTextDebugSelectDiffSection: ScrollToDiffSection(hwnd, static_cast<size_t>(wp)); return TRUE;
        case WndMsg::kViewerTextDebugSelectDiffHunk: ScrollToDiffHunk(hwnd, static_cast<size_t>(wp)); return TRUE;
        case WndMsg::kViewerTextDebugClickTextLogicalLine: return (_hEdit && DebugClickTextLogicalLine(_hEdit.get(), static_cast<uint32_t>(wp))) ? TRUE : FALSE;
#endif
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_DPICHANGED: OnDpiChanged(hwnd, static_cast<UINT>(LOWORD(wp)), reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 520, 320);
            }
            return 0;
        case WM_COMMAND: OnCommand(hwnd, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp)); return 0;
        case WM_CONTEXTMENU: OnContextMenu(hwnd, RedSalamander::DxUi::ResolveNativeContextMenuScreenPoint(hwnd, lp)); return 0;
        case WM_NOTIFY: return OnNotify(reinterpret_cast<const NMHDR*>(lp));
        case WM_SYSKEYDOWN:
            if ((wp == VK_F10 || wp == VK_MENU) && _menuBarHost.FocusFirstItem())
            {
                return 0;
            }
            break;
        case WM_SYSCHAR:
            if (wp >= 0x20u && _menuBarHost.ActivateMnemonic(static_cast<wchar_t>(wp)))
            {
                return 0;
            }
            break;
        case WM_KEYDOWN:
            if (HandleShortcutKey(hwnd, wp))
            {
                return 0;
            }
            break;
        case WM_CTLCOLORLISTBOX:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLORBTN: return OnCtlColor(msg, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_MOUSEMOVE:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnMouseMove(pt.x, pt.y);
            return 0;
        }
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_LBUTTONDOWN:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnLButtonDown(pt.x, pt.y);
            return 0;
        }
        case WM_LBUTTONUP:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnLButtonUp(pt.x, pt.y);
            return 0;
        }
        case WM_TIMER: OnTimer(static_cast<UINT_PTR>(wp)); return 0;
        case WM_SETCURSOR:
            if (OnSetCursor(hwnd, lp))
            {
                return TRUE;
            }
            break;
        case kAsyncOpenCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncOpenResult>(lp);
            OnAsyncOpenComplete(std::move(result));
            return 0;
        }
        case WM_PAINT: OnPaint(); return 0;
        case WM_ERASEBKGND: return _allowEraseBkgnd ? DefWindowProcW(hwnd, msg, wp, lp) : 1;
        case WM_CLOSE: CommandExit(hwnd); return 0;
        case WM_NCACTIVATE: OnNcActivate(wp != FALSE); return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_NCDESTROY: return OnNcDestroy(hwnd, wp, lp);
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void ViewerText::OnNcActivate(bool windowActive) noexcept
{
    ApplyTitleBarTheme(windowActive);
}

LRESULT ViewerText::OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    OnDestroy();
    static_cast<void>(DrainPostedPayloadsForWindow(hwnd));

    _menuBarHost.Detach();
    _menuHandle.reset();
    UnhookFileComboHostWindow(_hFileComboHost.get());
    _fileComboHost.Detach();
    _hFileComboHost.release();
    _hEdit.release();
    _hHex.release();
    _hWnd.release();
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

    Release();
    return DefWindowProcW(hwnd, WM_NCDESTROY, wp, lp);
}

LRESULT ViewerText::HandleFileComboHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept
{
    const bool popupWasOpen      = _fileComboControl && _fileComboControl->DebugIsPopupOpen();
    const bool preExpandForPopup = ! popupWasOpen && _fileComboControl && MessageMayOpenWindowComboPopup(msg, wp);
    if (preExpandForPopup)
    {
        _fileComboHostPreExpandPopup = true;
        if (_hWnd)
        {
            Layout(_hWnd.get());
        }
    }

    const LRESULT dxResult = _fileComboHost.HandleMessage(hwnd, msg, wp, lp, handled);
    if (msg == WM_NCDESTROY)
    {
        handled = true;
        _fileComboHost.ReleaseMouseCapture();
        _fileComboControl            = nullptr;
        _fileComboHostPreExpandPopup = false;
        _lastSyncedFileComboIndex.reset();
        _hFileComboHost.release();
        return dxResult;
    }

    const bool popupIsOpen = _fileComboControl && _fileComboControl->DebugIsPopupOpen();
    if (popupIsOpen != popupWasOpen || (preExpandForPopup && ! popupIsOpen))
    {
        _fileComboHostPreExpandPopup = false;
        if (_hWnd)
        {
            Layout(_hWnd.get());
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
    return dxResult;
}

void ViewerText::FocusMainSurfaceFromFileCombo(HWND hwnd) noexcept
{
    if (_embeddedMode)
    {
        return;
    }

    if (_viewMode == ViewMode::Hex && _hHex && IsWindow(_hHex.get()) != FALSE)
    {
        SetFocus(_hHex.get());
        return;
    }

    if (_hEdit && IsWindow(_hEdit.get()) != FALSE)
    {
        SetFocus(_hEdit.get());
        return;
    }

    if (hwnd && IsWindow(hwnd) != FALSE)
    {
        SetFocus(hwnd);
    }
}

void ViewerText::OnCreate(HWND hwnd)
{
    _allowEraseBkgnd         = true;
    _allowEraseBkgndTextView = true;
    _allowEraseBkgndHexView  = true;

    const DWORD comboHostStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | SS_NOTIFY;
    _hFileComboHost.reset(CreateWindowExW(
        0, L"Static", L"", comboHostStyle, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_VIEWERTEXT_FILE_COMBO)), g_hInstance, nullptr));
    if (! _hFileComboHost)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed for DxUi file combo host.");
    }
    else if (! _fileComboHost.Attach(_hFileComboHost.get()))
    {
        Debug::Error(L"ViewerText: failed to attach DxUi host for file combo.");
        _hFileComboHost.reset();
    }
    else if (! SetPropW(_hFileComboHost.get(), kFileComboHostStateProp, reinterpret_cast<HANDLE>(this)) ||
             ! InstallWndProcHook(_hFileComboHost.get(), kFileComboHostOriginalWndProcProp, FileComboHostWndProc))
    {
        RemovePropW(_hFileComboHost.get(), kFileComboHostStateProp);
        Debug::ErrorWithLastError(L"ViewerText: failed to install WNDPROC hook for DxUi file combo host.");
        _fileComboHost.Detach();
        _hFileComboHost.reset();
    }
    else
    {
        auto combo        = std::make_unique<ComboBox>();
        _fileComboControl = combo.get();
        _fileComboControl->SetVariant(ComboBoxVariant::Window);
        _fileComboControl->SetOnSelectionChanged([this, hwnd](size_t selectedIndex)
        {
            if (_syncingFileCombo)
            {
                return;
            }

            if (UseDiffSectionFileCombo())
            {
                ScrollToDiffSection(hwnd, selectedIndex);
            }
            else
            {
                if (selectedIndex >= _otherFiles.size())
                {
                    return;
                }

                _otherIndex = selectedIndex;
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
            }

            if (_embeddedMode)
            {
                return;
            }

            if (_viewMode == ViewMode::Hex)
            {
                if (_hHex)
                {
                    SetFocus(_hHex.get());
                }
            }
            else if (_hEdit)
            {
                SetFocus(_hEdit.get());
            }
            else
            {
                SetFocus(hwnd);
            }
        });
        _fileComboHost.SetOnTabBoundary([this, hwnd](bool) noexcept
        {
            FocusMainSurfaceFromFileCombo(hwnd);
            return true;
        });
        _fileComboHost.SetOnEscape([hwnd]() noexcept
        {
            PostMessageW(hwnd, WM_CLOSE, 0, 0);
            return true;
        });
        _fileComboHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
        _fileComboHost.SetRoot(std::move(combo));
    }

    if (! _menuHandle)
    {
        _menuHandle.reset(GetMenu(hwnd));
    }
    if (_menuHandle)
    {
        _menuBarHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
        _menuBarHost.SetRefreshMenuStateCallback([this, hwnd] { UpdateMenuChecks(hwnd, false); });
        static_cast<void>(_menuBarHost.Attach(g_hInstance, hwnd, _menuHandle.get()));
    }

    static_cast<void>(RegisterTextViewClass(g_hInstance));
    const DWORD textStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL;
    _hEdit.reset(CreateWindowExW(0, kTextViewClassName, nullptr, textStyle, 0, 0, 0, 0, hwnd, nullptr, g_hInstance, this));
    if (! _hEdit)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed for DirectX text view.");
    }

    static_cast<void>(RegisterHexViewClass(g_hInstance));
    const DWORD hexStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL;
    _hHex.reset(CreateWindowExW(0, kHexViewClassName, nullptr, hexStyle, 0, 0, 0, 0, hwnd, nullptr, g_hInstance, this));
    if (! _hHex)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed for DirectX hex view.");
    }

    ApplyTheme(hwnd);
    RefreshFileCombo(hwnd);
    Layout(hwnd);
    SetViewMode(hwnd, _viewMode);
    SetWrap(hwnd, _wrap);
}

void ViewerText::OnDestroy()
{
    EndLoadingUi();
    DiscardDirect2D();
    DiscardTextViewDirect2D();
    DiscardHexViewDirect2D();
    ResetHexState();
    _windowIconSmall.reset();
    _windowIconBig.reset();

    NotifyViewerClosed();
}

void ViewerText::StartAsyncOpen(HWND hwnd, const std::filesystem::path& path, bool updateOtherFiles, UINT displayEncodingMenuSelection) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (path.empty())
    {
        Debug::Error(L"ViewerText: StartAsyncOpen called with an empty path.");
        return;
    }

    if (! _fileSystem)
    {
        Debug::Error(L"ViewerText: StartAsyncOpen failed because file system is missing.");
        return;
    }

    const bool pathChanged = (_currentPath != path);
    _currentPath           = path;

    if (updateOtherFiles)
    {
        _otherFiles.clear();
        _otherFiles.push_back(path);
        _otherIndex = 0;
        RefreshFileCombo(hwnd);
    }
    else
    {
        SyncFileComboSelection();
    }

    const std::wstring title = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_TITLE_FORMAT, path.filename().wstring());
    if (! title.empty())
    {
        SetWindowTextW(hwnd, title.c_str());
    }

    _statusMessage.clear();
    _fileReader.reset();
    _fileSize              = 0;
    _encoding              = FileEncoding::Unknown;
    _bomBytes              = 0;
    _textStreamActive      = false;
    _textStreamSkipBytes   = 0;
    _textStreamStartOffset = 0;
    _textStreamEndOffset   = 0;
    _textTotalLineCount.reset();
    _textStreamLineCountedEndOffset = 0;
    _textStreamLineCountedNewlines  = 0;
    _textStreamLineCountLastWasCR   = false;
    _detectedCodePage               = 0;
    _detectedCodePageValid          = false;
    _detectedCodePageIsGuess        = false;
    ResetDiffState();

    _textBuffer.clear();
    _searchMatchStarts.clear();
    _textLineStarts.clear();
    _textLineEnds.clear();
    _textVisualLineStarts.clear();
    _textVisualLineLogical.clear();
    _textTopVisualLine   = 0;
    _textLeftColumn      = 0;
    _textCaretIndex      = 0;
    _textSelAnchor       = 0;
    _textSelActive       = 0;
    _textPreferredColumn = 0;
    _textSelecting       = false;
    _textMaxLineLength   = 0;

    ResetHexState();

    BeginLoadingUi();
    SetViewMode(hwnd, _viewMode);
    InvalidateRect(hwnd, nullptr, TRUE);

    const uint64_t requestId  = _asyncOpenRequestId.fetch_add(1, std::memory_order_relaxed) + 1u;
    _activeAsyncOpenRequestId = requestId;

    const ViewMode desiredViewMode              = _viewMode;
    const UINT previousDisplayEncodingSelection = _displayEncodingMenuSelection;
    const uint32_t textBufferMiB                = _config.textBufferMiB;
    const uint32_t hexBufferMiB                 = _config.hexBufferMiB;
    const DiffDefaultLayout diffDefaultLayout   = _config.diffDefaultLayout;
    const DiffAutoOpenMode diffAutoOpenMode     = _config.diffAutoOpenMode;
    const bool allowHexFallback                 = static_cast<bool>(_hHex);

    wil::com_ptr<IFileSystem> fileSystem = _fileSystem;

    AddRef();

    struct AsyncOpenWorkItem final
    {
        AsyncOpenWorkItem()                                    = default;
        AsyncOpenWorkItem(const AsyncOpenWorkItem&)            = delete;
        AsyncOpenWorkItem& operator=(const AsyncOpenWorkItem&) = delete;

        wil::unique_hmodule moduleKeepAlive;
        std::function<void()> work;
    };

    auto ctx = std::unique_ptr<AsyncOpenWorkItem>(new (std::nothrow) AsyncOpenWorkItem{});

    ctx->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerTextModuleAnchor);
    ctx->work            = [this,
                            hwnd,
                            requestId,
                            fileSystem = std::move(fileSystem),
                            path,
                            pathChanged,
                            desiredViewMode,
                            updateOtherFiles,
                            displayEncodingMenuSelection,
                            previousDisplayEncodingSelection,
                            textBufferMiB,
                            hexBufferMiB,
                            diffDefaultLayout,
                            diffAutoOpenMode,
                            allowHexFallback]() mutable
    {
        auto releaseSelf = wil::scope_exit([&] { Release(); });

        // Sleep(15000); // Simulate long operation for testing purposes.
        std::unique_ptr<AsyncOpenResult> result(new (std::nothrow) AsyncOpenResult{});
        if (! result)
        {
            return;
        }

        result->viewer           = this;
        result->requestId        = requestId;
        result->path             = path;
        result->updateOtherFiles = updateOtherFiles;
        result->viewMode         = desiredViewMode;
        result->hr               = E_FAIL;

        wil::com_ptr<IFileSystemIO> fileIo;
        const HRESULT fileIoHr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
        if (FAILED(fileIoHr) || ! fileIo)
        {
            Debug::Error(L"ViewerText: Active filesystem does not implement IFileSystemIO (hr=0x{:08X}).", static_cast<unsigned long>(fileIoHr));
            result->hr = FAILED(fileIoHr) ? fileIoHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            return;
        }

        const HRESULT openReaderHr = fileIo->CreateFileReader(path.c_str(), result->fileReader.put());
        if (FAILED(openReaderHr) || ! result->fileReader)
        {
            Debug::Error(L"ViewerText: Failed to create file reader for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(openReaderHr));
            result->hr = FAILED(openReaderHr) ? openReaderHr : E_FAIL;
            return;
        }

        FileEncoding encoding     = FileEncoding::Unknown;
        uint64_t bomBytes         = 0;
        uint64_t detectedFileSize = 0;

        uint64_t sizeBytes   = 0;
        const HRESULT sizeHr = result->fileReader->GetSize(&sizeBytes);
        if (FAILED(sizeHr))
        {
            Debug::Error(L"ViewerText: GetSize failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(sizeHr));
            result->hr = sizeHr;
            return;
        }

        detectedFileSize = sizeBytes;

        BYTE bom[4]{};
        unsigned long read = 0;

        uint64_t pos         = 0;
        const HRESULT seekHr = result->fileReader->Seek(0, FILE_BEGIN, &pos);
        if (FAILED(seekHr))
        {
            Debug::Error(L"ViewerText: Seek(FILE_BEGIN, 0) failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(seekHr));
            result->hr = seekHr;
            return;
        }

        const HRESULT readHr = result->fileReader->Read(bom, static_cast<unsigned long>(std::size(bom)), &read);
        if (FAILED(readHr))
        {
            Debug::Error(L"ViewerText: Read failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(readHr));
            result->hr = readHr;
            return;
        }

        if (read >= 4 && bom[0] == 0xFF && bom[1] == 0xFE && bom[2] == 0x00 && bom[3] == 0x00)
        {
            encoding = FileEncoding::Utf32LE;
            bomBytes = 4;
        }
        else if (read >= 4 && bom[0] == 0x00 && bom[1] == 0x00 && bom[2] == 0xFE && bom[3] == 0xFF)
        {
            encoding = FileEncoding::Utf32BE;
            bomBytes = 4;
        }
        else if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
        {
            encoding = FileEncoding::Utf8;
            bomBytes = 3;
        }
        else if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
        {
            encoding = FileEncoding::Utf16LE;
            bomBytes = 2;
        }
        else if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
        {
            encoding = FileEncoding::Utf16BE;
            bomBytes = 2;
        }

        result->encoding = encoding;
        result->bomBytes = bomBytes;
        result->fileSize = detectedFileSize;

        UINT selection = previousDisplayEncodingSelection;
        if (displayEncodingMenuSelection != 0 && IsEncodingMenuSelectionValid(displayEncodingMenuSelection))
        {
            selection = displayEncodingMenuSelection;
        }
        else if (pathChanged)
        {
            selection = IDM_VIEWER_ENCODING_DISPLAY_ANSI;
            switch (encoding)
            {
                case FileEncoding::Utf8: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM; break;
                case FileEncoding::Utf16LE: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM; break;
                case FileEncoding::Utf16BE: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM; break;
                case FileEncoding::Utf32LE: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM; break;
                case FileEncoding::Utf32BE: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM; break;
                case FileEncoding::Unknown:
                default:
                {
                    constexpr unsigned long kProbeSize = 64u * 1024u;
                    std::array<uint8_t, kProbeSize> probe{};
                    unsigned long probeRead = 0;

                    uint64_t probePos = 0;
                    if (SUCCEEDED(result->fileReader->Seek(0, FILE_BEGIN, &probePos)))
                    {
                        static_cast<void>(result->fileReader->Read(probe.data(), static_cast<unsigned long>(probe.size()), &probeRead));
                        if (probeRead != 0 && IsValidUtf8(probe.data(), static_cast<size_t>(probeRead)))
                        {
                            selection = IDM_VIEWER_ENCODING_DISPLAY_UTF8;
                        }
                    }
                    break;
                }
            }
        }

        if (! IsEncodingMenuSelectionValid(selection))
        {
            selection = IDM_VIEWER_ENCODING_DISPLAY_ANSI;
        }

        result->displayEncodingMenuSelection = selection;

        const uint64_t streamSkipBytes = ::BytesToSkipForDisplayEncoding(selection, encoding, bomBytes);
        result->textStreamSkipBytes    = streamSkipBytes;

        const uint64_t clampedStart   = std::min<uint64_t>(streamSkipBytes, detectedFileSize);
        result->textStreamStartOffset = clampedStart;
        result->textStreamEndOffset   = clampedStart;
        result->textStreamActive      = false;

        const FileEncoding displayEncoding = DisplayEncodingFileEncodingForSelection(selection);
        const UINT displayCodePage         = CodePageForSelection(selection);
        const uint64_t maxChunkBytes       = ::TextStreamChunkBytes(textBufferMiB, displayEncoding);
        const uint64_t maxParsedDiffBytes  = kMaxFullyBufferedParsedDiffBytes;
        const bool diffByExtension         = HasDiffLikeExtension(path);

        const uint64_t availableBytes        = (detectedFileSize > clampedStart) ? (detectedFileSize - clampedStart) : 0;
        const bool readWholeFileForDiffProbe = diffByExtension && availableBytes <= maxParsedDiffBytes;
        const uint64_t wantBytes64           = readWholeFileForDiffProbe ? availableBytes : std::min<uint64_t>(availableBytes, maxChunkBytes);
        std::vector<uint8_t> bytes;

        const auto readBytesFromOffset = [&](const uint64_t bytesToRead) noexcept -> bool
        {
            const size_t nextWantBytes = static_cast<size_t>(std::min<uint64_t>(bytesToRead, static_cast<uint64_t>(std::numeric_limits<size_t>::max())));

            uint64_t nextPosition     = 0;
            const HRESULT seekBytesHr = result->fileReader->Seek(static_cast<__int64>(clampedStart), FILE_BEGIN, &nextPosition);
            if (FAILED(seekBytesHr))
            {
                Debug::Error(L"ViewerText: Seek to data start offset failed (0x{:016X}) for '{}' (hr=0x{:08X}).",
                             clampedStart,
                             path.c_str(),
                             static_cast<unsigned long>(seekBytesHr));
                result->hr = seekBytesHr;
                return false;
            }

            bytes.assign(nextWantBytes, 0u);
            size_t bytesReadTotal = 0u;
            while (bytesReadTotal < bytes.size())
            {
                const size_t remaining   = bytes.size() - bytesReadTotal;
                const unsigned long want = remaining > static_cast<size_t>(std::numeric_limits<unsigned long>::max())
                                               ? std::numeric_limits<unsigned long>::max()
                                               : static_cast<unsigned long>(remaining);

                unsigned long chunkRead = 0u;
                const HRESULT chunkHr   = result->fileReader->Read(bytes.data() + bytesReadTotal, want, &chunkRead);
                if (FAILED(chunkHr))
                {
                    Debug::Error(L"ViewerText: Read failed for '{}' at offset 0x{:016X} (hr=0x{:08X}).",
                                 path.c_str(),
                                 clampedStart + bytesReadTotal,
                                 static_cast<unsigned long>(chunkHr));
                    result->hr = chunkHr;
                    return false;
                }

                if (chunkRead == 0u)
                {
                    break;
                }

                bytesReadTotal += static_cast<size_t>(chunkRead);
            }

            bytes.resize(bytesReadTotal);
            return true;
        };

        if (clampedStart > static_cast<uint64_t>(std::numeric_limits<__int64>::max()))
        {
            Debug::Error(L"ViewerText: File is too large to open (start offset 0x{:016X} exceeds maximum supported offset).", clampedStart);
            result->hr = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            return;
        }

        if (! readBytesFromOffset(wantBytes64))
        {
            return;
        }

        ViewMode targetViewMode = desiredViewMode;
        if (targetViewMode == ViewMode::Text && allowHexFallback)
        {
            const bool unicodeDecode = (displayEncoding == FileEncoding::Utf16LE || displayEncoding == FileEncoding::Utf16BE ||
                                        displayEncoding == FileEncoding::Utf32LE || displayEncoding == FileEncoding::Utf32BE);
            if (! unicodeDecode && LooksLikeBinaryData(bytes.data(), bytes.size()))
            {
                targetViewMode = ViewMode::Hex;
            }
        }

        size_t carryBytes = 0;
        if (displayEncoding == FileEncoding::Utf16LE || displayEncoding == FileEncoding::Utf16BE)
        {
            carryBytes = bytes.size() % 2;
        }
        else if (displayEncoding == FileEncoding::Utf32LE || displayEncoding == FileEncoding::Utf32BE)
        {
            carryBytes = bytes.size() % 4;
        }
        else if (displayCodePage == CP_UTF8)
        {
            carryBytes = Utf8IncompleteTailSize(bytes.data(), bytes.size());
        }

        carryBytes                = std::min(carryBytes, bytes.size());
        const size_t convertBytes = bytes.size() - carryBytes;

        result->textBuffer.clear();
        if (convertBytes > 0)
        {
            if ((displayEncoding == FileEncoding::Utf16LE || displayEncoding == FileEncoding::Utf16BE) && (convertBytes % 2) == 0)
            {
                const size_t wcharCount = convertBytes / 2;
                result->textBuffer.resize(wcharCount);
                memcpy(result->textBuffer.data(), bytes.data(), convertBytes);

                if (displayEncoding == FileEncoding::Utf16BE)
                {
                    for (size_t i = 0; i < result->textBuffer.size(); ++i)
                    {
                        const wchar_t v       = result->textBuffer[i];
                        result->textBuffer[i] = static_cast<wchar_t>((static_cast<uint16_t>(v) >> 8) | (static_cast<uint16_t>(v) << 8));
                    }
                }
            }
            else if ((displayEncoding == FileEncoding::Utf32LE || displayEncoding == FileEncoding::Utf32BE) && (convertBytes % 4) == 0)
            {
                const bool bigEndian = (displayEncoding == FileEncoding::Utf32BE);
                result->textBuffer.reserve(convertBytes / 4);

                for (size_t i = 0; i + 3 < convertBytes; i += 4)
                {
                    uint32_t cp = 0;
                    if (bigEndian)
                    {
                        cp = (static_cast<uint32_t>(bytes[i]) << 24) | (static_cast<uint32_t>(bytes[i + 1]) << 16) |
                             (static_cast<uint32_t>(bytes[i + 2]) << 8) | static_cast<uint32_t>(bytes[i + 3]);
                    }
                    else
                    {
                        cp = static_cast<uint32_t>(bytes[i]) | (static_cast<uint32_t>(bytes[i + 1]) << 8) | (static_cast<uint32_t>(bytes[i + 2]) << 16) |
                             (static_cast<uint32_t>(bytes[i + 3]) << 24);
                    }

                    if (cp <= 0xFFFFu)
                    {
                        if (cp >= 0xD800u && cp <= 0xDFFFu)
                        {
                            result->textBuffer.push_back(static_cast<wchar_t>(0xFFFDu));
                        }
                        else
                        {
                            result->textBuffer.push_back(static_cast<wchar_t>(cp));
                        }
                    }
                    else if (cp <= 0x10FFFFu)
                    {
                        const uint32_t v = cp - 0x10000u;
                        result->textBuffer.push_back(static_cast<wchar_t>(0xD800u + (v >> 10)));
                        result->textBuffer.push_back(static_cast<wchar_t>(0xDC00u + (v & 0x3FFu)));
                    }
                    else
                    {
                        result->textBuffer.push_back(static_cast<wchar_t>(0xFFFDu));
                    }
                }
            }
            else
            {
                if (convertBytes > static_cast<size_t>(std::numeric_limits<int>::max()))
                {
                    Debug::Error(L"ViewerText: File is too large to open (data size 0x{:016X} exceeds maximum supported size).", convertBytes);
                    result->hr = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                    return;
                }

                const int srcLen       = static_cast<int>(convertBytes);
                const int requiredWide = MultiByteToWideChar(displayCodePage, 0, reinterpret_cast<LPCCH>(bytes.data()), srcLen, nullptr, 0);
                if (requiredWide <= 0)
                {
                    auto lastError = Debug::ErrorWithLastError(
                        L"ViewerText: MultiByteToWideChar failed to calculate required buffer size for '{}' (hr=0x{:08X}).", path.c_str(), result->hr);
                    result->hr = HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_INVALID_DATA);
                    return;
                }

                result->textBuffer.resize(static_cast<size_t>(requiredWide));
                const int written =
                    MultiByteToWideChar(displayCodePage, 0, reinterpret_cast<LPCCH>(bytes.data()), srcLen, result->textBuffer.data(), requiredWide);
                if (written <= 0)
                {
                    auto lastError =
                        Debug::ErrorWithLastError(L"ViewerText: MultiByteToWideChar failed to convert data for '{}' (hr=0x{:08X}).", path.c_str(), result->hr);
                    result->hr = HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_INVALID_DATA);
                    return;
                }

                result->textBuffer.resize(static_cast<size_t>(written));
            }
        }

        result->textStreamStartOffset = clampedStart;
        if (bytes.size() >= carryBytes)
        {
            const uint64_t consumed     = static_cast<uint64_t>(bytes.size() - carryBytes);
            result->textStreamEndOffset = std::min<uint64_t>(clampedStart + consumed, detectedFileSize);
        }
        else
        {
            result->textStreamEndOffset = clampedStart;
        }

        result->textStreamActive = (detectedFileSize > streamSkipBytes) && ((detectedFileSize - streamSkipBytes) > maxChunkBytes);

        const UINT defaultCodePage      = GetACP();
        result->detectedCodePage        = 0;
        result->detectedCodePageValid   = false;
        result->detectedCodePageIsGuess = false;

        switch (encoding)
        {
            case FileEncoding::Utf8:
                result->detectedCodePage        = CP_UTF8;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Utf16LE:
                result->detectedCodePage        = 1200u;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Utf16BE:
                result->detectedCodePage        = 1201u;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Utf32LE:
                result->detectedCodePage        = 12000u;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Utf32BE:
                result->detectedCodePage        = 12001u;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Unknown:
            default:
            {
                result->detectedCodePageIsGuess = true;
                if (! bytes.empty() && IsValidUtf8(bytes.data(), bytes.size()))
                {
                    result->detectedCodePage = CP_UTF8;
                }
                else
                {
                    result->detectedCodePage = defaultCodePage;
                }
                result->detectedCodePageValid = true;
                break;
            }
        }

        BuildTextLineIndex(result->textBuffer, result->textLineStarts, result->textLineEnds, result->textMaxLineLength);

        bool looksLikeDiffDocument            = diffByExtension || LooksLikeUnifiedDiffText(result->textBuffer);
        const bool canPromoteFullBufferedDiff = targetViewMode == ViewMode::Text && looksLikeDiffDocument && availableBytes <= maxParsedDiffBytes &&
                                                clampedStart == streamSkipBytes && result->textStreamEndOffset < detectedFileSize;
        if (canPromoteFullBufferedDiff && ! readWholeFileForDiffProbe)
        {
            if (! readBytesFromOffset(availableBytes))
            {
                return;
            }

            HRESULT decodeHr     = S_OK;
            std::wstring decoded = DecodeBytesToWide(bytes, displayEncoding, displayCodePage, decodeHr);
            if (FAILED(decodeHr))
            {
                result->hr = decodeHr;
                return;
            }

            result->textBuffer            = std::move(decoded);
            result->textStreamStartOffset = clampedStart;
            result->textStreamEndOffset   = detectedFileSize;
            result->textStreamActive      = false;
            BuildTextLineIndex(result->textBuffer, result->textLineStarts, result->textLineEnds, result->textMaxLineLength);
            looksLikeDiffDocument = diffByExtension || LooksLikeUnifiedDiffText(result->textBuffer);
        }
        if (looksLikeDiffDocument)
        {
            result->documentKind = DocumentKind::Diff;
        }

        if (targetViewMode == ViewMode::Text && looksLikeDiffDocument && availableBytes <= maxParsedDiffBytes && clampedStart == streamSkipBytes &&
            result->textStreamEndOffset >= detectedFileSize)
        {
            auto parsedDocument = std::make_shared<ParsedDiffDocument>();
            if (ParseUnifiedDiffDocument(result->textBuffer, *parsedDocument))
            {
                result->diffParsedAvailable     = true;
                result->parsedDiffDocument      = parsedDocument;
                result->diffInlineHunksOnly     = BuildInlineDiffText(*parsedDocument, DiffContextMode::HunksOnly, path, fileIo.get()).variant;
                result->diffSideBySideHunksOnly = BuildSideBySideDiffText(*parsedDocument, DiffContextMode::HunksOnly, path, fileIo.get()).variant;
#ifdef _DEBUG
                result->diffParseCount += 1u;
#endif
                result->initialDiffMode =
                    diffAutoOpenMode == DiffAutoOpenMode::RawText
                        ? DiffPresentationMode::RawText
                        : (diffDefaultLayout == DiffDefaultLayout::Inline ? DiffPresentationMode::Inline : DiffPresentationMode::SideBySide);

                if (diffAutoOpenMode == DiffAutoOpenMode::RawText)
                {
                    result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_SHOWING_RAW_TEXT);
                }

                result->textStreamActive = false;
            }
            else if (diffByExtension)
            {
                result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_PARSE_UNAVAILABLE);
            }
        }
        else if (diffByExtension && diffAutoOpenMode == DiffAutoOpenMode::Parsed)
        {
            result->documentKind  = DocumentKind::Diff;
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_DIFF_PARSE_UNAVAILABLE);
        }

        if (targetViewMode == ViewMode::Text && result->documentKind == DocumentKind::Diff && ! result->diffParsedAvailable && result->textStreamActive)
        {
            static_cast<void>(BuildBoundedStreamedDiffSectionIndex(
                result->fileReader.get(), detectedFileSize, streamSkipBytes, displayEncoding, displayCodePage, result->streamedDiffSections));
        }

        result->viewMode = targetViewMode;

        const bool needHex = (targetViewMode == ViewMode::Hex);
        if (needHex && detectedFileSize > 0)
        {
            if (detectedFileSize <= kMaxHexLoadBytes && detectedFileSize <= static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
            {
                result->hexBytes.resize(static_cast<size_t>(detectedFileSize));

                uint64_t ignored        = 0;
                const HRESULT seekHexHr = result->fileReader->Seek(0, FILE_BEGIN, &ignored);
                if (FAILED(seekHexHr))
                {
                    Debug::Warning(
                        L"ViewerText: Seek(FILE_BEGIN, 0) failed for HEX preload of '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(seekHexHr));
                    result->hexBytes.clear();
                }
                else
                {
                    size_t offset = 0;
                    while (offset < result->hexBytes.size())
                    {
                        const unsigned long want = static_cast<unsigned long>(std::min<size_t>(256 * 1024, result->hexBytes.size() - offset));
                        unsigned long readHex    = 0;
                        const HRESULT readHexHr  = result->fileReader->Read(result->hexBytes.data() + offset, want, &readHex);
                        if (FAILED(readHexHr))
                        {
                            Debug::Warning(
                                L"ViewerText: Read failed for HEX preload of '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(readHexHr));
                            result->hexBytes.clear();
                            break;
                        }
                        if (readHex == 0)
                        {
                            break;
                        }
                        offset += static_cast<size_t>(readHex);
                    }
                }
            }
            else
            {
                uint64_t cacheBytes = static_cast<uint64_t>(hexBufferMiB) * 1024u * 1024u;
                cacheBytes          = std::clamp<uint64_t>(cacheBytes, 256u * 1024u, 256u * 1024u * 1024u);

                const uint64_t remaining = detectedFileSize;
                const uint64_t want64    = std::min<uint64_t>(remaining, cacheBytes);
                const unsigned long want = want64 > static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()) ? std::numeric_limits<unsigned long>::max()
                                                                                                                     : static_cast<unsigned long>(want64);

                result->hexCacheOffset = 0;
                result->hexCacheValid  = 0;
                if (want > 0)
                {
                    result->hexCache.resize(static_cast<size_t>(want));

                    uint64_t ignored        = 0;
                    const HRESULT seekHexHr = result->fileReader->Seek(0, FILE_BEGIN, &ignored);
                    if (FAILED(seekHexHr))
                    {
                        Debug::Warning(L"ViewerText: Seek(FILE_BEGIN, 0) failed for HEX cache preload of '{}' (hr=0x{:08X}).",
                                       path.c_str(),
                                       static_cast<unsigned long>(seekHexHr));
                        result->hexCache.clear();
                    }
                    else
                    {
                        unsigned long readHex   = 0;
                        const HRESULT readHexHr = result->fileReader->Read(result->hexCache.data(), want, &readHex);
                        if (FAILED(readHexHr))
                        {
                            Debug::Warning(
                                L"ViewerText: Read failed for HEX cache preload of '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(readHexHr));
                            result->hexCache.clear();
                        }
                        else
                        {
                            result->hasHexCache   = true;
                            result->hexCacheValid = static_cast<size_t>(readHex);
                        }
                    }
                }
            }
        }

        result->hr = S_OK;

        if (FAILED(result->hr) && result->hr != E_OUTOFMEMORY && allowHexFallback && result->fileReader && result->fileSize > 0)
        {
            result->hexBytes.clear();
            result->hexCache.clear();
            result->hexCacheOffset = 0;
            result->hexCacheValid  = 0;
            result->hasHexCache    = false;

            const uint64_t hexFallbackSize = result->fileSize;
            if (hexFallbackSize <= kMaxHexLoadBytes && hexFallbackSize <= static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
            {
                result->hexBytes.resize(static_cast<size_t>(hexFallbackSize));

                uint64_t ignored        = 0;
                const HRESULT seekHexHr = result->fileReader->Seek(0, FILE_BEGIN, &ignored);
                if (FAILED(seekHexHr))
                {
                    result->hexBytes.clear();
                    Debug::Error(
                        L"ViewerText: Seek(FILE_BEGIN, 0) failed for HEX fallback of '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(seekHexHr));
                    return;
                }

                size_t offset = 0;
                while (offset < result->hexBytes.size())
                {
                    const unsigned long want = static_cast<unsigned long>(std::min<size_t>(256 * 1024, result->hexBytes.size() - offset));
                    unsigned long readHex    = 0;
                    const HRESULT readHexHr  = result->fileReader->Read(result->hexBytes.data() + offset, want, &readHex);
                    if (FAILED(readHexHr))
                    {
                        result->hexBytes.clear();
                        Debug::Error(L"ViewerText: Read failed for HEX fallback of '{}' at offset 0x{:016X} (hr=0x{:08X}).",
                                     path.c_str(),
                                     offset,
                                     static_cast<unsigned long>(readHexHr));
                        return;
                    }
                    if (readHex == 0)
                    {
                        break;
                    }
                    offset += static_cast<size_t>(readHex);
                }
            }
            else
            {
                uint64_t cacheBytes = static_cast<uint64_t>(hexBufferMiB) * 1024u * 1024u;
                cacheBytes          = std::clamp<uint64_t>(cacheBytes, 256u * 1024u, 256u * 1024u * 1024u);

                const uint64_t aligned   = 0;
                const uint64_t remaining = (detectedFileSize > aligned) ? (detectedFileSize - aligned) : 0;
                const uint64_t want64    = std::min<uint64_t>(remaining, cacheBytes);
                const unsigned long want = want64 > static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()) ? std::numeric_limits<unsigned long>::max()
                                                                                                                     : static_cast<unsigned long>(want64);

                result->hexCacheOffset = aligned;
                if (want > 0)
                {
                    result->hexCache.resize(static_cast<size_t>(want));

                    uint64_t ignored        = 0;
                    const HRESULT seekHexHr = result->fileReader->Seek(static_cast<__int64>(aligned), FILE_BEGIN, &ignored);
                    if (FAILED(seekHexHr))
                    {
                        Debug::Error(L"ViewerText: Seek to offset 0x{:016X} failed for HEX cache fallback of '{}' (hr=0x{:08X}).",
                                     aligned,
                                     path.c_str(),
                                     static_cast<unsigned long>(seekHexHr));
                        return;
                    }

                    unsigned long readHex   = 0;
                    const HRESULT readHexHr = result->fileReader->Read(result->hexCache.data(), want, &readHex);
                    if (FAILED(readHexHr))
                    {
                        Debug::Error(L"ViewerText: Read failed for HEX cache fallback of '{}' at offset 0x{:016X} (hr=0x{:08X}).",
                                     path.c_str(),
                                     aligned,
                                     static_cast<unsigned long>(readHexHr));
                        return;
                    }

                    result->hasHexCache   = true;
                    result->hexCacheValid = static_cast<size_t>(readHex);
                }
            }

            Debug::Warning(
                L"ViewerText: Failed to load '{}' as text (hr=0x{:08X}); falling back to HEX view.", path.c_str(), static_cast<unsigned long>(result->hr));
            result->viewMode = ViewMode::Hex;
            result->hr       = S_OK;
        }

        if (! hwnd || GetWindowLongPtrW(hwnd, GWLP_USERDATA) != reinterpret_cast<LONG_PTR>(this))
        {
            return;
        }

        static_cast<void>(PostMessagePayload(hwnd, kAsyncOpenCompleteMessage, 0, std::move(result)));
    };

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
    {
        std::unique_ptr<AsyncOpenWorkItem> ctx(static_cast<AsyncOpenWorkItem*>(context));
        if (! ctx)
        {
            return;
        }

        static_cast<void>(ctx->moduleKeepAlive);
        if (ctx->work)
        {
            ctx->work();
        }
    },
        ctx.get(),
        nullptr);

    if (queued == 0)
    {
        Debug::Error(L"ViewerText: Failed to queue async open work item for '{}'.", path.c_str());
        return;
    }

    ctx.release();
}

void ViewerText::OnAsyncOpenComplete(std::unique_ptr<AsyncOpenResult> result) noexcept
{
    if (! result)
    {
        return;
    }
    if (result->viewer != this)
    {
        return;
    }

    if (result->requestId != _activeAsyncOpenRequestId)
    {
        return;
    }

    EndLoadingUi();

    if (FAILED(result->hr))
    {
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_ERR_OPEN_FAILED);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, TRUE);
        }

        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_OPEN_FAILED);
        return;
    }

    _fileReader                   = std::move(result->fileReader);
    _fileSize                     = result->fileSize;
    _encoding                     = result->encoding;
    _bomBytes                     = result->bomBytes;
    _displayEncodingMenuSelection = result->displayEncodingMenuSelection;
    _detectedCodePage             = result->detectedCodePage;
    _detectedCodePageValid        = result->detectedCodePageValid;
    _detectedCodePageIsGuess      = result->detectedCodePageIsGuess;

    _statusMessage       = result->statusMessage;
    _documentKind        = result->documentKind;
    _diffParsedAvailable = result->diffParsedAvailable;
    _diffPresentation    = result->initialDiffMode;
    _lastParsedDiffPresentation =
        result->initialDiffMode == DiffPresentationMode::RawText
            ? (_config.diffDefaultLayout == DiffDefaultLayout::Inline ? DiffPresentationMode::Inline : DiffPresentationMode::SideBySide)
            : result->initialDiffMode;
    _parsedDiffDocument = std::move(result->parsedDiffDocument);
    _diffReferenceCache.reset();
    _diffStreamSections          = std::move(result->streamedDiffSections);
    _diffInlineHunksOnly         = std::move(result->diffInlineHunksOnly);
    _diffInlineExpanded          = std::move(result->diffInlineExpanded);
    _diffSideBySideHunksOnly     = std::move(result->diffSideBySideHunksOnly);
    _diffSideBySideExpanded      = std::move(result->diffSideBySideExpanded);
    _diffInlineExpandedBuilt     = false;
    _diffSideBySideExpandedBuilt = false;
#ifdef _DEBUG
    _debugDiffParseCount = result->diffParseCount;
#endif

    _textStreamSkipBytes   = result->textStreamSkipBytes;
    _textStreamStartOffset = result->textStreamStartOffset;
    _textStreamEndOffset   = result->textStreamEndOffset;
    _textStreamActive      = result->textStreamActive;

    _textTotalLineCount.reset();
    _textStreamLineCountedEndOffset = _textStreamStartOffset;
    _textStreamLineCountedNewlines  = 0;
    _textStreamLineCountLastWasCR   = false;

    if (_documentKind == DocumentKind::Diff && _diffParsedAvailable)
    {
        _diffRawTextBuffer = std::move(result->textBuffer);
        ApplyCurrentTextPresentation(_hEdit.get());
    }
    else
    {
        _diffRawTextBuffer.clear();

        _textBuffer        = std::move(result->textBuffer);
        _textLineStarts    = std::move(result->textLineStarts);
        _textLineEnds      = std::move(result->textLineEnds);
        _textMaxLineLength = result->textMaxLineLength;

        UpdateTextStreamTotalLineCountAfterLoad();

        _textVisualLineStarts.clear();
        _textVisualLineLogical.clear();
        _textTopVisualLine   = 0;
        _textLeftColumn      = 0;
        _textCaretIndex      = 0;
        _textSelAnchor       = 0;
        _textSelActive       = 0;
        _textPreferredColumn = 0;
        _textSelecting       = false;
        _searchMatchStarts.clear();

        if (_hEdit)
        {
            RebuildTextVisualLines(_hEdit.get());
            UpdateTextViewScrollBars(_hEdit.get());
            UpdateSearchHighlights();
            InvalidateRect(_hEdit.get(), nullptr, TRUE);
        }
    }

    _hexBytes = std::move(result->hexBytes);
    if (result->hasHexCache)
    {
        _hexCache       = std::move(result->hexCache);
        _hexCacheOffset = result->hexCacheOffset;
        _hexCacheValid  = result->hexCacheValid;
    }

    if (_hHex)
    {
        UpdateHexViewScrollBars(_hHex.get());
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }

    if (_hWnd)
    {
        SetViewMode(_hWnd.get(), result->viewMode);
    }
}

void ViewerText::BeginLoadingUi() noexcept
{
    ClearInlineAlert();

    _isLoading                = true;
    _showLoadingOverlay       = false;
    _loadingSpinnerAngleDeg   = 0.0f;
    _loadingSpinnerLastTickMs = GetTickCount64();

    _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_LOADING);

    if (! _hWnd)
    {
        return;
    }

    KillTimer(_hWnd.get(), kLoadingDelayTimerId);
    KillTimer(_hWnd.get(), kLoadingAnimTimerId);
    SetTimer(_hWnd.get(), kLoadingDelayTimerId, kLoadingDelayMs, nullptr);
}

void ViewerText::EndLoadingUi() noexcept
{
    if (_hWnd)
    {
        KillTimer(_hWnd.get(), kLoadingDelayTimerId);
        KillTimer(_hWnd.get(), kLoadingAnimTimerId);
    }

    _isLoading          = false;
    _showLoadingOverlay = false;
}

void ViewerText::UpdateLoadingSpinner() noexcept
{
    if (! _isLoading || ! _showLoadingOverlay)
    {
        return;
    }

    const ULONGLONG now       = GetTickCount64();
    const ULONGLONG last      = _loadingSpinnerLastTickMs;
    _loadingSpinnerLastTickMs = now;

    double deltaSec = 0.0;
    if (now > last)
    {
        deltaSec = static_cast<double>(now - last) / 1000.0;
    }

    _loadingSpinnerAngleDeg += static_cast<float>(deltaSec * static_cast<double>(kLoadingSpinnerDegPerSec));
    while (_loadingSpinnerAngleDeg >= 360.0f)
    {
        _loadingSpinnerAngleDeg -= 360.0f;
    }

    if (_viewMode == ViewMode::Text && _hEdit)
    {
        InvalidateRect(_hEdit.get(), nullptr, FALSE);
    }
    else if (_viewMode == ViewMode::Hex && _hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, FALSE);
    }
}

void ViewerText::DrawLoadingOverlay(ID2D1HwndRenderTarget* target, ID2D1SolidColorBrush* brush, float widthDip, float heightDip) noexcept
{
    if (! _isLoading || ! _showLoadingOverlay || ! target || ! brush)
    {
        return;
    }

    if (widthDip <= 0.0f || heightDip <= 0.0f)
    {
        return;
    }

    const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

    const std::wstring seed = _currentPath.empty() ? std::wstring(L"viewer") : _currentPath.filename().wstring();
    const COLORREF accent   = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);

    if (! (_hasTheme && _theme.highContrast))
    {
        const uint8_t tintAlpha = (_hasTheme && _theme.darkMode) ? 28u : 18u;
        const COLORREF tint     = BlendColor(bg, accent, tintAlpha);
        const float overlayA    = (_hasTheme && _theme.darkMode) ? 0.85f : 0.75f;
        brush->SetColor(ColorFFromColorRef(tint, overlayA));
        target->FillRectangle(D2D1::RectF(0.0f, 0.0f, widthDip, heightDip), brush);
    }

    const float minDim = std::min(widthDip, heightDip);
    const float radius = std::clamp(minDim * 0.08f, 18.0f, 44.0f);
    const float stroke = std::clamp(radius * 0.20f, 3.0f, 6.0f);
    const float innerR = radius * 0.55f;
    const float outerR = radius;

    const float textHeightDip  = 34.0f;
    const float spacingDip     = 14.0f;
    const float groupHeightDip = outerR * 2.0f + spacingDip + textHeightDip;
    const float groupTopDip    = std::max(0.0f, (heightDip - groupHeightDip) * 0.5f);

    const float cx = widthDip * 0.5f;
    const float cy = groupTopDip + outerR;

    constexpr int kSegments = 12;
    constexpr float kPi     = 3.14159265358979323846f;
    const float baseRad     = (_loadingSpinnerAngleDeg - 90.0f) * (kPi / 180.0f);

    const bool rainbowSpinner = _hasTheme && ! _theme.highContrast && _theme.rainbowMode;
    float rainbowHue          = 0.0f;
    float rainbowSat          = 0.0f;
    float rainbowVal          = 0.0f;
    if (rainbowSpinner)
    {
        const uint32_t h = StableHash32(seed);
        rainbowHue       = static_cast<float>(h % 360u);
        rainbowSat       = _theme.darkBase ? 0.70f : 0.55f;
        rainbowVal       = _theme.darkBase ? 0.95f : 0.85f;
    }

    for (int i = 0; i < kSegments; ++i)
    {
        const float t     = static_cast<float>(i) / static_cast<float>(kSegments);
        const float alpha = 0.15f + 0.85f * (1.0f - t);
        const float angle = baseRad + t * (2.0f * kPi);
        const float s     = std::sin(angle);
        const float c     = std::cos(angle);

        const D2D1_POINT_2F p1 = D2D1::Point2F(cx + c * innerR, cy + s * innerR);
        const D2D1_POINT_2F p2 = D2D1::Point2F(cx + c * outerR, cy + s * outerR);

        COLORREF segmentColor = accent;
        if (rainbowSpinner)
        {
            const float hueStep    = 360.0f / static_cast<float>(kSegments);
            const float hueDegrees = rainbowHue + static_cast<float>(i) * hueStep;
            segmentColor           = ColorFromHSV(hueDegrees, rainbowSat, rainbowVal);
        }

        brush->SetColor(ColorFFromColorRef(segmentColor, alpha));
        target->DrawLine(p1, p2, brush, stroke);
    }

    std::wstring loadingText = _statusMessage;
    if (loadingText.empty())
    {
        loadingText = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_LOADING);
    }

    if (loadingText.empty())
    {
        return;
    }

    if (! _loadingOverlayFormat && _dwriteFactory)
    {
        wil::com_ptr<IDWriteTextFormat> format;
        const HRESULT hr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(22.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), format.put());
        if (SUCCEEDED(hr) && format)
        {
            static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
            static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
            _loadingOverlayFormat = std::move(format);
        }
    }

    if (! _loadingOverlayFormat)
    {
        return;
    }

    const float textTopDip   = groupTopDip + outerR * 2.0f + spacingDip;
    const D2D1_RECT_F textRc = D2D1::RectF(0.0f, textTopDip, widthDip, std::min(heightDip, textTopDip + textHeightDip));

    brush->SetColor(ColorFFromColorRef(fg, 0.90f));
    const UINT32 len = static_cast<UINT32>(std::min<size_t>(loadingText.size(), std::numeric_limits<UINT32>::max()));
    target->DrawTextW(loadingText.c_str(), len, _loadingOverlayFormat.get(), textRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
}

void ViewerText::ShowInlineAlert(InlineAlertSeverity severity, UINT titleId, UINT messageId) noexcept
{
    if (! _hostAlerts)
    {
        return;
    }

    const std::wstring title   = LoadStringResource(g_hInstance, titleId);
    const std::wstring message = LoadStringResource(g_hInstance, messageId);
    if (message.empty())
    {
        return;
    }

    HostAlertSeverity hostSeverity = HOST_ALERT_ERROR;
    switch (severity)
    {
        case InlineAlertSeverity::Warning: hostSeverity = HOST_ALERT_WARNING; break;
        case InlineAlertSeverity::Info: hostSeverity = HOST_ALERT_INFO; break;
        case InlineAlertSeverity::Error:
        default: hostSeverity = HOST_ALERT_ERROR; break;
    }

    HWND targetWindow = nullptr;
    if (_viewMode == ViewMode::Hex && _hHex)
    {
        targetWindow = _hHex.get();
    }
    else if (_hEdit)
    {
        targetWindow = _hEdit.get();
    }
    else if (_hWnd)
    {
        targetWindow = _hWnd.get();
    }

    if (! targetWindow)
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODAL;
    request.severity     = hostSeverity;
    request.targetWindow = targetWindow;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(_hostAlerts->ShowAlert(&request, nullptr));
}

void ViewerText::ClearInlineAlert() noexcept
{
    if (! _hostAlerts)
    {
        return;
    }

    if (_hEdit)
    {
        static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, reinterpret_cast<void*>(_hEdit.get())));
    }
    if (_hHex)
    {
        static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, reinterpret_cast<void*>(_hHex.get())));
    }
    if (_hWnd)
    {
        static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, reinterpret_cast<void*>(_hWnd.get())));
    }
}

void ViewerText::OnTimer(UINT_PTR timerId) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (timerId == kLoadingDelayTimerId)
    {
        KillTimer(_hWnd.get(), kLoadingDelayTimerId);
        if (! _isLoading)
        {
            return;
        }

        _showLoadingOverlay       = true;
        _loadingSpinnerAngleDeg   = 0.0f;
        _loadingSpinnerLastTickMs = GetTickCount64();
        SetTimer(_hWnd.get(), kLoadingAnimTimerId, kLoadingAnimIntervalMs, nullptr);

        if (_hEdit)
        {
            InvalidateRect(_hEdit.get(), nullptr, FALSE);
        }
        if (_hHex)
        {
            InvalidateRect(_hHex.get(), nullptr, FALSE);
        }
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
        return;
    }

    if (timerId == kLoadingAnimTimerId)
    {
        UpdateLoadingSpinner();
        return;
    }
}

void ViewerText::OnSize(UINT width, UINT height)
{
    if (! _hWnd)
    {
        return;
    }

    if (_d2dTarget && width > 0 && height > 0)
    {
        const HRESULT hr = _d2dTarget->Resize(D2D1::SizeU(width, height));
        if (FAILED(hr))
        {
            DiscardDirect2D();
        }
    }

    Layout(_hWnd.get());
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void ViewerText::OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggested) noexcept
{
    if (! hwnd)
    {
        return;
    }

    static_cast<void>(newDpi);

    if (suggested)
    {
        const int width  = std::max(1L, suggested->right - suggested->left);
        const int height = std::max(1L, suggested->bottom - suggested->top);
        SetWindowPos(hwnd, nullptr, suggested->left, suggested->top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    if (_hFileComboHost)
    {
        _fileComboHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
    }

    UpdateHexColumns(hwnd);
    Layout(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
}

void ViewerText::Layout(HWND hwnd) noexcept
{
    RECT client{};
    GetClientRect(hwnd, &client);
    _menuBarHost.UpdateLayout();
    client.top += _menuBarHost.GetHwnd() ? _menuBarHost.GetVisibleHeightPx() : 0;

    const UINT dpi                   = GetDpiForWindow(hwnd);
    const bool showStandaloneHeader  = ! _embeddedMode;
    const int baseHeaderHeight       = showStandaloneHeader ? PxFromDip(kHeaderHeightDip, dpi) : 0;
    const int statusHeight     = PxFromDip(kStatusHeightDip, dpi);
    const int accentHeight     = std::max(1, PxFromDip(2, dpi));
    const int accentGap        = std::max(1, PxFromDip(1, dpi));
    const int minPadding       = PxFromDip(3, dpi);
    const int minChromeHeight  = showStandaloneHeader ? PxFromDip(22, dpi) + accentHeight + accentGap + 2 * minPadding : 0;

    const bool showCombo         = (showStandaloneHeader && _hFileComboHost && ActiveFileComboEntryCount() > 1u);
    const int desiredComboHeight = showCombo ? std::max(1, PxFromDip(32, dpi)) : 0;

    int headerHeight = baseHeaderHeight;
    headerHeight     = std::max(headerHeight, minChromeHeight);
    if (showCombo)
    {
        headerHeight = std::max(headerHeight, desiredComboHeight + accentHeight + accentGap + 2 * minPadding);
    }

    for (int pass = 0; pass < 2; ++pass)
    {
        _headerRect        = client;
        _headerRect.bottom = std::min(client.bottom, client.top + std::max(0, headerHeight));

        _statusRect     = client;
        _statusRect.top = std::max(client.top, client.bottom - std::max(0, statusHeight));

        _contentRect        = client;
        _contentRect.top    = _headerRect.bottom;
        _contentRect.bottom = _statusRect.top;

        ClampRectNonNegative(_headerRect);
        ClampRectNonNegative(_statusRect);
        ClampRectNonNegative(_contentRect);

        RECT headerContentRect{};
        headerContentRect        = _headerRect;
        headerContentRect.top    = std::min(headerContentRect.bottom, headerContentRect.top + minPadding);
        headerContentRect.bottom = std::max(headerContentRect.top, headerContentRect.bottom - accentHeight - accentGap - minPadding);

        const int headerContentH = std::max(0L, headerContentRect.bottom - headerContentRect.top);
        const int margin         = PxFromDip(10, dpi);
        if (showStandaloneHeader)
        {
            const int buttonH      = std::min(headerContentH, PxFromDip(22, dpi));
            const int buttonW      = PxFromDip(72, dpi);
            const int buttonY      = headerContentRect.top + std::max(0, (headerContentH - buttonH) / 2);
            const int buttonX      = std::max<LONG>(headerContentRect.left, headerContentRect.right - margin - buttonW);
            _modeButtonRect.left   = buttonX;
            _modeButtonRect.top    = buttonY;
            _modeButtonRect.right  = std::min<LONG>(headerContentRect.right, buttonX + buttonW);
            _modeButtonRect.bottom = std::min<LONG>(headerContentRect.bottom, buttonY + buttonH);
        }
        else
        {
            _modeButtonRect = {};
        }

        int measuredComboHeight = 0;
        if (_hFileComboHost)
        {
            ShowWindow(_hFileComboHost.get(), showCombo ? SW_SHOW : SW_HIDE);
            EnableWindow(_hFileComboHost.get(), showCombo ? TRUE : FALSE);
            if (! showCombo)
            {
                _fileComboHostPreExpandPopup = false;
            }

            if (showCombo)
            {
                int comboH = desiredComboHeight;
                comboH     = std::clamp(comboH, 1, std::max(1, headerContentH));

                const int comboX    = headerContentRect.left + margin;
                const int comboW    = std::max(0, static_cast<int>(_modeButtonRect.left) - margin - comboX);
                measuredComboHeight = comboH;
                int comboY          = headerContentRect.top + std::max(0, (headerContentH - comboH) / 2);

                const int maxBottom = std::max(static_cast<int>(headerContentRect.top), static_cast<int>(headerContentRect.bottom));
                if (comboY + comboH > maxBottom)
                {
                    comboY = std::max(static_cast<int>(headerContentRect.top), maxBottom - comboH);
                }

                const bool expandPopupHost = _fileComboHostPreExpandPopup || (_fileComboControl && _fileComboControl->DebugIsPopupOpen());
                const int popupExtraHeight = expandPopupHost ? ComputeWindowComboPopupHeightPx(ActiveFileComboEntryCount(), dpi) : 0;
                const int hostHeight       = std::max(comboH, comboH + popupExtraHeight);
                SetWindowPos(_hFileComboHost.get(), HWND_TOP, comboX, comboY, comboW, hostHeight, SWP_NOACTIVATE);
                if (_fileComboControl)
                {
                    _fileComboControl->SetBounds(D2D1::RectF(0.0f,
                                                             0.0f,
                                                             static_cast<float>(comboW) * 96.0f / static_cast<float>(dpi),
                                                             static_cast<float>(comboH) * 96.0f / static_cast<float>(dpi)));
                    _fileComboHost.Invalidate();
                }
            }
        }

        const int currentHeaderHeight = headerHeight;
        int requiredHeaderHeight      = currentHeaderHeight;
        if (showCombo && measuredComboHeight > 0)
        {
            requiredHeaderHeight = std::max(minChromeHeight, measuredComboHeight + accentHeight + accentGap + 2 * minPadding);
        }

        if (requiredHeaderHeight > currentHeaderHeight && pass == 0)
        {
            headerHeight = requiredHeaderHeight;
            continue;
        }

        break;
    }

    const int contentW = std::max(0L, _contentRect.right - _contentRect.left);
    const int contentH = std::max(0L, _contentRect.bottom - _contentRect.top);

    if (_hEdit)
    {
        SetWindowPos(_hEdit.get(), nullptr, _contentRect.left, _contentRect.top, contentW, contentH, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (_hHex)
    {
        SetWindowPos(_hHex.get(), nullptr, _contentRect.left, _contentRect.top, contentW, contentH, SWP_NOZORDER | SWP_NOACTIVATE);
    }
}

void ViewerText::RefreshFileCombo(HWND hwnd) noexcept
{
    if (! _fileComboControl)
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    const bool useDiffSections = UseDiffSectionFileCombo();
    const size_t entryCount    = ActiveFileComboEntryCount();
    if (entryCount <= 1u)
    {
        _fileComboControl->SetItems({});
        _fileComboControl->SetSelectedIndex(std::nullopt);
        _lastSyncedFileComboIndex.reset();
        if (hwnd)
        {
            Layout(hwnd);
            InvalidateRect(hwnd, nullptr, TRUE);
        }
        return;
    }

    std::vector<ComboBox::Item> items;
    items.reserve(entryCount);
    if (useDiffSections)
    {
        if (const auto* variant = CurrentDiffVariant())
        {
            for (const auto& section : variant->sectionNavigation)
            {
                items.push_back(ComboBox::Item{section.label, section.label});
            }
        }
        else
        {
            for (const auto& section : _diffStreamSections)
            {
                items.push_back(ComboBox::Item{section.label, section.label});
            }
        }
    }
    else
    {
        for (const auto& path : _otherFiles)
        {
            std::wstring itemText = path.filename().wstring();
            if (itemText.empty())
            {
                itemText = path.wstring();
            }
            items.push_back(ComboBox::Item{path.wstring(), std::move(itemText)});
        }
    }
    _fileComboControl->SetItems(std::move(items));

    size_t selectedIndex = 0u;
    if (useDiffSections)
    {
        selectedIndex = CurrentDiffSectionIndex();
        _fileComboControl->SetSelectedIndex(selectedIndex);
    }
    else
    {
        if (_otherIndex >= _otherFiles.size())
        {
            _otherIndex = 0;
        }

        selectedIndex = _otherIndex;
        _fileComboControl->SetSelectedIndex(selectedIndex);
    }
    _lastSyncedFileComboUsesDiffSections = useDiffSections;
    _lastSyncedFileComboIndex            = selectedIndex;
    _fileComboHost.Invalidate();

    if (hwnd)
    {
        Layout(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
    }
}

void ViewerText::SyncFileComboSelection() noexcept
{
    if (! _fileComboControl)
    {
        return;
    }

    const bool useDiffSections = UseDiffSectionFileCombo();
    const size_t entryCount    = ActiveFileComboEntryCount();
    if (entryCount <= 1u)
    {
        return;
    }

    size_t selectedIndex = 0u;
    if (useDiffSections)
    {
        selectedIndex = CurrentDiffSectionIndex();
    }
    else
    {
        if (_otherIndex >= _otherFiles.size())
        {
            return;
        }

        selectedIndex = _otherIndex;
    }

    if (_lastSyncedFileComboIndex.has_value() && _lastSyncedFileComboUsesDiffSections == useDiffSections && _lastSyncedFileComboIndex.value() == selectedIndex)
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });
    _fileComboControl->SetSelectedIndex(selectedIndex);
    _lastSyncedFileComboUsesDiffSections = useDiffSections;
    _lastSyncedFileComboIndex            = selectedIndex;
    _fileComboHost.Invalidate();
}

bool ViewerText::EnsureDirect2D(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    const UINT dpi   = GetDpiForWindow(hwnd);
    const float dpiF = static_cast<float>(dpi);

    if (! _d2dFactory)
    {
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, _d2dFactory.put());
        if (FAILED(hr) || ! _d2dFactory)
        {
            _d2dFactory.reset();
            return false;
        }
    }

    if (! _dwriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.put()));
        if (FAILED(hr) || ! _dwriteFactory)
        {
            _dwriteFactory.reset();
            return false;
        }
    }

    if (! _headerFormat)
    {
        const HRESULT hr =
            Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(12.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _headerFormat.put());
        if (FAILED(hr) || ! _headerFormat)
        {
            _headerFormat.reset();
            return false;
        }

        static_cast<void>(_headerFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
        static_cast<void>(_headerFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    if (! _headerFormatRight)
    {
        const HRESULT hr =
            Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(12.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _headerFormatRight.put());
        if (FAILED(hr) || ! _headerFormatRight)
        {
            _headerFormatRight.reset();
            return false;
        }

        static_cast<void>(_headerFormatRight->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING));
        static_cast<void>(_headerFormatRight->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    if (! _modeButtonFormat)
    {
        const HRESULT hr =
            Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(12.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _modeButtonFormat.put());
        if (FAILED(hr) || ! _modeButtonFormat)
        {
            _modeButtonFormat.reset();
            return false;
        }

        static_cast<void>(_modeButtonFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
        static_cast<void>(_modeButtonFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    if (! _statusFormat)
    {
        const HRESULT hr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(11.0f), _statusFormat.put());
        if (FAILED(hr) || ! _statusFormat)
        {
            _statusFormat.reset();
            return false;
        }

        static_cast<void>(_statusFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
        static_cast<void>(_statusFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    if (! _watermarkFormat)
    {
        const HRESULT hr = Typography::CreateTextFormat(
            _dwriteFactory.get(), Typography::MakeUiTextSpec(kWatermarkFontSizeDip, DWRITE_FONT_WEIGHT_SEMI_BOLD), _watermarkFormat.put());
        if (FAILED(hr) || ! _watermarkFormat)
        {
            _watermarkFormat.reset();
            return false;
        }

        static_cast<void>(_watermarkFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
        static_cast<void>(_watermarkFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    if (! _d2dTarget)
    {
        RECT client{};
        GetClientRect(hwnd, &client);

        const UINT32 width     = static_cast<UINT32>(std::max<LONG>(0, client.right - client.left));
        const UINT32 height    = static_cast<UINT32>(std::max<LONG>(0, client.bottom - client.top));
        const D2D1_SIZE_U size = D2D1::SizeU(width, height);

        D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        props.dpiX                          = dpiF;
        props.dpiY                          = dpiF;

        const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, size);

        const HRESULT hr = _d2dFactory->CreateHwndRenderTarget(props, hwndProps, _d2dTarget.put());
        if (FAILED(hr) || ! _d2dTarget)
        {
            _d2dTarget.reset();
            _d2dBrush.reset();
            return false;
        }

        _d2dTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }
    else
    {
        _d2dTarget->SetDpi(dpiF, dpiF);
    }

    if (! _d2dBrush)
    {
        const HRESULT hr = _d2dTarget->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::Black), _d2dBrush.put());
        if (FAILED(hr) || ! _d2dBrush)
        {
            _d2dBrush.reset();
            return false;
        }
    }

    return _d2dTarget && _d2dBrush && _headerFormat && _headerFormatRight && _statusFormat && _watermarkFormat;
}

void ViewerText::DiscardDirect2D() noexcept
{
    _d2dBrush.reset();
    _headerFormat.reset();
    _headerFormatRight.reset();
    _modeButtonFormat.reset();
    _statusFormat.reset();
    _watermarkFormat.reset();
    _d2dTarget.reset();
}

void ViewerText::OnPaint()
{
    if (! _hWnd)
    {
        return;
    }

    Debug::Perf::Scope paintPerf(L"viewer.chrome.paint_us");

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd.get(), &ps);
    _allowEraseBkgnd          = false;

    const UINT dpi   = GetDpiForWindow(_hWnd.get());
    const int dpiInt = static_cast<int>(dpi);

    const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

    COLORREF headerBg = bg;
    COLORREF statusBg = bg;
    if (_hasTheme && _theme.darkMode)
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 10), std::max(0, GetGValue(bg) - 10), std::max(0, GetBValue(bg) - 10));
        statusBg = RGB(std::min(255, GetRValue(bg) + 5), std::min(255, GetGValue(bg) + 5), std::min(255, GetBValue(bg) + 5));
    }
    else
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 5), std::max(0, GetGValue(bg) - 5), std::max(0, GetBValue(bg) - 5));
        statusBg = RGB(std::min(255, GetRValue(bg) + 5), std::min(255, GetGValue(bg) + 5), std::min(255, GetBValue(bg) + 5));
    }

    const std::wstring seed = _currentPath.empty() ? std::wstring(L"viewer") : _currentPath.filename().wstring();
    const COLORREF accent   = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);

    std::wstring titleText;
    if (! _currentPath.empty())
    {
        titleText = _currentPath.filename().wstring();
    }

    UINT modeId = IDS_VIEWERTEXT_MODE_TEXT;
    if (_viewMode == ViewMode::Hex)
    {
        modeId = IDS_VIEWERTEXT_MODE_HEX;
    }
    else if (_documentKind == DocumentKind::Diff)
    {
        modeId = _diffPresentation == DiffPresentationMode::RawText ? IDS_VIEWERTEXT_MODE_RAW : IDS_VIEWERTEXT_MODE_DIFF;
    }
    const std::wstring modeText   = LoadStringResource(g_hInstance, modeId);
    const std::wstring statusText = BuildStatusText();

    const auto drawChrome = [&](ID2D1RenderTarget* target, ID2D1SolidColorBrush* brush) noexcept
    {
        if (! target || ! brush)
        {
            return;
        }

        target->SetTransform(D2D1::Matrix3x2F::Identity());
        target->Clear(ColorFFromColorRef(bg));

        const D2D1_RECT_F headerRc = RectFFromPixels(_headerRect, dpi);
        const D2D1_RECT_F statusRc = RectFFromPixels(_statusRect, dpi);

        brush->SetColor(ColorFFromColorRef(headerBg));
        target->FillRectangle(headerRc, brush);

        brush->SetColor(ColorFFromColorRef(statusBg));
        target->FillRectangle(statusRc, brush);

        const int accentHeightPx = std::max(1, PxFromDip(2, dpi));
        RECT accentPx            = _headerRect;
        accentPx.top             = std::max(accentPx.top, accentPx.bottom - accentHeightPx);
        ClampRectNonNegative(accentPx);
        const D2D1_RECT_F accentRc = RectFFromPixels(accentPx, dpi);

        brush->SetColor(ColorFFromColorRef(accent));
        target->FillRectangle(accentRc, brush);

        const float marginDip    = 10.0f;
        D2D1_RECT_F headerTextRc = headerRc;
        headerTextRc.left += marginDip;
        headerTextRc.right -= marginDip;

        const D2D1_RECT_F modeButtonRc = RectFFromPixels(_modeButtonRect, dpi);
        const float radius             = 2.0f;

        float modeAlpha = 0.16f;
        if (_modeButtonPressed)
        {
            modeAlpha = 0.30f;
        }
        else if (_modeButtonHot)
        {
            modeAlpha = 0.22f;
        }

        brush->SetColor(ColorFFromColorRef(accent, modeAlpha));
        target->FillRoundedRectangle(D2D1::RoundedRect(modeButtonRc, radius, radius), brush);

        brush->SetColor(ColorFFromColorRef(accent, 0.85f));
        target->DrawRoundedRectangle(D2D1::RoundedRect(modeButtonRc, radius, radius), brush, 1.0f);

        brush->SetColor(ColorFFromColorRef(fg));
        target->DrawTextW(modeText.c_str(),
                          static_cast<UINT32>(modeText.size()),
                          _modeButtonFormat ? _modeButtonFormat.get() : _headerFormatRight.get(),
                          modeButtonRc,
                          brush,
                          D2D1_DRAW_TEXT_OPTIONS_CLIP);

        if (ActiveFileComboEntryCount() <= 1u)
        {
            D2D1_RECT_F fileRc = headerTextRc;
            fileRc.right       = std::max(fileRc.left, modeButtonRc.left - marginDip);
            target->DrawTextW(titleText.c_str(), static_cast<UINT32>(titleText.size()), _headerFormat.get(), fileRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }

        D2D1_RECT_F statusTextRc = statusRc;
        statusTextRc.left += marginDip;
        statusTextRc.right -= marginDip;
        target->DrawTextW(statusText.c_str(), static_cast<UINT32>(statusText.size()), _statusFormat.get(), statusTextRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);

        if (! _isLoading && _fileReader && _fileSize == 0 && ! _currentPath.empty())
        {
            const std::wstring emptyText = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_EMPTY_WATERMARK);
            if (! emptyText.empty())
            {
                const D2D1_RECT_F contentRc = RectFFromPixels(_contentRect, dpi);
                const float centerX         = (contentRc.left + contentRc.right) / 2.0f;
                const float centerY         = (contentRc.top + contentRc.bottom) / 2.0f;
                const float alpha           = _hasTheme && _theme.darkMode ? 0.28f : 0.20f;
                brush->SetColor(ColorFFromColorRef(fg, alpha));

                auto restoreTransform          = wil::scope_exit([&] { target->SetTransform(D2D1::Matrix3x2F::Identity()); });
                const D2D1_MATRIX_3X2_F rotate = D2D1::Matrix3x2F::Rotation(kWatermarkAngleDegrees, D2D1::Point2F(centerX, centerY));
                target->SetTransform(rotate * D2D1::Matrix3x2F::Identity());
                target->DrawTextW(
                    emptyText.c_str(), static_cast<UINT32>(emptyText.size()), _watermarkFormat.get(), contentRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }
        }
    };

    const auto drawWithTarget = [&](ID2D1RenderTarget* target, ID2D1SolidColorBrush* brush, bool sharedTarget) noexcept -> bool
    {
        if (! target || ! brush)
        {
            return false;
        }

        HRESULT hr = S_OK;
        {
            target->BeginDraw();
            auto endDraw = wil::scope_exit([&]
            {
                hr = target->EndDraw();
                if (hr == D2DERR_RECREATE_TARGET && sharedTarget)
                {
                    DiscardDirect2D();
                }
            });
            drawChrome(target, brush);
        }

        if (hr == D2DERR_RECREATE_TARGET)
        {
            return false;
        }

        return SUCCEEDED(hr);
    };

    static_cast<void>(EnsureDirect2D(_hWnd.get()));
    const bool hasDirectWriteChromeResources =
        _d2dFactory && _headerFormat && _headerFormatRight && _statusFormat && _watermarkFormat && (_modeButtonFormat || _headerFormatRight);
    if (hasDirectWriteChromeResources)
    {
        if (_d2dTarget && _d2dBrush && drawWithTarget(_d2dTarget.get(), _d2dBrush.get(), true))
        {
            return;
        }

        RECT client{};
        GetClientRect(_hWnd.get(), &client);

        wil::com_ptr<ID2D1DCRenderTarget> dcTarget;
        const D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        if (SUCCEEDED(_d2dFactory->CreateDCRenderTarget(&props, dcTarget.put())) && dcTarget)
        {
            dcTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
            if (SUCCEEDED(dcTarget->BindDC(hdc.get(), &client)))
            {
                wil::com_ptr<ID2D1SolidColorBrush> dcBrush;
                if (SUCCEEDED(dcTarget->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::Black), dcBrush.put())) && dcBrush &&
                    drawWithTarget(dcTarget.get(), dcBrush.get(), false))
                {
                    return;
                }
            }
        }
    }

    FillRect(hdc.get(), &ps.rcPaint, _backgroundBrush.get());

    FillRect(hdc.get(), &_headerRect, _headerBrush.get());
    FillRect(hdc.get(), &_statusRect, _statusBrush.get());

    const int lineThickness = std::max(1, MulDiv(2, dpiInt, USER_DEFAULT_SCREEN_DPI));

    RECT line{};
    line.left   = 0;
    line.right  = ps.rcPaint.right;
    line.top    = _headerRect.bottom - lineThickness;
    line.bottom = _headerRect.bottom;

    wil::unique_hbrush accentBrush(CreateSolidBrush(accent));
    FillRect(hdc.get(), &line, accentBrush.get());
}

void ViewerText::OnCommand(HWND hwnd, UINT commandId, UINT notifyCode, HWND control) noexcept
{
    if (! hwnd)
    {
        return;
    }

    static_cast<void>(notifyCode);
    static_cast<void>(control);

    if (IsEncodingMenuSelectionValid(commandId))
    {
        SetDisplayEncodingMenuSelection(hwnd, commandId, true);
        return;
    }

    if (IsSaveEncodingMenuSelectionValid(commandId))
    {
        SetSaveEncodingMenuSelection(hwnd, commandId);
        return;
    }

    switch (commandId)
    {
        case IDM_VIEWER_FILE_OPEN: CommandOpen(hwnd); break;
        case IDM_VIEWER_FILE_SAVE_AS: CommandSaveAs(hwnd); break;
        case IDM_VIEWER_FILE_REFRESH: CommandRefresh(hwnd); break;
        case IDM_VIEWER_FILE_EXIT: CommandExit(hwnd); break;

        case IDM_VIEWER_OTHER_NEXT: CommandOtherNext(hwnd); break;
        case IDM_VIEWER_OTHER_PREVIOUS: CommandOtherPrevious(hwnd); break;
        case IDM_VIEWER_OTHER_FIRST: CommandOtherFirst(hwnd); break;
        case IDM_VIEWER_OTHER_LAST: CommandOtherLast(hwnd); break;

        case IDM_VIEWER_SEARCH_FIND: CommandFind(hwnd); break;
        case IDM_VIEWER_SEARCH_FIND_NEXT: CommandFindNext(hwnd, false); break;
        case IDM_VIEWER_SEARCH_FIND_PREVIOUS: CommandFindNext(hwnd, true); break;

        case IDM_VIEWER_VIEW_TEXT:
            if (_documentKind == DocumentKind::Diff && _diffParsedAvailable)
            {
                SetDiffPresentation(hwnd, DiffPresentationMode::RawText);
            }
            else
            {
                SetViewMode(hwnd, ViewMode::Text);
            }
            break;
        case IDM_VIEWER_VIEW_DIFF_SIDE_BY_SIDE: SetDiffPresentation(hwnd, DiffPresentationMode::SideBySide); break;
        case IDM_VIEWER_VIEW_DIFF_INLINE: SetDiffPresentation(hwnd, DiffPresentationMode::Inline); break;
        case IDM_VIEWER_VIEW_DIFF_RAW_TEXT: SetDiffPresentation(hwnd, DiffPresentationMode::RawText); break;
        case IDM_VIEWER_VIEW_DIFF_SHOW_UNCHANGED:
            SetDiffContextMode(
                hwnd, _config.diffContextMode == DiffContextMode::FullFileWhenAvailable ? DiffContextMode::HunksOnly : DiffContextMode::FullFileWhenAvailable);
            break;
        case IDM_VIEWER_VIEW_DIFF_NEXT_HUNK: static_cast<void>(NavigateDiffHunk(hwnd, false)); break;
        case IDM_VIEWER_VIEW_DIFF_PREVIOUS_HUNK: static_cast<void>(NavigateDiffHunk(hwnd, true)); break;
        case IDM_VIEWER_VIEW_HEX: SetViewMode(hwnd, ViewMode::Hex); break;
        case IDM_VIEWER_VIEW_HEX_BYTE_COLORS_LEADING_NIBBLE: SetHexByteColorMode(hwnd, HexByteColorMode::LeadingNibble); break;
        case IDM_VIEWER_VIEW_HEX_BYTE_COLORS_OFF: SetHexByteColorMode(hwnd, HexByteColorMode::Off); break;
        case IDM_VIEWER_VIEW_GOTO_TOP: CommandGoToTop(hwnd, false); break;
        case IDM_VIEWER_VIEW_GOTO_BOTTOM: CommandGoToBottom(hwnd, false); break;
        case IDM_VIEWER_VIEW_GOTO_OFFSET: CommandGoToOffset(hwnd); break;
        case IDM_VIEWER_VIEW_LINE_NUMBERS: SetShowLineNumbers(hwnd, ! _config.showLineNumbers); break;
        case IDM_VIEWER_VIEW_WRAP: SetWrap(hwnd, ! _wrap); break;

        case IDM_VIEWER_ENCODING_NEXT: CommandCycleDisplayEncoding(hwnd, false); break;
        case IDM_VIEWER_ENCODING_PREVIOUS: CommandCycleDisplayEncoding(hwnd, true); break;
    }
}

void ViewerText::OnContextMenu(HWND hwnd, POINT screenPt) noexcept
{
    if (! _menuHandle)
    {
        _menuHandle.reset(Localization::LoadMenuResource(g_hInstance, IDR_VIEWERTEXT_MENU));
    }

    HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    UpdateMenuChecks(hwnd, false);
    static constexpr std::array<int, 8> kPreviewContextMenuExcludedCommandIds{{
        IDM_VIEWER_FILE_OPEN,
        IDM_VIEWER_FILE_EXIT,
        IDM_VIEWER_OTHER_NEXT,
        IDM_VIEWER_OTHER_PREVIOUS,
        IDM_VIEWER_OTHER_FIRST,
        IDM_VIEWER_OTHER_LAST,
        IDM_VIEWER_ENCODING_NEXT,
        IDM_VIEWER_ENCODING_PREVIOUS,
    }};
    RedSalamander::DxUi::NativeMenuFlyoutOptions previewMenuOptions{};
    previewMenuOptions.includeAcceleratorText = false;
    previewMenuOptions.omitEmptySubmenus      = true;
    previewMenuOptions.trimSeparators         = true;
    previewMenuOptions.excludedCommandIds     = kPreviewContextMenuExcludedCommandIds;

    const auto result = RedSalamander::DxUi::ShowNativeHMenuContextMenu(
        hwnd, screenPt, menu, _hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false), previewMenuOptions);
    if (result.has_value() && result.value() > 0)
    {
        OnCommand(hwnd, static_cast<UINT>(result.value()), 0, nullptr);
    }
}

LRESULT ViewerText::OnNotify(const NMHDR* header)
{
    static_cast<void>(header);
    return 0;
}

void ViewerText::UpdateMenuChecks(HWND hwnd, bool syncDxMenuBar) noexcept
{
    HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    const auto findSubmenuContainingCommand = [&](auto&& self, HMENU currentMenu, UINT commandId) noexcept -> HMENU
    {
        if (! currentMenu)
        {
            return nullptr;
        }

        const int count = GetMenuItemCount(currentMenu);
        if (count <= 0)
        {
            return nullptr;
        }

        for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
        {
            MENUITEMINFOW info{};
            info.cbSize = sizeof(info);
            info.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_SUBMENU;
            if (GetMenuItemInfoW(currentMenu, pos, TRUE, &info) == 0)
            {
                continue;
            }

            if ((info.fType & MFT_SEPARATOR) == 0 && ! info.hSubMenu && info.wID == commandId)
            {
                return currentMenu;
            }

            if (info.hSubMenu)
            {
                if (HMENU foundMenu = self(self, info.hSubMenu, commandId))
                {
                    return foundMenu;
                }
            }
        }

        return nullptr;
    };

    const UINT selectedDisplay = EffectiveDisplayEncodingMenuSelection();
    HMENU encodingMenu         = findSubmenuContainingCommand(findSubmenuContainingCommand, menu, IDM_VIEWER_ENCODING_DISPLAY_ANSI);

    if (encodingMenu)
    {
        auto updateEncodingChecks = [&](auto&& self, HMENU currentMenu) noexcept -> void
        {
            if (! currentMenu)
            {
                return;
            }

            const int count = GetMenuItemCount(currentMenu);
            if (count <= 0)
            {
                Debug::Error(L"Encoding menu has no items");
                return;
            }

            for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
            {
                MENUITEMINFOW info{};
                info.cbSize = sizeof(info);
                info.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_ID | MIIM_SUBMENU;
                if (GetMenuItemInfoW(currentMenu, pos, TRUE, &info) == 0)
                {
                    continue;
                }

                if (info.hSubMenu)
                {
                    self(self, info.hSubMenu);
                    continue;
                }

                if ((info.fType & MFT_SEPARATOR) != 0)
                {
                    continue;
                }

                if (! IsEncodingMenuSelectionValid(info.wID))
                {
                    continue;
                }

                info.fType |= MFT_RADIOCHECK;

                info.fState &= ~MFS_CHECKED;
                if (info.wID == selectedDisplay)
                {
                    info.fState |= MFS_CHECKED;
                }

                static_cast<void>(SetMenuItemInfoW(currentMenu, pos, TRUE, &info));
            }
        };

        updateEncodingChecks(updateEncodingChecks, encodingMenu);
    }

    const UINT selectedSave = EffectiveSaveEncodingMenuSelection();
    CheckMenuRadioItem(menu, IDM_VIEWER_ENCODING_SAVE_FIRST, IDM_VIEWER_ENCODING_SAVE_LAST, selectedSave, MF_BYCOMMAND);

    if (HMENU hexByteColorMenu = findSubmenuContainingCommand(findSubmenuContainingCommand, menu, IDM_VIEWER_VIEW_HEX_BYTE_COLORS_LEADING_NIBBLE))
    {
        const UINT selectedHexByteColorMode = (_config.hexByteColorMode == HexByteColorMode::LeadingNibble) ? IDM_VIEWER_VIEW_HEX_BYTE_COLORS_LEADING_NIBBLE
                                                                                                            : IDM_VIEWER_VIEW_HEX_BYTE_COLORS_OFF;
        CheckMenuRadioItem(
            hexByteColorMenu, IDM_VIEWER_VIEW_HEX_BYTE_COLORS_FIRST, IDM_VIEWER_VIEW_HEX_BYTE_COLORS_LAST, selectedHexByteColorMode, MF_BYCOMMAND);
    }

    CheckMenuItem(menu,
                  IDM_VIEWER_VIEW_TEXT,
                  static_cast<UINT>(MF_BYCOMMAND |
                                    (_viewMode == ViewMode::Text && (_documentKind != DocumentKind::Diff || _diffPresentation == DiffPresentationMode::RawText)
                                         ? MF_CHECKED
                                         : MF_UNCHECKED)));
    CheckMenuItem(menu, IDM_VIEWER_VIEW_HEX, static_cast<UINT>(MF_BYCOMMAND | (_viewMode == ViewMode::Hex ? MF_CHECKED : MF_UNCHECKED)));
    const bool parsedDiffAvailable = _documentKind == DocumentKind::Diff && _diffParsedAvailable;
    CheckMenuItem(
        menu,
        IDM_VIEWER_VIEW_DIFF_SIDE_BY_SIDE,
        static_cast<UINT>(MF_BYCOMMAND | (_viewMode == ViewMode::Text && _diffPresentation == DiffPresentationMode::SideBySide ? MF_CHECKED : MF_UNCHECKED)));
    CheckMenuItem(
        menu,
        IDM_VIEWER_VIEW_DIFF_INLINE,
        static_cast<UINT>(MF_BYCOMMAND | (_viewMode == ViewMode::Text && _diffPresentation == DiffPresentationMode::Inline ? MF_CHECKED : MF_UNCHECKED)));
    CheckMenuItem(menu,
                  IDM_VIEWER_VIEW_DIFF_SHOW_UNCHANGED,
                  static_cast<UINT>(MF_BYCOMMAND | (_config.diffContextMode == DiffContextMode::FullFileWhenAvailable ? MF_CHECKED : MF_UNCHECKED)));
    CheckMenuItem(
        menu, IDM_VIEWER_VIEW_LINE_NUMBERS, static_cast<UINT>(MF_BYCOMMAND | (ShowTextLineNumbersInCurrentPresentation() ? MF_CHECKED : MF_UNCHECKED)));
    CheckMenuItem(menu, IDM_VIEWER_VIEW_WRAP, static_cast<UINT>(MF_BYCOMMAND | (_wrap ? MF_CHECKED : MF_UNCHECKED)));

    EnableMenuItem(menu, IDM_VIEWER_VIEW_DIFF_SIDE_BY_SIDE, static_cast<UINT>(MF_BYCOMMAND | (parsedDiffAvailable ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWER_VIEW_DIFF_INLINE, static_cast<UINT>(MF_BYCOMMAND | (parsedDiffAvailable ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWER_VIEW_DIFF_SHOW_UNCHANGED, static_cast<UINT>(MF_BYCOMMAND | (parsedDiffAvailable ? MF_ENABLED : MF_GRAYED)));
    const auto* currentVariant         = CurrentDiffVariant();
    const bool hunkNavigationAvailable = _viewMode == ViewMode::Text && _documentKind == DocumentKind::Diff &&
                                         _diffPresentation != DiffPresentationMode::RawText && currentVariant && ! currentVariant->hunkNavigation.empty();
    EnableMenuItem(menu, IDM_VIEWER_VIEW_DIFF_NEXT_HUNK, static_cast<UINT>(MF_BYCOMMAND | (hunkNavigationAvailable ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWER_VIEW_DIFF_PREVIOUS_HUNK, static_cast<UINT>(MF_BYCOMMAND | (hunkNavigationAvailable ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu,
                   IDM_VIEWER_VIEW_LINE_NUMBERS,
                   static_cast<UINT>(MF_BYCOMMAND | (_viewMode == ViewMode::Text && ! HasParsedDiffPresentation() ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWER_VIEW_WRAP, static_cast<UINT>(MF_BYCOMMAND | (_viewMode == ViewMode::Text ? MF_ENABLED : MF_GRAYED)));

    if (syncDxMenuBar && _menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
}

LRESULT ViewerText::OnCtlColor([[maybe_unused]] UINT msg, HDC hdc, HWND control) noexcept
{
    if (! hdc || ! control || ! _hasTheme)
    {
        return 0;
    }

    if (_theme.highContrast)
    {
        return 0;
    }

    if (_hFileComboHost && control == _hFileComboHost.get())
    {
        SetBkMode(hdc, OPAQUE);
        SetTextColor(hdc, _uiText);
        SetBkColor(hdc, _uiHeaderBg);
        return reinterpret_cast<LRESULT>(_headerBrush.get());
    }

    return 0;
}

void ViewerText::OnMouseMove(int x, int y) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const POINT pt{.x = x, .y = y};
    const bool hot = PtInRect(&_modeButtonRect, pt) != 0;

    if (hot != _modeButtonHot)
    {
        _modeButtonHot = hot;
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
    }

    if (! _trackingMouseLeave)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = _hWnd.get();
        if (TrackMouseEvent(&tme) != 0)
        {
            _trackingMouseLeave = true;
        }
    }
}

void ViewerText::OnMouseLeave() noexcept
{
    _trackingMouseLeave = false;
    if (_modeButtonHot)
    {
        _modeButtonHot = false;
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        }
    }
}

void ViewerText::OnLButtonDown(int x, int y) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const POINT pt{.x = x, .y = y};
    if (PtInRect(&_modeButtonRect, pt) == 0)
    {
        return;
    }

    _modeButtonPressed = true;
    SetCapture(_hWnd.get());
    InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
}

void ViewerText::OnLButtonUp(int x, int y) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const HWND captured = GetCapture();
    if (captured == _hWnd.get())
    {
        ReleaseCapture();
    }

    const bool wasPressed = _modeButtonPressed;
    _modeButtonPressed    = false;

    if (wasPressed)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);

        const POINT pt{.x = x, .y = y};
        if (PtInRect(&_modeButtonRect, pt) != 0)
        {
            if (_documentKind == DocumentKind::Diff && _diffParsedAvailable)
            {
                if (_viewMode == ViewMode::Hex)
                {
                    SetDiffPresentation(_hWnd.get(), PreferredParsedDiffPresentation());
                }
                else if (_diffPresentation == DiffPresentationMode::RawText)
                {
                    SetViewMode(_hWnd.get(), ViewMode::Hex);
                }
                else
                {
                    SetDiffPresentation(_hWnd.get(), DiffPresentationMode::RawText);
                }
            }
            else
            {
                SetViewMode(_hWnd.get(), _viewMode == ViewMode::Hex ? ViewMode::Text : ViewMode::Hex);
            }
        }
    }
}

bool ViewerText::OnSetCursor(HWND hwnd, LPARAM lParam) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    if (LOWORD(lParam) != HTCLIENT)
    {
        return false;
    }

    POINT pt{};
    if (GetCursorPos(&pt) == 0)
    {
        return false;
    }

    if (ScreenToClient(hwnd, &pt) == 0)
    {
        return false;
    }

    if (PtInRect(&_modeButtonRect, pt) == 0)
    {
        return false;
    }

    SetCursor(LoadCursorW(nullptr, IDC_HAND));
    return true;
}

void ViewerText::ApplyTheme(HWND hwnd) noexcept
{
    const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

    _backgroundBrush.reset(CreateSolidBrush(bg));

    COLORREF headerBg = bg;
    COLORREF statusBg = bg;
    if (_hasTheme && _theme.darkMode)
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 10), std::max(0, GetGValue(bg) - 10), std::max(0, GetBValue(bg) - 10));
        statusBg = RGB(std::min(255, GetRValue(bg) + 5), std::min(255, GetGValue(bg) + 5), std::min(255, GetBValue(bg) + 5));
    }
    else
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 5), std::max(0, GetGValue(bg) - 5), std::max(0, GetBValue(bg) - 5));
        statusBg = RGB(std::min(255, GetRValue(bg) + 5), std::min(255, GetGValue(bg) + 5), std::min(255, GetBValue(bg) + 5));
    }

    _headerBrush.reset(CreateSolidBrush(headerBg));
    _statusBrush.reset(CreateSolidBrush(statusBg));

    _uiBackground = bg;
    _uiText       = fg;
    _uiHeaderBg   = headerBg;
    _uiStatusBg   = statusBg;

    if (_hasTheme && _hWnd)
    {
        const bool windowActive = GetActiveWindow() == _hWnd.get();
        ApplyTitleBarTheme(windowActive);
    }

    const wchar_t* winTheme = L"Explorer";
    if (_hasTheme && _theme.highContrast)
    {
        winTheme = L"";
    }
    else if (_hasTheme && _theme.darkMode)
    {
        winTheme = L"DarkMode_Explorer";
    }

    if (_hEdit)
    {
        SetWindowTheme(_hEdit.get(), winTheme, nullptr);
        SendMessageW(_hEdit.get(), WM_THEMECHANGED, 0, 0);
    }
    if (_hHex)
    {
        SetWindowTheme(_hHex.get(), winTheme, nullptr);
        SendMessageW(_hHex.get(), WM_THEMECHANGED, 0, 0);
    }

    if (_hFileComboHost)
    {
        _fileComboHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
    }
    _menuBarHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));

    ApplyMenuTheme(hwnd);
    UpdateMenuChecks(hwnd);

    if (_hEdit)
    {
        InvalidateRect(_hEdit.get(), nullptr, TRUE);
    }
    if (_hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }
}

void ViewerText::ApplyTitleBarTheme(bool windowActive) noexcept
{
    if (! _hasTheme || ! _hWnd)
    {
        return;
    }

    static constexpr DWORD kDwmwaUseImmersiveDarkMode19 = 19u;
    static constexpr DWORD kDwmwaUseImmersiveDarkMode20 = 20u;
    static constexpr DWORD kDwmwaBorderColor            = 34u;
    static constexpr DWORD kDwmwaCaptionColor           = 35u;
    static constexpr DWORD kDwmwaTextColor              = 36u;
    static constexpr DWORD kDwmColorDefault             = 0xFFFFFFFFu;

    const BOOL darkMode = (_theme.darkMode && ! _theme.highContrast) ? TRUE : FALSE;
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaUseImmersiveDarkMode20, &darkMode, sizeof(darkMode));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaUseImmersiveDarkMode19, &darkMode, sizeof(darkMode));

    DWORD borderValue  = kDwmColorDefault;
    DWORD captionValue = kDwmColorDefault;
    DWORD textValue    = kDwmColorDefault;
    if (! _theme.highContrast && _theme.rainbowMode)
    {
        COLORREF accent = ResolveAccentColor(_theme, L"title");
        if (! windowActive)
        {
            static constexpr uint8_t kInactiveTitleBlendAlpha = 223u; // ~7/8 toward background
            const COLORREF bg                                 = ColorRefFromArgb(_theme.backgroundArgb);
            accent                                            = BlendColor(accent, bg, kInactiveTitleBlendAlpha);
        }

        const COLORREF text = ContrastingTextColor(accent);
        borderValue         = static_cast<DWORD>(accent);
        captionValue        = static_cast<DWORD>(accent);
        textValue           = static_cast<DWORD>(text);
    }

    DwmSetWindowAttribute(_hWnd.get(), kDwmwaBorderColor, &borderValue, sizeof(borderValue));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaCaptionColor, &captionValue, sizeof(captionValue));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaTextColor, &textValue, sizeof(textValue));
}

void ViewerText::ResetDiffState() noexcept
{
    _documentKind               = DocumentKind::PlainText;
    _diffPresentation           = DiffPresentationMode::RawText;
    _lastParsedDiffPresentation = _config.diffDefaultLayout == DiffDefaultLayout::Inline ? DiffPresentationMode::Inline : DiffPresentationMode::SideBySide;
    _diffParsedAvailable        = false;
    _diffRawTextBuffer.clear();
    _parsedDiffDocument.reset();
    _diffReferenceCache.reset();
    _diffStreamSections.clear();
    _diffExpandedSectionIndex.reset();
    _diffInlineHunksOnly         = {};
    _diffInlineExpanded          = {};
    _diffSideBySideHunksOnly     = {};
    _diffSideBySideExpanded      = {};
    _diffInlineExpandedBuilt     = false;
    _diffSideBySideExpandedBuilt = false;
#ifdef _DEBUG
    _debugDiffParseCount = 0u;
#endif
}

bool ViewerText::EnsureDiffVariantBuilt(DiffPresentationMode presentation, DiffContextMode contextMode) noexcept
{
    if (_documentKind != DocumentKind::Diff || ! _diffParsedAvailable || presentation == DiffPresentationMode::RawText ||
        contextMode != DiffContextMode::FullFileWhenAvailable)
    {
        return true;
    }

    DiffTextVariant* variant = nullptr;
    bool* variantBuilt       = nullptr;
    if (presentation == DiffPresentationMode::Inline)
    {
        variant      = &_diffInlineExpanded;
        variantBuilt = &_diffInlineExpandedBuilt;
    }
    else
    {
        variant      = &_diffSideBySideExpanded;
        variantBuilt = &_diffSideBySideExpandedBuilt;
    }

    if (! variant || ! variantBuilt || *variantBuilt)
    {
        return true;
    }

    if (! _parsedDiffDocument)
    {
        auto parsedDocument = std::make_shared<ParsedDiffDocument>();
        if (! ParseUnifiedDiffDocument(_diffRawTextBuffer, *parsedDocument))
        {
            return false;
        }
        _parsedDiffDocument = std::move(parsedDocument);
#ifdef _DEBUG
        _debugDiffParseCount += 1u;
#endif
    }

    if (! _diffReferenceCache)
    {
        _diffReferenceCache = std::make_shared<DiffReferenceCache>();
    }

    std::optional<size_t> hydratedSectionIndex;
    if (_parsedDiffDocument && ! _parsedDiffDocument->files.empty())
    {
        size_t targetSection      = _diffExpandedSectionIndex.value_or(CurrentDiffSectionIndex());
        targetSection             = std::min(targetSection, _parsedDiffDocument->files.size() - 1u);
        hydratedSectionIndex      = targetSection;
        _diffExpandedSectionIndex = targetSection;
    }
    else
    {
        _diffExpandedSectionIndex.reset();
    }

    const auto hydratedLogicalRange = ComputeVisibleDiffHydrationLogicalRange(_hEdit.get());

    wil::com_ptr<IFileSystemIO> fileIo;
    IFileSystemIO* fileIoRaw = nullptr;
    if (_fileSystem && SUCCEEDED(_fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void())) && fileIo)
    {
        fileIoRaw = fileIo.get();
    }

    if (presentation == DiffPresentationMode::Inline)
    {
        *variant = BuildInlineDiffText(
                       *_parsedDiffDocument, contextMode, _currentPath, fileIoRaw, _diffReferenceCache.get(), hydratedSectionIndex, hydratedLogicalRange)
                       .variant;
    }
    else
    {
        *variant = BuildSideBySideDiffText(
                       *_parsedDiffDocument, contextMode, _currentPath, fileIoRaw, _diffReferenceCache.get(), hydratedSectionIndex, hydratedLogicalRange)
                       .variant;
    }

    *variantBuilt = true;
    return true;
}

namespace
{
[[nodiscard]] bool ShouldHydrateExpandedDiffSection(ViewerText::DiffContextMode contextMode,
                                                    std::optional<size_t> hydratedSectionIndex,
                                                    size_t fileIndex) noexcept
{
    return contextMode == ViewerText::DiffContextMode::FullFileWhenAvailable &&
           (! hydratedSectionIndex.has_value() || hydratedSectionIndex.value() == fileIndex);
}

[[nodiscard]] bool ShouldHydrateExpandedDiffLogicalLine(std::optional<std::pair<uint32_t, uint32_t>> hydratedLogicalRange, uint32_t logicalLine) noexcept
{
    return ! hydratedLogicalRange.has_value() || (logicalLine >= hydratedLogicalRange->first && logicalLine < hydratedLogicalRange->second);
}

[[nodiscard]] std::wstring MaskDeferredDiffText(std::wstring_view text)
{
    return std::wstring(text.size(), L' ');
}
} // namespace

bool ViewerText::EnsureCurrentDiffVariantBuilt() noexcept
{
    return EnsureDiffVariantBuilt(_diffPresentation, _config.diffContextMode);
}

const ViewerText::DiffTextVariant* ViewerText::CurrentDiffVariant() const noexcept
{
    if (_documentKind != DocumentKind::Diff || ! _diffParsedAvailable || _diffPresentation == DiffPresentationMode::RawText)
    {
        return nullptr;
    }

    const bool expanded = _config.diffContextMode == DiffContextMode::FullFileWhenAvailable;
    if (_diffPresentation == DiffPresentationMode::Inline)
    {
        if (expanded && ! _diffInlineExpandedBuilt)
        {
            return nullptr;
        }
        return expanded ? &_diffInlineExpanded : &_diffInlineHunksOnly;
    }

    if (expanded && ! _diffSideBySideExpandedBuilt)
    {
        return nullptr;
    }
    return expanded ? &_diffSideBySideExpanded : &_diffSideBySideHunksOnly;
}

bool ViewerText::UseDiffSectionFileCombo() const noexcept
{
    const auto* variant = CurrentDiffVariant();
    if (variant && variant->sectionNavigation.size() > 1u)
    {
        return true;
    }

    return _documentKind == DocumentKind::Diff && ! _diffParsedAvailable && _textStreamActive && _diffStreamSections.size() > 1u;
}

size_t ViewerText::ActiveFileComboEntryCount() const noexcept
{
    if (UseDiffSectionFileCombo())
    {
        if (const auto* variant = CurrentDiffVariant())
        {
            return variant->sectionNavigation.size();
        }

        return _diffStreamSections.size();
    }

    return _otherFiles.size();
}

size_t ViewerText::CurrentDiffSectionIndex() const noexcept
{
    const auto* variant = CurrentDiffVariant();
    if (variant && ! variant->sectionNavigation.empty())
    {
        uint32_t topLogicalLine = 0u;
        if (! _textVisualLineLogical.empty())
        {
            const size_t topVisual = std::min<size_t>(_textTopVisualLine, _textVisualLineLogical.size() - 1u);
            topLogicalLine         = _textVisualLineLogical[topVisual];
        }

        const auto it =
            std::upper_bound(variant->sectionNavigation.begin(),
                             variant->sectionNavigation.end(),
                             topLogicalLine,
                             [](uint32_t line, const DiffTextVariant::SectionNavigationEntry& entry) noexcept { return line < entry.startLogicalLine; });
        if (it == variant->sectionNavigation.begin())
        {
            return 0u;
        }

        return static_cast<size_t>(std::distance(variant->sectionNavigation.begin(), it) - 1);
    }

    if (_documentKind == DocumentKind::Diff && ! _diffParsedAvailable && _textStreamActive && ! _diffStreamSections.empty())
    {
        const auto it = std::upper_bound(_diffStreamSections.begin(),
                                         _diffStreamSections.end(),
                                         _textStreamStartOffset,
                                         [](uint64_t offset, const StreamedDiffSectionEntry& entry) noexcept { return offset < entry.startOffset; });
        if (it == _diffStreamSections.begin())
        {
            return 0u;
        }

        return static_cast<size_t>(std::distance(_diffStreamSections.begin(), it) - 1);
    }

    return 0u;
}

size_t ViewerText::CurrentDiffHunkIndex() const noexcept
{
    const auto* variant = CurrentDiffVariant();
    if (variant && ! variant->hunkNavigation.empty())
    {
        uint32_t topLogicalLine = 0u;
        if (! _textVisualLineLogical.empty())
        {
            const size_t topVisual = std::min<size_t>(_textTopVisualLine, _textVisualLineLogical.size() - 1u);
            topLogicalLine         = _textVisualLineLogical[topVisual];
        }

        const auto it =
            std::upper_bound(variant->hunkNavigation.begin(),
                             variant->hunkNavigation.end(),
                             topLogicalLine,
                             [](uint32_t line, const DiffTextVariant::HunkNavigationEntry& entry) noexcept { return line < entry.startLogicalLine; });
        if (it == variant->hunkNavigation.begin())
        {
            return 0u;
        }

        return static_cast<size_t>(std::distance(variant->hunkNavigation.begin(), it) - 1);
    }

    return 0u;
}

bool ViewerText::ShowTextLineNumbersInCurrentPresentation() const noexcept
{
    return _config.showLineNumbers && ! HasParsedDiffPresentation();
}

bool ViewerText::HasParsedDiffPresentation() const noexcept
{
    return _documentKind == DocumentKind::Diff && _diffParsedAvailable && _diffPresentation != DiffPresentationMode::RawText;
}

ViewerText::DiffPresentationMode ViewerText::PreferredParsedDiffPresentation() const noexcept
{
    if (_lastParsedDiffPresentation == DiffPresentationMode::Inline || _lastParsedDiffPresentation == DiffPresentationMode::SideBySide)
    {
        return _lastParsedDiffPresentation;
    }

    return _config.diffDefaultLayout == DiffDefaultLayout::Inline ? DiffPresentationMode::Inline : DiffPresentationMode::SideBySide;
}

void ViewerText::ApplyCurrentTextPresentation(HWND hwnd, bool preserveViewport) noexcept
{
    const bool preserveParsedDiffViewport = preserveViewport && _documentKind == DocumentKind::Diff && _diffParsedAvailable;

    std::wstring nextText;
    if (_documentKind == DocumentKind::Diff && _diffParsedAvailable)
    {
        static_cast<void>(EnsureCurrentDiffVariantBuilt());
        if (const auto* variant = CurrentDiffVariant())
        {
            nextText = variant->text;
        }
        else
        {
            nextText = _diffRawTextBuffer;
        }
    }
    else
    {
        nextText = _textBuffer;
    }

    const uint32_t savedTopVisualLine = _textTopVisualLine;
    const uint32_t savedLeftColumn    = _textLeftColumn;
    const size_t savedCaretIndex      = _textCaretIndex;
    const size_t savedSelAnchor       = _textSelAnchor;
    const size_t savedSelActive       = _textSelActive;
    const size_t savedPreferredColumn = _textPreferredColumn;
    const bool savedSelecting         = _textSelecting;

    _textBuffer = std::move(nextText);
    BuildTextLineIndex(_textBuffer, _textLineStarts, _textLineEnds, _textMaxLineLength);

    _textTotalLineCount.reset();
    _textStreamLineCountedEndOffset = _textStreamStartOffset;
    _textStreamLineCountedNewlines  = 0;
    _textStreamLineCountLastWasCR   = false;
    UpdateTextStreamTotalLineCountAfterLoad();

    _textVisualLineStarts.clear();
    _textVisualLineLogical.clear();
    _textTopVisualLine   = preserveViewport ? savedTopVisualLine : 0u;
    _textLeftColumn      = preserveViewport ? savedLeftColumn : 0u;
    _textCaretIndex      = preserveViewport ? std::min(savedCaretIndex, _textBuffer.size()) : 0u;
    _textSelAnchor       = preserveViewport ? std::min(savedSelAnchor, _textBuffer.size()) : 0u;
    _textSelActive       = preserveViewport ? std::min(savedSelActive, _textBuffer.size()) : 0u;
    _textPreferredColumn = preserveViewport ? savedPreferredColumn : 0u;
    _textSelecting       = preserveViewport ? savedSelecting : false;
    _searchMatchStarts.clear();

    const HWND textWindow = hwnd ? hwnd : _hEdit.get();
    if (textWindow)
    {
        RebuildTextVisualLines(textWindow);
        if (! _textVisualLineStarts.empty())
        {
            _textTopVisualLine = std::min<uint32_t>(_textTopVisualLine, static_cast<uint32_t>(_textVisualLineStarts.size() - 1u));
        }
        else
        {
            _textTopVisualLine = 0u;
        }
        UpdateTextViewScrollBars(textWindow);
        UpdateSearchHighlights();
        InvalidateRect(textWindow, nullptr, TRUE);
    }

    if (_hWnd)
    {
        if (preserveParsedDiffViewport)
        {
            SyncFileComboSelection();
        }
        else
        {
            RefreshFileCombo(_hWnd.get());
            if (textWindow)
            {
                RebuildTextVisualLines(textWindow);
                UpdateTextViewScrollBars(textWindow);
                InvalidateRect(textWindow, nullptr, TRUE);
            }
        }
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
}

void ViewerText::SetDiffPresentation(HWND hwnd, DiffPresentationMode mode) noexcept
{
    if (_documentKind != DocumentKind::Diff)
    {
        return;
    }

    if (mode != DiffPresentationMode::RawText && ! _diffParsedAvailable)
    {
        return;
    }

    const size_t currentSection = CurrentDiffSectionIndex();
    if (mode == DiffPresentationMode::Inline)
    {
        _config.diffDefaultLayout = DiffDefaultLayout::Inline;
        RefreshConfigurationJson();
    }
    else if (mode == DiffPresentationMode::SideBySide)
    {
        _config.diffDefaultLayout = DiffDefaultLayout::SideBySide;
        RefreshConfigurationJson();
    }

    if (mode != DiffPresentationMode::RawText && _config.diffContextMode == DiffContextMode::FullFileWhenAvailable && _parsedDiffDocument &&
        ! _parsedDiffDocument->files.empty())
    {
        const size_t targetSection = std::min(CurrentDiffSectionIndex(), _parsedDiffDocument->files.size() - 1u);
        if (! _diffExpandedSectionIndex.has_value() || _diffExpandedSectionIndex.value() != targetSection)
        {
            _diffExpandedSectionIndex = targetSection;
            _diffReferenceCache.reset();
            _diffInlineExpandedBuilt     = false;
            _diffSideBySideExpandedBuilt = false;
        }
    }

    _diffPresentation = mode;
    if (mode != DiffPresentationMode::RawText)
    {
        _lastParsedDiffPresentation = mode;
    }
    ApplyCurrentTextPresentation(_hEdit.get());
    if (mode != DiffPresentationMode::RawText && _diffParsedAvailable)
    {
        ScrollToDiffSection(_hEdit.get(), currentSection);
    }

    const HWND root = hwnd ? hwnd : _hWnd.get();
    if (root)
    {
        SetViewMode(root, ViewMode::Text);
    }
}

void ViewerText::SetDiffContextMode(HWND hwnd, DiffContextMode mode) noexcept
{
    const DiffContextMode previousMode = _config.diffContextMode;
    const size_t currentSection        = CurrentDiffSectionIndex();
    _config.diffContextMode            = mode;
    RefreshConfigurationJson();

    if (_documentKind == DocumentKind::Diff && _diffParsedAvailable)
    {
        if (mode == DiffContextMode::FullFileWhenAvailable && _parsedDiffDocument && ! _parsedDiffDocument->files.empty())
        {
            _diffExpandedSectionIndex = std::min(currentSection, _parsedDiffDocument->files.size() - 1u);
        }
        else
        {
            _diffExpandedSectionIndex.reset();
        }

        if (previousMode != mode)
        {
            _diffReferenceCache.reset();
            _diffInlineExpandedBuilt     = false;
            _diffSideBySideExpandedBuilt = false;
        }
    }

    if (HasParsedDiffPresentation())
    {
        ApplyCurrentTextPresentation(_hEdit.get());
        ScrollToDiffSection(_hEdit.get(), currentSection);
    }

    const HWND root = hwnd ? hwnd : _hWnd.get();
    if (root)
    {
        UpdateMenuChecks(root);
        InvalidateRect(root, &_statusRect, FALSE);
    }
}

void ViewerText::SetViewMode(HWND hwnd, ViewMode mode) noexcept
{
    const ViewMode previous = _viewMode;
    _viewMode               = mode;

    const bool showContent = _isLoading || ! (_fileSize == 0 && ! _currentPath.empty());

    if (_hEdit)
    {
        ShowWindow(_hEdit.get(), showContent && _viewMode == ViewMode::Text ? SW_SHOW : SW_HIDE);
    }
    if (_hHex)
    {
        ShowWindow(_hHex.get(), showContent && _viewMode == ViewMode::Hex ? SW_SHOW : SW_HIDE);
    }

    if (_embeddedMode)
    {
        if (_viewMode == ViewMode::Hex && previous != _viewMode && ! _isLoading && _hexBytes.empty() && _hexCacheValid == 0)
        {
            static_cast<void>(LoadHexData(hwnd));
        }
    }
    else if (! showContent)
    {
        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
    }
    else if (_viewMode == ViewMode::Hex)
    {
        if (previous != _viewMode && ! _isLoading && _hexBytes.empty() && _hexCacheValid == 0)
        {
            static_cast<void>(LoadHexData(hwnd));
        }
        if (_hHex)
        {
            SetFocus(_hHex.get());
        }
    }
    else
    {
        if (_hEdit)
        {
            SetFocus(_hEdit.get());
        }
    }

    UpdateMenuChecks(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
}

void ViewerText::SetHexByteColorMode(HWND hwnd, HexByteColorMode mode) noexcept
{
    _config.hexByteColorMode = mode;
    RefreshConfigurationJson();

    const HWND menuWindow = hwnd ? hwnd : _hWnd.get();
    if (menuWindow)
    {
        UpdateMenuChecks(menuWindow);
    }

    if (_hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, TRUE);
        if (_viewMode == ViewMode::Hex)
        {
            static_cast<void>(RedrawWindow(_hHex.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW));
        }
    }
}

void ViewerText::CommandExit(HWND /*hwnd*/) noexcept
{
    static_cast<void>(Close());
}

std::optional<std::filesystem::path> ViewerText::ShowOpenDialog(HWND hwnd) noexcept
{
    wil::com_ptr<IFileOpenDialog> dialog;
    const HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return std::nullopt;
    }

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DIALOG_OPEN_TITLE);
    if (! title.empty())
    {
        static_cast<void>(dialog->SetTitle(title.c_str()));
    }

    DWORD options = 0;
    static_cast<void>(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST | FOS_PATHMUSTEXIST;
    static_cast<void>(dialog->SetOptions(options));

    COMDLG_FILTERSPEC spec{};
    const std::wstring allFiles = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DIALOG_FILTER_ALL_FILES);
    spec.pszName                = allFiles.c_str();
    spec.pszSpec                = L"*.*";
    static_cast<void>(dialog->SetFileTypes(1, &spec));

    const HRESULT showHr = dialog->Show(hwnd);
    if (FAILED(showHr))
    {
        return std::nullopt;
    }

    wil::com_ptr<IShellItem> item;
    const HRESULT itemHr = dialog->GetResult(item.put());
    if (FAILED(itemHr) || ! item)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string path;
    const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
    if (FAILED(nameHr) || ! path)
    {
        return std::nullopt;
    }

    return std::filesystem::path(path.get());
}

std::optional<ViewerText::SaveAsResult> ViewerText::ShowSaveAsDialog(HWND hwnd) noexcept
{
    wil::com_ptr<IFileSaveDialog> dialog;
    const HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return std::nullopt;
    }

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DIALOG_SAVE_TITLE);
    if (! title.empty())
    {
        static_cast<void>(dialog->SetTitle(title.c_str()));
    }

    DWORD options = 0;
    static_cast<void>(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_OVERWRITEPROMPT;
    static_cast<void>(dialog->SetOptions(options));

    if (! _currentPath.empty())
    {
        const std::wstring fileName = _currentPath.filename().wstring();
        if (! fileName.empty())
        {
            static_cast<void>(dialog->SetFileName(fileName.c_str()));
        }
    }

    UINT initialEncodingSelection = IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL;
    switch (EffectiveSaveEncodingMenuSelection())
    {
        case IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL: initialEncodingSelection = IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL; break;
        case IDM_VIEWER_ENCODING_SAVE_ANSI: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_ANSI; break;
        case IDM_VIEWER_ENCODING_SAVE_UTF8: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_UTF8; break;
        case IDM_VIEWER_ENCODING_SAVE_UTF8_BOM: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM; break;
        case IDM_VIEWER_ENCODING_SAVE_UTF16BE_BOM: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM; break;
        case IDM_VIEWER_ENCODING_SAVE_UTF16LE_BOM: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM; break;
        default: initialEncodingSelection = IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL; break;
    }

    static constexpr DWORD kEncodingComboId = 6100u;

    wil::com_ptr<IFileDialogCustomize> customize;
    static_cast<void>(dialog->QueryInterface(IID_PPV_ARGS(customize.put())));
    if (customize)
    {
        const std::wstring encodingLabel = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_SAVEAS_ENCODING_LABEL);
        static_cast<void>(customize->AddComboBox(kEncodingComboId));
        if (! encodingLabel.empty())
        {
            static_cast<void>(customize->SetControlLabel(kEncodingComboId, encodingLabel.c_str()));
        }

        auto stripMenuText = [](std::wstring_view text) -> std::wstring
        {
            const size_t tabPos = text.find(L'\t');
            if (tabPos != std::wstring_view::npos)
            {
                text = text.substr(0, tabPos);
            }

            std::wstring result;
            result.reserve(text.size());

            for (size_t i = 0; i < text.size(); ++i)
            {
                const wchar_t ch = text[i];
                if (ch != L'&')
                {
                    result.push_back(ch);
                    continue;
                }

                if ((i + 1) < text.size() && text[i + 1] == L'&')
                {
                    result.push_back(L'&');
                    i += 1;
                }
            }

            while (! result.empty() && result.front() == L' ')
            {
                result.erase(result.begin());
            }
            while (! result.empty() && result.back() == L' ')
            {
                result.pop_back();
            }

            return result;
        };

        auto addMenuItemToCombo = [&](UINT commandId) noexcept
        {
            if (! hwnd)
            {
                return;
            }

            HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
            if (! menu)
            {
                return;
            }

            auto findMenuText = [&](auto&& self, HMENU currentMenu, UINT targetId) -> std::wstring
            {
                if (! currentMenu)
                {
                    return {};
                }

                const int count = GetMenuItemCount(currentMenu);
                if (count <= 0)
                {
                    Debug::Error(L"findMenuText: Menu has no items");
                    return {};
                }

                for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
                {
                    MENUITEMINFOW info{};
                    info.cbSize = sizeof(info);
                    info.fMask  = MIIM_ID | MIIM_SUBMENU;
                    if (GetMenuItemInfoW(currentMenu, pos, TRUE, &info) == 0)
                    {
                        continue;
                    }

                    if (info.hSubMenu)
                    {
                        std::wstring sub = self(self, info.hSubMenu, targetId);
                        if (! sub.empty())
                        {
                            return sub;
                        }
                    }

                    if (info.wID != targetId)
                    {
                        continue;
                    }

                    wchar_t raw[256]{};
                    const int len = GetMenuStringW(currentMenu, pos, raw, static_cast<int>(std::size(raw)), MF_BYPOSITION);
                    if (len <= 0)
                    {
                        return {};
                    }

                    return stripMenuText(std::wstring_view(raw, static_cast<size_t>(len)));
                }

                return {};
            };

            std::wstring text = findMenuText(findMenuText, menu, commandId);
            if (text.empty())
            {
                return;
            }

            static_cast<void>(customize->AddControlItem(kEncodingComboId, commandId, text.c_str()));
        };

        addMenuItemToCombo(IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL);
        if (hwnd)
        {
            HMENU rootMenu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
            if (rootMenu)
            {
                HMENU encodingMenu = nullptr;
                const int topCount = GetMenuItemCount(rootMenu);
                if (topCount <= 0)
                {
                    Debug::Error(L"addMenuItemToCombo: No top-level menu items");
                    return std::nullopt;
                }

                for (UINT pos = 0; pos < static_cast<UINT>(topCount); ++pos)
                {
                    MENUITEMINFOW info{};
                    info.cbSize = sizeof(info);
                    info.fMask  = MIIM_SUBMENU;
                    if (GetMenuItemInfoW(rootMenu, pos, TRUE, &info) == 0)
                    {
                        continue;
                    }

                    if (! info.hSubMenu)
                    {
                        continue;
                    }

                    if (GetMenuState(info.hSubMenu, IDM_VIEWER_ENCODING_DISPLAY_ANSI, MF_BYCOMMAND) != static_cast<UINT>(-1))
                    {
                        encodingMenu = info.hSubMenu;
                        break;
                    }
                }

                if (encodingMenu)
                {
                    auto addEncodingItems = [&](auto&& self, HMENU currentMenu) noexcept -> void
                    {
                        if (! currentMenu)
                        {
                            return;
                        }

                        const int count = GetMenuItemCount(currentMenu);
                        if (count <= 0)
                        {
                            Debug::Error(L"addMenuItemToCombo: Encoding menu has no items");
                            return;
                        }

                        for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
                        {
                            MENUITEMINFOW info{};
                            info.cbSize = sizeof(info);
                            info.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_SUBMENU;
                            if (GetMenuItemInfoW(currentMenu, pos, TRUE, &info) == 0)
                            {
                                continue;
                            }

                            if (info.hSubMenu)
                            {
                                self(self, info.hSubMenu);
                                continue;
                            }

                            if ((info.fType & MFT_SEPARATOR) != 0)
                            {
                                continue;
                            }

                            if (! IsEncodingMenuSelectionValid(info.wID))
                            {
                                continue;
                            }

                            wchar_t raw[256]{};
                            const int len = GetMenuStringW(currentMenu, pos, raw, static_cast<int>(std::size(raw)), MF_BYPOSITION);
                            if (len <= 0)
                            {
                                continue;
                            }

                            std::wstring text = stripMenuText(std::wstring_view(raw, static_cast<size_t>(len)));
                            if (text.empty())
                            {
                                continue;
                            }

                            static_cast<void>(customize->AddControlItem(kEncodingComboId, info.wID, text.c_str()));
                        }
                    };

                    addEncodingItems(addEncodingItems, encodingMenu);
                }
            }
        }

        static_cast<void>(customize->SetSelectedControlItem(kEncodingComboId, initialEncodingSelection));
    }

    const HRESULT showHr = dialog->Show(hwnd);
    if (FAILED(showHr))
    {
        return std::nullopt;
    }

    wil::com_ptr<IShellItem> item;
    const HRESULT itemHr = dialog->GetResult(item.put());
    if (FAILED(itemHr) || ! item)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string path;
    const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
    if (FAILED(nameHr) || ! path)
    {
        return std::nullopt;
    }

    DWORD selectedEncoding = initialEncodingSelection;
    if (customize)
    {
        static_cast<void>(customize->GetSelectedControlItem(kEncodingComboId, &selectedEncoding));
    }

    SaveAsResult result{};
    result.path              = std::filesystem::path(path.get());
    result.encodingSelection = static_cast<UINT>(selectedEncoding);
    return result;
}

void ViewerText::CommandOpen(HWND hwnd)
{
    const auto path = ShowOpenDialog(hwnd);
    if (! path.has_value())
    {
        return;
    }

    static_cast<void>(OpenPath(hwnd, path.value(), true));
}

void ViewerText::CommandSaveAs(HWND hwnd)
{
    if (_currentPath.empty())
    {
        return;
    }

    const auto dest = ShowSaveAsDialog(hwnd);
    if (! dest.has_value())
    {
        return;
    }

    const UINT encodingSelection = dest.value().encodingSelection;
    if (encodingSelection == IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL)
    {
        if (! _fileSystem)
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        wil::unique_handle outFile(
            CreateFileW(dest.value().path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! outFile)
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        wil::com_ptr<IFileSystemIO> fileIo;
        if (FAILED(_fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void())) || ! fileIo)
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        wil::com_ptr<IFileReader> reader;
        const HRESULT openHr = fileIo->CreateFileReader(_currentPath.c_str(), reader.put());
        if (FAILED(openHr) || ! reader)
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        uint64_t ignored     = 0;
        const HRESULT seekHr = reader->Seek(0, FILE_BEGIN, &ignored);
        if (FAILED(seekHr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        std::vector<uint8_t> buffer(256u * 1024u);
        for (;;)
        {
            unsigned long read   = 0;
            const HRESULT readHr = reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &read);
            if (FAILED(readHr))
            {
                ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
                return;
            }

            if (read == 0)
            {
                break;
            }

            DWORD written = 0;
            if (WriteFile(outFile.get(), buffer.data(), read, &written, nullptr) == 0 || written != read)
            {
                ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
                return;
            }
        }
        return;
    }

    if (! _hEdit)
    {
        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
        return;
    }

    wil::unique_handle outFile(CreateFileW(dest.value().path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! outFile)
    {
        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
        return;
    }

    struct SaveEncoding
    {
        enum class Kind : uint8_t
        {
            CodePage,
            Utf16LE,
            Utf16BE,
            Utf32LE,
            Utf32BE,
        };

        Kind kind     = Kind::CodePage;
        UINT codePage = CP_UTF8;
        bool writeBom = false;
    };

    auto resolveSaveEncoding = [&](UINT selection) noexcept -> SaveEncoding
    {
        if (selection == IDM_VIEWER_ENCODING_SAVE_UTF16LE_BOM || selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::Utf16LE};
        }
        if (selection == IDM_VIEWER_ENCODING_SAVE_UTF16BE_BOM || selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::Utf16BE};
        }
        if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::Utf32LE};
        }
        if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::Utf32BE};
        }
        if (selection == IDM_VIEWER_ENCODING_SAVE_UTF8_BOM || selection == IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::CodePage, .codePage = CP_UTF8, .writeBom = true};
        }
        if (selection == IDM_VIEWER_ENCODING_SAVE_UTF8 || selection == IDM_VIEWER_ENCODING_DISPLAY_UTF8)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::CodePage, .codePage = CP_UTF8, .writeBom = false};
        }
        if (selection == IDM_VIEWER_ENCODING_SAVE_ANSI || selection == IDM_VIEWER_ENCODING_DISPLAY_ANSI)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::CodePage, .codePage = CP_ACP, .writeBom = false};
        }

        const UINT codePage = CodePageForMenuSelection(selection);
        return SaveEncoding{.kind = SaveEncoding::Kind::CodePage, .codePage = codePage, .writeBom = false};
    };

    const SaveEncoding saveEncoding = resolveSaveEncoding(encodingSelection);

    if (saveEncoding.kind == SaveEncoding::Kind::CodePage && saveEncoding.writeBom && saveEncoding.codePage == CP_UTF8)
    {
        static constexpr uint8_t kBom[] = {0xEFu, 0xBBu, 0xBFu};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }
    else if (saveEncoding.kind == SaveEncoding::Kind::Utf16LE)
    {
        static constexpr uint8_t kBom[] = {0xFFu, 0xFEu};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }
    else if (saveEncoding.kind == SaveEncoding::Kind::Utf16BE)
    {
        static constexpr uint8_t kBom[] = {0xFEu, 0xFFu};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }
    else if (saveEncoding.kind == SaveEncoding::Kind::Utf32LE)
    {
        static constexpr uint8_t kBom[] = {0xFFu, 0xFEu, 0x00u, 0x00u};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }
    else if (saveEncoding.kind == SaveEncoding::Kind::Utf32BE)
    {
        static constexpr uint8_t kBom[] = {0x00u, 0x00u, 0xFEu, 0xFFu};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }

    struct SaveCookie
    {
        SaveCookie()                             = default;
        SaveCookie(const SaveCookie&)            = delete;
        SaveCookie& operator=(const SaveCookie&) = delete;
        SaveCookie(SaveCookie&&)                 = delete;
        SaveCookie& operator=(SaveCookie&&)      = delete;
        ~SaveCookie()                            = default;

        HANDLE file = nullptr;
        SaveEncoding encoding{};
        HRESULT error = S_OK;
        std::optional<wchar_t> pendingHighSurrogate;
        std::vector<wchar_t> wideScratch;
        std::vector<uint8_t> byteScratch;

        static bool IsHighSurrogate(wchar_t ch) noexcept
        {
            return ch >= 0xD800u && ch <= 0xDBFFu;
        }

        static bool IsLowSurrogate(wchar_t ch) noexcept
        {
            return ch >= 0xDC00u && ch <= 0xDFFFu;
        }

        static HRESULT WriteChunk(SaveCookie& cookie, const wchar_t* data, size_t count) noexcept
        {
            if (! data || count == 0)
            {
                return S_OK;
            }

            if (cookie.encoding.kind == SaveEncoding::Kind::Utf16LE)
            {
                return WriteAllHandle(cookie.file, data, count * sizeof(wchar_t));
            }

            if (cookie.encoding.kind == SaveEncoding::Kind::Utf16BE)
            {
                cookie.byteScratch.resize(count * sizeof(wchar_t));
                for (size_t i = 0; i < count; ++i)
                {
                    const uint16_t value            = static_cast<uint16_t>(data[i]);
                    cookie.byteScratch[i * 2u + 0u] = static_cast<uint8_t>((value >> 8u) & 0xFFu);
                    cookie.byteScratch[i * 2u + 1u] = static_cast<uint8_t>(value & 0xFFu);
                }

                return WriteAllHandle(cookie.file, cookie.byteScratch.data(), cookie.byteScratch.size());
            }

            if (cookie.encoding.kind == SaveEncoding::Kind::Utf32LE || cookie.encoding.kind == SaveEncoding::Kind::Utf32BE)
            {
                cookie.byteScratch.clear();
                cookie.byteScratch.reserve(count * 4u);

                for (size_t i = 0; i < count; ++i)
                {
                    const wchar_t ch = data[i];

                    uint32_t cp = 0;
                    if (IsHighSurrogate(ch))
                    {
                        if ((i + 1) < count && IsLowSurrogate(data[i + 1]))
                        {
                            const uint32_t hi = static_cast<uint32_t>(ch) - 0xD800u;
                            const uint32_t lo = static_cast<uint32_t>(data[i + 1]) - 0xDC00u;
                            cp                = 0x10000u + ((hi << 10u) | lo);
                            i += 1;
                        }
                        else
                        {
                            cp = 0xFFFDu;
                        }
                    }
                    else if (IsLowSurrogate(ch))
                    {
                        cp = 0xFFFDu;
                    }
                    else
                    {
                        cp = static_cast<uint32_t>(ch);
                    }

                    if (cookie.encoding.kind == SaveEncoding::Kind::Utf32LE)
                    {
                        cookie.byteScratch.push_back(static_cast<uint8_t>(cp & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 8u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 16u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 24u) & 0xFFu));
                    }
                    else
                    {
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 24u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 16u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 8u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>(cp & 0xFFu));
                    }
                }

                return WriteAllHandle(cookie.file, cookie.byteScratch.data(), cookie.byteScratch.size());
            }

            const UINT codePage = cookie.encoding.codePage;

            if (count > static_cast<size_t>(std::numeric_limits<int>::max()))
            {
                count = static_cast<size_t>(std::numeric_limits<int>::max());
            }

            const int srcLen   = static_cast<int>(count);
            const int required = WideCharToMultiByte(codePage, 0, data, srcLen, nullptr, 0, nullptr, nullptr);
            if (required <= 0)
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }

            cookie.byteScratch.resize(static_cast<size_t>(required));
            const int written = WideCharToMultiByte(codePage, 0, data, srcLen, reinterpret_cast<LPSTR>(cookie.byteScratch.data()), required, nullptr, nullptr);
            if (written <= 0)
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }

            return WriteAllHandle(cookie.file, cookie.byteScratch.data(), static_cast<size_t>(written));
        }

        static DWORD CALLBACK StreamOutCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG* pcb) noexcept
        {
            auto* cookie = reinterpret_cast<SaveCookie*>(dwCookie);
            if (! cookie || ! pbBuff || cb <= 0 || ! pcb)
            {
                return 1;
            }

            const size_t byteCount = static_cast<size_t>(cb);
            const size_t wideBytes = (byteCount / sizeof(wchar_t)) * sizeof(wchar_t);
            const size_t wideCount = wideBytes / sizeof(wchar_t);

            *pcb = static_cast<LONG>(wideBytes);
            if (wideCount == 0)
            {
                return 0;
            }

            const wchar_t* wide = reinterpret_cast<const wchar_t*>(pbBuff);

            cookie->wideScratch.clear();
            cookie->wideScratch.reserve(wideCount + 1);

            if (cookie->pendingHighSurrogate.has_value())
            {
                cookie->wideScratch.push_back(cookie->pendingHighSurrogate.value());
                cookie->pendingHighSurrogate.reset();
            }

            cookie->wideScratch.insert(cookie->wideScratch.end(), wide, wide + wideCount);

            if (! cookie->wideScratch.empty() && IsHighSurrogate(cookie->wideScratch.back()))
            {
                cookie->pendingHighSurrogate = cookie->wideScratch.back();
                cookie->wideScratch.pop_back();
            }

            const HRESULT hr = WriteChunk(*cookie, cookie->wideScratch.data(), cookie->wideScratch.size());
            if (FAILED(hr))
            {
                cookie->error = hr;
                return 1;
            }

            return 0;
        }
    };

    if (_textBuffer.empty() && _fileSize > _textStreamSkipBytes)
    {
        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
        return;
    }

    SaveCookie cookie;
    cookie.file     = outFile.get();
    cookie.encoding = saveEncoding;

    cookie.error = SaveCookie::WriteChunk(cookie, _textBuffer.data(), _textBuffer.size());

    if (cookie.pendingHighSurrogate.has_value() && SUCCEEDED(cookie.error))
    {
        static constexpr wchar_t kReplacement = static_cast<wchar_t>(0xFFFDu);
        cookie.error                          = SaveCookie::WriteChunk(cookie, &kReplacement, 1);
        cookie.pendingHighSurrogate.reset();
    }

    if (FAILED(cookie.error))
    {
        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
        return;
    }

    if (_textStreamActive)
    {
        ShowInlineAlert(InlineAlertSeverity::Info, IDS_VIEWERTEXT_NAME, IDS_VIEWERTEXT_MSG_STREAM_TRUNCATED);
    }
}

void ViewerText::CommandRefresh(HWND hwnd)
{
    if (_currentPath.empty())
    {
        return;
    }

    static_cast<void>(OpenPath(hwnd, _currentPath, false));
}

void ViewerText::CommandOtherNext(HWND hwnd)
{
    if (_otherFiles.size() <= 1)
    {
        return;
    }

    _otherIndex = (_otherIndex + 1) % _otherFiles.size();
    static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
}

void ViewerText::CommandOtherPrevious(HWND hwnd)
{
    if (_otherFiles.size() <= 1)
    {
        return;
    }

    if (_otherIndex == 0)
    {
        _otherIndex = _otherFiles.size() - 1;
    }
    else
    {
        _otherIndex -= 1;
    }

    static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
}

void ViewerText::CommandOtherFirst(HWND hwnd)
{
    if (_otherFiles.empty())
    {
        return;
    }

    _otherIndex = 0;
    static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
}

void ViewerText::CommandOtherLast(HWND hwnd)
{
    if (_otherFiles.empty())
    {
        return;
    }

    _otherIndex = _otherFiles.size() - 1;
    static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
}

void ViewerText::CommandFind(HWND hwnd)
{
    std::wstring query;
    const HRESULT hr =
        ShowViewerTextPromptDialog(hwnd, _hasTheme ? &_theme : nullptr, IDS_VIEWERTEXT_FIND_CAPTION, IDS_VIEWERTEXT_FIND_LABEL, _searchQuery, query);
    if (hr != S_OK)
    {
        return;
    }

    _searchQuery = std::move(query);
    UpdateSearchHighlights();
    if (_searchQuery.empty())
    {
        return;
    }

    CommandFindNext(hwnd, false);
}

void ViewerText::UpdateSearchHighlights() noexcept
{
    _searchMatchStarts.clear();

    _hexSearchNeedle.clear();
    _hexSearchNeedleValid = false;
    if (! _searchQuery.empty())
    {
        std::vector<uint8_t> needle;
        if (TryParseHexSearchNeedle(_searchQuery, needle))
        {
            if (! HexBigEndian())
            {
                std::reverse(needle.begin(), needle.end());
            }

            _hexSearchNeedle      = std::move(needle);
            _hexSearchNeedleValid = ! _hexSearchNeedle.empty();
        }
    }

    if (! _searchQuery.empty() && ! _textBuffer.empty() && _searchQuery.size() <= _textBuffer.size())
    {
        const size_t queryLen = _searchQuery.size();

        size_t pos = 0;
        while (pos < _textBuffer.size())
        {
            const size_t found = _textBuffer.find(_searchQuery, pos);
            if (found == std::wstring::npos)
            {
                break;
            }

            _searchMatchStarts.push_back(found);
            pos = found + queryLen;
        }
    }

    if (_hEdit)
    {
        InvalidateRect(_hEdit.get(), nullptr, TRUE);
    }
    if (_hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }
}

void ViewerText::CommandGoToOffset(HWND hwnd)
{
    std::wstring text;
    const HRESULT hr = ShowViewerTextPromptDialog(hwnd, _hasTheme ? &_theme : nullptr, IDS_VIEWERTEXT_GOTO_CAPTION, IDS_VIEWERTEXT_GOTO_LABEL, L"0", text);
    if (hr != S_OK)
    {
        return;
    }

    uint64_t value = 0;
    if (! TryParseOffset(text, value))
    {
        return;
    }

    CommandGoToOffsetValue(hwnd, value);
}

void ViewerText::CommandGoToTop(HWND hwnd, bool extendSelection) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (_viewMode == ViewMode::Hex)
    {
        if (! _hHex || _fileSize == 0)
        {
            return;
        }

        uint64_t offset = 0;
        if (_hexSelectedOffset.has_value())
        {
            offset = _hexSelectedOffset.value();
        }
        else
        {
            offset = _hexTopLine * static_cast<uint64_t>(kHexBytesPerLine);
            if (offset >= _fileSize)
            {
                offset = _fileSize - 1;
            }
        }

        const uint64_t nextOffset = 0;
        if (extendSelection)
        {
            if (! _hexSelectionAnchorOffset.has_value())
            {
                _hexSelectionAnchorOffset = offset;
            }
        }
        else
        {
            _hexSelectionAnchorOffset = nextOffset;
        }

        _hexSelectedOffset = nextOffset;
        _hexTopLine        = 0;
        UpdateHexViewScrollBars(_hHex.get());
        InvalidateRect(_hHex.get(), nullptr, TRUE);
        InvalidateRect(hwnd, &_statusRect, FALSE);
        return;
    }

    if (! _hEdit)
    {
        return;
    }

    if (_textStreamActive && _textStreamStartOffset > _textStreamSkipBytes)
    {
        static_cast<void>(LoadTextToEdit(hwnd, _textStreamSkipBytes, false));
        return;
    }

    _textTopVisualLine = 0;
    _textLeftColumn    = 0;

    const size_t newCaret = 0;
    _textCaretIndex       = newCaret;
    if (! extendSelection)
    {
        _textSelAnchor = newCaret;
    }
    _textSelActive       = newCaret;
    _textPreferredColumn = 0;

    UpdateTextViewScrollBars(_hEdit.get());
    InvalidateRect(_hEdit.get(), nullptr, TRUE);
    InvalidateRect(hwnd, &_statusRect, FALSE);
}

void ViewerText::CommandGoToBottom(HWND hwnd, bool extendSelection) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (_viewMode == ViewMode::Hex)
    {
        if (! _hHex || _fileSize == 0)
        {
            return;
        }

        uint64_t offset = 0;
        if (_hexSelectedOffset.has_value())
        {
            offset = _hexSelectedOffset.value();
        }
        else
        {
            offset = _hexTopLine * static_cast<uint64_t>(kHexBytesPerLine);
            if (offset >= _fileSize)
            {
                offset = _fileSize - 1;
            }
        }

        const uint64_t nextOffset = _fileSize - 1;
        if (extendSelection)
        {
            if (! _hexSelectionAnchorOffset.has_value())
            {
                _hexSelectionAnchorOffset = offset;
            }
        }
        else
        {
            _hexSelectionAnchorOffset = nextOffset;
        }

        _hexSelectedOffset = nextOffset;

        const uint64_t targetLine = nextOffset / static_cast<uint64_t>(kHexBytesPerLine);

        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_PAGE;
        static_cast<void>(GetScrollInfo(_hHex.get(), SB_VERT, &si));
        const uint64_t pageLines = std::max<uint64_t>(1u, static_cast<uint64_t>(si.nPage == 0 ? 1u : si.nPage));

        if (targetLine < _hexTopLine)
        {
            _hexTopLine = targetLine;
        }
        else if (targetLine >= _hexTopLine + pageLines)
        {
            _hexTopLine = targetLine - pageLines + 1;
        }

        UpdateHexViewScrollBars(_hHex.get());
        InvalidateRect(_hHex.get(), nullptr, TRUE);
        InvalidateRect(hwnd, &_statusRect, FALSE);
        return;
    }

    if (! _hEdit)
    {
        return;
    }

    if (_textStreamActive && _fileSize > 0 && _textStreamEndOffset < _fileSize)
    {
        uint64_t lastStart        = _textStreamSkipBytes;
        const uint64_t chunkBytes = TextStreamChunkBytes();
        if (_fileSize > chunkBytes)
        {
            lastStart = _fileSize - chunkBytes;
        }
        lastStart = AlignTextStreamOffset(lastStart);
        static_cast<void>(LoadTextToEdit(hwnd, lastStart, true));
    }

    if (_textVisualLineStarts.empty())
    {
        return;
    }

    RECT client{};
    GetClientRect(_hEdit.get(), &client);
    const UINT dpi        = GetDpiForWindow(_hEdit.get());
    const float heightDip = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
    const float marginDip = 6.0f;
    const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
    const float usableDip = std::max(0.0f, heightDip - 2.0f * marginDip);
    const uint32_t rows   = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(usableDip / std::max(1.0f, lineH))));

    const uint32_t totalVisual = static_cast<uint32_t>(_textVisualLineStarts.size());
    const uint32_t lastVisual  = totalVisual > 0 ? (totalVisual - 1) : 0;
    const uint32_t desiredTop  = (totalVisual > rows) ? (totalVisual - rows) : 0;

    _textTopVisualLine = std::min<uint32_t>(desiredTop, lastVisual);

    const size_t newCaret = _textBuffer.size();
    _textCaretIndex       = newCaret;
    if (! extendSelection)
    {
        _textSelAnchor = newCaret;
    }
    _textSelActive       = newCaret;
    _textPreferredColumn = 0;

    UpdateTextViewScrollBars(_hEdit.get());
    InvalidateRect(_hEdit.get(), nullptr, TRUE);
    InvalidateRect(hwnd, &_statusRect, FALSE);
}

HRESULT ViewerText::DetectEncodingAndSize(const std::filesystem::path& path, FileEncoding& encoding, uint64_t& bomBytes, uint64_t& fileSize) noexcept
{
    encoding = FileEncoding::Unknown;
    bomBytes = 0;
    fileSize = 0;

    if (! _fileReader)
    {
        Debug::Error(L"ViewerText: DetectEncodingAndSize failed because file reader is missing for '{}'.", path.c_str());
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }

    uint64_t sizeBytes   = 0;
    const HRESULT sizeHr = _fileReader->GetSize(&sizeBytes);
    if (FAILED(sizeHr))
    {
        Debug::Error(L"ViewerText: GetSize failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(sizeHr));
        return sizeHr;
    }

    fileSize = sizeBytes;

    BYTE bom[4]{};
    unsigned long read = 0;

    uint64_t pos         = 0;
    const HRESULT seekHr = _fileReader->Seek(0, FILE_BEGIN, &pos);
    if (FAILED(seekHr))
    {
        Debug::Error(L"ViewerText: Seek(FILE_BEGIN, 0) failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(seekHr));
        return seekHr;
    }

    const HRESULT readHr = _fileReader->Read(bom, static_cast<unsigned long>(std::size(bom)), &read);
    if (FAILED(readHr))
    {
        Debug::Error(L"ViewerText: Read failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(readHr));
        return readHr;
    }

    if (read >= 4 && bom[0] == 0xFF && bom[1] == 0xFE && bom[2] == 0x00 && bom[3] == 0x00)
    {
        encoding = FileEncoding::Utf32LE;
        bomBytes = 4;
        return S_OK;
    }
    if (read >= 4 && bom[0] == 0x00 && bom[1] == 0x00 && bom[2] == 0xFE && bom[3] == 0xFF)
    {
        encoding = FileEncoding::Utf32BE;
        bomBytes = 4;
        return S_OK;
    }
    if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
    {
        encoding = FileEncoding::Utf8;
        bomBytes = 3;
        return S_OK;
    }
    if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
    {
        encoding = FileEncoding::Utf16LE;
        bomBytes = 2;
        return S_OK;
    }
    if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
    {
        encoding = FileEncoding::Utf16BE;
        bomBytes = 2;
        return S_OK;
    }

    encoding = FileEncoding::Unknown;
    bomBytes = 0;
    return S_OK;
}

std::wstring ViewerText::EncodingLabel() const
{
    UINT id = IDS_VIEWERTEXT_ENCODING_UNKNOWN;
    switch (_encoding)
    {
        case FileEncoding::Utf8: id = IDS_VIEWERTEXT_ENCODING_UTF8; break;
        case FileEncoding::Utf16LE: id = IDS_VIEWERTEXT_ENCODING_UTF16LE; break;
        case FileEncoding::Utf16BE: id = IDS_VIEWERTEXT_ENCODING_UTF16BE; break;
        case FileEncoding::Utf32LE: id = IDS_VIEWERTEXT_ENCODING_UTF32LE; break;
        case FileEncoding::Utf32BE: id = IDS_VIEWERTEXT_ENCODING_UTF32BE; break;
        case FileEncoding::Unknown: id = IDS_VIEWERTEXT_ENCODING_UNKNOWN; break;
        default: id = IDS_VIEWERTEXT_ENCODING_UNKNOWN; break;
    }

    return LoadStringResource(g_hInstance, id);
}

std::wstring ViewerText::BuildStatusText() const
{
    auto withStatusMessage = [&](std::wstring base) -> std::wstring
    {
        std::wstring combined = std::move(base);

        if (_viewMode == ViewMode::Text && _textStreamActive && ! _isLoading)
        {
            const std::wstring streamingMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_STREAM_TRUNCATED);
            if (! streamingMessage.empty())
            {
                const std::wstring streamingCombined = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_STATUS_WITH_MESSAGE_FORMAT, streamingMessage, combined);
                if (! streamingCombined.empty())
                {
                    combined = streamingCombined;
                }
            }
        }

        if (! _statusMessage.empty())
        {
            const std::wstring statusCombined = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_STATUS_WITH_MESSAGE_FORMAT, _statusMessage, combined);
            if (! statusCombined.empty())
            {
                combined = statusCombined;
            }
        }

        return combined;
    };

    std::wstring detected;
    if (_encoding != FileEncoding::Unknown)
    {
        detected = EncodingLabel();

        if (_bomBytes > 0)
        {
            detected.append(LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DETECTED_SUFFIX_BOM));
        }
    }
    else if (_detectedCodePageValid)
    {
        if (_detectedCodePage == CP_UTF8)
        {
            detected = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_ENCODING_UTF8);
        }
        else
        {
            detected = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_CODEPAGE_FORMAT, _detectedCodePage);
        }

        if (_detectedCodePageIsGuess)
        {
            detected.append(LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DETECTED_SUFFIX_GUESS));
        }
    }
    else
    {
        detected = EncodingLabel();
    }

    auto stripMenuText = [](std::wstring_view text) -> std::wstring
    {
        const size_t tabPos = text.find(L'\t');
        if (tabPos != std::wstring_view::npos)
        {
            text = text.substr(0, tabPos);
        }

        std::wstring result;
        result.reserve(text.size());

        for (size_t i = 0; i < text.size(); ++i)
        {
            const wchar_t ch = text[i];
            if (ch != L'&')
            {
                result.push_back(ch);
                continue;
            }

            if ((i + 1) < text.size() && text[i + 1] == L'&')
            {
                result.push_back(L'&');
                i += 1;
            }
        }

        while (! result.empty() && result.front() == L' ')
        {
            result.erase(result.begin());
        }
        while (! result.empty() && result.back() == L' ')
        {
            result.pop_back();
        }

        return result;
    };

    const UINT selection = EffectiveDisplayEncodingMenuSelection();
    std::wstring active;
    if (_hWnd)
    {
        HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(_hWnd.get());
        if (menu)
        {
            wchar_t buffer[256]{};
            const int len = GetMenuStringW(menu, selection, buffer, static_cast<int>(std::size(buffer)), MF_BYCOMMAND);
            if (len > 0)
            {
                active.assign(buffer, static_cast<size_t>(len));
            }
        }
    }

    active = stripMenuText(active);

    const std::wstring sizeText = FormatBytesCompact(_fileSize);

    if (_viewMode == ViewMode::Hex)
    {
        uint64_t topOffset    = 0;
        uint64_t bottomOffset = 0;

        if (_fileSize > 0 && _hHex)
        {
            const uint64_t maxByte  = _fileSize - 1;
            const uint64_t topStart = _hexTopLine * static_cast<uint64_t>(kHexBytesPerLine);

            RECT client{};
            GetClientRect(_hHex.get(), &client);
            const UINT dpi        = GetDpiForWindow(_hHex.get());
            const float heightDip = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
            const float marginDip = 6.0f;
            const float lineH     = (_hexLineHeightDip > 0.0f) ? _hexLineHeightDip : 14.0f;
            const float headerH   = lineH;
            const float usableDip = std::max(0.0f, heightDip - headerH - 2.0f * marginDip);
            const uint32_t rows   = std::max<uint32_t>(1u, static_cast<uint32_t>(std::ceil(usableDip / std::max(1.0f, lineH))));

            const uint64_t bottomLine  = (_hexTopLine + static_cast<uint64_t>(rows) > 0) ? (_hexTopLine + static_cast<uint64_t>(rows) - 1) : 0;
            const uint64_t bottomStart = bottomLine * static_cast<uint64_t>(kHexBytesPerLine);

            topOffset    = std::min(topStart, maxByte);
            bottomOffset = std::min(maxByte, std::min(bottomStart, maxByte) + static_cast<uint64_t>(kHexBytesPerLine - 1));
        }

        return withStatusMessage(FormatStringResource(g_hInstance,
                                                      IDS_VIEWERTEXT_STATUS_HEX_FORMAT,
                                                      _fileSystemName,
                                                      detected,
                                                      active,
                                                      sizeText,
                                                      FormatFileOffset(topOffset),
                                                      FormatFileOffset(bottomOffset)));
    }

    int topLine    = 1;
    int bottomLine = 1;

    if (_hEdit && ! _textVisualLineStarts.empty() && ! _textVisualLineLogical.empty())
    {
        RECT client{};
        GetClientRect(_hEdit.get(), &client);
        const UINT dpi        = GetDpiForWindow(_hEdit.get());
        const float heightDip = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
        const float marginDip = 6.0f;
        const float usableDip = std::max(0.0f, heightDip - 2.0f * marginDip);
        const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
        const uint32_t rows   = std::max<uint32_t>(1u, static_cast<uint32_t>(std::ceil(usableDip / std::max(1.0f, lineH))));

        const uint32_t totalVisual  = static_cast<uint32_t>(_textVisualLineStarts.size());
        const uint32_t topVisual    = std::min<uint32_t>(_textTopVisualLine, totalVisual - 1);
        const uint32_t bottomVisual = std::min<uint32_t>(totalVisual - 1, topVisual + rows - 1);

        const uint32_t topLogical    = std::min<uint32_t>(_textVisualLineLogical[topVisual], static_cast<uint32_t>(_textLineStarts.size() - 1));
        const uint32_t bottomLogical = std::min<uint32_t>(_textVisualLineLogical[bottomVisual], static_cast<uint32_t>(_textLineStarts.size() - 1));

        topLine    = static_cast<int>(topLogical) + 1;
        bottomLine = static_cast<int>(bottomLogical) + 1;
    }

    std::wstring totalLinesText = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_UNKNOWN);
    if (! _isLoading)
    {
        if (_textTotalLineCount.has_value())
        {
            totalLinesText = std::format(L"{:L}", _textTotalLineCount.value());
        }
        else if (! _textStreamActive && ! _textLineStarts.empty())
        {
            const uint64_t totalLines = static_cast<uint64_t>(_textLineStarts.size());
            totalLinesText            = std::format(L"{:L}", totalLines);
        }
    }

    return withStatusMessage(
        FormatStringResource(g_hInstance, IDS_VIEWERTEXT_STATUS_TEXT_FORMAT, _fileSystemName, detected, active, sizeText, topLine, bottomLine, totalLinesText));
}

HRESULT ViewerText::OpenPath(HWND hwnd, const std::filesystem::path& path, bool updateOtherFiles) noexcept
{
    if (path.empty())
    {
        Debug::Error(L"ViewerText: OpenPath called with an empty path.");
        return E_INVALIDARG;
    }

    StartAsyncOpen(hwnd, path, updateOtherFiles, 0);
    return S_OK;
}

bool ViewerText::HandleShortcutKey(HWND hwnd, WPARAM vk) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

    if (vk == VK_ESCAPE)
    {
        CommandExit(hwnd);
        return true;
    }

    if (vk == VK_SPACE)
    {
        CommandOtherNext(hwnd);
        return true;
    }

    if (vk == VK_BACK)
    {
        CommandOtherPrevious(hwnd);
        return true;
    }

    if (ctrl && vk == VK_RIGHT)
    {
        return false;
    }

    if (ctrl && vk == VK_LEFT)
    {
        return false;
    }

    if (ctrl && vk == VK_UP)
    {
        CommandOtherPrevious(hwnd);
        return true;
    }

    if (ctrl && vk == VK_DOWN)
    {
        CommandOtherNext(hwnd);
        return true;
    }

    if (ctrl && vk == VK_HOME)
    {
        CommandOtherFirst(hwnd);
        return true;
    }

    if (ctrl && vk == VK_END)
    {
        CommandOtherLast(hwnd);
        return true;
    }

    if (ctrl && (vk == 'F' || vk == 'f'))
    {
        CommandFind(hwnd);
        return true;
    }

    if (vk == VK_F7)
    {
        return NavigateDiffHunk(hwnd, shift);
    }

    if (vk == VK_F3)
    {
        CommandFindNext(hwnd, shift);
        return true;
    }

    if (ctrl && (vk == 'G' || vk == 'g'))
    {
        CommandGoToOffset(hwnd);
        return true;
    }

    if (ctrl && (vk == 'O' || vk == 'o'))
    {
        CommandOpen(hwnd);
        return true;
    }

    if (ctrl && (vk == 'S' || vk == 's'))
    {
        CommandSaveAs(hwnd);
        return true;
    }

    if (vk == VK_F5)
    {
        CommandRefresh(hwnd);
        return true;
    }

    if (vk == VK_F8)
    {
        CommandCycleDisplayEncoding(hwnd, shift);
        return true;
    }

    return false;
}

HRESULT STDMETHODCALLTYPE ViewerText::Open(const ViewerOpenContext* context) noexcept
{
    if (! context || ! context->focusedPath || context->focusedPath[0] == L'\0')
    {
        Debug::Error(L"ViewerText: Open called with an invalid context (focusedPath missing).");
        return E_INVALIDARG;
    }

    if (! context->fileSystem)
    {
        Debug::Error(L"ViewerText: Open called with an invalid context (fileSystem missing).");
        return E_INVALIDARG;
    }

    _fileSystem = context->fileSystem;

    _fileSystemName.clear();
    if (context->fileSystemName && context->fileSystemName[0] != L'\0')
    {
        _fileSystemName = context->fileSystemName;
    }

    _selection.clear();
    if (context->selectionPaths && context->selectionCount > 0)
    {
        for (unsigned long i = 0; i < context->selectionCount; ++i)
        {
            const wchar_t* p = context->selectionPaths[i];
            if (p && p[0] != L'\0')
            {
                _selection.emplace_back(p);
            }
        }
    }

    _otherFiles.clear();
    if (context->otherFiles && context->otherFileCount > 0)
    {
        for (unsigned long i = 0; i < context->otherFileCount; ++i)
        {
            const wchar_t* p = context->otherFiles[i];
            if (p && p[0] != L'\0')
            {
                _otherFiles.emplace_back(p);
            }
        }
    }

    _otherIndex = 0;
    if (! _otherFiles.empty() && context->focusedOtherFileIndex < _otherFiles.size())
    {
        _otherIndex = static_cast<size_t>(context->focusedOtherFileIndex);
    }

    if ((context->flags & VIEWER_OPEN_FLAG_START_HEX) != 0)
    {
        _viewMode = ViewMode::Hex;
    }

    const bool embeddedMode = IsEmbeddedOpen(*context);
    const HWND embeddedParent = embeddedMode ? context->ownerWindow : nullptr;
    if (embeddedMode && (! embeddedParent || IsWindow(embeddedParent) == FALSE))
    {
        Debug::Error(L"ViewerText: embedded Open called without a valid parent window.");
        return E_INVALIDARG;
    }

    if (ShouldRecreateViewerWindow(embeddedMode, embeddedParent))
    {
        static_cast<void>(Close());
    }

    const std::filesystem::path path(context->focusedPath);

    if (! _hWnd)
    {
        if (! RegisterWndClass(g_hInstance))
        {
            return E_FAIL;
        }

        _embeddedMode = embeddedMode;
        HWND ownerWindow = context->ownerWindow;
        const std::wstring initialTitle = embeddedMode ? std::wstring{} : (_metaName.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERTEXT_NAME) : _metaName);

        RECT ownerRect{};
        if (embeddedMode)
        {
            RECT parentClient{};
            GetClientRect(embeddedParent, &parentClient);
            HWND window = CreateWindowExW(0,
                                          kClassName,
                                          L"",
                                          WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                          0,
                                          0,
                                          std::max(1L, parentClient.right - parentClient.left),
                                          std::max(1L, parentClient.bottom - parentClient.top),
                                          embeddedParent,
                                          nullptr,
                                          g_hInstance,
                                          this);
            if (! window)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerText: embedded CreateWindowExW failed.");
                return HRESULT_FROM_WIN32(lastError);
            }

            _hWnd.reset(window);
            ApplyTheme(_hWnd.get());
            ApplyPendingViewerTextClassBackgroundBrush(_hWnd.get(), _hEdit.get(), _hHex.get());

            AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
            ShowWindow(_hWnd.get(), SW_SHOWNA);
        }
        else if (ownerWindow && GetWindowRect(ownerWindow, &ownerRect) != 0)
        {
            const int w = ownerRect.right - ownerRect.left;
            const int h = ownerRect.bottom - ownerRect.top;

            wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(Localization::LoadMenuResource(g_hInstance, IDR_VIEWERTEXT_MENU));
            HWND window = CreateWindowExW(0,
                                          kClassName,
                                          initialTitle.c_str(),
                                          WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                                          ownerRect.left,
                                          ownerRect.top,
                                          std::max(1, w),
                                          std::max(1, h),
                                          nullptr,
                                          menu.get(),
                                          g_hInstance,
                                          this);
            if (! window)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed.");
                return HRESULT_FROM_WIN32(lastError);
            }

            menu.release();

            _hWnd.reset(window);

            if (! _windowIconSmall)
            {
                _windowIconSmall = CreateViewerTextIcon(16);
            }
            if (! _windowIconBig)
            {
                _windowIconBig = CreateViewerTextIcon(32);
            }
            if (_windowIconSmall)
            {
                static_cast<void>(SendMessageW(_hWnd.get(), WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(_windowIconSmall.get())));
            }
            if (_windowIconBig)
            {
                static_cast<void>(SendMessageW(_hWnd.get(), WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(_windowIconBig.get())));
            }

            ApplyTheme(_hWnd.get());
            ApplyPendingViewerTextClassBackgroundBrush(_hWnd.get(), _hEdit.get(), _hHex.get());

            AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
        else
        {
            wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(Localization::LoadMenuResource(g_hInstance, IDR_VIEWERTEXT_MENU));
            HWND window = CreateWindowExW(0,
                                          kClassName,
                                          initialTitle.c_str(),
                                          WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          900,
                                          700,
                                          nullptr,
                                          menu.get(),
                                          g_hInstance,
                                          this);
            if (! window)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed.");
                return HRESULT_FROM_WIN32(lastError);
            }

            menu.release();
            _hWnd.reset(window);

            if (! _windowIconSmall)
            {
                _windowIconSmall = CreateViewerTextIcon(16);
            }
            if (! _windowIconBig)
            {
                _windowIconBig = CreateViewerTextIcon(32);
            }
            if (_windowIconSmall)
            {
                static_cast<void>(SendMessageW(_hWnd.get(), WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(_windowIconSmall.get())));
            }
            if (_windowIconBig)
            {
                static_cast<void>(SendMessageW(_hWnd.get(), WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(_windowIconBig.get())));
            }

            ApplyTheme(_hWnd.get());
            ApplyPendingViewerTextClassBackgroundBrush(_hWnd.get(), _hEdit.get(), _hHex.get());

            AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }
    else
    {
        ApplyPendingViewerTextClassBackgroundBrush(_hWnd.get(), _hEdit.get(), _hHex.get());
        if (_embeddedMode)
        {
            RECT parentClient{};
            if (embeddedParent && GetClientRect(embeddedParent, &parentClient) != FALSE)
            {
                SetWindowPos(_hWnd.get(),
                             HWND_TOP,
                             0,
                             0,
                             std::max(1L, parentClient.right - parentClient.left),
                             std::max(1L, parentClient.bottom - parentClient.top),
                             SWP_NOACTIVATE);
            }
            ShowWindow(_hWnd.get(), SW_SHOWNA);
        }
        else
        {
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }

    if (! _hWnd)
    {
        Debug::Error(L"ViewerText: Open failed because viewer window is missing after creation.");
        return E_FAIL;
    }

    StartAsyncOpen(_hWnd.get(), path, false, 0);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::Close() noexcept
{
    AddRef();
    const auto releaseSelf = wil::scope_exit([&]() noexcept { Release(); });
    _hWnd.reset();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::SetTheme(const ViewerTheme* theme) noexcept
{
    if (! theme || theme->version < 2u || theme->version > 4u)
    {
        return E_INVALIDARG;
    }

    _theme = *theme;
    PopulateViewerDiffThemeDefaults(_theme);
    _hasTheme = true;

    RequestViewerTextClassBackgroundColor(ColorRefFromArgb(_theme.backgroundArgb));
    ApplyPendingViewerTextClassBackgroundBrush(_hWnd.get(), _hEdit.get(), _hHex.get());

    if (_hWnd)
    {
        ApplyTheme(_hWnd.get());
        static_cast<void>(RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW));
    }

    return S_OK;
}

namespace
{
[[nodiscard]] WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept
{
    return RedSalamander::Win32Callback::GetStoredWndProc(hwnd, propName);
}

[[nodiscard]] bool InstallWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp, WNDPROC hookWndProc) noexcept
{
    if (! hwnd || ! originalWndProcProp || ! hookWndProc)
    {
        return false;
    }

    if (GetStoredWndProc(hwnd, originalWndProcProp))
    {
        return true;
    }

    const auto originalWndProc = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(hwnd, GWLP_WNDPROC));
    if (! originalWndProc)
    {
        return false;
    }

    if (! RedSalamander::Win32Callback::SetPropNoThrow(hwnd, originalWndProcProp, reinterpret_cast<HANDLE>(originalWndProc)))
    {
        return false;
    }

    const auto previousWndProc =
        reinterpret_cast<WNDPROC>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(hookWndProc)));
    if (previousWndProc != originalWndProc)
    {
        RemovePropW(hwnd, originalWndProcProp);
        if (previousWndProc)
        {
            static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(previousWndProc)));
        }
        return false;
    }

    return true;
}

[[nodiscard]] LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* originalWndProcProp, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (const auto originalWndProc = GetStoredWndProc(hwnd, originalWndProcProp))
    {
        return RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK FileComboHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

void UnhookFileComboHostWindow(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    RemovePropW(hwnd, kFileComboHostStateProp);
    RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kFileComboHostOriginalWndProcProp, FileComboHostWndProc);
}

[[nodiscard]] bool MessageMayOpenWindowComboPopup(UINT msg, WPARAM wp) noexcept
{
    switch (msg)
    {
        case WM_LBUTTONDOWN:
        case WM_LBUTTONDBLCLK: return true;
        case WM_SYSKEYDOWN:
        {
            const UINT vk = static_cast<UINT>(wp);
            return vk == VK_DOWN || vk == VK_UP;
        }
        case WM_KEYDOWN:
        {
            const UINT vk = static_cast<UINT>(wp);
            return vk == VK_SPACE || vk == VK_RETURN || vk == VK_F4 || vk == VK_DOWN || vk == VK_UP;
        }
        default: return false;
    }
}

[[nodiscard]] int ComputeWindowComboPopupHeightPx(size_t itemCount, UINT dpi) noexcept
{
    const size_t visibleRows = std::max<size_t>(1u, std::min(itemCount, kViewerComboPopupMaxVisibleItems));
    const int popupHeightDip = 2 + 8 + (24 * static_cast<int>(visibleRows));
    return std::max(0, MulDiv(popupHeightDip, static_cast<int>(dpi), 96));
}
} // namespace
