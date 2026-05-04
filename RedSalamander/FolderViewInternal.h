#pragma once

// Internal implementation header for FolderView split across multiple .cpp files.
// Keep this header private to the FolderView translation units.

#include <algorithm>
#include <array>
#include <atomic>
#include <cassert>
#include <chrono>
#include <cmath>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <exception>
#include <execution>
#include <format>
#include <iterator>
#include <limits>
#include <memory>
#include <new>
#include <unordered_map>
#include <unordered_set>

#define WINDOWS_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <commctrl.h>
#include <commdlg.h>
#include <dwmapi.h>
#include <propvarutil.h>
#include <shellapi.h>
#include <shellscalingapi.h>
#include <shlobj_core.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <uxtheme.h>

#include <sstream>
#include <system_error>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/result.h>
#pragma warning(pop)

#include <wincodec.h>
#include <wincodecsdk.h>

#include "Helpers.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "Ui/AlertOverlay.h"

#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "FolderView.h"
#include "FolderViewEmptyStateLayout.h"
#include "FolderViewVisualState.h"
#include "Helpers.h"
#include "HostServices.h"
#include "IconCache.h"
#include "ThemedInputFrames.h"
#include "UiMetrics.h"
#include "WindowMessages.h"
#include "resource.h"

#ifndef GET_X_LPARAM
#define GET_X_LPARAM(lp) ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp) ((int)(short)HIWORD(lp))
#endif

#ifndef CLSID_WICImagingFactory2
#define CLSID_WICImagingFactory2 CLSID_WICImagingFactory
#endif

#ifndef WICBitmapAlphaChannelOptionUseBitmapAlpha
// Fallback definition for older Windows SDK versions
#pragma warning(push)
#pragma warning(disable : 5264) // C5264: 'const' variable is not used
inline constexpr WICBitmapAlphaChannelOption WICBitmapAlphaChannelOptionUseBitmapAlpha_Fallback = static_cast<WICBitmapAlphaChannelOption>(2);
#pragma warning(pop)
#define WICBitmapAlphaChannelOptionUseBitmapAlpha WICBitmapAlphaChannelOptionUseBitmapAlpha_Fallback
#endif

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Dwmapi.lib")

using wil::unique_hbitmap;

#pragma warning(push)
// 5245 : unreferenced function with internal linkage has been removed
#pragma warning(disable : 5245)

namespace
{
constexpr wchar_t kFolderViewClassName[]              = L"RedSalamanderFolderView";
constexpr float kLabelHorizontalPaddingDip            = 12.0f;
constexpr float kLabelVerticalPaddingDip              = 4.0f;
constexpr float kFocusStrokeThicknessDip              = 2.0f;
constexpr float kFocusStrokeThicknessUnfocusedDip     = 1.0f;
constexpr float kFocusBorderOpacityUnfocused          = FolderViewVisualState::kFocusBorderOpacityUnfocused;
constexpr float kUnfocusedPaneTextOpacity             = FolderViewVisualState::kUnfocusedPaneTextOpacity;
constexpr float kUnfocusedPaneIconOpacity             = FolderViewVisualState::kUnfocusedPaneIconOpacity;
constexpr float kSelectionCornerRadiusDip             = 2.0f;
constexpr float kIconTextGapDip                       = 12.0f;
constexpr float kFolderViewListIconSizeDip            = 16.0f;
constexpr float kFolderViewThumbnailIconSizeDip       = 64.0f;
constexpr float kColumnSpacingDip                     = 18.0f;
constexpr float kRowSpacingDip                        = 4.0f;
constexpr float kCompactRowSpacingDip                 = 0.0f;
constexpr float kDetailsGapDip                        = 2.0f;
constexpr float kDetailsTextAlpha                     = 0.75f;
constexpr float kMetadataTextAlpha                    = 0.55f;
constexpr UINT kSwapChainBufferCount                  = 2;
constexpr UINT_PTR kOverlayTimerId                    = 1;
constexpr uint64_t kBusyOverlayDelayMs                = 300;
constexpr uint64_t kOperationInfoOverlayAutoDismissMs = 1800;
constexpr uint64_t kOverlayTimerRetryMs               = 120;

[[nodiscard]] constexpr float GetFolderViewRowSpacingDip(const AppTheme& appTheme) noexcept
{
    return appTheme.compactMode ? kCompactRowSpacingDip : kRowSpacingDip;
}

bool ConfirmNonRevertableFileOperation(HWND owner,
                                       [[maybe_unused]] IFileSystem* fileSystem,
                                       FileSystemOperation operation,
                                       const std::vector<std::filesystem::path>& sourcePaths,
                                       const std::filesystem::path& destinationFolder) noexcept
{
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return true;
    }

    if (sourcePaths.empty())
    {
        return true;
    }

    // Avoid I/O in the confirmation prompt path (plugins may require network access to answer GetAttributes).
    // Best-effort: treat item types as unknown.
    unsigned long long fileCount    = 0;
    unsigned long long folderCount  = 0;
    unsigned long long unknownCount = static_cast<unsigned long long>(sourcePaths.size());
    std::filesystem::path sampleFile;
    bool hasSampleFile = false;

    auto suffixFor = [](unsigned long long count) noexcept -> std::wstring_view { return count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s"); };

    const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
    std::wstring what;
    if (unknownCount > 0)
    {
        const std::wstring_view itemSuffix = suffixFor(itemCount);
        what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ITEM, itemCount, itemSuffix);
    }
    else if (fileCount > 0 && folderCount > 0)
    {
        const std::wstring_view fileSuffix   = suffixFor(fileCount);
        const std::wstring_view folderSuffix = suffixFor(folderCount);
        what = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILES_FOLDERS, fileCount, fileSuffix, folderCount, folderSuffix);
    }
    else if (fileCount > 0)
    {
        const std::wstring_view fileSuffix = suffixFor(fileCount);
        what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILE, fileCount, fileSuffix);
    }
    else
    {
        const std::wstring_view folderSuffix = suffixFor(folderCount);
        what                                 = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FOLDER, folderCount, folderSuffix);
    }

    auto ensureTrailingSeparator = [](std::wstring text) noexcept -> std::wstring
    {
        if (text.empty())
        {
            return text;
        }

        const wchar_t last = text.back();
        if (last == L'\\' || last == L'/')
        {
            return text;
        }

        text.push_back(L'\\');
        return text;
    };

    auto normalizeSlashes = [](std::wstring& text) noexcept
    {
        for (auto& ch : text)
        {
            if (ch == L'/')
            {
                ch = L'\\';
            }
        }
    };

    std::wstring fromText;
    if (sourcePaths.size() == 1u)
    {
        fromText = sourcePaths.front().wstring();
        if (unknownCount == 0 && folderCount == 1ull && fileCount == 0ull)
        {
            fromText = ensureTrailingSeparator(std::move(fromText));
        }
    }
    else
    {
        std::filesystem::path commonParent = sourcePaths.front().parent_path();
        bool multipleParents               = false;
        for (size_t index = 1; index < sourcePaths.size(); ++index)
        {
            const std::filesystem::path parent = sourcePaths[index].parent_path();
            if (CompareStringOrdinal(commonParent.c_str(), -1, parent.c_str(), -1, TRUE) != CSTR_EQUAL)
            {
                multipleParents = true;
                break;
            }
        }

        if (multipleParents)
        {
            fromText = LoadStringResource(nullptr, IDS_FILEOPS_LOCATION_MULTIPLE);
        }
        else if (unknownCount == 0 && fileCount > 0 && folderCount > 0 && hasSampleFile)
        {
            fromText = sampleFile.wstring();
        }
        else
        {
            fromText = ensureTrailingSeparator(commonParent.wstring());
        }
    }

    std::wstring toText = ensureTrailingSeparator(destinationFolder.wstring());
    normalizeSlashes(fromText);
    normalizeSlashes(toText);

    const UINT messageId = operation == FILESYSTEM_COPY ? static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_COPY) : static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_MOVE);
    const std::wstring message = FormatStringResource(nullptr, messageId, what, fromText, toText);

    const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_CONFIRM);
    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = (owner && IsWindow(owner)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    prompt.severity      = HOST_ALERT_INFO;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = (prompt.scope == HOST_ALERT_SCOPE_WINDOW) ? owner : nullptr;
    prompt.title         = caption.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_OK;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hr              = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hr))
    {
        return false;
    }

    return promptResult == HOST_PROMPT_RESULT_OK;
}

bool IsOverlaySampleEnabled() noexcept
{
#if defined(_DEBUG) || defined(DEBUG)
    return true;
#else
    return false;
#endif
}

enum FolderCommands : UINT
{
    CmdOpen                             = IDM_FOLDERVIEW_CONTEXT_OPEN,
    CmdOpenWith                         = IDM_FOLDERVIEW_CONTEXT_OPEN_WITH,
    CmdViewSpace                        = IDM_FOLDERVIEW_CONTEXT_VIEW_SPACE,
    CmdDelete                           = IDM_FOLDERVIEW_CONTEXT_DELETE,
    CmdRename                           = IDM_FOLDERVIEW_CONTEXT_RENAME,
    CmdCopy                             = IDM_FOLDERVIEW_CONTEXT_COPY,
    CmdPaste                            = IDM_FOLDERVIEW_CONTEXT_PASTE,
    CmdSelectAll                        = IDM_FOLDERVIEW_CONTEXT_SELECT_ALL,
    CmdUnselectAll                      = IDM_FOLDERVIEW_CONTEXT_UNSELECT_ALL,
    CmdProperties                       = IDM_FOLDERVIEW_CONTEXT_PROPERTIES,
    CmdMove                             = IDM_FOLDERVIEW_CONTEXT_MOVE,
    CmdOverlaySampleError               = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_ERROR,
    CmdOverlaySampleWarning             = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_WARNING,
    CmdOverlaySampleInformation         = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_INFORMATION,
    CmdOverlaySampleBusy                = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_BUSY,
    CmdOverlaySampleHide                = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_HIDE,
    CmdOverlaySampleErrorNonModal       = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_ERROR_NONMODAL,
    CmdOverlaySampleWarningNonModal     = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_WARNING_NONMODAL,
    CmdOverlaySampleInformationNonModal = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_INFORMATION_NONMODAL,
    CmdOverlaySampleCanceled            = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_CANCELED,
    CmdOverlaySampleBusyWithCancel      = IDM_FOLDERVIEW_CONTEXT_OVERLAY_SAMPLE_BUSY_WITH_CANCEL,
};

std::wstring FormatHResult(HRESULT hr)
{
    return FormatHResultMessage(hr);
}

HRESULT HrFromErrorCode(const std::error_code& ec)
{
    if (! ec)
    {
        return S_OK;
    }
    if (&ec.category() == &std::system_category())
    {
        return HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value()));
    }
    return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
}

std::wstring FormatLocalTime(int64_t fileTime)
{
    if (fileTime <= 0)
    {
        return {};
    }

    ULARGE_INTEGER uli{};
    uli.QuadPart = static_cast<ULONGLONG>(fileTime);

    FILETIME ft{};
    ft.dwLowDateTime  = uli.LowPart;
    ft.dwHighDateTime = uli.HighPart;

    FILETIME local{};
    SYSTEMTIME st{};
    if (! FileTimeToLocalFileTime(&ft, &local) || ! FileTimeToSystemTime(&local, &st))
    {
        return {};
    }

    return std::format(L"{:04d}-{:02d}-{:02d} {:02d}:{:02d}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute);
}

std::wstring FormatFileAttributes(DWORD attrs)
{
    std::wstring result;
    result.reserve(10);

    auto add = [&](DWORD flag, wchar_t ch)
    {
        if ((attrs & flag) != 0)
        {
            result.push_back(ch);
        }
    };

    add(FILE_ATTRIBUTE_READONLY, L'R');
    add(FILE_ATTRIBUTE_HIDDEN, L'H');
    add(FILE_ATTRIBUTE_SYSTEM, L'S');
    add(FILE_ATTRIBUTE_ARCHIVE, L'A');
    add(FILE_ATTRIBUTE_COMPRESSED, L'C');
    add(FILE_ATTRIBUTE_ENCRYPTED, L'E');
    add(FILE_ATTRIBUTE_TEMPORARY, L'T');
    add(FILE_ATTRIBUTE_OFFLINE, L'O');
    add(FILE_ATTRIBUTE_REPARSE_POINT, L'P');

    if (result.empty())
    {
        result = L"-";
    }

    return result;
}

std::wstring FileTypeLabel(std::wstring_view extension, bool isDirectory)
{
    if (isDirectory)
    {
        return LoadStringResource(nullptr, IDS_FOLDERVIEW_TYPE_FOLDER);
    }

    std::wstring type(extension);
    if (! type.empty() && type.front() == L'.')
    {
        type.erase(type.begin());
    }
    if (type.empty())
    {
        return LoadStringResource(nullptr, IDS_FOLDERVIEW_TYPE_FILE);
    }

    for (auto& ch : type)
    {
        ch = static_cast<wchar_t>(towupper(ch));
    }

    return type;
}

std::wstring PadLeftToWidth(std::wstring_view text, size_t width)
{
    if (text.size() >= width)
    {
        return std::wstring(text);
    }

    std::wstring result;
    result.reserve(width);
    result.append(width - text.size(), L' ');
    result.append(text);
    return result;
}

std::wstring BuildDetailsText(bool isDirectory, uint64_t sizeBytes, int64_t lastWriteTime, DWORD fileAttributes, size_t sizeSlotChars)
{
    const std::wstring timeText  = FormatLocalTime(lastWriteTime);
    const std::wstring attrsText = FormatFileAttributes(fileAttributes);

    if (isDirectory)
    {
        return std::format(L"{} • {}", timeText, attrsText);
    }

    std::wstring sizeField;
    if (sizeSlotChars > 0)
    {
        const std::wstring sizeText = FormatBytesCompact(sizeBytes);
        sizeField                   = PadLeftToWidth(sizeText, sizeSlotChars);
    }
    else
    {
        sizeField = FormatBytesCompact(sizeBytes);
    }

    return std::format(L"{} • {} • {}", timeText, sizeField, attrsText);
}

[[nodiscard]] HWND NormalizeOwnerWindow(HWND owner) noexcept
{
    if (! owner || IsWindow(owner) == FALSE)
    {
        return nullptr;
    }

    if (const HWND rootOwner = GetAncestor(owner, GA_ROOT); rootOwner && IsWindow(rootOwner) != FALSE)
    {
        return rootOwner;
    }

    return owner;
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

[[nodiscard]] std::wstring TrimRenameText(std::wstring text) noexcept
{
    text.erase(text.begin(), std::find_if(text.begin(), text.end(), [](wchar_t ch) { return ! iswspace(ch); }));
    text.erase(std::find_if(text.rbegin(), text.rend(), [](wchar_t ch) { return ! iswspace(ch); }).base(), text.end());
    return text;
}

constexpr wchar_t kFolderViewRenamePromptClassName[] = L"RedSalamander.FolderView.RenamePrompt";

#ifdef ENABLE_TESTS
enum class FolderViewRenamePromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetText,
    Confirm,
    Cancel,
};
#endif

class FolderViewRenamePromptWindow final
{
public:
    FolderViewRenamePromptWindow(const FolderViewRenamePromptWindow&)            = delete;
    FolderViewRenamePromptWindow& operator=(const FolderViewRenamePromptWindow&) = delete;
    FolderViewRenamePromptWindow(FolderViewRenamePromptWindow&&)                 = delete;
    FolderViewRenamePromptWindow& operator=(FolderViewRenamePromptWindow&&)      = delete;

    FolderViewRenamePromptWindow(HWND ownerWindow, std::wstring initialText, bool isDirectory, const AppTheme& theme) noexcept
        : _ownerWindow(NormalizeOwnerWindow(ownerWindow)),
          _restoreFocusWindow(ownerWindow && IsWindow(ownerWindow) != FALSE ? ownerWindow : nullptr),
          _initialText(std::move(initialText)),
          _isDirectory(isDirectory),
          _captionText(LoadStringResource(nullptr, IDS_FOLDERVIEW_RENAME_CAPTION)),
          _theme(theme)
    {
        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            const HWND focused = GetFocus();
            if (focused && IsWindow(focused) != FALSE && (focused == _ownerWindow || IsChild(_ownerWindow, focused) != FALSE))
            {
                _restoreFocusWindow = focused;
            }
            else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE ||
                     (_restoreFocusWindow != _ownerWindow && IsChild(_ownerWindow, _restoreFocusWindow) == FALSE))
            {
                _restoreFocusWindow = _ownerWindow;
            }
        }
        else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE)
        {
            _restoreFocusWindow = nullptr;
        }
    }

    [[nodiscard]] std::optional<std::wstring> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style        = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle      = WS_EX_DLGMODALFRAME;
        const UINT dpi           = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();
        const int clientWidthPx  = ScaleForDpi(dpi, 420);
        const int clientHeightPx = ScaleForDpi(dpi, 164);

        RECT bounds{0, 0, clientWidthPx, clientHeightPx};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kFolderViewRenamePromptClassName,
                                          _captionText.c_str(),
                                          style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        CenterWindowOnOwner(_hWnd.get(), _ownerWindow);
        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return std::nullopt;
            }
            if (getMessageResult == 0)
            {
                _done = true;
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _acceptedText;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<FolderViewRenamePromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<FolderViewRenamePromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
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
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
#ifdef ENABLE_TESTS
            case WndMsg::kFolderViewRenamePromptDebug: return self->OnDebugCommand(static_cast<FolderViewRenamePromptDebugCommand>(wParam), lParam);
#endif
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
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, self->_theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
                if (self->_ownerWindow && IsWindow(self->_ownerWindow) != FALSE)
                {
                    static_cast<void>(SetActiveWindow(self->_ownerWindow));

                    const HWND restoreFocus =
                        (self->_restoreFocusWindow && IsWindow(self->_restoreFocusWindow) != FALSE &&
                         (self->_restoreFocusWindow == self->_ownerWindow || IsChild(self->_ownerWindow, self->_restoreFocusWindow) != FALSE))
                            ? self->_restoreFocusWindow
                            : self->_ownerWindow;
                    static_cast<void>(SetFocus(restoreFocus));
                }
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
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
        wc.lpfnWndProc   = FolderViewRenamePromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kFolderViewRenamePromptClassName;
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
        if (! _captionText.empty())
        {
            SetWindowTextW(hwnd, _captionText.c_str());
        }
        ApplyTheme();
        Layout();
        if (_field)
        {
            _field->SetSelectionRange(0u, ResolveInitialSelectionEnd());
        }
        _dxHost.SetFocusControl(_field);
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi() noexcept
    {
        if (_root != nullptr)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _label = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_FOLDERVIEW_RENAME_LABEL));
        _label->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _field = _root->AddChild<TextField>(_initialText);
        _field->SetMultiline(false);
        _field->SetAccessibleName(LoadStringResource(nullptr, IDS_FOLDERVIEW_RENAME_LABEL));

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_hWnd)
        {
            if (! _captionText.empty())
            {
                SetWindowTextW(_hWnd.get(), _captionText.c_str());
            }
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
            static_cast<void>(RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW));
        }
    }

    [[nodiscard]] size_t ResolveInitialSelectionEnd() const noexcept
    {
        if (_isDirectory)
        {
            return _initialText.size();
        }

        const size_t dot = _initialText.find_last_of(L'.');
        if (dot == std::wstring::npos || dot == 0u)
        {
            return _initialText.size();
        }

        return dot;
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
        std::wstring text = _field ? std::wstring(_field->GetText()) : std::wstring{};
        text              = TrimRenameText(std::move(text));
        if (text.empty())
        {
            MessageBeep(MB_ICONWARNING);
            if (_field)
            {
                _dxHost.SetFocusControl(_field);
            }
            return;
        }

        _acceptedText = std::move(text);
        _done         = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

    void Cancel() noexcept
    {
        _acceptedText.reset();
        _done = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(FolderViewRenamePromptDebugCommand command, LPARAM lParam) noexcept
    {
        switch (command)
        {
            case FolderViewRenamePromptDebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<FolderViewRenamePromptDebugSnapshot*>(lParam);
                if (! snapshot)
                {
                    return FALSE;
                }

                snapshot->usesDxUiHost            = _dxHost.GetRoot() != nullptr;
                snapshot->visibleChildWindowCount = 0u;
                if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
                {
                    EnumChildWindows(_hWnd.get(),
                                     [](HWND child, LPARAM cookie) noexcept -> BOOL
                    {
                        if (IsWindowVisible(child) == FALSE)
                        {
                            return TRUE;
                        }
                        auto* count = reinterpret_cast<size_t*>(cookie);
                        if (count)
                        {
                            *count += 1u;
                        }
                        return TRUE;
                    },
                                     reinterpret_cast<LPARAM>(&snapshot->visibleChildWindowCount));
                }
                snapshot->text = _field ? std::wstring(_field->GetText()) : std::wstring{};
                if (_field)
                {
                    if (const auto selection = _field->GetSelectionRange(); selection.has_value())
                    {
                        snapshot->selectionStart = selection->first;
                        snapshot->selectionEnd   = selection->second;
                    }
                    else
                    {
                        snapshot->selectionStart = snapshot->text.size();
                        snapshot->selectionEnd   = snapshot->text.size();
                    }

                    RedSalamander::DxUi::TextFieldDebugSingleLinePaintState paintState{};
                    if (_field->DebugGetSingleLinePaintState(_dxHost, paintState))
                    {
                        snapshot->textRect              = paintState.textRect;
                        snapshot->selectionPaintRect    = paintState.selectionPaintRect;
                        snapshot->horizontalScrollDip   = paintState.horizontalScrollDip;
                        snapshot->hasSelectionPaintRect = paintState.hasSelectionPaintRect;
                    }
                }
                return TRUE;
            }
            case FolderViewRenamePromptDebugCommand::SetText:
            {
                const auto* text = reinterpret_cast<const std::wstring*>(lParam);
                if (! text || ! _field)
                {
                    return FALSE;
                }

                _field->SetTextAndNotify(*text);
                _dxHost.SetFocusControl(_field);
                return TRUE;
            }
            case FolderViewRenamePromptDebugCommand::Confirm: Confirm(); return TRUE;
            case FolderViewRenamePromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

    HWND _ownerWindow        = nullptr;
    HWND _restoreFocusWindow = nullptr;
    std::wstring _initialText;
    bool _isDirectory = false;
    std::wstring _captionText;
    AppTheme _theme{};
    RedSalamander::DxUi::ThemePalette _palette{};
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root          = nullptr;
    RedSalamander::DxUi::Label* _label         = nullptr;
    RedSalamander::DxUi::TextField* _field     = nullptr;
    RedSalamander::DxUi::Button* _okButton     = nullptr;
    RedSalamander::DxUi::Button* _cancelButton = nullptr;
    bool _done                                 = false;
    std::optional<std::wstring> _acceptedText;
};

std::optional<std::wstring> PromptForRename(HWND owner, const std::wstring& currentName, bool isDirectory, const AppTheme& theme)
{
    FolderViewRenamePromptWindow prompt(owner, currentName, isDirectory, theme);
    return prompt.ShowModal();
}

void AppendMultiSz(std::wstring& buffer, const std::wstring& path)
{
    buffer.append(path);
    buffer.push_back(L'\0');
}

std::wstring BuildMultiSz(const std::vector<std::filesystem::path>& paths)
{
    std::wstring buffer;
    for (const auto& p : paths)
    {
        AppendMultiSz(buffer, p.c_str());
    }
    buffer.push_back(L'\0');
    return buffer;
}

HRESULT
BuildPathArrayArena(const std::vector<std::filesystem::path>& paths, FileSystemArenaOwner& arenaOwner, const wchar_t*** outPaths, unsigned long* outCount)
{
    if (! outPaths || ! outCount)
    {
        return E_POINTER;
    }

    *outPaths = nullptr;
    *outCount = 0;

    if (paths.empty())
    {
        return S_OK;
    }

    const uint64_t count64 = static_cast<uint64_t>(paths.size());
    if (count64 > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const uint64_t arrayBytes64 = count64 * static_cast<uint64_t>(sizeof(const wchar_t*));
    if (arrayBytes64 > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    unsigned long totalBytes = static_cast<unsigned long>(arrayBytes64);

    for (const auto& path : paths)
    {
        const std::wstring& text = path.native();
        const size_t length      = text.size();
        if (length > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)) - 1u)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        const unsigned long bytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
        if (totalBytes > std::numeric_limits<unsigned long>::max() - bytes)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        totalBytes += bytes;
    }

    HRESULT hr = arenaOwner.Initialize(totalBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    FileSystemArena* arena = arenaOwner.Get();
    auto* array            = static_cast<const wchar_t**>(
        AllocateFromFileSystemArena(arena, static_cast<unsigned long>(arrayBytes64), static_cast<unsigned long>(alignof(const wchar_t*))));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (size_t index = 0; index < paths.size(); ++index)
    {
        const std::wstring& text  = paths[index].native();
        const size_t length       = text.size();
        const unsigned long bytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
        auto* buffer              = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, bytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! buffer)
        {
            return E_OUTOFMEMORY;
        }

        if (length > 0)
        {
            ::CopyMemory(buffer, text.data(), length * sizeof(wchar_t));
        }
        buffer[length] = L'\0';
        array[index]   = buffer;
    }

    *outPaths = array;
    *outCount = static_cast<unsigned long>(count64);
    return S_OK;
}

std::filesystem::path GenerateShortcutPath(const std::filesystem::path& folder, const std::filesystem::path& target, int attempt)
{
    std::wstring stem = target.stem().wstring();
    if (stem.empty())
    {
        stem = target.filename().wstring();
    }
    std::wstring suffix;
    if (attempt > 0)
    {
        suffix = std::format(L" ({})", attempt + 1);
    }
    std::wstring candidate = std::format(L"{} - Shortcut{}.lnk", stem, suffix);
    return folder / candidate;
}

void ConfigureLabelLayout(IDWriteTextLayout* layout, IDWriteInlineObject* ellipsisSign, bool enableEllipsisTrimming = true)
{
    if (! layout)
    {
        return;
    }
    layout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
    layout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
    layout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
    DWRITE_TRIMMING trimming{};
    trimming.granularity = enableEllipsisTrimming ? DWRITE_TRIMMING_GRANULARITY_CHARACTER : DWRITE_TRIMMING_GRANULARITY_NONE;
    layout->SetTrimming(&trimming, enableEllipsisTrimming ? ellipsisSign : nullptr);
}

UINT PreferredDropEffectFormat()
{
    static const UINT format = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
    return format;
}

UINT RedSalamanderInternalFileDropFormat()
{
    static const UINT format = RegisterClipboardFormatW(L"RedSalamander.InternalFileDrop.V1");
    return format;
}

class FormatEnumerator final : public IEnumFORMATETC
{
public:
    explicit FormatEnumerator(const std::vector<FORMATETC>& formats) : _refCount(1), _formats(formats)
    {
    }

    FormatEnumerator(const FormatEnumerator& other) : _refCount(1), _formats(other._formats), _index(other._index)
    {
    }

    // Explicitly delete assignment operators (COM objects are not assignable)
    FormatEnumerator& operator=(const FormatEnumerator&) = delete;
    FormatEnumerator(FormatEnumerator&&)                 = delete;
    FormatEnumerator& operator=(FormatEnumerator&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IEnumFORMATETC)
        {
            *ppvObject = static_cast<IEnumFORMATETC*>(this);
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
        ULONG remaining = --_refCount;
        if (remaining == 0)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE Next(ULONG celt, FORMATETC* rgelt, ULONG* pceltFetched) override
    {
        if (! rgelt)
        {
            return E_POINTER;
        }
        ULONG fetched = 0;
        while (fetched < celt && _index < _formats.size())
        {
            rgelt[fetched]     = _formats[_index];
            rgelt[fetched].ptd = nullptr;
            ++_index;
            ++fetched;
        }
        if (pceltFetched)
        {
            *pceltFetched = fetched;
        }
        return fetched == celt ? S_OK : S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE Skip(ULONG celt) override
    {
        const size_t remaining = _formats.size() - std::min(_index, _formats.size());
        if (celt > remaining)
        {
            _index = _formats.size();
            return S_FALSE;
        }
        _index += celt;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Reset() override
    {
        _index = 0;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Clone(IEnumFORMATETC** ppenum) override
    {
        if (! ppenum)
        {
            return E_POINTER;
        }
        auto* clone = new (std::nothrow) FormatEnumerator(*this);
        if (! clone)
        {
            return E_OUTOFMEMORY;
        }
        clone->_index = _index;
        *ppenum       = clone;
        return S_OK;
    }

private:
    std::atomic<ULONG> _refCount;
    std::vector<FORMATETC> _formats;
    size_t _index = 0;
};

class FolderViewDataObject final : public IDataObject
{
public:
    FolderViewDataObject(
        std::vector<std::filesystem::path> paths, std::wstring pluginId, std::wstring instanceContext, DWORD preferredEffect, bool includeHDrop)
        : _refCount(1),
          _paths(std::move(paths)),
          _pluginId(std::move(pluginId)),
          _instanceContext(std::move(instanceContext)),
          _preferredEffect(preferredEffect),
          _includeHDrop(includeHDrop)
    {
    }

    // Explicitly delete copy/move operations (COM objects are not copyable/movable)
    FolderViewDataObject(const FolderViewDataObject&)            = delete;
    FolderViewDataObject(FolderViewDataObject&&)                 = delete;
    FolderViewDataObject& operator=(const FolderViewDataObject&) = delete;
    FolderViewDataObject& operator=(FolderViewDataObject&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IDataObject)
        {
            *ppvObject = static_cast<IDataObject*>(this);
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
        ULONG remaining = --_refCount;
        if (remaining == 0)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE GetData(FORMATETC* format, STGMEDIUM* medium) override
    {
        if (! format || ! medium)
        {
            return E_POINTER;
        }

        if ((format->tymed & TYMED_HGLOBAL) == 0)
        {
            return DV_E_TYMED;
        }

        if (format->cfFormat == static_cast<CLIPFORMAT>(RedSalamanderInternalFileDropFormat()))
        {
            auto data = CreateInternalFileDrop();
            if (! data)
            {
                return E_OUTOFMEMORY;
            }
            medium->tymed          = TYMED_HGLOBAL;
            medium->hGlobal        = data.release();
            medium->pUnkForRelease = nullptr;
            return S_OK;
        }

        if (format->cfFormat == CF_HDROP)
        {
            if (! _includeHDrop)
            {
                return DV_E_FORMATETC;
            }

            auto data = CreateHDrop();
            if (! data)
            {
                return E_OUTOFMEMORY;
            }
            medium->tymed          = TYMED_HGLOBAL;
            medium->hGlobal        = data.release();
            medium->pUnkForRelease = nullptr;
            return S_OK;
        }

        if (format->cfFormat == PreferredDropEffectFormat())
        {
            auto data = CreatePreferredEffect();
            if (! data)
            {
                return E_OUTOFMEMORY;
            }
            medium->tymed          = TYMED_HGLOBAL;
            medium->hGlobal        = data.release();
            medium->pUnkForRelease = nullptr;
            return S_OK;
        }

        return DV_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC*, STGMEDIUM*) override
    {
        return DATA_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* format) override
    {
        if (! format)
        {
            return E_POINTER;
        }
        if ((format->tymed & TYMED_HGLOBAL) == 0)
        {
            return DV_E_TYMED;
        }
        if (format->cfFormat == static_cast<CLIPFORMAT>(RedSalamanderInternalFileDropFormat()) || format->cfFormat == PreferredDropEffectFormat())
        {
            return S_OK;
        }
        if (format->cfFormat == CF_HDROP)
        {
            return _includeHDrop ? S_OK : DV_E_FORMATETC;
        }
        return DV_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC*, FORMATETC* result) override
    {
        if (! result)
        {
            return E_POINTER;
        }
        *result = {};
        return DATA_S_SAMEFORMATETC;
    }

    HRESULT STDMETHODCALLTYPE SetData(FORMATETC*, STGMEDIUM*, BOOL) override
    {
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD direction, IEnumFORMATETC** enumerator) override
    {
        if (! enumerator)
        {
            return E_POINTER;
        }
        *enumerator = nullptr;
        if (direction != DATADIR_GET)
        {
            return E_NOTIMPL;
        }

        std::vector<FORMATETC> formats;
        FORMATETC hdrop{};
        hdrop.dwAspect = DVASPECT_CONTENT;
        hdrop.lindex   = -1;
        hdrop.ptd      = nullptr;
        hdrop.tymed    = TYMED_HGLOBAL;

        hdrop.cfFormat = static_cast<CLIPFORMAT>(RedSalamanderInternalFileDropFormat());
        formats.push_back(hdrop);

        if (_includeHDrop)
        {
            hdrop.cfFormat = CF_HDROP;
            formats.push_back(hdrop);
        }

        hdrop.cfFormat = static_cast<CLIPFORMAT>(PreferredDropEffectFormat());
        formats.push_back(hdrop);

        auto* enumFormats = new (std::nothrow) FormatEnumerator(formats);
        if (! enumFormats)
        {
            return E_OUTOFMEMORY;
        }
        *enumerator = enumFormats;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

    HRESULT STDMETHODCALLTYPE DUnadvise(DWORD) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

    HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA**) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

private:
    wil::unique_hglobal CreateInternalFileDrop() const
    {
        struct Header
        {
            uint32_t version              = 1;
            uint32_t pluginIdChars        = 0;
            uint32_t instanceContextChars = 0;
            uint32_t pathCount            = 0;
        };

        const size_t pluginIdChars = _pluginId.size();
        const size_t instanceChars = _instanceContext.size();
        const size_t pathCount     = _paths.size();
        if (pluginIdChars > std::numeric_limits<uint32_t>::max() || instanceChars > std::numeric_limits<uint32_t>::max() ||
            pathCount > std::numeric_limits<uint32_t>::max())
        {
            return nullptr;
        }

        size_t totalBytes = sizeof(Header);
        auto addString    = [&](size_t chars) -> bool
        {
            const size_t add = (chars + 1u) * sizeof(wchar_t);
            if (totalBytes > (std::numeric_limits<size_t>::max)() - add)
            {
                return false;
            }
            totalBytes += add;
            return true;
        };

        if (! addString(pluginIdChars) || ! addString(instanceChars))
        {
            return nullptr;
        }

        for (const auto& path : _paths)
        {
            const std::wstring& text = path.native();
            const size_t chars       = text.size();
            if (chars > std::numeric_limits<uint32_t>::max())
            {
                return nullptr;
            }

            if (totalBytes > (std::numeric_limits<size_t>::max)() - sizeof(uint32_t))
            {
                return nullptr;
            }
            totalBytes += sizeof(uint32_t);

            if (! addString(chars))
            {
                return nullptr;
            }
        }

        wil::unique_hglobal data(GlobalAlloc(GHND, totalBytes));
        if (! data)
        {
            return nullptr;
        }

        void* raw = GlobalLock(data.get());
        if (! raw)
        {
            return nullptr;
        }
        // Copy the handle value now; don’t reference the local `data`.
        HGLOBAL h   = data.get();
        auto unlock = wil::scope_exit([h]()
        {
            if (h)
                GlobalUnlock(h);
        });

        auto* header                 = static_cast<Header*>(raw);
        header->version              = 1;
        header->pluginIdChars        = static_cast<uint32_t>(pluginIdChars);
        header->instanceContextChars = static_cast<uint32_t>(instanceChars);
        header->pathCount            = static_cast<uint32_t>(pathCount);

        std::byte* cursor = reinterpret_cast<std::byte*>(header + 1u);

        const auto writeString = [&](std::wstring_view text)
        {
            const size_t bytes = (text.size() + 1u) * sizeof(wchar_t);
            memcpy(cursor, text.data(), text.size() * sizeof(wchar_t));
            cursor += text.size() * sizeof(wchar_t);
            *reinterpret_cast<wchar_t*>(cursor) = L'\0';
            cursor += sizeof(wchar_t);
            static_cast<void>(bytes);
        };

        writeString(_pluginId);
        writeString(_instanceContext);

        for (const auto& path : _paths)
        {
            const std::wstring& text = path.native();
            const uint32_t chars32   = static_cast<uint32_t>(text.size());
            memcpy(cursor, &chars32, sizeof(chars32));
            cursor += sizeof(chars32);
            writeString(text);
        }

        return data;
    }

    wil::unique_hglobal CreateHDrop() const
    {
        size_t totalChars = 0;
        for (const auto& path : _paths)
        {
            totalChars += path.native().size() + 1;
        }
        totalChars += 1; // double-null terminator

        const size_t bytes = sizeof(DROPFILES) + totalChars * sizeof(wchar_t);
        wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, bytes));
        if (! memory)
        {
            return nullptr;
        }

        auto* dropFiles = static_cast<DROPFILES*>(GlobalLock(memory.get()));
        if (! dropFiles)
        {
            return nullptr;
        }

        dropFiles->pFiles = sizeof(DROPFILES);
        dropFiles->pt     = POINT{};
        dropFiles->fNC    = FALSE;
        dropFiles->fWide  = TRUE;

        auto* buffer = reinterpret_cast<wchar_t*>(reinterpret_cast<BYTE*>(dropFiles) + dropFiles->pFiles);
        for (const auto& path : _paths)
        {
            const std::wstring wide = path.native();
            std::copy(wide.begin(), wide.end(), buffer);
            buffer += wide.size();
            *buffer++ = L'\0';
        }
        *buffer = L'\0';
        GlobalUnlock(memory.get());
        return memory;
    }

    wil::unique_hglobal CreatePreferredEffect() const
    {
        wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD)));
        if (! memory)
        {
            return nullptr;
        }
        auto* effect = static_cast<DWORD*>(GlobalLock(memory.get()));
        if (! effect)
        {
            return nullptr;
        }
        *effect = _preferredEffect;
        GlobalUnlock(memory.get());
        return memory;
    }

    std::atomic<ULONG> _refCount;
    std::vector<std::filesystem::path> _paths;
    std::wstring _pluginId;
    std::wstring _instanceContext;
    DWORD _preferredEffect = DROPEFFECT_COPY;
    bool _includeHDrop     = false;
};

class FolderViewDropSource final : public IDropSource
{
public:
    FolderViewDropSource() : _refCount(1)
    {
    }

    // Explicitly delete copy/move operations (COM objects are not copyable/movable)
    FolderViewDropSource(const FolderViewDropSource&)            = delete;
    FolderViewDropSource(FolderViewDropSource&&)                 = delete;
    FolderViewDropSource& operator=(const FolderViewDropSource&) = delete;
    FolderViewDropSource& operator=(FolderViewDropSource&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IDropSource)
        {
            *ppvObject = static_cast<IDropSource*>(this);
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
        ULONG remaining = --_refCount;
        if (remaining == 0)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL escapePressed, DWORD keyState) override
    {
        if (escapePressed)
        {
            return DRAGDROP_S_CANCEL;
        }
        if ((keyState & MK_LBUTTON) == 0)
        {
            return DRAGDROP_S_DROP;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD) override
    {
        return DRAGDROP_S_USEDEFAULTCURSORS;
    }

private:
    std::atomic<ULONG> _refCount;
};

} // namespace

#pragma warning(pop)
