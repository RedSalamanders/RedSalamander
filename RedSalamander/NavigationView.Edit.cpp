#include "NavigationViewInternal.h"

#include "DxUi/DxUi.Typography.h"
#include <windowsx.h>

#include "ConnectionSecrets.h"
#include "DirectoryInfoCache.h"
#include "FileSystemPluginManager.h"
#include "FluentIcons.h"
#include "Helpers.h"
#include "HostServices.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "SettingsStore.h"
#include "UiMetrics.h"
#include "resource.h"

namespace
{
constexpr wchar_t kNavigationEditOriginalWndProcProp[] = L"RS.NavigationView.Edit.OriginalWndProc";
constexpr wchar_t kNavigationEditOwnerProp[]           = L"RS.NavigationView.Edit.Owner";
constexpr UINT kDxUiModifierAlt                        = 0x0100u;
constexpr int kValidationPopupPreferredWidthDip        = 420;
constexpr int kValidationPopupMinWidthDip              = 280;
constexpr int kValidationPopupMinHeightDip             = 36;
constexpr int kValidationPopupMarginDip                = 8;
constexpr int kValidationPopupGapDip                   = 2;
constexpr int kValidationPopupPaddingXDip              = 10;
constexpr int kValidationPopupPaddingYDip              = 7;
constexpr int kValidationPopupIconSizeDip              = 16;
constexpr int kValidationPopupIconGapDip               = 8;
constexpr int kValidationPopupCornerRadiusDip          = 6;

[[nodiscard]] WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* originalWndProcProp) noexcept
{
    return RedSalamander::Win32Callback::GetStoredWndProc(hwnd, originalWndProcProp);
}

[[nodiscard]] bool InstallWndProcHook(HWND hwnd, WNDPROC wndProc, const wchar_t* originalWndProcProp) noexcept
{
    if (! hwnd || ! wndProc)
    {
        return false;
    }

    if (GetStoredWndProc(hwnd, originalWndProcProp))
    {
        return true;
    }

    const auto previous =
        reinterpret_cast<WNDPROC>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(wndProc)));
    if (! previous)
    {
        return false;
    }

    if (! RedSalamander::Win32Callback::SetPropNoThrow(hwnd, originalWndProcProp, reinterpret_cast<HANDLE>(previous)))
    {
        static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(previous)));
        return false;
    }

    return true;
}

void RestoreWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp, const wchar_t* ownerProp) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const auto originalWndProc = GetStoredWndProc(hwnd, originalWndProcProp);
    if (originalWndProc && IsWindow(hwnd))
    {
        static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(originalWndProc)));
    }

    RemovePropW(hwnd, originalWndProcProp);
    RemovePropW(hwnd, ownerProp);
}

[[nodiscard]] LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* originalWndProcProp, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    const auto originalWndProc = GetStoredWndProc(hwnd, originalWndProcProp);
    return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
}

[[nodiscard]] UINT32 ClampTextLengthForDWrite(std::wstring_view text) noexcept
{
    return static_cast<UINT32>(std::min(text.size(), static_cast<size_t>(std::numeric_limits<UINT32>::max())));
}

[[nodiscard]] wil::com_ptr<IDWriteTextFormat> CreateValidationPopupTextFormat(IDWriteFactory* dwriteFactory, UINT dpi) noexcept
{
    if (! dwriteFactory)
    {
        return {};
    }

    const float textSizeDip = static_cast<float>(MulDiv(13, static_cast<int>(std::max<UINT>(USER_DEFAULT_SCREEN_DPI, dpi)), USER_DEFAULT_SCREEN_DPI));
    wil::com_ptr<IDWriteTextFormat> format;
    const HRESULT hr =
        RedSalamander::DxUi::Typography::CreateTextFormat(dwriteFactory, RedSalamander::DxUi::Typography::MakeUiTextSpec(textSizeDip), format.put(), L"");
    if (FAILED(hr) || ! format)
    {
        return {};
    }

    static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
    static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
    static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP));
    return format;
}

[[nodiscard]] wil::com_ptr<IDWriteTextFormat> CreateValidationPopupIconTextFormat(IDWriteFactory* dwriteFactory,
                                                                                  UINT dpi,
                                                                                  bool& usesFluentIcon,
                                                                                  wchar_t& iconGlyph) noexcept
{
    usesFluentIcon = false;
    iconGlyph      = FluentIcons::kFallbackWarning;
    if (! dwriteFactory)
    {
        return {};
    }

    const float iconSizeDip =
        static_cast<float>(MulDiv(kValidationPopupIconSizeDip, static_cast<int>(std::max<UINT>(USER_DEFAULT_SCREEN_DPI, dpi)), USER_DEFAULT_SCREEN_DPI));
    usesFluentIcon = RedSalamander::DxUi::Typography::IsFontFamilyAvailable(dwriteFactory, RedSalamander::DxUi::Typography::kSegoeFluentIconsFamily);
    iconGlyph      = usesFluentIcon ? FluentIcons::kWarning : FluentIcons::kFallbackWarning;

    wil::com_ptr<IDWriteTextFormat> format;
    const HRESULT hr = RedSalamander::DxUi::Typography::CreateTextFormat(
        dwriteFactory,
        usesFluentIcon
            ? RedSalamander::DxUi::Typography::MakeUiIconSpec(iconSizeDip)
            : RedSalamander::DxUi::Typography::TypographySpec{.familyName = L"Segoe UI Symbol", .weight = DWRITE_FONT_WEIGHT_NORMAL, .sizeDip = iconSizeDip},
        format.put(),
        L"");
    if (FAILED(hr) || ! format)
    {
        usesFluentIcon = false;
        iconGlyph      = L'\0';
        return {};
    }

    static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
    static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    return format;
}

[[nodiscard]] bool ApplyValidationPopupRoundedRegion(HWND hwnd, int widthPx, int heightPx, UINT dpi) noexcept
{
    if (! hwnd || widthPx <= 0 || heightPx <= 0)
    {
        return false;
    }

    const int radiusPx   = std::max(1, DipsToPixelsInt(kValidationPopupCornerRadiusDip, dpi));
    const int diameterPx = std::max(1, radiusPx * 2);
    wil::unique_hrgn region(CreateRoundRectRgn(0, 0, widthPx + 1, heightPx + 1, diameterPx, diameterPx));
    if (! region)
    {
        return false;
    }

    if (SetWindowRgn(hwnd, region.get(), TRUE) == 0)
    {
        return false;
    }

    region.release();
    return true;
}

[[nodiscard]] bool IsForwardableEditHostMouseActivateMessage(UINT msg) noexcept
{
    switch (msg)
    {
        case WM_MOUSEMOVE:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_RBUTTONDBLCLK:
        case WM_MBUTTONDOWN:
        case WM_MBUTTONUP:
        case WM_MBUTTONDBLCLK:
        case WM_XBUTTONDOWN:
        case WM_XBUTTONUP:
        case WM_XBUTTONDBLCLK: return true;
        default: return false;
    }
}

[[nodiscard]] WPARAM BuildForwardedMouseKeyState(UINT msg) noexcept
{
    WPARAM state = 0;
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        state |= MK_CONTROL;
    }
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        state |= MK_SHIFT;
    }
    if ((GetKeyState(VK_LBUTTON) & 0x8000) != 0 || msg == WM_LBUTTONDOWN || msg == WM_LBUTTONDBLCLK)
    {
        state |= MK_LBUTTON;
    }
    if ((GetKeyState(VK_RBUTTON) & 0x8000) != 0 || msg == WM_RBUTTONDOWN || msg == WM_RBUTTONDBLCLK)
    {
        state |= MK_RBUTTON;
    }
    if ((GetKeyState(VK_MBUTTON) & 0x8000) != 0 || msg == WM_MBUTTONDOWN || msg == WM_MBUTTONDBLCLK)
    {
        state |= MK_MBUTTON;
    }
    if ((GetKeyState(VK_XBUTTON1) & 0x8000) != 0)
    {
        state |= MK_XBUTTON1;
    }
    if ((GetKeyState(VK_XBUTTON2) & 0x8000) != 0)
    {
        state |= MK_XBUTTON2;
    }
    return state;
}
} // namespace

ATOM NavigationView::RegisterEditSuggestPopupWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = EditSuggestPopupWndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kSuggestPopupClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

ATOM NavigationView::RegisterEditValidationPopupWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = EditValidationPopupWndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kValidationPopupClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

LRESULT CALLBACK NavigationView::EditSuggestPopupWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    NavigationView* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<NavigationView*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    }
    else
    {
        self = reinterpret_cast<NavigationView*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self)
    {
        return self->EditSuggestPopupWndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK NavigationView::EditValidationPopupWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    NavigationView* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<NavigationView*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    }
    else
    {
        self = reinterpret_cast<NavigationView*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self)
    {
        return self->EditValidationPopupWndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT NavigationView::EditValidationPopupWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_MOUSEACTIVATE: return MA_NOACTIVATE;
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: PaintEditValidationPopup(hwnd); return 0;
        case WM_NCDESTROY:
            if (_editValidationPopup.get() == hwnd)
            {
                static_cast<void>(_editValidationPopup.release());
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            return 0;
        default: break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT NavigationView::OnCtlColorEdit(HDC hdc, HWND hwndControl)
{
    if (_pathEdit && hwndControl == _pathEdit->GetTextInputHwnd())
    {
        SetTextColor(hdc, ColorToCOLORREF(_theme.text));
        SetBkColor(hdc, _theme.gdiBackground);
        return reinterpret_cast<LRESULT>(_backgroundBrush.get());
    }

    if (! _hWnd)
    {
        return 0;
    }

    return DefWindowProcW(_hWnd.get(), WM_CTLCOLOREDIT, reinterpret_cast<WPARAM>(hdc), reinterpret_cast<LPARAM>(hwndControl));
}

LRESULT NavigationView::OnEditSuggestResults(std::unique_ptr<EditSuggestResultsPayload> owned)
{
    if (! owned)
    {
        return 0;
    }
    if (! _editMode || ! _pathEdit || ! _pathEdit->field || owned->requestId != _editSuggestRequestId.load(std::memory_order_acquire) ||
        owned->editSessionId != _editSuggestEditSessionId || owned->queryText != _pathEdit->field->GetText())
    {
        return 0;
    }

    const size_t count = std::min(owned->displayItems.size(), owned->insertItems.size());

    std::vector<EditSuggestItem> merged;
    merged.reserve(kEditSuggestMaxItems);

    const size_t maxWithoutEllipsis = (owned->hasMore && kEditSuggestMaxItems > 0u) ? (kEditSuggestMaxItems - 1u) : kEditSuggestMaxItems;

    if (_editSuggestAdditionalRequestId == owned->requestId && ! _editSuggestAdditionalItems.empty())
    {
        for (auto& item : _editSuggestAdditionalItems)
        {
            if (merged.size() >= maxWithoutEllipsis)
            {
                break;
            }
            merged.push_back(std::move(item));
        }
        _editSuggestAdditionalItems.clear();
        _editSuggestAdditionalRequestId = 0;
    }

    for (size_t i = 0; i < count && merged.size() < maxWithoutEllipsis; ++i)
    {
        EditSuggestItem item{};
        item.display            = std::move(owned->displayItems[i]);
        item.insertText         = std::move(owned->insertItems[i]);
        item.directorySeparator = owned->directorySeparator;
        merged.push_back(std::move(item));
    }

    if (owned->hasMore && merged.size() < kEditSuggestMaxItems)
    {
        EditSuggestItem item{};
        item.display            = std::wstring(kEllipsisText);
        item.enabled            = false;
        item.directorySeparator = L'\0';
        merged.push_back(std::move(item));
    }

    _editSuggestHoveredIndex  = -1;
    _editSuggestSelectedIndex = -1;
    _editSuggestHighlightText = std::move(owned->highlightText);
    _editSuggestItems         = std::move(merged);

    if (_editSuggestItems.empty())
    {
        CloseEditSuggestPopup();
    }
    else
    {
        UpdateEditSuggestPopupWindow();
    }

    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupCreate()
{
    _editSuggestHoveredIndex  = -1;
    _editSuggestSelectedIndex = -1;
    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupNcDestroy()
{
    DiscardEditSuggestPopupD2DResources();
    _editSuggestPopup.release();
    _editSuggestPopupClientSize  = {0, 0};
    _editSuggestPopupRowHeightPx = 0;
    _editSuggestItems.clear();
    _editSuggestHoveredIndex  = -1;
    _editSuggestSelectedIndex = -1;
    _editSuggestHighlightText.clear();
    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupSize(HWND hwnd, UINT width, UINT height)
{
    _editSuggestPopupClientSize.cx = static_cast<LONG>(width);
    _editSuggestPopupClientSize.cy = static_cast<LONG>(height);

    if (_editSuggestPopupTarget)
    {
        _editSuggestPopupTarget->Resize(D2D1::SizeU(static_cast<UINT32>(_editSuggestPopupClientSize.cx), static_cast<UINT32>(_editSuggestPopupClientSize.cy)));
    }

    InvalidateRect(hwnd, nullptr, FALSE);
    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupMouseMove(HWND hwnd, POINT pt)
{
    TRACKMOUSEEVENT tme{};
    tme.cbSize    = sizeof(tme);
    tme.dwFlags   = TME_LEAVE;
    tme.hwndTrack = hwnd;
    TrackMouseEvent(&tme);

    const int itemHeight =
        std::max(1, _editSuggestPopupRowHeightPx > 0 ? _editSuggestPopupRowHeightPx : static_cast<int>(_sectionPathRect.bottom - _sectionPathRect.top));
    const int index = itemHeight > 0 ? (pt.y / itemHeight) : -1;

    int newHovered = -1;
    if (index >= 0 && static_cast<size_t>(index) < _editSuggestItems.size() && _editSuggestItems[static_cast<size_t>(index)].enabled)
    {
        newHovered = index;
    }

    if (newHovered != _editSuggestHoveredIndex)
    {
        _editSuggestHoveredIndex = newHovered;
        if (newHovered >= 0 && newHovered != _editSuggestSelectedIndex)
        {
            _editSuggestSelectedIndex = newHovered;
        }
        InvalidateRect(hwnd, nullptr, FALSE);
    }

    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupMouseLeave(HWND hwnd)
{
    if (_editSuggestHoveredIndex != -1)
    {
        _editSuggestHoveredIndex = -1;
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupLButtonDown(HWND /*hwnd*/, POINT pt)
{
    const int itemHeight =
        std::max(1, _editSuggestPopupRowHeightPx > 0 ? _editSuggestPopupRowHeightPx : static_cast<int>(_sectionPathRect.bottom - _sectionPathRect.top));
    const int index = itemHeight > 0 ? (pt.y / itemHeight) : -1;
    if (index >= 0 && static_cast<size_t>(index) < _editSuggestItems.size() && _editSuggestItems[static_cast<size_t>(index)].enabled)
    {
        ApplyEditSuggestIndex(static_cast<size_t>(index));
    }
    return 0;
}

LRESULT NavigationView::EditSuggestPopupWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_CREATE: return OnEditSuggestPopupCreate();
        case WM_NCDESTROY: return OnEditSuggestPopupNcDestroy();
        case WM_ERASEBKGND: return 1;
        case WM_MOUSEACTIVATE: return MA_NOACTIVATE;
        case WM_PAINT: RenderEditSuggestPopup(); return 0;
        case WM_SIZE: return OnEditSuggestPopupSize(hwnd, LOWORD(lp), HIWORD(lp));
        case WM_MOUSEMOVE: return OnEditSuggestPopupMouseMove(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_MOUSELEAVE: return OnEditSuggestPopupMouseLeave(hwnd);
        case WM_LBUTTONDOWN: return OnEditSuggestPopupLButtonDown(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void NavigationView::EnterEditMode()
{
#ifdef ENABLE_TESTS
    ++_debugEnterEditAttemptCount;
    _debugLastEnterEditAbortReason.clear();
#endif

    if (_editMode)
    {
#ifdef ENABLE_TESTS
        _debugLastEnterEditAbortReason = L"already-edit-mode";
#endif
        return;
    }

    if (! _currentPath)
    {
#ifdef ENABLE_TESTS
        _debugLastEnterEditAbortReason = L"missing-current-path";
#endif
        return;
    }

    const auto abortEnterEditMode = [this](std::wstring_view reason) noexcept
    {
        static_cast<void>(reason);
#ifdef ENABLE_TESTS
        ++_debugEnterEditAbortCount;
        _debugLastEnterEditAbortReason.assign(reason);
#endif
        _editMode                        = false;
        _pathEditBlurSuppressActive      = false;
        _pathEditBlurSuppressUntilTickMs = 0;
        _renderMode                      = RenderMode::Breadcrumb;
        _currentEditPath.reset();
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        UpdateHoverTimerState();
    };

    if (_pathEdit && (! _pathEdit->field || ! _pathEdit->hwnd || IsWindow(_pathEdit->hwnd.get()) == FALSE || IsWindowVisible(_pathEdit->hwnd.get()) == FALSE))
    {
        _pathEdit->DeactivateForHideOrDestroy();
        _pathEdit.reset();
    }

    ++_editSuggestEditSessionId;
    _editMode                        = true;
    _renderMode                      = RenderMode::Edit;
    _pathEditBlurSuppressActive      = ! _embeddedDestinationMode;
    _pathEditBlurSuppressUntilTickMs = _pathEditBlurSuppressActive ? (GetTickCount64() + 250u) : 0u;
    _editSuggestItems.clear();
    _editSuggestHighlightText.clear();
    CloseEditSuggestPopup();

    if (! _currentEditPath.has_value())
    {
        _currentEditPath = _currentPath.value();
    }
    const std::filesystem::path& currentPath = _currentEditPath.value();

    if (! _pathEdit)
    {
        if (! RegisterDxHostWndClass(_hInstance))
        {
            abortEnterEditMode(L"register-dx-host-class");
            return;
        }

        const RECT editBounds = GetPathEditBoundsRect(_sectionPathRect, _sectionHistoryRect);
        const auto chrome     = ComputeEditChromeRects(editBounds, _dpi);
        const int hostWidth   = static_cast<int>((std::max)(0L, chrome.editRect.right - chrome.editRect.left));
        const int hostHeight  = static_cast<int>((std::max)(0L, chrome.editRect.bottom - chrome.editRect.top));
        if (hostWidth <= 0 || hostHeight <= 0)
        {
            abortEnterEditMode(L"empty-edit-host-bounds");
            return;
        }

        auto hostState = std::make_unique<NavigationDxTextHost>();
        HWND hwnd      = CreateWindowExW(0,
                                         kDxHostClassName,
                                         L"",
                                         WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                         chrome.editRect.left,
                                         chrome.editRect.top,
                                         hostWidth,
                                         hostHeight,
                                         _hWnd.get(),
                                         reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_PATH_EDIT)),
                                         _hInstance,
                                         &hostState->host);
        if (! hwnd)
        {
            abortEnterEditMode(L"create-edit-host-window");
            return;
        }

        hostState->hwnd.reset(hwnd);
        hostState->host.SetTextInputBackend(RedSalamander::DxUi::TextInputBackend::Native);
        hostState->host.SetTheme(MakeNavigationDxEditPalette(_appTheme, _theme));

        auto field       = std::make_unique<RedSalamander::DxUi::TextField>();
        hostState->field = field.get();
        hostState->field->SetMultiline(false);
        hostState->field->SetClearButtonEnabled(false);
        hostState->field->SetCaretColor(_theme.text);
        hostState->field->SetHorizontalTextPadding(2.0f, 2.0f);
        hostState->field->SetVerticalTextPadding(1.0f, 1.0f);
        hostState->host.SetRoot(std::move(field));
        _pathEdit = std::move(hostState);
    }

    if (! _pathEdit || ! _pathEdit->field || ! _pathEdit->hwnd)
    {
        abortEnterEditMode(L"missing-edit-host-state");
        return;
    }

    _pathEdit->field->SetOnTextChanged([this](std::wstring_view text)
    {
        if (_pathEdit)
        {
            ClearEditValidationError(*_pathEdit);
        }
        _currentEditPath = std::filesystem::path(text);
        UpdateEditSuggest();
    });
    _pathEdit->field->SetOnSubmitted([this]()
    {
        if (_editSuggestSelectedIndex >= 0 && static_cast<size_t>(_editSuggestSelectedIndex) < _editSuggestItems.size())
        {
            ApplyEditSuggestIndex(static_cast<size_t>(_editSuggestSelectedIndex));
            return;
        }

        ExitEditMode(true, L"submitted");
        if (! _editMode && _requestFolderViewFocusCallback)
        {
            _requestFolderViewFocusCallback();
        }
    });
    _pathEdit->field->SetOnPreviewKeyDown([this](RedSalamander::DxUi::WindowHost& /*host*/, UINT virtualKey, UINT modifiers) -> bool
    {
        if ((virtualKey != VK_DOWN && virtualKey != VK_UP) || (modifiers & (MK_CONTROL | MK_SHIFT | kDxUiModifierAlt)) != 0u || ! _editSuggestPopup ||
            _editSuggestItems.empty())
        {
            return false;
        }

        const int count = static_cast<int>(_editSuggestItems.size());
        if (count > 0)
        {
            int next = _editSuggestSelectedIndex;
            if (virtualKey == VK_DOWN)
            {
                next = (next < 0) ? 0 : std::min(next + 1, count - 1);
                while (next < count && ! _editSuggestItems[static_cast<size_t>(next)].enabled)
                {
                    ++next;
                }
                if (next >= count)
                {
                    next = _editSuggestSelectedIndex;
                }
            }
            else
            {
                next = (next < 0) ? (count - 1) : std::max(next - 1, 0);
                while (next >= 0 && ! _editSuggestItems[static_cast<size_t>(next)].enabled)
                {
                    --next;
                }
                if (next < 0)
                {
                    next = _editSuggestSelectedIndex;
                }
            }

            if (next != _editSuggestSelectedIndex)
            {
                _editSuggestSelectedIndex = next;
                InvalidateRect(_editSuggestPopup.get(), nullptr, FALSE);
                return true;
            }

            if (_pathEdit && _pathEdit->field)
            {
                const size_t caretIndex = (virtualKey == VK_DOWN) ? _pathEdit->field->GetText().size() : 0u;
                _pathEdit->field->SetSelectionRange(caretIndex, caretIndex);
                _pathEdit->host.SyncTextInput(_pathEdit->field);
                if (_pathEdit->hwnd)
                {
                    InvalidateRect(_pathEdit->hwnd.get(), nullptr, FALSE);
                }
                return true;
            }
        }
        return false;
    });
    _pathEdit->field->SetOnBlur([this]() noexcept
    {
        if (_editMode && _pathEdit && _pathEdit->hwnd && IsWindowVisible(_pathEdit->hwnd.get()) == FALSE)
        {
            return;
        }

        const HWND focused = GetFocus();
        if (! focused)
        {
            return;
        }
        if (_pathEdit && _pathEdit->hwnd &&
            (focused == _pathEdit->hwnd.get() || focused == _pathEdit->GetTextInputHwnd() || IsChild(_pathEdit->hwnd.get(), focused) != FALSE))
        {
            return;
        }
        if (IsEditValidationPopupWindow(focused))
        {
            return;
        }
        if (! _embeddedDestinationMode && _pathEditBlurSuppressActive && _pathEdit && _pathEdit->hwnd && IsWindowVisible(_pathEdit->hwnd.get()) != FALSE)
        {
            if (GetTickCount64() <= _pathEditBlurSuppressUntilTickMs)
            {
                SetFocus(_pathEdit->hwnd.get());
                return;
            }
            _pathEditBlurSuppressActive = false;
        }
        if (_editMode)
        {
            ExitEditMode(false, L"blur");
        }
    });
    _pathEdit->host.SetOnEscape([this]() -> bool
    {
        if (_editSuggestPopup)
        {
            static_cast<void>(_editSuggestRequestId.fetch_add(1, std::memory_order_acq_rel));
            {
                std::lock_guard lock(_editSuggestMutex);
                _editSuggestPendingQuery.reset();
            }
            _editSuggestAdditionalRequestId = 0;
            _editSuggestAdditionalItems.clear();
            CloseEditSuggestPopup();
            return true;
        }

        ExitEditMode(false, L"escape");
        if (_requestFolderViewFocusCallback)
        {
            _requestFolderViewFocusCallback();
        }
        return true;
    });
    _pathEdit->host.SetOnTabBoundary([this](bool reverse) -> bool
    {
        ExitEditMode(false, L"tab-boundary");
        if (_requestFolderViewFocusCallback)
        {
            _requestFolderViewFocusCallback();
            return true;
        }

        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
        MoveFocus(! reverse);
        return true;
    });

    _pathEdit->field->SetText(currentPath.native());
    _pathEdit->field->SetSelectionRange(0u, currentPath.native().size());
    UpdatePathEditHostLayout();
    ApplyDxEditHostThemes();
    ShowWindow(_pathEdit->hwnd.get(), SW_SHOW);
    InstallEditHostHook(*_pathEdit);
    _pathEdit->host.SetFocusControl(_pathEdit->field);
    SetFocus(_pathEdit->hwnd.get());

    const std::wstring_view currentPathText = currentPath.native();
    const bool endsWithSeparator            = ! currentPathText.empty() && (currentPathText.back() == L'\\' || currentPathText.back() == L'/');
    if (endsWithSeparator)
    {
        UpdateEditSuggest();
    }

    if (_hWnd)
    {
        const RECT editBounds = GetPathEditBoundsRect(_sectionPathRect, _sectionHistoryRect);
        InvalidateRect(_hWnd.get(), &editBounds, FALSE);
    }

#ifdef ENABLE_TESTS
    ++_debugEnterEditSuccessCount;
    _debugLastEnterEditAbortReason.clear();
#endif

    UpdateHoverTimerState();
}

void NavigationView::ExitEditMode(bool accept, std::wstring_view reason)
{
    static_cast<void>(reason);

    if (! _editMode)
        return;

#ifdef ENABLE_TESTS
    ++_debugExitEditCount;
    _debugLastExitEditAccepted = accept;
    _debugLastExitEditReason.assign(reason);
#endif

    CloseEditSuggestPopup();
    static_cast<void>(_editSuggestRequestId.fetch_add(1, std::memory_order_acq_rel));
    ++_editSuggestEditSessionId;
    {
        std::lock_guard lock(_editSuggestMutex);
        _editSuggestPendingQuery.reset();
    }
    _editSuggestMountedInstance.reset();

    std::optional<std::filesystem::path> pendingPathChange;

    if (accept && _pathEdit && _pathEdit->field)
    {
        const std::wstring acceptedText = NavigationLocation::NormalizeUserTypedLocationText(std::wstring(_pathEdit->field->GetText()));

        if (ValidatePath(acceptedText))
        {
            std::filesystem::path newPath(acceptedText);

            const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");
            const std::wstring_view typedText(acceptedText);
            if (! isFilePlugin && ! _currentInstanceContext.empty() && typedText.find(L'|') != std::wstring_view::npos)
            {
                std::wstring_view typedPrefix;
                std::wstring_view typedRemainder;
                if (! TryParsePluginPrefix(typedText, typedPrefix, typedRemainder))
                {
                    std::wstring canonical;
                    canonical.reserve(_pluginShortId.size() + 1u + typedText.size());
                    canonical.append(_pluginShortId);
                    canonical.push_back(L':');
                    canonical.append(typedText);
                    newPath = std::filesystem::path(std::move(canonical));
                }
            }
            pendingPathChange = std::move(newPath);
        }
        else
        {
            const std::wstring message = FormatStringResource(nullptr, IDS_FMT_INVALID_PATH, acceptedText.c_str());
            ShowEditValidationError(*_pathEdit, message);

            if (const HWND inputHwnd = _pathEdit->GetTextInputHwnd(); inputHwnd && IsWindow(inputHwnd) != FALSE)
            {
                SetFocus(inputHwnd);
            }
            else if (_pathEdit->hwnd)
            {
                SetFocus(_pathEdit->hwnd.get());
            }

            _renderMode = RenderMode::Edit;
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
            }
            UpdateHoverTimerState();
            return;
        }
    }

    _editMode                        = false;
    _pathEditBlurSuppressActive      = false;
    _pathEditBlurSuppressUntilTickMs = 0;

    if (_pathEdit && _pathEdit->hwnd)
    {
        if (_pathEdit->field)
        {
            ClearEditValidationError(*_pathEdit);
            _pathEdit->field->SetOnTextChanged({});
            _pathEdit->field->SetOnSubmitted({});
            _pathEdit->field->SetOnPreviewKeyDown({});
            _pathEdit->field->SetOnBlur({});
        }
        _pathEdit->host.SetOnEscape({});
        _pathEdit->host.SetOnTabBoundary({});
        _pathEdit->DeactivateForHideOrDestroy();

        const HWND focused = GetFocus();
        if (focused && (focused == _pathEdit->hwnd.get() || focused == _pathEdit->GetTextInputHwnd() || IsChild(_pathEdit->hwnd.get(), focused) != FALSE) &&
            _hWnd)
        {
            SetFocus(_hWnd.get());
        }
        ShowWindow(_pathEdit->hwnd.get(), SW_HIDE);
    }
    _renderMode = RenderMode::Breadcrumb;
    InvalidateRect(_hWnd.get(), nullptr, FALSE);

    if (pendingPathChange.has_value())
    {
        RequestPathChange(pendingPathChange.value());
    }

    UpdateHoverTimerState();
}

void NavigationView::UpdateEditSuggest()
{
    if (! _editMode || ! _pathEdit || ! _pathEdit->field)
    {
        _editSuggestItems.clear();
        _editSuggestHighlightText.clear();
        CloseEditSuggestPopup();
        return;
    }

    std::wstring text(_pathEdit->field->GetText());

    const uint64_t requestId        = _editSuggestRequestId.fetch_add(1, std::memory_order_acq_rel) + 1u;
    _editSuggestAdditionalRequestId = 0;
    _editSuggestAdditionalItems.clear();

    std::wstring normalizedInput = TrimWhitespace(text);
    if (normalizedInput.size() >= 2u && normalizedInput.front() == L'"' && normalizedInput.back() == L'"')
    {
        normalizedInput = normalizedInput.substr(1, normalizedInput.size() - 2u);
        normalizedInput = TrimWhitespace(normalizedInput);
    }

    const auto startsWithNoCase = [](std::wstring_view value, std::wstring_view prefix) noexcept
    { return ! prefix.empty() && OrdinalString::StartsWithNoCase(value, prefix); };

    auto showStaticSuggestions = [&](std::vector<EditSuggestItem>&& items, std::wstring&& highlightText)
    {
        _editSuggestItems         = std::move(items);
        _editSuggestHoveredIndex  = -1;
        _editSuggestSelectedIndex = -1;
        _editSuggestHighlightText = std::move(highlightText);

        if (_editSuggestItems.empty())
        {
            CloseEditSuggestPopup();
        }
        else
        {
            UpdateEditSuggestPopupWindow();
        }
    };

    const auto buildProtocolAndDriveSuggestions = [&](std::wstring_view filterText) -> std::vector<EditSuggestItem>
    {
        std::vector<EditSuggestItem> items;
        if (filterText.empty())
        {
            return items;
        }

        // Host-level reserved prefix to route to Connection Manager profiles.
        if (filterText.front() == L'@' && startsWithNoCase(L"@conn:", filterText))
        {
            EditSuggestItem item{};
            item.display            = L"@conn:";
            item.insertText         = L"@conn:";
            item.directorySeparator = L'\0';
            items.push_back(std::move(item));
        }

        // File system plugins (shortId:)
        FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
        const auto& plugins                    = pluginManager.GetPlugins();
        for (const auto& entry : plugins)
        {
            if (entry.shortId.empty() || ! entry.loadable || entry.disabled)
            {
                continue;
            }

            if (! startsWithNoCase(entry.shortId, filterText))
            {
                continue;
            }

            EditSuggestItem item{};
            item.display            = entry.shortId + L":";
            item.insertText         = item.display;
            item.directorySeparator = L'\0';
            items.push_back(std::move(item));
        }

        // Drive roots (C:\)
        const DWORD drives = GetLogicalDrives();
        if (drives != 0)
        {
            const auto isAlpha = [](wchar_t ch) noexcept { return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z'); };

            const bool driveQuery =
                ! filterText.empty() && isAlpha(filterText.front()) && (filterText.size() == 1u || (filterText.size() == 2u && filterText[1] == L':'));
            if (driveQuery)
            {
                const wchar_t wanted = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(filterText.front())));
                for (int i = 0; i < 26; ++i)
                {
                    if ((drives & (static_cast<DWORD>(1u) << i)) == 0)
                    {
                        continue;
                    }

                    const wchar_t driveLetter = static_cast<wchar_t>(L'A' + i);
                    if (driveLetter != wanted)
                    {
                        continue;
                    }

                    std::wstring root;
                    root.push_back(driveLetter);
                    root.append(L":\\");

                    EditSuggestItem item{};
                    item.display            = root;
                    item.insertText         = std::move(root);
                    item.directorySeparator = L'\\';
                    items.push_back(std::move(item));
                }
            }
        }

        std::sort(
            items.begin(), items.end(), [](const EditSuggestItem& a, const EditSuggestItem& b) { return _wcsicmp(a.display.c_str(), b.display.c_str()) < 0; });

        if (items.size() > kEditSuggestMaxItems)
        {
            items.resize(kEditSuggestMaxItems);
        }

        return items;
    };

    auto buildConnectionSuggestions = [&](std::wstring_view filterText,
                                          std::wstring_view insertPrefix,
                                          std::wstring_view filterPluginId,
                                          bool usePluginFilter,
                                          wchar_t directorySeparator) -> std::vector<EditSuggestItem>
    {
        std::vector<EditSuggestItem> items;
        if (! _settings)
        {
            return items;
        }

        struct Candidate
        {
            std::wstring sortKey;
            std::wstring display;
            std::wstring name;
        };

        std::vector<Candidate> candidates;

        const auto& plugins                 = FileSystemPluginManager::GetInstance().GetPlugins();
        const auto tryGetShortIdForPluginId = [&](std::wstring_view pluginId) noexcept -> std::wstring_view
        {
            for (const auto& entry : plugins)
            {
                if (! entry.id.empty() && EqualsNoCase(entry.id, pluginId) && ! entry.shortId.empty())
                {
                    return entry.shortId;
                }
            }
            return {};
        };

        const auto buildPreview = [&](const Common::Settings::ConnectionProfile& profile) -> std::wstring
        {
            const std::wstring_view shortId = tryGetShortIdForPluginId(profile.pluginId);
            if (shortId.empty() || profile.host.empty())
            {
                return {};
            }

            std::wstring host = profile.host;
            if (profile.port != 0u)
            {
                host = std::format(L"{}:{}", profile.host, profile.port);
            }

            if (! profile.userName.empty())
            {
                return std::format(L"{}://{}@{}", shortId, profile.userName, host);
            }

            return std::format(L"{}://{}", shortId, host);
        };

        const auto tryAddProfile = [&](std::wstring_view name, const Common::Settings::ConnectionProfile& profile, std::wstring_view labelOverride)
        {
            if (name.empty())
            {
                return;
            }

            if (usePluginFilter && ! filterPluginId.empty() && ! EqualsNoCase(profile.pluginId, filterPluginId))
            {
                return;
            }

            const std::wstring_view labelView = labelOverride.empty() ? std::wstring_view(profile.name) : labelOverride;

            if (! filterText.empty() && ! ContainsInsensitive(name, filterText) && ! ContainsInsensitive(labelView, filterText))
            {
                return;
            }

            Candidate c{};
            c.sortKey = std::wstring(name);
            c.name    = std::wstring(name);

            const std::wstring preview = buildPreview(profile);
            if (! labelOverride.empty())
            {
                c.display = preview.empty() ? std::format(L"{} — {}", name, labelOverride) : std::format(L"{} — {} — {}", name, labelOverride, preview);
            }
            else
            {
                c.display = preview.empty() ? std::wstring(name) : std::format(L"{} — {}", name, preview);
            }

            candidates.push_back(std::move(c));
        };

        // Quick Connect (session-only)
        {
            Common::Settings::ConnectionProfile quick{};
            const std::wstring_view preferredPluginId =
                usePluginFilter && ! filterPluginId.empty() ? filterPluginId : FileSystemPluginManager::GetInstance().GetActivePluginId();
            RedSalamander::Connections::EnsureQuickConnectProfile(preferredPluginId);
            RedSalamander::Connections::GetQuickConnectProfile(quick);

            const std::wstring quickLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
            tryAddProfile(RedSalamander::Connections::kQuickConnectConnectionName, quick, quickLabel);
        }

        // Persisted profiles
        if (_settings->connections)
        {
            for (const auto& profile : _settings->connections->items)
            {
                if (profile.name.empty() || profile.pluginId.empty())
                {
                    continue;
                }
                tryAddProfile(profile.name, profile, {});
            }
        }

        std::sort(
            candidates.begin(), candidates.end(), [](const Candidate& a, const Candidate& b) { return _wcsicmp(a.sortKey.c_str(), b.sortKey.c_str()) < 0; });

        const size_t maxVisible = std::min(kEditSuggestMaxItems, static_cast<size_t>(10u));
        for (const auto& c : candidates)
        {
            if (items.size() >= maxVisible)
            {
                break;
            }
            EditSuggestItem item{};
            item.display            = c.display;
            item.insertText         = std::format(L"{}{}", insertPrefix, c.name);
            item.directorySeparator = directorySeparator;
            items.push_back(std::move(item));
        }

        if (candidates.size() > items.size() && items.size() < kEditSuggestMaxItems)
        {
            EditSuggestItem item{};
            item.display            = std::wstring(kEllipsisText);
            item.enabled            = false;
            item.directorySeparator = L'\0';
            items.push_back(std::move(item));
        }

        return items;
    };

    // `nav:` / `nav://` (Connection Manager routing)
    if (startsWithNoCase(normalizedInput, L"nav:"))
    {
        std::wstring rest = TrimWhitespace(std::wstring_view(normalizedInput).substr(4u));
        if (rest.size() >= 2u && rest[0] == L'/' && rest[1] == L'/')
        {
            rest.erase(0, 2);
        }

        std::wstring highlight = rest;
        std::wstring prefix    = startsWithNoCase(normalizedInput, L"nav://") ? L"nav://" : L"nav:";
        showStaticSuggestions(buildConnectionSuggestions(rest, prefix, {}, false, L'\0'), std::move(highlight));
        return;
    }

    // `@conn:` (Connection Manager routing alias)
    if (startsWithNoCase(normalizedInput, L"@conn:"))
    {
        std::wstring rest = TrimWhitespace(std::wstring_view(normalizedInput).substr(6u));
        showStaticSuggestions(buildConnectionSuggestions(rest, L"@conn:", {}, false, L'\0'), std::move(rest));
        return;
    }

    // Protocol-local Connection Manager prefix (ex: `ftp:/@conn:`)
    {
        std::wstring_view typedPrefix;
        std::wstring_view typedRemainder;
        if (TryParsePluginPrefix(normalizedInput, typedPrefix, typedRemainder))
        {
            const bool supportsConnections = EqualsNoCase(typedPrefix, L"ftp") || EqualsNoCase(typedPrefix, L"sftp") || EqualsNoCase(typedPrefix, L"scp") ||
                                             EqualsNoCase(typedPrefix, L"imap") || EqualsNoCase(typedPrefix, L"gdrive") ||
                                             EqualsNoCase(typedPrefix, L"onedrive") || EqualsNoCase(typedPrefix, L"onedrive-pro") ||
                                             EqualsNoCase(typedPrefix, L"sharepoint");

            if (supportsConnections && ! typedRemainder.empty() && typedRemainder.find(L'|') == std::wstring_view::npos)
            {
                std::wstring rem(typedRemainder);
                for (wchar_t& ch : rem)
                {
                    if (ch == L'\\')
                    {
                        ch = L'/';
                    }
                }

                if (! rem.empty() && rem.front() == L'@')
                {
                    rem.insert(rem.begin(), L'/');
                }

                std::wstring_view remView(rem);
                if (startsWithNoCase(remView, L"/@conn:"))
                {
                    std::wstring_view after(remView);
                    after.remove_prefix(7u); // "/@conn:"

                    const size_t nextSlash = after.find(L'/');
                    if (nextSlash == std::wstring_view::npos)
                    {
                        std::wstring insertPrefix = std::wstring(typedPrefix) + L":/@conn:";
                        std::wstring_view pluginIdFilter;
                        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
                        for (const auto& entry : plugins)
                        {
                            if (! entry.shortId.empty() && EqualsNoCase(entry.shortId, typedPrefix) && ! entry.id.empty())
                            {
                                pluginIdFilter = entry.id;
                                break;
                            }
                        }

                        std::wstring highlight = std::wstring(after);
                        showStaticSuggestions(buildConnectionSuggestions(after, insertPrefix, pluginIdFilter, true, L'/'), std::move(highlight));
                        return;
                    }
                }
                else if (startsWithNoCase(remView, L"/@") && remView.find(L'/') == 0u)
                {
                    // Complete `/@` to the reserved Connection Manager prefix.
                    std::wstring_view after(remView);
                    after.remove_prefix(2u); // "/@"

                    if (startsWithNoCase(L"conn:", after) || startsWithNoCase(L"conn", after) || after.empty())
                    {
                        EditSuggestItem item{};
                        item.display            = L"@conn:";
                        item.insertText         = std::wstring(typedPrefix) + L":/@conn:";
                        item.directorySeparator = L'\0';

                        std::vector<EditSuggestItem> items;
                        items.push_back(std::move(item));
                        showStaticSuggestions(std::move(items), std::wstring(after));
                        return;
                    }
                }
            }
        }
    }

    EditSuggestParseResult parseResult{};
    if (! TryParseEditSuggestQuery(normalizedInput, _pluginShortId, _currentEditPath, parseResult))
    {
        auto items = buildProtocolAndDriveSuggestions(normalizedInput);
        showStaticSuggestions(std::move(items), std::move(normalizedInput));
        return;
    }

    const auto isFileShortId = [](std::wstring_view shortId) noexcept { return shortId.empty() || EqualsNoCase(shortId, L"file"); };

    wil::com_ptr<IFileSystem> fileSystem = nullptr;
    std::shared_ptr<EditSuggestFileSystemInstance> keepAlive;

    const bool needsInstanceContext =
        parseResult.instanceContextSpecified && ! parseResult.instanceContext.empty() && ! isFileShortId(parseResult.enumerationShortId);

    if (! needsInstanceContext && _fileSystemPlugin &&
        (EqualsNoCase(parseResult.enumerationShortId, _pluginShortId) || (isFileShortId(parseResult.enumerationShortId) && isFileShortId(_pluginShortId))))
    {
        fileSystem = _fileSystemPlugin;
    }
    else if (needsInstanceContext && _fileSystemPlugin && EqualsNoCase(parseResult.enumerationShortId, _pluginShortId) &&
             EqualsNoCase(parseResult.instanceContext, _currentInstanceContext))
    {
        fileSystem = _fileSystemPlugin;
    }
    else
    {
        FileSystemPluginManager& plugins = FileSystemPluginManager::GetInstance();
        const auto& allPlugins           = plugins.GetPlugins();

        const FileSystemPluginManager::PluginEntry* entry = nullptr;
        for (const auto& candidate : allPlugins)
        {
            if (candidate.shortId.empty())
            {
                continue;
            }

            if (EqualsNoCase(candidate.shortId, parseResult.enumerationShortId))
            {
                entry = &candidate;
                break;
            }
        }

        if (entry != nullptr)
        {
            if (needsInstanceContext)
            {
                if (_editSuggestMountedInstance && EqualsNoCase(_editSuggestMountedInstance->pluginShortId, parseResult.enumerationShortId) &&
                    EqualsNoCase(_editSuggestMountedInstance->instanceContext, parseResult.instanceContext))
                {
                    keepAlive  = _editSuggestMountedInstance;
                    fileSystem = keepAlive->fileSystem;
                }
                else
                {
                    if (! entry->path.empty())
                    {
                        wil::unique_hmodule module(LoadLibraryExW(entry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));

                        if (module)
                        {
                            using CreateFactoryFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

#pragma warning(push)
#pragma warning(disable : 4191) // C4191: unsafe conversion from FARPROC
                            const auto createFactory = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(module.get(), "RedSalamanderCreate"));
#pragma warning(pop)

                            if (createFactory)
                            {
                                FactoryOptions options{};
                                options.debugLevel = DEBUG_LEVEL_NONE;

                                wil::com_ptr<IFileSystem> created;
                                const std::wstring requestedPluginId = entry->factoryPluginId.empty() ? entry->id : entry->factoryPluginId;
                                HRESULT createHr                     = E_INVALIDARG;
                                if (! requestedPluginId.empty())
                                {
                                    createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), requestedPluginId.c_str(), created.put_void());
                                }
                                if (SUCCEEDED(createHr) && created)
                                {
                                    std::string configurationJsonUtf8;
                                    if (entry->informations)
                                    {
                                        const char* configuration = nullptr;
                                        if (SUCCEEDED(entry->informations->GetConfiguration(&configuration)) && configuration != nullptr &&
                                            configuration[0] != '\0')
                                        {
                                            configurationJsonUtf8 = configuration;
                                        }
                                    }

                                    if (! configurationJsonUtf8.empty())
                                    {
                                        wil::com_ptr<IInformations> createdInfos;
                                        if (SUCCEEDED(created->QueryInterface(__uuidof(IInformations), createdInfos.put_void())) && createdInfos)
                                        {
                                            static_cast<void>(createdInfos->SetConfiguration(configurationJsonUtf8.c_str()));
                                        }
                                    }

                                    wil::com_ptr<IFileSystemInitialize> initializer;
                                    const HRESULT initQi = created->QueryInterface(__uuidof(IFileSystemInitialize), initializer.put_void());
                                    if (SUCCEEDED(initQi) && initializer && SUCCEEDED(initializer->Initialize(parseResult.instanceContext.c_str(), nullptr)))
                                    {
                                        auto instance        = std::make_shared<EditSuggestFileSystemInstance>();
                                        instance->module     = std::move(module);
                                        instance->fileSystem = created;
                                        instance->pluginShortId.assign(parseResult.enumerationShortId);
                                        instance->instanceContext = parseResult.instanceContext;

                                        _editSuggestMountedInstance = instance;
                                        keepAlive                   = std::move(instance);
                                        fileSystem                  = created;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                fileSystem = entry->fileSystem;
            }
        }
    }

    if (! fileSystem)
    {
        _editSuggestItems.clear();
        _editSuggestHighlightText.clear();
        CloseEditSuggestPopup();
        return;
    }

    std::vector<EditSuggestItem> additionalItems;
    {
        const std::wstring_view view(normalizedInput);
        const bool hasSeparator = view.find_first_of(L"\\/") != std::wstring_view::npos;
        const size_t colonPos   = view.find(L':');

        const auto isAlpha = [](wchar_t ch) noexcept { return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z'); };

        const bool driveLike = view.size() <= 2u && ! view.empty() && isAlpha(view.front()) && (view.size() == 1u || view[1] == L':');

        if (! hasSeparator && (! view.empty()) && (view.front() == L'@' || colonPos == std::wstring_view::npos || driveLike))
        {
            additionalItems = buildProtocolAndDriveSuggestions(view);
        }
    }

    std::vector<std::wstring> names;
    bool usedCache = false;

    if (fileSystem)
    {
        auto borrowed =
            DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(fileSystem.get(), parseResult.pluginFolder, DirectoryInfoCache::BorrowMode::CacheOnly);
        IFilesInformation* info = borrowed.Get();
        if (borrowed.Status() == S_OK && info)
        {
            usedCache = true;
            AppendMatchingDirectoryNamesFromFilesInformation(info, parseResult.filter, names);
        }
    }

    if (usedCache)
    {
        const bool hasMore = SortAndTrimEditSuggestNames(names);

        std::vector<std::wstring> displayItems;
        std::vector<std::wstring> insertItems;
        BuildEditSuggestLists(parseResult.displayFolder, names, parseResult.directorySeparator, displayItems, insertItems);

        const size_t count = std::min(displayItems.size(), insertItems.size());

        std::vector<EditSuggestItem> merged;
        merged.reserve(kEditSuggestMaxItems);

        const size_t maxWithoutEllipsis = (hasMore && kEditSuggestMaxItems > 0u) ? (kEditSuggestMaxItems - 1u) : kEditSuggestMaxItems;

        for (auto& item : additionalItems)
        {
            if (merged.size() >= maxWithoutEllipsis)
            {
                break;
            }
            merged.push_back(std::move(item));
        }

        for (size_t i = 0; i < count && merged.size() < maxWithoutEllipsis; ++i)
        {
            EditSuggestItem item{};
            item.display            = std::move(displayItems[i]);
            item.insertText         = std::move(insertItems[i]);
            item.directorySeparator = parseResult.directorySeparator;
            merged.push_back(std::move(item));
        }

        if (hasMore && merged.size() < kEditSuggestMaxItems)
        {
            EditSuggestItem item{};
            item.display            = std::wstring(kEllipsisText);
            item.enabled            = false;
            item.directorySeparator = L'\0';
            merged.push_back(std::move(item));
        }

        _editSuggestHoveredIndex  = -1;
        _editSuggestSelectedIndex = -1;
        _editSuggestHighlightText = parseResult.filter;
        _editSuggestItems         = std::move(merged);
        UpdateEditSuggestPopupWindow();
        return;
    }

    EnsureEditSuggestWorker();
    {
        std::lock_guard lock(_editSuggestMutex);
        EditSuggestQuery query{};
        query.requestId          = requestId;
        query.editSessionId      = _editSuggestEditSessionId;
        query.fileSystem         = fileSystem;
        query.displayFolder      = parseResult.displayFolder;
        query.pluginFolder       = parseResult.pluginFolder;
        query.queryText          = text;
        query.prefix             = std::move(parseResult.filter);
        query.directorySeparator = parseResult.directorySeparator;
        query.keepAlive          = keepAlive;
        _editSuggestPendingQuery = std::move(query);
    }
    _editSuggestCv.notify_one();

    if (! additionalItems.empty())
    {
        _editSuggestAdditionalRequestId = requestId;
        _editSuggestAdditionalItems     = additionalItems;

        _editSuggestItems         = std::move(additionalItems);
        _editSuggestHoveredIndex  = -1;
        _editSuggestSelectedIndex = -1;
        _editSuggestHighlightText = parseResult.filter;
        UpdateEditSuggestPopupWindow();
    }
    else
    {
        _editSuggestItems.clear();
        _editSuggestHoveredIndex  = -1;
        _editSuggestSelectedIndex = -1;
        _editSuggestHighlightText.clear();
        CloseEditSuggestPopup();
    }
}

void NavigationView::UpdateEditSuggestPopupWindow()
{
    if (! _hWnd || ! _editMode || ! _pathEdit || _editSuggestItems.empty())
    {
        CloseEditSuggestPopup();
        return;
    }

    EnsureD2DResources();
    if (! _d2dFactory || ! _dwriteFactory || ! _pathFormat)
    {
        return;
    }

    if (! RegisterEditSuggestPopupWndClass(_hInstance))
    {
        return;
    }

    const RECT editBounds = GetPathEditBoundsRect(_sectionPathRect, _sectionHistoryRect);
    const auto chrome     = ComputeEditChromeRects(editBounds, _dpi);

    const int navHeightPx         = std::max(1, static_cast<int>(_sectionPathRect.bottom - _sectionPathRect.top));
    const int minRowHeightPx      = std::max(1, DipsToPixelsInt(40, _dpi));
    const int itemHeight          = std::max(navHeightPx, minRowHeightPx);
    _editSuggestPopupRowHeightPx  = itemHeight;
    const int desiredClientWidth  = std::max(1l, chrome.editRect.right - chrome.editRect.left);
    const size_t itemCount        = std::min(kEditSuggestMaxItems, _editSuggestItems.size());
    const int desiredClientHeight = std::max(1, static_cast<int>(itemCount) * itemHeight);

    const DWORD style   = WS_POPUP;
    const DWORD exStyle = WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;

    RECT windowRect = {0, 0, desiredClientWidth, desiredClientHeight};
    if (! AdjustWindowRectExForDpi(&windowRect, style, FALSE, exStyle, _dpi))
    {
        AdjustWindowRectEx(&windowRect, style, FALSE, exStyle);
    }

    const int winWidth  = windowRect.right - windowRect.left;
    const int winHeight = windowRect.bottom - windowRect.top;

    POINT anchor = {chrome.editRect.left, editBounds.bottom};
    ClientToScreen(_hWnd.get(), &anchor);

    HMONITOR hMon = MonitorFromPoint(anchor, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (! GetMonitorInfoW(hMon, &mi))
    {
        return;
    }

    const RECT work = mi.rcWork;

    int x = anchor.x;
    int y = anchor.y;

    if (y + winHeight > work.bottom)
    {
        const int aboveY = anchor.y - winHeight;
        if (aboveY >= work.top)
        {
            y = aboveY;
        }
        else
        {
            y = std::max(static_cast<int>(work.top), static_cast<int>(work.bottom - winHeight));
        }
    }

    if (x + winWidth > work.right)
    {
        x = std::max(static_cast<int>(work.left), static_cast<int>(work.right - winWidth));
    }

    const int minX = static_cast<int>(work.left);
    const int minY = static_cast<int>(work.top);
    const int maxX = std::max(minX, static_cast<int>(work.right - winWidth));
    const int maxY = std::max(minY, static_cast<int>(work.bottom - winHeight));
    x              = std::clamp(x, minX, maxX);
    y              = std::clamp(y, minY, maxY);

    if (! _editSuggestPopup)
    {
        HWND popup = CreateWindowExW(exStyle, kSuggestPopupClassName, L"", style, x, y, winWidth, winHeight, _hWnd.get(), nullptr, _hInstance, this);
        if (! popup)
        {
            return;
        }

        _editSuggestPopup.reset(popup);
    }
    else
    {
        SetWindowPos(_editSuggestPopup.get(), HWND_TOP, x, y, winWidth, winHeight, SWP_NOACTIVATE);
    }

    RECT clientRect{};
    GetClientRect(_editSuggestPopup.get(), &clientRect);
    _editSuggestPopupClientSize.cx = clientRect.right - clientRect.left;
    _editSuggestPopupClientSize.cy = clientRect.bottom - clientRect.top;

    ShowWindow(_editSuggestPopup.get(), SW_SHOWNOACTIVATE);
    InvalidateRect(_editSuggestPopup.get(), nullptr, FALSE);
}

void NavigationView::CloseEditSuggestPopup()
{
    if (_editSuggestPopup)
    {
        _editSuggestPopup.reset();
        return;
    }

    _editSuggestItems.clear();
    _editSuggestHoveredIndex  = -1;
    _editSuggestSelectedIndex = -1;
    _editSuggestHighlightText.clear();
}

void NavigationView::EnsureEditSuggestPopupD2DResources()
{
    if (! _editSuggestPopup)
    {
        return;
    }

    EnsureD2DResources();
    if (! _d2dFactory)
    {
        return;
    }

    if (! _editSuggestPopupTarget)
    {
        D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        props.dpiX                          = 96.0f;
        props.dpiY                          = 96.0f;

        D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(
            _editSuggestPopup.get(), D2D1::SizeU(static_cast<UINT32>(_editSuggestPopupClientSize.cx), static_cast<UINT32>(_editSuggestPopupClientSize.cy)));

        wil::com_ptr<ID2D1HwndRenderTarget> target;
        if (FAILED(_d2dFactory->CreateHwndRenderTarget(props, hwndProps, target.addressof())))
        {
            return;
        }

        target->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
        _editSuggestPopupTarget = std::move(target);
    }

    if (_editSuggestPopupTarget)
    {
        if (! _editSuggestPopupBackgroundBrush)
        {
            const COLORREF surface = _appTheme.systemHighContrast ? GetSysColor(COLOR_WINDOW) : _appTheme.menu.background;
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(surface), _editSuggestPopupBackgroundBrush.addressof());
        }
        if (! _editSuggestPopupTextBrush)
        {
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(_appTheme.menu.text), _editSuggestPopupTextBrush.addressof());
        }
        if (! _editSuggestPopupDisabledTextBrush)
        {
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(_appTheme.menu.disabledText), _editSuggestPopupDisabledTextBrush.addressof());
        }
        if (! _editSuggestPopupHighlightBrush)
        {
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(_appTheme.menu.selectionBg), _editSuggestPopupHighlightBrush.addressof());
        }
        if (! _editSuggestPopupHoverBrush)
        {
            const COLORREF surface    = _appTheme.systemHighContrast ? GetSysColor(COLOR_WINDOW) : _appTheme.menu.background;
            const int highlightWeight = _appTheme.dark ? 30 : 18;
            const COLORREF highlightColor =
                _appTheme.systemHighContrast ? GetSysColor(COLOR_HIGHLIGHT) : UiMetrics::BlendColorRefWeightedTruncate(surface, _appTheme.menu.text, highlightWeight, 255);
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(highlightColor), _editSuggestPopupHoverBrush.addressof());
        }
        if (! _editSuggestPopupBorderBrush)
        {
            if (! _appTheme.systemHighContrast)
            {
                const COLORREF surface = _appTheme.menu.background;
                const COLORREF border  = UiMetrics::BlendColorRefWeightedTruncate(surface, _appTheme.menu.text, _appTheme.dark ? 60 : 40, 255);
                _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(border), _editSuggestPopupBorderBrush.addressof());
            }
        }
    }
}

void NavigationView::DiscardEditSuggestPopupD2DResources()
{
    _editSuggestPopupBorderBrush       = nullptr;
    _editSuggestPopupBackgroundBrush   = nullptr;
    _editSuggestPopupHoverBrush        = nullptr;
    _editSuggestPopupHighlightBrush    = nullptr;
    _editSuggestPopupDisabledTextBrush = nullptr;
    _editSuggestPopupTextBrush         = nullptr;
    _editSuggestPopupTarget            = nullptr;
}

void NavigationView::RenderEditSuggestPopup()
{
    if (! _editSuggestPopup)
    {
        return;
    }

    PAINTSTRUCT ps;
    wil::unique_hdc_paint hdc = wil::BeginPaint(_editSuggestPopup.get(), &ps);
    static_cast<void>(hdc);

    EnsureEditSuggestPopupD2DResources();
    if (! _editSuggestPopupTarget || ! _dwriteFactory || ! _pathFormat || ! _editSuggestPopupBackgroundBrush || ! _editSuggestPopupTextBrush)
    {
        return;
    }

    _editSuggestPopupTarget->BeginDraw();
    auto endDraw = wil::scope_exit([&]
    {
        const HRESULT hr = _editSuggestPopupTarget->EndDraw();
        if (hr == D2DERR_RECREATE_TARGET)
        {
            DiscardEditSuggestPopupD2DResources();
        }
    });

    const float width            = static_cast<float>(_editSuggestPopupClientSize.cx);
    const float height           = static_cast<float>(_editSuggestPopupClientSize.cy);
    const D2D1_RECT_F clientRect = D2D1::RectF(0.0f, 0.0f, width, height);

    _editSuggestPopupTarget->FillRectangle(clientRect, _editSuggestPopupBackgroundBrush.get());

    const float rowHeight = static_cast<float>(
        std::max(1, _editSuggestPopupRowHeightPx > 0 ? _editSuggestPopupRowHeightPx : static_cast<int>(_sectionPathRect.bottom - _sectionPathRect.top)));

    const float highlightInsetX = DipsToPixels(6.0f, _dpi);
    const float highlightInsetY = DipsToPixels(2.0f, _dpi);
    const float highlightRadius = DipsToPixels(8.0f, _dpi);

    const float barWidth  = DipsToPixels(5.0f, _dpi);
    const float barInsetX = DipsToPixels(4.0f, _dpi);
    const float barInsetY = DipsToPixels(4.0f, _dpi);
    const float barRadius = DipsToPixels(4.0f, _dpi);

    const float textInsetX       = DipsToPixels(22.0f, _dpi);
    const float textPaddingRight = DipsToPixels(22.0f, _dpi);

    const int activeIndex = _editSuggestSelectedIndex >= 0 ? _editSuggestSelectedIndex : _editSuggestHoveredIndex;

    const size_t count = std::min(kEditSuggestMaxItems, _editSuggestItems.size());
    for (size_t i = 0; i < count; ++i)
    {
        const float top     = rowHeight * static_cast<float>(i);
        D2D1_RECT_F rowRect = D2D1::RectF(0.0f, top, width, top + rowHeight);

        const auto& item    = _editSuggestItems[i];
        const bool enabled  = item.enabled;
        const bool selected = enabled && (static_cast<int>(i) == activeIndex);
        if (selected && _editSuggestPopupHoverBrush)
        {
            const D2D1_RECT_F highlightRect = InsetRectF(rowRect, highlightInsetX, highlightInsetY);
            _editSuggestPopupTarget->FillRoundedRectangle(RoundedRect(highlightRect, highlightRadius), _editSuggestPopupHoverBrush.get());

            if (_editSuggestPopupHighlightBrush)
            {
                D2D1_RECT_F barRect = highlightRect;
                barRect.left        = std::min(barRect.right, barRect.left + barInsetX);
                barRect.right       = std::min(barRect.right, barRect.left + barWidth);
                barRect.top         = std::min(barRect.bottom, barRect.top + barInsetY);
                barRect.bottom      = std::max(barRect.top, barRect.bottom - barInsetY);

                _editSuggestPopupTarget->FillRoundedRectangle(RoundedRect(barRect, barRadius), _editSuggestPopupHighlightBrush.get());
            }
        }

        D2D1_RECT_F textRect = rowRect;
        textRect.left        = std::min(textRect.right, textRect.left + textInsetX);
        textRect.right       = std::max(textRect.left, textRect.right - textPaddingRight);

        const auto& text = item.display;
        if (text.empty())
        {
            continue;
        }

        wil::com_ptr<IDWriteTextLayout> layout;
        const float layoutWidth  = std::max(1.0f, textRect.right - textRect.left);
        const float layoutHeight = std::max(1.0f, rowHeight);
        if (SUCCEEDED(_dwriteFactory->CreateTextLayout(
                text.data(), static_cast<UINT32>(text.size()), _pathFormat.get(), layoutWidth, layoutHeight, layout.addressof())) &&
            layout)
        {
            layout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
            layout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

            if (enabled && ! _editSuggestHighlightText.empty() && _editSuggestPopupHighlightBrush)
            {
                size_t searchStart = 0;
                while (searchStart < text.size())
                {
                    const size_t remaining = text.size() - searchStart;
                    const size_t maxInt    = static_cast<size_t>(std::numeric_limits<int>::max());
                    const int sourceLen    = static_cast<int>(std::min(remaining, maxInt));
                    const int valueLen     = static_cast<int>(std::min(_editSuggestHighlightText.size(), maxInt));
                    if (valueLen <= 0 || sourceLen < valueLen)
                    {
                        break;
                    }

                    const int foundAt = FindStringOrdinal(0, text.data() + searchStart, sourceLen, _editSuggestHighlightText.data(), valueLen, TRUE);
                    if (foundAt < 0)
                    {
                        break;
                    }

                    const size_t matchStart  = searchStart + static_cast<size_t>(foundAt);
                    const size_t matchLength = std::min(_editSuggestHighlightText.size(), text.size() - matchStart);
                    if (matchLength == 0u)
                    {
                        break;
                    }

                    const size_t maxUInt32 = static_cast<size_t>(std::numeric_limits<UINT32>::max());
                    const UINT32 startPos  = static_cast<UINT32>(std::min(matchStart, maxUInt32));
                    const UINT32 len       = static_cast<UINT32>(std::min(matchLength, maxUInt32));
                    const DWRITE_TEXT_RANGE range{startPos, len};

                    static_cast<void>(layout->SetDrawingEffect(_editSuggestPopupHighlightBrush.get(), range));
                    static_cast<void>(layout->SetFontWeight(DWRITE_FONT_WEIGHT_SEMI_BOLD, range));

                    searchStart = matchStart + matchLength;
                }
            }

            ID2D1SolidColorBrush* brush = _editSuggestPopupTextBrush.get();
            if (! enabled && _editSuggestPopupDisabledTextBrush)
            {
                brush = _editSuggestPopupDisabledTextBrush.get();
            }

            _editSuggestPopupTarget->DrawTextLayout(D2D1::Point2F(textRect.left, rowRect.top), layout.get(), brush, D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
        }
    }

    if (_editSuggestPopupBorderBrush)
    {
        const D2D1_RECT_F borderRect = InsetRectF(clientRect, 0.5f, 0.5f);
        _editSuggestPopupTarget->DrawRoundedRectangle(RoundedRect(borderRect, highlightRadius), _editSuggestPopupBorderBrush.get(), 1.0f);
    }
}

void NavigationView::ApplyEditSuggestIndex(size_t index)
{
    if (! _editMode || ! _pathEdit || ! _pathEdit->field || index >= _editSuggestItems.size())
    {
        return;
    }

    const auto& item = _editSuggestItems[index];
    if (! item.enabled || item.insertText.empty())
    {
        return;
    }

    std::wstring text = item.insertText;
    if (! text.empty() && text.back() != L'\\' && text.back() != L'/')
    {
        if (item.directorySeparator != L'\0')
        {
            text.push_back(item.directorySeparator);
        }
    }

    _pathEdit->field->SetText(text);
    _pathEdit->field->SetSelectionRange(text.size(), text.size());
    _currentEditPath = std::filesystem::path(text);
    _pathEdit->host.SetFocusControl(_pathEdit->field);
    SetFocus(_pathEdit->hwnd.get());
    _editSuggestHoveredIndex  = -1;
    _editSuggestSelectedIndex = -1;
    _editSuggestHighlightText.clear();
    static_cast<void>(_editSuggestRequestId.fetch_add(1, std::memory_order_acq_rel));
    {
        std::lock_guard lock(_editSuggestMutex);
        _editSuggestPendingQuery.reset();
    }
    _editSuggestAdditionalRequestId = 0;
    _editSuggestAdditionalItems.clear();
    CloseEditSuggestPopup();
}

bool NavigationView::TryHandleEditClipboardCommand(UINT commandId) noexcept
{
    if (! _editMode || ! _pathEdit || ! _pathEdit->field || ! _pathEdit->hwnd)
    {
        return false;
    }

    const HWND editHwnd = _pathEdit->hwnd.get();
    const HWND focus    = GetFocus();
    if (focus != editHwnd && (focus == nullptr || IsChild(editHwnd, focus) == FALSE))
    {
        return false;
    }

    SetFocus(editHwnd);
    _pathEdit->host.SetFocusControl(_pathEdit->field);
    const auto syncTextInput = [this]() noexcept { _pathEdit->host.SyncTextInput(_pathEdit->field); };

    switch (commandId)
    {
        case IDM_PANE_SELECTION_SELECT_ALL:
            static_cast<void>(_pathEdit->field->OnSelectAll(_pathEdit->host));
            syncTextInput();
            return true;
        case IDM_PANE_CLIPBOARD_COPY:
            static_cast<void>(_pathEdit->field->OnCopy(_pathEdit->host));
            syncTextInput();
            return true;
        case IDM_PANE_CLIPBOARD_CUT:
            static_cast<void>(_pathEdit->field->OnKeyDown(_pathEdit->host, 'X', MK_CONTROL));
            syncTextInput();
            return true;
        case IDM_PANE_CLIPBOARD_PASTE:
        case IDM_PANE_CLIPBOARD_PASTE_SHORTCUT:
            static_cast<void>(_pathEdit->field->OnKeyDown(_pathEdit->host, 'V', MK_CONTROL));
            syncTextInput();
            return true;
        default: return false;
    }
}

void NavigationView::EnsureEditSuggestWorker()
{
    if (_editSuggestThread.joinable())
    {
        return;
    }

    _editSuggestThread = std::jthread([this](std::stop_token stopToken) { EditSuggestWorker(stopToken); });
}

void NavigationView::EnsureSiblingPrefetchWorker()
{
    if (_siblingPrefetchThread.joinable())
    {
        return;
    }

    _siblingPrefetchThread = std::jthread([this](std::stop_token stopToken) { SiblingPrefetchWorker(stopToken); });
}

void NavigationView::QueueSiblingPrefetchForPath(const std::filesystem::path& displayPath)
{
    if (! _fileSystemPlugin)
    {
        return;
    }

    // /@conn: is a host-reserved prefix used by connection manager routing.
    // Prefetching parents like "/@conn:" or "/" triggers invalid enumerations for curl-backed protocols
    // (they require either an authority //host/... or a concrete /@conn:<name>/...), and can also cause
    // redundant remote calls right after Connect.
    const std::wstring_view displayText(displayPath.native());
    if (displayText.starts_with(L"/@conn:"))
    {
        return;
    }

    constexpr size_t kMaxFolders = 16u;

    const auto parts = SplitPathComponents(displayPath);
    if (parts.size() < 2u)
    {
        return;
    }

    std::vector<std::filesystem::path> folders;
    folders.reserve(std::min(parts.size(), kMaxFolders));

    for (size_t index = parts.size() - 1; index > 0; --index)
    {
        const std::filesystem::path normalized = NormalizeDirectoryPath(parts[index].fullPath);
        const std::filesystem::path parent     = normalized.parent_path();
        if (parent.empty())
        {
            continue;
        }

        const std::filesystem::path pluginParent = ToPluginPath(parent);
        if (pluginParent.empty())
        {
            continue;
        }

        const std::wstring_view pluginText(pluginParent.native());
        bool alreadyQueued = false;
        for (const auto& existing : folders)
        {
            if (EqualsNoCase(existing.native(), pluginText))
            {
                alreadyQueued = true;
                break;
            }
        }
        if (alreadyQueued)
        {
            continue;
        }

        folders.push_back(pluginParent);
        if (folders.size() >= kMaxFolders)
        {
            break;
        }
    }

    if (folders.empty())
    {
        return;
    }

    EnsureSiblingPrefetchWorker();
    const uint64_t requestId = _siblingPrefetchRequestId.fetch_add(1, std::memory_order_acq_rel) + 1u;

    {
        std::lock_guard lock(_siblingPrefetchMutex);
        SiblingPrefetchQuery query{};
        query.requestId              = requestId;
        query.fileSystem             = _fileSystemPlugin;
        query.folders                = std::move(folders);
        _siblingPrefetchPendingQuery = std::move(query);
    }

    _siblingPrefetchCv.notify_one();
}

void NavigationView::QueueSiblingPrefetchForParent(const std::filesystem::path& parentPath)
{
    if (! _fileSystemPlugin)
    {
        return;
    }

    const std::filesystem::path pluginParent = ToPluginPath(parentPath);
    if (pluginParent.empty())
    {
        return;
    }

    std::vector<std::filesystem::path> folders;
    folders.push_back(pluginParent);

    EnsureSiblingPrefetchWorker();
    const uint64_t requestId = _siblingPrefetchRequestId.fetch_add(1, std::memory_order_acq_rel) + 1u;

    {
        std::lock_guard lock(_siblingPrefetchMutex);
        SiblingPrefetchQuery query{};
        query.requestId              = requestId;
        query.fileSystem             = _fileSystemPlugin;
        query.folders                = std::move(folders);
        _siblingPrefetchPendingQuery = std::move(query);
    }

    _siblingPrefetchCv.notify_one();
}

void NavigationView::SiblingPrefetchWorker(std::stop_token stopToken)
{
    std::stop_callback stopCallback(stopToken, [this] { _siblingPrefetchCv.notify_all(); });

    for (;;)
    {
        if (stopToken.stop_requested())
        {
            return;
        }

        SiblingPrefetchQuery query{};
        {
            std::unique_lock lock(_siblingPrefetchMutex);
            _siblingPrefetchCv.wait(lock, [&] { return stopToken.stop_requested() || _siblingPrefetchPendingQuery.has_value(); });
            if (stopToken.stop_requested())
            {
                return;
            }

            query = std::move(_siblingPrefetchPendingQuery.value());
            _siblingPrefetchPendingQuery.reset();
        }

        if (! query.fileSystem)
        {
            continue;
        }

        for (const auto& folder : query.folders)
        {
            if (stopToken.stop_requested())
            {
                return;
            }

            const uint64_t latest = _siblingPrefetchRequestId.load(std::memory_order_acquire);
            if (query.requestId != latest)
            {
                break;
            }

            auto borrowed =
                DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(query.fileSystem.get(), folder, DirectoryInfoCache::BorrowMode::AllowEnumerate);
            static_cast<void>(borrowed);
        }
    }
}

void NavigationView::EditSuggestWorker(std::stop_token stopToken)
{
    std::stop_callback stopCallback(stopToken, [this] { _editSuggestCv.notify_all(); });

    for (;;)
    {
        if (stopToken.stop_requested())
        {
            return;
        }

        EditSuggestQuery query{};
        {
            std::unique_lock lock(_editSuggestMutex);
            _editSuggestCv.wait(lock, [&] { return stopToken.stop_requested() || _editSuggestPendingQuery.has_value(); });
            if (stopToken.stop_requested())
            {
                return;
            }

            query = std::move(_editSuggestPendingQuery.value());
            _editSuggestPendingQuery.reset();
        }

        std::vector<std::wstring> names;

        if (query.fileSystem)
        {
            auto borrowed = DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(
                query.fileSystem.get(), query.pluginFolder, DirectoryInfoCache::BorrowMode::AllowEnumerate);
            IFilesInformation* info = borrowed.Get();
            if (borrowed.Status() == S_OK && info)
            {
                AppendMatchingDirectoryNamesFromFilesInformation(info, query.prefix, names);
            }
        }

        const bool hasMore = SortAndTrimEditSuggestNames(names);

        std::vector<std::wstring> displayItems;
        std::vector<std::wstring> insertItems;
        BuildEditSuggestLists(query.displayFolder, names, query.directorySeparator, displayItems, insertItems);

        if (stopToken.stop_requested())
        {
            return;
        }

        PostEditSuggestResults(query.requestId,
                               query.editSessionId,
                               hasMore,
                               query.directorySeparator,
                               std::move(query.queryText),
                               std::move(query.prefix),
                               std::move(displayItems),
                               std::move(insertItems));
    }
}

void NavigationView::PostEditSuggestResults(uint64_t requestId,
                                            uint64_t editSessionId,
                                            bool hasMore,
                                            wchar_t directorySeparator,
                                            std::wstring&& queryText,
                                            std::wstring&& highlightText,
                                            std::vector<std::wstring>&& displayItems,
                                            std::vector<std::wstring>&& insertItems)
{
    if (! _hWnd)
    {
        return;
    }

    auto payload                = std::make_unique<EditSuggestResultsPayload>();
    payload->requestId          = requestId;
    payload->editSessionId      = editSessionId;
    payload->hasMore            = hasMore;
    payload->directorySeparator = directorySeparator;
    payload->queryText          = std::move(queryText);
    payload->highlightText      = std::move(highlightText);
    payload->displayItems       = std::move(displayItems);
    payload->insertItems        = std::move(insertItems);
    static_cast<void>(PostMessagePayload(_hWnd.get(), WndMsg::kEditSuggestResults, 0, std::move(payload)));
}

#ifdef ENABLE_TESTS
bool NavigationView::DebugPostCurrentEditSuggestResultForSelfTest()
{
    if (! _hWnd || ! _editMode || ! _pathEdit || ! _pathEdit->field)
    {
        return false;
    }

    auto payload           = std::make_unique<EditSuggestResultsPayload>();
    payload->requestId     = _editSuggestRequestId.load(std::memory_order_acquire);
    payload->editSessionId = _editSuggestEditSessionId;
    payload->queryText     = std::wstring(_pathEdit->field->GetText());
    payload->highlightText = L"stale";
    payload->displayItems.push_back(L"stale result");
    payload->insertItems.push_back(L"stale result");
    return PostMessagePayload(_hWnd.get(), WndMsg::kEditSuggestResults, 0, std::move(payload));
}
#endif

bool NavigationView::ValidatePath(const std::wstring& pathStr)
{
    const std::wstring_view text(pathStr);

    if (text.size() >= 6u)
    {
        constexpr std::wstring_view kConnPrefix = L"@conn:";
        if (OrdinalString::StartsWithNoCase(text, kConnPrefix))
        {
            return true;
        }
    }

    const size_t colon = text.find(L':');
    if (colon != std::wstring_view::npos && colon >= 2)
    {
        const size_t sep = text.find_first_of(L"\\/");
        if (sep == std::wstring_view::npos || sep > colon)
        {
            bool ok = true;
            for (size_t i = 0; i < colon; ++i)
            {
                if (std::iswalnum(text[i]) == 0)
                {
                    ok = false;
                    break;
                }
            }

            if (ok)
            {
                return true;
            }
        }
    }

    if (! EqualsNoCase(_pluginShortId, L"file") && LooksLikeWindowsAbsolutePath(text))
    {
        // Allow switching to the file plugin; validation will happen during plugin enumeration.
        return true;
    }

    if (! _pluginShortId.empty() && ! EqualsNoCase(_pluginShortId, L"file"))
    {
        return false;
    }

    if (! _fileSystemIo)
    {
        return false;
    }

    unsigned long attrs = 0;
    const HRESULT hr    = _fileSystemIo->GetAttributes(pathStr.c_str(), &attrs);
    if (FAILED(hr))
    {
        return false;
    }

    return (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

NavigationDxTextHost* NavigationView::ResolveEditTextHost(HWND hwnd) noexcept
{
    return const_cast<NavigationDxTextHost*>(std::as_const(*this).ResolveEditTextHost(hwnd));
}

const NavigationDxTextHost* NavigationView::ResolveEditTextHost(HWND hwnd) const noexcept
{
    if (! hwnd)
    {
        return nullptr;
    }

    const auto matches = [hwnd](const std::unique_ptr<NavigationDxTextHost>& textHost) noexcept
    {
        if (! textHost || ! textHost->hwnd)
        {
            return false;
        }

        const HWND hostHwnd  = textHost->hwnd.get();
        const HWND inputHwnd = textHost->GetTextInputHwnd();
        return hwnd == hostHwnd || hwnd == inputHwnd || IsChild(hostHwnd, hwnd) != FALSE;
    };

    if (matches(_pathEdit))
    {
        return _pathEdit.get();
    }
    if (matches(_fullPathPopupEdit))
    {
        return _fullPathPopupEdit.get();
    }
    return nullptr;
}

void NavigationView::InstallEditHostHook(NavigationDxTextHost& textHost) noexcept
{
    const HWND hwnd = textHost.hwnd.get();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    SetPropW(hwnd, kNavigationEditOwnerProp, reinterpret_cast<HANDLE>(this));
    if (! InstallWndProcHook(hwnd, NavigationView::EditWndProc, kNavigationEditOriginalWndProcProp))
    {
        RemovePropW(hwnd, kNavigationEditOwnerProp);
    }
}

void NavigationView::ReplaceEditTextSelection(NavigationDxTextHost& textHost, std::wstring_view replacement) noexcept
{
    if (! textHost.field)
    {
        return;
    }

    textHost.field->ReplaceSelectionAndNotify(replacement);
    textHost.host.SyncTextInput(textHost.field);
    if (textHost.hwnd)
    {
        InvalidateRect(textHost.hwnd.get(), nullptr, FALSE);
    }
}

void NavigationView::ShowEditValidationError(NavigationDxTextHost& textHost, const std::wstring& message) noexcept
{
    if (textHost.field)
    {
        textHost.field->SetAccessibleHelpText(message);
    }

    _editValidationMessage = message;
    UpdateEditValidationPopupWindow(textHost);
}

void NavigationView::ClearEditValidationError(NavigationDxTextHost& textHost) noexcept
{
    if (textHost.field && ! textHost.field->GetAccessibleHelpText().empty())
    {
        textHost.field->SetAccessibleHelpText({});
    }
    if (textHost.host.HasTooltip())
    {
        static_cast<void>(textHost.host.ClearTooltip());
    }

    CloseEditValidationPopup();
}

void NavigationView::RefreshActiveEditHostAfterParentPaint() noexcept
{
    if (! _editMode || ! _pathEdit || ! _pathEdit->hwnd || IsWindow(_pathEdit->hwnd.get()) == FALSE)
    {
        return;
    }

    const HWND hostHwnd                = _pathEdit->hwnd.get();
    const HWND inputHwnd               = _pathEdit->GetTextInputHwnd();
    const HWND focused                 = GetFocus();
    const HWND root                    = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    const bool focusStillBelongsToEdit = focused == _hWnd.get() || focused == hostHwnd || focused == inputHwnd || IsChild(hostHwnd, focused) != FALSE;
    const bool focusStillBelongsToRoot = root && focused && (focused == root || GetAncestor(focused, GA_ROOT) == root || IsChild(root, focused) != FALSE);

    if (! _embeddedDestinationMode)
    {
        _pathEditBlurSuppressActive      = true;
        _pathEditBlurSuppressUntilTickMs = GetTickCount64() + 2000u;
    }

    ShowWindow(hostHwnd, SW_SHOWNOACTIVATE);
    SetWindowPos(hostHwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);

    if (_pathEdit->field)
    {
        _pathEdit->host.SetFocusControl(_pathEdit->field);
        _pathEdit->host.SyncTextInput(_pathEdit->field);
        _pathEdit->host.Invalidate();
    }

    const HWND foreground     = GetForegroundWindow();
    DWORD foregroundProcessId = 0;
    if (foreground)
    {
        static_cast<void>(GetWindowThreadProcessId(foreground, &foregroundProcessId));
    }
    const bool canRestoreNullFocus = focused == nullptr && foregroundProcessId == GetCurrentProcessId();
    if (_hWnd && root && (focusStillBelongsToEdit || focusStillBelongsToRoot || canRestoreNullFocus))
    {
        SetFocus(hostHwnd);
    }

    RedrawWindow(hostHwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_ALLCHILDREN);

    if (_pathEdit->field && ! _pathEdit->field->GetAccessibleHelpText().empty())
    {
        UpdateEditValidationPopupWindow(*_pathEdit);
    }
}

void NavigationView::UpdateEditValidationPopupWindow(NavigationDxTextHost& textHost) noexcept
{
    const bool ownerEditActive = (_pathEdit.get() == &textHost && _editMode) || (_fullPathPopupEdit.get() == &textHost && _fullPathPopupEditMode);
    if (! _hWnd || ! ownerEditActive || _editValidationMessage.empty() || ! textHost.hwnd || IsWindow(textHost.hwnd.get()) == FALSE)
    {
        CloseEditValidationPopup();
        return;
    }

    if (! RegisterEditValidationPopupWndClass(_hInstance))
    {
        return;
    }

    EnsureD2DResources();
    if (! _d2dFactory || ! _dwriteFactory || ! _pathFormat)
    {
        return;
    }
    wil::com_ptr<IDWriteTextFormat> validationFormat = CreateValidationPopupTextFormat(_dwriteFactory.get(), _dpi);
    if (! validationFormat)
    {
        return;
    }
    bool usesFluentIcon                        = false;
    wchar_t iconGlyph                          = L'\0';
    wil::com_ptr<IDWriteTextFormat> iconFormat = CreateValidationPopupIconTextFormat(_dwriteFactory.get(), _dpi, usesFluentIcon, iconGlyph);

    RECT anchorRect{};
    if (GetWindowRect(textHost.hwnd.get(), &anchorRect) == FALSE)
    {
        CloseEditValidationPopup();
        return;
    }

    const int anchorWidthPx = static_cast<int>(std::max<LONG>(1, anchorRect.right - anchorRect.left));
    const int marginPx      = std::max(1, DipsToPixelsInt(kValidationPopupMarginDip, _dpi));
    const int gapPx         = std::max(0, DipsToPixelsInt(kValidationPopupGapDip, _dpi));
    const int paddingXPx    = std::max(1, DipsToPixelsInt(kValidationPopupPaddingXDip, _dpi));
    const int paddingYPx    = std::max(1, DipsToPixelsInt(kValidationPopupPaddingYDip, _dpi));
    const int iconSizePx    = iconFormat ? std::max(1, DipsToPixelsInt(kValidationPopupIconSizeDip, _dpi)) : 0;
    const int iconGapPx     = iconFormat ? std::max(0, DipsToPixelsInt(kValidationPopupIconGapDip, _dpi)) : 0;

    HMONITOR monitor = MonitorFromRect(&anchorRect, MONITOR_DEFAULTTONEAREST);
    MONITORINFO monitorInfo{};
    monitorInfo.cbSize = sizeof(monitorInfo);
    if (GetMonitorInfoW(monitor, &monitorInfo) == FALSE)
    {
        CloseEditValidationPopup();
        return;
    }

    const RECT workRect    = monitorInfo.rcWork;
    const int workWidthPx  = static_cast<int>(std::max<LONG>(1, workRect.right - workRect.left));
    const int workHeightPx = static_cast<int>(std::max<LONG>(1, workRect.bottom - workRect.top));
    const int maxWidthPx   = std::max(1, workWidthPx - (marginPx * 2));
    const int minWidthPx   = std::min(maxWidthPx, std::max(anchorWidthPx, DipsToPixelsInt(kValidationPopupMinWidthDip, _dpi)));

    int popupWidthPx = std::min(maxWidthPx, std::max(minWidthPx, DipsToPixelsInt(kValidationPopupPreferredWidthDip, _dpi)));
    popupWidthPx     = std::max(1, popupWidthPx);

    int textHeightPx = DipsToPixelsInt(16, _dpi);
    wil::com_ptr<IDWriteTextLayout> layout;
    const float textWidthDip = static_cast<float>(std::max(1, popupWidthPx - (paddingXPx * 2) - iconSizePx - iconGapPx));
    if (SUCCEEDED(_dwriteFactory->CreateTextLayout(_editValidationMessage.c_str(),
                                                   ClampTextLengthForDWrite(_editValidationMessage),
                                                   validationFormat.get(),
                                                   textWidthDip,
                                                   8192.0f,
                                                   layout.addressof())))
    {
        DWRITE_TEXT_METRICS metrics{};
        if (SUCCEEDED(layout->GetMetrics(&metrics)))
        {
            textHeightPx = std::max(1, static_cast<int>(std::ceil(metrics.height)));
        }
    }

    const int minHeightPx = DipsToPixelsInt(kValidationPopupMinHeightDip, _dpi);
    const int maxHeightPx = std::max(1, workHeightPx - (marginPx * 2));
    int popupHeightPx     = std::min(maxHeightPx, std::max(minHeightPx, std::max(textHeightPx, iconSizePx) + (paddingYPx * 2)));
    popupHeightPx         = std::max(1, popupHeightPx);

    const int workLeftPx   = static_cast<int>(workRect.left);
    const int workRightPx  = static_cast<int>(workRect.right);
    const int workTopPx    = static_cast<int>(workRect.top);
    const int workBottomPx = static_cast<int>(workRect.bottom);

    const int minX = workLeftPx + marginPx;
    const int maxX = std::max(minX, workRightPx - marginPx - popupWidthPx);
    int x          = std::clamp(static_cast<int>(anchorRect.left), minX, maxX);

    int y = static_cast<int>(anchorRect.bottom) + gapPx;
    if (y + popupHeightPx > workBottomPx - marginPx)
    {
        const int aboveY = static_cast<int>(anchorRect.top) - gapPx - popupHeightPx;
        y                = (aboveY >= workTopPx + marginPx) ? aboveY : std::max(workTopPx + marginPx, workBottomPx - marginPx - popupHeightPx);
    }

    const DWORD style   = WS_POPUP;
    const DWORD exStyle = WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
    if (! _editValidationPopup)
    {
        HWND popup = CreateWindowExW(exStyle, kValidationPopupClassName, L"", style, x, y, popupWidthPx, popupHeightPx, _hWnd.get(), nullptr, _hInstance, this);
        if (! popup)
        {
            _editValidationPopupScreenRect = {};
            return;
        }

        _editValidationPopup.reset(popup);
    }
    else
    {
        SetWindowPos(_editValidationPopup.get(), HWND_TOP, x, y, popupWidthPx, popupHeightPx, SWP_NOACTIVATE);
    }

    _editValidationPopupIconUsesFluent = iconFormat && usesFluentIcon;
    _editValidationPopupIconGlyph      = iconFormat ? iconGlyph : L'\0';
    _editValidationPopupRoundedRegion  = ApplyValidationPopupRoundedRegion(_editValidationPopup.get(), popupWidthPx, popupHeightPx, _dpi);
    ShowWindow(_editValidationPopup.get(), SW_SHOWNOACTIVATE);
    SetWindowPos(_editValidationPopup.get(), HWND_TOP, x, y, popupWidthPx, popupHeightPx, SWP_NOACTIVATE | SWP_SHOWWINDOW);
    GetWindowRect(_editValidationPopup.get(), &_editValidationPopupScreenRect);
    InvalidateRect(_editValidationPopup.get(), nullptr, FALSE);
}

void NavigationView::CloseEditValidationPopup() noexcept
{
    if (_editValidationPopup)
    {
        _editValidationPopup.reset();
    }
    _editValidationMessage.clear();
    _editValidationPopupScreenRect     = {};
    _editValidationPopupRoundedRegion  = false;
    _editValidationPopupIconUsesFluent = false;
    _editValidationPopupIconGlyph      = L'\0';
}

bool NavigationView::IsEditValidationPopupWindow(HWND hwnd) const noexcept
{
    return hwnd && _editValidationPopup && (hwnd == _editValidationPopup.get() || IsChild(_editValidationPopup.get(), hwnd) != FALSE);
}

void NavigationView::PaintEditValidationPopup(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
    static_cast<void>(hdc);

    EnsureD2DResources();
    if (! _d2dFactory || ! _dwriteFactory)
    {
        return;
    }
    wil::com_ptr<IDWriteTextFormat> validationFormat = CreateValidationPopupTextFormat(_dwriteFactory.get(), _dpi);
    if (! validationFormat)
    {
        return;
    }
    bool usesFluentIcon                        = false;
    wchar_t iconGlyph                          = L'\0';
    wil::com_ptr<IDWriteTextFormat> iconFormat = CreateValidationPopupIconTextFormat(_dwriteFactory.get(), _dpi, usesFluentIcon, iconGlyph);

    RECT clientRect{};
    if (GetClientRect(hwnd, &clientRect) == FALSE)
    {
        return;
    }

    const UINT32 width  = static_cast<UINT32>(std::max<LONG>(1, clientRect.right - clientRect.left));
    const UINT32 height = static_cast<UINT32>(std::max<LONG>(1, clientRect.bottom - clientRect.top));

    D2D1_RENDER_TARGET_PROPERTIES props                = D2D1::RenderTargetProperties();
    props.dpiX                                         = 96.0f;
    props.dpiY                                         = 96.0f;
    const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, D2D1::SizeU(width, height));

    wil::com_ptr<ID2D1HwndRenderTarget> target;
    if (FAILED(_d2dFactory->CreateHwndRenderTarget(props, hwndProps, target.addressof())))
    {
        return;
    }
    target->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);

    wil::com_ptr<ID2D1SolidColorBrush> backgroundBrush;
    wil::com_ptr<ID2D1SolidColorBrush> textBrush;
    wil::com_ptr<ID2D1SolidColorBrush> borderBrush;
    if (FAILED(target->CreateSolidColorBrush(_appTheme.folderView.warningBackground, backgroundBrush.addressof())) ||
        FAILED(target->CreateSolidColorBrush(_appTheme.folderView.warningText, textBrush.addressof())) ||
        FAILED(target->CreateSolidColorBrush(_appTheme.folderView.warningText, borderBrush.addressof())))
    {
        return;
    }

    target->BeginDraw();
    auto endDraw = wil::scope_exit([&] { static_cast<void>(target->EndDraw()); });

    const float widthDip            = static_cast<float>(width);
    const float heightDip           = static_cast<float>(height);
    const D2D1_RECT_F borderRect    = D2D1::RectF(0.5f, 0.5f, std::max(0.5f, widthDip - 0.5f), std::max(0.5f, heightDip - 0.5f));
    const float cornerRadiusDip     = static_cast<float>(std::max(1, DipsToPixelsInt(kValidationPopupCornerRadiusDip, _dpi)));
    const D2D1_ROUNDED_RECT rounded = D2D1::RoundedRect(borderRect, cornerRadiusDip, cornerRadiusDip);
    const float paddingXDip         = static_cast<float>(std::max(1, DipsToPixelsInt(kValidationPopupPaddingXDip, _dpi)));
    const float paddingYDip         = static_cast<float>(std::max(1, DipsToPixelsInt(kValidationPopupPaddingYDip, _dpi)));
    const float iconSizeDip         = iconFormat ? static_cast<float>(std::max(1, DipsToPixelsInt(kValidationPopupIconSizeDip, _dpi))) : 0.0f;
    const float iconGapDip          = iconFormat ? static_cast<float>(std::max(0, DipsToPixelsInt(kValidationPopupIconGapDip, _dpi))) : 0.0f;
    const float textLeftDip         = paddingXDip + (iconFormat ? iconSizeDip + iconGapDip : 0.0f);
    const D2D1_RECT_F textRect =
        D2D1::RectF(textLeftDip, paddingYDip, std::max(textLeftDip, widthDip - paddingXDip), std::max(paddingYDip, heightDip - paddingYDip));

    target->FillRoundedRectangle(rounded, backgroundBrush.get());
    target->DrawRoundedRectangle(rounded, borderBrush.get(), 1.0f);
    if (iconFormat && iconGlyph != L'\0')
    {
        const float iconTopDip     = std::max(paddingYDip, (heightDip - iconSizeDip) * 0.5f);
        const D2D1_RECT_F iconRect = D2D1::RectF(paddingXDip, iconTopDip, paddingXDip + iconSizeDip, iconTopDip + iconSizeDip);
        const wchar_t iconText[2]  = {iconGlyph, L'\0'};
        target->DrawTextW(iconText, 1u, iconFormat.get(), iconRect, textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_NO_SNAP, DWRITE_MEASURING_MODE_NATURAL);
    }
    constexpr auto textOptions = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
    target->DrawTextW(
        _editValidationMessage.c_str(), ClampTextLengthForDWrite(_editValidationMessage), validationFormat.get(), textRect, textBrush.get(), textOptions);
}

LRESULT NavigationView::HandleEditSubclassGetTextLength(HWND editHwnd) const noexcept
{
    const NavigationDxTextHost* textHost = ResolveEditTextHost(editHwnd);
    return (textHost && textHost->field) ? static_cast<LRESULT>(textHost->field->GetText().size()) : 0;
}

LRESULT NavigationView::HandleEditSubclassGetText(HWND editHwnd, WPARAM maxChars, LPARAM buffer) const noexcept
{
    auto* output = reinterpret_cast<wchar_t*>(buffer);
    if (! output || maxChars == 0u)
    {
        return 0;
    }

    const NavigationDxTextHost* textHost = ResolveEditTextHost(editHwnd);
    const std::wstring_view text         = (textHost && textHost->field) ? textHost->field->GetText() : std::wstring_view{};
    const size_t copyCount               = std::min(text.size(), static_cast<size_t>(maxChars - 1u));
    if (copyCount > 0u)
    {
        std::copy_n(text.data(), copyCount, output);
    }
    output[copyCount] = L'\0';
    return static_cast<LRESULT>(copyCount);
}

LRESULT NavigationView::HandleEditSubclassSetText(HWND editHwnd, LPARAM text) noexcept
{
    NavigationDxTextHost* textHost = ResolveEditTextHost(editHwnd);
    if (! textHost || ! textHost->field)
    {
        return FALSE;
    }

    const auto* input = reinterpret_cast<const wchar_t*>(text);
    textHost->field->SetTextAndNotify(input ? std::wstring(input) : std::wstring{});
    textHost->host.SyncTextInput(textHost->field);
    if (textHost->hwnd)
    {
        InvalidateRect(textHost->hwnd.get(), nullptr, FALSE);
    }
    return TRUE;
}

LRESULT NavigationView::HandleEditSubclassGetSelection(HWND editHwnd, WPARAM selectionStart, LPARAM selectionEnd) const noexcept
{
    const NavigationDxTextHost* textHost = ResolveEditTextHost(editHwnd);
    if (! textHost || ! textHost->field)
    {
        return 0;
    }

    size_t start = textHost->field->GetCaretIndex();
    size_t end   = start;
    if (const auto selection = textHost->field->GetSelectionRange(); selection.has_value())
    {
        start = selection->first;
        end   = selection->second;
    }

    if (auto* outStart = reinterpret_cast<DWORD*>(selectionStart))
    {
        *outStart = static_cast<DWORD>(std::min<size_t>(start, std::numeric_limits<DWORD>::max()));
    }
    if (auto* outEnd = reinterpret_cast<DWORD*>(selectionEnd))
    {
        *outEnd = static_cast<DWORD>(std::min<size_t>(end, std::numeric_limits<DWORD>::max()));
    }
    return static_cast<LRESULT>(MAKELONG(static_cast<WORD>(std::min<size_t>(start, std::numeric_limits<WORD>::max())),
                                         static_cast<WORD>(std::min<size_t>(end, std::numeric_limits<WORD>::max()))));
}

LRESULT NavigationView::HandleEditSubclassSetSelection(HWND editHwnd, WPARAM selectionStart, LPARAM selectionEnd) noexcept
{
    NavigationDxTextHost* textHost = ResolveEditTextHost(editHwnd);
    if (! textHost || ! textHost->field)
    {
        return 0;
    }

    const size_t textLength = textHost->field->GetText().size();
    size_t start            = selectionStart == static_cast<WPARAM>(-1) ? textLength : std::min(static_cast<size_t>(selectionStart), textLength);
    size_t end              = selectionEnd == static_cast<LPARAM>(-1) ? textLength : std::min(static_cast<size_t>(selectionEnd), textLength);
    if (selectionStart == static_cast<WPARAM>(-1))
    {
        end = start;
    }

    textHost->field->SetSelectionRange(start, end);
    textHost->host.SyncTextInput(textHost->field);
    if (textHost->hwnd)
    {
        InvalidateRect(textHost->hwnd.get(), nullptr, FALSE);
    }
    return 0;
}

LRESULT NavigationView::HandleEditSubclassReplaceSelection(HWND editHwnd, LPARAM replacement) noexcept
{
    NavigationDxTextHost* textHost = ResolveEditTextHost(editHwnd);
    if (! textHost || ! textHost->field)
    {
        return FALSE;
    }

    const auto* input = reinterpret_cast<const wchar_t*>(replacement);
    ReplaceEditTextSelection(*textHost, input ? std::wstring_view(input) : std::wstring_view{});
    return TRUE;
}

bool NavigationView::IsEditTextHostPointerInputActive(HWND hwnd) const noexcept
{
    const NavigationDxTextHost* textHost = ResolveEditTextHost(hwnd);
    if (! textHost || ! textHost->hwnd)
    {
        return false;
    }

    const bool pathEditHost     = _pathEdit.get() == textHost;
    const bool fullPathEditHost = _fullPathPopupEdit.get() == textHost;
    const bool editModeActive   = (pathEditHost && _editMode) || (fullPathEditHost && _fullPathPopupEditMode);
    if (! editModeActive)
    {
        return false;
    }

    const HWND focused = GetFocus();
    if (! focused)
    {
        return true;
    }

    const HWND hostHwnd  = textHost->hwnd.get();
    const HWND inputHwnd = textHost->GetTextInputHwnd();
    if (focused == _hWnd.get() || focused == hostHwnd || focused == inputHwnd || IsChild(hostHwnd, focused) != FALSE)
    {
        return true;
    }

    return IsEditValidationPopupWindow(focused);
}

LRESULT NavigationView::ForwardEditTextHostPointerMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (! _hWnd)
    {
        return 0;
    }

    LPARAM forwardedLParam = lp;
    switch (msg)
    {
        case WM_MOUSEMOVE:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_RBUTTONDBLCLK:
        {
            POINT pt{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
            if (MapWindowPoints(hwnd, _hWnd.get(), &pt, 1u) != 0)
            {
                forwardedLParam = MAKELPARAM(pt.x, pt.y);
            }
            break;
        }
        default: break;
    }

    return SendMessageW(_hWnd.get(), msg, wp, forwardedLParam);
}

LRESULT NavigationView::ForwardEditTextHostPointerMessageFromScreen(UINT msg, WPARAM wp, LPARAM screenPointLParam) noexcept
{
    if (! _hWnd)
    {
        return 0;
    }

    POINT pt{GET_X_LPARAM(screenPointLParam), GET_Y_LPARAM(screenPointLParam)};
    if (ScreenToClient(_hWnd.get(), &pt) == FALSE)
    {
        return 0;
    }

    return SendMessageW(_hWnd.get(), msg, wp, MAKELPARAM(pt.x, pt.y));
}

bool NavigationView::TryRetireInactiveEditTextHostForPointer(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, LRESULT& result) noexcept
{
    switch (msg)
    {
        case WM_NCHITTEST:
        case WM_MOUSEACTIVATE:
        case WM_SETCURSOR:
        case WM_MOUSEMOVE:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_RBUTTONDBLCLK: break;
        default: return false;
    }

    if (IsEditTextHostPointerInputActive(hwnd))
    {
        return false;
    }

    NavigationDxTextHost* textHost = ResolveEditTextHost(hwnd);
    if (! textHost || ! textHost->hwnd)
    {
        return false;
    }

    TraceNavigationViewMenuDiagnostics(L"navigation.edit-host-retire",
                                       L"host={:#x} msg=0x{:x} editMode={} fullPathEditMode={} focus={:#x} active={:#x} foreground={:#x}",
                                       reinterpret_cast<uintptr_t>(textHost->hwnd.get()),
                                       msg,
                                       _editMode ? 1 : 0,
                                       _fullPathPopupEditMode ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()));

    if (_pathEdit.get() == textHost && _editMode)
    {
        ExitEditMode(false, L"retire-pointer");
    }
    else if (_fullPathPopupEdit.get() == textHost && _fullPathPopupEditMode)
    {
        ExitFullPathPopupEditMode(false);
    }
    else
    {
        textHost->DeactivateForHideOrDestroy();
        ShowWindow(textHost->hwnd.get(), SW_HIDE);
    }

    switch (msg)
    {
        case WM_NCHITTEST: result = HTTRANSPARENT; return true;
        case WM_MOUSEACTIVATE:
        {
            const UINT mouseMessage = HIWORD(lp);
            if (IsForwardableEditHostMouseActivateMessage(mouseMessage))
            {
                const LPARAM screenPoint = static_cast<LPARAM>(GetMessagePos());
                TraceNavigationViewMenuDiagnostics(L"navigation.edit-host-forward-activate",
                                                   L"host={:#x} mouseMessage=0x{:x} screen=({}, {})",
                                                   reinterpret_cast<uintptr_t>(textHost->hwnd.get()),
                                                   mouseMessage,
                                                   GET_X_LPARAM(screenPoint),
                                                   GET_Y_LPARAM(screenPoint));
                static_cast<void>(ForwardEditTextHostPointerMessageFromScreen(mouseMessage, BuildForwardedMouseKeyState(mouseMessage), screenPoint));
                result = MA_NOACTIVATEANDEAT;
                return true;
            }
            result = MA_NOACTIVATE;
            return true;
        }
        case WM_SETCURSOR: result = _hWnd ? SendMessageW(_hWnd.get(), WM_SETCURSOR, reinterpret_cast<WPARAM>(_hWnd.get()), lp) : TRUE; return true;
        default: result = ForwardEditTextHostPointerMessage(hwnd, msg, wp, lp); return true;
    }
}

namespace
{
void NotifyPaneFocusChangedForEdit(const NavigationView* self) noexcept
{
    if (! self || ! self->GetHwnd())
    {
        return;
    }

    const HWND paneWindow = GetParent(self->GetHwnd());
    if (! paneWindow)
    {
        return;
    }

    PostMessageW(paneWindow, WndMsg::kPaneFocusChanged, 0, 0);
}
} // namespace

LRESULT CALLBACK NavigationView::EditWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* self                     = reinterpret_cast<NavigationView*>(GetPropW(hwnd, kNavigationEditOwnerProp));
    LRESULT inactiveEditHostResult = 0;
    if (self && self->TryRetireInactiveEditTextHostForPointer(hwnd, msg, wp, lp, inactiveEditHostResult))
    {
        return inactiveEditHostResult;
    }

    switch (msg)
    {
        case WM_MOUSEMOVE:
            if (self)
            {
                self->TraceNavigationInputState(L"edit-host.mouse-move");
            }
            TraceNavigationViewMenuDiagnostics(L"navigation.edit-host-message",
                                               L"hwnd={:#x} msg=WM_MOUSEMOVE pt=({}, {}) focus={:#x} active={:#x} foreground={:#x} capture={:#x}",
                                               reinterpret_cast<uintptr_t>(hwnd),
                                               GET_X_LPARAM(lp),
                                               GET_Y_LPARAM(lp),
                                               reinterpret_cast<uintptr_t>(GetFocus()),
                                               reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                               reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                               reinterpret_cast<uintptr_t>(GetCapture()));
            break;
        case WM_LBUTTONDOWN:
            if (self)
            {
                self->TraceNavigationInputState(L"edit-host.lbutton-down");
            }
            TraceNavigationViewMenuDiagnostics(L"navigation.edit-host-message",
                                               L"hwnd={:#x} msg=WM_LBUTTONDOWN pt=({}, {}) focus={:#x} active={:#x} foreground={:#x} capture={:#x}",
                                               reinterpret_cast<uintptr_t>(hwnd),
                                               GET_X_LPARAM(lp),
                                               GET_Y_LPARAM(lp),
                                               reinterpret_cast<uintptr_t>(GetFocus()),
                                               reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                               reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                               reinterpret_cast<uintptr_t>(GetCapture()));
            break;
        case WM_SETFOCUS:
            if (self)
            {
                self->TraceNavigationInputState(L"edit-host.set-focus");
            }
            TraceNavigationViewMenuDiagnostics(L"navigation.edit-host-message",
                                               L"hwnd={:#x} msg=WM_SETFOCUS oldFocus={:#x} focus={:#x} active={:#x}",
                                               reinterpret_cast<uintptr_t>(hwnd),
                                               reinterpret_cast<uintptr_t>(reinterpret_cast<HWND>(wp)),
                                               reinterpret_cast<uintptr_t>(GetFocus()),
                                               reinterpret_cast<uintptr_t>(GetActiveWindow()));
            break;
        case WM_KILLFOCUS:
            if (self)
            {
                self->TraceNavigationInputState(L"edit-host.kill-focus");
            }
            TraceNavigationViewMenuDiagnostics(L"navigation.edit-host-message",
                                               L"hwnd={:#x} msg=WM_KILLFOCUS newFocus={:#x} focus={:#x} active={:#x}",
                                               reinterpret_cast<uintptr_t>(hwnd),
                                               reinterpret_cast<uintptr_t>(reinterpret_cast<HWND>(wp)),
                                               reinterpret_cast<uintptr_t>(GetFocus()),
                                               reinterpret_cast<uintptr_t>(GetActiveWindow()));
            break;
        case WM_CAPTURECHANGED:
            if (self)
            {
                self->TraceNavigationInputState(L"edit-host.capture-changed");
            }
            TraceNavigationViewMenuDiagnostics(L"navigation.edit-host-message",
                                               L"hwnd={:#x} msg=WM_CAPTURECHANGED newCapture={:#x} focus={:#x} active={:#x}",
                                               reinterpret_cast<uintptr_t>(hwnd),
                                               reinterpret_cast<uintptr_t>(reinterpret_cast<HWND>(lp)),
                                               reinterpret_cast<uintptr_t>(GetFocus()),
                                               reinterpret_cast<uintptr_t>(GetActiveWindow()));
            break;
        default: break;
    }

    switch (msg)
    {
        case WM_SETFOCUS: NotifyPaneFocusChangedForEdit(self); break;
        case WM_KILLFOCUS: NotifyPaneFocusChangedForEdit(self); break;
        case WM_GETTEXTLENGTH: return self ? self->HandleEditSubclassGetTextLength(hwnd) : 0;
        case WM_GETTEXT: return self ? self->HandleEditSubclassGetText(hwnd, wp, lp) : 0;
        case WM_SETTEXT: return self ? self->HandleEditSubclassSetText(hwnd, lp) : FALSE;
        case EM_GETSEL: return self ? self->HandleEditSubclassGetSelection(hwnd, wp, lp) : 0;
        case EM_SETSEL: return self ? self->HandleEditSubclassSetSelection(hwnd, wp, lp) : 0;
        case EM_REPLACESEL: return self ? self->HandleEditSubclassReplaceSelection(hwnd, lp) : FALSE;
        case WM_NCDESTROY:
        {
            const LRESULT result = CallStoredWndProc(hwnd, kNavigationEditOriginalWndProcProp, msg, wp, lp);
            RestoreWndProcHook(hwnd, kNavigationEditOriginalWndProcProp, kNavigationEditOwnerProp);
            return result;
        }
    }

    return CallStoredWndProc(hwnd, kNavigationEditOriginalWndProcProp, msg, wp, lp);
}
