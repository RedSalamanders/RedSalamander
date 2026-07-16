#include "DxUi.Internal.h"

#include <algorithm>
#include <cstdint>
#include <cwctype>
#include <limits>

#include "Helpers.h"

namespace RedSalamander::DxUi
{
namespace
{
[[nodiscard]] wchar_t NormalizeMnemonicChar(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch)));
}

[[nodiscard]] bool ContinueModalLoopByDefault(void*) noexcept
{
    return true;
}
} // namespace

DxUiModalLoopResult RunDxUiModalLoop(HWND hwnd, const DxUiModalLoopOptions& options) noexcept
{
    const DxUiModalLoopContinueCallback shouldContinue = options.shouldContinue ? options.shouldContinue : ContinueModalLoopByDefault;
    const std::wstring_view diagnosticName             = options.diagnosticName.empty() ? std::wstring_view(L"modal") : options.diagnosticName;

    MSG msg{};
    while (shouldContinue(options.context))
    {
        const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
        if (getMessageResult == -1)
        {
            const DWORD lastError = GetLastError();
            Debug::Warning(L"DxUi::RunDxUiModalLoop: GetMessageW failed for '{0}' (hwnd=0x{1:X}, lastError={2})",
                           diagnosticName,
                           reinterpret_cast<uintptr_t>(hwnd),
                           lastError);
            SetLastError(lastError);
            return DxUiModalLoopResult::GetMessageFailed;
        }

        if (getMessageResult == 0)
        {
            if (options.onQuit)
            {
                options.onQuit(msg.wParam, options.context);
            }
            PostQuitMessage(static_cast<int>(msg.wParam));
            return DxUiModalLoopResult::Quit;
        }

        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    return DxUiModalLoopResult::Completed;
}

bool CaptureBackdropScreenRegion(const RECT& screenRect, WindowHostBitmapCapture& outCapture, std::wstring_view componentName) noexcept
{
    outCapture = {};

    const LONG widthPx  = screenRect.right - screenRect.left;
    const LONG heightPx = screenRect.bottom - screenRect.top;
    if (widthPx <= 0 || heightPx <= 0)
    {
        return false;
    }

    const uint64_t pixelCount = static_cast<uint64_t>(widthPx) * static_cast<uint64_t>(heightPx);
    if (pixelCount > static_cast<uint64_t>((std::numeric_limits<size_t>::max)() / 4u))
    {
        Debug::Warning(L"DxUi::{}: popup backdrop capture is too large (widthPx={} heightPx={})", componentName, widthPx, heightPx);
        return false;
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = widthPx;
    bmi.bmiHeader.biHeight      = -heightPx;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hdc_window screenDc{GetDC(nullptr)};
    if (! screenDc)
    {
        Debug::Warning(L"DxUi::{}: unable to acquire screen DC for popup backdrop capture", componentName);
        return false;
    }

    wil::unique_hdc memoryDc{CreateCompatibleDC(screenDc.get())};
    if (! memoryDc)
    {
        Debug::Warning(L"DxUi::{}: unable to create memory DC for popup backdrop capture", componentName);
        return false;
    }

    wil::unique_hbitmap bitmap{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        Debug::Warning(L"DxUi::{}: unable to create DIB section for popup backdrop capture", componentName);
        return false;
    }

    [[maybe_unused]] const auto oldBitmap = wil::SelectObject(memoryDc.get(), bitmap.get());
    if (BitBlt(memoryDc.get(), 0, 0, widthPx, heightPx, screenDc.get(), screenRect.left, screenRect.top, SRCCOPY | CAPTUREBLT) == FALSE)
    {
        Debug::Warning(L"DxUi::{}: BitBlt failed for popup backdrop capture (lastError={})", componentName, GetLastError());
        return false;
    }

    outCapture.widthPx  = static_cast<UINT>(widthPx);
    outCapture.heightPx = static_cast<UINT>(heightPx);
    outCapture.bgraPixels.resize(static_cast<size_t>(pixelCount) * 4u);

    const auto* const sourceBytes = static_cast<const uint8_t*>(bits);
    std::copy_n(sourceBytes, outCapture.bgraPixels.size(), outCapture.bgraPixels.data());
    for (size_t offset = 3u; offset < outCapture.bgraPixels.size(); offset += 4u)
    {
        outCapture.bgraPixels[offset] = 0xFFu;
    }

    return true;
}

void Control::SetBounds(const D2D1_RECT_F& bounds) noexcept
{
    if (_bounds.left != bounds.left || _bounds.top != bounds.top || _bounds.right != bounds.right || _bounds.bottom != bounds.bottom)
    {
        if (IsDxUiRenderStageActiveForDebug())
        {
            EmitDxUiRenderMutationBlockedForDebug();
            return;
        }

        _bounds = bounds;
        OnBoundsChanged();
        RequestInvalidate();
    }
}

D2D1_RECT_F Control::GetBounds() const noexcept
{
    return _bounds;
}

D2D1_RECT_F Control::GetHitBounds() const noexcept
{
    return _bounds;
}

std::optional<D2D1_RECT_F> Control::TryGetTextInputViewportRect() const noexcept
{
    return GetTextInputViewportRect();
}

std::optional<D2D1_RECT_F> Control::TryGetTextInputCaretRect(const WindowHost& host, size_t controlTextIndex) const noexcept
{
    return GetTextInputCaretRect(host, controlTextIndex);
}

std::optional<std::vector<D2D1_RECT_F>> Control::TryGetTextInputRangeRects(const WindowHost& host,
                                                                           size_t controlTextStartIndex,
                                                                           size_t controlTextEndIndex) const
{
    return GetTextInputRangeRects(host, controlTextStartIndex, controlTextEndIndex);
}

std::optional<size_t> Control::TryHitTestTextInputPoint(const WindowHost& host, D2D1_POINT_2F point) const noexcept
{
    return HitTestTextInputPoint(host, point);
}

void Control::SetVisible(bool visible) noexcept
{
    if (_visible != visible)
    {
        _visible = visible;
        RequestInvalidate();
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

bool Control::IsVisible() const noexcept
{
    return _visible;
}

void Control::SetEnabled(bool enabled) noexcept
{
    if (_enabled != enabled)
    {
        _enabled = enabled;
        OnEnabledChanged(enabled);
        RequestInvalidate();
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

bool Control::IsEnabled() const noexcept
{
    return _enabled;
}

void Control::SetFocusable(bool focusable) noexcept
{
    if (_focusable != focusable)
    {
        _focusable = focusable;
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

bool Control::IsFocusable() const noexcept
{
    return _focusable && _enabled && _visible;
}

bool Control::HasFocus() const noexcept
{
    return _hasFocus;
}

bool Control::IsHovered() const noexcept
{
    return _hovered;
}

void Control::PaintOverlay(WindowHost& /*host*/) const
{
}

bool Control::Tick(WindowHost& /*host*/, uint64_t /*nowTickMs*/)
{
    return false;
}

bool Control::OnMouseMove(WindowHost& /*host*/, D2D1_POINT_2F /*point*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnMouseLeave(WindowHost& /*host*/)
{
    return false;
}

bool Control::OnMouseDown(WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool /*rightButton*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    return OnMouseDown(host, point, rightButton, modifiers);
}

bool Control::OnMouseUp(WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool /*rightButton*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnMouseWheel(WindowHost& /*host*/, D2D1_POINT_2F /*point*/, float /*wheelDelta*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnKeyDown(WindowHost& /*host*/, UINT /*virtualKey*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnKeyUp(WindowHost& /*host*/, UINT /*virtualKey*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnChar(WindowHost& /*host*/, wchar_t /*ch*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip)
{
    if (! IsEnabled() || ! IsVisible() || ! _onContextMenu)
    {
        return false;
    }

    _onContextMenu(ResolveContextMenuAnchor(host, keyboardInvocation, pointDip), keyboardInvocation);
    return true;
}

bool Control::OnCopy(WindowHost& /*host*/)
{
    return false;
}

bool Control::OnSelectAll(WindowHost& /*host*/)
{
    return false;
}

bool Control::OnMnemonic(WindowHost& host)
{
    if (! IsEnabled() || ! IsVisible())
    {
        return false;
    }

    if (IsFocusable())
    {
        if (const HWND hwnd = host.GetHwnd())
        {
            SetFocus(hwnd);
        }
        host.SetFocusControl(this);
        return true;
    }

    return false;
}

size_t Control::GetLogicalChildCount() const noexcept
{
    return 0u;
}

Control* Control::GetLogicalChild(size_t /*index*/) noexcept
{
    return nullptr;
}

const Control* Control::GetLogicalChild(size_t /*index*/) const noexcept
{
    return nullptr;
}

WindowHost* Control::GetHost() const noexcept
{
    return _host;
}

void Control::SetMnemonic(wchar_t mnemonic) noexcept
{
    _mnemonic = NormalizeMnemonicChar(mnemonic);
}

wchar_t Control::GetMnemonic() const noexcept
{
    return _mnemonic;
}

void Control::SetFlowDirection(FlowDirection direction) noexcept
{
    if (_explicitFlowDirection.has_value() && _explicitFlowDirection.value() == direction)
    {
        return;
    }

    _explicitFlowDirection = direction;
    OnFlowDirectionChanged();
    RequestInvalidate();
}

void Control::ClearFlowDirection() noexcept
{
    if (! _explicitFlowDirection.has_value())
    {
        return;
    }

    _explicitFlowDirection.reset();
    OnFlowDirectionChanged();
    RequestInvalidate();
}

bool Control::HasExplicitFlowDirection() const noexcept
{
    return _explicitFlowDirection.has_value();
}

FlowDirection Control::GetFlowDirection() const noexcept
{
    if (_explicitFlowDirection.has_value())
    {
        return _explicitFlowDirection.value();
    }

    return _parent ? _parent->GetFlowDirection() : FlowDirection::LeftToRight;
}

bool Control::IsRightToLeft() const noexcept
{
    return GetFlowDirection() == FlowDirection::RightToLeft;
}

void Control::SetDensity(Density density) noexcept
{
    if (_explicitDensity.has_value() && _explicitDensity.value() == density)
    {
        return;
    }

    _explicitDensity = density;
    OnDensityChanged();
    RequestInvalidate();
}

void Control::ClearDensity() noexcept
{
    if (! _explicitDensity.has_value())
    {
        return;
    }

    _explicitDensity.reset();
    OnDensityChanged();
    RequestInvalidate();
}

bool Control::HasExplicitDensity() const noexcept
{
    return _explicitDensity.has_value();
}

Density Control::GetDensity() const noexcept
{
    if (_explicitDensity.has_value())
    {
        return _explicitDensity.value();
    }

    if (_parent)
    {
        return _parent->GetDensity();
    }

    return _host ? _host->GetTheme().density : Density::Standard;
}

bool Control::IsCompactDensity() const noexcept
{
    return GetDensity() == Density::Compact;
}

void Control::SetConnectedAnimationKey(std::wstring key)
{
    if (_connectedAnimationKey != key)
    {
        _connectedAnimationKey = std::move(key);
        RequestInvalidate();
    }
}

std::wstring_view Control::GetConnectedAnimationKey() const noexcept
{
    return _connectedAnimationKey;
}

void Control::SetAccessibleName(std::wstring name)
{
    if (_accessibleName != name)
    {
        _accessibleName = std::move(name);
        RequestInvalidate();
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

std::wstring_view Control::GetAccessibleName() const noexcept
{
    return _accessibleName;
}

void Control::SetAccessibleHelpText(std::wstring helpText)
{
    if (_accessibleHelpText != helpText)
    {
        _accessibleHelpText = std::move(helpText);
        RequestInvalidate();
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

std::wstring_view Control::GetAccessibleHelpText() const noexcept
{
    return _accessibleHelpText;
}

void Control::SetOnContextMenu(std::function<void(POINT screenPoint, bool keyboardInvocation)> onContextMenu)
{
    _onContextMenu = std::move(onContextMenu);
}

Control* Control::HitTest(D2D1_POINT_2F point)
{
    return (_visible && _enabled && PointInRect(GetHitBounds(), point)) ? this : nullptr;
}

const Control* Control::HitTest(D2D1_POINT_2F point) const
{
    return (_visible && _enabled && PointInRect(GetHitBounds(), point)) ? this : nullptr;
}

Control* Control::HitTestOverlay(D2D1_POINT_2F /*point*/)
{
    return nullptr;
}

const Control* Control::HitTestOverlay(D2D1_POINT_2F /*point*/) const
{
    return nullptr;
}

bool Control::DismissOverlayOnPointerDown(WindowHost& host, D2D1_POINT_2F point)
{
    if (! IsVisible() || ! IsEnabled())
    {
        return false;
    }

    for (size_t childIndex = GetLogicalChildCount(); childIndex > 0u; --childIndex)
    {
        Control* const child = GetLogicalChild(childIndex - 1u);
        if (child && child->DismissOverlayOnPointerDown(host, point))
        {
            return true;
        }
    }

    return false;
}

POINT Control::ResolveContextMenuAnchor(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip) const noexcept
{
    if (! keyboardInvocation)
    {
        return host.DipPointToScreenPoint(pointDip);
    }

    const D2D1_RECT_F bounds = GetHitBounds();
    const float anchorX      = (std::min)(bounds.right, bounds.left + 16.0f);
    const float anchorY      = bounds.top + ((bounds.bottom - bounds.top) * 0.5f);
    return host.DipPointToScreenPoint(D2D1::Point2F(anchorX, anchorY));
}

WindowHostCursorKind Control::ResolveCursorKind(WindowHost& /*host*/, D2D1_POINT_2F /*pointDip*/) const noexcept
{
    return WindowHostCursorKind::Default;
}

void Control::Invalidate(WindowHost& host) const
{
    host.Invalidate();
}

void Control::RequestInvalidate() const noexcept
{
    if (_host)
    {
        _host->Invalidate();
    }
}

void Control::PropagateHost(WindowHost* host) noexcept
{
    _host = host;
}

void Control::SetParent(Panel* parent) noexcept
{
    _parent = parent;
    PropagateHost(parent ? parent->_host : nullptr);
}

Panel* Control::GetParent() const noexcept
{
    return _parent;
}

void Control::OnBoundsChanged() noexcept
{
}

void Control::OnFlowDirectionChanged() noexcept
{
    if (_host && _host->GetFocusControl() == this && SupportsTextInput())
    {
        _host->SyncTextInput(this);
    }
    RequestInvalidate();
}

void Control::OnDensityChanged() noexcept
{
    if (_host && _host->GetFocusControl() == this && SupportsTextInput())
    {
        _host->SyncTextInput(this);
    }
    RequestInvalidate();
}

void Control::OnEnabledChanged(bool /*enabled*/) noexcept
{
}

void Control::OnHostDpiChanged(WindowHost& /*host*/) noexcept
{
}

void Control::OnFocusChanged(WindowHost& /*host*/, bool focused)
{
    _hasFocus = focused;
}

void Control::OnHoverChanged(WindowHost& /*host*/, bool hovered)
{
    _hovered = hovered;
}

void Control::OnCaptureLost(WindowHost& /*host*/)
{
}

bool Control::SupportsTextInput() const noexcept
{
    return false;
}

std::optional<D2D1_RECT_F> Control::GetTextInputViewportRect() const noexcept
{
    return std::nullopt;
}

std::optional<D2D1_RECT_F> Control::GetTextInputCaretRect(const WindowHost& /*host*/, size_t /*controlTextIndex*/) const noexcept
{
    return std::nullopt;
}

std::optional<std::vector<D2D1_RECT_F>> Control::GetTextInputRangeRects(const WindowHost& /*host*/,
                                                                        size_t /*controlTextStartIndex*/,
                                                                        size_t /*controlTextEndIndex*/) const
{
    return std::nullopt;
}

std::optional<size_t> Control::HitTestTextInputPoint(const WindowHost& /*host*/, D2D1_POINT_2F /*point*/) const noexcept
{
    return std::nullopt;
}

bool Control::ExportTextInputState(TextInputState& /*outState*/) const
{
    return false;
}

bool Control::ImportTextInputState(WindowHost& /*host*/, const TextInputState& /*state*/, bool /*notifyChange*/)
{
    return false;
}

std::vector<std::unique_ptr<Control>>& Panel::AccessChildren() noexcept
{
    return _children;
}

const std::vector<std::unique_ptr<Control>>& Panel::AccessChildren() const noexcept
{
    return _children;
}

} // namespace RedSalamander::DxUi
