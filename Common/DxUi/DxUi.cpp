#include "DxUi.Internal.h"

#include <algorithm>
#include <cwctype>

#include "Helpers.h"

namespace RedSalamander::DxUi
{
namespace
{
[[nodiscard]] wchar_t NormalizeMnemonicChar(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch)));
}
} // namespace

void Control::SetBounds(const D2D1_RECT_F& bounds) noexcept
{
    if (_bounds.left != bounds.left || _bounds.top != bounds.top || _bounds.right != bounds.right || _bounds.bottom != bounds.bottom)
    {
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

void Control::SetVisible(bool visible) noexcept
{
    if (_visible != visible)
    {
        _visible = visible;
        RequestInvalidate();
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
        RequestInvalidate();
    }
}

bool Control::IsEnabled() const noexcept
{
    return _enabled;
}

void Control::SetFocusable(bool focusable) noexcept
{
    _focusable = focusable;
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
    }
}

std::wstring_view Control::GetAccessibleName() const noexcept
{
    return _accessibleName;
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
    RequestInvalidate();
}

void Control::OnDensityChanged() noexcept
{
    RequestInvalidate();
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

bool Control::SupportsTextInputBridge() const noexcept
{
    return false;
}

std::optional<D2D1_RECT_F> Control::GetTextInputBridgeViewportRect() const noexcept
{
    return std::nullopt;
}

std::optional<D2D1_RECT_F> Control::GetTextInputBridgeCaretRect(const WindowHost& /*host*/, size_t /*controlTextIndex*/) const noexcept
{
    return std::nullopt;
}

bool Control::ExportTextInputBridgeState(TextInputBridgeState& /*outState*/) const
{
    return false;
}

bool Control::ImportTextInputBridgeState(WindowHost& /*host*/, const TextInputBridgeState& /*state*/, bool /*notifyChange*/)
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
