#include "DxUiTestHelpers.h"

#include <fstream>

namespace
{
using RedSalamander::DxUi::WindowHostBitmapCapture;

class StripedBackdropControl final : public RedSalamander::DxUi::Control
{
public:
    explicit StripedBackdropControl(LONG stripeWidthPx = 12) noexcept : _stripeWidthPx((std::max)(1l, stripeWidthPx))
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& host) const override
    {
        auto* const dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        const D2D1_RECT_F bounds      = host.GetClientBoundsDip();
        const LONG left               = static_cast<LONG>(std::floor(bounds.left));
        const LONG right              = static_cast<LONG>(std::ceil(bounds.right));
        const D2D1_COLOR_F evenStripe = D2D1::ColorF(0.94f, 0.97f, 1.00f, 1.0f);
        const D2D1_COLOR_F oddStripe  = D2D1::ColorF(0.03f, 0.16f, 0.34f, 1.0f);

        for (LONG x = left; x < right; ++x)
        {
            const LONG stripeIndex = (x - left) / _stripeWidthPx;
            if (auto* const brush = host.GetSolidBrush((stripeIndex & 1) == 0 ? evenStripe : oddStripe))
            {
                dc->FillRectangle(D2D1::RectF(static_cast<float>(x), bounds.top, static_cast<float>(x + 1), bounds.bottom), brush);
            }
        }
    }

private:
    LONG _stripeWidthPx = 12;
};

class OverlayPaintProbeControl final : public RedSalamander::DxUi::Control
{
public:
    void Paint(RedSalamander::DxUi::WindowHost&) const override
    {
    }

    void PaintOverlay(RedSalamander::DxUi::WindowHost& host) const override
    {
        auto* const dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        auto* const brush = host.GetSolidBrush(D2D1::ColorF(1.0f, 0.0f, 0.0f, 1.0f));
        if (brush)
        {
            dc->FillRectangle(GetBounds(), brush);
        }
    }
};

class PointerSinkControl final : public RedSalamander::DxUi::Control
{
public:
    void Paint(RedSalamander::DxUi::WindowHost&) const override
    {
    }

    bool OnMouseDown(RedSalamander::DxUi::WindowHost&, D2D1_POINT_2F, bool rightButton, UINT) override
    {
        if (rightButton)
        {
            return false;
        }

        ++_mouseDownCount;
        return true;
    }

    [[nodiscard]] size_t MouseDownCount() const noexcept
    {
        return _mouseDownCount;
    }

private:
    size_t _mouseDownCount = 0u;
};

WindowHostBitmapCapture CaptureAttachedHostWindowBitmapForComboSuite(AttachedHostWindow& window, const char* context)
{
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

    WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), context);
    return capture;
}

[[nodiscard]] uint32_t GetCapturePixelBgra(const WindowHostBitmapCapture& capture, UINT xPx, UINT yPx) noexcept
{
    if (xPx >= capture.widthPx || yPx >= capture.heightPx)
    {
        return 0u;
    }

    const size_t base = (static_cast<size_t>(yPx) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(xPx)) * 4u;
    if ((base + 3u) >= capture.bgraPixels.size())
    {
        return 0u;
    }

    return static_cast<uint32_t>(capture.bgraPixels[base + 0u]) | (static_cast<uint32_t>(capture.bgraPixels[base + 1u]) << 8u) |
           (static_cast<uint32_t>(capture.bgraPixels[base + 2u]) << 16u) | (static_cast<uint32_t>(capture.bgraPixels[base + 3u]) << 24u);
}

[[nodiscard]] uint32_t ComputeRgbDelta(uint32_t lhsBgra, uint32_t rhsBgra) noexcept
{
    const auto channel = [](uint32_t color, UINT shift) noexcept { return static_cast<int>((color >> shift) & 0xFFu); };

    return static_cast<uint32_t>(std::abs(channel(lhsBgra, 0u) - channel(rhsBgra, 0u)) + std::abs(channel(lhsBgra, 8u) - channel(rhsBgra, 8u)) +
                                 std::abs(channel(lhsBgra, 16u) - channel(rhsBgra, 16u)));
}

[[nodiscard]] size_t CountSubstring(std::string_view text, std::string_view needle) noexcept
{
    if (needle.empty())
    {
        return 0u;
    }

    size_t count  = 0u;
    size_t offset = 0u;
    while ((offset = text.find(needle, offset)) != std::string_view::npos)
    {
        ++count;
        offset += needle.size();
    }
    return count;
}

[[nodiscard]] bool PostHostLeftClickForComboSuite(HWND hwnd, POINT screenPoint) noexcept
{
    POINT clientPoint = screenPoint;
    if (! hwnd || ScreenToClient(hwnd, &clientPoint) == FALSE)
    {
        return false;
    }

    const LPARAM lParam = MAKELPARAM(static_cast<WORD>(static_cast<SHORT>(clientPoint.x)), static_cast<WORD>(static_cast<SHORT>(clientPoint.y)));
    return PostMessageW(hwnd, WM_MOUSEMOVE, 0, lParam) != FALSE && PostMessageW(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, lParam) != FALSE &&
           PostMessageW(hwnd, WM_LBUTTONUP, 0, lParam) != FALSE;
}

template <typename Predicate> [[nodiscard]] bool PumpComboSuiteMessagesUntil(AttachedHostWindow& window, DWORD timeoutMs, Predicate&& predicate)
{
    const ULONGLONG deadline = GetTickCount64() + timeoutMs;
    while (GetTickCount64() <= deadline)
    {
        window.PumpMessages();
        if (predicate())
        {
            return true;
        }
        Sleep(10);
    }
    window.PumpMessages();
    return predicate();
}

[[nodiscard]] size_t CountWarmSaturatedPixels(const WindowHostBitmapCapture& capture) noexcept
{
    size_t count = 0u;
    for (size_t pixelIndex = 0u; pixelIndex < static_cast<size_t>(capture.widthPx) * static_cast<size_t>(capture.heightPx); ++pixelIndex)
    {
        const size_t base = pixelIndex * 4u;
        if ((base + 3u) >= capture.bgraPixels.size())
        {
            break;
        }

        const uint8_t b = capture.bgraPixels[base + 0u];
        const uint8_t g = capture.bgraPixels[base + 1u];
        const uint8_t r = capture.bgraPixels[base + 2u];
        const uint8_t a = capture.bgraPixels[base + 3u];
        if (a >= 240u && r >= 170u && g >= 40u && g <= 220u && b <= 120u && r >= static_cast<uint8_t>((std::min)(255, g + 20)) &&
            g >= static_cast<uint8_t>((std::min)(255, b + 15)))
        {
            ++count;
        }
    }
    return count;
}

void EmitColorGlyphPixelCountForTest(std::wstring_view detail, const WindowHostBitmapCapture& capture, size_t warmPixelCount) noexcept
{
    if (! Debug::Perf::IsCaptureEnabled())
    {
        return;
    }

    const size_t pixelCount = static_cast<size_t>(capture.widthPx) * static_cast<size_t>(capture.heightPx);
    Debug::Perf::Emit(L"dxui.textinput.color_glyph_pixel_count", detail, 0, warmPixelCount, pixelCount, S_OK);
}

void TestComboRightClickInvokesContextMenuWithoutOpeningPopup()
{
    using namespace RedSalamander::DxUi;

    RecordingContextMenuInvocation contextMenu;
    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});
    combo->SetSelectedIndex(0u);
    combo->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 40.0f));
    host.SetFocusControl(combo);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_RBUTTONDOWN, 0, MAKELPARAM(172, 14), handled));
    Require(handled, "non-editable combo right-click is handled");
    Require(contextMenu.count == 1u, "non-editable combo right-click invokes one context menu");
    Require(! contextMenu.lastKeyboardInvocation, "non-editable combo right-click reports pointer invocation");
    Require(contextMenu.lastPoint.x == 172 && contextMenu.lastPoint.y == 14,
            "non-editable combo right-click uses the hit point as its client anchor on the detached host");
    Require(host.GetFocusControl() == combo, "non-editable combo right-click keeps focus on the combo");
    Require(! host.HasActiveTextInput(), "non-editable combo right-click does not activate text input");
    Require(! combo->DebugIsPopupOpen(), "non-editable combo right-click does not open the popup");
}

void TestComboBoxClosesOnFocusLoss()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* combo  = root->AddChild<ComboBox>();
    auto* button = root->AddChild<Button>(L"Other");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    button->SetBounds(D2D1::RectF(0.0f, 36.0f, 120.0f, 64.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    host.SetRoot(std::move(root));
    host.SetFocusControl(combo);

    Require(combo->OnKeyDown(host, VK_RETURN, 0), "combo enter opens popup");
    const auto openBounds = combo->GetHitBounds();
    Require(openBounds.bottom > combo->GetBounds().bottom, "open combo expands hit bounds");

    host.SetFocusControl(button);
    const auto closedBounds = combo->GetHitBounds();
    Require(closedBounds.bottom == combo->GetBounds().bottom, "combo popup closes on focus loss");
}

void TestComboBoxSecondClickTogglesPopupClosed()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(172, 12), handled));
    Require(handled, "first combo click handled");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(172, 12), handled));
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "first combo click opens popup");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(172, 12), handled));
    Require(handled, "second combo click handled");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(172, 12), handled));
    Require(combo->GetHitBounds().bottom == combo->GetBounds().bottom, "second combo click closes popup");
}

// Drives the attached-host combo popup entirely through delivered host messages
// (PostHost* helpers). The ComboBox dropdown is an in-host overlay rather than a
// separate top-level window, so production routes hover/click/light-dismiss purely
// from the delivered message point -- no live cursor injection is required, which
// keeps this test from warping the interactive user's cursor or stealing foreground.
void TestAttachedComboBoxSentMouseHoverAndOutsideDismiss()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    Require(SetWindowPos(window.Hwnd(), nullptr, 120, 120, 360, 260, SWP_NOZORDER | SWP_NOACTIVATE) != FALSE,
            "attached combo pointer window can be placed on-screen");
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);

    auto root         = std::make_unique<Panel>();
    auto* combo       = root->AddChild<ComboBox>();
    auto* pointerSink = root->AddChild<PointerSinkControl>();
    combo->SetBounds(D2D1::RectF(24.0f, 24.0f, 216.0f, 56.0f));
    combo->SetItems({
        ComboBox::Item{L"one", L"One"},
        ComboBox::Item{L"two", L"Two"},
        ComboBox::Item{L"three", L"Three"},
    });
    pointerSink->SetBounds(D2D1::RectF(244.0f, 24.0f, 336.0f, 88.0f));

    window.Host().SetRoot(std::move(root));
    static_cast<Panel*>(window.Host().GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 260.0f));
    window.PumpMessages();

    const POINT openPoint = window.Host().DipPointToScreenPoint(D2D1::Point2F(208.0f, 40.0f));
    Require(PostHostLeftClickForComboSuite(window.Hwnd(), openPoint), "combo click sends a host button sequence");
    const bool popupOpened = PumpComboSuiteMessagesUntil(window, 500u, [&]() noexcept { return combo->DebugIsPopupOpen(); });
    Require(popupOpened, "combo click opens the popup through the attached host message pump");

    const D2D1_RECT_F secondPopupItemRect = combo->DebugGetPopupItemRect(1u, &window.Host());
    RequireRectHasArea(secondPopupItemRect, "combo popup exposes the second popup item geometry");
    const D2D1_POINT_2F secondPopupItemCenter =
        D2D1::Point2F((secondPopupItemRect.left + secondPopupItemRect.right) * 0.5f, (secondPopupItemRect.top + secondPopupItemRect.bottom) * 0.5f);
    Require(combo->OnMouseMove(window.Host(), secondPopupItemCenter, 0), "combo popup hover move over the second item is handled");
    window.PumpMessages();
    Require(combo->DebugGetHoveredPopupIndex().has_value() && combo->DebugGetHoveredPopupIndex().value() == 1u,
            "combo popup hover updates the hovered index to the second item from the delivered point");

    const POINT outsidePoint = window.Host().DipPointToScreenPoint(D2D1::Point2F(288.0f, 56.0f));
    Require(PostHostLeftClickForComboSuite(window.Hwnd(), outsidePoint), "outside combo click sends a host button sequence");
    const bool popupClosed = PumpComboSuiteMessagesUntil(window, 500u, [&]() noexcept { return ! combo->DebugIsPopupOpen(); });
    Require(popupClosed, "outside click light-dismisses the combo popup even when the click lands on a handled non-focusable control");
    Require(pointerSink->MouseDownCount() == 1u, "combo outside click reaches the underlying sink after dismissing the popup");
}

void TestComboBoxPopupOverlayBlocksUnderlyingHover()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState underlyingState;
    auto root        = std::make_unique<Panel>();
    auto* combo      = root->AddChild<ComboBox>();
    auto* underlying = root->AddChild<TrackingControl>(underlyingState);
    combo->SetBounds(D2D1::RectF(12.0f, 0.0f, 192.0f, 28.0f));
    underlying->SetBounds(D2D1::RectF(0.0f, 30.0f, 220.0f, 120.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}, ComboBox::Item{L"three", L"Three"}});

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 180.0f));
    host.SetFocusControl(combo);
    Require(combo->OnKeyDown(host, VK_RETURN, 0), "combo enter opens popup for overlay hover test");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(48, 50), handled));
    Require(handled, "popup hover move handled");
    Require(combo->IsHovered(), "open combo owns hovered state above overlapping sibling content");
    Require(underlyingState.hoverEnterCount == 0u, "underlying control does not receive hover while combo popup is open above it");
    Require(underlyingState.mouseMoveCount == 0u, "underlying control does not receive popup hover mouse moves");
}

void TestComboBoxPopupOverlayBlocksUnderlyingClick()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState underlyingState;
    auto root        = std::make_unique<Panel>();
    auto* combo      = root->AddChild<ComboBox>();
    auto* underlying = root->AddChild<TrackingControl>(underlyingState);
    combo->SetBounds(D2D1::RectF(12.0f, 0.0f, 192.0f, 28.0f));
    underlying->SetBounds(D2D1::RectF(0.0f, 30.0f, 220.0f, 120.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}, ComboBox::Item{L"three", L"Three"}});

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 180.0f));
    host.SetFocusControl(combo);
    Require(combo->OnKeyDown(host, VK_RETURN, 0), "combo enter opens popup for overlay click test");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(48, 78), handled));
    Require(handled, "popup click handled");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(48, 78), handled));

    Require(combo->GetSelectedIndex().has_value(), "popup click selects a combo item");
    Require(combo->GetSelectedIndex().value() == 1u, "popup click selects the visible item above the overlapping sibling");
    Require(host.GetFocusControl() == combo, "popup click keeps focus on the combo rather than the overlapping sibling");
    Require(underlyingState.mouseDownCount == 0u, "underlying control does not receive mouse down through the popup");
    Require(underlyingState.mouseUpCount == 0u, "underlying control does not receive mouse up through the popup");
    Require(underlyingState.focusGainCount == 0u, "underlying control does not steal focus through the popup");
}

void TestComboBoxPopupInsideScrolledPanelUsesViewportCoordinates()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(420, 220), handled));
    Require(handled, "host size update handled for scrolled combo popup test");

    auto root                = std::make_unique<Panel>();
    auto* scroll             = root->AddChild<ScrollPanel>();
    auto* wrapper            = scroll->AddChild<Panel>();
    auto* combo              = wrapper->AddChild<ComboBox>();
    const float scrollOffset = 220.0f;
    scroll->SetBounds(D2D1::RectF(0.0f, 0.0f, 420.0f, 220.0f));
    scroll->SetContentHeight(720.0f);
    scroll->SetScrollOffset(scrollOffset);
    wrapper->SetBounds(D2D1::RectF(0.0f, 0.0f, 420.0f, 720.0f));
    combo->SetBounds(D2D1::RectF(260.0f, 300.0f, 404.0f, 328.0f));
    combo->SetItems({
        ComboBox::Item{L"custom", L"Custom"},
        ComboBox::Item{L"errors", L"Errors only"},
        ComboBox::Item{L"warnings", L"Errors and warnings"},
        ComboBox::Item{L"all", L"All types"},
    });
    combo->SetSelectedIndex(0u);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 420.0f, 220.0f));

    const D2D1_RECT_F comboBounds = combo->GetBounds();
    const LONG openX              = static_cast<LONG>(comboBounds.right - 10.0f);
    const LONG openY              = static_cast<LONG>(((comboBounds.top + comboBounds.bottom) * 0.5f) - scroll->GetScrollOffset());
    handled                       = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(openX, openY), handled));
    Require(handled, "combo field click inside scrolled panel is handled");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(openX, openY), handled));
    Require(combo->DebugIsPopupOpen(), "combo popup opens from a scrolled panel field");

    const D2D1_RECT_F popupBounds     = combo->DebugGetPopupBounds();
    const float comboBottomInViewport = comboBounds.bottom - scroll->GetScrollOffset();
    const float popupTopInViewport    = popupBounds.top - scroll->GetScrollOffset();
    Require(popupTopInViewport >= comboBottomInViewport && popupTopInViewport <= (comboBottomInViewport + 8.0f),
            "scrolled combo popup placement uses viewport space and opens adjacent to the visible field when there is room below");

    const D2D1_RECT_F secondPopupItemRect = combo->DebugGetPopupItemRect(1u, &host);
    RequireRectHasArea(secondPopupItemRect, "scrolled combo popup exposes second-row geometry");
    const LONG popupX = static_cast<LONG>((secondPopupItemRect.left + secondPopupItemRect.right) * 0.5f);
    const LONG popupY = static_cast<LONG>(((secondPopupItemRect.top + secondPopupItemRect.bottom) * 0.5f) - scroll->GetScrollOffset());
    handled           = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(popupX, popupY), handled));
    Require(handled, "combo popup row click inside a scrolled panel is handled at the visual position");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(popupX, popupY), handled));

    Require(combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == 1u,
            "scrolled combo popup selects the row clicked at the visual popup position");
}

void TestScrollPanelPaintsOverlaysInScrolledViewportCoordinates()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* scroll = root->AddChild<ScrollPanel>();
    auto* probe  = scroll->AddChild<OverlayPaintProbeControl>();
    scroll->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 180.0f));
    scroll->SetContentHeight(480.0f);
    scroll->SetScrollOffset(120.0f);
    probe->SetBounds(D2D1::RectF(40.0f, 150.0f, 160.0f, 176.0f));

    window.Host().SetRoot(std::move(root));
    static_cast<Panel*>(window.Host().GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 180.0f));

    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmapForComboSuite(window, "scroll panel overlay paint transform capture succeeds");
    const uint32_t visualPixel            = GetCapturePixelBgra(capture, 60u, 42u);
    const uint32_t rawPixel               = GetCapturePixelBgra(capture, 60u, 162u);
    const auto isRedPixel                 = [](uint32_t bgra) noexcept
    {
        const uint8_t blue  = static_cast<uint8_t>((bgra >> 0u) & 0xFFu);
        const uint8_t green = static_cast<uint8_t>((bgra >> 8u) & 0xFFu);
        const uint8_t red   = static_cast<uint8_t>((bgra >> 16u) & 0xFFu);
        return red > 180u && green < 80u && blue < 80u;
    };

    Require(isRedPixel(visualPixel), "scroll panel paints child overlays at the scrolled visual position");
    Require(! isRedPixel(rawPixel), "scroll panel does not paint child overlays at the unscrolled content position");
}

void TestAttachedComboBoxPopupMouseLeaveRestoresKeyboardHighlightedItem()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(12.0f, 12.0f, 192.0f, 40.0f));
    combo->SetItems({
        ComboBox::Item{L"one", L"One"},
        ComboBox::Item{L"two", L"Two"},
        ComboBox::Item{L"three", L"Three"},
        ComboBox::Item{L"four", L"Four"},
    });
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(172, 24), handled));
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(172, 24), handled));
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "attached combo popup opens before keyboard highlight restore test");

    Require(combo->OnKeyDown(host, VK_DOWN, 0), "attached combo popup handles keyboard down while open");
    Require(combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == 1u, "attached combo popup keyboard down updates the selected row");
    Require(combo->DebugGetHoveredPopupIndex().has_value() && combo->DebugGetHoveredPopupIndex().value() == 1u,
            "attached combo popup keyboard down highlights the second popup row");
    const size_t keyboardHighlightedPopupIndex = combo->DebugGetHoveredPopupIndex().value();

    const D2D1_RECT_F thirdPopupItemRect = combo->DebugGetPopupItemRect(2u, &host);
    RequireRectHasArea(thirdPopupItemRect, "attached combo popup exposes third-row geometry for keyboard highlight restore");
    const D2D1_POINT_2F thirdRowHoverPoint = D2D1::Point2F(thirdPopupItemRect.left + ((thirdPopupItemRect.right - thirdPopupItemRect.left) * 0.35f),
                                                           (thirdPopupItemRect.top + thirdPopupItemRect.bottom) * 0.5f);

    Require(combo->OnMouseMove(host, thirdRowHoverPoint, 0), "attached combo popup handles pointer hover over a different row");
    Require(combo->DebugGetHoveredPopupIndex().has_value() && combo->DebugGetHoveredPopupIndex().value() == 2u,
            "attached combo popup pointer hover temporarily overrides the keyboard-highlighted row");

#ifdef _DEBUG
    const uint64_t invalidateCountBeforeMouseLeave = host.DebugGetInvalidateCount();
#endif

    Require(combo->OnMouseLeave(host), "attached combo popup handles mouse leave while the popup is open");
    Require(combo->DebugGetHoveredPopupIndex().has_value() && combo->DebugGetHoveredPopupIndex().value() == keyboardHighlightedPopupIndex,
            "attached combo popup mouse leave restores the keyboard-highlighted row");
    Require(combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == keyboardHighlightedPopupIndex,
            "attached combo popup mouse leave preserves the keyboard-selected row");

#ifdef _DEBUG
    Require(host.DebugGetInvalidateCount() > invalidateCountBeforeMouseLeave,
            "attached combo popup mouse leave invalidates when it restores the keyboard-highlighted row");
    // Render count not asserted: headless WindowHost has no WM_PAINT loop,
    // so Render() is never invoked.  The invalidate-count check above proves
    // the control requests a repaint; actual rendering is platform-level.
#endif
}

void TestComboBoxPopupFlipsAboveWhenBelowSpaceIsInsufficient()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(200, 120), handled));
    Require(handled, "host size update handled for popup placement test");

    ComboBox combo;
    combo.SetBounds(D2D1::RectF(10.0f, 90.0f, 190.0f, 118.0f));
    combo.SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}, ComboBox::Item{L"three", L"Three"}, ComboBox::Item{L"four", L"Four"}});

    Require(combo.OnMouseDown(host, D2D1::Point2F(182.0f, 104.0f), false, 0), "combo opens for flipped popup test");
    Require(combo.GetHitBounds().top < combo.GetBounds().top, "popup extends hit bounds above combo near bottom edge");
    const D2D1_RECT_F secondPopupItemRect = combo.DebugGetPopupItemRect(1u, &host);
    RequireRectHasArea(secondPopupItemRect, "flipped popup exposes second-row geometry under shared menu metrics");
    const D2D1_POINT_2F secondPopupItemPoint =
        D2D1::Point2F((secondPopupItemRect.left + secondPopupItemRect.right) * 0.5f, (secondPopupItemRect.top + secondPopupItemRect.bottom) * 0.5f);
    Require(combo.OnMouseDown(host, secondPopupItemPoint, false, 0), "click inside flipped popup is handled");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 1u, "flipped popup hit-testing selects row above combo");
}

void TestAttachedComboBoxFlippedPopupItemRectsStayInsidePopupBounds()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(320, 200), handled));
    Require(handled, "host size update handled for attached flipped combo popup geometry test");

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}, ComboBox::Item{L"three", L"Three"}, ComboBox::Item{L"four", L"Four"}});
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 200.0f));

    const D2D1_RECT_F clientBounds = host.GetClientBoundsDip();
    combo->SetBounds(D2D1::RectF(10.0f, clientBounds.bottom - 30.0f, clientBounds.right - 10.0f, clientBounds.bottom - 2.0f));
    host.Invalidate();

    const D2D1_RECT_F comboBounds = combo->GetBounds();
    const LONG comboDropX         = static_cast<LONG>(comboBounds.right - 8.0f);
    const LONG comboCenterY       = static_cast<LONG>((comboBounds.top + comboBounds.bottom) * 0.5f);
    handled                       = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(comboDropX, comboCenterY), handled));
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(comboDropX, comboCenterY), handled));

    const D2D1_RECT_F popupBounds         = combo->DebugGetPopupBounds();
    const D2D1_RECT_F firstPopupItemRect  = combo->DebugGetPopupItemRect(0u, &host);
    const D2D1_RECT_F secondPopupItemRect = combo->DebugGetPopupItemRect(1u, &host);
    RequireRectHasArea(popupBounds, "attached flipped combo popup exposes popup bounds");
    RequireRectHasArea(firstPopupItemRect, "attached flipped combo popup exposes first-row geometry");
    RequireRectHasArea(secondPopupItemRect, "attached flipped combo popup exposes second-row geometry");
    Require(popupBounds.bottom <= combo->GetBounds().top, "attached flipped combo popup stays above the combo near the bottom edge");
    Require(firstPopupItemRect.top >= popupBounds.top && firstPopupItemRect.bottom <= popupBounds.bottom,
            "attached flipped combo first-row geometry stays inside the popup bounds");
    Require(secondPopupItemRect.top >= popupBounds.top && secondPopupItemRect.bottom <= popupBounds.bottom,
            "attached flipped combo second-row geometry stays inside the popup bounds");
    Require(firstPopupItemRect.bottom <= combo->GetBounds().top, "attached flipped combo first-row geometry stays above the combo field");
}

void TestComboBoxPageDownUsesClampedVisiblePopupRows()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(200, 120), handled));
    Require(handled, "host size update handled for paged popup placement test");

    ComboBox combo;
    combo.SetBounds(D2D1::RectF(10.0f, 90.0f, 190.0f, 118.0f));
    std::vector<ComboBox::Item> items;
    for (size_t index = 0; index < 12u; ++index)
    {
        items.push_back(ComboBox::Item{std::format(L"value{:02}", index), std::format(L"Item {:02}", index)});
    }
    combo.SetItems(std::move(items));

    Require(combo.OnMouseDown(host, D2D1::Point2F(182.0f, 104.0f), false, 0), "combo opens for constrained page-down test");
    size_t visibleRowCount = 0u;
    for (size_t index = 0u; index < 12u; ++index)
    {
        const D2D1_RECT_F itemRect = combo.DebugGetPopupItemRect(index, &host);
        if (itemRect.right > itemRect.left && itemRect.bottom > itemRect.top)
        {
            ++visibleRowCount;
        }
    }
    Require(visibleRowCount > 0u, "constrained popup exposes at least one visible row for page-down stepping");
    Require(combo.OnKeyDown(host, VK_NEXT, 0), "page down handled for constrained popup");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == std::min<size_t>(visibleRowCount, 11u),
            "page down advances by clamped visible popup rows");
}

void TestComboBoxPopupWidensForLongVisibleEntries()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(360, 180), handled));
    Require(handled, "host size update handled for popup width test");

    ComboBox combo;
    combo.SetBounds(D2D1::RectF(10.0f, 10.0f, 150.0f, 38.0f));
    combo.SetItems(
        {ComboBox::Item{L"short", L"Short"}, ComboBox::Item{L"very-long", L"Very long popup entry used to verify wider history dropdown hit-testing"}});

    Require(combo.OnMouseDown(host, D2D1::Point2F(142.0f, 24.0f), false, 0), "combo opens for widened popup test");

    const auto openBounds = combo.GetHitBounds();
    Require(openBounds.right > combo.GetBounds().right, "popup widens beyond combo width for long visible entries");
    Require(combo.OnMouseDown(host, D2D1::Point2F(openBounds.right - 6.0f, combo.GetBounds().bottom + 42.0f), false, 0),
            "click inside widened popup area is handled");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 1u, "widened popup hit-testing selects the long entry");
}

void TestComboBoxPopupWidthRemainsStableAcrossPopupScroll()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(360, 180), handled));
    Require(handled, "host size update handled for popup scroll width test");

    ComboBox combo;
    combo.SetBounds(D2D1::RectF(10.0f, 10.0f, 150.0f, 38.0f));
    std::vector<ComboBox::Item> items;
    items.push_back(ComboBox::Item{L"long-entry", L"Very long popup entry that must keep the dropdown width stable while scrolling"});
    for (size_t index = 0u; index < 16u; ++index)
    {
        items.push_back(ComboBox::Item{std::format(L"item-{:02}", index), std::format(L"Item {:02}", index)});
    }
    combo.SetItems(std::move(items));

    Require(combo.OnMouseDown(host, D2D1::Point2F(142.0f, 24.0f), false, 0), "combo opens for popup scroll width test");
    const auto openBounds    = combo.GetHitBounds();
    const float openWidthDip = openBounds.right - openBounds.left;

    Require(combo.OnKeyDown(host, VK_NEXT, 0), "combo handles page-down while popup is open");
    const auto scrolledBounds    = combo.GetHitBounds();
    const float scrolledWidthDip = scrolledBounds.right - scrolledBounds.left;

    RequireFloatNear(scrolledWidthDip, openWidthDip, 0.5f, "popup width remains stable after scrolling past the long entry");
}

void TestComboBoxPopupScrollbarPagesLongLists()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(220, 260), handled));
    Require(handled, "host size update handled for popup scrollbar test");

    ComboBox combo;
    combo.SetBounds(D2D1::RectF(10.0f, 10.0f, 170.0f, 38.0f));
    std::vector<ComboBox::Item> items;
    for (size_t index = 0; index < 12u; ++index)
    {
        items.push_back(ComboBox::Item{std::format(L"value{:02}", index), std::format(L"Item {:02}", index)});
    }
    combo.SetItems(std::move(items));

    Require(combo.OnMouseDown(host, D2D1::Point2F(162.0f, 24.0f), false, 0), "combo opens for popup scrollbar test");
    const auto openBounds = combo.GetHitBounds();
    Require(openBounds.right > combo.GetBounds().right, "popup scrollbar expands hit bounds beyond the combo width");
    size_t visibleRowCount = 0u;
    for (size_t index = 0u; index < 12u; ++index)
    {
        const D2D1_RECT_F itemRect = combo.DebugGetPopupItemRect(index, &host);
        if (itemRect.right > itemRect.left && itemRect.bottom > itemRect.top)
        {
            ++visibleRowCount;
        }
    }
    Require(visibleRowCount > 0u && visibleRowCount < 12u, "popup scrollbar test resolves a partially visible viewport");
    Require(combo.OnMouseDown(host, D2D1::Point2F(openBounds.right - 4.0f, combo.GetBounds().bottom + 176.0f), false, 0),
            "popup scrollbar track click is handled");
    Require(combo.OnMouseDown(host, D2D1::Point2F(20.0f, combo.GetBounds().bottom + 14.0f), false, 0),
            "click first visible row after scrollbar paging is handled");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == (12u - visibleRowCount),
            "popup scrollbar track paging reveals later rows");
}

void TestComboBoxPopupScrollbarThumbGutterDragThroughWindowHost()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(220, 260), handled));
    Require(handled, "host size update handled for combo thumb gutter drag");

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(10.0f, 10.0f, 170.0f, 38.0f));
    std::vector<ComboBox::Item> items;
    for (size_t index = 0; index < 20u; ++index)
    {
        items.push_back(ComboBox::Item{std::format(L"value{:02}", index), std::format(L"Item {:02}", index)});
    }
    combo->SetItems(std::move(items));

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 260.0f));

    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(162, 24), handled));
    Require(handled, "combo opens for popup thumb gutter drag");

    const D2D1_RECT_F popup = combo->DebugGetPopupBounds();
    RequireRectHasArea(popup, "combo popup exposes geometry for thumb gutter drag");

    const LONG gutterX = static_cast<LONG>(popup.right - 11.0f);
    const LONG gutterY = static_cast<LONG>(popup.top + 5.0f);
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(gutterX, gutterY), handled));
    Require(handled, "combo popup handles scrollbar thumb gutter mouse-down as a drag");

    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, MK_LBUTTON, MAKELPARAM(gutterX, gutterY + 48), handled));
    Require(handled, "combo popup handles captured scrollbar thumb gutter mouse-move");
    const D2D1_RECT_F firstItemRect = combo->DebugGetPopupItemRect(0u, &host);
    Require(firstItemRect.right <= firstItemRect.left || firstItemRect.bottom <= firstItemRect.top,
            "combo popup thumb gutter drag scrolls past the first item");

    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(gutterX, gutterY + 48), handled));
    Require(handled, "combo popup handles captured scrollbar thumb gutter mouse-up");
}

void TestComboBoxPopupCapsVisibleItemsAndScrolls()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    std::vector<ComboBox::Item> items;
    for (size_t index = 0; index < 12u; ++index)
    {
        items.push_back(ComboBox::Item{std::wstring(1u, static_cast<wchar_t>(L'a' + index)), std::format(L"Item {:02}", index)});
    }
    combo.SetItems(std::move(items));

    Require(combo.OnMouseDown(host, D2D1::Point2F(172.0f, 12.0f), false, 0), "long combo popup opens");
    const auto openBounds = combo.GetHitBounds();
    Require(openBounds.bottom < 28.0f + 2.0f + 8.0f + (24.0f * 12.0f), "long combo popup height is capped");

    const D2D1_POINT_2F firstVisibleItemPoint = D2D1::Point2F(16.0f, 42.0f);
    Require(combo.OnMouseWheel(host, firstVisibleItemPoint, -static_cast<float>(WHEEL_DELTA), 0), "mouse wheel scrolls open combo popup");
    Require(combo.OnMouseWheel(host, firstVisibleItemPoint, -static_cast<float>(WHEEL_DELTA), 0), "second wheel scrolls open combo popup");
    Require(combo.OnMouseWheel(host, firstVisibleItemPoint, -static_cast<float>(WHEEL_DELTA), 0), "third wheel scrolls open combo popup");
    Require(combo.OnMouseDown(host, firstVisibleItemPoint, false, 0), "click selects first visible item after scrolling");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 3u, "scrolling changes which popup item is selected");
}

void TestComboBoxTypeaheadSelectsMatchingItem()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetItems(
        {ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}, ComboBox::Item{L"gamma", L"Gamma"}, ComboBox::Item{L"delta", L"Delta"}});

    Require(combo.OnChar(host, L'g', 0), "typeahead character handled");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 2u, "typeahead selects matching later item");
    Require(combo.GetSelectedValue() == L"gamma", "typeahead updates combo value");
}

void TestComboBoxEightItemsAllFitInDefaultPopup()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(200, 280), handled));

    ComboBox combo;
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    // Exactly kComboBoxMaxVisibleItems = 8 items; all should fit without scrolling.
    combo.SetItems({ComboBox::Item{L"a", L"Alpha"},
                    ComboBox::Item{L"b", L"Beta"},
                    ComboBox::Item{L"c", L"Gamma"},
                    ComboBox::Item{L"d", L"Delta"},
                    ComboBox::Item{L"e", L"Epsilon"},
                    ComboBox::Item{L"f", L"Zeta"},
                    ComboBox::Item{L"g", L"Eta"},
                    ComboBox::Item{L"h", L"Theta"}});

    Require(combo.OnMouseDown(host, D2D1::Point2F(172.0f, 14.0f), false, 0), "8-item combo opens popup");
    const D2D1_RECT_F lastItemRect = combo.DebugGetPopupItemRect(7u, &host);
    RequireRectHasArea(lastItemRect, "last item of 8-item combo is visible without scrolling");
    Require(combo.OnMouseDown(host, D2D1::Point2F((lastItemRect.left + lastItemRect.right) * 0.5f, (lastItemRect.top + lastItemRect.bottom) * 0.5f), false, 0),
            "click on last item handled");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 7u,
            "last item of 8-item combo is directly clickable without scrolling");
}

void TestComboBoxSetMaxVisibleItemsAllowsLargerPopup()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(200, 320), handled));

    ComboBox combo;
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    // 9 items — one more than the default cap of 8. Without SetMaxVisibleItems the 9th would be unreachable.
    combo.SetItems({ComboBox::Item{L"a", L"Alpha"},
                    ComboBox::Item{L"b", L"Beta"},
                    ComboBox::Item{L"c", L"Gamma"},
                    ComboBox::Item{L"d", L"Delta"},
                    ComboBox::Item{L"e", L"Epsilon"},
                    ComboBox::Item{L"f", L"Zeta"},
                    ComboBox::Item{L"g", L"Eta"},
                    ComboBox::Item{L"h", L"Theta"},
                    ComboBox::Item{L"i", L"Iota"}});
    combo.SetMaxVisibleItems(9u);

    Require(combo.OnMouseDown(host, D2D1::Point2F(172.0f, 14.0f), false, 0), "9-item combo opens popup");
    const D2D1_RECT_F lastItemRect = combo.DebugGetPopupItemRect(8u, &host);
    RequireRectHasArea(lastItemRect, "SetMaxVisibleItems(9) keeps the 9th item visible");
    Require(combo.OnMouseDown(host, D2D1::Point2F((lastItemRect.left + lastItemRect.right) * 0.5f, (lastItemRect.top + lastItemRect.bottom) * 0.5f), false, 0),
            "click on 9th item handled");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 8u, "SetMaxVisibleItems(9) makes the 9th item directly clickable");
}

void TestComboBoxCompactEditableTextRectPreservesInsetAndWidth()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(240, 80), handled));
    Require(handled, "compact editable combo host size update handled");

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetVariant(ComboBoxVariant::Edit);
    combo->SetBounds(D2D1::RectF(12.0f, 12.0f, 204.0f, 40.0f));
    combo->SetText(L"Compact mode");
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 80.0f));

    ThemePalette theme = MakeDefaultThemePalette(true);
    theme.density      = Density::Compact;
    host.SetTheme(theme);

    const D2D1_RECT_F comboBounds = combo->GetBounds();
    const D2D1_RECT_F textRect    = combo->DebugGetEditableTextRect();
    RequireRectHasArea(textRect, "compact editable combo exposes a usable text rect");
    Require(textRect.left >= comboBounds.left + 9.0f, "compact editable combo keeps a left inset for the visible text");
    Require(textRect.right <= comboBounds.right - 20.0f, "compact editable combo keeps clear space for the drop button");
    Require((textRect.right - textRect.left) >= 120.0f, "compact editable combo keeps a non-trivial compact text width");
    Require(textRect.top >= comboBounds.top + 2.0f, "compact editable combo keeps vertical inset above the text");
    Require(textRect.bottom <= comboBounds.bottom - 2.0f, "compact editable combo keeps vertical inset below the text");
}

void TestNativeEditableComboBoxEmojiUsesColorFontRendering()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    constexpr std::wstring_view kFireEmoji = L"\xD83D\xDD25";
    std::wstring text                      = L"combo ";
    for (int index = 0; index < 10; ++index)
    {
        text.append(kFireEmoji);
    }

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetBounds(D2D1::RectF(20.0f, 20.0f, 520.0f, 64.0f));
    combo->SetText(text);
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);

    Require(window.Host().GetTextInputBackend() == TextInputBackend::Native, "native editable combo color-font test uses the native backend");
    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native editable combo color-font test does not create a bridge child");

    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmapForComboSuite(window, "native editable combo emoji color-font capture succeeds");
    const size_t warmPixels               = CountWarmSaturatedPixels(capture);
    EmitColorGlyphPixelCountForTest(L"native-editable-combo", capture, warmPixels);
    Require(warmPixels >= 24u, "native editable combo renders warm emoji color-font pixels");
}

void TestEditableComboBoxPaintReusesSingleLineLayoutAndCaretOffset()
{
    const std::filesystem::path repoRoot = FindRepoRootForDxUiTests();

    std::ifstream headerInput(repoRoot / L"Common" / L"DxUi" / L"DxUi.h");
    Require(headerInput.good(), "DxUi header is readable for editable ComboBox paint-cache guard");
    const std::string header((std::istreambuf_iterator<char>(headerInput)), std::istreambuf_iterator<char>());

    std::ifstream sourceInput(repoRoot / L"Common" / L"DxUi" / L"DxUi.ComboBox.cpp");
    Require(sourceInput.good(), "ComboBox source is readable for editable paint-cache guard");
    const std::string source((std::istreambuf_iterator<char>(sourceInput)), std::istreambuf_iterator<char>());

    const auto requireBlock = [](const std::string& text, const char* beginMarker, const char* endMarker, const char* description)
    {
        const size_t begin = text.find(beginMarker);
        const size_t end   = text.find(endMarker, begin == std::string::npos ? 0u : begin + 1u);
        Require(begin != std::string::npos && end != std::string::npos && begin < end, description);
        return text.substr(begin, end - begin);
    };

    const std::string comboClassBlock = requireBlock(header, "class ComboBox : public Control", "class Tree final", "ComboBox class block is found");
    Require(comboClassBlock.find("SingleLineTextLayoutCache _singleLineLayoutCache") != std::string::npos,
            "editable ComboBox owns the shared single-line layout cache");
    Require(comboClassBlock.find("InvalidateSingleLineLayoutCache() const noexcept") != std::string::npos,
            "editable ComboBox declares a single-line cache invalidator");
    Require(comboClassBlock.find("GetOrCreateSingleLineLayout(") != std::string::npos, "editable ComboBox declares a cache-backed single-line layout accessor");

    const std::string paintBlock = requireBlock(source, "void ComboBox::Paint", "void ComboBox::PaintOverlay", "ComboBox paint block is found");
    Require(paintBlock.find("const wil::com_ptr<IDWriteTextLayout> editableLayout") != std::string::npos,
            "editable ComboBox paint creates or reuses one layout for text, selection, and caret");
    Require(paintBlock.find("DrawSingleLineSelectionWithLayout") != std::string::npos, "editable ComboBox paint draws selection through the cached layout");
    Require(paintBlock.find("editableCaretOffsetDip = EnsureEditableCaretVisible") != std::string::npos,
            "editable ComboBox paint reuses the caret offset from its single visibility pass");
    Require(CountSubstring(paintBlock, "EnsureEditableCaretVisible(") == 1u, "editable ComboBox paint ensures caret visibility only once per frame");
    Require(paintBlock.find("MeasureCaretOffsetDip(&host") == std::string::npos, "editable ComboBox paint no longer creates a separate caret layout");

    const std::string setTextBlock = requireBlock(source, "void ComboBox::SetText", "void ComboBox::SetTextAndNotify", "ComboBox SetText block is found");
    Require(setTextBlock.find("InvalidateSingleLineLayoutCache()") != std::string::npos, "editable ComboBox text changes invalidate the cached layout");

    const std::string invalidateBlock = requireBlock(
        source, "void ComboBox::InvalidateSingleLineLayoutCache", "float ComboBox::EnsureEditableCaretVisible", "ComboBox cache invalidator block is found");
    Require(invalidateBlock.find("ClearSingleLineTextLayoutCache(_singleLineLayoutCache)") != std::string::npos,
            "editable ComboBox cache invalidator clears the shared cache");

    const std::string ensureBlock = requireBlock(
        source, "float ComboBox::EnsureEditableCaretVisible", "void ComboBox::ResetEditableCaretBlink", "ComboBox caret-visibility block is found");
    Require(ensureBlock.find("GetOrCreateSingleLineLayout(") != std::string::npos, "editable ComboBox caret visibility uses the shared layout cache");
    Require(ensureBlock.find("MeasureCaretOffsetDip(layout.get()") != std::string::npos, "editable ComboBox caret visibility measures from the cached layout");
    Require(ensureBlock.find("MeasureCaretOffsetDip(host") == std::string::npos,
            "editable ComboBox caret visibility no longer creates a separate uncached layout");
}

void TestPopupBackdropCaptureUsesSharedBoundedHelper()
{
    const std::filesystem::path repoRoot = FindRepoRootForDxUiTests();

    const auto readSource = [](const std::filesystem::path& path, const char* context)
    {
        std::ifstream input(path);
        Require(input.good(), context);
        return std::string((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
    };

    const std::string internalHeader =
        readSource(repoRoot / L"Common" / L"DxUi" / L"DxUi.Internal.h", "DxUi internal header is readable for backdrop-capture guard");
    const std::string sharedSource = readSource(repoRoot / L"Common" / L"DxUi" / L"DxUi.cpp", "DxUi shared source is readable for backdrop-capture guard");
    const std::string comboSource = readSource(repoRoot / L"Common" / L"DxUi" / L"DxUi.ComboBox.cpp", "ComboBox source is readable for backdrop-capture guard");
    const std::string menuSource  = readSource(repoRoot / L"Common" / L"DxUi" / L"DxUi.Menu.cpp", "Menu source is readable for backdrop-capture guard");

    Require(internalHeader.find("CaptureBackdropScreenRegion(") != std::string::npos, "bounded popup backdrop capture is declared as one shared DxUi helper");
    Require(sharedSource.find("bool CaptureBackdropScreenRegion(") != std::string::npos,
            "bounded popup backdrop capture is implemented once in shared DxUi code");
    Require(sharedSource.find("CreateDIBSection(") != std::string::npos && sharedSource.find("BitBlt(") != std::string::npos,
            "shared backdrop helper owns the GDI screen-copy implementation");

    Require(comboSource.find("CaptureComboBoxBackdropScreenRegion") == std::string::npos, "ComboBox does not keep a duplicate backdrop capture helper");
    Require(menuSource.find("CaptureMenuBackdropScreenRegion") == std::string::npos, "Menu does not keep a duplicate backdrop capture helper");
    Require(comboSource.find("CreateDIBSection(") == std::string::npos && comboSource.find("BitBlt(") == std::string::npos,
            "ComboBox calls the shared bounded helper instead of owning GDI screen-copy code");
    Require(menuSource.find("CreateDIBSection(") == std::string::npos && menuSource.find("BitBlt(") == std::string::npos,
            "Menu calls the shared bounded helper instead of owning GDI screen-copy code");

    const size_t comboLayout    = comboSource.find("UpdatePopupLayout(&host)");
    const size_t comboPopupRect = comboSource.find("const RECT popupRectPx", comboLayout);
    const size_t comboCapture   = comboSource.find("CaptureBackdropScreenRegion(popupRectPx, _popupBackdropCapture", comboPopupRect);
    Require(comboLayout != std::string::npos && comboPopupRect != std::string::npos && comboCapture != std::string::npos && comboLayout < comboPopupRect &&
                comboPopupRect < comboCapture,
            "ComboBox captures only the current bounded popup rectangle after layout is updated");

    const size_t menuCreate      = menuSource.find("CreateMenuPopupWindow");
    const size_t menuSurfaceRect = menuSource.find("const RECT surfaceRectPx", menuCreate);
    const size_t menuWindowRect  = menuSource.find("const RECT windowRect", menuSurfaceRect);
    const size_t menuCapture     = menuSource.find("CaptureBackdropScreenRegion(surfaceRectPx, popup->backdropSnapshot.capture", menuWindowRect);
    Require(menuCreate != std::string::npos && menuSurfaceRect != std::string::npos && menuWindowRect != std::string::npos &&
                menuCapture != std::string::npos && menuSurfaceRect < menuWindowRect && menuWindowRect < menuCapture,
            "Menu captures the visible surface rectangle, not the larger popup window shadow bounds");
    Require(menuSource.find("CaptureBackdropScreenRegion(popup->windowRectPx") == std::string::npos &&
                menuSource.find("CaptureBackdropScreenRegion(windowRect") == std::string::npos,
            "Menu backdrop capture does not use the larger window rectangle");
}

void TestComboBoxCompactPopupItemTextRectPreservesInsetAndWidth()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(240, 180), handled));
    Require(handled, "compact combo popup host size update handled");

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(12.0f, 12.0f, 204.0f, 40.0f));
    combo->SetItems({ComboBox::Item{L"one", L"System"}, ComboBox::Item{L"two", L"On"}, ComboBox::Item{L"three", L"Off"}});
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 180.0f));

    ThemePalette theme = MakeDefaultThemePalette(true);
    theme.density      = Density::Compact;
    host.SetTheme(theme);

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(196, 26), handled));
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(196, 26), handled));

    const D2D1_RECT_F itemRect     = combo->DebugGetPopupItemRect(0u, &host);
    const D2D1_RECT_F itemTextRect = combo->DebugGetPopupItemTextRect(0u, &host);
    RequireRectHasArea(itemRect, "compact combo popup exposes first-row bounds");
    RequireRectHasArea(itemTextRect, "compact combo popup exposes first-row text bounds");
    Require(itemTextRect.left >= itemRect.left + 9.0f, "compact combo popup keeps left padding before item text");
    Require(itemTextRect.right <= itemRect.right - 7.0f, "compact combo popup keeps right padding after item text");
    Require((itemTextRect.right - itemTextRect.left) >= 60.0f, "compact combo popup keeps usable text width");
    Require((itemRect.bottom - itemRect.top) >= 23.0f, "compact combo popup matches the shared compact menu row height contract");
}

void TestComboBoxPopupUsesRoundSmallCornerRadius()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    ThemePalette theme    = MakeDefaultThemePalette(false);
    theme.overlayMaterial = OverlayMaterial::Solid;
    window.Host().SetTheme(theme);

    auto root      = std::make_unique<Panel>();
    auto* backdrop = root->AddChild<StripedBackdropControl>(10);
    auto* combo    = root->AddChild<ComboBox>();
    backdrop->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 200.0f));
    combo->SetBounds(D2D1::RectF(24.0f, 24.0f, 216.0f, 56.0f));
    combo->SetItems({
        ComboBox::Item{L"system", L"System"},
        ComboBox::Item{L"on", L"On"},
        ComboBox::Item{L"off", L"Off"},
    });
    combo->SetSelectedIndex(1u);
    window.Host().SetRoot(std::move(root));

    Require(combo->OnMouseDown(window.Host(), D2D1::Point2F(208.0f, 40.0f), false, 0), "combo popup opens for RoundSmall corner validation");
    const D2D1_RECT_F popupBounds = combo->DebugGetPopupBounds();
    RequireRectHasArea(popupBounds, "combo popup exposes popup bounds for RoundSmall corner validation");

    const WindowHostBitmapCapture capture =
        CaptureAttachedHostWindowBitmapForComboSuite(window, "combo popup RoundSmall corner validation captures the attached host");
    const UINT cornerX         = static_cast<UINT>((std::max)(0l, std::lround(popupBounds.left + 2.0f)));
    const UINT cornerY         = static_cast<UINT>((std::max)(0l, std::lround(popupBounds.top + 2.0f)));
    const UINT fillX           = cornerX;
    const UINT fillY           = static_cast<UINT>((std::max)(0l, std::lround(popupBounds.top + 14.0f)));
    const uint32_t cornerPixel = GetCapturePixelBgra(capture, cornerX, cornerY);
    const uint32_t fillPixel   = GetCapturePixelBgra(capture, fillX, fillY);
    Require(ComputeRgbDelta(cornerPixel, fillPixel) <= 72u,
            "combo popup keeps a RoundSmall-style near-corner body pixel instead of falling back to the older deeper top-left cutout");
}

void TestComboBoxRainbowPopupUsesAccentDerivedHighlight()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF101513u;
    viewerTheme.textArgb                   = 0xFFF4FBF6u;
    viewerTheme.selectionBackgroundArgb    = 0xFF225E36u;
    viewerTheme.selectionTextArgb          = 0xFFF8FFF9u;
    viewerTheme.accentArgb                 = 0xFF2EE861u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = TRUE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette rainbowTheme        = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ComboBoxVisualStyle rainbowStyle = ResolveComboBoxVisualStyle(rainbowTheme, ComboBoxVariant::Window, true, false, true, true, true);

    RequireColorDifferent(
        rainbowStyle.popupActiveFill, rainbowTheme.selectionFill, "rainbow combo popup active row does not fall back to the ordinary selection fill");
    RequireColorDifferent(rainbowStyle.popupSelectedFill,
                          rainbowTheme.selectionInactiveFill,
                          "rainbow combo popup selected row does not fall back to the ordinary inactive selection fill");

    ThemePalette highContrastTheme              = rainbowTheme;
    highContrastTheme.highContrast              = true;
    const ComboBoxVisualStyle highContrastStyle = ResolveComboBoxVisualStyle(highContrastTheme, ComboBoxVariant::Window, true, false, true, true, true);
    RequireColorNear(
        highContrastStyle.popupActiveFill, highContrastTheme.selectionFill, "high-contrast combo popup active row falls back to the shared selection fill");
    RequireColorNear(highContrastStyle.popupSelectedFill,
                     highContrastTheme.selectionInactiveFill,
                     "high-contrast combo popup selected row falls back to the shared inactive selection fill");
}

} // namespace

void RunComboBoxTests()
{
    auto runTest = [](const char* name, void (*fn)())
    {
        std::cerr << "  [START] " << name << '\n' << std::flush;
        fn();
        std::cerr << "  [DONE] " << name << '\n' << std::flush;
    };

    runTest("TestComboRightClickInvokesContextMenuWithoutOpeningPopup", TestComboRightClickInvokesContextMenuWithoutOpeningPopup);
    runTest("TestComboBoxClosesOnFocusLoss", TestComboBoxClosesOnFocusLoss);
    runTest("TestComboBoxSecondClickTogglesPopupClosed", TestComboBoxSecondClickTogglesPopupClosed);
    runTest("TestAttachedComboBoxSentMouseHoverAndOutsideDismiss", TestAttachedComboBoxSentMouseHoverAndOutsideDismiss);
    runTest("TestComboBoxPopupOverlayBlocksUnderlyingHover", TestComboBoxPopupOverlayBlocksUnderlyingHover);
    runTest("TestComboBoxPopupOverlayBlocksUnderlyingClick", TestComboBoxPopupOverlayBlocksUnderlyingClick);
    runTest("TestComboBoxPopupInsideScrolledPanelUsesViewportCoordinates", TestComboBoxPopupInsideScrolledPanelUsesViewportCoordinates);
    runTest("TestScrollPanelPaintsOverlaysInScrolledViewportCoordinates", TestScrollPanelPaintsOverlaysInScrolledViewportCoordinates);
    runTest("TestAttachedComboBoxPopupMouseLeaveRestoresKeyboardHighlightedItem", TestAttachedComboBoxPopupMouseLeaveRestoresKeyboardHighlightedItem);
    runTest("TestComboBoxPopupFlipsAboveWhenBelowSpaceIsInsufficient", TestComboBoxPopupFlipsAboveWhenBelowSpaceIsInsufficient);
    runTest("TestAttachedComboBoxFlippedPopupItemRectsStayInsidePopupBounds", TestAttachedComboBoxFlippedPopupItemRectsStayInsidePopupBounds);
    runTest("TestComboBoxPageDownUsesClampedVisiblePopupRows", TestComboBoxPageDownUsesClampedVisiblePopupRows);
    runTest("TestComboBoxPopupWidensForLongVisibleEntries", TestComboBoxPopupWidensForLongVisibleEntries);
    runTest("TestComboBoxPopupWidthRemainsStableAcrossPopupScroll", TestComboBoxPopupWidthRemainsStableAcrossPopupScroll);
    runTest("TestComboBoxPopupScrollbarPagesLongLists", TestComboBoxPopupScrollbarPagesLongLists);
    runTest("TestComboBoxPopupScrollbarThumbGutterDragThroughWindowHost", TestComboBoxPopupScrollbarThumbGutterDragThroughWindowHost);
    runTest("TestComboBoxPopupCapsVisibleItemsAndScrolls", TestComboBoxPopupCapsVisibleItemsAndScrolls);
    runTest("TestComboBoxTypeaheadSelectsMatchingItem", TestComboBoxTypeaheadSelectsMatchingItem);
    runTest("TestComboBoxEightItemsAllFitInDefaultPopup", TestComboBoxEightItemsAllFitInDefaultPopup);
    runTest("TestComboBoxSetMaxVisibleItemsAllowsLargerPopup", TestComboBoxSetMaxVisibleItemsAllowsLargerPopup);
    runTest("TestComboBoxCompactEditableTextRectPreservesInsetAndWidth", TestComboBoxCompactEditableTextRectPreservesInsetAndWidth);
    runTest("TestNativeEditableComboBoxEmojiUsesColorFontRendering", TestNativeEditableComboBoxEmojiUsesColorFontRendering);
    runTest("TestEditableComboBoxPaintReusesSingleLineLayoutAndCaretOffset", TestEditableComboBoxPaintReusesSingleLineLayoutAndCaretOffset);
    runTest("TestPopupBackdropCaptureUsesSharedBoundedHelper", TestPopupBackdropCaptureUsesSharedBoundedHelper);
    runTest("TestComboBoxCompactPopupItemTextRectPreservesInsetAndWidth", TestComboBoxCompactPopupItemTextRectPreservesInsetAndWidth);
    runTest("TestComboBoxPopupUsesRoundSmallCornerRadius", TestComboBoxPopupUsesRoundSmallCornerRadius);
    runTest("TestComboBoxRainbowPopupUsesAccentDerivedHighlight", TestComboBoxRainbowPopupUsesAccentDerivedHighlight);
}
