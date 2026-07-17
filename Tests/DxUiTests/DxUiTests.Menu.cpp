#include "DxUi/DxUi.PointerInput.h"
#include "DxUi/DxUiNativeMenuInterop.h"
#include "DxUiTestHelpers.h"
#include "FolderViewEmptyStateLayout.h"
#include "FolderViewIncrementalSearch.h"
#include "FolderViewVisualState.h"
#include "Ui/AnimationDispatcher.h"

#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <commctrl.h>
#include <format>
#include <fstream>
#include <iterator>
#include <string>
#include <thread>
#include <vector>

#pragma comment(lib, "comctl32.lib")

namespace
{
[[nodiscard]] std::string NarrowAsciiForFailureMessage(std::wstring_view text)
{
    std::string result;
    result.reserve(text.size());
    for (const wchar_t ch : text)
    {
        const auto value = static_cast<unsigned int>(ch);
        result.push_back(value >= 0x20u && value <= 0x7eu ? static_cast<char>(value) : '?');
    }

    return result;
}

[[nodiscard]] std::string WideCodeUnitsForFailureMessage(std::wstring_view text)
{
    std::string result;
    for (const wchar_t ch : text)
    {
        std::format_to(std::back_inserter(result), "{:04X} ", static_cast<unsigned int>(ch));
    }

    return result;
}

void TestFolderViewIncrementalSearchKeepsContainsHighlightButUsesPrefixFocus()
{
    using namespace FolderViewIncrementalSearch;

    constexpr std::array<std::wstring_view, 4> names{{
        L"notes-a.txt",
        L"zeta.txt",
        L"ABC.txt",
        L"beta-a.txt",
    }};

    const auto displayNameAt = [&](size_t index) noexcept -> std::wstring_view { return index < names.size() ? names[index] : std::wstring_view{}; };

    const std::optional<UINT32> containsOffset = FindContainsOffsetNoCase(names[0], L"a");
    Require(containsOffset.has_value() && containsOffset.value() > 0u, "contains match remains available for incremental-search highlighting");
    Require(! StartsWithNoCase(names[0], L"a"), "a non-prefix contains match is not treated as a focus target");
    Require(StartsWithNoCase(names[2], L"a"), "prefix matching is case-insensitive");

    const std::optional<size_t> prefixIndex = FindNextPrefixMatchIndex(names.size(), 0u, true, displayNameAt, L"a");
    Require(prefixIndex.has_value() && prefixIndex.value() == 2u, "focus navigation prefers the next item whose name starts with the query");
}

void TestFolderViewInactiveVisualStateDimsNormalTextAndIcons()
{
    using namespace FolderViewVisualState;

    const auto same = [](float left, float right) noexcept { return std::fabs(left - right) <= 0.0001f; };

    Require(same(ResolveNormalTextAlpha(1.0f, true, false), 1.0f), "focused pane keeps normal text fully opaque");
    Require(same(ResolveNormalTextAlpha(1.0f, false, false), kUnfocusedPaneTextOpacity), "unfocused pane dims normal text");
    Require(same(ResolveNormalTextAlpha(1.0f, false, true), 1.0f), "selected text uses inactive-selection colors instead of normal text dimming");
    Require(same(ResolveNormalIconOpacity(1.0f, true), 1.0f), "focused pane keeps normal icons fully opaque");
    Require(same(ResolveNormalIconOpacity(1.0f, false), kUnfocusedPaneIconOpacity), "unfocused pane dims normal icons");
    Require(same(ResolveNormalIconOpacity(0.5f, false), 0.5f * kUnfocusedPaneIconOpacity), "hidden icons keep their hidden dim and get pane dimming");
    Require(same(ResolvePlaceholderIconOpacity(false), 0.4f * kUnfocusedPaneIconOpacity), "placeholder icons also dim in an unfocused pane");
    Require(same(ResolveFocusBorderAlpha(1.0f, false), kFocusBorderOpacityUnfocused), "unfocused current item keeps a dim focus border");
}

void TestFolderViewEmptyPlaceholderMetricsUseCurrentEmptyLayout()
{
    using namespace FolderViewEmptyStateLayout;

    constexpr PlaceholderItemMetricsInput brief{
        .clientWidthDip          = 640.0f,
        .clientHeightDip         = 360.0f,
        .iconSizeDip             = 16.0f,
        .estimatedCharWidthDip   = 8.0f,
        .estimatedLabelHeightDip = 18.0f,
        .detailsLineHeightDip    = 12.0f,
        .metadataLineHeightDip   = 10.0f,
        .titleLength             = 12u,
        .includeDetailsLine      = false,
        .includeMetadataLine     = false,
    };
    const auto expectedTileHeight = [](const PlaceholderItemMetricsInput& input) noexcept
    {
        constexpr float kLabelVerticalPaddingDip = 4.0f;
        constexpr float kDetailsGapDip           = 2.0f;

        float textBlockHeightDip = input.estimatedLabelHeightDip;
        if (input.includeDetailsLine)
        {
            textBlockHeightDip += kDetailsGapDip + input.detailsLineHeightDip;
        }
        if (input.includeMetadataLine)
        {
            textBlockHeightDip += kDetailsGapDip + input.metadataLineHeightDip;
        }

        return (std::max)(input.iconSizeDip, textBlockHeightDip) + (kLabelVerticalPaddingDip * 2.0f);
    };

    const PlaceholderItemMetrics briefMetrics = ResolvePlaceholderItemMetrics(brief);
    Require(briefMetrics.tileWidthDip == brief.clientWidthDip, "empty-folder placeholder uses the current client width as a full-view focus item");
    Require(briefMetrics.tileHeightDip == expectedTileHeight(brief), "empty-folder placeholder uses the current display-mode row height");

    PlaceholderItemMetricsInput narrow         = brief;
    narrow.clientWidthDip                      = 80.0f;
    const PlaceholderItemMetrics narrowMetrics = ResolvePlaceholderItemMetrics(narrow);
    Require(narrowMetrics.tileWidthDip == 80.0f, "empty-folder placeholder width follows the current client width");

    PlaceholderItemMetricsInput zeroWidth         = brief;
    zeroWidth.clientWidthDip                      = 0.0f;
    const PlaceholderItemMetrics zeroWidthMetrics = ResolvePlaceholderItemMetrics(zeroWidth);
    Require(zeroWidthMetrics.tileWidthDip == 0.0f && zeroWidthMetrics.tileHeightDip == 0.0f && zeroWidthMetrics.labelHeightDip == 0.0f,
            "empty-folder placeholder metrics clear when the current client width is zero");

    PlaceholderItemMetricsInput zeroHeight         = brief;
    zeroHeight.clientHeightDip                     = 0.0f;
    const PlaceholderItemMetrics zeroHeightMetrics = ResolvePlaceholderItemMetrics(zeroHeight);
    Require(zeroHeightMetrics.tileWidthDip == 0.0f && zeroHeightMetrics.tileHeightDip == 0.0f && zeroHeightMetrics.labelHeightDip == 0.0f,
            "empty-folder placeholder metrics clear when the current client height is zero");

    PlaceholderItemMetricsInput detailed         = brief;
    detailed.includeDetailsLine                  = true;
    const PlaceholderItemMetrics detailedMetrics = ResolvePlaceholderItemMetrics(detailed);
    Require(detailedMetrics.tileHeightDip == expectedTileHeight(detailed), "empty-folder placeholder height follows detailed display-mode row height");

    PlaceholderItemMetricsInput extraDetailed         = detailed;
    extraDetailed.includeMetadataLine                 = true;
    const PlaceholderItemMetrics extraDetailedMetrics = ResolvePlaceholderItemMetrics(extraDetailed);
    Require(extraDetailedMetrics.tileHeightDip == expectedTileHeight(extraDetailed),
            "empty-folder placeholder height follows extra-detailed display-mode row height");
}

void TestPointerInputEventMouseMoveUsesDeliveredPoint()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    SetWindowPos(window.Hwnd(), nullptr, 240, 180, 320, 200, SWP_NOZORDER | SWP_NOACTIVATE);

    constexpr POINT deliveredClientPoint{37, 49};
    POINT expectedScreenPoint = deliveredClientPoint;
    Require(ClientToScreen(window.Hwnd(), &expectedScreenPoint) != FALSE, "pointer move test converts delivered client point to screen");

    const std::optional<PointerInputEvent> event =
        TryBuildPointerInputEvent(window.Hwnd(), WM_MOUSEMOVE, MK_CONTROL, MAKELPARAM(deliveredClientPoint.x, deliveredClientPoint.y));

    Require(event.has_value(), "pointer move event is built from a mouse message");
    Require(event.value().kind == PointerInputKind::Move, "pointer move event records move kind");
    Require(event.value().targetHwnd == window.Hwnd(), "pointer move event records target HWND");
    Require(event.value().rootHwnd == GetAncestor(window.Hwnd(), GA_ROOT), "pointer move event records root HWND");
    Require(event.value().captureHwnd == GetCapture(), "pointer move event records current capture HWND");
    Require(event.value().message == WM_MOUSEMOVE, "pointer move event records message");
    Require(event.value().wParam == MK_CONTROL, "pointer move event records flags");
    Require(event.value().lParam == MAKELPARAM(deliveredClientPoint.x, deliveredClientPoint.y), "pointer move event records raw lParam");
    Require(event.value().messageTime == static_cast<DWORD>(GetMessageTime()), "pointer move event stores message time metadata");
    Require(event.value().hasClientPoint, "pointer move event has a client point");
    Require(event.value().hasScreenPoint, "pointer move event has a screen point");
    Require(event.value().clientPointPx.x == deliveredClientPoint.x && event.value().clientPointPx.y == deliveredClientPoint.y,
            "pointer move event uses delivered client point");
    Require(event.value().screenPointPx.x == expectedScreenPoint.x && event.value().screenPointPx.y == expectedScreenPoint.y,
            "pointer move event derives screen point from delivered client point");
}

void TestPointerInputEventButtonUsesDeliveredPointAndFlags()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    SetWindowPos(window.Hwnd(), nullptr, 260, 210, 320, 200, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    SetCapture(window.Hwnd());
    const auto releaseCapture = wil::scope_exit([]() noexcept { static_cast<void>(ReleaseCapture()); });

    constexpr POINT deliveredClientPoint{11, 23};
    POINT expectedScreenPoint = deliveredClientPoint;
    Require(ClientToScreen(window.Hwnd(), &expectedScreenPoint) != FALSE, "button test converts delivered client point to screen");

    const std::optional<PointerInputEvent> event =
        TryBuildPointerInputEvent(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON | MK_SHIFT, MAKELPARAM(deliveredClientPoint.x, deliveredClientPoint.y));

    Require(event.has_value(), "button event is built from a mouse button message");
    Require(event.value().kind == PointerInputKind::LeftDown, "button event records left-down kind");
    Require(event.value().targetHwnd == window.Hwnd(), "button event records target HWND");
    Require(event.value().captureHwnd == window.Hwnd(), "button event records current capture HWND");
    Require((event.value().wParam & MK_LBUTTON) != 0, "button event records MK_LBUTTON flag");
    Require(event.value().clientPointPx.x == deliveredClientPoint.x && event.value().clientPointPx.y == deliveredClientPoint.y,
            "button event uses delivered client point");
    Require(event.value().screenPointPx.x == expectedScreenPoint.x && event.value().screenPointPx.y == expectedScreenPoint.y,
            "button event derives screen point from delivered client point");
}

void TestPointerInputEventWheelUsesDeliveredScreenPoint()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    SetWindowPos(window.Hwnd(), nullptr, 300, 250, 320, 200, SWP_NOZORDER | SWP_NOACTIVATE);

    POINT deliveredScreenPoint{351, 289};
    POINT expectedClientPoint = deliveredScreenPoint;
    Require(ScreenToClient(window.Hwnd(), &expectedClientPoint) != FALSE, "wheel test converts delivered screen point to client");

    const WPARAM wheelFlags                      = MAKEWPARAM(MK_RBUTTON, WHEEL_DELTA);
    const LPARAM wheelPoint                      = MAKELPARAM(deliveredScreenPoint.x, deliveredScreenPoint.y);
    const std::optional<PointerInputEvent> event = TryBuildPointerInputEvent(window.Hwnd(), WM_MOUSEWHEEL, wheelFlags, wheelPoint);

    Require(event.has_value(), "wheel event is built from a wheel message");
    Require(event.value().kind == PointerInputKind::Wheel, "wheel event records wheel kind");
    Require(event.value().wheelDelta == WHEEL_DELTA, "wheel event records delivered wheel delta");
    Require(event.value().hasScreenPoint, "wheel event has a screen point");
    Require(event.value().hasClientPoint, "wheel event has a client point");
    Require(event.value().screenPointPx.x == deliveredScreenPoint.x && event.value().screenPointPx.y == deliveredScreenPoint.y,
            "wheel event uses delivered screen point");
    Require(event.value().clientPointPx.x == expectedClientPoint.x && event.value().clientPointPx.y == expectedClientPoint.y,
            "wheel event derives client point from delivered screen point");
}

void TestPointerInputEventHasNoLiveCursorState()
{
    using namespace RedSalamander::DxUi;

    MSG message{};
    message.hwnd    = nullptr;
    message.message = WM_TIMER;
    message.wParam  = 0;
    message.lParam  = 0;
    Require(! TryBuildPointerInputEvent(message.hwnd, message.message, message.wParam, message.lParam).has_value(),
            "non-pointer messages do not build pointer input events");
    Require(! PointerInputKindFromMessage(WM_TIMER).has_value(), "non-pointer messages do not map to pointer input kinds");

    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.PointerInput.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "pointer input implementation source is readable for live-cursor guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
    Require(source.find("GetCursorPos") == std::string::npos, "pointer input core does not sample the live cursor");
}

void TestPointerInputOriginAbstractionIsRemoved()
{
    const std::filesystem::path repoRoot    = FindRepoRootForDxUiTests();
    const std::string removedSourceTypeName = std::string("PointerInput") + "Source";

    std::ifstream headerInput(repoRoot / L"Common" / L"DxUi" / L"DxUi.PointerInput.h");
    Require(headerInput.good(), "pointer input header is readable for source-abstraction guard");
    const std::string header((std::istreambuf_iterator<char>(headerInput)), std::istreambuf_iterator<char>());

    std::ifstream sourceInput(repoRoot / L"Common" / L"DxUi" / L"DxUi.PointerInput.cpp");
    Require(sourceInput.good(), "pointer input implementation is readable for source-abstraction guard");
    const std::string source((std::istreambuf_iterator<char>(sourceInput)), std::istreambuf_iterator<char>());

    Require(header.find(removedSourceTypeName) == std::string::npos, "single-value pointer input source enum is removed");
    Require(header.find("event.source") == std::string::npos && source.find("event.source") == std::string::npos,
            "pointer input events no longer carry source metadata");
    Require(source.find(removedSourceTypeName) == std::string::npos, "pointer input implementation has no pass-through source parameter");
    Require(header.find("TryBuildPointerInputEventFromMsg") == std::string::npos && source.find("TryBuildPointerInputEventFromMsg") == std::string::npos,
            "unused MSG-based pointer input builder is removed");
    Require(source.find("TryBuildPointerInputEventWithMessageTime") == std::string::npos,
            "pointer input builder no longer keeps a private message-time adapter for a removed MSG path");

    std::ifstream navigationInput(repoRoot / L"RedSalamander" / L"NavigationView.cpp");
    Require(navigationInput.good(), "NavigationView source is readable for source-abstraction consumer guard");
    const std::string navigationSource((std::istreambuf_iterator<char>(navigationInput)), std::istreambuf_iterator<char>());
    Require(navigationSource.find(removedSourceTypeName) == std::string::npos, "NavigationView pointer routing does not pass source metadata");
}

void TestNavigationViewPointerRoutingHasNoSyntheticGenerationGate()
{
    const std::filesystem::path repoRoot                   = FindRepoRootForDxUiTests();
    const std::array<std::filesystem::path, 5> sourcePaths = {
        repoRoot / L"RedSalamander" / L"NavigationView.h",
        repoRoot / L"RedSalamander" / L"NavigationView.cpp",
        repoRoot / L"RedSalamander" / L"NavigationView.Interaction.cpp",
        repoRoot / L"RedSalamander" / L"NavigationView.Edit.cpp",
        repoRoot / L"RedSalamander" / L"NavigationView.Menus.cpp",
    };

    for (const std::filesystem::path& sourcePath : sourcePaths)
    {
        std::ifstream input(sourcePath);
        Require(input.good(), "NavigationView source is readable for pointer-generation guard");
        const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
        Require(source.find("BumpInputGeneration") == std::string::npos, "NavigationView does not maintain synthetic input generations");
        Require(source.find("CurrentInputGeneration") == std::string::npos, "NavigationView does not stamp delivered pointer input with synthetic generations");
        Require(source.find("_inputGeneration") == std::string::npos, "NavigationView does not store synthetic pointer input generation state");
    }

    const std::array<std::filesystem::path, 4> tokenSurfacePaths = {
        repoRoot / L"Common" / L"DxUi" / L"DxUi.PointerInput.h",
        repoRoot / L"Common" / L"DxUi" / L"DxUi.PointerInput.cpp",
        repoRoot / L"Specs" / L"UI" / L"UI_NavigationView.md",
        repoRoot / L"Specs" / L"Testing" / L"Testing_TestCoverage.md",
    };
    for (const std::filesystem::path& sourcePath : tokenSurfacePaths)
    {
        std::ifstream input(sourcePath);
        Require(input.good(), "pointer input token surface is readable for generation guard");
        const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
        Require(source.find("InputGeneration") == std::string::npos, "pointer routing contract does not expose a vestigial InputGeneration token");
    }
}

void TestMenuWindowClassRegistrationCachesOnlySuccess()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Menu.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Menu source is readable for window-class registration guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t ensureClass = source.find("void EnsureMenuWindowClass");
    const size_t nextSection = source.find("// ---------------------------------------------------------------------------", ensureClass + 1u);
    Require(ensureClass != std::string::npos && nextSection != std::string::npos && ensureClass < nextSection, "EnsureMenuWindowClass source block is found");
    const std::string block = source.substr(ensureClass, nextSection - ensureClass);

    const size_t alreadyRegistered  = block.find("s_classRegistered.load");
    const size_t registerClass      = block.find("RegisterClassExW(&wc)");
    const size_t classAlreadyExists = block.find("ERROR_CLASS_ALREADY_EXISTS");
    const size_t markRegistered     = block.find("s_classRegistered.store(true");
    Require(alreadyRegistered != std::string::npos, "menu window-class guard checks the cached registration state without mutating it");
    Require(registerClass != std::string::npos, "menu window-class guard calls RegisterClassExW");
    Require(classAlreadyExists != std::string::npos, "menu window-class guard treats ERROR_CLASS_ALREADY_EXISTS as success");
    Require(markRegistered != std::string::npos && registerClass < markRegistered,
            "menu window-class guard marks the class registered only after RegisterClassExW succeeds");
    Require(block.find("s_classRegistered.exchange(true") == std::string::npos,
            "menu window-class guard does not cache registration before RegisterClassExW succeeds");
}

void TestMenuPopupWindowRegionTransfersOwnershipOnlyAfterSuccess()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Menu.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Menu source is readable for window-region ownership guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t applyRegion  = source.find("void ApplyMenuPopupWindowRegion");
    const size_t nextFunction = source.find("[[nodiscard]] UINT ResolveMenuPopupMessageDpi", applyRegion);
    Require(applyRegion != std::string::npos && nextFunction != std::string::npos && applyRegion < nextFunction,
            "ApplyMenuPopupWindowRegion source block is found");
    const std::string block = source.substr(applyRegion, nextFunction - applyRegion);

    const size_t setWindowRgn  = block.find("SetWindowRgn(hwnd, region.get(), FALSE)");
    const size_t releaseRegion = block.find("region.release()");
    Require(setWindowRgn != std::string::npos, "menu popup region passes a borrowed HRGN handle to SetWindowRgn");
    Require(releaseRegion != std::string::npos && setWindowRgn < releaseRegion, "menu popup region transfers HRGN ownership only after SetWindowRgn succeeds");
    Require(block.find("SetWindowRgn(hwnd, region.release(), FALSE)") == std::string::npos,
            "menu popup region does not release HRGN before checking SetWindowRgn");
}

void TestContextMenuModalLoopDismissesWhenRootPopupDisappearsBeforeWaiting()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Menu.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Menu source is readable for modal-loop root-popup guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t modalLoop = source.find("void RunMenuModalLoop");
    const size_t waitCall  = source.find("MsgWaitForMultipleObjectsEx", modalLoop);
    Require(modalLoop != std::string::npos && waitCall != std::string::npos && modalLoop < waitCall, "context-menu modal-loop idle wait block is found");

    const size_t currentRoot = source.rfind("MenuPopup* currentRoot = controller.GetRootPopup();", waitCall);
    Require(currentRoot != std::string::npos && modalLoop < currentRoot, "context-menu modal-loop revalidates the root popup before its idle wait");
    const std::string idleBlock = source.substr(currentRoot, waitCall - currentRoot);

    const size_t invalidWindowCheck = idleBlock.find("IsWindow(currentRoot->hwnd) == FALSE");
    const size_t dismissTrace       = idleBlock.find("menu.loop-dismiss-missing-root");
    const size_t dismissCall        = idleBlock.find("controller.Dismiss();");
    Require(invalidWindowCheck != std::string::npos, "context-menu modal-loop checks whether the root popup HWND is still valid");
    Require(dismissTrace != std::string::npos, "context-menu modal-loop traces dismissal when the root popup is gone");
    Require(dismissCall != std::string::npos && invalidWindowCheck < dismissCall,
            "context-menu modal-loop dismisses before waiting when the root popup is gone");
}

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

template <typename TPredicate>
bool WaitForContextMenuPopupState(HWND popupHwnd,
                                  TPredicate&& predicate,
                                  RedSalamander::DxUi::ContextMenuPopupDebugState& outState,
                                  std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popupHwnd, outState) && predicate(outState))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

bool WaitForContextMenuPopupItemRect(HWND popupHwnd,
                                     size_t itemIndex,
                                     D2D1_RECT_F& outRectDip,
                                     std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemRect(popupHwnd, itemIndex, outRectDip))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

bool WaitForContextMenuPopupItemLayout(HWND popupHwnd,
                                       size_t itemIndex,
                                       RedSalamander::DxUi::ContextMenuPopupItemLayoutDebugState& outLayout,
                                       std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemLayout(popupHwnd, itemIndex, outLayout))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

bool WaitForContextMenuPopupItemPaint(HWND popupHwnd,
                                      size_t itemIndex,
                                      RedSalamander::DxUi::ContextMenuPopupItemPaintDebugState& outPaint,
                                      std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemPaint(popupHwnd, itemIndex, outPaint))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

bool WaitForContextMenuPopupItemText(HWND popupHwnd,
                                     size_t itemIndex,
                                     std::wstring& outText,
                                     std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemText(popupHwnd, itemIndex, outText))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    outText.clear();
    return false;
}

bool WaitForContextMenuPopupBitmapCapture(HWND popupHwnd,
                                          RedSalamander::DxUi::WindowHostBitmapCapture& outCapture,
                                          std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugCaptureContextMenuPopupBitmap(popupHwnd, outCapture) && outCapture.widthPx > 0u && outCapture.heightPx > 0u &&
            ! outCapture.bgraPixels.empty())
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

[[nodiscard]] uint8_t GetCapturePixelAlpha(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, UINT xPx, UINT yPx) noexcept
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

    return capture.bgraPixels[base + 3u];
}

[[nodiscard]] uint32_t GetCapturePixelBgra(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, UINT xPx, UINT yPx) noexcept
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

[[nodiscard]] UINT DipToPixelForPopup(float dip, UINT dpi) noexcept
{
    return static_cast<UINT>((std::max)(0l, std::lround(static_cast<double>(dip) * static_cast<double>(dpi) / 96.0)));
}

[[nodiscard]] POINT ClientScreenPointForTest(HWND hwnd, LONG x, LONG y, const char* context)
{
    POINT point{x, y};
    Require(ClientToScreen(hwnd, &point) != FALSE, context);
    return point;
}

[[nodiscard]] LPARAM MousePointLParamForMenuSuite(LONG x, LONG y) noexcept
{
    return MAKELPARAM(static_cast<WORD>(static_cast<SHORT>(x)), static_cast<WORD>(static_cast<SHORT>(y)));
}

[[nodiscard]] LRESULT SendCapturedMouseMessageForMenuSuite(HWND hwnd, UINT message, WPARAM wParam, POINT screenPoint) noexcept
{
    RECT windowRect{};
    if (! hwnd || GetWindowRect(hwnd, &windowRect) == FALSE)
    {
        return 0;
    }

    const LPARAM lParam = MousePointLParamForMenuSuite(screenPoint.x - windowRect.left, screenPoint.y - windowRect.top);
    return SendMessageW(hwnd, message, wParam, lParam);
}

[[nodiscard]] bool PostCapturedMouseMessageForMenuSuite(HWND hwnd, UINT message, WPARAM wParam, POINT screenPoint) noexcept
{
    RECT windowRect{};
    if (! hwnd || GetWindowRect(hwnd, &windowRect) == FALSE)
    {
        return false;
    }

    const LPARAM lParam = MousePointLParamForMenuSuite(screenPoint.x - windowRect.left, screenPoint.y - windowRect.top);
    return PostMessageW(hwnd, message, wParam, lParam) != FALSE;
}

[[nodiscard]] bool SendClientMouseMoveForMenuSuite(HWND hwnd, LONG x, LONG y) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    static_cast<void>(SendMessageW(hwnd, WM_MOUSEMOVE, 0, MousePointLParamForMenuSuite(x, y)));
    return true;
}

[[nodiscard]] bool SendSettledClientMouseMoveForMenuSuite(HWND hwnd, LONG x, LONG y) noexcept
{
    // Deliver the client-space mouse move that production hover routing reads from the
    // message lParam. This previously also warped the live OS cursor, but routing never
    // samples the cursor (see DxUi.PointerInput.cpp), so the warp was incidental and
    // intrusive -- it is omitted to keep the menu suite from disturbing the interactive
    // user's pointer.
    return SendClientMouseMoveForMenuSuite(hwnd, x, y);
}

void DrainPendingMouseMessagesForMenuSuite() noexcept
{
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE) != FALSE)
    {
    }
}

[[nodiscard]] bool WaitForPendingSubmenuCloseTimerForMenuSuite(HWND popupHwnd, RedSalamander::DxUi::ContextMenuPopupDebugState& outState)
{
    return WaitForContextMenuPopupState(popupHwnd, [](const RedSalamander::DxUi::ContextMenuPopupDebugState& state) noexcept {
        return state.hoverTimerActive && state.hoverTimerPendingClose;
    }, outState);
}

[[nodiscard]] bool WaitForNoSubmenuHoverTimerForMenuSuite(HWND popupHwnd, RedSalamander::DxUi::ContextMenuPopupDebugState& outState)
{
    return WaitForContextMenuPopupState(
        popupHwnd, [](const RedSalamander::DxUi::ContextMenuPopupDebugState& state) noexcept { return ! state.hoverTimerActive; }, outState);
}

[[nodiscard]] bool FirePendingSubmenuHoverTimerForMenuSuite(HWND popupHwnd) noexcept
{
    return RedSalamander::DxUi::DebugFireContextMenuPopupHoverTimer(popupHwnd);
}

HWND FindOwnedContextMenuPopupWindowByFirstItemText(HWND ownerHwnd, std::wstring_view firstItemText);

[[nodiscard]] bool WaitForSubmenuOpenAfterHoverForMenuSuite(HWND ownerHwnd,
                                                            HWND popupHwnd,
                                                            size_t itemIndex,
                                                            std::wstring_view submenuFirstItemText,
                                                            HWND& submenuHwnd,
                                                            RedSalamander::DxUi::ContextMenuPopupDebugState& outState,
                                                            std::chrono::milliseconds timeout = std::chrono::milliseconds(1200))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    bool timerPosted    = false;
    do
    {
        submenuHwnd = FindOwnedContextMenuPopupWindowByFirstItemText(ownerHwnd, submenuFirstItemText);
        if (submenuHwnd)
        {
            return true;
        }

        if (! timerPosted && RedSalamander::DxUi::DebugGetContextMenuPopupState(popupHwnd, outState) && outState.hoverTimerActive &&
            outState.hoverTimerPendingOpen && outState.hoverTimerItemIndex.has_value() && outState.hoverTimerItemIndex.value() == itemIndex)
        {
            if (! FirePendingSubmenuHoverTimerForMenuSuite(popupHwnd))
            {
                return false;
            }
            timerPosted = true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    submenuHwnd = FindOwnedContextMenuPopupWindowByFirstItemText(ownerHwnd, submenuFirstItemText);
    return submenuHwnd != nullptr;
}

[[nodiscard]] bool WaitForSubmenuClosedAfterHoverForMenuSuite(HWND popupHwnd,
                                                              HWND submenuHwnd,
                                                              RedSalamander::DxUi::ContextMenuPopupDebugState& outState,
                                                              std::chrono::milliseconds timeout = std::chrono::milliseconds(1200))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    bool timerPosted    = false;
    do
    {
        if (IsWindow(submenuHwnd) == FALSE)
        {
            return true;
        }

        if (! timerPosted && RedSalamander::DxUi::DebugGetContextMenuPopupState(popupHwnd, outState) && outState.hoverTimerActive &&
            outState.hoverTimerPendingClose)
        {
            if (! FirePendingSubmenuHoverTimerForMenuSuite(popupHwnd))
            {
                return false;
            }
            timerPosted = true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return IsWindow(submenuHwnd) == FALSE;
}

RedSalamander::DxUi::WindowHostBitmapCapture CropWindowHostBitmapCaptureForTest(const RedSalamander::DxUi::WindowHostBitmapCapture& source,
                                                                                const RECT& cropRect)
{
    RedSalamander::DxUi::WindowHostBitmapCapture capture{};
    const LONG widthPx        = cropRect.right - cropRect.left;
    const LONG heightPx       = cropRect.bottom - cropRect.top;
    const LONG sourceWidthPx  = static_cast<LONG>(source.widthPx);
    const LONG sourceHeightPx = static_cast<LONG>(source.heightPx);
    const bool cropIsInBounds =
        widthPx > 0 && heightPx > 0 && cropRect.left >= 0 && cropRect.top >= 0 && cropRect.right <= sourceWidthPx && cropRect.bottom <= sourceHeightPx;
    const std::string cropBoundsMessage = std::format("popup backdrop crop stays inside the owner capture crop=({},{} {}x{}) source={}x{}",
                                                      cropRect.left,
                                                      cropRect.top,
                                                      widthPx,
                                                      heightPx,
                                                      sourceWidthPx,
                                                      sourceHeightPx);
    Require(cropIsInBounds, cropBoundsMessage.c_str());

    capture.widthPx  = static_cast<UINT>(widthPx);
    capture.heightPx = static_cast<UINT>(heightPx);
    capture.bgraPixels.resize(static_cast<size_t>(capture.widthPx) * static_cast<size_t>(capture.heightPx) * 4u);

    const size_t destinationStride = static_cast<size_t>(capture.widthPx) * 4u;
    const size_t sourceStride      = static_cast<size_t>(source.widthPx) * 4u;
    for (LONG y = 0; y < heightPx; ++y)
    {
        const size_t sourceOffset      = (static_cast<size_t>(cropRect.top + y) * sourceStride) + (static_cast<size_t>(cropRect.left) * 4u);
        const size_t destinationOffset = static_cast<size_t>(y) * destinationStride;
        std::copy_n(source.bgraPixels.data() + sourceOffset, destinationStride, capture.bgraPixels.data() + destinationOffset);
    }

    return capture;
}

[[nodiscard]] RECT FindOpaqueBoundsInCapture(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, uint8_t alphaThreshold = 8u) noexcept
{
    RECT bounds{static_cast<LONG>(capture.widthPx), static_cast<LONG>(capture.heightPx), 0, 0};
    bool foundOpaque = false;

    for (UINT y = 0u; y < capture.heightPx; ++y)
    {
        for (UINT x = 0u; x < capture.widthPx; ++x)
        {
            if (GetCapturePixelAlpha(capture, x, y) <= alphaThreshold)
            {
                continue;
            }

            bounds.left   = (std::min)(bounds.left, static_cast<LONG>(x));
            bounds.top    = (std::min)(bounds.top, static_cast<LONG>(y));
            bounds.right  = (std::max)(bounds.right, static_cast<LONG>(x + 1u));
            bounds.bottom = (std::max)(bounds.bottom, static_cast<LONG>(y + 1u));
            foundOpaque   = true;
        }
    }

    return foundOpaque ? bounds : RECT{};
}

[[nodiscard]] RECT ComputeRightStripSampleRect(const RECT& opaqueBounds) noexcept
{
    const LONG widthPx  = opaqueBounds.right - opaqueBounds.left;
    const LONG heightPx = opaqueBounds.bottom - opaqueBounds.top;
    if (widthPx <= 24 || heightPx <= 24)
    {
        return RECT{};
    }

    const LONG horizontalInset = (std::max)(10l, widthPx / 12l);
    const LONG verticalInset   = (std::max)(10l, heightPx / 6l);
    const LONG sampleLeft      = opaqueBounds.left + (widthPx * 5l / 8l);
    const LONG sampleRight     = opaqueBounds.right - horizontalInset;
    const LONG sampleTop       = opaqueBounds.top + verticalInset;
    const LONG sampleBottom    = opaqueBounds.bottom - verticalInset;

    return RECT{sampleLeft, sampleTop, (std::max)(sampleLeft + 2l, sampleRight), (std::max)(sampleTop + 2l, sampleBottom)};
}

[[nodiscard]] uint64_t ComputeAverageAdjacentRgbDelta(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, const RECT& sampleRect) noexcept
{
    if (sampleRect.right - sampleRect.left < 2 || sampleRect.bottom - sampleRect.top < 1)
    {
        return 0u;
    }

    uint64_t accumulatedDelta = 0u;
    uint64_t sampleCount      = 0u;

    const LONG clampedLeft   = (std::clamp)(sampleRect.left, 0l, static_cast<LONG>(capture.widthPx));
    const LONG clampedTop    = (std::clamp)(sampleRect.top, 0l, static_cast<LONG>(capture.heightPx));
    const LONG clampedRight  = (std::clamp)(sampleRect.right, clampedLeft, static_cast<LONG>(capture.widthPx));
    const LONG clampedBottom = (std::clamp)(sampleRect.bottom, clampedTop, static_cast<LONG>(capture.heightPx));

    for (LONG y = clampedTop; y < clampedBottom; ++y)
    {
        for (LONG x = clampedLeft + 1; x < clampedRight; ++x)
        {
            const size_t currentBase  = (static_cast<size_t>(y) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(x)) * 4u;
            const size_t previousBase = currentBase - 4u;
            accumulatedDelta += static_cast<uint64_t>(
                std::abs(static_cast<int>(capture.bgraPixels[currentBase + 0u]) - static_cast<int>(capture.bgraPixels[previousBase + 0u])));
            accumulatedDelta += static_cast<uint64_t>(
                std::abs(static_cast<int>(capture.bgraPixels[currentBase + 1u]) - static_cast<int>(capture.bgraPixels[previousBase + 1u])));
            accumulatedDelta += static_cast<uint64_t>(
                std::abs(static_cast<int>(capture.bgraPixels[currentBase + 2u]) - static_cast<int>(capture.bgraPixels[previousBase + 2u])));
            sampleCount += 3u;
        }
    }

    return sampleCount == 0u ? 0u : (accumulatedDelta / sampleCount);
}

[[nodiscard]] uint64_t ComputeAverageAbsoluteRgbDeltaBetweenCaptures(const RedSalamander::DxUi::WindowHostBitmapCapture& lhs,
                                                                     const RedSalamander::DxUi::WindowHostBitmapCapture& rhs,
                                                                     const RECT& sampleRect) noexcept
{
    if (lhs.widthPx != rhs.widthPx || lhs.heightPx != rhs.heightPx)
    {
        return 0u;
    }

    uint64_t accumulatedDelta = 0u;
    uint64_t sampleCount      = 0u;

    const LONG clampedLeft   = (std::clamp)(sampleRect.left, 0l, static_cast<LONG>(lhs.widthPx));
    const LONG clampedTop    = (std::clamp)(sampleRect.top, 0l, static_cast<LONG>(lhs.heightPx));
    const LONG clampedRight  = (std::clamp)(sampleRect.right, clampedLeft, static_cast<LONG>(lhs.widthPx));
    const LONG clampedBottom = (std::clamp)(sampleRect.bottom, clampedTop, static_cast<LONG>(lhs.heightPx));

    for (LONG y = clampedTop; y < clampedBottom; ++y)
    {
        for (LONG x = clampedLeft; x < clampedRight; ++x)
        {
            const size_t base = (static_cast<size_t>(y) * static_cast<size_t>(lhs.widthPx) + static_cast<size_t>(x)) * 4u;
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(lhs.bgraPixels[base + 0u]) - static_cast<int>(rhs.bgraPixels[base + 0u])));
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(lhs.bgraPixels[base + 1u]) - static_cast<int>(rhs.bgraPixels[base + 1u])));
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(lhs.bgraPixels[base + 2u]) - static_cast<int>(rhs.bgraPixels[base + 2u])));
            sampleCount += 3u;
        }
    }

    return sampleCount == 0u ? 0u : (accumulatedDelta / sampleCount);
}

[[nodiscard]] uint64_t ComputeAverageAbsoluteRgbDeltaToColor(const RedSalamander::DxUi::WindowHostBitmapCapture& capture,
                                                             const RECT& sampleRect,
                                                             const D2D1_COLOR_F& color) noexcept
{
    uint64_t accumulatedDelta = 0u;
    uint64_t sampleCount      = 0u;

    const LONG clampedLeft   = (std::clamp)(sampleRect.left, 0l, static_cast<LONG>(capture.widthPx));
    const LONG clampedTop    = (std::clamp)(sampleRect.top, 0l, static_cast<LONG>(capture.heightPx));
    const LONG clampedRight  = (std::clamp)(sampleRect.right, clampedLeft, static_cast<LONG>(capture.widthPx));
    const LONG clampedBottom = (std::clamp)(sampleRect.bottom, clampedTop, static_cast<LONG>(capture.heightPx));

    const int expectedBlue  = static_cast<int>((std::clamp)(color.b, 0.0f, 1.0f) * 255.0f + 0.5f);
    const int expectedGreen = static_cast<int>((std::clamp)(color.g, 0.0f, 1.0f) * 255.0f + 0.5f);
    const int expectedRed   = static_cast<int>((std::clamp)(color.r, 0.0f, 1.0f) * 255.0f + 0.5f);

    for (LONG y = clampedTop; y < clampedBottom; ++y)
    {
        for (LONG x = clampedLeft; x < clampedRight; ++x)
        {
            const size_t base = (static_cast<size_t>(y) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(x)) * 4u;
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(capture.bgraPixels[base + 0u]) - expectedBlue));
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(capture.bgraPixels[base + 1u]) - expectedGreen));
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(capture.bgraPixels[base + 2u]) - expectedRed));
            sampleCount += 3u;
        }
    }

    return sampleCount == 0u ? 0u : (accumulatedDelta / sampleCount);
}

[[nodiscard]] D2D1_COLOR_F ResolveExpectedAcrylicMenuSlabColor(const RedSalamander::DxUi::ThemePalette& theme) noexcept
{
    return RedSalamander::DxUi::BlendColor(theme.overlayBackground, theme.headerHovered, theme.dark ? 0.14f : 0.10f);
}

HWND FindOwnedContextMenuPopupWindow(HWND ownerHwnd)
{
    HWND popupHwnd               = nullptr;
    const DWORD currentProcessId = GetCurrentProcessId();
    while ((popupHwnd = FindWindowExW(nullptr, popupHwnd, L"DxUi_ContextMenu", nullptr)) != nullptr)
    {
        DWORD popupProcessId = 0;
        static_cast<void>(GetWindowThreadProcessId(popupHwnd, &popupProcessId));
        if (popupProcessId != currentProcessId)
        {
            continue;
        }

        if (GetWindow(popupHwnd, GW_OWNER) == ownerHwnd)
        {
            return popupHwnd;
        }
    }

    return nullptr;
}

HWND WaitForOwnedContextMenuPopupWindow(HWND ownerHwnd, std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (HWND popupHwnd = FindOwnedContextMenuPopupWindow(ownerHwnd))
        {
            return popupHwnd;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return nullptr;
}

bool WaitForWindowDestroyed(HWND hwnd, std::chrono::milliseconds timeout);

void DismissOwnedContextMenuPopupChain(HWND ownerHwnd) noexcept
{
    for (int attempt = 0; attempt < 16; ++attempt)
    {
        if (HWND popupHwnd = FindOwnedContextMenuPopupWindow(ownerHwnd))
        {
            DWORD_PTR unused = 0;
            static_cast<void>(SendMessageTimeoutW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0, SMTO_ABORTIFHUNG, 1000u, &unused));
            static_cast<void>(WaitForWindowDestroyed(popupHwnd, std::chrono::milliseconds(120)));
            continue;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(30));
    }
}

[[nodiscard]] bool SendMenuKeyForMenuSuite(HWND popupHwnd, WPARAM virtualKey) noexcept
{
    if (! popupHwnd || IsWindow(popupHwnd) == FALSE)
    {
        return false;
    }

    DWORD_PTR unused = 0;
    return SendMessageTimeoutW(popupHwnd, WM_KEYDOWN, virtualKey, 0, SMTO_ABORTIFHUNG, 1000u, &unused) != 0;
}

HWND FindOwnedContextMenuPopupWindowByFirstItemText(HWND ownerHwnd, std::wstring_view firstItemText)
{
    HWND popupHwnd               = nullptr;
    const DWORD currentProcessId = GetCurrentProcessId();
    while ((popupHwnd = FindWindowExW(nullptr, popupHwnd, L"DxUi_ContextMenu", nullptr)) != nullptr)
    {
        DWORD popupProcessId = 0;
        static_cast<void>(GetWindowThreadProcessId(popupHwnd, &popupProcessId));
        if (popupProcessId != currentProcessId)
        {
            continue;
        }

        if (GetWindow(popupHwnd, GW_OWNER) != ownerHwnd)
        {
            continue;
        }

        std::wstring popupFirstItemText;
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemText(popupHwnd, 0u, popupFirstItemText) && popupFirstItemText == firstItemText)
        {
            return popupHwnd;
        }
    }

    return nullptr;
}

HWND WaitForOwnedContextMenuPopupWindowByFirstItemText(HWND ownerHwnd,
                                                       std::wstring_view firstItemText,
                                                       std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (HWND popupHwnd = FindOwnedContextMenuPopupWindowByFirstItemText(ownerHwnd, firstItemText))
        {
            return popupHwnd;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return nullptr;
}

bool WaitForWindowDestroyed(HWND hwnd, std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (IsWindow(hwnd) == FALSE)
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return IsWindow(hwnd) == FALSE;
}

bool WaitForFocusedWindow(HWND hwnd, std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    return WaitForDxUiThreadFocus(hwnd, static_cast<DWORD>(timeout.count()));
}

void TestContextMenuDebugStateProbeBoundsWedgedWindowThread()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 220, SWP_NOZORDER);
    if (! TryActivateDxUiTestWindow(ownerWindow.Hwnd()))
    {
        SkipDxUiTest("DxUi menu debug-state timeout requires an interactive desktop");
        return;
    }

    wil::unique_event_nothrow handlerEntered;
    handlerEntered.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(handlerEntered != nullptr, "menu debug-state timeout test creates the handler-entered event");
    wil::unique_event_nothrow releaseHandler;
    releaseHandler.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(releaseHandler != nullptr, "menu debug-state timeout test creates the release event");

    std::string driverFailure;
    bool unwedgedStateObserved = false;
    bool wedgedProbeFailed     = false;
    std::chrono::milliseconds wedgedElapsed{};
    std::thread releaser([&]
    {
        if (WaitForSingleObject(handlerEntered.get(), 3000u) == WAIT_OBJECT_0)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(1200));
        }
        static_cast<void>(SetEvent(releaseHandler.get()));
    });
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });
        const HWND popupHwnd     = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"State probe");
        if (! popupHwnd)
        {
            driverFailure = "menu debug-state timeout popup appears";
            return;
        }

        ContextMenuPopupDebugState state{};
        unwedgedStateObserved = DebugGetContextMenuPopupState(popupHwnd, state) && state.itemTexts.size() == 1u && state.itemTexts[0u] == L"State probe";
        if (! unwedgedStateObserved)
        {
            driverFailure = "unwedged menu debug-state probe returns the correct popup state";
            return;
        }

        DebugSetContextMenuStateProbeStallForTest(handlerEntered.get(), releaseHandler.get());
        const auto started = std::chrono::steady_clock::now();
        ContextMenuPopupDebugState wedgedState{};
        wedgedProbeFailed = ! DebugGetContextMenuPopupState(popupHwnd, wedgedState);
        wedgedElapsed     = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started);
    });

    const std::vector<MenuFlyoutItem> items{{.text = L"State probe", .enabled = true, .commandId = 91501}};
    const POINT menuAnchor = ClientScreenPointForTest(ownerWindow.Hwnd(), 24, 60, "menu debug-state timeout anchor converts to screen coordinates");
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), menuAnchor, items, ownerWindow.Host().GetTheme());
    driver.join();
    releaser.join();
    DebugSetContextMenuStateProbeStallForTest(nullptr, nullptr);

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(unwedgedStateObserved, "unwedged menu debug-state probe returns the correct popup state");
    Require(wedgedProbeFailed, "wedged menu debug-state probe returns failure");
    Require(wedgedElapsed >= std::chrono::milliseconds(900) && wedgedElapsed < std::chrono::milliseconds(1800),
            "wedged menu debug-state probe returns within the bounded test timeout");
    Require(! result.has_value(), "dismissed menu debug-state timeout popup returns no command");
}

void TestEmbeddedViewerContextMenuNativeConversionFiltersStandaloneCommands()
{
    using namespace RedSalamander::DxUi;

    wil::unique_hmenu menu{CreatePopupMenu()};
    Require(menu != nullptr, "embedded viewer context-menu conversion creates a native popup menu");

    wil::unique_hmenu standaloneSubmenu{CreatePopupMenu()};
    Require(standaloneSubmenu != nullptr, "embedded viewer context-menu conversion creates a standalone-only submenu");
    Require(AppendMenuW(standaloneSubmenu.get(), MF_STRING, 91701u, L"&Standalone action\tAlt+S") != FALSE,
            "embedded viewer context-menu conversion populates a submenu that should become empty after filtering");

    Require(AppendMenuW(menu.get(), MF_SEPARATOR, 0, nullptr) != FALSE, "embedded viewer context-menu conversion can start with a separator");
    Require(AppendMenuW(menu.get(), MF_STRING, 91702u, L"&Open standalone\tCtrl+O") != FALSE,
            "embedded viewer context-menu conversion populates a standalone-only command");
    Require(AppendMenuW(menu.get(), MF_SEPARATOR, 0, nullptr) != FALSE, "embedded viewer context-menu conversion can trim interior separators");
    Require(AppendMenuW(menu.get(), MF_STRING, 91703u, L"&Copy path\tCtrl+C") != FALSE,
            "embedded viewer context-menu conversion populates an embedded-safe command");
    Require(AppendMenuW(menu.get(), MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(standaloneSubmenu.get()), L"&Standalone") != FALSE,
            "embedded viewer context-menu conversion attaches the standalone-only submenu");
    static_cast<void>(standaloneSubmenu.release());
    Require(AppendMenuW(menu.get(), MF_SEPARATOR, 0, nullptr) != FALSE, "embedded viewer context-menu conversion can trim trailing separators");
    Require(AppendMenuW(menu.get(), MF_STRING | MF_CHECKED, 91704u, L"Show &metadata\tF4") != FALSE,
            "embedded viewer context-menu conversion preserves checked embedded-safe commands");

    static constexpr std::array<int, 2> kPreviewContextMenuExcludedCommandIds{{91701, 91702}};
    const NativeMenuFlyoutOptions options{
        .includeAcceleratorText = false,
        .omitEmptySubmenus      = true,
        .trimSeparators         = true,
        .excludedCommandIds     = kPreviewContextMenuExcludedCommandIds,
    };

    const std::vector<MenuFlyoutItem> items = ConvertNativeHMenuToFlyoutItems(menu.get(), options);
    Require(items.size() == 3u, "embedded viewer context-menu conversion removes standalone-only commands, empty submenus, and outer separator runs");
    Require(items[0].text == L"Copy path", "embedded viewer context-menu conversion strips mnemonics from retained command text");
    Require(items[0].commandId == 91703, "embedded viewer context-menu conversion preserves retained command IDs");
    Require(items[0].acceleratorText.empty(), "embedded viewer context-menu conversion suppresses standalone keyboard shortcut text");
    Require(items[1].kind == MenuItemKind::Separator, "embedded viewer context-menu conversion preserves one intentional separator between retained groups");
    Require(items[2].text == L"Show metadata", "embedded viewer context-menu conversion keeps the second embedded-safe command");
    Require(items[2].kind == MenuItemKind::Toggle && items[2].checked, "embedded viewer context-menu conversion preserves checked state");
}

void RunMenuDismissalKeyScenario(UINT message, WPARAM virtualKey, const char* appearExpectation, const char* dismissExpectation, const char* focusExpectation)
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOW);
    ownerWindow.PumpMessages();

    wil::unique_hwnd focusedChild{CreateWindowExW(0,
                                                  L"BUTTON",
                                                  L"DismissalTarget",
                                                  WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                                  12,
                                                  80,
                                                  140,
                                                  28,
                                                  ownerWindow.Hwnd(),
                                                  nullptr,
                                                  GetModuleHandleW(nullptr),
                                                  nullptr)};
    Require(focusedChild != nullptr, "dismissal validation creates a focusable child control");

    wil::unique_hmenu menu{CreateMenu()};
    Require(menu != nullptr, "dismissal validation creates a top-level native menu");

    wil::unique_hmenu filePopup{CreatePopupMenu()};
    Require(filePopup != nullptr, "dismissal validation creates a native popup menu");
    Require(AppendMenuW(filePopup.get(), MF_STRING, 3701u, L"&Open") != FALSE, "dismissal validation populates the native popup menu");
    Require(AppendMenuW(filePopup.get(), MF_STRING, 3702u, L"&Close") != FALSE, "dismissal validation populates a second native popup item");
    Require(AppendMenuW(menu.get(), MF_POPUP, reinterpret_cast<UINT_PTR>(filePopup.get()), L"&File") != FALSE,
            "dismissal validation attaches the popup menu to the native menu bar");
    static_cast<void>(filePopup.release());

    NativeMenuBarHost menuBarHost;
    Require(menuBarHost.Attach(GetModuleHandleW(nullptr), ownerWindow.Hwnd(), menu.get()), "dismissal validation attaches a native menu bar host");
    ownerWindow.PumpMessages();

    if (! TryActivateDxUiTestWindow(ownerWindow.Hwnd()))
    {
        SkipDxUiTest("DxUi menu popup requires an interactive desktop for native menu-bar focus dismissal");
        return;
    }
    static_cast<void>(SetFocus(focusedChild.get()));
    Require(WaitForFocusedWindow(focusedChild.get()), "dismissal validation starts with focus on the child control");

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = appearExpectation;
            return;
        }

        PostMessageW(popupHwnd, message, virtualKey, 0);
        if (! WaitForWindowDestroyed(popupHwnd))
        {
            driverFailure = dismissExpectation;
        }
    });

    Require(menuBarHost.FocusFirstItem(), "dismissal validation enters native menu-bar mode");
    static_cast<void>(SendMessageW(menuBarHost.GetHwnd(), WM_KEYDOWN, VK_DOWN, 0));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(WaitForFocusedWindow(focusedChild.get()), focusExpectation);
}

void TestSplitButtonContextMenuSentMouseMessagesHoverAndOutsideDismiss()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER);
    if (! TryActivateDxUiTestWindow(ownerWindow.Hwnd()))
    {
        SkipDxUiTest("DxUi menu popup requires an interactive desktop for split-button outside-dismiss routing");
        return;
    }

    const std::vector<MenuFlyoutItem> items{
        {.text = L"Find Now", .enabled = true, .commandId = 4101},
        {.text = L"Refine - Intersect with Found Items", .enabled = true, .commandId = 4102},
        {.text = L"Refine - Subtract from Found Items", .enabled = true, .commandId = 4103},
        {.text = L"Append to Found Items", .enabled = true, .commandId = 4104},
    };
    ownerWindow.PumpMessages();
    const POINT menuAnchor = ClientScreenPointForTest(ownerWindow.Hwnd(), 24, 60, "split-button-style menu anchor converts to screen coordinates");

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Find Now");
        if (! popupHwnd)
        {
            driverFailure = "split-button opens the owned DxUi context menu";
            return;
        }

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f;
        }, popupState))
        {
            driverFailure = "split-button popup exposes geometry before live pointer routing validation";
            return;
        }

        D2D1_RECT_F refineRectDip{};
        if (! WaitForContextMenuPopupItemRect(popupHwnd, 1u, refineRectDip))
        {
            driverFailure = "split-button popup exposes the Refine row bounds";
            return;
        }

        POINT refineCenter{
            static_cast<LONG>(std::lround((refineRectDip.left + refineRectDip.right) * 0.5f * static_cast<float>(popupState.dpi) /
                                          static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
            static_cast<LONG>(std::lround((refineRectDip.top + refineRectDip.bottom) * 0.5f * static_cast<float>(popupState.dpi) /
                                          static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
        };
        if (ClientToScreen(popupHwnd, &refineCenter) == FALSE)
        {
            driverFailure = "split-button popup Refine row converts to screen coordinates";
            return;
        }

        // Deliver the hover via the popup's window proc (the production routing input);
        // no live cursor traversal is needed because routing reads the delivered point.
        static_cast<void>(SendCapturedMouseMessageForMenuSuite(popupHwnd, WM_MOUSEMOVE, 0, refineCenter));

        ContextMenuPopupItemPaintDebugState paintState{};
        const auto hoverDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(800);
        bool hoverObserved       = false;
        const auto observeHover  = [&]() noexcept
        {
            ContextMenuPopupDebugState hoverState{};
            return DebugGetContextMenuPopupState(popupHwnd, hoverState) && hoverState.hoveredIndex == std::optional<size_t>{1u} &&
                   DebugGetContextMenuPopupItemPaint(popupHwnd, 1u, paintState) && paintState.usesHighlightFill;
        };
        hoverObserved = observeHover();
        do
        {
            if (hoverObserved)
            {
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(50));
            hoverObserved = observeHover();
        } while (std::chrono::steady_clock::now() < hoverDeadline);
        if (! hoverObserved)
        {
            return;
        }

        POINT outsidePoint = ownerWindow.Host().DipPointToScreenPoint(D2D1::Point2F(32.0f, 196.0f));
        RECT popupRect{};
        if (GetWindowRect(popupHwnd, &popupRect) != FALSE && PtInRect(&popupRect, outsidePoint) != FALSE)
        {
            outsidePoint = ownerWindow.Host().DipPointToScreenPoint(D2D1::Point2F(388.0f, 196.0f));
        }
        static_cast<void>(SendCapturedMouseMessageForMenuSuite(popupHwnd, WM_LBUTTONDOWN, MK_LBUTTON, outsidePoint));
        if (! WaitForWindowDestroyed(popupHwnd))
        {
            driverFailure = "outside click dismisses the split-button popup";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), menuAnchor, items, ownerWindow.Host().GetTheme());
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "outside-dismissed split-button-style popup returns no command");
}

void TestSplitButtonContextMenuSentMouseMessagesHoverAndInvokeImmediately()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 140, 140, 460, 280, SWP_NOZORDER);
    if (! TryActivateDxUiTestWindow(ownerWindow.Hwnd()))
    {
        SkipDxUiTest("DxUi menu popup requires an interactive desktop for split-button invoke routing");
        return;
    }

    const std::vector<MenuFlyoutItem> items{
        {.text = L"Find Now", .enabled = true, .commandId = 4201},
        {.text = L"Refine - Intersect with Found Items", .enabled = true, .commandId = 4202},
        {.text = L"Refine - Subtract from Found Items", .enabled = true, .commandId = 4203},
        {.text = L"Append to Found Items", .enabled = true, .commandId = 4204},
    };
    ownerWindow.PumpMessages();
    const POINT menuAnchor = ClientScreenPointForTest(ownerWindow.Hwnd(), 24, 60, "sent-message menu anchor converts to screen coordinates");

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Find Now");
        if (! popupHwnd)
        {
            driverFailure = "sent-message split-button popup opens";
            return;
        }

        ContextMenuPopupDebugState popupState{};
        D2D1_RECT_F refineRectDip{};
        if (! WaitForContextMenuPopupState(popupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupItemRect(popupHwnd, 1u, refineRectDip))
        {
            driverFailure = "sent-message split-button popup exposes the Refine row";
            return;
        }

        POINT refineCenter{
            static_cast<LONG>(std::lround((refineRectDip.left + refineRectDip.right) * 0.5f * static_cast<float>(popupState.dpi) /
                                          static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
            static_cast<LONG>(std::lround((refineRectDip.top + refineRectDip.bottom) * 0.5f * static_cast<float>(popupState.dpi) /
                                          static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
        };
        if (ClientToScreen(popupHwnd, &refineCenter) == FALSE)
        {
            driverFailure = "sent-message split-button Refine row converts to screen coordinates";
            return;
        }

        static_cast<void>(SendCapturedMouseMessageForMenuSuite(popupHwnd, WM_MOUSEMOVE, 0, refineCenter));

        ContextMenuPopupItemPaintDebugState paintState{};
        const auto hoverDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(800);
        bool hoverObserved       = false;
        const auto observeHover  = [&]() noexcept
        {
            ContextMenuPopupDebugState hoverState{};
            return DebugGetContextMenuPopupState(popupHwnd, hoverState) && hoverState.hoveredIndex == std::optional<size_t>{1u} &&
                   DebugGetContextMenuPopupItemPaint(popupHwnd, 1u, paintState) && paintState.usesHighlightFill;
        };
        hoverObserved = observeHover();
        do
        {
            if (hoverObserved)
            {
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(50));
            hoverObserved = observeHover();
        } while (std::chrono::steady_clock::now() < hoverDeadline);
        if (! hoverObserved)
        {
            ContextMenuPopupDebugState finalState{};
            static_cast<void>(DebugGetContextMenuPopupState(popupHwnd, finalState));
            RECT popupRect{};
            static_cast<void>(GetWindowRect(popupHwnd, &popupRect));
            driverFailure = std::format("sent WM_MOUSEMOVE over the split-button popup Refine row produces visible hover highlight "
                                        "(hovered={}, keyboard={}, center=({},{}), popup=({},{} {}x{}), row=({:.1f},{:.1f},{:.1f},{:.1f}), renders={})",
                                        finalState.hoveredIndex.has_value() ? static_cast<long long>(finalState.hoveredIndex.value()) : -1ll,
                                        finalState.keyboardIndex.has_value() ? static_cast<long long>(finalState.keyboardIndex.value()) : -1ll,
                                        refineCenter.x,
                                        refineCenter.y,
                                        popupRect.left,
                                        popupRect.top,
                                        popupRect.right - popupRect.left,
                                        popupRect.bottom - popupRect.top,
                                        refineRectDip.left,
                                        refineRectDip.top,
                                        refineRectDip.right,
                                        refineRectDip.bottom,
                                        finalState.renderCount);
            return;
        }

        static_cast<void>(SendCapturedMouseMessageForMenuSuite(popupHwnd, WM_LBUTTONDOWN, MK_LBUTTON, refineCenter));
        static_cast<void>(SendCapturedMouseMessageForMenuSuite(popupHwnd, WM_LBUTTONUP, 0, refineCenter));

        if (! WaitForWindowDestroyed(popupHwnd, std::chrono::milliseconds(800)))
        {
            driverFailure = "sent split-button menu item click closes the popup immediately";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), menuAnchor, items, ownerWindow.Host().GetTheme());
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(result == std::optional<int>{4202}, "sent split-button-style menu item click invokes immediately without a later activation poke");
}

constexpr UINT kMenuOwnerMessageFloodTestMessage    = WM_APP + 0x53Du;
constexpr UINT_PTR kMenuOwnerMessageFloodSubclassId = 0x53Du;

LRESULT CALLBACK MenuOwnerMessageFloodTestSubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR) noexcept
{
    if (message == kMenuOwnerMessageFloodTestMessage)
    {
        Sleep(1);
        return 0;
    }

    return DefSubclassProc(hwnd, message, wParam, lParam);
}

void TestSplitButtonContextMenuOwnerMessageFloodDoesNotStarvePointerInput()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 140, 140, 460, 280, SWP_NOZORDER);
    if (! TryActivateDxUiTestWindow(ownerWindow.Hwnd()))
    {
        SkipDxUiTest("DxUi menu popup requires an interactive desktop for owner-message-flood routing");
        return;
    }

    Require(SetWindowSubclass(ownerWindow.Hwnd(), MenuOwnerMessageFloodTestSubclassProc, kMenuOwnerMessageFloodSubclassId, 0) != FALSE,
            "owner-message-flood validation subclasses the owner window");
    const auto restoreOwnerWndProc = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(RemoveWindowSubclass(ownerWindow.Hwnd(), MenuOwnerMessageFloodTestSubclassProc, kMenuOwnerMessageFloodSubclassId));
        MSG drainMessage{};
        while (PeekMessageW(&drainMessage, ownerWindow.Hwnd(), kMenuOwnerMessageFloodTestMessage, kMenuOwnerMessageFloodTestMessage, PM_REMOVE) != FALSE)
        {
        }
    });

    const std::vector<MenuFlyoutItem> items{
        {.text = L"Find Now", .enabled = true, .commandId = 4201},
        {.text = L"Refine - Intersect with Found Items", .enabled = true, .commandId = 4202},
        {.text = L"Refine - Subtract from Found Items", .enabled = true, .commandId = 4203},
        {.text = L"Append to Found Items", .enabled = true, .commandId = 4204},
    };
    ownerWindow.PumpMessages();
    const POINT menuAnchor = ClientScreenPointForTest(ownerWindow.Hwnd(), 24, 60, "owner-message-flood menu anchor converts to screen coordinates");

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Find Now");
        if (! popupHwnd)
        {
            driverFailure = "owner-message-flood split-button popup opens";
            return;
        }

        ContextMenuPopupDebugState popupState{};
        D2D1_RECT_F refineRectDip{};
        if (! WaitForContextMenuPopupState(popupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupItemRect(popupHwnd, 1u, refineRectDip))
        {
            driverFailure = "owner-message-flood split-button popup exposes the Refine row";
            return;
        }

        POINT refineCenter{
            static_cast<LONG>(std::lround((refineRectDip.left + refineRectDip.right) * 0.5f * static_cast<float>(popupState.dpi) /
                                          static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
            static_cast<LONG>(std::lround((refineRectDip.top + refineRectDip.bottom) * 0.5f * static_cast<float>(popupState.dpi) /
                                          static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
        };
        if (ClientToScreen(popupHwnd, &refineCenter) == FALSE)
        {
            driverFailure = "owner-message-flood split-button Refine row converts to screen coordinates";
            return;
        }
        static constexpr int kFloodMessageCount = 2000;
        for (int i = 0; i < kFloodMessageCount; ++i)
        {
            if (PostMessageW(ownerWindow.Hwnd(), kMenuOwnerMessageFloodTestMessage, 0, 0) == FALSE)
            {
                driverFailure = "owner-message-flood posts owner-window traffic";
                return;
            }
        }

        if (! PostCapturedMouseMessageForMenuSuite(popupHwnd, WM_MOUSEMOVE, 0, refineCenter))
        {
            driverFailure = "owner-message-flood posts popup mouse move";
            return;
        }

        ContextMenuPopupItemPaintDebugState paintState{};
        const auto hoverDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(800);
        bool hoverObserved       = false;
        const auto observeHover  = [&]() noexcept
        {
            ContextMenuPopupDebugState hoverState{};
            return DebugGetContextMenuPopupState(popupHwnd, hoverState) && hoverState.hoveredIndex == std::optional<size_t>{1u} &&
                   DebugGetContextMenuPopupItemPaint(popupHwnd, 1u, paintState) && paintState.usesHighlightFill;
        };
        do
        {
            hoverObserved = observeHover();
            if (hoverObserved)
            {
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(25));
        } while (std::chrono::steady_clock::now() < hoverDeadline);
        if (! hoverObserved)
        {
            driverFailure = "owner-window message flood does not starve popup hover highlight";
            return;
        }

        if (! PostCapturedMouseMessageForMenuSuite(popupHwnd, WM_LBUTTONDOWN, MK_LBUTTON, refineCenter) ||
            ! PostCapturedMouseMessageForMenuSuite(popupHwnd, WM_LBUTTONUP, 0, refineCenter))
        {
            driverFailure = "owner-message-flood posts popup item click";
            return;
        }

        if (! WaitForWindowDestroyed(popupHwnd, std::chrono::milliseconds(800)))
        {
            driverFailure = "owner-window message flood does not delay popup invocation";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), menuAnchor, items, ownerWindow.Host().GetTheme());
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(result == std::optional<int>{4202}, "owner-window message flood cannot delay popup pointer hover or invocation");
}

std::shared_ptr<RedSalamander::DxUi::MenuFlyoutItem::BitmapIcon> CreateSyntheticMenuBitmapIcon(UINT sizePx)
{
    if (sizePx == 0u)
    {
        return {};
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = static_cast<LONG>(sizePx);
    bmi.bmiHeader.biHeight      = -static_cast<LONG>(sizePx);
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hdc_window screenDc{GetDC(nullptr)};
    wil::unique_hbitmap bitmap{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        return {};
    }

    const auto pixelCount = static_cast<size_t>(sizePx) * static_cast<size_t>(sizePx);
    auto* const pixels    = static_cast<uint32_t*>(bits);
    std::fill_n(pixels, pixelCount, 0xFF1F7AE0u);

    const UINT inset = sizePx > 6u ? 2u : 1u;
    for (UINT y = inset; y < (sizePx - inset); ++y)
    {
        for (UINT x = inset; x < (sizePx - inset); ++x)
        {
            pixels[static_cast<size_t>(y) * sizePx + x] = 0xFFFFFFFFu;
        }
    }

    return std::make_shared<RedSalamander::DxUi::MenuFlyoutItem::BitmapIcon>(std::move(bitmap), sizePx, sizePx);
}

RedSalamander::DxUi::WindowHostBitmapCapture CaptureMenuPopupBitmapForTheme(const RedSalamander::DxUi::ThemePalette& theme,
                                                                            const std::vector<RedSalamander::DxUi::MenuFlyoutItem>& items)
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    WindowHostBitmapCapture capture{};
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for material capture validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
        {
            driverFailure = "menu popup bitmap capture succeeds for material validation";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the material-capture menu with Escape returns no invoked command");
    return capture;
}

[[nodiscard]] RECT GetPrimaryMonitorWorkArea() noexcept
{
    MONITORINFO monitorInfo{};
    monitorInfo.cbSize     = sizeof(monitorInfo);
    const HMONITOR monitor = MonitorFromPoint(POINT{0, 0}, MONITOR_DEFAULTTOPRIMARY);
    if (monitor && GetMonitorInfoW(monitor, &monitorInfo) != FALSE)
    {
        return monitorInfo.rcWork;
    }

    return RECT{0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN)};
}

[[nodiscard]] RECT GetVirtualScreenBounds() noexcept
{
    const int left   = GetSystemMetrics(SM_XVIRTUALSCREEN);
    const int top    = GetSystemMetrics(SM_YVIRTUALSCREEN);
    const int width  = (std::max)(1, GetSystemMetrics(SM_CXVIRTUALSCREEN));
    const int height = (std::max)(1, GetSystemMetrics(SM_CYVIRTUALSCREEN));
    return RECT{left, top, left + width, top + height};
}

[[nodiscard]] RECT GetNearestMonitorWorkArea(POINT screenPoint) noexcept
{
    MONITORINFO monitorInfo{};
    monitorInfo.cbSize     = sizeof(monitorInfo);
    const HMONITOR monitor = MonitorFromPoint(screenPoint, MONITOR_DEFAULTTONEAREST);
    if (monitor && GetMonitorInfoW(monitor, &monitorInfo) != FALSE)
    {
        return monitorInfo.rcWork;
    }

    return GetPrimaryMonitorWorkArea();
}

void TestMenuKeyboardNavigationSkipsInfoRows()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Info, .text = L"Used Space:", .acceleratorText = L"561 GB"},
        {.kind = MenuItemKind::Standard, .text = L"Disk Properties...", .commandId = 3001},
        {.kind = MenuItemKind::Info, .text = L"Free Space:", .acceleratorText = L"1.27 TB"},
        {.kind = MenuItemKind::Standard, .text = L"Disk Cleanup...", .commandId = 3002},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for info-row keyboard navigation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        PostMessageW(popupHwnd, WM_KEYDOWN, VK_DOWN, 0);
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 1u;
        }, popupState))
        {
            driverFailure = "first VK_DOWN skips the leading info row and focuses the first actionable item";
            return;
        }

        PostMessageW(popupHwnd, WM_KEYDOWN, VK_DOWN, 0);
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 3u;
        }, popupState))
        {
            driverFailure = "second VK_DOWN skips the middle info row and focuses the next actionable item";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the info-row navigation menu with Escape returns no invoked command");
}

void TestMenuKeyboardRightArrowMatchesWindowsMenuLoop()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();
    ownerWindow.PumpMessages();
    DrainPendingMouseMessagesForMenuSuite();

    const std::vector<std::vector<MenuFlyoutItem>> rootMenus = {
        {
            {.text = L"A1", .commandId = 3101},
            {.text = L"A2", .commandId = 3102},
        },
        {
            {.text      = L"B1",
             .commandId = 3201,
             .children =
                 {
                     MenuFlyoutItem{.text = L"B11", .commandId = 3211},
                     MenuFlyoutItem{.text = L"B12", .commandId = 3212},
                 }},
            {.text = L"B2", .commandId = 3202},
        },
        {
            {.text = L"C1", .commandId = 3301},
            {.text = L"C2", .commandId = 3302},
            {.text = L"C3", .commandId = 3303},
        },
    };

    size_t activeRootIndex = 0u;
    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromDirection = [&](bool forward) -> std::optional<ContextMenuRootSwitchRequest>
    {
        activeRootIndex = forward ? ((activeRootIndex + 1u) % rootMenus.size()) : ((activeRootIndex + rootMenus.size() - 1u) % rootMenus.size());

        ContextMenuRootSwitchRequest request{};
        request.screenPoint = POINT{180, 180};
        request.items       = rootMenus[activeRootIndex];
        return request;
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND aPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"A1");
        if (! aPopupHwnd)
        {
            driverFailure = "menu popup window appears for right-arrow root switching validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        ContextMenuPopupDebugState popupState{};
        PostMessageW(aPopupHwnd, WM_KEYDOWN, VK_HOME, 0);
        if (! WaitForContextMenuPopupState(aPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Home selects the first A root item before directional switching";
            return;
        }

        PostMessageW(aPopupHwnd, WM_KEYDOWN, VK_DOWN, 0);
        if (! WaitForContextMenuPopupState(aPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 1u;
        }, popupState))
        {
            driverFailure = "Down selects A2 before directional switching";
            return;
        }

        PostMessageW(aPopupHwnd, WM_KEYDOWN, VK_RIGHT, 0);
        const HWND bPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! bPopupHwnd)
        {
            driverFailure = "Right arrow on A2 opens the B root popup";
            return;
        }

        if (! WaitForContextMenuPopupState(bPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Right arrow on A2 focuses B1 in the next root popup";
            return;
        }

        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_RIGHT, 0);
        const HWND bSubmenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B11");
        if (! bSubmenuHwnd)
        {
            driverFailure = "Right arrow on B1 opens the B submenu";
            return;
        }

        if (! WaitForContextMenuPopupState(bSubmenuHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Right arrow on B1 focuses B11 in the submenu";
            return;
        }

        PostMessageW(bSubmenuHwnd, WM_KEYDOWN, VK_RIGHT, 0);
        const HWND cPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"C1");
        if (! cPopupHwnd)
        {
            driverFailure = "Right arrow on B11 switches from the leaf submenu to the C root popup";
            return;
        }

        if (! WaitForContextMenuPopupState(cPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Right arrow on B11 focuses C1 after switching to the next root popup";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the right-arrow menu loop validation popup with Escape returns no invoked command");
}

void TestStationaryMouseDoesNotOverrideKeyboardRootSwitch()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<std::vector<MenuFlyoutItem>> rootMenus = {
        {
            {.text = L"A1", .commandId = 3651},
            {.text = L"A2", .commandId = 3652},
        },
        {
            {.text = L"B1", .commandId = 3661},
            {.text = L"B2", .commandId = 3662},
        },
        {
            {.text = L"C1", .commandId = 3671},
            {.text = L"C2", .commandId = 3672},
        },
    };
    const std::array<POINT, 3> rootPopupPoints = {POINT{170, 180}, POINT{250, 180}, POINT{330, 180}};

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromDirection = [&](bool forward) -> std::optional<ContextMenuRootSwitchRequest>
    {
        activeRootIndex = forward ? ((activeRootIndex + 1u) % rootMenus.size()) : ((activeRootIndex + rootMenus.size() - 1u) % rootMenus.size());
        return buildRequest(activeRootIndex);
    };
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE)
        {
            return std::nullopt;
        }
        if (clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 40 && clientPoint.x < 120)
        {
            hitIndex = 0u;
        }
        else if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 2u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex)
        {
            return std::nullopt;
        }

        activeRootIndex = hitIndex;
        return buildRequest(activeRootIndex);
    };

    DrainPendingMouseMessagesForMenuSuite();

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND aPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"A1");
        if (! aPopupHwnd)
        {
            driverFailure = "initial A popup opens before stationary-mouse keyboard switching validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        if (PostMessageW(ownerWindow.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(160, 16)) == FALSE)
        {
            driverFailure = "posting the initial B hover before stationary-mouse keyboard switching validation succeeds";
            return;
        }

        const HWND bPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! bPopupHwnd)
        {
            driverFailure = "initial owner mouse move switches from A to B before stationary-mouse keyboard switching validation";
            return;
        }

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(
                bPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept { return state.rootPointerSwitchCount == 1u; }, popupState))
        {
            driverFailure = "initial B hover is processed before keyboard root switching validation";
            return;
        }

        if (! SendMenuKeyForMenuSuite(bPopupHwnd, VK_HOME))
        {
            driverFailure = "Home key is delivered before stationary-mouse keyboard switching validation";
            return;
        }
        if (! WaitForContextMenuPopupState(bPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Home selects B1 before stationary-mouse keyboard switching validation";
            return;
        }

        if (! SendMenuKeyForMenuSuite(bPopupHwnd, VK_RIGHT))
        {
            driverFailure = "Right arrow key is delivered before stationary-mouse validation";
            return;
        }
        const HWND cPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"C1");
        if (! cPopupHwnd)
        {
            driverFailure = "Right arrow switches from B to C before stationary-mouse validation";
            return;
        }

        if (! WaitForContextMenuPopupState(cPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "switching from B to C focuses C1 before stationary-mouse validation";
            return;
        }

        if (PostMessageW(ownerWindow.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(160, 16)) == FALSE)
        {
            driverFailure = "posting the stationary B hover after switching to C succeeds";
            return;
        }
        if (WaitForWindowDestroyed(cPopupHwnd, std::chrono::milliseconds(120)))
        {
            driverFailure = "a stationary mouse over B does not pull keyboard navigation back from C";
            return;
        }

        if (! SendMenuKeyForMenuSuite(cPopupHwnd, VK_LEFT))
        {
            driverFailure = "Left arrow key is delivered before stationary-mouse previous-root validation";
            return;
        }
        const HWND bPopupAfterLeftHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! bPopupAfterLeftHwnd)
        {
            driverFailure = "Left arrow returns from C to B before stationary-mouse previous-root validation";
            return;
        }

        if (! SendMenuKeyForMenuSuite(bPopupAfterLeftHwnd, VK_LEFT))
        {
            driverFailure = "second Left arrow key is delivered before stationary-mouse previous-root validation";
            return;
        }
        const HWND aPopupAfterSecondLeftHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"A1");
        if (! aPopupAfterSecondLeftHwnd)
        {
            driverFailure = "a second Left arrow switches from B to A before stationary-mouse previous-root validation";
            return;
        }

        if (! WaitForContextMenuPopupState(aPopupAfterSecondLeftHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "switching from B to A focuses A1 before stationary-mouse previous-root validation";
            return;
        }

        if (PostMessageW(ownerWindow.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(160, 16)) == FALSE)
        {
            driverFailure = "posting the stationary B hover after switching to A succeeds";
            return;
        }
        if (WaitForWindowDestroyed(aPopupAfterSecondLeftHwnd, std::chrono::milliseconds(120)))
        {
            driverFailure = "a stationary mouse over B does not pull keyboard navigation back from A";
        }
    });

    const ThemePalette theme = MakeDefaultThemePalette(true);
    const std::optional<int> result =
        ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the stationary-mouse keyboard switching validation popup with Escape returns no invoked command");
}

void TestMenuPointerOverSiblingRootSwitchesOpenMenu()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<std::vector<MenuFlyoutItem>> rootMenus = {
        {
            {.text = L"View one", .commandId = 36701},
            {.text = L"View two", .commandId = 36702},
        },
        {
            {.text = L"Plugins one", .commandId = 36711},
            {.text = L"Plugins two", .commandId = 36712},
        },
    };
    const std::array<POINT, 2> rootPopupPoints = {POINT{270, 180}, POINT{190, 180}};

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE)
        {
            return std::nullopt;
        }
        if (clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 0u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex)
        {
            return std::nullopt;
        }

        activeRootIndex = hitIndex;
        return buildRequest(activeRootIndex);
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before pointer root-switch validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        POINT pluginsMenuClientPoint{180, 20};
        POINT pluginsMenuScreenPoint = pluginsMenuClientPoint;
        if (ClientToScreen(ownerWindow.Hwnd(), &pluginsMenuScreenPoint) == FALSE)
        {
            driverFailure = "Plugins root test point converts to screen coordinates";
            return;
        }

        RECT viewPopupRect{};
        if (GetWindowRect(viewPopupHwnd, &viewPopupRect) == FALSE)
        {
            driverFailure = "View root popup exposes a screen rect before pointer root-switch validation";
            return;
        }

        const LPARAM capturedMouseMovePoint = MAKELPARAM(pluginsMenuScreenPoint.x - viewPopupRect.left, pluginsMenuScreenPoint.y - viewPopupRect.top);
        PostMessageW(viewPopupHwnd, WM_MOUSEMOVE, 0, capturedMouseMovePoint);

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one");
        if (! pluginsPopupHwnd)
        {
            driverFailure = "moving the pointer from an open View root to the Plugins root opens the Plugins popup";
            return;
        }

        ContextMenuPopupDebugState pluginsPopupState{};
        if (! DebugGetContextMenuPopupState(pluginsPopupHwnd, pluginsPopupState) || pluginsPopupState.rootSwitchImmediateRenderCount == 0u ||
            pluginsPopupState.renderCount == 0u)
        {
            driverFailure = "root-switched Plugins popup is painted before the menu loop accepts another pointer move";
            return;
        }

        if (IsWindow(viewPopupHwnd) != FALSE)
        {
            driverFailure = "switching to the Plugins root closes the previous View popup";
        }
    });

    const ThemePalette theme = MakeDefaultThemePalette(true);
    const std::optional<int> result =
        ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the pointer root-switch validation popup with Escape returns no invoked command");
}

void TestMenuPopupMouseMoveUsesDeliveredPointForRootSwitch()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36761}, {.text = L"View two", .commandId = 36762}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36771}, {.text = L"Plugins two", .commandId = 36772}},
    };
    const std::array<POINT, 2> rootPopupPoints = {POINT{270, 180}, POINT{190, 180}};

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE || clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 0u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex)
        {
            return std::nullopt;
        }

        activeRootIndex = hitIndex;
        return buildRequest(activeRootIndex);
    };

    DrainPendingMouseMessagesForMenuSuite();

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before delivered root-switch validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const POINT pluginsMenuScreenPoint =
            ClientScreenPointForTest(ownerWindow.Hwnd(), 180, 20, "Plugins delivered root-switch point converts to screen coordinates");
        if (! PostCapturedMouseMessageForMenuSuite(viewPopupHwnd, WM_MOUSEMOVE, 0, pluginsMenuScreenPoint))
        {
            driverFailure = "View popup receives delivered root-switch mouse move";
            return;
        }

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one");
        if (! pluginsPopupHwnd)
        {
            driverFailure = "popup mouse-move root switching must use the delivered lParam point even when the live cursor is elsewhere";
            return;
        }

        if (IsWindow(viewPopupHwnd) != FALSE)
        {
            driverFailure = "delivered root-switch closes the previous View popup";
        }
    });

    const ThemePalette theme = MakeDefaultThemePalette(true);
    const std::optional<int> result =
        ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(activeRootIndex == 1u, "delivered root-switch validation records the accepted root index");
    Require(! result.has_value(), "closing the delivered root-switch validation popup with Escape returns no invoked command");
}

void TestMenuOwnerMouseMoveRoutesRootSwitchImmediately()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36781}, {.text = L"View two", .commandId = 36782}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36791}, {.text = L"Plugins two", .commandId = 36792}},
    };
    const std::array<POINT, 2> rootPopupPoints = {
        ClientScreenPointForTest(ownerWindow.Hwnd(), 270, 60, "owner-move View root popup point converts to screen coordinates"),
        ClientScreenPointForTest(ownerWindow.Hwnd(), 190, 60, "owner-move Plugins root popup point converts to screen coordinates"),
    };

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE || clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 0u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex)
        {
            return std::nullopt;
        }

        activeRootIndex = hitIndex;
        return buildRequest(activeRootIndex);
    };

    DrainPendingMouseMessagesForMenuSuite();

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before owner mouse-move root switching validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        if (PostMessageW(ownerWindow.Hwnd(), WM_MOUSEMOVE, 0, MousePointLParamForMenuSuite(180, 20)) == 0)
        {
            driverFailure = "owner window receives a mouse-move over the Plugins root while the menu loop is active";
            return;
        }

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one");
        if (! pluginsPopupHwnd)
        {
            driverFailure = "owner-window WM_MOUSEMOVE over a sibling root switches the open menu immediately";
            return;
        }

        ContextMenuPopupDebugState pluginsPopupState{};
        if (! DebugGetContextMenuPopupState(pluginsPopupHwnd, pluginsPopupState) || pluginsPopupState.rootSwitchImmediateRenderCount == 0u ||
            pluginsPopupState.renderCount == 0u)
        {
            driverFailure = "owner-window mouse-move root switch paints the replacement popup immediately";
            return;
        }

        if (IsWindow(viewPopupHwnd) != FALSE)
        {
            driverFailure = "owner-window mouse-move root switch closes the previous View popup";
        }
    });

    const ThemePalette theme = MakeDefaultThemePalette(true);
    const std::optional<int> result =
        ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(activeRootIndex == 1u, "owner-window mouse-move root switch records the accepted root index");
    Require(! result.has_value(), "closing the owner-window mouse-move root switch popup with Escape returns no invoked command");
}

void TestMenuBarHoverMessageSwitchesRootWhenCursorOutsidePopup()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36741}, {.text = L"View two", .commandId = 36742}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36751}, {.text = L"Plugins two", .commandId = 36752}},
    };
    const std::array<POINT, 2> rootPopupPoints = {
        ClientScreenPointForTest(ownerWindow.Hwnd(), 270, 60, "View synthetic hover root popup point converts to screen coordinates"),
        ClientScreenPointForTest(ownerWindow.Hwnd(), 190, 60, "Plugins synthetic hover root popup point converts to screen coordinates"),
    };

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    std::atomic<int> pendingMenuBarHoverRootSwitch{-1};
    std::atomic<std::uintptr_t> pendingMenuBarHoverSequence{0u};
    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromMenuBarHover = [&](size_t hoverIndex, std::uintptr_t sequence) -> std::optional<ContextMenuRootSwitchRequest>
    {
        const int expectedHoverIndex = pendingMenuBarHoverRootSwitch.load(std::memory_order_acquire);
        const std::uintptr_t expectedSequence = pendingMenuBarHoverSequence.load(std::memory_order_acquire);
        if (expectedHoverIndex < 0 || static_cast<size_t>(expectedHoverIndex) != hoverIndex || expectedSequence != sequence)
        {
            return std::nullopt;
        }

        pendingMenuBarHoverRootSwitch.store(-1, std::memory_order_release);
        activeRootIndex = hoverIndex;
        return buildRequest(activeRootIndex);
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before direct synthetic menu-bar hover validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        pendingMenuBarHoverRootSwitch.store(1, std::memory_order_release);
        pendingMenuBarHoverSequence.store(2u, std::memory_order_release);
        if (PostMessageW(viewPopupHwnd, WndMsg::kDxUiContextMenuRootHoverChanged, 0u, 1u) == 0 ||
            PostMessageW(viewPopupHwnd, WndMsg::kDxUiContextMenuRootHoverChanged, 1u, 2u) == 0)
        {
            driverFailure = "View popup receives the direct synthetic menu-bar hover switch messages";
            return;
        }

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one");
        if (! pluginsPopupHwnd)
        {
            driverFailure = "direct synthetic menu-bar hover message switches from View to Plugins while cursor is outside the popup";
            return;
        }

        ContextMenuPopupDebugState pluginsPopupState{};
        if (! DebugGetContextMenuPopupState(pluginsPopupHwnd, pluginsPopupState) || pluginsPopupState.rootSwitchImmediateRenderCount == 0u ||
            pluginsPopupState.renderCount == 0u)
        {
            driverFailure = "synthetic menu-bar hover root switch paints the replacement Plugins popup immediately";
            return;
        }

        if (IsWindow(viewPopupHwnd) != FALSE)
        {
            driverFailure = "synthetic menu-bar hover root switch closes the previous View popup";
        }
    });

    const ThemePalette theme = MakeDefaultThemePalette(true);
    const std::optional<int> result =
        ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(activeRootIndex == 1u, "direct synthetic menu-bar hover validation records the accepted root index");
    Require(! result.has_value(), "closing the direct synthetic menu-bar hover popup with Escape returns no invoked command");
}

void TestMenuBarHoverMessageSwitchesRootWhilePopupOverlapsMenuBar()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36721}, {.text = L"View two", .commandId = 36722}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36731}, {.text = L"Plugins two", .commandId = 36732}},
    };
    const std::array<POINT, 2> rootPopupPoints = {
        ClientScreenPointForTest(ownerWindow.Hwnd(), 90, 0, "View overlapping root popup point converts to screen coordinates"),
        ClientScreenPointForTest(ownerWindow.Hwnd(), 190, 0, "Plugins overlapping root popup point converts to screen coordinates"),
    };
    constexpr LONG kMainMenuStripHeightPx = 96;

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    std::atomic_bool pointerRootSwitchArmed{false};
    std::atomic<int> pendingMenuBarHoverRootSwitch{-1};
    std::atomic<std::uintptr_t> pendingMenuBarHoverSequence{0u};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        if (! pointerRootSwitchArmed.load(std::memory_order_acquire))
        {
            return std::nullopt;
        }

        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE || clientPoint.y < 0 || clientPoint.y >= kMainMenuStripHeightPx)
        {
            return std::nullopt;
        }
        if (activeRootIndex == 1u)
        {
            return std::nullopt;
        }

        activeRootIndex = 1u;
        return buildRequest(activeRootIndex);
    };
    sessionCallbacks.switchRootFromMenuBarHover = [&](size_t hoverIndex, std::uintptr_t sequence) -> std::optional<ContextMenuRootSwitchRequest>
    {
        const int expectedHoverIndex = pendingMenuBarHoverRootSwitch.load(std::memory_order_acquire);
        const std::uintptr_t expectedSequence = pendingMenuBarHoverSequence.load(std::memory_order_acquire);
        if (expectedHoverIndex < 0 || static_cast<size_t>(expectedHoverIndex) != hoverIndex || expectedSequence != sequence)
        {
            return std::nullopt;
        }

        pendingMenuBarHoverRootSwitch.store(-1, std::memory_order_release);
        activeRootIndex = hoverIndex;
        return buildRequest(activeRootIndex);
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before overlapping-popup validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(viewPopupHwnd, [](const ContextMenuPopupDebugState&) noexcept { return true; }, popupState))
        {
            driverFailure = "overlapping View popup exposes debug state";
            return;
        }

        D2D1_RECT_F firstItemRectDip{};
        if (! WaitForContextMenuPopupItemRect(viewPopupHwnd, 0u, firstItemRectDip))
        {
            driverFailure = "overlapping View popup exposes its first item rect";
            return;
        }

        const float itemMidXDip = (firstItemRectDip.left + firstItemRectDip.right) * 0.5f;
        const float itemMidYDip = (firstItemRectDip.top + firstItemRectDip.bottom) * 0.5f;
        const POINT popupItemScreenPoint{
            popupState.surfaceRectPx.left + static_cast<LONG>(DipToPixelForPopup(itemMidXDip, popupState.dpi)),
            popupState.surfaceRectPx.top + static_cast<LONG>(DipToPixelForPopup(itemMidYDip, popupState.dpi)),
        };

        POINT ownerClientPoint = popupItemScreenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &ownerClientPoint) == FALSE)
        {
            driverFailure = "overlapping popup sample point converts to owner client coordinates";
            return;
        }
        if (ownerClientPoint.y < 0 || ownerClientPoint.y >= kMainMenuStripHeightPx)
        {
            driverFailure = "overlapping popup sample point sits over the owner main-menu strip";
            return;
        }

        RECT popupWindowRect{};
        if (GetWindowRect(viewPopupHwnd, &popupWindowRect) == FALSE)
        {
            driverFailure = "overlapping View popup exposes a screen rect";
            return;
        }

        pointerRootSwitchArmed.store(true, std::memory_order_release);
        PostMessageW(viewPopupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(popupItemScreenPoint.x - popupWindowRect.left, popupItemScreenPoint.y - popupWindowRect.top));

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one", std::chrono::milliseconds(180));
        if (pluginsPopupHwnd)
        {
            driverFailure = "moving inside a frontmost popup must not switch the root menu underneath it";
            return;
        }

        pendingMenuBarHoverRootSwitch.store(1, std::memory_order_release);
        pendingMenuBarHoverSequence.store(1u, std::memory_order_release);
        if (PostMessageW(viewPopupHwnd, WndMsg::kDxUiContextMenuRootHoverChanged, 1u, 1u) == 0)
        {
            driverFailure = "overlapping popup can receive the synthetic menu-bar hover switch message";
            return;
        }
        const HWND pluginsPopupFromHoverHwnd =
            WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one", std::chrono::milliseconds(900));
        if (! pluginsPopupFromHoverHwnd)
        {
            driverFailure = "an explicit menu-bar hover message must switch root while a frontmost popup overlaps the menu strip";
            return;
        }

        if (IsWindow(viewPopupHwnd) != FALSE && IsWindowVisible(viewPopupHwnd) != FALSE)
        {
            driverFailure = "View popup should close after explicit menu-bar hover switches to the Plugins root";
        }
    });

    const ThemePalette theme = MakeDefaultThemePalette(true);
    const std::optional<int> result =
        ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the overlapping popup validation with Escape returns no invoked command");
}

void TestMenuRootSwitchUsesDeliveredOwnerMouseMoveAfterPopupSwitch()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36801}, {.text = L"View two", .commandId = 36802}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36811}, {.text = L"Plugins two", .commandId = 36812}},
    };
    const std::array<POINT, 2> rootPopupPoints = {POINT{270, 180}, POINT{190, 180}};

    std::atomic<size_t> activeRootIndex{0u};
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE || clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 0u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex.load(std::memory_order_acquire))
        {
            return std::nullopt;
        }

        activeRootIndex.store(hitIndex, std::memory_order_release);
        return buildRequest(hitIndex);
    };

    DrainPendingMouseMessagesForMenuSuite();

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before delivered owner mouse-move validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        POINT pluginsMenuClientPoint{180, 20};
        POINT pluginsMenuScreenPoint = pluginsMenuClientPoint;
        if (ClientToScreen(ownerWindow.Hwnd(), &pluginsMenuScreenPoint) == FALSE)
        {
            driverFailure = "Plugins root test point converts to screen coordinates for delivered owner mouse-move validation";
            return;
        }

        POINT viewMenuClientPoint{280, 20};
        POINT viewMenuScreenPoint = viewMenuClientPoint;
        if (ClientToScreen(ownerWindow.Hwnd(), &viewMenuScreenPoint) == FALSE)
        {
            driverFailure = "View root delivered owner point converts to screen coordinates";
            return;
        }

        if (! PostCapturedMouseMessageForMenuSuite(viewPopupHwnd, WM_MOUSEMOVE, 0, pluginsMenuScreenPoint))
        {
            driverFailure = "View popup receives the delivered Plugins root-switch move";
            return;
        }

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one");
        if (! pluginsPopupHwnd)
        {
            driverFailure = "delivered popup move opens the Plugins popup before delivered owner mouse-move validation";
            return;
        }
        if (activeRootIndex.load(std::memory_order_acquire) != 1u)
        {
            driverFailure = "delivered popup move records the Plugins root as active before owner mouse-move validation";
            return;
        }

        POINT deliveredViewClientPoint = viewMenuScreenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &deliveredViewClientPoint) == FALSE)
        {
            driverFailure = "View delivered owner point converts to owner client coordinates";
            return;
        }
        if (PostMessageW(ownerWindow.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(deliveredViewClientPoint.x, deliveredViewClientPoint.y)) == FALSE)
        {
            driverFailure = "owner receives the delivered View root-switch move";
            return;
        }

        const HWND deliveredViewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one", std::chrono::milliseconds(900));
        if (! deliveredViewPopupHwnd)
        {
            driverFailure = "a fresh delivered owner View mouse-move should switch root after the Plugins popup opened";
            return;
        }

        if (activeRootIndex.load(std::memory_order_acquire) != 0u)
        {
            driverFailure = "fresh delivered owner View mouse-move should record the View root as active";
            return;
        }

        if (IsWindow(pluginsPopupHwnd) != FALSE && IsWindowVisible(pluginsPopupHwnd) != FALSE)
        {
            driverFailure = "Plugins popup should close after the fresh delivered owner View mouse-move message";
        }
    });

    const ThemePalette theme      = MakeDefaultThemePalette(true);
    const size_t initialRootIndex = activeRootIndex.load(std::memory_order_acquire);
    const std::optional<int> result =
        ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[initialRootIndex], rootMenus[initialRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the delivered owner mouse-move validation popup with Escape returns no invoked command");
}

void TestMenuRootSwitchDoesNotPollCursorWhileIdle()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36901}, {.text = L"View two", .commandId = 36902}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36911}, {.text = L"Plugins two", .commandId = 36912}},
    };
    const std::array<POINT, 2> rootPopupPoints = {POINT{270, 180}, POINT{190, 180}};

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE || clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 0u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex)
        {
            return std::nullopt;
        }

        activeRootIndex = hitIndex;
        return buildRequest(activeRootIndex);
    };

    DrainPendingMouseMessagesForMenuSuite();

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before idle cursor polling validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one", std::chrono::milliseconds(180));
        if (pluginsPopupHwnd)
        {
            driverFailure = "waiting without a mouse-move message must not switch root menus by idle polling";
            return;
        }

        if (IsWindow(viewPopupHwnd) == FALSE)
        {
            driverFailure = "View popup remains alive when no mouse-move message is delivered";
        }
    });

    const ThemePalette theme = MakeDefaultThemePalette(true);
    const std::optional<int> result =
        ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the idle cursor polling validation popup with Escape returns no invoked command");
}

void TestMenuHoveringSiblingClosesOpenSubmenuAfterDelay()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();
    DrainPendingMouseMessagesForMenuSuite();
    const std::vector<MenuFlyoutItem> items = {
        {.text      = L"B1",
         .commandId = 3681,
         .children =
             {
                 MenuFlyoutItem{.text = L"B11", .commandId = 36811},
                 MenuFlyoutItem{.text = L"B12", .commandId = 36812},
             }},
        {.text = L"B2", .commandId = 3682},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND rootPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! rootPopupHwnd)
        {
            driverFailure = "menu popup window appears for delayed submenu-close validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        D2D1_RECT_F firstItemRectDip  = D2D1::RectF();
        D2D1_RECT_F secondItemRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(rootPopupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 0u, firstItemRectDip) || ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 1u, secondItemRectDip))
        {
            driverFailure = "menu popup exposes geometry for delayed submenu-close validation";
            return;
        }

        const float scale = static_cast<float>(popupState.dpi) / 96.0f;
        const LONG b1X    = static_cast<LONG>(std::lround(((firstItemRectDip.left + firstItemRectDip.right) * 0.5f) * scale));
        const LONG b1Y    = static_cast<LONG>(std::lround(((firstItemRectDip.top + firstItemRectDip.bottom) * 0.5f) * scale));
        const LONG b2X    = static_cast<LONG>(std::lround(((secondItemRectDip.left + secondItemRectDip.right) * 0.5f) * scale));
        const LONG b2Y    = static_cast<LONG>(std::lround(((secondItemRectDip.top + secondItemRectDip.bottom) * 0.5f) * scale));

        if (! SendSettledClientMouseMoveForMenuSuite(rootPopupHwnd, b1X, b1Y) ||
            ! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "hovering B1 updates the root-popup hover target before delayed submenu close";
            return;
        }

        HWND submenuHwnd = nullptr;
        if (! WaitForSubmenuOpenAfterHoverForMenuSuite(ownerWindow.Hwnd(), rootPopupHwnd, 0u, L"B11", submenuHwnd, popupState))
        {
            driverFailure = "hovering B1 long enough opens its submenu before delayed close validation";
            return;
        }

        if (! SendSettledClientMouseMoveForMenuSuite(rootPopupHwnd, b2X, b2Y) ||
            ! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
        }, popupState))
        {
            driverFailure = "hovering B2 updates the root-popup hover target before delayed submenu close";
            return;
        }

        if (! WaitForSubmenuClosedAfterHoverForMenuSuite(rootPopupHwnd, submenuHwnd, popupState))
        {
            driverFailure = "hovering B2 long enough closes the already-open submenu from B1";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the delayed submenu-close validation popup with Escape returns no invoked command");
}

void TestMenuHoveringSiblingWithChildrenReplacesOpenSubmenuAfterDelay()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();
    DrainPendingMouseMessagesForMenuSuite();
    const std::vector<MenuFlyoutItem> items = {
        {.text      = L"B1",
         .commandId = 3691,
         .children =
             {
                 MenuFlyoutItem{.text = L"B11", .commandId = 36911},
                 MenuFlyoutItem{.text = L"B12", .commandId = 36912},
             }},
        {.text      = L"B2",
         .commandId = 3692,
         .children =
             {
                 MenuFlyoutItem{.text = L"B21", .commandId = 36921},
                 MenuFlyoutItem{.text = L"B22", .commandId = 36922},
             }},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND rootPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! rootPopupHwnd)
        {
            driverFailure = "menu popup window appears for delayed submenu-replacement validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        D2D1_RECT_F firstItemRectDip  = D2D1::RectF();
        D2D1_RECT_F secondItemRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(rootPopupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 0u, firstItemRectDip) || ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 1u, secondItemRectDip))
        {
            driverFailure = "menu popup exposes geometry for delayed submenu-replacement validation";
            return;
        }

        const float scale = static_cast<float>(popupState.dpi) / 96.0f;
        const LONG b1X    = static_cast<LONG>(std::lround(((firstItemRectDip.left + firstItemRectDip.right) * 0.5f) * scale));
        const LONG b1Y    = static_cast<LONG>(std::lround(((firstItemRectDip.top + firstItemRectDip.bottom) * 0.5f) * scale));
        const LONG b2X    = static_cast<LONG>(std::lround(((secondItemRectDip.left + secondItemRectDip.right) * 0.5f) * scale));
        const LONG b2Y    = static_cast<LONG>(std::lround(((secondItemRectDip.top + secondItemRectDip.bottom) * 0.5f) * scale));

        if (! SendSettledClientMouseMoveForMenuSuite(rootPopupHwnd, b1X, b1Y) ||
            ! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "hovering B1 updates the root-popup hover target before delayed submenu replacement";
            return;
        }

        HWND firstSubmenuHwnd = nullptr;
        if (! WaitForSubmenuOpenAfterHoverForMenuSuite(ownerWindow.Hwnd(), rootPopupHwnd, 0u, L"B11", firstSubmenuHwnd, popupState))
        {
            driverFailure = "hovering B1 long enough opens its submenu before delayed replacement validation";
            return;
        }

        if (! SendSettledClientMouseMoveForMenuSuite(rootPopupHwnd, b2X, b2Y) ||
            ! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
        }, popupState))
        {
            driverFailure = "hovering B2 updates the root-popup hover target before delayed submenu replacement";
            return;
        }

        HWND replacementSubmenuHwnd = nullptr;
        if (! WaitForSubmenuOpenAfterHoverForMenuSuite(ownerWindow.Hwnd(), rootPopupHwnd, 1u, L"B21", replacementSubmenuHwnd, popupState))
        {
            driverFailure = "hovering B2 long enough opens B2's replacement submenu";
            return;
        }

        if (! WaitForWindowDestroyed(firstSubmenuHwnd, std::chrono::milliseconds(1200)))
        {
            driverFailure = "opening B2's submenu closes the older submenu from B1";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the delayed submenu-replacement validation popup with Escape returns no invoked command");
}

void TestMenuPointerInsideSubmenuAndParentItemCancelPendingCloseDelay()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();
    DrainPendingMouseMessagesForMenuSuite();
    const std::vector<MenuFlyoutItem> items = {
        {.text      = L"B1",
         .commandId = 3701,
         .children =
             {
                 MenuFlyoutItem{.text = L"B11", .commandId = 37011},
                 MenuFlyoutItem{.text = L"B12", .commandId = 37012},
             }},
        {.text = L"B2", .commandId = 3702},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND rootPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! rootPopupHwnd)
        {
            driverFailure = "menu popup window appears for submenu hover-retention validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        D2D1_RECT_F firstItemRectDip  = D2D1::RectF();
        D2D1_RECT_F secondItemRectDip = D2D1::RectF();
        ContextMenuPopupDebugState rootState{};
        if (! WaitForContextMenuPopupState(rootPopupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           rootState) ||
            ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 0u, firstItemRectDip) || ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 1u, secondItemRectDip))
        {
            driverFailure = "root popup exposes geometry for submenu hover-retention validation";
            return;
        }

        const float rootScale = static_cast<float>(rootState.dpi) / 96.0f;
        const LONG b1X        = static_cast<LONG>(std::lround(((firstItemRectDip.left + firstItemRectDip.right) * 0.5f) * rootScale));
        const LONG b1Y        = static_cast<LONG>(std::lround(((firstItemRectDip.top + firstItemRectDip.bottom) * 0.5f) * rootScale));
        const LONG b2X        = static_cast<LONG>(std::lround(((secondItemRectDip.left + secondItemRectDip.right) * 0.5f) * rootScale));
        const LONG b2Y        = static_cast<LONG>(std::lround(((secondItemRectDip.top + secondItemRectDip.bottom) * 0.5f) * rootScale));

        if (! SendSettledClientMouseMoveForMenuSuite(rootPopupHwnd, b1X, b1Y) ||
            ! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 0u;
        }, rootState))
        {
            driverFailure = "hovering B1 updates the root-popup hover target before hover-retention validation";
            return;
        }

        HWND submenuHwnd = nullptr;
        if (! WaitForSubmenuOpenAfterHoverForMenuSuite(ownerWindow.Hwnd(), rootPopupHwnd, 0u, L"B11", submenuHwnd, rootState))
        {
            driverFailure = "hovering B1 long enough opens its submenu before hover-retention validation";
            return;
        }

        if (! SendSettledClientMouseMoveForMenuSuite(rootPopupHwnd, b2X, b2Y) ||
            ! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
        }, rootState))
        {
            driverFailure = "hovering B2 starts the delayed close path before entering the submenu";
            return;
        }
        if (! WaitForPendingSubmenuCloseTimerForMenuSuite(rootPopupHwnd, rootState))
        {
            driverFailure = "hovering B2 schedules the parent delayed-close timer before entering the submenu";
            return;
        }

        D2D1_RECT_F submenuItemRectDip = D2D1::RectF();
        ContextMenuPopupDebugState submenuState{};
        if (! WaitForContextMenuPopupState(submenuHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           submenuState) ||
            ! WaitForContextMenuPopupItemRect(submenuHwnd, 0u, submenuItemRectDip))
        {
            driverFailure = "submenu exposes geometry before hover-retention validation";
            return;
        }

        const float submenuScale = static_cast<float>(submenuState.dpi) / 96.0f;
        const LONG submenuX      = static_cast<LONG>(std::lround(((submenuItemRectDip.left + submenuItemRectDip.right) * 0.5f) * submenuScale));
        const LONG submenuY      = static_cast<LONG>(std::lround(((submenuItemRectDip.top + submenuItemRectDip.bottom) * 0.5f) * submenuScale));
        if (! SendSettledClientMouseMoveForMenuSuite(submenuHwnd, submenuX, submenuY) ||
            ! WaitForContextMenuPopupState(submenuHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 0u;
        }, submenuState))
        {
            driverFailure = "moving into an open submenu updates the submenu hover target";
            return;
        }

        if (! WaitForNoSubmenuHoverTimerForMenuSuite(rootPopupHwnd, rootState) || IsWindow(submenuHwnd) == FALSE)
        {
            driverFailure = "moving into an open submenu cancels the parent delayed-close timer";
            return;
        }

        if (! SendSettledClientMouseMoveForMenuSuite(rootPopupHwnd, b1X, b1Y) ||
            ! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 0u;
        }, rootState))
        {
            driverFailure = "moving from an open submenu back to its parent menu item updates root hover";
            return;
        }
        if (! WaitForNoSubmenuHoverTimerForMenuSuite(rootPopupHwnd, rootState) || IsWindow(submenuHwnd) == FALSE)
        {
            driverFailure = "moving from an open submenu back to its parent menu item keeps the submenu open";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the submenu hover-retention validation popup with Escape returns no invoked command");
}

void TestMenuKeyboardTabExitsMenuLoop()
{
    RunMenuDismissalKeyScenario(WM_KEYDOWN,
                                VK_TAB,
                                "menu popup window appears for Tab dismissal validation",
                                "Tab dismisses the active menu popup chain",
                                "dismissing the menu popup with Tab preserves or restores focus to the previously focused control");
}

void TestMenuKeyboardF10ExitsMenuLoop()
{
    RunMenuDismissalKeyScenario(WM_SYSKEYDOWN,
                                VK_F10,
                                "menu popup window appears for F10 dismissal validation",
                                "F10 dismisses the active menu popup chain",
                                "dismissing the menu popup with F10 preserves or restores focus to the previously focused control");
}

void TestMenuKeyboardAltExitsMenuLoop()
{
    RunMenuDismissalKeyScenario(WM_SYSKEYDOWN,
                                VK_MENU,
                                "menu popup window appears for Alt dismissal validation",
                                "Alt dismisses the active menu popup chain",
                                "dismissing the menu popup with Alt preserves or restores focus to the previously focused control");
}

void TestMenuKeyboardLeftArrowMatchesWindowsMenuLoop()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<std::vector<MenuFlyoutItem>> rootMenus = {
        {
            {.text = L"A1", .commandId = 3401},
            {.text = L"A2", .commandId = 3402},
        },
        {
            {.text      = L"B1",
             .commandId = 3501,
             .children =
                 {
                     MenuFlyoutItem{.text = L"B11", .commandId = 3511},
                     MenuFlyoutItem{.text = L"B12", .commandId = 3512},
                 }},
            {.text = L"B2", .commandId = 3502},
        },
        {
            {.text = L"C1", .commandId = 3601},
            {.text = L"C2", .commandId = 3602},
        },
    };

    size_t activeRootIndex = 1u;
    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromDirection = [&](bool forward) -> std::optional<ContextMenuRootSwitchRequest>
    {
        activeRootIndex = forward ? ((activeRootIndex + 1u) % rootMenus.size()) : ((activeRootIndex + rootMenus.size() - 1u) % rootMenus.size());

        ContextMenuRootSwitchRequest request{};
        request.screenPoint = POINT{180, 180};
        request.items       = rootMenus[activeRootIndex];
        return request;
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND bPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! bPopupHwnd)
        {
            driverFailure = "menu popup window appears for left-arrow menu loop validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        ContextMenuPopupDebugState popupState{};
        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_HOME, 0);
        if (! WaitForContextMenuPopupState(bPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Home selects B1 before left-arrow submenu validation";
            return;
        }

        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_RIGHT, 0);
        const HWND bSubmenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B11");
        if (! bSubmenuHwnd)
        {
            driverFailure = "Right arrow on B1 opens the B submenu before left-arrow validation";
            return;
        }

        if (! WaitForContextMenuPopupState(bSubmenuHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Right arrow on B1 focuses B11 before left-arrow validation";
            return;
        }

        PostMessageW(bSubmenuHwnd, WM_KEYDOWN, VK_LEFT, 0);
        if (! WaitForWindowDestroyed(bSubmenuHwnd))
        {
            driverFailure = "Left arrow closes the active submenu before switching roots";
            return;
        }

        if (! WaitForContextMenuPopupState(bPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "closing the submenu restores keyboard focus to B1";
            return;
        }

        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_LEFT, 0);
        const HWND aPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"A1");
        if (! aPopupHwnd)
        {
            driverFailure = "a second Left arrow on B switches to the previous root popup";
            return;
        }

        if (! WaitForContextMenuPopupState(aPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "switching to the previous root popup focuses A1";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the left-arrow menu loop validation popup with Escape returns no invoked command");
}

void TestNativeMenuBarRestoresFocusAfterMenuDismiss()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOW);
    ownerWindow.PumpMessages();

    wil::unique_hwnd focusedChild{CreateWindowExW(
        0, L"BUTTON", L"Pane", WS_CHILD | WS_VISIBLE | WS_TABSTOP, 12, 80, 120, 28, ownerWindow.Hwnd(), nullptr, GetModuleHandleW(nullptr), nullptr)};
    Require(focusedChild != nullptr, "focus-restoration validation creates a focusable child control");

    wil::unique_hmenu menu{CreateMenu()};
    Require(menu != nullptr, "focus-restoration validation creates a top-level native menu");

    wil::unique_hmenu filePopup{CreatePopupMenu()};
    Require(filePopup != nullptr, "focus-restoration validation creates a native popup menu");
    Require(AppendMenuW(filePopup.get(), MF_STRING, 3801u, L"&Open") != FALSE, "focus-restoration validation populates the native popup menu");
    Require(AppendMenuW(menu.get(), MF_POPUP, reinterpret_cast<UINT_PTR>(filePopup.get()), L"&File") != FALSE,
            "focus-restoration validation attaches the popup menu to the native menu bar");
    static_cast<void>(filePopup.release());

    NativeMenuBarHost menuBarHost;
    Require(menuBarHost.Attach(GetModuleHandleW(nullptr), ownerWindow.Hwnd(), menu.get()), "native menu bar host attaches for focus-restoration validation");
    ownerWindow.PumpMessages();

    if (! TryActivateDxUiTestWindow(ownerWindow.Hwnd()))
    {
        SkipDxUiTest("DxUi menu popup requires an interactive desktop for native menu-bar focus restoration");
        return;
    }
    static_cast<void>(SetFocus(focusedChild.get()));
    Require(WaitForFocusedWindow(focusedChild.get()), "focus-restoration validation starts with focus on the child control");

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "native menu bar opens a popup for focus-restoration validation";
            return;
        }

        PostMessageW(popupHwnd, WM_KEYDOWN, VK_TAB, 0);
        if (! WaitForWindowDestroyed(popupHwnd))
        {
            driverFailure = "Tab dismisses the native menu bar popup during focus-restoration validation";
        }
    });

    Require(menuBarHost.FocusFirstItem(), "native menu bar host enters menu mode for focus-restoration validation");
    static_cast<void>(SendMessageW(menuBarHost.GetHwnd(), WM_KEYDOWN, VK_DOWN, 0));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(WaitForFocusedWindow(focusedChild.get()), "dismissing the native menu bar popup restores focus to the previously focused child control");
}

constexpr UINT kDestroyNativeMenuBarHostMessage    = WM_APP + 0x53Eu;
constexpr UINT_PTR kDestroyNativeMenuBarSubclassId = 0x53Eu;

struct DestroyNativeMenuBarHostState final
{
    DestroyNativeMenuBarHostState()                                                   = default;
    DestroyNativeMenuBarHostState(const DestroyNativeMenuBarHostState&)                = delete;
    DestroyNativeMenuBarHostState(DestroyNativeMenuBarHostState&&)                     = delete;
    DestroyNativeMenuBarHostState& operator=(const DestroyNativeMenuBarHostState&)     = delete;
    DestroyNativeMenuBarHostState& operator=(DestroyNativeMenuBarHostState&&)          = delete;

    std::unique_ptr<RedSalamander::DxUi::NativeMenuBarHost>* menuBarHost = nullptr;
    std::atomic_bool destroyed = false;
};

LRESULT CALLBACK DestroyNativeMenuBarHostSubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR refData) noexcept
{
    if (message == kDestroyNativeMenuBarHostMessage)
    {
        auto* const state = reinterpret_cast<DestroyNativeMenuBarHostState*>(refData);
        if (state && state->menuBarHost)
        {
            state->menuBarHost->reset();
            state->destroyed.store(true, std::memory_order_release);
        }
        return 0;
    }

    return DefSubclassProc(hwnd, message, wParam, lParam);
}

void TestNativeMenuBarNestedPopupCanDestroyHostSafely()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOW);
    ownerWindow.PumpMessages();
    if (! TryActivateDxUiTestWindow(ownerWindow.Hwnd()))
    {
        SkipDxUiTest("DxUi native menu destruction proof requires an interactive desktop");
        return;
    }

    wil::unique_hmenu menu{CreateMenu()};
    Require(menu != nullptr, "native menu destruction proof creates a top-level menu");
    wil::unique_hmenu filePopup{CreatePopupMenu()};
    Require(filePopup != nullptr, "native menu destruction proof creates a popup menu");
    Require(AppendMenuW(filePopup.get(), MF_STRING, 3901u, L"&Open") != FALSE, "native menu destruction proof populates the popup");
    Require(AppendMenuW(menu.get(), MF_POPUP, reinterpret_cast<UINT_PTR>(filePopup.get()), L"&File") != FALSE,
            "native menu destruction proof attaches the popup to the menu bar");
    static_cast<void>(filePopup.release());

    auto menuBarHost = std::make_unique<NativeMenuBarHost>();
    Require(menuBarHost->Attach(GetModuleHandleW(nullptr), ownerWindow.Hwnd(), menu.get()), "native menu destruction proof attaches the menu bar host");
    ownerWindow.PumpMessages();

    DestroyNativeMenuBarHostState destroyState;
    destroyState.menuBarHost = &menuBarHost;
    Require(SetWindowSubclass(ownerWindow.Hwnd(), DestroyNativeMenuBarHostSubclassProc, kDestroyNativeMenuBarSubclassId,
                              reinterpret_cast<DWORD_PTR>(&destroyState)) != FALSE,
            "native menu destruction proof subclasses the owner window");
    const auto removeSubclass = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(RemoveWindowSubclass(ownerWindow.Hwnd(), DestroyNativeMenuBarHostSubclassProc, kDestroyNativeMenuBarSubclassId));
    });

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Open");
        if (! popupHwnd)
        {
            driverFailure = "native menu popup opens before destroying its menu-bar host";
            return;
        }

        if (PostMessageW(ownerWindow.Hwnd(), kDestroyNativeMenuBarHostMessage, 0, 0) == FALSE)
        {
            driverFailure = "native menu destruction request posts to the owner";
            return;
        }

        const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(1500);
        while (! destroyState.destroyed.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < deadline)
        {
            Sleep(5);
        }
        if (! destroyState.destroyed.load(std::memory_order_acquire))
        {
            driverFailure = "native menu-bar host is destroyed inside the nested popup loop";
            return;
        }

        if (IsWindow(popupHwnd) != FALSE)
        {
            SendMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
        }
    });

    const bool activated = menuBarHost->ActivateMnemonic(L'F');
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(activated, "native menu activation returns after its host is destroyed inside the popup loop");
    Require(destroyState.destroyed.load(std::memory_order_acquire) && ! menuBarHost,
            "native menu nested-loop proof destroys the menu-bar host without a stale post-loop dereference");
}

void TestMenuInfoRowsDoNotDismissOnClick()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Info, .text = L"Volume Label:", .acceleratorText = L"Home"},
        {.kind = MenuItemKind::Standard, .text = L"Disk Properties...", .commandId = 3001},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for info-row click validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        D2D1_RECT_F infoRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupItemRect(popupHwnd, 0u, infoRectDip))
        {
            driverFailure = "info-row popup exposes geometry and debug state";
            return;
        }

        const float scale       = static_cast<float>(popupState.dpi) / 96.0f;
        const LONG clientX      = static_cast<LONG>(std::lround(((infoRectDip.left + infoRectDip.right) * 0.5f) * scale));
        const LONG clientY      = static_cast<LONG>(std::lround(((infoRectDip.top + infoRectDip.bottom) * 0.5f) * scale));
        const LPARAM clickPoint = MAKELPARAM(clientX, clientY);

        if (! SendSettledClientMouseMoveForMenuSuite(popupHwnd, clientX, clientY))
        {
            driverFailure = "moving the pointer to the info row before clicking succeeds";
            return;
        }
        PostMessageW(popupHwnd, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        PostMessageW(popupHwnd, WM_LBUTTONUP, 0, clickPoint);

        std::this_thread::sleep_for(std::chrono::milliseconds(80));
        if (IsWindow(popupHwnd) == FALSE)
        {
            driverFailure = "clicking an info row does not dismiss the popup";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "clicking an info row does not invoke a command");
}

void TestMenuPopupPositionClampsAcrossDpiMatrix()
{
    using namespace RedSalamander::DxUi;

    const RECT virtualScreen    = GetVirtualScreenBounds();
    constexpr UINT kDpiMatrix[] = {96u, 120u, 144u, 168u, 192u};

    for (const UINT dpi : kDpiMatrix)
    {
        RECT popupRect{};
        const POINT topLeftAnchor{virtualScreen.left - 80, virtualScreen.top - 60};
        const RECT topLeftWorkArea = GetNearestMonitorWorkArea(topLeftAnchor);
        Require(DebugComputeContextMenuPopupPosition(
                    topLeftAnchor, 240.0f, 180.0f, dpi, false, nullptr, nullptr, ContextMenuRootHorizontalAlignment::Start, popupRect),
                std::format("root popup position resolves for {} DPI near the top-left edge", dpi).c_str());
        Require(popupRect.left == topLeftWorkArea.left, std::format("root popup clamps left to the work area at {} DPI", dpi).c_str());
        Require(popupRect.top == topLeftWorkArea.top, std::format("root popup clamps top to the work area at {} DPI", dpi).c_str());
        Require(popupRect.right <= topLeftWorkArea.right && popupRect.bottom <= topLeftWorkArea.bottom,
                std::format("root popup stays inside the work area after top-left clamping at {} DPI", dpi).c_str());

        popupRect = RECT{};
        const POINT bottomRightAnchor{virtualScreen.right + 80, virtualScreen.bottom + 60};
        const RECT bottomRightWorkArea = GetNearestMonitorWorkArea(bottomRightAnchor);
        Require(DebugComputeContextMenuPopupPosition(
                    bottomRightAnchor, 240.0f, 180.0f, dpi, false, nullptr, nullptr, ContextMenuRootHorizontalAlignment::Start, popupRect),
                std::format("root popup position resolves for {} DPI near the bottom-right edge", dpi).c_str());
        Require(popupRect.right == bottomRightWorkArea.right, std::format("root popup clamps right to the work area at {} DPI", dpi).c_str());
        Require(popupRect.bottom == bottomRightWorkArea.bottom, std::format("root popup clamps bottom to the work area at {} DPI", dpi).c_str());
        Require(popupRect.left >= bottomRightWorkArea.left && popupRect.top >= bottomRightWorkArea.top,
                std::format("root popup stays inside the work area after bottom-right clamping at {} DPI", dpi).c_str());
    }
}

void TestMenuPopupPositionSupportsRightAlignedRootAnchors()
{
    using namespace RedSalamander::DxUi;

    const RECT workArea         = GetPrimaryMonitorWorkArea();
    constexpr UINT kDpiMatrix[] = {96u, 120u, 144u, 168u, 192u};

    for (const UINT dpi : kDpiMatrix)
    {
        RECT popupRect{};
        const int anchorRightPx = workArea.right - 24;
        Require(DebugComputeContextMenuPopupPosition(
                    POINT{anchorRightPx, workArea.top + 48}, 220.0f, 160.0f, dpi, false, nullptr, nullptr, ContextMenuRootHorizontalAlignment::End, popupRect),
                std::format("right-aligned root popup position resolves at {} DPI", dpi).c_str());
        Require(popupRect.right == anchorRightPx, std::format("right-aligned root popup keeps its visible right edge on the anchor at {} DPI", dpi).c_str());
        Require(popupRect.left >= workArea.left && popupRect.top >= workArea.top && popupRect.bottom <= workArea.bottom,
                std::format("right-aligned root popup still stays inside the work area at {} DPI", dpi).c_str());

        const int anchorLeftPx = workArea.left + 1;
        Require(DebugComputeContextMenuPopupPosition(
                    POINT{anchorLeftPx, workArea.top + 92}, 220.0f, 160.0f, dpi, false, nullptr, nullptr, ContextMenuRootHorizontalAlignment::End, popupRect),
                std::format("right-aligned root popup resolves near the left edge at {} DPI", dpi).c_str());
        Require(popupRect.left == anchorLeftPx,
                std::format("right-aligned root popup falls back to a start-aligned placement when trailing alignment would cross the left edge at {} DPI "
                            "(left={}, right={}, anchor={})",
                            dpi,
                            popupRect.left,
                            popupRect.right,
                            anchorLeftPx)
                    .c_str());
        Require(popupRect.right <= workArea.right && popupRect.top >= workArea.top && popupRect.bottom <= workArea.bottom,
                std::format("right-aligned root popup fallback still stays inside the work area at {} DPI", dpi).c_str());
    }
}

void TestMenuPopupPositionSupportsAboveRightAlignedRootAnchors()
{
    using namespace RedSalamander::DxUi;

    const RECT workArea         = GetPrimaryMonitorWorkArea();
    constexpr UINT kDpiMatrix[] = {96u, 120u, 144u, 168u, 192u};

    for (const UINT dpi : kDpiMatrix)
    {
        RECT popupRect{};
        const int anchorRightPx = workArea.left + ((workArea.right - workArea.left) / 2);
        const int anchorTopPx   = workArea.bottom - 96;
        Require(DebugComputeContextMenuPopupPosition(POINT{anchorRightPx, anchorTopPx},
                                                     220.0f,
                                                     160.0f,
                                                     dpi,
                                                     false,
                                                     nullptr,
                                                     nullptr,
                                                     ContextMenuRootHorizontalAlignment::End,
                                                     ContextMenuRootVerticalPlacement::Above,
                                                     popupRect),
                std::format("above/right-aligned root popup position resolves at {} DPI", dpi).c_str());
        Require(popupRect.right == anchorRightPx,
                std::format("above/right-aligned root popup keeps its visible right edge on the anchor at {} DPI", dpi).c_str());
        Require(popupRect.bottom == anchorTopPx,
                std::format("above/right-aligned root popup keeps its visible bottom edge above the anchor at {} DPI", dpi).c_str());
        Require(popupRect.left >= workArea.left && popupRect.top >= workArea.top && popupRect.right <= workArea.right,
                std::format("above/right-aligned root popup stays inside the work area at {} DPI", dpi).c_str());
    }
}

void TestMenuInfoRowsUseMeasuredValueColumnWidth()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Info, .text = L"Total Space:", .acceleratorText = L"1.82 TB (1999377526784 bytes)"},
        {.kind = MenuItemKind::Info, .text = L"Free Space:", .acceleratorText = L"1.27 TB"},
        {.kind = MenuItemKind::Standard, .text = L"Disk Properties...", .commandId = 3001},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for measured info-row column validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupItemLayoutDebugState firstLayout{};
        ContextMenuPopupItemLayoutDebugState secondLayout{};
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, firstLayout) || ! WaitForContextMenuPopupItemLayout(popupHwnd, 1u, secondLayout))
        {
            driverFailure = "info-row popup exposes measured text and value-column layout";
            return;
        }

        const float firstValueWidth  = firstLayout.acceleratorRectDip.right - firstLayout.acceleratorRectDip.left;
        const float secondValueWidth = secondLayout.acceleratorRectDip.right - secondLayout.acceleratorRectDip.left;
        if (firstValueWidth <= 120.5f || secondValueWidth <= 120.5f)
        {
            driverFailure = "info rows expand the measured value column beyond the old fixed-width accelerator slot";
            return;
        }

        if (std::fabs(firstLayout.acceleratorRectDip.left - secondLayout.acceleratorRectDip.left) > 0.5f ||
            std::fabs(firstLayout.acceleratorRectDip.right - secondLayout.acceleratorRectDip.right) > 0.5f)
        {
            driverFailure = "info rows share a stable aligned value column";
            return;
        }

        if (firstLayout.textRectDip.right > firstLayout.acceleratorRectDip.left || secondLayout.textRectDip.right > secondLayout.acceleratorRectDip.left)
        {
            driverFailure = "info-row labels stay separated from the right-aligned value column";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the measured info-row layout menu with Escape returns no invoked command");
}

void TestMenuStandardRowsDeriveAndAlignShortcutColumnFromTabbedText()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Standard, .text = L"&Open\tCtrl+O", .commandId = 3001},
        {.kind = MenuItemKind::Standard, .text = L"E&xit\t  Alt+F4  ", .commandId = 3002},
        {.kind = MenuItemKind::Standard, .text = L"Preferences...", .commandId = 3003},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for tabbed shortcut layout validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupItemLayoutDebugState firstLayout{};
        ContextMenuPopupItemLayoutDebugState secondLayout{};
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, firstLayout) || ! WaitForContextMenuPopupItemLayout(popupHwnd, 1u, secondLayout))
        {
            driverFailure = "standard rows expose shared shortcut-column layout";
            return;
        }

        const float firstShortcutWidth  = firstLayout.acceleratorRectDip.right - firstLayout.acceleratorRectDip.left;
        const float secondShortcutWidth = secondLayout.acceleratorRectDip.right - secondLayout.acceleratorRectDip.left;
        if (firstShortcutWidth <= 12.0f || secondShortcutWidth <= 12.0f)
        {
            driverFailure = "tabbed standard rows allocate a visible right-aligned shortcut column";
            return;
        }

        if (std::fabs(firstLayout.acceleratorRectDip.left - secondLayout.acceleratorRectDip.left) > 0.5f ||
            std::fabs(firstLayout.acceleratorRectDip.right - secondLayout.acceleratorRectDip.right) > 0.5f)
        {
            driverFailure = "tabbed standard rows share a stable aligned shortcut column";
            return;
        }

        if (firstLayout.textRectDip.right > firstLayout.acceleratorRectDip.left || secondLayout.textRectDip.right > secondLayout.acceleratorRectDip.left)
        {
            driverFailure = "tabbed standard row labels stay separated from the derived shortcut column";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the tabbed standard-row layout menu with Escape returns no invoked command");
}

void TestMenuShortcutRowsReserveChevronLaneWhenAnySubmenuExists()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Standard, .text = L"&More", .children = {{.kind = MenuItemKind::Standard, .text = L"Child", .commandId = 3101}}},
        {.kind = MenuItemKind::Standard, .text = L"&Open\tCtrl+O", .commandId = 3102},
        {.kind = MenuItemKind::Standard, .text = L"E&xit\tAlt+F4", .commandId = 3103},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for shared chevron-lane layout validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupItemLayoutDebugState submenuLayout{};
        ContextMenuPopupItemLayoutDebugState firstShortcutLayout{};
        ContextMenuPopupItemLayoutDebugState secondShortcutLayout{};
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, submenuLayout) || ! WaitForContextMenuPopupItemLayout(popupHwnd, 1u, firstShortcutLayout) ||
            ! WaitForContextMenuPopupItemLayout(popupHwnd, 2u, secondShortcutLayout))
        {
            driverFailure = "menu popup exposes chevron and shortcut row layout";
            return;
        }

        const float chevronWidthDip = submenuLayout.chevronRectDip.right - submenuLayout.chevronRectDip.left;
        if (chevronWidthDip <= 8.0f)
        {
            driverFailure = "submenu row exposes a visible chevron lane";
            return;
        }

        if (firstShortcutLayout.acceleratorRectDip.right > submenuLayout.chevronRectDip.left + 0.5f ||
            secondShortcutLayout.acceleratorRectDip.right > submenuLayout.chevronRectDip.left + 0.5f)
        {
            driverFailure = "plain shortcut rows reserve the submenu chevron lane";
            return;
        }

        if (std::fabs(firstShortcutLayout.acceleratorRectDip.left - secondShortcutLayout.acceleratorRectDip.left) > 0.5f ||
            std::fabs(firstShortcutLayout.acceleratorRectDip.right - secondShortcutLayout.acceleratorRectDip.right) > 0.5f)
        {
            driverFailure = "plain shortcut rows keep the shared shortcut column aligned";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the shared chevron-lane layout menu with Escape returns no invoked command");
}

void TestMenuBitmapIconsReachPopupLayout()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    std::vector<MenuFlyoutItem> items;
    items.push_back(MenuFlyoutItem{.text = L"Downloads", .iconBitmap = CreateSyntheticMenuBitmapIcon(16u), .commandId = 3001});

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for bitmap-icon validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupItemLayoutDebugState layout{};
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, layout))
        {
            driverFailure = "menu popup exposes layout for bitmap-icon item";
            return;
        }

        if (! layout.hasBitmapIcon)
        {
            driverFailure = "bitmap icon payload is preserved for popup rendering";
            return;
        }

        if (layout.iconRectDip.left < (layout.itemRectDip.left + 3.5f))
        {
            driverFailure = "bitmap icon slot begins inside the hovered row instead of before it";
            return;
        }

        if (layout.iconRectDip.right > (layout.textRectDip.left - 6.0f))
        {
            driverFailure = "bitmap icon slot leaves spacing before the menu text column";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the bitmap-icon menu with Escape returns no invoked command");
}

void TestMenuPopupMaterialsProduceDistinctCaptures()
{
    using namespace RedSalamander::DxUi;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 3001},
        {.text = L"Rename", .commandId = 3002},
        {.kind = MenuItemKind::Separator},
        {.text = L"Properties", .commandId = 3003},
    };

    ThemePalette micaTheme    = MakeDefaultThemePalette(true);
    micaTheme.overlayMaterial = OverlayMaterial::Mica;

    ThemePalette micaAltTheme    = micaTheme;
    micaAltTheme.overlayMaterial = OverlayMaterial::MicaAlt;

    ThemePalette acrylicTheme    = micaTheme;
    acrylicTheme.overlayMaterial = OverlayMaterial::Acrylic;

    const WindowHostBitmapCapture micaCapture    = CaptureMenuPopupBitmapForTheme(micaTheme, items);
    const WindowHostBitmapCapture micaAltCapture = CaptureMenuPopupBitmapForTheme(micaAltTheme, items);
    const WindowHostBitmapCapture acrylicCapture = CaptureMenuPopupBitmapForTheme(acrylicTheme, items);

    const BitmapComparisonStats micaVsMicaAlt = CompareWindowHostBitmapCapturesForTest(micaCapture, micaAltCapture, 4u);
    Require(micaVsMicaAlt.totalPixels > 0u, "menu material captures share the same geometry for comparison");
    Require(micaVsMicaAlt.DifferenceRatio() > 0.01, "Mica and Mica Alt popup materials produce visibly different menu captures");

    const BitmapComparisonStats micaVsAcrylic = CompareWindowHostBitmapCapturesForTest(micaCapture, acrylicCapture, 4u);
    Require(micaVsAcrylic.totalPixels > 0u, "menu acrylic capture shares the same geometry for comparison");
    Require(micaVsAcrylic.DifferenceRatio() > 0.01, "Mica and Acrylic popup materials produce visibly different menu captures");
}

void TestMenuRainbowHoverUsesSeededHighlightContrast()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.rainbowMode     = true;
    theme.overlayMaterial = OverlayMaterial::Solid;

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Radio, .text = L"Rainbow", .checked = true, .commandId = 8051},
        {.kind = MenuItemKind::Radio, .text = L"Forest Mist", .checked = false, .commandId = 8052},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    ContextMenuPopupDebugState popupState{};
    ContextMenuPopupItemPaintDebugState paintState{};
    D2D1_RECT_F firstRowRectDip = D2D1::RectF();
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for rainbow hover highlight validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        if (! WaitForContextMenuPopupItemRect(popupHwnd, 0u, firstRowRectDip))
        {
            driverFailure = "menu popup exposes row bounds for rainbow hover highlight validation";
            return;
        }

        const LONG hoverX = static_cast<LONG>(std::lround((firstRowRectDip.left + firstRowRectDip.right) * 0.5f));
        const LONG hoverY = static_cast<LONG>(std::lround((firstRowRectDip.top + firstRowRectDip.bottom) * 0.5f));
        if (! SendSettledClientMouseMoveForMenuSuite(popupHwnd, hoverX, hoverY))
        {
            driverFailure = "moving the pointer onto the rainbow row succeeds";
            return;
        }

        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "moving the pointer onto the rainbow row sets popup hover";
            return;
        }

        if (! WaitForContextMenuPopupItemPaint(popupHwnd, 0u, paintState))
        {
            driverFailure = "menu popup exposes hovered paint state for rainbow highlight validation";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the rainbow hover highlight validation menu with Escape returns no invoked command");

    const D2D1_COLOR_F expectedFill       = RainbowMenuSelectionTint(L"Rainbow", theme.dark);
    const D2D1_COLOR_F expectedForeground = ChooseContrastingTextColor(expectedFill);
    Require(paintState.hovered, "rainbow hover highlight paint state reports the row as hovered");
    Require(paintState.usesHighlightFill, "rainbow hover highlight paint state reports the highlight backplate");
    Require(paintState.usesRainbowHighlight, "rainbow hover highlight paint state reports the seeded rainbow backplate");
    RequireColorNear(paintState.fillColor, expectedFill, "rainbow hover highlight uses the stable seeded menu rainbow fill");
    RequireColorNear(paintState.compositeFillColor, expectedFill, "opaque rainbow hover highlight keeps its composite fill identical to the seeded color");
    RequireColorNear(paintState.textColor, expectedForeground, "rainbow hover highlight text uses the contrasting foreground");
    RequireColorNear(paintState.checkColor, expectedForeground, "rainbow hover highlight radio indicator uses the contrasting foreground");
}

void TestMenuHoverContrastAppliesToGlyphsAcrossThemes()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme      = MakeDefaultThemePalette(false);
    theme.rainbowMode       = false;
    theme.highContrast      = false;
    theme.overlayMaterial   = OverlayMaterial::Acrylic;
    theme.overlayBackground = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    theme.headerHovered     = D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f);
    theme.accent            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f);
    theme.text              = D2D1::ColorF(0.98f, 0.98f, 0.98f, 1.0f);
    theme.subduedText       = theme.text;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open With\tCtrl+Enter", .iconGlyph = L"\uE8A7", .commandId = 8151, .children = {MenuFlyoutItem{.text = L"Viewer", .commandId = 8152}}},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    ContextMenuPopupDebugState popupState{};
    ContextMenuPopupItemPaintDebugState paintState{};
    D2D1_RECT_F firstRowRectDip = D2D1::RectF();
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for hover foreground contrast validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        if (! WaitForContextMenuPopupItemRect(popupHwnd, 0u, firstRowRectDip))
        {
            driverFailure = "menu popup exposes row bounds for hover foreground contrast validation";
            return;
        }

        const LONG hoverX = static_cast<LONG>(std::lround((firstRowRectDip.left + firstRowRectDip.right) * 0.5f));
        const LONG hoverY = static_cast<LONG>(std::lround((firstRowRectDip.top + firstRowRectDip.bottom) * 0.5f));
        if (! SendSettledClientMouseMoveForMenuSuite(popupHwnd, hoverX, hoverY))
        {
            driverFailure = "moving the pointer onto the contrast row succeeds";
            return;
        }

        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "moving the pointer onto the custom-theme row sets popup hover";
            return;
        }

        if (! WaitForContextMenuPopupItemPaint(popupHwnd, 0u, paintState))
        {
            driverFailure = "menu popup exposes hovered paint state for hover foreground contrast validation";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the hover foreground contrast validation menu with Escape returns no invoked command");

    const D2D1_COLOR_F expectedForeground = ChooseContrastingTextColor(paintState.compositeFillColor);
    RequireColorDifferent(theme.text, expectedForeground, "custom non-rainbow hover validation requires a foreground different from the base text token");
    Require(! paintState.usesRainbowHighlight, "non-rainbow hover validation keeps the shared highlight path out of rainbow mode");
    RequireColorNear(paintState.textColor, expectedForeground, "non-rainbow menu hover text uses the contrasting foreground");
    RequireColorNear(paintState.acceleratorColor, expectedForeground, "non-rainbow menu hover accelerator uses the contrasting foreground");
    RequireColorNear(paintState.iconColor, expectedForeground, "non-rainbow menu hover icon glyph uses the contrasting foreground");
    RequireColorNear(paintState.chevronColor, expectedForeground, "non-rainbow menu hover submenu chevron uses the contrasting foreground");
}

void TestMenuRainbowCheckedItemUsesAccentIndicator()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.rainbowMode     = true;
    theme.overlayMaterial = OverlayMaterial::Solid;

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Radio, .text = L"Rainbow", .checked = true, .commandId = 8101},
        {.kind = MenuItemKind::Radio, .text = L"High Contrast", .checked = false, .commandId = 8102},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    WindowHostBitmapCapture capture{};
    ContextMenuPopupItemLayoutDebugState checkedLayout{};
    ContextMenuPopupItemLayoutDebugState uncheckedLayout{};
    ContextMenuPopupDebugState popupState{};
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for rainbow checked-row validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState&) noexcept { return true; }, popupState))
        {
            driverFailure = "menu popup exposes debug state for rainbow checked-row validation";
            return;
        }
        if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
        {
            driverFailure = "menu popup bitmap capture succeeds for rainbow checked-row validation";
            return;
        }
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, checkedLayout))
        {
            driverFailure = "menu popup exposes checked-row layout for rainbow checked-row validation";
            return;
        }
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 1u, uncheckedLayout))
        {
            driverFailure = "menu popup exposes unchecked-row layout for rainbow checked-row validation";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the rainbow checked-row menu with Escape returns no invoked command");

    const float iconCenterXDip = (checkedLayout.iconRectDip.left + checkedLayout.iconRectDip.right) * 0.5f;
    const UINT sampleX         = (std::min)(capture.widthPx - 1u, DipToPixelForPopup(iconCenterXDip, popupState.dpi));
    const UINT checkedY =
        (std::min)(capture.heightPx - 1u, DipToPixelForPopup((checkedLayout.iconRectDip.top + checkedLayout.iconRectDip.bottom) * 0.5f, popupState.dpi));
    const UINT uncheckedY =
        (std::min)(capture.heightPx - 1u, DipToPixelForPopup((uncheckedLayout.iconRectDip.top + uncheckedLayout.iconRectDip.bottom) * 0.5f, popupState.dpi));

    Require(GetCapturePixelBgra(capture, sampleX, checkedY) != GetCapturePixelBgra(capture, sampleX, uncheckedY),
            "rainbow checked menu rows render an accent-colored check indicator distinct from unchecked rows");
}

void TestMenuCheckedRowsDoNotPaintSecondFullRowSelection()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.rainbowMode     = true;
    theme.overlayMaterial = OverlayMaterial::Solid;

    const std::vector<MenuFlyoutItem> uncheckedItems = {
        {.kind = MenuItemKind::Radio, .text = L"Rainbow", .checked = false, .commandId = 8201},
        {.kind = MenuItemKind::Radio, .text = L"Forest Mist", .checked = false, .commandId = 8202},
        {.kind = MenuItemKind::Radio, .text = L"Neon Tokyo", .checked = false, .commandId = 8203},
    };
    const std::vector<MenuFlyoutItem> checkedItems = {
        {.kind = MenuItemKind::Radio, .text = L"Rainbow", .checked = true, .commandId = 8201},
        {.kind = MenuItemKind::Radio, .text = L"Forest Mist", .checked = false, .commandId = 8202},
        {.kind = MenuItemKind::Radio, .text = L"Neon Tokyo", .checked = false, .commandId = 8203},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const auto captureTrailingPixel = [&](const std::vector<MenuFlyoutItem>& items, uint32_t& trailingPixel, std::string& failure) noexcept
    {
        WindowHostBitmapCapture capture{};
        D2D1_RECT_F firstRowRectDip = D2D1::RectF();
        D2D1_RECT_F hoverRowRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        std::thread driver([&]
        {
            const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
            if (! popupHwnd)
            {
                failure = "menu popup window appears for checked-row selection validation";
                return;
            }

            const auto dismissPopup = wil::scope_exit([&]() noexcept
            {
                if (IsWindow(popupHwnd) != FALSE)
                {
                    PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
                }
            });

            if (! WaitForContextMenuPopupItemRect(popupHwnd, 0u, firstRowRectDip) || ! WaitForContextMenuPopupItemRect(popupHwnd, 1u, hoverRowRectDip))
            {
                failure = "menu popup exposes row bounds for checked-row selection validation";
                return;
            }

            const LONG hoverX = static_cast<LONG>(std::lround((hoverRowRectDip.left + hoverRowRectDip.right) * 0.5f));
            const LONG hoverY = static_cast<LONG>(std::lround((hoverRowRectDip.top + hoverRowRectDip.bottom) * 0.5f));
            if (! SendSettledClientMouseMoveForMenuSuite(popupHwnd, hoverX, hoverY))
            {
                failure = "moving the pointer to a plain row before checked-row selection capture succeeds";
                return;
            }

            if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
                return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
            }, popupState))
            {
                failure = "moving the pointer to a plain row transfers popup hover there";
                return;
            }

            if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
            {
                failure = "menu popup bitmap capture succeeds for checked-row selection validation";
            }
        });

        const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
        driver.join();

        if (! failure.empty())
        {
            return;
        }

        if (result.has_value())
        {
            failure = "closing the checked-row selection validation menu with Escape returns no invoked command";
            return;
        }

        const UINT sampleX = (std::min)(capture.widthPx - 1u, DipToPixelForPopup(firstRowRectDip.right - 18.0f, popupState.dpi));
        const UINT sampleY = (std::min)(capture.heightPx - 1u, DipToPixelForPopup((firstRowRectDip.top + firstRowRectDip.bottom) * 0.5f, popupState.dpi));
        trailingPixel      = GetCapturePixelBgra(capture, sampleX, sampleY);
    };

    uint32_t uncheckedTrailingPixel = 0u;
    uint32_t checkedTrailingPixel   = 0u;
    std::string driverFailure;
    captureTrailingPixel(uncheckedItems, uncheckedTrailingPixel, driverFailure);
    Require(driverFailure.empty(), driverFailure.c_str());
    captureTrailingPixel(checkedItems, checkedTrailingPixel, driverFailure);
    Require(driverFailure.empty(), driverFailure.c_str());

    Require(uncheckedTrailingPixel == checkedTrailingPixel,
            "checked menu rows do not paint a second full-width selection backplate when another row is hovered");
}

void TestMenuCheckedRowsDoNotPaintLeadingCheckedBox()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.overlayMaterial = OverlayMaterial::Solid;

    const std::vector<MenuFlyoutItem> uncheckedItems = {
        {.kind = MenuItemKind::Toggle, .text = L"Show Hidden Files", .checked = false, .commandId = 8301},
        {.kind = MenuItemKind::Toggle, .text = L"Show System Files", .checked = false, .commandId = 8302},
    };
    const std::vector<MenuFlyoutItem> checkedItems = {
        {.kind = MenuItemKind::Toggle, .text = L"Show Hidden Files", .checked = true, .commandId = 8301},
        {.kind = MenuItemKind::Toggle, .text = L"Show System Files", .checked = false, .commandId = 8302},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const auto captureLeadingPixel = [&](const std::vector<MenuFlyoutItem>& items, uint32_t& leadingPixel, std::string& failure) noexcept
    {
        WindowHostBitmapCapture capture{};
        ContextMenuPopupItemLayoutDebugState firstLayout{};
        D2D1_RECT_F hoverRowRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        std::thread driver([&]
        {
            const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
            if (! popupHwnd)
            {
                failure = "menu popup window appears for checked-indicator box validation";
                return;
            }

            const auto dismissPopup = wil::scope_exit([&]() noexcept
            {
                if (IsWindow(popupHwnd) != FALSE)
                {
                    PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
                }
            });

            if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, firstLayout) || ! WaitForContextMenuPopupItemRect(popupHwnd, 1u, hoverRowRectDip))
            {
                failure = "menu popup exposes leading indicator geometry for checked-indicator box validation";
                return;
            }

            const LONG hoverX = static_cast<LONG>(std::lround((hoverRowRectDip.left + hoverRowRectDip.right) * 0.5f));
            const LONG hoverY = static_cast<LONG>(std::lround((hoverRowRectDip.top + hoverRowRectDip.bottom) * 0.5f));
            if (! SendSettledClientMouseMoveForMenuSuite(popupHwnd, hoverX, hoverY))
            {
                failure = "moving the pointer away from the checked row before checked-indicator capture succeeds";
                return;
            }

            if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
                return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
            }, popupState))
            {
                failure = "moving the pointer away from the checked row transfers popup hover to the plain row";
                return;
            }

            if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
            {
                failure = "menu popup bitmap capture succeeds for checked-indicator box validation";
            }
        });

        const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
        driver.join();

        if (! failure.empty())
        {
            return;
        }

        if (result.has_value())
        {
            failure = "closing the checked-indicator box validation menu with Escape returns no invoked command";
            return;
        }

        const UINT sampleX = (std::min)(capture.widthPx - 1u, DipToPixelForPopup(firstLayout.iconRectDip.left + 2.0f, popupState.dpi));
        const UINT sampleY =
            (std::min)(capture.heightPx - 1u, DipToPixelForPopup((firstLayout.iconRectDip.top + firstLayout.iconRectDip.bottom) * 0.5f, popupState.dpi));
        leadingPixel = GetCapturePixelBgra(capture, sampleX, sampleY);
    };

    uint32_t uncheckedLeadingPixel = 0u;
    uint32_t checkedLeadingPixel   = 0u;
    std::string driverFailure;
    captureLeadingPixel(uncheckedItems, uncheckedLeadingPixel, driverFailure);
    Require(driverFailure.empty(), driverFailure.c_str());
    captureLeadingPixel(checkedItems, checkedLeadingPixel, driverFailure);
    Require(driverFailure.empty(), driverFailure.c_str());

    Require(uncheckedLeadingPixel == checkedLeadingPixel, "checked menu rows do not paint a rounded leading box behind the check indicator");
}

void TestMenuPopupCompositionHostUsesTransparentShadowMargins()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Desktop", .commandId = 3001},
        {.text = L"Documents", .commandId = 3002},
        {.text = L"Downloads", .commandId = 3003},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for composition shadow-margin validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        WindowHostBitmapCapture capture{};
        if (! WaitForContextMenuPopupState(popupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
        {
            driverFailure = "composition popup exposes geometry and bitmap capture";
            return;
        }

        const double scale         = static_cast<double>(popupState.dpi) / 96.0;
        const UINT visibleWidthPx  = static_cast<UINT>((std::max)(1ll, std::llround(static_cast<double>(popupState.visibleWidthDip) * scale)));
        const UINT visibleHeightPx = static_cast<UINT>((std::max)(1ll, std::llround(static_cast<double>(popupState.visibleHeightDip) * scale)));
        if (capture.widthPx <= visibleWidthPx || capture.heightPx <= visibleHeightPx)
        {
            driverFailure = "composition popup capture includes transparent window margins outside the visible menu surface";
            return;
        }

        const uint8_t topLeftAlpha       = GetCapturePixelAlpha(capture, 0u, 0u);
        const uint8_t topRightAlpha      = GetCapturePixelAlpha(capture, capture.widthPx - 1u, 0u);
        const uint8_t centerAlpha        = GetCapturePixelAlpha(capture, capture.widthPx / 2u, capture.heightPx / 2u);
        const UINT visibleLeftPx         = static_cast<UINT>((std::max)(1ll, std::llround(10.0 * scale)));
        const UINT visibleTopPx          = static_cast<UINT>((std::max)(1ll, std::llround(8.0 * scale)));
        const UINT visibleCornerProbeX   = (std::min)(capture.widthPx - 1u, visibleLeftPx + 2u);
        const UINT visibleCornerProbeY   = (std::min)(capture.heightPx - 1u, visibleTopPx + 2u);
        const uint8_t visibleCornerAlpha = GetCapturePixelAlpha(capture, visibleCornerProbeX, visibleCornerProbeY);
        if (topLeftAlpha > 8u || topRightAlpha > 8u)
        {
            driverFailure = "composition popup keeps outer window corners transparent";
            return;
        }
        if (centerAlpha < 32u)
        {
            driverFailure = "composition popup keeps the menu surface itself visibly opaque inside the transparent window";
            return;
        }
        if (visibleCornerAlpha < 64u)
        {
            driverFailure = "composition popup keeps a RoundSmall-style near-corner body pixel inside the visible menu surface instead of clipping too deeply";
            return;
        }

        RECT popupRect{};
        if (GetWindowRect(popupHwnd, &popupRect) == FALSE)
        {
            driverFailure = "composition popup exposes a valid window rectangle for region validation";
            return;
        }

        const int popupWidthPx  = popupRect.right - popupRect.left;
        const int popupHeightPx = popupRect.bottom - popupRect.top;
        if (popupWidthPx <= 0 || popupHeightPx <= 0)
        {
            driverFailure = "composition popup keeps a positive-sized window rectangle";
            return;
        }

        wil::unique_hrgn windowRegion(CreateRectRgn(0, 0, 0, 0));
        if (! windowRegion)
        {
            driverFailure = "composition popup test can allocate a region handle";
            return;
        }

        if (GetWindowRgn(popupHwnd, windowRegion.get()) == ERROR)
        {
            driverFailure = "composition popup exposes a non-rectangular host window region";
            return;
        }

        if (PtInRegion(windowRegion.get(), 0, 0) != FALSE || PtInRegion(windowRegion.get(), popupWidthPx - 1, 0) != FALSE)
        {
            driverFailure = "composition popup window region clips the top window corners";
            return;
        }

        if (PtInRegion(windowRegion.get(), 1, popupHeightPx / 2) == FALSE || PtInRegion(windowRegion.get(), popupWidthPx - 2, popupHeightPx / 2) == FALSE ||
            PtInRegion(windowRegion.get(), popupWidthPx / 2, 1) == FALSE)
        {
            driverFailure = "composition popup window region keeps the outer host margins available for the popup shadow";
            return;
        }

        const int topShadowShoulderX = (std::max)(1, static_cast<int>(std::llround(3.0 * scale)));
        const int topShadowShoulderY = (std::max)(1, static_cast<int>(std::llround(8.0 * scale)));
        if (PtInRegion(windowRegion.get(), topShadowShoulderX, topShadowShoulderY) == FALSE ||
            PtInRegion(windowRegion.get(), popupWidthPx - 1 - topShadowShoulderX, topShadowShoulderY) == FALSE)
        {
            driverFailure =
                "composition popup window region preserves the thinner top shadow shoulder instead of over-rounding it to match the thicker bottom margin";
            return;
        }

        if (PtInRegion(windowRegion.get(), popupWidthPx / 2, popupHeightPx / 2) == FALSE)
        {
            driverFailure = "composition popup window region still contains the menu body";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the composition-shadow menu with Escape returns no invoked command");
}

void TestMenuPopupWindowClassDoesNotUseNativeDropShadow()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 4101},
        {.text = L"Copy", .commandId = 4102},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for native-drop-shadow validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        const ULONG_PTR classStyle = GetClassLongPtrW(popupHwnd, GCL_STYLE);
        if ((classStyle & CS_DROPSHADOW) != 0u)
        {
            driverFailure = "composition popup window class does not opt into the native CS_DROPSHADOW effect";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the native-drop-shadow validation menu with Escape returns no invoked command");
}

void TestMenuPopupAcrylicLightVisualBaseline()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(false);
    theme.overlayMaterial = OverlayMaterial::Acrylic;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 4201},
        {.text = L"Copy", .commandId = 4202},
        {.text = L"Move", .commandId = 4203},
        {.text = L"Properties", .commandId = 4204},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ownerWindow.Host().SetRoot(std::make_unique<StripedBackdropControl>(10));
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();
    DrainPendingMouseMessagesForMenuSuite();
    WindowHostBitmapCapture ownerBackdropCapture{};
    Require(ownerWindow.Host().DebugCaptureBitmap(ownerBackdropCapture), "acrylic visual-baseline owner backdrop renders before popup capture");
    const POINT menuPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 72, 48, "acrylic baseline popup anchor maps to screen coordinates");

    WindowHostBitmapCapture capture{};
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for acrylic visual-baseline capture";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f && ! state.hoveredIndex.has_value();
        }, popupState))
        {
            driverFailure = "menu popup exposes unhovered acrylic geometry for visual-baseline capture";
            return;
        }

        POINT popupClientOrigin{menuPoint.x, menuPoint.y};
        if (ScreenToClient(ownerWindow.Hwnd(), &popupClientOrigin) == FALSE)
        {
            driverFailure = "acrylic visual-baseline popup origin maps to owner client coordinates";
            return;
        }

        const UINT surfaceWidthPx  = DipToPixelForPopup(popupState.visibleWidthDip, popupState.dpi);
        const UINT surfaceHeightPx = DipToPixelForPopup(popupState.visibleHeightDip, popupState.dpi);
        const RECT backdropCropRect{
            popupClientOrigin.x,
            popupClientOrigin.y,
            popupClientOrigin.x + static_cast<LONG>(surfaceWidthPx),
            popupClientOrigin.y + static_cast<LONG>(surfaceHeightPx),
        };
        const WindowHostBitmapCapture deterministicBackdrop = CropWindowHostBitmapCaptureForTest(ownerBackdropCapture, backdropCropRect);
        if (! DebugSetContextMenuPopupBackdropCapture(popupHwnd, deterministicBackdrop))
        {
            driverFailure = "menu popup accepts deterministic acrylic backdrop for visual-baseline capture";
            return;
        }

        if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
        {
            driverFailure = "menu popup bitmap capture succeeds for acrylic visual-baseline capture";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), menuPoint, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the acrylic visual-baseline menu with Escape returns no invoked command");
    VerifyOrUpdateBaselineForTest("acrylic light menu popup baseline matches", L"menu_popup_acrylic_light.png", capture, 0.06, 10u);
}

void TestMenuPopupKeepsSystemBackdropDisabledForAppRenderedMaterials()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.overlayMaterial = OverlayMaterial::Acrylic;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 3001},
        {.text = L"Rename", .commandId = 3002},
        {.text = L"Properties", .commandId = 3003},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for popup-backdrop validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f;
        }, popupState))
        {
            driverFailure = "menu popup exposes debug state for popup-backdrop validation";
            return;
        }

        if (popupState.usesSystemBackdrop)
        {
            driverFailure = "app-rendered popup materials keep the popup HWND free of DWM system backdrops";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the popup-backdrop validation menu with Escape returns no invoked command");
}

void TestMenuPopupSubmenuFlipsLeftNearRightEdge()
{
    using namespace RedSalamander::DxUi;

    const RECT workArea   = GetPrimaryMonitorWorkArea();
    const RECT parentRect = {
        workArea.right - 220,
        workArea.top + 80,
        workArea.right - 24,
        workArea.top + 260,
    };
    const RECT parentItemRect = {
        workArea.right - 34,
        workArea.top + 112,
        workArea.right - 12,
        workArea.top + 146,
    };

    RECT popupRect{};
    Require(DebugComputeContextMenuPopupPosition(POINT{parentItemRect.right, parentItemRect.top},
                                                 200.0f,
                                                 160.0f,
                                                 144u,
                                                 true,
                                                 &parentRect,
                                                 &parentItemRect,
                                                 ContextMenuRootHorizontalAlignment::Start,
                                                 popupRect),
            "submenu popup position resolves near the right screen edge");
    Require(popupRect.right <= parentItemRect.left, "submenu popup flips to the left side when opening right would overflow the work area");
    Require(popupRect.left >= workArea.left && popupRect.top >= workArea.top && popupRect.bottom <= workArea.bottom,
            "flipped submenu popup remains fully inside the monitor work area");
}

void TestMenuMnemonicHonorsExplicitAmpersandLabels()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Save && Close\tCtrl+S", .commandId = 3001},
        {.text = L"E&xit\tAlt+F4", .commandId = 3002},
        {.text = L"Rock &", .commandId = 3003},
    };
    std::wstring normalizedFirstItemText;
    const bool normalizedFirstItem = DebugGetContextMenuItemDisplayText(items[0], normalizedFirstItemText);
    if (! normalizedFirstItem || normalizedFirstItemText != L"Save & Close")
    {
        const std::string failure =
            std::format("menu item parser preserves escaped ampersands while stripping tabbed accelerator text actual='{}' codeUnits='{}'",
                        NarrowAsciiForFailureMessage(normalizedFirstItemText),
                        WideCodeUnitsForFailureMessage(normalizedFirstItemText));
        Require(false, failure.c_str());
    }

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Save & Close");
        if (! popupHwnd)
        {
            std::wstring popupFirstItemText;
            if (HWND fallbackPopupHwnd = FindOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
                fallbackPopupHwnd && RedSalamander::DxUi::DebugGetContextMenuPopupItemText(fallbackPopupHwnd, 0u, popupFirstItemText))
            {
                driverFailure = std::format("menu popup window appears for explicit ampersand mnemonic validation actual='{}' codeUnits='{}'",
                                            NarrowAsciiForFailureMessage(popupFirstItemText),
                                            WideCodeUnitsForFailureMessage(popupFirstItemText));
            }
            else
            {
                driverFailure = "menu popup window appears for explicit ampersand mnemonic validation";
            }
            return;
        }

        std::wstring firstItemText;
        std::wstring secondItemText;
        std::wstring thirdItemText;
        if (! WaitForContextMenuPopupItemText(popupHwnd, 0u, firstItemText) || ! WaitForContextMenuPopupItemText(popupHwnd, 1u, secondItemText) ||
            ! WaitForContextMenuPopupItemText(popupHwnd, 2u, thirdItemText))
        {
            driverFailure = "menu popup exposes item text for explicit ampersand mnemonic validation";
            return;
        }

        if (firstItemText != L"Save & Close")
        {
            driverFailure = std::format("menu popup preserves escaped ampersands while stripping tabbed accelerator text actual='{}' codeUnits='{}'",
                                        NarrowAsciiForFailureMessage(firstItemText),
                                        WideCodeUnitsForFailureMessage(firstItemText));
            return;
        }

        if (secondItemText != L"Exit")
        {
            driverFailure = std::format("menu popup strips explicit ampersand markers from displayed item text actual='{}' codeUnits='{}'",
                                        NarrowAsciiForFailureMessage(secondItemText),
                                        WideCodeUnitsForFailureMessage(secondItemText));
            return;
        }

        if (thirdItemText != L"Rock &")
        {
            driverFailure = std::format("menu popup preserves a trailing literal ampersand in displayed item text actual='{}' codeUnits='{}'",
                                        NarrowAsciiForFailureMessage(thirdItemText),
                                        WideCodeUnitsForFailureMessage(thirdItemText));
            return;
        }

        if (! SendMenuKeyForMenuSuite(popupHwnd, 'X'))
        {
            driverFailure = "explicit ampersand mnemonic key is delivered to the popup";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(result.has_value() && result.value() == 3002, "explicit ampersand mnemonic invokes the matching popup item");
}

void TestMenuOpeningPointerUpCanBeIgnoredOutsideVisibleSurface()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Ignore Release", .commandId = 3001},
        {.text = L"Close", .commandId = 3002},
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.ignoreInitialLeftButtonUp = true;

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Ignore Release");
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for initial pointer-up validation";
            return;
        }

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f;
        }, popupState))
        {
            driverFailure = "menu popup exposes geometry for initial pointer-up validation";
            return;
        }

        PostMessageW(popupHwnd, WM_LBUTTONUP, 0, MAKELPARAM(1, 1));
        std::this_thread::sleep_for(std::chrono::milliseconds(80));

        if (IsWindow(popupHwnd) == FALSE)
        {
            driverFailure = "menu popup ignores the opening button-up before applying light-dismiss rules";
            return;
        }

        PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true), sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the initial pointer-up validation menu returns no invoked command");
}

void TestMenuShadowMarginMouseUpLightDismissesAfterInitialRelease()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Shadow Desktop", .commandId = 3001},
        {.text = L"Documents", .commandId = 3002},
        {.text = L"Downloads", .commandId = 3003},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Shadow Desktop");
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for shadow-margin light-dismiss validation";
            return;
        }

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f;
        }, popupState))
        {
            driverFailure = "menu popup exposes geometry for shadow-margin light-dismiss validation";
            return;
        }

        PostMessageW(popupHwnd, WM_LBUTTONUP, 0, MAKELPARAM(1, 1));
        if (! WaitForWindowDestroyed(popupHwnd))
        {
            driverFailure = "menu popup treats transparent shadow margins as outside the visible menu surface for light-dismiss";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "shadow-margin light-dismiss returns no invoked command");
}

void TestMenuAcrylicBackdropScenarioEmitsMetrics()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 96, 96, 480, 360, SWP_NOZORDER | SWP_NOACTIVATE);
    ownerWindow.Host().SetRoot(std::make_unique<StripedBackdropControl>(12));
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    UpdateWindow(ownerWindow.Hwnd());
    ownerWindow.PumpMessages();
    WindowHostBitmapCapture ownerBackdropBeforePopup{};
    Require(ownerWindow.Host().DebugCaptureBitmap(ownerBackdropBeforePopup), "owner backdrop renders before acrylic metric capture");
    const POINT menuPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 96, 72, "acrylic metric popup anchor maps to screen coordinates");

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.overlayMaterial = OverlayMaterial::Acrylic;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 3001},
        {.text = L"Copy", .commandId = 3002},
        {.text = L"Move", .commandId = 3003},
    };

    std::string driverFailure;
    WindowHostBitmapCapture popupCapture{};
    uint64_t openToCaptureUs = 0u;
    const auto startedAt     = std::chrono::steady_clock::now();

    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for acrylic backdrop metric capture";
            return;
        }
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(
                popupHwnd, [](const ContextMenuPopupDebugState& state) { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; }, popupState))
        {
            driverFailure = "menu popup enables app-rendered backdrop blur for acrylic backdrop metric capture";
            return;
        }

        POINT popupClientOrigin{menuPoint.x, menuPoint.y};
        if (ScreenToClient(ownerWindow.Hwnd(), &popupClientOrigin) == FALSE)
        {
            driverFailure = "acrylic backdrop metric popup origin maps to owner client coordinates";
            return;
        }

        const UINT surfaceWidthPx  = DipToPixelForPopup(popupState.visibleWidthDip, popupState.dpi);
        const UINT surfaceHeightPx = DipToPixelForPopup(popupState.visibleHeightDip, popupState.dpi);
        const RECT backdropCropRect{
            popupClientOrigin.x,
            popupClientOrigin.y,
            popupClientOrigin.x + static_cast<LONG>(surfaceWidthPx),
            popupClientOrigin.y + static_cast<LONG>(surfaceHeightPx),
        };
        const WindowHostBitmapCapture deterministicBackdrop = CropWindowHostBitmapCaptureForTest(ownerBackdropBeforePopup, backdropCropRect);
        if (! DebugSetContextMenuPopupBackdropCapture(popupHwnd, deterministicBackdrop))
        {
            driverFailure = "menu popup accepts deterministic acrylic backdrop for metric capture";
            return;
        }

        if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, popupCapture))
        {
            driverFailure = "menu popup bitmap capture succeeds for acrylic backdrop metric capture";
            return;
        }

        openToCaptureUs = Debug::Perf::ElapsedUs(startedAt);
        PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), menuPoint, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the acrylic backdrop metric menu with Escape returns no invoked command");
    Require(popupCapture.widthPx > 0u && popupCapture.heightPx > 0u && ! popupCapture.bgraPixels.empty(),
            "acrylic backdrop metric scenario captures the popup bitmap");

    const RECT opaqueBounds = FindOpaqueBoundsInCapture(popupCapture);
    Require(opaqueBounds.right > opaqueBounds.left && opaqueBounds.bottom > opaqueBounds.top, "acrylic backdrop metric scenario resolves opaque popup bounds");
    const RECT sampleRect = ComputeRightStripSampleRect(opaqueBounds);
    Require(sampleRect.right > sampleRect.left && sampleRect.bottom > sampleRect.top,
            "acrylic backdrop metric scenario resolves a valid right-strip sample region");

    std::this_thread::sleep_for(std::chrono::milliseconds(30));
    InvalidateRect(ownerWindow.Hwnd(), nullptr, FALSE);
    UpdateWindow(ownerWindow.Hwnd());
    ownerWindow.PumpMessages();
    WindowHostBitmapCapture ownerBackdropCapture{};
    Require(ownerWindow.Host().DebugCaptureBitmap(ownerBackdropCapture), "owner backdrop capture succeeds for acrylic metrics");
    POINT popupClientOrigin{menuPoint.x, menuPoint.y};
    Require(ScreenToClient(ownerWindow.Hwnd(), &popupClientOrigin) != FALSE, "popup origin maps to owner client coordinates for acrylic metrics");
    const RECT rawCaptureRect = {
        popupClientOrigin.x - opaqueBounds.left,
        popupClientOrigin.y - opaqueBounds.top,
        popupClientOrigin.x - opaqueBounds.left + static_cast<LONG>(popupCapture.widthPx),
        popupClientOrigin.y - opaqueBounds.top + static_cast<LONG>(popupCapture.heightPx),
    };
    const WindowHostBitmapCapture rawCapture = CropWindowHostBitmapCaptureForTest(ownerBackdropCapture, rawCaptureRect);
    Require(rawCapture.widthPx == popupCapture.widthPx && rawCapture.heightPx == popupCapture.heightPx,
            "raw backdrop capture matches the popup capture geometry for acrylic metrics");

    const uint64_t popupAdjacentDelta = ComputeAverageAdjacentRgbDelta(popupCapture, sampleRect);
    const uint64_t rawAdjacentDelta   = ComputeAverageAdjacentRgbDelta(rawCapture, sampleRect);
    const uint64_t popupVsRawDelta    = ComputeAverageAbsoluteRgbDeltaBetweenCaptures(popupCapture, rawCapture, sampleRect);
    const uint64_t popupVsSlabDelta   = ComputeAverageAbsoluteRgbDeltaToColor(popupCapture, sampleRect, ResolveExpectedAcrylicMenuSlabColor(theme));
    const uint64_t popupToRawDeltaPermille =
        rawAdjacentDelta == 0u ? 0u : static_cast<uint64_t>((popupAdjacentDelta * 1000u + (rawAdjacentDelta / 2u)) / rawAdjacentDelta);
    const uint64_t minStrongBlurDelta = rawAdjacentDelta <= 1u ? 48u : 56u;

    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_popup_adjacent_rgb_delta", popupAdjacentDelta, S_OK);
    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_raw_adjacent_rgb_delta", rawAdjacentDelta, S_OK);
    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_popup_vs_raw_rgb_delta", popupVsRawDelta, S_OK);
    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_popup_vs_slab_rgb_delta", popupVsSlabDelta, S_OK);
    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_popup_to_raw_delta_permille", popupToRawDeltaPermille, S_OK);

    Require(rawAdjacentDelta > 0u, "acrylic backdrop metric scenario records non-zero raw backdrop variance");
    Require(popupVsRawDelta > 0u, "acrylic backdrop metric scenario materially changes the captured backdrop sample");
    Require(popupAdjacentDelta < rawAdjacentDelta, "acrylic backdrop metric scenario blurs the popup backdrop below the raw screen variance");
    Require(popupVsRawDelta >= minStrongBlurDelta,
            "acrylic backdrop metric scenario applies a visibly strong blur instead of a barely-changed transparent tint");
    Require(popupVsSlabDelta >= 24u,
            "acrylic backdrop metric scenario stays visually tied to the captured raw backdrop instead of collapsing into an opaque tint");

    Debug::Perf::Emit(L"dxui.menu.selftest.acrylic_open_to_capture_us",
                      L"",
                      openToCaptureUs,
                      static_cast<uint64_t>(popupCapture.widthPx),
                      static_cast<uint64_t>(popupCapture.heightPx),
                      S_OK);
}

void TestContextMenuShowAsyncKeepsOwnerPaintableWhileOpen()
{
    using namespace RedSalamander::DxUi;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Async One", .commandId = 7701},
        {.text = L"Async Two", .commandId = 7702},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 160, 160, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ownerWindow.Host().SetRoot(std::make_unique<StripedBackdropControl>(8));
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    InvalidateRect(ownerWindow.Hwnd(), nullptr, FALSE);
    UpdateWindow(ownerWindow.Hwnd());
    ownerWindow.PumpMessages();

    const POINT menuPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 72, 48, "async context menu anchor maps to screen coordinates");
    bool callbackInvoked  = false;
    std::optional<int> callbackResult;
    const bool shown = ContextMenu::ShowAsync(ownerWindow.Hwnd(),
                                              menuPoint,
                                              items,
                                              ownerWindow.Host().GetTheme(),
                                              [&](std::optional<int> commandId) noexcept
    {
        callbackInvoked = true;
        callbackResult  = commandId;
    });
    Require(shown, "async context menu show succeeds");
    Require(! callbackInvoked, "async context menu returns before any item is invoked");

    const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Async One");
    Require(popupHwnd != nullptr, "async context menu popup window appears");

    const uint64_t renderCountBefore = ownerWindow.Host().DebugGetRenderCount();
    InvalidateRect(ownerWindow.Hwnd(), nullptr, FALSE);
    UpdateWindow(ownerWindow.Hwnd());
    ownerWindow.PumpMessages();
    Require(ownerWindow.Host().DebugGetRenderCount() > renderCountBefore, "owner host repaints while async context menu is open");

    ContextMenuPopupDebugState popupState{};
    D2D1_RECT_F rowRectDip{};
    Require(WaitForContextMenuPopupState(popupHwnd,
                                         [](const ContextMenuPopupDebugState& state) noexcept
    { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                         popupState) &&
                WaitForContextMenuPopupItemRect(popupHwnd, 1u, rowRectDip),
            "async context menu exposes the second row geometry");

    POINT rowCenter{
        static_cast<LONG>(
            std::lround((rowRectDip.left + rowRectDip.right) * 0.5f * static_cast<float>(popupState.dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
        static_cast<LONG>(
            std::lround((rowRectDip.top + rowRectDip.bottom) * 0.5f * static_cast<float>(popupState.dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
    };
    Require(ClientToScreen(popupHwnd, &rowCenter) != FALSE, "async context menu row center maps to screen coordinates");

    static_cast<void>(SendCapturedMouseMessageForMenuSuite(popupHwnd, WM_LBUTTONDOWN, MK_LBUTTON, rowCenter));
    static_cast<void>(SendCapturedMouseMessageForMenuSuite(popupHwnd, WM_LBUTTONUP, 0, rowCenter));
    ownerWindow.PumpMessages();

    Require(callbackInvoked, "async context menu invokes callback after item click");
    Require(callbackResult == std::optional<int>{7702}, "async context menu callback receives invoked command");
    Require(WaitForWindowDestroyed(popupHwnd), "async context menu closes after item click");
}

void TestLargeMenuPaintsOnlyVisibleRowsWithCachedOffsets()
{
    using namespace RedSalamander::DxUi;

    constexpr size_t kItemCount = 4096u;
    std::vector<MenuFlyoutItem> items;
    items.reserve(kItemCount);
    for (size_t i = 0; i < kItemCount; ++i)
    {
        items.push_back(MenuFlyoutItem{.text = std::format(L"Folder {:04}", i), .commandId = static_cast<int>(80'000u + i)});
    }

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 180, 180, 420, 300, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    ContextMenuSessionCallbacks callbacks{};
    callbacks.maxRootHeightDip        = 240.0f;
    callbacks.focusFirstNavigableItem = true;

    bool callbackInvoked = false;
    const POINT menuPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 64, 48, "large context menu anchor maps to screen coordinates");
    const auto openStarted = std::chrono::steady_clock::now();
    const bool shown = ContextMenu::ShowAsync(ownerWindow.Hwnd(),
                                              menuPoint,
                                              items,
                                              ownerWindow.Host().GetTheme(),
                                              [&](std::optional<int>) noexcept { callbackInvoked = true; },
                                              callbacks);
    const uint64_t openToFirstPaintUs = Debug::Perf::ElapsedUs(openStarted);
    Require(shown, "large async context menu opens");
    Require(openToFirstPaintUs < 5'000'000u, "large context menu open-to-first-paint remains bounded");

    const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Folder 0000");
    Require(popupHwnd != nullptr, "large context menu popup window appears");

    ContextMenuPopupDebugState initialState{};
    Require(WaitForContextMenuPopupState(popupHwnd,
                                         [](const ContextMenuPopupDebugState& state) noexcept
    { return state.renderCount > 0u && state.lastPaintedItemCount > 0u; },
                                         initialState),
            "large context menu completes its first visible-row paint");
    Require(initialState.hasScrollbar, "large context menu uses a bounded viewport");
    Require(initialState.itemTexts.size() == kItemCount, "large context menu retains the immutable command model");
    Require(initialState.lastPaintedItemCount <= 32u, "large context menu first paint touches only viewport rows");

    const auto endStarted = std::chrono::steady_clock::now();
    SendMessageW(popupHwnd, WM_KEYDOWN, VK_END, 0);
    ownerWindow.PumpMessages();
    ContextMenuPopupDebugState endState{};
    Require(WaitForContextMenuPopupState(popupHwnd,
                                         [](const ContextMenuPopupDebugState& state) noexcept
    { return state.keyboardIndex == std::optional<size_t>{kItemCount - 1u} && state.scrollOffsetDip > 0.0f; },
                                         endState),
            "large context menu resolves the last row through cached offsets");
    const uint64_t endToVisibleUs = Debug::Perf::ElapsedUs(endStarted);
    Require(endToVisibleUs < 1'000'000u, "large context menu End-to-visible latency remains bounded");
    Require(endState.lastPaintedItemCount <= 32u, "large context menu scrolled paint remains limited to viewport rows");

    D2D1_RECT_F lastRowRect{};
    Require(WaitForContextMenuPopupItemRect(popupHwnd, kItemCount - 1u, lastRowRect), "large context menu exposes last-row geometry");
    Require(lastRowRect.bottom > endState.viewportRectDip.top && lastRowRect.top < endState.viewportRectDip.bottom,
            "large context menu keeps the keyboard-selected last row inside the viewport");

    Debug::Perf::Emit(L"dxui.menu.selftest.large_open_to_first_paint_us",
                      L"4096-items",
                      openToFirstPaintUs,
                      static_cast<uint64_t>(kItemCount),
                      static_cast<uint64_t>(initialState.lastPaintedItemCount),
                      S_OK);
    Debug::Perf::Emit(L"dxui.menu.selftest.large_end_to_visible_us",
                      L"4096-items",
                      endToVisibleUs,
                      static_cast<uint64_t>(kItemCount),
                      static_cast<uint64_t>(endState.lastPaintedItemCount),
                      S_OK);

    PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
    const auto closeDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
    while (! callbackInvoked && std::chrono::steady_clock::now() < closeDeadline)
    {
        ownerWindow.PumpMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(5));
    }
    Require(callbackInvoked, "large context menu closes through Escape");
    Require(WaitForWindowDestroyed(popupHwnd), "large context menu popup is destroyed after Escape");
}

void TestContextMenuShowAsyncSubmenuCloseKeepsSessionAlive()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 140, 140, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();
    DrainPendingMouseMessagesForMenuSuite();

    const std::vector<MenuFlyoutItem> items = {
        {.text      = L"S1",
         .commandId = 7901,
         .children =
             {
                 MenuFlyoutItem{.text = L"S11", .commandId = 79011},
                 MenuFlyoutItem{.text = L"S12", .commandId = 79012},
             }},
        {.text = L"S2", .commandId = 7902},
    };

    bool callbackInvoked = false;
    std::optional<int> callbackResult;
    const POINT menuPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 60, 40, "async submenu-close menu anchor maps to screen coordinates");
    Require(ContextMenu::ShowAsync(ownerWindow.Hwnd(),
                                   menuPoint,
                                   items,
                                   ownerWindow.Host().GetTheme(),
                                   [&](std::optional<int> commandId) noexcept
    {
        callbackInvoked = true;
        callbackResult  = commandId;
    }),
            "async submenu-close context menu opens");

    const HWND rootPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"S1");
    Require(rootPopupHwnd != nullptr, "async submenu-close root popup window appears");

    const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

    ContextMenuPopupDebugState popupState{};
    D2D1_RECT_F firstItemRectDip  = D2D1::RectF();
    D2D1_RECT_F secondItemRectDip = D2D1::RectF();
    Require(WaitForContextMenuPopupState(rootPopupHwnd,
                                         [](const ContextMenuPopupDebugState& state) noexcept
    { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                         popupState) &&
                WaitForContextMenuPopupItemRect(rootPopupHwnd, 0u, firstItemRectDip) && WaitForContextMenuPopupItemRect(rootPopupHwnd, 1u, secondItemRectDip),
            "async submenu-close root popup exposes its geometry");

    const float scale = static_cast<float>(popupState.dpi) / 96.0f;
    const LONG s1X    = static_cast<LONG>(std::lround(((firstItemRectDip.left + firstItemRectDip.right) * 0.5f) * scale));
    const LONG s1Y    = static_cast<LONG>(std::lround(((firstItemRectDip.top + firstItemRectDip.bottom) * 0.5f) * scale));
    const LONG s2X    = static_cast<LONG>(std::lround(((secondItemRectDip.left + secondItemRectDip.right) * 0.5f) * scale));
    const LONG s2Y    = static_cast<LONG>(std::lround(((secondItemRectDip.top + secondItemRectDip.bottom) * 0.5f) * scale));

    // Hover the parent item until the deferred-open hover timer is pending, then
    // fire it so the submenu opens through the async WndProc timer path.
    Require(SendSettledClientMouseMoveForMenuSuite(rootPopupHwnd, s1X, s1Y) && WaitForContextMenuPopupState(rootPopupHwnd,
                                                                                                            [](const ContextMenuPopupDebugState& state) noexcept
    { return state.hoverTimerActive && state.hoverTimerPendingOpen; },
                                                                                                            popupState),
            "hovering the async parent item schedules the submenu open timer");
    Require(FirePendingSubmenuHoverTimerForMenuSuite(rootPopupHwnd), "async submenu open timer fires");
    ownerWindow.PumpMessages();
    const HWND submenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"S11");
    Require(submenuHwnd != nullptr, "async submenu opens after the hover timer fires");
    Require(! callbackInvoked, "async session stays open while the submenu is open");

    // Hover the sibling without children: this schedules the delayed close, and
    // firing it destroys the submenu window from inside the controller. That
    // controller-initiated destroy must not finalize the async session.
    Require(SendSettledClientMouseMoveForMenuSuite(rootPopupHwnd, s2X, s2Y) && WaitForContextMenuPopupState(rootPopupHwnd,
                                                                                                            [](const ContextMenuPopupDebugState& state) noexcept
    { return state.hoverTimerActive && state.hoverTimerPendingClose; },
                                                                                                            popupState),
            "hovering the async sibling item schedules the submenu close timer");
    Require(FirePendingSubmenuHoverTimerForMenuSuite(rootPopupHwnd), "async submenu close timer fires");
    ownerWindow.PumpMessages();

    Require(WaitForWindowDestroyed(submenuHwnd), "async submenu closes after the close timer fires");
    Require(IsWindow(rootPopupHwnd) != FALSE, "async root popup survives the controller-initiated submenu close");
    Require(! callbackInvoked, "async session does not finalize when its submenu closes");

    // The session is still interactive: clicking the sibling item delivers its command.
    const POINT siblingCenter = ClientScreenPointForTest(rootPopupHwnd, s2X, s2Y, "async sibling item maps to screen coordinates");
    static_cast<void>(SendCapturedMouseMessageForMenuSuite(rootPopupHwnd, WM_LBUTTONDOWN, MK_LBUTTON, siblingCenter));
    static_cast<void>(SendCapturedMouseMessageForMenuSuite(rootPopupHwnd, WM_LBUTTONUP, 0, siblingCenter));
    ownerWindow.PumpMessages();

    Require(callbackInvoked, "async context menu invokes the closed callback after the post-close click");
    Require(callbackResult == std::optional<int>{7902}, "async context menu delivers the sibling command after the submenu close");
    Require(WaitForWindowDestroyed(rootPopupHwnd), "async root popup closes after invoking the sibling command");
}

void TestMenuGraphicalSliderSupportsClickDragAndAnimation()
{
    using namespace RedSalamander::DxUi;

    const auto makeSliderItems = []
    {
        MenuFlyoutItem slider{
            .kind            = MenuItemKind::Slider,
            .text            = L"Thumbnail size",
            .acceleratorText = L"Medium",
            .sliderStops =
                {
                    {.text = L"Small", .commandId = 8201},
                    {.text = L"Medium", .commandId = 8202},
                    {.text = L"Large", .commandId = 8203},
                    {.text = L"Extra Large", .commandId = 8204},
                },
            .sliderValue = 1u,
        };
        return std::vector<MenuFlyoutItem>{std::move(slider)};
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 180, 180, 340, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const auto openSlider = [&](bool& callbackInvoked, std::optional<int>& callbackResult)
    {
        const POINT menuPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 80, 56, "graphical slider context menu anchor maps to screen coordinates");
        Require(ContextMenu::ShowAsync(ownerWindow.Hwnd(),
                                       menuPoint,
                                       makeSliderItems(),
                                       ownerWindow.Host().GetTheme(),
                                       [&](std::optional<int> commandId) noexcept
        {
            callbackInvoked = true;
            callbackResult  = commandId;
        }),
                "graphical slider context menu opens asynchronously");
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Thumbnail size");
        Require(popupHwnd != nullptr, "graphical slider popup window appears");
        return popupHwnd;
    };

    const auto sliderPoint = [](HWND popupHwnd, const ContextMenuPopupDebugState& state, const D2D1_RECT_F& rowRectDip, float stopPosition)
    {
        const float scale = static_cast<float>(state.dpi) / 96.0f;
        POINT point{
            static_cast<LONG>(std::lround(std::lerp(rowRectDip.left + 16.0f, rowRectDip.right - 16.0f, stopPosition / 3.0f) * scale)),
            static_cast<LONG>(std::lround((rowRectDip.top + 50.0f) * scale)),
        };
        Require(ClientToScreen(popupHwnd, &point) != FALSE, "graphical slider stop maps to screen coordinates");
        return point;
    };

    bool clickCallbackInvoked = false;
    std::optional<int> clickCallbackResult;
    const HWND clickPopup = openSlider(clickCallbackInvoked, clickCallbackResult);
    ContextMenuPopupDebugState clickState{};
    D2D1_RECT_F clickRowRectDip{};
    Require(DebugGetContextMenuPopupState(clickPopup, clickState) && DebugGetContextMenuPopupItemRect(clickPopup, 0u, clickRowRectDip),
            "graphical slider click exposes state and row geometry");
    Require(clickState.sliderStopVisualExtentsDip.size() == 1u && clickState.sliderStopVisualExtentsDip[0u].size() == 4u &&
                clickState.sliderStopVisualExtentsDip[0u][0u] < clickState.sliderStopVisualExtentsDip[0u][1u] &&
                clickState.sliderStopVisualExtentsDip[0u][1u] < clickState.sliderStopVisualExtentsDip[0u][2u] &&
                clickState.sliderStopVisualExtentsDip[0u][2u] < clickState.sliderStopVisualExtentsDip[0u][3u],
            "slider stops use progressively larger graphical thumbnail representations");

    const POINT largePoint = sliderPoint(clickPopup, clickState, clickRowRectDip, 2.0f);
    static_cast<void>(SendCapturedMouseMessageForMenuSuite(clickPopup, WM_LBUTTONDOWN, MK_LBUTTON, largePoint));
    Require(DebugGetContextMenuPopupState(clickPopup, clickState) && clickState.sliderDragging[0u] && clickState.sliderValues[0u] == 2u &&
                clickState.sliderTargetPositions[0u] == 2.0f && clickState.sliderAnimatedPositions[0u] < clickState.sliderTargetPositions[0u],
            "direct slider press previews the selected stop and starts an animated transition");
    static_cast<void>(SendCapturedMouseMessageForMenuSuite(clickPopup, WM_LBUTTONUP, 0, largePoint));
    ownerWindow.PumpMessages();
    Require(clickCallbackInvoked && clickCallbackResult == std::optional<int>{8203}, "direct slider click commits the selected stop on release");
    Require(WaitForWindowDestroyed(clickPopup), "direct slider click closes the popup after committing");

    bool dragCallbackInvoked = false;
    std::optional<int> dragCallbackResult;
    const HWND dragPopup = openSlider(dragCallbackInvoked, dragCallbackResult);
    ContextMenuPopupDebugState dragState{};
    D2D1_RECT_F dragRowRectDip{};
    Require(DebugGetContextMenuPopupState(dragPopup, dragState) && DebugGetContextMenuPopupItemRect(dragPopup, 0u, dragRowRectDip),
            "graphical slider drag exposes state and row geometry");
    const POINT smallPoint = sliderPoint(dragPopup, dragState, dragRowRectDip, 0.0f);
    const POINT extraLargePoint = sliderPoint(dragPopup, dragState, dragRowRectDip, 3.0f);
    static_cast<void>(SendCapturedMouseMessageForMenuSuite(dragPopup, WM_LBUTTONDOWN, MK_LBUTTON, smallPoint));
    static_cast<void>(SendCapturedMouseMessageForMenuSuite(dragPopup, WM_MOUSEMOVE, MK_LBUTTON, extraLargePoint));
    Require(DebugGetContextMenuPopupState(dragPopup, dragState) && dragState.sliderDragging[0u] && dragState.sliderValues[0u] == 3u &&
                dragState.sliderTargetPositions[0u] == 3.0f && dragState.sliderAnimatedPositions[0u] < dragState.sliderTargetPositions[0u] &&
                dragState.itemAcceleratorTexts[0u] == L"Extra Large",
            "held pointer drag updates the live label and animates toward the graphical endpoint");
    static_cast<void>(SendCapturedMouseMessageForMenuSuite(dragPopup, WM_LBUTTONUP, 0, extraLargePoint));
    ownerWindow.PumpMessages();
    Require(dragCallbackInvoked && dragCallbackResult == std::optional<int>{8204}, "slider drag commits the final stop on release");
    Require(WaitForWindowDestroyed(dragPopup), "slider drag closes the popup after committing");
}

void TestContextMenuPopupRelayoutsOnDpiChanged()
{
    using namespace RedSalamander::DxUi;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"DPI One", .acceleratorText = L"Ctrl+1", .commandId = 7801},
        {.text = L"DPI Two", .acceleratorText = L"Ctrl+2", .commandId = 7802},
        {.text = L"DPI Three", .acceleratorText = L"Ctrl+3", .commandId = 7803},
        {.text = L"DPI Four", .acceleratorText = L"Ctrl+4", .commandId = 7804},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 180, 180, 340, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    bool callbackInvoked = false;
    std::optional<int> callbackResult;
    const POINT menuPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 80, 56, "dpi context menu anchor maps to screen coordinates");
    Require(ContextMenu::ShowAsync(ownerWindow.Hwnd(),
                                   menuPoint,
                                   items,
                                   ownerWindow.Host().GetTheme(),
                                   [&](std::optional<int> commandId) noexcept
    {
        callbackInvoked = true;
        callbackResult  = commandId;
    }),
            "async context menu opens for DPI relayout validation");

    const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"DPI One");
    Require(popupHwnd != nullptr, "DPI relayout context menu popup window appears");

    const auto dismissPopup = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(popupHwnd) != FALSE)
        {
            SendMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
        }
    });

    ContextMenuPopupDebugState initialState{};
    Require(WaitForContextMenuPopupState(popupHwnd,
                                         [](const ContextMenuPopupDebugState& state) noexcept
    { return state.dpi > 0u && state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f && ! state.hasScrollbar; },
                                         initialState),
            "DPI relayout popup exposes initial non-scrolling geometry");

    const UINT sourceDpi            = initialState.dpi == 0u ? USER_DEFAULT_SCREEN_DPI : initialState.dpi;
    const UINT targetDpi            = sourceDpi == 144u ? USER_DEFAULT_SCREEN_DPI : 144u;
    const int initialWindowWidthPx  = initialState.windowRectPx.right - initialState.windowRectPx.left;
    const int initialWindowHeightPx = initialState.windowRectPx.bottom - initialState.windowRectPx.top;
    RECT suggestedWindowRect{
        initialState.windowRectPx.left,
        initialState.windowRectPx.top,
        initialState.windowRectPx.left + MulDiv(initialWindowWidthPx, static_cast<int>(targetDpi), static_cast<int>(sourceDpi)),
        initialState.windowRectPx.top + MulDiv(initialWindowHeightPx, static_cast<int>(targetDpi), static_cast<int>(sourceDpi)),
    };

    SendMessageW(
        popupHwnd, WM_DPICHANGED, MAKEWPARAM(static_cast<WORD>(targetDpi), static_cast<WORD>(targetDpi)), reinterpret_cast<LPARAM>(&suggestedWindowRect));

    ContextMenuPopupDebugState relayoutState{};
    Require(WaitForContextMenuPopupState(popupHwnd,
                                         [targetDpi](const ContextMenuPopupDebugState& state) noexcept
    { return state.dpi == targetDpi && state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f && ! state.hasScrollbar; },
                                         relayoutState),
            "DPI relayout popup updates its debug DPI and visible geometry");

    const int surfaceWidthPx  = relayoutState.surfaceRectPx.right - relayoutState.surfaceRectPx.left;
    const int surfaceHeightPx = relayoutState.surfaceRectPx.bottom - relayoutState.surfaceRectPx.top;
    Require(surfaceWidthPx == static_cast<int>(DipToPixelForPopup(relayoutState.visibleWidthDip, targetDpi)),
            "DPI relayout popup surface width matches the new DPI scale");
    Require(surfaceHeightPx == static_cast<int>(DipToPixelForPopup(relayoutState.visibleHeightDip, targetDpi)),
            "DPI relayout popup surface height matches the new DPI scale");

    D2D1_RECT_F lastItemRectDip{};
    Require(WaitForContextMenuPopupItemRect(popupHwnd, items.size() - 1u, lastItemRectDip), "DPI relayout popup exposes the last item rect");
    Require(lastItemRectDip.bottom <= relayoutState.viewportRectDip.bottom + 0.5f, "DPI relayout keeps the last visible menu item inside the popup viewport");

    WindowHostBitmapCapture capture{};
    Require(WaitForContextMenuPopupBitmapCapture(popupHwnd, capture), "DPI relayout popup remains renderable");
    Require(capture.widthPx == static_cast<UINT>(relayoutState.windowRectPx.right - relayoutState.windowRectPx.left) &&
                capture.heightPx == static_cast<UINT>(relayoutState.windowRectPx.bottom - relayoutState.windowRectPx.top),
            "DPI relayout popup capture dimensions match the resized window rectangle");

    SendMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
    ownerWindow.PumpMessages();
    Require(callbackInvoked, "DPI relayout async popup closes through Escape");
    Require(! callbackResult.has_value(), "DPI relayout async popup closes without invoking a command");
    Require(WaitForWindowDestroyed(popupHwnd), "DPI relayout async popup window is destroyed after Escape");
}

} // namespace

void RunMenuTests()
{
    auto runTest = [](const char* name, void (*fn)())
    {
        std::cerr << "  [START] " << name << '\n' << std::flush;
        fn();
        RedSalamander::Ui::AnimationDispatcher::GetInstance().Shutdown();
        std::cerr << "  [DONE] " << name << '\n' << std::flush;
    };

    runTest("TestFolderViewIncrementalSearchKeepsContainsHighlightButUsesPrefixFocus", TestFolderViewIncrementalSearchKeepsContainsHighlightButUsesPrefixFocus);
    runTest("TestFolderViewInactiveVisualStateDimsNormalTextAndIcons", TestFolderViewInactiveVisualStateDimsNormalTextAndIcons);
    runTest("TestFolderViewEmptyPlaceholderMetricsUseCurrentEmptyLayout", TestFolderViewEmptyPlaceholderMetricsUseCurrentEmptyLayout);
    runTest("TestPointerInputEventMouseMoveUsesDeliveredPoint", TestPointerInputEventMouseMoveUsesDeliveredPoint);
    runTest("TestPointerInputEventButtonUsesDeliveredPointAndFlags", TestPointerInputEventButtonUsesDeliveredPointAndFlags);
    runTest("TestPointerInputEventWheelUsesDeliveredScreenPoint", TestPointerInputEventWheelUsesDeliveredScreenPoint);
    runTest("TestPointerInputEventHasNoLiveCursorState", TestPointerInputEventHasNoLiveCursorState);
    runTest("TestPointerInputOriginAbstractionIsRemoved", TestPointerInputOriginAbstractionIsRemoved);
    runTest("TestNavigationViewPointerRoutingHasNoSyntheticGenerationGate", TestNavigationViewPointerRoutingHasNoSyntheticGenerationGate);
    runTest("TestMenuWindowClassRegistrationCachesOnlySuccess", TestMenuWindowClassRegistrationCachesOnlySuccess);
    runTest("TestMenuPopupWindowRegionTransfersOwnershipOnlyAfterSuccess", TestMenuPopupWindowRegionTransfersOwnershipOnlyAfterSuccess);
    runTest("TestContextMenuModalLoopDismissesWhenRootPopupDisappearsBeforeWaiting", TestContextMenuModalLoopDismissesWhenRootPopupDisappearsBeforeWaiting);
    runTest("TestContextMenuDebugStateProbeBoundsWedgedWindowThread", TestContextMenuDebugStateProbeBoundsWedgedWindowThread);
    runTest("TestContextMenuShowAsyncKeepsOwnerPaintableWhileOpen", TestContextMenuShowAsyncKeepsOwnerPaintableWhileOpen);
    runTest("TestLargeMenuPaintsOnlyVisibleRowsWithCachedOffsets", TestLargeMenuPaintsOnlyVisibleRowsWithCachedOffsets);
    runTest("TestContextMenuShowAsyncSubmenuCloseKeepsSessionAlive", TestContextMenuShowAsyncSubmenuCloseKeepsSessionAlive);
    runTest("TestMenuGraphicalSliderSupportsClickDragAndAnimation", TestMenuGraphicalSliderSupportsClickDragAndAnimation);
    runTest("TestMenuMnemonicHonorsExplicitAmpersandLabels", TestMenuMnemonicHonorsExplicitAmpersandLabels);
    runTest("TestMenuOpeningPointerUpCanBeIgnoredOutsideVisibleSurface", TestMenuOpeningPointerUpCanBeIgnoredOutsideVisibleSurface);
    runTest("TestEmbeddedViewerContextMenuNativeConversionFiltersStandaloneCommands", TestEmbeddedViewerContextMenuNativeConversionFiltersStandaloneCommands);
    runTest("TestMenuShadowMarginMouseUpLightDismissesAfterInitialRelease", TestMenuShadowMarginMouseUpLightDismissesAfterInitialRelease);
    runTest("TestSplitButtonContextMenuSentMouseMessagesHoverAndInvokeImmediately", TestSplitButtonContextMenuSentMouseMessagesHoverAndInvokeImmediately);
    runTest("TestSplitButtonContextMenuOwnerMessageFloodDoesNotStarvePointerInput", TestSplitButtonContextMenuOwnerMessageFloodDoesNotStarvePointerInput);
    runTest("TestMenuKeyboardNavigationSkipsInfoRows", TestMenuKeyboardNavigationSkipsInfoRows);
    runTest("TestMenuKeyboardRightArrowMatchesWindowsMenuLoop", TestMenuKeyboardRightArrowMatchesWindowsMenuLoop);
    runTest("TestStationaryMouseDoesNotOverrideKeyboardRootSwitch", TestStationaryMouseDoesNotOverrideKeyboardRootSwitch);
    runTest("TestMenuPointerOverSiblingRootSwitchesOpenMenu", TestMenuPointerOverSiblingRootSwitchesOpenMenu);
    runTest("TestMenuPopupMouseMoveUsesDeliveredPointForRootSwitch", TestMenuPopupMouseMoveUsesDeliveredPointForRootSwitch);
    runTest("TestMenuOwnerMouseMoveRoutesRootSwitchImmediately", TestMenuOwnerMouseMoveRoutesRootSwitchImmediately);
    runTest("TestMenuBarHoverMessageSwitchesRootWhenCursorOutsidePopup", TestMenuBarHoverMessageSwitchesRootWhenCursorOutsidePopup);
    runTest("TestMenuBarHoverMessageSwitchesRootWhilePopupOverlapsMenuBar", TestMenuBarHoverMessageSwitchesRootWhilePopupOverlapsMenuBar);
    runTest("TestMenuRootSwitchUsesDeliveredOwnerMouseMoveAfterPopupSwitch", TestMenuRootSwitchUsesDeliveredOwnerMouseMoveAfterPopupSwitch);
    runTest("TestMenuRootSwitchDoesNotPollCursorWhileIdle", TestMenuRootSwitchDoesNotPollCursorWhileIdle);
    runTest("TestMenuHoveringSiblingClosesOpenSubmenuAfterDelay", TestMenuHoveringSiblingClosesOpenSubmenuAfterDelay);
    runTest("TestMenuHoveringSiblingWithChildrenReplacesOpenSubmenuAfterDelay", TestMenuHoveringSiblingWithChildrenReplacesOpenSubmenuAfterDelay);
    runTest("TestMenuPointerInsideSubmenuAndParentItemCancelPendingCloseDelay", TestMenuPointerInsideSubmenuAndParentItemCancelPendingCloseDelay);
    runTest("TestMenuKeyboardTabExitsMenuLoop", TestMenuKeyboardTabExitsMenuLoop);
    runTest("TestMenuKeyboardF10ExitsMenuLoop", TestMenuKeyboardF10ExitsMenuLoop);
    runTest("TestMenuKeyboardAltExitsMenuLoop", TestMenuKeyboardAltExitsMenuLoop);
    runTest("TestMenuKeyboardLeftArrowMatchesWindowsMenuLoop", TestMenuKeyboardLeftArrowMatchesWindowsMenuLoop);
    runTest("TestNativeMenuBarRestoresFocusAfterMenuDismiss", TestNativeMenuBarRestoresFocusAfterMenuDismiss);
    runTest("TestNativeMenuBarNestedPopupCanDestroyHostSafely", TestNativeMenuBarNestedPopupCanDestroyHostSafely);
    runTest("TestMenuInfoRowsDoNotDismissOnClick", TestMenuInfoRowsDoNotDismissOnClick);
    runTest("TestMenuPopupPositionClampsAcrossDpiMatrix", TestMenuPopupPositionClampsAcrossDpiMatrix);
    runTest("TestMenuPopupPositionSupportsRightAlignedRootAnchors", TestMenuPopupPositionSupportsRightAlignedRootAnchors);
    runTest("TestMenuPopupPositionSupportsAboveRightAlignedRootAnchors", TestMenuPopupPositionSupportsAboveRightAlignedRootAnchors);
    runTest("TestMenuInfoRowsUseMeasuredValueColumnWidth", TestMenuInfoRowsUseMeasuredValueColumnWidth);
    runTest("TestMenuStandardRowsDeriveAndAlignShortcutColumnFromTabbedText", TestMenuStandardRowsDeriveAndAlignShortcutColumnFromTabbedText);
    runTest("TestMenuShortcutRowsReserveChevronLaneWhenAnySubmenuExists", TestMenuShortcutRowsReserveChevronLaneWhenAnySubmenuExists);
    runTest("TestMenuBitmapIconsReachPopupLayout", TestMenuBitmapIconsReachPopupLayout);
    runTest("TestMenuPopupMaterialsProduceDistinctCaptures", TestMenuPopupMaterialsProduceDistinctCaptures);
    runTest("TestMenuRainbowHoverUsesSeededHighlightContrast", TestMenuRainbowHoverUsesSeededHighlightContrast);
    runTest("TestMenuHoverContrastAppliesToGlyphsAcrossThemes", TestMenuHoverContrastAppliesToGlyphsAcrossThemes);
    runTest("TestMenuRainbowCheckedItemUsesAccentIndicator", TestMenuRainbowCheckedItemUsesAccentIndicator);
    runTest("TestMenuCheckedRowsDoNotPaintSecondFullRowSelection", TestMenuCheckedRowsDoNotPaintSecondFullRowSelection);
    runTest("TestMenuCheckedRowsDoNotPaintLeadingCheckedBox", TestMenuCheckedRowsDoNotPaintLeadingCheckedBox);
    runTest("TestMenuPopupCompositionHostUsesTransparentShadowMargins", TestMenuPopupCompositionHostUsesTransparentShadowMargins);
    runTest("TestMenuPopupWindowClassDoesNotUseNativeDropShadow", TestMenuPopupWindowClassDoesNotUseNativeDropShadow);
    runTest("TestMenuPopupSubmenuFlipsLeftNearRightEdge", TestMenuPopupSubmenuFlipsLeftNearRightEdge);
    runTest("TestMenuPopupAcrylicLightVisualBaseline", TestMenuPopupAcrylicLightVisualBaseline);
    runTest("TestMenuPopupKeepsSystemBackdropDisabledForAppRenderedMaterials", TestMenuPopupKeepsSystemBackdropDisabledForAppRenderedMaterials);
    runTest("TestMenuAcrylicBackdropScenarioEmitsMetrics", TestMenuAcrylicBackdropScenarioEmitsMetrics);
    runTest("TestSplitButtonContextMenuSentMouseMessagesHoverAndOutsideDismiss", TestSplitButtonContextMenuSentMouseMessagesHoverAndOutsideDismiss);
    runTest("TestContextMenuPopupRelayoutsOnDpiChanged", TestContextMenuPopupRelayoutsOnDpiChanged);
}
