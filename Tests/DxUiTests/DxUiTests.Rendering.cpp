#include "DxUiTestHelpers.h"

namespace
{

using RedSalamander::DxUi::WindowHostBitmapCapture;

WindowHostBitmapCapture CaptureAttachedHostWindowBitmap(AttachedHostWindow& window, const char* context)
{
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

    WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), context);
    return capture;
}

struct CoreControlScene
{
    CoreControlScene()                                   = default;
    CoreControlScene(const CoreControlScene&)            = delete;
    CoreControlScene(CoreControlScene&&)                 = default;
    CoreControlScene& operator=(const CoreControlScene&) = delete;
    CoreControlScene& operator=(CoreControlScene&&)      = default;

    std::unique_ptr<RedSalamander::DxUi::Panel> root;
    ExposedButton* hoverButton                   = nullptr;
    ExposedButton* pressedButton                 = nullptr;
    ExposedButton* focusButton                   = nullptr;
    RedSalamander::DxUi::ProgressBar* marqueeBar = nullptr;
};

CoreControlScene BuildCoreControlScene()
{
    using namespace RedSalamander::DxUi;

    CoreControlScene scene{};
    scene.root = std::make_unique<Panel>();

    auto* restButton     = scene.root->AddChild<ExposedButton>(L"Rest");
    scene.hoverButton    = scene.root->AddChild<ExposedButton>(L"Hover");
    scene.pressedButton  = scene.root->AddChild<ExposedButton>(L"Pressed");
    scene.focusButton    = scene.root->AddChild<ExposedButton>(L"Primary");
    auto* disabledButton = scene.root->AddChild<ExposedButton>(L"Disabled");
    auto* toggle         = scene.root->AddChild<Toggle>(L"Wi-Fi");
    auto* checkbox       = scene.root->AddChild<Checkbox>(L"Updates");
    auto* radio          = scene.root->AddChild<RadioButton>(L"Daily");
    auto* field          = scene.root->AddChild<TextField>(L"");
    auto* progress       = scene.root->AddChild<ProgressBar>();
    scene.marqueeBar     = scene.root->AddChild<ProgressBar>();

    restButton->SetBounds(D2D1::RectF(16.0f, 16.0f, 112.0f, 48.0f));
    scene.hoverButton->SetBounds(D2D1::RectF(124.0f, 16.0f, 220.0f, 48.0f));
    scene.pressedButton->SetBounds(D2D1::RectF(232.0f, 16.0f, 328.0f, 48.0f));
    scene.focusButton->SetBounds(D2D1::RectF(340.0f, 16.0f, 456.0f, 48.0f));
    scene.focusButton->SetPrimary(true);
    disabledButton->SetBounds(D2D1::RectF(468.0f, 16.0f, 580.0f, 48.0f));
    disabledButton->SetEnabled(false);

    toggle->SetBounds(D2D1::RectF(16.0f, 70.0f, 210.0f, 102.0f));
    toggle->SetChecked(true);
    checkbox->SetBounds(D2D1::RectF(16.0f, 110.0f, 210.0f, 142.0f));
    checkbox->SetIndeterminate(true);
    radio->SetBounds(D2D1::RectF(16.0f, 150.0f, 210.0f, 182.0f));
    radio->SetChecked(true);

    field->SetBounds(D2D1::RectF(232.0f, 70.0f, 580.0f, 102.0f));
    field->SetPlaceholder(L"Search");
    progress->SetBounds(D2D1::RectF(232.0f, 116.0f, 580.0f, 138.0f));
    progress->SetValue(0.62);
    scene.marqueeBar->SetBounds(D2D1::RectF(232.0f, 150.0f, 580.0f, 172.0f));
    scene.marqueeBar->SetIndeterminate(true);

    return scene;
}

std::unique_ptr<RedSalamander::DxUi::Panel> BuildAnimatedPage(std::wstring title, const D2D1_RECT_F& heroBounds, std::wstring_view heroKey)
{
    using namespace RedSalamander::DxUi;

    auto page    = std::make_unique<Panel>();
    auto* label  = page->AddChild<Label>(std::move(title));
    auto* button = page->AddChild<Button>(L"Open");
    label->SetBounds(D2D1::RectF(12.0f, 12.0f, 240.0f, 40.0f));
    button->SetBounds(heroBounds);
    button->SetPrimary(true);
    button->SetConnectedAnimationKey(std::wstring(heroKey));
    return page;
}

void TestDxUiCoreControlsDarkVisualBaseline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    ThemePalette theme  = MakeDefaultThemePalette(true);
    theme.reducedMotion = true;
    window.Host().SetTheme(theme);

    CoreControlScene scene = BuildCoreControlScene();
    window.Host().SetRoot(std::move(scene.root));
    scene.hoverButton->OnHoverChanged(window.Host(), true);
    scene.pressedButton->OnHoverChanged(window.Host(), true);
    scene.pressedButton->SetPressed(true);
    window.Host().SetFocusControl(scene.focusButton);
    static_cast<void>(scene.marqueeBar->Tick(window.Host(), 250u));

    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmap(window, "dark core controls visual baseline capture succeeds");
    VerifyOrUpdateBaselineForTest("dark core controls visual baseline matches", L"core_controls_dark.png", capture);
}

void TestDxUiCoreControlsLightVisualBaseline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    ThemePalette theme  = MakeDefaultThemePalette(false);
    theme.reducedMotion = true;
    window.Host().SetTheme(theme);

    CoreControlScene scene = BuildCoreControlScene();
    window.Host().SetRoot(std::move(scene.root));
    scene.hoverButton->OnHoverChanged(window.Host(), true);
    scene.pressedButton->OnHoverChanged(window.Host(), true);
    scene.pressedButton->SetPressed(true);
    window.Host().SetFocusControl(scene.focusButton);
    static_cast<void>(scene.marqueeBar->Tick(window.Host(), 250u));

    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmap(window, "light core controls visual baseline capture succeeds");
    VerifyOrUpdateBaselineForTest("light core controls visual baseline matches", L"core_controls_light.png", capture);
}

void TestDxUiHighContrastVisualBaseline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    ThemePalette theme  = MakeDefaultThemePalette(false);
    theme.highContrast  = true;
    theme.reducedMotion = true;
    window.Host().SetTheme(theme);

    CoreControlScene scene = BuildCoreControlScene();
    window.Host().SetRoot(std::move(scene.root));
    scene.hoverButton->OnHoverChanged(window.Host(), true);
    scene.pressedButton->OnHoverChanged(window.Host(), true);
    scene.pressedButton->SetPressed(true);
    window.Host().SetFocusControl(scene.focusButton);

    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmap(window, "high-contrast visual baseline capture succeeds");
    VerifyOrUpdateBaselineForTest("high-contrast visual baseline matches", L"core_controls_high_contrast.png", capture);
}

void TestDxUiPopupAndBarsVisualBaseline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTheme(MakeDefaultThemePalette(true));

    auto root     = std::make_unique<Panel>();
    auto* menuBar = root->AddChild<MenuBar>();
    auto* toolbar = root->AddChild<Toolbar>();
    auto* combo   = root->AddChild<ComboBox>();
    auto* strip   = root->AddChild<StatusStrip>(L"Ready");

    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 32.0f));
    menuBar->SetItems({
        MenuBarItem{.text = L"File", .mnemonic = L'F', .enabled = true},
        MenuBarItem{.text = L"Edit", .mnemonic = L'E', .enabled = true},
        MenuBarItem{.text = L"View", .mnemonic = L'V', .enabled = true},
    });

    toolbar->SetBounds(D2D1::RectF(12.0f, 44.0f, 252.0f, 84.0f));
    toolbar->AddButton(L"Refresh", L"\xE72C");
    toolbar->AddButton(L"Copy", L"\xE8C8");
    toolbar->AddButton(L"Delete", L"\xE74D");

    combo->SetBounds(D2D1::RectF(280.0f, 48.0f, 520.0f, 80.0f));
    combo->SetItems({
        ComboBox::Item{L"one", L"One"},
        ComboBox::Item{L"two", L"Two"},
        ComboBox::Item{L"three", L"Three"},
        ComboBox::Item{L"four", L"Four"},
    });
    combo->SetSelectedIndex(1u);

    strip->SetBounds(D2D1::RectF(0.0f, 178.0f, 640.0f, 200.0f));
    strip->SetSections({
        StatusStrip::Section{.text = L"Ready", .widthDip = 0.0f},
        StatusStrip::Section{.text = L"UTF-8", .widthDip = 80.0f},
        StatusStrip::Section{.text = L"Ln 8, Col 42", .widthDip = 120.0f},
    });

    window.Host().SetRoot(std::move(root));

    Require(combo->OnMouseDown(window.Host(), D2D1::Point2F(500.0f, 64.0f), false, 0), "popup and bars baseline opens the combo popup");
    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmap(window, "popup and bars visual baseline capture succeeds");
    VerifyOrUpdateBaselineForTest("popup and bars visual baseline matches", L"popup_and_bars_dark.png", capture);
}

void TestDxUiPopupAndBarsAcrylicLightVisualBaseline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    ThemePalette theme    = MakeDefaultThemePalette(false);
    theme.overlayMaterial = OverlayMaterial::Acrylic;
    window.Host().SetTheme(theme);

    auto root     = std::make_unique<Panel>();
    auto* menuBar = root->AddChild<MenuBar>();
    auto* toolbar = root->AddChild<Toolbar>();
    auto* combo   = root->AddChild<ComboBox>();
    auto* strip   = root->AddChild<StatusStrip>(L"Ready");

    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 32.0f));
    menuBar->SetItems({
        MenuBarItem{.text = L"File", .mnemonic = L'F', .enabled = true},
        MenuBarItem{.text = L"Edit", .mnemonic = L'E', .enabled = true},
        MenuBarItem{.text = L"View", .mnemonic = L'V', .enabled = true},
    });

    toolbar->SetBounds(D2D1::RectF(12.0f, 44.0f, 252.0f, 84.0f));
    toolbar->AddButton(L"Refresh", L"\xE72C");
    toolbar->AddButton(L"Copy", L"\xE8C8");
    toolbar->AddButton(L"Delete", L"\xE74D");

    combo->SetBounds(D2D1::RectF(280.0f, 48.0f, 520.0f, 80.0f));
    combo->SetItems({
        ComboBox::Item{L"system", L"System"},
        ComboBox::Item{L"on", L"On"},
        ComboBox::Item{L"off", L"Off"},
    });
    combo->SetSelectedIndex(1u);

    strip->SetBounds(D2D1::RectF(0.0f, 178.0f, 640.0f, 200.0f));
    strip->SetSections({
        StatusStrip::Section{.text = L"Ready", .widthDip = 0.0f},
        StatusStrip::Section{.text = L"UTF-8", .widthDip = 80.0f},
        StatusStrip::Section{.text = L"Ln 8, Col 42", .widthDip = 120.0f},
    });

    window.Host().SetRoot(std::move(root));

    Require(combo->OnMouseDown(window.Host(), D2D1::Point2F(500.0f, 64.0f), false, 0), "acrylic popup and bars baseline opens the combo popup");
    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmap(window, "acrylic popup and bars visual baseline capture succeeds");
    VerifyOrUpdateBaselineForTest("acrylic popup and bars visual baseline matches", L"popup_and_bars_acrylic_light.png", capture);
}

void TestDxUiPageTransitionVisualBaseline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTheme(MakeDefaultThemePalette(true));

    auto root      = std::make_unique<Panel>();
    auto* pageHost = root->AddChild<PageHost>();
    pageHost->SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 200.0f));
    window.Host().SetRoot(std::move(root));

    pageHost->SetPage(BuildAnimatedPage(L"Overview", D2D1::RectF(24.0f, 48.0f, 160.0f, 92.0f), L"hero"));
    pageHost->SetPage(BuildAnimatedPage(L"Details", D2D1::RectF(420.0f, 112.0f, 592.0f, 156.0f), L"hero"), L"hero");
    pageHost->DebugFreezeTransitionProgress(0.5f);

    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmap(window, "page transition visual baseline capture succeeds");
    VerifyOrUpdateBaselineForTest("page transition visual baseline matches", L"page_transition_dark.png", capture);
}

void TestDxUiAdvancedControlsVisualBaseline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTheme(MakeDefaultThemePalette(true));
    SetWindowPos(window.Hwnd(), nullptr, 0, 0, 820, 280, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    window.PumpMessages();

    auto root        = std::make_unique<Panel>();
    auto* tabControl = root->AddChild<TabControl>();
    tabControl->SetBounds(D2D1::RectF(16.0f, 16.0f, 392.0f, 220.0f));
    tabControl->AddTab<Label>(L"Overview", L"Overview page");
    tabControl->AddTab<Label>(L"Details", L"Details page");
    tabControl->AddTab<Label>(L"Activity", L"Activity page");
    tabControl->AddTab<Label>(L"Permissions", L"Permissions page");
    tabControl->AddTab<Label>(L"Diagnostics", L"Diagnostics page");
    tabControl->AddTab<Label>(L"Advanced", L"Advanced page");
    tabControl->SetTabClosable(4u, true);
    tabControl->SetSelectedIndex(3u);

    auto* horizontalSlider = root->AddChild<Slider>();
    horizontalSlider->SetBounds(D2D1::RectF(440.0f, 40.0f, 760.0f, 72.0f));
    horizontalSlider->SetValue(68.0);
    horizontalSlider->SetTickMarks({0.0, 25.0, 50.0, 75.0, 100.0});

    auto* rtlSlider = root->AddChild<Slider>();
    rtlSlider->SetBounds(D2D1::RectF(440.0f, 108.0f, 760.0f, 140.0f));
    rtlSlider->SetFlowDirection(FlowDirection::RightToLeft);
    rtlSlider->SetValue(28.0);
    rtlSlider->SetTickMarks({0.0, 25.0, 50.0, 75.0, 100.0});

    auto* verticalSlider = root->AddChild<Slider>();
    verticalSlider->SetOrientation(SliderOrientation::Vertical);
    verticalSlider->SetBounds(D2D1::RectF(780.0f, 28.0f, 812.0f, 220.0f));
    verticalSlider->SetValue(74.0);

    auto* rtlStack = root->AddChild<StackPanel>();
    rtlStack->SetBounds(D2D1::RectF(440.0f, 172.0f, 760.0f, 212.0f));
    rtlStack->SetOrientation(StackOrientation::Horizontal);
    rtlStack->SetGap(8.0f);
    rtlStack->SetFlowDirection(FlowDirection::RightToLeft);
    auto* primaryButton = rtlStack->AddChild<Button>(L"Primary");
    primaryButton->SetPrimary(true);
    auto* secondaryButton = rtlStack->AddChild<Button>(L"Secondary");
    rtlStack->SetChildExtent(primaryButton, 112.0f);
    rtlStack->SetChildExtent(secondaryButton, 112.0f);
    rtlStack->ApplyLayout();

    window.Host().SetRoot(std::move(root));

    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmap(window, "advanced controls visual baseline capture succeeds");
    VerifyOrUpdateBaselineForTest("advanced controls visual baseline matches", L"advanced_controls_dark.png", capture);
}

void TestAttachedComboBoxPopupHoverDoesNotRepaintWhenHoveredItemStaysTheSame()
{
    using namespace RedSalamander::DxUi;

    std::cerr << "    [TRACE] rendering combo hover: begin\n" << std::flush;
    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(12.0f, 12.0f, 192.0f, 40.0f));
    combo->SetItems({
        ComboBox::Item{L"one", L"One"},
        ComboBox::Item{L"two", L"Two"},
        ComboBox::Item{L"three", L"Three"},
        ComboBox::Item{L"four", L"Four"},
    });
    window.Host().SetRoot(std::move(root));
    std::cerr << "    [TRACE] rendering combo hover: root attached\n" << std::flush;

    std::cerr << "    [TRACE] rendering combo hover: before ShowWindow\n" << std::flush;
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    std::cerr << "    [TRACE] rendering combo hover: after ShowWindow\n" << std::flush;
    MSG showMsg{};
    while (PeekMessageW(&showMsg, nullptr, 0, 0, PM_REMOVE) != FALSE)
    {
        wchar_t className[128]{};
        const BOOL isWindow = showMsg.hwnd != nullptr ? IsWindow(showMsg.hwnd) : FALSE;
        if (isWindow != FALSE)
        {
            static_cast<void>(GetClassNameW(showMsg.hwnd, className, static_cast<int>(std::size(className))));
        }
        std::cerr << "    [TRACE] rendering combo hover: dispatching initial msg=" << showMsg.message << " hwnd=" << showMsg.hwnd
                  << " isWindow=" << static_cast<int>(isWindow) << " class=";
        if (isWindow != FALSE)
        {
            std::wcerr << className;
        }
        else
        {
            std::cerr << "<invalid>";
        }
        std::cerr << " wp=" << showMsg.wParam << " lp=" << showMsg.lParam << '\n' << std::flush;
        if (DispatchQueuedMessageForTest(showMsg))
        {
            std::cerr << "    [TRACE] rendering combo hover: dispatched initial msg=" << showMsg.message << " hwnd=" << showMsg.hwnd << '\n' << std::flush;
        }
        else
        {
            std::cerr << "    [TRACE] rendering combo hover: skipped stale initial msg=" << showMsg.message << " hwnd=" << showMsg.hwnd << '\n' << std::flush;
        }
    }
    std::cerr << "    [TRACE] rendering combo hover: window shown\n" << std::flush;

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONDOWN, 0, MAKELPARAM(172, 24), handled));
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONUP, 0, MAKELPARAM(172, 24), handled));
    window.PumpMessages();
    std::cerr << "    [TRACE] rendering combo hover: popup toggle handled\n" << std::flush;
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "attached combo popup opens before hover repaint stability test");
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();
    std::cerr << "    [TRACE] rendering combo hover: initial redraw complete\n" << std::flush;
    const D2D1_RECT_F firstPopupItemRect  = combo->DebugGetPopupItemRect(0u, &window.Host());
    const D2D1_RECT_F secondPopupItemRect = combo->DebugGetPopupItemRect(1u, &window.Host());
    RequireRectHasArea(firstPopupItemRect, "attached combo popup exposes the first popup item geometry");
    RequireRectHasArea(secondPopupItemRect, "attached combo popup exposes the second popup item geometry");
    std::cerr << "    [TRACE] rendering combo hover: popup geometry ready\n" << std::flush;
    const D2D1_POINT_2F firstHoverPoint           = D2D1::Point2F(firstPopupItemRect.left + ((firstPopupItemRect.right - firstPopupItemRect.left) * 0.35f),
                                                                  (firstPopupItemRect.top + firstPopupItemRect.bottom) * 0.5f);
    const D2D1_POINT_2F nextRowHoverPoint         = D2D1::Point2F(secondPopupItemRect.left + ((secondPopupItemRect.right - secondPopupItemRect.left) * 0.35f),
                                                                  (secondPopupItemRect.top + secondPopupItemRect.bottom) * 0.5f);
    const D2D1_POINT_2F secondRowRepeatHoverPoint = D2D1::Point2F(secondPopupItemRect.left + ((secondPopupItemRect.right - secondPopupItemRect.left) * 0.65f),
                                                                  (secondPopupItemRect.top + secondPopupItemRect.bottom) * 0.5f);

#ifdef _DEBUG
    const uint64_t initialInvalidateCount = window.Host().DebugGetInvalidateCount();
    const uint64_t initialRenderCount     = window.Host().DebugGetRenderCount();
#endif
    Require(combo->DebugGetHoveredPopupIndex().has_value(), "attached combo popup opens with an initial hovered popup item");
    const size_t initialHoveredPopupIndex = combo->DebugGetHoveredPopupIndex().value();

    Require(combo->OnMouseMove(window.Host(), firstHoverPoint, 0), "attached combo popup first hover is handled");
    UpdateWindow(window.Hwnd());
    window.PumpMessages();
    std::cerr << "    [TRACE] rendering combo hover: first hover complete\n" << std::flush;
    Require(combo->DebugGetHoveredPopupIndex().has_value(), "attached combo popup first hover targets a popup item");
    Require(combo->DebugGetHoveredPopupIndex().value() == initialHoveredPopupIndex,
            "attached combo popup first hover stays on the initially highlighted popup item");

#ifdef _DEBUG
    Require(window.Host().DebugGetInvalidateCount() == initialInvalidateCount,
            "attached combo popup hover does not invalidate when the pointer stays on the initial popup item");
    Require(window.Host().DebugGetRenderCount() == initialRenderCount,
            "attached combo popup hover does not repaint when the pointer stays on the initial popup item");
#endif

    Require(combo->OnMouseMove(window.Host(), nextRowHoverPoint, 0), "attached combo popup next-row hover is handled");
    UpdateWindow(window.Hwnd());
    window.PumpMessages();
    std::cerr << "    [TRACE] rendering combo hover: second-row hover complete\n" << std::flush;
    Require(combo->DebugGetHoveredPopupIndex().has_value() && combo->DebugGetHoveredPopupIndex().value() != initialHoveredPopupIndex,
            "attached combo popup hover moves to the second popup item when the pointer changes rows");

#ifdef _DEBUG
    const uint64_t invalidateCountAfterRowChange = window.Host().DebugGetInvalidateCount();
    Require(invalidateCountAfterRowChange > initialInvalidateCount, "attached combo popup hover invalidates when the hovered popup item changes rows");
    const uint64_t renderCountAfterRowChange = window.Host().DebugGetRenderCount();
    Require(renderCountAfterRowChange > initialRenderCount, "attached combo popup hover repaints when the hovered popup item changes rows");
#endif

    const size_t hoveredPopupIndexAfterRowChange = combo->DebugGetHoveredPopupIndex().value();
    Require(combo->OnMouseMove(window.Host(), secondRowRepeatHoverPoint, 0), "attached combo popup repeated next-row hover is handled");
    UpdateWindow(window.Hwnd());
    window.PumpMessages();
    std::cerr << "    [TRACE] rendering combo hover: repeat hover complete\n" << std::flush;
    Require(combo->DebugGetHoveredPopupIndex().has_value() && combo->DebugGetHoveredPopupIndex().value() == hoveredPopupIndexAfterRowChange,
            "attached combo popup repeated next-row hover stays on the same popup item");

#ifdef _DEBUG
    Require(window.Host().DebugGetInvalidateCount() == invalidateCountAfterRowChange,
            "attached combo popup hover does not invalidate when the pointer stays on the second popup item");
    Require(window.Host().DebugGetRenderCount() == renderCountAfterRowChange,
            "attached combo popup hover does not repaint when the pointer stays on the second popup item");
    Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached combo popup hover repaint stability does not cause DX host resize failures");
#endif
}

void TestAttachedGridPaintHandlesDegenerateScrollbarTracks()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root  = std::make_unique<Grid>();
    auto* grid = root.get();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 24.0f, 40.0f));

    LargeGridModel model(40u, 8u, 120.0f);
    grid->SetModel(&model);
    window.Host().SetRoot(std::move(root));

    const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
    Require(metrics.hasVerticalScrollbar, "tiny attached grid still reports a vertical scrollbar when rows overflow");
    Require(metrics.hasHorizontalScrollbar, "tiny attached grid still reports a horizontal scrollbar when columns overflow");

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

#ifdef _DEBUG
    const uint64_t initialRenderCount = window.Host().DebugGetRenderCount();
#endif
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetRenderCount() > initialRenderCount,
            "tiny attached grid paints successfully when scrollbar tracks are smaller than the minimum thumb size");
    Require(window.Host().DebugGetResizeFailureCount() == 0u, "tiny attached grid paint does not cause DX host resize failures");
#endif
}

void TestGridHidesBottomClippedTrailingRow()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 170.0f));
    grid.SetRowHeightDip(28.0f);

    MultiRowGridModel model(10u);
    grid.SetModel(&model);

    Require(grid.GetVisibleRowCount() == 4u, "grid hides a bottom-clipped trailing row instead of counting it as visible");
    Require(grid.GetVisibleRowAt(3u).has_value() && grid.GetVisibleRowAt(3u).value() == 3u, "grid keeps the last fully visible row addressable");
    Require(! grid.GetVisibleRowAt(4u).has_value(), "grid does not expose a partially clipped trailing row as visible");
    Require(! grid.GetVisibleRowRect(4u).has_value(), "grid does not report geometry for a bottom-clipped trailing row");
}

class LargeIconBadgeGridModel final : public RedSalamander::DxUi::IDxGridModel
{
public:
    LargeIconBadgeGridModel(size_t rowCount, size_t columnCount, float columnWidthDip)
        : _rowCount(rowCount),
          _columnCount(columnCount),
          _columnWidthDip(columnWidthDip)
    {
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rowCount;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columnCount;
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        RedSalamander::DxUi::GridColumnDesc column;
        column.id       = std::format(L"state-{}", columnIndex);
        column.title    = std::format(L"State {}", columnIndex);
        column.widthDip = _columnWidthDip;
        column.kind     = (columnIndex % 3u == 0u) ? RedSalamander::DxUi::GridColumnKind::StateImage : RedSalamander::DxUi::GridColumnKind::Text;
        return column;
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        outCell.kind      = RedSalamander::DxUi::GridCellKind::IconText;
        outCell.iconText  = (columnIndex % 2u == 0u) ? L"*" : L"!";
        outCell.text      = std::format(L"Plugin {}:{}", rowIndex, columnIndex);
        outCell.badgeText = (rowIndex % 2u == 0u) ? L"Beta" : L"Live";
        outCell.badgeTone = (columnIndex % 2u == 0u) ? RedSalamander::DxUi::AdornmentTone::Info : RedSalamander::DxUi::AdornmentTone::Warning;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId >= _rowCount)
        {
            return std::nullopt;
        }
        return static_cast<size_t>(rowId);
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex);
    }

private:
    size_t _rowCount      = 0u;
    size_t _columnCount   = 0u;
    float _columnWidthDip = 96.0f;
};

void TestAttachedLargeGridVisibleWorkStaysBoundedAfterScroll()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root  = std::make_unique<Grid>();
    auto* grid = root.get();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    grid->SetRowHeightDip(24.0f);

    LargeGridModel model(1'000'000u, 64u, 96.0f);
    grid->SetModel(&model);
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    for (int i = 0; i < 24; ++i)
    {
        Require(grid->OnMouseWheel(window.Host(), D2D1::Point2F(24.0f, 48.0f), -static_cast<float>(WHEEL_DELTA), 0),
                "attached large grid handles wheel scrolling");
    }

    const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
    Require(metrics.visibleRowCount == 4u, "attached large grid keeps visible row count bounded after scrolling");
    Require(metrics.visibleColumnCount == 4u, "attached large grid keeps visible column count bounded after scrolling");
    Require(metrics.visibleCellCount == 16u, "attached large grid keeps visible cell work bounded after scrolling");
    Require(metrics.hasVerticalScrollbar, "attached large grid still reports a vertical scrollbar after scrolling");
    Require(metrics.hasHorizontalScrollbar, "attached large grid still reports a horizontal scrollbar after scrolling");

#ifdef _DEBUG
    const uint64_t initialRenderCount = window.Host().DebugGetRenderCount();
#endif
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetRenderCount() > initialRenderCount, "attached large grid repaints after scrolling");
    Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached large grid scrolling and repainting does not cause DX host resize failures");
#endif
}

void TestAttachedLargeIconBadgeGridVisibleWorkStaysBoundedAfterScroll()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root  = std::make_unique<Grid>();
    auto* grid = root.get();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    grid->SetRowHeightDip(24.0f);

    LargeIconBadgeGridModel model(250'000u, 48u, 96.0f);
    grid->SetModel(&model);
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    for (int i = 0; i < 24; ++i)
    {
        Require(grid->OnMouseWheel(window.Host(), D2D1::Point2F(24.0f, 48.0f), -static_cast<float>(WHEEL_DELTA), 0),
                "attached large icon/badge grid handles wheel scrolling");
    }

    const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
    Require(metrics.visibleRowCount == 4u, "attached large icon/badge grid keeps visible row count bounded after scrolling");
    Require(metrics.visibleColumnCount == 4u, "attached large icon/badge grid keeps visible column count bounded after scrolling");
    Require(metrics.visibleCellCount == 16u, "attached large icon/badge grid keeps visible cell work bounded after scrolling");
    Require(metrics.visibleIconCellCount == metrics.visibleCellCount, "attached large icon/badge grid reports bounded visible icon work");
    Require(metrics.visibleBadgeCellCount == metrics.visibleCellCount, "attached large icon/badge grid reports bounded visible badge work");
    Require(metrics.hasVerticalScrollbar, "attached large icon/badge grid still reports a vertical scrollbar after scrolling");
    Require(metrics.hasHorizontalScrollbar, "attached large icon/badge grid still reports a horizontal scrollbar after scrolling");

#ifdef _DEBUG
    const uint64_t initialRenderCount = window.Host().DebugGetRenderCount();
#endif
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetRenderCount() > initialRenderCount, "attached large icon/badge grid repaints after scrolling");
    Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached large icon/badge grid scrolling and repainting does not cause DX host resize failures");
#endif
}

void TestAttachedGridBottomScrollKeepsFirstVisibleRowFlushWithHeader()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root  = std::make_unique<Grid>();
    auto* grid = root.get();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 170.0f));
    grid->SetRowHeightDip(28.0f);

    MultiRowGridModel model(10u);
    grid->SetModel(&model);
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    bool reachedScrollEdge = false;
    for (int stepIndex = 0; stepIndex < 8 && ! reachedScrollEdge; ++stepIndex)
    {
        reachedScrollEdge =
            ! grid->OnMouseWheel(window.Host(), D2D1::Point2F(24.0f, 48.0f), -static_cast<float>(WHEEL_DELTA), 0);
    }
    Require(reachedScrollEdge, "attached grid bottom-alignment test reaches the bottom scroll edge");

    const std::optional<D2D1_RECT_F> headerRect = grid->GetVisibleColumnHeaderRect(0u);
    Require(headerRect.has_value(), "attached grid bottom-alignment test exposes the visible header rect");

    const std::optional<size_t> firstVisibleRowIndex = grid->GetVisibleRowAt(0u);
    Require(firstVisibleRowIndex.has_value(), "attached grid bottom-alignment test exposes the first visible row");
    Require(firstVisibleRowIndex.value() == 6u, "attached grid bottom-alignment test snaps the trailing viewport to the next full row boundary");

    const std::optional<D2D1_RECT_F> firstVisibleRowRect = grid->GetVisibleRowRect(firstVisibleRowIndex.value());
    Require(firstVisibleRowRect.has_value(), "attached grid bottom-alignment test exposes the first visible row rect");
    RequireFloatNear(firstVisibleRowRect->top,
                     headerRect->bottom,
                     0.01f,
                     "attached grid bottom-alignment test keeps the first visible row flush with the header at bottom scroll");

    const std::optional<D2D1_RECT_F> lastRowRect = grid->GetVisibleRowRect(9u);
    Require(lastRowRect.has_value(), "attached grid bottom-alignment test keeps the trailing row visible");
    Require(lastRowRect->bottom <= grid->GetBounds().bottom + 0.01f, "attached grid bottom-alignment test keeps the trailing row fully visible");
}

void TestAttachedLargeGridLongRunScrollingStaysBoundedWithoutResizeChurn()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root  = std::make_unique<Grid>();
    auto* grid = root.get();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    grid->SetRowHeightDip(24.0f);

    LargeGridModel model(1'000'000u, 64u, 96.0f);
    grid->SetModel(&model);
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    const D2D1_POINT_2F wheelPoint = D2D1::Point2F(24.0f, 48.0f);
#ifdef _DEBUG
    const uint64_t initialResizeCount = window.Host().DebugGetResizeCount();
    uint64_t lastRenderCount          = window.Host().DebugGetRenderCount();
#endif

    for (int chunkIndex = 0; chunkIndex < 12; ++chunkIndex)
    {
        for (int stepIndex = 0; stepIndex < 24; ++stepIndex)
        {
            Require(grid->OnMouseWheel(window.Host(), wheelPoint, -static_cast<float>(WHEEL_DELTA), 0),
                    "attached large grid handles sustained repeated wheel scrolling");
        }

        const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
        Require(metrics.visibleRowCount == 4u, "attached large grid long-run scrolling keeps visible row count bounded");
        Require(metrics.visibleColumnCount == 4u, "attached large grid long-run scrolling keeps visible column count bounded");
        Require(metrics.visibleCellCount == 16u, "attached large grid long-run scrolling keeps visible cell work bounded");
        Require(metrics.hasVerticalScrollbar, "attached large grid long-run scrolling keeps the vertical scrollbar active");
        Require(metrics.hasHorizontalScrollbar, "attached large grid long-run scrolling keeps the horizontal scrollbar active");

        RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        window.PumpMessages();

#ifdef _DEBUG
        const uint64_t renderCount = window.Host().DebugGetRenderCount();
        Require(renderCount > lastRenderCount, "attached large grid long-run scrolling repaints after each sustained chunk");
        Require(window.Host().DebugGetResizeCount() == initialResizeCount, "attached large grid long-run scrolling does not churn swapchain resizes");
        Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached large grid long-run scrolling does not cause DX host resize failures");
        lastRenderCount = renderCount;
#endif
    }
}

void TestAttachedLargeGroupedGridVisibleWorkStaysBoundedAfterScroll()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root  = std::make_unique<Grid>();
    auto* grid = root.get();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid->SetRowHeightDip(24.0f);
    grid->SetHeaderHeightDip(32.0f);

    GroupedGridModel model(2'000u);
    std::vector<GroupedGridModel::Group> groups;
    groups.reserve(100u);
    for (size_t groupIndex = 0; groupIndex < 100u; ++groupIndex)
    {
        groups.push_back(GroupedGridModel::Group{
            .stableId      = static_cast<uint64_t>(groupIndex + 1u),
            .title         = std::format(L"Group {:03}", groupIndex),
            .startRowIndex = groupIndex * 20u,
            .rowCount      = 20u,
        });
    }
    model.SetGroups(std::move(groups));
    grid->SetModel(&model);
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    for (int i = 0; i < 24; ++i)
    {
        Require(grid->OnMouseWheel(window.Host(), D2D1::Point2F(24.0f, 48.0f), -static_cast<float>(WHEEL_DELTA), 0),
                "attached large grouped grid handles wheel scrolling");
    }

    const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
    Require(metrics.visibleRowCount > 0u && metrics.visibleRowCount <= 7u, "attached large grouped grid keeps visible row count bounded after scrolling");
    Require(metrics.visibleGroupHeaderCount <= 2u, "attached large grouped grid keeps visible group-header count bounded after scrolling");
    Require(metrics.visibleColumnCount == 1u, "attached large grouped grid keeps visible column count bounded after scrolling");
    Require(metrics.visibleCellCount == metrics.visibleRowCount,
            "attached large grouped grid keeps visible cell work proportional to visible rows after scrolling");
    Require(metrics.hasVerticalScrollbar, "attached large grouped grid still reports a vertical scrollbar after scrolling");
    Require(! metrics.hasHorizontalScrollbar, "attached large grouped grid avoids a horizontal scrollbar when the single visible column fits");

#ifdef _DEBUG
    const uint64_t initialRenderCount = window.Host().DebugGetRenderCount();
#endif
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetRenderCount() > initialRenderCount, "attached large grouped grid repaints after scrolling");
    Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached large grouped grid scrolling and repainting does not cause DX host resize failures");
#endif
}

void TestAttachedLargeCheckboxGridVisibleWorkStaysBoundedAfterScroll()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root  = std::make_unique<Grid>();
    auto* grid = root.get();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    grid->SetRowHeightDip(24.0f);

    LargeCheckboxGridModel model(100'000u, 1u);
    grid->SetModel(&model);
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    for (int i = 0; i < 24; ++i)
    {
        Require(grid->OnMouseWheel(window.Host(), D2D1::Point2F(24.0f, 48.0f), -static_cast<float>(WHEEL_DELTA), 0),
                "attached large checkbox grid handles wheel scrolling");
    }

    const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
    Require(metrics.visibleRowCount == 4u, "attached large checkbox grid keeps visible row count bounded after scrolling");
    Require(metrics.visibleColumnCount == 2u, "attached large checkbox grid keeps visible column count bounded after scrolling");
    Require(metrics.visibleCellCount == 8u, "attached large checkbox grid keeps visible cell work bounded after scrolling");
    Require(metrics.hasVerticalScrollbar, "attached large checkbox grid still reports a vertical scrollbar after scrolling");

#ifdef _DEBUG
    const uint64_t initialRenderCount = window.Host().DebugGetRenderCount();
#endif
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetRenderCount() > initialRenderCount, "attached large checkbox grid repaints after scrolling");
    Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached large checkbox grid scrolling and repainting does not cause DX host resize failures");
#endif
}

void TestAttachedLargeGroupedGridLongRunScrollingStaysBoundedWithoutResizeChurn()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root  = std::make_unique<Grid>();
    auto* grid = root.get();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid->SetRowHeightDip(24.0f);
    grid->SetHeaderHeightDip(32.0f);

    GroupedGridModel model(2'000u);
    std::vector<GroupedGridModel::Group> groups;
    groups.reserve(100u);
    for (size_t groupIndex = 0; groupIndex < 100u; ++groupIndex)
    {
        groups.push_back(GroupedGridModel::Group{
            .stableId      = static_cast<uint64_t>(groupIndex + 1u),
            .title         = std::format(L"Group {:03}", groupIndex),
            .startRowIndex = groupIndex * 20u,
            .rowCount      = 20u,
        });
    }
    model.SetGroups(std::move(groups));
    grid->SetModel(&model);
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    const D2D1_POINT_2F wheelPoint = D2D1::Point2F(24.0f, 48.0f);
#ifdef _DEBUG
    const uint64_t initialResizeCount = window.Host().DebugGetResizeCount();
    uint64_t lastRenderCount          = window.Host().DebugGetRenderCount();
#endif

    for (int chunkIndex = 0; chunkIndex < 12; ++chunkIndex)
    {
        for (int stepIndex = 0; stepIndex < 24; ++stepIndex)
        {
            Require(grid->OnMouseWheel(window.Host(), wheelPoint, -static_cast<float>(WHEEL_DELTA), 0),
                    "attached large grouped grid handles sustained repeated wheel scrolling");
        }

        const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
        Require(metrics.visibleRowCount > 0u && metrics.visibleRowCount <= 7u,
                "attached large grouped grid long-run scrolling keeps visible row count bounded");
        Require(metrics.visibleGroupHeaderCount <= 2u, "attached large grouped grid long-run scrolling keeps visible group-header count bounded");
        Require(metrics.visibleColumnCount == 1u, "attached large grouped grid long-run scrolling keeps visible column count bounded");
        Require(metrics.visibleCellCount == metrics.visibleRowCount,
                "attached large grouped grid long-run scrolling keeps visible cell work proportional to visible rows");
        Require(metrics.hasVerticalScrollbar, "attached large grouped grid long-run scrolling keeps the vertical scrollbar active");
        Require(! metrics.hasHorizontalScrollbar, "attached large grouped grid long-run scrolling avoids horizontal scrollbar churn");

        RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        window.PumpMessages();

#ifdef _DEBUG
        const uint64_t renderCount = window.Host().DebugGetRenderCount();
        Require(renderCount > lastRenderCount, "attached large grouped grid long-run scrolling repaints after each sustained chunk");
        Require(window.Host().DebugGetResizeCount() == initialResizeCount, "attached large grouped grid long-run scrolling does not churn swapchain resizes");
        Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached large grouped grid long-run scrolling does not cause DX host resize failures");
        lastRenderCount = renderCount;
#endif
    }
}

void TestAttachedLargeCheckboxGridLongRunScrollingStaysBoundedWithoutResizeChurn()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root  = std::make_unique<Grid>();
    auto* grid = root.get();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    grid->SetRowHeightDip(24.0f);

    LargeCheckboxGridModel model(100'000u, 1u);
    grid->SetModel(&model);
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    const D2D1_POINT_2F wheelPoint = D2D1::Point2F(24.0f, 48.0f);
#ifdef _DEBUG
    const uint64_t initialResizeCount = window.Host().DebugGetResizeCount();
    uint64_t lastRenderCount          = window.Host().DebugGetRenderCount();
#endif

    for (int chunkIndex = 0; chunkIndex < 12; ++chunkIndex)
    {
        for (int stepIndex = 0; stepIndex < 24; ++stepIndex)
        {
            Require(grid->OnMouseWheel(window.Host(), wheelPoint, -static_cast<float>(WHEEL_DELTA), 0),
                    "attached large checkbox grid handles sustained repeated wheel scrolling");
        }

        const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
        Require(metrics.visibleRowCount == 4u, "attached large checkbox grid long-run scrolling keeps visible row count bounded");
        Require(metrics.visibleColumnCount == 2u, "attached large checkbox grid long-run scrolling keeps visible column count bounded");
        Require(metrics.visibleCellCount == 8u, "attached large checkbox grid long-run scrolling keeps visible cell work bounded");
        Require(metrics.hasVerticalScrollbar, "attached large checkbox grid long-run scrolling keeps the vertical scrollbar active");

        RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        window.PumpMessages();

#ifdef _DEBUG
        const uint64_t renderCount = window.Host().DebugGetRenderCount();
        Require(renderCount > lastRenderCount, "attached large checkbox grid long-run scrolling repaints after each sustained chunk");
        Require(window.Host().DebugGetResizeCount() == initialResizeCount, "attached large checkbox grid long-run scrolling does not churn swapchain resizes");
        Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached large checkbox grid long-run scrolling does not cause DX host resize failures");
        lastRenderCount = renderCount;
#endif
    }
}

void TestAttachedComboBoxPopupScrollingStaysStable()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    std::vector<ComboBox::Item> items;
    items.push_back(ComboBox::Item{L"long-entry", L"Very long popup entry that must keep the dropdown width stable while attached scrolling runs"});
    for (size_t index = 0u; index < 47u; ++index)
    {
        items.push_back(ComboBox::Item{std::format(L"value-{:02}", index), std::format(L"Item {:02}", index)});
    }
    combo->SetItems(std::move(items));
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    Require(combo->OnMouseDown(window.Host(), D2D1::Point2F(172.0f, 12.0f), false, 0), "attached combo opens for popup scrolling stability test");
    const D2D1_RECT_F openBounds = combo->GetHitBounds();
    const float openWidthDip     = openBounds.right - openBounds.left;
    Require(openBounds.bottom > combo->GetBounds().bottom, "attached combo popup is open before repeated wheel scrolling");

    const D2D1_POINT_2F firstVisibleItemPoint = D2D1::Point2F(16.0f, 42.0f);
    for (int i = 0; i < 24; ++i)
    {
        Require(combo->OnMouseWheel(window.Host(), firstVisibleItemPoint, -static_cast<float>(WHEEL_DELTA), 0),
                "attached combo popup handles repeated wheel scrolling");
    }

    const D2D1_RECT_F scrolledBounds = combo->GetHitBounds();
    Require(scrolledBounds.bottom > combo->GetBounds().bottom, "attached combo popup remains open after repeated wheel scrolling");
    RequireFloatNear(
        scrolledBounds.right - scrolledBounds.left, openWidthDip, 0.5f, "attached combo popup width remains stable after repeated wheel scrolling");

    Require(combo->OnMouseDown(window.Host(), firstVisibleItemPoint, false, 0),
            "attached combo click selects the first visible popup item after repeated scrolling");
    Require(combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == 24u,
            "attached combo repeated scrolling changes which popup item the first visible row selects");

#ifdef _DEBUG
    const uint64_t initialRenderCount = window.Host().DebugGetRenderCount();
#endif
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetRenderCount() > initialRenderCount, "attached combo popup scrolling repaints after selection");
    Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached combo popup scrolling and repainting does not cause DX host resize failures");
#endif
}

void TestAttachedComboBoxPopupLongRunScrollingStaysStable()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    std::vector<ComboBox::Item> items;
    items.push_back(ComboBox::Item{L"long-entry", L"Very long popup entry that must keep the dropdown width stable while attached scrolling runs"});
    for (size_t index = 0u; index < 47u; ++index)
    {
        items.push_back(ComboBox::Item{std::format(L"value-{:02}", index), std::format(L"Item {:02}", index)});
    }
    combo->SetItems(std::move(items));
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    Require(combo->OnMouseDown(window.Host(), D2D1::Point2F(172.0f, 12.0f), false, 0), "attached combo opens for long-run popup scrolling stability test");
    const D2D1_RECT_F openBounds = combo->GetHitBounds();
    const float openWidthDip     = openBounds.right - openBounds.left;
    Require(openBounds.bottom > combo->GetBounds().bottom, "attached combo popup is open before long-run scrolling");

    const D2D1_POINT_2F firstVisibleItemPoint = D2D1::Point2F(16.0f, 42.0f);
#ifdef _DEBUG
    const uint64_t initialResizeCount = window.Host().DebugGetResizeCount();
    uint64_t lastRenderCount          = window.Host().DebugGetRenderCount();
#endif

    for (int cycleIndex = 0; cycleIndex < 4; ++cycleIndex)
    {
        for (int stepIndex = 0; stepIndex < 12; ++stepIndex)
        {
            Require(combo->OnMouseWheel(window.Host(), firstVisibleItemPoint, -static_cast<float>(WHEEL_DELTA), 0),
                    "attached combo popup handles sustained downward wheel scrolling");
        }
        for (int stepIndex = 0; stepIndex < 10; ++stepIndex)
        {
            Require(combo->OnMouseWheel(window.Host(), firstVisibleItemPoint, static_cast<float>(WHEEL_DELTA), 0),
                    "attached combo popup handles sustained upward wheel scrolling");
        }

        const D2D1_RECT_F scrolledBounds = combo->GetHitBounds();
        Require(scrolledBounds.bottom > combo->GetBounds().bottom, "attached combo popup remains open during long-run scrolling");
        RequireFloatNear(scrolledBounds.right - scrolledBounds.left, openWidthDip, 0.5f, "attached combo popup width remains stable during long-run scrolling");

        RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        window.PumpMessages();

#ifdef _DEBUG
        const uint64_t renderCount = window.Host().DebugGetRenderCount();
        Require(renderCount > lastRenderCount, "attached combo popup long-run scrolling repaints after each sustained cycle");
        Require(window.Host().DebugGetResizeCount() == initialResizeCount, "attached combo popup long-run scrolling does not churn swapchain resizes");
        Require(window.Host().DebugGetResizeFailureCount() == 0u, "attached combo popup long-run scrolling does not cause DX host resize failures");
        lastRenderCount = renderCount;
#endif
    }

    Require(combo->OnMouseDown(window.Host(), firstVisibleItemPoint, false, 0), "attached combo click selects the first visible item after long-run scrolling");
    Require(combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() > 0u,
            "attached combo long-run scrolling keeps the popup selection target away from the initial top item");
}

void TestAttachedHostSameSizeRepaintDoesNotResizeSwapChain()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* label = root->AddChild<Label>(L"alpha");
    label->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    const uint64_t initialResizeCount = window.Host().DebugGetResizeCount();
    const uint64_t initialRenderCount = window.Host().DebugGetRenderCount();
#endif

    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetRenderCount() > initialRenderCount, "attached host repaint renders again without resizing");
    Require(window.Host().DebugGetResizeCount() == initialResizeCount, "attached host repaint at the same size does not call ResizeBuffers");
#endif

    SetWindowPos(window.Hwnd(), nullptr, 0, 0, 420, 240, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetResizeCount() == initialResizeCount + 1u, "attached host resize performs one swapchain resize");
#endif
}

void TestAttachedHostResizeDoesNotFlushD2DInWrongState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* label = root->AddChild<Label>(L"resize");
    label->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 28.0f));
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    const uint64_t initialFlushFailureCount = window.Host().DebugGetSwapChainPrepareD2DFlushFailureCount();
#endif

    SetWindowPos(window.Hwnd(), nullptr, 0, 0, 420, 240, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetSwapChainPrepareD2DFlushFailureCount() == initialFlushFailureCount,
            "attached host resize releases the D2D target without flushing in a wrong state");
#endif
}

void TestAttachedHostRecoversAfterSimulatedDeviceLoss()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root = std::make_unique<Panel>();

    // Build a rich control tree so recovery exercises Grid, Tree, TextField, and Label painting
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 80.0f));
    MultiRowGridModel gridModel(3u);
    grid->SetModel(&gridModel);

    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 80.0f, 280.0f, 140.0f));
    MutableTreeModel treeModel;
    treeModel.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Viewers"},
    });
    tree->SetModel(&treeModel);

    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 140.0f, 280.0f, 168.0f));

    auto* label = root->AddChild<Label>(L"beta");
    label->SetBounds(D2D1::RectF(0.0f, 168.0f, 280.0f, 196.0f));

    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    // Verify resources are populated after the initial render
    Require(window.Host().DebugHasD2DContext(), "device loss: D2D context exists before device loss");
    Require(window.Host().DebugHasFallbackBrush(), "device loss: fallback brush exists before device loss");
    Require(window.Host().DebugGetBrushCacheSize() > 0u, "device loss: brush cache is populated before device loss");
    const size_t textFormatCountBefore = window.Host().DebugGetConfiguredTextFormatCount();

    const uint64_t initialResizeCount = window.Host().DebugGetResizeCount();
    const uint64_t initialRenderCount = window.Host().DebugGetRenderCount();

    // Simulate device loss — caches must be cleared before recovery render
    window.Host().DebugSimulateDeviceLoss();

    Require(! window.Host().DebugHasD2DContext(), "device loss: D2D context is null after device loss");
    Require(! window.Host().DebugHasFallbackBrush(), "device loss: fallback brush is null after device loss");
    Require(window.Host().DebugGetBrushCacheSize() == 0u, "device loss: brush cache is empty after device loss");

    window.PumpMessages();
#endif

    // Force a full recovery render with all four control types
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

#ifdef _DEBUG
    Require(window.Host().DebugGetRenderCount() > initialRenderCount, "device loss: host repaints successfully after simulated device loss");
    Require(window.Host().DebugGetResizeCount() == initialResizeCount, "device loss: no swap chain resize churn during recovery");
    Require(window.Host().DebugGetResizeFailureCount() == 0u, "device loss: recovery does not introduce resize failures");

    // Verify D2D context and caches were rebuilt
    Require(window.Host().DebugHasD2DContext(), "device loss: D2D context is restored after recovery");
    Require(window.Host().DebugHasFallbackBrush(), "device loss: fallback brush is restored after recovery");
    Require(window.Host().DebugGetBrushCacheSize() > 0u, "device loss: brush cache is repopulated after recovery");

    // DWrite text formats are device-independent and must survive device loss
    Require(window.Host().DebugGetConfiguredTextFormatCount() == textFormatCountBefore,
            "device loss: configured text format cache persists through device loss recovery");
#endif
}

} // namespace

void RunRenderingTests()
{
    const auto runTest = [](const char* name, void (*fn)())
    {
        std::cerr << "  [START] " << name << '\n' << std::flush;
        fn();
        std::cerr << "  [DONE] " << name << '\n' << std::flush;
    };

    runTest("TestDxUiCoreControlsDarkVisualBaseline", TestDxUiCoreControlsDarkVisualBaseline);
    runTest("TestDxUiCoreControlsLightVisualBaseline", TestDxUiCoreControlsLightVisualBaseline);
    runTest("TestDxUiHighContrastVisualBaseline", TestDxUiHighContrastVisualBaseline);
    runTest("TestDxUiPopupAndBarsVisualBaseline", TestDxUiPopupAndBarsVisualBaseline);
    runTest("TestDxUiPopupAndBarsAcrylicLightVisualBaseline", TestDxUiPopupAndBarsAcrylicLightVisualBaseline);
    runTest("TestDxUiPageTransitionVisualBaseline", TestDxUiPageTransitionVisualBaseline);
    runTest("TestDxUiAdvancedControlsVisualBaseline", TestDxUiAdvancedControlsVisualBaseline);
    runTest("TestAttachedComboBoxPopupHoverDoesNotRepaintWhenHoveredItemStaysTheSame", TestAttachedComboBoxPopupHoverDoesNotRepaintWhenHoveredItemStaysTheSame);
    runTest("TestAttachedGridPaintHandlesDegenerateScrollbarTracks", TestAttachedGridPaintHandlesDegenerateScrollbarTracks);
    runTest("TestGridHidesBottomClippedTrailingRow", TestGridHidesBottomClippedTrailingRow);
    runTest("TestAttachedLargeGridVisibleWorkStaysBoundedAfterScroll", TestAttachedLargeGridVisibleWorkStaysBoundedAfterScroll);
    runTest("TestAttachedLargeIconBadgeGridVisibleWorkStaysBoundedAfterScroll", TestAttachedLargeIconBadgeGridVisibleWorkStaysBoundedAfterScroll);
    runTest("TestAttachedGridBottomScrollKeepsFirstVisibleRowFlushWithHeader", TestAttachedGridBottomScrollKeepsFirstVisibleRowFlushWithHeader);
    runTest("TestAttachedLargeGridLongRunScrollingStaysBoundedWithoutResizeChurn", TestAttachedLargeGridLongRunScrollingStaysBoundedWithoutResizeChurn);
    runTest("TestAttachedLargeGroupedGridVisibleWorkStaysBoundedAfterScroll", TestAttachedLargeGroupedGridVisibleWorkStaysBoundedAfterScroll);
    runTest("TestAttachedLargeCheckboxGridVisibleWorkStaysBoundedAfterScroll", TestAttachedLargeCheckboxGridVisibleWorkStaysBoundedAfterScroll);
    runTest("TestAttachedLargeGroupedGridLongRunScrollingStaysBoundedWithoutResizeChurn",
            TestAttachedLargeGroupedGridLongRunScrollingStaysBoundedWithoutResizeChurn);
    runTest("TestAttachedLargeCheckboxGridLongRunScrollingStaysBoundedWithoutResizeChurn",
            TestAttachedLargeCheckboxGridLongRunScrollingStaysBoundedWithoutResizeChurn);
    runTest("TestAttachedComboBoxPopupScrollingStaysStable", TestAttachedComboBoxPopupScrollingStaysStable);
    runTest("TestAttachedComboBoxPopupLongRunScrollingStaysStable", TestAttachedComboBoxPopupLongRunScrollingStaysStable);
    runTest("TestAttachedHostSameSizeRepaintDoesNotResizeSwapChain", TestAttachedHostSameSizeRepaintDoesNotResizeSwapChain);
    runTest("TestAttachedHostResizeDoesNotFlushD2DInWrongState", TestAttachedHostResizeDoesNotFlushD2DInWrongState);
    runTest("TestAttachedHostRecoversAfterSimulatedDeviceLoss", TestAttachedHostRecoversAfterSimulatedDeviceLoss);
}
