#include "ViewerPE.Tests.Scope.Tests.ViewerSpaceAnalysisHost.h"
#include "FileMetadataFormatting.h"
#include "ViewerSpace/ViewerSpace.Export.h"

#include <algorithm>
#include <chrono>
#include <cstddef>
#include <exception>
#include <format>
#include <iostream>
#include <limits>
#include <map>
#include <new>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
// WIL headers only: deleted special members, padding and SAL analysis. Owner: ViewerPETests; review on WIL/MSVC updates.
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "TestSupport/TestSupport.Tests.Runner.h"
#include "ViewerSpace/resource.h"
#include "WindowMessages.h"

namespace
{
using namespace std::chrono_literals;

[[nodiscard]] bool ReadViewerSpaceSnapshot(HWND viewerWindow, WndMsg::ViewerSpacePerfDebugSnapshot& snapshot) noexcept
{
    snapshot = {};
    return viewerWindow && IsWindow(viewerWindow) != FALSE &&
           SendMessageW(viewerWindow, WndMsg::kViewerSpaceDebugGetPerfSnapshot, 0u, reinterpret_cast<LPARAM>(&snapshot)) != FALSE;
}

[[nodiscard]] HRESULT FindByAutomationId(IUIAutomation* automation,
                                         IUIAutomationElement* root,
                                         const wchar_t* automationId,
                                         IUIAutomationElement** result) noexcept
{
    if (result == nullptr)
    {
        return E_POINTER;
    }
    *result = nullptr;
    if (! automation || ! root || ! automationId)
    {
        return E_INVALIDARG;
    }

    VARIANT value{};
    VariantInit(&value);
    value.vt      = VT_BSTR;
    value.bstrVal = SysAllocString(automationId);
    if (! value.bstrVal)
    {
        return E_OUTOFMEMORY;
    }
    const auto clearValue = wil::scope_exit([&]() noexcept { VariantClear(&value); });

    wil::com_ptr<IUIAutomationCondition> condition;
    HRESULT hr = automation->CreatePropertyCondition(UIA_AutomationIdPropertyId, value, condition.put());
    if (FAILED(hr) || ! condition)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }
    return root->FindFirst(TreeScope_Subtree, condition.get(), result);
}

[[nodiscard]] bool ReadAnalysisRow(HWND viewerWindow, size_t index, WndMsg::ViewerSpaceAnalysisRowDebugQuery& query) noexcept
{
    query       = {};
    query.index = index;
    return SendMessageW(viewerWindow, WndMsg::kViewerSpaceDebugGetAnalysisRow, 0u, reinterpret_cast<LPARAM>(&query)) != FALSE;
}

[[nodiscard]] bool WaitForAnalysisBuild(HWND viewerWindow, WndMsg::ViewerSpacePerfDebugSnapshot& snapshot) noexcept
{
    RedSalamander::TestSupport::MessagePumpWaitOptions options;
    options.timeout       = 5000ms;
    options.pollInterval  = 10ms;
    options.operationName = L"ViewerSpace analysis projection settle";
    const auto result     = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept { return ReadViewerSpaceSnapshot(viewerWindow, snapshot) && ! snapshot.analysisBuildActive && ! snapshot.analysisDirty; }, options);
    return result.conditionMet;
}

void PumpAnalysisFilterDebounce()
{
    RedSalamander::TestSupport::MessagePumpWaitOptions options;
    options.timeout       = 220ms;
    options.pollInterval  = 10ms;
    options.operationName = L"ViewerSpace analysis filter debounce";
    static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept { return false; }, options));
}

[[nodiscard]] bool CreateAutomationClient(wil::com_ptr<IUIAutomation>& automation, bool& uninitializeCom) noexcept
{
    const HRESULT coHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (coHr == RPC_E_CHANGED_MODE)
    {
        uninitializeCom = false;
    }
    else if (SUCCEEDED(coHr))
    {
        uninitializeCom = true;
    }
    else
    {
        return false;
    }
    return SUCCEEDED(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(automation.put()))) && automation;
}

[[nodiscard]] HRESULT FindVisibleDataRows(IUIAutomation* automation, IUIAutomationElement* grid, IUIAutomationElementArray** rows, int& rowCount) noexcept
{
    rowCount = 0;
    if (rows == nullptr)
    {
        return E_POINTER;
    }
    *rows = nullptr;
    VARIANT value{};
    VariantInit(&value);
    value.vt   = VT_I4;
    value.lVal = UIA_DataItemControlTypeId;
    wil::com_ptr<IUIAutomationCondition> condition;
    HRESULT hr = automation->CreatePropertyCondition(UIA_ControlTypePropertyId, value, condition.put());
    if (FAILED(hr) || ! condition)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }
    hr = grid->FindAll(TreeScope_Children, condition.get(), rows);
    if (FAILED(hr) || ! *rows)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }
    return (*rows)->get_Length(&rowCount);
}
} // namespace

bool CheckViewerSpaceDesktopObservationContracts(std::wstring_view& failure) noexcept
{
    using namespace RedSalamander::TestSupport;
    failure            = {};
    const auto require = [&](bool condition, std::wstring_view reason) noexcept
    {
        if (! condition)
        {
            failure = reason;
        }
        return condition;
    };
    constexpr DWORD owner = 42u;
    DesktopForegroundEvents events;
    if (! require(EvaluateDesktopObservation(true, true, owner, 7u, events) == DesktopObservationFailure::None, L"Foreign foreground is permitted") ||
        ! require(EvaluateDesktopObservation(false, true, owner, 7u, events) == DesktopObservationFailure::GuardUnavailable,
                  L"Missing no-activation guard fails closed") ||
        ! require(EvaluateDesktopObservation(true, false, owner, 7u, events) == DesktopObservationFailure::ObservationUnavailable,
                  L"Unavailable observations fail closed") ||
        ! require(EvaluateDesktopObservation(true, true, 0u, 7u, events) == DesktopObservationFailure::ObservationUnavailable,
                  L"Missing process identity fails closed") ||
        ! require(EvaluateDesktopObservation(true, true, owner, owner, events) == DesktopObservationFailure::OwnedForeground,
                  L"Current owned foreground is rejected"))
    {
        return false;
    }
    events.Record(EVENT_OBJECT_NAMECHANGE);
    if (! require(events.transitions == 0u, L"Non-foreground events cannot manufacture an activation"))
    {
        return false;
    }
    events.Record(EVENT_SYSTEM_FOREGROUND);
    events.Record(EVENT_OBJECT_NAMECHANGE);
    if (! require(events.transitions == 1u && EvaluateDesktopObservation(true, true, owner, 7u, events) == DesktopObservationFailure::OwnedForeground,
                  L"A transient owned activation remains a failure after foreign foreground returns"))
    {
        return false;
    }

    struct Fixture
    {
        unsigned int samples = 0u;
        bool readable        = true;
        bool owned           = false;
        bool inputAvailable  = true;
    } fixture;
    DesktopSnapshotSource source;
    source.context = &fixture;
    source.capture = [](void* context, DesktopStateSnapshot& state) noexcept
    {
        auto& value = *static_cast<Fixture*>(context);
        state       = {};
        if (! value.readable)
        {
            return false;
        }
        ++value.samples;
        // Synthetic observations only: no real pointer or foreground changes.
        state.foregroundWindow    = value.samples % 2u ? nullptr : GetDesktopWindow();
        state.foregroundProcessId = value.owned ? GetCurrentProcessId() : value.samples % 2u ? 0u : GetCurrentProcessId() + 4u;
        state.cursor              = {static_cast<LONG>(value.samples), -static_cast<LONG>(value.samples)};
        state.lastInputTick       = 1000u - value.samples; // Input ticks need not increase.
        state.lastInputAvailable  = value.inputAvailable ? TRUE : FALSE;
        return true;
    };
    ScopedNonActivatingDesktopObserver observer;
    if (! require(observer.Start(source), L"Process-scoped observer starts behind the real activation guard"))
    {
        return false;
    }
    {
        ScopedNonActivatingDesktopObserver nested;
        if (! require(! nested.Start(source) && nested.Failure() == DesktopObservationFailure::GuardUnavailable,
                      L"Nested observer refuses without displacing its live owner"))
        {
            return false;
        }
    }
    if (! require(observer.Sample() && observer.Finish() && observer.Finish(), L"Repeated event-drain barriers complete without desktop input") ||
        ! require(observer.Stats().cursorChanges > 0u && observer.Stats().foregroundChanges > 0u && observer.Stats().inputChanges > 0u &&
                      observer.OwnedForegroundTransitions() == 0u && observer.BarrierCount() == 2u,
                  L"Ambient cursor, foreign foreground and nonmonotonic input changes are diagnostics, not owned activation"))
    {
        return false;
    }
    observer.Stop();
    fixture.readable = false;
    if (! require(! observer.Start(source) && observer.Failure() == DesktopObservationFailure::ObservationUnavailable,
                  L"A failed capture cannot produce passing desktop evidence"))
    {
        return false;
    }
    fixture.readable = true;
    fixture.owned    = true;
    if (! require(! observer.Start(source) && observer.Failure() == DesktopObservationFailure::OwnedForeground,
                  L"Injected owned foreground is rejected without activating any real window"))
    {
        return false;
    }
    fixture.owned          = false;
    fixture.inputAvailable = false;
    if (! require(observer.Start(source) && observer.Finish(), L"Observer restarts after failed captures; unavailable input ticks grant no activation waiver"))
    {
        return false;
    }
    fixture.readable = false;
    if (! require(! observer.Sample() && observer.Failure() == DesktopObservationFailure::ObservationUnavailable,
                  L"Loss of observation during a live scope fails closed"))
    {
        return false;
    }
    observer.Stop();
    fixture.readable = true;
    return require(observer.Start(source) && observer.Finish(), L"Stop and restart retire callback ownership and require a fresh barrier");
}

bool CheckViewerSpaceAnalysisUia(HWND viewerWindow,
                                 std::wstring_view expectedRootPath,
                                 std::wstring_view& failure,
                                 wil::com_ptr<IUIAutomationElement>& retainedGridCell,
                                 uint64_t& cellRequestsBeforeDetach,
                                 RedSalamander::TestSupport::DesktopSnapshotSource desktopSource) noexcept
{
    failure = {};
    retainedGridCell.reset();
    cellRequestsBeforeDetach = 0u;
    RedSalamander::TestSupport::ScopedNonActivatingDesktopObserver desktop;
    const auto logDesktop = [&]()
    {
        const auto& stats = desktop.Stats();
        std::wcout << std::format(L"[INFO] ViewerSpace desktop witness: samples={} cursorChanges={} foregroundChanges={} inputChanges={} ownedTransitions={} "
                                  L"barrierUs={} failure={}\n",
                                  stats.samples,
                                  stats.cursorChanges,
                                  stats.foregroundChanges,
                                  stats.inputChanges,
                                  desktop.OwnedForegroundTransitions(),
                                  stats.barrierUs,
                                  static_cast<unsigned int>(desktop.Failure()));
    };
    const auto fail = [&failure, &desktop, &logDesktop](std::wstring_view reason)
    {
        using RedSalamander::TestSupport::DesktopObservationFailure;
        switch (desktop.Failure())
        {
            case DesktopObservationFailure::GuardUnavailable: failure = L"ViewerSpace UIA requires its active no-activation guard"; break;
            case DesktopObservationFailure::ObservationUnavailable: failure = L"ViewerSpace desktop witness capture and event drain must be available"; break;
            case DesktopObservationFailure::OwnedForeground: failure = L"ViewerSpace UIA never owns desktop foreground, including transient transitions"; break;
            case DesktopObservationFailure::None: failure = reason; break;
        }
        logDesktop();
        return false;
    };

    try
    {
        if (! viewerWindow || IsWindow(viewerWindow) == FALSE)
        {
            return fail(L"ViewerSpace analysis UIA test receives a live viewer window");
        }

        if (! desktop.Start(desktopSource))
        {
            return fail(L"ViewerSpace nonactivating desktop witness starts");
        }

        WndMsg::ViewerSpacePerfDebugSnapshot initial{};
        if (! WaitForAnalysisBuild(viewerWindow, initial) || ! initial.analysisRowCount)
        {
            return fail(L"ViewerSpace analysis projection is published before UIA checks");
        }
        if (initial.analysisHost.visible)
        {
            RedSalamander::TestSupport::MessagePumpWaitOptions frameOptions;
            frameOptions.timeout       = 2000ms;
            frameOptions.pollInterval  = 10ms;
            frameOptions.operationName = L"ViewerSpace visible analysis host first frame";
            const auto frameResult     = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return ReadViewerSpaceSnapshot(viewerWindow, initial) && ! initial.analysisBuildActive && initial.analysisHost.renderCount > 0u;
            }, frameOptions);
            if (! frameResult.conditionMet)
            {
                return fail(L"Visible ViewerSpace analysis host presents its first DxUi frame after rows are ready");
            }
        }
        RECT viewerClient{};
        static_cast<void>(GetClientRect(viewerWindow, &viewerClient));
        const UINT viewerDpi = GetDpiForWindow(viewerWindow);
        const auto logHost   = [](std::wstring_view hostName, const WndMsg::ViewerSpacePerfDebugSnapshot::ChildHost& host)
        {
            std::wcout << std::format(
                L"[INFO] ViewerSpace {} host: available={} hwnd={} visible={} size={}x{} px root={} D2D={} "
                L"render={} resize={} resizeFailure={} presentFailure={} flushFailure={} palette=0x{:08X}/0x{:08X} dark={} highContrast={} mode={}\n",
                hostName,
                host.diagnosticsAvailable,
                host.hasWindow,
                host.visible,
                host.clientWidthPx,
                host.clientHeightPx,
                host.hasRoot,
                host.hasD2DContext,
                host.renderCount,
                host.resizeCount,
                host.resizeFailureCount,
                host.presentFailureCount,
                host.prepareD2DFlushFailureCount,
                host.windowBackgroundArgb,
                host.textArgb,
                host.dark,
                host.highContrast,
                host.presentationMode);
        };
        std::wcout << std::format(
            L"[INFO] ViewerSpace analysis test outer client: {}x{} px ({}x{} DIP at {} DPI), rows={}, bounds={}x{} px.\n",
            viewerClient.right - viewerClient.left,
            viewerClient.bottom - viewerClient.top,
            MulDiv(viewerClient.right - viewerClient.left, USER_DEFAULT_SCREEN_DPI, static_cast<int>(viewerDpi ? viewerDpi : USER_DEFAULT_SCREEN_DPI)),
            MulDiv(viewerClient.bottom - viewerClient.top, USER_DEFAULT_SCREEN_DPI, static_cast<int>(viewerDpi ? viewerDpi : USER_DEFAULT_SCREEN_DPI)),
            viewerDpi,
            initial.analysisRowCount,
            initial.analysisBoundsPx.right - initial.analysisBoundsPx.left,
            initial.analysisBoundsPx.bottom - initial.analysisBoundsPx.top);
        logHost(L"analysis", initial.analysisHost);
        logHost(L"navigation", initial.navigationHost);

        bool uninitializeCom    = false;
        const auto uninitialize = wil::scope_exit([&uninitializeCom]() noexcept
        {
            if (uninitializeCom)
            {
                CoUninitialize();
            }
        });
        wil::com_ptr<IUIAutomation> automation;
        if (! CreateAutomationClient(automation, uninitializeCom))
        {
            return fail(L"UI Automation client initializes for ViewerSpace analysis");
        }

        wil::com_ptr<IUIAutomationElement> root;
        if (FAILED(automation->ElementFromHandle(viewerWindow, root.put())) || ! root)
        {
            return fail(L"ViewerSpace root UIA element is available");
        }
        const auto requireInvokableButton = [&](const wchar_t* automationId) noexcept
        {
            wil::com_ptr<IUIAutomationElement> button;
            if (FAILED(FindByAutomationId(automation.get(), root.get(), automationId, button.put())) || ! button)
            {
                return false;
            }
            int controlType      = 0;
            BSTR name            = nullptr;
            const HRESULT nameHr = button->get_CurrentName(&name);
            const auto freeName  = wil::scope_exit([&name]() noexcept
            {
                if (name != nullptr)
                {
                    SysFreeString(name);
                }
            });
            wil::com_ptr<IUIAutomationInvokePattern> invoke;
            return SUCCEEDED(button->get_CurrentControlType(&controlType)) && controlType == UIA_ButtonControlTypeId && SUCCEEDED(nameHr) && name != nullptr &&
                   name[0] != L'\0' && SUCCEEDED(button->GetCurrentPatternAs(UIA_InvokePatternId, __uuidof(IUIAutomationInvokePattern), invoke.put_void())) &&
                   invoke;
        };
        if (! requireInvokableButton(L"ViewerSpace.Navigation.Back") || ! requireInvokableButton(L"ViewerSpace.Navigation.Forward") ||
            ! requireInvokableButton(L"ViewerSpace.Navigation.Layout") || ! requireInvokableButton(L"ViewerSpace.Navigation.Folders0") ||
            ! requireInvokableButton(L"ViewerSpace.Navigation.Crumb0"))
        {
            return fail(L"ViewerSpace breadcrumb exposes named history, folder-menu, panel disclosure and path controls");
        }
        for (const auto* duplicateId : {L"ViewerSpaceAnalysis.Mode", L"ViewerSpaceAnalysis.Close"})
        {
            wil::com_ptr<IUIAutomationElement> duplicate;
            if (FAILED(FindByAutomationId(automation.get(), root.get(), duplicateId, duplicate.put())) || duplicate)
                return fail(L"ViewerSpace analysis toolbar contains no duplicate mode or close-panel actions");
        }
        wil::com_ptr<IUIAutomationElement> grid;
        if (FAILED(FindByAutomationId(automation.get(), root.get(), L"ViewerSpaceAnalysis.Grid", grid.put())) || ! grid)
        {
            return fail(L"ViewerSpace exposes its named analysis DataGrid");
        }

        wil::com_ptr<IUIAutomationGridPattern> gridPattern;
        if (FAILED(grid->GetCurrentPatternAs(UIA_GridPatternId, __uuidof(IUIAutomationGridPattern), gridPattern.put_void())) || ! gridPattern)
        {
            return fail(L"ViewerSpace analysis grid exposes GridPattern");
        }
        int rowCount    = 0;
        int columnCount = 0;
        if (FAILED(gridPattern->get_CurrentRowCount(&rowCount)) || FAILED(gridPattern->get_CurrentColumnCount(&columnCount)) || rowCount < 1 ||
            columnCount != 4 || static_cast<size_t>(rowCount) != initial.analysisRowCount)
        {
            return fail(L"ViewerSpace GridPattern reports the complete published row and four-column projection including cleanup marks");
        }

        wil::com_ptr<IUIAutomationElementArray> initialVisibleRows;
        int initialVisibleRowCount = 0;
        if (FAILED(FindVisibleDataRows(automation.get(), grid.get(), initialVisibleRows.put(), initialVisibleRowCount)) || initialVisibleRowCount < 1 ||
            static_cast<size_t>(initialVisibleRowCount) > initial.analysisRowCount)
        {
            return fail(L"ViewerSpace UIA materializes a bounded visible row set");
        }
        if (initial.analysisRowCount >= 64u && static_cast<size_t>(initialVisibleRowCount) >= initial.analysisRowCount)
        {
            return fail(L"ViewerSpace UIA does not materialize the offscreen analysis rows");
        }

        size_t selectedRowIndex = initial.analysisRowCount;
        WndMsg::ViewerSpaceAnalysisRowDebugQuery selectedRow{};
        for (size_t index = 0u; index < initial.analysisRowCount; ++index)
        {
            WndMsg::ViewerSpaceAnalysisRowDebugQuery candidate{};
            if (ReadAnalysisRow(viewerWindow, index, candidate) && ! candidate.other && ! candidate.directory && candidate.name[0] != L'\0')
            {
                selectedRowIndex = index;
                selectedRow      = candidate;
                break;
            }
        }
        if (selectedRowIndex >= initial.analysisRowCount)
        {
            return fail(L"ViewerSpace fixture contains a named retained row for selection and filtering");
        }

        WndMsg::ViewerSpaceAnalysisRowDebugQuery filterTarget{};
        bool haveFilterTarget = false;
        for (size_t index = selectedRowIndex + 1u; index < initial.analysisRowCount; ++index)
        {
            WndMsg::ViewerSpaceAnalysisRowDebugQuery candidate{};
            if (! ReadAnalysisRow(viewerWindow, index, candidate) || candidate.other || candidate.name[0] == L'\0')
            {
                continue;
            }
            const std::wstring_view candidateName(candidate.name);
            if (candidateName != std::wstring_view(selectedRow.name) && std::wstring_view(selectedRow.name).find(candidateName) == std::wstring_view::npos)
            {
                filterTarget     = candidate;
                haveFilterTarget = true;
                break;
            }
        }
        if (! haveFilterTarget)
        {
            return fail(L"ViewerSpace fixture contains distinct retained rows for persistent selection and filtering");
        }

        wil::com_ptr<IUIAutomationElement> selectionCell;
        if (selectedRowIndex > static_cast<size_t>((std::numeric_limits<int>::max)()) ||
            FAILED(gridPattern->GetItem(static_cast<int>(selectedRowIndex), 0, selectionCell.put())) || ! selectionCell)
        {
            return fail(L"ViewerSpace exposes the row selected for persistence checks through GridItem navigation");
        }
        wil::com_ptr<IUIAutomationTreeWalker> selectionWalker;
        if (FAILED(automation->get_ControlViewWalker(selectionWalker.put())) || ! selectionWalker)
        {
            return fail(L"UI Automation control-view walker is available for selection persistence checks");
        }
        wil::com_ptr<IUIAutomationElement> selectionRow;
        wil::com_ptr<IUIAutomationSelectionItemPattern> selectionPattern;
        if (FAILED(selectionWalker->GetParentElement(selectionCell.get(), selectionRow.put())) || ! selectionRow ||
            FAILED(selectionRow->GetCurrentPatternAs(UIA_SelectionItemPatternId, __uuidof(IUIAutomationSelectionItemPattern), selectionPattern.put_void())) ||
            ! selectionPattern || FAILED(selectionPattern->Select()))
        {
            return fail(L"ViewerSpace row selection can be established through UIA before filtering");
        }
        WndMsg::ViewerSpacePerfDebugSnapshot selectedBeforeFilter{};
        if (! ReadViewerSpaceSnapshot(viewerWindow, selectedBeforeFilter) || selectedBeforeFilter.selectedItemId != selectedRow.itemId)
        {
            return fail(L"ViewerSpace starts filter persistence checks with the row linked to the treemap selection");
        }

        wil::com_ptr<IUIAutomationElement> filter;
        if (FAILED(FindByAutomationId(automation.get(), root.get(), L"ViewerSpaceAnalysis.Filter", filter.put())) || ! filter)
        {
            return fail(L"ViewerSpace exposes its named analysis filter");
        }
        wil::com_ptr<IUIAutomationValuePattern> filterValue;
        if (FAILED(filter->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), filterValue.put_void())) || ! filterValue)
        {
            return fail(L"ViewerSpace analysis filter exposes ValuePattern");
        }

        const wil::unique_bstr filterTargetValue(SysAllocString(filterTarget.name));
        if (! filterTargetValue || FAILED(filterValue->SetValue(filterTargetValue.get())))
        {
            return fail(L"ViewerSpace analysis filter accepts a distinct retained row name through UIA");
        }
        PumpAnalysisFilterDebounce();
        WndMsg::ViewerSpacePerfDebugSnapshot filtered{};
        if (! WaitForAnalysisBuild(viewerWindow, filtered) || filtered.analysisRowCount < 1u || filtered.analysisRowCount > initial.analysisRowCount)
        {
            return fail(L"ViewerSpace filter publishes a settled retained-row projection");
        }
        bool foundFilteredTarget    = false;
        bool selectedRowFilteredOut = true;
        for (size_t index = 0u; index < filtered.analysisRowCount; ++index)
        {
            WndMsg::ViewerSpaceAnalysisRowDebugQuery candidate{};
            if (! ReadAnalysisRow(viewerWindow, index, candidate))
            {
                return fail(L"ViewerSpace filtered row identities remain queryable");
            }
            if (candidate.other)
            {
                continue;
            }
            const std::wstring_view candidateName(candidate.name);
            if (candidateName.find(filterTarget.name) == std::wstring_view::npos)
            {
                return fail(L"ViewerSpace retains only names matching the filter within the current-folder scope");
            }
            foundFilteredTarget    = foundFilteredTarget || candidate.itemId == filterTarget.itemId;
            selectedRowFilteredOut = selectedRowFilteredOut && candidate.itemId != selectedRow.itemId;
        }
        if (! foundFilteredTarget || ! selectedRowFilteredOut)
        {
            return fail(L"ViewerSpace filter includes its target while temporarily excluding the selected row");
        }

        // Clear through the same UIA ValuePattern so the full projection and its virtual snapshot can be checked.
        const wil::unique_bstr emptyFilterValue(SysAllocString(L""));
        if (! emptyFilterValue || FAILED(filterValue->SetValue(emptyFilterValue.get())))
        {
            return fail(L"ViewerSpace analysis filter clears through UIA");
        }
        PumpAnalysisFilterDebounce();
        WndMsg::ViewerSpacePerfDebugSnapshot cleared{};
        if (! WaitForAnalysisBuild(viewerWindow, cleared) || cleared.analysisRowCount != initial.analysisRowCount)
        {
            return fail(L"ViewerSpace clearing the retained-name filter restores the complete row projection");
        }
        bool stableTargetToken = false;
        selectedRowIndex       = cleared.analysisRowCount;
        for (size_t index = 0u; index < cleared.analysisRowCount; ++index)
        {
            WndMsg::ViewerSpaceAnalysisRowDebugQuery candidate{};
            if (ReadAnalysisRow(viewerWindow, index, candidate) && candidate.itemId == selectedRow.itemId)
            {
                stableTargetToken = candidate.token == selectedRow.token && candidate.generation == selectedRow.generation;
                selectedRowIndex  = index;
                break;
            }
        }
        if (! stableTargetToken)
        {
            return fail(L"ViewerSpace stable analysis row tokens survive filter removal and reprojection");
        }
        if (cleared.selectedItemId != selectedRow.itemId)
        {
            return fail(L"ViewerSpace clearing a filter restores the retained treemap and grid selection");
        }
        wil::com_ptr<IUIAutomationElement> restoredCell;
        wil::com_ptr<IUIAutomationGridPattern> restoredGridPattern;
        if (selectedRowIndex >= cleared.analysisRowCount || selectedRowIndex > static_cast<size_t>((std::numeric_limits<int>::max)()) ||
            FAILED(grid->GetCurrentPatternAs(UIA_GridPatternId, __uuidof(IUIAutomationGridPattern), restoredGridPattern.put_void())) || ! restoredGridPattern ||
            FAILED(restoredGridPattern->GetItem(static_cast<int>(selectedRowIndex), 0, restoredCell.put())) || ! restoredCell)
        {
            return fail(L"ViewerSpace restores the selected row's virtual UIA cell after filter removal");
        }
        wil::com_ptr<IUIAutomationTreeWalker> restoredWalker;
        wil::com_ptr<IUIAutomationElement> restoredRow;
        wil::com_ptr<IUIAutomationSelectionItemPattern> restoredSelection;
        BOOL isRestoredSelection = FALSE;
        if (FAILED(automation->get_ControlViewWalker(restoredWalker.put())) || ! restoredWalker ||
            FAILED(restoredWalker->GetParentElement(restoredCell.get(), restoredRow.put())) || ! restoredRow ||
            FAILED(restoredRow->GetCurrentPatternAs(UIA_SelectionItemPatternId, __uuidof(IUIAutomationSelectionItemPattern), restoredSelection.put_void())) ||
            ! restoredSelection || FAILED(restoredSelection->get_CurrentIsSelected(&isRestoredSelection)) || isRestoredSelection == FALSE)
        {
            return fail(L"ViewerSpace UIA SelectionItemPattern reports the retained row selected after filter removal");
        }

        ViewerSpaceExport::Row exportRow;
        exportRow.name         = "=SUM(1,2)\".txt";
        exportRow.path         = "atlas/東京";
        exportRow.bytes        = 1234u;
        exportRow.excluded     = true;
        const std::string json = ViewerSpaceExport::JsonRow(exportRow);
        Common::Json::UniqueDocument exported{yyjson_read(json.data(), json.size(), 0u)};
        auto* object = exported ? yyjson_doc_get_root(exported.get()) : nullptr;
        if (! object || ! yyjson_equals_str(yyjson_obj_get(object, "name"), exportRow.name.c_str()) ||
            ! yyjson_equals_str(yyjson_obj_get(object, "path"), exportRow.path.c_str()) || yyjson_get_uint(yyjson_obj_get(object, "logicalBytes")) != 1234u ||
            ! yyjson_is_null(yyjson_obj_get(object, "allocationBytes")) || ! yyjson_is_null(yyjson_obj_get(object, "fileCount")))
            return fail(L"JSON export preserves Unicode, literal names and unavailable facts");
        const std::string csv = ViewerSpaceExport::CsvRow(exportRow);
        if (! csv.starts_with("\"'=SUM(1,2)\"\".txt\",\"atlas/東京\",false,false,1234,0,,,") || ! csv.ends_with("true,false,\r\n"))
            return fail(L"CSV export quotes literal strings and identifies spreadsheet escaping");
        exportRow.synthetic         = true;
        const std::string otherJson = ViewerSpaceExport::JsonRow(exportRow);
        Common::Json::UniqueDocument otherExport{yyjson_read(otherJson.data(), otherJson.size(), 0u)};
        if (! otherExport || ! yyjson_is_null(yyjson_obj_get(yyjson_doc_get_root(otherExport.get()), "path")))
            return fail(L"Synthetic Other never gains an invented export path");

        wil::com_ptr<IUIAutomationElement> layoutButton;
        wil::unique_bstr initialMode;
        if (FAILED(FindByAutomationId(automation.get(), root.get(), L"ViewerSpace.Navigation.Layout", layoutButton.put())) || ! layoutButton ||
            FAILED(layoutButton->get_CurrentName(initialMode.put())) || ! initialMode)
            return fail(L"ViewerSpace exposes a compact, named Layout button");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_VIEW_EXPLORER, 0);
        WndMsg::ViewerSpacePerfDebugSnapshot analystMode{};
        wil::unique_bstr changedMode;
        if (! WaitForAnalysisBuild(viewerWindow, analystMode) || analystMode.selectedItemId != selectedRow.itemId ||
            FAILED(layoutButton->get_CurrentName(changedMode.put())) || ! changedMode || std::wstring_view(changedMode.get()) == initialMode.get())
            return fail(L"ViewerSpace layout command updates the compact control and preserves selection");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_VIEW_ANALYST, 0);
        if (! WaitForAnalysisBuild(viewerWindow, analystMode))
            return fail(L"Bottom layout settles");

        const auto workspace = [viewerWindow](WndMsg::ViewerSpaceWorkspaceDebugSnapshot& state) noexcept
        { return SendMessageW(viewerWindow, WndMsg::kViewerSpaceDebugGetWorkspace, 0u, reinterpret_cast<LPARAM>(&state)) != FALSE; };
        WndMsg::ViewerSpaceWorkspaceDebugSnapshot original{};
        if (! workspace(original) || original.exclusions || original.originalViewBytes != original.shownViewBytes)
            return fail(L"Workspace starts with original and shown totals equal");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_TYPE_FIRST, 0);
        WndMsg::ViewerSpacePerfDebugSnapshot typeFiltered{};
        if (! WaitForAnalysisBuild(viewerWindow, typeFiltered) || typeFiltered.analysisRowCount != 0u)
            return fail(L"Type filtering removes known nonmatching retained current-folder entries");
        WndMsg::ViewerSpaceWorkspaceDebugSnapshot filterTotals{};
        if (! workspace(filterTotals) || filterTotals.shownViewBytes != original.shownViewBytes || filterTotals.discoveredBytes != original.discoveredBytes)
            return fail(L"A current-folder type filter does not change map or discovered totals");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_TYPE_ALL, 0);
        if (! WaitForAnalysisBuild(viewerWindow, typeFiltered))
            return fail(L"Clearing the type filter restores the projection");
        if (! workspace(filterTotals) || ! filterTotals.factsReady || filterTotals.shownFileCount != 2002u)
            return fail(L"Derived file count accounts for all retained nested and direct files");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_WEIGHT_COUNT, 0);
        if (! workspace(filterTotals) || ! filterTotals.fileCountWeight || filterTotals.shownViewBytes != original.shownViewBytes)
            return fail(L"File count weighting preserves original logical bytes");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_WEIGHT_SIZE, 0);
        if (! WaitForAnalysisBuild(viewerWindow, typeFiltered))
            return fail(L"Normal Size weighting is restored before exclusion");
        SendMessageW(viewerWindow, WM_KEYDOWN, VK_RETURN, 0); // Dispatch only; real foreground keys are covered separately.
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_EXCLUDE, 0);
        RedSalamander::TestSupport::MessagePumpWaitOptions exclusionWait;
        exclusionWait.timeout       = 5000ms;
        exclusionWait.pollInterval  = 10ms;
        exclusionWait.operationName = L"ViewerSpace exclusion retained-path binding";
        WndMsg::ViewerSpaceWorkspaceDebugSnapshot excluded{};
        if (! RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept { return workspace(excluded) && ! excluded.rebinding && excluded.exclusions == 1u; },
                                                            exclusionWait)
                  .conditionMet ||
            excluded.discoveredBytes != original.discoveredBytes || excluded.originalViewBytes != original.originalViewBytes ||
            excluded.shownViewBytes + selectedRow.logicalBytes != original.shownViewBytes || excluded.unresolved || excluded.ownedBytes > 4u * 1024u * 1024u)
        {
            std::wcout << std::format(L"[INFO] Exclusion: count={} unresolved={} binding={} discovered={}/{} original={}/{} shown={}/{} selected={} owned={}\n",
                                      excluded.exclusions,
                                      excluded.unresolved,
                                      excluded.rebinding,
                                      excluded.discoveredBytes,
                                      original.discoveredBytes,
                                      excluded.originalViewBytes,
                                      original.originalViewBytes,
                                      excluded.shownViewBytes,
                                      original.shownViewBytes,
                                      selectedRow.logicalBytes,
                                      excluded.ownedBytes);
            return fail(L"Exclusion subtracts exactly one retained item, preserves discovered totals and stays bounded");
        }
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_MARK_CLEANUP, 0);
        if (! workspace(excluded) || excluded.cleanupMarks)
            return fail(L"Excluded rows cannot become cleanup targets");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_RESTORE_ALL, 0);
        WndMsg::ViewerSpaceWorkspaceDebugSnapshot restored{};
        if (! workspace(restored) || restored.exclusions || restored.shownViewBytes != original.shownViewBytes)
            return fail(L"Restore all restores the exact calculation");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_MARK_CLEANUP, 0);
        if (! workspace(restored) || restored.cleanupMarks != 1u || restored.reviewing)
            return fail(L"Marking is independent from exclusion and does not execute cleanup");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_REVIEW_CLEANUP, 0);
        if (! workspace(restored) || ! restored.reviewing || ! WaitForAnalysisBuild(viewerWindow, analystMode) || analystMode.analysisRowCount != 1u)
            return fail(L"Cleanup review freezes the marked real target list");
        SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_SEND_RECYCLE, 0);
        if (! workspace(restored) || restored.reviewing || restored.cleanupMarks)
            return fail(L"Confirmed cleanup hands the reviewed list to the host and clears marks");
        if (! WaitForAnalysisBuild(viewerWindow, analystMode))
            return fail(L"Folder analysis returns after cleanup handoff");

        const auto readByteSortPrefix =
            [viewerWindow](bool ascending, std::map<uint64_t, std::vector<uint32_t>>& tiedItemIds, std::wstring_view& failureText) noexcept
        {
            if (SendMessageW(viewerWindow, WndMsg::kViewerSpaceDebugSetAnalysisSort, 1u, ascending ? TRUE : FALSE) == FALSE)
            {
                failureText = L"ViewerSpace accepts the deterministic logical-byte sort hook";
                return false;
            }
            PumpAnalysisFilterDebounce();
            WndMsg::ViewerSpacePerfDebugSnapshot sorted{};
            if (! WaitForAnalysisBuild(viewerWindow, sorted) || sorted.analysisRowCount == 0u)
            {
                failureText = L"ViewerSpace publishes rows after byte-sort changes";
                return false;
            }
            const size_t prefixCount = std::min<size_t>(sorted.analysisRowCount, 128u);
            uint64_t previousBytes   = 0u;
            bool havePrevious        = false;
            for (size_t index = 0u; index < prefixCount; ++index)
            {
                WndMsg::ViewerSpaceAnalysisRowDebugQuery row{};
                if (! ReadAnalysisRow(viewerWindow, index, row))
                {
                    failureText = L"ViewerSpace debug row query remains available during byte-sort checks";
                    return false;
                }
                if (havePrevious && (ascending ? row.logicalBytes < previousBytes : row.logicalBytes > previousBytes))
                {
                    failureText = L"ViewerSpace logical-byte rows follow the requested sort direction";
                    return false;
                }
                previousBytes = row.logicalBytes;
                havePrevious  = true;
                tiedItemIds[row.logicalBytes].push_back(row.itemId);
            }
            return true;
        };
        std::map<uint64_t, std::vector<uint32_t>> ascendingTieOrder;
        std::map<uint64_t, std::vector<uint32_t>> descendingTieOrder;
        if (! readByteSortPrefix(true, ascendingTieOrder, failure) || ! readByteSortPrefix(false, descendingTieOrder, failure))
        {
            return false;
        }
        for (const auto& [byteCount, ascendingIds] : ascendingTieOrder)
        {
            const auto descending = descendingTieOrder.find(byteCount);
            if (descending != descendingTieOrder.end())
            {
                const size_t sharedCount = std::min(ascendingIds.size(), descending->second.size());
                if (! std::equal(ascendingIds.begin(), ascendingIds.begin() + static_cast<ptrdiff_t>(sharedCount), descending->second.begin()))
                {
                    return fail(L"ViewerSpace equal-byte tie ordering stays ascending and stable in either byte direction");
                }
            }
        }
        WndMsg::ViewerSpacePerfDebugSnapshot sortedSnapshot{};
        if (! ReadViewerSpaceSnapshot(viewerWindow, sortedSnapshot) || sortedSnapshot.analysisRowCount != cleared.analysisRowCount)
        {
            return fail(L"ViewerSpace byte sort preserves the current-folder denominator");
        }
        selectedRowIndex = sortedSnapshot.analysisRowCount;
        for (size_t index = 0u; index < sortedSnapshot.analysisRowCount; ++index)
        {
            WndMsg::ViewerSpaceAnalysisRowDebugQuery candidate{};
            if (ReadAnalysisRow(viewerWindow, index, candidate) && candidate.itemId == selectedRow.itemId && candidate.token == selectedRow.token)
            {
                selectedRowIndex = index;
                break;
            }
        }
        if (selectedRowIndex >= sortedSnapshot.analysisRowCount)
        {
            return fail(L"ViewerSpace byte sort preserves the selected stable row token");
        }

        // A forced accessibility refresh deterministically measures model cell callbacks without repaint sampling.
        wil::com_ptr<IUIAutomationElementArray> visibleRows;
        int visibleRowCount = 0;
        if (FAILED(FindVisibleDataRows(automation.get(), grid.get(), visibleRows.put(), visibleRowCount)) || visibleRowCount < 1 ||
            static_cast<size_t>(visibleRowCount) > cleared.analysisRowCount)
        {
            return fail(L"ViewerSpace full projection exposes visible UIA rows after filter clear");
        }
        WndMsg::ViewerSpacePerfDebugSnapshot beforeRefresh{};
        if (! ReadViewerSpaceSnapshot(viewerWindow, beforeRefresh) ||
            SendMessageW(viewerWindow, WndMsg::kViewerSpaceDebugRefreshAnalysisAccessibility, 0u, 0u) == FALSE)
        {
            return fail(L"ViewerSpace answers the forced analysis accessibility-refresh hook");
        }
        WndMsg::ViewerSpacePerfDebugSnapshot afterRefresh{};
        if (! ReadViewerSpaceSnapshot(viewerWindow, afterRefresh) || afterRefresh.analysisCellRequests < beforeRefresh.analysisCellRequests)
        {
            return fail(L"ViewerSpace analysis cell-request counter remains monotonic across UIA refresh");
        }
        const uint64_t requestDelta     = afterRefresh.analysisCellRequests - beforeRefresh.analysisCellRequests;
        const uint64_t cellRequestBound = static_cast<uint64_t>(visibleRowCount + 1) * static_cast<uint64_t>(columnCount) * 2u;
        if (requestDelta > cellRequestBound)
        {
            return fail(L"ViewerSpace accessibility refresh reads only visible rows plus selection overscan");
        }

        wil::com_ptr<IUIAutomationGridPattern> currentGridPattern;
        if (FAILED(grid->GetCurrentPatternAs(UIA_GridPatternId, __uuidof(IUIAutomationGridPattern), currentGridPattern.put_void())) || ! currentGridPattern)
        {
            return fail(L"ViewerSpace grid remains available after its projection refresh");
        }
        wil::com_ptr<IUIAutomationElement> selectedCell;
        if (selectedRowIndex >= cleared.analysisRowCount || selectedRowIndex > static_cast<size_t>((std::numeric_limits<int>::max)()) ||
            FAILED(currentGridPattern->GetItem(static_cast<int>(selectedRowIndex), 0, selectedCell.put())) || ! selectedCell)
        {
            return fail(L"ViewerSpace exposes the selected retained row through GridItem navigation");
        }
        wil::com_ptr<IUIAutomationGridItemPattern> cellPattern;
        if (FAILED(selectedCell->GetCurrentPatternAs(UIA_GridItemPatternId, __uuidof(IUIAutomationGridItemPattern), cellPattern.put_void())) || ! cellPattern)
        {
            return fail(L"ViewerSpace analysis cells expose GridItemPattern");
        }
        wil::com_ptr<IUIAutomationTreeWalker> walker;
        if (FAILED(automation->get_ControlViewWalker(walker.put())) || ! walker)
        {
            return fail(L"UI Automation control-view walker is available");
        }
        wil::com_ptr<IUIAutomationElement> selectedRowElement;
        if (FAILED(walker->GetParentElement(selectedCell.get(), selectedRowElement.put())) || ! selectedRowElement)
        {
            return fail(L"ViewerSpace GridItem resolves to its semantic row element");
        }
        wil::com_ptr<IUIAutomationSelectionItemPattern> selectionItem;
        if (FAILED(
                selectedRowElement->GetCurrentPatternAs(UIA_SelectionItemPatternId, __uuidof(IUIAutomationSelectionItemPattern), selectionItem.put_void())) ||
            ! selectionItem || FAILED(selectionItem->Select()))
        {
            return fail(L"ViewerSpace analysis row selection works through UIA SelectionItemPattern");
        }
        WndMsg::ViewerSpacePerfDebugSnapshot selected{};
        if (! ReadViewerSpaceSnapshot(viewerWindow, selected) || selected.selectedItemId != selectedRow.itemId)
        {
            return fail(L"ViewerSpace UIA row selection updates the linked treemap selection");
        }
        wil::com_ptr<IUIAutomationSelectionPattern> gridSelection;
        if (FAILED(grid->GetCurrentPatternAs(UIA_SelectionPatternId, __uuidof(IUIAutomationSelectionPattern), gridSelection.put_void())) || ! gridSelection)
        {
            return fail(L"ViewerSpace analysis grid exposes SelectionPattern");
        }
        wil::com_ptr<IUIAutomationElementArray> selectedElements;
        int selectedElementCount = 0;
        if (FAILED(gridSelection->GetCurrentSelection(selectedElements.put())) || ! selectedElements ||
            FAILED(selectedElements->get_Length(&selectedElementCount)) || selectedElementCount != 1)
        {
            return fail(L"ViewerSpace UIA SelectionPattern reports the single linked selected row");
        }
        wil::com_ptr<IUIAutomationElement> inspectorName;
        if (FAILED(FindByAutomationId(automation.get(), root.get(), L"ViewerSpaceAnalysis.Inspector.Name", inspectorName.put())) || ! inspectorName)
        {
            return fail(L"ViewerSpace exposes the selected-item inspector name");
        }
        wil::unique_bstr inspectorNameText;
        if (FAILED(inspectorName->get_CurrentName(inspectorNameText.put())) || ! inspectorNameText ||
            std::wstring_view(inspectorNameText.get()) != std::wstring_view(selectedRow.name))
        {
            return fail(L"ViewerSpace inspector follows the selected UIA row");
        }

        const auto checkModified = [&](std::wstring_view name, std::wstring_view expected) noexcept
        {
            size_t rowIndex = initial.analysisRowCount;
            for (size_t index = 0u; index < initial.analysisRowCount; ++index)
            {
                WndMsg::ViewerSpaceAnalysisRowDebugQuery row{};
                if (ReadAnalysisRow(viewerWindow, index, row) && std::wstring_view(row.name) == name)
                {
                    rowIndex = index;
                    break;
                }
            }
            wil::com_ptr<IUIAutomationElement> cell;
            wil::com_ptr<IUIAutomationElement> rowElement;
            wil::com_ptr<IUIAutomationSelectionItemPattern> rowSelection;
            wil::com_ptr<IUIAutomationElement> modified;
            wil::unique_bstr value;
            if (rowIndex >= initial.analysisRowCount || FAILED(currentGridPattern->GetItem(static_cast<int>(rowIndex), 0, cell.put())) || ! cell ||
                FAILED(restoredWalker->GetParentElement(cell.get(), rowElement.put())) || ! rowElement ||
                FAILED(rowElement->GetCurrentPatternAs(UIA_SelectionItemPatternId, __uuidof(IUIAutomationSelectionItemPattern), rowSelection.put_void())) ||
                ! rowSelection || FAILED(rowSelection->Select()) ||
                FAILED(FindByAutomationId(automation.get(), root.get(), L"ViewerSpaceAnalysis.Inspector.Modified", modified.put())) || ! modified ||
                FAILED(modified->get_CurrentName(value.put())) || ! value || std::wstring_view(value.get()) != expected)
            {
                return false;
            }
            BOOL offscreen = TRUE;
            RECT bounds{};
            return SUCCEEDED(modified->get_CurrentIsOffscreen(&offscreen)) && offscreen == FALSE &&
                   SUCCEEDED(modified->get_CurrentBoundingRectangle(&bounds)) && bounds.right > bounds.left && bounds.bottom > bounds.top;
        };
        const std::wstring expectedFileTime =
            Common::FileMetadata::FormatDisplayFields({kViewerSpaceFileModifiedTime, 0u}, Common::FileMetadata::DisplayProfile::CompactDetails).localTime;
        const std::wstring expectedFolderTime =
            Common::FileMetadata::FormatDisplayFields({kViewerSpaceFolderModifiedTime, 0u}, Common::FileMetadata::DisplayProfile::CompactDetails).localTime;
        if (expectedFileTime.empty() || expectedFolderTime.empty() || ! checkModified(L"folder", expectedFolderTime) ||
            ! checkModified(L"item_0000.bin", expectedFileTime) || ! checkModified(L"item_0001.bin", L"Unavailable") ||
            ! checkModified(L"item_0002.bin", L"Unavailable") || FAILED(selectionItem->Select()))
        {
            return fail(L"ViewerSpace shows retained folder/file modified times and rejects missing or invalid dates in its visible inspector");
        }
        std::wcout << L"[PASS] ViewerSpace modified timestamps: folder, file, missing and invalid provider dates.\n";

        // Verify a deliberately long virtual path keeps its accessible value through
        // narrow/wide/narrow resize transitions; visual ellipsis belongs to the UI.
        if (expectedRootPath.empty())
        {
            return fail(L"ViewerSpace long-path fixture supplies a nonempty virtual root");
        }
        std::wstring expectedPath(expectedRootPath);
        expectedPath.append(L"/");
        expectedPath.append(selectedRow.name);
        RECT windowRect{};
        RECT clientRect{};
        MONITORINFO monitor{.cbSize = sizeof(MONITORINFO)};
        const HMONITOR targetMonitor = MonitorFromWindow(viewerWindow, MONITOR_DEFAULTTONEAREST);
        if (GetWindowRect(viewerWindow, &windowRect) == FALSE || GetClientRect(viewerWindow, &clientRect) == FALSE ||
            ! GetMonitorInfoW(targetMonitor, &monitor))
        {
            return fail(L"ViewerSpace long-path resize fixture resolves its owned viewer and monitor bounds");
        }
        const UINT observedDpi     = GetDpiForWindow(viewerWindow);
        const UINT dpi             = observedDpi ? observedDpi : USER_DEFAULT_SCREEN_DPI;
        const int nonClientWidthPx = (windowRect.right - windowRect.left) - (clientRect.right - clientRect.left);
        const int maxClientWidthDip =
            MulDiv((monitor.rcWork.right - monitor.rcWork.left) - nonClientWidthPx - 8, static_cast<int>(USER_DEFAULT_SCREEN_DPI), static_cast<int>(dpi));
        constexpr int kNarrowWidthDip = 504;
        const int wideWidthDip        = (std::min)(900, maxClientWidthDip);
        if (wideWidthDip <= kNarrowWidthDip)
        {
            return fail(L"ViewerSpace test monitor can host both narrow and wide inspector layouts");
        }
        const auto verifyPathAtWidth = [&](int widthDip) noexcept
        {
            RECT beforeWindow{};
            RECT beforeClient{};
            if (GetWindowRect(viewerWindow, &beforeWindow) == FALSE || GetClientRect(viewerWindow, &beforeClient) == FALSE)
            {
                return false;
            }
            const int targetClientWidthPx = MulDiv(widthDip, static_cast<int>(dpi), static_cast<int>(USER_DEFAULT_SCREEN_DPI));
            const int targetWindowWidthPx = (beforeWindow.right - beforeWindow.left) + targetClientWidthPx - (beforeClient.right - beforeClient.left);
            if (targetWindowWidthPx <= 0 || targetWindowWidthPx > monitor.rcWork.right - monitor.rcWork.left)
            {
                return false;
            }
            const int targetX = (std::clamp)(beforeWindow.left, monitor.rcWork.left, monitor.rcWork.right - targetWindowWidthPx);
            if (SetWindowPos(viewerWindow,
                             nullptr,
                             targetX,
                             beforeWindow.top,
                             targetWindowWidthPx,
                             beforeWindow.bottom - beforeWindow.top,
                             SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSENDCHANGING) == FALSE)
            {
                return false;
            }
            RedSalamander::TestSupport::MessagePumpWaitOptions resizeOptions;
            resizeOptions.timeout       = 2000ms;
            resizeOptions.pollInterval  = 10ms;
            resizeOptions.operationName = L"ViewerSpace inspector accessibility after client resize";
            const auto resized          = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                RECT currentClient{};
                WndMsg::ViewerSpacePerfDebugSnapshot current{};
                if (GetClientRect(viewerWindow, &currentClient) == FALSE || ! ReadViewerSpaceSnapshot(viewerWindow, current) || current.analysisBuildActive ||
                    current.analysisHwnd == nullptr || IsWindowVisible(current.analysisHwnd) == FALSE)
                {
                    return false;
                }
                const int actualClientWidthPx = currentClient.right - currentClient.left;
                return actualClientWidthPx >= targetClientWidthPx - 1 && actualClientWidthPx <= targetClientWidthPx + 1;
            },
                resizeOptions);
            if (! resized.conditionMet)
            {
                return false;
            }
            wil::com_ptr<IUIAutomationElement> pathValue;
            wil::unique_bstr accessibleName;
            wil::unique_bstr accessibleHelp;
            if (FAILED(FindByAutomationId(automation.get(), root.get(), L"ViewerSpaceAnalysis.Inspector.Path", pathValue.put())) || ! pathValue ||
                FAILED(pathValue->get_CurrentName(accessibleName.put())) || ! accessibleName || FAILED(pathValue->get_CurrentHelpText(accessibleHelp.put())) ||
                ! accessibleHelp || std::wstring_view(accessibleName.get()) != expectedPath || std::wstring_view(accessibleHelp.get()) != expectedPath)
            {
                return false;
            }
            return checkModified(selectedRow.name, expectedFileTime) && desktop.Sample();
        };
        if (! verifyPathAtWidth(kNarrowWidthDip) || ! verifyPathAtWidth(wideWidthDip) || ! verifyPathAtWidth(kNarrowWidthDip))
        {
            return fail(L"ViewerSpace selected inspector path retains its full UIA Name and HelpText through no-activate narrow/wide/narrow resizes");
        }

        const auto invoke = [&](const wchar_t* id) noexcept
        {
            wil::com_ptr<IUIAutomationElement> element;
            wil::com_ptr<IUIAutomationInvokePattern> pattern;
            return SUCCEEDED(FindByAutomationId(automation.get(), root.get(), id, element.put())) && element &&
                   SUCCEEDED(element->GetCurrentPatternAs(UIA_InvokePatternId, __uuidof(IUIAutomationInvokePattern), pattern.put_void())) && pattern &&
                   SUCCEEDED(pattern->Invoke());
        };
        const auto resize = [&](const wchar_t* id) noexcept
        {
            wil::com_ptr<IUIAutomationElement> element;
            wil::com_ptr<IUIAutomationRangeValuePattern> pattern;
            double before = 0.0, after = 0.0, minimum = 0.0, maximum = 0.0;
            if (FAILED(FindByAutomationId(automation.get(), root.get(), id, element.put())) || ! element ||
                FAILED(element->GetCurrentPatternAs(UIA_RangeValuePatternId, __uuidof(IUIAutomationRangeValuePattern), pattern.put_void())) || ! pattern ||
                FAILED(pattern->get_CurrentValue(&before)) || FAILED(pattern->get_CurrentMinimum(&minimum)) || FAILED(pattern->get_CurrentMaximum(&maximum)))
                return false;
            const double target = before + 16.0 <= maximum ? before + 16.0 : std::max(minimum, before - 16.0);
            if (target == before || FAILED(pattern->SetValue(target)) || FAILED(pattern->get_CurrentValue(&after)) || std::abs(after - target) > 1.0)
                return false;
            return SUCCEEDED(pattern->SetValue(before));
        };
        const auto togglePanel = [viewerWindow]() noexcept
        {
            SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_VIEW_SHOW_ANALYSIS, 0);
            return true;
        };
        if (! verifyPathAtWidth(wideWidthDip) || ! resize(L"ViewerSpaceAnalysis.PanelSplitter") || ! resize(L"ViewerSpaceAnalysis.InspectorSplitter"))
            return fail(L"ViewerSpace panel and inspector splitters resize through UIA RangeValue without activation");
        if (maxClientWidthDip >= 1100)
        {
            SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_VIEW_EXPLORER, 0);
            if (! verifyPathAtWidth((std::min)(1200, maxClientWidthDip)) || ! resize(L"ViewerSpaceAnalysis.PanelSplitter") ||
                ! resize(L"ViewerSpaceAnalysis.InspectorSplitter") || ! togglePanel() || IsWindowVisible(initial.analysisHwnd) != FALSE || ! togglePanel() ||
                IsWindowVisible(initial.analysisHwnd) == FALSE)
                return fail(L"ViewerSpace right panel and its inspector resize and collapse without losing the selected path");
            SendMessageW(viewerWindow, WM_COMMAND, IDM_VIEWERSPACE_VIEW_ANALYST, 0);
            if (! verifyPathAtWidth(wideWidthDip))
                return fail(L"ViewerSpace mode commands restore the bottom workspace and its selected path");
        }
        else
        {
            std::wcout << L"[SKIP] Right-panel resize requires at least 1100 DIP of owned monitor width.\n";
        }
        if (! invoke(L"ViewerSpaceAnalysis.InspectorToggle"))
            return fail(L"ViewerSpace inspector collapse is invokable");
        wil::com_ptr<IUIAutomationElement> hiddenPath;
        BOOL offscreen = FALSE;
        if (SUCCEEDED(FindByAutomationId(automation.get(), root.get(), L"ViewerSpaceAnalysis.Inspector.Path", hiddenPath.put())) && hiddenPath &&
            (FAILED(hiddenPath->get_CurrentIsOffscreen(&offscreen)) || ! offscreen))
            return fail(L"Collapsed inspector hides its fields from the visible UIA tree");
        if (! invoke(L"ViewerSpaceAnalysis.InspectorToggle") || ! verifyPathAtWidth(kNarrowWidthDip))
            return fail(L"Reopened inspector restores the selected full path");
        wil::com_ptr<IUIAutomationElement> panelToggle;
        wil::unique_bstr visiblePanelCaption;
        if (FAILED(FindByAutomationId(automation.get(), root.get(), L"ViewerSpace.Navigation.Layout", panelToggle.put())) || ! panelToggle ||
            FAILED(panelToggle->get_CurrentName(visiblePanelCaption.put())) || ! visiblePanelCaption)
            return fail(L"Panel toggle has an explicit accessible action label");
        if (! togglePanel() || IsWindowVisible(initial.analysisHwnd) != FALSE)
            return fail(L"Analysis disclosure collapses the entire bottom panel");
        wil::unique_bstr hiddenPanelCaption;
        if (FAILED(panelToggle->get_CurrentName(hiddenPanelCaption.put())) || ! hiddenPanelCaption ||
            std::wstring_view(hiddenPanelCaption.get()) == visiblePanelCaption.get())
            return fail(L"Hidden panel leaves a top-right toggle labeled with the inverse action");
        if (! togglePanel() || IsWindowVisible(initial.analysisHwnd) == FALSE || ! WaitForAnalysisBuild(viewerWindow, initial))
            return fail(L"Analysis disclosure restores the panel and derived rows");
        wil::unique_bstr restoredPanelCaption;
        if (FAILED(panelToggle->get_CurrentName(restoredPanelCaption.put())) || ! restoredPanelCaption ||
            std::wstring_view(restoredPanelCaption.get()) != visiblePanelCaption.get())
            return fail(L"Restored panel restores its collapse action label");
        if (! desktop.Sample())
            return fail(L"Panel resizing and collapse never take owned foreground");

        retainedGridCell = selectedCell;
        WndMsg::ViewerSpacePerfDebugSnapshot beforeDetach{};
        if (! ReadViewerSpaceSnapshot(viewerWindow, beforeDetach))
        {
            return fail(L"ViewerSpace debug snapshot is available before analysis-host detach");
        }
        cellRequestsBeforeDetach = beforeDetach.analysisCellRequests;
        if (! desktop.Finish())
        {
            return fail(L"ViewerSpace UIA desktop witness drains owned events");
        }
        logDesktop();
        return true;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        return fail(L"ViewerSpace analysis UIA test fails closed on an unexpected test exception");
    }
}

bool CheckViewerSpaceAnalysisDetached(HWND viewerWindow,
                                      IUIAutomationElement* retainedGridCell,
                                      uint64_t cellRequestsBeforeDetach,
                                      std::wstring_view& failure) noexcept
{
    failure         = {};
    const auto fail = [&failure](std::wstring_view reason) noexcept
    {
        failure = reason;
        return false;
    };

    try
    {
        if (! viewerWindow || IsWindow(viewerWindow) == FALSE || ! retainedGridCell)
        {
            return fail(L"ViewerSpace detached-provider test retains the outer window and a Grid cell provider");
        }
        WndMsg::ViewerSpacePerfDebugSnapshot beforeStaleRead{};
        if (! ReadViewerSpaceSnapshot(viewerWindow, beforeStaleRead) || beforeStaleRead.analysisCellRequests != cellRequestsBeforeDetach)
        {
            return fail(L"AnalysisHost detach leaves the source callback count unchanged before stale UIA access");
        }
        wil::unique_bstr staleName;
        const HRESULT staleHr = retainedGridCell->get_CurrentName(staleName.put());
        if (staleHr != UIA_E_ELEMENTNOTAVAILABLE)
        {
            return fail(L"Retained ViewerSpace Grid cell UIA provider reports unavailable after Detach");
        }
        WndMsg::ViewerSpacePerfDebugSnapshot afterStaleRead{};
        if (! ReadViewerSpaceSnapshot(viewerWindow, afterStaleRead) || afterStaleRead.analysisCellRequests != beforeStaleRead.analysisCellRequests)
        {
            return fail(L"Stale UIA provider access after Detach does not call the model");
        }
        return true;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        return fail(L"ViewerSpace detached-provider test fails closed on an unexpected test exception");
    }
}
