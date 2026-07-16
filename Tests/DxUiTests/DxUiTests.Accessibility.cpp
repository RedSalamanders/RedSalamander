#include "DxUiTestHelpers.h"

#include <atomic>
#include <chrono>
#include <fstream>
#include <iterator>
#include <numeric>
#include <thread>

namespace
{

void TestAccessibilityTargetPublishesImmutableSnapshotBeforeTreeTeardown()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for snapshot-publication guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    Require(source.find("struct AccessibilitySnapshot") != std::string::npos, "accessibility provider defines an immutable snapshot payload");
    Require(source.find("std::shared_ptr<const AccessibilitySnapshot>") != std::string::npos,
            "accessibility target/provider publish immutable snapshot instances");
    Require(source.find("CaptureAccessibilitySnapshot(target, hwnd)") != std::string::npos,
            "accessibility providers capture a snapshot instead of relying only on a live target");

    const size_t unregisterFunction = source.find("void UnregisterWindowHostAccessibilityTarget");
    const size_t returnProvider     = source.find("LRESULT ReturnWindowHostAccessibilityProvider", unregisterFunction);
    Require(unregisterFunction != std::string::npos && returnProvider != std::string::npos && unregisterFunction < returnProvider,
            "accessibility unregister source block is found");
    const std::string unregisterBlock = source.substr(unregisterFunction, returnProvider - unregisterFunction);

    const size_t publishEmpty = unregisterBlock.find("PublishEmptyAccessibilitySnapshot(*target)");
    const size_t clearHost    = unregisterBlock.find("target->host.store(nullptr");
    const size_t retireProviderMap = unregisterBlock.find("UiaReturnRawElementProvider(hwnd, 0, 0, nullptr)");
    Require(publishEmpty != std::string::npos, "accessibility unregister publishes an empty snapshot");
    Require(clearHost != std::string::npos && publishEmpty < clearHost,
            "accessibility unregister publishes the empty snapshot before clearing the live host pointer");
    Require(retireProviderMap != std::string::npos && clearHost < retireProviderMap,
            "accessibility unregister retires the HWND provider map after clearing the live host pointer");

    const std::filesystem::path windowHostSourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.WindowHost.cpp";
    std::ifstream windowHostInput(windowHostSourcePath);
    Require(windowHostInput.good(), "WindowHost source is readable for provider-destruction guard");
    const std::string windowHostSource((std::istreambuf_iterator<char>(windowHostInput)), std::istreambuf_iterator<char>());
    const size_t destroyCase          = windowHostSource.find("case WM_DESTROY:");
    const size_t destroyNotification  = windowHostSource.find("NotifyWindowHostAccessibilityDestroyed(hwnd)", destroyCase);
    const size_t nonClientDestroyCase = windowHostSource.find("case WM_NCDESTROY:", destroyCase);
    Require(destroyCase != std::string::npos && destroyNotification != std::string::npos && nonClientDestroyCase != std::string::npos &&
                destroyCase < destroyNotification && destroyNotification < nonClientDestroyCase,
            "WindowHost retires the UIA HWND provider map during WM_DESTROY before final non-client teardown");

    Require(source.find("focusedFragment") != std::string::npos, "accessibility snapshot captures focused-fragment identity");
    const size_t getFocusFunction = source.find("HRESULT AccessibilityProvider::GetFocus");
    const size_t invokeFunction   = source.find("HRESULT AccessibilityProvider::Invoke", getFocusFunction);
    Require(getFocusFunction != std::string::npos && invokeFunction != std::string::npos && getFocusFunction < invokeFunction,
            "accessibility GetFocus source block is found");
    const std::string getFocusBlock = source.substr(getFocusFunction, invokeFunction - getFocusFunction);
    Require(getFocusBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "accessibility GetFocus reads an immutable published snapshot");
    Require(getFocusBlock.find("ResolveHost(") == std::string::npos, "accessibility GetFocus does not resolve the live host");
    Require(getFocusBlock.find("ResolveRootControl(") == std::string::npos, "accessibility GetFocus does not resolve the live root");
    Require(getFocusBlock.find("GetFocusControl(") == std::string::npos, "accessibility GetFocus does not read live focus state");
    Require(getFocusBlock.find("FindPathForTarget(") == std::string::npos, "accessibility GetFocus does not walk the live tree");
}

void TestAccessibilityLiveHostResolutionIsWindowThreadOnly()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for live-host threading guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t resolveHostFunction = source.find("[[nodiscard]] WindowHost* ResolveHost() const noexcept");
    const size_t targetMutexFunction = source.find("[[nodiscard]] std::recursive_mutex& GetAccessibilityTargetMutex", resolveHostFunction);
    Require(resolveHostFunction != std::string::npos && targetMutexFunction != std::string::npos && resolveHostFunction < targetMutexFunction,
            "accessibility target ResolveHost source block is found");
    const std::string resolveHostBlock = source.substr(resolveHostFunction, targetMutexFunction - resolveHostFunction);

    Require(resolveHostBlock.find("GetWindowThreadProcessId(hwnd") != std::string::npos,
            "accessibility target ResolveHost checks the owning window thread before returning a live host");
    Require(resolveHostBlock.find("GetCurrentThreadId()") != std::string::npos,
            "accessibility target ResolveHost compares live access against the current thread");
    Require(resolveHostBlock.find("return nullptr;") != std::string::npos,
            "accessibility target ResolveHost refuses live host access when the thread/window check fails");
}

void TestAccessibilityProviderTraversalSurvivesConcurrentRootReplacement()
{
    using namespace RedSalamander::DxUi;

    struct StressRoot
    {
        StressRoot() = default;
        StressRoot(std::unique_ptr<Panel> rootValue, Control* focusValue) noexcept : root(std::move(rootValue)), focusTarget(focusValue)
        {
        }
        StressRoot(const StressRoot&)                = delete;
        StressRoot& operator=(const StressRoot&)     = delete;
        StressRoot(StressRoot&&) noexcept            = default;
        StressRoot& operator=(StressRoot&&) noexcept = default;

        std::unique_ptr<Panel> root;
        Control* focusTarget = nullptr;
    };

    auto buildRoot = [](MutableTreeModel& treeModel, MultiRowGridModel& gridModel)
    {
        auto root = std::make_unique<Panel>();

        auto* field = root->AddChild<TextField>(L"alpha beta gamma");
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 32.0f));

        auto* button = root->AddChild<Button>(L"Run");
        button->SetBounds(D2D1::RectF(188.0f, 0.0f, 260.0f, 32.0f));

        auto* tree = root->AddChild<Tree>();
        tree->SetBounds(D2D1::RectF(0.0f, 40.0f, 140.0f, 112.0f));
        tree->SetModel(&treeModel);
        tree->SetSelectedItemId(2u);

        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(148.0f, 40.0f, 300.0f, 132.0f));
        grid->SetModel(&gridModel);

        StressRoot result;
        result.focusTarget = field;
        result.root        = std::move(root);
        return result;
    };

    AttachedHostWindow window;
    MutableTreeModel treeModel;
    MultiRowGridModel gridModel(8u);
    treeModel.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Panes"},
        TreeItemData{.id = 3u, .text = L"Viewers"},
    });

    StressRoot firstRoot = buildRoot(treeModel, gridModel);
    Control* firstFocus  = firstRoot.focusTarget;
    window.Host().SetRoot(std::move(firstRoot.root));
    window.Host().SetFocusControl(firstFocus);

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "concurrent accessibility traversal creates a root provider");
    wil::com_ptr_nothrow<IRawElementProviderFragment> rootFragment;
    RequireSucceeded(rootProvider.query_to(rootFragment.put()), "concurrent accessibility traversal root supports fragment navigation");

    POINT hitPoint{24, 20};
    Require(ClientToScreen(window.Hwnd(), &hitPoint) != FALSE, "concurrent accessibility traversal computes a screen point");

    std::atomic<bool> stopWorker{false};
    std::atomic<HRESULT> workerFailure{S_OK};
    std::atomic<int> workerStep{0};
    std::atomic<uint32_t> traversalCount{0u};
    std::thread worker([&]
    {
        while (! stopWorker.load(std::memory_order_acquire))
        {
            wil::com_ptr_nothrow<IRawElementProviderFragment> focusedProvider;
            HRESULT hr = rootProvider->GetFocus(focusedProvider.put());
            if (FAILED(hr))
            {
                workerStep.store(1, std::memory_order_release);
                workerFailure.store(hr, std::memory_order_release);
                break;
            }

            wil::com_ptr_nothrow<IRawElementProviderFragment> hitProvider;
            hr = rootProvider->ElementProviderFromPoint(static_cast<double>(hitPoint.x), static_cast<double>(hitPoint.y), hitProvider.put());
            if (FAILED(hr))
            {
                workerStep.store(2, std::memory_order_release);
                workerFailure.store(hr, std::memory_order_release);
                break;
            }

            wil::com_ptr_nothrow<IRawElementProviderFragment> childProvider;
            hr = rootFragment->Navigate(NavigateDirection_FirstChild, childProvider.put());
            if (FAILED(hr))
            {
                workerStep.store(3, std::memory_order_release);
                workerFailure.store(hr, std::memory_order_release);
                break;
            }

            for (int depth = 0; childProvider && depth < 4; ++depth)
            {
                wil::com_ptr_nothrow<IRawElementProviderSimple> childSimple;
                hr = childProvider.query_to(childSimple.put());
                if (FAILED(hr))
                {
                    workerStep.store(4, std::memory_order_release);
                    workerFailure.store(hr, std::memory_order_release);
                    break;
                }

                VARIANT propertyValue;
                VariantInit(&propertyValue);
                hr = childSimple->GetPropertyValue(UIA_NamePropertyId, &propertyValue);
                VariantClear(&propertyValue);
                if (FAILED(hr))
                {
                    workerStep.store(5, std::memory_order_release);
                    workerFailure.store(hr, std::memory_order_release);
                    break;
                }

                wil::com_ptr_nothrow<IRawElementProviderFragment> nextProvider;
                hr = childProvider->Navigate(NavigateDirection_NextSibling, nextProvider.put());
                if (FAILED(hr))
                {
                    workerStep.store(6, std::memory_order_release);
                    workerFailure.store(hr, std::memory_order_release);
                    break;
                }

                childProvider = std::move(nextProvider);
            }

            if (FAILED(workerFailure.load(std::memory_order_acquire)))
            {
                break;
            }
            traversalCount.fetch_add(1u, std::memory_order_acq_rel);
        }
    });

    for (uint32_t iteration = 0u; iteration < 80u && SUCCEEDED(workerFailure.load(std::memory_order_acquire)); ++iteration)
    {
        treeModel.SetVisibleItems({
            TreeItemData{.id = 1u, .text = L"General"},
            TreeItemData{.id = 2u, .text = (iteration % 2u == 0u) ? L"Panes" : L"Layout"},
            TreeItemData{.id = 3u, .text = L"Viewers"},
            TreeItemData{.id = 4u, .text = L"Network"},
        });

        StressRoot replacement = buildRoot(treeModel, gridModel);
        Control* focusTarget   = replacement.focusTarget;
        window.Host().SetRoot(std::move(replacement.root));
        window.Host().SetFocusControl(focusTarget);
        window.PumpMessages();

        if ((iteration % 5u) == 0u)
        {
            window.Host().SetRoot(nullptr);
            window.PumpMessages();
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    stopWorker.store(true, std::memory_order_release);
    worker.join();

    RequireSucceeded(workerFailure.load(std::memory_order_acquire), "concurrent accessibility traversal survives root replacement");
    Require(workerStep.load(std::memory_order_acquire) == 0, "concurrent accessibility traversal reports no failed provider read step");
    Require(traversalCount.load(std::memory_order_acquire) > 0u, "concurrent accessibility traversal performs provider reads while roots churn");
}

void TestAccessibilityTreeHitTestUsesCheapVisibleIndexLookup()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Tree.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Tree source is readable for tree hit-test hot-path guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t helperFunction = source.find("std::optional<size_t> Tree::FindVisibleItemAtPoint");
    const size_t nextFunction   = source.find("float Tree::GetVerticalScrollableExtent", helperFunction);
    Require(helperFunction != std::string::npos && nextFunction != std::string::npos && helperFunction < nextFunction,
            "Tree visible-index point helper source block is found");

    const std::string helperBlock = source.substr(helperFunction, nextFunction - helperFunction);
    Require(helperBlock.find("offsetDip / _rowHeightDip") != std::string::npos, "Tree point lookup computes the visible index from row-height geometry");
    Require(helperBlock.find("GetItemLayoutMetrics(") == std::string::npos, "Tree point lookup does not materialize per-row layout metrics");
    Require(helperBlock.find("GetVisibleItem(") == std::string::npos, "Tree point lookup does not materialize TreeItemData while hit testing");
}

void TestAccessibilityProviderFactoriesUseSharedMakeProviderHelper()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for provider factory RAII guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    Require(source.find("MakeProvider(") != std::string::npos, "accessibility provider defines a shared provider factory helper");

    const size_t rootFactory = source.find("IRawElementProviderFragmentRoot* AccessibilityProvider::CreateRootProvider");
    const size_t pathHelper  = source.find("[[nodiscard]] bool FindAccessibilityPathForTarget", rootFactory);
    Require(rootFactory != std::string::npos && pathHelper != std::string::npos && rootFactory < pathHelper, "Accessibility provider factory block is found");
    const std::string factoryBlock = source.substr(rootFactory, pathHelper - rootFactory);
    Require(factoryBlock.find("new (std::nothrow)") == std::string::npos, "Accessibility provider factory methods do not duplicate nothrow allocation");
    Require(factoryBlock.find("target->Release()") == std::string::npos,
            "Accessibility provider factory methods do not duplicate target release-on-allocation-failure");
    Require(factoryBlock.find("MakeProvider<IRawElementProviderFragmentRoot, AccessibilityProvider>") != std::string::npos,
            "Accessibility root provider factory uses the shared helper");
    Require(factoryBlock.find("MakeProvider<ITextRangeProvider, AccessibilityTextRangeProvider>") != std::string::npos,
            "Accessibility text-range provider factory uses the shared helper");
}

void TestAccessibilityRuntimeIdsUseSharedBuilder()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for runtime-id RAII guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    Require(source.find("HRESULT BuildRuntimeId(SAFEARRAY** outArray, std::span<const LONG> values)") != std::string::npos,
            "Accessibility runtime-id code defines a shared SAFEARRAY builder");
    Require(source.find("AppendControlPathRuntimeIdPrefix") != std::string::npos, "Accessibility runtime-id code defines one control-path prefix helper");
    Require(source.find("kAccessibilityRuntimeIdTreeItem") != std::string::npos, "Tree item runtime-id discriminator is named");
    Require(source.find("kAccessibilityRuntimeIdGridRow") != std::string::npos, "Grid row runtime-id discriminator is named");
    Require(source.find("kAccessibilityRuntimeIdGridCell") != std::string::npos, "Grid cell runtime-id discriminator is named");
    Require(source.find("kAccessibilityRuntimeIdGridHeader") != std::string::npos, "Grid header runtime-id discriminator is named");
    Require(source.find("kAccessibilityRuntimeIdPasswordRevealButton") != std::string::npos, "Password reveal runtime-id discriminator is named");

    const size_t runtimeIdBlockStart = source.find("HRESULT BuildRuntimeId");
    const size_t providerArrayStart  = source.find("HRESULT SetProviderArray", runtimeIdBlockStart);
    Require(runtimeIdBlockStart != std::string::npos && providerArrayStart != std::string::npos && runtimeIdBlockStart < providerArrayStart,
            "Accessibility runtime-id helper block is found");
    const std::string runtimeIdBlock = source.substr(runtimeIdBlockStart, providerArrayStart - runtimeIdBlockStart);

    const size_t firstCreate = runtimeIdBlock.find("SafeArrayCreateVector(VT_I4");
    Require(firstCreate != std::string::npos, "Accessibility runtime-id helper owns VT_I4 SAFEARRAY creation");
    Require(runtimeIdBlock.find("SafeArrayCreateVector(VT_I4", firstCreate + 1u) == std::string::npos,
            "Accessibility runtime-id wrappers do not duplicate VT_I4 SAFEARRAY creation");

    const size_t firstPut = runtimeIdBlock.find("SafeArrayPutElement(");
    Require(firstPut != std::string::npos, "Accessibility runtime-id helper owns SAFEARRAY population");
    Require(runtimeIdBlock.find("SafeArrayPutElement(", firstPut + 1u) == std::string::npos,
            "Accessibility runtime-id wrappers do not duplicate SAFEARRAY population loops");
}

void TestAccessibilityPatternDispatchUsesSharedQueryPattern()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for pattern-dispatch guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    Require(source.find("enum class AccessibilityPatternKind") != std::string::npos, "Accessibility pattern dispatch names one internal pattern kind enum");
    Require(source.find("AccessibilityPatternQueryResult AccessibilityProvider::QueryPattern") != std::string::npos,
            "Accessibility provider defines one shared QueryPattern implementation");

    const size_t queryFunction   = source.find("HRESULT AccessibilityProvider::QueryInterface");
    const size_t optionsFunction = source.find("HRESULT AccessibilityProvider::get_ProviderOptions", queryFunction);
    Require(queryFunction != std::string::npos && optionsFunction != std::string::npos && queryFunction < optionsFunction,
            "Accessibility QueryInterface source block is found for pattern-dispatch guard");
    const std::string queryBlock = source.substr(queryFunction, optionsFunction - queryFunction);
    Require(queryBlock.find("PatternKindFromInterfaceId(riid)") != std::string::npos,
            "Accessibility QueryInterface maps interface ids into the shared pattern dispatch");
    Require(queryBlock.find("QueryPattern(patternKind.value())") != std::string::npos,
            "Accessibility QueryInterface uses the shared QueryPattern implementation");
    Require(queryBlock.find("SupportsInvokePattern(") == std::string::npos, "Accessibility QueryInterface no longer owns invoke-pattern eligibility");
    Require(queryBlock.find("SnapshotGridCellSupports") == std::string::npos, "Accessibility QueryInterface no longer owns grid-cell pattern eligibility");

    const size_t patternFunction  = source.find("HRESULT AccessibilityProvider::GetPatternProvider");
    const size_t propertyFunction = source.find("HRESULT AccessibilityProvider::GetPropertyValue", patternFunction);
    Require(patternFunction != std::string::npos && propertyFunction != std::string::npos && patternFunction < propertyFunction,
            "Accessibility GetPatternProvider source block is found for pattern-dispatch guard");
    const std::string patternBlock = source.substr(patternFunction, propertyFunction - patternFunction);
    Require(patternBlock.find("PatternKindFromPatternId(patternId)") != std::string::npos,
            "Accessibility GetPatternProvider maps pattern ids into the shared pattern dispatch");
    Require(patternBlock.find("QueryPattern(patternKind.value())") != std::string::npos,
            "Accessibility GetPatternProvider uses the shared QueryPattern implementation");
    Require(patternBlock.find("SupportsInvokePattern(") == std::string::npos, "Accessibility GetPatternProvider no longer owns invoke-pattern eligibility");
    Require(patternBlock.find("SnapshotGridCellSupports") == std::string::npos,
            "Accessibility GetPatternProvider no longer owns grid-cell pattern eligibility");
}

void TestAccessibilityElementProviderFromPointUsesSnapshot()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for point-provider snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t pointFunction = source.find("HRESULT AccessibilityProvider::ElementProviderFromPoint");
    const size_t focusFunction = source.find("HRESULT AccessibilityProvider::GetFocus", pointFunction);
    Require(pointFunction != std::string::npos && focusFunction != std::string::npos && pointFunction < focusFunction,
            "Accessibility ElementProviderFromPoint source block is found");

    const std::string pointBlock = source.substr(pointFunction, focusFunction - pointFunction);
    Require(pointBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility ElementProviderFromPoint reads an immutable published snapshot");
    Require(pointBlock.find("FindSnapshotPointHit") != std::string::npos,
            "Accessibility ElementProviderFromPoint resolves providers from snapshot hit records");
    Require(pointBlock.find("ResolveHost(") == std::string::npos, "Accessibility ElementProviderFromPoint does not resolve the live host");
    Require(pointBlock.find("ResolveRootControl(") == std::string::npos, "Accessibility ElementProviderFromPoint does not resolve the live root");
    Require(pointBlock.find("ResolveControlAtPath(") == std::string::npos, "Accessibility ElementProviderFromPoint does not walk live controls");
    Require(pointBlock.find("FindSemanticControlAtPoint(") == std::string::npos, "Accessibility ElementProviderFromPoint does not live-hit-test controls");
    Require(pointBlock.find("GetModel(") == std::string::npos, "Accessibility ElementProviderFromPoint does not read Tree/Grid models");
}

void TestAccessibilityBoundingRectangleUsesSnapshot()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for bounding-rectangle snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t boundsFunction   = source.find("HRESULT AccessibilityProvider::get_BoundingRectangle");
    const size_t embeddedFunction = source.find("HRESULT AccessibilityProvider::GetEmbeddedFragmentRoots", boundsFunction);
    Require(boundsFunction != std::string::npos && embeddedFunction != std::string::npos && boundsFunction < embeddedFunction,
            "Accessibility get_BoundingRectangle source block is found");

    const std::string boundsBlock = source.substr(boundsFunction, embeddedFunction - boundsFunction);
    Require(boundsBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility get_BoundingRectangle reads an immutable published snapshot");
    Require(boundsBlock.find("FindSnapshotFragmentBounds") != std::string::npos,
            "Accessibility get_BoundingRectangle resolves fragment bounds from snapshot records");
    Require(boundsBlock.find("ResolveHost(") == std::string::npos, "Accessibility get_BoundingRectangle does not resolve the live host");
    Require(boundsBlock.find("ResolveControl(") == std::string::npos, "Accessibility get_BoundingRectangle does not resolve live controls");
    Require(boundsBlock.find("ResolveTreeControl(") == std::string::npos, "Accessibility get_BoundingRectangle does not resolve live trees");
    Require(boundsBlock.find("ResolveGridControl(") == std::string::npos, "Accessibility get_BoundingRectangle does not resolve live grids");
    Require(boundsBlock.find("ResolveGridCellData(") == std::string::npos, "Accessibility get_BoundingRectangle does not read Grid cell data");
    Require(boundsBlock.find("ResolveTreeItemData(") == std::string::npos, "Accessibility get_BoundingRectangle does not read Tree item data");
    Require(boundsBlock.find("GetItemLayoutMetrics(") == std::string::npos, "Accessibility get_BoundingRectangle does not build Tree layout metrics");
    Require(boundsBlock.find("GetModel(") == std::string::npos, "Accessibility get_BoundingRectangle does not read Tree/Grid models");
}

void TestAccessibilityGridPatternReadsUseSnapshots()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for grid-pattern snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t tableHeadersFunction = source.find("HRESULT AccessibilityProvider::GetColumnHeaders");
    const size_t rowOrColumnFunction  = source.find("HRESULT AccessibilityProvider::get_RowOrColumnMajor", tableHeadersFunction);
    Require(tableHeadersFunction != std::string::npos && rowOrColumnFunction != std::string::npos && tableHeadersFunction < rowOrColumnFunction,
            "Accessibility table column-header source block is found");
    const std::string tableHeadersBlock = source.substr(tableHeadersFunction, rowOrColumnFunction - tableHeadersFunction);
    Require(tableHeadersBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetColumnHeaders reads an immutable published snapshot");
    Require(tableHeadersBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility GetColumnHeaders reads visible grid columns from snapshot records");
    Require(tableHeadersBlock.find("ResolveControlPath(") == std::string::npos, "Accessibility GetColumnHeaders does not resolve live control paths");
    Require(tableHeadersBlock.find("ResolveGridControl(") == std::string::npos, "Accessibility GetColumnHeaders does not resolve live grids");
    Require(tableHeadersBlock.find("GetVisibleColumnAt(") == std::string::npos, "Accessibility GetColumnHeaders does not read live visible columns");

    const size_t rowHeadersFunction    = source.find("HRESULT AccessibilityProvider::GetRowHeaders");
    const size_t columnHeadersFunction = source.find("HRESULT AccessibilityProvider::GetColumnHeaders", rowHeadersFunction);
    Require(rowHeadersFunction != std::string::npos && columnHeadersFunction != std::string::npos && rowHeadersFunction < columnHeadersFunction,
            "Accessibility table row-header source block is found");
    const std::string rowHeadersBlock = source.substr(rowHeadersFunction, columnHeadersFunction - rowHeadersFunction);
    Require(rowHeadersBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetRowHeaders reads an immutable published snapshot");
    Require(rowHeadersBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility GetRowHeaders validates the grid from snapshot records");
    Require(rowHeadersBlock.find("ResolveControlPath(") == std::string::npos, "Accessibility GetRowHeaders does not resolve live control paths");
    Require(rowHeadersBlock.find("SupportsGridTablePattern(") == std::string::npos, "Accessibility GetRowHeaders does not re-resolve live table support");

    const size_t rowFunction        = source.find("HRESULT AccessibilityProvider::get_Row(int*");
    const size_t containingFunction = source.find("HRESULT AccessibilityProvider::get_ContainingGrid", rowFunction);
    Require(rowFunction != std::string::npos && containingFunction != std::string::npos && rowFunction < containingFunction,
            "Accessibility grid item row/column source block is found");
    const std::string gridItemBlock = source.substr(rowFunction, containingFunction - rowFunction);
    Require(gridItemBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility grid item row/column reads immutable published snapshots");
    Require(gridItemBlock.find("FindSnapshotGridCellRecord") != std::string::npos,
            "Accessibility grid item row/column resolves cell metadata from snapshot records");
    Require(gridItemBlock.find("ResolveGridCellData(") == std::string::npos, "Accessibility grid item row/column does not read live Grid cell data");
    Require(gridItemBlock.find("SupportsGridCellPattern(") == std::string::npos,
            "Accessibility grid item row/column does not re-resolve live Grid pattern support");

    const size_t containingGridFunction = source.find("HRESULT AccessibilityProvider::get_ContainingGrid");
    const size_t nextFunction           = source.find("HRESULT AccessibilityProvider::GetRowHeaderItems", containingGridFunction);
    Require(containingGridFunction != std::string::npos && nextFunction != std::string::npos && containingGridFunction < nextFunction,
            "Accessibility containing-grid source block is found");
    const std::string containingGridBlock = source.substr(containingGridFunction, nextFunction - containingGridFunction);
    Require(containingGridBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility containing-grid reads an immutable published snapshot");
    Require(containingGridBlock.find("FindSnapshotGridCellRecord") != std::string::npos,
            "Accessibility containing-grid validates the cell from snapshot records");
    Require(containingGridBlock.find("SupportsGridCellPattern(") == std::string::npos,
            "Accessibility containing-grid does not re-resolve live Grid pattern support");

    const size_t tableItemFunction = source.find("HRESULT AccessibilityProvider::GetColumnHeaderItems");
    const size_t helperFunction    = source.find("WindowHost* AccessibilityProvider::ResolveHost", tableItemFunction);
    Require(tableItemFunction != std::string::npos && helperFunction != std::string::npos && tableItemFunction < helperFunction,
            "Accessibility table-item column-header source block is found");
    const std::string tableItemBlock = source.substr(tableItemFunction, helperFunction - tableItemFunction);
    Require(tableItemBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility table item header reads immutable published snapshots");
    Require(tableItemBlock.find("FindSnapshotGridCellRecord") != std::string::npos,
            "Accessibility table item header resolves cell metadata from snapshot records");
    Require(tableItemBlock.find("ResolveGridCellData(") == std::string::npos, "Accessibility table item header does not read live Grid cell data");
    Require(tableItemBlock.find("SupportsGridCellTableItemPattern(") == std::string::npos,
            "Accessibility table item header does not re-resolve live table-item support");

    const size_t propertyValueFunction = source.find("HRESULT AccessibilityProvider::GetPropertyValue");
    const size_t hostProviderFunction  = source.find("HRESULT AccessibilityProvider::get_HostRawElementProvider", propertyValueFunction);
    Require(propertyValueFunction != std::string::npos && hostProviderFunction != std::string::npos && propertyValueFunction < hostProviderFunction,
            "Accessibility GetPropertyValue source block is found for grid count properties");
    const std::string propertyValueBlock = source.substr(propertyValueFunction, hostProviderFunction - propertyValueFunction);
    const size_t gridRowCountProperty    = propertyValueBlock.find("UIA_GridRowCountPropertyId");
    const size_t propertyLiveRootResolve = propertyValueBlock.find("ResolveRootControl(");
    Require(gridRowCountProperty != std::string::npos && (propertyLiveRootResolve == std::string::npos || gridRowCountProperty < propertyLiveRootResolve),
            "Accessibility GetPropertyValue handles Grid row/column counts before any live root resolve");
    const std::string gridCountPropertyBlock = propertyValueBlock.substr(0u, propertyLiveRootResolve);
    Require(gridCountPropertyBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetPropertyValue Grid counts read an immutable published snapshot");
    Require(gridCountPropertyBlock.find("gridRowCount") != std::string::npos, "Accessibility GetPropertyValue Grid row count reads snapshot records");
    Require(gridCountPropertyBlock.find("gridColumnCount") != std::string::npos, "Accessibility GetPropertyValue Grid column count reads snapshot records");
    Require(gridCountPropertyBlock.find("GetModel(") == std::string::npos, "Accessibility GetPropertyValue Grid counts do not read live Grid models");

    const size_t gridHeaderBranch = propertyValueBlock.find("if (_kind == AccessibilityFragmentKind::GridHeader)");
    Require(gridHeaderBranch != std::string::npos && (propertyLiveRootResolve == std::string::npos || gridHeaderBranch < propertyLiveRootResolve),
            "Accessibility GetPropertyValue handles GridHeader state before any live root resolve");
    const std::string gridHeaderPropertyBlock = propertyValueBlock.substr(gridHeaderBranch, propertyLiveRootResolve - gridHeaderBranch);
    Require(gridHeaderPropertyBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetPropertyValue GridHeader state reads an immutable published snapshot");
    Require(gridHeaderPropertyBlock.find("FindSnapshotGridHeaderRecord") != std::string::npos,
            "Accessibility GetPropertyValue GridHeader state resolves metadata from snapshot records");
    Require(gridHeaderPropertyBlock.find("gridHeaderName") != std::string::npos, "Accessibility GetPropertyValue GridHeader name reads snapshot records");
    Require(gridHeaderPropertyBlock.find("ResolveGridControl(") == std::string::npos,
            "Accessibility GetPropertyValue GridHeader state does not resolve live grids");
    Require(gridHeaderPropertyBlock.find("ResolveGridHeaderColumn(") == std::string::npos,
            "Accessibility GetPropertyValue GridHeader state does not read live Grid columns");

    const std::string legacyFallbackBlock = propertyLiveRootResolve == std::string::npos ? std::string{} : propertyValueBlock.substr(propertyLiveRootResolve);
    Require(legacyFallbackBlock.find("AccessibilityFragmentKind::TreeItem") == std::string::npos,
            "Accessibility GetPropertyValue has no legacy live TreeItem property fallback after snapshot handling");
    Require(legacyFallbackBlock.find("AccessibilityFragmentKind::GridHeader") == std::string::npos,
            "Accessibility GetPropertyValue has no legacy live GridHeader property fallback after snapshot handling");
    Require(legacyFallbackBlock.find("UIA_GridRowCountPropertyId") == std::string::npos,
            "Accessibility GetPropertyValue has no legacy live Grid row-count fallback after snapshot handling");
    Require(legacyFallbackBlock.find("UIA_GridColumnCountPropertyId") == std::string::npos,
            "Accessibility GetPropertyValue has no legacy live Grid column-count fallback after snapshot handling");
}

void TestAccessibilityControlStateReadsUseSnapshots()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for control snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t queryPatternFunction = source.find("AccessibilityPatternQueryResult AccessibilityProvider::QueryPattern");
    const size_t queryFunction        = source.find("HRESULT AccessibilityProvider::QueryInterface", queryPatternFunction);
    Require(queryPatternFunction != std::string::npos && queryFunction != std::string::npos && queryPatternFunction < queryFunction,
            "Accessibility shared QueryPattern source block is found for control snapshot guard");
    const std::string queryPatternBlock = source.substr(queryPatternFunction, queryFunction - queryPatternFunction);
    const size_t queryLiveRoot          = queryPatternBlock.find("ResolveRootControl(");
    const size_t querySnapshotRecord    = queryPatternBlock.find("ResolveSnapshotControlRecord");
    Require(querySnapshotRecord != std::string::npos && (queryLiveRoot == std::string::npos || querySnapshotRecord < queryLiveRoot),
            "Accessibility shared QueryPattern resolves ordinary control pattern support from snapshots before any live root resolve");
    const std::string queryLegacyBlock = queryLiveRoot == std::string::npos ? std::string{} : queryPatternBlock.substr(queryLiveRoot);
    Require(queryLegacyBlock.find("SupportsValuePattern(control)") == std::string::npos,
            "Accessibility shared QueryPattern has no legacy live ValuePattern support fallback");
    Require(queryLegacyBlock.find("SupportsTextPattern(control)") == std::string::npos,
            "Accessibility shared QueryPattern has no legacy live TextPattern support fallback");
    Require(queryLegacyBlock.find("SupportsRangeValuePattern(control)") == std::string::npos,
            "Accessibility shared QueryPattern has no legacy live RangeValue support fallback");

    const size_t propertyValueFunction = source.find("HRESULT AccessibilityProvider::GetPropertyValue");
    const size_t hostProviderFunction  = source.find("HRESULT AccessibilityProvider::get_HostRawElementProvider", propertyValueFunction);
    Require(propertyValueFunction != std::string::npos && hostProviderFunction != std::string::npos && propertyValueFunction < hostProviderFunction,
            "Accessibility GetPropertyValue source block is found for control snapshot guard");
    const std::string propertyValueBlock = source.substr(propertyValueFunction, hostProviderFunction - propertyValueFunction);
    const size_t propertyLiveRoot        = propertyValueBlock.find("ResolveRootControl(");
    const size_t propertySnapshotRecord  = propertyValueBlock.find("ResolveSnapshotControlRecord");
    Require(propertySnapshotRecord != std::string::npos && (propertyLiveRoot == std::string::npos || propertySnapshotRecord < propertyLiveRoot),
            "Accessibility GetPropertyValue resolves ordinary control properties from snapshots before any live root resolve");
    const std::string propertyLegacyBlock = propertyLiveRoot == std::string::npos ? std::string{} : propertyValueBlock.substr(propertyLiveRoot);
    Require(propertyLegacyBlock.find("GetControlAccessibleName(") == std::string::npos,
            "Accessibility GetPropertyValue has no legacy live accessible-name fallback");
    Require(propertyLegacyBlock.find("GetControlAccessibleValue(") == std::string::npos,
            "Accessibility GetPropertyValue has no legacy live text value fallback");
    Require(propertyLegacyBlock.find("dynamic_cast<const TextField*>") == std::string::npos,
            "Accessibility GetPropertyValue has no legacy live TextField password-state fallback");
    Require(propertyLegacyBlock.find("IsValueReadOnly(control)") == std::string::npos, "Accessibility GetPropertyValue has no legacy live read-only fallback");

    const size_t stringValueFunction = source.find("HRESULT AccessibilityProvider::get_Value(BSTR*");
    const size_t doubleValueFunction = source.find("HRESULT AccessibilityProvider::get_Value(double*", stringValueFunction);
    Require(stringValueFunction != std::string::npos && doubleValueFunction != std::string::npos && stringValueFunction < doubleValueFunction,
            "Accessibility string ValuePattern source block is found");
    const std::string stringValueBlock = source.substr(stringValueFunction, doubleValueFunction - stringValueFunction);
    Require(stringValueBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility string ValuePattern reads ordinary control values from snapshots");
    Require(stringValueBlock.find("GetControlAccessibleValue(control)") == std::string::npos,
            "Accessibility string ValuePattern has no legacy live text value fallback");

    const size_t readOnlyFunction = source.find("HRESULT AccessibilityProvider::get_IsReadOnly");
    const size_t maximumFunction  = source.find("HRESULT AccessibilityProvider::get_Maximum", readOnlyFunction);
    Require(readOnlyFunction != std::string::npos && maximumFunction != std::string::npos && readOnlyFunction < maximumFunction,
            "Accessibility ValuePattern read-only source block is found");
    const std::string readOnlyBlock = source.substr(readOnlyFunction, maximumFunction - readOnlyFunction);
    Require(readOnlyBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility ValuePattern read-only state reads ordinary control state from snapshots");
    Require(readOnlyBlock.find("IsValueReadOnly(control)") == std::string::npos,
            "Accessibility ValuePattern read-only state has no legacy live control fallback");
}

void TestAccessibilityGridCellValueReadsUseSnapshots()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for grid-cell snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t queryPatternFunction = source.find("AccessibilityPatternQueryResult AccessibilityProvider::QueryPattern");
    const size_t queryFunction        = source.find("HRESULT AccessibilityProvider::QueryInterface", queryPatternFunction);
    Require(queryPatternFunction != std::string::npos && queryFunction != std::string::npos && queryPatternFunction < queryFunction,
            "Accessibility shared QueryPattern source block is found for grid-cell guard");
    const std::string queryPatternBlock = source.substr(queryPatternFunction, queryFunction - queryPatternFunction);
    const size_t queryGridCellBranch    = queryPatternBlock.find("if (_kind == AccessibilityFragmentKind::GridCell)");
    const size_t queryLiveRootResolve   = queryPatternBlock.find("ResolveRootControl(");
    Require(queryGridCellBranch != std::string::npos && (queryLiveRootResolve == std::string::npos || queryGridCellBranch < queryLiveRootResolve),
            "Accessibility shared QueryPattern handles GridCell patterns before any live root resolve");
    const std::string queryGridCellBlock = queryPatternBlock.substr(
        queryGridCellBranch, queryLiveRootResolve == std::string::npos ? std::string::npos : queryLiveRootResolve - queryGridCellBranch);
    Require(queryGridCellBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility shared QueryPattern GridCell patterns read an immutable published snapshot");
    Require(queryGridCellBlock.find("FindSnapshotGridCellRecord") != std::string::npos,
            "Accessibility shared QueryPattern GridCell patterns validate snapshot cell records");
    Require(queryGridCellBlock.find("SnapshotGridCellSupportsTogglePattern") != std::string::npos,
            "Accessibility shared QueryPattern GridCell toggle pattern support comes from snapshot records");
    Require(queryGridCellBlock.find("SnapshotGridCellSupportsValuePattern") != std::string::npos,
            "Accessibility shared QueryPattern GridCell value pattern support comes from snapshot records");
    Require(queryGridCellBlock.find("SnapshotGridCellSupportsRangeValuePattern") != std::string::npos,
            "Accessibility shared QueryPattern GridCell range pattern support comes from snapshot records");
    Require(queryGridCellBlock.find("SupportsGridCell") == std::string::npos,
            "Accessibility shared QueryPattern GridCell patterns do not re-resolve live Grid cell support");
    Require(queryGridCellBlock.find("ResolveGridCellData(") == std::string::npos,
            "Accessibility shared QueryPattern GridCell patterns do not read live Grid cell data");

    const size_t propertyValueFunction = source.find("HRESULT AccessibilityProvider::GetPropertyValue");
    const size_t hostProviderFunction  = source.find("HRESULT AccessibilityProvider::get_HostRawElementProvider", propertyValueFunction);
    Require(propertyValueFunction != std::string::npos && hostProviderFunction != std::string::npos && propertyValueFunction < hostProviderFunction,
            "Accessibility GetPropertyValue source block is found for grid-cell guard");
    const std::string propertyValueBlock = source.substr(propertyValueFunction, hostProviderFunction - propertyValueFunction);
    const size_t propertyGridCellBranch  = propertyValueBlock.find("if (_kind == AccessibilityFragmentKind::GridCell)");
    const size_t propertyLiveRootResolve = propertyValueBlock.find("ResolveRootControl(");
    Require(propertyGridCellBranch != std::string::npos && (propertyLiveRootResolve == std::string::npos || propertyGridCellBranch < propertyLiveRootResolve),
            "Accessibility GetPropertyValue handles GridCell state before any live root resolve");
    const std::string propertyGridCellBlock = propertyValueBlock.substr(
        propertyGridCellBranch, propertyLiveRootResolve == std::string::npos ? std::string::npos : propertyLiveRootResolve - propertyGridCellBranch);
    Require(propertyGridCellBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetPropertyValue GridCell state reads an immutable published snapshot");
    Require(propertyGridCellBlock.find("FindSnapshotGridCellRecord") != std::string::npos,
            "Accessibility GetPropertyValue GridCell state resolves metadata from snapshot records");
    Require(propertyGridCellBlock.find("gridCellAccessibleText") != std::string::npos, "Accessibility GetPropertyValue GridCell text reads snapshot records");
    Require(propertyGridCellBlock.find("SnapshotGridCellSupportsTogglePattern") != std::string::npos,
            "Accessibility GetPropertyValue GridCell toggle support comes from snapshot records");
    Require(propertyGridCellBlock.find("SnapshotGridCellSupportsValuePattern") != std::string::npos,
            "Accessibility GetPropertyValue GridCell value support comes from snapshot records");
    Require(propertyGridCellBlock.find("SnapshotGridCellSupportsRangeValuePattern") != std::string::npos,
            "Accessibility GetPropertyValue GridCell range support comes from snapshot records");
    Require(propertyGridCellBlock.find("ResolveGridControl(") == std::string::npos,
            "Accessibility GetPropertyValue GridCell state does not resolve live grids");
    Require(propertyGridCellBlock.find("ResolveGridCellData(") == std::string::npos,
            "Accessibility GetPropertyValue GridCell state does not read live Grid cell data");
    Require(propertyGridCellBlock.find("SupportsGridCell") == std::string::npos,
            "Accessibility GetPropertyValue GridCell state does not re-resolve live Grid cell support");

    const size_t toggleFunction   = source.find("HRESULT AccessibilityProvider::get_ToggleState");
    const size_t setValueFunction = source.find("HRESULT AccessibilityProvider::SetValue(LPCWSTR", toggleFunction);
    Require(toggleFunction != std::string::npos && setValueFunction != std::string::npos && toggleFunction < setValueFunction,
            "Accessibility get_ToggleState source block is found");
    const std::string toggleBlock = source.substr(toggleFunction, setValueFunction - toggleFunction);
    Require(toggleBlock.find("FindSnapshotGridCellRecord") != std::string::npos, "Accessibility get_ToggleState GridCell read uses snapshot cell records");
    Require(toggleBlock.find("SnapshotGridCellSupportsTogglePattern") != std::string::npos,
            "Accessibility get_ToggleState GridCell support comes from snapshot records");
    Require(toggleBlock.find("ResolveGridCellData(") == std::string::npos, "Accessibility get_ToggleState does not read live Grid cell data");

    const size_t stringValueFunction = source.find("HRESULT AccessibilityProvider::get_Value(BSTR* outValue)");
    const size_t rangeValueFunction  = source.find("HRESULT AccessibilityProvider::get_Value(double* outValue)", stringValueFunction);
    Require(stringValueFunction != std::string::npos && rangeValueFunction != std::string::npos && stringValueFunction < rangeValueFunction,
            "Accessibility string get_Value source block is found");
    const std::string stringValueBlock = source.substr(stringValueFunction, rangeValueFunction - stringValueFunction);
    Require(stringValueBlock.find("FindSnapshotGridCellRecord") != std::string::npos,
            "Accessibility string get_Value GridCell read uses snapshot cell records");
    Require(stringValueBlock.find("SnapshotGridCellSupportsValuePattern") != std::string::npos,
            "Accessibility string get_Value GridCell support comes from snapshot records");
    Require(stringValueBlock.find("ResolveGridCellData(") == std::string::npos, "Accessibility string get_Value does not read live Grid cell data");

    const size_t isReadOnlyFunction = source.find("HRESULT AccessibilityProvider::get_IsReadOnly", rangeValueFunction);
    Require(rangeValueFunction != std::string::npos && isReadOnlyFunction != std::string::npos && rangeValueFunction < isReadOnlyFunction,
            "Accessibility range get_Value source block is found");
    const std::string rangeValueBlock = source.substr(rangeValueFunction, isReadOnlyFunction - rangeValueFunction);
    Require(rangeValueBlock.find("FindSnapshotGridCellRecord") != std::string::npos, "Accessibility range get_Value GridCell read uses snapshot cell records");
    Require(rangeValueBlock.find("SnapshotGridCellSupportsRangeValuePattern") != std::string::npos,
            "Accessibility range get_Value GridCell support comes from snapshot records");
    Require(rangeValueBlock.find("ResolveGridCellData(") == std::string::npos, "Accessibility range get_Value does not read live Grid cell data");

    const size_t maximumFunction = source.find("HRESULT AccessibilityProvider::get_Maximum", isReadOnlyFunction);
    Require(isReadOnlyFunction != std::string::npos && maximumFunction != std::string::npos && isReadOnlyFunction < maximumFunction,
            "Accessibility get_IsReadOnly source block is found");
    const std::string readOnlyBlock = source.substr(isReadOnlyFunction, maximumFunction - isReadOnlyFunction);
    Require(readOnlyBlock.find("FindSnapshotGridCellRecord") != std::string::npos, "Accessibility get_IsReadOnly GridCell read uses snapshot cell records");
    Require(readOnlyBlock.find("SnapshotGridCellSupportsValuePattern") != std::string::npos,
            "Accessibility get_IsReadOnly GridCell value support comes from snapshot records");
    Require(readOnlyBlock.find("SnapshotGridCellSupportsRangeValuePattern") != std::string::npos,
            "Accessibility get_IsReadOnly GridCell range support comes from snapshot records");
    Require(readOnlyBlock.find("ResolveGridCellData(") == std::string::npos, "Accessibility get_IsReadOnly does not read live Grid cell data");

    const size_t selectionFunction = source.find("HRESULT AccessibilityProvider::GetSelection", maximumFunction);
    Require(maximumFunction != std::string::npos && selectionFunction != std::string::npos && maximumFunction < selectionFunction,
            "Accessibility range metadata source block is found");
    const std::string rangeMetadataBlock = source.substr(maximumFunction, selectionFunction - maximumFunction);
    Require(rangeMetadataBlock.find("FindSnapshotGridCellRecord") != std::string::npos,
            "Accessibility range metadata GridCell reads use snapshot cell records");
    Require(rangeMetadataBlock.find("SnapshotGridCellSupportsRangeValuePattern") != std::string::npos,
            "Accessibility range metadata GridCell support comes from snapshot records");
    Require(rangeMetadataBlock.find("SupportsGridCellRangeValuePattern()") == std::string::npos,
            "Accessibility range metadata does not re-resolve live Grid cell range support");
    Require(rangeMetadataBlock.find("ResolveGridCellData(") == std::string::npos, "Accessibility range metadata does not read live Grid cell data");
}

void TestAccessibilityGridRowPropertyReadsUseSnapshots()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for grid-row snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t propertyValueFunction = source.find("HRESULT AccessibilityProvider::GetPropertyValue");
    const size_t hostProviderFunction  = source.find("HRESULT AccessibilityProvider::get_HostRawElementProvider", propertyValueFunction);
    Require(propertyValueFunction != std::string::npos && hostProviderFunction != std::string::npos && propertyValueFunction < hostProviderFunction,
            "Accessibility GetPropertyValue source block is found for grid-row guard");
    const std::string propertyValueBlock = source.substr(propertyValueFunction, hostProviderFunction - propertyValueFunction);

    const size_t propertyGridRowBranch   = propertyValueBlock.find("if (_kind == AccessibilityFragmentKind::GridRow)");
    const size_t propertyLiveRootResolve = propertyValueBlock.find("ResolveRootControl(");
    Require(propertyGridRowBranch != std::string::npos && (propertyLiveRootResolve == std::string::npos || propertyGridRowBranch < propertyLiveRootResolve),
            "Accessibility GetPropertyValue handles GridRow state before any live root resolve");
    const std::string propertyGridRowBlock = propertyValueBlock.substr(
        propertyGridRowBranch, propertyLiveRootResolve == std::string::npos ? std::string::npos : propertyLiveRootResolve - propertyGridRowBranch);
    Require(propertyGridRowBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetPropertyValue GridRow state reads an immutable published snapshot");
    Require(propertyGridRowBlock.find("FindSnapshotGridRowRecord") != std::string::npos,
            "Accessibility GetPropertyValue GridRow state resolves row data from snapshot records");
    Require(propertyGridRowBlock.find("gridRowAccessibleName") != std::string::npos, "Accessibility GetPropertyValue GridRow name reads snapshot records");
    Require(propertyGridRowBlock.find("SnapshotGridRowIsSelected") != std::string::npos,
            "Accessibility GetPropertyValue GridRow selection reads snapshot records");
    Require(propertyGridRowBlock.find("ResolveGridControl(") == std::string::npos, "Accessibility GetPropertyValue GridRow state does not resolve live grids");
    Require(propertyGridRowBlock.find("ResolveGridRowIndex(") == std::string::npos,
            "Accessibility GetPropertyValue GridRow state does not query live Grid rows");
    Require(propertyGridRowBlock.find("BuildGridRowAccessibleName(") == std::string::npos,
            "Accessibility GetPropertyValue GridRow state does not build names from live Grid models");
}

void TestAccessibilityTreeGridSelectionReadsUseSnapshots()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for selection snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t getSelectionFunction = source.find("HRESULT AccessibilityProvider::GetSelection");
    const size_t canSelectFunction    = source.find("HRESULT AccessibilityProvider::get_CanSelectMultiple", getSelectionFunction);
    Require(getSelectionFunction != std::string::npos && canSelectFunction != std::string::npos && getSelectionFunction < canSelectFunction,
            "Accessibility GetSelection source block is found");
    const std::string getSelectionBlock = source.substr(getSelectionFunction, canSelectFunction - getSelectionFunction);
    Require(getSelectionBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetSelection reads an immutable published snapshot");
    const size_t getSelectionSnapshotRead = getSelectionBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)");
    const size_t getSelectionLiveResolve  = getSelectionBlock.find("ResolveControl(");
    Require(getSelectionLiveResolve == std::string::npos || getSelectionSnapshotRead < getSelectionLiveResolve,
            "Accessibility GetSelection handles snapshot selection before any live control resolve");
    Require(getSelectionBlock.find("selectedGridRowIds") != std::string::npos, "Accessibility GetSelection reads selected Grid rows from snapshot records");
    Require(getSelectionBlock.find("selectedTreeVisibleIndex") != std::string::npos,
            "Accessibility GetSelection reads selected Tree item identity from snapshot records");
    Require(getSelectionBlock.find("GetSelectionModel(") == std::string::npos, "Accessibility GetSelection does not read live Grid selection");
    Require(getSelectionBlock.find("GetSelectedItemId(") == std::string::npos, "Accessibility GetSelection does not read live Tree selection");
    Require(getSelectionBlock.find("FindVisibleItemById(") == std::string::npos, "Accessibility GetSelection does not query live Tree model selection");
    Require(getSelectionBlock.find("FindRowByStableId(") == std::string::npos, "Accessibility GetSelection does not query live Grid model rows");

    const size_t isSelectedFunction = source.find("HRESULT AccessibilityProvider::get_IsSelected");
    const size_t containerFunction  = source.find("HRESULT AccessibilityProvider::get_SelectionContainer", isSelectedFunction);
    Require(isSelectedFunction != std::string::npos && containerFunction != std::string::npos && isSelectedFunction < containerFunction,
            "Accessibility get_IsSelected source block is found");
    const std::string isSelectedBlock = source.substr(isSelectedFunction, containerFunction - isSelectedFunction);
    Require(isSelectedBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility get_IsSelected reads an immutable published snapshot");
    Require(isSelectedBlock.find("SnapshotGridRowIsSelected") != std::string::npos,
            "Accessibility get_IsSelected resolves Grid selected state from snapshot records");
    Require(isSelectedBlock.find("SnapshotTreeItemIsSelected") != std::string::npos,
            "Accessibility get_IsSelected resolves Tree selected state from snapshot records");
    Require(isSelectedBlock.find("ResolveTreeControl(") == std::string::npos, "Accessibility get_IsSelected does not resolve live trees");
    Require(isSelectedBlock.find("ResolveTreeItemData(") == std::string::npos, "Accessibility get_IsSelected does not read Tree item data");
    Require(isSelectedBlock.find("ResolveGridControl(") == std::string::npos, "Accessibility get_IsSelected does not resolve live grids");
    Require(isSelectedBlock.find("ResolveGridRowIndex(") == std::string::npos, "Accessibility get_IsSelected does not query live Grid rows");

    const size_t canSelectNextFunction = source.find("HRESULT AccessibilityProvider::get_IsSelectionRequired", canSelectFunction);
    Require(canSelectNextFunction != std::string::npos && canSelectFunction < canSelectNextFunction,
            "Accessibility get_CanSelectMultiple source block is found");
    const std::string canSelectBlock = source.substr(canSelectFunction, canSelectNextFunction - canSelectFunction);
    Require(canSelectBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility get_CanSelectMultiple reads an immutable published snapshot");
    Require(canSelectBlock.find("ResolveControl(") == std::string::npos, "Accessibility get_CanSelectMultiple does not resolve live controls");
    Require(canSelectBlock.find("GetSelectionMode(") == std::string::npos, "Accessibility get_CanSelectMultiple does not read live Grid selection mode");

    const size_t selectionRequiredNextFunction = source.find("HRESULT AccessibilityProvider::GetRowHeaders", canSelectNextFunction);
    Require(selectionRequiredNextFunction != std::string::npos && canSelectNextFunction < selectionRequiredNextFunction,
            "Accessibility get_IsSelectionRequired source block is found");
    const std::string selectionRequiredBlock = source.substr(canSelectNextFunction, selectionRequiredNextFunction - canSelectNextFunction);
    Require(selectionRequiredBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility get_IsSelectionRequired reads an immutable published snapshot");
    Require(selectionRequiredBlock.find("ResolveControl(") == std::string::npos, "Accessibility get_IsSelectionRequired does not resolve live controls");

    const size_t containerNextFunction = source.find("HRESULT AccessibilityProvider::Expand", containerFunction);
    Require(containerNextFunction != std::string::npos && containerFunction < containerNextFunction,
            "Accessibility get_SelectionContainer source block is found");
    const std::string containerBlock = source.substr(containerFunction, containerNextFunction - containerFunction);
    Require(containerBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility get_SelectionContainer reads an immutable published snapshot");
    Require(containerBlock.find("SupportsTreeItemSelectionPattern(") == std::string::npos,
            "Accessibility get_SelectionContainer does not re-resolve live Tree selection support");
    Require(containerBlock.find("SupportsGridRowSelectionPattern(") == std::string::npos,
            "Accessibility get_SelectionContainer does not re-resolve live Grid selection support");

    const size_t queryPatternFunction = source.find("AccessibilityPatternQueryResult AccessibilityProvider::QueryPattern");
    const size_t queryFunction        = source.find("HRESULT AccessibilityProvider::QueryInterface", queryPatternFunction);
    Require(queryPatternFunction != std::string::npos && queryFunction != std::string::npos && queryPatternFunction < queryFunction,
            "Accessibility shared QueryPattern source block is found");
    const std::string queryPatternBlock = source.substr(queryPatternFunction, queryFunction - queryPatternFunction);
    const size_t liveRootResolve        = queryPatternBlock.find("ResolveRootControl(");
    const size_t treeSelectionBranch =
        queryPatternBlock.find("_kind == AccessibilityFragmentKind::TreeItem && patternKind == AccessibilityPatternKind::SelectionItem");
    Require(treeSelectionBranch != std::string::npos && (liveRootResolve == std::string::npos || treeSelectionBranch < liveRootResolve),
            "Accessibility shared QueryPattern handles TreeItem selection before any live root resolve");
    const size_t gridRowBranch = queryPatternBlock.find("if (_kind == AccessibilityFragmentKind::GridRow)");
    Require(gridRowBranch != std::string::npos && (liveRootResolve == std::string::npos || gridRowBranch < liveRootResolve),
            "Accessibility shared QueryPattern handles GridRow selection before any live root resolve");
    const std::string treeSelectionPatternBlock = queryPatternBlock.substr(treeSelectionBranch, gridRowBranch - treeSelectionBranch);
    Require(treeSelectionPatternBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility shared QueryPattern TreeItem selection reads an immutable published snapshot");
    Require(treeSelectionPatternBlock.find("SnapshotContainsTreeItem") != std::string::npos,
            "Accessibility shared QueryPattern TreeItem selection validates tree items from snapshot records");
    Require(treeSelectionPatternBlock.find("SupportsTreeItemSelectionPattern(") == std::string::npos,
            "Accessibility shared QueryPattern TreeItem selection does not re-resolve live Tree item support");
    Require(treeSelectionPatternBlock.find("ResolveTreeItemData(") == std::string::npos,
            "Accessibility shared QueryPattern TreeItem selection does not read Tree item data");
    const size_t controlSelectionResolver = queryPatternBlock.find("ResolveSnapshotControlRecord", gridRowBranch);
    const size_t controlSelectionBranch =
        controlSelectionResolver == std::string::npos
            ? std::string::npos
            : queryPatternBlock.rfind("const std::shared_ptr<const AccessibilitySnapshot> snapshot", controlSelectionResolver);
    Require(controlSelectionBranch != std::string::npos && (liveRootResolve == std::string::npos || controlSelectionBranch < liveRootResolve),
            "Accessibility shared QueryPattern handles Tree/Grid container patterns before any live root resolve");
    const std::string gridRowPatternBlock = queryPatternBlock.substr(gridRowBranch, controlSelectionBranch - gridRowBranch);
    Require(gridRowPatternBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility shared QueryPattern GridRow selection reads an immutable published snapshot");
    Require(gridRowPatternBlock.find("SnapshotContainsGridRow") != std::string::npos,
            "Accessibility shared QueryPattern GridRow selection validates visible rows from snapshot records");
    Require(gridRowPatternBlock.find("SnapshotGridRowIsSelected") != std::string::npos,
            "Accessibility shared QueryPattern GridRow selection validates selected rows from snapshot records");
    Require(gridRowPatternBlock.find("SupportsGridRowSelectionPattern(") == std::string::npos,
            "Accessibility shared QueryPattern GridRow selection does not re-resolve live Grid row support");
    Require(gridRowPatternBlock.find("ResolveGridRowIndex(") == std::string::npos,
            "Accessibility shared QueryPattern GridRow selection does not query live Grid rows");
    const std::string controlPatternBlock = queryPatternBlock.substr(controlSelectionBranch, liveRootResolve - controlSelectionBranch);
    Require(controlPatternBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility shared QueryPattern Tree/Grid container patterns read an immutable published snapshot");
    Require(controlPatternBlock.find("controlSupportsSelection") != std::string::npos,
            "Accessibility shared QueryPattern Tree/Grid SelectionPattern validates snapshot records");
    Require(controlPatternBlock.find("controlSupportsTable") != std::string::npos,
            "Accessibility shared QueryPattern Grid TablePattern validates snapshot column count");
    Require(controlPatternBlock.find("SupportsSelectionProviderPattern(") == std::string::npos,
            "Accessibility shared QueryPattern Tree/Grid container patterns do not re-resolve live selection support");
    Require(controlPatternBlock.find("SupportsGridTablePattern(") == std::string::npos,
            "Accessibility shared QueryPattern Grid TablePattern does not re-resolve live table support");

    const size_t propertyValueFunction = source.find("HRESULT AccessibilityProvider::GetPropertyValue");
    const size_t hostProviderFunction  = source.find("HRESULT AccessibilityProvider::get_HostRawElementProvider", propertyValueFunction);
    Require(propertyValueFunction != std::string::npos && hostProviderFunction != std::string::npos && propertyValueFunction < hostProviderFunction,
            "Accessibility GetPropertyValue source block is found");
    const std::string propertyValueBlock = source.substr(propertyValueFunction, hostProviderFunction - propertyValueFunction);
    const size_t selectedProperty        = propertyValueBlock.find("UIA_SelectionItemIsSelectedPropertyId");
    const size_t propertyLiveRootResolve = propertyValueBlock.find("ResolveRootControl(");
    Require(selectedProperty != std::string::npos && (propertyLiveRootResolve == std::string::npos || selectedProperty < propertyLiveRootResolve),
            "Accessibility GetPropertyValue handles Tree/Grid selected state before any live root resolve");
    const std::string selectedPropertyBlock = propertyValueBlock.substr(0u, propertyLiveRootResolve);
    Require(selectedPropertyBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetPropertyValue selected state reads an immutable published snapshot");
    Require(selectedPropertyBlock.find("SnapshotTreeItemIsSelected") != std::string::npos,
            "Accessibility GetPropertyValue selected state resolves Tree state from snapshot records");
    Require(selectedPropertyBlock.find("SnapshotGridRowIsSelected") != std::string::npos,
            "Accessibility GetPropertyValue selected state resolves Grid state from snapshot records");
    Require(selectedPropertyBlock.find("ResolveTreeControl(") == std::string::npos,
            "Accessibility GetPropertyValue selected state does not resolve live trees");
    Require(selectedPropertyBlock.find("ResolveGridControl(") == std::string::npos,
            "Accessibility GetPropertyValue selected state does not resolve live grids");
}

void TestAccessibilityTreeItemStateReadsUseSnapshots()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for TreeItem state snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t propertyFunction = source.find("HRESULT AccessibilityProvider::GetPropertyValue");
    const size_t hostFunction     = source.find("HRESULT AccessibilityProvider::get_HostRawElementProvider", propertyFunction);
    Require(propertyFunction != std::string::npos && hostFunction != std::string::npos && propertyFunction < hostFunction,
            "Accessibility GetPropertyValue source block is found for TreeItem state");
    const std::string propertyBlock = source.substr(propertyFunction, hostFunction - propertyFunction);
    const size_t propertyLiveRoot   = propertyBlock.find("ResolveRootControl(");
    const size_t treeItemBranch     = propertyBlock.find("if (_kind == AccessibilityFragmentKind::TreeItem)");
    Require(treeItemBranch != std::string::npos && (propertyLiveRoot == std::string::npos || treeItemBranch < propertyLiveRoot),
            "Accessibility GetPropertyValue handles TreeItem state before any live root resolve");
    const std::string treeItemPropertyBlock = propertyBlock.substr(treeItemBranch, propertyLiveRoot - treeItemBranch);
    Require(treeItemPropertyBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetPropertyValue TreeItem state reads an immutable published snapshot");
    Require(treeItemPropertyBlock.find("FindSnapshotTreeItemRecord") != std::string::npos,
            "Accessibility GetPropertyValue TreeItem state resolves item data from snapshot records");
    Require(treeItemPropertyBlock.find("UIA_NamePropertyId") != std::string::npos, "Accessibility GetPropertyValue TreeItem name is snapshot-backed");
    Require(treeItemPropertyBlock.find("UIA_LevelPropertyId") != std::string::npos, "Accessibility GetPropertyValue TreeItem level is snapshot-backed");
    Require(treeItemPropertyBlock.find("UIA_ExpandCollapseExpandCollapseStatePropertyId") != std::string::npos,
            "Accessibility GetPropertyValue TreeItem expand-collapse state is snapshot-backed");
    Require(treeItemPropertyBlock.find("ResolveTreeControl(") == std::string::npos,
            "Accessibility GetPropertyValue TreeItem state does not resolve live trees");
    Require(treeItemPropertyBlock.find("ResolveTreeItemData(") == std::string::npos,
            "Accessibility GetPropertyValue TreeItem state does not read Tree item data");
    Require(treeItemPropertyBlock.find("GetSelectedItemId(") == std::string::npos,
            "Accessibility GetPropertyValue TreeItem state does not read live Tree selection");

    const size_t queryPatternFunction = source.find("AccessibilityPatternQueryResult AccessibilityProvider::QueryPattern");
    const size_t queryFunction        = source.find("HRESULT AccessibilityProvider::QueryInterface", queryPatternFunction);
    Require(queryPatternFunction != std::string::npos && queryFunction != std::string::npos && queryPatternFunction < queryFunction,
            "Accessibility shared QueryPattern source block is found for TreeItem expand-collapse");
    const std::string queryPatternBlock = source.substr(queryPatternFunction, queryFunction - queryPatternFunction);
    const size_t patternLiveRoot        = queryPatternBlock.find("ResolveRootControl(");
    const size_t expandPatternBranch =
        queryPatternBlock.find("_kind == AccessibilityFragmentKind::TreeItem && patternKind == AccessibilityPatternKind::ExpandCollapse");
    Require(expandPatternBranch != std::string::npos && (patternLiveRoot == std::string::npos || expandPatternBranch < patternLiveRoot),
            "Accessibility shared QueryPattern handles TreeItem expand-collapse before any live root resolve");
    const std::string expandPatternBlock = queryPatternBlock.substr(expandPatternBranch, patternLiveRoot - expandPatternBranch);
    Require(expandPatternBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility shared QueryPattern TreeItem expand-collapse reads an immutable published snapshot");
    Require(expandPatternBlock.find("FindSnapshotTreeItemRecord") != std::string::npos,
            "Accessibility shared QueryPattern TreeItem expand-collapse validates snapshot records");
    Require(expandPatternBlock.find("SupportsTreeItemExpandCollapsePattern(") == std::string::npos,
            "Accessibility shared QueryPattern TreeItem expand-collapse does not re-resolve live Tree item support");

    const size_t stateFunction = source.find("HRESULT AccessibilityProvider::get_ExpandCollapseState");
    const size_t rowFunction   = source.find("HRESULT AccessibilityProvider::get_Row(int*", stateFunction);
    Require(stateFunction != std::string::npos && rowFunction != std::string::npos && stateFunction < rowFunction,
            "Accessibility get_ExpandCollapseState source block is found");
    const std::string stateBlock = source.substr(stateFunction, rowFunction - stateFunction);
    Require(stateBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility get_ExpandCollapseState reads an immutable published snapshot");
    Require(stateBlock.find("FindSnapshotTreeItemRecord") != std::string::npos, "Accessibility get_ExpandCollapseState resolves state from snapshot records");
    Require(stateBlock.find("ResolveTreeItemData(") == std::string::npos, "Accessibility get_ExpandCollapseState does not read Tree item data");
}

void TestAccessibilityTextPatternDocumentAndSelectionReadsUseSnapshots()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for TextPattern snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    Require(source.find("controlTextSelectionStart") != std::string::npos, "Accessibility snapshots capture TextPattern selection start");
    Require(source.find("controlTextSelectionEnd") != std::string::npos, "Accessibility snapshots capture TextPattern selection end");
    Require(source.find("controlTextSelectionBoundsDip") != std::string::npos, "Accessibility snapshots capture TextPattern selection bounding rectangles");

    const size_t getSelectionFunction = source.find("HRESULT AccessibilityProvider::GetSelection");
    const size_t canSelectFunction    = source.find("HRESULT AccessibilityProvider::get_CanSelectMultiple", getSelectionFunction);
    Require(getSelectionFunction != std::string::npos && canSelectFunction != std::string::npos && getSelectionFunction < canSelectFunction,
            "Accessibility GetSelection source block is found for TextPattern snapshot guard");
    const std::string getSelectionBlock = source.substr(getSelectionFunction, canSelectFunction - getSelectionFunction);
    Require(getSelectionBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility GetSelection resolves TextPattern selection from snapshot records");
    Require(getSelectionBlock.find("ResolveControl()") == std::string::npos,
            "Accessibility GetSelection does not resolve live controls for TextPattern selection");
    Require(getSelectionBlock.find("SupportsTextPattern(control)") == std::string::npos,
            "Accessibility GetSelection does not re-check live TextPattern support");
    Require(getSelectionBlock.find("GetControlAccessibleSelectionRange(") == std::string::npos,
            "Accessibility GetSelection does not read live selection state");
    Require(getSelectionBlock.find("GetControlAccessibleTextRangeText(control)") == std::string::npos, "Accessibility GetSelection does not read live text");
    Require(getSelectionBlock.find("controlTextSelectionBoundsDip") != std::string::npos,
            "Accessibility GetSelection gives selection ranges snapshot bounding rectangles");

    const size_t boundingRectanglesFunction = source.find("HRESULT AccessibilityTextRangeProvider::GetBoundingRectangles");
    const size_t enclosingElementFunction   = source.find("HRESULT AccessibilityTextRangeProvider::GetEnclosingElement", boundingRectanglesFunction);
    Require(boundingRectanglesFunction != std::string::npos && enclosingElementFunction != std::string::npos &&
                boundingRectanglesFunction < enclosingElementFunction,
            "AccessibilityTextRangeProvider GetBoundingRectangles source block is found");
    const std::string boundingRectanglesBlock = source.substr(boundingRectanglesFunction, enclosingElementFunction - boundingRectanglesFunction);
    const size_t snapshotBoundsRead           = boundingRectanglesBlock.find("_boundsOverrideDip");
    const size_t liveHostRead                 = boundingRectanglesBlock.find("ResolveHost()");
    const size_t liveControlRead              = boundingRectanglesBlock.find("ResolveControl()");
    Require(snapshotBoundsRead != std::string::npos, "AccessibilityTextRangeProvider GetBoundingRectangles checks snapshot selection bounds");
    Require(liveHostRead == std::string::npos || snapshotBoundsRead < liveHostRead,
            "AccessibilityTextRangeProvider GetBoundingRectangles uses snapshot bounds before live host fallback");
    Require(liveControlRead == std::string::npos || snapshotBoundsRead < liveControlRead,
            "AccessibilityTextRangeProvider GetBoundingRectangles uses snapshot bounds before live control fallback");

    const size_t getEnclosingElementFunction = source.find("HRESULT AccessibilityTextRangeProvider::GetEnclosingElement");
    const size_t getTextFunction             = source.find("HRESULT AccessibilityTextRangeProvider::GetText", getEnclosingElementFunction);
    Require(getEnclosingElementFunction != std::string::npos && getTextFunction != std::string::npos && getEnclosingElementFunction < getTextFunction,
            "AccessibilityTextRangeProvider GetEnclosingElement source block is found");
    const std::string enclosingElementBlock = source.substr(getEnclosingElementFunction, getTextFunction - getEnclosingElementFunction);
    Require(enclosingElementBlock.find("FindControlNavigationRecord") != std::string::npos,
            "AccessibilityTextRangeProvider GetEnclosingElement validates the enclosing element from the snapshot");
    Require(enclosingElementBlock.find("ResolveControl()") == std::string::npos,
            "AccessibilityTextRangeProvider GetEnclosingElement does not resolve live controls");

    const size_t visibleRangesFunction  = source.find("HRESULT AccessibilityProvider::GetVisibleRanges");
    const size_t rangeFromChildFunction = source.find("HRESULT AccessibilityProvider::RangeFromChild", visibleRangesFunction);
    Require(visibleRangesFunction != std::string::npos && rangeFromChildFunction != std::string::npos && visibleRangesFunction < rangeFromChildFunction,
            "Accessibility GetVisibleRanges source block is found");
    const std::string visibleRangesBlock = source.substr(visibleRangesFunction, rangeFromChildFunction - visibleRangesFunction);
    Require(visibleRangesBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility GetVisibleRanges resolves text length from snapshot records");
    Require(visibleRangesBlock.find("ResolveControl()") == std::string::npos, "Accessibility GetVisibleRanges does not resolve live controls");
    Require(visibleRangesBlock.find("SupportsTextPattern(control)") == std::string::npos,
            "Accessibility GetVisibleRanges does not re-check live TextPattern support");
    Require(visibleRangesBlock.find("GetControlAccessibleTextRangeText(control)") == std::string::npos,
            "Accessibility GetVisibleRanges does not read live text");

    const size_t rangeFromPointFunction = source.find("HRESULT AccessibilityProvider::RangeFromPoint");
    const size_t rangeFromPointHelper   = source.find("HRESULT AccessibilityProvider::ExecuteResolveTextRangeFromPointOnWindowThread", rangeFromPointFunction);
    Require(rangeFromPointFunction != std::string::npos && rangeFromPointHelper != std::string::npos && rangeFromPointFunction < rangeFromPointHelper,
            "Accessibility RangeFromPoint source block is found");
    const std::string rangeFromPointBlock = source.substr(rangeFromPointFunction, rangeFromPointHelper - rangeFromPointFunction);
    Require(rangeFromPointBlock.find("DispatchActionToWindowThread") != std::string::npos,
            "Accessibility RangeFromPoint dispatches exact hit testing to the window thread");
    Require(rangeFromPointBlock.find("ResolveHost()") == std::string::npos,
            "Accessibility RangeFromPoint public provider method does not resolve the live host");
    Require(rangeFromPointBlock.find("ResolveControl()") == std::string::npos,
            "Accessibility RangeFromPoint public provider method does not resolve live controls");
    Require(rangeFromPointBlock.find("GetControlAccessibleTextRangeText(control)") == std::string::npos,
            "Accessibility RangeFromPoint public provider method does not read live text");

    const size_t documentRangeFunction      = source.find("HRESULT AccessibilityProvider::get_DocumentRange", rangeFromPointHelper);
    const size_t supportedSelectionFunction = source.find("HRESULT AccessibilityProvider::get_SupportedTextSelection", documentRangeFunction);
    Require(documentRangeFunction != std::string::npos && supportedSelectionFunction != std::string::npos && documentRangeFunction < supportedSelectionFunction,
            "Accessibility get_DocumentRange source block is found");
    const std::string documentRangeBlock = source.substr(documentRangeFunction, supportedSelectionFunction - documentRangeFunction);
    Require(documentRangeBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility get_DocumentRange resolves text length from snapshot records");
    Require(documentRangeBlock.find("ResolveControl()") == std::string::npos, "Accessibility get_DocumentRange does not resolve live controls");
    Require(documentRangeBlock.find("SupportsTextPattern(control)") == std::string::npos,
            "Accessibility get_DocumentRange does not re-check live TextPattern support");
    Require(documentRangeBlock.find("GetControlAccessibleTextRangeText(control)") == std::string::npos,
            "Accessibility get_DocumentRange does not read live text");

    const size_t supportedSelectionEnd = source.find("HRESULT AccessibilityProvider::GetActiveComposition", supportedSelectionFunction);
    Require(supportedSelectionEnd != std::string::npos && supportedSelectionFunction < supportedSelectionEnd,
            "Accessibility get_SupportedTextSelection source block is found");
    const std::string supportedSelectionBlock = source.substr(supportedSelectionFunction, supportedSelectionEnd - supportedSelectionFunction);
    Require(supportedSelectionBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility get_SupportedTextSelection resolves support from snapshot records");
    Require(supportedSelectionBlock.find("ResolveControl()") == std::string::npos, "Accessibility get_SupportedTextSelection does not resolve live controls");
    Require(supportedSelectionBlock.find("SupportsTextPattern(control)") == std::string::npos,
            "Accessibility get_SupportedTextSelection does not re-check live TextPattern support");

    const size_t resolveTextFunction = source.find("std::wstring AccessibilityTextRangeProvider::ResolveText");
    const size_t clampFunction       = source.find("TextRangeSpan AccessibilityTextRangeProvider::ClampCurrentRange", resolveTextFunction);
    Require(resolveTextFunction != std::string::npos && clampFunction != std::string::npos && resolveTextFunction < clampFunction,
            "AccessibilityTextRangeProvider ResolveText source block is found");
    const std::string resolveTextBlock = source.substr(resolveTextFunction, clampFunction - resolveTextFunction);
    Require(resolveTextBlock.find("FindControlNavigationRecord") != std::string::npos, "AccessibilityTextRangeProvider resolves text from snapshot records");
    Require(resolveTextBlock.find("ResolveControl()") == std::string::npos, "AccessibilityTextRangeProvider does not resolve live controls to read text");
    Require(resolveTextBlock.find("GetControlAccessibleTextRangeText(control)") == std::string::npos,
            "AccessibilityTextRangeProvider does not read live control text");
}

void TestAccessibilityTextEditCompositionReadsUseSnapshots()
{
    const std::filesystem::path repoRoot   = FindRepoRootForDxUiTests();
    const std::filesystem::path sourcePath = repoRoot / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for TextEdit composition snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    Require(source.find("controlTextCompositionStart") != std::string::npos, "Accessibility snapshots capture TextEdit active-composition start");
    Require(source.find("controlTextCompositionEnd") != std::string::npos, "Accessibility snapshots capture TextEdit active-composition end");
    Require(source.find("controlTextConversionTargetStart") != std::string::npos, "Accessibility snapshots capture TextEdit conversion-target start");
    Require(source.find("controlTextConversionTargetEnd") != std::string::npos, "Accessibility snapshots capture TextEdit conversion-target end");

    const size_t activeCompositionFunction = source.find("HRESULT AccessibilityProvider::GetActiveComposition");
    const size_t conversionTargetFunction  = source.find("HRESULT AccessibilityProvider::GetConversionTarget", activeCompositionFunction);
    Require(activeCompositionFunction != std::string::npos && conversionTargetFunction != std::string::npos &&
                activeCompositionFunction < conversionTargetFunction,
            "Accessibility GetActiveComposition source block is found");
    const std::string activeCompositionBlock = source.substr(activeCompositionFunction, conversionTargetFunction - activeCompositionFunction);
    Require(activeCompositionBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility GetActiveComposition resolves native IME state from snapshot records");
    Require(activeCompositionBlock.find("controlTextCompositionStart") != std::string::npos &&
                activeCompositionBlock.find("controlTextCompositionEnd") != std::string::npos,
            "Accessibility GetActiveComposition reads snapshot active-composition endpoints");
    Require(activeCompositionBlock.find("ResolveControl()") == std::string::npos, "Accessibility GetActiveComposition does not resolve live controls");
    Require(activeCompositionBlock.find("ResolveHost()") == std::string::npos, "Accessibility GetActiveComposition does not resolve the live host");
    Require(activeCompositionBlock.find("TryReadNativeTextInputState(") == std::string::npos,
            "Accessibility GetActiveComposition does not read live native text-input state");

    const size_t setValueFunction = source.find("HRESULT AccessibilityProvider::SetValue(LPCWSTR", conversionTargetFunction);
    Require(setValueFunction != std::string::npos && conversionTargetFunction < setValueFunction, "Accessibility GetConversionTarget source block is found");
    const std::string conversionTargetBlock = source.substr(conversionTargetFunction, setValueFunction - conversionTargetFunction);
    Require(conversionTargetBlock.find("ResolveSnapshotControlRecord") != std::string::npos,
            "Accessibility GetConversionTarget resolves native IME state from snapshot records");
    Require(conversionTargetBlock.find("controlTextConversionTargetStart") != std::string::npos &&
                conversionTargetBlock.find("controlTextConversionTargetEnd") != std::string::npos,
            "Accessibility GetConversionTarget reads snapshot conversion-target endpoints");
    Require(conversionTargetBlock.find("ResolveControl()") == std::string::npos, "Accessibility GetConversionTarget does not resolve live controls");
    Require(conversionTargetBlock.find("ResolveHost()") == std::string::npos, "Accessibility GetConversionTarget does not resolve the live host");
    Require(conversionTargetBlock.find("TryReadNativeTextInputState(") == std::string::npos,
            "Accessibility GetConversionTarget does not read live native text-input state");

    const std::filesystem::path nativeInputPath = repoRoot / L"Common" / L"DxUi" / L"DxUi.NativeTextInput.cpp";
    std::ifstream nativeInput(nativeInputPath);
    Require(nativeInput.good(), "NativeTextInput source is readable for TextEdit snapshot publication guard");
    const std::string nativeSource((std::istreambuf_iterator<char>(nativeInput)), std::istreambuf_iterator<char>());

    const size_t raiseEventsFunction = nativeSource.find("void WindowHost::RaiseNativeTextInputAccessibilityEvents");
    const size_t nextFunction        = nativeSource.find("void WindowHost::ApplyNativeTextInputCompositionStateToCache", raiseEventsFunction);
    Require(raiseEventsFunction != std::string::npos && nextFunction != std::string::npos && raiseEventsFunction < nextFunction,
            "native text-input accessibility event source block is found");
    const std::string raiseEventsBlock = nativeSource.substr(raiseEventsFunction, nextFunction - raiseEventsFunction);
    const size_t snapshotRefresh       = raiseEventsBlock.find("RefreshWindowHostAccessibilitySnapshot(_hwnd, this)");
    const size_t compositionEvent = raiseEventsBlock.find("RaiseNativeTextInputAccessibilityEvent(TextInputAutomationEventKind::TextEditCompositionChanged)");
    const size_t conversionEvent =
        raiseEventsBlock.find("RaiseNativeTextInputAccessibilityEvent(TextInputAutomationEventKind::TextEditConversionTargetChanged)");
    Require(snapshotRefresh != std::string::npos, "native text-input accessibility events publish a fresh snapshot");
    Require(compositionEvent != std::string::npos && snapshotRefresh < compositionEvent,
            "TextEdit composition-changed events are raised after snapshot publication");
    Require(conversionEvent != std::string::npos && snapshotRefresh < conversionEvent,
            "TextEdit conversion-target events are raised after snapshot publication");
}

void TestAccessibilityNavigateUsesSnapshot()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for Navigate snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t navigateFunction = source.find("HRESULT AccessibilityProvider::Navigate");
    const size_t runtimeFunction  = source.find("HRESULT AccessibilityProvider::GetRuntimeId", navigateFunction);
    Require(navigateFunction != std::string::npos && runtimeFunction != std::string::npos && navigateFunction < runtimeFunction,
            "Accessibility Navigate source block is found");

    const std::string navigateBlock = source.substr(navigateFunction, runtimeFunction - navigateFunction);
    Require(navigateBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility Navigate reads an immutable published snapshot");
    Require(navigateBlock.find("ResolveSnapshotNavigationTarget") != std::string::npos,
            "Accessibility Navigate resolves provider identity from snapshot navigation records");
    Require(navigateBlock.find("ResolveRootControl(") == std::string::npos, "Accessibility Navigate does not resolve the live root");
    Require(navigateBlock.find("ResolveControl(") == std::string::npos, "Accessibility Navigate does not resolve live controls");
    Require(navigateBlock.find("FindFirstSemanticControl(") == std::string::npos, "Accessibility Navigate does not walk live first semantic controls");
    Require(navigateBlock.find("FindNextSemanticControl(") == std::string::npos, "Accessibility Navigate does not walk live next semantic controls");
    Require(navigateBlock.find("FindPreviousSemanticControl(") == std::string::npos, "Accessibility Navigate does not walk live previous semantic controls");
    Require(navigateBlock.find("GetModel(") == std::string::npos, "Accessibility Navigate does not read Tree/Grid models");
}

void TestAccessibilityTreeSnapshotHitRecordsUseVisibleGeometryOnly()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for tree snapshot hot-path guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t treeFunction = source.find("void AppendTreeAccessibilityPointHits");
    const size_t gridFunction = source.find("void AppendGridAccessibilityPointHits", treeFunction);
    Require(treeFunction != std::string::npos && gridFunction != std::string::npos && treeFunction < gridFunction,
            "Accessibility Tree snapshot hit-record source block is found");

    const std::string treeBlock = source.substr(treeFunction, gridFunction - treeFunction);
    Require(treeBlock.find("GetFirstVisibleItemIndex()") != std::string::npos, "Tree snapshot hit records start at the first visible row");
    Require(treeBlock.find("GetVisibleItemHitRect(") != std::string::npos, "Tree snapshot hit records use geometry-only row rects");
    Require(treeBlock.find("GetItemLayoutMetrics(") == std::string::npos, "Tree snapshot hit records do not build full layout metrics");
    Require(treeBlock.find("GetVisibleItem(") == std::string::npos, "Tree snapshot hit records do not materialize TreeItemData");
}

void TestAttachedWindowHostWmGetObjectReturnsAccessibilityProvider()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    const LRESULT result = SendMessageW(window.Hwnd(), WM_GETOBJECT, 0, static_cast<LPARAM>(UiaRootObjectId));
    Require(result != 0, "attached DX host returns a UIA provider from WM_GETOBJECT");
}

void TestAccessibilityRootRuntimeIdIncludesProviderSpecificValues()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "runtime-id test creates a root accessibility provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> rootFragment;
    RequireSucceeded(rootProvider.query_to(rootFragment.put()), "runtime-id test root provider exposes fragment navigation");

    SAFEARRAY* runtimeId = nullptr;
    RequireSucceeded(rootFragment->GetRuntimeId(&runtimeId), "root provider runtime-id lookup succeeds");
    Require(runtimeId != nullptr, "root provider returns a runtime-id array");
    const auto destroyRuntimeId = wil::scope_exit([&] { SafeArrayDestroy(runtimeId); });

    LONG lowerBound = 0;
    LONG upperBound = -1;
    RequireSucceeded(SafeArrayGetLBound(runtimeId, 1, &lowerBound), "root runtime-id lower bound lookup succeeds");
    RequireSucceeded(SafeArrayGetUBound(runtimeId, 1, &upperBound), "root runtime-id upper bound lookup succeeds");
    Require(lowerBound == 0, "root runtime-id starts at index zero");
    Require(upperBound >= 2, "root runtime-id includes provider-specific values beyond UiaAppendRuntimeId");

    LONG firstValue  = 0;
    LONG secondValue = 0;
    LONG thirdValue  = 0;
    LONG index       = 0;
    RequireSucceeded(SafeArrayGetElement(runtimeId, &index, &firstValue), "root runtime-id first element lookup succeeds");
    index = 1;
    RequireSucceeded(SafeArrayGetElement(runtimeId, &index, &secondValue), "root runtime-id second element lookup succeeds");
    index = 2;
    RequireSucceeded(SafeArrayGetElement(runtimeId, &index, &thirdValue), "root runtime-id third element lookup succeeds");
    Require(firstValue == UiaAppendRuntimeId, "root runtime-id starts with UiaAppendRuntimeId");
    Require(secondValue != 0 || thirdValue != 0, "root runtime-id appends non-zero provider-specific identity values");
}

std::wstring ReadTextRangeText(ITextRangeProvider& range, int maxLength, const char* context)
{
    BSTR text = nullptr;
    RequireSucceeded(range.GetText(maxLength, &text), context);
    const auto freeText = wil::scope_exit([&] { SysFreeString(text); });
    return std::wstring(text ? text : L"");
}

wil::com_ptr_nothrow<ITextRangeProvider> GetSingleTextRangeFromArray(SAFEARRAY* array, const char* context)
{
    Require(array != nullptr, context);

    LONG lowerBound = 0;
    LONG upperBound = -1;
    RequireSucceeded(SafeArrayGetLBound(array, 1, &lowerBound), context);
    RequireSucceeded(SafeArrayGetUBound(array, 1, &upperBound), context);
    Require(lowerBound == upperBound, context);

    IUnknown* rawUnknown = nullptr;
    LONG index           = lowerBound;
    RequireSucceeded(SafeArrayGetElement(array, &index, &rawUnknown), context);
    wil::com_ptr_nothrow<IUnknown> unknown;
    unknown.attach(rawUnknown);
    Require(unknown != nullptr, context);

    wil::com_ptr_nothrow<ITextRangeProvider> range;
    RequireSucceeded(unknown.query_to(range.put()), context);
    Require(range != nullptr, context);
    return range;
}

std::vector<double> ReadDoubleArray(SAFEARRAY* array, const char* context)
{
    Require(array != nullptr, context);

    LONG lowerBound = 0;
    LONG upperBound = -1;
    RequireSucceeded(SafeArrayGetLBound(array, 1, &lowerBound), context);
    RequireSucceeded(SafeArrayGetUBound(array, 1, &upperBound), context);
    Require(upperBound >= lowerBound, context);

    std::vector<double> values(static_cast<size_t>(upperBound - lowerBound + 1));
    for (LONG index = lowerBound; index <= upperBound; ++index)
    {
        double value = 0.0;
        RequireSucceeded(SafeArrayGetElement(array, &index, &value), context);
        values[static_cast<size_t>(index - lowerBound)] = value;
    }
    return values;
}

std::vector<size_t> ResolveVisualLineStarts(const RedSalamander::DxUi::WindowHost& host,
                                            const RedSalamander::DxUi::TextField& field,
                                            std::wstring_view text,
                                            const char* context)
{
    std::vector<size_t> starts{0u};
    std::optional<D2D1_RECT_F> lineRect = field.TryGetTextInputCaretRect(host, 0u);
    Require(lineRect.has_value(), context);

    for (size_t index = 1u; index <= text.size(); ++index)
    {
        const std::optional<D2D1_RECT_F> rect = field.TryGetTextInputCaretRect(host, index);
        Require(rect.has_value(), context);
        constexpr float kLineToleranceDip = 1.0f;
        const bool sameVisualLine =
            std::fabs(rect->top - lineRect->top) <= kLineToleranceDip && std::fabs(rect->bottom - lineRect->bottom) <= kLineToleranceDip;
        if (! sameVisualLine)
        {
            starts.push_back(index);
            lineRect = rect;
        }
    }

    return starts;
}

[[nodiscard]] RECT DipRectToScreenRect(AttachedHostWindow& window, const D2D1_RECT_F& rectDip)
{
    POINT points[2]{
        {static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.left))), static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.top)))},
        {static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.right))),
         static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.bottom)))}};
    MapWindowPoints(window.Hwnd(), nullptr, points, 2);
    return RECT{points[0].x, points[0].y, points[1].x, points[1].y};
}

[[nodiscard]] D2D1_RECT_F UnionRects(const D2D1_RECT_F& first, const D2D1_RECT_F& second) noexcept
{
    return D2D1::RectF(
        (std::min)(first.left, second.left), (std::min)(first.top, second.top), (std::max)(first.right, second.right), (std::max)(first.bottom, second.bottom));
}

void TestAccessibilityPasswordRevealButtonReadsUseSnapshots()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Accessibility.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Accessibility source is readable for password reveal snapshot guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t queryPatternFunction = source.find("AccessibilityPatternQueryResult AccessibilityProvider::QueryPattern");
    const size_t queryFunction        = source.find("HRESULT AccessibilityProvider::QueryInterface", queryPatternFunction);
    Require(queryPatternFunction != std::string::npos && queryFunction != std::string::npos && queryPatternFunction < queryFunction,
            "Accessibility shared QueryPattern source block is found for password reveal guard");
    const std::string queryPatternBlock = source.substr(queryPatternFunction, queryFunction - queryPatternFunction);
    const size_t queryRevealBranch      = queryPatternBlock.find("if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)");
    const size_t queryLiveRoot          = queryPatternBlock.find("ResolveRootControl(");
    Require(queryRevealBranch != std::string::npos && (queryLiveRoot == std::string::npos || queryRevealBranch < queryLiveRoot),
            "Accessibility shared QueryPattern handles password reveal Invoke before any live root resolve");
    const std::string queryRevealBlock = queryPatternBlock.substr(queryRevealBranch, queryLiveRoot - queryRevealBranch);
    Require(queryRevealBlock.find("patternKind != AccessibilityPatternKind::Invoke") != std::string::npos,
            "Accessibility shared QueryPattern restricts password reveal to Invoke");
    Require(queryRevealBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility shared QueryPattern password reveal Invoke reads an immutable published snapshot");
    Require(queryRevealBlock.find("hasPasswordRevealButton") != std::string::npos,
            "Accessibility shared QueryPattern password reveal Invoke validates snapshot visibility");
    Require(queryRevealBlock.find("IsTextFieldPasswordRevealButtonVisible(") == std::string::npos,
            "Accessibility shared QueryPattern password reveal Invoke does not read live TextField reveal state");
    Require(queryRevealBlock.find("ResolveControl(") == std::string::npos,
            "Accessibility shared QueryPattern password reveal Invoke does not resolve live controls");

    const size_t propertyValueFunction = source.find("HRESULT AccessibilityProvider::GetPropertyValue");
    const size_t hostProviderFunction  = source.find("HRESULT AccessibilityProvider::get_HostRawElementProvider", propertyValueFunction);
    Require(propertyValueFunction != std::string::npos && hostProviderFunction != std::string::npos && propertyValueFunction < hostProviderFunction,
            "Accessibility GetPropertyValue source block is found for password reveal guard");
    const std::string propertyValueBlock = source.substr(propertyValueFunction, hostProviderFunction - propertyValueFunction);
    const size_t propertyRevealBranch    = propertyValueBlock.find("if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)");
    const size_t propertyLiveRoot        = propertyValueBlock.find("ResolveRootControl(");
    Require(propertyRevealBranch != std::string::npos && (propertyLiveRoot == std::string::npos || propertyRevealBranch < propertyLiveRoot),
            "Accessibility GetPropertyValue handles password reveal state before any live root resolve");
    const std::string propertyRevealBlock = propertyValueBlock.substr(propertyRevealBranch, propertyLiveRoot - propertyRevealBranch);
    Require(propertyRevealBlock.find("CaptureAccessibilitySnapshot(_target, _hwnd)") != std::string::npos,
            "Accessibility GetPropertyValue password reveal state reads an immutable published snapshot");
    Require(propertyRevealBlock.find("passwordRevealButtonAccessibleName") != std::string::npos,
            "Accessibility GetPropertyValue password reveal name reads snapshot records");
    Require(propertyRevealBlock.find("passwordRevealButtonEnabled") != std::string::npos,
            "Accessibility GetPropertyValue password reveal enabled state reads snapshot records");
    Require(propertyRevealBlock.find("GetPasswordRevealAccessibleName(") == std::string::npos,
            "Accessibility GetPropertyValue password reveal state does not read live TextField name");
    Require(propertyRevealBlock.find("IsPasswordRevealButtonVisibleForAccessibility(") == std::string::npos,
            "Accessibility GetPropertyValue password reveal state does not read live TextField visibility");
    Require(propertyRevealBlock.find("dynamic_cast") == std::string::npos, "Accessibility GetPropertyValue password reveal state does not cast live controls");
}

void TestAccessibilityProviderExposesInvokeToggleAndLabeledValuePatterns()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 140.0f, 32.0f));

    auto* toggle = root->AddChild<Toggle>(L"Menu bar");
    toggle->SetBounds(D2D1::RectF(0.0f, 40.0f, 220.0f, 88.0f));
    toggle->SetChecked(true);

    auto* label = root->AddChild<Label>(L"Search");
    label->SetBounds(D2D1::RectF(0.0f, 100.0f, 120.0f, 124.0f));
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 128.0f, 220.0f, 156.0f));
    field->SetAccessibleHelpText(L"Type a local path");
    label->SetMnemonicTarget(field);

    auto* comboLabel = root->AddChild<Label>(L"Mode");
    comboLabel->SetBounds(D2D1::RectF(0.0f, 168.0f, 120.0f, 192.0f));
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"current");
    combo->SetBounds(D2D1::RectF(0.0f, 196.0f, 220.0f, 224.0f));
    comboLabel->SetMnemonicTarget(combo);

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "debug accessibility provider is created for attached DX host");

    wil::com_ptr_nothrow<IRawElementProviderFragment> buttonProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 24.0f, 16.0f, "button accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> buttonSimple;
    RequireSucceeded(buttonProvider.query_to(buttonSimple.put()), "button accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*buttonSimple.get(), UIA_ControlTypePropertyId, "button exposes UIA control type") == UIA_ButtonControlTypeId,
            "button accessibility provider reports button control type");
    Require(ReadProviderStringProperty(*buttonSimple.get(), UIA_NamePropertyId, "button exposes accessibility name") == L"Run",
            "button accessibility provider reports button text as the accessible name");
    wil::com_ptr_nothrow<IUnknown> invokePattern;
    RequireSucceeded(buttonSimple->GetPatternProvider(UIA_InvokePatternId, invokePattern.put()), "button invoke pattern lookup succeeds");
    Require(invokePattern != nullptr, "button accessibility provider exposes invoke pattern");

    wil::com_ptr_nothrow<IRawElementProviderFragment> toggleProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 200.0f, 64.0f, "toggle accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> toggleSimple;
    RequireSucceeded(toggleProvider.query_to(toggleSimple.put()), "toggle accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*toggleSimple.get(), UIA_NamePropertyId, "toggle exposes accessibility name") == L"Menu bar",
            "toggle accessibility provider reports displayed label text");
    wil::com_ptr_nothrow<IUnknown> togglePatternUnknown;
    RequireSucceeded(toggleSimple->GetPatternProvider(UIA_TogglePatternId, togglePatternUnknown.put()), "toggle pattern lookup succeeds");
    Require(togglePatternUnknown != nullptr, "toggle accessibility provider exposes toggle pattern");
    wil::com_ptr_nothrow<IToggleProvider> togglePattern;
    RequireSucceeded(togglePatternUnknown.query_to(togglePattern.put()), "toggle pattern supports IToggleProvider");
    ToggleState toggleState = ToggleState_Off;
    RequireSucceeded(togglePattern->get_ToggleState(&toggleState), "toggle state query succeeds");
    Require(toggleState == ToggleState_On, "toggle accessibility provider reports the checked state");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 142.0f, "text field accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "text field accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*fieldSimple.get(), UIA_ControlTypePropertyId, "text field exposes UIA control type") == UIA_EditControlTypeId,
            "text field accessibility provider reports edit control type");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_NamePropertyId, "text field exposes accessibility name") == L"Search",
            "text field accessibility provider uses its associated label as the accessible name");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_HelpTextPropertyId, "text field exposes accessibility help text") == L"Type a local path",
            "text field accessibility provider reports explicit help text");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_ValueValuePropertyId, "text field exposes current value") == L"alpha",
            "text field accessibility provider reports the current value");
    Require(! ReadProviderBoolProperty(*fieldSimple.get(), UIA_ValueIsReadOnlyPropertyId, "text field exposes editable state"),
            "text field accessibility provider reports editable state");

    wil::com_ptr_nothrow<IUnknown> valuePatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_ValuePatternId, valuePatternUnknown.put()), "text field value pattern lookup succeeds");
    Require(valuePatternUnknown != nullptr, "text field accessibility provider exposes value pattern");
    wil::com_ptr_nothrow<IValueProvider> valuePattern;
    RequireSucceeded(valuePatternUnknown.query_to(valuePattern.put()), "text field value pattern supports IValueProvider");
    RequireSucceeded(valuePattern->SetValue(L"beta"), "text field accessibility provider can set the value");
    Require(field->GetText() == L"beta", "text field accessibility SetValue updates the underlying DX control");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_ValueValuePropertyId, "text field exposes value after SetValue") == L"beta",
            "text field accessibility provider reports the updated value after SetValue");

    wil::com_ptr_nothrow<IRawElementProviderFragment> comboLabelProvider;
    RequireSucceeded(fieldProvider->Navigate(NavigateDirection_NextSibling, comboLabelProvider.put()),
                     "text field accessibility provider navigates to the combo label");
    Require(comboLabelProvider != nullptr, "text field accessibility provider returns the combo label as the next sibling");

    wil::com_ptr_nothrow<IRawElementProviderFragment> comboProvider;
    RequireSucceeded(comboLabelProvider->Navigate(NavigateDirection_NextSibling, comboProvider.put()),
                     "combo label accessibility provider navigates to the combo");
    Require(comboProvider != nullptr, "combo label accessibility provider returns the combo as the next sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> comboSimple;
    RequireSucceeded(comboProvider.query_to(comboSimple.put()), "combo accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*comboSimple.get(), UIA_ControlTypePropertyId, "combo exposes UIA control type") == UIA_ComboBoxControlTypeId,
            "editable combo accessibility provider reports combo-box control type");
    Require(ReadProviderStringProperty(*comboSimple.get(), UIA_NamePropertyId, "combo exposes accessibility name") == L"Mode",
            "editable combo accessibility provider uses its associated label as the accessible name");
    Require(ReadProviderStringProperty(*comboSimple.get(), UIA_ValueValuePropertyId, "combo exposes current value") == L"current",
            "editable combo accessibility provider reports the current value");

    wil::com_ptr_nothrow<IUnknown> comboValuePatternUnknown;
    RequireSucceeded(comboSimple->GetPatternProvider(UIA_ValuePatternId, comboValuePatternUnknown.put()), "combo value pattern lookup succeeds");
    Require(comboValuePatternUnknown != nullptr, "editable combo accessibility provider exposes value pattern");
    wil::com_ptr_nothrow<IValueProvider> comboValuePattern;
    RequireSucceeded(comboValuePatternUnknown.query_to(comboValuePattern.put()), "combo value pattern supports IValueProvider");
    RequireSucceeded(comboValuePattern->SetValue(L"updated"), "editable combo accessibility provider can set the value");
    Require(combo->GetText() == L"updated", "editable combo accessibility SetValue updates the underlying DX control");
    Require(ReadProviderStringProperty(*comboSimple.get(), UIA_ValueValuePropertyId, "combo exposes value after SetValue") == L"updated",
            "editable combo accessibility provider reports the updated value after SetValue");
}

void TestAccessibilityProviderRefreshesButtonSemanticProperties()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>();
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 140.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "button semantic refresh test creates an accessibility provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> buttonProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 24.0f, 16.0f, "button semantic refresh provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> buttonSimple;
    RequireSucceeded(buttonProvider.query_to(buttonSimple.put()), "button semantic refresh provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*buttonSimple.get(), UIA_NamePropertyId, "empty button exposes initial accessibility name").empty(),
            "button accessibility provider starts with an empty name");

    button->SetText(L"Cancel");
    Require(ReadProviderStringProperty(*buttonSimple.get(), UIA_NamePropertyId, "button text refresh updates accessibility name") == L"Cancel",
            "button accessibility provider refreshes its name after SetText");

    button->SetEnabled(false);
    Require(! ReadProviderBoolProperty(*buttonSimple.get(), UIA_IsEnabledPropertyId, "button enabled refresh updates accessibility state"),
            "button accessibility provider refreshes its enabled state after SetEnabled");

    button->SetAccessibleName(L"Dismiss");
    Require(ReadProviderStringProperty(*buttonSimple.get(), UIA_NamePropertyId, "explicit button accessible name refreshes") == L"Dismiss",
            "button accessibility provider refreshes explicit accessible name changes");
}

void TestAccessibilityProviderRefreshesLabelAssociations()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root     = std::make_unique<Panel>();
    auto* label   = root->AddChild<Label>(L"Ignore files");
    auto* toggle  = root->AddChild<Toggle>(L"Off");
    auto* pattern = root->AddChild<TextField>(L"*.tmp");
    label->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 24.0f));
    toggle->SetBounds(D2D1::RectF(0.0f, 32.0f, 160.0f, 64.0f));
    pattern->SetBounds(D2D1::RectF(0.0f, 72.0f, 220.0f, 104.0f));
    label->SetMnemonicTarget(toggle);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "label association refresh test creates an accessibility provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 88.0f, "label association field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "label association field provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_NamePropertyId, "unassociated field exposes fallback name") == L"*.tmp",
            "text field starts with its fallback accessible name while the label targets the toggle");

    label->SetMnemonicTarget(pattern);
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_NamePropertyId, "retargeted label updates field name") == L"Ignore files",
            "text field accessibility name refreshes when a label retargets to it");

    label->SetText(L"File patterns");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_NamePropertyId, "label text update refreshes associated field name") == L"File patterns",
            "text field accessibility name refreshes when its associated label text changes");
}

void TestAccessibilityProviderExposesDirectSemanticRootControls()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto combo = std::make_unique<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"current");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    auto* comboRaw = combo.get();
    window.Host().SetRoot(std::move(combo));
    window.Host().SetFocusControl(comboRaw);

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "direct-root accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> rootFragment;
    RequireSucceeded(rootProvider.query_to(rootFragment.put()), "direct-root accessibility root exposes IRawElementProviderFragment");

    wil::com_ptr_nothrow<IRawElementProviderSimple> rootSimple;
    RequireSucceeded(rootProvider.query_to(rootSimple.put()), "direct-root root provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*rootSimple.get(), UIA_ControlTypePropertyId, "direct-root combo exposes UIA control type") == UIA_ComboBoxControlTypeId,
            "direct-root combo accessibility provider reports combo-box control type");
    Require(ReadProviderStringProperty(*rootSimple.get(), UIA_ValueValuePropertyId, "direct-root combo exposes current value") == L"current",
            "direct-root combo accessibility provider reports the current value");

    wil::com_ptr_nothrow<IUnknown> comboValuePatternUnknown;
    RequireSucceeded(rootSimple->GetPatternProvider(UIA_ValuePatternId, comboValuePatternUnknown.put()), "direct-root combo value pattern lookup succeeds");
    Require(comboValuePatternUnknown != nullptr, "direct-root combo accessibility provider exposes value pattern");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstChildProvider;
    RequireSucceeded(rootFragment->Navigate(NavigateDirection_FirstChild, firstChildProvider.put()),
                     "direct-root collapsed provider first-child lookup succeeds");
    Require(firstChildProvider == nullptr, "direct-root collapsed provider does not expose the same semantic control as a duplicate child");

    wil::com_ptr_nothrow<IRawElementProviderFragment> hitProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 16.0f, "direct-root combo provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> hitRoot;
    RequireSucceeded(hitProvider.query_to(hitRoot.put()), "direct-root point-hit provider is the collapsed root provider");
    wil::com_ptr_nothrow<IRawElementProviderSimple> hitSimple;
    RequireSucceeded(hitProvider.query_to(hitSimple.put()), "direct-root point-hit provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*hitSimple.get(), UIA_ControlTypePropertyId, "direct-root point-hit provider exposes combo control type") ==
                UIA_ComboBoxControlTypeId,
            "direct-root point-hit provider resolves the combo control");

    wil::com_ptr_nothrow<IRawElementProviderFragment> focusedProvider;
    RequireSucceeded(rootProvider->GetFocus(focusedProvider.put()), "direct-root focus lookup succeeds");
    Require(focusedProvider != nullptr, "direct-root focus lookup returns the combo control");
    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> focusedRoot;
    RequireSucceeded(focusedProvider.query_to(focusedRoot.put()), "direct-root focused provider is the collapsed root provider");
    wil::com_ptr_nothrow<IRawElementProviderSimple> focusedSimple;
    RequireSucceeded(focusedProvider.query_to(focusedSimple.put()), "direct-root focused provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*focusedSimple.get(), UIA_ValueValuePropertyId, "direct-root focused combo exposes value") == L"current",
            "direct-root focus lookup returns the semantic root combo provider");
}

void TestAccessibilityLabelOnlyRootDoesNotUseDirectSemanticRootCollapse()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto label = std::make_unique<Label>(L"Standalone label");
    label->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 24.0f));
    window.Host().SetRoot(std::move(label));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "label-only accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> rootFragment;
    RequireSucceeded(rootProvider.query_to(rootFragment.put()), "label-only accessibility root exposes IRawElementProviderFragment");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstChildProvider;
    RequireSucceeded(rootFragment->Navigate(NavigateDirection_FirstChild, firstChildProvider.put()),
                     "label-only root first-child lookup succeeds");
    Require(firstChildProvider != nullptr, "label-only root exposes the label as a child instead of collapsing it into the root");

    wil::com_ptr_nothrow<IRawElementProviderSimple> childSimple;
    RequireSucceeded(firstChildProvider.query_to(childSimple.put()), "label-only child provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*childSimple.get(), UIA_ControlTypePropertyId, "label-only child exposes text control type") == UIA_TextControlTypeId,
            "label-only child provider is the label text element");
    Require(ReadProviderStringProperty(*childSimple.get(), UIA_NamePropertyId, "label-only child exposes label text") == L"Standalone label",
            "label-only child provider exposes the label text");

    wil::com_ptr_nothrow<IRawElementProviderFragment> duplicateGrandchild;
    RequireSucceeded(firstChildProvider->Navigate(NavigateDirection_FirstChild, duplicateGrandchild.put()),
                     "label-only child first-child lookup succeeds");
    Require(duplicateGrandchild == nullptr, "label-only label child does not expose a duplicate nested label");
}

void TestAccessibilityDirectSemanticRootMatchesUiAutomationClientTree()
{
    using namespace RedSalamander::DxUi;

    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    Require(coinitHr == S_OK || coinitHr == S_FALSE || coinitHr == RPC_E_CHANGED_MODE, "UIAutomation direct-root client test initializes COM");
    const bool uninitializeCom       = coinitHr == S_OK || coinitHr == S_FALSE;
    const auto uninitializeComOnExit = wil::scope_exit([&]
    {
        if (uninitializeCom)
        {
            CoUninitialize();
        }
    });

    AttachedHostWindow window;
    auto combo = std::make_unique<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"current");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    window.Host().SetRoot(std::move(combo));

    wil::com_ptr_nothrow<IUIAutomation> automation;
    RequireSucceeded(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(automation.put())),
                     "UIAutomation direct-root client is created");
    Require(automation != nullptr, "UIAutomation direct-root client instance is available");

    wil::com_ptr_nothrow<IUIAutomationElement> rootElement;
    RequireSucceeded(automation->ElementFromHandle(window.Hwnd(), rootElement.put()), "UIAutomation direct-root ElementFromHandle succeeds");
    Require(rootElement != nullptr, "UIAutomation direct-root ElementFromHandle returns an element");

    CONTROLTYPEID rootControlType = 0;
    RequireSucceeded(rootElement->get_CurrentControlType(&rootControlType), "UIAutomation direct-root current control type lookup succeeds");
    Require(rootControlType == UIA_ComboBoxControlTypeId, "UIAutomation direct-root element is the semantic combo box, not a wrapper pane");

    wil::com_ptr_nothrow<IUIAutomationValuePattern> valuePattern;
    RequireSucceeded(rootElement->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void()),
                     "UIAutomation direct-root value pattern lookup succeeds");
    Require(valuePattern != nullptr, "UIAutomation direct-root exposes the combo value pattern");

    BSTR currentValue = nullptr;
    RequireSucceeded(valuePattern->get_CurrentValue(&currentValue), "UIAutomation direct-root value lookup succeeds");
    const auto freeCurrentValue = wil::scope_exit([&] { SysFreeString(currentValue); });
    Require(std::wstring_view(currentValue ? currentValue : L"", SysStringLen(currentValue)) == L"current", "UIAutomation direct-root reports the combo value");

    BOOL isContentElement = FALSE;
    RequireSucceeded(rootElement->get_CurrentIsContentElement(&isContentElement), "UIAutomation direct-root content-element lookup succeeds");
    Require(isContentElement == TRUE, "UIAutomation direct-root element is a content element");

    VARIANT contentElementProperty{};
    VariantInit(&contentElementProperty);
    contentElementProperty.vt      = VT_BOOL;
    contentElementProperty.boolVal = VARIANT_TRUE;
    wil::com_ptr_nothrow<IUIAutomationCondition> contentViewCondition;
    RequireSucceeded(automation->CreatePropertyCondition(UIA_IsContentElementPropertyId, contentElementProperty, contentViewCondition.put()),
                     "UIAutomation direct-root content-view condition is created");
    Require(contentViewCondition != nullptr, "UIAutomation direct-root content-view condition instance is available");

    wil::com_ptr_nothrow<IUIAutomationElementArray> contentChildren;
    RequireSucceeded(rootElement->FindAll(TreeScope_Children, contentViewCondition.get(), contentChildren.put()),
                     "UIAutomation direct-root content-view child enumeration succeeds");
    Require(contentChildren != nullptr, "UIAutomation direct-root content-view child enumeration returns an array");

    int contentChildCount = -1;
    RequireSucceeded(contentChildren->get_Length(&contentChildCount), "UIAutomation direct-root content-view child count lookup succeeds");
    Require(contentChildCount == 0, "UIAutomation direct-root content view exposes exactly one content element with no duplicate child");
}

void TestAccessibilityDirectSemanticRootTreeSelectionMatchesUiAutomationClientTree()
{
    using namespace RedSalamander::DxUi;

    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    Require(coinitHr == S_OK || coinitHr == S_FALSE || coinitHr == RPC_E_CHANGED_MODE,
            "UIAutomation direct-root tree selection test initializes COM");
    const bool uninitializeCom = coinitHr == S_OK || coinitHr == S_FALSE;
    const auto uninitializeComOnExit = wil::scope_exit([&]
    {
        if (uninitializeCom)
        {
            CoUninitialize();
        }
    });

    AttachedHostWindow window;
    MutableTreeModel treeModel;
    treeModel.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Panes"},
        TreeItemData{.id = 3u, .text = L"Viewers"},
    });

    auto tree = std::make_unique<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    tree->SetModel(&treeModel);
    Tree* const liveTree = tree.get();
    window.Host().SetRoot(std::move(tree));
    liveTree->SetSelectedItemId(2u);

    wil::com_ptr_nothrow<IUIAutomation> automation;
    RequireSucceeded(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(automation.put())),
                     "UIAutomation direct-root tree selection client is created");
    Require(automation != nullptr, "UIAutomation direct-root tree selection client instance is available");

    wil::com_ptr_nothrow<IUIAutomationElement> rootElement;
    RequireSucceeded(automation->ElementFromHandle(window.Hwnd(), rootElement.put()), "UIAutomation direct-root tree ElementFromHandle succeeds");
    Require(rootElement != nullptr, "UIAutomation direct-root tree ElementFromHandle returns an element");

    CONTROLTYPEID rootControlType = 0;
    RequireSucceeded(rootElement->get_CurrentControlType(&rootControlType), "UIAutomation direct-root tree control type lookup succeeds");
    Require(rootControlType == UIA_TreeControlTypeId, "UIAutomation direct-root tree element is the semantic tree");

    wil::com_ptr_nothrow<IUIAutomationSelectionPattern> selectionPattern;
    RequireSucceeded(rootElement->GetCurrentPatternAs(UIA_SelectionPatternId, __uuidof(IUIAutomationSelectionPattern), selectionPattern.put_void()),
                     "UIAutomation direct-root tree exposes SelectionPattern");
    Require(selectionPattern != nullptr, "UIAutomation direct-root tree selection pattern is available");

    BOOL canSelectMultiple = TRUE;
    RequireSucceeded(selectionPattern->get_CurrentCanSelectMultiple(&canSelectMultiple),
                     "UIAutomation direct-root tree selection pattern reports multi-select capability");
    Require(canSelectMultiple == FALSE, "UIAutomation direct-root tree selection pattern reports single-selection behavior");

    wil::com_ptr_nothrow<IUIAutomationElementArray> selection;
    RequireSucceeded(selectionPattern->GetCurrentSelection(selection.put()), "UIAutomation direct-root tree selected-items lookup succeeds");
    Require(selection != nullptr, "UIAutomation direct-root tree selected-items lookup returns an array");

    int selectionLength = 0;
    RequireSucceeded(selection->get_Length(&selectionLength), "UIAutomation direct-root tree selected-items array reports length");
    Require(selectionLength == 1, "UIAutomation direct-root tree returns exactly one selected item");

    wil::com_ptr_nothrow<IUIAutomationElement> selectedElement;
    RequireSucceeded(selection->GetElement(0, selectedElement.put()), "UIAutomation direct-root tree selected item lookup succeeds");
    Require(selectedElement != nullptr, "UIAutomation direct-root tree selected item is available");

    CONTROLTYPEID selectedControlType = 0;
    RequireSucceeded(selectedElement->get_CurrentControlType(&selectedControlType),
                     "UIAutomation direct-root tree selected item control type lookup succeeds");
    Require(selectedControlType == UIA_TreeItemControlTypeId, "UIAutomation direct-root tree selected element is a tree item");

    BSTR selectedName = nullptr;
    RequireSucceeded(selectedElement->get_CurrentName(&selectedName), "UIAutomation direct-root tree selected item name lookup succeeds");
    const auto freeSelectedName = wil::scope_exit([&] { SysFreeString(selectedName); });
    Require(std::wstring_view(selectedName ? selectedName : L"", SysStringLen(selectedName)) == L"Panes",
            "UIAutomation direct-root tree selected item reports the selected visible item name");
}

void TestAccessibilityProviderReportsFocusedControl()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 140.0f, 32.0f));
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 40.0f, 220.0f, 68.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "focused-control accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> focusedProvider;
    RequireSucceeded(rootProvider->GetFocus(focusedProvider.put()), "root provider focus lookup succeeds");
    Require(focusedProvider != nullptr, "root provider returns the focused control");

    wil::com_ptr_nothrow<IRawElementProviderSimple> focusedSimple;
    RequireSucceeded(focusedProvider.query_to(focusedSimple.put()), "focused provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*focusedSimple.get(), UIA_ValueValuePropertyId, "focused text field exposes value") == L"alpha",
            "root provider focus lookup returns the focused text field provider");
}

void TestAccessibilityProviderMasksPasswordTextFieldValue()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetAccessibleName(L"Password");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "masked text field accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 16.0f, "masked text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "masked text field provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*fieldSimple.get(), UIA_ControlTypePropertyId, "masked text field exposes edit control type") == UIA_EditControlTypeId,
            "masked text field provider reports edit control type");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_NamePropertyId, "masked text field exposes accessible name") == L"Password",
            "masked text field provider does not use the secret value as its accessible name");
    Require(ReadProviderBoolProperty(*fieldSimple.get(), UIA_IsPasswordPropertyId, "masked text field exposes password state"),
            "masked text field provider reports IsPassword");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_ValueValuePropertyId, "masked text field suppresses UIA value").empty(),
            "masked text field provider does not expose the secret through ValuePattern");
}

void TestAccessibilityProviderExposesMaskedRevealButton()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetAccessibleName(L"Password");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "masked reveal accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> revealProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 206.0f, 16.0f, "masked reveal button provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> revealSimple;
    RequireSucceeded(revealProvider.query_to(revealSimple.put()), "masked reveal button provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*revealSimple.get(), UIA_ControlTypePropertyId, "masked reveal button exposes control type") == UIA_ButtonControlTypeId,
            "masked reveal button provider reports button control type");
    Require(ReadProviderStringProperty(*revealSimple.get(), UIA_NamePropertyId, "masked reveal button exposes accessible name") == L"Show password",
            "masked reveal button provider reports the reveal affordance name");

    wil::com_ptr_nothrow<IUnknown> invokePatternUnknown;
    RequireSucceeded(revealSimple->GetPatternProvider(UIA_InvokePatternId, invokePatternUnknown.put()), "masked reveal button invoke pattern lookup succeeds");
    Require(invokePatternUnknown != nullptr, "masked reveal button exposes InvokePattern");
    wil::com_ptr_nothrow<IInvokeProvider> invokePattern;
    RequireSucceeded(invokePatternUnknown.query_to(invokePattern.put()), "masked reveal button invoke pattern supports IInvokeProvider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider;
    RequireSucceeded(revealProvider->Navigate(NavigateDirection_Parent, fieldProvider.put()), "masked reveal button navigates to parent field");
    Require(fieldProvider != nullptr, "masked reveal button returns its owning field as parent");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "masked reveal parent provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*fieldSimple.get(), UIA_ControlTypePropertyId, "masked reveal parent exposes edit type") == UIA_EditControlTypeId,
            "masked reveal button parent is the edit field provider");

    RequireSucceeded(invokePattern->Invoke(), "masked reveal button Invoke succeeds");
    Require(field->GetPasswordRevealState() == PasswordRevealState::Visible, "masked reveal button Invoke reveals the field visually");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_ValueValuePropertyId, "revealed masked text field still suppresses UIA value").empty(),
            "masked text field provider still does not expose the secret after reveal Invoke");
}

void TestAccessibilityProviderExposesTextPatternForTextField()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root       = std::make_unique<Panel>();
    auto* textLabel = root->AddChild<Label>(L"Query");
    textLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    field->SetSelectionRange(0u, 5u);
    textLabel->SetMnemonicTarget(field);

    auto* passwordLabel = root->AddChild<Label>(L"Password");
    passwordLabel->SetBounds(D2D1::RectF(0.0f, 72.0f, 120.0f, 96.0f));
    auto* maskedField = root->AddChild<TextField>(L"secret");
    maskedField->SetMasked(true);
    maskedField->SetBounds(D2D1::RectF(0.0f, 100.0f, 260.0f, 132.0f));
    passwordLabel->SetMnemonicTarget(maskedField);

    auto* multilineLabel = root->AddChild<Label>(L"Notes");
    multilineLabel->SetBounds(D2D1::RectF(0.0f, 132.0f, 120.0f, 136.0f));
    auto* multilineField = root->AddChild<TextField>(L"red\ngreen\nblue");
    multilineField->SetMultiline(true);
    multilineField->SetBounds(D2D1::RectF(0.0f, 136.0f, 260.0f, 168.0f));
    multilineField->SetSelectionRange(0u, 3u);
    multilineLabel->SetMnemonicTarget(multilineField);

    std::wstring emojiText;
    emojiText.reserve(7u);
    emojiText.push_back(L'A');
    emojiText.push_back(static_cast<wchar_t>(0xD83D));
    emojiText.push_back(static_cast<wchar_t>(0xDC69));
    emojiText.push_back(static_cast<wchar_t>(0x200D));
    emojiText.push_back(static_cast<wchar_t>(0xD83D));
    emojiText.push_back(static_cast<wchar_t>(0xDCBB));
    emojiText.push_back(L'Z');
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "text pattern accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 44.0f, "text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "text field provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "text field TextPattern supports ITextProvider");

    SupportedTextSelection supportedSelection = SupportedTextSelection_None;
    RequireSucceeded(textPattern->get_SupportedTextSelection(&supportedSelection), "text field TextPattern reports supported selection mode");
    Require(supportedSelection == SupportedTextSelection_Single, "text field TextPattern supports one selection range");

    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "text field TextPattern exposes a document range");
    Require(documentRange != nullptr, "text field TextPattern returns a document range");
    Require(ReadTextRangeText(*documentRange.get(), -1, "text field document range exposes text") == L"alpha beta",
            "text field TextPattern document range returns the current value");
    wil::com_ptr_nothrow<ITextRangeProvider> clonedDocumentRange;
    RequireSucceeded(documentRange->Clone(clonedDocumentRange.put()), "text field TextPattern document range clones");
    Require(clonedDocumentRange != nullptr, "text field TextPattern document range clone is returned");
    BOOL sameRange = FALSE;
    RequireSucceeded(documentRange->Compare(clonedDocumentRange.get(), &sameRange), "text field TextPattern cloned range compares");
    Require(sameRange == TRUE, "text field TextPattern cloned range compares equal by content");
    int endpointComparison = 0;
    RequireSucceeded(
        documentRange->CompareEndpoints(TextPatternRangeEndpoint_Start, clonedDocumentRange.get(), TextPatternRangeEndpoint_End, &endpointComparison),
        "text field TextPattern endpoint comparison succeeds");
    Require(endpointComparison < 0, "text field TextPattern start endpoint compares before document end");
    wil::com_ptr_nothrow<ITextRangeProvider> wordRange;
    RequireSucceeded(documentRange->Clone(wordRange.put()), "text field TextPattern document range clones for word movement");
    Require(wordRange != nullptr, "text field TextPattern word movement range clone is returned");
    int moved = 0;
    RequireSucceeded(wordRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Word, 1, &moved),
                     "text field TextPattern start endpoint moves by word");
    Require(moved == 1, "text field TextPattern start endpoint reports moved words");
    Require(ReadTextRangeText(*wordRange.get(), -1, "text field document range exposes moved-word text") == L"beta",
            "text field TextPattern word endpoint movement narrows to the next word");
    moved = 0;
    RequireSucceeded(wordRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Word, -1, &moved),
                     "text field TextPattern start endpoint moves backward by word");
    Require(moved == -1, "text field TextPattern start endpoint reports moved words backward");
    Require(ReadTextRangeText(*wordRange.get(), -1, "text field document range restores after moved-word text") == L"alpha beta",
            "text field TextPattern word endpoint movement restores the document text");
    const POINT rangePointScreen = window.Host().DipPointToScreenPoint(D2D1::Point2F(0.0f, 44.0f));
    const UiaPoint rangePoint{static_cast<double>(rangePointScreen.x), static_cast<double>(rangePointScreen.y)};
    wil::com_ptr_nothrow<ITextRangeProvider> pointRange;
    RequireSucceeded(textPattern->RangeFromPoint(rangePoint, pointRange.put()), "text field TextPattern RangeFromPoint succeeds");
    Require(pointRange != nullptr, "text field TextPattern RangeFromPoint returns a range");
    Require(ReadTextRangeText(*pointRange.get(), -1, "text field RangeFromPoint range is collapsed").empty(),
            "text field TextPattern RangeFromPoint returns a collapsed caret range");
    moved = 0;
    wil::com_ptr_nothrow<ITextRangeProvider> wordCaretRange;
    RequireSucceeded(pointRange->Clone(wordCaretRange.put()), "text field RangeFromPoint range clones for word movement");
    Require(wordCaretRange != nullptr, "text field RangeFromPoint word movement clone is returned");
    RequireSucceeded(wordCaretRange->Move(TextUnit_Word, 1, &moved), "text field collapsed range moves forward by word");
    Require(moved == 1, "text field collapsed range reports moved words");
    RequireSucceeded(wordCaretRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, 1, &moved),
                     "text field word-moved collapsed range expands by one character");
    Require(ReadTextRangeText(*wordCaretRange.get(), -1, "text field word-moved collapsed range exposes text") == L"b",
            "text field collapsed word movement lands at the next word");
    moved = 0;
    RequireSucceeded(pointRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, 1, &moved),
                     "text field RangeFromPoint range endpoint expands by one character");
    Require(moved == 1, "text field RangeFromPoint range reports one expanded character");
    Require(ReadTextRangeText(*pointRange.get(), -1, "text field RangeFromPoint expanded range exposes text") == L"a",
            "text field TextPattern RangeFromPoint maps the leading point to the first character");
    moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 6, &moved),
                     "text field TextPattern start endpoint moves by character");
    Require(moved == 6, "text field TextPattern start endpoint reports moved characters");
    Require(ReadTextRangeText(*documentRange.get(), -1, "text field document range exposes moved-start text") == L"beta",
            "text field TextPattern start endpoint movement narrows the range");
    moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, -2, &moved),
                     "text field TextPattern end endpoint moves backward by character");
    Require(moved == -2, "text field TextPattern end endpoint reports moved characters");
    Require(ReadTextRangeText(*documentRange.get(), -1, "text field document range exposes moved-end text") == L"be",
            "text field TextPattern end endpoint movement narrows the range from the end");

    AttachedHostWindow emojiWindow;
    auto emojiRoot   = std::make_unique<Panel>();
    auto* emojiField = emojiRoot->AddChild<TextField>(emojiText);
    emojiField->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    emojiWindow.Host().SetRoot(std::move(emojiRoot));
    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> emojiRootProvider;
    emojiRootProvider.attach(emojiWindow.Host().DebugCreateAccessibilityProvider());
    Require(emojiRootProvider != nullptr, "emoji text pattern test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> emojiProvider =
        GetProviderAtDipPoint(emojiWindow.Hwnd(), emojiWindow.Host(), *emojiRootProvider.get(), 40.0f, 44.0f, "emoji text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> emojiSimple;
    RequireSucceeded(emojiProvider.query_to(emojiSimple.put()), "emoji text field provider exposes IRawElementProviderSimple");
    wil::com_ptr_nothrow<IUnknown> emojiTextPatternUnknown;
    RequireSucceeded(emojiSimple->GetPatternProvider(UIA_TextPatternId, emojiTextPatternUnknown.put()), "emoji text field TextPattern lookup succeeds");
    Require(emojiTextPatternUnknown != nullptr, "emoji text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> emojiTextPattern;
    RequireSucceeded(emojiTextPatternUnknown.query_to(emojiTextPattern.put()), "emoji text field TextPattern supports ITextProvider");
    wil::com_ptr_nothrow<ITextRangeProvider> emojiDocumentRange;
    RequireSucceeded(emojiTextPattern->get_DocumentRange(emojiDocumentRange.put()), "emoji text field TextPattern exposes a document range");
    Require(emojiDocumentRange != nullptr, "emoji text field TextPattern returns a document range");
    Require(ReadTextRangeText(*emojiDocumentRange.get(), -1, "emoji text field document range exposes text") == emojiText,
            "emoji text field TextPattern document range returns the current value");
    moved = 0;
    RequireSucceeded(emojiDocumentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 2, &moved),
                     "emoji text field TextPattern start endpoint moves by text elements");
    Require(moved == 2, "emoji text field TextPattern start endpoint reports moved text elements");
    Require(ReadTextRangeText(*emojiDocumentRange.get(), -1, "emoji text field document range exposes text-element moved text") ==
                emojiText.substr(emojiText.size() - 1u),
            "emoji text field TextPattern character movement treats the ZWJ emoji cluster as one text element");
    moved = 0;
    RequireSucceeded(emojiDocumentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, -1, &moved),
                     "emoji text field TextPattern start endpoint moves backward by text element");
    Require(moved == -1, "emoji text field TextPattern start endpoint reports moved text element backward");
    Require(ReadTextRangeText(*emojiDocumentRange.get(), -1, "emoji text field document range exposes backward text-element moved text") ==
                emojiText.substr(1u),
            "emoji text field TextPattern backward character movement restores the full ZWJ emoji cluster");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "text field TextPattern selection lookup succeeds");
    const auto destroySelectionRanges                      = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange = GetSingleTextRangeFromArray(selectionRanges, "text field TextPattern exposes one selection range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "text field selected range exposes text") == L"alpha",
            "text field TextPattern selection range exposes the retained selection");
    wil::com_ptr_nothrow<ITextRangeProvider> selectedWordRange;
    RequireSucceeded(selectedRange->Clone(selectedWordRange.put()), "text field selected range clones for word-range movement");
    Require(selectedWordRange != nullptr, "text field selected word movement range clone is returned");
    moved = 0;
    RequireSucceeded(selectedWordRange->Move(TextUnit_Word, 1, &moved), "text field selected range moves by word");
    Require(moved == 1, "text field selected range reports moved words");
    Require(ReadTextRangeText(*selectedWordRange.get(), -1, "text field selected range exposes word-moved text") == L"beta",
            "text field selected range word movement lands on the next word");
    moved = 0;
    RequireSucceeded(selectedRange->Move(TextUnit_Character, 1, &moved), "text field selected range moves by character");
    Require(moved == 1, "text field selected range reports moved characters");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "text field selected range exposes moved text") == L"lpha ",
            "text field selected range movement preserves range length");
    RequireSucceeded(selectedRange->Select(), "text field selected range Select succeeds");
    const std::optional<std::pair<size_t, size_t>> selectedAfterRangeSelect = field->GetSelectionRange();
    Require(selectedAfterRangeSelect.has_value(), "text field selected range Select applies a retained selection");
    Require(selectedAfterRangeSelect.value().first == 1u && selectedAfterRangeSelect.value().second == 6u,
            "text field selected range Select applies the UIA range to the retained TextField");
    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "text field selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles              = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues = ReadDoubleArray(selectedRectangles, "text field selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() >= 4u && selectedRectangleValues.size() % 4u == 0u,
            "text field selected range returns complete bounding rectangle tuples");
    Require(selectedRectangleValues[2] > 0.0 && selectedRectangleValues[3] > 0.0, "text field selected range returns a non-empty bounding rectangle");

    wil::com_ptr_nothrow<IUnknown> textEditPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextEditPatternId, textEditPatternUnknown.put()), "text field TextEditPattern lookup succeeds");
    Require(textEditPatternUnknown != nullptr, "text field exposes TextEditPattern");
    wil::com_ptr_nothrow<ITextEditProvider> textEditPattern;
    RequireSucceeded(textEditPatternUnknown.query_to(textEditPattern.put()), "text field TextEditPattern supports ITextEditProvider");
    wil::com_ptr_nothrow<ITextRangeProvider> activeComposition;
    RequireSucceeded(textEditPattern->GetActiveComposition(activeComposition.put()), "inactive TextEditPattern active-composition lookup succeeds");
    Require(activeComposition == nullptr, "inactive TextEditPattern has no active composition range");

    wil::com_ptr_nothrow<IRawElementProviderFragment> maskedProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 116.0f, "masked text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> maskedSimple;
    RequireSucceeded(maskedProvider.query_to(maskedSimple.put()), "masked text field provider exposes IRawElementProviderSimple");
    wil::com_ptr_nothrow<IUnknown> maskedTextPatternUnknown;
    RequireSucceeded(maskedSimple->GetPatternProvider(UIA_TextPatternId, maskedTextPatternUnknown.put()), "masked text field TextPattern lookup succeeds");
    Require(maskedTextPatternUnknown == nullptr, "masked text field does not expose TextPattern over protected content");
    wil::com_ptr_nothrow<IUnknown> maskedTextEditPatternUnknown;
    RequireSucceeded(maskedSimple->GetPatternProvider(UIA_TextEditPatternId, maskedTextEditPatternUnknown.put()),
                     "masked text field TextEditPattern lookup succeeds");
    Require(maskedTextEditPatternUnknown == nullptr, "masked text field does not expose TextEditPattern over protected content");

    wil::com_ptr_nothrow<IRawElementProviderFragment> multilineProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 148.0f, "multiline text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> multilineSimple;
    RequireSucceeded(multilineProvider.query_to(multilineSimple.put()), "multiline text field provider exposes IRawElementProviderSimple");
    wil::com_ptr_nothrow<IUnknown> multilineValuePatternUnknown;
    RequireSucceeded(multilineSimple->GetPatternProvider(UIA_ValuePatternId, multilineValuePatternUnknown.put()),
                     "multiline text field ValuePattern lookup succeeds");
    Require(multilineValuePatternUnknown == nullptr, "multiline text field does not expose ValuePattern");
    wil::com_ptr_nothrow<IUnknown> multilineTextPatternUnknown;
    RequireSucceeded(multilineSimple->GetPatternProvider(UIA_TextPatternId, multilineTextPatternUnknown.put()),
                     "multiline text field TextPattern lookup succeeds");
    Require(multilineTextPatternUnknown != nullptr, "multiline text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> multilineTextPattern;
    RequireSucceeded(multilineTextPatternUnknown.query_to(multilineTextPattern.put()), "multiline text field TextPattern supports ITextProvider");
    wil::com_ptr_nothrow<ITextRangeProvider> multilineDocumentRange;
    RequireSucceeded(multilineTextPattern->get_DocumentRange(multilineDocumentRange.put()), "multiline text field TextPattern exposes a document range");
    Require(multilineDocumentRange != nullptr, "multiline text field TextPattern returns a document range");
    moved = 0;
    RequireSucceeded(multilineDocumentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, 1, &moved),
                     "multiline text field start endpoint moves by line");
    Require(moved == 1, "multiline text field start endpoint reports moved lines");
    Require(ReadTextRangeText(*multilineDocumentRange.get(), -1, "multiline text field document range exposes moved-line text") == L"green\nblue",
            "multiline text field line endpoint movement narrows to the next logical line");
    moved = 0;
    RequireSucceeded(multilineDocumentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, -1, &moved),
                     "multiline text field start endpoint moves backward by line");
    Require(moved == -1, "multiline text field start endpoint reports moved lines backward");
    Require(ReadTextRangeText(*multilineDocumentRange.get(), -1, "multiline text field document range restores after moved-line text") == L"red\ngreen\nblue",
            "multiline text field line endpoint movement restores the document text");
    SAFEARRAY* multilineSelectionRanges = nullptr;
    RequireSucceeded(multilineTextPattern->GetSelection(&multilineSelectionRanges), "multiline text field TextPattern selection lookup succeeds");
    const auto destroyMultilineSelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(multilineSelectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> multilineSelectedRange =
        GetSingleTextRangeFromArray(multilineSelectionRanges, "multiline text field TextPattern exposes one selection range");
    moved = 0;
    RequireSucceeded(multilineSelectedRange->Move(TextUnit_Line, 1, &moved), "multiline selected range moves by line");
    Require(moved == 1, "multiline selected range reports moved lines");
    Require(ReadTextRangeText(*multilineSelectedRange.get(), -1, "multiline selected range exposes line-moved text") == L"green",
            "multiline selected range line movement lands on the next logical line");
}

void TestAccessibilityTextFieldSimpleRangeBoundingRectanglesUseCaretGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta gamma");
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 420.0f, 56.0f));
    field->SetSelectionRange(6u, 10u);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "simple range rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 80.0f, 40.0f, "simple text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "simple text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "simple text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "simple text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "simple text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "simple text field selection lookup succeeds");
    const auto destroySelectionRanges                      = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange = GetSingleTextRangeFromArray(selectionRanges, "simple text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "simple selected range exposes text") == L"beta",
            "simple selected range preserves logical selected text");

    D2D1_RECT_F startRectDip{};
    D2D1_RECT_F endRectDip{};
    Require(field->DebugGetCaretRect(window.Host(), 6u, startRectDip), "simple selected range measures the start caret");
    Require(field->DebugGetCaretRect(window.Host(), 10u, endRectDip), "simple selected range measures the end caret");
    const RECT expectedScreen = DipRectToScreenRect(window, UnionRects(startRectDip, endRectDip));

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "simple selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles              = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues = ReadDoubleArray(selectedRectangles, "simple selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == 4u, "simple selected range returns one caret-geometry rectangle tuple");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[0]),
                     static_cast<float>(expectedScreen.left),
                     1.0f,
                     "simple selected range rectangle follows the native start caret x");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[1]),
                     static_cast<float>(expectedScreen.top),
                     1.0f,
                     "simple selected range rectangle follows the native caret top");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[2]),
                     static_cast<float>(expectedScreen.right - expectedScreen.left),
                     1.0f,
                     "simple selected range rectangle width follows native caret geometry");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[3]),
                     static_cast<float>(expectedScreen.bottom - expectedScreen.top),
                     1.0f,
                     "simple selected range rectangle height follows native caret geometry");
}

void TestAccessibilityTextFieldMultilineRangeFromPointUsesNativeHitTest()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"one\ntwo three\nfour");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 280.0f, 112.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "multiline RangeFromPoint test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "multiline text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "multiline text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "multiline text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "multiline text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "multiline text field TextPattern supports ITextProvider");

    D2D1_RECT_F caretRectDip{};
    Require(field->DebugGetCaretRect(window.Host(), 8u, caretRectDip), "multiline RangeFromPoint test measures the target caret");
    const POINT queryPoint = window.Host().DipPointToScreenPoint(D2D1::Point2F(caretRectDip.left, (caretRectDip.top + caretRectDip.bottom) * 0.5f));
    const UiaPoint rangePoint{static_cast<double>(queryPoint.x), static_cast<double>(queryPoint.y)};

    wil::com_ptr_nothrow<ITextRangeProvider> pointRange;
    RequireSucceeded(textPattern->RangeFromPoint(rangePoint, pointRange.put()), "multiline text field RangeFromPoint succeeds");
    Require(pointRange != nullptr, "multiline text field RangeFromPoint returns a range");
    int moved = 0;
    RequireSucceeded(pointRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, 1, &moved),
                     "multiline RangeFromPoint caret range expands by one character");
    Require(moved == 1, "multiline RangeFromPoint caret range reports one expanded character");
    Require(ReadTextRangeText(*pointRange.get(), -1, "multiline RangeFromPoint expanded range exposes text") == L"t",
            "multiline RangeFromPoint maps the second-line point to the native logical ACP");
}

void TestAccessibilityTextRangeFromPointDispatchesToWindowThread()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"one\ntwo three\nfour");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 280.0f, 112.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "cross-thread RangeFromPoint test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider = GetProviderAtDipPoint(
        window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "cross-thread RangeFromPoint field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "cross-thread RangeFromPoint provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "cross-thread RangeFromPoint TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "cross-thread RangeFromPoint field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "cross-thread RangeFromPoint TextPattern supports ITextProvider");

    D2D1_RECT_F caretRectDip{};
    Require(field->DebugGetCaretRect(window.Host(), 8u, caretRectDip), "cross-thread RangeFromPoint test measures the target caret");
    const POINT queryPoint = window.Host().DipPointToScreenPoint(D2D1::Point2F(caretRectDip.left, (caretRectDip.top + caretRectDip.bottom) * 0.5f));
    const UiaPoint rangePoint{static_cast<double>(queryPoint.x), static_cast<double>(queryPoint.y)};

    constexpr HRESULT kPendingRange = E_PENDING;
    std::atomic<bool> workerStarted{false};
    std::atomic<HRESULT> rangeResult{kPendingRange};
    ITextRangeProvider* returnedRange = nullptr;
    std::thread worker([&]
    {
        workerStarted.store(true, std::memory_order_release);
        const HRESULT hr = textPattern->RangeFromPoint(rangePoint, &returnedRange);
        rangeResult.store(hr, std::memory_order_release);
    });

    const auto startDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(500);
    while (! workerStarted.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < startDeadline)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
    Require(workerStarted.load(std::memory_order_acquire), "cross-thread RangeFromPoint worker starts");

    std::this_thread::sleep_for(std::chrono::milliseconds(50));
    if (rangeResult.load(std::memory_order_acquire) != kPendingRange)
    {
        worker.join();
        Require(false, "cross-thread RangeFromPoint waits for host window-thread dispatch");
    }

    const auto rangeDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
    while (rangeResult.load(std::memory_order_acquire) == kPendingRange && std::chrono::steady_clock::now() < rangeDeadline)
    {
        window.PumpMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    const HRESULT result = rangeResult.load(std::memory_order_acquire);
    worker.join();
    Require(result != kPendingRange, "cross-thread RangeFromPoint completes after host window-thread dispatch");
    RequireSucceeded(result, "cross-thread RangeFromPoint succeeds after dispatch");
    wil::com_ptr_nothrow<ITextRangeProvider> pointRange;
    pointRange.attach(returnedRange);
    Require(pointRange != nullptr, "cross-thread RangeFromPoint returns a range after dispatch");

    int moved = 0;
    RequireSucceeded(pointRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, 1, &moved),
                     "cross-thread RangeFromPoint returned range expands by one character");
    Require(moved == 1, "cross-thread RangeFromPoint range expands by one character");
    const std::wstring expandedText = ReadTextRangeText(*pointRange.get(), -1, "cross-thread RangeFromPoint expanded range exposes text");
    Require(expandedText == L"t", "cross-thread RangeFromPoint maps the second-line point to the native logical ACP");
}

void TestAccessibilityTextFieldMultilineSameLineRangeBoundingRectanglesUseCaretGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"one\ntwo three\nfour");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 280.0f, 112.0f));
    field->SetSelectionRange(4u, 13u);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "multiline same-line rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "multiline same-line text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "multiline same-line text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "multiline same-line text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "multiline same-line text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "multiline same-line text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "multiline same-line text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "multiline same-line text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "multiline same-line selected range exposes text") == L"two three",
            "multiline same-line selected range preserves logical selected text");

    D2D1_RECT_F startRectDip{};
    D2D1_RECT_F endRectDip{};
    Require(field->DebugGetCaretRect(window.Host(), 4u, startRectDip), "multiline same-line selected range measures the start caret");
    Require(field->DebugGetCaretRect(window.Host(), 13u, endRectDip), "multiline same-line selected range measures the end caret");
    const RECT expectedScreen = DipRectToScreenRect(window, UnionRects(startRectDip, endRectDip));

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "multiline same-line selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "multiline same-line selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == 4u, "multiline same-line selected range returns one caret-geometry rectangle tuple");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[0]),
                     static_cast<float>(expectedScreen.left),
                     1.0f,
                     "multiline same-line selected range rectangle follows the native start caret x");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[1]),
                     static_cast<float>(expectedScreen.top),
                     1.0f,
                     "multiline same-line selected range rectangle follows the native caret top");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[2]),
                     static_cast<float>(expectedScreen.right - expectedScreen.left),
                     1.0f,
                     "multiline same-line selected range rectangle width follows native caret geometry");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[3]),
                     static_cast<float>(expectedScreen.bottom - expectedScreen.top),
                     1.0f,
                     "multiline same-line selected range rectangle height follows native caret geometry");
}

void TestAccessibilityTextFieldMultilineRangeBoundingRectanglesUseLineCaretGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"one\ntwo three\nfour");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 280.0f, 112.0f));
    field->SetSelectionRange(1u, 7u);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "multiline cross-line rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "multiline cross-line text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "multiline cross-line text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "multiline cross-line text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "multiline cross-line text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "multiline cross-line text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "multiline cross-line text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "multiline cross-line text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "multiline cross-line selected range exposes text") == L"ne\ntwo",
            "multiline cross-line selected range preserves logical selected text");

    D2D1_RECT_F firstStartRectDip{};
    D2D1_RECT_F firstEndRectDip{};
    D2D1_RECT_F secondStartRectDip{};
    D2D1_RECT_F secondEndRectDip{};
    Require(field->DebugGetCaretRect(window.Host(), 1u, firstStartRectDip), "multiline cross-line selected range measures first-line start caret");
    Require(field->DebugGetCaretRect(window.Host(), 3u, firstEndRectDip), "multiline cross-line selected range measures first-line end caret");
    Require(field->DebugGetCaretRect(window.Host(), 4u, secondStartRectDip), "multiline cross-line selected range measures second-line start caret");
    Require(field->DebugGetCaretRect(window.Host(), 7u, secondEndRectDip), "multiline cross-line selected range measures second-line end caret");
    const std::array<RECT, 2> expectedScreenRects{
        DipRectToScreenRect(window, UnionRects(firstStartRectDip, firstEndRectDip)),
        DipRectToScreenRect(window, UnionRects(secondStartRectDip, secondEndRectDip)),
    };

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "multiline cross-line selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "multiline cross-line selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == 8u, "multiline cross-line selected range returns one rectangle tuple per logical line");
    for (size_t rectIndex = 0u; rectIndex < expectedScreenRects.size(); ++rectIndex)
    {
        const size_t valueIndex = rectIndex * 4u;
        const RECT& expected    = expectedScreenRects[rectIndex];
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex]),
                         static_cast<float>(expected.left),
                         1.0f,
                         "multiline cross-line selected range rectangle follows the native line start caret x");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 1u]),
                         static_cast<float>(expected.top),
                         1.0f,
                         "multiline cross-line selected range rectangle follows the native line caret top");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 2u]),
                         static_cast<float>(expected.right - expected.left),
                         1.0f,
                         "multiline cross-line selected range rectangle width follows native line caret geometry");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 3u]),
                         static_cast<float>(expected.bottom - expected.top),
                         1.0f,
                         "multiline cross-line selected range rectangle height follows native line caret geometry");
    }
}

void TestAccessibilityTextFieldWrappedRangeBoundingRectanglesUseVisualLineGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    constexpr std::wstring_view kWrappedText = L"alpha beta gamma delta epsilon zeta eta theta";
    auto root                                = std::make_unique<Panel>();
    auto* field                              = root->AddChild<TextField>(std::wstring(kWrappedText));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 150.0f, 140.0f));
    field->SetSelectionRange(0u, kWrappedText.size());
    window.Host().SetRoot(std::move(root));

    TextFieldDebugMultilineState multilineState{};
    Require(field->DebugGetMultilineState(window.Host(), multilineState), "wrapped selected range reads multiline debug state");
    Require(multilineState.totalLineCount > 1u, "wrapped selected range fixture wraps onto multiple visual lines");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "wrapped rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "wrapped text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "wrapped text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "wrapped text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "wrapped text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "wrapped text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "wrapped text field selection lookup succeeds");
    const auto destroySelectionRanges                      = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange = GetSingleTextRangeFromArray(selectionRanges, "wrapped text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "wrapped selected range exposes text") == kWrappedText,
            "wrapped selected range preserves the logical selected text");

    D2D1_RECT_F firstCaretDip{};
    D2D1_RECT_F lastCaretDip{};
    Require(field->DebugGetCaretRect(window.Host(), 0u, firstCaretDip), "wrapped selected range measures the first caret");
    Require(field->DebugGetCaretRect(window.Host(), kWrappedText.size(), lastCaretDip), "wrapped selected range measures the final caret");
    const RECT firstCaretScreen = DipRectToScreenRect(window, firstCaretDip);
    const RECT lastCaretScreen  = DipRectToScreenRect(window, lastCaretDip);
    Require(firstCaretScreen.top != lastCaretScreen.top, "wrapped selected range starts and ends on different visual lines");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "wrapped selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles              = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues = ReadDoubleArray(selectedRectangles, "wrapped selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() >= 8u, "wrapped selected range returns multiple rectangle tuples");
    Require(selectedRectangleValues.size() % 4u == 0u, "wrapped selected range returns complete rectangle tuples");

    RequireFloatNear(static_cast<float>(selectedRectangleValues[1]),
                     static_cast<float>(firstCaretScreen.top),
                     1.0f,
                     "wrapped selected range first rectangle follows the first visual-line caret top");
    const size_t lastTuple = selectedRectangleValues.size() - 4u;
    RequireFloatNear(static_cast<float>(selectedRectangleValues[lastTuple + 1u]),
                     static_cast<float>(lastCaretScreen.top),
                     1.0f,
                     "wrapped selected range last rectangle follows the final visual-line caret top");
}

void TestAccessibilityTextFieldWrappedCrossLineRangeBoundingRectanglesUseVisualLineGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    constexpr std::wstring_view kWrappedText = L"alpha beta gamma delta epsilon zeta\nomega";
    auto root                                = std::make_unique<Panel>();
    auto* field                              = root->AddChild<TextField>(std::wstring(kWrappedText));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 150.0f, 150.0f));
    field->SetSelectionRange(0u, kWrappedText.size());
    window.Host().SetRoot(std::move(root));

    TextFieldDebugMultilineState multilineState{};
    Require(field->DebugGetMultilineState(window.Host(), multilineState), "wrapped cross-line range reads multiline debug state");
    Require(multilineState.totalLineCount > 2u, "wrapped cross-line range fixture wraps one logical line and includes a newline");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "wrapped cross-line rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "wrapped cross-line text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "wrapped cross-line text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "wrapped cross-line text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "wrapped cross-line text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "wrapped cross-line text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "wrapped cross-line text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "wrapped cross-line text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "wrapped cross-line selected range exposes text") == kWrappedText,
            "wrapped cross-line selected range preserves the logical selected text");

    D2D1_RECT_F firstCaretDip{};
    D2D1_RECT_F lastCaretDip{};
    Require(field->DebugGetCaretRect(window.Host(), 0u, firstCaretDip), "wrapped cross-line selected range measures the first caret");
    Require(field->DebugGetCaretRect(window.Host(), kWrappedText.size(), lastCaretDip), "wrapped cross-line selected range measures the final caret");
    const RECT firstCaretScreen = DipRectToScreenRect(window, firstCaretDip);
    const RECT lastCaretScreen  = DipRectToScreenRect(window, lastCaretDip);
    Require(firstCaretScreen.top != lastCaretScreen.top, "wrapped cross-line selected range starts and ends on different visual lines");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "wrapped cross-line selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "wrapped cross-line selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() >= 12u, "wrapped cross-line selected range returns visual-line tuples across the newline");
    Require(selectedRectangleValues.size() % 4u == 0u, "wrapped cross-line selected range returns complete rectangle tuples");

    RequireFloatNear(static_cast<float>(selectedRectangleValues[1]),
                     static_cast<float>(firstCaretScreen.top),
                     1.0f,
                     "wrapped cross-line first rectangle follows the first visual-line caret top");
    const size_t lastTuple = selectedRectangleValues.size() - 4u;
    RequireFloatNear(static_cast<float>(selectedRectangleValues[lastTuple + 1u]),
                     static_cast<float>(lastCaretScreen.top),
                     1.0f,
                     "wrapped cross-line last rectangle follows the final visual-line caret top");
}

void TestAccessibilityTextFieldWrappedLineMovementUsesVisualLines()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    const std::wstring text = L"alpha beta gamma delta epsilon zeta";
    auto root               = std::make_unique<Panel>();
    auto* field             = root->AddChild<TextField>(text);
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 118.0f, 112.0f));
    window.Host().SetRoot(std::move(root));

    const std::vector<size_t> visualLineStarts =
        ResolveVisualLineStarts(window.Host(), *field, text, "wrapped line movement test resolves native visual-line starts");
    Require(visualLineStarts.size() >= 3u, "wrapped line movement fixture creates at least three visual lines");
    const size_t secondLineStart = visualLineStarts[1];
    const size_t thirdLineStart  = visualLineStarts[2];
    Require(secondLineStart > 0u && secondLineStart < thirdLineStart && thirdLineStart < text.size(),
            "wrapped line movement fixture exposes stable wrapped visual-line boundaries");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "wrapped line movement test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "wrapped line movement field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "wrapped line movement provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "wrapped line movement TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "wrapped line movement field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "wrapped line movement TextPattern supports ITextProvider");

    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "wrapped line movement document range lookup succeeds");
    int moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, 1, &moved),
                     "wrapped line movement start endpoint moves by visual line");
    Require(moved == 1, "wrapped line movement endpoint reports one visual line");
    Require(ReadTextRangeText(*documentRange.get(), -1, "wrapped line movement document range exposes moved text") == text.substr(secondLineStart),
            "wrapped line movement endpoint lands on the second visual line");

    field->SetSelectionRange(0u, secondLineStart);
    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "wrapped line movement selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "wrapped line movement TextPattern exposes one selection range");
    moved = 0;
    RequireSucceeded(selectedRange->Move(TextUnit_Line, 1, &moved), "wrapped selected range moves by visual line");
    Require(moved == 1, "wrapped selected range reports one visual line");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "wrapped selected range exposes moved visual-line text") ==
                text.substr(secondLineStart, thirdLineStart - secondLineStart),
            "wrapped selected range lands on the next visual line span");
}

void TestAccessibilityTextRangeEndpointLineMovementDispatchesToWindowThread()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    const std::wstring text = L"alpha beta gamma delta epsilon zeta";
    auto root               = std::make_unique<Panel>();
    auto* field             = root->AddChild<TextField>(text);
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 118.0f, 112.0f));
    window.Host().SetRoot(std::move(root));

    const std::vector<size_t> visualLineStarts =
        ResolveVisualLineStarts(window.Host(), *field, text, "cross-thread endpoint line movement resolves native visual-line starts");
    Require(visualLineStarts.size() >= 2u, "cross-thread endpoint line movement fixture creates wrapped visual lines");
    const size_t secondLineStart = visualLineStarts[1];

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "cross-thread endpoint line movement creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider = GetProviderAtDipPoint(
        window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "cross-thread endpoint line movement field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "cross-thread endpoint line movement provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "cross-thread endpoint line movement TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "cross-thread endpoint line movement field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "cross-thread endpoint line movement TextPattern supports ITextProvider");

    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "cross-thread endpoint line movement TextPattern exposes a document range");

    constexpr HRESULT kPendingMove = E_PENDING;
    std::atomic<bool> workerStarted{false};
    std::atomic<HRESULT> moveResult{kPendingMove};
    std::atomic<int> movedResult{0};
    std::thread worker([&]
    {
        int moved = 0;
        workerStarted.store(true, std::memory_order_release);
        const HRESULT hr = documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, 1, &moved);
        movedResult.store(moved, std::memory_order_release);
        moveResult.store(hr, std::memory_order_release);
    });

    const auto startDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(500);
    while (! workerStarted.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < startDeadline)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
    Require(workerStarted.load(std::memory_order_acquire), "cross-thread TextRange endpoint line movement worker starts");

    std::this_thread::sleep_for(std::chrono::milliseconds(50));
    if (moveResult.load(std::memory_order_acquire) != kPendingMove)
    {
        worker.join();
        Require(false, "cross-thread TextRange endpoint line movement waits for host window-thread dispatch");
    }

    const auto moveDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
    while (moveResult.load(std::memory_order_acquire) == kPendingMove && std::chrono::steady_clock::now() < moveDeadline)
    {
        window.PumpMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    const HRESULT result = moveResult.load(std::memory_order_acquire);
    worker.join();
    Require(result != kPendingMove, "cross-thread TextRange endpoint line movement completes after host window-thread dispatch");
    RequireSucceeded(result, "cross-thread TextRange endpoint line movement succeeds after dispatch");
    Require(movedResult.load(std::memory_order_acquire) == 1, "cross-thread TextRange endpoint line movement reports one visual line");
    Require(ReadTextRangeText(*documentRange.get(), -1, "cross-thread endpoint line movement exposes moved text") == text.substr(secondLineStart),
            "cross-thread TextRange endpoint line movement lands on the second visual line");
}

void TestAccessibilityTextRangeSpanLineMovementDispatchesToWindowThread()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    const std::wstring text = L"alpha beta gamma delta epsilon zeta";
    auto root               = std::make_unique<Panel>();
    auto* field             = root->AddChild<TextField>(text);
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 118.0f, 112.0f));
    window.Host().SetRoot(std::move(root));

    const std::vector<size_t> visualLineStarts =
        ResolveVisualLineStarts(window.Host(), *field, text, "cross-thread span line movement resolves native visual-line starts");
    Require(visualLineStarts.size() >= 3u, "cross-thread span line movement fixture creates wrapped visual lines");
    const size_t secondLineStart = visualLineStarts[1];
    const size_t thirdLineStart  = visualLineStarts[2];

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "cross-thread span line movement creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider = GetProviderAtDipPoint(
        window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "cross-thread span line movement field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "cross-thread span line movement provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "cross-thread span line movement TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "cross-thread span line movement field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "cross-thread span line movement TextPattern supports ITextProvider");

    field->SetSelectionRange(0u, secondLineStart);
    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "cross-thread span line movement selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "cross-thread span line movement TextPattern exposes one selection range");

    constexpr HRESULT kPendingMove = E_PENDING;
    std::atomic<bool> workerStarted{false};
    std::atomic<HRESULT> moveResult{kPendingMove};
    std::atomic<int> movedResult{0};
    std::thread worker([&]
    {
        int moved = 0;
        workerStarted.store(true, std::memory_order_release);
        const HRESULT hr = selectedRange->Move(TextUnit_Line, 1, &moved);
        movedResult.store(moved, std::memory_order_release);
        moveResult.store(hr, std::memory_order_release);
    });

    const auto startDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(500);
    while (! workerStarted.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < startDeadline)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
    Require(workerStarted.load(std::memory_order_acquire), "cross-thread TextRange span line movement worker starts");

    std::this_thread::sleep_for(std::chrono::milliseconds(50));
    if (moveResult.load(std::memory_order_acquire) != kPendingMove)
    {
        worker.join();
        Require(false, "cross-thread TextRange span line movement waits for host window-thread dispatch");
    }

    const auto moveDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
    while (moveResult.load(std::memory_order_acquire) == kPendingMove && std::chrono::steady_clock::now() < moveDeadline)
    {
        window.PumpMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    const HRESULT result = moveResult.load(std::memory_order_acquire);
    worker.join();
    Require(result != kPendingMove, "cross-thread TextRange span line movement completes after host window-thread dispatch");
    RequireSucceeded(result, "cross-thread TextRange span line movement succeeds after dispatch");
    Require(movedResult.load(std::memory_order_acquire) == 1, "cross-thread TextRange span line movement reports one visual line");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "cross-thread span line movement exposes moved text") ==
                text.substr(secondLineStart, thirdLineStart - secondLineStart),
            "cross-thread TextRange span line movement lands on the next visual line span");
}

void TestAccessibilityTextFieldSingleLineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root = std::make_unique<Panel>();
    root->SetFlowDirection(FlowDirection::RightToLeft);
    auto* field = root->AddChild<TextField>(L"abc \x05D0\x05D1\x05D2 123");
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 380.0f, 56.0f));
    field->SetSelectionRange(4u, 7u);
    window.Host().SetRoot(std::move(root));

    const std::optional<std::vector<D2D1_RECT_F>> expectedRectsDip = field->TryGetTextInputRangeRects(window.Host(), 4u, 7u);
    Require(expectedRectsDip.has_value() && ! expectedRectsDip->empty(), "single-line mixed-BiDi selected range has retained DirectWrite range geometry");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "single-line mixed-BiDi range rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider = GetProviderAtDipPoint(
        window.Hwnd(), window.Host(), *rootProvider.get(), 80.0f, 40.0f, "single-line mixed-BiDi text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "single-line mixed-BiDi text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "single-line mixed-BiDi text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "single-line mixed-BiDi text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "single-line mixed-BiDi text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "single-line mixed-BiDi text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "single-line mixed-BiDi text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "single-line mixed-BiDi selected range exposes text") == L"\x05D0\x05D1\x05D2",
            "single-line mixed-BiDi selected range preserves logical UTF-16 selected text");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "single-line mixed-BiDi selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "single-line mixed-BiDi selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == expectedRectsDip->size() * 4u,
            "single-line mixed-BiDi selected range returns the retained DirectWrite rectangle tuple count");
    for (size_t rectIndex = 0u; rectIndex < expectedRectsDip->size(); ++rectIndex)
    {
        const RECT expected     = DipRectToScreenRect(window, expectedRectsDip->at(rectIndex));
        const size_t valueIndex = rectIndex * 4u;
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex]),
                         static_cast<float>(expected.left),
                         1.0f,
                         "single-line mixed-BiDi selected range rectangle uses DirectWrite left edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 1u]),
                         static_cast<float>(expected.top),
                         1.0f,
                         "single-line mixed-BiDi selected range rectangle uses DirectWrite top edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 2u]),
                         static_cast<float>(expected.right - expected.left),
                         1.0f,
                         "single-line mixed-BiDi selected range rectangle uses DirectWrite width");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 3u]),
                         static_cast<float>(expected.bottom - expected.top),
                         1.0f,
                         "single-line mixed-BiDi selected range rectangle uses DirectWrite height");
    }
}

void TestAccessibilityTextFieldMultilineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root = std::make_unique<Panel>();
    root->SetFlowDirection(FlowDirection::RightToLeft);
    auto* field = root->AddChild<TextField>(L"latin \x05D0\x05D1\x05D2 span\nsecond line");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 300.0f, 112.0f));
    field->SetSelectionRange(6u, 9u);
    window.Host().SetRoot(std::move(root));

    const std::optional<std::vector<D2D1_RECT_F>> expectedRectsDip = field->TryGetTextInputRangeRects(window.Host(), 6u, 9u);
    Require(expectedRectsDip.has_value() && ! expectedRectsDip->empty(), "multiline mixed-BiDi selected range has retained DirectWrite range geometry");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "multiline mixed-BiDi range rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 80.0f, 40.0f, "multiline mixed-BiDi text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "multiline mixed-BiDi text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "multiline mixed-BiDi text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "multiline mixed-BiDi text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "multiline mixed-BiDi text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "multiline mixed-BiDi text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "multiline mixed-BiDi text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "multiline mixed-BiDi selected range exposes text") == L"\x05D0\x05D1\x05D2",
            "multiline mixed-BiDi selected range preserves logical UTF-16 selected text");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "multiline mixed-BiDi selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "multiline mixed-BiDi selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == expectedRectsDip->size() * 4u,
            "multiline mixed-BiDi selected range returns the retained DirectWrite rectangle tuple count");

    for (size_t index = 0u; index < expectedRectsDip->size(); ++index)
    {
        const RECT expectedScreen = DipRectToScreenRect(window, expectedRectsDip.value()[index]);
        const size_t valueIndex   = index * 4u;
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 0u]),
                         static_cast<float>(expectedScreen.left),
                         1.0f,
                         "multiline mixed-BiDi selected range rectangle uses DirectWrite left edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 1u]),
                         static_cast<float>(expectedScreen.top),
                         1.0f,
                         "multiline mixed-BiDi selected range rectangle uses DirectWrite top edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 2u]),
                         static_cast<float>(expectedScreen.right - expectedScreen.left),
                         1.0f,
                         "multiline mixed-BiDi selected range rectangle uses DirectWrite width");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 3u]),
                         static_cast<float>(expectedScreen.bottom - expectedScreen.top),
                         1.0f,
                         "multiline mixed-BiDi selected range rectangle uses DirectWrite height");
    }
}

void TestAccessibilityEditableComboBoxSingleLineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root = std::make_unique<Panel>();
    root->SetFlowDirection(FlowDirection::RightToLeft);
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"abc \x05D0\x05D1\x05D2 123");
    combo->SetEditableSelectionRange(4u, 7u);
    combo->SetBounds(D2D1::RectF(20.0f, 24.0f, 380.0f, 56.0f));
    window.Host().SetRoot(std::move(root));

    const std::optional<std::vector<D2D1_RECT_F>> expectedRectsDip = combo->TryGetTextInputRangeRects(window.Host(), 4u, 7u);
    Require(expectedRectsDip.has_value() && ! expectedRectsDip->empty(),
            "editable combo single-line mixed-BiDi selected range has retained DirectWrite range geometry");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "editable combo mixed-BiDi range rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> comboProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 80.0f, 40.0f, "editable combo mixed-BiDi provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> comboSimple;
    RequireSucceeded(comboProvider.query_to(comboSimple.put()), "editable combo mixed-BiDi provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(comboSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "editable combo mixed-BiDi TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "editable combo mixed-BiDi exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "editable combo mixed-BiDi TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "editable combo mixed-BiDi selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "editable combo mixed-BiDi exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "editable combo mixed-BiDi selected range exposes text") == L"\x05D0\x05D1\x05D2",
            "editable combo mixed-BiDi selected range preserves logical UTF-16 selected text");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "editable combo mixed-BiDi selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "editable combo mixed-BiDi selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == expectedRectsDip->size() * 4u,
            "editable combo mixed-BiDi selected range returns the retained DirectWrite rectangle tuple count");

    for (size_t rectIndex = 0u; rectIndex < expectedRectsDip->size(); ++rectIndex)
    {
        const RECT expected     = DipRectToScreenRect(window, expectedRectsDip->at(rectIndex));
        const size_t valueIndex = rectIndex * 4u;
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex]),
                         static_cast<float>(expected.left),
                         1.0f,
                         "editable combo mixed-BiDi selected range rectangle uses DirectWrite left edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 1u]),
                         static_cast<float>(expected.top),
                         1.0f,
                         "editable combo mixed-BiDi selected range rectangle uses DirectWrite top edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 2u]),
                         static_cast<float>(expected.right - expected.left),
                         1.0f,
                         "editable combo mixed-BiDi selected range rectangle uses DirectWrite width");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 3u]),
                         static_cast<float>(expected.bottom - expected.top),
                         1.0f,
                         "editable combo mixed-BiDi selected range rectangle uses DirectWrite height");
    }
}

void TestAccessibilityProviderExposesTextPatternForEditableComboBox()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root        = std::make_unique<Panel>();
    auto* comboLabel = root->AddChild<Label>(L"Mode");
    comboLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"current");
    combo->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    comboLabel->SetMnemonicTarget(combo);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "editable combo text pattern test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> comboProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 44.0f, "editable combo provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> comboSimple;
    RequireSucceeded(comboProvider.query_to(comboSimple.put()), "editable combo provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(comboSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "editable combo TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "editable combo exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "editable combo TextPattern supports ITextProvider");

    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "editable combo TextPattern exposes a document range");
    Require(documentRange != nullptr, "editable combo TextPattern returns a document range");
    Require(ReadTextRangeText(*documentRange.get(), -1, "editable combo document range exposes text") == L"current",
            "editable combo TextPattern document range returns editable text");

    const D2D1_RECT_F editableTextRect = combo->DebugGetEditableTextRect();
    const POINT rangePointScreen =
        window.Host().DipPointToScreenPoint(D2D1::Point2F(editableTextRect.left + 1.0f, (editableTextRect.top + editableTextRect.bottom) * 0.5f));
    const UiaPoint rangePoint{static_cast<double>(rangePointScreen.x), static_cast<double>(rangePointScreen.y)};
    wil::com_ptr_nothrow<ITextRangeProvider> pointRange;
    RequireSucceeded(textPattern->RangeFromPoint(rangePoint, pointRange.put()), "editable combo TextPattern RangeFromPoint succeeds");
    Require(pointRange != nullptr, "editable combo TextPattern RangeFromPoint returns a range");
    Require(ReadTextRangeText(*pointRange.get(), -1, "editable combo RangeFromPoint range is collapsed").empty(),
            "editable combo TextPattern RangeFromPoint returns a collapsed caret range");
    int moved = 0;
    RequireSucceeded(pointRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, 1, &moved),
                     "editable combo RangeFromPoint range endpoint expands by one character");
    Require(moved == 1, "editable combo RangeFromPoint range reports one expanded character");
    Require(ReadTextRangeText(*pointRange.get(), -1, "editable combo RangeFromPoint expanded range exposes text") == L"c",
            "editable combo TextPattern RangeFromPoint maps the editable text point to the first character");

    Require(combo->OnSelectAll(window.Host()), "editable combo can select all before UIA selection lookup");
    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "editable combo TextPattern selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "editable combo TextPattern exposes one selection range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "editable combo selected range exposes text") == L"current",
            "editable combo TextPattern selection range exposes the retained editable selection");
    moved = 0;
    RequireSucceeded(selectedRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 2, &moved),
                     "editable combo selected range start moves by character");
    Require(moved == 2, "editable combo selected range start reports moved characters");
    moved = 0;
    RequireSucceeded(selectedRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, -2, &moved),
                     "editable combo selected range end moves backward by character");
    Require(moved == -2, "editable combo selected range end reports moved characters");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "editable combo selected range exposes narrowed text") == L"rre",
            "editable combo selected range endpoint movement narrows the range");
    RequireSucceeded(selectedRange->Select(), "editable combo selected range Select succeeds");

    SAFEARRAY* appliedSelectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&appliedSelectionRanges), "editable combo TextPattern selection lookup succeeds after Select");
    const auto destroyAppliedSelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(appliedSelectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> appliedSelectedRange =
        GetSingleTextRangeFromArray(appliedSelectionRanges, "editable combo TextPattern exposes one applied selection range");
    Require(ReadTextRangeText(*appliedSelectedRange.get(), -1, "editable combo applied selected range exposes text") == L"rre",
            "editable combo TextPattern Select applies the UIA range to retained editable selection");

    wil::com_ptr_nothrow<IUnknown> textEditPatternUnknown;
    RequireSucceeded(comboSimple->GetPatternProvider(UIA_TextEditPatternId, textEditPatternUnknown.put()), "editable combo TextEditPattern lookup succeeds");
    Require(textEditPatternUnknown != nullptr, "editable combo exposes TextEditPattern");
}

void TestAccessibilityTextRangeSelectDispatchesToWindowThread()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    field->SetSelectionRange(0u, 5u);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "cross-thread text range test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 44.0f, "cross-thread text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "cross-thread text field provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "cross-thread text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "cross-thread text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "cross-thread text field TextPattern supports ITextProvider");
    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "cross-thread text field TextPattern exposes a document range");

    int moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 6, &moved),
                     "cross-thread text range start endpoint moves to beta");
    Require(moved == 6, "cross-thread text range start endpoint reports moved characters");

    constexpr HRESULT kPendingSelect = E_PENDING;
    std::atomic<bool> workerStarted{false};
    std::atomic<HRESULT> selectResult{kPendingSelect};
    std::thread worker([&]
    {
        workerStarted.store(true, std::memory_order_release);
        selectResult.store(documentRange->Select(), std::memory_order_release);
    });

    const auto startDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(500);
    while (! workerStarted.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < startDeadline)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
    Require(workerStarted.load(std::memory_order_acquire), "cross-thread TextRange Select worker starts");

    std::this_thread::sleep_for(std::chrono::milliseconds(50));
    if (selectResult.load(std::memory_order_acquire) != kPendingSelect)
    {
        worker.join();
        Require(false, "cross-thread TextRange Select waits for host window-thread dispatch");
    }

    const auto selectDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
    while (selectResult.load(std::memory_order_acquire) == kPendingSelect && std::chrono::steady_clock::now() < selectDeadline)
    {
        window.PumpMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    const HRESULT result = selectResult.load(std::memory_order_acquire);
    worker.join();
    Require(result != kPendingSelect, "cross-thread TextRange Select completes after host window-thread dispatch");
    RequireSucceeded(result, "cross-thread TextRange Select succeeds after dispatch");

    const std::optional<std::pair<size_t, size_t>> selectedRange = field->GetSelectionRange();
    Require(selectedRange.has_value(), "cross-thread TextRange Select applies a retained selection");
    Require(selectedRange.value().first == 6u && selectedRange.value().second == 10u,
            "cross-thread TextRange Select applies the range on the host window thread");
}

struct AccessibilityTextRangeSelectFixture
{
    explicit AccessibilityTextRangeSelectFixture(AttachedHostWindow& window)
    {
        using namespace RedSalamander::DxUi;

        auto root = std::make_unique<Panel>();
        field     = root->AddChild<TextField>(L"alpha beta");
        field->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
        field->SetSelectionRange(0u, 5u);
        window.Host().SetRoot(std::move(root));

        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "UIA Select dispatch fixture creates a root provider");
        fieldProvider = GetProviderAtDipPoint(
            window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 44.0f, "UIA Select dispatch fixture resolves the text field provider");
        RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "UIA Select dispatch fixture exposes IRawElementProviderSimple");
        RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "UIA Select dispatch fixture gets TextPattern");
        Require(textPatternUnknown != nullptr, "UIA Select dispatch fixture exposes TextPattern");
        RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "UIA Select dispatch fixture gets ITextProvider");
        RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "UIA Select dispatch fixture gets the document range");

        int moved = 0;
        RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 6, &moved),
                         "UIA Select dispatch fixture moves the range start to beta");
        Require(moved == 6, "UIA Select dispatch fixture reports the moved range start");
    }

    RedSalamander::DxUi::TextField* field = nullptr;
    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider;
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
};

void TestAccessibilityTimedOutTextRangeSelectDoesNotExecuteLater()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    AccessibilityTextRangeSelectFixture fixture(window);

    wil::unique_event_nothrow posted;
    posted.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(posted != nullptr, "timed-out Select test creates the posted event");
    wil::unique_event_nothrow handlerEntered;
    handlerEntered.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(handlerEntered != nullptr, "timed-out Select test creates the handler-entered event");
    wil::unique_event_nothrow releaseHandler;
    releaseHandler.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(releaseHandler != nullptr, "timed-out Select test creates the release event");

    DebugResetAccessibilityUiActionExecutionCountForTest();
    DebugSetAccessibilityUiActionDispatchTimeoutForTest(100u);
    DebugSetAccessibilityUiActionPostedEventForTest(posted.get());
    DebugSetAccessibilityUiActionHandlerStallForTest(handlerEntered.get(), releaseHandler.get());
    const auto clearHooks = wil::scope_exit([]() noexcept
    {
        DebugSetAccessibilityUiActionHandlerStallForTest(nullptr, nullptr);
        DebugSetAccessibilityUiActionPostedEventForTest(nullptr);
        DebugSetAccessibilityUiActionDispatchTimeoutForTest(0u);
    });

    constexpr HRESULT kPending = E_PENDING;
    std::atomic<HRESULT> result{kPending};
    std::atomic<bool> enteredObserved{false};
    std::thread worker([&] { result.store(fixture.documentRange->Select(), std::memory_order_release); });

    Require(WaitForSingleObject(posted.get(), 2000u) == WAIT_OBJECT_0, "timed-out Select request is posted");
    std::thread releaser([&]
    {
        enteredObserved.store(WaitForSingleObject(handlerEntered.get(), 2000u) == WAIT_OBJECT_0, std::memory_order_release);
        const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
        while (result.load(std::memory_order_acquire) == kPending && std::chrono::steady_clock::now() < deadline)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
        }
        static_cast<void>(SetEvent(releaseHandler.get()));
    });

    window.PumpMessages();
    worker.join();
    releaser.join();

    Require(enteredObserved.load(std::memory_order_acquire), "timed-out Select handler reaches the pre-take stall");
    Require(result.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_TIMEOUT), "timed-out Select reports ERROR_TIMEOUT");
    Require(DebugGetAccessibilityUiActionExecutionCountForTest() == 0u, "abandoned Select is not executed after its caller times out");
    const auto selection = fixture.field->GetSelectionRange();
    Require(selection == std::optional<std::pair<size_t, size_t>>{{0u, 5u}}, "abandoned Select leaves the retained selection unchanged");
}

void TestAccessibilityTakenTextRangeSelectExecutesOnlyOnce()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    AccessibilityTextRangeSelectFixture fixture(window);

    wil::unique_event_nothrow posted;
    posted.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(posted != nullptr, "taken Select test creates the posted event");
    wil::unique_event_nothrow handlerTaken;
    handlerTaken.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(handlerTaken != nullptr, "taken Select test creates the handler-taken event");
    wil::unique_event_nothrow releaseHandler;
    releaseHandler.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(releaseHandler != nullptr, "taken Select test creates the release event");

    DebugResetAccessibilityUiActionExecutionCountForTest();
    DebugSetAccessibilityUiActionDispatchTimeoutForTest(100u);
    DebugSetAccessibilityUiActionPostedEventForTest(posted.get());
    DebugSetAccessibilityUiActionHandlerTakenStallForTest(handlerTaken.get(), releaseHandler.get());
    const auto clearHooks = wil::scope_exit([]() noexcept
    {
        DebugSetAccessibilityUiActionHandlerTakenStallForTest(nullptr, nullptr);
        DebugSetAccessibilityUiActionPostedEventForTest(nullptr);
        DebugSetAccessibilityUiActionDispatchTimeoutForTest(0u);
    });

    constexpr HRESULT kPending = E_PENDING;
    std::atomic<HRESULT> result{kPending};
    std::atomic<bool> takenObserved{false};
    std::thread worker([&] { result.store(fixture.documentRange->Select(), std::memory_order_release); });

    Require(WaitForSingleObject(posted.get(), 2000u) == WAIT_OBJECT_0, "taken Select request is posted");
    std::thread releaser([&]
    {
        takenObserved.store(WaitForSingleObject(handlerTaken.get(), 2000u) == WAIT_OBJECT_0, std::memory_order_release);
        const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
        while (result.load(std::memory_order_acquire) == kPending && std::chrono::steady_clock::now() < deadline)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
        }
        static_cast<void>(SetEvent(releaseHandler.get()));
    });

    window.PumpMessages();
    worker.join();
    releaser.join();

    Require(takenObserved.load(std::memory_order_acquire), "taken Select handler owns the dispatch before timeout");
    Require(result.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_TIMEOUT), "taken-but-incomplete Select reports ERROR_TIMEOUT");
    Require(DebugGetAccessibilityUiActionExecutionCountForTest() == 1u, "taken Select executes exactly once after the timeout race");
    const auto selection = fixture.field->GetSelectionRange();
    Require(selection == std::optional<std::pair<size_t, size_t>>{{6u, 10u}}, "taken Select applies its range exactly once");
}

void TestAccessibilityDestroyWithPendingDispatchReturnsCancelled()
{
    using namespace RedSalamander::DxUi;

    auto window = std::make_unique<AttachedHostWindow>();
    AccessibilityTextRangeSelectFixture fixture(*window);

    wil::unique_event_nothrow posted;
    posted.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(posted != nullptr, "destroy-pending Select test creates the posted event");
    DebugSetAccessibilityUiActionPostedEventForTest(posted.get());
    const auto clearHook = wil::scope_exit([]() noexcept { DebugSetAccessibilityUiActionPostedEventForTest(nullptr); });

    constexpr HRESULT kPending = E_PENDING;
    std::atomic<HRESULT> result{kPending};
    std::thread worker([&] { result.store(fixture.documentRange->Select(), std::memory_order_release); });

    Require(WaitForSingleObject(posted.get(), 2000u) == WAIT_OBJECT_0, "destroy-pending Select request is posted");
    window.reset();
    worker.join();

    Require(result.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
            "destroying a window drains its pending Select request with ERROR_CANCELLED");
}

void TestAccessibilityTextRangeBoundingRectanglesDispatchesToWindowThread()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "cross-thread text range bounds test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 44.0f, "cross-thread bounds text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "cross-thread bounds text field provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "cross-thread bounds TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "cross-thread bounds text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "cross-thread bounds TextPattern supports ITextProvider");

    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "cross-thread bounds TextPattern exposes a document range");

    int moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 6, &moved),
                     "cross-thread bounds range start endpoint moves to beta");
    Require(moved == 6, "cross-thread bounds range start endpoint reports moved characters");

    constexpr HRESULT kPendingBounds = E_PENDING;
    std::atomic<bool> workerStarted{false};
    std::atomic<HRESULT> boundsResult{kPendingBounds};
    SAFEARRAY* workerRectangles = nullptr;
    std::thread worker([&]
    {
        workerStarted.store(true, std::memory_order_release);
        boundsResult.store(documentRange->GetBoundingRectangles(&workerRectangles), std::memory_order_release);
    });

    const auto startDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(500);
    while (! workerStarted.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < startDeadline)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
    Require(workerStarted.load(std::memory_order_acquire), "cross-thread TextRange bounds worker starts");

    std::this_thread::sleep_for(std::chrono::milliseconds(50));
    if (boundsResult.load(std::memory_order_acquire) != kPendingBounds)
    {
        worker.join();
        const auto destroyEarlyRectangles = wil::scope_exit([&]
        {
            if (workerRectangles)
            {
                SafeArrayDestroy(workerRectangles);
            }
        });
        Require(false, "cross-thread TextRange GetBoundingRectangles waits for host window-thread dispatch");
    }

    const auto boundsDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
    while (boundsResult.load(std::memory_order_acquire) == kPendingBounds && std::chrono::steady_clock::now() < boundsDeadline)
    {
        window.PumpMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    const HRESULT result = boundsResult.load(std::memory_order_acquire);
    worker.join();
    const auto destroyRectangles = wil::scope_exit([&]
    {
        if (workerRectangles)
        {
            SafeArrayDestroy(workerRectangles);
        }
    });
    Require(result != kPendingBounds, "cross-thread TextRange GetBoundingRectangles completes after host window-thread dispatch");
    RequireSucceeded(result, "cross-thread TextRange GetBoundingRectangles succeeds after dispatch");

    const std::vector<double> rectangleValues = ReadDoubleArray(workerRectangles, "cross-thread moved range returns bounding rectangle values");
    Require(rectangleValues.size() == 4u, "cross-thread moved range returns one bounding rectangle");
    Require(rectangleValues[2] > 0.0 && rectangleValues[3] > 0.0, "cross-thread moved range rectangle is non-empty");
}

void TestAccessibilityTextRangeBoundingRectanglesTimeoutKeepsLateHandlerStorageAlive()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "late-handler text range bounds test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 44.0f, "late-handler bounds text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "late-handler bounds text field provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "late-handler bounds TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "late-handler bounds text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "late-handler bounds TextPattern supports ITextProvider");

    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "late-handler bounds TextPattern exposes a document range");

    int moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 6, &moved),
                     "late-handler bounds range start endpoint moves to beta");
    Require(moved == 6, "late-handler bounds range start endpoint reports moved characters");

    wil::unique_event_nothrow handlerEntered;
    handlerEntered.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(handlerEntered != nullptr, "late-handler bounds test creates handler-entered event");
    wil::unique_event_nothrow releaseHandler;
    releaseHandler.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    Require(releaseHandler != nullptr, "late-handler bounds test creates release event");

    DebugSetAccessibilityUiActionHandlerStallForTest(handlerEntered.get(), releaseHandler.get());
    DebugSetAccessibilityUiActionDispatchTimeoutForTest(100u);
    const auto clearStallHook = wil::scope_exit([]() noexcept
    {
        DebugSetAccessibilityUiActionHandlerStallForTest(nullptr, nullptr);
        DebugSetAccessibilityUiActionDispatchTimeoutForTest(0u);
    });

    constexpr HRESULT kPendingBounds = E_PENDING;
    constexpr HRESULT kExpectedTimeout = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    std::atomic<bool> workerStarted{false};
    std::atomic<bool> handlerEnteredObserved{false};
    std::atomic<bool> timeoutObservedBeforeRelease{false};
    std::atomic<HRESULT> boundsResult{kPendingBounds};
    SAFEARRAY* workerRectangles = nullptr;

    std::thread worker([&]
    {
        workerStarted.store(true, std::memory_order_release);
        boundsResult.store(documentRange->GetBoundingRectangles(&workerRectangles), std::memory_order_release);
    });

    std::thread releaser([&]
    {
        const DWORD entered = ::WaitForSingleObject(handlerEntered.get(), 2000u);
        handlerEnteredObserved.store(entered == WAIT_OBJECT_0, std::memory_order_release);
        if (entered == WAIT_OBJECT_0)
        {
            const auto timeoutDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(7);
            while (boundsResult.load(std::memory_order_acquire) == kPendingBounds && std::chrono::steady_clock::now() < timeoutDeadline)
            {
                std::this_thread::sleep_for(std::chrono::milliseconds(1));
            }
            timeoutObservedBeforeRelease.store(boundsResult.load(std::memory_order_acquire) == kExpectedTimeout, std::memory_order_release);
        }
        static_cast<void>(::SetEvent(releaseHandler.get()));
    });

    const auto startDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(500);
    while (! workerStarted.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < startDeadline)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
    Require(workerStarted.load(std::memory_order_acquire), "late-handler TextRange bounds worker starts");

    const auto completionDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(9);
    while (boundsResult.load(std::memory_order_acquire) == kPendingBounds && std::chrono::steady_clock::now() < completionDeadline)
    {
        window.PumpMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    if (worker.joinable())
    {
        worker.join();
    }
    if (releaser.joinable())
    {
        releaser.join();
    }
    const auto destroyRectangles = wil::scope_exit([&]
    {
        if (workerRectangles)
        {
            SafeArrayDestroy(workerRectangles);
        }
    });

    Require(handlerEnteredObserved.load(std::memory_order_acquire), "late-handler UIA action is dequeued before the sender times out");
    Require(timeoutObservedBeforeRelease.load(std::memory_order_acquire), "late-handler UIA action sender times out before handler completion");
    Require(boundsResult.load(std::memory_order_acquire) == kExpectedTimeout,
            "late-handler TextRange GetBoundingRectangles reports timeout instead of reading late output");
    Require(workerRectangles == nullptr, "late-handler timeout leaves caller SAFEARRAY output untouched");
}

void TestAccessibilityProviderExposesNativeImeTextEditRanges()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(5u, 5u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputImePayload payload;
    payload.hasCompositionString  = true;
    payload.compositionString     = L"-ime";
    payload.compositionAttributes = {ATTR_INPUT, ATTR_TARGET_CONVERTED, ATTR_TARGET_CONVERTED, ATTR_INPUT};
    payload.hasCursorPosition     = true;
    payload.cursorPosition        = 3u;
    window.Host().DebugSetNativeTextInputImePayloadForTest(payload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_COMPSTR | GCS_COMPATTR | GCS_CURSORPOS));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "native ime TextEditPattern test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 16.0f, "native ime text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "native ime text field provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textEditPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextEditPatternId, textEditPatternUnknown.put()), "native ime TextEditPattern lookup succeeds");
    Require(textEditPatternUnknown != nullptr, "native ime text field exposes TextEditPattern");
    wil::com_ptr_nothrow<ITextEditProvider> textEditPattern;
    RequireSucceeded(textEditPatternUnknown.query_to(textEditPattern.put()), "native ime TextEditPattern supports ITextEditProvider");

    wil::com_ptr_nothrow<ITextRangeProvider> activeComposition;
    RequireSucceeded(textEditPattern->GetActiveComposition(activeComposition.put()), "native ime active-composition range lookup succeeds");
    Require(activeComposition != nullptr, "native ime TextEditPattern exposes active composition range");
    Require(ReadTextRangeText(*activeComposition.get(), -1, "native ime active-composition range exposes text") == L"-ime",
            "native ime active-composition range returns the preview string");

    wil::com_ptr_nothrow<ITextRangeProvider> conversionTarget;
    RequireSucceeded(textEditPattern->GetConversionTarget(conversionTarget.put()), "native ime conversion-target range lookup succeeds");
    Require(conversionTarget != nullptr, "native ime TextEditPattern exposes conversion target range");
    Require(ReadTextRangeText(*conversionTarget.get(), -1, "native ime conversion-target range exposes text") == L"im",
            "native ime conversion-target range returns the target-converted span");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_ENDCOMPOSITION, 0, 0));
}

void TestAccessibilityNativeTextInputRaisesTextAndTextEditEventCounters()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInput(field);

    const NativeTextInputEventCounters baselineCounters = window.Host().DebugGetNativeTextInputEventCounters();

    field->SetTextAndNotify(L"alpha beta edited");
    window.Host().SyncTextInput(field);

    NativeTextInputEventCounters counters = window.Host().DebugGetNativeTextInputEventCounters();
    Require(counters.uiaTextChangedCount == baselineCounters.uiaTextChangedCount + 1u,
            "native text input raises a UIA TextPattern text-changed event for retained text mutations");

    const NativeTextInputEventCounters afterTextCounters = counters;
    field->SetSelectionRange(6u, 10u);
    window.Host().SyncTextInput(field);
    counters = window.Host().DebugGetNativeTextInputEventCounters();
    Require(counters.uiaTextSelectionChangedCount == afterTextCounters.uiaTextSelectionChangedCount + 1u,
            "native text input raises a UIA TextPattern selection-changed event for retained selection mutations");

    const NativeTextInputEventCounters afterSelectionCounters = counters;
    field->SetSelectionRange(3u, 3u);
    window.Host().SyncTextInput(field);
    counters = window.Host().DebugGetNativeTextInputEventCounters();
    Require(counters.uiaActiveTextPositionChangedCount == afterSelectionCounters.uiaActiveTextPositionChangedCount + 1u,
            "native text input raises a UIA active text position event for retained caret moves");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputImePayload payload;
    payload.hasCompositionString  = true;
    payload.compositionString     = L"-ime";
    payload.compositionAttributes = {ATTR_INPUT, ATTR_TARGET_CONVERTED, ATTR_TARGET_CONVERTED, ATTR_INPUT};
    window.Host().DebugSetNativeTextInputImePayloadForTest(payload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_COMPSTR | GCS_COMPATTR));

    counters = window.Host().DebugGetNativeTextInputEventCounters();
    Require(counters.uiaTextEditTextChangedCount >= baselineCounters.uiaTextEditTextChangedCount + 1u,
            "native IME composition raises a UIA TextEdit text-changed event");
    Require(counters.uiaTextEditConversionTargetChangedCount == baselineCounters.uiaTextEditConversionTargetChangedCount + 1u,
            "native IME target conversion raises a UIA TextEdit conversion-target-changed event");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_ENDCOMPOSITION, 0, 0));
}

void TestAccessibilityGridSnapshotRebuildMeetsTenThousandRowSelectionBudget()
{
    using namespace RedSalamander::DxUi;

    constexpr size_t kRowCount = 10'000u;
    MultiRowGridModel gridModel(kRowCount);
    AttachedHostWindow window;
    auto grid = std::make_unique<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid->SetModel(&gridModel);
    Grid* const liveGrid = grid.get();
    window.Host().SetRoot(std::move(grid));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "10k-row grid snapshot test creates an accessibility provider");

    std::vector<uint64_t> selectedRowIds(kRowCount);
    std::iota(selectedRowIds.begin(), selectedRowIds.end(), uint64_t{0u});
    liveGrid->GetSelectionModel().SetRange(selectedRowIds, selectedRowIds.front(), selectedRowIds.back());

    DebugSetAccessibilityOffscreenSelectedRowMaterializationLimitForTest(kRowCount);
    const auto resetMaterializationLimit = wil::scope_exit([]() noexcept
    { DebugSetAccessibilityOffscreenSelectedRowMaterializationLimitForTest(0u); });
    const auto baselineStarted = std::chrono::steady_clock::now();
    liveGrid->RefreshAccessibilitySnapshot();
    const auto baselineElapsed = std::chrono::steady_clock::now() - baselineStarted;

    DebugSetAccessibilityOffscreenSelectedRowMaterializationLimitForTest(256u);
    const auto candidateStarted = std::chrono::steady_clock::now();
    liveGrid->RefreshAccessibilitySnapshot();
    const auto candidateElapsed = std::chrono::steady_clock::now() - candidateStarted;

    Require(liveGrid->GetSelectionModel().GetCount() == kRowCount, "10k-row grid snapshot retains the complete Ctrl+A selection");
    Require(candidateElapsed < std::chrono::milliseconds(250), "10k-row grid Ctrl+A accessibility snapshot rebuild stays under the 250 ms Debug budget");
    Require(candidateElapsed < baselineElapsed, "capped offscreen row materialization improves the 10k-row Ctrl+A snapshot rebuild");
}

void TestAccessibilityProviderExposesTreeAndGridMetadata()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root = std::make_unique<Panel>();

    auto* treeLabel = root->AddChild<Label>(L"Categories");
    treeLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 28.0f, 240.0f, 88.0f));
    MutableTreeModel treeModel;
    treeModel.SetVisibleItems({
        RedSalamander::DxUi::TreeItemData{.id = 1u, .text = L"General"},
        RedSalamander::DxUi::TreeItemData{.id = 2u, .text = L"Panes"},
        RedSalamander::DxUi::TreeItemData{.id = 3u, .text = L"Viewers"},
    });
    tree->SetModel(&treeModel);
    tree->SetSelectedItemId(2u);
    treeLabel->SetMnemonicTarget(tree);

    auto* gridLabel = root->AddChild<Label>(L"Results");
    gridLabel->SetBounds(D2D1::RectF(0.0f, 92.0f, 120.0f, 116.0f));
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 120.0f, 240.0f, 188.0f));
    MultiRowGridModel gridModel(6u);
    grid->SetModel(&gridModel);
    gridLabel->SetMnemonicTarget(grid);

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "tree/grid accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> treeLabelProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 48.0f, 12.0f, "tree label accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderFragment> treeProvider;
    RequireSucceeded(treeLabelProvider->Navigate(NavigateDirection_NextSibling, treeProvider.put()), "tree label accessibility provider navigates to the tree");
    Require(treeProvider != nullptr, "tree label accessibility provider returns the tree as the next sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> treeSimple;
    RequireSucceeded(treeProvider.query_to(treeSimple.put()), "tree accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*treeSimple.get(), UIA_ControlTypePropertyId, "tree exposes UIA control type") == UIA_TreeControlTypeId,
            "tree accessibility provider reports tree control type");
    Require(ReadProviderStringProperty(*treeSimple.get(), UIA_NamePropertyId, "tree exposes accessibility name") == L"Categories",
            "tree accessibility provider uses its associated label as the accessible name");

    wil::com_ptr_nothrow<IRawElementProviderFragment> treeItemProvider;
    RequireSucceeded(treeProvider->Navigate(NavigateDirection_FirstChild, treeItemProvider.put()),
                     "tree accessibility provider navigates to the first visible tree item");
    Require(treeItemProvider != nullptr, "tree accessibility provider returns a first tree-item child");
    wil::com_ptr_nothrow<IRawElementProviderSimple> treeItemSimple;
    RequireSucceeded(treeItemProvider.query_to(treeItemSimple.put()), "tree item accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*treeItemSimple.get(), UIA_ControlTypePropertyId, "tree item exposes UIA control type") == UIA_TreeItemControlTypeId,
            "tree item accessibility provider reports tree-item control type");
    Require(ReadProviderStringProperty(*treeItemSimple.get(), UIA_NamePropertyId, "tree item exposes accessibility name") == L"General",
            "tree item accessibility provider exposes the visible item text as its accessible name");
    Require(ReadProviderLongProperty(*treeItemSimple.get(), UIA_LevelPropertyId, "tree item exposes depth level") == 1,
            "tree item accessibility provider exposes a 1-based tree level");
    Require(! ReadProviderBoolProperty(*treeItemSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "first tree item exposes selected state"),
            "tree item accessibility provider reports the unselected first item");

    wil::com_ptr_nothrow<IRawElementProviderFragment> selectedTreeItemProvider;
    RequireSucceeded(treeItemProvider->Navigate(NavigateDirection_NextSibling, selectedTreeItemProvider.put()),
                     "tree item accessibility provider navigates to the next visible tree item");
    Require(selectedTreeItemProvider != nullptr, "tree item accessibility provider returns the next sibling item");
    wil::com_ptr_nothrow<IRawElementProviderSimple> selectedTreeItemSimple;
    RequireSucceeded(selectedTreeItemProvider.query_to(selectedTreeItemSimple.put()),
                     "selected tree item accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*selectedTreeItemSimple.get(), UIA_NamePropertyId, "selected tree item exposes accessibility name") == L"Panes",
            "tree item accessibility provider exposes the selected visible item text");
    Require(ReadProviderBoolProperty(*selectedTreeItemSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "selected tree item exposes selected state"),
            "tree item accessibility provider reports the selected item");

    wil::com_ptr_nothrow<IRawElementProviderFragment> hitTreeItemProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 48.0f, 70.0f, "tree item accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> hitTreeItemSimple;
    RequireSucceeded(hitTreeItemProvider.query_to(hitTreeItemSimple.put()), "tree point-hit provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*hitTreeItemSimple.get(), UIA_ControlTypePropertyId, "tree point-hit provider exposes item control type") ==
                UIA_TreeItemControlTypeId,
            "tree hit-testing resolves the visible tree item provider instead of only the tree container");
    Require(ReadProviderStringProperty(*hitTreeItemSimple.get(), UIA_NamePropertyId, "tree point-hit provider exposes item name") == L"Panes",
            "tree hit-testing resolves the expected visible tree item provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> gridLabelProvider;
    RequireSucceeded(treeProvider->Navigate(NavigateDirection_NextSibling, gridLabelProvider.put()), "tree accessibility provider navigates to the grid label");
    Require(gridLabelProvider != nullptr, "tree accessibility provider returns the grid label as the next sibling");

    wil::com_ptr_nothrow<IRawElementProviderFragment> gridProvider;
    RequireSucceeded(gridLabelProvider->Navigate(NavigateDirection_NextSibling, gridProvider.put()), "grid label accessibility provider navigates to the grid");
    Require(gridProvider != nullptr, "grid label accessibility provider returns the grid as the next sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> gridSimple;
    RequireSucceeded(gridProvider.query_to(gridSimple.put()), "grid accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*gridSimple.get(), UIA_ControlTypePropertyId, "grid exposes UIA control type") == UIA_DataGridControlTypeId,
            "grid accessibility provider reports data-grid control type");
    Require(ReadProviderStringProperty(*gridSimple.get(), UIA_NamePropertyId, "grid exposes accessibility name") == L"Results",
            "grid accessibility provider uses its associated label as the accessible name");
    Require(ReadProviderLongProperty(*gridSimple.get(), UIA_GridRowCountPropertyId, "grid exposes row count") == 6,
            "grid accessibility provider reports model row count");
    Require(ReadProviderLongProperty(*gridSimple.get(), UIA_GridColumnCountPropertyId, "grid exposes column count") == 1,
            "grid accessibility provider reports model column count");

    window.Host().SetFocusControl(tree);
    wil::com_ptr_nothrow<IRawElementProviderFragment> focusedProvider;
    RequireSucceeded(rootProvider->GetFocus(focusedProvider.put()), "root provider focus lookup succeeds for the tree");
    Require(focusedProvider != nullptr, "root provider returns the focused tree-item provider");
    wil::com_ptr_nothrow<IRawElementProviderSimple> focusedSimple;
    RequireSucceeded(focusedProvider.query_to(focusedSimple.put()), "focused tree provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*focusedSimple.get(), UIA_ControlTypePropertyId, "focused tree exposes UIA control type") == UIA_TreeItemControlTypeId,
            "root provider focus lookup returns the selected tree item provider for a focused tree");
    Require(ReadProviderStringProperty(*focusedSimple.get(), UIA_NamePropertyId, "focused tree item exposes accessibility name") == L"Panes",
            "root provider focus lookup returns the selected visible tree item");
}

void TestAccessibilityProviderExposesTreeItemSelectionAndExpandCollapsePatterns()
{
    using namespace RedSalamander::DxUi;

    class ExpandableTreeModel final : public IDxTreeModel
    {
    public:
        void SetExpanded(bool expanded)
        {
            _expanded = expanded;
        }

        [[nodiscard]] size_t GetVisibleItemCount() const noexcept override
        {
            return _expanded ? 3u : 2u;
        }

        void GetVisibleItem(size_t visibleIndex, TreeItemData& outItem) const override
        {
            switch (visibleIndex)
            {
                case 0u: outItem = TreeItemData{.id = 10u, .text = L"Plugins", .hasChildren = true, .expanded = _expanded}; return;
                case 1u:
                    if (_expanded)
                    {
                        outItem = TreeItemData{.id = 11u, .parentId = 10u, .text = L"FTP", .depth = 1u};
                    }
                    else
                    {
                        outItem = TreeItemData{.id = 12u, .text = L"Search"};
                    }
                    return;
                case 2u: outItem = TreeItemData{.id = 12u, .text = L"Search"}; return;
                default: throw std::out_of_range("invalid visible tree item");
            }
        }

    private:
        bool _expanded = false;
    };

    class ExpandableTreeDelegate final : public IDxTreeDelegate
    {
    public:
        ExpandableTreeDelegate(ExpandableTreeModel& model, Tree& tree) : _model(model), _tree(tree)
        {
        }

        ExpandableTreeDelegate(const ExpandableTreeDelegate&)            = delete;
        ExpandableTreeDelegate& operator=(const ExpandableTreeDelegate&) = delete;
        ExpandableTreeDelegate(ExpandableTreeDelegate&&)                 = delete;
        ExpandableTreeDelegate& operator=(ExpandableTreeDelegate&&)      = delete;

        void OnTreeSelectionChanged(uint64_t itemId) override
        {
            ++selectionChangedCount;
            lastSelectedItemId = itemId;
        }

        void OnTreeToggleExpanded(uint64_t itemId, bool expanded) override
        {
            ++toggleCount;
            lastToggledItemId = itemId;
            lastExpandedState = expanded;
            _model.SetExpanded(expanded);
            _tree.NotifyDataChanged();
        }

        size_t selectionChangedCount = 0u;
        std::optional<uint64_t> lastSelectedItemId;
        size_t toggleCount = 0u;
        std::optional<uint64_t> lastToggledItemId;
        std::optional<bool> lastExpandedState;

    private:
        ExpandableTreeModel& _model;
        Tree& _tree;
    };

    AttachedHostWindow window;
    auto root       = std::make_unique<Panel>();
    auto* treeLabel = root->AddChild<Label>(L"Categories");
    treeLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 28.0f, 240.0f, 120.0f));

    ExpandableTreeModel treeModel;
    ExpandableTreeDelegate delegate(treeModel, *tree);
    tree->SetModel(&treeModel);
    tree->SetDelegate(&delegate);
    tree->SetSelectedItemId(10u);
    treeLabel->SetMnemonicTarget(tree);

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "tree-item pattern accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> treeLabelProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 48.0f, 12.0f, "tree label accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderFragment> treeProvider;
    RequireSucceeded(treeLabelProvider->Navigate(NavigateDirection_NextSibling, treeProvider.put()), "tree label accessibility provider navigates to the tree");
    Require(treeProvider != nullptr, "tree label accessibility provider returns the tree as the next sibling");

    wil::com_ptr_nothrow<IRawElementProviderFragment> parentItemProvider;
    RequireSucceeded(treeProvider->Navigate(NavigateDirection_FirstChild, parentItemProvider.put()),
                     "tree accessibility provider navigates to the expandable parent item");
    Require(parentItemProvider != nullptr, "tree accessibility provider returns the parent tree item");
    wil::com_ptr_nothrow<IRawElementProviderSimple> parentItemSimple;
    RequireSucceeded(parentItemProvider.query_to(parentItemSimple.put()), "parent tree-item provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> selectionPatternUnknown;
    RequireSucceeded(parentItemSimple->GetPatternProvider(UIA_SelectionItemPatternId, selectionPatternUnknown.put()),
                     "tree item selection pattern lookup succeeds");
    Require(selectionPatternUnknown != nullptr, "tree item accessibility provider exposes selection-item pattern");
    wil::com_ptr_nothrow<ISelectionItemProvider> selectionPattern;
    RequireSucceeded(selectionPatternUnknown.query_to(selectionPattern.put()), "tree item selection pattern supports ISelectionItemProvider");

    BOOL isSelected = FALSE;
    RequireSucceeded(selectionPattern->get_IsSelected(&isSelected), "tree item selected-state query succeeds");
    Require(isSelected == TRUE, "parent tree item pattern reports the initial selection");

    wil::com_ptr_nothrow<IRawElementProviderSimple> selectionContainer;
    RequireSucceeded(selectionPattern->get_SelectionContainer(selectionContainer.put()), "tree item selection container lookup succeeds");
    Require(selectionContainer != nullptr, "tree item selection pattern exposes the tree container");
    Require(ReadProviderStringProperty(*selectionContainer.get(), UIA_NamePropertyId, "tree selection container exposes accessibility name") == L"Categories",
            "tree item selection container resolves to the labeled tree host");
    wil::com_ptr_nothrow<IUnknown> treeSelectionContainerUnknown;
    RequireSucceeded(selectionContainer->GetPatternProvider(UIA_SelectionPatternId, treeSelectionContainerUnknown.put()),
                     "tree selection container selection-pattern lookup succeeds");
    Require(treeSelectionContainerUnknown != nullptr, "tree selection container exposes the selection pattern");
    wil::com_ptr_nothrow<ISelectionProvider> treeSelectionProvider;
    RequireSucceeded(treeSelectionContainerUnknown.query_to(treeSelectionProvider.put()), "tree selection container pattern supports ISelectionProvider");
    BOOL canSelectMultiple = TRUE;
    RequireSucceeded(treeSelectionProvider->get_CanSelectMultiple(&canSelectMultiple), "tree selection provider reports multi-select capability");
    Require(canSelectMultiple == FALSE, "tree selection provider reports single-selection behavior");
    Require(ReadSelectionProviderNames(*treeSelectionProvider.get(), "tree selection provider returns selected item names") ==
                std::vector<std::wstring>{L"Plugins"},
            "tree selection provider resolves the currently selected tree item");

    RequireSucceeded(selectionPattern->RemoveFromSelection(), "tree item selection pattern can remove the current selection");
    Require(! tree->GetSelectedItemId().has_value(), "tree selection removal clears the selected item");
    Require(! ReadProviderBoolProperty(*parentItemSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "tree item selected state updates after removal"),
            "tree item provider reports deselection after RemoveFromSelection");

    RequireSucceeded(selectionPattern->Select(), "tree item selection pattern can restore the selection");
    Require(tree->GetSelectedItemId() && tree->GetSelectedItemId().value() == 10u, "tree item selection pattern selects the parent item");
    Require(delegate.selectionChangedCount == 1u && delegate.lastSelectedItemId && delegate.lastSelectedItemId.value() == 10u,
            "tree selection pattern uses the shared delegate-driven selection path");

    wil::com_ptr_nothrow<IUnknown> expandPatternUnknown;
    RequireSucceeded(parentItemSimple->GetPatternProvider(UIA_ExpandCollapsePatternId, expandPatternUnknown.put()),
                     "tree item expand-collapse pattern lookup succeeds");
    Require(expandPatternUnknown != nullptr, "expandable tree item exposes expand-collapse pattern");
    wil::com_ptr_nothrow<IExpandCollapseProvider> expandPattern;
    RequireSucceeded(expandPatternUnknown.query_to(expandPattern.put()), "expand-collapse pattern supports IExpandCollapseProvider");

    ExpandCollapseState expandState = ExpandCollapseState_LeafNode;
    RequireSucceeded(expandPattern->get_ExpandCollapseState(&expandState), "tree item expand state query succeeds");
    Require(expandState == ExpandCollapseState_Collapsed, "tree item expand-collapse pattern reports the collapsed state");

    Require(! window.Host().DebugHasActiveAnimationSubscription(), "tree item expand-collapse pattern starts without an active host animation");
    RequireSucceeded(expandPattern->Expand(), "tree item expand-collapse pattern can expand the parent item");
    Require(delegate.toggleCount == 1u && delegate.lastToggledItemId && delegate.lastToggledItemId.value() == 10u && delegate.lastExpandedState == true,
            "tree item expand-collapse pattern uses the shared delegate-driven expansion path");
    Require(window.Host().DebugHasActiveAnimationSubscription(), "tree item expand-collapse pattern requests host animation for expansion visuals");
    Require(treeModel.GetVisibleItemCount() == 3u, "tree model exposes the child item after Expand");
    RequireSucceeded(expandPattern->get_ExpandCollapseState(&expandState), "expanded tree item state query succeeds");
    Require(expandState == ExpandCollapseState_Expanded, "tree item expand-collapse pattern reports the expanded state");

    wil::com_ptr_nothrow<IRawElementProviderFragment> childItemProvider;
    RequireSucceeded(parentItemProvider->Navigate(NavigateDirection_NextSibling, childItemProvider.put()),
                     "expanded parent tree item navigates to its first visible child");
    Require(childItemProvider != nullptr, "expanded parent tree item returns the child provider");
    wil::com_ptr_nothrow<IRawElementProviderSimple> childItemSimple;
    RequireSucceeded(childItemProvider.query_to(childItemSimple.put()), "child tree-item provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*childItemSimple.get(), UIA_NamePropertyId, "child tree item exposes accessibility name") == L"FTP",
            "expanded tree item navigation reaches the expected child");

    wil::com_ptr_nothrow<IUnknown> childSelectionUnknown;
    RequireSucceeded(childItemSimple->GetPatternProvider(UIA_SelectionItemPatternId, childSelectionUnknown.put()),
                     "child tree-item selection pattern lookup succeeds");
    Require(childSelectionUnknown != nullptr, "child tree item exposes selection-item pattern");
    wil::com_ptr_nothrow<ISelectionItemProvider> childSelectionPattern;
    RequireSucceeded(childSelectionUnknown.query_to(childSelectionPattern.put()), "child selection pattern supports ISelectionItemProvider");

    RequireSucceeded(childSelectionPattern->AddToSelection(), "child tree-item selection pattern can select the child");
    Require(tree->GetSelectedItemId() && tree->GetSelectedItemId().value() == 11u, "child tree item selection updates the tree selection");
    Require(delegate.selectionChangedCount == 2u && delegate.lastSelectedItemId && delegate.lastSelectedItemId.value() == 11u,
            "child tree item selection continues to use the shared delegate path");
    Require(ReadSelectionProviderNames(*treeSelectionProvider.get(), "tree selection provider updates after child selection") ==
                std::vector<std::wstring>{L"FTP"},
            "tree selection provider tracks the newly selected visible tree item");

    wil::com_ptr_nothrow<IUnknown> childExpandUnknown;
    RequireSucceeded(childItemSimple->GetPatternProvider(UIA_ExpandCollapsePatternId, childExpandUnknown.put()),
                     "leaf tree-item expand-collapse lookup succeeds");
    Require(childExpandUnknown == nullptr, "leaf tree item does not expose expand-collapse pattern");

    RequireSucceeded(expandPattern->Collapse(), "tree item expand-collapse pattern can collapse the parent item");
    Require(delegate.toggleCount == 2u && delegate.lastExpandedState == false, "tree item collapse again uses the shared delegate path");
    Require(treeModel.GetVisibleItemCount() == 2u, "tree model hides the child item after Collapse");
    Require(! tree->GetSelectedItemId().has_value(), "tree collapse clears a selection that is no longer visible");
    RequireSucceeded(expandPattern->get_ExpandCollapseState(&expandState), "collapsed tree item state query succeeds after Collapse");
    Require(expandState == ExpandCollapseState_Collapsed, "tree item expand-collapse pattern reports the collapsed state after Collapse");
}

void TestAccessibilityProviderExposesGridRowSelectionPatterns()
{
    using namespace RedSalamander::DxUi;

    class AccessibleGridModel final : public IDxGridModel
    {
    public:
        struct Row
        {
            uint64_t stableId = 0u;
            std::wstring name;
            std::wstring status;
        };

        explicit AccessibleGridModel(std::vector<Row> rows) : _rows(std::move(rows))
        {
        }

        [[nodiscard]] size_t GetRowCount() const noexcept override
        {
            return _rows.size();
        }

        [[nodiscard]] size_t GetColumnCount() const noexcept override
        {
            return 2u;
        }

        [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
        {
            GridColumnDesc column;
            if (columnIndex == 0u)
            {
                column.id       = L"name";
                column.title    = L"Name";
                column.widthDip = 140.0f;
            }
            else
            {
                column.id       = L"status";
                column.title    = L"Status";
                column.widthDip = 100.0f;
            }
            return column;
        }

        void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
        {
            const Row& row = _rows.at(rowIndex);
            outCell.kind   = GridCellKind::Text;
            outCell.text   = (columnIndex == 0u) ? row.name : row.status;
        }

        [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
        {
            return _rows[rowIndex].stableId;
        }

        [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
        {
            for (size_t rowIndex = 0u; rowIndex < _rows.size(); ++rowIndex)
            {
                if (_rows[rowIndex].stableId == rowId)
                {
                    return rowIndex;
                }
            }

            return std::nullopt;
        }

    private:
        std::vector<Row> _rows;
    };

    class AccessibleGridDelegate final : public IDxGridDelegate
    {
    public:
        using IDxGridDelegate::OnGridSelectionChanged;

        void OnGridSelectionChanged(Grid& sender) override
        {
            ++selectionChangedCount;
            selectionCounts.push_back(sender.GetSelectionModel().GetCount());
            orderedSelection.assign(sender.GetSelectionModel().GetOrderedSelection().begin(), sender.GetSelectionModel().GetOrderedSelection().end());
        }

        size_t selectionChangedCount = 0u;
        std::vector<size_t> selectionCounts;
        std::vector<uint64_t> orderedSelection;
    };

    AttachedHostWindow window;
    auto root       = std::make_unique<Panel>();
    auto* gridLabel = root->AddChild<Label>(L"Results");
    gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 140.0f));

    AccessibleGridModel gridModel({AccessibleGridModel::Row{100u, L"Alpha", L"Ready"},
                                   AccessibleGridModel::Row{200u, L"Beta", L"Busy"},
                                   AccessibleGridModel::Row{300u, L"Gamma", L"Idle"}});
    AccessibleGridDelegate delegate;
    grid->SetModel(&gridModel);
    grid->SetDelegate(&delegate);
    gridLabel->SetMnemonicTarget(grid);

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "grid-row accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> gridLabelProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 12.0f, "grid label accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderFragment> gridProvider;
    RequireSucceeded(gridLabelProvider->Navigate(NavigateDirection_NextSibling, gridProvider.put()), "grid label accessibility provider navigates to the grid");
    Require(gridProvider != nullptr, "grid label accessibility provider returns the grid as the next sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> gridSimple;
    RequireSucceeded(gridProvider.query_to(gridSimple.put()), "grid accessibility provider exposes IRawElementProviderSimple");
    wil::com_ptr_nothrow<IUnknown> tablePatternUnknown;
    RequireSucceeded(gridSimple->GetPatternProvider(UIA_TablePatternId, tablePatternUnknown.put()), "grid table-pattern lookup succeeds");
    Require(tablePatternUnknown != nullptr, "grid accessibility provider exposes the table pattern");
    wil::com_ptr_nothrow<ITableProvider> tablePattern;
    RequireSucceeded(tablePatternUnknown.query_to(tablePattern.put()), "grid table pattern supports ITableProvider");

    SAFEARRAY* columnHeadersArray = nullptr;
    RequireSucceeded(tablePattern->GetColumnHeaders(&columnHeadersArray), "grid table pattern returns visible column headers");
    Require(columnHeadersArray != nullptr, "grid table pattern returns a column-header array");
    const auto destroyColumnHeadersArray = wil::scope_exit([&] { SafeArrayDestroy(columnHeadersArray); });
    Require(ReadProviderArrayNames(columnHeadersArray, "grid table column headers expose header names") == std::vector<std::wstring>({L"Name", L"Status"}),
            "grid table pattern exposes visible grid header fragments in display order");

    SAFEARRAY* rowHeadersArray = nullptr;
    RequireSucceeded(tablePattern->GetRowHeaders(&rowHeadersArray), "grid table pattern row-header lookup succeeds");
    Require(rowHeadersArray != nullptr, "grid table pattern returns a row-header array");
    const auto destroyRowHeadersArray = wil::scope_exit([&] { SafeArrayDestroy(rowHeadersArray); });
    Require(ReadProviderArrayNames(rowHeadersArray, "grid table row headers return an empty array").empty(),
            "grid table pattern reports no row-header fragments for row-headerless grids");

    RowOrColumnMajor rowOrColumnMajor = RowOrColumnMajor_RowMajor;
    RequireSucceeded(tablePattern->get_RowOrColumnMajor(&rowOrColumnMajor), "grid table pattern row-or-column-major query succeeds");
    Require(rowOrColumnMajor == RowOrColumnMajor_Indeterminate, "grid table pattern reports indeterminate row/column major order");

    const std::optional<D2D1_RECT_F> firstHeaderRect = grid->GetVisibleColumnHeaderRect(0u);
    Require(firstHeaderRect.has_value(), "grid exposes a visible header rect for point hit-testing");
    const float firstHeaderCenterXDip                                = (firstHeaderRect->left + firstHeaderRect->right) * 0.5f;
    const float firstHeaderCenterYDip                                = (firstHeaderRect->top + firstHeaderRect->bottom) * 0.5f;
    wil::com_ptr_nothrow<IRawElementProviderFragment> headerProvider = GetProviderAtDipPoint(window.Hwnd(),
                                                                                             window.Host(),
                                                                                             *rootProvider.get(),
                                                                                             firstHeaderCenterXDip,
                                                                                             firstHeaderCenterYDip,
                                                                                             "grid header accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> headerSimple;
    RequireSucceeded(headerProvider.query_to(headerSimple.put()), "grid header accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*headerSimple.get(), UIA_ControlTypePropertyId, "grid header exposes UIA control type") == UIA_HeaderItemControlTypeId,
            "grid hit-testing resolves the visible column-header fragment");
    Require(ReadProviderStringProperty(*headerSimple.get(), UIA_NamePropertyId, "grid header exposes accessibility name") == L"Name",
            "grid header accessibility provider exposes the visible column header title");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstHeaderFragmentProvider;
    RequireSucceeded(gridProvider->Navigate(NavigateDirection_FirstChild, firstHeaderFragmentProvider.put()),
                     "grid accessibility provider navigates to the first visible header");
    Require(firstHeaderFragmentProvider != nullptr, "grid accessibility provider returns a first header child");
    wil::com_ptr_nothrow<IRawElementProviderSimple> firstHeaderFragmentSimple;
    RequireSucceeded(firstHeaderFragmentProvider.query_to(firstHeaderFragmentSimple.put()), "grid header fragment exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*firstHeaderFragmentSimple.get(), UIA_ControlTypePropertyId, "grid header fragment exposes UIA control type") ==
                UIA_HeaderItemControlTypeId,
            "grid accessibility provider reports a header-item first child when visible headers are present");
    Require(ReadProviderStringProperty(*firstHeaderFragmentSimple.get(), UIA_NamePropertyId, "grid header fragment exposes accessibility name") == L"Name",
            "grid accessibility provider returns the first visible column header before row fragments");

    wil::com_ptr_nothrow<IRawElementProviderFragment> secondHeaderProvider;
    RequireSucceeded(firstHeaderFragmentProvider->Navigate(NavigateDirection_NextSibling, secondHeaderProvider.put()),
                     "grid header fragment navigates to the next visible header");
    Require(secondHeaderProvider != nullptr, "grid header fragment returns the next visible header sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> secondHeaderSimple;
    RequireSucceeded(secondHeaderProvider.query_to(secondHeaderSimple.put()), "second grid header fragment exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*secondHeaderSimple.get(), UIA_NamePropertyId, "second grid header exposes accessibility name") == L"Status",
            "grid accessibility provider exposes the remaining visible column headers before row fragments");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstRowProvider;
    RequireSucceeded(secondHeaderProvider->Navigate(NavigateDirection_NextSibling, firstRowProvider.put()),
                     "grid header fragment navigates to the first visible row after the last header");
    Require(firstRowProvider != nullptr, "grid accessibility provider returns a first row child after visible headers");
    wil::com_ptr_nothrow<IRawElementProviderSimple> firstRowSimple;
    RequireSucceeded(firstRowProvider.query_to(firstRowSimple.put()), "grid row accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*firstRowSimple.get(), UIA_ControlTypePropertyId, "grid row exposes UIA control type") == UIA_DataItemControlTypeId,
            "grid row accessibility provider reports data-item control type");
    Require(ReadProviderStringProperty(*firstRowSimple.get(), UIA_NamePropertyId, "grid row exposes accessibility name") == L"Alpha | Ready",
            "grid row accessibility provider exposes joined visible cell text");
    Require(! ReadProviderBoolProperty(*firstRowSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "grid row exposes selected state"),
            "grid row accessibility provider reports the unselected initial row");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstCellProvider;
    RequireSucceeded(firstRowProvider->Navigate(NavigateDirection_FirstChild, firstCellProvider.put()),
                     "grid row accessibility provider navigates to the first visible cell");
    Require(firstCellProvider != nullptr, "grid row accessibility provider returns a first cell child");
    wil::com_ptr_nothrow<IRawElementProviderSimple> firstCellSimple;
    RequireSucceeded(firstCellProvider.query_to(firstCellSimple.put()), "grid cell accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*firstCellSimple.get(), UIA_ControlTypePropertyId, "grid cell exposes UIA control type") == UIA_TextControlTypeId,
            "grid cell accessibility provider reports text control type for text cells");
    Require(ReadProviderStringProperty(*firstCellSimple.get(), UIA_NamePropertyId, "grid cell exposes accessibility name") == L"Alpha",
            "grid cell accessibility provider exposes the visible cell text");
    Require(ReadProviderStringProperty(*firstCellSimple.get(), UIA_ValueValuePropertyId, "grid cell exposes value") == L"Alpha",
            "grid text cell accessibility provider exposes a read-only value");
    Require(ReadProviderBoolProperty(*firstCellSimple.get(), UIA_ValueIsReadOnlyPropertyId, "grid cell exposes read-only state"),
            "grid text cell accessibility provider reports the value pattern as read-only");

    wil::com_ptr_nothrow<IUnknown> firstCellValueUnknown;
    RequireSucceeded(firstCellSimple->GetPatternProvider(UIA_ValuePatternId, firstCellValueUnknown.put()), "grid cell value-pattern lookup succeeds");
    Require(firstCellValueUnknown != nullptr, "grid text cell accessibility provider exposes the value pattern");
    wil::com_ptr_nothrow<IValueProvider> firstCellValuePattern;
    RequireSucceeded(firstCellValueUnknown.query_to(firstCellValuePattern.put()), "grid cell value pattern supports IValueProvider");
    BSTR firstCellValue = nullptr;
    RequireSucceeded(firstCellValuePattern->get_Value(&firstCellValue), "grid cell value pattern returns the visible cell text");
    const auto freeFirstCellValue = wil::scope_exit([&] { SysFreeString(firstCellValue); });
    Require(std::wstring(firstCellValue ? firstCellValue : L"") == L"Alpha", "grid cell value pattern returns the expected visible text value");
    BOOL firstCellReadOnly = FALSE;
    RequireSucceeded(firstCellValuePattern->get_IsReadOnly(&firstCellReadOnly), "grid cell value pattern read-only lookup succeeds");
    Require(firstCellReadOnly == TRUE, "grid cell value pattern reports a read-only value");

    wil::com_ptr_nothrow<IUnknown> firstCellGridItemUnknown;
    RequireSucceeded(firstCellSimple->GetPatternProvider(UIA_GridItemPatternId, firstCellGridItemUnknown.put()), "grid cell grid-item pattern lookup succeeds");
    Require(firstCellGridItemUnknown != nullptr, "grid cell accessibility provider exposes the grid-item pattern");
    wil::com_ptr_nothrow<IGridItemProvider> firstCellGridItemPattern;
    RequireSucceeded(firstCellGridItemUnknown.query_to(firstCellGridItemPattern.put()), "grid cell grid-item pattern supports IGridItemProvider");
    wil::com_ptr_nothrow<IUnknown> firstCellTableItemUnknown;
    RequireSucceeded(firstCellSimple->GetPatternProvider(UIA_TableItemPatternId, firstCellTableItemUnknown.put()),
                     "grid cell table-item pattern lookup succeeds");
    Require(firstCellTableItemUnknown != nullptr, "grid cell accessibility provider exposes the table-item pattern");
    wil::com_ptr_nothrow<ITableItemProvider> firstCellTableItemPattern;
    RequireSucceeded(firstCellTableItemUnknown.query_to(firstCellTableItemPattern.put()), "grid cell table-item pattern supports ITableItemProvider");

    int firstCellRow        = -1;
    int firstCellColumn     = -1;
    int firstCellRowSpan    = 0;
    int firstCellColumnSpan = 0;
    RequireSucceeded(firstCellGridItemPattern->get_Row(&firstCellRow), "grid cell grid-item row query succeeds");
    RequireSucceeded(firstCellGridItemPattern->get_Column(&firstCellColumn), "grid cell grid-item column query succeeds");
    RequireSucceeded(firstCellGridItemPattern->get_RowSpan(&firstCellRowSpan), "grid cell grid-item row-span query succeeds");
    RequireSucceeded(firstCellGridItemPattern->get_ColumnSpan(&firstCellColumnSpan), "grid cell grid-item column-span query succeeds");
    Require(firstCellRow == 0 && firstCellColumn == 0, "grid cell grid-item metadata reports the expected row and column");
    Require(firstCellRowSpan == 1 && firstCellColumnSpan == 1, "grid cell grid-item metadata reports single-cell spans");

    wil::com_ptr_nothrow<IRawElementProviderSimple> firstCellContainingGrid;
    RequireSucceeded(firstCellGridItemPattern->get_ContainingGrid(firstCellContainingGrid.put()), "grid cell containing-grid lookup succeeds");
    Require(firstCellContainingGrid != nullptr, "grid cell grid-item pattern resolves the containing grid");
    Require(ReadProviderStringProperty(*firstCellContainingGrid.get(), UIA_NamePropertyId, "grid cell containing grid exposes accessibility name") ==
                L"Results",
            "grid cell grid-item pattern resolves the labeled grid container");

    SAFEARRAY* firstCellColumnHeadersArray = nullptr;
    RequireSucceeded(firstCellTableItemPattern->GetColumnHeaderItems(&firstCellColumnHeadersArray),
                     "grid cell table-item pattern returns the owning column header");
    Require(firstCellColumnHeadersArray != nullptr, "grid cell table-item pattern returns a column-header array");
    const auto destroyFirstCellColumnHeadersArray = wil::scope_exit([&] { SafeArrayDestroy(firstCellColumnHeadersArray); });
    Require(ReadProviderArrayNames(firstCellColumnHeadersArray, "grid cell table-item column headers expose the owning header") ==
                std::vector<std::wstring>{L"Name"},
            "grid cell table-item pattern resolves the visible owning column header");

    SAFEARRAY* firstCellRowHeadersArray = nullptr;
    RequireSucceeded(firstCellTableItemPattern->GetRowHeaderItems(&firstCellRowHeadersArray), "grid cell table-item row-header lookup succeeds");
    Require(firstCellRowHeadersArray != nullptr, "grid cell table-item pattern returns a row-header array");
    const auto destroyFirstCellRowHeadersArray = wil::scope_exit([&] { SafeArrayDestroy(firstCellRowHeadersArray); });
    Require(ReadProviderArrayNames(firstCellRowHeadersArray, "grid cell table-item row headers return an empty array").empty(),
            "grid cell table-item pattern reports no row-header fragments for row-headerless grids");

    wil::com_ptr_nothrow<IRawElementProviderFragment> secondCellProvider;
    RequireSucceeded(firstCellProvider->Navigate(NavigateDirection_NextSibling, secondCellProvider.put()),
                     "grid cell accessibility provider navigates to the next visible cell");
    Require(secondCellProvider != nullptr, "grid cell accessibility provider returns the next visible cell");
    wil::com_ptr_nothrow<IRawElementProviderSimple> secondCellSimple;
    RequireSucceeded(secondCellProvider.query_to(secondCellSimple.put()), "second grid cell accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*secondCellSimple.get(), UIA_NamePropertyId, "second grid cell exposes accessibility name") == L"Ready",
            "grid cell navigation reaches the expected second visible cell");

    wil::com_ptr_nothrow<IRawElementProviderFragment> previousCellProvider;
    RequireSucceeded(secondCellProvider->Navigate(NavigateDirection_PreviousSibling, previousCellProvider.put()),
                     "second grid cell accessibility provider navigates back to the previous visible cell");
    Require(previousCellProvider != nullptr, "grid cell accessibility provider returns the previous visible cell");
    wil::com_ptr_nothrow<IRawElementProviderSimple> previousCellSimple;
    RequireSucceeded(previousCellProvider.query_to(previousCellSimple.put()), "previous grid cell accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*previousCellSimple.get(), UIA_NamePropertyId, "previous grid cell exposes accessibility name") == L"Alpha",
            "grid cell previous-sibling navigation returns to the first cell");

    wil::com_ptr_nothrow<IRawElementProviderFragment> parentRowFromCell;
    RequireSucceeded(firstCellProvider->Navigate(NavigateDirection_Parent, parentRowFromCell.put()),
                     "grid cell accessibility provider navigates back to its row");
    Require(parentRowFromCell != nullptr, "grid cell accessibility provider returns its parent row");
    wil::com_ptr_nothrow<IRawElementProviderSimple> parentRowSimple;
    RequireSucceeded(parentRowFromCell.query_to(parentRowSimple.put()), "parent row provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*parentRowSimple.get(), UIA_NamePropertyId, "parent row provider exposes accessibility name") == L"Alpha | Ready",
            "grid cell parent navigation returns the owning row provider");

    wil::com_ptr_nothrow<IUnknown> firstSelectionUnknown;
    RequireSucceeded(firstRowSimple->GetPatternProvider(UIA_SelectionItemPatternId, firstSelectionUnknown.put()), "grid row selection pattern lookup succeeds");
    Require(firstSelectionUnknown != nullptr, "grid row accessibility provider exposes selection-item pattern");
    wil::com_ptr_nothrow<ISelectionItemProvider> firstSelectionPattern;
    RequireSucceeded(firstSelectionUnknown.query_to(firstSelectionPattern.put()), "grid row selection pattern supports ISelectionItemProvider");

    wil::com_ptr_nothrow<IRawElementProviderSimple> firstSelectionContainer;
    RequireSucceeded(firstSelectionPattern->get_SelectionContainer(firstSelectionContainer.put()), "grid row selection container lookup succeeds");
    Require(firstSelectionContainer != nullptr, "grid row selection pattern exposes the grid container");
    Require(ReadProviderStringProperty(*firstSelectionContainer.get(), UIA_NamePropertyId, "grid selection container exposes accessibility name") == L"Results",
            "grid row selection container resolves to the labeled grid host");
    wil::com_ptr_nothrow<IUnknown> gridSelectionContainerUnknown;
    RequireSucceeded(firstSelectionContainer->GetPatternProvider(UIA_SelectionPatternId, gridSelectionContainerUnknown.put()),
                     "grid selection container selection-pattern lookup succeeds");
    Require(gridSelectionContainerUnknown != nullptr, "grid selection container exposes the selection pattern");
    wil::com_ptr_nothrow<ISelectionProvider> gridSelectionProvider;
    RequireSucceeded(gridSelectionContainerUnknown.query_to(gridSelectionProvider.put()), "grid selection container pattern supports ISelectionProvider");
    BOOL canSelectMultiple = FALSE;
    RequireSucceeded(gridSelectionProvider->get_CanSelectMultiple(&canSelectMultiple), "grid selection provider reports multi-select capability");
    Require(canSelectMultiple == TRUE, "grid selection provider reports extended multi-selection behavior");

    RequireSucceeded(firstSelectionPattern->Select(), "grid row selection pattern can select the first row");
    Require(grid->IsRowSelected(0u), "grid row selection pattern selects the first row");
    Require(delegate.selectionChangedCount == 1u && delegate.selectionCounts == std::vector<size_t>{1u} &&
                delegate.orderedSelection == std::vector<uint64_t>{100u},
            "grid row selection pattern uses the shared delegate-driven selection path");
    Require(ReadProviderBoolProperty(*firstRowSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "grid row selected state updates after Select"),
            "grid row accessibility provider reports the selected row after Select");
    Require(ReadSelectionProviderNames(*gridSelectionProvider.get(), "grid selection provider returns selected row names") ==
                std::vector<std::wstring>{L"Alpha | Ready"},
            "grid selection provider resolves the currently selected visible row");

    const std::optional<D2D1_RECT_F> firstCellRect = grid->GetVisibleCellRect(0u, 0u);
    Require(firstCellRect.has_value(), "grid exposes a visible cell rect for point hit-testing");
    const float firstCellCenterXDip                                   = (firstCellRect->left + firstCellRect->right) * 0.5f;
    const float firstCellCenterYDip                                   = (firstCellRect->top + firstCellRect->bottom) * 0.5f;
    wil::com_ptr_nothrow<IRawElementProviderFragment> hitCellProvider = GetProviderAtDipPoint(
        window.Hwnd(), window.Host(), *rootProvider.get(), firstCellCenterXDip, firstCellCenterYDip, "grid cell accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> hitCellSimple;
    RequireSucceeded(hitCellProvider.query_to(hitCellSimple.put()), "grid point-hit provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*hitCellSimple.get(), UIA_ControlTypePropertyId, "grid point-hit provider exposes cell control type") ==
                UIA_TextControlTypeId,
            "grid hit-testing resolves the visible cell provider instead of only the row or grid container");
    Require(ReadProviderStringProperty(*hitCellSimple.get(), UIA_NamePropertyId, "grid point-hit provider exposes cell name") == L"Alpha",
            "grid hit-testing resolves the expected visible cell provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> secondRowProvider;
    RequireSucceeded(firstRowProvider->Navigate(NavigateDirection_NextSibling, secondRowProvider.put()),
                     "first grid row provider navigates to the next visible row");
    Require(secondRowProvider != nullptr, "first grid row provider returns the next sibling row");
    wil::com_ptr_nothrow<IRawElementProviderSimple> secondRowSimple;
    RequireSucceeded(secondRowProvider.query_to(secondRowSimple.put()), "second grid row provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*secondRowSimple.get(), UIA_NamePropertyId, "second grid row exposes accessibility name") == L"Beta | Busy",
            "grid row navigation reaches the expected second row");

    wil::com_ptr_nothrow<IUnknown> secondSelectionUnknown;
    RequireSucceeded(secondRowSimple->GetPatternProvider(UIA_SelectionItemPatternId, secondSelectionUnknown.put()),
                     "second grid row selection pattern lookup succeeds");
    Require(secondSelectionUnknown != nullptr, "second grid row exposes selection-item pattern");
    wil::com_ptr_nothrow<ISelectionItemProvider> secondSelectionPattern;
    RequireSucceeded(secondSelectionUnknown.query_to(secondSelectionPattern.put()), "second grid row selection pattern supports ISelectionItemProvider");

    RequireSucceeded(secondSelectionPattern->AddToSelection(), "grid row selection pattern can extend the selection");
    Require(grid->IsRowSelected(0u) && grid->IsRowSelected(1u), "grid row AddToSelection preserves the first row and adds the second");
    Require(delegate.selectionChangedCount == 2u && delegate.selectionCounts == std::vector<size_t>({1u, 2u}) &&
                delegate.orderedSelection == std::vector<uint64_t>({100u, 200u}),
            "grid row AddToSelection continues to use the shared delegate path");
    Require(ReadProviderBoolProperty(*secondRowSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "second grid row exposes selected state"),
            "second grid row provider reports selection after AddToSelection");
    Require(ReadSelectionProviderNames(*gridSelectionProvider.get(), "grid selection provider updates after AddToSelection") ==
                std::vector<std::wstring>({L"Alpha | Ready", L"Beta | Busy"}),
            "grid selection provider tracks the ordered visible row selection");

    window.Host().SetFocusControl(grid);
    wil::com_ptr_nothrow<IRawElementProviderFragment> focusedProvider;
    RequireSucceeded(rootProvider->GetFocus(focusedProvider.put()), "root provider focus lookup succeeds for the grid");
    Require(focusedProvider != nullptr, "root provider returns the focused grid row provider");
    wil::com_ptr_nothrow<IRawElementProviderSimple> focusedSimple;
    RequireSucceeded(focusedProvider.query_to(focusedSimple.put()), "focused grid row provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*focusedSimple.get(), UIA_ControlTypePropertyId, "focused grid row exposes UIA control type") == UIA_DataItemControlTypeId,
            "root provider focus lookup returns the selected grid row provider for a focused grid");
    Require(ReadProviderStringProperty(*focusedSimple.get(), UIA_NamePropertyId, "focused grid row exposes accessibility name") == L"Beta | Busy",
            "root provider focus lookup returns the most recently selected visible grid row");

    RequireSucceeded(secondSelectionPattern->RemoveFromSelection(), "grid row selection pattern can remove the second row from the selection");
    Require(grid->IsRowSelected(0u) && ! grid->IsRowSelected(1u), "grid row RemoveFromSelection preserves the remaining visible selection");
    Require(delegate.selectionChangedCount == 3u && delegate.selectionCounts == std::vector<size_t>({1u, 2u, 1u}) &&
                delegate.orderedSelection == std::vector<uint64_t>({100u}),
            "grid row RemoveFromSelection continues to use the shared delegate path");
    Require(! ReadProviderBoolProperty(*secondRowSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "grid row selected state updates after removal"),
            "grid row accessibility provider reports deselection after RemoveFromSelection");
    Require(ReadSelectionProviderNames(*gridSelectionProvider.get(), "grid selection provider updates after removal") ==
                std::vector<std::wstring>{L"Alpha | Ready"},
            "grid selection provider drops the removed row and preserves the remaining selection");
}

void TestAccessibilityProviderExposesHorizontallyScrolledGridRowStructure()
{
    using namespace RedSalamander::DxUi;

    class WideGridModel final : public IDxGridModel
    {
    public:
        [[nodiscard]] size_t GetRowCount() const noexcept override
        {
            return 1u;
        }

        [[nodiscard]] size_t GetColumnCount() const noexcept override
        {
            return 3u;
        }

        [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
        {
            GridColumnDesc column;
            switch (columnIndex)
            {
            case 0u:
                column.id    = L"name";
                column.title = L"Name";
                break;
            case 1u:
                column.id    = L"status";
                column.title = L"Status";
                break;
            default:
                column.id    = L"state";
                column.title = L"State";
                break;
            }
            column.widthDip = 120.0f;
            return column;
        }

        void GetCellData(size_t /*rowIndex*/, size_t columnIndex, GridCellData& outCell) const override
        {
            outCell.kind = GridCellKind::Text;
            switch (columnIndex)
            {
            case 0u:
                outCell.text = L"Alpha";
                break;
            case 1u:
                outCell.text = L"Ready";
                break;
            default:
                outCell.text = L"Archived";
                break;
            }
        }

        [[nodiscard]] uint64_t GetStableRowId(size_t /*rowIndex*/) const noexcept override
        {
            return 100u;
        }

        [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
        {
            return (rowId == 100u) ? std::optional<size_t>{0u} : std::nullopt;
        }
    };

    AttachedHostWindow window;
    auto root       = std::make_unique<Panel>();
    auto* gridLabel = root->AddChild<Label>(L"Results");
    gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 190.0f, 128.0f));

    WideGridModel gridModel;
    grid->SetModel(&gridModel);
    grid->DebugSetScrollOffsets(0.0f, 130.0f);
    gridLabel->SetMnemonicTarget(grid);

    Require(! grid->GetVisibleCellRect(0u, 0u).has_value(), "scrolled grid keeps the first column outside the visible cell viewport");
    Require(grid->GetVisibleCellRect(0u, 1u).has_value(), "scrolled grid keeps later columns visible for point-hit comparison");

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "horizontally scrolled grid accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> gridLabelProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 12.0f, "scrolled grid label provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderFragment> gridProvider;
    RequireSucceeded(gridLabelProvider->Navigate(NavigateDirection_NextSibling, gridProvider.put()),
                     "scrolled grid label accessibility provider navigates to the grid");
    Require(gridProvider != nullptr, "scrolled grid label provider returns the grid as the next sibling");

    wil::com_ptr_nothrow<IRawElementProviderFragment> childProvider;
    RequireSucceeded(gridProvider->Navigate(NavigateDirection_FirstChild, childProvider.put()),
                     "scrolled grid provider navigates to its first structural child");
    Require(childProvider != nullptr, "scrolled grid provider exposes at least one structural child");

    wil::com_ptr_nothrow<IRawElementProviderFragment> rowProvider;
    for (size_t step = 0u; childProvider && step < 8u; ++step)
    {
        wil::com_ptr_nothrow<IRawElementProviderSimple> childSimple;
        RequireSucceeded(childProvider.query_to(childSimple.put()), "scrolled grid child exposes IRawElementProviderSimple");
        if (ReadProviderLongProperty(*childSimple.get(), UIA_ControlTypePropertyId, "scrolled grid child exposes UIA control type") ==
            UIA_DataItemControlTypeId)
        {
            rowProvider = childProvider;
            break;
        }

        wil::com_ptr_nothrow<IRawElementProviderFragment> nextProvider;
        RequireSucceeded(childProvider->Navigate(NavigateDirection_NextSibling, nextProvider.put()),
                         "scrolled grid child navigates to the next structural sibling");
        childProvider = nextProvider;
    }

    Require(rowProvider != nullptr, "scrolled grid structure exposes a row provider after visible headers");
    wil::com_ptr_nothrow<IRawElementProviderSimple> rowSimple;
    RequireSucceeded(rowProvider.query_to(rowSimple.put()), "scrolled grid row provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*rowSimple.get(), UIA_NamePropertyId, "scrolled grid row exposes accessibility name") ==
                L"Alpha | Ready | Archived",
            "scrolled grid row name includes all model columns, including horizontally off-view cells");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstCellProvider;
    RequireSucceeded(rowProvider->Navigate(NavigateDirection_FirstChild, firstCellProvider.put()),
                     "scrolled grid row provider navigates to its first structural cell");
    Require(firstCellProvider != nullptr, "scrolled grid row exposes a first structural cell");
    wil::com_ptr_nothrow<IRawElementProviderSimple> firstCellSimple;
    RequireSucceeded(firstCellProvider.query_to(firstCellSimple.put()), "scrolled grid first cell provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*firstCellSimple.get(), UIA_NamePropertyId, "scrolled grid first cell exposes accessibility name") == L"Alpha",
            "scrolled grid first row child is the horizontally off-view first model column");
    Require(ReadProviderBoolProperty(*firstCellSimple.get(), UIA_IsOffscreenPropertyId, "scrolled grid first cell exposes offscreen state"),
            "scrolled grid off-view cell fragment reports itself offscreen");

    wil::com_ptr_nothrow<IRawElementProviderFragment> secondCellProvider;
    RequireSucceeded(firstCellProvider->Navigate(NavigateDirection_NextSibling, secondCellProvider.put()),
                     "scrolled grid first cell navigates to the second structural cell");
    Require(secondCellProvider != nullptr, "scrolled grid first cell returns the next model-column cell");
    wil::com_ptr_nothrow<IRawElementProviderSimple> secondCellSimple;
    RequireSucceeded(secondCellProvider.query_to(secondCellSimple.put()), "scrolled grid second cell provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*secondCellSimple.get(), UIA_NamePropertyId, "scrolled grid second cell exposes accessibility name") == L"Ready",
            "scrolled grid cell navigation continues through the full model column set");
}

void TestAccessibilityProviderPointHitsClipAndTranslateScrollPanelChildren()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* scroll = root->AddChild<ScrollPanel>();
    scroll->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));
    scroll->SetContentHeight(260.0f);
    scroll->SetScrollOffset(80.0f);

    auto* offscreenButton = scroll->AddChild<Button>(L"Hidden above");
    offscreenButton->SetBounds(D2D1::RectF(12.0f, 12.0f, 180.0f, 48.0f));
    auto* visibleButton = scroll->AddChild<Button>(L"Visible after scroll");
    visibleButton->SetBounds(D2D1::RectF(12.0f, 112.0f, 180.0f, 148.0f));

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "scrolled panel point-hit test creates a root provider");

    const POINT rawContentSpacePointPx = window.Host().DipPointToScreenPoint(D2D1::Point2F(24.0f, 24.0f));
    wil::com_ptr_nothrow<IRawElementProviderFragment> rawContentSpaceProvider;
    RequireSucceeded(rootProvider->ElementProviderFromPoint(static_cast<double>(rawContentSpacePointPx.x),
                                                            static_cast<double>(rawContentSpacePointPx.y),
                                                            rawContentSpaceProvider.put()),
                     "scrolled panel raw content-space point query succeeds");
    Require(rawContentSpaceProvider != nullptr, "scrolled panel empty viewport point resolves the root provider");
    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rawContentSpaceRoot;
    RequireSucceeded(rawContentSpaceProvider.query_to(rawContentSpaceRoot.put()), "scrolled panel empty viewport provider is the root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> visibleProvider = GetProviderAtDipPoint(window.Hwnd(),
                                                                                              window.Host(),
                                                                                              *rootProvider.get(),
                                                                                              24.0f,
                                                                                              44.0f,
                                                                                              "scrolled panel visible child point is queryable");
    Require(visibleProvider != nullptr, "scrolled ScrollPanel child is hit-testable at its viewport-translated position");
    wil::com_ptr_nothrow<IRawElementProviderSimple> visibleSimple;
    RequireSucceeded(visibleProvider.query_to(visibleSimple.put()), "scrolled panel visible provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*visibleSimple.get(), UIA_NamePropertyId, "scrolled panel visible provider exposes accessibility name") ==
                L"Visible after scroll",
            "scrolled panel point-hit translation resolves the visible child rather than its raw content-space position");
}

void TestAccessibilityProviderExposesGridCellToggleAndRangePatterns()
{
    using namespace RedSalamander::DxUi;

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"Rules");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        CheckboxGridModel gridModel(0u);
        gridModel.SetRows({CheckboxGridModel::Row{L"Rule A", true, true}});
        RecordingCheckboxGridDelegate delegate(gridModel);
        grid->SetModel(&gridModel);
        grid->SetDelegate(&delegate);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "grid checkbox accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> checkboxCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(checkboxCellRect.has_value(), "grid exposes a visible checkbox cell rect for accessibility hit-testing");
        const float checkboxCellCenterXDip = (checkboxCellRect->left + checkboxCellRect->right) * 0.5f;
        const float checkboxCellCenterYDip = (checkboxCellRect->top + checkboxCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> checkboxCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  checkboxCellCenterXDip,
                                  checkboxCellCenterYDip,
                                  "grid checkbox cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> checkboxCellSimple;
        RequireSucceeded(checkboxCellProvider.query_to(checkboxCellSimple.put()),
                         "grid checkbox cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderLongProperty(*checkboxCellSimple.get(), UIA_ControlTypePropertyId, "grid checkbox cell exposes UIA control type") ==
                    UIA_CheckBoxControlTypeId,
                "grid checkbox cell accessibility provider reports checkbox control type");
        Require(ReadProviderStringProperty(*checkboxCellSimple.get(), UIA_NamePropertyId, "grid checkbox cell exposes accessibility name") == L"[x] Enabled",
                "grid checkbox cell accessibility provider exposes the checked cell text");
        Require(ReadProviderLongProperty(*checkboxCellSimple.get(), UIA_ToggleToggleStatePropertyId, "grid checkbox cell exposes toggle state") ==
                    ToggleState_On,
                "grid checkbox cell accessibility provider reports the checked toggle state");

        wil::com_ptr_nothrow<IUnknown> checkboxToggleUnknown;
        RequireSucceeded(checkboxCellSimple->GetPatternProvider(UIA_TogglePatternId, checkboxToggleUnknown.put()),
                         "grid checkbox cell toggle-pattern lookup succeeds");
        Require(checkboxToggleUnknown != nullptr, "grid checkbox cell accessibility provider exposes the toggle pattern");
        wil::com_ptr_nothrow<IToggleProvider> checkboxTogglePattern;
        RequireSucceeded(checkboxToggleUnknown.query_to(checkboxTogglePattern.put()), "grid checkbox cell toggle pattern supports IToggleProvider");

        ToggleState toggleState = ToggleState_Off;
        RequireSucceeded(checkboxTogglePattern->get_ToggleState(&toggleState), "grid checkbox cell toggle-state lookup succeeds");
        Require(toggleState == ToggleState_On, "grid checkbox cell toggle pattern reports the initial checked state");

        RequireSucceeded(checkboxTogglePattern->Toggle(), "grid checkbox cell toggle pattern can toggle the visible checkbox");
        Require(delegate.toggleCount == 1u && delegate.lastToggleRow == 0u && delegate.lastToggleColumn == 0u && ! delegate.lastToggleChecked,
                "grid checkbox cell toggle pattern routes through the shared delegate checkbox path");
        Require(! gridModel.IsChecked(0u), "grid checkbox cell toggle pattern updates the backing model state");
        Require(ReadProviderLongProperty(*checkboxCellSimple.get(), UIA_ToggleToggleStatePropertyId, "grid checkbox cell toggle state updates after Toggle") ==
                    ToggleState_Off,
                "grid checkbox cell accessibility provider reports the toggled unchecked state");
    }

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"Rules");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        CheckboxGridModel gridModel(0u);
        gridModel.SetRows({CheckboxGridModel::Row{L"Rule A", false, false}});
        RecordingCheckboxGridDelegate delegate(gridModel);
        grid->SetModel(&gridModel);
        grid->SetDelegate(&delegate);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "disabled grid checkbox accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> checkboxCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(checkboxCellRect.has_value(), "disabled grid checkbox exposes a visible cell rect for accessibility hit-testing");
        const float checkboxCellCenterXDip = (checkboxCellRect->left + checkboxCellRect->right) * 0.5f;
        const float checkboxCellCenterYDip = (checkboxCellRect->top + checkboxCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> checkboxCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  checkboxCellCenterXDip,
                                  checkboxCellCenterYDip,
                                  "disabled grid checkbox cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> checkboxCellSimple;
        RequireSucceeded(checkboxCellProvider.query_to(checkboxCellSimple.put()),
                         "disabled grid checkbox cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderLongProperty(*checkboxCellSimple.get(), UIA_ControlTypePropertyId, "disabled grid checkbox cell exposes UIA control type") ==
                    UIA_CheckBoxControlTypeId,
                "disabled grid checkbox cell accessibility provider reports checkbox control type");
        Require(ReadProviderStringProperty(*checkboxCellSimple.get(), UIA_NamePropertyId, "disabled grid checkbox cell exposes accessibility name") ==
                    L"[ ] Enabled",
                "disabled grid checkbox cell accessibility provider exposes the unchecked cell text");
        Require(! ReadProviderBoolProperty(*checkboxCellSimple.get(), UIA_IsEnabledPropertyId, "disabled grid checkbox cell exposes disabled state"),
                "disabled grid checkbox cell accessibility provider reports disabled state");

        wil::com_ptr_nothrow<IUnknown> checkboxToggleUnknown;
        RequireSucceeded(checkboxCellSimple->GetPatternProvider(UIA_TogglePatternId, checkboxToggleUnknown.put()),
                         "disabled grid checkbox cell toggle-pattern lookup succeeds");
        Require(checkboxToggleUnknown == nullptr, "disabled grid checkbox cell accessibility provider does not expose the toggle pattern");
        Require(delegate.toggleCount == 0u && ! gridModel.IsChecked(0u),
                "disabled grid checkbox cell accessibility provider leaves the backing checkbox state unchanged");
    }

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"Plugins");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        GridCellData pluginCell{};
        pluginCell.kind        = GridCellKind::IconText;
        pluginCell.iconText    = L"*";
        pluginCell.text        = L"Plugin";
        pluginCell.badgeText   = L"Beta";
        pluginCell.tooltipText = L"Plugin is disabled by policy.";
        SingleCellGridModel pluginModel(std::move(pluginCell));
        grid->SetModel(&pluginModel);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "grid infotip accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> pluginCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(pluginCellRect.has_value(), "grid exposes a visible infotip cell rect for accessibility hit-testing");
        const float pluginCellCenterXDip = (pluginCellRect->left + pluginCellRect->right) * 0.5f;
        const float pluginCellCenterYDip = (pluginCellRect->top + pluginCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> pluginCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  pluginCellCenterXDip,
                                  pluginCellCenterYDip,
                                  "grid infotip cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> pluginCellSimple;
        RequireSucceeded(pluginCellProvider.query_to(pluginCellSimple.put()), "grid infotip cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderStringProperty(*pluginCellSimple.get(), UIA_NamePropertyId, "grid infotip cell exposes accessibility name") == L"Plugin [Beta]",
                "grid infotip cell accessibility provider keeps icon and badge text in the accessible name");
        Require(ReadProviderStringProperty(*pluginCellSimple.get(), UIA_HelpTextPropertyId, "grid infotip cell exposes help text") ==
                    L"Plugin is disabled by policy.",
                "grid infotip cell accessibility provider exposes explicit tooltip text as UIA HelpText");
    }

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"States");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        StateImageColumnGridModel stateImageModel;
        grid->SetModel(&stateImageModel);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "grid state-image accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> stateImageCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(stateImageCellRect.has_value(), "grid exposes a visible state-image cell rect for accessibility hit-testing");
        const float stateImageCellCenterXDip = (stateImageCellRect->left + stateImageCellRect->right) * 0.5f;
        const float stateImageCellCenterYDip = (stateImageCellRect->top + stateImageCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> stateImageCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  stateImageCellCenterXDip,
                                  stateImageCellCenterYDip,
                                  "grid state-image cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> stateImageCellSimple;
        RequireSucceeded(stateImageCellProvider.query_to(stateImageCellSimple.put()),
                         "grid state-image cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderLongProperty(*stateImageCellSimple.get(), UIA_ControlTypePropertyId, "grid state-image cell exposes UIA control type") ==
                    UIA_ImageControlTypeId,
                "grid state-image cell accessibility provider reports image control type");
        Require(ReadProviderStringProperty(*stateImageCellSimple.get(), UIA_NamePropertyId, "grid state-image cell exposes accessibility name") == L"!",
                "grid state-image cell accessibility provider keeps the icon glyph as its accessible name");

        wil::com_ptr_nothrow<IUnknown> stateImageValueUnknown;
        RequireSucceeded(stateImageCellSimple->GetPatternProvider(UIA_ValuePatternId, stateImageValueUnknown.put()),
                         "grid state-image cell value-pattern lookup succeeds");
        Require(stateImageValueUnknown == nullptr, "grid state-image cell accessibility provider does not expose a text value pattern");
    }

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"Jobs");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        GridCellData progressCell{};
        progressCell.kind     = GridCellKind::Marquee;
        progressCell.text     = L"Halfway";
        progressCell.progress = 0.5f;
        SingleCellGridModel progressModel(std::move(progressCell));
        grid->SetModel(&progressModel);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "grid progress accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> progressCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(progressCellRect.has_value(), "grid exposes a visible progress cell rect for accessibility hit-testing");
        const float progressCellCenterXDip = (progressCellRect->left + progressCellRect->right) * 0.5f;
        const float progressCellCenterYDip = (progressCellRect->top + progressCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> progressCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  progressCellCenterXDip,
                                  progressCellCenterYDip,
                                  "grid progress cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> progressCellSimple;
        RequireSucceeded(progressCellProvider.query_to(progressCellSimple.put()),
                         "grid progress cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderLongProperty(*progressCellSimple.get(), UIA_ControlTypePropertyId, "grid progress cell exposes UIA control type") ==
                    UIA_ProgressBarControlTypeId,
                "grid progress cell accessibility provider reports progress-bar control type");
        Require(ReadProviderStringProperty(*progressCellSimple.get(), UIA_NamePropertyId, "grid progress cell exposes accessibility name") == L"Halfway",
                "grid progress cell accessibility provider exposes the determinate progress label");
        Require(ReadProviderBoolProperty(*progressCellSimple.get(), UIA_ValueIsReadOnlyPropertyId, "grid progress cell exposes read-only state"),
                "grid progress cell accessibility provider reports read-only range semantics");

        wil::com_ptr_nothrow<IUnknown> rangeValueUnknown;
        RequireSucceeded(progressCellSimple->GetPatternProvider(UIA_RangeValuePatternId, rangeValueUnknown.put()),
                         "grid progress cell range-value lookup succeeds");
        Require(rangeValueUnknown != nullptr, "grid progress cell accessibility provider exposes the range-value pattern");
        wil::com_ptr_nothrow<IRangeValueProvider> rangeValuePattern;
        RequireSucceeded(rangeValueUnknown.query_to(rangeValuePattern.put()), "grid progress cell range-value pattern supports IRangeValueProvider");

        double rangeValue       = 0.0;
        double rangeMinimum     = 0.0;
        double rangeMaximum     = 0.0;
        double rangeSmallChange = 1.0;
        double rangeLargeChange = 1.0;
        BOOL rangeReadOnly      = FALSE;
        RequireSucceeded(rangeValuePattern->get_Value(&rangeValue), "grid progress cell range-value query succeeds");
        RequireSucceeded(rangeValuePattern->get_Minimum(&rangeMinimum), "grid progress cell minimum query succeeds");
        RequireSucceeded(rangeValuePattern->get_Maximum(&rangeMaximum), "grid progress cell maximum query succeeds");
        RequireSucceeded(rangeValuePattern->get_SmallChange(&rangeSmallChange), "grid progress cell small-change query succeeds");
        RequireSucceeded(rangeValuePattern->get_LargeChange(&rangeLargeChange), "grid progress cell large-change query succeeds");
        RequireSucceeded(rangeValuePattern->get_IsReadOnly(&rangeReadOnly), "grid progress cell range read-only query succeeds");
        Require(rangeValue == 0.5 && rangeMinimum == 0.0 && rangeMaximum == 1.0,
                "grid progress cell range-value pattern reports the determinate 0..1 progress value");
        Require(rangeSmallChange == 0.0 && rangeLargeChange == 0.0 && rangeReadOnly == TRUE,
                "grid progress cell range-value pattern reports a read-only non-adjustable progress range");
    }
}

void TestAccessibilityProviderExposesSliderRangeValuePattern()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root         = std::make_unique<Panel>();
    auto* sliderLabel = root->AddChild<Label>(L"Opacity");
    sliderLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* slider = root->AddChild<Slider>();
    slider->SetBounds(D2D1::RectF(0.0f, 32.0f, 240.0f, 64.0f));
    slider->SetMinimum(10.0);
    slider->SetMaximum(90.0);
    slider->SetValue(42.0);
    slider->SetStep(2.0);
    slider->SetLargeStep(10.0);
    sliderLabel->SetMnemonicTarget(slider);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "slider accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> sliderProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 120.0f, 48.0f, "slider accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> sliderSimple;
    RequireSucceeded(sliderProvider.query_to(sliderSimple.put()), "slider accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*sliderSimple.get(), UIA_ControlTypePropertyId, "slider exposes UIA control type") == UIA_SliderControlTypeId,
            "slider accessibility provider reports slider control type");
    Require(ReadProviderStringProperty(*sliderSimple.get(), UIA_NamePropertyId, "slider exposes accessibility name") == L"Opacity",
            "slider accessibility provider uses its associated label as the accessible name");
    Require(! ReadProviderBoolProperty(*sliderSimple.get(), UIA_ValueIsReadOnlyPropertyId, "slider exposes writable range state"),
            "slider accessibility provider reports an adjustable range value");

    wil::com_ptr_nothrow<IUnknown> rangeValueUnknown;
    RequireSucceeded(sliderSimple->GetPatternProvider(UIA_RangeValuePatternId, rangeValueUnknown.put()), "slider range-value lookup succeeds");
    Require(rangeValueUnknown != nullptr, "slider accessibility provider exposes the range-value pattern");
    wil::com_ptr_nothrow<IRangeValueProvider> rangeValuePattern;
    RequireSucceeded(rangeValueUnknown.query_to(rangeValuePattern.put()), "slider range-value pattern supports IRangeValueProvider");

    double rangeValue       = 0.0;
    double rangeMinimum     = 0.0;
    double rangeMaximum     = 0.0;
    double rangeSmallChange = 0.0;
    double rangeLargeChange = 0.0;
    BOOL rangeReadOnly      = TRUE;
    RequireSucceeded(rangeValuePattern->get_Value(&rangeValue), "slider range-value query succeeds");
    RequireSucceeded(rangeValuePattern->get_Minimum(&rangeMinimum), "slider minimum query succeeds");
    RequireSucceeded(rangeValuePattern->get_Maximum(&rangeMaximum), "slider maximum query succeeds");
    RequireSucceeded(rangeValuePattern->get_SmallChange(&rangeSmallChange), "slider small-change query succeeds");
    RequireSucceeded(rangeValuePattern->get_LargeChange(&rangeLargeChange), "slider large-change query succeeds");
    RequireSucceeded(rangeValuePattern->get_IsReadOnly(&rangeReadOnly), "slider read-only query succeeds");
    Require(rangeValue == 42.0 && rangeMinimum == 10.0 && rangeMaximum == 90.0, "slider range-value pattern reports the configured min/max/value");
    Require(rangeSmallChange == 2.0 && rangeLargeChange == 10.0 && rangeReadOnly == FALSE, "slider range-value pattern reports the configured step values");

    RequireSucceeded(rangeValuePattern->SetValue(68.0), "slider range-value SetValue succeeds");
    Require(slider->GetValue() == 68.0, "slider range-value SetValue updates the underlying control value");
}

} // namespace

void RunAccessibilityTests()
{
    TestAccessibilityTargetPublishesImmutableSnapshotBeforeTreeTeardown();
    TestAccessibilityLiveHostResolutionIsWindowThreadOnly();
    TestAccessibilityProviderTraversalSurvivesConcurrentRootReplacement();
    TestAccessibilityTreeHitTestUsesCheapVisibleIndexLookup();
    TestAccessibilityProviderFactoriesUseSharedMakeProviderHelper();
    TestAccessibilityRuntimeIdsUseSharedBuilder();
    TestAccessibilityPatternDispatchUsesSharedQueryPattern();
    TestAccessibilityElementProviderFromPointUsesSnapshot();
    TestAccessibilityBoundingRectangleUsesSnapshot();
    TestAccessibilityGridPatternReadsUseSnapshots();
    TestAccessibilityControlStateReadsUseSnapshots();
    TestAccessibilityGridCellValueReadsUseSnapshots();
    TestAccessibilityGridRowPropertyReadsUseSnapshots();
    TestAccessibilityTreeGridSelectionReadsUseSnapshots();
    TestAccessibilityTreeItemStateReadsUseSnapshots();
    TestAccessibilityTextPatternDocumentAndSelectionReadsUseSnapshots();
    TestAccessibilityTextEditCompositionReadsUseSnapshots();
    TestAccessibilityPasswordRevealButtonReadsUseSnapshots();
    TestAccessibilityNavigateUsesSnapshot();
    TestAccessibilityTreeSnapshotHitRecordsUseVisibleGeometryOnly();
    TestAttachedWindowHostWmGetObjectReturnsAccessibilityProvider();
    TestAccessibilityRootRuntimeIdIncludesProviderSpecificValues();
    TestAccessibilityProviderExposesInvokeToggleAndLabeledValuePatterns();
    TestAccessibilityProviderRefreshesButtonSemanticProperties();
    TestAccessibilityProviderRefreshesLabelAssociations();
    TestAccessibilityProviderExposesDirectSemanticRootControls();
    TestAccessibilityLabelOnlyRootDoesNotUseDirectSemanticRootCollapse();
    TestAccessibilityDirectSemanticRootMatchesUiAutomationClientTree();
    TestAccessibilityDirectSemanticRootTreeSelectionMatchesUiAutomationClientTree();
    TestAccessibilityProviderReportsFocusedControl();
    TestAccessibilityProviderMasksPasswordTextFieldValue();
    TestAccessibilityProviderExposesMaskedRevealButton();
    TestAccessibilityProviderExposesTextPatternForTextField();
    TestAccessibilityTextFieldSimpleRangeBoundingRectanglesUseCaretGeometry();
    TestAccessibilityTextFieldMultilineRangeFromPointUsesNativeHitTest();
    TestAccessibilityTextRangeFromPointDispatchesToWindowThread();
    TestAccessibilityTextFieldMultilineSameLineRangeBoundingRectanglesUseCaretGeometry();
    TestAccessibilityTextFieldMultilineRangeBoundingRectanglesUseLineCaretGeometry();
    TestAccessibilityTextFieldWrappedRangeBoundingRectanglesUseVisualLineGeometry();
    TestAccessibilityTextFieldWrappedCrossLineRangeBoundingRectanglesUseVisualLineGeometry();
    TestAccessibilityTextFieldWrappedLineMovementUsesVisualLines();
    TestAccessibilityTextRangeEndpointLineMovementDispatchesToWindowThread();
    TestAccessibilityTextRangeSpanLineMovementDispatchesToWindowThread();
    TestAccessibilityTextFieldSingleLineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry();
    TestAccessibilityTextFieldMultilineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry();
    TestAccessibilityEditableComboBoxSingleLineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry();
    TestAccessibilityProviderExposesTextPatternForEditableComboBox();
    TestAccessibilityTextRangeSelectDispatchesToWindowThread();
    TestAccessibilityTimedOutTextRangeSelectDoesNotExecuteLater();
    TestAccessibilityTakenTextRangeSelectExecutesOnlyOnce();
    TestAccessibilityDestroyWithPendingDispatchReturnsCancelled();
    TestAccessibilityTextRangeBoundingRectanglesDispatchesToWindowThread();
    TestAccessibilityTextRangeBoundingRectanglesTimeoutKeepsLateHandlerStorageAlive();
    TestAccessibilityProviderExposesNativeImeTextEditRanges();
    TestAccessibilityNativeTextInputRaisesTextAndTextEditEventCounters();
    TestAccessibilityGridSnapshotRebuildMeetsTenThousandRowSelectionBudget();
    TestAccessibilityProviderExposesTreeAndGridMetadata();
    TestAccessibilityProviderExposesTreeItemSelectionAndExpandCollapsePatterns();
    TestAccessibilityProviderExposesGridRowSelectionPatterns();
    TestAccessibilityProviderExposesHorizontallyScrolledGridRowStructure();
    TestAccessibilityProviderPointHitsClipAndTranslateScrollPanelChildren();
    TestAccessibilityProviderExposesGridCellToggleAndRangePatterns();
    TestAccessibilityProviderExposesSliderRangeValuePattern();
}
