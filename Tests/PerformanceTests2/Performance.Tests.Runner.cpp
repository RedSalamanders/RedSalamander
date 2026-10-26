#include "Performance.Tests.pch.h"

#include "FileSystemPluginManager.h"
#include "PluginModuleLifecycle.h"
#include "SettingsStore.h"
#include "SplashScreen.h"
#include "TestSupport/TestSupport.Tests.FailureReports.h"
#include "ViewerPluginManager.h"

#include <algorithm>
#include <array>
#include <stop_token>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/resource.h>
#pragma warning(pop)

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

// The test host loads this module, so there is no entry point of ours to route a failed Debug-CRT or STL check away from the modal
// Abort/Retry/Ignore box: the module initializer runs before any test of the module and does it for the host process.
TEST_MODULE_INITIALIZE(PerformanceTests2ModuleRouteFailureReports)
{
    RedSalamander::TestSupport::RouteFailureReportsAwayFromDialogs();
}

namespace
{
const int kFakeViewerModuleAnchor = 0;

struct FakeViewerEntry
{
    FakeViewerEntry()                                      = default;
    FakeViewerEntry(const FakeViewerEntry&)                = delete;
    FakeViewerEntry& operator=(const FakeViewerEntry&)     = delete;
    FakeViewerEntry(FakeViewerEntry&&) noexcept            = default;
    FakeViewerEntry& operator=(FakeViewerEntry&&) noexcept = default;

    enum class Origin : uint8_t
    {
        Embedded,
    };

    Origin origin = Origin::Embedded;
    std::filesystem::path path;
    std::wstring factoryPluginId;
    std::wstring id;
    std::wstring shortId;
    std::wstring name;
    std::wstring description;
    std::wstring author;
    std::wstring version;
    bool loadable       = true;
    bool disabled       = false;
    bool unloadDeferred = false;
    std::wstring loadError;
    wil::unique_hmodule module;
};

[[nodiscard]] wil::unique_hmodule AcquireFakeViewerModulePin() noexcept
{
    HMODULE module      = nullptr;
    const BOOL acquired = GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, reinterpret_cast<LPCWSTR>(&kFakeViewerModuleAnchor), &module);
    return acquired != FALSE ? wil::unique_hmodule(module) : wil::unique_hmodule{};
}

[[nodiscard]] FakeViewerEntry MakeFakeViewerEntry()
{
    FakeViewerEntry entry;
    entry.path            = LR"(C:\RedSalamanderSelfTest\ViewerFirebreak.dll)";
    entry.factoryPluginId = L"selftest/viewer-firebreak";
    entry.id              = entry.factoryPluginId;
    entry.shortId         = L"firebreak";
    entry.name            = L"Firebreak Viewer";
    entry.description     = L"Plugin lifecycle regression fixture";
    entry.author          = L"RedSalamander";
    entry.version         = L"1";
    entry.module          = AcquireFakeViewerModulePin();
    return entry;
}
} // namespace

namespace PerformanceTests2
{
TEST_CLASS(PerformanceTests2){
    public : TEST_METHOD(SplashScreenCloseGuardTriggersWhenCloseEventWasSignaled){wil::unique_handle closeEvent(::CreateEventW(nullptr, TRUE, TRUE, nullptr));
Assert::IsNotNull(closeEvent.get(), L"failed to create close event for splash guard test");

std::stop_source stopSource;
Assert::IsTrue(SplashScreen::Detail::ShouldAbortPendingOpen(stopSource.get_token(), closeEvent.get()),
               L"signaled close event should suppress splash open even after the delay elapsed");
} // namespace PerformanceTests2

TEST_METHOD(FileSystemPluginManagerInitializeFailsWhenNoPluginsAreDiscovered)
{
    Common::Settings::Settings settings;
    auto& manager = FileSystemPluginManager::GetInstance();
    manager.Shutdown(settings);

    const HRESULT hr = manager.Initialize(settings);

    Assert::IsTrue(FAILED(hr), L"initialization must fail when discovery resolves to zero file-system plugins");
    manager.Shutdown(settings);
}

TEST_METHOD(ViewerPluginManagerInitializeFailsWhenNoPluginsAreDiscovered)
{
    Common::Settings::Settings settings;
    auto& manager = ViewerPluginManager::GetInstance();
    manager.Shutdown(settings);

    const HRESULT hr = manager.Initialize(settings);

    Assert::IsTrue(FAILED(hr), L"initialization must fail when discovery resolves to zero viewer plugins");
    manager.Shutdown(settings);
}

TEST_METHOD(PluginModuleLifecycleDefersAndHealsFakeViewer)
{
    FakeViewerEntry viewer = MakeFakeViewerEntry();
    Assert::IsNotNull(viewer.module.get(), L"failed to acquire the fake viewer module pin");

    std::vector<FakeViewerEntry> active;
    std::vector<FakeViewerEntry> deferred;
    active.push_back(std::move(viewer));

    bool busy                   = true;
    uint32_t unloadAttempts     = 0u;
    const auto tryRuntimeUnload = [&](FakeViewerEntry& entry, PluginModuleLifecycle::ModuleUnloadMode mode) noexcept
    {
        ++unloadAttempts;
        if (busy && mode == PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary)
        {
            PluginModuleLifecycle::MarkDeferred(entry);
            return false;
        }

        entry.module.reset();
        entry.unloadDeferred = false;
        return true;
    };

    PluginModuleLifecycle::UnloadAll(active, deferred, PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary, tryRuntimeUnload);
    active.clear();
    Assert::AreEqual<size_t>(1u, deferred.size(), L"busy fake viewer was not retained in the deferred set");
    Assert::IsNotNull(deferred.front().module.get(), L"busy fake viewer lost its module pin during deferral");
    Assert::IsTrue(PluginModuleLifecycle::IsPathDeferred(deferred, deferred.front().path), L"busy fake viewer path was not reported as deferred");

    const FakeViewerEntry placeholder = PluginModuleLifecycle::MakeDeferredPlaceholder(deferred.front());
    Assert::IsFalse(placeholder.loadable, L"deferred fake viewer placeholder remained loadable");
    Assert::IsTrue(placeholder.unloadDeferred, L"deferred fake viewer placeholder lost its deferred marker");
    Assert::AreEqual(
        PluginModuleLifecycle::kDeferredUnloadError.data(), placeholder.loadError.c_str(), L"deferred fake viewer placeholder did not expose the shared error");

    PluginModuleLifecycle::SweepDeferred(deferred, PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary, tryRuntimeUnload);
    Assert::AreEqual<size_t>(1u, deferred.size(), L"on-demand sweep dropped a viewer that was still busy");
    Assert::AreEqual<uint32_t>(2u, unloadAttempts, L"on-demand sweep did not retry the deferred viewer");

    busy = false;
    PluginModuleLifecycle::SweepDeferred(deferred, PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary, tryRuntimeUnload);
    Assert::IsTrue(deferred.empty(), L"healed fake viewer remained deferred after the on-demand sweep");
    Assert::AreEqual<uint32_t>(3u, unloadAttempts, L"healed fake viewer was not retried exactly once");
}

TEST_METHOD(PluginModuleLifecycleRetainsSharedDllAliasesAcrossPartialSweeps)
{
    // Synthetic paths are identity labels only; this test performs no filesystem
    // I/O. Real RAII module pins expose accidental ownership duplication/release.
    constexpr std::array<std::wstring_view, 5> ids{{L"alias-a", L"alias-b", L"alias-c", L"unrelated-d", L"unrelated-e"}};
    for (const bool reverseOrder : {false, true})
    {
        std::vector<FakeViewerEntry> active;
        std::vector<FakeViewerEntry> deferred;
        std::vector<FakeViewerEntry> metadata;
        for (size_t index = 0; index < ids.size(); ++index)
        {
            auto entry            = MakeFakeViewerEntry();
            entry.id              = ids[index];
            entry.factoryPluginId = entry.id;
            entry.shortId         = std::wstring(1u, static_cast<wchar_t>(static_cast<size_t>(L'a') + index));
            entry.disabled        = index == 0u;
            entry.path            = index < 3u    ? LR"(C:\RedSalamander.Perf\synthetic\Shared.dll)"
                                    : index == 3u ? LR"(C:\RedSalamander.Perf\synthetic\Unrelated.dll)"
                                                  : LR"(C:\RedSalamander.Perf\synthetic\OtherBusy.dll)";
            if (index == 1u)
            {
                entry.path = LR"(c:\redsalamander.perf\synthetic\SHARED.DLL)";
            }
            Assert::IsNotNull(entry.module.get(), L"failed to acquire an alias module pin");
            metadata.push_back(PluginModuleLifecycle::MakeDeferredPlaceholder(entry));
            active.push_back(std::move(entry));
        }
        if (reverseOrder)
        {
            std::reverse(active.begin(), active.end());
        }

        uint32_t phase = 0u;
        std::array<uint32_t, 5> releases{};
        const auto tryUnload = [&](FakeViewerEntry& entry, PluginModuleLifecycle::ModuleUnloadMode mode) noexcept
        {
            if (! entry.module)
            {
                entry.unloadDeferred = false;
                return true;
            }
            const size_t index = static_cast<size_t>(entry.shortId.front() - L'a');
            const bool busy    = (index == 1u && phase < 1u) || (index == 2u && phase < 3u) || (index == 4u && phase < 2u);
            if (busy && mode == PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary)
            {
                PluginModuleLifecycle::MarkDeferred(entry);
                return false;
            }
            ++releases[index];
            entry.module.reset();
            entry.unloadDeferred = false;
            return true;
        };
        const auto verifyDeferred = [&](size_t expectedAliases, size_t expectedPins)
        {
            Assert::AreEqual(expectedAliases, deferred.size(), L"logical aliases were lost or duplicated");
            size_t pins = 0u;
            for (const auto& entry : deferred)
            {
                pins += entry.module ? 1u : 0u;
                Assert::IsTrue(entry.unloadDeferred && ! entry.loadable, L"a retained alias was advertised as loadable");
                Assert::AreEqual(PluginModuleLifecycle::kDeferredUnloadError.data(), entry.loadError.c_str());
                const auto sameIdCount = std::count_if(deferred.begin(), deferred.end(), [&](const auto& other) { return other.id == entry.id; });
                Assert::AreEqual<size_t>(1u, static_cast<size_t>(sameIdCount), L"refresh duplicated an alias identity");
                const auto original = std::find_if(metadata.begin(), metadata.end(), [&](const auto& known) { return known.id == entry.id; });
                Assert::IsTrue(original != metadata.end(), L"retained an unknown alias");
                Assert::IsTrue(entry.path == original->path && entry.factoryPluginId == original->factoryPluginId && entry.shortId == original->shortId &&
                                   entry.name == original->name && entry.description == original->description && entry.author == original->author &&
                                   entry.version == original->version && entry.origin == original->origin && entry.disabled == original->disabled,
                               L"alias recovery metadata changed during deferral");
            }
            Assert::AreEqual(expectedPins, pins, L"alias retention changed real module ownership");
            Assert::IsFalse(PluginModuleLifecycle::IsPathDeferred(deferred, metadata[3].path), L"an unrelated unloaded DLL was retained");
        };

        PluginModuleLifecycle::UnloadAll(active, deferred, PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary, tryUnload);
        active.clear();
        verifyDeferred(4u, 3u);
        Assert::AreEqual<uint32_t>(1u, releases[0], L"the already released alias must remain visible without a module pin");

        for (phase = 0u; phase < 3u; ++phase)
        {
            for (uint32_t refresh = 0u; refresh < 3u; ++refresh)
            {
                PluginModuleLifecycle::SweepDeferred(deferred, PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary, tryUnload);
                for (const auto& entry : deferred)
                {
                    active.push_back(PluginModuleLifecycle::MakeDeferredPlaceholder(entry));
                }
                PluginModuleLifecycle::UnloadAll(active, deferred, PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary, tryUnload);
                active.clear();
                verifyDeferred(phase < 2u ? 4u : 3u, static_cast<size_t>(3u - phase));
                Assert::IsTrue(PluginModuleLifecycle::IsPathDeferred(deferred, metadata[0].path), L"a partially released shared DLL became reloadable");
            }
        }

        PluginModuleLifecycle::SweepDeferred(deferred, PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary, tryUnload);
        verifyDeferred(0u, 0u);
        for (size_t index = 0; index < ids.size(); ++index)
        {
            Assert::AreEqual<uint32_t>(1u, releases[index], L"a real module reference was released more than once or leaked");
            Assert::IsFalse(PluginModuleLifecycle::IsPathDeferred(deferred, metadata[index].path), L"a fully released DLL remained blocked");
        }
        PluginModuleLifecycle::SweepDeferred(deferred, PluginModuleLifecycle::ModuleUnloadMode::FreeLibrary, tryUnload);
        verifyDeferred(0u, 0u);
    }
}

TEST_METHOD(PluginModuleLifecycleRetainsBusyFakeViewerAtProcessShutdown)
{
    FakeViewerEntry viewer = MakeFakeViewerEntry();
    Assert::IsNotNull(viewer.module.get(), L"failed to acquire the process-shutdown fake viewer module pin");
    PluginModuleLifecycle::MarkDeferred(viewer);

    std::vector<FakeViewerEntry> deferred;
    deferred.push_back(std::move(viewer));
    HMODULE retainedModule       = nullptr;
    bool usedProcessShutdownMode = false;
    PluginModuleLifecycle::SweepDeferred(deferred,
                                         PluginModuleLifecycle::ModuleUnloadMode::ProcessShutdown,
                                         [&](FakeViewerEntry& entry, PluginModuleLifecycle::ModuleUnloadMode mode) noexcept
    {
        usedProcessShutdownMode = mode == PluginModuleLifecycle::ModuleUnloadMode::ProcessShutdown;
        retainedModule          = entry.module.release();
        entry.unloadDeferred    = false;
        return true;
    });

    wil::unique_hmodule releaseAfterProof(retainedModule);
    Assert::IsTrue(usedProcessShutdownMode, L"busy viewer shutdown sweep used runtime-unload semantics");
    Assert::IsTrue(deferred.empty(), L"process-shutdown sweep left the busy fake viewer in the deferred vector");
    Assert::IsNotNull(retainedModule, L"process-shutdown sweep freed the busy fake viewer instead of retaining its module ownership");
}
}
;
} // namespace PerformanceTests2
