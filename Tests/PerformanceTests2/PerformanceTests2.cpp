#include "pch.h"

#include "FileSystemPluginManager.h"
#include "SettingsStore.h"
#include "SplashScreen.h"
#include "ViewerPluginManager.h"

#include <stop_token>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/resource.h>
#pragma warning(pop)

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace PerformanceTests2
{
TEST_CLASS(PerformanceTests2)
{
public:
    TEST_METHOD(SplashScreenCloseGuardTriggersWhenCloseEventWasSignaled)
    {
        wil::unique_handle closeEvent(::CreateEventW(nullptr, TRUE, TRUE, nullptr));
        Assert::IsNotNull(closeEvent.get(), L"failed to create close event for splash guard test");

        std::stop_source stopSource;
        Assert::IsTrue(SplashScreen::Detail::ShouldAbortPendingOpen(stopSource.get_token(), closeEvent.get()),
                       L"signaled close event should suppress splash open even after the delay elapsed");
    }

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
};
} // namespace PerformanceTests2
