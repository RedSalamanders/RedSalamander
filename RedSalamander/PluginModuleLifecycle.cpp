#include "Framework.h"

#include "PluginModuleLifecycle.h"

#include "Helpers.h"
#include "LocalizationManager.h"

namespace PluginModuleLifecycle
{
namespace
{
using PluginShutdownExportFunc                     = void(__stdcall*)();
using PluginCanUnloadNowExportFunc                 = BOOL(__stdcall*)();
using PluginRetainModuleUntilProcessExitExportFunc = BOOL(__stdcall*)();
} // namespace

bool UnloadModule(wil::unique_hmodule& module, ModuleUnloadMode mode, const std::filesystem::path& path, std::wstring_view pluginKind) noexcept
{
    if (! module)
    {
        return true;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // C4191: unsafe conversion from FARPROC
    if (const auto shutdown = reinterpret_cast<PluginShutdownExportFunc>(GetProcAddress(module.get(), "RedSalamanderPluginShutdown")); shutdown != nullptr)
    {
        shutdown();
    }

    bool canUnloadNow = true;
    if (const auto canUnload = reinterpret_cast<PluginCanUnloadNowExportFunc>(GetProcAddress(module.get(), "RedSalamanderPluginCanUnloadNow"));
        canUnload != nullptr)
    {
        canUnloadNow = canUnload() != FALSE;
    }
#pragma warning(pop)

    if (! canUnloadNow && mode == ModuleUnloadMode::FreeLibrary)
    {
        Debug::Warning(L"{} '{}' deferred unload because RedSalamanderPluginCanUnloadNow returned FALSE.", pluginKind, path.wstring());
        return false;
    }

    if (! canUnloadNow)
    {
        Debug::Warning(L"{} '{}' still reported live work during process shutdown; leaving the module mapped for OS teardown.", pluginKind, path.wstring());
        static_cast<void>(module.release());
        return true;
    }

    Localization::UnregisterResourceOwner(module.get());

    bool retainUntilProcessExit = false;
    if (mode == ModuleUnloadMode::ProcessShutdown)
    {
#pragma warning(push)
#pragma warning(disable : 4191) // C4191: unsafe conversion from FARPROC
        if (const auto retain =
                reinterpret_cast<PluginRetainModuleUntilProcessExitExportFunc>(GetProcAddress(module.get(), "RedSalamanderPluginRetainModuleUntilProcessExit"));
            retain != nullptr)
        {
            retainUntilProcessExit = retain() != FALSE;
        }
#pragma warning(pop)
    }

    if (retainUntilProcessExit)
    {
        static_cast<void>(module.release());
    }
    else
    {
        module.reset();
    }
    return true;
}

bool PathsEqualNoCase(const std::filesystem::path& left, const std::filesystem::path& right) noexcept
{
    const std::wstring leftText  = left.wstring();
    const std::wstring rightText = right.wstring();
    return CompareStringOrdinal(leftText.c_str(), -1, rightText.c_str(), -1, TRUE) == CSTR_EQUAL;
}
} // namespace PluginModuleLifecycle
