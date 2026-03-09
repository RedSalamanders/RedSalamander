#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <new>
#include <string>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"
#include "FileSystemGoogleDriveResources.h"

#include "FileSystemGoogleDrive.h"

extern HINSTANCE g_hInstance;

namespace
{
[[nodiscard]] const PluginMetaData& GetPluginMetaData() noexcept
{
    static const std::wstring name        = LoadStringResource(g_hInstance, IDS_FILESYSTEMGOOGLEDRIVE_NAME);
    static const std::wstring description = LoadStringResource(g_hInstance, IDS_FILESYSTEMGOOGLEDRIVE_DESCRIPTION);
    static const PluginMetaData metaData  = {
        .id          = L"builtin/file-system-gdrive",
        .shortId     = L"gdrive",
        .name        = name.c_str(),
        .description = description.c_str(),
        .author      = L"RedSalamander",
        .version     = L"0.1",
    };
    return metaData;
}
}

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, void** result)
{
    if (! result)
    {
        return E_POINTER;
    }

    *result = nullptr;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    auto* instance = new (std::nothrow) FileSystemGoogleDrive(host);
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = instance->QueryInterface(riid, result);
    instance->Release();
    return hr;
}

extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)
{
    if (! metaData || ! count)
    {
        return E_POINTER;
    }

    *metaData = nullptr;
    *count    = 0;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    *metaData = &GetPluginMetaData();
    *count    = 1;
    return S_OK;
}

extern "C" HRESULT __stdcall RedSalamanderCreateEx(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result)
{
    if (! result)
    {
        return E_POINTER;
    }

    *result = nullptr;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }
    if (! pluginId || pluginId[0] == L'\0')
    {
        return E_INVALIDARG;
    }
    if (! OrdinalString::EqualsNoCase(pluginId, GetPluginMetaData().id))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    return RedSalamanderCreate(riid, factoryOptions, host, result);
}
