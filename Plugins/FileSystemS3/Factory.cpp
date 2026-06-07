#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <array>
#include <new>
#include <optional>
#include <string>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move / unused inline Helpers / Deferencing NULL Pointer
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "FileSystemS3Resources.h"
#include "Helpers.h"

#include "FileSystemS3.h"

extern HINSTANCE g_hInstance;

namespace
{
struct LocalizedPluginMetaDataSet
{
    std::wstring s3Name;
    std::wstring s3Description;
    std::wstring s3TableName;
    std::wstring s3TableDescription;
    std::array<PluginMetaData, 2> plugins{};

    LocalizedPluginMetaDataSet()
    {
        LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMS3_NAME, s3Name);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMS3_DESCRIPTION, s3Description);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMS3TABLE_NAME, s3TableName);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMS3TABLE_DESCRIPTION, s3TableDescription);

        // PluginMetaData keeps raw wchar_t* pointers, so bind them only after the
        // backing strings live in their final static storage.
        plugins = {{
            {
                .id          = L"builtin/file-system-s3",
                .shortId     = L"s3",
                .name        = s3Name.c_str(),
                .description = s3Description.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
            {
                .id          = L"builtin/file-system-s3table",
                .shortId     = L"s3table",
                .name        = s3TableName.c_str(),
                .description = s3TableDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
        }};
    }
};

[[nodiscard]] const LocalizedPluginMetaDataSet& GetPluginMetaDataSet() noexcept
{
    static const LocalizedPluginMetaDataSet data;

    return data;
}

[[nodiscard]] std::optional<FileSystemS3Mode> ModeFromPluginId(std::wstring_view pluginId) noexcept
{
    if (pluginId == L"builtin/file-system-s3")
    {
        return FileSystemS3Mode::S3;
    }
    if (pluginId == L"builtin/file-system-s3table")
    {
        return FileSystemS3Mode::S3Table;
    }
    return std::nullopt;
}

[[nodiscard]] const char* GetPluginSchema(std::wstring_view pluginId) noexcept
{
    const auto mode = ModeFromPluginId(pluginId);
    if (! mode.has_value())
    {
        return nullptr;
    }

    return GetFileSystemS3StaticConfigurationSchema(mode.value());
}

HRESULT CreatePluginInstance(REFIID riid, IHost* host, std::wstring_view pluginId, void** result)
{
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    const auto mode = ModeFromPluginId(pluginId);
    if (! mode.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    auto* instance = new (std::nothrow) FileSystemS3(mode.value(), host);
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = instance->QueryInterface(riid, result);
    instance->Release();
    return hr;
}
} // namespace

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

    const auto& plugins = GetPluginMetaDataSet().plugins;
    *metaData           = plugins.data();
    *count              = static_cast<unsigned int>(plugins.size());
    return S_OK;
}

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, const wchar_t* pluginId, void** result)
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

    return CreatePluginInstance(riid, host, pluginId, result);
}

extern "C" HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8)
{
    if (! schemaJsonUtf8)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = nullptr;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }
    if (! pluginId || pluginId[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const char* schema = GetPluginSchema(pluginId);
    if (! schema)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    *schemaJsonUtf8 = schema;
    return S_OK;
}

extern "C" PLUGFACTORY_API void __stdcall RedSalamanderPluginShutdown() noexcept
{
    // AWS SDK lifetime is reference-counted by FileSystemS3 instances and I/O objects.
}

extern "C" PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginRetainModuleUntilProcessExit() noexcept
{
    return TRUE;
}
