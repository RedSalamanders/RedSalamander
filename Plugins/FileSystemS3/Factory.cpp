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

#include "FileSystemS3.Internal.h"

#include "PlugInterfaces/FactoryImpl.h"

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

// Per-entry metadata thunks (return contiguous array elements).
const PluginMetaData* GetMetaDataS3() noexcept
{
    return &GetPluginMetaDataSet().plugins[0];
}
const PluginMetaData* GetMetaDataS3Table() noexcept
{
    return &GetPluginMetaDataSet().plugins[1];
}

// Per-entry schema thunks.
const char* GetSchemaS3() noexcept
{
    return GetFileSystemS3StaticConfigurationSchema(FileSystemS3Mode::S3);
}
const char* GetSchemaS3Table() noexcept
{
    return GetFileSystemS3StaticConfigurationSchema(FileSystemS3Mode::S3Table);
}

// Per-entry creation thunks.
HRESULT CreateInstanceS3(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3, host);
    if (! instance)
        return E_OUTOFMEMORY;
    const HRESULT initializationHr = instance->InitializationStatus();
    if (FAILED(initializationHr))
    {
        instance->Release();
        return initializationHr;
    }
    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}
HRESULT CreateInstanceS3Table(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3Table, host);
    if (! instance)
        return E_OUTOFMEMORY;
    const HRESULT initializationHr = instance->InitializationStatus();
    if (FAILED(initializationHr))
    {
        instance->Release();
        return initializationHr;
    }
    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}

const PluginFactoryEntry kEntries[] = {
    {&GetMetaDataS3, &GetSchemaS3, &CreateInstanceS3},
    {&GetMetaDataS3Table, &GetSchemaS3Table, &CreateInstanceS3Table},
};
} // namespace

extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)
{
    return FactoryEnumeratePlugins<IFileSystem>(kEntries, riid, metaData, count);
}

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result)
{
    return FactoryCreate<IFileSystem>(kEntries, riid, factoryOptions, host, pluginId, result);
}

extern "C" HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8)
{
    return FactoryGetConfigurationSchema<IFileSystem>(kEntries, riid, pluginId, schemaJsonUtf8);
}

extern "C" PLUGFACTORY_API void __stdcall RedSalamanderPluginShutdown() noexcept
{
    FileSystemS3Internal::SchedulePendingMultipartAbortCleanup();
    FileSystemS3Internal::AwsSdkLifetime::BeginShutdown();
}

extern "C" PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginCanUnloadNow() noexcept
{
    FileSystemS3Internal::SchedulePendingMultipartAbortCleanup();
    return FileSystemS3Internal::CanUnloadPendingMultipartAbortCleanup() && FileSystemS3Internal::AwsSdkLifetime::CanUnloadNow() ? TRUE : FALSE;
}

extern "C" PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginRetainModuleUntilProcessExit() noexcept
{
    return TRUE;
}
