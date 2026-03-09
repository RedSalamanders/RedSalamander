#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <new>
#include <array>
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
#include "Helpers.h"
#include "FileSystemS3Resources.h"

#include "FileSystemS3.h"

extern HINSTANCE g_hInstance;

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, void** result)
{
    if (! result)
    {
        return E_POINTER;
    }

    *result = nullptr;

    if (riid == __uuidof(IFileSystem))
    {
        // Backward-compatible single-plugin entry point.
        // Prefer RedSalamanderEnumeratePlugins + RedSalamanderCreateEx to select S3 vs S3 Table.
        auto* instance = new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3, host);
        if (! instance)
        {
            return E_OUTOFMEMORY;
        }

        const HRESULT hr = instance->QueryInterface(riid, result);
        instance->Release();
        return hr;
    }

    return E_NOINTERFACE;
}

namespace
{
struct LocalizedPluginMetaDataSet
{
    std::wstring s3Name;
    std::wstring s3Description;
    std::wstring s3TableName;
    std::wstring s3TableDescription;
    std::array<PluginMetaData, 2> plugins{};
};

[[nodiscard]] const LocalizedPluginMetaDataSet& GetPluginMetaDataSet() noexcept
{
    static const LocalizedPluginMetaDataSet data = [] {
        LocalizedPluginMetaDataSet value{};
        value.s3Name             = LoadStringResource(g_hInstance, IDS_FILESYSTEMS3_NAME);
        value.s3Description      = LoadStringResource(g_hInstance, IDS_FILESYSTEMS3_DESCRIPTION);
        value.s3TableName        = LoadStringResource(g_hInstance, IDS_FILESYSTEMS3TABLE_NAME);
        value.s3TableDescription = LoadStringResource(g_hInstance, IDS_FILESYSTEMS3TABLE_DESCRIPTION);
        value.plugins            = {{
            {
                .id          = L"builtin/file-system-s3",
                .shortId     = L"s3",
                .name        = value.s3Name.c_str(),
                .description = value.s3Description.c_str(),
                .author      = L"RedSalamander",
                .version     = L"0.1",
            },
            {
                .id          = L"builtin/file-system-s3table",
                .shortId     = L"s3table",
                .name        = value.s3TableName.c_str(),
                .description = value.s3TableDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = L"0.1",
            },
        }};
        return value;
    }();

    return data;
}

static std::optional<FileSystemS3Mode> ModeFromPluginId(std::wstring_view pluginId) noexcept
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

extern "C" HRESULT __stdcall RedSalamanderCreateEx(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, const wchar_t* pluginId, void** result)
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
