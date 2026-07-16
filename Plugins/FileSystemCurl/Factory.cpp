#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <array>
#include <iterator>
#include <new>
#include <optional>
#include <string>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "FileSystemCurlResources.h"
#include "Helpers.h"

#include "FileSystemCurl.h"

#include "PlugInterfaces/FactoryImpl.h"

extern HINSTANCE g_hInstance;

namespace
{
struct LocalizedPluginMetaDataSet
{
    std::wstring ftpName;
    std::wstring ftpDescription;
    std::wstring sftpName;
    std::wstring sftpDescription;
    std::wstring scpName;
    std::wstring scpDescription;
    std::wstring imapName;
    std::wstring imapDescription;
    std::array<PluginMetaData, 4> plugins{};

    LocalizedPluginMetaDataSet()
    {
        LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMCURL_FTP_NAME, ftpName);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_FTP_DESCRIPTION, ftpDescription);
        LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMCURL_SFTP_NAME, sftpName);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_SFTP_DESCRIPTION, sftpDescription);
        LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMCURL_SCP_NAME, scpName);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_SCP_DESCRIPTION, scpDescription);
        LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMCURL_IMAP_NAME, imapName);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_IMAP_DESCRIPTION, imapDescription);

        // PluginMetaData keeps raw wchar_t* pointers, so bind them only after the
        // backing strings live in their final static storage.
        plugins = {{
            {
                .id          = L"builtin/file-system-ftp",
                .shortId     = L"ftp",
                .name        = ftpName.c_str(),
                .description = ftpDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
            {
                .id          = L"builtin/file-system-sftp",
                .shortId     = L"sftp",
                .name        = sftpName.c_str(),
                .description = sftpDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
            {
                .id          = L"builtin/file-system-scp",
                .shortId     = L"scp",
                .name        = scpName.c_str(),
                .description = scpDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
            {
                .id          = L"builtin/file-system-imap",
                .shortId     = L"imap",
                .name        = imapName.c_str(),
                .description = imapDescription.c_str(),
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
const PluginMetaData* GetMetaDataFtp() noexcept
{
    return &GetPluginMetaDataSet().plugins[0];
}
const PluginMetaData* GetMetaDataSftp() noexcept
{
    return &GetPluginMetaDataSet().plugins[1];
}
const PluginMetaData* GetMetaDataScp() noexcept
{
    return &GetPluginMetaDataSet().plugins[2];
}
const PluginMetaData* GetMetaDataImap() noexcept
{
    return &GetPluginMetaDataSet().plugins[3];
}

// Per-entry schema thunks.
const char* GetSchemaFtp() noexcept
{
    return GetFileSystemCurlStaticConfigurationSchema(FileSystemCurlProtocol::Ftp);
}
const char* GetSchemaSftp() noexcept
{
    return GetFileSystemCurlStaticConfigurationSchema(FileSystemCurlProtocol::Sftp);
}
const char* GetSchemaScp() noexcept
{
    return GetFileSystemCurlStaticConfigurationSchema(FileSystemCurlProtocol::Scp);
}
const char* GetSchemaImap() noexcept
{
    return GetFileSystemCurlStaticConfigurationSchema(FileSystemCurlProtocol::Imap);
}

// Per-entry creation thunks (bake in the plugin id for dispatch to the existing helper).
HRESULT CreateInstanceFtp(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemCurl(FileSystemCurlProtocol::Ftp, host);
    if (! instance)
        return E_OUTOFMEMORY;
    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}
HRESULT CreateInstanceSftp(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemCurl(FileSystemCurlProtocol::Sftp, host);
    if (! instance)
        return E_OUTOFMEMORY;
    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}
HRESULT CreateInstanceScp(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemCurl(FileSystemCurlProtocol::Scp, host);
    if (! instance)
        return E_OUTOFMEMORY;
    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}
HRESULT CreateInstanceImap(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemCurl(FileSystemCurlProtocol::Imap, host);
    if (! instance)
        return E_OUTOFMEMORY;
    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}

const PluginFactoryEntry kEntries[] = {
    {&GetMetaDataFtp, &GetSchemaFtp, &CreateInstanceFtp},
    {&GetMetaDataSftp, &GetSchemaSftp, &CreateInstanceSftp},
    {&GetMetaDataScp, &GetSchemaScp, &CreateInstanceScp},
    {&GetMetaDataImap, &GetSchemaImap, &CreateInstanceImap},
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
