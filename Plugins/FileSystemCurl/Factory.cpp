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

[[nodiscard]] std::optional<FileSystemCurlProtocol> ProtocolFromPluginId(std::wstring_view pluginId) noexcept
{
    if (pluginId == L"builtin/file-system-ftp")
    {
        return FileSystemCurlProtocol::Ftp;
    }
    if (pluginId == L"builtin/file-system-sftp")
    {
        return FileSystemCurlProtocol::Sftp;
    }
    if (pluginId == L"builtin/file-system-scp")
    {
        return FileSystemCurlProtocol::Scp;
    }
    if (pluginId == L"builtin/file-system-imap")
    {
        return FileSystemCurlProtocol::Imap;
    }
    return std::nullopt;
}

[[nodiscard]] const char* GetPluginSchema(std::wstring_view pluginId) noexcept
{
    const auto protocol = ProtocolFromPluginId(pluginId);
    if (! protocol.has_value())
    {
        return nullptr;
    }

    return GetFileSystemCurlStaticConfigurationSchema(protocol.value());
}

HRESULT CreatePluginInstance(REFIID riid, IHost* host, std::wstring_view pluginId, void** result)
{
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    const auto protocol = ProtocolFromPluginId(pluginId);
    if (! protocol.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    auto* instance = new (std::nothrow) FileSystemCurl(protocol.value(), host);
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
