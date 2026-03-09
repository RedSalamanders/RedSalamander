#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <iterator>
#include <new>
#include <array>
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
#include "Helpers.h"
#include "FileSystemCurlResources.h"

#include "FileSystemCurl.h"

extern HINSTANCE g_hInstance;

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, void** result)
{
    if (result == nullptr)
    {
        return E_POINTER;
    }

    *result = nullptr;

    if (riid == __uuidof(IFileSystem))
    {
        // Backward-compatible single-plugin entry point.
        // Prefer RedSalamanderEnumeratePlugins + RedSalamanderCreateEx for selecting ftp/sftp/scp.
        auto* instance = new (std::nothrow) FileSystemCurl(FileSystemCurlProtocol::Sftp, host);
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
    std::wstring ftpName;
    std::wstring ftpDescription;
    std::wstring sftpName;
    std::wstring sftpDescription;
    std::wstring scpName;
    std::wstring scpDescription;
    std::wstring imapName;
    std::wstring imapDescription;
    std::array<PluginMetaData, 4> plugins{};
};

[[nodiscard]] const LocalizedPluginMetaDataSet& GetPluginMetaDataSet() noexcept
{
    static const LocalizedPluginMetaDataSet data = [] {
        LocalizedPluginMetaDataSet value{};
        value.ftpName         = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_FTP_NAME);
        value.ftpDescription  = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_FTP_DESCRIPTION);
        value.sftpName        = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_SFTP_NAME);
        value.sftpDescription = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_SFTP_DESCRIPTION);
        value.scpName         = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_SCP_NAME);
        value.scpDescription  = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_SCP_DESCRIPTION);
        value.imapName        = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_IMAP_NAME);
        value.imapDescription = LoadStringResource(g_hInstance, IDS_FILESYSTEMCURL_IMAP_DESCRIPTION);
        value.plugins         = {{
            {
                .id          = L"builtin/file-system-ftp",
                .shortId     = L"ftp",
                .name        = value.ftpName.c_str(),
                .description = value.ftpDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = L"0.3",
            },
            {
                .id          = L"builtin/file-system-sftp",
                .shortId     = L"sftp",
                .name        = value.sftpName.c_str(),
                .description = value.sftpDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = L"0.3",
            },
            {
                .id          = L"builtin/file-system-scp",
                .shortId     = L"scp",
                .name        = value.scpName.c_str(),
                .description = value.scpDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = L"0.3",
            },
            {
                .id          = L"builtin/file-system-imap",
                .shortId     = L"imap",
                .name        = value.imapName.c_str(),
                .description = value.imapDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = L"0.3",
            },
        }};
        return value;
    }();

    return data;
}

static std::optional<FileSystemCurlProtocol> ProtocolFromPluginId(std::wstring_view pluginId) noexcept
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
