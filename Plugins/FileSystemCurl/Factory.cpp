#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <iterator>
#include <new>
#include <optional>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"

#include "FileSystemCurl.h"

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
static constexpr PluginMetaData kFileSystemCurlPlugins[] = {
    {
        .id          = L"builtin/file-system-ftp",
        .shortId     = L"ftp",
        .name        = L"FTP",
        .description = L"FTP virtual file system.",
        .author      = L"RedSalamander",
        .version     = L"0.3",
    },
    {
        .id          = L"builtin/file-system-sftp",
        .shortId     = L"sftp",
        .name        = L"SFTP",
        .description = L"SFTP virtual file system (SSH File Transfer Protocol).",
        .author      = L"RedSalamander",
        .version     = L"0.3",
    },
    {
        .id          = L"builtin/file-system-scp",
        .shortId     = L"scp",
        .name        = L"SCP",
        .description = L"SCP virtual file system (secure copy over SSH).",
        .author      = L"RedSalamander",
        .version     = L"0.3",
    },
    {
        .id          = L"builtin/file-system-imap",
        .shortId     = L"imap",
        .name        = L"IMAP",
        .description = L"IMAP virtual mail file system.",
        .author      = L"RedSalamander",
        .version     = L"0.3",
    },
};

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

    *metaData = kFileSystemCurlPlugins;
    *count    = static_cast<unsigned int>(std::size(kFileSystemCurlPlugins));
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
