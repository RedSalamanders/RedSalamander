#include "FileSystemS3.Internal.h"

namespace FsS3 = FileSystemS3Internal;

// FileSystemS3

FileSystemS3::FileSystemS3(FileSystemS3Mode mode, IHost* host) : _mode(mode)
{
    FsS3::AwsSdkLifetime::AddRef();

    switch (_mode)
    {
        case FileSystemS3Mode::S3:
            _metaData.id          = kPluginIdS3;
            _metaData.shortId     = kPluginShortIdS3;
            _metaData.name        = kPluginNameS3;
            _metaData.description = kPluginDescS3;
            break;
        case FileSystemS3Mode::S3Table:
            _metaData.id          = kPluginIdS3Table;
            _metaData.shortId     = kPluginShortIdS3Table;
            _metaData.name        = kPluginNameS3Table;
            _metaData.description = kPluginDescS3Table;
            break;
    }

    _metaData.author  = kPluginAuthor;
    _metaData.version = kPluginVersion;

    _configurationJson = "{}";
    _driveFileSystem   = _metaData.shortId ? _metaData.shortId : L"";

    if (host)
    {
        static_cast<void>(host->QueryInterface(__uuidof(IHostConnections), _hostConnections.put_void()));
    }
}

FileSystemS3::~FileSystemS3()
{
    FsS3::AwsSdkLifetime::Release();
}

HRESULT STDMETHODCALLTYPE FileSystemS3::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
    {
        *ppvObject = static_cast<IFileSystem*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemIO))
    {
        *ppvObject = static_cast<IFileSystemIO*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemDirectoryOperations))
    {
        *ppvObject = static_cast<IFileSystemDirectoryOperations*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(INavigationMenu))
    {
        *ppvObject = static_cast<INavigationMenu*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IDriveInfo))
    {
        *ppvObject = static_cast<IDriveInfo*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE FileSystemS3::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FileSystemS3::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = (_mode == FileSystemS3Mode::S3) ? kSchemaJsonS3 : kSchemaJsonS3Table;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *jsonUtf8 = kCapabilitiesJson;
    return S_OK;
}
