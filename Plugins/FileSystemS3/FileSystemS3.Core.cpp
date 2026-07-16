#include "FileSystemS3.Internal.h"
#include "FileSystemS3Resources.h"

namespace FsS3 = FileSystemS3Internal;

extern HINSTANCE g_hInstance;

namespace
{
[[nodiscard]] const wchar_t* LocalizedPluginName(FileSystemS3Mode mode) noexcept
{
    switch (mode)
    {
        case FileSystemS3Mode::S3:
        {
            static const std::wstring text = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMS3_NAME);
            return text.c_str();
        }
        case FileSystemS3Mode::S3Table:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMS3TABLE_NAME);
            return text.c_str();
        }
    }

    static const std::wstring fallback = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMS3_NAME);
    return fallback.c_str();
}

[[nodiscard]] const wchar_t* LocalizedPluginDescription(FileSystemS3Mode mode) noexcept
{
    switch (mode)
    {
        case FileSystemS3Mode::S3:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMS3_DESCRIPTION);
            return text.c_str();
        }
        case FileSystemS3Mode::S3Table:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMS3TABLE_DESCRIPTION);
            return text.c_str();
        }
    }

    static const std::wstring fallback = LoadStringResource(g_hInstance, IDS_FILESYSTEMS3_DESCRIPTION);
    return fallback.c_str();
}
} // namespace

// FileSystemS3

FileSystemS3::FileSystemS3(FileSystemS3Mode mode, IHost* host) : _mode(mode)
{
    FsS3::AwsSdkLifetime::AddRef();

    switch (_mode)
    {
        case FileSystemS3Mode::S3:
            _metaData.id          = kPluginIdS3;
            _metaData.shortId     = kPluginShortIdS3;
            _metaData.name        = LocalizedPluginName(FileSystemS3Mode::S3);
            _metaData.description = LocalizedPluginDescription(FileSystemS3Mode::S3);
            break;
        case FileSystemS3Mode::S3Table:
            _metaData.id          = kPluginIdS3Table;
            _metaData.shortId     = kPluginShortIdS3Table;
            _metaData.name        = LocalizedPluginName(FileSystemS3Mode::S3Table);
            _metaData.description = LocalizedPluginDescription(FileSystemS3Mode::S3Table);
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
    // Destroy cached AWS clients before shutting down the AWS SDK to avoid UAF/UB during teardown.
    {
        std::lock_guard lock(_stateMutex);
        _s3ClientsByCtxKey.clear();
    }
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

    if (riid == __uuidof(IFileSystemDirectoryWatch))
    {
        *ppvObject = static_cast<IFileSystemDirectoryWatch*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemAtomicWriter))
    {
        *ppvObject = static_cast<IFileSystemAtomicWriter*>(this);
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

std::wstring FileSystemS3::MakeWatchPathKey(std::wstring_view path) noexcept
{
    std::wstring key(path);
    for (wchar_t& ch : key)
    {
        ch = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
    }
    return key;
}

std::wstring FileSystemS3::RelativeWatchPath(std::wstring_view watchedPath, std::wstring_view fullPath) noexcept
{
    if (fullPath == watchedPath)
    {
        return {};
    }

    size_t offset = watchedPath.size();
    if (offset < fullPath.size() && fullPath[offset] == L'/')
    {
        ++offset;
    }
    return std::wstring(fullPath.substr(offset));
}

void FileSystemS3::EmitSyntheticWatchNotification(std::wstring_view watchedPath, const std::vector<SyntheticWatchChange>& changes, bool overflow) noexcept
{
    std::shared_ptr<SyntheticWatchRegistration> registration;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring watchedKey = MakeWatchPathKey(watchedPath);
        for (const auto& candidate : _syntheticWatches)
        {
            if (! candidate || ! candidate->active.load(std::memory_order_acquire))
            {
                continue;
            }
            if (candidate->watchedPathKey == watchedKey)
            {
                registration = candidate;
                break;
            }
        }
    }

    if (! registration || changes.empty())
    {
        return;
    }

    std::vector<FileSystemDirectoryChange> rawChanges;
    rawChanges.reserve(changes.size());
    for (const SyntheticWatchChange& change : changes)
    {
        rawChanges.push_back(FileSystemDirectoryChange{
            .action           = change.action,
            .relativePath     = change.relativePath.empty() ? nullptr : change.relativePath.c_str(),
            .relativePathSize = static_cast<unsigned long>(change.relativePath.size() * sizeof(wchar_t)),
        });
    }

    registration->inFlight.fetch_add(1u, std::memory_order_acq_rel);
    if (! registration->active.load(std::memory_order_acquire))
    {
        if (registration->inFlight.fetch_sub(1u, std::memory_order_acq_rel) == 1u)
        {
            std::lock_guard lock(registration->drainMutex);
            registration->drainCv.notify_all();
        }
        return;
    }

    FileSystemDirectoryChangeNotification notification{};
    notification.sizeBytes       = sizeof(notification);
    notification.watchedPath     = registration->watchedPath.c_str();
    notification.watchedPathSize = static_cast<unsigned long>(registration->watchedPath.size() * sizeof(wchar_t));
    notification.changes         = rawChanges.data();
    notification.changeCount     = static_cast<unsigned long>(rawChanges.size());
    notification.overflow        = overflow ? TRUE : FALSE;

    registration->callbackThreadId.store(GetCurrentThreadId(), std::memory_order_release);
    if (registration->callback)
    {
        static_cast<void>(registration->callback->FileSystemDirectoryChanged(&notification, registration->cookie));
    }
    registration->callbackThreadId.store(0u, std::memory_order_release);

    if (registration->inFlight.fetch_sub(1u, std::memory_order_acq_rel) == 1u)
    {
        std::lock_guard lock(registration->drainMutex);
        registration->drainCv.notify_all();
    }
}

void FileSystemS3::NotifySyntheticPathCreated(std::wstring_view fullPath) noexcept
{
    const std::wstring normalized          = FsS3::NormalizePluginPath(fullPath);
    {
        std::lock_guard lock(_stateMutex);
        _writableDirectoryValidationTicks.erase(normalized);
    }
    const std::filesystem::path parentPath = std::filesystem::path(normalized).parent_path();
    const std::wstring parent              = parentPath.empty() ? L"/" : FsS3::NormalizePluginPath(parentPath.native());

    std::vector<std::pair<std::wstring, std::vector<SyntheticWatchChange>>> notifications;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring parentKey = MakeWatchPathKey(parent);
        for (const auto& registration : _syntheticWatches)
        {
            if (! registration || ! registration->active.load(std::memory_order_acquire) || registration->watchedPathKey != parentKey)
            {
                continue;
            }

            std::wstring relative = RelativeWatchPath(registration->watchedPath, normalized);
            std::vector<SyntheticWatchChange> changes(1);
            changes[0].action       = FILESYSTEM_DIR_CHANGE_ADDED;
            changes[0].relativePath = std::move(relative);
            notifications.emplace_back(registration->watchedPath, std::move(changes));
        }
    }

    for (const auto& [watchedPath, changes] : notifications)
    {
        EmitSyntheticWatchNotification(watchedPath, changes, false);
    }
}

void FileSystemS3::RememberWritableDirectoryValidation(std::wstring_view fullPath) noexcept
{
    constexpr size_t kMaxCachedDirectories = 4096u;
    const std::wstring normalized = FsS3::NormalizePluginPath(fullPath);
    std::lock_guard lock(_stateMutex);
    if (_writableDirectoryValidationTicks.size() >= kMaxCachedDirectories)
    {
        _writableDirectoryValidationTicks.clear();
    }
    _writableDirectoryValidationTicks.insert_or_assign(normalized, GetTickCount64());
}

bool FileSystemS3::HasFreshWritableDirectoryValidation(std::wstring_view fullPath) noexcept
{
    constexpr ULONGLONG kValidationLifetimeMs = 60'000ull;
    const std::wstring normalized = FsS3::NormalizePluginPath(fullPath);
    const ULONGLONG nowTick = GetTickCount64();
    std::lock_guard lock(_stateMutex);
    const auto found = _writableDirectoryValidationTicks.find(normalized);
    if (found == _writableDirectoryValidationTicks.end())
    {
        return false;
    }
    if (nowTick < found->second || nowTick - found->second > kValidationLifetimeMs)
    {
        _writableDirectoryValidationTicks.erase(found);
        return false;
    }
    return true;
}

void FileSystemS3::NotifySyntheticPathDeleted(std::wstring_view fullPath) noexcept
{
    const std::wstring normalized          = FsS3::NormalizePluginPath(fullPath);
    const std::filesystem::path parentPath = std::filesystem::path(normalized).parent_path();
    const std::wstring parent              = parentPath.empty() ? L"/" : FsS3::NormalizePluginPath(parentPath.native());

    std::vector<std::pair<std::wstring, std::vector<SyntheticWatchChange>>> notifications;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring parentKey = MakeWatchPathKey(parent);
        for (const auto& registration : _syntheticWatches)
        {
            if (! registration || ! registration->active.load(std::memory_order_acquire) || registration->watchedPathKey != parentKey)
            {
                continue;
            }

            std::wstring relative = RelativeWatchPath(registration->watchedPath, normalized);
            std::vector<SyntheticWatchChange> changes(1);
            changes[0].action       = FILESYSTEM_DIR_CHANGE_REMOVED;
            changes[0].relativePath = std::move(relative);
            notifications.emplace_back(registration->watchedPath, std::move(changes));
        }
    }

    for (const auto& [watchedPath, changes] : notifications)
    {
        EmitSyntheticWatchNotification(watchedPath, changes, false);
    }
}

void FileSystemS3::NotifySyntheticFolderChanged(std::wstring_view folderPath) noexcept
{
    const std::wstring normalized = FsS3::NormalizePluginPath(folderPath);

    std::vector<std::pair<std::wstring, std::vector<SyntheticWatchChange>>> notifications;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring folderKey = MakeWatchPathKey(normalized);
        for (const auto& registration : _syntheticWatches)
        {
            if (! registration || ! registration->active.load(std::memory_order_acquire) || registration->watchedPathKey != folderKey)
            {
                continue;
            }

            std::vector<SyntheticWatchChange> changes(1);
            changes[0].action = FILESYSTEM_DIR_CHANGE_MODIFIED;
            notifications.emplace_back(registration->watchedPath, std::move(changes));
        }
    }

    for (const auto& [watchedPath, changes] : notifications)
    {
        EmitSyntheticWatchNotification(watchedPath, changes, false);
    }
}

HRESULT STDMETHODCALLTYPE FileSystemS3::WatchDirectory(const wchar_t* path, IFileSystemDirectoryWatchCallback* callback, void* cookie) noexcept
{
    if (! path || ! callback)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring normalized = FsS3::NormalizePluginPath(path);
    if (normalized.empty())
    {
        return E_INVALIDARG;
    }

    auto registration            = std::make_shared<SyntheticWatchRegistration>();
    registration->watchedPath    = normalized;
    registration->watchedPathKey = MakeWatchPathKey(normalized);
    registration->callback       = callback;
    registration->cookie         = cookie;

    std::lock_guard lock(_watchMutex);
    for (const auto& existing : _syntheticWatches)
    {
        if (existing && existing->active.load(std::memory_order_acquire) && existing->watchedPathKey == registration->watchedPathKey)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
    }

    _syntheticWatches.push_back(std::move(registration));
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::UnwatchDirectory(const wchar_t* path) noexcept
{
    if (! path)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring normalized = FsS3::NormalizePluginPath(path);
    if (normalized.empty())
    {
        return E_INVALIDARG;
    }

    std::shared_ptr<SyntheticWatchRegistration> registration;
    {
        std::lock_guard lock(_watchMutex);
        const std::wstring watchKey = MakeWatchPathKey(normalized);
        auto it = std::find_if(_syntheticWatches.begin(), _syntheticWatches.end(), [&](const std::shared_ptr<SyntheticWatchRegistration>& candidate) noexcept {
            return candidate && candidate->active.load(std::memory_order_acquire) && candidate->watchedPathKey == watchKey;
        });
        if (it == _syntheticWatches.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        registration = *it;
        _syntheticWatches.erase(it);
        registration->active.store(false, std::memory_order_release);
    }

    const DWORD currentThreadId       = GetCurrentThreadId();
    const unsigned int targetInFlight = registration->callbackThreadId.load(std::memory_order_acquire) == currentThreadId ? 1u : 0u;
    std::unique_lock drainLock(registration->drainMutex);
    registration->drainCv.wait(drainLock, [&]() noexcept { return registration->inFlight.load(std::memory_order_acquire) <= targetInFlight; });
    return S_OK;
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

    *schemaJsonUtf8 = StaticConfigurationSchema(_mode);
    return S_OK;
}

const char* GetFileSystemS3StaticConfigurationSchema(FileSystemS3Mode mode) noexcept
{
    return FileSystemS3::StaticConfigurationSchema(mode);
}

const char* FileSystemS3::StaticConfigurationSchema(FileSystemS3Mode mode) noexcept
{
    return (mode == FileSystemS3Mode::S3) ? kSchemaJsonS3 : kSchemaJsonS3Table;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *jsonUtf8 = (_mode == FileSystemS3Mode::S3) ? kCapabilitiesJsonS3 : kCapabilitiesJsonS3Table;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetTransferHints([[maybe_unused]] const wchar_t* path,
                                                         [[maybe_unused]] FileSystemOperation operationType,
                                                         [[maybe_unused]] FileSystemTransferEndpoint endpoint,
                                                         FileSystemTransferHints* hints) noexcept
{
    if (path == nullptr || path[0] == L'\0' || hints == nullptr)
    {
        return E_INVALIDARG;
    }
    if (hints->sizeBytes < sizeof(FileSystemTransferHints))
    {
        return E_INVALIDARG;
    }

    hints->latencyClass = FILESYSTEM_TRANSFER_LATENCY_CLOUD;
    hints->flags =
        FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS | FILESYSTEM_TRANSFER_HINT_PREFERS_SEQUENTIAL_IO | FILESYSTEM_TRANSFER_HINT_HIGH_METADATA_COST;
    hints->preferredBufferBytes      = 8u * 1024u * 1024u;
    hints->preferredProgressPeriodMs = 200u;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetStorageCharacteristics([[maybe_unused]] const wchar_t* path,
                                                                  FileSystemStorageCharacteristics* characteristics) noexcept
{
    if (path == nullptr || path[0] == L'\0' || characteristics == nullptr)
    {
        return E_INVALIDARG;
    }
    if (characteristics->sizeBytes < sizeof(FileSystemStorageCharacteristics))
    {
        return E_INVALIDARG;
    }

    characteristics->storageKind = FILESYSTEM_STORAGE_CLOUD;
    characteristics->flags = FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY | FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO | FILESYSTEM_STORAGE_FLAG_SUPPORTS_DEEP_QUEUE;
    characteristics->queueDepthHint               = 8u;
    characteristics->preferredCopyMoveConcurrency = 8u;
    characteristics->preferredDeleteConcurrency   = 8u;
    return S_OK;
}
