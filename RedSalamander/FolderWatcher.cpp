#include <memory>
#include <new>

#include "FolderWatcher.h"

#include "Helpers.h"

FolderWatcher::FolderWatcher(wil::com_ptr<IFileSystemDirectoryWatch> directoryWatch, std::wstring folderPath, Callback callback)
    : _folderPath(std::move(folderPath)),
      _callback(std::move(callback)),
      _pluginWatch(std::move(directoryWatch)),
      _pluginCallback()
{
    _pluginCallback.SetOwner(this);
}

FolderWatcher::~FolderWatcher()
{
    Stop();
}

HRESULT STDMETHODCALLTYPE FolderWatcher::PluginCallback::FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification,
                                                                                    void* /*cookie*/) noexcept
{
    FolderWatcherNotification owned{};
    if (! notification || notification->sizeBytes != sizeof(FileSystemDirectoryChangeNotification))
    {
        if (_owner)
        {
            owned.overflow = true;
            _owner->OnPluginDirectoryChanged(std::move(owned));
        }
        return notification ? E_INVALIDARG : E_POINTER;
    }

    owned.overflow = notification->overflow != FALSE;
    if (notification->changes && notification->changeCount > 0)
    {
        owned.changes.reserve(notification->changeCount);
        for (unsigned long index = 0; index < notification->changeCount; ++index)
        {
            const FileSystemDirectoryChange& change = notification->changes[index];

            FolderWatcherNotification::Change ownedChange{};
            ownedChange.action = change.action;
            if (change.relativePath && change.relativePathSize > 0 && change.relativePathSize % sizeof(wchar_t) == 0)
            {
                const size_t charCount = static_cast<size_t>(change.relativePathSize) / sizeof(wchar_t);
                ownedChange.relativePath.assign(change.relativePath, change.relativePath + charCount);
            }
            owned.changes.push_back(std::move(ownedChange));
        }
    }

    if (_owner)
    {
        _owner->OnPluginDirectoryChanged(std::move(owned));
    }

    return S_OK;
}

HRESULT FolderWatcher::Start() noexcept
{
    std::unique_lock lock(_mutex);

    if (_running.load(std::memory_order_acquire))
    {
        return S_OK;
    }

    if (_folderPath.empty())
    {
        return E_INVALIDARG;
    }

    _stopping.store(false, std::memory_order_release);

    wil::com_ptr<IFileSystemDirectoryWatch> watch = _pluginWatch;
    const std::wstring folderPath                 = _folderPath;

    lock.unlock();

    const HRESULT hr = watch ? watch->WatchDirectory(folderPath.c_str(), &_pluginCallback, this) : E_POINTER;
    if (FAILED(hr))
    {
        Debug::Warning(L"FolderWatcher: Failed to start plugin watch for '{}' (hr=0x{:08X})", folderPath, static_cast<unsigned long>(hr));
        return hr;
    }

    _running.store(true, std::memory_order_release);
    return S_OK;
}

void FolderWatcher::Stop() noexcept
{
    std::unique_lock lock(_mutex);

    if (! _running.load(std::memory_order_acquire))
    {
        return;
    }

    _stopping.store(true, std::memory_order_release);

    wil::com_ptr<IFileSystemDirectoryWatch> watch = _pluginWatch;
    const std::wstring folderPath                 = _folderPath;
    _running.store(false, std::memory_order_release);

    lock.unlock();

    if (watch)
    {
        static_cast<void>(watch->UnwatchDirectory(folderPath.c_str()));
    }
}

void FolderWatcher::OnPluginDirectoryChanged(FolderWatcherNotification notification) noexcept
{
    if (_stopping.load(std::memory_order_acquire))
    {
        return;
    }

    if (notification.overflow)
    {
        _overflowCount.fetch_add(1ull, std::memory_order_relaxed);
        const ULONGLONG nowTick               = GetTickCount64();
        const ULONGLONG lastTick              = _lastOverflowLogTick.load(std::memory_order_acquire);
        constexpr ULONGLONG kMinLogIntervalMs = 5'000ull;
        if (lastTick == 0 || (nowTick >= lastTick && (nowTick - lastTick) >= kMinLogIntervalMs))
        {
            _lastOverflowLogTick.store(nowTick, std::memory_order_release);
            Debug::Warning(L"FolderWatcher: directory watch overflow for '{}' (events dropped/coalesced); scheduling full refresh", _folderPath);
        }
    }

    if (_callback)
    {
        // Preserve plugin callback ordering here so rename old/new pairs stay adjacent when the cache routes the batch.
        // The host intentionally avoids adding a second async hop after IFileSystemDirectoryWatch delivery.
        _callback(notification);
    }
}
