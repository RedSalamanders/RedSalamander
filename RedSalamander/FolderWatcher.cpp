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
    if (_owner)
    {
        const bool overflow = notification && notification->overflow;
        _owner->OnPluginDirectoryChanged(overflow);
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

void FolderWatcher::OnPluginDirectoryChanged(bool overflow) noexcept
{
    if (_stopping.load(std::memory_order_acquire))
    {
        return;
    }

    if (overflow)
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
        auto deferred = std::make_unique<Callback>(_callback);

        Callback* raw = deferred.release();
        const BOOL ok = TrySubmitThreadpoolCallback(
            [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
            {
                std::unique_ptr<Callback> callback(static_cast<Callback*>(context));
                if (callback && *callback)
                {
                    (*callback)();
                }
            },
            raw,
            nullptr);

        if (! ok)
        {
            // Reclaim ownership via unique_ptr destructor
            std::unique_ptr<Callback> reclaimed(raw);
            _callback();
        }
    }
}
