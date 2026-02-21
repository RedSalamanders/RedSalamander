#pragma once
#include <atomic>
#include <functional>
#include <mutex>
#include <string>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/FileSystem.h"

class FolderWatcher
{
public:
    using Callback = std::function<void()>;

    FolderWatcher(wil::com_ptr<IFileSystemDirectoryWatch> directoryWatch, std::wstring folderPath, Callback callback);
    ~FolderWatcher();

    FolderWatcher(const FolderWatcher&)            = delete;
    FolderWatcher& operator=(const FolderWatcher&) = delete;
    FolderWatcher(FolderWatcher&&)                 = delete;
    FolderWatcher& operator=(FolderWatcher&&)      = delete;

    HRESULT Start() noexcept;
    void Stop() noexcept;

private:
    struct PluginCallback final : public IFileSystemDirectoryWatchCallback
    {
        PluginCallback() noexcept = default;

        void SetOwner(FolderWatcher* owner) noexcept
        {
            _owner = owner;
        }

        HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification, void* cookie) noexcept override;

    private:
        FolderWatcher* _owner = nullptr;
    };

    void OnPluginDirectoryChanged(bool overflow) noexcept;

    std::wstring _folderPath;
    Callback _callback;

    wil::com_ptr<IFileSystemDirectoryWatch> _pluginWatch;
    PluginCallback _pluginCallback;

    std::atomic<bool> _running{false};
    std::atomic<bool> _stopping{false};
    std::atomic<ULONGLONG> _lastOverflowLogTick{0};
    std::atomic<uint64_t> _overflowCount{0};
    std::mutex _mutex;
};
