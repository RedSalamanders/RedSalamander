#include "FileSystem.Internal.h"

#include <array>
#include <deque>
#include <limits>

using namespace FileSystemInternal;

namespace
{
constexpr size_t kDefaultWatchBufferBytes = 64u * 1024u;
constexpr size_t kDefaultWatchBufferPool  = 4u;
constexpr size_t kMaxPendingWatchBuffers  = 4u;
constexpr DWORD kDefaultWatchFilter = FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_DIR_NAME | FILE_NOTIFY_CHANGE_ATTRIBUTES | FILE_NOTIFY_CHANGE_SIZE |
                                      FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_CREATION | FILE_NOTIFY_CHANGE_SECURITY;

FileSystemDirectoryChangeAction MapDirectoryWatchAction(DWORD action) noexcept
{
    switch (action)
    {
        case FILE_ACTION_ADDED: return FILESYSTEM_DIR_CHANGE_ADDED;
        case FILE_ACTION_REMOVED: return FILESYSTEM_DIR_CHANGE_REMOVED;
        case FILE_ACTION_MODIFIED: return FILESYSTEM_DIR_CHANGE_MODIFIED;
        case FILE_ACTION_RENAMED_OLD_NAME: return FILESYSTEM_DIR_CHANGE_RENAMED_OLD_NAME;
        case FILE_ACTION_RENAMED_NEW_NAME: return FILESYSTEM_DIR_CHANGE_RENAMED_NEW_NAME;
        default: return FILESYSTEM_DIR_CHANGE_UNKNOWN;
    }
}
} // namespace

class FileSystem::DirectoryWatch final
{
public:
    DirectoryWatch(std::wstring watchedPath, std::wstring extendedPath, IFileSystemDirectoryWatchCallback* callback, void* cookie) noexcept
        : _watchedPath(std::move(watchedPath)),
          _extendedPath(std::move(extendedPath)),
          _callback(callback),
          _cookie(cookie),
          _activeBuffer(kDefaultWatchBufferBytes),
          _filter(kDefaultWatchFilter)
    {
        _freeBuffers.reserve(kDefaultWatchBufferPool > 0 ? (kDefaultWatchBufferPool - 1u) : 0u);
        for (size_t i = 1u; i < kDefaultWatchBufferPool; ++i)
        {
            _freeBuffers.emplace_back(kDefaultWatchBufferBytes);
        }
    }

    DirectoryWatch(const DirectoryWatch&)            = delete;
    DirectoryWatch(DirectoryWatch&&)                 = delete;
    DirectoryWatch& operator=(const DirectoryWatch&) = delete;
    DirectoryWatch& operator=(DirectoryWatch&&)      = delete;

    ~DirectoryWatch()
    {
        Stop();
    }

    HRESULT Start() noexcept
    {
        std::unique_lock lock(_mutex);

        if (_running.load(std::memory_order_acquire))
        {
            return S_OK;
        }

        if (_extendedPath.empty())
        {
            return E_INVALIDARG;
        }

        _stopping.store(false, std::memory_order_release);

        _directory.reset(::CreateFileW(_extendedPath.c_str(),
                                       FILE_LIST_DIRECTORY,
                                       FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                       nullptr,
                                       OPEN_EXISTING,
                                       FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED,
                                       nullptr));
        if (! _directory)
        {
            const DWORD lastError = Debug::ErrorWithLastError(L"FileSystem: Failed to open directory handle for '{}'", _watchedPath);
            return HRESULT_FROM_WIN32(lastError);
        }

        _tpIo.reset(::CreateThreadpoolIo(_directory.get(), &DirectoryWatch::IoCallback, this, nullptr));
        if (! _tpIo)
        {
            const DWORD lastError = Debug::ErrorWithLastError(L"FileSystem: Failed to create thread pool I/O for '{}'", _watchedPath);
            _directory.reset();
            return HRESULT_FROM_WIN32(lastError);
        }

        _tpWork.reset(::CreateThreadpoolWork(&DirectoryWatch::WorkCallback, this, nullptr));
        if (! _tpWork)
        {
            const DWORD lastError = Debug::ErrorWithLastError(L"FileSystem: Failed to create thread pool work for '{}'", _watchedPath);
            _tpIo.reset();
            _directory.reset();
            return HRESULT_FROM_WIN32(lastError);
        }

        const HRESULT hr = IssueRead();
        if (FAILED(hr))
        {
            Debug::Warning(L"FileSystem: Failed to start directory watch for '{}' (hr=0x{:08X})", _watchedPath, static_cast<unsigned long>(hr));
            _tpWork.reset();
            _tpIo.reset();
            _directory.reset();
            std::memset(&_overlapped, 0, sizeof(_overlapped));
            _running.store(false, std::memory_order_release);
            return hr;
        }

        _watchStartTick.store(GetTickCount64(), std::memory_order_relaxed);
        _peakPendingEvents.store(0, std::memory_order_relaxed);
        _droppedPendingBuffers.store(0, std::memory_order_relaxed);
        _overflowEnqueued.store(0, std::memory_order_relaxed);
        _overflowDelivered.store(0, std::memory_order_relaxed);
        _changedDelivered.store(0, std::memory_order_relaxed);
        _queueLatencyMaxUs.store(0, std::memory_order_relaxed);
        _queueLatencyTotalUs.store(0, std::memory_order_relaxed);
        _queueLatencySampleCount.store(0, std::memory_order_relaxed);

        _running.store(true, std::memory_order_release);
        return S_OK;
    }

    void Stop() noexcept
    {
        std::unique_lock lock(_mutex);

        if (! _running.load(std::memory_order_acquire) && ! _tpIo && ! _directory)
        {
            return;
        }

        _stopping.store(true, std::memory_order_release);

        if (_directory)
        {
            ::CancelIoEx(_directory.get(), &_overlapped);
        }

        if (_tpIo)
        {
            PTP_IO tpIo = _tpIo.get();
            lock.unlock();
            ::WaitForThreadpoolIoCallbacks(tpIo, TRUE);
            lock.lock();
        }

        if (_tpWork)
        {
            PTP_WORK tpWork = _tpWork.get();
            lock.unlock();
            ::WaitForThreadpoolWorkCallbacks(tpWork, TRUE);
            lock.lock();
        }

        _workSubmitted  = false;
        _overflowQueued = false;
        for (auto& pending : _pendingEvents)
        {
            ReturnBufferLocked(std::move(pending.buffer));
        }
        _pendingEvents.clear();

        _tpWork.reset();
        _tpIo.reset();
        _directory.reset();
        std::memset(&_overlapped, 0, sizeof(_overlapped));
        _running.store(false, std::memory_order_release);

        const ULONGLONG startTick = _watchStartTick.exchange(0, std::memory_order_relaxed);
        if (startTick > 0 && Debug::Perf::IsEnabled())
        {
            const ULONGLONG endTick   = GetTickCount64();
            const ULONGLONG elapsedMs = (endTick >= startTick) ? (endTick - startTick) : 0;
            const uint64_t durationUs = static_cast<uint64_t>(elapsedMs) * 1000ull;

            const uint64_t peakQueueDepth    = _peakPendingEvents.load(std::memory_order_relaxed);
            const uint64_t droppedBuffers    = _droppedPendingBuffers.load(std::memory_order_relaxed);
            const uint64_t overflowEnqueued  = _overflowEnqueued.load(std::memory_order_relaxed);
            const uint64_t overflowDelivered = _overflowDelivered.load(std::memory_order_relaxed);
            const uint64_t changedDelivered  = _changedDelivered.load(std::memory_order_relaxed);

            const uint64_t latencyTotalUs = _queueLatencyTotalUs.load(std::memory_order_relaxed);
            const uint64_t latencySamples = _queueLatencySampleCount.load(std::memory_order_relaxed);
            const uint64_t avgLatencyUs   = (latencySamples > 0) ? (latencyTotalUs / latencySamples) : 0;
            const uint64_t maxLatencyUs   = _queueLatencyMaxUs.load(std::memory_order_relaxed);

            const HRESULT hr          = (overflowDelivered > 0 || droppedBuffers > 0) ? S_FALSE : S_OK;
            const std::wstring detail = std::format(L"{} changed={} overflowDelivered={} overflowEnqueued={} dropped={} peakQ={} maxQUs={} avgQUs={}",
                                                    _watchedPath,
                                                    changedDelivered,
                                                    overflowDelivered,
                                                    overflowEnqueued,
                                                    droppedBuffers,
                                                    peakQueueDepth,
                                                    maxLatencyUs,
                                                    avgLatencyUs);
            Debug::Perf::Emit(L"FileSystem.Watch", detail, durationUs, changedDelivered, overflowDelivered, hr);
        }
    }

private:
    struct PendingEvent
    {
        enum class Kind : uint8_t
        {
            Overflow,
            Changed,
        };

        Kind kind              = Kind::Overflow;
        ULONGLONG enqueuedTick = 0;
        std::vector<std::byte> buffer;
        size_t bytesTransferred = 0;
    };

    static void CALLBACK IoCallback([[maybe_unused]] PTP_CALLBACK_INSTANCE instance,
                                    void* context,
                                    [[maybe_unused]] void* overlapped,
                                    ULONG ioResult,
                                    ULONG_PTR numberOfBytesTransferred,
                                    [[maybe_unused]] PTP_IO io) noexcept
    {
        auto* self = static_cast<DirectoryWatch*>(context);
        if (! self)
        {
            return;
        }

        self->OnIoCompleted(ioResult, numberOfBytesTransferred);
    }

    static void CALLBACK WorkCallback([[maybe_unused]] PTP_CALLBACK_INSTANCE instance, void* context, [[maybe_unused]] PTP_WORK work) noexcept
    {
        auto* self = static_cast<DirectoryWatch*>(context);
        if (! self)
        {
            return;
        }

        self->ProcessPendingEvents();
    }

    HRESULT IssueRead() noexcept
    {
        if (! _directory || ! _tpIo)
        {
            return E_HANDLE;
        }

        if (_activeBuffer.empty())
        {
            return E_OUTOFMEMORY;
        }

        std::memset(&_overlapped, 0, sizeof(_overlapped));
        ::StartThreadpoolIo(_tpIo.get());

        DWORD bytesReturned = 0;
        const BOOL ok       = ::ReadDirectoryChangesW(
            _directory.get(), _activeBuffer.data(), static_cast<DWORD>(_activeBuffer.size()), FALSE, _filter, &bytesReturned, &_overlapped, nullptr);
        if (ok)
        {
            return S_OK;
        }

        const DWORD err = ::GetLastError();
        if (err == ERROR_IO_PENDING)
        {
            return S_OK;
        }

        ::CancelThreadpoolIo(_tpIo.get());
        return HRESULT_FROM_WIN32(err);
    }

    void OnIoCompleted(ULONG ioResult, ULONG_PTR numberOfBytesTransferred) noexcept
    {
        if (ioResult == ERROR_OPERATION_ABORTED)
        {
            return;
        }

        const ULONGLONG nowTick = GetTickCount64();
        bool submitWork         = false;

        {
            std::lock_guard guard(_mutex);

            if (_stopping.load(std::memory_order_acquire))
            {
                return;
            }

            if (ioResult != 0)
            {
                if (ioResult != ERROR_NOTIFY_ENUM_DIR)
                {
                    Debug::Warning(L"FileSystem: ReadDirectoryChangesW failed for '{}' (err={})", _watchedPath, ioResult);
                }
                EnqueueOverflowLocked();
            }
            else if (numberOfBytesTransferred > 0)
            {
                const size_t bytesTransferred = static_cast<size_t>(numberOfBytesTransferred);
                if (bytesTransferred == 0 || bytesTransferred > _activeBuffer.size())
                {
                    EnqueueOverflowLocked();
                }
                else if (_pendingEvents.size() >= kMaxPendingWatchBuffers)
                {
                    DropPendingBuffersLocked();
                    EnqueueOverflowLocked();
                }
                else
                {
                    PendingEvent pending{};
                    pending.kind             = PendingEvent::Kind::Changed;
                    pending.buffer           = std::move(_activeBuffer);
                    pending.bytesTransferred = bytesTransferred;
                    pending.enqueuedTick     = nowTick;
                    _pendingEvents.emplace_back(std::move(pending));

                    AcquireNextActiveBufferLocked();
                }
            }

            // Re-arm read immediately before dispatching callbacks.
            const HRESULT hr = IssueRead();
            if (FAILED(hr))
            {
                Debug::Warning(L"FileSystem: Failed to re-issue directory watch for '{}' (hr=0x{:08X})", _watchedPath, static_cast<unsigned long>(hr));
            }

            if (! _pendingEvents.empty() && _tpWork && ! _workSubmitted)
            {
                _workSubmitted = true;
                submitWork     = true;
            }

            const uint64_t depth = static_cast<uint64_t>(_pendingEvents.size());
            uint64_t peak        = _peakPendingEvents.load(std::memory_order_relaxed);
            while (depth > peak && ! _peakPendingEvents.compare_exchange_weak(peak, depth, std::memory_order_relaxed))
            {
            }
        }

        if (submitWork)
        {
            ::SubmitThreadpoolWork(_tpWork.get());
        }
    }

    void NotifyOverflow() noexcept
    {
        if (_stopping.load(std::memory_order_acquire))
        {
            return;
        }

        if (! _callback)
        {
            return;
        }

        if (_watchedPath.size() > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)))
        {
            return;
        }

        FileSystemDirectoryChangeNotification notification{};
        notification.watchedPath     = _watchedPath.c_str();
        notification.watchedPathSize = static_cast<unsigned long>(_watchedPath.size() * sizeof(wchar_t));
        notification.changes         = nullptr;
        notification.changeCount     = 0;
        notification.overflow        = TRUE;

        _callback->FileSystemDirectoryChanged(&notification, _cookie);
    }

    void NotifyChanged(const std::byte* bufferBegin, size_t bytesTransferred) noexcept
    {
        if (_stopping.load(std::memory_order_acquire))
        {
            return;
        }

        if (! _callback)
        {
            return;
        }

        if (_watchedPath.size() > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)))
        {
            return;
        }

        if (bytesTransferred == 0 || ! bufferBegin)
        {
            NotifyOverflow();
            return;
        }

        const std::byte* bufferEnd = bufferBegin + bytesTransferred;

        constexpr size_t kMaxChanges = 128u;
        std::array<FileSystemDirectoryChange, kMaxChanges> changes{};
        unsigned long changeCount = 0;
        BOOL overflow             = FALSE;

        const FILE_NOTIFY_INFORMATION* entry = reinterpret_cast<const FILE_NOTIFY_INFORMATION*>(bufferBegin);
        while (true)
        {
            const std::byte* entryBytes  = reinterpret_cast<const std::byte*>(entry);
            const std::byte* entryMinEnd = (entryBytes >= bufferBegin) ? (entryBytes + offsetof(FILE_NOTIFY_INFORMATION, FileName)) : bufferEnd;
            if (entryBytes < bufferBegin || entryMinEnd > bufferEnd)
            {
                overflow = TRUE;
                break;
            }

            const std::byte* fileNameBytes = reinterpret_cast<const std::byte*>(entry->FileName);
            const std::byte* fileNameEnd   = fileNameBytes + entry->FileNameLength;
            if (fileNameEnd > bufferEnd)
            {
                overflow = TRUE;
                break;
            }

            if (changeCount < static_cast<unsigned long>(changes.size()))
            {
                FileSystemDirectoryChange change{};
                change.action           = MapDirectoryWatchAction(entry->Action);
                change.relativePath     = entry->FileName;
                change.relativePathSize = entry->FileNameLength;
                changes[changeCount]    = change;
                ++changeCount;
            }
            else
            {
                overflow = TRUE;
                break;
            }

            if (entry->NextEntryOffset == 0)
            {
                break;
            }

            const std::byte* nextBytes = entryBytes + entry->NextEntryOffset;
            if (nextBytes <= entryBytes || nextBytes > bufferEnd)
            {
                overflow = TRUE;
                break;
            }

            entry = reinterpret_cast<const FILE_NOTIFY_INFORMATION*>(nextBytes);
        }

        FileSystemDirectoryChangeNotification notification{};
        notification.watchedPath     = _watchedPath.c_str();
        notification.watchedPathSize = static_cast<unsigned long>(_watchedPath.size() * sizeof(wchar_t));
        notification.changes         = changeCount > 0 ? changes.data() : nullptr;
        notification.changeCount     = changeCount;
        notification.overflow        = overflow;

        _callback->FileSystemDirectoryChanged(&notification, _cookie);
    }

    void AcquireNextActiveBufferLocked() noexcept
    {
        if (! _activeBuffer.empty())
        {
            return;
        }

        if (! _freeBuffers.empty())
        {
            _activeBuffer = std::move(_freeBuffers.back());
            _freeBuffers.pop_back();
            return;
        }

        _activeBuffer.resize(kDefaultWatchBufferBytes);
    }

    void ReturnBufferLocked(std::vector<std::byte> buffer) noexcept
    {
        if (buffer.empty())
        {
            return;
        }

        if (_freeBuffers.size() >= (kDefaultWatchBufferPool > 0 ? (kDefaultWatchBufferPool - 1u) : 0u))
        {
            return;
        }

        if (buffer.size() != kDefaultWatchBufferBytes)
        {
            buffer.resize(kDefaultWatchBufferBytes);
        }

        _freeBuffers.emplace_back(std::move(buffer));
    }

    void DropPendingBuffersLocked() noexcept
    {
        uint64_t dropped = 0;
        for (auto& pending : _pendingEvents)
        {
            if (pending.kind == PendingEvent::Kind::Changed)
            {
                ++dropped;
                ReturnBufferLocked(std::move(pending.buffer));
            }
        }
        _pendingEvents.clear();
        _overflowQueued = false;
        if (dropped > 0)
        {
            _droppedPendingBuffers.fetch_add(dropped, std::memory_order_relaxed);
        }
    }

    void EnqueueOverflowLocked() noexcept
    {
        if (_overflowQueued)
        {
            return;
        }

        PendingEvent pending{};
        pending.kind         = PendingEvent::Kind::Overflow;
        pending.enqueuedTick = GetTickCount64();
        _pendingEvents.emplace_back(std::move(pending));
        _overflowQueued = true;
        _overflowEnqueued.fetch_add(1, std::memory_order_relaxed);

        const uint64_t depth = static_cast<uint64_t>(_pendingEvents.size());
        uint64_t peak        = _peakPendingEvents.load(std::memory_order_relaxed);
        while (depth > peak && ! _peakPendingEvents.compare_exchange_weak(peak, depth, std::memory_order_relaxed))
        {
        }
    }

    void ProcessPendingEvents() noexcept
    {
        for (;;)
        {
            PendingEvent pending{};
            {
                std::lock_guard guard(_mutex);

                if (_stopping.load(std::memory_order_acquire))
                {
                    _workSubmitted = false;
                    return;
                }

                if (_pendingEvents.empty())
                {
                    _workSubmitted = false;
                    return;
                }

                pending = std::move(_pendingEvents.front());
                _pendingEvents.pop_front();
                if (pending.kind == PendingEvent::Kind::Overflow)
                {
                    _overflowQueued = false;
                }
            }

            if (_stopping.load(std::memory_order_acquire))
            {
                if (pending.kind == PendingEvent::Kind::Changed)
                {
                    std::lock_guard guard(_mutex);
                    ReturnBufferLocked(std::move(pending.buffer));
                }
                return;
            }

            const ULONGLONG nowTick = GetTickCount64();
            if (pending.enqueuedTick > 0 && nowTick >= pending.enqueuedTick)
            {
                const uint64_t queuedUs = static_cast<uint64_t>(nowTick - pending.enqueuedTick) * 1000ull;
                _queueLatencyTotalUs.fetch_add(queuedUs, std::memory_order_relaxed);
                _queueLatencySampleCount.fetch_add(1, std::memory_order_relaxed);

                uint64_t maxUs = _queueLatencyMaxUs.load(std::memory_order_relaxed);
                while (queuedUs > maxUs && ! _queueLatencyMaxUs.compare_exchange_weak(maxUs, queuedUs, std::memory_order_relaxed))
                {
                }
            }

            if (pending.kind == PendingEvent::Kind::Overflow)
            {
                _overflowDelivered.fetch_add(1, std::memory_order_relaxed);
                NotifyOverflow();
            }
            else
            {
                _changedDelivered.fetch_add(1, std::memory_order_relaxed);
                NotifyChanged(pending.buffer.data(), pending.bytesTransferred);
            }

            if (pending.kind == PendingEvent::Kind::Changed)
            {
                std::lock_guard guard(_mutex);
                ReturnBufferLocked(std::move(pending.buffer));
            }
        }
    }

    std::wstring _watchedPath;
    std::wstring _extendedPath;

    IFileSystemDirectoryWatchCallback* _callback = nullptr;
    void* _cookie                                = nullptr;

    wil::unique_handle _directory;
    wil::unique_any<PTP_IO, decltype(&::CloseThreadpoolIo), ::CloseThreadpoolIo> _tpIo;
    wil::unique_any<PTP_WORK, decltype(&::CloseThreadpoolWork), ::CloseThreadpoolWork> _tpWork;
    std::vector<std::byte> _activeBuffer;
    std::vector<std::vector<std::byte>> _freeBuffers;
    std::deque<PendingEvent> _pendingEvents;
    bool _workSubmitted  = false;
    bool _overflowQueued = false;
    OVERLAPPED _overlapped{};
    DWORD _filter = 0;

    std::atomic<bool> _running{false};
    std::atomic<bool> _stopping{false};

    std::atomic<ULONGLONG> _watchStartTick{0};
    std::atomic<uint64_t> _peakPendingEvents{0};
    std::atomic<uint64_t> _droppedPendingBuffers{0};
    std::atomic<uint64_t> _overflowEnqueued{0};
    std::atomic<uint64_t> _overflowDelivered{0};
    std::atomic<uint64_t> _changedDelivered{0};
    std::atomic<uint64_t> _queueLatencyMaxUs{0};
    std::atomic<uint64_t> _queueLatencyTotalUs{0};
    std::atomic<uint64_t> _queueLatencySampleCount{0};
    std::mutex _mutex;
};

FileSystem::FileSystem()
{
    _metaData.id          = kPluginId;
    _metaData.shortId     = kPluginShortId;
    _metaData.name        = kPluginName;
    _metaData.description = kPluginDescription;
    _metaData.author      = kPluginAuthor;
    _metaData.version     = kPluginVersion;

    {
        std::lock_guard lock(_stateMutex);
        _configurationJson = "{}";
        UpdateCapabilitiesJson();
    }
}

FileSystem::~FileSystem()
{
    std::vector<std::unique_ptr<DirectoryWatch>> watchersToStop;
    {
        std::lock_guard lock(_watchMutex);
        watchersToStop.reserve(_directoryWatches.size());

        for (auto& entry : _directoryWatches)
        {
            if (entry.second)
            {
                watchersToStop.emplace_back(std::move(entry.second));
            }
        }

        _directoryWatches.clear();
    }

    for (auto& watcher : watchersToStop)
    {
        if (watcher)
        {
            watcher->Stop();
        }
    }
}

HRESULT STDMETHODCALLTYPE FileSystem::WatchDirectory(const wchar_t* path, IFileSystemDirectoryWatchCallback* callback, void* cookie) noexcept
{
    if (! path || ! callback)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring watchedPathText = path;
    std::wstring watchKey              = ToExtendedPath(watchedPathText);
    if (watchKey.empty())
    {
        return E_INVALIDARG;
    }

    {
        std::lock_guard lock(_watchMutex);
        if (_directoryWatches.contains(watchKey))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
    }

    auto watch       = std::make_unique<DirectoryWatch>(watchedPathText, watchKey, callback, cookie);
    const HRESULT hr = watch->Start();
    if (FAILED(hr))
    {
        return hr;
    }

    bool inserted = false;
    {
        std::lock_guard lock(_watchMutex);
        const auto insertResult = _directoryWatches.emplace(watchKey, std::move(watch));
        inserted                = insertResult.second;
    }

    if (! inserted)
    {
        if (watch)
        {
            watch->Stop();
        }
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::UnwatchDirectory(const wchar_t* path) noexcept
{
    if (! path)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring watchedPathText = path;
    const std::wstring watchKey        = ToExtendedPath(watchedPathText);
    if (watchKey.empty())
    {
        return E_INVALIDARG;
    }

    std::unique_ptr<DirectoryWatch> watch;
    {
        std::lock_guard lock(_watchMutex);
        auto it = _directoryWatches.find(watchKey);
        if (it == _directoryWatches.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        watch = std::move(it->second);
        _directoryWatches.erase(it);
    }

    if (watch)
    {
        watch->Stop();
    }

    return S_OK;
}
