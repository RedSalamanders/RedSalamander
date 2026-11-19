#include "FileOperationTraversalPolicy.h"
#include "FileSystemCurl.Internal.h"

#include <chrono>
#include <condition_variable>
#include <system_error>
#include <thread>
#include <utility>
#include <vector>

using namespace FileSystemCurlInternal;

namespace
{
[[nodiscard]] bool ReadRequestReachesCommittedEnd(size_t bufferCapacity, unsigned long bytesToRead, uint64_t remainingCommittedBytes) noexcept
{
    return remainingCommittedBytes <= static_cast<uint64_t>(bufferCapacity) && static_cast<uint64_t>(bytesToRead) >= remainingCommittedBytes;
}

class TempFileReader final : public IFileReader
{
public:
    TempFileReader(wil::unique_hfile file, uint64_t sizeBytes) noexcept : _file(std::move(file)), _sizeBytes(sizeBytes)
    {
    }

    TempFileReader(const TempFileReader&)            = delete;
    TempFileReader(TempFileReader&&)                 = delete;
    TempFileReader& operator=(const TempFileReader&) = delete;
    TempFileReader& operator=(TempFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (! sizeBytes)
        {
            return E_POINTER;
        }

        *sizeBytes = _sizeBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (! newPosition)
        {
            return E_POINTER;
        }

        *newPosition = 0;

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        if (origin != FILE_BEGIN && origin != FILE_CURRENT && origin != FILE_END)
        {
            return E_INVALIDARG;
        }

        LARGE_INTEGER distance{};
        distance.QuadPart = offset;

        LARGE_INTEGER moved{};
        if (SetFilePointerEx(_file.get(), distance, &moved, origin) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (moved.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        *newPosition = static_cast<uint64_t>(moved.QuadPart);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (! bytesRead)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        if (bytesToRead == 0)
        {
            return S_OK;
        }

        if (! buffer)
        {
            return E_POINTER;
        }

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        DWORD read = 0;
        if (ReadFile(_file.get(), buffer, bytesToRead, &read, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        *bytesRead = static_cast<unsigned long>(read);
        return S_OK;
    }

private:
    ~TempFileReader() = default;

    std::atomic_ulong _refCount{1};
    wil::unique_hfile _file;
    uint64_t _sizeBytes = 0;
};

class TempFileWriter final : public IFileWriter, public IFileWriterExpectedReplacement
{
public:
    TempFileWriter(wil::com_ptr<FileSystemCurl> owner,
                   wil::unique_hfile file,
                   FileSystemCurlProtocol protocol,
                   FileSystemCurl::Settings settings,
                   wil::com_ptr<IHostConnections> hostConnections,
                   std::wstring pluginPath,
                   FileSystemFlags flags) noexcept
        : _owner(std::move(owner)),
          _file(std::move(file)),
          _protocol(protocol),
          _settings(std::move(settings)),
          _hostConnections(std::move(hostConnections)),
          _pluginPath(std::move(pluginPath)),
          _flags(flags)
    {
    }

    TempFileWriter(const TempFileWriter&)            = delete;
    TempFileWriter(TempFileWriter&&)                 = delete;
    TempFileWriter& operator=(const TempFileWriter&) = delete;
    TempFileWriter& operator=(TempFileWriter&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileWriter))
        {
            *ppvObject = static_cast<IFileWriter*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileWriterExpectedReplacement))
        {
            *ppvObject = static_cast<IFileWriterExpectedReplacement*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetPosition(uint64_t* positionBytes) noexcept override
    {
        if (! positionBytes)
        {
            return E_POINTER;
        }

        *positionBytes = 0;

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        LARGE_INTEGER zero{};
        LARGE_INTEGER moved{};
        if (SetFilePointerEx(_file.get(), zero, &moved, FILE_CURRENT) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (moved.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        *positionBytes = static_cast<uint64_t>(moved.QuadPart);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept override
    {
        if (! bytesWritten)
        {
            return E_POINTER;
        }

        *bytesWritten = 0;

        if (bytesToWrite == 0)
        {
            return S_OK;
        }

        if (! buffer)
        {
            return E_POINTER;
        }

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        DWORD written = 0;
        if (WriteFile(_file.get(), buffer, bytesToWrite, &written, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        *bytesWritten = static_cast<unsigned long>(written);
        return S_OK;
    }

    // R3-1: the occupant the host showed the user; Commit promotes only over that occupant.
    HRESULT STDMETHODCALLTYPE SetExpectedReplacement(const FileSystemBasicInformation* expected) noexcept override
    {
        if (expected == nullptr || expected->sizeBytes < sizeof(FileSystemBasicInformation))
        {
            return E_INVALIDARG;
        }
        if (! HasFlag(_flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE) || _committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        _replaceExpectation = CurlReplaceExpectation{.lastWriteTime = expected->lastWriteTime};
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override
    {
        if (_committed)
        {
            return S_OK;
        }

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        uint64_t sizeBytes = 0;
        HRESULT hr         = GetFileSizeBytes(_file.get(), sizeBytes);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ResetFilePointerToStart(_file.get());
        if (FAILED(hr))
        {
            return hr;
        }

        const bool allowOverwrite                        = HasFlag(_flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        const CurlReplaceExpectation* replaceExpectation = _replaceExpectation.has_value() ? &_replaceExpectation.value() : nullptr;
        CurlWriterPublicationMetrics publicationMetrics{};
        CurlPublicationAccumulator publicationAccumulator;

        hr = ResolveLocationWithAuthRetry(_protocol,
                                          _settings,
                                          _pluginPath.c_str(),
                                          _hostConnections.get(),
                                          true,
                                          [&](const ResolvedLocation& resolved) noexcept
        {
            const CurlPublicationResult attemptResult = PublishCurlWriterTransaction(
                resolved.connection, resolved.remotePath, _file.get(), sizeBytes, allowOverwrite, replaceExpectation, publicationMetrics);
            publicationAccumulator.Merge(attemptResult);
            return attemptResult.OperationResult();
        });

        constexpr std::wstring_view detail = L"writer commit";
        Debug::Perf::Emit(L"FileOps.Curl.Writer.CommitUs", detail, publicationMetrics.commitUs, publicationMetrics.requestCount, sizeBytes, hr);
        Debug::Perf::Emit(L"FileOps.Curl.Writer.StagedBytes", detail, 0u, publicationMetrics.stagedBytes, sizeBytes, hr);
        Debug::Perf::Emit(L"FileOps.Curl.Writer.RequestCount", detail, 0u, publicationMetrics.requestCount, sizeBytes, hr);
        Debug::Perf::Emit(L"FileOps.Curl.Writer.CleanupAttemptCount", detail, 0u, publicationMetrics.cleanupAttemptCount, sizeBytes, hr);
        Debug::Perf::Emit(L"FileOps.Curl.Writer.CancellationCount", detail, 0u, hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) ? 1u : 0u, sizeBytes, hr);
        const CurlPublicationResult publicationResult = publicationAccumulator.Snapshot();
        if (FAILED(hr))
        {
            if (_owner)
            {
                _owner->ObserveCurlCleanupDebt(publicationResult);
            }
            return hr;
        }

        _committed = true;
        if (_owner)
        {
            _owner->NotifySyntheticPathCreated(_pluginPath);
            _owner->ObserveCurlCleanupDebt(publicationResult);
        }
        return S_OK;
    }

private:
    ~TempFileWriter() = default;

    wil::com_ptr<FileSystemCurl> _owner;
    std::atomic_ulong _refCount{1};
    wil::unique_hfile _file;
    FileSystemCurlProtocol _protocol = FileSystemCurlProtocol::Ftp;
    FileSystemCurl::Settings _settings{};
    wil::com_ptr<IHostConnections> _hostConnections;
    std::wstring _pluginPath;
    FileSystemFlags _flags = FILESYSTEM_FLAG_NONE;
    bool _committed        = false;
    std::optional<CurlReplaceExpectation> _replaceExpectation;
};

class CurlStreamingReader final : public IFileReader, public IFileReaderOperationControl
{
public:
    CurlStreamingReader(ConnectionInfo conn, std::wstring remotePath, uint64_t sizeBytes, bool sizeKnown) noexcept
        : _conn(std::move(conn)),
          _remotePath(std::move(remotePath)),
          _sizeBytes(sizeBytes),
          _sizeKnown(sizeKnown)
    {
    }

    CurlStreamingReader(const CurlStreamingReader&)            = delete;
    CurlStreamingReader(CurlStreamingReader&&)                 = delete;
    CurlStreamingReader& operator=(const CurlStreamingReader&) = delete;
    CurlStreamingReader& operator=(CurlStreamingReader&&)      = delete;

    HRESULT Initialize() noexcept
    {
        constexpr size_t kBufferBytes = 1024u * 1024u;
        _buffer.reset(new (std::nothrow) std::byte[kBufferBytes]);
        if (! _buffer)
        {
            return E_OUTOFMEMORY;
        }
        _bufferCapacity = kBufferBytes;

        // Pin the module so the DLL cannot be unloaded while the worker thread is running.
        _modulePin = AcquireModuleReferenceFromAddress(&kFileSystemCurlModuleAnchor);
        if (! _modulePin)
        {
            Debug::Error(L"FileSystemCurl: Failed to pin module for CurlStreamingReader");
            return E_FAIL;
        }

        try
        {
            _worker = std::jthread([this](std::stop_token stopToken) noexcept { WorkerMain(stopToken); });
        }
        catch (const std::system_error&)
        {
            // Module pin released via RAII if thread creation fails.
            return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileReaderOperationControl))
        {
            *ppvObject = static_cast<IFileReaderOperationControl*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    // R0f-Curl-OR1: the host's operation control for a reader created without options. The worker's
    // progress callback polls it on every libcurl call (at least once per second, also while a
    // command waits for the server), so a cancel ends the transfer and the in-flight Read.
    HRESULT STDMETHODCALLTYPE SetOperationControl(const FileSystemOptions* options) noexcept override
    {
        _operationOptions.store(options, std::memory_order_release);
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (! sizeBytes)
        {
            return E_POINTER;
        }

        *sizeBytes = _sizeBytes;
        return _sizeKnown ? S_OK : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (! newPosition)
        {
            return E_POINTER;
        }

        *newPosition = 0;

        if (origin != FILE_BEGIN && origin != FILE_CURRENT && origin != FILE_END)
        {
            return E_INVALIDARG;
        }

        std::unique_lock lock(_mutex);

        uint64_t base = 0;
        if (origin == FILE_BEGIN)
        {
            base = 0;
        }
        else if (origin == FILE_CURRENT)
        {
            base = _positionBytes;
        }
        else
        {
            if (! _sizeKnown)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
            base = _sizeBytes;
        }

        if (offset == (std::numeric_limits<__int64>::min)())
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        if (offset < 0)
        {
            const uint64_t magnitude = static_cast<uint64_t>(-(offset + 1)) + 1u;
            if (base < magnitude)
            {
                return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
            }
        }
        else
        {
            const uint64_t add = static_cast<uint64_t>(offset);
            if (base > (std::numeric_limits<uint64_t>::max)() - add)
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
        }

        const uint64_t newPos = (offset < 0) ? (base - (static_cast<uint64_t>(-(offset + 1)) + 1u)) : (base + static_cast<uint64_t>(offset));

        if (newPos == _positionBytes && SUCCEEDED(_workerHr))
        {
            *newPosition = newPos;
            return S_OK;
        }

        _positionBytes = newPos;
        _readPos       = 0;
        _writePos      = 0;
        _bufferedBytes = 0;
        _eof           = false;
        _workerHr      = S_OK;

        _generation.fetch_add(1, std::memory_order_acq_rel);

        lock.unlock();
        _cvReadable.notify_all();
        _cvWritable.notify_all();

        *newPosition = newPos;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (! bytesRead)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        if (bytesToRead == 0)
        {
            return S_OK;
        }

        if (! buffer)
        {
            return E_POINTER;
        }

        std::unique_lock lock(_mutex);
        if (_stopping.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        uint64_t remainingCommittedBytes = (std::numeric_limits<uint64_t>::max)();
        if (_sizeKnown)
        {
            if (_positionBytes > _sizeBytes)
            {
                _workerHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                return _workerHr;
            }

            remainingCommittedBytes = _sizeBytes - _positionBytes;
            if (remainingCommittedBytes == 0u)
            {
                while (! _eof && SUCCEEDED(_workerHr) && ! _stopping.load(std::memory_order_acquire))
                {
#if defined(ENABLE_TESTS)
                    _debugReadableWaitEntered.store(true, std::memory_order_release);
#endif
                    _cvReadable.wait(lock, [&]() noexcept { return _eof || FAILED(_workerHr) || _stopping.load(std::memory_order_acquire); });
                }
                if (_stopping.load(std::memory_order_acquire) && SUCCEEDED(_workerHr))
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                return FAILED(_workerHr) ? _workerHr : S_OK;
            }
        }

        const size_t desiredBufferedBytes =
            (std::min)({static_cast<size_t>(bytesToRead),
                        _bufferCapacity,
                        static_cast<size_t>((std::min)(remainingCommittedBytes, static_cast<uint64_t>((std::numeric_limits<size_t>::max)())))});
        const bool readReachesCommittedEnd = _sizeKnown && ReadRequestReachesCommittedEnd(_bufferCapacity, bytesToRead, remainingCommittedBytes);
        while (_bufferedBytes < desiredBufferedBytes || (readReachesCommittedEnd && ! _eof))
        {
            if (_stopping.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            if (FAILED(_workerHr))
            {
                return _workerHr;
            }
            if (_eof)
            {
                break;
            }
#if defined(ENABLE_TESTS)
            _debugReadableWaitEntered.store(true, std::memory_order_release);
#endif
            _cvReadable.wait(lock,
                             [&]() noexcept
            {
                return _stopping.load(std::memory_order_acquire) || FAILED(_workerHr) || _eof ||
                       (_bufferedBytes >= desiredBufferedBytes && (! readReachesCommittedEnd || _eof));
            });
        }

        if (_sizeKnown && _bufferedBytes > remainingCommittedBytes)
        {
            _workerHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return _workerHr;
        }
        if (_eof && _bufferedBytes == 0u && _sizeKnown && _positionBytes < _sizeBytes)
        {
            _workerHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            return _workerHr;
        }

        const size_t take = (std::min)(static_cast<size_t>(bytesToRead), _bufferedBytes);
        if (_positionBytes > (std::numeric_limits<uint64_t>::max)() - static_cast<uint64_t>(take))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        const size_t first = (std::min)(take, _bufferCapacity - _readPos);

        std::memcpy(buffer, _buffer.get() + _readPos, first);
        _readPos = (_readPos + first) % _bufferCapacity;
        _bufferedBytes -= first;

        const size_t remaining = take - first;
        if (remaining > 0)
        {
            std::memcpy(static_cast<std::byte*>(buffer) + first, _buffer.get() + _readPos, remaining);
            _readPos = (_readPos + remaining) % _bufferCapacity;
            _bufferedBytes -= remaining;
        }

        _positionBytes += static_cast<uint64_t>(take);

        lock.unlock();
        _cvWritable.notify_all();

        *bytesRead = static_cast<unsigned long>(take);
        return S_OK;
    }

#if defined(ENABLE_TESTS)
    [[nodiscard]] HRESULT DebugPrimeForSelfTest(std::span<const std::byte> bytes, bool eof) noexcept
    {
        if (bytes.empty())
        {
            return E_INVALIDARG;
        }

        _buffer.reset(new (std::nothrow) std::byte[bytes.size()]);
        if (! _buffer)
        {
            return E_OUTOFMEMORY;
        }

        std::memcpy(_buffer.get(), bytes.data(), bytes.size());
        _bufferCapacity = bytes.size();
        _readPos        = 0;
        _writePos       = 0;
        _bufferedBytes  = bytes.size();
        _eof            = eof;
        return S_OK;
    }

    [[nodiscard]] HRESULT DebugPrepareEmptyForSelfTest(const size_t capacity) noexcept
    {
        if (capacity == 0u)
        {
            return E_INVALIDARG;
        }
        std::unique_ptr<std::byte[]> buffer(new (std::nothrow) std::byte[capacity]);
        if (! buffer)
        {
            return E_OUTOFMEMORY;
        }
        std::scoped_lock lock(_mutex);
        _buffer         = std::move(buffer);
        _bufferCapacity = capacity;
        _readPos        = 0u;
        _writePos       = 0u;
        _bufferedBytes  = 0u;
        _eof            = false;
        _workerHr       = S_OK;
        _stopping.store(false, std::memory_order_release);
        _debugReadableWaitEntered.store(false, std::memory_order_release);
        _debugWritableWaitEntered.store(false, std::memory_order_release);
        return S_OK;
    }

    void DebugRequestStopForSelfTest() noexcept
    {
        {
            std::scoped_lock lock(_mutex);
            _stopping.store(true, std::memory_order_release);
        }
        _cvReadable.notify_all();
        _cvWritable.notify_all();
    }

    void DebugFailForSelfTest(const HRESULT failure) noexcept
    {
        {
            std::scoped_lock lock(_mutex);
            _workerHr = FAILED(failure) ? failure : E_FAIL;
        }
        _cvReadable.notify_all();
        _cvWritable.notify_all();
    }

    [[nodiscard]] bool DebugReadableWaitEnteredForSelfTest() const noexcept
    {
        return _debugReadableWaitEntered.load(std::memory_order_acquire);
    }

    [[nodiscard]] bool DebugWritableWaitEnteredForSelfTest() const noexcept
    {
        return _debugWritableWaitEntered.load(std::memory_order_acquire);
    }

    [[nodiscard]] size_t DebugWriteForSelfTest(const std::span<const std::byte> bytes) noexcept
    {
        return OnCurlWrite(bytes.data(), bytes.size(), std::stop_token{});
    }
#endif

private:
    ~CurlStreamingReader()
    {
        {
            // Publish the predicate change under the same mutex used by the
            // condition-variable waiters. Otherwise teardown can notify after
            // a waiter observes false but before it atomically starts waiting.
            std::scoped_lock lock(_mutex);
            _stopping.store(true, std::memory_order_release);
        }
        if (_worker.joinable())
        {
            _worker.request_stop();
        }
        _cvReadable.notify_all();
        _cvWritable.notify_all();
    }

    [[nodiscard]] size_t OnCurlWrite(const std::byte* data, size_t bytes, std::stop_token stopToken) noexcept
    {
        if (! data || bytes == 0)
        {
            return 0;
        }

        const uint64_t activeGen = _transferGeneration.load(std::memory_order_acquire);

        if (_activeTransferSizeKnown && (_activeTransferReceivedBytes > _activeTransferExpectedBytes ||
                                         static_cast<uint64_t>(bytes) > _activeTransferExpectedBytes - _activeTransferReceivedBytes))
        {
            _activeTransferValidationHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return 0;
        }

        size_t offset = 0;
        while (offset < bytes)
        {
            if (_stopping.load(std::memory_order_acquire) || stopToken.stop_requested())
            {
                return 0;
            }

            if (_generation.load(std::memory_order_acquire) != activeGen)
            {
                return 0;
            }

            std::unique_lock lock(_mutex);
#if defined(ENABLE_TESTS)
            _debugWritableWaitEntered.store(true, std::memory_order_release);
#endif
            _cvWritable.wait(lock,
                             [&]() noexcept
            {
                return _stopping.load(std::memory_order_acquire) || stopToken.stop_requested() || _generation.load(std::memory_order_acquire) != activeGen ||
                       _bufferedBytes < _bufferCapacity;
            });

            if (_stopping.load(std::memory_order_acquire) || stopToken.stop_requested())
            {
                return 0;
            }

            if (_generation.load(std::memory_order_acquire) != activeGen)
            {
                return 0;
            }

            const size_t space = _bufferCapacity - _bufferedBytes;
            const size_t chunk = (std::min)(space, bytes - offset);
            if (chunk == 0)
            {
                continue;
            }

            const size_t first = (std::min)(chunk, _bufferCapacity - _writePos);
            std::memcpy(_buffer.get() + _writePos, data + offset, first);
            _writePos = (_writePos + first) % _bufferCapacity;
            _bufferedBytes += first;
            offset += first;

            const size_t second = chunk - first;
            if (second > 0)
            {
                std::memcpy(_buffer.get() + _writePos, data + offset, second);
                _writePos = (_writePos + second) % _bufferCapacity;
                _bufferedBytes += second;
                offset += second;
            }

            lock.unlock();
            _cvReadable.notify_all();
        }

        _activeTransferReceivedBytes += static_cast<uint64_t>(bytes);
        return bytes;
    }

    static size_t CurlWriteToStream(void* ptr, size_t size, size_t nmemb, void* userdata) noexcept
    {
        auto* self = static_cast<CurlStreamingReader*>(userdata);
        if (! self || ! ptr || size == 0 || nmemb == 0)
        {
            return 0;
        }

        if (nmemb > (std::numeric_limits<size_t>::max)() / size)
        {
            self->_activeTransferValidationHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            return 0;
        }

        const size_t total = size * nmemb;
        return self->OnCurlWrite(static_cast<const std::byte*>(ptr), total, self->_worker.get_stop_token());
    }

    static int CurlProgress(void* clientp, curl_off_t /*dltotal*/, curl_off_t /*dlnow*/, curl_off_t /*ultotal*/, curl_off_t /*ulnow*/) noexcept
    {
        auto* self = static_cast<CurlStreamingReader*>(clientp);
        if (! self)
        {
            return 0;
        }

        if (self->_stopping.load(std::memory_order_acquire))
        {
            return 1;
        }
        if (const FileSystemOptions* const options = self->_operationOptions.load(std::memory_order_acquire);
            options != nullptr && FAILED(FileSystemCheckOperationControl(options)))
        {
            // R0f-Curl-OR1: the owning call was cancelled or ran out of time; stop like a host stop.
            self->_stopping.store(true, std::memory_order_release);
            return 1;
        }

        const uint64_t activeGen = self->_transferGeneration.load(std::memory_order_acquire);
        return self->_generation.load(std::memory_order_acquire) != activeGen ? 1 : 0;
    }

    void WorkerMain(std::stop_token stopToken) noexcept
    {
        const HRESULT initHr = EnsureCurlInitialized();
        if (FAILED(initHr))
        {
            std::scoped_lock lock(_mutex);
            _workerHr = initHr;
            _cvReadable.notify_all();
            return;
        }

        for (;;)
        {
            if (_stopping.load(std::memory_order_acquire) || stopToken.stop_requested())
            {
                std::scoped_lock lock(_mutex);
                _workerHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                _cvReadable.notify_all();
                return;
            }

            const uint64_t gen = _generation.load(std::memory_order_acquire);
            _transferGeneration.store(gen, std::memory_order_release);

            uint64_t startOffset = 0;
            {
                std::scoped_lock lock(_mutex);
                startOffset = _positionBytes;
            }

            _activeTransferSizeKnown     = _sizeKnown;
            _activeTransferExpectedBytes = (_sizeKnown && startOffset < _sizeBytes) ? (_sizeBytes - startOffset) : 0u;
            _activeTransferReceivedBytes = 0u;
            _activeTransferValidationHr  = S_OK;

            if (_sizeKnown && startOffset > _sizeBytes)
            {
                {
                    std::scoped_lock lock(_mutex);
                    if (_generation.load(std::memory_order_acquire) != gen)
                    {
                        continue;
                    }
                    _workerHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    _eof      = false;
                }
                _cvReadable.notify_all();

                std::unique_lock lock(_mutex);
                _cvWritable.wait(lock, [&]() noexcept {
                    return _stopping.load(std::memory_order_acquire) || stopToken.stop_requested() || _generation.load(std::memory_order_acquire) != gen;
                });
                continue;
            }

            // A zero-byte committed range still performs the transfer. This is
            // required for an initially empty file and for an explicit seek to
            // the committed end: a contradictory body must be rejected rather
            // than silently reported as successful EOF.

            const std::string url = BuildUrl(_conn, _remotePath, false, false);
            if (url.empty())
            {
                std::scoped_lock lock(_mutex);
                _workerHr = E_INVALIDARG;
                _cvReadable.notify_all();
                return;
            }

            auto curl = GetCurlEasyPool().Borrow(_conn.limiterKey);
            if (! curl)
            {
                std::scoped_lock lock(_mutex);
                _workerHr = E_OUTOFMEMORY;
                _cvReadable.notify_all();
                return;
            }

            curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
            curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, CurlWriteToStream);
            curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, this);
            curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

            curl_easy_setopt(curl.get(), CURLOPT_XFERINFOFUNCTION, CurlProgress);
            curl_easy_setopt(curl.get(), CURLOPT_XFERINFODATA, this);
            curl_easy_setopt(curl.get(), CURLOPT_NOPROGRESS, 0L);

            ApplyCommonCurlOptions(curl.get(), _conn, nullptr, false);

            if (_activeTransferSizeKnown)
            {
                // The reader's probed size is the commitment. Observe the real
                // data-stream terminator so a contradictory short or overlong
                // body cannot be hidden by libcurl's advertised-length handling.
                curl_easy_setopt(curl.get(), CURLOPT_IGNORE_CONTENT_LENGTH, 1L);
            }

            unique_curl_slist preTransferCommands;
            if (startOffset > 0)
            {
                if (_conn.protocol == Protocol::Ftp)
                {
                    constexpr std::string_view kRestPrefix = "REST ";
                    std::array<char, 32u> command{};
                    std::memcpy(command.data(), kRestPrefix.data(), kRestPrefix.size());
                    const auto [commandEnd, commandError] =
                        std::to_chars(command.data() + kRestPrefix.size(), command.data() + command.size() - 1u, startOffset);
                    if (commandError != std::errc{})
                    {
                        std::scoped_lock lock(_mutex);
                        _workerHr = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                        _cvReadable.notify_all();
                        return;
                    }
                    *commandEnd = '\0';

                    preTransferCommands.reset(curl_slist_append(nullptr, command.data()));
                    if (! preTransferCommands)
                    {
                        std::scoped_lock lock(_mutex);
                        _workerHr = E_OUTOFMEMORY;
                        _cvReadable.notify_all();
                        return;
                    }
                    curl_easy_setopt(curl.get(), CURLOPT_PREQUOTE, preTransferCommands.get());
                }
                else
                {
                    constexpr uint64_t kCurlOffMax = static_cast<uint64_t>((std::numeric_limits<curl_off_t>::max)());
                    if (startOffset > kCurlOffMax)
                    {
                        std::scoped_lock lock(_mutex);
                        _workerHr = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                        _cvReadable.notify_all();
                        return;
                    }
                    curl_easy_setopt(curl.get(), CURLOPT_RESUME_FROM_LARGE, static_cast<curl_off_t>(startOffset));
                }
            }

            const CURLcode code = curl_easy_perform(curl.get());

            if (_stopping.load(std::memory_order_acquire) || stopToken.stop_requested())
            {
                std::scoped_lock lock(_mutex);
                _workerHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                _cvReadable.notify_all();
                return;
            }

            if (_generation.load(std::memory_order_acquire) != gen)
            {
                continue;
            }

            HRESULT transferHr = _activeTransferValidationHr;
            if (SUCCEEDED(transferHr) && _activeTransferSizeKnown && _activeTransferReceivedBytes > _activeTransferExpectedBytes)
            {
                transferHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            if (SUCCEEDED(transferHr) && _activeTransferSizeKnown && _activeTransferReceivedBytes < _activeTransferExpectedBytes &&
                (code == CURLE_OK || code == CURLE_PARTIAL_FILE))
            {
                transferHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }
            if (SUCCEEDED(transferHr) && code != CURLE_OK)
            {
                transferHr = HResultFromCurl(code);
            }

            {
                std::scoped_lock lock(_mutex);
                _workerHr = transferHr;
                _eof      = SUCCEEDED(transferHr);
            }
            _cvReadable.notify_all();

            std::unique_lock lock(_mutex);
            _cvWritable.wait(lock, [&]() noexcept {
                return _stopping.load(std::memory_order_acquire) || stopToken.stop_requested() || _generation.load(std::memory_order_acquire) != gen;
            });
        }
    }

    std::atomic_ulong _refCount{1};

    ConnectionInfo _conn;
    std::wstring _remotePath;

    uint64_t _sizeBytes = 0;
    bool _sizeKnown     = false;

    std::mutex _mutex;
    std::condition_variable _cvReadable;
    std::condition_variable _cvWritable;

    std::unique_ptr<std::byte[]> _buffer;
    size_t _bufferCapacity = 0;

    size_t _readPos       = 0;
    size_t _writePos      = 0;
    size_t _bufferedBytes = 0;

    uint64_t _positionBytes = 0;

    std::atomic<uint64_t> _generation{0};
    std::atomic<uint64_t> _transferGeneration{0};
    uint64_t _activeTransferExpectedBytes = 0u;
    uint64_t _activeTransferReceivedBytes = 0u;
    bool _activeTransferSizeKnown         = false;
    HRESULT _activeTransferValidationHr   = S_OK;
    std::atomic_bool _stopping{false};
    std::atomic<const FileSystemOptions*> _operationOptions{nullptr}; // R0f-Curl-OR1

    bool _eof         = false;
    HRESULT _workerHr = S_OK;

#if defined(ENABLE_TESTS)
    std::atomic_bool _debugReadableWaitEntered{false};
    std::atomic_bool _debugWritableWaitEntered{false};
#endif

    wil::unique_hmodule _modulePin;
    std::jthread _worker;
};

} // namespace

#if defined(ENABLE_TESTS)
namespace FileSystemCurlInternal
{
void RunDebugCurlStreamingReaderContractSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept
    {
        if (condition)
        {
            ++passed;
            return;
        }

        ++failed;
        Debug::Error(L"FileSystemCurl streaming-reader selftest failed: {}", message);
    };

    constexpr size_t kReaderBufferBytes = 1024u * 1024u;
    check(ReadRequestReachesCommittedEnd(kReaderBufferBytes, 4099u, 4099u),
          L"a request that can consume the buffered committed tail should await terminal validation");
    check(! ReadRequestReachesCommittedEnd(kReaderBufferBytes, (std::numeric_limits<unsigned long>::max)(), static_cast<uint64_t>(kReaderBufferBytes) + 1u),
          L"a request larger than the ring buffer must drain a chunk instead of deadlocking while awaiting terminal validation");

    ConnectionInfo connection{};
    wil::com_ptr<CurlStreamingReader> reader;
    reader.attach(new (std::nothrow) CurlStreamingReader(std::move(connection), L"/selftest", 0u, false));
    check(static_cast<bool>(reader), L"reader allocation should succeed");
    if (! reader)
    {
        return;
    }

    uint64_t sizeBytes = 123u;
    check(reader->GetSize(&sizeBytes) == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), L"unknown size should not be reported as zero");

    uint64_t position = 123u;
    check(reader->Seek(0, FILE_END, &position) == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), L"FILE_END seek should reject unknown size");

    constexpr size_t kHintBytes = 1024u * 1024u;
    std::vector<std::byte> fullRequest(kHintBytes, std::byte{0x5a});
    check(SUCCEEDED(reader->DebugPrimeForSelfTest(fullRequest, false)), L"full-request test buffer should initialize");

    std::vector<std::byte> output(kHintBytes);
    unsigned long bytesRead = 0;
    check(SUCCEEDED(reader->Read(output.data(), static_cast<unsigned long>(output.size()), &bytesRead)) && bytesRead == output.size(),
          L"reader should fill the advertised 1 MiB request when data is available");

    constexpr std::array<std::byte, 7> tail{{std::byte{1}, std::byte{2}, std::byte{3}, std::byte{4}, std::byte{5}, std::byte{6}, std::byte{7}}};
    check(SUCCEEDED(reader->DebugPrimeForSelfTest(tail, true)), L"EOF-tail test buffer should initialize");

    std::array<std::byte, 16> tailOutput{};
    bytesRead = 0;
    check(SUCCEEDED(reader->Read(tailOutput.data(), static_cast<unsigned long>(tailOutput.size()), &bytesRead)) && bytesRead == tail.size() &&
              std::equal(tail.begin(), tail.end(), tailOutput.begin()),
          L"reader should return the final buffered tail before EOF");

    const auto makeKnownReader = [](uint64_t sizeBytes) noexcept
    {
        wil::com_ptr<CurlStreamingReader> result;
        ConnectionInfo knownConnection{};
        result.attach(new (std::nothrow) CurlStreamingReader(std::move(knownConnection), L"/known-selftest", sizeBytes, true));
        return result;
    };

    const auto waitUntil = [](const auto& predicate, const std::chrono::milliseconds timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (! predicate() && std::chrono::steady_clock::now() < deadline)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds{1});
        }
        return predicate();
    };

    {
        auto waitingReader = makeKnownReader(1u);
        check(static_cast<bool>(waitingReader) && SUCCEEDED(waitingReader->DebugPrepareEmptyForSelfTest(16u)), L"read-wait teardown fixture should initialize");
        if (waitingReader)
        {
            std::atomic_bool returned{false};
            std::atomic<HRESULT> readHr{E_PENDING};
            std::jthread readThread;
            try
            {
                readThread = std::jthread([&]() noexcept
                {
                    std::array<std::byte, 1u> byte{};
                    unsigned long read = 0u;
                    readHr.store(waitingReader->Read(byte.data(), 1u, &read), std::memory_order_release);
                    returned.store(true, std::memory_order_release);
                });
            }
            catch (const std::bad_alloc&)
            {
                std::terminate();
            }
            catch (const std::system_error& error)
            {
                Debug::Error(L"FileSystemCurl streaming-reader selftest could not create the read-wait worker (code={}).", error.code().value());
            }
            const bool entered = waitUntil([&]() noexcept { return waitingReader->DebugReadableWaitEnteredForSelfTest(); }, std::chrono::seconds{2});
            waitingReader->DebugRequestStopForSelfTest();
            const bool stopped = waitUntil([&]() noexcept { return returned.load(std::memory_order_acquire); }, std::chrono::seconds{2});
            if (! stopped)
            {
                waitingReader->DebugFailForSelfTest(E_ABORT);
            }
            if (readThread.joinable())
            {
                readThread.join();
            }
            check(entered && stopped && readHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  L"reader stopping should wake a consumer blocked in Read and return Cancelled");
        }
    }

    {
        auto fullReader = makeKnownReader(2u);
        constexpr std::array<std::byte, 1u> full{{std::byte{0x2a}}};
        check(static_cast<bool>(fullReader) && SUCCEEDED(fullReader->DebugPrimeForSelfTest(full, false)), L"full-buffer writer-stop fixture should initialize");
        if (fullReader)
        {
            std::atomic_bool returned{false};
            std::atomic_size_t written{1u};
            std::jthread writerThread;
            try
            {
                writerThread = std::jthread([&]() noexcept
                {
                    written.store(fullReader->DebugWriteForSelfTest(full), std::memory_order_release);
                    returned.store(true, std::memory_order_release);
                });
            }
            catch (const std::bad_alloc&)
            {
                std::terminate();
            }
            catch (const std::system_error& error)
            {
                Debug::Error(L"FileSystemCurl streaming-reader selftest could not create the writer-wait worker (code={}).", error.code().value());
            }
            const bool entered = waitUntil([&]() noexcept { return fullReader->DebugWritableWaitEnteredForSelfTest(); }, std::chrono::seconds{2});
            fullReader->DebugRequestStopForSelfTest();
            const bool stopped = waitUntil([&]() noexcept { return returned.load(std::memory_order_acquire); }, std::chrono::seconds{2});
            if (! stopped)
            {
                fullReader->DebugFailForSelfTest(E_ABORT);
            }
            if (writerThread.joinable())
            {
                writerThread.join();
            }
            check(entered && stopped && written.load(std::memory_order_acquire) == 0u,
                  L"reader stopping should wake a Curl writer blocked on a full ring buffer");
        }
    }

    {
        auto knownReader = makeKnownReader(tail.size());
        check(static_cast<bool>(knownReader) && SUCCEEDED(knownReader->DebugPrimeForSelfTest(tail, true)), L"known-size exact reader should initialize");
        std::array<std::byte, 16> exactOutput{};
        bytesRead = 0;
        check(knownReader && SUCCEEDED(knownReader->Read(exactOutput.data(), static_cast<unsigned long>(exactOutput.size()), &bytesRead)) &&
                  bytesRead == tail.size() && std::equal(tail.begin(), tail.end(), exactOutput.begin()),
              L"known-size exact reader should return every committed byte");
        bytesRead = 99u;
        check(knownReader && SUCCEEDED(knownReader->Read(exactOutput.data(), static_cast<unsigned long>(exactOutput.size()), &bytesRead)) && bytesRead == 0u,
              L"known-size exact reader should report EOF only after the commitment is fulfilled");
    }

    {
        auto shortReader = makeKnownReader(tail.size() + 1u);
        check(static_cast<bool>(shortReader) && SUCCEEDED(shortReader->DebugPrimeForSelfTest(tail, true)), L"known-size short reader should initialize");
        std::array<std::byte, 16> shortOutput{};
        bytesRead = 0;
        check(shortReader && SUCCEEDED(shortReader->Read(shortOutput.data(), static_cast<unsigned long>(shortOutput.size()), &bytesRead)) &&
                  bytesRead == tail.size(),
              L"known-size short reader may return its final nonempty prefix");
        bytesRead = 99u;
        check(shortReader &&
                  shortReader->Read(shortOutput.data(), static_cast<unsigned long>(shortOutput.size()), &bytesRead) == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                  bytesRead == 0u,
              L"known-size short reader must fail instead of reporting premature successful EOF");
    }

    {
        auto overlongReader = makeKnownReader(tail.size() - 1u);
        check(static_cast<bool>(overlongReader) && SUCCEEDED(overlongReader->DebugPrimeForSelfTest(tail, true)),
              L"known-size overlong reader should initialize");
        std::array<std::byte, 16> overlongOutput{};
        bytesRead = 99u;
        check(overlongReader &&
                  overlongReader->Read(overlongOutput.data(), static_cast<unsigned long>(overlongOutput.size()), &bytesRead) ==
                      HRESULT_FROM_WIN32(ERROR_INVALID_DATA) &&
                  bytesRead == 0u,
              L"known-size reader must reject buffered bytes beyond its commitment");
    }

    {
        auto seekReader       = makeKnownReader(tail.size());
        uint64_t seekPosition = 0u;
        check(static_cast<bool>(seekReader) && SUCCEEDED(seekReader->Seek(3, FILE_BEGIN, &seekPosition)) && seekPosition == 3u,
              L"known-size seeked-range reader should accept an in-range restart");
        const std::span<const std::byte> seekTail{tail.data() + 3u, tail.size() - 3u};
        check(seekReader && SUCCEEDED(seekReader->DebugPrimeForSelfTest(seekTail, true)), L"known-size seeked range should initialize");
        std::array<std::byte, 16> seekOutput{};
        bytesRead = 0;
        check(seekReader && SUCCEEDED(seekReader->Read(seekOutput.data(), static_cast<unsigned long>(seekOutput.size()), &bytesRead)) &&
                  bytesRead == seekTail.size() && std::equal(seekTail.begin(), seekTail.end(), seekOutput.begin()),
              L"known-size seeked range should validate only its committed remainder");
    }

    {
        constexpr std::array<std::byte, 8> restartBytes{
            {std::byte{8}, std::byte{7}, std::byte{6}, std::byte{5}, std::byte{4}, std::byte{3}, std::byte{2}, std::byte{1}}};
        auto restartReader = makeKnownReader(restartBytes.size());
        const std::span<const std::byte> shortGeneration{restartBytes.data(), restartBytes.size() - 1u};
        check(static_cast<bool>(restartReader) && SUCCEEDED(restartReader->DebugPrimeForSelfTest(shortGeneration, true)),
              L"generation-restart short range should initialize");
        std::array<std::byte, 16> restartOutput{};
        bytesRead = 0;
        check(restartReader && SUCCEEDED(restartReader->Read(restartOutput.data(), static_cast<unsigned long>(restartOutput.size()), &bytesRead)) &&
                  bytesRead == shortGeneration.size(),
              L"generation-restart fixture should consume the short generation prefix");
        bytesRead = 0;
        check(restartReader && restartReader->Read(restartOutput.data(), static_cast<unsigned long>(restartOutput.size()), &bytesRead) ==
                                   HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
              L"generation-restart fixture should latch its first premature completion");

        uint64_t restartPosition = 99u;
        check(restartReader && SUCCEEDED(restartReader->Seek(0, FILE_BEGIN, &restartPosition)) && restartPosition == 0u &&
                  SUCCEEDED(restartReader->DebugPrimeForSelfTest(restartBytes, true)),
              L"seek should reset a failed generation before retry");
        bytesRead = 0;
        check(restartReader && SUCCEEDED(restartReader->Read(restartOutput.data(), static_cast<unsigned long>(restartOutput.size()), &bytesRead)) &&
                  bytesRead == restartBytes.size() && std::equal(restartBytes.begin(), restartBytes.end(), restartOutput.begin()),
              L"a restarted generation should expose only its own exact committed bytes");
    }
}

void RunDebugCurlWriterOwnershipContractSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept
    {
        if (condition)
        {
            ++passed;
            return;
        }

        ++failed;
        Debug::Error(L"FileSystemCurl writer-ownership selftest failed: {}", message);
    };

    wil::com_ptr<FileSystemCurl> owner;
    owner.attach(new (std::nothrow) FileSystemCurl(FileSystemCurlProtocol::Imap, nullptr));
    check(static_cast<bool>(owner), L"temporary writer owner allocation should succeed");
    wil::unique_hfile file;
    const HRESULT tempHr = Common::Files::CreateDeleteOnCloseTemporaryFile(kCurlTemporaryFileOptions, file);
    check(SUCCEEDED(tempHr) && static_cast<bool>(file), L"temporary writer file should be created");
    if (! owner || ! file)
    {
        return;
    }

    wil::com_ptr<IFileWriter> unsupportedImapWriter;
    check(owner->CreateFileWriter(L"//selftest/unsupported-writer.eml", FILESYSTEM_FLAG_ALLOW_OVERWRITE, unsupportedImapWriter.put()) ==
                  HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) &&
              ! unsupportedImapWriter,
          L"IMAP writer publication should fail closed because the protocol has no recoverable rename primitive");

    wil::com_ptr<IFileWriter> writer;
    writer.attach(new (std::nothrow) TempFileWriter(
        owner, std::move(file), FileSystemCurlProtocol::Imap, FileSystemCurl::Settings{}, nullptr, L"//selftest/owned-writer.eml", FILESYSTEM_FLAG_NONE));
    check(static_cast<bool>(writer), L"temporary writer allocation should succeed");
    owner.reset();
    if (! writer)
    {
        return;
    }

    constexpr std::array<std::byte, 4u> bytes{{std::byte{0x52}, std::byte{0x53}, std::byte{0x21}, std::byte{0x0a}}};
    unsigned long written = 0u;
    check(writer->Write(bytes.data(), static_cast<unsigned long>(bytes.size()), &written) == S_OK && written == bytes.size(),
          L"temporary writer should remain usable after the external filesystem owner is released");
    check(FAILED(writer->Commit()), L"temporary writer commit should fail safely through its retained owner when the synthetic endpoint is unresolved");
    writer.reset();
    check(true, L"temporary writer destruction should release the final filesystem owner without use-after-free");
}
} // namespace FileSystemCurlInternal
#endif

HRESULT STDMETHODCALLTYPE FileSystemCurl::ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept
{
    if (ppFilesInformation == nullptr)
    {
        return E_POINTER;
    }

    *ppFilesInformation = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    std::vector<FilesInformationCurl::Entry> entries;
    const HRESULT hr = ResolveLocationWithAuthRetry(_protocol,
                                                    settings,
                                                    path,
                                                    _hostConnections.get(),
                                                    true,
                                                    [&](const ResolvedLocation& resolved) noexcept
    {
        entries.clear();
        return ReadDirectoryEntries(resolved.connection, resolved.remotePath, entries);
    });
    if (FAILED(hr))
    {
        return hr;
    }

    auto infoImpl = std::unique_ptr<FilesInformationCurl>(new (std::nothrow) FilesInformationCurl());
    if (! infoImpl)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT buildHr = infoImpl->BuildFromEntries(std::move(entries));
    if (FAILED(buildHr))
    {
        return buildHr;
    }

    *ppFilesInformation = infoImpl.release();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept
{
    if (fileAttributes == nullptr)
    {
        return E_POINTER;
    }

    *fileAttributes = 0;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FilesInformationCurl::Entry entry{};
    const HRESULT hr = ResolveLocationWithAuthRetry(_protocol,
                                                    settings,
                                                    path,
                                                    _hostConnections.get(),
                                                    true,
                                                    [&](const ResolvedLocation& resolved) noexcept
    {
        entry = {};
        return GetEntryInfo(resolved.connection, resolved.remotePath, entry);
    });
    if (FAILED(hr))
    {
        return hr;
    }

    *fileAttributes = entry.attributes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept
{
    if (reader == nullptr)
    {
        return E_POINTER;
    }

    *reader = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    return ResolveLocationWithAuthRetry(_protocol,
                                        settings,
                                        path,
                                        _hostConnections.get(),
                                        true,
                                        [&](const ResolvedLocation& resolved) noexcept
    {
        FilesInformationCurl::Entry entry{};
        const HRESULT attrHr = GetEntryInfo(resolved.connection, resolved.remotePath, entry);
        if (FAILED(attrHr))
        {
            return attrHr;
        }

        if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        if (resolved.connection.protocol != Protocol::Imap)
        {
            CurlSourceSizeCommitment sourceSize{};
            const HRESULT sourceSizeHr =
                ResolveCurlSourceSizeCommitment(resolved.connection, resolved.remotePath, entry.sizeBytes, entry.sizeKnown, sourceSize);
            if (FAILED(sourceSizeHr))
            {
                return sourceSizeHr;
            }

            auto* impl = new (std::nothrow) CurlStreamingReader(resolved.connection, resolved.remotePath, sourceSize.sizeBytes, sourceSize.known);
            if (! impl)
            {
                return E_OUTOFMEMORY;
            }

            const HRESULT initHr = impl->Initialize();
            if (FAILED(initHr))
            {
                impl->Release();
                return initHr;
            }

            *reader = impl;
            return S_OK;
        }

        wil::unique_hfile file;
        if (const HRESULT tempHr = Common::Files::CreateDeleteOnCloseTemporaryFile(kCurlTemporaryFileOptions, file); FAILED(tempHr))
        {
            return tempHr;
        }

        HRESULT dlHr = S_OK;
        if (resolved.connection.protocol == Protocol::Imap)
        {
            dlHr = ImapDownloadMessageToFile(resolved.connection, resolved.remotePath, file.get());
        }
        else
        {
            dlHr = CurlDownloadToFile(resolved.connection, resolved.remotePath, file.get(), nullptr, nullptr);
        }
        if (FAILED(dlHr))
        {
            return dlHr;
        }

        uint64_t sizeBytes = 0;
        HRESULT hr         = GetFileSizeBytes(file.get(), sizeBytes);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = ResetFilePointerToStart(file.get());
        if (FAILED(hr))
        {
            return hr;
        }

        auto* impl = new (std::nothrow) TempFileReader(std::move(file), sizeBytes);
        if (! impl)
        {
            return E_OUTOFMEMORY;
        }

        *reader = impl;
        return S_OK;
    });
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::CreateFileWriter([[maybe_unused]] const wchar_t* path,
                                                           [[maybe_unused]] FileSystemFlags flags,
                                                           IFileWriter** writer) noexcept
{
    if (writer == nullptr)
    {
        return E_POINTER;
    }

    *writer = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) && ! HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE))
    {
        return E_INVALIDARG;
    }

    // R0f-Curl: FTP RNTO and libcurl's SFTP rename path have no portable no-replace primitive.
    // A no-overwrite Commit probes the destination and fails with ERROR_FILE_EXISTS when it is
    // present; a creator that appears between that probe and the rename is overwritten. That
    // residual race is documented in the FTP/SFTP/SCP spec and is the same window every FTP
    // client has; refusing new-name writers outright left these destinations unusable.
    if (_protocol == FileSystemCurlProtocol::Imap)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    wil::com_ptr<FileSystemCurl> writerOwner;
    writerOwner = this;
    wil::unique_hfile file;
    if (const HRESULT tempHr = Common::Files::CreateDeleteOnCloseTemporaryFile(kCurlTemporaryFileOptions, file); FAILED(tempHr))
    {
        return tempHr;
    }

    auto* impl = new (std::nothrow) TempFileWriter(std::move(writerOwner), std::move(file), _protocol, settings, _hostConnections, path, flags);
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *writer = impl;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::SupportsAtomicWriterCommit(const wchar_t* path, FileSystemFlags flags, BOOL* supported) noexcept
{
    if (supported == nullptr)
    {
        return E_POINTER;
    }
    *supported = FALSE;
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    constexpr uint32_t knownFlags =
        FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR;
    const uint32_t requestedFlags = static_cast<uint32_t>(flags);
    if ((requestedFlags & ~knownFlags) != 0u ||
        ((requestedFlags & FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) != 0u && (requestedFlags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) == 0u))
    {
        return E_INVALIDARG;
    }

    // The temp-file writer publishes with one server-side rename of a unique staged sibling, so
    // the requested path appears atomically for FTP/SFTP/SCP. IMAP has no writer at all.
    *supported = _protocol != FileSystemCurlProtocol::Imap ? TRUE : FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetFileBasicInformation([[maybe_unused]] const wchar_t* path, FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    if (info->sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }

    info->creationTime   = 0;
    info->lastAccessTime = 0;
    info->lastWriteTime  = 0;
    info->attributes     = 0;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FilesInformationCurl::Entry entry{};
    const HRESULT hr = ResolveLocationWithAuthRetry(_protocol,
                                                    settings,
                                                    path,
                                                    _hostConnections.get(),
                                                    true,
                                                    [&](const ResolvedLocation& resolved) noexcept
    {
        entry = {};
        return GetEntryInfo(resolved.connection, resolved.remotePath, entry);
    });
    if (FAILED(hr))
    {
        return hr;
    }

    // Avoid propagating zero times (would map to 1601-01-01 if applied on a Win32 destination).
    if (entry.lastWriteTime == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    info->lastWriteTime  = entry.lastWriteTime;
    info->creationTime   = entry.creationTime != 0 ? entry.creationTime : entry.lastWriteTime;
    info->lastAccessTime = entry.lastAccessTime != 0 ? entry.lastAccessTime : entry.lastWriteTime;
    info->attributes     = entry.attributes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::SetFileBasicInformation([[maybe_unused]] const wchar_t* path,
                                                                  [[maybe_unused]] const FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    if (info->sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::CreateDirectory(const wchar_t* path) noexcept
{
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    const HRESULT hr = ResolveLocationWithAuthRetry(_protocol,
                                                    settings,
                                                    path,
                                                    _hostConnections.get(),
                                                    true,
                                                    [&](const ResolvedLocation& resolved) noexcept
    {
        if (resolved.remotePath == L"/")
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        const HRESULT hr = RemoteMkdir(resolved.connection, resolved.remotePath);
        if (SUCCEEDED(hr))
        {
            return S_OK;
        }

        FilesInformationCurl::Entry existing{};
        const HRESULT existsHr = GetEntryInfo(resolved.connection, resolved.remotePath, existing);
        if (SUCCEEDED(existsHr) && (existing.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        return hr;
    });

    if (SUCCEEDED(hr))
    {
        NotifySyntheticPathCreated(path);
    }
    return hr;
}

namespace
{
// Size callbacks have a different raw ABI from operation control. This adapter
// borrows both caller objects on the one calling worker, coalesces progress, and
// carries its first callback failure through Curl's transport abort machinery.
class CurlDirectorySizeControl final : public IFileSystemOperationControl
{
public:
    CurlDirectorySizeControl(FileSystemDirectorySizeResult& result, IFileSystemDirectorySizeCallback* callback, void* cookie, const wchar_t* path) noexcept
        : _result(result),
          _callback(callback),
          _cookie(cookie),
          _path(path)
    {
    }
    CurlDirectorySizeControl(const CurlDirectorySizeControl&)            = delete;
    CurlDirectorySizeControl& operator=(const CurlDirectorySizeControl&) = delete;
    CurlDirectorySizeControl(CurlDirectorySizeControl&&)                 = delete;
    CurlDirectorySizeControl& operator=(CurlDirectorySizeControl&&)      = delete;

    // Borrowed until the next SetPath; the walker restores the caller-owned root
    // before destroying a child path or frame. No worker outlives this call.
    void SetPath(const wchar_t* path) noexcept
    {
        _path = path;
    }
    [[nodiscard]] HRESULT CallbackStatus() const noexcept
    {
        return _callbackStatus;
    }

    [[nodiscard]] HRESULT ObserveEntry(bool directory) noexcept
    {
        uint64_t& count = directory ? _result.directoryCount : _result.fileCount;
        if (_scanned == (std::numeric_limits<uint64_t>::max)() || count == (std::numeric_limits<uint64_t>::max)())
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        ++_scanned;
        ++count;
        return S_OK;
    }

    [[nodiscard]] HRESULT Checkpoint() noexcept
    {
        if (_finished || FAILED(_callbackStatus) || ! _callback)
        {
            return _callbackStatus;
        }
        if (_reported && _scanned - _reportedScanned < 100u && GetTickCount64() - _reportedTick < 200u)
        {
            return S_OK;
        }
        return Report(_path);
    }

    [[nodiscard]] HRESULT Finish(HRESULT primary, bool skippedChildren = false) noexcept
    {
        if (SUCCEEDED(primary) && FAILED(_callbackStatus))
        {
            primary = _callbackStatus;
        }
        // The final snapshot is useful on failure too, but must not replace the
        // earlier provider/callback verdict. Check cancellation after this report.
        const HRESULT finalHr = Report(nullptr);
        // Transport teardown can still ask for progress. The final snapshot is
        // terminal: never re-enter caller feedback while releasing the cursors.
        _finished = true;
        _path     = nullptr;
        if (SUCCEEDED(primary) && FAILED(finalHr))
        {
            primary = finalHr;
        }
        // Skipped-child totals are not a global failure. Final feedback can still
        // cancel/fail the traversal; only an earlier global verdict wins over it.
        _result.status = SUCCEEDED(primary) && skippedChildren ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : NormalizeCancellation(primary);
        return _result.status;
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
    {
        if (! abort)
        {
            return E_POINTER;
        }
        const HRESULT hr = Checkpoint();
        *abort           = FAILED(hr) ? TRUE : FALSE;
        return hr;
    }
    HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void*) noexcept override
    {
        if (! mode)
        {
            return E_POINTER;
        }
        *mode = FILESYSTEM_DISCOVERY_AHEAD;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress*, void*) noexcept override
    {
        return S_OK;
    }

private:
    [[nodiscard]] HRESULT Report(const wchar_t* path) noexcept
    {
        if (! _callback)
        {
            return S_OK;
        }
        _reported        = true;
        _reportedScanned = _scanned;
        _reportedTick    = GetTickCount64();
        HRESULT hr       = _callback->DirectorySizeProgress(_scanned, _result.totalBytes, _result.fileCount, _result.directoryCount, path, _cookie);
        if (SUCCEEDED(hr))
        {
            BOOL cancel = FALSE;
            hr          = _callback->DirectorySizeShouldCancel(&cancel, _cookie);
            if (SUCCEEDED(hr) && cancel != FALSE)
            {
                hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
        }
        if (SUCCEEDED(_callbackStatus) && FAILED(hr))
        {
            _callbackStatus = NormalizeCancellation(hr);
        }
        return hr;
    }

    FileSystemDirectorySizeResult& _result;
    IFileSystemDirectorySizeCallback* _callback;
    void* _cookie;
    const wchar_t* _path;
    uint64_t _scanned         = 0u;
    uint64_t _reportedScanned = 0u;
    ULONGLONG _reportedTick   = 0u;
    HRESULT _callbackStatus   = S_OK;
    bool _reported            = false;
    bool _finished            = false;
};

[[nodiscard]] bool IsRecoverableCurlSizeChildFailure(HRESULT hr) noexcept
{
    // Curl's HRESULT transport mapping is not the Local provider's DWORD policy.
    // Authentication, malformed structure, overflow and callback failures are
    // global; the caller excludes callback failures before consulting this list.
    switch (hr)
    {
        case HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED):
        case HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND):
        case HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND):
        case HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION):
        case HRESULT_FROM_WIN32(ERROR_LOCK_VIOLATION):
        case HRESULT_FROM_WIN32(ERROR_NETWORK_BUSY):
        case HRESULT_FROM_WIN32(ERROR_BAD_NETPATH):
        case HRESULT_FROM_WIN32(ERROR_NETNAME_DELETED):
        case HRESULT_FROM_WIN32(ERROR_DEV_NOT_EXIST):
        case HRESULT_FROM_WIN32(ERROR_NOT_READY):
        case HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE):
        case HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED):
        case HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED):
        case HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT): return true;
        default: return false;
    }
}
} // namespace

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetDirectorySize(
    const wchar_t* path, FileSystemFlags flags, IFileSystemDirectorySizeCallback* callback, void* cookie, FileSystemDirectorySizeResult* result) noexcept
{
    if (result == nullptr)
    {
        return E_POINTER;
    }

    if (result->sizeBytes != sizeof(FileSystemDirectorySizeResult))
    {
        return E_INVALIDARG;
    }

    result->totalBytes     = 0;
    result->fileCount      = 0;
    result->directoryCount = 0;
    result->status         = S_OK;

    if (path == nullptr || path[0] == L'\0')
    {
        result->status = path == nullptr ? E_POINTER : E_INVALIDARG;
        return result->status;
    }

    using namespace Common::FileOperations;
    CurlDirectorySizeControl control(*result, callback, cookie, path);
    FileSystemOptions options{};
    options.sizeBytes        = sizeof(options);
    options.operationControl = &control;
    const CurlOperationOptionsScope operationScope(&options);
    const auto started         = std::chrono::steady_clock::now();
    uint64_t peakFrames        = 0u;
    uint64_t peakPathBytes     = 0u;
    uint64_t peakMetadataBytes = 0u;
    const auto reportMetrics   = wil::scope_exit([&]() noexcept
    {
        const std::wstring detail = std::format(L"protocol={};imapBuffered={}", ProtocolToDisplay(_protocol), _protocol == Protocol::Imap);
        Debug::Perf::Emit(
            L"FileOps.Curl.DirectorySize.Traversal", detail.c_str(), Debug::Perf::ElapsedUs(started), peakFrames, result->fileCount, result->status);
        Debug::Perf::Emit(L"FileOps.Curl.DirectorySize.Retained", detail.c_str(), 0u, peakPathBytes, peakMetadataBytes, result->status);
    });
    HRESULT hr                 = control.Checkpoint();
    if (FAILED(hr))
    {
        return control.Finish(hr);
    }
    if (std::wstring_view(path).size() > kTraversalMaxQueuedPathBytes / sizeof(wchar_t))
    {
        return control.Finish(HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW));
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    ResolvedLocation rootResolved{};
    FilesInformationCurl::Entry rootInfo{};
    const HRESULT rootHr = ResolveLocationWithAuthRetry(_protocol,
                                                        settings,
                                                        path,
                                                        _hostConnections.get(),
                                                        true,
                                                        [&](const ResolvedLocation& resolved) noexcept
    {
        rootResolved = resolved;
        CurlEntryLookupMetrics lookup{};
        const HRESULT probeHr =
            GetEntryInfo(resolved.connection, resolved.remotePath, rootInfo, &lookup, [&control]() noexcept { return control.Checkpoint(); });
        peakFrames        = lookup.metadataBytes == 0u ? 0u : 1u;
        peakPathBytes     = (std::max)(peakPathBytes, lookup.pathBytes);
        peakMetadataBytes = (std::max)(peakMetadataBytes, lookup.metadataBytes);
        return probeHr;
    });
    if (FAILED(rootHr))
    {
        return control.Finish(rootHr);
    }

    const bool recursive             = HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE);
    const ConnectionInfo& connection = rootResolved.connection;
    const auto addFile               = [&](const FilesInformationCurl::Entry& entry, std::wstring_view remotePath) noexcept -> HRESULT
    {
        HRESULT fileHr = control.ObserveEntry(false);
        if (FAILED(fileHr))
        {
            return fileHr;
        }
        uint64_t bytes = entry.sizeBytes;
        if (! entry.sizeKnown)
        {
            fileHr = control.Checkpoint();
            if (FAILED(fileHr))
            {
                return fileHr;
            }
            CurlSourceSizeCommitment size{};
            fileHr = ResolveCurlSourceSizeCommitment(connection, remotePath, 0u, false, size);
            if (FAILED(fileHr))
            {
                return fileHr;
            }
            if (! size.known)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
            bytes = size.sizeBytes;
        }
        if (bytes > (std::numeric_limits<uint64_t>::max)() - result->totalBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        result->totalBytes += bytes;
        return control.Checkpoint();
    };

    if ((rootInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
    {
        return control.Finish(addFile(rootInfo, rootResolved.remotePath));
    }

    struct Frame final
    {
        Frame()                        = default;
        Frame(const Frame&)            = delete;
        Frame& operator=(const Frame&) = delete;
        Frame(Frame&&)                 = delete;
        Frame& operator=(Frame&&)      = delete;
        std::wstring remotePath;
        std::wstring displayPath;
        CurlDirectoryCursor cursor;
        // IMAP's current mailbox/message facade is buffered internally. Charging
        // its returned storage does not qualify the facade's transient allocations.
        std::vector<FilesInformationCurl::Entry> imapEntries;
        std::vector<std::pair<std::wstring, std::wstring>> pendingDirectories;
        size_t imapIndex       = 0u;
        uint64_t pathBytes     = 0u;
        uint64_t metadataBytes = 0u;
    };
    std::vector<std::unique_ptr<Frame>> frames;
    frames.reserve(static_cast<size_t>(kTraversalMaxDepth + 1u));
    uint64_t retainedPaths    = 0u;
    uint64_t retainedMetadata = static_cast<uint64_t>(frames.capacity()) * sizeof(std::unique_ptr<Frame>);
    bool partial              = false;
    const bool imap           = connection.protocol == Protocol::Imap;
    const auto pushFrame      = [&](std::wstring remotePath, std::wstring displayPath) noexcept -> HRESULT
    {
        const uint64_t reservation = sizeof(Frame) + (imap ? 0u : CurlDirectoryCursor::kMetadataReservationBytes);
        if (frames.size() >= kTraversalMaxDepth + 1u || reservation > kTraversalMaxMetadataBytes - retainedMetadata)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        auto frame           = std::make_unique<Frame>();
        frame->remotePath    = std::move(remotePath);
        frame->displayPath   = std::move(displayPath);
        frame->pathBytes     = static_cast<uint64_t>(frame->remotePath.capacity() + frame->displayPath.capacity() + 2u) * sizeof(wchar_t);
        frame->metadataBytes = reservation;
        if (frame->pathBytes > kTraversalMaxQueuedPathBytes - retainedPaths)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        HRESULT openHr = control.Checkpoint();
        if (FAILED(openHr))
        {
            return openHr;
        }
        if (imap)
        {
            openHr = ReadDirectoryEntries(connection, frame->remotePath, frame->imapEntries);
            frame->metadataBytes += static_cast<uint64_t>(frame->imapEntries.capacity()) * sizeof(FilesInformationCurl::Entry);
            for (const auto& entry : frame->imapEntries)
            {
                frame->pathBytes += static_cast<uint64_t>(entry.name.capacity() + 1u) * sizeof(wchar_t);
            }
        }
        else
        {
            openHr = frame->cursor.Open(connection, frame->remotePath, [&control]() noexcept { return control.Checkpoint(); });
            frame->pathBytes += frame->cursor.RetainedPathBytes();
        }
        if (FAILED(openHr))
        {
            return openHr;
        }
        if (frame->pathBytes > kTraversalMaxQueuedPathBytes - retainedPaths || frame->metadataBytes > kTraversalMaxMetadataBytes - retainedMetadata)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        retainedPaths += frame->pathBytes;
        retainedMetadata += frame->metadataBytes;
        frames.push_back(std::move(frame));
        peakFrames        = (std::max)(peakFrames, static_cast<uint64_t>(frames.size()));
        peakPathBytes     = (std::max)(peakPathBytes, retainedPaths);
        peakMetadataBytes = (std::max)(peakMetadataBytes, retainedMetadata);
        return S_OK;
    };
    const auto popFrame = [&]() noexcept
    {
        control.SetPath(path);
        retainedPaths -= frames.back()->pathBytes;
        retainedMetadata -= frames.back()->metadataBytes;
        frames.pop_back();
    };
    const auto recoverable = [&](HRESULT childHr) noexcept { return SUCCEEDED(control.CallbackStatus()) && IsRecoverableCurlSizeChildFailure(childHr); };
    hr                     = pushFrame(rootResolved.remotePath, std::wstring(path));
    if (FAILED(hr))
    {
        return control.Finish(hr);
    }
    while (! frames.empty())
    {
        Frame& frame = *frames.back();
        control.SetPath(frame.displayPath.c_str());
        hr = control.Checkpoint();
        if (FAILED(hr))
        {
            return control.Finish(hr);
        }
        FilesInformationCurl::Entry entry{};
        if (imap)
        {
            hr = frame.imapIndex < frame.imapEntries.size() ? S_OK : S_FALSE;
            if (hr == S_OK)
            {
                entry = frame.imapEntries[frame.imapIndex++];
            }
        }
        else
        {
            hr = frame.cursor.Next(entry);
        }
        if (hr != S_OK)
        {
            if (hr != S_FALSE && (frames.size() == 1u || ! recoverable(hr)))
            {
                return control.Finish(hr);
            }
            if (hr == S_FALSE && ! frame.pendingDirectories.empty())
            {
                auto child                  = std::move(frame.pendingDirectories.front());
                const uint64_t pendingBytes = static_cast<uint64_t>(child.first.capacity() + child.second.capacity() + 2u) * sizeof(wchar_t);
                frame.pendingDirectories.erase(frame.pendingDirectories.begin());
                if (pendingBytes > frame.pathBytes || pendingBytes > retainedPaths)
                {
                    return control.Finish(E_UNEXPECTED);
                }
                frame.pathBytes -= pendingBytes;
                retainedPaths -= pendingBytes;
                hr = pushFrame(std::move(child.first), std::move(child.second));
                if (FAILED(hr))
                {
                    if (! recoverable(hr))
                    {
                        return control.Finish(hr);
                    }
                    partial = true;
                }
                continue;
            }
            partial = partial || FAILED(hr);
            popFrame();
            continue;
        }
        if ((entry.attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
        {
            continue;
        }
        const std::wstring childRemote  = JoinPluginPath(frame.remotePath, entry.name);
        const std::wstring childDisplay = JoinPluginPath(frame.displayPath, entry.name);
        control.SetPath(childDisplay.c_str());
        const auto restorePath = wil::scope_exit([&]() noexcept { control.SetPath(path); });
        const bool directory   = (entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (directory)
        {
            hr = control.ObserveEntry(true);
            if (SUCCEEDED(hr))
            {
                hr = control.Checkpoint();
            }
            if (SUCCEEDED(hr) && recursive)
            {
                const uint64_t pendingBytes = static_cast<uint64_t>(childRemote.capacity() + childDisplay.capacity() + 2u) * sizeof(wchar_t);
                if (pendingBytes > kTraversalMaxQueuedPathBytes - retainedPaths)
                {
                    hr = HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
                }
                else
                {
                    frame.pendingDirectories.emplace_back(childRemote, childDisplay);
                    frame.pathBytes += pendingBytes;
                    retainedPaths += pendingBytes;
                    hr = S_OK;
                }
            }
        }
        else if ((entry.attributes & FILE_ATTRIBUTE_DEVICE) != 0)
        {
            partial = true;
            continue;
        }
        else
        {
            hr = addFile(entry, childRemote);
            if (hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && SUCCEEDED(control.CallbackStatus()))
            {
                partial = true; // Observed file, unknown bytes: never fabricate a complete zero.
                continue;
            }
        }
        if (FAILED(hr))
        {
            if (! recoverable(hr))
            {
                return control.Finish(hr);
            }
            partial = true;
        }
    }
    return control.Finish(S_OK, partial);
}
