#include "FileSystemCurl.Internal.h"

#include <condition_variable>
#include <system_error>
#include <thread>

using namespace FileSystemCurlInternal;

namespace
{
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

class TempFileWriter final : public IFileWriter
{
public:
    TempFileWriter(wil::unique_hfile file,
                   FileSystemCurlProtocol protocol,
                   FileSystemCurl::Settings settings,
                   wil::com_ptr<IHostConnections> hostConnections,
                   std::wstring pluginPath,
                   FileSystemFlags flags) noexcept
        : _file(std::move(file)),
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

        const bool allowOverwrite = HasFlag(_flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);

        hr = ResolveLocationWithAuthRetry(_protocol,
                                          _settings,
                                          _pluginPath.c_str(),
                                          _hostConnections.get(),
                                          true,
                                          [&](const ResolvedLocation& resolved) noexcept
                                          {
                                              const HRESULT overwriteHr = EnsureOverwriteTargetFile(resolved.connection, resolved.remotePath, allowOverwrite);
                                              if (FAILED(overwriteHr))
                                              {
                                                  return overwriteHr;
                                              }

                                              return CurlUploadFromFile(resolved.connection, resolved.remotePath, _file.get(), sizeBytes, nullptr, nullptr);
                                          });
        if (FAILED(hr))
        {
            return hr;
        }

        _committed = true;
        return S_OK;
    }

private:
    ~TempFileWriter() = default;

    std::atomic_ulong _refCount{1};
    wil::unique_hfile _file;
    FileSystemCurlProtocol _protocol = FileSystemCurlProtocol::Ftp;
    FileSystemCurl::Settings _settings{};
    wil::com_ptr<IHostConnections> _hostConnections;
    std::wstring _pluginPath;
    FileSystemFlags _flags = FILESYSTEM_FLAG_NONE;
    bool _committed        = false;
};

class CurlStreamingReader final : public IFileReader
{
public:
    CurlStreamingReader(ConnectionInfo conn, std::wstring remotePath, uint64_t sizeBytes) noexcept
        : _conn(std::move(conn)),
          _remotePath(std::move(remotePath)),
          _sizeBytes(sizeBytes)
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

        try
        {
            _worker = std::jthread([this](std::stop_token stopToken) noexcept { WorkerMain(stopToken); });
        }
        catch (const std::system_error&)
        {
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

        if (newPos == _positionBytes)
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
        while (_bufferedBytes == 0)
        {
            if (FAILED(_workerHr))
            {
                return _workerHr;
            }
            if (_eof)
            {
                return S_OK;
            }
            _cvReadable.wait(lock);
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

private:
    ~CurlStreamingReader()
    {
        _stopping.store(true, std::memory_order_release);
        _cvReadable.notify_all();
        _cvWritable.notify_all();
        if (_worker.joinable())
        {
            _worker.request_stop();
        }
    }

    [[nodiscard]] size_t OnCurlWrite(const std::byte* data, size_t bytes, std::stop_token stopToken) noexcept
    {
        if (! data || bytes == 0)
        {
            return 0;
        }

        const uint64_t activeGen = _transferGeneration.load(std::memory_order_acquire);

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
            _cvWritable.wait(lock,
                             [&]() noexcept
                             {
                                 return _stopping.load(std::memory_order_acquire) || stopToken.stop_requested() ||
                                        _generation.load(std::memory_order_acquire) != activeGen || _bufferedBytes < _bufferCapacity;
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

        return bytes;
    }

    static size_t CurlWriteToStream(void* ptr, size_t size, size_t nmemb, void* userdata) noexcept
    {
        auto* self = static_cast<CurlStreamingReader*>(userdata);
        if (! self || ! ptr || size == 0 || nmemb == 0)
        {
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

            const std::string url = BuildUrl(_conn, _remotePath, false, false);
            if (url.empty())
            {
                std::scoped_lock lock(_mutex);
                _workerHr = E_INVALIDARG;
                _cvReadable.notify_all();
                return;
            }

            unique_curl_easy curl{curl_easy_init()};
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

            if (startOffset > 0)
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

            if (code == CURLE_OK)
            {
                {
                    std::scoped_lock lock(_mutex);
                    _eof = true;
                }
                _cvReadable.notify_all();

                std::unique_lock lock(_mutex);
                const uint64_t eofGen = _generation.load(std::memory_order_acquire);
                _cvWritable.wait(lock,
                                 [&]() noexcept
                                 {
                                     return _stopping.load(std::memory_order_acquire) || stopToken.stop_requested() ||
                                            _generation.load(std::memory_order_acquire) != eofGen;
                                 });
                continue;
            }

            const HRESULT hr = HResultFromCurl(code);
            std::scoped_lock lock(_mutex);
            _workerHr = hr;
            _cvReadable.notify_all();
            return;
        }
    }

    std::atomic_ulong _refCount{1};

    ConnectionInfo _conn;
    std::wstring _remotePath;

    uint64_t _sizeBytes = 0;

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
    std::atomic_bool _stopping{false};

    bool _eof         = false;
    HRESULT _workerHr = S_OK;

    std::jthread _worker;
};

class CurlStreamingWriter final : public IFileWriter
{
public:
    CurlStreamingWriter(ConnectionInfo conn, std::wstring remotePath) noexcept : _conn(std::move(conn)), _remotePath(std::move(remotePath))
    {
    }

    CurlStreamingWriter(const CurlStreamingWriter&)            = delete;
    CurlStreamingWriter(CurlStreamingWriter&&)                 = delete;
    CurlStreamingWriter& operator=(const CurlStreamingWriter&) = delete;
    CurlStreamingWriter& operator=(CurlStreamingWriter&&)      = delete;

    HRESULT Initialize() noexcept
    {
        constexpr size_t kBufferBytes = 1024u * 1024u;
        _buffer.reset(new (std::nothrow) std::byte[kBufferBytes]);
        if (! _buffer)
        {
            return E_OUTOFMEMORY;
        }
        _bufferCapacity = kBufferBytes;

        try
        {
            _worker = std::jthread([this](std::stop_token stopToken) noexcept { WorkerMain(stopToken); });
        }
        catch (const std::system_error&)
        {
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

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileWriter))
        {
            *ppvObject = static_cast<IFileWriter*>(this);
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

        std::scoped_lock lock(_mutex);
        *positionBytes = _positionBytes;
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

        size_t offset = 0;
        while (offset < bytesToWrite)
        {
            if (_stopping.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            std::unique_lock lock(_mutex);

            if (_closedForWrite)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
            }

            if (FAILED(_workerHr))
            {
                return _workerHr;
            }

            _cvWritable.wait(lock,
                             [&]() noexcept
                             { return _stopping.load(std::memory_order_acquire) || _closedForWrite || FAILED(_workerHr) || _bufferedBytes < _bufferCapacity; });

            if (_stopping.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            if (_closedForWrite)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
            }

            if (FAILED(_workerHr))
            {
                return _workerHr;
            }

            const size_t space = _bufferCapacity - _bufferedBytes;
            const size_t chunk = (std::min)(space, static_cast<size_t>(bytesToWrite) - offset);
            if (chunk == 0)
            {
                continue;
            }

            const auto* data = static_cast<const std::byte*>(buffer);

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

            if (_positionBytes > (std::numeric_limits<uint64_t>::max)() - static_cast<uint64_t>(chunk))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            _positionBytes += static_cast<uint64_t>(chunk);

            lock.unlock();
            _cvReadable.notify_all();
        }

        *bytesWritten = bytesToWrite;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override
    {
        if (_committed)
        {
            return S_OK;
        }

        {
            std::scoped_lock lock(_mutex);
            _closedForWrite = true;
        }
        _cvReadable.notify_all();
        _cvWritable.notify_all();

        if (_worker.joinable())
        {
            _worker.join();
        }

        const HRESULT hr = _workerHr;
        if (SUCCEEDED(hr))
        {
            _committed = true;
        }
        return hr;
    }

private:
    ~CurlStreamingWriter()
    {
        _stopping.store(true, std::memory_order_release);
        {
            std::scoped_lock lock(_mutex);
            _closedForWrite = true;
        }
        _cvReadable.notify_all();
        _cvWritable.notify_all();
        if (_worker.joinable())
        {
            _worker.request_stop();
        }
    }

    [[nodiscard]] size_t OnCurlRead(char* buffer, size_t bytesToRead, std::stop_token stopToken) noexcept
    {
        if (! buffer || bytesToRead == 0)
        {
            return 0;
        }

        std::unique_lock lock(_mutex);
        while (_bufferedBytes == 0)
        {
            if (_stopping.load(std::memory_order_acquire) || stopToken.stop_requested())
            {
                return CURL_READFUNC_ABORT;
            }
            if (FAILED(_workerHr))
            {
                return CURL_READFUNC_ABORT;
            }
            if (_closedForWrite)
            {
                return 0;
            }
            _cvReadable.wait(lock);
        }

        const size_t take  = (std::min)(bytesToRead, _bufferedBytes);
        const size_t first = (std::min)(take, _bufferCapacity - _readPos);

        std::memcpy(buffer, _buffer.get() + _readPos, first);
        _readPos = (_readPos + first) % _bufferCapacity;
        _bufferedBytes -= first;

        const size_t remaining = take - first;
        if (remaining > 0)
        {
            std::memcpy(buffer + first, _buffer.get() + _readPos, remaining);
            _readPos = (_readPos + remaining) % _bufferCapacity;
            _bufferedBytes -= remaining;
        }

        lock.unlock();
        _cvWritable.notify_all();

        return take;
    }

    static size_t CurlReadFromStream(char* buffer, size_t size, size_t nitems, void* instream) noexcept
    {
        auto* self = static_cast<CurlStreamingWriter*>(instream);
        if (! self || ! buffer || size == 0 || nitems == 0)
        {
            return 0;
        }

        const size_t total = size * nitems;
        return self->OnCurlRead(buffer, total, self->_worker.get_stop_token());
    }

    static int CurlProgress(void* clientp, curl_off_t /*dltotal*/, curl_off_t /*dlnow*/, curl_off_t /*ultotal*/, curl_off_t /*ulnow*/) noexcept
    {
        auto* self = static_cast<CurlStreamingWriter*>(clientp);
        if (! self)
        {
            return 0;
        }

        return (self->_stopping.load(std::memory_order_acquire) || self->_worker.get_stop_token().stop_requested()) ? 1 : 0;
    }

    void WorkerMain(std::stop_token stopToken) noexcept
    {
        const HRESULT initHr = EnsureCurlInitialized();
        if (FAILED(initHr))
        {
            std::scoped_lock lock(_mutex);
            _workerHr = initHr;
            _cvWritable.notify_all();
            return;
        }

        const std::string url = BuildUrl(_conn, _remotePath, false, false);
        if (url.empty())
        {
            std::scoped_lock lock(_mutex);
            _workerHr = E_INVALIDARG;
            _cvWritable.notify_all();
            return;
        }

        unique_curl_easy curl{curl_easy_init()};
        if (! curl)
        {
            std::scoped_lock lock(_mutex);
            _workerHr = E_OUTOFMEMORY;
            _cvWritable.notify_all();
            return;
        }

        curl_easy_setopt(curl.get(), CURLOPT_URL, url.c_str());
        curl_easy_setopt(curl.get(), CURLOPT_UPLOAD, 1L);
        curl_easy_setopt(curl.get(), CURLOPT_READFUNCTION, CurlReadFromStream);
        curl_easy_setopt(curl.get(), CURLOPT_READDATA, this);
        curl_easy_setopt(curl.get(), CURLOPT_FAILONERROR, 1L);

        curl_easy_setopt(curl.get(), CURLOPT_XFERINFOFUNCTION, CurlProgress);
        curl_easy_setopt(curl.get(), CURLOPT_XFERINFODATA, this);
        curl_easy_setopt(curl.get(), CURLOPT_NOPROGRESS, 0L);

        ApplyCommonCurlOptions(curl.get(), _conn, nullptr, true);

        const CURLcode code = curl_easy_perform(curl.get());

        if (_stopping.load(std::memory_order_acquire) || stopToken.stop_requested())
        {
            std::scoped_lock lock(_mutex);
            _workerHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            _cvWritable.notify_all();
            return;
        }

        const HRESULT hr = HResultFromCurl(code);
        std::scoped_lock lock(_mutex);
        _workerHr = hr;
        _cvWritable.notify_all();
    }

    std::atomic_ulong _refCount{1};

    ConnectionInfo _conn;
    std::wstring _remotePath;

    std::mutex _mutex;
    std::condition_variable _cvReadable;
    std::condition_variable _cvWritable;

    std::unique_ptr<std::byte[]> _buffer;
    size_t _bufferCapacity = 0;

    size_t _readPos       = 0;
    size_t _writePos      = 0;
    size_t _bufferedBytes = 0;

    uint64_t _positionBytes = 0;

    std::atomic_bool _stopping{false};
    bool _closedForWrite = false;
    bool _committed      = false;
    HRESULT _workerHr    = S_OK;

    std::jthread _worker;
};
} // namespace

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
                                                auto* impl = new (std::nothrow) CurlStreamingReader(resolved.connection, resolved.remotePath, entry.sizeBytes);
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

                                            wil::unique_hfile file = CreateTemporaryDeleteOnCloseFile();
                                            if (! file)
                                            {
                                                return HRESULT_FROM_WIN32(GetLastError());
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

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    const bool allowOverwrite = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);

    return ResolveLocationWithAuthRetry(_protocol,
                                        settings,
                                        path,
                                        _hostConnections.get(),
                                        true,
                                        [&](const ResolvedLocation& resolved) noexcept
                                        {
                                            const HRESULT overwriteHr = EnsureOverwriteTargetFile(resolved.connection, resolved.remotePath, allowOverwrite);
                                            if (FAILED(overwriteHr))
                                            {
                                                return overwriteHr;
                                            }

                                            if (resolved.connection.protocol != Protocol::Imap)
                                            {
                                                auto* impl = new (std::nothrow) CurlStreamingWriter(resolved.connection, resolved.remotePath);
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

                                                *writer = impl;
                                                return S_OK;
                                            }

                                            wil::unique_hfile file = CreateTemporaryDeleteOnCloseFile();
                                            if (! file)
                                            {
                                                return HRESULT_FROM_WIN32(GetLastError());
                                            }

                                            auto* impl = new (std::nothrow) TempFileWriter(std::move(file), _protocol, settings, _hostConnections, path, flags);
                                            if (! impl)
                                            {
                                                return E_OUTOFMEMORY;
                                            }

                                            *writer = impl;
                                            return S_OK;
                                        });
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetFileBasicInformation([[maybe_unused]] const wchar_t* path, FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    *info = {};

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

    return ResolveLocationWithAuthRetry(_protocol,
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
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::GetDirectorySize(
    const wchar_t* path, FileSystemFlags flags, IFileSystemDirectorySizeCallback* callback, void* cookie, FileSystemDirectorySizeResult* result) noexcept
{
    if (result == nullptr)
    {
        return E_POINTER;
    }

    *result        = {};
    result->status = S_OK;

    if (path == nullptr || path[0] == L'\0')
    {
        result->status = E_INVALIDARG;
        return result->status;
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
                                                            rootInfo     = {};
                                                            return GetEntryInfo(resolved.connection, resolved.remotePath, rootInfo);
                                                        });
    if (FAILED(rootHr))
    {
        result->status = rootHr;
        return result->status;
    }

    const bool recursive    = HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE);
    uint64_t scannedEntries = 0;

    auto shouldCancel = [&]() noexcept -> bool
    {
        if (! callback)
        {
            return false;
        }

        BOOL cancel = FALSE;
        if (FAILED(callback->DirectorySizeShouldCancel(&cancel, cookie)))
        {
            return false;
        }
        return cancel != FALSE;
    };

    auto reportProgress = [&](const wchar_t* currentPath) noexcept
    {
        if (! callback)
        {
            return;
        }

        callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, currentPath, cookie);
    };

    if ((rootInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
    {
        scannedEntries     = 1;
        result->totalBytes = rootInfo.sizeBytes;
        result->fileCount  = 1;

        reportProgress(path);
        if (shouldCancel())
        {
            result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            reportProgress(nullptr);
            return result->status;
        }

        reportProgress(nullptr);
        return result->status;
    }

    std::function<HRESULT(const std::wstring&)> scan;
    scan = [&](const std::wstring& directory) noexcept -> HRESULT
    {
        if (shouldCancel())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        ResolvedLocation directoryResolved{};
        const HRESULT resolveHr = ResolveLocation(_protocol, settings, directory, _hostConnections.get(), true, directoryResolved);
        if (FAILED(resolveHr))
        {
            return resolveHr;
        }

        std::vector<FilesInformationCurl::Entry> entries;
        HRESULT hr = ReadDirectoryEntries(directoryResolved.connection, directoryResolved.remotePath, entries);
        if (FAILED(hr))
        {
            return hr;
        }

        for (const auto& entry : entries)
        {
            ++scannedEntries;

            const std::wstring childPath = JoinPluginPath(directory, entry.name);

            if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
            {
                ++result->directoryCount;
                if (recursive)
                {
                    const std::wstring subDir = EnsureTrailingSlash(childPath);
                    hr                        = scan(subDir);
                    if (FAILED(hr))
                    {
                        return hr;
                    }
                }
            }
            else
            {
                ++result->fileCount;
                result->totalBytes += entry.sizeBytes;
            }

            if ((scannedEntries % 128u) == 0u)
            {
                reportProgress(childPath.c_str());
            }
        }

        return S_OK;
    };

    const std::wstring startDir = EnsureTrailingSlash(path);
    HRESULT hr                  = scan(startDir);
    hr                          = NormalizeCancellation(hr);
    result->status              = hr;

    reportProgress(nullptr);
    return result->status;
}
