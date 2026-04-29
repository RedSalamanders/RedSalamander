#include "FileSystemS3.Internal.h"

#include <aws/s3-crt/model/GetObjectRequest.h>
#include <aws/s3-crt/model/HeadObjectRequest.h>
#include <aws/s3-crt/model/ListObjectsV2Request.h>

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstring>
#include <format>
#include <limits>
#include <string>
#include <vector>

namespace FsS3 = FileSystemS3Internal;

static HRESULT TryGetS3ObjectSummaryFromClient(Aws::S3Crt::S3CrtClient& client,
                                               const FsS3::ResolvedAwsContext& bucketCtx,
                                               std::string_view bucket,
                                               std::string_view key,
                                               uint64_t& outSizeBytes,
                                               __int64& outLastWriteTime,
                                               bool& outFound) noexcept
{
    outSizeBytes     = 0;
    outLastWriteTime = 0;
    outFound         = false;

    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

    Aws::S3Crt::Model::HeadObjectRequest req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetKey(Aws::String(key.data(), key.size()));

    const auto outcome = client.HeadObject(req);
    if (! outcome.IsSuccess())
    {
        const auto& err  = outcome.GetError();
        const HRESULT hr = FsS3::HresultFromAwsError(err);
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            outFound = false;
            return S_OK;
        }

        const std::wstring details = std::format(L"bucket='{}' key='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(key));
        FsS3::LogAwsFailure(L"S3", L"HeadObject", bucketCtx, err, details);
        return hr;
    }

    const auto& result = outcome.GetResult();
    outFound           = true;
    outSizeBytes       = static_cast<uint64_t>(result.GetContentLength());
    outLastWriteTime   = FsS3::AwsDateTimeToFileTime64(result.GetLastModified());
    return S_OK;
}

HRESULT FsS3::TryGetS3ObjectSummary(FileSystemS3& fs,
                                    const ResolvedAwsContext& bucketCtx,
                                    std::string_view bucket,
                                    std::string_view key,
                                    uint64_t& outSizeBytes,
                                    __int64& outLastWriteTime,
                                    bool& outFound) noexcept
{
    const auto client = GetS3Client(fs, bucketCtx);
    return TryGetS3ObjectSummaryFromClient(*client, bucketCtx, bucket, key, outSizeBytes, outLastWriteTime, outFound);
}

namespace
{
[[nodiscard]] HRESULT ReadFileToStringUtf8(HANDLE file, std::string& out) noexcept
{
    out.clear();

    if (file == nullptr || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    const HRESULT seekHr = FsS3::ResetFilePointerToStart(file);
    if (FAILED(seekHr))
    {
        return seekHr;
    }

    std::array<char, 64 * 1024> buffer{};
    while (true)
    {
        DWORD read = 0;
        if (ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (read == 0)
        {
            break;
        }

        out.append(buffer.data(), static_cast<size_t>(read));
    }

    return S_OK;
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
        if (! ppvObject)
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

class S3RangedFileReader final : public IFileReader
{
public:
    S3RangedFileReader(FsS3::ResolvedAwsContext bucketCtx, std::string bucket, std::string key, std::shared_ptr<Aws::S3Crt::S3CrtClient> client) noexcept
        : _bucketCtx(std::move(bucketCtx)),
          _bucket(std::move(bucket)),
          _key(std::move(key)),
          _client(std::move(client))
    {
        FsS3::AwsSdkLifetime::AddRef();
    }

    S3RangedFileReader(const S3RangedFileReader&)            = delete;
    S3RangedFileReader(S3RangedFileReader&&)                 = delete;
    S3RangedFileReader& operator=(const S3RangedFileReader&) = delete;
    S3RangedFileReader& operator=(S3RangedFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
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

        *sizeBytes = 0;

        const HRESULT hr = EnsureSizeKnown();
        if (FAILED(hr))
        {
            return hr;
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

        uint64_t base = 0;
        if (origin == FILE_BEGIN)
        {
            base = 0;
        }
        else if (origin == FILE_CURRENT)
        {
            base = _position;
        }
        else
        {
            const HRESULT hr = EnsureSizeKnown();
            if (FAILED(hr))
            {
                return hr;
            }
            base = _sizeBytes;
        }

        uint64_t moved = 0;
        if (offset >= 0)
        {
            const uint64_t delta = static_cast<uint64_t>(offset);
            if (base > std::numeric_limits<uint64_t>::max() - delta)
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            moved = base + delta;
        }
        else
        {
            const uint64_t delta = static_cast<uint64_t>(-(offset + 1)) + 1u;
            if (base < delta)
            {
                return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
            }
            moved = base - delta;
        }

        _position    = moved;
        _bufferHave  = 0;
        _bufferStart = _position;
        *newPosition = _position;
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

        if (_sizeKnown && _position >= _sizeBytes)
        {
            return S_OK;
        }

        unsigned long totalRead = 0;
        auto* out               = static_cast<std::byte*>(buffer);

        while (totalRead < bytesToRead)
        {
            if (_bufferHave == 0 || _position < _bufferStart || _position >= (_bufferStart + _bufferHave))
            {
                const HRESULT hr = FillBufferFrom(_position);
                if (FAILED(hr))
                {
                    return hr;
                }

                if (_bufferHave == 0)
                {
                    break; // EOF
                }
            }

            const uint64_t offset64 = _position - _bufferStart;
            if (offset64 > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }

            const size_t offset = static_cast<size_t>(offset64);
            if (offset >= _bufferHave)
            {
                _bufferHave = 0;
                continue;
            }

            const size_t available = _bufferHave - offset;
            const size_t wanted    = static_cast<size_t>(bytesToRead - totalRead);
            const size_t take      = std::min(available, wanted);

            std::memcpy(out + totalRead, _buffer.data() + offset, take);
            totalRead += static_cast<unsigned long>(take);
            _position += static_cast<uint64_t>(take);

            if (_sizeKnown && _position >= _sizeBytes)
            {
                break;
            }
        }

        *bytesRead = totalRead;
        return S_OK;
    }

private:
    ~S3RangedFileReader()
    {
        FsS3::AwsSdkLifetime::Release();
    }

    [[nodiscard]] HRESULT EnsureSizeKnown() noexcept
    {
        if (_sizeKnown)
        {
            return S_OK;
        }

        uint64_t sizeBytes    = 0;
        __int64 lastWriteTime = 0;
        bool found            = false;
        const HRESULT hr      = TryGetS3ObjectSummaryFromClient(*_client, _bucketCtx, _bucket, _key, sizeBytes, lastWriteTime, found);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! found)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        _sizeBytes = sizeBytes;
        _sizeKnown = true;
        return S_OK;
    }

    [[nodiscard]] HRESULT FillBufferFrom(uint64_t start) noexcept
    {
        constexpr size_t kChunkBytes = 8 * 1024 * 1024;

        _bufferStart = start;
        _bufferHave  = 0;

        if (_sizeKnown && start >= _sizeBytes)
        {
            return S_OK;
        }

        uint64_t maxBytes = kChunkBytes;
        if (_sizeKnown)
        {
            const uint64_t remaining = _sizeBytes - start;
            maxBytes                 = std::min<uint64_t>(maxBytes, remaining);
        }

        if (maxBytes == 0u)
        {
            return S_OK;
        }

        if (maxBytes > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        if (_buffer.size() < static_cast<size_t>(maxBytes))
        {
            _buffer.resize(static_cast<size_t>(maxBytes));
        }

        const uint64_t endInclusive = start + maxBytes - 1u;
        const std::string range     = std::string("bytes=") + std::to_string(start) + "-" + std::to_string(endInclusive);

        Aws::S3Crt::Model::GetObjectRequest req;
        req.SetBucket(Aws::String(_bucket.data(), _bucket.size()));
        req.SetKey(Aws::String(_key.data(), _key.size()));
        req.SetRange(Aws::String(range.data(), range.size()));

        auto outcome = _client->GetObject(req);
        if (! outcome.IsSuccess())
        {
            const auto& err = outcome.GetError();
            const std::wstring details =
                std::format(L"bucket='{}' key='{}' range='{}'", FsS3::Utf16FromUtf8(_bucket), FsS3::Utf16FromUtf8(_key), FsS3::Utf16FromUtf8(range));
            FsS3::LogAwsFailure(L"S3", L"GetObject", _bucketCtx, err, details);
            return FsS3::HresultFromAwsError(err);
        }

        auto result           = outcome.GetResultWithOwnership();
        Aws::IOStream& stream = result.GetBody();

        size_t total = 0;
        while (stream.good() && total < static_cast<size_t>(maxBytes))
        {
            const size_t want = static_cast<size_t>(maxBytes) - total;
            stream.read(reinterpret_cast<char*>(_buffer.data()) + total, static_cast<std::streamsize>(want));
            const std::streamsize got = stream.gcount();
            if (got <= 0)
            {
                break;
            }
            total += static_cast<size_t>(got);
        }

        _bufferHave = total;
        if (! _sizeKnown && total < static_cast<size_t>(maxBytes))
        {
            _sizeBytes = _bufferStart + static_cast<uint64_t>(total);
            _sizeKnown = true;
        }

        return S_OK;
    }

    std::atomic_ulong _refCount{1};

    FsS3::ResolvedAwsContext _bucketCtx;
    std::string _bucket;
    std::string _key;
    std::shared_ptr<Aws::S3Crt::S3CrtClient> _client;

    bool _sizeKnown     = false;
    uint64_t _sizeBytes = 0;

    uint64_t _position    = 0;
    uint64_t _bufferStart = 0;
    size_t _bufferHave    = 0;
    std::vector<std::byte> _buffer;
};

[[nodiscard]] HRESULT EnsureWritableS3Target(
    FileSystemS3& owner, const FsS3::ResolvedAwsContext& bucketCtx, std::string_view bucket, std::string_view key, bool allowOverwrite) noexcept
{
    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

    uint64_t existingSize     = 0;
    __int64 existingLastWrite = 0;
    bool found                = false;
    const HRESULT existsHr    = FsS3::TryGetS3ObjectSummary(owner, bucketCtx, bucket, key, existingSize, existingLastWrite, found);
    if (FAILED(existsHr))
    {
        return existsHr;
    }
    if (found && ! allowOverwrite)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }

    std::string trimmedKey(key);
    while (! trimmedKey.empty() && trimmedKey.back() == '/')
    {
        trimmedKey.pop_back();
    }

    for (size_t slash = trimmedKey.find('/'); slash != std::string::npos; slash = trimmedKey.find('/', slash + 1u))
    {
        const std::string ancestor = trimmedKey.substr(0, slash);
        uint64_t ancestorSize      = 0;
        __int64 ancestorLastWrite  = 0;
        bool ancestorFound         = false;
        const HRESULT ancestorHr   = FsS3::TryGetS3ObjectSummary(owner, bucketCtx, bucket, ancestor, ancestorSize, ancestorLastWrite, ancestorFound);
        if (FAILED(ancestorHr))
        {
            return ancestorHr;
        }
        if (ancestorFound)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
    }

    std::string prefix(key);
    if (prefix.back() != '/')
    {
        prefix.push_back('/');
    }

    const auto client = FsS3::GetS3Client(owner, bucketCtx);
    Aws::S3Crt::Model::ListObjectsV2Request req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetPrefix(Aws::String(prefix.data(), prefix.size()));
    req.SetMaxKeys(1);

    const auto outcome = client->ListObjectsV2(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' prefix='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(prefix));
        FsS3::LogAwsFailure(L"S3", L"ListObjectsV2", bucketCtx, err, details);
        return FsS3::HresultFromAwsError(err);
    }

    if (! outcome.GetResult().GetContents().empty())
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    return S_OK;
}

class MultipartS3FileWriter final : public IFileWriter
{
public:
    MultipartS3FileWriter(
        FileSystemS3* owner, FsS3::ResolvedAwsContext bucketCtx, std::string bucket, std::string key, std::wstring pluginPath, bool allowOverwrite) noexcept
        : _bucketCtx(std::move(bucketCtx)),
          _bucket(std::move(bucket)),
          _key(std::move(key)),
          _pluginPath(std::move(pluginPath)),
          _allowOverwrite(allowOverwrite)
    {
        if (owner)
        {
            _owner = owner;
        }
    }

    MultipartS3FileWriter(const MultipartS3FileWriter&)            = delete;
    MultipartS3FileWriter(MultipartS3FileWriter&&)                 = delete;
    MultipartS3FileWriter& operator=(const MultipartS3FileWriter&) = delete;
    MultipartS3FileWriter& operator=(MultipartS3FileWriter&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
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

        *positionBytes = _position;
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

        if (_committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        if (FAILED(_failedHr))
        {
            return _failedHr;
        }

        if (bytesToWrite > (std::numeric_limits<size_t>::max)() - _buffer.size())
        {
            _failedHr = E_OUTOFMEMORY;
            return _failedHr;
        }

        const auto* src = static_cast<const std::byte*>(buffer);
        _buffer.insert(_buffer.end(), src, src + bytesToWrite);

        const HRESULT hr = FlushBufferedParts();
        if (FAILED(hr))
        {
            _failedHr = hr;
            return hr;
        }

        _position += bytesToWrite;
        *bytesWritten = bytesToWrite;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override
    {
        if (_committed)
        {
            return S_OK;
        }

        if (FAILED(_failedHr))
        {
            return _failedHr;
        }

        if (! _owner)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        HRESULT hr = EnsureWritableS3Target(*_owner.get(), _bucketCtx, _bucket, _key, _allowOverwrite);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! _hasSession)
        {
            hr = FsS3::PutS3ObjectFromMemory(*_owner.get(), _bucketCtx, _bucket, _key, _buffer.empty() ? nullptr : _buffer.data(), _buffer.size());
            if (FAILED(hr))
            {
                _failedHr = hr;
                return hr;
            }
        }
        else
        {
            if (! _buffer.empty())
            {
                hr = UploadPart(_buffer.data(), _buffer.size());
                if (FAILED(hr))
                {
                    const HRESULT abortHr = AbortMultipartUpload();
                    _failedHr             = FAILED(abortHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : hr;
                    return _failedHr;
                }
                _buffer.clear();
            }

            if (_parts.empty())
            {
                const HRESULT abortHr = AbortMultipartUpload();
                return FAILED(abortHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            hr = FsS3::CompleteS3MultipartUpload(*_owner.get(), _session, _parts);
            if (FAILED(hr))
            {
                const HRESULT abortHr = AbortMultipartUpload();
                _failedHr             = FAILED(abortHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : hr;
                return _failedHr;
            }

            _hasSession = false;
            _parts.clear();
        }

        _committed = true;
        _owner->NotifySyntheticPathCreated(_pluginPath);
        return S_OK;
    }

private:
    ~MultipartS3FileWriter()
    {
        if (! _committed)
        {
            const HRESULT hr = AbortMultipartUpload();
            if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_INVALID_STATE))
            {
                Debug::Warning(L"S3: multipart upload cleanup failed path='{}' hr=0x{:08X}.", _pluginPath, static_cast<unsigned long>(hr));
            }
        }
    }

    HRESULT EnsureMultipartSession() noexcept
    {
        if (_hasSession)
        {
            return S_OK;
        }

        if (! _owner)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        HRESULT hr = FsS3::BeginS3MultipartUpload(*_owner.get(), _bucketCtx, _bucket, _key, _session);
        if (FAILED(hr))
        {
            return hr;
        }

        _hasSession = true;
        return S_OK;
    }

    HRESULT UploadPart(const void* data, size_t sizeBytes) noexcept
    {
        if (! _hasSession)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        if (_nextPartNumber > 10000)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }

        std::string etag;
        HRESULT hr = FsS3::UploadS3MultipartPartFromMemory(*_owner.get(), _session, _nextPartNumber, data, sizeBytes, etag);
        if (FAILED(hr))
        {
            return hr;
        }

        FsS3::S3MultipartUploadedPart part{};
        part.partNumber = _nextPartNumber;
        part.eTag       = std::move(etag);
        _parts.push_back(std::move(part));
        ++_nextPartNumber;
        return S_OK;
    }

    HRESULT FlushBufferedParts() noexcept
    {
        while (_buffer.size() >= FsS3::kMultipartMinPartSizeBytes)
        {
            HRESULT hr = EnsureMultipartSession();
            if (FAILED(hr))
            {
                return hr;
            }

            hr = UploadPart(_buffer.data(), FsS3::kMultipartMinPartSizeBytes);
            if (FAILED(hr))
            {
                return hr;
            }

            const size_t remainingBytes = _buffer.size() - FsS3::kMultipartMinPartSizeBytes;
            if (remainingBytes > 0)
            {
                std::memmove(_buffer.data(), _buffer.data() + FsS3::kMultipartMinPartSizeBytes, remainingBytes);
            }
            _buffer.resize(remainingBytes);
        }

        return S_OK;
    }

    HRESULT AbortMultipartUpload() noexcept
    {
        if (! _hasSession || ! _owner)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        const HRESULT hr = FsS3::AbortS3MultipartUpload(*_owner.get(), _session);
        _hasSession      = false;
        _parts.clear();
        return hr;
    }

    std::atomic_ulong _refCount{1};
    wil::com_ptr<FileSystemS3> _owner;
    FsS3::ResolvedAwsContext _bucketCtx;
    std::string _bucket;
    std::string _key;
    std::wstring _pluginPath;
    std::vector<std::byte> _buffer;
    std::vector<FsS3::S3MultipartUploadedPart> _parts;
    FsS3::S3MultipartUploadSession _session{};
    uint64_t _position   = 0;
    HRESULT _failedHr    = S_OK;
    int _nextPartNumber  = 1;
    bool _allowOverwrite = false;
    bool _committed      = false;
    bool _hasSession     = false;
};
} // namespace

HRESULT STDMETHODCALLTYPE FileSystemS3::GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept
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

    FsS3::ResolvedAwsContext ctx{};
    std::wstring canonical;
    const HRESULT hr = FsS3::ResolveAwsContext(_mode, settings, path, _hostConnections.get(), true, ctx, canonical);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring normalized = FsS3::NormalizePluginPath(canonical);

    if (_mode == FileSystemS3Mode::S3)
    {
        if (normalized == L"/" || normalized.empty())
        {
            *fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
            return S_OK;
        }

        if (! normalized.empty() && normalized.back() == L'/')
        {
            *fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
            return S_OK;
        }

        const auto segments = FsS3::SplitPathSegments(normalized);
        if (segments.size() <= 1)
        {
            *fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
            return S_OK;
        }

        const std::string bucket = FsS3::Utf8FromUtf16(segments[0]);
        if (bucket.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }

        std::wstring keyWide;
        for (size_t i = 1; i < segments.size(); ++i)
        {
            if (i > 1)
            {
                keyWide.push_back(L'/');
            }
            keyWide.append(segments[i]);
        }

        const std::string key = FsS3::Utf8FromUtf16(keyWide);
        if (key.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }

        FsS3::ResolvedAwsContext bucketCtx{};
        HRESULT bucketHr = FsS3::ResolveS3ContextForBucket(*this, ctx, segments[0], bucketCtx);
        if (FAILED(bucketHr))
        {
            return bucketHr;
        }

        uint64_t sizeBytes    = 0;
        __int64 lastWriteTime = 0;
        bool foundFile        = false;
        const HRESULT objHr   = FsS3::TryGetS3ObjectSummary(*this, bucketCtx, bucket, key, sizeBytes, lastWriteTime, foundFile);
        if (FAILED(objHr))
        {
            return objHr;
        }
        if (foundFile)
        {
            *fileAttributes = FILE_ATTRIBUTE_NORMAL;
            return S_OK;
        }

        // S3 has no intrinsic directories; treat a non-empty prefix as a directory.
        std::string prefix = key;
        if (prefix.back() != '/')
        {
            prefix.push_back('/');
        }

        const auto client = FsS3::GetS3Client(*this, bucketCtx);
        Aws::S3Crt::Model::ListObjectsV2Request req;
        req.SetBucket(Aws::String(bucket.data(), bucket.size()));
        req.SetPrefix(Aws::String(prefix.data(), prefix.size()));
        req.SetMaxKeys(1);

        const auto outcome = client->ListObjectsV2(req);
        if (! outcome.IsSuccess())
        {
            const auto& err            = outcome.GetError();
            const std::wstring details = std::format(L"bucket='{}' prefix='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(prefix));
            FsS3::LogAwsFailure(L"S3", L"ListObjectsV2", bucketCtx, err, details);
            return FsS3::HresultFromAwsError(err);
        }

        if (! outcome.GetResult().GetContents().empty())
        {
            *fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
            return S_OK;
        }

        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    // S3 Tables
    if (normalized == L"/" || normalized.empty())
    {
        *fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        return S_OK;
    }

    const auto segments = FsS3::SplitPathSegments(normalized);
    if (segments.size() <= 2)
    {
        *fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        return S_OK;
    }
    if (segments.size() == 3)
    {
        *fileAttributes = FILE_ATTRIBUTE_NORMAL;
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
}

HRESULT STDMETHODCALLTYPE FileSystemS3::CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept
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

    FsS3::ResolvedAwsContext ctx{};
    std::wstring canonical;
    HRESULT hr = FsS3::ResolveAwsContext(_mode, settings, path, _hostConnections.get(), true, ctx, canonical);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring normalized = FsS3::NormalizePluginPath(canonical);
    if (normalized == L"/" || normalized.empty() || (! normalized.empty() && normalized.back() == L'/'))
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    if (_mode == FileSystemS3Mode::S3)
    {
        const auto segments = FsS3::SplitPathSegments(normalized);
        if (segments.size() < 2)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        std::string bucket = FsS3::Utf8FromUtf16(segments[0]);
        if (bucket.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }

        std::wstring keyWide;
        for (size_t i = 1; i < segments.size(); ++i)
        {
            if (i > 1)
            {
                keyWide.push_back(L'/');
            }
            keyWide.append(segments[i]);
        }

        std::string key = FsS3::Utf8FromUtf16(keyWide);
        if (key.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }

        FsS3::ResolvedAwsContext bucketCtx{};
        hr = FsS3::ResolveS3ContextForBucket(*this, ctx, segments[0], bucketCtx);
        if (FAILED(hr))
        {
            return hr;
        }

        const auto client = FsS3::GetS3Client(*this, bucketCtx);

        auto* impl = new (std::nothrow) S3RangedFileReader(std::move(bucketCtx), std::move(bucket), std::move(key), client);
        if (! impl)
        {
            return E_OUTOFMEMORY;
        }

        *reader = impl;
        return S_OK;
    }

    // S3 Tables: materialize the JSON document into a temp file (small, seekable).
    const auto segments = FsS3::SplitPathSegments(normalized);
    if (segments.size() != 3)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    std::wstring_view tableLeaf         = segments[2];
    constexpr std::wstring_view kSuffix = L".table.json";
    if (tableLeaf.size() >= kSuffix.size() && OrdinalString::EqualsNoCase(tableLeaf.substr(tableLeaf.size() - kSuffix.size()), kSuffix))
    {
        tableLeaf = tableLeaf.substr(0, tableLeaf.size() - kSuffix.size());
    }

    if (tableLeaf.empty())
    {
        return E_INVALIDARG;
    }

    wil::unique_hfile file = FsS3::CreateTemporaryDeleteOnCloseFile();
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    hr = FsS3::WriteS3TableInfoJson(*this, ctx, segments[0], segments[1], tableLeaf, file);
    if (FAILED(hr))
    {
        return hr;
    }

    uint64_t sizeBytes = 0;
    hr                 = FsS3::GetFileSizeBytes(file.get(), sizeBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = FsS3::ResetFilePointerToStart(file.get());
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
}

HRESULT STDMETHODCALLTYPE FileSystemS3::CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept
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

    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool allowOverwrite = (static_cast<unsigned long>(flags) & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;

    unsigned long existingAttrs = 0;
    const HRESULT hrAttr        = GetAttributes(path, &existingAttrs);
    if (SUCCEEDED(hrAttr))
    {
        if ((existingAttrs & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        if (! allowOverwrite)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        }
    }
    else if (hrAttr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && hrAttr != HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
    {
        return hrAttr;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FsS3::ResolvedAwsContext ctx{};
    std::wstring canonical;
    HRESULT hr = FsS3::ResolveAwsContext(_mode, settings, path, _hostConnections.get(), true, ctx, canonical);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring normalized = FsS3::NormalizePluginPath(canonical);
    if (normalized == L"/" || normalized.empty() || (! normalized.empty() && normalized.back() == L'/'))
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    const auto segments = FsS3::SplitPathSegments(normalized);
    if (segments.size() < 2)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    const std::string bucket = FsS3::Utf8FromUtf16(segments[0]);
    if (bucket.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    std::wstring keyWide;
    for (size_t i = 1; i < segments.size(); ++i)
    {
        if (i > 1)
        {
            keyWide.push_back(L'/');
        }
        keyWide.append(segments[i]);
    }

    const std::string key = FsS3::Utf8FromUtf16(keyWide);
    if (key.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    FsS3::ResolvedAwsContext bucketCtx{};
    hr = FsS3::ResolveS3ContextForBucket(*this, ctx, segments[0], bucketCtx);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = EnsureWritableS3Target(*this, bucketCtx, bucket, key, allowOverwrite);
    if (FAILED(hr))
    {
        return hr;
    }

    auto* impl = new (std::nothrow) MultipartS3FileWriter(this, std::move(bucketCtx), std::string(bucket), std::string(key), path, allowOverwrite);
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *writer = impl;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetFileBasicInformation([[maybe_unused]] const wchar_t* path, FileSystemBasicInformation* info) noexcept
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

    FsS3::ResolvedAwsContext ctx{};
    std::wstring canonical;
    const HRESULT hr = FsS3::ResolveAwsContext(_mode, settings, path, _hostConnections.get(), true, ctx, canonical);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring normalized = FsS3::NormalizePluginPath(canonical);

    // Only file paths provide meaningful basic info for cross-FS metadata propagation.
    if (normalized == L"/" || normalized.empty() || (! normalized.empty() && normalized.back() == L'/'))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const auto segments = FsS3::SplitPathSegments(normalized);
    if (segments.size() < 2)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const std::string bucket = FsS3::Utf8FromUtf16(segments[0]);
    if (bucket.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    std::wstring keyWide;
    for (size_t i = 1; i < segments.size(); ++i)
    {
        if (i > 1)
        {
            keyWide.push_back(L'/');
        }
        keyWide.append(segments[i]);
    }

    const std::string key = FsS3::Utf8FromUtf16(keyWide);
    if (key.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    FsS3::ResolvedAwsContext bucketCtx{};
    HRESULT bucketHr = FsS3::ResolveS3ContextForBucket(*this, ctx, segments[0], bucketCtx);
    if (FAILED(bucketHr))
    {
        return bucketHr;
    }

    uint64_t sizeBytes    = 0;
    __int64 lastWriteTime = 0;
    bool found            = false;
    const HRESULT objHr   = FsS3::TryGetS3ObjectSummary(*this, bucketCtx, bucket, key, sizeBytes, lastWriteTime, found);
    if (FAILED(objHr))
    {
        return objHr;
    }
    if (! found)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    // Avoid propagating zero times (would map to 1601-01-01 if applied on a Win32 destination).
    if (lastWriteTime == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    info->attributes     = FILE_ATTRIBUTE_NORMAL;
    info->lastWriteTime  = lastWriteTime;
    info->creationTime   = lastWriteTime;
    info->lastAccessTime = lastWriteTime;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::SetFileBasicInformation([[maybe_unused]] const wchar_t* path,
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

HRESULT STDMETHODCALLTYPE FileSystemS3::GetItemProperties([[maybe_unused]] const wchar_t* path, const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *jsonUtf8 = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FsS3::ResolvedAwsContext ctx{};
    std::wstring canonical;
    HRESULT hr = FsS3::ResolveAwsContext(_mode, settings, path, _hostConnections.get(), true, ctx, canonical);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring normalized = FsS3::NormalizePluginPath(canonical);
    const auto segments           = FsS3::SplitPathSegments(normalized);

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }
    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    yyjson_mut_doc_set_root(doc, root);

    yyjson_mut_obj_add_int(doc, root, "version", 1);
    yyjson_mut_obj_add_str(doc, root, "title", "properties");

    yyjson_mut_val* sections = yyjson_mut_arr(doc);
    yyjson_mut_obj_add_val(doc, root, "sections", sections);

    auto addSection = [&](const char* title) -> yyjson_mut_val*
    {
        yyjson_mut_val* section = yyjson_mut_obj(doc);
        yyjson_mut_obj_add_str(doc, section, "title", title);

        yyjson_mut_val* fields = yyjson_mut_arr(doc);
        yyjson_mut_obj_add_val(doc, section, "fields", fields);

        yyjson_mut_arr_add_val(sections, section);
        return fields;
    };

    auto addField = [&](yyjson_mut_val* fields, const char* key, const std::string& value)
    {
        yyjson_mut_val* field = yyjson_mut_obj(doc);
        yyjson_mut_obj_add_strcpy(doc, field, "key", key);
        yyjson_mut_obj_add_strncpy(doc, field, "value", value.data(), value.size());
        yyjson_mut_arr_add_val(fields, field);
    };

    yyjson_mut_val* general = addSection("general");
    if (normalized == L"/" || normalized.empty())
    {
        addField(general, "name", "/");
    }
    else if (! segments.empty())
    {
        addField(general, "name", FsS3::Utf8FromUtf16(segments.back()));
    }
    addField(general, "path", FsS3::Utf8FromUtf16(normalized));

    const char* mode = (_mode == FileSystemS3Mode::S3) ? "s3" : "s3table";
    addField(general, "mode", mode);

    yyjson_mut_val* connection = addSection("connection");
    addField(connection, "connectionName", FsS3::Utf8FromUtf16(ctx.connectionName));
    addField(connection, "region", ctx.region);
    addField(connection, "endpointOverride", ctx.endpointOverride);
    addField(connection, "useHttps", ctx.useHttps ? "true" : "false");
    addField(connection, "verifyTls", ctx.verifyTls ? "true" : "false");
    addField(connection, "useVirtualAddressing", ctx.useVirtualAddressing ? "true" : "false");
    addField(connection, "maxKeys", std::format("{}", ctx.maxKeys));
    addField(connection, "maxTableResults", std::format("{}", ctx.maxTableResults));
    addField(connection, "hasExplicitRegion", ctx.explicitRegion.has_value() ? "true" : "false");
    addField(connection, "hasAccessKeyId", ctx.accessKeyId.has_value() ? "true" : "false");
    addField(connection, "hasSecretAccessKey", ctx.secretAccessKey.has_value() ? "true" : "false");

    if (_mode == FileSystemS3Mode::S3)
    {
        bool isDirectory = normalized == L"/" || normalized.empty();
        if (! isDirectory && ! normalized.empty() && normalized.back() == L'/')
        {
            isDirectory = true;
        }

        const std::wstring_view bucketName = segments.empty() ? std::wstring_view{} : segments[0];

        yyjson_mut_val* s3 = addSection("s3");
        addField(s3, "bucket", FsS3::Utf8FromUtf16(bucketName));

        if (segments.size() <= 1)
        {
            addField(general, "type", "directory");
        }
        else
        {
            std::wstring keyWide;
            for (size_t i = 1; i < segments.size(); ++i)
            {
                if (i > 1)
                {
                    keyWide.push_back(L'/');
                }
                keyWide.append(segments[i]);
            }

            const std::string bucketUtf8 = FsS3::Utf8FromUtf16(bucketName);
            const std::string keyUtf8    = FsS3::Utf8FromUtf16(keyWide);
            if (bucketUtf8.empty() || keyUtf8.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
            }

            FsS3::ResolvedAwsContext bucketCtx{};
            hr = FsS3::ResolveS3ContextForBucket(*this, ctx, bucketName, bucketCtx);
            if (FAILED(hr))
            {
                return hr;
            }

            uint64_t sizeBytes    = 0;
            __int64 lastWriteTime = 0;
            bool found            = false;
            hr                    = FsS3::TryGetS3ObjectSummary(*this, bucketCtx, bucketUtf8, keyUtf8, sizeBytes, lastWriteTime, found);
            if (FAILED(hr))
            {
                return hr;
            }

            if (found)
            {
                addField(general, "type", "file");
                addField(general, "sizeBytes", std::format("{}", sizeBytes));
                if (lastWriteTime != 0)
                {
                    addField(general, "lastWriteTime", std::format("{}", lastWriteTime));
                }
                addField(s3, "key", keyUtf8);
            }
            else
            {
                isDirectory = true;
                addField(general, "type", "directory");
                addField(s3, "prefix", keyUtf8 + "/");
            }
        }
    }
    else
    {
        // S3 Tables
        yyjson_mut_val* s3t = addSection("s3table");

        if (segments.size() < 1)
        {
            addField(general, "type", "directory");
        }
        else
        {
            addField(s3t, "bucket", FsS3::Utf8FromUtf16(segments[0]));
        }

        if (segments.size() >= 2)
        {
            addField(s3t, "namespace", FsS3::Utf8FromUtf16(segments[1]));
        }

        if (segments.size() == 3)
        {
            std::wstring_view tableLeaf         = segments[2];
            constexpr std::wstring_view kSuffix = L".table.json";
            if (tableLeaf.size() >= kSuffix.size() && OrdinalString::EqualsNoCase(tableLeaf.substr(tableLeaf.size() - kSuffix.size()), kSuffix))
            {
                tableLeaf = tableLeaf.substr(0, tableLeaf.size() - kSuffix.size());
            }

            wil::unique_hfile infoFile;
            hr = FsS3::WriteS3TableInfoJson(*this, ctx, segments[0], segments[1], tableLeaf, infoFile);
            if (FAILED(hr))
            {
                return hr;
            }

            std::string infoText;
            hr = ReadFileToStringUtf8(infoFile.get(), infoText);
            if (SUCCEEDED(hr) && ! infoText.empty())
            {
                yyjson_read_err err{};
                yyjson_doc* infoDoc = yyjson_read_opts(infoText.data(), infoText.size(), YYJSON_READ_NOFLAG, nullptr, &err);
                if (infoDoc)
                {
                    auto freeInfoDoc     = wil::scope_exit([&] { yyjson_doc_free(infoDoc); });
                    yyjson_val* infoRoot = yyjson_doc_get_root(infoDoc);
                    if (infoRoot && yyjson_is_obj(infoRoot))
                    {
                        if (const auto name = FsS3::TryGetJsonString(infoRoot, "name"); name.has_value())
                        {
                            addField(s3t, "tableName", FsS3::Utf8FromUtf16(name.value()));
                        }
                        if (const auto arn = FsS3::TryGetJsonString(infoRoot, "tableArn"); arn.has_value())
                        {
                            addField(s3t, "tableArn", FsS3::Utf8FromUtf16(arn.value()));
                        }
                        if (const auto metaLoc = FsS3::TryGetJsonString(infoRoot, "metadataLocation"); metaLoc.has_value())
                        {
                            addField(s3t, "metadataLocation", FsS3::Utf8FromUtf16(metaLoc.value()));
                        }
                        if (const auto whLoc = FsS3::TryGetJsonString(infoRoot, "warehouseLocation"); whLoc.has_value())
                        {
                            addField(s3t, "warehouseLocation", FsS3::Utf8FromUtf16(whLoc.value()));
                        }
                        if (const auto ver = FsS3::TryGetJsonString(infoRoot, "versionToken"); ver.has_value())
                        {
                            addField(s3t, "versionToken", FsS3::Utf8FromUtf16(ver.value()));
                        }
                        if (const auto managed = FsS3::TryGetJsonString(infoRoot, "managedByService"); managed.has_value())
                        {
                            addField(s3t, "managedByService", FsS3::Utf8FromUtf16(managed.value()));
                        }
                        if (const auto created = FsS3::TryGetJsonString(infoRoot, "createdAt"); created.has_value())
                        {
                            addField(s3t, "createdAt", FsS3::Utf8FromUtf16(created.value()));
                        }
                    }
                }
            }

            addField(general, "type", "file");
        }
        else
        {
            addField(general, "type", "directory");
        }
    }

    const char* written = yyjson_mut_write(doc, YYJSON_WRITE_NOFLAG, nullptr);
    if (! written)
    {
        return E_OUTOFMEMORY;
    }
    auto freeWritten = wil::scope_exit([&] { free(const_cast<char*>(written)); });

    {
        std::scoped_lock lock(_propertiesMutex);
        _lastPropertiesJson.assign(written);
        *jsonUtf8 = _lastPropertiesJson.c_str();
    }

    return S_OK;
}
