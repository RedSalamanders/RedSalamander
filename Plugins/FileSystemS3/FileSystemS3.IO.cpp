#include "FileSystemS3.Internal.h"

#include <aws/s3-crt/model/GetObjectRequest.h>
#include <aws/s3-crt/model/HeadObjectRequest.h>
#include <aws/s3-crt/model/ListObjectsV2Request.h>

#include <algorithm>
#include <array>
#include <charconv>
#include <chrono>
#include <condition_variable>
#include <cstddef>
#include <cstring>
#include <deque>
#include <format>
#include <limits>
#include <semaphore>
#include <string>
#include <system_error>
#include <thread>
#include <vector>

namespace FsS3 = FileSystemS3Internal;

static HRESULT TryGetS3ObjectSummaryFromClient(Aws::S3Crt::S3CrtClient& client,
                                               const FsS3::ResolvedAwsContext& bucketCtx,
                                               std::string_view bucket,
                                               std::string_view key,
                                               uint64_t& outSizeBytes,
                                               __int64& outLastWriteTime,
                                               bool& outFound,
                                               Aws::String* outEtag      = nullptr,
                                               Aws::String* outVersionId = nullptr) noexcept
{
    outSizeBytes     = 0;
    outLastWriteTime = 0;
    outFound         = false;
    if (outEtag != nullptr)
    {
        outEtag->clear();
    }
    if (outVersionId != nullptr)
    {
        outVersionId->clear();
    }

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
    if (outEtag != nullptr)
    {
        *outEtag = result.GetETag();
    }
    if (outVersionId != nullptr)
    {
        *outVersionId = result.GetVersionId();
    }
    return S_OK;
}

void ApplyS3ReadVersionPin(Aws::S3Crt::Model::GetObjectRequest& request, const Aws::String& etag, const Aws::String& versionId) noexcept
{
    if (! versionId.empty())
    {
        request.SetVersionId(versionId);
    }
    else if (! etag.empty())
    {
        request.SetIfMatch(etag);
    }
}

[[nodiscard]] HRESULT ObserveS3ReadVersion(Aws::String& etag,
                                           Aws::String& versionId,
                                           const Aws::String& responseEtag,
                                           const Aws::String& responseVersionId) noexcept
{
    if (etag.empty() && versionId.empty())
    {
        if (responseEtag.empty() && responseVersionId.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        etag      = responseEtag;
        versionId = responseVersionId;
        return S_OK;
    }
    if ((! etag.empty() && ! responseEtag.empty() && etag != responseEtag) ||
        (! versionId.empty() && ! responseVersionId.empty() && versionId != responseVersionId))
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
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

HRESULT FsS3::ValidateS3RangeResponseLength(uint64_t expectedBytes, long long responseContentLength, uint64_t bodyBytesRead) noexcept
{
    if (responseContentLength < 0)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const uint64_t responseBytes = static_cast<uint64_t>(responseContentLength);
    if (responseBytes != expectedBytes || bodyBytesRead != expectedBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT FsS3::ValidateS3UploadReadResult(uint64_t declaredBytes, uint64_t consumedBytes, HRESULT readStatus) noexcept
{
    if (FAILED(readStatus))
    {
        return readStatus;
    }
    return consumedBytes == declaredBytes ? S_OK : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
}

namespace
{
struct S3ContentRange final
{
    uint64_t first = 0;
    uint64_t last  = 0;
    uint64_t total = 0;
};

[[nodiscard]] HRESULT ParseS3ContentRange(std::string_view text, S3ContentRange& range) noexcept
{
    constexpr std::string_view prefix = "bytes ";
    if (! text.starts_with(prefix))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    text.remove_prefix(prefix.size());
    const size_t dash  = text.find('-');
    const size_t slash = text.find('/');
    if (dash == std::string_view::npos || slash == std::string_view::npos || dash == 0u || slash <= dash + 1u || slash + 1u >= text.size())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const auto parse = [](std::string_view value, uint64_t& output) noexcept
    {
        const char* const begin    = value.data();
        const char* const end      = begin + value.size();
        const auto [cursor, error] = std::from_chars(begin, end, output);
        return error == std::errc{} && cursor == end;
    };

    if (! parse(text.substr(0u, dash), range.first) || ! parse(text.substr(dash + 1u, slash - dash - 1u), range.last) ||
        ! parse(text.substr(slash + 1u), range.total) || range.first > range.last || range.last >= range.total)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    return S_OK;
}

[[nodiscard]] bool ShouldDiscoverS3SizeFromRange(HRESULT headHr) noexcept
{
    return headHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
}
} // namespace

#if defined(_DEBUG)
void FsS3::RunDebugRangeReadContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept -> bool
    {
        if (condition)
        {
            ++passed;
            return true;
        }

        ++failed;
        Debug::Error(L"FileSystemS3 debug selftest failed: {}", message);
        return false;
    };

    check(ValidateS3RangeResponseLength(10u, 10, 10u) == S_OK, L"S3 range read contract should accept exact Content-Length and exact body bytes");
    check(ValidateS3RangeResponseLength(10u, 9, 9u) == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
          L"S3 range read contract should fail a short ranged response length");
    check(ValidateS3RangeResponseLength(10u, 10, 9u) == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
          L"S3 range read contract should fail early EOF despite matching Content-Length");
    check(ValidateS3RangeResponseLength(10u, 12, 10u) == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
          L"S3 range read contract should fail an overlong ranged response length");
    check(ValidateS3RangeResponseLength(10u, -1, 0u) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
          L"S3 range read contract should fail missing or invalid Content-Length");
    check(ValidateS3UploadReadResult(10u, 10u, S_OK) == S_OK, L"S3 upload should accept exact declared-byte consumption");
    check(ValidateS3UploadReadResult(10u, 9u, S_OK) == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
          L"S3 upload should reject fake SDK success after early source EOF");
    check(ValidateS3UploadReadResult(10u, 10u, HRESULT_FROM_WIN32(ERROR_READ_FAULT)) == HRESULT_FROM_WIN32(ERROR_READ_FAULT),
          L"S3 upload should preserve the underlying source read failure");

    S3ContentRange contentRange{};
    check(ShouldDiscoverS3SizeFromRange(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)), L"S3 reader should defer a HEAD access denial to ranged GET size discovery");
    check(! ShouldDiscoverS3SizeFromRange(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)), L"S3 reader should not hide non-permission HEAD failures");
    check(ParseS3ContentRange("bytes 0-9/25", contentRange) == S_OK && contentRange.first == 0u && contentRange.last == 9u && contentRange.total == 25u,
          L"S3 reader should discover object size from a validated Content-Range");
    check(ParseS3ContentRange("bytes 10-9/25", contentRange) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"S3 reader should reject a reversed Content-Range");
    check(ParseS3ContentRange("bytes 0-25/25", contentRange) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
          L"S3 reader should reject a Content-Range ending beyond the object size");

    Aws::String etag;
    Aws::String versionId;
    check(ObserveS3ReadVersion(etag, versionId, "\"v1\"", {}) == S_OK, L"S3 reader should capture the first response ETag");
    Aws::S3Crt::Model::GetObjectRequest etagRequest;
    ApplyS3ReadVersionPin(etagRequest, etag, versionId);
    check(etagRequest.GetIfMatch() == "\"v1\"", L"S3 reader should apply If-Match to subsequent ranges");
    check(ObserveS3ReadVersion(etag, versionId, "\"v2\"", {}) == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH),
          L"S3 reader should reject a same-size object version swap");

    etag.clear();
    versionId.clear();
    check(ObserveS3ReadVersion(etag, versionId, "\"v1\"", "version-1") == S_OK, L"S3 reader should capture an object version ID");
    Aws::S3Crt::Model::GetObjectRequest versionRequest;
    ApplyS3ReadVersionPin(versionRequest, etag, versionId);
    check(versionRequest.GetVersionId() == "version-1", L"S3 reader should prefer immutable versionId pinning when available");
}
#endif

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
        _awsRuntimeStatus = FsS3::AwsSdkLifetime::Acquire();
    }

    [[nodiscard]] HRESULT InitializationStatus() const noexcept
    {
        return _awsRuntimeStatus;
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
        if (! _sizeKnown)
        {
            const HRESULT rangeHr = FillBufferFrom(0u);
            if (FAILED(rangeHr))
            {
                return rangeHr;
            }
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
            if (! _sizeKnown)
            {
                const HRESULT rangeHr = FillBufferFrom(0u);
                if (FAILED(rangeHr))
                {
                    return rangeHr;
                }
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

        const HRESULT sizeHr = EnsureSizeKnown();
        if (FAILED(sizeHr))
        {
            return sizeHr;
        }
        if (! _sizeKnown)
        {
            const HRESULT rangeHr = FillBufferFrom(_position);
            if (FAILED(rangeHr))
            {
                return rangeHr;
            }
        }

        if (_position >= _sizeBytes)
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
        if (SUCCEEDED(_awsRuntimeStatus))
        {
            FsS3::AwsSdkLifetime::Release();
        }
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
        const HRESULT hr      = TryGetS3ObjectSummaryFromClient(*_client, _bucketCtx, _bucket, _key, sizeBytes, lastWriteTime, found, &_etag, &_versionId);
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

        if (start > (std::numeric_limits<uint64_t>::max)() - (maxBytes - 1u))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        const uint64_t endInclusive = start + maxBytes - 1u;
        const std::string range     = std::string("bytes=") + std::to_string(start) + "-" + std::to_string(endInclusive);

        Aws::S3Crt::Model::GetObjectRequest req;
        req.SetBucket(Aws::String(_bucket.data(), _bucket.size()));
        req.SetKey(Aws::String(_key.data(), _key.size()));
        req.SetRange(Aws::String(range.data(), range.size()));
        ApplyS3ReadVersionPin(req, _etag, _versionId);

        auto outcome = _client->GetObject(req);
        if (! outcome.IsSuccess())
        {
            const auto& err = outcome.GetError();
            const std::wstring details =
                std::format(L"bucket='{}' key='{}' range='{}'", FsS3::Utf16FromUtf8(_bucket), FsS3::Utf16FromUtf8(_key), FsS3::Utf16FromUtf8(range));
            FsS3::LogAwsFailure(L"S3", L"GetObject", _bucketCtx, err, details);
            return FsS3::HresultFromAwsError(err);
        }

        auto result                         = outcome.GetResultWithOwnership();
        const Aws::String responseEtag      = result.GetETag();
        const Aws::String responseVersionId = result.GetVersionId();
        const HRESULT versionHr             = ObserveS3ReadVersion(_etag, _versionId, responseEtag, responseVersionId);
        if (FAILED(versionHr))
        {
            return versionHr;
        }
        const auto responseContentLength       = result.GetContentLength();
        uint64_t expectedBytes                 = maxBytes;
        const Aws::String responseContentRange = result.GetContentRange();
        S3ContentRange contentRange{};
        const HRESULT contentRangeHr = ParseS3ContentRange(std::string_view(responseContentRange.data(), responseContentRange.size()), contentRange);
        if (FAILED(contentRangeHr) || contentRange.first != start)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (! _sizeKnown)
        {
            _sizeBytes = contentRange.total;
            _sizeKnown = true;
        }
        else if (_sizeBytes != contentRange.total)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }

        expectedBytes = std::min<uint64_t>(maxBytes, _sizeBytes - start);
        if (contentRange.last != start + expectedBytes - 1u)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
        Aws::IOStream& stream = result.GetBody();

        size_t total = 0;
        while (stream.good() && total < static_cast<size_t>(expectedBytes))
        {
            const size_t want = static_cast<size_t>(expectedBytes) - total;
            stream.read(reinterpret_cast<char*>(_buffer.data()) + total, static_cast<std::streamsize>(want));
            const std::streamsize got = stream.gcount();
            if (got <= 0)
            {
                break;
            }
            total += static_cast<size_t>(got);
        }

        _bufferHave              = total;
        const HRESULT validateHr = FsS3::ValidateS3RangeResponseLength(expectedBytes, responseContentLength, static_cast<uint64_t>(total));
        if (FAILED(validateHr))
        {
            _bufferHave = 0;
            return validateHr;
        }

        return S_OK;
    }

    std::atomic_ulong _refCount{1};

    FsS3::ResolvedAwsContext _bucketCtx;
    std::string _bucket;
    std::string _key;
    std::shared_ptr<Aws::S3Crt::S3CrtClient> _client;
    HRESULT _awsRuntimeStatus = E_UNEXPECTED;

    bool _sizeKnown     = false;
    uint64_t _sizeBytes = 0;

    uint64_t _position    = 0;
    uint64_t _bufferStart = 0;
    size_t _bufferHave    = 0;
    std::vector<std::byte> _buffer;
    Aws::String _etag;
    Aws::String _versionId;
};

[[nodiscard]] HRESULT EnsureWritableS3Target(FileSystemS3& owner,
                                             const FsS3::ResolvedAwsContext& bucketCtx,
                                             std::string_view bucket,
                                             std::string_view key,
                                             std::wstring_view pluginPath,
                                             bool allowOverwrite) noexcept
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

    std::vector<std::wstring> ancestorPluginPaths;
    std::wstring currentPluginPath = FsS3::NormalizePluginPath(pluginPath);
    const size_t ancestorCount     = static_cast<size_t>(std::count(trimmedKey.begin(), trimmedKey.end(), '/'));
    ancestorPluginPaths.reserve(ancestorCount);
    for (size_t index = 0; index < ancestorCount; ++index)
    {
        const size_t slash = currentPluginPath.find_last_of(L'/');
        if (slash == std::wstring::npos)
        {
            ancestorPluginPaths.clear();
            break;
        }
        currentPluginPath.resize(slash == 0u ? 1u : slash);
        ancestorPluginPaths.push_back(currentPluginPath);
    }
    std::reverse(ancestorPluginPaths.begin(), ancestorPluginPaths.end());

    size_t ancestorIndex = 0;
    for (size_t slash = trimmedKey.find('/'); slash != std::string::npos; slash = trimmedKey.find('/', slash + 1u), ++ancestorIndex)
    {
        if (ancestorIndex < ancestorPluginPaths.size() && owner.HasFreshWritableDirectoryValidation(ancestorPluginPaths[ancestorIndex]))
        {
            continue;
        }

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
        if (ancestorIndex < ancestorPluginPaths.size())
        {
            owner.RememberWritableDirectoryValidation(ancestorPluginPaths[ancestorIndex]);
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

#if defined(_DEBUG)
struct MultipartWriterDebugTransport final
{
    void* cookie                                                                           = nullptr;
    HRESULT (*begin)(void*, FsS3::S3MultipartUploadSession&) noexcept                      = nullptr;
    HRESULT (*upload)(void*, int, size_t, std::string&) noexcept                           = nullptr;
    HRESULT (*complete)(void*, const std::vector<FsS3::S3MultipartUploadedPart>&) noexcept = nullptr;
    HRESULT (*abort)(void*) noexcept                                                       = nullptr;
};

std::atomic<const MultipartWriterDebugTransport*> g_multipartWriterDebugTransport{nullptr};

class MultipartWriterDebugTransportScope final
{
public:
    explicit MultipartWriterDebugTransportScope(const MultipartWriterDebugTransport& transport) noexcept
    {
        g_multipartWriterDebugTransport.store(&transport, std::memory_order_release);
    }

    ~MultipartWriterDebugTransportScope() noexcept
    {
        g_multipartWriterDebugTransport.store(nullptr, std::memory_order_release);
    }

    MultipartWriterDebugTransportScope(const MultipartWriterDebugTransportScope&)            = delete;
    MultipartWriterDebugTransportScope(MultipartWriterDebugTransportScope&&)                 = delete;
    MultipartWriterDebugTransportScope& operator=(const MultipartWriterDebugTransportScope&) = delete;
    MultipartWriterDebugTransportScope& operator=(MultipartWriterDebugTransportScope&&)      = delete;
};

struct MultipartWriterDebugContext final
{
    MultipartWriterDebugContext()                                              = default;
    ~MultipartWriterDebugContext()                                             = default;
    MultipartWriterDebugContext(const MultipartWriterDebugContext&)            = delete;
    MultipartWriterDebugContext(MultipartWriterDebugContext&&)                 = delete;
    MultipartWriterDebugContext& operator=(const MultipartWriterDebugContext&) = delete;
    MultipartWriterDebugContext& operator=(MultipartWriterDebugContext&&)      = delete;

    std::mutex mutex;
    std::condition_variable cv;
    bool uploadStarted                  = false;
    bool allowUploadCompletion          = false;
    unsigned int beginCalls             = 0u;
    unsigned int uploadCalls            = 0u;
    unsigned int completeCalls          = 0u;
    unsigned int abortCalls             = 0u;
    unsigned int abortFailuresRemaining = 0u;
};
#endif

struct PendingMultipartAbort final
{
    wil::com_ptr<FileSystemS3> owner;
    FsS3::S3MultipartUploadSession session;
    std::wstring pluginPath;
    unsigned int attemptCount = 0u;
#if defined(_DEBUG)
    void* debugCookie                     = nullptr;
    HRESULT (*debugAbort)(void*) noexcept = nullptr;
#endif
};

const int kMultipartAbortCleanupModuleAnchor    = 0;
constexpr auto kMultipartAbortRetryDelay        = std::chrono::seconds(1);
constexpr size_t kMultipartAbortMaxItemsPerPass = 16u;

class PendingMultipartAbortQueue final
{
public:
    PendingMultipartAbortQueue()                                             = default;
    ~PendingMultipartAbortQueue()                                            = default;
    PendingMultipartAbortQueue(const PendingMultipartAbortQueue&)            = delete;
    PendingMultipartAbortQueue(PendingMultipartAbortQueue&&)                 = delete;
    PendingMultipartAbortQueue& operator=(const PendingMultipartAbortQueue&) = delete;
    PendingMultipartAbortQueue& operator=(PendingMultipartAbortQueue&&)      = delete;

    void Queue(FileSystemS3* owner,
               FsS3::S3MultipartUploadSession session,
               std::wstring_view pluginPath
#if defined(_DEBUG)
               ,
               void* debugCookie,
               HRESULT (*debugAbort)(void*) noexcept
#endif
               ) noexcept
    {
        auto pending        = std::make_unique<PendingMultipartAbort>();
        pending->owner      = owner;
        pending->session    = std::move(session);
        pending->pluginPath = pluginPath;
#if defined(_DEBUG)
        pending->debugCookie = debugCookie;
        pending->debugAbort  = debugAbort;
#endif

        {
            std::lock_guard lock(_mutex);
            _pending.push_back(std::move(pending));
            _nextAttempt = std::chrono::steady_clock::now();
        }
        _changed.notify_all();
        Schedule();
    }

    void Schedule() noexcept
    {
        {
            std::lock_guard lock(_mutex);
            if (_workerScheduled || _pending.empty() || std::chrono::steady_clock::now() < _nextAttempt)
            {
                return;
            }
            _workerScheduled = true;
        }

        auto work             = std::make_unique<WorkItem>();
        work->queue           = this;
        work->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kMultipartAbortCleanupModuleAnchor);
        if (! work->moduleKeepAlive)
        {
            MarkSubmissionFailed();
            Debug::Error(L"S3: unable to pin the plugin module for multipart-abort cleanup.");
            return;
        }

        const BOOL submitted = TrySubmitThreadpoolCallback(
            [](PTP_CALLBACK_INSTANCE instance, void* context) noexcept
        {
            std::unique_ptr<WorkItem> work(static_cast<WorkItem*>(context));
            if (! work)
            {
                return;
            }
            TransferModulePinToCallbackReturn(instance, work->moduleKeepAlive);
            work->queue->RunOnePass();
        },
            work.get(),
            nullptr);
        if (submitted == FALSE)
        {
            MarkSubmissionFailed();
            Debug::ErrorWithLastError(L"S3: unable to submit multipart-abort cleanup.");
            return;
        }
        work.release();
    }

    [[nodiscard]] bool IsDrained() noexcept
    {
        std::lock_guard lock(_mutex);
        return _pending.empty() && ! _workerScheduled;
    }

#if defined(_DEBUG)
    [[nodiscard]] bool WaitUntilDrained(unsigned long timeoutMs) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(timeoutMs);
        while (true)
        {
            Schedule();
            std::unique_lock lock(_mutex);
            if (_pending.empty() && ! _workerScheduled)
            {
                return true;
            }
            if (_changed.wait_until(lock, deadline) == std::cv_status::timeout)
            {
                return _pending.empty() && ! _workerScheduled;
            }
        }
    }
#endif

private:
    struct WorkItem final
    {
        WorkItem()                           = default;
        ~WorkItem()                          = default;
        WorkItem(const WorkItem&)            = delete;
        WorkItem(WorkItem&&)                 = delete;
        WorkItem& operator=(const WorkItem&) = delete;
        WorkItem& operator=(WorkItem&&)      = delete;

        PendingMultipartAbortQueue* queue = nullptr;
        wil::unique_hmodule moduleKeepAlive;
    };

    [[nodiscard]] static HRESULT Abort(const PendingMultipartAbort& pending) noexcept
    {
#if defined(_DEBUG)
        if (pending.debugAbort != nullptr)
        {
            return pending.debugAbort(pending.debugCookie);
        }
#endif
        return pending.owner ? FsS3::AbortS3MultipartUpload(*pending.owner.get(), pending.session) : HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }

    void RunOnePass() noexcept
    {
        size_t itemCount = 0u;
        {
            std::lock_guard lock(_mutex);
            itemCount = (std::min)(_pending.size(), kMultipartAbortMaxItemsPerPass);
        }

        bool retryNeeded = false;
        for (size_t index = 0u; index < itemCount; ++index)
        {
            std::unique_ptr<PendingMultipartAbort> pending;
            {
                std::lock_guard lock(_mutex);
                if (_pending.empty())
                {
                    break;
                }
                pending = std::move(_pending.front());
                _pending.pop_front();
            }

            ++pending->attemptCount;
            const HRESULT hr = Abort(*pending);
            if (FAILED(hr))
            {
                Debug::Warning(L"S3: multipart-abort cleanup still pending path='{}' attempt={} hr=0x{:08X}.",
                               pending->pluginPath,
                               pending->attemptCount,
                               static_cast<unsigned long>(hr));
                std::lock_guard lock(_mutex);
                _pending.push_back(std::move(pending));
                retryNeeded = true;
            }
        }

        {
            std::lock_guard lock(_mutex);
            _workerScheduled = false;
            _nextAttempt     = retryNeeded ? std::chrono::steady_clock::now() + kMultipartAbortRetryDelay : std::chrono::steady_clock::now();
        }
        _changed.notify_all();
    }

    void MarkSubmissionFailed() noexcept
    {
        {
            std::lock_guard lock(_mutex);
            _workerScheduled = false;
            _nextAttempt     = std::chrono::steady_clock::now() + kMultipartAbortRetryDelay;
        }
        _changed.notify_all();
    }

    std::mutex _mutex;
    std::condition_variable _changed;
    std::deque<std::unique_ptr<PendingMultipartAbort>> _pending;
    std::chrono::steady_clock::time_point _nextAttempt{};
    bool _workerScheduled = false;
};

[[nodiscard]] PendingMultipartAbortQueue& MultipartAbortQueue() noexcept
{
    static PendingMultipartAbortQueue queue;
    return queue;
}

constexpr ptrdiff_t kMultipartWriterBufferSlots = 4;
constexpr uint64_t kMultipartWriterBufferBudgetBytes =
    static_cast<uint64_t>(kMultipartWriterBufferSlots) * static_cast<uint64_t>(FsS3::kMultipartMinPartSizeBytes);
static_assert(kMultipartWriterBufferBudgetBytes == 256ull * 1024ull * 1024ull);

std::counting_semaphore<kMultipartWriterBufferSlots>& GetMultipartWriterBufferBudget() noexcept
{
    static std::counting_semaphore<kMultipartWriterBufferSlots> budget(kMultipartWriterBufferSlots);
    return budget;
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

        if (! _hasBufferBudget)
        {
            GetMultipartWriterBufferBudget().acquire();
            _hasBufferBudget = true;
        }

        const auto* src         = static_cast<const std::byte*>(buffer);
        unsigned long remaining = bytesToWrite;
        while (remaining > 0u)
        {
            if (_uploadThread.joinable())
            {
                const HRESULT collectHr = CollectAsyncPart();
                if (FAILED(collectHr))
                {
                    _failedHr = collectHr;
                    return collectHr;
                }
            }

            const size_t available   = FsS3::kMultipartMinPartSizeBytes - _buffer.size();
            const size_t appendBytes = (std::min)(available, static_cast<size_t>(remaining));
            _buffer.insert(_buffer.end(), src, src + appendBytes);
            src += appendBytes;
            remaining -= static_cast<unsigned long>(appendBytes);
            _position += appendBytes;
            *bytesWritten += static_cast<unsigned long>(appendBytes);

            const HRESULT flushHr = FlushBufferedParts();
            if (FAILED(flushHr))
            {
                _failedHr = flushHr;
                return flushHr;
            }
        }

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

        HRESULT hr = S_OK;

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
            hr = CollectAsyncPart();
            if (FAILED(hr))
            {
                const HRESULT abortHr = AbortMultipartUpload();
                _failedHr             = FAILED(abortHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : hr;
                return _failedHr;
            }

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

#if defined(_DEBUG)
            const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
            hr = (debugTransport != nullptr && debugTransport->complete != nullptr) ? debugTransport->complete(debugTransport->cookie, _parts)
                                                                                    : FsS3::CompleteS3MultipartUpload(*_owner.get(), _session, _parts);
#else
            hr = FsS3::CompleteS3MultipartUpload(*_owner.get(), _session, _parts);
#endif
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
        ReleaseBufferBudget();
        return S_OK;
    }

private:
    ~MultipartS3FileWriter()
    {
        static_cast<void>(CollectAsyncPart());
        if (! _committed)
        {
            const HRESULT hr = AbortMultipartUpload();
            if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_INVALID_STATE))
            {
                Debug::Warning(L"S3: multipart upload cleanup failed path='{}' hr=0x{:08X}.", _pluginPath, static_cast<unsigned long>(hr));
            }
        }
        ReleaseBufferBudget();
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

#if defined(_DEBUG)
        const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
        HRESULT hr                                          = (debugTransport != nullptr && debugTransport->begin != nullptr)
                                                                  ? debugTransport->begin(debugTransport->cookie, _session)
                                                                  : FsS3::BeginS3MultipartUpload(*_owner.get(), _bucketCtx, _bucket, _key, _session);
#else
        HRESULT hr = FsS3::BeginS3MultipartUpload(*_owner.get(), _bucketCtx, _bucket, _key, _session);
#endif
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
#if defined(_DEBUG)
        const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
        HRESULT hr = (debugTransport != nullptr && debugTransport->upload != nullptr)
                         ? debugTransport->upload(debugTransport->cookie, _nextPartNumber, sizeBytes, etag)
                         : FsS3::UploadS3MultipartPartFromMemory(*_owner.get(), _session, _nextPartNumber, data, sizeBytes, etag);
#else
        HRESULT hr = FsS3::UploadS3MultipartPartFromMemory(*_owner.get(), _session, _nextPartNumber, data, sizeBytes, etag);
#endif
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

    HRESULT CollectAsyncPart() noexcept
    {
        if (! _uploadThread.joinable())
        {
            return S_OK;
        }

        _uploadThread.join();
        std::vector<std::byte>().swap(_inFlightData);
        if (FAILED(_inFlightHr))
        {
            return _inFlightHr;
        }

        FsS3::S3MultipartUploadedPart part{};
        part.partNumber = _inFlightPartNumber;
        part.eTag       = std::move(_inFlightEtag);
        _parts.push_back(std::move(part));
        ++_nextPartNumber;
        _inFlightPartNumber = 0;
        return S_OK;
    }

    HRESULT ScheduleAsyncPart() noexcept
    {
        if (_uploadThread.joinable() || _buffer.size() < FsS3::kMultipartMinPartSizeBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        HRESULT hr = EnsureMultipartSession();
        if (FAILED(hr))
        {
            if (ShouldDiscoverS3SizeFromRange(hr))
            {
                Debug::Warning(L"S3: HeadObject was denied for bucket='{}' key='{}'; discovering size from ranged GetObject.",
                               FsS3::Utf16FromUtf8(_bucket),
                               FsS3::Utf16FromUtf8(_key));
                return S_OK;
            }
            return hr;
        }
        if (_nextPartNumber > 10000)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }

        _inFlightData = std::move(_buffer);
        _buffer.clear();

        _inFlightPartNumber = _nextPartNumber;
        _inFlightHr         = S_OK;
        _inFlightEtag.clear();
        try
        {
            _uploadThread = std::jthread([this](std::stop_token) noexcept
            {
#if defined(_DEBUG)
                const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
                _inFlightHr = (debugTransport != nullptr && debugTransport->upload != nullptr)
                                  ? debugTransport->upload(debugTransport->cookie, _inFlightPartNumber, _inFlightData.size(), _inFlightEtag)
                                  : FsS3::UploadS3MultipartPartFromMemory(
                                        *_owner.get(), _session, _inFlightPartNumber, _inFlightData.data(), _inFlightData.size(), _inFlightEtag);
#else
                _inFlightHr = FsS3::UploadS3MultipartPartFromMemory(
                    *_owner.get(), _session, _inFlightPartNumber, _inFlightData.data(), _inFlightData.size(), _inFlightEtag);
#endif
            });
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::system_error& error)
        {
            // std::jthread construction is the ABI boundary here; fail the upload cleanly when
            // Windows cannot create the worker instead of unwinding through IFileWriter::Write.
            Debug::Error(L"S3: unable to start asynchronous multipart upload worker. error={}", error.code().value());
            _inFlightData.clear();
            _inFlightPartNumber = 0;
            return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        }
        return S_OK;
    }

    HRESULT FlushBufferedParts() noexcept
    {
        if (_buffer.size() < FsS3::kMultipartMinPartSizeBytes)
        {
            return S_OK;
        }
        if (_buffer.size() != FsS3::kMultipartMinPartSizeBytes || _uploadThread.joinable())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        return ScheduleAsyncPart();
    }

    void ReleaseBufferBudget() noexcept
    {
        if (! _hasBufferBudget)
        {
            return;
        }

        std::vector<std::byte>().swap(_buffer);
        std::vector<std::byte>().swap(_inFlightData);
        _hasBufferBudget = false;
        GetMultipartWriterBufferBudget().release();
    }

    HRESULT AbortMultipartUpload() noexcept
    {
        if (! _hasSession || ! _owner)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

#if defined(_DEBUG)
        const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
        void* debugCookie                                   = debugTransport != nullptr ? debugTransport->cookie : nullptr;
        HRESULT (*debugAbort)(void*) noexcept               = debugTransport != nullptr ? debugTransport->abort : nullptr;
        const HRESULT hr = (debugTransport != nullptr && debugTransport->abort != nullptr) ? debugTransport->abort(debugTransport->cookie)
                                                                                           : FsS3::AbortS3MultipartUpload(*_owner.get(), _session);
#else
        const HRESULT hr = FsS3::AbortS3MultipartUpload(*_owner.get(), _session);
#endif
        if (FAILED(hr))
        {
            Debug::Warning(L"S3: queued failed multipart-abort cleanup path='{}' hr=0x{:08X}.", _pluginPath, static_cast<unsigned long>(hr));
            MultipartAbortQueue().Queue(_owner.get(),
                                        std::move(_session),
                                        _pluginPath
#if defined(_DEBUG)
                                        ,
                                        debugCookie,
                                        debugAbort
#endif
            );
        }
        else
        {
            _session = {};
        }
        _hasSession = false;
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
    std::vector<std::byte> _inFlightData;
    std::vector<FsS3::S3MultipartUploadedPart> _parts;
    std::jthread _uploadThread;
    FsS3::S3MultipartUploadSession _session{};
    std::string _inFlightEtag;
    uint64_t _position      = 0;
    HRESULT _failedHr       = S_OK;
    HRESULT _inFlightHr     = S_OK;
    int _nextPartNumber     = 1;
    int _inFlightPartNumber = 0;
    bool _allowOverwrite    = false;
    bool _committed         = false;
    bool _hasSession        = false;
    bool _hasBufferBudget   = false;
};
} // namespace

void FsS3::SchedulePendingMultipartAbortCleanup() noexcept
{
    MultipartAbortQueue().Schedule();
}

bool FsS3::CanUnloadPendingMultipartAbortCleanup() noexcept
{
    return MultipartAbortQueue().IsDrained();
}

#if defined(_DEBUG)
bool FsS3::WaitForPendingMultipartAbortCleanupForTest(unsigned long timeoutMs) noexcept
{
    return MultipartAbortQueue().WaitUntilDrained(timeoutMs);
}

void FsS3::RunDebugMultipartWriterContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept
    {
        if (condition)
        {
            ++passed;
            return;
        }

        ++failed;
        Debug::Error(L"FileSystemS3 multipart-writer selftest failed: {}", message);
    };

    wil::com_ptr<FileSystemS3> owner;
    owner.attach(new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3, nullptr));
    check(static_cast<bool>(owner), L"S3 filesystem allocation should succeed");
    if (! owner)
    {
        return;
    }

    wil::com_ptr<IFileSystemAtomicWriter> atomicWriter;
    check(owner.try_query_to(atomicWriter.put()) && atomicWriter, L"S3 should expose the explicit atomic-writer capability");
    BOOL atomicSupported = FALSE;
    check(atomicWriter && SUCCEEDED(atomicWriter->SupportsAtomicWriterCommit(L"/bucket/final.bin", FILESYSTEM_FLAG_NONE, &atomicSupported)) &&
              atomicSupported == TRUE,
          L"S3 final-key multipart commit should be advertised as atomic");

    MultipartWriterDebugContext context{};
    const MultipartWriterDebugTransport transport{
        .cookie = &context,
        .begin  = [](void* cookie, FsS3::S3MultipartUploadSession& session) noexcept -> HRESULT
    {
        auto* value = static_cast<MultipartWriterDebugContext*>(cookie);
        ++value->beginCalls;
        session.bucket   = "bucket";
        session.key      = "final.bin";
        session.uploadId = "debug-upload";
        return S_OK;
    },
        .upload = [](void* cookie, int partNumber, size_t sizeBytes, std::string& eTag) noexcept -> HRESULT
    {
        auto* value = static_cast<MultipartWriterDebugContext*>(cookie);
        if (partNumber != 1 || sizeBytes != FsS3::kMultipartMinPartSizeBytes)
        {
            return E_INVALIDARG;
        }

        std::unique_lock lock(value->mutex);
        ++value->uploadCalls;
        value->uploadStarted = true;
        value->cv.notify_all();
        if (! value->cv.wait_for(lock, std::chrono::seconds(5), [&] noexcept { return value->allowUploadCompletion; }))
        {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }
        eTag = "debug-etag";
        return S_OK;
    },
        .complete = [](void* cookie, const std::vector<FsS3::S3MultipartUploadedPart>& parts) noexcept -> HRESULT
    {
        auto* value = static_cast<MultipartWriterDebugContext*>(cookie);
        ++value->completeCalls;
        return parts.size() == 1u && parts.front().partNumber == 1 && parts.front().eTag == "debug-etag" ? S_OK : E_INVALIDARG;
    },
        .abort = [](void* cookie) noexcept -> HRESULT
    {
        auto* value = static_cast<MultipartWriterDebugContext*>(cookie);
        std::lock_guard lock(value->mutex);
        ++value->abortCalls;
        const bool fail = value->abortFailuresRemaining != 0u;
        if (fail)
        {
            --value->abortFailuresRemaining;
        }
        value->cv.notify_all();
        if (fail)
        {
            return HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE);
        }
        return S_OK;
    },
    };
    const MultipartWriterDebugTransportScope transportScope(transport);

    wil::com_ptr<IFileWriter> writer;
    writer.attach(new (std::nothrow) MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "final.bin", L"/bucket/final.bin", true));
    check(static_cast<bool>(writer), L"multipart writer allocation should succeed");
    if (! writer)
    {
        return;
    }

    std::vector<std::byte> part(static_cast<size_t>(FsS3::kMultipartMinPartSizeBytes), std::byte{0x5a});
    unsigned long bytesWritten = 0u;
    const HRESULT writeHr      = writer->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten);
    check(SUCCEEDED(writeHr) && bytesWritten == part.size(), L"multipart Write should schedule a complete part without waiting for upload completion");

    bool observedInFlight = false;
    {
        std::unique_lock lock(context.mutex);
        observedInFlight              = context.cv.wait_for(lock, std::chrono::seconds(2), [&] noexcept { return context.uploadStarted; });
        context.allowUploadCompletion = true;
    }
    context.cv.notify_all();
    check(observedInFlight, L"multipart upload should overlap the caller after Write returns");

    const HRESULT commitHr = writer->Commit();
    check(SUCCEEDED(commitHr), L"multipart Commit should join the in-flight part and complete the upload");
    check(context.beginCalls == 1u && context.uploadCalls == 1u && context.completeCalls == 1u && context.abortCalls == 0u,
          L"multipart writer should use one begin, one overlapped part, one complete, and no abort");
    writer.reset();

    {
        std::lock_guard lock(context.mutex);
        context.uploadStarted          = false;
        context.allowUploadCompletion  = true;
        context.abortFailuresRemaining = 1u;
    }

    wil::com_ptr<IFileWriter> abandonedWriter;
    abandonedWriter.attach(new (std::nothrow)
                               MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "orphan.bin", L"/bucket/orphan.bin", true));
    check(static_cast<bool>(abandonedWriter), L"multipart cleanup test writer allocation should succeed");
    if (! abandonedWriter)
    {
        return;
    }

    bytesWritten                 = 0u;
    const HRESULT abandonedWrite = abandonedWriter->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten);
    check(SUCCEEDED(abandonedWrite) && bytesWritten == part.size(), L"multipart cleanup test should establish an upload session");
    abandonedWriter.reset();

    const bool cleanupDrained = FsS3::WaitForPendingMultipartAbortCleanupForTest(5000u);
    unsigned int abortCalls   = 0u;
    {
        std::lock_guard lock(context.mutex);
        abortCalls = context.abortCalls;
    }
    check(cleanupDrained && abortCalls == 2u, L"a failed destructor abort should be retried asynchronously and drained before plugin unload");
}
#endif

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

        const HRESULT initializationHr = impl->InitializationStatus();
        if (FAILED(initializationHr))
        {
            impl->Release();
            return initializationHr;
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

    hr = EnsureWritableS3Target(*this, bucketCtx, bucket, key, path, allowOverwrite);
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

HRESULT STDMETHODCALLTYPE FileSystemS3::SupportsAtomicWriterCommit(const wchar_t* path, [[maybe_unused]] FileSystemFlags flags, BOOL* supported) noexcept
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

    *supported = _mode == FileSystemS3Mode::S3 ? TRUE : FALSE;
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
