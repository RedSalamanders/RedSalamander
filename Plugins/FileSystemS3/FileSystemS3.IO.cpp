#include "FileSystemS3.Internal.h"
#include "ContentDigest.h"

#include <aws/s3-crt/model/GetObjectRequest.h>
#include <aws/s3-crt/model/HeadObjectRequest.h>
#include <aws/s3-crt/model/ListObjectsV2Request.h>

#include <algorithm>
#include <array>
#include <chrono>
#include <charconv>
#include <cstddef>
#include <condition_variable>
#include <cstring>
#include <deque>
#include <format>
#include <limits>
#include <optional>
#include <semaphore>
#include <string>
#include <system_error>
#include <thread>
#include <utility>
#include <vector>

namespace FsS3 = FileSystemS3Internal;

static HRESULT TryGetS3ObjectSummaryFromClient(Aws::S3Crt::S3CrtClient& client,
                                               const FsS3::ResolvedAwsContext& bucketCtx,
                                               std::string_view bucket,
                                               std::string_view key,
                                               uint64_t& outSizeBytes,
                                               __int64& outLastWriteTime,
                                               bool& outFound,
                                               Aws::String* outEtag = nullptr,
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

    FsS3::ArmS3RequestControl(req);
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
                                    bool& outFound,
                                    S3ObjectRevision* outRevision) noexcept
{
    const auto client = GetS3Client(fs, bucketCtx);
    Aws::String etag;
    Aws::String versionId;
    const HRESULT hr = TryGetS3ObjectSummaryFromClient(*client,
                                                       bucketCtx,
                                                       bucket,
                                                       key,
                                                       outSizeBytes,
                                                       outLastWriteTime,
                                                       outFound,
                                                       outRevision != nullptr ? &etag : nullptr,
                                                       outRevision != nullptr ? &versionId : nullptr);
    if (outRevision != nullptr)
    {
        outRevision->etag.assign(etag.c_str(), etag.size());
        outRevision->versionId.assign(versionId.c_str(), versionId.size());
    }
    return hr;
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
        const char* const begin = value.data();
        const char* const end   = begin + value.size();
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

#if defined(_DEBUG)
[[nodiscard]] bool ShouldDiscoverS3SizeFromRange(HRESULT headHr) noexcept
{
    return headHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
}
#endif
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
    check(ShouldDiscoverS3SizeFromRange(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)),
          L"S3 reader should defer a HEAD access denial to ranged GET size discovery");
    check(! ShouldDiscoverS3SizeFromRange(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)),
          L"S3 reader should not hide non-permission HEAD failures");
    check(ParseS3ContentRange("bytes 0-9/25", contentRange) == S_OK && contentRange.first == 0u && contentRange.last == 9u && contentRange.total == 25u,
          L"S3 reader should discover object size from a validated Content-Range");
    check(ParseS3ContentRange("bytes 10-9/25", contentRange) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
          L"S3 reader should reject a reversed Content-Range");
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

class S3RangedFileReader final : public IFileReader, public IFileReaderOperationControl
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
        if (riid == __uuidof(IFileReaderOperationControl))
        {
            *ppvObject = static_cast<IFileReaderOperationControl*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    // R0f-S3 (C8): CreateFileReader carries no FileSystemOptions, so the host hands the task's
    // operation control here before the first Read. Read establishes the per-thread options scope
    // from this pointer, which is what arms the CRT continue handler on the reader's own thread.
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

        // The host's control travels with this reader, not with the calling thread: scope it here so
        // every request below is armed and a cancel ends an in-flight GetObject, not only the gap
        // between two Reads.
        const FileSystemOptions* const operationOptions = _operationOptions.load(std::memory_order_acquire);
        FsS3::S3OperationOptionsScope operationScope(operationOptions);
        if (const HRESULT control = FileSystemCheckOperationControl(operationOptions); FAILED(control))
        {
            return control;
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
                if (const HRESULT control = FileSystemCheckOperationControl(operationOptions); FAILED(control))
                {
                    return control;
                }
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
    std::atomic<const FileSystemOptions*> _operationOptions{nullptr}; // R0f-S3 (C8), host-owned lifetime

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
        const HRESULT hr =
            TryGetS3ObjectSummaryFromClient(*_client, _bucketCtx, _bucket, _key, sizeBytes, lastWriteTime, found, &_etag, &_versionId);
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

        FsS3::ArmS3RequestControl(req);
        auto outcome = _client->GetObject(req);
        if (! outcome.IsSuccess())
        {
            const auto& err = outcome.GetError();
            const std::wstring details =
                std::format(L"bucket='{}' key='{}' range='{}'", FsS3::Utf16FromUtf8(_bucket), FsS3::Utf16FromUtf8(_key), FsS3::Utf16FromUtf8(range));
            FsS3::LogAwsFailure(L"S3", L"GetObject", _bucketCtx, err, details);
            return FsS3::HresultFromAwsError(err);
        }

        auto result                      = outcome.GetResultWithOwnership();
        const Aws::String responseEtag   = result.GetETag();
        const Aws::String responseVersionId = result.GetVersionId();
        const HRESULT versionHr = ObserveS3ReadVersion(_etag, _versionId, responseEtag, responseVersionId);
        if (FAILED(versionHr))
        {
            return versionHr;
        }
        const auto responseContentLength = result.GetContentLength();
        uint64_t expectedBytes           = maxBytes;
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
        Aws::IOStream& stream            = result.GetBody();

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

// R3-1: the object that occupied the destination when an overwrite-capable writer was created.
struct S3ReplaceOccupant final
{
    bool found             = false;
    uint64_t sizeBytes     = 0;
    __int64 lastWriteTime  = 0;
    FsS3::S3ObjectRevision revision;
};

[[nodiscard]] HRESULT EnsureWritableS3Target(
    FileSystemS3& owner,
    const FsS3::ResolvedAwsContext& bucketCtx,
    std::string_view bucket,
    std::string_view key,
    std::wstring_view pluginPath,
    bool allowOverwrite,
    S3ReplaceOccupant* occupantOut) noexcept
{
    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

    uint64_t existingSize     = 0;
    __int64 existingLastWrite = 0;
    bool found                = false;
    FsS3::S3ObjectRevision existingRevision;
    const HRESULT existsHr =
        FsS3::TryGetS3ObjectSummary(owner, bucketCtx, bucket, key, existingSize, existingLastWrite, found, occupantOut != nullptr ? &existingRevision : nullptr);
    if (FAILED(existsHr))
    {
        return existsHr;
    }
    if (found && ! allowOverwrite)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }
    if (occupantOut != nullptr)
    {
        occupantOut->found         = found;
        occupantOut->sizeBytes     = existingSize;
        occupantOut->lastWriteTime = existingLastWrite;
        occupantOut->revision      = std::move(existingRevision);
    }

    std::string trimmedKey(key);
    while (! trimmedKey.empty() && trimmedKey.back() == '/')
    {
        trimmedKey.pop_back();
    }

    std::vector<std::wstring> ancestorPluginPaths;
    std::wstring currentPluginPath = FsS3::NormalizePluginPath(pluginPath);
    const size_t ancestorCount = static_cast<size_t>(std::count(trimmedKey.begin(), trimmedKey.end(), '/'));
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

    FsS3::ArmS3RequestControl(req);
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

#if defined(ENABLE_TESTS)
struct MultipartWriterDebugTransport final
{
    void* cookie = nullptr;
    HRESULT (*begin)(void*, FsS3::S3MultipartUploadSession&) noexcept = nullptr;
    HRESULT (*put)(void*, const void*, size_t, bool) noexcept = nullptr;
    HRESULT (*upload)(void*, int, size_t, std::string&) noexcept = nullptr;
    HRESULT (*complete)(void*, const std::vector<FsS3::S3MultipartUploadedPart>&, bool) noexcept = nullptr;
    HRESULT (*abort)(void*) noexcept = nullptr;
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
    MultipartWriterDebugContext()  = default;
    ~MultipartWriterDebugContext() = default;
    MultipartWriterDebugContext(const MultipartWriterDebugContext&)            = delete;
    MultipartWriterDebugContext(MultipartWriterDebugContext&&)                 = delete;
    MultipartWriterDebugContext& operator=(const MultipartWriterDebugContext&) = delete;
    MultipartWriterDebugContext& operator=(MultipartWriterDebugContext&&)      = delete;

    std::mutex mutex;
    std::condition_variable cv;
    std::array<unsigned long, 8> partDelayMs{};
    std::vector<int> startedParts;
    std::vector<int> completedParts;
    std::vector<int> committedParts;
    int failingPart = 0;
    HRESULT failingPartHr = HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE);
    int secondaryFailingPart = 0;
    HRESULT secondaryFailingPartHr = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    bool failingPartObserved = false;
    unsigned long defaultPartDelayMs = 0u;
    unsigned int activeUploads = 0u;
    unsigned int activeUploadHighWater = 0u;
    unsigned int beginCalls = 0u;
    unsigned int putCalls = 0u;
    unsigned int uploadCalls = 0u;
    unsigned int completeCalls = 0u;
    unsigned int abortCalls = 0u;
    unsigned int abortFailuresRemaining = 0u;
    HRESULT abortFailureHr = HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE);
    bool publicationDestinationExists = false;
    bool publicationDestinationPreserved = false;
    bool putDestinationMustNotExist = false;
    bool completeDestinationMustNotExist = false;
};
#endif

struct PendingMultipartAbort final
{
    wil::com_ptr<FileSystemS3> owner;
    FsS3::S3MultipartUploadSession session;
    std::wstring pluginPath;
    unsigned int attemptCount = 0u;
#if defined(ENABLE_TESTS)
    void* debugCookie = nullptr;
    HRESULT (*debugAbort)(void*) noexcept = nullptr;
#endif
};

const int kMultipartAbortCleanupModuleAnchor = 0;
constexpr auto kMultipartAbortRetryDelay = std::chrono::seconds(1);
constexpr size_t kMultipartAbortMaxItemsPerPass = 16u;
using UniqueThreadpoolTimer = wil::unique_any<PTP_TIMER, decltype(&::CloseThreadpoolTimer), ::CloseThreadpoolTimer>;

class PendingMultipartAbortQueue final
{
public:
    PendingMultipartAbortQueue() = default;
    ~PendingMultipartAbortQueue() = default;
    PendingMultipartAbortQueue(const PendingMultipartAbortQueue&)            = delete;
    PendingMultipartAbortQueue(PendingMultipartAbortQueue&&)                 = delete;
    PendingMultipartAbortQueue& operator=(const PendingMultipartAbortQueue&) = delete;
    PendingMultipartAbortQueue& operator=(PendingMultipartAbortQueue&&)      = delete;

    void Queue(FileSystemS3* owner,
               FsS3::S3MultipartUploadSession session,
               std::wstring_view pluginPath
#if defined(ENABLE_TESTS)
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
#if defined(ENABLE_TESTS)
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
        std::chrono::steady_clock::time_point timerDue{};
        bool armTimer = false;
        {
            std::lock_guard lock(_mutex);
            if (_workerScheduled || _timerScheduled || _pending.empty())
            {
                return;
            }
            if (const auto now = std::chrono::steady_clock::now(); now < _nextAttempt)
            {
                _timerScheduled = true;
                timerDue = _nextAttempt;
                armTimer = true;
            }
            else
            {
                _workerScheduled = true;
            }
        }

        if (armTimer)
        {
            ArmRetryTimer(timerDue);
            return;
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
        return _pending.empty() && ! _workerScheduled && ! _timerScheduled;
    }

#if defined(ENABLE_TESTS)
    [[nodiscard]] bool WaitUntilDrained(unsigned long timeoutMs) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(timeoutMs);
        while (true)
        {
            std::unique_lock lock(_mutex);
            if (_pending.empty() && ! _workerScheduled && ! _timerScheduled)
            {
                return true;
            }
            if (_changed.wait_until(lock, deadline) == std::cv_status::timeout)
            {
                return _pending.empty() && ! _workerScheduled && ! _timerScheduled;
            }
        }
    }
#endif

private:
    struct WorkItem final
    {
        WorkItem() = default;
        ~WorkItem() = default;
        WorkItem(const WorkItem&)            = delete;
        WorkItem(WorkItem&&)                 = delete;
        WorkItem& operator=(const WorkItem&) = delete;
        WorkItem& operator=(WorkItem&&)      = delete;

        PendingMultipartAbortQueue* queue = nullptr;
        wil::unique_hmodule moduleKeepAlive;
    };

    struct TimerItem final
    {
        TimerItem() = default;
        ~TimerItem() = default;
        TimerItem(const TimerItem&) = delete;
        TimerItem(TimerItem&&) = delete;
        TimerItem& operator=(const TimerItem&) = delete;
        TimerItem& operator=(TimerItem&&) = delete;

        PendingMultipartAbortQueue* queue = nullptr;
        wil::unique_hmodule moduleKeepAlive;
        UniqueThreadpoolTimer timer;
    };

    void ArmRetryTimer(std::chrono::steady_clock::time_point due) noexcept
    {
        auto timerItem = std::make_unique<TimerItem>();
        timerItem->queue = this;
        timerItem->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kMultipartAbortCleanupModuleAnchor);
        if (! timerItem->moduleKeepAlive)
        {
            MarkTimerSubmissionFailed();
            Debug::Error(L"S3: unable to pin the plugin module for delayed multipart-abort cleanup.");
            return;
        }
        timerItem->timer.reset(CreateThreadpoolTimer(
            [](PTP_CALLBACK_INSTANCE instance, void* context, PTP_TIMER) noexcept
            {
                std::unique_ptr<TimerItem> timer(static_cast<TimerItem*>(context));
                if (! timer)
                {
                    return;
                }
                TransferModulePinToCallbackReturn(instance, timer->moduleKeepAlive);
                PendingMultipartAbortQueue* queue = timer->queue;
                {
                    std::lock_guard lock(queue->_mutex);
                    queue->_timerScheduled = false;
                }
                queue->_changed.notify_all();
                queue->Schedule();
            },
            timerItem.get(),
            nullptr));
        if (! timerItem->timer)
        {
            MarkTimerSubmissionFailed();
            Debug::ErrorWithLastError(L"S3: unable to create delayed multipart-abort cleanup timer.");
            return;
        }

        const auto delay = due > std::chrono::steady_clock::now() ? due - std::chrono::steady_clock::now()
                                                                   : std::chrono::steady_clock::duration::zero();
        const uint64_t delayMs = static_cast<uint64_t>(
            std::max<int64_t>(1, std::chrono::duration_cast<std::chrono::milliseconds>(delay).count()));
        LARGE_INTEGER relativeDue{};
        const uint64_t maximumDelayMs = static_cast<uint64_t>((std::numeric_limits<LONGLONG>::max)() / 10'000ll);
        relativeDue.QuadPart = -static_cast<LONGLONG>(std::min(delayMs, maximumDelayMs) * 10'000u);
        FILETIME dueFileTime{};
        dueFileTime.dwLowDateTime = relativeDue.LowPart;
        dueFileTime.dwHighDateTime = static_cast<DWORD>(relativeDue.HighPart);
        TimerItem* rawTimer = timerItem.release();
        SetThreadpoolTimer(rawTimer->timer.get(), &dueFileTime, 0u, 0u);
    }

    [[nodiscard]] static HRESULT Abort(const PendingMultipartAbort& pending) noexcept
    {
#if defined(ENABLE_TESTS)
        if (pending.debugAbort != nullptr)
        {
            return pending.debugAbort(pending.debugCookie);
        }
#endif
        return pending.owner ? FsS3::AbortS3MultipartUpload(*pending.owner.get(), pending.session)
                             : HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
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
            if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
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
            _nextAttempt = retryNeeded ? std::chrono::steady_clock::now() + kMultipartAbortRetryDelay : std::chrono::steady_clock::now();
        }
        _changed.notify_all();
        Schedule();
    }

    void MarkSubmissionFailed() noexcept
    {
        {
            std::lock_guard lock(_mutex);
            _workerScheduled = false;
            _nextAttempt     = std::chrono::steady_clock::now() + kMultipartAbortRetryDelay;
        }
        _changed.notify_all();
        Schedule();
    }

    void MarkTimerSubmissionFailed() noexcept
    {
        {
            std::lock_guard lock(_mutex);
            _timerScheduled = false;
            // Timer allocation is an optimization, not a cleanup dependency. Fall back to
            // immediate worker submission so a queued abort cannot become wakeup-dependent.
            _nextAttempt = std::chrono::steady_clock::now();
        }
        _changed.notify_all();
        Schedule();
    }

    std::mutex _mutex;
    std::condition_variable _changed;
    std::deque<std::unique_ptr<PendingMultipartAbort>> _pending;
    std::chrono::steady_clock::time_point _nextAttempt{};
    bool _workerScheduled = false;
    bool _timerScheduled = false;
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
constexpr size_t kMultipartWriterMaxInFlightParts = 2u;

std::counting_semaphore<kMultipartWriterBufferSlots>& GetMultipartWriterBufferBudget() noexcept
{
    static std::counting_semaphore<kMultipartWriterBufferSlots> budget(kMultipartWriterBufferSlots);
    return budget;
}

std::atomic<uint64_t> g_multipartWriterBuffersInUse{0u};
std::atomic<uint64_t> g_multipartWriterBufferPeak{0u};

// FileSystemS3-local telemetry helper: these counters are observation-only, deliberately use relaxed ordering,
// and are covered by the multipart focused selftest rather than serving as a synchronization primitive.
void UpdateRelaxedMultipartTelemetryPeak(std::atomic<uint64_t>& peak, uint64_t candidate) noexcept
{
    uint64_t observed = peak.load(std::memory_order_relaxed);
    while (observed < candidate && ! peak.compare_exchange_weak(observed, candidate, std::memory_order_relaxed))
    {
    }
}

class MultipartWriterBufferLease final
{
public:
    MultipartWriterBufferLease() = default;
    ~MultipartWriterBufferLease() noexcept
    {
        Reset();
    }
    MultipartWriterBufferLease(const MultipartWriterBufferLease&)            = delete;
    MultipartWriterBufferLease& operator=(const MultipartWriterBufferLease&) = delete;
    MultipartWriterBufferLease(MultipartWriterBufferLease&& other) noexcept : _held(std::exchange(other._held, false))
    {
    }
    MultipartWriterBufferLease& operator=(MultipartWriterBufferLease&& other) noexcept
    {
        if (this != &other)
        {
            Reset();
            _held = std::exchange(other._held, false);
        }
        return *this;
    }

    [[nodiscard]] bool TryAcquire() noexcept
    {
        if (_held)
        {
            return true;
        }
        if (! GetMultipartWriterBufferBudget().try_acquire())
        {
            return false;
        }
        MarkAcquired();
        return true;
    }

    void Acquire() noexcept
    {
        if (_held)
        {
            return;
        }
        GetMultipartWriterBufferBudget().acquire();
        MarkAcquired();
    }

    void Reset() noexcept
    {
        if (! _held)
        {
            return;
        }
        _held = false;
        g_multipartWriterBuffersInUse.fetch_sub(1u, std::memory_order_relaxed);
        GetMultipartWriterBufferBudget().release();
    }

    [[nodiscard]] bool IsHeld() const noexcept
    {
        return _held;
    }

private:
    void MarkAcquired() noexcept
    {
        _held                 = true;
        const uint64_t inUse = g_multipartWriterBuffersInUse.fetch_add(1u, std::memory_order_relaxed) + 1u;
        UpdateRelaxedMultipartTelemetryPeak(g_multipartWriterBufferPeak, inUse);
    }

    bool _held = false;
};

#if defined(ENABLE_TESTS)
std::atomic<size_t> g_multipartWriterMaxInFlightOverride{0u};

class MultipartWriterMaxInFlightScope final
{
public:
    explicit MultipartWriterMaxInFlightScope(size_t maxInFlight) noexcept
        : _previous(g_multipartWriterMaxInFlightOverride.exchange(maxInFlight, std::memory_order_acq_rel))
    {
    }
    ~MultipartWriterMaxInFlightScope() noexcept
    {
        g_multipartWriterMaxInFlightOverride.store(_previous, std::memory_order_release);
    }
    MultipartWriterMaxInFlightScope(const MultipartWriterMaxInFlightScope&)            = delete;
    MultipartWriterMaxInFlightScope& operator=(const MultipartWriterMaxInFlightScope&) = delete;
    MultipartWriterMaxInFlightScope(MultipartWriterMaxInFlightScope&&)                 = delete;
    MultipartWriterMaxInFlightScope& operator=(MultipartWriterMaxInFlightScope&&)      = delete;

private:
    size_t _previous = 0u;
};
#endif

[[nodiscard]] size_t ResolveMultipartWriterMaxInFlightParts() noexcept
{
#if defined(ENABLE_TESTS)
    const size_t overrideValue = g_multipartWriterMaxInFlightOverride.load(std::memory_order_acquire);
    if (overrideValue != 0u)
    {
        return std::clamp(overrideValue, size_t{1u}, kMultipartWriterMaxInFlightParts);
    }
#endif
    return kMultipartWriterMaxInFlightParts;
}

struct S3WriterUploadPlan final
{
    uint64_t partSizeBytes = FsS3::kMultipartMinPartSizeBytes;
    uint64_t partCount = 0u;
};

[[nodiscard]] HRESULT PlanS3WriterUpload(uint64_t expectedSize, S3WriterUploadPlan& plan) noexcept
{
    constexpr uint64_t maximumParts = 10'000u;
    constexpr uint64_t fixedPartSize = FsS3::kMultipartMinPartSizeBytes;
    plan = {.partSizeBytes = fixedPartSize,
            .partCount = expectedSize == 0u ? 0u : expectedSize / fixedPartSize + (expectedSize % fixedPartSize == 0u ? 0u : 1u)};
    return plan.partCount <= maximumParts ? S_OK : HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
}

class MultipartS3FileWriter final : public IFileWriter, public IFileWriterExpectedSize, public IFileWriterCommitSizeProof, public IFileWriterExpectedReplacement, public IFileWriterContentProof
{
public:
    MultipartS3FileWriter(FileSystemS3* owner,
                          FsS3::ResolvedAwsContext bucketCtx,
                          std::string bucket,
                          std::string key,
                          std::wstring pluginPath,
                          bool allowOverwrite,
                          std::optional<S3ReplaceOccupant> replaceOccupant = std::nullopt) noexcept
        : _bucketCtx(std::move(bucketCtx)),
          _bucket(std::move(bucket)),
          _key(std::move(key)),
          _pluginPath(std::move(pluginPath)),
          _maxInFlightParts(ResolveMultipartWriterMaxInFlightParts()),
          _allowOverwrite(allowOverwrite),
          _replaceOccupant(std::move(replaceOccupant))
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
        if (riid == __uuidof(IFileWriterCommitSizeProof))
        {
            *ppvObject = static_cast<IFileWriterCommitSizeProof*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileWriterExpectedSize))
        {
            *ppvObject = static_cast<IFileWriterExpectedSize*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileWriterExpectedReplacement))
        {
            *ppvObject = static_cast<IFileWriterExpectedReplacement*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(IFileWriterContentProof))
        {
            *ppvObject = static_cast<IFileWriterContentProof*>(this);
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
        if (_expectedSize.has_value() &&
            (static_cast<uint64_t>(bytesToWrite) > _expectedSize.value() - std::min(_expectedSize.value(), _position)))
        {
            return RememberFailure(HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE));
        }
        EnsureMetricsStarted();

        const auto* src = static_cast<const std::byte*>(buffer);
        unsigned long remaining = bytesToWrite;
        while (remaining > 0u)
        {
            const HRESULT observedFailure = _firstWorkerFailure.load(std::memory_order_acquire);
            if (FAILED(observedFailure))
            {
                static_cast<void>(DrainPendingParts());
                return RememberFailure(observedFailure);
            }

            if (_buffer.empty())
            {
                const HRESULT leaseHr = AcquireAssemblyBuffer();
                if (FAILED(leaseHr))
                {
                    return RememberFailure(leaseHr);
                }
            }

            const size_t available   = static_cast<size_t>(_partSizeBytes) - _buffer.size();
            const size_t appendBytes = (std::min)(available, static_cast<size_t>(remaining));
            _buffer.insert(_buffer.end(), src, src + appendBytes);
            src += appendBytes;
            remaining -= static_cast<unsigned long>(appendBytes);
            _position += appendBytes;
            *bytesWritten += static_cast<unsigned long>(appendBytes);

            const HRESULT flushHr = FlushBufferedParts();
            if (FAILED(flushHr))
            {
                static_cast<void>(DrainPendingParts());
                return RememberFailure(flushHr);
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
        EnsureMetricsStarted();

        if (_expectedSize.has_value() && _position != _expectedSize.value())
        {
            return FinishFailure(HRESULT_FROM_WIN32(ERROR_HANDLE_EOF));
        }

        if (! _owner)
        {
            return FinishFailure(HRESULT_FROM_WIN32(ERROR_INVALID_STATE));
        }

        HRESULT hr = DrainPendingParts();
        if (FAILED(hr))
        {
            return FinishFailure(hr);
        }
        const HRESULT observedFailure = _firstWorkerFailure.load(std::memory_order_acquire);
        if (FAILED(_failedHr) || FAILED(observedFailure))
        {
            return FinishFailure(FAILED(_failedHr) ? _failedHr : observedFailure);
        }

        if (! _hasSession)
        {
#if defined(ENABLE_TESTS)
            const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
            hr = (debugTransport != nullptr && debugTransport->put != nullptr)
                     ? debugTransport->put(debugTransport->cookie, _buffer.empty() ? nullptr : _buffer.data(), _buffer.size(), ! _allowOverwrite)
                     : FsS3::PutS3ObjectFromMemory(*_owner.get(),
                                                   _bucketCtx,
                                                   _bucket,
                                                   _key,
                                                   _buffer.empty() ? nullptr : _buffer.data(),
                                                   _buffer.size(),
                                                   ! _allowOverwrite,
                                                   ReplaceIfMatchEtag(),
                                                   &_committedRevision);
#else
            hr = FsS3::PutS3ObjectFromMemory(*_owner.get(),
                                             _bucketCtx,
                                             _bucket,
                                             _key,
                                             _buffer.empty() ? nullptr : _buffer.data(),
                                             _buffer.size(),
                                             ! _allowOverwrite,
                                             ReplaceIfMatchEtag(),
                                             &_committedRevision);
#endif
            if (FAILED(hr))
            {
                return FinishFailure(hr);
            }
        }
        else
        {
            if (! _buffer.empty())
            {
                hr = ScheduleBufferedPart(true);
                if (FAILED(hr))
                {
                    return FinishFailure(hr);
                }
            }

            hr = DrainPendingParts();
            if (FAILED(hr))
            {
                return FinishFailure(hr);
            }

            if (_parts.empty())
            {
                return FinishFailure(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
            }

            std::sort(_parts.begin(), _parts.end(), [](const auto& left, const auto& right) noexcept { return left.partNumber < right.partNumber; });

#if defined(ENABLE_TESTS)
            const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
            hr = (debugTransport != nullptr && debugTransport->complete != nullptr)
                     ? debugTransport->complete(debugTransport->cookie, _parts, ! _allowOverwrite)
                     : FsS3::CompleteS3MultipartUpload(*_owner.get(), _session, _parts, ! _allowOverwrite, &_committedRevision, ReplaceIfMatchEtag());
#else
            hr = FsS3::CompleteS3MultipartUpload(*_owner.get(), _session, _parts, ! _allowOverwrite, &_committedRevision, ReplaceIfMatchEtag());
#endif
            if (FAILED(hr))
            {
                return FinishFailure(hr);
            }

            _hasSession = false;
            _session    = {};
            _parts.clear();
        }

        _committed = true;
        _owner->NotifySyntheticPathCreated(_pluginPath);
        EmitMetrics(S_OK);
        ReleaseBufferStorage();
        return S_OK;
    }

    // R3-2: S3 returns the full-object CRC-64/NVME it computed for the stored object when the upload
    // declares that algorithm; the host compares it with the bytes it streamed.
    HRESULT STDMETHODCALLTYPE GetContentProofAlgorithms(uint32_t* algorithmMask) noexcept override
    {
        if (algorithmMask == nullptr)
        {
            return E_POINTER;
        }
        *algorithmMask = 1u << FILESYSTEM_CONTENT_PROOF_CRC64NVME_64;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetCommittedContentProof(FileSystemContentProof* proof) noexcept override
    {
        if (proof == nullptr)
        {
            return E_POINTER;
        }
        if (! _committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        std::vector<std::byte> digest;
        if (! Common::Crypto::DecodeBase64Digest(_committedRevision.crc64NvmeBase64, digest) || digest.size() != 8u)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        uint64_t committedBytes = 0u;
        RETURN_IF_FAILED(GetCommittedSize(&committedBytes));
        *proof                  = {};
        proof->sizeBytes        = sizeof(*proof);
        proof->algorithm        = FILESYSTEM_CONTENT_PROOF_CRC64NVME_64;
        proof->contentSizeBytes = committedBytes;
        std::memcpy(proof->digest, digest.data(), digest.size());
        return S_OK;
    }

    // R3-1: the occupant the host showed the user. It must be the object present when the writer
    // opened and unchanged since; publication then carries If-Match with that object's ETag.
    HRESULT STDMETHODCALLTYPE SetExpectedReplacement(const FileSystemBasicInformation* expected) noexcept override
    {
        if (expected == nullptr || expected->sizeBytes < sizeof(FileSystemBasicInformation))
        {
            return E_INVALIDARG;
        }
        if (! _allowOverwrite || _committed || _position != 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        if (! _replaceOccupant.has_value() || ! _replaceOccupant->found)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        // R0-RC4: a zero timestamp on either side leaves no identity to compare; refuse rather than
        // replace whichever object the writer observed when it opened.
        if (expected->lastWriteTime == 0 || _replaceOccupant->lastWriteTime == 0 || expected->lastWriteTime != _replaceOccupant->lastWriteTime)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        _replaceConditional = true;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetCommittedSize(uint64_t* sizeBytes) noexcept override
    {
        if (! sizeBytes)
        {
            return E_POINTER;
        }

        *sizeBytes = 0u;
        if (! _committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        *sizeBytes = _position;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetExpectedSize(uint64_t sizeBytes) noexcept override
    {
        if (_position != 0u || _hasSession || ! _buffer.empty() || ! _pendingParts.empty() || _committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        S3WriterUploadPlan plan{};
        const HRESULT planHr = PlanS3WriterUpload(sizeBytes, plan);
        if (FAILED(planHr))
        {
            return planHr;
        }
        _partSizeBytes = plan.partSizeBytes;
        _expectedSize = sizeBytes;
        return S_OK;
    }

#if defined(ENABLE_TESTS)
    [[nodiscard]] size_t DebugAssemblyBufferCapacity() const noexcept
    {
        return _buffer.capacity();
    }
#endif

private:
    struct PendingPart final
    {
        PendingPart() = default;
        ~PendingPart() = default;
        PendingPart(const PendingPart&)            = delete;
        PendingPart& operator=(const PendingPart&) = delete;
        PendingPart(PendingPart&&)                 = delete;
        PendingPart& operator=(PendingPart&&)      = delete;

        MultipartWriterBufferLease bufferLease;
        std::vector<std::byte> data;
        std::jthread uploadThread;
        std::string eTag;
        uint64_t uploadUs = 0u;
        HRESULT hr        = S_OK;
        int partNumber    = 0;
    };

    ~MultipartS3FileWriter()
    {
        const HRESULT drainHr = DrainPendingParts();
        if (FAILED(drainHr))
        {
            static_cast<void>(RememberFailure(drainHr));
        }
        if (! _committed)
        {
            if (_hasSession)
            {
                const HRESULT abortHr = AbortMultipartUpload();
                if (FAILED(abortHr))
                {
                    Debug::Warning(L"S3: multipart upload cleanup failed path='{}' hr=0x{:08X}.", _pluginPath, static_cast<unsigned long>(abortHr));
                }
            }
            const HRESULT finalHr = FAILED(_failedHr) ? _failedHr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
            EmitMetrics(finalHr);
        }
        ReleaseBufferStorage();
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

#if defined(ENABLE_TESTS)
        const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
        HRESULT hr = (debugTransport != nullptr && debugTransport->begin != nullptr)
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

    [[nodiscard]] HRESULT RememberFailure(HRESULT hr) noexcept
    {
        if (FAILED(hr) && SUCCEEDED(_failedHr))
        {
            _failedHr = hr;
        }
        return _failedHr;
    }

    void RecordWorkerFailure(HRESULT hr) noexcept
    {
        if (SUCCEEDED(hr))
        {
            return;
        }
        HRESULT expected = S_OK;
        static_cast<void>(_firstWorkerFailure.compare_exchange_strong(expected, hr, std::memory_order_acq_rel));
    }

    HRESULT CollectOldestPendingPart(bool capacityWait) noexcept
    {
        if (_pendingParts.empty())
        {
            return S_OK;
        }

        const auto waitStartedAt              = std::chrono::steady_clock::now();
        std::unique_ptr<PendingPart> pending = std::move(_pendingParts.front());
        _pendingParts.pop_front();
        if (pending->uploadThread.joinable())
        {
            pending->uploadThread.join();
        }
        if (capacityWait)
        {
            _partWindowWaitUs += Debug::Perf::ElapsedUs(waitStartedAt);
        }

        _partUploadUs += pending->uploadUs;
        const size_t payloadBytes = pending->data.size();
        if (FAILED(pending->hr))
        {
            const HRESULT firstWorkerFailure = _firstWorkerFailure.load(std::memory_order_acquire);
            return RememberFailure(FAILED(firstWorkerFailure) ? firstWorkerFailure : pending->hr);
        }

        FsS3::S3MultipartUploadedPart part{};
        part.partNumber = pending->partNumber;
        part.eTag       = std::move(pending->eTag);
        _parts.push_back(std::move(part));
        _uploadedPartBytes += static_cast<uint64_t>(payloadBytes);
        return S_OK;
    }

    HRESULT DrainPendingParts() noexcept
    {
        HRESULT firstFailure = S_OK;
        while (! _pendingParts.empty())
        {
            const HRESULT hr = CollectOldestPendingPart(false);
            if (FAILED(hr) && SUCCEEDED(firstFailure))
            {
                firstFailure = hr;
            }
        }
        const HRESULT workerFailure = _firstWorkerFailure.load(std::memory_order_acquire);
        return FAILED(firstFailure) ? firstFailure : workerFailure;
    }

    HRESULT AcquireAssemblyBuffer() noexcept
    {
        if (_bufferLease.IsHeld())
        {
            return S_OK;
        }

        if (_pendingParts.size() >= _maxInFlightParts)
        {
            const HRESULT collectHr = CollectOldestPendingPart(true);
            if (FAILED(collectHr))
            {
                return collectHr;
            }
        }

        const auto waitStartedAt = std::chrono::steady_clock::now();
        if (! _bufferLease.TryAcquire())
        {
            if (! _pendingParts.empty())
            {
                const HRESULT collectHr = CollectOldestPendingPart(true);
                if (FAILED(collectHr))
                {
                    return collectHr;
                }
            }
            _bufferLease.Acquire();
        }
        _bufferBudgetWaitUs += Debug::Perf::ElapsedUs(waitStartedAt);
        uint64_t reserveBytes = _partSizeBytes;
        if (_expectedSize.has_value())
        {
            reserveBytes = (std::min)(_partSizeBytes,
                                      _expectedSize.value() > _position ? _expectedSize.value() - _position : uint64_t{0u});
        }
        if (reserveBytes != 0u)
        {
            _buffer.reserve(static_cast<size_t>(reserveBytes));
        }
        return S_OK;
    }

    HRESULT ScheduleBufferedPart(bool allowPartial) noexcept
    {
        if (_buffer.empty() || (! allowPartial && _buffer.size() != _partSizeBytes) || ! _bufferLease.IsHeld())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        const HRESULT observedFailure = _firstWorkerFailure.load(std::memory_order_acquire);
        if (FAILED(observedFailure))
        {
            return observedFailure;
        }
        if (_pendingParts.size() >= _maxInFlightParts)
        {
            const HRESULT collectHr = CollectOldestPendingPart(true);
            if (FAILED(collectHr))
            {
                return collectHr;
            }
        }

        HRESULT hr = EnsureMultipartSession();
        if (FAILED(hr))
        {
            return hr;
        }
        if (_nextPartNumber > 10000)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }

        auto pending          = std::make_unique<PendingPart>();
        pending->bufferLease  = std::move(_bufferLease);
        pending->data         = std::move(_buffer);
        pending->partNumber   = _nextPartNumber;
        PendingPart* pendingPtr = pending.get();
        try
        {
            pending->uploadThread = std::jthread([this, pendingPtr](std::stop_token) noexcept
            {
                const uint64_t activeUploads = _activeUploads.fetch_add(1u, std::memory_order_acq_rel) + 1u;
                UpdateRelaxedMultipartTelemetryPeak(_inFlightHighWater, activeUploads);
                auto activeGuard = wil::scope_exit([&]() noexcept { _activeUploads.fetch_sub(1u, std::memory_order_acq_rel); });
                const auto uploadStartedAt = std::chrono::steady_clock::now();
#if defined(ENABLE_TESTS)
                const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
                pendingPtr->hr = (debugTransport != nullptr && debugTransport->upload != nullptr)
                                     ? debugTransport->upload(
                                           debugTransport->cookie, pendingPtr->partNumber, pendingPtr->data.size(), pendingPtr->eTag)
                                     : FsS3::UploadS3MultipartPartFromMemory(*_owner.get(),
                                                                            _session,
                                                                            pendingPtr->partNumber,
                                                                            pendingPtr->data.data(),
                                                                            pendingPtr->data.size(),
                                                                            pendingPtr->eTag);
#else
                pendingPtr->hr = FsS3::UploadS3MultipartPartFromMemory(*_owner.get(),
                                                                       _session,
                                                                       pendingPtr->partNumber,
                                                                       pendingPtr->data.data(),
                                                                       pendingPtr->data.size(),
                                                                       pendingPtr->eTag);
#endif
                pendingPtr->uploadUs = Debug::Perf::ElapsedUs(uploadStartedAt);
                RecordWorkerFailure(pendingPtr->hr);
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
            return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        }

        ++_nextPartNumber;
        ++_scheduledPartCount;
        _pendingParts.push_back(std::move(pending));
        return S_OK;
    }

    HRESULT FlushBufferedParts() noexcept
    {
        if (_buffer.size() < _partSizeBytes)
        {
            return S_OK;
        }
        if (_buffer.size() != _partSizeBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        return ScheduleBufferedPart(false);
    }

    [[nodiscard]] std::string_view ReplaceIfMatchEtag() const noexcept
    {
        return _replaceConditional && _replaceOccupant.has_value() ? std::string_view(_replaceOccupant->revision.etag) : std::string_view{};
    }

    HRESULT FinishFailure(HRESULT hr) noexcept
    {
        const HRESULT drainHr = DrainPendingParts();
        const HRESULT primaryHr = FAILED(hr) ? hr : drainHr;
        static_cast<void>(RememberFailure(primaryHr));
        if (_hasSession)
        {
            static_cast<void>(AbortMultipartUpload());
        }
        EmitMetrics(_failedHr);
        ReleaseBufferStorage();
        return _failedHr;
    }

    void ReleaseBufferStorage() noexcept
    {
        std::vector<std::byte>().swap(_buffer);
        _bufferLease.Reset();
    }

    HRESULT AbortMultipartUpload() noexcept
    {
        if (! _hasSession || ! _owner)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

#if defined(ENABLE_TESTS)
        const MultipartWriterDebugTransport* debugTransport = g_multipartWriterDebugTransport.load(std::memory_order_acquire);
        void* debugCookie                     = debugTransport != nullptr ? debugTransport->cookie : nullptr;
        HRESULT (*debugAbort)(void*) noexcept = debugTransport != nullptr ? debugTransport->abort : nullptr;
        const HRESULT hr = (debugTransport != nullptr && debugTransport->abort != nullptr) ? debugTransport->abort(debugTransport->cookie)
                                                                                           : FsS3::AbortS3MultipartUpload(*_owner.get(), _session);
#else
        const HRESULT hr = FsS3::AbortS3MultipartUpload(*_owner.get(), _session);
#endif
        if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            // A failed destructor abort should be retried asynchronously so cleanup reaches the unload quiet point.
            Debug::Warning(L"S3: queued failed multipart-abort cleanup path='{}' hr=0x{:08X}.", _pluginPath, static_cast<unsigned long>(hr));
            MultipartAbortQueue().Queue(_owner.get(),
                                        std::move(_session),
                                        _pluginPath
#if defined(ENABLE_TESTS)
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

    void EnsureMetricsStarted() noexcept
    {
        if (! _metricsStarted)
        {
            _metricsStarted = true;
            _metricsStartedAt = std::chrono::steady_clock::now();
        }
    }

    void EmitMetrics(HRESULT hr) noexcept
    {
        if (_metricsEmitted)
        {
            return;
        }
        _metricsEmitted = true;
        if (! Debug::Perf::IsCaptureEnabled())
        {
            return;
        }

        EnsureMetricsStarted();
        const uint64_t totalUs = Debug::Perf::ElapsedUs(_metricsStartedAt);
        const uint64_t highWater = _inFlightHighWater.load(std::memory_order_acquire);
        const std::wstring detail = std::format(L"path={} maxInFlight={} scheduledParts={} uploadedPartBytes={} bufferBudgetBytes={}",
                                                _pluginPath,
                                                _maxInFlightParts,
                                                _scheduledPartCount,
                                                _uploadedPartBytes,
                                                kMultipartWriterBufferBudgetBytes);
        Debug::Perf::Emit(L"FileOps.S3.Multipart.TotalUs", detail, totalUs, _position, _scheduledPartCount, hr);
        Debug::Perf::Emit(L"FileOps.S3.Multipart.InFlightHighWater", detail, 0u, highWater, _maxInFlightParts, hr);
        Debug::Perf::Emit(
            L"FileOps.S3.Multipart.CapacityWaitUs", detail, _bufferBudgetWaitUs + _partWindowWaitUs, _bufferBudgetWaitUs, _partWindowWaitUs, hr);
        Debug::Perf::Emit(L"FileOps.S3.Multipart.PartUploadUs", detail, _partUploadUs, _scheduledPartCount, _uploadedPartBytes, hr);
        if (FAILED(hr))
        {
            Debug::Perf::Emit(L"FileOps.S3.Multipart.FirstFailure", detail, 0u, static_cast<uint32_t>(hr), _scheduledPartCount, hr);
        }
    }

    std::atomic_ulong _refCount{1};
    wil::com_ptr<FileSystemS3> _owner;
    FsS3::ResolvedAwsContext _bucketCtx;
    std::string _bucket;
    std::string _key;
    std::wstring _pluginPath;
    MultipartWriterBufferLease _bufferLease;
    std::vector<std::byte> _buffer;
    std::vector<FsS3::S3MultipartUploadedPart> _parts;
    std::deque<std::unique_ptr<PendingPart>> _pendingParts;
    FsS3::S3MultipartUploadSession _session{};
    std::chrono::steady_clock::time_point _metricsStartedAt{};
    std::atomic<uint64_t> _activeUploads{0u};
    std::atomic<uint64_t> _inFlightHighWater{0u};
    std::atomic<HRESULT> _firstWorkerFailure{S_OK};
    uint64_t _position            = 0u;
    uint64_t _partSizeBytes       = FsS3::kMultipartMinPartSizeBytes;
    uint64_t _uploadedPartBytes   = 0u;
    uint64_t _bufferBudgetWaitUs  = 0u;
    uint64_t _partWindowWaitUs    = 0u;
    uint64_t _partUploadUs        = 0u;
    HRESULT _failedHr             = S_OK;
    std::optional<uint64_t> _expectedSize;
    size_t _scheduledPartCount    = 0u;
    size_t _maxInFlightParts      = kMultipartWriterMaxInFlightParts;
    int _nextPartNumber           = 1;
    bool _allowOverwrite = false;
    std::optional<S3ReplaceOccupant> _replaceOccupant; // R3-1: occupant at creation (overwrite writers only)
    bool _replaceConditional = false;                  // R3-1: the host granted a replacement of that occupant
    FsS3::S3ObjectRevision _committedRevision;         // R3-2: ETag/version/CRC-64 S3 reported at publication
    bool _committed      = false;
    bool _hasSession     = false;
    bool _metricsStarted = false;
    bool _metricsEmitted = false;
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

#if defined(ENABLE_TESTS)
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
        .begin = [](void* cookie, FsS3::S3MultipartUploadSession& session) noexcept -> HRESULT
        {
            auto* value = static_cast<MultipartWriterDebugContext*>(cookie);
            {
                std::lock_guard lock(value->mutex);
                ++value->beginCalls;
            }
            session.bucket   = "bucket";
            session.key      = "final.bin";
            session.uploadId = "debug-upload";
            return S_OK;
        },
        .put = [](void* cookie, const void* buffer, size_t sizeBytes, bool destinationMustNotExist) noexcept -> HRESULT
        {
            auto* value = static_cast<MultipartWriterDebugContext*>(cookie);
            std::lock_guard lock(value->mutex);
            ++value->putCalls;
            value->putDestinationMustNotExist = destinationMustNotExist;
            if (buffer == nullptr || sizeBytes != 1024u * 1024u)
            {
                return E_INVALIDARG;
            }
            if (value->publicationDestinationExists)
            {
                value->publicationDestinationPreserved = destinationMustNotExist;
                Aws::Client::AWSError<Aws::Client::CoreErrors> error;
                error.SetResponseCode(Aws::Http::HttpResponseCode::PRECONDITION_FAILED);
                return destinationMustNotExist ? FsS3::HresultFromS3ConditionalPublicationError(error, true) : S_OK;
            }
            return S_OK;
        },
        .upload = [](void* cookie, int partNumber, size_t sizeBytes, std::string& eTag) noexcept -> HRESULT
        {
            auto* value = static_cast<MultipartWriterDebugContext*>(cookie);
            if (partNumber <= 0 || sizeBytes != FsS3::kMultipartMinPartSizeBytes)
            {
                return E_INVALIDARG;
            }

            unsigned long delayMs = 0u;
            HRESULT configuredFailure = S_OK;
            {
                std::lock_guard lock(value->mutex);
                ++value->uploadCalls;
                ++value->activeUploads;
                value->activeUploadHighWater = (std::max)(value->activeUploadHighWater, value->activeUploads);
                value->startedParts.push_back(partNumber);
                delayMs = value->defaultPartDelayMs;
                if (static_cast<size_t>(partNumber) < value->partDelayMs.size() && value->partDelayMs[static_cast<size_t>(partNumber)] != 0u)
                {
                    delayMs = value->partDelayMs[static_cast<size_t>(partNumber)];
                }
                if (partNumber == value->failingPart)
                {
                    configuredFailure = value->failingPartHr;
                }
                else if (partNumber == value->secondaryFailingPart)
                {
                    configuredFailure = value->secondaryFailingPartHr;
                }
                value->cv.notify_all();
            }

            if (delayMs != 0u)
            {
                std::this_thread::sleep_for(std::chrono::milliseconds(delayMs));
            }

            {
                std::lock_guard lock(value->mutex);
                --value->activeUploads;
                value->completedParts.push_back(partNumber);
                if (FAILED(configuredFailure))
                {
                    value->failingPartObserved = true;
                }
                value->cv.notify_all();
            }
            if (FAILED(configuredFailure))
            {
                return configuredFailure;
            }
            eTag = std::format("debug-etag-{}", partNumber);
            return S_OK;
        },
        .complete = [](void* cookie, const std::vector<FsS3::S3MultipartUploadedPart>& parts, bool destinationMustNotExist) noexcept -> HRESULT
        {
            auto* value = static_cast<MultipartWriterDebugContext*>(cookie);
            std::lock_guard lock(value->mutex);
            ++value->completeCalls;
            value->completeDestinationMustNotExist = destinationMustNotExist;
            value->committedParts.clear();
            int expectedPart = 1;
            for (const auto& part : parts)
            {
                value->committedParts.push_back(part.partNumber);
                if (part.partNumber != expectedPart || part.eTag != std::format("debug-etag-{}", expectedPart))
                {
                    return E_INVALIDARG;
                }
                ++expectedPart;
            }
            if (value->publicationDestinationExists)
            {
                value->publicationDestinationPreserved = destinationMustNotExist;
                Aws::Client::AWSError<Aws::Client::CoreErrors> error;
                error.SetResponseCode(Aws::Http::HttpResponseCode::CONFLICT);
                return destinationMustNotExist ? FsS3::HresultFromS3ConditionalPublicationError(error, true) : S_OK;
            }
            return parts.empty() ? E_INVALIDARG : S_OK;
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
            return fail ? value->abortFailureHr : S_OK;
        },
    };
    const MultipartWriterDebugTransportScope transportScope(transport);

    const auto resetContext = [&](unsigned long defaultDelayMs = 0u) noexcept
    {
        std::lock_guard lock(context.mutex);
        context.partDelayMs.fill(0u);
        context.startedParts.clear();
        context.completedParts.clear();
        context.committedParts.clear();
        context.failingPart            = 0;
        context.failingPartHr          = HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE);
        context.secondaryFailingPart   = 0;
        context.secondaryFailingPartHr = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        context.failingPartObserved    = false;
        context.defaultPartDelayMs     = defaultDelayMs;
        context.activeUploads          = 0u;
        context.activeUploadHighWater = 0u;
        context.beginCalls             = 0u;
        context.putCalls               = 0u;
        context.uploadCalls            = 0u;
        context.completeCalls          = 0u;
        context.abortCalls             = 0u;
        context.abortFailuresRemaining = 0u;
        context.abortFailureHr         = HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE);
        context.publicationDestinationExists    = false;
        context.publicationDestinationPreserved = false;
        context.putDestinationMustNotExist       = false;
        context.completeDestinationMustNotExist  = false;
    };

    resetContext();
    {
        constexpr uint64_t maximumPlannedBytes = FsS3::kMultipartMinPartSizeBytes * 10'000u;
        S3WriterUploadPlan plan{};
        check(PlanS3WriterUpload(0u, plan) == S_OK && plan.partCount == 0u,
              L"S3 expected-size planner should model an empty upload without parts");
        check(PlanS3WriterUpload(maximumPlannedBytes, plan) == S_OK && plan.partCount == 10'000u &&
                  plan.partSizeBytes == FsS3::kMultipartMinPartSizeBytes,
              L"S3 expected-size planner should accept the fixed-memory 10,000-part boundary");
        check(PlanS3WriterUpload(maximumPlannedBytes + 1u, plan) == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE),
              L"S3 expected-size planner should reject the first unrepresentable byte");

        wil::com_ptr<IFileWriter> rejectedWriter;
        rejectedWriter.attach(new (std::nothrow)
                                  MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "too-large.bin", L"/bucket/too-large.bin", true));
        wil::com_ptr<IFileWriterExpectedSize> expectedSizeWriter;
        check(rejectedWriter && rejectedWriter.try_query_to(expectedSizeWriter.put()) && expectedSizeWriter,
              L"S3 writer should expose expected-size planning through QueryInterface");
        check(expectedSizeWriter && expectedSizeWriter->SetExpectedSize(maximumPlannedBytes + 1u) ==
                                        HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE),
              L"S3 writer should reject an unrepresentable upload before the first Write");
        check(context.beginCalls == 0u && context.putCalls == 0u && context.uploadCalls == 0u,
              L"rejected S3 expected sizes should perform zero network activity");
    }

    resetContext();
    {
        constexpr size_t kTinyPayloadBytes = 4u * 1024u;
        constexpr size_t kTinyWriterCount  = 4u;
        std::array<wil::com_ptr<IFileWriter>, kTinyWriterCount> writers;
        std::array<MultipartS3FileWriter*, kTinyWriterCount> implementations{};
        std::array<std::byte, kTinyPayloadBytes> payload{};
        size_t totalCapacity = 0u;
        bool prepared = true;
        for (size_t index = 0u; index < writers.size(); ++index)
        {
            auto* implementation = new (std::nothrow) MultipartS3FileWriter(
                owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "tiny.bin", std::format(L"/bucket/tiny-{}.bin", index), true);
            implementations[index] = implementation;
            writers[index].attach(implementation);
            wil::com_ptr<IFileWriterExpectedSize> expectedSizeWriter;
            unsigned long written = 0u;
            prepared = prepared && writers[index] && writers[index].try_query_to(expectedSizeWriter.put()) && expectedSizeWriter &&
                       SUCCEEDED(expectedSizeWriter->SetExpectedSize(payload.size())) &&
                       SUCCEEDED(writers[index]->Write(payload.data(), static_cast<unsigned long>(payload.size()), &written)) &&
                       written == payload.size();
            if (implementation != nullptr)
            {
                totalCapacity += implementation->DebugAssemblyBufferCapacity();
            }
        }
        check(prepared && totalCapacity <= kTinyWriterCount * kTinyPayloadBytes,
              L"four known 4-KiB S3 writers should reserve no more than 16 KiB of assembly buffers");
    }

    resetContext();
    {
        wil::com_ptr<IFileWriter> writer;
        writer.attach(new (std::nothrow)
                          MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "proof.bin", L"/bucket/proof.bin", true));
        check(static_cast<bool>(writer), L"committed-size proof writer allocation should succeed");
        wil::com_ptr<IFileWriterCommitSizeProof> proof;
        check(writer && writer.try_query_to(proof.put()) && proof, L"S3 writer should expose the committed-size proof contract");
        wil::com_ptr<IFileWriterExpectedSize> expectedSizeWriter;
        check(writer && writer.try_query_to(expectedSizeWriter.put()) && expectedSizeWriter,
              L"S3 writer should expose the expected-size contract for production bridge planning");
        uint64_t committedSize = 0u;
        check(proof && proof->GetCommittedSize(&committedSize) == HRESULT_FROM_WIN32(ERROR_INVALID_STATE),
              L"S3 committed-size proof should reject queries before Commit");

        std::vector<std::byte> payload(1024u * 1024u, std::byte{0x5a});
        check(expectedSizeWriter && SUCCEEDED(expectedSizeWriter->SetExpectedSize(payload.size())),
              L"S3 writer should accept a representable size before the first Write");
        unsigned long bytesWritten = 0u;
        HRESULT hr = writer ? writer->Write(payload.data(), static_cast<unsigned long>(payload.size()), &bytesWritten) : E_POINTER;
        if (SUCCEEDED(hr) && writer)
        {
            hr = writer->Commit();
        }
        check(SUCCEEDED(hr) && bytesWritten == payload.size(), L"S3 committed-size proof scenario should commit its payload");
        committedSize = 0u;
        check(proof && SUCCEEDED(proof->GetCommittedSize(&committedSize)) && committedSize == payload.size(),
              L"S3 committed-size proof should return the exact successful Commit byte count");
    }

    struct ScenarioResult final
    {
        uint64_t durationUs = 0u;
        uint64_t bufferPeak = 0u;
        unsigned int beginCalls = 0u;
        unsigned int uploadCalls = 0u;
        unsigned int completeCalls = 0u;
        unsigned int abortCalls = 0u;
        unsigned int activeUploadHighWater = 0u;
        HRESULT hr = E_FAIL;
    };

    std::vector<std::byte> part(static_cast<size_t>(FsS3::kMultipartMinPartSizeBytes), std::byte{0x5a});
    const auto runFourPartScenario = [&](size_t maxInFlight, std::wstring_view path) noexcept -> ScenarioResult
    {
        resetContext(150u);
        g_multipartWriterBufferPeak.store(g_multipartWriterBuffersInUse.load(std::memory_order_acquire), std::memory_order_release);
        ScenarioResult result{};
        MultipartWriterMaxInFlightScope maxInFlightScope(maxInFlight);
        wil::com_ptr<IFileWriter> writer;
        writer.attach(new (std::nothrow)
                          MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "final.bin", std::wstring(path), true));
        if (! writer)
        {
            result.hr = E_OUTOFMEMORY;
            return result;
        }

        const auto startedAt = std::chrono::steady_clock::now();
        for (unsigned int partIndex = 0u; partIndex < 4u; ++partIndex)
        {
            unsigned long bytesWritten = 0u;
            result.hr = writer->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten);
            if (FAILED(result.hr) || bytesWritten != part.size())
            {
                if (SUCCEEDED(result.hr))
                {
                    result.hr = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
                }
                break;
            }
        }
        if (SUCCEEDED(result.hr))
        {
            result.hr = writer->Commit();
        }
        result.durationUs = Debug::Perf::ElapsedUs(startedAt);
        writer.reset();
        result.bufferPeak = g_multipartWriterBufferPeak.load(std::memory_order_acquire);
        {
            std::lock_guard lock(context.mutex);
            result.beginCalls             = context.beginCalls;
            result.uploadCalls            = context.uploadCalls;
            result.completeCalls          = context.completeCalls;
            result.abortCalls             = context.abortCalls;
            result.activeUploadHighWater = context.activeUploadHighWater;
        }
        return result;
    };

    const ScenarioResult baseline = runFourPartScenario(1u, L"/bucket/baseline.bin");
    const ScenarioResult candidate = runFourPartScenario(2u, L"/bucket/candidate.bin");
    constexpr uint64_t kScenarioBytes = 4ull * static_cast<uint64_t>(FsS3::kMultipartMinPartSizeBytes);
    constexpr uint64_t kMaximumCandidatePercent = 70u;
    check(SUCCEEDED(baseline.hr) && baseline.beginCalls == 1u && baseline.uploadCalls == 4u && baseline.completeCalls == 1u &&
              baseline.abortCalls == 0u && baseline.activeUploadHighWater == 1u,
          L"single-worker baseline should commit four ordered multipart payloads with high-water one");
    check(SUCCEEDED(candidate.hr) && candidate.beginCalls == 1u && candidate.uploadCalls == 4u && candidate.completeCalls == 1u &&
              candidate.abortCalls == 0u && candidate.activeUploadHighWater >= 2u,
          L"bounded candidate should commit four ordered multipart payloads with at least two concurrent uploads");
    check(candidate.durationUs * 100u <= baseline.durationUs * kMaximumCandidatePercent,
          L"two-worker candidate should improve the fixed-latency four-part baseline by at least thirty percent");
    check(baseline.bufferPeak <= static_cast<uint64_t>(kMultipartWriterBufferSlots) &&
              candidate.bufferPeak <= static_cast<uint64_t>(kMultipartWriterBufferSlots) &&
              g_multipartWriterBuffersInUse.load(std::memory_order_acquire) == 0u,
          L"multipart scenarios should remain inside the four-payload process budget and release every permit");
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.S3.Multipart.Baseline", L"single-worker; four 64-MiB parts; fixed 150-ms transport", baseline.durationUs, kScenarioBytes, 4u, baseline.hr);
        Debug::Perf::Emit(L"FileOps.S3.Multipart.Candidate", L"two-worker; four 64-MiB parts; fixed 150-ms transport", candidate.durationUs, kScenarioBytes, 4u, candidate.hr);
        Debug::Perf::Emit(L"FileOps.S3.Multipart.Improvement",
                          L"candidate gate <= 70 percent of baseline",
                          baseline.durationUs - (std::min)(baseline.durationUs, candidate.durationUs),
                          baseline.durationUs,
                          candidate.durationUs,
                          candidate.hr);
        Debug::Perf::Emit(L"FileOps.S3.Multipart.BufferHighWater",
                          L"global four-payload semaphore budget",
                          0u,
                          candidate.bufferPeak,
                          kMultipartWriterBufferSlots,
                          candidate.hr);
    }

    resetContext();
    {
        std::lock_guard lock(context.mutex);
        context.partDelayMs[1] = 160u;
        context.partDelayMs[2] = 10u;
    }
    {
        MultipartWriterMaxInFlightScope maxInFlightScope(2u);
        wil::com_ptr<IFileWriter> writer;
        writer.attach(new (std::nothrow)
                          MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "ordered.bin", L"/bucket/ordered.bin", true));
        check(static_cast<bool>(writer), L"out-of-order completion writer allocation should succeed");
        if (writer)
        {
            HRESULT hr = S_OK;
            for (unsigned int partIndex = 0u; partIndex < 2u && SUCCEEDED(hr); ++partIndex)
            {
                unsigned long bytesWritten = 0u;
                hr = writer->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten);
                if (SUCCEEDED(hr) && bytesWritten != part.size())
                {
                    hr = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
                }
            }
            if (SUCCEEDED(hr))
            {
                hr = writer->Commit();
            }
            check(SUCCEEDED(hr), L"out-of-order worker completion should still commit successfully");
            writer.reset();
        }
    }
    {
        std::lock_guard lock(context.mutex);
        check(context.completedParts.size() == 2u && context.completedParts[0] == 2 && context.completedParts[1] == 1,
              L"deterministic transport should complete part two before part one");
        check(context.committedParts == std::vector<int>({1, 2}), L"CompleteMultipartUpload input should remain ordered by assigned part number");
    }

    resetContext();
    constexpr HRESULT kInjectedPartFailure = HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE);
    {
        std::lock_guard lock(context.mutex);
        context.partDelayMs[1] = 120u;
        context.partDelayMs[2] = 10u;
        context.failingPart    = 2;
        context.failingPartHr  = kInjectedPartFailure;
        context.secondaryFailingPart   = 1;
        context.secondaryFailingPartHr = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    }
    {
        MultipartWriterMaxInFlightScope maxInFlightScope(2u);
        wil::com_ptr<IFileWriter> writer;
        writer.attach(new (std::nothrow)
                          MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "failure.bin", L"/bucket/failure.bin", true));
        check(static_cast<bool>(writer), L"first-failure writer allocation should succeed");
        if (writer)
        {
            HRESULT hr = S_OK;
            for (unsigned int partIndex = 0u; partIndex < 2u && SUCCEEDED(hr); ++partIndex)
            {
                unsigned long bytesWritten = 0u;
                hr = writer->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten);
            }
            bool failureObserved = false;
            {
                std::unique_lock lock(context.mutex);
                failureObserved = context.cv.wait_for(lock, std::chrono::seconds(2), [&]() noexcept { return context.failingPartObserved; });
            }
            unsigned long bytesWritten = 0u;
            const HRESULT failureHr = writer->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten);
            check(failureObserved && failureHr == kInjectedPartFailure && bytesWritten == 0u,
                  L"the first worker failure should stop later scheduling and propagate its exact HRESULT");
            writer.reset();
        }
    }
    {
        std::lock_guard lock(context.mutex);
        check(context.uploadCalls == 2u && context.completeCalls == 0u && context.abortCalls == 1u && context.activeUploads == 0u,
              L"first worker failure should join both workers, abort exactly once, and never complete");
        check(g_multipartWriterBuffersInUse.load(std::memory_order_acquire) == 0u,
              L"first worker failure should release all multipart payload permits");
    }

    resetContext();
    std::vector<std::byte> smallObject(1024u * 1024u, std::byte{0x2a});
    {
        wil::com_ptr<IFileWriter> writer;
        writer.attach(new (std::nothrow)
                          MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "small.bin", L"/bucket/small.bin", true));
        check(static_cast<bool>(writer), L"small-object writer allocation should succeed");
        if (writer)
        {
            unsigned long bytesWritten = 0u;
            const HRESULT writeHr = writer->Write(smallObject.data(), static_cast<unsigned long>(smallObject.size()), &bytesWritten);
            const HRESULT commitHr = SUCCEEDED(writeHr) ? writer->Commit() : writeHr;
            check(SUCCEEDED(commitHr) && bytesWritten == smallObject.size(), L"small-object PutObject path should remain successful");
            writer.reset();
        }
    }
    {
        std::lock_guard lock(context.mutex);
        check(context.putCalls == 1u && context.beginCalls == 0u && context.uploadCalls == 0u && context.completeCalls == 0u && context.abortCalls == 0u,
              L"small-object writes should retain the one-request PutObject path");
    }

    resetContext();
    {
        std::lock_guard lock(context.mutex);
        context.publicationDestinationExists = true;
    }
    {
        wil::com_ptr<IFileWriter> writer;
        writer.attach(new (std::nothrow)
                          MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "small-race.bin", L"/bucket/small-race.bin", false));
        check(static_cast<bool>(writer), L"small-object no-overwrite race writer allocation should succeed");
        if (writer)
        {
            unsigned long bytesWritten = 0u;
            const HRESULT writeHr = writer->Write(smallObject.data(), static_cast<unsigned long>(smallObject.size()), &bytesWritten);
            const HRESULT commitHr = SUCCEEDED(writeHr) ? writer->Commit() : writeHr;
            check(commitHr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) && bytesWritten == smallObject.size(),
                  L"small-object no-overwrite publication should reject a destination injected after preflight");
        }
    }
    {
        std::lock_guard lock(context.mutex);
        check(context.putCalls == 1u && context.putDestinationMustNotExist && context.publicationDestinationPreserved,
              L"small-object no-overwrite publication should require absence and preserve the injected object");
    }

    resetContext();
    {
        std::lock_guard lock(context.mutex);
        context.publicationDestinationExists = true;
    }
    {
        wil::com_ptr<IFileWriter> writer;
        writer.attach(new (std::nothrow)
                          MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "small-overwrite.bin", L"/bucket/small-overwrite.bin", true));
        check(static_cast<bool>(writer), L"small-object overwrite-authorized race writer allocation should succeed");
        if (writer)
        {
            unsigned long bytesWritten = 0u;
            const HRESULT writeHr = writer->Write(smallObject.data(), static_cast<unsigned long>(smallObject.size()), &bytesWritten);
            const HRESULT commitHr = SUCCEEDED(writeHr) ? writer->Commit() : writeHr;
            check(SUCCEEDED(commitHr) && bytesWritten == smallObject.size(),
                  L"small-object overwrite-authorized publication should replace an injected destination");
        }
    }
    {
        std::lock_guard lock(context.mutex);
        check(context.putCalls == 1u && ! context.putDestinationMustNotExist && ! context.publicationDestinationPreserved,
              L"small-object overwrite-authorized publication should remain unconditional");
    }

    resetContext();
    {
        std::lock_guard lock(context.mutex);
        context.publicationDestinationExists = true;
    }
    {
        wil::com_ptr<IFileWriter> writer;
        writer.attach(new (std::nothrow)
                          MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "multipart-race.bin", L"/bucket/multipart-race.bin", false));
        check(static_cast<bool>(writer), L"multipart no-overwrite race writer allocation should succeed");
        if (writer)
        {
            unsigned long bytesWritten = 0u;
            const HRESULT writeHr = writer->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten);
            const HRESULT commitHr = SUCCEEDED(writeHr) ? writer->Commit() : writeHr;
            check(commitHr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) && bytesWritten == part.size(),
                  L"multipart no-overwrite completion should reject a destination injected after preflight");
        }
    }
    {
        std::lock_guard lock(context.mutex);
        check(context.completeCalls == 1u && context.completeDestinationMustNotExist && context.publicationDestinationPreserved && context.abortCalls == 1u,
              L"multipart no-overwrite completion should require absence, preserve the injected object, and abort the upload");
    }

    resetContext();
    {
        std::lock_guard lock(context.mutex);
        context.publicationDestinationExists = true;
    }
    {
        wil::com_ptr<IFileWriter> writer;
        writer.attach(new (std::nothrow)
                          MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "multipart-overwrite.bin", L"/bucket/multipart-overwrite.bin", true));
        check(static_cast<bool>(writer), L"multipart overwrite-authorized race writer allocation should succeed");
        if (writer)
        {
            unsigned long bytesWritten = 0u;
            const HRESULT writeHr = writer->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten);
            const HRESULT commitHr = SUCCEEDED(writeHr) ? writer->Commit() : writeHr;
            check(SUCCEEDED(commitHr) && bytesWritten == part.size(),
                  L"multipart overwrite-authorized completion should replace an injected destination");
        }
    }
    {
        std::lock_guard lock(context.mutex);
        check(context.completeCalls == 1u && ! context.completeDestinationMustNotExist && ! context.publicationDestinationPreserved && context.abortCalls == 0u,
              L"multipart overwrite-authorized completion should remain unconditional");
    }

    resetContext(10u);
    {
        std::lock_guard lock(context.mutex);
        context.abortFailuresRemaining = 1u;
    }
    wil::com_ptr<IFileWriter> abandonedWriter;
    abandonedWriter.attach(new (std::nothrow)
                               MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "orphan.bin", L"/bucket/orphan.bin", true));
    check(static_cast<bool>(abandonedWriter), L"multipart cleanup test writer allocation should succeed");
    if (abandonedWriter)
    {
        unsigned long bytesWritten = 0u;
        const HRESULT abandonedWrite = abandonedWriter->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten);
        check(SUCCEEDED(abandonedWrite) && bytesWritten == part.size(), L"multipart cleanup test should establish an upload session");
        abandonedWriter.reset();
    }

    const bool cleanupDrained = FsS3::WaitForPendingMultipartAbortCleanupForTest(5000u);
    unsigned int abortCalls   = 0u;
    unsigned int activeUploads = 0u;
    {
        std::lock_guard lock(context.mutex);
        abortCalls    = context.abortCalls;
        activeUploads = context.activeUploads;
    }
    check(cleanupDrained && abortCalls == 2u && activeUploads == 0u && g_multipartWriterBuffersInUse.load(std::memory_order_acquire) == 0u,
          L"a failed destructor abort should retry, drain workers and payloads, and reach the plugin unload quiet point");

    resetContext(10u);
    {
        std::lock_guard lock(context.mutex);
        context.abortFailuresRemaining = 1u;
        context.abortFailureHr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    abandonedWriter.attach(new (std::nothrow)
                               MultipartS3FileWriter(owner.get(), FsS3::ResolvedAwsContext{}, "bucket", "already-aborted.bin", L"/bucket/already-aborted.bin", true));
    check(static_cast<bool>(abandonedWriter), L"terminal abort-status writer allocation should succeed");
    if (abandonedWriter)
    {
        unsigned long bytesWritten = 0u;
        check(SUCCEEDED(abandonedWriter->Write(part.data(), static_cast<unsigned long>(part.size()), &bytesWritten)) &&
                  bytesWritten == part.size(),
              L"terminal abort-status test should establish an upload session");
        abandonedWriter.reset();
    }
    const bool terminalCleanupDrained = FsS3::WaitForPendingMultipartAbortCleanupForTest(5000u);
    {
        std::lock_guard lock(context.mutex);
        abortCalls = context.abortCalls;
    }
    check(terminalCleanupDrained && abortCalls == 1u,
          L"an already-absent multipart upload should be terminal success and must not retry");
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

#if defined(ENABLE_TESTS)
    if (_mode == FileSystemS3Mode::S3)
    {
        bool debugHandled = false;
        const HRESULT debugHr = FsS3::TryGetDebugS3Attributes(path, debugHandled, *fileAttributes);
        if (debugHandled)
        {
            return debugHr;
        }
    }
#endif

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

        FsS3::ArmS3RequestControl(req);
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

    wil::unique_hfile file;
    if (const HRESULT tempHr = Common::Files::CreateDeleteOnCloseTemporaryFile(FsS3::kS3TemporaryFileOptions, file); FAILED(tempHr))
    {
        return tempHr;
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

    S3ReplaceOccupant occupant{};
    hr = EnsureWritableS3Target(*this, bucketCtx, bucket, key, path, allowOverwrite, allowOverwrite ? &occupant : nullptr);
    if (FAILED(hr))
    {
        return hr;
    }
    std::optional<S3ReplaceOccupant> replaceOccupant;
    if (allowOverwrite)
    {
        replaceOccupant = std::move(occupant);
    }

    auto* impl = new (std::nothrow)
        MultipartS3FileWriter(this, std::move(bucketCtx), std::string(bucket), std::string(key), path, allowOverwrite, std::move(replaceOccupant));
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *writer = impl;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::SupportsAtomicWriterCommit(const wchar_t* path,
                                                                   [[maybe_unused]] FileSystemFlags flags,
                                                                   BOOL* supported) noexcept
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
