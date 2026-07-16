#include "FileSystemS3.Internal.h"
#include "HandleIo.h"

#include <aws/core/utils/StringUtils.h>
#include <aws/core/utils/memory/stl/AWSStringStream.h>
#include <aws/s3-crt/model/AbortMultipartUploadRequest.h>
#include <aws/s3-crt/model/BucketLocationConstraint.h>
#include <aws/s3-crt/model/CompleteMultipartUploadRequest.h>
#include <aws/s3-crt/model/CompletedMultipartUpload.h>
#include <aws/s3-crt/model/CompletedPart.h>
#include <aws/s3-crt/model/CopyObjectRequest.h>
#include <aws/s3-crt/model/CreateMultipartUploadRequest.h>
#include <aws/s3-crt/model/GetBucketLocationRequest.h>
#include <aws/s3-crt/model/GetObjectRequest.h>
#include <aws/s3-crt/model/ListBucketsRequest.h>
#include <aws/s3-crt/model/ListObjectsV2Request.h>
#include <aws/s3-crt/model/PutObjectRequest.h>
#include <aws/s3-crt/model/UploadPartCopyRequest.h>
#include <aws/s3-crt/model/UploadPartRequest.h>

#include <algorithm>
#include <unordered_set>

std::optional<std::string> LookupS3BucketRegion(FileSystemS3& fs, std::wstring_view bucketName) noexcept
{
    std::lock_guard lock(fs._stateMutex);
    if (const auto it = fs._s3BucketRegionByName.find(std::wstring(bucketName)); it != fs._s3BucketRegionByName.end())
    {
        return it->second;
    }
    return std::nullopt;
}

void SetS3BucketRegion(FileSystemS3& fs, std::wstring_view bucketName, std::string region) noexcept
{
    if (bucketName.empty() || region.empty())
    {
        return;
    }

    std::lock_guard lock(fs._stateMutex);
    fs._s3BucketRegionByName[std::wstring(bucketName)] = std::move(region);
}

namespace FileSystemS3Internal
{
namespace
{
[[nodiscard]] HRESULT PutS3ObjectFromHandle(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, HANDLE file, uint64_t sizeBytes) noexcept;

[[nodiscard]] std::string NormalizeBucketLocationRegion(const Aws::S3Crt::Model::BucketLocationConstraint value) noexcept
{
    using Aws::S3Crt::Model::BucketLocationConstraint;

    if (value == BucketLocationConstraint::NOT_SET || value == BucketLocationConstraint::us_east_1)
    {
        return "us-east-1";
    }

    if (value == BucketLocationConstraint::EU)
    {
        // Legacy alias for eu-west-1.
        return "eu-west-1";
    }

    const Aws::String name = Aws::S3Crt::Model::BucketLocationConstraintMapper::GetNameForBucketLocationConstraint(value);
    if (name.empty() || name == "NOT_SET")
    {
        return "us-east-1";
    }

    if (name == "EU")
    {
        return "eu-west-1";
    }

    return std::string(name.c_str(), name.size());
}

[[nodiscard]] HRESULT EnsureS3BucketRegion(FileSystemS3& fs, const ResolvedAwsContext& ctx, std::wstring_view bucketNameWide, std::string& outRegion) noexcept
{
    outRegion.clear();

    if (bucketNameWide.empty())
    {
        return E_INVALIDARG;
    }

    if (const auto cached = LookupS3BucketRegion(fs, bucketNameWide); cached.has_value())
    {
        outRegion = cached.value();
        return S_OK;
    }

    const std::string bucket = Utf8FromUtf16(bucketNameWide);
    if (bucket.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    const std::shared_ptr<Aws::S3Crt::S3CrtClient> client = GetS3Client(fs, ctx);
    Aws::S3Crt::Model::GetBucketLocationRequest req;
    req.SetBucket(bucket);

    const auto outcome = client->GetBucketLocation(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}'", bucketNameWide);
        LogAwsFailure(L"S3", L"GetBucketLocation", ctx, err, details);
        return HresultFromAwsError(err);
    }

    const auto region = NormalizeBucketLocationRegion(outcome.GetResult().GetLocationConstraint());
    if (region.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    outRegion = region;
    SetS3BucketRegion(fs, bucketNameWide, region);
    return S_OK;
}
} // namespace

[[nodiscard]] uint64_t ComputeMultipartPartSize(uint64_t sizeBytes) noexcept
{
    const uint64_t minPartSize = kMultipartMinPartSizeBytes;
    const uint64_t maxParts    = 10'000ull;
    const uint64_t roundedUp   = (sizeBytes == 0u) ? minPartSize : std::max<uint64_t>(minPartSize, (sizeBytes + maxParts - 1u) / maxParts);
    return roundedUp;
}

[[nodiscard]] bool IsSameAwsContextIdentity(const ResolvedAwsContext& left, const ResolvedAwsContext& right) noexcept
{
    return left.connectionName == right.connectionName && left.region == right.region && left.explicitRegion == right.explicitRegion &&
           left.endpointOverride == right.endpointOverride && left.useHttps == right.useHttps && left.verifyTls == right.verifyTls &&
           left.useVirtualAddressing == right.useVirtualAddressing && left.accessKeyId == right.accessKeyId && left.secretAccessKey == right.secretAccessKey;
}

[[nodiscard]] std::string UrlEncodeS3CopySource(std::string_view bucket, std::string_view key) noexcept
{
    std::string raw(bucket);
    raw.push_back('/');
    raw.append(key);

    const Aws::String encoded = Aws::Utils::StringUtils::URLEncode(Aws::String(raw.data(), raw.size()));
    return std::string(encoded.c_str(), encoded.size());
}

[[nodiscard]] HRESULT ListS3BucketsForConnection(FileSystemS3& fs, const ResolvedAwsContext& ctx, std::vector<FilesInformationS3::Entry>& out) noexcept
{
    const HRESULT hr = ListS3Buckets(fs, ctx, out);
    if (FAILED(hr) || ! ctx.explicitRegion.has_value())
    {
        return hr;
    }

    // Region filtering only makes sense when using AWS endpoints (custom endpoints may not support GetBucketLocation).
    if (! ctx.endpointOverride.empty())
    {
        return S_OK;
    }

    const std::string filterRegion = ctx.explicitRegion.value();
    if (filterRegion.empty())
    {
        return S_OK;
    }

    const std::wstring filterWide = Utf16FromUtf8(filterRegion);
    if (filterWide.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    std::vector<FilesInformationS3::Entry> filtered;
    filtered.reserve(out.size());

    for (auto& e : out)
    {
        if (e.name.empty())
        {
            continue;
        }

        std::string bucketRegion;
        const HRESULT regionHr = EnsureS3BucketRegion(fs, ctx, e.name, bucketRegion);
        if (FAILED(regionHr))
        {
            continue;
        }

        if (OrdinalString::EqualsNoCase(Utf16FromUtf8(bucketRegion), filterWide))
        {
            filtered.push_back(std::move(e));
        }
    }

    out = std::move(filtered);
    return S_OK;
}

[[nodiscard]] HRESULT ResolveS3ContextForBucket(FileSystemS3& fs, const ResolvedAwsContext& ctx, std::wstring_view bucketName, ResolvedAwsContext& out) noexcept
{
    out = ctx;

    if (! ctx.endpointOverride.empty() || ctx.explicitRegion.has_value())
    {
        return S_OK;
    }

    std::string bucketRegion;
    const HRESULT hr = EnsureS3BucketRegion(fs, ctx, bucketName, bucketRegion);
    if (FAILED(hr))
    {
        return hr;
    }

    out.region = std::move(bucketRegion);
    return S_OK;
}

[[nodiscard]] HRESULT ListS3Buckets(FileSystemS3& fs, const ResolvedAwsContext& ctx, std::vector<FilesInformationS3::Entry>& out) noexcept
{
    out.clear();

    const std::shared_ptr<Aws::S3Crt::S3CrtClient> client = GetS3Client(fs, ctx);
    Aws::S3Crt::Model::ListBucketsRequest req;
    const auto outcome = client->ListBuckets(req);
    if (! outcome.IsSuccess())
    {
        const auto& err = outcome.GetError();
        LogAwsFailure(L"S3", L"ListBuckets", ctx, err, L"buckets");
        return HresultFromAwsError(err);
    }

    const auto& buckets = outcome.GetResult().GetBuckets();
    out.reserve(buckets.size());

    for (const auto& bucket : buckets)
    {
        FilesInformationS3::Entry e{};
        e.name          = Utf16FromUtf8(bucket.GetName());
        e.attributes    = FILE_ATTRIBUTE_DIRECTORY;
        e.creationTime  = AwsDateTimeToFileTime64(bucket.GetCreationDate());
        e.lastWriteTime = e.creationTime;
        e.changeTime    = e.creationTime;
        out.push_back(std::move(e));
    }

    return S_OK;
}

[[nodiscard]] HRESULT ParseS3LocationForDirectory(std::wstring_view canonicalPath, S3Location& out) noexcept
{
    out = {};

    const std::wstring normalized = NormalizePluginPath(canonicalPath);
    if (normalized == L"/" || normalized.empty())
    {
        out.isRoot = true;
        return S_OK;
    }

    const auto segments = SplitPathSegments(normalized);
    if (segments.empty())
    {
        out.isRoot = true;
        return S_OK;
    }

    out.bucket = Utf8FromUtf16(segments[0]);
    if (out.bucket.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    if (segments.size() == 1)
    {
        out.keyOrPrefix.clear();
        return S_OK;
    }

    // For a directory listing, treat the remainder as a prefix ending in '/'.
    std::wstring suffix;
    for (size_t i = 1; i < segments.size(); ++i)
    {
        if (i > 1)
        {
            suffix.push_back(L'/');
        }
        suffix.append(segments[i]);
    }

    std::string prefix = Utf8FromUtf16(suffix);
    if (prefix.empty() && ! suffix.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }
    if (! prefix.empty() && prefix.back() != '/')
    {
        prefix.push_back('/');
    }
    out.keyOrPrefix = std::move(prefix);
    return S_OK;
}

[[nodiscard]] HRESULT ListS3Objects(FileSystemS3& fs,
                                    const ResolvedAwsContext& ctx,
                                    const S3Location& loc,
                                    std::vector<FilesInformationS3::Entry>& out) noexcept
{
    out.clear();

    const std::shared_ptr<Aws::S3Crt::S3CrtClient> client = GetS3Client(fs, ctx);
    Aws::S3Crt::Model::ListObjectsV2Request req;
    req.SetBucket(loc.bucket);
    req.SetDelimiter("/");
    if (! loc.keyOrPrefix.empty())
    {
        req.SetPrefix(loc.keyOrPrefix);
    }
    req.SetMaxKeys(static_cast<int>(std::min<unsigned long>(ctx.maxKeys, 1000u)));

    // With delimiter set, CommonPrefixes can repeat across pages if a page boundary lands inside a prefix group.
    // Dedupe by the immediate-child name (UTF-8, case-sensitive) to avoid emitting duplicate directory entries.
    std::unordered_set<std::string> seenDirectoryNamesUtf8;
    seenDirectoryNamesUtf8.reserve(256);

    while (true)
    {
        const auto outcome = client->ListObjectsV2(req);
        if (! outcome.IsSuccess())
        {
            const auto& err            = outcome.GetError();
            const std::wstring details = std::format(L"bucket='{}' prefix='{}'", Utf16FromUtf8(loc.bucket), Utf16FromUtf8(loc.keyOrPrefix));
            LogAwsFailure(L"S3", L"ListObjectsV2", ctx, err, details);
            return HresultFromAwsError(err);
        }

        const auto& result = outcome.GetResult();

        // Directories (common prefixes)
        for (const auto& cp : result.GetCommonPrefixes())
        {
            const Aws::String& full = cp.GetPrefix();
            std::string_view fullView(full.c_str(), full.size());

            if (! loc.keyOrPrefix.empty() && fullView.rfind(loc.keyOrPrefix, 0) == 0)
            {
                fullView.remove_prefix(loc.keyOrPrefix.size());
            }

            while (! fullView.empty() && fullView.back() == '/')
            {
                fullView.remove_suffix(1);
            }

            if (fullView.empty())
            {
                continue;
            }

            if (! seenDirectoryNamesUtf8.emplace(fullView).second)
            {
                continue;
            }

            FilesInformationS3::Entry e{};
            e.name       = Utf16FromUtf8(fullView);
            e.attributes = FILE_ATTRIBUTE_DIRECTORY;
            out.push_back(std::move(e));
        }

        // Files
        for (const auto& obj : result.GetContents())
        {
            const Aws::String& key = obj.GetKey();
            std::string_view keyView(key.c_str(), key.size());

            // Skip the "folder marker" for the current prefix.
            if (! loc.keyOrPrefix.empty() && keyView == loc.keyOrPrefix)
            {
                continue;
            }

            if (! loc.keyOrPrefix.empty() && keyView.rfind(loc.keyOrPrefix, 0) == 0)
            {
                keyView.remove_prefix(loc.keyOrPrefix.size());
            }

            if (keyView.empty())
            {
                continue;
            }

            // With delimiter set, keys should not contain '/', but be defensive.
            const size_t slash = keyView.find('/');
            if (slash != std::string_view::npos)
            {
                keyView = keyView.substr(0, slash);
            }

            FilesInformationS3::Entry e{};
            e.name          = Utf16FromUtf8(keyView);
            e.attributes    = FILE_ATTRIBUTE_NORMAL;
            e.sizeBytes     = static_cast<uint64_t>(obj.GetSize());
            e.lastWriteTime = AwsDateTimeToFileTime64(obj.GetLastModified());
            e.changeTime    = e.lastWriteTime;
            out.push_back(std::move(e));
        }

        if (! result.GetIsTruncated())
        {
            break;
        }

        const Aws::String& token = result.GetNextContinuationToken();
        req.SetContinuationToken(token);
    }

    return S_OK;
}

[[nodiscard]] HRESULT DownloadS3ObjectToTempFile(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, wil::unique_hfile& outFile) noexcept
{
    outFile.reset();

    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

    wil::unique_hfile file = CreateTemporaryDeleteOnCloseFile();
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const std::shared_ptr<Aws::S3Crt::S3CrtClient> client = GetS3Client(fs, ctx);
    Aws::S3Crt::Model::GetObjectRequest req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetKey(Aws::String(key.data(), key.size()));

    auto outcome = client->GetObject(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' key='{}'", Utf16FromUtf8(bucket), Utf16FromUtf8(key));
        LogAwsFailure(L"S3", L"GetObject", ctx, err, details);
        return HresultFromAwsError(err);
    }

    auto result           = outcome.GetResultWithOwnership();
    Aws::IOStream& stream = result.GetBody();

    std::array<char, 64 * 1024> buffer{};
    while (stream.good())
    {
        stream.read(buffer.data(), static_cast<std::streamsize>(buffer.size()));
        const std::streamsize got = stream.gcount();
        if (got <= 0)
        {
            break;
        }

        if (got > static_cast<std::streamsize>((std::numeric_limits<DWORD>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        const HRESULT writeHr = Common::HandleIo::WriteAll(file.get(), buffer.data(), static_cast<size_t>(got));
        if (FAILED(writeHr))
        {
            return writeHr;
        }
    }

    const HRESULT seekHr = ResetFilePointerToStart(file.get());
    if (FAILED(seekHr))
    {
        return seekHr;
    }

    outFile = std::move(file);
    return S_OK;
}

namespace
{
[[nodiscard]] HRESULT ReadExactBytesFromFile(HANDLE file, void* buffer, size_t sizeBytes) noexcept
{
    return Common::HandleIo::ReadExact(file, buffer, sizeBytes);
}

class HandleReadStreamBuf final : public std::streambuf
{
public:
    explicit HandleReadStreamBuf(HANDLE file) noexcept : _file(file)
    {
        setg(_buffer.data(), _buffer.data(), _buffer.data());
    }

    HandleReadStreamBuf(const HandleReadStreamBuf&)            = delete;
    HandleReadStreamBuf(HandleReadStreamBuf&&)                 = delete;
    HandleReadStreamBuf& operator=(const HandleReadStreamBuf&) = delete;
    HandleReadStreamBuf& operator=(HandleReadStreamBuf&&)      = delete;

    [[nodiscard]] HRESULT GetReadError() const noexcept
    {
        return _readErrorHr;
    }

protected:
    int_type underflow() override
    {
        if (FAILED(_readErrorHr))
        {
            return traits_type::eof();
        }

        if (gptr() < egptr())
        {
            return traits_type::to_int_type(*gptr());
        }

        if (! _file || _file == INVALID_HANDLE_VALUE)
        {
            _readErrorHr = HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
            return traits_type::eof();
        }

        DWORD read = 0;
        if (ReadFile(_file, _buffer.data(), static_cast<DWORD>(_buffer.size()), &read, nullptr) == 0)
        {
            const DWORD lastError = GetLastError();
            _readErrorHr          = HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_READ_FAULT);
            return traits_type::eof();
        }

        if (read == 0)
        {
            return traits_type::eof();
        }

        setg(_buffer.data(), _buffer.data(), _buffer.data() + read);
        return traits_type::to_int_type(*gptr());
    }

private:
    HANDLE _file         = nullptr;
    HRESULT _readErrorHr = S_OK;
    std::array<char, 64 * 1024> _buffer{};
};

class HandleReadIStream final : public Aws::IOStream
{
public:
    explicit HandleReadIStream(HANDLE file) noexcept : Aws::IOStream(nullptr), _buf(file)
    {
        rdbuf(&_buf);
    }

    HandleReadIStream(const HandleReadIStream&)            = delete;
    HandleReadIStream(HandleReadIStream&&)                 = delete;
    HandleReadIStream& operator=(const HandleReadIStream&) = delete;
    HandleReadIStream& operator=(HandleReadIStream&&)      = delete;

    [[nodiscard]] HRESULT GetReadError() const noexcept
    {
        return _buf.GetReadError();
    }

private:
    HandleReadStreamBuf _buf;
};

[[nodiscard]] Aws::String ToAwsString(std::string_view value)
{
    return Aws::String(value.data(), value.size());
}

[[nodiscard]] HRESULT UploadPartCopy(FileSystemS3& fs,
                                     const S3MultipartUploadSession& session,
                                     std::string_view sourceBucket,
                                     std::string_view sourceKey,
                                     uint64_t startOffset,
                                     uint64_t endInclusive,
                                     int partNumber,
                                     std::string& outETag) noexcept
{
    outETag.clear();

    const auto client = GetS3Client(fs, session.ctx);

    Aws::S3Crt::Model::UploadPartCopyRequest req;
    req.SetBucket(ToAwsString(session.bucket));
    req.SetKey(ToAwsString(session.key));
    req.SetUploadId(ToAwsString(session.uploadId));
    req.SetPartNumber(partNumber);
    req.SetCopySource(ToAwsString(UrlEncodeS3CopySource(sourceBucket, sourceKey)));

    const std::string range = std::string("bytes=") + std::to_string(startOffset) + "-" + std::to_string(endInclusive);
    req.SetCopySourceRange(ToAwsString(range));

    const auto outcome = client->UploadPartCopy(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"src='{}:{}' dst='{}:{}' part={} range='{}'",
                                                 Utf16FromUtf8(sourceBucket),
                                                 Utf16FromUtf8(sourceKey),
                                                 Utf16FromUtf8(session.bucket),
                                                 Utf16FromUtf8(session.key),
                                                 partNumber,
                                                 Utf16FromUtf8(range));
        LogAwsFailure(L"S3", L"UploadPartCopy", session.ctx, err, details);
        return HresultFromAwsError(err);
    }

    const Aws::String& etag = outcome.GetResult().GetCopyPartResult().GetETag();
    outETag.assign(etag.c_str(), etag.size());
    return S_OK;
}
} // namespace

[[nodiscard]] HRESULT PutS3ObjectFromMemory(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, const void* data, size_t sizeBytes) noexcept
{
    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

    if (sizeBytes > 0 && data == nullptr)
    {
        return E_POINTER;
    }

    if (sizeBytes > static_cast<size_t>((std::numeric_limits<long long>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const std::shared_ptr<Aws::S3Crt::S3CrtClient> client = GetS3Client(fs, ctx);
    Aws::S3Crt::Model::PutObjectRequest req;
    req.SetBucket(ToAwsString(bucket));
    req.SetKey(ToAwsString(key));
    req.SetContentLength(static_cast<long long>(sizeBytes));

    auto body = Aws::MakeShared<Aws::StringStream>("rs3-put-memory");
    if (sizeBytes > 0)
    {
        body->write(static_cast<const char*>(data), static_cast<std::streamsize>(sizeBytes));
        body->seekg(0, std::ios::beg);
    }
    req.SetBody(body);

    const auto outcome = client->PutObject(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' key='{}'", Utf16FromUtf8(bucket), Utf16FromUtf8(key));
        LogAwsFailure(L"S3", L"PutObject", ctx, err, details);
        return HresultFromAwsError(err);
    }

    return S_OK;
}

[[nodiscard]] HRESULT BeginS3MultipartUpload(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, S3MultipartUploadSession& outSession) noexcept
{
    outSession = {};

    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

    const auto client = GetS3Client(fs, ctx);

    Aws::S3Crt::Model::CreateMultipartUploadRequest req;
    req.SetBucket(ToAwsString(bucket));
    req.SetKey(ToAwsString(key));

    const auto outcome = client->CreateMultipartUpload(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' key='{}'", Utf16FromUtf8(bucket), Utf16FromUtf8(key));
        LogAwsFailure(L"S3", L"CreateMultipartUpload", ctx, err, details);
        return HresultFromAwsError(err);
    }

    outSession.ctx      = ctx;
    outSession.bucket   = std::string(bucket);
    outSession.key      = std::string(key);
    outSession.uploadId = std::string(outcome.GetResult().GetUploadId().c_str(), outcome.GetResult().GetUploadId().size());
    return outSession.uploadId.empty() ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : S_OK;
}

[[nodiscard]] HRESULT UploadS3MultipartPartFromMemory(
    FileSystemS3& fs, const S3MultipartUploadSession& session, int partNumber, const void* data, size_t sizeBytes, std::string& outETag) noexcept
{
    outETag.clear();

    if (partNumber <= 0 || sizeBytes == 0u)
    {
        return E_INVALIDARG;
    }

    if (data == nullptr)
    {
        return E_POINTER;
    }

    if (sizeBytes > static_cast<size_t>((std::numeric_limits<long long>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const auto client = GetS3Client(fs, session.ctx);

    Aws::S3Crt::Model::UploadPartRequest req;
    req.SetBucket(ToAwsString(session.bucket));
    req.SetKey(ToAwsString(session.key));
    req.SetUploadId(ToAwsString(session.uploadId));
    req.SetPartNumber(partNumber);
    req.SetContentLength(static_cast<long long>(sizeBytes));

    auto body = Aws::MakeShared<Aws::StringStream>("rs3-upload-part");
    body->write(static_cast<const char*>(data), static_cast<std::streamsize>(sizeBytes));
    body->seekg(0, std::ios::beg);
    req.SetBody(body);

    const auto outcome = client->UploadPart(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' key='{}' uploadId='{}' part={}",
                                                 Utf16FromUtf8(session.bucket),
                                                 Utf16FromUtf8(session.key),
                                                 Utf16FromUtf8(session.uploadId),
                                                 partNumber);
        LogAwsFailure(L"S3", L"UploadPart", session.ctx, err, details);
        return HresultFromAwsError(err);
    }

    const Aws::String& etag = outcome.GetResult().GetETag();
    outETag.assign(etag.c_str(), etag.size());
    return outETag.empty() ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : S_OK;
}

[[nodiscard]] HRESULT CompleteS3MultipartUpload(FileSystemS3& fs,
                                                const S3MultipartUploadSession& session,
                                                const std::vector<S3MultipartUploadedPart>& parts) noexcept
{
    if (session.uploadId.empty() || parts.empty())
    {
        return E_INVALIDARG;
    }

    const auto client = GetS3Client(fs, session.ctx);

    Aws::S3Crt::Model::CompletedMultipartUpload completed;
    for (const auto& part : parts)
    {
        Aws::S3Crt::Model::CompletedPart completedPart;
        completedPart.SetPartNumber(part.partNumber);
        completedPart.SetETag(ToAwsString(part.eTag));
        completed.AddParts(std::move(completedPart));
    }

    Aws::S3Crt::Model::CompleteMultipartUploadRequest req;
    req.SetBucket(ToAwsString(session.bucket));
    req.SetKey(ToAwsString(session.key));
    req.SetUploadId(ToAwsString(session.uploadId));
    req.SetMultipartUpload(std::move(completed));

    const auto outcome = client->CompleteMultipartUpload(req);
    if (! outcome.IsSuccess())
    {
        const auto& err = outcome.GetError();
        const std::wstring details =
            std::format(L"bucket='{}' key='{}' uploadId='{}'", Utf16FromUtf8(session.bucket), Utf16FromUtf8(session.key), Utf16FromUtf8(session.uploadId));
        LogAwsFailure(L"S3", L"CompleteMultipartUpload", session.ctx, err, details);
        return HresultFromAwsError(err);
    }

    return S_OK;
}

[[nodiscard]] HRESULT AbortS3MultipartUpload(FileSystemS3& fs, const S3MultipartUploadSession& session) noexcept
{
    if (session.uploadId.empty())
    {
        return S_OK;
    }

    const auto client = GetS3Client(fs, session.ctx);

    Aws::S3Crt::Model::AbortMultipartUploadRequest req;
    req.SetBucket(ToAwsString(session.bucket));
    req.SetKey(ToAwsString(session.key));
    req.SetUploadId(ToAwsString(session.uploadId));

    const auto outcome = client->AbortMultipartUpload(req);
    if (! outcome.IsSuccess())
    {
        const auto& err = outcome.GetError();
        const std::wstring details =
            std::format(L"bucket='{}' key='{}' uploadId='{}'", Utf16FromUtf8(session.bucket), Utf16FromUtf8(session.key), Utf16FromUtf8(session.uploadId));
        LogAwsFailure(L"S3", L"AbortMultipartUpload", session.ctx, err, details);
        return HresultFromAwsError(err);
    }

    return S_OK;
}

[[nodiscard]] HRESULT CopyS3ObjectServerSide(FileSystemS3& fs,
                                             const ResolvedAwsContext& destinationCtx,
                                             std::string_view sourceBucket,
                                             std::string_view sourceKey,
                                             std::string_view destinationBucket,
                                             std::string_view destinationKey,
                                             uint64_t sourceSizeBytes) noexcept
{
    if (sourceBucket.empty() || sourceKey.empty() || destinationBucket.empty() || destinationKey.empty())
    {
        return E_INVALIDARG;
    }

    const auto client            = GetS3Client(fs, destinationCtx);
    const std::string copySource = UrlEncodeS3CopySource(sourceBucket, sourceKey);

    if (sourceSizeBytes < kMultipartMinPartSizeBytes)
    {
        Aws::S3Crt::Model::CopyObjectRequest req;
        req.SetBucket(ToAwsString(destinationBucket));
        req.SetKey(ToAwsString(destinationKey));
        req.SetCopySource(ToAwsString(copySource));

        const auto outcome = client->CopyObject(req);
        if (! outcome.IsSuccess())
        {
            const auto& err            = outcome.GetError();
            const std::wstring details = std::format(L"src='{}:{}' dst='{}:{}'",
                                                     Utf16FromUtf8(sourceBucket),
                                                     Utf16FromUtf8(sourceKey),
                                                     Utf16FromUtf8(destinationBucket),
                                                     Utf16FromUtf8(destinationKey));
            LogAwsFailure(L"S3", L"CopyObject", destinationCtx, err, details);
            return HresultFromAwsError(err);
        }

        return S_OK;
    }

    S3MultipartUploadSession session{};
    HRESULT hr = BeginS3MultipartUpload(fs, destinationCtx, destinationBucket, destinationKey, session);
    if (FAILED(hr))
    {
        return hr;
    }

    bool completeSucceeded    = false;
    const auto abortOnFailure = wil::scope_exit([&]() noexcept
    {
        if (! completeSucceeded)
        {
            const HRESULT abortHr = AbortS3MultipartUpload(fs, session);
            if (FAILED(abortHr))
            {
                Debug::Warning(L"S3: failed to abort multipart copy upload '{}' (hr={:#x})", Utf16FromUtf8(session.key), static_cast<unsigned long>(abortHr));
            }
        }
    });

    const uint64_t partSize = ComputeMultipartPartSize(sourceSizeBytes);
    std::vector<S3MultipartUploadedPart> parts;
    for (uint64_t offset = 0, partIndex = 1; offset < sourceSizeBytes; offset += partSize, ++partIndex)
    {
        const uint64_t partBytes    = std::min<uint64_t>(partSize, sourceSizeBytes - offset);
        const uint64_t endInclusive = offset + partBytes - 1u;

        std::string etag;
        hr = UploadPartCopy(fs, session, sourceBucket, sourceKey, offset, endInclusive, static_cast<int>(partIndex), etag);
        if (FAILED(hr))
        {
            return hr;
        }

        S3MultipartUploadedPart uploaded{};
        uploaded.partNumber = static_cast<int>(partIndex);
        uploaded.eTag       = std::move(etag);
        parts.push_back(std::move(uploaded));
    }

    hr = CompleteS3MultipartUpload(fs, session, parts);
    if (FAILED(hr))
    {
        return hr;
    }

    completeSucceeded = true;
    return S_OK;
}

[[nodiscard]] HRESULT UploadS3ObjectFromFile(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, HANDLE file, uint64_t sizeBytes) noexcept
{
    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

    if (! file || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    if (sizeBytes > static_cast<uint64_t>((std::numeric_limits<long long>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    if (sizeBytes >= kMultipartMinPartSizeBytes)
    {
        const HRESULT resetHr = ResetFilePointerToStart(file);
        if (FAILED(resetHr))
        {
            return resetHr;
        }

        S3MultipartUploadSession session{};
        HRESULT hr = BeginS3MultipartUpload(fs, ctx, bucket, key, session);
        if (FAILED(hr))
        {
            return hr;
        }

        bool completeSucceeded    = false;
        const auto abortOnFailure = wil::scope_exit([&]() noexcept
        {
            if (! completeSucceeded)
            {
                const HRESULT abortHr = AbortS3MultipartUpload(fs, session);
                if (FAILED(abortHr))
                {
                    Debug::Warning(L"S3: failed to abort multipart upload '{}' (hr={:#x})", Utf16FromUtf8(session.key), static_cast<unsigned long>(abortHr));
                }
            }
        });

        const uint64_t partSize = ComputeMultipartPartSize(sizeBytes);
        std::vector<S3MultipartUploadedPart> parts;
        std::vector<std::byte> buffer;

        uint64_t remaining = sizeBytes;
        int partNumber     = 1;
        while (remaining > 0)
        {
            const uint64_t chunkBytes64 = std::min<uint64_t>(partSize, remaining);
            if (chunkBytes64 > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }

            const size_t chunkBytes = static_cast<size_t>(chunkBytes64);
            buffer.resize(chunkBytes);

            hr = ReadExactBytesFromFile(file, buffer.data(), chunkBytes);
            if (FAILED(hr))
            {
                return hr;
            }

            std::string etag;
            hr = UploadS3MultipartPartFromMemory(fs, session, partNumber, buffer.data(), chunkBytes, etag);
            if (FAILED(hr))
            {
                return hr;
            }

            S3MultipartUploadedPart uploaded{};
            uploaded.partNumber = partNumber;
            uploaded.eTag       = std::move(etag);
            parts.push_back(std::move(uploaded));

            remaining -= chunkBytes64;
            ++partNumber;
        }

        hr = CompleteS3MultipartUpload(fs, session, parts);
        if (FAILED(hr))
        {
            return hr;
        }

        completeSucceeded = true;
        return S_OK;
    }

    return PutS3ObjectFromHandle(fs, ctx, bucket, key, file, sizeBytes);
}

namespace
{
[[nodiscard]] HRESULT PutS3ObjectFromHandle(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, HANDLE file, uint64_t sizeBytes) noexcept
{
    const std::shared_ptr<Aws::S3Crt::S3CrtClient> client = GetS3Client(fs, ctx);
    Aws::S3Crt::Model::PutObjectRequest req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetKey(Aws::String(key.data(), key.size()));
    req.SetContentLength(static_cast<long long>(sizeBytes));

    auto body = Aws::MakeShared<HandleReadIStream>("rs3-put", file);
    req.SetBody(body);

    const auto outcome = client->PutObject(req);

    const HRESULT readHr = body->GetReadError();
    if (FAILED(readHr))
    {
        return readHr;
    }

    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' key='{}'", Utf16FromUtf8(bucket), Utf16FromUtf8(key));
        LogAwsFailure(L"S3", L"PutObject", ctx, err, details);
        return HresultFromAwsError(err);
    }

    return S_OK;
}
} // namespace
} // namespace FileSystemS3Internal
