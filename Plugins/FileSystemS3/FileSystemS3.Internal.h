#pragma once

#include "FileSystemS3.h"

// Windows headers define ERROR as a macro; AWS Outcome.h declares an ERROR typedef.
#pragma warning(push)
#pragma warning(disable : 4574) // Intentional probe of Win32's zero-valued ERROR macro before undefining it.
#ifdef ERROR
#undef ERROR
#endif
#pragma warning(pop)

#include "Helpers.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <format>
#include <limits>
#include <memory>
#include <mutex>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include <aws/core/Aws.h>
#include <aws/core/auth/AWSCredentials.h>
#include <aws/core/client/AWSError.h>
#include <aws/core/client/ClientConfiguration.h>
#include <aws/core/client/CoreErrors.h>
#include <aws/core/utils/DateTime.h>
#include <aws/s3-crt/S3CrtClient.h>
#include <aws/s3tables/S3TablesClient.h>

namespace FileSystemS3Internal
{
struct ResolvedAwsContext
{
    // Set when resolved from a Connection Manager profile. Empty means defaults + AWS default credential chain.
    std::wstring connectionName;

    // Region used for signing and regional endpoints. Always set to a non-empty value.
    std::string region;

    // When set, the user explicitly selected a region (Connection Manager host field). When not set,
    // the S3 plugin may auto-resolve bucket regions when using AWS endpoints.
    std::optional<std::string> explicitRegion;
    std::string endpointOverride;
    bool useHttps                 = true;
    bool verifyTls                = true;
    bool useVirtualAddressing     = true;
    unsigned long maxKeys         = 1000;
    unsigned long maxTableResults = 1000;
    uint32_t connectTimeoutMs     = 10'000;
    uint32_t requestTimeoutMs     = 30'000;

    std::optional<std::string> accessKeyId;
    std::optional<std::string> secretAccessKey;

    ResolvedAwsContext()                              = default;
    ResolvedAwsContext(const ResolvedAwsContext&)     = default;
    ResolvedAwsContext(ResolvedAwsContext&&) noexcept = default;

    ResolvedAwsContext& operator=(const ResolvedAwsContext& other)
    {
        if (this == &other)
        {
            return *this;
        }

        SecureClearSensitiveFields();

        connectionName       = other.connectionName;
        region               = other.region;
        explicitRegion       = other.explicitRegion;
        endpointOverride     = other.endpointOverride;
        useHttps             = other.useHttps;
        verifyTls            = other.verifyTls;
        useVirtualAddressing = other.useVirtualAddressing;
        maxKeys              = other.maxKeys;
        maxTableResults      = other.maxTableResults;
        connectTimeoutMs     = other.connectTimeoutMs;
        requestTimeoutMs     = other.requestTimeoutMs;
        accessKeyId          = other.accessKeyId;
        secretAccessKey      = other.secretAccessKey;
        return *this;
    }

    ResolvedAwsContext& operator=(ResolvedAwsContext&& other) noexcept
    {
        if (this == &other)
        {
            return *this;
        }

        SecureClearSensitiveFields();

        connectionName       = std::move(other.connectionName);
        region               = std::move(other.region);
        explicitRegion       = std::move(other.explicitRegion);
        endpointOverride     = std::move(other.endpointOverride);
        useHttps             = other.useHttps;
        verifyTls            = other.verifyTls;
        useVirtualAddressing = other.useVirtualAddressing;
        maxKeys              = other.maxKeys;
        maxTableResults      = other.maxTableResults;
        connectTimeoutMs     = other.connectTimeoutMs;
        requestTimeoutMs     = other.requestTimeoutMs;
        accessKeyId          = std::move(other.accessKeyId);
        secretAccessKey      = std::move(other.secretAccessKey);
        other.SecureClearSensitiveFields();
        return *this;
    }

    ~ResolvedAwsContext() noexcept
    {
        SecureClearSensitiveFields();
    }

private:
    void SecureClearSensitiveFields() noexcept
    {
        if (accessKeyId.has_value())
        {
            SecureWipe::SecureClear(accessKeyId.value());
            accessKeyId.reset();
        }
        if (secretAccessKey.has_value())
        {
            SecureWipe::SecureClear(secretAccessKey.value());
            secretAccessKey.reset();
        }
    }
};

struct S3Location
{
    std::string bucket;
    std::string keyOrPrefix; // no leading '/'
    bool isRoot = false;     // true when listing buckets
};

struct S3MultipartUploadedPart
{
    int partNumber = 0;
    std::string eTag;
};

struct S3MultipartUploadSession
{
    ResolvedAwsContext ctx;
    std::string bucket;
    std::string key;
    std::string uploadId;
};

struct AwsSdkLifetime
{
    [[nodiscard]] static HRESULT Acquire() noexcept;
    static void Release() noexcept;
    static void BeginShutdown() noexcept;
    [[nodiscard]] static bool CanUnloadNow() noexcept;
};

void SchedulePendingMultipartAbortCleanup() noexcept;
[[nodiscard]] bool CanUnloadPendingMultipartAbortCleanup() noexcept;

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept;
[[nodiscard]] std::wstring Utf16FromUtf8(const char* text) noexcept;
[[nodiscard]] std::wstring Utf16FromUtf8(const Aws::String& text) noexcept;
[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept;

[[nodiscard]] std::wstring NormalizePluginPath(std::wstring_view rawPath) noexcept;

[[nodiscard]] __int64 UnixMsToFileTime64(uint64_t unixMs) noexcept;
[[nodiscard]] __int64 AwsDateTimeToFileTime64(const Aws::Utils::DateTime& t) noexcept;

[[nodiscard]] wil::unique_hfile CreateTemporaryDeleteOnCloseFile() noexcept;
[[nodiscard]] HRESULT GetFileSizeBytes(HANDLE file, uint64_t& out) noexcept;
[[nodiscard]] HRESULT ResetFilePointerToStart(HANDLE file) noexcept;
[[nodiscard]] HRESULT WriteUtf8ToFile(HANDLE file, std::string_view text) noexcept;

[[nodiscard]] std::optional<std::wstring> TryGetJsonString(yyjson_val* root, const char* key) noexcept;
[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt(yyjson_val* root, const char* key) noexcept;
[[nodiscard]] std::optional<bool> TryGetJsonBool(yyjson_val* root, const char* key) noexcept;

[[nodiscard]] bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept;

[[nodiscard]] HRESULT ResolveAwsContext(FileSystemS3Mode mode,
                                        const FileSystemS3::Settings& defaults,
                                        std::wstring_view pluginPath,
                                        IHostConnections* hostConnections,
                                        bool acquireSecrets,
                                        ResolvedAwsContext& outContext,
                                        std::wstring& outCanonicalPath) noexcept;

template <typename AwsErrors> [[nodiscard]] HRESULT HresultFromAwsError(const Aws::Client::AWSError<AwsErrors>& err) noexcept
{
    using Aws::Http::HttpResponseCode;

    const HttpResponseCode code = err.GetResponseCode();
    if (code == HttpResponseCode::NOT_FOUND)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    if (code == HttpResponseCode::FORBIDDEN)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    if (code == HttpResponseCode::UNAUTHORIZED)
    {
        return HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
    }
    if (code == HttpResponseCode::REQUEST_TIMEOUT)
    {
        return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    }
    if (code == HttpResponseCode::PRECONDITION_FAILED || code == HttpResponseCode::REQUESTED_RANGE_NOT_SATISFIABLE)
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }

    const int errType = static_cast<int>(err.GetErrorType());
    if (errType == static_cast<int>(Aws::Client::CoreErrors::NETWORK_CONNECTION))
    {
        return HRESULT_FROM_WIN32(ERROR_UNEXP_NET_ERR);
    }
    if (errType == static_cast<int>(Aws::Client::CoreErrors::ENDPOINT_RESOLUTION_FAILURE))
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME);
    }
    if (errType == static_cast<int>(Aws::Client::CoreErrors::REQUEST_TIMEOUT))
    {
        return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    }
    if (errType == static_cast<int>(Aws::Client::CoreErrors::USER_CANCELLED))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    if (errType == static_cast<int>(Aws::Client::CoreErrors::ACCESS_DENIED))
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    if (errType == static_cast<int>(Aws::Client::CoreErrors::INVALID_ACCESS_KEY_ID) ||
        errType == static_cast<int>(Aws::Client::CoreErrors::INVALID_SIGNATURE) ||
        errType == static_cast<int>(Aws::Client::CoreErrors::SIGNATURE_DOES_NOT_MATCH) ||
        errType == static_cast<int>(Aws::Client::CoreErrors::UNRECOGNIZED_CLIENT))
    {
        return HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
    }

    return E_FAIL;
}

[[nodiscard]] Aws::Client::ClientConfiguration MakeClientConfig(const ResolvedAwsContext& ctx) noexcept;
[[nodiscard]] std::shared_ptr<Aws::S3Crt::S3CrtClient> MakeS3Client(const ResolvedAwsContext& ctx) noexcept;
[[nodiscard]] Aws::S3Tables::S3TablesClient MakeS3TablesClient(const ResolvedAwsContext& ctx) noexcept;
[[nodiscard]] std::shared_ptr<Aws::S3Crt::S3CrtClient> GetS3Client(FileSystemS3& fs, const ResolvedAwsContext& ctx) noexcept;
[[nodiscard]] HRESULT TryGetS3ObjectSummary(FileSystemS3& fs,
                                            const ResolvedAwsContext& bucketCtx,
                                            std::string_view bucket,
                                            std::string_view key,
                                            uint64_t& outSizeBytes,
                                            __int64& outLastWriteTime,
                                            bool& outFound) noexcept;
[[nodiscard]] HRESULT ValidateS3RangeResponseLength(uint64_t expectedBytes, long long responseContentLength, uint64_t bodyBytesRead) noexcept;
[[nodiscard]] HRESULT ValidateS3UploadReadResult(uint64_t declaredBytes, uint64_t consumedBytes, HRESULT readStatus) noexcept;

#if defined(_DEBUG)
void RunDebugRangeReadContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugMultipartWriterContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugDirectorySizeCallbackContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugAwsSdkLifetimeContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
[[nodiscard]] bool WaitForPendingMultipartAbortCleanupForTest(unsigned long timeoutMs) noexcept;
#endif

[[nodiscard]] constexpr std::wstring_view CoreErrorNameFromInt(int code) noexcept
{
    switch (code)
    {
        case static_cast<int>(Aws::Client::CoreErrors::NETWORK_CONNECTION): return L"NETWORK_CONNECTION";
        case static_cast<int>(Aws::Client::CoreErrors::ENDPOINT_RESOLUTION_FAILURE): return L"ENDPOINT_RESOLUTION_FAILURE";
        case static_cast<int>(Aws::Client::CoreErrors::REQUEST_TIMEOUT): return L"REQUEST_TIMEOUT";
        case static_cast<int>(Aws::Client::CoreErrors::ACCESS_DENIED): return L"ACCESS_DENIED";
        case static_cast<int>(Aws::Client::CoreErrors::INVALID_ACCESS_KEY_ID): return L"INVALID_ACCESS_KEY_ID";
        case static_cast<int>(Aws::Client::CoreErrors::SIGNATURE_DOES_NOT_MATCH): return L"SIGNATURE_DOES_NOT_MATCH";
        case static_cast<int>(Aws::Client::CoreErrors::UNRECOGNIZED_CLIENT): return L"UNRECOGNIZED_CLIENT";
        case static_cast<int>(Aws::Client::CoreErrors::UNKNOWN): return L"UNKNOWN";
        default: return {};
    }
}

template <typename AwsErrors>
inline void LogAwsFailure(std::wstring_view prefix,
                          std::wstring_view operation,
                          const ResolvedAwsContext& ctx,
                          const Aws::Client::AWSError<AwsErrors>& err,
                          std::wstring_view details) noexcept
{
    const Aws::Client::ClientConfiguration cfg = MakeClientConfig(ctx);
    const wchar_t* scheme                      = (cfg.scheme == Aws::Http::Scheme::HTTPS) ? L"https" : L"http";
    const int errType                          = static_cast<int>(err.GetErrorType());
    const std::wstring_view errTypeName        = CoreErrorNameFromInt(errType);

    Debug::Error(L"{}: {} failed {}"
                 L" conn='{}' region='{}' endpoint='{}' scheme='{}' verifyTls={} connectTimeoutMs={} requestTimeoutMs={} virtualAddressing={}"
                 L" errType={} errTypeName='{}' http={} retry={} requestId='{}' remoteIp='{}' exception='{}' message='{}'",
                 prefix,
                 operation,
                 details,
                 ctx.connectionName,
                 Utf16FromUtf8(cfg.region),
                 Redaction::ForLog(Utf16FromUtf8(cfg.endpointOverride)),
                 scheme,
                 cfg.verifySSL ? 1 : 0,
                 cfg.connectTimeoutMs,
                 cfg.requestTimeoutMs,
                 ctx.useVirtualAddressing ? 1 : 0,
                 errType,
                 errTypeName,
                 static_cast<int>(err.GetResponseCode()),
                 err.ShouldRetry() ? 1 : 0,
                 Utf16FromUtf8(err.GetRequestId().c_str()),
                 Utf16FromUtf8(err.GetRemoteHostIpAddress().c_str()),
                 Utf16FromUtf8(err.GetExceptionName().c_str()),
                 Utf16FromUtf8(err.GetMessage().c_str()));
}

[[nodiscard]] std::vector<std::wstring_view> SplitPathSegments(std::wstring_view path) noexcept;

inline constexpr uint64_t kMultipartMinPartSizeBytes = 64ull * 1024ull * 1024ull;

[[nodiscard]] uint64_t ComputeMultipartPartSize(uint64_t sizeBytes) noexcept;
[[nodiscard]] bool IsSameAwsContextIdentity(const ResolvedAwsContext& left, const ResolvedAwsContext& right) noexcept;
[[nodiscard]] std::string UrlEncodeS3CopySource(std::string_view bucket, std::string_view key) noexcept;

[[nodiscard]] HRESULT ParseS3LocationForDirectory(std::wstring_view canonicalPath, S3Location& out) noexcept;
[[nodiscard]] HRESULT ListS3Buckets(FileSystemS3& fs, const ResolvedAwsContext& ctx, std::vector<FilesInformationS3::Entry>& out) noexcept;
[[nodiscard]] HRESULT ListS3BucketsForConnection(FileSystemS3& fs, const ResolvedAwsContext& ctx, std::vector<FilesInformationS3::Entry>& out) noexcept;
[[nodiscard]] HRESULT ListS3Objects(FileSystemS3& fs,
                                    const ResolvedAwsContext& ctx,
                                    const S3Location& loc,
                                    std::vector<FilesInformationS3::Entry>& out) noexcept;
[[nodiscard]] HRESULT ResolveS3ContextForBucket(FileSystemS3& fs,
                                                const ResolvedAwsContext& ctx,
                                                std::wstring_view bucketName,
                                                ResolvedAwsContext& out) noexcept;
[[nodiscard]] HRESULT DownloadS3ObjectToTempFile(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, wil::unique_hfile& outFile) noexcept;
[[nodiscard]] HRESULT PutS3ObjectFromMemory(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, const void* data, size_t sizeBytes) noexcept;
[[nodiscard]] HRESULT BeginS3MultipartUpload(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, S3MultipartUploadSession& outSession) noexcept;
[[nodiscard]] HRESULT UploadS3MultipartPartFromMemory(
    FileSystemS3& fs, const S3MultipartUploadSession& session, int partNumber, const void* data, size_t sizeBytes, std::string& outETag) noexcept;
[[nodiscard]] HRESULT CompleteS3MultipartUpload(FileSystemS3& fs,
                                                const S3MultipartUploadSession& session,
                                                const std::vector<S3MultipartUploadedPart>& parts) noexcept;
[[nodiscard]] HRESULT AbortS3MultipartUpload(FileSystemS3& fs, const S3MultipartUploadSession& session) noexcept;
[[nodiscard]] HRESULT CopyS3ObjectServerSide(FileSystemS3& fs,
                                             const ResolvedAwsContext& destinationCtx,
                                             std::string_view sourceBucket,
                                             std::string_view sourceKey,
                                             std::string_view destinationBucket,
                                             std::string_view destinationKey,
                                             uint64_t sourceSizeBytes) noexcept;
[[nodiscard]] HRESULT UploadS3ObjectFromFile(
    FileSystemS3& fs, const ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key, HANDLE file, uint64_t sizeBytes) noexcept;

[[nodiscard]] HRESULT ListS3TableNamespaces(FileSystemS3& fs,
                                            const ResolvedAwsContext& ctx,
                                            std::wstring_view bucketName,
                                            std::vector<FilesInformationS3::Entry>& out) noexcept;

[[nodiscard]] HRESULT ListS3TableTables(FileSystemS3& fs,
                                        const ResolvedAwsContext& ctx,
                                        std::wstring_view bucketName,
                                        std::wstring_view nsName,
                                        std::vector<FilesInformationS3::Entry>& out) noexcept;

[[nodiscard]] HRESULT WriteS3TableInfoJson(FileSystemS3& fs,
                                           const ResolvedAwsContext& ctx,
                                           std::wstring_view bucketName,
                                           std::wstring_view nsName,
                                           std::wstring_view tableName,
                                           wil::unique_hfile& outFile) noexcept;
} // namespace FileSystemS3Internal
