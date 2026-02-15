#include "FileSystemS3.Internal.h"

#include <aws/core/utils/json/JsonSerializer.h>
#include <aws/s3tables/model/GetTableRequest.h>
#include <aws/s3tables/model/ListNamespacesRequest.h>
#include <aws/s3tables/model/ListTableBucketsRequest.h>
#include <aws/s3tables/model/ListTablesRequest.h>

HRESULT ListS3TableBuckets(FileSystemS3& fs, const FileSystemS3Internal::ResolvedAwsContext& ctx, std::vector<FilesInformationS3::Entry>& out) noexcept
{
    out.clear();

    Aws::S3Tables::S3TablesClient client = FileSystemS3Internal::MakeS3TablesClient(ctx);
    const auto outcome                   = client.ListTableBuckets(Aws::S3Tables::Model::ListTableBucketsRequest());
    if (! outcome.IsSuccess())
    {
        const auto& err = outcome.GetError();
        FileSystemS3Internal::LogAwsFailure(L"S3Tables", L"ListTableBuckets", ctx, err, L"tableBuckets");
        return FileSystemS3Internal::HresultFromAwsError(err);
    }

    const auto& buckets = outcome.GetResult().GetTableBuckets();
    out.reserve(buckets.size());

    std::unordered_map<std::wstring, std::string> cache;
    cache.reserve(buckets.size());

    for (const auto& bucket : buckets)
    {
        FilesInformationS3::Entry e{};
        e.name          = FileSystemS3Internal::Utf16FromUtf8(bucket.GetName());
        e.attributes    = FILE_ATTRIBUTE_DIRECTORY;
        e.creationTime  = FileSystemS3Internal::AwsDateTimeToFileTime64(bucket.GetCreatedAt());
        e.lastWriteTime = e.creationTime;
        e.changeTime    = e.creationTime;
        out.push_back(std::move(e));

        if (! e.name.empty() && bucket.ArnHasBeenSet())
        {
            cache.emplace(e.name, std::string(bucket.GetArn().c_str(), bucket.GetArn().size()));
        }
    }

    {
        std::lock_guard lock(fs._stateMutex);
        fs._s3TableBucketArnByName = std::move(cache);
    }

    return S_OK;
}

std::optional<std::string> LookupS3TableBucketArn(FileSystemS3& fs, std::wstring_view bucketName) noexcept
{
    std::lock_guard lock(fs._stateMutex);
    if (const auto it = fs._s3TableBucketArnByName.find(std::wstring(bucketName)); it != fs._s3TableBucketArnByName.end())
    {
        return it->second;
    }
    return std::nullopt;
}

namespace FileSystemS3Internal
{
namespace
{
[[nodiscard]] HRESULT EnsureS3TableBucketArn(FileSystemS3& fs, const ResolvedAwsContext& ctx, std::wstring_view bucketName, std::string& outArn) noexcept
{
    outArn.clear();

    if (const auto cached = LookupS3TableBucketArn(fs, bucketName); cached.has_value())
    {
        outArn = cached.value();
        return S_OK;
    }

    std::vector<FilesInformationS3::Entry> unused;
    const HRESULT refreshHr = ListS3TableBuckets(fs, ctx, unused);
    if (FAILED(refreshHr))
    {
        return refreshHr;
    }

    if (const auto cached = LookupS3TableBucketArn(fs, bucketName); cached.has_value())
    {
        outArn = cached.value();
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}
} // namespace

[[nodiscard]] HRESULT
ListS3TableNamespaces(FileSystemS3& fs, const ResolvedAwsContext& ctx, std::wstring_view bucketName, std::vector<FilesInformationS3::Entry>& out) noexcept
{
    out.clear();

    std::string bucketArn;
    const HRESULT arnHr = EnsureS3TableBucketArn(fs, ctx, bucketName, bucketArn);
    if (FAILED(arnHr))
    {
        return arnHr;
    }

    Aws::S3Tables::S3TablesClient client = MakeS3TablesClient(ctx);
    Aws::S3Tables::Model::ListNamespacesRequest req;
    req.SetTableBucketARN(bucketArn);
    req.SetMaxNamespaces(static_cast<int>(std::min<unsigned long>(ctx.maxTableResults, 1000u)));

    const auto outcome = client.ListNamespaces(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}'", bucketName);
        LogAwsFailure(L"S3Tables", L"ListNamespaces", ctx, err, details);
        return HresultFromAwsError(err);
    }

    for (const auto& ns : outcome.GetResult().GetNamespaces())
    {
        std::string joined;
        for (const auto& part : ns.GetNamespace())
        {
            if (! joined.empty())
            {
                joined.push_back('.');
            }
            joined.append(part.c_str(), part.size());
        }

        FilesInformationS3::Entry e{};
        e.name          = Utf16FromUtf8(joined);
        e.attributes    = FILE_ATTRIBUTE_DIRECTORY;
        e.creationTime  = AwsDateTimeToFileTime64(ns.GetCreatedAt());
        e.lastWriteTime = e.creationTime;
        e.changeTime    = e.creationTime;
        out.push_back(std::move(e));
    }

    return S_OK;
}

[[nodiscard]] HRESULT ListS3TableTables(FileSystemS3& fs,
                                        const ResolvedAwsContext& ctx,
                                        std::wstring_view bucketName,
                                        std::wstring_view nsName,
                                        std::vector<FilesInformationS3::Entry>& out) noexcept
{
    out.clear();

    std::string bucketArn;
    const HRESULT arnHr = EnsureS3TableBucketArn(fs, ctx, bucketName, bucketArn);
    if (FAILED(arnHr))
    {
        return arnHr;
    }

    const std::string nsUtf8 = Utf8FromUtf16(nsName);
    if (nsUtf8.empty() && ! nsName.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    Aws::S3Tables::S3TablesClient client = MakeS3TablesClient(ctx);
    Aws::S3Tables::Model::ListTablesRequest req;
    req.SetTableBucketARN(bucketArn);
    if (! nsUtf8.empty())
    {
        req.SetNamespace(nsUtf8);
    }
    req.SetMaxTables(static_cast<int>(std::min<unsigned long>(ctx.maxTableResults, 1000u)));

    const auto outcome = client.ListTables(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' namespace='{}'", bucketName, nsName);
        LogAwsFailure(L"S3Tables", L"ListTables", ctx, err, details);
        return HresultFromAwsError(err);
    }

    for (const auto& table : outcome.GetResult().GetTables())
    {
        const std::wstring name = Utf16FromUtf8(table.GetName());
        if (name.empty())
        {
            continue;
        }

        FilesInformationS3::Entry e{};
        e.name          = std::format(L"{}.table.json", name);
        e.attributes    = FILE_ATTRIBUTE_NORMAL;
        e.creationTime  = AwsDateTimeToFileTime64(table.GetCreatedAt());
        e.lastWriteTime = AwsDateTimeToFileTime64(table.GetModifiedAt());
        e.changeTime    = e.lastWriteTime;
        out.push_back(std::move(e));
    }

    return S_OK;
}

[[nodiscard]] HRESULT WriteS3TableInfoJson(FileSystemS3& fs,
                                           const ResolvedAwsContext& ctx,
                                           std::wstring_view bucketName,
                                           std::wstring_view nsName,
                                           std::wstring_view tableName,
                                           wil::unique_hfile& outFile) noexcept
{
    outFile.reset();

    std::string bucketArn;
    const HRESULT arnHr = EnsureS3TableBucketArn(fs, ctx, bucketName, bucketArn);
    if (FAILED(arnHr))
    {
        return arnHr;
    }

    const std::string nsUtf8 = Utf8FromUtf16(nsName);
    if (nsUtf8.empty() && ! nsName.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    const std::string tableUtf8 = Utf8FromUtf16(tableName);
    if (tableUtf8.empty() && ! tableName.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    Aws::S3Tables::S3TablesClient client = MakeS3TablesClient(ctx);
    Aws::S3Tables::Model::GetTableRequest req;
    req.SetTableBucketARN(bucketArn);
    req.SetNamespace(nsUtf8);
    req.SetName(tableUtf8);

    const auto outcome = client.GetTable(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' namespace='{}' table='{}'", bucketName, nsName, tableName);
        LogAwsFailure(L"S3Tables", L"GetTable", ctx, err, details);
        return HresultFromAwsError(err);
    }

    const auto& r = outcome.GetResult();

    Aws::Utils::Json::JsonValue json;
    json.WithString("name", r.GetName());
    json.WithString("tableArn", r.GetTableARN());

    // Namespace is returned as an array of strings.
    Aws::Utils::Array<Aws::String> nsArr(r.GetNamespace().size());
    for (size_t i = 0; i < r.GetNamespace().size(); ++i)
    {
        nsArr[i] = r.GetNamespace()[i];
    }
    json.WithArray("namespace", nsArr);

    json.WithString("metadataLocation", r.GetMetadataLocation());
    json.WithString("warehouseLocation", r.GetWarehouseLocation());
    json.WithString("versionToken", r.GetVersionToken());
    json.WithString("managedByService", r.GetManagedByService());

    json.WithString("createdAt", r.GetCreatedAt().ToGmtString(Aws::Utils::DateFormat::ISO_8601));

    const Aws::String readable = json.View().WriteReadable();
    const std::string utf8(readable.c_str(), readable.size());

    wil::unique_hfile file = CreateTemporaryDeleteOnCloseFile();
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    HRESULT hr = WriteUtf8ToFile(file.get(), utf8);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ResetFilePointerToStart(file.get());
    if (FAILED(hr))
    {
        return hr;
    }

    outFile = std::move(file);
    return S_OK;
}
} // namespace FileSystemS3Internal
