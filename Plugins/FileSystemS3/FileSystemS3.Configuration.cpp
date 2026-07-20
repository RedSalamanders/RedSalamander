#include "FileSystemS3.Internal.h"
#include "YyjsonHelpers.h"

namespace FsS3 = FileSystemS3Internal;

HRESULT STDMETHODCALLTYPE FileSystemS3::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    Settings nextSettings{};
    std::string nextConfiguration = "{}";

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        nextConfiguration = configurationJsonUtf8;
        Common::Json::ObjectDocument parsed = Common::Json::ParseObjectDocument(nextConfiguration, YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (! parsed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        yyjson_val* root = parsed.root;
        if (const auto value = FsS3::TryGetJsonString(root, "defaultRegion"); value.has_value())
        {
            nextSettings.defaultRegion = value.value();
            if (nextSettings.defaultRegion.empty())
            {
                nextSettings.defaultRegion = L"us-east-1";
            }
        }

        if (const auto value = FsS3::TryGetJsonString(root, "defaultEndpointOverride"); value.has_value())
        {
            nextSettings.defaultEndpointOverride = value.value();
        }

        if (const auto value = FsS3::TryGetJsonBool(root, "useHttps"); value.has_value())
        {
            nextSettings.useHttps = value.value();
        }

        if (const auto value = FsS3::TryGetJsonBool(root, "verifyTls"); value.has_value())
        {
            nextSettings.verifyTls = value.value();
        }

        if (const auto value = FsS3::TryGetJsonBool(root, "useVirtualAddressing"); value.has_value())
        {
            nextSettings.useVirtualAddressing = value.value();
        }

        if (const auto value = FsS3::TryGetJsonUInt(root, "maxKeys"); value.has_value() && value.value() >= 1u)
        {
            nextSettings.maxKeys = static_cast<unsigned long>((std::min)(value.value(), 1000ull));
        }

        if (const auto value = FsS3::TryGetJsonUInt(root, "maxTableResults"); value.has_value() && value.value() >= 1u)
        {
            nextSettings.maxTableResults = static_cast<unsigned long>((std::min)(value.value(), 1000ull));
        }

        if (const auto value = FsS3::TryGetJsonUInt(root, "connectTimeoutMs"); value.has_value() && value.value() >= 1u)
        {
            nextSettings.connectTimeoutMs = static_cast<uint32_t>((std::min)(value.value(), 600'000ull));
        }

        if (const auto value = FsS3::TryGetJsonUInt(root, "requestTimeoutMs"); value.has_value() && value.value() >= 1u)
        {
            nextSettings.requestTimeoutMs = static_cast<uint32_t>((std::min)(value.value(), 600'000ull));
        }
    }

    std::lock_guard lock(_stateMutex);
    _settings          = std::move(nextSettings);
    _configurationJson = std::move(nextConfiguration);
    _s3BucketRegionByName.clear();
    _s3ClientsByCtxKey.clear();
    _writableDirectoryValidationTicks.clear();
    _s3TableBucketArnByName.clear();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    const bool hasNonDefault = ! _configurationJson.empty() && _configurationJson != "{}";
    *pSomethingToSave        = hasNonDefault ? TRUE : FALSE;
    return S_OK;
}
