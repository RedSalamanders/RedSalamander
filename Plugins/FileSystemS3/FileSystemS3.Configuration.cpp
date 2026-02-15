#include "FileSystemS3.Internal.h"

namespace FsS3 = FileSystemS3Internal;

HRESULT STDMETHODCALLTYPE FileSystemS3::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    std::lock_guard lock(_stateMutex);

    _settings = {};

    if (configurationJsonUtf8 == nullptr || configurationJsonUtf8[0] == '\0')
    {
        _configurationJson = "{}";
        return S_OK;
    }

    _configurationJson = configurationJsonUtf8;

    yyjson_read_err err{};
    yyjson_doc* doc = yyjson_read_opts(_configurationJson.data(), _configurationJson.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
    if (! doc)
    {
        return S_OK;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return S_OK;
    }

    if (const auto v = FsS3::TryGetJsonString(root, "defaultRegion"); v.has_value())
    {
        _settings.defaultRegion = v.value();
        if (_settings.defaultRegion.empty())
        {
            _settings.defaultRegion = L"us-east-1";
        }
    }

    if (const auto v = FsS3::TryGetJsonString(root, "defaultEndpointOverride"); v.has_value())
    {
        _settings.defaultEndpointOverride = v.value();
    }

    if (const auto v = FsS3::TryGetJsonBool(root, "useHttps"); v.has_value())
    {
        _settings.useHttps = v.value();
    }

    if (const auto v = FsS3::TryGetJsonBool(root, "verifyTls"); v.has_value())
    {
        _settings.verifyTls = v.value();
    }

    if (const auto v = FsS3::TryGetJsonBool(root, "useVirtualAddressing"); v.has_value())
    {
        _settings.useVirtualAddressing = v.value();
    }

    if (const auto v = FsS3::TryGetJsonUInt(root, "maxKeys"); v.has_value())
    {
        const uint64_t raw = v.value();
        if (raw >= 1u)
        {
            _settings.maxKeys = static_cast<unsigned long>(std::min<uint64_t>(raw, 1000u));
        }
    }

    if (const auto v = FsS3::TryGetJsonUInt(root, "maxTableResults"); v.has_value())
    {
        const uint64_t raw = v.value();
        if (raw >= 1u)
        {
            _settings.maxTableResults = static_cast<unsigned long>(std::min<uint64_t>(raw, 1000u));
        }
    }

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
