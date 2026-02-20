#include "FileSystemS3.Internal.h"

namespace FileSystemS3Internal
{
void AwsSdkLifetime::AddRef() noexcept
{
    const unsigned long prev = s_refCount.fetch_add(1, std::memory_order_acq_rel);
    if (prev != 0)
    {
        return;
    }

    // Mandatory: noexcept boundary. AWS SDK init is best-effort but must not throw.
    try
    {
        Debug::Warning(L"S3: Initializing AWS SDK");
        Aws::InitAPI(s_options);
    }
    catch (const std::bad_alloc&)
    {
        // Out-of-memory is treated as fatal. Fail-fast so the crash pipeline can capture a dump.
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"S3: Aws::InitAPI threw an exception");
    }
}

void AwsSdkLifetime::Release() noexcept
{
    const unsigned long prev = s_refCount.fetch_sub(1, std::memory_order_acq_rel);
    if (prev != 1)
    {
        return;
    }

    // Mandatory: noexcept boundary. AWS SDK shutdown is best-effort but must not throw.
    try
    {
        Debug::Warning(L"S3: Shutting down AWS SDK");
        const uint64_t startTickMs = GetTickCount64();
        Aws::ShutdownAPI(s_options);
        const uint64_t elapsedMs = GetTickCount64() - startTickMs;
        Debug::Warning(L"S3: AWS SDK shutdown complete ({} ms)", elapsedMs);
    }
    catch (const std::bad_alloc&)
    {
        // Out-of-memory is treated as fatal. Fail-fast so the crash pipeline can capture a dump.
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"S3: Aws::ShutdownAPI threw an exception");
    }
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring out(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), out.data(), required);
    if (written != required)
    {
        return {};
    }

    return out;
}

[[nodiscard]] std::wstring Utf16FromUtf8(const char* text) noexcept
{
    if (! text)
    {
        return {};
    }

    return Utf16FromUtf8(std::string_view(text));
}

[[nodiscard]] std::wstring Utf16FromUtf8(const Aws::String& text) noexcept
{
    return Utf16FromUtf8(std::string_view(text.c_str(), text.size()));
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string out(static_cast<size_t>(required), '\0');
    const int written = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), out.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return out;
}

[[nodiscard]] std::wstring NormalizePluginPath(std::wstring_view rawPath) noexcept
{
    std::wstring path(rawPath);
    if (path.empty())
    {
        return L"/";
    }

    for (wchar_t& ch : path)
    {
        if (ch == L'\\')
        {
            ch = L'/';
        }
    }

    const bool hasAuthorityPrefix = path.size() >= 2u && path[0] == L'/' && path[1] == L'/';

    if (! path.empty() && path.front() != L'/')
    {
        path.insert(path.begin(), L'/');
    }

    std::wstring collapsed;
    collapsed.reserve(path.size());

    bool prevSlash = false;
    size_t index   = 0;
    if (hasAuthorityPrefix)
    {
        collapsed.append(L"//");
        prevSlash = true;
        index     = 2;
        while (index < path.size() && path[index] == L'/')
        {
            ++index;
        }
    }

    for (; index < path.size(); ++index)
    {
        const wchar_t ch = path[index];
        const bool slash = (ch == L'/');
        if (slash && prevSlash)
        {
            continue;
        }
        collapsed.push_back(ch);
        prevSlash = slash;
    }

    if (collapsed.empty())
    {
        return L"/";
    }

    return collapsed;
}

[[nodiscard]] __int64 UnixMsToFileTime64(uint64_t unixMs) noexcept
{
    // FILETIME is 100-ns intervals since 1601-01-01 UTC.
    // Unix epoch is 1970-01-01 UTC.
    constexpr uint64_t kEpochDiff100ns = 116444736000000000ull;
    constexpr uint64_t kMsTo100ns      = 10000ull;

    const uint64_t fileTime100ns = kEpochDiff100ns + (unixMs * kMsTo100ns);
    if (fileTime100ns > static_cast<uint64_t>((std::numeric_limits<__int64>::max)()))
    {
        return static_cast<__int64>((std::numeric_limits<__int64>::max)());
    }
    return static_cast<__int64>(fileTime100ns);
}

[[nodiscard]] __int64 AwsDateTimeToFileTime64(const Aws::Utils::DateTime& t) noexcept
{
    const uint64_t ms = static_cast<uint64_t>(t.Millis());
    return UnixMsToFileTime64(ms);
}

[[nodiscard]] wil::unique_hfile CreateTemporaryDeleteOnCloseFile() noexcept
{
    wchar_t path[MAX_PATH + 1] = {};
    const DWORD len            = GetTempPathW(static_cast<DWORD>(std::size(path)), path);
    if (len == 0 || len >= std::size(path))
    {
        return {};
    }

    wchar_t name[MAX_PATH + 1] = {};
    if (GetTempFileNameW(path, L"rs3", 0, name) == 0)
    {
        return {};
    }

    wil::unique_hfile file(CreateFileW(
        name, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE, nullptr));
    if (! file)
    {
        DeleteFileW(name);
        return {};
    }

    return file;
}

[[nodiscard]] HRESULT GetFileSizeBytes(HANDLE file, uint64_t& out) noexcept
{
    out = 0;
    if (! file || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    LARGE_INTEGER size{};
    if (GetFileSizeEx(file, &size) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if (size.QuadPart < 0)
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    out = static_cast<uint64_t>(size.QuadPart);
    return S_OK;
}

[[nodiscard]] HRESULT ResetFilePointerToStart(HANDLE file) noexcept
{
    if (! file || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    LARGE_INTEGER zero{};
    if (SetFilePointerEx(file, zero, nullptr, FILE_BEGIN) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    return S_OK;
}

[[nodiscard]] HRESULT WriteUtf8ToFile(HANDLE file, std::string_view text) noexcept
{
    if (! file || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    size_t offset = 0;
    while (offset < text.size())
    {
        const size_t chunk = std::min<size_t>(text.size() - offset, static_cast<size_t>((std::numeric_limits<DWORD>::max)()));
        DWORD written      = 0;
        if (WriteFile(file, text.data() + offset, static_cast<DWORD>(chunk), &written, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (written != chunk)
        {
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }

        offset += chunk;
    }

    return S_OK;
}

[[nodiscard]] std::optional<std::wstring> TryGetJsonString(yyjson_val* root, const char* key) noexcept
{
    if (! root || ! yyjson_is_obj(root) || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* v = yyjson_obj_get(root, key);
    if (! v || ! yyjson_is_str(v))
    {
        return std::nullopt;
    }

    const char* s    = yyjson_get_str(v);
    const size_t len = yyjson_get_len(v);
    if (! s)
    {
        return std::nullopt;
    }

    return Utf16FromUtf8(std::string_view(s, len));
}

[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt(yyjson_val* root, const char* key) noexcept
{
    if (! root || ! yyjson_is_obj(root) || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* v = yyjson_obj_get(root, key);
    if (! v || ! yyjson_is_uint(v))
    {
        return std::nullopt;
    }

    return yyjson_get_uint(v);
}

[[nodiscard]] std::optional<bool> TryGetJsonBool(yyjson_val* root, const char* key) noexcept
{
    if (! root || ! yyjson_is_obj(root) || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* v = yyjson_obj_get(root, key);
    if (! v || ! yyjson_is_bool(v))
    {
        return std::nullopt;
    }

    return yyjson_get_bool(v) != 0;
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    return OrdinalString::EqualsNoCase(a, b);
}

namespace
{
[[nodiscard]] HRESULT ResolveConnectionManagerProfile(IHostConnections* hostConnections,
                                                      FileSystemS3Mode mode,
                                                      std::wstring_view connectionName,
                                                      bool acquireSecrets,
                                                      const FileSystemS3::Settings& defaults,
                                                      ResolvedAwsContext& out) noexcept
{
    if (! hostConnections)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    wil::unique_cotaskmem_ptr<char> json;
    {
        char* rawJson    = nullptr;
        const HRESULT hr = hostConnections->GetConnectionJsonUtf8(std::wstring(connectionName).c_str(), &rawJson);
        if (FAILED(hr))
        {
            return hr;
        }
        json.reset(rawJson);
    }

    if (! json || ! json.get()[0])
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    yyjson_doc* doc = yyjson_read(json.get(), strlen(json.get()), YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const auto pluginId = TryGetJsonString(root, "pluginId");
    if (! pluginId.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    constexpr std::wstring_view kPluginIdS3      = L"builtin/file-system-s3";
    constexpr std::wstring_view kPluginIdS3Table = L"builtin/file-system-s3table";
    const std::wstring_view expectedId           = (mode == FileSystemS3Mode::S3) ? kPluginIdS3 : kPluginIdS3Table;
    if (! EqualsNoCase(*pluginId, expectedId))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    out                = {};
    out.connectionName = std::wstring(connectionName);

    const auto regionWide = TryGetJsonString(root, "host");
    if (regionWide.has_value() && ! regionWide->empty())
    {
        out.region = Utf8FromUtf16(*regionWide);
        if (out.region.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }
        out.explicitRegion = out.region;
    }
    else
    {
        out.region = Utf8FromUtf16(defaults.defaultRegion);
        if (out.region.empty())
        {
            out.region = "us-east-1";
        }
    }

    const auto accessKeyWide = TryGetJsonString(root, "userName");
    if (accessKeyWide.has_value() && ! accessKeyWide->empty())
    {
        const std::string key = Utf8FromUtf16(*accessKeyWide);
        if (key.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }
        out.accessKeyId = key;
    }

    // Default config (may be overridden by connection.extra)
    out.endpointOverride     = Utf8FromUtf16(defaults.defaultEndpointOverride);
    out.useHttps             = defaults.useHttps;
    out.verifyTls            = defaults.verifyTls;
    out.useVirtualAddressing = defaults.useVirtualAddressing;
    out.maxKeys              = defaults.maxKeys;
    out.maxTableResults      = defaults.maxTableResults;

    // extra payload (optional; forwarded by host as `extra`)
    if (yyjson_val* extra = yyjson_obj_get(root, "extra"); extra && yyjson_is_obj(extra))
    {
        if (const auto v = TryGetJsonString(extra, "endpointOverride"); v.has_value())
        {
            out.endpointOverride = Utf8FromUtf16(*v);
        }
        if (const auto v = TryGetJsonBool(extra, "useHttps"); v.has_value())
        {
            out.useHttps = v.value();
        }
        if (const auto v = TryGetJsonBool(extra, "verifyTls"); v.has_value())
        {
            out.verifyTls = v.value();
        }
        if (const auto v = TryGetJsonBool(extra, "useVirtualAddressing"); v.has_value())
        {
            out.useVirtualAddressing = v.value();
        }
    }

    if (out.accessKeyId.has_value() && acquireSecrets)
    {
        wil::unique_cotaskmem_string secret;
        wchar_t* rawSecret = nullptr;
        HRESULT secretHr   = hostConnections->GetConnectionSecret(std::wstring(connectionName).c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, &rawSecret);
        if (secretHr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
        {
            rawSecret = nullptr;
            secretHr  = hostConnections->PromptForConnectionSecret(std::wstring(connectionName).c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, &rawSecret);
            if (secretHr == S_FALSE)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
        }
        if (FAILED(secretHr))
        {
            Debug::Error(L"S3: GetConnectionSecret failed conn='{}' hr=0x{:08X}", connectionName, static_cast<unsigned long>(secretHr));
            return secretHr;
        }

        secret.reset(rawSecret);
        if (! secret.get() || secret.get()[0] == L'\0')
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
        }

        const std::string secretUtf8 = Utf8FromUtf16(secret.get());
        if (secretUtf8.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }
        out.secretAccessKey = secretUtf8;
    }

    return S_OK;
}
} // namespace

[[nodiscard]] HRESULT ResolveAwsContext(FileSystemS3Mode mode,
                                        const FileSystemS3::Settings& defaults,
                                        std::wstring_view pluginPath,
                                        IHostConnections* hostConnections,
                                        bool acquireSecrets,
                                        ResolvedAwsContext& outContext,
                                        std::wstring& outCanonicalPath) noexcept
{
    outContext = {};
    outCanonicalPath.clear();

    const std::wstring normalizedFull = NormalizePluginPath(pluginPath);

    // Split optional URI authority: //<authority>/<path>
    std::wstring_view authority;
    std::wstring_view pathPart;

    if (normalizedFull.size() >= 2u && normalizedFull[0] == L'/' && normalizedFull[1] == L'/')
    {
        std::wstring_view after(normalizedFull);
        after.remove_prefix(2);
        const size_t slashPos = after.find(L'/');
        authority             = slashPos == std::wstring_view::npos ? after : after.substr(0, slashPos);
        pathPart              = slashPos == std::wstring_view::npos ? std::wstring_view(L"/") : after.substr(slashPos);
    }
    else
    {
        pathPart = normalizedFull;
    }

    // Connection Manager prefix: /@conn:<name>/...
    bool hasConnPrefix = false;
    std::wstring_view connectionName;
    std::wstring_view connPath = pathPart;

    constexpr std::wstring_view kConnPrefix = L"@conn:";
    if (! pathPart.empty())
    {
        std::wstring_view rest = pathPart;
        while (! rest.empty() && rest.front() == L'/')
        {
            rest.remove_prefix(1);
        }

        if (rest.rfind(kConnPrefix, 0) == 0)
        {
            rest.remove_prefix(kConnPrefix.size());
            const size_t slashPos = rest.find(L'/');
            connectionName        = slashPos == std::wstring_view::npos ? rest : rest.substr(0, slashPos);
            connPath              = slashPos == std::wstring_view::npos ? std::wstring_view(L"/") : rest.substr(slashPos);
            hasConnPrefix         = true;
        }
    }
    else if (EqualsNoCase(authority, L"@conn"))
    {
        // URI-style shorthand: // @conn / <connectionName> / ...
        std::wstring_view rest = pathPart;
        while (! rest.empty() && rest.front() == L'/')
        {
            rest.remove_prefix(1);
        }

        const size_t slashPos = rest.find(L'/');
        connectionName        = slashPos == std::wstring_view::npos ? rest : rest.substr(0, slashPos);
        connPath              = slashPos == std::wstring_view::npos ? std::wstring_view(L"/") : rest.substr(slashPos);
        hasConnPrefix         = true;
    }

    if (hasConnPrefix)
    {
        if (connectionName.empty())
        {
            return E_INVALIDARG;
        }

        HRESULT hr = ResolveConnectionManagerProfile(hostConnections, mode, connectionName, acquireSecrets, defaults, outContext);
        if (FAILED(hr))
        {
            return hr;
        }

        outCanonicalPath = NormalizePluginPath(connPath);
        if (outCanonicalPath.empty())
        {
            outCanonicalPath = L"/";
        }

        return S_OK;
    }

    // No Connection Manager profile: use defaults and AWS default credential chain.
    outContext.region = Utf8FromUtf16(defaults.defaultRegion);
    if (outContext.region.empty())
    {
        outContext.region = "us-east-1";
    }

    outContext.endpointOverride     = Utf8FromUtf16(defaults.defaultEndpointOverride);
    outContext.useHttps             = defaults.useHttps;
    outContext.verifyTls            = defaults.verifyTls;
    outContext.useVirtualAddressing = defaults.useVirtualAddressing;
    outContext.maxKeys              = defaults.maxKeys;
    outContext.maxTableResults      = defaults.maxTableResults;

    // Canonicalize authority-based paths (s3://bucket/...) into "/bucket/..."
    if (! authority.empty())
    {
        std::wstring tmp;
        tmp.reserve(1u + authority.size() + pathPart.size());
        tmp.push_back(L'/');
        tmp.append(authority);
        tmp.append(pathPart);

        outCanonicalPath = NormalizePluginPath(tmp);
    }
    else
    {
        outCanonicalPath = NormalizePluginPath(pathPart);
    }

    if (outCanonicalPath.empty())
    {
        outCanonicalPath = L"/";
    }

    return S_OK;
}

[[nodiscard]] Aws::Client::ClientConfiguration MakeClientConfig(const ResolvedAwsContext& ctx) noexcept
{
    Aws::Client::ClientConfiguration cfg;
    cfg.region = ctx.region;

    cfg.scheme    = ctx.useHttps ? Aws::Http::Scheme::HTTPS : Aws::Http::Scheme::HTTP;
    cfg.verifySSL = ctx.verifyTls;

    if (! ctx.endpointOverride.empty())
    {
        std::string endpoint = ctx.endpointOverride;

        constexpr std::string_view kHttp  = "http://";
        constexpr std::string_view kHttps = "https://";

        if (endpoint.rfind(kHttp, 0) == 0)
        {
            cfg.scheme = Aws::Http::Scheme::HTTP;
            endpoint.erase(0, kHttp.size());
        }
        else if (endpoint.rfind(kHttps, 0) == 0)
        {
            cfg.scheme = Aws::Http::Scheme::HTTPS;
            endpoint.erase(0, kHttps.size());
        }

        while (! endpoint.empty() && endpoint.back() == '/')
        {
            endpoint.pop_back();
        }

        if (! endpoint.empty())
        {
            cfg.endpointOverride = std::move(endpoint);
        }
    }

    // Conservative defaults; can be made configurable later.
    cfg.connectTimeoutMs = 10'000;
    cfg.requestTimeoutMs = 0;
    return cfg;
}

[[nodiscard]] Aws::S3Crt::S3CrtClient MakeS3Client(const ResolvedAwsContext& ctx) noexcept
{
    Aws::Client::ClientConfiguration legacy = MakeClientConfig(ctx);
    Aws::S3Crt::ClientConfiguration s3cfg(
        legacy, Aws::Client::AWSAuthV4Signer::PayloadSigningPolicy::Never, ctx.useVirtualAddressing, Aws::S3Crt::US_EAST_1_REGIONAL_ENDPOINT_OPTION::NOT_SET);

    if (ctx.accessKeyId.has_value() && ctx.secretAccessKey.has_value())
    {
        Aws::Auth::AWSCredentials creds(ctx.accessKeyId.value(), ctx.secretAccessKey.value());
        return Aws::S3Crt::S3CrtClient(creds, s3cfg);
    }

    return Aws::S3Crt::S3CrtClient(s3cfg);
}

[[nodiscard]] Aws::S3Tables::S3TablesClient MakeS3TablesClient(const ResolvedAwsContext& ctx) noexcept
{
    Aws::Client::ClientConfiguration cfg = MakeClientConfig(ctx);
    if (ctx.accessKeyId.has_value() && ctx.secretAccessKey.has_value())
    {
        Aws::Auth::AWSCredentials creds(ctx.accessKeyId.value(), ctx.secretAccessKey.value());
        return Aws::S3Tables::S3TablesClient(creds, cfg);
    }

    return Aws::S3Tables::S3TablesClient(cfg);
}

[[nodiscard]] std::vector<std::wstring_view> SplitPathSegments(std::wstring_view path) noexcept
{
    std::vector<std::wstring_view> segments;
    while (! path.empty() && path.front() == L'/')
    {
        path.remove_prefix(1);
    }
    while (! path.empty())
    {
        const size_t slash           = path.find(L'/');
        const std::wstring_view part = (slash == std::wstring_view::npos) ? path : path.substr(0, slash);
        if (! part.empty())
        {
            segments.push_back(part);
        }
        if (slash == std::wstring_view::npos)
        {
            break;
        }
        path.remove_prefix(slash + 1);
        while (! path.empty() && path.front() == L'/')
        {
            path.remove_prefix(1);
        }
    }
    return segments;
}
} // namespace FileSystemS3Internal
