#include "FileSystemS3.Internal.h"
#include "HandleIo.h"
#include "YyjsonHelpers.h"

#include <functional>
#include <stdexcept>

namespace FileSystemS3Internal
{
namespace
{
class AwsSdkRuntime final
{
public:
    using Action = void (*)(void* cookie);

    AwsSdkRuntime(Action initialize, Action shutdown, void* cookie) noexcept : _initialize(initialize), _shutdown(shutdown), _cookie(cookie) {}

    [[nodiscard]] HRESULT Acquire() noexcept
    {
        while (true)
        {
            AcquireSRWLockExclusive(&_lock);
            if (_shutdownRequested)
            {
                ReleaseSRWLockExclusive(&_lock);
                return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
            }

            if (_state == State::Initialized)
            {
                ++_refCount;
                ReleaseSRWLockExclusive(&_lock);
                return S_OK;
            }
            if (_state == State::Failed)
            {
                const HRESULT failure = _failureStatus;
                ReleaseSRWLockExclusive(&_lock);
                return failure;
            }
            if (_state == State::Uninitialized)
            {
                _state = State::Initializing;
                ReleaseSRWLockExclusive(&_lock);
                break;
            }

            const BOOL waited = SleepConditionVariableSRW(&_changed, &_lock, INFINITE, 0u);
            const DWORD waitError = waited != FALSE ? ERROR_SUCCESS : GetLastError();
            ReleaseSRWLockExclusive(&_lock);
            if (waited == FALSE)
            {
                return HRESULT_FROM_WIN32(waitError != ERROR_SUCCESS ? waitError : ERROR_GEN_FAILURE);
            }
        }

        const HRESULT initializeStatus = InvokeAction(_initialize, L"Aws::InitAPI");

        AcquireSRWLockExclusive(&_lock);
        if (FAILED(initializeStatus))
        {
            _failureStatus = initializeStatus;
            _state         = State::Failed;
            WakeAllConditionVariable(&_changed);
            ReleaseSRWLockExclusive(&_lock);
            return initializeStatus;
        }

        if (! _shutdownRequested)
        {
            _refCount = 1u;
            _state    = State::Initialized;
            WakeAllConditionVariable(&_changed);
            ReleaseSRWLockExclusive(&_lock);
            return S_OK;
        }

        _state = State::ShuttingDown;
        WakeAllConditionVariable(&_changed);
        ReleaseSRWLockExclusive(&_lock);
        const HRESULT shutdownStatus = InvokeAction(_shutdown, L"Aws::ShutdownAPI");
        FinishShutdown(shutdownStatus);
        return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
    }

    void Release() noexcept
    {
        AcquireSRWLockExclusive(&_lock);
        if (_state != State::Initialized || _refCount == 0u)
        {
            ReleaseSRWLockExclusive(&_lock);
            Debug::Error(L"S3: AWS SDK lifetime release without a matching successful acquire");
            return;
        }

        --_refCount;
        if (_refCount != 0u)
        {
            ReleaseSRWLockExclusive(&_lock);
            return;
        }

        _state = State::ShuttingDown;
        ReleaseSRWLockExclusive(&_lock);

        const HRESULT shutdownStatus = InvokeAction(_shutdown, L"Aws::ShutdownAPI");
        FinishShutdown(shutdownStatus);
    }

    void BeginShutdown() noexcept
    {
        AcquireSRWLockExclusive(&_lock);
        _shutdownRequested = true;
        if (_state == State::Uninitialized)
        {
            _state = State::ShutdownComplete;
        }
        WakeAllConditionVariable(&_changed);
        ReleaseSRWLockExclusive(&_lock);
    }

    [[nodiscard]] bool CanUnloadNow() noexcept
    {
        AcquireSRWLockShared(&_lock);
        const bool canUnload = _shutdownRequested && _refCount == 0u && _state == State::ShutdownComplete;
        ReleaseSRWLockShared(&_lock);
        return canUnload;
    }

private:
    enum class State
    {
        Uninitialized,
        Initializing,
        Initialized,
        ShuttingDown,
        ShutdownComplete,
        Failed,
    };

    [[nodiscard]] HRESULT InvokeAction(Action action, std::wstring_view name) noexcept
    {
        if (action == nullptr)
        {
            return E_POINTER;
        }
        try
        {
            action(_cookie);
            return S_OK;
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::exception& error)
        {
            // AWS actions are third-party noexcept boundaries. Preserve a failed state so callers
            // cannot use or unload a partially initialized runtime.
            Debug::Error(L"S3: {} threw an exception: {}", name, Utf16FromUtf8(error.what()));
            return E_FAIL;
        }
    }

    void FinishShutdown(HRESULT status) noexcept
    {
        AcquireSRWLockExclusive(&_lock);
        if (FAILED(status))
        {
            _failureStatus = status;
            _state         = State::Failed;
        }
        else
        {
            _state = _shutdownRequested ? State::ShutdownComplete : State::Uninitialized;
        }
        WakeAllConditionVariable(&_changed);
        ReleaseSRWLockExclusive(&_lock);
    }

    SRWLOCK _lock = SRWLOCK_INIT;
    CONDITION_VARIABLE _changed = CONDITION_VARIABLE_INIT;
    Action _initialize = nullptr;
    Action _shutdown   = nullptr;
    void* _cookie      = nullptr;
    State _state       = State::Uninitialized;
    unsigned long _refCount = 0u;
    HRESULT _failureStatus  = E_FAIL;
    bool _shutdownRequested = false;
};

[[nodiscard]] Aws::SDKOptions& AwsOptions() noexcept
{
    static Aws::SDKOptions options{};
    return options;
}

void InitializeAwsSdk(void* cookie)
{
    auto* options = static_cast<Aws::SDKOptions*>(cookie);
    Aws::InitAPI(*options);
}

void ShutdownAwsSdk(void* cookie)
{
    auto* options = static_cast<Aws::SDKOptions*>(cookie);
    Aws::ShutdownAPI(*options);
}

[[nodiscard]] AwsSdkRuntime& Runtime() noexcept
{
    static AwsSdkRuntime runtime(&InitializeAwsSdk, &ShutdownAwsSdk, &AwsOptions());
    return runtime;
}
} // namespace

HRESULT AwsSdkLifetime::Acquire() noexcept
{
    return Runtime().Acquire();
}

void AwsSdkLifetime::Release() noexcept
{
    Runtime().Release();
}

void AwsSdkLifetime::BeginShutdown() noexcept
{
    Runtime().BeginShutdown();
}

bool AwsSdkLifetime::CanUnloadNow() noexcept
{
    return Runtime().CanUnloadNow();
}

#if defined(_DEBUG)
void RunDebugAwsSdkLifetimeContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept
    {
        if (condition)
        {
            ++passed;
        }
        else
        {
            ++failed;
            Debug::Error(L"FileSystemS3 AWS-lifetime selftest failed: {}", message);
        }
    };

    struct Context
    {
        unsigned int initializeCalls = 0u;
        unsigned int shutdownCalls   = 0u;
        bool throwDuringInitialize   = false;
    };

    const auto initialize = [](void* cookie)
    {
        auto* context = static_cast<Context*>(cookie);
        ++context->initializeCalls;
        if (context->throwDuringInitialize)
        {
            throw std::runtime_error("injected init failure");
        }
    };
    const auto shutdown = [](void* cookie)
    {
        auto* context = static_cast<Context*>(cookie);
        ++context->shutdownCalls;
    };

    Context failedContext{.throwDuringInitialize = true};
    AwsSdkRuntime failedRuntime(initialize, shutdown, &failedContext);
    check(failedRuntime.Acquire() == E_FAIL, L"AWS runtime should convert an InitAPI exception into factory-visible failure");
    check(failedRuntime.Acquire() == E_FAIL && failedContext.initializeCalls == 1u,
          L"AWS runtime should not increment or retry a partially initialized failed state");
    failedRuntime.BeginShutdown();
    check(! failedRuntime.CanUnloadNow(), L"AWS runtime should keep a partially initialized failed module mapped");

    Context normalContext{};
    AwsSdkRuntime normalRuntime(initialize, shutdown, &normalContext);
    check(normalRuntime.Acquire() == S_OK && normalRuntime.Acquire() == S_OK && normalContext.initializeCalls == 1u,
          L"AWS runtime should initialize once for concurrent owners");
    normalRuntime.BeginShutdown();
    check(! normalRuntime.CanUnloadNow(), L"AWS runtime should reject unload while acquired owners remain");
    normalRuntime.Release();
    check(! normalRuntime.CanUnloadNow(), L"AWS runtime should reject unload until the final owner releases");
    normalRuntime.Release();
    check(normalContext.shutdownCalls == 1u && normalRuntime.CanUnloadNow(),
          L"AWS runtime should permit unload only after ShutdownAPI returns for the final owner");
}
#endif

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
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
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
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
    return Common::HandleIo::GetFileSizeBounded(file, (std::numeric_limits<uint64_t>::max)(), out);
}

[[nodiscard]] HRESULT ResetFilePointerToStart(HANDLE file) noexcept
{
    return Common::HandleIo::Rewind(file);
}

[[nodiscard]] HRESULT WriteUtf8ToFile(HANDLE file, std::string_view text) noexcept
{
    return Common::HandleIo::WriteAll(file, text.data(), text.size());
}

[[nodiscard]] std::optional<std::wstring> TryGetJsonString(yyjson_val* root, const char* key) noexcept
{
    const Common::Json::MemberResult<std::wstring> value =
        Common::Json::GetUtf16StringMemberStrict(root, key, Common::Json::MemberRequirement::Optional);
    return value.HasValue() ? std::optional<std::wstring>{value.value} : std::nullopt;
}

[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt(yyjson_val* root, const char* key) noexcept
{
    const Common::Json::MemberResult<uint64_t> parsed = Common::Json::GetUInt64Member(root,
                                                                                      key,
                                                                                      Common::Json::MemberRequirement::Optional,
                                                                                      Common::Json::NumericStringPolicy::Reject,
                                                                                      Common::Json::UnsignedIntegerPolicy::RequireUnsignedStorage);
    return parsed.HasValue() ? std::optional<uint64_t>{parsed.value} : std::nullopt;
}

[[nodiscard]] std::optional<bool> TryGetJsonBool(yyjson_val* root, const char* key) noexcept
{
    const Common::Json::MemberResult<bool> value =
        Common::Json::GetBoolMember(root, key, Common::Json::MemberRequirement::Optional);
    return value.HasValue() ? std::optional<bool>{value.value} : std::nullopt;
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
    Common::Json::UniqueDocument docOwner{doc};

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
    if (! OrdinalString::EqualsNoCase(*pluginId, expectedId))
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
        std::string key = Utf8FromUtf16(*accessKeyWide);
        auto clearKey   = wil::scope_exit([&]() noexcept { SecureWipe::SecureClear(key); });
        if (key.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }
        if (out.accessKeyId.has_value())
        {
            SecureWipe::SecureClear(out.accessKeyId.value());
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
    out.connectTimeoutMs     = defaults.connectTimeoutMs;
    out.requestTimeoutMs     = defaults.requestTimeoutMs;

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
        auto clearSecretWide = wil::scope_exit([&]() noexcept
        {
            if (secret)
            {
                const int length = lstrlenW(secret.get());
                if (length > 0)
                {
                    SecureZeroMemory(secret.get(), static_cast<size_t>(length) * sizeof(wchar_t));
                }
            }
        });
        if (! secret.get() || secret.get()[0] == L'\0')
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
        }

        std::string secretUtf8 = Utf8FromUtf16(secret.get());
        auto clearSecretUtf8   = wil::scope_exit([&]() noexcept { SecureWipe::SecureClear(secretUtf8); });
        if (secretUtf8.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }
        if (out.secretAccessKey.has_value())
        {
            SecureWipe::SecureClear(out.secretAccessKey.value());
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
    else if (OrdinalString::EqualsNoCase(authority, L"@conn"))
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
    outContext.connectTimeoutMs     = defaults.connectTimeoutMs;
    outContext.requestTimeoutMs     = defaults.requestTimeoutMs;

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

    cfg.connectTimeoutMs = static_cast<long>(std::min<uint64_t>(ctx.connectTimeoutMs, static_cast<uint64_t>((std::numeric_limits<long>::max)())));
    cfg.requestTimeoutMs = static_cast<long>(std::min<uint64_t>(ctx.requestTimeoutMs, static_cast<uint64_t>((std::numeric_limits<long>::max)())));
    return cfg;
}

[[nodiscard]] std::shared_ptr<Aws::S3Crt::S3CrtClient> MakeS3Client(const ResolvedAwsContext& ctx) noexcept
{
    Aws::Client::ClientConfiguration legacy = MakeClientConfig(ctx);
    Aws::S3Crt::ClientConfiguration s3cfg(
        legacy, Aws::Client::AWSAuthV4Signer::PayloadSigningPolicy::Never, ctx.useVirtualAddressing, Aws::S3Crt::US_EAST_1_REGIONAL_ENDPOINT_OPTION::NOT_SET);

    if (ctx.accessKeyId.has_value() && ctx.secretAccessKey.has_value())
    {
        Aws::Auth::AWSCredentials creds(ctx.accessKeyId.value(), ctx.secretAccessKey.value());
        return std::make_shared<Aws::S3Crt::S3CrtClient>(creds, s3cfg);
    }

    return std::make_shared<Aws::S3Crt::S3CrtClient>(s3cfg);
}

namespace
{
void AppendHexU64(std::string& out, uint64_t value) noexcept
{
    constexpr std::string_view kHex = "0123456789ABCDEF";
    for (int shift = 60; shift >= 0; shift -= 4)
    {
        out.push_back(kHex[(value >> shift) & 0xFu]);
    }
}

[[nodiscard]] std::string MakeS3ClientCacheKey(const ResolvedAwsContext& ctx) noexcept
{
    uint64_t secretHash = 0;
    if (ctx.secretAccessKey.has_value())
    {
        secretHash = std::hash<std::string>{}(ctx.secretAccessKey.value());
    }

    const std::string conn = Utf8FromUtf16(ctx.connectionName);

    std::string key;
    key.reserve(conn.size() + ctx.region.size() + ctx.endpointOverride.size() + 64u + (ctx.accessKeyId.has_value() ? ctx.accessKeyId->size() : 0u));

    key.append(conn);
    key.push_back('|');
    key.append(ctx.region);
    key.push_back('|');
    key.append(ctx.endpointOverride);
    key.push_back('|');
    key.push_back(ctx.useHttps ? 'H' : 'h');
    key.push_back(ctx.verifyTls ? '1' : '0');
    key.push_back(ctx.useVirtualAddressing ? '1' : '0');
    key.push_back('|');
    if (ctx.accessKeyId.has_value())
    {
        key.append(ctx.accessKeyId.value());
    }
    key.push_back('|');
    AppendHexU64(key, secretHash);

    return key;
}
} // namespace

std::shared_ptr<Aws::S3Crt::S3CrtClient> GetS3Client(FileSystemS3& fs, const ResolvedAwsContext& ctx) noexcept
{
    const std::string key = MakeS3ClientCacheKey(ctx);

    {
        std::lock_guard lock(fs._stateMutex);
        if (const auto it = fs._s3ClientsByCtxKey.find(key); it != fs._s3ClientsByCtxKey.end() && it->second)
        {
            return it->second;
        }
    }

    // Client creation can be expensive; do it outside the lock.
    const std::shared_ptr<Aws::S3Crt::S3CrtClient> created = MakeS3Client(ctx);

    std::lock_guard lock(fs._stateMutex);
    auto it = fs._s3ClientsByCtxKey.find(key);
    if (it != fs._s3ClientsByCtxKey.end() && it->second)
    {
        return it->second;
    }

    fs._s3ClientsByCtxKey.emplace(key, created);
    return created;
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
